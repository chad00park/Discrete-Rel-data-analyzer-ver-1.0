#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Discrete Reliability Data Analyzer
==================================
경량 스택: Tkinter + Matplotlib + csv/openpyxl (Pandas/PySide6 미사용)
- 여러 Read-out 측정 파일(CSV/XLSX)을 읽어 Parameter별 변화/산포 분석
- Interactive Line Graph / Box Plot + 편집(Undo/Redo) + PDF Report
"""

import csv
import math
import os
import re
import sys
import threading
import traceback

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser, simpledialog

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.ticker import MultipleLocator

# Drag & Drop (선택적 — 미설치 시 Browse만 동작)
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except Exception:
    HAS_DND = False

# ----------------------------------------------------------------------------
# 제한사항
# ----------------------------------------------------------------------------
MAX_SAMPLES = 500
MAX_PARAMS = 500
MAX_READOUTS = 20

READOUT_COLORS = [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
    "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf",
    "#aec7e8", "#ffbb78", "#98df8a", "#ff9896", "#c5b0d5",
    "#c49c94", "#f7b6d2", "#c7c7c7", "#dbdb8d", "#9edae5",
]


# ============================================================================
# 1. 파일명 파싱: 신뢰성명 + Lot번호 + Read-out (구분자 _ - + 공백, 순서 무관)
# ============================================================================
READOUT_RE = re.compile(r"^(\d+(?:\.\d+)?)\s*(hr|hrs|h|hour|hours|cyc|cycle|cycles|cy)$", re.I)
LOT_RE = re.compile(r"^lot\s*([A-Za-z0-9]+)$", re.I)


def parse_filename(path):
    """파일명에서 (reliability, lot, readout_label, readout_value) 추출.
    실패 시 ValueError."""
    base = os.path.splitext(os.path.basename(path))[0]
    tokens = [t for t in re.split(r"[_\-\+\s]+", base) if t]
    readout_label = None
    readout_value = None
    lot = None
    rest = []
    for tok in tokens:
        m = READOUT_RE.match(tok)
        if m and readout_label is None:
            readout_value = float(m.group(1))
            unit = m.group(2).lower()
            unit = "hr" if unit.startswith("h") else "cyc"
            num = m.group(1)
            readout_label = f"{num}{unit}"
            continue
        m = LOT_RE.match(tok)
        if m and lot is None:
            lot = "LOT" + m.group(1).upper()
            continue
        rest.append(tok)
    if readout_label is None or lot is None or not rest:
        raise ValueError(
            f"파일명 인식 실패: '{os.path.basename(path)}'\n"
            "파일 이름은 신뢰성명 + Lot번호 + Read-out 형식이어야 합니다.\n"
            "예: HTRB_Lot1_0hr, TC+Lot2+500cyc"
        )
    reliability = "_".join(rest)
    return reliability, lot, readout_label, readout_value


# ============================================================================
# 2. 데이터 파싱 (좌표 규칙: Column6 키워드 탐색)
# ============================================================================
def _read_rows(path):
    """CSV/XLSX → 모든 셀 strip()된 문자열 2차원 리스트."""
    ext = os.path.splitext(path)[1].lower()
    rows = []
    if ext == ".csv":
        # 인코딩 유연 처리 (utf-8-sig 우선, cp949 폴백)
        for enc in ("utf-8-sig", "cp949", "latin-1"):
            try:
                with open(path, newline="", encoding=enc) as f:
                    rows = [[(c or "").strip() for c in r] for r in csv.reader(f)]
                break
            except UnicodeDecodeError:
                rows = []
                continue
    elif ext in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        for r in ws.iter_rows(values_only=True):
            rows.append([("" if c is None else str(c)).strip() for c in r])
        wb.close()
    else:
        raise ValueError(f"지원하지 않는 확장자: {ext}")
    return rows


def _cell(row, idx):
    return row[idx].strip() if idx < len(row) else ""


def parse_data_file(path):
    """좌표 규칙에 따라 파일 파싱.
    반환: (columns, data)
      columns: [f"Item@Bias1 (Unit)"] — Unit이 빈 열 제외
      data: {sample_no(int): {colname: float|None}}
    실패 시 ValueError(파일명 포함 메시지)."""
    fname = os.path.basename(path)
    rows = _read_rows(path)

    item_row = bias_row = unit_row = None
    data_start = None
    for i, row in enumerate(rows):
        c6 = _cell(row, 6)
        if item_row is None and c6 == "Item":
            item_row = row  # 첫 번째 Item 행만 사용 (통계 블록의 두 번째 Item 무시)
        elif c6 == "Bias1" and bias_row is None:
            bias_row = row
        elif c6 == "Unit" and unit_row is None:
            unit_row = row
        if data_start is None and _cell(row, 0) == "Test No.":
            data_start = i + 1

    missing = []
    if item_row is None:
        missing.append("Item")
    if bias_row is None:
        missing.append("Bias1")
    if unit_row is None:
        missing.append("Unit")
    if data_start is None:
        missing.append("Test No.")
    if missing:
        raise ValueError(f"'{fname}' 에서 {', '.join(missing)} 행을 찾을 수 없습니다.")

    # Column7부터 Item/Bias1/Unit이 1:1 매핑. Unit이 빈 열은 제거.
    ncol = max(len(item_row), len(bias_row), len(unit_row))
    col_map = []  # (colname, column_index)
    name_count = {}
    for j in range(7, ncol):
        item = _cell(item_row, j)
        unit = _cell(unit_row, j)
        bias = _cell(bias_row, j)
        if not item or not unit:
            continue
        name = f"{item}@{bias} ({unit})"
        # Item+Bias1+Unit이 완전히 동일한 컬럼이 중복 등장하면 #2, #3... 접미사로 구분
        if name in name_count:
            name_count[name] += 1
            name = f"{name} #{name_count[name]}"
        else:
            name_count[name] = 1
        col_map.append((name, j))

    data = {}
    for row in rows[data_start:]:
        s = _cell(row, 0)
        if not s:
            continue
        try:
            sample = int(float(s))
        except ValueError:
            continue
        vals = {}
        for colname, j in col_map:
            v = _cell(row, j)
            try:
                vals[colname] = float(v)
            except ValueError:
                vals[colname] = None  # 결측 → NaN 처리
        data[sample] = vals

    columns = [c for c, _ in col_map]
    return columns, data


# ============================================================================
# 3. 데이터 모델 (경량 테이블 + diff 기반 Undo/Redo)
# ============================================================================
class DataModel:
    def __init__(self):
        self.reliability = None
        self.lot = None
        self.readouts = []          # 정렬된 label 목록
        self.columns = []           # 전체 Parameter 컬럼명
        self.data = {}              # data[readout][colname][sample] = float|None
        self.samples = []           # 정렬된 전체 sample 목록
        # 편집 상태
        self.deleted = set()        # (readout, colname, sample)
        self.color_over = {}        # (readout, colname, sample) -> hex color
        self.ylim = {}              # colname -> (ymin, ymax)
        self._undo = []
        self._redo = []

    # ---- 로드 ------------------------------------------------------------
    def load(self, files):
        """files 파싱. 반환: (errors:list[str])"""
        errors = []
        parsed = []  # (readout_label, readout_value, columns, data)
        rel = lot = None
        for p in files:
            try:
                r, l, rl, rv = parse_filename(p)
                cols, d = parse_data_file(p)
            except ValueError as e:
                errors.append(str(e))
                continue
            if rel is None:
                rel, lot = r, l
            parsed.append((rl, rv, cols, d))

        if not parsed:
            return errors

        if len(parsed) > MAX_READOUTS:
            errors.append(f"Read-out 파일이 {MAX_READOUTS}개를 초과합니다 ({len(parsed)}개).")
            return errors

        parsed.sort(key=lambda x: x[1])
        self.reliability, self.lot = rel, lot
        self.readouts = [p[0] for p in parsed]
        col_union, seen = [], set()
        samp_union = set()
        self.data = {}
        for rl, rv, cols, d in parsed:
            for c in cols:
                if c not in seen:
                    seen.add(c)
                    col_union.append(c)
            table = {}
            for s, vals in d.items():
                samp_union.add(s)
            self.data[rl] = d

        if len(col_union) > MAX_PARAMS:
            errors.append(f"Parameter가 {MAX_PARAMS}개를 초과합니다 ({len(col_union)}개).")
            return errors
        if len(samp_union) > MAX_SAMPLES:
            errors.append(f"Sample이 {MAX_SAMPLES}개를 초과합니다 ({len(samp_union)}개).")
            return errors

        self.columns = col_union
        self.samples = sorted(samp_union)
        self.deleted.clear()
        self.color_over.clear()
        self.ylim.clear()
        self._undo.clear()
        self._redo.clear()
        return errors

    # ---- 값 조회 -----------------------------------------------------------
    def value(self, readout, col, sample):
        if (readout, col, sample) in self.deleted:
            return None
        return self.data.get(readout, {}).get(sample, {}).get(col)

    def series(self, readout, col):
        """(samples, values) — 삭제된 포인트는 NaN, 축은 유지."""
        xs, ys = [], []
        for s in self.samples:
            xs.append(s)
            v = self.value(readout, col, s)
            ys.append(math.nan if v is None else v)
        return xs, ys

    def box_values(self, readout, col):
        return [v for s in self.samples
                if (v := self.value(readout, col, s)) is not None]

    # ---- diff 기반 Undo/Redo ------------------------------------------------
    def _apply(self, action, forward=True):
        kind = action[0]
        if kind == "del":
            _, keys = action
            if forward:
                self.deleted.update(keys)
            else:
                self.deleted.difference_update(keys)
        elif kind == "color":
            _, key, old, new = action
            c = new if forward else old
            if c is None:
                self.color_over.pop(key, None)
            else:
                self.color_over[key] = c
        elif kind == "ylim":
            _, col, old, new = action
            v = new if forward else old
            if v is None:
                self.ylim.pop(col, None)
            else:
                self.ylim[col] = v

    def do(self, action):
        self._apply(action, True)
        self._undo.append(action)
        self._redo.clear()

    def undo(self):
        if not self._undo:
            return False
        a = self._undo.pop()
        self._apply(a, False)
        self._redo.append(a)
        return True

    def redo(self):
        if not self._redo:
            return False
        a = self._redo.pop()
        self._apply(a, True)
        self._undo.append(a)
        return True

    # ---- 편집 액션 -----------------------------------------------------------
    def delete_point(self, readout, col, sample):
        self.do(("del", frozenset({(readout, col, sample)})))

    def delete_sample_all_readouts(self, col, sample):
        keys = frozenset((r, col, sample) for r in self.readouts)
        self.do(("del", keys))

    def set_color(self, readout, col, sample, color):
        key = (readout, col, sample)
        self.do(("color", key, self.color_over.get(key), color))

    def set_ylim(self, col, ymin, ymax):
        self.do(("ylim", col, self.ylim.get(col),
                 None if ymin is None else (ymin, ymax)))

    # ---- 통계 -----------------------------------------------------------------
    def stats(self, readout, col):
        vals = self.box_values(readout, col)
        n = len(vals)
        if n == 0:
            return dict(SS=0, Min=math.nan, Max=math.nan, AVG=math.nan, STD=math.nan)
        avg = sum(vals) / n
        std = (sum((v - avg) ** 2 for v in vals) / (n - 1)) ** 0.5 if n > 1 else 0.0
        return dict(SS=n, Min=min(vals), Max=max(vals), AVG=avg, STD=std)


# ============================================================================
# 4. 그래프 렌더링 (화면/PDF 공용 Matplotlib 파이프라인)
# ============================================================================
def readout_color(model, readout):
    return READOUT_COLORS[model.readouts.index(readout) % len(READOUT_COLORS)]


def draw_line(ax, model, col, picker=False):
    """Line Graph: X=Sample(짝수 Label/홀수 Minor Tick), Read-out별 색상."""
    artists = {}
    for r in model.readouts:
        c = readout_color(model, r)
        xs, ys = model.series(r, col)
        kw = dict(marker="o", ms=4, lw=1, color=c, label=r)
        if picker:
            kw["picker"] = 5
        line, = ax.plot(xs, ys, **kw)
        artists[r] = line
        # 색 변경된 마커는 삼각형으로 덧그림
        for s, v in zip(xs, ys):
            oc = model.color_over.get((r, col, s))
            if oc and not math.isnan(v):
                ax.plot([s], [v], marker="^", ms=8, color=oc, ls="none", zorder=5)
    ax.set_title(col, fontsize=9)
    ax.set_xlabel("Sample No.", fontsize=8)
    ax.set_ylabel(col.split("@")[0], fontsize=8)
    if model.samples:
        ax.set_xlim(min(model.samples) - 1, max(model.samples) + 1)
        even = [s for s in model.samples if s % 2 == 0]
        ax.set_xticks(even)
        ax.xaxis.set_minor_locator(MultipleLocator(1))
    ax.tick_params(labelsize=7)
    if col in model.ylim:
        ax.set_ylim(*model.ylim[col])
    ax.legend(fontsize=6, ncol=2)
    ax.grid(True, alpha=0.3)
    return artists


def draw_box(ax, model, col, picker=False, stats_table=True):
    """Box Plot: X=Read-out, Line Graph와 동일 색상, Median/Outlier 표시.
    stats_table=True면 각 Read-out 박스 위치에 맞춰 아래에 통계표 표시."""
    groups = [model.box_values(r, col) for r in model.readouts]
    bp = ax.boxplot(groups, labels=model.readouts, patch_artist=True,
                    showfliers=True, medianprops=dict(color="black"))
    for patch, r in zip(bp["boxes"], model.readouts):
        c = readout_color(model, r)
        patch.set_facecolor(c)
        patch.set_alpha(0.5)
    for fl, r in zip(bp["fliers"], model.readouts):
        fl.set(marker="o", markerfacecolor=readout_color(model, r),
               markeredgecolor="black", markersize=5)
        if picker:
            fl.set_picker(5)
    ax.set_title(col, fontsize=9)
    ax.tick_params(labelsize=7)
    if col in model.ylim:
        ax.set_ylim(*model.ylim[col])
    ax.grid(True, alpha=0.3)

    if stats_table:
        # Read-out별 통계를 박스 x위치에 1:1 정렬된 표로 표시
        rows = [("S/S", "SS"), ("Min", "Min"), ("Max", "Max"),
                ("AVG", "AVG"), ("STD", "STD")]
        cell = []
        for label, key in rows:
            line = []
            for r in model.readouts:
                st = model.stats(r, col)
                v = st[key]
                line.append(str(v) if key == "SS" else f"{v:.4g}")
            cell.append(line)
        tbl = ax.table(cellText=cell, rowLabels=[l for l, _ in rows],
                       colLabels=model.readouts,
                       cellLoc="center", rowLoc="center",
                       bbox=[0.0, -0.62, 1.0, 0.44])
        tbl.auto_set_font_size(False)
        tbl.set_fontsize(6)
        ax.set_xlabel("")  # 표의 colLabels가 Read-out 라벨을 대신함
        ax.set_xticklabels([])
    else:
        ax.set_xlabel("Read-out", fontsize=8)
    return bp


def stats_text(model, col):
    lines = [f"{'Read-out':<10}{'S/S':>5}{'Min':>12}{'Max':>12}{'AVG':>12}{'STD':>12}"]
    for r in model.readouts:
        st = model.stats(r, col)
        lines.append(f"{r:<10}{st['SS']:>5}{st['Min']:>12.4g}{st['Max']:>12.4g}"
                     f"{st['AVG']:>12.4g}{st['STD']:>12.4g}")
    return "\n".join(lines)


# ============================================================================
# 5. PDF 출력 (Line 1×4 / Box 4×3, A4 Landscape, 마지막 페이지 크기 유지)
# ============================================================================
A4L = (11.69, 8.27)


def export_pdf(model, cols, path, progress_cb=None):
    total = math.ceil(len(cols) / 4) + math.ceil(len(cols) / 12)
    done = 0
    with PdfPages(path) as pdf:
        title = f"{model.reliability}_{model.lot} Data Analysis"
        # Line Graph: 1×4 (한 페이지 4개, 세로 배치)
        for i in range(0, len(cols), 4):
            fig = Figure(figsize=A4L)
            fig.suptitle(title, fontsize=12)
            for k in range(4):  # 마지막 페이지도 4칸 유지 → 동일 크기
                ax = fig.add_subplot(4, 1, k + 1)
                if i + k < len(cols):
                    draw_line(ax, model, cols[i + k])
                else:
                    ax.axis("off")
            fig.tight_layout(rect=(0, 0, 1, 0.95))
            pdf.savefig(fig)
            done += 1
            if progress_cb:
                progress_cb(done, total)
        # Box Plot: 4×3
        for i in range(0, len(cols), 12):
            fig = Figure(figsize=A4L)
            fig.suptitle(title, fontsize=12)
            for k in range(12):
                ax = fig.add_subplot(4, 3, k + 1)
                if i + k < len(cols):
                    draw_box(ax, model, cols[i + k], stats_table=True)
                else:
                    ax.axis("off")
            # 각 subplot 아래 통계표 공간 확보
            fig.subplots_adjust(top=0.92, bottom=0.10, left=0.06, right=0.98,
                                hspace=1.15, wspace=0.35)
            pdf.savefig(fig)
            done += 1
            if progress_cb:
                progress_cb(done, total)


# ============================================================================
# 6. GUI
# ============================================================================
BaseTk = TkinterDnD.Tk if HAS_DND else tk.Tk


class App(BaseTk):
    def __init__(self):
        super().__init__()
        self.title("Discrete Reliability Data Analyzer")
        self.geometry("1200x800")
        self.model = DataModel()
        self.files = []
        self.selected_cols = []
        self.cur_idx = 0
        self._build_start()

    # ---- 화면 1: 파일 선택 -------------------------------------------------
    def _build_start(self):
        for w in self.winfo_children():
            w.destroy()
        frm = ttk.Frame(self, padding=20)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Discrete Reliability Data Analyzer",
                  font=("", 16, "bold")).pack(pady=10)
        ttk.Label(frm, text="파일 이름은 신뢰성명 + Lot번호 + Read-out 형식이어야 합니다.\n"
                            "예: HTRB_Lot1_0hr.csv, TC+Lot2+500cyc.xlsx").pack(pady=5)

        drop = tk.Label(frm, text="여기에 파일을 Drag && Drop 하세요"
                        if HAS_DND else "Browse 버튼으로 파일을 선택하세요",
                        relief="ridge", height=6, bg="#f0f0f0")
        drop.pack(fill="x", pady=10)
        if HAS_DND:
            drop.drop_target_register(DND_FILES)
            drop.dnd_bind("<<Drop>>", lambda e: self._add_files(self.tk.splitlist(e.data)))

        ttk.Button(frm, text="Browse...", command=self._browse).pack()
        self.file_list = tk.Listbox(frm, height=8)
        self.file_list.pack(fill="both", expand=True, pady=10)
        btns = ttk.Frame(frm)
        btns.pack()
        ttk.Button(btns, text="선택 제거", command=self._remove_file).pack(side="left", padx=5)
        ttk.Button(btns, text="다음 (파일 읽기)", command=self._load_files).pack(side="left", padx=5)

    def _browse(self):
        paths = filedialog.askopenfilenames(
            filetypes=[("Data files", "*.csv *.xlsx *.xlsm"), ("All", "*.*")])
        self._add_files(paths)

    def _add_files(self, paths):
        for p in paths:
            p = p.strip("{}")
            if p and p not in self.files:
                self.files.append(p)
                self.file_list.insert("end", os.path.basename(p))

    def _remove_file(self):
        for i in reversed(self.file_list.curselection()):
            self.file_list.delete(i)
            del self.files[i]

    def _load_files(self):
        if not self.files:
            messagebox.showwarning("알림", "파일을 먼저 선택하세요.")
            return
        errors = self.model.load(self.files)
        if errors:
            messagebox.showerror("파일 오류", "\n\n".join(errors))
        if not self.model.columns:
            return
        self._build_param_select()

    # ---- 화면 2: Parameter 선택 ---------------------------------------------
    def _build_param_select(self):
        for w in self.winfo_children():
            w.destroy()
        frm = ttk.Frame(self, padding=20)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text=f"{self.model.reliability}_{self.model.lot}  |  "
                            f"Read-out: {', '.join(self.model.readouts)}  |  "
                            f"Sample: {len(self.model.samples)}",
                  font=("", 11, "bold")).pack(pady=5)
        ttk.Label(frm, text="분석할 Parameter를 선택하세요 (Ctrl/Shift 다중 선택)").pack()
        self.param_lb = tk.Listbox(frm, selectmode="extended")
        for c in self.model.columns:
            self.param_lb.insert("end", c)
        self.param_lb.pack(fill="both", expand=True, pady=10)
        btns = ttk.Frame(frm)
        btns.pack()
        ttk.Button(btns, text="전체 선택",
                   command=lambda: self.param_lb.select_set(0, "end")).pack(side="left", padx=5)
        ttk.Button(btns, text="분석 시작", command=self._analyze).pack(side="left", padx=5)
        ttk.Button(btns, text="← 파일 다시 선택", command=self._build_start).pack(side="left", padx=5)
        self.pbar = ttk.Progressbar(frm, mode="determinate")
        self.pbar.pack(fill="x", pady=5)

    def _analyze(self):
        sel = self.param_lb.curselection()
        if not sel:
            messagebox.showwarning("알림", "Parameter를 선택하세요.")
            return
        self.selected_cols = [self.model.columns[i] for i in sel]
        self.cur_idx = 0
        # Progress 시뮬레이션(렌더 준비) 후 그래프 화면으로
        self.pbar["maximum"] = len(self.selected_cols)
        self.pbar["value"] = len(self.selected_cols)
        self.update_idletasks()
        messagebox.showinfo("완료", "Data 분석이 완료되었습니다.")
        self._build_graphs()

    # ---- 화면 3: 그래프 -------------------------------------------------------
    def _build_graphs(self):
        for w in self.winfo_children():
            w.destroy()
        top = ttk.Frame(self, padding=5)
        top.pack(fill="x")
        ttk.Button(top, text="← Parameter", command=self._build_param_select).pack(side="left")
        ttk.Button(top, text="◀ 이전", command=lambda: self._nav(-1)).pack(side="left", padx=3)
        self.param_var = tk.StringVar()
        cb = ttk.Combobox(top, textvariable=self.param_var,
                          values=self.selected_cols, width=45, state="readonly")
        cb.pack(side="left", padx=3)
        cb.bind("<<ComboboxSelected>>",
                lambda e: self._goto(self.selected_cols.index(self.param_var.get())))
        ttk.Button(top, text="다음 ▶", command=lambda: self._nav(1)).pack(side="left", padx=3)
        ttk.Separator(top, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Button(top, text="Y축 Min/Max", command=self._set_ylim).pack(side="left", padx=3)
        ttk.Button(top, text="Undo", command=self._undo).pack(side="left", padx=3)
        ttk.Button(top, text="Redo", command=self._redo).pack(side="left", padx=3)
        ttk.Button(top, text="Export PDF", command=self._export_pdf).pack(side="right", padx=3)

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True)
        self.line_tab = ttk.Frame(self.nb)
        self.box_tab = ttk.Frame(self.nb)
        self.nb.add(self.line_tab, text="Line Graph")
        self.nb.add(self.box_tab, text="Box Plot")

        self.line_fig = Figure(figsize=(10, 5))
        self.line_canvas = FigureCanvasTkAgg(self.line_fig, self.line_tab)
        self.line_canvas.get_tk_widget().pack(fill="both", expand=True)
        NavigationToolbar2Tk(self.line_canvas, self.line_tab)
        self.line_canvas.mpl_connect("pick_event", self._on_pick_line)

        self.box_fig = Figure(figsize=(10, 5))
        self.box_canvas = FigureCanvasTkAgg(self.box_fig, self.box_tab)
        self.box_canvas.get_tk_widget().pack(fill="both", expand=True)
        NavigationToolbar2Tk(self.box_canvas, self.box_tab)

        self._goto(0)

    def _nav(self, d):
        self._goto((self.cur_idx + d) % len(self.selected_cols))

    def _goto(self, idx):
        self.cur_idx = idx
        self.param_var.set(self.selected_cols[idx])
        self._redraw()

    def _redraw(self):
        col = self.selected_cols[self.cur_idx]
        self.line_fig.clear()
        ax = self.line_fig.add_subplot(111)
        self._line_artists = draw_line(ax, self.model, col, picker=True)
        self.line_fig.tight_layout()
        self.line_canvas.draw()

        self.box_fig.clear()
        ax2 = self.box_fig.add_subplot(111)
        draw_box(ax2, self.model, col, picker=False, stats_table=True)
        self.box_fig.subplots_adjust(top=0.9, bottom=0.35, left=0.12, right=0.95)
        self.box_canvas.draw()

    # ---- 편집 --------------------------------------------------------------
    def _on_pick_line(self, event):
        col = self.selected_cols[self.cur_idx]
        artist = event.artist
        readout = None
        for r, ln in self._line_artists.items():
            if ln is artist:
                readout = r
                break
        if readout is None or not len(event.ind):
            return
        sample = self.model.samples[event.ind[0]]
        self._point_menu(readout, col, sample)

    def _point_menu(self, readout, col, sample):
        m = tk.Menu(self, tearoff=0)
        m.add_command(label=f"Sample {sample} @ {readout}", state="disabled")
        m.add_separator()
        m.add_command(label="Marker 색 변경 (삼각형 표시)",
                      command=lambda: self._change_color(readout, col, sample))
        m.add_command(label="이 Read-out Marker만 삭제",
                      command=lambda: (self.model.delete_point(readout, col, sample),
                                       self._redraw()))
        m.add_command(label="Sample 전체 삭제 (모든 Read-out)",
                      command=lambda: (self.model.delete_sample_all_readouts(col, sample),
                                       self._redraw()))
        m.tk_popup(self.winfo_pointerx(), self.winfo_pointery())

    def _change_color(self, readout, col, sample):
        c = colorchooser.askcolor()[1]
        if c:
            self.model.set_color(readout, col, sample, c)
            self._redraw()

    def _set_ylim(self):
        col = self.selected_cols[self.cur_idx]
        cur = self.model.ylim.get(col, (None, None))
        ymin = simpledialog.askfloat("Y축", "Y Min (취소=자동):",
                                     initialvalue=cur[0])
        if ymin is None:
            self.model.set_ylim(col, None, None)
            self._redraw()
            return
        ymax = simpledialog.askfloat("Y축", "Y Max:", initialvalue=cur[1])
        if ymax is None or ymax <= ymin:
            messagebox.showwarning("알림", "Y Max는 Y Min보다 커야 합니다.")
            return
        self.model.set_ylim(col, ymin, ymax)
        self._redraw()

    def _undo(self):
        if self.model.undo():
            self._redraw()

    def _redo(self):
        if self.model.redo():
            self._redraw()

    # ---- PDF ------------------------------------------------------------------
    def _export_pdf(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            initialfile=f"{self.model.reliability}_{self.model.lot}_Data_Analysis.pdf",
            filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        win = tk.Toplevel(self)
        win.title("PDF 생성 중")
        win.geometry("360x90")
        ttk.Label(win, text="PDF Report 생성 중...").pack(pady=8)
        pbar = ttk.Progressbar(win, mode="determinate", length=320)
        pbar.pack(pady=5)

        def cb(done, total):
            pbar["maximum"] = total
            pbar["value"] = done
            win.update_idletasks()

        def work():
            try:
                export_pdf(self.model, self.selected_cols, path, cb)
                self.after(0, lambda: (win.destroy(),
                                       messagebox.showinfo("완료", f"PDF 저장 완료:\n{path}")))
            except Exception as e:
                err = traceback.format_exc()
                self.after(0, lambda: (win.destroy(),
                                       messagebox.showerror("오류", f"PDF 생성 실패:\n{e}\n\n{err}")))

        threading.Thread(target=work, daemon=True).start()


# ============================================================================
if __name__ == "__main__":
    app = App()
    app.mainloop()
