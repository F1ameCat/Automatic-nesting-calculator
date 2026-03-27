from __future__ import annotations

import bisect
import math
import tkinter as tk
from tkinter import ttk
from typing import Any, Optional, Sequence

from dxf_nesting_tab import DxfNestingTab


MIN_SHEET_SIDE = 150

# --- 折弯中性层因子：R/T   ---
RT_RATIO: tuple[float, ...] = (
    0.2,
    0.3,
    0.4,
    0.5,
    0.6,
    0.7,
    0.8,
    0.9,
    1.0,
    1.2,
    1.5,
    1.8,
    2.0,
    2.2,
    2.5,
    2.8,
    3.0,
    3.5,
    4.0,
    4.5,
    5.0,
    6.0,
    7.0,
    8.0,
    9.0,
    10.0,
)

# 钢 / 铜 / 铝 对应中性层因子 V
SHEET_NEUTRAL_STEEL: tuple[float, ...] = (
    0.248,
    0.28,
    0.303,
    0.32,
    0.333,
    0.343,
    0.352,
    0.36,
    0.366,
    0.376,
    0.388,
    0.396,
    0.4,
    0.403,
    0.408,
    0.413,
    0.415,
    0.419,
    0.423,
    0.425,
    0.427,
    0.431,
    0.434,
    0.437,
    0.438,
    0.44,
)
SHEET_NEUTRAL_COPPER: tuple[float, ...] = (
    0.26,
    0.293,
    0.318,
    0.334,
    0.347,
    0.358,
    0.366,
    0.374,
    0.38,
    0.392,
    0.402,
    0.41,
    0.415,
    0.419,
    0.423,
    0.427,
    0.429,
    0.431,
    0.438,
    0.441,
    0.442,
    0.446,
    0.449,
    0.45,
    0.452,
    0.453,
)
SHEET_NEUTRAL_ALUMINUM: tuple[float, ...] = (
    0.299,
    0.334,
    0.358,
    0.376,
    0.389,
    0.4,
    0.409,
    0.417,
    0.424,
    0.434,
    0.446,
    0.455,
    0.46,
    0.463,
    0.468,
    0.472,
    0.475,
    0.478,
    0.482,
    0.486,
    0.488,
    0.491,
    0.493,
    0.495,
    0.497,
    0.498,
)

NEUTRAL_SHEETS: tuple[tuple[float, ...], ...] = (
    SHEET_NEUTRAL_STEEL,
    SHEET_NEUTRAL_COPPER,
    SHEET_NEUTRAL_ALUMINUM,
)


def neutral_layer_factor_v(k: float, sheet: Sequence[float]) -> float:
    """
    根据 R/T（内径/厚度）在表中插值得到中性层因子 V。
    K 大于最大节点时取 0.5；K 小于最小节点时按首段线性外推。
    """
    rt = RT_RATIO
    if k > rt[-1]:
        return 0.5
    if k <= rt[0]:
        t = (k - rt[0]) / (rt[1] - rt[0])
        return float(sheet[0] + t * (sheet[1] - sheet[0]))
    i = bisect.bisect_right(rt, k) - 1
    if i >= len(rt) - 1:
        return float(sheet[-1])
    t = (k - rt[i]) / (rt[i + 1] - rt[i])
    return float(sheet[i] + t * (sheet[i + 1] - sheet[i]))


def maximize_toplevel(win: tk.Toplevel | tk.Tk) -> None:
    """尽量以最大化显示主窗口（Windows 为 zoomed；其它平台尝试等价方式）。"""
    win.update_idletasks()
    try:
        win.state("zoomed")
    except tk.TclError:
        try:
            win.attributes("-zoomed", True)
        except tk.TclError:
            w = win.winfo_screenwidth()
            h = win.winfo_screenheight()
            win.geometry(f"{w}x{h}+0+0")


def safe_int(value: str, default: int) -> int:
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return default


def max_parts_along_axis(inner_len: int, part_len: int, gap_between: int) -> int:
    """
    沿单轴可密铺的最多零件个数。
    inner_len：该方向净长（板长减去两侧「板边间隙」）。
    gap_between：相邻零件之间的间隙（可为 0）。
    """
    if inner_len < part_len:
        return 0
    step = part_len + gap_between
    if step <= 0:
        return 0
    return (inner_len + gap_between) // step


def tail_board_size_mm(
    cols: int,
    rows: int,
    part_w: int,
    part_h: int,
    gap_edge: int,
    gap_part: int,
) -> tuple[int, int]:
    """给定列数、行数时，含两侧板边与零件间隙的板料宽、高（mm）。"""
    w = 2 * gap_edge + cols * part_w + max(0, cols - 1) * gap_part
    h = 2 * gap_edge + rows * part_h + max(0, rows - 1) * gap_part
    return w, h


class CuttingCalculatorTab(ttk.Frame):
    """密铺下料计算与预览（作为 Notebook 中的一个选项卡）。"""

    _LEFT_OUT_PAD = 8
    _LEFT_LBL_W = 9
    _LEFT_ENT_W = 9

    def __init__(self, parent: tk.Misc, **kwargs: Any) -> None:
        super().__init__(parent, **kwargs)

        self.vars = {
            "part_w": tk.StringVar(value="58"),
            "part_h": tk.StringVar(value="96"),
            "qty": tk.StringVar(value="685"),
            "gap_edge": tk.StringVar(value="25"),
            "gap_part": tk.StringVar(value="5"),
            "sheet_w": tk.StringVar(value="1200"),
            "sheet_h": tk.StringVar(value="2000"),
        }

        self.result_var = tk.StringVar(
            value="请输入参数；按 Enter 或离开输入框后再计算。"
        )
        self.detail_var = tk.StringVar(value="")

        # 每块预览画布独立：缩放、平移（相对“居中适配”原点的偏移，像素）
        self.zoom_full = 1.0
        self.zoom_last = 1.0
        self.pan_full: list[float] = [0.0, 0.0]
        self.pan_last: list[float] = [0.0, 0.0]

        self._layout: Optional[dict[str, Any]] = None
        self._drag: Optional[dict[str, Any]] = None

        self._build_ui()
        self._bind_events()
        self.recalculate()

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=12)
        outer.pack(fill="both", expand=True)
        ttk.Label(
            outer,
            text="自动排料计算器",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", pady=(0, 6))

        # main 必须放在 outer 内，否则与 outer 并列 expand 会把标题栏撑出大片空白
        main = ttk.Frame(outer)
        main.pack(fill="both", expand=True)

        left = ttk.Frame(main, padding=(0, 0, self._LEFT_OUT_PAD, 0))
        left.pack(side="left", fill="y")

        right = ttk.Frame(main)
        right.pack(side="right", fill="both", expand=True)

        group = ttk.LabelFrame(left, text="输入参数（单位：mm）", padding=6)
        group.pack(side="top", fill="x")

        self._add_entry(group, "零件宽", "part_w")
        self._add_entry(group, "零件高", "part_h")
        swap_row = ttk.Frame(group)
        swap_row.pack(fill="x", pady=(2, 6))
        ttk.Button(
            swap_row,
            text="宽↔高",
            command=self._swap_part_dimensions,
        ).pack(side="right")
        self._add_entry(group, "零件数量", "qty")
        self._add_entry(group, "板边间隙", "gap_edge")
        self._add_entry(group, "零件间隙", "gap_part")
        self._add_entry(group, "板材宽", "sheet_w")
        self._add_entry(group, "板材高", "sheet_h")

        lower = ttk.Frame(left)
        lower.pack(fill="both", expand=True, pady=(8, 0))
        lower.rowconfigure(0, weight=1)
        lower.rowconfigure(1, weight=1)
        lower.columnconfigure(0, weight=1)

        result_lf = ttk.LabelFrame(lower, text="计算结果", padding=(6, 6))
        result_lf.grid(row=0, column=0, sticky="nsew", pady=(0, 6))
        result_lf.rowconfigure(0, weight=1)
        result_lf.columnconfigure(0, weight=1)
        self._result_text = tk.Text(
            result_lf,
            height=3,
            wrap="word",
            font=("Segoe UI", 9),
            relief="flat",
            highlightthickness=0,
            padx=4,
            pady=4,
            takefocus=False,
        )
        self._result_text.tag_configure("head", font=("Segoe UI", 9, "bold"))
        rsb = ttk.Scrollbar(result_lf, command=self._result_text.yview)
        self._result_text.configure(yscrollcommand=rsb.set)
        self._result_text.grid(row=0, column=0, sticky="nsew")
        rsb.grid(row=0, column=1, sticky="ns")

        help_lf = ttk.LabelFrame(lower, text="说明", padding=(6, 6))
        help_lf.grid(row=1, column=0, sticky="nsew")
        help_lf.rowconfigure(0, weight=1)
        help_lf.columnconfigure(0, weight=1)
        help_txt = tk.Text(
            help_lf,
            height=3,
            wrap="word",
            font=("Segoe UI", 9),
            relief="flat",
            highlightthickness=0,
            padx=4,
            pady=4,
            takefocus=False,
        )
        help_sb = ttk.Scrollbar(help_lf, command=help_txt.yview)
        help_txt.configure(yscrollcommand=help_sb.set)
        help_txt.grid(row=0, column=0, sticky="nsew")
        help_sb.grid(row=0, column=1, sticky="ns")
        help_txt.insert(
            "1.0",
            "「板边间隙」为零件到板边的距离；「零件间隙」为零件与零件之间距离。\n"
            "尾板宽和高都至少为 150mm（设备限制）。\n"
            "示意图仅用于验算展示，不代表切割路径。\n"
            "右侧预览框内：滚轮缩放，左键拖动平移；改数字后按 Enter 或离开输入框再计算。",
        )
        help_txt.configure(state="disabled")

        # 右侧：工具栏 + 双画布（子画布自带裁剪，放大后内容不会画出边框外）
        toolbar = ttk.Frame(right)
        toolbar.pack(fill="x", pady=(0, 8))
        ttk.Button(
            toolbar,
            text="重置显示比例",
            command=self._reset_view,
        ).pack(side="right")

        preview = ttk.Frame(right)
        preview.pack(fill="both", expand=True)
        preview.columnconfigure(0, weight=1)
        preview.columnconfigure(1, weight=1)
        preview.rowconfigure(0, weight=1)

        col_full = ttk.Frame(preview)
        col_full.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        col_last = ttk.Frame(preview)
        col_last.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        self.lbl_full_title = ttk.Label(col_full, font=("Segoe UI", 11, "bold"))
        self.lbl_full_title.grid(row=0, column=0, sticky="w")
        self.lbl_last_title = ttk.Label(col_last, font=("Segoe UI", 11, "bold"))
        self.lbl_last_title.grid(row=0, column=0, sticky="w")

        self.lbl_full_sub = ttk.Label(col_full, font=("Segoe UI", 9), foreground="#555")
        self.lbl_full_sub.grid(row=1, column=0, sticky="w", pady=(0, 4))
        self.lbl_last_sub = ttk.Label(col_last, font=("Segoe UI", 9), foreground="#555")
        self.lbl_last_sub.grid(row=1, column=0, sticky="w", pady=(0, 4))

        # 使用 Frame 包一层带边框的容器；子 Canvas 只在内侧绘制，放大后内容会被窗口裁剪在边框内
        frame_cv_full = tk.Frame(col_full, bg="#505050", padx=2, pady=2)
        frame_cv_full.grid(row=2, column=0, sticky="nsew")
        col_full.rowconfigure(2, weight=1)
        self.canvas_full = tk.Canvas(frame_cv_full, bg="#f7f7f7", highlightthickness=0)
        self.canvas_full.pack(fill="both", expand=True)

        frame_cv_last = tk.Frame(col_last, bg="#505050", padx=2, pady=2)
        frame_cv_last.grid(row=2, column=0, sticky="nsew")
        col_last.rowconfigure(2, weight=1)
        self.canvas_last = tk.Canvas(frame_cv_last, bg="#f7f7f7", highlightthickness=0)
        self.canvas_last.pack(fill="both", expand=True)

    def _add_entry(self, parent: ttk.Widget, label: str, key: str) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=4)
        ttk.Label(row, text=label, width=self._LEFT_LBL_W, anchor="w").pack(
            side="left"
        )
        ent = ttk.Entry(row, textvariable=self.vars[key], width=self._LEFT_ENT_W)
        ent.pack(side="right")
        ent.bind("<FocusOut>", lambda _e: self.recalculate(), add="+")
        ent.bind("<Return>", lambda _e: self.recalculate(), add="+")

    def _bind_events(self) -> None:
        self.canvas_full.bind("<Configure>", lambda *_: self._on_preview_configure())
        self.canvas_last.bind("<Configure>", lambda *_: self._on_preview_configure())

        for key, cv in (("full", self.canvas_full), ("last", self.canvas_last)):
            cv.bind("<MouseWheel>", lambda e, k=key: self._on_mousewheel(k, e))
            cv.bind("<Button-4>", lambda e, k=key: self._apply_zoom(k, 120))
            cv.bind("<Button-5>", lambda e, k=key: self._apply_zoom(k, -120))
            cv.bind("<ButtonPress-1>", lambda e, k=key: self._on_drag_press(k, e))
            cv.bind("<B1-Motion>", self._on_drag_motion)
            cv.bind("<ButtonRelease-1>", self._on_drag_release)

    def _on_preview_configure(self) -> None:
        if self._layout is not None:
            self._draw_previews()

    def _on_mousewheel(self, key: str, event: tk.Event) -> str:
        self._apply_zoom(key, event.delta)
        return "break"

    def _apply_zoom(self, key: str, delta: int) -> None:
        if self._layout is None:
            return
        direction = 1 if delta > 0 else -1
        factor = 1.15 if direction > 0 else 1 / 1.15
        if key == "full":
            self.zoom_full = min(max(self.zoom_full * factor, 0.25), 8.0)
        else:
            self.zoom_last = min(max(self.zoom_last * factor, 0.25), 8.0)
        self._draw_previews()

    def _on_drag_press(self, key: str, event: tk.Event) -> None:
        if self._layout is None:
            return
        self._drag = {"key": key, "x": event.x, "y": event.y}

    def _on_drag_motion(self, event: tk.Event) -> None:
        if self._drag is None or self._layout is None:
            return
        key = self._drag["key"]
        dx = event.x - self._drag["x"]
        dy = event.y - self._drag["y"]
        self._drag["x"] = event.x
        self._drag["y"] = event.y
        pan = self.pan_full if key == "full" else self.pan_last
        pan[0] += dx
        pan[1] += dy
        self._draw_previews()

    def _on_drag_release(self, _event: tk.Event) -> None:
        self._drag = None

    def _reset_view(self) -> None:
        self.zoom_full = 1.0
        self.zoom_last = 1.0
        self.pan_full[0] = self.pan_full[1] = 0.0
        self.pan_last[0] = self.pan_last[1] = 0.0
        if self._layout is not None:
            self._draw_previews()
        else:
            self.recalculate()

    def _swap_part_dimensions(self) -> None:
        w = self.vars["part_w"].get()
        h = self.vars["part_h"].get()
        self.vars["part_w"].set(h)
        self.vars["part_h"].set(w)
        self.recalculate()

    def _sync_cutting_result_text(self) -> None:
        t = getattr(self, "_result_text", None)
        if t is None:
            return
        head = self.result_var.get()
        detail = self.detail_var.get()
        t.configure(state="normal")
        t.delete("1.0", "end")
        if head:
            t.insert("1.0", head, ("head",))
        if detail:
            t.insert("end", ("\n\n" if head else "") + detail)
        t.configure(state="disabled")

    def recalculate(self) -> None:
        try:
            part_w = safe_int(self.vars["part_w"].get(), 0)
            part_h = safe_int(self.vars["part_h"].get(), 0)
            qty = safe_int(self.vars["qty"].get(), 0)
            gap_edge = safe_int(self.vars["gap_edge"].get(), 0)
            gap_part = safe_int(self.vars["gap_part"].get(), 0)
            sheet_w = safe_int(self.vars["sheet_w"].get(), 0)
            sheet_h = safe_int(self.vars["sheet_h"].get(), 0)

            self.canvas_full.delete("all")
            self.canvas_last.delete("all")
            self._layout = None
            self.lbl_full_title.config(text="")
            self.lbl_full_sub.config(text="")
            self.lbl_last_title.config(text="")
            self.lbl_last_sub.config(text="")

            if min(part_w, part_h, qty, sheet_w, sheet_h) <= 0:
                self.result_var.set("请输入大于 0 的数值。")
                self.detail_var.set("")
                return
            if gap_edge < 0 or gap_part < 0:
                self.result_var.set("板边间隙、零件间隙不能为负数。")
                self.detail_var.set("")
                return

            inner_w = sheet_w - 2 * gap_edge
            inner_h = sheet_h - 2 * gap_edge
            cols = max_parts_along_axis(inner_w, part_w, gap_part)
            rows = max_parts_along_axis(inner_h, part_h, gap_part)
            if cols <= 0 or rows <= 0:
                self.result_var.set(
                    "目标板材太小，无法放下 1 个零件（已考虑边距和间距）。"
                )
                self.detail_var.set(
                    f"当前最多列数={max(cols, 0)}，行数={max(rows, 0)}"
                )
                return

            per_sheet = cols * rows
            full_count = qty // per_sheet
            remain = qty % per_sheet

            if remain == 0:
                total_sheets = full_count
                last_w, last_h = sheet_w, sheet_h
                last_cols, last_rows, last_cap = cols, rows, per_sheet
                last_used = per_sheet
            else:
                total_sheets = full_count + 1
                (last_w, last_h), (last_cols, last_rows, last_cap) = (
                    self._best_last_sheet(
                        remain, part_w, part_h, gap_edge, gap_part
                    )
                )
                last_used = remain

            self.result_var.set(
                f"共需要板材：{total_sheets} 张\n"
                f"其中完整目标板：{full_count} 张\n"
                f"最后一张尺寸：{last_w} x {last_h} mm"
            )
            self.detail_var.set(
                f"目标板单张可排：{cols} 列 x {rows} 行 = {per_sheet} 件\n"
                f"尾板排版：{last_cols} 列 x {last_rows} 行，可放 {last_cap} 件，实际放 {last_used} 件"
            )

            self.lbl_full_title.config(text=f"完整目标板（{sheet_w} x {sheet_h}）")
            self.lbl_full_sub.config(
                text=f"每张可放 {per_sheet} 件，完整板数量 {full_count} 张"
            )
            self.lbl_last_title.config(text=f"最后一张（{last_w} x {last_h}）")
            self.lbl_last_sub.config(
                text=f"实际放 {last_used} 件，总板数 {total_sheets} 张"
            )

            self._layout = {
                "sheet_w": sheet_w,
                "sheet_h": sheet_h,
                "cols": cols,
                "rows": rows,
                "used_full": per_sheet,
                "last_w": last_w,
                "last_h": last_h,
                "last_cols": last_cols,
                "last_rows": last_rows,
                "used_last": last_used,
                "part_w": part_w,
                "part_h": part_h,
                "gap_edge": gap_edge,
                "gap_part": gap_part,
            }
            self._draw_previews()
        finally:
            self._sync_cutting_result_text()

    def _best_last_sheet(
        self,
        remain: int,
        part_w: int,
        part_h: int,
        gap_edge: int,
        gap_part: int,
    ) -> tuple[tuple[int, int], tuple[int, int, int]]:
        best = None
        for c in range(1, remain + 1):
            r = math.ceil(remain / c)
            req_w, req_h = tail_board_size_mm(c, r, part_w, part_h, gap_edge, gap_part)
            final_w = max(req_w, MIN_SHEET_SIDE)
            final_h = max(req_h, MIN_SHEET_SIDE)
            area = final_w * final_h
            score = (area, final_w + final_h, final_w)
            item = (score, final_w, final_h, c, r)
            if best is None or item[0] < best[0]:
                best = item

        assert best is not None
        _, final_w, final_h, best_cols, best_rows = best
        capacity = best_cols * best_rows
        return (final_w, final_h), (best_cols, best_rows, capacity)

    def _draw_previews(self) -> None:
        if self._layout is None:
            return
        L = self._layout
        self._draw_single_panel(
            self.canvas_full,
            zoom=self.zoom_full,
            pan=self.pan_full,
            sheet_w=L["sheet_w"],
            sheet_h=L["sheet_h"],
            cols=L["cols"],
            rows=L["rows"],
            used=L["used_full"],
            part_w=L["part_w"],
            part_h=L["part_h"],
            gap_edge=L["gap_edge"],
            gap_part=L["gap_part"],
        )
        self._draw_single_panel(
            self.canvas_last,
            zoom=self.zoom_last,
            pan=self.pan_last,
            sheet_w=L["last_w"],
            sheet_h=L["last_h"],
            cols=L["last_cols"],
            rows=L["last_rows"],
            used=L["used_last"],
            part_w=L["part_w"],
            part_h=L["part_h"],
            gap_edge=L["gap_edge"],
            gap_part=L["gap_part"],
        )

    # 画布内边距：为 H=（左侧竖排）、W=（板下方）尺寸标注留出空间，避免默认缩放下被裁切
    _PREVIEW_PAD_L = 56
    _PREVIEW_PAD_R = 12
    _PREVIEW_PAD_T = 12
    _PREVIEW_PAD_B = 40

    def _draw_single_panel(
        self,
        canvas: tk.Canvas,
        zoom: float,
        pan: list[float],
        sheet_w: int,
        sheet_h: int,
        cols: int,
        rows: int,
        used: int,
        part_w: int,
        part_h: int,
        gap_edge: int,
        gap_part: int,
    ) -> None:
        canvas.delete("all")
        cw = max(canvas.winfo_width(), 120)
        ch = max(canvas.winfo_height(), 120)

        pl = self._PREVIEW_PAD_L
        pr = self._PREVIEW_PAD_R
        pt = self._PREVIEW_PAD_T
        pb = self._PREVIEW_PAD_B
        avail_w = max(cw - pl - pr, 40)
        avail_h = max(ch - pt - pb, 40)
        scale_fit = min(avail_w / sheet_w, avail_h / sheet_h)
        scale = scale_fit * zoom
        draw_w = sheet_w * scale
        draw_h = sheet_h * scale
        ox = pl + (avail_w - draw_w) / 2 + pan[0]
        oy = pt + (avail_h - draw_h) / 2 + pan[1]

        canvas.create_rectangle(
            ox, oy, ox + draw_w, oy + draw_h, fill="#fffef6", outline="#303030", width=2
        )

        idx = 0
        for r in range(rows):
            for c in range(cols):
                if idx >= used:
                    break
                px = ox + (gap_edge + c * (part_w + gap_part)) * scale
                py = oy + (gap_edge + r * (part_h + gap_part)) * scale
                pw = part_w * scale
                ph = part_h * scale
                canvas.create_rectangle(
                    px, py, px + pw, py + ph, fill="#9fd3ff", outline="#1a4d7a"
                )
                if pw >= 44 and ph >= 22:
                    canvas.create_text(
                        px + pw / 2,
                        py + ph / 2,
                        text=f"{part_w}x{part_h}",
                        font=("Segoe UI", max(7, int(8 * scale))),
                    )
                idx += 1
            if idx >= used:
                break

        canvas.create_text(
            ox + draw_w / 2,
            oy + draw_h + 14,
            text=f"W={sheet_w}",
            fill="#111",
            font=("Segoe UI", 9, "bold"),
        )
        canvas.create_text(
            ox - 20,
            oy + draw_h / 2,
            text=f"H={sheet_h}",
            angle=90,
            fill="#111",
            font=("Segoe UI", 9, "bold"),
        )


class NeutralLayerFactorTab(ttk.Frame):
    """金属材料折弯中性层因子：K=R/T 查表插值 + 曲线示意。"""

    _K_AXIS_MAX = 10.0
    _V_AXIS_MIN = 0.22
    _V_AXIS_MAX = 0.52

    def __init__(self, parent: tk.Misc, **kwargs: Any) -> None:
        super().__init__(parent, **kwargs)
        self.mat_var = tk.IntVar(value=2)
        self.var_r = tk.StringVar(value="")
        self.var_t = tk.StringVar(value="")
        self.var_k = tk.StringVar(value="—")
        self.var_v = tk.StringVar(value="—")
        self._k_cur: float | None = None
        self._v_cur: float | None = None

        self._build_ui()
        self.mat_var.trace_add("write", lambda *_: self._refresh())
        self.var_r.trace_add("write", lambda *_: self._refresh())
        self.var_t.trace_add("write", lambda *_: self._refresh())
        self.chart.bind("<Configure>", lambda *_: self._draw_chart())

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=12)
        outer.pack(fill="both", expand=True)

        ttk.Label(
            outer,
            text="金属材料折弯中性层因子计算器",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", pady=(0, 12))

        body = ttk.Frame(outer)
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        left = ttk.Frame(body, padding=(0, 0, 12, 0))
        left.grid(row=0, column=0, sticky="nsew")

        group = ttk.LabelFrame(left, text="输入参数（单位：mm）", padding=10)
        group.pack(fill="x")

        mat_row = ttk.Frame(group)
        mat_row.pack(fill="x", pady=4)
        ttk.Label(mat_row, text="材料类型", width=18, anchor="w").pack(side="left")
        rb_fr = ttk.Frame(mat_row)
        rb_fr.pack(side="right")
        for i, name in enumerate(("钢 ST", "铜 CU", "铝 AL")):
            ttk.Radiobutton(
                rb_fr,
                text=name,
                variable=self.mat_var,
                value=i,
            ).pack(side="left", padx=(0, 8))

        self._add_labeled_entry(group, "内径 R", self.var_r)
        self._add_labeled_entry(group, "厚度 T", self.var_t)

        ttk.Label(
            left,
            text="说明：\n"
            "- K 超出表格上限时按原表规则取 V=0.5；\n"
            "- 图表纵轴为 V，横轴为 K（内径/厚度）。\n"
            "- 数据来源于《航空制造手册 飞机钣金工艺》\n第 215 页表 10.8（V 形件中性层系数）。",
            justify="left",
            wraplength=300,
        ).pack(fill="x", pady=(12, 8))

        out_fr = ttk.LabelFrame(left, text="计算结果", padding=10)
        out_fr.pack(fill="x")
        ttk.Label(out_fr, text="R / T = K", foreground="#333").pack(anchor="w")
        ttk.Label(
            out_fr,
            textvariable=self.var_k,
            justify="left",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w")
        ttk.Label(out_fr, text="中性层因子 V", foreground="#333").pack(anchor="w", pady=(8, 0))
        ttk.Label(
            out_fr,
            textvariable=self.var_v,
            justify="left",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w")

        right = ttk.LabelFrame(body, text="曲线图", padding=10)
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)

        border = tk.Frame(right, bg="#505050", padx=2, pady=2)
        border.grid(row=0, column=0, sticky="nsew")
        self.chart = tk.Canvas(border, bg="#f7f7f7", highlightthickness=0)
        self.chart.pack(fill="both", expand=True)

    def _add_labeled_entry(
        self, parent: ttk.Widget, label: str, var: tk.StringVar
    ) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=4)
        ttk.Label(row, text=label, width=18, anchor="w").pack(side="left")
        ttk.Entry(row, textvariable=var, width=16).pack(side="right")

    @staticmethod
    def _parse_positive_float(s: str) -> float | None:
        s = s.strip()
        if not s:
            return None
        try:
            x = float(s)
        except ValueError:
            return None
        if x <= 0:
            return None
        return x

    def _refresh(self) -> None:
        r = self._parse_positive_float(self.var_r.get())
        t = self._parse_positive_float(self.var_t.get())
        if r is None or t is None:
            self.var_k.set("—")
            self.var_v.set("—")
            self._k_cur = None
            self._v_cur = None
            self._draw_chart()
            return
        k = r / t
        mat = self.mat_var.get()
        if mat not in (0, 1, 2):
            mat = 2
        v = neutral_layer_factor_v(k, NEUTRAL_SHEETS[mat])
        self.var_k.set(f"{k:.3f}")
        self.var_v.set(f"{v:.3f}")
        self._k_cur = k
        self._v_cur = v
        self._draw_chart()

    def _data_to_canvas(
        self, k: float, v: float, plot: tuple[float, float, float, float]
    ) -> tuple[float, float]:
        pl, pt, pr, pb = plot
        pw = pr - pl
        ph = pb - pt
        x = pl + (k / self._K_AXIS_MAX) * pw
        y = pb - (v - self._V_AXIS_MIN) / (self._V_AXIS_MAX - self._V_AXIS_MIN) * ph
        return x, y

    def _draw_chart(self) -> None:
        c = self.chart
        c.delete("all")
        cw = max(c.winfo_width(), 400)
        ch = max(c.winfo_height(), 280)
        pl, pt, pr, pb = 52.0, 16.0, float(cw - 16), float(ch - 36)
        plot = (pl, pt, pr, pb)

        c.create_rectangle(pl, pt, pr, pb, outline="#b0c4c4", width=1, fill="#fffef6")

        # 网格与刻度（字号与下料预览画布标注一致）
        tick_font = ("Segoe UI", 9)
        for kk in (0, 2, 4, 6, 8, 10):
            x, _ = self._data_to_canvas(float(kk), self._V_AXIS_MIN, plot)
            c.create_line(x, pt, x, pb, fill="#ddeeee", width=1)
            c.create_text(x, pb + 14, text=str(kk), font=tick_font, fill="#333")
        for vv in (0.25, 0.3, 0.35, 0.4, 0.45, 0.5):
            _, y = self._data_to_canvas(0.0, vv, plot)
            c.create_line(pl, y, pr, y, fill="#ddeeee", width=1)
            c.create_text(pl - 12, y, text=f"{vv:.2f}", font=tick_font, fill="#333")

        c.create_line(pl, pb, pr, pb, fill="#cc3333", width=2)
        c.create_line(pl, pt, pl, pb, fill="#33aa33", width=2)
        axis_font = ("Segoe UI", 9, "bold")
        c.create_text(pr - 4, pb + 30, text="K = R/T", anchor="e", font=axis_font, fill="#333")
        c.create_text(pl - 4, pt - 6, text="V", anchor="w", font=axis_font, fill="#333")

        colors_sel = ("#1a6fd4", "#ffd500", "#777777")
        colors_dim = ("#9dbfe8", "#f0c09a", "#a8a8a8")
        sel = self.mat_var.get()
        if sel not in (0, 1, 2):
            sel = 2

        for mi, row in enumerate(NEUTRAL_SHEETS):
            pts: list[float] = []
            for i, rk in enumerate(RT_RATIO):
                vk = float(row[i])
                x, y = self._data_to_canvas(rk, vk, plot)
                pts.extend((x, y))
            active = mi == sel
            fill = colors_sel[mi] if active else colors_dim[mi]
            width = 3 if active else 1
            if len(pts) >= 4:
                c.create_line(*pts, fill=fill, width=width, smooth=False)

        lx, ly = pr - 118, pt + 6
        for mi, name in enumerate(("钢", "铜", "铝")):
            yy = ly + mi * 16
            c.create_line(lx, yy, lx + 20, yy, fill=colors_sel[mi], width=3 if mi == sel else 1)
            c.create_text(lx + 26, yy, text=name, anchor="w", font=("Segoe UI", 9), fill="#333")

        if self._k_cur is not None and self._v_cur is not None:
            x0, y0 = self._data_to_canvas(self._k_cur, self._v_cur, plot)
            xb, yb = self._data_to_canvas(self._k_cur, self._V_AXIS_MIN, plot)
            xl, yl = self._data_to_canvas(0.0, self._v_cur, plot)
            c.create_line(x0, y0, xb, yb, fill="#aa2222", width=1, dash=(4, 3))
            c.create_line(x0, y0, xl, yl, fill="#228822", width=1, dash=(4, 3))
            r = 5
            c.create_oval(x0 - r, y0 - r, x0 + r, y0 + r, outline="#c00000", width=2, fill="#ffeeee")
            c.create_text(
                x0 + 8,
                y0 - 18,
                text=f"({self._k_cur:.3f}, {self._v_cur:.3f})",
                anchor="w",
                font=("Segoe UI", 9),
                fill="#111",
            )


def main() -> None:
    root = tk.Tk()
    root.title("简易工程计算器")
    root.minsize(980, 700)

    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True)

    tab_dxf = DxfNestingTab(notebook, padding=0)
    notebook.add(tab_dxf, text="下料（DXF轮廓）")

    tab_cutting = CuttingCalculatorTab(notebook, padding=0)
    notebook.add(tab_cutting, text="下料计算器（原型）")

    tab_neutral = NeutralLayerFactorTab(notebook, padding=0)
    notebook.add(tab_neutral, text="中心层因子")

    maximize_toplevel(root)

    root.mainloop()


if __name__ == "__main__":
    main()
