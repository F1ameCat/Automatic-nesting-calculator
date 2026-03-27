from __future__ import annotations

import math
import queue
import threading
import tkinter as tk
from tkinter import filedialog, ttk
from typing import Any, Optional

import dxf_outline as dx

MIN_SHEET_SIDE = 150


def _safe_int(value: str, default: int) -> int:
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return default


class DxfNestingTab(ttk.Frame):
    """根据 DXF 轮廓密铺排料（多种密铺 + 真实多边形绘制）。"""

    _LEFT_OUT_PAD = 8
    _LEFT_LBL_W = 9
    _LEFT_ENT_W = 9

    _PACK_MODES: list[tuple[str, str]] = [
        (dx.PACKING_MODE_INTERLOCK_COL, "列向180°互嵌（省料）"),
        (dx.PACKING_MODE_GRID, "标准网格"),
        (dx.PACKING_MODE_COMPACT, "紧凑（凹槽互嵌）"),
        (dx.PACKING_MODE_BRICK, "交错行"),
    ]

    _PREVIEW_PAD_L = 56
    _PREVIEW_PAD_R = 12
    _PREVIEW_PAD_T = 12
    _PREVIEW_PAD_B = 40

    def __init__(self, parent: tk.Misc, **kwargs: Any) -> None:
        super().__init__(parent, **kwargs)

        self.vars = {
            "dxf_path": tk.StringVar(value=""),
            "gap_edge": tk.StringVar(value="25"),
            "gap_part": tk.StringVar(value="5"),
            "sheet_w": tk.StringVar(value="1200"),
            "sheet_h": tk.StringVar(value="2000"),
            "qty": tk.StringVar(value="10"),
            "packing_mode": tk.StringVar(value=dx.PACKING_MODE_INTERLOCK_COL),
            "fiber_deg": tk.IntVar(value=0),
        }
        self.result_var = tk.StringVar(
            value="请选择 DXF；改数字后按 Enter 或离开输入框再计算。"
        )
        self.detail_var = tk.StringVar(value="")

        self._layout: Optional[dict[str, Any]] = None
        self.zoom_full = 1.0
        self.zoom_last = 1.0
        self.pan_full: list[float] = [0.0, 0.0]
        self.pan_last: list[float] = [0.0, 0.0]
        self._drag: Optional[dict[str, Any]] = None
        self._dxf_recalc_gen = 0
        self._active_busy_win: Optional[tk.Toplevel] = None

        self.lbl_full_title: ttk.Label
        self.lbl_full_sub: ttk.Label
        self.lbl_last_title: ttk.Label
        self.lbl_last_sub: ttk.Label
        self.canvas_full: tk.Canvas
        self.canvas_last: tk.Canvas

        self._build_ui()
        if dx.deps_available():
            self._bind_events()

    def _build_ui(self) -> None:
        if not dx.deps_available():
            ttk.Label(
                self,
                text="本页需要额外依赖，请在项目目录执行：\npip install ezdxf shapely\n安装后重新打开程序。",
                justify="left",
                padding=24,
            ).pack(anchor="nw")
            return

        outer = ttk.Frame(self, padding=12)
        outer.pack(fill="both", expand=True)
        ttk.Label(
            outer,
            text="自动排料（DXF 轮廓）",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", pady=(0, 6))

        main = ttk.Frame(outer)
        main.pack(fill="both", expand=True)

        left = ttk.Frame(main, padding=(0, 0, self._LEFT_OUT_PAD, 0))
        left.pack(side="left", fill="y")
        right = ttk.Frame(main)
        right.pack(side="right", fill="both", expand=True)

        group = ttk.LabelFrame(left, text="DXF 与参数（单位：mm）", padding=6)
        group.pack(side="top", fill="x")

        row = ttk.Frame(group)
        row.pack(fill="x", pady=4)
        ttk.Label(row, text="DXF", width=self._LEFT_LBL_W, anchor="w").pack(
            side="left"
        )
        ent_dxf = ttk.Entry(row, textvariable=self.vars["dxf_path"], width=14)
        ent_dxf.pack(side="left", fill="x", expand=True)
        ent_dxf.bind("<FocusOut>", lambda _e: self._recalculate(), add="+")
        ent_dxf.bind("<Return>", lambda _e: self._recalculate(), add="+")
        ttk.Button(row, text="浏览…", command=self._browse_dxf).pack(
            side="right", padx=(6, 0)
        )

        self._add_entry(group, "板边间隙", "gap_edge")
        self._add_entry(group, "零件间隙", "gap_part")
        self._add_entry(group, "板材宽", "sheet_w")
        self._add_entry(group, "板材高", "sheet_h")
        self._add_entry(group, "零件数量", "qty")

        row_fb = ttk.Frame(group)
        row_fb.pack(fill="x", pady=4)
        ttk.Label(row_fb, text="纤维°", width=self._LEFT_LBL_W, anchor="w").pack(
            side="left"
        )
        ttk.Label(row_fb, textvariable=self.vars["fiber_deg"], width=3).pack(
            side="left", padx=(2, 0)
        )
        ttk.Button(
            row_fb,
            text="+90°",
            command=self._fiber_add_90,
        ).pack(side="right", padx=(4, 0))

        row_pk = ttk.Frame(group)
        row_pk.pack(fill="x", pady=4)
        ttk.Label(row_pk, text="密铺", width=self._LEFT_LBL_W, anchor="w").pack(
            side="left"
        )
        sub_pk = ttk.Frame(row_pk)
        sub_pk.pack(side="left", fill="x", expand=True)
        self._packing_combo = ttk.Combobox(
            sub_pk,
            state="readonly",
            width=11,
            values=[lab for _, lab in self._PACK_MODES],
        )
        self._packing_combo.pack(side="left", fill="x", expand=True)
        self._packing_combo.set(self._PACK_MODES[0][1])  # 默认：列向180°互嵌
        self._packing_combo.bind("<<ComboboxSelected>>", self._on_packing_combo)
        ttk.Button(sub_pk, text="下一", command=self._cycle_packing_mode).pack(
            side="right", padx=(4, 0)
        )

        lower = ttk.Frame(left)
        lower.pack(fill="both", expand=True, pady=(8, 0))
        lower.rowconfigure(0, weight=1)
        lower.rowconfigure(1, weight=1)
        lower.columnconfigure(0, weight=1)

        result_lf = ttk.LabelFrame(lower, text="计算结果", padding=(6, 6))
        result_lf.grid(row=0, column=0, sticky="nsew", pady=(0, 6))
        result_lf.rowconfigure(0, weight=1)
        result_lf.columnconfigure(0, weight=1)
        self._dxf_result_text = tk.Text(
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
        self._dxf_result_text.tag_configure("head", font=("Segoe UI", 9, "bold"))
        dxf_rsb = ttk.Scrollbar(result_lf, command=self._dxf_result_text.yview)
        self._dxf_result_text.configure(yscrollcommand=dxf_rsb.set)
        self._dxf_result_text.grid(row=0, column=0, sticky="nsew")
        dxf_rsb.grid(row=0, column=1, sticky="ns")

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
            "外轮廓：在多种封闭图元中选外层；外形内完全落入的封闭区域视为孔（避免把孔圆当成零件）。\n"
            "图元：LWPOLYLINE、闭合 POLYLINE、圆、椭圆、SPLINE、HATCH（外环+孔）；"
            "LINE/ARC 等散线若首尾相接会尝试拼成闭合（端点宜精确相接）。\n"
            "密铺自动优选 0°/180° 与横/纵镜像（不含 90°/270°），使单张件数最多。\n"
            "「+90°」在排样前将轮廓绕质心旋转（可连点累加）。零件间隙用等距半宽缓冲，净距不小于设定值。\n"
            "密铺含列向互嵌、标准网格、紧凑、交错行等；预览区滚轮缩放、左键拖动、重置比例；"
            "改数字后按 Enter 或离开输入框再计算。",
        )
        help_txt.configure(state="disabled")

        toolbar = ttk.Frame(right)
        toolbar.pack(fill="x", pady=(0, 8))
        ttk.Button(toolbar, text="重置显示比例", command=self._reset_view).pack(
            side="right"
        )

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
        ent.bind("<FocusOut>", lambda _e: self._recalculate(), add="+")
        ent.bind("<Return>", lambda _e: self._recalculate(), add="+")

    def _fiber_add_90(self) -> None:
        self.vars["fiber_deg"].set((int(self.vars["fiber_deg"].get()) + 90) % 360)
        self._recalculate()

    def _on_packing_combo(self, _event: Any = None) -> None:
        label = self._packing_combo.get()
        for key, lab in self._PACK_MODES:
            if lab == label:
                if self.vars["packing_mode"].get() != key:
                    self.vars["packing_mode"].set(key)
                self._recalculate()
                break

    def _cycle_packing_mode(self) -> None:
        cur = self.vars["packing_mode"].get()
        keys = [k for k, _ in self._PACK_MODES]
        try:
            i = keys.index(cur)
        except ValueError:
            i = 0
        nxt = keys[(i + 1) % len(keys)]
        self.vars["packing_mode"].set(nxt)
        for key, lab in self._PACK_MODES:
            if key == nxt:
                self._packing_combo.set(lab)
                break
        self._recalculate()

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

    def _browse_dxf(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 DXF 文件",
            filetypes=[("DXF 图纸", "*.dxf"), ("所有文件", "*.*")],
        )
        if path:
            self.vars["dxf_path"].set(path)
            self._recalculate()

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

    def _show_busy_overlay(self) -> tk.Toplevel:
        """屏幕居中显示不确定进度条；计算在后台线程时由主线程 periodic update 驱动滚动。"""
        root = self.winfo_toplevel()
        win = tk.Toplevel(root)
        win.title("请稍候")
        win.transient(root)
        win.resizable(False, False)
        win.protocol("WM_DELETE_WINDOW", lambda: None)
        frm = ttk.Frame(win, padding=(28, 22, 28, 24))
        frm.pack()
        ttk.Label(frm, text="计算中…", font=("Segoe UI", 11)).pack(pady=(0, 12))
        pb = ttk.Progressbar(frm, length=280, mode="indeterminate")
        pb.pack()
        pb.start(12)
        win._busy_pb = pb  # type: ignore[attr-defined]
        win._busy_root = root  # type: ignore[attr-defined]
        try:
            win.grab_set()
        except tk.TclError:
            pass
        win.update_idletasks()
        ww = max(win.winfo_reqwidth(), 200)
        wh = max(win.winfo_reqheight(), 80)
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        x = max(0, (sw - ww) // 2)
        y = max(0, (sh - wh) // 2)
        win.geometry(f"{ww}x{wh}+{x}+{y}")
        win.lift()
        root.update()
        return win

    @staticmethod
    def _destroy_busy_overlay(win: tk.Toplevel) -> None:
        pb = getattr(win, "_busy_pb", None)
        if pb is not None:
            try:
                pb.stop()
            except tk.TclError:
                pass
        try:
            win.grab_release()
        except tk.TclError:
            pass
        try:
            win.destroy()
        except tk.TclError:
            pass
        r = getattr(win, "_busy_root", None)
        if r is not None:
            try:
                r.update_idletasks()
            except tk.TclError:
                pass

    @staticmethod
    def _best_last_sheet_dxf(
        remain: int,
        cell_w: float,
        cell_h: float,
        part_wkt: str,
        gap_edge: int,
        stagger_x: float = 0.0,
        part_wkt_180: str = "",
        interlock_dy: float = 0.0,
    ) -> tuple[tuple[int, int], tuple[int, int, int]]:
        from shapely import wkt

        pd0 = wkt.loads(part_wkt)
        pd1 = wkt.loads(part_wkt_180) if (part_wkt_180 or "").strip() else None
        best = None
        for c in range(1, remain + 1):
            if pd1 is not None:
                tw, th = dx.union_tail_bounds_mm_interlock_col(
                    remain,
                    c,
                    cell_w,
                    cell_h,
                    pd0,
                    pd1,
                    float(gap_edge),
                    stagger_x,
                    float(interlock_dy),
                )
            else:
                tw, th = dx.union_tail_bounds_mm(
                    remain, c, cell_w, cell_h, pd0, float(gap_edge), stagger_x
                )
            final_w = max(int(math.ceil(tw)), MIN_SHEET_SIDE)
            final_h = max(int(math.ceil(th)), MIN_SHEET_SIDE)
            score = (final_w * final_h, final_w + final_h, final_w)
            item = (score, final_w, final_h, c)
            if best is None or item[0] < best[0]:
                best = item
        assert best is not None
        _, final_w, final_h, best_cols = best
        best_rows = int(math.ceil(remain / best_cols))
        capacity = best_cols * best_rows
        return (final_w, final_h), (best_cols, best_rows, capacity)

    @staticmethod
    def _compute_layout_job(
        path: str,
        gap_edge: int,
        gap_part: int,
        sheet_w: int,
        sheet_h: int,
        qty: int,
        fiber_deg: int,
        pack_mode: str,
    ) -> dict[str, Any]:
        """在后台线程中运行；只返回可序列化数据，不触碰 Tk。"""
        from shapely import wkt as shapely_wkt
        from shapely.affinity import rotate as shp_rotate

        fd = int(fiber_deg) % 360
        try:
            poly, unit_note = dx.load_largest_outline_polygon(path)
        except Exception as ex:
            return {
                "kind": "err",
                "result_var": f"读取 DXF 失败：{ex}",
                "detail_var": "",
            }

        if fd:
            poly = shp_rotate(poly, -float(fd), origin="centroid")

        inner_w = float(sheet_w - 2 * gap_edge)
        inner_h = float(sheet_h - 2 * gap_edge)
        if inner_w <= 0 or inner_h <= 0:
            return {
                "kind": "err",
                "result_var": "板边间隙过大，有效区域无效。",
                "detail_var": "",
            }

        try:
            (
                rot_deg,
                cell_w,
                cell_h,
                poly_draw,
                cols,
                rows,
                stagger_x,
                pack_hint,
                mirror_x,
                mirror_y,
                wkt180,
                interlock_dy,
            ) = dx.layout_dxf_packing(
                poly, float(gap_part), inner_w, inner_h, pack_mode
            )
        except Exception as ex:
            return {
                "kind": "err",
                "result_var": f"轮廓处理失败：{ex}",
                "detail_var": "",
            }

        if cols <= 0 or rows <= 0:
            return {
                "kind": "small",
                "result_var": "目标板材太小，无法放下 1 件（已考虑间隙与缓冲）。",
                "detail_var": unit_note,
            }

        part_wkt = shapely_wkt.dumps(poly_draw)
        part_wkt_180 = (wkt180 or "").strip()
        alternate_col_180 = bool(part_wkt_180)
        idy = float(interlock_dy) if alternate_col_180 else 0.0
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
                DxfNestingTab._best_last_sheet_dxf(
                    remain,
                    cell_w,
                    cell_h,
                    part_wkt,
                    gap_edge,
                    stagger_x,
                    part_wkt_180,
                    idy,
                )
            )
            last_used = remain

        mir_parts: list[str] = []
        if mirror_x:
            mir_parts.append("横向镜像")
        if mirror_y:
            mir_parts.append("纵向镜像")
        mir_txt = "、".join(mir_parts) if mir_parts else "无镜像"

        detail_var = (
            f"{unit_note}\n"
            f"纤维向：相对 DXF 已旋 {fd}°（排样前）\n"
            f"{pack_hint}\n"
            f"排样姿态：再旋 {rot_deg}°；{mir_txt}；步距约 {cell_w:.1f} x {cell_h:.1f} mm"
            + (f"；行错开 {stagger_x:.1f} mm" if stagger_x > 1e-3 else "")
            + "\n"
            f"目标板单张：{cols} 列 x {rows} 行 = {per_sheet} 件\n"
            f"尾板：{last_cols} 列 x {last_rows} 行，可放 {last_cap} 件，实际放 {last_used} 件"
        )

        layout = {
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
            "gap_edge": gap_edge,
            "cell_w": cell_w,
            "cell_h": cell_h,
            "stagger_x": stagger_x,
            "part_wkt": part_wkt,
            "part_wkt_180": part_wkt_180,
            "alternate_col_180": alternate_col_180,
            "interlock_dy": idy,
        }

        return {
            "kind": "ok",
            "result_var": (
                f"共需要板材：{total_sheets} 张\n"
                f"其中完整目标板：{full_count} 张\n"
                f"最后一张尺寸：{last_w} x {last_h} mm"
            ),
            "detail_var": detail_var,
            "lbl_full_title": f"完整目标板（{sheet_w} x {sheet_h}）",
            "lbl_full_sub": f"每张可放 {per_sheet} 件，完整板 {full_count} 张",
            "lbl_last_title": f"最后一张（{last_w} x {last_h}）",
            "lbl_last_sub": f"实际放 {last_used} 件，总板数 {total_sheets} 张",
            "layout": layout,
        }

    def _poll_dxf_recalc(
        self,
        thread: threading.Thread,
        busy_win: tk.Toplevel,
        out_q: queue.Queue[dict[str, Any]],
        gen: int,
    ) -> None:
        if gen != self._dxf_recalc_gen:
            try:
                self._destroy_busy_overlay(busy_win)
            except Exception:
                pass
            if self._active_busy_win is busy_win:
                self._active_busy_win = None
            return
        if thread.is_alive():
            try:
                busy_win.update()
            except tk.TclError:
                return
            self.after(
                50,
                lambda: self._poll_dxf_recalc(thread, busy_win, out_q, gen),
            )
            return
        try:
            data = out_q.get_nowait()
        except queue.Empty:
            data = {
                "kind": "err",
                "result_var": "计算未完成或异常中断。",
                "detail_var": "",
            }
        self._destroy_busy_overlay(busy_win)
        if self._active_busy_win is busy_win:
            self._active_busy_win = None
        if gen != self._dxf_recalc_gen:
            return
        self._apply_dxf_compute_result(data)

    def _sync_dxf_result_text(self) -> None:
        t = getattr(self, "_dxf_result_text", None)
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

    def _apply_dxf_compute_result(self, data: dict[str, Any]) -> None:
        try:
            kind = data.get("kind", "err")
            if kind == "err":
                self.result_var.set(str(data.get("result_var", "出错")))
                self.detail_var.set(str(data.get("detail_var", "")))
                return
            if kind == "small":
                self.result_var.set(str(data.get("result_var", "")))
                self.detail_var.set(str(data.get("detail_var", "")))
                return
            if kind != "ok":
                self.result_var.set("未知结果类型。")
                self.detail_var.set("")
                return
            self.result_var.set(str(data["result_var"]))
            self.detail_var.set(str(data["detail_var"]))
            self.lbl_full_title.config(text=str(data["lbl_full_title"]))
            self.lbl_full_sub.config(text=str(data["lbl_full_sub"]))
            self.lbl_last_title.config(text=str(data["lbl_last_title"]))
            self.lbl_last_sub.config(text=str(data["lbl_last_sub"]))
            self._layout = data["layout"]
            self._draw_previews()
        finally:
            self._sync_dxf_result_text()

    def _recalculate(self) -> None:
        if not dx.deps_available():
            return

        try:
            path = self.vars["dxf_path"].get().strip()
            gap_edge = _safe_int(self.vars["gap_edge"].get(), -1)
            gap_part = _safe_int(self.vars["gap_part"].get(), -1)
            sheet_w = _safe_int(self.vars["sheet_w"].get(), 0)
            sheet_h = _safe_int(self.vars["sheet_h"].get(), 0)
            qty = _safe_int(self.vars["qty"].get(), 0)

            self.canvas_full.delete("all")
            self.canvas_last.delete("all")
            self._layout = None
            self.lbl_full_title.config(text="")
            self.lbl_full_sub.config(text="")
            self.lbl_last_title.config(text="")
            self.lbl_last_sub.config(text="")

            if not path:
                self.result_var.set("请选择 DXF 文件。")
                self.detail_var.set("")
                return

            if min(sheet_w, sheet_h, qty) <= 0:
                self.result_var.set("板材尺寸、数量须大于 0。")
                self.detail_var.set("")
                return
            if gap_edge < 0 or gap_part < 0:
                self.result_var.set("板边间隙、零件间隙不能为负数。")
                self.detail_var.set("")
                return

            self._dxf_recalc_gen += 1
            gen = self._dxf_recalc_gen
            prev_busy = self._active_busy_win
            if prev_busy is not None:
                try:
                    self._destroy_busy_overlay(prev_busy)
                except Exception:
                    pass
                self._active_busy_win = None

            fiber_deg = int(self.vars["fiber_deg"].get()) % 360
            pack_mode = self.vars["packing_mode"].get().strip().lower()

            busy_win = self._show_busy_overlay()
            self._active_busy_win = busy_win
            out_q: queue.Queue[dict[str, Any]] = queue.Queue(maxsize=1)

            def worker() -> None:
                try:
                    r = DxfNestingTab._compute_layout_job(
                        path,
                        gap_edge,
                        gap_part,
                        sheet_w,
                        sheet_h,
                        qty,
                        fiber_deg,
                        pack_mode,
                    )
                    out_q.put(r)
                except Exception as e:
                    out_q.put(
                        {
                            "kind": "err",
                            "result_var": f"计算异常：{e}",
                            "detail_var": "",
                        }
                    )

            th = threading.Thread(target=worker, daemon=True)
            th.start()
            self.after(50, lambda: self._poll_dxf_recalc(th, busy_win, out_q, gen))
        finally:
            self._sync_dxf_result_text()

    def _draw_previews(self) -> None:
        if self._layout is None:
            return
        L = self._layout
        self._draw_dxf_panel(
            self.canvas_full,
            self.zoom_full,
            self.pan_full,
            L["sheet_w"],
            L["sheet_h"],
            L["cols"],
            L["rows"],
            L["used_full"],
            L["gap_edge"],
            L["cell_w"],
            L["cell_h"],
            L["stagger_x"],
            L["part_wkt"],
            L.get("part_wkt_180") or "",
            bool(L.get("alternate_col_180")),
            float(L.get("interlock_dy") or 0.0),
        )
        self._draw_dxf_panel(
            self.canvas_last,
            self.zoom_last,
            self.pan_last,
            L["last_w"],
            L["last_h"],
            L["last_cols"],
            L["last_rows"],
            L["used_last"],
            L["gap_edge"],
            L["cell_w"],
            L["cell_h"],
            L["stagger_x"],
            L["part_wkt"],
            L.get("part_wkt_180") or "",
            bool(L.get("alternate_col_180")),
            float(L.get("interlock_dy") or 0.0),
        )

    def _draw_dxf_panel(
        self,
        canvas: tk.Canvas,
        zoom: float,
        pan: list[float],
        sheet_w: int,
        sheet_h: int,
        cols: int,
        rows: int,
        used: int,
        gap_edge: int,
        cell_w: float,
        cell_h: float,
        stagger_x: float,
        part_wkt: str,
        part_wkt_180: str = "",
        alternate_col_180: bool = False,
        interlock_dy: float = 0.0,
    ) -> None:
        from shapely import wkt as shapely_wkt

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

        part_poly = shapely_wkt.loads(part_wkt)
        part_poly_180 = (
            shapely_wkt.loads(part_wkt_180)
            if alternate_col_180 and (part_wkt_180 or "").strip()
            else None
        )

        def ring_to_canvas_flat(bx: float, by: float, ring: list[tuple[float, float]]) -> list[float]:
            flat: list[float] = []
            for px, py in ring:
                flat.extend([bx + px * scale, by + py * scale])
            return flat

        ge = float(gap_edge)
        st = float(stagger_x)
        idx = 0
        for r in range(rows):
            for c in range(cols):
                if idx >= used:
                    break
                bx = ox + (ge + c * cell_w + (r % 2) * st) * scale
                y_mm = ge + r * cell_h + ((c % 2) * interlock_dy if alternate_col_180 else 0.0)
                by = oy + y_mm * scale
                pp = (
                    part_poly_180
                    if part_poly_180 is not None and (c % 2 == 1)
                    else part_poly
                )
                ext = list(pp.exterior.coords)
                fe = ring_to_canvas_flat(bx, by, ext)
                if len(fe) >= 6:
                    canvas.create_polygon(
                        *fe,
                        fill="#9fd3ff",
                        outline="#1a4d7a",
                        width=1,
                    )
                for intr in pp.interiors:
                    hi = ring_to_canvas_flat(bx, by, list(intr.coords))
                    if len(hi) >= 6:
                        canvas.create_polygon(
                            *hi,
                            fill="#fffef6",
                            outline="#1a4d7a",
                            width=1,
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
