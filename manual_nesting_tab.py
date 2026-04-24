"""手动排样：导入 DXF 外轮廓、参数区、画布交互放置与阵列复制。"""

from __future__ import annotations

import math
import os
import tkinter as tk
from dataclasses import dataclass
from tkinter import filedialog, messagebox, ttk
from typing import Any, Optional

import dxf_outline as dx
from shapely.affinity import rotate, translate
from shapely.geometry import Point, box
from shapely.geometry.base import BaseGeometry

_GAP_TOL = 1e-4
_PICK_TOL_MM = 3.0
# 导入对话框 / 顶栏缩略图预览画布（像素）
_THUMB_IMPORT_CANVAS = 120
_THUMB_BAR_CANVAS = 96
_DRAG_START_PX = 5
_LONG_PRESS_MS = 100


def _gap_ok(p0: BaseGeometry, p1: BaseGeometry, gap_mm: float) -> bool:
    return float(p0.distance(p1)) >= float(gap_mm) - _GAP_TOL


def _safe_float(s: str, default: float) -> float:
    try:
        return float(s)
    except (TypeError, ValueError):
        return default


@dataclass
class ImportedShape:
    path: str
    name: str
    poly_at_centroid: Any  # Polygon, centroid at (0,0)
    # 源 DXF 中与 poly_at_centroid 对应的轮廓质心（图面坐标），用于导出时还原圆弧/圆
    src_cx: float
    src_cy: float
    inventory: int
    placed_count: int = 0


@dataclass
class PlacedInstance:
    id: int
    src_idx: int
    rot: int  # 0,90,180,270 — 绕质心顺时针累计（与 Shapely rotate(..., -rot) 一致）
    cx: float
    cy: float


class ManualNestingTab(ttk.Frame):
    """手动在板内放置 DXF 外轮廓；支持旋转、拖动、框选与沿拖动的二维阵列复制。"""

    _PREVIEW_PAD_L = 16
    _PREVIEW_PAD_R = 16
    _PREVIEW_PAD_T = 12
    _PREVIEW_PAD_B = 36

    def __init__(self, parent: tk.Misc, **kwargs: Any) -> None:
        super().__init__(parent, **kwargs)
        self._shapes: list[ImportedShape] = []
        self._placed: list[PlacedInstance] = []
        self._next_id = 1
        self._selection: set[int] = set()

        self.var_gap_part = tk.StringVar(value="5")
        self.var_gap_edge = tk.StringVar(value="25")
        self.var_sheet_w = tk.StringVar(value="2000")
        self.var_sheet_h = tk.StringVar(value="1200")
        # 板料宽/高在输入完成后再参与画布缩放（避免改数字过程中不停重算）
        self._applied_sheet_w = 2000.0
        self._applied_sheet_h = 1200.0

        self._ox = 0.0
        self._oy = 0.0
        self._scale = 1.0
        self._sheet_w = 1200.0
        self._sheet_h = 2000.0

        # idle | palette_ghost | move_one | replica_drag | rubber
        self._mode = "idle"
        self._ghost_src: Optional[int] = None
        self._ghost_rot = 0
        self._ghost_cx = 0.0
        self._ghost_cy = 0.0
        self._move_anchor_id: Optional[int] = None
        self._move_ids: set[int] = set()
        self._move_dx: dict[int, float] = {}
        self._move_dy: dict[int, float] = {}
        self._move_initial_rots: dict[int, int] = {}

        self._rubber: Optional[tuple[float, float, float, float]] = None  # canvas x0,y0,x1,y1
        self._press_canvas: Optional[tuple[float, float]] = None
        self._replica_press_mm: Optional[tuple[float, float]] = None
        self._replica_ghost_offsets: list[tuple[float, float]] = []
        self._long_press_timer: Optional[str] = None
        self._replica_from_long_press = False
        self._replica_locked_sx: Optional[float] = None
        self._replica_locked_sy: Optional[float] = None

        self._thumb_images: list[tk.PhotoImage] = []

        self._build_ui()

        if dx.deps_available():
            for v in (self.var_gap_part, self.var_gap_edge):
                v.trace_add("write", lambda *_: self._schedule_redraw())
            self._canvas.bind("<Configure>", lambda *_: self._redraw())
            self._canvas.bind("<Motion>", self._on_motion)
            self._canvas.bind("<ButtonPress-1>", self._on_b1_press)
            self._canvas.bind("<B1-Motion>", self._on_b1_motion)
            self._canvas.bind("<ButtonRelease-1>", self._on_b1_release)
            self._canvas.bind("<ButtonPress-3>", self._on_b3_press)
            self._canvas.bind("<Double-Button-1>", self._on_double_b1)
            self.winfo_toplevel().bind(
                "<Escape>", self._on_escape_toplevel, add="+"
            )
            root = self.winfo_toplevel()
            root.bind_all("<Delete>", self._on_delete_all, add="+")
            root.bind_all("<KP_Delete>", self._on_delete_all, add="+")

    def _build_ui(self) -> None:
        if not dx.deps_available():
            ttk.Label(
                self,
                text="本页需要：pip install ezdxf shapely",
                padding=24,
            ).pack(anchor="nw")
            return

        outer = ttk.Frame(self, padding=10)
        outer.pack(fill="both", expand=True)
        self._outer = outer

        top_bar = ttk.Frame(outer)
        top_bar.pack(fill="x", pady=(0, 8))
        ttk.Button(top_bar, text="导入 DXF…", command=self._import_dxf).pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(top_bar, text="保存排样 DXF…", command=self._save_layout_dxf).pack(
            side="left", padx=(0, 8)
        )
        ttk.Label(
            top_bar,
            text="双击缩略图放置（虚影在有效区内）；框选后双击其一可整体移动；右键 90°；Esc 取消；Delete 删除选中；板料宽高失焦生效；长按拖动复制阵列。",
            foreground="#444",
        ).pack(side="left")

        self._thumb_row = ttk.Frame(outer)
        self._thumb_row.pack(fill="x", pady=(0, 6))

        main = ttk.Frame(outer)
        main.pack(fill="both", expand=True)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        left = ttk.LabelFrame(main, text="参数（mm）", padding=8)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self._left_panel = left
        self._add_row(left, "零件间距", self.var_gap_part)
        self._add_row(left, "零件与边缘间距", self.var_gap_edge)
        self._add_sheet_commit_row(left, "板料宽", self.var_sheet_w)
        self._add_sheet_commit_row(left, "板料高", self.var_sheet_h)

        right = ttk.Frame(main)
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)
        frame_cv = tk.Frame(right, bg="#505050", padx=2, pady=2)
        frame_cv.grid(row=0, column=0, sticky="nsew")
        self._canvas = tk.Canvas(
            frame_cv, bg="#f7f7f7", highlightthickness=0, takefocus=True
        )
        self._canvas.pack(fill="both", expand=True)

        self._redraw_job: Optional[str] = None

    def _add_row(self, parent: ttk.Widget, label: str, var: tk.StringVar) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=4)
        row.columnconfigure(1, weight=1)
        ttk.Label(row, text=f"{label}：", width=14, anchor="w").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Entry(row, textvariable=var, width=12).grid(row=0, column=1, sticky="ew")

    def _add_sheet_commit_row(self, parent: ttk.Widget, label: str, var: tk.StringVar) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=4)
        row.columnconfigure(1, weight=1)
        ttk.Label(row, text=f"{label}：", width=14, anchor="w").grid(
            row=0, column=0, sticky="w"
        )
        ent = ttk.Entry(row, textvariable=var, width=12)
        ent.grid(row=0, column=1, sticky="ew")

        def apply_sheet(_event: Any = None) -> None:
            if var is self.var_sheet_w:
                self._applied_sheet_w = max(1.0, _safe_float(var.get(), self._applied_sheet_w))
            else:
                self._applied_sheet_h = max(1.0, _safe_float(var.get(), self._applied_sheet_h))
            self._schedule_redraw()

        ent.bind("<FocusOut>", apply_sheet, add="+")
        ent.bind("<Return>", apply_sheet, add="+")

    def _on_delete_all(self, ev: tk.Event) -> Optional[str]:
        if not self._widget_is_under_tab(ev.widget):
            return None
        try:
            wc = ev.widget.winfo_class()
        except tk.TclError:
            return None
        if wc in ("TEntry", "Entry", "Text", "Spinbox", "TSpinbox"):
            return None
        return self._on_delete(ev)

    def _widget_is_under_tab(self, widget: Any) -> bool:
        w: Any = widget
        while w is not None:
            if w == self:
                return True
            try:
                w = w.master
            except (tk.TclError, AttributeError):
                break
        return False

    def _on_escape_toplevel(self, ev: tk.Event) -> Optional[str]:
        if not self._widget_is_under_tab(ev.widget):
            return None
        return self._on_escape(ev)

    def _on_escape(self, _ev: Optional[tk.Event] = None) -> Optional[str]:
        if self._mode in ("palette_ghost", "move_one"):
            self._exit_ghost_placement(reset_mode=True)
            return "break"
        if self._mode == "replica_drag":
            self._clear_rubber_and_replica_state()
            self._mode = "idle"
            self._press_canvas = None
            self._redraw()
            return "break"
        if self._mode == "rubber":
            self._rubber = None
            self._mode = "idle"
            self._press_canvas = None
            self._redraw()
            return "break"
        if self._mode == "idle" and self._selection:
            self._selection.clear()
            self._redraw()
            return "break"
        return None

    def _exit_ghost_placement(self, *, reset_mode: bool) -> None:
        self._ghost_src = None
        self._ghost_rot = 0
        self._move_anchor_id = None
        self._move_ids.clear()
        self._move_dx.clear()
        self._move_dy.clear()
        self._move_initial_rots.clear()
        if reset_mode:
            self._mode = "idle"
        self._clear_rubber_and_replica_state()

    def _on_delete(self, _ev: Optional[tk.Event] = None) -> Optional[str]:
        if self._mode == "replica_drag":
            self._clear_rubber_and_replica_state()
            self._mode = "idle"
            self._press_canvas = None
            self._redraw()
        if self._mode == "rubber":
            self._rubber = None
            self._mode = "idle"
            self._press_canvas = None
            self._redraw()
        if not self._selection:
            return None
        if self._mode in ("palette_ghost", "move_one"):
            self._exit_ghost_placement(reset_mode=True)
        remove_ids = {p.id for p in self._placed if p.id in self._selection}
        if not remove_ids:
            self._selection.clear()
            return "break"
        self._placed = [p for p in self._placed if p.id not in remove_ids]
        self._recount_placed_by_src()
        self._selection.clear()
        self._rebuild_thumbs()
        self._redraw()
        return "break"

    def _recount_placed_by_src(self) -> None:
        n = len(self._shapes)
        cnt = [0] * n
        for p in self._placed:
            if 0 <= p.src_idx < n:
                cnt[p.src_idx] += 1
        for i, sh in enumerate(self._shapes):
            sh.placed_count = cnt[i]

    def _schedule_redraw(self) -> None:
        if self._redraw_job is not None:
            try:
                self.after_cancel(self._redraw_job)
            except tk.TclError:
                pass
        self._redraw_job = self.after_idle(self._do_scheduled_redraw)

    def _do_scheduled_redraw(self) -> None:
        self._redraw_job = None
        self._redraw()

    def _nums(self) -> tuple[float, float, float, float]:
        gp = max(0.0, _safe_float(self.var_gap_part.get(), 5.0))
        ge = max(0.0, _safe_float(self.var_gap_edge.get(), 25.0))
        sw = max(1.0, float(self._applied_sheet_w))
        sh = max(1.0, float(self._applied_sheet_h))
        return gp, ge, sw, sh

    def _inner_poly(self) -> Any:
        gp, ge, sw, sh = self._nums()
        if sw <= 2 * ge or sh <= 2 * ge:
            return box(0, 0, max(sw, 1), max(sh, 1))
        return box(ge, ge, sw - ge, sh - ge)

    def _world_poly(self, ins: PlacedInstance) -> Any:
        p0 = self._shapes[ins.src_idx].poly_at_centroid
        pr = rotate(p0, -float(ins.rot % 360), origin=(0.0, 0.0))
        return translate(pr, xoff=ins.cx, yoff=ins.cy)

    def _placement_matrix_44(self, ins: PlacedInstance) -> Any:
        """与 _world_poly 一致的平面刚体变换（绕轮廓质心旋转 + 平移到放置中心）。"""
        from ezdxf.math import Matrix44

        sh = self._shapes[ins.src_idx]
        t = math.radians(float(ins.rot % 360))
        c, s = math.cos(t), math.sin(t)
        scx, scy = float(sh.src_cx), float(sh.src_cy)
        tx = float(ins.cx) - c * scx - s * scy
        ty = float(ins.cy) + s * scx - c * scy
        return Matrix44.from_2d_transformation([c, -s, s, c, tx, ty])

    def _export_add_outline_polylines(self, msp: Any, ins: PlacedInstance) -> None:
        wp = self._world_poly(ins)
        if not getattr(wp, "exterior", None):
            return
        ext = list(wp.exterior.coords)
        if len(ext) < 3:
            return
        if ext[0] == ext[-1]:
            ext = ext[:-1]
        pts = [(float(x), float(y)) for x, y in ext]
        msp.add_lwpolyline(pts, dxfattribs={"layer": "PARTS"}, close=True)
        for intr in getattr(wp, "interiors", []) or []:
            ir = list(intr.coords)
            if len(ir) < 3:
                continue
            if ir[0] == ir[-1]:
                ir = ir[:-1]
            ipt = [(float(x), float(y)) for x, y in ir]
            msp.add_lwpolyline(ipt, dxfattribs={"layer": "PARTS"}, close=True)

    def _export_ensure_layer(
        self, doc: Any, src_doc: Any, layer_name: str
    ) -> None:
        if layer_name in doc.layers:
            return
        if layer_name in src_doc.layers:
            sl = src_doc.layers.get(layer_name)
            doc.layers.add(
                layer_name,
                color=int(sl.color),
                linetype=sl.dxf.linetype,
            )
        else:
            doc.layers.add(layer_name, color=5)

    def _export_try_add_native_modelspace(
        self, doc: Any, msp: Any, ins: PlacedInstance, src_doc: Any
    ) -> bool:
        """将源文件模型空间中的几何实体变换后写入目标图；失败返回 False。"""
        import ezdxf

        m44 = self._placement_matrix_44(ins)
        shape = self._shapes[ins.src_idx]
        added = 0
        try:
            for entity in src_doc.modelspace():
                if not entity.is_alive or getattr(entity, "is_virtual", False):
                    continue
                if not isinstance(entity, ezdxf.entities.DXFGraphic):
                    continue
                try:
                    self._export_ensure_layer(doc, src_doc, entity.dxf.layer)
                    ne = entity.copy()
                    ne.transform(m44)
                    msp.add_entity(ne)
                    added += 1
                except (ezdxf.DXFError, ezdxf.DXFStructureError, ValueError, TypeError):
                    continue
                except Exception:
                    continue
        except Exception:
            return False
        return added > 0

    def _ghost_poly_at(self, src_idx: int, rot_deg: int, cx: float, cy: float) -> Any:
        p0 = self._shapes[src_idx].poly_at_centroid
        pr = rotate(p0, -float(rot_deg % 360), origin=(0.0, 0.0))
        return translate(pr, xoff=cx, yoff=cy)

    def _ghost_world_poly(self) -> Optional[Any]:
        if self._ghost_src is None:
            return None
        return self._ghost_poly_at(
            self._ghost_src, int(self._ghost_rot), self._ghost_cx, self._ghost_cy
        )

    def _palette_ghost_ok_at(self, cx: float, cy: float) -> bool:
        if self._ghost_src is None:
            return False
        gh = self._ghost_poly_at(
            int(self._ghost_src), int(self._ghost_rot), cx, cy
        )
        gp, _ge, _sw, _sh = self._nums()
        inner = self._inner_poly()
        return self._placement_ok(gh, inner, gp, self._others_list(), None)

    def _constrain_palette_ghost_fallback_inner(
        self, mmx: float, mmy: float
    ) -> tuple[float, float]:
        """上一位置不可用或与目标重合时：沿目标→板内中心细步查找可行点（旋转后卡死等）。"""
        inner = self._inner_poly()
        icx = (inner.bounds[0] + inner.bounds[2]) * 0.5
        icy = (inner.bounds[1] + inner.bounds[3]) * 0.5
        for step in range(1, 81):
            t = step / 80.0
            tx = mmx * (1.0 - t) + icx * t
            ty = mmy * (1.0 - t) + icy * t
            if self._palette_ghost_ok_at(tx, ty):
                return tx, ty
        if self._palette_ghost_ok_at(icx, icy):
            return icx, icy
        return self._ghost_cx, self._ghost_cy

    def _constrain_palette_ghost(self, mmx: float, mmy: float) -> tuple[float, float]:
        if self._ghost_src is None:
            return mmx, mmy
        if self._palette_ghost_ok_at(mmx, mmy):
            return mmx, mmy
        px, py = float(self._ghost_cx), float(self._ghost_cy)
        dx, dy = mmx - px, mmy - py
        # 沿「上一帧虚影质心 → 当前鼠标」线段取尽量靠近鼠标仍合法的位置，贴边时不再整体被拉向板心
        if (dx * dx + dy * dy) > 1e-12 and self._palette_ghost_ok_at(px, py):
            lo, hi = 0.0, 1.0
            for _ in range(28):
                mid = (lo + hi) * 0.5
                tx = px + dx * mid
                ty = py + dy * mid
                if self._palette_ghost_ok_at(tx, ty):
                    lo = mid
                else:
                    hi = mid
            return px + dx * lo, py + dy * lo
        return self._constrain_palette_ghost_fallback_inner(mmx, mmy)

    def _moved_part_poly_at(self, pid: int, anchor_cx: float, anchor_cy: float) -> Any:
        ins = next(p for p in self._placed if p.id == pid)
        eff = (self._move_initial_rots[pid] + int(self._ghost_rot)) % 360
        cx = anchor_cx + float(self._move_dx[pid])
        cy = anchor_cy + float(self._move_dy[pid])
        return self._ghost_poly_at(ins.src_idx, eff, cx, cy)

    def _move_group_fully_valid(self, anchor_cx: float, anchor_cy: float) -> bool:
        inner = self._inner_poly()
        gp, _ge, _sw, _sh = self._nums()
        polys: list[tuple[int, Any]] = []
        for pid in self._move_ids:
            polys.append(
                (pid, self._moved_part_poly_at(pid, anchor_cx, anchor_cy))
            )
        for pid, poly in polys:
            if not self._poly_fits_inner(poly, inner):
                return False
            for oid, op in self._others_list(self._move_ids):
                if not _gap_ok(poly, op, gp):
                    return False
            for pid2, poly2 in polys:
                if pid2 == pid:
                    continue
                if not _gap_ok(poly, poly2, gp):
                    return False
        return True

    def _poly_fits_inner(self, poly: Any, inner: Any) -> bool:
        try:
            if not inner.is_valid:
                inner = inner.buffer(0)
            if not poly.is_valid:
                poly = poly.buffer(0)
            return bool(inner.covers(poly))
        except Exception:
            return False

    def _placement_ok(
        self,
        poly: Any,
        inner: Any,
        gap_part: float,
        others: list[tuple[int, Any]],
        skip_ids: Optional[set[int]] = None,
    ) -> bool:
        if not self._poly_fits_inner(poly, inner):
            return False
        skip_ids = skip_ids or set()
        for pid, op in others:
            if pid in skip_ids:
                continue
            if not _gap_ok(poly, op, gap_part):
                return False
        return True

    def _others_list(self, skip: Optional[set[int]] = None) -> list[tuple[int, Any]]:
        skip = skip or set()
        out: list[tuple[int, Any]] = []
        for ins in self._placed:
            if ins.id in skip:
                continue
            out.append((ins.id, self._world_poly(ins)))
        return out

    def _hit_placed(self, mmx: float, mmy: float) -> Optional[int]:
        pt = Point(mmx, mmy)
        for ins in reversed(self._placed):
            wp = self._world_poly(ins)
            try:
                if wp.distance(pt) <= _PICK_TOL_MM or wp.contains(pt):
                    return ins.id
            except Exception:
                continue
        return None

    def _canvas_to_mm(self, cx: float, cy: float) -> tuple[float, float]:
        return (cx - self._ox) / self._scale, (cy - self._oy) / self._scale

    def _import_dxf(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择 DXF（可多选）",
            filetypes=[("DXF", "*.dxf"), ("所有文件", "*.*")],
        )
        if not paths:
            return
        dlg = tk.Toplevel(self.winfo_toplevel())
        dlg.title("定义库存数量（提示用）")
        dlg.transient(self.winfo_toplevel())
        entries: list[tuple[str, ttk.Entry]] = []
        frm = ttk.Frame(dlg, padding=10)
        frm.pack(fill="both", expand=True)
        for p in paths:
            row = ttk.Frame(frm)
            row.pack(fill="x", pady=4)
            row.columnconfigure(0, weight=1)
            ttk.Label(row, text=os.path.basename(p)[:44], anchor="w").grid(
                row=0, column=0, sticky="w", padx=(0, 6)
            )
            e = ttk.Entry(row, width=8)
            e.insert(0, "1")
            e.grid(row=0, column=1, sticky="e", padx=(0, 8))
            pv = tk.Canvas(
                row,
                width=_THUMB_IMPORT_CANVAS,
                height=_THUMB_IMPORT_CANVAS,
                bg="#fafafa",
                highlightthickness=1,
                highlightbackground="#c0c0c0",
            )
            pv.grid(row=0, column=2, sticky="e")
            try:
                poly, _n = dx.load_largest_outline_polygon(p)
                self._draw_thumb(pv, poly, size=_THUMB_IMPORT_CANVAS)
            except Exception:
                s2 = _THUMB_IMPORT_CANVAS // 2
                pv.create_text(
                    s2, s2, text="!", fill="#c00", font=("Segoe UI", 16, "bold")
                )
            entries.append((p, e))

        def ok() -> None:
            for p, e in entries:
                try:
                    q = max(0, int(float(e.get().strip() or "0")))
                except (TypeError, ValueError):
                    q = 0
                self._load_one_dxf(p, q)
            dlg.destroy()
            self._rebuild_thumbs()
            self._redraw()

        def cancel() -> None:
            dlg.destroy()

        bt = ttk.Frame(frm)
        bt.pack(fill="x", pady=(12, 0))
        ttk.Button(bt, text="确定", command=ok).pack(side="left")
        ttk.Button(bt, text="取消", command=cancel).pack(side="right")
        dlg.grab_set()

    def _save_layout_dxf(self) -> None:
        if not dx.deps_available():
            return
        if not self._placed:
            messagebox.showinfo(
                "无内容可保存",
                "当前没有已放置的零件，请先排样后再导出。",
                parent=self.winfo_toplevel(),
            )
            return
        path = filedialog.asksaveasfilename(
            title="保存排样为 DXF",
            defaultextension=".dxf",
            filetypes=[("DXF 图纸", "*.dxf"), ("所有文件", "*.*")],
            initialfile="手动排样.dxf",
        )
        if not path:
            return
        try:
            import ezdxf
            from ezdxf import units as ezdxf_units
        except ImportError:
            messagebox.showerror(
                "无法导出",
                "需要安装 ezdxf：pip install ezdxf",
                parent=self.winfo_toplevel(),
            )
            return
        try:
            doc = ezdxf.new("R2010", setup=True)
            try:
                doc.units = ezdxf_units.MM
            except (AttributeError, TypeError, ValueError):
                doc.header["$INSUNITS"] = 4
            doc.layers.add("SHEET", color=8)
            doc.layers.add("PARTS", color=5)
            msp = doc.modelspace()
            _gp, _ge, sw, sh = self._nums()
            sheet_ring = [
                (0.0, 0.0),
                (float(sw), 0.0),
                (float(sw), float(sh)),
                (0.0, float(sh)),
            ]
            msp.add_lwpolyline(
                sheet_ring,
                dxfattribs={"layer": "SHEET"},
                close=True,
            )
            src_open: dict[str, Any] = {}
            outline_fallback = 0
            for ins in self._placed:
                shp = self._shapes[ins.src_idx]
                fp = shp.path
                ok_native = False
                if fp and os.path.isfile(fp):
                    try:
                        if fp not in src_open:
                            src_open[fp] = ezdxf.readfile(fp)
                        ok_native = self._export_try_add_native_modelspace(
                            doc, msp, ins, src_open[fp]
                        )
                    except Exception:
                        ok_native = False
                if not ok_native:
                    self._export_add_outline_polylines(msp, ins)
                    outline_fallback += 1
            doc.saveas(path)
        except Exception as ex:
            messagebox.showerror(
                "保存失败",
                str(ex),
                parent=self.winfo_toplevel(),
            )
            return
        msg = (
            f"已导出 {len(self._placed)} 个零件及板料外框（mm）。\n"
            "几何尽量保留源 DXF 中的直线/圆弧/圆等实体（与排样时轮廓变换一致）。\n"
        )
        if outline_fallback:
            msg += (
                f"其中 {outline_fallback} 件无法写入源实体，已改为轮廓折线近似（PARTS 层）。\n"
            )
        msg += path
        messagebox.showinfo("已保存", msg, parent=self.winfo_toplevel())

    def _load_one_dxf(self, path: str, inventory: int) -> None:
        try:
            poly, _note = dx.load_largest_outline_polygon(path)
        except Exception as ex:
            messagebox.showerror(
                "DXF 无法使用",
                f"{os.path.basename(path)}\n{ex}",
                parent=self.winfo_toplevel(),
            )
            return
        c = poly.centroid
        sx, sy = float(c.x), float(c.y)
        poly_c = translate(poly, xoff=-sx, yoff=-sy)
        name = os.path.basename(path)
        self._shapes.append(
            ImportedShape(
                path=path,
                name=name,
                poly_at_centroid=poly_c,
                src_cx=sx,
                src_cy=sy,
                inventory=inventory,
                placed_count=0,
            )
        )

    def _rebuild_thumbs(self) -> None:
        for w in self._thumb_row.winfo_children():
            w.destroy()
        self._thumb_images.clear()
        for i, sh in enumerate(self._shapes):
            wrap = ttk.Frame(self._thumb_row, relief="groove", borderwidth=1, padding=4)
            wrap.pack(side="left", padx=4, pady=2)
            cv = tk.Canvas(
                wrap,
                width=_THUMB_BAR_CANVAS,
                height=_THUMB_BAR_CANVAS,
                bg="#fafafa",
                highlightthickness=0,
            )
            cv.pack()
            self._draw_thumb(cv, sh.poly_at_centroid, size=_THUMB_BAR_CANVAS)
            ttk.Label(wrap, text=f"{sh.name[:18]}", font=("Segoe UI", 8)).pack()
            ttk.Label(
                wrap, text=f"库存 {sh.inventory} / 已放 {sh.placed_count}", font=("Segoe UI", 7)
            ).pack()
            cv.bind("<Double-Button-1>", lambda _e, idx=i: self._start_palette_ghost(idx))
            wrap.bind("<Double-Button-1>", lambda _e, idx=i: self._start_palette_ghost(idx))

    @staticmethod
    def _draw_thumb(cv: tk.Canvas, poly_at_c: Any, size: int = 72) -> None:
        """实心填充外轮廓 + 孔洞（与背景同色）+ 描边，缩放以完整落入画布。"""
        cv.delete("all")
        b = poly_at_c.bounds
        w = max(b[2] - b[0], 1e-6)
        h = max(b[3] - b[1], 1e-6)
        margin = max(5, min(size // 7, 16))
        sc = min((size - 2 * margin) / w, (size - 2 * margin) / h)
        cx0, cy0 = size * 0.5, size * 0.5
        bcx = (b[0] + b[2]) * 0.5
        bcy = (b[1] + b[3]) * 0.5
        hole_fill = cv.cget("bg")
        if not hole_fill or str(hole_fill).lower() in ("", "systembuttonface"):
            hole_fill = "#fafafa"
        ow = 2 if size >= 88 else 1

        def ring_flat(ring: list[tuple[float, float]]) -> list[float]:
            flat: list[float] = []
            for px, py in ring:
                flat.append(cx0 + (px - bcx) * sc)
                flat.append(cy0 + (py - bcy) * sc)
            return flat

        ext = list(poly_at_c.exterior.coords)
        fe = ring_flat(ext)
        if len(fe) >= 6:
            cv.create_polygon(
                *fe,
                fill="#5a9fd4",
                outline="#0d2847",
                width=ow,
                activefill="#6eb0e0",
            )
        try:
            interiors = list(poly_at_c.interiors)
        except Exception:
            interiors = []
        for intr in interiors:
            hi = ring_flat(list(intr.coords))
            if len(hi) >= 6:
                cv.create_polygon(
                    *hi,
                    fill=hole_fill,
                    outline="#0d2847",
                    width=max(1, ow - 1),
                )

    def _start_palette_ghost(self, src_idx: int) -> None:
        self._clear_rubber_and_replica_state()
        self._mode = "palette_ghost"
        self._ghost_src = src_idx
        self._ghost_rot = 0
        self._selection.clear()
        _gp, ge, sw, sh = self._nums()
        cx, cy = sw / 2.0, sh / 2.0
        self._ghost_cx, self._ghost_cy = self._constrain_palette_ghost(cx, cy)
        self._redraw()

    def _clear_rubber_and_replica_state(self) -> None:
        self._rubber = None
        self._press_canvas = None
        self._replica_press_mm = None
        self._replica_ghost_offsets = []
        self._replica_locked_sx = None
        self._replica_locked_sy = None
        self._cancel_long_press_timer()
        self._replica_from_long_press = False

    def _cancel_long_press_timer(self) -> None:
        if self._long_press_timer is not None:
            try:
                self.after_cancel(self._long_press_timer)
            except tk.TclError:
                pass
            self._long_press_timer = None

    def _on_motion(self, ev: tk.Event) -> None:
        mmx, mmy = self._canvas_to_mm(float(ev.x), float(ev.y))
        if self._mode == "palette_ghost" and self._ghost_src is not None:
            cx, cy = self._constrain_palette_ghost(mmx, mmy)
            self._ghost_cx = cx
            self._ghost_cy = cy
            self._redraw()
        elif self._mode == "move_one" and self._move_anchor_id is not None:
            if self._move_group_fully_valid(mmx, mmy):
                self._ghost_cx = mmx
                self._ghost_cy = mmy
            self._redraw()

    def _on_b3_press(self, ev: tk.Event) -> str:
        if self._mode in ("palette_ghost", "move_one") and (
            self._ghost_src is not None or self._move_anchor_id is not None
        ):
            self._ghost_rot = (self._ghost_rot + 90) % 360
            if self._mode == "palette_ghost" and self._ghost_src is not None:
                self._ghost_cx, self._ghost_cy = self._constrain_palette_ghost(
                    self._ghost_cx, self._ghost_cy
                )
            elif self._mode == "move_one" and self._move_anchor_id is not None:
                if not self._move_group_fully_valid(self._ghost_cx, self._ghost_cy):
                    self._ghost_rot = (self._ghost_rot - 90) % 360
            self._redraw()
        return "break"

    def _on_double_b1(self, ev: tk.Event) -> None:
        mmx, mmy = self._canvas_to_mm(float(ev.x), float(ev.y))
        pid = self._hit_placed(mmx, mmy)
        if pid is None:
            return
        self._clear_rubber_and_replica_state()
        ins = next(p for p in self._placed if p.id == pid)
        if pid in self._selection and len(self._selection) > 1:
            ids = set(self._selection)
        else:
            ids = {pid}
        hins = ins
        self._mode = "move_one"
        self._move_anchor_id = pid
        self._move_ids = set(ids)
        self._move_dx.clear()
        self._move_dy.clear()
        self._move_initial_rots.clear()
        for qid in ids:
            q = next(p for p in self._placed if p.id == qid)
            self._move_dx[qid] = float(q.cx - hins.cx)
            self._move_dy[qid] = float(q.cy - hins.cy)
            self._move_initial_rots[qid] = int(q.rot % 360)
        self._ghost_src = None
        self._ghost_rot = 0
        self._ghost_cx = float(hins.cx)
        self._ghost_cy = float(hins.cy)
        self._selection = set(ids)
        try:
            self._canvas.focus_set()
        except tk.TclError:
            pass
        self._redraw()

    def _on_b1_press(self, ev: tk.Event) -> None:
        if self._mode == "replica_drag":
            return
        mmx, mmy = self._canvas_to_mm(float(ev.x), float(ev.y))
        self._press_canvas = (float(ev.x), float(ev.y))

        if self._mode == "palette_ghost" and self._ghost_src is not None:
            self._try_place_ghost()
            return
        if self._mode == "move_one" and self._move_anchor_id is not None:
            self._commit_move_one()
            return

        hit = self._hit_placed(mmx, mmy)
        if hit is not None and hit in self._selection:
            self._replica_press_mm = (mmx, mmy)
            self._schedule_long_press()
        else:
            self._replica_press_mm = None

        if hit is not None:
            if not (hit in self._selection and self._replica_press_mm):
                self._selection = {hit}
            try:
                self._canvas.focus_set()
            except tk.TclError:
                pass
            self._redraw()
            return

        self._selection.clear()
        self._rubber = (float(ev.x), float(ev.y), float(ev.x), float(ev.y))
        self._mode = "rubber"
        try:
            self._canvas.focus_set()
        except tk.TclError:
            pass
        self._redraw()

    def _schedule_long_press(self) -> None:
        self._cancel_long_press_timer()
        self._replica_from_long_press = False

        def fire() -> None:
            self._long_press_timer = None
            if self._mode not in ("idle", "rubber"):
                return
            if self._replica_press_mm is None:
                return
            self._replica_from_long_press = True

        self._long_press_timer = self.after(_LONG_PRESS_MS, fire)

    def _replica_batch_polys_at(self, ox: float, oy: float) -> list[Any]:
        batch: list[Any] = []
        for ins in self._placed:
            if ins.id not in self._selection:
                continue
            p0 = self._shapes[ins.src_idx].poly_at_centroid
            pr = rotate(p0, -float(ins.rot % 360), origin=(0.0, 0.0))
            batch.append(translate(pr, xoff=ins.cx + ox, yoff=ins.cy + oy))
        return batch

    def _replica_batch_valid(
        self,
        batch_polys: list[Any],
        inner: Any,
        gp: float,
        dyn_others: list[tuple[int, Any]],
    ) -> bool:
        for a in batch_polys:
            if not self._poly_fits_inner(a, inner):
                return False
            for _pid, op in dyn_others:
                if not _gap_ok(a, op, gp):
                    return False
            for b in batch_polys:
                if a is b:
                    continue
                if not _gap_ok(a, b, gp):
                    return False
        return True

    def _replica_try_add_offset(
        self,
        ox: float,
        oy: float,
        valid: list[tuple[float, float]],
        dyn_others: list[tuple[int, Any]],
        inner: Any,
        gp: float,
    ) -> bool:
        batch = self._replica_batch_polys_at(ox, oy)
        if not batch:
            return False
        if not self._replica_batch_valid(batch, inner, gp, dyn_others):
            return False
        valid.append((ox, oy))
        for wp in batch:
            dyn_others.append((-1, wp))
        return True

    def _on_b1_motion(self, ev: tk.Event) -> None:
        if self._press_canvas is None:
            return
        dx = float(ev.x) - self._press_canvas[0]
        dy = float(ev.y) - self._press_canvas[1]
        dist = math.hypot(dx, dy)

        if self._long_press_timer is not None and dist > 10:
            self._cancel_long_press_timer()

        if self._mode == "replica_drag":
            mmx, mmy = self._canvas_to_mm(float(ev.x), float(ev.y))
            self._update_replica_ghosts(mmx, mmy)
            self._redraw()
            return

        if self._mode == "rubber" and self._rubber is not None:
            x0, y0, _, _ = self._rubber
            self._rubber = (x0, y0, float(ev.x), float(ev.y))
            self._redraw()
            return

        if (
            self._replica_from_long_press
            and self._replica_press_mm is not None
            and dist >= _DRAG_START_PX
            and self._selection
        ):
            self._cancel_long_press_timer()
            self._mode = "replica_drag"
            self._replica_locked_sx = None
            self._replica_locked_sy = None
            mmx, mmy = self._canvas_to_mm(float(ev.x), float(ev.y))
            self._update_replica_ghosts(mmx, mmy)
            self._redraw()

    def _update_replica_ghosts(self, cur_mmx: float, cur_mmy: float) -> None:
        if not self._selection or self._replica_press_mm is None:
            self._replica_ghost_offsets = []
            return
        gp, _ge, _sw, _sh = self._nums()
        inner = self._inner_poly()

        polys: list[Any] = []
        for ins in self._placed:
            if ins.id not in self._selection:
                continue
            polys.append(self._world_poly(ins))

        if not polys:
            self._replica_ghost_offsets = []
            return

        u = polys[0]
        for q in polys[1:]:
            u = u.union(q)

        if self._replica_locked_sx is None or self._replica_locked_sy is None:
            minx, miny, maxx, maxy = u.bounds
            bw = max(maxx - minx, 1e-6)
            bh = max(maxy - miny, 1e-6)
            sx_bb = bw + gp
            sy_bb = bh + gp
            try:
                px = float(dx._min_positive_period_along_axis(u, gp, "x"))
                py = float(dx._min_positive_period_along_axis(u, gp, "y"))
            except Exception:
                px, py = sx_bb, sy_bb
            # 步长不得小于外包络+间隙，避免 min(周期,包络) 被异常小的周期拉成≈gp 导致与原件重叠、整排校验失败
            sx_try = min(max(px, gp + 1e-6), sx_bb)
            sy_try = min(max(py, gp + 1e-6), sy_bb)
            self._replica_locked_sx = float(max(sx_bb, sx_try))
            self._replica_locked_sy = float(max(sy_bb, sy_try))

        sx = float(self._replica_locked_sx)
        sy = float(self._replica_locked_sy)

        ddx = cur_mmx - self._replica_press_mm[0]
        ddy = cur_mmy - self._replica_press_mm[1]
        nx = int(math.floor(abs(ddx) / sx + 1e-9))
        ny = int(math.floor(abs(ddy) / sy + 1e-9))
        sgx = 1.0 if ddx >= 0 else -1.0
        sgy = 1.0 if ddy >= 0 else -1.0

        valid: list[tuple[float, float]] = []
        dyn_others: list[tuple[int, Any]] = list(self._others_list(skip=None))
        adx, ady = abs(ddx), abs(ddy)

        if adx >= ady:
            max_kx = 0
            for k in range(1, nx + 1):
                ox = k * sx * sgx
                if self._replica_try_add_offset(ox, 0.0, valid, dyn_others, inner, gp):
                    max_kx = k
                else:
                    break
            if max_kx == 0:
                for k in range(1, ny + 1):
                    oy = k * sy * sgy
                    if not self._replica_try_add_offset(0.0, oy, valid, dyn_others, inner, gp):
                        break
            else:
                for iy in range(1, ny + 1):
                    oy = iy * sy * sgy
                    if not self._replica_try_add_offset(0.0, oy, valid, dyn_others, inner, gp):
                        break
                    for k in range(1, max_kx + 1):
                        ox = k * sx * sgx
                        if not self._replica_try_add_offset(ox, oy, valid, dyn_others, inner, gp):
                            break
                    else:
                        continue
                    break
        else:
            max_ky = 0
            for k in range(1, ny + 1):
                oy = k * sy * sgy
                if self._replica_try_add_offset(0.0, oy, valid, dyn_others, inner, gp):
                    max_ky = k
                else:
                    break
            if max_ky == 0:
                for k in range(1, nx + 1):
                    ox = k * sx * sgx
                    if not self._replica_try_add_offset(ox, 0.0, valid, dyn_others, inner, gp):
                        break
            else:
                for ix in range(1, nx + 1):
                    ox = ix * sx * sgx
                    if not self._replica_try_add_offset(ox, 0.0, valid, dyn_others, inner, gp):
                        break
                    for k in range(1, max_ky + 1):
                        oy = k * sy * sgy
                        if not self._replica_try_add_offset(ox, oy, valid, dyn_others, inner, gp):
                            break
                    else:
                        continue
                    break

        self._replica_ghost_offsets = valid

    def _on_b1_release(self, ev: tk.Event) -> None:
        self._cancel_long_press_timer()

        if self._mode == "rubber" and self._rubber is not None:
            x0, y0, x1, y1 = self._rubber
            self._rubber = None
            self._mode = "idle"
            self._press_canvas = None
            xr0, yr0 = self._canvas_to_mm(min(x0, x1), min(y0, y1))
            xr1, yr1 = self._canvas_to_mm(max(x0, x1), max(y0, y1))
            if abs(x1 - x0) < _DRAG_START_PX and abs(y1 - y0) < _DRAG_START_PX:
                self._redraw()
                return
            rb = box(xr0, yr0, xr1, yr1)
            new_sel: set[int] = set()
            for ins in self._placed:
                try:
                    if rb.intersects(self._world_poly(ins).envelope):
                        new_sel.add(ins.id)
                except Exception:
                    continue
            self._selection = new_sel
            try:
                self._canvas.focus_set()
            except tk.TclError:
                pass
            self._redraw()
            return

        if self._mode == "replica_drag":
            self._commit_replicas()
            self._mode = "idle"
            self._replica_press_mm = None
            self._replica_ghost_offsets = []
            self._replica_locked_sx = None
            self._replica_locked_sy = None
            self._press_canvas = None
            self._replica_from_long_press = False
            self._redraw()
            return

        self._press_canvas = None
        self._replica_press_mm = None
        self._replica_from_long_press = False

    def _try_place_ghost(self) -> None:
        gh = self._ghost_world_poly()
        if gh is None or self._ghost_src is None:
            return
        gp, _ge, _sw, _sh = self._nums()
        inner = self._inner_poly()
        if not self._placement_ok(gh, inner, gp, self._others_list(), None):
            return
        ins = PlacedInstance(
            id=self._next_id,
            src_idx=self._ghost_src,
            rot=self._ghost_rot % 360,
            cx=self._ghost_cx,
            cy=self._ghost_cy,
        )
        self._next_id += 1
        self._placed.append(ins)
        sh = self._shapes[self._ghost_src]
        sh.placed_count += 1
        if sh.placed_count > sh.inventory:
            messagebox.showwarning(
                "超过库存提示",
                f"「{sh.name}」库存为 {sh.inventory}，当前已放置 {sh.placed_count} 件。\n仍允许继续放置。",
                parent=self.winfo_toplevel(),
            )
        self._mode = "idle"
        self._ghost_src = None
        self._ghost_rot = 0
        self._rebuild_thumbs()
        self._redraw()

    def _commit_move_one(self) -> None:
        if self._move_anchor_id is None or not self._move_ids:
            return
        if not self._move_group_fully_valid(self._ghost_cx, self._ghost_cy):
            messagebox.showinfo(
                "无法放置",
                "该位置与板边、其它零件间距或重叠要求不符。",
                parent=self.winfo_toplevel(),
            )
            return
        acx, acy = self._ghost_cx, self._ghost_cy
        new_placed: list[PlacedInstance] = []
        for p in self._placed:
            if p.id not in self._move_ids:
                new_placed.append(p)
                continue
            new_placed.append(
                PlacedInstance(
                    id=p.id,
                    src_idx=p.src_idx,
                    rot=(self._move_initial_rots[p.id] + int(self._ghost_rot)) % 360,
                    cx=acx + self._move_dx[p.id],
                    cy=acy + self._move_dy[p.id],
                )
            )
        self._placed = new_placed
        self._mode = "idle"
        self._move_anchor_id = None
        self._move_ids.clear()
        self._move_dx.clear()
        self._move_dy.clear()
        self._move_initial_rots.clear()
        self._ghost_src = None
        self._ghost_rot = 0
        self._selection.clear()
        self._redraw()

    def _commit_replicas(self) -> None:
        if not self._replica_ghost_offsets:
            return
        gp, _ge, _sw, _sh = self._nums()
        inner = self._inner_poly()
        new_instances: list[PlacedInstance] = []
        dyn_others: list[tuple[int, Any]] = list(self._others_list(skip=None))
        for ox, oy in self._replica_ghost_offsets:
            batch_polys: list[Any] = []
            batch_rows: list[PlacedInstance] = []
            for ins in self._placed:
                if ins.id not in self._selection:
                    continue
                p0 = self._shapes[ins.src_idx].poly_at_centroid
                pr = rotate(p0, -float(ins.rot % 360), origin=(0.0, 0.0))
                wp = translate(pr, xoff=ins.cx + ox, yoff=ins.cy + oy)
                batch_polys.append(wp)
                batch_rows.append(
                    PlacedInstance(
                        id=-1,
                        src_idx=ins.src_idx,
                        rot=ins.rot % 360,
                        cx=ins.cx + ox,
                        cy=ins.cy + oy,
                    )
                )
            ok = True
            for a in batch_polys:
                if not self._poly_fits_inner(a, inner):
                    ok = False
                    break
                for _pid, op in dyn_others:
                    if not _gap_ok(a, op, gp):
                        ok = False
                        break
                if not ok:
                    break
                for b in batch_polys:
                    if a is b:
                        continue
                    if not _gap_ok(a, b, gp):
                        ok = False
                        break
                if not ok:
                    break
            if not ok:
                continue
            for row, wp in zip(batch_rows, batch_polys):
                row.id = self._next_id
                self._next_id += 1
                new_instances.append(row)
                self._shapes[row.src_idx].placed_count += 1
                dyn_others.append((row.id, wp))

        for ni in new_instances:
            self._placed.append(ni)

        warned: set[int] = set()
        for ni in new_instances:
            sh = self._shapes[ni.src_idx]
            if sh.placed_count > sh.inventory and ni.src_idx not in warned:
                warned.add(ni.src_idx)
                messagebox.showwarning(
                    "超过库存提示",
                    f"「{sh.name}」库存为 {sh.inventory}，当前已放置 {sh.placed_count} 件。\n仍允许继续放置。",
                    parent=self.winfo_toplevel(),
                )
        self._rebuild_thumbs()

    def _ring_to_canvas_flat(
        self, bx: float, by: float, ring: list[tuple[float, float]]
    ) -> list[float]:
        flat: list[float] = []
        for px, py in ring:
            flat.extend([bx + px * self._scale, by + py * self._scale])
        return flat

    def _draw_poly_on_canvas(
        self,
        poly: Any,
        fill: str,
        outline: str,
        dash: Optional[tuple[int, ...]] = None,
    ) -> None:
        bx, by = self._ox, self._oy
        ext = list(poly.exterior.coords)
        fe = self._ring_to_canvas_flat(bx, by, ext)
        if len(fe) >= 6:
            kw: dict[str, Any] = {"outline": outline, "width": 1}
            if fill:
                kw["fill"] = fill
            else:
                kw["fill"] = ""
            if dash:
                kw["dash"] = dash
            self._canvas.create_polygon(*fe, **kw)
        for intr in poly.interiors:
            hi = self._ring_to_canvas_flat(bx, by, list(intr.coords))
            if len(hi) >= 6:
                self._canvas.create_polygon(
                    *hi, fill="#f7f7f7", outline=outline, width=1, dash=dash
                )

    def _redraw(self) -> None:
        if not dx.deps_available():
            return
        self._canvas.delete("all")
        gp, ge, sw, sh = self._nums()
        self._sheet_w, self._sheet_h = sw, sh

        cw = max(self._canvas.winfo_width(), 120)
        ch = max(self._canvas.winfo_height(), 120)
        pl, pr, pt, pb = (
            self._PREVIEW_PAD_L,
            self._PREVIEW_PAD_R,
            self._PREVIEW_PAD_T,
            self._PREVIEW_PAD_B,
        )
        avail_w = max(cw - pl - pr, 40)
        avail_h = max(ch - pt - pb, 40)
        scale_fit = min(avail_w / sw, avail_h / sh)
        self._scale = scale_fit
        draw_w = sw * self._scale
        draw_h = sh * self._scale
        self._ox = pl + (avail_w - draw_w) / 2
        self._oy = pt + (avail_h - draw_h) / 2

        self._canvas.create_rectangle(
            self._ox,
            self._oy,
            self._ox + draw_w,
            self._oy + draw_h,
            fill="#fffef6",
            outline="#303030",
            width=2,
        )

        ix0 = self._ox + ge * self._scale
        iy0 = self._oy + ge * self._scale
        ix1 = self._ox + (sw - ge) * self._scale
        iy1 = self._oy + (sh - ge) * self._scale
        self._canvas.create_rectangle(
            ix0, iy0, ix1, iy1, outline="#2a7a2a", width=1, dash=(4, 4)
        )

        for ins in self._placed:
            if self._mode == "move_one" and ins.id in self._move_ids:
                continue
            wp = self._world_poly(ins)
            sel = ins.id in self._selection
            self._draw_poly_on_canvas(
                wp,
                fill="#9fd3ff" if not sel else "#ffd080",
                outline="#1a4d7a",
            )

        if self._mode == "palette_ghost":
            gh = self._ghost_world_poly()
            if gh is not None:
                self._draw_poly_on_canvas(
                    gh, fill="", outline="#8040c0", dash=(4, 3)
                )
        elif self._mode == "move_one" and self._move_anchor_id is not None:
            for mpid in self._move_ids:
                gh = self._moved_part_poly_at(mpid, self._ghost_cx, self._ghost_cy)
                self._draw_poly_on_canvas(
                    gh, fill="", outline="#8040c0", dash=(4, 3)
                )

        for ox, oy in self._replica_ghost_offsets:
            for ins in self._placed:
                if ins.id not in self._selection:
                    continue
                p0 = self._shapes[ins.src_idx].poly_at_centroid
                pr = rotate(p0, -float(ins.rot % 360), origin=(0.0, 0.0))
                ghost = translate(pr, xoff=ins.cx + ox, yoff=ins.cy + oy)
                self._draw_poly_on_canvas(
                    ghost, fill="", outline="#a060a0", dash=(2, 4)
                )

        if self._rubber is not None:
            x0, y0, x1, y1 = self._rubber
            self._canvas.create_rectangle(
                x0, y0, x1, y1, outline="#0066cc", width=1, dash=(3, 2)
            )

        self._canvas.create_text(
            self._ox + draw_w / 2,
            self._oy + draw_h + 12,
            text=f"板料 W={sw:.0f} mm",
            fill="#111",
            font=("Segoe UI", 9, "bold"),
        )
        self._canvas.create_text(
            self._ox - 12,
            self._oy + draw_h / 2,
            text=f"H={sh:.0f}",
            angle=90,
            fill="#111",
            font=("Segoe UI", 9, "bold"),
        )
