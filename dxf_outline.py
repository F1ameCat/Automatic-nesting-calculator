"""从 DXF 提取零件封闭轮廓，并计算密铺单元尺寸（与 Shapely 配合）。"""

from __future__ import annotations

import atexit
import math
import os
import threading
from typing import Any, Optional

try:
    import ezdxf
    from ezdxf import const as ezdxf_const
except ImportError:
    ezdxf = None  # type: ignore
    ezdxf_const = None  # type: ignore

try:
    from shapely.affinity import rotate, scale, translate
    from shapely.geometry import LineString, MultiLineString, Polygon
    from shapely.ops import linemerge, polygonize, unary_union
except ImportError:
    LineString = None  # type: ignore
    MultiLineString = None  # type: ignore
    Polygon = None  # type: ignore
    translate = None  # type: ignore
    rotate = None  # type: ignore
    scale = None  # type: ignore
    linemerge = None  # type: ignore
    polygonize = None  # type: ignore
    unary_union = None  # type: ignore


def deps_available() -> bool:
    return ezdxf is not None and Polygon is not None


# 散线/圆弧拼合：Path 展平精度（DXF 图面单位，通常为 mm）
_WIRE_FLATTEN_DIST = 0.12
# 端点量化小数位，用于消除“肉眼相接但坐标不完全相等”导致的无法 linemerge
_WIRE_QUANTIZE_DECIMALS = 4
# 拼合出的碎面、噪点面积下限（图面单位²）
_WIRE_MIN_AREA = 1e-8

# 可用 make_path 再展平为折线、参与拓扑闭合的实体类型
_WIRE_ENTITY_TYPES = frozenset(
    {"LINE", "ARC", "CIRCLE", "LWPOLYLINE", "POLYLINE", "ELLIPSE", "SPLINE"}
)


def _path_entity_to_linestrings(entity: Any, flatten_dist: float) -> list[LineString]:
    """将单个 DXF 实体转为若干条 Shapely LineString（弧、圆已展平为折线）。"""
    if ezdxf is None or LineString is None:
        return []
    try:
        from ezdxf.path import make_path
    except ImportError:
        return []

    try:
        path = make_path(entity, segments=8, level=5)
    except (TypeError, ValueError, Exception):
        return []

    out: list[LineString] = []
    try:
        for sub in path.sub_paths():
            pts = [(float(v.x), float(v.y)) for v in sub.flattening(flatten_dist)]
            if len(pts) >= 2:
                out.append(LineString(pts))
    except Exception:
        return out
    return out


def _quantize_linestring(ls: LineString, ndp: int) -> LineString:
    coords = [(round(x, ndp), round(y, ndp)) for x, y in ls.coords]
    return LineString(coords)


def _linestrings_from_merged(merged: Any) -> list[LineString]:
    if merged is None or merged.is_empty:
        return []
    gt = merged.geom_type
    if gt == "LineString":
        return [merged]
    if gt == "MultiLineString":
        return list(merged.geoms)
    if gt == "GeometryCollection":
        return [
            g
            for g in merged.geoms
            if g is not None and not g.is_empty and g.geom_type == "LineString"
        ]
    return []


def _polygons_from_connected_wires(
    msp: Any,
    flatten_dist: float = _WIRE_FLATTEN_DIST,
    quantize_ndp: int = _WIRE_QUANTIZE_DECIMALS,
    min_area: float = _WIRE_MIN_AREA,
) -> list[Polygon]:
    """
    将模型空间中相互端点相接的 LINE / ARC / CIRCLE / 轻量多段线等拼成闭合环，
    再用 Shapely linemerge + polygonize 得到多边形（适用于未做成一条闭合多段线的情况）。
    """
    if (
        ezdxf is None
        or Polygon is None
        or LineString is None
        or MultiLineString is None
        or linemerge is None
        or polygonize is None
    ):
        return []

    segments: list[LineString] = []
    for e in msp:
        dt = e.dxftype()
        if dt == "HATCH":
            continue
        if dt == "POLYLINE" and getattr(e, "is_3d_polyline", False):
            continue
        if dt not in _WIRE_ENTITY_TYPES:
            continue
        for ls in _path_entity_to_linestrings(e, flatten_dist):
            if ls.length <= 0:
                continue
            segments.append(_quantize_linestring(ls, quantize_ndp))

    if len(segments) < 2:
        return []

    try:
        merged = linemerge(MultiLineString(segments))
    except Exception:
        return []

    lines = _linestrings_from_merged(merged)
    if not lines:
        return []

    polys: list[Polygon] = []
    try:
        for poly in polygonize(lines):
            fp = _fix_poly(poly)
            if fp is not None and fp.area >= min_area:
                polys.append(fp)
    except Exception:
        return []

    return polys


# $INSUNITS：将图面单位换算为 mm 的乘数
_INSUNITS_TO_MM: dict[int, float] = {
    0: 1.0,
    1: 25.4,
    2: 304.8,
    4: 1.0,
    5: 10.0,
    6: 1000.0,
    7: 1e-6,
    8: 1e-3,
    9: 0.0254,
    10: 914.4,
}


def _units_to_mm_factor(doc: Any) -> float:
    try:
        code = int(doc.header.get("$INSUNITS", 4))
    except (TypeError, ValueError):
        return 1.0
    return float(_INSUNITS_TO_MM.get(code, 1.0))


def _scale_poly(poly: Polygon, factor: float) -> Polygon:
    if abs(factor - 1.0) < 1e-12:
        return poly
    from shapely.affinity import scale

    return scale(poly, xfact=factor, yfact=factor, origin=(0.0, 0.0))


def _poly_from_lwpline(e: Any) -> Optional[Polygon]:
    if not e.closed:
        return None
    pts = [(float(p[0]), float(p[1])) for p in e.get_points("xy")]
    if len(pts) < 3:
        return None
    if pts[0] != pts[-1]:
        pts = pts + [pts[0]]
    try:
        p = Polygon(pts)
    except Exception:
        return None
    if not p.is_valid:
        p = p.buffer(0)
    return p if p.is_valid and not p.is_empty and p.area > 0 else None


def _poly_from_2d_polyline(e: Any) -> Optional[Polygon]:
    if not getattr(e, "is_closed", False):
        return None
    pts = []
    try:
        for v in e.vertices:
            loc = v.dxf.location
            pts.append((float(loc.x), float(loc.y)))
    except Exception:
        return None
    if len(pts) < 3:
        return None
    if pts[0] != pts[-1]:
        pts = pts + [pts[0]]
    try:
        p = Polygon(pts)
    except Exception:
        return None
    if not p.is_valid:
        p = p.buffer(0)
    return p if p.is_valid and not p.is_empty and p.area > 0 else None


def _poly_from_circle(e: Any) -> Polygon:
    c = e.dxf.center
    r = float(e.dxf.radius)
    cx, cy = float(c.x), float(c.y)
    n = 48
    pts = [
        (cx + r * math.cos(2 * math.pi * i / n), cy + r * math.sin(2 * math.pi * i / n))
        for i in range(n)
    ]
    return Polygon(pts)


def _poly_from_spline(e: Any) -> Optional[Polygon]:
    pts: list[tuple[float, float]] = []
    try:
        for p in e.flattening(distance=0.25):
            pts.append((float(p[0]), float(p[1])))
    except TypeError:
        try:
            for p in e.flattening(0.25):
                pts.append((float(p[0]), float(p[1])))
        except Exception:
            return None
    except Exception:
        return None
    if len(pts) < 3:
        return None
    if pts[0] != pts[-1]:
        pts.append(pts[0])
    try:
        p = Polygon(pts)
        if not p.is_valid:
            p = p.buffer(0)
        return p if p.is_valid and not p.is_empty and p.area > 0 else None
    except Exception:
        return None


def _fix_poly(p: Polygon) -> Optional[Polygon]:
    if p is None or p.is_empty:
        return None
    if not p.is_valid:
        p = p.buffer(0)
    return p if p.is_valid and not p.is_empty and p.area > 0 else None


def _select_outer_outline(candidates: list[Polygon]) -> Polygon:
    """
    在多个封闭图形中选取「零件外轮廓」：
    若某封闭区域完全落在外层更大的封闭区域内（典型：孔洞圆在板材矩形内），则排除内层，
    避免把孔误判为零件。
    """
    fixed: list[Polygon] = []
    for p in candidates:
        q = _fix_poly(p)
        if q is not None:
            fixed.append(q)
    if not fixed:
        raise ValueError("没有有效的封闭轮廓")

    fixed.sort(key=lambda x: x.area, reverse=True)
    kept: list[Polygon] = []
    for p in fixed:
        rp = p.representative_point()
        inside_larger = False
        for q in fixed:
            if q is p:
                continue
            if q.area <= p.area + 1e-9:
                continue
            try:
                if q.contains(rp) or q.covers(p):
                    inside_larger = True
                    break
            except Exception:
                try:
                    if q.buffer(1e-5).contains(rp):
                        inside_larger = True
                        break
                except Exception:
                    pass
        if not inside_larger:
            kept.append(p)

    if not kept:
        return max(fixed, key=lambda x: x.area)
    return max(kept, key=lambda x: x.area)


def _boundary_path_to_ring(path: Any, flatten_dist: float = 0.25) -> Optional[list[tuple[float, float]]]:
    """HATCH 边界路径 → 闭合点列（OCS XY）。"""
    if ezdxf is None:
        return None
    try:
        from ezdxf.entities.boundary_paths import EdgePath, PolylinePath, flatten_to_polyline_path
    except ImportError:
        return None

    pts: list[tuple[float, float]]
    if isinstance(path, PolylinePath):
        pts = [(float(t[0]), float(t[1])) for t in path.vertices]
    elif isinstance(path, EdgePath):
        try:
            np = flatten_to_polyline_path(path, flatten_dist)
            pts = [(float(t[0]), float(t[1])) for t in np.vertices]
        except Exception:
            return None
    else:
        return None

    if len(pts) < 3:
        return None
    if pts[0] != pts[-1]:
        pts = pts + [pts[0]]
    return pts


def _polygon_from_hatch_entity(h: Any) -> Optional[Polygon]:
    """
    从 HATCH 构造 Shapely 多边形（外环 + 孔）。
    - 带 EXTERNAL 标志的边界视为外轮廓；
    - 其余非 TEXTBOX 边界视为孔（与 ezdxf/geo 中孔用 OUTERMOST 的常见导出一致）。
    """
    if ezdxf is None or ezdxf_const is None or Polygon is None:
        return None

    paths = list(h.paths)
    if not paths:
        return None

    outers: list[list[tuple[float, float]]] = []
    holes: list[list[tuple[float, float]]] = []
    fallback_rings: list[list[tuple[float, float]]] = []

    for path in paths:
        ring = _boundary_path_to_ring(path)
        if not ring:
            continue
        fl = int(getattr(path, "path_type_flags", 0))
        if fl & ezdxf_const.BOUNDARY_PATH_TEXTBOX:
            continue
        if fl & ezdxf_const.BOUNDARY_PATH_EXTERNAL:
            outers.append(ring)
        else:
            holes.append(ring)
        fallback_rings.append(ring)

    if not outers:
        polys: list[Polygon] = []
        for ring in fallback_rings:
            try:
                polys.append(Polygon(ring))
            except Exception:
                continue
        if not polys:
            return None
        return _fix_poly(_select_outer_outline(polys))

    shell = outers[0]
    hole_rings = [hr for hr in holes if len(hr) >= 4]
    try:
        if hole_rings:
            p = Polygon(shell, holes=[tuple(hr) for hr in hole_rings])
        else:
            p = Polygon(shell)
    except Exception:
        p = Polygon(shell)
    return _fix_poly(p)


def _poly_from_ellipse(e: Any) -> Polygon:
    c = e.dxf.center
    ax = float(e.major_axis[0])
    ay = float(e.major_axis[1])
    ratio = float(e.dxf.ratio)
    major = math.hypot(ax, ay)
    minor = major * ratio
    ang = math.atan2(ay, ax)
    n = 48
    pts = []
    for i in range(n):
        t = 2 * math.pi * i / n
        x = major * math.cos(t)
        y = minor * math.sin(t)
        xr = x * math.cos(ang) - y * math.sin(ang)
        yr = x * math.sin(ang) + y * math.cos(ang)
        pts.append((float(c.x) + xr, float(c.y) + yr))
    return Polygon(pts)


def _merge_shell_with_inner_voids(shell: Polygon, candidates: list[Polygon]) -> Polygon:
    """
    当外轮廓已是最大外包络、但图内还有完全落在外形内的封闭图形（孔）时，
    将其并入为 Polygon 的内环（用于带孔板材）。
    若 shell 已由 HATCH 带孔，则不再合并。
    """
    if shell.interiors:
        return shell
    holes_rings: list[tuple[tuple[float, float], ...]] = []
    shell_area = shell.area
    for h in candidates:
        try:
            if h.equals(shell):
                continue
        except Exception:
            pass
        ha = h.area
        if ha <= 0 or ha >= shell_area * 0.98:
            continue
        try:
            rp = h.representative_point()
        except Exception:
            continue
        try:
            if not (shell.contains(rp) or h.within(shell)):
                continue
        except Exception:
            continue
        try:
            ring = tuple(h.exterior.coords)
            if len(ring) < 4:
                continue
            holes_rings.append(ring)
        except Exception:
            continue
    if not holes_rings:
        return shell
    try:
        merged = Polygon(shell.exterior.coords, holes=holes_rings)
        out = _fix_poly(merged)
        return out if out is not None else shell
    except Exception:
        return shell


def load_largest_outline_polygon(path: str) -> tuple[Polygon, str]:
    """
    读取 DXF 模型空间零件外轮廓：
    - 支持 LWPOLYLINE / POLYLINE / CIRCLE / ELLIPSE / SPLINE / HATCH；
    - 首尾相接的 LINE / ARC 等散线会尝试 linemerge + polygonize 拼成闭合区域；
    - 多个封闭图形时排除落在外层内部的图形（避免把孔洞圆当成零件）；
    - 若最大外形为实面、其内还有更小封闭区域，则自动合并为带孔多边形。
    """
    if not deps_available():
        raise RuntimeError("需要安装 ezdxf 与 shapely：pip install ezdxf shapely")

    doc = ezdxf.readfile(path)
    msp = doc.modelspace()
    fac = _units_to_mm_factor(doc)
    unit_note = f"按 DXF $INSUNITS 换算为 mm（系数 {fac}）"

    candidates: list[Polygon] = []

    wire_polys = _polygons_from_connected_wires(msp)
    if wire_polys:
        unit_note += "；已尝试将 LINE/ARC 等散线拼合为闭合区域"
    for wp in wire_polys:
        fp = _fix_poly(wp)
        if fp is not None:
            candidates.append(fp)

    for e in msp:
        dt = e.dxftype()
        p = None
        if dt == "HATCH":
            p = _polygon_from_hatch_entity(e)
        elif dt == "LWPOLYLINE":
            p = _poly_from_lwpline(e)
        elif dt == "POLYLINE" and not e.is_3d_polyline:
            p = _poly_from_2d_polyline(e)
        elif dt == "CIRCLE":
            p = _poly_from_circle(e)
        elif dt == "ELLIPSE":
            p = _poly_from_ellipse(e)
        elif dt == "SPLINE":
            p = _poly_from_spline(e)
        if p is not None:
            p = _fix_poly(p)
            if p is not None:
                candidates.append(p)

    if not candidates:
        raise ValueError(
            "未找到可用的封闭轮廓。可检查：\n"
            "1）端点是否精确相接（或尝试在 CAD 中用 JOIN/PEDIT 连接后重存）；\n"
            "2）或使用闭合 LWPOLYLINE、HATCH、整圆等；\n"
            "3）带孔件可用 HATCH 或「外轮廓+孔」图元组合。"
        )

    best = _select_outer_outline(candidates)
    best = _merge_shell_with_inner_voids(best, candidates)
    best = _scale_poly(best, fac)
    if not best.is_valid:
        best = best.buffer(0)
    if best.is_empty or best.area <= 0:
        raise ValueError("轮廓无效")

    return best, unit_note


def prepare_cell(
    poly_mm: Polygon, gap_part_mm: float
) -> tuple[float, float, Polygon]:
    """
    用 gap_part/2 作等距缓冲，保证相邻两件轮廓之间净距至少为 gap_part；
    返回 (单元宽, 单元高, 平移后轮廓：缓冲包络左下角在原点)。
    """
    if gap_part_mm < 0:
        raise ValueError("零件间隙不能为负数")
    buf = float(gap_part_mm) / 2.0
    q = poly_mm.buffer(buf)
    if q.is_empty:
        raise ValueError("缓冲后轮廓为空")
    minx, miny, maxx, maxy = q.bounds
    cell_w = maxx - minx
    cell_h = maxy - miny
    if cell_w <= 0 or cell_h <= 0:
        raise ValueError("轮廓尺寸无效")
    p_draw = translate(poly_mm, xoff=-minx, yoff=-miny)
    return cell_w, cell_h, p_draw


def _prepare_interlock_cell(
    poly_mm: Polygon, gap_part_mm: float
) -> tuple[float, float, Polygon]:
    """
    列向互嵌专用：以轮廓真实轴对齐包络为基准，宽高各加一整份零件间隙（与「净距≥gap」的矩形格下限一致），
    不做 buffer 外包络，避免单元虚大导致列距/行距偏松、互嵌不紧。
    返回 (cell_w, cell_h, 平移后轮廓：包络左下角在原点)。
    """
    if gap_part_mm < 0:
        raise ValueError("零件间隙不能为负数")
    gap = float(gap_part_mm)
    minx, miny, _, _ = poly_mm.bounds
    p0 = translate(poly_mm, xoff=-minx, yoff=-miny)
    bx = p0.bounds
    if bx[0] < -1e-6 or bx[1] < -1e-6:
        p0 = translate(p0, xoff=-bx[0], yoff=-bx[1])
        bx = p0.bounds
    w = bx[2] - bx[0]
    h = bx[3] - bx[1]
    return w + gap, h + gap, p0


def _transform_pose(
    poly: Polygon, rot_deg: int, mirror_x: bool, mirror_y: bool
) -> Polygon:
    """
    以轮廓初始质心为基准：先水平/垂直镜像（可选），再旋转 rot_deg°
    （旋转方向与历史逻辑一致：rotate(..., -rot_deg, origin=质心)）。
    """
    if scale is None or rotate is None:
        raise RuntimeError("需要 shapely.affinity 的 scale 与 rotate")
    ox, oy = float(poly.centroid.x), float(poly.centroid.y)
    origin = (ox, oy)
    p = poly
    if mirror_x:
        p = scale(p, -1.0, 1.0, origin=origin)
    if mirror_y:
        p = scale(p, 1.0, -1.0, origin=origin)
    rd = int(rot_deg) % 360
    if rd:
        p = rotate(p, -float(rd), origin=origin)
    return p


def best_orientation_and_cell(
    poly_mm: Polygon,
    gap_part_mm: float,
    inner_w: float,
    inner_h: float,
) -> tuple[int, bool, bool, float, float, Polygon, int, int]:
    """
    在 0°/180° 及横向/纵向镜像组合中选单张可放件数最多的姿态（不尝试 90°/270°）。
    返回 (rotation_deg, mirror_x, mirror_y, cell_w, cell_h, poly_draw, cols, rows)。
    """
    best_n = -1
    best_area = float("inf")
    best_pack: tuple[int, bool, bool, float, float, Polygon, int, int] | None = None

    for deg in (0, 180):
        for mx in (False, True):
            for my in (False, True):
                pr = _transform_pose(poly_mm, deg, mx, my)
                cw, ch, pd = prepare_cell(pr, gap_part_mm)
                cols = int(math.floor((inner_w + 1e-6) / cw)) if cw > 0 else 0
                rows = int(math.floor((inner_h + 1e-6) / ch)) if ch > 0 else 0
                n = cols * rows
                cell_a = cw * ch
                if n > best_n or (n == best_n and cell_a < best_area):
                    best_n = n
                    best_area = cell_a
                    best_pack = (deg, mx, my, cw, ch, pd, cols, rows)

    assert best_pack is not None
    return best_pack


# --- 密铺模式：标准网格 / 紧凑 / 交错行 / 列向互嵌（旋转+镜像 或 仅180°） ---
PACKING_MODE_GRID = "grid"
PACKING_MODE_COMPACT = "compact"
PACKING_MODE_BRICK = "brick"
# 奇数列可在「水平镜像 / 180° / 垂直镜像」中选优（实际常为 180°+镜像 组合）
PACKING_MODE_INTERLOCK_COL = "interlock_col"
# 奇数列仅相对偶数列做 180°，不使用镜像类奇列变体
PACKING_MODE_INTERLOCK_COL_ROT180 = "interlock_col_rot180"

_PACKING_GAP_TOL = 1e-4
_PACKING_GRID_CHECK_RADIUS = 3
# 交错行模式计算量远大于矩形格，单独用更小邻域与采样，避免界面卡死
_BRICK_LATTICE_RADIUS = 2
_BRICK_VY_SAMPLES = 36
_BRICK_VY_BINARY_STEPS = 18
_BRICK_GROW_MAX = 36
# 交错行仅尝试 0°/180°（与主密铺一致，不含 90°/270°）
_BRICK_ROTATIONS = (0, 180)


def _gap_satisfied(p0: Polygon, p1: Polygon, gap_mm: float) -> bool:
    return p0.distance(p1) >= gap_mm - _PACKING_GAP_TOL


def _poly_rot180_reanchor(p: Polygon) -> Polygon:
    """绕质心旋转 180° 后，将包络左下角移回原点（与 prepare_cell 后件位姿配套）。"""
    c = p.centroid
    r = rotate(p, 180.0, origin=(float(c.x), float(c.y)))
    minx, miny, _, _ = r.bounds
    return translate(r, xoff=-minx, yoff=-miny)


def _poly_mirror_x_reanchor(p: Polygon) -> Polygon:
    """关于过质心的竖直线镜像，包络左下角移回原点（左右对调，利于 L 形对扣）。"""
    if scale is None:
        raise RuntimeError("需要 shapely.affinity.scale")
    c = p.centroid
    ox, oy = float(c.x), float(c.y)
    r = scale(p, -1.0, 1.0, origin=(ox, oy))
    minx, miny, _, _ = r.bounds
    return translate(r, xoff=-minx, yoff=-miny)


def _poly_mirror_y_reanchor(p: Polygon) -> Polygon:
    """关于过质心的水平线镜像，包络左下角移回原点。"""
    if scale is None:
        raise RuntimeError("需要 shapely.affinity.scale")
    c = p.centroid
    ox, oy = float(c.x), float(c.y)
    r = scale(p, 1.0, -1.0, origin=(ox, oy))
    minx, miny, _, _ = r.bounds
    return translate(r, xoff=-minx, yoff=-miny)


def _poly_odd_for_interlock(p0: Polygon, kind: str) -> Polygon:
    if kind == "rot180":
        return _poly_rot180_reanchor(p0)
    if kind == "mirx":
        return _poly_mirror_x_reanchor(p0)
    if kind == "miry":
        return _poly_mirror_y_reanchor(p0)
    raise ValueError(f"unknown odd kind {kind!r}")


# 列向互嵌：奇数列相对 p0 的变体集合
INTERLOCK_ODD_KINDS_MIRROR = ("mirx", "rot180", "miry")
INTERLOCK_ODD_KINDS_ROT180_ONLY = ("rot180",)
_INTERLOCK_MAX_PARALLEL_TASKS = 24  # 2*2*2*3，进程池上限参考
# 列向错移 dy：可接受约 1mm 级离散化；滑动步长不得小于 1mm（更细无意义且拖慢）
_INTERLOCK_DY_STEP_MIN_MM = 1.0
# 奇数列 Y 向滑动：首轮/细化目标点数 + 缩窗轮数（实际点数随步长≥1mm 自动变少）
_INTERLOCK_DY_SLIDE_PTS_FIRST = 28
_INTERLOCK_DY_SLIDE_PTS_REFINE = 16
_INTERLOCK_DY_SLIDE_PASSES = 3
_INTERLOCK_DY_WINDOW_SHRINK = 0.34
# 缩窗半宽至少 1mm，与步长一致
_INTERLOCK_DY_MIN_HALF_WIDTH_MM = 1.0
# 寻优阶段用略小邻域加速，最后对若干最优候选用标准半径复核
_INTERLOCK_DY_SCAN_LATTICE_RADIUS = 2
_INTERLOCK_DY_VERIFY_TOP_K = 6
# 列向互嵌专用：列距二分搜索略疏，显著减少 translate+distance 次数（仅互嵌调用）
_INTERLOCK_MIN_HX_SAMPLES = 96
_INTERLOCK_MIN_HX_BINARY = 28

# 列向互嵌多进程池：复用工作进程，避免每次排版在 Windows 上反复 spawn（开销极大）
_interlock_executor = None
_interlock_executor_lock = threading.Lock()


def _shutdown_interlock_executor() -> None:
    global _interlock_executor
    with _interlock_executor_lock:
        ex = _interlock_executor
        _interlock_executor = None
    if ex is not None:
        ex.shutdown(wait=True)


def _get_interlock_executor():
    """懒创建全局 ProcessPoolExecutor；DXF_INTERLOCK_PARALLEL=0 时不使用。"""
    global _interlock_executor
    with _interlock_executor_lock:
        if _interlock_executor is None:
            import multiprocessing as mp
            from concurrent.futures import ProcessPoolExecutor

            nw = max(
                1,
                min(_INTERLOCK_MAX_PARALLEL_TASKS, (os.cpu_count() or 1)),
            )
            ctx = mp.get_context("spawn")
            _interlock_executor = ProcessPoolExecutor(
                max_workers=nw,
                mp_context=ctx,
            )
    return _interlock_executor


atexit.register(_shutdown_interlock_executor)


def _interlock_parity_metrics(
    p0: Polygon,
    p_odd: Polygon,
    gap_part_mm: float,
    inner_w: float,
    inner_h: float,
    bbox_cw: float,
    bbox_ch: float,
    dy_odd: float,
    lattice_radius: int = _PACKING_GRID_CHECK_RADIUS,
    row_pitch_ch: Optional[float] = None,
) -> tuple[int, float, float, float, int, int]:
    """固定 dy_odd 时：邻域校验后的 (件数 n, 单元面积 cw*ch, cw, ch, cols, rows)。"""
    cw = _min_hx_alternating_cols_dy(
        p0,
        p_odd,
        gap_part_mm,
        dy_odd,
        samples=_INTERLOCK_MIN_HX_SAMPLES,
        binary_steps=_INTERLOCK_MIN_HX_BINARY,
    )
    if row_pitch_ch is not None:
        ch = float(row_pitch_ch)
    else:
        ch = _min_row_pitch_pair(p0, p_odd, gap_part_mm)
    cw0, ch0 = cw, ch
    cap_cw = max(
        bbox_cw * 2,
        p0.bounds[2] + p_odd.bounds[2] + gap_part_mm * 3,
    )
    cap_ch = max(
        bbox_ch * 2,
        bbox_ch + max(p0.bounds[3], p_odd.bounds[3]) + abs(dy_odd),
    )
    cw, ch = _grow_until_parity_lattice_ok(
        p0,
        p_odd,
        gap_part_mm,
        cw,
        ch,
        cap_cw,
        cap_ch,
        dy_odd,
        lattice_radius=lattice_radius,
    )
    cw, ch = _tighten_parity_cw_ch(
        p0,
        p_odd,
        gap_part_mm,
        dy_odd,
        lattice_radius,
        cw0,
        ch0,
        cw,
        ch,
    )
    cols, rows = _count_cols_rows_parity(
        p0, p_odd, inner_w, inner_h, cw, ch, dy_odd
    )
    cols, rows = _parity_shrink_count_if_union_overflow(
        inner_w, inner_h, cols, rows, cw, ch, p0, p_odd, dy_odd
    )
    n = cols * rows
    return n, cw * ch, cw, ch, cols, rows


def _dy_odd_search_bounds(p0: Polygon, p_odd: Polygon, gap_mm: float) -> tuple[float, float]:
    """奇数列相对偶数列的竖直错移搜索区间（对称，单位 mm）。"""
    h0 = float(p0.bounds[3])
    h1 = float(p_odd.bounds[3])
    span = max(h0, h1, 1e-3)
    margin = max(span * 0.62, (h0 + h1) * 0.42 + abs(gap_mm) * 3.0)
    return -margin, margin


def _interlock_dy_slide_positions(lo: float, hi: float, npts: int) -> list[float]:
    """
    在 [lo, hi] 上生成 dy 采样点；相邻点间距至少 _INTERLOCK_DY_STEP_MIN_MM，
    并保证包含两端（在容差内）。
    """
    lo = float(lo)
    hi = float(hi)
    if hi < lo:
        lo, hi = hi, lo
    w = hi - lo
    st_min = float(_INTERLOCK_DY_STEP_MIN_MM)
    if w <= 1e-12:
        return [round(lo, 6)]
    denom = max(int(npts) - 1, 1)
    step = max(w / denom, st_min)
    seq: list[float] = []
    x = lo
    guard = 0
    while x <= hi + 1e-9 and guard < 60000:
        seq.append(round(x, 6))
        x += step
        guard += 1
    if not seq:
        return [round(lo, 6)]
    if abs(seq[-1] - hi) > 1e-3:
        seq.append(round(hi, 6))
    # 去重（近重合点）
    seq.sort()
    out: list[float] = []
    for v in seq:
        if not out or abs(v - out[-1]) > 1e-6:
            out.append(v)
    return out


def _best_dy_odd_continuous(
    p0: Polygon,
    p_odd: Polygon,
    gap_part_mm: float,
    inner_w: float,
    inner_h: float,
    bbox_cw: float,
    bbox_ch: float,
) -> tuple[float, int, float, float, float, int, int]:
    """
    在 [lo,hi] 上多轮缩窗 + 采样，对 dy 寻优（相邻采样点间距至少 _INTERLOCK_DY_STEP_MIN_MM，默认 1mm）。
    主目标单张件数 n，次目标单元面积 cw*ch。
    寻优阶段用较小邻域半径加速；结束后对若干最优 dy 用标准邻域复核，避免漏检远处碰撞。
    """
    orig_lo, orig_hi = _dy_odd_search_bounds(p0, p_odd, gap_part_mm)
    lo, hi = orig_lo, orig_hi
    best_dy = 0.0
    best_n = -1
    best_area = float("inf")
    best_cw = best_ch = 0.0
    best_cols = best_rows = 0
    scan_r = _INTERLOCK_DY_SCAN_LATTICE_RADIUS
    ch_fix = _min_row_pitch_pair(p0, p_odd, gap_part_mm)
    top_pool: list[tuple[int, float, float]] = []  # (n, cell_a, dy)

    def _pool_push(n: int, cell_a: float, dy: float) -> None:
        top_pool.append((n, cell_a, dy))
        top_pool.sort(key=lambda t: (t[0], -t[1]), reverse=True)
        del top_pool[_INTERLOCK_DY_VERIFY_TOP_K:]

    for pass_i in range(_INTERLOCK_DY_SLIDE_PASSES):
        width = hi - lo
        if width <= 1e-12:
            break
        npts = (
            _INTERLOCK_DY_SLIDE_PTS_FIRST
            if pass_i == 0
            else _INTERLOCK_DY_SLIDE_PTS_REFINE
        )
        for dy in _interlock_dy_slide_positions(lo, hi, npts):
            n, cell_a, cw, ch, cols, rows = _interlock_parity_metrics(
                p0,
                p_odd,
                gap_part_mm,
                inner_w,
                inner_h,
                bbox_cw,
                bbox_ch,
                dy,
                lattice_radius=scan_r,
                row_pitch_ch=ch_fix,
            )
            _pool_push(n, cell_a, dy)
            if n > best_n or (n == best_n and cell_a < best_area):
                best_n = n
                best_area = cell_a
                best_dy = dy
                best_cw, best_ch = cw, ch
                best_cols, best_rows = cols, rows

        half = max(
            width * _INTERLOCK_DY_WINDOW_SHRINK * 0.5,
            _INTERLOCK_DY_MIN_HALF_WIDTH_MM,
        )
        lo = max(orig_lo, best_dy - half)
        hi = min(orig_hi, best_dy + half)

    merge_eps = max(float(_INTERLOCK_DY_STEP_MIN_MM) * 0.999, 1e-3)
    seen_dy: list[float] = []
    for _n_s, _a_s, dy_s in top_pool:
        if any(abs(dy_s - prev) < merge_eps for prev in seen_dy):
            continue
        seen_dy.append(dy_s)

    if not seen_dy:
        seen_dy = [best_dy]

    v_n, v_a = -1, float("inf")
    v_dy, v_cw, v_ch = 0.0, 0.0, 0.0
    v_cols, v_rows = 0, 0
    full_r = _PACKING_GRID_CHECK_RADIUS
    for dy_v in seen_dy:
        n, cell_a, cw, ch, cols, rows = _interlock_parity_metrics(
            p0,
            p_odd,
            gap_part_mm,
            inner_w,
            inner_h,
            bbox_cw,
            bbox_ch,
            dy_v,
            lattice_radius=full_r,
            row_pitch_ch=ch_fix,
        )
        if n > v_n or (n == v_n and cell_a < v_a):
            v_n, v_a = n, cell_a
            v_dy, v_cw, v_ch = dy_v, cw, ch
            v_cols, v_rows = cols, rows

    return v_dy, v_n, v_a, v_cw, v_ch, v_cols, v_rows


def _poly_at_grid_parity(
    p0: Polygon,
    p_odd: Polygon,
    col: int,
    row: int,
    cw: float,
    ch: float,
    dy_odd: float = 0.0,
) -> Polygon:
    if col % 2 == 0:
        return translate(p0, xoff=col * cw, yoff=row * ch)
    return translate(p_odd, xoff=col * cw, yoff=row * ch + dy_odd)


def _min_hx_alternating_cols_dy(
    p0: Polygon,
    p_odd: Polygon,
    gap_mm: float,
    dy_odd: float,
    samples: int = 120,
    binary_steps: int = 36,
) -> float:
    """偶数列 p0、奇数列 p_odd 且奇数列相对行基准下移 dy_odd 时，最小列距。"""
    w0 = p0.bounds[2] - p0.bounds[0]
    w1 = p_odd.bounds[2] - p_odd.bounds[0]
    lo = 1e-4
    hi = max(w0, w1) * 2.5 + abs(gap_mm) * 8.0 + abs(dy_odd)
    best = hi
    n = max(samples, 2)
    for k in range(n):
        hx = lo + (hi - lo) * k / (n - 1)
        q = translate(p_odd, xoff=hx, yoff=dy_odd)
        if _gap_satisfied(p0, q, gap_mm):
            best = min(best, hx)
    if best >= hi - 0.05:
        return max(w0, w1) + max(gap_mm, 0.0) * 2.0
    step = (hi - lo) / max(n - 1, 1)
    lo2 = max(lo, best - step * 2.0)
    hi2 = best
    nb = max(binary_steps, 8)
    for _ in range(nb):
        mid = (lo2 + hi2) / 2.0
        q = translate(p_odd, xoff=mid, yoff=dy_odd)
        if _gap_satisfied(p0, q, gap_mm):
            hi2 = mid
        else:
            lo2 = mid
    return max(hi2, lo)


def _min_row_pitch_pair(
    p0: Polygon, p_odd: Polygon, gap_mm: float
) -> float:
    """同列姿态下行距：取两种轮廓竖直安全步距的较大者。"""
    v0 = _min_positive_period_along_axis(p0, gap_mm, "y")
    v1 = _min_positive_period_along_axis(p_odd, gap_mm, "y")
    return max(v0, v1)


def _parity_col_lattice_multi_ok(
    p0: Polygon,
    p_odd: Polygon,
    cw: float,
    ch: float,
    gap_mm: float,
    radius: int,
    dy_odd: float = 0.0,
) -> bool:
    origin = _poly_at_grid_parity(p0, p_odd, 0, 0, cw, ch, dy_odd)
    for dr in range(-radius, radius + 1):
        for dc in range(-radius, radius + 1):
            if dr == 0 and dc == 0:
                continue
            other = _poly_at_grid_parity(p0, p_odd, dc, dr, cw, ch, dy_odd)
            if not _gap_satisfied(origin, other, gap_mm):
                return False
    return True


def _grow_until_parity_lattice_ok(
    p0: Polygon,
    p_odd: Polygon,
    gap_mm: float,
    cw: float,
    ch: float,
    cap_cw: float,
    cap_ch: float,
    dy_odd: float = 0.0,
    lattice_radius: int = _PACKING_GRID_CHECK_RADIUS,
) -> tuple[float, float]:
    if _parity_col_lattice_multi_ok(
        p0, p_odd, cw, ch, gap_mm, lattice_radius, dy_odd
    ):
        return cw, ch
    for _ in range(72):
        cw = min(cw * 1.035, cap_cw)
        ch = min(ch * 1.035, cap_ch)
        if _parity_col_lattice_multi_ok(
            p0, p_odd, cw, ch, gap_mm, lattice_radius, dy_odd
        ):
            return cw, ch
    return cap_cw, cap_ch


def _tighten_parity_cw_ch(
    p0: Polygon,
    p_odd: Polygon,
    gap_mm: float,
    dy_odd: float,
    lattice_radius: int,
    cw0: float,
    ch0: float,
    cw1: float,
    ch1: float,
) -> tuple[float, float]:
    """
    grow 后 cw/ch 往往偏大；在仍满足邻域间隙的前提下交替二分缩小列距与行距，使互嵌更紧。
    """
    cw, ch = cw1, ch1
    if not _parity_col_lattice_multi_ok(
        p0, p_odd, cw, ch, gap_mm, lattice_radius, dy_odd
    ):
        return cw1, ch1
    tol = 0.05
    steps = 28
    for _ in range(4):
        lo, hi = cw0, cw
        if hi > lo + tol:
            for __ in range(steps):
                if hi - lo < tol:
                    break
                mid = (lo + hi) / 2.0
                if _parity_col_lattice_multi_ok(
                    p0, p_odd, mid, ch, gap_mm, lattice_radius, dy_odd
                ):
                    hi = mid
                else:
                    lo = mid
            cw = hi
        lo, hi = ch0, ch
        if hi > lo + tol:
            for __ in range(steps):
                if hi - lo < tol:
                    break
                mid = (lo + hi) / 2.0
                if _parity_col_lattice_multi_ok(
                    p0, p_odd, cw, mid, gap_mm, lattice_radius, dy_odd
                ):
                    hi = mid
                else:
                    lo = mid
            ch = hi
    return cw, ch


def _parity_true_aabb(
    cols: int,
    rows: int,
    cw: float,
    ch: float,
    p0: Polygon,
    p_odd: Polygon,
    dy_odd: float,
) -> tuple[float, float, float, float]:
    minx = miny = float("inf")
    maxx = maxy = float("-inf")
    for c in range(cols):
        for r in range(rows):
            p = _poly_at_grid_parity(p0, p_odd, c, r, cw, ch, dy_odd)
            b = p.bounds
            minx = min(minx, b[0])
            miny = min(miny, b[1])
            maxx = max(maxx, b[2])
            maxy = max(maxy, b[3])
    return minx, miny, maxx, maxy


def _parity_shrink_count_if_union_overflow(
    inner_w: float,
    inner_h: float,
    cols: int,
    rows: int,
    cw: float,
    ch: float,
    p0: Polygon,
    p_odd: Polygon,
    dy_odd: float,
    eps_mm: float = 1.0,
) -> tuple[int, int]:
    """用真实轮廓并集轴对齐包络校验；略大于解析 footprint 时回退件数，避免预览/出图越界。"""
    if cols <= 0 or rows <= 0:
        return cols, rows
    eps = float(eps_mm)
    guard = 0
    while guard < (cols + rows) * 4 + 8:
        guard += 1
        minx, miny, maxx, maxy = _parity_true_aabb(
            cols, rows, cw, ch, p0, p_odd, dy_odd
        )
        if (
            maxx <= inner_w + eps
            and maxy <= inner_h + eps
            and minx >= -eps
            and miny >= -eps
        ):
            return cols, rows
        if rows <= 1 and cols <= 1:
            return cols, rows
        over_w = maxx - inner_w
        over_h = maxy - inner_h
        if over_h >= over_w and rows > 1:
            rows -= 1
        elif cols > 1:
            cols -= 1
        elif rows > 1:
            rows -= 1
        else:
            break
    return max(0, cols), max(0, rows)


def _footprint_parity_grid_dy(
    cols: int,
    rows: int,
    cw: float,
    ch: float,
    p0: Polygon,
    p_odd: Polygon,
    dy_odd: float,
) -> tuple[float, float]:
    """列向交替 + 奇数列竖直错移 dy_odd 时的外包宽高。"""
    if cols <= 0 or rows <= 0:
        return 0.0, 0.0
    w0, h0 = p0.bounds[2], p0.bounds[3]
    w1, h1 = p_odd.bounds[2], p_odd.bounds[3]
    max_right = 0.0
    for cc in range(cols):
        wcc = w0 if (cc % 2 == 0) else w1
        max_right = max(max_right, cc * cw + wcc)
    if cols == 1:
        max_top = (rows - 1) * ch + h0
    else:
        row_ext = max(h0, dy_odd + h1)
        max_top = (rows - 1) * ch + row_ext
    return max_right, max_top


def _count_cols_rows_parity(
    p0: Polygon,
    p_odd: Polygon,
    inner_w: float,
    inner_h: float,
    cw: float,
    ch: float,
    dy_odd: float = 0.0,
) -> tuple[int, int]:
    best_c, best_r, best_n = 0, 0, 0
    lim = 220
    h0, h1 = p0.bounds[3], p_odd.bounds[3]
    for c in range(1, lim + 1):
        if c == 1:
            row_ext = h0
        else:
            row_ext = max(h0, dy_odd + h1)
        if ch <= 0:
            continue
        r_cap = int(math.floor((inner_h - row_ext + 1e-9) / ch)) + 1
        r_cap = max(1, min(r_cap, lim))
        for r in range(r_cap, 0, -1):
            fw, fh = _footprint_parity_grid_dy(c, r, cw, ch, p0, p_odd, dy_odd)
            if fw <= inner_w + 1e-9 and fh <= inner_h + 1e-9:
                n = c * r
                if n > best_n:
                    best_n = n
                    best_c, best_r = c, r
                break
    return best_c, best_r


def _interlock_col_combo_worker(
    task: tuple[str, float, float, float, int, bool, bool, str],
) -> tuple[tuple[int, float], tuple[int, bool, bool, float, float, Polygon, Polygon, int, int, float, str]]:
    """
    子进程任务：对单一 (姿态, 奇列变体) 做 Y 向寻优（步长≥1mm）。
    入参为 (poly_wkt, gap_part_mm, inner_w, inner_h, deg, mx, my, odd_kind)。
    返回 (排序键 (n, -cell_area), 与 best_orientation_and_cell_interlock_cols 相同的 payload 元组)。
    """
    from shapely import wkt as shapely_wkt

    (
        poly_wkt,
        gap_part_mm,
        inner_w,
        inner_h,
        deg,
        mx,
        my,
        odd_kind,
    ) = task
    poly_mm = shapely_wkt.loads(poly_wkt)
    pr = _transform_pose(poly_mm, deg, mx, my)
    bbox_cw, bbox_ch, p0 = _prepare_interlock_cell(pr, gap_part_mm)
    p_odd = _poly_odd_for_interlock(p0, odd_kind)
    dy_odd, n, cell_a, cw, ch, cols, rows = _best_dy_odd_continuous(
        p0,
        p_odd,
        gap_part_mm,
        inner_w,
        inner_h,
        bbox_cw,
        bbox_ch,
    )
    key = (n, -cell_a)
    payload = (
        deg,
        mx,
        my,
        cw,
        ch,
        p0,
        p_odd,
        cols,
        rows,
        dy_odd,
        odd_kind,
    )
    return key, payload


def best_orientation_and_cell_interlock_cols(
    poly_mm: Polygon,
    gap_part_mm: float,
    inner_w: float,
    inner_h: float,
    *,
    odd_kinds: tuple[str, ...] = INTERLOCK_ODD_KINDS_MIRROR,
) -> tuple[int, bool, bool, float, float, Polygon, Polygon, int, int, float, str]:
    """
    相邻列交替：偶数列基准件 p0，奇数列在 odd_kinds 指定变体中选优
    （默认可含水平镜像 / 180° / 垂直镜像；仅 180° 互嵌时传入 INTERLOCK_ODD_KINDS_ROT180_ONLY）。
    奇数列竖直错移 dy 多轮缩窗采样（步长≥1mm；主：件数，次：单元面积）。
    返回 (rot_deg, mirror_x, mirror_y, cell_w, cell_h, p0_draw, p_odd_draw, cols, rows, dy_odd, odd_kind)。
    """
    from shapely import wkt as shapely_wkt

    if not odd_kinds:
        raise ValueError("odd_kinds 不能为空")

    poly_wkt = shapely_wkt.dumps(poly_mm)
    tasks: list[tuple[str, float, float, float, int, bool, bool, str]] = [
        (poly_wkt, gap_part_mm, inner_w, inner_h, deg, mx, my, odd_kind)
        for deg in (0, 180)
        for mx in (False, True)
        for my in (False, True)
        for odd_kind in odd_kinds
    ]

    env_par = os.environ.get("DXF_INTERLOCK_PARALLEL", "1").strip().lower()
    use_parallel = env_par not in ("0", "false", "no", "off")

    results: list[
        tuple[tuple[int, float], tuple[int, bool, bool, float, float, Polygon, Polygon, int, int, float, str]]
    ] = []
    if use_parallel and len(tasks) > 2:
        try:
            ex = _get_interlock_executor()
            results = list(ex.map(_interlock_col_combo_worker, tasks, chunksize=1))
        except Exception:
            results = []

    if not results:
        results = [_interlock_col_combo_worker(t) for t in tasks]

    best_key: tuple[int, float] = (-1, float("-inf"))
    best: tuple[int, bool, bool, float, float, Polygon, Polygon, int, int, float, str] | None = None
    for key, payload in results:
        if key > best_key:
            best_key = key
            best = payload

    assert best is not None
    return best


def _min_positive_period_along_axis(
    poly: Polygon, gap_mm: float, axis: str, samples: int = 120
) -> float:
    """沿 +X 或 +Y 平移一份相同轮廓，使两件净距≥gap 的最小正步距（用于凹槽互嵌时的列距/行距）。"""
    minx, miny, maxx, maxy = poly.bounds
    span = max(maxx - minx, maxy - miny, 1e-6)
    lo = 1e-4
    hi = span * 3.0 + abs(gap_mm) * 8.0
    best = hi
    n = max(samples, 2)
    for k in range(n):
        t = lo + (hi - lo) * k / (n - 1)
        if axis == "x":
            q = translate(poly, xoff=t, yoff=0)
        else:
            q = translate(poly, xoff=0, yoff=t)
        if _gap_satisfied(poly, q, gap_mm):
            best = min(best, t)
    if best >= hi - 0.05:
        return span + max(gap_mm, 0.0) * 2.0
    step = (hi - lo) / max(n - 1, 1)
    lo2 = max(lo, best - step * 2.0)
    hi2 = best
    for _ in range(48):
        mid = (lo2 + hi2) / 2.0
        if axis == "x":
            q = translate(poly, xoff=mid, yoff=0)
        else:
            q = translate(poly, xoff=0, yoff=mid)
        if _gap_satisfied(poly, q, gap_mm):
            hi2 = mid
        else:
            lo2 = mid
    return max(hi2, lo)


def _rect_lattice_multi_ok(
    poly: Polygon, gap_mm: float, cw: float, ch: float, radius: int
) -> bool:
    """矩形格点 (i*cw, j*ch) 上，检查 |i|,|j|≤radius 的平移是否与原点件满足间隙。"""
    for i in range(-radius, radius + 1):
        for j in range(-radius, radius + 1):
            if i == 0 and j == 0:
                continue
            q = translate(poly, xoff=i * cw, yoff=j * ch)
            if not _gap_satisfied(poly, q, gap_mm):
                return False
    return True


def _grow_until_rect_lattice_ok(
    poly: Polygon,
    gap_mm: float,
    cw: float,
    ch: float,
    cap_cw: float,
    cap_ch: float,
) -> tuple[float, float]:
    if _rect_lattice_multi_ok(poly, gap_mm, cw, ch, _PACKING_GRID_CHECK_RADIUS):
        return cw, ch
    for _ in range(72):
        cw = min(cw * 1.035, cap_cw)
        ch = min(ch * 1.035, cap_ch)
        if _rect_lattice_multi_ok(poly, gap_mm, cw, ch, _PACKING_GRID_CHECK_RADIUS):
            return cw, ch
    return cap_cw, cap_ch


def _count_cols_rows_rect(
    poly_draw: Polygon, inner_w: float, inner_h: float, cw: float, ch: float
) -> tuple[int, int]:
    minx, miny, maxx, maxy = poly_draw.bounds
    if cw <= 0 or ch <= 0:
        return 0, 0
    if inner_w < maxx - 1e-9 or inner_h < maxy - 1e-9:
        return 0, 0
    cols = int(math.floor((inner_w - maxx + 1e-9) / cw)) + 1
    rows = int(math.floor((inner_h - maxy + 1e-9) / ch)) + 1
    return max(0, cols), max(0, rows)


def _brick_cell_delta(dc: int, dr: int, hx: float, vy: float, stagger_x: float) -> tuple[float, float]:
    dx = dc * hx + (dr % 2) * stagger_x
    dy = dr * vy
    return dx, dy


def _brick_lattice_multi_ok(
    poly: Polygon,
    gap_mm: float,
    hx: float,
    vy: float,
    stagger_x: float,
    radius: int,
) -> bool:
    for dr in range(-radius, radius + 1):
        for dc in range(-radius, radius + 1):
            if dr == 0 and dc == 0:
                continue
            dx, dy = _brick_cell_delta(dc, dr, hx, vy, stagger_x)
            q = translate(poly, xoff=dx, yoff=dy)
            if not _gap_satisfied(poly, q, gap_mm):
                return False
    return True


def _min_vy_brick(
    poly: Polygon,
    gap_mm: float,
    hx: float,
    stagger_x: float,
    samples: int = 96,
    lattice_radius: int = _PACKING_GRID_CHECK_RADIUS,
    binary_steps: int = 40,
) -> float:
    span = max(poly.bounds[2] - poly.bounds[0], poly.bounds[3] - poly.bounds[1], 1e-6)
    lo = 1e-4
    hi = span * 3.5 + abs(gap_mm) * 8.0
    best = hi
    n = max(samples, 2)
    for k in range(n):
        vy = lo + (hi - lo) * k / (n - 1)
        if _brick_lattice_multi_ok(
            poly, gap_mm, hx, vy, stagger_x, lattice_radius
        ):
            best = min(best, vy)
    if best >= hi - 0.05:
        return span + max(gap_mm, 0.0) * 2.0
    step = (hi - lo) / max(n - 1, 1)
    lo2 = max(lo, best - step * 2.0)
    hi2 = best
    for _ in range(max(1, binary_steps)):
        mid = (lo2 + hi2) / 2.0
        if _brick_lattice_multi_ok(
            poly, gap_mm, hx, mid, stagger_x, lattice_radius
        ):
            hi2 = mid
        else:
            lo2 = mid
    return max(hi2, lo)


def _max_brick_extent(
    cols: int, rows: int, hx: float, vy: float, stagger_x: float, poly_draw: Polygon
) -> tuple[float, float]:
    """
    砖形排布：单元 (i,j) 左下角参考 (i*hx + (j%2)*st, j*vy)，O(1) 求整体包络宽高。
    """
    if cols <= 0 or rows <= 0:
        return 0.0, 0.0
    minx, miny, maxx, maxy = poly_draw.bounds
    st = max(stagger_x, 0.0)
    extra_st = st if rows >= 2 else 0.0
    min_l = minx
    max_r = (cols - 1) * hx + extra_st + maxx
    min_b = miny
    max_t = (rows - 1) * vy + maxy
    return max_r - min_l, max_t - min_b


def _count_cols_rows_brick(
    poly_draw: Polygon,
    inner_w: float,
    inner_h: float,
    hx: float,
    vy: float,
    stagger_x: float,
) -> tuple[int, int]:
    """在板内求最多件数：对每个列数取可行最大行数，避免 O(lim²) 全枚举。"""
    minx, miny, maxx, maxy = poly_draw.bounds
    ph = maxy - miny
    pw = maxx - minx
    if hx <= 0 or vy <= 0 or inner_w < pw - 1e-9 or inner_h < ph - 1e-9:
        return 0, 0
    r_cap = int(math.floor((inner_h - ph + 1e-9) / vy)) + 1
    r_cap = max(1, min(r_cap, 400))
    c_cap = min(
        400,
        max(2, int(math.floor((inner_w + max(stagger_x, 0.0)) / max(hx * 0.2, 1e-6))) + 3),
    )
    best_c, best_r, best_n = 0, 0, 0
    for c in range(1, c_cap + 1):
        for r in range(r_cap, 0, -1):
            w_need, h_need = _max_brick_extent(c, r, hx, vy, stagger_x, poly_draw)
            if w_need <= inner_w + 1e-9 and h_need <= inner_h + 1e-9:
                n = c * r
                if n > best_n:
                    best_n = n
                    best_c, best_r = c, r
                break
    return best_c, best_r


def best_orientation_and_cell_compact(
    poly_mm: Polygon,
    gap_part_mm: float,
    inner_w: float,
    inner_h: float,
) -> tuple[int, bool, bool, float, float, Polygon, int, int]:
    """
    在 0°/180° 及横向/纵向镜像下用「最小安全列距、行距」做矩形密铺；
    若多格点邻域检验无法通过，则在该姿态上退回与标准网格相同的包络步距。
    返回 (rot_deg, mirror_x, mirror_y, cell_w, cell_h, poly_draw, cols, rows)。
    """
    best_n = -1
    best_area = float("inf")
    best: tuple[int, bool, bool, float, float, Polygon, int, int] | None = None

    for deg in (0, 180):
        for mx in (False, True):
            for my in (False, True):
                pr = _transform_pose(poly_mm, deg, mx, my)
                bbox_cw, bbox_ch, pd = prepare_cell(pr, gap_part_mm)
                cw = _min_positive_period_along_axis(pd, gap_part_mm, "x")
                ch = _min_positive_period_along_axis(pd, gap_part_mm, "y")
                cw, ch = _grow_until_rect_lattice_ok(
                    pd, gap_part_mm, cw, ch, bbox_cw, bbox_ch
                )
                cols, rows = _count_cols_rows_rect(pd, inner_w, inner_h, cw, ch)
                n = cols * rows
                cell_a = cw * ch
                if n > best_n or (n == best_n and cell_a < best_area):
                    best_n = n
                    best_area = cell_a
                    best = (deg, mx, my, cw, ch, pd, cols, rows)

    assert best is not None
    return best


def best_orientation_and_cell_brick(
    poly_mm: Polygon,
    gap_part_mm: float,
    inner_w: float,
    inner_h: float,
) -> tuple[int, bool, bool, float, float, Polygon, int, int, float]:
    """
    交错行密铺：仅在 0°/180° 与镜像组合下搜索；邻域校验使用较小半径。
    返回 (rot_deg, mirror_x, mirror_y, hx, vy, poly_draw, cols, rows, stagger_x)。
    """
    best_n = -1
    best_area = float("inf")
    best: tuple[int, bool, bool, float, float, Polygon, int, int, float] | None = None
    lr = _BRICK_LATTICE_RADIUS

    for deg in _BRICK_ROTATIONS:
        for mx in (False, True):
            for my in (False, True):
                pr = _transform_pose(poly_mm, deg, mx, my)
                bbox_cw, bbox_ch, pd = prepare_cell(pr, gap_part_mm)
                hx0 = _min_positive_period_along_axis(pd, gap_part_mm, "x")
                hx0 = min(hx0, bbox_cw)
                stagger_candidates = [0.0]
                for k in (0.25, 0.5, 0.75):
                    s = hx0 * k
                    if s > 1e-3:
                        stagger_candidates.append(s)

                for st in stagger_candidates:
                    vy = _min_vy_brick(
                        pd,
                        gap_part_mm,
                        hx0,
                        st,
                        samples=_BRICK_VY_SAMPLES,
                        lattice_radius=lr,
                        binary_steps=_BRICK_VY_BINARY_STEPS,
                    )
                    hx = hx0
                    hx, vy = _grow_brick_hx_vy(
                        pd,
                        gap_part_mm,
                        hx,
                        vy,
                        st,
                        bbox_cw,
                        bbox_ch,
                        lattice_radius=lr,
                        max_grow=_BRICK_GROW_MAX,
                    )
                    cols, rows = _count_cols_rows_brick(
                        pd, inner_w, inner_h, hx, vy, st
                    )
                    n = cols * rows
                    cell_a = hx * vy
                    if n > best_n or (n == best_n and cell_a < best_area):
                        best_n = n
                        best_area = cell_a
                        best = (deg, mx, my, hx, vy, pd, cols, rows, st)

    assert best is not None
    return best


def _grow_brick_hx_vy(
    poly: Polygon,
    gap_mm: float,
    hx: float,
    vy: float,
    stagger_x: float,
    cap_hx: float,
    cap_vy: float,
    lattice_radius: int = _PACKING_GRID_CHECK_RADIUS,
    max_grow: int = 72,
) -> tuple[float, float]:
    if _brick_lattice_multi_ok(poly, gap_mm, hx, vy, stagger_x, lattice_radius):
        return hx, vy
    for _ in range(max(1, max_grow)):
        hx = min(hx * 1.03, cap_hx)
        vy = min(vy * 1.03, cap_vy)
        if _brick_lattice_multi_ok(poly, gap_mm, hx, vy, stagger_x, lattice_radius):
            return hx, vy
    return cap_hx, cap_vy


def layout_dxf_packing(
    poly_mm: Polygon,
    gap_part_mm: float,
    inner_w: float,
    inner_h: float,
    mode: str,
) -> tuple[int, float, float, Polygon, int, int, float, str, bool, bool, str, float]:
    """
    按密铺模式计算排版（含 0°/180° 与横/纵镜像优选，不含 90°/270°）。
    返回 (rotation_deg, cell_w, cell_h, poly_draw, cols, rows, stagger_x, hint,
          mirror_x, mirror_y, part_odd_wkt, interlock_dy_mm)。
    互嵌模式：奇数列 WKT + 列向竖直错移（mm）；其它模式末两项为 "" 与 0。
    """
    from shapely import wkt as shapely_wkt

    _kind_cn = {"mirx": "水平镜像", "rot180": "180°", "miry": "垂直镜像"}

    m = (mode or PACKING_MODE_GRID).strip().lower()
    if m == PACKING_MODE_INTERLOCK_COL:
        deg, mx, my, cw, ch, p0, p_odd, cols, rows, dy_odd, odd_kind = (
            best_orientation_and_cell_interlock_cols(
                poly_mm,
                gap_part_mm,
                inner_w,
                inner_h,
                odd_kinds=INTERLOCK_ODD_KINDS_MIRROR,
            )
        )
        w_odd = shapely_wkt.dumps(p_odd)
        kcn = _kind_cn.get(odd_kind, odd_kind)
        dy_note = f"；列向错移 {dy_odd:.2f} mm" if abs(dy_odd) > 1e-6 else ""
        hint = (
            f"列向180°旋转+镜像：奇数列相对偶数列为「{kcn}」"
            f"{dy_note}（奇数列 Y 向步长≥1mm 寻优，邻域间隙校验）"
        )
        return (
            deg,
            cw,
            ch,
            p0,
            cols,
            rows,
            0.0,
            hint,
            mx,
            my,
            w_odd,
            float(dy_odd),
        )
    if m == PACKING_MODE_INTERLOCK_COL_ROT180:
        deg, mx, my, cw, ch, p0, p_odd, cols, rows, dy_odd, odd_kind = (
            best_orientation_and_cell_interlock_cols(
                poly_mm,
                gap_part_mm,
                inner_w,
                inner_h,
                odd_kinds=INTERLOCK_ODD_KINDS_ROT180_ONLY,
            )
        )
        w_odd = shapely_wkt.dumps(p_odd)
        dy_note = f"；列向错移 {dy_odd:.2f} mm" if abs(dy_odd) > 1e-6 else ""
        hint = (
            "列向180°互嵌（无镜像奇列）：奇数列相对偶数列仅 180° 姿态"
            f"{dy_note}（Y 向步长≥1mm 寻优，邻域间隙校验）"
        )
        return (
            deg,
            cw,
            ch,
            p0,
            cols,
            rows,
            0.0,
            hint,
            mx,
            my,
            w_odd,
            float(dy_odd),
        )
    if m == PACKING_MODE_COMPACT:
        deg, mx, my, cw, ch, pd, cols, rows = best_orientation_and_cell_compact(
            poly_mm, gap_part_mm, inner_w, inner_h
        )
        return (
            deg,
            cw,
            ch,
            pd,
            cols,
            rows,
            0.0,
            "紧凑轴对齐（缩短列距/行距，已做多邻域间隙校验）",
            mx,
            my,
            "",
            0.0,
        )
    if m == PACKING_MODE_BRICK:
        deg, mx, my, hx, vy, pd, cols, rows, st = best_orientation_and_cell_brick(
            poly_mm, gap_part_mm, inner_w, inner_h
        )
        st_note = f"，行错开 {st:.2f} mm" if st > 1e-3 else "（无行错开）"
        return (
            deg,
            hx,
            vy,
            pd,
            cols,
            rows,
            st,
            f"交错行密铺{st_note}（0°/180°+镜像，邻域半径 {_BRICK_LATTICE_RADIUS}）",
            mx,
            my,
            "",
            0.0,
        )
    deg, mx, my, cw, ch, pd, cols, rows = best_orientation_and_cell(
        poly_mm, gap_part_mm, inner_w, inner_h
    )
    return (
        deg,
        cw,
        ch,
        pd,
        cols,
        rows,
        0.0,
        "标准轴对齐包络网格",
        mx,
        my,
        "",
        0.0,
    )


def union_tail_bounds_mm(
    remain: int,
    cols: int,
    cell_w: float,
    cell_h: float,
    poly_draw: Polygon,
    gap_edge_mm: float,
    stagger_x: float = 0.0,
) -> tuple[float, float]:
    """尾板上放 remain 件时，包住所有轮廓外形的轴对齐最小矩形尺寸（含两侧板边）。"""
    rows = int(math.ceil(remain / cols))
    ge = float(gap_edge_mm)
    st = float(stagger_x)
    parts: list[Polygon] = []
    for idx in range(remain):
        c = idx % cols
        r = idx // cols
        if r >= rows:
            break
        t = translate(
            poly_draw,
            xoff=ge + c * cell_w + (r % 2) * st,
            yoff=ge + r * cell_h,
        )
        parts.append(t)
    if not parts:
        return 2 * ge, 2 * ge
    u = unary_union(parts)
    minx, miny, maxx, maxy = u.bounds
    return maxx - minx + 2 * ge, maxy - miny + 2 * ge


def union_tail_bounds_mm_interlock_col(
    remain: int,
    cols: int,
    cell_w: float,
    cell_h: float,
    poly_even: Polygon,
    poly_odd: Polygon,
    gap_edge_mm: float,
    stagger_x: float = 0.0,
    dy_odd: float = 0.0,
) -> tuple[float, float]:
    """列向互嵌尾板包络（偶数列 poly_even，奇数列 poly_odd + 竖直错移 dy_odd）。"""
    rows = int(math.ceil(remain / cols))
    ge = float(gap_edge_mm)
    st = float(stagger_x)
    dy = float(dy_odd)
    parts: list[Polygon] = []
    for idx in range(remain):
        c = idx % cols
        r = idx // cols
        if r >= rows:
            break
        p = poly_odd if (c % 2) else poly_even
        y_extra = dy if (c % 2) else 0.0
        t = translate(
            p,
            xoff=ge + c * cell_w + (r % 2) * st,
            yoff=ge + r * cell_h + y_extra,
        )
        parts.append(t)
    if not parts:
        return 2 * ge, 2 * ge
    u = unary_union(parts)
    minx, miny, maxx, maxy = u.bounds
    return maxx - minx + 2 * ge, maxy - miny + 2 * ge
