"""Draw a parsed Chart onto a reportlab canvas.

The canvas must already be translated so that (0, 0) is the bottom-left of the
chart frame and y grows upward; ``draw`` receives the frame width and height in
points. Everything is drawn with plain vector primitives.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

from .chart import Axis, Chart, DataLabels, PlotGroup, Series, ShapeProps, TextStyle
from .colors import RGB, hex_to_rgb, lum_mod_off
from .numfmt import format_value
from .text import FontSpec, Line, TextEngine
from .theme import Theme

WHITE: RGB = (1.0, 1.0, 1.0)
BLACK: RGB = (0.0, 0.0, 0.0)
AXIS_GRAY: RGB = hex_to_rgb("868686")
GRID_GRAY: RGB = hex_to_rgb("D9D9D9")
BORDER_GRAY: RGB = hex_to_rgb("868686")

_TIERS = [(1.0, 0.0), (0.6, 0.0), (0.8, 0.2), (0.8, 0.0), (0.6, 0.4), (0.5, 0.0)]
_AUTO_MARKERS = ["diamond", "square", "triangle", "x", "star", "circle", "plus"]
_Z_ORDER = {"area": 0, "bar": 1, "bubble": 2, "line": 3, "scatter": 4, "radar": 5}

_DASHES = {
    "dash": [4, 3], "sysDash": [3, 1], "sysDot": [1, 1], "dot": [1, 2], "lgDash": [8, 3],
    "dashDot": [4, 2, 1, 2], "lgDashDot": [8, 3, 1, 3], "sysDashDot": [3, 1, 1, 1],
    "lgDashDotDot": [8, 3, 1, 3, 1, 3], "sysDashDotDot": [3, 1, 1, 1, 1, 1],
}


@dataclass
class Scale:
    lo: float
    hi: float
    step: float
    minor: float
    log_base: Optional[float] = None
    reverse: bool = False

    def frac(self, v: float) -> float:
        if self.log_base:
            v = max(v, 1e-300)
            lo, hi = math.log(self.lo, self.log_base), math.log(self.hi, self.log_base)
            f = (math.log(v, self.log_base) - lo) / ((hi - lo) or 1.0)
        else:
            f = (v - self.lo) / ((self.hi - self.lo) or 1.0)
        return 1.0 - f if self.reverse else f

    def ticks(self) -> List[float]:
        out: List[float] = []
        if self.log_base:
            v = self.lo
            while v <= self.hi * (1 + 1e-9):
                out.append(v)
                v *= self.log_base
            return out
        n = int(round((self.hi - self.lo) / self.step))
        for i in range(n + 1):
            out.append(_clean(self.lo + i * self.step))
        return out

    def minor_ticks(self) -> List[float]:
        if self.log_base or self.minor <= 0:
            return []
        n = int(round((self.hi - self.lo) / self.minor))
        return [_clean(self.lo + i * self.minor) for i in range(n + 1)]


def _clean(v: float) -> float:
    return float(f"{v:.12g}")


def _nice_step(raw: float) -> float:
    if raw <= 0 or not math.isfinite(raw):
        return 1.0
    exp = math.floor(math.log10(raw))
    base = 10 ** exp
    for m in (1, 2, 2.5, 5, 10):
        if raw <= m * base * (1 + 1e-9):
            if m == 2.5 and exp < 0:
                continue
            return m * base
    return 10 * base


@dataclass
class AxisInfo:
    axis: Axis
    role: str  # cat | val
    horizontal: bool
    side: str  # b t l r
    style: FontSpec
    fmt: str = "General"
    scale: Optional[Scale] = None
    n: int = 0
    labels: List[str] = field(default_factory=list)
    between: bool = True
    reverse: bool = False
    rotation: float = 0.0
    wrapped: Optional[List[List[Line]]] = None
    skip: int = 1
    thickness: float = 0.0
    title_thickness: float = 0.0

    def pos(self, i: float) -> float:
        """Fraction along the axis for category index i (may be fractional)."""
        if self.between:
            f = (i + 0.5) / max(self.n, 1)
        else:
            f = i / max(self.n - 1, 1)
        return 1.0 - f if self.reverse else f


class ChartRenderer:
    def __init__(self, text: TextEngine, theme: Theme) -> None:
        self.text = text
        self.theme = theme
        self.font = theme.minor_font or "Calibri"

    # ------------------------------------------------------------------
    # helpers
    def color_for(self, i: int) -> RGB:
        accent = hex_to_rgb(self.theme.colors.get(f"accent{i % 6 + 1}", "4472C4"))
        mod, off = _TIERS[(i // 6) % len(_TIERS)]
        return lum_mod_off(accent, mod, off) if (mod, off) != (1.0, 0.0) else accent

    def spec(self, st: TextStyle, size: float, bold: bool = False, color: RGB = BLACK) -> FontSpec:
        return FontSpec(
            family=st.font or self.font,
            size=st.size if st.size else size,
            bold=st.bold if st.bold is not None else bold,
            italic=bool(st.italic),
            color=st.color if st.color is not None else color,
        )

    def _set_line(self, color: RGB, width: float, dash: Optional[str] = None) -> None:
        c = self.c
        c.setStrokeColorRGB(*color)
        c.setLineWidth(width)
        c.setLineCap(0)
        c.setLineJoin(1)
        if dash and dash in _DASHES:
            c.setDash([d * max(width, 0.75) for d in _DASHES[dash]], 0)
        else:
            c.setDash([], 0)

    def _line_of(self, props: ShapeProps, default: Optional[RGB], width: float) -> Optional[Tuple[RGB, float, Optional[str]]]:
        if props.line_none:
            return None
        color = props.line if props.line is not None else default
        if color is None:
            return None
        return color, props.line_width if props.line_width is not None else width, props.dash

    def _fill_of(self, props: ShapeProps, default: Optional[RGB]) -> Optional[RGB]:
        if props.fill_none:
            return None
        return props.fill if props.fill is not None else default

    def _rect(self, x, y, w, h, fill: Optional[RGB], line: Optional[Tuple[RGB, float, Optional[str]]], radius: float = 0.0) -> None:
        c = self.c
        if fill is None and line is None:
            return
        if fill is not None:
            c.setFillColorRGB(*fill)
        if line is not None:
            self._set_line(*line)
        if radius > 0:
            c.roundRect(x, y, w, h, radius, stroke=1 if line else 0, fill=1 if fill else 0)
        else:
            c.rect(x, y, w, h, stroke=1 if line else 0, fill=1 if fill else 0)

    def _draw_text_block(self, lines: List[Line], x: float, y: float, align: str = "center",
                         valign: str = "middle", rotation: float = 0.0) -> None:
        """Draw lines around anchor (x, y). rotation in degrees counter-clockwise."""
        if not lines:
            return
        total_h = sum(ln.height for ln in lines)
        width = max(ln.width for ln in lines)
        c = self.c
        c.saveState()
        c.translate(x, y)
        if rotation:
            c.rotate(rotation)
        if valign == "middle":
            top = total_h / 2
        elif valign == "top":
            top = 0.0
        else:  # bottom
            top = total_h
        cur = top
        for ln in lines:
            baseline = cur - ln.ascent - 0.75
            if align == "center":
                lx = -ln.width / 2
            elif align == "right":
                lx = -ln.width
            else:
                lx = 0.0
            self.text.draw_line(c, ln, lx, baseline)
            cur -= ln.height
        c.restoreState()

    def _lines(self, text: str, spec: FontSpec, width: Optional[float] = None) -> List[Line]:
        return self.text.layout([(text, spec)], width)

    def _block_size(self, lines: List[Line]) -> Tuple[float, float]:
        return (max((ln.width for ln in lines), default=0.0), sum(ln.height for ln in lines))

    # ------------------------------------------------------------------
    def draw(self, canvas, chart: Chart, W: float, H: float) -> None:
        self.c = canvas
        self.chart = chart
        base = TextStyle(size=10.0, bold=False, italic=False, color=BLACK, font=self.font)
        self.base = chart.text_style.merged(base) if chart.has_text_props else base

        # chart area
        ap = chart.area_props
        fill = self._fill_of(ap, WHITE)
        line = self._line_of(ap, BORDER_GRAY if not ap.specified else None, 0.75)
        self._rect(0, 0, W, H, fill, line, radius=6.0 if chart.rounded else 0.0)

        pad = max(4.0, min(W, H) * 0.03)
        x0, y0, x1, y1 = pad, pad, W - pad, H - pad

        # title
        if chart.title is not None and chart.title.runs:
            t = chart.title
            st = t.style.merged(self.base)
            runs = []
            # Excel draws an unstyled title bold at the chart text size
            default_size = (st.size or 10.0) * 1.4 if chart.has_text_props else (self.base.size or 10.0)
            title_spec = self.spec(t.style.merged(TextStyle(font=self.base.font, color=self.base.color)),
                                   default_size, bold=True if t.style.bold is None else bool(t.style.bold))
            for text, rst in t.runs:
                rs = rst.merged(t.style)
                runs.append((text, self.spec(rs.merged(TextStyle(font=self.base.font, color=self.base.color)),
                                             title_spec.size, bold=title_spec.bold)))
            lines = self.text.layout(runs, (x1 - x0) * 0.9)
            _, th = self._block_size(lines)
            self._draw_text_block(lines, W / 2, y1, "center", "top")
            if not t.overlay:
                y1 -= th + 4

        # legend
        if chart.legend is not None:
            x0, y0, x1, y1 = self._legend(x0, y0, x1, y1)

        groups = [g for g in chart.groups if g.series]
        if not groups:
            return
        pie_like = [g for g in groups if g.type in ("pie", "doughnut")]
        if pie_like and len(pie_like) == len(groups):
            self._pie_chart(pie_like[0], x0, y0, x1, y1)
            return
        groups = [g for g in groups if g.type not in ("pie", "doughnut")]
        if all(g.type == "radar" for g in groups):
            self._radar_chart(groups, x0, y0, x1, y1)
            return
        self._axis_chart([g for g in groups if g.type != "radar"], W, H, x0, y0, x1, y1)

    # ------------------------------------------------------------------
    # legend
    def _legend_entries(self) -> List[Tuple[str, str, RGB, Optional[Tuple], Optional[str]]]:
        """(label, kind, color, line_info, marker_symbol)"""
        chart = self.chart
        groups = [g for g in chart.groups if g.series]
        out = []
        deleted = chart.legend.deleted if chart.legend else set()
        if len(groups) == 1 and groups[0].vary_colors and len(groups[0].series) == 1 and groups[0].type in ("pie", "doughnut", "bar"):
            s = groups[0].series[0]
            n = max(len(s.values), len(s.categories))
            for i in range(n):
                if i in deleted:
                    continue
                label = s.categories[i] if i < len(s.categories) and s.categories[i] else str(i + 1)
                pp = s.point_props.get(i, ShapeProps())
                color = self._fill_of(pp, None) or self._fill_of(s.props, None) if pp.specified else self.color_for(i)
                out.append((label, "box", color or self.color_for(i), None, None))
            return out
        for g in groups:
            for s in g.series:
                if s.idx in deleted:
                    continue
                kind = "line" if g.type in ("line", "scatter", "radar") else "box"
                color = self.series_color(g, s)
                line_info = None
                marker = None
                if kind == "line":
                    line_info = self.series_line(g, s)
                    marker = self.marker_symbol(g, s)
                    if g.type == "radar" and g.radar_style == "filled":
                        kind = "box"
                out.append((s.name, kind, color, line_info, marker))
        if chart.legend and chart.legend.pos == "r" and any(g.type == "bar" and g.bar_dir == "bar" and g.grouping != "clustered" for g in groups):
            out.reverse()
        return out

    def _legend(self, x0, y0, x1, y1):
        lg = self.chart.legend
        entries = self._legend_entries()
        if not entries:
            return x0, y0, x1, y1
        spec = self.spec(lg.style.merged(self.base), self.base.size or 10.0)
        sym_w = {"box": spec.size * 0.7, "line": spec.size * 2.0}
        gap = 4.0
        items = []
        for label, kind, color, line_info, marker in entries:
            ln = self._lines(label, spec)
            w, h = self._block_size(ln)
            items.append((ln, kind, color, line_info, marker, sym_w[kind] + gap + w, h))
        row_h = max(it[6] for it in items)
        c = self.c
        if lg.pos in ("r", "l", "tr"):
            max_w = (x1 - x0) * 0.4
            lw = min(max(it[5] for it in items), max_w)
            lh = row_h * len(items)
            lx = x1 - lw if lg.pos in ("r", "tr") else x0
            if lg.pos == "tr":
                ly_top = y1
            else:
                ly_top = (y0 + y1) / 2 + lh / 2
            cy = ly_top
            for ln, kind, color, line_info, marker, w, h in items:
                self._legend_symbol(lx, cy - row_h / 2, kind, color, line_info, marker, sym_w[kind], spec.size)
                c.saveState()
                p = c.beginPath()
                p.rect(lx, cy - row_h, lw, row_h)
                c.clipPath(p, stroke=0, fill=0)
                self._draw_text_block(ln, lx + sym_w[kind] + gap, cy - row_h / 2, "left", "middle")
                c.restoreState()
                cy -= row_h
            if not lg.overlay:
                if lg.pos in ("r", "tr"):
                    x1 = lx - 8
                else:
                    x0 = lx + lw + 8
        else:
            avail = x1 - x0
            rows: List[List] = [[]]
            rw = 0.0
            for it in items:
                if rows[-1] and rw + it[5] > avail:
                    rows.append([])
                    rw = 0.0
                rows[-1].append(it)
                rw += it[5] + 10
            lh = row_h * len(rows)
            ly_top = y0 + lh if lg.pos == "b" else y1
            cy = ly_top
            for row in rows:
                total = sum(it[5] for it in row) + 10 * (len(row) - 1)
                cx = (x0 + x1) / 2 - total / 2
                for ln, kind, color, line_info, marker, w, h in row:
                    self._legend_symbol(cx, cy - row_h / 2, kind, color, line_info, marker, sym_w[kind], spec.size)
                    self._draw_text_block(ln, cx + sym_w[kind] + gap, cy - row_h / 2, "left", "middle")
                    cx += w + 10
                cy -= row_h
            if not lg.overlay:
                if lg.pos == "b":
                    y0 = y0 + lh + 6
                else:
                    y1 = y1 - lh - 6
        return x0, y0, x1, y1

    def _legend_symbol(self, x, ymid, kind, color, line_info, marker, w, size) -> None:
        c = self.c
        if kind == "box":
            s = size * 0.7
            c.setFillColorRGB(*color)
            c.rect(x, ymid - s / 2, s, s, stroke=0, fill=1)
        else:
            if line_info:
                lc, lw, dash = line_info
                self._set_line(lc, min(lw, 3.0), dash)
                c.line(x, ymid, x + w, ymid)
                c.setDash([], 0)
            if marker:
                self._marker(x + w / 2, ymid, marker, min(size * 0.6, 7.0), color, color)

    # ------------------------------------------------------------------
    # series styling
    def series_color(self, g: PlotGroup, s: Series) -> RGB:
        if g.type in ("line", "scatter", "radar") and not (g.type == "radar" and g.radar_style == "filled"):
            if s.props.line is not None:
                return s.props.line
            if s.marker and s.marker.props.fill is not None and s.props.line_none:
                return s.marker.props.fill
            return self.color_for(s.idx)
        if s.props.fill is not None:
            return s.props.fill
        return self.color_for(s.idx)

    def series_line(self, g: PlotGroup, s: Series):
        if s.props.line_none:
            return None
        color = s.props.line if s.props.line is not None else self.color_for(s.idx)
        width = s.props.line_width if s.props.line_width is not None else 2.25
        return color, width, s.props.dash

    def marker_symbol(self, g: PlotGroup, s: Series) -> Optional[str]:
        sym = s.marker.symbol if s.marker else None
        if sym == "none":
            return None
        if sym and sym != "auto":
            return sym
        if g.type == "line" and not g.show_markers:
            return None
        if g.type == "scatter" and g.scatter_style in ("line", "smooth"):
            return None
        if g.type == "radar" and g.radar_style != "marker":
            return None
        return _AUTO_MARKERS[s.idx % len(_AUTO_MARKERS)]

    def _marker(self, x, y, symbol, size, fill: Optional[RGB], line: Optional[RGB]) -> None:
        c = self.c
        r = size / 2
        if fill is not None:
            c.setFillColorRGB(*fill)
        stroke = 1 if line is not None else 0
        if line is not None:
            self._set_line(line, 0.75)
        filled = 1 if fill is not None else 0
        if symbol in ("circle", "dot"):
            c.circle(x, y, r if symbol == "circle" else r / 2, stroke=stroke, fill=filled)
        elif symbol == "square":
            c.rect(x - r, y - r, size, size, stroke=stroke, fill=filled)
        elif symbol in ("diamond", "triangle"):
            p = c.beginPath()
            if symbol == "diamond":
                p.moveTo(x, y + r); p.lineTo(x + r, y); p.lineTo(x, y - r); p.lineTo(x - r, y)
            else:
                p.moveTo(x, y + r); p.lineTo(x + r, y - r); p.lineTo(x - r, y - r)
            p.close()
            c.drawPath(p, stroke=stroke, fill=filled)
        elif symbol in ("x", "star", "plus"):
            self._set_line(line or fill or BLACK, 1.0)
            if symbol in ("x", "star"):
                c.line(x - r, y - r, x + r, y + r)
                c.line(x - r, y + r, x + r, y - r)
            if symbol in ("plus", "star"):
                c.line(x, y - r, x, y + r)
            if symbol == "plus":
                c.line(x - r, y, x + r, y)
        elif symbol == "dash":
            c.rect(x - r, y - r / 4, size, r / 2, stroke=0, fill=1)

    # ------------------------------------------------------------------
    # axis charts
    def _scale(self, lo: float, hi: float, ax: Axis, length: float, horizontal: bool,
               spec: FontSpec, fmt: str) -> Scale:
        if ax.log_base and ax.log_base > 1:
            b = ax.log_base
            lo_v = ax.min if ax.min else b ** math.floor(math.log(max(lo, 1e-12), b))
            hi_v = ax.max if ax.max else b ** math.ceil(math.log(max(hi, 1e-12), b))
            if hi_v <= lo_v:
                hi_v = lo_v * b
            return Scale(lo_v, hi_v, 1, 0, b, ax.reverse)
        if not math.isfinite(lo) or not math.isfinite(hi):
            lo, hi = 0.0, 1.0
        fixed_lo, fixed_hi = ax.min, ax.max
        if lo == hi:
            if lo == 0:
                hi = 1.0
            elif lo > 0:
                lo = 0.0
            else:
                hi = 0.0
        # Excel starts the axis at zero unless the data sits in a narrow band
        if fixed_lo is None and lo > 0 and (hi - lo) / hi > 1 / 6:
            lo = 0.0
        if fixed_hi is None and hi < 0 and (hi - lo) / -lo > 1 / 6:
            hi = 0.0
        lo_d = fixed_lo if fixed_lo is not None else lo
        hi_d = fixed_hi if fixed_hi is not None else hi
        if hi_d <= lo_d:
            hi_d = lo_d + 1
        if ax.major_unit:
            step = ax.major_unit
        else:
            if horizontal:
                sample = format_value(max(abs(lo_d), abs(hi_d)), fmt).text
                per = self.text.measure([(sample + "00", spec)]) + spec.size
            else:
                per = spec.size * 2.0
            intervals = max(1.0, length / max(per, 1.0))
            span = (hi_d - lo_d)
            if fixed_hi is None:
                span *= 1.05
            step = _nice_step(span / intervals)
        span = hi_d - lo_d
        if fixed_lo is None:
            if lo_d < 0:
                lo_v = math.floor((lo_d - span * 0.05) / step + 1e-9) * step
            else:
                lo_v = math.floor(lo_d / step + 1e-9) * step
        else:
            lo_v = fixed_lo
        if fixed_hi is None:
            if hi_d > 0:
                hi_v = math.ceil((hi_d + span * 0.05) / step - 1e-9) * step
            else:
                hi_v = math.ceil(hi_d / step - 1e-9) * step
            if hi_v <= lo_v:
                hi_v = lo_v + step
        else:
            hi_v = fixed_hi
        minor = ax.minor_unit or step / 5
        return Scale(_clean(lo_v), _clean(hi_v), step, minor, None, ax.reverse)

    def _value_format(self, ax: Axis, groups: List[PlotGroup]) -> str:
        if any(g.grouping == "percentStacked" for g in groups):
            if not ax.num_fmt or ax.source_linked:
                return "0%"
        if ax.num_fmt and not ax.source_linked:
            return ax.num_fmt
        for g in groups:
            for s in g.series:
                if s.format_code and s.format_code != "General":
                    return s.format_code
        return ax.num_fmt or "General"

    def _axis_chart(self, groups: List[PlotGroup], W, H, x0, y0, x1, y1) -> None:
        chart = self.chart
        c = self.c
        groups = sorted(groups, key=lambda g: _Z_ORDER.get(g.type, 9))
        horizontal_bars = any(g.type == "bar" and g.bar_dir == "bar" for g in groups)
        xy = any(g.type in ("scatter", "bubble") for g in groups)

        axes = dict(chart.axes)
        group_axes: List[Tuple[PlotGroup, str, str]] = []
        for g in groups:
            ids = list(g.ax_ids) + ["__c", "__v"]
            cid, vid = ids[0], ids[1]
            if cid not in axes:
                axes[cid] = Axis(id=cid, kind="val" if g.type in ("scatter", "bubble") else "cat",
                                 pos="l" if horizontal_bars else "b")
            if vid not in axes:
                axes[vid] = Axis(id=vid, kind="val", pos="b" if horizontal_bars else "l")
            group_axes.append((g, cid, vid))

        # ---- categories
        n_cat = 0
        cat_labels: List[str] = []
        for g, _, _ in group_axes:
            if g.type in ("scatter", "bubble"):
                continue
            for s in g.series:
                n_cat = max(n_cat, len(s.values), len(s.categories))
                if not cat_labels and any(s.categories):
                    cat_labels = list(s.categories)
        cat_labels = (cat_labels + [str(i + 1) for i in range(len(cat_labels), n_cat)])[:n_cat]

        # ---- value ranges per value axis
        ranges: Dict[str, List[float]] = {}

        def extend(aid: str, *vals: float) -> None:
            r = ranges.setdefault(aid, [math.inf, -math.inf])
            for v in vals:
                if v is not None and math.isfinite(v):
                    r[0] = min(r[0], v)
                    r[1] = max(r[1], v)

        for g, cid, vid in group_axes:
            if g.type in ("scatter", "bubble"):
                for s in g.series:
                    xs = self._xvalues(s)
                    extend(cid, *[x for x, y in zip(xs, s.values) if y is not None])
                    extend(vid, *[v for v in s.values if v is not None])
                continue
            stacked = g.grouping in ("stacked", "percentStacked")
            if stacked:
                for i in range(n_cat):
                    pos = sum(max(0.0, _at(s.values, i) or 0.0) for s in g.series)
                    neg = sum(min(0.0, _at(s.values, i) or 0.0) for s in g.series)
                    if g.grouping == "percentStacked":
                        tot = sum(abs(_at(s.values, i) or 0.0) for s in g.series) or 1.0
                        pos, neg = pos / tot, neg / tot
                    extend(vid, pos, neg)
            else:
                for s in g.series:
                    extend(vid, *[v for v in s.values if v is not None])
        if any(g.grouping == "percentStacked" for g in groups):
            for g, _, vid in group_axes:
                if g.grouping == "percentStacked":
                    r = ranges.setdefault(vid, [0.0, 1.0])
                    r[0] = min(r[0], 0.0)
                    r[1] = max(r[1], 1.0) if r[1] > 0 else r[1]

        # ---- axis infos
        infos: Dict[str, AxisInfo] = {}
        for g, cid, vid in group_axes:
            if cid not in infos:
                ax = axes[cid]
                if g.type in ("scatter", "bubble"):
                    horiz = True
                    infos[cid] = AxisInfo(ax, "val", horiz, _side(ax.pos, horiz),
                                          self.spec(ax.style.merged(self.base), self.base.size or 10.0),
                                          fmt=ax.num_fmt if ax.num_fmt and not ax.source_linked else (self._x_format(g) or "General"))
                else:
                    horiz = not horizontal_bars
                    info = AxisInfo(ax, "cat", horiz, _side(ax.pos, horiz),
                                    self.spec(ax.style.merged(self.base), self.base.size or 10.0))
                    info.n = max(n_cat, 1)
                    info.labels = cat_labels
                    info.reverse = ax.reverse
                    infos[cid] = info
            if vid not in infos:
                ax = axes[vid]
                horiz = horizontal_bars and not xy
                vgroups = [gg for gg, _, v in group_axes if v == vid]
                side = _side(ax.pos, horiz)
                if ax.crosses == "max":
                    # an axis that crosses the category axis at its maximum sits on the far side
                    side = {"l": "r", "b": "t"}.get(side, side)
                infos[vid] = AxisInfo(ax, "val", horiz, side,
                                      self.spec(ax.style.merged(self.base), self.base.size or 10.0),
                                      fmt=self._value_format(ax, vgroups))
        # crossBetween comes from the value axis crossing the category axis
        for g, cid, vid in group_axes:
            ci = infos[cid]
            if ci.role == "cat":
                cb = axes[vid].cross_between
                if cb is None:
                    cb = "midCat" if g.type == "area" and not any(gg.type == "bar" for gg in groups) else "between"
                if any(gg.type == "bar" for gg in groups):
                    cb = "between"
                ci.between = cb != "midCat"

        # ---- layout loop
        if chart.plot_layout:
            lx, ly, lw, lh = chart.plot_layout
            plot = (W * lx, H * (1 - ly - lh), W * (lx + lw), H * (1 - ly))
            self._compute_axes(infos, ranges, plot)
        else:
            plot = (x0, y0, x1, y1)
            for _ in range(3):
                self._compute_axes(infos, ranges, plot)
                m = {"l": 0.0, "r": 0.0, "t": 0.0, "b": 0.0}
                for info in infos.values():
                    if info.axis.delete:
                        continue
                    m[info.side] += info.thickness + info.title_thickness
                # horizontal value-axis labels overhang the plot edges by half a label
                new_plot = (x0 + m["l"] + 2, y0 + m["b"] + 2, x1 - m["r"] - 2, y1 - m["t"] - 2)
                if new_plot[2] - new_plot[0] < 20 or new_plot[3] - new_plot[1] < 20:
                    new_plot = (x0, y0, x1, y1)
                if all(abs(a - b) < 0.5 for a, b in zip(new_plot, plot)):
                    plot = new_plot
                    break
                plot = new_plot
            self._compute_axes(infos, ranges, plot)
        px0, py0, px1, py1 = plot
        pw, ph = px1 - px0, py1 - py0
        if pw <= 0 or ph <= 0:
            return

        # ---- plot area background
        pp = chart.plot_props
        self._rect(px0, py0, pw, ph, self._fill_of(pp, None), self._line_of(pp, None, 0.75))

        # ---- gridlines
        for info in infos.values():
            self._gridlines(info, plot)

        # ---- series
        c.saveState()
        p = c.beginPath()
        p.rect(px0 - 1, py0 - 1, pw + 2, ph + 2)
        c.clipPath(p, stroke=0, fill=0)
        label_jobs: List = []
        for g, cid, vid in group_axes:
            ci, vi = infos[cid], infos[vid]
            if g.type == "bar":
                self._bars(g, ci, vi, plot, label_jobs)
            elif g.type in ("line",):
                self._lines_series(g, ci, vi, plot, label_jobs)
            elif g.type == "area":
                self._areas(g, ci, vi, plot)
            elif g.type in ("scatter", "bubble"):
                self._scatter(g, ci, vi, plot, label_jobs)
        c.restoreState()

        # ---- axes
        primary_val = next((infos[v] for _, _, v in group_axes), None)
        for info in infos.values():
            self._axis(info, plot, infos, group_axes, primary_val)

        # ---- data labels on top
        for job in label_jobs:
            self._draw_text_block(*job)

    def _x_format(self, g: PlotGroup) -> Optional[str]:
        return None

    def _xvalues(self, s: Series) -> List[float]:
        if s.cat_values and all(v is not None for v in s.cat_values[: len(s.values)]) and len(s.cat_values) >= len(s.values):
            return [float(v) for v in s.cat_values]  # type: ignore[arg-type]
        return [float(i + 1) for i in range(len(s.values))]

    def _compute_axes(self, infos: Dict[str, AxisInfo], ranges: Dict[str, List[float]], plot) -> None:
        px0, py0, px1, py1 = plot
        for aid, info in infos.items():
            length = (px1 - px0) if info.horizontal else (py1 - py0)
            ax = info.axis
            if info.role == "val":
                lo, hi = ranges.get(aid, [0.0, 1.0])
                if lo == math.inf:
                    lo, hi = 0.0, 1.0
                info.scale = self._scale(lo, hi, ax, length, info.horizontal, info.style, info.fmt)
                info.labels = [format_value(v, info.fmt).text for v in info.scale.ticks()]
            self._plan_labels(info, length)
            if ax.title is not None and ax.title.runs and not ax.delete:
                tspec = self.spec(ax.title.style.merged(self.base), self.base.size or 10.0, bold=True)
                tl = self.text.layout([(t, tspec) for t, _ in ax.title.runs], length)
                info.title_thickness = self._block_size(tl)[1] + 4
            else:
                info.title_thickness = 0.0

    def _plan_labels(self, info: AxisInfo, length: float) -> None:
        ax = info.axis
        spec = info.style
        info.wrapped = None
        info.rotation = 0.0
        info.skip = 1
        if ax.delete or ax.label_pos == "none":
            info.thickness = 0.0
            return
        tick = 4.0 if ax.major_tick in ("out", "cross") else 0.0
        line_h = self.text.line_height(spec)
        labels = info.labels
        if not labels:
            info.thickness = tick
            return
        widths = [self.text.measure([(t, spec)]) for t in labels]
        explicit_rot = ax.style.rotation
        if info.role == "cat":
            band = length / max(info.n, 1)
            if info.horizontal:
                if explicit_rot not in (None, 0.0):
                    info.rotation = -explicit_rot
                elif max(widths) > band * 0.95:
                    wrapped = [self._lines(t, spec, max(band * 0.95, spec.size)) for t in labels]
                    ok = all(len(w) <= 3 and all(ln.width <= band for ln in w) for w in wrapped)
                    # wrapping only works when no single word is wider than the band
                    if ok and all(self.text.measure([(word, spec)]) <= band for t in labels for word in t.split()):
                        info.wrapped = wrapped
                    else:
                        info.rotation = 45.0
                if ax.tick_skip:
                    info.skip = ax.tick_skip
                elif info.rotation:
                    info.skip = max(1, math.ceil(line_h * 1.1 / max(band, 0.1)))
                if info.wrapped is not None:
                    th = max(sum(ln.height for ln in w) for w in info.wrapped)
                elif info.rotation:
                    a = math.radians(abs(info.rotation))
                    th = max(w * math.sin(a) + line_h * math.cos(a) for w in widths)
                    th = min(th, length * 0.6 if length > 0 else th)
                else:
                    th = line_h
                info.thickness = tick + 2 + th
            else:
                info.skip = ax.tick_skip or max(1, math.ceil(line_h / max(band, 0.1)))
                info.thickness = tick + 4 + max(widths)
        else:
            if info.horizontal:
                info.thickness = tick + 2 + line_h
            else:
                info.thickness = tick + 4 + max(widths)

    def _gridlines(self, info: AxisInfo, plot) -> None:
        px0, py0, px1, py1 = plot
        c = self.c
        ax = info.axis
        for grid, minor in ((ax.minor_grid, True), (ax.major_grid, False)):
            if grid is None:
                continue
            line = self._line_of(grid, AXIS_GRAY if not minor else GRID_GRAY, 0.75)
            if line is None:
                continue
            self._set_line(*line)
            positions: List[float] = []
            if info.role == "val" and info.scale:
                ticks = info.scale.minor_ticks() if minor else info.scale.ticks()
                positions = [info.scale.frac(v) for v in ticks]
            elif info.role == "cat" and not minor:
                n = max(info.n, 1)
                if info.between:
                    positions = [i / n for i in range(n + 1)]
                else:
                    positions = [info.pos(i) for i in range(n)]
            for f in positions:
                if not -1e-6 <= f <= 1 + 1e-6:
                    continue
                if info.horizontal:
                    x = px0 + f * (px1 - px0)
                    c.line(x, py0, x, py1)
                else:
                    y = py0 + f * (py1 - py0)
                    c.line(px0, y, px1, y)
            c.setDash([], 0)

    def _cross_position(self, info: AxisInfo, plot, infos, group_axes, primary_val) -> float:
        """Where along the perpendicular direction this axis line sits (absolute coordinate)."""
        px0, py0, px1, py1 = plot
        if info.side in ("l", "r") and info.role == "val" and not info.horizontal:
            return px0 if info.side == "l" else px1
        if info.role == "cat" or (info.role == "val" and info.horizontal):
            # category axis (or X axis) crosses the value axis at zero
            other = None
            for g, cid, vid in group_axes:
                if info.axis.id == cid:
                    other = infos[vid]
                    break
                if info.axis.id == vid and not info.horizontal:
                    other = infos[cid]
            if other is not None and other.scale is not None:
                ax = info.axis
                if ax.crosses == "max":
                    f = 1.0 if not other.scale.reverse else 0.0
                elif ax.crosses_at is not None:
                    f = other.scale.frac(ax.crosses_at)
                elif ax.crosses == "min":
                    f = 0.0 if not other.scale.reverse else 1.0
                else:
                    zero = 0.0 if other.scale.lo <= 0 <= other.scale.hi else (other.scale.lo if other.scale.lo > 0 else other.scale.hi)
                    f = other.scale.frac(zero) if not other.scale.log_base else 0.0
                f = min(max(f, 0.0), 1.0)
                if other.horizontal:
                    return px0 + f * (px1 - px0)
                return py0 + f * (py1 - py0)
        if info.horizontal:
            return py0 if info.side == "b" else py1
        return px0 if info.side == "l" else px1

    def _axis(self, info: AxisInfo, plot, infos, group_axes, primary_val) -> None:
        ax = info.axis
        if ax.delete:
            return
        c = self.c
        px0, py0, px1, py1 = plot
        pw, ph = px1 - px0, py1 - py0
        cross = self._cross_position(info, plot, infos, group_axes, primary_val)
        line = self._line_of(ax.props, AXIS_GRAY if not ax.props.specified else None, 0.75)
        if line is not None:
            self._set_line(*line)
            if info.horizontal:
                c.line(px0, cross, px1, cross)
            else:
                c.line(cross, py0, cross, py1)
            # tick marks
            tick_fracs: List[float] = []
            if info.role == "val" and info.scale:
                tick_fracs = [info.scale.frac(v) for v in info.scale.ticks()]
            else:
                n = max(info.n, 1)
                tick_fracs = [i / n for i in range(n + 1)] if info.between else [info.pos(i) for i in range(n)]
            out_sign = -1.0 if info.side in ("b", "l") else 1.0
            t = ax.major_tick
            if t != "none":
                a = 4.0 if t in ("out", "cross") else 0.0
                b = 4.0 if t in ("in", "cross") else 0.0
                for f in tick_fracs:
                    if info.horizontal:
                        x = px0 + f * pw
                        c.line(x, cross + out_sign * a, x, cross - out_sign * b)
                    else:
                        y = py0 + f * ph
                        c.line(cross + out_sign * a, y, cross - out_sign * b, y)
            c.setDash([], 0)

        # labels at the plot edge on the axis side
        if ax.label_pos != "none" and info.labels:
            tick = 4.0 if ax.major_tick in ("out", "cross") else 0.0
            if ax.label_pos == "high":
                edge_side = {"b": "t", "t": "b", "l": "r", "r": "l"}[info.side]
            else:
                edge_side = info.side
            if info.horizontal:
                edge = py0 if edge_side == "b" else py1
                if ax.label_pos == "nextTo" and info.role == "val":
                    edge = cross if edge_side == "b" else edge
            else:
                edge = px0 if edge_side == "l" else px1
            if info.role == "val" and info.scale:
                fracs = [info.scale.frac(v) for v in info.scale.ticks()]
            else:
                fracs = [info.pos(i) for i in range(len(info.labels))]
            for i, (label, f) in enumerate(zip(info.labels, fracs)):
                if info.role == "cat" and i % info.skip:
                    continue
                if not label:
                    continue
                if info.horizontal:
                    x = px0 + f * pw
                    if edge_side == "b":
                        y = edge - tick - 2
                        if info.wrapped is not None:
                            self._draw_text_block(info.wrapped[i], x, y, "center", "top")
                        elif info.rotation:
                            lines = self._lines(label, info.style)
                            self._draw_text_block(lines, x + 2, y, "right", "middle", rotation=info.rotation)
                        else:
                            self._draw_text_block(self._lines(label, info.style), x, y, "center", "top")
                    else:
                        y = edge + tick + 2
                        self._draw_text_block(self._lines(label, info.style), x, y, "center", "bottom")
                else:
                    y = py0 + f * ph
                    if edge_side == "l":
                        self._draw_text_block(self._lines(label, info.style), edge - tick - 3, y, "right", "middle")
                    else:
                        self._draw_text_block(self._lines(label, info.style), edge + tick + 3, y, "left", "middle")

        # axis title
        if ax.title is not None and ax.title.runs:
            tspec = self.spec(ax.title.style.merged(self.base), self.base.size or 10.0, bold=True)
            lines = self.text.layout([(t, tspec) for t, _ in ax.title.runs], pw if info.horizontal else ph)
            if info.horizontal:
                if info.side == "b":
                    y = py0 - info.thickness - 2
                    self._draw_text_block(lines, (px0 + px1) / 2, y, "center", "top")
                else:
                    y = py1 + info.thickness + 2
                    self._draw_text_block(lines, (px0 + px1) / 2, y, "center", "bottom")
            else:
                if info.side == "l":
                    x = px0 - info.thickness - 2
                    self._draw_text_block(lines, x, (py0 + py1) / 2, "center", "bottom", rotation=90)
                else:
                    x = px1 + info.thickness + 2
                    self._draw_text_block(lines, x, (py0 + py1) / 2, "center", "bottom", rotation=-90)

    # ------------------------------------------------------------------
    def _value_to(self, vi: AxisInfo, v: float, plot) -> float:
        px0, py0, px1, py1 = plot
        f = vi.scale.frac(v) if vi.scale else 0.0
        return px0 + f * (px1 - px0) if vi.horizontal else py0 + f * (py1 - py0)

    def _cat_to(self, ci: AxisInfo, i: float, plot) -> float:
        px0, py0, px1, py1 = plot
        f = ci.pos(i)
        return px0 + f * (px1 - px0) if ci.horizontal else py0 + f * (py1 - py0)

    def _base_value(self, vi: AxisInfo) -> float:
        s = vi.scale
        if s is None:
            return 0.0
        if s.log_base:
            return s.lo
        if s.lo <= 0 <= s.hi:
            return 0.0
        return s.lo if s.lo > 0 else s.hi

    def _labels_for(self, g: PlotGroup, s: Series) -> Optional[DataLabels]:
        if s.labels is not None:
            return s.labels if s.labels.any else None
        if g.labels is not None and g.labels.any:
            return g.labels
        return None

    def _label_text(self, dl: DataLabels, s: Series, i: int, value: Optional[float], pct: Optional[float] = None) -> str:
        parts = []
        if dl.show_ser:
            parts.append(s.name)
        if dl.show_cat:
            parts.append(s.categories[i] if i < len(s.categories) else str(i + 1))
        if dl.show_val and value is not None:
            parts.append(format_value(value, dl.num_fmt or s.format_code or "General").text)
        if dl.show_pct and pct is not None:
            parts.append(format_value(pct, dl.num_fmt if dl.num_fmt and "%" in dl.num_fmt else "0%").text)
        return (dl.separator or ", ").join(parts)

    def _label_spec(self, dl: DataLabels) -> FontSpec:
        return self.spec(dl.style.merged(self.base), (self.base.size or 10.0) * 0.9)

    def _bars(self, g: PlotGroup, ci: AxisInfo, vi: AxisInfo, plot, label_jobs) -> None:
        c = self.c
        n = max(ci.n, 1)
        px0, py0, px1, py1 = plot
        length = (px1 - px0) if ci.horizontal else (py1 - py0)
        band = length / n
        stacked = g.grouping in ("stacked", "percentStacked")
        S = len(g.series)
        overlap = (g.overlap if g.overlap is not None else (100.0 if stacked else 0.0)) / 100.0
        gap = g.gap_width / 100.0
        k = 1 if stacked else S
        bw = band / (k - (k - 1) * overlap + gap) if k else band
        cluster = bw * (k - (k - 1) * overlap)
        base_v = self._base_value(vi)
        pos_cum = [0.0] * n
        neg_cum = [0.0] * n
        totals = [sum(abs(_at(s.values, i) or 0.0) for s in g.series) or 1.0 for i in range(n)]
        vary = g.vary_colors and S == 1
        for si, s in enumerate(g.series):
            color = self.series_color(g, s)
            line = self._line_of(s.props, None, 0.75)
            dl = self._labels_for(g, s)
            for i in range(n):
                v = _at(s.values, i)
                if v is None:
                    continue
                if g.grouping == "percentStacked":
                    v_plot = v / totals[i]
                else:
                    v_plot = v
                if stacked:
                    if v_plot >= 0:
                        a, b = pos_cum[i], pos_cum[i] + v_plot
                        pos_cum[i] = b
                    else:
                        a, b = neg_cum[i], neg_cum[i] + v_plot
                        neg_cum[i] = b
                    offset = 0.0
                else:
                    a, b = base_v, v_plot
                    offset = si * bw * (1 - overlap)
                center = self._cat_to(ci, i, plot)
                start = center - band / 2 + (band - cluster) / 2 + offset if not ci.reverse else \
                    center + band / 2 - (band - cluster) / 2 - offset - bw
                va, vb = self._value_to(vi, a, plot), self._value_to(vi, b, plot)
                fill = color
                pp = s.point_props.get(i)
                if vary and (pp is None or not pp.specified):
                    fill = self.color_for(i)
                lp = line
                if pp is not None and pp.specified:
                    fill = self._fill_of(pp, fill)
                    lp = self._line_of(pp, line[0] if line else None, line[1] if line else 0.75)
                if s.invert_if_negative and v < 0 and not (pp and pp.specified):
                    fill = WHITE
                    lp = lp or (color, 0.75, None)
                if ci.horizontal:
                    rx, ry, rw, rh = start, min(va, vb), bw, abs(vb - va)
                else:
                    rx, ry, rw, rh = min(va, vb), start, abs(vb - va), bw
                self._rect(rx, ry, rw, rh, fill, lp)
                if dl is not None:
                    text = self._label_text(dl, s, i, v)
                    if text:
                        spec = self._label_spec(dl)
                        pos = dl.position or ("ctr" if stacked else "outEnd")
                        lines = self._lines(text, spec)
                        positive = vb >= va
                        if ci.horizontal:
                            x = rx + rw / 2
                            if pos == "outEnd":
                                job = (lines, x, vb + (2 if positive else -2), "center", "bottom" if positive else "top")
                            elif pos == "inEnd":
                                job = (lines, x, vb - (2 if positive else -2), "center", "top" if positive else "bottom")
                            elif pos == "inBase":
                                job = (lines, x, va + (2 if positive else -2), "center", "bottom" if positive else "top")
                            else:
                                job = (lines, x, (va + vb) / 2, "center", "middle")
                        else:
                            y = ry + rh / 2
                            if pos == "outEnd":
                                job = (lines, vb + (3 if positive else -3), y, "left" if positive else "right", "middle")
                            elif pos == "inEnd":
                                job = (lines, vb - (3 if positive else -3), y, "right" if positive else "left", "middle")
                            elif pos == "inBase":
                                job = (lines, va + (3 if positive else -3), y, "left" if positive else "right", "middle")
                            else:
                                job = (lines, (va + vb) / 2, y, "center", "middle")
                        label_jobs.append(job)

    def _polyline(self, pts: List[Tuple[float, float]], smooth: bool) -> None:
        c = self.c
        if len(pts) < 2:
            return
        p = c.beginPath()
        p.moveTo(*pts[0])
        if smooth and len(pts) > 2:
            for i in range(len(pts) - 1):
                p0 = pts[i - 1] if i > 0 else pts[i]
                p1, p2 = pts[i], pts[i + 1]
                p3 = pts[i + 2] if i + 2 < len(pts) else p2
                c1 = (p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6)
                c2 = (p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6)
                p.curveTo(c1[0], c1[1], c2[0], c2[1], p2[0], p2[1])
        else:
            for pt in pts[1:]:
                p.lineTo(*pt)
        c.drawPath(p, stroke=1, fill=0)

    def _lines_series(self, g: PlotGroup, ci: AxisInfo, vi: AxisInfo, plot, label_jobs) -> None:
        n = max(ci.n, 1)
        stacked = g.grouping in ("stacked", "percentStacked")
        cum = [0.0] * n
        totals = [sum(abs(_at(s.values, i) or 0.0) for s in g.series) or 1.0 for i in range(n)]
        blanks = self.chart.display_blanks
        for s in g.series:
            pts_groups: List[List[Tuple[float, float, int, float]]] = [[]]
            for i in range(n):
                v = _at(s.values, i)
                if v is None:
                    if blanks == "zero" or stacked:
                        v = 0.0
                    elif blanks == "span":
                        continue
                    else:
                        if pts_groups[-1]:
                            pts_groups.append([])
                        continue
                orig = v
                if g.grouping == "percentStacked":
                    v = v / totals[i]
                if stacked:
                    cum[i] += v
                    v = cum[i]
                a = self._cat_to(ci, i, plot)
                b = self._value_to(vi, v, plot)
                pt = (a, b) if ci.horizontal else (b, a)
                pts_groups[-1].append((pt[0], pt[1], i, orig))
            line = self.series_line(g, s)
            if line is not None:
                self._set_line(*line)
                for grp in pts_groups:
                    self._polyline([(x, y) for x, y, _, _ in grp], s.smooth)
                self.c.setDash([], 0)
            self._markers_for(g, s, [pt for grp in pts_groups for pt in grp])
            self._point_labels(g, s, [pt for grp in pts_groups for pt in grp], label_jobs, "r")

    def _markers_for(self, g: PlotGroup, s: Series, pts) -> None:
        symbol = self.marker_symbol(g, s)
        if not symbol:
            return
        color = self.series_color(g, s)
        mp = s.marker.props if s.marker else None
        fill = self._fill_of(mp, color) if mp else color
        line = (mp.line if mp and mp.line is not None else color) if not (mp and mp.line_none) else None
        size = s.marker.size if s.marker else 5.0
        for x, y, *_ in pts:
            self._marker(x, y, symbol, size, fill, line)

    def _point_labels(self, g: PlotGroup, s: Series, pts, label_jobs, default_pos: str) -> None:
        dl = self._labels_for(g, s)
        if dl is None:
            return
        spec = self._label_spec(dl)
        pos = dl.position or default_pos
        for x, y, i, v in pts:
            text = self._label_text(dl, s, i, v)
            if not text:
                continue
            lines = self._lines(text, spec)
            if pos == "t":
                label_jobs.append((lines, x, y + 4, "center", "bottom"))
            elif pos == "b":
                label_jobs.append((lines, x, y - 4, "center", "top"))
            elif pos == "l":
                label_jobs.append((lines, x - 5, y, "right", "middle"))
            elif pos == "ctr":
                label_jobs.append((lines, x, y, "center", "middle"))
            else:
                label_jobs.append((lines, x + 5, y, "left", "middle"))

    def _areas(self, g: PlotGroup, ci: AxisInfo, vi: AxisInfo, plot) -> None:
        c = self.c
        n = max(ci.n, 1)
        stacked = g.grouping in ("stacked", "percentStacked")
        totals = [sum(abs(_at(s.values, i) or 0.0) for s in g.series) or 1.0 for i in range(n)]
        lower = [0.0] * n
        base_v = self._base_value(vi)
        for s in g.series:
            upper = []
            for i in range(n):
                v = _at(s.values, i) or 0.0
                if g.grouping == "percentStacked":
                    v = v / totals[i]
                upper.append(lower[i] + v if stacked else v)
            low = lower if stacked else [base_v] * n
            color = self.series_color(g, s)
            fill = self._fill_of(s.props, color)
            line = self._line_of(s.props, None, 0.75)
            p = c.beginPath()
            pts_up = []
            pts_low = []
            for i in range(n):
                a = self._cat_to(ci, i, plot)
                pts_up.append((a, self._value_to(vi, upper[i], plot)) if ci.horizontal else (self._value_to(vi, upper[i], plot), a))
                pts_low.append((a, self._value_to(vi, low[i], plot)) if ci.horizontal else (self._value_to(vi, low[i], plot), a))
            if n == 1:
                # a single category spans the plot width
                px0, py0, px1, py1 = plot
                x_a, x_b = (px0, px1) if ci.horizontal else (py0, py1)
                pts_up = [(x_a, pts_up[0][1]), (x_b, pts_up[0][1])] if ci.horizontal else [(pts_up[0][0], x_a), (pts_up[0][0], x_b)]
                pts_low = [(x_a, pts_low[0][1]), (x_b, pts_low[0][1])] if ci.horizontal else [(pts_low[0][0], x_a), (pts_low[0][0], x_b)]
            p.moveTo(*pts_up[0])
            for pt in pts_up[1:]:
                p.lineTo(*pt)
            for pt in reversed(pts_low):
                p.lineTo(*pt)
            p.close()
            if fill is not None:
                c.setFillColorRGB(*fill)
            if line is not None:
                self._set_line(*line)
            c.drawPath(p, stroke=1 if line else 0, fill=1 if fill else 0)
            if stacked:
                lower = upper

    def _scatter(self, g: PlotGroup, xi: AxisInfo, yi: AxisInfo, plot, label_jobs) -> None:
        c = self.c
        px0, py0, px1, py1 = plot
        max_size = 0.0
        if g.type == "bubble":
            max_size = max((abs(v) for s in g.series for v in s.sizes if v is not None), default=1.0) or 1.0
        for s in g.series:
            xs = self._xvalues(s)
            pts = []
            for i, y in enumerate(s.values):
                if y is None or i >= len(xs):
                    continue
                fx = xi.scale.frac(xs[i]) if xi.scale else 0.0
                fy = yi.scale.frac(y) if yi.scale else 0.0
                pts.append((px0 + fx * (px1 - px0), py0 + fy * (py1 - py0), i, y))
            if g.type == "bubble":
                color = self.series_color(g, s)
                fill = self._fill_of(s.props, color)
                line = self._line_of(s.props, None, 0.75)
                dmax = min(px1 - px0, py1 - py0) * 0.25 * g.bubble_scale / 100.0
                for x, y, i, _ in pts:
                    size = _at(s.sizes, i) or 0.0
                    r = dmax / 2 * math.sqrt(abs(size) / max_size)
                    if fill is not None:
                        c.setFillColorRGB(*fill)
                    if line is not None:
                        self._set_line(*line)
                    c.circle(x, y, r, stroke=1 if line else 0, fill=1 if fill else 0)
                self._point_labels(g, s, pts, label_jobs, "ctr")
                continue
            line = self.series_line(g, s)
            if g.scatter_style == "marker" and not s.props.specified:
                line = None
            if line is not None:
                self._set_line(*line)
                self._polyline([(x, y) for x, y, _, _ in pts], s.smooth or g.scatter_style in ("smooth", "smoothMarker") and not s.props.specified)
                c.setDash([], 0)
            self._markers_for(g, s, pts)
            self._point_labels(g, s, pts, label_jobs, "r")

    # ------------------------------------------------------------------
    def _pie_chart(self, g: PlotGroup, x0, y0, x1, y1) -> None:
        c = self.c
        s = g.series[0]
        values = [max(0.0, v or 0.0) for v in s.values]
        total = sum(values)
        dl = self._labels_for(g, s)
        outside = dl is not None and (dl.position in ("outEnd", "bestFit", None)) and dl.any
        spec = self._label_spec(dl) if dl else None
        w, h = x1 - x0, y1 - y0
        r = min(w, h) / 2 * (0.78 if outside and dl and dl.position == "outEnd" else 0.92)
        max_expl = max([s.explosion] + list(s.point_explosion.values()) + [0.0]) / 100.0
        r = r / (1 + max_expl)
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        if total <= 0:
            return
        angle = g.first_slice_angle  # clockwise from 12 o'clock
        pcts = _pie_percents(values, total, dl.num_fmt if dl else None)
        hole = g.hole_size / 100.0 if g.type == "doughnut" else 0.0
        default_line = None
        jobs = []
        for i, v in enumerate(values):
            sweep = 360.0 * v / total
            pp = s.point_props.get(i, ShapeProps())
            color = self.color_for(i) if g.vary_colors else self.series_color(g, s)
            if not g.vary_colors and s.props.fill is not None:
                color = s.props.fill
            fill = self._fill_of(pp, color) if pp.specified else color
            line = self._line_of(pp if pp.specified else s.props, default_line, 0.75)
            mid = angle + sweep / 2
            expl = s.point_explosion.get(i, s.explosion) / 100.0 * r
            mx = cx + expl * math.sin(math.radians(mid))
            my = cy + expl * math.cos(math.radians(mid))
            if sweep > 0:
                start_ccw = 90.0 - angle
                p = c.beginPath()
                if hole > 0:
                    p.arc(mx - r, my - r, mx + r, my + r, start_ccw, -sweep)
                    p.arcTo(mx - r * hole, my - r * hole, mx + r * hole, my + r * hole, start_ccw - sweep, sweep)
                    p.close()
                elif sweep >= 359.999:
                    p.circle(mx, my, r)
                else:
                    p.moveTo(mx, my)
                    p.arcTo(mx - r, my - r, mx + r, my + r, start_ccw, -sweep)
                    p.close()
                if fill is not None:
                    c.setFillColorRGB(*fill)
                if line is not None:
                    self._set_line(*line)
                c.drawPath(p, stroke=1 if line else 0, fill=1 if fill else 0)
            if dl is not None and v > 0:
                text = self._label_text(dl, s, i, s.values[i], pcts[i])
                if text:
                    lines = self._lines(text, spec)
                    lw, lh = self._block_size(lines)
                    pos = dl.position or "bestFit"
                    inside_fits = sweep > 25 and lw < r * 0.6
                    if pos in ("ctr", "inEnd") or (pos == "bestFit" and inside_fits) or hole > 0:
                        rr = r * ((1 + hole) / 2 if hole > 0 else (0.62 if pos != "inEnd" else 0.8))
                        jobs.append((lines, mx + rr * math.sin(math.radians(mid)), my + rr * math.cos(math.radians(mid)), "center", "middle"))
                    else:
                        rr = r * 1.08
                        lx = mx + rr * math.sin(math.radians(mid))
                        ly = my + rr * math.cos(math.radians(mid))
                        align = "left" if math.sin(math.radians(mid)) >= 0 else "right"
                        jobs.append((lines, lx, ly, align, "middle"))
            angle += sweep
        for job in jobs:
            self._draw_text_block(*job)

    def _radar_chart(self, groups: List[PlotGroup], x0, y0, x1, y1) -> None:
        c = self.c
        g0 = groups[0]
        series = [s for g in groups for s in g.series]
        n = max(max(len(s.values), len(s.categories)) for s in series)
        if n < 3:
            n = max(n, 3)
        labels = next((s.categories for s in series if any(s.categories)), [str(i + 1) for i in range(n)])
        cat_ax = self.chart.axes.get(g0.ax_ids[0]) if g0.ax_ids else None
        val_ax = self.chart.axes.get(g0.ax_ids[1]) if len(g0.ax_ids) > 1 else None
        val_ax = val_ax or Axis(id="v", kind="val")
        spec = self.spec((cat_ax.style if cat_ax else TextStyle()).merged(self.base), self.base.size or 10.0)
        lab_h = self.text.line_height(spec)
        lab_w = max((self.text.measure([(t, spec)]) for t in labels), default=0.0)
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        r = max(10.0, min((x1 - x0) / 2 - lab_w - 6, (y1 - y0) / 2 - lab_h - 4))
        vals = [v for s in series for v in s.values if v is not None]
        lo, hi = (min(vals + [0.0]), max(vals + [0.0])) if vals else (0.0, 1.0)
        vspec = self.spec(val_ax.style.merged(self.base), self.base.size or 10.0)
        scale = self._scale(lo, hi, val_ax, r, False, vspec, "General")

        def point(i: int, v: float) -> Tuple[float, float]:
            ang = math.radians(90 - 360.0 * i / n)
            rr = r * scale.frac(v)
            return cx + rr * math.cos(ang), cy + rr * math.sin(ang)

        grid = self._line_of(val_ax.major_grid, GRID_GRAY, 0.75) if val_ax.major_grid else (GRID_GRAY, 0.75, None)
        if grid:
            self._set_line(*grid)
            for t in scale.ticks():
                pts = [point(i, t) for i in range(n)]
                p = c.beginPath()
                p.moveTo(*pts[0])
                for pt in pts[1:]:
                    p.lineTo(*pt)
                p.close()
                c.drawPath(p, stroke=1, fill=0)
            for i in range(n):
                c.line(cx, cy, *point(i, scale.hi))
        for i, lab in enumerate(labels[:n]):
            ang = math.radians(90 - 360.0 * i / n)
            lx, ly = cx + (r + 6) * math.cos(ang), cy + (r + 6) * math.sin(ang)
            align = "center" if abs(math.cos(ang)) < 0.2 else ("left" if math.cos(ang) > 0 else "right")
            valign = "bottom" if math.sin(ang) > 0.5 else ("top" if math.sin(ang) < -0.5 else "middle")
            self._draw_text_block(self._lines(lab, spec), lx, ly, align, valign)
        if not val_ax.delete:
            for t in scale.ticks():
                x, y = point(0, t)
                self._draw_text_block(self._lines(format_value(t, val_ax.num_fmt or "General").text, vspec), x - 3, y, "right", "middle")
        for g in groups:
            for s in g.series:
                pts = [point(i, _at(s.values, i) or 0.0) for i in range(n)]
                color = self.series_color(g, s)
                p = c.beginPath()
                p.moveTo(*pts[0])
                for pt in pts[1:]:
                    p.lineTo(*pt)
                p.close()
                if g.radar_style == "filled":
                    c.setFillColorRGB(*self._fill_of(s.props, color) or color)
                    c.drawPath(p, stroke=0, fill=1)
                else:
                    line = self.series_line(g, s)
                    if line:
                        self._set_line(*line)
                        c.drawPath(p, stroke=1, fill=0)
                        c.setDash([], 0)
                    self._markers_for(g, s, [(x, y) for x, y in pts])


def _pie_percents(values: List[float], total: float, fmt: Optional[str]) -> List[float]:
    """Slice shares, rounded like Excel's labels: the last visible slice takes the remainder
    so the displayed percentages add up to 100%."""
    if total <= 0:
        return [0.0 for _ in values]
    m = re.search(r"\.(0+)%", fmt or "")
    decimals = len(m.group(1)) if m else 0
    unit = 10 ** -(decimals + 2)
    shares = [v / total for v in values]
    rounded = [round(round(x / unit) * unit, decimals + 4) for x in shares]
    last = max((i for i, v in enumerate(values) if v > 0), default=None)
    if last is not None:
        rounded[last] = round(1.0 - sum(r for i, r in enumerate(rounded) if i != last), decimals + 4)
    return rounded


def _at(values: Sequence[Optional[float]], i: int) -> Optional[float]:
    return values[i] if 0 <= i < len(values) else None


def _side(pos: str, horizontal: bool) -> str:
    if horizontal:
        return pos if pos in ("b", "t") else "b"
    return pos if pos in ("l", "r") else "l"
