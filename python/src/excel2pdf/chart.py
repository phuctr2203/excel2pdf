"""Parse DrawingML chart parts (xl/charts/chartN.xml) into a plain model."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Sequence, Tuple

from .colors import RGB, ColorResolver
from .drawing import A, C, EMU_PER_PT
from .numfmt import format_value

# resolve(ref) -> list of (value, number_format) or None when unresolvable
RefResolver = Callable[[str], Optional[List[Tuple[object, Optional[str]]]]]


@dataclass
class TextStyle:
    size: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    color: Optional[RGB] = None
    font: Optional[str] = None
    rotation: Optional[float] = None  # degrees, from bodyPr rot

    def merged(self, base: "TextStyle") -> "TextStyle":
        return TextStyle(
            self.size if self.size is not None else base.size,
            self.bold if self.bold is not None else base.bold,
            self.italic if self.italic is not None else base.italic,
            self.color if self.color is not None else base.color,
            self.font if self.font is not None else base.font,
            self.rotation if self.rotation is not None else base.rotation,
        )


@dataclass
class ShapeProps:
    fill: Optional[RGB] = None
    fill_none: bool = False
    line: Optional[RGB] = None
    line_none: bool = False
    line_width: Optional[float] = None
    dash: Optional[str] = None
    specified: bool = False


@dataclass
class Marker:
    symbol: Optional[str] = None  # none, circle, square, diamond, triangle, x, star, dot, dash, plus, auto
    size: float = 5.0
    props: ShapeProps = field(default_factory=ShapeProps)


@dataclass
class DataLabels:
    show_val: bool = False
    show_pct: bool = False
    show_cat: bool = False
    show_ser: bool = False
    position: Optional[str] = None
    num_fmt: Optional[str] = None
    style: TextStyle = field(default_factory=TextStyle)
    separator: Optional[str] = None

    @property
    def any(self) -> bool:
        return self.show_val or self.show_pct or self.show_cat or self.show_ser


@dataclass
class Series:
    idx: int
    order: int
    name: str = ""
    categories: List[str] = field(default_factory=list)
    cat_values: List[Optional[float]] = field(default_factory=list)  # numeric categories / scatter x
    values: List[Optional[float]] = field(default_factory=list)
    sizes: List[Optional[float]] = field(default_factory=list)
    format_code: Optional[str] = None
    props: ShapeProps = field(default_factory=ShapeProps)
    marker: Optional[Marker] = None
    point_props: Dict[int, ShapeProps] = field(default_factory=dict)
    point_explosion: Dict[int, float] = field(default_factory=dict)
    labels: Optional[DataLabels] = None
    smooth: bool = False
    explosion: float = 0.0
    invert_if_negative: bool = False


@dataclass
class Axis:
    id: str
    kind: str  # cat | val | date | ser
    delete: bool = False
    pos: str = "b"
    cross_ax: Optional[str] = None
    min: Optional[float] = None
    max: Optional[float] = None
    reverse: bool = False
    log_base: Optional[float] = None
    major_grid: Optional[ShapeProps] = None
    minor_grid: Optional[ShapeProps] = None
    num_fmt: Optional[str] = None
    source_linked: bool = True
    major_tick: str = "out"
    minor_tick: str = "none"
    label_pos: str = "nextTo"
    title: Optional["Title"] = None
    major_unit: Optional[float] = None
    minor_unit: Optional[float] = None
    crosses: str = "autoZero"
    crosses_at: Optional[float] = None
    cross_between: Optional[str] = None
    style: TextStyle = field(default_factory=TextStyle)
    props: ShapeProps = field(default_factory=ShapeProps)
    tick_skip: Optional[int] = None


@dataclass
class PlotGroup:
    type: str  # bar line area pie doughnut scatter radar bubble
    series: List[Series] = field(default_factory=list)
    bar_dir: str = "col"
    grouping: str = "clustered"
    vary_colors: bool = False
    gap_width: float = 150.0
    overlap: Optional[float] = None
    hole_size: float = 50.0
    first_slice_angle: float = 0.0
    scatter_style: str = "lineMarker"
    radar_style: str = "marker"
    ax_ids: List[str] = field(default_factory=list)
    show_markers: bool = True
    labels: Optional[DataLabels] = None
    bubble_scale: float = 100.0
    is_3d: bool = False


@dataclass
class Title:
    runs: List[Tuple[str, TextStyle]] = field(default_factory=list)
    style: TextStyle = field(default_factory=TextStyle)
    overlay: bool = False
    auto: bool = False  # no text given: Excel derives it


@dataclass
class Legend:
    pos: str = "r"
    overlay: bool = False
    deleted: set = field(default_factory=set)
    style: TextStyle = field(default_factory=TextStyle)


@dataclass
class Chart:
    groups: List[PlotGroup] = field(default_factory=list)
    axes: Dict[str, Axis] = field(default_factory=dict)
    title: Optional[Title] = None
    legend: Optional[Legend] = None
    plot_visible_only: bool = True
    display_blanks: str = "gap"
    area_props: ShapeProps = field(default_factory=ShapeProps)
    plot_props: ShapeProps = field(default_factory=ShapeProps)
    rounded: bool = True
    plot_layout: Optional[Tuple[float, float, float, float]] = None  # x, y, w, h fractions (inner)
    text_style: TextStyle = field(default_factory=TextStyle)
    has_text_props: bool = False

    @property
    def all_series(self) -> List[Series]:
        return [s for g in self.groups for s in g.series]


_GROUP_TAGS = {
    "barChart": "bar", "bar3DChart": "bar", "lineChart": "line", "line3DChart": "line",
    "areaChart": "area", "area3DChart": "area", "pieChart": "pie", "pie3DChart": "pie",
    "ofPieChart": "pie", "doughnutChart": "doughnut", "scatterChart": "scatter",
    "radarChart": "radar", "bubbleChart": "bubble", "stockChart": "line", "surfaceChart": "area",
    "surface3DChart": "area",
}


def _val(el: Optional[ET.Element], tag: str, default=None):
    if el is None:
        return default
    child = el.find(C + tag)
    if child is None:
        return default
    return child.get("val", default if default is not None else "1")


def _bool(el: Optional[ET.Element], tag: str, default: bool) -> bool:
    v = _val(el, tag, None)
    if v is None:
        return default
    return v in ("1", "true")


def _float(el: Optional[ET.Element], tag: str, default: Optional[float]) -> Optional[float]:
    v = _val(el, tag, None)
    try:
        return float(v) if v is not None else default
    except ValueError:
        return default


class ChartReader:
    def __init__(self, colors: ColorResolver, resolve: RefResolver, date1904: bool = False) -> None:
        self.colors = colors
        self.theme = colors.theme
        self.resolve = resolve
        self.date1904 = date1904

    # ------------------------------------------------------------------
    def read(self, root: ET.Element) -> Optional[Chart]:
        chart_el = root.find(C + "chart")
        if chart_el is None:
            return None
        ch = Chart()
        ch.rounded = _bool(root, "roundedCorners", True)
        ch.area_props = self.sppr(root.find(C + "spPr"))
        txpr = root.find(C + "txPr")
        ch.text_style = self.text_props(txpr)
        ch.has_text_props = txpr is not None

        plot = chart_el.find(C + "plotArea")
        if plot is None:
            return ch
        ch.plot_props = self.sppr(plot.find(C + "spPr"))
        ch.plot_layout = self._layout(plot.find(C + "layout"))
        ch.plot_visible_only = _bool(chart_el, "plotVisOnly", True)
        ch.display_blanks = _val(chart_el, "dispBlanksAs", "gap") or "gap"

        for child in plot:
            tag = child.tag.replace(C, "")
            if tag in _GROUP_TAGS:
                ch.groups.append(self._group(child, _GROUP_TAGS[tag], tag))
            elif tag in ("catAx", "valAx", "dateAx", "serAx"):
                ax = self._axis(child, tag[:-2])
                ch.axes[ax.id] = ax

        title_el = chart_el.find(C + "title")
        auto_deleted = _bool(chart_el, "autoTitleDeleted", False)
        if title_el is not None:
            ch.title = self._title(title_el)
        elif not auto_deleted and len(ch.all_series) == 1 and _val(chart_el, "autoTitleDeleted", None) == "0":
            ch.title = Title(auto=True)
        if ch.title is not None and ch.title.auto:
            series = ch.all_series
            text = series[0].name if len(series) == 1 and series[0].name else "Chart Title"
            ch.title.runs = [(text, TextStyle())]

        leg = chart_el.find(C + "legend")
        if leg is not None:
            lg = Legend(pos=_val(leg, "legendPos", "r") or "r", overlay=_bool(leg, "overlay", False))
            for entry in leg.findall(C + "legendEntry"):
                if _bool(entry, "delete", False):
                    lg.deleted.add(int(_val(entry, "idx", "0")))
            lg.style = self.text_props(leg.find(C + "txPr"))
            ch.legend = lg
        return ch

    # ------------------------------------------------------------------
    def _layout(self, el: Optional[ET.Element]) -> Optional[Tuple[float, float, float, float]]:
        if el is None:
            return None
        ml = el.find(C + "manualLayout")
        if ml is None:
            return None
        if (_val(ml, "xMode", "factor") != "edge") or (_val(ml, "yMode", "factor") != "edge"):
            return None
        try:
            x = float(_val(ml, "x", "0"))
            y = float(_val(ml, "y", "0"))
            w = float(_val(ml, "w", "0"))
            h = float(_val(ml, "h", "0"))
        except ValueError:
            return None
        if w <= 0 or h <= 0:
            return None
        return (x, y, w, h)

    def _group(self, el: ET.Element, kind: str, tag: str) -> PlotGroup:
        g = PlotGroup(type=kind, is_3d="3D" in tag)
        g.bar_dir = _val(el, "barDir", "col") or "col"
        g.grouping = _val(el, "grouping", "clustered" if kind == "bar" else "standard") or "standard"
        g.vary_colors = _bool(el, "varyColors", kind in ("pie", "doughnut"))
        g.gap_width = _float(el, "gapWidth", 150.0) or 0.0
        g.overlap = _float(el, "overlap", None)
        g.hole_size = _float(el, "holeSize", 50.0) or 50.0
        g.first_slice_angle = _float(el, "firstSliceAng", 0.0) or 0.0
        g.scatter_style = _val(el, "scatterStyle", "lineMarker") or "lineMarker"
        g.radar_style = _val(el, "radarStyle", "marker") or "marker"
        g.show_markers = _bool(el, "marker", True)
        g.bubble_scale = _float(el, "bubbleScale", 100.0) or 100.0
        g.ax_ids = [a.get("val", "") for a in el.findall(C + "axId")]
        dl = el.find(C + "dLbls")
        if dl is not None:
            g.labels = self._labels(dl)
        for s in el.findall(C + "ser"):
            g.series.append(self._series(s, kind))
        g.series.sort(key=lambda s: s.order)
        return g

    def _series(self, el: ET.Element, kind: str) -> Series:
        s = Series(idx=int(_val(el, "idx", "0")), order=int(_val(el, "order", "0")))
        tx = el.find(C + "tx")
        if tx is not None:
            sref = tx.find(C + "strRef")
            if sref is not None:
                vals = self._ref_values(sref)
                s.name = " ".join(str(v) for v, _ in vals if v is not None) if vals else ""
            else:
                v = tx.find(C + "v")
                s.name = v.text if v is not None and v.text else ""
        if not s.name:
            s.name = f"Series{s.idx + 1}"
        s.props = self.sppr(el.find(C + "spPr"))
        m = el.find(C + "marker")
        if m is not None:
            s.marker = Marker(symbol=_val(m, "symbol", None), size=_float(m, "size", 5.0) or 5.0,
                              props=self.sppr(m.find(C + "spPr")))
        for dpt in el.findall(C + "dPt"):
            idx = int(_val(dpt, "idx", "0"))
            s.point_props[idx] = self.sppr(dpt.find(C + "spPr"))
            exp = _float(dpt, "explosion", None)
            if exp is not None:
                s.point_explosion[idx] = exp
        dl = el.find(C + "dLbls")
        if dl is not None:
            s.labels = self._labels(dl)
        s.smooth = _bool(el, "smooth", False)
        s.explosion = _float(el, "explosion", 0.0) or 0.0
        s.invert_if_negative = _bool(el, "invertIfNegative", False)

        cat_tag = "xVal" if kind in ("scatter", "bubble") else "cat"
        val_tag = "yVal" if kind in ("scatter", "bubble") else "val"
        cat = el.find(C + cat_tag)
        if cat is not None:
            s.categories, s.cat_values = self._categories(cat)
        val = el.find(C + val_tag)
        if val is not None:
            s.values, s.format_code = self._numbers(val)
        size = el.find(C + "bubbleSize")
        if size is not None:
            s.sizes, _ = self._numbers(size)
        return s

    def _labels(self, el: ET.Element) -> DataLabels:
        dl = DataLabels(
            show_val=_bool(el, "showVal", False),
            show_pct=_bool(el, "showPercent", False),
            show_cat=_bool(el, "showCatName", False),
            show_ser=_bool(el, "showSerName", False),
            position=_val(el, "dLblPos", None),
        )
        if _bool(el, "delete", False):
            dl.show_val = dl.show_pct = dl.show_cat = dl.show_ser = False
        nf = el.find(C + "numFmt")
        if nf is not None:
            dl.num_fmt = nf.get("formatCode")
        sep = el.find(C + "separator")
        if sep is not None and sep.text:
            dl.separator = sep.text
        dl.style = self.text_props(el.find(C + "txPr"))
        return dl

    def _axis(self, el: ET.Element, kind: str) -> Axis:
        ax = Axis(id=_val(el, "axId", "") or "", kind=kind)
        ax.delete = _bool(el, "delete", False)
        ax.pos = _val(el, "axPos", "b") or "b"
        ax.cross_ax = _val(el, "crossAx", None)
        scaling = el.find(C + "scaling")
        if scaling is not None:
            ax.min = _float(scaling, "min", None)
            ax.max = _float(scaling, "max", None)
            ax.reverse = _val(scaling, "orientation", "minMax") == "maxMin"
            ax.log_base = _float(scaling, "logBase", None)
        mg = el.find(C + "majorGridlines")
        if mg is not None:
            ax.major_grid = self.sppr(mg.find(C + "spPr"))
        mng = el.find(C + "minorGridlines")
        if mng is not None:
            ax.minor_grid = self.sppr(mng.find(C + "spPr"))
        nf = el.find(C + "numFmt")
        if nf is not None:
            ax.num_fmt = nf.get("formatCode")
            ax.source_linked = nf.get("sourceLinked", "0") in ("1", "true")
        ax.major_tick = _val(el, "majorTickMark", "out") or "out"
        ax.minor_tick = _val(el, "minorTickMark", "none") or "none"
        ax.label_pos = _val(el, "tickLblPos", "nextTo") or "nextTo"
        t = el.find(C + "title")
        if t is not None:
            ax.title = self._title(t)
            if ax.title.auto:
                ax.title.runs = [("Axis Title", TextStyle())]
        ax.major_unit = _float(el, "majorUnit", None)
        ax.minor_unit = _float(el, "minorUnit", None)
        ax.crosses = _val(el, "crosses", "autoZero") or "autoZero"
        ax.crosses_at = _float(el, "crossesAt", None)
        ax.cross_between = _val(el, "crossBetween", None)
        ax.style = self.text_props(el.find(C + "txPr"))
        ax.props = self.sppr(el.find(C + "spPr"))
        skip = _float(el, "tickLblSkip", None)
        ax.tick_skip = int(skip) if skip else None
        return ax

    def _title(self, el: ET.Element) -> Title:
        t = Title(overlay=_bool(el, "overlay", False))
        txpr = el.find(C + "txPr")
        tx = el.find(C + "tx")
        if tx is not None:
            rich = tx.find(C + "rich")
            if rich is not None:
                base = self.text_props(rich)
                t.style = base
                for p_i, p in enumerate(rich.findall(A + "p")):
                    if p_i > 0:
                        t.runs.append(("\n", base))
                    for r in p:
                        if r.tag in (A + "r", A + "fld"):
                            te = r.find(A + "t")
                            st = self._run_style(r.find(A + "rPr")).merged(base)
                            t.runs.append((te.text if te is not None and te.text else "", st))
            else:
                sref = tx.find(C + "strRef")
                if sref is not None:
                    vals = self._ref_values(sref)
                    text = " ".join(str(v) for v, _ in vals if v is not None) if vals else ""
                    t.runs = [(text, TextStyle())]
        elif txpr is not None:
            t.style = self.text_props(txpr)
        if txpr is not None and tx is None:
            t.style = self.text_props(txpr)
        if not t.runs:
            t.auto = True
        return t

    # ------------------------------------------------------------------
    def text_props(self, el: Optional[ET.Element]) -> TextStyle:
        """Read <c:txPr>/<c:rich>: bodyPr rotation and the first paragraph's defRPr."""
        st = TextStyle()
        if el is None:
            return st
        body = el.find(A + "bodyPr")
        if body is not None and body.get("rot") not in (None, "-60000000"):
            try:
                st.rotation = int(body.get("rot", "0")) / 60000.0
            except ValueError:
                pass
        if body is not None and body.get("vert") in ("vert", "eaVert"):
            st.rotation = 90.0
        d = el.find(f"{A}p/{A}pPr/{A}defRPr")
        if d is not None:
            st = self._run_style(d).merged(st)
            st.rotation = st.rotation
        return st

    def _run_style(self, rpr: Optional[ET.Element]) -> TextStyle:
        st = TextStyle()
        if rpr is None:
            return st
        if rpr.get("sz"):
            st.size = int(rpr.get("sz")) / 100.0
        if rpr.get("b") is not None:
            st.bold = rpr.get("b") in ("1", "true")
        if rpr.get("i") is not None:
            st.italic = rpr.get("i") in ("1", "true")
        sf = rpr.find(A + "solidFill")
        if sf is not None:
            st.color = self.colors.dml(sf)
        latin = rpr.find(A + "latin")
        if latin is not None:
            st.font = self.theme.font_for(latin.get("typeface"))
        return st

    def sppr(self, el: Optional[ET.Element]) -> ShapeProps:
        sp = ShapeProps()
        if el is None:
            return sp
        sp.specified = True
        for child in el:
            t = child.tag.replace(A, "")
            if t == "noFill":
                sp.fill_none = True
            elif t == "solidFill":
                sp.fill = self.colors.dml(child)
            elif t == "gradFill":
                stops = child.findall(f"{A}gsLst/{A}gs")
                if stops:
                    sp.fill = self.colors.dml(stops[len(stops) // 2])
            elif t == "pattFill":
                sp.fill = self.colors.dml(child.find(A + "fgClr"))
            elif t == "ln":
                if child.find(A + "noFill") is not None:
                    sp.line_none = True
                sf = child.find(A + "solidFill")
                if sf is not None:
                    sp.line = self.colors.dml(sf)
                gf = child.find(A + "gradFill")
                if gf is not None:
                    stops = gf.findall(f"{A}gsLst/{A}gs")
                    if stops:
                        sp.line = self.colors.dml(stops[0])
                if child.get("w"):
                    sp.line_width = int(child.get("w")) / EMU_PER_PT
                dash = child.find(A + "prstDash")
                if dash is not None:
                    sp.dash = dash.get("val")
        return sp

    # ------------------------------------------------------------------
    def _ref_values(self, ref_el: ET.Element) -> List[Tuple[object, Optional[str]]]:
        """Values for a strRef/numRef: live cells first, the saved cache as fallback."""
        f = ref_el.find(C + "f")
        if f is not None and f.text:
            vals = self.resolve(f.text.strip())
            if vals is not None:
                return vals
        return self._cache(ref_el)

    def _cache(self, ref_el: ET.Element) -> List[Tuple[object, Optional[str]]]:
        cache = None
        for tag in ("numCache", "strCache", "multiLvlStrCache"):
            cache = ref_el.find(C + tag)
            if cache is not None:
                break
        if cache is None:
            return []
        if cache.tag == C + "multiLvlStrCache":
            lvl = cache.find(C + "lvl")  # innermost level
            if lvl is None:
                return []
            cache_pts = lvl
        else:
            cache_pts = cache
        fmt_el = cache.find(C + "formatCode")
        fmt = fmt_el.text if fmt_el is not None and fmt_el.text else None
        count = int(_val(cache, "ptCount", "0") or 0)
        out: List[Tuple[object, Optional[str]]] = [(None, fmt)] * count
        for pt in cache_pts.findall(C + "pt"):
            idx = int(pt.get("idx", "0"))
            v = pt.find(C + "v")
            text = v.text if v is not None else None
            value: object = text
            if cache.tag == C + "numCache" and text is not None:
                try:
                    value = float(text)
                except ValueError:
                    value = None
            pt_fmt = pt.get("formatCode") or fmt
            while idx >= len(out):
                out.append((None, fmt))
            out[idx] = (value, pt_fmt)
        return out

    def _literal(self, lit_el: ET.Element, numeric: bool) -> List[Tuple[object, Optional[str]]]:
        return self._cache_like(lit_el, numeric)

    def _cache_like(self, el: ET.Element, numeric: bool) -> List[Tuple[object, Optional[str]]]:
        fmt_el = el.find(C + "formatCode")
        fmt = fmt_el.text if fmt_el is not None else None
        out: List[Tuple[object, Optional[str]]] = []
        for pt in el.findall(C + "pt"):
            idx = int(pt.get("idx", "0"))
            v = pt.find(C + "v")
            value: object = v.text if v is not None else None
            if numeric and value is not None:
                try:
                    value = float(value)  # type: ignore[arg-type]
                except ValueError:
                    value = None
            while idx >= len(out):
                out.append((None, fmt))
            out[idx] = (value, fmt)
        return out

    def _source(self, el: ET.Element) -> List[Tuple[object, Optional[str]]]:
        for tag in ("numRef", "strRef", "multiLvlStrRef"):
            ref = el.find(C + tag)
            if ref is not None:
                return self._ref_values(ref)
        nl = el.find(C + "numLit")
        if nl is not None:
            return self._cache_like(nl, True)
        sl = el.find(C + "strLit")
        if sl is not None:
            return self._cache_like(sl, False)
        return []

    def _numbers(self, el: ET.Element) -> Tuple[List[Optional[float]], Optional[str]]:
        vals = self._source(el)
        out: List[Optional[float]] = []
        fmt = None
        for v, f in vals:
            if fmt is None and f:
                fmt = f
            out.append(_to_float(v))
        return out, fmt

    def _categories(self, el: ET.Element) -> Tuple[List[str], List[Optional[float]]]:
        vals = self._source(el)
        labels: List[str] = []
        nums: List[Optional[float]] = []
        for v, f in vals:
            nums.append(_to_float(v))
            if v is None:
                labels.append("")
            elif isinstance(v, str):
                labels.append(v)
            else:
                labels.append(format_value(v, f or "General", self.date1904).text)
        return labels, nums


def _to_float(v) -> Optional[float]:
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    from .numfmt import to_serial

    s = to_serial(v)
    if s is not None:
        return s
    try:
        return float(str(v))
    except ValueError:
        return None
