"""Parse SpreadsheetML drawings (xl/drawings/drawingN.xml): pictures, charts, shapes, groups."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

from .colors import RGB, ColorResolver
from .package import NS_R, Package
from .theme import NS_A

NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

X = "{%s}" % NS_XDR
A = "{%s}" % NS_A
R = "{%s}" % NS_R
C = "{%s}" % NS_C
MC = "{%s}" % NS_MC

EMU_PER_PT = 12700.0

# mc:Choice blocks we can read as plain DrawingML (their extras are ignored).
_UNDERSTOOD = {"a14", "a15", "x14", "x15", "xdr14", "v", "o", "x", "w14", "r", "c14", "a"}


@dataclass
class Marker:
    col: int  # 0-based
    col_off: float  # points
    row: int  # 0-based
    row_off: float  # points


@dataclass
class Anchor:
    kind: str  # two | one | abs
    frm: Optional[Marker] = None
    to: Optional[Marker] = None
    ext: Tuple[float, float] = (0.0, 0.0)  # points
    pos: Tuple[float, float] = (0.0, 0.0)  # points


@dataclass
class Xfrm:
    x: float = 0.0
    y: float = 0.0
    w: float = 0.0
    h: float = 0.0
    rot: float = 0.0  # degrees clockwise
    flip_h: bool = False
    flip_v: bool = False
    ch_x: float = 0.0
    ch_y: float = 0.0
    ch_w: float = 0.0
    ch_h: float = 0.0


@dataclass
class LineStyle:
    color: RGB
    width: float  # points
    dash: Optional[str] = None
    head: Optional[str] = None  # arrow type at start
    tail: Optional[str] = None  # arrow type at end
    alpha: float = 1.0


@dataclass
class TextRun:
    text: str
    size: float
    bold: bool
    italic: bool
    underline: bool
    strike: bool
    color: RGB
    font: Optional[str]


@dataclass
class Paragraph:
    runs: List[TextRun]
    align: str = "l"
    end_size: float = 11.0


@dataclass
class TextBody:
    paragraphs: List[Paragraph]
    anchor: str = "t"  # t | ctr | b
    insets: Tuple[float, float, float, float] = (7.2, 3.6, 7.2, 3.6)  # l t r b
    wrap: bool = True
    vertical: Optional[str] = None


@dataclass
class DrawingObject:
    kind: str  # picture | chart | shape | connector | group
    xfrm: Xfrm
    anchor: Optional[Anchor] = None
    name: str = ""
    hidden: bool = False
    # picture
    image: Optional[bytes] = None
    crop: Tuple[float, float, float, float] = (0.0, 0.0, 0.0, 0.0)
    # chart
    chart_part: Optional[str] = None
    # shape / connector
    geometry: str = "rect"
    adj: dict = field(default_factory=dict)
    fill: Optional[RGB] = None
    fill_alpha: float = 1.0
    line: Optional[LineStyle] = None
    text: Optional[TextBody] = None
    # group
    children: List["DrawingObject"] = field(default_factory=list)


def _emu(v: Optional[str]) -> float:
    try:
        return int(v or 0) / EMU_PER_PT
    except ValueError:
        return 0.0


def _marker(el: Optional[ET.Element]) -> Optional[Marker]:
    if el is None:
        return None

    def num(tag: str) -> str:
        child = el.find(X + tag)
        return child.text.strip() if child is not None and child.text else "0"

    return Marker(int(num("col")), _emu(num("colOff")), int(num("row")), _emu(num("rowOff")))


def _resolve_alternate(el: ET.Element) -> List[ET.Element]:
    """Pick the children of mc:AlternateContent we can understand."""
    for choice in el.findall(MC + "Choice"):
        req = set((choice.get("Requires") or "").split())
        if req <= _UNDERSTOOD:
            return list(choice)
    fb = el.find(MC + "Fallback")
    return list(fb) if fb is not None else []


def _expand(children: List[ET.Element]) -> List[ET.Element]:
    out: List[ET.Element] = []
    for ch in children:
        if ch.tag == MC + "AlternateContent":
            out.extend(_expand(_resolve_alternate(ch)))
        else:
            out.append(ch)
    return out


class DrawingReader:
    def __init__(self, pkg: Package, colors: ColorResolver) -> None:
        self.pkg = pkg
        self.colors = colors
        self.theme = colors.theme

    def read(self, part: str) -> List[DrawingObject]:
        root = self.pkg.xml(part)
        if root is None:
            return []
        self._part = part
        self._rels = self.pkg.rels(part)
        out: List[DrawingObject] = []
        for anc in root:
            tag = anc.tag.replace(X, "")
            if tag not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
                continue
            anchor = self._anchor(anc, tag)
            for el in _expand(list(anc)):
                obj = self._object(el)
                if obj is not None:
                    obj.anchor = anchor
                    out.append(obj)
                    break
        return out

    def _anchor(self, anc: ET.Element, tag: str) -> Anchor:
        if tag == "twoCellAnchor":
            return Anchor("two", _marker(anc.find(X + "from")), _marker(anc.find(X + "to")))
        ext_el = anc.find(X + "ext")
        ext = (_emu(ext_el.get("cx")), _emu(ext_el.get("cy"))) if ext_el is not None else (0.0, 0.0)
        if tag == "oneCellAnchor":
            return Anchor("one", _marker(anc.find(X + "from")), ext=ext)
        pos_el = anc.find(X + "pos")
        pos = (_emu(pos_el.get("x")), _emu(pos_el.get("y"))) if pos_el is not None else (0.0, 0.0)
        return Anchor("abs", ext=ext, pos=pos)

    # ------------------------------------------------------------------
    def _object(self, el: ET.Element) -> Optional[DrawingObject]:
        tag = el.tag.replace(X, "")
        if tag == "pic":
            return self._picture(el)
        if tag == "graphicFrame":
            return self._graphic_frame(el)
        if tag in ("sp", "cxnSp"):
            return self._shape(el, connector=(tag == "cxnSp"))
        if tag == "grpSp":
            return self._group(el)
        return None

    def _nv(self, el: ET.Element) -> Tuple[str, bool]:
        for nv in el:
            cnv = nv.find(X + "cNvPr")
            if cnv is not None:
                return cnv.get("name", ""), cnv.get("hidden") in ("1", "true")
        return "", False

    def _xfrm(self, sppr: Optional[ET.Element], tag: str = A + "xfrm") -> Xfrm:
        xf = Xfrm()
        if sppr is None:
            return xf
        el = sppr.find(tag)
        if el is None:
            return xf
        xf.rot = int(el.get("rot", "0")) / 60000.0
        xf.flip_h = el.get("flipH") in ("1", "true")
        xf.flip_v = el.get("flipV") in ("1", "true")
        off = el.find(A + "off")
        ext = el.find(A + "ext")
        if off is not None:
            xf.x, xf.y = _emu(off.get("x")), _emu(off.get("y"))
        if ext is not None:
            xf.w, xf.h = _emu(ext.get("cx")), _emu(ext.get("cy"))
        choff = el.find(A + "chOff")
        chext = el.find(A + "chExt")
        if choff is not None:
            xf.ch_x, xf.ch_y = _emu(choff.get("x")), _emu(choff.get("y"))
        if chext is not None:
            xf.ch_w, xf.ch_h = _emu(chext.get("cx")), _emu(chext.get("cy"))
        return xf

    def _picture(self, el: ET.Element) -> Optional[DrawingObject]:
        name, hidden = self._nv(el)
        blip_fill = el.find(X + "blipFill")
        if blip_fill is None:
            return None
        blip = blip_fill.find(A + "blip")
        if blip is None:
            return None
        rid = blip.get(R + "embed")
        rel = self._rels.get(rid or "")
        if rel is None or rel.external:
            return None
        data = self.pkg.read(rel.target)
        if not data:
            return None
        crop = (0.0, 0.0, 0.0, 0.0)
        src = blip_fill.find(A + "srcRect")
        if src is not None:
            crop = tuple(int(src.get(k, "0")) / 100000.0 for k in ("l", "t", "r", "b"))  # type: ignore[assignment]
        sppr = el.find(X + "spPr")
        obj = DrawingObject("picture", self._xfrm(sppr), name=name, hidden=hidden, image=data, crop=crop)
        obj.line = self._line(sppr.find(A + "ln") if sppr is not None else None, None)
        return obj

    def _graphic_frame(self, el: ET.Element) -> Optional[DrawingObject]:
        name, hidden = self._nv(el)
        gd = el.find(f"{A}graphic/{A}graphicData")
        if gd is None:
            return None
        chart = gd.find(C + "chart")
        if chart is None:
            return None
        rel = self._rels.get(chart.get(R + "id", ""))
        if rel is None:
            return None
        return DrawingObject("chart", self._xfrm(el, X + "xfrm"), name=name, hidden=hidden, chart_part=rel.target)

    def _group(self, el: ET.Element) -> DrawingObject:
        name, hidden = self._nv(el)
        gsppr = el.find(X + "grpSpPr")
        obj = DrawingObject("group", self._xfrm(gsppr), name=name, hidden=hidden)
        for child in _expand(list(el)):
            if child.tag in (X + "nvGrpSpPr", X + "grpSpPr"):
                continue
            o = self._object(child)
            if o is not None:
                obj.children.append(o)
        return obj

    # ------------------------------------------------------------------
    def _shape(self, el: ET.Element, connector: bool) -> DrawingObject:
        name, hidden = self._nv(el)
        sppr = el.find(X + "spPr")
        style = el.find(X + "style")
        obj = DrawingObject("connector" if connector else "shape", self._xfrm(sppr), name=name, hidden=hidden)

        style_fill = style_line = style_font = None
        line_idx = fill_idx = 0
        if style is not None:
            lr = style.find(A + "lnRef")
            fr = style.find(A + "fillRef")
            fo = style.find(A + "fontRef")
            if lr is not None:
                line_idx = int(lr.get("idx", "0"))
                style_line = self.colors.dml(lr)
            if fr is not None:
                fill_idx = int(fr.get("idx", "0"))
                style_fill = self.colors.dml(fr)
            if fo is not None:
                style_font = self.colors.dml(fo)

        geom = "line" if connector else "rect"
        if sppr is not None:
            pg = sppr.find(A + "prstGeom")
            if pg is not None:
                geom = pg.get("prst", geom)
                for gd in pg.iter(A + "gd"):
                    fmla = gd.get("fmla", "")
                    if fmla.startswith("val "):
                        try:
                            obj.adj[gd.get("name", "")] = int(fmla[4:]) / 100000.0
                        except ValueError:
                            pass
            elif sppr.find(A + "custGeom") is not None:
                geom = "rect"
        obj.geometry = geom

        # fill
        fill_set = False
        if sppr is not None:
            for child in sppr:
                t = child.tag.replace(A, "")
                if t == "noFill":
                    fill_set = True
                elif t == "solidFill":
                    obj.fill = self.colors.dml(child, style_fill)
                    obj.fill_alpha = self.colors.dml_alpha(child)
                    fill_set = True
                elif t == "gradFill":
                    stops = child.findall(f"{A}gsLst/{A}gs")
                    if stops:
                        obj.fill = self.colors.dml(stops[0], style_fill)
                    fill_set = True
                elif t == "pattFill":
                    obj.fill = self.colors.dml(child.find(A + "fgClr"), style_fill)
                    fill_set = True
        if not fill_set and fill_idx > 0 and not connector:
            obj.fill = style_fill

        # line
        ln = sppr.find(A + "ln") if sppr is not None else None
        default_line = None
        if line_idx > 0 and style_line is not None:
            default_line = LineStyle(style_line, [0.5, 0.75, 1.0, 1.5][min(line_idx, 3)])
        obj.line = self._line(ln, default_line)

        # text
        tx = el.find(X + "txBody")
        if tx is not None:
            obj.text = self._text_body(tx, style_font)
        return obj

    def _line(self, ln: Optional[ET.Element], default: Optional[LineStyle]) -> Optional[LineStyle]:
        if ln is None:
            return default
        if ln.find(A + "noFill") is not None:
            return None
        color = None
        sf = ln.find(A + "solidFill")
        alpha = 1.0
        if sf is not None:
            color = self.colors.dml(sf, default.color if default else None)
            alpha = self.colors.dml_alpha(sf)
        if color is None:
            if default is None:
                return None
            color = default.color
        width = _emu(ln.get("w")) if ln.get("w") else (default.width if default else 0.75)
        dash_el = ln.find(A + "prstDash")
        head = ln.find(A + "headEnd")
        tail = ln.find(A + "tailEnd")
        return LineStyle(
            color=color,
            width=width,
            dash=dash_el.get("val") if dash_el is not None else None,
            head=head.get("type") if head is not None and head.get("type") != "none" else None,
            tail=tail.get("type") if tail is not None and tail.get("type") != "none" else None,
            alpha=alpha,
        )

    def _text_body(self, tx: ET.Element, default_color: Optional[RGB]) -> TextBody:
        body = tx.find(A + "bodyPr")
        tb = TextBody(paragraphs=[])
        if body is not None:
            tb.anchor = body.get("anchor", "t")
            tb.insets = (
                _emu(body.get("lIns", "91440")),
                _emu(body.get("tIns", "45720")),
                _emu(body.get("rIns", "91440")),
                _emu(body.get("bIns", "45720")),
            )
            tb.wrap = body.get("wrap", "square") != "none"
            tb.vertical = body.get("vert")
        color = default_color or (0.0, 0.0, 0.0)
        for p in tx.findall(A + "p"):
            ppr = p.find(A + "pPr")
            para = Paragraph(runs=[], align=ppr.get("algn", "l") if ppr is not None else "l")
            for r in p:
                t = r.tag.replace(A, "")
                if t in ("r", "fld"):
                    te = r.find(A + "t")
                    text = te.text if te is not None and te.text else ""
                    para.runs.append(self._run(text, r.find(A + "rPr"), color))
                elif t == "br":
                    para.runs.append(self._run("\n", r.find(A + "rPr"), color))
                elif t == "endParaRPr":
                    sz = r.get("sz")
                    if sz:
                        para.end_size = int(sz) / 100.0
            tb.paragraphs.append(para)
        return tb

    def _run(self, text: str, rpr: Optional[ET.Element], color: RGB) -> TextRun:
        size, bold, italic, underline, strike, font = 11.0, False, False, False, False, None
        if rpr is not None:
            if rpr.get("sz"):
                size = int(rpr.get("sz")) / 100.0
            bold = rpr.get("b") in ("1", "true")
            italic = rpr.get("i") in ("1", "true")
            underline = rpr.get("u") not in (None, "none")
            strike = rpr.get("strike") not in (None, "noStrike")
            sf = rpr.find(A + "solidFill")
            if sf is not None:
                color = self.colors.dml(sf) or color
            latin = rpr.find(A + "latin")
            if latin is not None:
                font = self.theme.font_for(latin.get("typeface"))
        return TextRun(text, size, bold, italic, underline, strike, color, font)
