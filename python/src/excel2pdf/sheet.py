"""Worksheet geometry and resolved cell styles."""

from __future__ import annotations

import math
import re
from bisect import bisect_right
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from openpyxl.utils import range_boundaries

from .colors import BLACK, RGB, WHITE, ColorResolver, blend
from .text import FontSpec

PX_TO_PT = 0.75

# approximate ink coverage of Excel's pattern fills
_PATTERN_DENSITY = {
    "gray0625": 0.0625, "gray125": 0.125, "lightGray": 0.25, "mediumGray": 0.5, "darkGray": 0.75,
    "lightHorizontal": 0.25, "lightVertical": 0.25, "lightDown": 0.25, "lightUp": 0.25,
    "lightGrid": 0.4, "lightTrellis": 0.4, "darkHorizontal": 0.5, "darkVertical": 0.5,
    "darkDown": 0.5, "darkUp": 0.5, "darkGrid": 0.6, "darkTrellis": 0.7,
}


@dataclass(frozen=True)
class BorderSide:
    style: str
    color: RGB


@dataclass
class CellStyle:
    font: FontSpec
    fill: Optional[RGB] = None
    gradient: Optional[Tuple[str, float, List[Tuple[float, RGB]]]] = None
    left: Optional[BorderSide] = None
    right: Optional[BorderSide] = None
    top: Optional[BorderSide] = None
    bottom: Optional[BorderSide] = None
    diag_up: Optional[BorderSide] = None
    diag_down: Optional[BorderSide] = None
    h_align: str = "general"
    v_align: str = "bottom"
    wrap: bool = False
    shrink: bool = False
    indent: int = 0
    rotation: int = 0
    number_format: str = "General"
    data_bar: Optional[Tuple[float, RGB]] = None  # set by conditional formatting

    @property
    def has_visual(self) -> bool:
        return bool(self.fill or self.gradient or self.left or self.right or self.top or self.bottom
                    or self.diag_up or self.diag_down)


class StyleCache:
    """Resolve openpyxl styles once per distinct style index combination."""

    def __init__(self, wb, colors: ColorResolver, default_font: FontSpec) -> None:
        self.wb = wb
        self.colors = colors
        self.default_font = default_font
        self._cache: Dict[tuple, CellStyle] = {}

    def get(self, obj) -> CellStyle:
        key = tuple(obj._style) if getattr(obj, "_style", None) is not None else ()
        st = self._cache.get(key)
        if st is None:
            st = self._resolve(obj)
            self._cache[key] = st
        return st

    def font_spec(self, font, base: Optional[FontSpec] = None) -> FontSpec:
        base = base or self.default_font
        if font is None:
            return base
        name = getattr(font, "name", None) or getattr(font, "rFont", None) or base.family
        size = getattr(font, "sz", None) or getattr(font, "size", None) or base.size
        u = getattr(font, "u", None)
        color = self.colors.xl(getattr(font, "color", None), None)
        return FontSpec(
            family=str(name),
            size=float(size),
            bold=bool(font.b) if getattr(font, "b", None) is not None else base.bold,
            italic=bool(font.i) if getattr(font, "i", None) is not None else base.italic,
            underline=u if u not in (None, "none") else (None if u == "none" else base.underline),
            strike=bool(font.strike) if getattr(font, "strike", None) is not None else base.strike,
            color=color if color is not None else base.color,
            vert=getattr(font, "vertAlign", None) or base.vert,
        )

    def _resolve(self, obj) -> CellStyle:
        font = self.font_spec(obj.font)
        st = CellStyle(font=font)
        fill = obj.fill
        ftype = getattr(fill, "fill_type", None) if fill is not None else None
        if fill is not None and getattr(fill, "tagname", "") == "gradientFill":
            stops = []
            for stop in fill.stop or []:
                col = self.colors.xl(stop.color, WHITE) or WHITE
                stops.append((float(stop.position or 0.0), col))
            if stops:
                st.gradient = (fill.type or "linear", float(fill.degree or 0.0), stops)
        elif ftype == "solid":
            st.fill = self.colors.xl(fill.fgColor, None)
            if st.fill is None:
                st.fill = self.colors.xl(fill.bgColor, None)
        elif ftype and ftype != "none":
            fg = self.colors.xl(fill.fgColor, BLACK) or BLACK
            bg = self.colors.xl(fill.bgColor, WHITE) or WHITE
            st.fill = blend(bg, fg, _PATTERN_DENSITY.get(ftype, 0.5))
        border = obj.border
        if border is not None:
            st.left = self._side(border.left)
            st.right = self._side(border.right)
            st.top = self._side(border.top)
            st.bottom = self._side(border.bottom)
            diag = self._side(border.diagonal)
            if diag is not None:
                if border.diagonalUp:
                    st.diag_up = diag
                if border.diagonalDown:
                    st.diag_down = diag
        al = obj.alignment
        if al is not None:
            st.h_align = al.horizontal or "general"
            st.v_align = al.vertical or "bottom"
            st.wrap = bool(al.wrap_text)
            st.shrink = bool(al.shrink_to_fit)
            st.indent = int(al.indent or 0)
            st.rotation = int(al.text_rotation or 0)
        st.number_format = obj.number_format or "General"
        return st

    def _side(self, side) -> Optional[BorderSide]:
        if side is None or not side.style or side.style == "none":
            return None
        return BorderSide(side.style, self.colors.xl(side.color, BLACK) or BLACK)


def width_to_points(width: float, mdw: int) -> float:
    """Column width (characters, as stored in the file) to points."""
    if width <= 0:
        return 0.0
    px = math.trunc(((256 * width + math.trunc(128 / mdw)) / 256) * mdw)
    return px * PX_TO_PT


_AREA_RE = re.compile(r"(?:'(?:[^']|'')+'|[^!,]+)?!?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)")


def parse_areas(text: Optional[str]) -> List[Tuple[int, int, int, int]]:
    """'Sheet1'!$A$1:$D$20,'Sheet1'!$F$1:$G$5 -> [(min_col, min_row, max_col, max_row), ...]."""
    out = []
    if not text:
        return out
    for part in _split_top(text):
        ref = part.split("!")[-1].replace("$", "").strip()
        if not ref:
            continue
        if ":" not in ref:
            ref = f"{ref}:{ref}"
        try:
            c0, r0, c1, r1 = range_boundaries(ref)
        except ValueError:
            continue
        out.append((c0 or 1, r0 or 1, c1 or 16384, r1 or 1048576))
    return out


def _split_top(text: str) -> List[str]:
    parts, cur, quoted = [], [], False
    for ch in text:
        if ch == "'":
            quoted = not quoted
        if ch == "," and not quoted:
            parts.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
    parts.append("".join(cur))
    return parts


class SheetGrid:
    """Column/row positions in points, hidden flags and merged ranges."""

    def __init__(self, ws, styles: StyleCache, text_engine, mdw: int, cols=()) -> None:
        self.ws = ws
        self.styles = styles
        self.text = text_engine
        self.mdw = mdw
        fmt = ws.sheet_format
        if fmt.defaultColWidth:
            self.default_col = width_to_points(float(fmt.defaultColWidth), mdw)
        else:
            base = int(fmt.baseColWidth or 8)
            px = math.ceil((base * mdw + 5) / 8.0) * 8
            self.default_col = px * PX_TO_PT
        self.default_row = float(fmt.defaultRowHeight or 15.0)
        self.zero_height = bool(getattr(fmt, "zeroHeight", False))

        self._col_w: Dict[int, float] = {}
        for lo, hi, width, hidden in cols:
            if hidden:
                w = 0.0
            elif width is not None:
                w = width_to_points(width, mdw)
            else:
                continue
            for c in range(lo, min(hi, 16384) + 1):
                self._col_w[c] = w
        self._row_h: Dict[int, float] = {}
        for r, dim in ws.row_dimensions.items():
            if dim.hidden:
                self._row_h[r] = 0.0
            elif dim.ht is not None:
                self._row_h[r] = float(dim.ht)

        self.merged: Dict[Tuple[int, int], Tuple[int, int, int, int]] = {}
        self.merged_cover: Dict[Tuple[int, int], Tuple[int, int]] = {}
        for rng in ws.merged_cells.ranges:
            c0, r0, c1, r1 = rng.min_col, rng.min_row, rng.max_col, rng.max_row
            self.merged[(r0, c0)] = (r0, c0, r1, c1)
            if (r1 - r0 + 1) * (c1 - c0 + 1) <= 200000:
                for r in range(r0, r1 + 1):
                    for c in range(c0, c1 + 1):
                        self.merged_cover[(r, c)] = (r0, c0)

        self._col_x: List[float] = [0.0, 0.0]  # _col_x[c] = left edge of column c (1-based)
        self._row_y: List[float] = [0.0, 0.0]
        self._autofit_done = False

    # ---- sizes -----------------------------------------------------------
    def col_width(self, c: int) -> float:
        return self._col_w.get(c, self.default_col)

    def row_height(self, r: int) -> float:
        h = self._row_h.get(r)
        if h is not None:
            return h
        return 0.0 if self.zero_height else self.default_row

    def col_x(self, c: int) -> float:
        xs = self._col_x
        while len(xs) <= c:
            n = len(xs) - 1
            xs.append(xs[-1] + self.col_width(n))
        return xs[c]

    def row_y(self, r: int) -> float:
        ys = self._row_y
        while len(ys) <= r:
            n = len(ys) - 1
            ys.append(ys[-1] + self.row_height(n))
        return ys[r]

    def col_at(self, x: float) -> int:
        c = 1
        while self.col_x(c + 1) <= x and c < 16384:
            c += 1
        return c

    def row_at(self, y: float) -> int:
        self.row_y(1)
        r = bisect_right(self._row_y, y) - 1
        while self.row_y(r + 1) <= y and r < 1048576:
            r += 1
        return max(1, r)

    # ---- content -----------------------------------------------------------
    def used_range(self) -> Optional[Tuple[int, int, int, int]]:
        """(min_col, min_row, max_col, max_row) of cells with values or visible formatting."""
        ws = self.ws
        min_r = min_c = 10 ** 9
        max_r = max_c = 0
        for (r, c), cell in ws._cells.items():
            if cell.value is None or cell.value == "":
                if not self.styles.get(cell).has_visual:
                    continue
            if r < min_r:
                min_r = r
            if r > max_r:
                max_r = r
            if c < min_c:
                min_c = c
            if c > max_c:
                max_c = c
        for (r0, c0, r1, c1) in self.merged.values():
            if max_r and (r0 <= max_r + 1):
                min_r, min_c = min(min_r, r0), min(min_c, c0)
                max_r, max_c = max(max_r, r1), max(max_c, c1)
        if max_r == 0:
            return None
        return (min_c, min_r, max_c, max_r)

    def autofit_rows(self, rows: range, value_text) -> None:
        """Give rows without an explicit height the height Excel would auto-fit."""
        ws = self.ws
        auto_rows = [r for r in rows if r not in self._row_h]
        if not auto_rows:
            return
        need: Dict[int, float] = {}
        default_lh = self.text.line_height(self.styles.default_font)
        for (r, c), cell in ws._cells.items():
            if r not in rows or r in self._row_h or cell.value is None:
                continue
            if (r, c) in self.merged_cover:
                continue
            st = self.styles.get(cell)
            if st.rotation and st.rotation != 255:
                continue
            lh = self.text.line_height(st.font)
            if st.wrap or (isinstance(cell.value, str) and "\n" in cell.value and st.wrap):
                width = self.col_width(c) - 4.0 - st.indent * self.mdw * 1.0
                runs = value_text(cell, st)
                lines = self.text.layout(runs, max(width, 1.0))
                h = sum(ln.height for ln in lines)
            else:
                h = lh
            if h > need.get(r, 0.0):
                need[r] = h
        for r, h in need.items():
            if h > default_lh + 0.5:
                # rows snap to whole pixels (0.75pt)
                self._row_h[r] = max(self.default_row, math.ceil(h / PX_TO_PT) * PX_TO_PT)
        # cumulative positions are now stale
        self._row_y = [0.0, 0.0]
