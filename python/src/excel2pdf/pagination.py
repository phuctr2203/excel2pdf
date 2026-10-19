"""Split a print area into pages the way Excel's page layout does."""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from typing import List, Optional, Sequence, Tuple

from .sheet import SheetGrid

MM = 72.0 / 25.4

# Excel paperSize codes -> (width, height) in points, portrait.
PAPER_SIZES = {
    1: (612, 792), 2: (612, 792), 3: (792, 1224), 4: (1224, 792), 5: (612, 1008),
    6: (396, 612), 7: (522, 756), 8: (297 * MM, 420 * MM), 9: (210 * MM, 297 * MM),
    10: (210 * MM, 297 * MM), 11: (148 * MM, 210 * MM), 12: (257 * MM, 364 * MM),
    13: (182 * MM, 257 * MM), 14: (612, 936), 15: (215 * MM, 275 * MM), 16: (720, 1008),
    17: (792, 1224), 18: (612, 792), 19: (279, 639), 20: (297, 684), 21: (324, 747),
    22: (342, 792), 23: (360, 828), 24: (1224, 1584), 25: (1584, 2448), 26: (2448, 3168),
    27: (110 * MM, 220 * MM), 28: (162 * MM, 229 * MM), 29: (324 * MM, 458 * MM),
    30: (229 * MM, 324 * MM), 31: (114 * MM, 162 * MM), 32: (114 * MM, 229 * MM),
    33: (250 * MM, 353 * MM), 34: (176 * MM, 250 * MM), 35: (176 * MM, 125 * MM),
    36: (110 * MM, 230 * MM), 37: (279, 540), 38: (261, 468), 39: (1071, 792),
    40: (612, 864), 41: (612, 936), 66: (420 * MM, 594 * MM), 70: (105 * MM, 148 * MM),
}

NAMED_PAPER = {
    "letter": 1, "legal": 5, "a3": 8, "a4": 9, "a5": 11, "b4": 12, "b5": 13, "tabloid": 3,
    "ledger": 4, "executive": 7, "a2": 66, "a6": 70,
}


@dataclass
class PageSetup:
    paper: Tuple[float, float] = (612.0, 792.0)
    landscape: bool = False
    margins: Tuple[float, float, float, float] = (0.7 * 72, 0.75 * 72, 0.7 * 72, 0.75 * 72)  # l t r b
    header_margin: float = 0.3 * 72
    footer_margin: float = 0.3 * 72
    scale: float = 1.0
    fit_width: Optional[int] = None  # pages wide (0 = unconstrained) when fit-to-page is on
    fit_height: Optional[int] = None
    fit: bool = False
    center_h: bool = False
    center_v: bool = False
    gridlines: bool = False
    over_then_down: bool = False
    row_breaks: Sequence[int] = ()
    col_breaks: Sequence[int] = ()
    title_rows: Optional[Tuple[int, int]] = None
    title_cols: Optional[Tuple[int, int]] = None
    first_page_number: Optional[int] = None

    @property
    def page_size(self) -> Tuple[float, float]:
        w, h = self.paper
        if self.landscape:
            return (max(w, h), min(w, h))
        return (min(w, h), max(w, h))

    @property
    def printable(self) -> Tuple[float, float]:
        w, h = self.page_size
        l, t, r, b = self.margins
        return (max(10.0, w - l - r), max(10.0, h - t - b))


@dataclass
class Page:
    rows: Tuple[int, int]
    cols: Tuple[int, int]
    scale: float
    title_rows: Optional[Tuple[int, int]] = None
    title_cols: Optional[Tuple[int, int]] = None


def paper_from_option(value) -> Optional[Tuple[float, float]]:
    if value is None:
        return None
    if isinstance(value, (tuple, list)) and len(value) == 2:
        return (float(value[0]), float(value[1]))
    if isinstance(value, int):
        return PAPER_SIZES.get(value)
    code = NAMED_PAPER.get(str(value).strip().lower())
    if code:
        return PAPER_SIZES[code]
    raise ValueError(f"Unknown paper size: {value!r} (use e.g. 'A4', 'Letter' or (width_pt, height_pt))")


def _rowspan(text: Optional[str]) -> Optional[Tuple[int, int]]:
    if not text:
        return None
    m = re.search(r"\$?(\d+):\$?(\d+)", text)
    return (int(m.group(1)), int(m.group(2))) if m else None


def _colspan(text: Optional[str]) -> Optional[Tuple[int, int]]:
    from openpyxl.utils import column_index_from_string

    if not text:
        return None
    m = re.search(r"\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})", text)
    if not m:
        return None
    return (column_index_from_string(m.group(1).upper()), column_index_from_string(m.group(2).upper()))


def setup_from_sheet(ws) -> PageSetup:
    ps = ws.page_setup
    setup = PageSetup()
    try:
        code = int(ps.paperSize) if ps.paperSize is not None else 1
    except (TypeError, ValueError):
        code = 1
    setup.paper = PAPER_SIZES.get(code, PAPER_SIZES[1])
    setup.landscape = ps.orientation == "landscape"
    m = ws.page_margins
    setup.margins = (float(m.left) * 72, float(m.top) * 72, float(m.right) * 72, float(m.bottom) * 72)
    setup.header_margin = float(m.header or 0.3) * 72
    setup.footer_margin = float(m.footer or 0.3) * 72
    try:
        setup.scale = max(10, min(400, int(ps.scale))) / 100.0 if ps.scale else 1.0
    except (TypeError, ValueError):
        setup.scale = 1.0
    props = ws.sheet_properties.pageSetUpPr
    setup.fit = bool(props is not None and props.fitToPage)
    if setup.fit:
        setup.fit_width = int(ps.fitToWidth) if ps.fitToWidth is not None else 1
        setup.fit_height = int(ps.fitToHeight) if ps.fitToHeight is not None else 1
    po = ws.print_options
    setup.center_h = bool(po.horizontalCentered)
    setup.center_v = bool(po.verticalCentered)
    setup.gridlines = bool(po.gridLines)
    setup.over_then_down = ps.pageOrder == "overThenDown"
    setup.row_breaks = sorted(b.id for b in ws.row_breaks.brk if b.id)
    setup.col_breaks = sorted(b.id for b in ws.col_breaks.brk if b.id)
    setup.title_rows = _rowspan(ws.print_title_rows)
    setup.title_cols = _colspan(ws.print_title_cols)
    if ps.useFirstPageNumber and ps.firstPageNumber:
        setup.first_page_number = int(ps.firstPageNumber)
    return setup


def _bands(start: int, end: int, size_of, avail: float, breaks: set, title: Optional[Tuple[int, int]],
           title_size: float) -> List[Tuple[int, int]]:
    bands: List[Tuple[int, int]] = []
    cur_start = start
    used = 0.0
    i = start
    while i <= end:
        limit = avail
        if title is not None and cur_start > title[1]:
            limit = max(avail - title_size, avail * 0.2)
        s = size_of(i)
        if used + s > limit + 0.01 and i > cur_start:
            bands.append((cur_start, i - 1))
            cur_start, used = i, 0.0
            continue
        used += s
        if i in breaks and i < end:
            bands.append((cur_start, i))
            cur_start, used = i + 1, 0.0
        i += 1
    if cur_start <= end:
        bands.append((cur_start, end))
    # drop bands that are entirely hidden
    return [b for b in bands if sum(size_of(k) for k in range(b[0], b[1] + 1)) > 0] or bands[:1]


def paginate(grid: SheetGrid, area: Tuple[int, int, int, int], setup: PageSetup) -> List[Page]:
    c0, r0, c1, r1 = area
    avail_w, avail_h = setup.printable
    title_rows = setup.title_rows
    title_cols = setup.title_cols
    tr_h = sum(grid.row_height(r) for r in range(title_rows[0], title_rows[1] + 1)) if title_rows else 0.0
    tc_w = sum(grid.col_width(c) for c in range(title_cols[0], title_cols[1] + 1)) if title_cols else 0.0

    def layout(scale: float, use_breaks: bool):
        rb = set(setup.row_breaks) if use_breaks else set()
        cb = set(setup.col_breaks) if use_breaks else set()
        cols = _bands(c0, c1, grid.col_width, avail_w / scale, cb, title_cols, tc_w)
        rows = _bands(r0, r1, grid.row_height, avail_h / scale, rb, title_rows, tr_h)
        return cols, rows

    scale = setup.scale
    if setup.fit and (setup.fit_width or setup.fit_height):
        total_w = sum(grid.col_width(c) for c in range(c0, c1 + 1)) or 1.0
        total_h = sum(grid.row_height(r) for r in range(r0, r1 + 1)) or 1.0
        fw, fh = setup.fit_width or 0, setup.fit_height or 0
        est = 1.0
        if fw:
            est = min(est, fw * avail_w / total_w)
        if fh:
            est = min(est, fh * avail_h / total_h)
        pct = max(10, min(100, int(math.floor(est * 100 + 1e-6))))
        while True:
            cols, rows = layout(pct / 100.0, False)
            if (not fw or len(cols) <= fw) and (not fh or len(rows) <= fh):
                break
            if pct <= 10:
                break
            pct -= 1
        scale = pct / 100.0
        cols, rows = layout(scale, False)
    else:
        cols, rows = layout(scale, True)

    pages: List[Page] = []
    order = ([(rb, cb) for cb in cols for rb in rows] if not setup.over_then_down
             else [(rb, cb) for rb in rows for cb in cols])
    for rb, cb in order:
        pages.append(Page(
            rows=rb, cols=cb, scale=scale,
            title_rows=title_rows if title_rows and rb[0] > title_rows[1] else None,
            title_cols=title_cols if title_cols and cb[0] > title_cols[1] else None,
        ))
    return pages
