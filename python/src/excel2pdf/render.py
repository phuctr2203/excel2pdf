"""Render worksheets (cells, drawings, charts) onto reportlab canvas pages."""

from __future__ import annotations

import datetime as _dt
import io
import math
import re
import warnings
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

from openpyxl.cell.rich_text import CellRichText, TextBlock
from reportlab.lib.utils import ImageReader

from .chart import ChartReader
from .chart_render import ChartRenderer
from .colors import BLACK, RGB, WHITE, ColorResolver
from .condfmt import ConditionalFormats
from .drawing import DrawingObject, DrawingReader, TextBody
from .numfmt import Formatted, format_general, format_value
from .pagination import Page, PageSetup, paginate
from .sheet import PX_TO_PT, BorderSide, CellStyle, SheetGrid, StyleCache
from .text import FontSpec, Line, TextEngine

CELL_PAD = 2.0  # points of horizontal padding inside a cell

_BORDER_WEIGHT = {
    "hair": 1, "dotted": 2, "thin": 2, "dashDotDot": 3, "dashDot": 3, "dashed": 3,
    "mediumDashDotDot": 4, "slantDashDot": 4, "mediumDashDot": 4, "mediumDashed": 4,
    "medium": 5, "double": 6, "thick": 7,
}
_BORDER_LINE = {
    # style: (width, dash)
    "hair": (0.3, [0.75, 0.75]),
    "thin": (0.5, None),
    "dotted": (0.5, [0.75, 0.75]),
    "dashed": (0.5, [2.25, 0.75]),
    "dashDot": (0.5, [4.5, 1.5, 1.5, 1.5]),
    "dashDotDot": (0.5, [4.5, 1.5, 1.5, 1.5, 1.5, 1.5]),
    "medium": (1.0, None),
    "mediumDashed": (1.0, [6.75, 2.25]),
    "mediumDashDot": (1.0, [6.75, 2.25, 2.25, 2.25]),
    "mediumDashDotDot": (1.0, [6.75, 2.25, 2.25, 2.25, 2.25, 2.25]),
    "slantDashDot": (1.0, [6.75, 1.5, 2.25, 1.5]),
    "thick": (1.5, None),
    "double": (0.5, None),
}
_SHAPE_DASH = {
    "dash": [4, 3], "sysDash": [3, 1], "sysDot": [1, 1], "dot": [1, 2], "lgDash": [8, 3],
    "dashDot": [4, 2, 1, 2], "lgDashDot": [8, 3, 1, 3], "sysDashDot": [3, 1, 1, 1],
}


@dataclass
class Region:
    """Maps sheet coordinates (points, y down) into the page transform (y up)."""

    x0: float
    y0: float
    x1: float
    y1: float
    dx: float
    dy: float

    def X(self, x: float) -> float:
        return self.dx + (x - self.x0)

    def Y(self, y: float) -> float:
        return -(self.dy + (y - self.y0))

    @property
    def width(self) -> float:
        return self.x1 - self.x0

    @property
    def height(self) -> float:
        return self.y1 - self.y0


@dataclass
class PlacedDrawing:
    obj: DrawingObject
    rect: Tuple[float, float, float, float]  # x, y, w, h in sheet points


@dataclass
class SheetPlan:
    name: str
    pages: List[Page]
    setup: PageSetup
    renderer: "SheetRenderer"


class RenderContext:
    def __init__(self, wb, pkg, theme, colors: ColorResolver, text: TextEngine, styles: StyleCache,
                 options, chart_reader: ChartReader, drawing_reader: DrawingReader, filename: str) -> None:
        self.wb = wb
        self.pkg = pkg
        self.theme = theme
        self.colors = colors
        self.text = text
        self.styles = styles
        self.options = options
        self.chart_reader = chart_reader
        self.drawing_reader = drawing_reader
        self.chart_renderer = ChartRenderer(text, theme)
        self.filename = filename
        self.date1904 = bool(getattr(wb, "epoch", None) and wb.epoch.year == 1904)
        self._images: Dict[Tuple[int, Tuple[float, ...]], Optional[ImageReader]] = {}
        self._charts: Dict[str, object] = {}

    def image(self, data: bytes, crop: Tuple[float, float, float, float]) -> Optional[ImageReader]:
        key = (id(data), crop)
        if key in self._images:
            return self._images[key]
        reader = None
        try:
            from PIL import Image

            img = Image.open(io.BytesIO(data))
            img.load()
            if any(crop):
                l, t, r, b = (max(0.0, v) for v in crop)
                w, h = img.size
                box = (int(w * l), int(h * t), max(int(w * l) + 1, int(w * (1 - r))), max(int(h * t) + 1, int(h * (1 - b))))
                img = img.crop(box)
            if img.mode not in ("RGB", "RGBA", "L", "LA"):
                img = img.convert("RGBA")
            reader = ImageReader(img)
        except Exception as exc:
            warnings.warn(f"excel2pdf: skipping an image that cannot be decoded ({exc})")
        self._images[key] = reader
        return reader

    def chart(self, part: str):
        if part not in self._charts:
            root = self.pkg.xml(part)
            try:
                self._charts[part] = self.chart_reader.read(root) if root is not None else None
            except Exception as exc:  # never fail a whole document on one chart
                warnings.warn(f"excel2pdf: cannot read chart {part}: {exc}")
                self._charts[part] = None
        return self._charts[part]


# --------------------------------------------------------------------------
class SheetRenderer:
    def __init__(self, ctx: RenderContext, ws, part: str, setup: PageSetup, kind: str = "worksheet") -> None:
        self.ctx = ctx
        self.ws = ws
        self.part = part
        self.setup = setup
        self.kind = kind
        self.text = ctx.text
        self.styles = ctx.styles
        self.grid: Optional[SheetGrid] = None
        self.drawings: List[PlacedDrawing] = []
        self.cf: Optional[ConditionalFormats] = None
        self._eff: Dict[Tuple[int, int], Optional[CellStyle]] = {}

    # ---- planning -----------------------------------------------------------
    def plan(self, areas_override=None) -> List[Page]:
        ctx = self.ctx
        opts = ctx.options
        objs: List[DrawingObject] = []
        dpart = ctx.pkg.drawing_part(self.part)
        if dpart:
            try:
                objs = ctx.drawing_reader.read(dpart)
            except Exception as exc:
                warnings.warn(f"excel2pdf: cannot read drawing {dpart}: {exc}")
        if self.kind == "chartsheet":
            self.drawings = [PlacedDrawing(o, (0, 0, 0, 0)) for o in objs if o.kind == "chart"]
            return [Page(rows=(1, 1), cols=(1, 1), scale=1.0)] if self.drawings else []

        mdw = ctx.text.fonts.max_digit_width_px(ctx.styles.default_font.family, ctx.styles.default_font.size)
        self.grid = grid = SheetGrid(self.ws, self.styles, self.text, mdw, ctx.pkg.read_cols(self.part))
        self.cf = ConditionalFormats(self.ws, ctx.colors, self.styles) if opts.conditional_formatting else None

        for o in objs:
            if o.hidden:
                continue
            if o.kind == "picture" and not opts.images:
                continue
            if o.kind == "chart" and not opts.charts:
                continue
            if o.kind in ("shape", "connector", "group") and not opts.shapes:
                continue
            rect = self._anchor_rect(o)
            if rect is not None:
                self.drawings.append(PlacedDrawing(o, rect))

        areas = areas_override if areas_override is not None else self._print_areas()
        if not areas:
            return []
        for c0, r0, c1, r1 in areas:
            grid.autofit_rows(range(r0, r1 + 1), self.cell_runs_for_fit)
        pages: List[Page] = []
        for area in areas:
            pages.extend(paginate(grid, area, self.setup))
        return pages

    def _print_areas(self) -> List[Tuple[int, int, int, int]]:
        from .sheet import parse_areas

        grid = self.grid
        opts = self.ctx.options
        if not opts.ignore_print_area:
            areas = parse_areas(self.ws.print_area) if self.ws.print_area else []
            if areas:
                return [(c0, r0, min(c1, 16384), min(r1, 1048576)) for c0, r0, c1, r1 in areas]
        used = grid.used_range()
        if self.drawings:
            # first make sure row/col positions exist far enough
            dmax_c = dmax_r = 0
            dmin_c = dmin_r = 10 ** 9
            for d in self.drawings:
                x, y, w, h = d.rect
                if w <= 0 or h <= 0:
                    continue
                dmin_c = min(dmin_c, grid.col_at(x))
                dmin_r = min(dmin_r, grid.row_at(y))
                dmax_c = max(dmax_c, grid.col_at(max(x, x + w - 0.01)))
                dmax_r = max(dmax_r, grid.row_at(max(y, y + h - 0.01)))
            if dmax_c:
                if used is None:
                    used = (dmin_c, dmin_r, dmax_c, dmax_r)
                else:
                    used = (min(used[0], dmin_c), min(used[1], dmin_r), max(used[2], dmax_c), max(used[3], dmax_r))
        if used is None:
            return []
        # Excel prints from A1 unless the content starts far away; keep the top-left at A1
        c0, r0, c1, r1 = used
        return [(1, 1, c1, r1)]

    def _anchor_rect(self, o: DrawingObject) -> Optional[Tuple[float, float, float, float]]:
        a = o.anchor
        grid = self.grid
        if a is None or grid is None:
            return None
        if a.kind == "two" and a.frm and a.to:
            x1 = grid.col_x(a.frm.col + 1) + min(a.frm.col_off, grid.col_width(a.frm.col + 1))
            y1 = grid.row_y(a.frm.row + 1) + min(a.frm.row_off, grid.row_height(a.frm.row + 1))
            x2 = grid.col_x(a.to.col + 1) + min(a.to.col_off, grid.col_width(a.to.col + 1))
            y2 = grid.row_y(a.to.row + 1) + min(a.to.row_off, grid.row_height(a.to.row + 1))
            return (x1, y1, max(0.0, x2 - x1), max(0.0, y2 - y1))
        if a.kind == "one" and a.frm:
            x1 = grid.col_x(a.frm.col + 1) + a.frm.col_off
            y1 = grid.row_y(a.frm.row + 1) + a.frm.row_off
            return (x1, y1, a.ext[0], a.ext[1])
        if a.kind == "abs":
            return (a.pos[0], a.pos[1], a.ext[0], a.ext[1])
        return None

    # ---- cell content -----------------------------------------------------------
    def cell_runs(self, cell, st: CellStyle, font: Optional[FontSpec] = None) -> Tuple[List[Tuple[str, FontSpec]], Optional[Formatted]]:
        v = cell.value
        font = font or st.font
        if isinstance(v, CellRichText):
            runs = []
            for part in v:
                if isinstance(part, TextBlock):
                    runs.append((part.text, self.styles.font_spec(part.font, font)))
                else:
                    runs.append((str(part), font))
            return runs, None
        f = format_value(v, st.number_format, self.ctx.date1904, self.ctx.colors.indexed_color)
        spec = font
        if f.color is not None:
            spec = FontSpec(**{**font.__dict__, "color": f.color})
        return [(f.text, spec)], f

    def cell_runs_for_fit(self, cell, st: CellStyle):
        return self.cell_runs(cell, st)[0]

    def style_at(self, r: int, c: int, cell=None) -> Optional[CellStyle]:
        if cell is not None:
            return self.styles.get(cell)
        rd = self.ws.row_dimensions.get(r)
        if rd is not None and rd.customFormat and rd.has_style:
            return self.styles.get(rd)
        return None

    # ---- drawing a page -----------------------------------------------------------
    def draw_page(self, canvas, page: Page, page_no: int, total: int, sheet_page_index: int) -> None:
        setup = self.setup
        pw, ph = setup.page_size
        canvas.setPageSize((pw, ph))
        if self.kind == "chartsheet":
            self._draw_chartsheet(canvas, pw, ph)
            self._header_footer(canvas, pw, ph, page_no, total, sheet_page_index)
            return
        grid = self.grid
        s = page.scale
        rows, cols = page.rows, page.cols
        tr, tc = page.title_rows, page.title_cols
        tr_h = (grid.row_y(tr[1] + 1) - grid.row_y(tr[0])) if tr else 0.0
        tc_w = (grid.col_x(tc[1] + 1) - grid.col_x(tc[0])) if tc else 0.0
        body_w = grid.col_x(cols[1] + 1) - grid.col_x(cols[0])
        body_h = grid.row_y(rows[1] + 1) - grid.row_y(rows[0])
        content_w = (tc_w + body_w) * s
        content_h = (tr_h + body_h) * s
        ml, mt, mr, mb = setup.margins
        avail_w, avail_h = setup.printable
        ox = ml + (max(0.0, (avail_w - content_w) / 2) if setup.center_h else 0.0)
        oy = mt + (max(0.0, (avail_h - content_h) / 2) if setup.center_v else 0.0)

        canvas.saveState()
        canvas.translate(ox, ph - oy)
        canvas.scale(s, s)
        regions = []
        if tr and tc:
            regions.append(((tr[0], tr[1]), (tc[0], tc[1]), 0.0, 0.0))
        if tr:
            regions.append(((tr[0], tr[1]), cols, tc_w, 0.0))
        if tc:
            regions.append((rows, (tc[0], tc[1]), 0.0, tr_h))
        regions.append((rows, cols, tc_w, tr_h))
        for (r0, r1), (c0, c1), dx, dy in regions:
            reg = Region(grid.col_x(c0), grid.row_y(r0), grid.col_x(c1 + 1), grid.row_y(r1 + 1), dx, dy)
            self._draw_region(canvas, reg, r0, r1, c0, c1)
        canvas.restoreState()
        self._header_footer(canvas, pw, ph, page_no, total, sheet_page_index)

    def _draw_region(self, c, reg: Region, r0: int, r1: int, c0: int, c1: int) -> None:
        if reg.width <= 0 or reg.height <= 0:
            return
        c.saveState()
        p = c.beginPath()
        p.rect(reg.dx, -(reg.dy + reg.height), reg.width, reg.height)
        c.clipPath(p, stroke=0, fill=0)
        grid = self.grid
        rows = [r for r in range(r0, r1 + 1) if grid.row_height(r) > 0]
        cols = [k for k in range(c0, c1 + 1) if grid.col_width(k) > 0]
        cells = self.ws._cells

        if self.setup.gridlines:
            self._gridlines(c, reg, rows, cols)
        self._fills(c, reg, rows, cols, cells)
        self._texts(c, reg, rows, cols, cells, r0, r1, c0, c1)
        self._borders(c, reg, rows, cols, cells, r0, r1, c0, c1)
        self._drawings(c, reg)
        c.restoreState()

    # ---- fills / gridlines -----------------------------------------------------------
    def _gridlines(self, c, reg: Region, rows, cols) -> None:
        grid = self.grid
        c.setStrokeColorRGB(0.65, 0.65, 0.65)
        c.setLineWidth(0.3)
        x_a, x_b = reg.X(reg.x0), reg.X(reg.x1)
        y_a, y_b = reg.Y(reg.y0), reg.Y(reg.y1)
        for r in rows + [rows[-1] + 1] if rows else []:
            y = reg.Y(grid.row_y(r))
            c.line(x_a, y, x_b, y)
        for k in cols + [cols[-1] + 1] if cols else []:
            x = reg.X(grid.col_x(k))
            c.line(x, y_a, x, y_b)

    def _effective_style(self, r: int, k: int, cell) -> Optional[CellStyle]:
        key = (r, k)
        if key in self._eff:
            return self._eff[key]
        st = self.style_at(r, k, cell)
        if self.cf is not None and cell is not None:
            st = self.cf.apply(r, k, cell, st)
        self._eff[key] = st
        return st

    def _fills(self, c, reg: Region, rows, cols, cells) -> None:
        grid = self.grid
        done_merges = set()
        for r in rows:
            for k in cols:
                cell = cells.get((r, k))
                anchor = grid.merged_cover.get((r, k))
                if anchor is not None:
                    if anchor in done_merges:
                        continue
                    done_merges.add(anchor)
                    r0, c0, r1, c1 = grid.merged[anchor]
                    st = self._effective_style(anchor[0], anchor[1], cells.get(anchor))
                    x, y = grid.col_x(c0), grid.row_y(r0)
                    w, h = grid.col_x(c1 + 1) - x, grid.row_y(r1 + 1) - y
                else:
                    st = self._effective_style(r, k, cell)
                    x, y = grid.col_x(k), grid.row_y(r)
                    w, h = grid.col_width(k), grid.row_height(r)
                if st is None:
                    continue
                if st.gradient:
                    self._gradient(c, reg, x, y, w, h, st.gradient)
                elif st.fill is not None:
                    c.setFillColorRGB(*st.fill)
                    c.rect(reg.X(x), reg.Y(y + h), w, h, stroke=0, fill=1)
                db = st.data_bar
                if db is not None:
                    frac, color = db
                    c.setFillColorRGB(*color)
                    c.rect(reg.X(x) + 1, reg.Y(y + h) + 1.5, max(0.0, (w - 2) * frac), h - 3, stroke=0, fill=1)

    def _gradient(self, c, reg: Region, x, y, w, h, gradient) -> None:
        kind, degree, stops = gradient
        c.saveState()
        p = c.beginPath()
        p.rect(reg.X(x), reg.Y(y + h), w, h)
        c.clipPath(p, stroke=0, fill=0)
        colors = [col for _, col in stops]
        positions = [pos for pos, _ in stops]
        from reportlab.lib.colors import Color

        rl_colors = [Color(*col) for col in colors]
        cx, cy = reg.X(x) + w / 2, reg.Y(y + h) + h / 2
        if kind == "path":
            c.radialGradient(cx, cy, max(w, h) / 2, rl_colors, positions, extend=True)
        else:
            ang = math.radians(degree)
            dx, dy = math.cos(ang) * w / 2, -math.sin(ang) * h / 2
            c.linearGradient(cx - dx, cy - dy, cx + dx, cy + dy, rl_colors, positions, extend=True)
        c.restoreState()

    # ---- text -----------------------------------------------------------
    def _is_empty(self, r: int, k: int) -> bool:
        cell = self.ws._cells.get((r, k))
        if cell is not None and cell.value not in (None, ""):
            return False
        return (r, k) not in self.grid.merged_cover

    def _texts(self, c, reg: Region, rows, cols, cells, r0, r1, c0, c1) -> None:
        grid = self.grid
        row_set = set(rows)
        col_set = set(cols)
        # cells inside the region plus merged anchors that start outside but overlap it
        targets = []
        for r in rows:
            for k in cols:
                cell = cells.get((r, k))
                if cell is None or cell.value is None or cell.value == "":
                    continue
                if (r, k) in grid.merged_cover and grid.merged_cover[(r, k)] != (r, k):
                    continue
                targets.append((r, k, cell))
        for (ar, ak), (mr0, mc0, mr1, mc1) in grid.merged.items():
            if (ar in row_set and ak in col_set) or mr1 < r0 or mr0 > r1 or mc1 < c0 or mc0 > c1:
                continue
            cell = cells.get((ar, ak))
            if cell is not None and cell.value not in (None, ""):
                targets.append((ar, ak, cell))
        # also catch text overflowing in from the left/right neighbours of the region
        for r in rows:
            for k in (c0 - 1, c1 + 1):
                cell = cells.get((r, k))
                if cell is not None and isinstance(cell.value, str) and cell.value and (r, k) not in grid.merged_cover:
                    targets.append((r, k, cell))
        for r, k, cell in targets:
            st = self._effective_style(r, k, cell)
            if st is None:
                continue
            try:
                self._cell_text(c, reg, r, k, cell, st)
            except Exception as exc:  # keep going on odd content
                warnings.warn(f"excel2pdf: cannot draw cell {cell.coordinate}: {exc}")
            link = getattr(cell, "hyperlink", None)
            if link is not None and getattr(link, "target", None) and self.ctx.options.hyperlinks:
                x, y, w, h = self._cell_rect(r, k)
                c.linkURL(link.target, (reg.X(x), reg.Y(y + h), reg.X(x + w), reg.Y(y)), relative=1)

    def _cell_rect(self, r: int, k: int) -> Tuple[float, float, float, float]:
        grid = self.grid
        m = grid.merged.get((r, k))
        if m:
            r0, c0, r1, c1 = m
            x, y = grid.col_x(c0), grid.row_y(r0)
            return x, y, grid.col_x(c1 + 1) - x, grid.row_y(r1 + 1) - y
        return grid.col_x(k), grid.row_y(r), grid.col_width(k), grid.row_height(r)

    def _cell_text(self, c, reg: Region, r: int, k: int, cell, st: CellStyle) -> None:
        grid = self.grid
        x, y, w, h = self._cell_rect(r, k)
        if w <= 0 or h <= 0:
            return
        runs, fmt = self.cell_runs(cell, st)
        if not runs or all(not t for t, _ in runs):
            return
        value = cell.value
        is_number = bool(fmt and fmt.is_number)
        halign = st.h_align
        if halign == "general":
            if isinstance(value, bool) or cell.data_type == "e":
                halign = "center"
            elif is_number:
                halign = "right"
            else:
                halign = "left"
        indent = st.indent * grid.mdw * PX_TO_PT * 1.4
        pad = CELL_PAD
        avail = max(1.0, w - 2 * pad - (indent if halign in ("left", "right", "distributed") else 0))
        merged = (r, k) in grid.merged

        if st.rotation:
            self._rotated_text(c, reg, x, y, w, h, runs, st, halign)
            return

        wrap = st.wrap or halign in ("justify", "distributed") or st.v_align in ("justify", "distributed")
        if not wrap:
            runs = [(t.replace("\r\n", "").replace("\n", "").replace("\r", ""), s) for t, s in runs]
        if halign == "fill":
            unit = self.text.measure(runs)
            if unit > 0:
                reps = max(1, int(avail // unit))
                runs = runs * reps
            halign = "left"

        width = self.text.measure(runs)
        if is_number and not wrap and width > avail and fmt is not None:
            runs, width = self._shrink_number(value, st, runs, avail, fmt)
            fmt = None  # the accounting fill layout no longer applies

        if st.shrink and not wrap and width > avail:
            factor = max(0.1, avail / width)
            runs = [(t, FontSpec(**{**s.__dict__, "size": s.size * factor})) for t, s in runs]
            width = self.text.measure(runs)

        lines = self.text.layout(runs, avail if wrap else None)
        total_h = sum(ln.height for ln in lines)

        # clip rectangle, extended for overflowing text
        clip_x0, clip_x1 = x, x + w
        if not wrap and not merged and not is_number and width > avail and halign in ("left", "right", "center", "centerContinuous"):
            need = width - avail
            if halign == "left":
                clip_x1 = self._extend(r, k, +1, need, x + w)
            elif halign == "right":
                clip_x0 = self._extend(r, k, -1, need, x)
            else:
                clip_x1 = self._extend(r, k, +1, need / 2, x + w)
                clip_x0 = self._extend(r, k, -1, need / 2, x)
        if halign == "centerContinuous":
            halign = "center"

        c.saveState()
        p = c.beginPath()
        p.rect(reg.X(clip_x0), reg.Y(y + h), clip_x1 - clip_x0, h)
        c.clipPath(p, stroke=0, fill=0)

        # vertical placement (sheet coords, y down)
        if st.v_align == "top" or st.v_align in ("justify", "distributed"):
            top = y + 1.0
        elif st.v_align == "center":
            top = y + (h - total_h) / 2
        else:
            top = y + h - total_h - 0.5
        cur = top
        for ln in lines:
            baseline_sheet = cur + ln.ascent + 0.75
            if fmt is not None and fmt.fill_char is not None and len(lines) == 1 and 0 <= fmt.fill_at <= len(fmt.text):
                self._draw_filled(c, reg, x + pad, x + w - pad, baseline_sheet, fmt, runs[0][1])
            else:
                if halign in ("right",):
                    lx = x + w - pad - indent - ln.width
                elif halign == "center":
                    lx = x + (w - ln.width) / 2
                else:
                    lx = x + pad + indent
                self.text.draw_line(c, ln, reg.X(lx), reg.Y(baseline_sheet))
            cur += ln.height
        c.restoreState()

    def _shrink_number(self, value, st: CellStyle, runs, avail: float, fmt: Formatted):
        spec = runs[0][1]
        from .numfmt import to_serial

        serial = to_serial(value, self.ctx.date1904)
        if st.number_format in ("General", "general", None) and serial is not None:
            for chars in range(10, 0, -1):
                text = format_general(serial, chars)
                width = self.text.measure([(text, spec)])
                if width <= avail:
                    return [(text, spec)], width
        hash_w = self.text.measure([("#", spec)]) or 1.0
        n = max(1, int(avail // hash_w))
        text = "#" * n
        return [(text, spec)], self.text.measure([(text, spec)])

    def _extend(self, r: int, k: int, step: int, need: float, edge: float) -> float:
        grid = self.grid
        nk = k + step
        while need > 0 and 1 <= nk <= 16384:
            if not self._is_empty(r, nk):
                break
            w = grid.col_width(nk)
            edge += w * step
            need -= w
            nk += step
        return edge

    def _draw_filled(self, c, reg: Region, left: float, right: float, baseline: float, fmt: Formatted, spec: FontSpec) -> None:
        a, b = fmt.text[: fmt.fill_at], fmt.text[fmt.fill_at:]
        la = self.text.layout([(a, spec)], None)[0] if a else None
        lb = self.text.layout([(b, spec)], None)[0] if b else None
        if la is not None:
            self.text.draw_line(c, la, reg.X(left), reg.Y(baseline))
        if lb is not None:
            self.text.draw_line(c, lb, reg.X(right - lb.width), reg.Y(baseline))
        ch = fmt.fill_char
        if ch and ch != " ":
            gap_l = left + (la.width if la else 0.0)
            gap_r = right - (lb.width if lb else 0.0)
            unit = self.text.measure([(ch, spec)])
            if unit > 0 and gap_r > gap_l:
                reps = int((gap_r - gap_l) // unit)
                if reps > 0:
                    ln = self.text.layout([(ch * reps, spec)], None)[0]
                    self.text.draw_line(c, ln, reg.X(gap_l), reg.Y(baseline))

    def _rotated_text(self, c, reg: Region, x, y, w, h, runs, st: CellStyle, halign: str) -> None:
        rot = st.rotation
        c.saveState()
        p = c.beginPath()
        p.rect(reg.X(x), reg.Y(y + h), w, h)
        c.clipPath(p, stroke=0, fill=0)
        if rot == 255:
            chars = [(ch, s) for t, s in runs for ch in t if ch not in "\r\n"]
            lines = [self.text.layout([(ch, s)], None)[0] for ch, s in chars]
            total = sum(ln.height for ln in lines)
            cur = y + (h - total) / 2 if st.v_align == "center" else (y + 1 if st.v_align == "top" else y + h - total - 1)
            for ln in lines:
                self.text.draw_line(c, ln, reg.X(x + (w - ln.width) / 2), reg.Y(cur + ln.ascent + 0.75))
                cur += ln.height
            c.restoreState()
            return
        angle = rot if rot <= 90 else -(rot - 90)
        lines = self.text.layout([(t.replace("\n", " "), s) for t, s in runs], None)
        ln = lines[0]
        a = math.radians(abs(angle))
        bw = ln.width * math.cos(a) + ln.height * math.sin(a)
        bh = ln.width * math.sin(a) + ln.height * math.cos(a)
        if halign == "left":
            cx = x + CELL_PAD + bw / 2
        elif halign == "right":
            cx = x + w - CELL_PAD - bw / 2
        else:
            cx = x + w / 2
        if st.v_align == "top":
            cy = y + 1 + bh / 2
        elif st.v_align == "center":
            cy = y + h / 2
        else:
            cy = y + h - 1 - bh / 2
        c.translate(reg.X(cx), reg.Y(cy))
        c.rotate(angle)
        self.text.draw_line(c, ln, -ln.width / 2, -(ln.ascent - ln.descent) / 2)
        c.restoreState()

    # ---- borders -----------------------------------------------------------
    def _borders(self, c, reg: Region, rows, cols, cells, r0, r1, c0, c1) -> None:
        grid = self.grid
        h_edges: Dict[Tuple[int, int], BorderSide] = {}  # (row of top edge, col)
        v_edges: Dict[Tuple[int, int], BorderSide] = {}  # (row, col of left edge)
        diagonals = []

        def put(store, key, side: Optional[BorderSide]) -> None:
            if side is None:
                return
            cur = store.get(key)
            if cur is None or _BORDER_WEIGHT.get(side.style, 2) >= _BORDER_WEIGHT.get(cur.style, 2):
                store[key] = side

        row_set = set(rows) | {r1 + 1, r0 - 1}
        col_set = set(cols) | {c1 + 1, c0 - 1}
        for (r, k), cell in cells.items():
            if r not in row_set or k not in col_set:
                continue
            st = self._effective_style(r, k, cell)
            if st is None or not (st.left or st.right or st.top or st.bottom or st.diag_up or st.diag_down):
                continue
            m = grid.merged_cover.get((r, k))
            mr = grid.merged.get(m) if m else None
            if mr is None or r == mr[0]:
                put(h_edges, (r, k), st.top)
            if mr is None or r == mr[2]:
                put(h_edges, (r + 1, k), st.bottom)
            if mr is None or k == mr[1]:
                put(v_edges, (r, k), st.left)
            if mr is None or k == mr[3]:
                put(v_edges, (r, k + 1), st.right)
            if (st.diag_up or st.diag_down) and (mr is None or (r, k) == (mr[0], mr[1])):
                diagonals.append((r, k, st))
        # merged anchors: apply their outer border to the whole range
        for (ar, ak), (mr0, mc0, mr1, mc1) in grid.merged.items():
            if mr1 < r0 - 1 or mr0 > r1 + 1 or mc1 < c0 - 1 or mc0 > c1 + 1:
                continue
            cell = cells.get((ar, ak))
            st = self._effective_style(ar, ak, cell) if cell is not None else None
            if st is None:
                continue
            for kk in range(mc0, mc1 + 1):
                if (mr0, kk) not in h_edges:
                    put(h_edges, (mr0, kk), st.top)
                if (mr1 + 1, kk) not in h_edges and cells.get((mr1, kk)) is None:
                    put(h_edges, (mr1 + 1, kk), st.bottom)
            for rr in range(mr0, mr1 + 1):
                if (rr, mc0) not in v_edges:
                    put(v_edges, (rr, mc0), st.left)
                if (rr, mc1 + 1) not in v_edges and cells.get((rr, mc1)) is None:
                    put(v_edges, (rr, mc1 + 1), st.right)

        items = []
        for (r, k), side in h_edges.items():
            if grid.col_width(k) <= 0 or not (r0 <= r <= r1 + 1) or not (c0 <= k <= c1):
                continue
            items.append((_BORDER_WEIGHT.get(side.style, 2), "h", r, k, side))
        for (r, k), side in v_edges.items():
            if grid.row_height(r) <= 0 or not (r0 <= r <= r1) or not (c0 <= k <= c1 + 1):
                continue
            items.append((_BORDER_WEIGHT.get(side.style, 2), "v", r, k, side))
        items.sort(key=lambda t: t[0])
        for _, kind, r, k, side in items:
            width, dash = _BORDER_LINE.get(side.style, (0.5, None))
            c.setStrokeColorRGB(*side.color)
            c.setLineWidth(width)
            c.setDash(dash or [], 0)
            c.setLineCap(2 if dash is None else 0)
            if kind == "h":
                y = reg.Y(grid.row_y(r))
                xa, xb = reg.X(grid.col_x(k)), reg.X(grid.col_x(k + 1))
                if side.style == "double":
                    c.line(xa, y + 0.75, xb, y + 0.75)
                    c.line(xa, y - 0.75, xb, y - 0.75)
                else:
                    c.line(xa, y, xb, y)
            else:
                x = reg.X(grid.col_x(k))
                ya, yb = reg.Y(grid.row_y(r)), reg.Y(grid.row_y(r + 1))
                if side.style == "double":
                    c.line(x - 0.75, ya, x - 0.75, yb)
                    c.line(x + 0.75, ya, x + 0.75, yb)
                else:
                    c.line(x, ya, x, yb)
        for r, k, st in diagonals:
            x, y, w, h = self._cell_rect(r, k)
            for side, up in ((st.diag_up, True), (st.diag_down, False)):
                if side is None:
                    continue
                width, dash = _BORDER_LINE.get(side.style, (0.5, None))
                c.setStrokeColorRGB(*side.color)
                c.setLineWidth(width)
                c.setDash(dash or [], 0)
                if up:
                    c.line(reg.X(x), reg.Y(y + h), reg.X(x + w), reg.Y(y))
                else:
                    c.line(reg.X(x), reg.Y(y), reg.X(x + w), reg.Y(y + h))
        c.setDash([], 0)
        c.setLineCap(0)

    # ---- drawings -----------------------------------------------------------
    def _drawings(self, c, reg: Region) -> None:
        for d in self.drawings:
            x, y, w, h = d.rect
            if x >= reg.x1 or x + w <= reg.x0 or y >= reg.y1 or y + h <= reg.y0:
                continue
            try:
                self._draw_object(c, d.obj, reg, x, y, w, h)
            except Exception as exc:
                warnings.warn(f"excel2pdf: cannot draw {d.obj.kind} '{d.obj.name}': {exc}")

    def _draw_object(self, c, o: DrawingObject, reg: Region, x, y, w, h) -> None:
        if o.hidden or w <= 0 or h <= 0:
            if o.kind not in ("connector",) or (w <= 0 and h <= 0):
                return
        if o.kind == "picture":
            self._picture(c, o, reg, x, y, w, h)
        elif o.kind == "chart":
            chart = self.ctx.chart(o.chart_part) if o.chart_part else None
            if chart is not None:
                c.saveState()
                c.translate(reg.X(x), reg.Y(y + h))
                self.ctx.chart_renderer.draw(c, chart, w, h)
                c.restoreState()
        elif o.kind in ("shape", "connector"):
            self._shape(c, o, reg, x, y, w, h)
        elif o.kind == "group":
            xf = o.xfrm
            sx = w / xf.ch_w if xf.ch_w else 1.0
            sy = h / xf.ch_h if xf.ch_h else 1.0
            for child in o.children:
                cx = x + (child.xfrm.x - xf.ch_x) * sx
                cy = y + (child.xfrm.y - xf.ch_y) * sy
                self._draw_object(c, child, reg, cx, cy, child.xfrm.w * sx, child.xfrm.h * sy)

    def _picture(self, c, o: DrawingObject, reg: Region, x, y, w, h) -> None:
        img = self.ctx.image(o.image, o.crop) if o.image else None
        if img is None:
            return
        c.saveState()
        c.translate(reg.X(x + w / 2), reg.Y(y + h / 2))
        if o.xfrm.rot:
            c.rotate(-o.xfrm.rot)
        c.scale(-1 if o.xfrm.flip_h else 1, -1 if o.xfrm.flip_v else 1)
        c.drawImage(img, -w / 2, -h / 2, w, h, mask="auto")
        if o.line is not None:
            c.setStrokeColorRGB(*o.line.color)
            c.setLineWidth(o.line.width)
            c.rect(-w / 2, -h / 2, w, h, stroke=1, fill=0)
        c.restoreState()

    def _shape(self, c, o: DrawingObject, reg: Region, x, y, w, h) -> None:
        c.saveState()
        c.translate(reg.X(x + w / 2), reg.Y(y + h / 2))
        if o.xfrm.rot:
            c.rotate(-o.xfrm.rot)
        c.saveState()
        c.scale(-1 if o.xfrm.flip_h else 1, -1 if o.xfrm.flip_v else 1)
        path = _shape_path(c, o.geometry, w, h, o.adj)
        fill = o.fill if (o.kind == "shape" and o.geometry not in ("line", "straightConnector1")) else None
        line = o.line
        if fill is not None:
            c.setFillColorRGB(*fill)
            if o.fill_alpha < 1:
                c.setFillAlpha(o.fill_alpha)
        if line is not None:
            c.setStrokeColorRGB(*line.color)
            c.setLineWidth(line.width)
            c.setDash([d * max(line.width, 0.75) for d in _SHAPE_DASH.get(line.dash or "", [])] or [], 0)
            if line.alpha < 1:
                c.setStrokeAlpha(line.alpha)
        if path is not None and (fill is not None or line is not None):
            c.drawPath(path, stroke=1 if line else 0, fill=1 if fill else 0)
        if line is not None and o.geometry in _LINE_GEOMS:
            if line.tail:
                _arrow_head(c, (-w / 2, h / 2), (w / 2, -h / 2), line)
            if line.head:
                _arrow_head(c, (w / 2, -h / 2), (-w / 2, h / 2), line)
        c.restoreState()
        if o.text is not None:
            self._shape_text(c, o.text, w, h)
        c.restoreState()

    def _shape_text(self, c, tb: TextBody, w: float, h: float) -> None:
        li, ti, ri, bi = tb.insets
        avail = max(1.0, w - li - ri) if tb.wrap else None
        font_default = self.ctx.theme.minor_font
        para_lines: List[Tuple[List[Line], str]] = []
        for para in tb.paragraphs:
            runs = [(r.text, FontSpec(family=r.font or font_default, size=r.size, bold=r.bold, italic=r.italic,
                                      underline="single" if r.underline else None, strike=r.strike, color=r.color))
                    for r in para.runs if r.text]
            if not runs:
                runs = [("", FontSpec(family=font_default, size=para.end_size))]
            para_lines.append((self.text.layout(runs, avail), para.align))
        total = sum(ln.height for lines, _ in para_lines for ln in lines)
        top = h / 2 - ti
        if tb.anchor == "ctr":
            top = (h / 2 - ti + (-h / 2 + bi)) / 2 + total / 2
        elif tb.anchor == "b":
            top = -h / 2 + bi + total
        cur = top
        for lines, align in para_lines:
            for ln in lines:
                baseline = cur - ln.ascent - 0.75
                if align == "ctr":
                    lx = -ln.width / 2 + (li - ri) / 2
                elif align == "r":
                    lx = w / 2 - ri - ln.width
                else:
                    lx = -w / 2 + li
                self.text.draw_line(c, ln, lx, baseline)
                cur -= ln.height

    def _draw_chartsheet(self, c, pw, ph) -> None:
        ml, mt, mr, mb = self.setup.margins
        for d in self.drawings:
            chart = self.ctx.chart(d.obj.chart_part) if d.obj.chart_part else None
            if chart is None:
                continue
            c.saveState()
            c.translate(ml, mb)
            self.ctx.chart_renderer.draw(c, chart, pw - ml - mr, ph - mt - mb)
            c.restoreState()

    # ---- header / footer -----------------------------------------------------------
    def _header_footer(self, c, pw, ph, page_no: int, total: int, sheet_page_index: int) -> None:
        if not self.ctx.options.headers_footers:
            return
        ws = self.ws
        hf = getattr(ws, "HeaderFooter", None)
        if hf is None:
            return
        first = sheet_page_index == 0 and getattr(hf, "differentFirst", False)
        even = getattr(hf, "differentOddEven", False) and page_no % 2 == 0
        header = hf.firstHeader if first else (hf.evenHeader if even else hf.oddHeader)
        footer = hf.firstFooter if first else (hf.evenFooter if even else hf.oddFooter)
        ml, mt, mr, mb = self.setup.margins
        for item, is_header in ((header, True), (footer, False)):
            if item is None:
                continue
            for part, align in ((item.left, "l"), (item.center, "c"), (item.right, "r")):
                if part is None or not part.text:
                    continue
                text = _expand_hf(part.text, page_no, total, ws.title, self.ctx.filename)
                if not text.strip():
                    continue
                family, bold, italic = _hf_font(part.font, self.ctx.theme.minor_font)
                color = (0.0, 0.0, 0.0)
                if part.color:
                    from .colors import hex_to_rgb

                    color = hex_to_rgb(part.color)
                spec = FontSpec(family=family, size=float(part.size or 11), bold=bold, italic=italic, color=color)
                lines = self.text.layout([(text, spec)], None)
                total_h = sum(ln.height for ln in lines)
                if is_header:
                    cur = ph - self.setup.header_margin
                else:
                    cur = self.setup.footer_margin + total_h
                for ln in lines:
                    baseline = cur - ln.ascent - 0.75
                    if align == "l":
                        lx = ml
                    elif align == "r":
                        lx = pw - mr - ln.width
                    else:
                        lx = (pw - ln.width) / 2
                    self.text.draw_line(c, ln, lx, baseline)
                    cur -= ln.height


_LINE_GEOMS = {"line", "straightConnector1", "bentConnector2", "bentConnector3", "curvedConnector3", "lineInv"}


def _shape_path(c, geom: str, w: float, h: float, adj: dict):
    p = c.beginPath()
    hw, hh = w / 2, h / 2
    if geom in _LINE_GEOMS:
        if geom == "lineInv":
            p.moveTo(-hw, -hh)
            p.lineTo(hw, hh)
        elif geom == "bentConnector3":
            mid = -hw + w * adj.get("adj1", 0.5)
            p.moveTo(-hw, hh)
            p.lineTo(mid, hh)
            p.lineTo(mid, -hh)
            p.lineTo(hw, -hh)
        else:
            p.moveTo(-hw, hh)
            p.lineTo(hw, -hh)
        return p
    if geom in ("ellipse", "flowChartConnector"):
        p.ellipse(-hw, -hh, w, h)
        return p
    if geom in ("roundRect", "flowChartAlternateProcess"):
        r = min(w, h) * adj.get("adj", 0.16667)
        p.roundRect(-hw, -hh, w, h, r)
        return p
    pts: Optional[List[Tuple[float, float]]] = None
    if geom == "triangle":
        ax = -hw + w * adj.get("adj", 0.5)
        pts = [(ax, hh), (hw, -hh), (-hw, -hh)]
    elif geom == "rtTriangle":
        pts = [(-hw, hh), (hw, -hh), (-hw, -hh)]
    elif geom in ("diamond", "flowChartDecision"):
        pts = [(0, hh), (hw, 0), (0, -hh), (-hw, 0)]
    elif geom in ("parallelogram", "flowChartInputOutput"):
        o = min(w, h) * adj.get("adj", 0.25)
        pts = [(-hw + o, hh), (hw, hh), (hw - o, -hh), (-hw, -hh)]
    elif geom == "trapezoid":
        o = min(w, h) * adj.get("adj", 0.25)
        pts = [(-hw + o, hh), (hw - o, hh), (hw, -hh), (-hw, -hh)]
    elif geom in ("homePlate", "pentagon"):
        o = min(w, h) * adj.get("adj", 0.5)
        pts = [(-hw, hh), (hw - o, hh), (hw, 0), (hw - o, -hh), (-hw, -hh)]
    elif geom == "chevron":
        o = min(w, h) * adj.get("adj", 0.5)
        pts = [(-hw, hh), (hw - o, hh), (hw, 0), (hw - o, -hh), (-hw, -hh), (-hw + o, 0)]
    elif geom == "hexagon":
        o = min(w, h) * adj.get("adj", 0.25)
        pts = [(-hw + o, hh), (hw - o, hh), (hw, 0), (hw - o, -hh), (-hw + o, -hh), (-hw, 0)]
    elif geom == "octagon":
        o = min(w, h) * adj.get("adj", 0.29289)
        pts = [(-hw + o, hh), (hw - o, hh), (hw, hh - o), (hw, -hh + o), (hw - o, -hh), (-hw + o, -hh), (-hw, -hh + o), (-hw, hh - o)]
    elif geom in ("rightArrow", "leftArrow", "upArrow", "downArrow"):
        horizontal = geom in ("rightArrow", "leftArrow")
        L, T = (w, h) if horizontal else (h, w)
        shaft = T * adj.get("adj1", 0.5) / 2
        head = min(w, h) * adj.get("adj2", 0.5)
        base = [(-L / 2, shaft), (L / 2 - head, shaft), (L / 2 - head, T / 2), (L / 2, 0),
                (L / 2 - head, -T / 2), (L / 2 - head, -shaft), (-L / 2, -shaft)]
        if geom == "leftArrow":
            base = [(-px, py) for px, py in base]
        elif geom == "upArrow":
            base = [(py, px) for px, py in base]
        elif geom == "downArrow":
            base = [(py, -px) for px, py in base]
        pts = base
    elif geom in ("star5", "star4", "star6", "star8"):
        n = int(geom[4:])
        inner = 0.38 if n == 5 else 0.5
        pts = []
        for i in range(2 * n):
            ang = math.pi / 2 + i * math.pi / n
            rr = 1.0 if i % 2 == 0 else inner
            pts.append((math.cos(ang) * hw * rr, math.sin(ang) * hh * rr))
    elif geom in ("plus", "mathPlus"):
        t = min(w, h) * adj.get("adj", 0.25)
        pts = [(-hw + t, hh), (hw - t, hh), (hw - t, hh - t), (hw, hh - t), (hw, -hh + t), (hw - t, -hh + t),
               (hw - t, -hh), (-hw + t, -hh), (-hw + t, -hh + t), (-hw, -hh + t), (-hw, hh - t), (-hw + t, hh - t)]
    if pts is not None:
        p.moveTo(*pts[0])
        for pt in pts[1:]:
            p.lineTo(*pt)
        p.close()
        return p
    p.rect(-hw, -hh, w, h)
    return p


def _arrow_head(c, start, end, line) -> None:
    dx, dy = end[0] - start[0], end[1] - start[1]
    length = math.hypot(dx, dy)
    if length == 0:
        return
    ux, uy = dx / length, dy / length
    size = max(4.0, line.width * 3.5)
    bx, by = end[0] - ux * size, end[1] - uy * size
    px, py = -uy * size / 2, ux * size / 2
    c.setFillColorRGB(*line.color)
    p = c.beginPath()
    p.moveTo(*end)
    p.lineTo(bx + px, by + py)
    p.lineTo(bx - px, by - py)
    p.close()
    c.drawPath(p, stroke=0, fill=1)


_HF_CODE = re.compile(r"&([PNDTFAZGBIUSXYEKO&]|\"[^\"]*\"|\d+)")


def _expand_hf(text: str, page: int, total: int, sheet: str, filename: str) -> str:
    import os

    def repl(m: re.Match) -> str:
        code = m.group(1)
        if code == "P":
            return str(page)
        if code == "N":
            return str(total)
        if code == "D":
            return _dt.date.today().strftime("%m/%d/%Y")
        if code == "T":
            return _dt.datetime.now().strftime("%I:%M %p").lstrip("0")
        if code == "F":
            return os.path.basename(filename)
        if code == "Z":
            d = os.path.dirname(filename)
            return d + os.sep if d else ""
        if code == "A":
            return sheet
        if code == "&":
            return "&"
        return ""

    text = text.replace("&[Page]", "&P").replace("&[Pages]", "&N").replace("&[Date]", "&D")
    text = text.replace("&[Time]", "&T").replace("&[File]", "&F").replace("&[Tab]", "&A").replace("&[Path]", "&Z")
    # page arithmetic such as &P+1
    text = re.sub(r"&P([+-]\d+)", lambda m: str(page + int(m.group(1))), text)
    return _HF_CODE.sub(repl, text)


def _hf_font(font: Optional[str], default: str) -> Tuple[str, bool, bool]:
    if not font:
        return default, False, False
    name, _, style = font.partition(",")
    style = style.lower()
    if name in ("-", ""):
        name = default
    return name, "bold" in style, "italic" in style
