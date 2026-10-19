"""Rich-text layout: measuring, wrapping and drawing runs of styled text."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional, Sequence, Tuple

from .fonts import FontFace, FontManager

RGB = Tuple[float, float, float]

# Excel adds ~1.5pt of leading on top of the font's win ascent+descent.
LINE_PADDING = 1.5


@dataclass(frozen=True)
class FontSpec:
    family: str = "Calibri"
    size: float = 11.0
    bold: bool = False
    italic: bool = False
    underline: Optional[str] = None  # single, double, singleAccounting, doubleAccounting
    strike: bool = False
    color: RGB = (0.0, 0.0, 0.0)
    vert: Optional[str] = None  # superscript / subscript


@dataclass
class Piece:
    text: str
    spec: FontSpec
    face: FontFace
    size: float  # effective size (after super/subscript)
    rise: float
    width: float


@dataclass
class Line:
    pieces: List[Piece] = field(default_factory=list)
    width: float = 0.0
    ascent: float = 0.0
    descent: float = 0.0

    @property
    def height(self) -> float:
        return self.ascent + self.descent + LINE_PADDING


Run = Tuple[str, FontSpec]


class TextEngine:
    def __init__(self, fonts: FontManager) -> None:
        self.fonts = fonts

    def face_for(self, spec: FontSpec) -> FontFace:
        return self.fonts.face(spec.family, spec.bold, spec.italic)

    # ------------------------------------------------------------------
    def _pieces(self, text: str, spec: FontSpec) -> List[Piece]:
        face = self.face_for(spec)
        size = spec.size
        rise = 0.0
        if spec.vert == "superscript":
            size, rise = spec.size * 0.65, spec.size * 0.33
        elif spec.vert == "subscript":
            size, rise = spec.size * 0.65, -spec.size * 0.12
        out = []
        for seg, f in self.fonts.split_by_coverage(text, face):
            out.append(Piece(seg, spec, f, size, rise, f.width(seg, size)))
        return out

    def line_metrics(self, spec: FontSpec) -> Tuple[float, float]:
        face = self.face_for(spec)
        return face.ascent * spec.size, face.descent * spec.size

    def line_height(self, spec: FontSpec) -> float:
        a, d = self.line_metrics(spec)
        return a + d + LINE_PADDING

    def measure(self, runs: Sequence[Run]) -> float:
        return max((ln.width for ln in self.layout(runs, None)), default=0.0)

    # ------------------------------------------------------------------
    def layout(self, runs: Sequence[Run], max_width: Optional[float]) -> List[Line]:
        """Break runs into lines on explicit newlines and, if max_width is set, by wrapping."""
        # character stream with style
        chars: List[Tuple[str, FontSpec]] = []
        for text, spec in runs:
            for ch in text.replace("\r\n", "\n").replace("\r", "\n"):
                chars.append((ch, spec))
        paragraphs: List[List[Tuple[str, FontSpec]]] = [[]]
        for ch, spec in chars:
            if ch == "\n":
                paragraphs.append([])
            else:
                paragraphs[-1].append((ch, spec))
        default_spec = runs[0][1] if runs else FontSpec()
        lines: List[Line] = []
        for para in paragraphs:
            if max_width is None or not para:
                lines.append(self._make_line(para, default_spec))
                continue
            lines.extend(self._wrap(para, max_width, default_spec))
        return lines

    def _char_width(self, ch: str, spec: FontSpec) -> float:
        face = self.face_for(spec)
        size = spec.size * (0.65 if spec.vert in ("superscript", "subscript") else 1.0)
        if not face.covers(ch):
            for fb in self.fonts.fallbacks():
                if fb.covers(ch):
                    return fb.char_width(ch) * size
        return face.char_width(ch) * size

    def _wrap(self, para: List[Tuple[str, FontSpec]], max_width: float, default_spec: FontSpec) -> List[Line]:
        lines: List[Line] = []
        start = 0
        n = len(para)
        while start < n:
            width = 0.0
            last_break = -1
            i = start
            while i < n:
                ch, spec = para[i]
                w = self._char_width(ch, spec)
                if width + w > max_width and i > start and ch != " ":
                    break
                width += w
                if ch in " -​":
                    last_break = i
                i += 1
            if i >= n:
                lines.append(self._make_line(para[start:], default_spec))
                break
            if last_break >= start:
                end = last_break + 1
            else:
                end = max(i, start + 1)  # break inside a long word
            lines.append(self._make_line(para[start:end], default_spec))
            start = end
            while start < n and para[start][0] == " ":
                start += 1
        return lines

    def _make_line(self, chars: List[Tuple[str, FontSpec]], default_spec: FontSpec) -> Line:
        line = Line()
        # trailing spaces do not count toward alignment width
        while chars and chars[-1][0] == " ":
            chars = chars[:-1]
        groups: List[Tuple[str, FontSpec]] = []
        for ch, spec in chars:
            if groups and groups[-1][1] is spec:
                groups[-1] = (groups[-1][0] + ch, spec)
            else:
                groups.append((ch, spec))
        if not groups:
            a, d = self.line_metrics(default_spec)
            line.ascent, line.descent = a, d
            return line
        for text, spec in groups:
            for p in self._pieces(text, spec):
                line.pieces.append(p)
                line.width += p.width
                line.ascent = max(line.ascent, p.face.ascent * spec.size)
                line.descent = max(line.descent, p.face.descent * spec.size)
        return line

    # ------------------------------------------------------------------
    def draw_line(self, canvas, line: Line, x: float, baseline: float) -> None:
        for p in line.pieces:
            self.draw_piece(canvas, p, x, baseline)
            x += p.width

    def draw_piece(self, canvas, p: Piece, x: float, baseline: float) -> None:
        r, g, b = p.spec.color
        canvas.setFillColorRGB(r, g, b)
        y = baseline + p.rise
        face = p.face
        if face.fake_bold or face.fake_italic:
            canvas.saveState()
            canvas.translate(x, y)
            if face.fake_italic:
                canvas.transform(1, 0, 0.2, 1, 0, 0)
            t = canvas.beginText(0, 0)
            t.setFont(face.name, p.size)
            if face.fake_bold:
                canvas.setStrokeColorRGB(r, g, b)
                canvas.setLineWidth(p.size * 0.04)
                t.setTextRenderMode(2)
            t.textOut(p.text)
            canvas.drawText(t)
            canvas.restoreState()
        else:
            canvas.setFont(face.name, p.size)
            canvas.drawString(x, y, p.text)
        spec = p.spec
        if spec.underline or spec.strike:
            canvas.saveState()
            canvas.setStrokeColorRGB(r, g, b)
            lw = max(0.4, p.size * 0.055)
            canvas.setLineWidth(lw)
            if spec.underline:
                acc = "Accounting" in spec.underline
                uy = y - p.size * (0.22 if acc else 0.12)
                canvas.line(x, uy, x + p.width, uy)
                if spec.underline.startswith("double"):
                    uy2 = uy - lw * 2.2
                    canvas.line(x, uy2, x + p.width, uy2)
            if spec.strike:
                sy = y + p.size * 0.28
                canvas.line(x, sy, x + p.width, sy)
            canvas.restoreState()
