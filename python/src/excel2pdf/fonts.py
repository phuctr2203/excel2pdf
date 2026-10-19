"""Font discovery, registration with reportlab, metrics and glyph fallback.

Excel files name fonts by family ("Calibri", "Arial", "Aptos Narrow"). We look
for matching TrueType files in the system/Office font folders, register them
with reportlab (which embeds a subset in the PDF) and fall back to metric-
compatible substitutes, then to a Unicode-capable sans font, then to the PDF
base-14 Helvetica.
"""

from __future__ import annotations

import json
import logging
import os
import sys
import threading
import warnings
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

_FONT_EXT = (".ttf", ".ttc", ".otf")

# Metric-compatible / visually close substitutes, tried in order.
SUBSTITUTES: Dict[str, List[str]] = {
    "calibri": ["Carlito", "Aptos", "Arial", "Liberation Sans"],
    "calibri light": ["Calibri", "Carlito", "Arial"],
    "cambria": ["Caladea", "Georgia", "Times New Roman", "Liberation Serif"],
    "aptos": ["Calibri", "Carlito", "Arial"],
    "aptos narrow": ["Aptos", "Arial Narrow", "Calibri", "Carlito"],
    "aptos display": ["Aptos", "Calibri", "Carlito"],
    "arial": ["Liberation Sans", "Arimo", "Helvetica", "DejaVu Sans"],
    "helvetica": ["Arial", "Liberation Sans", "Arimo"],
    "arial narrow": ["Liberation Sans Narrow", "Arial"],
    "times new roman": ["Liberation Serif", "Tinos", "Times", "DejaVu Serif"],
    "times": ["Times New Roman", "Liberation Serif"],
    "courier new": ["Liberation Mono", "Cousine", "Courier", "DejaVu Sans Mono"],
    "consolas": ["Courier New", "Liberation Mono", "Menlo", "DejaVu Sans Mono"],
    "segoe ui": ["Arial", "Helvetica", "DejaVu Sans"],
    "tahoma": ["Verdana", "DejaVu Sans", "Arial"],
    "verdana": ["DejaVu Sans", "Tahoma", "Arial"],
    "georgia": ["Gelasio", "Times New Roman", "Liberation Serif"],
    "century gothic": ["URW Gothic", "Arial"],
    "garamond": ["EB Garamond", "Times New Roman"],
}

# Tried for characters the chosen font cannot draw (CJK, symbols, ...).
GLYPH_FALLBACKS = [
    "Arial Unicode MS", "Segoe UI", "Arial", "DejaVu Sans", "Noto Sans",
    "Liberation Sans", "Tahoma", "Microsoft YaHei", "SimSun", "MS Gothic",
    "Malgun Gothic", "Arial Unicode", "Segoe UI Symbol", "Noto Sans Symbols",
]

GENERIC_SANS = ["Arial", "Helvetica", "Liberation Sans", "DejaVu Sans", "Noto Sans", "Verdana"]

_SERIF_HINTS = ("times", "cambria", "georgia", "garamond", "serif", "book", "roman", "mincho", "song")
_MONO_HINTS = ("courier", "consolas", "mono", "menlo", "lucida console")

_BASE14 = {
    "sans": ("Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique"),
    "serif": ("Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic"),
    "mono": ("Courier", "Courier-Bold", "Courier-Oblique", "Courier-BoldOblique"),
}
# win ascent/descent (em fractions) roughly matching Arial/Times/Courier New
_BASE14_METRICS = {"sans": (0.905, 0.212), "serif": (0.891, 0.216), "mono": (0.833, 0.300)}


def system_font_dirs() -> List[Path]:
    home = Path.home()
    dirs: List[Path] = []
    if sys.platform == "win32":
        windir = os.environ.get("WINDIR", r"C:\Windows")
        dirs.append(Path(windir) / "Fonts")
        local = os.environ.get("LOCALAPPDATA")
        if local:
            dirs.append(Path(local) / "Microsoft" / "Windows" / "Fonts")
    elif sys.platform == "darwin":
        dirs += [
            Path("/System/Library/Fonts"),
            Path("/System/Library/Fonts/Supplemental"),
            Path("/Library/Fonts"),
            home / "Library" / "Fonts",
        ]
        # Microsoft Office for Mac ships Calibri, Cambria, Aptos... inside the app bundles.
        for app in ("Microsoft Excel", "Microsoft Word", "Microsoft PowerPoint", "Microsoft Outlook"):
            dirs.append(Path(f"/Applications/{app}.app/Contents/Resources/DFonts"))
    else:
        dirs += [
            Path("/usr/share/fonts"),
            Path("/usr/local/share/fonts"),
            home / ".fonts",
            home / ".local" / "share" / "fonts",
        ]
    return [d for d in dirs if d.is_dir()]


@dataclass
class FontEntry:
    path: str
    index: int
    family: str
    bold: bool
    italic: bool
    ascent: float  # em fraction
    descent: float  # em fraction, positive
    prio: int = 0


def _cache_file() -> Path:
    base = os.environ.get("XDG_CACHE_HOME")
    if sys.platform == "win32":
        base = base or os.environ.get("LOCALAPPDATA")
    elif sys.platform == "darwin":
        base = base or str(Path.home() / "Library" / "Caches")
    base = base or str(Path.home() / ".cache")
    return Path(base) / "excel2pdf" / "font-index-v2.json"


def _scan_file(path: str) -> List[dict]:
    """Read family/style/metrics of every face in a font file (TrueType outlines only)."""
    from fontTools.ttLib import TTCollection, TTFont as FTFont

    logging.getLogger("fontTools").setLevel(logging.ERROR)

    out: List[dict] = []
    try:
        if path.lower().endswith(".ttc"):
            with TTCollection(path, lazy=True) as coll:
                faces = list(coll.fonts)
                for i, f in enumerate(faces):
                    out += _describe(f, path, i)
        else:
            f = FTFont(path, lazy=True)
            try:
                out += _describe(f, path, 0)
            finally:
                f.close()
    except Exception:
        pass
    return out


def _describe(f, path: str, index: int) -> List[dict]:
    if "glyf" not in f or "name" not in f:
        return []  # reportlab can only embed TrueType outlines
    name = f["name"]

    def get(nid: int) -> Optional[str]:
        rec = name.getDebugName(nid)
        return rec.strip() if rec else None

    os2 = f["OS/2"] if "OS/2" in f else None
    head = f["head"]
    upm = float(head.unitsPerEm or 1000)
    if os2 is not None:
        fs = os2.fsSelection
        bold = bool(fs & 0x20) or os2.usWeightClass >= 600
        italic = bool(fs & 0x01)
        asc = os2.usWinAscent / upm
        desc = os2.usWinDescent / upm
        if os2.fsType & 0x0002 and not (os2.fsType & 0x000C):
            return []  # restricted license: cannot be embedded
    else:
        mac = head.macStyle
        bold, italic = bool(mac & 1), bool(mac & 2)
        hhea = f["hhea"]
        asc, desc = hhea.ascent / upm, abs(hhea.descent) / upm
    sub = (get(2) or "").lower()
    if "bold" in sub:
        bold = True
    if "italic" in sub or "oblique" in sub:
        italic = True
    ribbi = {"regular", "bold", "italic", "bold italic", "normal", "oblique", "bold oblique", "roman", "book"}
    typo_sub = (get(17) or get(2) or "").lower()
    out: List[dict] = []
    seen = set()
    # priority: legacy family (ID 1) beats typographic family (ID 16) beats full name (ID 4)
    for prio, fam in ((0, get(1)), (1 if typo_sub in ribbi else 3, get(16)), (2, get(4))):
        if fam and fam.lower() not in seen:
            seen.add(fam.lower())
            out.append(dict(path=path, index=index, family=fam, bold=bold, italic=italic,
                            ascent=asc, descent=desc, prio=prio))
    return out


class FontIndex:
    """Lazy, disk-cached index of installed TrueType fonts."""

    _lock = threading.Lock()

    def __init__(self, dirs: Sequence[Path], use_cache: bool = True) -> None:
        self.dirs = list(dirs)
        self.use_cache = use_cache
        self._by_family: Optional[Dict[str, List[FontEntry]]] = None

    def _files(self) -> Iterable[Path]:
        for d in self.dirs:
            for root, _, files in os.walk(d):
                for fn in files:
                    if fn.lower().endswith(_FONT_EXT):
                        yield Path(root) / fn

    def _build(self) -> Dict[str, List[FontEntry]]:
        cache_path = _cache_file()
        cache: Dict[str, dict] = {}
        if self.use_cache:
            try:
                cache = json.loads(cache_path.read_text("utf-8"))
            except Exception:
                cache = {}
        fresh: Dict[str, dict] = {}
        changed = False
        for p in self._files():
            key = str(p)
            try:
                mtime = p.stat().st_mtime
            except OSError:
                continue
            hit = cache.get(key)
            if hit and hit.get("mtime") == mtime:
                fresh[key] = hit
            else:
                fresh[key] = {"mtime": mtime, "faces": _scan_file(key)}
                changed = True
        if self.use_cache and (changed or len(fresh) != len(cache)):
            try:
                cache_path.parent.mkdir(parents=True, exist_ok=True)
                cache_path.write_text(json.dumps(fresh), "utf-8")
            except Exception:
                pass
        by_family: Dict[str, List[FontEntry]] = {}
        for rec in fresh.values():
            for face in rec["faces"]:
                by_family.setdefault(face["family"].lower(), []).append(FontEntry(**face))
        return by_family

    @property
    def families(self) -> Dict[str, List[FontEntry]]:
        if self._by_family is None:
            with self._lock:
                if self._by_family is None:
                    self._by_family = self._build()
        return self._by_family

    def add_file(self, path: str, family: Optional[str] = None) -> None:
        for face in _scan_file(path):
            if family:
                face["family"] = family
            self.families.setdefault(face["family"].lower(), []).append(FontEntry(**face))

    def find(self, family: str, bold: bool, italic: bool) -> Tuple[Optional[FontEntry], bool, bool]:
        """Return (entry, needs_fake_bold, needs_fake_italic)."""
        entries = self.families.get(family.lower())
        if not entries:
            return None, False, False
        entries = sorted(entries, key=lambda e: e.prio)
        for want_b, want_i in ((bold, italic), (bold, False), (False, italic), (False, False)):
            for e in entries:
                if e.bold == want_b and e.italic == want_i:
                    return e, bool(bold and not e.bold), bool(italic and not e.italic)
        e = entries[0]
        return e, bool(bold and not e.bold), bool(italic and not e.italic)


@dataclass
class FontFace:
    """A concrete font usable on a reportlab canvas."""

    name: str  # reportlab font name
    ascent: float  # em fraction
    descent: float  # em fraction (positive)
    fake_bold: bool = False
    fake_italic: bool = False
    is_ttf: bool = False
    _char_to_glyph: Optional[dict] = field(default=None, repr=False)
    _widths: Dict[str, float] = field(default_factory=dict, repr=False)

    def covers(self, ch: str) -> bool:
        if self._char_to_glyph is None:
            # base-14 fonts: WinAnsi only
            try:
                ch.encode("cp1252")
                return True
            except UnicodeEncodeError:
                return False
        return ord(ch) in self._char_to_glyph or ch in "\r\n\t"

    def char_width(self, ch: str) -> float:
        """Advance width for size 1."""
        w = self._widths.get(ch)
        if w is None:
            w = pdfmetrics.stringWidth(ch, self.name, 1.0)
            self._widths[ch] = w
        return w

    def width(self, text: str, size: float) -> float:
        return sum(self.char_width(c) for c in text) * size


class FontManager:
    def __init__(
        self,
        fonts: Optional[Dict[str, str]] = None,
        font_dirs: Optional[Sequence[str]] = None,
        use_system_fonts: bool = True,
        use_cache: bool = True,
    ) -> None:
        dirs = [Path(d) for d in (font_dirs or [])]
        if use_system_fonts:
            dirs += system_font_dirs()
        self.index = FontIndex(dirs, use_cache=use_cache)
        self._faces: Dict[Tuple[str, bool, bool], FontFace] = {}
        self._registered: Dict[Tuple[str, int], str] = {}
        self._fallbacks: Optional[List[FontFace]] = None
        self._explicit = {k.lower(): v for k, v in (fonts or {}).items()}
        self._warned: set = set()

    # ------------------------------------------------------------------
    def face(self, family: Optional[str], bold: bool = False, italic: bool = False) -> FontFace:
        family = (family or "Calibri").strip() or "Calibri"
        key = (family.lower(), bold, italic)
        f = self._faces.get(key)
        if f is None:
            f = self._resolve(family, bold, italic)
            self._faces[key] = f
        return f

    def _resolve(self, family: str, bold: bool, italic: bool) -> FontFace:
        fam_l = family.lower()
        if fam_l in self._explicit:
            entry = self._explicit_entry(fam_l, bold, italic)
            if entry is not None:
                face = self._register(*entry)
                if face is not None:
                    return face
        candidates = [family] + SUBSTITUTES.get(fam_l, [])
        if any(h in fam_l for h in _MONO_HINTS):
            candidates += ["Courier New", "Liberation Mono", "Menlo", "DejaVu Sans Mono"]
        elif any(h in fam_l for h in _SERIF_HINTS):
            candidates += ["Times New Roman", "Liberation Serif", "Times", "DejaVu Serif"]
        candidates += GENERIC_SANS
        for cand in candidates:
            if cand.lower() in self._explicit:
                entry = self._explicit_entry(cand.lower(), bold, italic)
                if entry:
                    face = self._register(*entry)
                    if face is not None:
                        return face
            e, fb, fi = self.index.find(cand, bold, italic)
            if e is not None:
                face = self._register(e, fb, fi)
                if face is not None:
                    if cand is not family and fam_l not in self._warned:
                        self._warned.add(fam_l)
                    return face
        return self._base14(fam_l, bold, italic)

    def _explicit_entry(self, fam_l: str, bold: bool, italic: bool):
        spec = self._explicit[fam_l]
        if isinstance(spec, dict):
            key = ("bolditalic" if bold and italic else "bold" if bold else "italic" if italic else "regular")
            path = spec.get(key) or spec.get("regular")
            fake_b = bold and not spec.get("bold") and key != "regular"
            fake_i = italic and not spec.get("italic")
        else:
            path, fake_b, fake_i = spec, bold, italic
        if not path:
            return None
        faces = _scan_file(str(path))
        if not faces:
            return None
        f = faces[0]
        return FontEntry(**f), fake_b, fake_i

    def _register(self, e: FontEntry, fake_bold: bool, fake_italic: bool) -> Optional[FontFace]:
        key = (e.path, e.index)
        name = self._registered.get(key)
        if name is None:
            name = f"E2P-{len(self._registered)}-{Path(e.path).stem[:24]}-{e.index}".replace(" ", "_")
            try:
                pdfmetrics.registerFont(TTFont(name, e.path, subfontIndex=e.index))
            except Exception as exc:  # unsupported/protected font
                warnings.warn(f"excel2pdf: cannot use font {e.path}: {exc}")
                self._registered[key] = ""
                return None
            self._registered[key] = name
        if not name:
            return None
        font = pdfmetrics.getFont(name)
        return FontFace(
            name=name,
            ascent=e.ascent or 0.9,
            descent=e.descent or 0.22,
            fake_bold=fake_bold,
            fake_italic=fake_italic,
            is_ttf=True,
            _char_to_glyph=font.face.charToGlyph,
        )

    def _base14(self, fam_l: str, bold: bool, italic: bool) -> FontFace:
        kind = "mono" if any(h in fam_l for h in _MONO_HINTS) else (
            "serif" if any(h in fam_l for h in _SERIF_HINTS) else "sans")
        names = _BASE14[kind]
        name = names[(1 if bold else 0) + (2 if italic else 0)]
        asc, desc = _BASE14_METRICS[kind]
        return FontFace(name=name, ascent=asc, descent=desc)

    # ------------------------------------------------------------------
    def fallbacks(self) -> List[FontFace]:
        if self._fallbacks is None:
            out: List[FontFace] = []
            seen = set()
            for fam in GLYPH_FALLBACKS:
                e, _, _ = self.index.find(fam, False, False)
                if e is None or (e.path, e.index) in seen:
                    continue
                seen.add((e.path, e.index))
                face = self._register(e, False, False)
                if face is not None:
                    out.append(face)
            self._fallbacks = out
        return self._fallbacks

    def split_by_coverage(self, text: str, face: FontFace) -> List[Tuple[str, FontFace]]:
        """Split text into segments, swapping in fallback faces for missing glyphs."""
        if all(face.covers(c) for c in text):
            return [(text, face)]
        segs: List[Tuple[str, FontFace]] = []
        cur_face, cur = face, []
        for ch in text:
            f = face
            if not face.covers(ch):
                f = next((fb for fb in self.fallbacks() if fb.covers(ch)), face)
                if f.fake_bold != face.fake_bold:
                    f = FontFace(**{**f.__dict__, "fake_bold": face.fake_bold})
            if f is not cur_face and cur:
                segs.append(("".join(cur), cur_face))
                cur = []
            cur_face = f
            cur.append(ch if f.covers(ch) else "?")
        if cur:
            segs.append(("".join(cur), cur_face))
        return segs

    def max_digit_width_px(self, family: str, size: float) -> int:
        """Excel's 'maximum digit width' in pixels at 96 dpi, used for column widths."""
        known = {("calibri", 11): 7, ("aptos narrow", 11): 7, ("aptos", 11): 7, ("arial", 10): 7,
                 ("calibri", 10): 7, ("arial", 11): 8, ("times new roman", 12): 8, ("cambria", 11): 8}
        k = (family.lower(), round(size))
        if k in known:
            return known[k]
        face = self.face(family)
        w = max(face.char_width(d) for d in "0123456789") * size * 96.0 / 72.0
        return max(1, int(round(w)))
