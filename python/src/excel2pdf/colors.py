"""Color resolution for SpreadsheetML (<color rgb/theme/indexed tint>) and DrawingML."""

from __future__ import annotations

import colorsys
import xml.etree.ElementTree as ET
from typing import Optional, Tuple

from openpyxl.styles.colors import COLOR_INDEX

from .theme import NS_A, Theme

RGB = Tuple[float, float, float]
_A = "{%s}" % NS_A

BLACK: RGB = (0.0, 0.0, 0.0)
WHITE: RGB = (1.0, 1.0, 1.0)

_PRESET = {
    "black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "008000",
    "blue": "0000FF", "yellow": "FFFF00", "gray": "808080", "grey": "808080",
    "darkGray": "A9A9A9", "lightGray": "D3D3D3", "orange": "FFA500",
}


def hex_to_rgb(value: str) -> RGB:
    value = value.strip().lstrip("#")[-6:].rjust(6, "0")
    return (int(value[0:2], 16) / 255.0, int(value[2:4], 16) / 255.0, int(value[4:6], 16) / 255.0)


def rgb_to_hex(rgb: RGB) -> str:
    return "".join(f"{max(0, min(255, round(c * 255))):02X}" for c in rgb)


def apply_tint(rgb: RGB, tint: float) -> RGB:
    """SpreadsheetML tint: lighten (tint>0) or darken (tint<0) in HLS space."""
    if not tint:
        return rgb
    h, l, s = colorsys.rgb_to_hls(*rgb)
    if tint < 0:
        l = l * (1.0 + tint)
    else:
        l = l * (1.0 - tint) + tint
    return colorsys.hls_to_rgb(h, max(0.0, min(1.0, l)), s)


def lum_mod_off(rgb: RGB, mod: float = 1.0, off: float = 0.0) -> RGB:
    h, l, s = colorsys.rgb_to_hls(*rgb)
    l = max(0.0, min(1.0, l * mod + off))
    return colorsys.hls_to_rgb(h, l, s)


def blend(a: RGB, b: RGB, t: float) -> RGB:
    """Mix: t=0 gives a, t=1 gives b."""
    return tuple(x + (y - x) * t for x, y in zip(a, b))  # type: ignore[return-value]


class ColorResolver:
    def __init__(self, theme: Theme, indexed: Optional[list] = None) -> None:
        self.theme = theme
        self.indexed = indexed or list(COLOR_INDEX)

    # ---- SpreadsheetML -------------------------------------------------
    def xl(self, color, default: Optional[RGB] = None) -> Optional[RGB]:
        """Resolve an openpyxl Color object."""
        if color is None:
            return default
        ctype = getattr(color, "type", None)
        tint = getattr(color, "tint", 0.0) or 0.0
        base: Optional[RGB] = None
        if ctype == "rgb":
            value = color.rgb
            if not isinstance(value, str):
                return default
            base = hex_to_rgb(value)
        elif ctype == "theme":
            base = hex_to_rgb(self.theme.by_index(int(color.theme)))
        elif ctype == "indexed":
            idx = int(color.indexed)
            if idx in (64, 65) or idx >= len(self.indexed):
                return default  # system foreground/background
            base = hex_to_rgb(self.indexed[idx])
        else:  # auto
            return default
        return apply_tint(base, tint)

    def indexed_color(self, n: int) -> Optional[RGB]:
        if 0 <= n < len(self.indexed):
            return hex_to_rgb(self.indexed[n])
        return None

    # ---- DrawingML -----------------------------------------------------
    def dml(self, parent: Optional[ET.Element], placeholder: Optional[RGB] = None) -> Optional[RGB]:
        """Resolve the color child of a DrawingML element such as <a:solidFill>."""
        if parent is None:
            return None
        for child in parent:
            tag = child.tag.replace(_A, "")
            base: Optional[RGB] = None
            if tag == "srgbClr":
                base = hex_to_rgb(child.get("val", "000000"))
            elif tag == "schemeClr":
                val = child.get("val", "tx1")
                if val == "phClr":
                    base = placeholder or BLACK
                else:
                    hexval = self.theme.by_name(val)
                    base = hex_to_rgb(hexval) if hexval else BLACK
            elif tag == "sysClr":
                base = hex_to_rgb(child.get("lastClr") or ("000000" if child.get("val") == "windowText" else "FFFFFF"))
            elif tag == "prstClr":
                base = hex_to_rgb(_PRESET.get(child.get("val", "black"), "000000"))
            elif tag == "scrgbClr":
                base = tuple(int(child.get(k, "0")) / 100000.0 for k in ("r", "g", "b"))  # type: ignore[assignment]
            elif tag == "hslClr":
                base = colorsys.hls_to_rgb(
                    int(child.get("hue", "0")) / 21600000.0,
                    int(child.get("lum", "0")) / 100000.0,
                    int(child.get("sat", "0")) / 100000.0,
                )
            if base is not None:
                return _apply_dml_mods(base, child)
        return None

    @staticmethod
    def dml_alpha(parent: Optional[ET.Element]) -> float:
        if parent is None:
            return 1.0
        for child in parent:
            for mod in child:
                if mod.tag == f"{_A}alpha":
                    return int(mod.get("val", "100000")) / 100000.0
        return 1.0


def _apply_dml_mods(rgb: RGB, el: ET.Element) -> RGB:
    mod, off = 1.0, 0.0
    has_lum = False
    for m in el:
        tag = m.tag.replace(_A, "")
        val = int(m.get("val", "0")) / 100000.0
        if tag == "lumMod":
            mod, has_lum = val, True
        elif tag == "lumOff":
            off, has_lum = val, True
        elif tag == "tint":
            rgb = tuple(c * val + (1.0 - val) for c in rgb)  # type: ignore[assignment]
        elif tag == "shade":
            rgb = tuple(c * val for c in rgb)  # type: ignore[assignment]
        elif tag == "satMod":
            h, l, s = colorsys.rgb_to_hls(*rgb)
            rgb = colorsys.hls_to_rgb(h, l, max(0.0, min(1.0, s * val)))
    if has_lum:
        rgb = lum_mod_off(rgb, mod, off)
    return rgb
