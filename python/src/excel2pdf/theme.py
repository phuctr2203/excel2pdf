"""Workbook theme: the color scheme and major/minor fonts from xl/theme/theme1.xml."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from typing import Dict, Optional

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_A = "{%s}" % NS_A

# Office 2013-2022 defaults, used when a workbook ships without a theme.
DEFAULT_COLORS: Dict[str, str] = {
    "dk1": "000000",
    "lt1": "FFFFFF",
    "dk2": "44546A",
    "lt2": "E7E6E6",
    "accent1": "4472C4",
    "accent2": "ED7D31",
    "accent3": "A5A5A5",
    "accent4": "FFC000",
    "accent5": "5B9BD5",
    "accent6": "70AD47",
    "hlink": "0563C1",
    "folHlink": "954F72",
}

# SpreadsheetML theme="n" indexes swap the first two pairs relative to the
# order inside <a:clrScheme>.
SPREADSHEET_ORDER = [
    "lt1", "dk1", "lt2", "dk2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
]

SCHEME_ALIASES = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}


class Theme:
    def __init__(
        self,
        colors: Optional[Dict[str, str]] = None,
        major_font: str = "Calibri Light",
        minor_font: str = "Calibri",
    ) -> None:
        self.colors = dict(DEFAULT_COLORS)
        if colors:
            self.colors.update(colors)
        self.major_font = major_font
        self.minor_font = minor_font

    @classmethod
    def from_xml(cls, data: Optional[bytes]) -> "Theme":
        if not data:
            return cls()
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            return cls()
        colors: Dict[str, str] = {}
        scheme = root.find(f".//{_A}clrScheme")
        if scheme is not None:
            for child in scheme:
                name = child.tag.replace(_A, "")
                value = _color_value(child)
                if value:
                    colors[name] = value
        major = root.find(f".//{_A}majorFont/{_A}latin")
        minor = root.find(f".//{_A}minorFont/{_A}latin")
        return cls(
            colors,
            major.get("typeface") if major is not None and major.get("typeface") else "Calibri Light",
            minor.get("typeface") if minor is not None and minor.get("typeface") else "Calibri",
        )

    def by_index(self, index: int) -> str:
        if 0 <= index < len(SPREADSHEET_ORDER):
            return self.colors[SPREADSHEET_ORDER[index]]
        return "000000"

    def by_name(self, name: str) -> Optional[str]:
        name = SCHEME_ALIASES.get(name, name)
        return self.colors.get(name)

    def font_for(self, typeface: Optional[str]) -> Optional[str]:
        """Resolve theme font references such as '+mn-lt' / '+mj-lt'."""
        if typeface in ("+mn-lt", "+mn-ea", "+mn-cs"):
            return self.minor_font
        if typeface in ("+mj-lt", "+mj-ea", "+mj-cs"):
            return self.major_font
        return typeface


def _color_value(el: ET.Element) -> Optional[str]:
    for child in el:
        tag = child.tag.replace(_A, "")
        if tag == "srgbClr":
            return child.get("val")
        if tag == "sysClr":
            return child.get("lastClr") or ("000000" if child.get("val") == "windowText" else "FFFFFF")
    return None
