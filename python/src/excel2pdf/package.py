"""Low-level access to the .xlsx zip package (parts and relationships).

openpyxl covers cells and styles; drawings, charts and shapes are parsed
directly from the package here so that files with shapes openpyxl does not
understand still render their pictures and charts.
"""

from __future__ import annotations

import posixpath
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

REL_THEME = "/theme"
REL_DRAWING = "/drawing"


@dataclass
class Rel:
    id: str
    type: str
    target: str  # absolute part name inside the zip (no leading slash)
    external: bool = False


@dataclass
class SheetInfo:
    name: str
    part: str
    kind: str  # worksheet | chartsheet | dialogsheet | macrosheet
    state: str  # visible | hidden | veryHidden


class Package:
    def __init__(self, zf: zipfile.ZipFile) -> None:
        self.zf = zf
        self._names = set(zf.namelist())
        self._rels: Dict[str, Dict[str, Rel]] = {}
        self._xml: Dict[str, Optional[ET.Element]] = {}

    def has(self, part: str) -> bool:
        return part in self._names

    def read(self, part: str) -> Optional[bytes]:
        if part not in self._names:
            return None
        return self.zf.read(part)

    def xml(self, part: str) -> Optional[ET.Element]:
        if part not in self._xml:
            data = self.read(part)
            try:
                self._xml[part] = ET.fromstring(data) if data else None
            except ET.ParseError:
                self._xml[part] = None
        return self._xml[part]

    def rels(self, part: str) -> Dict[str, Rel]:
        if part in self._rels:
            return self._rels[part]
        folder, base = posixpath.split(part)
        rels_part = posixpath.join(folder, "_rels", base + ".rels")
        out: Dict[str, Rel] = {}
        root = self.xml(rels_part)
        if root is not None:
            for r in root:
                target = r.get("Target", "")
                external = r.get("TargetMode") == "External"
                if not external:
                    if target.startswith("/"):
                        target = target[1:]
                    else:
                        target = posixpath.normpath(posixpath.join(folder, target))
                out[r.get("Id", "")] = Rel(r.get("Id", ""), r.get("Type", ""), target, external)
        self._rels[part] = out
        return out

    def rel_by_type(self, part: str, suffix: str) -> Optional[Rel]:
        for rel in self.rels(part).values():
            if rel.type.endswith(suffix):
                return rel
        return None

    def workbook_part(self) -> str:
        rel = self.rel_by_type("", "/officeDocument")
        return rel.target if rel else "xl/workbook.xml"

    def sheets(self) -> List[SheetInfo]:
        wb = self.workbook_part()
        root = self.xml(wb)
        if root is None:
            return []
        rels = self.rels(wb)
        out = []
        for s in root.iter(f"{{{NS_MAIN}}}sheet"):
            rid = s.get(f"{{{NS_R}}}id", "")
            rel = rels.get(rid)
            if rel is None:
                continue
            kind = rel.type.rsplit("/", 1)[-1]
            out.append(SheetInfo(s.get("name", ""), rel.target, kind, s.get("state", "visible")))
        return out

    def theme_xml(self) -> Optional[bytes]:
        rel = self.rel_by_type(self.workbook_part(), REL_THEME)
        return self.read(rel.target) if rel else None

    def read_cols(self, sheet_part: str) -> List[Tuple[int, int, Optional[float], bool]]:
        """<cols> entries as (min, max, width or None, hidden), read without loading sheetData."""
        out: List[Tuple[int, int, Optional[float], bool]] = []
        if sheet_part not in self._names:
            return out
        col_tag = f"{{{NS_MAIN}}}col"
        data_tag = f"{{{NS_MAIN}}}sheetData"
        with self.zf.open(sheet_part) as fh:
            try:
                for event, el in ET.iterparse(fh, events=("start", "end")):
                    if event == "start" and el.tag == data_tag:
                        break
                    if event == "end" and el.tag == col_tag:
                        try:
                            lo = int(el.get("min", "1"))
                            hi = int(el.get("max", str(lo)))
                        except ValueError:
                            continue
                        w = el.get("width")
                        out.append((lo, hi, float(w) if w else None, el.get("hidden") in ("1", "true")))
            except ET.ParseError:
                pass
        return out

    def drawing_part(self, sheet_part: str) -> Optional[str]:
        rel = self.rel_by_type(sheet_part, REL_DRAWING)
        return rel.target if rel else None
