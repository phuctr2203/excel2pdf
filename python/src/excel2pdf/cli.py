"""Command line: excel2pdf input.xlsx [-o output.pdf] [options]"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import List, Optional

from . import __version__
from .api import ConvertOptions, convert
from .capture import CaptureError, available_engines


def _parse_sheets(value: Optional[str]):
    if not value:
        return None
    out = []
    for part in value.split(","):
        part = part.strip()
        if part.isdigit():
            out.append(int(part))
        elif part:
            out.append(part)
    return out


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="excel2pdf", description="Convert Excel workbooks (.xlsx/.xlsm) to PDF.")
    p.add_argument("inputs", nargs="*", help="workbook file(s)")
    p.add_argument("-o", "--output", help="output PDF (single input) or directory (several inputs)")
    p.add_argument("--engine", default="native", choices=["native", "capture", "excel", "libreoffice"],
                   help="native (default, pure Python) or capture with an installed Excel/LibreOffice")
    p.add_argument("--sheets", help="comma-separated sheet names or 0-based indexes")
    p.add_argument("--paper", help="paper size: A4, Letter, Legal, A3...")
    orient = p.add_mutually_exclusive_group()
    orient.add_argument("--landscape", action="store_true")
    orient.add_argument("--portrait", action="store_true")
    p.add_argument("--fit-page", action="store_true", help="fit each print area on one page")
    p.add_argument("--fit-width", type=int, help="fit to N pages wide")
    p.add_argument("--fit-height", type=int, help="fit to N pages tall")
    p.add_argument("--scale", type=int, help="print scale in percent")
    p.add_argument("--margins", type=float, help="all margins in inches")
    grid = p.add_mutually_exclusive_group()
    grid.add_argument("--gridlines", action="store_true", default=None)
    grid.add_argument("--no-gridlines", dest="gridlines", action="store_false")
    p.add_argument("--ignore-print-area", action="store_true")
    p.add_argument("--hidden-sheets", action="store_true", help="include hidden sheets")
    p.add_argument("--no-images", action="store_true")
    p.add_argument("--no-charts", action="store_true")
    p.add_argument("--no-shapes", action="store_true")
    p.add_argument("--font", action="append", default=[], metavar="FAMILY=PATH",
                   help="map a font family to a .ttf file (repeatable)")
    p.add_argument("--font-dir", action="append", default=[], help="extra folder to search for fonts")
    p.add_argument("--list-engines", action="store_true", help="show capture engines available here")
    p.add_argument("--version", action="version", version=f"excel2pdf {__version__}")
    return p


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    if args.list_engines:
        engines = available_engines()
        print("native (always available)")
        for e in engines:
            print(e)
        return 0
    if not args.inputs:
        build_parser().print_usage()
        return 2

    fonts = {}
    for item in args.font:
        fam, _, path = item.partition("=")
        if not path:
            print(f"--font expects FAMILY=PATH, got {item!r}", file=sys.stderr)
            return 2
        fonts[fam] = path

    orientation = "landscape" if args.landscape else ("portrait" if args.portrait else None)
    opts = ConvertOptions(
        sheets=_parse_sheets(args.sheets),
        engine=args.engine,
        paper_size=args.paper,
        orientation=orientation,
        margins=args.margins,
        scale=args.scale,
        fit_to_width=args.fit_width,
        fit_to_height=args.fit_height,
        fit_to_page=args.fit_page,
        gridlines=args.gridlines,
        ignore_print_area=args.ignore_print_area,
        include_hidden_sheets=args.hidden_sheets,
        images=not args.no_images,
        charts=not args.no_charts,
        shapes=not args.no_shapes,
        fonts=fonts or None,
        font_dirs=args.font_dir or None,
    )

    inputs = [Path(i) for i in args.inputs]
    out = Path(args.output) if args.output else None
    many = len(inputs) > 1
    if many and out is not None:
        out.mkdir(parents=True, exist_ok=True)
    status = 0
    for src in inputs:
        if many or out is None:
            dst = (out or src.parent) / (src.stem + ".pdf")
        else:
            dst = out
        try:
            convert(src, dst, opts)
            print(f"{src} -> {dst}")
        except (OSError, ValueError, KeyError, CaptureError) as exc:
            print(f"{src}: {exc}", file=sys.stderr)
            status = 1
    return status


if __name__ == "__main__":
    sys.exit(main())
