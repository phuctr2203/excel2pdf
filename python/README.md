# excel2pdf (Python)

Convert Excel workbooks (`.xlsx`, `.xlsm`) to PDF in pure Python. No Excel, no
LibreOffice, no native binaries: `pip install` and import. MIT licensed.

Not on PyPI yet. Install from GitHub (needs Python 3.9+ and Git):

```bash
pip install "git+https://github.com/phuctr2203/excel2pdf.git#subdirectory=python"
```

To pin a specific commit, add `@<commit>` after `.git`.

```python
import excel2pdf

excel2pdf.convert("report.xlsx", "report.pdf")

# bytes in, bytes out
pdf_bytes = excel2pdf.convert(open("report.xlsx", "rb").read())

# override page setup
excel2pdf.convert("report.xlsx", "report.pdf", paper_size="A4", orientation="landscape", fit_to_width=1)

# only some sheets
excel2pdf.convert("report.xlsx", "summary.pdf", sheets=["Summary", 2])
```

### Example: Excel vs. PDF

| The workbook in Microsoft Excel | The PDF written by excel2pdf |
|---|---|
| ![sales-report.xlsx open in Microsoft Excel](https://raw.githubusercontent.com/phuctr2203/excel2pdf/main/docs/images/excel.png) | ![PDF page produced by excel2pdf](https://raw.githubusercontent.com/phuctr2203/excel2pdf/main/docs/images/pdf.png) |

Made with `excel2pdf.convert("sales-report.xlsx", "sales-report.pdf")` from
[samples/sales-report.xlsx](https://github.com/phuctr2203/excel2pdf/blob/main/samples/sales-report.xlsx)
([PDF](https://github.com/phuctr2203/excel2pdf/blob/main/docs/sales-report.pdf)).

Command line:

```bash
excel2pdf report.xlsx -o report.pdf --paper A4 --fit-width 1
excel2pdf *.xlsx -o pdfs/
```

## What gets rendered

| Area | Supported |
|---|---|
| Cells | values, formula results (cached), rich text, Unicode with font fallback |
| Number formats | General, decimals, thousands, %, scientific, fractions, currency/accounting, dates/times, elapsed time, colors, conditions, text sections |
| Formatting | fonts, bold/italic/underline/strike, super/subscript, fills (solid, patterns, gradients), all border styles, diagonal borders, theme colors and tints |
| Layout | column widths, row heights (with auto-fit), hidden rows/columns, merged cells, wrap, shrink-to-fit, indent, rotation, vertical text, text overflow into empty neighbours, `####` for numbers that do not fit |
| Conditional formatting | cell value rules, text rules, top/bottom, above/below average, duplicates, color scales, data bars, formula rules (common functions) |
| Page setup | paper size, orientation, margins, scale, fit-to-pages, print area, print titles, manual page breaks, page order, centering, gridlines, headers/footers (`&P`, `&N`, `&A`, `&F`, `&D`, `&T`) |
| Pictures | PNG, JPEG, GIF, BMP, TIFF, with cropping, rotation and flips |
| Charts | column, bar (clustered/stacked/100%), line, area, pie, doughnut, scatter, bubble, radar, combo charts, secondary axes, titles, legends, axis titles, gridlines, data labels, number formats, log scale, reversed axes |
| Shapes | rectangles, rounded rectangles, ellipses, arrows, stars, lines/connectors with arrowheads, text boxes, groups |
| Chart sheets | rendered as a full page |

## Capture engine (optional)

For pixel-perfect output, or for things the native renderer does not draw,
`excel2pdf` can hand the workbook to a spreadsheet application that is already
installed on the machine:

```python
excel2pdf.convert("report.xlsx", "report.pdf", engine="capture")      # Excel if present, else LibreOffice
excel2pdf.convert("report.xlsx", "report.pdf", engine="excel")        # Microsoft Excel (Windows/macOS)
excel2pdf.convert("report.xlsx", "report.pdf", engine="libreoffice")  # LibreOffice (any OS)
excel2pdf.available_engines()  # e.g. ['excel', 'libreoffice']
```

The capture engine also opens legacy `.xls` files. Nothing extra is installed:
Excel is driven through PowerShell/COM on Windows and AppleScript on macOS, and
LibreOffice through `soffice --headless`. Set `EXCEL2PDF_SOFFICE` to point at a
LibreOffice binary in a non-standard place.

## Options

All options are keyword arguments of `convert()` (or fields of `ConvertOptions`).

| Option | Default | Meaning |
|---|---|---|
| `sheets` | all visible | sheet names or 0-based indexes |
| `engine` | `"native"` | `native`, `capture`, `excel`, `libreoffice` |
| `paper_size` | from workbook | `"A4"`, `"Letter"`, `"Legal"`, `"A3"`..., an Excel paper code, or `(width_pt, height_pt)` |
| `orientation` | from workbook | `"portrait"` or `"landscape"` |
| `margins` | from workbook | inches, one number or `(left, top, right, bottom)` |
| `scale` | from workbook | percent |
| `fit_to_width` / `fit_to_height` | from workbook | pages wide / tall (0 = no limit) |
| `fit_to_page` | `False` | each print area on one page |
| `gridlines` | from workbook | print cell gridlines |
| `ignore_print_area` | `False` | print the used range instead of the defined print area |
| `include_hidden_sheets` | `False` | also render hidden sheets |
| `images`, `charts`, `shapes`, `conditional_formatting`, `headers_footers`, `hyperlinks` | `True` | turn features off |
| `fonts` | `None` | `{"Calibri": "/path/Carlito-Regular.ttf"}` or `{"Calibri": {"regular": ..., "bold": ...}}` |
| `font_dirs` | `None` | extra folders searched for `.ttf`/`.ttc` files |
| `use_system_fonts` | `True` | search the OS (and Microsoft Office) font folders |

## Fonts

PDFs embed the fonts used. `excel2pdf` looks for the workbook's fonts in the
system font folders (and in Microsoft Office's own font folder on macOS). If a
font is missing it substitutes a metric-compatible one (for example *Carlito*
for *Calibri*, *Liberation Sans* for *Arial*), then any installed sans-serif
font, then the PDF built-in Helvetica. On servers and in Docker images install
`fonts-crosextra-carlito fonts-liberation` (Debian/Ubuntu) or pass `fonts=`.
Characters missing from a font (CJK, symbols) fall back to a Unicode font such
as Arial Unicode MS, Noto Sans or DejaVu Sans when present.

## Known limitations of the native engine

* Formulas are not recalculated: the value Excel saved last is printed. Files
  written by tools that do not calculate (e.g. openpyxl) show blanks for formulas.
* Not drawn: icon sets, sparklines, SmartArt, form controls, WordArt effects,
  3-D chart perspective (3-D charts are drawn flat), stock/surface charts
  (drawn as line/area), the new Excel 2016 chart types (waterfall, treemap,
  funnel...), EMF/WMF pictures on non-Windows systems.
* Numbers use en-US separators (`1,234.50`).

Use `engine="capture"` for those.

## License

MIT. All runtime dependencies are permissively licensed: openpyxl (MIT),
ReportLab (BSD), Pillow (MIT-CMU), fontTools (MIT).
