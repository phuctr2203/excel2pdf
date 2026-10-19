"""excel2pdf - convert Excel workbooks to PDF. MIT licensed.

    >>> import excel2pdf
    >>> excel2pdf.convert("report.xlsx", "report.pdf")
    >>> pdf_bytes = excel2pdf.convert("report.xlsx", paper_size="A4", fit_to_width=1)
    >>> excel2pdf.convert("report.xlsx", "report.pdf", engine="capture")  # use installed Excel/LibreOffice
"""

__version__ = "0.1.0"

from .api import ConvertOptions, convert  # noqa: E402
from .capture import CaptureError, available_engines  # noqa: E402

__all__ = ["convert", "ConvertOptions", "available_engines", "CaptureError", "__version__"]
