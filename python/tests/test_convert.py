import io
import os
import sys
from pathlib import Path

import pytest
from pypdf import PdfReader

import excel2pdf

sys.path.insert(0, str(Path(__file__).parent))
import make_samples  # noqa: E402


@pytest.fixture(scope="session")
def samples(tmp_path_factory):
    d = tmp_path_factory.mktemp("samples")
    make_samples.formatting_workbook(d / "formatting.xlsx")
    make_samples.charts_workbook(d / "charts.xlsx")
    return d


def text_of(pdf: bytes) -> str:
    return "\n".join(p.extract_text() or "" for p in PdfReader(io.BytesIO(pdf)).pages)


def test_formatting_workbook(samples):
    pdf = excel2pdf.convert(samples / "formatting.xlsx")
    assert pdf.startswith(b"%PDF")
    reader = PdfReader(io.BytesIO(pdf))
    assert len(reader.pages) >= 4
    text = text_of(pdf)
    for expected in ["Quarterly Report", "1,234,567.89", "(456,789.12)", "15-Jan-2026", "25.6%",
                     "Xin chào", "Customer 120", "Page 1 of", "42:00:00", "5 1/4"]:
        assert expected in text, expected


def test_print_titles_repeat(samples):
    pdf = excel2pdf.convert(samples / "formatting.xlsx", sheets=["Long table"])
    reader = PdfReader(io.BytesIO(pdf))
    assert len(reader.pages) > 1
    for page in reader.pages:
        assert "Comment" in page.extract_text()


def test_charts_workbook(samples):
    pdf = excel2pdf.convert(samples / "charts.xlsx", fit_to_page=True)
    text = text_of(pdf)
    assert len(PdfReader(io.BytesIO(pdf)).pages) == 1
    for title in ["Clustered column", "Stacked bar", "Line with markers", "Stacked area", "Pie (North)",
                  "Doughnut (South)", "Scatter", "Radar", "Bubble", "Combo"]:
        assert title in text, title
    assert "13%" in text  # pie data label


def test_options_paper_and_orientation(samples):
    pdf = excel2pdf.convert(samples / "formatting.xlsx", sheets=[0], paper_size="A4", orientation="landscape")
    box = PdfReader(io.BytesIO(pdf)).pages[0].mediabox
    assert round(float(box.width)) == 842 and round(float(box.height)) == 595


def test_inputs_and_outputs(samples, tmp_path):
    data = (samples / "formatting.xlsx").read_bytes()
    out = tmp_path / "out.pdf"
    assert excel2pdf.convert(data, out, sheets=["Conditional"]) is None
    assert out.read_bytes().startswith(b"%PDF")
    buf = io.BytesIO()
    excel2pdf.convert(io.BytesIO(data), buf, sheets=["Conditional"])
    assert buf.getvalue().startswith(b"%PDF")


def test_errors(samples, tmp_path):
    with pytest.raises(TypeError):
        excel2pdf.convert(samples / "formatting.xlsx", not_an_option=True)
    with pytest.raises(KeyError):
        excel2pdf.convert(samples / "formatting.xlsx", sheets=["Nope"])
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"not a zip")
    with pytest.raises(ValueError):
        excel2pdf.convert(bad)


def test_cli(samples, tmp_path):
    from excel2pdf.cli import main

    out = tmp_path / "cli.pdf"
    assert main([str(samples / "charts.xlsx"), "-o", str(out), "--fit-page"]) == 0
    assert out.read_bytes().startswith(b"%PDF")


@pytest.mark.skipif(not os.environ.get("EXCEL2PDF_TEST_CAPTURE"), reason="set EXCEL2PDF_TEST_CAPTURE=1 to drive Excel/LibreOffice")
def test_capture(samples):
    if not excel2pdf.available_engines():
        pytest.skip("no Excel or LibreOffice installed")
    pdf = excel2pdf.convert(samples / "formatting.xlsx", engine="capture")
    assert "Quarterly Report" in text_of(pdf)
