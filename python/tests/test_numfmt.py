import datetime as dt

import pytest

from excel2pdf.numfmt import format_general, format_value

ACCOUNTING = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'


@pytest.mark.parametrize(
    "value, fmt, expected",
    [
        (1234.5678, "General", "1234.5678"),
        (0.1 + 0.2, "General", "0.3"),
        (123456789012, "General", "1.23457E+11"),
        (0.000012345, "General", "1.2345E-05"),
        (1234.567, "#,##0.00", "1,234.57"),
        (-1234.567, "#,##0.00;(#,##0.00)", "(1,234.57)"),
        (0, '#,##0;-#,##0;"-"', "-"),
        (0.256, "0.0%", "25.6%"),
        (12345678, '#,##0,"K"', "12,346K"),
        (1234.5, "0.00E+00", "1.23E+03"),
        (5.25, "# ?/?", "5 1/4"),
        (1.5, "# ?/8", "1 4/8"),
        (5551234, "000-0000", "555-1234"),
        (0.5, "#.00", ".50"),
        (3.14159, "0.###", "3.142"),
        (-0.001, "0", "-0"),
        (12, "[$€-407] #,##0.00", "€ 12.00"),
        (1500, '[>=1000]#,##0,"K";0', "2K"),
        ("abc", '"Name: "@', "Name: abc"),
        (42, "@", "42"),
        (True, "General", "TRUE"),
        (dt.datetime(2026, 10, 19, 14, 5, 9), "yyyy-mm-dd hh:mm:ss", "2026-10-19 14:05:09"),
        (dt.datetime(2026, 10, 19, 14, 5), "dddd, mmmm d, yyyy h:mm AM/PM", "Monday, October 19, 2026 2:05 PM"),
        (dt.date(2026, 1, 3), "d-mmm-yy", "3-Jan-26"),
        (1.75, "[h]:mm", "42:00"),
        (dt.time(9, 30), "h:mm AM/PM", "9:30 AM"),
        (45000, "mm/dd/yyyy", "03/15/2023"),
        (dt.datetime(1900, 1, 1), "yyyy-mm-dd", "1900-01-01"),
        (dt.datetime(1900, 3, 1), "yyyy-mm-dd", "1900-03-01"),
    ],
)
def test_format(value, fmt, expected):
    assert format_value(value, fmt).text == expected


def test_section_color():
    r = format_value(-5, "0;[Red]-0")
    assert r.text == "-5"
    assert r.color == (1.0, 0.0, 0.0)


def test_accounting_fill_marker():
    r = format_value(-1234.5, ACCOUNTING)
    assert r.text == " $(1,234.50)"
    assert r.fill_char == " " and r.fill_at == 2


def test_general_shrinks():
    assert format_general(1234.5678, 6) == "1234.6"
