"""Unit tests for the connector table auto-write helpers.

Covers:
  - _format_cell_value: numeric -> string formatting (0%, 0.0%, 0, etc.)
  - _refresh_templated_count_with_records: 1x1 cells with embedded counts
  - _auto_compute_cell_values: 1xN vertical, Nx1 horizontal, NxM matrix
    layouts (the latter via integration with python-pptx tables)

The 1x1 templated-cell tests are pure-function and don't need a real
table shape. The N-dimensional layout tests need a python-pptx table
fixture; we build one in-memory via Presentation(blank).
"""
from __future__ import annotations

from io import BytesIO

import pandas as pd
import pytest
from pptx import Presentation
from pptx.util import Inches

from slidegen.intelligent_refresh import (
    _auto_compute_cell_values,
    _format_cell_value,
    _refresh_templated_count_with_records,
)


# ─────────────────────────────────────────────────────────────────────
# _format_cell_value
# ─────────────────────────────────────────────────────────────────────

def test_format_decimal_as_percent():
    assert _format_cell_value(0.73, "0%") == "73%"
    assert _format_cell_value(0.5, "0%") == "50%"


def test_format_whole_as_percent():
    """Heuristic: value > 1 in a "0%" format is taken as already-whole."""
    assert _format_cell_value(73, "0%") == "73%"


def test_format_one_decimal_percent():
    """0.0% format: one decimal place + percent sign."""
    out = _format_cell_value(0.733, "0.0%")
    assert out.endswith("%") and "73." in out


def test_format_integer():
    assert _format_cell_value(73.0, "0") == "73"


def test_format_thousands():
    assert _format_cell_value(12345, "#,##0") == "12,345"


def test_format_none_or_nan():
    assert _format_cell_value(None, "0%") == ""
    assert _format_cell_value(float("nan"), "0%") == ""


# ─────────────────────────────────────────────────────────────────────
# _refresh_templated_count_with_records
# ─────────────────────────────────────────────────────────────────────

@pytest.fixture
def df120():
    return pd.DataFrame({
        "respondents": [120, 120, 120],
        "base": [100, 100, 100],
        "count": [120, 80, 60],
    })


def test_count_substitution_n_paren(df120):
    assert _refresh_templated_count_with_records(
        "CARDs (n = 51)", df120) == "CARDs (n = 120)"


def test_count_substitution_capital_n(df120):
    assert _refresh_templated_count_with_records(
        "L CARDs (N = 51)", df120) == "L CARDs (N = 120)"


def test_count_substitution_no_paren(df120):
    assert _refresh_templated_count_with_records(
        "n=51", df120) == "n=120"


def test_count_substitution_sample_size(df120):
    assert _refresh_templated_count_with_records(
        "Sample size: 51", df120) == "Sample size: 120"


def test_count_substitution_base(df120):
    assert _refresh_templated_count_with_records(
        "Base: 51", df120) == "Base: 120"


def test_count_substitution_total_with_label(df120):
    assert _refresh_templated_count_with_records(
        "Total = 51 patients", df120) == "Total = 120 patients"


def test_count_substitution_based_on(df120):
    assert _refresh_templated_count_with_records(
        "based on 51 respondents", df120) == "based on 120 respondents"


def test_count_substitution_inline_with_unit(df120):
    assert _refresh_templated_count_with_records(
        "51 HCPs", df120) == "120 HCPs"


def test_count_substitution_no_pattern(df120):
    """No matching pattern -> returns None (caller preserves source)."""
    assert _refresh_templated_count_with_records(
        "Some unrelated text", df120) is None


def test_count_substitution_empty_text(df120):
    assert _refresh_templated_count_with_records("", df120) is None
    assert _refresh_templated_count_with_records(None, df120) is None


def test_count_substitution_empty_df():
    assert _refresh_templated_count_with_records(
        "(n = 51)", pd.DataFrame()) is None


# ─────────────────────────────────────────────────────────────────────
# _auto_compute_cell_values — layout detection
# ─────────────────────────────────────────────────────────────────────

def _new_table(rows: int, cols: int):
    """Create a minimal python-pptx table for layout tests."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout
    shape = slide.shapes.add_table(
        rows, cols, Inches(1), Inches(1),
        Inches(cols * 1.2), Inches(rows * 0.4),
    )
    return shape, slide, prs


def test_vertical_label_table_layout():
    """5x1 'L' table: row 0 = label preserved, rows 1-4 = formula values
    formatted as '0%'."""
    shape, _, _ = _new_table(5, 1)
    shape.table.cell(0, 0).text = "L"
    for i in range(1, 5):
        shape.table.cell(i, 0).text = "0%"  # placeholder — would be overwritten

    col_defs = [
        {"Name": "alias_label"},
        {"Name": "<blank:L>", "Formula": "=B2/100", "Alias": "L",
         "Format": "0%"},
    ]
    pivot = {"columnDefinitions": col_defs, "RowFields": ["alias_label"]}
    mapping = {"selectedColumns": ["<blank:L>"]}
    categories = ["A", "B", "C", "D"]
    series = [("L", [0.70, 0.71, 0.66, 0.65])]

    cv = _auto_compute_cell_values(shape, categories, series, pivot, mapping)
    assert cv == [["L"], ["70%"], ["71%"], ["66%"], ["65%"]]


def test_horizontal_label_table_layout():
    """1x5 horizontal: col 0 stays, cols 1-4 take series[0] values."""
    shape, _, _ = _new_table(1, 5)
    shape.table.cell(0, 0).text = "Header"
    for c in range(1, 5):
        shape.table.cell(0, c).text = "?"

    col_defs = [
        {"Name": "alias_label"},
        {"Name": "<blank:V>", "Format": "0%", "Alias": "V"},
    ]
    pivot = {"columnDefinitions": col_defs, "RowFields": ["alias_label"]}
    mapping = {"selectedColumns": ["<blank:V>"]}
    series = [("V", [0.5, 0.6, 0.7, 0.8])]
    cv = _auto_compute_cell_values(shape, ["a","b","c","d"], series, pivot, mapping)
    assert cv == [["Header", "50%", "60%", "70%", "80%"]]


def test_one_by_one_templated():
    """1x1 cell with '(n = X)' template — substitution via post-filter records."""
    shape, _, _ = _new_table(1, 1)
    shape.table.cell(0, 0).text = "CARDs (n = 51)"

    df = pd.DataFrame({"respondents": [120, 120], "base": [100, 100]})
    cv = _auto_compute_cell_values(
        shape, ["x"], [("S", [0.5])],
        {"columnDefinitions": []}, {"selectedColumns": []},
        filtered_records=df,
    )
    assert cv == [["CARDs (n = 120)"]]


def test_one_by_one_no_template_returns_none():
    """1x1 cell without a recognized count pattern -> None (preserves source)."""
    shape, _, _ = _new_table(1, 1)
    shape.table.cell(0, 0).text = "Static label"
    df = pd.DataFrame({"respondents": [120]})
    cv = _auto_compute_cell_values(
        shape, ["x"], [("S", [0.5])],
        {"columnDefinitions": []}, {"selectedColumns": []},
        filtered_records=df,
    )
    assert cv is None
