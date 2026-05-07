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


def test_row_field_only_selected_columns_preserves_source():
    """CR_MR_Table case: selectedColumns=['y_code','y_label'] (both RowFields).

    Spec §17.9.5 says selectedColumns is an exclusive filter — when no
    selectedColumns entry maps to a value column, no value series should
    be written. Without this guard, the matrix layout would write
    series[0] (the first wave column) into physical col 1, overwriting
    the hand-curated message-text column on CREON slide 15 / 19 / 20 /
    21 / 22 with a wave percentage.

    Auto-compute must return None so the caller preserves source cells.
    """
    shape, _, _ = _new_table(13, 3)
    # Source table: id | message text | %
    shape.table.cell(0, 0).text = "ID"
    shape.table.cell(0, 1).text = "Message"
    shape.table.cell(0, 2).text = "%"
    shape.table.cell(1, 0).text = "C6"
    shape.table.cell(1, 1).text = "Take CREON every meal..."
    shape.table.cell(1, 2).text = "57%"

    pivot = {
        "RowFields": ["y_code", "y_label"],
        "ColumnFields": ["time_period_name"],
        "ValueFields": ["decimal"],
        "columnDefinitions": [
            {"Name": "y_code", "IsDefaultAlias": True},
            {"Name": "y_label", "IsDefaultAlias": True},
            {"Name": "Mar'26", "Format": "0%", "IsDefaultAlias": True},
        ],
    }
    mapping = {"selectedColumns": ["y_code", "y_label"], "selectAllRows": False}

    # Series passed in as if from the pivot — wave columns the user
    # explicitly excluded by setting selectedColumns to row-fields only.
    series = [("Mar'26", [0.566, 0.548]), ("Feb'26", [0.50, 0.48])]
    categories = ["msg1", "msg2"]

    cv = _auto_compute_cell_values(shape, categories, series, pivot, mapping)
    assert cv is None, (
        f"Expected None (preserve source) when all selectedColumns are RowFields, "
        f"got {cv!r}. Without this guard, col 1 of the table would be overwritten "
        f"with the first wave's percentage value."
    )


def test_multi_row_field_columns_preserved_in_matrix_layout():
    """Datroway slide 25 Table 4 case: 5-col table with selectedColumns =
    ['ID','Tags','y_label','Q4_2025_value','Q1_2026_value']. Three leading
    RowField cols must be preserved from source; only the 2 trailing wave
    value cols should be written from series. Legacy matrix layout would
    write series[0]/series[1] into the Tags and y_label columns,
    overwriting hand-curated message text with wave %."""
    shape, _, _ = _new_table(11, 5)
    # Source table layout: ID | Tags | Abbreviated Message | Q4'25 | Q1'26
    shape.table.cell(0, 0).text = "ID"
    shape.table.cell(0, 1).text = "Tags"
    shape.table.cell(0, 2).text = "Abbreviated Message"
    shape.table.cell(0, 3).text = "Q4 '25"
    shape.table.cell(0, 4).text = "Q1 '26"
    shape.table.cell(1, 0).text = "D6"
    shape.table.cell(1, 1).text = "Efficacy - ORR"
    shape.table.cell(1, 2).text = "45% ORR was observed with DATROWAY..."
    shape.table.cell(1, 3).text = "66%"
    shape.table.cell(1, 4).text = "70%"

    pivot = {
        "RowFields": ["Tags", "ID", "y_label"],
        "ColumnFields": ["time_period_name", "x_label"],
        "ValueFields": ["decimal", "me_score"],
        "columnDefinitions": [
            {"Name": "ID", "IsDefaultAlias": True},
            {"Name": "Tags", "IsDefaultAlias": True},
            {"Name": "y_label", "IsDefaultAlias": True},
            {"Name": "Q4 2025 @:@ Believable @:@ Average of me_score",
             "Format": "0%", "IsDefaultAlias": True},
            {"Name": "Q1 2026 @:@ Believable @:@ Average of me_score",
             "Format": "0%", "IsDefaultAlias": True},
        ],
    }
    mapping = {
        "selectedColumns": [
            "ID", "Tags", "y_label",
            "Q4 2025 @:@ Believable @:@ Average of me_score",
            "Q1 2026 @:@ Believable @:@ Average of me_score",
        ],
        "selectAllRows": False,
    }
    series = [
        ("Q4 2025 @:@ Believable @:@ Average of me_score", [0.65, 0.59]),
        ("Q1 2026 @:@ Believable @:@ Average of me_score", [0.71, 0.66]),
    ]
    categories = ["msg1", "msg2"]

    cv = _auto_compute_cell_values(shape, categories, series, pivot, mapping)
    assert cv is not None, "Expected matrix layout to fire (5 cols, 11 rows)"
    # Header row: preserved verbatim
    assert cv[0] == ["ID", "Tags", "Abbreviated Message", "Q4 '25", "Q1 '26"]
    # Data row 1: cols 0-2 preserved (RowFields), cols 3-4 refreshed (waves)
    assert cv[1][0] == "D6", f"col 0 (ID) must be preserved, got {cv[1][0]!r}"
    assert cv[1][1] == "Efficacy - ORR", (
        f"col 1 (Tags) must be preserved as text, got {cv[1][1]!r}. "
        f"Without the multi-row-field guard, series[0] (Q4 2025 = 65%) "
        f"would have been written here."
    )
    assert cv[1][2] == "45% ORR was observed with DATROWAY...", (
        f"col 2 (y_label/Message) must be preserved as text, got {cv[1][2]!r}"
    )
    assert cv[1][3] == "65%", f"col 3 (Q4'25) must be refreshed, got {cv[1][3]!r}"
    assert cv[1][4] == "71%", f"col 4 (Q1'26) must be refreshed, got {cv[1][4]!r}"


def test_legacy_single_row_label_matrix_unchanged():
    """Sanity: classic 'col 0 = label, cols 1+ = series' tables (no
    explicit RowField cols in selectedColumns) keep their legacy
    behavior — cols 1+ get series[col-1]."""
    shape, _, _ = _new_table(3, 3)
    shape.table.cell(0, 0).text = "Topic"
    shape.table.cell(0, 1).text = "CARD"
    shape.table.cell(0, 2).text = "ONC"
    shape.table.cell(1, 0).text = "Efficacy"
    shape.table.cell(2, 0).text = "Safety"

    pivot = {
        "RowFields": ["y_label"],
        "ColumnFields": ["segment_1"],
        "ValueFields": ["decimal"],
        "columnDefinitions": [
            {"Name": "y_label", "IsDefaultAlias": True},
            {"Name": "CARD @:@ Average of decimal",
             "Format": "0%", "IsDefaultAlias": True},
            {"Name": "ONC @:@ Average of decimal",
             "Format": "0%", "IsDefaultAlias": True},
        ],
    }
    # Note: selectedColumns has only the value cols, NOT the row label —
    # the legacy "col 0 implicit row label" pattern.
    mapping = {
        "selectedColumns": [
            "CARD @:@ Average of decimal",
            "ONC @:@ Average of decimal",
        ],
        "selectAllRows": True,
    }
    series = [
        ("CARD @:@ Average of decimal", [0.42, 0.58]),
        ("ONC @:@ Average of decimal", [0.55, 0.61]),
    ]

    cv = _auto_compute_cell_values(shape, ["Efficacy", "Safety"], series, pivot, mapping)
    assert cv is not None
    assert cv[0] == ["Topic", "CARD", "ONC"]
    # cols 1, 2 fall through to positional fallback in col_actions:
    # selected[1]/selected[2] map to selected_visible entries → series[0]/[1]
    assert cv[1][1] == "42%"
    assert cv[1][2] == "55%"
    assert cv[2][1] == "58%"
    assert cv[2][2] == "61%"


def test_mixed_row_field_and_value_selected_columns_still_writes():
    """Sanity: when selectedColumns mixes a RowField AND a value column,
    the value column must still be written (the row-field-only guard
    must not over-fire)."""
    shape, _, _ = _new_table(3, 2)
    shape.table.cell(0, 0).text = "Region"
    shape.table.cell(0, 1).text = "%"
    shape.table.cell(1, 0).text = "NA"
    shape.table.cell(1, 1).text = "?"

    pivot = {
        "RowFields": ["region"],
        "ColumnFields": [],
        "ValueFields": ["decimal"],
        "columnDefinitions": [
            {"Name": "region", "IsDefaultAlias": True},
            {"Name": "Sum of decimal", "Format": "0%", "IsDefaultAlias": True},
        ],
    }
    mapping = {"selectedColumns": ["region", "Sum of decimal"], "selectAllRows": False}
    series = [("Sum of decimal", [0.42, 0.58])]
    categories = ["NA", "EU"]

    cv = _auto_compute_cell_values(shape, categories, series, pivot, mapping)
    # NOT None — the value column should still be written
    assert cv is not None
    # Layout 3 (matrix): row 0 header preserved, row 1 col 0 preserved,
    # row 1 col 1 = formatted series[0][0]
    assert cv[1][1] == "42%"


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
