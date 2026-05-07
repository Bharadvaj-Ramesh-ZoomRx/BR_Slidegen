"""Unit tests for the connector-tag formula-column evaluator.

Covers the small Excel-formula language the mapper supports for derived
columns declared via columnDefinitions[].Formula:

    =BN, =BN/X, =BN*X, =BN+CN, =BN-CN, =BN+X, =BN-X

Plus the column-letter resolver, the per-row evaluator, and the full
_build_formula_columns flow with realistic columnDefinitions including
Connector default-alias renames ("Project Wave 7" -> "Wave 7").
"""
from __future__ import annotations

import pytest

from slidegen.synapse_chart_mapper import (
    _build_formula_columns,
    _col_letter_to_index,
    _eval_formula_for_row,
    _tokenize_formula,
)


# ─────────────────────────────────────────────────────────────────────
# Column-letter resolver
# ─────────────────────────────────────────────────────────────────────

def test_col_letter_single_letters():
    assert _col_letter_to_index("A") == 0
    assert _col_letter_to_index("B") == 1
    assert _col_letter_to_index("Z") == 25


def test_col_letter_two_letters():
    assert _col_letter_to_index("AA") == 26
    assert _col_letter_to_index("AB") == 27
    assert _col_letter_to_index("AZ") == 51
    assert _col_letter_to_index("BA") == 52


# ─────────────────────────────────────────────────────────────────────
# Tokenizer
# ─────────────────────────────────────────────────────────────────────

def test_tokenize_simple_col_ref():
    toks = _tokenize_formula("B2")
    assert toks == [("col", 1)]


def test_tokenize_division_by_constant():
    toks = _tokenize_formula("B2/100")
    assert toks == [("col", 1), ("op", "/"), ("num", 100.0)]


def test_tokenize_multiplication_by_constant():
    toks = _tokenize_formula("B2*100")
    assert toks == [("col", 1), ("op", "*"), ("num", 100.0)]


def test_tokenize_addition_two_cols():
    toks = _tokenize_formula("B2+C2")
    assert toks == [("col", 1), ("op", "+"), ("col", 2)]


def test_tokenize_handles_inline_whitespace():
    """Inline whitespace between tokens — the patterns Connector emits
    in the wild ('=B2 / 100')."""
    toks = _tokenize_formula("B2 / 100")
    assert toks == [("col", 1), ("op", "/"), ("num", 100.0)]


def test_tokenize_invalid_returns_none():
    assert _tokenize_formula("$$$") is None


# ─────────────────────────────────────────────────────────────────────
# Per-row evaluator
# ─────────────────────────────────────────────────────────────────────

def test_eval_simple_col_ref():
    toks = _tokenize_formula("B2")
    assert _eval_formula_for_row(toks, [None, 73]) == 73


def test_eval_divide_by_100():
    toks = _tokenize_formula("B2/100")
    assert _eval_formula_for_row(toks, [None, 73]) == 0.73


def test_eval_multiply_by_100():
    toks = _tokenize_formula("B2*100")
    assert _eval_formula_for_row(toks, [None, 0.73]) == 73.0


def test_eval_two_cols_addition():
    toks = _tokenize_formula("B2+C2")
    assert _eval_formula_for_row(toks, [None, 10, 20]) == 30


def test_eval_two_cols_subtraction():
    toks = _tokenize_formula("B2-C2")
    assert _eval_formula_for_row(toks, [None, 50, 30]) == 20


def test_eval_mixed_precedence():
    """*/ binds tighter than +-."""
    toks = _tokenize_formula("B2*100+5")
    assert _eval_formula_for_row(toks, [None, 0.5]) == 55.0


def test_eval_missing_column_returns_none():
    """Reference to an out-of-range column letter -> None, doesn't crash."""
    toks = _tokenize_formula("Z2")
    assert _eval_formula_for_row(toks, [None, 73]) is None


def test_eval_division_by_zero_returns_none():
    toks = _tokenize_formula("B2/C2")
    assert _eval_formula_for_row(toks, [None, 73, 0]) is None


# ─────────────────────────────────────────────────────────────────────
# Full _build_formula_columns flow
# ─────────────────────────────────────────────────────────────────────

def test_avgpercent_formula_creon_atu():
    """Repatha ATU slide 11 'Overall CARDs' chart.

    columnDefinitions:
      A: alias_label (row label)
      B: Project Wave 7 (data)
      C: <blank:AvgPercent> Formula =B2/100, Format 0%

    Pivot output (after Connector default-alias rename "Project Wave 7"
    -> "Wave 7"):
      series = [("Wave 7", [73, 81, 65])]

    Expected: AvgPercent column derived as [0.73, 0.81, 0.65].
    """
    col_defs = [
        {"Name": "alias_label", "IsDefaultAlias": True},
        {"Name": "Project Wave 7", "IsDefaultAlias": True},
        {"Name": "<blank:AvgPercent>", "Formula": "=B2/100",
         "Alias": "AvgPercent", "Format": "0%"},
    ]
    series = [("Wave 7", [73.0, 81.0, 65.0])]
    out = _build_formula_columns(col_defs, series, ["a", "b", "c"])
    assert out == [("AvgPercent", [0.73, 0.81, 0.65])]


def test_l_table_formula_with_specialty_prefix():
    """Repatha ATU slide 11 'L' label_table.

    columnDefinitions[1] is "Specialty (C/PCP Segments) + Tier Detailed
    - CARD - L" but pivot output names it "CARD - L" after default-alias.
    Resolver should still match via progressive-suffix.
    """
    col_defs = [
        {"Name": "alias_label", "IsDefaultAlias": True},
        {"Name": "Specialty (C/PCP Segments) + Tier Detailed - CARD - L",
         "IsDefaultAlias": True},
        {"Name": "<blank:L>", "Formula": "=B2/100",
         "Alias": "L", "Format": "0%"},
    ]
    series = [("CARD - L", [70.0, 71.0, 66.0, 65.0])]
    out = _build_formula_columns(col_defs, series, ["p1", "p2", "p3", "p4"])
    assert out == [("L", [0.70, 0.71, 0.66, 0.65])]


def test_no_formula_returns_empty():
    """columnDefinitions with no Formula entries -> no derived columns."""
    col_defs = [
        {"Name": "alias_label", "IsDefaultAlias": True},
        {"Name": "Wave 7", "IsDefaultAlias": True},
    ]
    series = [("Wave 7", [50.0, 60.0])]
    assert _build_formula_columns(col_defs, series, ["a", "b"]) == []


def test_formula_referencing_missing_col_returns_zero():
    """When the formula's column letter doesn't resolve to a series in
    the pivot, the row's value is None and gets coerced to 0.0 (we
    don't drop the column — surfaces visibly that something is off)."""
    col_defs = [
        {"Name": "alias_label", "IsDefaultAlias": True},
        {"Name": "<blank:Computed>", "Formula": "=Z2*2",
         "Alias": "Computed"},
    ]
    series = [("Wave 7", [73.0])]
    out = _build_formula_columns(col_defs, series, ["x"])
    assert out == [("Computed", [0.0])]


def test_two_formula_columns_independent():
    """Two derived columns side by side use the same source columns
    independently."""
    col_defs = [
        {"Name": "alias_label"},
        {"Name": "Wave 6"},
        {"Name": "Wave 7"},
        {"Name": "<blank:Pct6>", "Formula": "=B2/100", "Alias": "Pct6"},
        {"Name": "<blank:Delta>", "Formula": "=C2-B2", "Alias": "Delta"},
    ]
    series = [("Wave 6", [70.0, 80.0]), ("Wave 7", [73.0, 85.0])]
    out = _build_formula_columns(col_defs, series, ["a", "b"])
    assert ("Pct6", [0.70, 0.80]) in out
    assert ("Delta", [3.0, 5.0]) in out
