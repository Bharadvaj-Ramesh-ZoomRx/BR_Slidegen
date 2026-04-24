"""
test_propose_pivot_config.py — Unit tests for propose_pivot_config().

Tests the DataFrame-only inference path (no chart context) and the
chart-validated path (with chart_shape from read_slide_context).

Usage:
    python tests/test_propose_pivot_config.py
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO))

from slidegen.intelligent_refresh import propose_pivot_config, propose_raw_configs


# ══════════════════════════════════════════════════════════════════════════════
# Test DataFrames — representative Synapse shapes
# ══════════════════════════════════════════════════════════════════════════════

def _creon_stacked_df() -> pd.DataFrame:
    """CREON Practice Setting — column_stacked_100_vertical."""
    return pd.DataFrame([
        {"time_period_name": "Jan'26", "option": "Office-based",   "decimal": 0.45, "analysis_id": 641211, "base": 150},
        {"time_period_name": "Jan'26", "option": "Hospital-based", "decimal": 0.35, "analysis_id": 641211, "base": 150},
        {"time_period_name": "Jan'26", "option": "Academic",       "decimal": 0.20, "analysis_id": 641211, "base": 150},
        {"time_period_name": "Feb'26", "option": "Office-based",   "decimal": 0.50, "analysis_id": 641211, "base": 160},
        {"time_period_name": "Feb'26", "option": "Hospital-based", "decimal": 0.30, "analysis_id": 641211, "base": 160},
        {"time_period_name": "Feb'26", "option": "Academic",       "decimal": 0.20, "analysis_id": 641211, "base": 160},
        {"time_period_name": "Mar'26", "option": "Office-based",   "decimal": 0.48, "analysis_id": 641211, "base": 155},
        {"time_period_name": "Mar'26", "option": "Hospital-based", "decimal": 0.32, "analysis_id": 641211, "base": 155},
        {"time_period_name": "Mar'26", "option": "Academic",       "decimal": 0.20, "analysis_id": 641211, "base": 155},
    ])


def _repatha_atu_df() -> pd.DataFrame:
    """Repatha ATU — bar_stacked_100_horizontal with L/N/H/0 series."""
    return pd.DataFrame([
        {"y_label": "Repatha",  "measure": "Low",         "percentage": 0.37, "base": 91,  "time_period_name": "Q1'26"},
        {"y_label": "Repatha",  "measure": "Neutral",     "percentage": 0.22, "base": 91,  "time_period_name": "Q1'26"},
        {"y_label": "Repatha",  "measure": "High",        "percentage": 0.30, "base": 91,  "time_period_name": "Q1'26"},
        {"y_label": "Repatha",  "measure": "Don't Know",  "percentage": 0.11, "base": 91,  "time_period_name": "Q1'26"},
        {"y_label": "Praluent", "measure": "Low",         "percentage": 0.40, "base": 126, "time_period_name": "Q1'26"},
        {"y_label": "Praluent", "measure": "Neutral",     "percentage": 0.25, "base": 126, "time_period_name": "Q1'26"},
        {"y_label": "Praluent", "measure": "High",        "percentage": 0.20, "base": 126, "time_period_name": "Q1'26"},
        {"y_label": "Praluent", "measure": "Don't Know",  "percentage": 0.15, "base": 126, "time_period_name": "Q1'26"},
    ])


def _segment_df() -> pd.DataFrame:
    """Segment comparison — bar_clustered with segment series."""
    return pd.DataFrame([
        {"y_label": "Awareness",    "segment_1": "Cardiologist", "decimal": 0.80, "time_period_name": "Q1'26"},
        {"y_label": "Awareness",    "segment_1": "PCP",          "decimal": 0.60, "time_period_name": "Q1'26"},
        {"y_label": "Awareness",    "segment_1": "Overall",      "decimal": 0.70, "time_period_name": "Q1'26"},
        {"y_label": "Trial",        "segment_1": "Cardiologist", "decimal": 0.50, "time_period_name": "Q1'26"},
        {"y_label": "Trial",        "segment_1": "PCP",          "decimal": 0.30, "time_period_name": "Q1'26"},
        {"y_label": "Trial",        "segment_1": "Overall",      "decimal": 0.40, "time_period_name": "Q1'26"},
        {"y_label": "Adoption",     "segment_1": "Cardiologist", "decimal": 0.35, "time_period_name": "Q1'26"},
        {"y_label": "Adoption",     "segment_1": "PCP",          "decimal": 0.15, "time_period_name": "Q1'26"},
        {"y_label": "Adoption",     "segment_1": "Overall",      "decimal": 0.25, "time_period_name": "Q1'26"},
    ])


# ══════════════════════════════════════════════════════════════════════════════
# Tests
# ══════════════════════════════════════════════════════════════════════════════

def test_creon_no_chart():
    """CREON stacked — infer from DataFrame alone."""
    df = _creon_stacked_df()
    result = propose_pivot_config(df)

    assert result["error"] is None, f"Unexpected error: {result['error']}"
    assert result["row_field"] == "time_period_name", f"row_field: {result['row_field']}"
    assert result["series_column"] == "option", f"series_column: {result['series_column']}"
    assert result["value_field"] == "decimal", f"value_field: {result['value_field']}"
    assert result["val_format"] == "0%"
    print(f"  PASS  CREON no-chart: row={result['row_field']}, series={result['series_column']}, "
          f"val={result['value_field']}, confidence={result['confidence']}")


def test_creon_with_chart():
    """CREON stacked — infer with chart context for validation."""
    df = _creon_stacked_df()
    chart_shape = {
        "categories": ["Jan'26", "Feb'26", "Mar'26"],
        "series_names": ["Office-based", "Hospital-based", "Academic"],
        "series_values": {"Office-based": [0.45, 0.50, 0.48]},
    }
    result = propose_pivot_config(df, chart_shape)

    assert result["error"] is None
    assert result["row_field"] == "time_period_name"
    assert result["series_column"] == "option"
    assert result["value_field"] == "decimal"
    assert result["confidence"] >= 0.8, f"Expected high confidence, got {result['confidence']}"
    print(f"  PASS  CREON with-chart: confidence={result['confidence']}, "
          f"row={result['derivation']['row_field'][1]}")


def test_repatha_no_chart():
    """Repatha ATU — y_label as row, measure as series."""
    df = _repatha_atu_df()
    result = propose_pivot_config(df)

    assert result["error"] is None
    assert result["row_field"] == "y_label", f"row_field: {result['row_field']}"
    assert result["series_column"] == "measure", f"series_column: {result['series_column']}"
    assert result["value_field"] == "percentage", f"value_field: {result['value_field']}"
    print(f"  PASS  Repatha no-chart: row={result['row_field']}, series={result['series_column']}, "
          f"val={result['value_field']}, confidence={result['confidence']}")


def test_segment_no_chart():
    """Segment comparison — y_label as row, segment_1 as series."""
    df = _segment_df()
    result = propose_pivot_config(df)

    assert result["error"] is None
    assert result["row_field"] == "y_label", f"row_field: {result['row_field']}"
    assert result["series_column"] == "segment_1", f"series_column: {result['series_column']}"
    assert result["value_field"] == "decimal", f"value_field: {result['value_field']}"
    print(f"  PASS  Segment no-chart: row={result['row_field']}, series={result['series_column']}, "
          f"val={result['value_field']}, confidence={result['confidence']}")


def test_propose_raw_configs_no_chart():
    """propose_raw_configs with chart_shape=None — pure DataFrame inference."""
    df = _creon_stacked_df()
    result = propose_raw_configs(None, df)

    assert result["error"] is None, f"Unexpected error: {result['error']}"
    pivot = result["raw_pivot_config"]
    assert pivot["RowFields"] == ["time_period_name"]
    assert pivot["ColumnFields"] == ["option"]
    assert pivot["ValueFields"] == ["decimal"]
    assert result["confidence"] > 0
    print(f"  PASS  raw_configs no-chart: RowFields={pivot['RowFields']}, "
          f"ColumnFields={pivot['ColumnFields']}, confidence={result['confidence']}")


def test_propose_raw_configs_with_chart():
    """propose_raw_configs with chart_shape — backward compat."""
    df = _creon_stacked_df()
    chart_shape = {
        "categories": ["Jan'26", "Feb'26", "Mar'26"],
        "series_names": ["Office-based", "Hospital-based", "Academic"],
        "series_values": {},
    }
    result = propose_raw_configs(chart_shape, df)

    assert result["error"] is None
    pivot = result["raw_pivot_config"]
    assert pivot["RowFields"] == ["time_period_name"]
    assert pivot["ColumnFields"] == ["option"]
    # Check series are sorted alphabetically (Connector convention) and row_field included (chart mode)
    mapping = result["raw_mapping_config"]
    assert mapping["selectedColumns"][0] == "time_period_name"
    assert sorted(mapping["selectedColumns"][1:]) == mapping["selectedColumns"][1:]
    assert set(mapping["selectedColumns"][1:]) == {"Academic", "Hospital-based", "Office-based"}
    print(f"  PASS  raw_configs with-chart: selectedColumns sorted, row_field included")


def test_empty_df():
    """Empty DataFrame should return error, not crash."""
    result = propose_pivot_config(pd.DataFrame())
    assert result.get("error") is not None
    print(f"  PASS  empty df: {result['error']}")


def test_column_roles_exposed():
    """Verify column_roles is returned for debugging."""
    df = _creon_stacked_df()
    result = propose_pivot_config(df)
    roles = result["column_roles"]

    assert "temporal" in roles
    assert "categorical" in roles
    assert "numeric" in roles
    assert "identifier" in roles
    # analysis_id should be classified as identifier
    id_cols = roles["identifier"]
    assert "analysis_id" in id_cols, f"analysis_id not in identifiers: {id_cols}"
    print(f"  PASS  column_roles: temporal={[c for c,_ in roles['temporal']]}, "
          f"identifiers={roles['identifier']}")


def main():
    print("=" * 60)
    print("propose_pivot_config() unit tests")
    print("=" * 60)

    tests = [
        test_creon_no_chart,
        test_creon_with_chart,
        test_repatha_no_chart,
        test_segment_no_chart,
        test_propose_raw_configs_no_chart,
        test_propose_raw_configs_with_chart,
        test_empty_df,
        test_column_roles_exposed,
    ]

    passed = 0
    failed = 0
    for t in tests:
        try:
            t()
            passed += 1
        except (AssertionError, Exception) as e:
            print(f"  FAIL  {t.__name__}: {e}")
            failed += 1

    print(f"\n{'=' * 60}")
    print(f"Results: {passed} passed, {failed} failed")
    print(f"{'=' * 60}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
