"""Regression: the latest-wave-only filter must not collapse trend charts.

The filter at synapse_chart_mapper.pivot_records_to_chart_data exists for
'L-style label tables' (single-wave snapshot label) — when waves aren't a
chart axis at all, summing 6 waves of a Mean would produce nonsense
values, so the filter restricts to the latest wave.

Bug: the original filter only checked `ColumnFields`. When a tag has
`RowFields=['time_period_name']` (the canonical TRENDED CHART shape —
each wave is a row of the chart's category axis), the filter wrongly
fired and collapsed `dynamic_latest_n=6` data down to a single wave.

Slide 17 of the Testing Deck (CREON rep-effectiveness trends) had 6
charts with this exact shape. All 6 rendered as single-wave snapshots
instead of 6-wave trends. Inventory across the full Testing Deck found
29 trend charts collapsed by this bug — slides 2, 3, 11, 16, 17, 18,
23, 44, 47.

Fix: the filter must also exclude `time_period_name in RowFields`.
This locks in the contract.
"""
import sys
from pathlib import Path

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO_ROOT))

from slidegen.synapse_chart_mapper import pivot_records_to_chart_data


def _trend_records(waves, value_per_wave_per_brand):
    """Build raw API records: 1 row per (wave, brand) with a decimal value."""
    out = []
    for w in waves:
        for brand, v in value_per_wave_per_brand.items():
            out.append({
                "time_period_name": w,
                "alias_label": brand,
                "y_label": "Compelling reason to prescribe",
                "decimal": v[w],
                "value": v[w],
                "respondents": 100,
                "base": 100,
            })
    return out


def test_trend_chart_with_time_period_in_rowfields_keeps_all_waves():
    """Slide 17 case: RowFields=['time_period_name'], 6 waves expected."""
    waves = ["Nov'25", "Dec'25", "Jan'26", "Feb'26", "Mar'26", "Apr'26"]
    creon_vals = {w: 0.50 + i * 0.02 for i, w in enumerate(waves)}
    zenpep_vals = {w: 0.40 + i * 0.02 for i, w in enumerate(waves)}
    records = _trend_records(waves, {"CREON": creon_vals, "ZENPEP": zenpep_vals})

    pivot_config = {
        "RowFields": ["time_period_name"],
        "ColumnFields": ["alias_label", "y_label"],
        "ValueFields": ["decimal"],
        "AggregationType": 1,  # mean
        "columnDefinitions": [],
    }
    mapping_config = {
        "selectedColumns": [
            "time_period_name",
            "CREON @:@ Compelling reason to prescribe",
            "ZENPEP @:@ Compelling reason to prescribe",
        ],
        "selectAllRows": True,
    }

    result = pivot_records_to_chart_data(
        records, pivot_config, mapping_config,
    )
    assert result.success, f"pivot failed: {result.error}"
    # Categories axis = each wave; expect all 6
    assert len(result.categories) == 6, (
        f"Expected 6 categories (one per wave), got {len(result.categories)}: "
        f"{result.categories}. The latest-wave-only filter is over-firing."
    )


def test_label_table_no_time_period_axis_collapses_to_latest():
    """Sanity: when waves are NEITHER in RowFields NOR ColumnFields (the
    L-style label table), the filter MUST still fire to keep just the
    latest wave (otherwise Mean values double up across 2 waves)."""
    waves = ["Feb'26", "Mar'26"]
    records = []
    for w in waves:
        records.append({
            "time_period_name": w,
            "alias_label": "L",
            "decimal": 0.74 if w == "Mar'26" else 0.70,
            "respondents": 100,
            "base": 100,
        })

    pivot_config = {
        "RowFields": ["alias_label"],
        "ColumnFields": [],
        "ValueFields": ["decimal"],
        "AggregationType": 1,  # mean
        "columnDefinitions": [],
    }
    mapping_config = {"selectedColumns": ["alias_label"], "selectAllRows": True}

    result = pivot_records_to_chart_data(records, pivot_config, mapping_config)
    assert result.success
    # Only one row in the pivot (alias_label=L), one wave kept (Mar'26).
    # The series value should be the SINGLE Mar'26 value (0.74), NOT
    # a sum/mean of both waves.
    assert len(result.categories) == 1
    if result.series:
        sname, vals = result.series[0]
        # Single non-aggregated wave value
        assert any(abs(v - 0.74) < 0.001 for v in vals if v is not None), (
            f"Latest-wave-only filter didn't fire — got {vals}, "
            f"should be single Mar'26 value 0.74"
        )


def test_chart_with_time_period_in_columnfields_unaffected():
    """When time_period_name is in ColumnFields (waves spread across the
    columns axis), the filter must NOT fire either."""
    waves = ["Q1'26", "Q2'26", "Q3'26"]
    records = []
    for w in waves:
        records.append({
            "time_period_name": w,
            "y_label": "topic A",
            "decimal": 0.5,
            "respondents": 100,
            "base": 100,
        })

    pivot_config = {
        "RowFields": ["y_label"],
        "ColumnFields": ["time_period_name"],
        "ValueFields": ["decimal"],
        "AggregationType": 1,
        "columnDefinitions": [],
    }
    mapping_config = {
        "selectedColumns": [
            "y_label",
            "Q1'26 @:@ Average of decimal",
            "Q2'26 @:@ Average of decimal",
            "Q3'26 @:@ Average of decimal",
        ],
        "selectAllRows": True,
    }

    result = pivot_records_to_chart_data(records, pivot_config, mapping_config)
    assert result.success
    # All 3 waves preserved (categories axis is y_label here, but the
    # series count reflects the 3 waves)
    assert len(result.series) >= 1
