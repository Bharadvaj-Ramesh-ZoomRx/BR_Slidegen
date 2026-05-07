"""Unit tests for the connector mapper's Filters resolver and the
latest-wave-only filter.

Tests via pivot_records_to_chart_data with synthetic records:
  - ;-joined IN filters with EXACT matches
  - ;-joined IN filters with compound-suffix matches (long form in tag,
    short form in API)
  - ;-joined IN filters where no value matches (df left untouched)
  - measure-column filter via Filters list (Mean only)
  - latest-wave-only filter when time_period_name not in ColumnFields
"""
from __future__ import annotations

from slidegen.synapse_chart_mapper import pivot_records_to_chart_data


def _records(rows):
    """Helper: build records with sensible defaults."""
    out = []
    for r in rows:
        out.append({
            "time_period_name": r.get("time_period_name", "Wave 7"),
            "alias_label": r.get("alias_label", "X"),
            "value": r.get("value", 0),
            "measure": r.get("measure", "Mean"),
            "segment_1": r.get("segment_1", "Overall"),
        })
    return out


_PIVOT_BASE = {
    "AggregationType": 1,  # mean
    "ColumnFields": ["segment_1"],
    "RowFields": ["alias_label"],
    "ValueFields": ["value"],
    "Filters": [],
    "columnDefinitions": [
        {"Name": "alias_label"},
    ],
}

_MAPPING_PASS = {
    "selectedColumns": [],
    "selectedRows": [],
    "selectAllRows": True,
    "applyTranspose": False,
}


# ─────────────────────────────────────────────────────────────────────
# ;-joined IN filter
# ─────────────────────────────────────────────────────────────────────

def test_semicolon_filter_exact_matches():
    """When ;-joined values match df values exactly, isin filtering keeps
    only the listed values."""
    records = _records([
        {"alias_label": "Risk1", "segment_1": "CARD - L", "value": 0.7},
        {"alias_label": "Risk1", "segment_1": "CARD - M", "value": 0.8},
        {"alias_label": "Risk1", "segment_1": "PCP - Believer", "value": 0.5},
    ])
    pivot = dict(_PIVOT_BASE)
    pivot["Filters"] = [{
        "ColumnKey": "segment_1",
        "Type": 1,
        "Value": "CARD - L;CARD - M",
        "filterCriteria": 1,
    }]
    res = pivot_records_to_chart_data(records, pivot, _MAPPING_PASS)
    # Only the 2 matching segments should be present
    series_names = {n for n, _ in res.series}
    assert any("CARD - L" in n for n in series_names)
    assert any("CARD - M" in n for n in series_names)
    assert not any("Believer" in n for n in series_names)


def test_semicolon_filter_compound_suffix_match():
    """Repatha ATU L-table case: filter holds full-form long names
    delimited by ';' but API has just the short tail ('CARD - L'). The
    new ;-resolver should match via compound-suffix tier."""
    records = _records([
        {"alias_label": "Risk1", "segment_1": "CARD - L", "value": 0.7},
        {"alias_label": "Risk1", "segment_1": "CARD - M", "value": 0.8},
        {"alias_label": "Risk1", "segment_1": "PCP - Reserver", "value": 0.5},
    ])
    pivot = dict(_PIVOT_BASE)
    pivot["Filters"] = [{
        "ColumnKey": "segment_1",
        "Type": 1,
        "Value": (
            "Specialty (C/PCP Segments) + Tier Detailed - CARD - L;"
            "Specialty (C/PCP Segments) + Tier Detailed - CARD - M"
        ),
        "filterCriteria": 1,
    }]
    res = pivot_records_to_chart_data(records, pivot, _MAPPING_PASS)
    series_names = {n for n, _ in res.series}
    # Both CARD tiers kept (matched via compound-suffix); Reserver filtered out.
    assert any("CARD - L" in n for n in series_names)
    assert any("CARD - M" in n for n in series_names)
    assert not any("Reserver" in n for n in series_names)


def test_semicolon_filter_no_matches_leaves_df_intact():
    """When NONE of the ;-joined values match (broken tag), the resolver
    should NOT wipe df. Records remain so the rest of the pipeline can
    continue (alignment_failed / tag_mismatch will catch it later)."""
    records = _records([
        {"alias_label": "Risk1", "segment_1": "CARD - L", "value": 0.7},
        {"alias_label": "Risk1", "segment_1": "PCP - Reserver", "value": 0.5},
    ])
    pivot = dict(_PIVOT_BASE)
    pivot["Filters"] = [{
        "ColumnKey": "segment_1",
        "Type": 1,
        "Value": "Nonexistent Segment 1;Nonexistent Segment 2",
        "filterCriteria": 1,
    }]
    res = pivot_records_to_chart_data(records, pivot, _MAPPING_PASS)
    # Both original segments should still appear in the pivot output —
    # the broken filter didn't wipe them.
    series_names = {n for n, _ in res.series}
    assert any("CARD - L" in n or "Reserver" in n for n in series_names)


# ─────────────────────────────────────────────────────────────────────
# Single-value filter (measure column)
# ─────────────────────────────────────────────────────────────────────

def test_measure_filter_keeps_only_mean_rows():
    """Filter measure='Mean' against rows with mixed measures."""
    records = _records([
        {"alias_label": "Risk1", "value": 0.73, "measure": "Mean"},
        {"alias_label": "Risk1", "value": 100, "measure": "Sum"},
        {"alias_label": "Risk1", "value": 5, "measure": "Count"},
    ])
    pivot = dict(_PIVOT_BASE)
    pivot["Filters"] = [{
        "ColumnKey": "measure", "Type": 1, "Value": "Mean", "filterCriteria": 1,
    }]
    res = pivot_records_to_chart_data(records, pivot, _MAPPING_PASS)
    # Only the Mean row's value (0.73) should be aggregated, not the
    # Sum / Count values.
    assert res.success
    # Series value should be 0.73 (mean of 1 row), not the average of
    # all 3 rows.
    vals = res.series[0][1] if res.series else []
    assert vals and abs(vals[0] - 0.73) < 0.01


# ─────────────────────────────────────────────────────────────────────
# Latest-wave-only filter
# ─────────────────────────────────────────────────────────────────────

def test_latest_wave_only_when_time_not_in_column_fields():
    """When df has multiple waves but ColumnFields doesn't include
    time_period_name, the mapper should keep ONLY the latest wave's
    rows (otherwise sum across 2 waves of Mean values produces 148%)."""
    records = _records([
        {"alias_label": "Risk1", "value": 0.7, "time_period_name": "Project Wave 6"},
        {"alias_label": "Risk1", "value": 0.74, "time_period_name": "Project Wave 7"},
    ])
    pivot = dict(_PIVOT_BASE)
    # ColumnFields is segment_1 (default), NOT time_period_name -> latest filter active
    res = pivot_records_to_chart_data(records, pivot, _MAPPING_PASS)
    assert res.success
    # Only Wave 7's value (0.74) should survive — not the sum (1.44).
    vals = res.series[0][1] if res.series else []
    assert vals and abs(vals[0] - 0.74) < 0.01


def test_no_latest_filter_when_time_in_column_fields():
    """When time_period_name IS in ColumnFields, multiple waves stay as
    separate columns (the chart shows them as separate series)."""
    records = _records([
        {"alias_label": "Risk1", "value": 0.7, "time_period_name": "Wave 6"},
        {"alias_label": "Risk1", "value": 0.74, "time_period_name": "Wave 7"},
    ])
    pivot = dict(_PIVOT_BASE)
    pivot["ColumnFields"] = ["time_period_name"]
    res = pivot_records_to_chart_data(records, pivot, _MAPPING_PASS)
    assert res.success
    # Both waves should be separate series
    series_names = {n for n, _ in res.series}
    assert len(series_names) == 2  # Wave 6 + Wave 7
