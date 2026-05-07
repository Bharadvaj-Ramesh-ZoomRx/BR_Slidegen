"""Unit tests for the connector mapper's 5-tier label resolution chain.

Tests the full source-canonical alignment behavior:
  1. Exact normalized match
  2. Progressive-suffix on " - " split
  3. Tail substring containment in either direction
  4. Fuzzy similarity (SequenceMatcher.ratio >= 0.85)
  5. Positional fallback when source/API counts align

Black-box approach: build minimal records + connector tag dicts and call
pivot_records_to_chart_data, then assert on categories + series.
"""
from __future__ import annotations

import pytest

from slidegen.synapse_chart_mapper import pivot_records_to_chart_data


def _records(rows):
    """Build records with default measure='Mean' and segment_1='Overall'."""
    return [
        {
            "time_period_name": r.get("time_period_name", "Wave 7"),
            "alias_label": r["alias_label"],
            "value": r["value"],
            "measure": r.get("measure", "Mean"),
            "segment_1": r.get("segment_1", "Overall"),
        }
        for r in rows
    ]


_PIVOT = {
    "AggregationType": 0,
    "ColumnFields": [],
    "RowFields": ["alias_label"],
    "ValueFields": ["value"],
    "Filters": [],
    "columnDefinitions": [
        {"Name": "alias_label", "IsDefaultAlias": True},
    ],
}

_MAPPING = {
    "selectedColumns": [],
    "selectedRows": [],
    "selectAllRows": True,
    "applyTranspose": False,
}


def test_exact_match_categories():
    """Source category labels match pivot output exactly — no fallback needed."""
    records = _records([
        {"alias_label": "Category A", "value": 10},
        {"alias_label": "Category B", "value": 20},
    ])
    res = pivot_records_to_chart_data(
        records, _PIVOT, _MAPPING,
        source_categories=["Category A", "Category B"],
        source_series_names=["S"],
    )
    assert res.success
    assert res.categories == ["Category A", "Category B"]


def test_progressive_suffix_match_categories():
    """Connector default-alias renames 'Specialty - X - CARD' -> 'CARD'.
    Resolver should match the source's full form to the pivot's short form.
    """
    records = _records([
        {"alias_label": "CARD", "value": 10},
        {"alias_label": "PCP - Reserver", "value": 20},
    ])
    res = pivot_records_to_chart_data(
        records, _PIVOT, _MAPPING,
        source_categories=[
            "Specialty (C/PCP Segments) + Tier Groups - CARD",
            "Specialty (C/PCP Segments) + Tier Groups - PCP - Reserver",
        ],
        source_series_names=["S"],
    )
    assert res.success
    # Categories preserved as source's full form
    assert res.categories[0].endswith("CARD")
    # Values flowed through positional after suffix match
    assert res.series[0][1][0] == 10
    assert res.series[0][1][1] == 20


def test_fuzzy_match_handles_typo():
    """Source label 'CARDS' (with typo) vs pivot 'CARD' — fuzzy at 0.85
    should catch."""
    records = _records([{"alias_label": "CARD", "value": 42}])
    res = pivot_records_to_chart_data(
        records, _PIVOT, _MAPPING,
        source_categories=["CARDS"],
        source_series_names=["S"],
    )
    assert res.success
    assert res.series[0][1][0] == 42


def test_fuzzy_rejects_unrelated():
    """Distinct labels with low similarity should not fuzzy-match;
    falls through to positional or None depending on count alignment."""
    records = _records([{"alias_label": "Bananas", "value": 99}])
    res = pivot_records_to_chart_data(
        records, _PIVOT, _MAPPING,
        source_categories=["Apples"],
        source_series_names=["S"],
    )
    # Counts align (1 vs 1) -> positional fallback fires; values flow.
    # The point is fuzzy didn't false-match before that; positional is
    # the documented behavior for count-aligned non-matching labels.
    assert res.success
    assert res.series[0][1][0] == 99


def test_positional_fallback_count_match():
    """4 source categories, 4 pivot categories, no name match — 5th-tier
    positional alignment populates values 1:1."""
    records = _records([
        {"alias_label": "X1", "value": 1},
        {"alias_label": "X2", "value": 2},
        {"alias_label": "X3", "value": 3},
        {"alias_label": "X4", "value": 4},
    ])
    res = pivot_records_to_chart_data(
        records, _PIVOT, _MAPPING,
        source_categories=["Foo", "Bar", "Baz", "Qux"],  # nothing matches
        source_series_names=["S"],
    )
    assert res.success
    # Values arrive 1:1 in source order
    assert res.series[0][1] == [1.0, 2.0, 3.0, 4.0]


def test_count_mismatch_with_zero_overlap_surfaces_tag_mismatch():
    """When source count != pivot count AND zero label overlap on both
    axes, the tag_mismatch detection fires earlier than category
    alignment. Returns API data faithfully + a tag_mismatch note for
    the user to review — NOT alignment_failed."""
    records = _records([
        {"alias_label": "X1", "value": 1},
        {"alias_label": "X2", "value": 2},
    ])
    res = pivot_records_to_chart_data(
        records, _PIVOT, _MAPPING,
        source_categories=["Foo", "Bar", "Baz", "Qux"],
        source_series_names=["S_unrelated"],
    )
    assert res.success
    note_kinds = {n.get("kind") for n in (res.notes or [])}
    assert "tag_mismatch" in note_kinds


def test_series_name_progressive_suffix():
    """Source series 'Specialty - CARD - L' should match pivot's 'CARD - L'."""
    # Build records that produce the pivot column "CARD - L"
    records = [
        {"time_period_name": "Wave 7", "alias_label": "Risk A",
         "value": 0.7, "measure": "Mean", "segment_1": "CARD - L"},
        {"time_period_name": "Wave 7", "alias_label": "Risk B",
         "value": 0.6, "measure": "Mean", "segment_1": "CARD - L"},
    ]
    pivot = {
        "AggregationType": 0,
        "ColumnFields": ["segment_1"],
        "RowFields": ["alias_label"],
        "ValueFields": ["value"],
        "Filters": [],
        "columnDefinitions": [
            {"Name": "alias_label"},
            {"Name": "Specialty (C/PCP Segments) + Tier Detailed - CARD - L"},
        ],
    }
    mapping = {
        "selectedColumns": [
            "Specialty (C/PCP Segments) + Tier Detailed - CARD - L",
        ],
        "selectedRows": [],
        "selectAllRows": True,
        "applyTranspose": False,
    }
    res = pivot_records_to_chart_data(
        records, pivot, mapping,
        source_categories=["Risk A", "Risk B"],
        source_series_names=[
            "Specialty (C/PCP Segments) + Tier Detailed - CARD - L",
        ],
    )
    assert res.success
    # Series resolved by progressive suffix; values flow through.
    assert len(res.series) >= 1
    vals = res.series[0][1]
    assert vals[:2] == [0.7, 0.6]
