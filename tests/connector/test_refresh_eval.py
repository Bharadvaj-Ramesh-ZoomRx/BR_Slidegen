"""Unit tests for the post-refresh quality eval (slidegen.refresh_eval).

The eval reads a refresh_status.json sidecar and produces a verdict on
what % of components refreshed cleanly vs need user review vs failed.
Tests cover:
  - status-level classification (passed, preserved, review, error)
  - note-level review downgrade (tag_mismatch / selectedColumns_drift
    on a component whose status is "ok")
  - rate calculations (refreshable denominator excludes preserved)
  - empty / missing data
  - resolve_status_path helper
"""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from slidegen.refresh_eval import (
    ERROR,
    PASSED,
    PRESERVED,
    REVIEW,
    REVIEW_NOTE_KINDS,
    eval_refresh_status,
    format_report,
)


def _write_status(tmp_path, slides):
    p = tmp_path / "status.json"
    p.write_text(json.dumps({"slides": slides}), encoding="utf-8")
    return p


def test_all_passed(tmp_path):
    p = _write_status(tmp_path, [
        {"slide_index": 0, "charts": [
            {"name": "C1", "status": "ok"},
            {"name": "C2", "status": "ok"},
        ], "tables": []},
    ])
    r = eval_refresh_status(p)
    assert r["total_components"] == 2
    assert r["passed"] == 2
    assert r["pass_rate_pct"] == 100.0
    assert r["review"] == 0


def test_static_preserved_excluded_from_denominator(tmp_path):
    """static-pinned components don't count against pass rate — the
    denominator is "refreshable" (total - preserved)."""
    p = _write_status(tmp_path, [
        {"slide_index": 0, "charts": [
            {"name": "C1", "status": "ok"},
        ], "tables": [
            {"name": "T1", "status": "static_pinned_skipped"},
            {"name": "T2", "status": "static_pinned_skipped"},
        ]},
    ])
    r = eval_refresh_status(p)
    assert r["passed"] == 1
    assert r["preserved"] == 2
    # 1 passed of 1 refreshable -> 100%
    assert r["pass_rate_pct"] == 100.0
    # Inclusive pass rate counts preserved as fine: 3 / 3 = 100%
    assert r["inclusive_pass_rate_pct"] == 100.0


def test_review_status_classified(tmp_path):
    p = _write_status(tmp_path, [
        {"slide_index": 5, "charts": [
            {"name": "C1", "status": "alignment_failed"},
            {"name": "C2", "status": "tag_mismatch"},
            {"name": "C3", "status": "ambiguous_tag"},
        ], "tables": []},
    ])
    r = eval_refresh_status(p)
    assert r["review"] == 3
    assert r["passed"] == 0
    # All 3 review kinds appear in the breakdown
    assert "alignment_failed" in r["by_review_kind"]
    assert "tag_mismatch" in r["by_review_kind"]
    assert "ambiguous_tag" in r["by_review_kind"]


def test_note_level_downgrade_tag_mismatch(tmp_path):
    """status='ok' with a tag_mismatch note -> classified REVIEW, not
    PASSED. This is the Repatha ATU slide 73 / slide 47 case."""
    p = _write_status(tmp_path, [
        {"slide_index": 46, "charts": [
            {"name": "Chart 46", "status": "ok",
             "notes": [{"kind": "tag_mismatch",
                        "detail": "fetch had no overlap with source"}]},
        ], "tables": []},
    ])
    r = eval_refresh_status(p)
    assert r["passed"] == 0
    assert r["review"] == 1
    assert r["by_status"].get("tag_mismatch") == 1
    assert 46 in r["by_review_kind"]["tag_mismatch"]


def test_note_level_downgrade_columns_drift(tmp_path):
    """selectedColumns_drift note -> REVIEW classification."""
    p = _write_status(tmp_path, [
        {"slide_index": 0, "charts": [
            {"name": "C1", "status": "ok",
             "notes": [{"kind": "selectedColumns_drift",
                        "detail": "filter intent partially survived"}]},
        ], "tables": []},
    ])
    r = eval_refresh_status(p)
    assert r["review"] == 1
    assert r["passed"] == 0


def test_dynamic_added_counts_as_passed(tmp_path):
    """ok_with_dynamic_added is a clean refresh — extra waves flowed in
    per the dynamic tag, by user contract."""
    p = _write_status(tmp_path, [
        {"slide_index": 0, "charts": [
            {"name": "C1", "status": "ok_with_dynamic_added"},
        ], "tables": []},
    ])
    r = eval_refresh_status(p)
    assert r["passed"] == 1
    assert r["pass_rate_pct"] == 100.0


def test_error_statuses_classified(tmp_path):
    p = _write_status(tmp_path, [
        {"slide_index": 0, "charts": [
            {"name": "C1", "status": "not_found"},
            {"name": "C2", "status": "empty"},
            {"name": "C3", "status": "error"},
        ], "tables": []},
    ])
    r = eval_refresh_status(p)
    assert r["error"] == 3
    assert r["passed"] == 0


def test_empty_status(tmp_path):
    p = _write_status(tmp_path, [])
    r = eval_refresh_status(p)
    assert r["total_components"] == 0
    assert r["pass_rate_pct"] == 0.0
    assert r["review_rate_pct"] == 0.0


def test_mixed_real_world(tmp_path):
    """Mirrors the Repatha ATU v10 distribution: most pass, some review,
    some error."""
    p = _write_status(tmp_path, [
        {"slide_index": i, "charts": [
            {"name": f"C{i}", "status": "ok"},
        ], "tables": []}
        for i in range(8)
    ] + [
        {"slide_index": 17, "charts": [
            {"name": "Overall Reach", "status": "alignment_failed",
             "error": "every value None"},
        ], "tables": []},
        {"slide_index": 46, "charts": [
            {"name": "Chart 46", "status": "ok",
             "notes": [{"kind": "tag_mismatch", "detail": "..."}]},
        ], "tables": []},
    ])
    r = eval_refresh_status(p)
    assert r["total_components"] == 10
    assert r["passed"] == 8
    assert r["review"] == 2
    assert r["pass_rate_pct"] == 80.0


def test_format_report_renders_review_kinds(tmp_path):
    """format_report outputs a human-readable string with review
    breakdown — used by the CLI and pipeline summary."""
    p = _write_status(tmp_path, [
        {"slide_index": 17, "charts": [
            {"name": "Overall Reach", "status": "alignment_failed"},
        ], "tables": []},
    ])
    r = eval_refresh_status(p)
    out = format_report(r)
    assert "alignment_failed" in out
    # Slide 18 is 1-based (0-indexed slide_index=17)
    assert "18" in out


def test_constants_no_overlap():
    """Sanity: the four classification sets don't accidentally overlap."""
    sets = [PASSED, PRESERVED, REVIEW, ERROR, REVIEW_NOTE_KINDS]
    for i, s1 in enumerate(sets):
        for s2 in sets[i + 1:]:
            # REVIEW and REVIEW_NOTE_KINDS share tag_mismatch by design
            # (it can be either a status or a note); skip that pair.
            if s1 is REVIEW and s2 is REVIEW_NOTE_KINDS:
                continue
            assert s1.isdisjoint(s2), \
                f"sets overlap: {s1} vs {s2}"
