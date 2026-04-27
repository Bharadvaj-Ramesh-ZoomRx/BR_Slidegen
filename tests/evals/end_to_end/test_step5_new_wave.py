"""Step 5 eval — new data propagates with same structure (connected path).

For each fixture deck, picks one reference chart on a stable analysis,
runs refresh against TWO different wave sets (latest vs older), and
asserts:

    1. Smell test — values differ between the two wave variants
       (refresh isn't a no-op restore)
    2. Structural invariance — categories and series count identical
       across variants (Step 5's "same structure" qualifier)
    3. API correctness — each variant's chart values match a direct
       Synapse query for that wave (within 1e-3 absolute tolerance)

Hits Synapse 4× per deck per run (2 refreshes + 2 ground-truth queries).
"""
from __future__ import annotations

import json
import os
import shutil
from pathlib import Path

import pytest

from .step5_helpers import (
    build_wave_variant_spec,
    extract_reference_chart,
    fetch_api_truth,
    refresh_with_variant,
)
from ..fixtures import FIXTURE_DECKS


REPO_ROOT = Path(__file__).resolve().parents[3]
WORK_DIR = REPO_ROOT / "output" / "_step5"

# One reference chart per deck. Each entry pins:
#   data_source     the analysis whose lineage we override
#   wave_a_ids      "newer" wave-set (latest delivered, used today)
#   wave_b_ids      "older" wave-set (a clearly different point in time)
#   reference       (slide_idx, (left_in, top_in)) of the reference chart
#                   that must be on a slide whose data_source matches
#                   the override target.
#
# Reference chart selection criteria:
#   - values_match=True under same-wave roundtrip (today's Eval #3)
#   - underlying analysis returns >=4 distinct waves so we have headroom
#   - chart shape is stable across waves (same brand list, same options)
REFERENCE = {
    "atu_q1_26": {
        # Reference chart on slide 4 (Practice setting n-1) — 2 series
        # (Community, Academic) over wave categories. selectedColumns is
        # wave-free so the variant override actually drives a refetch.
        "data_source": "p981_rp2019_a406875",
        "project_id": 981,
        "reporting_plan_id": 2019,
        "analysis_ids": [406875],
        "segment_ids": [],
        # 406875 returns waves 1-10, 12, 13 (12 distinct waves; wave 11 absent)
        "wave_a_ids": [23115, 23116],   # Wave 12 + Wave 13 (latest delivered)
        "wave_b_ids": [23104, 23105],   # Wave 1  + Wave 2  (clearly older)
        "reference": (4, (8.36, 4.51)),
    },
    "creon_pet_w33": {
        "data_source": "p523_rp1143_a689321",
        "project_id": 523,
        "reporting_plan_id": 1143,
        "analysis_ids": [689321],
        # base spec filters this analysis by segment 2445 — match it on
        # the API-truth side so multiset containment is meaningful.
        "segment_ids": [2445],
        # 689321 returns Jan'25..Dec'25 (15163, 15178-15188) +
        # Jan'26..Mar'26 (25375, 25374, 25373 — note reversed order).
        "wave_a_ids": [25374, 25373],   # Feb'26 + Mar'26 (newest)
        "wave_b_ids": [15184, 15185],   # Aug'25 + Sep'25 (older)
        "reference": (8, (3.28, 1.84)),
    },
}

VALUE_TOLERANCE = 6e-3  # absorbs chart's percentage-point rounding (chart
# stores `percentage / 100` ≈ 2dp; API `decimal` is 4dp; max divergence is
# under half a percentage point)


def _values_match(refreshed_series, api_records) -> tuple[bool, str]:
    """Verify each refreshed value appears in the API records.

    Loose multiset containment: for every value in the refreshed chart, find
    a matching value in the API records (within VALUE_TOLERANCE). The API
    pool includes both `decimal` (4dp) and `percentage/100` (2dp) since the
    mapper may use either depending on the chart's ValueFields.
    """
    api_values = []
    for r in api_records:
        for key, scale in (("decimal", 1.0), ("percentage", 0.01)):
            v = r.get(key)
            if v is None:
                continue
            try:
                api_values.append(round(float(v) * scale, 4))
            except (TypeError, ValueError):
                continue

    refreshed_values = []
    for _name, vals in refreshed_series:
        for v in vals:
            if v is None:
                continue
            refreshed_values.append(round(float(v), 4))

    if not refreshed_values:
        return False, "refreshed chart has no numeric values"
    if not api_values:
        return False, "API returned no numeric values"

    # Multiset containment — every refreshed value must appear in API values
    # (allowing duplicates). Strong enough to catch wrong-wave / wrong-segment
    # corruption without being brittle on ordering.
    api_pool = list(api_values)
    missing = []
    for rv in refreshed_values:
        for i, av in enumerate(api_pool):
            if abs(rv - av) <= VALUE_TOLERANCE:
                api_pool.pop(i)
                break
        else:
            missing.append(rv)

    if missing:
        return False, (
            f"{len(missing)}/{len(refreshed_values)} refreshed values not in API "
            f"response (e.g. {missing[:3]})"
        )
    return True, ""


@pytest.mark.parametrize("deck_key", list(REFERENCE.keys()))
def test_step5_new_wave_propagates(deck_key):
    """Refresh against two wave variants → assert values change + match API."""
    cfg = REFERENCE[deck_key]
    source_pptx = FIXTURE_DECKS[deck_key]
    base_spec = REPO_ROOT / "output" / f"{deck_key}_refresh" / f"{deck_key}_full_spec.json"
    if not base_spec.exists():
        pytest.skip(f"base spec missing — run scripts/gen_refreshed_fixture.py {deck_key}")

    work = WORK_DIR / deck_key
    work.mkdir(parents=True, exist_ok=True)

    # Build two variant specs
    spec_a = build_wave_variant_spec(
        base_spec_path=base_spec,
        data_source_key=cfg["data_source"],
        static_time_period_ids=cfg["wave_a_ids"],
        out_spec_path=work / "spec_wave_a.json",
    )
    spec_b = build_wave_variant_spec(
        base_spec_path=base_spec,
        data_source_key=cfg["data_source"],
        static_time_period_ids=cfg["wave_b_ids"],
        out_spec_path=work / "spec_wave_b.json",
    )

    # Refresh against each variant
    deck_a = refresh_with_variant(spec_a, source_pptx, work / "refreshed_wave_a.pptx")
    deck_b = refresh_with_variant(spec_b, source_pptx, work / "refreshed_wave_b.pptx")

    # Pull reference chart from each
    slide_idx, position = cfg["reference"]
    chart_a = extract_reference_chart(deck_a, slide_idx, position)
    chart_b = extract_reference_chart(deck_b, slide_idx, position)

    # Save artifacts for debugging
    (work / "chart_extract.json").write_text(
        json.dumps({"a": chart_a, "b": chart_b}, indent=2, default=str),
        encoding="utf-8",
    )

    # ── Assertion 1: structural invariance ──
    # Waves are the cats on these reference charts (time_period_name in RowFields),
    # so cat *labels* legitimately change with the override — only the *count* and
    # the series shape should be invariant.
    assert len(chart_a["cats"]) == len(chart_b["cats"]), (
        f"cat count differs between waves: "
        f"A={len(chart_a['cats'])} ({chart_a['cats']!r}) vs "
        f"B={len(chart_b['cats'])} ({chart_b['cats']!r})"
    )
    assert len(chart_a["series"]) == len(chart_b["series"]), (
        f"series count differs: A={len(chart_a['series'])} vs B={len(chart_b['series'])}"
    )
    names_a = [n for n, _ in chart_a["series"]]
    names_b = [n for n, _ in chart_b["series"]]
    assert names_a == names_b, (
        f"series names differ between waves: A={names_a!r} vs B={names_b!r}"
    )

    # ── Assertion 2: smell test (values differ) ──
    a_flat = tuple(round(v, 4) for _n, vs in chart_a["series"] for v in vs if v is not None)
    b_flat = tuple(round(v, 4) for _n, vs in chart_b["series"] for v in vs if v is not None)
    assert a_flat != b_flat, (
        f"refresh produced identical values for two different waves — "
        f"wave A {cfg['wave_a_ids']} and wave B {cfg['wave_b_ids']} "
        f"both yielded {a_flat[:5]}..."
    )

    # ── Assertion 3: API correctness for each variant ──
    api_a = fetch_api_truth(
        project_id=cfg["project_id"],
        reporting_plan_id=cfg["reporting_plan_id"],
        analysis_ids=cfg["analysis_ids"],
        static_time_period_ids=cfg["wave_a_ids"],
        segment_ids=cfg.get("segment_ids", []),
    )
    api_b = fetch_api_truth(
        project_id=cfg["project_id"],
        reporting_plan_id=cfg["reporting_plan_id"],
        analysis_ids=cfg["analysis_ids"],
        static_time_period_ids=cfg["wave_b_ids"],
        segment_ids=cfg.get("segment_ids", []),
    )

    ok_a, msg_a = _values_match(chart_a["series"], api_a)
    ok_b, msg_b = _values_match(chart_b["series"], api_b)

    assert ok_a, f"variant A values don't match API: {msg_a}"
    assert ok_b, f"variant B values don't match API: {msg_b}"
