"""Helpers for the Step 5 "new data propagates" eval.

Step 5 of the 7-step refresh framework asserts: when source and refresh
target different waves, values update to the targeted wave (not just
preserve source). Eval #3 only verifies same-wave roundtrip; this
module supports variant-vs-variant comparison plus direct API
correctness checks.

Public surface:
    build_wave_variant_spec()   write a modified full_spec.json with
                                static_time_period_ids overridden on a
                                specific data_source
    refresh_with_variant()      run dual-mode refresh against a variant
                                spec, return path to refreshed pptx
    extract_reference_chart()   pull (cats, series) out of a chart at
                                a known (slide_idx, position)
    fetch_api_truth()           direct Synapse query for one analysis
                                + wave set, returns the same pivot the
                                mapper would produce
"""
from __future__ import annotations

import json
import shutil
import sys
from copy import deepcopy
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


def build_wave_variant_spec(
    base_spec_path: Path,
    data_source_key: str,
    static_time_period_ids: list[int],
    out_spec_path: Path,
) -> Path:
    """Clone the base spec, override one data_source's wave lineage, write to out path.

    Forces static-mode fetching for the named data_source by setting
    static_time_period_ids and clearing dynamic_latest_n. Other data
    sources are untouched so non-reference charts behave normally.
    """
    spec = json.loads(base_spec_path.read_text(encoding="utf-8"))
    spec = deepcopy(spec)
    if data_source_key not in spec.get("data_sources", {}):
        raise KeyError(f"data_source {data_source_key!r} not in spec")
    ds = spec["data_sources"][data_source_key]
    ds["static_time_period_ids"] = list(static_time_period_ids)
    ds["static_time_period_names"] = []  # let the mapper resolve names from IDs
    ds["dynamic_latest_n"] = 0
    ds["include_live_wave"] = False
    out_spec_path.write_text(
        json.dumps(spec, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    return out_spec_path


def refresh_with_variant(
    spec_path: Path,
    source_pptx: Path,
    out_pptx: Path,
) -> Path:
    """Run dual-mode refresh against a variant spec. Returns out_pptx path."""
    from slidegen.intelligent_refresh import refresh_deck_from_spec

    refresh_deck_from_spec(
        spec_path=str(spec_path),
        pptx_path=str(source_pptx),
        output_path=str(out_pptx),
    )
    return out_pptx


def extract_reference_chart(
    pptx_path: Path, slide_idx: int, position: tuple[float, float],
) -> dict:
    """Pull (cats, [(series_name, values)]) for the chart at (slide_idx, position).

    Position match uses a 0.1in tolerance (same as compare_decks).
    Returns {'cats': list[str], 'series': list[(name, [float])]}.
    """
    from pptx import Presentation
    from .compare_decks import _pos, _extract_chart_data

    pres = Presentation(str(pptx_path))
    if slide_idx >= len(pres.slides):
        raise IndexError(f"slide_idx {slide_idx} out of range")
    target_l, target_t = position
    for shape in pres.slides[slide_idx].shapes:
        if not shape.has_chart:
            continue
        sl, st = _pos(shape)
        if abs(sl - target_l) <= 0.1 and abs(st - target_t) <= 0.1:
            cats, series = _extract_chart_data(shape)
            return {"cats": cats, "series": series}
    raise LookupError(
        f"no chart at slide {slide_idx} position ({target_l},{target_t})"
    )


def fetch_api_truth(
    project_id: int,
    reporting_plan_id: int,
    analysis_ids: list[int],
    static_time_period_ids: list[int],
    segment_ids: list[int] | None = None,
) -> list[dict]:
    """Direct Synapse query for the (analysis × wave [× segment]) tuple.

    Bypasses the mapper — returns raw records so the test can compare
    refreshed values against the API source of truth. `segment_ids` must
    match the base spec's segment filter on the same data_source so that
    the API records are comparable to the refreshed chart values.
    """
    from slidegen.intelligent_refresh import fetch_synapse_data

    lineage = {
        "project_id": project_id,
        "reporting_plan_id": reporting_plan_id,
        "analysis_ids": list(analysis_ids),
        "segment_ids": list(segment_ids) if segment_ids else [],
        "dynamic_latest_n": 0,
        "static_time_period_ids": list(static_time_period_ids),
        "include_live_wave": False,
    }
    records, _ = fetch_synapse_data(lineage)
    return records
