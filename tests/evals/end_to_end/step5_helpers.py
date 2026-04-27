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


def build_deckwide_wave_variant_spec(
    base_spec_path: Path,
    static_time_period_ids: list[int],
    out_spec_path: Path,
    *,
    project_id: int | None = None,
) -> Path:
    """Clone the base spec, override the wave lineage on EVERY data_source.

    Used by the deck-wide wave-shift eval: simulates rolling the entire deck
    forward (or back) to a new wave window. Every analysis is asked to fetch
    the same target waves; analyses that don't have those waves available
    will return empty records and the chart falls back to source.

    `project_id` filters which data_sources to override — useful when a spec
    has data_sources from multiple projects and we only want to shift one.
    """
    spec = json.loads(base_spec_path.read_text(encoding="utf-8"))
    spec = deepcopy(spec)
    overridden = 0
    for ds_key, ds in spec.get("data_sources", {}).items():
        if project_id is not None and ds.get("project_id") != project_id:
            continue
        ds["static_time_period_ids"] = list(static_time_period_ids)
        ds["static_time_period_names"] = []
        ds["dynamic_latest_n"] = 0
        ds["include_live_wave"] = False
        overridden += 1
    if overridden == 0:
        raise ValueError(
            f"No data_sources matched project_id={project_id} in {base_spec_path}"
        )
    out_spec_path.write_text(
        json.dumps(spec, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    return out_spec_path


def build_deckwide_per_ds_variant_spec(
    base_spec_path: Path,
    ds_overrides: dict[str, list[int]],
    out_spec_path: Path,
    *,
    clear_table_cell_values: bool = True,
) -> Path:
    """Override static_time_period_ids on specific data_sources only.

    Used when each data_source needs a different wave override (e.g. the
    deck-wide one-wave-back shift, where each analysis has its own
    chronological wave order). data_sources NOT in `ds_overrides` are left
    untouched and refresh against their base config.

    `clear_table_cell_values` (default True): tables in the spec carry
    pre-computed `cell_values` from the source render. When the wave
    window shifts, those cached values are stale — clearing them forces
    the refresh pipeline to re-pivot from API data. Without this flag,
    every table appears "welded" in the wave-shift eval because its
    cell_values are reused verbatim.
    """
    spec = json.loads(base_spec_path.read_text(encoding="utf-8"))
    spec = deepcopy(spec)
    for ds_key, ids in ds_overrides.items():
        if ds_key not in spec.get("data_sources", {}):
            raise KeyError(f"data_source {ds_key!r} not in spec")
        ds = spec["data_sources"][ds_key]
        ds["static_time_period_ids"] = list(ids)
        ds["static_time_period_names"] = []
        ds["dynamic_latest_n"] = 0
        ds["include_live_wave"] = False

    if clear_table_cell_values:
        for slide in spec.get("slides", []):
            for comp in slide.get("components", []):
                ctype = comp.get("type", "")
                if ctype in ("table", "value_table") and "cell_values" in comp:
                    comp.pop("cell_values", None)

    out_spec_path.write_text(
        json.dumps(spec, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    return out_spec_path


def compute_one_wave_back_shift(
    spec: dict, project_id: int,
) -> tuple[dict[str, list[int]], list[tuple[str, str]]]:
    """Per-data_source: shift each wave back by one in chronological order.

    Preserves each chart's source wave count — matches the April-end claim
    that the deck rolls forward (or back) by one wave without changing
    chart shape.

    Strategy per data_source:
      - If `static_time_period_ids` is set, shift each ID back by one position
        in the analysis's chronological wave order. If any ID is at position 0
        (no older wave), skip this data_source.
      - Else if `dynamic_latest_n=K`, return the K waves that would have been
        latest one wave earlier (positions [-K-1 : -1]). If K+1 > total waves,
        skip.
      - Else, skip (no wave config).

    Skipped data_sources fall through to base-spec refresh (same waves as
    source) — the chart will appear `welded` in the eval, which is the
    correct classification.

    Returns (overrides, skipped) where:
      overrides : dict[ds_key, shifted_tp_ids]
      skipped   : list[(ds_key, reason)] for visibility
    """
    from slidegen.intelligent_refresh import fetch_synapse_data

    overrides: dict[str, list[int]] = {}
    skipped: list[tuple[str, str]] = []
    cache: dict[int, list[int]] = {}  # analysis_id -> chronologically-ordered tp_ids

    for ds_key, ds in spec.get("data_sources", {}).items():
        if ds.get("project_id") != project_id:
            continue
        aids = ds.get("analysis_ids") or []
        if not aids:
            skipped.append((ds_key, "no analysis_ids"))
            continue
        aid = aids[0]

        if aid not in cache:
            try:
                _, df = fetch_synapse_data({
                    "project_id": project_id,
                    "reporting_plan_id": ds.get("reporting_plan_id"),
                    "analysis_ids": [aid],
                    "segment_ids": ds.get("segment_ids") or [],
                    "dynamic_latest_n": 0,
                    "static_time_period_ids": [],
                    "static_time_period_names": [],
                    "include_live_wave": False,
                })
            except Exception as exc:
                skipped.append((ds_key, f"a{aid}: fetch failed: {exc}"))
                cache[aid] = []
                continue
            if df.empty or "time_period_id" not in df.columns:
                cache[aid] = []
            else:
                cache[aid] = (df.groupby("time_period_name")["time_period_id"]
                              .max().sort_values().tolist())
        order = cache[aid]
        if not order:
            skipped.append((ds_key, f"a{aid}: no waves available"))
            continue

        static_ids = ds.get("static_time_period_ids") or []
        dyn_n = ds.get("dynamic_latest_n") or 0

        if static_ids:
            shifted: list[int] = []
            failure = None
            for sid in static_ids:
                try:
                    idx = order.index(sid)
                except ValueError:
                    failure = f"id {sid} not in chronological order"
                    break
                if idx == 0:
                    failure = f"id {sid} is oldest, cannot shift back"
                    break
                shifted.append(order[idx - 1])
            if failure is None:
                overrides[ds_key] = shifted
            else:
                skipped.append((ds_key, f"a{aid}: {failure}"))
        elif dyn_n > 0:
            if dyn_n + 1 > len(order):
                skipped.append(
                    (ds_key, f"a{aid}: only {len(order)} waves, need >={dyn_n + 1}")
                )
            else:
                overrides[ds_key] = order[-(dyn_n + 1):-1]
        else:
            skipped.append((ds_key, f"a{aid}: no static ids and no dynamic_latest_n"))

    return overrides, skipped


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
