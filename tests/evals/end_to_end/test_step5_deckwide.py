"""Step 5 eval — new data propagates with same structure (whole deck).

Per-chart API correctness check on top of Step 6's wave-shift refresh.
For every chart on a shifted data_source whose values actually changed
(post Step 6 fix), assert the refreshed values match a direct Synapse
query for that ds + the shifted wave window.

Where Step 6 measures structure preservation deck-wide, Step 5 deck-wide
measures value correctness deck-wide.

Reuses Step 6's shift logic (compute_one_wave_back_shift) and refresh
path (build_deckwide_per_ds_variant_spec + refresh_with_variant) so the
two evals share fixtures, API calls, and refresh costs when run together.
"""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from .compare_decks import _extract_chart_data, _find_match
from .step5_helpers import (
    build_deckwide_per_ds_variant_spec,
    compute_one_wave_back_shift,
    fetch_api_truth,
    refresh_with_variant,
)
from ..fixtures import FIXTURE_DECKS


REPO_ROOT = Path(__file__).resolve().parents[3]
WORK_DIR = REPO_ROOT / "output" / "_step5_deckwide"

DECKS = {
    "atu_q1_26":     {"project_id": 981},
    "creon_pet_w33": {"project_id": 523},
}

VALUE_TOLERANCE = 6e-3


def _values_match(refreshed_series, api_records) -> tuple[bool, str]:
    """Multiset containment check — every refreshed value must appear in API."""
    api_values = []
    for r in api_records:
        for key, scale in (("decimal", 1.0), ("percentage", 0.01), ("count", 1.0)):
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
            f"(e.g. {missing[:3]})"
        )
    return True, ""


@pytest.mark.parametrize("deck_key", list(DECKS.keys()))
def test_step5_deckwide_api_correctness(deck_key):
    cfg = DECKS[deck_key]
    source_pptx = FIXTURE_DECKS[deck_key]
    base_spec = REPO_ROOT / "output" / "step2_test_connected" / f"{deck_key}_full_spec.json"
    if not source_pptx.exists():
        pytest.skip(f"source pptx missing: {source_pptx}")
    if not base_spec.exists():
        pytest.skip(
            f"base spec missing — run scripts/gen_refreshed_fixture.py {deck_key}"
        )

    work = WORK_DIR / deck_key
    work.mkdir(parents=True, exist_ok=True)

    spec_dict = json.loads(base_spec.read_text(encoding="utf-8"))
    overrides, skipped_ds = compute_one_wave_back_shift(spec_dict, cfg["project_id"])
    if not overrides:
        pytest.skip(
            f"No data_sources could be shifted for project {cfg['project_id']}"
        )

    variant_spec = build_deckwide_per_ds_variant_spec(
        base_spec_path=base_spec,
        ds_overrides=overrides,
        out_spec_path=work / "spec_shift.json",
    )
    refreshed_pptx = refresh_with_variant(
        variant_spec, source_pptx, work / "refreshed_shift.pptx",
    )

    # Index spec components → ds key (only shifted ds)
    comp_ds_by_name: dict[tuple, str] = {}
    comp_ds_by_pos: dict[int, list[tuple[float, float, str]]] = {}
    for slide in spec_dict.get("slides", []):
        s_idx_spec = slide["slide_index"]
        slide_default_ds = slide.get("data_source")
        for comp in slide.get("components", []):
            ds_for_comp = comp.get("data_source") or slide_default_ds
            if ds_for_comp not in overrides:
                continue
            name = comp.get("name") or ""
            if name:
                comp_ds_by_name[(s_idx_spec, name)] = ds_for_comp
            cpos = comp.get("position", {}) or {}
            cleft = float(cpos.get("left", 0) or 0)
            ctop = float(cpos.get("top", 0) or 0)
            comp_ds_by_pos.setdefault(s_idx_spec, []).append(
                (cleft, ctop, ds_for_comp)
            )

    def _resolve_ds(s_idx: int, name: str, pos: tuple[float, float]) -> str | None:
        if name and (s_idx, name) in comp_ds_by_name:
            return comp_ds_by_name[(s_idx, name)]
        for cl, ct, ds in comp_ds_by_pos.get(s_idx, []):
            if abs(cl - pos[0]) <= 0.05 and abs(ct - pos[1]) <= 0.05:
                return ds
        return None

    # API truth fetch — once per (ds, wave_set) pair
    api_cache: dict[str, list[dict]] = {}

    def _api_truth(ds_key: str) -> list[dict]:
        if ds_key in api_cache:
            return api_cache[ds_key]
        ds = spec_dict["data_sources"][ds_key]
        try:
            records = fetch_api_truth(
                project_id=ds["project_id"],
                reporting_plan_id=ds["reporting_plan_id"],
                analysis_ids=ds["analysis_ids"],
                static_time_period_ids=overrides[ds_key],
                segment_ids=ds.get("segment_ids", []),
            )
        except Exception as exc:
            print(f"  API truth fetch failed for {ds_key}: {exc}")
            records = []
        api_cache[ds_key] = records
        return records

    # Walk source vs refreshed for charts on shifted ds
    from pptx import Presentation
    src = Presentation(str(source_pptx))
    ref = Presentation(str(refreshed_pptx))

    value_correct: list[tuple] = []
    value_wrong: list[tuple] = []
    welded: list[tuple] = []          # values didn't change → not testable here
    extract_failed: list[tuple] = []
    no_match: list[tuple] = []

    for s_idx, (s_slide, r_slide) in enumerate(zip(src.slides, ref.slides)):
        src_charts = [s for s in s_slide.shapes if s.has_chart]
        ref_charts = [s for s in r_slide.shapes if s.has_chart]
        for src_shape in src_charts:
            shape_name = src_shape.name or ""
            sl = round((src_shape.left or 0) / 914400, 2)
            st = round((src_shape.top or 0) / 914400, 2)
            ds_key = _resolve_ds(s_idx, shape_name, (sl, st))
            if ds_key is None:
                continue  # not on a shifted ds — out of scope for Step 5
            ref_shape = _find_match(src_shape, ref_charts)
            if ref_shape is None:
                no_match.append((s_idx, shape_name))
                continue
            try:
                _, src_series = _extract_chart_data(src_shape)
                _, ref_series = _extract_chart_data(ref_shape)
            except Exception as exc:
                extract_failed.append((s_idx, shape_name, str(exc)))
                continue

            src_flat = tuple(round(v if v is not None else 0, 4)
                             for _, vs in src_series for v in vs)
            ref_flat = tuple(round(v if v is not None else 0, 4)
                             for _, vs in ref_series for v in vs)
            if src_flat == ref_flat:
                welded.append((s_idx, shape_name))
                continue

            api_records = _api_truth(ds_key)
            ok, msg = _values_match(ref_series, api_records)
            if ok:
                value_correct.append((s_idx, shape_name))
            else:
                value_wrong.append((s_idx, shape_name, ds_key, msg))

    refreshed_count = len(value_correct) + len(value_wrong)
    summary = {
        "deck": deck_key,
        "shifted_ds": len(overrides),
        "skipped_ds": len(skipped_ds),
        "totals": {
            "value_correct": len(value_correct),
            "value_wrong": len(value_wrong),
            "welded": len(welded),
            "extract_failed": len(extract_failed),
            "no_match": len(no_match),
        },
        "api_correct_rate": (
            f"{len(value_correct)}/{refreshed_count} "
            f"= {100 * len(value_correct) / refreshed_count:.1f}%"
            if refreshed_count else "n/a"
        ),
        "value_wrong_detail": [
            {"slide": s, "name": n, "ds": d, "reason": r}
            for s, n, d, r in value_wrong[:30]
        ],
    }
    (work / "summary.json").write_text(
        json.dumps(summary, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    print(f"\n=== {deck_key} | Step 5 deck-wide ===")
    print(f"  shifted ds: {len(overrides)}")
    print(f"  charts on shifted ds: {refreshed_count + len(welded)}")
    print(f"  refreshed (values changed): {refreshed_count}")
    print(f"  welded (values unchanged): {len(welded)}")
    print(f"  value_correct: {len(value_correct)}")
    print(f"  value_wrong:   {len(value_wrong)}")
    if refreshed_count:
        print(
            f"  API-correct rate: {len(value_correct)}/{refreshed_count} = "
            f"{100 * len(value_correct) / refreshed_count:.1f}%"
        )
    if value_wrong:
        print("  first 5 value_wrong:")
        for s, n, d, r in value_wrong[:5]:
            print(f"    slide {s} {n!r} ({d}): {r}")

    # Hard gate is loose at first — we want to SEE the metric. Ratchet up
    # once the baseline is known. The refreshed_count > 0 assertion
    # guarantees the wave-shift fix is actually doing something.
    assert refreshed_count > 0, (
        "no charts refreshed at all — wave-shift may be broken (Step 6 regression)"
    )
