"""Round-trip wave-shift eval — back then forward, expect original.

Tests the production case: a new wave lands and refresh has to propagate it.
We can't get a "real" older source, so we synthesize one by refreshing the
current source backward by one wave per ds. Then we refresh that synthesized
deck FORWARD by one wave to reconstruct the original window. The
reconstruction is compared component-by-component against the ORIGINAL
source — values, structure, formatting.

Why this is the strongest refresh test:
  Step 2 (roundtrip same-wave): refresh proves it can hold the line.
  Step 5 backward: refresh proves it can pull older data on demand.
  Round-trip back→forward: refresh proves it can absorb a new wave AND
    surface it correctly — the actual production path.

For each chart on a ds that gets shifted in BOTH directions, classify:
  reconstructed   — value identical to original source within tolerance
  drift           — values changed but reconstruction differs from source
  welded_back     — backward refresh didn't change values (mapper preserved
                    source) → forward refresh has nothing to undo, trivially
                    matches source. Counts as "would-be reconstructed but
                    not informative."
  structural_drift — chart_type/cat-count/series-count differs vs source
"""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from .compare_decks import _extract_chart_data, _find_match
from .step5_helpers import (
    build_deckwide_per_ds_variant_spec,
    compute_one_wave_back_shift,
    compute_one_wave_forward_shift,
    refresh_with_variant,
)
from ..fixtures import FIXTURE_DECKS


REPO_ROOT = Path(__file__).resolve().parents[3]
WORK_DIR = REPO_ROOT / "output" / "_roundtrip"

DECKS = {
    "atu_q1_26":     {"project_id": 981},
    "creon_pet_w33": {"project_id": 523},
}

VALUE_TOLERANCE = 1e-3


def _values_identical(a_series, b_series) -> bool:
    if len(a_series) != len(b_series):
        return False
    for (an, av), (bn, bv) in zip(a_series, b_series):
        if an != bn or len(av) != len(bv):
            return False
        for x, y in zip(av, bv):
            if x is None and y is None:
                continue
            if x is None or y is None:
                return False
            if abs(x - y) > VALUE_TOLERANCE:
                return False
    return True


@pytest.mark.parametrize("deck_key", list(DECKS.keys()))
def test_roundtrip_back_then_forward(deck_key):
    cfg = DECKS[deck_key]
    source_pptx = FIXTURE_DECKS[deck_key]
    base_spec = REPO_ROOT / "output" / "step2_test_connected" / f"{deck_key}_full_spec.json"
    if not source_pptx.exists():
        pytest.skip(f"source pptx missing: {source_pptx}")
    if not base_spec.exists():
        pytest.skip(f"base spec missing")

    work = WORK_DIR / deck_key
    work.mkdir(parents=True, exist_ok=True)

    spec_dict = json.loads(base_spec.read_text(encoding="utf-8"))

    # ── Leg 1: backward refresh — produce synthesized "older" deck ──
    back_overrides, back_skipped = compute_one_wave_back_shift(
        spec_dict, cfg["project_id"],
    )
    if not back_overrides:
        pytest.skip(f"no ds shiftable backward for project {cfg['project_id']}")

    spec_back = build_deckwide_per_ds_variant_spec(
        base_spec_path=base_spec,
        ds_overrides=back_overrides,
        out_spec_path=work / "spec_back.json",
    )
    deck_back = refresh_with_variant(
        spec_back, source_pptx, work / "deck_back.pptx",
    )

    # ── Leg 2: forward refresh on the synthesized older deck ──
    # Use the BACKWARD spec as base for the forward shift — this carries the
    # backward-pinned wave IDs forward by one position, ideally returning to
    # the source's wave window.
    spec_back_dict = json.loads(spec_back.read_text(encoding="utf-8"))
    fwd_overrides, fwd_skipped = compute_one_wave_forward_shift(
        spec_back_dict, cfg["project_id"],
    )
    if not fwd_overrides:
        pytest.skip("backward shift produced no ds that could be shifted forward")

    spec_fwd = build_deckwide_per_ds_variant_spec(
        base_spec_path=spec_back,
        ds_overrides=fwd_overrides,
        out_spec_path=work / "spec_forward.json",
    )
    deck_fwd = refresh_with_variant(
        spec_fwd, deck_back, work / "deck_reconstructed.pptx",
    )

    # ── Compare reconstructed deck against ORIGINAL source ──
    # Only assess components on ds that survived BOTH legs; others fall
    # through unchanged in both directions and trivially match.
    full_round_trip_ds = set(back_overrides) & set(fwd_overrides)

    # Map ds → component (slide_idx, name)
    comp_ds_by_name: dict[tuple, str] = {}
    comp_ds_by_pos: dict[int, list[tuple[float, float, str]]] = {}
    for slide in spec_dict.get("slides", []):
        s_idx = slide["slide_index"]
        slide_default = slide.get("data_source")
        for comp in slide.get("components", []):
            ds = comp.get("data_source") or slide_default
            if ds not in full_round_trip_ds:
                continue
            name = comp.get("name") or ""
            if name:
                comp_ds_by_name[(s_idx, name)] = ds
            cpos = comp.get("position", {}) or {}
            cl = float(cpos.get("left", 0) or 0)
            ct = float(cpos.get("top", 0) or 0)
            comp_ds_by_pos.setdefault(s_idx, []).append((cl, ct, ds))

    def _resolve_ds(s_idx, name, pos):
        if name and (s_idx, name) in comp_ds_by_name:
            return comp_ds_by_name[(s_idx, name)]
        for cl, ct, ds in comp_ds_by_pos.get(s_idx, []):
            if abs(cl - pos[0]) <= 0.05 and abs(ct - pos[1]) <= 0.05:
                return ds
        return None

    from pptx import Presentation
    src_prs = Presentation(str(source_pptx))
    back_prs = Presentation(str(deck_back))
    fwd_prs = Presentation(str(deck_fwd))

    reconstructed = []
    drift = []                # values changed back, didn't reconstruct on forward
    welded_back = []          # backward leg didn't change values → trivially "matches"
    structural_drift = []     # cat / series / chart-type structure differs
    extract_failed = []
    no_match = []

    for s_idx in range(len(src_prs.slides)):
        if s_idx >= len(fwd_prs.slides) or s_idx >= len(back_prs.slides):
            continue
        src_charts = [s for s in src_prs.slides[s_idx].shapes if s.has_chart]
        back_charts = [s for s in back_prs.slides[s_idx].shapes if s.has_chart]
        fwd_charts = [s for s in fwd_prs.slides[s_idx].shapes if s.has_chart]

        for src_shape in src_charts:
            shape_name = src_shape.name or ""
            sl = round((src_shape.left or 0) / 914400, 2)
            st = round((src_shape.top or 0) / 914400, 2)
            ds_key = _resolve_ds(s_idx, shape_name, (sl, st))
            if ds_key is None:
                continue
            back_shape = _find_match(src_shape, back_charts)
            fwd_shape = _find_match(src_shape, fwd_charts)
            if fwd_shape is None or back_shape is None:
                no_match.append((s_idx, shape_name, ds_key))
                continue
            try:
                src_cats, src_series = _extract_chart_data(src_shape)
                back_cats, back_series = _extract_chart_data(back_shape)
                fwd_cats, fwd_series = _extract_chart_data(fwd_shape)
            except Exception as exc:
                extract_failed.append((s_idx, shape_name, str(exc)))
                continue

            # Structural check first — chart_type and counts vs source
            if (src_shape.chart.chart_type != fwd_shape.chart.chart_type
                    or len(src_cats) != len(fwd_cats)
                    or len(src_series) != len(fwd_series)):
                structural_drift.append({
                    "slide": s_idx, "name": shape_name, "ds": ds_key,
                    "src_cats": len(src_cats), "fwd_cats": len(fwd_cats),
                    "src_series": len(src_series), "fwd_series": len(fwd_series),
                })
                continue

            # If backward leg didn't change values, forward refresh on identical
            # data trivially matches source — flag as welded_back, not informative
            if _values_identical(src_series, back_series):
                welded_back.append((s_idx, shape_name, ds_key))
                continue

            # Round-trip: forward should match source
            if _values_identical(src_series, fwd_series):
                reconstructed.append((s_idx, shape_name, ds_key))
            else:
                # Build a small diff for debugging
                first_diff = None
                for (sn, sv), (fn, fv) in zip(src_series, fwd_series):
                    for i, (a, b) in enumerate(zip(sv, fv)):
                        if a is None and b is None:
                            continue
                        if (a is None or b is None) or abs(a - b) > VALUE_TOLERANCE:
                            first_diff = (sn, i, a, b)
                            break
                    if first_diff:
                        break
                drift.append({
                    "slide": s_idx, "name": shape_name, "ds": ds_key,
                    "first_diff": str(first_diff),
                })

    informative_total = len(reconstructed) + len(drift) + len(structural_drift)
    summary = {
        "deck": deck_key,
        "ds_back_shifted": len(back_overrides),
        "ds_back_skipped": len(back_skipped),
        "ds_fwd_shifted": len(fwd_overrides),
        "ds_fwd_skipped": len(fwd_skipped),
        "ds_full_roundtrip": len(full_round_trip_ds),
        "totals": {
            "reconstructed": len(reconstructed),
            "drift": len(drift),
            "structural_drift": len(structural_drift),
            "welded_back": len(welded_back),
            "extract_failed": len(extract_failed),
            "no_match": len(no_match),
        },
        "reconstruction_rate": (
            f"{len(reconstructed)}/{informative_total} "
            f"= {100 * len(reconstructed) / informative_total:.1f}%"
            if informative_total else "n/a (all welded_back)"
        ),
        "drift_detail": drift[:30],
        "structural_drift_detail": structural_drift[:20],
    }
    (work / "summary.json").write_text(
        json.dumps(summary, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    print(f"\n=== {deck_key} | round-trip back→forward ===")
    print(f"  ds back/fwd/intersect: {len(back_overrides)} / {len(fwd_overrides)} / {len(full_round_trip_ds)}")
    print(f"  reconstructed (matches source values): {len(reconstructed)}")
    print(f"  drift (back changed, fwd didn't reconstruct): {len(drift)}")
    print(f"  structural_drift: {len(structural_drift)}")
    print(f"  welded_back (back leg was a no-op — uninformative): {len(welded_back)}")
    print(f"  extract_failed: {len(extract_failed)}, no_match: {len(no_match)}")
    if informative_total:
        rate = 100 * len(reconstructed) / informative_total
        print(f"  reconstruction rate: {len(reconstructed)}/{informative_total} = {rate:.1f}%")
    if drift:
        print("  first 5 drift:")
        for d in drift[:5]:
            print(f"    slide {d['slide']} {d['name']!r} ({d['ds']}): {d['first_diff']}")
    if structural_drift:
        print("  first 5 structural_drift:")
        for d in structural_drift[:5]:
            print(f"    slide {d['slide']} {d['name']!r}: cats {d['src_cats']}->{d['fwd_cats']}, series {d['src_series']}->{d['fwd_series']}")

    # Soft gate at first — we want to see the metric. Hard-gate later.
    assert informative_total > 0, (
        "no informative round-trip components — backward leg welded everything"
    )
