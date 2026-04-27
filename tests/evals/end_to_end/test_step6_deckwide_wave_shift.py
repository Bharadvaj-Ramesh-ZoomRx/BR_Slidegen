"""Step 6 eval — whole-deck wave shift, structure preserved.

Asserts the April-end claim: "a deck containing waves [..., N-1, N] can be
refreshed against waves [..., N-2, N-1] without anything growing or shrinking
or breaking, and the data propagates through every chart that isn't welded by
wave-pinned selectedColumns."

Per-data_source shift: each chart's analysis is shifted back by one position
in its OWN chronological wave order, preserving the source chart's wave count.
A 6-wave chart on CREON slide 8 stays 6-wave; a 2-wave chart stays 2-wave.

For each fixture deck:
  1. Query the API once per unique analysis to get chronological wave order.
  2. Compute a one-wave-back override per data_source, preserving wave count.
  3. Refresh against the per-ds variant spec.
  4. Walk source vs refreshed component-by-component and classify each:
       refreshed   — values changed, structure invariant
       welded      — values identical to source (mapper silent-fallback,
                     usually wave-pinned selectedColumns post-Fix-L holdouts
                     or analyses too short to shift)
       drift       — structure changed (chart_type / series count / cat
                     count / table dims) — genuine failure

The welded count is the metric we want to drive down via Connector-side
fixes (Vijay's lane). drift is the metric this eval gates on.

Backward shift is symmetric to forward shift (the production case): if a
chart accepts [W12, W13] -> [W10, W12], it accepts [W11, W12] -> [W12, W13].
Backward shifts let us test today; forward shifts can use this eval's output
decks as synthesized "older sources."

Hits Synapse once per unique analysis (cached) plus the refresh — ~3-6 min
per deck.
"""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from .compare_decks import _extract_chart_data, _extract_table_rows, _find_match, _pos
from .step5_helpers import (
    build_deckwide_per_ds_variant_spec,
    compute_one_wave_back_shift,
    refresh_with_variant,
)
from ..fixtures import FIXTURE_DECKS


REPO_ROOT = Path(__file__).resolve().parents[3]
WORK_DIR = REPO_ROOT / "output" / "_step6"

SHIFT = {
    "atu_q1_26":     {"project_id": 981, "label": "ATU one-wave-back per ds"},
    "creon_pet_w33": {"project_id": 523, "label": "CREON one-wave-back per ds"},
}

VALUE_TOLERANCE = 1e-6


def _values_identical(src_series, ref_series) -> bool:
    """All values match within VALUE_TOLERANCE → mapper silent-fell-back to source."""
    if len(src_series) != len(ref_series):
        return False
    for (sn, sv), (rn, rv) in zip(src_series, ref_series):
        if sn != rn or len(sv) != len(rv):
            return False
        for x, y in zip(sv, rv):
            if x is None and y is None:
                continue
            if x is None or y is None:
                return False
            if abs(x - y) > VALUE_TOLERANCE:
                return False
    return True


@pytest.mark.parametrize("deck_key", list(SHIFT.keys()))
def test_deckwide_wave_shift(deck_key):
    cfg = SHIFT[deck_key]
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
    (work / "shift_overrides.json").write_text(
        json.dumps(
            {"overrides": overrides, "skipped": skipped_ds},
            indent=2, ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    if not overrides:
        pytest.skip(
            f"No data_sources could be shifted for project {cfg['project_id']}"
        )
    print(f"\n[step6] {deck_key}: shifted {len(overrides)} ds, "
          f"skipped {len(skipped_ds)} (see shift_overrides.json)")

    variant_spec = build_deckwide_per_ds_variant_spec(
        base_spec_path=base_spec,
        ds_overrides=overrides,
        out_spec_path=work / "spec_shift.json",
    )

    refreshed_pptx = refresh_with_variant(
        variant_spec, source_pptx, work / "refreshed_shift.pptx",
    )

    # Walk source vs refreshed
    from pptx import Presentation
    src = Presentation(str(source_pptx))
    ref = Presentation(str(refreshed_pptx))

    # Structural assertion: same slide count
    assert len(src.slides) == len(ref.slides), (
        f"slide count drift: src={len(src.slides)} vs ref={len(ref.slides)}"
    )

    # Build the set of shifted ds_keys for classification
    shifted_ds = set(overrides.keys())

    # Index spec components by (slide_idx, name) AND (slide_idx, position)
    # → data_source. Some specs use `name` reliably, others lose name in
    # round-trip; positions can diverge by tiny rounding amounts. Try name
    # first (exact), fall back to position with a tolerance match.
    comp_ds_by_name: dict[tuple, str | None] = {}
    comp_ds_by_pos: dict[int, list[tuple[float, float, str | None]]] = {}
    for slide in spec_dict.get("slides", []):
        s_idx_spec = slide["slide_index"]
        slide_default_ds = slide.get("data_source")
        for comp in slide.get("components", []):
            ds_for_comp = comp.get("data_source") or slide_default_ds
            name = comp.get("name") or ""
            if name:
                comp_ds_by_name[(s_idx_spec, name)] = ds_for_comp
            cpos = comp.get("position", {}) or {}
            cleft = float(cpos.get("left", 0) or 0)
            ctop = float(cpos.get("top", 0) or 0)
            comp_ds_by_pos.setdefault(s_idx_spec, []).append((cleft, ctop, ds_for_comp))

    def _classify_ds(slide_idx: int, pos: tuple, name: str = "") -> str:
        """Return 'shifted' | 'unshifted' | 'unknown' for this component."""
        ds_key = None
        if name:
            ds_key = comp_ds_by_name.get((slide_idx, name))
        if ds_key is None:
            for cl, ct, ds in comp_ds_by_pos.get(slide_idx, []):
                if abs(cl - pos[0]) <= 0.05 and abs(ct - pos[1]) <= 0.05:
                    ds_key = ds
                    break
        if ds_key is None:
            return "unknown"
        return "shifted" if ds_key in shifted_ds else "unshifted"

    refreshed_n = 0
    welded_shifted_n = 0       # ds was shifted but values didn't change → real welded
    unshifted_n = 0            # ds was skipped → correctly identical to source
    unknown_n = 0              # could not match component to a ds (rare)
    drift_on_shifted = []      # drift caused by the wave shift itself — gates the test
    drift_on_unshifted = []    # drift on a non-shifted ds — pre-existing Step 2 noise
    drift_unknown = []         # drift where ds couldn't be resolved
    no_match = []              # source shape with no refreshed counterpart
    chart_total = 0
    table_total = 0

    def _record_drift(s_idx, pos, kind, reason, name=""):
        state = _classify_ds(s_idx, pos, name=name)
        item = (s_idx, pos, kind, reason)
        if state == "shifted":
            drift_on_shifted.append(item)
        elif state == "unshifted":
            drift_on_unshifted.append(item)
        else:
            drift_unknown.append(item)

    for s_idx, (s_slide, r_slide) in enumerate(zip(src.slides, ref.slides)):
        src_charts = [s for s in s_slide.shapes if s.has_chart]
        ref_charts = [s for s in r_slide.shapes if s.has_chart]
        src_tables = [s for s in s_slide.shapes if s.has_table]
        ref_tables = [s for s in r_slide.shapes if s.has_table]

        # Component-count parity per slide
        if len(src_charts) != len(ref_charts):
            _record_drift(s_idx, (0.0, 0.0), "chart-count",
                          f"src {len(src_charts)} vs ref {len(ref_charts)}")
        if len(src_tables) != len(ref_tables):
            _record_drift(s_idx, (0.0, 0.0), "table-count",
                          f"src {len(src_tables)} vs ref {len(ref_tables)}")

        # Charts
        for src_shape in src_charts:
            chart_total += 1
            pos = _pos(src_shape)
            ref_shape = _find_match(src_shape, ref_charts)
            if ref_shape is None:
                no_match.append((s_idx, pos, "chart"))
                continue
            # Structural checks
            shape_name = src_shape.name or ""
            if src_shape.chart.chart_type != ref_shape.chart.chart_type:
                _record_drift(s_idx, pos, "chart-type",
                              f"{src_shape.chart.chart_type} -> {ref_shape.chart.chart_type}",
                              name=shape_name)
                continue
            try:
                src_cats, src_series = _extract_chart_data(src_shape)
                ref_cats, ref_series = _extract_chart_data(ref_shape)
            except Exception as e:
                _record_drift(s_idx, pos, "extract-fail", str(e), name=shape_name)
                continue
            if len(src_cats) != len(ref_cats):
                _record_drift(s_idx, pos, "cat-count",
                              f"{len(src_cats)} -> {len(ref_cats)}",
                              name=shape_name)
                continue
            if len(src_series) != len(ref_series):
                _record_drift(s_idx, pos, "series-count",
                              f"{len(src_series)} -> {len(ref_series)}",
                              name=shape_name)
                continue
            # Classification by ds shifted state, then by values diff
            ds_state = _classify_ds(s_idx, pos, name=src_shape.name or "")
            if _values_identical(src_series, ref_series):
                if ds_state == "shifted":
                    welded_shifted_n += 1
                elif ds_state == "unshifted":
                    unshifted_n += 1
                else:
                    unknown_n += 1
            else:
                refreshed_n += 1

        # Tables
        for src_shape in src_tables:
            table_total += 1
            pos = _pos(src_shape)
            ref_shape = _find_match(src_shape, ref_tables)
            if ref_shape is None:
                no_match.append((s_idx, pos, "table"))
                continue
            shape_name = src_shape.name or ""
            try:
                src_rows = _extract_table_rows(src_shape)
                ref_rows = _extract_table_rows(ref_shape)
            except Exception as e:
                _record_drift(s_idx, pos, "table-extract-fail", str(e), name=shape_name)
                continue
            if len(src_rows) != len(ref_rows):
                _record_drift(s_idx, pos, "table-row-count",
                              f"{len(src_rows)} -> {len(ref_rows)}",
                              name=shape_name)
                continue
            if src_rows and ref_rows and len(src_rows[0]) != len(ref_rows[0]):
                _record_drift(s_idx, pos, "table-col-count",
                              f"{len(src_rows[0])} -> {len(ref_rows[0])}",
                              name=shape_name)
                continue
            # Classify table by cell-text identity + ds shifted state
            ds_state = _classify_ds(s_idx, pos, name=src_shape.name or "")
            if src_rows == ref_rows:
                if ds_state == "shifted":
                    welded_shifted_n += 1
                elif ds_state == "unshifted":
                    unshifted_n += 1
                else:
                    unknown_n += 1
            else:
                refreshed_n += 1

    summary = {
        "deck": deck_key,
        "shift": cfg["label"],
        "ds_overrides": len(overrides),
        "ds_skipped": len(skipped_ds),
        "totals": {
            "charts": chart_total,
            "tables": table_total,
            "components": chart_total + table_total,
            "refreshed": refreshed_n,
            "welded_shifted": welded_shifted_n,
            "unshifted": unshifted_n,
            "unknown": unknown_n,
            "drift_on_shifted": len(drift_on_shifted),
            "drift_on_unshifted": len(drift_on_unshifted),
            "drift_unknown": len(drift_unknown),
            "no_match": len(no_match),
        },
        "drift_on_shifted_detail": [
            {"slide": s, "pos": str(p), "kind": k, "reason": r}
            for s, p, k, r in drift_on_shifted
        ],
        "drift_on_unshifted_detail": [
            {"slide": s, "pos": str(p), "kind": k, "reason": r}
            for s, p, k, r in drift_on_unshifted
        ],
        "no_match_detail": [
            {"slide": s, "pos": str(p), "kind": k} for s, p, k in no_match
        ],
    }
    (work / "summary.json").write_text(
        json.dumps(summary, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    print(f"\n=== {deck_key} | {cfg['label']} ===")
    print(f"  ds shifted/skipped: {len(overrides)}/{len(skipped_ds)}")
    print(f"  components: {chart_total} charts + {table_total} tables = "
          f"{chart_total + table_total}")
    print(f"  refreshed       (shifted ds, values changed): {refreshed_n}")
    print(f"  welded_shifted  (shifted ds, values == src):  {welded_shifted_n}")
    print(f"  unshifted       (ds skipped, values == src):  {unshifted_n}")
    print(f"  unknown         (ds could not be matched):    {unknown_n}")
    print(f"  drift_on_shifted   (Step 6 wave-shift failure):  {len(drift_on_shifted)}")
    print(f"  drift_on_unshifted (pre-existing Step 2 noise):  {len(drift_on_unshifted)}")
    print(f"  drift_unknown                                    {len(drift_unknown)}")
    print(f"  no-match        (missing in ref):                {len(no_match)}")
    shiftable_total = refreshed_n + welded_shifted_n
    if shiftable_total:
        print(
            f"  shift success rate: {refreshed_n}/{shiftable_total} = "
            f"{100 * refreshed_n / shiftable_total:.1f}% of shiftable components refreshed"
        )
    if drift_on_shifted:
        print("  drift_on_shifted detail (gates the test, first 10):")
        for s, p, k, r in drift_on_shifted[:10]:
            print(f"    slide {s} {p} [{k}]: {r}")
    if drift_on_unshifted:
        print(f"  drift_on_unshifted: {len(drift_on_unshifted)} pre-existing Step 2 issues, not gating")

    # Gate ONLY on drift caused by the wave shift itself. Drift on
    # data_sources we couldn't shift (and so refreshed against base config)
    # is a pre-existing Step 2 same-wave roundtrip issue, not a Step 6
    # wave-shift bug — fail the test on it would punish this eval for
    # someone else's known problem.
    assert not drift_on_shifted, (
        f"{len(drift_on_shifted)} structural drift(s) caused by the wave "
        f"shift — see {work / 'summary.json'}"
    )
    assert not no_match, (
        f"{len(no_match)} source components with no refreshed counterpart"
    )
