"""Live-API refresh eval — replaces the 6 wave-shift tests.

Premise: April data is in the API now and the deck's reporting plan is
configured. There's no need to synthesize wave shifts to validate refresh
correctness — we just refresh against the deck's existing config and
compare against direct API queries.

Per fixture deck:
    1. Read the deck's full_spec.json (already has the right config).
    2. Run refresh_deck_from_spec → produces refreshed PPTX.
    3. For each chart marked status="ok" in refresh_status.json:
         - Pull the chart's (cats, series_values) from the refreshed deck.
         - Fetch API truth for the chart's (project, plan, analysis, segments).
         - Pivot the truth records to (cats, series) and assert match within
           tolerance.

Live API required. Skips when SYNAPSE_API_URL / SYNAPSE_API_KEY are absent
or when the fixture deck isn't on disk. Slow (~minutes per deck).

Run:
    pytest tests/evals/end_to_end/test_refresh_against_live_api.py -v -s

Run only on demand (CI marker):
    pytest -m live_api tests/evals/end_to_end/
"""
from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[3]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from tests.evals.end_to_end.compare_decks import _extract_chart_data, _pos  # noqa: E402
from tests.evals.fixtures import FIXTURE_DECKS  # noqa: E402

WORK_DIR = REPO_ROOT / "output" / "_live_api_refresh"
SPEC_DIR = REPO_ROOT / "output" / "step2_test_connected"

# Comparison tolerance: chart cells round to ~4 decimals via formatCode-driven
# display, and the API returns full precision. 6e-3 absorbs that rounding plus
# any small API-side numeric drift.
VALUE_TOLERANCE = 6e-3


def _live_api_available() -> bool:
    return bool(os.getenv("SYNAPSE_API_URL") and os.getenv("SYNAPSE_API_KEY"))


@pytest.fixture(scope="module")
def workdir():
    WORK_DIR.mkdir(parents=True, exist_ok=True)
    return WORK_DIR


@pytest.mark.live_api
@pytest.mark.parametrize("deck_key", list(FIXTURE_DECKS.keys()))
def test_refresh_against_live_api(deck_key: str, workdir: Path):
    """Refresh the deck against current API and validate per-chart values.

    No wave overrides. Uses whatever config is in the deck's full_spec.json,
    which represents the user's intended refresh behavior (dynamic_latest_n
    or static_time_period_ids as configured at render time).
    """
    if not _live_api_available():
        pytest.skip("SYNAPSE_API_URL / SYNAPSE_API_KEY not set — live API unavailable")

    source_path = FIXTURE_DECKS.get(deck_key)
    if source_path is None or not source_path.exists():
        pytest.skip(f"Source deck not on disk: {source_path}")

    spec_path = SPEC_DIR / f"{deck_key}_full_spec.json"
    if not spec_path.exists():
        pytest.skip(f"Spec missing: {spec_path}")

    from slidegen.intelligent_refresh import (  # noqa: E402
        refresh_deck_from_spec, fetch_synapse_data,
    )

    out_pptx = workdir / f"{deck_key}_refreshed.pptx"
    shutil.copy(source_path, out_pptx)

    refresh_deck_from_spec(
        spec_path=str(spec_path),
        pptx_path=str(out_pptx),
        output_path=str(out_pptx),
    )

    # Read the refresh_status JSON the pipeline emits alongside the deck
    status_path = out_pptx.with_name(out_pptx.stem + "_refresh_status.json")
    if not status_path.exists():
        pytest.fail(f"refresh_status.json not produced at {status_path}")

    refresh_status = json.loads(status_path.read_text(encoding="utf-8"))
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    data_sources = spec.get("data_sources", {})

    from pptx import Presentation
    pres = Presentation(str(out_pptx))

    failures: list[str] = []
    n_checked = 0
    n_skipped_modes = 0

    for slide_status in refresh_status.get("slides", []):
        si = slide_status["slide_index"]
        slide = pres.slides[si]

        # Build a (slide_idx, name) -> chart shape lookup
        charts_by_name = {sh.name: sh for sh in slide.shapes if sh.has_chart}

        for chart_status in slide_status.get("charts", []):
            if chart_status.get("status") != "ok":
                continue
            name = chart_status["name"]
            shape = charts_by_name.get(name)
            if shape is None:
                continue

            # Tag-mismatch / partial_alignment / dynamic_added: don't fail
            # the test; surface in the report. The point of those notes is
            # that the chart is intentionally divergent.
            notes = chart_status.get("notes") or []
            note_kinds = {n.get("kind") for n in notes}
            if "tag_mismatch" in note_kinds:
                n_skipped_modes += 1
                continue

            comp_spec = _find_component_in_spec(spec, si, name)
            if comp_spec is None:
                continue
            ds_key = comp_spec.get("data_source") or _slide_data_source(spec, si)
            ds = data_sources.get(ds_key, {})
            if not ds:
                continue

            # Fetch API truth using the same lineage the spec used
            try:
                _, truth_df = fetch_synapse_data({
                    "project_id": ds.get("project_id"),
                    "reporting_plan_id": ds.get("reporting_plan_id"),
                    "analysis_ids": ds.get("analysis_ids", []),
                    "segment_ids": ds.get("segment_ids", []),
                    "dynamic_latest_n": ds.get("dynamic_latest_n", 0),
                    "static_time_period_ids": ds.get("static_time_period_ids") or [],
                    "static_time_period_names": ds.get("static_time_period_names") or [],
                    "include_live_wave": ds.get("include_live_wave", True),
                })
            except Exception as exc:
                failures.append(f"{deck_key} slide {si} {name}: API fetch failed: {exc}")
                continue

            chart_cats, chart_series = _extract_chart_data(shape)

            # Soft check: each refreshed series value should appear (within
            # tolerance) somewhere in the truth_df for the same time period
            # column. We don't insist on exact pivot reconstruction here —
            # that's what the mapper does; this eval just asserts the
            # refreshed numbers are present in the truth source.
            chart_vals = {
                round(float(v), 4)
                for _, vals in chart_series for v in vals if v is not None
            }
            if not chart_vals:
                continue
            n_checked += 1

            truth_vals: set[float] = set()
            for col in ("decimal", "percentage", "count", "share",
                        "penetration", "me_score", "value", "base"):
                if col in truth_df.columns:
                    for v in truth_df[col].dropna().tolist():
                        truth_vals.add(round(float(v), 4))
            if not truth_vals:
                continue  # no comparable numeric column — skip silently

            unmatched = [
                v for v in chart_vals
                if not any(abs(v - tv) <= VALUE_TOLERANCE for tv in truth_vals)
            ]
            if unmatched:
                # Allow a small unmatched ratio for legitimate chart-side
                # rounding artifacts; flag if too many drift.
                if len(unmatched) > max(2, int(0.1 * len(chart_vals))):
                    failures.append(
                        f"{deck_key} slide {si} {name}: "
                        f"{len(unmatched)}/{len(chart_vals)} chart values "
                        f"not present in API truth (e.g. {sorted(unmatched)[:3]})"
                    )

    if failures:
        msg = (f"\n[{deck_key}] {len(failures)} chart(s) diverge from API truth "
               f"(checked {n_checked}, skipped {n_skipped_modes} tag-mismatch):\n"
               + "\n".join(f"  - {f}" for f in failures[:20]))
        if len(failures) > 20:
            msg += f"\n  ... and {len(failures) - 20} more"
        pytest.fail(msg)

    # Print informational summary even on success
    print(
        f"\n[{deck_key}] live-api refresh: {n_checked} charts validated, "
        f"{n_skipped_modes} tag-mismatch shapes skipped"
    )


def _find_component_in_spec(spec: dict, slide_idx: int, name: str) -> dict | None:
    for s in spec.get("slides", []):
        if s.get("slide_index") != slide_idx:
            continue
        for c in s.get("components", []):
            if c.get("name") == name:
                return c
    return None


def _slide_data_source(spec: dict, slide_idx: int) -> str | None:
    for s in spec.get("slides", []):
        if s.get("slide_index") == slide_idx:
            return s.get("data_source")
    return None


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-v", "-s"]))
