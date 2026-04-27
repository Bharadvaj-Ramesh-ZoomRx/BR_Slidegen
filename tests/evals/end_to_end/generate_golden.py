"""One-time golden generator for Eval #3 (refresh execution).

Runs compare_decks(source, refreshed) on each fixture pair and writes the
per-component match report as the golden. The eval diffs future runs
against this.

Run to (re)generate:
    python -m tests.evals.end_to_end.generate_golden

Re-run deliberately when:
  - the source or refreshed deck has been regenerated on purpose
  - the comparison logic in compare_decks.py has been tightened/relaxed

Review the JSON diff before committing updated goldens.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT))

from tests.evals.end_to_end.compare_decks import compare_decks  # noqa: E402
from tests.evals.fixtures import FIXTURE_DECKS, REFRESHED_DECKS  # noqa: E402

GOLDEN_DIR = Path(__file__).parent / "goldens"
SPEC_GOLDEN_DIR = Path(__file__).resolve().parents[1] / "spec_extraction" / "goldens"
FULL_SPEC_DIR = Path(__file__).resolve().parents[3] / "output" / "step2_test_connected"


def connected_slide_indices_for(deck_key: str) -> set[int] | None:
    """Return the set of 0-based slide indices that have Connector tags.

    Loaded from the Step 1 spec golden (connected slides only). Returns None
    if the spec golden doesn't exist yet, which causes compare_decks to fall
    back to comparing all slides.
    """
    spec_golden = SPEC_GOLDEN_DIR / f"{deck_key}_spec.json"
    if not spec_golden.exists():
        return None
    specs = json.loads(spec_golden.read_text(encoding="utf-8"))
    return {s["slide_index"] for s in specs}


def classify_chart_modes(deck_key: str) -> dict[tuple[int, str], str] | None:
    """Read the deck's full_spec.json and classify each chart as 'dynamic' or 'static'.

    Used by compare_decks to apply the right success criterion per chart:
      - 'static':  strict identity comparison (source values == refreshed values).
                   Includes charts with static_time_period_ids set, OR charts
                   without dynamic_latest_n.
      - 'dynamic': structural comparison only. Wave labels legitimately differ
                   after auto-roll-forward (Wave 11+12 -> Wave 12+13 is correct
                   product behavior, not a refresh bug).

    Returns dict mapping (slide_index, chart_shape_name) -> mode.
    Returns None if the full_spec is missing — compare_decks then falls back
    to strict comparison for everything (current behavior).
    """
    full_spec_path = FULL_SPEC_DIR / f"{deck_key}_full_spec.json"
    if not full_spec_path.exists():
        return None
    spec = json.loads(full_spec_path.read_text(encoding="utf-8"))
    data_sources = spec.get("data_sources", {})
    modes: dict[tuple[int, str], str] = {}
    for slide in spec.get("slides", []):
        si = slide["slide_index"]
        slide_ds_key = slide.get("data_source")
        for comp in slide.get("components", []):
            if comp.get("type") != "chart":
                continue
            name = comp.get("name") or ""
            if not name:
                continue
            ds_key = comp.get("data_source") or slide_ds_key
            ds = data_sources.get(ds_key, {})
            sids = ds.get("static_time_period_ids") or []
            snames = ds.get("static_time_period_names") or []
            dyn = ds.get("dynamic_latest_n") or 0
            if sids or snames or not dyn:
                modes[(si, name)] = "static"
            else:
                modes[(si, name)] = "dynamic"
    return modes


def main():
    GOLDEN_DIR.mkdir(parents=True, exist_ok=True)

    for key, source_path in FIXTURE_DECKS.items():
        refreshed_path = REFRESHED_DECKS.get(key)
        if refreshed_path is None:
            print(f"[skip] {key}: no refreshed fixture registered")
            continue
        if not source_path.exists():
            print(f"[skip] {key}: source not found at {source_path}")
            continue
        if not refreshed_path.exists():
            print(f"[skip] {key}: refreshed fixture not found at {refreshed_path}")
            continue

        connected = connected_slide_indices_for(key)
        chart_modes = classify_chart_modes(key)
        print(f"[gen]  {key}: comparing {source_path.name} vs {refreshed_path.name}")
        if connected is not None:
            print(f"       restricting to {len(connected)} connected slides")
        if chart_modes is not None:
            n_dyn = sum(1 for m in chart_modes.values() if m == "dynamic")
            n_stat = sum(1 for m in chart_modes.values() if m == "static")
            print(f"       chart modes: {n_dyn} dynamic, {n_stat} static")
        report = compare_decks(
            source_path, refreshed_path,
            connected_slide_indices=connected,
            chart_modes=chart_modes,
        )
        report_dict = report.to_dict()

        out_file = GOLDEN_DIR / f"{key}_refresh.json"
        out_file.write_text(
            json.dumps(report_dict, indent=2, sort_keys=True),
            encoding="utf-8",
        )
        tot = report_dict["totals"]
        size_kb = out_file.stat().st_size / 1024
        print(
            f"       wrote {out_file.relative_to(REPO_ROOT)} "
            f"({size_kb:.0f} KB)"
        )
        print(
            f"       components: {tot['components_compared']}  "
            f"categories: {tot['categories_match']}  "
            f"series: {tot['series_names_match']}  "
            f"values: {tot['values_match']}  "
            f"errors: {tot['components_with_errors']}"
        )


if __name__ == "__main__":
    main()
