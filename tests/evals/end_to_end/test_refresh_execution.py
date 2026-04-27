"""Eval #3 — Refresh execution regression test.

Runs component-level comparison between source and refreshed decks and diffs
the resulting match report against a committed golden. This catches
regressions in the refresh pipeline that Path A/B spec-extraction evals
cannot see, because those evals only test spec READING, not pipeline
EXECUTION.

Deeper than Vijay's Stage 3 (categories-only): also checks series names and
series values. The wave-filter + label-transform bug (~155 charts on ATU)
was silent under categories-only comparison; values_match would have caught it.

Mutation test corrupts the golden in-memory and verifies compare_reports flags
it — proves the comparison is not always-green.

To update the golden intentionally (e.g., after re-running Vijay's refresh
pipeline on a deck):
    python -m tests.evals.end_to_end.generate_golden

Usage:
    pytest tests/evals/end_to_end/ -v
"""
from __future__ import annotations

import copy
import json
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT))

from tests.evals.end_to_end.compare_decks import compare_decks, compare_reports  # noqa: E402
from tests.evals.end_to_end.generate_golden import (  # noqa: E402
    GOLDEN_DIR,
    classify_chart_modes,
    connected_slide_indices_for,
)
from tests.evals.fixtures import FIXTURE_DECKS, REFRESHED_DECKS  # noqa: E402


@pytest.mark.parametrize("deck_key", list(REFRESHED_DECKS.keys()))
def test_refresh_matches_golden(deck_key: str):
    source_path = FIXTURE_DECKS.get(deck_key)
    refreshed_path = REFRESHED_DECKS.get(deck_key)

    if source_path is None or not source_path.exists():
        pytest.skip(f"Source deck not on disk: {source_path}")
    if refreshed_path is None or not refreshed_path.exists():
        pytest.skip(
            f"Refreshed fixture not on disk: {refreshed_path}\n"
            "Generate by running Vijay's pipeline on this deck, or see"
            " tests/test_spec_refresh_pipeline.py for the 3-stage flow."
        )

    golden_file = GOLDEN_DIR / f"{deck_key}_refresh.json"
    assert golden_file.exists(), (
        f"Golden file missing: {golden_file}. "
        "Run `python -m tests.evals.end_to_end.generate_golden` first."
    )
    golden = json.loads(golden_file.read_text(encoding="utf-8"))

    connected = connected_slide_indices_for(deck_key)
    chart_modes = classify_chart_modes(deck_key)
    report = compare_decks(
        source_path, refreshed_path,
        connected_slide_indices=connected,
        chart_modes=chart_modes,
    )
    actual = report.to_dict()

    errors = compare_reports(actual, golden)

    if errors:
        lines = [
            f"[{deck_key}] {len(errors)} regression(s) vs golden:",
            f"  golden  totals: {golden['totals']}",
            f"  current totals: {actual['totals']}",
        ]
        for err in errors[:20]:
            lines.append(f"  {err}")
        if len(errors) > 20:
            lines.append(f"  ... and {len(errors)-20} more")
        pytest.fail("\n".join(lines))


def test_mutation_is_caught():
    """Sanity check: corrupt the golden in-memory, verify compare_reports flags it.

    Uses the first available refreshed fixture. Skips if none are on disk.
    """
    target_key = None
    target_golden = None
    for key in REFRESHED_DECKS:
        golden_file = GOLDEN_DIR / f"{key}_refresh.json"
        if golden_file.exists():
            target_key = key
            target_golden = json.loads(golden_file.read_text(encoding="utf-8"))
            break

    if target_golden is None:
        pytest.skip("No Eval #3 goldens on disk — run generate_golden first.")

    assert target_golden["components"], f"Golden for {target_key} has no components"

    # Mutation: pretend the "actual" run regressed on the first passing component
    # The golden says True; we mutate a deep copy to say False → compare should flag it.
    mutated_actual = copy.deepcopy(target_golden)
    flipped = False
    for comp in mutated_actual["components"]:
        for signal in ("categories_match", "series_names_match", "values_match"):
            if comp.get(signal) is True:
                comp[signal] = False
                flipped = True
                # also decrement the total so the totals-regression path triggers too
                mutated_actual["totals"][signal] = max(
                    0, mutated_actual["totals"][signal] - 1
                )
                break
        if flipped:
            break

    if not flipped:
        pytest.skip(
            f"Golden for {target_key} has no passing signals to mutate — "
            "refresh may have produced a baseline of all-fail."
        )

    # Compare the mutated "actual" against the pristine golden.
    errors = compare_reports(mutated_actual, target_golden)
    assert len(errors) > 0, (
        "Mutation test failed: compare_reports did not detect a regression. "
        "The comparison logic is not sensitive — eval would always pass."
    )


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-v"]))
