"""Eval #2 — Path A (connector-tag spec extraction) regression test.

Runs `generate_config_specs()` on each fixture deck and diffs the output
against the committed golden snapshot. Fails loud on any drift.

Includes a mutation test that proves the comparison logic is sensitive —
corrupts the golden in-memory and asserts the comparison reports drift.

To update the golden intentionally:
    python -m tests.evals.spec_extraction.generate_golden

Usage:
    pytest tests/evals/spec_extraction/ -v
"""
from __future__ import annotations

import copy
import json
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT))

from tests.evals.fixtures import FIXTURE_DECKS  # noqa: E402
from tests.evals.spec_extraction.generate_golden import (  # noqa: E402
    GOLDEN_DIR,
    build_specs_snapshot,
    compare_specs,
)


@pytest.mark.parametrize("deck_key", list(FIXTURE_DECKS.keys()))
def test_path_a_matches_golden(deck_key: str):
    deck_path = FIXTURE_DECKS[deck_key]
    if not deck_path.exists():
        pytest.skip(f"Fixture deck not on disk: {deck_path}")

    golden_file = GOLDEN_DIR / f"{deck_key}_spec.json"
    assert golden_file.exists(), (
        f"Golden file missing: {golden_file}. "
        "Run `python -m tests.evals.spec_extraction.generate_golden` first."
    )
    golden = json.loads(golden_file.read_text(encoding="utf-8"))

    actual = build_specs_snapshot(deck_path)

    errors = compare_specs(actual, golden)

    if errors:
        lines = [f"[{deck_key}] {len(errors)} drift(s) vs golden:"]
        for err in errors[:20]:
            lines.append(f"  {err}")
        if len(errors) > 20:
            lines.append(f"  ... and {len(errors)-20} more")
        pytest.fail("\n".join(lines))


def test_mutation_is_caught():
    """Sanity check: corrupt the golden in-memory, verify compare_specs flags it.

    Proves the comparison is actually sensitive — not always-green. Uses the
    first fixture deck with an on-disk golden; skips if none are available.
    """
    # Pick any fixture with a golden on disk
    target_key = None
    target_golden = None
    for key in FIXTURE_DECKS:
        golden_file = GOLDEN_DIR / f"{key}_spec.json"
        if golden_file.exists():
            target_key = key
            target_golden = json.loads(golden_file.read_text(encoding="utf-8"))
            break

    if target_golden is None:
        pytest.skip("No goldens on disk — run generate_golden first.")

    assert len(target_golden) > 0, f"Golden for {target_key} is empty"

    # Mutation: corrupt one slide's slide_id in a deep copy
    corrupted = copy.deepcopy(target_golden)
    corrupted[0]["slide_id"] = "MUTATED_SLIDE_ID"

    # The actual output (which matches the original golden) should differ from the corrupted version
    errors = compare_specs(target_golden, corrupted)
    assert len(errors) > 0, (
        "Mutation test failed: compare_specs did not detect corrupted golden. "
        "The comparison logic is not sensitive — eval would always pass."
    )


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-v"]))
