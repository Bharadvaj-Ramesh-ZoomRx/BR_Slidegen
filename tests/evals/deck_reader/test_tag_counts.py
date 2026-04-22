"""Eval #1 — Tag count per slide (regression test).

Runs deck-reader on fixture decks and verifies that the tagged/untagged shape
counts per slide match the committed golden snapshot.

Fails loud if any count drifts. To update the golden intentionally, re-run
`python -m tests.evals.deck_reader.generate_golden` and commit the diff.

Usage:
    pytest tests/evals/deck_reader/test_tag_counts.py -v
    # or:
    python -m tests.evals.deck_reader.test_tag_counts
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT))

from tests.evals.deck_reader.generate_golden import build_tag_counts  # noqa: E402
from tests.evals.fixtures import FIXTURE_DECKS, GOLDEN_DIR  # noqa: E402


@pytest.mark.parametrize("deck_key", list(FIXTURE_DECKS.keys()))
def test_tag_counts_match_golden(deck_key: str):
    deck_path = FIXTURE_DECKS[deck_key]
    if not deck_path.exists():
        pytest.skip(f"Fixture deck not on disk: {deck_path}")

    golden_file = GOLDEN_DIR / f"{deck_key}_tag_counts.json"
    assert golden_file.exists(), (
        f"Golden file missing: {golden_file}. "
        "Run `python -m tests.evals.deck_reader.generate_golden` first."
    )
    golden = json.loads(golden_file.read_text(encoding="utf-8"))

    actual = build_tag_counts(deck_path)

    errors = []

    # ── Totals ──
    for key, expected in golden["totals"].items():
        got = actual["totals"].get(key)
        if got != expected:
            errors.append(f"totals.{key}: expected {expected}, got {got}")

    # ── Per-slide drift ──
    for slide_num, expected_counts in golden["per_slide"].items():
        got_counts = actual["per_slide"].get(slide_num)
        if got_counts != expected_counts:
            errors.append(
                f"slide {slide_num}: expected {expected_counts}, got {got_counts}"
            )

    # ── New slides not in golden ──
    extra = set(actual["per_slide"]) - set(golden["per_slide"])
    for slide_num in sorted(extra, key=int):
        errors.append(f"slide {slide_num}: NEW (not in golden), got {actual['per_slide'][slide_num]}")

    if errors:
        lines = [f"[{deck_key}] {len(errors)} drift(s) vs golden:"]
        for err in errors[:20]:
            lines.append(f"  {err}")
        if len(errors) > 20:
            lines.append(f"  ... and {len(errors)-20} more")
        pytest.fail("\n".join(lines))


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-v"]))
