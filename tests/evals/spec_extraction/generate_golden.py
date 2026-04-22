"""One-time golden file generator for Path A (connector-tag spec extraction).

Runs `generate_config_specs()` on each fixture deck and serializes the full
spec list as the golden snapshot. The eval in test_path_a.py diffs future
extraction output against these goldens.

Run to (re)generate:
    python -m tests.evals.spec_extraction.generate_golden

Re-run only when Path A intentionally changes (and review the diff before
committing the updated goldens).
"""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT))

from slidegen.deck_reader.tag_reader import generate_config_specs  # noqa: E402
from slidegen.slide_spec.schema import dump_spec  # noqa: E402
from tests.evals.fixtures import FIXTURE_DECKS  # noqa: E402

GOLDEN_DIR = Path(__file__).parent / "goldens"


def build_specs_snapshot(deck_path: Path) -> list[dict[str, Any]]:
    """Run Path A on a deck and return the spec list as JSON-serializable dicts."""
    specs, _summary = generate_config_specs(str(deck_path))
    return [json.loads(dump_spec(s)) for s in specs]


def compare_specs(
    actual: list[dict[str, Any]],
    expected: list[dict[str, Any]],
) -> list[str]:
    """Diff two spec lists and return a list of drift messages (empty = match).

    Reports differences at the slide level. For each slide that differs, lists
    the top-level keys that changed.
    """
    errors: list[str] = []

    if len(actual) != len(expected):
        errors.append(
            f"spec count: expected {len(expected)}, got {len(actual)}"
        )

    for i in range(min(len(actual), len(expected))):
        a, e = actual[i], expected[i]
        if a == e:
            continue
        slide_id = a.get("slide_id") or e.get("slide_id") or f"index_{i}"
        diff_keys = sorted(
            k for k in set(a.keys()) | set(e.keys())
            if a.get(k) != e.get(k)
        )
        errors.append(
            f"slide {slide_id} (index {i}): differs in keys {diff_keys}"
        )

    if len(actual) > len(expected):
        for i in range(len(expected), len(actual)):
            slide_id = actual[i].get("slide_id", f"index_{i}")
            errors.append(f"slide {slide_id} (index {i}): NEW (not in golden)")

    if len(expected) > len(actual):
        for i in range(len(actual), len(expected)):
            slide_id = expected[i].get("slide_id", f"index_{i}")
            errors.append(f"slide {slide_id} (index {i}): MISSING (in golden, not in actual)")

    return errors


def main():
    GOLDEN_DIR.mkdir(parents=True, exist_ok=True)

    for key, path in FIXTURE_DECKS.items():
        if not path.exists():
            print(f"[skip] {key}: deck not found at {path}")
            continue
        print(f"[gen]  {key}: reading {path.name}")
        snapshot = build_specs_snapshot(path)
        out_file = GOLDEN_DIR / f"{key}_spec.json"
        out_file.write_text(
            json.dumps(snapshot, indent=2, sort_keys=True),
            encoding="utf-8",
        )
        size_kb = out_file.stat().st_size / 1024
        print(f"       wrote {out_file.relative_to(REPO_ROOT)} ({len(snapshot)} specs, {size_kb:.0f} KB)")


if __name__ == "__main__":
    main()
