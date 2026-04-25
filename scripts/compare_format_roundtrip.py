"""Run Step 3 formatting comparison for an Eval #3 deck key and print totals.

Usage:
    python scripts/compare_format_roundtrip.py <deck_key>

Reads source from tests.evals.fixtures.FIXTURE_DECKS and refreshed from
output/<deck_key>_refresh/<deck_key>_refreshed.pptx, then prints
chart_type / series_colors match counts and writes a per-component
report next to the existing roundtrip JSON.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))


def main():
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(2)

    deck_key = sys.argv[1]

    from tests.evals.fixtures import FIXTURE_DECKS
    from tests.evals.end_to_end.compare_formatting import compare_formatting

    source = FIXTURE_DECKS[deck_key]
    refreshed = REPO_ROOT / "output" / f"{deck_key}_refresh" / f"{deck_key}_refreshed.pptx"

    if not source.exists():
        print(f"ERROR: source not found: {source}")
        sys.exit(2)
    if not refreshed.exists():
        print(f"ERROR: refreshed not found: {refreshed}")
        sys.exit(2)

    print(f"Comparing formatting:\n  source:    {source.name}\n  refreshed: {refreshed.name}\n")
    report = compare_formatting(source, refreshed)
    totals = report.totals()
    n = totals["components_compared"]
    type_ok = totals["chart_type_match"]
    col_ok = totals["series_colors_match"]

    print(f"=== {deck_key} (Step 3 — formatting) ===")
    print(f"Charts compared: {n}")
    print(f"  chart_type_match:    {type_ok}/{n} ({100*type_ok/max(n,1):.1f}%)")
    print(f"  series_colors_match: {col_ok}/{n} ({100*col_ok/max(n,1):.1f}%)")
    print(f"  Extraction errors:   {totals['components_with_errors']}")

    out = REPO_ROOT / "output" / f"{deck_key}_refresh" / f"{deck_key}_format_2026-04-25.json"
    out.write_text(json.dumps(report.to_dict(), indent=2), encoding="utf-8")
    print(f"\nFull report: {out.relative_to(REPO_ROOT)}")


if __name__ == "__main__":
    main()
