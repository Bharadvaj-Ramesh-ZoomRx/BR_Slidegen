"""Run source-vs-refreshed comparison for Eval #3 deck keys and print totals.

Usage:
    python scripts/compare_refresh_roundtrip.py <deck_key>

Reads source from tests.evals.fixtures.FIXTURE_DECKS and refreshed from
output/<deck_key>_refresh/<deck_key>_refreshed.pptx (the path written by
gen_refreshed_fixture.py), then prints per-deck totals.
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
    from tests.evals.end_to_end.compare_decks import compare_decks

    source = FIXTURE_DECKS[deck_key]
    refreshed = REPO_ROOT / "output" / f"{deck_key}_refresh" / f"{deck_key}_refreshed.pptx"

    if not source.exists():
        print(f"ERROR: source not found: {source}")
        sys.exit(2)
    if not refreshed.exists():
        print(f"ERROR: refreshed not found: {refreshed}")
        sys.exit(2)

    print(f"Comparing:\n  source:    {source.name}\n  refreshed: {refreshed.name}\n")
    report = compare_decks(source, refreshed)
    totals = report.totals()

    charts = [c for c in report.components if c.kind == "chart"]
    tables = [c for c in report.components if c.kind == "table"]

    chart_cat = sum(1 for c in charts if c.categories_match is True)
    chart_srs = sum(1 for c in charts if c.series_names_match is True)
    chart_val = sum(1 for c in charts if c.values_match is True)
    tbl_val = sum(1 for c in tables if c.values_match is True)

    print(f"=== {deck_key} ===")
    print(f"Components compared: {totals['components_compared']}")
    print(f"  Charts: {len(charts)}")
    print(f"    categories_match:   {chart_cat}/{len(charts)} ({100*chart_cat/max(len(charts),1):.1f}%)")
    print(f"    series_names_match: {chart_srs}/{len(charts)} ({100*chart_srs/max(len(charts),1):.1f}%)")
    print(f"    values_match:       {chart_val}/{len(charts)} ({100*chart_val/max(len(charts),1):.1f}%)")
    print(f"  Tables: {len(tables)}")
    print(f"    values_match:       {tbl_val}/{len(tables)} ({100*tbl_val/max(len(tables),1):.1f}%)")
    print(f"  Extraction errors: {totals['components_with_errors']}")

    out = REPO_ROOT / "output" / f"{deck_key}_refresh" / f"{deck_key}_roundtrip_2026-04-24.json"
    out.write_text(json.dumps(report.to_dict(), indent=2), encoding="utf-8")
    print(f"\nFull report: {out.relative_to(REPO_ROOT)}")


if __name__ == "__main__":
    main()
