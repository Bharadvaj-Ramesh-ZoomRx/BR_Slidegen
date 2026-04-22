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

        print(f"[gen]  {key}: comparing {source_path.name} vs {refreshed_path.name}")
        report = compare_decks(source_path, refreshed_path)
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
