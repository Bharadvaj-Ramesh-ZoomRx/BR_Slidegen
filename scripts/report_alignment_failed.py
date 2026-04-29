"""Analyst-facing report: slides where charts need a Synapse-side fix.

After today's mapper fix, the refresh pipeline correctly preserves source
values when it can't align API columns to source chart series — and emits
a clear "alignment_failed" status in refresh_status.json. Those slides
are the "user must fix Synapse setup" cases (e.g. 5 API rating columns
need to be bucketed into 3 source series via custom virtual question;
neither connected nor non-connected pipeline can recover the bucketing
recipe automatically).

This script reads the latest refresh_status.json for each deck and lists
the affected slides + chart names so an analyst knows which Synapse
configurations to revisit.

Usage:
    python scripts/report_alignment_failed.py
"""
from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
BASE = REPO / "output" / "step2_test_connected"

DECKS = ("atu_q1_26", "creon_pet_w33")


def report(deck_key: str) -> None:
    path = BASE / f"{deck_key}_refresh_status.json"
    if not path.exists():
        print(f"\n=== {deck_key} ===")
        print(f"  refresh_status.json not on disk: {path}")
        return

    data = json.loads(path.read_text(encoding="utf-8"))

    by_slide: dict[int, list[tuple[str, str]]] = defaultdict(list)
    counts: dict[str, int] = defaultdict(int)

    for s in data.get("slides", []):
        s_idx = s["slide_index"]
        for c in s.get("charts", []) + s.get("tables", []):
            status = c.get("status", "?")
            counts[status] += 1
            if status == "alignment_failed":
                by_slide[s_idx].append((c.get("name", "?"), status))

    n_total = sum(counts.values())
    n_aligned = counts.get("ok", 0) + counts.get("static_pinned_skipped", 0)
    print(f"\n=== {deck_key} ===")
    print(f"  Components: {n_total}")
    for k in ("ok", "static_pinned_skipped", "alignment_failed",
             "no_data_for_shifted_window", "empty"):
        if counts.get(k):
            print(f"    {k}: {counts[k]}")

    if not by_slide:
        print("  No alignment_failed components — nothing requires Synapse-side fix.")
        return

    n_charts = sum(len(v) for v in by_slide.values())
    print(
        f"\n  [!] {n_charts} chart(s) on {len(by_slide)} slide(s) need a "
        f"Synapse-side fix:"
    )
    print(
        "    (refresh preserved source values for these — they're flagged "
        "so an analyst can\n"
        "     create a virtual question on the Synapse side that returns "
        "the chart's expected\n"
        "     bucketed columns directly, or update the Connector tag to "
        "reference one)\n"
    )
    for s_idx in sorted(by_slide):
        items = by_slide[s_idx]
        first_names = ", ".join(n for n, _ in items[:3])
        more = f" (+{len(items) - 3} more)" if len(items) > 3 else ""
        print(f"    slide {s_idx + 1}:  {first_names}{more}")


for deck_key in DECKS:
    report(deck_key)
