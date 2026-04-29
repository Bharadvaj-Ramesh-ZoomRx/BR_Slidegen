"""List slides where the refresh+headline pair doesn't fit the expected contract."""
from __future__ import annotations
import json
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]


def classify(refresh_slide, headline_update):
    if refresh_slide is None and (headline_update is None or headline_update.get("status") == "not_in_spec"):
        return "untagged"

    components = (refresh_slide or {}).get("charts", []) + (refresh_slide or {}).get("tables", [])
    statuses = [c.get("status", "?") for c in components]
    n = len(statuses) or 1
    n_ok = sum(1 for s in statuses if s == "ok")
    n_static = sum(1 for s in statuses if s == "static_pinned_skipped")
    n_nodata = sum(1 for s in statuses if s == "no_data_for_shifted_window")
    n_empty = sum(1 for s in statuses if s == "empty")

    if n_ok == n:
        r = "ok"
    elif n_static == n:
        r = "static"
    elif n_nodata == n:
        r = "no_data"
    elif n_empty == n:
        r = "empty"
    else:
        r = f"mixed(ok={n_ok},static={n_static},nodata={n_nodata},empty={n_empty})"

    h = (headline_update or {}).get("status", "missing")
    return f"refresh={r} | headline={h}"


for deck in ("atu_q1_26", "creon_pet_w33"):
    base = REPO / "output" / "step2_test_connected"
    refresh = json.loads((base / f"{deck}_refresh_status.json").read_text(encoding="utf-8"))
    headline = json.loads((base / f"{deck}_with_headlines_status.json").read_text(encoding="utf-8"))

    refresh_by_idx = {s["slide_index"]: s for s in refresh["slides"]}
    headline_by_idx = {u["slide_idx"]: u for u in headline["updates"]}

    print(f"\n========== {deck} ==========")
    all_idx = set(refresh_by_idx) | set(headline_by_idx)
    by_class = {}
    for s_idx in sorted(all_idx):
        cls = classify(refresh_by_idx.get(s_idx), headline_by_idx.get(s_idx))
        by_class.setdefault(cls, []).append(s_idx)

    for cls in sorted(by_class):
        idx_str = ", ".join(str(i + 1) for i in by_class[cls][:30])
        if len(by_class[cls]) > 30:
            idx_str += f"  ...(+{len(by_class[cls]) - 30} more)"
        print(f"  [{len(by_class[cls]):3d}] {cls}")
        print(f"        slides (1-indexed): {idx_str}")
