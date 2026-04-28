"""Verify the backward wave-shift refresh landed each chart on the
expected previous-available waves.

For each shifted ds, look at every chart on that ds in the refreshed deck
and check whether its category labels (or compound series labels) contain
the wave names corresponding to the shifted_ids. Flags any chart that
ended up showing the SOURCE waves instead of the shifted ones.

Run: python scripts/verify_backward_shift.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pptx import Presentation

from slidegen.intelligent_refresh import fetch_synapse_data

REPO = Path(__file__).resolve().parents[1]


def _wave_names_for_ids(ds: dict, ids: list[str]) -> list[str]:
    """Resolve time_period_id strings to their human wave names via Synapse."""
    _, df = fetch_synapse_data({
        "project_id": ds["project_id"],
        "reporting_plan_id": ds["reporting_plan_id"],
        "analysis_ids": ds["analysis_ids"],
        "segment_ids": ds.get("segment_ids", []),
        "dynamic_latest_n": 0,
        "static_time_period_ids": list(ids),
        "static_time_period_names": [],
        "include_live_wave": False,
    })
    if df.empty or "time_period_name" not in df.columns:
        return []
    out = []
    for sid in ids:
        match = df[df["time_period_id"].astype(str) == str(sid)]
        if not match.empty:
            name = str(match["time_period_name"].iloc[0])
            out.append(name)
            # Also record the "Wave N" stripped form
            stripped = name.replace("Project Wave ", "Wave ")
            if stripped != name:
                out.append(stripped)
    return out


def verify(deck_key: str, project_id: int) -> dict:
    spec_path = REPO / "output" / "_step6" / deck_key / "spec_shift.json"
    refreshed_path = REPO / "output" / "_step6" / deck_key / "refreshed_shift.pptx"
    overrides_path = REPO / "output" / "_step6" / deck_key / "shift_overrides.json"
    if not all(p.exists() for p in (spec_path, refreshed_path, overrides_path)):
        return {"error": f"missing artifact for {deck_key}"}

    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    overrides = json.loads(overrides_path.read_text(encoding="utf-8"))["overrides"]

    # Cache of expected wave-name labels per ds
    expected_names: dict[str, list[str]] = {}
    for ds_key, ids in overrides.items():
        ds = spec["data_sources"][ds_key]
        if ds.get("project_id") != project_id:
            continue
        expected_names[ds_key] = _wave_names_for_ids(ds, ids)

    # Map each component name → its ds (only shifted ds)
    comp_ds_by_name: dict[tuple, str] = {}
    comp_ds_by_pos: dict[int, list[tuple[float, float, str]]] = {}
    for slide in spec.get("slides", []):
        s_idx = slide["slide_index"]
        slide_default = slide.get("data_source")
        for comp in slide.get("components", []):
            ds = comp.get("data_source") or slide_default
            if ds not in expected_names:
                continue
            name = comp.get("name") or ""
            if name:
                comp_ds_by_name[(s_idx, name)] = ds
            cpos = comp.get("position", {}) or {}
            cl = float(cpos.get("left", 0) or 0)
            ct = float(cpos.get("top", 0) or 0)
            comp_ds_by_pos.setdefault(s_idx, []).append((cl, ct, ds))

    def _resolve_ds(s_idx: int, name: str, pos: tuple[float, float]) -> str | None:
        if name and (s_idx, name) in comp_ds_by_name:
            return comp_ds_by_name[(s_idx, name)]
        for cl, ct, ds in comp_ds_by_pos.get(s_idx, []):
            if abs(cl - pos[0]) <= 0.05 and abs(ct - pos[1]) <= 0.05:
                return ds
        return None

    pres = Presentation(str(refreshed_path))
    shifted_correct = []
    shifted_wrong = []
    welded_skip = []  # values may be source-preserved (welded) — out of scope here

    for s_idx, slide in enumerate(pres.slides):
        for shape in slide.shapes:
            if not shape.has_chart:
                continue
            shape_name = shape.name or ""
            sl = round((shape.left or 0) / 914400, 2)
            st = round((shape.top or 0) / 914400, 2)
            ds_key = _resolve_ds(s_idx, shape_name, (sl, st))
            if ds_key is None:
                continue
            expected = expected_names.get(ds_key, [])
            if not expected:
                continue
            cats = [str(c) for c in shape.chart.plots[0].categories]
            series_names = [s.name or "" for s in shape.chart.plots[0].series]
            haystack = " | ".join(cats) + " || " + " | ".join(series_names)
            # Check whether ANY expected wave name appears in cats OR series labels
            hits = [w for w in expected if w in haystack]
            if hits:
                shifted_correct.append({
                    "slide": s_idx, "name": shape_name,
                    "ds": ds_key, "expected": expected, "hits": hits,
                })
            else:
                shifted_wrong.append({
                    "slide": s_idx, "name": shape_name,
                    "ds": ds_key, "expected": expected,
                    "cats": cats, "series_names": series_names,
                })

    return {
        "deck": deck_key,
        "shifted_ds": len(expected_names),
        "charts_on_shifted_ds": len(shifted_correct) + len(shifted_wrong),
        "wave_label_hit": len(shifted_correct),
        "wave_label_miss": len(shifted_wrong),
        "miss_detail": shifted_wrong[:15],
    }


if __name__ == "__main__":
    for k, pid in [("atu_q1_26", 981), ("creon_pet_w33", 523)]:
        print(f"\n=== {k} ===")
        rpt = verify(k, pid)
        print(f"shifted ds:           {rpt.get('shifted_ds')}")
        print(f"charts on shifted ds: {rpt.get('charts_on_shifted_ds')}")
        print(f"  wave label hit (= cat/series shows an expected new wave): {rpt.get('wave_label_hit')}")
        print(f"  wave label miss (= shows source waves instead):          {rpt.get('wave_label_miss')}")
        if rpt.get("miss_detail"):
            print("  first 5 misses:")
            for m in rpt["miss_detail"][:5]:
                print(f"    slide {m['slide']} {m['name']!r} ({m['ds']})")
                print(f"      expected any of: {m['expected']}")
                print(f"      cats: {m['cats'][:6]}")
                print(f"      series: {m['series_names'][:4]}")
