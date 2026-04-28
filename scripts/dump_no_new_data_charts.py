"""Dump every WELDED · NO NEW DATA chart to a CSV with:
  slide_idx, slide_idx_1based, chart_name, ds_key, analysis_id, segment_ids,
  source_waves (from chart cats/series labels), api_waves (what Synapse returned)

Lets Bharadvaj spot-check whether each chart is genuinely "no new data" or
whether the API is returning waves the chart should be showing but isn't.
"""
from __future__ import annotations

import csv
import json
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO))

from pptx import Presentation
from slidegen.intelligent_refresh import fetch_synapse_data
from tests.evals.end_to_end.compare_decks import (
    _extract_chart_data, _find_match,
)

DECKS = {
    "atu_q1_26": REPO / "projects" / "J&J Rybrevant PET" / "Template" / "ZoomRx_UC_ATU_Report_Q1_'26.pptx",
    "creon_pet_w33": REPO / "projects" / "J&J Rybrevant PET" / "Template" / "CREON Share of Voice Study - W33.pptx",
}


def _values_eq(a, b, tol=1e-3):
    if len(a) != len(b):
        return False
    for x, y in zip(a, b):
        if x is None and y is None:
            continue
        if x is None or y is None:
            return False
        if abs(x - y) > tol:
            return False
    return True


def _series_eq(a_series, b_series):
    if len(a_series) != len(b_series):
        return False
    for (an, av), (bn, bv) in zip(a_series, b_series):
        if an != bn:
            return False
        if not _values_eq(av, bv):
            return False
    return True


def dump_deck(deck_key: str) -> None:
    src_path = DECKS[deck_key]
    ref_path = REPO / "output" / "step2_test_connected" / f"{deck_key}.pptx"
    spec_path = REPO / "output" / "step2_test_connected" / f"{deck_key}_full_spec.json"
    refresh_status_path = REPO / "output" / "step2_test_connected" / f"{deck_key}_refresh_status.json"

    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    chart_status_by_name: dict[tuple, str] = {}
    if refresh_status_path.exists():
        rs = json.loads(refresh_status_path.read_text(encoding="utf-8"))
        for slide_res in rs.get("slides", []):
            s_idx = slide_res.get("slide_index")
            for c in slide_res.get("charts", []):
                chart_status_by_name[(s_idx, c.get("name", ""))] = c.get("status", "")

    comp_ds_by_name: dict[tuple, str] = {}
    comp_ds_by_pos: dict[int, list] = {}
    for slide in spec.get("slides", []):
        s_idx = slide["slide_index"]
        slide_default = slide.get("data_source")
        for comp in slide.get("components", []):
            if comp.get("type") != "chart":
                continue
            ds = comp.get("data_source") or slide_default
            name = comp.get("name") or ""
            if name:
                comp_ds_by_name[(s_idx, name)] = ds
            cpos = comp.get("position", {}) or {}
            cl = float(cpos.get("left", 0) or 0)
            ct = float(cpos.get("top", 0) or 0)
            comp_ds_by_pos.setdefault(s_idx, []).append((cl, ct, ds))

    static_ds = {
        ds_key for ds_key, d in spec.get("data_sources", {}).items()
        if d.get("static_time_period_ids") and not d.get("include_live_wave")
    }

    api_wave_cache: dict[str, list[str]] = {}

    def _api_waves_list(ds_key: str) -> list[str]:
        if ds_key in api_wave_cache:
            return api_wave_cache[ds_key]
        d = spec["data_sources"].get(ds_key, {})
        try:
            records, _ = fetch_synapse_data({
                "project_id": d.get("project_id"),
                "reporting_plan_id": d.get("reporting_plan_id"),
                "analysis_ids": list(d.get("analysis_ids") or []),
                "segment_ids": list(d.get("segment_ids") or []),
                "dynamic_latest_n": d.get("dynamic_latest_n") or 0,
                "static_time_period_ids": d.get("static_time_period_ids") or [],
                "static_time_period_names": [],
                "include_live_wave": bool(d.get("include_live_wave", False)),
            })
        except Exception:
            records = []
        names = sorted({str(r.get("time_period_name")) for r in records
                       if r.get("time_period_name") is not None})
        api_wave_cache[ds_key] = names
        return names

    def _resolve_ds(s_idx, name, pos):
        if name and (s_idx, name) in comp_ds_by_name:
            return comp_ds_by_name[(s_idx, name)]
        for cl, ct, ds in comp_ds_by_pos.get(s_idx, []):
            if abs(cl - pos[0]) <= 0.05 and abs(ct - pos[1]) <= 0.05:
                return ds
        return None

    src = Presentation(str(src_path))
    ref = Presentation(str(ref_path))

    rows = []
    for s_idx, (s_slide, r_slide) in enumerate(zip(src.slides, ref.slides)):
        src_charts = [s for s in s_slide.shapes if s.has_chart]
        ref_charts = [s for s in r_slide.shapes if s.has_chart]
        for src_shape in src_charts:
            shape_name = src_shape.name or ""
            sl = round((src_shape.left or 0) / 914400, 2)
            st = round((src_shape.top or 0) / 914400, 2)
            ds_key = _resolve_ds(s_idx, shape_name, (sl, st))
            if ds_key is None:
                continue
            if ds_key in static_ds:
                continue  # those are WELDED · STATIC, not WELDED · NO NEW DATA
            ref_shape = _find_match(src_shape, ref_charts)
            if ref_shape is None:
                continue
            try:
                src_cats, src_series = _extract_chart_data(src_shape)
                _, ref_series = _extract_chart_data(ref_shape)
            except Exception:
                continue
            if not _series_eq(src_series, ref_series):
                continue   # not welded
            # Welded — case 1 (status=='ok') or case 2 (status='error' or other)
            mapper_status = chart_status_by_name.get((s_idx, shape_name), "")
            if mapper_status != "ok":
                continue   # case 2 — mapper failed; not what we're listing here

            ds = spec["data_sources"][ds_key]
            api_waves = _api_waves_list(ds_key)
            src_labels = list(src_cats) + [n for n, _ in src_series]
            rows.append({
                "deck": deck_key,
                "slide_idx_0based": s_idx,
                "slide_idx_1based": s_idx + 1,
                "chart_name": shape_name,
                "ds_key": ds_key,
                "analysis_id": (ds.get("analysis_ids") or [None])[0],
                "segment_ids": ",".join(str(x) for x in (ds.get("segment_ids") or [])),
                "dynamic_latest_n": ds.get("dynamic_latest_n") or 0,
                "include_live_wave": bool(ds.get("include_live_wave", False)),
                "api_waves": ", ".join(api_waves[:8]),
                "n_api_waves": len(api_waves),
                "src_labels_sample": " | ".join(str(s) for s in src_labels[:6]),
            })

    out_path = REPO / "output" / "_forward" / deck_key / "no_new_data_charts.csv"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8", newline="") as f:
        if rows:
            w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)
    print(f"{deck_key}: {len(rows)} WELDED · NO NEW DATA charts -> {out_path.relative_to(REPO)}")


if __name__ == "__main__":
    for k in DECKS:
        dump_deck(k)
