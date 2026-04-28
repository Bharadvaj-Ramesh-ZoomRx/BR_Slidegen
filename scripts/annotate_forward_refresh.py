"""Annotate a forward-refresh deck (Step 2 production refresh output)
with per-slide status badges.

Labels (one per slide, aggregated worst-case across components):
  REFRESHED                 dynamic chart values updated vs source
  WELDED · STATIC CONFIG    static ds, source preserved correctly
  WELDED · NO NEW DATA      dynamic ds, but no value change (no new
                            wave delivered, OR mapper safety net fired)
  STRUCTURAL DRIFT          cat / series count differs from source
  NON-CONNECTED             slide outside spec coverage

Usage:
    python scripts/annotate_forward_refresh.py <deck_key>

Reads:
  output/step2_test_connected/<deck_key>.pptx
  output/step2_test_connected/<deck_key>_full_spec.json

Writes:
  output/step2_test_connected/<deck_key>_annotated.pptx
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

GREEN = RGBColor(0x2E, 0x7D, 0x32)
TEAL = RGBColor(0x00, 0x69, 0x6B)
AMBER = RGBColor(0xF9, 0xA8, 0x25)
RED = RGBColor(0xC6, 0x28, 0x28)
GREY = RGBColor(0x9E, 0x9E, 0x9E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

PRIORITY = {
    "STRUCTURAL_DRIFT": 0,
    "MAPPER_FAILED":    1,   # dynamic chart, API has new waves but mapper couldn't apply
    "REFRESHED":        2,
    "NO_NEW_DATA":      3,   # dynamic chart, API returned same waves as source — correct
    "WELDED_STATIC":    4,
    "NONE":             5,
}

_BADGE = {
    "REFRESHED":          (GREEN, "REFRESHED"),
    "WELDED_STATIC":      (TEAL,  "WELDED · STATIC CONFIG"),
    "NO_NEW_DATA":        (TEAL,  "WELDED · NO NEW DATA"),
    "MAPPER_FAILED":      (AMBER, "WELDED · MAPPER FAILED"),
    "STRUCTURAL_DRIFT":   (RED,   "STRUCTURAL DRIFT"),
    "NONE":               (GREY,  "NON-CONNECTED"),
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


def _classify_slides(deck_key: str) -> dict[int, str]:
    from tests.evals.end_to_end.compare_decks import _extract_chart_data, _find_match
    from slidegen.intelligent_refresh import fetch_synapse_data
    src_path_map = {
        "atu_q1_26": REPO_ROOT / "projects" / "J&J Rybrevant PET" / "Template" / "ZoomRx_UC_ATU_Report_Q1_'26.pptx",
        "creon_pet_w33": REPO_ROOT / "projects" / "J&J Rybrevant PET" / "Template" / "CREON Share of Voice Study - W33.pptx",
    }
    src_path = src_path_map[deck_key]
    ref_path = REPO_ROOT / "output" / "step2_test_connected" / f"{deck_key}.pptx"
    spec_path = REPO_ROOT / "output" / "step2_test_connected" / f"{deck_key}_full_spec.json"

    spec = json.loads(spec_path.read_text(encoding="utf-8"))

    comp_ds_by_name: dict[tuple, str] = {}
    comp_ds_by_pos: dict[int, list[tuple[float, float, str]]] = {}
    spec_slide_idx = set()
    for slide in spec.get("slides", []):
        s_idx = slide["slide_index"]
        spec_slide_idx.add(s_idx)
        slide_default = slide.get("data_source")
        for comp in slide.get("components", []):
            ds = comp.get("data_source") or slide_default
            name = comp.get("name") or ""
            if name:
                comp_ds_by_name[(s_idx, name)] = ds
            cpos = comp.get("position", {}) or {}
            cl = float(cpos.get("left", 0) or 0)
            ct = float(cpos.get("top", 0) or 0)
            comp_ds_by_pos.setdefault(s_idx, []).append((cl, ct, ds))

    static_ds = {
        ds_key for ds_key, ds in spec.get("data_sources", {}).items()
        if ds.get("static_time_period_ids") and not ds.get("include_live_wave")
    }

    def _resolve(s_idx, name, pos):
        if name and (s_idx, name) in comp_ds_by_name:
            return comp_ds_by_name[(s_idx, name)]
        for cl, ct, ds in comp_ds_by_pos.get(s_idx, []):
            if abs(cl - pos[0]) <= 0.05 and abs(ct - pos[1]) <= 0.05:
                return ds
        return None

    src = Presentation(str(src_path))
    ref = Presentation(str(ref_path))

    api_wave_cache: dict[str, set] = {}

    def _api_wave_names(ds_key: str) -> set:
        if ds_key in api_wave_cache:
            return api_wave_cache[ds_key]
        ds = spec["data_sources"].get(ds_key, {})
        try:
            records, _ = fetch_synapse_data({
                "project_id": ds.get("project_id"),
                "reporting_plan_id": ds.get("reporting_plan_id"),
                "analysis_ids": list(ds.get("analysis_ids") or []),
                "segment_ids": list(ds.get("segment_ids") or []),
                "dynamic_latest_n": ds.get("dynamic_latest_n") or 0,
                "static_time_period_ids": ds.get("static_time_period_ids") or [],
                "static_time_period_names": [],
                "include_live_wave": bool(ds.get("include_live_wave", False)),
            })
        except Exception:
            records = []
        names: set = set()
        for r in records:
            n = r.get("time_period_name")
            if n is not None:
                names.add(str(n))
                stripped = str(n).replace("Project Wave ", "Wave ")
                if stripped != str(n):
                    names.add(stripped)
        api_wave_cache[ds_key] = names
        return names

    def _all_api_waves_in_source(api_waves: set, src_labels: list[str]) -> bool:
        if not api_waves:
            return False
        haystack = " | ".join(str(s) for s in src_labels)
        return all(w in haystack for w in api_waves)

    slide_status: dict[int, str] = {}
    for s_idx in range(len(src.slides)):
        if s_idx not in spec_slide_idx:
            slide_status[s_idx] = "NONE"
            continue
        s_charts = [s for s in src.slides[s_idx].shapes if s.has_chart]
        r_charts = [s for s in ref.slides[s_idx].shapes if s.has_chart]
        statuses: list[str] = []
        for src_shape in s_charts:
            shape_name = src_shape.name or ""
            sl = round((src_shape.left or 0) / 914400, 2)
            st = round((src_shape.top or 0) / 914400, 2)
            ds_key = _resolve(s_idx, shape_name, (sl, st))
            if ds_key is None:
                continue
            ref_shape = _find_match(src_shape, r_charts)
            if ref_shape is None:
                statuses.append("MAPPER_FAILED")
                continue
            try:
                src_cats, src_series = _extract_chart_data(src_shape)
                _, ref_series = _extract_chart_data(ref_shape)
            except Exception:
                statuses.append("MAPPER_FAILED")
                continue
            if ds_key in static_ds:
                if _series_eq(src_series, ref_series):
                    statuses.append("WELDED_STATIC")
                else:
                    statuses.append("STRUCTURAL_DRIFT")
                continue
            if (src_shape.chart.chart_type != ref_shape.chart.chart_type
                    or len(src_series) != len(ref_series)):
                statuses.append("STRUCTURAL_DRIFT")
                continue
            if _series_eq(src_series, ref_series):
                # Welded dynamic — split case 1 vs case 2 using API wave check
                src_labels = list(src_cats) + [n for n, _ in src_series]
                api_waves = _api_wave_names(ds_key)
                if _all_api_waves_in_source(api_waves, src_labels):
                    statuses.append("NO_NEW_DATA")
                else:
                    statuses.append("MAPPER_FAILED")
            else:
                statuses.append("REFRESHED")

        if not statuses:
            slide_status[s_idx] = "NONE"
        else:
            slide_status[s_idx] = min(statuses, key=lambda x: PRIORITY[x])
    return slide_status


def _add_badge(slide, color, label, slide_w_emu):
    width = Inches(2.6)
    height = Inches(0.28)
    left = slide_w_emu - width - Inches(0.15).emu
    top = Inches(0.1)
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    box.fill.solid()
    box.fill.fore_color.rgb = color
    box.line.fill.background()
    box.shadow.inherit = False
    tf = box.text_frame
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = label
    run.font.size = Pt(9)
    run.font.bold = True
    run.font.color.rgb = WHITE


def main():
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(2)
    deck_key = sys.argv[1]
    in_pptx = REPO_ROOT / "output" / "step2_test_connected" / f"{deck_key}.pptx"
    out_pptx = REPO_ROOT / "output" / "step2_test_connected" / f"{deck_key}_annotated.pptx"
    if not in_pptx.exists():
        print(f"ERROR: refreshed deck not found: {in_pptx}")
        sys.exit(2)

    slide_status = _classify_slides(deck_key)
    pres = Presentation(str(in_pptx))
    slide_w = pres.slide_width
    counts = {k: 0 for k in _BADGE}
    for s_idx, slide in enumerate(pres.slides):
        status = slide_status.get(s_idx, "NONE")
        color, label = _BADGE[status]
        counts[status] += 1
        _add_badge(slide, color, label, slide_w)
    pres.save(str(out_pptx))
    total = sum(counts.values())
    print(f"=== {deck_key} | annotated {total} slides ===")
    for k, n in counts.items():
        if n:
            print(f"  {k:<20s}: {n} slides")
    print(f"\nWrote: {out_pptx}")


if __name__ == "__main__":
    main()
