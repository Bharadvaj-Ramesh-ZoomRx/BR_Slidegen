"""Annotate every slide in the round-trip reconstructed deck with a
text-box badge stating the refresh outcome.

Labels (one per slide, aggregated from per-component statuses):
  RECONSTRUCTION DONE      — all dynamic components reconstructed back to
                             source values via back→forward round-trip
  REFRESH DRIFT            — some components moved during back leg but
                             didn't restore to source on forward leg
  STRUCTURAL DRIFT         — cat or series count differs from source
  WELDED · STATIC CONFIG   — slide's components are on static data sources;
                             welding is correct (the ds is configured to
                             show frozen waves)
  WELDED · REFRESH FAILED  — slide's dynamic components welded — back leg
                             didn't change values, so the refresh path
                             couldn't be exercised
  NON-CONNECTED            — slide has no Connector spec coverage

Aggregation rule per slide:
  - All components in one bucket → that bucket
  - Mixed → use worst-case priority:
      DRIFT > STRUCTURAL_DRIFT > REFRESH_FAILED > RECONSTRUCTION > STATIC

Usage:
    python scripts/annotate_round_trip_status.py <deck_key>

Reads:
  output/_roundtrip/<deck_key>/summary.json
  output/_roundtrip/<deck_key>/deck_reconstructed.pptx
  output/step2_test_connected/<deck_key>_full_spec.json

Writes:
  output/_roundtrip/<deck_key>/deck_reconstructed_annotated.pptx
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

REPO_ROOT = Path(__file__).resolve().parents[1]

# Status palette
GREEN = RGBColor(0x2E, 0x7D, 0x32)   # reconstruction done
TEAL = RGBColor(0x00, 0x69, 0x6B)    # static (welded by design)
AMBER = RGBColor(0xF9, 0xA8, 0x25)   # refresh drift
RED = RGBColor(0xC6, 0x28, 0x28)     # refresh failed / structural drift
GREY = RGBColor(0x9E, 0x9E, 0x9E)    # non-connected
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

# Priority for mixed-state slides (lower = higher priority = wins)
PRIORITY = {
    "STRUCTURAL_DRIFT": 0,
    "DRIFT": 1,
    "REFRESH_FAILED": 2,
    "RECONSTRUCTION": 3,
    "STATIC": 4,
    "NONE": 5,
}


def _build_component_status(deck_key: str) -> dict[tuple, str]:
    """Build (slide_idx, name) -> status from the round-trip summary."""
    summary_path = REPO_ROOT / "output" / "_roundtrip" / deck_key / "summary.json"
    if not summary_path.exists():
        return {}
    summary = json.loads(summary_path.read_text(encoding="utf-8"))
    # The summary doesn't store per-component results in detail beyond
    # drift_detail. We need to read the spec + reconstructed deck and
    # cross-reference. For now, rely on drift_detail and welded_back is
    # implicit (any chart not in drift_detail and not in structural_drift_detail).
    # Better: re-run the comparison here using compare_decks helpers.
    return summary  # placeholder — real classification happens below


def _classify_slide(deck_key: str) -> dict[int, str]:
    """Walk source vs reconstructed deck and classify each slide.

    Returns: {slide_idx: status_label}
    Status values: STRUCTURAL_DRIFT | DRIFT | REFRESH_FAILED | RECONSTRUCTION
                   | STATIC | NONE (non-connected)
    """
    from tests.evals.end_to_end.compare_decks import _extract_chart_data, _find_match
    source_path_map = {
        "atu_q1_26": REPO_ROOT / "projects" / "J&J Rybrevant PET" / "Template" / "ZoomRx_UC_ATU_Report_Q1_'26.pptx",
        "creon_pet_w33": REPO_ROOT / "projects" / "J&J Rybrevant PET" / "Template" / "CREON Share of Voice Study - W33.pptx",
    }
    src_path = source_path_map[deck_key]
    back_path = REPO_ROOT / "output" / "_roundtrip" / deck_key / "deck_back.pptx"
    fwd_path = REPO_ROOT / "output" / "_roundtrip" / deck_key / "deck_reconstructed.pptx"
    spec_path = REPO_ROOT / "output" / "step2_test_connected" / f"{deck_key}_full_spec.json"

    spec = json.loads(spec_path.read_text(encoding="utf-8"))

    # Build (slide_idx, name) -> ds_key for components in spec
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

    # Static ds set: those with static_time_period_ids configured
    static_ds = {
        ds_key for ds_key, ds in spec.get("data_sources", {}).items()
        if ds.get("static_time_period_ids")
    }

    def _resolve_ds(s_idx, name, pos):
        if name and (s_idx, name) in comp_ds_by_name:
            return comp_ds_by_name[(s_idx, name)]
        for cl, ct, ds in comp_ds_by_pos.get(s_idx, []):
            if abs(cl - pos[0]) <= 0.05 and abs(ct - pos[1]) <= 0.05:
                return ds
        return None

    src = Presentation(str(src_path))
    back = Presentation(str(back_path)) if back_path.exists() else None
    fwd = Presentation(str(fwd_path)) if fwd_path.exists() else None

    VALUE_TOLERANCE = 1e-3

    def _values_eq(a, b):
        if len(a) != len(b):
            return False
        for x, y in zip(a, b):
            if x is None and y is None:
                continue
            if x is None or y is None:
                return False
            if abs(x - y) > VALUE_TOLERANCE:
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

    slide_status: dict[int, str] = {}

    for s_idx in range(len(src.slides)):
        if s_idx not in spec_slide_idx:
            slide_status[s_idx] = "NONE"
            continue

        s_charts = [s for s in src.slides[s_idx].shapes if s.has_chart]
        b_charts = [s for s in back.slides[s_idx].shapes if s.has_chart] if back else []
        f_charts = [s for s in fwd.slides[s_idx].shapes if s.has_chart] if fwd else []

        statuses: list[str] = []
        for src_shape in s_charts:
            shape_name = src_shape.name or ""
            sl = round((src_shape.left or 0) / 914400, 2)
            st = round((src_shape.top or 0) / 914400, 2)
            ds_key = _resolve_ds(s_idx, shape_name, (sl, st))
            if ds_key is None:
                continue  # not in spec coverage — skip this chart
            if ds_key in static_ds:
                statuses.append("STATIC")
                continue
            # Dynamic: classify via round-trip outcome
            b_shape = _find_match(src_shape, b_charts)
            f_shape = _find_match(src_shape, f_charts)
            if b_shape is None or f_shape is None:
                statuses.append("REFRESH_FAILED")
                continue
            try:
                _, src_series = _extract_chart_data(src_shape)
                _, back_series = _extract_chart_data(b_shape)
                _, fwd_series = _extract_chart_data(f_shape)
            except Exception:
                statuses.append("REFRESH_FAILED")
                continue
            # Structural check
            if (src_shape.chart.chart_type != f_shape.chart.chart_type
                    or len(src_series) != len(fwd_series)):
                statuses.append("STRUCTURAL_DRIFT")
                continue
            # Did back leg change values?
            if _series_eq(src_series, back_series):
                statuses.append("REFRESH_FAILED")
                continue
            # Round-trip: forward should match source
            if _series_eq(src_series, fwd_series):
                statuses.append("RECONSTRUCTION")
            else:
                statuses.append("DRIFT")

        if not statuses:
            # Slide is in spec but no chart components matched (could be table-only).
            # Fall through to "NONE" — for now annotate as non-connected-equivalent.
            slide_status[s_idx] = "NONE"
            continue
        # Pick worst-priority (lowest priority number)
        worst = min(statuses, key=lambda x: PRIORITY[x])
        slide_status[s_idx] = worst

    return slide_status


_BADGE = {
    "RECONSTRUCTION":     (GREEN, "RECONSTRUCTION DONE"),
    "STATIC":             (TEAL,  "WELDED · STATIC CONFIG"),
    "DRIFT":              (AMBER, "REFRESH DRIFT"),
    "STRUCTURAL_DRIFT":   (RED,   "STRUCTURAL DRIFT"),
    "REFRESH_FAILED":     (RED,   "WELDED · REFRESH FAILED"),
    "NONE":               (GREY,  "NON-CONNECTED"),
}


def _add_badge(slide, color: RGBColor, label: str, slide_w_emu: int) -> None:
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
    in_pptx = REPO_ROOT / "output" / "_roundtrip" / deck_key / "deck_reconstructed.pptx"
    out_pptx = REPO_ROOT / "output" / "_roundtrip" / deck_key / "deck_reconstructed_annotated.pptx"

    if not in_pptx.exists():
        print(f"ERROR: reconstructed deck not found: {in_pptx}")
        print("Run the round-trip eval first.")
        sys.exit(2)

    slide_status = _classify_slide(deck_key)

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
