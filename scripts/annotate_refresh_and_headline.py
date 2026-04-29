"""Stamp refresh + headline status badges on the with_headlines deck.

Reads:
    output/step2_test_connected/<deck>_with_headlines.pptx       (input)
    output/step2_test_connected/<deck>_refresh_status.json       (refresh outcome)
    output/step2_test_connected/<deck>_with_headlines_status.json (headline outcome)
    output/step2_test_connected/<deck>_full_spec.json            (connected slides)

Writes:
    output/step2_test_connected/<deck>_with_headlines_annotated.pptx

Each slide gets two stacked badges in the top-right corner:
    [REFRESH · ...]    -- refresh-execution outcome from refresh_status.json
    [HEADLINE · ...]   -- whether the headline was rewritten / preserved / skipped

The two badges together let you inspect the contract:
    REFRESH ok        + HEADLINE rewritten  =  values changed + narrative existed
    REFRESH ok        + HEADLINE preserved  =  values unchanged (within threshold)
    REFRESH ok        + HEADLINE no_narr    =  values changed but slide had no narrative
    REFRESH no_data   + HEADLINE preserved  =  no new wave -> headline must not change
    REFRESH static    + HEADLINE preserved  =  static-pinned, headline rightly untouched
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

REPO = Path(__file__).resolve().parents[1]

# Palette
GREEN = RGBColor(0x2E, 0x7D, 0x32)
BLUE = RGBColor(0x15, 0x65, 0xC0)
AMBER = RGBColor(0xF9, 0xA8, 0x25)
RED = RGBColor(0xC6, 0x28, 0x28)
GREY = RGBColor(0x9E, 0x9E, 0x9E)
DARK = RGBColor(0x42, 0x42, 0x42)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def _refresh_badge(slide_status: dict | None) -> tuple[RGBColor, str]:
    """Classify a slide's refresh execution outcome from refresh_status.json."""
    if slide_status is None:
        return GREY, "REFRESH - not connected"

    components = slide_status.get("charts", []) + slide_status.get("tables", [])
    if not components:
        return GREY, "REFRESH - no components"

    statuses = [c.get("status", "missing") for c in components]
    n = len(statuses)
    n_ok = sum(1 for s in statuses if s == "ok")
    n_static = sum(1 for s in statuses if s == "static_pinned_skipped")
    n_nodata = sum(1 for s in statuses if s == "no_data_for_shifted_window")
    n_empty = sum(1 for s in statuses if s == "empty")

    if n_ok == n:
        return GREEN, f"REFRESH ok  {n_ok}/{n}"
    if n_static == n:
        return BLUE, f"REFRESH static-pinned  {n_static}/{n}"
    if n_nodata == n:
        return BLUE, f"REFRESH no new data  {n_nodata}/{n}"
    if n_empty == n:
        return RED, f"REFRESH empty  {n_empty}/{n}"

    parts = []
    if n_ok: parts.append(f"{n_ok} ok")
    if n_static: parts.append(f"{n_static} static")
    if n_nodata: parts.append(f"{n_nodata} no-data")
    if n_empty: parts.append(f"{n_empty} empty")
    color = AMBER if (n_ok > 0 or n_static + n_nodata > 0) else RED
    return color, "REFRESH " + " / ".join(parts)


def _headline_badge(headline_status: dict | None) -> tuple[RGBColor, str]:
    """Classify a slide's headline outcome from with_headlines_status.json."""
    if headline_status is None:
        return GREY, "HEADLINE - not in scope"

    s = headline_status.get("status", "")
    if s == "updated":
        return GREEN, "HEADLINE rewritten"
    if s == "unchanged":
        return BLUE, "HEADLINE preserved (data unchanged)"
    if s == "no_headline":
        return AMBER, "HEADLINE skipped (no narrative)"
    if s == "no_chart":
        return DARK, "HEADLINE skipped (no chart)"
    if s == "extract_error":
        return RED, "HEADLINE skipped (extract error)"
    if s == "not_in_spec":
        return GREY, "HEADLINE - not connected"
    return DARK, f"HEADLINE {s}"


def _add_badge(slide, color: RGBColor, label: str, slide_w_emu: int,
               row: int) -> None:
    """Drop a colored badge in the top-right; row=0 is upper, row=1 is below it."""
    width = Inches(2.85)
    height = Inches(0.28)
    left = slide_w_emu - width - Inches(0.15).emu
    top = Inches(0.1) + row * (height + Inches(0.04))

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


def annotate(deck_key: str) -> None:
    base = REPO / "output" / "step2_test_connected"
    in_pptx = base / f"{deck_key}_with_headlines.pptx"
    refresh = json.loads((base / f"{deck_key}_refresh_status.json").read_text(encoding="utf-8"))
    headline = json.loads((base / f"{deck_key}_with_headlines_status.json").read_text(encoding="utf-8"))

    refresh_by_idx = {s["slide_index"]: s for s in refresh.get("slides", [])}
    headline_by_idx = {u["slide_idx"]: u for u in headline.get("updates", [])}

    pres = Presentation(str(in_pptx))
    slide_w = pres.slide_width

    counts = {"both_green": 0, "ok+preserved": 0, "ok+no_narr": 0,
              "no_data+preserved": 0, "static+preserved": 0,
              "mismatch": 0, "untagged": 0}

    for s_idx, slide in enumerate(pres.slides):
        r_color, r_label = _refresh_badge(refresh_by_idx.get(s_idx))
        h_color, h_label = _headline_badge(headline_by_idx.get(s_idx))
        _add_badge(slide, r_color, r_label, slide_w, row=0)
        _add_badge(slide, h_color, h_label, slide_w, row=1)

        # Bucket for summary
        if r_color == GREY and h_color == GREY:
            counts["untagged"] += 1
        elif "ok" in r_label and h_color == GREEN:
            counts["both_green"] += 1
        elif "ok" in r_label and h_color == BLUE:
            counts["ok+preserved"] += 1
        elif "ok" in r_label and h_color == AMBER:
            counts["ok+no_narr"] += 1
        elif "no new data" in r_label and h_color == BLUE:
            counts["no_data+preserved"] += 1
        elif "static" in r_label and h_color == BLUE:
            counts["static+preserved"] += 1
        else:
            counts["mismatch"] += 1

    # If the existing _annotated file is open in PowerPoint, fall back to v2
    out = base / f"{deck_key}_with_headlines_annotated.pptx"
    try:
        pres.save(str(out))
    except PermissionError:
        out = base / f"{deck_key}_with_headlines_annotated_v2.pptx"
        pres.save(str(out))
        print(f"  (existing _annotated.pptx was locked — wrote v2 instead)")

    print(f"=== {deck_key} ===")
    for k, v in counts.items():
        print(f"  {k}: {v}")
    print(f"\nWrote: {out.relative_to(REPO)}")


if __name__ == "__main__":
    if len(sys.argv) != 2 or sys.argv[1] not in {"atu_q1_26", "creon_pet_w33"}:
        print("Usage: python scripts/annotate_refresh_and_headline.py {atu_q1_26|creon_pet_w33}")
        sys.exit(2)
    annotate(sys.argv[1])
