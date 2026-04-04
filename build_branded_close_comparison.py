"""
Build Slide 4 (Branded Close) two ways for comparison:
  A) Pipeline-driven: config.yaml → renderer → PPTX
  B) Hand-crafted: direct python-pptx with SlideGen utils → PPTX

Output:
  Projects/Project - Magnum Opus/output/Q1 2026/slide4_approach_a.pptx
  Projects/Project - Magnum Opus/output/Q1 2026/slide4_approach_b.pptx
"""
import os, sys, json, copy

sys.path.insert(0, os.path.dirname(__file__))

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# ── SlideGen imports ─────────────────────────────────────────────────────────
from slidegen.pipeline.project_config import load_project_config
from slidegen.pipeline.data_loaders import load_all_data
from slidegen.pipeline.slide_renderers import RENDERERS
from slidegen.pptx_utils.deck import load_template
from slidegen.pptx_utils.brand import (
    C_RED, C_GREEN, C_GREY, C_WHITE, C_FTGREY, C_LBGREY, C_HDRGREY,
    FONT_DISPLAY, FONT_TEXT,
)
from slidegen.pptx_utils.shapes import textbox, solidrect, callout_box
from slidegen.pptx_utils.charts import add_clustered_bar_chart, enable_data_labels
from slidegen.pptx_utils.layout import slide_header, slide_footer, section_header_bar, manual_legend
from slidegen.pptx_utils.tables import add_delta_table
from slidegen.pptx_utils.text import delta_format, add_run

# ── Paths ────────────────────────────────────────────────────────────────────
PROJECT = "Projects/Project - Magnum Opus"
CONFIG  = f"{PROJECT}/config.yaml"
OUT_DIR = f"{PROJECT}/output/Q1 2026"
os.makedirs(OUT_DIR, exist_ok=True)

# ── New insight-first headline (from narrative_threads_v3.md) ────────────────
HEADLINE = (
    "Q4\u2019s coaching emphasis on closing technique is paying off \u2014 "
    "RYB overtook TAG on branded close for the first time, with the "
    "competitive swing validating that reps execute the ask when "
    "interactions reach sufficient quality"
)

# ── Colors ───────────────────────────────────────────────────────────────────
C_RYB      = RGBColor(0xF7, 0x58, 0x24)
C_RYB_PRIOR= RGBColor(0xFF, 0xC1, 0x99)
C_TAG      = RGBColor(0x70, 0x30, 0xA0)
C_TAG_PRIOR= RGBColor(0xAD, 0x88, 0xC8)
C_HII_GREEN= RGBColor(0x2E, 0x7D, 0x32)
C_HII_GREY = RGBColor(0xBD, 0xBD, 0xBD)

# ════════════════════════════════════════════════════════════════════════════
#  APPROACH A — Pipeline-driven (config → renderer → PPTX)
# ════════════════════════════════════════════════════════════════════════════

def approach_a():
    print("\n═══ APPROACH A: Pipeline-driven ═══")
    config = load_project_config(CONFIG)
    data = load_all_data(config)

    # Find branded_closing ask
    ask = None
    for a in config.asks:
        if a.id == "branded_closing":
            ask = copy.deepcopy(a)
            break
    if ask is None:
        print("  ERROR: branded_closing ask not found in config.yaml")
        return

    # Override headline with v3 insight-first version
    ask.headline = HEADLINE

    # Create fresh presentation from template
    prs, blank_layout = load_template(config)
    slide = prs.slides.add_slide(blank_layout)

    # Call the pipeline renderer
    RENDERERS["clustered_compare"](slide, config, ask, data)

    out = os.path.join(OUT_DIR, "slide4_approach_a.pptx")
    prs.save(out)
    print(f"  Saved: {out}")


# ════════════════════════════════════════════════════════════════════════════
#  APPROACH B — Hand-crafted (direct python-pptx + SlideGen utils)
# ════════════════════════════════════════════════════════════════════════════

def approach_b():
    print("\n═══ APPROACH B: Hand-crafted ═══")
    config = load_project_config(CONFIG)

    # ── Data ─────────────────────────────────────────────────────────────────
    ryb = {"label": "Asked to prescribe\n(branded close)", "current": 53.0, "prior": 46.0}
    tag = {"label": "Asked to prescribe\n(branded close)", "current": 44.0, "prior": 55.0}
    compelling_ryb = {"label": "Compelling reason\nto prescribe", "current": 53.0, "prior": 49.0}
    compelling_tag = {"label": "Compelling reason\nto prescribe", "current": 51.0, "prior": 53.0}
    opinion_ryb = {"label": "Changed HCP\nopinion", "current": 56.0, "prior": 57.0}
    opinion_tag = {"label": "Changed HCP\nopinion", "current": 58.0, "prior": 52.0}

    metrics = [
        ("Asked to prescribe (branded close)", 53.0, 46.0, 44.0, 55.0),
        ("Compelling reason to prescribe",     53.0, 49.0, 51.0, 53.0),
        ("Changed HCP opinion",                56.0, 57.0, 58.0, 52.0),
    ]
    labels    = [m[0] for m in metrics]
    ryb_cur   = [m[1] for m in metrics]
    ryb_pri   = [m[2] for m in metrics]
    tag_cur   = [m[3] for m in metrics]
    tag_pri   = [m[4] for m in metrics]
    ryb_delta = [round(c - p, 1) for c, p in zip(ryb_cur, ryb_pri)]
    tag_delta = [round(c - p, 1) for c, p in zip(tag_cur, tag_pri)]

    # ── Create presentation ─────────────────────────────────────────────────
    prs, blank_layout = load_template(config)
    slide = prs.slides.add_slide(blank_layout)

    # ── 1. HEADLINE ─────────────────────────────────────────────────────────
    # Red headline bar at top
    hl_top = 0.16
    hl_left = 0.20
    hl_w = 10.50
    hl_h = 1.15
    tb = textbox(slide, HEADLINE, hl_left, hl_top, hl_w, hl_h,
                 fsize=14, bold=True, color=C_RED, font=FONT_DISPLAY, wrap=True)

    # ── 2. MODULE BADGE ─────────────────────────────────────────────────────
    badge_w, badge_h = 2.20, 0.32
    badge_l = 13.333 - badge_w - 0.15
    badge_t = 0.16
    solidrect(slide, badge_l, badge_t, badge_w, badge_h, fill=C_RED)
    textbox(slide, "Personal Promotion", badge_l, badge_t, badge_w, badge_h,
            fsize=9, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER,
            font=FONT_TEXT)

    # ── 3. SECTION HEADER BAR ───────────────────────────────────────────────
    sec_top = 1.40
    sec_h = 0.32
    solidrect(slide, 0.20, sec_top, 12.93, sec_h, fill=RGBColor(0xE8, 0xE8, 0xE8))
    textbox(slide, "PRESCRIBING CONVERSION — CALL-TO-ACTION", 0.30, sec_top, 8.0, sec_h,
            fsize=9, bold=True, color=C_GREY, font=FONT_TEXT)

    # ── 4. LABEL TABLE (left) ───────────────────────────────────────────────
    chart_top = 1.85
    hdr_h = 0.36
    n = len(labels)
    # Fill ~65% of available vertical space (6.78 - 1.85 - 0.55 = 4.38")
    avail_h = 6.78 - chart_top - 0.55
    row_h = min(1.0, max(0.55, (avail_h * 0.65) / max(n, 1)))
    body_h = n * row_h

    label_w = 3.10
    label_l = 0.25

    # Header
    solidrect(slide, label_l, chart_top, label_w, hdr_h, fill=C_HDRGREY)
    textbox(slide, "Call-to-Action Metric", label_l + 0.06, chart_top, label_w - 0.12, hdr_h,
            fsize=9, bold=True, color=C_WHITE, font=FONT_TEXT)

    # Rows
    for i, lbl in enumerate(labels):
        row_top = chart_top + hdr_h + i * row_h
        bg = C_LBGREY if i % 2 == 0 else C_WHITE
        solidrect(slide, label_l, row_top, label_w, row_h, fill=bg)
        tb = textbox(slide, lbl, label_l + 0.08, row_top, label_w - 0.16, row_h,
                     fsize=9, color=C_GREY, font=FONT_TEXT, wrap=True)

    # ── 5. CLUSTERED BAR CHART ──────────────────────────────────────────────
    chart_l = label_l + label_w + 0.08
    chart_w = 6.80

    cf, ch = add_clustered_bar_chart(
        slide,
        categories=[f"R{i}" for i in range(n)],
        series_list=[
            ("RYB+LAZ Q1'26", ryb_cur),
            ("TAG Q1'26", tag_cur),
        ],
        left=chart_l,
        top=chart_top + hdr_h,
        width=chart_w,
        height=body_h,
        colors=[C_RYB, C_TAG],
        legend=False,
        gap=100,
        overlap=0,
        cat_font_size=1,
        label_fsize=9,
        font_name=FONT_TEXT,
    )

    # Hide category axis labels (label table provides them)
    from slidegen.pptx_utils.lxml_helpers import hide_cat_labels, invert_cat_axis, hide_axis, set_chart_plot_area, set_overlap
    hide_axis(ch, "val")
    hide_cat_labels(ch)
    invert_cat_axis(ch)
    set_chart_plot_area(ch, x=0.0, y=0.0, w=1.0, h=1.0)
    set_overlap(ch, -15)

    # ── 6. DELTA COLUMNS ────────────────────────────────────────────────────
    delta_w = 0.65
    gap = 0.06
    d1_l = chart_l + chart_w + gap
    d2_l = d1_l + delta_w + gap

    add_delta_table(slide, ryb_delta, left=d1_l, top=chart_top, width=delta_w,
                    row_height=row_h, header_text="RYB Δ", font_name=FONT_TEXT)
    add_delta_table(slide, tag_delta, left=d2_l, top=chart_top, width=delta_w,
                    row_height=row_h, header_text="TAG Δ", font_name=FONT_TEXT)

    # ── 7. LEGEND ───────────────────────────────────────────────────────────
    legend_top = chart_top + hdr_h + body_h + 0.15
    legend_items = [
        (C_RYB, "RYB+LAZ Q1'26"),
        (C_TAG, "TAG Q1'26"),
    ]
    ly = legend_top
    lx = chart_l
    for color, label in legend_items:
        solidrect(slide, lx, ly, 0.20, 0.16, fill=color)
        textbox(slide, label, lx + 0.24, ly - 0.02, 1.50, 0.20,
                fsize=8.5, color=C_GREY, font=FONT_TEXT)
        lx += 2.0

    # ── 8. HII CALLOUT BOX ─────────────────────────────────────────────────
    callout_top = legend_top + 0.35
    callout_l = 0.25
    callout_w = 6.50
    callout_h = 0.75

    callout_box(slide, callout_l, callout_top, callout_w, callout_h,
                border_color=C_HII_GREEN, dashed=True, fsize=9)

    # Add text inside callout
    hii_text = (
        "HII Insight: Branded close in HII interactions 65% vs Others 35% "
        "(+30pp gap) — confirming the conversion mechanism activates at "
        "high-impact quality"
    )
    textbox(slide, hii_text, callout_l + 0.10, callout_top + 0.05,
            callout_w - 0.20, callout_h - 0.10,
            fsize=8.5, color=RGBColor(0x1B, 0x5E, 0x20), font=FONT_TEXT, wrap=True)

    # ── 9. ACTION ITEM ──────────────────────────────────────────────────────
    action_top = callout_top + callout_h + 0.10
    action_text = (
        "[ACTION ITEM] Tests Rec 5 from Q4'25: "
        "\"STRENGTHEN closing to further reinforce intent to increase "
        "RYB+LAZ prescriptions\" — CONFIRMED: branded close surged +7pp"
    )
    textbox(slide, action_text, 0.25, action_top, 10.0, 0.35,
            fsize=8, italic=True, color=C_FTGREY, font=FONT_TEXT, wrap=True)

    # ── 10. SOURCE FOOTER ───────────────────────────────────────────────────
    footer_text = (
        "Source: ZoomRx PET Q1.83 / C1_81 / C1_82 — "
        "RYB+LAZ (n=102) vs TAG (n=70) | Q4'25 vs Q1'26"
    )
    slide_footer(slide, footer_text, font=FONT_TEXT)

    # ── 11. SPEAKER NOTES ───────────────────────────────────────────────────
    notes_slide = slide.notes_slide
    notes_tf = notes_slide.notes_text_frame
    notes_tf.text = (
        "Slide 4 — Branded Close Competitive Reversal\n"
        "Arc: Thread 1 — Conversion Engine (TENSION — ACT NOW)\n"
        "Hypotheses: H2 (CONFIRMED), H4 (PARTIALLY CONFIRMED)\n"
        "Question: Q1.83 — Did the rep specifically ask you to prescribe?\n"
        "Key data: RYB 53% (+7pp) vs TAG 44% (-11pp) — 18pp swing\n"
        "HII: 65% vs Others 35% (+30pp gap)\n"
        "Action Item: Tests Rec 5 — STRENGTHEN closing"
    )

    out = os.path.join(OUT_DIR, "slide4_approach_b.pptx")
    prs.save(out)
    print(f"  Saved: {out}")


# ════════════════════════════════════════════════════════════════════════════
#  RUN BOTH
# ════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    approach_a()
    approach_b()
    print(f"\n✓ Both files saved to {OUT_DIR}/")
    print("  Open both in PowerPoint and compare side-by-side.")
