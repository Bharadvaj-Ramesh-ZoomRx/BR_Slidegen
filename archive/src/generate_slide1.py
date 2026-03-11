#!/usr/bin/env python3
"""
generate_slide1.py
───────────────────
Single perfected client-delivery slide  –  Associate Ask #1
RYB+LAZ Message Recall & Effectiveness

Design decisions
  • Template   : Q4'25 JJ Report (theme, fonts, color palette)
  • Charts     : Native PPT charts (editable in PowerPoint – no images)
  • Labels     : Full message text from Q2_10Z banner plan (not coded labels)
  • Delta table: Single delta (Δ) column, no repeated message-label column
  • Order      : Sorted by MR Q4 descending, highest recall at top
  • Alignment  : All objects on a shared grid; delta tables sized/positioned
                 to match chart height so rows align with bars
"""

import os, pickle, textwrap
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import ChartData
from pptx.oxml.ns import qn

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(BASE, "JJ PET RYBREVANT+LAZCLUZE Q4'25 Report.pptx")
PKL      = os.path.join(BASE, "output", "slide_data_v2.pkl")
OUT      = os.path.join(BASE, "output", "Slide1_MR_ME_Final_v5.pptx")

# ── Brand colours ─────────────────────────────────────────────────────────────
C_RYB_Q4  = RGBColor(0xF7, 0x58, 0x24)   # deep orange  – Q4 bars
C_RYB_Q3  = RGBColor(0xFF, 0xC1, 0x99)   # pale orange  – Q3 bars
C_RED     = RGBColor(0xFF, 0x00, 0x00)    # J&J red
C_WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
C_GREY    = RGBColor(0x50, 0x50, 0x50)
C_FTGREY  = RGBColor(0x7F, 0x7F, 0x7F)
C_GREEN   = RGBColor(0x00, 0xB0, 0x50)   # positive delta
C_LBGREY  = RGBColor(0xF4, 0xF4, 0xF4)   # alt table row bg
C_HDRGREY = RGBColor(0x40, 0x40, 0x40)   # delta table header bg

# ── Full message texts from Q2_10Z banner plan (RYB sheet rows 236-245) ───────
FULL_TEXT = {
    "NCCN - NSCLC":
        "RYBREVANT® + LAZCLUZE® is NCCN Category 1 recommended 1L treatment for EGFR+ mNSCLC",
    "Efficacy – mPFS":
        "7.1-month mPFS improvement vs osimertinib (23.7 mo vs 16.6 mo)",
    "Indication":
        "Indicated for 1L treatment of adults with locally advanced or metastatic EGFR+ NSCLC",
    "Safety – Prophylaxis protocol":
        "Proactive therapy management (SKIPPirr & COCOON) reduces IRRs and dermatologic ARs",
    "Median OS Not Reached":
        "Median overall survival not reached; projected to exceed 4 years",
    "OS Headline":
        "Unmatched survival: superior OS vs osimertinib with proven durability",
    "MOA – inhibition":
        "Reduces MET amplification & secondary EGFR alterations (chemo-free, multi-targeted)",
    "Efficacy – CNS":
        "2x Intracranial PFS at 36 months: 36% with RYB+LAZ vs 18% with osimertinib",
    "NCCN - CNS":
        "NCCN preferred: only RYBREVANT®-based regimen for brain mets in EGFR+ mNSCLC",
    "Safety – ARs":
        "ARs peaked in the first 4 months and declined over the next 4 months",
}

# ── Wrap label text for chart axis (forces newlines so PPT wraps cleanly) ─────
def wrap_label(text, width=50):
    return "\n".join(textwrap.wrap(text, width))

# ── XML helpers ───────────────────────────────────────────────────────────────
def _get_or_add(parent, tag):
    el = parent.find(qn(tag))
    if el is None:
        el = etree.SubElement(parent, qn(tag))
    return el

def invert_cat_axis(chart):
    """Show first category at top (maxMin orientation)."""
    catAx = chart.category_axis._element
    scaling = _get_or_add(catAx, "c:scaling")
    orient  = _get_or_add(scaling, "c:orientation")
    orient.set("val", "maxMin")

def hide_cat_labels(chart):
    """Hide category-axis tick labels (Y-axis on horizontal bar)."""
    catAx = chart.category_axis._element
    tlp   = _get_or_add(catAx, "c:tickLblPos")
    tlp.set("val", "none")

def set_datalabel_pos_outside_end(series):
    """Force data labels to appear outside-end (right of bar)."""
    dLbls = series._element.find(qn("c:dLbls"))
    if dLbls is not None:
        pos = _get_or_add(dLbls, "c:dLblPos")
        pos.set("val", "outEnd")

def set_series_no_border(series):
    """Remove visible border on bar series."""
    spPr = series._element.get_or_add_spPr()
    ln   = _get_or_add(spPr, "a:ln")
    _get_or_add(ln, "a:noFill")

def set_val_axis_number_format(axis, fmt="0"):
    """Set number format on the value axis tick labels."""
    axEl    = axis._element
    numFmt  = _get_or_add(axEl, "c:numFmt")
    numFmt.set("formatCode", fmt)
    numFmt.set("sourceLinked", "0")

def set_plot_area_gap(chart, gap_pct=80):
    """Set gap between bar clusters (as % of bar width, lower = fatter bars)."""
    barChart = chart.plots[0]._element
    gapWidth = barChart.find(qn("c:gapWidth"))
    if gapWidth is None:
        gapWidth = etree.SubElement(barChart, qn("c:gapWidth"))
    gapWidth.set("val", str(gap_pct))

def set_overlap(chart, overlap=0):
    """Set bar overlap percentage within a cluster."""
    barChart = chart.plots[0]._element
    ov = barChart.find(qn("c:overlap"))
    if ov is None:
        ov = etree.SubElement(barChart, qn("c:overlap"))
    ov.set("val", str(overlap))

# ── Shape helpers ─────────────────────────────────────────────────────────────
def textbox(slide, text, left, top, width, height,
            fsize=9, bold=False, color=C_GREY,
            align=PP_ALIGN.LEFT, italic=False, wrap=True):
    shape = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height))
    tf = shape.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(fsize)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return shape

def solidrect(slide, left, top, width, height, fill, line=None):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    shape = slide.shapes.add_shape(
        1,  # RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()  # no border
    else:
        shape.line.color.rgb = line
    return shape

def horiz_line(slide, left, top, width, color=C_RED, width_pt=1.5):
    shape = slide.shapes.add_connector(
        1,  # STRAIGHT
        Inches(left), Inches(top), Inches(left + width), Inches(top))
    shape.line.color.rgb = color
    shape.line.width = Pt(width_pt)
    return shape

# ── Delta table (single column: header + N delta values) ─────────────────────
def add_delta_col(slide, deltas, left, top, width, height, header,
                  hdr_h_frac=0.06):
    """
    Creates a 1-column PPT table aligned with the chart.
    deltas: list of float|None  (same order as chart categories, top-to-bottom)
    hdr_h_frac: header row as fraction of total height
    """
    n      = len(deltas)
    hdr_h  = height * hdr_h_frac        # small header
    body_h = height - hdr_h
    row_h  = body_h / n

    tbl = slide.shapes.add_table(
        n + 1, 1,
        Inches(left), Inches(top), Inches(width), Inches(height)
    ).table

    # ── header row ────────────────────────────────────────────────────────────
    tbl.rows[0].height = Emu(int(hdr_h * 914400))
    hc = tbl.cell(0, 0)
    hc.fill.solid()
    hc.fill.fore_color.rgb = C_HDRGREY
    hc.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    run = hc.text_frame.paragraphs[0].add_run()
    run.text = header
    run.font.size = Pt(7)
    run.font.bold = True
    run.font.color.rgb = C_WHITE

    # ── data rows ─────────────────────────────────────────────────────────────
    for i, d in enumerate(deltas):
        tbl.rows[i + 1].height = Emu(int(row_h * 914400))
        cell = tbl.cell(i + 1, 0)

        # Alternating background
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_LBGREY if i % 2 == 0 else C_WHITE

        # Value + colour
        if d is None:
            text, fcolor = "N/A", C_FTGREY
        elif d > 0:
            text, fcolor = f"+{d:.1f}", C_GREEN
        elif d < 0:
            text, fcolor = f"{d:.1f}", C_RED
        else:
            text, fcolor = "0.0", C_GREY

        tf = cell.text_frame
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        run = tf.paragraphs[0].add_run()
        run.text = text
        run.font.size = Pt(8)
        run.font.bold = True
        run.font.color.rgb = fcolor

    # Column width
    tbl.columns[0].width = Inches(width)
    return tbl

# ══════════════════════════════════════════════════════════════════════════════
# LOAD DATA
# ══════════════════════════════════════════════════════════════════════════════
with open(PKL, "rb") as f:
    D = pickle.load(f)

msgs     = D["s3_ryb_messaging"]   # already sorted by mr_q4 desc
meta     = D["meta"]
RYB_N3   = meta["ryb_n_q3"]        # 103
RYB_N4   = meta["ryb_n_q4"]        # 100

# Build ordered lists (descending by MR Q4, highest recall first)
tags      = [m["tag"]      for m in msgs]
full_labs = [wrap_label(FULL_TEXT.get(m["tag"], m["tag"])) for m in msgs]

mr_q4     = [m["mr_q4"]    for m in msgs]   # None kept → chart shows gap
mr_q3     = [m["mr_q3"]    for m in msgs]   # None for NCCN (new msgs)
mr_delta  = [m["mr_delta"] for m in msgs]

me_q4     = [m["me_q4"]    for m in msgs]
me_q3     = [m["me_q3"]    for m in msgs]   # None for NCCN
me_delta  = [m["me_delta"] for m in msgs]

# PPT horizontal bar default: first category → bottom.
# We want first (highest) at TOP, so we reverse everything fed to the chart.
# The delta table is built in visual top-to-bottom order (same reversed list).
rev_full  = list(reversed(full_labs))
rev_mr_q4 = list(reversed(mr_q4));  rev_mr_q3 = list(reversed(mr_q3))
rev_me_q4 = list(reversed(me_q4));  rev_me_q3 = list(reversed(me_q3))
rev_mr_d  = list(reversed(mr_delta))   # visual top→bottom = highest recall first
rev_me_d  = list(reversed(me_delta))

# ══════════════════════════════════════════════════════════════════════════════
# PRESENTATION – open template, add blank slide, strip other slides
# ══════════════════════════════════════════════════════════════════════════════
prs = Presentation(TEMPLATE)

# Find the Blank layout
blank_layout = next(
    (l for l in prs.slide_layouts if "blank" in l.name.lower()),
    prs.slide_layouts[-1]
)

# Add our slide (will be appended at the end)
slide = prs.slides.add_slide(blank_layout)

# Remove all slides EXCEPT the last one (ours)
sldIdLst = prs.slides._sldIdLst
all_ids  = list(sldIdLst)
for sldId in all_ids[:-1]:
    sldIdLst.remove(sldId)

# ══════════════════════════════════════════════════════════════════════════════
# LAYOUT GRID  (all inches)
# ══════════════════════════════════════════════════════════════════════════════
W, H = 13.333, 7.500   # slide dimensions

CHART_TOP   = 1.47
CHART_H     = 5.05

MR_LEFT     = 0.20
MR_W        = 7.30      # wider frame: labels (~40%) + bars (~60%) = legible and spacious
MR_DELTA_L  = MR_LEFT + MR_W + 0.04
MR_DELTA_W  = 0.62

# ME has no Y-axis labels, so its full frame width = bar width.
# MR bar area ≈ MR_W × 0.60 ≈ 4.38" → set ME_W = 4.40" to match physically.
ME_LEFT     = MR_DELTA_L + MR_DELTA_W + 0.10   # centre gap
ME_W        = 4.40      # matches MR's estimated bar-only area for visual parity
ME_DELTA_L  = ME_LEFT + ME_W + 0.04
ME_DELTA_W  = 0.62

DELTA_H     = CHART_H - 0.12   # marginal bottom trim on delta tables

# ══════════════════════════════════════════════════════════════════════════════
# HEADER SECTION
# ══════════════════════════════════════════════════════════════════════════════

# Thin coloured band across full top of slide (acts as title bar)
solidrect(slide, 0, 0.15, W, 0.02, C_RED)   # red top accent line

# Section / module label (small, grey, top-right)
textbox(slide, "Personal Promotion Module  |  Message Strategy",
        5.5, 0.01, 7.70, 0.22,
        fsize=7.5, color=C_FTGREY, align=PP_ALIGN.RIGHT)

# Headline text
HEADLINE = (
    "Across all RYB+LAZ messages, recall declined Q3→Q4; "
    "NCCN Category 1 leads Q4 recall while OS and CNS efficacy messages "
    "sustain the highest effectiveness ratings"
)
textbox(slide, HEADLINE,
        0.20, 0.20, 10.55, 1.05,
        fsize=12, bold=True, color=C_RED, align=PP_ALIGN.LEFT)

# Module badge (red rectangle + white text, top-right)
solidrect(slide, 10.90, 0.20, 2.25, 0.95, C_RED)
textbox(slide, "PERSONAL\nPROMOTION\nMODULE",
        10.90, 0.20, 2.25, 0.95,
        fsize=8, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

# Separator line below header
horiz_line(slide, 0.0, 1.32, W, color=C_RED, width_pt=1.0)

# ══════════════════════════════════════════════════════════════════════════════
# CHART COLUMN TITLES  (textboxes above each chart)
# ══════════════════════════════════════════════════════════════════════════════
COL_LBL_TOP = CHART_TOP - 0.20
COL_LBL_H   = 0.22

textbox(slide, "Message Recall (% Recalled)  |  sorted by Q4'25 descending",
        MR_LEFT, COL_LBL_TOP, MR_W + MR_DELTA_W + 0.04, COL_LBL_H,
        fsize=8.5, bold=True, color=C_GREY, align=PP_ALIGN.CENTER)

textbox(slide, "Message Effectiveness (% Effective, Top-2 Box†)",
        ME_LEFT, COL_LBL_TOP, ME_W + ME_DELTA_W + 0.04, COL_LBL_H,
        fsize=8.5, bold=True, color=C_GREY, align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════════════════════
# MR CHART
# ══════════════════════════════════════════════════════════════════════════════
mr_data = ChartData()
mr_data.categories = rev_full
mr_data.add_series("Q3'25", rev_mr_q3)
mr_data.add_series("Q4'25", rev_mr_q4)

mr_chart_shape = slide.shapes.add_chart(
    XL_CHART_TYPE.BAR_CLUSTERED,
    Inches(MR_LEFT), Inches(CHART_TOP), Inches(MR_W), Inches(CHART_H),
    mr_data
)
mr_chart = mr_chart_shape.chart
mr_chart.has_title  = False
mr_chart.has_legend = False

# Series colours
for i, s in enumerate(mr_chart.series):
    s.format.fill.solid()
    s.format.fill.fore_color.rgb = [C_RYB_Q3, C_RYB_Q4][i]
    set_series_no_border(s)

# Data labels on both Q3 and Q4 series
q3_mr = mr_chart.series[0]
q3_mr.data_labels.show_value        = True
q3_mr.data_labels.number_format     = "0"
q3_mr.data_labels.font.size         = Pt(7)
q3_mr.data_labels.font.color.rgb    = C_FTGREY
set_datalabel_pos_outside_end(q3_mr)

q4_mr = mr_chart.series[1]
q4_mr.data_labels.show_value        = True
q4_mr.data_labels.number_format     = "0"
q4_mr.data_labels.font.size         = Pt(7)
q4_mr.data_labels.font.color.rgb    = C_GREY
set_datalabel_pos_outside_end(q4_mr)

# Category axis (Y) – full text labels, small font
mr_cat = mr_chart.category_axis
mr_cat.tick_labels.font.size      = Pt(6.5)
mr_cat.tick_labels.font.color.rgb = C_GREY

# Value axis (X)
mr_val = mr_chart.value_axis
mr_val.maximum_scale = 100.0   # same scale as ME so bars are directly comparable
mr_val.minimum_scale = 0.0
mr_val.has_major_gridlines = False
mr_val.tick_labels.font.size = Pt(7)
set_val_axis_number_format(mr_val, "0")

set_plot_area_gap(mr_chart, 70)
set_overlap(mr_chart, -10)

# ══════════════════════════════════════════════════════════════════════════════
# ME CHART
# ══════════════════════════════════════════════════════════════════════════════
me_data = ChartData()
me_data.categories = rev_full
me_data.add_series("Q3'25", rev_me_q3)
me_data.add_series("Q4'25", rev_me_q4)

me_chart_shape = slide.shapes.add_chart(
    XL_CHART_TYPE.BAR_CLUSTERED,
    Inches(ME_LEFT), Inches(CHART_TOP), Inches(ME_W), Inches(CHART_H),
    me_data
)
me_chart = me_chart_shape.chart
me_chart.has_title  = False
me_chart.has_legend = False

# Series colours
for i, s in enumerate(me_chart.series):
    s.format.fill.solid()
    s.format.fill.fore_color.rgb = [C_RYB_Q3, C_RYB_Q4][i]
    set_series_no_border(s)

# Data labels on both Q3 and Q4 series
q3_me = me_chart.series[0]
q3_me.data_labels.show_value        = True
q3_me.data_labels.number_format     = "0"
q3_me.data_labels.font.size         = Pt(7)
q3_me.data_labels.font.color.rgb    = C_FTGREY
set_datalabel_pos_outside_end(q3_me)

q4_me = me_chart.series[1]
q4_me.data_labels.show_value        = True
q4_me.data_labels.number_format     = "0"
q4_me.data_labels.font.size         = Pt(7)
q4_me.data_labels.font.color.rgb    = C_GREY
set_datalabel_pos_outside_end(q4_me)

# Category axis – HIDDEN (labels already on MR chart)
hide_cat_labels(me_chart)
me_cat = me_chart.category_axis
me_cat.tick_labels.font.size = Pt(1)   # belt-and-suspenders: near-invisible

# Value axis (X)
me_val = me_chart.value_axis
me_val.maximum_scale = 100.0
me_val.minimum_scale = 0.0    # same baseline as MR so bars are directly comparable
me_val.has_major_gridlines = False
me_val.tick_labels.font.size = Pt(7)
set_val_axis_number_format(me_val, "0")

set_plot_area_gap(me_chart, 70)
set_overlap(me_chart, -10)

# ══════════════════════════════════════════════════════════════════════════════
# DELTA TABLES  (single Δ column, visual top-to-bottom = highest MR first)
# ══════════════════════════════════════════════════════════════════════════════
add_delta_col(slide, mr_delta,
              MR_DELTA_L, CHART_TOP, MR_DELTA_W, DELTA_H, "MR Δ")

add_delta_col(slide, me_delta,
              ME_DELTA_L, CHART_TOP, ME_DELTA_W, DELTA_H, "ME Δ")

# ══════════════════════════════════════════════════════════════════════════════
# MANUAL LEGEND  (shared, centred below charts)
# ══════════════════════════════════════════════════════════════════════════════
LEG_TOP = CHART_TOP + CHART_H + 0.10
LEG_CTR = (MR_LEFT + ME_DELTA_L + ME_DELTA_W) / 2   # centre of content area
LEG_L   = LEG_CTR - 2.2

solidrect(slide, LEG_L,        LEG_TOP,       0.20, 0.14, C_RYB_Q4)
textbox(slide, f"Q4'25  (n={RYB_N4})",
        LEG_L + 0.25, LEG_TOP - 0.01, 1.20, 0.18, fsize=8, color=C_GREY)

solidrect(slide, LEG_L + 1.65, LEG_TOP,       0.20, 0.14, C_RYB_Q3)
textbox(slide, f"Q3'25  (n={RYB_N3})",
        LEG_L + 1.90, LEG_TOP - 0.01, 1.20, 0.18, fsize=8, color=C_GREY)

solidrect(slide, LEG_L + 3.35, LEG_TOP,       0.20, 0.14, C_GREEN)
textbox(slide, "Δ increase (vs Q3)",
        LEG_L + 3.60, LEG_TOP - 0.01, 1.40, 0.18, fsize=8, color=C_GREY)

solidrect(slide, LEG_L + 5.15, LEG_TOP,       0.20, 0.14, C_RED)
textbox(slide, "Δ decrease (vs Q3)   N/A = new message (no Q3)",
        LEG_L + 5.40, LEG_TOP - 0.01, 3.00, 0.18, fsize=8, color=C_GREY)

# ══════════════════════════════════════════════════════════════════════════════
# FOOTER
# ══════════════════════════════════════════════════════════════════════════════
# Footer positioned below the J&J logo (logo bottom ~7.20"); slide height = 7.50"
FOOTER = (
    f"† ME = geometric mean composite effectiveness (Top-2 Box).  "
    f"* NCCN messages new Nov 2025 — no Q3 baseline (N/A).  "
    f"Source: RYB IM Survey (738902) | Q3 n={RYB_N3} | Q4 n={RYB_N4}.  "
    f"Delta = Q4'25 minus Q3'25 (pp). Sorted by Message Recall Q4 descending."
)
textbox(slide, FOOTER,
        0.15, 7.20, 13.0, 0.28,
        fsize=6.0, color=C_FTGREY, align=PP_ALIGN.LEFT)

# ══════════════════════════════════════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════════════════════════════════════
prs.save(OUT)
print(f"\nSaved: {OUT}")
print(f"   Slide count: {len(prs.slides)}")
print(f"   Messages: {len(msgs)}")
print(f"\nMessage order (top to bottom in chart):")
for i, m in enumerate(msgs):
    delta_str = f"{m['mr_delta']:+.1f}" if m['mr_delta'] is not None else "N/A"
    print(f"  {i+1:2d}.  MR={m['mr_q4']:5.1f}%  ME={m['me_q4'] or 0:5.1f}%  dMR={delta_str:>6}  {m['tag']}")
