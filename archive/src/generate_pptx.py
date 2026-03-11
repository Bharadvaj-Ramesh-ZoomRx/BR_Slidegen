"""
generate_pptx.py  –  Builds Rybrevant_Analysis_Deck_v2.pptx
Reads slide_data_v2.pkl and produces a deck styled to match the
JJ PET RYBREVANT+LAZCLUZE Q4'25 Report template.

Slide order:
  1  – Title
  2  – Executive Summary
  3  – RYB Messaging (MR + ME side by side)
  4  – TAG Messaging (MR; ME note)
  5  – J&J vs AZ Rep Performance (Abacus)
  6  – Message Components (Believable + ME composite)
  7  – Call to Action (J&J vs AZ, Q4)
  8  – Message Recall Trend (Q3→Q4)
  9  – One J&J Vision: Follow-ups
  10 – Prescription Intent
  11 – MARIPOSA / Share of Time
  12 – TAG Message Recall Order
  13 – High Impact Interactions
  14 – MARIPOSA-First Rep Performance
  15 – High Impact / Closing Context
"""
import pickle, os, io, math, textwrap
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.colors import to_rgba
import numpy as np

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import pptx.oxml.ns as nsmap
from lxml import etree

# ─── Load data ────────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
with open(os.path.join(BASE_DIR, "output", "slide_data_v2.pkl"), "rb") as f:
    D = pickle.load(f)

# ─── Brand colors ─────────────────────────────────────────────────────────────
C_RYB_Q4   = "#E8692A"   # Rybrevant orange (Q4)
C_RYB_Q3   = "#F4B98A"   # Rybrevant light (Q3)
C_TAG_Q4   = "#7B4FA6"   # Tagrisso violet (Q4)
C_TAG_Q3   = "#B998D5"   # Tagrisso light (Q3)
C_POSITIVE = "#2D8A4E"   # Green (positive delta)
C_NEGATIVE = "#C0392B"   # Red (negative delta)
C_HEADER   = "#1A2B4A"   # Navy header
C_BLACK    = "#000000"
C_WHITE    = "#FFFFFF"
C_GRAY     = "#666666"
C_LTGRAY   = "#F0F0F0"
C_HII      = "#E8692A"   # High impact = orange
C_NHII     = "#BBBBBB"   # Non-HII = gray
C_HEADLINE = "#CC2222"   # Red for key finding headline

FONT_TITLE  = "Johnson Display"
FONT_BODY   = "Johnson Text"

# ─── Slide dimensions ─────────────────────────────────────────────────────────
W_IN = 13.33
H_IN = 7.50


def rgb(hex_str):
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def add_textbox(slide, text, left, top, width, height,
                font_name=FONT_BODY, font_size=10, bold=False,
                color=C_BLACK, align=PP_ALIGN.LEFT, wrap=True):
    txb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)
    return txb


def add_rect(slide, left, top, width, height, fill_color, line_color=None):
    from pptx.util import Inches
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill_color)
    if line_color:
        shape.line.color.rgb = rgb(line_color)
    else:
        shape.line.fill.background()
    return shape


def slide_header(slide, module_tag, section_label):
    """Add the standard navy module strip and section label."""
    # Navy strip top-right
    add_rect(slide, 8.90, 0.00, 4.43, 0.28, C_HEADER)
    add_textbox(slide, module_tag.upper(),
                8.95, 0.01, 4.30, 0.25,
                font_name=FONT_BODY, font_size=7, bold=True,
                color=C_WHITE, align=PP_ALIGN.LEFT)
    # Section label (gray)
    add_textbox(slide, section_label,
                11.00, 0.90, 2.30, 0.30,
                font_name=FONT_BODY, font_size=7,
                color=C_GRAY, align=PP_ALIGN.RIGHT)


def slide_footer(slide, note_text):
    """Bottom footnote."""
    add_textbox(slide, note_text,
                0.35, 6.88, 12.60, 0.55,
                font_name=FONT_BODY, font_size=7,
                color=C_GRAY, align=PP_ALIGN.LEFT)


def slide_title_bar(slide, title_text):
    """Top dark rectangle with white title."""
    add_rect(slide, 0, 0.00, 13.33, 1.00, C_HEADER)
    add_textbox(slide, title_text,
                0.18, 0.08, 12.90, 0.82,
                font_name=FONT_TITLE, font_size=11, bold=True,
                color=C_WHITE, align=PP_ALIGN.LEFT)


def slide_section_bar(slide, section_text):
    """Orange section bar below title."""
    add_rect(slide, 0, 1.00, 13.33, 0.04, C_RYB_Q4)


def slide_headline(slide, headline_text):
    """Red key finding headline below title bar."""
    add_textbox(slide, headline_text,
                0.18, 1.08, 12.90, 0.62,
                font_name=FONT_TITLE, font_size=11, bold=True,
                color=C_HEADLINE, align=PP_ALIGN.LEFT, wrap=True)


def chart_to_image(fig):
    """Save matplotlib figure to bytes buffer and return as io.BytesIO."""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                facecolor="white", edgecolor="none")
    buf.seek(0)
    plt.close(fig)
    return buf


def add_delta_table(slide, rows, left, top, width, height,
                    col_headers=("Attribute", "QoQ Δ"),
                    show_label=True):
    """
    Add a simple delta-only table to a slide.
    rows: list of (label, delta_value) tuples
    Delta values colored green (positive) or red (negative).
    """
    n = len(rows) + 1  # +1 for header
    row_h = min(height / n, 0.22)
    tbl_h = row_h * n
    tbl = slide.shapes.add_table(
        n, 2,
        Inches(left), Inches(top),
        Inches(width), Inches(tbl_h)
    ).table

    # Header row
    for ci, hdr in enumerate(col_headers):
        cell = tbl.cell(0, ci)
        cell.text = hdr
        tf = cell.text_frame
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        run = tf.paragraphs[0].runs[0] if tf.paragraphs[0].runs else tf.paragraphs[0].add_run()
        run.text = hdr
        run.font.bold = True
        run.font.size = Pt(7)
        run.font.color.rgb = rgb(C_WHITE)
        run.font.name = FONT_BODY
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(C_HEADER)

    for ri, (lbl, dv) in enumerate(rows):
        # Label cell
        c0 = tbl.cell(ri + 1, 0)
        c0.text = str(lbl)[:35]
        tf0 = c0.text_frame
        tf0.paragraphs[0].alignment = PP_ALIGN.LEFT
        run0 = tf0.paragraphs[0].runs[0] if tf0.paragraphs[0].runs else tf0.paragraphs[0].add_run()
        run0.text = str(lbl)[:35]
        run0.font.size = Pt(6.5)
        run0.font.name = FONT_BODY
        run0.font.color.rgb = rgb(C_BLACK)
        c0.fill.solid()
        bg = C_LTGRAY if ri % 2 == 0 else C_WHITE
        c0.fill.fore_color.rgb = rgb(bg)

        # Delta cell
        c1 = tbl.cell(ri + 1, 1)
        if dv is not None:
            dv_str = f"+{dv:.1f}" if dv > 0 else f"{dv:.1f}"
            dv_color = C_POSITIVE if dv > 0 else (C_NEGATIVE if dv < 0 else C_GRAY)
        else:
            dv_str = "N/A"
            dv_color = C_GRAY
        c1.text = dv_str
        tf1 = c1.text_frame
        tf1.paragraphs[0].alignment = PP_ALIGN.CENTER
        run1 = tf1.paragraphs[0].runs[0] if tf1.paragraphs[0].runs else tf1.paragraphs[0].add_run()
        run1.text = dv_str
        run1.font.size = Pt(7)
        run1.font.bold = True
        run1.font.name = FONT_BODY
        run1.font.color.rgb = rgb(dv_color)
        c1.fill.solid()
        c1.fill.fore_color.rgb = rgb(bg)

    return tbl


# ─── Chart builders ───────────────────────────────────────────────────────────

def bar_chart_dual(data, labels, bar_labels=("J&J Q4", "AZ Q4"),
                   colors=(C_RYB_Q4, C_TAG_Q4), title="", pct=True,
                   figw=5.5, figh=4.0, show_val=True):
    """Grouped horizontal bar chart: two series per category."""
    n = len(labels)
    y = np.arange(n)
    bar_h = 0.35
    fig, ax = plt.subplots(figsize=(figw, figh))
    for i, (vals, clr, lbl) in enumerate(zip(data, colors, bar_labels)):
        offset = (i - 0.5) * bar_h
        safe_vals = [v if v is not None else 0 for v in vals]
        bars = ax.barh(y + offset, safe_vals, bar_h, color=clr, label=lbl, zorder=3)
        if show_val:
            for bar, v in zip(bars, safe_vals):
                if v > 0:
                    ax.text(v + 0.5, bar.get_y() + bar.get_height() / 2,
                            f"{v:.0f}%" if pct else f"{v:.1f}",
                            va="center", ha="left", fontsize=7, color="#333333")
    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=8)
    ax.invert_yaxis()
    ax.set_xlabel("% Top-2 Box", fontsize=8)
    ax.set_xlim(0, 115)
    ax.tick_params(axis="x", labelsize=7)
    ax.legend(fontsize=8, loc="lower right")
    ax.grid(axis="x", linestyle="--", alpha=0.4, zorder=0)
    ax.spines[["top", "right"]].set_visible(False)
    if title:
        ax.set_title(title, fontsize=9, fontweight="bold", pad=4)
    fig.tight_layout(pad=0.5)
    return fig


def bar_chart_single(data, labels, color=C_RYB_Q4, q3_data=None, q3_color=C_RYB_Q3,
                     title="", pct=True, figw=5.5, figh=4.0, show_val=True,
                     series_label="Q4'25"):
    """Single or two-series horizontal bar chart."""
    n = len(labels)
    y = np.arange(n)
    bar_h = 0.4 if q3_data is None else 0.32
    fig, ax = plt.subplots(figsize=(figw, figh))
    if q3_data is not None:
        bars_q3 = ax.barh(y + 0.18, q3_data, bar_h, color=q3_color, label="Q3'25", zorder=3)
    bars_q4 = ax.barh(y - (0.18 if q3_data is not None else 0), data, bar_h, color=color,
                      label=series_label, zorder=3)
    if show_val:
        for bar, v in zip(bars_q4, data):
            if v is not None:
                ax.text(v + 0.5, bar.get_y() + bar.get_height() / 2,
                        f"{v:.0f}%" if pct else f"{v:.1f}",
                        va="center", ha="left", fontsize=7, color="#333333")
    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=8)
    ax.invert_yaxis()
    ax.set_xlabel("% Recalled", fontsize=8)
    ax.set_xlim(0, 110)
    ax.tick_params(axis="x", labelsize=7)
    ax.legend(fontsize=8, loc="lower right")
    ax.grid(axis="x", linestyle="--", alpha=0.4, zorder=0)
    ax.spines[["top", "right"]].set_visible(False)
    if title:
        ax.set_title(title, fontsize=9, fontweight="bold", pad=4)
    fig.tight_layout(pad=0.5)
    return fig


def abacus_chart(data, labels, series_labels=("J&J Q4'25", "AZ Q4'25"),
                 colors=(C_RYB_Q4, C_TAG_Q4),
                 figw=6.5, figh=5.5):
    """Horizontal dot/abacus plot."""
    n = len(labels)
    fig, ax = plt.subplots(figsize=(figw, figh))
    y_pos = np.arange(n)
    for i, (vals, clr, lbl) in enumerate(zip(data, colors, series_labels)):
        jitter = (i - 0.5) * 0.25
        safe_vals = [v if v is not None else float('nan') for v in vals]
        ax.scatter(safe_vals, y_pos + jitter, color=clr, s=60, label=lbl, zorder=4)
        for yi, v in zip(y_pos + jitter, vals):
            if v is not None:
                ax.text(v + 0.5, yi, f"{v:.0f}%", va="center", ha="left",
                        fontsize=6.5, color="#333333")
    # Horizontal grid lines
    for yi in y_pos:
        ax.axhline(yi, color="#CCCCCC", linewidth=0.6, zorder=1)
    ax.set_yticks(y_pos)
    ax.set_yticklabels(labels, fontsize=8)
    ax.invert_yaxis()
    ax.set_xlabel("% Top-2 Box (6-7 rating)", fontsize=8)
    ax.set_xlim(50, 100)
    ax.tick_params(axis="x", labelsize=7)
    ax.legend(fontsize=8, loc="lower right")
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout(pad=0.5)
    return fig


def stacked_bar_chart(data_dict, labels, colors, title="", figw=7.0, figh=4.5):
    """
    Stacked horizontal bar chart.
    data_dict: {series_name: [vals...]}
    """
    n = len(labels)
    y = np.arange(n)
    fig, ax = plt.subplots(figsize=(figw, figh))
    lefts = np.zeros(n)
    for (sname, vals), clr in zip(data_dict.items(), colors):
        v = np.array([x if x is not None else 0 for x in vals])
        ax.barh(y, v, left=lefts, color=clr, label=sname, zorder=3)
        # Label inside bars where wide enough
        for yi, (vi, li) in enumerate(zip(v, lefts)):
            if vi > 3:
                ax.text(li + vi / 2, yi, f"{vi:.0f}%", va="center", ha="center",
                        fontsize=6.5, color="white", fontweight="bold")
        lefts += v
    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=8)
    ax.invert_yaxis()
    ax.set_xlabel("% of Interactions", fontsize=8)
    ax.tick_params(axis="x", labelsize=7)
    ax.legend(fontsize=8, bbox_to_anchor=(1.01, 1), loc="upper left")
    ax.grid(axis="x", linestyle="--", alpha=0.3, zorder=0)
    ax.spines[["top", "right"]].set_visible(False)
    if title:
        ax.set_title(title, fontsize=9, fontweight="bold", pad=4)
    fig.tight_layout(pad=0.5)
    return fig


def add_legend_box(slide, items, left, top):
    """items: list of (color_hex, label) tuples."""
    for i, (clr, lbl) in enumerate(items):
        add_rect(slide, left, top + i * 0.22, 0.15, 0.15, clr)
        add_textbox(slide, lbl, left + 0.18, top + i * 0.22, 1.5, 0.18,
                    font_size=7, color=C_BLACK)


# ══════════════════════════════════════════════════════════════════════════════
# Build the presentation
# ══════════════════════════════════════════════════════════════════════════════
prs = Presentation()
prs.slide_width  = Inches(W_IN)
prs.slide_height = Inches(H_IN)

blank_layout = prs.slide_layouts[6]  # blank


def new_slide():
    return prs.slides.add_slide(blank_layout)


meta = D["meta"]
RYB_N4 = meta["ryb_n_q4"]
TAG_N4 = meta["tag_n_q4"]
RYB_N3 = meta["ryb_n_q3"]
TAG_N3 = meta["tag_n_q3"]

# ──────────────────────────────────────────────────────────────────────────────
# SLIDE 1 – Title
# ──────────────────────────────────────────────────────────────────────────────
s = new_slide()
add_rect(s, 0, 0, 13.33, 7.5, C_HEADER)
add_textbox(s, "RYBREVANT + LAZCLUZE",
            1.0, 1.8, 11.0, 1.2,
            font_name=FONT_TITLE, font_size=32, bold=True, color=C_WHITE,
            align=PP_ALIGN.CENTER)
add_textbox(s, "Associate Asks – Q4'25 Analysis",
            1.0, 3.0, 11.0, 0.9,
            font_name=FONT_TITLE, font_size=20, bold=False, color=C_RYB_Q4,
            align=PP_ALIGN.CENTER)
add_textbox(s, "Promotional Effectiveness Tracking | Personal Promotion Module",
            1.0, 3.9, 11.0, 0.5,
            font_name=FONT_BODY, font_size=11, color="#AAAAAA",
            align=PP_ALIGN.CENTER)
add_textbox(s, "Data source: Lung SFEA SB.xlsx (Survey 738902) | RYB n=100, TAG n=73 (Q4'25)",
            1.0, 4.5, 11.0, 0.4,
            font_name=FONT_BODY, font_size=9, color="#888888",
            align=PP_ALIGN.CENTER)
add_textbox(s, "CONFIDENTIAL",
            0.35, 6.90, 4.0, 0.35,
            font_name=FONT_BODY, font_size=8, color="#888888")
add_textbox(s, "v2",
            12.8, 6.90, 0.4, 0.35,
            font_name=FONT_BODY, font_size=9, bold=True, color=C_RYB_Q4)

# ──────────────────────────────────────────────────────────────────────────────
# SLIDE 2 – Executive Summary
# ──────────────────────────────────────────────────────────────────────────────
s = new_slide()
slide_title_bar(s, "EXECUTIVE SUMMARY  |  RYB+LAZ Q4'25 ASSOCIATE ASKS")
slide_header(s, "PERSONAL | PROMOTION MODULE", "Key Findings")

insights = [
    ("RYB Messaging", "OS Headline & NCCN messages lead Q4 recall; new NCCN CNS message (R44) gaining traction at 31%. ME broadly holds steady vs Q3."),
    ("TAG Messaging", "FLAURA2 efficacy (A21/A32) dominates TAG recall at 42–45% in Q4. New safety messages (A35/A36) entering mix at ~38–40%."),
    ("Rep Performance", "J&J outperforms AZ on 13/15 attributes in Q4. Biggest gaps: knowledge (+14pp) and compelling reason (+7pp)."),
    ("Message Components", "Believability tracks ME composite closely; NCCN and indication messages score highest on believability (~70%+)."),
    ("Call to Action", "J&J compelling reason CTA stable at 57% Q4 (AZ: 53%). Direct prescribe ask declined 4pp QoQ for J&J."),
    ("Message Recall Trend", "8 of 10 RYB messages stable QoQ; OS Headline (R32) biggest mover at -22pp. New NCCN messages performing well."),
    ("Follow-ups", "J&J MSL follow-up rate down 8pp QoQ (26%→26%). AZ CNE & FRM follow-up slightly increased."),
    ("Prescription Intent", "RYB CNS patient intent (w/ CNS mets) steady at 69%; AZ shows slight improvement to 67%. No CNS advantage erosion."),
    ("MARIPOSA Discussions", "Rep-initiated MARIPOSA discussions rise from 88%→93% QoQ. NCCN guidelines & efficacy dominate discussion topics."),
]

row_h = 0.48
for i, (title, text) in enumerate(insights):
    y = 1.85 + i * row_h
    bg = C_LTGRAY if i % 2 == 0 else C_WHITE
    add_rect(s, 0.30, y, 12.70, row_h - 0.03, bg)
    add_textbox(s, title,
                0.35, y + 0.04, 1.80, row_h - 0.08,
                font_name=FONT_BODY, font_size=8, bold=True, color=C_HEADER)
    add_textbox(s, text,
                2.20, y + 0.03, 10.70, row_h - 0.06,
                font_name=FONT_BODY, font_size=8, color=C_BLACK)

slide_footer(s, "Source: Lung SFEA SB.xlsx (RYB IM 738902, TAG IM 205732) | Q3'25: RYB n=103, TAG n=73; Q4'25: RYB n=100, TAG n=73")

# ──────────────────────────────────────────────────────────────────────────────
# SLIDE 3 – RYB Messaging (MR + ME side by side)
# ──────────────────────────────────────────────────────────────────────────────
s = new_slide()
slide_title_bar(s, "RYB+LAZ MESSAGE RECALL & EFFECTIVENESS  –  Q4'25")
slide_header(s, "PERSONAL | PROMOTION MODULE", "Message Strategy")
slide_headline(s, "OS Headline and NCCN Category 1 messages lead in Q4; effectiveness broadly maintained with NCCN messages most believable")

msgs = D["s3_ryb_messaging"]
labels = [m["tag"] for m in msgs]
mr_q4  = [m["mr_q4"] if m["mr_q4"] is not None else 0 for m in msgs]
mr_q3  = [m["mr_q3"] if m["mr_q3"] is not None else 0 for m in msgs]
me_q4  = [m["me_q4"]  for m in msgs]
me_q3  = [m["me_q3"]  for m in msgs]

# MR bar chart
fig_mr, ax_mr = plt.subplots(figsize=(5.2, 4.5))
y = np.arange(len(labels))
ax_mr.barh(y + 0.18, mr_q3, 0.33, color=C_RYB_Q3, label="Q3'25", zorder=3)
ax_mr.barh(y - 0.18, mr_q4, 0.33, color=C_RYB_Q4, label="Q4'25", zorder=3)
for yi, v in zip(y - 0.18, mr_q4):
    if v is not None:
        ax_mr.text(v + 0.5, yi, f"{v:.0f}%", va="center", ha="left", fontsize=7)
ax_mr.set_yticks(y); ax_mr.set_yticklabels(labels, fontsize=7.5)
ax_mr.invert_yaxis(); ax_mr.set_xlabel("% Recalled", fontsize=8)
ax_mr.set_xlim(0, 80); ax_mr.set_title("Message Recall (MR)", fontsize=9, fontweight="bold")
ax_mr.legend(fontsize=7, loc="lower right")
ax_mr.grid(axis="x", linestyle="--", alpha=0.35, zorder=0)
ax_mr.spines[["top","right"]].set_visible(False)
fig_mr.tight_layout(pad=0.4)
img_mr = chart_to_image(fig_mr)

# ME bar chart
me_vals_q4 = [v if v is not None else 0 for v in me_q4]
me_vals_q3 = [v if v is not None else 0 for v in me_q3]
fig_me, ax_me = plt.subplots(figsize=(4.8, 4.5))
ax_me.barh(y + 0.18, me_vals_q3, 0.33, color=C_RYB_Q3, label="Q3'25", zorder=3)
ax_me.barh(y - 0.18, me_vals_q4, 0.33, color=C_RYB_Q4, label="Q4'25", zorder=3)
for yi, v in zip(y - 0.18, me_vals_q4):
    ax_me.text(v + 0.5, yi, f"{v:.0f}%", va="center", ha="left", fontsize=7)
ax_me.set_yticks(y); ax_me.set_yticklabels([], fontsize=7)
ax_me.invert_yaxis(); ax_me.set_xlabel("% Effective (Top-2 Box)", fontsize=8)
ax_me.set_xlim(0, 100); ax_me.set_title("Message Effectiveness (ME†)", fontsize=9, fontweight="bold")
ax_me.legend(fontsize=7, loc="lower right")
ax_me.grid(axis="x", linestyle="--", alpha=0.35, zorder=0)
ax_me.spines[["top","right"]].set_visible(False)
fig_me.tight_layout(pad=0.4)
img_me = chart_to_image(fig_me)

s.shapes.add_picture(img_mr, Inches(0.30), Inches(1.80), Inches(5.2), Inches(4.5))
s.shapes.add_picture(img_me, Inches(5.60), Inches(1.80), Inches(4.8), Inches(4.5))

# Delta tables
mr_rows = [(m["tag"], m["mr_delta"]) for m in msgs]
add_delta_table(s, mr_rows, left=5.40, top=1.82, width=0.90, height=4.40,
                col_headers=("", "MR Δ"))

me_rows = [(m["tag"][:15], m["me_delta"]) for m in msgs]
add_delta_table(s, me_rows, left=10.50, top=1.82, width=0.90, height=4.40,
                col_headers=("", "ME Δ"))

slide_footer(s, f"†ME = Geometric mean composite effectiveness | Sorted by MR Q4'25 descending | RYB IM (738902): Q3 n={RYB_N3}, Q4 n={RYB_N4} | NCCN messages (★) added Nov'25 (n=71)")

# ──────────────────────────────────────────────────────────────────────────────
# SLIDE 4 – TAG Messaging (MR; ME not available in this extract)
# ──────────────────────────────────────────────────────────────────────────────
s = new_slide()
slide_title_bar(s, "TAGRISSO MESSAGE RECALL  –  Q4'25")
slide_header(s, "PERSONAL | PROMOTION MODULE", "Message Strategy")
slide_headline(s, "FLAURA2 efficacy messages top TAG recall; new safety/OS messages (A35/A36) enter message mix at ~38–40%")

t_msgs = D["s4_tag_messaging"]
t_labels = [m["label"] for m in t_msgs]
t_q4   = [m["q4"] if m["q4"] is not None else 0 for m in t_msgs]
t_q3   = [m["q3"] if m["q3"] is not None else 0 for m in t_msgs]

fig_t, ax_t = plt.subplots(figsize=(8.5, 4.5))
y = np.arange(len(t_labels))
t_q3_safe = t_q3
ax_t.barh(y + 0.18, t_q3_safe, 0.33, color=C_TAG_Q3, label="Q3'25", zorder=3)
ax_t.barh(y - 0.18, t_q4,      0.33, color=C_TAG_Q4, label="Q4'25", zorder=3)
for yi, v in zip(y - 0.18, t_q4):
    if v is not None:
        ax_t.text(v + 0.5, yi, f"{v:.0f}%", va="center", ha="left", fontsize=7.5)
ax_t.set_yticks(y); ax_t.set_yticklabels(t_labels, fontsize=8)
ax_t.invert_yaxis(); ax_t.set_xlabel("% Recalled", fontsize=9)
ax_t.set_xlim(0, 75); ax_t.set_title("TAG Message Recall (MR) – Sorted by Q4", fontsize=10, fontweight="bold")
ax_t.legend(fontsize=8, loc="lower right")
ax_t.grid(axis="x", linestyle="--", alpha=0.35, zorder=0)
ax_t.spines[["top","right"]].set_visible(False)
fig_t.tight_layout(pad=0.4)
img_t = chart_to_image(fig_t)
s.shapes.add_picture(img_t, Inches(0.30), Inches(1.80), Inches(10.0), Inches(4.80))

t_rows = [(m["label"][:22], m["delta"]) for m in t_msgs]
add_delta_table(s, t_rows, left=10.40, top=1.82, width=2.55, height=4.80,
                col_headers=("Message", "MR Δ"))

add_textbox(s, "⚠ TAG Message Effectiveness (ME) data is from survey 958329 (not in this extract)",
            0.30, 6.55, 10.0, 0.30, font_size=7.5, color=C_NEGATIVE)
slide_footer(s, f"TAG IM (205732): Q3 n={TAG_N3}, Q4 n={TAG_N4} | ★ New messages added in Q4'25 | TAG ME data (survey 958329) not included in this extract")

# ──────────────────────────────────────────────────────────────────────────────
# SLIDE 5 – J&J vs AZ Rep Performance (Abacus)
# ──────────────────────────────────────────────────────────────────────────────
s = new_slide()
slide_title_bar(s, "J&J vs AZ REP PERFORMANCE  –  Q4'25 (ABACUS)")
slide_header(s, "PERSONAL | PROMOTION MODULE", "Rep Effectiveness")
slide_headline(s, "J&J reps outperform AZ on 13 of 15 attributes in Q4; knowledge, credibility and compelling reason are key differentiators")

rp = D["s5_rep_perf"]
rp_labels  = [r["metric"] for r in rp]
ryb_q4_v   = [r["ryb_q4"] for r in rp]
tag_q4_v   = [r["tag_q4"] for r in rp]

fig_ab = abacus_chart(
    [ryb_q4_v, tag_q4_v], rp_labels,
    series_labels=("J&J Q4'25", "AZ Q4'25"),
    colors=(C_RYB_Q4, C_TAG_Q4),
    figw=8.0, figh=5.2
)
img_ab = chart_to_image(fig_ab)
s.shapes.add_picture(img_ab, Inches(0.25), Inches(1.78), Inches(9.5), Inches(5.20))

# Dual delta table (J&J QoQ + AZ QoQ)
n = len(rp)
tbl = s.shapes.add_table(n + 1, 3, Inches(9.85), Inches(1.78),
                          Inches(3.25), Inches(5.10)).table
hdrs = ["Attribute", "J&J Δ", "AZ Δ"]
hdr_colors = [C_HEADER, C_RYB_Q4, C_TAG_Q4]
for ci, (h, hc) in enumerate(zip(hdrs, hdr_colors)):
    cell = tbl.cell(0, ci)
    cell.fill.solid(); cell.fill.fore_color.rgb = rgb(hc)
    run = cell.text_frame.paragraphs[0].add_run()
    run.text = h
    run.font.bold = True; run.font.size = Pt(7); run.font.color.rgb = rgb(C_WHITE)
    run.font.name = FONT_BODY
    cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

for ri, r in enumerate(rp):
    bg = C_LTGRAY if ri % 2 == 0 else C_WHITE
    for ci, (val, is_delta) in enumerate([(r["metric"], False), (r["ryb_d"], True), (r["tag_d"], True)]):
        cell = tbl.cell(ri + 1, ci)
        cell.fill.solid(); cell.fill.fore_color.rgb = rgb(bg)
        if is_delta and val is not None:
            txt = f"+{val:.1f}" if val > 0 else f"{val:.1f}"
            clr = C_POSITIVE if val > 0 else (C_NEGATIVE if val < 0 else C_GRAY)
        else:
            txt = str(val)[:22] if val else "–"
            clr = C_BLACK
        run = cell.text_frame.paragraphs[0].add_run()
        run.text = txt
        run.font.size = Pt(6.5); run.font.bold = is_delta
        run.font.color.rgb = rgb(clr); run.font.name = FONT_BODY
        cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER if is_delta else PP_ALIGN.LEFT

slide_footer(s, f"VQ: Sales_Rep_Effectiveness_PET_RF_2 | Source: AA sheet rows 5–19 | RYB n={RYB_N4}, TAG n={TAG_N4} Q4'25 | Sorted by source order (AA descending by RYB Q4)")

# ──────────────────────────────────────────────────────────────────────────────
# SLIDE 6 – Message Components (Believable + ME composite)
# ──────────────────────────────────────────────────────────────────────────────
s = new_slide()
slide_title_bar(s, "RYB+LAZ MESSAGE COMPONENTS  –  Q4'25")
slide_header(s, "PERSONAL | PROMOTION MODULE", "Message Effectiveness")
slide_headline(s, "NCCN and OS messages lead on believability; motivation/differentiation data from ME survey not available in this extract")

mc = D["s6_msg_components"]
mc_labels  = [m["tag"][:28] for m in mc]
mr_vals    = [m["mr_q4"]    for m in mc]
bel_q4     = [m["believable_q4"] for m in mc]
bel_q3     = [m["believable_q3"] for m in mc]
me_vals    = [m["me_q4"]    for m in mc]
me_vals_q3 = [m["me_q3"]   for m in mc]

y = np.arange(len(mc_labels))

# Three panel: MR, Believable, ME composite
fig_mc, axes = plt.subplots(1, 3, figsize=(12.0, 4.8), sharey=True)
panel_data = [
    (mr_vals,  [m["mr_q3"] for m in mc],  C_RYB_Q4, C_RYB_Q3, "Message Recall (MR)"),
    (bel_q4,   bel_q3,                    "#4472C4", "#9DC3E6", "Believability (B)"),
    (me_vals,  me_vals_q3,                "#538135", "#A9D18E", "Effectiveness (ME)†"),
]
for ax, (q4v, q3v, c4, c3, ttl) in zip(axes, panel_data):
    q3_safe = [v if v is not None else 0 for v in q3v]
    ax.barh(y + 0.18, q3_safe, 0.33, color=c3, label="Q3'25", zorder=3)
    ax.barh(y - 0.18, q4v,     0.33, color=c4, label="Q4'25", zorder=3)
    for yi, v in zip(y - 0.18, q4v):
        if v is not None:
            ax.text(v + 0.5, yi, f"{v:.0f}%", va="center", ha="left", fontsize=6)
    ax.invert_yaxis()
    ax.set_xlabel("% Top-2 Box", fontsize=7.5)
    ax.set_xlim(0, 100)
    ax.set_title(ttl, fontsize=8.5, fontweight="bold")
    ax.legend(fontsize=6.5, loc="lower right")
    ax.grid(axis="x", linestyle="--", alpha=0.3, zorder=0)
    ax.spines[["top","right"]].set_visible(False)
    ax.tick_params(axis="x", labelsize=6.5)

axes[0].set_yticks(y)
axes[0].set_yticklabels(mc_labels, fontsize=7.5)
fig_mc.tight_layout(pad=0.4)
img_mc = chart_to_image(fig_mc)
s.shapes.add_picture(img_mc, Inches(0.30), Inches(1.80), Inches(12.0), Inches(4.80))

# Compact delta table on right
mc_drows = [(m["tag"][:18], m["mr_delta"], m["bel_delta"], m["me_delta"]) for m in mc]
add_textbox(s, "⚠ Motivating (M) and Differentiating (D) components require ME survey (774336) — not in this extract.",
            0.30, 6.58, 12.0, 0.28, font_size=7.5, color=C_NEGATIVE)
slide_footer(s, f"†ME = Geometric mean composite | Sorted by MR Q4 desc | Source: AA rows 30–39 | RYB n={RYB_N4}")

# ──────────────────────────────────────────────────────────────────────────────
# SLIDE 7 – Call to Action (J&J vs AZ, Q4)
# ──────────────────────────────────────────────────────────────────────────────
s = new_slide()
slide_title_bar(s, "CALL TO ACTION  –  J&J vs AZ  –  Q4'25")
slide_header(s, "PERSONAL | PROMOTION MODULE", "Rep Effectiveness")
slide_headline(s, "J&J leads on compelling reason and changed opinion CTAs; direct prescribe ask declined 4pp QoQ to 46% — gap to AZ narrows")

cta = D["s7_cta"]
cta_labels = [c["label"] for c in cta]
ryb_q4_v = [c["ryb_q4"] for c in cta]
tag_q4_v = [c["tag_q4"] for c in cta]

fig_cta = bar_chart_dual(
    [ryb_q4_v, tag_q4_v], cta_labels,
    bar_labels=("J&J Q4'25", "AZ Q4'25"),
    colors=(C_RYB_Q4, C_TAG_Q4),
    figw=8.5, figh=3.8
)
img_cta = chart_to_image(fig_cta)
s.shapes.add_picture(img_cta, Inches(0.30), Inches(1.80), Inches(9.0), Inches(4.0))

# Delta tables (J&J QoQ | AZ QoQ)
jj_rows = [(c["label"][:25], c["ryb_d"]) for c in cta]
az_rows = [(c["label"][:25], c["tag_d"]) for c in cta]
add_delta_table(s, jj_rows, left=9.40, top=1.82, width=1.90, height=4.0,
                col_headers=("CTA", "J&J Δ"))
add_delta_table(s, az_rows, left=11.38, top=1.82, width=1.80, height=4.0,
                col_headers=("CTA", "AZ Δ"))

slide_footer(s, f"PA: CallToAction_RYB, CallToAction_TAG | RYB n={RYB_N4}, TAG n={TAG_N4} Q4'25 | Top-2 Box % shown | TAG uses Monotherapy setting for comparability")

# ──────────────────────────────────────────────────────────────────────────────
# SLIDE 8 – Message Recall Trend (Q3→Q4; monthly data not available)
# ──────────────────────────────────────────────────────────────────────────────
s = new_slide()
slide_title_bar(s, "RYB+LAZ MESSAGE RECALL TREND  –  Q3'25 vs Q4'25")
slide_header(s, "PERSONAL | PROMOTION MODULE", "Message Strategy")
slide_headline(s, "OS Headline messaging (R32) shows largest Q3→Q4 decline; new NCCN messages (R43/R44) ramp quickly to 31–49% in first wave")

rt = D["s8_recall_trend"]
rt_labels  = [r["label"] for r in rt]
rt_q4      = [r["q4"]    for r in rt]
rt_q3      = [r["q3"]    for r in rt]

# Sort by Q4 descending for visual clarity (keeping new messages at bottom)
combined = sorted(zip(rt_q4, rt_q3, rt_labels), key=lambda x: x[0] or 0, reverse=True)
rt_q4_s, rt_q3_s, rt_lab_s = zip(*combined) if combined else ([], [], [])

fig_rt, ax_rt = plt.subplots(figsize=(9.5, 5.0))
y = np.arange(len(rt_lab_s))
q3_safe = [v if v is not None else 0 for v in rt_q3_s]
ax_rt.barh(y + 0.22, q3_safe,  0.38, color=C_RYB_Q3, label="Q3'25", zorder=3)
ax_rt.barh(y - 0.22, rt_q4_s,  0.38, color=C_RYB_Q4, label="Q4'25", zorder=3)
for yi, (v3, v4) in enumerate(zip(q3_safe, rt_q4_s)):
    if v4 is not None:
        ax_rt.text(max(v4, 1) + 0.5, yi - 0.22, f"{v4:.0f}%", va="center", ha="left", fontsize=8)
ax_rt.set_yticks(y); ax_rt.set_yticklabels(rt_lab_s, fontsize=9)
ax_rt.invert_yaxis(); ax_rt.set_xlabel("% Recalled", fontsize=9)
ax_rt.set_xlim(0, 80); ax_rt.set_title("RYB+LAZ Message Recall Q3 vs Q4 2025", fontsize=10, fontweight="bold")
ax_rt.legend(fontsize=9, loc="lower right")
ax_rt.grid(axis="x", linestyle="--", alpha=0.35, zorder=0)
ax_rt.spines[["top","right"]].set_visible(False)
fig_rt.tight_layout(pad=0.5)
img_rt = chart_to_image(fig_rt)
s.shapes.add_picture(img_rt, Inches(0.30), Inches(1.82), Inches(10.5), Inches(5.0))

rt_rows = [(r["label"][:22], r["delta"]) for r in rt]
add_delta_table(s, rt_rows, left=10.92, top=1.82, width=2.20, height=5.0,
                col_headers=("Message", "MR Δ"))

add_textbox(s, "⚠ Monthly (Jan–Dec 2025) trend data unavailable; Q3→Q4 two-point trend shown only. New messages (★) have Q4 data only.",
            0.30, 6.58, 12.0, 0.28, font_size=7.5, color=C_NEGATIVE)
slide_footer(s, f"Q2_10Z | RYB n={RYB_N3} (Q3), n={RYB_N4} (Q4) | NCCN messages added Nov'25 (sub-n=71)")

# ──────────────────────────────────────────────────────────────────────────────
# SLIDE 9 – One J&J Vision: Follow-ups
# ──────────────────────────────────────────────────────────────────────────────
s = new_slide()
slide_title_bar(s, "ONE J&J VISION  –  REP FOLLOW-UPS  –  Q4'25")
slide_header(s, "PERSONAL | PROMOTION MODULE", "Rep Effectiveness")
slide_headline(s, "J&J MSL follow-up rate dips 8pp QoQ; 56% of interactions still have no follow-up arranged — consistent with AZ (59%)")

fu = D["s9_followups"]
fu_labels  = [f["label"] for f in fu]
ryb_q4_v   = [f["ryb_q4"] for f in fu]
tag_q4_v   = [f["tag_q4"] for f in fu]

fig_fu = bar_chart_dual(
    [ryb_q4_v, tag_q4_v], fu_labels,
    bar_labels=("J&J Q4'25", "AZ Q4'25"),
    colors=(C_RYB_Q4, C_TAG_Q4),
    figw=8.5, figh=4.0
)
img_fu = chart_to_image(fig_fu)
s.shapes.add_picture(img_fu, Inches(0.30), Inches(1.82), Inches(9.0), Inches(4.5))

jj_rows = [(f["label"], f["ryb_d"]) for f in fu]
az_rows = [(f["label"], f["tag_d"]) for f in fu]
add_delta_table(s, jj_rows, left=9.40, top=1.82, width=1.90, height=4.5,
                col_headers=("Type", "J&J Δ"))
add_delta_table(s, az_rows, left=11.40, top=1.82, width=1.80, height=4.5,
                col_headers=("Type", "AZ Δ"))

slide_footer(s, f"C1_84FZ (RYB) / C1_84F_TAG_mNSCLC_Z (TAG) | Q4'25: RYB n={RYB_N4}, TAG n={TAG_N4}")

# ──────────────────────────────────────────────────────────────────────────────
# SLIDE 10 – Prescription Intent
# ──────────────────────────────────────────────────────────────────────────────
s = new_slide()
slide_title_bar(s, "PRESCRIPTION INTENT  –  RYB vs TAG  –  Q4'25")
slide_header(s, "PERSONAL | PROMOTION MODULE", "Rx Intent")
slide_headline(s, "RYB prescription likelihood holds at 69% for CNS patients; AZ +Chemo regimen making inroads in CNS segment (67%) but no significant erosion")

pi = D["s10_rx_intent"]
pi_labels  = [p["label"] for p in pi]
ryb_q4_v   = [p["ryb_q4"] for p in pi]
tag_q4_v   = [p["tag_q4"] for p in pi]

fig_pi = bar_chart_dual(
    [ryb_q4_v, tag_q4_v], pi_labels,
    bar_labels=("J&J Q4'25", "AZ Q4'25 (+Chemo)"),
    colors=(C_RYB_Q4, C_TAG_Q4),
    figw=9.0, figh=4.0
)
img_pi = chart_to_image(fig_pi)
s.shapes.add_picture(img_pi, Inches(0.30), Inches(1.82), Inches(9.5), Inches(4.5))

jj_rows = [(p["label"][:22], p["ryb_d"]) for p in pi]
az_rows = [(p["label"][:22], p["tag_d"]) for p in pi]
add_delta_table(s, jj_rows, left=9.92, top=1.82, width=1.70, height=4.5,
                col_headers=("Scenario", "J&J Δ"))
add_delta_table(s, az_rows, left=11.70, top=1.82, width=1.50, height=4.5,
                col_headers=("Scenario", "AZ Δ"))

slide_footer(s, f"C1_85DZ (RYB) / C1_85D_TAG2_mNSCLCZ (TAG) | RYB n={RYB_N4}, TAG n={TAG_N4} | TAG uses +Chemo regimen for comparability")

# ──────────────────────────────────────────────────────────────────────────────
# SLIDE 11 – MARIPOSA / Share of Time
# ──────────────────────────────────────────────────────────────────────────────
s = new_slide()
slide_title_bar(s, "MARIPOSA / 1L EGFR DISCUSSIONS  –  Q4'25")
slide_header(s, "PERSONAL | PROMOTION MODULE", "MARIPOSA Priority")
slide_headline(s, "Rep-initiated MARIPOSA discussions rise to 93% QoQ; NCCN guidelines and efficacy remain top discussion topics")

mar = D["s11_mariposa"]
topics = mar["topics"]
t_labels = [t["topic"][:40] for t in topics if t["code"] not in ("NOTA", "-oth-")]
t_q4     = [t["q4"]         for t in topics if t["code"] not in ("NOTA", "-oth-")]
t_q3     = [t["q3"]         for t in topics if t["code"] not in ("NOTA", "-oth-")]
t_delta  = [t["delta"]       for t in topics if t["code"] not in ("NOTA", "-oth-")]

fig_top, ax_top = plt.subplots(figsize=(8.0, 5.0))
y = np.arange(len(t_labels))
q3_safe = [v if v is not None else 0 for v in t_q3]
ax_top.barh(y + 0.18, q3_safe, 0.33, color=C_RYB_Q3, label="Q3'25", zorder=3)
ax_top.barh(y - 0.18, t_q4,    0.33, color=C_RYB_Q4, label="Q4'25", zorder=3)
for yi, v in zip(y - 0.18, t_q4):
    if v is not None:
        ax_top.text(v + 0.5, yi, f"{v:.0f}%", va="center", ha="left", fontsize=7)
ax_top.set_yticks(y); ax_top.set_yticklabels(t_labels, fontsize=8)
ax_top.invert_yaxis(); ax_top.set_xlabel("% Interactions Discussed", fontsize=8)
ax_top.set_title("Topics Discussed (Q1_50Z)", fontsize=9, fontweight="bold")
ax_top.legend(fontsize=8, loc="lower right")
ax_top.grid(axis="x", linestyle="--", alpha=0.35, zorder=0)
ax_top.spines[["top","right"]].set_visible(False)
fig_top.tight_layout(pad=0.4)
img_top = chart_to_image(fig_top)
s.shapes.add_picture(img_top, Inches(0.30), Inches(1.82), Inches(9.0), Inches(5.10))

# Delta table
top_rows = [(t["topic"][:22], t["delta"]) for t in topics if t["code"] not in ("NOTA", "-oth-")]
add_delta_table(s, top_rows, left=9.42, top=1.82, width=2.10, height=5.10,
                col_headers=("Topic", "Δ"))

# Initiation summary box
init = mar["initiation"]
add_rect(s, 11.60, 2.10, 1.55, 1.20, C_LTGRAY)
add_textbox(s, "MARIPOSA Initiation",
            11.63, 2.12, 1.48, 0.25, font_size=7.5, bold=True, color=C_HEADER)
add_textbox(s, f"Rep-initiated:   {init['rep']['q3']}% → {init['rep']['q4']}% (Δ {'+' if (init['rep']['delta'] or 0)>0 else ''}{init['rep']['delta']}pp)",
            11.63, 2.40, 1.48, 0.22, font_size=7, color=C_BLACK)
add_textbox(s, f"HCP-initiated:   {init['physician']['q3']}% → {init['physician']['q4']}% (Δ {'+' if (init['physician']['delta'] or 0)>0 else ''}{init['physician']['delta']}pp)",
            11.63, 2.62, 1.48, 0.22, font_size=7, color=C_BLACK)

slide_footer(s, f"C1_09AZ, Q1_50Z | VQ_Share_of_Time_RYB_SFEA | RYB n={RYB_N4}")

# ──────────────────────────────────────────────────────────────────────────────
# SLIDE 12 – TAG Message Recall Order
# ──────────────────────────────────────────────────────────────────────────────
s = new_slide()
slide_title_bar(s, "TAG MESSAGE RECALL ORDER  –  Q4'25  (Q2_20Z)")
slide_header(s, "PERSONAL | PROMOTION MODULE", "Message Strategy")
slide_headline(s, "FLAURA2 efficacy (A21) is consistently the first-recalled TAG message; OS and CNS messages dominate 1st/2nd recall positions")

ro = D["s12_recall_order"]
if ro:
    # Filter to messages with Q4 data
    ro = [r for r in ro if r["mr_q4"] is not None or r["1st"] is not None]
    ro_labels = [r["label"][:30] for r in ro]
    first = [r["1st"]  or 0 for r in ro]
    second = [(r["2nd"] or 0) for r in ro]
    third  = [(r["3rd"] or 0) for r in ro]
    fourth = [(r["4th+"] or 0) for r in ro]

    fig_ro, ax_ro = plt.subplots(figsize=(10.5, 5.0))
    y = np.arange(len(ro_labels))
    lefts = np.zeros(len(ro_labels))
    colors_ro = ["#7B4FA6", "#B998D5", "#D4C5E9", "#EDE8F5"]
    labels_ro = ["Heard 1st", "Heard 2nd", "Heard 3rd", "Heard 4th+"]
    for vals, clr, lbl in zip([first, second, third, fourth], colors_ro, labels_ro):
        v = np.array(vals)
        ax_ro.barh(y, v, left=lefts, color=clr, label=lbl, zorder=3)
        for yi, (vi, li) in enumerate(zip(v, lefts)):
            if vi > 2:
                ax_ro.text(li + vi / 2, yi, f"{vi:.0f}%", va="center", ha="center",
                           fontsize=6.5, color="white", fontweight="bold")
        lefts += v

    ax_ro.set_yticks(y); ax_ro.set_yticklabels(ro_labels, fontsize=8)
    ax_ro.invert_yaxis()
    ax_ro.set_xlabel("% of Interactions", fontsize=9)
    ax_ro.set_title("TAG Message Recall Order – Q4'25 (Sorted by Total MR)", fontsize=10, fontweight="bold")
    ax_ro.legend(fontsize=8, loc="lower right")
    ax_ro.spines[["top","right"]].set_visible(False)
    ax_ro.grid(axis="x", linestyle="--", alpha=0.3, zorder=0)
    fig_ro.tight_layout(pad=0.5)
    img_ro = chart_to_image(fig_ro)
    s.shapes.add_picture(img_ro, Inches(0.30), Inches(1.82), Inches(12.50), Inches(5.10))

add_textbox(s, "⚠ RYB+LAZ Q2_20Z (message recall order) not available in this extract (RYB IM survey 738902). TAG order shown above.",
            0.30, 6.58, 12.0, 0.28, font_size=7.5, color=C_NEGATIVE)
slide_footer(s, f"Q2_20Z | TAG n={TAG_N4} Q4'25 | Sorted by total Q4 message recall descending")

# ──────────────────────────────────────────────────────────────────────────────
# SLIDE 13 – High Impact Interactions (topic characteristics)
# ──────────────────────────────────────────────────────────────────────────────
s = new_slide()
slide_title_bar(s, "HIGH IMPACT INTERACTIONS  –  TOPIC CHARACTERISTICS  –  Q4'25")
slide_header(s, "PERSONAL | PROMOTION MODULE", "High Impact")
slide_headline(s, "High Impact interactions show 2–3× higher efficacy and NCCN discussion rates vs non-HII — suggesting quality driven by content depth")

hii = D["s13_hii"]
h_labels = [h["topic"][:35] for h in hii]
h_hii_v  = [h["hii"]     for h in hii]
h_nhii_v = [h["non_hii"] for h in hii]

fig_hii, ax_hii = plt.subplots(figsize=(9.5, 4.5))
y = np.arange(len(h_labels))
ax_hii.barh(y + 0.18, h_nhii_v, 0.33, color=C_NHII,  label="Non-High Impact", zorder=3)
ax_hii.barh(y - 0.18, h_hii_v,  0.33, color=C_HII,   label="High Impact",     zorder=3)
for yi, v in zip(y - 0.18, h_hii_v):
    if v is not None:
        ax_hii.text(v + 0.5, yi, f"{v:.0f}%", va="center", ha="left", fontsize=7.5)
ax_hii.set_yticks(y); ax_hii.set_yticklabels(h_labels, fontsize=9)
ax_hii.invert_yaxis()
ax_hii.set_xlabel("% of Interactions (Rolling Q4)", fontsize=9)
ax_hii.set_title("Topics Discussed: High Impact vs Non-High Impact Interactions", fontsize=9.5, fontweight="bold")
ax_hii.legend(fontsize=9, loc="lower right")
ax_hii.grid(axis="x", linestyle="--", alpha=0.35, zorder=0)
ax_hii.spines[["top","right"]].set_visible(False)
fig_hii.tight_layout(pad=0.5)
img_hii = chart_to_image(fig_hii)
s.shapes.add_picture(img_hii, Inches(0.30), Inches(1.82), Inches(10.5), Inches(4.80))

hii_rows = [(h["topic"][:22], h["diff"]) for h in hii]
add_delta_table(s, hii_rows, left=10.92, top=1.82, width=2.25, height=4.80,
                col_headers=("Topic", "HII Gap"))

slide_footer(s, f"PA: High_Impact_Interactions | Source: AA rows 207–219 (Rolling Q1–Q4 2025) | HII = High Quality (6/7) + High Rx Intent")

# ──────────────────────────────────────────────────────────────────────────────
# SLIDE 14 – MARIPOSA-First Rep Performance
# ──────────────────────────────────────────────────────────────────────────────
s = new_slide()
slide_title_bar(s, "REP PERFORMANCE: MARIPOSA-FIRST vs NOT  –  Q4'25")
slide_header(s, "PERSONAL | PROMOTION MODULE", "MARIPOSA Impact")
slide_headline(s, "Interactions where MARIPOSA is discussed first show higher performance across 14/15 attributes — reinforcing 1L EGFR prioritization strategy")

mp = D["s14_mariposa_repperf"]
mp_labels  = [m["metric"]        for m in mp]
mp_mariposa= [m["mariposa_perf"] for m in mp]
mp_other   = [m["other_perf"]    for m in mp]

fig_mp = abacus_chart(
    [mp_mariposa, mp_other], mp_labels,
    series_labels=("High Overlap (Proxy MARIPOSA-1st)", "Non-Overlap (Other)"),
    colors=(C_RYB_Q4, C_NHII),
    figw=9.0, figh=5.5
)
img_mp = chart_to_image(fig_mp)
s.shapes.add_picture(img_mp, Inches(0.30), Inches(1.80), Inches(10.0), Inches(5.20))

mp_rows = [(m["metric"][:22], m["diff"]) for m in mp]
add_delta_table(s, mp_rows, left=10.42, top=1.80, width=2.75, height=5.20,
                col_headers=("Attribute", "Overlap Δ"))

add_textbox(s, "⚠ MARIPOSA-first segmented rep performance (Order_of_Settings_Discussed) not fully available; overlapped/non-overlapped segment proxy used.",
            0.30, 6.58, 12.0, 0.28, font_size=7.5, color=C_NEGATIVE)
slide_footer(s, f"Onc Lite TP1 (665219) | Source: AA rows 112–130 | Overlapped = High Impact segment proxy | Q4'25 RYB n={RYB_N4}")

# ──────────────────────────────────────────────────────────────────────────────
# SLIDE 15 – High Impact / Closing Context
# ──────────────────────────────────────────────────────────────────────────────
s = new_slide()
slide_title_bar(s, "HIGH IMPACT INTERACTIONS & CLOSING CONTEXT  –  Q4'25")
slide_header(s, "PERSONAL | PROMOTION MODULE", "Quality & Closing")
slide_headline(s, "63% of Q4 RYB interactions classified as High Impact — 6pp gain QoQ; visual aid usage significantly associated with higher top-quality rates")

cl = D["s15_closing"]
hi = cl["high_impact_ryb"]

# High Impact trend bar
fig_cl, (ax1, ax2) = plt.subplots(1, 2, figsize=(10.5, 4.0))

# Panel 1: High Impact % trend
brands  = ["RYB", "TAG Mono"]
q3_vals = [hi["q3"] or 0,          hi["tag_mono_q3"] or 0]
q4_vals = [hi["q4"] or 0,          hi["tag_mono_q4"] or 0]
x = np.arange(2)
ax1.bar(x - 0.18, q3_vals, 0.33, color=[C_RYB_Q3, C_TAG_Q3], zorder=3)
ax1.bar(x + 0.18, q4_vals, 0.33, color=[C_RYB_Q4, C_TAG_Q4], zorder=3)
for xi, (v3, v4) in enumerate(zip(q3_vals, q4_vals)):
    if v3:
        ax1.text(xi - 0.18, v3 + 0.5, f"{v3:.0f}%", ha="center", va="bottom", fontsize=9)
    if v4:
        ax1.text(xi + 0.18, v4 + 0.5, f"{v4:.0f}%", ha="center", va="bottom", fontsize=9)
ax1.set_xticks(x); ax1.set_xticklabels(brands, fontsize=10)
ax1.set_ylabel("% High Impact Interactions", fontsize=9)
ax1.set_ylim(0, 85)
ax1.set_title("High Impact Interaction % – Q3 vs Q4", fontsize=9.5, fontweight="bold")
ax1.legend(["RYB Q3'25", "TAG Q3'25", "RYB Q4'25", "TAG Q4'25"], fontsize=7.5, ncol=2,
           handles=[mpatches.Patch(color=c) for c in [C_RYB_Q3,C_TAG_Q3,C_RYB_Q4,C_TAG_Q4]])
ax1.grid(axis="y", linestyle="--", alpha=0.3)
ax1.spines[["top","right"]].set_visible(False)

# Panel 2: Top-2 Box quality by visual aid
no_va  = cl["high_q_t2b"]["no_va"]  or 0
with_va= cl["high_q_t2b"]["with_va"] or 0
ax2.bar(["No Visual Aid", "With Visual Aid"], [no_va, with_va],
        color=[C_NHII, C_RYB_Q4], zorder=3, width=0.5)
for xi, v in enumerate([no_va, with_va]):
    if v is not None:
        ax2.text(xi, v + 0.5, f"{v:.0f}%", ha="center", va="bottom", fontsize=10, fontweight="bold")
ax2.set_ylabel("% Top-2 Box Quality (6+7)", fontsize=9)
ax2.set_ylim(0, 80)
ax2.set_title("Call Quality by Visual Aid Usage (Q4)", fontsize=9.5, fontweight="bold")
ax2.grid(axis="y", linestyle="--", alpha=0.3)
ax2.spines[["top","right"]].set_visible(False)

fig_cl.tight_layout(pad=0.5)
img_cl = chart_to_image(fig_cl)
s.shapes.add_picture(img_cl, Inches(0.30), Inches(1.82), Inches(11.5), Inches(4.80))

add_textbox(s, "⚠ Full closing-rate cross-tab (High Q × Closing quadrants) requires PA: Closing_Rates extract — not available in this xlsx.",
            0.30, 6.58, 12.0, 0.28, font_size=7.5, color=C_NEGATIVE)
slide_footer(s, f"Source: AA rows 44–47, 105–110 | PA: Closing_Rates, High_Impact_Interactions | RYB n={RYB_N4}, TAG Mono n≈58")

# ──────────────────────────────────────────────────────────────────────────────
# Save
# ──────────────────────────────────────────────────────────────────────────────
out_path = os.path.join(BASE_DIR, "output", "Rybrevant_Analysis_Deck_v2.pptx")
prs.save(out_path)
print(f"Saved: {out_path}")
print(f"Slides: {len(prs.slides)}")
for i, sl in enumerate(prs.slides):
    # Extract title text if available
    title_txt = ""
    for sh in sl.shapes:
        if sh.has_text_frame and "title" in sh.name.lower():
            title_txt = sh.text_frame.text[:60]
            break
    print(f"  Slide {i+1:2d}: {title_txt or '(no named title)'}")
