"""
generate_pptx.py
Generates the Rybrevant Associate Asks 9-slide PowerPoint deck.
Requires slide_data.pkl (produced by extract_data.py).

Layout per content slide:
  - Slide title bar  (top 0-0.7")
  - Chart image      (0.7"-5.3")
  - Delta table      (5.3"-6.8")
  - Footer           (6.8"-7.5")
"""
import pickle, os, io, math, textwrap
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyArrowPatch
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.oxml.ns import qn
from lxml import etree
import copy

# ─── Paths ────────────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PKL_PATH = os.path.join(BASE_DIR, "slide_data.pkl")
OUT_PATH = os.path.join(BASE_DIR, "Rybrevant_Analysis_Deck.pptx")

with open(PKL_PATH, "rb") as f:
    D = pickle.load(f)

META  = D["meta"]

# ─── Color Palette ────────────────────────────────────────────────────────────
RYB_COLOR    = "#E8692A"      # Rybrevant orange
TAG_COLOR    = "#7B4FA6"      # Tagrisso violet
RYB_LIGHT    = "#F4B98A"      # Q3 / lighter orange
TAG_LIGHT    = "#B998D5"      # Q3 / lighter violet
GREEN_DELTA  = "#2D8A4E"
RED_DELTA    = "#C0392B"
GREY_DELTA   = "#7F8C8D"
BG_WHITE     = "#FFFFFF"
HEADER_BG    = "#1A2B4A"      # dark navy header

def hex2rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def rgb_color(h):
    r, g, b = hex2rgb(h)
    return RGBColor(r, g, b)

# ─── Presentation Setup ───────────────────────────────────────────────────────
prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]   # blank layout

# ─── Helpers: pptx primitives ─────────────────────────────────────────────────
def add_textbox(slide, text, left, top, width, height,
                font_size=11, bold=False, color="#000000",
                align=PP_ALIGN.LEFT, wrap=True, italic=False):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = rgb_color(color)
    return tb

def add_rect(slide, left, top, width, height, fill_hex, line_hex=None):
    shape = slide.shapes.add_shape(
        1, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb_color(fill_hex)
    if line_hex:
        shape.line.color.rgb = rgb_color(line_hex)
    else:
        shape.line.fill.background()
    return shape

def add_image_from_buf(slide, buf, left, top, width, height):
    buf.seek(0)
    slide.shapes.add_picture(buf, Inches(left), Inches(top), Inches(width), Inches(height))

def set_cell_fill(cell, hex_color):
    from pptx.oxml.ns import qn
    from lxml import etree
    tc = cell._tc
    # Remove existing fills
    for old in tc.findall(qn('a:solidFill'), tc.nsmap):
        tc.remove(old)
    spPr = tc.get_or_add_spPr() if hasattr(tc, 'get_or_add_spPr') else None
    # Use XML directly
    r, g, b = hex2rgb(hex_color)
    color_str = f"{r:02X}{g:02X}{b:02X}"
    fill_xml = (
        f'<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:srgbClr val="{color_str}"/>'
        f'</a:solidFill>'
    )
    fill_el = etree.fromstring(fill_xml)
    tc.insert(0, fill_el)

def delta_color(val):
    if val is None: return GREY_DELTA
    if val > 0:     return GREEN_DELTA
    if val < 0:     return RED_DELTA
    return GREY_DELTA

def delta_str(val):
    if val is None: return "—"
    sign = "+" if val > 0 else ""
    return f"{sign}{val:.1f}pp"

def pct_str(val):
    if val is None: return "N/A"
    return f"{val:.1f}%"

# ─── Slide Header ─────────────────────────────────────────────────────────────
def add_slide_header(slide, title, subtitle=""):
    add_rect(slide, 0, 0, 13.33, 0.65, HEADER_BG)
    add_textbox(slide, title, 0.15, 0.04, 9.5, 0.55,
                font_size=16, bold=True, color="#FFFFFF", align=PP_ALIGN.LEFT)
    if subtitle:
        add_textbox(slide, subtitle, 9.7, 0.1, 3.5, 0.45,
                    font_size=9, color="#CCDDEE", align=PP_ALIGN.RIGHT)

def add_slide_footer(slide, note="", ryb_n_q3=103, ryb_n_q4=100, tag_n_q3=73, tag_n_q4=73, show_tag=True):
    add_rect(slide, 0, 6.95, 13.33, 0.55, "#F0F2F5")
    ryb_txt = f"J&J (Rybrevant): N={ryb_n_q3} (Q3) / N={ryb_n_q4} (Q4)"
    tag_txt = f"  |  AZ (Tagrisso): N={tag_n_q3} (Q3) / N={tag_n_q4} (Q4)" if show_tag else ""
    base = f"Source: Lung SFEA SB.xlsx  |  {ryb_txt}{tag_txt}"
    full = f"{base}  |  {note}" if note else base
    add_textbox(slide, full, 0.15, 7.0, 13.0, 0.45,
                font_size=7, color="#444444", align=PP_ALIGN.LEFT)

# ─── Delta Table ──────────────────────────────────────────────────────────────
def add_delta_table(slide, headers, rows,
                    left=0.15, top=5.35, width=13.0, row_height=0.22,
                    delta_col_idx=None):
    """
    headers: list of column header strings
    rows: list of tuples; each value is (text, bg_color_hex or None)
    delta_col_idx: list of column indices that contain delta values (for auto-coloring)
    """
    n_cols = len(headers)
    n_rows = len(rows)
    col_w  = [Inches(width / n_cols)] * n_cols
    table = slide.shapes.add_table(
        n_rows + 1, n_cols,
        Inches(left), Inches(top),
        Inches(width), Inches((n_rows + 1) * row_height)
    ).table
    table.horz_banding = False

    # Header row
    for ci, hdr in enumerate(headers):
        cell = table.cell(0, ci)
        cell.text = hdr
        cell.text_frame.paragraphs[0].runs[0].font.size = Pt(8)
        cell.text_frame.paragraphs[0].runs[0].font.bold = True
        cell.text_frame.paragraphs[0].runs[0].font.color.rgb = rgb_color("#FFFFFF")
        cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        set_cell_fill(cell, HEADER_BG)
        table.columns[ci].width = col_w[ci]

    # Data rows
    for ri, row in enumerate(rows):
        for ci, (text, bg) in enumerate(row):
            cell = table.cell(ri + 1, ci)
            cell.text = str(text)
            tf = cell.text_frame
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            run = tf.paragraphs[0].runs[0]
            run.font.size = Pt(8)
            if bg:
                set_cell_fill(cell, bg)
                if bg in (GREEN_DELTA, RED_DELTA, HEADER_BG):
                    run.font.color.rgb = rgb_color("#FFFFFF")
                else:
                    run.font.color.rgb = rgb_color("#222222")
            else:
                run.font.color.rgb = rgb_color("#222222")


# ══════════════════════════════════════════════════════════════════════════════
# TITLE SLIDE
# ══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(blank_layout)
add_rect(slide, 0, 0, 13.33, 7.5, HEADER_BG)
add_rect(slide, 0.3, 2.0, 12.73, 0.08, RYB_COLOR)

add_textbox(slide, "RYBREVANT® (amivantamab-vmjw) + LAZCLUZE® (lazertinib)",
            0.5, 1.2, 12.0, 0.8, font_size=22, bold=True, color="#FFFFFF", align=PP_ALIGN.CENTER)
add_textbox(slide, "Associate Asks Analysis  |  Q4 2025",
            0.5, 2.3, 12.0, 0.7, font_size=18, bold=False, color="#F0C8A0", align=PP_ALIGN.CENTER)
add_textbox(slide, "J&J (Rybrevant) vs AstraZeneca (Tagrisso)  |  Lung SFEA Survey 738902",
            0.5, 3.1, 12.0, 0.5, font_size=12, color="#AABBCC", align=PP_ALIGN.CENTER)

# Legend
for x_off, color, label in [(3.8, RYB_COLOR, "Rybrevant (J&J)"), (7.5, TAG_COLOR, "Tagrisso (AZ)")]:
    add_rect(slide, x_off, 4.2, 0.3, 0.22, color)
    add_textbox(slide, label, x_off + 0.35, 4.15, 3.0, 0.32, font_size=12, color="#FFFFFF")

add_textbox(slide, "Straight Ask Analyses  |  9 Slides",
            0.5, 5.5, 12.0, 0.4, font_size=10, color="#667788", align=PP_ALIGN.CENTER)
add_textbox(slide, "Confidential – For Internal Use Only",
            0.5, 6.8, 12.0, 0.4, font_size=9, italic=True, color="#445566", align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
# CHART HELPERS
# ══════════════════════════════════════════════════════════════════════════════
CHART_W = 12.8   # inches
CHART_H = 4.4    # inches

def fig_to_buf(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                facecolor="white", edgecolor="none")
    plt.close(fig)
    buf.seek(0)
    return buf

def wrap_label(text, width=22):
    return "\n".join(textwrap.wrap(text, width))


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 – RYB + Lazcluze Messaging
# ══════════════════════════════════════════════════════════════════════════════
data1 = D["slide1_ryb_messaging"]
tags1  = [d["tag"] for d in data1]
mr_q3  = [d["mr_q3"]  or 0 for d in data1]
mr_q4  = [d["mr_q4"]  or 0 for d in data1]
me_q3  = [d["me_q3"]  or 0 for d in data1]
me_q4  = [d["me_q4"]  or 0 for d in data1]

fig, axes = plt.subplots(1, 2, figsize=(CHART_W, CHART_H))
fig.suptitle("", fontsize=1)

for ax, q3_vals, q4_vals, title, color_q3, color_q4, unit in [
    (axes[0], mr_q3, mr_q4, "Message Recall (MR)", RYB_LIGHT, RYB_COLOR, "%"),
    (axes[1], me_q3, me_q4, "Message Effectiveness (ME – Top 2 Box)", "#C5A0C5", TAG_COLOR, "%"),
]:
    y = np.arange(len(tags1))
    h = 0.35
    bars_q3 = ax.barh(y + h/2, q3_vals, h, label="Q3 2025", color=color_q3, alpha=0.85)
    bars_q4 = ax.barh(y - h/2, q4_vals, h, label="Q4 2025", color=color_q4, alpha=0.95)
    ax.set_yticks(y)
    ax.set_yticklabels([wrap_label(t, 28) for t in tags1], fontsize=8.5)
    ax.set_xlabel("% (Top 2 Box / Recall)", fontsize=8)
    ax.set_title(title, fontsize=10, fontweight="bold", pad=6)
    ax.set_xlim(0, 105)
    ax.axvline(0, color="#CCCCCC", linewidth=0.5)
    ax.grid(axis="x", linestyle="--", alpha=0.4, color="#DDDDDD")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for bar in bars_q4:
        w = bar.get_width()
        if w > 0:
            ax.text(w + 1.2, bar.get_y() + bar.get_height()/2,
                    f"{w:.0f}%", va="center", ha="left", fontsize=7.5)
    ax.legend(fontsize=8, loc="lower right")

plt.tight_layout(rect=[0, 0, 1, 1])
buf1 = fig_to_buf(fig)

slide = prs.slides.add_slide(blank_layout)
add_slide_header(slide, "Ask #1 – RYB + Lazcluze Messaging  |  Q3 vs Q4 2025",
                 subtitle="Source: Additional Analysis | N=103 (Q3) / N=100 (Q4)")
add_image_from_buf(slide, buf1, 0.25, 0.72, CHART_W, CHART_H)

# Delta table
hdrs1 = ["Message Tag", "MR Q3", "MR Q4", "MR Δ", "ME Q3", "ME Q4", "ME Δ"]
tbl1_rows = []
for d in data1:
    dc  = delta_color(d["mr_delta"])
    dc2 = delta_color(d["me_delta"])
    tbl1_rows.append([
        (d["tag"][:30],     None),
        (pct_str(d["mr_q3"]), None), (pct_str(d["mr_q4"]), None),
        (delta_str(d["mr_delta"]), dc),
        (pct_str(d["me_q3"]), None), (pct_str(d["me_q4"]), None),
        (delta_str(d["me_delta"]), dc2),
    ])
add_delta_table(slide, hdrs1, tbl1_rows, top=5.25, row_height=0.19)
add_slide_footer(slide, show_tag=False, note="*R43/R44 new in Q4 2025; Q3 MR not applicable")


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 – Tagrisso Messaging
# ══════════════════════════════════════════════════════════════════════════════
data2  = D["slide2_tag_messaging"]
labels2 = [wrap_label(d["label"], 30) for d in data2]
q3_2 = [d["q3"] or 0 for d in data2]
q4_2 = [d["q4"] or 0 for d in data2]

fig, ax = plt.subplots(figsize=(CHART_W, CHART_H + 0.5))
y = np.arange(len(data2))
h = 0.35
ax.barh(y + h/2, q3_2, h, label="Q3 2025", color=TAG_LIGHT, alpha=0.85)
bars = ax.barh(y - h/2, q4_2, h, label="Q4 2025", color=TAG_COLOR, alpha=0.95)
ax.set_yticks(y)
ax.set_yticklabels(labels2, fontsize=8.5)
ax.set_xlabel("Message Recall % (Total)", fontsize=9)
ax.set_title("Tagrisso (osimertinib) Message Recall – Q3 vs Q4 2025", fontsize=11, fontweight="bold")
ax.set_xlim(0, 75)
ax.grid(axis="x", linestyle="--", alpha=0.4, color="#DDDDDD")
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
for bar in bars:
    w = bar.get_width()
    if w > 0:
        ax.text(w + 0.8, bar.get_y() + bar.get_height()/2,
                f"{w:.0f}%", va="center", ha="left", fontsize=7.5)
ax.legend(fontsize=9, loc="lower right")
plt.tight_layout()
buf2 = fig_to_buf(fig)

slide = prs.slides.add_slide(blank_layout)
add_slide_header(slide, "Ask #2 – Tagrisso Messaging  |  Q3 vs Q4 2025",
                 subtitle="Source: TAG Sheet (Q2_10Z) | N=73 (Q3) / N=73 (Q4)")
add_image_from_buf(slide, buf2, 0.25, 0.72, CHART_W, CHART_H + 0.2)

hdrs2 = ["Message", "Q3 Recall", "Q4 Recall", "Δ (pp)"]
tbl2_rows = []
for d in data2:
    tbl2_rows.append([
        (d["label"][:35], None),
        (pct_str(d["q3"]), None),
        (pct_str(d["q4"]), None),
        (delta_str(d["delta"]), delta_color(d["delta"])),
    ])
add_delta_table(slide, hdrs2, tbl2_rows, top=5.1, row_height=0.172)
add_slide_footer(slide, show_tag=False, note="*A35/A36 new in Q4; Q3 N/A",
                 ryb_n_q3=None, ryb_n_q4=None, tag_n_q3=73, tag_n_q4=73)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 – Rep Performance (Abacus / Dumbbell chart)
# ══════════════════════════════════════════════════════════════════════════════
data3 = D["slide3_rep_perf"]
metrics3 = [d["metric"] for d in data3]
ryb_q4_3 = [d["ryb_q4"] or 0 for d in data3]
tag_q4_3 = [d["tag_q4"] or 0 for d in data3]
ryb_q3_3 = [d["ryb_q3"] or 0 for d in data3]
tag_q3_3 = [d["tag_q3"] or 0 for d in data3]

fig, ax = plt.subplots(figsize=(CHART_W, CHART_H + 0.8))
y = np.arange(len(data3))

# Draw connecting lines (dumbbell)
for i in range(len(data3)):
    lo = min(ryb_q4_3[i], tag_q4_3[i])
    hi = max(ryb_q4_3[i], tag_q4_3[i])
    ax.hlines(y[i], lo, hi, colors="#CCCCCC", linewidths=1.5, zorder=1)

# Q3 hollow dots
ax.scatter(ryb_q3_3, y, s=40, color=RYB_LIGHT, zorder=2, marker="o",
           edgecolors=RYB_COLOR, linewidths=0.8, alpha=0.6, label="RYB Q3")
ax.scatter(tag_q3_3, y, s=40, color=TAG_LIGHT, zorder=2, marker="o",
           edgecolors=TAG_COLOR, linewidths=0.8, alpha=0.6, label="TAG Q3")

# Q4 filled dots
ax.scatter(ryb_q4_3, y, s=90, color=RYB_COLOR, zorder=3, marker="o", label="RYB Q4 (J&J)")
ax.scatter(tag_q4_3, y, s=90, color=TAG_COLOR, zorder=3, marker="D", label="TAG Q4 (AZ)")

ax.set_yticks(y)
ax.set_yticklabels([wrap_label(m, 35) for m in metrics3], fontsize=8)
ax.set_xlabel("Top 2 Box Score (7-pt scale)", fontsize=9)
ax.set_title("J&J vs AZ Rep Performance – Q4 2025 (Abacus)", fontsize=11, fontweight="bold")
ax.set_xlim(55, 100)
ax.grid(axis="x", linestyle="--", alpha=0.35, color="#DDDDDD")
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.legend(fontsize=8, loc="lower right", ncol=2)
for i, (rv, tv) in enumerate(zip(ryb_q4_3, tag_q4_3)):
    ax.text(rv + 0.5, y[i] + 0.18, f"{rv:.0f}%", fontsize=7, color=RYB_COLOR, ha="left", va="bottom")
    ax.text(tv + 0.5, y[i] - 0.28, f"{tv:.0f}%", fontsize=7, color=TAG_COLOR, ha="left", va="top")
plt.tight_layout()
buf3 = fig_to_buf(fig)

slide = prs.slides.add_slide(blank_layout)
add_slide_header(slide, "Ask #3 – J&J vs AZ Rep Performance  |  Q3 vs Q4 2025",
                 subtitle="Source: Additional Analysis | N=103→100 (J&J) / N=73 (AZ)")
add_image_from_buf(slide, buf3, 0.25, 0.72, CHART_W, CHART_H + 0.5)

hdrs3 = ["Metric", "RYB Q3", "RYB Q4", "RYB Δ", "TAG Q3", "TAG Q4", "TAG Δ"]
tbl3_rows = []
for d in data3:
    tbl3_rows.append([
        (d["metric"][:35], None),
        (pct_str(d["ryb_q3"]), None), (pct_str(d["ryb_q4"]), None),
        (delta_str(d["ryb_d"]), delta_color(d["ryb_d"])),
        (pct_str(d["tag_q3"]), None), (pct_str(d["tag_q4"]), None),
        (delta_str(d["tag_d"]), delta_color(d["tag_d"])),
    ])
add_delta_table(slide, hdrs3, tbl3_rows, top=5.35, row_height=0.172)
add_slide_footer(slide)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 – Message Component Analysis (MR vs ME, Q4)
# ══════════════════════════════════════════════════════════════════════════════
data4 = D["slide4_msg_components"]
tags4   = [d["tag"] for d in data4]
mr_q4_4 = [d["mr_q4"]     or 0 for d in data4]
me_b_4  = [d["me_bel_q4"] or 0 for d in data4]
me_t_4  = [d["me_tb_q4"]  or 0 for d in data4]

fig, ax = plt.subplots(figsize=(CHART_W, CHART_H))
y = np.arange(len(data4))
h = 0.28

ax.barh(y + h,   mr_q4_4, h, label="Message Recall (MR)",          color=RYB_COLOR, alpha=0.9)
ax.barh(y,       me_b_4,  h, label="ME – Believable %",            color=TAG_LIGHT, alpha=0.9)
ax.barh(y - h,   me_t_4,  h, label="ME – Top 2 Box (overall)",    color=TAG_COLOR, alpha=0.9)

ax.set_yticks(y)
ax.set_yticklabels([wrap_label(t, 28) for t in tags4], fontsize=8.5)
ax.set_xlabel("Q4 2025 %", fontsize=9)
ax.set_title("Message Component Analysis – Q4 2025 (MR / Believable / ME Top 2 Box)", fontsize=10, fontweight="bold")
ax.set_xlim(0, 105)
ax.grid(axis="x", linestyle="--", alpha=0.35, color="#DDDDDD")
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.legend(fontsize=8.5, loc="lower right")
plt.tight_layout()
buf4 = fig_to_buf(fig)

slide = prs.slides.add_slide(blank_layout)
add_slide_header(slide, "Ask #4 – Message Component Analysis (MR / Believable / ME)  |  Q4 2025",
                 subtitle="Source: Additional Analysis | N=100 (Q4)")
add_image_from_buf(slide, buf4, 0.25, 0.72, CHART_W, CHART_H)

hdrs4 = ["Message", "MR Q4", "ME Bel Q4", "ME T2B Q4"]
tbl4_rows = [[
    (d["tag"][:30], None),
    (pct_str(d["mr_q4"]), None),
    (pct_str(d["me_bel_q4"]), None),
    (pct_str(d["me_tb_q4"]), None),
] for d in data4]
add_delta_table(slide, hdrs4, tbl4_rows, top=5.28, row_height=0.19)
add_slide_footer(slide, show_tag=False, note="ME = Message Effectiveness. Believable = single-attribute score.")


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 – J&J vs AZ Call to Action
# ══════════════════════════════════════════════════════════════════════════════
data5 = D["slide5_cta"]

fig, axes = plt.subplots(1, 2, figsize=(CHART_W, 3.5))

for ax_idx, period, label in [(0, "q3", "Q3 2025"), (1, "q4", "Q4 2025")]:
    ax = axes[ax_idx]
    labels5 = [d["label"] for d in data5]
    ryb_vals = [d[f"ryb_{period}"] or 0 for d in data5]
    tag_vals = [d[f"tag_{period}"] or 0 for d in data5]
    x = np.arange(len(labels5))
    w = 0.35
    ax.bar(x - w/2, ryb_vals, w, label="J&J (Rybrevant)", color=RYB_COLOR, alpha=0.9)
    ax.bar(x + w/2, tag_vals, w, label="AZ (Tagrisso)",   color=TAG_COLOR, alpha=0.9)
    ax.set_xticks(x)
    ax.set_xticklabels([wrap_label(l, 14) for l in labels5], fontsize=8.5)
    ax.set_ylabel("Yes % (Incidence)", fontsize=8)
    ax.set_ylim(0, 80)
    ax.set_title(f"CTA Incidence – {label}", fontsize=10, fontweight="bold")
    ax.grid(axis="y", linestyle="--", alpha=0.35, color="#DDDDDD")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for rect, val in zip(ax.patches, ryb_vals + tag_vals):
        if val > 0:
            ax.text(rect.get_x() + rect.get_width()/2, rect.get_height() + 0.5,
                    f"{val:.0f}%", ha="center", va="bottom", fontsize=7.5)
    ax.legend(fontsize=8)
plt.tight_layout()
buf5 = fig_to_buf(fig)

slide = prs.slides.add_slide(blank_layout)
add_slide_header(slide, "Ask #5 – J&J vs AZ Call to Action  |  Q3 & Q4 2025",
                 subtitle="Source: RYB / TAG Sheets | N(J&J)=103→100, N(AZ)=73")
add_image_from_buf(slide, buf5, 0.25, 0.72, CHART_W, 3.6)

hdrs5 = ["CTA Type", "J&J Q3", "J&J Q4", "J&J Δ", "AZ Q3", "AZ Q4", "AZ Δ"]
tbl5_rows = []
for d in data5:
    tbl5_rows.append([
        (d["label"], None),
        (pct_str(d["ryb_q3"]), None), (pct_str(d["ryb_q4"]), None),
        (delta_str(d["ryb_d"]), delta_color(d["ryb_d"])),
        (pct_str(d["tag_q3"]), None), (pct_str(d["tag_q4"]), None),
        (delta_str(d["tag_d"]), delta_color(d["tag_d"])),
    ])
add_delta_table(slide, hdrs5, tbl5_rows, top=4.5, row_height=0.3)
add_slide_footer(slide, note="TAG values = Tagrisso Monotherapy setting. Yes% = incidence of rep performing that CTA.")


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 – RYB Message Recall Trend (Q3 → Q4)
# ══════════════════════════════════════════════════════════════════════════════
data6 = D["slide6_recall_trend"]
# Separate Q3-available vs Q4-only
data6_both = [d for d in data6 if d["q3"] is not None]
data6_new  = [d for d in data6 if d["q3"] is None]

fig, ax = plt.subplots(figsize=(CHART_W, CHART_H))
x_pts = [0, 1]
x_labels = ["Q3 2025", "Q4 2025"]
cmap = plt.get_cmap("tab10")

for i, d in enumerate(data6_both):
    color = cmap(i % 10)
    ax.plot(x_pts, [d["q3"], d["q4"]], marker="o", markersize=8,
            linewidth=2, color=color, label=d["label"][:30])
    ax.text(1.02, d["q4"], f"  {d['q4']:.0f}%", va="center", fontsize=7.5, color=color)

# New Q4 messages (plot only Q4 point)
for d in data6_new:
    ax.scatter([1], [d["q4"]], marker="*", s=100, color="gray", zorder=5)
    ax.text(1.02, d["q4"] - 1, f"  ★ {d['label'][:20]} {d['q4']:.0f}%",
            va="top", fontsize=7, color="gray")

ax.set_xticks(x_pts)
ax.set_xticklabels(x_labels, fontsize=10)
ax.set_ylabel("Message Recall %", fontsize=9)
ax.set_title("RYB + Lazcluze Message Recall Trend  |  Q3 → Q4 2025", fontsize=11, fontweight="bold")
ax.set_ylim(0, 70)
ax.grid(axis="y", linestyle="--", alpha=0.4, color="#DDDDDD")
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.legend(fontsize=7.5, bbox_to_anchor=(1.01, 1), loc="upper left", ncol=1)
plt.tight_layout()
buf6 = fig_to_buf(fig)

slide = prs.slides.add_slide(blank_layout)
add_slide_header(slide, "Ask #6 – RYB Message Recall Trend  |  Q3 → Q4 2025",
                 subtitle="Source: RYB Sheet (Q2_10Z) | N=103 (Q3) / N=100 (Q4)")
add_image_from_buf(slide, buf6, 0.25, 0.72, CHART_W, CHART_H + 0.3)

hdrs6 = ["Message", "Q3 Recall", "Q4 Recall", "Δ (pp)"]
tbl6_rows = [[
    (d["label"][:35], None),
    (pct_str(d["q3"]), None),
    (pct_str(d["q4"]), None),
    (delta_str(d["delta"]), delta_color(d["delta"])),
] for d in data6]
add_delta_table(slide, hdrs6, tbl6_rows, top=5.25, row_height=0.172)
add_slide_footer(slide, show_tag=False,
                 note="⚠ Monthly trend data unavailable. Two-point Q3→Q4 trend shown. ★ = new Q4 message.")


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 – One J&J Vision – Follow-ups
# ══════════════════════════════════════════════════════════════════════════════
data7  = D["slide7_followups"]
labels7 = [d["label"] for d in data7]
ryb_q3_7 = [d["ryb_q3"] or 0 for d in data7]
ryb_q4_7 = [d["ryb_q4"] or 0 for d in data7]
tag_q3_7 = [d["tag_q3"] or 0 for d in data7]
tag_q4_7 = [d["tag_q4"] or 0 for d in data7]

fig, axes = plt.subplots(1, 2, figsize=(CHART_W, 3.8))
for ax, q3_r, q4_r, q3_t, q4_t, period in [
    (axes[0], ryb_q3_7, None, tag_q3_7, None, "Q3 2025"),
    (axes[1], None, ryb_q4_7, None, tag_q4_7, "Q4 2025"),
]:
    vals_ryb = q3_r if q3_r is not None else q4_r
    vals_tag = q3_t if q3_t is not None else q4_t
    x = np.arange(len(labels7))
    w = 0.35
    ax.bar(x - w/2, vals_ryb, w, label="J&J (Rybrevant)", color=RYB_COLOR, alpha=0.9)
    ax.bar(x + w/2, vals_tag, w, label="AZ (Tagrisso)",   color=TAG_COLOR, alpha=0.9)
    ax.set_xticks(x)
    ax.set_xticklabels([wrap_label(l, 10) for l in labels7], fontsize=8.5)
    ax.set_ylabel("% respondents", fontsize=8)
    ax.set_ylim(0, 75)
    ax.set_title(f"Follow-up Setup – {period}", fontsize=10, fontweight="bold")
    ax.grid(axis="y", linestyle="--", alpha=0.35, color="#DDDDDD")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for rect in ax.patches:
        h_val = rect.get_height()
        if h_val > 0:
            ax.text(rect.get_x() + rect.get_width()/2, h_val + 0.5,
                    f"{h_val:.0f}%", ha="center", va="bottom", fontsize=7.5)
    ax.legend(fontsize=8)
plt.tight_layout()
buf7 = fig_to_buf(fig)

slide = prs.slides.add_slide(blank_layout)
add_slide_header(slide, "Ask #7 – One J&J Vision: Follow-up Connect Setup  |  Q3 vs Q4 2025",
                 subtitle="Source: RYB/TAG Sheets (C1_84FZ) | N(J&J)=103→100, N(AZ)=73")
add_image_from_buf(slide, buf7, 0.25, 0.72, CHART_W, 4.0)

hdrs7 = ["Follow-up Type", "J&J Q3", "J&J Q4", "J&J Δ", "AZ Q3", "AZ Q4", "AZ Δ"]
tbl7_rows = []
for d in data7:
    tbl7_rows.append([
        (d["label"], None),
        (pct_str(d["ryb_q3"]), None), (pct_str(d["ryb_q4"]), None),
        (delta_str(d["ryb_d"]), delta_color(d["ryb_d"])),
        (pct_str(d["tag_q3"]), None), (pct_str(d["tag_q4"]), None),
        (delta_str(d["tag_d"]), delta_color(d["tag_d"])),
    ])
add_delta_table(slide, hdrs7, tbl7_rows, top=4.9, row_height=0.27)
add_slide_footer(slide, note="MSL=Medical Science Liaison, KAM=Key Account Mgr, FRM=Field Reimbursement Mgr, CNE=Clinical Nurse Educator")


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 – RYB vs TAG Prescription Intent
# ══════════════════════════════════════════════════════════════════════════════
data8 = D["slide8_rx_intent"]
labels8 = [d["label"] for d in data8]
ryb_q3_8 = [d["ryb_q3"] or 0 for d in data8]
ryb_q4_8 = [d["ryb_q4"] or 0 for d in data8]
tag_q3_8 = [d["tag_q3"] or 0 for d in data8]
tag_q4_8 = [d["tag_q4"] or 0 for d in data8]

fig, axes = plt.subplots(1, 2, figsize=(CHART_W, 4.0))
for ax, ry, ty, period in [
    (axes[0], ryb_q3_8, tag_q3_8, "Q3 2025"),
    (axes[1], ryb_q4_8, tag_q4_8, "Q4 2025"),
]:
    x = np.arange(len(labels8))
    w = 0.35
    ax.bar(x - w/2, ry, w, label="J&J (Rybrevant)", color=RYB_COLOR, alpha=0.9)
    ax.bar(x + w/2, ty, w, label="AZ (Tagrisso)",   color=TAG_COLOR, alpha=0.9)
    ax.set_xticks(x)
    ax.set_xticklabels([wrap_label(l, 14) for l in labels8], fontsize=7.5)
    ax.set_ylabel("Top 2 Box % (7-pt scale)", fontsize=8)
    ax.set_ylim(50, 95)
    ax.set_title(f"Prescription Intent – {period}", fontsize=10, fontweight="bold")
    ax.grid(axis="y", linestyle="--", alpha=0.35, color="#DDDDDD")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for rect in ax.patches:
        h_val = rect.get_height()
        if h_val > 0:
            ax.text(rect.get_x() + rect.get_width()/2, h_val + 0.3,
                    f"{h_val:.0f}%", ha="center", va="bottom", fontsize=7)
    ax.legend(fontsize=8)
plt.tight_layout()
buf8 = fig_to_buf(fig)

slide = prs.slides.add_slide(blank_layout)
add_slide_header(slide, "Ask #8 – RYB vs TAG Prescription Intent by Patient Type  |  Q3 vs Q4 2025",
                 subtitle="Source: RYB/TAG Sheets (C1_85A-D) | TAG = Monotherapy")
add_image_from_buf(slide, buf8, 0.25, 0.72, CHART_W, 4.1)

hdrs8 = ["Scenario", "J&J Q3", "J&J Q4", "J&J Δ", "AZ Q3", "AZ Q4", "AZ Δ"]
tbl8_rows = []
for d in data8:
    tbl8_rows.append([
        (d["label"].replace("\n", " ")[:35], None),
        (pct_str(d["ryb_q3"]), None), (pct_str(d["ryb_q4"]), None),
        (delta_str(d["ryb_d"]), delta_color(d["ryb_d"])),
        (pct_str(d["tag_q3"]), None), (pct_str(d["tag_q4"]), None),
        (delta_str(d["tag_d"]), delta_color(d["tag_d"])),
    ])
add_delta_table(slide, hdrs8, tbl8_rows, top=5.0, row_height=0.23)
add_slide_footer(slide, note="TAG values = Tagrisso Monotherapy patient groups. T2B = % rating 6-7 on 7-pt scale.")


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 9 – RYB 1L EGFR / Mariposa Discussions
# ══════════════════════════════════════════════════════════════════════════════
data9 = D["slide9_mariposa"]
init9  = data9["initiation"]
mar9   = data9["mariposa_q4"]
top9   = data9["topics"]

fig, axes = plt.subplots(1, 2, figsize=(CHART_W, 4.2))

# Left: Initiation stacked bar (Q3 vs Q4)
ax = axes[0]
q3_phy = init9["physician_initiated"]["q3"] or 0
q4_phy = init9["physician_initiated"]["q4"] or 0
q3_rep = init9["rep_initiated"]["q3"] or 0
q4_rep = init9["rep_initiated"]["q4"] or 0
x_init = np.arange(2)
ax.bar(x_init, [q3_rep, q4_rep], 0.5, label="Rep-initiated", color=RYB_COLOR, alpha=0.9)
ax.bar(x_init, [q3_phy, q4_phy], 0.5, bottom=[q3_rep, q4_rep],
       label="Physician-initiated", color=RYB_LIGHT, alpha=0.9)
ax.set_xticks(x_init)
ax.set_xticklabels(["Q3 2025", "Q4 2025"], fontsize=10)
ax.set_ylabel("%", fontsize=9)
ax.set_ylim(0, 115)
ax.set_title("MARIPOSA Discussion Initiation\n(Who initiated?)", fontsize=10, fontweight="bold")
ax.legend(fontsize=8)
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
# Add labels
for xi, (r_val, p_val) in enumerate([(q3_rep, q3_phy), (q4_rep, q4_phy)]):
    ax.text(xi, r_val/2, f"{r_val:.0f}%", ha="center", va="center",
            fontsize=9, fontweight="bold", color="white")
    ax.text(xi, r_val + p_val/2, f"{p_val:.0f}%", ha="center", va="center",
            fontsize=9, fontweight="bold", color="#333333")

# Right: MARIPOSA discussion quality Top 2 Box (Q4 only)
ax2 = axes[1]
mar_labels = list(mar9.keys())
mar_t2b    = [mar9[k]["top2box"] or 0 for k in mar_labels]
short_labels = ["Disc. 1st\n(First)", "Disc. 1st\n(Later)", "Disc. Most\n(Less)", "Disc. Most\n(Most)"]
bars = ax2.bar(range(len(mar_t2b)), mar_t2b, width=0.5,
               color=[RYB_COLOR, RYB_LIGHT, TAG_LIGHT, TAG_COLOR], alpha=0.9)
ax2.set_xticks(range(len(mar_t2b)))
ax2.set_xticklabels(short_labels, fontsize=9)
ax2.set_ylabel("Top 2 Box %", fontsize=9)
ax2.set_ylim(0, 70)
ax2.set_title("MARIPOSA Discussion Quality\n(Q4 2025 Top 2 Box, N≈200)", fontsize=10, fontweight="bold")
ax2.spines["top"].set_visible(False)
ax2.spines["right"].set_visible(False)
ax2.grid(axis="y", linestyle="--", alpha=0.35)
for bar, val in zip(bars, mar_t2b):
    ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
             f"{val:.0f}%", ha="center", va="bottom", fontsize=9, fontweight="bold")
plt.tight_layout()
buf9 = fig_to_buf(fig)

slide = prs.slides.add_slide(blank_layout)
add_slide_header(slide, "Ask #9 – RYB 1L EGFR / MARIPOSA Discussions  |  Q3 vs Q4 2025",
                 subtitle="Source: RYB Sheet (C1_09AZ) + Additional Analysis | N=103→100")
add_image_from_buf(slide, buf9, 0.25, 0.72, CHART_W, 4.3)

# Delta table: initiation
hdrs9 = ["Who Initiated?", "Q3", "Q4", "Δ (pp)"]
tbl9_rows = [
    [("Rep initiated",      None), (pct_str(q3_rep), None), (pct_str(q4_rep), None),
     (delta_str(round(q4_rep - q3_rep, 1)), delta_color(round(q4_rep - q3_rep, 1)))],
    [("Physician initiated", None), (pct_str(q3_phy), None), (pct_str(q4_phy), None),
     (delta_str(round(q4_phy - q3_phy, 1)), delta_color(round(q4_phy - q3_phy, 1)))],
]
add_delta_table(slide, hdrs9, tbl9_rows, top=5.1, width=6.0, row_height=0.28)

# Topics discussed (select top ones for context)
top_topics = sorted(top9, key=lambda x: x["q4"] or 0, reverse=True)[:6]
hdrs9b = ["Topic Discussed (Q1_50Z)", "Q3", "Q4", "Δ (pp)"]
tbl9b_rows = [[
    (d["topic"][:30], None),
    (pct_str(d["q3"]), None), (pct_str(d["q4"]), None),
    (delta_str(d["delta"]), delta_color(d["delta"])),
] for d in top_topics]
add_delta_table(slide, hdrs9b, tbl9b_rows, left=6.5, top=5.1, width=6.6, row_height=0.28)

add_slide_footer(slide, show_tag=False,
                 note="MARIPOSA = 1L RYB+LAZ setting. Quality scores based on Q4 N≈200 (pooled Q3+Q4 cohort).")


# ══════════════════════════════════════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════════════════════════════════════
prs.save(OUT_PATH)
print(f"\nDeck saved to: {OUT_PATH}")
print(f"  Total slides: {len(prs.slides)}")
