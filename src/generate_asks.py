"""
generate_asks.py
────────────────
Generates the Rybrevant ask-response deck from source_data.xlsx.
Each slide answers a specific ask from projects/jnj_rybrevant/reference/asks.md.

Usage:
    python src/generate_asks.py
"""

import os, sys
import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt, Emu
from lxml import etree
from pptx.oxml.ns import qn

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from slidegen.pptx_utils import (
    SLIDE_W_IN, SLIDE_H_IN,
    C_RYB_Q4, C_RYB_Q3, C_TAG, C_RED, C_GREEN, C_WHITE, C_GREY,
    C_FTGREY, C_LBGREY, C_HDRGREY, C_LTGREY,
    FONT_DISPLAY, FONT_TEXT,
    textbox, solidrect, horiz_line, slide_header, slide_footer,
    section_header_bar, add_delta_col, manual_legend, cover_slide,
    divider_slide, callout_box, hide_axis, set_series_color,
    set_plot_area_gap, set_overlap, set_series_no_border,
    invert_cat_axis, hide_cat_labels, set_datalabel_pos_outside_end,
    set_data_label_color, set_series_line_style, set_series_marker,
    set_series_smooth, _get_or_add, EMU_PER_IN,
)

# ── Paths ──
XLSX = os.path.join(ROOT, "projects", "jnj_rybrevant", "data", "source_data.xlsx")
TMPL = os.path.join(ROOT, "projects", "jnj_rybrevant", "templates", "template.pptx")
OUT  = os.path.join(ROOT, "projects", "jnj_rybrevant", "output", "deck_legacy.pptx")

# ── Extra brand colours ──
C_TAG_Q3     = RGBColor(0xAD, 0x88, 0xC8)   # light purple for Q3
C_TAG_Q4     = RGBColor(0x70, 0x30, 0xA0)   # deep purple for Q4
C_RYB_DARK   = RGBColor(0xF7, 0x58, 0x24)   # deep orange
C_RYB_LIGHT  = RGBColor(0xFF, 0xC1, 0x99)   # pale orange for Q3
C_GREY_BAR   = RGBColor(0xBF, 0xBF, 0xBF)   # neutral grey bars
C_ORANGE_HI  = RGBColor(0xF7, 0x58, 0x24)   # high impact orange
C_GREY_LO    = RGBColor(0xC0, 0xC0, 0xC0)   # non-high-impact grey

# ── Label shortening for rep performance / quality metrics ──
def shorten_label(label, max_len=35):
    """Shorten a metric label to fit chart y-axis."""
    # Keyword-based matching (case-insensitive) — order matters (most specific first)
    SUBS = [
        ("overall quality of sales call", "Overall call quality"),
        ("knowledgeable", "competitor", "Knowledge: competitors"),
        ("knowledgeable", "exon 20", "Knowledge: EGFR Exon 20"),
        ("knowledgeable", "product", "Knowledge: product (RYB)"),
        ("knowledgeable", "egfr", "Knowledge: EGFR mNSCLC"),
        ("knowledgeable", None, "Knowledge: product"),
        ("organized", None, "Organized presentation"),
        ("compelling reason to prescribe", None, "Compelling reason to Rx"),
        ("compelling reason to use broader", None, "Compelling: mol. testing"),
        ("credible support", None, "Credible product claims"),
        ("prepared", "conversation", "Prepared for conversation"),
        ("prepared", None, "Prepared presentation"),
        ("engaging", "discussing", "Engaging discussion"),
        ("meaningful questions", None, "Meaningful questions"),
        ("valuable use of my time", None, "Valuable use of time"),
        ("address my questions", None, "Addressed Rx questions"),
        ("delivering a clear message", None, "Clear product message"),
        ("tailor", "discussions", "Tailored discussion"),
        ("relevant to your practice", None, "Relevant to practice"),
    ]
    ll = label.lower()
    for entry in SUBS:
        if len(entry) == 2:
            kw, short = entry
            if kw.lower() in ll:
                return short
        else:
            kw1, kw2, short = entry
            if kw1.lower() in ll and (kw2 is None or kw2.lower() in ll):
                return short
    # Generic cleanup
    label = label.replace("Johnson & Johnson (Formerly Janssen)", "J&J")
    label = label.replace("Johnson &amp; Johnson", "J&J")
    label = label.replace("[COMPANY]", "J&J")
    label = label.replace("[PRODUCT]", "RYB")
    label = label.replace("How well the ", "").replace("How ", "")
    label = label.strip()
    if len(label) > max_len:
        label = label[:max_len-1] + "\u2026"
    return label

# ── Helpers ──
def pct(v):
    """Convert decimal 0-1 to percentage rounded to 1 decimal."""
    if v is None: return None
    try:
        return round(float(v) * 100, 1)
    except (ValueError, TypeError):
        return None

def sf(v):
    """Straight float — already a percentage, just round."""
    if v is None: return None
    try:
        return round(float(v), 1)
    except (ValueError, TypeError):
        return None

def delta(q4, q3):
    """Q4 minus Q3 in ppts, rounded to 1 decimal."""
    if q4 is None or q3 is None: return None
    return round(q4 - q3, 1)


# ══════════════════════════════════════════════════════════════════════════════
# DATA EXTRACTION
# ══════════════════════════════════════════════════════════════════════════════

def load_data():
    """Load all required data from Excel."""
    data = {}

    # ── RYB sheet ──
    ryb = pd.read_excel(XLSX, sheet_name="RYB", header=None)
    # ── TAG sheet ──
    tag = pd.read_excel(XLSX, sheet_name="TAG", header=None)
    # ── Additional Analysis ──
    aa = pd.read_excel(XLSX, sheet_name="Additonal Analysis", header=None)

    # --- Ask 1 & 2: Message Recall & Effectiveness ---
    # RYB Messages (Q2_10Z) - find the section
    ryb_mr = []
    for i in range(len(ryb)):
        if ryb.iloc[i, 0] == 'Q2_10Z':
            # Messages start at next row
            j = i + 1
            while j < len(ryb) and pd.notna(ryb.iloc[j, 1]):
                code = ryb.iloc[j, 0] if pd.notna(ryb.iloc[j, 0]) else ""
                desc = str(ryb.iloc[j, 1]).strip()
                q3 = ryb.iloc[j, 7]  # col H = index 7
                q4 = ryb.iloc[j, 17]  # col R = index 17
                if desc and not desc.startswith("Base"):
                    ryb_mr.append({
                        "code": code, "desc": desc,
                        "q3": pct(q3) if pd.notna(q3) else None,
                        "q4": pct(q4) if pd.notna(q4) else None,
                    })
                j += 1
                if j - i > 20:
                    break
            break
    data["ryb_mr"] = ryb_mr

    # TAG Messages (Q2_10Z) — use short labels
    TAG_MSG_SHORT = [
        ("CNS metastasis at diag", "CNS mPFS w/ chemo"),
        ("29.4", "mPFS 29.4mo (chemo)"),
        ("mPFS", "mPFS (chemo)"),
        ("89%", "89% stay on treatment"),
        ("Grade 1", "ARs mostly Gr1/2"),
        ("onset frequency and severity", "Onset/severity reduces"),
        ("Longest", "Longest mPFS & mOS"),
        ("47.5", "Final OS 47.5mo"),
        ("NCCN", "NCCN Cat 1 (mono+chemo)"),
        ("30.5", "Median exposure 30.5mo"),
        ("43%", "43% risk reduction (chemo)"),
        ("once-daily", "Once-daily oral dosing"),
        ("convenient", "Once-daily oral dosing"),
    ]
    def tag_short(desc):
        for key, short in TAG_MSG_SHORT:
            if key.lower() in desc.lower():
                return short
        return desc[:35]

    tag_mr = []
    for i in range(len(tag)):
        if tag.iloc[i, 0] == 'Q2_10Z':
            j = i + 1
            while j < len(tag) and pd.notna(tag.iloc[j, 1]):
                code = tag.iloc[j, 0] if pd.notna(tag.iloc[j, 0]) else ""
                desc = str(tag.iloc[j, 1]).strip()
                q3 = tag.iloc[j, 7]  # col H
                q4 = tag.iloc[j, 13]  # col N
                if desc and not desc.startswith("Base"):
                    tag_mr.append({
                        "code": code, "desc": tag_short(desc),
                        "q3": pct(q3) if pd.notna(q3) else None,
                        "q4": pct(q4) if pd.notna(q4) else None,
                    })
                j += 1
                if j - i > 20:
                    break
            break
    data["tag_mr"] = tag_mr

    # RYB Message Effectiveness from AA sheet (rows 30-39, 0-indexed)
    # c1=desc, c2=short tag, c3=MR_Q3, c4=MR_Q4, c5=Believ_Q3, c6=Believ_Q4, c7=ME_Q3, c8=ME_Q4
    ryb_me = []
    for i in range(30, 40):
        if i >= len(aa):
            break
        desc = aa.iloc[i, 1]
        short = aa.iloc[i, 2]
        if not pd.notna(desc) or not isinstance(desc, str) or len(desc) < 10:
            continue
        mr_q3 = aa.iloc[i, 3] if pd.notna(aa.iloc[i, 3]) else None
        mr_q4 = aa.iloc[i, 4] if pd.notna(aa.iloc[i, 4]) else None
        believ_q3 = aa.iloc[i, 5] if pd.notna(aa.iloc[i, 5]) else None
        believ_q4 = aa.iloc[i, 6] if pd.notna(aa.iloc[i, 6]) else None
        me_q3 = aa.iloc[i, 7] if pd.notna(aa.iloc[i, 7]) else None
        me_q4 = aa.iloc[i, 8] if pd.notna(aa.iloc[i, 8]) else None
        short_lbl = str(short).strip() if pd.notna(short) else desc[:30]
        ryb_me.append({
            "label": str(desc).strip(),
            "short": short_lbl,
            "mr_q3": pct(mr_q3), "mr_q4": pct(mr_q4),
            "me_q3": pct(me_q3), "me_q4": pct(me_q4),
            "believ_q3": pct(believ_q3), "believ_q4": pct(believ_q4),
            "me_comp_q3": pct(me_q3), "me_comp_q4": pct(me_q4),
        })
    data["ryb_me"] = ryb_me

    # --- Ask 3: Rep Performance (from AA sheet rows 5-20) ---
    rep_perf = []
    for i in range(4, min(25, len(aa))):
        metric = aa.iloc[i, 1]
        if pd.notna(metric) and isinstance(metric, str) and len(metric) > 5:
            ryb_q3 = aa.iloc[i, 2]
            ryb_q4 = aa.iloc[i, 3]
            tag_q3 = aa.iloc[i, 4]
            tag_q4 = aa.iloc[i, 5]
            if pd.notna(ryb_q4) and pd.notna(tag_q4):
                rep_perf.append({
                    "metric": shorten_label(metric.strip()),
                    "ryb_q3": pct(ryb_q3), "ryb_q4": pct(ryb_q4),
                    "tag_q3": pct(tag_q3), "tag_q4": pct(tag_q4),
                })
    data["rep_perf"] = rep_perf

    # --- Ask 4: M/B/D from AA sheet (rows 30-39) ---
    # Same data as ryb_me but we use believability + composite columns
    mbd = []
    for m in ryb_me:
        mbd.append({
            "label": m["label"][:50],
            "short": m["short"],
            "believ_q3": m.get("believ_q3"),
            "believ_q4": m.get("believ_q4"),
            "composite_q3": m.get("me_comp_q3"),
            "composite_q4": m.get("me_comp_q4"),
        })
    data["mbd"] = mbd

    # --- Ask 5: CTA (Call to Action) ---
    # RYB CTA: C1_81Z (compelling), C1_82Z (changed opinion), C1_83D1Z (closing), Q1_84bZ (education)
    ryb_cta = []
    cta_codes = ['C1_81Z', 'C1_82Z', 'C1_83D1Z', 'Q1_84bZ']
    cta_labels = ['Compelling reason to prescribe', 'Changed opinion',
                  'Asked to prescribe (Branded closing)', 'Educational ask']
    for code, label in zip(cta_codes, cta_labels):
        for i in range(len(ryb)):
            if ryb.iloc[i, 0] == code:
                # Top box is typically the row after the code row or the "Top 2 Box" row
                for j in range(i, min(i+5, len(ryb))):
                    desc = str(ryb.iloc[j, 1]) if pd.notna(ryb.iloc[j, 1]) else ""
                    if "top" in desc.lower() or "yes" in desc.lower() or j == i + 1:
                        q3 = ryb.iloc[j, 7]
                        q4 = ryb.iloc[j, 17]
                        if pd.notna(q3) and pd.notna(q4):
                            ryb_cta.append({
                                "label": label, "code": code,
                                "q3": pct(q3), "q4": pct(q4),
                            })
                            break
                break
    data["ryb_cta"] = ryb_cta

    # TAG CTA: equivalent codes
    tag_cta = []
    tag_cta_codes_labels = [
        ('C1_81_TAG_mNSCLC_Z', 'Compelling reason to prescribe'),
        ('C1_82_TAG_mNSCLC_Z', 'Changed opinion'),
        ('C1_83D_TAG_mNSCLC_Z', 'Asked to prescribe (Branded closing)'),
        ('C1_84b_TAG_mNSCLC_Z', 'Educational ask'),
    ]
    for code, label in tag_cta_codes_labels:
        for i in range(len(tag)):
            if tag.iloc[i, 0] == code:
                for j in range(i, min(i+8, len(tag))):
                    desc = str(tag.iloc[j, 1]) if pd.notna(tag.iloc[j, 1]) else ""
                    if "top" in desc.lower() or "yes" in desc.lower() or j == i + 1:
                        q3 = tag.iloc[j, 7]
                        q4 = tag.iloc[j, 13]
                        if pd.notna(q3) and pd.notna(q4):
                            tag_cta.append({
                                "label": label, "code": code,
                                "q3": pct(q3), "q4": pct(q4),
                            })
                            break
                break
    data["tag_cta"] = tag_cta

    # --- Ask 7: Follow-up reps (C1_84FZ) ---
    FU_SHORT = {
        "Medical Science Liaison": "MSL",
        "Key Account Manager": "KAM",
        "Field Reimbursement": "FRM",
        "Clinical Nurse Educator": "CNE",
        "No, I did not": "No follow-up set",
    }
    def fu_short(desc):
        for key, short in FU_SHORT.items():
            if key.lower() in desc.lower():
                return short
        return desc[:30]

    ryb_followup = []
    for i in range(len(ryb)):
        if ryb.iloc[i, 0] == 'C1_84FZ':
            j = i + 1
            while j < len(ryb) and pd.notna(ryb.iloc[j, 1]):
                desc = str(ryb.iloc[j, 1]).strip()
                q3 = ryb.iloc[j, 7]
                q4 = ryb.iloc[j, 17]
                if desc and not desc.startswith("Base") and pd.notna(q4):
                    ryb_followup.append({
                        "desc": fu_short(desc),
                        "q3": pct(q3) if pd.notna(q3) else None,
                        "q4": pct(q4) if pd.notna(q4) else None,
                    })
                j += 1
                if j - i > 12:
                    break
            break
    data["ryb_followup"] = ryb_followup

    # --- Ask 8: Prescription intent (C1_85DZ for RYB, C1_85D_TAG2_mNSCLCZ for TAG) ---
    RX_SHORT = {"without CNS": "Without CNS metastasis", "with CNS": "With CNS metastasis"}
    def rx_short(desc):
        for key, short in RX_SHORT.items():
            if key.lower() in desc.lower():
                return short
        return desc[:30]

    ryb_rx = []
    for i in range(len(ryb)):
        if ryb.iloc[i, 0] == 'C1_85DZ':
            j = i + 1
            while j < len(ryb) and pd.notna(ryb.iloc[j, 1]):
                desc = str(ryb.iloc[j, 1]).strip()
                q3 = ryb.iloc[j, 7]
                q4 = ryb.iloc[j, 17]
                if desc and not desc.startswith("Base") and pd.notna(q4):
                    ryb_rx.append({
                        "desc": rx_short(desc),
                        "q3": pct(q3) if pd.notna(q3) else None,
                        "q4": pct(q4) if pd.notna(q4) else None,
                    })
                j += 1
                if j - i > 8:
                    break
            break
    data["ryb_rx"] = ryb_rx

    tag_rx = []
    for i in range(len(tag)):
        if tag.iloc[i, 0] == 'C1_85D_TAG2_mNSCLCZ':
            j = i + 1
            while j < len(tag) and pd.notna(tag.iloc[j, 1]):
                desc = str(tag.iloc[j, 1]).strip()
                q3 = tag.iloc[j, 7]
                q4 = tag.iloc[j, 13]
                if desc and not desc.startswith("Base") and pd.notna(q4):
                    tag_rx.append({
                        "desc": rx_short(desc),
                        "q3": pct(q3) if pd.notna(q3) else None,
                        "q4": pct(q4) if pd.notna(q4) else None,
                    })
                j += 1
                if j - i > 8:
                    break
            break
    data["tag_rx"] = tag_rx

    # --- Ask 9: Share of Time / Order of Settings (from AA rows 197-204) ---
    # c3=topic label, c4=Q3%, c7=Q4% (whole-number percentages)
    share_of_time = []
    for i in range(197, 205):
        if i >= len(aa):
            break
        label = aa.iloc[i, 3]  # col 3 = topic name
        if pd.notna(label) and isinstance(label, str) and len(label) > 3:
            q3 = aa.iloc[i, 4] if pd.notna(aa.iloc[i, 4]) else None
            q4 = aa.iloc[i, 7] if pd.notna(aa.iloc[i, 7]) else None
            share_of_time.append({
                "label": str(label).strip(),
                "q3": sf(q3) if q3 is not None else None,
                "q4": sf(q4) if q4 is not None else None,
            })
    data["share_of_time"] = share_of_time

    # --- Nebulous 1: Quality (Q1_87Z from RYB) + Closing ---
    ryb_quality = []
    for i in range(len(ryb)):
        if ryb.iloc[i, 0] == 'Q1_87Z':
            j = i + 1
            while j < len(ryb) and pd.notna(ryb.iloc[j, 1]):
                desc = str(ryb.iloc[j, 1]).strip()
                q3 = ryb.iloc[j, 7]
                q4 = ryb.iloc[j, 17]
                if desc and not desc.startswith("Base") and pd.notna(q4):
                    ryb_quality.append({
                        "desc": shorten_label(desc),
                        "q3": pct(q3) if pd.notna(q3) else None,
                        "q4": pct(q4) if pd.notna(q4) else None,
                    })
                j += 1
                if j - i > 20:
                    break
            break
    data["ryb_quality"] = ryb_quality

    # --- Nebulous 2: High Impact vs Non-High Impact (from RYB sheet) ---
    # col 9=Others Q3, col 10=HI Q3, col 19=Others Q4, col 20=HI Q4
    hii_compare = []
    for i in range(len(ryb)):
        if ryb.iloc[i, 0] == 'Q1_87Z':
            j = i + 1
            while j < len(ryb) and pd.notna(ryb.iloc[j, 1]):
                desc = str(ryb.iloc[j, 1]).strip()
                hi_q4 = ryb.iloc[j, 20]   # High Impact Q4
                oth_q4 = ryb.iloc[j, 19]  # Others Q4
                if desc and not desc.startswith("Base") and pd.notna(hi_q4) and pd.notna(oth_q4):
                    hi_pct = pct(hi_q4)
                    oth_pct = pct(oth_q4)
                    if hi_pct is not None and oth_pct is not None:
                        hii_compare.append({
                            "desc": shorten_label(desc),
                            "hi_q4": hi_pct,
                            "other_q4": oth_pct,
                            "diff": round(hi_pct - oth_pct, 1),
                        })
                j += 1
                if j - i > 20:
                    break
            break
    # Filter for sizeable differences (>5pp)
    hii_compare = [h for h in hii_compare if abs(h["diff"]) > 3]
    hii_compare.sort(key=lambda x: abs(x["diff"]), reverse=True)
    data["hii_compare"] = hii_compare

    # --- HII percentages (AA row ~48) ---
    hii_pcts = {}
    for i in range(43, min(55, len(aa))):
        label = aa.iloc[i, 1]
        if pd.notna(label) and isinstance(label, str) and "high impact" in label.lower():
            hii_pcts["ryb_q3"] = sf(aa.iloc[i, 2]) if pd.notna(aa.iloc[i, 2]) else None
            hii_pcts["ryb_q4"] = sf(aa.iloc[i, 3]) if pd.notna(aa.iloc[i, 3]) else None
            hii_pcts["tag_q3"] = sf(aa.iloc[i, 4]) if pd.notna(aa.iloc[i, 4]) else None
            hii_pcts["tag_q4"] = sf(aa.iloc[i, 5]) if pd.notna(aa.iloc[i, 5]) else None
            break
    data["hii_pcts"] = hii_pcts

    # --- TAG Q2_20Z: Order of message recall (rows 250-302) ---
    # Each message has 4 sub-rows: 1st, 2nd, 3rd, 4th recalled
    # Group by message code to get recall-order breakdown
    tag_recall_order = {}  # code -> {desc, 1st_q4, 2nd_q4, 3rd_q4, 4th_q4, total_q4, ...}
    for i in range(250, min(303, len(tag))):
        code = str(tag.iloc[i, 0]).strip() if pd.notna(tag.iloc[i, 0]) else ""
        desc = str(tag.iloc[i, 1]).strip() if pd.notna(tag.iloc[i, 1]) else ""
        ordinal = str(tag.iloc[i, 2]).strip() if pd.notna(tag.iloc[i, 2]) else ""
        q4 = tag.iloc[i, 13]
        q3 = tag.iloc[i, 7]
        if not code or not desc:
            continue
        if code not in tag_recall_order:
            tag_recall_order[code] = {
                "code": code,
                "desc": tag_short(desc),
                "1st_q4": 0, "2nd_q4": 0, "3rd_q4": 0, "4th_q4": 0,
                "1st_q3": 0, "2nd_q3": 0, "3rd_q3": 0, "4th_q3": 0,
            }
        entry = tag_recall_order[code]
        q4v = pct(q4) if pd.notna(q4) else 0
        q3v = pct(q3) if pd.notna(q3) else 0
        if ordinal.startswith("1"):
            entry["1st_q4"] = q4v; entry["1st_q3"] = q3v
        elif ordinal.startswith("2"):
            entry["2nd_q4"] = q4v; entry["2nd_q3"] = q3v
        elif ordinal.startswith("3"):
            entry["3rd_q4"] = q4v; entry["3rd_q3"] = q3v
        elif ordinal.startswith("4"):
            entry["4th_q4"] = q4v; entry["4th_q3"] = q3v
    # Calculate totals and build list
    tag_order = []
    for code, entry in tag_recall_order.items():
        entry["total_q4"] = round(entry["1st_q4"] + entry["2nd_q4"] + entry["3rd_q4"] + entry["4th_q4"], 1)
        entry["total_q3"] = round(entry["1st_q3"] + entry["2nd_q3"] + entry["3rd_q3"] + entry["4th_q3"], 1)
        if entry["total_q4"] > 0:
            tag_order.append(entry)
    # Sort by overall recall descending
    tag_order.sort(key=lambda x: x["total_q4"], reverse=True)
    data["tag_order"] = tag_order

    # --- Nebulous Ask 4: Mariposa-segmented rep performance (AA rows 134-151) ---
    # c3=overall Q4, c4=overlapped (Mariposa discussed), c5=non-overlapped
    mariposa_rep = []
    for i in range(136, min(151, len(aa))):
        metric_desc = aa.iloc[i, 2]
        if pd.notna(metric_desc) and isinstance(metric_desc, str) and len(metric_desc) > 10:
            overall = aa.iloc[i, 3]
            overlap = aa.iloc[i, 4]
            non_overlap = aa.iloc[i, 5]
            if pd.notna(overall) and pd.notna(overlap) and pd.notna(non_overlap):
                mariposa_rep.append({
                    "metric": shorten_label(metric_desc.strip()),
                    "overall": pct(overall),
                    "mariposa_1st": pct(overlap),
                    "mariposa_not": pct(non_overlap),
                })
    data["mariposa_rep"] = mariposa_rep

    # Mariposa HII & quality (AA rows 59-63, 173)
    mariposa_hii = {}
    if 173 < len(aa):
        # Row 173: HII by segment (c3=Less, c4=Most, c5=1st, c6=Later)
        mariposa_hii["hii_less"] = sf(aa.iloc[173, 3]) if pd.notna(aa.iloc[173, 3]) else None
        mariposa_hii["hii_most"] = sf(aa.iloc[173, 4]) if pd.notna(aa.iloc[173, 4]) else None
        mariposa_hii["hii_first"] = sf(aa.iloc[173, 5]) if pd.notna(aa.iloc[173, 5]) else None
        mariposa_hii["hii_later"] = sf(aa.iloc[173, 6]) if pd.notna(aa.iloc[173, 6]) else None
    data["mariposa_hii"] = mariposa_hii

    # --- Sample sizes ---
    data["ryb_n_q3"] = 103
    data["ryb_n_q4"] = 100
    data["tag_n_q3"] = 73
    data["tag_n_q4"] = 73

    return data


# ══════════════════════════════════════════════════════════════════════════════
# CHART HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _enable_data_labels(series, color, fsize=8, num_fmt='0"%"', pos="outEnd"):
    """Enable and configure data labels on a chart series.
    pos: 'outEnd' for regular bars, 'ctr' for stacked bars."""
    plot = series._element.getparent()
    # Enable at plot level
    plot_dLbls = plot.find(qn("c:dLbls"))
    if plot_dLbls is None:
        plot_dLbls = etree.SubElement(plot, qn("c:dLbls"))
    _get_or_add(plot_dLbls, "c:showVal").set("val", "1")
    _get_or_add(plot_dLbls, "c:showCatName").set("val", "0")
    _get_or_add(plot_dLbls, "c:showSerName").set("val", "0")
    _get_or_add(plot_dLbls, "c:showPercent").set("val", "0")

    # Enable at series level
    dLbls = _get_or_add(series._element, "c:dLbls")
    _get_or_add(dLbls, "c:showVal").set("val", "1")
    _get_or_add(dLbls, "c:showCatName").set("val", "0")
    _get_or_add(dLbls, "c:showSerName").set("val", "0")
    _get_or_add(dLbls, "c:showPercent").set("val", "0")
    numFmt = _get_or_add(dLbls, "c:numFmt")
    numFmt.set("formatCode", num_fmt)
    numFmt.set("sourceLinked", "0")
    # Set label position
    dLblPos = _get_or_add(dLbls, "c:dLblPos")
    dLblPos.set("val", pos)
    set_data_label_color(series, color)

    # Font
    txPr = _get_or_add(dLbls, "c:txPr")
    _get_or_add(txPr, "a:bodyPr")
    _get_or_add(txPr, "a:lstStyle")
    p = _get_or_add(txPr, "a:p")
    pPr = _get_or_add(p, "a:pPr")
    defRPr = _get_or_add(pPr, "a:defRPr")
    defRPr.set("sz", str(int(fsize * 100)))
    defRPr.set("b", "1")
    sf = _get_or_add(defRPr, "a:solidFill")
    clr = _get_or_add(sf, "a:srgbClr")
    clr.set("val", str(color))
    latin = _get_or_add(defRPr, "a:latin")
    latin.set("typeface", FONT_TEXT)


def add_bar_chart(slide, categories, values, left, top, width, height,
                  fill_color=C_RYB_Q4, chart_title=None, max_val=100):
    """Add a horizontal bar chart showing Q4 values."""
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series("Q4'25", values)

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED, Inches(left), Inches(top),
        Inches(width), Inches(height), chart_data)
    chart = chart_frame.chart
    chart.has_legend = False

    # Style
    series = chart.series[0]
    set_series_color(series, fill_color)
    set_series_no_border(series)
    set_plot_area_gap(chart, 80)

    # Data labels
    _enable_data_labels(series, fill_color)

    # Axes
    hide_axis(chart, "val")
    chart.category_axis.has_major_gridlines = False
    chart.category_axis.tick_labels.font.size = Pt(7)
    chart.category_axis.tick_labels.font.name = FONT_TEXT
    invert_cat_axis(chart)

    if chart_title:
        chart.has_title = True
        chart.chart_title.text_frame.paragraphs[0].text = chart_title
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(9)
        chart.chart_title.text_frame.paragraphs[0].font.bold = True
        chart.chart_title.text_frame.paragraphs[0].font.name = FONT_DISPLAY

    return chart_frame, chart


def add_clustered_bar_chart(slide, categories, series_data, left, top, width, height,
                            colors=None, legend=True, max_val=100):
    """Add a clustered horizontal bar chart with multiple series."""
    chart_data = CategoryChartData()
    chart_data.categories = categories
    for name, vals in series_data:
        chart_data.add_series(name, vals)

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED, Inches(left), Inches(top),
        Inches(width), Inches(height), chart_data)
    chart = chart_frame.chart

    if colors is None:
        colors = [C_RYB_Q4, C_TAG_Q4]

    for idx, series in enumerate(chart.series):
        c = colors[idx] if idx < len(colors) else C_GREY_BAR
        set_series_color(series, c)
        set_series_no_border(series)
        _enable_data_labels(series, c, fsize=7)

    set_plot_area_gap(chart, 100)
    set_overlap(chart, 0)
    hide_axis(chart, "val")
    chart.category_axis.has_major_gridlines = False
    chart.category_axis.tick_labels.font.size = Pt(7)
    chart.category_axis.tick_labels.font.name = FONT_TEXT
    invert_cat_axis(chart)

    if legend:
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(8)
        chart.legend.font.name = FONT_TEXT
    else:
        chart.has_legend = False

    return chart_frame, chart


def add_delta_table(slide, labels, deltas, left, top, width, row_height, header_text="QoQ Delta"):
    """Add a delta column table with green/red coloring."""
    n = len(deltas)
    total_h = 0.28 + n * row_height  # header + data rows

    tbl_shape = slide.shapes.add_table(n + 1, 1,
        Inches(left), Inches(top), Inches(width), Inches(total_h))
    tbl = tbl_shape.table

    # Header
    tbl.rows[0].height = Inches(0.28)
    hc = tbl.cell(0, 0)
    hc.fill.solid()
    hc.fill.fore_color.rgb = C_HDRGREY
    p = hc.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = header_text
    run.font.size = Pt(7)
    run.font.bold = True
    run.font.color.rgb = C_WHITE
    run.font.name = FONT_TEXT

    # Data rows
    for i, d in enumerate(deltas):
        tbl.rows[i + 1].height = Inches(row_height)
        cell = tbl.cell(i + 1, 0)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_LBGREY if i % 2 == 0 else C_WHITE

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
        run.font.name = FONT_TEXT

    tbl.columns[0].width = Inches(width)
    return tbl_shape


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE BUILDERS
# ══════════════════════════════════════════════════════════════════════════════

def build_cover(prs):
    """Slide 1: Title/Cover."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
    cover_slide(slide,
                "RYB+LAZ Promotional Effectiveness\nTracking Q4'25 Report",
                "Answers to key strategic asks",
                "March 2026",
                "Johnson & Johnson")


def build_es(prs, data):
    """Slide 2: Executive Summary."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
    slide_header(slide, "Executive Summary — Key Insights from Q4'25")

    insights = [
        "Messaging: RYB message recall declined across most messages in Q4, with new NCCN messages (Cat 1: 49%, CNS preferred: 31%) gaining traction. Avg MR dropped from ~44% to 36.5%, though ME remained strong at 72.5%.",
        "Tagrisso messaging: TAG also saw MR declines, with two new messages (43% risk reduction, once-daily dosing) entering at 38-40%. OS headline recall dropped from 55% to 42%.",
        "Rep Performance: J&J reps outperform AZ on all 15 rep quality metrics. Overall quality at 82% (vs TAG 73%), with knowledge and organization at 80-81%.",
        "Call to Action: RYB reps maintain strong CTA execution — changed opinion rose to 57% (+8pp). However, branded closing dipped slightly to 46% (-4pp).",
        "Prescription Intent: RYB intent to increase Rx improved for patients without CNS (59% to 69%, +10pp). TAG chemo intent held steady.",
        "One J&J Vision: 44% of RYB HCPs did NOT set follow-up with other J&J reps, a slight decline from Q3 (49% had no follow-up).",
        "High Impact Interactions: RYB HII rose to 63% (+6pp QoQ), outpacing TAG mono (57%) and TAG+chemo (58%).",
    ]

    y = 1.50
    for i, ins in enumerate(insights):
        # Bullet marker
        solidrect(slide, 0.40, y + 0.06, 0.08, 0.08, C_RYB_Q4)
        textbox(slide, ins, 0.60, y, 12.40, 0.50,
                fsize=9, color=C_GREY, font=FONT_TEXT)
        y += 0.62

    slide_footer(slide, "Source: Lung SFEA SB — Q3 & Q4 2025 | RYB N=103/100, TAG N=73/73")


def build_ryb_mr_me(prs, data):
    """Slide 3 — Ask 1: RYB Message Recall & Effectiveness."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide,
        "RYB message recall declined broadly in Q4; new NCCN messages gained traction while effectiveness remained strong")

    section_header_bar(slide,
        "RYB+LAZ MESSAGE RECALL & EFFECTIVENESS — Q4'25 vs Q3'25", top=1.40)

    msgs = data["ryb_me"] if data["ryb_me"] else data["ryb_mr"]
    # Sort by MR Q4 descending
    msgs_sorted = sorted(msgs, key=lambda x: (x.get("mr_q4") or x.get("q4") or 0), reverse=True)

    labels = [m.get("short", m.get("desc", ""))[:35] for m in msgs_sorted]
    mr_q4 = [m.get("mr_q4") or m.get("q4") or 0 for m in msgs_sorted]
    mr_q3 = [m.get("mr_q3") or m.get("q3") or 0 for m in msgs_sorted]
    me_q4 = [m.get("me_q4") or m.get("me_comp_q4") or 0 for m in msgs_sorted]
    me_q3 = [m.get("me_q3") or m.get("me_comp_q3") or 0 for m in msgs_sorted]

    n = len(labels)
    chart_top = 1.80
    chart_h = min(4.2, n * 0.42)
    row_h = chart_h / max(n, 1)

    # Left chart: Message Recall
    textbox(slide, "Message Recall (%)", 0.30, 1.72, 3.0, 0.25,
            fsize=9, bold=True, color=C_GREY, font=FONT_DISPLAY)

    chart_data = CategoryChartData()
    chart_data.categories = labels
    chart_data.add_series("Q4'25", mr_q4)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.30), Inches(chart_top), Inches(5.0), Inches(chart_h), chart_data)
    ch = cf.chart
    ch.has_legend = False
    s = ch.series[0]
    set_series_color(s, C_RYB_Q4)
    set_series_no_border(s)
    _enable_data_labels(s, C_RYB_Q4)
    hide_axis(ch, "val")
    ch.category_axis.tick_labels.font.size = Pt(7)
    ch.category_axis.tick_labels.font.name = FONT_TEXT
    invert_cat_axis(ch)
    set_plot_area_gap(ch, 80)

    # MR Delta table
    mr_deltas = [delta(q4, q3) for q4, q3 in zip(mr_q4, mr_q3)]
    add_delta_table(slide, labels, mr_deltas, 5.35, chart_top, 0.65, row_h * 0.85,
                    header_text="MR \u0394")

    # Right chart: Message Effectiveness
    textbox(slide, "Message Effectiveness (%)", 6.20, 1.72, 3.5, 0.25,
            fsize=9, bold=True, color=C_GREY, font=FONT_DISPLAY)

    chart_data2 = CategoryChartData()
    chart_data2.categories = labels
    chart_data2.add_series("Q4'25", me_q4)

    cf2 = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(6.20), Inches(chart_top), Inches(5.0), Inches(chart_h), chart_data2)
    ch2 = cf2.chart
    ch2.has_legend = False
    s2 = ch2.series[0]
    set_series_color(s2, C_RYB_Q4)
    set_series_no_border(s2)
    _enable_data_labels(s2, C_RYB_Q4)
    hide_axis(ch2, "val")
    hide_cat_labels(ch2)
    invert_cat_axis(ch2)
    set_plot_area_gap(ch2, 80)

    # ME Delta table
    me_deltas = [delta(q4, q3) for q4, q3 in zip(me_q4, me_q3)]
    add_delta_table(slide, labels, me_deltas, 11.25, chart_top, 0.65, row_h * 0.85,
                    header_text="ME \u0394")

    # Legend
    ly = chart_top + chart_h + 0.15
    solidrect(slide, 2.0, ly, 0.18, 0.12, C_RYB_Q4)
    textbox(slide, f"Q4'25 (n={data['ryb_n_q4']})", 2.25, ly - 0.02, 1.5, 0.18,
            fsize=7, color=C_GREY)
    solidrect(slide, 4.0, ly, 0.18, 0.12, C_GREEN)
    textbox(slide, "Positive \u0394", 4.25, ly - 0.02, 0.8, 0.18, fsize=7, color=C_GREY)
    solidrect(slide, 5.2, ly, 0.18, 0.12, C_RED)
    textbox(slide, "Negative \u0394", 5.45, ly - 0.02, 0.8, 0.18, fsize=7, color=C_GREY)
    textbox(slide, "N/A = New message (no Q3 baseline)", 6.5, ly - 0.02, 3.0, 0.18,
            fsize=7, color=C_FTGREY)

    slide_footer(slide, "Source: Lung SFEA SB Q2_10Z | MR = Message Recall, ME = Message Effectiveness | RYB+LAZ Q3 N=103, Q4 N=100")


def build_tag_mr_me(prs, data):
    """Slide 4 — Ask 2: TAG Message Recall & Effectiveness."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide,
        "Tagrisso message recall declined in Q4; new messages on risk reduction and daily dosing entered mid-range")

    section_header_bar(slide,
        "TAGRISSO MESSAGE RECALL — Q4'25 vs Q3'25", top=1.40)

    msgs = data["tag_mr"]
    msgs_sorted = sorted(msgs, key=lambda x: (x.get("q4") or 0), reverse=True)

    labels = [m["desc"] for m in msgs_sorted]
    mr_q4 = [m["q4"] or 0 for m in msgs_sorted]
    mr_q3 = [m["q3"] or 0 for m in msgs_sorted]

    n = len(labels)
    chart_top = 1.80
    chart_h = min(4.5, n * 0.40)
    row_h = chart_h / max(n, 1)

    # Message Recall chart — full width
    textbox(slide, "Message Recall (%)", 0.30, 1.72, 3.0, 0.25,
            fsize=9, bold=True, color=C_GREY, font=FONT_DISPLAY)

    chart_data = CategoryChartData()
    chart_data.categories = labels
    chart_data.add_series("Q4'25", mr_q4)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.30), Inches(chart_top), Inches(9.0), Inches(chart_h), chart_data)
    ch = cf.chart
    ch.has_legend = False
    s = ch.series[0]
    set_series_color(s, C_TAG_Q4)
    set_series_no_border(s)
    _enable_data_labels(s, C_TAG_Q4)
    hide_axis(ch, "val")
    ch.category_axis.tick_labels.font.size = Pt(7.5)
    ch.category_axis.tick_labels.font.name = FONT_TEXT
    invert_cat_axis(ch)
    set_plot_area_gap(ch, 80)

    # Delta table
    mr_deltas = [delta(q4, q3) for q4, q3 in zip(mr_q4, mr_q3)]
    add_delta_table(slide, labels, mr_deltas, 9.40, chart_top, 0.65, row_h * 0.85,
                    header_text="MR \u0394")

    # Callout
    callout_box(slide, 10.30, chart_top, 2.80, 2.5,
        text="Two new Q4 messages entered:\n"
             "\u2022 43% risk reduction: 38%\n"
             "\u2022 Once-daily oral dosing: 40%\n\n"
             "Largest MR declines:\n"
             "\u2022 OS headline: -13pp\n"
             "\u2022 mPFS: -10pp\n"
             "\u2022 Longest mPFS/mOS: -9pp",
        border_color=C_TAG_Q4, fsize=7.5)

    # Legend
    ly = chart_top + chart_h + 0.15
    solidrect(slide, 2.0, ly, 0.18, 0.12, C_TAG_Q4)
    textbox(slide, f"Q4'25 (n={data['tag_n_q4']})", 2.25, ly - 0.02, 1.5, 0.18,
            fsize=7, color=C_GREY)
    solidrect(slide, 4.0, ly, 0.18, 0.12, C_GREEN)
    textbox(slide, "Positive \u0394", 4.25, ly - 0.02, 0.8, 0.18, fsize=7, color=C_GREY)
    solidrect(slide, 5.2, ly, 0.18, 0.12, C_RED)
    textbox(slide, "Negative \u0394", 5.45, ly - 0.02, 0.8, 0.18, fsize=7, color=C_GREY)
    textbox(slide, "N/A = New message (no Q3 baseline)", 6.5, ly - 0.02, 3.0, 0.18,
            fsize=7, color=C_FTGREY)

    slide_footer(slide, "Source: Lung SFEA SB Q2_10Z | TAG Q3 N=73, Q4 N=73")


def build_rep_perf(prs, data):
    """Slide 5 — Ask 3: Rep Performance Abacus (J&J vs AZ)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide,
        "J&J reps outperform AZ across all 15 rep quality metrics; overall quality at 82% vs 73%")

    section_header_bar(slide,
        "REP PERFORMANCE — J&J vs AZ — Q4'25", top=1.40)

    rp = data["rep_perf"]
    # Sort by RYB Q4 descending
    rp_sorted = sorted(rp, key=lambda x: x["ryb_q4"], reverse=True)

    labels = [r["metric"][:45] for r in rp_sorted]
    ryb_vals = [r["ryb_q4"] for r in rp_sorted]
    tag_vals = [r["tag_q4"] for r in rp_sorted]
    ryb_deltas = [delta(r["ryb_q4"], r["ryb_q3"]) for r in rp_sorted]
    tag_deltas = [delta(r["tag_q4"], r["tag_q3"]) for r in rp_sorted]

    n = len(labels)
    chart_top = 1.80
    chart_h = min(4.8, n * 0.32)
    row_h = chart_h / max(n, 1)

    # Abacus chart (clustered horizontal bar)
    chart_data = CategoryChartData()
    chart_data.categories = labels
    chart_data.add_series("J&J (RYB+LAZ)", ryb_vals)
    chart_data.add_series("AZ (Tagrisso)", tag_vals)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.30), Inches(chart_top), Inches(8.5), Inches(chart_h), chart_data)
    ch = cf.chart
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.font.size = Pt(8)
    ch.legend.font.name = FONT_TEXT

    s0 = ch.series[0]
    set_series_color(s0, C_RYB_Q4)
    set_series_no_border(s0)
    _enable_data_labels(s0, C_RYB_Q4)

    s1 = ch.series[1]
    set_series_color(s1, C_TAG_Q4)
    set_series_no_border(s1)
    _enable_data_labels(s1, C_TAG_Q4)

    hide_axis(ch, "val")
    ch.category_axis.tick_labels.font.size = Pt(6.5)
    ch.category_axis.tick_labels.font.name = FONT_TEXT
    invert_cat_axis(ch)
    set_plot_area_gap(ch, 65)
    set_overlap(ch, -15)

    # Delta tables
    textbox(slide, "J&J \u0394", 9.10, chart_top - 0.15, 0.65, 0.18,
            fsize=7, bold=True, color=C_RYB_Q4, align=PP_ALIGN.CENTER)
    add_delta_table(slide, labels, ryb_deltas, 9.10, chart_top, 0.65, row_h * 0.85,
                    header_text="J&J \u0394")

    textbox(slide, "AZ \u0394", 9.85, chart_top - 0.15, 0.65, 0.18,
            fsize=7, bold=True, color=C_TAG_Q4, align=PP_ALIGN.CENTER)
    add_delta_table(slide, labels, tag_deltas, 9.85, chart_top, 0.65, row_h * 0.85,
                    header_text="AZ \u0394")

    slide_footer(slide, "Source: Lung SFEA SB — Sales_Rep_Effectiveness_PET_RF_2 | Top 2 Box (6,7) | RYB N=100, TAG N=73")


def build_mbd(prs, data):
    """Slide 6 — Ask 4: Message M/B/D Components."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide,
        "Message effectiveness driven primarily by believability; composite effectiveness remained stable QoQ")

    section_header_bar(slide,
        "MESSAGE BELIEVABILITY & COMPOSITE EFFECTIVENESS — RYB+LAZ — Q4'25 vs Q3'25", top=1.40)

    msgs = data["mbd"]
    if not msgs:
        textbox(slide, "M/B/D data not available in current extract", 2, 3, 8, 1,
                fsize=14, color=C_RED)
        return

    # Sort by composite Q4 descending (as proxy for Motivation)
    msgs_sorted = sorted(msgs, key=lambda x: x.get("composite_q4") or 0, reverse=True)

    labels = [m["short"][:35] for m in msgs_sorted]
    n = len(labels)
    chart_top = 1.80
    chart_h = min(4.2, n * 0.42)
    row_h = chart_h / max(n, 1)

    believ_q4 = [m.get("believ_q4") or 0 for m in msgs_sorted]
    believ_q3 = [m.get("believ_q3") or 0 for m in msgs_sorted]
    comp_q4 = [m.get("composite_q4") or 0 for m in msgs_sorted]
    comp_q3 = [m.get("composite_q3") or 0 for m in msgs_sorted]

    # Left: Believability
    textbox(slide, "Believability (%)", 0.30, 1.72, 3.0, 0.25,
            fsize=9, bold=True, color=C_GREY, font=FONT_DISPLAY)

    cd1 = CategoryChartData()
    cd1.categories = labels
    cd1.add_series("Q4'25", believ_q4)
    cd1.add_series("Q3'25", believ_q3)

    cf1 = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.30), Inches(chart_top), Inches(4.8), Inches(chart_h), cd1)
    ch1 = cf1.chart
    ch1.has_legend = False
    set_series_color(ch1.series[0], C_RYB_Q4)
    set_series_no_border(ch1.series[0])
    set_series_color(ch1.series[1], C_RYB_LIGHT)
    set_series_no_border(ch1.series[1])
    _enable_data_labels(ch1.series[0], C_RYB_Q4, fsize=7)
    _enable_data_labels(ch1.series[1], C_RYB_LIGHT, fsize=7)
    hide_axis(ch1, "val")
    ch1.category_axis.tick_labels.font.size = Pt(7)
    ch1.category_axis.tick_labels.font.name = FONT_TEXT
    invert_cat_axis(ch1)
    set_plot_area_gap(ch1, 80)
    set_overlap(ch1, -10)

    # Believability delta
    b_deltas = [delta(q4, q3) for q4, q3 in zip(believ_q4, believ_q3)]
    add_delta_table(slide, labels, b_deltas, 5.15, chart_top, 0.55, row_h * 0.85,
                    header_text="B \u0394")

    # Right: Composite Effectiveness
    textbox(slide, "Composite Effectiveness (%)", 5.90, 1.72, 3.5, 0.25,
            fsize=9, bold=True, color=C_GREY, font=FONT_DISPLAY)

    cd2 = CategoryChartData()
    cd2.categories = labels
    cd2.add_series("Q4'25", comp_q4)
    cd2.add_series("Q3'25", comp_q3)

    cf2 = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(5.90), Inches(chart_top), Inches(4.8), Inches(chart_h), cd2)
    ch2 = cf2.chart
    ch2.has_legend = False
    set_series_color(ch2.series[0], C_RYB_Q4)
    set_series_no_border(ch2.series[0])
    set_series_color(ch2.series[1], C_RYB_LIGHT)
    set_series_no_border(ch2.series[1])
    _enable_data_labels(ch2.series[0], C_RYB_Q4, fsize=7)
    _enable_data_labels(ch2.series[1], C_RYB_LIGHT, fsize=7)
    hide_axis(ch2, "val")
    hide_cat_labels(ch2)
    invert_cat_axis(ch2)
    set_plot_area_gap(ch2, 80)
    set_overlap(ch2, -10)

    # Composite delta
    c_deltas = [delta(q4, q3) for q4, q3 in zip(comp_q4, comp_q3)]
    add_delta_table(slide, labels, c_deltas, 10.75, chart_top, 0.55, row_h * 0.85,
                    header_text="CE \u0394")

    # Legend
    ly = chart_top + chart_h + 0.15
    solidrect(slide, 2.5, ly, 0.18, 0.12, C_RYB_Q4)
    textbox(slide, "Q4'25", 2.75, ly - 0.02, 0.6, 0.18, fsize=7, color=C_GREY)
    solidrect(slide, 3.5, ly, 0.18, 0.12, C_RYB_LIGHT)
    textbox(slide, "Q3'25", 3.75, ly - 0.02, 0.6, 0.18, fsize=7, color=C_GREY)

    slide_footer(slide, "Source: Lung SFEA SB — Additonal Analysis | B = Believability, CE = Composite Effectiveness")


def build_cta(prs, data):
    """Slide 7 — Ask 5: Call to Action (J&J vs AZ)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide,
        "RYB reps improved on 'changed opinion'; branded closing dipped slightly vs Q3")

    section_header_bar(slide,
        "CALL TO ACTION — J&J (RYB+LAZ) vs AZ (TAGRISSO) — Q4'25", top=1.40)

    ryb = data["ryb_cta"]
    tag = data["tag_cta"]

    if not ryb or not tag:
        textbox(slide, "CTA data extraction pending", 2, 3, 8, 1,
                fsize=14, color=C_RED)
        return

    # Align labels
    labels = [r["label"] for r in ryb]
    ryb_q4 = [r["q4"] for r in ryb]
    tag_q4_vals = []
    for r in ryb:
        match = [t for t in tag if t["label"] == r["label"]]
        tag_q4_vals.append(match[0]["q4"] if match else 0)
    ryb_deltas_v = [delta(r["q4"], r["q3"]) for r in ryb]
    tag_deltas_v = []
    for r in ryb:
        match = [t for t in tag if t["label"] == r["label"]]
        tag_deltas_v.append(delta(match[0]["q4"], match[0]["q3"]) if match else None)

    n = len(labels)
    chart_top = 1.85
    chart_h = min(3.5, n * 0.75)
    row_h = chart_h / max(n, 1)

    # Clustered bar
    cd = CategoryChartData()
    cd.categories = labels
    cd.add_series("J&J (RYB+LAZ)", ryb_q4)
    cd.add_series("AZ (Tagrisso)", tag_q4_vals)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.40), Inches(chart_top), Inches(8.0), Inches(chart_h), cd)
    ch = cf.chart
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.font.size = Pt(8)
    ch.legend.font.name = FONT_TEXT

    s0 = ch.series[0]
    set_series_color(s0, C_RYB_Q4)
    set_series_no_border(s0)
    _enable_data_labels(s0, C_RYB_Q4)

    s1 = ch.series[1]
    set_series_color(s1, C_TAG_Q4)
    set_series_no_border(s1)
    _enable_data_labels(s1, C_TAG_Q4)

    hide_axis(ch, "val")
    ch.category_axis.tick_labels.font.size = Pt(8)
    ch.category_axis.tick_labels.font.name = FONT_TEXT
    invert_cat_axis(ch)
    set_plot_area_gap(ch, 100)
    set_overlap(ch, -15)

    # Delta tables
    add_delta_table(slide, labels, ryb_deltas_v, 8.60, chart_top, 0.65, row_h * 0.85,
                    header_text="J&J \u0394")
    add_delta_table(slide, labels, tag_deltas_v, 9.35, chart_top, 0.65, row_h * 0.85,
                    header_text="AZ \u0394")

    # Insight callout
    callout_box(slide, 10.30, chart_top, 2.80, 2.5,
        text="Key CTA shifts:\n"
             "\u2022 'Changed opinion' rose +8pp for J&J\n"
             "\u2022 Branded closing dipped -4pp\n"
             "\u2022 TAG CTA largely flat QoQ",
        border_color=C_RYB_Q4, fsize=7.5)

    slide_footer(slide, "Source: Lung SFEA SB — C1_81Z, C1_82Z, C1_83D1Z, Q1_84bZ | Top Box / Yes% | RYB N=100, TAG N=73")


def build_followup_reps(prs, data):
    """Slide 8 — Ask 7: One J&J Vision — Follow-up Reps."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide,
        "Over half of HCPs did not set follow-up with other J&J reps; MSL referrals declined in Q4")

    section_header_bar(slide,
        "ONE J&J VISION — FOLLOW-UP WITH OTHER J&J REPRESENTATIVES — Q4'25", top=1.40)

    fu = data["ryb_followup"]
    if not fu:
        textbox(slide, "Follow-up data not available", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    labels = [f["desc"][:40] for f in fu]
    q4_vals = [f["q4"] for f in fu]
    q3_vals = [f["q3"] for f in fu]
    n = len(labels)
    chart_top = 1.85
    chart_h = min(3.5, n * 0.55)
    row_h = chart_h / max(n, 1)

    # Bar chart
    cd = CategoryChartData()
    cd.categories = labels
    cd.add_series("Q4'25", q4_vals)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.40), Inches(chart_top), Inches(8.0), Inches(chart_h), cd)
    ch = cf.chart
    ch.has_legend = False
    s = ch.series[0]
    set_series_color(s, C_RYB_Q4)
    set_series_no_border(s)
    _enable_data_labels(s, C_RYB_Q4)
    hide_axis(ch, "val")
    ch.category_axis.tick_labels.font.size = Pt(8)
    ch.category_axis.tick_labels.font.name = FONT_TEXT
    invert_cat_axis(ch)
    set_plot_area_gap(ch, 80)

    # Delta table
    fu_deltas = [delta(q4, q3) for q4, q3 in zip(q4_vals, q3_vals)]
    add_delta_table(slide, labels, fu_deltas, 8.60, chart_top, 0.65, row_h * 0.85,
                    header_text="QoQ \u0394")

    slide_footer(slide, "Source: Lung SFEA SB — C1_84FZ | % of HCPs who set follow-up | RYB Q3 N=103, Q4 N=100")


def build_rx_intent(prs, data):
    """Slide 9 — Ask 8: Prescription Intent by Patient Type."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide,
        "RYB intent to increase Rx rose sharply for patients w/o CNS (+10pp); TAG chemo intent held steady")

    section_header_bar(slide,
        "PRESCRIPTION INTENT (INCREASE Rx) BY PATIENT TYPE — Q4'25", top=1.40)

    ryb = data["ryb_rx"]
    tag = data["tag_rx"]

    chart_top = 1.85

    # RYB section
    textbox(slide, "RYB+LAZ — Intent to Increase Rx", 0.40, 1.72, 5.0, 0.25,
            fsize=9, bold=True, color=C_RYB_Q4, font=FONT_DISPLAY)

    if ryb:
        labels_r = [r["desc"][:35] for r in ryb]
        q4_r = [r["q4"] for r in ryb]
        q3_r = [r["q3"] for r in ryb]
        n_r = len(labels_r)
        ch_h_r = max(1.2, n_r * 0.55)

        cd_r = CategoryChartData()
        cd_r.categories = labels_r
        cd_r.add_series("Q4'25", q4_r)
        cf_r = slide.shapes.add_chart(
            XL_CHART_TYPE.BAR_CLUSTERED,
            Inches(0.40), Inches(chart_top), Inches(5.5), Inches(ch_h_r), cd_r)
        ch_r = cf_r.chart
        ch_r.has_legend = False
        s = ch_r.series[0]
        set_series_color(s, C_RYB_Q4)
        set_series_no_border(s)
        _enable_data_labels(s, C_RYB_Q4)
        hide_axis(ch_r, "val")
        ch_r.category_axis.tick_labels.font.size = Pt(8)
        ch_r.category_axis.tick_labels.font.name = FONT_TEXT
        invert_cat_axis(ch_r)
        set_plot_area_gap(ch_r, 80)

        # RYB delta
        r_deltas = [delta(q4, q3) for q4, q3 in zip(q4_r, q3_r)]
        add_delta_table(slide, labels_r, r_deltas, 6.0, chart_top, 0.55, ch_h_r / max(n_r, 1) * 0.85,
                        header_text="\u0394")
    else:
        ch_h_r = 1.2

    # TAG section
    tag_top = chart_top + ch_h_r + 0.50
    textbox(slide, "Tagrisso+Chemo — Intent to Increase Rx", 0.40, tag_top - 0.15, 5.0, 0.25,
            fsize=9, bold=True, color=C_TAG_Q4, font=FONT_DISPLAY)

    if tag:
        labels_t = [t["desc"][:35] for t in tag]
        q4_t = [t["q4"] for t in tag]
        q3_t = [t["q3"] for t in tag]
        n_t = len(labels_t)
        ch_h_t = max(1.2, n_t * 0.55)

        cd_t = CategoryChartData()
        cd_t.categories = labels_t
        cd_t.add_series("Q4'25", q4_t)
        cf_t = slide.shapes.add_chart(
            XL_CHART_TYPE.BAR_CLUSTERED,
            Inches(0.40), Inches(tag_top), Inches(5.5), Inches(ch_h_t), cd_t)
        ch_t = cf_t.chart
        ch_t.has_legend = False
        s = ch_t.series[0]
        set_series_color(s, C_TAG_Q4)
        set_series_no_border(s)
        _enable_data_labels(s, C_TAG_Q4)
        hide_axis(ch_t, "val")
        ch_t.category_axis.tick_labels.font.size = Pt(8)
        ch_t.category_axis.tick_labels.font.name = FONT_TEXT
        invert_cat_axis(ch_t)
        set_plot_area_gap(ch_t, 80)

        # TAG delta
        t_deltas = [delta(q4, q3) for q4, q3 in zip(q4_t, q3_t)]
        add_delta_table(slide, labels_t, t_deltas, 6.0, tag_top, 0.55, ch_h_t / max(n_t, 1) * 0.85,
                        header_text="\u0394")

    slide_footer(slide, "Source: Lung SFEA SB — C1_85DZ (RYB), C1_85D_TAG2_mNSCLCZ (TAG) | Top 2 Box (6,7)")


def build_share_of_time(prs, data):
    """Slide 10 — Ask 9: Share of Interactions by Indication/Topic."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide,
        "Efficacy discussions surged in Q4 (+22pp); safety and guidelines topics saw shifts")

    section_header_bar(slide,
        "SHARE OF INTERACTION TIME BY TOPIC — RYB+LAZ — Q4'25 vs Q3'25", top=1.40)

    sot = data["share_of_time"]
    if not sot:
        textbox(slide, "Share of time data not available", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    # Sort by Q4 descending
    sot_sorted = sorted(sot, key=lambda x: x["q4"] or 0, reverse=True)
    labels = [s["label"][:35] for s in sot_sorted]
    q4_vals = [s["q4"] or 0 for s in sot_sorted]
    q3_vals = [s["q3"] or 0 for s in sot_sorted]

    n = len(labels)
    chart_top = 1.85
    chart_h = min(4.0, n * 0.50)
    row_h = chart_h / max(n, 1)

    # Clustered bar (Q4 vs Q3)
    cd = CategoryChartData()
    cd.categories = labels
    cd.add_series("Q4'25", q4_vals)
    cd.add_series("Q3'25", q3_vals)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.40), Inches(chart_top), Inches(8.0), Inches(chart_h), cd)
    ch = cf.chart
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.font.size = Pt(8)
    ch.legend.font.name = FONT_TEXT

    s0 = ch.series[0]
    set_series_color(s0, C_RYB_Q4)
    set_series_no_border(s0)
    _enable_data_labels(s0, C_RYB_Q4)

    s1 = ch.series[1]
    set_series_color(s1, C_RYB_LIGHT)
    set_series_no_border(s1)
    _enable_data_labels(s1, C_RYB_LIGHT)

    hide_axis(ch, "val")
    ch.category_axis.tick_labels.font.size = Pt(8)
    ch.category_axis.tick_labels.font.name = FONT_TEXT
    invert_cat_axis(ch)
    set_plot_area_gap(ch, 80)
    set_overlap(ch, -10)

    # Delta
    sot_deltas = [delta(q4, q3) for q4, q3 in zip(q4_vals, q3_vals)]
    add_delta_table(slide, labels, sot_deltas, 8.60, chart_top, 0.65, row_h * 0.85,
                    header_text="QoQ \u0394")

    slide_footer(slide, "Source: Lung SFEA SB — VQ_Share_of_Time_RYB_SFEA | % of interactions | Whole-number percentages")


def build_hii_compare(prs, data):
    """Slide 11 — Nebulous Ask 2: High Impact vs Non-High Impact characteristics."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide,
        "High-impact interactions drive significantly higher rep quality scores across all key metrics")

    section_header_bar(slide,
        "HIGH IMPACT vs NON-HIGH IMPACT INTERACTIONS — KEY DIFFERENTIATORS — Q4'25", top=1.40)

    hii = data["hii_compare"]
    if not hii:
        textbox(slide, "HII comparison data not available", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    # Take top 10 differentiators
    hii_top = hii[:min(10, len(hii))]
    labels = [h["desc"] for h in hii_top]
    hi_vals = [h["hi_q4"] for h in hii_top]
    other_vals = [h["other_q4"] for h in hii_top]

    n = len(labels)
    chart_top = 1.85
    chart_h = min(4.5, n * 0.42)

    # Clustered bar
    cd = CategoryChartData()
    cd.categories = labels
    cd.add_series("High Impact", hi_vals)
    cd.add_series("Non-High Impact", other_vals)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.40), Inches(chart_top), Inches(9.5), Inches(chart_h), cd)
    ch = cf.chart
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.font.size = Pt(8)
    ch.legend.font.name = FONT_TEXT

    s0 = ch.series[0]
    set_series_color(s0, C_ORANGE_HI)
    set_series_no_border(s0)
    _enable_data_labels(s0, C_ORANGE_HI)

    s1 = ch.series[1]
    set_series_color(s1, C_GREY_LO)
    set_series_no_border(s1)
    _enable_data_labels(s1, C_GREY_LO)

    hide_axis(ch, "val")
    ch.category_axis.tick_labels.font.size = Pt(7)
    ch.category_axis.tick_labels.font.name = FONT_TEXT
    invert_cat_axis(ch)
    set_plot_area_gap(ch, 80)
    set_overlap(ch, -10)

    # Difference column
    diff_top = chart_top
    diff_vals = [h["diff"] for h in hii_top]
    row_h = chart_h / max(n, 1)
    add_delta_table(slide, labels, diff_vals, 10.10, diff_top, 0.70, row_h * 0.85,
                    header_text="Gap (pp)")

    # Callout
    callout_box(slide, 10.90, diff_top, 2.20, 2.0,
        text="Orange = High Impact\nGrey = Non-High Impact\n\n"
             "Attributes with largest gaps are the strongest differentiators of high-quality interactions",
        border_color=C_ORANGE_HI, fsize=7)

    slide_footer(slide, "Source: Lung SFEA SB — Q1_87Z | High Impact = Top Box Overall Quality + Intent to Prescribe | RYB Q4 N=100")


def build_quality_overview(prs, data):
    """Slide 12 — Overall Quality Metrics (supports nebulous asks)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide,
        "Overall quality metrics for RYB reps remained stable to positive in Q4; key areas held strong")

    section_header_bar(slide,
        "OVERALL CALL QUALITY METRICS — RYB+LAZ — Q4'25", top=1.40)

    qual = data["ryb_quality"]
    if not qual:
        textbox(slide, "Quality data not available", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    # Sort by Q4 descending
    qual_sorted = sorted(qual, key=lambda x: x["q4"] or 0, reverse=True)
    labels = [q["desc"] for q in qual_sorted]
    q4_vals = [q["q4"] for q in qual_sorted]
    q3_vals = [q["q3"] for q in qual_sorted]

    n = len(labels)
    chart_top = 1.85
    chart_h = min(4.8, n * 0.35)
    row_h = chart_h / max(n, 1)

    # Bar chart
    cd = CategoryChartData()
    cd.categories = labels
    cd.add_series("Q4'25", q4_vals)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.30), Inches(chart_top), Inches(9.0), Inches(chart_h), cd)
    ch = cf.chart
    ch.has_legend = False
    s = ch.series[0]
    set_series_color(s, C_RYB_Q4)
    set_series_no_border(s)
    _enable_data_labels(s, C_RYB_Q4)
    hide_axis(ch, "val")
    ch.category_axis.tick_labels.font.size = Pt(6.5)
    ch.category_axis.tick_labels.font.name = FONT_TEXT
    invert_cat_axis(ch)
    set_plot_area_gap(ch, 65)

    # Delta
    q_deltas = [delta(q4, q3) for q4, q3 in zip(q4_vals, q3_vals)]
    add_delta_table(slide, labels, q_deltas, 9.40, chart_top, 0.65, row_h * 0.85,
                    header_text="QoQ \u0394")

    slide_footer(slide, "Source: Lung SFEA SB — Q1_87Z | Top 2 Box (6,7) | RYB Q3 N=103, Q4 N=100")


def build_mariposa_rep_perf(prs, data):
    """Slide 13 — Nebulous Ask 4: Rep Performance segmented by Mariposa 1st Discussed vs NOT."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide,
        "Mariposa-first interactions show higher rep performance on compelling reason to Rx and knowledge metrics")

    section_header_bar(slide,
        "REP PERFORMANCE — MARIPOSA 1st DISCUSSED vs NOT — Q4'25", top=1.40)

    mrep = data["mariposa_rep"]
    mhii = data["mariposa_hii"]

    if not mrep:
        textbox(slide, "Mariposa segmented data not available", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    # Sort by Mariposa-1st descending
    mrep_sorted = sorted(mrep, key=lambda x: x["mariposa_1st"] or 0, reverse=True)

    labels = [m["metric"] for m in mrep_sorted]
    m1st_vals = [m["mariposa_1st"] for m in mrep_sorted]
    mnot_vals = [m["mariposa_not"] for m in mrep_sorted]

    n = len(labels)
    chart_top = 1.85
    chart_h = min(4.8, n * 0.32)
    row_h = chart_h / max(n, 1)

    # Clustered bar: Mariposa 1st vs NOT
    cd = CategoryChartData()
    cd.categories = labels
    cd.add_series("Mariposa 1st (n=58)", m1st_vals)
    cd.add_series("Not 1st (n=42)", mnot_vals)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.30), Inches(chart_top), Inches(8.5), Inches(chart_h), cd)
    ch = cf.chart
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.font.size = Pt(8)
    ch.legend.font.name = FONT_TEXT

    s0 = ch.series[0]
    set_series_color(s0, C_RYB_Q4)
    set_series_no_border(s0)
    _enable_data_labels(s0, C_RYB_Q4, fsize=7)

    s1 = ch.series[1]
    set_series_color(s1, C_GREY_BAR)
    set_series_no_border(s1)
    _enable_data_labels(s1, C_GREY_BAR, fsize=7)

    hide_axis(ch, "val")
    ch.category_axis.tick_labels.font.size = Pt(6.5)
    ch.category_axis.tick_labels.font.name = FONT_TEXT
    invert_cat_axis(ch)
    set_plot_area_gap(ch, 65)
    set_overlap(ch, -15)

    # Gap column
    gaps = [round((m1 or 0) - (mn or 0), 1) for m1, mn in zip(m1st_vals, mnot_vals)]
    add_delta_table(slide, labels, gaps, 9.10, chart_top, 0.65, row_h * 0.85,
                    header_text="Gap (pp)")

    # HII callout
    if mhii:
        callout_box(slide, 10.00, chart_top, 3.10, 2.5,
            text=f"High Impact Interactions by MARIPOSA segment:\n\n"
                 f"\u2022 MARIPOSA 1st discussed: {mhii.get('hii_first', 'N/A')}%\n"
                 f"\u2022 Discussed later: {mhii.get('hii_later', 'N/A')}%\n"
                 f"\u2022 Discussed most: {mhii.get('hii_most', 'N/A')}%\n"
                 f"\u2022 Discussed less: {mhii.get('hii_less', 'N/A')}%\n\n"
                 "Prioritizing MARIPOSA in discussions correlates with higher HII",
            border_color=C_RYB_Q4, fsize=7.5)

    slide_footer(slide, "Source: Lung SFEA SB — Additonal Analysis | Top 2 Box (6,7) | Overlapped=Mariposa 1st, Non-overlapped=Not 1st | RYB Q4 N=100")


def build_msg_recall_order(prs, data):
    """Slide 14 — Nebulous Ask 5: TAG Message Recall Order (1st, 2nd, 3rd recalled)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide,
        "TAG message recall order: NCCN and mPFS messages most frequently recalled first; OS headline leads 2nd recall")

    section_header_bar(slide,
        "TAGRISSO MESSAGE RECALL ORDER — 1st, 2nd, 3rd RECALLED — Q4'25", top=1.40)

    order = data["tag_order"]
    if not order:
        textbox(slide, "Message recall order data not available", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    # Take top 10 by total recall
    order_top = order[:min(10, len(order))]

    labels = [o["desc"][:40] for o in order_top]
    first_vals = [o["1st_q4"] for o in order_top]
    second_vals = [o["2nd_q4"] for o in order_top]
    third_vals = [o["3rd_q4"] for o in order_top]
    fourth_vals = [o["4th_q4"] for o in order_top]
    total_vals = [o["total_q4"] for o in order_top]

    n = len(labels)
    chart_top = 1.85
    chart_h = min(4.5, n * 0.42)
    row_h = chart_h / max(n, 1)

    # Stacked bar chart: 1st + 2nd + 3rd + 4th recalled
    cd = CategoryChartData()
    cd.categories = labels
    cd.add_series("1st Recalled", first_vals)
    cd.add_series("2nd Recalled", second_vals)
    cd.add_series("3rd Recalled", third_vals)
    cd.add_series("4th+ Recalled", fourth_vals)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_STACKED,
        Inches(0.30), Inches(chart_top), Inches(8.5), Inches(chart_h), cd)
    ch = cf.chart
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.font.size = Pt(7.5)
    ch.legend.font.name = FONT_TEXT

    # Color scheme: dark to light purple gradient (higher contrast between 3rd/4th)
    stack_colors = [
        RGBColor(0x3B, 0x0F, 0x6B),  # darkest purple - 1st
        RGBColor(0x70, 0x30, 0xA0),  # medium purple - 2nd
        RGBColor(0xB0, 0x80, 0xD0),  # light purple - 3rd
        RGBColor(0xE0, 0xD0, 0xF0),  # very pale lavender - 4th+
    ]
    label_colors = [
        RGBColor(0xFF, 0xFF, 0xFF),  # white on dark
        RGBColor(0xFF, 0xFF, 0xFF),  # white on medium
        RGBColor(0x30, 0x10, 0x50),  # dark on light
        RGBColor(0x30, 0x10, 0x50),  # dark on palest
    ]

    all_vals = [first_vals, second_vals, third_vals, fourth_vals]
    for idx, series in enumerate(ch.series):
        set_series_color(series, stack_colors[idx])
        set_series_no_border(series)
        _enable_data_labels(series, label_colors[idx], fsize=6, pos="ctr")
        # Hide labels on small segments (≤3%) to prevent overlap
        for pt_idx, val in enumerate(all_vals[idx]):
            if val <= 3:
                # Add a dLbl element with delete=1 for this point
                dLbls = series._element.find(qn("c:dLbls"))
                if dLbls is not None:
                    dLbl = etree.SubElement(dLbls, qn("c:dLbl"))
                    idx_el = etree.SubElement(dLbl, qn("c:idx"))
                    idx_el.set("val", str(pt_idx))
                    delete_el = etree.SubElement(dLbl, qn("c:delete"))
                    delete_el.set("val", "1")

    hide_axis(ch, "val")
    ch.category_axis.tick_labels.font.size = Pt(7)
    ch.category_axis.tick_labels.font.name = FONT_TEXT
    invert_cat_axis(ch)
    set_plot_area_gap(ch, 80)

    # Total recall column
    total_deltas = [round(t, 1) for t in total_vals]
    # Use a simple table for totals (not delta coloring)
    tbl_shape = slide.shapes.add_table(n + 1, 1,
        Inches(9.10), Inches(chart_top), Inches(0.70), Inches(0.28 + n * row_h * 0.85))
    tbl = tbl_shape.table
    tbl.rows[0].height = Inches(0.28)
    hc = tbl.cell(0, 0)
    hc.fill.solid()
    hc.fill.fore_color.rgb = C_HDRGREY
    p = hc.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "Total %"
    run.font.size = Pt(7)
    run.font.bold = True
    run.font.color.rgb = C_WHITE
    run.font.name = FONT_TEXT
    for i, t in enumerate(total_deltas):
        tbl.rows[i + 1].height = Inches(row_h * 0.85)
        cell = tbl.cell(i + 1, 0)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_LBGREY if i % 2 == 0 else C_WHITE
        tf = cell.text_frame
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        run = tf.paragraphs[0].add_run()
        run.text = f"{t:.0f}%"
        run.font.size = Pt(8)
        run.font.bold = True
        run.font.color.rgb = C_TAG_Q4
        run.font.name = FONT_TEXT
    tbl.columns[0].width = Inches(0.70)

    # QoQ delta column
    qoq_deltas = [delta(o["total_q4"], o["total_q3"]) for o in order_top]
    add_delta_table(slide, labels, qoq_deltas, 9.90, chart_top, 0.65, row_h * 0.85,
                    header_text="QoQ \u0394")

    slide_footer(slide, "Source: Lung SFEA SB — Q2_20Z | % recalled by order | TAG Q3 N=73, Q4 N=73")


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    print("Loading data...")
    data = load_data()

    print(f"  RYB messages: {len(data['ryb_mr'])}")
    print(f"  TAG messages: {len(data['tag_mr'])}")
    print(f"  Rep perf metrics: {len(data['rep_perf'])}")
    print(f"  MBD messages: {len(data['mbd'])}")
    print(f"  RYB CTA: {len(data['ryb_cta'])}")
    print(f"  TAG CTA: {len(data['tag_cta'])}")
    print(f"  Follow-up reps: {len(data['ryb_followup'])}")
    print(f"  RYB Rx intent: {len(data['ryb_rx'])}")
    print(f"  TAG Rx intent: {len(data['tag_rx'])}")
    print(f"  Share of time: {len(data['share_of_time'])}")
    print(f"  HII compare: {len(data['hii_compare'])}")
    print(f"  Quality metrics: {len(data['ryb_quality'])}")
    print(f"  Mariposa rep perf: {len(data['mariposa_rep'])}")
    print(f"  TAG recall order: {len(data['tag_order'])}")

    # Create fresh presentation with same dimensions as template
    print("\nCreating presentation...")
    prs = Presentation()
    prs.slide_width = Emu(12192000)   # 13.33 inches
    prs.slide_height = Emu(6858000)   # 7.50 inches

    print("Building slides...")

    # Slide 1: Cover
    build_cover(prs)
    print("  [1] Cover")

    # Slide 2: Executive Summary
    build_es(prs, data)
    print("  [2] Executive Summary")

    # Slide 3: RYB MR/ME (Ask 1)
    build_ryb_mr_me(prs, data)
    print("  [3] RYB Message Recall & Effectiveness")

    # Slide 4: TAG MR/ME (Ask 2)
    build_tag_mr_me(prs, data)
    print("  [4] TAG Message Recall & Effectiveness")

    # Slide 5: Rep Performance (Ask 3)
    build_rep_perf(prs, data)
    print("  [5] Rep Performance — J&J vs AZ")

    # Slide 6: M/B/D (Ask 4)
    build_mbd(prs, data)
    print("  [6] Message Believability & Composite Effectiveness")

    # Slide 7: CTA (Ask 5)
    build_cta(prs, data)
    print("  [7] Call to Action")

    # Slide 8: Follow-up Reps (Ask 7)
    build_followup_reps(prs, data)
    print("  [8] One J&J Vision — Follow-up Reps")

    # Slide 9: Rx Intent (Ask 8)
    build_rx_intent(prs, data)
    print("  [9] Prescription Intent by Patient Type")

    # Slide 10: Share of Time (Ask 9)
    build_share_of_time(prs, data)
    print("  [10] Share of Interaction Time")

    # Slide 11: HII Compare (Nebulous 2)
    build_hii_compare(prs, data)
    print("  [11] High Impact vs Non-High Impact")

    # Slide 12: Quality Overview
    build_quality_overview(prs, data)
    print("  [12] Overall Quality Metrics")

    # Slide 13: Mariposa Rep Performance (Nebulous 4)
    build_mariposa_rep_perf(prs, data)
    print("  [13] Mariposa 1st Discussed vs NOT — Rep Performance")

    # Slide 14: TAG Message Recall Order (Nebulous 5)
    build_msg_recall_order(prs, data)
    print("  [14] TAG Message Recall Order")

    # Save
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    prs.save(OUT)
    print(f"\nSaved to: {OUT}")
    print("Done!")


if __name__ == "__main__":
    main()
