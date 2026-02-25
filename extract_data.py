"""
extract_data.py
Extracts and structures all data needed for the 9-slide Rybrevant Associate Asks deck.
Reads from 'Lung SFEA SB.xlsx' and saves structured data to 'slide_data.pkl'.

Column mapping (all sheets use 0-indexed, no header):
  RYB sheet  : col 0=code, col 1=desc, col 7=Q3_Total, col 17=Q4_Total
  TAG sheet  : col 0=code, col 1=desc, col 7=Q3_Total, col 13=Q4_Total
  AA sheet   : col 1=metric/label, col 2=RYB_Q3, col 3=RYB_Q4,
               col 4=TAG_Q3,  col 5=TAG_Q4, col 7=RYB_delta, col 8=TAG_delta
               (for message table: col3=MR_Q3, col4=MR_Q4, col6=ME_Q3, col8=ME_Q4)
"""
import pandas as pd
import pickle
import os
import math

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(BASE_DIR, "Lung SFEA SB.xlsx")

# ─── Load sheets (no header) ──────────────────────────────────────────────────
ryb = pd.read_excel(XLSX, sheet_name="RYB",               header=None)
tag = pd.read_excel(XLSX, sheet_name="TAG",               header=None)
aa  = pd.read_excel(XLSX, sheet_name="Additonal Analysis", header=None)

def safe_float(val):
    try:
        f = float(val)
        return None if math.isnan(f) else round(f, 4)
    except (TypeError, ValueError):
        return None

def pct(val):
    """Convert 0-1 decimal to percentage string or None."""
    v = safe_float(val)
    return round(v * 100, 1) if v is not None else None

# ─── Helper: extract answer rows below a question header ──────────────────────
def get_answer_rows(df, q_col, v_col_q3, v_col_q4, start_row, n_rows):
    """
    From 'df', starting at 'start_row'+1, collect 'n_rows' answer rows.
    Returns list of dicts {code, desc, q3, q4, delta}.
    """
    results = []
    for i in range(start_row + 1, start_row + 1 + n_rows):
        if i >= len(df):
            break
        r = df.iloc[i]
        code = str(r[q_col]) if pd.notna(r[q_col]) else None
        if code in (None, 'nan', 'NaN'):
            break
        desc_raw = str(r[1]) if pd.notna(r[1]) else ''
        q3 = pct(r[v_col_q3])
        q4 = pct(r[v_col_q4])
        delta = round(q4 - q3, 1) if (q3 is not None and q4 is not None) else None
        results.append({"code": code, "desc": desc_raw[:80], "q3": q3, "q4": q4, "delta": delta})
    return results

# ─── Helper: find row index of a specific question code ───────────────────────
def find_row(df, code, col=0):
    for i, val in enumerate(df[col]):
        if str(val) == code:
            return i
    return None


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 – RYB + Lazcluze Messaging  (MR + ME)
# Source: Additional Analysis rows 30-39  (0-indexed, pandas)
# Cols: 1=msg_text, 2=tag, 3=MR_Q3, 4=MR_Q4, 6=ME_Q3(top-box), 8=ME_Q4(top-box)
# ══════════════════════════════════════════════════════════════════════════════
slide1 = []
for i in range(30, 40):
    r = aa.iloc[i]
    tag_label = str(r[2])[:30] if pd.notna(r[2]) else ''
    mr_q3  = pct(r[3])
    mr_q4  = pct(r[4])
    me_q3  = pct(r[6])    # ME top-box Q3
    me_q4  = pct(r[8])    # ME top-box Q4
    me_bel_q4 = pct(r[7]) # ME Believable Q4 (secondary)
    mr_d   = round(mr_q4 - mr_q3, 1) if (mr_q3 and mr_q4) else None
    me_d   = round(me_q4 - me_q3, 1) if (me_q3 and me_q4) else None
    slide1.append({
        "tag":       tag_label,
        "mr_q3":     mr_q3,
        "mr_q4":     mr_q4,
        "me_q3":     me_q3,
        "me_q4":     me_q4,
        "me_bel_q4": me_bel_q4,
        "mr_delta":  mr_d,
        "me_delta":  me_d,
    })

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 – Tagrisso Messaging  (MR only)
# Source: TAG sheet rows 224-234 (Q2_10Z answer rows)
# ══════════════════════════════════════════════════════════════════════════════
TAG_MSG_SHORT = {
    "A21": "mPFS 29.4mo vs mono",
    "A22": "CNS mPFS w/ chemo",
    "A28": "89% stay on treatment",
    "A29": "ARs manageable",
    "A30": "Onset frequency/severity",
    "A31": "Longest mPFS & mOS",
    "A32": "Final OS 47.5mo",
    "A33": "NCCN Cat 1 (mono+chemo)",
    "A34": "Median exposure 30.5mo",
    "A35": "43% risk reduction (new Q4)",
    "A36": "Once-daily oral (new Q4)",
}
slide2 = []
for i in range(224, 235):
    r = tag.iloc[i]
    code = str(r[0])
    if code in ('nan', 'NaN') or not code.startswith('A'):
        continue
    q3 = pct(r[7])
    q4 = pct(r[13])
    delta = round(q4 - q3, 1) if (q3 is not None and q4 is not None) else None
    short = TAG_MSG_SHORT.get(code, code)
    slide2.append({"code": code, "label": short, "q3": q3, "q4": q4, "delta": delta})

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 – J&J vs AZ Rep Performance  (Abacus chart)
# Source: Additional Analysis rows 5-19
# Cols: 1=metric, 2=RYB_Q3, 3=RYB_Q4, 4=TAG_Q3, 5=TAG_Q4
# ══════════════════════════════════════════════════════════════════════════════
METRIC_SHORT = {
    "Overall quality of sales call": "Overall quality",
    "How knowledgeable the [COMPANY] rep was about [PRODUCT] for EGFR+ mNSCLC": "Knowledge: product",
    "How organized the [COMPANY] rep appeared to be": "Organization",
    "How well the [COMPANY] rep performed in providing a compelling reason to prescribe [PRODUCT] for EGFR+ mNSCLC": "Compelling reason",
    "How well the [COMPANY] rep performed on providing credible support for [PRODUCT] claims in treating EGFR+ mNSCLC": "Credible support",
    "How prepared the [COMPANY] rep was for the conversation": "Preparedness",
    "How engaging the rep was while discussing [PRODUCT] for EGFR+ mNSCLC": "Engaging",
    "How well the rep was able to ask meaningful questions ": "Meaningful questions",
    "How knowledgeable the [COMPANY] rep was about EGFR mNSCLC Exon 20 mutation": "Knowledge: Exon 20",
    "How well the [COMPANY] rep made valuable use of my time": "Valuable use of time",
    "How well the rep was able to address my questions about [PRODUCT] for EGFR+ mNSCLC": "Addressed questions",
    "How well the rep performed in providing a compelling reason to use broader molecular testing": "Molecular testing reason",
    "How well the [COMPANY] rep performed on delivering a clear message about [PRODUCT] for EGFR+ mNSCLC": "Clear message",
    "How well the rep was able to tailor the discussions to my needs and/or the needs of my practice": "Tailored discussion",
    "How knowledgeable the rep was about competitor products": "Competitor knowledge",
}

slide3 = []
for i in range(5, 20):
    r = aa.iloc[i]
    metric = str(r[1]).strip() if pd.notna(r[1]) else ''
    if not metric or metric == 'nan':
        continue
    short = METRIC_SHORT.get(metric, metric[:45])
    ryb_q3 = pct(r[2])
    ryb_q4 = pct(r[3])
    tag_q3 = pct(r[4])
    tag_q4 = pct(r[5])
    ryb_d  = round(ryb_q4 - ryb_q3, 1) if (ryb_q3 and ryb_q4) else None
    tag_d  = round(tag_q4 - tag_q3, 1) if (tag_q3 and tag_q4) else None
    slide3.append({
        "metric":  short,
        "ryb_q3":  ryb_q3,
        "ryb_q4":  ryb_q4,
        "tag_q3":  tag_q3,
        "tag_q4":  tag_q4,
        "ryb_d":   ryb_d,
        "tag_d":   tag_d,
    })
# Sort by RYB Q4 descending
slide3.sort(key=lambda x: x["ryb_q4"] or 0, reverse=True)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 – Message Component Analysis (Q4 MR vs ME)
# Source: Additional Analysis rows 30-39 (same source as Slide 1 but Q4 focus)
# ══════════════════════════════════════════════════════════════════════════════
slide4 = []
for i in range(30, 40):
    r = aa.iloc[i]
    tag_label = str(r[2])[:30] if pd.notna(r[2]) else ''
    mr_q4     = pct(r[4])
    me_bel_q4 = pct(r[7])   # ME Believable
    me_tb_q4  = pct(r[8])   # ME top-box
    slide4.append({
        "tag":        tag_label,
        "mr_q4":      mr_q4,
        "me_bel_q4":  me_bel_q4,
        "me_tb_q4":   me_tb_q4,
    })

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 – J&J vs AZ Call to Action (CTA)
# RYB: first answer row (Yes%) for each CTA question
# TAG: first answer row (Yes%) for each CTA question
# ══════════════════════════════════════════════════════════════════════════════
CTA_LABELS = ["Compelling Reason", "Changed Opinion", "Direct Prescribe Ask", "Education Ask"]

ryb_cta_rows = [121, 127, 133, 149]   # "Yes" rows
tag_cta_rows = [81,  90,  99,  125]   # "Yes" rows (Mono for C1_81/82/84b, A1 for C1_83D)

slide5 = []
for label, ri, ti in zip(CTA_LABELS, ryb_cta_rows, tag_cta_rows):
    r_ryb = ryb.iloc[ri]
    r_tag = tag.iloc[ti]
    ryb_q3 = pct(r_ryb[7]);  ryb_q4 = pct(r_ryb[17])
    tag_q3 = pct(r_tag[7]);  tag_q4 = pct(r_tag[13])
    slide5.append({
        "label":   label,
        "ryb_q3":  ryb_q3, "ryb_q4": ryb_q4,
        "ryb_d":   round(ryb_q4 - ryb_q3, 1) if (ryb_q3 and ryb_q4) else None,
        "tag_q3":  tag_q3, "tag_q4": tag_q4,
        "tag_d":   round(tag_q4 - tag_q3, 1) if (tag_q3 and tag_q4) else None,
    })

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 – RYB Message Recall Trend (Q3 → Q4, two-point)
# Source: RYB sheet rows 236-245 (Q2_10Z answer rows)
# ══════════════════════════════════════════════════════════════════════════════
RYB_MSG_SHORT = {
    "R21": "Indication (1L EGFR+)",
    "R23": "mPFS +7.1mo vs TAG",
    "R26": "MOA inhibition",
    "R28": "Safety: ARs timeline",
    "R31": "Prophylaxis (SKIPPirr/COCOON)",
    "R32": "OS Headline (unmatched survival)",
    "R33": "Median OS not reached",
    "R36": "CNS PFS 2× at 36mo",
    "R43": "NCCN Cat 1 NSCLC (new Q4)",
    "R44": "NCCN CNS preferred (new Q4)",
}
slide6 = []
for i in range(236, 246):
    r = ryb.iloc[i]
    code = str(r[0])
    if code in ('nan', 'NaN'):
        continue
    q3 = pct(r[7])
    q4 = pct(r[17])
    delta = round(q4 - q3, 1) if (q3 is not None and q4 is not None) else None
    slide6.append({
        "code":  code,
        "label": RYB_MSG_SHORT.get(code, code),
        "q3":    q3,
        "q4":    q4,
        "delta": delta,
    })

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 – One J&J Vision – Follow-ups (J&J vs AZ)
# Source: RYB rows 167-172, TAG rows 139-144
# ══════════════════════════════════════════════════════════════════════════════
FOLLOWUP_LABELS = ["MSL", "KAM", "FRM", "CNE", "No follow-up"]
ryb_fu_rows = [167, 168, 169, 170, 171]
tag_fu_rows = [139, 140, 141, 142, 143]

slide7 = []
for label, ri, ti in zip(FOLLOWUP_LABELS, ryb_fu_rows, tag_fu_rows):
    r_ryb = ryb.iloc[ri]
    r_tag = tag.iloc[ti]
    ryb_q3 = pct(r_ryb[7]);   ryb_q4 = pct(r_ryb[17])
    tag_q3 = pct(r_tag[7]);   tag_q4 = pct(r_tag[13])
    slide7.append({
        "label":   label,
        "ryb_q3":  ryb_q3, "ryb_q4": ryb_q4,
        "ryb_d":   round(ryb_q4 - ryb_q3, 1) if (ryb_q3 is not None and ryb_q4 is not None) else None,
        "tag_q3":  tag_q3, "tag_q4": tag_q4,
        "tag_d":   round(tag_q4 - tag_q3, 1) if (tag_q3 is not None and tag_q4 is not None) else None,
    })

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 – RYB vs TAG Prescription Intent
# RYB: C1_85AZ (prescribe), C1_85BZ (increase), C1_85CZ (by patient type), C1_85DZ
# TAG: C1_85A, C1_85B, C1_85C_TAG1 (Mono by patient type), C1_85C_TAG2 (+Chemo)
# ══════════════════════════════════════════════════════════════════════════════

# Prescribe Likelihood (Top 2 Box % on 7-pt scale)
slide8 = {
    "prescribe_1l": {
        "label": "Prescribe Likelihood\n(1L EGFR+ mNSCLC)",
        "ryb_q3": pct(ryb.iloc[176][7]),   # C1_85AZ A1 Q3
        "ryb_q4": pct(ryb.iloc[176][17]),  # C1_85AZ A1 Q4
        "tag_q3": pct(tag.iloc[148][7]),   # C1_85A A1 (Mono) Q3
        "tag_q4": pct(tag.iloc[148][13]),  # C1_85A A1 (Mono) Q4
    },
    "increase_1l": {
        "label": "Increase Rx Likelihood\n(1L EGFR+ mNSCLC)",
        "ryb_q3": pct(ryb.iloc[180][7]),   # C1_85BZ A1 Q3
        "ryb_q4": pct(ryb.iloc[180][17]),  # C1_85BZ A1 Q4
        "tag_q3": pct(tag.iloc[153][7]),   # C1_85B A1 (Mono) Q3
        "tag_q4": pct(tag.iloc[153][13]),  # C1_85B A1 (Mono) Q4
    },
    "prescribe_wcns": {
        "label": "Prescribe Likelihood\n(With CNS Mets)",
        "ryb_q3": pct(ryb.iloc[184][7]),   # C1_85CZ WCNS Q3
        "ryb_q4": pct(ryb.iloc[184][17]),  # C1_85CZ WCNS Q4
        "tag_q3": pct(tag.iloc[158][7]),   # C1_85C_TAG1 WCNS Q3 (Mono)
        "tag_q4": pct(tag.iloc[158][13]),  # C1_85C_TAG1 WCNS Q4
    },
    "prescribe_wocns": {
        "label": "Prescribe Likelihood\n(Without CNS Mets)",
        "ryb_q3": pct(ryb.iloc[185][7]),   # C1_85CZ WOCNS Q3
        "ryb_q4": pct(ryb.iloc[185][17]),  # C1_85CZ WOCNS Q4
        "tag_q3": pct(tag.iloc[159][7]),   # C1_85C_TAG1 WOCNS Q3 (Mono)
        "tag_q4": pct(tag.iloc[159][13]),  # C1_85C_TAG1 WOCNS Q4
    },
    "increase_wcns": {
        "label": "Increase Rx Likelihood\n(With CNS Mets)",
        "ryb_q3": pct(ryb.iloc[189][7]),   # C1_85DZ WCNS Q3
        "ryb_q4": pct(ryb.iloc[189][17]),
        "tag_q3": pct(tag.iloc[168][7]),   # C1_85D_TAG1 WCNS Q3 (Mono)
        "tag_q4": pct(tag.iloc[168][13]),
    },
    "increase_wocns": {
        "label": "Increase Rx Likelihood\n(Without CNS Mets)",
        "ryb_q3": pct(ryb.iloc[190][7]),   # C1_85DZ WOCNS Q3
        "ryb_q4": pct(ryb.iloc[190][17]),
        "tag_q3": pct(tag.iloc[169][7]),   # C1_85D_TAG1 WOCNS Q3 (Mono)
        "tag_q4": pct(tag.iloc[169][13]),
    },
}
# Compute deltas
for k, v in slide8.items():
    v["ryb_d"] = round(v["ryb_q4"] - v["ryb_q3"], 1) if (v["ryb_q3"] and v["ryb_q4"]) else None
    v["tag_d"] = round(v["tag_q4"] - v["tag_q3"], 1) if (v["tag_q3"] and v["tag_q4"]) else None

# Convert to list for easy iteration
slide8_list = list(slide8.values())

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 9 – RYB 1L EGFR / Mariposa Discussions
# C1_09AZ (who initiated): RYB rows 8-10 (header + A1 + A2)
# AA MARIPOSA section rows 61-63 (scores 6, 7, top2box) with 4 columns
# ══════════════════════════════════════════════════════════════════════════════

# C1_09AZ: physician vs rep initiated
c109_header_row = 8
slide9_initiation = {
    "physician_initiated": {
        "q3": pct(ryb.iloc[9][7]),   # A1 Q3
        "q4": pct(ryb.iloc[9][17]),  # A1 Q4
    },
    "rep_initiated": {
        "q3": pct(ryb.iloc[10][7]),  # A2 Q3
        "q4": pct(ryb.iloc[10][17]), # A2 Q4
    },
}
for k, v in slide9_initiation.items():
    v["delta"] = round(v["q4"] - v["q3"], 1) if (v["q3"] and v["q4"]) else None

# MARIPOSA discussion quality (from AA, Q4 only)
# Row 61: score=6, [Discussed1st-first, Discussed1st-later, DiscussedMost-less, DiscussedMost-most]
# Row 62: score=7
# Row 63: top2box
aa_mar_cols = {
    "Discussed 1st – First": 2,
    "Discussed 1st – Later": 3,
    "Discussed Most – Less": 4,
    "Discussed Most – Most": 5,
}
slide9_mariposa_q4 = {}
for lbl, col in aa_mar_cols.items():
    score6 = safe_float(aa.iloc[61][col])
    score7 = safe_float(aa.iloc[62][col])
    t2b    = safe_float(aa.iloc[63][col])
    slide9_mariposa_q4[lbl] = {
        "score6": score6,
        "score7": score7,
        "top2box": t2b,
    }

# Q1_50Z topics discussed in RYB (Slide 9 context)
slide9_topics = []
for i in range(76, 90):
    r = ryb.iloc[i]
    code = str(r[0])
    if code in ('nan', 'NaN'):
        break
    topic = str(r[1])[:50] if pd.notna(r[1]) else code
    q3 = pct(r[7])
    q4 = pct(r[17])
    delta = round(q4 - q3, 1) if (q3 is not None and q4 is not None) else None
    slide9_topics.append({"code": code, "topic": topic, "q3": q3, "q4": q4, "delta": delta})

# ─── Bundle and save ──────────────────────────────────────────────────────────
all_data = {
    "slide1_ryb_messaging":    slide1,
    "slide2_tag_messaging":    slide2,
    "slide3_rep_perf":         slide3,
    "slide4_msg_components":   slide4,
    "slide5_cta":              slide5,
    "slide6_recall_trend":     slide6,
    "slide7_followups":        slide7,
    "slide8_rx_intent":        slide8_list,
    "slide9_mariposa": {
        "initiation":  slide9_initiation,
        "mariposa_q4": slide9_mariposa_q4,
        "topics":      slide9_topics,
    },
    "meta": {
        "ryb_n_q3": 103, "ryb_n_q4": 100,
        "tag_n_q3":  73, "tag_n_q4":  73,
    }
}

out_pkl = os.path.join(BASE_DIR, "slide_data.pkl")
with open(out_pkl, "wb") as f:
    pickle.dump(all_data, f)

print(f"Data extracted and saved to: {out_pkl}")
print("\nSlide data summary:")
print(f"  Slide 1 – RYB Messaging:       {len(slide1)} messages")
print(f"  Slide 2 – TAG Messaging:       {len(slide2)} messages")
print(f"  Slide 3 – Rep Performance:     {len(slide3)} metrics")
print(f"  Slide 4 – Msg Components:      {len(slide4)} messages")
print(f"  Slide 5 – CTA:                 {len(slide5)} items")
print(f"  Slide 6 – Recall Trend:        {len(slide6)} messages")
print(f"  Slide 7 – Follow-ups:          {len(slide7)} types")
print(f"  Slide 8 – Rx Intent:           {len(slide8_list)} scenarios")
print(f"  Slide 9 – Mariposa:            {len(slide9_topics)} topics, {len(slide9_mariposa_q4)} segments")
