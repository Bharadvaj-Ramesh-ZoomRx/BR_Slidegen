"""
extract_data.py  –  Data extraction for all updated Associate Asks slides.
Reads Lung SFEA SB.xlsx and saves structured data to output/slide_data_v2.pkl.

Key survey IDs in Lung SFEA SB.xlsx:
  RYB IM sheet  : col7=Q3_Total,  col17=Q4_Total
  TAG sheet     : col7=Q3_Total,  col13=Q4_Total
  AA sheet      : col1=label, col2=RYB_Q3, col3=RYB_Q4, col4=TAG_Q3, col5=TAG_Q4
                  (message rows 30-39): col3=MR_Q3, col4=MR_Q4, col6=ME_Q3, col8=ME_Q4
"""
import pandas as pd
import pickle
import math
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(BASE_DIR, "Lung SFEA SB.xlsx")

ryb = pd.read_excel(XLSX, sheet_name="RYB",               header=None)
tag = pd.read_excel(XLSX, sheet_name="TAG",               header=None)
aa  = pd.read_excel(XLSX, sheet_name="Additonal Analysis", header=None)


def sf(val):
    try:
        f = float(val)
        return None if math.isnan(f) else round(f, 4)
    except (TypeError, ValueError):
        return None


def pct(val):
    v = sf(val)
    return round(v * 100, 1) if v is not None else None


def delta(a, b):
    """b - a; both must be non-None."""
    return round(b - a, 1) if (a is not None and b is not None) else None


# ─── Short message labels ──────────────────────────────────────────────────────
RYB_MSG_SHORT = {
    "R21": "Indication (1L EGFR+)",
    "R23": "mPFS +7.1mo vs osimertinib",
    "R26": "MOA: biology of disease",
    "R28": "Safety: ARs timeline",
    "R31": "Prophylaxis (SKIPPirr/COCOON)",
    "R32": "OS Headline (unmatched survival)",
    "R33": "Median OS not reached (>4yr)",
    "R36": "CNS PFS 2× at 36 months",
    "R43": "NCCN Cat 1 NSCLC ★",
    "R44": "NCCN CNS preferred ★",
}

TAG_MSG_SHORT = {
    "A21": "mPFS 29.4mo (chemo combo)",
    "A22": "CNS mPFS w/ chemo 24.9mo",
    "A28": "89% stay on treatment",
    "A29": "ARs mostly Gr1/2",
    "A30": "Onset/severity reduces over time",
    "A31": "Longest mPFS & mOS reported",
    "A32": "Final OS 47.5mo (~4 years)",
    "A33": "NCCN Cat 1 (mono+chemo)",
    "A34": "Median exposure 30.5mo",
    "A35": "43% risk reduction (CNS≥3 met) ★",
    "A36": "Once-daily oral dosing ★",
}

METRIC_SHORT = {
    "Overall quality of sales call":                                              "Overall quality",
    "How knowledgeable the [COMPANY] rep was about [PRODUCT] for EGFR+ mNSCLC":  "Knowledge: Product",
    "How organized the [COMPANY] rep appeared to be":                             "Organization",
    "How well the [COMPANY] rep performed in providing a compelling reason to prescribe [PRODUCT] for EGFR+ mNSCLC": "Compelling: Prescribe",
    "How well the [COMPANY] rep performed on providing credible support for [PRODUCT] claims in treating EGFR+ mNSCLC": "Credible support",
    "How prepared the [COMPANY] rep was for the conversation":                    "Preparedness",
    "How engaging the rep was while discussing [PRODUCT] for EGFR+ mNSCLC":      "Engaging",
    "How well the rep was able to ask meaningful questions ":                     "Meaningful questions",
    "How well the rep was able to ask meaningful questions":                      "Meaningful questions",
    "How knowledgeable the [COMPANY] rep was about EGFR mNSCLC Exon 20 mutation":"Knowledge: EGFR",
    "How well the [COMPANY] rep made valuable use of my time":                    "Valuable use of time",
    "How well the rep was able to address my questions about [PRODUCT] for EGFR+ mNSCLC": "Addressed questions",
    "How well the rep performed in providing a compelling reason to use broader molecular testing": "Molecular testing",
    "How well the [COMPANY] rep performed on delivering a clear message about [PRODUCT] for EGFR+ mNSCLC": "Clear message",
    "How well the rep was able to tailor the discussions to my needs and/or the needs of my practice": "Tailored discussion",
    "How knowledgeable the rep was about competitor products":                    "Competitor knowledge",
}

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 – RYB Messaging (MR + ME side by side)
# Source: AA rows 30-39
# Cols: 1=desc, 2=tag, 3=MR_Q3, 4=MR_Q4, 6=ME_Q3(composite), 8=ME_Q4(composite)
#       col5=ME_Bel_Q3, col7=ME_Bel_Q4
# ══════════════════════════════════════════════════════════════════════════════
s3_ryb_msg = []
for i in range(30, 40):
    r = aa.iloc[i]
    if pd.isna(r[1]):
        continue
    tag_label = str(r[2])[:35] if pd.notna(r[2]) else ''
    mr_q3  = pct(r[3])
    mr_q4  = pct(r[4])
    me_q3  = pct(r[6])   # ME composite Q3
    me_q4  = pct(r[8])   # ME composite Q4
    me_bel_q3 = pct(r[5])
    me_bel_q4 = pct(r[7])
    s3_ryb_msg.append({
        "tag":        tag_label,
        "mr_q3":      mr_q3,
        "mr_q4":      mr_q4,
        "mr_delta":   delta(mr_q3, mr_q4),
        "me_q3":      me_q3,
        "me_q4":      me_q4,
        "me_delta":   delta(me_q3, me_q4),
        "me_bel_q3":  me_bel_q3,
        "me_bel_q4":  me_bel_q4,
    })
# Sort by MR Q4 descending
s3_ryb_msg.sort(key=lambda x: x["mr_q4"] or 0, reverse=True)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 – TAG Messaging (MR; ME not available in this extract)
# Source: TAG sheet rows 224-234
# ══════════════════════════════════════════════════════════════════════════════
s4_tag_msg = []
for i in range(224, 235):
    r = tag.iloc[i]
    code = str(r[0])
    if code in ('nan', 'NaN') or not code.startswith('A'):
        continue
    q3 = pct(r[7])
    q4 = pct(r[13])
    short = TAG_MSG_SHORT.get(code, code)
    s4_tag_msg.append({
        "code":  code,
        "label": short,
        "q3":    q3,
        "q4":    q4,
        "delta": delta(q3, q4),
        "me_q3": None,  # TAG ME not in this xlsx
        "me_q4": None,
        "me_delta": None,
    })
# Sort by Q4 MR descending
s4_tag_msg.sort(key=lambda x: x["q4"] or 0, reverse=True)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 – J&J vs AZ Rep Performance  (Abacus, Q4 + QoQ delta)
# Source: AA rows 5-19
# Cols: 1=metric, 2=RYB_Q3, 3=RYB_Q4, 4=TAG_Q3, 5=TAG_Q4, 7=RYB_delta, 8=TAG_delta
# ══════════════════════════════════════════════════════════════════════════════
s5_rep_perf = []
for i in range(5, 20):
    r = aa.iloc[i]
    metric = str(r[1]).strip() if pd.notna(r[1]) else ''
    if not metric or metric == 'nan':
        continue
    short = METRIC_SHORT.get(metric, metric[:50])
    ryb_q3 = pct(r[2])
    ryb_q4 = pct(r[3])
    tag_q3 = pct(r[4])
    tag_q4 = pct(r[5])
    s5_rep_perf.append({
        "metric":  short,
        "full":    metric,
        "ryb_q3":  ryb_q3,
        "ryb_q4":  ryb_q4,
        "tag_q3":  tag_q3,
        "tag_q4":  tag_q4,
        "ryb_d":   delta(ryb_q3, ryb_q4),
        "tag_d":   delta(tag_q3, tag_q4),
        "jj_az_gap": delta(tag_q4, ryb_q4),  # RYB Q4 - TAG Q4
    })
# Keep AA order (already sorted descending by RYB Q4 in source)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 – Message Components (Believable + ME composite; M/D not available)
# Source: AA rows 30-39
# ══════════════════════════════════════════════════════════════════════════════
s6_msg_components = []
for i in range(30, 40):
    r = aa.iloc[i]
    if pd.isna(r[1]):
        continue
    tag_label = str(r[2])[:35] if pd.notna(r[2]) else ''
    mr_q3  = pct(r[3])
    mr_q4  = pct(r[4])
    me_bel_q3 = pct(r[5])
    me_comp_q3 = pct(r[6])
    me_bel_q4 = pct(r[7])
    me_comp_q4 = pct(r[8])
    s6_msg_components.append({
        "tag":          tag_label,
        "mr_q3":        mr_q3,
        "mr_q4":        mr_q4,
        "believable_q3": me_bel_q3,
        "believable_q4": me_bel_q4,
        "me_q3":        me_comp_q3,
        "me_q4":        me_comp_q4,
        "mr_delta":     delta(mr_q3, mr_q4),
        "bel_delta":    delta(me_bel_q3, me_bel_q4),
        "me_delta":     delta(me_comp_q3, me_comp_q4),
    })
# Sort by MR Q4 descending (per Ask 4: sort by Motivation/MR for Q4)
s6_msg_components.sort(key=lambda x: x["mr_q4"] or 0, reverse=True)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 – Call to Action  (most recent quarter = Q4)
# RYB: C1_81Z(121), C1_82Z(127), C1_83D1Z(133), Q1_84bZ(149)
# TAG: C1_81_TAG(81), C1_82_TAG(90), C1_83D_TAG(99), C1_84b_TAG(125)
# Row values = first answer row (Yes/top-box)
# ══════════════════════════════════════════════════════════════════════════════
CTA_LABELS = [
    "Compelling Reason to Prescribe",
    "Changed Opinion",
    "Direct Ask to Prescribe (Branded)",
    "Educational Ask",
]
ryb_cta = [121, 127, 133, 149]
tag_cta = [81,  90,  99,  125]

s7_cta = []
for lbl, ri, ti in zip(CTA_LABELS, ryb_cta, tag_cta):
    rr = ryb.iloc[ri]
    rt = tag.iloc[ti]
    rq3 = pct(rr[7]);  rq4 = pct(rr[17])
    tq3 = pct(rt[7]);  tq4 = pct(rt[13])
    s7_cta.append({
        "label":   lbl,
        "ryb_q3":  rq3, "ryb_q4": rq4, "ryb_d": delta(rq3, rq4),
        "tag_q3":  tq3, "tag_q4": tq4, "tag_d": delta(tq3, tq4),
    })

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 – Message Recall Trend (Q3→Q4 only; monthly data not available)
# Source: RYB sheet rows 236-245 (Q2_10Z answer rows)
# ══════════════════════════════════════════════════════════════════════════════
s8_recall_trend = []
for i in range(236, 246):
    r = ryb.iloc[i]
    code = str(r[0])
    if code in ('nan', 'NaN'):
        continue
    q3 = pct(r[7])
    q4 = pct(r[17])
    s8_recall_trend.append({
        "code":    code,
        "label":   RYB_MSG_SHORT.get(code, code),
        "q3":      q3,
        "q4":      q4,
        "delta":   delta(q3, q4),
        "new_q4":  q3 is None,  # new messages only in Q4
    })

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 9 – One J&J Vision: Follow-ups  (J&J vs AZ)
# RYB: C1_84FZ rows 167-172, TAG: C1_84F_TAG rows 139-144
# ══════════════════════════════════════════════════════════════════════════════
FOLLOWUP_LABELS = ["MSL", "KAM", "FRM", "CNE", "No follow-up"]
ryb_fu = [167, 168, 169, 170, 171]
tag_fu = [139, 140, 141, 142, 143]

s9_followups = []
for lbl, ri, ti in zip(FOLLOWUP_LABELS, ryb_fu, tag_fu):
    rr = ryb.iloc[ri]
    rt = tag.iloc[ti]
    rq3 = pct(rr[7]);  rq4 = pct(rr[17])
    tq3 = pct(rt[7]);  tq4 = pct(rt[13])
    s9_followups.append({
        "label":  lbl,
        "ryb_q3": rq3, "ryb_q4": rq4, "ryb_d": delta(rq3, rq4),
        "tag_q3": tq3, "tag_q4": tq4, "tag_d": delta(tq3, tq4),
    })

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 10 – Prescription Intent  (by patient type, current quarter)
# RYB: C1_85DZ WCNS(189)/WOCNS(190); TAG: C1_85D_TAG2_mNSCLCZ WCNS(173)/WOCNS(174)
# ══════════════════════════════════════════════════════════════════════════════
PI_ROWS = [
    ("Prescribe Likelihood: 1L EGFR+",        176, 148),  # C1_85AZ / C1_85A_TAG A1
    ("Increase Rx: 1L EGFR+",                 180, 153),  # C1_85BZ / C1_85B_TAG A1
    ("Prescribe Likelihood: w/ CNS mets",      184, 158),  # C1_85CZ WCNS / TAG1 WCNS
    ("Prescribe Likelihood: w/o CNS mets",     185, 159),  # C1_85CZ WOCNS / TAG1 WOCNS
    ("Increase Rx: w/ CNS mets",               189, 173),  # C1_85DZ WCNS / TAG2 WCNS
    ("Increase Rx: w/o CNS mets",              190, 174),  # C1_85DZ WOCNS / TAG2 WOCNS
]
s10_rx_intent = []
for lbl, ri, ti in PI_ROWS:
    rr = ryb.iloc[ri]
    rt = tag.iloc[ti]
    rq3 = pct(rr[7]);  rq4 = pct(rr[17])
    tq3 = pct(rt[7]);  tq4 = pct(rt[13])
    s10_rx_intent.append({
        "label":  lbl,
        "ryb_q3": rq3, "ryb_q4": rq4, "ryb_d": delta(rq3, rq4),
        "tag_q3": tq3, "tag_q4": tq4, "tag_d": delta(tq3, tq4),
    })

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 11 – MARIPOSA / Share of time  (topics discussed + initiation)
# C1_09AZ: RYB rows 8-10  (who initiated MARIPOSA discussion)
# Q1_50Z:  RYB rows 76-89 (topics discussed)
# AA rows 197-204: Share of time (Q3, Q4 by overlap/non-overlap)
# ══════════════════════════════════════════════════════════════════════════════
# Who initiated
s11_initiation = {
    "physician": {"q3": pct(ryb.iloc[9][7]),  "q4": pct(ryb.iloc[9][17])},
    "rep":       {"q3": pct(ryb.iloc[10][7]), "q4": pct(ryb.iloc[10][17])},
}
for k in s11_initiation:
    s11_initiation[k]["delta"] = delta(s11_initiation[k]["q3"], s11_initiation[k]["q4"])

# Topics discussed (Q1_50Z)
s11_topics = []
for i in range(76, 91):
    r = ryb.iloc[i]
    code = str(r[0])
    if code in ('nan', 'NaN') or code == '-oth-':
        continue
    topic = str(r[1])[:55] if pd.notna(r[1]) else code
    q3 = pct(r[7])
    q4 = pct(r[17])
    s11_topics.append({"code": code, "topic": topic, "q3": q3, "q4": q4, "delta": delta(q3, q4)})

# Share of time AA rows 197-204
s11_share_time = []
for i in range(197, 205):
    r = aa.iloc[i]
    if pd.isna(r[1]) or str(r[1]) == 'nan':
        continue
    topic = str(r[2])[:35] if pd.notna(r[2]) else ''
    q3_total = pct(r[4]) if sf(r[4]) is not None else None   # col4 = % Q3
    q4_total = pct(r[5]) if sf(r[5]) is not None else None   # col5 = % Q4
    s11_share_time.append({
        "topic":    topic,
        "q3":       q3_total,
        "q4":       q4_total,
        "delta":    delta(q3_total, q4_total),
    })

# MARIPOSA discussion quality (AA rows 61-63)
aa_mar_cols = {
    "Discussed 1st – First":  2,
    "Discussed 1st – Later":  3,
    "Discussed Most – Less":  4,
    "Discussed Most – Most":  5,
}
s11_mariposa_q4 = {}
for lbl, col in aa_mar_cols.items():
    s11_mariposa_q4[lbl] = {
        "top2box": sf(aa.iloc[63][col]),
    }

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 12 – Message Recall Order  (TAG Q2_20Z; RYB Q2_20Z not in this extract)
# Source: TAG sheet rows 250-316, cols: 0=code, 2=position, 7=Q3, 13=Q4
# Aggregate: group by message code, collect {1st, 2nd, 3rd, 4th} Q4 %
# ══════════════════════════════════════════════════════════════════════════════
# First build mapping of message code -> MR Q4 for sorting
tag_mr_q4 = {str(tag.iloc[i][0]): sf(tag.iloc[i][13]) for i in range(224, 235)
             if pd.notna(tag.iloc[i][0]) and str(tag.iloc[i][0]).startswith('A')}

s12_recall_order = {}
for i in range(250, len(tag)):
    r = tag.iloc[i]
    code = str(r[0])
    if not code.startswith('A') or pd.isna(r[2]):
        continue
    pos = str(r[2]).strip()
    q4  = sf(r[13])
    if code not in s12_recall_order:
        s12_recall_order[code] = {
            "label":     TAG_MSG_SHORT.get(code, code),
            "mr_q4":     tag_mr_q4.get(code),
            "1st":  None, "2nd":  None, "3rd":  None, "4th+": None,
        }
    mapping = {"1st": "1st", "2nd": "2nd", "3rd": "3rd", "4th or later": "4th+", "4th": "4th+"}
    key = mapping.get(pos)
    if key:
        s12_recall_order[code][key] = round(q4 * 100, 1) if q4 is not None else None

# Convert to sorted list (by MR Q4 descending)
s12_order_list = sorted(s12_recall_order.values(),
                        key=lambda x: x["mr_q4"] or 0, reverse=True)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 13 – High Impact Interactions  (characteristics by HII vs non-HII)
# Source: AA rows 207-219 (Q1_50Z topics by HII status, rolling Q1-Q4 2025)
# ══════════════════════════════════════════════════════════════════════════════
s13_hii = []
for i in range(210, 218):
    r = aa.iloc[i]
    if pd.isna(r[1]) or str(r[1]) == 'nan':
        continue
    topic = str(r[2])[:35] if pd.notna(r[2]) else ''
    hii_val   = pct(r[4])   # col4: HII
    non_hii   = pct(r[6])   # col6: not HII
    s13_hii.append({
        "topic":   topic,
        "hii":     hii_val,
        "non_hii": non_hii,
        "diff":    delta(non_hii, hii_val),   # HII - non-HII gap
    })
# Sort by HII value desc
s13_hii.sort(key=lambda x: x["hii"] or 0, reverse=True)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 14 – Rep Performance: MARIPOSA First vs Not  (Abacus)
# Source: AA rows 112-130 (visual aid segments for RYB) – closest available
# Actual MARIPOSA 1st segmented rep perf from AA rows 169-173 (partial) +
# high impact data from AA rows 84-102 (visual aid x rep perf Q4)
# Best source: AA rows 5-19 (overall) – MARIPOSA-first specific not fully available
# Fallback: use AA rows 84-102 which show rep perf by visual aid (proxy available)
# ══════════════════════════════════════════════════════════════════════════════
# Rep perf with MARIPOSA-first segmentation: AA rows 112-130
s14_mariposa_repperf = []
MARIPOSA_COLS = {
    "RYB_Q3_Overall":            3,  # Q3 2025 total
    "RYB_Q4_Overlapped":         4,  # Q4 2025 overlapped (high impact segment)
    "RYB_Q4_NonOverlapped":      5,  # Q4 2025 non-overlapped
    "RYB_Q4_Total":              4,  # AA rows 134-150 col3 = Q4 total
}
# Use AA rows 112-130 for rep perf by MARIPOSA/visual aid segments
for i in range(115, 130):
    r = aa.iloc[i]
    metric = str(r[1]).strip() if pd.notna(r[1]) else ''
    if not metric or metric == 'nan':
        continue
    short = METRIC_SHORT.get(metric, metric[:50])
    q3_val    = pct(r[2])
    q4_over   = pct(r[3])   # overlapped segment Q4
    q4_nonover= pct(r[4])   # non-overlapped Q4
    q4_total  = pct(r[3])   # use overlapped as "MARIPOSA discussed" proxy
    s14_mariposa_repperf.append({
        "metric":        short,
        "mariposa_perf": q4_over,    # High Impact / Overlapped
        "other_perf":    q4_nonover, # Non-overlapped
        "q3_total":      q3_val,
        "diff":          delta(q4_nonover, q4_over),
    })

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 15 – Closing Rates  (High/Low Quality × Closing)
# Source: AA rows 105-110 (top2box by visual aid) – partial proxy
# Actual segmentation: High Q × Closing / High Q × Not Closing / Low Q × Closing / etc.
# Note: Full cross-tab not in this extract; using available high-quality rates
# ══════════════════════════════════════════════════════════════════════════════
# Row 108: score=7 (top quality)
# Row 109: score=6 (high quality)
# Row 110: top2box (6+7)
# Cols: col3 = No visual aid, col4 = Visual aid used (proxy for engagement)
s15_closing = {
    "high_q_7":    {"no_va": pct(aa.iloc[108][3]), "with_va": pct(aa.iloc[108][4])},
    "high_q_6":    {"no_va": pct(aa.iloc[109][3]), "with_va": pct(aa.iloc[109][4])},
    "high_q_t2b":  {"no_va": pct(aa.iloc[110][3]), "with_va": pct(aa.iloc[110][4])},
    "high_impact_ryb": {
        # AA row 47 stores values as whole-number percentages (e.g. 57, not 0.57)
        # so use sf() not pct() to avoid multiplying by 100
        "q3": sf(aa.iloc[47][2]),   # AA row 47: High Impact Interactions RYB Q3
        "q4": sf(aa.iloc[47][3]),   # RYB Q4
        "tag_mono_q3": sf(aa.iloc[47][4]),
        "tag_mono_q4": sf(aa.iloc[47][5]),
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# Metadata
# ══════════════════════════════════════════════════════════════════════════════
meta = {
    "ryb_n_q3": 103,
    "ryb_n_q4": 100,
    "tag_n_q3":  73,
    "tag_n_q4":  73,
}

# ─── Bundle and save ──────────────────────────────────────────────────────────
all_data = {
    "s3_ryb_messaging":     s3_ryb_msg,
    "s4_tag_messaging":     s4_tag_msg,
    "s5_rep_perf":          s5_rep_perf,
    "s6_msg_components":    s6_msg_components,
    "s7_cta":               s7_cta,
    "s8_recall_trend":      s8_recall_trend,
    "s9_followups":         s9_followups,
    "s10_rx_intent":        s10_rx_intent,
    "s11_mariposa": {
        "initiation":       s11_initiation,
        "topics":           s11_topics,
        "share_time":       s11_share_time,
        "mariposa_q4":      s11_mariposa_q4,
    },
    "s12_recall_order":     s12_order_list,
    "s13_hii":              s13_hii,
    "s14_mariposa_repperf": s14_mariposa_repperf,
    "s15_closing":          s15_closing,
    "meta":                 meta,
}

out_pkl = os.path.join(BASE_DIR, "output", "slide_data_v2.pkl")
os.makedirs(os.path.dirname(out_pkl), exist_ok=True)
with open(out_pkl, "wb") as f:
    pickle.dump(all_data, f)

print(f"Data saved to: {out_pkl}")
print(f"  S3  RYB Messaging:        {len(s3_ryb_msg)} messages")
print(f"  S4  TAG Messaging:        {len(s4_tag_msg)} messages")
print(f"  S5  Rep Performance:      {len(s5_rep_perf)} metrics")
print(f"  S6  Msg Components:       {len(s6_msg_components)} messages")
print(f"  S7  CTA:                  {len(s7_cta)} items")
print(f"  S8  Recall Trend:         {len(s8_recall_trend)} messages")
print(f"  S9  Follow-ups:           {len(s9_followups)} types")
print(f"  S10 Rx Intent:            {len(s10_rx_intent)} scenarios")
print(f"  S11 MARIPOSA topics:      {len(s11_topics)} topics")
print(f"  S12 Recall Order (TAG):   {len(s12_order_list)} messages")
print(f"  S13 HII:                  {len(s13_hii)} topics")
print(f"  S14 MARIPOSA Rep Perf:    {len(s14_mariposa_repperf)} metrics")
