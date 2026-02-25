"""
validate_data_v2.py  –  Verifies every data point used in Rybrevant_Analysis_Deck_v2.pptx
against the raw source in Lung SFEA SB.xlsx.

Output: validation_report_v2.txt (and console)

Checks:
  1. RYB Messaging MR (Q3/Q4) from AA sheet rows 30-39
  2. RYB Messaging ME (Q3/Q4) from AA sheet rows 30-39
  3. TAG Messaging MR (Q3/Q4) from TAG sheet rows 224-234
  4. Rep Performance (Q3/Q4) from AA sheet rows 5-19 for both brands
  5. Message Components (Believable/ME) from AA rows 30-39
  6. Call to Action rows from RYB/TAG sheets
  7. Message Recall Trend from RYB sheet Q2_10Z rows
  8. Follow-ups from RYB/TAG sheets
  9. Prescription Intent from RYB/TAG sheets
 10. Initiation (C1_09AZ) from RYB sheet
 11. Topics discussed (Q1_50Z) from RYB sheet
 12. High Impact data from AA rows 207-219
 13. TAG Recall Order from TAG Q2_20Z
"""
import pandas as pd
import pickle
import math
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(BASE_DIR, "Lung SFEA SB.xlsx")

ryb = pd.read_excel(XLSX, sheet_name="RYB",               header=None)
tag = pd.read_excel(XLSX, sheet_name="TAG",               header=None)
aa  = pd.read_excel(XLSX, sheet_name="Additonal Analysis", header=None)

with open(os.path.join(BASE_DIR, "slide_data_v2.pkl"), "rb") as f:
    D = pickle.load(f)


def sf(val):
    try:
        f = float(val)
        return None if math.isnan(f) else round(f, 4)
    except (TypeError, ValueError):
        return None


def pct(val):
    v = sf(val)
    return round(v * 100, 1) if v is not None else None


TOL = 0.15  # tolerance in percentage points for float comparison

results = []
pass_count = 0
fail_count = 0
skip_count = 0


def check(label, slide, expected, actual, source_ref):
    global pass_count, fail_count, skip_count
    if expected is None and actual is None:
        status = "SKIP (both None)"
        skip_count += 1
    elif expected is None:
        status = f"SKIP (source=None, deck={actual})"
        skip_count += 1
    elif actual is None:
        status = f"FAIL  (source={expected}, deck=None)"
        fail_count += 1
    elif abs(expected - actual) <= TOL:
        status = f"PASS  ({expected} == {actual})"
        pass_count += 1
    else:
        status = f"FAIL  (source={expected}, deck={actual}, diff={actual-expected:+.1f}pp)"
        fail_count += 1
    results.append((slide, label, source_ref, status))


# ══════════════════════════════════════════════════════════════════════════════
# 1. RYB Messaging MR + ME  (Slide 3)
#    Source: AA rows 30-39
# ══════════════════════════════════════════════════════════════════════════════
slide3_data = D["s3_ryb_messaging"]

# Rebuild raw from AA
aa_msgs_raw = []
for i in range(30, 40):
    r = aa.iloc[i]
    if pd.isna(r[1]):
        continue
    tag_lbl = str(r[2])[:35] if pd.notna(r[2]) else ''
    aa_msgs_raw.append({
        "tag":    tag_lbl,
        "mr_q3":  pct(r[3]),
        "mr_q4":  pct(r[4]),
        "me_q3":  pct(r[6]),
        "me_q4":  pct(r[8]),
    })

# Match by tag label
for d in slide3_data:
    match = next((x for x in aa_msgs_raw if x["tag"] == d["tag"]), None)
    lbl = d["tag"][:25]
    if match is None:
        results.append(("Slide 3", f"MR tag '{lbl}'", "AA 30-39", "FAIL  (no matching AA row)"))
        fail_count += 1
        continue
    check(f"MR Q3 '{lbl}'",    "Slide 3", match["mr_q3"], d["mr_q3"],  "AA col3")
    check(f"MR Q4 '{lbl}'",    "Slide 3", match["mr_q4"], d["mr_q4"],  "AA col4")
    check(f"ME Q3 '{lbl}'",    "Slide 3", match["me_q3"], d["me_q3"],  "AA col6")
    check(f"ME Q4 '{lbl}'",    "Slide 3", match["me_q4"], d["me_q4"],  "AA col8")
    # Delta checks
    if d["mr_q3"] is not None and d["mr_q4"] is not None:
        exp_delta = round(d["mr_q4"] - d["mr_q3"], 1)
        check(f"MR delta '{lbl}'", "Slide 3", exp_delta, d["mr_delta"], "computed")

# ══════════════════════════════════════════════════════════════════════════════
# 2. TAG Messaging MR  (Slide 4)
#    Source: TAG sheet rows 224-234
# ══════════════════════════════════════════════════════════════════════════════
slide4_data = D["s4_tag_messaging"]
tag_msg_raw = {}
for i in range(224, 235):
    r = tag.iloc[i]
    code = str(r[0])
    if not code.startswith('A'):
        continue
    tag_msg_raw[code] = {"q3": pct(r[7]), "q4": pct(r[13])}

for d in slide4_data:
    raw = tag_msg_raw.get(d["code"])
    lbl = d["code"]
    if raw is None:
        results.append(("Slide 4", f"TAG msg '{lbl}'", "TAG 224-234", "FAIL (not found in source)"))
        fail_count += 1
        continue
    check(f"TAG MR Q3 '{lbl}'", "Slide 4", raw["q3"], d["q3"], f"TAG row ~{224} col7")
    check(f"TAG MR Q4 '{lbl}'", "Slide 4", raw["q4"], d["q4"], f"TAG row ~{224} col13")
    if d["q3"] is not None and d["q4"] is not None:
        check(f"TAG MR delta '{lbl}'", "Slide 4",
              round(d["q4"] - d["q3"], 1), d["delta"], "computed")

# ══════════════════════════════════════════════════════════════════════════════
# 3. Rep Performance  (Slide 5)
#    Source: AA rows 5-19
# ══════════════════════════════════════════════════════════════════════════════
slide5_data = D["s5_rep_perf"]
aa_rep_raw = {}
for i in range(5, 20):
    r = aa.iloc[i]
    metric = str(r[1]).strip() if pd.notna(r[1]) else ''
    if not metric or metric == 'nan':
        continue
    aa_rep_raw[metric] = {
        "ryb_q3": pct(r[2]), "ryb_q4": pct(r[3]),
        "tag_q3": pct(r[4]), "tag_q4": pct(r[5]),
    }

for d in slide5_data:
    raw = aa_rep_raw.get(d["full"])
    lbl = d["metric"][:25]
    if raw is None:
        results.append(("Slide 5", f"Rep '{lbl}'", "AA 5-19", "FAIL (metric not found)"))
        fail_count += 1
        continue
    check(f"RYB Q3 '{lbl}'", "Slide 5", raw["ryb_q3"], d["ryb_q3"], "AA col2")
    check(f"RYB Q4 '{lbl}'", "Slide 5", raw["ryb_q4"], d["ryb_q4"], "AA col3")
    check(f"TAG Q3 '{lbl}'", "Slide 5", raw["tag_q3"], d["tag_q3"], "AA col4")
    check(f"TAG Q4 '{lbl}'", "Slide 5", raw["tag_q4"], d["tag_q4"], "AA col5")

# ══════════════════════════════════════════════════════════════════════════════
# 4. Message Components  (Slide 6)
#    Source: AA rows 30-39
# ══════════════════════════════════════════════════════════════════════════════
slide6_data = D["s6_msg_components"]
aa_mc_raw = []
for i in range(30, 40):
    r = aa.iloc[i]
    if pd.isna(r[1]):
        continue
    aa_mc_raw.append({
        "tag":      str(r[2])[:35] if pd.notna(r[2]) else '',
        "mr_q3":    pct(r[3]),
        "mr_q4":    pct(r[4]),
        "bel_q3":   pct(r[5]),
        "bel_q4":   pct(r[7]),
        "me_q3":    pct(r[6]),
        "me_q4":    pct(r[8]),
    })

for d in slide6_data:
    match = next((x for x in aa_mc_raw if x["tag"] == d["tag"]), None)
    lbl = d["tag"][:22]
    if match is None:
        continue
    check(f"MC MR Q4 '{lbl}'",  "Slide 6", match["mr_q4"],  d["mr_q4"],  "AA col4")
    check(f"MC Bel Q4 '{lbl}'", "Slide 6", match["bel_q4"], d["believable_q4"], "AA col7")
    check(f"MC ME Q4 '{lbl}'",  "Slide 6", match["me_q4"],  d["me_q4"],  "AA col8")

# ══════════════════════════════════════════════════════════════════════════════
# 5. Call to Action  (Slide 7)
#    Source: RYB rows 121,127,133,149  /  TAG rows 81,90,99,125
# ══════════════════════════════════════════════════════════════════════════════
slide7_data = D["s7_cta"]
CTA_LABELS = ["Compelling Reason to Prescribe", "Changed Opinion",
              "Direct Ask to Prescribe (Branded)", "Educational Ask"]
ryb_cta_rows = [121, 127, 133, 149]
tag_cta_rows = [81,  90,  99,  125]

for d, ri, ti in zip(slide7_data, ryb_cta_rows, tag_cta_rows):
    lbl = d["label"][:25]
    src_ryb_q3 = pct(ryb.iloc[ri][7]);  src_ryb_q4 = pct(ryb.iloc[ri][17])
    src_tag_q3 = pct(tag.iloc[ti][7]);  src_tag_q4 = pct(tag.iloc[ti][13])
    check(f"CTA RYB Q3 '{lbl}'", "Slide 7", src_ryb_q3, d["ryb_q3"], f"RYB row {ri} col7")
    check(f"CTA RYB Q4 '{lbl}'", "Slide 7", src_ryb_q4, d["ryb_q4"], f"RYB row {ri} col17")
    check(f"CTA TAG Q3 '{lbl}'", "Slide 7", src_tag_q3, d["tag_q3"], f"TAG row {ti} col7")
    check(f"CTA TAG Q4 '{lbl}'", "Slide 7", src_tag_q4, d["tag_q4"], f"TAG row {ti} col13")

# ══════════════════════════════════════════════════════════════════════════════
# 6. Message Recall Trend  (Slide 8)
#    Source: RYB rows 236-245
# ══════════════════════════════════════════════════════════════════════════════
slide8_data = D["s8_recall_trend"]
for d in slide8_data:
    # Find the row in RYB by code
    for i in range(236, 246):
        if str(ryb.iloc[i][0]) == d["code"]:
            src_q3 = pct(ryb.iloc[i][7])
            src_q4 = pct(ryb.iloc[i][17])
            lbl = d["code"]
            check(f"MR Q3 '{lbl}'", "Slide 8", src_q3, d["q3"], f"RYB row {i} col7")
            check(f"MR Q4 '{lbl}'", "Slide 8", src_q4, d["q4"], f"RYB row {i} col17")
            break

# ══════════════════════════════════════════════════════════════════════════════
# 7. Follow-ups  (Slide 9)
#    Source: RYB rows 167-171  /  TAG rows 139-143
# ══════════════════════════════════════════════════════════════════════════════
slide9_data = D["s9_followups"]
ryb_fu_rows = [167, 168, 169, 170, 171]
tag_fu_rows = [139, 140, 141, 142, 143]

for d, ri, ti in zip(slide9_data, ryb_fu_rows, tag_fu_rows):
    lbl = d["label"]
    check(f"FU RYB Q3 '{lbl}'", "Slide 9", pct(ryb.iloc[ri][7]),  d["ryb_q3"], f"RYB row {ri} col7")
    check(f"FU RYB Q4 '{lbl}'", "Slide 9", pct(ryb.iloc[ri][17]), d["ryb_q4"], f"RYB row {ri} col17")
    check(f"FU TAG Q3 '{lbl}'", "Slide 9", pct(tag.iloc[ti][7]),  d["tag_q3"], f"TAG row {ti} col7")
    check(f"FU TAG Q4 '{lbl}'", "Slide 9", pct(tag.iloc[ti][13]), d["tag_q4"], f"TAG row {ti} col13")

# ══════════════════════════════════════════════════════════════════════════════
# 8. Prescription Intent  (Slide 10)
#    Source: RYB rows 176,180,184,185,189,190  /  TAG rows 148,153,158,159,173,174
# ══════════════════════════════════════════════════════════════════════════════
slide10_data = D["s10_rx_intent"]
PI_ROWS = [
    (176, 148), (180, 153), (184, 158), (185, 159), (189, 173), (190, 174)
]
for d, (ri, ti) in zip(slide10_data, PI_ROWS):
    lbl = d["label"][:25]
    check(f"PI RYB Q3 '{lbl}'", "Slide 10", pct(ryb.iloc[ri][7]),  d["ryb_q3"], f"RYB row {ri} col7")
    check(f"PI RYB Q4 '{lbl}'", "Slide 10", pct(ryb.iloc[ri][17]), d["ryb_q4"], f"RYB row {ri} col17")
    check(f"PI TAG Q3 '{lbl}'", "Slide 10", pct(tag.iloc[ti][7]),  d["tag_q3"], f"TAG row {ti} col7")
    check(f"PI TAG Q4 '{lbl}'", "Slide 10", pct(tag.iloc[ti][13]), d["tag_q4"], f"TAG row {ti} col13")

# ══════════════════════════════════════════════════════════════════════════════
# 9. MARIPOSA Initiation (C1_09AZ)  (Slide 11)
#    Source: RYB rows 9-10
# ══════════════════════════════════════════════════════════════════════════════
mar_init = D["s11_mariposa"]["initiation"]
check("MARIPOSA rep init Q3",       "Slide 11", pct(ryb.iloc[10][7]),  mar_init["rep"]["q3"],       "RYB row 10 col7")
check("MARIPOSA rep init Q4",       "Slide 11", pct(ryb.iloc[10][17]), mar_init["rep"]["q4"],       "RYB row 10 col17")
check("MARIPOSA phys init Q3",      "Slide 11", pct(ryb.iloc[9][7]),   mar_init["physician"]["q3"], "RYB row 9 col7")
check("MARIPOSA phys init Q4",      "Slide 11", pct(ryb.iloc[9][17]),  mar_init["physician"]["q4"], "RYB row 9 col17")

# ══════════════════════════════════════════════════════════════════════════════
# 10. Topics discussed Q1_50Z  (Slide 11)
#     Source: RYB rows 76-89
# ══════════════════════════════════════════════════════════════════════════════
topics_data = D["s11_mariposa"]["topics"]
for d in topics_data:
    for i in range(76, 91):
        if str(ryb.iloc[i][0]) == d["code"]:
            lbl = d["code"]
            check(f"Topic Q3 '{lbl}'", "Slide 11", pct(ryb.iloc[i][7]),  d["q3"], f"RYB row {i} col7")
            check(f"Topic Q4 '{lbl}'", "Slide 11", pct(ryb.iloc[i][17]), d["q4"], f"RYB row {i} col17")
            break

# ══════════════════════════════════════════════════════════════════════════════
# 11. TAG Message Recall Order  (Slide 12)
#     Source: TAG rows 250-316 (Q2_20Z)
# ══════════════════════════════════════════════════════════════════════════════
order_data = D["s12_recall_order"]
# Build expected from TAG raw
pos_map = {"1st": "1st", "2nd": "2nd", "3rd": "3rd", "4th or later": "4th+", "4th": "4th+"}
expected_order = {}
for i in range(250, len(tag)):
    r = tag.iloc[i]
    code = str(r[0])
    if not code.startswith('A') or pd.isna(r[2]):
        continue
    pos = pos_map.get(str(r[2]).strip())
    val = sf(r[13])
    if code not in expected_order:
        expected_order[code] = {"1st": None, "2nd": None, "3rd": None, "4th+": None}
    if pos:
        expected_order[code][pos] = round(val * 100, 1) if val is not None else None

for d in order_data:
    # Find code by label matching TAG_MSG_SHORT inverse
    code = None
    for c, lbl in {
        "A21": "mPFS 29.4mo (chemo combo)", "A22": "CNS mPFS w/ chemo 24.9mo",
        "A28": "89% stay on treatment", "A29": "ARs mostly Gr1/2",
        "A30": "Onset/severity reduces over time", "A31": "Longest mPFS & mOS reported",
        "A32": "Final OS 47.5mo (~4 years)", "A33": "NCCN Cat 1 (mono+chemo)",
        "A34": "Median exposure 30.5mo", "A35": "43% risk reduction (CNS≥3 met) ★",
        "A36": "Once-daily oral dosing ★",
    }.items():
        if lbl == d["label"]:
            code = c
            break
    if code is None:
        continue
    exp = expected_order.get(code, {})
    for pos in ["1st", "2nd", "3rd", "4th+"]:
        check(f"TAG recall order {code} {pos}", "Slide 12",
              exp.get(pos), d[pos],
              f"TAG Q2_20Z code {code} pos {pos} col13")

# ══════════════════════════════════════════════════════════════════════════════
# 12. High Impact Interactions  (Slide 13)
#     Source: AA rows 210-217, col4=HII, col6=non-HII
# ══════════════════════════════════════════════════════════════════════════════
hii_data = D["s13_hii"]
aa_hii_raw = []
for i in range(210, 218):
    r = aa.iloc[i]
    if pd.isna(r[1]) or str(r[1]) == 'nan':
        continue
    topic = str(r[2])[:35] if pd.notna(r[2]) else ''
    aa_hii_raw.append({
        "topic": topic,
        "hii":     pct(r[4]),
        "non_hii": pct(r[6]),
    })

# All rows in AA 210-217 have col2="RYB", so topic-based matching always hits row 0.
# Both s13_hii (from extract) and aa_hii_raw are sorted by HII desc → match by index.
aa_hii_sorted = sorted(aa_hii_raw, key=lambda x: x["hii"] or 0, reverse=True)
for idx, d in enumerate(hii_data):
    if idx >= len(aa_hii_sorted):
        break
    match = aa_hii_sorted[idx]
    lbl = f"HII row {idx + 1}"
    check(f"HII val '{lbl}'",     "Slide 13", match["hii"],     d["hii"],     f"AA col4 (sorted row {idx+1})")
    check(f"Non-HII val '{lbl}'", "Slide 13", match["non_hii"], d["non_hii"], f"AA col6 (sorted row {idx+1})")

# ══════════════════════════════════════════════════════════════════════════════
# 13. High Impact % overall (Slide 15)
#     Source: AA row 47, cols 2-5
# ══════════════════════════════════════════════════════════════════════════════
cl = D["s15_closing"]
hi = cl["high_impact_ryb"]
# AA row 47 stores values as whole-number percentages (57, 63, 52, 57) — use sf() not pct()
check("High Impact RYB Q3", "Slide 15", sf(aa.iloc[47][2]), hi["q3"], "AA row 47 col2")
check("High Impact RYB Q4", "Slide 15", sf(aa.iloc[47][3]), hi["q4"], "AA row 47 col3")
check("High Impact TAG Q3", "Slide 15", sf(aa.iloc[47][4]), hi["tag_mono_q3"], "AA row 47 col4")
check("High Impact TAG Q4", "Slide 15", sf(aa.iloc[47][5]), hi["tag_mono_q4"], "AA row 47 col5")

# Top-2 box quality by visual aid (AA rows 110, cols 3-4)
check("Quality Top2Box no VA",   "Slide 15", pct(aa.iloc[110][3]), cl["high_q_t2b"]["no_va"],   "AA row 110 col3")
check("Quality Top2Box with VA", "Slide 15", pct(aa.iloc[110][4]), cl["high_q_t2b"]["with_va"], "AA row 110 col4")

# ══════════════════════════════════════════════════════════════════════════════
# Summary and report
# ══════════════════════════════════════════════════════════════════════════════
total = pass_count + fail_count + skip_count
pct_pass = round(pass_count / max(total - skip_count, 1) * 100, 1)

sep = "=" * 90
header = (
    sep + "\n" +
    "VALIDATION REPORT  -  Rybrevant_Analysis_Deck_v2.pptx vs Lung SFEA SB.xlsx\n" +
    sep + "\n" +
    f"Total checks : {total}\n" +
    f"  PASS       : {pass_count}\n" +
    f"  FAIL       : {fail_count}\n" +
    f"  SKIP (N/A) : {skip_count}\n" +
    f"  Match rate : {pct_pass}%  (excl. N/A)\n" +
    sep + "\n\n"
)

col_widths = (8, 35, 25, 55)
row_fmt = f"{{:<{col_widths[0]}}} {{:<{col_widths[1]}}} {{:<{col_widths[2]}}} {{}}\n"
detail_header = row_fmt.format("Slide", "Check", "Source Ref", "Status")
separator = "-" * 90 + "\n"

lines = [header, detail_header, separator]
current_slide = None
for (slide, lbl, src, status) in results:
    if slide != current_slide:
        lines.append(f"\n  ── {slide} ──\n")
        current_slide = slide
    lines.append(row_fmt.format(slide, lbl[:35], src[:25], status))

report_text = "".join(lines)

# Save report
report_path = os.path.join(BASE_DIR, "validation_report_v2.txt")
with open(report_path, "w", encoding="utf-8") as f:
    f.write(report_text)

import sys
sys.stdout.buffer.write(report_text.encode("utf-8", errors="replace"))
sys.stdout.buffer.write(f"\nReport saved to: {report_path}\n".encode("utf-8"))
