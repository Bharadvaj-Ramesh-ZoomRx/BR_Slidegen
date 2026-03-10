"""Explore Excel data for remaining slide asks."""
import pandas as pd
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(ROOT, "J_and_J_project", "data", "Lung SFEA SB.xlsx")

pd.set_option("display.max_columns", 30)
pd.set_option("display.width", 220)
pd.set_option("display.max_colwidth", 60)

# Load all sheets
xls = pd.ExcelFile(XLSX, engine="openpyxl")
print("=== SHEETS ===")
print(xls.sheet_names)

sheets = {}
for name in xls.sheet_names:
    sheets[name] = pd.read_excel(xls, sheet_name=name, header=None)

aa = sheets.get("Additonal Analysis")
ryb = sheets.get("RYB", sheets.get("Rybrevant"))
tag = sheets.get("TAG", sheets.get("Tagrisso"))

print(f"\nAA shape: {aa.shape if aa is not None else 'NOT FOUND'}")
print(f"RYB sheet names matching: {[s for s in xls.sheet_names if 'ryb' in s.lower() or 'rybrevant' in s.lower()]}")
print(f"TAG sheet names matching: {[s for s in xls.sheet_names if 'tag' in s.lower() or 'tagrisso' in s.lower()]}")

# ============================================================
# ASK 1: Closing Rates x Quality
# ============================================================
print("\n" + "="*80)
print("ASK 1: CLOSING RATES x QUALITY")
print("="*80)

if aa is not None:
    # Search for Closing_Rates, Q1_87Z, quality keywords
    for idx, row in aa.iterrows():
        row_str = " ".join(str(v) for v in row.values if pd.notna(v)).lower()
        if any(kw in row_str for kw in ["closing", "q1_87z", "quality", "high quality", "low quality"]):
            print(f"\nRow {idx}: {row.tolist()}")

    # Also scan broader area - check rows around where closing/quality might be
    print("\n--- Scanning AA rows 40-80 for closing/quality context ---")
    for idx in range(min(40, len(aa)), min(80, len(aa))):
        row = aa.iloc[idx]
        row_str = " ".join(str(v) for v in row.values if pd.notna(v)).lower()
        if any(kw in row_str for kw in ["closing", "quality", "87z", "high", "low"]):
            print(f"Row {idx}: {row.tolist()}")

    print("\n--- Scanning AA rows 80-130 ---")
    for idx in range(min(80, len(aa)), min(130, len(aa))):
        row = aa.iloc[idx]
        row_str = " ".join(str(v) for v in row.values if pd.notna(v)).lower()
        if any(kw in row_str for kw in ["closing", "quality", "87z", "high", "low"]):
            print(f"Row {idx}: {row.tolist()}")

    print("\n--- Scanning ALL AA rows for closing/quality ---")
    for idx in range(len(aa)):
        row = aa.iloc[idx]
        row_str = " ".join(str(v) for v in row.values if pd.notna(v)).lower()
        if any(kw in row_str for kw in ["closing_rate", "closing rate", "q1_87z"]):
            print(f"Row {idx}: {row.tolist()}")

# Search all sheets
print("\n--- Searching ALL sheets for closing/quality ---")
for sname, df in sheets.items():
    for idx, row in df.iterrows():
        row_str = " ".join(str(v) for v in row.values if pd.notna(v)).lower()
        if any(kw in row_str for kw in ["closing_rate", "closing rate", "q1_87z"]):
            print(f"[{sname}] Row {idx}: {row.tolist()}")

# ============================================================
# ASK 2: Rep performance sorted by importance / drivers
# ============================================================
print("\n" + "="*80)
print("ASK 2: REP PERFORMANCE - IMPORTANCE / DRIVER DATA")
print("="*80)

if aa is not None:
    # Print rows 0-25 to see rep performance area and headers
    print("\n--- AA rows 0-25 (rep performance area) ---")
    for idx in range(min(25, len(aa))):
        row = aa.iloc[idx]
        vals = [v for v in row.values if pd.notna(v)]
        if vals:
            print(f"Row {idx}: {row.tolist()}")

    # Search for importance/driver/weight keywords anywhere in AA
    print("\n--- AA: searching for importance/driver/weight ---")
    for idx in range(len(aa)):
        row = aa.iloc[idx]
        row_str = " ".join(str(v) for v in row.values if pd.notna(v)).lower()
        if any(kw in row_str for kw in ["importance", "driver", "weight", "impact", "correlation", "ranked"]):
            print(f"Row {idx}: {row.tolist()}")

# Search all sheets for importance/driver
print("\n--- ALL sheets: importance/driver/weight ---")
for sname, df in sheets.items():
    if sname == "Additonal Analysis":
        continue
    for idx, row in df.iterrows():
        row_str = " ".join(str(v) for v in row.values if pd.notna(v)).lower()
        if any(kw in row_str for kw in ["importance", "driver", "weight"]):
            print(f"[{sname}] Row {idx}: {row.tolist()}")

# ============================================================
# ASK 3: Mariposa prioritized x rep performance
# ============================================================
print("\n" + "="*80)
print("ASK 3: MARIPOSA DATA (rows 170-174 AA + keyword search)")
print("="*80)

if aa is not None:
    print("\n--- AA rows 165-180 ---")
    for idx in range(min(165, len(aa)), min(180, len(aa))):
        row = aa.iloc[idx]
        vals = [v for v in row.values if pd.notna(v)]
        if vals:
            print(f"Row {idx}: {row.tolist()}")

    # Search for Mariposa
    print("\n--- AA: searching for 'mariposa' ---")
    for idx in range(len(aa)):
        row = aa.iloc[idx]
        row_str = " ".join(str(v) for v in row.values if pd.notna(v)).lower()
        if "mariposa" in row_str:
            print(f"Row {idx}: {row.tolist()}")

# All sheets
print("\n--- ALL sheets: 'mariposa' ---")
for sname, df in sheets.items():
    for idx, row in df.iterrows():
        row_str = " ".join(str(v) for v in row.values if pd.notna(v)).lower()
        if "mariposa" in row_str:
            print(f"[{sname}] Row {idx}: {row.tolist()}")

# ============================================================
# ASK 4: Order of message recall (Q2_20Z)
# ============================================================
print("\n" + "="*80)
print("ASK 4: ORDER OF MESSAGE RECALL (Q2_20Z)")
print("="*80)

# Check TAG rows 238-317
if tag is not None:
    print(f"\nTAG shape: {tag.shape}")
    print("\n--- TAG rows 235-320 ---")
    for idx in range(min(235, len(tag)), min(320, len(tag))):
        row = tag.iloc[idx]
        vals = [v for v in row.values if pd.notna(v)]
        if vals:
            print(f"Row {idx}: {row.tolist()}")

# Search all sheets for Q2_20Z, order of recall, 1st recalled
print("\n--- ALL sheets: Q2_20Z / order / recall order ---")
for sname, df in sheets.items():
    for idx, row in df.iterrows():
        row_str = " ".join(str(v) for v in row.values if pd.notna(v)).lower()
        if any(kw in row_str for kw in ["q2_20z", "1st recalled", "2nd recalled", "3rd recalled",
                                          "first recalled", "order of recall", "recall order"]):
            print(f"[{sname}] Row {idx}: {row.tolist()}")

# Also check RYB for same range
if ryb is not None:
    print(f"\nRYB shape: {ryb.shape}")
    print("\n--- RYB rows 235-320 (if exist) ---")
    for idx in range(min(235, len(ryb)), min(320, len(ryb))):
        row = ryb.iloc[idx]
        vals = [v for v in row.values if pd.notna(v)]
        if vals:
            row_str = " ".join(str(v) for v in row.values if pd.notna(v)).lower()
            if any(kw in row_str for kw in ["q2_20", "recall", "order", "1st", "2nd", "3rd"]):
                print(f"Row {idx}: {row.tolist()}")

# ============================================================
# ASK 5: Monthly trend data
# ============================================================
print("\n" + "="*80)
print("ASK 5: MONTHLY TREND DATA")
print("="*80)

months = ["january", "february", "march", "april", "may", "june",
          "july", "august", "september", "october", "november", "december",
          "jan ", "feb ", "mar ", "apr ", "may ", "jun ",
          "jul ", "aug ", "sep ", "oct ", "nov ", "dec ",
          "monthly", "month", "trend"]

for sname, df in sheets.items():
    found = []
    for idx, row in df.iterrows():
        row_str = " ".join(str(v) for v in row.values if pd.notna(v)).lower()
        if any(kw in row_str for kw in months):
            found.append((idx, row.tolist()))
    if found:
        print(f"\n--- [{sname}] monthly/trend hits ({len(found)} rows) ---")
        for idx, vals in found[:30]:  # limit output
            print(f"Row {idx}: {vals}")

# ============================================================
# BONUS: Print AA column headers / structure overview
# ============================================================
print("\n" + "="*80)
print("BONUS: AA FULL STRUCTURE DUMP (first 3 cols, every row with data)")
print("="*80)
if aa is not None:
    print(f"AA total rows: {len(aa)}, cols: {aa.shape[1]}")
    # Print col 0 and col 1 for every non-empty row to see structure
    for idx in range(len(aa)):
        c0 = aa.iloc[idx, 0] if pd.notna(aa.iloc[idx, 0]) else ""
        c1 = aa.iloc[idx, 1] if aa.shape[1] > 1 and pd.notna(aa.iloc[idx, 1]) else ""
        c2 = aa.iloc[idx, 2] if aa.shape[1] > 2 and pd.notna(aa.iloc[idx, 2]) else ""
        if c0 or c1 or c2:
            print(f"Row {idx}: [{c0}] [{c1}] [{c2}]")

print("\n=== DONE ===")
