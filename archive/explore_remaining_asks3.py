"""Targeted follow-up #3: Closing rates x Quality cross-tab, and order of recall in RYB."""
import pandas as pd
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(ROOT, "J_and_J_project", "data", "Lung SFEA SB.xlsx")

pd.set_option("display.max_columns", 35)
pd.set_option("display.width", 250)
pd.set_option("display.max_colwidth", 100)

xls = pd.ExcelFile(XLSX, engine="openpyxl")
aa = pd.read_excel(xls, sheet_name="Additonal Analysis", header=None)
ryb = pd.read_excel(xls, sheet_name="RYB", header=None)
tag = pd.read_excel(xls, sheet_name="TAG", header=None)

# ============================================================
# ASK 1: Q1_87Z in RYB and TAG (quality ratings)
# ============================================================
print("="*80)
print("ASK 1: Q1_87Z sections in RYB sheet")
print("="*80)
for idx in range(len(ryb)):
    c0 = str(ryb.iloc[idx, 0]) if pd.notna(ryb.iloc[idx, 0]) else ""
    if "Q1_87Z" in c0 or "Q1_87" in c0:
        # Print this row + next 15 rows for context
        for j in range(idx, min(idx + 20, len(ryb))):
            row = ryb.iloc[j]
            vals = [v for v in row.values if pd.notna(v)]
            if vals:
                print(f"Row {j}: {row.tolist()}")
        print("---")

print("\n" + "="*80)
print("ASK 1: Q1_87Z sections in TAG sheet")
print("="*80)
for idx in range(len(tag)):
    c0 = str(tag.iloc[idx, 0]) if pd.notna(tag.iloc[idx, 0]) else ""
    if "Q1_87Z" in c0 or "Q1_87" in c0:
        for j in range(idx, min(idx + 20, len(tag))):
            row = tag.iloc[j]
            vals = [v for v in row.values if pd.notna(v)]
            if vals:
                print(f"Row {j}: {row.tolist()}")
        print("---")

# ============================================================
# ASK 1: AA rows 55-75 for MARIPOSA and quality cross-tab
# ============================================================
print("\n" + "="*80)
print("ASK 1: AA rows 55-100 (MARIPOSA / CTA / quality cross)")
print("="*80)
for idx in range(55, min(100, len(aa))):
    row = aa.iloc[idx]
    vals = [v for v in row.values if pd.notna(v)]
    if vals:
        print(f"Row {idx}: {row.tolist()}")

# ============================================================
# ASK 1: AA rows 95-155 (look for quality x closing cross-tab)
# ============================================================
print("\n" + "="*80)
print("ASK 1: AA rows 95-155")
print("="*80)
for idx in range(95, min(155, len(aa))):
    row = aa.iloc[idx]
    vals = [v for v in row.values if pd.notna(v)]
    if vals:
        print(f"Row {idx}: {row.tolist()}")

# ============================================================
# ASK 4: RYB rows with '1st', '2nd', '3rd' in col 2/3
# ============================================================
print("\n" + "="*80)
print("ASK 4: RYB rows with ordinal recall positions")
print("="*80)
for idx in range(len(ryb)):
    row = ryb.iloc[idx]
    row_str = " ".join(str(v) for v in row.values if pd.notna(v))
    if any(kw in row_str for kw in ["1st", "2nd", "3rd", "4th or later"]):
        c2 = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
        c3 = str(row.iloc[3]) if pd.notna(row.iloc[3]) else ""
        if c2 in ["1st", "2nd", "3rd", "4th"] or c3 in ["1st", "2nd", "3rd", "4th or later"]:
            print(f"Row {idx}: {row.tolist()}")

# ============================================================
# ASK 4: TAG header for recall-order section
# ============================================================
print("\n" + "="*80)
print("ASK 4: TAG rows with ordinal recall positions")
print("="*80)
for idx in range(len(tag)):
    row = tag.iloc[idx]
    c2 = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
    c3 = str(row.iloc[3]) if pd.notna(row.iloc[3]) else ""
    if c2 in ["1st", "2nd", "3rd", "4th"] or c3 in ["1st", "2nd", "3rd", "4th or later"]:
        print(f"Row {idx}: {row.tolist()}")

# ============================================================
# ASK 5: Search for Q2_20Z in RYB/TAG
# ============================================================
print("\n" + "="*80)
print("ASK 5: Q2_20Z in RYB")
print("="*80)
for idx in range(len(ryb)):
    c0 = str(ryb.iloc[idx, 0]) if pd.notna(ryb.iloc[idx, 0]) else ""
    if "Q2_20" in c0:
        for j in range(idx, min(idx+5, len(ryb))):
            print(f"Row {j}: {ryb.iloc[j].tolist()}")
        print("---")

print("\nQ2_20Z in TAG")
for idx in range(len(tag)):
    c0 = str(tag.iloc[idx, 0]) if pd.notna(tag.iloc[idx, 0]) else ""
    if "Q2_20" in c0:
        for j in range(idx, min(idx+5, len(tag))):
            print(f"Row {j}: {tag.iloc[j].tolist()}")
        print("---")

print("\n=== DONE ===")
