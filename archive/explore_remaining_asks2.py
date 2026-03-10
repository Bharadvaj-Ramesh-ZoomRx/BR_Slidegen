"""Targeted follow-up exploration for remaining asks."""
import pandas as pd
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(ROOT, "J_and_J_project", "data", "Lung SFEA SB.xlsx")

pd.set_option("display.max_columns", 30)
pd.set_option("display.width", 220)
pd.set_option("display.max_colwidth", 80)

xls = pd.ExcelFile(XLSX, engine="openpyxl")
aa = pd.read_excel(xls, sheet_name="Additonal Analysis", header=None)
ryb = pd.read_excel(xls, sheet_name="RYB", header=None)
tag = pd.read_excel(xls, sheet_name="TAG", header=None)

# ============================================================
# ASK 1: Closing Rates x Quality - Deep dive rows 150-170 AA
# ============================================================
print("="*80)
print("ASK 1: CLOSING RATES x QUALITY — AA rows 150-175")
print("="*80)
for idx in range(150, min(175, len(aa))):
    row = aa.iloc[idx]
    vals = [v for v in row.values if pd.notna(v)]
    if vals:
        print(f"Row {idx}: {row.tolist()}")

# ============================================================
# ASK 3: MARIPOSA — check AA rows 165-220
# ============================================================
print("\n" + "="*80)
print("ASK 3: MARIPOSA — AA rows 195-220")
print("="*80)
for idx in range(195, min(220, len(aa))):
    row = aa.iloc[idx]
    vals = [v for v in row.values if pd.notna(v)]
    if vals:
        print(f"Row {idx}: {row.tolist()}")

# Also search for MARIPOSA / clinical trial names
print("\n--- All sheets: MARIPOSA / FLAURA ---")
for sname in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sname, header=None)
    for idx in range(len(df)):
        row_str = " ".join(str(v) for v in df.iloc[idx].values if pd.notna(v)).lower()
        if "mariposa" in row_str or "flaura" in row_str:
            print(f"[{sname}] Row {idx}: {df.iloc[idx].tolist()}")

# ============================================================
# ASK 4: Order of recall — RYB rows 238-300
# ============================================================
print("\n" + "="*80)
print("ASK 4: ORDER OF RECALL — RYB rows 238-300")
print("="*80)
for idx in range(238, min(300, len(ryb))):
    row = ryb.iloc[idx]
    vals = [v for v in row.values if pd.notna(v)]
    if vals:
        print(f"Row {idx}: {row.tolist()}")

# ============================================================
# ASK 4: TAG header row for recall section
# ============================================================
print("\n" + "="*80)
print("ASK 4: TAG rows 238-242 (header area)")
print("="*80)
for idx in range(238, min(245, len(tag))):
    row = tag.iloc[idx]
    print(f"Row {idx}: {row.tolist()}")

# ============================================================
# ASK 5: Monthly — check Index sheet
# ============================================================
print("\n" + "="*80)
print("ASK 5: INDEX SHEET — full dump")
print("="*80)
idx_sheet = pd.read_excel(xls, sheet_name="Index", header=None)
print(f"Index shape: {idx_sheet.shape}")
for idx in range(len(idx_sheet)):
    row = idx_sheet.iloc[idx]
    vals = [v for v in row.values if pd.notna(v)]
    if vals:
        print(f"Row {idx}: {row.tolist()}")

# ============================================================
# Check for any wave/fieldwork date columns in RYB/TAG
# ============================================================
print("\n" + "="*80)
print("ASK 5: RYB column headers (row 0-3)")
print("="*80)
for idx in range(min(5, len(ryb))):
    print(f"RYB Row {idx}: {ryb.iloc[idx].tolist()}")

print("\nTAG column headers (row 0-3)")
for idx in range(min(5, len(tag))):
    print(f"TAG Row {idx}: {tag.iloc[idx].tolist()}")

print("\n=== DONE ===")
