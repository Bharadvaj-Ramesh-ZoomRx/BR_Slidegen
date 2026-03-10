"""Final targeted exploration: RYB recall order rows, AA MARIPOSA block, AA rows 173-195."""
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
# RYB: Find Q2_20Z section and recall order rows
# ============================================================
print("="*80)
print("RYB: Q2_20Z section")
print("="*80)
for idx in range(len(ryb)):
    c0 = str(ryb.iloc[idx, 0]) if pd.notna(ryb.iloc[idx, 0]) else ""
    if "Q2_20" in c0:
        for j in range(max(0, idx-2), min(idx+50, len(ryb))):
            row = ryb.iloc[j]
            vals = [v for v in row.values if pd.notna(v)]
            if vals:
                print(f"Row {j}: col0={row.iloc[0]}, col1={str(row.iloc[1])[:80]}, col2={row.iloc[2]}, col3={row.iloc[3]}, col7={row.iloc[7]}, col17={row.iloc[17]}")
        break  # just first occurrence

# ============================================================
# TAG: Find Q2_20Z section
# ============================================================
print("\n" + "="*80)
print("TAG: Q2_20Z section")
print("="*80)
for idx in range(len(tag)):
    c0 = str(tag.iloc[idx, 0]) if pd.notna(tag.iloc[idx, 0]) else ""
    if "Q2_20" in c0:
        for j in range(max(0, idx-2), min(idx+50, len(tag))):
            row = tag.iloc[j]
            vals = [v for v in row.values if pd.notna(v)]
            if vals:
                print(f"Row {j}: col0={row.iloc[0]}, col1={str(row.iloc[1])[:80]}, col2={row.iloc[2]}, col3={row.iloc[3]}, col7={row.iloc[7]}, col13={row.iloc[13]}")
        break

# ============================================================
# AA rows 55-80 (MARIPOSA discussed block)
# ============================================================
print("\n" + "="*80)
print("AA rows 55-85: MARIPOSA discussed / rep performance cross")
print("="*80)
for idx in range(55, min(85, len(aa))):
    row = aa.iloc[idx]
    vals = [v for v in row.values if pd.notna(v)]
    if vals:
        print(f"Row {idx}: {row.tolist()}")

# ============================================================
# AA rows 170-195: MARIPOSA x rep performance full block
# ============================================================
print("\n" + "="*80)
print("AA rows 169-195: MARIPOSA x rep perf full block")
print("="*80)
for idx in range(169, min(195, len(aa))):
    row = aa.iloc[idx]
    vals = [v for v in row.values if pd.notna(v)]
    if vals:
        print(f"Row {idx}: {row.tolist()}")

# ============================================================
# Check if there's wave/monthly data anywhere
# ============================================================
print("\n" + "="*80)
print("RYB: Check for wave columns (row 0 header)")
print("="*80)
print(f"RYB cols: {ryb.shape[1]}")
row0 = ryb.iloc[0]
for c in range(ryb.shape[1]):
    if pd.notna(row0.iloc[c]):
        print(f"  col {c}: {row0.iloc[c]}")

row1 = ryb.iloc[1]
for c in range(ryb.shape[1]):
    if pd.notna(row1.iloc[c]):
        print(f"  row1 col {c}: {row1.iloc[c]}")

# Check last columns of RYB for extra wave data
print("\nRYB row 0, last 10 cols:", ryb.iloc[0, -10:].tolist())
print("RYB row 1, last 10 cols:", ryb.iloc[1, -10:].tolist())

# Check if extra columns beyond col 27 have wave data
print("\nRYB cols 27-31 for row 4 (sample data row):", ryb.iloc[4, 27:].tolist() if ryb.shape[1] > 27 else "N/A")

print("\n=== DONE ===")
