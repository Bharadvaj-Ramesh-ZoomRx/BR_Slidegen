"""
discover_data.py
Reads all sheets from 'Lung SFEA SB.xlsx' and outputs a question-code mapping CSV.
Run this first to confirm column/code names before building charts.
"""
import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX_PATH = os.path.join(BASE_DIR, "Lung SFEA SB.xlsx")
OUT_DIR = BASE_DIR

def print_sheet_info(sheet_name, df, max_rows=5):
    print(f"\n{'='*60}")
    print(f"SHEET: {sheet_name}  |  Shape: {df.shape}")
    print(f"{'='*60}")
    print("COLUMNS:")
    for i, col in enumerate(df.columns):
        print(f"  [{i}] {repr(col)}")
    print(f"\nFIRST {max_rows} ROWS:")
    print(df.head(max_rows).to_string())

def main():
    xl = pd.ExcelFile(XLSX_PATH)
    sheet_names = xl.sheet_names
    print(f"Sheets found: {sheet_names}\n")

    all_mappings = []

    for sheet in sheet_names:
        df = pd.read_excel(XLSX_PATH, sheet_name=sheet, header=0)
        print_sheet_info(sheet, df)

        # Build a mapping row per column
        for col in df.columns:
            # Sample up to 3 non-null values
            vals = df[col].dropna().unique()[:3].tolist()
            all_mappings.append({
                "sheet": sheet,
                "column": str(col),
                "dtype": str(df[col].dtype),
                "non_null_count": int(df[col].notna().sum()),
                "sample_values": " | ".join(str(v) for v in vals),
            })

    mapping_df = pd.DataFrame(all_mappings)
    out_csv = os.path.join(OUT_DIR, "column_mapping.csv")
    mapping_df.to_csv(out_csv, index=False)
    print(f"\n\nMapping CSV written to: {out_csv}")

    # Also try to detect likely Q-code columns (question code pattern)
    print("\n\nQ-CODE COLUMNS DETECTED:")
    q_codes = mapping_df[mapping_df["column"].str.match(r"Q\d", na=False)]
    print(q_codes[["sheet","column","dtype","non_null_count","sample_values"]].to_string(index=False))

if __name__ == "__main__":
    main()
