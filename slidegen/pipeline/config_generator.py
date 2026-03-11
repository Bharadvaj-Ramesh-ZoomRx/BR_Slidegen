"""
config_generator.py — Data discovery and config scaffolding helpers.

These helpers are called by Claude Code during the "Create slides" workflow
to auto-discover Excel structure and generate a starter config.yaml.

Usage:
    from slidegen.pipeline.config_generator import discover_excel_structure
    summary = discover_excel_structure("projects/new_project/data/source_data.xlsx")
    # Claude reads the summary + reference docs → writes config.yaml
"""

from __future__ import annotations

import re
import pandas as pd
import yaml


# ── Question code detection ─────────────────────────────────────────────────

_CODE_PATTERN = re.compile(r'^[A-Z]\d')


def _detect_question_codes(df: pd.DataFrame, code_col: int = 0,
                           max_scan: int = 500) -> list[str]:
    """Scan a column for question code patterns (e.g. Q2_10Z, C1_81Z)."""
    codes = []
    for i in range(min(max_scan, len(df))):
        val = df.iloc[i, code_col]
        if pd.notna(val):
            s = str(val).strip()
            if _CODE_PATTERN.match(s) and len(s) >= 3:
                codes.append(s)
    return codes


def _detect_data_columns(df: pd.DataFrame, sample_rows: int = 100) -> list[dict]:
    """Identify columns that likely contain percentage data (0-1 range)."""
    results = []
    for col_idx in range(len(df.columns)):
        col_data = df.iloc[:sample_rows, col_idx]
        numeric = pd.to_numeric(col_data, errors='coerce').dropna()
        if len(numeric) < 5:
            continue
        mean_val = numeric.mean()
        # Percentages are typically 0-1 (decimal) or 0-100 (whole)
        if 0 < mean_val < 1.1 and numeric.min() >= 0:
            results.append({
                "index": col_idx,
                "mean": round(float(mean_val), 3),
                "non_null": int(len(numeric)),
                "format": "decimal_pct",
            })
        elif 0 < mean_val < 110 and numeric.min() >= 0:
            results.append({
                "index": col_idx,
                "mean": round(float(mean_val), 1),
                "non_null": int(len(numeric)),
                "format": "whole_pct",
            })
    return results


# ── Main discovery function ──────────────────────────────────────────────────

def discover_excel_structure(excel_path: str) -> dict:
    """Analyze an Excel file and return a structured summary.

    This gives Claude the information needed to auto-generate a config.yaml
    without reading the raw binary Excel file.

    Returns:
        {
            "file": "source_data.xlsx",
            "sheets": [
                {
                    "name": "RYB",
                    "rows": 303,
                    "cols": 21,
                    "column_headers": ["Q code", "Description", ...],
                    "question_codes": ["Q2_10Z", "C1_81Z", ...],
                    "data_columns": [{"index": 7, "mean": 0.45, ...}, ...],
                    "sample_rows": [[row0_col0, row0_col1, ...], ...]
                }
            ]
        }
    """
    import os

    xls = pd.ExcelFile(excel_path)
    result = {
        "file": os.path.basename(excel_path),
        "sheets": [],
    }

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)

        # Detect question codes in first column
        codes = _detect_question_codes(df, code_col=0)

        # Detect data columns
        data_cols = _detect_data_columns(df)

        # Get first row as potential headers
        first_row = []
        for j in range(min(25, len(df.columns))):
            val = df.iloc[0, j] if len(df) > 0 else None
            first_row.append(str(val) if pd.notna(val) else None)

        # Sample rows (first 5 data rows, first 10 columns)
        sample_rows = []
        for i in range(min(5, len(df))):
            row = []
            for j in range(min(10, len(df.columns))):
                val = df.iloc[i, j]
                if pd.notna(val):
                    row.append(str(val)[:50])
                else:
                    row.append(None)
            sample_rows.append(row)

        sheet_info = {
            "name": sheet_name,
            "rows": len(df),
            "cols": len(df.columns),
            "column_headers": first_row,
            "question_codes": codes[:50],  # cap at 50
            "question_code_count": len(codes),
            "data_columns": data_cols,
            "sample_rows": sample_rows,
        }
        result["sheets"].append(sheet_info)

    return result


# ── Config scaffold generator ────────────────────────────────────────────────

def generate_config_scaffold(
    excel_summary: dict,
    project_name: str,
    client: str,
    wave: str = "",
    period_current: str = "Q4'25",
    period_prior: str = "Q3'25",
    primary_brand: str = "Brand A",
    competitor_brand: str = "Brand B",
) -> str:
    """Generate a starter config.yaml from discovered Excel structure.

    Returns a YAML string with sheets and brands pre-filled.
    Claude should refine this with extraction and ask details.
    """
    sheets_config = {}
    for i, sheet in enumerate(excel_summary.get("sheets", [])):
        # Find likely prior/current columns from data_columns
        data_cols = sheet.get("data_columns", [])
        prior_col = data_cols[0]["index"] if len(data_cols) > 0 else 7
        current_col = data_cols[1]["index"] if len(data_cols) > 1 else prior_col + 1

        key = "primary" if i == 0 else ("competitor" if i == 1 else f"sheet_{i}")
        sheets_config[key] = {
            "name": sheet["name"],
            "q_prior_col": prior_col,
            "q_current_col": current_col,
        }

    project_block = {
        "name": project_name,
        "client": client,
        "period_current": period_current,
        "period_prior": period_prior,
    }
    if wave:
        project_block["wave"] = wave

    # Data/output use {{wave}} subfolders; template stays flat
    data_path = f"data/{{{{wave}}}}/source_data.xlsx" if wave else "data/source_data.xlsx"
    out_path = f"output/{{{{wave}}}}/deck.pptx" if wave else "output/deck.pptx"

    config = {
        "project": project_block,
        "data_source_path": data_path,
        "template_path": "templates/template.pptx",
        "output_path": out_path,
        "brands": {
            "primary": {
                "name": primary_brand,
                "full_name": primary_brand,
                "color_current": "#4472C4",
                "color_prior": "#A9C4EB",
            },
            "competitor": {
                "name": competitor_brand,
                "full_name": competitor_brand,
                "color_current": "#7030A0",
                "color_prior": "#AD88C8",
            },
        },
        "fonts": {
            "display": "Calibri",
            "body": "Calibri",
        },
        "sheets": sheets_config,
        "sample_sizes": {
            "primary": {"prior": 100, "current": 100},
            "competitor": {"prior": 100, "current": 100},
        },
        "extractions": [],
        "asks": [],
    }

    yaml_str = yaml.dump(config, default_flow_style=False, sort_keys=False,
                         allow_unicode=True)

    # Add comments for Claude to fill in
    yaml_str += """
# ── TODO: Add extractions ──────────────────────────────────────────────────
# Each extraction pulls a data block from the Excel into a named key.
# Example:
#   - id: brand_mr
#     method: question_code
#     sheet: primary
#     params: { code: "Q2_10Z", max_rows: 20 }
#
# Available methods: question_code, multi_question_code, row_range,
#   question_code_multi_col, nested_ordinal

# ── TODO: Add asks (slides) ───────────────────────────────────────────────
# Each ask defines one slide to generate.
# Example:
#   - id: brand_message_recall
#     slide_type: single_bar_with_delta
#     headline: "Brand message recall trends"
#     section: "MESSAGE RECALL"
#     source_text: "Source: Survey Data"
#     data_key: brand_mr
#     brand: primary
#     sort_by: current
#
# Available slide_types: cover, executive_summary, single_bar_with_delta,
#   dual_bar_with_delta, dual_bar_qoq, clustered_compare,
#   qoq_bar_with_delta, two_section_bar, stacked_order
"""
    return yaml_str
