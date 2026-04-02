"""
config_generator.py — Data discovery and config scaffolding helpers.

These helpers are called by Claude Code during the "Create slides" workflow
to auto-discover Excel structure and generate a starter config.yaml.

Usage:
    from slidegen.pipeline.config_generator import discover_excel_structure, scaffold_config_from_plan
    summary = discover_excel_structure("projects/new_project/data/source_data.xlsx")
    # Claude reads the summary + reference docs → writes config.yaml

    # Or auto-scaffold from slide plan:
    yaml_str = scaffold_config_from_plan(
        "projects/jnj_rybrevant/context/PET_Q3Q4_2025/slide_plan.md",
        "projects/jnj_rybrevant/context/PET_Q3Q4_2025/source_data.json",
    )
"""

from __future__ import annotations

import json as _json
import logging
import os
import re
from typing import Optional

import pandas as pd
import yaml

logger = logging.getLogger(__name__)


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


# ── Slide plan → config scaffold ─────────────────────────────────────────────

# Regex patterns for parsing slide_plan.md
_SLIDE_HEADING = re.compile(r'^###\s+Slide\s+(\d+)\s*[—–-]\s*(.+)', re.MULTILINE)
_SECTION_HEADING = re.compile(r'^##\s+SECTION:\s*(.+)', re.MULTILINE)
_FIELD_PATTERN = re.compile(r'^\*\*([^*]+)\*\*:?\s*(.+)', re.MULTILINE)
_Q_CODE_LINE = re.compile(r'^\s*-\s+([A-Z]\d[\w_]*)\s*[—–-]\s*"?(.+?)"?\s*$', re.MULTILINE)
_BACKTICK_VALUE = re.compile(r'`([^`]+)`')

# Import valid slide types from the single source of truth (RENDERERS registry)
from slidegen.pipeline.slide_renderers import RENDERERS as _RENDERERS
_VALID_SLIDE_TYPES = set(_RENDERERS.keys())

# Slide types that don't need data extractions
_NO_DATA_TYPES = {"cover", "executive_summary"}

def scaffold_config_from_plan(
    plan_path: str,
    json_path: str,
    base_config_path: Optional[str] = None,
) -> str:
    """Auto-generate config.yaml extractions + asks from a slide_plan.md.

    Parses the slide plan for slide specs and question codes, cross-checks
    codes against the _codes index in source_data.json, auto-selects
    extraction method and pct_mode, and returns a YAML string with
    extractions and asks ready for review.

    Args:
        plan_path:        Path to slide_plan.md.
        json_path:        Path to source_data.json (must have _codes index).
        base_config_path: Optional existing config.yaml to merge into.
                          If provided, project/brands/sheets/etc. are preserved
                          and only extractions + asks are replaced.

    Returns:
        YAML string with extractions and asks populated from the slide plan.
        Includes comments flagging any question codes not found in _codes.
    """
    # Read slide plan
    with open(plan_path, "r", encoding="utf-8") as f:
        plan_text = f.read()

    # Read _codes index from source_data.json
    codes_index = {}
    sheets_in_json = {}
    if json_path and os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                source_data = _json.load(f)
        except _json.JSONDecodeError as e:
            logger.warning("Corrupt %s — skipping code index: %s", json_path, e)
            source_data = {}
        codes_index = source_data.get("_codes", {})
        sheets_in_json = source_data.get("_sheets", {})

    # Parse slides from the plan
    slides = _parse_slide_plan(plan_text)

    # Build extractions and asks
    extractions = []
    asks = []
    extraction_ids = set()
    warnings = []

    for slide in slides:
        slide_type = slide["slide_type"]

        # Cover and ES don't need extractions
        if slide_type in _NO_DATA_TYPES:
            ask = _build_ask_entry(slide)
            asks.append(ask)
            continue

        # For data slides, create extraction(s) from question codes
        q_codes = slide.get("question_codes", [])
        if not q_codes:
            warnings.append(
                f"Slide {slide['number']} ({slide['title']}): "
                f"no question codes found — extraction must be added manually"
            )
            ask = _build_ask_entry(slide, data_key="")
            asks.append(ask)
            continue

        # Determine extraction strategy based on slide_type and code count
        ext_entries, data_key, extra, slide_warnings = _build_extractions(
            slide, q_codes, codes_index, extraction_ids
        )
        extractions.extend(ext_entries)
        warnings.extend(slide_warnings)

        ask = _build_ask_entry(slide, data_key=data_key, extra=extra)
        asks.append(ask)

    # Build output YAML
    if base_config_path and os.path.exists(base_config_path):
        yaml_str = _merge_into_base_config(base_config_path, extractions, asks)
    else:
        yaml_str = _standalone_scaffold(extractions, asks)

    # Append warnings as comments
    if warnings:
        yaml_str += "\n# ── Warnings (review required) ────────────────────────────────────────\n"
        for w in warnings:
            yaml_str += f"# ⚠ {w}\n"

    return yaml_str


def _parse_slide_plan(text: str) -> list[dict]:
    """Parse slide_plan.md into a list of slide dicts.

    Each dict has: number, title, slide_type, section, question_codes,
    insight, driving_question, cuts_needed, chart_description.
    """
    slides = []
    current_section = "Opening"

    # Split into slide blocks using ### Slide N headings
    # Find all slide heading positions
    heading_matches = list(_SLIDE_HEADING.finditer(text))
    section_matches = list(_SECTION_HEADING.finditer(text))

    # Build a position → section mapping
    section_at = []
    for m in section_matches:
        section_at.append((m.start(), m.group(1).strip()))

    for i, match in enumerate(heading_matches):
        slide_num = int(match.group(1))
        slide_title = match.group(2).strip()

        # Determine section for this slide
        for pos, sec_name in reversed(section_at):
            if pos < match.start():
                current_section = sec_name
                break

        # Extract block text (from this heading to the next heading or end)
        block_start = match.end()
        block_end = heading_matches[i + 1].start() if i + 1 < len(heading_matches) else len(text)
        block = text[block_start:block_end]

        # Parse fields from block
        slide = {
            "number": slide_num,
            "title": slide_title,
            "section": current_section,
            "slide_type": "single_bar_with_delta",  # default
            "question_codes": [],
            "insight": "",
            "driving_question": "",
            "cuts_needed": "",
            "chart_description": "",
            "brand": "primary",
        }

        # Extract **field:** value pairs
        for fm in _FIELD_PATTERN.finditer(block):
            field_name = fm.group(1).strip().rstrip(':').strip().lower()
            field_value = fm.group(2).strip()

            if field_name == "slide_type":
                # Extract from backticks if present
                bt = _BACKTICK_VALUE.search(field_value)
                slide["slide_type"] = bt.group(1) if bt else field_value
            elif field_name == "driving question":
                slide["driving_question"] = field_value
            elif field_name == "insight":
                slide["insight"] = field_value
            elif field_name == "cuts needed":
                slide["cuts_needed"] = field_value
            elif field_name == "chart description":
                slide["chart_description"] = field_value
            elif field_name == "content":
                # For cover/ES slides
                slide["content"] = field_value

        # Extract question codes from **Primary questions:** section
        # Find the "Primary questions:" field and parse the bullet list after it
        pq_match = re.search(r'\*\*Primary questions:\*\*', block)
        if pq_match:
            pq_block = block[pq_match.end():]
            # Take lines until next **field:** or end of block
            pq_end = re.search(r'^\*\*[^*]+\*\*:', pq_block, re.MULTILINE)
            pq_text = pq_block[:pq_end.start()] if pq_end else pq_block
            for qm in _Q_CODE_LINE.finditer(pq_text):
                code = qm.group(1).strip()
                q_text = qm.group(2).strip().rstrip('"')
                slide["question_codes"].append({"code": code, "text": q_text})

        # Detect brand from chart description, cuts, or title
        block_lower = block.lower()
        if "competitor" in block_lower or "tag" in block_lower:
            if "brand comparison" in block_lower or "vs" in slide_title.lower():
                slide["brand"] = "primary"  # comparison slides use primary
            elif "competitor" in slide.get("chart_description", "").lower():
                slide["brand"] = "competitor"

        slides.append(slide)

    return slides


def _build_extractions(
    slide: dict,
    q_codes: list[dict],
    codes_index: dict,
    seen_ids: set,
) -> tuple[list[dict], str, dict, list[str]]:
    """Build extraction entries for a slide's question codes.

    Returns: (extraction_list, primary_data_key, extra_dict, warnings)
    """
    extractions = []
    warnings = []
    slide_type = slide["slide_type"]

    # Look up each code in _codes to determine method and sheet
    code_lookups = []
    used_sheets = set()
    for qc in q_codes:
        code = qc["code"]
        q_text = qc.get("text", "")
        # Check for parenthetical brand hint like "(TAG)" or "(competitor)"
        brand_hint = _detect_brand_hint(q_text)
        found_sheet, found_meta = _lookup_code(
            code, codes_index, prefer_sheet=brand_hint, exclude_sheets=used_sheets
        )
        if found_sheet:
            used_sheets.add(found_sheet)
        code_lookups.append({
            "code": code,
            "text": q_text,
            "sheet": found_sheet,
            "meta": found_meta,
        })
        if not found_sheet:
            warnings.append(
                f"Slide {slide['number']}: code '{code}' not found in _codes index"
            )

    # Generate extraction ID from slide title
    base_id = _slugify(slide["title"])

    # Single-code slides: one extraction
    if len(code_lookups) == 1:
        cl = code_lookups[0]
        ext_id = _unique_id(base_id, seen_ids)
        ext = _make_extraction(ext_id, cl)
        extractions.append(ext)
        return extractions, ext_id, {}, warnings

    # Multi-code slides: strategy depends on slide_type
    extra = {}

    if slide_type in ("dual_bar_with_delta", "dual_bar_qoq", "dual_bar_compare"):
        # Two codes → left/right or primary/comp extractions
        if len(code_lookups) >= 2:
            left_id = _unique_id(base_id + "_left", seen_ids)
            right_id = _unique_id(base_id + "_right", seen_ids)
            extractions.append(_make_extraction(left_id, code_lookups[0]))
            extractions.append(_make_extraction(right_id, code_lookups[1]))

            if slide_type == "dual_bar_with_delta":
                extra = {
                    "left": {"data_key": left_id},
                    "right": {"data_key": right_id},
                }
                return extractions, left_id, extra, warnings
            else:
                extra = {"primary_key": left_id, "comp_key": right_id}
                return extractions, left_id, extra, warnings

    if slide_type == "clustered_compare":
        # Two brands, same code → two extractions (primary + comp sheet)
        if len(code_lookups) >= 1:
            code = code_lookups[0]["code"]
            primary_id = _unique_id(base_id + "_pri", seen_ids)
            comp_id = _unique_id(base_id + "_comp", seen_ids)

            # Primary sheet extraction
            ext_pri = _make_extraction(primary_id, code_lookups[0])
            extractions.append(ext_pri)

            # Competitor sheet extraction (same code, different sheet)
            if len(code_lookups) >= 2:
                ext_comp = _make_extraction(comp_id, code_lookups[1])
            else:
                # Same code on competitor sheet
                ext_comp = _make_extraction(comp_id, code_lookups[0])
                ext_comp["sheet"] = "competitor"
            extractions.append(ext_comp)

            extra = {"primary_key": primary_id, "comp_key": comp_id}
            return extractions, primary_id, extra, warnings

    # Default: first code is primary extraction, rest noted in extra
    primary_id = _unique_id(base_id, seen_ids)
    extractions.append(_make_extraction(primary_id, code_lookups[0]))

    if len(code_lookups) > 1:
        for idx, cl in enumerate(code_lookups[1:], 1):
            extra_id = _unique_id(f"{base_id}_{idx}", seen_ids)
            extractions.append(_make_extraction(extra_id, cl))

    return extractions, primary_id, extra, warnings


def _make_extraction(ext_id: str, code_lookup: dict) -> dict:
    """Create a single extraction dict from a code lookup result."""
    code = code_lookup["code"]
    sheet = code_lookup.get("sheet", "primary")
    meta = code_lookup.get("meta") or {}

    method = _infer_method(meta)
    pct_mode = _infer_pct_mode(meta)

    params = {"code": code}
    if method == "multi_question_code":
        params = {"codes": [code]}
    if pct_mode != "pct":
        params["pct_mode"] = pct_mode
    if meta.get("sub_row_count", 0) > 0:
        params["max_rows"] = min(meta["sub_row_count"] + 2, 30)

    ext = {
        "id": ext_id,
        "method": method,
        "sheet": sheet or "primary",
        "params": params,
    }
    return ext


def _infer_method(meta: dict) -> str:
    """Infer extraction method from _codes metadata."""
    if not meta:
        return "question_code"  # safe default

    if meta.get("has_sub_codes", False):
        return "multi_question_code"
    if meta.get("sub_row_count", 0) > 1:
        return "question_code"
    return "question_code"


def _infer_pct_mode(meta: dict) -> str:
    """Infer pct_mode from _codes value_range metadata."""
    if not meta:
        return "pct"  # default: decimal → percentage

    value_range = meta.get("value_range", "decimal")
    return "straight" if value_range == "whole" else "pct"


def _detect_brand_hint(q_text: str) -> Optional[str]:
    """Detect a parenthetical brand/sheet hint in question text.

    E.g., 'Unaided Message Recall (TAG)' → 'TAG'
    """
    m = re.search(r'\(([A-Z]{2,})\)\s*$', q_text)
    return m.group(1) if m else None


def _lookup_code(
    code: str,
    codes_index: dict,
    prefer_sheet: Optional[str] = None,
    exclude_sheets: Optional[set] = None,
) -> tuple[Optional[str], Optional[dict]]:
    """Look up a question code across all sheets in _codes index.

    Args:
        code: Question code to find.
        codes_index: {sheet_name: {code: metadata}}.
        prefer_sheet: If set, prefer this sheet name (case-insensitive partial match).
        exclude_sheets: If set, try to find the code on a sheet NOT in this set
                        (for getting different-sheet lookups on duplicate codes).

    Returns (sheet_key, code_metadata) or (None, None) if not found.
    """
    matches = []
    for sheet_name, sheet_codes in codes_index.items():
        if code in sheet_codes:
            matches.append((sheet_name, sheet_codes[code]))

    if not matches:
        return None, None

    # If prefer_sheet is specified, try to match it
    if prefer_sheet:
        for sheet_name, meta in matches:
            if prefer_sheet.lower() in sheet_name.lower():
                return sheet_name, meta

    # If exclude_sheets is specified, prefer a sheet not yet used
    if exclude_sheets:
        for sheet_name, meta in matches:
            if sheet_name not in exclude_sheets:
                return sheet_name, meta

    # Fallback: first match
    return matches[0]


def _build_ask_entry(slide: dict, data_key: str = "", extra: Optional[dict] = None) -> dict:
    """Build an ask dict from parsed slide data."""
    ask = {
        "id": _slugify(slide["title"]),
        "slide_type": slide["slide_type"],
        "headline": slide.get("insight") or slide["title"],
        "section": slide["section"].upper(),
        "source_text": "",  # needs manual entry
        "data_key": data_key,
    }

    if slide.get("brand") and slide["brand"] != "primary":
        ask["brand"] = slide["brand"]

    # Add sort_by for bar-type slides
    if slide["slide_type"] in (
        "single_bar_with_delta", "dual_bar_with_delta",
        "qoq_bar_with_delta", "abacus",
    ):
        ask["sort_by"] = "current"
        ask["sort_desc"] = True

    if extra:
        ask["extra"] = extra

    return ask


def _slugify(text: str) -> str:
    """Convert a slide title to a snake_case id."""
    # Remove special chars, lowercase, replace spaces with underscores
    slug = re.sub(r'[^a-zA-Z0-9\s]', '', text.lower())
    slug = re.sub(r'\s+', '_', slug.strip())
    # Truncate to reasonable length
    return slug[:40].rstrip('_')


def _unique_id(base: str, seen: set) -> str:
    """Ensure extraction ID is unique by appending a suffix if needed."""
    if base not in seen:
        seen.add(base)
        return base
    for i in range(2, 100):
        candidate = f"{base}_{i}"
        if candidate not in seen:
            seen.add(candidate)
            return candidate
    return base  # fallback


def _merge_into_base_config(
    base_config_path: str, extractions: list[dict], asks: list[dict]
) -> str:
    """Load an existing config.yaml and replace extractions + asks."""
    try:
        with open(base_config_path, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
    except yaml.YAMLError as e:
        raise ValueError(f"Malformed YAML in {base_config_path}: {e}") from e

    config["extractions"] = extractions
    config["asks"] = asks

    return yaml.dump(config, default_flow_style=False, sort_keys=False,
                     allow_unicode=True)


def _standalone_scaffold(extractions: list[dict], asks: list[dict]) -> str:
    """Generate a minimal YAML with just extractions and asks."""
    content = {
        "extractions": extractions,
        "asks": asks,
    }
    yaml_str = yaml.dump(content, default_flow_style=False, sort_keys=False,
                         allow_unicode=True)
    header = (
        "# ── Auto-scaffolded from slide_plan.md ──────────────────────────────────\n"
        "# Review and adjust: source_text, sort_by, extra params, sheet mappings.\n"
        "# Merge these extractions + asks into your project config.yaml.\n\n"
    )
    return header + yaml_str
