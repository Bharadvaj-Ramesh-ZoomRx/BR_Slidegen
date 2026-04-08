"""
qual_data_loader.py — Qualitative (verbatim) data extraction from raw survey data.

Reads the raw data Excel (same format as raw_data_loader.py), identifies columns
containing free-text verbatim responses, extracts all responses with respondent
metadata (quarter, segment tags), and saves as qualitative_data.json.

This is Stage 0q in the pipeline — runs after Stage 0 (index_excel), before
Stage 1 (build-project-context). If no raw data file exists or no verbatim
columns are found, skips gracefully.

Theme-coding happens downstream at the skill level (Stage 3: /sfea-insight-writer),
not here. This module provides the structured extraction that skills consume.

Excel layout (per sheet, same as raw_data_loader):
    Row 3 (idx 2): Column headers (Id, User Id, Quarter, ...)
    Row 5 (idx 4): Question codes
    Row 6 (idx 5): Full question text
    Row 8 (idx 7): Question type + attribute label
    Row 9+ (idx 8+): Respondent data

Verbatim detection heuristics:
    1. Type string contains "Text", "Free Text", "Long Free", "Short Free", "Voice"
    2. OR values are predominantly long strings (>20 chars) with high uniqueness
    3. Exclude columns with mostly numeric, single-char, or "---" placeholder values
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
import re
from collections import defaultdict
from datetime import datetime
from typing import Optional

logger = logging.getLogger(__name__)

# Reuse constants from raw_data_loader
_ROW_HEADERS = 2
_ROW_CODES = 4
_ROW_QTEXT = 5
_ROW_QTYPES = 7
_ROW_DATA_START = 8
_COL_QUARTER = 6
_COL_ID = 0
_COL_META_END = 33

# ── Verbatim detection ────────────────────────────────────────────────────

# Type strings that indicate free-text / verbatim columns
_VERBATIM_TYPE_PATTERNS = re.compile(
    r"(?i)(free\s*text|long\s*free|short\s*free|voice|open.?end|verbatim|text\s*input)",
)

# Minimum response length (chars) to consider a value as genuine verbatim text
_MIN_VERBATIM_LEN = 15

# Minimum number of non-placeholder responses to keep a verbatim column
_MIN_RESPONSES = 3

# Placeholder values to exclude
_PLACEHOLDERS = {
    "---", "--", "-", "N/A", "n/a", "NA", "na", "None", "none", "",
    "VOICE_RESPONSE_ANSWERED", "voice_response_answered",
    "VOICE_RESPONSE_NOT_ANSWERED", "voice_response_not_answered",
    "NO_RESPONSE", "no_response", "SKIPPED", "skipped",
}


def _is_verbatim_by_type(type_str: str) -> bool:
    """Check if the column type string indicates a verbatim/free-text question."""
    return bool(_VERBATIM_TYPE_PATTERNS.search(type_str))


def _is_verbatim_by_values(values: list) -> bool:
    """Heuristic: detect verbatim columns from response content.

    A column is likely verbatim if:
    - >50% of non-null values are strings longer than MIN_VERBATIM_LEN
    - Values have high uniqueness (>70% unique among non-null)
    - Values are not predominantly numeric
    """
    if not values:
        return False

    # Filter out placeholders
    clean = [v for v in values if str(v).strip() not in _PLACEHOLDERS]
    if len(clean) < _MIN_RESPONSES:
        return False

    str_vals = [str(v) for v in clean]

    # Check: mostly long strings?
    long_count = sum(1 for s in str_vals if len(s) >= _MIN_VERBATIM_LEN)
    if long_count / len(str_vals) < 0.5:
        return False

    # Check: high uniqueness?
    unique_ratio = len(set(str_vals)) / len(str_vals) if str_vals else 0
    if unique_ratio < 0.5:
        return False

    # Check: not predominantly numeric
    numeric_count = 0
    for s in str_vals:
        try:
            float(s)
            numeric_count += 1
        except (ValueError, TypeError):
            pass
    if numeric_count / len(str_vals) > 0.5:
        return False

    # Check: not structured/serialized data (e.g. highlighter JSON ranges)
    json_like = sum(1 for s in str_vals if "startIndex" in s or "charLength" in s)
    if json_like / len(str_vals) > 0.3:
        return False

    return True


# ── Segment tagging ────────────────────────────────────────────────────

def _detect_segment_columns(columns: dict, code_map: dict,
                            respondents: list, config_segments: dict | None = None) -> dict:
    """Detect segment indicator columns for tagging verbatim responses.

    Returns dict mapping segment_name → {col_index, value_map} for tagging.
    Checks config-specified segment codes first, then falls back to auto-detection.
    """
    segments = {}

    if config_segments:
        for seg_name, seg_cfg in config_segments.items():
            code = seg_cfg.get("code", "")
            col_indices = code_map.get(code, [])
            if col_indices:
                ci = str(col_indices[0])
                # Discover unique values
                values = set()
                for r in respondents[:50]:
                    v = r.get("values", {}).get(ci)
                    if v is not None:
                        values.add(str(v).strip())
                segments[seg_name] = {"col": ci, "values": sorted(values)}

    return segments


# ── AI probe pair detection ──────────────────────────────────────────────

def _detect_ai_probe_pairs(columns: dict) -> list[dict]:
    """Detect AI-probing follow-up column pairs (question + response).

    The survey platform generates dynamic follow-up probes based on ratings.
    These appear as columns with codes like `*_generative_prompt_1_question`
    paired with `*_generative_prompt_1` (the response). We match explicitly
    on the generative_prompt naming convention to avoid false positives.
    """
    pairs = []
    col_indices = sorted(columns.keys(), key=lambda x: int(x))

    # Build a map of code → col_index for quick lookup
    code_to_col = {}
    for ci in col_indices:
        code = columns[ci].get("code", "")
        if code:
            code_to_col[code] = ci

    # Find _generative_prompt_N_question columns and pair with response columns
    _PROBE_RE = re.compile(r'(.+_generative_prompt_\d+)_question$', re.IGNORECASE)

    for ci in col_indices:
        code = columns[ci].get("code", "")
        m = _PROBE_RE.match(code)
        if m:
            response_code = m.group(1)  # e.g., C1_85F_1Z_generative_prompt_1
            if response_code in code_to_col:
                pairs.append({
                    "probe_col": ci,
                    "response_col": code_to_col[response_code],
                    "code": response_code,
                    "base_code": response_code.split("_generative_prompt_")[0],
                })

    return pairs


# ── Core extraction ───────────────────────────────────────────────────

def _extract_sheet_verbatims(wb, sheet_name: str,
                             config_segments: dict | None = None) -> dict:
    """Extract all verbatim responses from a single sheet.

    Returns:
        {
            "n_respondents": int,
            "n_verbatim_cols": int,
            "questions": {
                "Q1.53A": {
                    "code": "Q1_53AZ",
                    "text": "Full question text...",
                    "type": "verbatim",
                    "n_responses": int,
                    "responses": [
                        {
                            "id": respondent_id,
                            "quarter": "Q1'26",
                            "text": "The rep discussed...",
                            "segments": {"setting": "Community", "hii": "High Impact"}
                        },
                        ...
                    ]
                }
            },
            "ai_probe_pairs": [
                {
                    "code": "Q1_53Z",
                    "n_pairs": int,
                    "pairs": [
                        {"id": resp_id, "probe": "Tell me more...", "response": "Well..."}
                    ]
                }
            ]
        }
    """
    ws = wb[sheet_name]

    # Read header rows
    header_rows = []
    for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=_ROW_DATA_START, values_only=True)):
        header_rows.append(list(row))

    if len(header_rows) < _ROW_DATA_START:
        return {"n_respondents": 0, "n_verbatim_cols": 0, "questions": {}, "ai_probe_pairs": []}

    codes_row = header_rows[_ROW_CODES]
    qtext_row = header_rows[_ROW_QTEXT]
    qtypes_row = header_rows[_ROW_QTYPES]

    # Build column metadata
    columns = {}
    code_map = defaultdict(list)
    for ci in range(_COL_META_END, len(codes_row)):
        code = codes_row[ci]
        if not code:
            continue
        code = str(code).strip()

        type_str = str(qtypes_row[ci]) if ci < len(qtypes_row) and qtypes_row[ci] else ""
        parts = type_str.split("-", 3)
        if len(parts) >= 4 and parts[0].strip() == code:
            q_type = parts[1].strip()
            attribute = parts[3].strip()
        elif len(parts) >= 3 and any(t in parts[1] for t in ("Input", "choice", "Radio", "Array", "Text", "Yes")):
            q_type = parts[1].strip()
            attribute = ""
        else:
            q_type = ""
            attribute = type_str.strip()

        q_text = str(qtext_row[ci])[:300] if ci < len(qtext_row) and qtext_row[ci] else ""

        columns[ci] = {
            "code": code,
            "text": q_text,
            "type": q_type,
            "attribute": attribute,
            "type_str": type_str,
        }
        code_map[code].append(ci)

    # Read respondent data
    respondents = []
    for row in ws.iter_rows(min_row=_ROW_DATA_START + 1, values_only=True):
        vals = list(row)
        resp_id = vals[_COL_ID] if _COL_ID < len(vals) else None
        if resp_id is None:
            continue
        quarter = str(vals[_COL_QUARTER]).strip() if _COL_QUARTER < len(vals) and vals[_COL_QUARTER] else None

        respondents.append({
            "id": resp_id,
            "quarter": quarter,
            "values": vals,
        })

    n_respondents = len(respondents)

    # Detect segment columns for tagging
    # Build a mini code_map for segment detection (col indices as strings)
    str_code_map = {code: [str(ci) for ci in indices] for code, indices in code_map.items()}
    # Build respondent values dict keyed by string col indices for segment detection
    resp_for_segments = []
    for r in respondents[:50]:
        vals_dict = {}
        for ci_str in [str(ci) for ci in columns]:
            ci_int = int(ci_str)
            if ci_int < len(r["values"]) and r["values"][ci_int] is not None:
                vals_dict[ci_str] = r["values"][ci_int]
        resp_for_segments.append({"values": vals_dict})
    segments_def = _detect_segment_columns(columns, str_code_map, resp_for_segments, config_segments)

    # Phase 1: Identify verbatim columns
    verbatim_cols = {}  # col_index → True

    for ci, col_info in columns.items():
        # Check by type string
        if _is_verbatim_by_type(col_info.get("type", "")) or _is_verbatim_by_type(col_info.get("type_str", "")):
            verbatim_cols[ci] = True
            continue

        # Check by value content (sample first 50 respondents)
        sample_values = []
        for r in respondents[:min(50, len(respondents))]:
            v = r["values"][ci] if ci < len(r["values"]) else None
            if v is not None and str(v).strip() not in _PLACEHOLDERS:
                sample_values.append(v)

        if _is_verbatim_by_values(sample_values):
            verbatim_cols[ci] = True

    # Phase 2: Extract verbatim responses
    questions = {}

    for ci in sorted(verbatim_cols.keys()):
        col_info = columns[ci]
        code = col_info.get("code", f"col_{ci}")
        q_text = col_info.get("text", "")
        attribute = col_info.get("attribute", "")

        # Build a clean question label
        q_label = attribute if attribute else q_text[:100]

        responses = []
        for r in respondents:
            val = r["values"][ci] if ci < len(r["values"]) else None
            if val is None:
                continue
            text = str(val).strip()
            if text in _PLACEHOLDERS or len(text) < 3:
                continue

            # Build segment tags for this respondent
            seg_tags = {}
            for seg_name, seg_def in segments_def.items():
                seg_ci = int(seg_def["col"])
                if seg_ci < len(r["values"]) and r["values"][seg_ci] is not None:
                    seg_tags[seg_name] = str(r["values"][seg_ci]).strip()

            responses.append({
                "id": r["id"],
                "quarter": r["quarter"],
                "text": text,
                "segments": seg_tags,
            })

        if len(responses) < _MIN_RESPONSES:
            continue

        # Build a stable key: use question code + short attribute hash
        q_key = code
        if attribute:
            attr_short = re.sub(r'[^a-zA-Z0-9]', '_', attribute[:30]).strip('_')
            q_key = f"{code}__{attr_short}"

        questions[q_key] = {
            "code": code,
            "text": q_text,
            "attribute": attribute,
            "label": q_label,
            "col_index": ci,
            "detection_method": "type_string" if _is_verbatim_by_type(
                col_info.get("type_str", "")) else "value_heuristic",
            "n_responses": len(responses),
            "responses": responses,
        }

    # Phase 3: Detect AI probe pairs
    ai_probe_pairs = _detect_ai_probe_pairs(columns)
    ai_pairs_extracted = []

    for pair_def in ai_probe_pairs:
        probe_ci = int(pair_def["probe_col"])
        resp_ci = int(pair_def["response_col"])
        pairs = []

        for r in respondents:
            probe_val = r["values"][probe_ci] if probe_ci < len(r["values"]) else None
            resp_val = r["values"][resp_ci] if resp_ci < len(r["values"]) else None

            if probe_val is not None and resp_val is not None:
                probe_text = str(probe_val).strip()
                resp_text = str(resp_val).strip()
                if probe_text not in _PLACEHOLDERS and resp_text not in _PLACEHOLDERS:
                    if len(probe_text) > 5 and len(resp_text) > 5:
                        pairs.append({
                            "id": r["id"],
                            "quarter": r["quarter"],
                            "probe": probe_text,
                            "response": resp_text,
                        })

        if pairs:
            ai_pairs_extracted.append({
                "code": pair_def["code"],
                "probe_col": pair_def["probe_col"],
                "response_col": pair_def["response_col"],
                "n_pairs": len(pairs),
                "pairs": pairs,
            })

    return {
        "n_respondents": n_respondents,
        "n_verbatim_cols": len(questions),
        "questions": questions,
        "ai_probe_pairs": ai_pairs_extracted,
    }


# ── Hash & caching ─────────────────────────────────────────────────────

def _file_hash(path: str) -> str:
    """Fast MD5 hash for staleness checks."""
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()[:12]


# ── Public API ─────────────────────────────────────────────────────────

_DEFAULT_SHEET_MAP = {
    "primary": "RYB IM",
    "competitor": "TAG IM",
    "primary_npp": "RYB NPP",
    "competitor_npp": "TAG NPP",
    "market_survey": "MS",
}


def index_qualitative(
    raw_excel_path: str,
    json_path: str,
    sheet_map: dict | None = None,
    config_segments: dict | None = None,
) -> dict | None:
    """Extract verbatim responses from raw data Excel → qualitative_data.json.

    This is Stage 0q in the pipeline. If the raw Excel doesn't exist or contains
    no verbatim columns, returns None and does not create the JSON file.

    Uses hash-based caching: if qualitative_data.json already exists and the
    source Excel hasn't changed, returns the cached data.

    Args:
        raw_excel_path: Path to the raw data Excel (e.g., Additional Source (Raw Data).xlsx)
        json_path: Path to write qualitative_data.json
        sheet_map: Optional override for sheet name → role mapping
        config_segments: Optional segment column definitions from config
            e.g., {"setting": {"code": "Q_SETTING"}, "hii": {"code": "Q_HII"}}

    Returns:
        dict with qualitative data, or None if no raw data / no verbatims found
    """
    # Check raw file exists
    if not raw_excel_path or not os.path.exists(raw_excel_path):
        logger.info("  Stage 0q: No raw data file found — skipping qualitative extraction")
        return None

    # Check cache freshness
    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                existing = json.load(f)
            meta = existing.get("_meta", {})
            if meta.get("raw_hash") == _file_hash(raw_excel_path):
                n_sheets = len(existing.get("sheets", {}))
                n_qs = sum(
                    len(s.get("questions", {}))
                    for s in existing.get("sheets", {}).values()
                )
                logger.info(
                    "  Stage 0q: qualitative_data.json up to date — %d sheets, %d verbatim questions",
                    n_sheets, n_qs,
                )
                return existing
        except (json.JSONDecodeError, KeyError):
            logger.warning("  Stage 0q: Corrupt qualitative cache — re-extracting")

    # Parse raw Excel
    logger.info("  Stage 0q: Extracting verbatims from %s", os.path.basename(raw_excel_path))

    import openpyxl
    wb = openpyxl.load_workbook(raw_excel_path, read_only=True, data_only=True)

    # Resolve sheet map
    smap = dict(sheet_map or _DEFAULT_SHEET_MAP)
    # Auto-discover any sheets not in the default map
    for sn in wb.sheetnames:
        if sn not in smap.values():
            key = sn.lower().replace(" ", "_")
            if key not in smap:
                smap[key] = sn

    sheets_data = {}
    total_verbatim_cols = 0
    total_responses = 0

    for sheet_key, sheet_name in smap.items():
        if sheet_name not in wb.sheetnames:
            logger.debug("    Sheet '%s' not found — skipping", sheet_name)
            continue

        logger.info("    Scanning %s (%s)...", sheet_name, sheet_key)
        sheet_result = _extract_sheet_verbatims(wb, sheet_name, config_segments)

        if sheet_result["n_verbatim_cols"] > 0:
            sheets_data[sheet_key] = sheet_result
            total_verbatim_cols += sheet_result["n_verbatim_cols"]
            total_responses += sum(
                q["n_responses"] for q in sheet_result["questions"].values()
            )
            n_probes = sum(p["n_pairs"] for p in sheet_result.get("ai_probe_pairs", []))
            logger.info(
                "      → %d verbatim columns, %d total responses, %d AI probe pairs",
                sheet_result["n_verbatim_cols"],
                sum(q["n_responses"] for q in sheet_result["questions"].values()),
                n_probes,
            )
        else:
            logger.info("      → No verbatim columns detected")

    wb.close()

    if not sheets_data:
        logger.info("  Stage 0q: No verbatim data found in any sheet — skipping")
        return None

    # Build output
    result = {
        "_meta": {
            "extracted_at": datetime.now().isoformat(),
            "raw_file": os.path.basename(raw_excel_path),
            "raw_hash": _file_hash(raw_excel_path),
            "total_verbatim_columns": total_verbatim_cols,
            "total_responses": total_responses,
            "sheets_with_verbatims": list(sheets_data.keys()),
        },
        "sheets": sheets_data,
    }

    # Save JSON
    os.makedirs(os.path.dirname(json_path), exist_ok=True)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False, default=str)

    size_kb = os.path.getsize(json_path) / 1024
    logger.info(
        "  Stage 0q: Saved qualitative_data.json — %d sheets, %d verbatim columns, "
        "%d total responses (%.0f KB)",
        len(sheets_data), total_verbatim_cols, total_responses, size_kb,
    )

    return result


# ── Query helpers (used by downstream skills) ─────────────────────────

def get_verbatims_for_code(qual_data: dict, code: str,
                           sheet_key: str | None = None,
                           quarter: str | None = None,
                           segment_filter: dict | None = None,
                           max_responses: int = 0) -> list[dict]:
    """Retrieve verbatim responses for a specific question code.

    Searches across all sheets (or a specific sheet) for the given code.
    Optionally filters by quarter and segment values.

    Args:
        qual_data: The qualitative_data.json dict
        code: Question code to search for (partial match supported)
        sheet_key: Optional sheet to restrict search to
        quarter: Optional quarter label to filter by
        segment_filter: Optional dict of {segment_name: value} to filter by
        max_responses: Max responses to return (0 = all)

    Returns:
        list of response dicts with {id, quarter, text, segments, sheet, question_key}
    """
    results = []

    sheets = qual_data.get("sheets", {})
    if sheet_key:
        sheets = {sheet_key: sheets[sheet_key]} if sheet_key in sheets else {}

    for sk, sheet_data in sheets.items():
        for q_key, q_data in sheet_data.get("questions", {}).items():
            # Match code (exact or partial)
            q_code = q_data.get("code", "")
            if code not in q_code and q_code not in code:
                continue

            for resp in q_data.get("responses", []):
                # Filter by quarter
                if quarter and resp.get("quarter") != quarter:
                    continue

                # Filter by segments
                if segment_filter:
                    resp_segs = resp.get("segments", {})
                    if not all(resp_segs.get(k) == v for k, v in segment_filter.items()):
                        continue

                results.append({
                    **resp,
                    "sheet": sk,
                    "question_key": q_key,
                    "question_text": q_data.get("text", ""),
                    "question_label": q_data.get("label", ""),
                })

    if max_responses > 0:
        results = results[:max_responses]

    return results


def summarize_qualitative(qual_data: dict) -> dict:
    """Generate a summary of the qualitative data for Stage 1/2 context files.

    Returns a compact summary suitable for inclusion in project_context.md
    or hypothesis_bank.md.
    """
    if not qual_data:
        return {"available": False}

    meta = qual_data.get("_meta", {})
    sheets = qual_data.get("sheets", {})

    summary = {
        "available": True,
        "total_verbatim_columns": meta.get("total_verbatim_columns", 0),
        "total_responses": meta.get("total_responses", 0),
        "sheets": {},
    }

    for sk, sheet_data in sheets.items():
        questions_summary = []
        for q_key, q_data in sheet_data.get("questions", {}).items():
            questions_summary.append({
                "code": q_data.get("code", ""),
                "label": q_data.get("label", ""),
                "n_responses": q_data.get("n_responses", 0),
            })
        n_probes = sum(p["n_pairs"] for p in sheet_data.get("ai_probe_pairs", []))
        summary["sheets"][sk] = {
            "n_respondents": sheet_data.get("n_respondents", 0),
            "n_verbatim_questions": len(questions_summary),
            "n_ai_probe_pairs": n_probes,
            "questions": questions_summary,
        }

    return summary
