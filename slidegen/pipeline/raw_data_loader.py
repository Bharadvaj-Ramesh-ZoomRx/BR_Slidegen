"""
raw_data_loader.py — Respondent-level raw data loader and aggregator.

Reads source_raw_data.xlsx (one row per HCP response) and computes
aggregated metrics (top-2-box %, yes %, recall %, mean) on the fly.

Excel layout (per sheet):
    Row 3: Column headers (Id, User Id, Quarter, ...)
    Row 5: Question codes (Q1_87Z, C1_81Z, Q2_10Z, ...)
    Row 6: Full question text
    Row 8: Question type + attribute label (e.g. "Q1_87Z-Multiple Numerical Input-N-Overall quality...")
    Row 9+: Respondent data

Sheets: RYB IM, TAG IM, RYB NPP, TAG NPP, MS
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
from collections import defaultdict
from datetime import datetime
from typing import Optional, Callable

logger = logging.getLogger(__name__)

# ── Header row indices (0-based within the Excel, 1-based in openpyxl) ────
_ROW_HEADERS = 2      # row 3: Id, User Id, Quarter, ...
_ROW_CODES = 4        # row 5: question codes
_ROW_QTEXT = 5        # row 6: full question text
_ROW_QTYPES = 7       # row 8: code-type-N-attribute_label
_ROW_DATA_START = 8   # row 9+: respondent data
_COL_QUARTER = 6      # column G: quarter label ("Q4'25", "Q1'26")
_COL_ID = 0           # column A: respondent ID
_COL_META_END = 33    # columns 0-32 are metadata; 33+ are question responses


# ── Aggregation functions ──────────────────────────────────────────────────

def _top2box(values: list) -> Optional[float]:
    """% of respondents rating 6 or 7 on a 1-7 scale."""
    nums = [v for v in values if isinstance(v, (int, float))]
    if not nums:
        return None
    return round(sum(1 for v in nums if v >= 6) / len(nums) * 100, 1)


def _yes_pct(values: list) -> Optional[float]:
    """% responding 'Yes' (for Yes/No questions)."""
    valid = [v for v in values if v is not None]
    if not valid:
        return None
    yes_count = sum(1 for v in valid if str(v).lower() in ("yes", "y", "1"))
    return round(yes_count / len(valid) * 100, 1)


def _recall_pct(values: list) -> Optional[float]:
    """% with value=1 (for multi-choice 0/1 binary questions)."""
    nums = [v for v in values if isinstance(v, (int, float))]
    if not nums:
        return None
    return round(sum(1 for v in nums if v == 1) / len(nums) * 100, 1)


def _mean_val(values: list) -> Optional[float]:
    """Mean of numerical values."""
    nums = [v for v in values if isinstance(v, (int, float))]
    if not nums:
        return None
    return round(sum(nums) / len(nums), 1)


_AGG_FUNCS = {
    "top2box": _top2box,
    "yes_pct": _yes_pct,
    "recall_pct": _recall_pct,
    "mean": _mean_val,
}


# ── Sheet name mapping ────────────────────────────────────────────────────

_DEFAULT_SHEET_MAP = {
    "primary": "RYB IM",
    "competitor": "TAG IM",
    "primary_npp": "RYB NPP",
    "competitor_npp": "TAG NPP",
    "market_share": "MS",
}


# ── Raw data parsing ──────────────────────────────────────────────────────

def _parse_raw_sheet(wb, sheet_name: str) -> dict:
    """Parse a single raw data sheet into a structured dict.

    Returns:
        {
            "columns": {
                col_index: {
                    "code": "Q1_87Z",
                    "text": "Full question text...",
                    "type": "Multiple Numerical Input",
                    "attribute": "Overall quality of sales call",
                }
            },
            "code_map": {
                "Q1_87Z": [col_index, col_index, ...],   # all columns for this code
            },
            "respondents": [
                {"id": 123, "quarter": "Q4'25", "values": {col_idx: value, ...}},
                ...
            ]
        }
    """
    ws = wb[sheet_name]

    # Read header rows
    header_rows = []
    for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=_ROW_DATA_START, values_only=True)):
        header_rows.append(list(row))

    if len(header_rows) < _ROW_DATA_START:
        return {"columns": {}, "code_map": {}, "respondents": []}

    codes_row = header_rows[_ROW_CODES]
    qtext_row = header_rows[_ROW_QTEXT]
    qtypes_row = header_rows[_ROW_QTYPES]

    # Build column map
    columns = {}
    code_map = defaultdict(list)
    for ci in range(_COL_META_END, len(codes_row)):
        code = codes_row[ci]
        if not code:
            continue
        code = str(code).strip()

        # Parse type string — two formats:
        #   RYB: "Q1_87Z-Multiple Numerical Input-N-attribute label"  (4+ parts)
        #   TAG: "attribute label"  (no prefix, the attribute IS the full string)
        type_str = str(qtypes_row[ci]) if ci < len(qtypes_row) and qtypes_row[ci] else ""
        parts = type_str.split("-", 3)
        if len(parts) >= 4 and parts[0].strip() == code:
            # Standard format: code-type-N-attribute
            q_type = parts[1].strip()
            attribute = parts[3].strip()
        elif len(parts) >= 3 and any(t in parts[1] for t in ("Input", "choice", "Radio", "Array", "Text", "Yes")):
            # Format: code-type-N (no attribute suffix) — attribute is empty
            q_type = parts[1].strip()
            attribute = ""
        else:
            # TAG-style: entire string is the attribute label
            q_type = ""
            attribute = type_str.strip()

        q_text = str(qtext_row[ci])[:200] if ci < len(qtext_row) and qtext_row[ci] else ""

        columns[ci] = {
            "code": code,
            "text": q_text,
            "type": q_type,
            "attribute": attribute,
        }
        code_map[code].append(ci)

    # Read respondent data
    respondents = []
    for row in ws.iter_rows(min_row=_ROW_DATA_START + 1, values_only=True):
        vals = list(row)
        resp_id = vals[_COL_ID] if _COL_ID < len(vals) else None
        if resp_id is None:
            continue
        quarter = vals[_COL_QUARTER] if _COL_QUARTER < len(vals) else None

        # Collect question response values (only non-None)
        resp_values = {}
        for ci in columns:
            if ci < len(vals) and vals[ci] is not None:
                resp_values[ci] = vals[ci]

        respondents.append({
            "id": resp_id,
            "quarter": quarter,
            "values": resp_values,
        })

    return {
        "columns": columns,
        "code_map": dict(code_map),
        "respondents": respondents,
    }


# ── Caching ─────────────────────────────────────────────────────────────

def _file_hash(path: str) -> str:
    """Fast MD5 hash for staleness checks."""
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()[:12]


def _raw_json_path(config) -> str:
    """Path to source_raw_data.json cache."""
    if config.context_path:
        return os.path.join(config.context_path, "source_raw_data.json")
    return os.path.join(
        os.path.dirname(getattr(config, "raw_data_source_path", "")),
        "source_raw_data.json",
    )


def load_raw_data(config) -> dict | None:
    """Load and cache parsed raw data from source_raw_data.xlsx.

    Returns dict mapping sheet_key → parsed sheet data, or None if no raw data file.
    """
    raw_path = getattr(config, "raw_data_source_path", "")
    if not raw_path or not os.path.exists(raw_path):
        return None

    json_path = _raw_json_path(config)

    # Check cache
    if os.path.exists(json_path):
        with open(json_path, "r", encoding="utf-8") as f:
            cached = json.load(f)
        meta = cached.get("_meta", {})
        if meta.get("raw_hash") == _file_hash(raw_path):
            n_sheets = sum(1 for k in cached if not k.startswith("_"))
            print(f"  Raw data loaded from cache: {json_path} ({n_sheets} sheets)")
            return cached

    # Parse from Excel
    print(f"  Parsing raw data: {os.path.basename(raw_path)}")
    import openpyxl
    wb = openpyxl.load_workbook(raw_path, read_only=True, data_only=True)

    # Build sheet map: use config's raw_sheets if defined, else auto-discover + defaults
    sheet_map = dict(_DEFAULT_SHEET_MAP)
    raw_sheets_cfg = getattr(config, "raw_sheets", None)
    if raw_sheets_cfg and isinstance(raw_sheets_cfg, dict):
        sheet_map.update(raw_sheets_cfg)
    else:
        # Auto-discover sheets not in default map
        for sn in wb.sheetnames:
            if sn not in sheet_map.values():
                key = sn.lower().replace(" ", "_")
                sheet_map[key] = sn

    raw_data = {}
    for sheet_key, sheet_name in sheet_map.items():
        if sheet_name in wb.sheetnames:
            parsed = _parse_raw_sheet(wb, sheet_name)
            raw_data[sheet_key] = parsed
            n_resp = len(parsed["respondents"])
            n_codes = len(parsed["code_map"])
            print(f"    {sheet_name} -> {sheet_key}: {n_resp} respondents, {n_codes} question codes")

    wb.close()

    # Save cache
    raw_data["_meta"] = {
        "parsed_at": datetime.now().isoformat(),
        "raw_file": os.path.basename(raw_path),
        "raw_hash": _file_hash(raw_path),
    }

    os.makedirs(os.path.dirname(json_path), exist_ok=True)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(raw_data, f, indent=2, default=str)
    print(f"  Saved raw data cache: {json_path}")

    return raw_data


# ── Quarter discovery & matching ──────────────────────────────────────

def discover_quarters(raw_data: dict, sheet_key: str = "primary") -> list[str]:
    """Return sorted list of quarter labels found in the data.

    Useful for auto-detecting available periods without hardcoding labels.
    """
    sheet = raw_data.get(sheet_key, {})
    respondents = sheet.get("respondents", [])
    quarters = sorted(set(r["quarter"] for r in respondents if r.get("quarter")))
    return quarters


def _match_quarter(label: str, available: list[str]) -> str:
    """Fuzzy-match a quarter label against available quarters.

    Handles common variations: "Q4'25" vs "Q4 2025" vs "Q4'2025" vs "Q4 '25".
    Returns the matched label or empty string if no match.
    """
    if not label or not available:
        return ""
    # Exact match first
    if label in available:
        return label
    # Normalize: strip spaces, lowercase
    def _norm(s):
        return s.replace(" ", "").replace("'", "").replace("\u2019", "").lower()
    norm_label = _norm(label)
    for q in available:
        if _norm(q) == norm_label:
            return q
    # Partial match: extract quarter+year digits
    import re
    m = re.search(r'[qQ](\d)[\'"]?\s*(\d{2,4})', label)
    if m:
        qn, yr = m.group(1), m.group(2)[-2:]  # last 2 digits of year
        for q in available:
            m2 = re.search(r'[qQ](\d)[\'"]?\s*(\d{2,4})', q)
            if m2 and m2.group(1) == qn and m2.group(2)[-2:] == yr:
                return q
    logger.warning("Quarter '%s' not found in data. Available: %s", label, available)
    return ""


def discover_codes(raw_data: dict, sheet_key: str = "primary") -> dict[str, int]:
    """Return dict of {question_code: n_columns} for a sheet.

    Useful for checking which codes are available before configuring extractions.
    """
    sheet = raw_data.get(sheet_key, {})
    code_map = sheet.get("code_map", {})
    return {code: len(cols) for code, cols in code_map.items()}


def validate_raw_extractions(config, raw_data: dict) -> list[str]:
    """Check all raw_aggregate extractions and report missing codes or quarters.

    Returns list of warning messages (empty if all OK).
    """
    warnings = []
    for ex in config.extractions:
        if ex.method != "raw_aggregate":
            continue
        params = ex.params
        sheet_key = params.get("raw_sheet", ex.sheet)
        code = params.get("code", "")
        sheet = raw_data.get(sheet_key, {})
        code_map = sheet.get("code_map", {})

        if not sheet or sheet_key not in raw_data:
            warnings.append(f"[{ex.id}] sheet '{sheet_key}' not found in raw data")
            continue
        if code and code not in code_map:
            # Try partial match
            matches = [c for c in code_map if code in c]
            if matches:
                warnings.append(
                    f"[{ex.id}] code '{code}' not found in '{sheet_key}', "
                    f"but similar codes exist: {matches[:5]}"
                )
            else:
                warnings.append(f"[{ex.id}] code '{code}' not found in '{sheet_key}'")

        # Check quarters
        quarters = discover_quarters(raw_data, sheet_key)
        for qkey in ("quarter_current", "quarter_prior"):
            qlabel = params.get(qkey, "")
            if qlabel and not _match_quarter(qlabel, quarters):
                warnings.append(
                    f"[{ex.id}] {qkey}='{qlabel}' not found. Available: {quarters}"
                )

    return warnings


# ── Aggregation engine ─────────────────────────────────────────────────

def aggregate_raw(raw_data: dict, sheet_key: str, code: str,
                   agg: str = "top2box",
                   quarter_current: str = "",
                   quarter_prior: str = "",
                   label_fn: Optional[Callable[[str], str]] = None,
                   ) -> list[dict]:
    """Aggregate respondent-level data for a question code.

    For multi-column questions (e.g. Q1_87Z with 14 attribute columns),
    returns one row per attribute with {desc, prior, current, code} —
    same format as existing extractors.

    For single-column questions (e.g. C1_81Z), returns one row.

    Args:
        raw_data: parsed raw data dict from load_raw_data()
        sheet_key: "primary", "competitor", "primary_npp", etc.
        code: question code (e.g. "Q1_87Z")
        agg: aggregation function — "top2box", "yes_pct", "recall_pct", "mean"
        quarter_current: quarter label for current period (e.g. "Q4'25")
        quarter_prior: quarter label for prior period (e.g. "Q3'25")
        label_fn: optional label shortening function

    Returns:
        list[dict] with {desc, code, prior, current, n_current, n_prior}
    """
    sheet = raw_data.get(sheet_key)
    if not sheet or isinstance(sheet, dict) and "_meta" in sheet and len(sheet) == 1:
        return []

    columns = sheet.get("columns", {})
    code_map = sheet.get("code_map", {})
    respondents = sheet.get("respondents", [])

    # Find columns for this code
    col_indices = code_map.get(code, [])
    if not col_indices:
        logger.warning("raw_aggregate: code '%s' not found in sheet '%s'", code, sheet_key)
        return []

    # Convert col_indices keys to int (JSON loads them as strings)
    col_indices = [int(ci) for ci in col_indices]

    agg_fn = _AGG_FUNCS.get(agg, _top2box)

    # Fuzzy-match quarter labels against available data
    available_quarters = discover_quarters(raw_data, sheet_key)
    q_cur = _match_quarter(quarter_current, available_quarters) if quarter_current else ""
    q_pri = _match_quarter(quarter_prior, available_quarters) if quarter_prior else ""

    # Split respondents by quarter
    current_resps = [r for r in respondents if r["quarter"] == q_cur] if q_cur else respondents
    prior_resps = [r for r in respondents if r["quarter"] == q_pri] if q_pri else []

    if q_cur and not current_resps:
        logger.warning("raw_aggregate: 0 respondents for quarter_current='%s' (matched='%s')",
                       quarter_current, q_cur)
    if q_pri and not prior_resps:
        logger.warning("raw_aggregate: 0 respondents for quarter_prior='%s' (matched='%s')",
                       quarter_prior, q_pri)

    results = []
    for ci in col_indices:
        ci_str = str(ci)  # JSON keys are strings
        col_info = columns.get(str(ci), columns.get(ci, {}))
        attribute = col_info.get("attribute", "")
        if not attribute:
            continue

        # Collect values for this column by period
        cur_vals = [r["values"].get(ci_str, r["values"].get(ci)) for r in current_resps
                    if ci_str in r["values"] or ci in r.get("values", {})]
        pri_vals = [r["values"].get(ci_str, r["values"].get(ci)) for r in prior_resps
                    if ci_str in r["values"] or ci in r.get("values", {})]

        current_val = agg_fn(cur_vals)
        prior_val = agg_fn(pri_vals) if pri_vals else None

        label = label_fn(attribute) if label_fn else attribute

        results.append({
            "desc": label,
            "code": code,
            "prior": prior_val,
            "current": current_val,
            "n_current": len(cur_vals),
            "n_prior": len(pri_vals),
        })

    return results


def aggregate_raw_multi_code(raw_data: dict, sheet_key: str,
                              codes: list[dict],
                              agg: str = "top2box",
                              quarter_current: str = "",
                              quarter_prior: str = "",
                              ) -> list[dict]:
    """Aggregate multiple independent question codes into one result list.

    Each code entry: {"code": "C1_81Z", "label": "Compelling reason to prescribe"}
    Returns one row per code with {desc, code, prior, current, n_current, n_prior}.

    Use for CTA-style slides where each metric is a different question code.
    """
    sheet = raw_data.get(sheet_key, {})
    if not sheet:
        return []

    code_map = sheet.get("code_map", {})
    columns = sheet.get("columns", {})
    respondents = sheet.get("respondents", [])
    agg_fn = _AGG_FUNCS.get(agg, _top2box)

    available_quarters = discover_quarters(raw_data, sheet_key)
    q_cur = _match_quarter(quarter_current, available_quarters) if quarter_current else ""
    q_pri = _match_quarter(quarter_prior, available_quarters) if quarter_prior else ""

    current_resps = [r for r in respondents if r["quarter"] == q_cur] if q_cur else respondents
    prior_resps = [r for r in respondents if r["quarter"] == q_pri] if q_pri else []

    results = []
    for entry in codes:
        code = entry["code"]
        label = entry["label"]
        col_indices = code_map.get(code, [])
        if not col_indices:
            logger.warning("raw_multi_code: code '%s' not found in '%s'", code, sheet_key)
            continue

        # For single-column codes (Yes/No), use the first column
        # For multi-column codes, use the first that has data
        ci_str = str(col_indices[0])
        cur_vals = [r["values"].get(ci_str) for r in current_resps if ci_str in r["values"]]
        pri_vals = [r["values"].get(ci_str) for r in prior_resps if ci_str in r["values"]]

        results.append({
            "desc": label,
            "code": code,
            "prior": agg_fn(pri_vals) if pri_vals else None,
            "current": agg_fn(cur_vals),
            "n_current": len(cur_vals),
            "n_prior": len(pri_vals),
        })

    return results


def aggregate_raw_cross_brand(raw_data: dict,
                               primary_sheet: str, comp_sheet: str,
                               code: str,
                               agg: str = "top2box",
                               quarter_current: str = "",
                               quarter_prior: str = "",
                               label_fn: Optional[Callable[[str], str]] = None,
                               ) -> list[dict]:
    """Aggregate the same question code across two brand sheets (e.g. RYB vs TAG).

    Returns rows with {desc, primary_prior, primary_current, comp_prior, comp_current}
    — same format as the rep_perf extraction from the Analysis sheet.

    Only includes attributes present in BOTH sheets.
    """
    primary_result = aggregate_raw(raw_data, primary_sheet, code, agg,
                                    quarter_current, quarter_prior, label_fn)
    comp_result = aggregate_raw(raw_data, comp_sheet, code, agg,
                                 quarter_current, quarter_prior, label_fn)

    # Join by desc (label)
    comp_by_desc = {r["desc"]: r for r in comp_result}
    results = []
    for pr in primary_result:
        cr = comp_by_desc.get(pr["desc"])
        row = {
            "desc": pr["desc"],
            "primary_prior": pr.get("prior"),
            "primary_current": pr.get("current"),
            "comp_prior": cr.get("prior") if cr else None,
            "comp_current": cr.get("current") if cr else None,
        }
        results.append(row)

    return results


def aggregate_raw_by_segment(raw_data: dict, sheet_key: str, code: str,
                              segment_code: str,
                              segment_values: dict[str, list],
                              agg: str = "top2box",
                              quarter: str = "",
                              label_fn: Optional[Callable[[str], str]] = None,
                              ) -> list[dict]:
    """Aggregate a question code split by a segment question.

    segment_code: question code used for segmenting (e.g. practice setting)
    segment_values: mapping of output field name to accepted raw values
        e.g. {"acad": ["Academic"], "comm": ["Community"]}

    Returns rows with {desc, acad_current, comm_current, diff} —
    same format as question_code_multi_col extractions.
    """
    sheet = raw_data.get(sheet_key, {})
    if not sheet:
        return []

    columns = sheet.get("columns", {})
    code_map = sheet.get("code_map", {})
    respondents = sheet.get("respondents", [])
    agg_fn = _AGG_FUNCS.get(agg, _top2box)

    available_quarters = discover_quarters(raw_data, sheet_key)
    q = _match_quarter(quarter, available_quarters) if quarter else ""

    resps = [r for r in respondents if r["quarter"] == q] if q else respondents

    # Find segment column
    seg_cols = code_map.get(segment_code, [])
    if not seg_cols:
        logger.warning("raw_by_segment: segment code '%s' not found in '%s'", segment_code, sheet_key)
        return []
    seg_ci_str = str(seg_cols[0])

    # Split respondents by segment
    seg_resps = {}
    for seg_name, accepted_vals in segment_values.items():
        accepted_lower = [str(v).lower() for v in accepted_vals]
        seg_resps[seg_name] = [
            r for r in resps
            if str(r["values"].get(seg_ci_str, "")).lower() in accepted_lower
        ]

    # Aggregate question code per segment
    target_cols = code_map.get(code, [])
    if not target_cols:
        logger.warning("raw_by_segment: code '%s' not found in '%s'", code, sheet_key)
        return []

    results = []
    for ci in [int(c) for c in target_cols]:
        ci_str = str(ci)
        col_info = columns.get(ci_str, {})
        attribute = col_info.get("attribute", "")
        if not attribute:
            continue

        row = {"desc": label_fn(attribute) if label_fn else attribute}
        seg_vals = {}
        for seg_name, seg_r in seg_resps.items():
            vals = [r["values"].get(ci_str) for r in seg_r if ci_str in r["values"]]
            agg_val = agg_fn(vals)
            row[f"{seg_name}_current"] = agg_val
            seg_vals[seg_name] = agg_val

        # Compute diff (first segment minus second)
        seg_keys = list(segment_values.keys())
        if len(seg_keys) >= 2:
            v1 = seg_vals.get(seg_keys[0])
            v2 = seg_vals.get(seg_keys[1])
            row["diff"] = round(v1 - v2, 1) if v1 is not None and v2 is not None else None

        results.append(row)

    return results
