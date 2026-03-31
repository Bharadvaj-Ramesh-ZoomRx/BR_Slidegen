"""
raw_data_loader.py — Respondent-level raw data loader and aggregator.

Reads source_raw_data.xlsx (one row per HCP response) and computes
aggregated metrics (top-2-box %, yes %, recall %, mean) on the fly.

Caching: Parsed data is cached as ``source_raw_data.pkl`` (pickle) for fast
reloading and native Python type preservation. The pkl auto-invalidates when
the source Excel changes (hash check). Delete the pkl to force re-parsing.

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
import pickle
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


# ── DataFrame → raw_data conversion ─────────────────────────────────────

def dataframes_to_raw_data(
    dataframes: dict,
    header_metadata: dict | None = None,
    sheet_map: dict | None = None,
) -> dict:
    """Convert Synapse JSON-sourced DataFrames to the dict format used by aggregators.

    The DataFrame columns match the Synapse download-responses layout: 32 static
    metadata columns followed by N dynamic question response columns.  The header
    metadata (from NDJSON header objects) provides question codes and type strings.

    Args:
        dataframes: {sheet_name: pd.DataFrame} from synapse_raw_fetcher.
        header_metadata: {sheet_name: header_rows list} from pkl ``_header_metadata``.
        sheet_map: Optional mapping of sheet_name → sheet_key.  If omitted, sheet
            names are normalised to lowercase keys.

    Returns:
        dict in the same shape as ``load_raw_data()`` — ``{sheet_key: {...}}``.
    """
    if not dataframes:
        return {}

    raw_data = {}
    header_metadata = header_metadata or {}

    for sheet_name, df in dataframes.items():
        if sheet_map:
            sheet_key = sheet_map.get(sheet_name, sheet_name.lower().replace(" ", "_"))
        else:
            sheet_key = sheet_name.lower().replace(" ", "_")

        col_names = list(df.columns)
        h_rows = header_metadata.get(sheet_name, [])

        # Try to extract question codes from header rows.
        # Synapse header_rows layout (8 rows): SGQA, QIDs, titles, text, option_groups, sub_q_text, abstract
        # The row at index _ROW_CODES (4) contains question codes.
        codes_row = h_rows[_ROW_CODES] if len(h_rows) > _ROW_CODES else []
        qtypes_row = h_rows[_ROW_QTYPES] if len(h_rows) > _ROW_QTYPES else []
        qtext_row = h_rows[_ROW_QTEXT] if len(h_rows) > _ROW_QTEXT else []

        columns = {}
        code_map = defaultdict(list)

        for ci in range(_COL_META_END, len(col_names)):
            code = codes_row[ci] if ci < len(codes_row) else ""
            if not code:
                # Fallback: use column name as code
                code = str(col_names[ci]).strip() if ci < len(col_names) else ""
            code = str(code).strip()
            if not code:
                continue

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

            q_text = str(qtext_row[ci])[:200] if ci < len(qtext_row) and qtext_row[ci] else ""

            ci_str = str(ci)
            columns[ci_str] = {
                "code": code,
                "text": q_text,
                "type": q_type,
                "attribute": attribute,
            }
            code_map[code].append(ci_str)

        # Build respondent list from DataFrame rows
        respondents = []
        id_col_idx = _COL_ID
        quarter_col_idx = _COL_QUARTER

        for _, row in df.iterrows():
            resp_id = row.iloc[id_col_idx] if id_col_idx < len(row) else None
            if resp_id is None:
                continue
            quarter = row.iloc[quarter_col_idx] if quarter_col_idx < len(row) else None

            resp_values = {}
            for ci_str in columns:
                ci = int(ci_str)
                if ci < len(row):
                    val = row.iloc[ci]
                    if val is not None and str(val) != "" and str(val) != "nan":
                        try:
                            resp_values[ci_str] = float(val) if isinstance(val, (int, float)) else val
                        except (ValueError, TypeError):
                            resp_values[ci_str] = val

            respondents.append({
                "id": resp_id,
                "quarter": quarter,
                "values": resp_values,
            })

        raw_data[sheet_key] = {
            "columns": columns,
            "code_map": dict(code_map),
            "respondents": respondents,
        }

        n_resp = len(respondents)
        n_codes = len(code_map)
        print(f"    {sheet_name} → {sheet_key}: {n_resp} respondents, {n_codes} question codes")

    return raw_data


# ── Caching ─────────────────────────────────────────────────────────────

def _file_hash(path: str) -> str:
    """Fast MD5 hash for staleness checks."""
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()[:12]


def _raw_pkl_path(config) -> str:
    """Path to source_raw_data.pkl cache."""
    if config.context_path:
        return os.path.join(config.context_path, "source_raw_data.pkl")
    return os.path.join(
        os.path.dirname(getattr(config, "raw_data_source_path", "")),
        "source_raw_data.pkl",
    )


def load_raw_data(config) -> dict | None:
    """Load and cache parsed raw data from source_raw_data.xlsx.

    Caching uses pickle (.pkl) for fast serialisation and native Python type
    preservation. The pkl auto-invalidates when the source Excel changes.

    Returns dict mapping sheet_key → parsed sheet data, or None if no raw data file.
    """
    raw_path = getattr(config, "raw_data_source_path", "")
    if not raw_path or not os.path.exists(raw_path):
        return None

    pkl_path = _raw_pkl_path(config)

    # Check pkl cache
    if os.path.exists(pkl_path):
        try:
            with open(pkl_path, "rb") as f:
                cached = pickle.load(f)
            meta = cached.get("_meta", {})
            if meta.get("raw_hash") == _file_hash(raw_path):
                n_sheets = sum(1 for k in cached if not k.startswith("_"))
                print(f"  Raw data loaded from pkl cache: {pkl_path} ({n_sheets} sheets)")
                return cached
            else:
                print("  Raw data pkl stale (Excel changed) — re-parsing")
        except (pickle.UnpicklingError, EOFError, Exception) as e:
            logger.warning("Failed to load pkl cache: %s — re-parsing", e)

    # Remove legacy JSON cache if present
    _remove_legacy_json_cache(config)

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

    # Save pkl cache
    raw_data["_meta"] = {
        "parsed_at": datetime.now().isoformat(),
        "raw_file": os.path.basename(raw_path),
        "raw_hash": _file_hash(raw_path),
    }

    os.makedirs(os.path.dirname(pkl_path), exist_ok=True)
    with open(pkl_path, "wb") as f:
        pickle.dump(raw_data, f, protocol=pickle.HIGHEST_PROTOCOL)
    size_mb = os.path.getsize(pkl_path) / (1024 * 1024)
    print(f"  Saved raw data pkl cache: {pkl_path} ({size_mb:.1f} MB)")

    return raw_data


def _remove_legacy_json_cache(config) -> None:
    """Remove old source_raw_data.json cache if present (migrated to pkl)."""
    if config.context_path:
        json_path = os.path.join(config.context_path, "source_raw_data.json")
    else:
        json_path = os.path.join(
            os.path.dirname(getattr(config, "raw_data_source_path", "")),
            "source_raw_data.json",
        )
    if os.path.exists(json_path):
        os.remove(json_path)
        print(f"  Removed legacy JSON cache: {os.path.basename(json_path)}")


# ── Quarter discovery & matching ──────────────────────────────────────

def discover_quarters(raw_data: dict, sheet_key: str = "primary") -> list[str]:
    """Return sorted list of quarter labels found in the data.

    Useful for auto-detecting available periods without hardcoding labels.
    """
    sheet = raw_data.get(sheet_key, {})
    respondents = sheet.get("respondents", [])
    quarters = sorted(set(r["quarter"] for r in respondents if r.get("quarter")))
    return quarters


_quarter_cache: dict[tuple[str, tuple[str, ...]], str] = {}


def _match_quarter(label: str, available: list[str]) -> str:
    """Fuzzy-match a quarter label against available quarters.

    Handles common variations: "Q4'25" vs "Q4 2025" vs "Q4'2025" vs "Q4 '25".
    Returns the matched label or empty string if no match.
    Results are cached to avoid repeated regex matching across extractions.
    """
    if not label or not available:
        return ""

    cache_key = (label, tuple(sorted(available)))
    if cache_key in _quarter_cache:
        return _quarter_cache[cache_key]

    result = _match_quarter_impl(label, available)
    _quarter_cache[cache_key] = result
    return result


def _match_quarter_impl(label: str, available: list[str]) -> str:
    """Internal quarter matching logic (uncached)."""
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
                   skip_zero: bool = False,
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
        skip_zero: if True, omit rows where both prior and current are 0 or None

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

    # Normalize all column indices to strings (JSON keys are always strings)
    col_indices = [str(ci) for ci in col_indices]

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
    for ci_str in col_indices:
        col_info = columns.get(ci_str, {})
        attribute = col_info.get("attribute", "")
        if not attribute:
            continue

        # Collect values for this column by period (all keys are strings)
        cur_vals = [r["values"][ci_str] for r in current_resps if ci_str in r["values"]]
        pri_vals = [r["values"][ci_str] for r in prior_resps if ci_str in r["values"]]

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

    if skip_zero:
        results = [r for r in results
                   if not (r["current"] in (0, 0.0, None)
                           and r["prior"] in (0, 0.0, None))]

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


def merge_vq_data(raw_data: dict, vq_data: dict, vq_questions: list[dict],
                   sheet_key: str = "primary") -> None:
    """Merge virtual question responses into parsed raw data in-place.

    Adds VQ columns to the sheet's columns dict and code_map, and adds
    VQ response values to each respondent's values dict.

    Args:
        raw_data: Parsed raw data dict from load_raw_data().
        vq_data: {qid: [{users_wave_id, value, ...}]} from VQ export.
        vq_questions: [{qid, title, question, type}] VQ metadata.
        sheet_key: Which sheet to merge into.
    """
    sheet = raw_data.get(sheet_key)
    if not sheet:
        return

    columns = sheet.get("columns", {})
    code_map = sheet.get("code_map", {})
    respondents = sheet.get("respondents", [])

    # Build respondent lookup: id → respondent dict
    resp_lookup = {}
    for r in respondents:
        rid = str(r.get("id", ""))
        if rid:
            resp_lookup[rid] = r

    # Assign virtual column indices starting after the last real column
    max_col = max((int(k) for k in columns.keys()), default=100) + 1
    vq_count = 0

    for vq in vq_questions:
        qid = vq["qid"]
        if qid not in vq_data:
            continue

        col_idx = str(max_col)
        vq_code = f"VQ_{qid}"

        columns[col_idx] = {
            "code": vq_code,
            "text": vq.get("question", vq.get("title", "")),
            "type": vq.get("type", "virtual"),
            "attribute": vq.get("title", f"VQ {qid}"),
        }
        code_map[vq_code] = [col_idx]

        # Merge values into respondents
        for vq_row in vq_data[qid]:
            rid = str(vq_row.get("users_wave_id", ""))
            if rid in resp_lookup:
                resp_lookup[rid]["values"][col_idx] = vq_row["value"]

        max_col += 1
        vq_count += 1

    if vq_count:
        print(f"  Merged {vq_count} virtual questions into {sheet_key}")


def apply_segment_filter(raw_data: dict, segment_cuts: list, sheet_key: str = "primary") -> dict:
    """Apply segment filter cuts to raw data, returning filtered respondents.

    For "filter" mode segments: keeps only respondents matching the specified values.
    For "groupby" mode segments: no filtering (groupby is applied during aggregation).

    Args:
        raw_data: Parsed raw data dict.
        segment_cuts: List of SegmentCutConfig-like dicts with id, mode, values.
        sheet_key: Which sheet to filter.

    Returns:
        New raw_data dict with filtered respondents (original not modified).
    """
    filter_cuts = [sc for sc in segment_cuts if sc.get("mode") == "filter" and sc.get("values")]
    if not filter_cuts:
        return raw_data

    sheet = raw_data.get(sheet_key)
    if not sheet:
        return raw_data

    code_map = sheet.get("code_map", {})
    respondents = sheet.get("respondents", [])

    # For each filter segment, find the column and filter respondents
    filtered = list(respondents)
    for sc in filter_cuts:
        seg_code = sc.get("segment_code", "")
        if not seg_code:
            # Try to find segment column by segment name in columns
            seg_name = sc.get("name", "").lower()
            for code, cols in code_map.items():
                if seg_name in code.lower():
                    seg_code = code
                    break

        if not seg_code:
            logger.warning("Segment filter: no segment_code for segment id=%s", sc.get("id"))
            continue

        seg_cols = code_map.get(seg_code, [])
        if not seg_cols:
            continue

        seg_ci = str(seg_cols[0])
        accepted = [str(v).lower() for v in sc["values"]]
        before = len(filtered)
        filtered = [
            r for r in filtered
            if str(r["values"].get(seg_ci, "")).lower() in accepted
        ]
        print(f"  Segment filter '{seg_code}' ({sc['values']}): {before} → {len(filtered)} respondents")

    # Return new raw_data with filtered respondents
    import copy
    new_data = copy.copy(raw_data)
    new_sheet = dict(sheet)
    new_sheet["respondents"] = filtered
    new_data[sheet_key] = new_sheet
    return new_data


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
    for ci_str in [str(c) for c in target_cols]:
        col_info = columns.get(ci_str, {})
        attribute = col_info.get("attribute", "")
        if not attribute:
            continue

        row = {"desc": label_fn(attribute) if label_fn else attribute}
        seg_vals = {}
        for seg_name, seg_r in seg_resps.items():
            vals = [r["values"][ci_str] for r in seg_r if ci_str in r["values"]]
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


def aggregate_raw_by_hii(raw_data: dict, sheet_key: str, code: str,
                          quality_code: str = "Q1_87Z",
                          ltip_code: str = "C1_85DZ",
                          agg: str = "top2box",
                          quarter: str = "",
                          label_fn: Optional[Callable[[str], str]] = None,
                          skip_zero: bool = False,
                          ) -> list[dict]:
    """Aggregate a question code split by HII vs Others.

    HII = quality_code overall quality == 7 (top-box) AND ltip_code any col >= 6.
    Others = everyone else with valid values for both.

    Returns rows with {desc, hi_current, other_current, diff}.
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

    # Find overall quality column
    quality_cols = code_map.get(quality_code, [])
    quality_ci = None
    for ci in quality_cols:
        attr = columns.get(str(ci), {}).get("attribute", "").lower()
        if "overall quality" in attr:
            quality_ci = str(ci)
            break
    if quality_ci is None and quality_cols:
        quality_ci = str(quality_cols[0])
    if quality_ci is None:
        logger.warning("by_hii: quality_code '%s' not found", quality_code)
        return []

    ltip_cols = [str(ci) for ci in code_map.get(ltip_code, [])]
    if not ltip_cols:
        logger.warning("by_hii: ltip_code '%s' not found", ltip_code)
        return []

    # Classify respondents
    hii_resps, other_resps = [], []
    for r in resps:
        q_val = r["values"].get(quality_ci)
        ltip_vals = [r["values"].get(c) for c in ltip_cols if c in r["values"]]
        if q_val is None or not ltip_vals:
            continue
        ltip_max = max((v for v in ltip_vals if isinstance(v, (int, float))), default=None)
        if ltip_max is None:
            continue
        if q_val == 7 and ltip_max >= 6:
            hii_resps.append(r)
        else:
            other_resps.append(r)

    logger.info("by_hii: %d HII, %d Others (of %d total)",
                len(hii_resps), len(other_resps), len(resps))

    # Aggregate target code per segment
    target_cols = code_map.get(code, [])
    if not target_cols:
        logger.warning("by_hii: code '%s' not found", code)
        return []

    results = []
    for ci in [int(c) for c in target_cols]:
        ci_str = str(ci)
        col_info = columns.get(ci_str, {})
        attribute = col_info.get("attribute", "")
        if not attribute:
            continue

        hi_vals = [r["values"].get(ci_str) for r in hii_resps if ci_str in r["values"]]
        ot_vals = [r["values"].get(ci_str) for r in other_resps if ci_str in r["values"]]

        hi_agg = agg_fn(hi_vals)
        ot_agg = agg_fn(ot_vals)

        row = {
            "desc": label_fn(attribute) if label_fn else attribute,
            "hi_current": hi_agg,
            "other_current": ot_agg,
        }
        if hi_agg is not None and ot_agg is not None:
            row["diff"] = round(hi_agg - ot_agg, 1)
        else:
            row["diff"] = None
        results.append(row)

    if skip_zero:
        results = [r for r in results
                   if not (r["hi_current"] in (0, 0.0, None)
                           and r["other_current"] in (0, 0.0, None))]

    return results
