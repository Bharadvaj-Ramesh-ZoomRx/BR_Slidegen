"""
data_loaders.py — Generic data extraction functions.

All the repeated "find code, walk rows, extract Q3/Q4" patterns
from generate_asks.py are consolidated here into reusable extractors.

Auto-JSON mode: On first run, data is extracted from Excel and saved as
`source_data.json` alongside the Excel file. Subsequent runs read the JSON
directly — no pandas, no column indices, no question-code walking. Delete
`source_data.json` to force re-extraction, or it auto-invalidates when the
Excel file changes.
"""

from __future__ import annotations
import hashlib
import json
import logging
import os
import pandas as pd
from datetime import datetime
from typing import Optional, Callable

logger = logging.getLogger(__name__)


# ── Value converters ─────────────────────────────────────────────────────────

def pct(v) -> Optional[float]:
    """Convert decimal 0-1 to percentage rounded to 1 decimal."""
    if v is None:
        return None
    try:
        return round(float(v) * 100, 1)
    except (ValueError, TypeError):
        return None


def straight(v) -> Optional[float]:
    """Value is already a percentage, just round."""
    if v is None:
        return None
    try:
        return round(float(v), 1)
    except (ValueError, TypeError):
        return None


def delta(current, prior) -> Optional[float]:
    """Current minus prior in ppts, rounded to 1 decimal."""
    if current is None or prior is None:
        return None
    return round(current - prior, 1)


# ── Label shortening ────────────────────────────────────────────────────────

def make_label_shortener(shortcuts: list[dict], max_len: int = 35) -> Callable[[str], str]:
    """Build a label shortening function from a list of keyword→short mappings.

    Each shortcut: {"keywords": ["kw1", "kw2"], "short": "Short Label"}
    """
    def shorten(label: str) -> str:
        ll = label.lower()
        for sc in shortcuts:
            kws = sc["keywords"]
            if all(kw.lower() in ll for kw in kws):
                return sc["short"]
        # Generic cleanup
        label = label.replace("Johnson & Johnson (Formerly Janssen)", "J&J")
        label = label.replace("Johnson &amp; Johnson", "J&J")
        label = label.replace("[COMPANY]", "J&J")
        label = label.replace("[PRODUCT]", "RYB")
        label = label.replace("How well the ", "").replace("How ", "")
        label = label.strip()
        if len(label) > max_len:
            label = label[:max_len - 1] + "\u2026"
        return label

    return shorten


# ── Generic extractors ───────────────────────────────────────────────────────

def extract_by_question_code(
    sheet: pd.DataFrame,
    code: str,
    q_prior_col: int,
    q_current_col: int,
    code_col: int = 0,
    desc_col: int = 1,
    max_rows: int = 20,
    label_fn: Optional[Callable[[str], str]] = None,
    converter: Callable = pct,
) -> list[dict]:
    """Find a question code row, walk sub-rows, extract prior/current values.

    Returns list of: {"desc": str, "prior": float, "current": float, "code": str}
    """
    results = []
    for i in range(len(sheet)):
        if sheet.iloc[i, code_col] == code:
            j = i + 1
            while j < len(sheet) and pd.notna(sheet.iloc[j, desc_col]):
                desc = str(sheet.iloc[j, desc_col]).strip()
                q_prior = sheet.iloc[j, q_prior_col]
                q_current = sheet.iloc[j, q_current_col]
                row_code = sheet.iloc[j, code_col] if pd.notna(sheet.iloc[j, code_col]) else ""

                if desc and not desc.startswith("Base") and pd.notna(q_current):
                    prior_val = converter(q_prior) if pd.notna(q_prior) else None
                    current_val = converter(q_current) if pd.notna(q_current) else None
                    # Skip stale rows where both prior and current are 0
                    if (prior_val is None or prior_val == 0) and (current_val is None or current_val == 0):
                        j += 1
                        continue
                    label = label_fn(desc) if label_fn else desc
                    results.append({
                        "desc": label,
                        "code": str(row_code),
                        "prior": prior_val,
                        "current": current_val,
                    })
                j += 1
                if j - i > max_rows:
                    break
            break
    return results


def extract_multi_question_code(
    sheet: pd.DataFrame,
    codes: list[dict],
    q_prior_col: int,
    q_current_col: int,
    code_col: int = 0,
    desc_col: int = 1,
    converter: Callable = pct,
) -> list[dict]:
    """Extract one row per question code (e.g. CTA metrics).

    Each code entry: {"code": "C1_81Z", "label": "Compelling reason to prescribe"}
    Optional "pick" key: match a specific sub-row code (e.g. "DIA", "WCNS", "A2").
    Without "pick", finds the code row, then looks for "top"/"yes" sub-row or takes first data row.
    """
    results = []
    for entry in codes:
        code = entry["code"]
        label = entry["label"]
        pick = entry.get("pick")
        for i in range(len(sheet)):
            if sheet.iloc[i, code_col] == code:
                for j in range(i, min(i + 15, len(sheet))):
                    cell = str(sheet.iloc[j, code_col]) if pd.notna(sheet.iloc[j, code_col]) else ""
                    desc = str(sheet.iloc[j, desc_col]) if pd.notna(sheet.iloc[j, desc_col]) else ""
                    if pick:
                        if cell == pick:
                            q_prior = sheet.iloc[j, q_prior_col]
                            q_current = sheet.iloc[j, q_current_col]
                            if pd.notna(q_prior) and pd.notna(q_current):
                                results.append({
                                    "desc": label,
                                    "code": code,
                                    "prior": converter(q_prior),
                                    "current": converter(q_current),
                                })
                            break
                    elif "top" in desc.lower() or "yes" in desc.lower() or j == i + 1:
                        q_prior = sheet.iloc[j, q_prior_col]
                        q_current = sheet.iloc[j, q_current_col]
                        if pd.notna(q_prior) and pd.notna(q_current):
                            results.append({
                                "desc": label,
                                "code": code,
                                "prior": converter(q_prior),
                                "current": converter(q_current),
                            })
                            break
                break
    return results


def extract_row_range(
    sheet: pd.DataFrame,
    row_start: int,
    row_end: int,
    col_map: dict[str, int],
    label_fn: Optional[Callable[[str], str]] = None,
    converter: Callable = pct,
    min_label_len: int = 3,
) -> list[dict]:
    """Extract data from a fixed row range with a column mapping.

    col_map: {"label": 1, "primary_prior": 2, "primary_current": 3, ...}
    The "label" key is required; all others become data fields.
    """
    results = []
    label_col = col_map["label"]
    data_cols = {k: v for k, v in col_map.items() if k != "label" and k != "short"}
    short_col = col_map.get("short")

    for i in range(row_start, min(row_end, len(sheet))):
        raw_label = sheet.iloc[i, label_col]
        if not pd.notna(raw_label) or not isinstance(raw_label, str) or len(raw_label) < min_label_len:
            continue

        label = label_fn(raw_label.strip()) if label_fn else raw_label.strip()
        row = {"desc": label}

        if short_col is not None:
            short_val = sheet.iloc[i, short_col]
            row["short"] = str(short_val).strip() if pd.notna(short_val) else label[:30]

        # Check at least one data col is non-null
        has_data = False
        for key, col_idx in data_cols.items():
            val = sheet.iloc[i, col_idx]
            row[key] = converter(val) if pd.notna(val) else None
            if pd.notna(val):
                has_data = True

        if has_data:
            results.append(row)

    return results


def extract_question_code_multi_col(
    sheet: pd.DataFrame,
    code: str,
    columns: dict[str, int],
    code_col: int = 0,
    desc_col: int = 1,
    max_rows: int = 20,
    label_fn: Optional[Callable[[str], str]] = None,
    converter: Callable = pct,
    min_diff: float = 0,
) -> list[dict]:
    """Extract multiple columns per row under a question code.

    Used for HII comparison: hi_current=col20, other_current=col19.
    Returns list with computed 'diff' field.
    """
    results = []
    for i in range(len(sheet)):
        if sheet.iloc[i, code_col] == code:
            j = i + 1
            while j < len(sheet) and pd.notna(sheet.iloc[j, desc_col]):
                desc = str(sheet.iloc[j, desc_col]).strip()
                if desc and not desc.startswith("Base"):
                    row = {"desc": label_fn(desc) if label_fn else desc}
                    all_valid = True
                    for key, col_idx in columns.items():
                        val = sheet.iloc[j, col_idx]
                        if pd.notna(val):
                            row[key] = converter(val)
                        else:
                            all_valid = False
                    if all_valid and len(columns) >= 2:
                        vals = list(row.values())
                        # diff = first data col - second data col
                        data_vals = [v for k, v in row.items() if k != "desc"]
                        if len(data_vals) >= 2:
                            row["diff"] = round(data_vals[0] - data_vals[1], 1)
                    # Skip stale rows where all data values are 0
                    if all_valid:
                        data_vals = [v for k, v in row.items() if k not in ("desc", "diff")]
                        if not all(v == 0 for v in data_vals):
                            results.append(row)
                j += 1
                if j - i > max_rows:
                    break
            break

    if min_diff > 0:
        pre_filter = len(results)
        results = [r for r in results if abs(r.get("diff", 0)) > min_diff]
        results.sort(key=lambda x: abs(x.get("diff", 0)), reverse=True)
        if pre_filter > 0 and len(results) == 0:
            logger.warning(
                "extract_question_code_multi_col: min_diff=%.1f filtered all "
                "%d rows to 0 for code '%s'", min_diff, pre_filter, code
            )

    return results


def extract_nested_ordinal(
    sheet: pd.DataFrame,
    row_start: int,
    row_end: int,
    code_col: int = 0,
    desc_col: int = 1,
    ordinal_col: int = 2,
    ordinals: list[str] = None,
    q_prior_col: int = 7,
    q_current_col: int = 13,
    label_fn: Optional[Callable[[str], str]] = None,
    converter: Callable = pct,
) -> list[dict]:
    """Extract grouped data with ordinal sub-rows (e.g. recall order).

    Each message code has sub-rows for 1st, 2nd, 3rd, 4th recalled.
    Returns list with per-ordinal fields and totals.
    """
    if ordinals is None:
        ordinals = ["1st", "2nd", "3rd", "4th"]

    groups = {}  # code → {desc, 1st_current, 2nd_current, ..., total_current}

    for i in range(row_start, min(row_end, len(sheet))):
        code = str(sheet.iloc[i, code_col]).strip() if pd.notna(sheet.iloc[i, code_col]) else ""
        desc = str(sheet.iloc[i, desc_col]).strip() if pd.notna(sheet.iloc[i, desc_col]) else ""
        ordinal = str(sheet.iloc[i, ordinal_col]).strip() if pd.notna(sheet.iloc[i, ordinal_col]) else ""

        if not code or not desc:
            continue

        if code not in groups:
            label = label_fn(desc) if label_fn else desc
            entry = {"code": code, "desc": label}
            for o in ordinals:
                entry[f"{o}_current"] = 0
                entry[f"{o}_prior"] = 0
            groups[code] = entry

        entry = groups[code]
        q_cur = converter(sheet.iloc[i, q_current_col]) if pd.notna(sheet.iloc[i, q_current_col]) else 0
        q_pri = converter(sheet.iloc[i, q_prior_col]) if pd.notna(sheet.iloc[i, q_prior_col]) else 0

        for o in ordinals:
            if ordinal.lower().startswith(o.lower()):
                entry[f"{o}_current"] = q_cur
                entry[f"{o}_prior"] = q_pri
                break

    # Compute totals
    results = []
    for entry in groups.values():
        entry["total_current"] = round(sum(entry.get(f"{o}_current", 0) for o in ordinals), 1)
        entry["total_prior"] = round(sum(entry.get(f"{o}_prior", 0) for o in ordinals), 1)
        if entry["total_current"] > 0:
            results.append(entry)

    results.sort(key=lambda x: x["total_current"], reverse=True)
    return results


# ── Excel → JSON indexing (Stage 0) ─────────────────────────────────────────

def _is_zero(v) -> bool:
    """Check if a cell value is effectively zero (0, 0.0, '0', '0%', etc.)."""
    try:
        return float(v) == 0.0
    except (ValueError, TypeError):
        s = str(v).strip().rstrip("%")
        try:
            return float(s) == 0.0
        except (ValueError, TypeError):
            return False


def _detect_column_layout(header_rows: list) -> dict | None:
    """Detect prior/current/segment column positions from Excel header rows.

    Scans header rows (0-3) for time period labels (e.g., "Quarterly - Q4 2025",
    "Total", "High Impact", "Community", "Academic") and returns a layout dict:
    {
        "prior_col": int,       # 0-indexed column for prior wave total
        "current_col": int,     # 0-indexed column for current wave total
        "prior_label": str,     # e.g., "Quarterly - Q4 2025"
        "current_label": str,   # e.g., "Quarterly - Q1 2026"
        "segments": [           # segment columns detected from rows 2-3
            {"name": "High Impact - LTIP", "columns": {"Others": col, "High Impact": col}},
            {"name": "Acad-Comm Lite", "columns": {"Community": col, "Academic": col}},
        ]
    }
    """
    import re
    if not header_rows or len(header_rows) < 2:
        return None

    layout = {"segments": []}
    row1 = header_rows[0] if len(header_rows) > 0 else ()
    row2 = header_rows[1] if len(header_rows) > 1 else ()
    row3 = header_rows[2] if len(header_rows) > 2 else ()

    # Find "Total" columns in row 2 — these are the main prior/current totals
    total_cols = []
    for ci, val in enumerate(row2):
        if val and str(val).strip().lower() == "total":
            total_cols.append(ci)

    # Match Total columns to period labels in row 1
    period_pattern = re.compile(r'(Quarterly|Monthly|Wave)\s*[-–]\s*(.+)', re.IGNORECASE)
    period_cols = []
    for ci, val in enumerate(row1):
        if val:
            m = period_pattern.match(str(val).strip())
            if m:
                period_cols.append((ci, m.group(2).strip()))

    # Map each Total column to its nearest period label
    if len(total_cols) >= 2 and len(period_cols) >= 2:
        layout["prior_col"] = total_cols[0]
        layout["current_col"] = total_cols[1]
        layout["prior_label"] = period_cols[0][1] if period_cols else ""
        layout["current_label"] = period_cols[1][1] if len(period_cols) > 1 else ""
    elif len(total_cols) == 1:
        layout["prior_col"] = total_cols[0]
        layout["current_col"] = total_cols[0]
        layout["prior_label"] = period_cols[0][1] if period_cols else ""
        layout["current_label"] = layout["prior_label"]
    else:
        return None  # Can't detect layout

    # Detect segment columns from row 2-3
    seen_segments = set()
    for ci, val in enumerate(row2):
        if val and str(val).strip().lower() not in ("total", "", "varying sample for overall data"):
            seg_name = str(val).strip()
            if seg_name.startswith("Varying Sample"):
                continue
            if seg_name not in seen_segments:
                seen_segments.add(seg_name)
                # Find sub-columns in row 3
                sub_cols = {}
                for ci2 in range(ci, min(ci + 4, len(row3))):
                    if row3[ci2]:
                        sub_cols[str(row3[ci2]).strip()] = ci2
                if sub_cols:
                    layout["segments"].append({"name": seg_name, "columns": sub_cols})

    return layout


def index_excel(excel_path: str, json_path: str) -> dict:
    """Convert Excel to JSON index — runs as Stage 0 before any analysis.

    Reads all sheets from the Excel file and builds a row-level index with
    code (col 0) and description (col 1) for every non-empty row. The result
    is saved as source_data.json with a ``_sheets`` key containing the index.

    Downstream stages (hypotheses, slide plan, config generation) can search
    ``_sheets`` for question codes without touching Excel again.

    If source_data.json already exists and the Excel hasn't changed (hash
    check), returns the existing data without re-indexing.

    Args:
        excel_path: Path to source_data.xlsx
        json_path:  Path to write/read source_data.json

    Returns:
        dict with ``_meta`` and ``_sheets`` keys (and any existing extractions)
    """
    import openpyxl

    # Check for existing fresh index
    if os.path.exists(json_path):
        with open(json_path, "r", encoding="utf-8") as f:
            existing = json.load(f)
        meta = existing.get("_meta", {})
        if meta.get("excel_hash") == _file_hash(excel_path) and "_sheets" in existing and "_codes" in existing:
            n_codes = sum(len(v) for v in existing.get("_codes", {}).values())
            print(f"  source_data.json up to date — {len(existing['_sheets'])} sheets, {n_codes} codes indexed")
            return existing

    # Build raw sheet index from Excel
    print(f"  Indexing Excel: {os.path.basename(excel_path)}")
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    sheets_index = {}
    codes_index = {}   # rich per-code metadata for Stage 7 auto-detection

    # ── Column layout detection (auto-discover prior/current/segment columns) ──
    column_layouts = {}  # sheet_name → {prior_col, current_col, segments: [...]}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = []
        skipped = 0
        # Store all parsed rows for _codes analysis
        all_rows = []  # (row_idx, code, desc, data_vals_tuple)

        # Parse header rows (0-3) to detect column layout
        header_rows = []
        for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
            if row_idx < 4:
                header_rows.append(row)
                continue
            code = str(row[0]).strip() if row[0] is not None else ""
            desc = str(row[1]).strip()[:120] if len(row) > 1 and row[1] is not None else ""
            if code or desc:
                data_cols = row[2:] if len(row) > 2 else ()
                data_vals = [v for v in data_cols if v is not None and str(v).strip() != ""]
                if data_vals and all(_is_zero(v) for v in data_vals):
                    skipped += 1
                    continue
                rows.append({"row": row_idx, "code": code, "desc": desc})
                all_rows.append((row_idx, code, desc, row))
        sheets_index[sheet_name] = rows
        skipped_msg = f" ({skipped} stale removed)" if skipped else ""
        print(f"    {sheet_name}: {len(rows)} rows{skipped_msg}")

        # Build _codes: analyze question code patterns (codes ending in Z are
        # typically parent codes with sub-rows beneath them)
        import re
        _code_re = re.compile(r'^[A-Z]\d')
        sheet_codes = {}
        for idx, (row_idx, code, desc, raw_row) in enumerate(all_rows):
            if not code or not _code_re.match(code):
                continue
            # Count sub-rows: rows after this one until next code row or gap
            sub_row_count = 0
            has_sub_codes = False
            for j in range(idx + 1, len(all_rows)):
                _, next_code, next_desc, _ = all_rows[j]
                if next_code and _code_re.match(next_code):
                    # Another parent code — stop counting
                    if next_code != code:
                        break
                    has_sub_codes = True
                sub_row_count += 1
                if sub_row_count > 30:
                    break

            # Sample data values from specific columns
            sample_values = {}
            for ci in (7, 13, 17):
                if ci < len(raw_row) and raw_row[ci] is not None:
                    try:
                        v = float(raw_row[ci])
                        sample_values[f"col_{ci}"] = round(v, 4)
                    except (ValueError, TypeError):
                        pass

            # Detect value range: decimal (0-1) vs whole (0-100+)
            # When header row looks like a base size (>1), check sub-row values
            sample_cols = [7, 13, 17]
            value_range = "unknown"
            numeric_samples = [v for v in sample_values.values() if isinstance(v, (int, float))]
            if numeric_samples:
                max_v = max(abs(v) for v in numeric_samples)
                if max_v <= 1.1:
                    value_range = "decimal"
                elif sub_row_count > 0:
                    # Header may be a base size — check first sub-row values
                    sub_vals = []
                    for j2 in range(idx + 1, min(idx + 4, len(all_rows))):
                        _, _, _, sub_raw_row = all_rows[j2]
                        for sc in sample_cols:
                            if sc < len(sub_raw_row) and sub_raw_row[sc] is not None:
                                try:
                                    sv = float(sub_raw_row[sc])
                                    if sv > 0:
                                        sub_vals.append(abs(sv))
                                except (ValueError, TypeError):
                                    pass
                    if sub_vals and max(sub_vals) <= 1.1:
                        value_range = "decimal"
                    else:
                        value_range = "whole"
                else:
                    value_range = "whole"

            sheet_codes[code] = {
                "row": row_idx,
                "desc": desc,
                "sub_row_count": sub_row_count,
                "has_sub_codes": has_sub_codes,
                "sample_values": sample_values,
                "value_range": value_range,
            }
        if sheet_codes:
            codes_index[sheet_name] = sheet_codes

        # Detect column layout from header rows
        if header_rows:
            layout = _detect_column_layout(header_rows)
            if layout:
                column_layouts[sheet_name] = layout

    wb.close()

    # Build or update payload
    payload = {}
    if os.path.exists(json_path):
        with open(json_path, "r", encoding="utf-8") as f:
            payload = json.load(f)

    payload["_meta"] = {
        "extracted_at": datetime.now().isoformat(),
        "excel_file": os.path.basename(excel_path),
        "excel_hash": _file_hash(excel_path),
    }
    payload["_sheets"] = sheets_index
    if codes_index:
        payload["_codes"] = codes_index
    if column_layouts:
        payload["_column_layouts"] = column_layouts

    os.makedirs(os.path.dirname(json_path), exist_ok=True)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, default=str)

    print(f"  Saved: {json_path}")
    return payload


# ── JSON auto-cache ───────────��─────────────────────────────────��───────────

def _file_hash(path: str) -> str:
    """Fast MD5 hash of a file for staleness checks."""
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()[:12]


def _extractions_hash(config) -> str:
    """MD5 hash of serialized extraction configs for cache invalidation.

    If extraction params change (different question code, column, etc.)
    but the Excel file hasn't changed, this hash will differ and trigger
    re-extraction.
    """
    import json as _json
    extractions_data = []
    for ex in config.extractions:
        extractions_data.append({
            "id": ex.id, "method": ex.method,
            "sheet": ex.sheet, "params": ex.params,
        })
    serialized = _json.dumps(extractions_data, sort_keys=True, default=str)
    return hashlib.md5(serialized.encode()).hexdigest()[:12]


def _json_path_for(config) -> str:
    """Return path to source_data.json in the context folder."""
    if config.context_path:
        return os.path.join(config.context_path, "source_data.json")
    # Fallback: next to the Excel file
    return os.path.join(
        os.path.dirname(config.data_source_path), "source_data.json"
    )


def _load_source_json(config) -> dict | None:
    """Load source_data.json if it exists and is not stale.

    Returns None if missing or if the Excel file has changed since extraction.
    """
    json_path = _json_path_for(config)
    if not os.path.exists(json_path):
        return None

    with open(json_path, "r", encoding="utf-8") as f:
        payload = json.load(f)

    meta = payload.get("_meta", {})

    # Invalidate if Excel file changed
    if os.path.exists(config.data_source_path):
        current_hash = _file_hash(config.data_source_path)
        if meta.get("excel_hash") != current_hash:
            print("  source_data.json stale (Excel changed) — re-extracting")
            return None

    # Invalidate if extraction config changed
    current_ext_hash = _extractions_hash(config)
    if meta.get("extractions_hash") and meta["extractions_hash"] != current_ext_hash:
        print("  source_data.json stale (extraction params changed) — re-extracting")
        return None

    # Keep _meta alongside data (renderers ignore it; orchestrator uses extracted_at)
    data = dict(payload)
    n_extractions = sum(1 for k in data if not k.startswith("_"))

    # If JSON only has _sheets (index-only from Stage 0) but no extractions,
    # signal re-extraction needed
    if n_extractions == 0:
        print("  source_data.json has index only (no extractions) — extracting from Excel")
        return None

    print(f"  Loaded from JSON: {json_path} ({n_extractions} extractions)")
    return data


def _save_source_json(data: dict, config) -> str:
    """Save extracted data as source_data.json next to the Excel file."""
    json_path = _json_path_for(config)
    if not json_path or json_path == "source_data.json":
        # No context or data path configured — skip JSON cache
        return ""
    os.makedirs(os.path.dirname(json_path), exist_ok=True)

    excel_hash = ""
    if os.path.exists(config.data_source_path):
        excel_hash = _file_hash(config.data_source_path)

    payload = {
        "_meta": {
            "extracted_at": datetime.now().isoformat(),
            "excel_file": os.path.basename(config.data_source_path),
            "excel_hash": excel_hash,
            "extractions_hash": _extractions_hash(config),
            "wave": config.wave,
        },
    }
    payload.update(data)

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, default=str)

    print(f"  Saved: {json_path}")
    return json_path


# ── Excel extraction (internal) ─────────────────────────────────────────────

def _extract_all_from_excel(config) -> dict:
    """Run all extraction methods against the Excel source.

    This is the heavy-lift function that reads Excel via pandas and runs
    the 5 extraction methods. Called only when source_data.json is missing
    or stale.
    """
    # Check if any extraction actually needs Excel (skip for all-mock / all-synapse configs)
    _SKIP_EXCEL_METHODS = {"mock", "synapse_report", "synapse_raw", "raw_aggregate"}
    needs_excel = any(ex.method not in _SKIP_EXCEL_METHODS for ex in config.extractions)

    # Load Excel sheets (only if needed)
    sheets_data = {}
    if needs_excel and config.data_source_path:
        for sheet_key, sheet_cfg in config.sheets.items():
            sheets_data[sheet_key] = pd.read_excel(
                config.data_source_path,
                sheet_name=sheet_cfg.name,
                header=None,
            )

    # Reuse _sheets from existing JSON if available and Excel hash matches,
    # avoiding a redundant scan that index_excel() already performed.
    raw_index = None
    json_path = _json_path_for(config)
    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                existing = json.load(f)
            meta = existing.get("_meta", {})
            if (meta.get("excel_hash") == _file_hash(config.data_source_path)
                    and "_sheets" in existing):
                raw_index = existing["_sheets"]
                print("  Reusing _sheets index from existing source_data.json")
        except (json.JSONDecodeError, OSError):
            pass

    if raw_index is None:
        # Build raw sheet index from pandas (fallback when no prior index exists)
        raw_index = {}
        for sheet_key, df in sheets_data.items():
            rows = []
            for i in range(len(df)):
                code = df.iloc[i, 0]
                desc = df.iloc[i, 1] if df.shape[1] > 1 else None
                code_str = str(code).strip() if pd.notna(code) else ""
                desc_str = str(desc).strip()[:120] if pd.notna(desc) else ""
                if code_str or desc_str:
                    rows.append({"row": i, "code": code_str, "desc": desc_str})
            raw_index[sheet_key] = rows

    # Build global label shortener
    global_shortener = None
    if config.label_shortcuts:
        shortcuts = [{"keywords": ls.keywords, "short": ls.short} for ls in config.label_shortcuts]
        global_shortener = make_label_shortener(shortcuts)

    data = {}

    for ex in config.extractions:
        params = ex.params

        # Mock extractions don't need Excel
        if ex.method == "mock":
            data[ex.id] = params.get("rows", [])
            continue

        sheet_cfg = config.sheets[ex.sheet]
        df = sheets_data[ex.sheet]

        # Build label function
        label_fn = None
        if params.get("use_label_shortcuts") and global_shortener:
            label_fn = global_shortener
        elif "label_shortcuts" in params:
            local_shortcuts = params["label_shortcuts"]
            label_fn = make_label_shortener(local_shortcuts)

        # Determine converter
        converter = straight if params.get("pct_mode") == "straight" else pct

        if ex.method == "question_code":
            data[ex.id] = extract_by_question_code(
                df,
                code=params["code"],
                q_prior_col=sheet_cfg.q_prior_col,
                q_current_col=sheet_cfg.q_current_col,
                code_col=sheet_cfg.code_col,
                desc_col=sheet_cfg.desc_col,
                max_rows=params.get("max_rows", 20),
                label_fn=label_fn,
                converter=converter,
            )

        elif ex.method == "multi_question_code":
            data[ex.id] = extract_multi_question_code(
                df,
                codes=params["codes"],
                q_prior_col=sheet_cfg.q_prior_col,
                q_current_col=sheet_cfg.q_current_col,
                code_col=sheet_cfg.code_col,
                desc_col=sheet_cfg.desc_col,
                converter=converter,
            )

        elif ex.method == "row_range":
            data[ex.id] = extract_row_range(
                df,
                row_start=params["row_start"],
                row_end=params["row_end"],
                col_map=params["col_map"],
                label_fn=label_fn,
                converter=converter,
            )

        elif ex.method == "question_code_multi_col":
            data[ex.id] = extract_question_code_multi_col(
                df,
                code=params["code"],
                columns=params["columns"],
                code_col=sheet_cfg.code_col,
                desc_col=sheet_cfg.desc_col,
                max_rows=params.get("max_rows", 20),
                label_fn=label_fn,
                converter=converter,
                min_diff=params.get("min_diff", 0),
            )

        elif ex.method == "nested_ordinal":
            data[ex.id] = extract_nested_ordinal(
                df,
                row_start=params["row_start"],
                row_end=params["row_end"],
                code_col=params.get("code_col", 0),
                desc_col=params.get("desc_col", 1),
                ordinal_col=params.get("ordinal_col", 2),
                ordinals=params.get("ordinals", ["1st", "2nd", "3rd", "4th"]),
                q_prior_col=sheet_cfg.q_prior_col,
                q_current_col=sheet_cfg.q_current_col,
                label_fn=label_fn,
                converter=converter,
            )

        elif ex.method == "raw_aggregate":
            # Deferred — handled after main extraction loop (needs raw_data)
            pass

        elif ex.method == "synapse_report":
            # Deferred — handled in load_all_data() via synapse_json_loader
            pass

        else:
            print(f"  [WARN] Unknown extraction method: {ex.method} for {ex.id}")

        # Warn if extraction returned no rows
        if ex.id in data and isinstance(data[ex.id], list) and len(data[ex.id]) == 0:
            logger.warning("Extraction '%s' (method=%s) returned 0 rows", ex.id, ex.method)

    # Attach raw sheet index for discovery by Stage 4 / config generation
    data["_sheets"] = raw_index

    return data


# ── Master loader ────────────────────────────────────────────────────────────

_STALENESS_THRESHOLD_HOURS = 24


def load_all_data(config, force_fresh: bool = False) -> dict:
    """Load all data extractions — from JSON if available, else from Excel.

    Flow:
      1. Check for source_data.json next to the Excel file
      2. If found and Excel unchanged → load JSON (fast, no pandas)
      3. If missing or stale → extract from Excel → save source_data.json

    Args:
        config: ProjectConfig instance
        force_fresh: If True, skip JSON cache and re-extract from Excel.

    Returns:
        dict mapping extraction.id → list[dict] (includes _meta and _sheets)
    """
    data = None

    if not force_fresh:
        # Try JSON first
        data = _load_source_json(config)
    else:
        print("  --fresh: forcing re-extraction from Excel")

    if data is None:
        # Extract from Excel and save JSON for next time
        print("  Extracting from Excel...")
        data = _extract_all_from_excel(config)
        _save_source_json(data, config)

    # ── Staleness reporting (PRD §9.3) ──
    meta = data.get("_meta", {})
    extracted_at = meta.get("extracted_at", "")
    if extracted_at:
        try:
            from datetime import datetime as _dt
            pull_time = _dt.fromisoformat(extracted_at)
            age_hours = (datetime.now() - pull_time).total_seconds() / 3600
            if age_hours > _STALENESS_THRESHOLD_HOURS:
                print(
                    f"  [STALE] Data extracted {age_hours:.0f}h ago "
                    f"({extracted_at[:16]}). "
                    f"Use force_fresh=True to re-extract."
                )
        except (ValueError, TypeError):
            pass

    # Always inject mock extractions (not stored in JSON)
    for ex in config.extractions:
        if ex.method == "mock":
            data[ex.id] = ex.params.get("rows", [])

    # ── Synapse JSON report extractions ──
    synapse_extractions = [ex for ex in config.extractions if ex.method == "synapse_report"]
    if synapse_extractions:
        from slidegen.pipeline.synapse_auth import resolve_api_key
        try:
            api_key = resolve_api_key()
        except ValueError:
            api_key = ""
        if api_key:
            from slidegen.pipeline.synapse_json_loader import fetch_data_as_json
            try:
                synapse_data = fetch_data_as_json(config, api_key=api_key)
                for ex in synapse_extractions:
                    if ex.id in synapse_data:
                        data[ex.id] = synapse_data[ex.id]
                        print(f"  {ex.id}: {len(synapse_data[ex.id])} rows (synapse_report)")
            except Exception as e:
                logger.warning("Synapse JSON fetch failed: %s — falling back to Excel", e)
                # Fallback: skip, data may already be in JSON cache from prior Excel extraction
        else:
            logger.info(
                "No Synapse API token available — synapse_report extractions skipped, using Excel only"
            )

    # ── Raw data aggregation (respondent-level) ──
    raw_extractions = [ex for ex in config.extractions if ex.method == "raw_aggregate"]
    if raw_extractions:
        from slidegen.pipeline.raw_data_loader import (
            load_raw_data, aggregate_raw, aggregate_raw_multi_code,
            aggregate_raw_cross_brand, aggregate_raw_by_segment,
            aggregate_raw_by_hii, validate_raw_extractions,
        )

        raw_data = load_raw_data(config)
        if raw_data is None:
            logger.warning("raw_data_source_path not set or file missing — skipping raw_aggregate extractions")
        else:
            # Validate all raw extractions upfront
            warnings = validate_raw_extractions(config, raw_data)
            for w in warnings:
                print(f"  [WARN] {w}")
            # Build global label shortener
            global_shortener = None
            if config.label_shortcuts:
                shortcuts = [{"keywords": ls.keywords, "short": ls.short} for ls in config.label_shortcuts]
                global_shortener = make_label_shortener(shortcuts)

            for ex in raw_extractions:
                params = ex.params
                label_fn = None
                if params.get("use_label_shortcuts") and global_shortener:
                    label_fn = global_shortener
                elif "label_shortcuts" in params:
                    label_fn = make_label_shortener(params["label_shortcuts"])

                sheet_key = params.get("raw_sheet", ex.sheet)
                raw_mode = params.get("mode", "single")

                if raw_mode == "multi_code":
                    result = aggregate_raw_multi_code(
                        raw_data, sheet_key=sheet_key,
                        codes=params["codes"],
                        agg=params.get("agg", "top2box"),
                        quarter_current=params.get("quarter_current", config.period_current),
                        quarter_prior=params.get("quarter_prior", config.period_prior),
                    )

                elif raw_mode == "cross_brand":
                    result = aggregate_raw_cross_brand(
                        raw_data,
                        primary_sheet=params.get("primary_sheet", "primary"),
                        comp_sheet=params.get("comp_sheet", "competitor"),
                        code=params["code"],
                        agg=params.get("agg", "top2box"),
                        quarter_current=params.get("quarter_current", config.period_current),
                        quarter_prior=params.get("quarter_prior", config.period_prior),
                        label_fn=label_fn,
                    )

                elif raw_mode == "by_segment":
                    result = aggregate_raw_by_segment(
                        raw_data, sheet_key=sheet_key,
                        code=params["code"],
                        segment_code=params["segment_code"],
                        segment_values=params["segment_values"],
                        agg=params.get("agg", "top2box"),
                        quarter=params.get("quarter_current", config.period_current),
                        label_fn=label_fn,
                    )

                elif raw_mode == "by_hii":
                    result = aggregate_raw_by_hii(
                        raw_data, sheet_key=sheet_key,
                        code=params["code"],
                        quality_code=params.get("quality_code", "Q1_87Z"),
                        ltip_code=params.get("ltip_code", "C1_85DZ"),
                        agg=params.get("agg", "top2box"),
                        quarter=params.get("quarter_current", config.period_current),
                        label_fn=label_fn,
                        skip_zero=params.get("skip_zero", False),
                    )

                else:  # single (default)
                    result = aggregate_raw(
                        raw_data, sheet_key=sheet_key,
                        code=params["code"],
                        agg=params.get("agg", "top2box"),
                        quarter_current=params.get("quarter_current", config.period_current),
                        quarter_prior=params.get("quarter_prior", config.period_prior),
                        label_fn=label_fn,
                        skip_zero=params.get("skip_zero", False),
                    )

                data[ex.id] = result
                print(f"  {ex.id}: {len(result)} rows (raw_aggregate/{raw_mode})")

    # ── Synapse raw data fetching (respondent-level via API) ──
    synapse_raw_extractions = [ex for ex in config.extractions if ex.method == "synapse_raw"]
    if synapse_raw_extractions:
        from slidegen.pipeline.synapse_auth import resolve_api_key as _resolve_raw_key
        try:
            api_key = _resolve_raw_key()
        except ValueError:
            api_key = ""
        if api_key and config.synapse:
            from slidegen.pipeline.synapse_raw_fetcher import fetch_all_raw, load_cached_pkl
            from slidegen.pipeline.raw_data_loader import (
                dataframes_to_raw_data, merge_vq_data, apply_segment_filter,
                aggregate_raw, aggregate_raw_multi_code,
                aggregate_raw_cross_brand, aggregate_raw_by_segment,
            )

            try:
                # Try pkl cache first, else fetch fresh from Synapse JSON API
                cached = load_cached_pkl(config)
                if cached and cached.get("dataframes"):
                    raw_result = cached
                    raw_result.setdefault("vq_data", {})
                    raw_result.setdefault("vq_questions", [])
                    raw_result.setdefault("segments", [])
                else:
                    raw_result = fetch_all_raw(config, api_key=api_key)

                # Convert DataFrames → dict format for aggregation
                raw_data = dataframes_to_raw_data(
                    raw_result.get("dataframes", {}),
                    header_metadata=raw_result.get("_header_metadata"),
                )

                if raw_data:
                    # Merge VQ data into respondent data
                    if raw_result.get("vq_data"):
                        for sheet_key in list(raw_data.keys()):
                            if not sheet_key.startswith("_"):
                                merge_vq_data(
                                    raw_data, raw_result["vq_data"],
                                    raw_result.get("vq_questions", []),
                                    sheet_key=sheet_key,
                                )

                    # Apply segment filters from config
                    if config.synapse and config.synapse.segments:
                        segment_cuts = [
                            {"id": sc.id, "name": sc.name, "mode": sc.mode,
                             "values": sc.values, "segment_code": sc.name}
                            for sc in config.synapse.segments
                        ]
                        for sheet_key in list(raw_data.keys()):
                            if not sheet_key.startswith("_"):
                                raw_data = apply_segment_filter(
                                    raw_data, segment_cuts, sheet_key=sheet_key
                                )

                    # Build global label shortener
                    global_shortener_raw = None
                    if config.label_shortcuts:
                        shortcuts = [{"keywords": ls.keywords, "short": ls.short}
                                     for ls in config.label_shortcuts]
                        global_shortener_raw = make_label_shortener(shortcuts)

                    # Run synapse_raw extractions
                    for ex in synapse_raw_extractions:
                        params = ex.params
                        label_fn = None
                        if params.get("use_label_shortcuts") and global_shortener_raw:
                            label_fn = global_shortener_raw
                        elif "label_shortcuts" in params:
                            label_fn = make_label_shortener(params["label_shortcuts"])

                        sheet_key = params.get("raw_sheet", ex.sheet)
                        raw_mode = params.get("mode", "single")

                        if raw_mode == "multi_code":
                            result = aggregate_raw_multi_code(
                                raw_data, sheet_key=sheet_key,
                                codes=params["codes"],
                                agg=params.get("agg", "top2box"),
                                quarter_current=params.get("quarter_current", config.period_current),
                                quarter_prior=params.get("quarter_prior", config.period_prior),
                            )
                        elif raw_mode == "cross_brand":
                            result = aggregate_raw_cross_brand(
                                raw_data,
                                primary_sheet=params.get("primary_sheet", "primary"),
                                comp_sheet=params.get("comp_sheet", "competitor"),
                                code=params["code"],
                                agg=params.get("agg", "top2box"),
                                quarter_current=params.get("quarter_current", config.period_current),
                                quarter_prior=params.get("quarter_prior", config.period_prior),
                                label_fn=label_fn,
                            )
                        elif raw_mode == "by_segment":
                            result = aggregate_raw_by_segment(
                                raw_data, sheet_key=sheet_key,
                                code=params["code"],
                                segment_code=params["segment_code"],
                                segment_values=params["segment_values"],
                                agg=params.get("agg", "top2box"),
                                quarter=params.get("quarter_current", config.period_current),
                                label_fn=label_fn,
                            )
                        else:  # single (default)
                            result = aggregate_raw(
                                raw_data, sheet_key=sheet_key,
                                code=params["code"],
                                agg=params.get("agg", "top2box"),
                                quarter_current=params.get("quarter_current", config.period_current),
                                quarter_prior=params.get("quarter_prior", config.period_prior),
                                label_fn=label_fn,
                            )

                        data[ex.id] = result
                        print(f"  {ex.id}: {len(result)} rows (synapse_raw/{raw_mode})")

            except Exception as e:
                logger.warning("Synapse raw fetch failed: %s — skipping synapse_raw extractions", e)
        else:
            if not api_key:
                logger.info("No Synapse API token available — synapse_raw extractions skipped")
            if not config.synapse:
                logger.info("No synapse config — synapse_raw extractions skipped")

    # Always attach sample sizes from config (not stored in JSON)
    data["_sample_sizes"] = config.sample_sizes

    return data
