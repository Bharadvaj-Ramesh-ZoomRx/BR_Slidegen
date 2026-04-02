"""
extractors.py — Value converters, label shortening, and Excel extraction functions.

Split from data_loaders.py for maintainability. These are the pure extraction
functions that transform pandas DataFrames into standardized data dicts.
"""

from __future__ import annotations
import logging
import pandas as pd
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

def make_label_shortener(
    shortcuts: list[dict],
    max_len: int = 35,
    brand_replacements: Optional[dict[str, str]] = None,
) -> Callable[[str], str]:
    """Build a label shortening function from a list of keyword→short mappings.

    Each shortcut: {"keywords": ["kw1", "kw2"], "short": "Short Label"}

    brand_replacements: optional dict of long→short brand name substitutions,
        e.g. {"Johnson & Johnson (Formerly Janssen)": "J&J", "[COMPANY]": "J&J"}.
        Replaces the old hardcoded brand strings.
    """
    replacements = brand_replacements or {}

    def shorten(label: str) -> str:
        ll = label.lower()
        for sc in shortcuts:
            kws = sc["keywords"]
            if all(kw.lower() in ll for kw in kws):
                return sc["short"]
        # Apply brand-specific replacements from config
        for old, new in replacements.items():
            label = label.replace(old, new)
        label = label.replace("How well the ", "").replace("How ", "")
        label = label.strip()
        if len(label) > max_len:
            label = label[:max_len - 1] + "\u2026"
        return label

    return shorten


def _build_brand_replacements(config) -> dict[str, str]:
    """Build label replacement dict from config brand info.

    Maps common long company/product strings to their short forms
    using config.client and config.brands.
    """
    replacements = {}
    client = getattr(config, "client", "")
    if client:
        # Common long-form variants → short client name
        replacements[f"{client} (Formerly Janssen)"] = client
        replacements[f"{client.replace(' & ', ' &amp; ')}"] = client
        replacements["[COMPANY]"] = client
    primary = config.brands.get("primary") if hasattr(config, "brands") else None
    if primary:
        replacements["[PRODUCT]"] = primary.name
    return replacements


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
    occurrence: int = 1,
    dim_col: Optional[int] = None,
) -> list[dict]:
    """Find a question code row, walk sub-rows, extract prior/current values.

    Returns list of: {"desc": str, "prior": float, "current": float, "code": str}
    The question code header row's values are captured as _base (sample size).
    ``occurrence`` selects the Nth match when a code appears multiple times.
    ``dim_col`` adds a "dimension" field from the specified column.
    """
    results = []
    match_count = 0
    for i in range(len(sheet)):
        if sheet.iloc[i, code_col] == code:
            match_count += 1
            if match_count < occurrence:
                continue
            # Capture base/sample size from the question code row itself
            base_prior = sheet.iloc[i, q_prior_col]
            base_current = sheet.iloc[i, q_current_col]
            base_prior = int(base_prior) if pd.notna(base_prior) and base_prior else None
            base_current = int(base_current) if pd.notna(base_current) and base_current else None

            j = i + 1
            while j < len(sheet) and pd.notna(sheet.iloc[j, desc_col]):
                desc = str(sheet.iloc[j, desc_col]).strip()
                q_prior = sheet.iloc[j, q_prior_col]
                q_current = sheet.iloc[j, q_current_col]
                row_code = sheet.iloc[j, code_col] if pd.notna(sheet.iloc[j, code_col]) else ""

                if desc and not desc.startswith("Base") and pd.notna(q_current):
                    prior_val = converter(q_prior) if pd.notna(q_prior) else None
                    current_val = converter(q_current) if pd.notna(q_current) else None
                    # Skip rows where both prior and current are missing
                    if prior_val is None and current_val is None:
                        j += 1
                        continue
                    label = label_fn(desc) if label_fn else desc
                    row_dict = {
                        "desc": label,
                        "code": str(row_code),
                        "prior": prior_val,
                        "current": current_val,
                    }
                    if dim_col is not None:
                        dim_val = sheet.iloc[j, dim_col] if dim_col < sheet.shape[1] else None
                        if pd.notna(dim_val):
                            row_dict["dimension"] = str(dim_val).strip()
                    results.append(row_dict)
                j += 1
                if j - i > max_rows:
                    break
            results = _attach_base(results, base_prior, base_current)
            break
    return results


def _attach_base(rows: list[dict], base_prior, base_current) -> list[dict]:
    """Attach _base metadata to a results list (stored as list attribute)."""
    class _ResultsWithBase(list):
        pass
    out = _ResultsWithBase(rows)
    out._base = {"prior": base_prior, "current": base_current}
    return out


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
