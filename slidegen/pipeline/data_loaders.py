"""
data_loaders.py — Generic data extraction functions.

All the repeated "find code, walk rows, extract Q3/Q4" patterns
from generate_asks.py are consolidated here into reusable extractors.
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
                    label = label_fn(desc) if label_fn else desc
                    results.append({
                        "desc": label,
                        "code": str(row_code),
                        "prior": converter(q_prior) if pd.notna(q_prior) else None,
                        "current": converter(q_current) if pd.notna(q_current) else None,
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
    Finds the code row, then looks for "top"/"yes" sub-row or takes first data row.
    """
    results = []
    for entry in codes:
        code = entry["code"]
        label = entry["label"]
        for i in range(len(sheet)):
            if sheet.iloc[i, code_col] == code:
                for j in range(i, min(i + 8, len(sheet))):
                    desc = str(sheet.iloc[j, desc_col]) if pd.notna(sheet.iloc[j, desc_col]) else ""
                    if "top" in desc.lower() or "yes" in desc.lower() or j == i + 1:
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
                    if all_valid:
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
            label = label_fn(desc) if label_fn else desc[:40]
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


# ── Master loader ────────────────────────────────────────────────────────────

def load_all_data(config) -> dict:
    """Load all data extractions defined in the project config.

    Args:
        config: ProjectConfig instance

    Returns:
        dict mapping extraction.id → list[dict]
    """
    # Load Excel sheets
    sheets_data = {}
    for sheet_key, sheet_cfg in config.sheets.items():
        sheets_data[sheet_key] = pd.read_excel(
            config.data_source_path,
            sheet_name=sheet_cfg.name,
            header=None,
        )

    # Build global label shortener
    global_shortener = None
    if config.label_shortcuts:
        shortcuts = [{"keywords": ls.keywords, "short": ls.short} for ls in config.label_shortcuts]
        global_shortener = make_label_shortener(shortcuts)

    data = {}

    for ex in config.extractions:
        sheet_cfg = config.sheets[ex.sheet]
        df = sheets_data[ex.sheet]
        params = ex.params

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

        else:
            print(f"  [WARN] Unknown extraction method: {ex.method} for {ex.id}")

        # Warn if extraction returned no rows
        if ex.id in data and isinstance(data[ex.id], list) and len(data[ex.id]) == 0:
            logger.warning("Extraction '%s' (method=%s) returned 0 rows", ex.id, ex.method)

    # Store sample sizes
    data["_sample_sizes"] = config.sample_sizes

    return data
