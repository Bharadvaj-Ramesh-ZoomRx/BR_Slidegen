"""
data_loaders.py — Data loading, JSON caching, and Excel extraction orchestration.

Coordinates the three data tiers: aggregated Excel, respondent-level local,
and Synapse API. Extraction functions have been moved to extractors.py;
this module handles caching, index building, and the master load_all_data().
"""

from __future__ import annotations
import hashlib
import json
import logging
import os
import tempfile
import pandas as pd
from datetime import datetime
from typing import Optional, Callable

# Re-export extraction functions for backward compatibility
from slidegen.pipeline.extractors import (  # noqa: F401
    pct, straight, delta,
    make_label_shortener, _build_brand_replacements,
    extract_by_question_code, extract_multi_question_code,
    extract_row_range, extract_question_code_multi_col,
    extract_nested_ordinal,
)

logger = logging.getLogger(__name__)


# Extractors moved to extractors.py — import for backward compat
from slidegen.pipeline.extractors import (  # noqa: F401, E402
    pct, straight, delta,
    extract_by_question_code, extract_multi_question_code,
    extract_row_range, extract_question_code_multi_col,
    extract_nested_ordinal, make_label_shortener,
)


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
            logger.info("  source_data.json up to date — %d sheets, %d codes indexed", len(existing['_sheets']), n_codes)
            return existing

    # Build raw sheet index from Excel
    logger.info("  Indexing Excel: %s", os.path.basename(excel_path))
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
        logger.info("    %s: %d rows%s", sheet_name, len(rows), skipped_msg)

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

    _atomic_json_write(json_path, payload)

    logger.info("  Saved: %s", json_path)
    return payload


# ── JSON auto-cache ───────────��─────────────────────────────────��───────────

def _atomic_json_write(path: str, payload: dict):
    """Write JSON atomically via temp-file + os.replace to prevent corruption."""
    dir_path = os.path.dirname(path)
    os.makedirs(dir_path, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=dir_path, suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2, default=str)
        os.replace(tmp, path)
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


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
            logger.info("  source_data.json stale (Excel changed) — re-extracting")
            return None

    # Invalidate if extraction config changed
    current_ext_hash = _extractions_hash(config)
    if meta.get("extractions_hash") and meta["extractions_hash"] != current_ext_hash:
        logger.info("  source_data.json stale (extraction params changed) — re-extracting")
        return None

    # Keep _meta alongside data (renderers ignore it; orchestrator uses extracted_at)
    data = dict(payload)
    n_extractions = sum(1 for k in data if not k.startswith("_"))

    # If JSON only has _sheets (index-only from Stage 0) but no extractions,
    # signal re-extraction needed
    if n_extractions == 0:
        logger.info("  source_data.json has index only (no extractions) — extracting from Excel")
        return None

    logger.info("  Loaded from JSON: %s (%d extractions)", json_path, n_extractions)
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

    _atomic_json_write(json_path, payload)

    logger.info("  Saved: %s", json_path)
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

    # Reuse _sheets and _codes from existing JSON if available and Excel hash
    # matches, avoiding a redundant scan that index_excel() already performed.
    raw_index = None
    codes_index = {}  # {sheet: {code: {value_range, ...}}} for pct_mode auto-detect
    json_path = _json_path_for(config)
    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                existing = json.load(f)
            meta = existing.get("_meta", {})
            if (meta.get("excel_hash") == _file_hash(config.data_source_path)
                    and "_sheets" in existing):
                raw_index = existing["_sheets"]
                logger.info("  Reusing _sheets index from existing source_data.json")
            codes_index = existing.get("_codes", {})
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

    # Build brand-specific label replacements from config
    brand_replacements = _build_brand_replacements(config)

    # Build global label shortener
    global_shortener = None
    if config.label_shortcuts:
        shortcuts = [{"keywords": ls.keywords, "short": ls.short} for ls in config.label_shortcuts]
        global_shortener = make_label_shortener(shortcuts, brand_replacements=brand_replacements)

    data = {}
    bases = {}  # extraction_id → {prior: int, current: int}

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
            label_fn = make_label_shortener(local_shortcuts, brand_replacements=brand_replacements)

        # Determine converter — explicit pct_mode wins; else auto-detect from _codes
        explicit_mode = params.get("pct_mode")
        if explicit_mode:
            converter = straight if explicit_mode == "straight" else pct
        else:
            # Auto-detect from _codes value_range metadata (if available)
            code_key = params.get("code", "")
            sheet_codes = codes_index.get(ex.sheet, {})
            vr = sheet_codes.get(code_key, {}).get("value_range", "")
            if vr == "whole":
                converter = straight
                logger.debug("  Auto pct_mode=straight for %s (value_range=whole)", ex.id)
            else:
                converter = pct

        if ex.method == "question_code":
            # Per-extraction desc_col override (e.g. VA/Starter Kit labels in col 0)
            effective_desc_col = params.get("desc_col", sheet_cfg.desc_col)
            data[ex.id] = extract_by_question_code(
                df,
                code=params["code"],
                q_prior_col=sheet_cfg.q_prior_col,
                q_current_col=sheet_cfg.q_current_col,
                code_col=sheet_cfg.code_col,
                desc_col=effective_desc_col,
                max_rows=params.get("max_rows", 20),
                label_fn=label_fn,
                converter=converter,
                occurrence=params.get("occurrence", 1),
                dim_col=params.get("dim_col"),
            )

            # Capture base sizes if available
            rows = data[ex.id]
            if hasattr(rows, '_base'):
                bases[ex.id] = rows._base
                data[ex.id] = list(rows)  # convert to plain list for JSON

            # Post-extraction aggregation
            aggregate = params.get("aggregate")
            if aggregate == "avg_by_desc":
                data[ex.id] = _aggregate_avg_by_desc(data[ex.id])
            elif aggregate == "mbd_pivot":
                data[ex.id] = _aggregate_mbd_pivot(data[ex.id])

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
            logger.warning("  [WARN] Unknown extraction method: %s for %s", ex.method, ex.id)

        # Warn if extraction returned no rows
        if ex.id in data and isinstance(data[ex.id], list) and len(data[ex.id]) == 0:
            logger.warning("Extraction '%s' (method=%s) returned 0 rows", ex.id, ex.method)

    # Attach raw sheet index for discovery by Stage 4 / config generation
    data["_sheets"] = raw_index

    # Attach base sizes extracted from question code header rows
    if bases:
        data["_bases"] = bases

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
        logger.info("  --fresh: forcing re-extraction from Excel")

    if data is None:
        # Extract from Excel and save JSON for next time
        logger.info("  Extracting from Excel...")
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
            threshold = getattr(config, 'staleness_hours', 0) or _STALENESS_THRESHOLD_HOURS
            if age_hours > threshold:
                logger.warning(
                    "[STALE] Data extracted %.0fh ago (%s). "
                    "Use force_fresh=True to re-extract.",
                    age_hours, extracted_at[:16],
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
                        logger.info("  %s: %d rows (synapse_report)", ex.id, len(synapse_data[ex.id]))
            except (RuntimeError, OSError, ValueError, KeyError) as e:
                logger.warning("Synapse JSON fetch failed: %s — falling back to Excel", e)
                # Fallback: skip, data may already be in JSON cache from prior Excel extraction
        else:
            logger.info(
                "No Synapse API token available — synapse_report extractions skipped, using Excel only"
            )

    # Build brand-specific label replacements for raw/synapse shorteners
    brand_replacements = _build_brand_replacements(config)

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
                logger.warning("  [WARN] %s", w)
            # Build global label shortener
            global_shortener = None
            if config.label_shortcuts:
                shortcuts = [{"keywords": ls.keywords, "short": ls.short} for ls in config.label_shortcuts]
                global_shortener = make_label_shortener(shortcuts, brand_replacements=brand_replacements)

            for ex in raw_extractions:
                params = ex.params
                label_fn = None
                if params.get("use_label_shortcuts") and global_shortener:
                    label_fn = global_shortener
                elif "label_shortcuts" in params:
                    label_fn = make_label_shortener(params["label_shortcuts"], brand_replacements=brand_replacements)

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
                logger.info("  %s: %d rows (raw_aggregate/%s)", ex.id, len(result), raw_mode)

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
                        global_shortener_raw = make_label_shortener(shortcuts, brand_replacements=brand_replacements)

                    # Run synapse_raw extractions
                    for ex in synapse_raw_extractions:
                        params = ex.params
                        label_fn = None
                        if params.get("use_label_shortcuts") and global_shortener_raw:
                            label_fn = global_shortener_raw
                        elif "label_shortcuts" in params:
                            label_fn = make_label_shortener(params["label_shortcuts"], brand_replacements=brand_replacements)

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
                        logger.info("  %s: %d rows (synapse_raw/%s)", ex.id, len(result), raw_mode)

            except (RuntimeError, OSError, ValueError, KeyError) as e:
                logger.warning("Synapse raw fetch failed: %s — skipping synapse_raw extractions", e)
        else:
            if not api_key:
                logger.info("No Synapse API token available — synapse_raw extractions skipped")
            if not config.synapse:
                logger.info("No synapse config — synapse_raw extractions skipped")

    # Always attach sample sizes from config (not stored in JSON)
    data["_sample_sizes"] = config.sample_sizes

    # ── Post-load completeness check ──
    _warnings: list[str] = []
    expected_ids = {ex.id for ex in config.extractions}
    loaded_ids = {k for k in data if not k.startswith("_")}
    missing = expected_ids - loaded_ids
    empty = {k for k in loaded_ids if isinstance(data.get(k), list) and len(data[k]) == 0}
    if missing:
        for m in sorted(missing):
            _warnings.append(f"Extraction '{m}' not loaded (missing from data)")
    if empty:
        for e in sorted(empty):
            _warnings.append(f"Extraction '{e}' returned 0 rows")
    data["_warnings"] = _warnings
    if _warnings:
        logger.warning("Data completeness: %d/%d extractions loaded, %d warnings:",
                        len(loaded_ids - empty), len(expected_ids), len(_warnings))
        for w in _warnings:
            logger.warning("  - %s", w)

    return data
