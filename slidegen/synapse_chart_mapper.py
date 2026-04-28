"""
synapse_chart_mapper.py — Map Synapse API records to chart data using Connector tags.

Takes flat Synapse records + PivotConfig + MappingConfig from the Connector tags
and produces CategoryChartData that can be written into a chart via replace_data().

The Connector's tag model:
  - ReportConfig: what to fetch (project, analysis, segments, time periods)
  - PivotConfig (DataFrameConfigHash): how to pivot (row/col/value fields, filters)
  - MappingConfig: what goes into the chart (which pivoted columns become series)

Field name mapping (Connector internal → Synapse API):
  - "value" → "y_label" (answer/option text)
  - "segment_1" → "segment_1" (already matches)
  - "time_period_name" → "time_period_name" (already matches)
"""
from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from typing import Any, Optional

import pandas as pd
import requests

logger = logging.getLogger(__name__)

# Connector field names → Synapse API record column names
FIELD_MAP = {
    "value": "y_label",
    "y_code": "y_code",
    "y_label": "y_label",
    "segment_1": "segment_1",
    "segment_2": "segment_2",
    "time_period_name": "time_period_name",
    "time_period_id": "time_period_id",
    "question": "y_label",
    "title": "y_label",       # Connector "title" = API "y_label"
    "measure": "measure",
    "n": "base",              # Connector "n" = API "base" (sample size)
    "n - value": "base",      # Computed "n cross value" → use base
}


def _remap_field(name: str) -> str:
    return FIELD_MAP.get(name, name)


@dataclass
class ChartRefreshData:
    """Data ready to write into a chart via replace_data()."""
    categories: list[str]
    series: list[tuple[str, list[float]]]  # [(series_name, values), ...]
    success: bool = True
    error: str = ""


def fetch_synapse_report(
    base_url: str,
    token: str,
    report_config: dict,
) -> list[dict]:
    """Fetch report data from Synapse API using ReportConfig from Connector tag.

    Uses StaticTimePeriodIds when available (matches the exact data the chart
    was rendered with). Falls back to DynamicTimePeriod for fresh/latest data.
    """
    headers = {"Authorization": token, "accept": "application/json",
               "Content-Type": "application/json"}

    payload = {
        "project_id": report_config.get("ProjectId"),
        "reporting_plan_id": report_config.get("ReportingPlanId"),
        "analysis_ids": report_config.get("AnalysisIds"),
        "segment_ids": report_config.get("SegmentIds"),
    }

    # Use DynamicTimePeriod for refresh (gets latest data including new waves)
    dtp = report_config.get("DynamicTimePeriod")
    if dtp:
        payload["setup_type"] = "DYNAMIC"
        payload["dynamic_time_period"] = {
            "latest_n_deliverables": dtp.get("LatestNDeliverables", 8),
            "include_live_wave": dtp.get("IncludeLiveWave", True),
        }
    else:
        # No dynamic config — use static
        payload["setup_type"] = "DYNAMIC"
        payload["dynamic_time_period"] = {
            "latest_n_deliverables": 20,
            "include_live_wave": True,
        }

    try:
        resp = requests.post(
            f"{base_url}/api/reports/generate",
            headers=headers,
            json=payload,
            timeout=60,
        )
        if resp.status_code in (200, 201):
            return resp.json().get("records", [])
        else:
            logger.warning(f"Synapse API {resp.status_code}: {resp.text[:200]}")
            return []
    except Exception as e:
        logger.warning(f"Synapse API error: {e}")
        return []


def _resolve_value_field(val_field: str, df_columns: list[str]) -> str | None:
    """Find a value field in the DataFrame, case-insensitive."""
    if val_field in df_columns:
        return val_field
    lc = val_field.lower()
    for col in df_columns:
        if col.lower() == lc:
            return col
    return None


def _normalize_key(k: str) -> str:
    """Normalize compound column key: @:@ -> ' - ', collapse whitespace.

    Tag selectedColumns strings and pivot column values often disagree on
    whitespace — newlines, double spaces, NBSP. Normalize both sides to a
    single-space form so string matching succeeds.
    """
    import re
    k = k.replace(" @:@ ", " - ").replace("@:@", " - ")
    # Replace any whitespace run (incl. newlines, tabs, NBSP) with a single space
    k = re.sub(r"\s+", " ", k)
    return k.strip()


def _match_pivot_col(sc_norm: str, pivot_columns) -> str | None:
    """Find pivot column matching a normalized selectedColumn key.

    The Connector's selectedColumns use full segment paths like:
        "Q1'26 - Specialty (C/PCPs) - Overall - CARD - L - Average of decimal"
    But pivot columns use short values from the API:
        "Q1'26 - CARD - L"

    The key insight: every PART of the pivot column name (split by " - ") must
    appear SOMEWHERE in the selectedColumn string. "CARD" appears inside
    "Specialty (C/PCPs) - Overall - CARD". "Overall" appears inside "Overall Data".
    """
    # Pass 1: exact match
    for pc_col in pivot_columns:
        pc_norm = _normalize_key(str(pc_col))
        if sc_norm == pc_norm:
            return pc_col

    # Pass 2: each pivot part must be a substring of the selectedColumn.
    # When multiple pivot columns match, prefer the one whose LAST part
    # matches the LAST segment of the selectedColumn (prevents "Overall"
    # matching PCP's path "...Overall - PCP").
    sc_parts = [p.strip() for p in sc_norm.split(" - ")]
    sc_last = sc_parts[-1] if sc_parts else ""

    best_match = None
    best_score = 0
    best_last_match = False
    for pc_col in pivot_columns:
        pc_norm = _normalize_key(str(pc_col))
        pc_parts = [p.strip() for p in pc_norm.split(" - ")]

        if not pc_parts:
            continue

        matched_parts = sum(1 for pp in pc_parts if pp in sc_norm)
        if matched_parts < len(pc_parts):
            continue  # not all parts found

        # Check if last parts match (stronger signal)
        pc_last = pc_parts[-1]
        last_match = (pc_last == sc_last) or (pc_last in sc_last) or (sc_last in pc_last)

        # Prefer: (1) last-part match, (2) higher part count
        if last_match and not best_last_match:
            best_match = pc_col
            best_score = matched_parts
            best_last_match = True
        elif last_match == best_last_match and matched_parts > best_score:
            best_match = pc_col
            best_score = matched_parts

    return best_match


# ── Fix L+M: wave-pinned selectedColumns rewrite (universal) ──
# Connector tags freeze wave labels into selectedColumns when waves are the
# column axis. Pradeep's non-connected inference produces the same shape
# because it copies the source chart's column headers into selectedColumns.
# When a refresh fetches different waves, those pinned entries no longer
# match the pivot output and the chart silently falls back to source-restore.
#
# Two flavors are handled:
#   Fix L — standalone wave entries: ["y_label", "Wave 12", "Wave 13"]
#   Fix M — compound entries: ["Codes", "Feb'26 @:@ X", "Mar'26 @:@ X"]
#                                          ^wave-prefix^^suffix^
#
# Wave detection is value-driven first (s in known_wave_labels — populated
# by the refresh pipeline from the analysis's full wave set, ground-truth
# regardless of project's wave naming convention) and falls back to a
# pattern set for callers that don't supply known_wave_labels.
import re as _re_wave

_WAVE_LABEL_PATTERNS = (
    _re_wave.compile(r"^(?:Project )?Wave \d+$"),                    # Wave 12, Project Wave 12
    _re_wave.compile(r"^W\d+$"),                                      # W28 (short form)
    _re_wave.compile(r"^[A-Z][a-z]{2}'\d{2}$"),                       # Jan'26
    _re_wave.compile(r"^Q[1-4]'\d{2}$"),                              # Q1'26
    _re_wave.compile(r"^Q[1-4] \d{4}$"),                              # Q1 2026
)

_WAVE_TEMPLATE_SENTINEL = "\x00WAVE\x00"
_COMPOUND_SEP = " @:@ "
# Some non-Connector / Pradeep-inferred selectedColumns use " - " as the
# compound separator instead (e.g. "Overall Efficacy - HIT - Project Wave 13").
# We try " @:@ " first; if no wave parts emerge, we try " - " as a fallback.
_COMPOUND_SEP_ALT = " - "


def _is_wave_label(s: str, known_wave_labels: set[str] | None = None) -> bool:
    if not isinstance(s, str):
        return False
    if known_wave_labels and s in known_wave_labels:
        return True
    return any(p.match(s) for p in _WAVE_LABEL_PATTERNS)


def _rewrite_wave_pinned_selected_columns(
    selected: list,
    df: pd.DataFrame,
    known_wave_labels: set[str] | None = None,
) -> tuple[list, int]:
    """Replace wave-label entries (standalone or compound) with fetched waves.

    Walks selectedColumns. For each entry, splits on the Connector compound
    separator (" @:@ ") and detects wave-label parts. An entry with at least
    one wave part is dropped; in its place, one new entry per fetched wave
    is emitted, with the wave-part substituted in (other parts preserved).
    Entries with no wave parts are kept unchanged in their original order.

    Compound templates are deduplicated, so two source entries like
    `"Feb'26 @:@ X"` and `"Mar'26 @:@ X"` collapse to one template
    `"<wave> @:@ X"` and produce one new entry per fetched wave (not two
    × number_of_waves).

    Returns (rewritten_list, n_dropped). When n_dropped is 0 the function is
    a no-op. Wave entries are appended in chronological order by
    `time_period_id`.
    """
    if not selected or "time_period_name" not in df.columns:
        return list(selected) if selected else selected, 0

    if "time_period_id" in df.columns:
        fetched_waves = (df.groupby("time_period_name")["time_period_id"]
                         .max().sort_values().index.tolist())
    else:
        fetched_waves = list(dict.fromkeys(df["time_period_name"].astype(str)))
    fetched_waves = [str(w) for w in fetched_waves]
    if not fetched_waves:
        return list(selected), 0

    non_wave_entries: list = []
    ordered_templates: list[str] = []  # preserves first-seen order
    seen_templates: set[str] = set()
    n_dropped = 0

    # Track templates with their separator so we re-emit using the right one.
    # `ordered_templates` items are (separator, template_string) tuples.
    # `seen_templates` mirrors that for dedup.
    ordered_templates: list[tuple[str, str]] = []  # (sep, template)
    seen_templates: set[tuple[str, str]] = set()

    for entry in selected:
        if not isinstance(entry, str):
            non_wave_entries.append(entry)
            continue
        # Try " @:@ " first (the Connector default), then " - " (Pradeep-
        # inferred / non-Connector decks). Pick whichever yields wave parts.
        # Fall through to whole-entry standalone check if neither does.
        chosen_sep: str | None = None
        chosen_parts: list[str] = []
        wave_idxs: list[int] = []
        for try_sep in (_COMPOUND_SEP, _COMPOUND_SEP_ALT):
            if try_sep not in entry:
                continue
            try_parts = entry.split(try_sep)
            try_wave_idxs = [i for i, p in enumerate(try_parts)
                             if _is_wave_label(p, known_wave_labels)]
            if try_wave_idxs:
                chosen_sep = try_sep
                chosen_parts = try_parts
                wave_idxs = try_wave_idxs
                break
        if not wave_idxs:
            # No compound wave parts. Standalone wave-label?
            if _is_wave_label(entry, known_wave_labels):
                chosen_sep = ""  # empty sep == standalone (single part)
                chosen_parts = [entry]
                wave_idxs = [0]
            else:
                non_wave_entries.append(entry)
                continue
        # Build template with wave-part(s) replaced by sentinel
        tmpl_parts = list(chosen_parts)
        for i in wave_idxs:
            tmpl_parts[i] = _WAVE_TEMPLATE_SENTINEL
        # Standalone (chosen_sep == "") joins to just the sentinel
        template = (chosen_sep.join(tmpl_parts) if chosen_sep
                    else tmpl_parts[0])
        key = (chosen_sep or "", template)
        if key not in seen_templates:
            ordered_templates.append(key)
            seen_templates.add(key)
        n_dropped += 1

    if n_dropped == 0:
        return list(selected), 0

    rewritten: list = list(non_wave_entries)
    for _sep, template in ordered_templates:
        for w in fetched_waves:
            rewritten.append(template.replace(_WAVE_TEMPLATE_SENTINEL, w))

    logger.info(
        "Fix L/M: dropped %d wave-pinned selectedColumns entries, expanded "
        "%d template(s) over %d fetched waves",
        n_dropped, len(ordered_templates), len(fetched_waves),
    )
    return rewritten, n_dropped


def pivot_records_to_chart_data(
    records: list[dict],
    pivot_config: dict,
    mapping_config: dict,
    static_time_period_names: list[str] | None = None,
    chart_pattern: str = "",
    split_order: int | None = None,
    rows_per_object: int | None = None,
    top_n_rows: int | None = None,
    static_time_period_ids: list[int] | None = None,
    known_wave_labels: set[str] | None = None,
    source_categories: list[str] | None = None,
    source_series_names: list[str] | None = None,
) -> ChartRefreshData:
    """Transform flat Synapse records into chart categories + series.

    Faithfully replicates the Connector's transformation logic using the raw
    PivotConfig and MappingConfig from the shape tags.

    Key Connector behaviors replicated:
      1. Compound row index (multiple RowFields) with selectedColumns[0] as display field
      2. @:@ separator normalization in selectedColumns compound keys
      3. applyTranspose from MappingConfig (explicit transpose for trended charts)
      4. selectedRows for row filtering + ordering
      5. columnDefinitions.sortCriteria.CustomList for explicit sort
      6. PivotConfig.Filters with filterCriteria (1=include, 2=exclude)
      7. Case-insensitive ValueField resolution with fallbacks
    """
    if not records:
        return ChartRefreshData(categories=[], series=[], success=False,
                                error="No records from Synapse")

    df = pd.DataFrame(records)

    # ── Wave-label normalization ──
    # Some Synapse projects expose waves as "Project Wave N" while source decks
    # display the shortened "Wave N". Since the underlying time_period_id is the
    # same record, stripping the "Project " prefix is semantically safe — it's a
    # display-label reconciliation, not a data change.
    if "time_period_name" in df.columns:
        df = df.copy()
        df["time_period_name"] = df["time_period_name"].astype(str).str.replace(
            r"^Project Wave ", "Wave ", regex=True
        )

    # Build name -> time_period_id map for chronological sorting later.
    # Uses max id per name so renames that kept the id still sort correctly.
    name_to_tp_id: dict[str, int] = {}
    if "time_period_name" in df.columns and "time_period_id" in df.columns:
        for name, group in df.groupby("time_period_name"):
            try:
                name_to_tp_id[str(name)] = int(group["time_period_id"].max())
            except (ValueError, TypeError):
                pass

    # ── Filter to static time periods by ID (preferred) or name ──
    # ID-based filter is safer than name-based because wave renames don't break it.
    if static_time_period_ids and "time_period_id" in df.columns:
        filtered = df[df["time_period_id"].isin(static_time_period_ids)]
        if not filtered.empty:
            df = filtered

    raw_row_fields = pivot_config.get("RowFields", [])
    raw_col_fields = pivot_config.get("ColumnFields", [])
    raw_val_fields = pivot_config.get("ValueFields", [])
    row_fields = [_remap_field(f) for f in raw_row_fields]
    col_fields = [_remap_field(f) for f in raw_col_fields]
    val_fields = [_remap_field(f) for f in raw_val_fields]

    # ── Field-name fallback for cross-tab analyses ──
    # The Connector tag names breakout dimensions "segment_1"/"segment_2" and
    # "options", but cross-tab Synapse analyses return those axes as "x_label"/
    # "x_code". When the requested field is absent but x_label is present,
    # substitute. The "options" substitution is scoped to avoid creating a
    # collision with the other axis — if both RowFields and ColumnFields would
    # end up claiming x_label, leave "options" unsubstituted so the mapper
    # returns success=False and the refresh pipeline preserves the source chart.
    def _substitute_missing(field_list: list[str], *, allow_options: bool) -> list[str]:
        out = []
        for f in field_list:
            if f in df.columns:
                out.append(f)
                continue
            if f in ("segment_1", "segment_2") and "x_label" in df.columns and "x_label" not in out:
                out.append("x_label")
                continue
            if f == "options" and allow_options and "x_label" in df.columns and "x_label" not in out:
                out.append("x_label")
                continue
            out.append(f)  # keep unchanged; downstream handles missing
        return out

    # Decide whether "options" substitution is safe on each axis. It would
    # collide if the OTHER axis already contains x_label or would gain it via
    # segment_1/segment_2 substitution.
    def _would_claim_x_label(fields: list[str]) -> bool:
        for f in fields:
            if f == "x_label":
                return True
            if f in ("segment_1", "segment_2") and f not in df.columns and "x_label" in df.columns:
                return True
        return False

    row_allow_opts = not _would_claim_x_label(col_fields)
    col_allow_opts = not _would_claim_x_label(row_fields)
    row_fields = _substitute_missing(row_fields, allow_options=row_allow_opts)
    col_fields = _substitute_missing(col_fields, allow_options=col_allow_opts)

    # ── Collision-avoidance for cross-tab analyses ──
    # Multiple Connector field names remap to "y_label" (e.g. "question", "value",
    # "title"). When both RowFields and ColumnFields resolve to the same Synapse
    # column, the pivot degenerates into a diagonal self-matrix. In cross-tab
    # analyses the second axis lives in x_label / x_code — substitute when
    # available. Preserves label↔label and code↔code alignment.
    row_set = set(row_fields)
    new_col_fields: list[str] = []
    for cf in col_fields:
        if cf in row_set:
            if cf == "y_label" and "x_label" in df.columns:
                new_col_fields.append("x_label")
                continue
            if cf == "y_code" and "x_code" in df.columns:
                new_col_fields.append("x_code")
                continue
        new_col_fields.append(cf)
    col_fields = new_col_fields

    if not row_fields or not val_fields:
        return ChartRefreshData(categories=[], series=[], success=False,
                                error="Missing RowFields or ValueFields")

    # ── Resolve value field (case-insensitive + fallbacks) ──
    primary_val = None
    for vf in val_fields:
        resolved = _resolve_value_field(vf, list(df.columns))
        if resolved:
            primary_val = resolved
            break
    if not primary_val:
        for fallback in ["percentage", "decimal", "value"]:
            if fallback in df.columns:
                primary_val = fallback
                break
    if not primary_val:
        return ChartRefreshData(categories=[], series=[], success=False,
                                error=f"ValueField '{val_fields}' not in records")

    # ── Resolve row fields — support compound index ──
    valid_row_fields = [rf for rf in row_fields if rf in df.columns]
    if not valid_row_fields:
        return ChartRefreshData(categories=[], series=[], success=False,
                                error=f"RowFields '{row_fields}' not in records")
    primary_row = valid_row_fields[0]

    # ── Apply PivotConfig.Filters ──
    # Connector encodes multi-value IN filters by joining values with ";".
    # When a semicolon is present we split and use isin() for inclusion
    # (or ~isin() for exclusion). Falls back to the original single-value
    # / first-word-contains behavior for atomic values.
    for filt in pivot_config.get("Filters", []):
        col_key = _remap_field(filt.get("ColumnKey", ""))
        fval = filt.get("Value", "")
        criteria = filt.get("filterCriteria", 1)
        if col_key and fval and col_key in df.columns:
            if ";" in fval:
                values = [v.strip() for v in fval.split(";") if v.strip()]
                if criteria == 2:
                    df = df[~df[col_key].isin(values)]
                else:
                    df = df[df[col_key].isin(values)]
            elif criteria == 2:
                df = df[df[col_key] != fval]
            else:
                if fval in df[col_key].values:
                    df = df[df[col_key] == fval]
                else:
                    mask = df[col_key].astype(str).str.contains(
                        fval.split()[0], case=False, na=False)
                    if mask.any():
                        df = df[mask]

    if df.empty:
        df = pd.DataFrame(records)

    # ── Filter to static time periods ──
    if static_time_period_names and "time_period_name" in df.columns:
        filtered = df[df["time_period_name"].isin(static_time_period_names)]
        if not filtered.empty:
            df = filtered

    # ── Build compound column key from ColumnFields ──
    # Connector format: "{col_field_values joined by ' @:@ '}"
    # With multiple ValueFields: creates separate columns per value field:
    #   "{col_prefix} @:@ Average of reach", "{col_prefix} @:@ Average of sov", etc.
    valid_col_fields = [f for f in col_fields if f in df.columns]
    df = df.copy()
    # Replace NaN with "(blank)" in column fields (Connector displays null as "(blank)")
    for cf in valid_col_fields:
        df[cf] = df[cf].fillna("(blank)")
    # Connector always uses "Average of" in column NAMES regardless of AggregationType.
    # (AggregationType controls the aggregation function, not the display prefix.)
    # Column alias may override display but Names are always "Average of {field}".
    agg_prefix = "Average of"

    # Resolve ALL available value fields (for multi-value pivots)
    resolved_val_fields = []
    for vf in val_fields:
        resolved = _resolve_value_field(vf, list(df.columns))
        if resolved:
            resolved_val_fields.append((vf, resolved))  # (original_name, df_column)
    if not resolved_val_fields:
        resolved_val_fields = [(primary_val, primary_val)]
    has_multi_vals = len(resolved_val_fields) > 1

    # Build column prefix from ColumnFields
    if len(valid_col_fields) >= 2:
        df["_col_prefix"] = df[valid_col_fields[0]].astype(str)
        for cf in valid_col_fields[1:]:
            df["_col_prefix"] = df["_col_prefix"] + " @:@ " + df[cf].astype(str)
    elif len(valid_col_fields) == 1:
        df["_col_prefix"] = df[valid_col_fields[0]].astype(str)
    else:
        df["_col_prefix"] = "value"

    if not has_multi_vals:
        # Single value field — column key = prefix only
        df["_col_key"] = df["_col_prefix"]
    else:
        # Multi-value: key includes aggregation suffix (pivot separately below)
        df["_col_key"] = df["_col_prefix"] + " @:@ " + agg_prefix + " " + primary_val

    # ── Determine display field for categories ──
    # For compound RowFields (e.g. ['y_label','alias5540','alias3486']),
    # selectedColumns lists the row fields to include, then series columns.
    # The LAST RowField in selectedColumns is the display field for categories.
    # E.g. selectedCols=['y_label','alias5540','Q1 2026'] → display alias5540 values.
    # Apply Fix L+M: rewrite wave-pinned entries (standalone or compound)
    # against the actually-fetched waves before any downstream consumer reads
    # selectedColumns. Detection is value-driven via known_wave_labels (the
    # full wave set for this analysis, supplied by the refresh pipeline) with
    # a regex fallback. No-op when selectedColumns has no wave-label entries.
    selected, _n_wave_rewrites = _rewrite_wave_pinned_selected_columns(
        mapping_config.get("selectedColumns", []), df,
        known_wave_labels=known_wave_labels,
    )
    display_row_field = primary_row  # default: first RowField

    if selected and len(valid_row_fields) > 1:
        # Find which selectedColumns entries match RowFields
        row_fields_in_sel = []
        for sc in selected:
            sc_remapped = _remap_field(sc)
            if sc_remapped in valid_row_fields:
                row_fields_in_sel.append(sc_remapped)
            elif sc in raw_row_fields:
                mapped = _remap_field(sc)
                if mapped in df.columns:
                    row_fields_in_sel.append(mapped)
        # Use the LAST matching RowField as display field
        if row_fields_in_sel:
            display_row_field = row_fields_in_sel[-1]
    elif selected and len(valid_row_fields) == 1:
        # Single RowField — check if selectedColumns[0] matches
        sc0_remapped = _remap_field(selected[0])
        if sc0_remapped in valid_row_fields:
            display_row_field = sc0_remapped

    # ── Pivot ──
    # Use compound row index when multiple RowFields exist (preserves duplicates
    # where display_field has repeated values but other row fields differ).
    # After pivot, extract display_field values as category labels.
    use_compound = len(valid_row_fields) > 1 and display_row_field != valid_row_fields[0]
    pivot_index = valid_row_fields if use_compound else display_row_field

    try:
        if has_multi_vals:
            pivot_frames = []
            for orig_name, df_col in resolved_val_fields:
                col_key = df["_col_prefix"] + " @:@ " + agg_prefix + " " + df_col
                sub = df.copy()
                sub["_col_key"] = col_key
                sub_pivot = sub.pivot_table(
                    index=pivot_index,
                    columns="_col_key",
                    values=df_col,
                    aggfunc="first",
                )
                pivot_frames.append(sub_pivot)
            pivot = pd.concat(pivot_frames, axis=1)
        else:
            pivot = df.pivot_table(
                index=pivot_index,
                columns="_col_key",
                values=primary_val,
                aggfunc="first",
            )
    except Exception as e:
        return ChartRefreshData(categories=[], series=[], success=False,
                                error=f"Pivot failed: {e}")

    # Extract display labels from compound index
    if use_compound and isinstance(pivot.index, pd.MultiIndex):
        display_idx = valid_row_fields.index(display_row_field)
        pivot.index = [idx[display_idx] if isinstance(idx, tuple) else idx
                       for idx in pivot.index]

    # ── Step 1: Filter COLUMNS by selectedColumns (Connector: GetSelectedColumnsInOrder) ──
    # This happens BEFORE transpose. Removes columns not in selectedColumns.
    row_field_names = set(_remap_field(f) for f in raw_row_fields) | set(raw_row_fields)
    series_columns = []  # non-row, non-blank selectedColumns entries
    for sc in selected:
        if sc.startswith("<blank:"):
            continue
        sc_remapped = _remap_field(sc)
        if sc_remapped in row_field_names or sc in row_field_names:
            continue
        series_columns.append(sc)

    if series_columns:
        # Keep only columns that match selectedColumns entries
        cols_to_keep = []
        for sc in series_columns:
            sc_norm = _normalize_key(sc)
            matched = _match_pivot_col(sc_norm, pivot.columns)
            if matched is not None:
                cols_to_keep.append(matched)
        if cols_to_keep:
            pivot = pivot[cols_to_keep]

    # ── Step 1b: Apply topNRows + rowsPerObject/splitOrder (split visualization) ──
    # Connector's SplitVisualizationService splits pivot rows across chart shapes.
    # topNRows: limit total rows. rowsPerObject=1 + splitOrder=K: take row K only.
    if top_n_rows and top_n_rows > 0 and len(pivot) > top_n_rows:
        pivot = pivot.iloc[:top_n_rows]

    if rows_per_object and rows_per_object > 0 and split_order is not None:
        start = split_order * rows_per_object
        end = start + rows_per_object
        if start < len(pivot):
            pivot = pivot.iloc[start:min(end, len(pivot))]

    # ── Step 2: Apply transpose (Connector: VisualizationService at render time) ──
    # After transpose, column names become categories.
    # Connector applies columnAliasMap: default alias replaces @:@ with -.
    # Then explicit columnDefinitions[].Alias overrides if set.
    apply_transpose = mapping_config.get("applyTranspose", False)
    if apply_transpose:
        # Build alias map: {raw_name: display_alias}
        col_alias_map = {}
        for cd in pivot_config.get("columnDefinitions", []):
            name = cd.get("Name", "")
            alias = cd.get("Alias")
            if name and alias:
                col_alias_map[name] = alias

        pivot = pivot.T

        # Apply aliases: explicit alias if available, else default (replace @:@ with -)
        new_index = []
        for idx in pivot.index:
            idx_str = str(idx)
            if idx_str in col_alias_map:
                new_index.append(col_alias_map[idx_str])
            else:
                # Default alias: @:@ -> -
                new_index.append(idx_str.replace(" @:@ ", " - "))
        pivot.index = new_index

    # ── Step 3: Extract categories + series from the (possibly transposed) pivot ──
    # Deduplicate columns (multi-value pivot can produce duplicates)
    pivot = pivot.loc[:, ~pivot.columns.duplicated()]
    categories = list(pivot.index.astype(str))
    series: list[tuple[str, list[float]]] = []

    for ci in range(len(pivot.columns)):
        col_name = str(pivot.columns[ci])
        col_data = pivot.iloc[:, ci]
        vals = [0.0 if pd.isna(v) else float(v) for v in col_data]
        display_name = col_name.replace(" @:@ ", " - ").replace("@:@", " - ")
        series.append((display_name, vals))

    # ── Apply selectedRows (filter + explicit order) ──
    select_all = mapping_config.get("selectAllRows", True)
    selected_rows_raw = mapping_config.get("selectedRows", [])
    rows_ordered = False

    if not select_all and selected_rows_raw:
        # selectedRows may use @:@ separator for compound keys
        row_order = []
        for sr in selected_rows_raw:
            sr_norm = _normalize_key(sr)
            sr_parts = [p.strip() for p in sr_norm.split(" - ")]
            for ci, cat in enumerate(categories):
                # Match: exact, or last part matches, or cat is in any part
                if (cat == sr_norm or cat in sr_parts or
                        any(cat == p or p in cat or cat in p for p in sr_parts)):
                    if ci not in row_order:
                        row_order.append(ci)
                    break
        if row_order:
            categories = [categories[i] for i in row_order]
            series = [(name, [vals[i] for i in row_order if i < len(vals)])
                      for name, vals in series]
            rows_ordered = True

    # ── Apply moveRowsToFirst / moveRowsToLast ──
    move_first = mapping_config.get("moveRowsToFirst", [])
    move_last = mapping_config.get("moveRowsToLast", [])
    if move_first or move_last:
        fi, li, mi = [], [], []
        for ci, cat in enumerate(categories):
            if any(cat == m or m in cat for m in move_first):
                fi.append(ci)
            elif any(cat == m or m in cat for m in move_last):
                li.append(ci)
            else:
                mi.append(ci)
        new_order = fi + mi + li
        if new_order:
            categories = [categories[i] for i in new_order]
            series = [(n, [v[i] for i in new_order if i < len(v)])
                      for n, v in series]
            rows_ordered = True

    # ── Sort (only when explicit CustomList exists) ──
    # Don't apply default value-based sort — preserve data order.
    # Only sort when columnDefinitions has a CustomList (explicit category order).
    if not rows_ordered and series and series[0][1] and len(series[0][1]) == len(categories):
        col_defs = pivot_config.get("columnDefinitions", [])
        custom_list = None
        for cd in col_defs:
            sc = cd.get("sortCriteria", {})
            if sc:
                cl = sc.get("CustomList", [])
                if cl:
                    custom_list = cl
                break

        if custom_list:
            order = []
            for cl_item in custom_list:
                for ci, cat in enumerate(categories):
                    if cat == cl_item or cl_item in cat or cat in cl_item:
                        if ci not in order:
                            order.append(ci)
                        break
            for ci in range(len(categories)):
                if ci not in order:
                    order.append(ci)
            categories = [categories[i] for i in order]
            series = [(n, [v[i] for i in order]) for n, v in series]
            rows_ordered = True

    # ── Chronological wave sort ──
    # If no explicit ordering was applied and every category is a time_period
    # label, sort by its underlying time_period_id. Prevents the lex ordering
    # "Wave 1, Wave 10, Wave 2, Wave 3, ..." on deep wave counts.
    if not rows_ordered and name_to_tp_id and categories:
        if all(cat in name_to_tp_id for cat in categories):
            tp_order = sorted(
                range(len(categories)),
                key=lambda i: name_to_tp_id[categories[i]],
            )
            if tp_order != list(range(len(categories))):
                categories = [categories[i] for i in tp_order]
                series = [(n, [v[i] for i in tp_order if i < len(v)])
                          for n, v in series]

    # ── Source-chart-canonical alignment for non-wave dimensions ──
    # The source chart's category list defines what rows to show. After refresh,
    # if a category that was in source no longer has data, keep the row with a
    # placeholder (None → renders as blank/"-"). If new categories appear that
    # weren't in source, drop them. Same for series. Only apply on dimensions
    # that aren't waves — wave dims must respect dynamic_latest_n and grow/
    # shrink as configured.
    def _is_wave_dim(labels: list[str]) -> bool:
        if not labels:
            return False
        return all(_is_wave_label(str(l), known_wave_labels) for l in labels)

    if source_categories and not _is_wave_dim(source_categories):
        cat_to_idx = {c: i for i, c in enumerate(categories)}
        new_series = []
        for sname, vals in series:
            new_vals = []
            for src_cat in source_categories:
                if src_cat in cat_to_idx:
                    j = cat_to_idx[src_cat]
                    new_vals.append(vals[j] if j < len(vals) else None)
                else:
                    new_vals.append(None)
            new_series.append((sname, new_vals))
        categories = list(source_categories)
        series = new_series

    if source_series_names and not _is_wave_dim(source_series_names):
        current_names = [n for n, _ in series]
        name_to_idx = {n: i for i, n in enumerate(current_names)}
        new_series = []
        for src_name in source_series_names:
            if src_name in name_to_idx:
                new_series.append(series[name_to_idx[src_name]])
            else:
                new_series.append((src_name, [None] * len(categories)))
        series = new_series

    return ChartRefreshData(categories=categories, series=series)


def refresh_chart_from_synapse(
    chart_shape,
    report_config: dict,
    pivot_config: dict,
    mapping_config: dict,
    base_url: str,
    token: str,
) -> ChartRefreshData:
    """Full pipeline: fetch from Synapse → pivot → map → return chart data.

    Does NOT write to the chart — caller uses replace_data() with the result.
    """
    records = fetch_synapse_report(base_url, token, report_config)
    if not records:
        return ChartRefreshData(categories=[], series=[], success=False,
                                error="No records from Synapse API")

    static_names = report_config.get("StaticTimePeriodNames")

    result = pivot_records_to_chart_data(
        records, pivot_config, mapping_config, static_names
    )

    return result
