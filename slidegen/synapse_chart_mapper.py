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
    """Data ready to write into a chart via replace_data().

    `notes` carries non-fatal annotations the caller surfaces via slide
    badges and per-shape Connector tags. Each entry is a dict with keys:
        kind:    "tag_mismatch" | "partial_alignment" | ...
        detail:  short human-readable string for the badge / tag value
    """
    categories: list[str]
    series: list[tuple[str, list[float]]]  # [(series_name, values), ...]
    success: bool = True
    error: str = ""
    notes: list[dict] = field(default_factory=list)


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
    _re_wave.compile(r"^(?:Project )?Wave \d+$"),                     # Wave 12, Project Wave 12
    _re_wave.compile(r"^W\d+$"),                                      # W28 (short form)
    # MMM'YY — accept straight ' and curly '/' (PowerPoint smart quotes).
    _re_wave.compile(r"^[A-Z][a-z]{2}['’‘]\d{2}$"),                  # Jan'26 / Jan'26
    _re_wave.compile(r"^Q[1-4]['’‘]\d{2}$"),                         # Q1'26
    _re_wave.compile(r"^Q[1-4] \d{4}$"),                              # Q1 2026
    # Rolling-period labels: e.g. "Oct'25 - Dec'25" or "Jan'26 - Mar'26".
    _re_wave.compile(
        r"^[A-Z][a-z]{2}['’‘]\d{2}\s*[-–]\s*[A-Z][a-z]{2}['’‘]\d{2}$"
    ),
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


def _all_wave_like(labels, known_wave_labels: set[str] | None = None) -> bool:
    """True iff every label in the iterable is recognized as a wave label."""
    if not labels:
        return False
    return all(_is_wave_label(str(l), known_wave_labels) for l in labels)


_CHRONO_MONTH = {m: i for i, m in enumerate(
    ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"], start=1)}

# Wave-token regex for extracting a wave-shaped substring from compound labels
# like "Nov'25 - Gastro" so chronology sort works on segment-suffixed labels.
# Order matters: rolling MMM-MMM and dated tokens come before generic forms.
_WAVE_TOKEN_RE = _re_wave.compile(
    r"(?:^|\s|[-–\(])\s*("
    r"[A-Za-z]{3}['’‘]\d{2}\s*[-–]\s*[A-Za-z]{3}['’‘]\d{2}|"  # rolling
    r"[A-Za-z]{3}['’‘]\d{2}|"                                # MMM'YY
    r"Q[1-4]['’‘]\d{2}|"                                     # Q1'26
    r"Q[1-4] \d{4}|"                                          # Q1 2026
    r"(?:Project )?Wave \d+|"                                 # Wave 12
    r"W\d+"                                                   # W28
    r")\b"
)


def _extract_wave_token(label: str) -> str | None:
    """Find a wave-shaped token inside `label`, or return None.

    For compound labels like "Nov'25 - Gastro" returns "Nov'25" so chronology
    sort can operate on the wave part. Returns None if no recognizable wave
    token is present (so the caller can fall back to alphabetical).
    """
    if not isinstance(label, str) or not label:
        return None
    s = label.replace("’", "'").replace("‘", "'")
    m = _WAVE_TOKEN_RE.search(s)
    return m.group(1) if m else None


# ─────────────────────────────────────────────────────────────────────
# Formula columns (Connector: derived columns from columnDefinitions)
# ─────────────────────────────────────────────────────────────────────
# columnDefinitions can carry a Formula like "=B2/100" that derives a
# new column from existing pivot columns. Excel column letters map to
# columnDefinitions positions (the connector treats the row label as
# column A; the first value column is B, etc.). The row number portion
# changes per row — same operation applied to every row's value.
#
# Supported expression grammar (right-hand side of "="):
#     atom    := COL_REF | NUMBER
#     term    := atom ('*' atom | '/' atom)*
#     expr    := term ('+' term | '-' term)*
# COL_REF is one or more uppercase letters followed by digits.
# This covers >95% of real connector formulas (=B2, =B2/100, =B2*100,
# =B2+C2, =B2-C2, etc.). More exotic formulas (IF, ROUND, abs()) fall
# through to None and the column is dropped with a warning.
import re as _re_formula

_FORMULA_TOKEN_RE = _re_formula.compile(
    r"\s*(?:(?P<colref>[A-Z]+)(?P<row>\d+)|(?P<num>-?\d+(?:\.\d+)?)|(?P<op>[+\-*/]))"
)


def _col_letter_to_index(letters: str) -> int:
    """A->0, B->1, ..., Z->25, AA->26, AB->27, ..."""
    n = 0
    for c in letters:
        n = n * 26 + (ord(c) - ord("A") + 1)
    return n - 1


def _tokenize_formula(expr: str) -> list[tuple[str, object]] | None:
    """Tokenize the right-hand side of a formula. Returns list of
    (kind, value) pairs or None on parse error."""
    tokens: list[tuple[str, object]] = []
    pos = 0
    while pos < len(expr):
        m = _FORMULA_TOKEN_RE.match(expr, pos)
        if m is None or m.end() == pos:
            return None
        if m.group("colref"):
            tokens.append(("col", _col_letter_to_index(m.group("colref"))))
        elif m.group("num") is not None:
            tokens.append(("num", float(m.group("num"))))
        elif m.group("op"):
            tokens.append(("op", m.group("op")))
        pos = m.end()
    return tokens


def _eval_formula_for_row(tokens: list[tuple[str, object]],
                          row_values: list[float | None]) -> float | None:
    """Evaluate tokens for one row. Column refs resolve via row_values
    (indexed by Excel column letter, 0-based). Returns None on missing
    operand or division by zero."""
    # Two-pass shunting-yard simplified to */ first, then +-.
    # Step 1: resolve atoms to numeric values (or None).
    nums: list[float | None] = []
    ops: list[str] = []
    for kind, val in tokens:
        if kind == "col":
            if val < 0 or val >= len(row_values):
                return None
            nums.append(row_values[val])
        elif kind == "num":
            nums.append(float(val))
        elif kind == "op":
            ops.append(str(val))
    if not nums or len(ops) != len(nums) - 1:
        return None
    if any(n is None for n in nums):
        return None

    # Step 2: */ pass
    i = 0
    while i < len(ops):
        if ops[i] in ("*", "/"):
            a, b = nums[i], nums[i + 1]
            if ops[i] == "*":
                r = a * b
            else:
                if b == 0:
                    return None
                r = a / b
            nums[i] = r
            del nums[i + 1]
            del ops[i]
        else:
            i += 1
    # Step 3: +- pass
    result = nums[0]
    for i, op in enumerate(ops):
        b = nums[i + 1]
        result = result + b if op == "+" else result - b
    return result


def _build_formula_columns(
    col_defs: list[dict],
    series: list[tuple[str, list[float]]],
    categories: list[str],
) -> list[tuple[str, list[float]]]:
    """Return derived series for every columnDefinition that carries a
    Formula. Excel column letters map to columnDefinitions positions:
    columnDefinitions[0] is column A, [1] is column B, etc. The pivot's
    `series` provides values for the data columns (excluding the row
    label which Excel treats as column A).

    For each row, builds row_values[] indexed by Excel column letter:
      - row_values[col_letter_idx] = pivot value at that letter
    Then evaluates the formula tokens against row_values.
    """
    if not col_defs or not series:
        return []
    n_rows = len(categories)
    if n_rows == 0:
        return []

    # Map columnDefinitions position -> series index (or None if it's a
    # row-label column not present in `series`).
    # Series display names from the pivot look like "Project Wave 7 - mean"
    # or "Project Wave 7 @:@ Average of mean", while columnDefinitions[].Name
    # is the bare wave label "Project Wave 7". So we accept exact match,
    # alias match, OR a startswith-match where the cd Name appears as the
    # leading compound segment of the series name.
    def _norm_compound(s: str) -> str:
        return s.replace(" @:@ ", " - ")

    series_by_exact: dict[str, int] = {}
    series_by_prefix: dict[str, int] = {}
    for i, (sname, _) in enumerate(series):
        ns = _norm_compound(sname)
        series_by_exact[sname] = i
        series_by_exact[ns] = i
        # Leading compound segment, e.g. "Project Wave 7" from
        # "Project Wave 7 - mean". First occurrence wins (rare collisions).
        head = ns.split(" - ", 1)[0]
        series_by_prefix.setdefault(head, i)

    cd_to_series_idx: list[int | None] = []
    for cd in col_defs:
        name = cd.get("Name") or ""
        if name.startswith("<blank:"):
            cd_to_series_idx.append(None)  # placeholder for derived col
            continue
        norm_name = _norm_compound(name)
        if name in series_by_exact:
            cd_to_series_idx.append(series_by_exact[name])
            continue
        if norm_name in series_by_exact:
            cd_to_series_idx.append(series_by_exact[norm_name])
            continue
        if name in series_by_prefix:
            cd_to_series_idx.append(series_by_prefix[name])
            continue
        if norm_name in series_by_prefix:
            cd_to_series_idx.append(series_by_prefix[norm_name])
            continue
        # Last-resort suffix match: Connector default-alias rules drop
        # leading classifier parts of the column name. Examples:
        #   "Project Wave 7" -> "Wave 7"
        #   "Specialty (C/PCP Segments) + Tier Detailed - CARD - L"
        #   -> "CARD - L"
        # Match by progressively shorter trailing suffixes — split cd by
        # " - " and check if any series name (or its head) equals the
        # last-K parts joined back. Falls back to checking if the series
        # name appears as a tail substring of the cd name.
        sub_match = None
        cd_parts = [p.strip() for p in norm_name.split(" - ") if p.strip()]
        for i, (sname, _) in enumerate(series):
            ns = _norm_compound(sname)
            head = ns.split(" - ", 1)[0]
            # Try progressive suffixes of cd_parts (longest first)
            for k in range(len(cd_parts), 0, -1):
                suffix = " - ".join(cd_parts[-k:])
                if suffix == ns or suffix == head:
                    sub_match = i
                    break
            if sub_match is not None:
                break
            # Or: series name is a literal tail of the cd name
            if ns and norm_name.endswith(ns):
                sub_match = i
                break
            if head and norm_name.endswith(head):
                sub_match = i
                break
        if sub_match is not None:
            cd_to_series_idx.append(sub_match)
            continue
        # Try alias match
        alias = cd.get("Alias")
        if alias and alias in series_by_exact:
            cd_to_series_idx.append(series_by_exact[alias])
            continue
        cd_to_series_idx.append(None)  # row-label or unmatched

    out: list[tuple[str, list[float]]] = []
    for cd_idx, cd in enumerate(col_defs):
        formula = (cd.get("Formula") or "").strip()
        if not formula or not formula.startswith("="):
            continue
        tokens = _tokenize_formula(formula[1:])
        if tokens is None:
            continue
        # Build per-row values indexed by Excel column letter (0-based)
        row_results: list[float] = []
        for r in range(n_rows):
            row_values: list[float | None] = []
            for s_idx in cd_to_series_idx:
                if s_idx is None:
                    row_values.append(None)
                else:
                    vals = series[s_idx][1]
                    row_values.append(vals[r] if r < len(vals) else None)
            v = _eval_formula_for_row(tokens, row_values)
            row_results.append(0.0 if v is None else v)
        # Display name: prefer Alias, else strip "<blank:" wrapper
        display = cd.get("Alias") or cd.get("Name", "")
        if display.startswith("<blank:") and display.endswith(">"):
            display = display[len("<blank:"):-1]
        out.append((display, row_results))
    return out


def _wave_chrono_key(label: str) -> tuple:
    """Sortable key for a wave label so chronologically-later waves sort later.

    Mirrors `intelligent_refresh._chronological_wave_key` but local to the
    mapper. Handles MMM'YY (incl. curly quotes), MMM'YY-MMM'YY rolling,
    Q[1-4]'YY, Q[1-4] YYYY, Wave N, W N, Project Wave N. Unknown formats
    sort to the front.
    """
    s = (str(label) if label is not None else "").strip()
    if not s:
        return (-1, 0, s)
    s_norm = s.replace("’", "'").replace("‘", "'")
    m = _re_wave.match(r"^[A-Za-z]+'?\d{2}\s*[-–]\s*([A-Za-z]+)'?(\d{2})$", s_norm)
    if m:
        mon = m.group(1).lower()[:3]
        yy = int(m.group(2))
        return (2000 + yy if yy < 80 else 1900 + yy, _CHRONO_MONTH.get(mon, 0), 0)
    m = _re_wave.match(r"^([A-Za-z]+)'?(\d{2})$", s_norm)
    if m:
        mon = m.group(1).lower()[:3]
        yy = int(m.group(2))
        if mon in _CHRONO_MONTH:
            return (2000 + yy if yy < 80 else 1900 + yy, _CHRONO_MONTH[mon], 0)
    m = _re_wave.match(r"^Q([1-4])'?(\d{2})$", s_norm)
    if m:
        q = int(m.group(1)); yy = int(m.group(2))
        return (2000 + yy if yy < 80 else 1900 + yy, q * 3, 0)
    m = _re_wave.match(r"^Q([1-4])\s+(\d{4})$", s_norm)
    if m:
        return (int(m.group(2)), int(m.group(1)) * 3, 0)
    m = _re_wave.match(r"^(?:Project\s+)?Wave\s+(\d+)$", s_norm, _re_wave.IGNORECASE)
    if m:
        return (10000, int(m.group(1)), 0)
    m = _re_wave.match(r"^W(\d+)$", s_norm)
    if m:
        return (10000, int(m.group(1)), 0)
    # Layer 2: year-less quarter (Q1, Q2, ..., Q10).
    m = _re_wave.match(r"^Q(\d+)$", s_norm, _re_wave.IGNORECASE)
    if m:
        return (10000, int(m.group(1)), "q")
    # Layer 3: generic prefix-N (Period 1, M1, etc.). Trailing int is sort
    # key; leading prefix is tiebreaker for mixed-prefix charts. Already-dated
    # labels were caught above and won't reach here.
    m = _re_wave.match(r"^([A-Za-z][A-Za-z\s_\-]*?)\s*(\d+)$", s_norm)
    if m:
        prefix = m.group(1).strip().lower()
        n = int(m.group(2))
        return (10000, n, prefix)
    return (-1, 0, s)


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
    source_series_values: list[list] | None = None,
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

    # Survey-style data: ValueFields can name a measure-type label (e.g.
    # "mean") that isn't a column but appears in the `measure` column with
    # the actual numeric in `value`. Filter df to rows whose measure
    # matches val_field, then aggregate `value`. Without this filter the
    # pivot averages across Mean+Sum+Count+% rows together — produces
    # nonsense values (Repatha ATU slide 11 saw 2313 instead of 73 for
    # a Mean-typed chart).
    if (not primary_val
            and "measure" in df.columns
            and "value" in df.columns):
        for vf in val_fields:
            measure_mask = df["measure"].astype(str).str.lower() == vf.lower()
            if measure_mask.any():
                df = df[measure_mask].copy()
                primary_val = "value"
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
    # (or ~isin() for exclusion). For each split value we also try
    # compound-suffix matching ("Specialty - CARD - L" -> "CARD - L")
    # because source decks store the long form but the API often
    # returns just the short suffix.
    def _resolve_filter_values(values: list[str], col_series) -> list[str]:
        actual_values = set(col_series.astype(str).values)
        resolved: list[str] = []
        for v in values:
            v = v.strip()
            if not v:
                continue
            if v in actual_values:
                resolved.append(v)
                continue
            # Try compound-part suffix: walk from longest suffix down,
            # split on " - " or " @:@ ". First match wins.
            parts = []
            for sep in (" - ", " @:@ "):
                if sep in v:
                    parts = [p.strip() for p in v.split(sep) if p.strip()]
                    break
            if parts:
                # Try progressively shorter suffixes joined by " - ".
                matched = False
                for k in range(len(parts), 0, -1):
                    candidate = " - ".join(parts[-k:])
                    if candidate in actual_values:
                        resolved.append(candidate)
                        matched = True
                        break
                    # Case-insensitive check
                    lc = candidate.lower()
                    for av in actual_values:
                        if av.lower() == lc:
                            resolved.append(av)
                            matched = True
                            break
                    if matched:
                        break
                if matched:
                    continue
        return resolved

    for filt in pivot_config.get("Filters", []):
        col_key = _remap_field(filt.get("ColumnKey", ""))
        fval = filt.get("Value", "")
        criteria = filt.get("filterCriteria", 1)
        if col_key and fval and col_key in df.columns:
            if ";" in fval:
                raw_values = [v.strip() for v in fval.split(";") if v.strip()]
                values = _resolve_filter_values(raw_values, df[col_key])
                if not values:
                    # Resolution failed — leave df untouched rather than
                    # wiping it. Empty filter = all rows pass.
                    continue
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
                    # Compound filter values: source decks sometimes carry
                    # "Speciality - Gastro" while the API now returns just
                    # "Gastro". Try matching each compound part (split on
                    # " - " or " @:@ ") against actual values; prefer the
                    # most specific (last part), then fall back to first-
                    # word substring match.
                    parts = []
                    for sep in (" - ", " @:@ "):
                        if sep in fval:
                            parts = [p.strip() for p in fval.split(sep) if p.strip()]
                            break
                    matched = False
                    for part in reversed(parts):  # most specific first
                        if part in df[col_key].values:
                            df = df[df[col_key] == part]
                            matched = True
                            break
                        # Case-insensitive exact match
                        ci_mask = df[col_key].astype(str).str.lower() == part.lower()
                        if ci_mask.any():
                            df = df[ci_mask]
                            matched = True
                            break
                    if not matched:
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

    # ── Latest-wave-only filter ──
    # When `time_period_name` is in the data but NOT in ColumnFields (i.e.
    # waves aren't a chart axis) AND multiple waves are present, the user
    # intent is "show the latest snapshot," not "sum/mean across waves."
    # Without this filter, an L-style label_table with dynamic_latest_n=2
    # gets two waves' Mean values summed (148% instead of 74%).
    # Sort by chronological wave key (NOT time_period_id — that gotcha).
    if (
        "time_period_name" in df.columns
        and "time_period_name" not in (col_fields or [])
        and df["time_period_name"].nunique() > 1
    ):
        unique_waves = list(df["time_period_name"].astype(str).unique())
        # Use the existing chronology key. Pick the wave with the
        # highest key, falling back to lexicographic if all keys are
        # the unknown sentinel.
        scored = [(w, _wave_chrono_key(w)) for w in unique_waves]
        scored.sort(key=lambda kv: kv[1])
        latest_wave = scored[-1][0]
        df = df[df["time_period_name"].astype(str) == latest_wave]

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

    # ── Pre-aggregated metric switch ──
    # When the spec says ValueFields=['decimal'] but the data has a
    # pre-aggregated metric column (me_score, share, penetration, etc.)
    # AND the chart's pivot dimensions exclude the dim that decimal varies
    # on (typically x_label for ME charts), Connector internally switches
    # to the pre-aggregated column. Replicate that here.
    #
    # Trigger: x_label is in df with multiple distinct values, neither
    # RowFields nor ColumnFields includes x_label, and decimal varies
    # across x_label within the same (row × col) group while a pre-
    # aggregated column is constant within that group → use the pre-
    # aggregated column.
    _PRE_AGG_BY_VARIANCE_DIM = {
        # variance_dim → preferred pre-aggregated column when present
        "x_label": "me_score",
    }
    if (primary_val == "decimal" and not has_multi_vals
            and len(df) > 0):
        for var_dim, pre_agg_col in _PRE_AGG_BY_VARIANCE_DIM.items():
            if (var_dim in df.columns
                    and pre_agg_col in df.columns
                    and var_dim not in row_fields
                    and var_dim not in col_fields
                    and df[var_dim].nunique() > 1):
                # Verify pre-agg is constant within (row × col) groups while
                # decimal varies — that's the signature of a pre-aggregated
                # metric folded across the variance dim.
                group_keys = [k for k in (row_fields + col_fields) if k in df.columns]
                if group_keys:
                    g = df.groupby(group_keys, dropna=False)
                    if (g[pre_agg_col].nunique().max() == 1
                            and g["decimal"].nunique().max() > 1):
                        primary_val = pre_agg_col
                        resolved_val_fields = [(pre_agg_col, pre_agg_col)]
                        break

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

    # Honor PivotConfig.AggregationType for cells with multiple source rows
    # (e.g. ME chart's ColumnFields excludes x_label so multiple x_label
    # rows collapse into one cell per Code+wave+alias). AggregationType
    # convention in the Connector tag:
    #   0  Sum
    #   1  Average  (Connector default)
    #   2  Count
    #   3  Min
    #   4  Max
    # Source-deck inferred specs sometimes encode "first" semantics — fall
    # back to mean if AggregationType=1, sum if 0, etc. The previous
    # `aggfunc="first"` silently picked the alphabetically-first row of each
    # group, which produced wrong values for ME charts (slide 25/26 issue).
    _agg_map = {0: "sum", 1: "mean", 2: "count", 3: "min", 4: "max"}
    _agg_type = pivot_config.get("AggregationType", 1)
    _aggfunc = _agg_map.get(_agg_type, "mean")

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
                    aggfunc=_aggfunc,
                )
                pivot_frames.append(sub_pivot)
            pivot = pd.concat(pivot_frames, axis=1)
        else:
            pivot = df.pivot_table(
                index=pivot_index,
                columns="_col_key",
                values=primary_val,
                aggfunc=_aggfunc,
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

    _selected_columns_unmatched: list[str] | None = None
    _selected_columns_zero_match: bool = False
    if series_columns:
        # Keep only columns that match selectedColumns entries.
        cols_to_keep = []
        unmatched = []
        for sc in series_columns:
            sc_norm = _normalize_key(sc)
            matched = _match_pivot_col(sc_norm, pivot.columns)
            if matched is not None:
                cols_to_keep.append(matched)
            else:
                unmatched.append(sc)
        if cols_to_keep:
            pivot = pivot[cols_to_keep]
            # When SOME selectedColumns matched but others didn't, the source's
            # filter intent partially survives. Surface the unmatched ones.
            if unmatched:
                _selected_columns_unmatched = unmatched
        elif series_columns:
            # ZERO matches → segment dim or column structure changed at the
            # API since the deck was rendered. Don't silently leak all pivot
            # columns into the chart; record so the caller can surface a
            # tag-mismatch/partial_alignment note.
            _selected_columns_unmatched = list(series_columns)
            _selected_columns_zero_match = True

    # ── Pre-split row reorder per columnDefinitions[row_field].sortCriteria.CustomList ──
    # Connector applies the CustomList from the row field's columnDefinition
    # to the pivot rows BEFORE slicing for split_order. Without this, pivot
    # rows arrive in groupby/alphabetical order and split_order=0 picks the
    # alphabetically-first y_label instead of the user-intended first row
    # (e.g. "Improves symptoms" listed first in the source table). The
    # downstream split_order=N then misaligns chart shapes vs the table.
    if rows_per_object and split_order is not None and len(valid_row_fields) >= 1:
        primary_rf = display_row_field if not use_compound else valid_row_fields[0]
        col_def = next(
            (cd for cd in pivot_config.get("columnDefinitions", [])
             if cd.get("Name") == primary_rf),
            None,
        )
        if col_def:
            custom_list = (col_def.get("sortCriteria") or {}).get("CustomList") or []
            if custom_list:
                order_map = {str(it): i for i, it in enumerate(custom_list)}
                existing_idx = list(pivot.index)
                sorted_idx = sorted(
                    existing_idx,
                    key=lambda x: order_map.get(str(x), len(custom_list) + 1),
                )
                if sorted_idx != existing_idx:
                    pivot = pivot.reindex(sorted_idx)

    # ── Step 1b: Apply topNRows + rowsPerObject/splitOrder (split visualization) ──
    # Connector's SplitVisualizationService splits pivot rows across chart shapes.
    # topNRows: limit total rows. rowsPerObject=1 + splitOrder=K: take row K only.
    if top_n_rows and top_n_rows > 0 and len(pivot) > top_n_rows:
        pivot = pivot.iloc[:top_n_rows]

    if rows_per_object and rows_per_object > 0 and split_order is not None:
        # Source-driven row selection: when the source chart shape's series
        # name (passed as the only entry in source_series_names for split-
        # viz) maps to a row in the pivot, pick THAT row instead of indexing
        # by split_order. Preserves the user contract: "each chart shape
        # keeps the message it had pre-refresh" — no surprise reassignment.
        #
        # Connector charts often display y_code (e.g. "CR6") as the series
        # label while pivoting on y_label (the full message text). When a
        # direct name match fails, try translating via the data's
        # code↔label columns.
        used_source_match = False
        if (rows_per_object == 1
                and source_series_names
                and len(source_series_names) == 1):
            target = str(source_series_names[0]).strip()
            target_low = target.lower()
            # Pass 1: direct match on pivot.index
            for idx_val in pivot.index:
                if (str(idx_val).strip() == target
                        or str(idx_val).strip().lower() == target_low):
                    pivot = pivot.loc[[idx_val]]
                    used_source_match = True
                    break
            # Pass 2: code↔label translation. If the source's series name
            # looks like a y_code value but pivot indexes on y_label (or vice
            # versa), translate via the source df.
            if not used_source_match:
                code_label_pairs = [
                    ("y_code", "y_label"), ("y_label", "y_code"),
                    ("Codes", "y_label"), ("y_label", "Codes"),
                    ("y_code", "Codes"), ("Codes", "y_code"),
                ]
                for from_col, to_col in code_label_pairs:
                    if from_col not in df.columns or to_col not in df.columns:
                        continue
                    matches = df[df[from_col].astype(str).str.strip() == target]
                    if matches.empty:
                        matches = df[df[from_col].astype(str).str.strip().str.lower()
                                     == target_low]
                    if not matches.empty:
                        translated = str(matches[to_col].iloc[0]).strip()
                        translated_low = translated.lower()
                        for idx_val in pivot.index:
                            if (str(idx_val).strip() == translated
                                    or str(idx_val).strip().lower() == translated_low):
                                pivot = pivot.loc[[idx_val]]
                                used_source_match = True
                                break
                    if used_source_match:
                        break
            # When source-driven match succeeded, also relabel the pivot row
            # to the source's series name (e.g. "CR6") so the post-transpose
            # series.name matches what the source chart had — keeps table-row
            # and chart-shape attributes aligned per user contract.
            if used_source_match and source_series_names:
                pivot.index = [source_series_names[0]] * len(pivot.index)
        if not used_source_match:
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

    # ── Step 3b: Apply formula columns (Connector: derived columns) ──
    # columnDefinitions can declare a derived column with an Excel-style
    # Formula. Example from Repatha ATU slide 11:
    #     {"Name": "<blank:AvgPercent>", "Formula": "=B2/100", "Format": "0%"}
    # The chart's selectedColumns references "<blank:AvgPercent>" but the
    # value in each row is computed via the formula on adjacent columns
    # (column letters map to columnDefinitions positions; row number is
    # the per-row substitution). Without applying the formula, the API's
    # raw "mean" (e.g. 73) leaks into the chart, then the chart's "0%"
    # formatCode displays it as "7300%".
    #
    # Supported patterns: =BN, =BN/X, =BN*X, =BN+CN, =BN-CN, =BN+X, =BN-X
    # where B,C are column letters (single-letter only) and X is numeric.
    col_defs = pivot_config.get("columnDefinitions", []) or []
    if col_defs:
        formula_cols = _build_formula_columns(col_defs, series, categories)
        if formula_cols:
            series.extend(formula_cols)

    # ── Apply selectedRows (filter + explicit order) ──
    select_all = mapping_config.get("selectAllRows", True)
    selected_rows_raw = mapping_config.get("selectedRows", [])
    rows_ordered = False
    _row_drop_note: dict | None = None

    if not select_all and selected_rows_raw:
        # selectedRows may use @:@ separator for compound keys.
        # Match case-insensitively + whitespace-normalized so source code
        # 'C10' aligns with pivot value 'C10 ' or 'c10' without dropping rows.
        def _norm_row(s):
            return str(s).strip().lower() if s is not None else ""

        cat_norm = [_norm_row(c) for c in categories]
        row_order = []
        sr_dropped: list[str] = []
        for sr in selected_rows_raw:
            sr_str = str(sr)
            sr_norm = _norm_row(_normalize_key(sr_str))
            sr_parts = [_norm_row(p) for p in sr_norm.split(" - ")]
            matched_ci = None
            # Pass 1: exact match (cat == sr_norm). Prevents 'C1' from
            # incorrectly substring-matching 'C10', 'C11', 'C17' which
            # silently dropped 4 codes on slide 25.
            for ci, cn in enumerate(cat_norm):
                if cn == sr_norm and ci not in row_order:
                    matched_ci = ci
                    break
            # Pass 2: cat is one of the compound parts (entire-part match).
            # E.g. cat='c10' matches sr_parts=['x', 'c10']; but does NOT
            # match sr_parts=['c1'] because 'c10' != 'c1'.
            if matched_ci is None:
                for ci, cn in enumerate(cat_norm):
                    if cn in sr_parts and ci not in row_order:
                        matched_ci = ci
                        break
            if matched_ci is not None:
                row_order.append(matched_ci)
            else:
                sr_dropped.append(sr_str)
        if row_order:
            categories = [categories[i] for i in row_order]
            series = [(name, [vals[i] for i in row_order if i < len(vals)])
                      for name, vals in series]
            rows_ordered = True
        if sr_dropped:
            # Stash on a local list — combined with cat/series notes at the
            # end of the function so the slide badge + REFRESH_NOTE tag sees
            # both row-drop and partial_alignment together.
            sample = sr_dropped[:5]
            more = len(sr_dropped) - len(sample)
            suffix = "..." if more > 0 else ""
            _row_drop_note = {
                "kind": "rows_dropped",
                "detail": (
                    f"{len(sr_dropped)} of {len(selected_rows_raw)} "
                    f"selectedRows had no match in API output: "
                    f"{', '.join(sample)}{suffix}"
                ),
            }
        else:
            _row_drop_note = None

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
        # Only honor the move when at least one cat actually matched a
        # move_first/move_last directive. If fi and li are both empty, the
        # directive doesn't apply to the current cat set (e.g. "Feb'26" was
        # listed in moveRowsToLast but Feb is no longer in the wave window).
        # Without this guard the no-op reorder set rows_ordered=True and
        # suppressed the chronological wave sort downstream.
        if new_order and (fi or li):
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
            # Only honor CustomList when at least one item matched. Post-
            # transpose, the CustomList may target the original axis (e.g.
            # y_label) while categories are now the OTHER axis (waves) —
            # no matches means the CustomList doesn't apply to this view.
            # Without this guard the fallback "append unmatched" branch
            # produced a no-op reorder while still setting rows_ordered=True,
            # which suppressed the chronological wave sort downstream.
            if order:
                for ci in range(len(categories)):
                    if ci not in order:
                        order.append(ci)
                categories = [categories[i] for i in order]
                series = [(n, [v[i] for i in order]) for n, v in series]
                rows_ordered = True

    # ── Chronological wave sort with direction preservation ──
    # Universal rule: whichever slot held the LATEST source wave still holds
    # the latest post-refresh. Practical effect: if source had Feb on LHS and
    # Mar on RHS (latest on RHS → ascending), refresh produces [Mar, Apr]. If
    # source had Mar on LHS and Feb on RHS (latest on LHS → descending),
    # refresh produces [Apr, Mar]. time_period_id is NOT chronological
    # (CLAUDE.md gotcha) — parse names instead.
    #
    # This block runs even when rows_ordered=True (i.e. CustomList or moveRows
    # already touched the order) because for wave dimensions the chronological
    # order with direction inferred from source is canonical. CustomLists in
    # source decks frequently encode stale "Feb, Mar" pairs that do not match
    # the actual rendered direction; trusting them blindly leaves the chart
    # with refreshed values in the wrong slots.
    # Compute chrono keys per category. Three modes:
    #   1) pure wave label ("Nov'25") → _wave_chrono_key returns a real key
    #   2) compound with wave token ("Nov'25 - Gastro") → extract token first
    #   3) neither → key starts with -1 (unsortable)
    # If every category has a chrono key under (1) or (2), the chart's
    # category axis is temporal and we run the chronological sort. The
    # compound case handles segment-suffixed cats from segment-applied charts
    # (slides 68-79) which previously fell through alphabetical.
    def _cat_chrono_key(c: str) -> tuple:
        k = _wave_chrono_key(c)
        if k[0] != -1:
            return k
        token = _extract_wave_token(c)
        if token:
            sub = _wave_chrono_key(token)
            if sub[0] != -1:
                return sub
        return k

    cat_chrono_keys = [_cat_chrono_key(c) for c in categories] if categories else []
    all_have_chrono = bool(cat_chrono_keys) and all(k[0] != -1 for k in cat_chrono_keys)

    if all_have_chrono:
        descending = False
        if source_categories and len(source_categories) >= 2:
            src_keys = [_cat_chrono_key(c) for c in source_categories]
            if all(k[0] != -1 for k in src_keys):
                # Position of the chronologically-latest source wave
                latest_pos = max(range(len(src_keys)), key=lambda i: src_keys[i])
                if latest_pos == 0:
                    descending = True
        tp_order = sorted(
            range(len(categories)),
            key=lambda i: cat_chrono_keys[i],
            reverse=descending,
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

    # Match labels case-insensitively + whitespace-normalized so source
    # 'NO' / 'YES' aligns with mapper 'No' / 'Yes' from y_label values.
    # Without this normalization, a casing mismatch in the source deck
    # (uppercase title-cased vs API mixed-case y_label) nulls out the
    # entire chart even though the data is present.
    def _norm_label(s) -> str:
        s = str(s) if s is not None else ""
        return _re_wave.sub(r"\s+", " ", s).strip().lower()

    # Snapshot the API-fetched data BEFORE source-canonical mutates it.
    # Used for (a) tag_mismatch detection — when source and API have zero
    # overlap on both axes, write API data faithfully per user contract
    # rather than preserve source (Bucket A: review tag or move to non-
    # connected); and (b) end-of-block notes (dynamic_added when the tag is
    # dynamic and new items flowed in; partial_alignment when the tag is
    # static and API extras exist that the user may want to opt into).
    api_cats_pre = list(categories)
    api_series_pre = [(n, list(v)) for n, v in series]

    # Dynamic vs static: if the tag pins specific time periods, treat the
    # whole chart as static — source-canonical alignment locks cat/series
    # structure to the slide. Otherwise the tag is dynamic and new items
    # from the API are allowed to flow into the slide on refresh.
    is_dynamic = not bool(static_time_period_ids)

    # Tag-mismatch detection: both axes have structure on both sides AND
    # no labels overlap → tag points at a different analysis than what was
    # rendered. Skip source-canonical, return API data with a note.
    # Fires regardless of dynamic/static — the diagnosis is the same.
    if (api_cats_pre and api_series_pre
            and (source_categories and not _is_wave_dim(source_categories))
            and (source_series_names and not _is_wave_dim(source_series_names))):
        api_cat_set = {_norm_label(c) for c in api_cats_pre}
        src_cat_set = {_norm_label(c) for c in source_categories}
        api_ser_set = {_norm_label(n) for n, _ in api_series_pre}
        src_ser_set = {_norm_label(n) for n in source_series_names}
        if (api_cat_set and src_cat_set and api_ser_set and src_ser_set
                and not (api_cat_set & src_cat_set)
                and not (api_ser_set & src_ser_set)):
            return ChartRefreshData(
                categories=api_cats_pre,
                series=api_series_pre,
                success=True,
                notes=[{
                    "kind": "tag_mismatch",
                    "detail": (
                        f"Tag fetched {len(api_cats_pre)} cats / "
                        f"{len(api_series_pre)} series with no overlap to source "
                        f"({len(source_categories)} / {len(source_series_names)}). "
                        "Review tag or move to non-connected path."
                    ),
                }],
            )

    if (source_categories and not _is_wave_dim(source_categories)
            and not is_dynamic):
        norm_to_idx = {_norm_label(c): i for i, c in enumerate(categories)}

        def _resolve_cat_idx(src_cat: str) -> int | None:
            """Find the pivot category that corresponds to a source category.

            Resolution tiers (each only fires when previous misses):
              1. Exact normalized match.
              2. Progressive-suffix on " - " split (handles default-alias
                 renames "Specialty - X - CARD" -> "CARD").
              3. Tail substring containment in either direction.
              4. Fuzzy similarity (SequenceMatcher.ratio >= 0.85) as a
                 final guard against minor rewording / whitespace drift.
                 Only applies when both labels have >= 4 chars to avoid
                 spurious matches between short tokens.
            """
            key = _norm_label(src_cat)
            if key in norm_to_idx:
                return norm_to_idx[key]
            parts = [p.strip() for p in key.split(" - ") if p.strip()]
            for k in range(len(parts), 0, -1):
                suffix = " - ".join(parts[-k:])
                if suffix in norm_to_idx:
                    return norm_to_idx[suffix]
            for cat_norm, idx in norm_to_idx.items():
                if cat_norm and (key.endswith(cat_norm) or cat_norm.endswith(key)):
                    return idx
            # Fuzzy fallback — pick the BEST match above the threshold,
            # not the first one. Skip very short labels.
            if len(key) >= 4:
                from difflib import SequenceMatcher
                best_idx, best_score = None, 0.0
                for cat_norm, idx in norm_to_idx.items():
                    if not cat_norm or len(cat_norm) < 4:
                        continue
                    score = SequenceMatcher(None, key, cat_norm).ratio()
                    if score > best_score:
                        best_score, best_idx = score, idx
                if best_score >= 0.85:
                    return best_idx
            return None

        new_series = []
        for sname, vals in series:
            new_vals = []
            for src_cat in source_categories:
                j = _resolve_cat_idx(src_cat)
                if j is not None:
                    new_vals.append(vals[j] if j < len(vals) else None)
                else:
                    new_vals.append(None)
            new_series.append((sname, new_vals))
        categories = list(source_categories)
        series = new_series

    # Always reorder series to source order when source_series_names map
    # 1:1 to pivot output — even for dynamic charts. Pivot's column
    # extraction is alphabetical by default; if the source had Wave 7
    # before Wave 6 (visual most-recent-first) and the pivot returns
    # them alphabetically, the chart's bars switch sides post-refresh
    # without any data justification (Repatha ATU slide 49 symptom).
    # This is purely a reorder — no value substitution, no preservation.
    if (source_series_names and not _is_wave_dim(source_series_names)
            and is_dynamic and len(source_series_names) == len(series)):
        current_names = [n for n, _ in series]
        norm_to_idx_dyn: dict[str, int] = {}
        for i, n in enumerate(current_names):
            norm_to_idx_dyn.setdefault(_norm_label(n), i)
        reorder_idx: list[int] = []
        used: set[int] = set()
        all_resolved = True
        for src_name in source_series_names:
            ni = norm_to_idx_dyn.get(_norm_label(src_name))
            # Try the same fuzzy compound-prefix logic the static path uses
            if ni is None:
                for i, cn in enumerate(current_names):
                    cn_norm = _norm_label(cn)
                    if (i not in used
                            and (_norm_label(src_name).endswith(cn_norm)
                                 or cn_norm.endswith(_norm_label(src_name)))):
                        ni = i
                        break
            if ni is None or ni in used:
                all_resolved = False
                break
            reorder_idx.append(ni)
            used.add(ni)
        if all_resolved and reorder_idx != list(range(len(series))):
            series = [series[i] for i in reorder_idx]

    if (source_series_names and not _is_wave_dim(source_series_names)
            and not is_dynamic):
        current_names = [n for n, _ in series]
        norm_to_idx = {_norm_label(n): i for i, n in enumerate(current_names)}

        def _split_compound(label: str) -> list[str]:
            """Normalize then split a compound label into parts.
            Tries " @:@ " first (Connector default), then " - " (Pradeep
            inferred). Returns the parts list."""
            s = _norm_label(label)
            for sep in (" @:@ ", " - "):
                if sep in s:
                    return [p.strip() for p in s.split(sep) if p.strip()]
            return [s]

        # Build prefix index: a tuple of leading-N parts → index for fuzzy
        # match when exact normalized fails (e.g., source has hardcoded
        # display alias on the last part that differs from mapper's raw
        # measure name: 'Feb'26 - ZENPEP ME - Average of Message
        # Effectiveness(%)' vs mapper 'Feb'26 - ZENPEP ME - Average of
        # me_score' should match on the first 2 parts).
        prefix_idx: dict[tuple, list[int]] = {}
        for i, n in enumerate(current_names):
            parts = _split_compound(n)
            for plen in range(1, len(parts)):
                prefix = tuple(parts[:plen])
                prefix_idx.setdefault(prefix, []).append(i)

        def _fuzzy_lookup(src_name: str) -> int | None:
            parts = _split_compound(src_name)
            for plen in range(len(parts) - 1, 0, -1):
                prefix = tuple(parts[:plen])
                hits = prefix_idx.get(prefix, [])
                if len(hits) == 1:
                    return hits[0]
            # Final fallback: text similarity. Pick the BEST current_name
            # whose normalized form has a SequenceMatcher.ratio >= 0.85
            # against src_name's normalized form. Guards against minor
            # rewording or punctuation drift while staying conservative
            # enough that distinct labels won't collide.
            sn = _norm_label(src_name)
            if len(sn) < 4:
                return None
            from difflib import SequenceMatcher
            best_idx, best_score = None, 0.0
            for i, n in enumerate(current_names):
                cn = _norm_label(n)
                if len(cn) < 4:
                    continue
                score = SequenceMatcher(None, sn, cn).ratio()
                if score > best_score:
                    best_score, best_idx = score, i
            return best_idx if best_score >= 0.85 else None

        # Resolve each source name → index into mapper series.
        matched: list[int | None] = []
        for src_name in source_series_names:
            key = _norm_label(src_name)
            if key in norm_to_idx:
                matched.append(norm_to_idx[key])
            else:
                matched.append(_fuzzy_lookup(src_name))

        # Positional fallback: when the source's series count exactly matches
        # the mapper's series count, and one or more source names couldn't be
        # resolved by name, line them up positionally against the leftover
        # mapper series. Handles two common source-deck quirks:
        #   (a) source series with empty label '' (slide 26 'Chart 29' family)
        #   (b) source label that's a stale literal (e.g. 'Project Wave - 8949'
        #       on slide 8 'Chart 27', a Connector pre-render artifact)
        if (len(source_series_names) == len(current_names)
                and any(m is None for m in matched)):
            used = {m for m in matched if m is not None}
            leftover = [i for i in range(len(current_names)) if i not in used]
            leftover_iter = iter(leftover)
            for k, m in enumerate(matched):
                if m is None:
                    try:
                        matched[k] = next(leftover_iter)
                    except StopIteration:
                        pass

        # When alignment fails for a series, prefer source values over Nones
        # so the chart visibly retains its previous state instead of
        # rendering empty. Per user contract: "if no usable new data, slide
        # stays untouched." Falling back to source values is only safe when
        # source values are aligned to the same category order we now have —
        # i.e., when source_categories was provided and accepted.
        can_preserve = bool(
            source_series_values
            and source_categories
            and len(categories) == len(source_categories)
        )

        new_series = []
        n_aligned = 0
        n_preserved = 0
        for k, (src_name, m) in enumerate(zip(source_series_names, matched)):
            if m is not None and m < len(series):
                _, mapper_vals = series[m]
                new_series.append((src_name, mapper_vals))
                n_aligned += 1
                continue
            if can_preserve and k < len(source_series_values):
                preserved = list(source_series_values[k])
                if len(preserved) < len(categories):
                    preserved += [None] * (len(categories) - len(preserved))
                elif len(preserved) > len(categories):
                    preserved = preserved[: len(categories)]
                new_series.append((src_name, preserved))
                n_preserved += 1
            else:
                new_series.append((src_name, [None] * len(categories)))
        series = new_series

        # If no series aligned at all, signal alignment failure so the caller
        # can preserve the entire source chart (status=alignment_failed) rather
        # than write a chart of all-preserved values that's identical to source.
        if n_aligned == 0 and n_preserved > 0:
            return ChartRefreshData(
                categories=categories, series=series, success=False,
                error="alignment_failed: no API series aligned with source",
            )

    # Chart-level safety net: if every series ended up all-None across all
    # categories (e.g. category alignment dropped every value because the API
    # cats and source cats are disjoint), don't write that empty data into
    # the chart — signal alignment failure so the caller preserves source.
    # This catches the second class of all-None corruption that the series
    # preserve path doesn't cover (per-category misses inside aligned series).
    if series and all(
        all(v is None for v in vals) for _name, vals in series
    ):
        return ChartRefreshData(
            categories=categories, series=series, success=False,
            error="alignment_failed: refreshed values are entirely None",
        )

    # ── Annotation notes for downstream slide badges + Connector tags ──
    # When API has more items than source on a non-wave dim, two cases:
    #   - Dynamic tag → items already flowed in (we skipped source-canonical).
    #     Emit `dynamic_added` so the slide visibly reflects the additions.
    #   - Static tag → items were dropped (source-canonical kept slide order).
    #     Emit `partial_alignment` so the user can opt in if desired.
    note_kind = "dynamic_added" if is_dynamic else "partial_alignment"
    notes: list[dict] = []

    def _format_note(api_only: list[str], dim: str) -> dict:
        sample = api_only[:3]
        more = len(api_only) - len(sample)
        suffix = "..." if more > 0 else ""
        noun = (
            "category" if (dim == "cats" and len(api_only) == 1)
            else "categories" if dim == "cats"
            else "series"
        )
        verb = "flowed in per dynamic tag" if is_dynamic else "not on slide"
        return {
            "kind": note_kind,
            "detail": (
                f"API has {len(api_only)} extra {noun} {verb}: "
                f"{', '.join(sample)}{suffix}"
            ),
        }

    if (source_categories and not _is_wave_dim(source_categories)
            and api_cats_pre):
        api_only = sorted(
            {_norm_label(c) for c in api_cats_pre}
            - {_norm_label(c) for c in source_categories}
        )
        if api_only:
            notes.append(_format_note(api_only, "cats"))
    if (source_series_names and not _is_wave_dim(source_series_names)
            and api_series_pre):
        api_only = sorted(
            {_norm_label(n) for n, _ in api_series_pre}
            - {_norm_label(n) for n in source_series_names}
        )
        if api_only:
            notes.append(_format_note(api_only, "series"))

    # Row-drop diagnostic from selectedRows match earlier in the function
    if _row_drop_note is not None:
        notes.append(_row_drop_note)

    # selectedColumns drift diagnostic: source listed columns that don't
    # match current pivot output (segment dim changed, column-structure
    # differs at the API since deck render, etc.).
    if _selected_columns_unmatched:
        sample = _selected_columns_unmatched[:3]
        more = len(_selected_columns_unmatched) - len(sample)
        suffix = "..." if more > 0 else ""
        notes.append({
            "kind": "selectedColumns_drift",
            "detail": (
                f"{len(_selected_columns_unmatched)} selectedColumns "
                f"could not be matched against current pivot output "
                f"(segment dim or column structure changed): "
                f"{', '.join(sample)}{suffix}. Review tag or move to non-connected."
            ),
        })

    # When the source's column-axis filter matched ZERO API columns, the
    # chart's pre-refresh structure can't be reconstructed (API analysis is
    # structurally different from what the deck was rendered against). The
    # current pivot would dump all API columns — including ones the source
    # was filtering out — producing a chart with the wrong attributes/series
    # (slide 80: each shape was filtered to 1 attribute × 2 brands but API
    # returns 2 attributes × 2 segment values, so all 4 leak into every
    # shape). Signal alignment_failed so the orchestrator preserves source.
    if (_selected_columns_zero_match
            and source_series_values
            and source_series_names
            and source_categories):
        return ChartRefreshData(
            categories=list(source_categories),
            series=[(n, list(v)) for n, v in zip(
                source_series_names, source_series_values)],
            success=False,
            error=(
                "alignment_failed: selectedColumns fully unmatched "
                "(API column structure differs from source)"
            ),
            notes=notes,
        )

    return ChartRefreshData(categories=categories, series=series, notes=notes)


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
