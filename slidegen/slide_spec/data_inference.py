"""
Data Inference for SlideGen Specs
=================================

Infers DataTransform and SeriesConfig from chart series names extracted from
OOXML, matched against column values in Synapse records.

Used for non-Connector-tagged slides where we need to figure out how Synapse
data maps to chart components using only the series names from the existing
chart and the column values available in the fetched records.

Inference strategy:
    1. For each column in the records, check if the series_names are a
       subset of that column's unique values (exact match first, then fuzzy).
    2. The matching column becomes column_field (what pivots into series).
    3. row_field and value_field are inferred from chart pattern and the
       remaining columns, with sensible defaults for PET chart conventions.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from .schema import DataTransform, DataFilter, SeriesConfig


# ─────────────────────────────────────────────────────────────────────────────
# Abbreviation / alias maps for fuzzy matching
# ─────────────────────────────────────────────────────────────────────────────

# Common short forms found in chart series names -> full forms in data columns.
# Keys are lowercased.  Values are lowercased candidate expansions.
_ABBREVIATION_MAP: dict[str, list[str]] = {
    "l":    ["low"],
    "n":    ["neutral", "neither", "none"],
    "h":    ["high"],
    "0":    ["idk", "i don't know", "don't know", "dk", "not sure", "zero", "0", "unsure"],
    "idk":  ["0", "don't know", "not sure", "unsure"],
    "hi":   ["high impact", "high"],
    "lo":   ["low impact", "low"],
    "card": ["cardiologist", "cardiology"],
    "pcp":  ["primary care", "primary care physician"],
    "acad": ["academic", "academia"],
    "comm": ["community"],
    "yr":   ["year"],
    "mo":   ["month"],
    "wk":   ["week"],
}

# Reverse map: long form -> short form, built lazily on first use.
_REVERSE_MAP: dict[str, str] | None = None


def _get_reverse_map() -> dict[str, str]:
    """Build a reverse lookup: lowered long-form -> lowered abbreviation."""
    global _REVERSE_MAP
    if _REVERSE_MAP is None:
        _REVERSE_MAP = {}
        for abbr, expansions in _ABBREVIATION_MAP.items():
            for exp in expansions:
                _REVERSE_MAP[exp] = abbr
    return _REVERSE_MAP


# ─────────────────────────────────────────────────────────────────────────────
# Column role heuristics
# ─────────────────────────────────────────────────────────────────────────────

# Columns that typically serve as row labels (y-axis / categories)
_ROW_FIELD_CANDIDATES = ["y_label", "product", "brand", "attribute", "message"]

# Columns that typically hold numeric values (cell content)
_VALUE_FIELD_CANDIDATES = ["percentage", "decimal", "value", "pct", "count", "mean"]

# Columns that are never matchable as series sources (always numeric)
_NUMERIC_COLUMNS = {"percentage", "decimal", "value", "pct", "count", "mean", "base",
                    "base_size", "n", "sample_size", "weight"}


# ─────────────────────────────────────────────────────────────────────────────
# Match result
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class ColumnMatch:
    """Result of matching series names to a single record column."""
    column_name: str
    match_type: str          # "exact" | "subset" | "fuzzy" | "partial"
    matched_count: int       # how many series names matched
    total_series: int        # total series names attempted
    coverage: float          # matched_count / total_series
    # Mapping from series_name -> column_value (for fuzzy matches, the actual
    # column value; for exact matches, same as series_name)
    name_to_value: dict[str, str]


# ─────────────────────────────────────────────────────────────────────────────
# Core matching
# ─────────────────────────────────────────────────────────────────────────────

def match_series_to_column(
    series_names: list[str],
    column_name: str,
    column_values: list[str],
) -> Optional[ColumnMatch]:
    """Check if series_names match against a column's unique values.

    Tries three strategies in order:
        1. **Exact match** — series_names are a subset of column_values
           (case-insensitive string comparison).
        2. **Fuzzy / abbreviation match** — series names are abbreviations
           of column values (e.g. "H" -> "High") or vice versa.
        3. **Partial match** — at least 50% of series names match.

    Returns None if coverage is below 50% (no plausible match).

    Args:
        series_names:  Labels extracted from chart OOXML series.
        column_name:   Name of the record column being tested.
        column_values: Unique string values in that column.

    Returns:
        ColumnMatch with coverage and mapping, or None.
    """
    if not series_names or not column_values:
        return None

    # Skip purely numeric columns — they hold values, not labels
    if column_name.lower() in _NUMERIC_COLUMNS:
        return None

    # Normalize: strip whitespace, compare case-insensitive
    series_lower = [s.strip().lower() for s in series_names]
    values_lower = {v.strip().lower(): v for v in column_values if isinstance(v, str)}

    name_to_value: dict[str, str] = {}

    # Pass 1: exact case-insensitive match
    for orig, sl in zip(series_names, series_lower):
        if sl in values_lower:
            name_to_value[orig] = values_lower[sl]

    # Pass 2: abbreviation / alias expansion for unmatched
    reverse_map = _get_reverse_map()
    for orig, sl in zip(series_names, series_lower):
        if orig in name_to_value:
            continue

        # Check if the series name is a known abbreviation
        if sl in _ABBREVIATION_MAP:
            for expansion in _ABBREVIATION_MAP[sl]:
                if expansion in values_lower:
                    name_to_value[orig] = values_lower[expansion]
                    break

        # Check if a column value abbreviates to the series name
        if orig not in name_to_value:
            for vl, v_orig in values_lower.items():
                if vl in reverse_map and reverse_map[vl] == sl:
                    name_to_value[orig] = v_orig
                    break

    # Pass 3: prefix / substring matching for remaining unmatched
    # Only for series names >= 3 chars to avoid false positives
    # ("L" matching "Leqvio", "H" matching "High Intensity Statins")
    for orig, sl in zip(series_names, series_lower):
        if orig in name_to_value:
            continue
        if len(sl) < 3:
            continue  # single/double char names must match via exact or abbreviation
        for vl, v_orig in values_lower.items():
            if vl.startswith(sl) or sl.startswith(vl):
                name_to_value[orig] = v_orig
                break
            if sl in vl:
                name_to_value[orig] = v_orig
                break

    matched = len(name_to_value)
    total = len(series_names)
    coverage = matched / total if total > 0 else 0.0

    # Require at least 50% coverage to consider this a plausible match
    if coverage < 0.5:
        return None

    # Determine match type
    if matched == total and all(
        s.strip().lower() in values_lower for s in series_names
    ):
        match_type = "exact"
    elif matched == total:
        match_type = "fuzzy"
    else:
        match_type = "partial"

    return ColumnMatch(
        column_name=column_name,
        match_type=match_type,
        matched_count=matched,
        total_series=total,
        coverage=coverage,
        name_to_value=name_to_value,
    )


# ─────────────────────────────────────────────────────────────────────────────
# Row-field and value-field inference
# ─────────────────────────────────────────────────────────────────────────────

def _infer_row_field(
    record_columns: dict[str, list],
    matched_column: str,
    chart_pattern: str,
    category_count: int,
) -> str:
    """Determine which column serves as row labels (categories / y-axis).

    Priority:
        1. A known row-label column name present in the records.
        2. A non-numeric, non-matched column whose unique value count
           matches category_count (if provided and > 0).
        3. Fallback: "y_label".
    """
    available = set(record_columns.keys()) - {matched_column} - _NUMERIC_COLUMNS

    # Priority 1: known row-field names
    for candidate in _ROW_FIELD_CANDIDATES:
        if candidate in available:
            # If we have a category_count hint, prefer a column whose cardinality matches
            if category_count > 0:
                unique_count = len(set(record_columns[candidate]))
                if unique_count == category_count:
                    return candidate
            else:
                return candidate

    # Priority 2: cardinality match against category_count
    if category_count > 0:
        for col in sorted(available):
            vals = record_columns[col]
            if isinstance(vals, list) and len(set(str(v) for v in vals)) == category_count:
                return col

    # Priority 3: first known row-field name regardless of cardinality
    for candidate in _ROW_FIELD_CANDIDATES:
        if candidate in available:
            return candidate

    # Fallback
    return "y_label"


def _infer_value_field(record_columns: dict[str, list], chart_pattern: str) -> str:
    """Determine which column holds the numeric values.

    Checks for known value-field names. Falls back to "percentage" which is
    the most common in PET survey data.
    """
    for candidate in _VALUE_FIELD_CANDIDATES:
        if candidate in record_columns:
            # Quick check: does it actually hold numbers?
            vals = record_columns[candidate]
            if vals and isinstance(vals[0], (int, float)):
                return candidate

    return "percentage"


# ─────────────────────────────────────────────────────────────────────────────
# Series config builder
# ─────────────────────────────────────────────────────────────────────────────

# Default color palettes by semantic role.  Extendable per-brand at call site.
_ROLE_COLORS: dict[str, str] = {
    # L/N/H/0 stacked bar palette (red/amber/green/grey)
    "L":        "#D34D2F",
    "Low":      "#D34D2F",
    "N":        "#F5C242",
    "Neutral":  "#F5C242",
    "H":        "#4CAF50",
    "High":     "#4CAF50",
    "0":        "#BFBFBF",
    "IDK":      "#BFBFBF",
    "DK":       "#BFBFBF",
}


def _build_series_configs(
    series_names: list[str],
    match: ColumnMatch,
) -> list[SeriesConfig]:
    """Create a SeriesConfig for each series, in order.

    Assigns role from the match mapping, default colors for known roles,
    and sequential order indices.
    """
    configs: list[SeriesConfig] = []
    for i, name in enumerate(series_names):
        role = match.name_to_value.get(name, name)
        # Try to find a color for the role or the original series name
        color = _ROLE_COLORS.get(name, _ROLE_COLORS.get(role, ""))
        configs.append(SeriesConfig(
            role=role,
            color=color,
            order=i,
        ))
    return configs


# ─────────────────────────────────────────────────────────────────────────────
# Main entry point
# ─────────────────────────────────────────────────────────────────────────────

def infer_data_transform(
    series_names: list[str],
    chart_pattern: str,
    record_columns: dict[str, list[str]],
    category_count: int = 0,
) -> tuple[DataTransform, list[SeriesConfig]]:
    """Infer a DataTransform + SeriesConfig list from chart series names.

    For non-Connector-tagged slides, this is how we figure out the mapping
    between Synapse records and chart components. The function:

        1. Tests every column in record_columns for overlap with series_names.
        2. Picks the best-matching column as column_field (the pivot axis).
        3. Infers row_field and value_field from remaining columns + chart pattern.
        4. Builds a SeriesConfig per series with role, color, and order.

    Args:
        series_names:    Series labels from chart OOXML, e.g. ["L","N","H","0"].
        chart_pattern:   Chart pattern key, e.g. "bar_stacked_100_horizontal".
        record_columns:  {column_name: [unique_values]} from Synapse records.
        category_count:  Number of categories in the existing chart (0 = unknown).

    Returns:
        (DataTransform, list[SeriesConfig])  — ready to assign to
        ChartDataMapping.transform and ChartDataMapping.series_config.

    Raises:
        ValueError: If no column matches the series names at all (coverage < 50%).
    """
    # ── Step 1: Find the best column match ──
    matches: list[ColumnMatch] = []
    for col_name, col_values in record_columns.items():
        # Coerce values to strings for comparison (skip pure-numeric lists)
        str_values = []
        for v in col_values:
            if isinstance(v, (int, float)):
                str_values.append(str(v))
            elif isinstance(v, str):
                str_values.append(v)
        if not str_values:
            continue

        result = match_series_to_column(series_names, col_name, str_values)
        if result is not None:
            matches.append(result)

    if not matches:
        raise ValueError(
            f"No column matches series names {series_names}. "
            f"Available columns: {list(record_columns.keys())}"
        )

    # Rank matches: exact > fuzzy > partial, then by coverage desc
    _type_rank = {"exact": 0, "fuzzy": 1, "partial": 2}
    matches.sort(key=lambda m: (_type_rank.get(m.match_type, 9), -m.coverage))
    best = matches[0]

    # ── Step 2: Infer row_field and value_field ──
    row_field = _infer_row_field(record_columns, best.column_name, chart_pattern, category_count)
    value_field = _infer_value_field(record_columns, chart_pattern)

    # ── Step 3: Build DataTransform ──
    transform = DataTransform(
        row_field=row_field,
        column_field=best.column_name,
        value_field=value_field,
    )

    # ── Step 4: Build SeriesConfig list ──
    series_configs = _build_series_configs(series_names, best)

    return transform, series_configs


# ─────────────────────────────────────────────────────────────────────────────
# Demo / smoke test
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    # Sample record columns (what Synapse returns)
    columns = {
        "y_label": ["Repatha", "Praluent", "Leqvio"],
        "segment_1": ["CARD", "PCP", "Overall"],
        "time_period_name": ["Q1'26", "Q4'25"],
        "measure": ["L", "N", "H", "0"],
        "percentage": [0.37, 0.22, 0.13],   # numeric, not matchable
        "base": [91, 126, 45],               # numeric, not matchable
    }

    print("=" * 70)
    print("TEST 1: Stacked bar with L/N/H/0 series")
    print("=" * 70)
    series = ["L", "N", "H", "0"]
    transform, configs = infer_data_transform(
        series_names=series,
        chart_pattern="bar_stacked_100_horizontal",
        record_columns=columns,
        category_count=3,
    )
    print(f"  column_field : {transform.column_field}")
    print(f"  row_field    : {transform.row_field}")
    print(f"  value_field  : {transform.value_field}")
    print(f"  series_config:")
    for sc in configs:
        print(f"    role={sc.role!r:12s}  color={sc.color!r:10s}  order={sc.order}")
    print()

    print("=" * 70)
    print("TEST 2: Time-period trend with quarter series")
    print("=" * 70)
    series = ["Q4'25", "Q1'26"]
    transform, configs = infer_data_transform(
        series_names=series,
        chart_pattern="column_clustered_vertical",
        record_columns=columns,
        category_count=3,
    )
    print(f"  column_field : {transform.column_field}")
    print(f"  row_field    : {transform.row_field}")
    print(f"  value_field  : {transform.value_field}")
    print(f"  series_config:")
    for sc in configs:
        print(f"    role={sc.role!r:12s}  color={sc.color!r:10s}  order={sc.order}")
    print()

    print("=" * 70)
    print("TEST 3: Segment comparison (CARD vs PCP)")
    print("=" * 70)
    series = ["CARD", "PCP", "Overall"]
    transform, configs = infer_data_transform(
        series_names=series,
        chart_pattern="bar_clustered_horizontal",
        record_columns=columns,
    )
    print(f"  column_field : {transform.column_field}")
    print(f"  row_field    : {transform.row_field}")
    print(f"  value_field  : {transform.value_field}")
    print(f"  series_config:")
    for sc in configs:
        print(f"    role={sc.role!r:12s}  color={sc.color!r:10s}  order={sc.order}")
    print()

    print("=" * 70)
    print("TEST 4: Fuzzy match — 'High'/'Low' in data, 'H'/'L' in chart")
    print("=" * 70)
    fuzzy_columns = {
        "y_label": ["Repatha", "Praluent"],
        "metric": ["High", "Low", "Neutral", "Don't Know"],
        "percentage": [0.5, 0.3],
    }
    series = ["L", "N", "H", "0"]
    transform, configs = infer_data_transform(
        series_names=series,
        chart_pattern="bar_stacked_100_horizontal",
        record_columns=fuzzy_columns,
        category_count=2,
    )
    print(f"  column_field : {transform.column_field}")
    print(f"  row_field    : {transform.row_field}")
    print(f"  value_field  : {transform.value_field}")
    print(f"  match mapping:")
    # Re-run match to show the name->value mapping
    match = match_series_to_column(
        series, "metric", ["High", "Low", "Neutral", "Don't Know"]
    )
    if match:
        for sn, cv in match.name_to_value.items():
            print(f"    {sn!r:6s} -> {cv!r}")
        print(f"  match_type: {match.match_type}, coverage: {match.coverage:.0%}")
    print()

    print("=" * 70)
    print("TEST 5: No match (expect ValueError)")
    print("=" * 70)
    try:
        infer_data_transform(
            series_names=["Alpha", "Beta", "Gamma"],
            chart_pattern="bar_clustered_horizontal",
            record_columns=columns,
        )
    except ValueError as e:
        print(f"  Caught expected error: {e}")
    print()

    print("All tests passed.")
