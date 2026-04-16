"""
layout_selector.py — Deterministic layout selector.

Picks a LAYOUTS{} key for a given set of slide components.
Pure function — no LLM calls, no I/O, no randomness.

Rules (from deck-analysis ground truth, 32 PET decks):
  - 1 chart + 1 table          → observed_1chart_1table  (145 slides)
  - 1 chart + 2 tables         → observed_1chart_2table  (110 slides)
  - 2 charts + 2 tables        → observed_2chart_2table  (68 slides)
  - 1 chart + 0 tables         → single_chart_with_delta (or observed_dual_chart_no_table)
  - 0 charts + 1 table (full)  → observed_full_width_table (139 slides)
  - 3 charts + 1 table         → observed_three_metric_scorecard (45 slides)
  - 2 charts + 0 tables        → observed_dual_chart_no_table (44 slides)
"""
from __future__ import annotations


# Component type constants
_CHART_TYPES = {"chart"}
_TABLE_TYPES = {"label_table", "value_table"}
# Note: delta_column is structural — it always pairs with a chart and is part of
# the layout preset (observed_*_rect positions). For layout selection we do NOT
# count it as a separate table component; the canonical `1chart_1table` pattern
# includes a chart + companion label_table + delta_column together (see
# `slidegen/slide_spec/examples/bar_clustered_with_delta.json`).


def select_layout(components: list[str]) -> str:
    """Select a LAYOUTS{} key based on component types present.

    Args:
        components: List of component type strings (e.g. ["chart", "label_table", "delta_column"]).
                    Values must be from SUPPORTED_COMPONENT_TYPES.

    Returns:
        A key from LAYOUTS{} in slidegen/pptx_utils/layout.py.
    """
    chart_count = sum(1 for c in components if c in _CHART_TYPES)
    table_count = sum(1 for c in components if c in _TABLE_TYPES)

    # Special case: no charts at all
    if chart_count == 0:
        if table_count >= 1:
            return "observed_full_width_table"
        # No charts, no tables — likely a textbox-only slide (cover, ES)
        return "observed_full_width_table"

    # 1 chart
    if chart_count == 1:
        if table_count == 0:
            return "single_chart_with_delta"
        elif table_count == 1:
            return "observed_1chart_1table"
        elif table_count == 2:
            return "observed_1chart_2table"
        else:
            # 3+ tables with one chart — extended compare layout
            return "observed_1chart_2table"

    # 2 charts
    if chart_count == 2:
        if table_count == 0:
            return "observed_dual_chart_no_table"
        elif table_count <= 2:
            return "observed_2chart_2table"
        else:
            return "observed_2chart_2table"

    # 3+ charts
    if chart_count >= 3:
        return "observed_three_metric_scorecard"

    # Fallback (should not be reached)
    return "observed_1chart_1table"


def describe_layout(layout_key: str) -> str:
    """Return a human-readable description of a layout key."""
    descriptions = {
        "observed_1chart_1table": "Single chart + label table + delta column (145 real slides)",
        "observed_1chart_2table": "Single chart + two companion tables (110 real slides)",
        "observed_2chart_2table": "Dual chart + two tables (68 real slides)",
        "observed_full_width_table": "Full-width table or text-only (139 real slides)",
        "observed_three_metric_scorecard": "Three-panel scorecard (45 real slides)",
        "observed_dual_chart_no_table": "Dual chart without tables (44 real slides)",
        "single_chart_with_delta": "Single chart with delta column (no companion table)",
        "dual_chart_with_delta": "Dual chart with paired delta columns",
        "clustered_compare": "Clustered comparison with gap + delta columns",
    }
    return descriptions.get(layout_key, f"Layout: {layout_key}")
