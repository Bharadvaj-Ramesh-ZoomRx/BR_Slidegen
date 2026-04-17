"""
Layout presets for slide compositions.

Generated from deck analysis of 551 decks (24,740 slides).
Coordinate medians from observed chart + table positions across the full corpus.
"""
from __future__ import annotations


SLIDE_WIDTH = 13.333
SLIDE_HEIGHT = 7.5

HEADLINE_RECT = {"left": 0.2, "top": 0.3, "width": 12.8, "height": 0.9}
FOOTER_RECT = {"left": 0.2, "top": 7.0, "width": 12.8, "height": 0.4}


LAYOUTS = {

    # ---- 0_chart_0_table (7606 slides, 30.7%) ----
    #   chart positions: 0, table positions: 0
    "observed_0_chart_0_table": {
        "delta_col_rect": {"left": 9.24, "top": 2.19, "width": 0.63, "height": 4.09},
        "_slides": 7606,
    },

    # ---- 0_chart_1_table (2576 slides, 10.4%) ----
    #   chart positions: 0, table positions: 0
    "observed_0_chart_1_table": {
        "delta_col_rect": {"left": 9.24, "top": 2.19, "width": 0.63, "height": 4.09},
        "_slides": 2576,
    },

    # ---- 1_chart_1_table (2056 slides, 8.3%) ----
    #   chart positions: 1291, table positions: 1335
    "observed_1_chart_1_table": {
        "chart_rect": {"left": 4.97, "top": 2.01, "width": 4.96, "height": 4.3},
        "table_rect": {"left": 1.59, "top": 1.89, "width": 6.22, "height": 4.33},
        "delta_col_rect": {"left": 9.24, "top": 2.19, "width": 0.63, "height": 4.09},
        "_slides": 2056,
    },

    # ---- 1_chart_2_table (1175 slides, 4.7%) ----
    #   chart positions: 839, table positions: 1687
    "observed_1_chart_2_table": {
        "chart_rect": {"left": 5.89, "top": 2.04, "width": 4.28, "height": 4.13},
        "table_rect": {"left": 4.22, "top": 1.99, "width": 3.65, "height": 3.98},
        "delta_col_rect": {"left": 9.24, "top": 2.19, "width": 0.63, "height": 4.09},
        "_slides": 1175,
    },

    # ---- 2_chart_2_table (859 slides, 3.5%) ----
    #   chart positions: 1165, table positions: 1214
    "observed_2_chart_2_table": {
        "chart_rect": {"left": 5.95, "top": 2.37, "width": 2.87, "height": 3.77},
        "table_rect": {"left": 2.52, "top": 2.33, "width": 3.15, "height": 3.47},
        "delta_col_rect": {"left": 9.24, "top": 2.19, "width": 0.63, "height": 4.09},
        "_slides": 859,
    },

    # ---- 2_chart_1_table (775 slides, 3.1%) ----
    #   chart positions: 995, table positions: 538
    "observed_2_chart_1_table": {
        "chart_rect": {"left": 6.25, "top": 2.23, "width": 3.03, "height": 3.74},
        "table_rect": {"left": 0.88, "top": 2.08, "width": 4.27, "height": 3.87},
        "delta_col_rect": {"left": 9.24, "top": 2.19, "width": 0.63, "height": 4.09},
        "_slides": 775,
    },

    # ---- 0_chart_2_table (624 slides, 2.5%) ----
    #   chart positions: 0, table positions: 0
    "observed_0_chart_2_table": {
        "delta_col_rect": {"left": 9.24, "top": 2.19, "width": 0.63, "height": 4.09},
        "_slides": 624,
    },

    # ---- 1_chart_0_table (609 slides, 2.5%) ----
    #   chart positions: 0, table positions: 0
    "observed_1_chart_0_table": {
        "delta_col_rect": {"left": 9.24, "top": 2.19, "width": 0.63, "height": 4.09},
        "_slides": 609,
    },

    # ---- 2_chart_0_table (541 slides, 2.2%) ----
    #   chart positions: 0, table positions: 0
    "observed_2_chart_0_table": {
        "delta_col_rect": {"left": 9.24, "top": 2.19, "width": 0.63, "height": 4.09},
        "_slides": 541,
    },

    # ---- 1_chart_3_table (527 slides, 2.1%) ----
    #   chart positions: 384, table positions: 1170
    "observed_1_chart_3_table": {
        "chart_rect": {"left": 5.49, "top": 2.0, "width": 4.08, "height": 4.3},
        "table_rect": {"left": 6.07, "top": 2.0, "width": 2.62, "height": 3.87},
        "delta_col_rect": {"left": 9.24, "top": 2.19, "width": 0.63, "height": 4.09},
        "_slides": 527,
    },

    # ---- 3_chart_1_table (526 slides, 2.1%) ----
    #   chart positions: 1148, table positions: 414
    "observed_3_chart_1_table": {
        "chart_rect": {"left": 6.95, "top": 2.34, "width": 2.47, "height": 3.7},
        "table_rect": {"left": 0.56, "top": 2.16, "width": 4.34, "height": 3.87},
        "delta_col_rect": {"left": 9.24, "top": 2.19, "width": 0.63, "height": 4.09},
        "_slides": 526,
    },

    # ---- 3_chart_0_table (445 slides, 1.8%) ----
    #   chart positions: 0, table positions: 0
    "observed_3_chart_0_table": {
        "delta_col_rect": {"left": 9.24, "top": 2.19, "width": 0.63, "height": 4.09},
        "_slides": 445,
    },

    # ---- 2_chart_3_table (406 slides, 1.6%) ----
    #   chart positions: 550, table positions: 852
    "observed_2_chart_3_table": {
        "chart_rect": {"left": 5.79, "top": 2.34, "width": 2.89, "height": 3.78},
        "table_rect": {"left": 5.2, "top": 2.27, "width": 2.05, "height": 3.05},
        "delta_col_rect": {"left": 9.24, "top": 2.19, "width": 0.63, "height": 4.09},
        "_slides": 406,
    },

    # ---- 3_chart_3_table (373 slides, 1.5%) ----
    #   chart positions: 697, table positions: 715
    "observed_3_chart_3_table": {
        "chart_rect": {"left": 6.11, "top": 2.41, "width": 2.59, "height": 3.41},
        "table_rect": {"left": 3.65, "top": 2.48, "width": 2.86, "height": 1.16},
        "delta_col_rect": {"left": 9.24, "top": 2.19, "width": 0.63, "height": 4.09},
        "_slides": 373,
    },

    # ---- 3_chart_2_table (342 slides, 1.4%) ----
    #   chart positions: 716, table positions: 511
    "observed_3_chart_2_table": {
        "chart_rect": {"left": 7.64, "top": 2.31, "width": 1.9, "height": 3.9},
        "table_rect": {"left": 2.97, "top": 2.19, "width": 3.14, "height": 3.64},
        "delta_col_rect": {"left": 9.24, "top": 2.19, "width": 0.63, "height": 4.09},
        "_slides": 342,
    },

}


def get_layout(layout_key: str) -> dict:
    """Look up a layout preset by key."""
    if layout_key not in LAYOUTS:
        raise KeyError(
            f"Unknown layout {layout_key!r}. Available: {sorted(LAYOUTS.keys())}"
        )
    return LAYOUTS[layout_key]
