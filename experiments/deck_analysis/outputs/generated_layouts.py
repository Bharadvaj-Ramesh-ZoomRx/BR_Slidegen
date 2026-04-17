"""
Layout presets for slide compositions.

Generated from deck analysis of 905 decks (36,914 slides).
Coordinate medians from observed chart + table positions across the full corpus.
"""
from __future__ import annotations


SLIDE_WIDTH = 13.333
SLIDE_HEIGHT = 7.5

HEADLINE_RECT = {"left": 0.2, "top": 0.3, "width": 12.8, "height": 0.9}
FOOTER_RECT = {"left": 0.2, "top": 7.0, "width": 12.8, "height": 0.4}


LAYOUTS = {

    # ---- 0_chart_0_table (12231 slides, 33.1%) ----
    #   chart positions: 0, table positions: 0
    "observed_0_chart_0_table": {
        "delta_col_rect": {"left": 9.0, "top": 2.19, "width": 0.65, "height": 4.02},
        "_slides": 12231,
    },

    # ---- 0_chart_1_table (3918 slides, 10.6%) ----
    #   chart positions: 0, table positions: 0
    "observed_0_chart_1_table": {
        "delta_col_rect": {"left": 9.0, "top": 2.19, "width": 0.65, "height": 4.02},
        "_slides": 3918,
    },

    # ---- 1_chart_1_table (3016 slides, 8.2%) ----
    #   chart positions: 1951, table positions: 2010
    "observed_1_chart_1_table": {
        "chart_rect": {"left": 4.97, "top": 2.04, "width": 4.9, "height": 4.22},
        "table_rect": {"left": 1.57, "top": 1.96, "width": 5.97, "height": 4.22},
        "delta_col_rect": {"left": 9.0, "top": 2.19, "width": 0.65, "height": 4.02},
        "_slides": 3016,
    },

    # ---- 1_chart_2_table (1666 slides, 4.5%) ----
    #   chart positions: 1178, table positions: 2379
    "observed_1_chart_2_table": {
        "chart_rect": {"left": 5.71, "top": 2.01, "width": 4.26, "height": 4.13},
        "table_rect": {"left": 4.26, "top": 1.98, "width": 3.55, "height": 3.94},
        "delta_col_rect": {"left": 9.0, "top": 2.19, "width": 0.65, "height": 4.02},
        "_slides": 1666,
    },

    # ---- 2_chart_2_table (1266 slides, 3.4%) ----
    #   chart positions: 1657, table positions: 1706
    "observed_2_chart_2_table": {
        "chart_rect": {"left": 5.94, "top": 2.35, "width": 2.93, "height": 3.78},
        "table_rect": {"left": 2.58, "top": 2.32, "width": 3.1, "height": 3.47},
        "delta_col_rect": {"left": 9.0, "top": 2.19, "width": 0.65, "height": 4.02},
        "_slides": 1266,
    },

    # ---- 2_chart_1_table (1239 slides, 3.4%) ----
    #   chart positions: 1600, table positions: 859
    "observed_2_chart_1_table": {
        "chart_rect": {"left": 6.16, "top": 2.28, "width": 3.05, "height": 3.72},
        "table_rect": {"left": 0.85, "top": 2.08, "width": 4.14, "height": 3.87},
        "delta_col_rect": {"left": 9.0, "top": 2.19, "width": 0.65, "height": 4.02},
        "_slides": 1239,
    },

    # ---- 1_chart_0_table (992 slides, 2.7%) ----
    #   chart positions: 0, table positions: 0
    "observed_1_chart_0_table": {
        "delta_col_rect": {"left": 9.0, "top": 2.19, "width": 0.65, "height": 4.02},
        "_slides": 992,
    },

    # ---- 0_chart_2_table (902 slides, 2.4%) ----
    #   chart positions: 0, table positions: 0
    "observed_0_chart_2_table": {
        "delta_col_rect": {"left": 9.0, "top": 2.19, "width": 0.65, "height": 4.02},
        "_slides": 902,
    },

    # ---- 2_chart_0_table (876 slides, 2.4%) ----
    #   chart positions: 0, table positions: 0
    "observed_2_chart_0_table": {
        "delta_col_rect": {"left": 9.0, "top": 2.19, "width": 0.65, "height": 4.02},
        "_slides": 876,
    },

    # ---- 3_chart_1_table (792 slides, 2.1%) ----
    #   chart positions: 1654, table positions: 598
    "observed_3_chart_1_table": {
        "chart_rect": {"left": 6.79, "top": 2.35, "width": 2.43, "height": 3.7},
        "table_rect": {"left": 0.56, "top": 2.14, "width": 4.21, "height": 3.89},
        "delta_col_rect": {"left": 9.0, "top": 2.19, "width": 0.65, "height": 4.02},
        "_slides": 792,
    },

    # ---- 1_chart_3_table (693 slides, 1.9%) ----
    #   chart positions: 496, table positions: 1508
    "observed_1_chart_3_table": {
        "chart_rect": {"left": 5.49, "top": 2.0, "width": 4.06, "height": 4.27},
        "table_rect": {"left": 5.74, "top": 1.99, "width": 2.7, "height": 3.86},
        "delta_col_rect": {"left": 9.0, "top": 2.19, "width": 0.65, "height": 4.02},
        "_slides": 693,
    },

    # ---- 3_chart_0_table (670 slides, 1.8%) ----
    #   chart positions: 0, table positions: 0
    "observed_3_chart_0_table": {
        "delta_col_rect": {"left": 9.0, "top": 2.19, "width": 0.65, "height": 4.02},
        "_slides": 670,
    },

    # ---- 2_chart_3_table (516 slides, 1.4%) ----
    #   chart positions: 715, table positions: 1106
    "observed_2_chart_3_table": {
        "chart_rect": {"left": 5.73, "top": 2.3, "width": 2.93, "height": 3.84},
        "table_rect": {"left": 5.41, "top": 2.24, "width": 1.93, "height": 3.25},
        "delta_col_rect": {"left": 9.0, "top": 2.19, "width": 0.65, "height": 4.02},
        "_slides": 516,
    },

    # ---- 0_chart_3_table (503 slides, 1.4%) ----
    #   chart positions: 0, table positions: 0
    "observed_0_chart_3_table": {
        "delta_col_rect": {"left": 9.0, "top": 2.19, "width": 0.65, "height": 4.02},
        "_slides": 503,
    },

    # ---- 3_chart_3_table (470 slides, 1.3%) ----
    #   chart positions: 848, table positions: 871
    "observed_3_chart_3_table": {
        "chart_rect": {"left": 6.11, "top": 2.39, "width": 2.5, "height": 3.43},
        "table_rect": {"left": 4.12, "top": 2.43, "width": 2.51, "height": 1.68},
        "delta_col_rect": {"left": 9.0, "top": 2.19, "width": 0.65, "height": 4.02},
        "_slides": 470,
    },

}


def get_layout(layout_key: str) -> dict:
    """Look up a layout preset by key."""
    if layout_key not in LAYOUTS:
        raise KeyError(
            f"Unknown layout {layout_key!r}. Available: {sorted(LAYOUTS.keys())}"
        )
    return LAYOUTS[layout_key]
