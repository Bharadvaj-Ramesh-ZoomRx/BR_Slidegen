"""
Layout presets for slide compositions.

Generated from real-deck analysis — coordinate clusters of 410 distinct chart
positions and 490 table positions across 2,333 slides.

Each LAYOUTS entry is a dict of (left, top, width, height) in inches. Use in
renderers as the source of truth for coordinate placement.

Conventions:
  - chart_rect: main chart position
  - table_rect: main label/data table position
  - delta_col_rect: narrow delta column (typically ~0.5" wide)
  - headline_rect: top-of-slide headline text box
"""
from __future__ import annotations


# Slide canvas (widescreen 16:9)
SLIDE_WIDTH = 13.333
SLIDE_HEIGHT = 7.5

# Universal headline position (from headline analysis — median top 0.3",
# median width 11.94", median char count 72)
HEADLINE_RECT = {"left": 0.2, "top": 0.3, "width": 12.8, "height": 0.9}

# Universal footer position (used for source + footnotes)
FOOTER_RECT = {"left": 0.2, "top": 7.0, "width": 12.8, "height": 0.4}


# ============================================================================
# LAYOUT PRESETS — by slide signature
# ============================================================================

LAYOUTS = {

    # ---- single_bar_with_delta (1 chart + 1 table, 145 slides) ----
    # Most common pattern — label table + bar chart side by side
    "single_bar_with_delta": {
        "chart_rect": {"left": 1.62, "top": 2.02, "width": 5.41, "height": 4.29},
        "table_rect": {"left": 1.48, "top": 1.82, "width": 6.79, "height": 4.55},
        "delta_col_rect": {"left": 12.52, "top": 2.17, "width": 0.58, "height": 4.5},  # from 40-occurrence cluster
        "_source": "1_chart_1_table signature, 145 slides",
    },

    # ---- clustered_compare (1 chart + 2 tables, 110 slides) ----
    # Label table + delta table side by side with chart on right
    "clustered_compare": {
        "chart_rect": {"left": 6.75, "top": 2.04, "width": 3.81, "height": 4.3},
        "primary_table_rect": {"left": 3.11, "top": 1.98, "width": 3.88, "height": 4.17},
        "secondary_table_rect": {"left": 7.09, "top": 1.98, "width": 3.88, "height": 4.17},
        "_source": "1_chart_2_table signature, 110 slides",
    },

    # ---- dual_bar_with_delta (2 charts + 2 tables, 68 slides) ----
    # Two brand comparison — two bars + two tables side by side
    "dual_bar_with_delta": {
        "left_chart_rect": {"left": 2.6, "top": 2.25, "width": 3.0, "height": 3.69},
        "right_chart_rect": {"left": 6.84, "top": 2.42, "width": 3.0, "height": 3.69},
        "left_table_rect": {"left": 2.6, "top": 2.25, "width": 2.92, "height": 3.66},
        "right_table_rect": {"left": 3.82, "top": 2.25, "width": 2.92, "height": 3.66},
        "_source": "2_chart_2_table signature, 68 slides",
    },

    # ---- full_width_table (1 table only, 139 slides) ----
    # Large single table — Executive Summary, Recommendations
    "full_width_table": {
        "table_rect": {"left": 0.91, "top": 1.52, "width": 11.11, "height": 4.92},
        "_source": "1_table signature, 139 slides",
    },

    # ---- three_metric_scorecard (3 charts + 1 table, 45 slides) ----
    # Three-panel scorecard comparing metrics side by side
    "three_metric_scorecard": {
        "label_table_rect": {"left": 0.39, "top": 2.17, "width": 3.85, "height": 3.76},
        "chart_panel_template": {"left": 7.02, "top": 2.36, "width": 2.57, "height": 3.87},
        "chart_panel_count": 3,
        "chart_panel_gap": 0.1,
        "_source": "3_chart_1_table signature, 45 slides",
    },

    # ---- dual_chart_no_table (2 charts, 44 slides) ----
    # Two charts only — comparison without label tables
    "dual_chart_no_table": {
        "left_chart_rect": {"left": 2.4, "top": 2.33, "width": 4.19, "height": 3.67},
        "right_chart_rect": {"left": 6.89, "top": 2.33, "width": 4.19, "height": 3.67},
        "_source": "2_chart signature, 44 slides",
    },

}


# ============================================================================
# NARROW DELTA COLUMN POSITIONS — top 10 observed positions
# ============================================================================
#
# These are recurring "narrow column" positions observed in real decks.
# Typically ~0.5" wide × ~4.5" tall at specific left positions, used as
# standalone delta columns adjacent to charts.

DELTA_COL_POSITIONS = [
    {"left": 12.52, "top": 2.17, "width": 0.58, "height": 4.5, "_occurrences": 40},
    {"left": 11.46, "top": 2.08, "width": 0.47, "height": 4.58, "_occurrences": 22},
    {"left": 7.96, "top": 2.07, "width": 0.49, "height": 4.58, "_occurrences": 19},
    {"left": 10.59, "top": 2.06, "width": 0.53, "height": 4.52, "_occurrences": 18},
    {"left": 12.07, "top": 2.03, "width": 0.64, "height": 4.47, "_occurrences": 17},
    {"left": 9.6, "top": 1.94, "width": 0.44, "height": 4.54, "_occurrences": 16},
    {"left": 11.98, "top": 2.0, "width": 0.88, "height": 4.54, "_occurrences": 15},
    {"left": 5.99, "top": 2.39, "width": 0.63, "height": 4.01, "_occurrences": 15},
    {"left": 7.51, "top": 1.88, "width": 0.44, "height": 4.51, "_occurrences": 13},
    {"left": 12.0, "top": 2.38, "width": 0.63, "height": 4.38, "_occurrences": 12},
]


def get_layout(slide_type: str) -> dict:
    """Look up a layout preset by slide type."""
    if slide_type not in LAYOUTS:
        raise KeyError(
            f"Unknown slide_type {slide_type!r}. Available: {sorted(LAYOUTS.keys())}"
        )
    return LAYOUTS[slide_type]
