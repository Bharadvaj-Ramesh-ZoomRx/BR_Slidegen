"""
Named chart patterns — canonical defaults for each chart type.

Generated from real-deck analysis. Each pattern encodes the defaults observed
in client decks for that chart type. Renderers use these as starting points;
renderer args can override any field.
"""
from __future__ import annotations

from pptx.enum.chart import XL_CHART_TYPE


CHART_PATTERNS = {

    # ---- bar_clustered (horizontal) ----
    # 35% of all PET charts. Horizontal bars with inverted cat axis,
    # cat labels hidden (shown in companion table), data labels inside-end.
    "bar_clustered_horizontal": {
        "chart_type": XL_CHART_TYPE.BAR_CLUSTERED,
        "bar_dir": "bar",  # horizontal
        "grouping": "clustered",
        "gap_width": 80,
        "overlap": 0,
        "invert_cat_axis": True,
        "hide_cat_labels": True,
        "datalabel_pos": "inEnd",
        "val_num_format": "0%",
        "show_legend": False,
        "show_title": False,
        "major_gridlines": False,
        "series_no_border": True,
        "invert_if_negative": False,
    },

    # ---- bar_stacked_100 (horizontal) ----
    # Horizontal 100% stacked, typical for intent distributions and composition.
    "bar_stacked_100_horizontal": {
        "chart_type": XL_CHART_TYPE.BAR_STACKED_100,
        "bar_dir": "bar",
        "grouping": "percentStacked",
        "gap_width": 80,
        "overlap": 100,  # fully stacked
        "invert_cat_axis": True,
        "hide_cat_labels": True,
        "datalabel_pos": "ctr",
        "val_num_format": "0%",
        "show_legend": False,
        "show_title": False,
        "major_gridlines": False,
        "series_no_border": True,
        "invert_if_negative": False,
    },

    # ---- column_stacked_100 (vertical) ----
    # 11% of PET charts. Vertical 100% stacked for proportions.
    "column_stacked_100_vertical": {
        "chart_type": XL_CHART_TYPE.COLUMN_STACKED_100,
        "bar_dir": "col",
        "grouping": "percentStacked",
        "gap_width": 100,
        "overlap": 100,
        "invert_cat_axis": False,
        "hide_cat_labels": False,
        "datalabel_pos": "ctr",
        "val_num_format": "0%",
        "show_legend": False,
        "show_title": False,
        "major_gridlines": False,
        "series_no_border": True,
        "invert_if_negative": False,
    },

    # ---- column_clustered (vertical, QoQ) ----
    # Side-by-side bars for quarter-over-quarter comparison.
    "column_clustered_vertical": {
        "chart_type": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "bar_dir": "col",
        "grouping": "clustered",
        "gap_width": 100,
        "overlap": -20,  # slight gap between series
        "invert_cat_axis": False,
        "hide_cat_labels": False,
        "datalabel_pos": "outEnd",
        "val_num_format": "0%",
        "show_legend": False,
        "show_title": False,
        "major_gridlines": False,
        "series_no_border": True,
        "invert_if_negative": False,
    },

    # ---- line_markers (trended) ----
    # 14% of PET charts. Multi-wave trend lines with circle markers.
    "line_markers_trended": {
        "chart_type": XL_CHART_TYPE.LINE_MARKERS,
        "grouping": "standard",
        "datalabel_pos": "t",  # top of marker
        "val_num_format": "0%",
        "show_legend": False,
        "show_title": False,
        "major_gridlines": False,
        "marker_symbol": "circle",
        "marker_size": 7,
        "line_width_emu": 25400,  # 2pt
    },

    # ---- xy_scatter (abacus) ----
    # 17% of PET charts. Scatter for multi-attribute comparisons (abacus, MBD).
    "xy_scatter_abacus": {
        "chart_type": XL_CHART_TYPE.XY_SCATTER,
        "datalabel_pos": "t",
        "val_num_format": "0%",
        "show_legend": False,
        "show_title": False,
        "major_gridlines": False,
        "marker_symbol": "circle",
        "marker_size": 10,
        "hide_val_labels": False,
        "hide_cat_labels": False,
    },

    # ---- doughnut ----
    # 138 occurrences. Used for segment composition breakdowns.
    "doughnut_default": {
        "chart_type": XL_CHART_TYPE.DOUGHNUT,
        "datalabel_pos": "ctr",
        "val_num_format": "0%",
        "show_legend": False,
        "show_title": False,
        "hole_size": 50,  # percent
        "series_no_border": True,
    },

}


def get_pattern(pattern_name: str) -> dict:
    """Look up chart pattern by name. Raises KeyError with available patterns."""
    if pattern_name not in CHART_PATTERNS:
        raise KeyError(
            f"Unknown pattern {pattern_name!r}. Available: {sorted(CHART_PATTERNS.keys())}"
        )
    return CHART_PATTERNS[pattern_name].copy()
