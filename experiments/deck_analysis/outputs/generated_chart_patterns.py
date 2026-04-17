"""
Chart pattern definitions — deterministic rendering configs per chart type.

Generated from deck analysis of 905 decks (50,072 charts total).
Frequency-ranked. OOXML defaults from corpus-wide mode values.

OOXML property summary (corpus-wide):
  gapWidth modes:     {'100': 5470, '50': 3889, '150': 3688, '80': 2452, '70': 2396}
  overlap modes:      {'100': 21992, '-20': 1118, '-27': 718, '-10': 516, '-5': 194}
  dLblPos modes:      {'ctr': 88318, 'outEnd': 29584, 't': 26192, 'inBase': 1871, 'r': 1702}
  numFmt modes:       {'0%': 52458, 'General': 3129, '0': 2641, '#,##0': 710, '0.0': 470}
  axis orientations:  {'val:minMax': 53945, 'cat:minMax': 26187, 'cat:maxMin': 20807, 'val:maxMin': 3063}
  tickLblPos modes:   {'val:nextTo': 55539, 'cat:nextTo': 42473, 'cat:none': 4319, 'val:high': 828, 'val:none': 622}
  marker types:       {'circle': 22043, 'square': 2166, 'diamond': 1221, 'triangle': 852, 'none': 631}
  line widths (EMU):  {'19050': 9736, '28575': 9381, '25400': 6422, '12700': 2121, '9525': 1257}
  charts w/ legend:   864/50072
  charts w/ gridlines:4323/50072
  charts w/ title:    516/50072
"""
from __future__ import annotations


# Corpus-wide OOXML defaults (mode values)
DEFAULT_GAP_WIDTH = 100
DEFAULT_OVERLAP = 100
DEFAULT_DLBL_POS = "ctr"
DEFAULT_NUM_FORMAT = "0%"
DEFAULT_MARKER_TYPE = "circle"


CHART_PATTERNS = {
    "bar_clustered_horizontal": {
        "occurrences": 13994,
        "pct": 27.9,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "column_stacked_100_vertical": {
        "occurrences": 8227,
        "pct": 16.4,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "line_markers_trended": {
        "occurrences": 6857,
        "pct": 13.7,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "bar_stacked_100_horizontal": {
        "occurrences": 5528,
        "pct": 11.0,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "xy_scatter_abacus": {
        "occurrences": 4260,
        "pct": 8.5,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "bar_stacked_horizontal": {
        "occurrences": 3637,
        "pct": 7.3,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "column_clustered_vertical": {
        "occurrences": 2869,
        "pct": 5.7,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "column_stacked_vertical": {
        "occurrences": 1801,
        "pct": 3.6,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "doughnut_default": {
        "occurrences": 1752,
        "pct": 3.5,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "pie_default": {
        "occurrences": 472,
        "pct": 0.9,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "xy_scatter_lines": {
        "occurrences": 207,
        "pct": 0.4,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "bubble": {
        "occurrences": 109,
        "pct": 0.2,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "unmapped_RADAR_FILLED": {
        "occurrences": 85,
        "pct": 0.2,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "unmapped_LINE_MARKERS_STACKED": {
        "occurrences": 77,
        "pct": 0.2,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "unmapped_DOUGHNUT_EXPLODED": {
        "occurrences": 71,
        "pct": 0.1,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "unmapped_PIE_EXPLODED": {
        "occurrences": 44,
        "pct": 0.1,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "unmapped_AREA_STACKED_100": {
        "occurrences": 25,
        "pct": 0.0,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "unmapped_RADAR_MARKERS": {
        "occurrences": 25,
        "pct": 0.0,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "area_stacked": {
        "occurrences": 14,
        "pct": 0.0,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "unmapped_XY_SCATTER_SMOOTH": {
        "occurrences": 6,
        "pct": 0.0,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "unmapped_AREA": {
        "occurrences": 5,
        "pct": 0.0,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "unmapped_XY_SCATTER_LINES_NO_MARKERS": {
        "occurrences": 2,
        "pct": 0.0,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "radar": {
        "occurrences": 2,
        "pct": 0.0,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "unknown": {
        "occurrences": 2,
        "pct": 0.0,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "unmapped_LINE_STACKED_100": {
        "occurrences": 1,
        "pct": 0.0,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
}
