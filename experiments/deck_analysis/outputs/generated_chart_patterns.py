"""
Chart pattern definitions — deterministic rendering configs per chart type.

Generated from deck analysis of 551 decks (35,352 charts total).
Frequency-ranked. OOXML defaults from corpus-wide mode values.

OOXML property summary (corpus-wide):
  gapWidth modes:     {'100': 3881, '50': 2743, '150': 2573, '80': 1808, '70': 1674}
  overlap modes:      {'100': 15032, '-20': 640, '-27': 565, '-10': 398, '-5': 186}
  dLblPos modes:      {'ctr': 60660, 'outEnd': 21703, 't': 19097, 'inBase': 1307, 'r': 1284}
  numFmt modes:       {'0%': 36910, 'General': 2170, '0': 2142, '#,##0': 550, '0.0': 373}
  axis orientations:  {'val:minMax': 38246, 'cat:minMax': 18244, 'cat:maxMin': 14700, 'val:maxMin': 2275}
  tickLblPos modes:   {'val:nextTo': 39521, 'cat:nextTo': 29992, 'cat:none': 2839, 'val:high': 497, 'val:none': 489}
  marker types:       {'circle': 15872, 'square': 1900, 'diamond': 1064, 'triangle': 616, 'none': 417}
  line widths (EMU):  {'28575': 6917, '19050': 6745, '25400': 4809, '12700': 1836, '9525': 866}
  charts w/ legend:   592/35352
  charts w/ gridlines:3170/35352
  charts w/ title:    306/35352
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
        "occurrences": 9825,
        "pct": 27.8,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "column_stacked_100_vertical": {
        "occurrences": 5966,
        "pct": 16.9,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "line_markers_trended": {
        "occurrences": 5055,
        "pct": 14.3,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "bar_stacked_100_horizontal": {
        "occurrences": 3515,
        "pct": 9.9,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "xy_scatter_abacus": {
        "occurrences": 3221,
        "pct": 9.1,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "bar_stacked_horizontal": {
        "occurrences": 2579,
        "pct": 7.3,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "column_clustered_vertical": {
        "occurrences": 2202,
        "pct": 6.2,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "column_stacked_vertical": {
        "occurrences": 1263,
        "pct": 3.6,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "doughnut_default": {
        "occurrences": 1106,
        "pct": 3.1,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "pie_default": {
        "occurrences": 251,
        "pct": 0.7,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "xy_scatter_lines": {
        "occurrences": 178,
        "pct": 0.5,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "bubble": {
        "occurrences": 76,
        "pct": 0.2,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "unmapped_DOUGHNUT_EXPLODED": {
        "occurrences": 36,
        "pct": 0.1,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "unmapped_LINE_MARKERS_STACKED": {
        "occurrences": 25,
        "pct": 0.1,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "unmapped_RADAR_MARKERS": {
        "occurrences": 25,
        "pct": 0.1,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "unmapped_PIE_EXPLODED": {
        "occurrences": 11,
        "pct": 0.0,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "unmapped_AREA_STACKED_100": {
        "occurrences": 10,
        "pct": 0.0,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "unmapped_XY_SCATTER_SMOOTH": {
        "occurrences": 4,
        "pct": 0.0,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "area_stacked": {
        "occurrences": 2,
        "pct": 0.0,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "radar": {
        "occurrences": 1,
        "pct": 0.0,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
    "unknown": {
        "occurrences": 1,
        "pct": 0.0,
        "gap_width": DEFAULT_GAP_WIDTH,
        "overlap": DEFAULT_OVERLAP,
        "dlbl_pos": DEFAULT_DLBL_POS,
        "num_format": DEFAULT_NUM_FORMAT,
    },
}
