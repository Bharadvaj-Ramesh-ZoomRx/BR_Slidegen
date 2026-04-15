"""
layout.py — Slide chrome, decorations, and layout presets.

Slide-level structural elements: headers, footers, section bars,
badges, cover slides, dividers, legends, trend arrows, quadrant fills.

PRD §4.4: LAYOUTS{} dict with named presets.
"""

from __future__ import annotations

import os
from typing import Optional

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from .brand import (
    SLIDE_W_IN, SLIDE_H_IN,
    C_RED, C_WHITE, C_GREY, C_FTGREY, C_LBGREY, C_LTGREY, C_GREEN,
    C_RYB_Q4, C_RYB_Q3,
    FONT_DISPLAY, FONT_TEXT,
)
from .shapes import textbox, solidrect, horiz_line
from .lxml_helpers import _get_or_add


# ── LAYOUTS{} dict (PRD §4.4) ───────────────────────────────────────────────

LAYOUTS = {
    "single_chart_with_delta": {
        "chart_left": 0.30,
        "chart_top": 1.85,
        "chart_w": 9.0,
        "delta_left": 9.40,
        "delta_w": 0.65,
        "legend_y_offset": 0.15,
    },
    "dual_chart_with_delta": {
        "left_chart_left": 0.30,
        "left_chart_w": 5.0,
        "left_delta_left": 5.35,
        "left_delta_w": 0.65,
        "right_chart_left": 6.20,
        "right_chart_w": 5.0,
        "right_delta_left": 11.25,
        "right_delta_w": 0.65,
        "chart_top": 1.85,
    },
    "clustered_compare": {
        "chart_left": 0.30,
        "chart_top": 1.85,
        "chart_w": 8.0,
        "gap_col_left": 8.35,
        "gap_col_w": 0.65,
        "delta_col_left": 9.05,
        "delta_col_w": 0.65,
    },
    "two_section": {
        "top_section_top": 1.85,
        "bottom_section_top": 4.50,
        "section_height": 2.40,
        "chart_left": 0.30,
        "chart_w": 9.0,
    },
    "stacked_order": {
        "chart_left": 0.30,
        "chart_top": 1.85,
        "chart_w": 10.0,
        "total_col_left": 10.35,
        "total_col_w": 0.65,
    },

    # ────────────────────────────────────────────────────────────────────────
    # Layouts derived from real-deck coordinate analysis (32 PET decks).
    # Each entry lists the median coordinates from its signature cluster.
    # Use these as the default geometry; renderers can override specific
    # fields at call time if a deck needs tweaking.
    # Source: experiments/deck_analysis/outputs/layout_clusters.md
    # ────────────────────────────────────────────────────────────────────────

    # 145 slides — canonical label-table + bar-chart pattern
    "observed_1chart_1table": {
        "chart_rect": {"left": 1.62, "top": 2.02, "width": 5.41, "height": 4.29},
        "table_rect": {"left": 1.48, "top": 1.82, "width": 6.79, "height": 4.55},
        "delta_col_rect": {"left": 12.52, "top": 2.17, "width": 0.58, "height": 4.50},
        "_source": "1_chart_1_table signature, 145 slides",
    },

    # 110 slides — clustered compare (chart right + two tables left)
    "observed_1chart_2table": {
        "chart_rect": {"left": 6.75, "top": 2.04, "width": 3.81, "height": 4.30},
        "primary_table_rect": {"left": 3.11, "top": 1.98, "width": 3.88, "height": 4.17},
        "secondary_table_rect": {"left": 7.09, "top": 1.98, "width": 3.88, "height": 4.17},
        "_source": "1_chart_2_table signature, 110 slides",
    },

    # 68 slides — dual bar comparison (two charts + two tables)
    "observed_2chart_2table": {
        "left_chart_rect": {"left": 2.60, "top": 2.25, "width": 3.00, "height": 3.69},
        "right_chart_rect": {"left": 6.84, "top": 2.42, "width": 3.00, "height": 3.69},
        "left_table_rect": {"left": 2.60, "top": 2.25, "width": 2.92, "height": 3.66},
        "right_table_rect": {"left": 3.84, "top": 2.25, "width": 2.92, "height": 3.66},
        "_source": "2_chart_2_table signature, 68 slides",
    },

    # 139 slides — large single table (Executive Summary / Recommendations)
    "observed_full_width_table": {
        "table_rect": {"left": 0.91, "top": 1.52, "width": 11.11, "height": 4.92},
        "_source": "1_table signature, 139 slides",
    },

    # 45 slides — 3-panel scorecard
    "observed_three_metric_scorecard": {
        "label_table_rect": {"left": 0.39, "top": 2.17, "width": 3.85, "height": 3.76},
        "chart_panel_template": {"left": 7.02, "top": 2.36, "width": 2.57, "height": 3.87},
        "chart_panel_count": 3,
        "chart_panel_gap": 0.10,
        "_source": "3_chart_1_table signature, 45 slides",
    },

    # 44 slides — two charts, no tables (side-by-side comparison)
    "observed_dual_chart_no_table": {
        "left_chart_rect": {"left": 2.40, "top": 2.33, "width": 4.19, "height": 3.67},
        "right_chart_rect": {"left": 6.89, "top": 2.33, "width": 4.19, "height": 3.67},
        "_source": "2_chart signature, 44 slides",
    },
}


# ── Universal positions (from deck analysis) ────────────────────────────────
# Applicable to most slides regardless of content type.

HEADLINE_RECT = {"left": 0.20, "top": 0.30, "width": 12.80, "height": 0.90}
FOOTER_RECT = {"left": 0.20, "top": 7.00, "width": 12.80, "height": 0.40}

# Narrow delta-column positions observed in real decks (top 5, most common first).
# Each has width ≤ 1.0", height ≥ 3.0" — canonical delta-column geometry.
OBSERVED_DELTA_COL_POSITIONS = [
    {"left": 12.52, "top": 2.17, "width": 0.58, "height": 4.50, "_occurrences": 40},
    {"left": 11.46, "top": 2.08, "width": 0.47, "height": 4.58, "_occurrences": 22},
    {"left": 7.96,  "top": 2.07, "width": 0.49, "height": 4.58, "_occurrences": 19},
    {"left": 10.59, "top": 2.06, "width": 0.53, "height": 4.52, "_occurrences": 18},
    {"left": 12.07, "top": 2.03, "width": 0.64, "height": 4.47, "_occurrences": 17},
]


def get_layout(slide_type: str) -> dict:
    """Look up a layout preset by slide type."""
    if slide_type not in LAYOUTS:
        raise KeyError(
            f"Unknown slide_type {slide_type!r}. Available: {sorted(LAYOUTS.keys())}"
        )
    return LAYOUTS[slide_type]


# ── Slide chrome functions ───────────────────────────────────────────────────

def slide_header(slide, headline: str, module_label: str = "Personal Promotion Module",
                 font: Optional[str] = None) -> None:
    """Add the standard ZoomRx slide header:
      - module label (top-right, small grey)
      - headline text (large, red, bold)
      - red module badge (top-right rectangle)
    """
    f_display = font or FONT_DISPLAY
    f_text = font or FONT_TEXT

    # Module label + badge (skip if module_label is empty)
    if module_label:
        # Short badge label: drop trailing "Module" if present for cleaner fit
        badge_label = module_label.replace(" Module", "").upper()
        textbox(slide, module_label,
                10.49, 0.02, 2.85, 0.25,
                fsize=9, color=C_FTGREY, align=PP_ALIGN.RIGHT, font=f_display)

        badge_w = 1.92
        badge_h = 0.27
        badge_l = SLIDE_W_IN - badge_w - 0.10
        badge_t = 1.06
        solidrect(slide, badge_l, badge_t, badge_w, badge_h, C_RED)
        textbox(slide, badge_label,
                badge_l, badge_t, badge_w, badge_h,
                fsize=10, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER,
                font=f_display)

    # Headline (full width — badge overlaps if needed)
    textbox(slide, headline,
            0.20, 0.16, 12.08, 1.21,
            fsize=22, bold=True, color=C_RED, align=PP_ALIGN.LEFT, font=f_display)


def slide_footer(slide, footer_text: str, font: Optional[str] = None) -> None:
    """Add standard footer text at bottom of slide.

    Full-width below the chart area. Font size 9pt for readability.
    """
    sh = textbox(slide, footer_text,
                 0.20, 7.10, 12.70, 0.25,
                 fsize=9.0, color=C_FTGREY, align=PP_ALIGN.LEFT, font=font or FONT_TEXT)
    _get_or_add(sh.text_frame._txBody, "a:bodyPr").set("anchor", "ctr")


def chart_header_row(slide, columns: list[dict], top: float, height: float = 0.30,
                     bg_color: Optional[RGBColor] = None,
                     font: Optional[str] = None) -> None:
    """Add a colored header row above the chart area with column titles.

    Each column dict has: label (str), left (float), width (float),
    and optional align (PP_ALIGN, default CENTER).

    The background spans the full slide width. Column labels are positioned
    individually within the row.

    Args:
        columns: [{"label": "Tag$", "left": 0.20, "width": 2.0}, ...]
        top: Y position in inches
        height: row height in inches (default 0.30)
        bg_color: background color (default C_RED)
        font: font name override
    """
    fill = bg_color or C_RED
    f = font or FONT_TEXT

    # Full-width background strip
    solidrect(slide, 0, top, SLIDE_W_IN, height, fill)

    # Column labels
    for col in columns:
        align = col.get("align", PP_ALIGN.CENTER)
        textbox(slide, col["label"],
                col["left"], top, col["width"], height,
                fsize=8, bold=True, color=C_WHITE, align=align, font=f)


def manual_legend(slide, q4_n: int, q3_n: int,
                  chart_left: float, chart_right: float,
                  chart_bottom: float) -> None:
    """Add the shared Q3/Q4/delta colour legend below a pair of charts."""
    leg_top = chart_bottom + 0.10
    leg_ctr = (chart_left + chart_right) / 2
    leg_l   = leg_ctr - 2.2

    solidrect(slide, leg_l,        leg_top, 0.20, 0.14, C_RYB_Q4)
    textbox(slide,  f"Q4'25  (n={q4_n})",
            leg_l + 0.25, leg_top - 0.01, 1.20, 0.18, fsize=8, color=C_GREY)

    solidrect(slide, leg_l + 1.65, leg_top, 0.20, 0.14, C_RYB_Q3)
    textbox(slide,  f"Q3'25  (n={q3_n})",
            leg_l + 1.90, leg_top - 0.01, 1.20, 0.18, fsize=8, color=C_GREY)

    solidrect(slide, leg_l + 3.35, leg_top, 0.20, 0.14, C_GREEN)
    textbox(slide,  "Delta increase (vs Q3)",
            leg_l + 3.60, leg_top - 0.01, 1.40, 0.18, fsize=8, color=C_GREY)

    solidrect(slide, leg_l + 5.15, leg_top, 0.20, 0.14, C_RED)
    textbox(slide,  "Delta decrease (vs Q3)   N/A = new message (no Q3)",
            leg_l + 5.40, leg_top - 0.01, 3.00, 0.18, fsize=8, color=C_GREY)


def section_header_bar(slide, label: str, top: float = 1.40,
                       icon_path: Optional[str] = None,
                       font: Optional[str] = None):
    """Add the gray icon+label strip used as a chart/question title."""
    bar_l = 0.15
    bar_w = SLIDE_W_IN - 0.30
    bar_h = 0.28

    bg = solidrect(slide, bar_l, top, bar_w, bar_h, C_LBGREY, line=C_LTGREY)

    text_l = bar_l + 0.10
    if icon_path and os.path.exists(icon_path):
        slide.shapes.add_picture(
            icon_path,
            Inches(bar_l + 0.05), Inches(top + 0.04),
            Inches(0.20),          Inches(0.20))
        text_l = bar_l + 0.32

    textbox(slide, label.upper(),
            text_l, top + 0.03, bar_w - 0.35, bar_h - 0.06,
            fsize=12, bold=True, color=C_GREY, font=font or FONT_TEXT)
    return bg


def module_badge(slide, label: str, color: Optional[RGBColor] = None) -> None:
    """Add a colored pill badge in the top-right corner."""
    if color is None:
        color = C_RED
    badge_w = 2.25
    badge_h = 0.28
    badge_l = SLIDE_W_IN - badge_w - 0.10
    badge_t = 0.05
    solidrect(slide, badge_l, badge_t, badge_w, badge_h, color)
    textbox(slide, label.upper(),
            badge_l, badge_t, badge_w, badge_h,
            fsize=7, bold=True, color=C_WHITE,
            align=PP_ALIGN.CENTER, font=FONT_TEXT)


def section_breadcrumb(slide, text: str) -> None:
    """Add small right-aligned breadcrumb text in the top-right."""
    textbox(slide, text,
            8.0, 0.36, 5.20, 0.20,
            fsize=7, color=C_FTGREY,
            align=PP_ALIGN.RIGHT, font=FONT_TEXT)


def divider_slide(slide, title: str, logo_path: Optional[str] = None) -> None:
    """Build a full-slide section divider."""
    solidrect(slide, 0, 0, SLIDE_W_IN, SLIDE_H_IN, RGBColor(0xF8, 0xF8, 0xF8))
    solidrect(slide, 0, 0, 0.08, SLIDE_H_IN, C_RED)
    textbox(slide, title,
            0.60, 2.30, 9.0, 3.0,
            fsize=36, bold=True, color=C_RED,
            align=PP_ALIGN.LEFT, font=FONT_DISPLAY)
    if logo_path and os.path.exists(logo_path):
        slide.shapes.add_picture(
            logo_path,
            Inches(0.30), Inches(SLIDE_H_IN - 0.80),
            Inches(1.50), Inches(0.50))


def cover_slide(slide, title: str, subtitle: str, date: str, client_name: str,
                jj_logo_path: Optional[str] = None,
                zrx_logo_path: Optional[str] = None) -> None:
    """Build the full red-background title cover slide."""
    solidrect(slide, 0, 0, SLIDE_W_IN, SLIDE_H_IN, C_RED)
    textbox(slide, title,
            0.60, 1.50, 11.0, 2.50,
            fsize=32, bold=True, color=C_WHITE,
            align=PP_ALIGN.LEFT, font=FONT_DISPLAY)
    textbox(slide, subtitle,
            0.60, 4.20, 10.0, 0.80,
            fsize=16, color=C_WHITE,
            align=PP_ALIGN.LEFT, font=FONT_TEXT)
    textbox(slide, date,
            0.60, 5.10, 5.0, 0.40,
            fsize=10, color=C_WHITE, font=FONT_TEXT)
    textbox(slide, client_name,
            7.0, 6.80, 6.10, 0.40,
            fsize=8, color=C_WHITE,
            align=PP_ALIGN.RIGHT, font=FONT_TEXT)
    if jj_logo_path and os.path.exists(jj_logo_path):
        slide.shapes.add_picture(
            jj_logo_path,
            Inches(0.40), Inches(6.70), Inches(1.50), Inches(0.50))
    if zrx_logo_path and os.path.exists(zrx_logo_path):
        slide.shapes.add_picture(
            zrx_logo_path,
            Inches(2.10), Inches(6.70), Inches(1.50), Inches(0.50))


def trend_arrow_icon(slide, direction: str, left: float, top: float) -> None:
    """Add a small directional trend arrow icon at a position."""
    colors = {"up": C_GREEN, "down": C_RED,
              "flat": RGBColor(0xFF, 0xC0, 0x00)}
    chars  = {"up": u"\u25b2", "down": u"\u25bc", "flat": u"\u25c4\u25ba"}
    color  = colors.get(direction, C_GREY)
    char   = chars.get(direction, "?")
    textbox(slide, char,
            left, top, 0.22, 0.22,
            fsize=9, bold=True, color=color,
            align=PP_ALIGN.CENTER, font=FONT_TEXT)


def scatter_quadrant_fills(slide, chart_left: float, chart_top: float,
                            chart_width: float, chart_height: float,
                            tl_color: Optional[RGBColor] = None,
                            tr_color: Optional[RGBColor] = None,
                            bl_color: Optional[RGBColor] = None,
                            br_color: Optional[RGBColor] = None) -> None:
    """Place four colored background rectangles behind a scatter chart."""
    if tl_color is None: tl_color = RGBColor(0xDE, 0xEB, 0xF7)
    if tr_color is None: tr_color = RGBColor(0xE2, 0xEF, 0xDA)
    if bl_color is None: bl_color = RGBColor(0xF4, 0xF4, 0xF4)
    if br_color is None: br_color = RGBColor(0xFF, 0xFF, 0xCC)

    cx = chart_left  + chart_width  / 2
    cy = chart_top   + chart_height / 2
    hw = chart_width  / 2
    hh = chart_height / 2

    solidrect(slide, chart_left, chart_top, hw, hh, tl_color)
    solidrect(slide, cx,         chart_top, hw, hh, tr_color)
    solidrect(slide, chart_left, cy,        hw, hh, bl_color)
    solidrect(slide, cx,         cy,        hw, hh, br_color)
