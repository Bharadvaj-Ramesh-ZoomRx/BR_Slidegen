"""
layout.py — Slide chrome, decorations, and layout presets.

Slide-level structural elements: headers, footers, section bars,
badges, cover slides, dividers, legends, trend arrows, quadrant fills.

PRD §4.4: LAYOUTS{} dict with named presets.
"""

import os

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
}


# ── Slide chrome functions ───────────────────────────────────────────────────

def slide_header(slide, headline, module_label="Personal Promotion Module",
                 font=None):
    """Add the standard ZoomRx slide header:
      - thin red accent line at very top
      - module label (top-right, small grey)
      - headline text (large, red, bold)
      - red module badge (top-right rectangle)
      - separator line below header
    """
    f_display = font or FONT_DISPLAY
    f_text = font or FONT_TEXT

    # Red accent line
    solidrect(slide, 0, 0.15, SLIDE_W_IN, 0.02, C_RED)

    # Module label
    textbox(slide, module_label,
            5.5, 0.01, 7.70, 0.22,
            fsize=7.5, color=C_FTGREY, align=PP_ALIGN.RIGHT, font=f_text)

    # Headline
    textbox(slide, headline,
            0.20, 0.20, 10.55, 1.05,
            fsize=12, bold=True, color=C_RED, align=PP_ALIGN.LEFT, font=f_display)

    # Module badge
    solidrect(slide, 10.90, 0.20, 2.25, 0.95, C_RED)
    textbox(slide, "PERSONAL\nPROMOTION\nMODULE",
            10.90, 0.20, 2.25, 0.95,
            fsize=8, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER, font=f_display)

    # Separator line
    horiz_line(slide, 0.0, 1.32, SLIDE_W_IN, color=C_RED, width_pt=1.0)


def slide_footer(slide, footer_text, font=None):
    """Add standard footer text at bottom of slide."""
    textbox(slide, footer_text,
            0.15, 7.20, 13.0, 0.28,
            fsize=6.0, color=C_FTGREY, align=PP_ALIGN.LEFT, font=font or FONT_TEXT)


def manual_legend(slide, q4_n, q3_n,
                  chart_left, chart_right, chart_bottom):
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


def section_header_bar(slide, label, top=1.40, icon_path=None, font=None):
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
            fsize=9, bold=True, color=C_GREY, font=font or FONT_TEXT)
    return bg


def module_badge(slide, label, color=None):
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


def section_breadcrumb(slide, text):
    """Add small right-aligned breadcrumb text in the top-right."""
    textbox(slide, text,
            8.0, 0.36, 5.20, 0.20,
            fsize=7, color=C_FTGREY,
            align=PP_ALIGN.RIGHT, font=FONT_TEXT)


def divider_slide(slide, title, logo_path=None):
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


def cover_slide(slide, title, subtitle, date, client_name,
                jj_logo_path=None, zrx_logo_path=None):
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


def trend_arrow_icon(slide, direction, left, top):
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


def scatter_quadrant_fills(slide, chart_left, chart_top,
                            chart_width, chart_height,
                            tl_color=None, tr_color=None,
                            bl_color=None, br_color=None):
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
