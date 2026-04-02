"""
shapes.py — Layout primitives for slide creation.

All positions/sizes in inches. Uses python-pptx API (no raw lxml except
for dashed_separator which needs prstDash).
"""

from __future__ import annotations

import os
from typing import Optional

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from .brand import C_GREY, C_RED, C_FTGREY, C_LTGREY, C_WHITE, FONT_TEXT, FONT_DISPLAY
from .lxml_helpers import _get_or_add


def textbox(slide, text: str, left: float, top: float, width: float, height: float,
            fsize: float = 9, bold: bool = False, color: Optional[RGBColor] = None,
            align=PP_ALIGN.LEFT, italic: bool = False, wrap: bool = True,
            font: str = FONT_TEXT):
    """Add a text box. All positions/sizes in inches."""
    if color is None:
        color = C_GREY
    shape = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height))
    tf = shape.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size    = Pt(fsize)
    run.font.bold    = bold
    run.font.italic  = italic
    run.font.color.rgb = color
    run.font.name    = font
    return shape


def solidrect(slide, left: float, top: float, width: float, height: float,
              fill: RGBColor, line: Optional[RGBColor] = None):
    """Add a filled rectangle. All positions/sizes in inches.
    line=None removes border; line=RGBColor draws a border."""
    shape = slide.shapes.add_shape(
        1,   # MSO_SHAPE_TYPE.RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
    return shape


def horiz_line(slide, left: float, top: float, width: float,
               color: Optional[RGBColor] = None, width_pt: float = 1.0):
    """Add a horizontal line. Positions in inches, width_pt in points."""
    if color is None:
        color = C_RED
    shape = slide.shapes.add_connector(
        1,   # MSO_CONNECTOR.STRAIGHT
        Inches(left), Inches(top), Inches(left + width), Inches(top))
    shape.line.color.rgb = color
    shape.line.width = Pt(width_pt)
    return shape


def dashed_separator(slide, left: float, top: float, length: float,
                      color: Optional[RGBColor] = None, width_pt: float = 0.75,
                      dash: str = "dash", vertical: bool = False):
    """Add a dashed line separator (horizontal or vertical).

    Args:
        length: line length in inches (width if horizontal, height if vertical)
        vertical: if True, draw a vertical line instead of horizontal
    """
    if color is None:
        color = C_LTGREY
    if vertical:
        end_left, end_top = left, top + length
    else:
        end_left, end_top = left + length, top
    shape = slide.shapes.add_connector(
        1,  # MSO_CONNECTOR.STRAIGHT
        Inches(left), Inches(top),
        Inches(end_left), Inches(end_top))
    shape.line.color.rgb = color
    shape.line.width = Pt(width_pt)
    spPr = shape._element.spPr
    ln   = _get_or_add(spPr, "a:ln")
    _get_or_add(ln, "a:prstDash").set("val", dash)
    return shape


def stat_callout(slide, value, delta: Optional[float], label: str,
                 left: float, top: float) -> None:
    """Add a large single-stat display: oversized number, delta, and label."""
    textbox(slide, str(value),
            left, top, 2.0, 0.80,
            fsize=40, bold=True, color=C_RED,
            align=PP_ALIGN.CENTER, font=FONT_DISPLAY)
    if delta is not None:
        delta_str = (f"({delta:+d})" if isinstance(delta, int)
                     else f"({delta:+.1f})")
        textbox(slide, delta_str,
                left, top + 0.75, 2.0, 0.35,
                fsize=14, color=C_FTGREY,
                align=PP_ALIGN.CENTER, font=FONT_TEXT)
    textbox(slide, label,
            left, top + 1.10, 2.0, 0.40,
            fsize=9, color=C_FTGREY,
            align=PP_ALIGN.CENTER, font=FONT_TEXT)


def callout_box(slide, left: float, top: float, width: float, height: float,
                text: Optional[str] = None, border_color: Optional[RGBColor] = None,
                dashed: bool = True, fill_color: Optional[RGBColor] = None,
                fsize: float = 8, text_color: Optional[RGBColor] = None):
    """Add a rounded-rectangle annotation callout box."""
    if border_color is None:
        border_color = C_RED
    if fill_color is None:
        bc = str(border_color)  # "RRGGBB"
        r = int(int(bc[0:2], 16) * 0.12 + 0xFF * 0.88)
        g = int(int(bc[2:4], 16) * 0.12 + 0xFF * 0.88)
        b = int(int(bc[4:6], 16) * 0.12 + 0xFF * 0.88)
        fill_color = RGBColor(r, g, b)
    if text_color is None:
        text_color = C_GREY

    # 5 = MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE
    shape = slide.shapes.add_shape(
        5, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.color.rgb = border_color
    shape.line.width = Pt(1.0)

    if dashed:
        spPr = shape._element.spPr
        ln   = _get_or_add(spPr, "a:ln")
        _get_or_add(ln, "a:prstDash").set("val", "dash")

    if text:
        tf = shape.text_frame
        tf.word_wrap = True
        p   = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text             = text
        run.font.size        = Pt(fsize)
        run.font.color.rgb   = text_color
        run.font.name        = FONT_TEXT

    return shape
