"""
narrative.py — Cover and executive summary slide renderers.
"""

from __future__ import annotations

from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from lxml import etree
from pptx.oxml.ns import qn

from ._shared import (
    _resolve_template,
    C_GREY, C_FTGREY, C_RED, C_WHITE, C_LBGREY,
    textbox, solidrect, slide_header, slide_footer,
    cover_slide, _get_or_add,
    ProjectConfig, AskConfig,
)


def render_cover(slide, config: ProjectConfig, ask: AskConfig, data: dict, *, namer=None):
    """Title/cover slide."""
    headline = _resolve_template(ask.headline, config)
    extra = ask.extra
    cover_slide(
        slide,
        headline,
        extra.get("subtitle", ""),
        extra.get("date", ""),
        _resolve_template(extra.get("client", ""), config),
    )


def render_executive_summary(slide, config: ProjectConfig, ask: AskConfig, data: dict, *, namer=None):
    """Executive summary with numbered insights in a full-width branded card.

    Layout:
      - Headline at top
      - Full-width rounded card with brand accent (left border stripe)
      - Numbered insights with alternating subtle background bands
      - Each insight has a brand-colored number and black text

    Config (ask.extra):
        insights: list of finding strings (supports {{template}} placeholders)
    """
    headline = _resolve_template(ask.headline, config)
    slide_header(slide, headline, module_label="", font=config.font_display)

    extra = ask.extra or {}
    font = config.font_body
    brand = config.primary.color_current
    insights = extra.get("insights", [])

    if not insights:
        textbox(slide, "No insights provided", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    n = len(insights)

    # Card dimensions
    card_l = 0.30
    card_w = 12.73
    card_top = 1.15
    card_bottom = 6.70
    card_h = card_bottom - card_top
    accent_w = 0.06  # thin left accent stripe

    # Brand accent stripe (left edge of card)
    solidrect(slide, card_l, card_top, accent_w, card_h, brand)

    # Row layout inside card
    row_h = card_h / n
    num_w = 0.45       # number column width
    text_l = card_l + accent_w + num_w + 0.10
    text_w = card_w - accent_w - num_w - 0.25

    brand_hex = f'{brand[0]:02X}{brand[1]:02X}{brand[2]:02X}'

    for i, insight in enumerate(insights):
        insight = _resolve_template(insight, config)
        y = card_top + i * row_h

        # Alternating row background
        bg = C_LBGREY if i % 2 == 0 else C_WHITE
        solidrect(slide, card_l + accent_w, y, card_w - accent_w, row_h, bg)

        # Number badge (brand-colored, bold, centered in its column)
        num_text = f"{i + 1}"
        _add_number_badge(slide, card_l + accent_w + 0.08, y + row_h / 2 - 0.14,
                          0.28, 0.28, num_text, brand, font)

        # Insight text (vertically centered in the row)
        tb = textbox(slide, insight, text_l, y, text_w, row_h,
                     fsize=10, color=RGBColor(0x1A, 0x1A, 0x1A), font=font)
        if tb and hasattr(tb, 'text_frame'):
            tb.text_frame.word_wrap = True
            bodyPr = tb.text_frame._txBody.find(qn('a:bodyPr'))
            if bodyPr is not None:
                bodyPr.set('anchor', 'ctr')

    # Thin bottom line to close the card
    solidrect(slide, card_l, card_bottom, card_w, 0.01, brand)

    source = _resolve_template(ask.source_text, config)
    if source:
        slide_footer(slide, source, font=config.font_body)


def _add_number_badge(slide, left, top, width, height, text, color, font):
    """Add a filled circle with a number inside (insight badge)."""
    # Oval shape (MSO_AUTO_SHAPE_TYPE.OVAL = 9)
    shape = slide.shapes.add_shape(
        9, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()

    tf = shape.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = C_WHITE
    run.font.name = font

    # Vertical center
    bodyPr = tf._txBody.find(qn('a:bodyPr'))
    if bodyPr is not None:
        bodyPr.set('anchor', 'ctr')
        bodyPr.set('lIns', '0')
        bodyPr.set('rIns', '0')
        bodyPr.set('tIns', '0')
        bodyPr.set('bIns', '0')
