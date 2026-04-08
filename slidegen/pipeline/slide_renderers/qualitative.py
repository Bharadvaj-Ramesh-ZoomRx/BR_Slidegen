"""
qualitative.py — Qualitative theme analysis slide renderer.

Renders a dedicated qualitative slide with:
  Left panel:  Theme frequency horizontal bars (sorted descending)
  Right panel: Representative quote boxes with respondent segment tags

Data source: qualitative_data.json themes from validated_analysis.md
  ask.extra.themes: [{"theme": str, "pct": float, "count": int}, ...]
  ask.extra.quotes: [{"text": str, "attribution": str}, ...]
  ask.extra.qual_source: str (e.g., "Q1.53A (n=68)")
"""

from __future__ import annotations

from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

from ._shared import (
    CHART_TOP_STD, SLIDE_W, FOOTER_TOP,
    _resolve_template, _slide_chrome,
    FONT_HDR, FONT_BODY,
    C_GREEN, C_GREY, C_LBGREY, C_HDRGREY, C_WHITE, C_RED,
    textbox, solidrect, callout_box,
    ProjectConfig, AskConfig,
    logger,
)


# ── Layout constants ─────────────────────────────────────────────────────

# Two-panel split: left 55% (theme bars), right 45% (quotes)
_MARGIN = 0.30
_PANEL_GAP = 0.25
_PANEL_TOP = 1.85
_PANEL_BOTTOM = 6.50

_LEFT_PANEL_L = _MARGIN
_LEFT_PANEL_W = (SLIDE_W - 2 * _MARGIN - _PANEL_GAP) * 0.55
_RIGHT_PANEL_L = _LEFT_PANEL_L + _LEFT_PANEL_W + _PANEL_GAP
_RIGHT_PANEL_W = SLIDE_W - _RIGHT_PANEL_L - _MARGIN

# Theme bar dimensions
_BAR_LABEL_W = 2.2     # theme label width
_BAR_MAX_W = _LEFT_PANEL_W - _BAR_LABEL_W - 0.6  # max bar width
_BAR_H = 0.32          # bar height
_BAR_GAP = 0.12        # vertical gap between bars
_BAR_PCT_W = 0.5       # width for percentage label
_BAR_COLOR = RGBColor(0xF7, 0x58, 0x24)  # RYB orange (default)
_BAR_BG = RGBColor(0xF0, 0xF0, 0xF0)     # light grey background

# Quote box dimensions
_QUOTE_GAP = 0.18
_QUOTE_BORDER = RGBColor(0x50, 0x50, 0x50)
_QUOTE_FILL = RGBColor(0xF8, 0xF8, 0xF8)
_QUOTE_TEXT_COLOR = RGBColor(0x33, 0x33, 0x33)
_ATTR_COLOR = RGBColor(0x70, 0x70, 0x70)

# Source label
_SOURCE_COLOR = RGBColor(0x80, 0x80, 0x80)


def render_qual_theme_analysis(slide, config: ProjectConfig, ask: AskConfig,
                                data: dict, *, namer=None):
    """Qualitative theme analysis: theme frequency bars + representative quotes."""
    _slide_chrome(slide, config, ask)

    extra = ask.extra or {}
    themes = extra.get("themes", [])
    quotes = extra.get("quotes", [])
    qual_source = extra.get("qual_source", "")

    if not themes:
        # Fallback: try to use data dict if themes were passed there
        if isinstance(data, dict):
            themes = data.get("themes", [])
            quotes = data.get("quotes", quotes)
            qual_source = data.get("qual_source", qual_source)

    if not themes:
        textbox(slide, "No qualitative theme data available",
                2.0, 3.0, 9.0, 0.5, fsize=12, color=C_GREY)
        return

    # Resolve bar color from brand config
    bar_color = config.primary.color_current if config.primary else _BAR_COLOR

    # ── Source label ──────────────────────────────────────────────────
    if qual_source:
        source_text = f"Source: {qual_source}"
        textbox(slide, source_text,
                _LEFT_PANEL_L, _PANEL_TOP - 0.30, _LEFT_PANEL_W, 0.22,
                fsize=8, italic=True, color=_SOURCE_COLOR,
                font=config.font_body)

    # ── Left panel: Theme frequency bars ──────────────────────────────
    max_pct = max((t.get("pct", 0) for t in themes), default=100) or 100

    y = _PANEL_TOP
    for i, theme in enumerate(themes):
        theme_name = theme.get("theme", f"Theme {i+1}")
        pct = theme.get("pct", 0)
        count = theme.get("count", 0)

        # Theme label
        textbox(slide, theme_name,
                _LEFT_PANEL_L, y, _BAR_LABEL_W, _BAR_H,
                fsize=8.5, bold=False, color=C_GREY,
                align=PP_ALIGN.RIGHT, font=config.font_body)

        # Bar background
        bar_left = _LEFT_PANEL_L + _BAR_LABEL_W + 0.12
        solidrect(slide, bar_left, y + 0.02, _BAR_MAX_W, _BAR_H - 0.04,
                  fill=_BAR_BG)

        # Bar fill (proportional to max)
        bar_w = max(0.05, (pct / max_pct) * _BAR_MAX_W) if max_pct > 0 else 0.05
        solidrect(slide, bar_left, y + 0.02, bar_w, _BAR_H - 0.04,
                  fill=bar_color)

        # Percentage label (right of bar)
        pct_text = f"{pct:.0f}%"
        if count:
            pct_text += f" ({count})"
        textbox(slide, pct_text,
                bar_left + _BAR_MAX_W + 0.06, y, _BAR_PCT_W, _BAR_H,
                fsize=8, bold=True, color=C_GREY,
                align=PP_ALIGN.LEFT, font=config.font_body)

        y += _BAR_H + _BAR_GAP

        # Stop if we'd overflow
        if y > _PANEL_BOTTOM - _BAR_H:
            break

    # ── Right panel: Representative quotes ────────────────────────────
    if not quotes:
        return

    # Calculate available height per quote
    n_quotes = min(len(quotes), 4)  # max 4 quotes
    available_h = _PANEL_BOTTOM - _PANEL_TOP
    quote_h = min(1.2, (available_h - (n_quotes - 1) * _QUOTE_GAP) / n_quotes)

    y = _PANEL_TOP
    for i, quote in enumerate(quotes[:n_quotes]):
        q_text = quote.get("text", "")
        attribution = quote.get("attribution", "")

        # Quote box background
        box = slide.shapes.add_shape(
            5,  # MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE
            Inches(_RIGHT_PANEL_L), Inches(y),
            Inches(_RIGHT_PANEL_W), Inches(quote_h))
        box.fill.solid()
        box.fill.fore_color.rgb = _QUOTE_FILL
        box.line.color.rgb = _QUOTE_BORDER
        box.line.width = Pt(0.75)

        # Left accent bar (colored stripe)
        accent_w = 0.04
        solidrect(slide, _RIGHT_PANEL_L, y, accent_w, quote_h,
                  fill=bar_color)

        # Quote text (with opening quote mark)
        q_display = f"\u201c{q_text}\u201d" if q_text else ""
        text_margin = 0.14
        textbox(slide, q_display,
                _RIGHT_PANEL_L + text_margin, y + 0.08,
                _RIGHT_PANEL_W - text_margin - 0.10, quote_h - 0.38,
                fsize=8, italic=True, color=_QUOTE_TEXT_COLOR,
                font=config.font_body)

        # Attribution line
        if attribution:
            textbox(slide, f"\u2014 {attribution}",
                    _RIGHT_PANEL_L + text_margin, y + quote_h - 0.28,
                    _RIGHT_PANEL_W - text_margin - 0.10, 0.20,
                    fsize=7, italic=False, color=_ATTR_COLOR,
                    align=PP_ALIGN.RIGHT, font=config.font_body)

        y += quote_h + _QUOTE_GAP

        if y > _PANEL_BOTTOM:
            break

    # ── Namer ─────────────────────────────────────────────────────────
    if namer:
        namer.done()
