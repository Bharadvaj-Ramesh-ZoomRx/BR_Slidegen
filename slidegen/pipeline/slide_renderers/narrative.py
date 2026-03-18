"""
narrative.py — Cover and executive summary slide renderers.
"""

from __future__ import annotations

from ._shared import (
    # helpers
    _resolve_template,
    # pptx_utils
    C_GREY, C_RED,
    textbox, solidrect, slide_header, slide_footer,
    cover_slide,
    # project_config types
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
    """Bullet-list executive summary. Insights are in ask.extra['insights']."""
    headline = _resolve_template(ask.headline, config)
    slide_header(slide, headline, font=config.font_display)

    color_current = config.primary.color_current
    insights = ask.extra.get("insights", [])

    y = 1.50
    for ins in insights:
        solidrect(slide, 0.40, y + 0.06, 0.08, 0.08, color_current)
        textbox(slide, ins, 0.60, y, 12.40, 0.50,
                fsize=9, color=C_GREY, font=config.font_body)
        y += 0.62

    source = _resolve_template(ask.source_text, config)
    if source:
        slide_footer(slide, source, font=config.font_body)
