"""
slidegen/pipeline/slide_renderers — Slide type renderer package.

Each renderer takes (slide, config, ask, data, *, namer=None) and builds
slide content generically from the ask definition + extracted data.

Slide types:
  - cover
  - executive_summary
  - single_bar_with_delta
  - dual_bar_with_delta
  - dual_bar_qoq
  - clustered_compare
  - dual_bar_compare       (generic; "dual_brand_compare" is a backward-compat alias)
  - qoq_bar_with_delta
  - two_section_bar
  - stacked_order
  - lollipop
  - abacus
"""

from .narrative import render_cover, render_executive_summary
from .bar import (
    render_single_bar_with_delta,
    render_qoq_bar_with_delta,
    render_two_section_bar,
)
from .bar_dual import render_dual_bar_with_delta, render_dual_bar_qoq
from .compare import render_clustered_compare, render_stacked_order, render_dual_bar_compare
from .dot import render_lollipop, render_abacus, render_message_mbd

# REGISTRY — maps slide_type string → renderer function
RENDERERS = {
    "cover": render_cover,
    "executive_summary": render_executive_summary,
    "single_bar_with_delta": render_single_bar_with_delta,
    "dual_bar_with_delta": render_dual_bar_with_delta,
    "dual_bar_qoq": render_dual_bar_qoq,
    "clustered_compare": render_clustered_compare,
    "dual_bar_compare": render_dual_bar_compare,
    "dual_brand_compare": render_dual_bar_compare,   # backward-compat alias
    "qoq_bar_with_delta": render_qoq_bar_with_delta,
    "two_section_bar": render_two_section_bar,
    "stacked_order": render_stacked_order,
    "lollipop": render_lollipop,
    "abacus": render_abacus,
    "message_mbd": render_message_mbd,
}

__all__ = [
    "RENDERERS",
    "render_cover",
    "render_executive_summary",
    "render_single_bar_with_delta",
    "render_dual_bar_with_delta",
    "render_dual_bar_qoq",
    "render_clustered_compare",
    "render_dual_bar_compare",
    "render_qoq_bar_with_delta",
    "render_two_section_bar",
    "render_stacked_order",
    "render_lollipop",
    "render_abacus",
    "render_message_mbd",
]
