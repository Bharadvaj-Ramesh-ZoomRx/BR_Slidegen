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
from .compare import render_clustered_compare, render_stacked_order, render_dual_bar_compare, render_hii_scorecard, render_dual_doughnut
from .dot import render_abacus, render_dual_abacus, render_followup_rep, render_message_mbd
from .line import render_trended_scorecard, render_trended_activity
from .quadrant import render_quadrant_scatter
from .heatmap import render_heatmap_table

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
    "abacus": render_abacus,
    "message_mbd": render_message_mbd,
    "dual_abacus": render_dual_abacus,
    "followup_rep": render_followup_rep,
    "hii_scorecard": render_hii_scorecard,
    "dual_doughnut": render_dual_doughnut,
    "trended_scorecard": render_trended_scorecard,
    "trended_activity": render_trended_activity,
    "quadrant_scatter": render_quadrant_scatter,
    "heatmap_table": render_heatmap_table,
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
    "render_abacus",
    "render_message_mbd",
    "render_dual_abacus",
    "render_hii_scorecard",
    "render_dual_doughnut",
    "render_trended_scorecard",
    "render_trended_activity",
    "render_quadrant_scatter",
    "render_heatmap_table",
]
