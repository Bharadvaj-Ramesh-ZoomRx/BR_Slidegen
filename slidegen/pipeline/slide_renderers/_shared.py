"""
_shared.py — Shared imports, layout constants, and private helpers for all
slide renderer submodules.
"""

from __future__ import annotations
import logging

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from lxml import etree
from pptx.oxml.ns import qn

from slidegen.pptx_utils import (
    C_GREEN, C_WHITE, C_GREY, C_FTGREY, C_LBGREY, C_HDRGREY, C_RED,
    FONT_TEXT,
    textbox, solidrect, horiz_line, slide_header, slide_footer,
    section_header_bar, section_breadcrumb, chart_header_row,
    callout_box, dashed_separator, hide_axis, set_series_color,
    set_plot_area_gap, set_overlap, set_series_no_border,
    invert_cat_axis, hide_cat_labels, set_data_label_color,
    set_chart_plot_area, set_val_axis_scale,
    set_series_marker, set_series_line_style,
    _get_or_add, suppress_para_bullets, suppress_cat_axis_bullets, cell_vcenter,
    enable_data_labels, delete_data_label,
    add_single_bar_chart, add_clustered_bar_chart,
    add_line_chart, add_stacked_column_chart, add_scatter_chart,
    add_delta_table, add_value_table,
    cover_slide,
)
from slidegen.pipeline.data_loaders import delta
from slidegen.pptx_utils.tables import HEADER_ROW_HEIGHT_IN
from slidegen.pipeline.project_config import ProjectConfig, AskConfig, parse_color

logger = logging.getLogger(__name__)


# ── Layout constants ─────────────────────────────────────────────────────────
# Y-axis positions (inches)
SECTION_BAR_TOP = 1.40
CHART_TOP_STD = 1.85       # standard chart top below section bar
LEGEND_GAP = 0.15           # gap below chart before legend
FOOTER_TOP = 6.93           # footer Y position (spec)

# Slide dimensions
SLIDE_W = 13.33             # slide width (inches)

# Chart dimensions (inches) — generic
SINGLE_BAR_WIDTH = 9.0
CLUSTERED_BAR_WIDTH = 8.5
DELTA_COL_WIDTH = 0.65
DELTA_COL_NARROW = 0.55
CHART_DELTA_GAP = 0.10      # gap between chart and delta column

# Dual bar (Archetype 3 — MR+ME) positions from spec
DUAL_HEADER_ROW_TOP = 1.40   # red header row Y (just below section bar)
DUAL_HEADER_ROW_H = 0.42     # header row height (taller for 2-line text + padding)
DUAL_CHART_TOP = 1.84        # chart top (= DUAL_HEADER_ROW_TOP + DUAL_HEADER_ROW_H + 0.02)
DUAL_MR_LEFT = 0.20          # left chart X
DUAL_MR_WIDTH = 7.30         # left chart width (wider — has cat labels)
DUAL_MR_DELTA_LEFT = 7.54    # left delta column X
DUAL_ME_LEFT = 8.26          # right chart X
DUAL_ME_WIDTH = 4.40         # right chart width (narrower — no cat labels)
DUAL_ME_DELTA_LEFT = 12.70   # right delta column X
DUAL_DELTA_WIDTH = 0.62      # delta column width
DUAL_SEPARATOR_X = 8.16      # vertical dashed separator X
DUAL_MAX_CHART_HEIGHT = 4.90  # max chart height (fills to footer area)

# Narrow variant — template-matched positions from template.pptx slide 15
# Category table + main table + overlaid charts + delta tables + callouts
DUAL_T_CAT_LEFT = 1.269        # Category (Tag) table left
DUAL_T_CAT_W = 2.200           # Category table width
DUAL_T_MAIN_LEFT = 3.472       # Main table left (Recall | Effectiveness cols)
DUAL_T_MR_COL_W = 3.320        # Recall column width in main table
DUAL_T_ME_COL_W = 3.320        # Effectiveness column width in main table
DUAL_T_HDR_H = 0.501           # Table header row height
DUAL_T_ROW_H = 0.438           # Table data row height
DUAL_T_MR_CHART_L = 3.457      # Recall chart left (overlaid, no axis labels)
DUAL_T_MR_CHART_W = 2.974      # Recall chart width
DUAL_T_MR_DELTA_L = 6.017      # Recall delta table left
DUAL_T_MR_DELTA_W = 0.656      # Recall delta table width
DUAL_T_ME_CHART_L = 6.742      # Effectiveness chart left
DUAL_T_ME_CHART_W = 2.784      # Effectiveness chart width
DUAL_T_ME_DELTA_L = 9.526      # Effectiveness delta table left
DUAL_T_ME_DELTA_W = 0.529      # Effectiveness delta table width
DUAL_T_CALLOUT_L = 10.326      # Callout boxes left edge
DUAL_T_CALLOUT_W = 2.677       # Callout boxes width

# Chart sizing constraints
MAX_CHART_HEIGHT = 4.5
MAX_CLUSTERED_HEIGHT = 4.8
MAX_QOQ_HEIGHT = 4.0
MIN_CHART_HEIGHT = 3.0       # minimum chart height to avoid excessive white space
ROW_SCALE_FACTOR = 0.85     # delta table row height as fraction of chart row

# Compat aliases for renderers not yet updated to Archetype 3 layout
CHART_TOP_DUAL = 1.80                # used by render_dual_bar_qoq
DUAL_BAR_WIDTH = 5.0                 # used by render_dual_bar_qoq
MAX_DUAL_CHART_HEIGHT = 4.2          # used by render_dual_bar_qoq

# Footer and content boundary
FOOTER_TOP = 6.78           # footer textbox top (from slide_footer y=6.78)
MAX_CHART_BOTTOM = 6.35     # latest a chart body can end (leaves room for legend + gap)
LEGEND_H = 0.18             # height of a manual legend row (swatch + label)

# Bar chart gap/overlap
BAR_GAP_STD = 80
CLUSTERED_OVERLAP = -10
CLUSTERED_GAP = 65

# Label truncation lengths
LABEL_MAX_SINGLE = 40
LABEL_MAX_DUAL = 35
LABEL_MAX_CLUSTERED = 45
LABEL_MAX_STACKED = 40

# Small segment threshold for stacked bars (%)
STACKED_HIDE_THRESHOLD = 3

# Shared dot-chart layout — lollipop and abacus use the same label panel + plot area
_DOT_LABEL_W   = 2.5          # category label column width (inches)
_DOT_PLOT_LEFT = 2.80         # left edge of plot area (0.30 margin + 2.5 label col)
_DOT_PLOT_W    = 6.20         # plot area width (leaves room for value labels)

# Lollipop chart layout constants
LOLLI_LABEL_W   = _DOT_LABEL_W
LOLLI_PLOT_LEFT = _DOT_PLOT_LEFT
LOLLI_PLOT_W    = _DOT_PLOT_W
LOLLI_DOT_R     = 0.09        # current-period circle radius (inches)
LOLLI_PRIOR_R   = 0.065       # prior-period circle radius (inches)

# Abacus chart layout constants (legacy shape-based — kept for lollipop shared dims)
ABACUS_LABEL_W   = _DOT_LABEL_W
ABACUS_PLOT_LEFT = _DOT_PLOT_LEFT
ABACUS_PLOT_W    = _DOT_PLOT_W
ABACUS_TRACK_H   = 0.065      # height of the grey track rectangle
ABACUS_DOT_R     = 0.10       # current-period bead radius
ABACUS_PRIOR_R   = 0.07       # prior-period bead radius

# Abacus XY scatter layout — template slide 20 (2-period, horizontally centered)
_ABS_HDR_H      = 0.28    # header row height
_ABS_ROW_H_MIN  = 0.25    # minimum per-row height
_ABS_ROW_H_MAX  = 0.38    # maximum per-row height
_ABS_TOP        = 1.88    # top of the entire block
_ABS_LABEL_W    = 3.80    # attribute label column width
_ABS_VAL_W      = 0.58    # each value column (prior or current)
_ABS_CHART_W    = 3.10    # XY scatter chart width
_ABS_DELTA_W    = 0.58    # delta column width
_ABS_GAP        = 0.08    # gap between element groups
_ABS_SLIDE_W    = 13.33   # slide width for centering

# Shape type constant (MSO_AUTO_SHAPE_TYPE.OVAL = 9)
_OVAL = 9

# Dual-brand compare layout — template-matched from template.pptx slide 23
# Two separate brand bar charts (left/right) with a category column and separator
DUAL_BC_TOP         = 2.30    # chart/table top (brand labels sit above at ~1.86)
DUAL_BC_ROW_H       = 0.930   # row height (template: 4 rows × 0.930" = 3.72")
DUAL_BC_CAT_LEFT    = 1.917   # category table left
DUAL_BC_CAT_W       = 2.710   # category table width
DUAL_BC_L_CHART_L   = 5.041   # left brand bar chart left
DUAL_BC_L_CHART_W   = 2.480   # left brand bar chart width
DUAL_BC_L_DELTA_L   = 6.976   # left brand delta table left
DUAL_BC_L_DELTA_W   = 0.529   # left brand delta table width
DUAL_BC_SEP_X       = 7.756   # vertical separator X
DUAL_BC_R_CHART_L   = 8.536   # right brand bar chart left
DUAL_BC_R_CHART_W   = 2.480   # right brand bar chart width
DUAL_BC_R_DELTA_L   = 10.651  # right brand delta table left
DUAL_BC_R_DELTA_W   = 0.529   # right brand delta table width


# ── Dynamic label width ──────────────────────────────────────────────────────

def _auto_label_width(labels: list[str], fsize: float = 7.5,
                      min_w: float = 2.50, max_w: float = 5.50,
                      max_line_chars: int = 45) -> float:
    """Calculate optimal label table width based on longest label text.

    At ~7.5pt Calibri, ~13 chars per inch. Labels longer than max_line_chars
    are assumed to wrap to 2 lines — width is based on half the length.
    """
    if not labels:
        return min_w
    max_len = max(len(l) for l in labels)
    cpi = 13.0 * (7.5 / max(fsize, 5.0))
    # If text would wrap to 2 lines, use half the length for width calc
    line_len = (max_len + 1) // 2 if max_len > max_line_chars else max_len
    width = line_len / cpi + 0.25  # add cell margin padding
    return max(min_w, min(max_w, round(width, 2)))


# ── Template resolution ───────────────────────────────────────────────────────

def _resolve_template(text: str, config: ProjectConfig) -> str:
    """Replace {{...}} placeholders in headline/section text."""
    if not text:
        return text
    text = text.replace("{{primary.name}}", config.primary.name)
    text = text.replace("{{primary.full_name}}", config.primary.full_name)
    text = text.replace("{{competitor.name}}", config.competitor.name)
    text = text.replace("{{competitor.full_name}}", config.competitor.full_name)
    text = text.replace("{{period_current}}", config.period_current)
    text = text.replace("{{period_prior}}", config.period_prior)
    text = text.replace("{{client}}", config.client)
    return text


# ── Shared layout helpers ─────────────────────────────────────────────────────

def _slide_chrome(slide, config: ProjectConfig, ask: AskConfig):
    """Add header, section bar, breadcrumb, and footer to a slide."""
    headline = _resolve_template(ask.headline, config)
    slide_header(slide, headline, font=config.font_display)

    section = _resolve_template(ask.section, config)
    if section:
        section_header_bar(slide, section, top=SECTION_BAR_TOP,
                           icon_path=config.section_icon_path or None,
                           font=config.font_body)

    # Breadcrumb (top-right, e.g. "Appendix" or "Key Findings")
    breadcrumb = ask.extra.get("breadcrumb", "") if ask.extra else ""
    if breadcrumb:
        section_breadcrumb(slide, _resolve_template(breadcrumb, config))

    source = _resolve_template(ask.source_text, config)
    if source:
        slide_footer(slide, source, font=config.font_body)


def _get_brand_colors(config: ProjectConfig, ask: AskConfig):
    """Return (color_current, color_prior) for the ask's brand."""
    brand_key = ask.brand if ask.brand else "primary"
    brand = config.brands.get(brand_key, config.primary)
    return brand.color_current, brand.color_prior


def _sort_data(rows: list[dict], sort_by: str | None, sort_desc: bool = True) -> list[dict]:
    """Sort rows by a field, handling None values."""
    if not sort_by:
        return rows
    return sorted(rows, key=lambda x: x.get(sort_by) or 0, reverse=sort_desc)


def _cap_chart_h(chart_h: float, chart_top: float, has_axis_label: bool = False) -> float:
    """Clamp chart height so the manual legend + gap fits above the footer."""
    budget = MAX_CHART_BOTTOM - chart_top
    if has_axis_label:
        budget -= 0.24          # axis sub-label height
    return max(MIN_CHART_HEIGHT, min(chart_h, budget))


def _make_legend(slide, items: list[tuple], left: float, top: float, font_name: str = None,
                 center_over: tuple[float, float] | None = None):
    """Add a manual legend row. items: [(color, label), ...]

    If center_over=(block_left, block_width), the legend row is horizontally
    centered over that block and the ``left`` parameter is ignored.
    """
    item_w = 2.0   # horizontal advance per legend item
    if center_over is not None:
        block_left, block_width = center_over
        total_legend_w = len(items) * item_w
        left = block_left + (block_width - total_legend_w) / 2
    x = left
    for color, label in items:
        solidrect(slide, x, top, 0.18, 0.12, color)
        textbox(slide, label, x + 0.25, top - 0.02, 1.8, 0.18,
                fsize=7, color=C_GREY, font=font_name)
        x += item_w


def _add_dot(slide, cx: float, cy: float, r: float, color):
    """Draw a filled borderless circle centered at (cx, cy) with radius r (inches)."""
    dot = slide.shapes.add_shape(
        _OVAL,
        Inches(cx - r), Inches(cy - r),
        Inches(r * 2), Inches(r * 2))
    dot.fill.solid()
    dot.fill.fore_color.rgb = color
    dot.line.fill.background()


def _pptx_table(slide, col_widths_in, row_heights_in, left, top):
    """Create a pptx table with exact column/row dimensions. Returns (shape, table)."""
    n_rows = len(row_heights_in)
    n_cols = len(col_widths_in)
    shape = slide.shapes.add_table(
        n_rows, n_cols,
        Inches(left), Inches(top),
        Inches(sum(col_widths_in)), Inches(sum(row_heights_in))
    )
    tbl = shape.table
    for ci, cw in enumerate(col_widths_in):
        tbl.columns[ci].width = Inches(cw)
    for ri, rh in enumerate(row_heights_in):
        tbl.rows[ri].height = Inches(rh)
    return shape, tbl


def _style_tbl_cell(cell, text, bg=None, fg=None, fsize=8.0, bold=False,
                    align=PP_ALIGN.CENTER, font=None, ml=0.05, mr=0.05):
    """Style a pptx table cell: fill, margins, vertical center, and multi-line text."""
    if bg is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg
    else:
        cell.fill.background()
    cell.margin_left = Inches(ml)
    cell.margin_right = Inches(mr)
    cell.margin_top = Inches(0.02)
    cell.margin_bottom = Inches(0.02)
    cell_vcenter(cell)
    tf = cell.text_frame
    tf.word_wrap = True
    lines = text.replace('\x0b', '\n').split('\n') if text else ['']
    for li, line in enumerate(lines):
        p = tf.paragraphs[li] if li < len(tf.paragraphs) else tf.add_paragraph()
        p.alignment = align
        suppress_para_bullets(p._p)
        run = p.runs[0] if p.runs else p.add_run()
        run.text = line
        run.font.size = Pt(fsize)
        run.font.bold = bold
        if fg:
            run.font.color.rgb = fg
        if font:
            run.font.name = font


def _cell_bottom_border(cell, hex_rgb, w_emu=6350):
    """Add a bottom border to a pptx table cell via XML."""
    tc = cell._tc
    tcPr = tc.find(qn('a:tcPr'))
    if tcPr is None:
        tcPr = etree.SubElement(tc, qn('a:tcPr'))
    for old in tcPr.findall(qn('a:lnB')):
        tcPr.remove(old)
    lnB = etree.SubElement(tcPr, qn('a:lnB'))
    lnB.set('w', str(w_emu))
    sf = etree.SubElement(lnB, qn('a:solidFill'))
    clr = etree.SubElement(sf, qn('a:srgbClr'))
    clr.set('val', hex_rgb)
