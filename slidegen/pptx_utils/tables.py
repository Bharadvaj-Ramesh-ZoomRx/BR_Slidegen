"""
tables.py — Table builders for delta columns and value columns.

Reusable table patterns used alongside charts in slide renderers.
"""

from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt, Emu

from .brand import (
    C_WHITE, C_GREY, C_FTGREY, C_GREEN, C_RED,
    C_LBGREY, C_HDRGREY, FONT_TEXT, EMU_PER_IN,
)

# ── Header row height ────────────────────────────────────────────────────────
HEADER_ROW_HEIGHT_IN = 0.28


# ── Shared helpers ───────────────────────────────────────────────────────────

def _render_header(tbl, header_text: str, font_name: str | None = None):
    """Style row 0 as a dark header with centered white bold text."""
    tbl.rows[0].height = Inches(HEADER_ROW_HEIGHT_IN)
    hc = tbl.cell(0, 0)
    hc.fill.solid()
    hc.fill.fore_color.rgb = C_HDRGREY
    p = hc.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = header_text
    run.font.size = Pt(7)
    run.font.bold = True
    run.font.color.rgb = C_WHITE
    run.font.name = font_name or FONT_TEXT


def _render_data_row(tbl, row_idx: int, text: str, color: RGBColor,
                     row_height_in: float, font_name: str | None = None):
    """Style a single data row: alternating bg, centered colored bold text."""
    tbl.rows[row_idx].height = Inches(row_height_in)
    cell = tbl.cell(row_idx, 0)
    cell.fill.solid()
    cell.fill.fore_color.rgb = C_LBGREY if (row_idx - 1) % 2 == 0 else C_WHITE
    tf = cell.text_frame
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    run = tf.paragraphs[0].add_run()
    run.text = text
    run.font.size = Pt(8)
    run.font.bold = True
    run.font.color.rgb = color
    run.font.name = font_name or FONT_TEXT


def _delta_text_color(d: float | None) -> tuple[str, RGBColor]:
    """Return (display_text, color) for a delta value."""
    if d is None:
        return "N/A", C_FTGREY
    if d > 0:
        return f"+{d:.1f}", C_GREEN
    if d < 0:
        return f"{d:.1f}", C_RED
    return "0.0", C_GREY


# ── Public API ───────────────────────────────────────────────────────────────

def add_delta_col(slide, deltas: list[float | None],
                  left: float, top: float, width: float, height: float,
                  header: str, hdr_h_frac: float = 0.06):
    """Add a single-column delta table aligned with a chart.

    Args:
        deltas: list of float|None  (same order as chart categories, top-to-bottom)
        left/top/width/height: inches — match chart dimensions exactly
        header: column header string (e.g. 'MR Delta')
        hdr_h_frac: header row as fraction of total height (default 0.06)

    Returns: the table object
    """
    n = len(deltas)
    hdr_h = height * hdr_h_frac
    body_h = height - hdr_h
    row_h = body_h / n

    tbl = slide.shapes.add_table(
        n + 1, 1,
        Inches(left), Inches(top), Inches(width), Inches(height)
    ).table

    # Header row (custom height for proportional mode)
    tbl.rows[0].height = Emu(int(hdr_h * EMU_PER_IN))
    hc = tbl.cell(0, 0)
    hc.fill.solid()
    hc.fill.fore_color.rgb = C_HDRGREY
    hc.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    run = hc.text_frame.paragraphs[0].add_run()
    run.text = header
    run.font.size = Pt(7)
    run.font.bold = True
    run.font.color.rgb = C_WHITE
    run.font.name = FONT_TEXT

    # Data rows
    for i, d in enumerate(deltas):
        row_idx = i + 1
        tbl.rows[row_idx].height = Emu(int(row_h * EMU_PER_IN))
        cell = tbl.cell(row_idx, 0)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_LBGREY if i % 2 == 0 else C_WHITE

        if d is None:
            text, fcolor = "N/A", C_FTGREY
        elif d > 0:
            text, fcolor = f"+{d:.0f}", C_GREEN
        elif d < 0:
            text, fcolor = f"{d:.0f}", C_RED
        else:
            text, fcolor = "0", C_GREY

        tf = cell.text_frame
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        run = tf.paragraphs[0].add_run()
        run.text = text
        run.font.size = Pt(8)
        run.font.bold = True
        run.font.color.rgb = fcolor
        run.font.name = FONT_TEXT

    tbl.columns[0].width = Inches(width)
    return tbl


def add_delta_table(slide, deltas: list[float | None],
                    left: float, top: float, width: float, row_height: float,
                    header_text: str = "QoQ \u0394", font_name: str | None = None,
                    show_header: bool = True):
    """Add a single-column delta table with green/red conditional coloring.

    Args:
        deltas: list of float/None values
        row_height: height of each data row in inches
        show_header: if False, omit the header row (use when header is
            already in a chart_header_row above).
    Returns the table shape.
    """
    n = len(deltas)
    if show_header:
        total_h = HEADER_ROW_HEIGHT_IN + n * row_height
        n_rows = n + 1
    else:
        total_h = n * row_height
        n_rows = n

    tbl_shape = slide.shapes.add_table(
        n_rows, 1,
        Inches(left), Inches(top), Inches(width), Inches(total_h))
    tbl = tbl_shape.table

    if show_header:
        _render_header(tbl, header_text, font_name)
        data_offset = 1
    else:
        data_offset = 0

    for i, d in enumerate(deltas):
        text, color = _delta_text_color(d)
        _render_data_row(tbl, i + data_offset, text, color, row_height, font_name)

    tbl.columns[0].width = Inches(width)
    return tbl_shape


def add_value_table(slide, values: list[float | None],
                    left: float, top: float, width: float, row_height: float,
                    header_text: str = "Total %", value_color: RGBColor | None = None,
                    font_name: str | None = None):
    """Add a single-column table showing plain values (not delta-colored).

    Useful for total percentages in stacked bar charts.
    """
    n = len(values)
    total_h = HEADER_ROW_HEIGHT_IN + n * row_height

    tbl_shape = slide.shapes.add_table(
        n + 1, 1,
        Inches(left), Inches(top), Inches(width), Inches(total_h))
    tbl = tbl_shape.table

    _render_header(tbl, header_text, font_name)

    vc = value_color or C_GREY
    for i, v in enumerate(values):
        text = f"{v:.0f}%" if v is not None else "N/A"
        _render_data_row(tbl, i + 1, text, vc, row_height, font_name)

    tbl.columns[0].width = Inches(width)
    return tbl_shape
