"""
tables.py — Table builders for delta columns and value columns.

Reusable table patterns used alongside charts in slide renderers.
"""

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt, Emu

from .brand import (
    C_WHITE, C_GREY, C_FTGREY, C_GREEN, C_RED,
    C_LBGREY, C_HDRGREY, FONT_TEXT, EMU_PER_IN,
)


def add_delta_col(slide, deltas, left, top, width, height, header,
                  hdr_h_frac=0.06):
    """Add a single-column delta table aligned with a chart.

    Args:
        deltas: list of float|None  (same order as chart categories, top-to-bottom)
        left/top/width/height: inches — match chart dimensions exactly
        header: column header string (e.g. 'MR Delta')
        hdr_h_frac: header row as fraction of total height (default 0.06)

    Returns: the table object
    """
    n      = len(deltas)
    hdr_h  = height * hdr_h_frac
    body_h = height - hdr_h
    row_h  = body_h / n

    tbl = slide.shapes.add_table(
        n + 1, 1,
        Inches(left), Inches(top), Inches(width), Inches(height)
    ).table

    # Header row
    tbl.rows[0].height = Emu(int(hdr_h * EMU_PER_IN))
    hc = tbl.cell(0, 0)
    hc.fill.solid()
    hc.fill.fore_color.rgb = C_HDRGREY
    hc.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    run = hc.text_frame.paragraphs[0].add_run()
    run.text = header
    run.font.size  = Pt(7)
    run.font.bold  = True
    run.font.color.rgb = C_WHITE
    run.font.name  = FONT_TEXT

    # Data rows
    for i, d in enumerate(deltas):
        tbl.rows[i + 1].height = Emu(int(row_h * EMU_PER_IN))
        cell = tbl.cell(i + 1, 0)

        # Alternating row background
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_LBGREY if i % 2 == 0 else C_WHITE

        # Value and colour
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
        run.font.size  = Pt(8)
        run.font.bold  = True
        run.font.color.rgb = fcolor
        run.font.name  = FONT_TEXT

    tbl.columns[0].width = Inches(width)
    return tbl


def add_delta_table(slide, deltas, left, top, width, row_height,
                    header_text="QoQ \u0394", font_name=None):
    """Add a single-column delta table with green/red conditional coloring.

    Args:
        deltas: list of float/None values
        row_height: height of each data row in inches
    Returns the table shape.
    """
    n = len(deltas)
    total_h = 0.28 + n * row_height

    tbl_shape = slide.shapes.add_table(
        n + 1, 1,
        Inches(left), Inches(top), Inches(width), Inches(total_h))
    tbl = tbl_shape.table

    # Header
    tbl.rows[0].height = Inches(0.28)
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

    # Data rows
    for i, d in enumerate(deltas):
        tbl.rows[i + 1].height = Inches(row_height)
        cell = tbl.cell(i + 1, 0)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_LBGREY if i % 2 == 0 else C_WHITE

        if d is None:
            text, fcolor = "N/A", C_FTGREY
        elif d > 0:
            text, fcolor = f"+{d:.1f}", C_GREEN
        elif d < 0:
            text, fcolor = f"{d:.1f}", RGBColor(0xFF, 0x00, 0x00)
        else:
            text, fcolor = "0.0", C_GREY

        tf = cell.text_frame
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        run = tf.paragraphs[0].add_run()
        run.text = text
        run.font.size = Pt(8)
        run.font.bold = True
        run.font.color.rgb = fcolor
        run.font.name = font_name or FONT_TEXT

    tbl.columns[0].width = Inches(width)
    return tbl_shape


def add_value_table(slide, values, left, top, width, row_height,
                    header_text="Total %", value_color=None, font_name=None):
    """Add a single-column table showing plain values (not delta-colored).

    Useful for total percentages in stacked bar charts.
    """
    n = len(values)
    total_h = 0.28 + n * row_height

    tbl_shape = slide.shapes.add_table(
        n + 1, 1,
        Inches(left), Inches(top), Inches(width), Inches(total_h))
    tbl = tbl_shape.table

    tbl.rows[0].height = Inches(0.28)
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

    for i, v in enumerate(values):
        tbl.rows[i + 1].height = Inches(row_height)
        cell = tbl.cell(i + 1, 0)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_LBGREY if i % 2 == 0 else C_WHITE
        tf = cell.text_frame
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        run = tf.paragraphs[0].add_run()
        run.text = f"{v:.0f}%" if v is not None else "N/A"
        run.font.size = Pt(8)
        run.font.bold = True
        run.font.color.rgb = value_color or C_GREY
        run.font.name = font_name or FONT_TEXT

    tbl.columns[0].width = Inches(width)
    return tbl_shape
