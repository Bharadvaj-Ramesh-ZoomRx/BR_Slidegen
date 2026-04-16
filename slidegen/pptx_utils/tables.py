"""
tables.py — Table builders for delta columns and value columns.

Reusable table patterns used alongside charts in slide renderers.
"""

from __future__ import annotations

from typing import Optional

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt, Emu

from .brand import (
    C_WHITE, C_GREY, C_FTGREY, C_GREEN, C_RED,
    C_LBGREY, C_HDRGREY, FONT_TEXT, EMU_PER_IN,
)
from .lxml_helpers import _get_or_add, suppress_para_bullets, cell_vcenter

# ── Header row height ────────────────────────────────────────────────────────
HEADER_ROW_HEIGHT_IN = 0.28


def _cell_vcenter(cell) -> None:
    """Vertically center-align text within a table cell."""
    cell_vcenter(cell)


# ── Shared helpers ───────────────────────────────────────────────────────────

def _render_header(tbl, header_text: str, font_name: Optional[str] = None,
                   display_font: Optional[str] = None, font_size_pt: float = 11.0):
    """Style row 0 as a dark header with centered white bold text."""
    tbl.rows[0].height = Inches(HEADER_ROW_HEIGHT_IN)
    hc = tbl.cell(0, 0)
    hc.fill.solid()
    hc.fill.fore_color.rgb = C_HDRGREY
    p = hc.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    suppress_para_bullets(p._p)
    run = p.add_run()
    run.text = header_text
    run.font.size = Pt(font_size_pt)
    run.font.bold = True
    run.font.color.rgb = C_WHITE
    run.font.name = display_font or font_name or FONT_TEXT
    _cell_vcenter(hc)


def _render_data_row(tbl, row_idx: int, text: str, color: RGBColor,
                     row_height_in: float, font_name: Optional[str] = None,
                     data_idx: Optional[int] = None, font_size_pt: float = 11.0):
    """Style a single data row: alternating bg, centered colored bold text."""
    tbl.rows[row_idx].height = Inches(row_height_in)
    cell = tbl.cell(row_idx, 0)
    cell.fill.solid()
    # Use data_idx for alternating if provided, otherwise infer from row_idx
    alt = data_idx if data_idx is not None else (row_idx - 1)
    cell.fill.fore_color.rgb = C_LBGREY if alt % 2 == 0 else C_WHITE
    tf = cell.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    suppress_para_bullets(p._p)
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size_pt)
    run.font.bold = True
    run.font.color.rgb = color
    run.font.name = font_name or FONT_TEXT
    _cell_vcenter(cell)


def _delta_text_color(d: float | None, *,
                      threshold: float = 5.0,
                      show_pct: bool = True) -> tuple[str, RGBColor]:
    """Return (display_text, color) for a delta value.

    Args:
        d: Delta value in percentage points (or None for N/A).
        threshold: Minimum absolute delta to apply green/red coloring.
            Deltas below this threshold render in neutral grey (no false
            signal). Set to 0 to color all non-zero deltas.
        show_pct: If True, append '%' suffix and show as integer.
    """
    if d is None:
        return "N/A", C_FTGREY
    suffix = "%" if show_pct else ""
    if show_pct:
        text = f"+{d:.0f}{suffix}" if d > 0 else (f"{d:.0f}{suffix}" if d < 0 else f"0{suffix}")
    else:
        text = f"+{d:.1f}" if d > 0 else (f"{d:.1f}" if d < 0 else "0.0")
    # Color only if magnitude meets threshold
    if abs(d) >= threshold:
        color = C_GREEN if d > 0 else C_RED
    elif d == 0:
        color = C_GREY
    else:
        color = C_GREY  # below threshold — neutral
    return text, color


# ── Public API ───────────────────────────────────────────────────────────────

def add_delta_col(slide, deltas: list[float | None],
                  left: float, top: float, width: float, height: float,
                  header: str, hdr_h_frac: float = 0.06,
                  font_size_pt: float = 11.0,
                  header_font_size_pt: Optional[float] = None,
                  font_name: Optional[str] = None):
    """Add a single-column delta table aligned with a chart.

    Args:
        deltas: list of float|None  (same order as chart categories, top-to-bottom)
        left/top/width/height: inches — match chart dimensions exactly
        header: column header string (e.g. 'MR Delta')
        hdr_h_frac: header row as fraction of total height (default 0.06)
        font_size_pt: font size for data cells (default 11pt, back-compat)
        header_font_size_pt: font size for header cell (defaults to font_size_pt)
        font_name: override font family for all cells

    Returns: the table object
    """
    n = len(deltas)
    hdr_h = height * hdr_h_frac
    body_h = height - hdr_h
    row_h = body_h / n
    hdr_size = header_font_size_pt if header_font_size_pt is not None else font_size_pt

    tbl = slide.shapes.add_table(
        n + 1, 1,
        Inches(left), Inches(top), Inches(width), Inches(height)
    ).table

    # Header row (custom height for proportional mode)
    tbl.rows[0].height = Emu(int(hdr_h * EMU_PER_IN))
    _render_header(tbl, header, font_name=font_name, font_size_pt=hdr_size)

    # Data rows (use integer-rounded delta format for proportional mode)
    for i, d in enumerate(deltas):
        if d is None:
            text, color = "N/A", C_FTGREY
        elif d > 0:
            text, color = f"+{d:.0f}", C_GREEN
        elif d < 0:
            text, color = f"{d:.0f}", C_RED
        else:
            text, color = "0", C_GREY
        row_idx = i + 1
        tbl.rows[row_idx].height = Emu(int(row_h * EMU_PER_IN))
        _render_data_row(tbl, row_idx, text, color, row_h,
                         data_idx=i, font_name=font_name, font_size_pt=font_size_pt)

    tbl.columns[0].width = Inches(width)
    return tbl


def add_delta_table(slide, deltas: list[float | None],
                    left: float, top: float, width: float, row_height: float,
                    header_text: str = "QoQ \u0394", font_name: Optional[str] = None,
                    display_font: Optional[str] = None,
                    show_header: bool = True,
                    delta_threshold: float = 5.0,
                    show_pct: bool = True):
    """Add a single-column delta table with conditional coloring.

    Args:
        deltas: list of float/None values (percentage points).
        row_height: height of each data row in inches.
        show_header: if False, omit the header row.
        delta_threshold: minimum |delta| to apply green/red coloring.
            Below this threshold, deltas render in neutral grey.
        show_pct: if True, format as integer with '%' suffix (e.g. "+7%").
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
        _render_header(tbl, header_text, font_name, display_font=display_font)
        data_offset = 1
    else:
        data_offset = 0

    for i, d in enumerate(deltas):
        text, color = _delta_text_color(d, threshold=delta_threshold,
                                        show_pct=show_pct)
        _render_data_row(tbl, i + data_offset, text, color, row_height, font_name,
                         data_idx=i)

    tbl.columns[0].width = Inches(width)
    return tbl_shape


def add_value_table(slide, values: list[float | None],
                    left: float, top: float, width: float, row_height: float,
                    header_text: str = "Total %", value_color: RGBColor | None = None,
                    font_name: Optional[str] = None):
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
        _render_data_row(tbl, i + 1, text, vc, row_height, font_name, data_idx=i)

    tbl.columns[0].width = Inches(width)
    return tbl_shape
