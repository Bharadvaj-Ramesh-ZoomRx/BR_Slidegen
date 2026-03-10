"""
pptx_utils.py
─────────────
SlideGen shared utility library.
Covers everything needed to create and live-edit client-delivery PowerPoint slides.

Sections:
  1. Brand constants
  2. lxml XML helpers  (things python-pptx cannot do natively)
  3. python-pptx shape builders  (creation)
  4. COM helpers  (live editing via win32com)
  5. Registry helpers

Import pattern:
    from slidegen.pptx_utils import *
"""

import json
import os
from datetime import datetime
from lxml import etree

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt, Emu

# ══════════════════════════════════════════════════════════════════════════════
# 1. BRAND CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════

# Slide dimensions
SLIDE_W_IN = 13.333
SLIDE_H_IN = 7.500
SLIDE_W_EMU = Emu(12192000)
SLIDE_H_EMU = Emu(6858000)

# J&J / Rybrevant brand colours
C_RYB_Q4   = RGBColor(0xF7, 0x58, 0x24)   # deep orange  — Q4 bars, primary accent
C_RYB_Q3   = RGBColor(0xFF, 0xC1, 0x99)   # pale orange  — Q3 bars
C_TAG      = RGBColor(0x70, 0x30, 0xA0)   # purple       — AZ / Tagrisso
C_RED      = RGBColor(0xFF, 0x00, 0x00)   # J&J red      — title bar, headline
C_GREEN    = RGBColor(0x00, 0xB0, 0x50)   # positive delta
C_WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
C_GREY     = RGBColor(0x50, 0x50, 0x50)   # body text
C_FTGREY   = RGBColor(0x7F, 0x7F, 0x7F)   # footer / faint text
C_LBGREY   = RGBColor(0xF4, 0xF4, 0xF4)   # alternating table row bg
C_HDRGREY  = RGBColor(0x40, 0x40, 0x40)   # delta table header bg
C_LTGREY   = RGBColor(0xBF, 0xBF, 0xBF)   # gridlines / borders

# COM colour equivalents (BGR order — red=0x0000FF, green=0x0050B0)
COM_RED    = 0x0000FF
COM_ORANGE = 0x2458F7
COM_GREEN  = 0x50B000
COM_GREY   = 0x505050

# Fonts
FONT_DISPLAY = "Johnson Display"
FONT_TEXT    = "Johnson Text"

# Registry path
try:
    from slidegen.config import REGISTRY_PATH
except ImportError:
    REGISTRY_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "slide_registry.json")

# Unit conversion
IN = 72          # 1 inch = 72 points (COM)
EMU_PER_IN = 914400

# ══════════════════════════════════════════════════════════════════════════════
# 2. lxml XML HELPERS
#    These cover the gaps in python-pptx's API.
#    Never write raw lxml in slide scripts — use these functions instead.
# ══════════════════════════════════════════════════════════════════════════════

def _get_or_add(parent, tag):
    """Get an existing child XML element or create it if absent."""
    el = parent.find(qn(tag))
    if el is None:
        el = etree.SubElement(parent, qn(tag))
    return el


def invert_cat_axis(chart):
    """Show first category at top of a horizontal bar chart (maxMin orientation).
    Call after chart creation. Without this, highest-value items appear at bottom."""
    catAx   = chart.category_axis._element
    scaling = _get_or_add(catAx, "c:scaling")
    orient  = _get_or_add(scaling, "c:orientation")
    orient.set("val", "maxMin")


def hide_cat_labels(chart):
    """Hide the category-axis tick labels (Y-axis on horizontal bar).
    Use on the right-hand chart when the left chart already shows the labels."""
    catAx = chart.category_axis._element
    tlp   = _get_or_add(catAx, "c:tickLblPos")
    tlp.set("val", "none")


def set_datalabel_pos_outside_end(series):
    """Force data labels to appear outside-end (right of bar for horizontal charts)."""
    dLbls = series._element.find(qn("c:dLbls"))
    if dLbls is not None:
        pos = _get_or_add(dLbls, "c:dLblPos")
        pos.set("val", "outEnd")


def set_series_no_border(series):
    """Remove the visible border line on a bar series."""
    spPr = series._element.get_or_add_spPr()
    ln   = _get_or_add(spPr, "a:ln")
    _get_or_add(ln, "a:noFill")


def set_val_axis_number_format(axis, fmt="0"):
    """Set number format on value-axis tick labels (e.g. '0' for integers, '0%')."""
    axEl   = axis._element
    numFmt = _get_or_add(axEl, "c:numFmt")
    numFmt.set("formatCode", fmt)
    numFmt.set("sourceLinked", "0")


def set_plot_area_gap(chart, gap_pct=80):
    """Set gap between bar clusters (% of bar width). Lower = fatter bars.
    Typical values: 50 (fat), 80 (standard), 150 (thin)."""
    barChart = chart.plots[0]._element
    gapWidth = barChart.find(qn("c:gapWidth"))
    if gapWidth is None:
        gapWidth = etree.SubElement(barChart, qn("c:gapWidth"))
    gapWidth.set("val", str(gap_pct))


def set_overlap(chart, overlap=0):
    """Set bar overlap within a cluster. Negative = gap between bars in cluster.
    Typical: 0 (touching), -10 (small gap), -30 (wider gap)."""
    barChart = chart.plots[0]._element
    ov = barChart.find(qn("c:overlap"))
    if ov is None:
        ov = etree.SubElement(barChart, qn("c:overlap"))
    ov.set("val", str(overlap))


def set_series_marker(series, marker_type="circle", size=10,
                      fill_color=None, line_color=None):
    """Set marker style on a chart series (line / scatter / dot-plot charts).

    Args:
        series: python-pptx Series object
        marker_type: OOXML symbol — "circle", "diamond", "square", "triangle",
                     "star", "dot", "dash", "plus", "x", "none"
        size: marker size in points (default 10)
        fill_color: RGBColor for fill, or None to skip
        line_color: RGBColor for border, or None to skip
    """
    ser_el = series._element
    marker = _get_or_add(ser_el, "c:marker")
    sym = _get_or_add(marker, "c:symbol")
    sym.set("val", marker_type)
    sz = _get_or_add(marker, "c:size")
    sz.set("val", str(size))
    if fill_color is not None or line_color is not None:
        spPr = _get_or_add(marker, "c:spPr")
        if fill_color is not None:
            sf = _get_or_add(spPr, "a:solidFill")
            clr = _get_or_add(sf, "a:srgbClr")
            clr.set("val", str(fill_color))
        if line_color is not None:
            ln = _get_or_add(spPr, "a:ln")
            lsf = _get_or_add(ln, "a:solidFill")
            lclr = _get_or_add(lsf, "a:srgbClr")
            lclr.set("val", str(line_color))


def set_series_line_style(series, width_pt=1.5, dash="solid", visible=True):
    """Set line style on a chart series (line / combo charts).

    Args:
        series: python-pptx Series object
        width_pt: line width in points (default 1.5)
        dash: "solid", "dash", "dashDot", "dot", "lgDash", "lgDashDot",
              "sysDash", "sysDot"
        visible: False to hide the connecting line entirely
    """
    spPr = series._element.get_or_add_spPr()
    ln = _get_or_add(spPr, "a:ln")
    if not visible:
        for child in list(ln):
            ln.remove(child)
        _get_or_add(ln, "a:noFill")
    else:
        ln.set("w", str(int(width_pt * 12700)))   # 1 pt = 12700 EMU
        noFill = ln.find(qn("a:noFill"))
        if noFill is not None:
            ln.remove(noFill)
        prstDash = _get_or_add(ln, "a:prstDash")
        prstDash.set("val", dash)


def set_series_smooth(series, smooth=True):
    """Toggle smooth (bezier) vs straight line segments for a line series.

    Args:
        series: python-pptx Series object
        smooth: True for smooth curves, False for straight segments
    """
    smooth_el = _get_or_add(series._element, "c:smooth")
    smooth_el.set("val", "1" if smooth else "0")


def set_marker_data_label_pos(series, pos="r"):
    """Set the position of data labels on a line+marker series.

    Args:
        series: python-pptx Series object
        pos: "r" (right), "l" (left), "t" (above), "b" (below), "ctr" (center)
    """
    dLbls = series._element.find(qn("c:dLbls"))
    if dLbls is not None:
        _get_or_add(dLbls, "c:dLblPos").set("val", pos)


def set_data_label_color(series, color_rgb):
    """Override text color inside data labels for a series.

    python-pptx exposes data labels but not their run-level font color.
    This writes a txPr/defRPr/solidFill element directly into the dLbls XML.

    Args:
        series: python-pptx Series object
        color_rgb: RGBColor (e.g. C_RYB_Q4 for orange J&J labels)
    """
    dLbls = series._element.find(qn("c:dLbls"))
    if dLbls is None:
        return
    txPr   = _get_or_add(dLbls, "c:txPr")
    _get_or_add(txPr, "a:bodyPr")
    _get_or_add(txPr, "a:lstStyle")
    p      = _get_or_add(txPr, "a:p")
    pPr    = _get_or_add(p, "a:pPr")
    defRPr = _get_or_add(pPr, "a:defRPr")
    sf     = _get_or_add(defRPr, "a:solidFill")
    clr    = _get_or_add(sf, "a:srgbClr")
    clr.set("val", str(color_rgb))


def add_val_axis_reference_line(chart, x_value, label="",
                                 color=None, dash="dash", width_pt=1.0):
    """Inject a vertical reference line at a fixed value axis position.

    Adds a supplementary scatter series pinned to x_value with y spanning
    -100 to 100, so it renders as a vertical dashed line across the plot.
    Use for "Industry Average = 35%" markers on horizontal bar charts.

    Args:
        chart: python-pptx Chart object
        x_value: float — position on the value axis (e.g. 35 for 35%)
        label: optional series name shown in legend (default "")
        color: RGBColor for the line (default C_LTGREY)
        dash: "solid", "dash", "dot" (default "dash")
        width_pt: line width in points (default 1.0)
    """
    if color is None:
        color = C_LTGREY
    plotArea = chart._element.find(qn("c:plotArea"))
    existing_sers = plotArea.findall(".//" + qn("c:ser"))
    idx  = len(existing_sers)
    hex_c = str(color)
    lw    = int(width_pt * 12700)

    C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

    scatter_xml = (
        f'<c:scatterChart xmlns:c="{C_NS}" xmlns:a="{A_NS}">'
        f'<c:scatterStyle val="line"/><c:varyColors val="0"/>'
        f'<c:ser>'
        f'<c:idx val="{idx}"/><c:order val="{idx}"/>'
        f'<c:spPr><a:ln w="{lw}"><a:solidFill><a:srgbClr val="{hex_c}"/></a:solidFill>'
        f'<a:prstDash val="{dash}"/></a:ln></c:spPr>'
        f'<c:marker><c:symbol val="none"/></c:marker>'
        f'<c:xVal><c:numRef><c:numCache>'
        f'<c:formatCode>General</c:formatCode><c:ptCount val="2"/>'
        f'<c:pt idx="0"><c:v>{x_value}</c:v></c:pt>'
        f'<c:pt idx="1"><c:v>{x_value}</c:v></c:pt>'
        f'</c:numCache></c:numRef></c:xVal>'
        f'<c:yVal><c:numRef><c:numCache>'
        f'<c:formatCode>General</c:formatCode><c:ptCount val="2"/>'
        f'<c:pt idx="0"><c:v>-100</c:v></c:pt>'
        f'<c:pt idx="1"><c:v>100</c:v></c:pt>'
        f'</c:numCache></c:numRef></c:yVal>'
        f'</c:ser>'
        f'<c:axId val="99991"/><c:axId val="99992"/>'
        f'</c:scatterChart>'
    )
    val_ax_xml = (
        f'<c:valAx xmlns:c="{C_NS}">'
        f'<c:axId val="99991"/>'
        f'<c:scaling><c:orientation val="minMax"/></c:scaling>'
        f'<c:delete val="1"/><c:axPos val="b"/>'
        f'<c:crossAx val="99992"/></c:valAx>'
    )
    cat_ax_xml = (
        f'<c:valAx xmlns:c="{C_NS}">'
        f'<c:axId val="99992"/>'
        f'<c:scaling><c:orientation val="minMax"/></c:scaling>'
        f'<c:delete val="1"/><c:axPos val="l"/>'
        f'<c:crossAx val="99991"/></c:valAx>'
    )
    plotArea.append(etree.fromstring(scatter_xml))
    plotArea.append(etree.fromstring(val_ax_xml))
    plotArea.append(etree.fromstring(cat_ax_xml))


def set_stacked_label_pos(series, pos="ctr"):
    """Set data label position inside stacked bar segments.

    The existing set_datalabel_pos_outside_end only handles outEnd.
    Use this for stacked/100%-stacked charts.

    Args:
        series: python-pptx Series object
        pos: "inBase", "inEnd", "ctr" (inside center), "outEnd"
    """
    dLbls = series._element.find(qn("c:dLbls"))
    if dLbls is not None:
        _get_or_add(dLbls, "c:dLblPos").set("val", pos)


def set_pie_slice_colors(chart, colors):
    """Set fill color per slice on a pie or donut chart by index.

    python-pptx creates default theme-colored slices; this writes
    <a:solidFill> into each <c:dPt> data-point element.

    Args:
        chart: python-pptx Chart object (pie or donut type)
        colors: list of RGBColor, one per slice in category order
    """
    ser_el = chart.series[0]._element
    for i, color in enumerate(colors):
        dPt    = etree.SubElement(ser_el, qn("c:dPt"))
        idx_el = etree.SubElement(dPt, qn("c:idx"))
        idx_el.set("val", str(i))
        b3d    = etree.SubElement(dPt, qn("c:bubble3D"))
        b3d.set("val", "0")
        spPr   = etree.SubElement(dPt, qn("c:spPr"))
        sf     = etree.SubElement(spPr, qn("a:solidFill"))
        clr    = etree.SubElement(sf, qn("a:srgbClr"))
        clr.set("val", str(color))
        ln     = etree.SubElement(spPr, qn("a:ln"))
        etree.SubElement(ln, qn("a:noFill"))


def set_donut_hole_size(chart, pct=50):
    """Control the inner hole radius of a donut chart.

    python-pptx doesn't expose <c:holeSize>; without this call the
    default hole may be wrong for the client template.

    Args:
        chart: python-pptx Chart object (donut type)
        pct: integer 10–90 — hole diameter as % of chart diameter (default 50)
    """
    plot_el = chart.plots[0]._element
    _get_or_add(plot_el, "c:holeSize").set("val", str(int(pct)))


def hide_axis(chart, axis="val"):
    """Fully suppress a chart axis (line, ticks, labels, and gridlines).

    The existing hide_cat_labels only removes label text; this deletes
    the axis entirely and strips major/minor gridlines.

    Args:
        chart: python-pptx Chart object
        axis: "val" (value axis) or "cat" (category axis)
    """
    ax_el = (chart.value_axis if axis == "val" else chart.category_axis)._element
    _get_or_add(ax_el, "c:delete").set("val", "1")
    _get_or_add(ax_el, "c:tickLblPos").set("val", "none")
    for tag in ("c:majorGridlines", "c:minorGridlines"):
        child = ax_el.find(qn(tag))
        if child is not None:
            ax_el.remove(child)


def set_gridlines(chart, axis="val", major=True, minor=False):
    """Enable or disable major/minor gridlines on a chart axis.

    Args:
        chart: python-pptx Chart object
        axis: "val" (value axis) or "cat" (category axis)
        major: True to show major gridlines (default True)
        minor: True to show minor gridlines (default False)
    """
    ax_el = (chart.value_axis if axis == "val" else chart.category_axis)._element
    for tag, show in (("c:majorGridlines", major), ("c:minorGridlines", minor)):
        existing = ax_el.find(qn(tag))
        if show and existing is None:
            etree.SubElement(ax_el, qn(tag))
        elif not show and existing is not None:
            ax_el.remove(existing)


def set_series_color(series, fill_color, line_color=None):
    """Set the fill (and optionally border) color of a bar or line series.

    There is no high-level python-pptx API for this; the function writes
    directly into the series spPr element.

    Args:
        series: python-pptx Series object
        fill_color: RGBColor for the series fill
        line_color: RGBColor for the series border, or None for no border
    """
    spPr = series._element.get_or_add_spPr()
    for tag in ("a:noFill", "a:solidFill", "a:gradFill", "a:pattFill"):
        existing = spPr.find(qn(tag))
        if existing is not None:
            spPr.remove(existing)
    sf  = etree.SubElement(spPr, qn("a:solidFill"))
    clr = etree.SubElement(sf, qn("a:srgbClr"))
    clr.set("val", str(fill_color))
    if line_color is not None:
        ln   = _get_or_add(spPr, "a:ln")
        lsf  = _get_or_add(ln, "a:solidFill")
        lclr = _get_or_add(lsf, "a:srgbClr")
        lclr.set("val", str(line_color))
    else:
        ln = _get_or_add(spPr, "a:ln")
        for child in list(ln):
            ln.remove(child)
        _get_or_add(ln, "a:noFill")


# ══════════════════════════════════════════════════════════════════════════════
# 3. python-pptx SHAPE BUILDERS  (creation — all positions in inches)
# ══════════════════════════════════════════════════════════════════════════════

def textbox(slide, text, left, top, width, height,
            fsize=9, bold=False, color=None, align=PP_ALIGN.LEFT,
            italic=False, wrap=True, font=FONT_TEXT):
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


def solidrect(slide, left, top, width, height, fill, line=None):
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


def horiz_line(slide, left, top, width, color=None, width_pt=1.0):
    """Add a horizontal line. Positions in inches, width_pt in points."""
    if color is None:
        color = C_RED
    shape = slide.shapes.add_connector(
        1,   # MSO_CONNECTOR.STRAIGHT
        Inches(left), Inches(top), Inches(left + width), Inches(top))
    shape.line.color.rgb = color
    shape.line.width = Pt(width_pt)
    return shape


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


def slide_header(slide, headline, module_label="Personal Promotion Module",
                 font=None):
    """Add the standard ZoomRx slide header:
      - thin red accent line at very top
      - module label (top-right, small grey)
      - headline text (large, red, bold)
      - red module badge (top-right rectangle)
      - separator line below header
    """
    f_display = font or FONT_DISPLAY
    f_text = font or FONT_TEXT

    # Red accent line
    solidrect(slide, 0, 0.15, SLIDE_W_IN, 0.02, C_RED)

    # Module label
    textbox(slide, module_label,
            5.5, 0.01, 7.70, 0.22,
            fsize=7.5, color=C_FTGREY, align=PP_ALIGN.RIGHT, font=f_text)

    # Headline
    textbox(slide, headline,
            0.20, 0.20, 10.55, 1.05,
            fsize=12, bold=True, color=C_RED, align=PP_ALIGN.LEFT, font=f_display)

    # Module badge
    solidrect(slide, 10.90, 0.20, 2.25, 0.95, C_RED)
    textbox(slide, "PERSONAL\nPROMOTION\nMODULE",
            10.90, 0.20, 2.25, 0.95,
            fsize=8, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER, font=f_display)

    # Separator line
    horiz_line(slide, 0.0, 1.32, SLIDE_W_IN, color=C_RED, width_pt=1.0)


def slide_footer(slide, footer_text, font=None):
    """Add standard footer text at bottom of slide."""
    textbox(slide, footer_text,
            0.15, 7.20, 13.0, 0.28,
            fsize=6.0, color=C_FTGREY, align=PP_ALIGN.LEFT, font=font or FONT_TEXT)


def manual_legend(slide, q4_n, q3_n,
                  chart_left, chart_right, chart_bottom):
    """Add the shared Q3/Q4/delta colour legend below a pair of charts.

    Args:
        q4_n, q3_n: sample sizes for label text
        chart_left: leftmost x of the content area (inches)
        chart_right: rightmost x of the content area (inches)
        chart_bottom: y position of chart bottom edge (inches)
    """
    leg_top = chart_bottom + 0.10
    leg_ctr = (chart_left + chart_right) / 2
    leg_l   = leg_ctr - 2.2

    solidrect(slide, leg_l,        leg_top, 0.20, 0.14, C_RYB_Q4)
    textbox(slide,  f"Q4'25  (n={q4_n})",
            leg_l + 0.25, leg_top - 0.01, 1.20, 0.18, fsize=8, color=C_GREY)

    solidrect(slide, leg_l + 1.65, leg_top, 0.20, 0.14, C_RYB_Q3)
    textbox(slide,  f"Q3'25  (n={q3_n})",
            leg_l + 1.90, leg_top - 0.01, 1.20, 0.18, fsize=8, color=C_GREY)

    solidrect(slide, leg_l + 3.35, leg_top, 0.20, 0.14, C_GREEN)
    textbox(slide,  "Delta increase (vs Q3)",
            leg_l + 3.60, leg_top - 0.01, 1.40, 0.18, fsize=8, color=C_GREY)

    solidrect(slide, leg_l + 5.15, leg_top, 0.20, 0.14, C_RED)
    textbox(slide,  "Delta decrease (vs Q3)   N/A = new message (no Q3)",
            leg_l + 5.40, leg_top - 0.01, 3.00, 0.18, fsize=8, color=C_GREY)


def callout_box(slide, left, top, width, height, text=None,
                border_color=None, dashed=True, fill_color=None,
                fsize=8, text_color=None):
    """Add a rounded-rectangle annotation callout box.

    Appears on ~30/45 data slides as a floating annotation with a dashed
    colored border and a light tint fill matching the border color.

    Args:
        slide: python-pptx Slide
        left/top/width/height: position and size in inches
        text: optional text string inside the box
        border_color: RGBColor for the border (default C_RED)
        dashed: True for dashed border, False for solid (default True)
        fill_color: RGBColor for fill; None auto-computes a 12% tint of border_color
        fsize: font size for text content (default 8)
        text_color: RGBColor for text (default C_GREY)
    Returns: the shape object
    """
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


def section_header_bar(slide, label, top=1.40, icon_path=None, font=None):
    """Add the gray icon+label strip used as a chart/question title on every data slide.

    Creates a light-gray rectangle spanning most of the slide width, with an
    optional 16×16px icon at the left edge and bold all-caps label text.

    Args:
        slide: python-pptx Slide
        label: text label shown in the bar (uppercased automatically)
        top: y position in inches (default 1.40, just below slide_header)
        icon_path: path to a small PNG icon; None skips the icon
        font: optional font name override
    Returns: the background rect shape
    """
    bar_l = 0.15
    bar_w = SLIDE_W_IN - 0.30
    bar_h = 0.28

    bg = solidrect(slide, bar_l, top, bar_w, bar_h, C_LBGREY, line=C_LTGREY)

    text_l = bar_l + 0.10
    if icon_path and os.path.exists(icon_path):
        slide.shapes.add_picture(
            icon_path,
            Inches(bar_l + 0.05), Inches(top + 0.04),
            Inches(0.20),          Inches(0.20))
        text_l = bar_l + 0.32

    textbox(slide, label.upper(),
            text_l, top + 0.03, bar_w - 0.35, bar_h - 0.06,
            fsize=9, bold=True, color=C_GREY, font=font or FONT_TEXT)
    return bg


def module_badge(slide, label, color=None):
    """Add a colored pill badge in the top-right corner labeling the current module.

    Distinct from slide_header — this is module-scoped, not slide-scoped.
    Examples: "PERSONAL PROMOTION MODULE" (red), "NON-SALES REP PROMOTIONS" (green).

    Args:
        slide: python-pptx Slide
        label: badge text (uppercased automatically)
        color: RGBColor fill color (default C_RED)
    """
    if color is None:
        color = C_RED
    badge_w = 2.25
    badge_h = 0.28
    badge_l = SLIDE_W_IN - badge_w - 0.10
    badge_t = 0.05
    solidrect(slide, badge_l, badge_t, badge_w, badge_h, color)
    textbox(slide, label.upper(),
            badge_l, badge_t, badge_w, badge_h,
            fsize=7, bold=True, color=C_WHITE,
            align=PP_ALIGN.CENTER, font=FONT_TEXT)


def section_breadcrumb(slide, text):
    """Add small right-aligned breadcrumb text in the top-right showing section context.

    E.g. "Key Findings - Personal Promotions". Different from module_badge.

    Args:
        slide: python-pptx Slide
        text: breadcrumb string
    """
    textbox(slide, text,
            8.0, 0.36, 5.20, 0.20,
            fsize=7, color=C_FTGREY,
            align=PP_ALIGN.RIGHT, font=FONT_TEXT)


def divider_slide(slide, title, logo_path=None):
    """Build a full-slide section divider.

    Off-white background with a red left-edge stripe, large red title,
    and optional J&J logo in the bottom-left corner.

    Args:
        slide: python-pptx Slide
        title: section title text
        logo_path: path to logo PNG; None skips it
    """
    solidrect(slide, 0, 0, SLIDE_W_IN, SLIDE_H_IN, RGBColor(0xF8, 0xF8, 0xF8))
    solidrect(slide, 0, 0, 0.08, SLIDE_H_IN, C_RED)
    textbox(slide, title,
            0.60, 2.30, 9.0, 3.0,
            fsize=36, bold=True, color=C_RED,
            align=PP_ALIGN.LEFT, font=FONT_DISPLAY)
    if logo_path and os.path.exists(logo_path):
        slide.shapes.add_picture(
            logo_path,
            Inches(0.30), Inches(SLIDE_H_IN - 0.80),
            Inches(1.50), Inches(0.50))


def cover_slide(slide, title, subtitle, date, client_name,
                jj_logo_path=None, zrx_logo_path=None):
    """Build the full red-background title cover slide.

    Args:
        slide: python-pptx Slide (should be blank layout)
        title: large white title text
        subtitle: smaller white subtitle
        date: date string (e.g. "Q4 2025")
        client_name: attribution line bottom-right (e.g. "Johnson & Johnson")
        jj_logo_path: path to J&J logo PNG; None skips it
        zrx_logo_path: path to ZoomRx logo PNG; None skips it
    """
    solidrect(slide, 0, 0, SLIDE_W_IN, SLIDE_H_IN, C_RED)
    textbox(slide, title,
            0.60, 1.50, 11.0, 2.50,
            fsize=32, bold=True, color=C_WHITE,
            align=PP_ALIGN.LEFT, font=FONT_DISPLAY)
    textbox(slide, subtitle,
            0.60, 4.20, 10.0, 0.80,
            fsize=16, color=C_WHITE,
            align=PP_ALIGN.LEFT, font=FONT_TEXT)
    textbox(slide, date,
            0.60, 5.10, 5.0, 0.40,
            fsize=10, color=C_WHITE, font=FONT_TEXT)
    textbox(slide, client_name,
            7.0, 6.80, 6.10, 0.40,
            fsize=8, color=C_WHITE,
            align=PP_ALIGN.RIGHT, font=FONT_TEXT)
    if jj_logo_path and os.path.exists(jj_logo_path):
        slide.shapes.add_picture(
            jj_logo_path,
            Inches(0.40), Inches(6.70), Inches(1.50), Inches(0.50))
    if zrx_logo_path and os.path.exists(zrx_logo_path):
        slide.shapes.add_picture(
            zrx_logo_path,
            Inches(2.10), Inches(6.70), Inches(1.50), Inches(0.50))


def stat_callout(slide, value, delta, label, left, top):
    """Add a large single-stat display: oversized number, delta, and label.

    Example: value="16", delta=-1, label="years avg treatment duration"
    renders as "16" in large red, "(-1)" in small grey, label beneath.

    Args:
        slide: python-pptx Slide
        value: string or number for the primary stat
        delta: int/float for the change vs prior period, or None to skip
        label: descriptor text beneath the stat
        left/top: position of the callout block in inches
    """
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


def insert_image(slide, img_path, left, top, width, height, name=None):
    """Place an external PNG or JPG at specified inch coordinates.

    Wrapper around add_picture that handles inch-to-EMU conversion and
    optional shape naming for the registry.

    Args:
        slide: python-pptx Slide
        img_path: path to image file; returns None if file does not exist
        left/top/width/height: position and size in inches
        name: optional shape name (e.g. 'zrx_img_001') for registry tracking
    Returns: the picture shape, or None if img_path does not exist
    """
    if not os.path.exists(img_path):
        return None
    pic = slide.shapes.add_picture(
        img_path,
        Inches(left), Inches(top), Inches(width), Inches(height))
    if name:
        pic.name = name
    return pic


def dashed_separator(slide, left, top, width,
                      color=None, width_pt=0.75, dash="dash"):
    """Add a horizontal dashed line separator (not solid — use horiz_line for solid).

    Used to visually divide a slide into upper/lower chart panels.

    Args:
        slide: python-pptx Slide
        left/top/width: position and length in inches
        color: RGBColor (default C_LTGREY)
        width_pt: line weight in points (default 0.75)
        dash: prstDash value — "dash", "dot", "dashDot", "lgDash" (default "dash")
    Returns: the connector shape
    """
    if color is None:
        color = C_LTGREY
    shape = slide.shapes.add_connector(
        1,  # MSO_CONNECTOR.STRAIGHT
        Inches(left), Inches(top),
        Inches(left + width), Inches(top))
    shape.line.color.rgb = color
    shape.line.width = Pt(width_pt)
    # Set dash via XML (python-pptx dash_style enum not always reliable)
    spPr = shape._element.spPr
    ln   = _get_or_add(spPr, "a:ln")
    _get_or_add(ln, "a:prstDash").set("val", dash)
    return shape


def trend_arrow_icon(slide, direction, left, top):
    """Add a small directional trend arrow icon at a position.

    Used inline in scorecard cells to indicate QoQ direction.

    Args:
        slide: python-pptx Slide
        direction: "up" (green triangle up), "down" (red triangle down),
                   "flat" (yellow double-arrow)
        left/top: position in inches; icon occupies ~0.22 x 0.22 inches
    """
    colors = {"up": C_GREEN, "down": C_RED,
              "flat": RGBColor(0xFF, 0xC0, 0x00)}
    chars  = {"up": u"\u25b2", "down": u"\u25bc", "flat": u"\u25c4\u25ba"}
    color  = colors.get(direction, C_GREY)
    char   = chars.get(direction, "?")
    textbox(slide, char,
            left, top, 0.22, 0.22,
            fsize=9, bold=True, color=color,
            align=PP_ALIGN.CENTER, font=FONT_TEXT)


def scatter_quadrant_fills(slide, chart_left, chart_top,
                            chart_width, chart_height,
                            tl_color=None, tr_color=None,
                            bl_color=None, br_color=None):
    """Place four colored background rectangles behind a scatter chart plot area.

    Produces the four-quadrant colored background used on the stated-vs-derived
    importance scatter maps (slides 30-31).  Call this BEFORE adding the chart
    so the rects sit behind it in z-order.

    Args:
        slide: python-pptx Slide
        chart_left/chart_top/chart_width/chart_height: chart dimensions in inches
        tl_color: top-left quadrant fill (default light blue)
        tr_color: top-right quadrant fill (default light green)
        bl_color: bottom-left quadrant fill (default light grey)
        br_color: bottom-right quadrant fill (default light yellow)
    """
    if tl_color is None: tl_color = RGBColor(0xDE, 0xEB, 0xF7)  # light blue
    if tr_color is None: tr_color = RGBColor(0xE2, 0xEF, 0xDA)  # light green
    if bl_color is None: bl_color = RGBColor(0xF4, 0xF4, 0xF4)  # light grey
    if br_color is None: br_color = RGBColor(0xFF, 0xFF, 0xCC)  # light yellow

    cx = chart_left  + chart_width  / 2
    cy = chart_top   + chart_height / 2
    hw = chart_width  / 2
    hh = chart_height / 2

    solidrect(slide, chart_left, chart_top, hw, hh, tl_color)  # top-left
    solidrect(slide, cx,         chart_top, hw, hh, tr_color)  # top-right
    solidrect(slide, chart_left, cy,        hw, hh, bl_color)  # bottom-left
    solidrect(slide, cx,         cy,        hw, hh, br_color)  # bottom-right


# ══════════════════════════════════════════════════════════════════════════════
# 4. COM HELPERS  (live editing via win32com — all positions in inches)
# ══════════════════════════════════════════════════════════════════════════════

def com_connect(target_filename):
    """Connect to a running PowerPoint instance and return the named presentation.

    Args:
        target_filename: bare filename, e.g. 'phase1_test.pptx' (not full path)
    Returns: win32com Presentation object
    Raises: RuntimeError if PowerPoint is not open or file not found
    """
    import win32com.client
    try:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    except Exception as e:
        raise RuntimeError(f"Could not connect to PowerPoint: {e}")

    for i in range(1, ppt_app.Presentations.Count + 1):
        p = ppt_app.Presentations(i)
        if target_filename in p.Name:
            return p

    open_files = [ppt_app.Presentations(i).Name
                  for i in range(1, ppt_app.Presentations.Count + 1)]
    raise RuntimeError(
        f"'{target_filename}' not found in open presentations.\n"
        f"  Open: {open_files}"
    )


def com_find_shape(com_slide, name):
    """Find a shape on a COM slide by its zrx_ name.

    Args:
        com_slide: win32com Slide object (1-indexed)
        name: shape name string, e.g. 'zrx_001'
    Returns: win32com Shape object
    Raises: RuntimeError with clear message if not found
    """
    for i in range(1, com_slide.Shapes.Count + 1):
        sh = com_slide.Shapes(i)
        if sh.Name == name:
            return sh
    raise RuntimeError(
        f"Shape '{name}' not found on slide. "
        f"Run reconcile_registry.py and check for renames."
    )


def com_set_text(shape, text, color_bgr=None, size_pt=None, bold=None):
    """Set text content and optional formatting on a COM shape.

    Args:
        shape: win32com Shape object with a TextFrame
        text: new text string
        color_bgr: int in BGR order (COM convention), e.g. 0x0000FF for red
        size_pt: font size in points
        bold: True/False
    """
    tr = shape.TextFrame.TextRange
    tr.Text = text
    if color_bgr is not None:
        tr.Font.Color.RGB = color_bgr
    if size_pt is not None:
        tr.Font.Size = size_pt
    if bold is not None:
        tr.Font.Bold = bold


def com_set_fill(shape, color_bgr):
    """Set fill colour on a COM shape (BGR int, e.g. 0x2458F7 for orange)."""
    shape.Fill.ForeColor.RGB = color_bgr


def com_move(shape, left_in, top_in):
    """Move a COM shape to a new position (inches)."""
    shape.Left = left_in * IN
    shape.Top  = top_in  * IN


def com_resize(shape, width_in, height_in):
    """Resize a COM shape (inches)."""
    shape.Width  = width_in  * IN
    shape.Height = height_in * IN


def com_get_position(shape):
    """Return current (left, top, width, height) in inches from a COM shape."""
    return (
        round(shape.Left   / IN, 4),
        round(shape.Top    / IN, 4),
        round(shape.Width  / IN, 4),
        round(shape.Height / IN, 4),
    )


# ══════════════════════════════════════════════════════════════════════════════
# 5. REGISTRY HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def load_registry(path=None):
    """Load slide_registry.json and return the dict."""
    if path is None:
        path = REGISTRY_PATH
    with open(path) as f:
        return json.load(f)


def save_registry(registry, path=None):
    """Save registry dict back to slide_registry.json."""
    if path is None:
        path = REGISTRY_PATH
    with open(path, "w") as f:
        json.dump(registry, f, indent=2)


def registry_get(name, path=None):
    """Return the record for a single shape by name. Raises KeyError if not found."""
    reg = load_registry(path)
    if name not in reg["shapes"]:
        raise KeyError(
            f"Shape '{name}' not in registry. "
            f"Run phase2_create.py or reconcile_registry.py first."
        )
    return reg["shapes"][name]


def registry_tag_slide(slide_idx, module, section, data_source="", path=None):
    """Store a slide-level metadata record in the registry.

    Enables 'rebuild slide 14 from scratch' without touching other slides.
    Records are stored under reg['slides'][str(slide_idx)].

    Args:
        slide_idx: int — 0-based slide index
        module: module name (e.g. "Personal Promotion Module")
        section: section name (e.g. "Key Findings - Personal Promotions")
        data_source: data source description (e.g. "Lung SFEA SB.xlsx / RYB sheet")
        path: registry file path (default REGISTRY_PATH)
    """
    reg = load_registry(path)
    if "slides" not in reg:
        reg["slides"] = {}
    reg["slides"][str(slide_idx)] = {
        "module":      module,
        "section":     section,
        "data_source": data_source,
        "tagged_at":   datetime.now().isoformat(),
    }
    save_registry(reg, path)


def registry_find_by_type(shape_type, slide_idx=None, path=None):
    """Query the registry for all shapes matching a type string.

    Enables bulk COM operations such as 'update all footer text on all slides'
    without knowing individual shape names.

    Args:
        shape_type: string — "chart", "textbox", "rect", "image", etc.
        slide_idx: int or None — if set, restrict results to one slide
        path: registry file path (default REGISTRY_PATH)
    Returns: list of (name, record) tuples matching the query
    """
    reg    = load_registry(path)
    shapes = reg.get("shapes", {})
    return [
        (name, record)
        for name, record in shapes.items()
        if record.get("type") == shape_type
        and (slide_idx is None or record.get("slide_idx") == slide_idx)
    ]


def registry_diff_slide(slide_idx, com_slide, path=None):
    """Diff the registry snapshot against live COM state for a single slide.

    More ergonomic than reconcile_registry.py when you only care about one slide.

    Args:
        slide_idx: int — 0-based slide index
        com_slide: win32com Slide object (pass prs.Slides(slide_idx + 1))
        path: registry file path (default REGISTRY_PATH)
    Returns: list of dicts with keys: name, field, registry_val, live_val
    """
    reg    = load_registry(path)
    shapes = reg.get("shapes", {})

    live = {}
    for i in range(1, com_slide.Shapes.Count + 1):
        sh = com_slide.Shapes(i)
        live[sh.Name] = {
            "left":   round(sh.Left   / IN, 4),
            "top":    round(sh.Top    / IN, 4),
            "width":  round(sh.Width  / IN, 4),
            "height": round(sh.Height / IN, 4),
        }

    diffs = []
    for name, record in shapes.items():
        if record.get("slide_idx") != slide_idx:
            continue
        if name not in live:
            diffs.append({"name": name, "field": "existence",
                          "registry_val": "present", "live_val": "missing"})
            continue
        for field in ("left", "top", "width", "height"):
            r_val = record.get(field)
            l_val = live[name].get(field)
            if r_val is not None and l_val is not None:
                if abs(r_val - l_val) > 0.01:
                    diffs.append({"name": name, "field": field,
                                  "registry_val": r_val, "live_val": l_val})
    return diffs


# ══════════════════════════════════════════════════════════════════════════════
# 6. HIGH-LEVEL CHART BUILDERS
#    Reusable chart + table patterns used by the pipeline slide renderers.
# ══════════════════════════════════════════════════════════════════════════════

from pptx.chart.data import CategoryChartData


def enable_data_labels(series, color, fsize=8, num_fmt='0"%"', pos="outEnd",
                       font_name=None):
    """Enable and style data labels on a chart series.

    Args:
        pos: 'outEnd' for regular bars, 'ctr' for stacked bars.
        font_name: override font (default uses FONT_TEXT).
    """
    plot = series._element.getparent()
    plot_dLbls = plot.find(qn("c:dLbls"))
    if plot_dLbls is None:
        plot_dLbls = etree.SubElement(plot, qn("c:dLbls"))
    _get_or_add(plot_dLbls, "c:showVal").set("val", "1")
    _get_or_add(plot_dLbls, "c:showCatName").set("val", "0")
    _get_or_add(plot_dLbls, "c:showSerName").set("val", "0")
    _get_or_add(plot_dLbls, "c:showPercent").set("val", "0")

    dLbls = _get_or_add(series._element, "c:dLbls")
    _get_or_add(dLbls, "c:showVal").set("val", "1")
    _get_or_add(dLbls, "c:showCatName").set("val", "0")
    _get_or_add(dLbls, "c:showSerName").set("val", "0")
    _get_or_add(dLbls, "c:showPercent").set("val", "0")
    numFmt = _get_or_add(dLbls, "c:numFmt")
    numFmt.set("formatCode", num_fmt)
    numFmt.set("sourceLinked", "0")
    dLblPos = _get_or_add(dLbls, "c:dLblPos")
    dLblPos.set("val", pos)
    set_data_label_color(series, color)

    # Font
    txPr = _get_or_add(dLbls, "c:txPr")
    _get_or_add(txPr, "a:bodyPr")
    _get_or_add(txPr, "a:lstStyle")
    p = _get_or_add(txPr, "a:p")
    pPr = _get_or_add(p, "a:pPr")
    defRPr = _get_or_add(pPr, "a:defRPr")
    defRPr.set("sz", str(int(fsize * 100)))
    defRPr.set("b", "1")
    sf = _get_or_add(defRPr, "a:solidFill")
    clr = _get_or_add(sf, "a:srgbClr")
    clr.set("val", str(color))
    latin = _get_or_add(defRPr, "a:latin")
    latin.set("typeface", font_name or FONT_TEXT)


def delete_data_label(series, point_idx):
    """Hide the data label for a specific point (e.g. hide small segments in stacked bars)."""
    dLbls = series._element.find(qn("c:dLbls"))
    if dLbls is not None:
        dLbl = etree.SubElement(dLbls, qn("c:dLbl"))
        idx_el = etree.SubElement(dLbl, qn("c:idx"))
        idx_el.set("val", str(point_idx))
        delete_el = etree.SubElement(dLbl, qn("c:delete"))
        delete_el.set("val", "1")


def add_single_bar_chart(slide, categories, values, left, top, width, height,
                         fill_color, cat_font_size=7, gap=80, font_name=None):
    """Add a horizontal bar chart with one series + data labels.

    Returns (chart_frame, chart).
    """
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series("Values", values)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(left), Inches(top), Inches(width), Inches(height),
        chart_data)
    ch = cf.chart
    ch.has_legend = False

    s = ch.series[0]
    set_series_color(s, fill_color)
    set_series_no_border(s)
    enable_data_labels(s, fill_color, font_name=font_name)

    hide_axis(ch, "val")
    ch.category_axis.has_major_gridlines = False
    ch.category_axis.tick_labels.font.size = Pt(cat_font_size)
    ch.category_axis.tick_labels.font.name = font_name or FONT_TEXT
    invert_cat_axis(ch)
    set_plot_area_gap(ch, gap)

    return cf, ch


def add_clustered_bar_chart(slide, categories, series_list, left, top, width, height,
                            colors=None, legend=True, gap=100, overlap=0,
                            cat_font_size=7, label_fsize=7, font_name=None):
    """Add a clustered horizontal bar chart with multiple series.

    series_list: [("Series Name", [vals...]), ...]
    colors: list of RGBColor, one per series.
    Returns (chart_frame, chart).
    """
    chart_data = CategoryChartData()
    chart_data.categories = categories
    for name, vals in series_list:
        chart_data.add_series(name, vals)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(left), Inches(top), Inches(width), Inches(height),
        chart_data)
    ch = cf.chart

    if colors is None:
        colors = [C_RYB_Q4, C_TAG]

    for idx, series in enumerate(ch.series):
        c = colors[idx] if idx < len(colors) else C_LTGREY
        set_series_color(series, c)
        set_series_no_border(series)
        enable_data_labels(series, c, fsize=label_fsize, font_name=font_name)

    set_plot_area_gap(ch, gap)
    set_overlap(ch, overlap)
    hide_axis(ch, "val")
    ch.category_axis.has_major_gridlines = False
    ch.category_axis.tick_labels.font.size = Pt(cat_font_size)
    ch.category_axis.tick_labels.font.name = font_name or FONT_TEXT
    invert_cat_axis(ch)

    if legend:
        ch.has_legend = True
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(8)
        ch.legend.font.name = font_name or FONT_TEXT
    else:
        ch.has_legend = False

    return cf, ch


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
