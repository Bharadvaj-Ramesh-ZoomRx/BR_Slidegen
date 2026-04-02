"""
lxml_helpers.py — Raw OOXML XML manipulation functions.

These cover the gaps in python-pptx's API. Never write raw lxml in slide
scripts — use these functions instead.

PRD §4.2 / Appendix D: All lxml manipulation centralized here.
"""

from __future__ import annotations

from typing import Optional

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

from .brand import C_LTGREY


def _get_or_add(parent: etree._Element, tag: str) -> etree._Element:
    """Get an existing child XML element or create it if absent."""
    el = parent.find(qn(tag))
    if el is None:
        el = etree.SubElement(parent, qn(tag))
    return el


def suppress_para_bullets(p_el: etree._Element) -> None:
    """Add <a:buNone/> to <a:pPr> to suppress inherited paragraph bullets."""
    pPr = _get_or_add(p_el, "a:pPr")
    _get_or_add(pPr, "a:buNone")


def cell_vcenter(cell) -> None:
    """Vertically center text in a pptx table cell via <a:tcPr anchor='ctr'/>.

    Table cell vertical alignment lives on <a:tcPr>, NOT on <a:bodyPr>.
    """
    tc = cell._tc
    tcPr = tc.find(qn("a:tcPr"))
    if tcPr is None:
        tcPr = etree.SubElement(tc, qn("a:tcPr"))
    tcPr.set("anchor", "ctr")


def suppress_cat_axis_bullets(ch) -> None:
    """Suppress inherited bullet markers on chart category axis tick labels."""
    cat_ax = ch.category_axis._element
    txPr = _get_or_add(cat_ax, "c:txPr")
    _get_or_add(txPr, "a:bodyPr")
    _get_or_add(txPr, "a:lstStyle")
    p = _get_or_add(txPr, "a:p")
    pPr = _get_or_add(p, "a:pPr")
    _get_or_add(pPr, "a:buNone")


def invert_cat_axis(chart) -> None:
    """Show first category at top of a horizontal bar chart (maxMin orientation).
    Call after chart creation. Without this, highest-value items appear at bottom."""
    catAx   = chart.category_axis._element
    scaling = _get_or_add(catAx, "c:scaling")
    orient  = _get_or_add(scaling, "c:orientation")
    orient.set("val", "maxMin")


def hide_cat_labels(chart) -> None:
    """Hide the category-axis tick labels (Y-axis on horizontal bar).
    Use on the right-hand chart when the left chart already shows the labels."""
    catAx = chart.category_axis._element
    tlp   = _get_or_add(catAx, "c:tickLblPos")
    tlp.set("val", "none")


def set_datalabel_pos_outside_end(series) -> None:
    """Force data labels to appear outside-end (right of bar for horizontal charts)."""
    dLbls = series._element.find(qn("c:dLbls"))
    if dLbls is not None:
        pos = _get_or_add(dLbls, "c:dLblPos")
        pos.set("val", "outEnd")


def set_series_no_border(series) -> None:
    """Remove the visible border line on a bar series."""
    spPr = series._element.get_or_add_spPr()
    ln   = _get_or_add(spPr, "a:ln")
    _get_or_add(ln, "a:noFill")


def set_val_axis_number_format(axis, fmt: str = "0") -> None:
    """Set number format on value-axis tick labels (e.g. '0' for integers, '0%')."""
    axEl   = axis._element
    numFmt = _get_or_add(axEl, "c:numFmt")
    numFmt.set("formatCode", fmt)
    numFmt.set("sourceLinked", "0")


def set_plot_area_gap(chart, gap_pct: int = 80) -> None:
    """Set gap between bar clusters (% of bar width). Lower = fatter bars.
    Typical values: 50 (fat), 80 (standard), 150 (thin)."""
    barChart = chart.plots[0]._element
    gapWidth = barChart.find(qn("c:gapWidth"))
    if gapWidth is None:
        gapWidth = etree.SubElement(barChart, qn("c:gapWidth"))
    gapWidth.set("val", str(gap_pct))


def set_overlap(chart, overlap: int = 0) -> None:
    """Set bar overlap within a cluster. Negative = gap between bars in cluster.
    Typical: 0 (touching), -10 (small gap), -30 (wider gap)."""
    barChart = chart.plots[0]._element
    ov = barChart.find(qn("c:overlap"))
    if ov is None:
        ov = etree.SubElement(barChart, qn("c:overlap"))
    ov.set("val", str(overlap))


_VALID_MARKER_TYPES = frozenset({
    "circle", "diamond", "square", "triangle", "star",
    "dot", "dash", "plus", "x", "none",
})


def set_series_marker(series, marker_type: str = "circle", size: int = 10,
                      fill_color: Optional[RGBColor] = None,
                      line_color: Optional[RGBColor] = None) -> None:
    """Set marker style on a chart series (line / scatter / dot-plot charts).

    Args:
        series: python-pptx Series object
        marker_type: OOXML symbol — "circle", "diamond", "square", "triangle",
                     "star", "dot", "dash", "plus", "x", "none"
        size: marker size in points (2-72, default 10)
        fill_color: RGBColor for fill, or None to skip
        line_color: RGBColor for border, or None to skip
    """
    if marker_type not in _VALID_MARKER_TYPES:
        marker_type = "circle"
    size = max(2, min(72, size))

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


def set_series_line_style(series, width_pt: float = 1.5, dash: str = "solid",
                          visible: bool = True) -> None:
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


def set_series_smooth(series, smooth: bool = True) -> None:
    """Toggle smooth (bezier) vs straight line segments for a line series."""
    smooth_el = _get_or_add(series._element, "c:smooth")
    smooth_el.set("val", "1" if smooth else "0")


def set_marker_data_label_pos(series, pos: str = "r") -> None:
    """Set the position of data labels on a line+marker series.

    Args:
        pos: "r" (right), "l" (left), "t" (above), "b" (below), "ctr" (center)
    """
    dLbls = series._element.find(qn("c:dLbls"))
    if dLbls is not None:
        _get_or_add(dLbls, "c:dLblPos").set("val", pos)


def set_data_label_color(series, color_rgb: RGBColor) -> None:
    """Override text color inside data labels for a series.

    python-pptx exposes data labels but not their run-level font color.
    This writes a txPr/defRPr/solidFill element directly into the dLbls XML.
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


def add_val_axis_reference_line(chart, x_value: float, label: str = "",
                                 color: Optional[RGBColor] = None,
                                 dash: str = "dash", width_pt: float = 1.0) -> None:
    """Inject a vertical reference line at a fixed value axis position.

    Adds a supplementary scatter series pinned to x_value with y spanning
    -100 to 100, so it renders as a vertical dashed line across the plot.
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


def set_stacked_label_pos(series, pos: str = "ctr") -> None:
    """Set data label position inside stacked bar segments."""
    dLbls = series._element.find(qn("c:dLbls"))
    if dLbls is not None:
        _get_or_add(dLbls, "c:dLblPos").set("val", pos)


def set_pie_slice_colors(chart, colors: list[RGBColor]) -> None:
    """Set fill color per slice on a pie or donut chart by index."""
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


def set_donut_hole_size(chart, pct: int = 50) -> None:
    """Control the inner hole radius of a donut chart."""
    plot_el = chart.plots[0]._element
    _get_or_add(plot_el, "c:holeSize").set("val", str(int(pct)))


def hide_axis(chart, axis: str = "val") -> None:
    """Fully suppress a chart axis (line, ticks, labels, and gridlines)."""
    ax_el = (chart.value_axis if axis == "val" else chart.category_axis)._element
    _get_or_add(ax_el, "c:delete").set("val", "1")
    _get_or_add(ax_el, "c:tickLblPos").set("val", "none")
    for tag in ("c:majorGridlines", "c:minorGridlines"):
        child = ax_el.find(qn(tag))
        if child is not None:
            ax_el.remove(child)


def set_gridlines(chart, axis: str = "val", major: bool = True,
                   minor: bool = False) -> None:
    """Enable or disable major/minor gridlines on a chart axis."""
    ax_el = (chart.value_axis if axis == "val" else chart.category_axis)._element
    for tag, show in (("c:majorGridlines", major), ("c:minorGridlines", minor)):
        existing = ax_el.find(qn(tag))
        if show and existing is None:
            etree.SubElement(ax_el, qn(tag))
        elif not show and existing is not None:
            ax_el.remove(existing)


def set_series_color(series, fill_color: RGBColor,
                     line_color: Optional[RGBColor] = None) -> None:
    """Set the fill (and optionally border) color of a bar or line series."""
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


def set_chart_plot_area(chart, x: float = 0.0, y: float = 0.0,
                        w: float = 1.0, h: float = 1.0) -> None:
    """Set chart plot area position using manual layout (fractions of chart frame).

    Args:
        x, y: top-left corner as fraction (0.0–1.0)
        w, h: width/height as fraction (0.0–1.0)
    """
    chart_el = chart._element.find(qn("c:chart"))
    plotArea = chart_el.find(qn("c:plotArea"))
    layout = _get_or_add(plotArea, "c:layout")
    ml = _get_or_add(layout, "c:manualLayout")
    for tag, val in [("c:xMode", "edge"), ("c:yMode", "edge"),
                     ("c:x", str(x)), ("c:y", str(y)),
                     ("c:w", str(w)), ("c:h", str(h))]:
        elem = _get_or_add(ml, tag)
        elem.set("val", val)


def set_val_axis_scale(chart, min_val: float = 0, max_val: float = 100) -> None:
    """Set explicit min/max scale on the value axis.

    Use to synchronize bar lengths across multiple charts so that
    the same percentage value produces the same visual bar width.
    """
    ax = chart.value_axis._element
    scaling = _get_or_add(ax, "c:scaling")
    min_el = _get_or_add(scaling, "c:min")
    min_el.set("val", str(min_val))
    max_el = _get_or_add(scaling, "c:max")
    max_el.set("val", str(max_val))
