"""
charts.py — High-level chart builders for slide renderers.

Reusable chart + table patterns: bar charts (single, clustered) with
data labels, and the CHART_PATTERNS{} configuration dict (PRD section 4.5).
"""

from __future__ import annotations

from typing import Optional

from lxml import etree

from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from .brand import C_RYB_Q4, C_TAG, C_LTGREY, C_WHITE, FONT_TEXT
from .lxml_helpers import (
    _get_or_add, suppress_para_bullets, suppress_cat_axis_bullets,
    set_data_label_color, set_series_color,
    set_series_no_border, hide_axis, invert_cat_axis, set_plot_area_gap, set_overlap,
)


# ── CHART_PATTERNS{} dict (PRD §4.5) ────────────────────────────────────────

CHART_PATTERNS = {
    "single_bar": {
        "gap": 80,
        "overlap": 0,
        "label_fsize": 8,
        "label_pos": "outEnd",
    },
    "clustered_bar": {
        "gap": 80,
        "overlap": -10,
        "label_fsize": 7,
        "label_pos": "outEnd",
    },
    "stacked_bar": {
        "gap": 80,
        "label_fsize": 6,
        "label_pos": "ctr",
    },
}


# ── Shared chart helpers ─────────────────────────────────────────────────────


def _configure_bar_axes(ch, gap: int = 80, overlap: int = 0,
                        cat_font_size: float = 7, invert: bool = True,
                        font_name: Optional[str] = None,
                        hide_cats: bool = False,
                        plot_area: Optional[tuple] = None,
                        val_scale: Optional[tuple] = None) -> None:
    """Apply common bar chart axis configuration: hide val axis, style cat axis, set gap.

    Args:
        hide_cats: If True, hide category labels (for charts with external label tables).
        plot_area: (x, y, w, h) fractions for set_chart_plot_area. Default: no change.
        val_scale: (min_val, max_val) for set_val_axis_scale. Default: auto.
    """
    from .lxml_helpers import hide_cat_labels, set_chart_plot_area, set_val_axis_scale
    hide_axis(ch, "val")
    ch.category_axis.has_major_gridlines = False
    if hide_cats:
        hide_cat_labels(ch)
    else:
        ch.category_axis.tick_labels.font.size = Pt(cat_font_size)
        ch.category_axis.tick_labels.font.name = font_name or FONT_TEXT
        suppress_cat_axis_bullets(ch)
    if invert:
        invert_cat_axis(ch)
    set_plot_area_gap(ch, gap)
    if overlap:
        set_overlap(ch, overlap)
    if plot_area:
        set_chart_plot_area(ch, x=plot_area[0], y=plot_area[1],
                            w=plot_area[2], h=plot_area[3])
    if val_scale:
        set_val_axis_scale(ch, val_scale[0], val_scale[1])


def _style_bar_series(series, fill_color: RGBColor, label_color: RGBColor = C_WHITE,
                      label_fsize: float = 8, label_pos: str = "inEnd",
                      num_fmt: str = '0"%"', font_name: Optional[str] = None) -> None:
    """Apply common series styling: color, no border, data labels."""
    set_series_color(series, fill_color)
    set_series_no_border(series)
    enable_data_labels(series, label_color, fsize=label_fsize, pos=label_pos,
                       num_fmt=num_fmt, font_name=font_name)


# ── Chart builder functions ──────────────────────────────────────────────────

def enable_data_labels(series, color: RGBColor, fsize: float = 8,
                       num_fmt: str = '0"%"', pos: str = "outEnd",
                       font_name: Optional[str] = None) -> None:
    """Enable and style data labels on a chart series.

    Args:
        pos: 'outEnd' for regular bars, 'ctr' for stacked bars.
        font_name: override font (default uses FONT_TEXT).
    """
    plot = series._element.getparent()
    plot_dLbls = plot.find(qn("c:dLbls"))
    if plot_dLbls is None:
        plot_dLbls = etree.SubElement(plot, qn("c:dLbls"))
    _get_or_add(plot_dLbls, "c:showLegendKey").set("val", "0")
    _get_or_add(plot_dLbls, "c:showVal").set("val", "1")
    _get_or_add(plot_dLbls, "c:showCatName").set("val", "0")
    _get_or_add(plot_dLbls, "c:showSerName").set("val", "0")
    _get_or_add(plot_dLbls, "c:showPercent").set("val", "0")

    dLbls = _get_or_add(series._element, "c:dLbls")
    _get_or_add(dLbls, "c:showLegendKey").set("val", "0")
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
    _get_or_add(pPr, "a:buNone")
    defRPr = _get_or_add(pPr, "a:defRPr")
    defRPr.set("sz", str(int(fsize * 100)))
    defRPr.set("b", "1")
    sf = _get_or_add(defRPr, "a:solidFill")
    clr = _get_or_add(sf, "a:srgbClr")
    clr.set("val", str(color))
    latin = _get_or_add(defRPr, "a:latin")
    latin.set("typeface", font_name or FONT_TEXT)


def delete_data_label(series, point_idx: int) -> None:
    """Hide the data label for a specific point (e.g. hide small segments in stacked bars)."""
    if point_idx < 0 or point_idx >= len(series.values):
        return
    dLbls = series._element.find(qn("c:dLbls"))
    if dLbls is not None:
        dLbl = etree.SubElement(dLbls, qn("c:dLbl"))
        idx_el = etree.SubElement(dLbl, qn("c:idx"))
        idx_el.set("val", str(point_idx))
        delete_el = etree.SubElement(dLbl, qn("c:delete"))
        delete_el.set("val", "1")


def add_single_bar_chart(slide, categories: list[str], values: list[float],
                         left: float, top: float, width: float, height: float,
                         fill_color: RGBColor, cat_font_size: float = 7,
                         gap: int = 80, font_name: Optional[str] = None):
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
    ch.has_title = False

    s = ch.series[0]
    _style_bar_series(s, fill_color, C_WHITE, label_pos="inEnd", font_name=font_name)
    _configure_bar_axes(ch, gap=gap, cat_font_size=cat_font_size, font_name=font_name)

    return cf, ch


def add_clustered_bar_chart(slide, categories: list[str],
                            series_list: list[tuple[str, list[float]]],
                            left: float, top: float, width: float, height: float,
                            colors: Optional[list[RGBColor]] = None,
                            legend: bool = True, gap: int = 100, overlap: int = 0,
                            cat_font_size: float = 7, label_fsize: float = 7,
                            font_name: Optional[str] = None):
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
    ch.has_title = False

    if colors is None:
        colors = [C_RYB_Q4, C_TAG]

    for idx, series in enumerate(ch.series):
        c = colors[idx] if idx < len(colors) else C_LTGREY
        _style_bar_series(series, c, c, label_fsize=label_fsize, label_pos="inEnd",
                          font_name=font_name)

    _configure_bar_axes(ch, gap=gap, overlap=overlap, cat_font_size=cat_font_size,
                        font_name=font_name)

    if legend:
        ch.has_legend = True
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(8)
        ch.legend.font.name = font_name or FONT_TEXT
    else:
        ch.has_legend = False

    return cf, ch


def add_line_chart(slide, categories: list[str],
                   series_list: list[tuple[str, list[float]]],
                   left: float, top: float, width: float, height: float,
                   colors: Optional[list[RGBColor]] = None,
                   marker_size: int = 5, line_width_pt: float = 2.25,
                   show_labels: bool = True, label_fsize: float = 9,
                   num_fmt: str = '0"%"',
                   label_positions: Optional[list[str]] = None,
                   scale_min: Optional[float] = None,
                   scale_max: Optional[float] = None,
                   hide_axes: bool = True,
                   font_name: Optional[str] = None):
    """Add a line chart with circle markers and data labels.

    series_list: [("Series Name", [vals...]), ...]
    colors: list of RGBColor, one per series.
    label_positions: per-series label position list — e.g. ["t", "b", "b"].
        Defaults to first series "t" (above), rest "b" (below) to avoid overlap.
    Returns (chart_frame, chart).
    """
    from .lxml_helpers import (
        set_series_marker, set_series_line_style, set_series_smooth,
    )

    chart_data = CategoryChartData()
    chart_data.categories = categories
    for name, vals in series_list:
        chart_data.add_series(name, vals)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        Inches(left), Inches(top), Inches(width), Inches(height),
        chart_data)
    ch = cf.chart
    ch.has_title = False
    ch.has_legend = False

    if colors is None:
        colors = [C_RYB_Q4, C_TAG]

    # Default label positions: first series above, rest below
    if label_positions is None:
        label_positions = ["t"] + ["b"] * (len(series_list) - 1)

    for idx, series in enumerate(ch.series):
        c = colors[idx] if idx < len(colors) else C_LTGREY
        set_series_color(series, c)
        set_series_marker(series, "circle", marker_size, fill_color=c, line_color=c)
        set_series_line_style(series, width_pt=line_width_pt)
        set_series_smooth(series, False)
        if show_labels:
            pos = label_positions[idx] if idx < len(label_positions) else "b"
            enable_data_labels(series, c, fsize=label_fsize, num_fmt=num_fmt,
                               pos=pos, font_name=font_name)

    if hide_axes:
        hide_axis(ch, "val")
        hide_axis(ch, "cat")
    else:
        ch.category_axis.has_major_gridlines = False
        ch.category_axis.tick_labels.font.size = Pt(8)
        ch.category_axis.tick_labels.font.name = font_name or FONT_TEXT
        suppress_cat_axis_bullets(ch)

    if scale_min is not None or scale_max is not None:
        ax = ch.value_axis._element
        scaling = _get_or_add(ax, "c:scaling")
        if scale_min is not None:
            _get_or_add(scaling, "c:min").set("val", str(scale_min))
        if scale_max is not None:
            _get_or_add(scaling, "c:max").set("val", str(scale_max))

    return cf, ch


def add_stacked_column_chart(slide, categories: list[str],
                              series_list: list[tuple[str, list[float]]],
                              left: float, top: float, width: float, height: float,
                              colors: Optional[list[RGBColor]] = None,
                              label_fsize: float = 10,
                              num_fmt: str = '0"%"',
                              gap: int = 80,
                              font_name: Optional[str] = None):
    """Add a stacked vertical column chart with data labels centered in segments.

    series_list: [("Series Name", [vals...]), ...]
    colors: list of RGBColor, one per series.
    Returns (chart_frame, chart).
    """
    chart_data = CategoryChartData()
    chart_data.categories = categories
    for name, vals in series_list:
        chart_data.add_series(name, vals)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_STACKED,
        Inches(left), Inches(top), Inches(width), Inches(height),
        chart_data)
    ch = cf.chart
    ch.has_title = False
    ch.has_legend = False

    if colors is None:
        colors = [C_RYB_Q4, C_TAG]

    for idx, series in enumerate(ch.series):
        c = colors[idx] if idx < len(colors) else C_LTGREY
        _style_bar_series(series, c, C_WHITE, label_fsize=label_fsize,
                          label_pos="ctr", num_fmt=num_fmt, font_name=font_name)

    _configure_bar_axes(ch, gap=gap, cat_font_size=8, invert=False, font_name=font_name)

    return cf, ch


def add_scatter_chart(slide, series_list: list[tuple[str, list[float], list[float]]],
                       left: float, top: float, width: float, height: float,
                       colors: Optional[list[RGBColor]] = None,
                       marker_size: int = 7,
                       x_min: float = 0, x_max: float = 1.0,
                       y_min: float = 0, y_max: float = 1.0,
                       font_name: Optional[str] = None):
    """Add an XY scatter chart with circle markers and no connecting lines.

    series_list: [("Series Name", [x_vals...], [y_vals...]), ...]
    Returns (chart_frame, chart).
    """
    from .lxml_helpers import set_series_marker, set_series_line_style

    chart_data = XyChartData()
    for name, x_vals, y_vals in series_list:
        ser = chart_data.add_series(name)
        for x, y in zip(x_vals, y_vals):
            ser.add_data_point(x, y)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.XY_SCATTER,
        Inches(left), Inches(top), Inches(width), Inches(height),
        chart_data)
    ch = cf.chart
    ch.has_title = False
    ch.has_legend = False

    if colors is None:
        colors = [C_RYB_Q4, C_TAG]

    for idx, series in enumerate(ch.series):
        c = colors[idx] if idx < len(colors) else C_LTGREY
        set_series_marker(series, "circle", marker_size, fill_color=c, line_color=c)
        set_series_line_style(series, visible=False)

    # Set axis scales — scatter has two value axes
    plot_el = ch._element.find(qn("c:chart")).find(qn("c:plotArea"))
    val_axes = plot_el.findall(qn("c:valAx"))
    for i, ax in enumerate(val_axes):
        scaling = _get_or_add(ax, "c:scaling")
        if i == 0:  # X axis
            _get_or_add(scaling, "c:min").set("val", str(x_min))
            _get_or_add(scaling, "c:max").set("val", str(x_max))
        else:  # Y axis
            _get_or_add(scaling, "c:min").set("val", str(y_min))
            _get_or_add(scaling, "c:max").set("val", str(y_max))
        # Hide axis
        _get_or_add(ax, "c:delete").set("val", "1")
        _get_or_add(ax, "c:tickLblPos").set("val", "none")
        for gtag in ("c:majorGridlines", "c:minorGridlines"):
            child = ax.find(qn(gtag))
            if child is not None:
                ax.remove(child)

    return cf, ch
