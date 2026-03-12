"""
charts.py — High-level chart builders for slide renderers.

Reusable chart + table patterns: bar charts (single, clustered) with
data labels, and the CHART_PATTERNS{} configuration dict (PRD section 4.5).
"""

from __future__ import annotations

from typing import Optional

from lxml import etree

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from .brand import C_RYB_Q4, C_TAG, C_LTGREY, FONT_TEXT
from .lxml_helpers import (
    _get_or_add, set_data_label_color, set_series_color, set_series_no_border,
    hide_axis, invert_cat_axis, set_plot_area_gap, set_overlap,
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


def delete_data_label(series, point_idx: int) -> None:
    """Hide the data label for a specific point (e.g. hide small segments in stacked bars)."""
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
