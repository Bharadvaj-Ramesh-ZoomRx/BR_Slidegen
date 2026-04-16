"""
OOXML helpers — raw XML manipulation for properties python-pptx does not expose.

Generated from real-deck analysis of 32 PET decks. Each function wraps a
specific OOXML property that appeared frequently enough to warrant a named helper.
See experiments/deck_analysis/outputs/ACTIONABLE_FINDINGS.md §2 for the
observed frequencies that justified each helper.

All functions operate on python-pptx Chart or Series objects and mutate their
underlying lxml element tree. No new presentation objects are created here.
"""
from __future__ import annotations

from lxml import etree


NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NSMAP = {"c": NS_C, "a": NS_A}


def qn_c(local: str) -> str:
    """Qualified chart namespace tag."""
    return f"{{{NS_C}}}{local}"


def qn_a(local: str) -> str:
    """Qualified drawingML namespace tag."""
    return f"{{{NS_A}}}{local}"


def _find_or_create(parent, tag: str, insert_before: list[str] | None = None):
    """Find child by qualified tag, or create it (inserted before listed siblings)."""
    el = parent.find(tag)
    if el is not None:
        return el
    el = etree.SubElement(parent, tag)
    return el


# ============================================================================
# BAR / COLUMN GEOMETRY
# ============================================================================

def set_plot_area_gap(chart, gap_width: int = 80) -> None:
    """Set <c:gapWidth> on bar/column plot.

    gap_width: percentage of bar thickness; 0 = bars touching, 200 = gap equals
    bar width. Observed range 40-150 in real decks, with 50/80/100 most common.
    Default 80 matches the most common PET bar chart.
    """
    plot_area = chart._chartSpace.find(f".//{qn_c('plotArea')}")
    if plot_area is None:
        return
    for tag in ('barChart', 'bar3DChart'):
        for bar_chart in plot_area.findall(qn_c(tag)):
            gw = bar_chart.find(qn_c('gapWidth'))
            if gw is None:
                gw = etree.SubElement(bar_chart, qn_c('gapWidth'))
            gw.set('val', str(gap_width))


def set_overlap(chart, overlap_pct: int = 0) -> None:
    """Set <c:overlap> on bar/column plot.

    overlap_pct: -100 to 100. 100 = fully overlapped (stacked). 0 = clustered
    side by side. Negative = gap between bars within a cluster. Observed: 100
    dominates (stacked charts), -20 next most common (slight separation).
    """
    plot_area = chart._chartSpace.find(f".//{qn_c('plotArea')}")
    if plot_area is None:
        return
    for tag in ('barChart', 'bar3DChart'):
        for bar_chart in plot_area.findall(qn_c(tag)):
            ov = bar_chart.find(qn_c('overlap'))
            if ov is None:
                ov = etree.SubElement(bar_chart, qn_c('overlap'))
            ov.set('val', str(overlap_pct))


def set_invert_if_negative(chart, value: bool = False) -> None:
    """Toggle <c:invertIfNegative> on all series. Default False disables the
    auto-color-flip on negative values (observed: 13,550 series want this off).
    """
    for ser in chart._chartSpace.findall(f".//{qn_c('ser')}"):
        iin = ser.find(qn_c('invertIfNegative'))
        if iin is None:
            iin = etree.SubElement(ser, qn_c('invertIfNegative'))
        iin.set('val', '1' if value else '0')


# ============================================================================
# AXIS CONFIGURATION
# ============================================================================

def invert_cat_axis(chart) -> None:
    """Set category axis orientation to maxMin (top-down).

    For horizontal bar charts, this makes the first category appear at the top,
    matching standard PET layout. Observed in 43% of charts (1,865 / 4,354).
    """
    _set_axis_orientation(chart, 'catAx', 'maxMin')


def invert_val_axis(chart) -> None:
    """Set value axis orientation to maxMin (rare — only 115 charts)."""
    _set_axis_orientation(chart, 'valAx', 'maxMin')


def _set_axis_orientation(chart, axis_tag: str, orientation: str) -> None:
    for ax in chart._chartSpace.findall(f".//{qn_c(axis_tag)}"):
        scaling = ax.find(qn_c('scaling'))
        if scaling is None:
            scaling = etree.SubElement(ax, qn_c('scaling'))
        orient = scaling.find(qn_c('orientation'))
        if orient is None:
            orient = etree.SubElement(scaling, qn_c('orientation'))
        orient.set('val', orientation)


def hide_cat_labels(chart) -> None:
    """Set <c:tickLblPos val="none"> on category axis.

    Used when category labels appear in a companion table instead of on the axis.
    Core clustered_compare pattern. Observed in 371 charts.
    """
    _set_tick_lbl_pos(chart, 'catAx', 'none')


def hide_val_labels(chart) -> None:
    """Hide value axis labels (less common, 48 charts)."""
    _set_tick_lbl_pos(chart, 'valAx', 'none')


def _set_tick_lbl_pos(chart, axis_tag: str, pos: str) -> None:
    for ax in chart._chartSpace.findall(f".//{qn_c(axis_tag)}"):
        tlp = ax.find(qn_c('tickLblPos'))
        if tlp is None:
            tlp = etree.SubElement(ax, qn_c('tickLblPos'))
        tlp.set('val', pos)


# ============================================================================
# DATA LABELS
# ============================================================================

def set_datalabel_pos_center(chart) -> None:
    """Center (ctr) — 4,508 occurrences, most common for stacked bars."""
    _set_datalabel_pos(chart, 'ctr')


def set_datalabel_pos_top(chart) -> None:
    """Top (t) — 2,778 occurrences, common for markers/scatter."""
    _set_datalabel_pos(chart, 't')


def set_datalabel_pos_outside_end(chart) -> None:
    """Outside end (outEnd) — 2,618 occurrences, common for clustered bars
    where labels sit outside bar ends.
    """
    _set_datalabel_pos(chart, 'outEnd')


def set_datalabel_pos_inside_end(chart) -> None:
    """Inside end (inEnd) — labels inside bar with white text, prevents overflow.
    143 occurrences, used on single_bar_with_delta patterns.
    """
    _set_datalabel_pos(chart, 'inEnd')


def _set_datalabel_pos(chart, pos: str) -> None:
    for ser in chart._chartSpace.findall(f".//{qn_c('ser')}"):
        dlbls = ser.find(qn_c('dLbls'))
        if dlbls is None:
            dlbls = etree.SubElement(ser, qn_c('dLbls'))
        dlp = dlbls.find(qn_c('dLblPos'))
        if dlp is None:
            dlp = etree.SubElement(dlbls, qn_c('dLblPos'))
        dlp.set('val', pos)


# ============================================================================
# NUMBER FORMATS
# ============================================================================

def set_val_axis_pct_format(chart) -> None:
    """Apply \"0%\" to val axis — 96% of all chart number formats in PET decks."""
    _set_val_axis_num_format(chart, '0%')


def set_val_axis_int_format(chart) -> None:
    """Apply \"0\" (integer) to val axis — 13% of charts."""
    _set_val_axis_num_format(chart, '0')


def _set_val_axis_num_format(chart, fmt: str) -> None:
    for ax in chart._chartSpace.findall(f".//{qn_c('valAx')}"):
        nf = ax.find(qn_c('numFmt'))
        if nf is None:
            nf = etree.SubElement(ax, qn_c('numFmt'))
        nf.set('formatCode', fmt)
        nf.set('sourceLinked', '0')


def set_datalabel_format(chart, fmt: str = '0%') -> None:
    """Apply a number format to all series data labels.

    Common formats:
      '0%'       — integer percent (96% of labels)
      '0'        — integer
      '0.0'      — one decimal
      '0%;\\-0%;\\ '  — percent with conditional sign (delta columns)
    """
    for ser in chart._chartSpace.findall(f".//{qn_c('ser')}"):
        dlbls = ser.find(qn_c('dLbls'))
        if dlbls is None:
            dlbls = etree.SubElement(ser, qn_c('dLbls'))
        nf = dlbls.find(qn_c('numFmt'))
        if nf is None:
            nf = etree.SubElement(dlbls, qn_c('numFmt'))
        nf.set('formatCode', fmt)
        nf.set('sourceLinked', '0')


# ============================================================================
# SERIES FORMATTING
# ============================================================================

def set_series_no_border(series) -> None:
    """Remove series border via <a:ln><a:noFill/></a:ln>. Universal pattern
    (74,367 noFill tags across all decks).
    """
    sp_pr = series._element.find(qn_c('spPr'))
    if sp_pr is None:
        sp_pr = etree.SubElement(series._element, qn_c('spPr'))
    ln = sp_pr.find(qn_a('ln'))
    if ln is None:
        ln = etree.SubElement(sp_pr, qn_a('ln'))
    # Remove existing fills and set noFill
    for child in list(ln):
        ln.remove(child)
    etree.SubElement(ln, qn_a('noFill'))


def set_series_fill_rgb(series, hex_color: str) -> None:
    """Set series fill to solid sRGB color. Prefer this over scheme colors.

    hex_color: 'RRGGBB' or '#RRGGBB' — case-insensitive.
    """
    hex_clean = hex_color.lstrip('#').upper()
    sp_pr = series._element.find(qn_c('spPr'))
    if sp_pr is None:
        sp_pr = etree.SubElement(series._element, qn_c('spPr'))
    # Remove existing fill children
    for tag in ('solidFill', 'gradFill', 'blipFill', 'pattFill', 'noFill'):
        existing = sp_pr.find(qn_a(tag))
        if existing is not None:
            sp_pr.remove(existing)
    solid = etree.SubElement(sp_pr, qn_a('solidFill'))
    rgb = etree.SubElement(solid, qn_a('srgbClr'))
    rgb.set('val', hex_clean)


def set_series_line_width(series, emu: int = 25400) -> None:
    """Set series line width in EMU. Common values:
      12700 = 1pt
      19050 = 1.5pt
      25400 = 2pt (PET default, 827 occurrences)
      28575 = 2.25pt (most common overall, 1,223 occurrences)
    """
    sp_pr = series._element.find(qn_c('spPr'))
    if sp_pr is None:
        sp_pr = etree.SubElement(series._element, qn_c('spPr'))
    ln = sp_pr.find(qn_a('ln'))
    if ln is None:
        ln = etree.SubElement(sp_pr, qn_a('ln'))
    ln.set('w', str(emu))


def set_series_marker_circle(series, size: int = 7) -> None:
    """Set marker to circle (94% of markers in scatter/line charts).

    size: marker size in points (observed range 3-12, default 7).
    """
    _set_series_marker(series, 'circle', size)


def _set_series_marker(series, symbol: str, size: int) -> None:
    marker = series._element.find(qn_c('marker'))
    if marker is None:
        marker = etree.SubElement(series._element, qn_c('marker'))
    sym = marker.find(qn_c('symbol'))
    if sym is None:
        sym = etree.SubElement(marker, qn_c('symbol'))
    sym.set('val', symbol)
    sz = marker.find(qn_c('size'))
    if sz is None:
        sz = etree.SubElement(marker, qn_c('size'))
    sz.set('val', str(size))


# ============================================================================
# GRIDLINES
# ============================================================================

def remove_major_gridlines(chart) -> None:
    """Remove major gridlines from all value axes (84% of PET charts have none)."""
    for ax in chart._chartSpace.findall(f".//{qn_c('valAx')}"):
        mg = ax.find(qn_c('majorGridlines'))
        if mg is not None:
            ax.remove(mg)


def add_major_gridlines(chart) -> None:
    """Add major gridlines (used in 16% of PET charts)."""
    for ax in chart._chartSpace.findall(f".//{qn_c('valAx')}"):
        if ax.find(qn_c('majorGridlines')) is None:
            etree.SubElement(ax, qn_c('majorGridlines'))
