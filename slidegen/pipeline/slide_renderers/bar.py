"""
bar.py — Single-bar, QoQ bar, and two-section bar slide renderers.
"""

from __future__ import annotations

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches, Pt

from ._shared import (
    # constants
    CHART_TOP_STD, LEGEND_GAP, SLIDE_W, CHART_DELTA_GAP,
    SINGLE_BAR_WIDTH, DELTA_COL_WIDTH, DELTA_COL_NARROW,
    MAX_CHART_HEIGHT, MAX_QOQ_HEIGHT, MIN_CHART_HEIGHT,
    ROW_SCALE_FACTOR,
    BAR_GAP_STD, CLUSTERED_OVERLAP,
    LABEL_MAX_SINGLE, LABEL_MAX_DUAL,
    # helpers
    _resolve_template, _slide_chrome, _get_brand_colors, _sort_data, _make_legend,
    _cap_chart_h,
    # pptx_utils
    C_GREEN, C_GREY, C_RED,
    textbox,
    hide_axis, set_series_color, set_plot_area_gap, set_overlap,
    set_series_no_border, invert_cat_axis, suppress_cat_axis_bullets,
    set_chart_plot_area,
    enable_data_labels,
    add_single_bar_chart, add_delta_table,
    # data_loaders
    delta,
    # project_config types
    ProjectConfig, AskConfig,
)


def render_single_bar_with_delta(slide, config: ProjectConfig, ask: AskConfig, data: dict, *, namer=None):
    """Single horizontal bar chart with a QoQ delta column."""
    _slide_chrome(slide, config, ask)

    rows = data.get(ask.data_key, [])
    rows = _sort_data(rows, ask.sort_by, ask.sort_desc)
    if not rows:
        textbox(slide, "Data not available", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    color_current, color_prior = _get_brand_colors(config, ask)
    font = config.font_body

    labels = [r.get("short", r.get("desc", ""))[:LABEL_MAX_SINGLE] for r in rows]
    current_vals = [r.get("current") or 0 for r in rows]
    prior_vals = [r.get("prior") for r in rows]

    n = len(labels)
    chart_top = CHART_TOP_STD
    chart_h = max(MIN_CHART_HEIGHT, min(MAX_CHART_HEIGHT, n * 0.40))
    row_h = chart_h / max(n, 1)

    # Center the chart + delta block
    block_w = SINGLE_BAR_WIDTH + CHART_DELTA_GAP + DELTA_COL_WIDTH
    chart_left = (SLIDE_W - block_w) / 2
    delta_left = chart_left + SINGLE_BAR_WIDTH + CHART_DELTA_GAP

    # Bar chart
    add_single_bar_chart(
        slide, labels, current_vals,
        left=chart_left, top=chart_top, width=SINGLE_BAR_WIDTH, height=chart_h,
        fill_color=color_current, font_name=font,
    )

    # Delta table
    deltas = [delta(c, p) if p is not None else None
              for c, p in zip(current_vals, prior_vals)]
    add_delta_table(
        slide, deltas,
        left=delta_left, top=chart_top, width=DELTA_COL_WIDTH,
        row_height=row_h * ROW_SCALE_FACTOR,
        header_text="QoQ Δ", font_name=font,
    )

    # Legend
    ly = chart_top + chart_h + LEGEND_GAP
    sample = config.sample_sizes.get(ask.brand or "primary")
    n_label = f" (n={sample.current})" if sample else ""
    _make_legend(slide, [
        (color_current, f"{config.period_current}{n_label}"),
        (C_GREEN, "Positive Δ"),
        (C_RED, "Negative Δ"),
    ], 0, ly, font, center_over=(chart_left, block_w))


def render_qoq_bar_with_delta(slide, config: ProjectConfig, ask: AskConfig, data: dict, *, namer=None):
    """Clustered Q4 vs Q3 bar chart with delta column."""
    _slide_chrome(slide, config, ask)

    rows = data.get(ask.data_key, [])
    rows = _sort_data(rows, ask.sort_by, ask.sort_desc)
    if not rows:
        textbox(slide, "Data not available", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    color_current, color_prior = _get_brand_colors(config, ask)
    font = config.font_body

    labels = [r.get("desc", "")[:LABEL_MAX_DUAL] for r in rows]
    current_vals = [r.get("current") or 0 for r in rows]
    prior_vals = [r.get("prior") or 0 for r in rows]

    n = len(labels)
    chart_top = CHART_TOP_STD
    chart_h = max(MIN_CHART_HEIGHT, min(MAX_QOQ_HEIGHT, n * 0.50))
    row_h = chart_h / max(n, 1)

    # Center the chart + delta block
    qoq_chart_w = 8.0
    block_w = qoq_chart_w + CHART_DELTA_GAP + DELTA_COL_WIDTH
    chart_left = (SLIDE_W - block_w) / 2
    delta_left = chart_left + qoq_chart_w + CHART_DELTA_GAP

    # Clustered bar
    cd = CategoryChartData()
    cd.categories = labels
    cd.add_series(config.period_current, current_vals)
    cd.add_series(config.period_prior, prior_vals)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(chart_left), Inches(chart_top), Inches(qoq_chart_w), Inches(chart_h), cd)
    ch = cf.chart
    ch.has_legend = False

    set_series_color(ch.series[0], color_current)
    set_series_no_border(ch.series[0])
    enable_data_labels(ch.series[0], color_current, font_name=font)

    set_series_color(ch.series[1], color_prior)
    set_series_no_border(ch.series[1])
    enable_data_labels(ch.series[1], color_prior, font_name=font)

    hide_axis(ch, "val")
    ch.category_axis.tick_labels.font.size = Pt(8)
    ch.category_axis.tick_labels.font.name = font
    suppress_cat_axis_bullets(ch)
    invert_cat_axis(ch)
    set_plot_area_gap(ch, BAR_GAP_STD)
    set_overlap(ch, CLUSTERED_OVERLAP)
    set_chart_plot_area(ch, x=0.0, y=0.0, w=1.0, h=1.0)

    # Delta
    deltas = [delta(c, p) for c, p in zip(current_vals, prior_vals)]
    add_delta_table(
        slide, deltas,
        left=delta_left, top=chart_top, width=DELTA_COL_WIDTH, row_height=row_h * ROW_SCALE_FACTOR,
        header_text="QoQ Δ", font_name=font,
    )

    # Manual legend below chart
    _make_legend(slide, [
        (color_current, config.period_current),
        (color_prior, config.period_prior),
    ], 0, chart_top + chart_h + LEGEND_GAP, font,
        center_over=(chart_left, block_w))


def render_two_section_bar(slide, config: ProjectConfig, ask: AskConfig, data: dict, *, namer=None):
    """Two vertically stacked single-bar sections (e.g. RYB Rx + TAG Rx)."""
    _slide_chrome(slide, config, ask)

    extra = ask.extra
    font = config.font_body
    chart_top = CHART_TOP_STD

    sections = [("top", extra.get("top", {})), ("bottom", extra.get("bottom", {}))]
    y_offset = chart_top

    for pos, sec_cfg in sections:
        sec_data_key = sec_cfg.get("data_key", "")
        sec_rows = data.get(sec_data_key, [])
        sec_label = _resolve_template(sec_cfg.get("label", ""), config)
        sec_brand = sec_cfg.get("brand", "primary")
        brand = config.brands.get(sec_brand, config.primary)

        # Center the chart + delta block
        _block_w = SINGLE_BAR_WIDTH + CHART_DELTA_GAP + DELTA_COL_NARROW
        _chart_left = (SLIDE_W - _block_w) / 2
        _delta_left = _chart_left + SINGLE_BAR_WIDTH + CHART_DELTA_GAP

        # Section label
        textbox(slide, sec_label, _chart_left, y_offset - 0.15, SINGLE_BAR_WIDTH, 0.25,
                fsize=9, bold=True, color=brand.color_current, font=config.font_display)

        if sec_rows:
            labels = [r.get("desc", "")[:LABEL_MAX_DUAL] for r in sec_rows]
            current_vals = [r.get("current") or 0 for r in sec_rows]
            prior_vals = [r.get("prior") for r in sec_rows]
            n = len(labels)
            ch_h = max(1.5, n * 0.55)

            cd = CategoryChartData()
            cd.categories = labels
            cd.add_series(config.period_current, current_vals)

            cf = slide.shapes.add_chart(
                XL_CHART_TYPE.BAR_CLUSTERED,
                Inches(_chart_left), Inches(y_offset), Inches(SINGLE_BAR_WIDTH), Inches(ch_h), cd)
            ch = cf.chart
            ch.has_legend = False
            s = ch.series[0]
            set_series_color(s, brand.color_current)
            set_series_no_border(s)
            enable_data_labels(s, brand.color_current, font_name=font)
            hide_axis(ch, "val")
            ch.category_axis.tick_labels.font.size = Pt(8)
            ch.category_axis.tick_labels.font.name = font
            suppress_cat_axis_bullets(ch)
            invert_cat_axis(ch)
            set_plot_area_gap(ch, BAR_GAP_STD)

            # Delta
            d = [delta(c, p) if p is not None else None
                 for c, p in zip(current_vals, prior_vals)]
            add_delta_table(
                slide, d,
                left=_delta_left, top=y_offset, width=DELTA_COL_NARROW,
                row_height=ch_h / max(n, 1) * ROW_SCALE_FACTOR,
                header_text="Δ", font_name=font,
            )

            y_offset += ch_h + 0.50
        else:
            y_offset += 1.7

