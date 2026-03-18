"""
compare.py — Clustered compare and stacked order slide renderers.
"""

from __future__ import annotations
import logging

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

from ._shared import (
    # constants
    CHART_TOP_STD, LEGEND_GAP, SLIDE_W, CHART_DELTA_GAP,
    CLUSTERED_BAR_WIDTH, DELTA_COL_WIDTH,
    MAX_CHART_HEIGHT, MAX_CLUSTERED_HEIGHT, MIN_CHART_HEIGHT,
    ROW_SCALE_FACTOR,
    BAR_GAP_STD, CLUSTERED_GAP, CLUSTERED_OVERLAP,
    LABEL_MAX_CLUSTERED, LABEL_MAX_STACKED, LABEL_MAX_DUAL,
    STACKED_HIDE_THRESHOLD,
    DUAL_BC_TOP, DUAL_BC_ROW_H,
    DUAL_BC_CAT_LEFT, DUAL_BC_CAT_W,
    DUAL_BC_L_CHART_L, DUAL_BC_L_CHART_W, DUAL_BC_L_DELTA_L, DUAL_BC_L_DELTA_W,
    DUAL_BC_SEP_X,
    DUAL_BC_R_CHART_L, DUAL_BC_R_CHART_W, DUAL_BC_R_DELTA_L, DUAL_BC_R_DELTA_W,
    # helpers
    _resolve_template, _slide_chrome, _get_brand_colors, _sort_data, _make_legend,
    _pptx_table, _style_tbl_cell, _cell_bottom_border, _cap_chart_h,
    # pptx_utils
    C_GREEN, C_WHITE, C_GREY, C_FTGREY, C_LBGREY, C_RED,
    PP_ALIGN,
    textbox, solidrect, dashed_separator,
    hide_axis, set_series_color, set_plot_area_gap, set_overlap,
    set_series_no_border, invert_cat_axis, hide_cat_labels,
    suppress_cat_axis_bullets,
    set_chart_plot_area, set_val_axis_scale,
    enable_data_labels, delete_data_label,
    add_single_bar_chart, add_delta_table, add_value_table,
    # data_loaders
    delta,
    # project_config types
    ProjectConfig, AskConfig, parse_color,
)

logger = logging.getLogger(__name__)


def render_clustered_compare(slide, config: ProjectConfig, ask: AskConfig, data: dict, *, namer=None):
    """Clustered horizontal bar comparing two groups + delta columns."""
    _slide_chrome(slide, config, ask)

    extra = ask.extra
    series_cfgs = extra.get("series", [])

    # Support combined data keys (CTA merges ryb_cta + tag_cta)
    rows = data.get(ask.data_key, [])

    # If data_key not found, check for primary_key + comp_key pattern
    if not rows and "primary_key" in extra:
        logger.warning(
            "clustered_compare '%s': data_key '%s' not found, merging "
            "primary_key '%s' + comp_key '%s' instead",
            ask.id, ask.data_key, extra.get("primary_key"), extra.get("comp_key"),
        )
        primary_rows = data.get(extra["primary_key"], [])
        comp_rows = data.get(extra["comp_key"], [])
        # Merge: align by desc/label
        comp_idx = {r.get("desc"): r for r in comp_rows}
        merged = []
        for pr in primary_rows:
            label = pr.get("desc", "")
            cr = comp_idx.get(label, {})
            merged.append({
                "desc": label,
                "primary_current": pr.get("current"),
                "primary_prior": pr.get("prior"),
                "comp_current": cr.get("current"),
                "comp_prior": cr.get("prior"),
            })
        rows = merged

    rows = _sort_data(rows, ask.sort_by, ask.sort_desc)
    if not rows:
        textbox(slide, "Data not available", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    font = config.font_body

    # Determine series fields and colors
    if len(series_cfgs) >= 2:
        s1_field = series_cfgs[0].get("field", "primary")
        s2_field = series_cfgs[1].get("field", "comp")
        s1_label = _resolve_template(series_cfgs[0].get("label", "Series 1"), config)
        s2_label = _resolve_template(series_cfgs[1].get("label", "Series 2"), config)
        s1_color = parse_color(series_cfgs[0]["color"]) if "color" in series_cfgs[0] else config.primary.color_current
        s2_color = parse_color(series_cfgs[1]["color"]) if "color" in series_cfgs[1] else config.competitor.color_current
    else:
        s1_field, s2_field = "primary", "comp"
        s1_label, s2_label = config.primary.name, config.competitor.name
        s1_color, s2_color = config.primary.color_current, config.competitor.color_current

    labels = [r.get("desc", "")[:LABEL_MAX_CLUSTERED] for r in rows]

    # Extract values — support both flat (hi_current, other_current) and
    # prefixed (primary_current, comp_current) naming
    s1_current = []
    s1_prior = []
    s2_current = []
    s2_prior = []
    for r in rows:
        s1_current.append(r.get(f"{s1_field}_current") or r.get(s1_field) or 0)
        s1_prior.append(r.get(f"{s1_field}_prior"))
        s2_current.append(r.get(f"{s2_field}_current") or r.get(s2_field) or 0)
        s2_prior.append(r.get(f"{s2_field}_prior"))

    n = len(labels)
    chart_top = CHART_TOP_STD
    chart_h = _cap_chart_h(min(MAX_CLUSTERED_HEIGHT, n * 0.42), chart_top)
    row_h = chart_h / max(n, 1)

    # Delta columns
    has_prior_1 = any(p is not None for p in s1_prior)
    has_prior_2 = any(p is not None for p in s2_prior)

    # Center the chart + delta block(s)
    if has_prior_1 and has_prior_2:
        # Two delta columns
        block_w = CLUSTERED_BAR_WIDTH + CHART_DELTA_GAP + DELTA_COL_WIDTH + CHART_DELTA_GAP + DELTA_COL_WIDTH
    else:
        # Single gap/delta column
        block_w = CLUSTERED_BAR_WIDTH + CHART_DELTA_GAP + DELTA_COL_WIDTH
    chart_left = (SLIDE_W - block_w) / 2
    delta1_left = chart_left + CLUSTERED_BAR_WIDTH + CHART_DELTA_GAP
    delta2_left = delta1_left + DELTA_COL_WIDTH + CHART_DELTA_GAP

    # Clustered bar
    cd = CategoryChartData()
    cd.categories = labels
    cd.add_series(s1_label, s1_current)
    cd.add_series(s2_label, s2_current)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(chart_left), Inches(chart_top), Inches(CLUSTERED_BAR_WIDTH), Inches(chart_h), cd)
    ch = cf.chart
    ch.has_legend = False

    set_series_color(ch.series[0], s1_color)
    set_series_no_border(ch.series[0])
    enable_data_labels(ch.series[0], s1_color, fsize=7, font_name=font)

    set_series_color(ch.series[1], s2_color)
    set_series_no_border(ch.series[1])
    enable_data_labels(ch.series[1], s2_color, fsize=7, font_name=font)

    hide_axis(ch, "val")
    ch.category_axis.tick_labels.font.size = Pt(6.5)
    ch.category_axis.tick_labels.font.name = font
    suppress_cat_axis_bullets(ch)
    invert_cat_axis(ch)
    set_plot_area_gap(ch, CLUSTERED_GAP)
    set_overlap(ch, -15)
    set_chart_plot_area(ch, x=0.0, y=0.0, w=1.0, h=1.0)

    # Delta columns
    if has_prior_1 and has_prior_2:
        # Two delta columns (QoQ for each series)
        d1 = [delta(c, p) if p is not None else None for c, p in zip(s1_current, s1_prior)]
        d2 = [delta(c, p) if p is not None else None for c, p in zip(s2_current, s2_prior)]
        add_delta_table(slide, d1, left=delta1_left, top=chart_top, width=DELTA_COL_WIDTH,
                        row_height=row_h * ROW_SCALE_FACTOR,
                        header_text=s1_label.split("(")[0].strip() + " Δ",
                        font_name=font)
        add_delta_table(slide, d2, left=delta2_left, top=chart_top, width=DELTA_COL_WIDTH,
                        row_height=row_h * ROW_SCALE_FACTOR,
                        header_text=s2_label.split("(")[0].strip() + " Δ",
                        font_name=font)
    else:
        # Gap column (difference between series)
        gaps = [round((c1 or 0) - (c2 or 0), 1) for c1, c2 in zip(s1_current, s2_current)]
        header = extra.get("gap_header", "Gap (pp)")
        add_delta_table(slide, gaps, left=delta1_left, top=chart_top, width=DELTA_COL_WIDTH,
                        row_height=row_h * ROW_SCALE_FACTOR, header_text=header,
                        font_name=font)

    # Manual legend below chart
    _make_legend(slide, [
        (s1_color, s1_label),
        (s2_color, s2_label),
    ], 0, chart_top + chart_h + LEGEND_GAP, font,
        center_over=(chart_left, block_w))


def render_stacked_order(slide, config: ProjectConfig, ask: AskConfig, data: dict, *, namer=None):
    """Stacked bar with ordinal breakdown (1st/2nd/3rd/4th recalled)."""
    _slide_chrome(slide, config, ask)

    rows = data.get(ask.data_key, [])
    rows = _sort_data(rows, ask.sort_by, ask.sort_desc)
    if not rows:
        textbox(slide, "Data not available", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    color_current, _ = _get_brand_colors(config, ask)
    font = config.font_body

    # Take top 10
    order_top = rows[:min(10, len(rows))]

    labels = [r.get("desc", "")[:LABEL_MAX_STACKED] for r in order_top]

    # Determine ordinals from data keys
    ordinals = ask.extra.get("ordinals", ["1st", "2nd", "3rd", "4th"])
    ordinal_vals = []
    for o in ordinals:
        ordinal_vals.append([r.get(f"{o}_current", 0) for r in order_top])
    total_vals = [r.get("total_current", 0) for r in order_top]

    n = len(labels)
    chart_top = CHART_TOP_STD
    chart_h = max(MIN_CHART_HEIGHT, min(MAX_CHART_HEIGHT, n * 0.42))
    row_h = chart_h / max(n, 1)

    # Center the chart + value col + delta col block
    value_col_w = 0.70
    block_w = CLUSTERED_BAR_WIDTH + CHART_DELTA_GAP + value_col_w + CHART_DELTA_GAP + DELTA_COL_WIDTH
    chart_left = (SLIDE_W - block_w) / 2
    value_left = chart_left + CLUSTERED_BAR_WIDTH + CHART_DELTA_GAP
    delta_left = value_left + value_col_w + CHART_DELTA_GAP

    # Build stacked bar chart
    cd = CategoryChartData()
    cd.categories = labels
    ordinal_labels = [f"{o} Recalled" for o in ordinals]
    if len(ordinal_labels) == 4:
        ordinal_labels[3] = f"{ordinals[3]}+ Recalled"
    for label, vals in zip(ordinal_labels, ordinal_vals):
        cd.add_series(label, vals)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_STACKED,
        Inches(chart_left), Inches(chart_top), Inches(CLUSTERED_BAR_WIDTH), Inches(chart_h), cd)
    ch = cf.chart
    ch.has_legend = False

    # Color scheme: gradient from dark to light based on brand color
    r, g, b = color_current[0], color_current[1], color_current[2]
    stack_colors = [
        RGBColor(max(0, r - 60), max(0, g - 30), max(0, b - 50)),  # darkest
        color_current,                                                # medium
        RGBColor(min(255, r + 50), min(255, g + 60), min(255, b + 50)),  # light
        RGBColor(min(255, r + 100), min(255, g + 120), min(255, b + 100)),  # palest
    ]
    label_colors = [C_WHITE, C_WHITE,
                    RGBColor(0x30, 0x10, 0x50), RGBColor(0x30, 0x10, 0x50)]

    for idx, series in enumerate(ch.series):
        if idx < len(stack_colors):
            set_series_color(series, stack_colors[idx])
        set_series_no_border(series)
        enable_data_labels(series, label_colors[idx] if idx < len(label_colors) else C_GREY,
                          fsize=6, pos="ctr", font_name=font)
        # Hide labels on small segments
        if idx < len(ordinal_vals):
            for pt_idx, val in enumerate(ordinal_vals[idx]):
                if val <= STACKED_HIDE_THRESHOLD:
                    delete_data_label(series, pt_idx)

    hide_axis(ch, "val")
    ch.category_axis.tick_labels.font.size = Pt(7)
    ch.category_axis.tick_labels.font.name = font
    suppress_cat_axis_bullets(ch)
    invert_cat_axis(ch)
    set_plot_area_gap(ch, BAR_GAP_STD)
    set_chart_plot_area(ch, x=0.0, y=0.0, w=1.0, h=1.0)

    # Total column
    add_value_table(
        slide, total_vals,
        left=value_left, top=chart_top, width=value_col_w, row_height=row_h * ROW_SCALE_FACTOR,
        header_text="Total %", value_color=color_current, font_name=font,
    )

    # QoQ delta column
    total_prior = [r.get("total_prior", 0) for r in order_top]
    qoq_deltas = [delta(c, p) for c, p in zip(total_vals, total_prior)]
    add_delta_table(
        slide, qoq_deltas,
        left=delta_left, top=chart_top, width=DELTA_COL_WIDTH, row_height=row_h * ROW_SCALE_FACTOR,
        header_text="QoQ Δ", font_name=font,
    )

    # Manual legend below chart
    ordinal_labels = [f"{o} Recalled" for o in ordinals]
    if len(ordinal_labels) == 4:
        ordinal_labels[3] = f"{ordinals[3]}+ Recalled"
    _make_legend(slide,
                 [(stack_colors[i], ordinal_labels[i]) for i in range(len(ordinal_labels))],
                 0, chart_top + chart_h + LEGEND_GAP, font,
                 center_over=(chart_left, block_w))


def render_dual_bar_compare(slide, config: ProjectConfig, ask: AskConfig, data: dict, *, namer=None):
    """Side-by-side dual bar comparison: [category col] [left bar+delta] | [right bar+delta].

    Generic renderer for any two-group horizontal bar comparison with a shared
    category column and vertical separator. Both charts are sized equally.
    Template-matched to template.pptx slide 23.

    Required extra keys:
      left.data_key   — extraction id for left series
      left.label      — header label above left chart
      right.data_key  — extraction id for right series
      right.label     — header label above right chart

    Optional extra keys:
      left.brand      — brand key for color (default: "primary")
      left.color      — explicit hex color override (e.g. "#F75824"), overrides brand
      right.brand     — brand key for color (default: "competitor")
      right.color     — explicit hex color override, overrides brand
      category_header  — small label above the category column
      axis_label       — axis sub-label below both charts (e.g. "% of Interactions")
      highlight_rows   — list of 0-indexed row numbers to shade (e.g. [1, 3])
    """
    _slide_chrome(slide, config, ask)

    extra = ask.extra or {}
    font = config.font_body
    left_cfg = extra.get("left", {})
    right_cfg = extra.get("right", {})

    # Data
    left_rows = data.get(left_cfg.get("data_key", ""), [])
    right_rows = data.get(right_cfg.get("data_key", ""), [])
    if not left_rows:
        textbox(slide, "Data not available", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    # Use full desc for matching; truncate only for chart display
    left_descs = [r.get("desc", "") for r in left_rows]
    labels = [d[:LABEL_MAX_DUAL] for d in left_descs]
    n = len(labels)

    # Colors — explicit hex override takes priority, then brand, then defaults
    left_brand  = config.brands.get(left_cfg.get("brand", "primary"),    config.primary)
    right_brand = config.brands.get(right_cfg.get("brand", "competitor"), config.competitor)
    left_color  = parse_color(left_cfg["color"])  if "color" in left_cfg  else left_brand.color_current
    right_color = parse_color(right_cfg["color"]) if "color" in right_cfg else right_brand.color_current

    # Values — match right rows to left order by full desc
    left_current  = [r.get("current") or 0 for r in left_rows]
    left_prior    = [r.get("prior") for r in left_rows]
    right_idx     = {r.get("desc"): r for r in right_rows}
    right_current = [right_idx.get(d, {}).get("current") or 0 for d in left_descs]
    right_prior   = [right_idx.get(d, {}).get("prior") for d in left_descs]

    # Shared axis scale
    all_vals = left_current + right_current
    axis_max = min(100, ((int(max(all_vals or [100])) // 10) + 1) * 10)

    left_deltas  = [delta(c, p) if p is not None else None for c, p in zip(left_current, left_prior)]
    right_deltas = [delta(c, p) if p is not None else None for c, p in zip(right_current, right_prior)]

    # Row height: use template value for ≤4 rows, scale down for more
    row_h    = min(DUAL_BC_ROW_H, 4.0 / max(n, 1))
    chart_h  = n * row_h
    chart_top = DUAL_BC_TOP
    label_top = chart_top - 0.44   # brand header labels just above chart area

    # Equal-width charts — both sides get the same width (min of available space)
    _GAP = 0.05
    _SLIDE_R = 13.10  # usable right edge
    _left_avail  = DUAL_BC_SEP_X - DUAL_BC_L_CHART_L - DUAL_BC_L_DELTA_W - _GAP
    _right_avail = _SLIDE_R - DUAL_BC_R_CHART_L - DUAL_BC_R_DELTA_W - _GAP
    chart_w      = min(_left_avail, _right_avail)
    left_delta_l  = DUAL_BC_L_CHART_L + chart_w + _GAP
    right_delta_l = DUAL_BC_R_CHART_L + chart_w + _GAP

    # 1. Optional highlight row bands (drawn first, behind everything else)
    for ri in extra.get("highlight_rows", []):
        if 0 <= ri < n:
            hy = chart_top + ri * row_h
            solidrect(slide,
                      DUAL_BC_CAT_LEFT, hy,
                      right_delta_l + DUAL_BC_R_DELTA_W - DUAL_BC_CAT_LEFT, row_h,
                      fill=C_LBGREY)

    # 2. Category column label (small text above the category table)
    cat_header = extra.get("category_header", "")
    if cat_header:
        textbox(slide, cat_header,
                DUAL_BC_CAT_LEFT, chart_top - 0.30, DUAL_BC_CAT_W, 0.28,
                fsize=8, bold=True, color=C_GREY, font=config.font_display)

    # 3. Category pptx table (no header row — labels align directly to row midpoints)
    _, cat_tbl = _pptx_table(slide, [DUAL_BC_CAT_W], [row_h] * n,
                              DUAL_BC_CAT_LEFT, chart_top)
    for i, label in enumerate(labels):
        cell = cat_tbl.cell(i, 0)
        _style_tbl_cell(cell, label, bg=None, fg=C_GREY,
                        fsize=7.5, align=PP_ALIGN.LEFT, font=font, ml=0.08, mr=0.04)
        if i < n - 1:
            _cell_bottom_border(cell, "D9D9D9")

    # 4. Column header labels (floating textboxes above each chart)
    left_label  = _resolve_template(left_cfg.get("label",  left_brand.name),  config)
    right_label = _resolve_template(right_cfg.get("label", right_brand.name), config)
    textbox(slide, left_label,
            DUAL_BC_L_CHART_L, label_top, chart_w + DUAL_BC_L_DELTA_W, 0.40,
            fsize=8, bold=True, color=left_color, align=PP_ALIGN.CENTER, font=font)
    textbox(slide, right_label,
            DUAL_BC_R_CHART_L, label_top, chart_w + DUAL_BC_R_DELTA_W, 0.40,
            fsize=8, bold=True, color=right_color, align=PP_ALIGN.CENTER, font=font)

    # 5. Left bar chart (no category labels — category table provides them)
    cf1, ch1 = add_single_bar_chart(
        slide, labels, left_current,
        left=DUAL_BC_L_CHART_L, top=chart_top,
        width=chart_w, height=chart_h,
        fill_color=left_color, font_name=font,
    )
    ch1.has_title = False
    hide_cat_labels(ch1)
    set_val_axis_scale(ch1, 0, axis_max)
    set_chart_plot_area(ch1, x=0.0, y=0.0, w=1.0, h=1.0)
    # Override to inEnd so labels sit inside bars and don't overflow into the delta column
    enable_data_labels(ch1.series[0], C_WHITE, pos="inEnd", font_name=font)

    # 6. Left brand delta column (headerless — brand label serves as header)
    left_delta_header = left_cfg.get("delta_header", "QoQ Δ")
    add_delta_table(
        slide, left_deltas,
        left=left_delta_l, top=chart_top,
        width=DUAL_BC_L_DELTA_W, row_height=row_h,
        header_text=left_delta_header, font_name=font,
        show_header=False,
    )

    # 7. Vertical separator
    dashed_separator(slide, DUAL_BC_SEP_X, label_top,
                     chart_h + 0.44, color=C_FTGREY, width_pt=0.5, vertical=True)

    # 8. Right bar chart
    cd2 = CategoryChartData()
    cd2.categories = labels
    cd2.add_series(config.period_current, right_current)
    cf2 = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(DUAL_BC_R_CHART_L), Inches(chart_top),
        Inches(chart_w), Inches(chart_h), cd2)
    ch2 = cf2.chart
    ch2.has_legend = False
    ch2.has_title = False
    s2 = ch2.series[0]
    set_series_color(s2, right_color)
    set_series_no_border(s2)
    enable_data_labels(s2, C_WHITE, pos="inEnd", font_name=font)
    hide_axis(ch2, "val")
    hide_cat_labels(ch2)
    invert_cat_axis(ch2)
    set_plot_area_gap(ch2, BAR_GAP_STD)
    set_val_axis_scale(ch2, 0, axis_max)
    set_chart_plot_area(ch2, x=0.02, y=0.0, w=0.98, h=1.0)

    # 9. Right brand delta column
    right_delta_header = right_cfg.get("delta_header", "QoQ Δ")
    add_delta_table(
        slide, right_deltas,
        left=right_delta_l, top=chart_top,
        width=DUAL_BC_R_DELTA_W, row_height=row_h,
        header_text=right_delta_header, font_name=font,
        show_header=False,
    )

    # 10. Axis sub-labels below both charts
    axis_label = extra.get("axis_label", "")
    ax_y = chart_top + chart_h + 0.04
    if axis_label:
        textbox(slide, axis_label,
                DUAL_BC_L_CHART_L, ax_y,
                chart_w + DUAL_BC_L_DELTA_W, 0.20,
                fsize=7, color=C_FTGREY, align=PP_ALIGN.CENTER, font=font)
        textbox(slide, axis_label,
                DUAL_BC_R_CHART_L, ax_y,
                chart_w + DUAL_BC_R_DELTA_W, 0.20,
                fsize=7, color=C_FTGREY, align=PP_ALIGN.CENTER, font=font)

    # 11. Legend — use configured labels (not brand name) for generality
    ly = ax_y + (0.24 if axis_label else LEGEND_GAP)
    _make_legend(slide, [
        (left_color,  left_label),
        (right_color, right_label),
        (C_GREEN, "Positive Δ"),
        (C_RED,   "Negative Δ"),
    ], 0, ly, font,
        center_over=(DUAL_BC_CAT_LEFT, right_delta_l + DUAL_BC_R_DELTA_W - DUAL_BC_CAT_LEFT))


# Backward-compat alias
render_dual_brand_compare = render_dual_bar_compare
