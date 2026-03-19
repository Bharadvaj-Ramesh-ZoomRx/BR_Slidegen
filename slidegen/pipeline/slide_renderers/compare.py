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
    CHART_TOP_STD, FOOTER_TOP, LEGEND_GAP, SLIDE_W, CHART_DELTA_GAP,
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
    _pptx_table, _style_tbl_cell, _cell_bottom_border, _cap_chart_h, _auto_label_width,
    # pptx_utils
    C_GREEN, C_WHITE, C_GREY, C_FTGREY, C_LBGREY, C_HDRGREY, C_RED,
    PP_ALIGN,
    textbox, solidrect, dashed_separator, callout_box,
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

    labels = [r.get("desc", "") for r in rows]

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

    # Layout: [Label table] [Clustered bar (no cat labels)] [Delta col(s)]
    gap = 0.08
    hdr_h = 0.30
    chart_top = CHART_TOP_STD + 0.05
    max_body_h = FOOTER_TOP - chart_top - 0.55
    row_h = min(0.42, max(0.28, max_body_h / max(n, 1)))
    body_h = n * row_h

    # Delta columns
    has_prior_1 = any(p is not None for p in s1_prior)
    has_prior_2 = any(p is not None for p in s2_prior)
    n_delta_cols = 2 if (has_prior_1 and has_prior_2) else 1
    delta_total_w = n_delta_cols * DELTA_COL_WIDTH + (n_delta_cols - 1) * gap

    # Dynamic label width + chart fills remaining
    label_w = _auto_label_width(labels)
    chart_w = SLIDE_W - 0.60 - label_w - delta_total_w - gap * 2

    # Center
    block_w = label_w + gap + chart_w + gap + delta_total_w
    origin = (SLIDE_W - block_w) / 2
    label_l = origin
    chart_l = label_l + label_w + gap
    delta1_l = chart_l + chart_w + gap
    delta2_l = delta1_l + DELTA_COL_WIDTH + gap

    # 1. Label table
    _, ltbl = _pptx_table(slide, [label_w], [hdr_h] + [row_h] * n,
                           label_l, chart_top)
    _style_tbl_cell(ltbl.cell(0, 0), "Attribute", bg=C_HDRGREY, fg=C_WHITE,
                    fsize=8, bold=True, align=PP_ALIGN.LEFT, font=font, ml=0.08, mr=0.05)
    for i, label in enumerate(labels):
        cell = ltbl.cell(i + 1, 0)
        _style_tbl_cell(cell, label,
                        bg=C_LBGREY if i % 2 == 0 else C_WHITE,
                        fg=C_GREY, fsize=7.5, align=PP_ALIGN.LEFT, font=font,
                        ml=0.08, mr=0.05)
        cell.text_frame.word_wrap = True

    # 2. Clustered bar (no category labels)
    cd = CategoryChartData()
    cd.categories = [f"R{i}" for i in range(n)]
    cd.add_series(s1_label, s1_current)
    cd.add_series(s2_label, s2_current)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(chart_l), Inches(chart_top + hdr_h), Inches(chart_w), Inches(body_h), cd)
    ch = cf.chart
    ch.has_legend = False

    set_series_color(ch.series[0], s1_color)
    set_series_no_border(ch.series[0])
    enable_data_labels(ch.series[0], s1_color, fsize=7, font_name=font)

    set_series_color(ch.series[1], s2_color)
    set_series_no_border(ch.series[1])
    enable_data_labels(ch.series[1], s2_color, fsize=7, font_name=font)

    hide_axis(ch, "val")
    hide_cat_labels(ch)
    invert_cat_axis(ch)
    set_plot_area_gap(ch, CLUSTERED_GAP)
    set_overlap(ch, -15)
    set_chart_plot_area(ch, x=0.0, y=0.0, w=1.0, h=1.0)

    # 3. Delta columns
    if has_prior_1 and has_prior_2:
        d1 = [delta(c, p) if p is not None else None for c, p in zip(s1_current, s1_prior)]
        d2 = [delta(c, p) if p is not None else None for c, p in zip(s2_current, s2_prior)]
        add_delta_table(slide, d1, left=delta1_l, top=chart_top, width=DELTA_COL_WIDTH,
                        row_height=row_h,
                        header_text=s1_label.split("(")[0].strip() + " Δ",
                        font_name=font)
        add_delta_table(slide, d2, left=delta2_l, top=chart_top, width=DELTA_COL_WIDTH,
                        row_height=row_h,
                        header_text=s2_label.split("(")[0].strip() + " Δ",
                        font_name=font)
    else:
        gaps = [round((c1 or 0) - (c2 or 0), 1) for c1, c2 in zip(s1_current, s2_current)]
        header = extra.get("gap_header", "Gap (pp)")
        add_delta_table(slide, gaps, left=delta1_l, top=chart_top, width=DELTA_COL_WIDTH,
                        row_height=row_h, header_text=header,
                        font_name=font)

    # 4. Legend
    ly = chart_top + hdr_h + body_h + LEGEND_GAP
    _make_legend(slide, [
        (s1_color, s1_label),
        (s2_color, s2_label),
    ], 0, ly, font,
        center_over=(label_l, block_w))


def render_stacked_order(slide, config: ProjectConfig, ask: AskConfig, data: dict, *, namer=None):
    """Stacked bar with ordinal breakdown (1st/2nd/3rd/4th recalled).

    Layout: [Label table] [Stacked bar chart (no cat labels)] [Total col] [Delta col]
    Label table shows full message text with proper font sizing.
    """
    _slide_chrome(slide, config, ask)

    rows = data.get(ask.data_key, [])
    rows = _sort_data(rows, ask.sort_by, ask.sort_desc)
    if not rows:
        textbox(slide, "Data not available", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    color_current, _ = _get_brand_colors(config, ask)
    font = config.font_body
    extra = ask.extra or {}

    # Take top 10
    order_top = rows[:min(10, len(rows))]
    labels = [r.get("desc", "") for r in order_top]

    # Determine ordinals from data keys
    ordinals = extra.get("ordinals", ["1st", "2nd", "3rd", "4th"])
    ordinal_vals = []
    for o in ordinals:
        ordinal_vals.append([r.get(f"{o}_current", 0) for r in order_top])
    total_vals = [r.get("total_current", 0) for r in order_top]

    n = len(labels)

    # Layout constants — table-based approach
    total_w = 0.65           # total % column
    delta_w = 0.60           # QoQ delta column
    gap = 0.08               # gap between elements
    label_w = _auto_label_width(labels)
    chart_w = SLIDE_W - 0.60 - label_w - total_w - delta_w - gap * 3  # fill remaining
    hdr_h = 0.30             # header row height
    chart_top = CHART_TOP_STD + 0.05
    max_body_h = FOOTER_TOP - chart_top - 0.60  # leave room for legend + footer
    row_h = min(0.42, max(0.30, max_body_h / max(n, 1)))
    body_h = n * row_h

    # Center the full block
    total_block_w = label_w + gap + chart_w + gap + total_w + gap + delta_w
    origin = (SLIDE_W - total_block_w) / 2
    label_l = origin
    chart_l = label_l + label_w + gap
    total_l = chart_l + chart_w + gap
    delta_l = total_l + total_w + gap

    # 1. Label table — full message text, 9pt, left-aligned
    _, ltbl = _pptx_table(slide, [label_w], [hdr_h] + [row_h] * n,
                           label_l, chart_top)
    # Header
    _style_tbl_cell(ltbl.cell(0, 0), "Message", bg=C_HDRGREY, fg=C_WHITE,
                    fsize=8, bold=True, align=PP_ALIGN.LEFT, font=font, ml=0.08, mr=0.05)
    for i, label in enumerate(labels):
        cell = ltbl.cell(i + 1, 0)
        _style_tbl_cell(cell, label,
                        bg=C_LBGREY if i % 2 == 0 else C_WHITE,
                        fg=C_GREY, fsize=7.5, align=PP_ALIGN.LEFT, font=font,
                        ml=0.08, mr=0.05)
        cell.text_frame.word_wrap = True

    # 2. Stacked bar chart — no category labels (table provides them)
    cd = CategoryChartData()
    cd.categories = [f"R{i}" for i in range(n)]  # dummy labels (hidden)
    ordinal_labels = [f"{o} Recalled" for o in ordinals]
    if len(ordinal_labels) == 4:
        ordinal_labels[3] = f"{ordinals[3]}+ Recalled"
    for label, vals in zip(ordinal_labels, ordinal_vals):
        cd.add_series(label, vals)

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_STACKED,
        Inches(chart_l), Inches(chart_top + hdr_h),
        Inches(chart_w), Inches(body_h), cd)
    ch = cf.chart
    ch.has_legend = False
    ch.has_title = False

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
                          fsize=7, pos="ctr", font_name=font)
        # Hide labels on small segments
        if idx < len(ordinal_vals):
            for pt_idx, val in enumerate(ordinal_vals[idx]):
                if val <= STACKED_HIDE_THRESHOLD:
                    delete_data_label(series, pt_idx)

    hide_axis(ch, "val")
    hide_cat_labels(ch)
    invert_cat_axis(ch)
    set_plot_area_gap(ch, 60)
    set_chart_plot_area(ch, x=0.0, y=0.0, w=1.0, h=1.0)

    # 3. Total column
    add_value_table(
        slide, total_vals,
        left=total_l, top=chart_top, width=total_w, row_height=row_h,
        header_text="Total %", value_color=color_current, font_name=font,
    )

    # 4. QoQ delta column
    total_prior = [r.get("total_prior", 0) for r in order_top]
    qoq_deltas = [delta(c, p) for c, p in zip(total_vals, total_prior)]
    add_delta_table(
        slide, qoq_deltas,
        left=delta_l, top=chart_top, width=delta_w, row_height=row_h,
        header_text="QoQ Δ", font_name=font,
    )

    # 5. Manual legend below
    ly = chart_top + hdr_h + body_h + LEGEND_GAP
    _make_legend(slide,
                 [(stack_colors[i], ordinal_labels[i]) for i in range(len(ordinal_labels))],
                 0, ly, font,
                 center_over=(label_l, total_block_w))


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

    # Use full desc for matching and display
    left_descs = [r.get("desc", "") for r in left_rows]
    labels = left_descs
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
    # Budget for elements below chart: axis(0.20+0.24) + legend(0.20+0.10) + callout(0.45+0.10) = 1.29
    has_insight = bool(extra.get("insight_text"))
    below_budget = 0.44 + (0.65 if has_insight else 0)  # axis+legend + optional callout
    max_chart_h = FOOTER_TOP - DUAL_BC_TOP - below_budget
    row_h    = min(DUAL_BC_ROW_H, max_chart_h / max(n, 1))
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
    legend_items = [
        (left_color,  left_label),
        (right_color, right_label),
        (C_GREEN, "Positive Δ"),
        (C_RED,   "Negative Δ"),
    ]
    _make_legend(slide, legend_items, 0, ly, font,
        center_over=(DUAL_BC_CAT_LEFT, right_delta_l + DUAL_BC_R_DELTA_W - DUAL_BC_CAT_LEFT))

    # 12. Insight callout box (optional — dashed border box below the legend)
    insight_text = extra.get("insight_text", "")
    if insight_text:
        insight_y = ly + 0.30
        insight_text = _resolve_template(insight_text, config)
        box_w = right_delta_l + DUAL_BC_R_DELTA_W - DUAL_BC_CAT_LEFT
        callout_box(slide, DUAL_BC_CAT_LEFT, insight_y, box_w, 0.45,
                    text=insight_text, border_color=C_GREY, dashed=True,
                    fsize=9, text_color=C_GREY)


# Backward-compat alias
render_dual_brand_compare = render_dual_bar_compare


# ══════════════════════════════════════════════════════════════════════════════
# HII Scorecard — multi-section clustered column chart (template slide 24)
# ══════════════════════════════════════════════════════════════════════════════

_SC_TOP       = 1.88    # chart area top
_SC_CHART_TOP = 3.28    # chart starts below section headers
_SC_CHART_H   = 2.58    # chart height
_SC_TBL_TOP   = 5.56    # label table top (below chart)
_SC_TBL_H     = 0.66    # label table height
_SC_LEFT      = 0.45    # left edge
_SC_RIGHT     = 12.85   # right edge
_SC_SEC_HDR_H = 0.70    # section header block height
_SC_SEC_HDR_TOP = 2.50  # section header Y


def render_hii_scorecard(slide, config: ProjectConfig, ask: AskConfig, data: dict, *, namer=None):
    """HII Drivers Scorecard — multi-section clustered column chart.

    Mirrors template slide 24: sections separated by vertical lines, each with
    a header group, clustered columns (HII vs Other), a label table below, and
    an optional insight callout.

    Config (ask.extra):
        sections: list of section defs, each with:
            label: section header text (e.g. "Visual Aid Types")
            subtitle: optional subtitle (e.g. "(% of Interactions)")
            summary: optional summary text shown in header box (e.g. "Avg: 3.7 | 3.1")
            items: list of category names to include from the data
        series: [
            { field: "hi", label: "High Impact (n=63)", color: "#F75824" },
            { field: "other", label: "Other (n=37)", color: "#C0C0C0" },
        ]
        insight_text: optional callout below chart
    """
    _slide_chrome(slide, config, ask)

    extra = ask.extra or {}
    font  = config.font_body
    sections = extra.get("sections", [])
    series_cfg = extra.get("series", [])

    rows = data.get(ask.data_key, [])
    if not rows:
        textbox(slide, "Data not available", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    # Build lookup: desc → row
    row_map = {r.get("desc", ""): r for r in rows}

    # Flatten sections into ordered categories
    all_cats = []
    section_ranges = []  # (start_idx, end_idx, section_cfg)
    for sec in sections:
        start = len(all_cats)
        for item in sec.get("items", []):
            all_cats.append(item)
        section_ranges.append((start, len(all_cats), sec))

    n_cats = len(all_cats)
    if n_cats == 0:
        textbox(slide, "No categories defined", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    # Extract series values
    series_data = []
    for s in series_cfg:
        field = s.get("field", "current")
        vals = []
        for cat in all_cats:
            r = row_map.get(cat, {})
            vals.append(r.get(f"{field}_current") or r.get(field) or 0)
        color = parse_color(s["color"]) if "color" in s else config.primary.color_current
        series_data.append({"label": s.get("label", field), "vals": vals, "color": color})

    # Chart dimensions — full width
    chart_w = _SC_RIGHT - _SC_LEFT
    chart_left = _SC_LEFT

    # Section divider positions
    cat_width = chart_w / n_cats
    sec_x_positions = []
    for start, end, sec in section_ranges:
        x_start = chart_left + start * cat_width
        x_end = chart_left + end * cat_width
        sec_x_positions.append((x_start, x_end, sec))

    # 1. Section header boxes and labels
    for x_start, x_end, sec in sec_x_positions:
        sec_w = x_end - x_start
        label = sec.get("label", "")
        subtitle = sec.get("subtitle", "")
        summary = sec.get("summary", "")

        # Header text
        header_text = label
        if subtitle:
            header_text += f"\n{subtitle}"

        textbox(slide, header_text,
                x_start + 0.05, _SC_SEC_HDR_TOP, sec_w - 0.10, _SC_SEC_HDR_H,
                fsize=7, color=C_GREY, align=PP_ALIGN.CENTER, font=font)

        # Summary box (if provided) — above the section header text
        if summary:
            summary = _resolve_template(summary, config)
            box_h = 0.48
            box_w = min(1.5, sec_w - 0.10)
            box_l = x_start + (sec_w - box_w) / 2
            box_top = _SC_SEC_HDR_TOP - box_h - 0.06  # above header text with gap
            callout_box(slide, box_l, box_top, box_w, box_h,
                        text=summary, border_color=C_FTGREY, dashed=True,
                        fsize=7, text_color=C_GREY)

    # 2. Vertical section dividers
    for i, (x_start, x_end, sec) in enumerate(sec_x_positions):
        if i > 0:
            dashed_separator(slide, x_start, _SC_TOP, _SC_TBL_TOP + _SC_TBL_H - _SC_TOP,
                             color=C_FTGREY, width_pt=0.75, vertical=True)

    # 3. Clustered column chart
    cd = CategoryChartData()
    cd.categories = [c[:20] for c in all_cats]
    for sd in series_data:
        # Values as decimals (0-1) for percentage display
        cd.add_series(sd["label"], [v / 100.0 for v in sd["vals"]])

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(chart_left), Inches(_SC_CHART_TOP),
        Inches(chart_w), Inches(_SC_CHART_H), cd)
    ch = cf.chart
    ch.has_legend = False
    ch.has_title = False

    # Style series + data labels
    for i, sd in enumerate(series_data):
        s = ch.series[i]
        set_series_color(s, sd["color"])
        set_series_no_border(s)
        enable_data_labels(s, C_GREY, pos="outEnd", font_name=font, num_fmt='0%', fsize=7)

    # Axis styling
    hide_cat_labels(ch)
    val_ax = ch.value_axis
    val_ax.maximum_scale = 0.80
    val_ax.minimum_scale = 0.0
    val_ax.has_title = False
    val_ax.format.line.fill.background()
    val_ax.major_gridlines.format.line.color.rgb = C_LBGREY
    val_ax.major_gridlines.format.line.width = Pt(0.5)
    # Number format as percentage
    val_ax.tick_labels.number_format = '0%'
    val_ax.tick_labels.font.size = Pt(7)
    val_ax.tick_labels.font.color.rgb = C_FTGREY

    set_plot_area_gap(ch, 80)
    set_overlap(ch, 0)
    set_chart_plot_area(ch, x=0.02, y=0.0, w=0.96, h=0.95)

    # 4. Category label table below chart
    n_cols = n_cats
    col_w = chart_w / n_cols
    _, tbl = _pptx_table(slide, [col_w] * n_cols, [_SC_TBL_H],
                          chart_left, _SC_TBL_TOP)
    for ci, cat in enumerate(all_cats):
        short = cat[:25].replace(" - ", "\n").replace(" / ", "\n")
        cell = tbl.cell(0, ci)
        _style_tbl_cell(cell, short, bg=C_LBGREY, fg=C_GREY,
                        fsize=6.5, align=PP_ALIGN.CENTER, font=font)

    # 5. Legend
    ly = _SC_TBL_TOP + _SC_TBL_H + 0.08
    legend_items = [(sd["color"], _resolve_template(sd["label"], config)) for sd in series_data]
    _make_legend(slide, legend_items, 0, ly, font,
                 center_over=(chart_left, chart_w))

    # 6. Insight callout (optional)
    insight_text = extra.get("insight_text", "")
    if insight_text:
        insight_y = ly + 0.24
        insight_text = _resolve_template(insight_text, config)
        callout_box(slide, chart_left, insight_y, chart_w, 0.40,
                    text=insight_text, border_color=C_GREY, dashed=True,
                    fsize=8, text_color=C_GREY)


# ══════════════════════════════════════════════════════════════════════════════
# Dual Doughnut — side-by-side doughnut pairs (template slide 50)
# ══════════════════════════════════════════════════════════════════════════════

def render_dual_doughnut(slide, config: ProjectConfig, ask: AskConfig, data: dict, *, namer=None):
    """Dual doughnut comparison: two patient segments, each with two brand doughnuts.

    Layout: [Left section: Brand1 donut + Brand2 donut] | [Right section: Brand1 donut + Brand2 donut]

    Config (ask.extra):
        left:
            label: section header (e.g. "With CNS Metastasis")
            items: list of 2 dicts, each with:
                data_key, brand_label, current, prior (field names or values)
        right:
            (same as left)
        legend_current / legend_prior: legend labels
    """
    from pptx.chart.data import ChartData
    from lxml import etree as _etree
    from pptx.oxml.ns import qn as _qn

    _slide_chrome(slide, config, ask)

    extra = ask.extra or {}
    font  = config.font_body

    left_cfg  = extra.get("left", {})
    right_cfg = extra.get("right", {})

    # Doughnut dimensions
    donut_size = 2.20
    donut_gap  = 0.20  # gap between two donuts in a section
    section_gap = 0.50  # gap between left and right sections
    section_top = 3.10
    label_h = 0.28

    # Calculate total width and center
    section_w = donut_size * 2 + donut_gap
    total_w = section_w * 2 + section_gap
    origin = (13.33 - total_w) / 2

    l_x1 = origin
    l_x2 = l_x1 + donut_size + donut_gap
    sep_x = origin + section_w + section_gap / 2
    r_x1 = origin + section_w + section_gap
    r_x2 = r_x1 + donut_size + donut_gap

    # Section headers
    hdr_top = 1.82
    hdr_h = 0.45
    l_label = _resolve_template(left_cfg.get("label", "Left"), config)
    r_label = _resolve_template(right_cfg.get("label", "Right"), config)
    textbox(slide, l_label, l_x1, hdr_top, section_w, hdr_h,
            fsize=9, bold=True, color=C_GREY, font=font)
    textbox(slide, r_label, r_x1, hdr_top, section_w, hdr_h,
            fsize=9, bold=True, color=C_GREY, font=font)

    # Vertical separator
    dashed_separator(slide, sep_x, 2.77, 2.62, color=C_FTGREY, width_pt=0.75, vertical=True)

    def _add_doughnut(x, y, items_cfg, section_data):
        """Add a single doughnut chart at (x, y)."""
        current_val = section_data.get("current", 0)
        prior_val = section_data.get("prior")
        brand_label = _resolve_template(section_data.get("brand_label", ""), config)
        brand_color = parse_color(section_data["color"]) if "color" in section_data else config.primary.color_current
        prior_color = parse_color(section_data["color_prior"]) if "color_prior" in section_data else config.primary.color_prior

        # Brand label above donut
        textbox(slide, brand_label, x, y - 0.02, donut_size, label_h,
                fsize=8, bold=True, color=brand_color, align=PP_ALIGN.CENTER, font=font)

        # Build doughnut data: outer ring = prior, inner ring = current
        cd = ChartData()
        cd.categories = [brand_label, "Other"]

        if prior_val is not None:
            cd.add_series("Prior", (prior_val / 100.0, 1.0 - prior_val / 100.0))
        cd.add_series("Current", (current_val / 100.0, 1.0 - current_val / 100.0))

        cf = slide.shapes.add_chart(
            XL_CHART_TYPE.DOUGHNUT,
            Inches(x), Inches(y + label_h),
            Inches(donut_size), Inches(donut_size), cd)
        ch = cf.chart
        ch.has_legend = False
        ch.has_title = False

        # Color the series
        for i, ser in enumerate(ch.series):
            has_two = prior_val is not None
            if has_two and i == 0:
                # Prior (outer) — lighter color
                c = prior_color
            else:
                # Current (inner or only) — brand color
                c = brand_color

            # Color first point (value), make second point (remainder) light grey
            for pi in range(2):
                pt = ser.points[pi]
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = c if pi == 0 else C_LBGREY
                pt.format.line.fill.background()

        # Center label showing current %
        center_y = y + label_h + donut_size / 2 - 0.15
        textbox(slide, f"{current_val:.0f}%", x, center_y, donut_size, 0.30,
                fsize=16, bold=True, color=brand_color, align=PP_ALIGN.CENTER, font=font)

    # Load data for each section
    def _get_section_items(cfg):
        items = cfg.get("items", [])
        result = []
        for item in items:
            dk = item.get("data_key", ask.data_key)
            desc_match = item.get("desc", "")
            rows = data.get(dk, [])
            row = next((r for r in rows if r.get("desc", "").startswith(desc_match)), {})
            result.append({
                "current": row.get("current") or item.get("current", 0),
                "prior": row.get("prior") if row.get("prior") is not None else item.get("prior"),
                "brand_label": item.get("brand_label", ""),
                "color": item.get("color", "#F75824"),
                "color_prior": item.get("color_prior", "#FFC199"),
                "sample_current": item.get("sample_current"),
                "sample_prior": item.get("sample_prior"),
            })
        return result

    l_items = _get_section_items(left_cfg)
    r_items = _get_section_items(right_cfg)

    # Draw donuts
    if len(l_items) >= 1:
        _add_doughnut(l_x1, section_top, left_cfg, l_items[0])
    if len(l_items) >= 2:
        _add_doughnut(l_x2, section_top, left_cfg, l_items[1])
    if len(r_items) >= 1:
        _add_doughnut(r_x1, section_top, right_cfg, r_items[0])
    if len(r_items) >= 2:
        _add_doughnut(r_x2, section_top, right_cfg, r_items[1])

    # Legend — show current and prior for each brand
    ly = section_top + label_h + donut_size + 0.25
    legend_current = extra.get("legend_current", config.period_current)
    legend_prior = extra.get("legend_prior", config.period_prior)

    # Collect unique brands from left items (RYB first, then TAG)
    all_items = l_items + (r_items if r_items != l_items else [])
    # Deduplicate by brand_label
    seen = set()
    unique_brands = []
    for item in all_items:
        bl = item.get("brand_label", "")
        if bl not in seen:
            seen.add(bl)
            unique_brands.append(item)

    legend_items = []
    for item in unique_brands:
        bl = _resolve_template(item.get("brand_label", ""), config)
        c_cur = parse_color(item["color"])
        c_pri = parse_color(item["color_prior"])
        s_cur = item.get("sample_current")
        s_pri = item.get("sample_prior")
        cur_suffix = f" (s={s_cur})" if s_cur is not None else ""
        pri_suffix = f" (s={s_pri})" if s_pri is not None else ""
        legend_items.append((c_cur, f"{bl} {legend_current}{cur_suffix}"))
        legend_items.append((c_pri, f"{bl} {legend_prior}{pri_suffix}"))

    _make_legend(slide, legend_items, 0, ly, font, center_over=(0, 13.33))
