"""
compare_special.py — HII scorecard and dual doughnut slide renderers.

Split from compare.py for maintainability.
"""

from __future__ import annotations
import logging

from pptx.chart.data import CategoryChartData, ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from lxml import etree as _etree
from pptx.oxml.ns import qn as _qn

from ._shared import (
    _resolve_template, _slide_chrome, _get_brand_colors, _make_legend,
    _pptx_table, _style_tbl_cell,
    _no_data_placeholder,
    C_GREEN, C_WHITE, C_GREY, C_FTGREY, C_LBGREY, C_HDRGREY, C_RED,
    PP_ALIGN,
    textbox, solidrect, dashed_separator, callout_box,
    hide_axis, set_series_color, set_plot_area_gap, set_overlap,
    set_series_no_border, hide_cat_labels,
    set_chart_plot_area,
    enable_data_labels,
    ProjectConfig, AskConfig, parse_color,
)

logger = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════════════════════
# HII Scorecard — layout constants
# ══════════════════════════════════════════════════════════════════════════════

_SC_TOP       = 1.88
_SC_CHART_TOP = 3.28
_SC_CHART_H   = 2.58
_SC_TBL_TOP   = 5.56
_SC_TBL_H     = 0.66
_SC_LEFT      = 0.45
_SC_RIGHT     = 12.85
_SC_SEC_HDR_H = 0.70
_SC_SEC_HDR_TOP = 2.50


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
        _no_data_placeholder(slide, ask.id)
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
