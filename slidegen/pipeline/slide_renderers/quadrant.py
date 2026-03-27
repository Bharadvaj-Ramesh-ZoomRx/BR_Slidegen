"""
quadrant.py — 2×2 quadrant scatter chart renderer.

Slide type:
  - quadrant_scatter   (template slide 30 — stated vs derived importance)
"""

from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from ._shared import (
    SLIDE_W, CHART_TOP_STD, FOOTER_TOP,
    C_WHITE, C_GREY, C_FTGREY, C_LBGREY, C_HDRGREY,
    FONT_TEXT,
    _resolve_template, _slide_chrome, _make_legend,
    textbox, solidrect, horiz_line, dashed_separator,
    add_scatter_chart, set_chart_plot_area,
    ProjectConfig, AskConfig, parse_color,
    logger,
)


def render_quadrant_scatter(slide, config: ProjectConfig, ask: AskConfig,
                             data: dict, *, namer=None):
    """Render a 2×2 quadrant scatter chart with labeled data points.

    Data format (mock rows): list of dicts, each a data point:
        {
            "label": "Efficacy",
            "x": 0.32,                 # stated importance (X axis)
            "y": 0.695,                # derived importance (Y axis)
            "color": "#0070C0"         # optional per-point color override
        }

    ask.extra keys:
        x_label: "Stated Importance Score"
        y_label: "Derived Importance"
        x_sub:   "(Share of points allocated across attributes)"
        y_sub:   "(Correlation with future usage)"
        x_min/x_max/y_min/y_max: axis scale (defaults: 0–0.40, 0.45–0.75)
        x_mid/y_mid: midpoint for quadrant dividers (defaults: midpoint of scale)
        quadrant_labels:
            tl: "Hidden drivers: ..."
            tr: "Critical drivers: ..."
            bl: "Non-essentials: ..."
            br: "Fundamentals: ..."
        marker_color: "#0070C0"       (default dot color)
        marker_size: 7
        average_point:                 (optional average crosshair)
            label: "SA"
            x: 0.126
            y: 0.611
            color: "#FF0000"
    """
    _slide_chrome(slide, config, ask)
    extra = ask.extra or {}

    rows = data.get(ask.data_key, [])
    if not rows:
        logger.warning("quadrant_scatter: no data for key=%s", ask.data_key)
        return

    # Layout
    chart_left = 1.35
    chart_top = CHART_TOP_STD + 0.15
    chart_w = 10.54
    chart_h = 4.19
    font = config.font_body or FONT_TEXT

    # Axis ranges
    x_min = extra.get("x_min", 0.0)
    x_max = extra.get("x_max", 0.40)
    y_min = extra.get("y_min", 0.45)
    y_max = extra.get("y_max", 0.75)
    x_mid = extra.get("x_mid", (x_min + x_max) / 2)
    y_mid = extra.get("y_mid", (y_min + y_max) / 2)

    marker_color = parse_color(extra.get("marker_color", "#0070C0"))
    marker_size = extra.get("marker_size", 7)

    # Build scatter series
    x_vals = [r.get("x", 0) for r in rows]
    y_vals = [r.get("y", 0) for r in rows]
    labels = [r.get("label", "") for r in rows]

    series_list = [("Data", x_vals, y_vals)]
    colors = [marker_color]

    # Add average point if specified
    avg = extra.get("average_point")
    if avg:
        series_list.append((avg.get("label", "Avg"),
                            [avg.get("x", 0)], [avg.get("y", 0)]))
        colors.append(parse_color(avg.get("color", "#FF0000")))

    # Draw quadrant background fills (light tints)
    q_colors = extra.get("quadrant_colors", {})
    tl_fill = parse_color(q_colors.get("tl", "#FFF2CC"))  # hidden drivers — light yellow
    tr_fill = parse_color(q_colors.get("tr", "#E2EFDA"))  # critical drivers — light green
    bl_fill = parse_color(q_colors.get("bl", "#F2F2F2"))  # non-essentials — light grey
    br_fill = parse_color(q_colors.get("br", "#DEEBF7"))  # fundamentals — light blue

    # Compute pixel positions for quadrant fills
    x_frac = (x_mid - x_min) / (x_max - x_min)  # fraction of chart width for midpoint
    y_frac = (y_mid - y_min) / (y_max - y_min)

    left_w = chart_w * x_frac
    right_w = chart_w - left_w
    top_h = chart_h * (1 - y_frac)  # chart Y is inverted (top = high Y)
    bot_h = chart_h - top_h

    solidrect(slide, chart_left, chart_top, left_w, top_h, tl_fill)         # TL
    solidrect(slide, chart_left + left_w, chart_top, right_w, top_h, tr_fill)  # TR
    solidrect(slide, chart_left, chart_top + top_h, left_w, bot_h, bl_fill)   # BL
    solidrect(slide, chart_left + left_w, chart_top + top_h, right_w, bot_h, br_fill)  # BR

    # Draw midpoint divider lines
    mid_color = C_GREY
    # Vertical divider at x_mid
    dashed_separator(slide, chart_left + left_w, chart_top, chart_h,
                     vertical=True, color=mid_color)
    # Horizontal divider at y_mid
    dashed_separator(slide, chart_left, chart_top + top_h, chart_w,
                     vertical=False, color=mid_color)

    # Add scatter chart on top (transparent background)
    cf, ch = add_scatter_chart(
        slide, series_list,
        left=chart_left, top=chart_top, width=chart_w, height=chart_h,
        colors=colors, marker_size=marker_size,
        x_min=x_min, x_max=x_max,
        y_min=y_min, y_max=y_max,
        font_name=font,
    )
    # Expand plot area to fill chart frame
    set_chart_plot_area(ch, x=0.02, y=0.04, w=0.96, h=0.92)

    # Add data point labels as textboxes (positioned relative to chart)
    for i, label in enumerate(labels):
        if not label:
            continue
        x_val = x_vals[i]
        y_val = y_vals[i]
        # Map data coords to slide inches
        x_frac_pt = (x_val - x_min) / (x_max - x_min)
        y_frac_pt = 1 - (y_val - y_min) / (y_max - y_min)  # invert Y
        lbl_x = chart_left + x_frac_pt * chart_w + 0.12  # offset right of dot
        lbl_y = chart_top + y_frac_pt * chart_h - 0.10
        # Clamp to chart bounds
        lbl_x = min(lbl_x, chart_left + chart_w - 1.2)
        lbl_y = max(lbl_y, chart_top)
        textbox(slide, label, lbl_x, lbl_y, 1.5, 0.22,
                fsize=7, bold=False, color=C_FTGREY, font=font)

    # Quadrant labels
    q_labels = extra.get("quadrant_labels", {})
    ql_fsize = 7.5
    ql_w = 3.2
    ql_h = 0.45
    if q_labels.get("tl"):
        textbox(slide, q_labels["tl"],
                chart_left + 0.15, chart_top + 0.05, ql_w, ql_h,
                fsize=ql_fsize, bold=False, color=C_FTGREY, font=font)
    if q_labels.get("tr"):
        textbox(slide, q_labels["tr"],
                chart_left + chart_w - ql_w - 0.15, chart_top + 0.05, ql_w, ql_h,
                fsize=ql_fsize, bold=False, color=C_FTGREY, font=font,
                align=PP_ALIGN.RIGHT)
    if q_labels.get("bl"):
        textbox(slide, q_labels["bl"],
                chart_left + 0.15, chart_top + chart_h - ql_h - 0.05, ql_w, ql_h,
                fsize=ql_fsize, bold=False, color=C_FTGREY, font=font)
    if q_labels.get("br"):
        textbox(slide, q_labels["br"],
                chart_left + chart_w - ql_w - 0.15, chart_top + chart_h - ql_h - 0.05,
                ql_w, ql_h, fsize=ql_fsize, bold=False, color=C_FTGREY, font=font,
                align=PP_ALIGN.RIGHT)

    # Axis labels — positioned to match template slide 30
    x_label = extra.get("x_label", "")
    y_label = extra.get("y_label", "")
    x_sub = extra.get("x_sub", "")
    y_sub = extra.get("y_sub", "")

    # X-axis label: centered below the chart area
    if x_label:
        x_lbl_w = 3.26
        x_lbl_left = chart_left + (chart_w - x_lbl_w) / 2
        x_lbl_top = chart_top + chart_h + 0.02
        shape = slide.shapes.add_textbox(
            Inches(x_lbl_left), Inches(x_lbl_top),
            Inches(x_lbl_w), Inches(0.38))
        tf = shape.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        run1 = p1.add_run()
        run1.text = x_label
        run1.font.size = Pt(8)
        run1.font.bold = True
        run1.font.color.rgb = C_FTGREY
        run1.font.name = font
        if x_sub:
            p2 = tf.add_paragraph()
            p2.alignment = PP_ALIGN.CENTER
            run2 = p2.add_run()
            run2.text = x_sub
            run2.font.size = Pt(7)
            run2.font.bold = False
            run2.font.color.rgb = C_GREY
            run2.font.name = font

    # Y-axis label: rotated 270° along the left edge of the chart (reads bottom-to-top)
    if y_label:
        from pptx.oxml.ns import qn as _qn
        y_lbl_w = 0.45        # narrow width (becomes height when rotated)
        y_lbl_h = chart_h * 0.5  # tall enough to span mid-chart
        y_lbl_left = chart_left - y_lbl_w - 0.08
        y_lbl_top = chart_top + (chart_h - y_lbl_h) / 2
        shape = slide.shapes.add_textbox(
            Inches(y_lbl_left), Inches(y_lbl_top),
            Inches(y_lbl_w), Inches(y_lbl_h))
        tf = shape.text_frame
        tf.word_wrap = True
        # Rotate text 270° (vertical, reading bottom-to-top)
        bodyPr = tf._txBody.find(_qn("a:bodyPr"))
        if bodyPr is not None:
            bodyPr.set("vert", "vert270")
            bodyPr.set("anchor", "ctr")
        p1 = tf.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        run1 = p1.add_run()
        run1.text = y_label
        run1.font.size = Pt(8)
        run1.font.bold = True
        run1.font.color.rgb = C_FTGREY
        run1.font.name = font
        if y_sub:
            p2 = tf.add_paragraph()
            p2.alignment = PP_ALIGN.CENTER
            run2 = p2.add_run()
            run2.text = y_sub
            run2.font.size = Pt(7)
            run2.font.bold = False
            run2.font.color.rgb = C_GREY
            run2.font.name = font

    # Legend
    legend_items = [(marker_color, "Data Points")]
    if avg:
        legend_items = [(marker_color, extra.get("legend_data", "Overall")),
                        (parse_color(avg.get("color", "#FF0000")),
                         avg.get("label", "Avg"))]
    legend_top = FOOTER_TOP - 0.18
    _make_legend(slide, legend_items, left=chart_left, top=legend_top,
                 font_name=font, center_over=(chart_left, chart_w))
