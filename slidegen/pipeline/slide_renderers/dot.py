"""
dot.py — Lollipop, abacus, and message MBD dot-chart slide renderers.
"""

from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from pptx.chart.data import XyChartData
from pptx.enum.chart import XL_CHART_TYPE

from ._shared import (
    # constants
    CHART_TOP_STD, LEGEND_GAP, SLIDE_W,
    MAX_CHART_HEIGHT, MIN_CHART_HEIGHT,
    LOLLI_LABEL_W, LOLLI_PLOT_LEFT, LOLLI_PLOT_W, LOLLI_DOT_R, LOLLI_PRIOR_R,
    LABEL_MAX_SINGLE, LABEL_MAX_DUAL,
    _ABS_HDR_H, _ABS_ROW_H_MIN, _ABS_ROW_H_MAX, _ABS_TOP,
    _ABS_LABEL_W, _ABS_VAL_W, _ABS_CHART_W, _ABS_DELTA_W, _ABS_GAP, _ABS_SLIDE_W,
    # helpers
    _resolve_template, _slide_chrome, _get_brand_colors, _sort_data, _make_legend, _add_dot,
    _pptx_table, _style_tbl_cell, _prepare_rows,
    _alt_row_bg, _no_data_placeholder,
    # pptx_utils
    C_GREEN, C_GREY, C_FTGREY, C_LBGREY, C_HDRGREY, C_WHITE, C_RED,
    FONT_TEXT,
    textbox, solidrect, horiz_line, dashed_separator,
    set_series_marker, set_series_line_style,
    set_chart_plot_area, add_delta_table,
    _get_or_add, suppress_para_bullets, cell_vcenter,
    Inches, Pt,
    # project_config types
    ProjectConfig, AskConfig, parse_color,
)
from slidegen.pipeline.data_loaders import delta
from pptx.oxml.ns import qn as _qn


def render_abacus(slide, config: ProjectConfig, ask: AskConfig, data: dict, *, namer=None):
    """Abacus chart using XY scatter + tables layout (mirrors template slide 20).

    Layout (left → right, horizontally centered on slide):
      Attribute labels | Prior % | Current % | XY scatter dots | QoQ Δ
    """
    _slide_chrome(slide, config, ask)

    rows = _prepare_rows(slide, data, ask)
    if rows is None:
        return

    font = config.font_body
    extra = ask.extra or {}

    # Custom field names (default: "current" / "prior")
    current_field = extra.get("current_field", "current")
    prior_field   = extra.get("prior_field", "prior")

    # Filter out rows where current value is None or 0
    rows = [r for r in rows if r.get(current_field) not in (None, 0)]

    # Custom colors (default: brand colors)
    color_current = parse_color(extra["color_current"]) if "color_current" in extra else None
    color_prior   = parse_color(extra["color_prior"]) if "color_prior" in extra else None
    if color_current is None or color_prior is None:
        bc, bp = _get_brand_colors(config, ask)
        color_current = color_current or bc
        color_prior   = color_prior or bp

    labels       = [r.get("short", r.get("desc", ""))[:LABEL_MAX_SINGLE] for r in rows]
    current_vals = [r.get(current_field) or 0 for r in rows]
    prior_vals   = [r.get(prior_field) for r in rows]
    n            = len(labels)
    has_prior    = any(p is not None for p in prior_vals)

    # Row / block heights
    row_h   = min(_ABS_ROW_H_MAX, max(_ABS_ROW_H_MIN, 4.50 / max(n, 1)))
    body_h  = n * row_h
    total_h = _ABS_HDR_H + body_h

    # Scale config (from ask.extra, same keys as old renderer)
    extra      = ask.extra or {}
    hide_val   = extra.get("hide_val_cols", False)

    # Compute column left edges — center the whole block on the 13.33" slide
    if hide_val:
        n_gaps  = 2
        total_w = _ABS_LABEL_W + _ABS_CHART_W + _ABS_DELTA_W + n_gaps * _ABS_GAP
        # Give extra space to the chart
        extra_w = (2 if has_prior else 1) * _ABS_VAL_W + _ABS_GAP
        chart_w = _ABS_CHART_W + extra_w
        total_w = _ABS_LABEL_W + chart_w + _ABS_DELTA_W + n_gaps * _ABS_GAP
    else:
        n_val_cols = 2 if has_prior else 1
        n_gaps     = 3 if has_prior else 2
        chart_w    = _ABS_CHART_W
        total_w    = (_ABS_LABEL_W + n_val_cols * _ABS_VAL_W
                      + chart_w + _ABS_DELTA_W + n_gaps * _ABS_GAP)
    label_l  = (_ABS_SLIDE_W - total_w) / 2
    if hide_val:
        chart_l = label_l + _ABS_LABEL_W + _ABS_GAP
    else:
        prior_l  = label_l + _ABS_LABEL_W + _ABS_GAP
        curr_l   = (prior_l + _ABS_VAL_W) if has_prior else prior_l
        chart_l  = curr_l + _ABS_VAL_W + _ABS_GAP
    delta_l  = chart_l + chart_w + _ABS_GAP
    chart_top = _ABS_TOP + _ABS_HDR_H
    scale_min  = int(extra.get("scale_min", 0))
    scale_max  = int(extra.get("scale_max", 100))
    scale_sfx  = extra.get("scale_suffix", "%")
    ticks      = extra.get("scale_ticks") or [scale_min,
                                              (scale_min + scale_max) // 2,
                                              scale_max]

    all_vals = [v for v in (prior_vals if has_prior else []) + current_vals
                if v is not None]
    x_min, x_max = _abs_x_range(all_vals, scale_min, scale_max)

    def _chart_x(pct):
        """Map a percentage (0–100) to slide-inches within the chart frame."""
        frac = max(0.0, min(1.0, (pct - x_min * 100) / ((x_max - x_min) * 100)))
        return chart_l + frac * chart_w

    # NOTE: Horizontal row lines are rendered as Y-axis majorGridlines inside
    # the scatter chart (see _abs_scatter) so they aren't hidden behind the
    # chart shape.  Vertical reference lines at scale ticks are still drawn
    # as shapes because they sit outside the chart's X-axis range.
    for tick in ticks:
        if x_min * 100 <= tick <= x_max * 100:
            tx = _chart_x(tick)
            dashed_separator(slide, tx, chart_top, body_h,
                             color=C_LBGREY, width_pt=0.5, vertical=True)

    # 1. Attribute labels table
    _abs_label_tbl(slide, labels, label_l, _ABS_TOP, _ABS_LABEL_W,
                   _ABS_HDR_H, row_h, "Attribute", font)

    # 2–3. Value columns (skip if hide_val_cols)
    if not hide_val:
        current_header = extra.get("current_label", config.period_current)
        prior_header   = extra.get("prior_label", config.period_prior)
        if has_prior:
            _abs_val_col(slide, prior_vals, prior_l, _ABS_TOP, _ABS_VAL_W,
                         _ABS_HDR_H, row_h, prior_header, color_prior, font)
        _abs_val_col(slide, current_vals, curr_l, _ABS_TOP, _ABS_VAL_W,
                     _ABS_HDR_H, row_h, current_header, color_current, font)

    # 4. XY scatter chart (transparent bg so background shapes show through)
    show_labels = extra.get("show_data_labels", True)
    _abs_scatter(slide, prior_vals if has_prior else None, current_vals, n,
                 chart_l, chart_top, chart_w, body_h,
                 color_prior, color_current, x_min, x_max,
                 data_labels=show_labels)

    # Percentage scale labels below the chart
    lbl_y = chart_top + body_h + 0.04
    for tick in ticks:
        if x_min * 100 <= tick <= x_max * 100:
            tx = _chart_x(tick)
            textbox(slide, f"{tick}{scale_sfx}", tx - 0.20, lbl_y,
                    0.40, 0.18, fsize=7, color=C_FTGREY,
                    align=PP_ALIGN.CENTER, font=font)

    # 5. Delta column
    delta_header = _resolve_template(extra.get("delta_header", "QoQ \u0394"), config)
    deltas = [
        (current_vals[i] - prior_vals[i]) if prior_vals[i] is not None else None
        for i in range(n)
    ]
    add_delta_table(slide, deltas, delta_l, _ABS_TOP, _ABS_DELTA_W, row_h,
                    header_text=delta_header, font_name=font)

    # Legend — use custom labels from extra if provided
    legend_current = extra.get("legend_current")
    legend_prior   = extra.get("legend_prior")
    if legend_current:
        legend_items = [(color_current, legend_current)]
        if has_prior and legend_prior:
            legend_items.append((color_prior, legend_prior))
        elif has_prior:
            legend_items.append((color_prior, prior_header))
    else:
        sample  = config.sample_sizes.get(ask.brand or "primary")
        n_curr  = f" (n={sample.current})" if sample else ""
        legend_items = [(color_current, f"{config.period_current}{n_curr}")]
        if has_prior:
            n_prior = f" (n={sample.prior})" if sample else ""
            legend_items.append((color_prior, f"{config.period_prior}{n_prior}"))
    _make_legend(slide, legend_items, 0, _ABS_TOP + total_h + LEGEND_GAP, font,
                 center_over=(0, _ABS_SLIDE_W))


# ── Private abacus helpers ────────────────────────────────────────────────────

def _abs_cell_mid(cell) -> None:
    """Vertically center text within a table cell."""
    cell_vcenter(cell)


def _abs_label_tbl(slide, labels, left, top, width, hdr_h, row_h, header, font):
    n      = len(labels)
    total  = hdr_h + n * row_h
    tbl    = slide.shapes.add_table(
        n + 1, 1, Inches(left), Inches(top), Inches(width), Inches(total)
    ).table
    tbl.columns[0].width = Inches(width)

    hdr = tbl.cell(0, 0)
    tbl.rows[0].height = Inches(hdr_h)
    hdr.fill.solid(); hdr.fill.fore_color.rgb = C_HDRGREY
    p = hdr.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    suppress_para_bullets(p._p)
    r = p.add_run()
    r.text = header; r.font.size = Pt(7); r.font.bold = True
    r.font.color.rgb = C_WHITE; r.font.name = font or FONT_TEXT
    _abs_cell_mid(hdr)

    for i, label in enumerate(labels):
        c = tbl.cell(i + 1, 0)
        tbl.rows[i + 1].height = Inches(row_h)
        c.fill.solid()
        c.fill.fore_color.rgb = _alt_row_bg(i)
        p = c.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
        suppress_para_bullets(p._p)
        r = p.add_run()
        r.text = label; r.font.size = Pt(7)
        r.font.color.rgb = C_GREY; r.font.name = font or FONT_TEXT
        _abs_cell_mid(c)


def _abs_val_col(slide, vals, left, top, width, hdr_h, row_h,
                 period_label, color, font):
    n     = len(vals)
    total = hdr_h + n * row_h
    tbl   = slide.shapes.add_table(
        n + 1, 1, Inches(left), Inches(top), Inches(width), Inches(total)
    ).table
    tbl.columns[0].width = Inches(width)

    hdr = tbl.cell(0, 0)
    tbl.rows[0].height = Inches(hdr_h)
    hdr.fill.solid(); hdr.fill.fore_color.rgb = C_HDRGREY
    p = hdr.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    suppress_para_bullets(p._p)
    r = p.add_run()
    r.text = period_label; r.font.size = Pt(7); r.font.bold = True
    r.font.color.rgb = C_WHITE; r.font.name = font or FONT_TEXT
    _abs_cell_mid(hdr)

    for i, val in enumerate(vals):
        c = tbl.cell(i + 1, 0)
        tbl.rows[i + 1].height = Inches(row_h)
        c.fill.solid()
        c.fill.fore_color.rgb = _alt_row_bg(i)
        text = f"{val:.0f}%" if val is not None else "\u2014"
        p = c.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        suppress_para_bullets(p._p)
        r = p.add_run()
        r.text = text; r.font.size = Pt(8); r.font.bold = True
        r.font.color.rgb = color; r.font.name = font or FONT_TEXT
        _abs_cell_mid(c)


def _abs_x_range(vals, scale_min=0, scale_max=100):
    """Compute X axis decimal range (0.0–1.0) for the scatter chart.

    Uses scale_min/scale_max from ask.extra when provided, otherwise
    auto-ranges from the data values.
    """
    if scale_min != 0 or scale_max != 100:
        return scale_min / 100.0, scale_max / 100.0
    if not vals:
        return 0.0, 1.0
    vals = [v for v in vals if v is not None and v != 0]
    if not vals:
        return 0.0, 1.0
    lo, hi = min(vals), max(vals)
    x_min = max(0.0, (int(lo // 10) - 1) * 10) / 100.0
    x_max = min(100.0, (int(hi // 10) + 1) * 10) / 100.0
    return x_min, x_max


def _abs_scatter(slide, prior_vals, current_vals, n,
                 left, top, width, height,
                 color_prior, color_current, x_min, x_max,
                 *, data_labels=False, label_color=None):
    """Add XY scatter chart matching template slides 18/20.

    Template axis pattern (slides 18 & 20):
        Y axis: orient=minMax, min=1, max=n, majorUnit=1, majorGridlines=True
        X axis: orient=minMax, min/max from data, hidden
        Plot area: default layout (natural padding keeps edge dots visible)

    Args:
        data_labels: If True, add per-point percentage labels (current series
            only) matching the template's dLbl rich-text pattern.
        label_color: RGBColor for data label text.  Defaults to C_GREY.
    """
    from lxml import etree as _etree

    chart_data = XyChartData()
    # Row 0 (top) = rank n; row n-1 (bottom) = rank 1
    # With orient=minMax, high Y values appear at the top → first row on top.
    ranks = list(range(n, 0, -1))

    if prior_vals is not None:
        s_p = chart_data.add_series("Prior")
        for rank, val in zip(ranks, prior_vals):
            if val is not None:
                s_p.add_data_point(val / 100.0, float(rank))

    s_c = chart_data.add_series("Current")
    for rank, val in zip(ranks, current_vals):
        s_c.add_data_point((val or 0) / 100.0, float(rank))

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.XY_SCATTER,
        Inches(left), Inches(top), Inches(width), Inches(height),
        chart_data)
    ch = cf.chart
    ch.has_legend = False
    ch.has_title  = False

    # Make chart background transparent so background shapes show through
    for spPr_parent, spPr_tag in (
        (ch._element, "c:spPr"),
        (ch._element.find(_qn("c:chart")).find(_qn("c:plotArea")), "c:spPr"),
    ):
        if spPr_parent is None:
            continue
        sp = _get_or_add(spPr_parent, spPr_tag)
        for fill_tag in ("a:solidFill", "a:gradFill", "a:pattFill", "a:blipFill"):
            el = sp.find(_qn(fill_tag))
            if el is not None:
                sp.remove(el)
        if sp.find(_qn("a:noFill")) is None:
            _etree.SubElement(sp, _qn("a:noFill"))

    # Style series: filled circles, no connecting line
    has_two = prior_vals is not None
    for i, series in enumerate(ch.series):
        c = color_prior if (i == 0 and has_two) else color_current
        set_series_marker(series, "circle", size=8, fill_color=c, line_color=c)
        set_series_line_style(series, visible=False)

    # ── Axis configuration (matches template slides 18/20) ────────────
    chart_el  = ch._element.find(_qn("c:chart"))
    plot_area = chart_el.find(_qn("c:plotArea"))
    scatter_el = plot_area.find(_qn("c:scatterChart"))
    ax_ids    = [el.get("val") for el in scatter_el.findall(_qn("c:axId"))]
    ax_by_id  = {
        ax.find(_qn("c:axId")).get("val"): ax
        for ax in plot_area.findall(_qn("c:valAx"))
        if ax.find(_qn("c:axId")) is not None
    }
    x_ax = ax_by_id.get(ax_ids[0]) if ax_ids else None
    y_ax = ax_by_id.get(ax_ids[1]) if len(ax_ids) > 1 else None

    # X axis — percentage range, hidden, no gridlines
    if x_ax is not None:
        sc = _get_or_add(x_ax, "c:scaling")
        _get_or_add(sc, "c:orientation").set("val", "minMax")
        _get_or_add(sc, "c:min").set("val", str(x_min))
        _get_or_add(sc, "c:max").set("val", str(x_max))
        _get_or_add(x_ax, "c:delete").set("val", "1")
        _get_or_add(x_ax, "c:tickLblPos").set("val", "none")
        _get_or_add(x_ax, "c:crossBetween").set("val", "midCat")
        for tag in ("c:majorGridlines", "c:minorGridlines"):
            child = x_ax.find(_qn(tag))
            if child is not None:
                x_ax.remove(child)

    # Y axis — orient=minMax, min=1, max=n, majorUnit=1
    # Template pattern: gridlines at Y=1,2,…,n pass exactly through dots.
    if y_ax is not None:
        sc = _get_or_add(y_ax, "c:scaling")
        _get_or_add(sc, "c:orientation").set("val", "minMax")
        _get_or_add(sc, "c:min").set("val", "1")
        _get_or_add(sc, "c:max").set("val", str(n))
        _get_or_add(y_ax, "c:delete").set("val", "1")
        _get_or_add(y_ax, "c:tickLblPos").set("val", "none")
        _get_or_add(y_ax, "c:majorUnit").set("val", "1")
        _get_or_add(y_ax, "c:crossBetween").set("val", "midCat")

        # Remove minor gridlines
        minor = y_ax.find(_qn("c:minorGridlines"))
        if minor is not None:
            y_ax.remove(minor)

        # Add light-grey major gridlines (matching template schemeClr bg1/85%)
        gl = _get_or_add(y_ax, "c:majorGridlines")
        gl_sp = _get_or_add(gl, "c:spPr")
        gl_ln = _get_or_add(gl_sp, "a:ln")
        for tag in ("a:solidFill", "a:noFill"):
            el = gl_ln.find(_qn(tag))
            if el is not None:
                gl_ln.remove(el)
        sf = _etree.SubElement(gl_ln, _qn("a:solidFill"))
        clr = _etree.SubElement(sf, _qn("a:srgbClr"))
        clr.set("val", "D9D9D9")

    # Use default plot area layout (natural padding ~4-6%) — do NOT force 0,0,1,1
    # This matches the template and keeps dots at min/max Y visible.

    # ── Per-point data labels (both series) ────────────────────────────
    if data_labels:
        lbl_rgb = label_color or C_GREY
        lbl_hex = f"{lbl_rgb[0]:02X}{lbl_rgb[1]:02X}{lbl_rgb[2]:02X}" if hasattr(lbl_rgb, '__getitem__') else "808080"
        ser_els = scatter_el.findall(_qn("c:ser"))

        for si, ser_el in enumerate(ser_els):
            is_prior = (si == 0 and has_two)
            # Position: labels above dots; prior slightly left, current slightly right
            # Dynamic y-offset: scale with row count to avoid overlapping previous row
            # Few rows (≤5): large offset; many rows (15+): tight offset
            y_frac = min(0.06, max(0.02, 0.30 / max(n, 1)))
            x_off = "-0.03" if is_prior else "0.01"
            y_off = f"-{y_frac:.3f}"

            # Get series color for label
            ser_color = color_prior if is_prior else color_current
            ser_hex = f"{ser_color[0]:02X}{ser_color[1]:02X}{ser_color[2]:02X}" if hasattr(ser_color, '__getitem__') else lbl_hex

            dLbls = _get_or_add(ser_el, "c:dLbls")
            for tag in ("showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize"):
                _get_or_add(dLbls, f"c:{tag}").set("val", "0")

            xVal_el = ser_el.find(_qn("c:xVal"))
            numRef = xVal_el.find(_qn("c:numRef")) if xVal_el is not None else None
            numCache = numRef.find(_qn("c:numCache")) if numRef is not None else None
            if numCache is None:
                continue

            for pt in numCache.findall(_qn("c:pt")):
                idx = pt.get("idx")
                raw_v = pt.find(_qn("c:v"))
                if raw_v is None or raw_v.text is None:
                    continue
                pct_val = float(raw_v.text) * 100
                label_text = f"{pct_val:.0f}%"

                dLbl = _etree.SubElement(dLbls, _qn("c:dLbl"))
                _etree.SubElement(dLbl, _qn("c:idx")).set("val", idx)
                layout_el = _etree.SubElement(dLbl, _qn("c:layout"))
                manLayout = _etree.SubElement(layout_el, _qn("c:manualLayout"))
                _etree.SubElement(manLayout, _qn("c:x")).set("val", x_off)
                _etree.SubElement(manLayout, _qn("c:y")).set("val", y_off)
                # Rich text with percentage
                tx = _etree.SubElement(dLbl, _qn("c:tx"))
                rich = _etree.SubElement(tx, _qn("c:rich"))
                bodyPr = _etree.SubElement(rich, _qn("a:bodyPr"))
                bodyPr.set("wrap", "square")
                bodyPr.set("lIns", "38100"); bodyPr.set("tIns", "19050")
                bodyPr.set("rIns", "38100"); bodyPr.set("bIns", "19050")
                bodyPr.set("anchor", "ctr")
                _etree.SubElement(bodyPr, _qn("a:spAutoFit"))
                _etree.SubElement(rich, _qn("a:lstStyle"))
                p = _etree.SubElement(rich, _qn("a:p"))
                pPr = _etree.SubElement(p, _qn("a:pPr"))
                defRPr = _etree.SubElement(pPr, _qn("a:defRPr"))
                defRPr.set("sz", "700"); defRPr.set("b", "1")
                sFill = _etree.SubElement(defRPr, _qn("a:solidFill"))
                sClr = _etree.SubElement(sFill, _qn("a:srgbClr"))
                sClr.set("val", ser_hex)
                r = _etree.SubElement(p, _qn("a:r"))
                rPr = _etree.SubElement(r, _qn("a:rPr"))
                rPr.set("lang", "en-US"); rPr.set("sz", "700"); rPr.set("b", "1")
                rFill = _etree.SubElement(rPr, _qn("a:solidFill"))
                rClr = _etree.SubElement(rFill, _qn("a:srgbClr"))
                rClr.set("val", ser_hex)
                t = _etree.SubElement(r, _qn("a:t"))
                t.text = label_text
                for tag in ("showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize"):
                    _get_or_add(dLbl, f"c:{tag}").set("val", "0")


# Follow-up rep, dual abacus, and message MBD renderers moved to dot_special.py
