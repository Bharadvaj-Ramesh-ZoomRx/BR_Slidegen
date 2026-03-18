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
    _slide_chrome, _get_brand_colors, _sort_data, _make_legend, _add_dot,
    # pptx_utils
    C_GREY, C_FTGREY, C_LBGREY, C_HDRGREY, C_WHITE, C_RED,
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


def render_lollipop(slide, config: ProjectConfig, ask: AskConfig, data: dict, *, namer=None):
    """Lollipop chart: horizontal stem (line) + filled circle at current value.

    Optionally overlays a smaller prior-period dot when prior data is available.

    Optional ask.extra keys:
      scale_max (int): x-axis maximum; default = rounded up to next 10 above max value
    """
    _slide_chrome(slide, config, ask)

    rows = data.get(ask.data_key, [])
    rows = _sort_data(rows, ask.sort_by, ask.sort_desc)
    if not rows:
        textbox(slide, "Data not available", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    color_current, color_prior = _get_brand_colors(config, ask)
    font = config.font_body
    extra = ask.extra or {}

    labels = [r.get("short", r.get("desc", ""))[:LABEL_MAX_SINGLE] for r in rows]
    current_vals = [r.get("current") or 0 for r in rows]
    prior_vals = [r.get("prior") for r in rows]
    n = len(labels)
    has_prior = any(p is not None for p in prior_vals)

    # Layout
    chart_top = CHART_TOP_STD
    chart_h = max(MIN_CHART_HEIGHT, min(MAX_CHART_HEIGHT, n * 0.42))
    row_h = chart_h / max(n, 1)

    # Scale
    raw_max = max(current_vals) if current_vals else 100
    scale_max = extra.get("scale_max") or min(100, (int(raw_max // 10) + 1) * 10)

    def _x(val):
        return LOLLI_PLOT_LEFT + (val / scale_max) * LOLLI_PLOT_W

    # Vertical gridlines + x-axis tick labels
    for pct in [0, 25, 50, 75, 100]:
        if pct > scale_max:
            break
        gx = _x(pct)
        dashed_separator(slide, gx, chart_top, chart_h,
                          color=C_LBGREY, width_pt=0.5, vertical=True)
        textbox(slide, f"{pct}%", gx - 0.18, chart_top + chart_h + 0.05,
                0.36, 0.18, fsize=7, color=C_FTGREY,
                align=PP_ALIGN.CENTER, font=font)

    # Lollipop rows
    for i, (label, val) in enumerate(zip(labels, current_vals)):
        y_mid = chart_top + (i + 0.5) * row_h
        dot_x = _x(val)

        # Category label (right-aligned, left of plot area)
        textbox(slide, label, 0.30, y_mid - 0.14, LOLLI_LABEL_W - 0.15, 0.28,
                fsize=8, color=C_GREY, align=PP_ALIGN.RIGHT, font=font)

        # Stem: horizontal line from axis to just before the dot
        stem_len = dot_x - LOLLI_PLOT_LEFT - LOLLI_DOT_R
        if stem_len > 0:
            horiz_line(slide, LOLLI_PLOT_LEFT, y_mid, stem_len,
                       color=color_current, width_pt=1.5)

        # Prior dot (smaller, drawn first so current sits on top)
        prior = prior_vals[i]
        if prior is not None:
            _add_dot(slide, _x(prior), y_mid, LOLLI_PRIOR_R, color_prior)

        # Current dot
        r = LOLLI_DOT_R
        _add_dot(slide, dot_x, y_mid, r, color_current)

        # Value label (right of dot)
        textbox(slide, f"{val:.0f}%", dot_x + r + 0.04, y_mid - 0.11,
                0.38, 0.22, fsize=7, bold=True, color=color_current, font=font)

    # Legend
    ly = chart_top + chart_h + LEGEND_GAP
    sample = config.sample_sizes.get(ask.brand or "primary")
    n_curr = f" (n={sample.current})" if sample else ""
    legend_items = [(color_current, f"{config.period_current}{n_curr}")]
    if has_prior:
        n_prior = f" (n={sample.prior})" if sample else ""
        legend_items.append((color_prior, f"{config.period_prior}{n_prior}"))
    _make_legend(slide, legend_items, 0, ly, font, center_over=(0, SLIDE_W))


def render_abacus(slide, config: ProjectConfig, ask: AskConfig, data: dict, *, namer=None):
    """Abacus chart using XY scatter + tables layout (mirrors template slide 20).

    Layout (left → right, horizontally centered on slide):
      Attribute labels | Prior % | Current % | XY scatter dots | QoQ Δ
    """
    _slide_chrome(slide, config, ask)

    rows = data.get(ask.data_key, [])
    rows = _sort_data(rows, ask.sort_by, ask.sort_desc)
    if not rows:
        textbox(slide, "Data not available", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    color_current, color_prior = _get_brand_colors(config, ask)
    font = config.font_body

    labels       = [r.get("short", r.get("desc", ""))[:LABEL_MAX_SINGLE] for r in rows]
    current_vals = [r.get("current") or 0 for r in rows]
    prior_vals   = [r.get("prior") for r in rows]
    n            = len(labels)
    has_prior    = any(p is not None for p in prior_vals)

    # Row / block heights
    row_h   = min(_ABS_ROW_H_MAX, max(_ABS_ROW_H_MIN, 4.50 / max(n, 1)))
    body_h  = n * row_h
    total_h = _ABS_HDR_H + body_h

    # Compute column left edges — center the whole block on the 13.33" slide
    n_val_cols = 2 if has_prior else 1
    n_gaps     = 3 if has_prior else 2
    total_w    = (_ABS_LABEL_W + n_val_cols * _ABS_VAL_W
                  + _ABS_CHART_W + _ABS_DELTA_W + n_gaps * _ABS_GAP)
    label_l  = (_ABS_SLIDE_W - total_w) / 2
    prior_l  = label_l + _ABS_LABEL_W + _ABS_GAP
    curr_l   = (prior_l + _ABS_VAL_W) if has_prior else prior_l
    chart_l  = curr_l + _ABS_VAL_W + _ABS_GAP
    delta_l  = chart_l + _ABS_CHART_W + _ABS_GAP
    chart_top = _ABS_TOP + _ABS_HDR_H

    # Scale config (from ask.extra, same keys as old renderer)
    extra      = ask.extra or {}
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
        return chart_l + frac * _ABS_CHART_W

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

    # 2. Prior value column
    if has_prior:
        _abs_val_col(slide, prior_vals, prior_l, _ABS_TOP, _ABS_VAL_W,
                     _ABS_HDR_H, row_h, config.period_prior, color_prior, font)

    # 3. Current value column
    _abs_val_col(slide, current_vals, curr_l, _ABS_TOP, _ABS_VAL_W,
                 _ABS_HDR_H, row_h, config.period_current, color_current, font)

    # 4. XY scatter chart (transparent bg so background shapes show through)
    _abs_scatter(slide, prior_vals if has_prior else None, current_vals, n,
                 chart_l, chart_top, _ABS_CHART_W, body_h,
                 color_prior, color_current, x_min, x_max)

    # Percentage scale labels below the chart
    lbl_y = chart_top + body_h + 0.04
    for tick in ticks:
        if x_min * 100 <= tick <= x_max * 100:
            tx = _chart_x(tick)
            textbox(slide, f"{tick}{scale_sfx}", tx - 0.20, lbl_y,
                    0.40, 0.18, fsize=7, color=C_FTGREY,
                    align=PP_ALIGN.CENTER, font=font)

    # 5. Delta column
    deltas = [
        (current_vals[i] - prior_vals[i]) if prior_vals[i] is not None else None
        for i in range(n)
    ]
    add_delta_table(slide, deltas, delta_l, _ABS_TOP, _ABS_DELTA_W, row_h,
                    header_text="QoQ \u0394", font_name=font)

    # Legend
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
        c.fill.fore_color.rgb = C_LBGREY if i % 2 == 0 else C_WHITE
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
        c.fill.fore_color.rgb = C_LBGREY if i % 2 == 0 else C_WHITE
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

    # ── Per-point data labels (current series only) ───────────────────
    if data_labels:
        lbl_rgb = label_color or C_GREY
        lbl_hex = f"{lbl_rgb[0]:02X}{lbl_rgb[1]:02X}{lbl_rgb[2]:02X}" if hasattr(lbl_rgb, '__getitem__') else "808080"
        # Find the last series element (current)
        ser_els = scatter_el.findall(_qn("c:ser"))
        cur_ser = ser_els[-1]  # current is always the last series
        dLbls = _get_or_add(cur_ser, "c:dLbls")
        # Default: hide all labels at series level
        _get_or_add(dLbls, "c:showLegendKey").set("val", "0")
        _get_or_add(dLbls, "c:showVal").set("val", "0")
        _get_or_add(dLbls, "c:showCatName").set("val", "0")
        _get_or_add(dLbls, "c:showSerName").set("val", "0")
        _get_or_add(dLbls, "c:showPercent").set("val", "0")
        _get_or_add(dLbls, "c:showBubbleSize").set("val", "0")

        # Read xVal percentage data from the series cache
        xVal_el = cur_ser.find(_qn("c:xVal"))
        numRef = xVal_el.find(_qn("c:numRef")) if xVal_el is not None else None
        numCache = numRef.find(_qn("c:numCache")) if numRef is not None else None
        if numCache is not None:
            for pt in numCache.findall(_qn("c:pt")):
                idx = pt.get("idx")
                raw_v = pt.find(_qn("c:v"))
                if raw_v is None or raw_v.text is None:
                    continue
                pct_val = float(raw_v.text) * 100
                label_text = f"{pct_val:.0f}%"

                dLbl = _etree.SubElement(dLbls, _qn("c:dLbl"))
                _etree.SubElement(dLbl, _qn("c:idx")).set("val", idx)
                # Offset: slightly above and left of dot (template pattern)
                layout_el = _etree.SubElement(dLbl, _qn("c:layout"))
                manLayout = _etree.SubElement(layout_el, _qn("c:manualLayout"))
                _etree.SubElement(manLayout, _qn("c:x")).set("val", "-0.05")
                _etree.SubElement(manLayout, _qn("c:y")).set("val", "-0.045")
                # Rich text with percentage
                tx = _etree.SubElement(dLbl, _qn("c:tx"))
                rich = _etree.SubElement(tx, _qn("c:rich"))
                bodyPr = _etree.SubElement(rich, _qn("a:bodyPr"))
                bodyPr.set("wrap", "square")
                bodyPr.set("lIns", "38100")
                bodyPr.set("tIns", "19050")
                bodyPr.set("rIns", "38100")
                bodyPr.set("bIns", "19050")
                bodyPr.set("anchor", "ctr")
                _etree.SubElement(bodyPr, _qn("a:spAutoFit"))
                _etree.SubElement(rich, _qn("a:lstStyle"))
                p = _etree.SubElement(rich, _qn("a:p"))
                pPr = _etree.SubElement(p, _qn("a:pPr"))
                defRPr = _etree.SubElement(pPr, _qn("a:defRPr"))
                defRPr.set("sz", "800")
                defRPr.set("b", "0")
                sFill = _etree.SubElement(defRPr, _qn("a:solidFill"))
                sClr = _etree.SubElement(sFill, _qn("a:srgbClr"))
                sClr.set("val", lbl_hex)
                r = _etree.SubElement(p, _qn("a:r"))
                rPr = _etree.SubElement(r, _qn("a:rPr"))
                rPr.set("lang", "en-US")
                rPr.set("sz", "800")
                rFill = _etree.SubElement(rPr, _qn("a:solidFill"))
                rClr = _etree.SubElement(rFill, _qn("a:srgbClr"))
                rClr.set("val", lbl_hex)
                t = _etree.SubElement(r, _qn("a:t"))
                t.text = label_text
                # Show flags for this point
                _get_or_add(dLbl, "c:showLegendKey").set("val", "0")
                _get_or_add(dLbl, "c:showVal").set("val", "0")
                _get_or_add(dLbl, "c:showCatName").set("val", "0")
                _get_or_add(dLbl, "c:showSerName").set("val", "0")
                _get_or_add(dLbl, "c:showPercent").set("val", "0")
                _get_or_add(dLbl, "c:showBubbleSize").set("val", "0")


# ══════════════════════════════════════════════════════════════════════════════
# Message MBD — Multi-column abacus layout (template slide 18)
# ══════════════════════════════════════════════════════════════════════════════

# Layout constants from template slide 18
_MBD_TOP         = 2.10      # top of tables/charts (below section + column headers)
_MBD_HDR_H       = _ABS_HDR_H  # reuse abacus header height
_MBD_ROW_H_MIN   = 0.30
_MBD_ROW_H_MAX   = 0.42
_MBD_LABEL_W     = 3.03      # category label table width (template: 3.034")
_MBD_LABEL_L     = 0.47      # category label table left
_MBD_CE_W        = 1.66      # CE value table width (2 cols: prior + current)
_MBD_CE_COL_W    = 0.83      # each CE column width
_MBD_SCATTER_W   = 1.90      # each scatter chart width (template: 1.900")
_MBD_GAP         = 0.09      # gap between elements
_MBD_COL_HDR_TOP = 1.78      # column header labels Y
_MBD_COL_HDR_H   = 0.29      # column header height
_MBD_SCALE_LBL_Y_OFF = 0.04  # offset below charts for "% of Responses" label

# Default MBD series colors
C_MOTIV  = RGBColor(0x4E, 0x79, 0xA7)   # steel blue — Motivation
C_BELIEV = RGBColor(0x59, 0xA1, 0x4F)   # green — Believability
C_DIFF   = RGBColor(0xE1, 0x57, 0x59)   # coral red — Differentiation


def render_message_mbd(slide, config: ProjectConfig, ask: AskConfig, data: dict, *, namer=None):
    """Multi-column abacus layout for Message Effectiveness MBD breakdown.

    Layout (matches template slide 18):
        Category labels | CE prior | CE current | Scatter 1 | Scatter 2 | Scatter 3

    Each scatter chart shows prior + current dots per message for one MBD metric.

    Expected data fields per row:
        desc / short   — message label
        motiv          — Motivation (current %)
        motiv_prior    — Motivation (prior %, optional)
        believ         — Believability (current %)
        believ_prior   — Believability (prior %, optional)
        diff           — Differentiation (current %)
        diff_prior     — Differentiation (prior %, optional)
        ce_current     — Composite Effectiveness current (%)
        ce_prior       — Composite Effectiveness prior (%)

    extra config keys (all optional):
        series:
          - { field: "believ", label: "Believability",   color: "#59A14F" }
          - { field: "diff",   label: "Differentiation", color: "#E15759" }
          - { field: "motiv",  label: "Motivation",      color: "#4E79A7" }
        ce_field:       "ce_current"
        ce_prior_field: "ce_prior"
        ce_header:      "Message Effectiveness^"
        scale_label:    "% of Responses (Rated 6, 7)"
    """
    _slide_chrome(slide, config, ask)

    rows = data.get(ask.data_key, [])
    rows = _sort_data(rows, ask.sort_by, ask.sort_desc)
    if not rows:
        textbox(slide, "Data not available", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    color_current, color_prior = _get_brand_colors(config, ask)
    font = config.font_body
    extra = ask.extra or {}

    # ── Series configuration ────────────────────────────────────────────
    default_series = [
        {"field": "believ", "label": "Believability",   "color": None},
        {"field": "diff",   "label": "Differentiation", "color": None},
        {"field": "motiv",  "label": "Motivation",      "color": None},
    ]
    default_colors = [C_BELIEV, C_DIFF, C_MOTIV]
    series_cfg = extra.get("series", default_series)

    series_colors = []
    for i, s in enumerate(series_cfg):
        if s.get("color"):
            series_colors.append(parse_color(s["color"]))
        elif i < len(default_colors):
            series_colors.append(default_colors[i])
        else:
            series_colors.append(C_GREY)

    ce_field = extra.get("ce_field", "ce_current")
    ce_prior_field = extra.get("ce_prior_field", "ce_prior")
    ce_header = extra.get("ce_header", "Message Effectiveness^")
    scale_label = extra.get("scale_label", "% of Responses (Rated 6, 7)")

    # ── Data preparation ────────────────────────────────────────────────
    labels = [r.get("short", r.get("desc", ""))[:LABEL_MAX_DUAL] for r in rows]
    n = len(labels)

    row_h = min(_MBD_ROW_H_MAX, max(_MBD_ROW_H_MIN, 4.15 / max(n, 1)))
    body_h = n * row_h

    ce_current_vals = [r.get(ce_field) for r in rows]
    ce_prior_vals = [r.get(ce_prior_field) for r in rows]
    has_ce_prior = any(v is not None for v in ce_prior_vals)

    # ── Compute column positions (left to right) ────────────────────────
    n_scatter = len(series_cfg)
    n_ce_cols = 2 if has_ce_prior else 1
    ce_total_w = n_ce_cols * _MBD_CE_COL_W

    label_l = _MBD_LABEL_L
    ce_l = label_l + _MBD_LABEL_W + _MBD_GAP
    scatter_start = ce_l + ce_total_w + _MBD_GAP

    chart_top = _MBD_TOP + _MBD_HDR_H

    # ── 1. Category labels table ────────────────────────────────────────
    _abs_label_tbl(slide, labels, label_l, _MBD_TOP, _MBD_LABEL_W,
                   _MBD_HDR_H, row_h, "Tag$", font)

    # ── 2. CE value columns ─────────────────────────────────────────────
    # Column header above CE columns
    textbox(slide, ce_header,
            ce_l, _MBD_COL_HDR_TOP - 0.20, ce_total_w, _MBD_COL_HDR_H,
            fsize=7, bold=True, color=C_GREY, font=font)

    if has_ce_prior:
        _abs_val_col(slide, ce_prior_vals, ce_l, _MBD_TOP, _MBD_CE_COL_W,
                     _MBD_HDR_H, row_h, config.period_prior, color_prior, font)
        _abs_val_col(slide, ce_current_vals, ce_l + _MBD_CE_COL_W, _MBD_TOP,
                     _MBD_CE_COL_W, _MBD_HDR_H, row_h,
                     config.period_current, color_current, font)
    else:
        _abs_val_col(slide, ce_current_vals, ce_l, _MBD_TOP, ce_total_w,
                     _MBD_HDR_H, row_h, config.period_current, color_current, font)

    # ── 3. Scatter charts (one per MBD metric) ──────────────────────────
    # Collect all values to compute shared X axis range
    all_scatter_vals = []
    for s in series_cfg:
        field = s.get("field", "")
        prior_field = s.get("prior_field", f"{field}_prior")
        for r in rows:
            v = r.get(field)
            if v is not None:
                all_scatter_vals.append(v)
            vp = r.get(prior_field)
            if vp is not None:
                all_scatter_vals.append(vp)

    x_min, x_max = _abs_x_range(all_scatter_vals, 0, 100)

    for si, s_cfg in enumerate(series_cfg):
        field = s_cfg.get("field", "")
        prior_field = s_cfg.get("prior_field", f"{field}_prior")
        label = s_cfg.get("label", field)
        sc_color = series_colors[si]

        sc_left = scatter_start + si * (_MBD_SCATTER_W + _MBD_GAP)

        # Column header
        textbox(slide, label,
                sc_left, _MBD_COL_HDR_TOP, _MBD_SCATTER_W, _MBD_COL_HDR_H,
                fsize=8, bold=True, color=sc_color,
                align=PP_ALIGN.CENTER, font=font)

        # Get current and prior values
        current_vals = [r.get(field) or 0 for r in rows]
        prior_vals = [r.get(prior_field) for r in rows]
        has_prior = any(p is not None for p in prior_vals)

        # Scatter chart with gridlines + data labels (matches template)
        _abs_scatter(slide,
                     prior_vals if has_prior else None,
                     current_vals, n,
                     sc_left, chart_top, _MBD_SCATTER_W, body_h,
                     color_prior, sc_color, x_min, x_max,
                     data_labels=True, label_color=sc_color)

        # Scale label below each chart
        textbox(slide, scale_label,
                sc_left - 0.15, chart_top + body_h + _MBD_SCALE_LBL_Y_OFF,
                _MBD_SCATTER_W + 0.30, 0.22,
                fsize=6.5, color=C_FTGREY, align=PP_ALIGN.CENTER, font=font)

    # ── Legend ───────────────────────────────────────────────────────────
    ly = chart_top + body_h + 0.32
    legend_items = [
        (color_current, config.period_current),
        (color_prior, config.period_prior),
    ]
    _make_legend(slide, legend_items, 0, ly, font, center_over=(0, SLIDE_W))
