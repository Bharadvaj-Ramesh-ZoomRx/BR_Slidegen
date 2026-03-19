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
    _pptx_table, _style_tbl_cell,
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

    rows = data.get(ask.data_key, [])
    rows = _sort_data(rows, ask.sort_by, ask.sort_desc)
    if not rows:
        textbox(slide, "Data not available", 2, 3, 8, 1, fsize=14, color=C_RED)
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


# ══════════════════════════════════════════════════════════════════════════════
# Follow-up Rep — template slide 51 layout
# ══════════════════════════════════════════════════════════════════════════════

# Layout constants from template slide 51
_FR_LABEL_L    = 1.87     # label table left
_FR_LABEL_W    = 3.26     # label table width
_FR_TOP        = 2.50     # table/chart top
_FR_ROW_H      = 0.709    # row height
_FR_CHART_L    = 5.13     # scatter chart left
_FR_CHART_W    = 4.04     # scatter chart width
_FR_DELTA_L    = 9.58     # delta table left
_FR_DELTA_W    = 1.22     # delta table width (2 cols)
_FR_DELTA_COL  = 0.61     # each delta column width
_FR_HDR_TOP    = 2.05     # "Representative Types" header Y
_FR_VS_HDR_TOP = 1.92     # "vs Q3'25" header Y
_FR_AX_LBL_TOP = 5.99     # "% of Interactions" axis label Y


def render_followup_rep(slide, config, ask, data, *, namer=None):
    """Follow-up rep abacus — matches template slide 51 layout.

    Two-brand scatter comparison with label table, XY scatter, and
    per-brand QoQ delta columns. Uses larger fonts and wider spacing
    than the generic abacus renderer.

    Config (ask.extra):
        current_field / prior_field: field names for brand 1 / brand 2
        current_label / prior_label: column headers ("J&J" / "AZ")
        color_current / color_prior: hex colors
        legend_current / legend_prior: legend labels
        delta_current_field / delta_prior_field: field names for deltas
        scale_min, scale_max, scale_ticks, scale_suffix
    """
    _slide_chrome(slide, config, ask)

    extra = ask.extra or {}
    font  = config.font_body

    current_field = extra.get("current_field", "current")
    prior_field   = extra.get("prior_field", "prior")
    delta_cur_field = extra.get("delta_current_field", "jnj_delta")
    delta_pri_field = extra.get("delta_prior_field", "az_delta")

    rows = data.get(ask.data_key, [])
    rows = _sort_data(rows, ask.sort_by, ask.sort_desc)
    rows = [r for r in rows if r.get(current_field) not in (None, 0) or r.get(prior_field) not in (None, 0)]
    if not rows:
        textbox(slide, "Data not available", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    # Colors
    color_current = parse_color(extra["color_current"]) if "color_current" in extra else config.primary.color_current
    color_prior   = parse_color(extra["color_prior"]) if "color_prior" in extra else config.competitor.color_current

    n = len(rows)
    labels      = [r.get("desc", "")[:40] for r in rows]
    cur_vals    = [r.get(current_field) or 0 for r in rows]
    pri_vals    = [r.get(prior_field) for r in rows]
    delta_cur   = [r.get(delta_cur_field) for r in rows]
    delta_pri   = [r.get(delta_pri_field) for r in rows]
    has_prior   = any(p is not None for p in pri_vals)

    body_h = n * _FR_ROW_H
    chart_h = body_h

    # Scale
    scale_min = int(extra.get("scale_min", 0))
    scale_max = int(extra.get("scale_max", 70))
    scale_sfx = extra.get("scale_suffix", "%")
    ticks = extra.get("scale_ticks") or [scale_min, (scale_min + scale_max) // 2, scale_max]

    all_vals = [v for v in cur_vals + (pri_vals if has_prior else []) if v is not None and v != 0]
    x_min, x_max = _abs_x_range(all_vals, scale_min, scale_max)

    # 1. "Representative Types" header
    textbox(slide, "Representative Types",
            _FR_LABEL_L + 0.15, _FR_HDR_TOP, _FR_LABEL_W - 0.30, 0.28,
            fsize=9, bold=True, color=C_GREY, font=font)

    # 2. Label table (10pt font, matching template)
    _, tbl = _pptx_table(slide, [_FR_LABEL_W], [_FR_ROW_H] * n,
                              _FR_LABEL_L, _FR_TOP)
    for i, label in enumerate(labels):
        cell = tbl.cell(i, 0)
        _style_tbl_cell(cell, label,
                        bg=C_LBGREY if i % 2 == 0 else C_WHITE,
                        fg=C_GREY, fsize=10, align=PP_ALIGN.LEFT, font=font,
                        ml=0.10, mr=0.05)

    # 3. Scatter chart (larger markers, matching template marker_size=7)
    chart_top = _FR_TOP + 0.02  # slight offset matching template

    # Vertical tick lines behind chart
    for tick in ticks:
        if x_min * 100 <= tick <= x_max * 100:
            frac = max(0.0, min(1.0, (tick - x_min * 100) / ((x_max - x_min) * 100)))
            tx = _FR_CHART_L + frac * _FR_CHART_W
            dashed_separator(slide, tx, chart_top, chart_h,
                             color=C_LBGREY, width_pt=0.5, vertical=True)

    _abs_scatter(slide, pri_vals if has_prior else None, cur_vals, n,
                 _FR_CHART_L, chart_top, _FR_CHART_W, chart_h,
                 color_prior, color_current, x_min, x_max,
                 data_labels=True)

    # 4. "vs Q4'25" delta header
    delta_header = _resolve_template(extra.get("delta_header", "vs {{period_prior}}"), config)
    cur_label = extra.get("current_label", "J&J")
    pri_label = extra.get("prior_label", "AZ")
    textbox(slide, delta_header,
            _FR_DELTA_L - 0.20, _FR_VS_HDR_TOP, _FR_DELTA_W + 0.40, 0.28,
            fsize=11, bold=True, color=C_GREY, align=PP_ALIGN.CENTER, font=font)

    # 5. Delta table (2 columns: J&J delta, AZ delta — 9pt, color-coded)
    _, dtbl = _pptx_table(slide, [_FR_DELTA_COL, _FR_DELTA_COL], [0.28] + [_FR_ROW_H] * n,
                                _FR_DELTA_L, _FR_TOP - 0.28)
    # Headers
    _style_tbl_cell(dtbl.cell(0, 0), cur_label, bg=C_HDRGREY, fg=C_WHITE,
                    fsize=8, bold=True, align=PP_ALIGN.CENTER, font=font)
    _style_tbl_cell(dtbl.cell(0, 1), pri_label, bg=C_HDRGREY, fg=C_WHITE,
                    fsize=8, bold=True, align=PP_ALIGN.CENTER, font=font)

    # Delta values
    for i in range(n):
        for ci, dval in enumerate([delta_cur[i], delta_pri[i]]):
            cell = dtbl.cell(i + 1, ci)
            if dval is not None:
                sign = "+" if dval > 0 else ""
                txt = f"{sign}{dval:.0f}%"
                fg = C_GREEN if dval > 0 else (C_RED if dval < 0 else C_GREY)
            else:
                txt = "\u2014"
                fg = C_GREY
            _style_tbl_cell(cell, txt,
                            bg=C_LBGREY if i % 2 == 0 else C_WHITE,
                            fg=fg, fsize=9, bold=True, align=PP_ALIGN.CENTER, font=font)

    # 6. "% of Interactions" axis label
    textbox(slide, "% of Interactions",
            _FR_CHART_L, _FR_AX_LBL_TOP, _FR_CHART_W, 0.27,
            fsize=10, color=C_GREY, align=PP_ALIGN.CENTER, font=font)

    # 7. Scale tick labels
    lbl_y = _FR_TOP + body_h + 0.06
    for tick in ticks:
        if x_min * 100 <= tick <= x_max * 100:
            frac = max(0.0, min(1.0, (tick - x_min * 100) / ((x_max - x_min) * 100)))
            tx = _FR_CHART_L + frac * _FR_CHART_W
            textbox(slide, f"{tick}{scale_sfx}", tx - 0.22, lbl_y,
                    0.44, 0.20, fsize=8, color=C_FTGREY,
                    align=PP_ALIGN.CENTER, font=font)

    # 8. Legend
    ly = _FR_AX_LBL_TOP + 0.30
    legend_items = [
        (color_current, extra.get("legend_current", cur_label)),
        (color_prior, extra.get("legend_prior", pri_label)),
    ]
    _make_legend(slide, legend_items, 0, ly, font, center_over=(0, _ABS_SLIDE_W))


# ══════════════════════════════════════════════════════════════════════════════
# Dual Abacus — two side-by-side abacus panels on one slide
# ══════════════════════════════════════════════════════════════════════════════

# Dual abacus layout constants
_DA_HDR_H       = _ABS_HDR_H
_DA_LABEL_W     = 2.60              # shared label column width
_DA_VAL_W       = 0.50              # value column width (narrower)
_DA_CHART_W     = 2.20              # scatter chart width per panel
_DA_DELTA_W     = 0.50              # delta column width
_DA_GAP         = 0.06              # gap between elements
_DA_SEP_GAP     = 0.25              # gap between panels (dashed separator)
_DA_PANEL_HDR_H = 0.30              # panel brand header height


def render_dual_abacus(slide, config: ProjectConfig, ask: AskConfig, data: dict, *, namer=None):
    """Dual abacus chart: two side-by-side XY scatter panels on one slide.

    Config (ask.extra):
        left:
            data_key: extraction id for left panel
            current_field / prior_field: field names (default "current"/"prior")
            current_label / prior_label: column headers
            color_current / color_prior: hex colors
            label: panel header text (e.g. "RYB+LAZ")
        right:
            (same keys as left)
        scale_min, scale_max, scale_ticks, scale_suffix: shared scale
        delta_header: header for gap column (default "Gap (pp)")
    """
    _slide_chrome(slide, config, ask)

    extra = ask.extra or {}
    font  = config.font_body

    left_cfg  = extra.get("left", {})
    right_cfg = extra.get("right", {})

    # Load and filter data for each panel
    def _load_panel(pcfg):
        dk = pcfg.get("data_key", ask.data_key)
        rows = data.get(dk, [])
        cf = pcfg.get("current_field", "current")
        pf = pcfg.get("prior_field", "prior")
        rows = _sort_data(rows, ask.sort_by, ask.sort_desc)
        rows = [r for r in rows if r.get(cf) not in (None, 0)]
        labels = [r.get("short") or r.get("desc", "") for r in rows]
        cur = [r.get(cf) or 0 for r in rows]
        pri = [r.get(pf) for r in rows]
        return labels, cur, pri

    l_labels, l_cur, l_pri = _load_panel(left_cfg)
    r_labels, r_cur, r_pri = _load_panel(right_cfg)

    # Use the longer panel's row count for consistent layout
    n = max(len(l_labels), len(r_labels))
    if n == 0:
        textbox(slide, "Data not available", 2, 3, 8, 1, fsize=14, color=C_RED)
        return

    # Pad shorter panel
    while len(l_labels) < n:
        l_labels.append(""); l_cur.append(None); l_pri.append(None)
    while len(r_labels) < n:
        r_labels.append(""); r_cur.append(None); r_pri.append(None)

    l_has_prior = any(p is not None for p in l_pri)
    r_has_prior = any(p is not None for p in r_pri)

    # Colors
    def _panel_colors(pcfg, brand_key):
        cc = parse_color(pcfg["color_current"]) if "color_current" in pcfg else None
        cp = parse_color(pcfg["color_prior"]) if "color_prior" in pcfg else None
        if cc is None or cp is None:
            b = config.brands.get(brand_key, config.primary)
            cc = cc or b.color_current
            cp = cp or b.color_prior
        return cc, cp

    l_cc, l_cp = _panel_colors(left_cfg, left_cfg.get("brand", "primary"))
    r_cc, r_cp = _panel_colors(right_cfg, right_cfg.get("brand", "competitor"))

    # Scale
    scale_min = int(extra.get("scale_min", 0))
    scale_max = int(extra.get("scale_max", 100))
    scale_sfx = extra.get("scale_suffix", "%")
    ticks = extra.get("scale_ticks") or [scale_min, (scale_min + scale_max) // 2, scale_max]

    all_vals = [v for v in l_cur + l_pri + r_cur + r_pri if v is not None and v != 0]
    x_min, x_max = _abs_x_range(all_vals, scale_min, scale_max)

    # Row heights
    row_h  = min(_ABS_ROW_H_MAX, max(_ABS_ROW_H_MIN, 4.20 / max(n, 1)))
    body_h = n * row_h
    total_h = _DA_HDR_H + body_h

    # Panel width calculation
    hide_val = extra.get("hide_val_cols", False)

    if hide_val:
        da_chart_w = _DA_CHART_W + 2 * _DA_VAL_W + _DA_GAP  # absorb val col space
        panel_w_fn = lambda nc: da_chart_w + _DA_DELTA_W + 2 * _DA_GAP
    else:
        da_chart_w = _DA_CHART_W
        panel_w_fn = lambda nc: nc * _DA_VAL_W + da_chart_w + _DA_DELTA_W + (nc + 1) * _DA_GAP

    l_ncols = 2 if l_has_prior else 1
    r_ncols = 2 if r_has_prior else 1
    l_panel_w = panel_w_fn(l_ncols)
    r_panel_w = panel_w_fn(r_ncols)
    total_w = _DA_LABEL_W + _DA_GAP + l_panel_w + _DA_SEP_GAP + r_panel_w

    # Center everything on the 13.33" slide
    origin = (_ABS_SLIDE_W - total_w) / 2
    label_l = origin
    l_start = label_l + _DA_LABEL_W + _DA_GAP

    # Left panel element positions
    if hide_val:
        l_chart_l = l_start + _DA_GAP
    else:
        l_pri_l  = l_start
        l_cur_l  = (l_pri_l + _DA_VAL_W) if l_has_prior else l_pri_l
        l_chart_l = l_cur_l + _DA_VAL_W + _DA_GAP
    l_delta_l = l_chart_l + da_chart_w + _DA_GAP

    sep_x = l_start + l_panel_w + _DA_SEP_GAP / 2

    r_start = l_start + l_panel_w + _DA_SEP_GAP
    if hide_val:
        r_chart_l = r_start + _DA_GAP
    else:
        r_pri_l  = r_start
        r_cur_l  = (r_pri_l + _DA_VAL_W) if r_has_prior else r_pri_l
        r_chart_l = r_cur_l + _DA_VAL_W + _DA_GAP
    r_delta_l = r_chart_l + da_chart_w + _DA_GAP

    # Panel brand headers sit between section bar and table headers
    # Section bar bottom ≈ 1.40 + 0.30 = 1.70; push headers to 1.72
    hdr_top = 1.72
    panel_top = hdr_top + _DA_PANEL_HDR_H + 0.02  # tables start below headers
    chart_top = panel_top + _DA_HDR_H

    l_label_text = left_cfg.get("label", "Left")
    r_label_text = right_cfg.get("label", "Right")
    textbox(slide, l_label_text, l_start, hdr_top, l_panel_w, _DA_PANEL_HDR_H,
            fsize=9, bold=True, color=l_cc, align=PP_ALIGN.CENTER, font=font)
    textbox(slide, r_label_text, r_start, hdr_top, r_panel_w, _DA_PANEL_HDR_H,
            fsize=9, bold=True, color=r_cc, align=PP_ALIGN.CENTER, font=font)

    # Recalculate body_h with adjusted top
    total_h = _DA_HDR_H + body_h

    # Shared label column
    _abs_label_tbl(slide, l_labels, label_l, panel_top, _DA_LABEL_W,
                   _DA_HDR_H, row_h, "Attribute", font)

    # ── Left panel ──
    if not hide_val:
        l_cur_hdr = left_cfg.get("current_label", "Acad")
        l_pri_hdr = left_cfg.get("prior_label", "Comm")
        if l_has_prior:
            _abs_val_col(slide, l_pri, l_pri_l, panel_top, _DA_VAL_W,
                         _DA_HDR_H, row_h, l_pri_hdr, l_cp, font)
        _abs_val_col(slide, l_cur, l_cur_l, panel_top, _DA_VAL_W,
                     _DA_HDR_H, row_h, l_cur_hdr, l_cc, font)

    # Left tick lines
    for tick in ticks:
        if x_min * 100 <= tick <= x_max * 100:
            frac = max(0.0, min(1.0, (tick - x_min * 100) / ((x_max - x_min) * 100)))
            tx = l_chart_l + frac * da_chart_w
            dashed_separator(slide, tx, chart_top, body_h,
                             color=C_LBGREY, width_pt=0.5, vertical=True)

    _abs_scatter(slide, l_pri if l_has_prior else None, l_cur, n,
                 l_chart_l, chart_top, da_chart_w, body_h,
                 l_cp, l_cc, x_min, x_max, data_labels=True)

    l_delta_hdr = extra.get("delta_header", "Gap (pp)")
    l_deltas = [
        round(l_cur[i] - l_pri[i], 1) if l_cur[i] is not None and l_pri[i] is not None else None
        for i in range(n)
    ]
    add_delta_table(slide, l_deltas, l_delta_l, panel_top, _DA_DELTA_W, row_h,
                    header_text=l_delta_hdr, font_name=font)

    # ── Vertical separator ──
    dashed_separator(slide, sep_x, panel_top, total_h,
                     color=C_GREY, width_pt=0.75, vertical=True)

    # ── Right panel ──
    if not hide_val:
        r_cur_hdr = right_cfg.get("current_label", "Acad")
        r_pri_hdr = right_cfg.get("prior_label", "Comm")
        if r_has_prior:
            _abs_val_col(slide, r_pri, r_pri_l, panel_top, _DA_VAL_W,
                         _DA_HDR_H, row_h, r_pri_hdr, r_cp, font)
        _abs_val_col(slide, r_cur, r_cur_l, panel_top, _DA_VAL_W,
                     _DA_HDR_H, row_h, r_cur_hdr, r_cc, font)

    # Right tick lines
    for tick in ticks:
        if x_min * 100 <= tick <= x_max * 100:
            frac = max(0.0, min(1.0, (tick - x_min * 100) / ((x_max - x_min) * 100)))
            tx = r_chart_l + frac * da_chart_w
            dashed_separator(slide, tx, chart_top, body_h,
                             color=C_LBGREY, width_pt=0.5, vertical=True)

    _abs_scatter(slide, r_pri if r_has_prior else None, r_cur, n,
                 r_chart_l, chart_top, da_chart_w, body_h,
                 r_cp, r_cc, x_min, x_max, data_labels=True)

    r_deltas = [
        round(r_cur[i] - r_pri[i], 1) if r_cur[i] is not None and r_pri[i] is not None else None
        for i in range(n)
    ]
    add_delta_table(slide, r_deltas, r_delta_l, panel_top, _DA_DELTA_W, row_h,
                    header_text=l_delta_hdr, font_name=font)

    # Scale labels below charts
    lbl_y = chart_top + body_h + 0.04
    for chart_left in (l_chart_l, r_chart_l):
        for tick in ticks:
            if x_min * 100 <= tick <= x_max * 100:
                frac = max(0.0, min(1.0, (tick - x_min * 100) / ((x_max - x_min) * 100)))
                tx = chart_left + frac * da_chart_w
                textbox(slide, f"{tick}{scale_sfx}", tx - 0.18, lbl_y,
                        0.36, 0.18, fsize=6, color=C_FTGREY,
                        align=PP_ALIGN.CENTER, font=font)

    # Legend with sample sizes
    _l_cur_hdr = left_cfg.get("current_label", "Acad")
    _l_pri_hdr = left_cfg.get("prior_label", "Com")
    legend_items = [
        (l_cc, left_cfg.get("legend_current", _l_cur_hdr)),
        (l_cp, left_cfg.get("legend_prior", _l_pri_hdr)),
    ]
    _make_legend(slide, legend_items, 0, panel_top + total_h + LEGEND_GAP + 0.10, font,
                 center_over=(0, _ABS_SLIDE_W))


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
