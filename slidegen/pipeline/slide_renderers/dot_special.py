"""
dot_special.py — Follow-up rep, dual abacus, and message MBD renderers.

Split from dot.py for maintainability. These are specialized variants of
the core abacus scatter chart layout.
"""

from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from ._shared import (
    CHART_TOP_STD, LEGEND_GAP, SLIDE_W,
    _ABS_HDR_H, _ABS_ROW_H_MIN, _ABS_ROW_H_MAX, _ABS_SLIDE_W,
    LABEL_MAX_DUAL,
    _resolve_template, _slide_chrome, _get_brand_colors, _sort_data, _make_legend,
    _pptx_table, _style_tbl_cell, _prepare_rows,
    _alt_row_bg, _no_data_placeholder,
    C_GREEN, C_GREY, C_FTGREY, C_LBGREY, C_HDRGREY, C_WHITE, C_RED,
    FONT_TEXT,
    textbox, dashed_separator,
    add_delta_table,
    ProjectConfig, AskConfig, parse_color,
)
from .dot import (
    _abs_x_range, _abs_scatter, _abs_label_tbl, _abs_val_col,
)


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
        _no_data_placeholder(slide, ask.id)
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
                        bg=_alt_row_bg(i),
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
                            bg=_alt_row_bg(i),
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
        _no_data_placeholder(slide, ask.id)
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
    hdr_top = 1.72
    panel_top = hdr_top + _DA_PANEL_HDR_H + 0.02
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

    rows = _prepare_rows(slide, data, ask)
    if rows is None:
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
