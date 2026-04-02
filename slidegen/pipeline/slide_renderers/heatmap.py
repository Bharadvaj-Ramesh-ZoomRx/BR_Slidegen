"""
heatmap.py — Heatmap table renderer.

Slide type:
  - heatmap_table   (template slide 41 — message recall by channel with green gradient)
"""

from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from ._shared import (
    SLIDE_W, CHART_TOP_STD, FOOTER_TOP, _vcenter_top,
    C_WHITE, C_GREY, C_FTGREY, C_LBGREY, C_HDRGREY, C_GREEN, C_RED,
    FONT_TEXT,
    _resolve_template, _slide_chrome,
    _pptx_table, _style_tbl_cell, _cell_bottom_border,
    textbox, solidrect,
    ProjectConfig, AskConfig, parse_color,
    logger,
)

# Template-matched colors
_C_ROW_EVEN = RGBColor(0xDC, 0xF4, 0xC4)   # light green alternating row
_C_ROW_ODD  = RGBColor(0xFD, 0xFF, 0xC3)   # light yellow alternating row
_C_DELTA_GREEN = RGBColor(0x00, 0x54, 0x26) # dark green for positive deltas
_C_DELTA_RED   = RGBColor(0xFF, 0x00, 0x00) # red for negative deltas
_DELTA_THRESHOLD = 5  # deltas > this get colored


def _heatmap_color(value: float, vmin: float, vmax: float,
                    low_rgb: tuple = (224, 242, 229),
                    high_rgb: tuple = (99, 190, 123)) -> RGBColor:
    """Interpolate between low_rgb (light) and high_rgb (dark) based on value."""
    if vmax == vmin:
        t = 0.5
    else:
        t = max(0, min(1, (value - vmin) / (vmax - vmin)))
    r = int(low_rgb[0] + t * (high_rgb[0] - low_rgb[0]))
    g = int(low_rgb[1] + t * (high_rgb[1] - low_rgb[1]))
    b = int(low_rgb[2] + t * (high_rgb[2] - low_rgb[2]))
    return RGBColor(r, g, b)


def _delta_color(d: float, threshold: int = _DELTA_THRESHOLD) -> RGBColor:
    """Return colored font for delta: green if positive > threshold, red if negative > threshold."""
    if d > threshold:
        return _C_DELTA_GREEN
    elif d < -threshold:
        return _C_DELTA_RED
    return C_GREY


def render_heatmap_table(slide, config: ProjectConfig, ask: AskConfig,
                          data: dict, *, namer=None):
    """Render a heatmap table with green gradient fills and QoQ delta columns.

    Data format (mock rows): list of dicts:
        {
            "label": "Indication",
            "values": {"Websites": 47, "Emails": 40, "Online Ads": 50, ...},
            "deltas": {"Websites": 2, "Emails": -3, "Online Ads": 7, ...}
        }

    ask.extra keys:
        columns:         ["Websites", "Emails", "Online Ads", ...]
        sample_sizes:    {"Websites": "19*", "Emails": "65", ...}
        label_header:    "Tag^"
        sample_header:   "s"
        show_deltas:     true
        value_suffix:    "%"
        delta_suffix:    "%"
        delta_threshold: 5    (abs delta > this gets colored green/red)
        heatmap_min:     0
        heatmap_max:     75
        low_color:       "#E0F2E5"
        high_color:      "#63BE7B"
    """
    _slide_chrome(slide, config, ask)
    extra = ask.extra or {}

    raw_rows = data.get(ask.data_key, [])
    if not raw_rows:
        logger.warning("heatmap_table: no data for key=%s", ask.data_key)
        return

    columns_cfg = extra.get("columns", [])
    if not columns_cfg:
        logger.warning("heatmap_table: no columns specified in extra")
        return

    # columns can be list of strings (legacy) or list of dicts with field/label
    if columns_cfg and isinstance(columns_cfg[0], dict):
        col_fields = [c["field"] for c in columns_cfg]
        columns = [c.get("label", c["field"]) for c in columns_cfg]
    else:
        col_fields = columns_cfg
        columns = columns_cfg

    # Auto-convert flat data ({desc, field1, field2, ...}) to heatmap format
    # ({label, values: {col: val}, deltas: {col: delta}})
    rows = []
    if raw_rows and "values" not in raw_rows[0]:
        for r in raw_rows:
            vals = {}
            for cf, cn in zip(col_fields, columns):
                v = r.get(cf)
                if v is not None:
                    vals[cn] = v
            rows.append({"label": r.get("desc", ""), "values": vals, "deltas": {}})
    else:
        rows = raw_rows

    sample_sizes = extra.get("sample_sizes", {})
    # Auto-derive sample sizes: _bases (from Excel) → sample_size_key (config) → sample_n (explicit)
    if not sample_sizes:
        _bases = data.get("_bases", {})
        _base = _bases.get(ask.data_key, {})
        _ss_key = extra.get("sample_size_key")  # e.g. "impact", "primary"
        _ss_n = extra.get("sample_n")            # e.g. 60 (explicit override)
        if _base and _base.get("current"):
            # From extraction's question code header row (most authoritative)
            sample_sizes = {col: f"n={_base['current']}" for col in columns}
        elif _ss_n:
            # Explicit N from config (for mock data where Excel base isn't available)
            sample_sizes = {col: f"n={_ss_n}" for col in columns}
        elif _ss_key and hasattr(config, 'sample_sizes'):
            _ss = config.sample_sizes.get(_ss_key)
            if _ss is not None:
                n_val = _ss.current if hasattr(_ss, 'current') else _ss.get('current')
                if n_val:
                    sample_sizes = {col: f"n={n_val}" for col in columns}
    label_header = extra.get("label_header", "Tag^")
    sample_header = extra.get("sample_header", "N" if sample_sizes else "s")
    show_deltas = extra.get("show_deltas", True)
    value_suffix = extra.get("value_suffix", "%")
    delta_suffix = extra.get("delta_suffix", "%")
    delta_threshold = extra.get("delta_threshold", _DELTA_THRESHOLD)
    hm_min = extra.get("heatmap_min", 0)
    hm_max = extra.get("heatmap_max", 75)
    font = config.font_body or FONT_TEXT

    low_rgb = _parse_rgb_tuple(extra.get("low_color", "#E0F2E5"))
    high_rgb = _parse_rgb_tuple(extra.get("high_color", "#63BE7B"))

    n_cols = len(columns)
    n_rows = len(rows)

    # ── Layout — dynamic dimensions based on row count ──
    label_col_w = extra.get("label_col_w", 3.00)
    value_col_w = 1.54
    delta_col_w = 0.63
    hdr_h = 0.34
    has_sample = bool(sample_sizes)
    sample_h = 0.24 if has_sample else 0.0

    # Dynamic row height: scale down for many rows to fit within slide
    available_h = FOOTER_TOP - CHART_TOP_STD - 0.60  # usable vertical space
    max_body_h = available_h - hdr_h - sample_h
    row_h = min(0.35, max(0.18, max_body_h / max(n_rows, 1)))

    # Total width of unified table: label + n_cols × (value + delta)
    if show_deltas:
        total_w = label_col_w + n_cols * (value_col_w + delta_col_w)
    else:
        total_w = label_col_w + n_cols * value_col_w

    # Center horizontally and vertically on slide
    table_left = (SLIDE_W - total_w) / 2
    content_h = hdr_h + sample_h + n_rows * row_h
    table_top = _vcenter_top(content_h, has_legend=False)

    # ── Build column widths list ──
    col_widths = [label_col_w]
    for ci in range(n_cols):
        col_widths.append(value_col_w)
        if show_deltas:
            col_widths.append(delta_col_w)
    total_cols = len(col_widths)

    # ── Header table (header row + optional sample row) ──
    hdr_rows = [hdr_h] + ([sample_h] if has_sample else [])
    hdr_shape, hdr_tbl = _pptx_table(
        slide, col_widths, hdr_rows,
        table_left, table_top)

    # Adaptive font size for header based on column count
    hdr_fsize = 8 if n_cols <= 4 else 7

    # Header row
    _style_tbl_cell(hdr_tbl.cell(0, 0), label_header,
                    bg=C_HDRGREY, fg=C_WHITE, fsize=hdr_fsize, bold=True,
                    align=PP_ALIGN.CENTER, font=font)

    if has_sample:
        _style_tbl_cell(hdr_tbl.cell(1, 0), sample_header,
                        bg=C_WHITE, fg=C_FTGREY, fsize=hdr_fsize, bold=True,
                        align=PP_ALIGN.CENTER, font=font)

    for ci, col_name in enumerate(columns):
        val_ci = 1 + ci * (2 if show_deltas else 1)
        _style_tbl_cell(hdr_tbl.cell(0, val_ci), col_name,
                        bg=C_HDRGREY, fg=C_WHITE, fsize=hdr_fsize, bold=True,
                        align=PP_ALIGN.CENTER, font=font)
        if has_sample:
            ss = sample_sizes.get(col_name, "")
            _style_tbl_cell(hdr_tbl.cell(1, val_ci), str(ss),
                            bg=C_WHITE, fg=C_FTGREY, fsize=hdr_fsize, bold=False,
                            align=PP_ALIGN.CENTER, font=font)

        if show_deltas:
            delta_ci = val_ci + 1
            _style_tbl_cell(hdr_tbl.cell(0, delta_ci), "",
                            bg=C_HDRGREY, fg=C_WHITE, fsize=7, bold=False,
                            align=PP_ALIGN.CENTER, font=font)
            if has_sample:
                _style_tbl_cell(hdr_tbl.cell(1, delta_ci), "",
                                bg=C_WHITE, fg=C_FTGREY, fsize=7, bold=False,
                                align=PP_ALIGN.CENTER, font=font)

    # ── Data table ──
    data_top = table_top + hdr_h + sample_h
    row_heights = [row_h] * n_rows
    data_shape, data_tbl = _pptx_table(
        slide, col_widths, row_heights,
        table_left, data_top)

    # Adaptive font sizes based on row count (dense tables get smaller text)
    label_fsize = 9 if n_rows <= 25 else (8 if n_rows <= 35 else 7)
    val_fsize = 9 if n_rows <= 25 else (8 if n_rows <= 35 else 7)

    # Collect all values for heatmap range
    all_values = []
    for row in rows:
        vals = row.get("values", {})
        for col in columns:
            v = vals.get(col)
            if v is not None:
                all_values.append(v)
    if all_values:
        if hm_min == 0 and hm_max == 75:
            hm_min = min(all_values)
            hm_max = max(all_values)

    for ri, row in enumerate(rows):
        label = row.get("label", "")
        vals = row.get("values", {})
        deltas = row.get("deltas", {})

        # Alternating row backgrounds (green / yellow like template)
        row_bg = _C_ROW_EVEN if ri % 2 == 0 else _C_ROW_ODD
        fc = RGBColor(0x00, 0x00, 0x00) if ri % 2 == 0 else RGBColor(0x33, 0x33, 0x33)

        # Label cell
        _style_tbl_cell(data_tbl.cell(ri, 0), label,
                        bg=row_bg, fg=fc, fsize=label_fsize, bold=False,
                        align=PP_ALIGN.LEFT, font=font, ml=0.08)

        for ci, col in enumerate(columns):
            val_ci = 1 + ci * (2 if show_deltas else 1)

            # Value cell with heatmap fill
            v = vals.get(col)
            if v is not None:
                fill = _heatmap_color(v, hm_min, hm_max, low_rgb, high_rgb)
                text = f"{v:.0f}{value_suffix}" if isinstance(v, float) else f"{v}{value_suffix}"
            else:
                fill = row_bg
                text = "-"
            _style_tbl_cell(data_tbl.cell(ri, val_ci), text,
                            bg=fill, fg=fc, fsize=val_fsize, bold=False,
                            align=PP_ALIGN.CENTER, font=font)

            # Delta cell
            if show_deltas:
                delta_ci = val_ci + 1
                d = deltas.get(col)
                if d is not None and d != 0:
                    sign = "+" if d > 0 else ""
                    d_text = f"{sign}{d:.0f}{delta_suffix}" if isinstance(d, float) else f"{sign}{d}{delta_suffix}"
                    d_color = _delta_color(d, delta_threshold)
                elif d == 0:
                    d_text = f"0{delta_suffix}"
                    d_color = C_GREY
                else:
                    d_text = "-"
                    d_color = C_GREY

                _style_tbl_cell(data_tbl.cell(ri, delta_ci), d_text,
                                bg=None, fg=d_color, fsize=val_fsize - 1, bold=False,
                                align=PP_ALIGN.CENTER, font=font)


def _parse_rgb_tuple(hex_color: str) -> tuple:
    """Parse '#RRGGBB' into (R, G, B) int tuple."""
    h = hex_color.lstrip("#")
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
