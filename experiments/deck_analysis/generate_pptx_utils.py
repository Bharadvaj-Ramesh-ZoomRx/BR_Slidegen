"""
Generate pptx_utils Python from Deck Analysis
==============================================

Reads the structured JSON outputs from deep_analyzer.py and layout_clusterer.py
and produces Python source for pptx_utils modules:

    outputs/generated_brand.py    — BRAND{} per client
    outputs/generated_layouts.py  — LAYOUTS{} from top coordinate clusters
    outputs/generated_lxml_helpers.py  — new helper functions with OOXML knowledge
    outputs/generated_chart_patterns.py — CHART_PATTERNS{} per chart type

Run after deep_analyzer.py and layout_clusterer.py have produced their outputs.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

HERE = Path(__file__).parent
OUTPUTS_DIR = HERE / "outputs"


def hex_to_rgb_tuple(hex_str: str) -> tuple[int, int, int]:
    """'FF9933' or '#FF9933' -> (0xFF, 0x99, 0x33)"""
    h = hex_str.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def client_key(client: str) -> str:
    """Normalize client name for Python dict key."""
    return client.upper().replace(" ", "_").replace("-", "_")


# ---------------------------------------------------------------------------
# BRAND{} generator
# ---------------------------------------------------------------------------

# Manual overrides where the observed series color isn't the "real" brand primary
# (e.g., a deck might use green for delta rather than the brand's actual color).
BRAND_PRIMARY_OVERRIDES = {
    "JJ": "0063C3",      # JJ Blue (official) — not the '7FB1E1' tint that showed up first
    "GSK": "F36633",     # GSK Orange — override the green which is contextual
}

BRAND_HEADING_COLOR = {
    "JJ": "001E60",
    "AZN": "595959",
    "Pfizer": "0063C3",
    "Amgen": "003C71",
    "Regeneron": "00745A",
    "LEO": "C014A3",
    "Alexion": "0E8779",
    "BL": "400286",
    "Novartis": "018E86",
    "Apellis": "FC3B6E",
    "DSI": "1E22AA",
    "GSK": "F36633",
    "Otsuka": "000000",
    "Ipsen": "54AC65",
    "CCA": "F28E2B",
    "Bone-HCP": "FF9933",
}


def generate_brand_py() -> str:
    data = json.loads((OUTPUTS_DIR / "deep_brand_proposals.json").read_text(encoding="utf-8"))

    lines: list[str] = []
    lines.append('"""')
    lines.append("Brand definitions per pharmaceutical client.")
    lines.append("")
    lines.append("Generated from real-deck analysis of 32 PET decks across 17 clients.")
    lines.append("See experiments/deck_analysis/outputs/deep_brand_proposals.md for")
    lines.append("the observation data that drove these constants.")
    lines.append("")
    lines.append("Each BRAND entry provides:")
    lines.append("  - primary:      main series color (used for 'current wave' bars)")
    lines.append("  - prior:        secondary/tint color (used for 'prior wave' bars)")
    lines.append("  - secondary:    additional accent for comparisons")
    lines.append("  - positive:     delta positive (universal green across decks)")
    lines.append("  - negative:     delta negative (universal red across decks)")
    lines.append("  - heading_color: headline text color")
    lines.append("  - font_heading: headline font")
    lines.append("  - font_body:    body text font")
    lines.append("")
    lines.append("Growth model: new clients are added to BRAND{} as their first engagement")
    lines.append("uses the system. Existing entries are updated when brand guides change.")
    lines.append('"""')
    lines.append("from __future__ import annotations")
    lines.append("")
    lines.append("from pptx.dml.color import RGBColor")
    lines.append("")
    lines.append("")
    lines.append("# Universal delta colors — observed across all 32 decks")
    lines.append("POSITIVE_GREEN = RGBColor(0x00, 0xB0, 0x50)  # 494 occurrences")
    lines.append("NEGATIVE_RED = RGBColor(0xFF, 0x00, 0x00)    # 913 occurrences")
    lines.append("NEGATIVE_DEEP_RED = RGBColor(0xC0, 0x00, 0x00)  # 61 occurrences (alt)")
    lines.append("")
    lines.append("# Standard greys (appear across all decks)")
    lines.append("GREY_DARK = RGBColor(0x40, 0x40, 0x40)")
    lines.append("GREY_MID = RGBColor(0x59, 0x59, 0x59)")
    lines.append("GREY_LIGHT = RGBColor(0xBF, 0xBF, 0xBF)")
    lines.append("GREY_ALT_ROW = RGBColor(0xF2, 0xF2, 0xF2)  # Standard table alt-row fill")
    lines.append("")
    lines.append("")
    lines.append("BRAND = {")

    # Sort by deck count, descending
    sorted_clients = sorted(
        data.items(),
        key=lambda kv: -kv[1].get("deck_count", 0)
    )

    for client, info in sorted_clients:
        deck_count = info.get("deck_count", 0)
        top_colors = list(info.get("top_series_colors", {}).keys())
        top_fonts = list(info.get("top_fonts", {}).keys())
        theme_major = list(info.get("theme_fonts_major", {}).keys())
        theme_minor = list(info.get("theme_fonts_minor", {}).keys())

        if not top_colors:
            continue

        primary = BRAND_PRIMARY_OVERRIDES.get(client, top_colors[0])
        secondary = top_colors[1] if len(top_colors) > 1 else "808080"
        prior_tint = top_colors[2] if len(top_colors) > 2 else "BFBFBF"

        # Font selection: skip theme placeholders (+mn-lt, +mj-lt)
        real_fonts = [f for f in top_fonts if not f.startswith("+")]
        font_body = real_fonts[0] if real_fonts else "Arial"
        font_heading = theme_major[0] if theme_major else (real_fonts[0] if real_fonts else "Arial")
        if font_heading.startswith("+"):
            font_heading = font_body

        heading_color = BRAND_HEADING_COLOR.get(client, "000000")

        key = client_key(client)
        lines.append(f'    # --- {client} ({deck_count} deck{"s" if deck_count != 1 else ""}) ---')
        lines.append(f'    "{key}": {{')
        r, g, b = hex_to_rgb_tuple(primary)
        lines.append(f'        "primary":       RGBColor(0x{r:02X}, 0x{g:02X}, 0x{b:02X}),')
        r, g, b = hex_to_rgb_tuple(secondary)
        lines.append(f'        "secondary":     RGBColor(0x{r:02X}, 0x{g:02X}, 0x{b:02X}),')
        r, g, b = hex_to_rgb_tuple(prior_tint)
        lines.append(f'        "prior":         RGBColor(0x{r:02X}, 0x{g:02X}, 0x{b:02X}),')
        lines.append('        "positive":      POSITIVE_GREEN,')
        lines.append('        "negative":      NEGATIVE_RED,')
        r, g, b = hex_to_rgb_tuple(heading_color)
        lines.append(f'        "heading_color": RGBColor(0x{r:02X}, 0x{g:02X}, 0x{b:02X}),')
        lines.append(f'        "font_heading":  "{font_heading}",')
        lines.append(f'        "font_body":     "{font_body}",')
        # Observed palette for reference
        palette_hex = ", ".join(f'"#{c}"' for c in top_colors[:8])
        lines.append(f'        "_observed_palette": [{palette_hex}],  # for reference')
        lines.append('    },')
        lines.append('')

    lines.append("}")
    lines.append("")
    lines.append("")
    lines.append("def get_brand(client_key_or_alias: str) -> dict:")
    lines.append('    """Look up brand config by client key. Raises KeyError with available keys."""')
    lines.append("    key = client_key_or_alias.upper().replace(' ', '_').replace('-', '_')")
    lines.append("    if key not in BRAND:")
    lines.append('        raise KeyError(')
    lines.append('            f"Unknown client {client_key_or_alias!r}. Available: {sorted(BRAND.keys())}"')
    lines.append("        )")
    lines.append("    return BRAND[key]")
    lines.append("")
    lines.append("")
    lines.append("def get_color(client: str, role: str) -> RGBColor:")
    lines.append('    """Shortcut: get_color("JJ", "primary") -> RGBColor."""')
    lines.append("    return get_brand(client)[role]")
    lines.append("")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# LAYOUTS{} generator
# ---------------------------------------------------------------------------


def generate_layouts_py() -> str:
    chart_clusters = json.loads((OUTPUTS_DIR / "chart_positions.json").read_text(encoding="utf-8"))
    table_clusters = json.loads((OUTPUTS_DIR / "table_positions.json").read_text(encoding="utf-8"))
    sig_stats = json.loads((OUTPUTS_DIR / "layout_clusters.json").read_text(encoding="utf-8"))

    lines: list[str] = []
    lines.append('"""')
    lines.append("Layout presets for slide compositions.")
    lines.append("")
    lines.append("Generated from real-deck analysis — coordinate clusters of 410 distinct chart")
    lines.append("positions and 490 table positions across 2,333 slides.")
    lines.append("")
    lines.append("Each LAYOUTS entry is a dict of (left, top, width, height) in inches. Use in")
    lines.append("renderers as the source of truth for coordinate placement.")
    lines.append("")
    lines.append("Conventions:")
    lines.append("  - chart_rect: main chart position")
    lines.append("  - table_rect: main label/data table position")
    lines.append("  - delta_col_rect: narrow delta column (typically ~0.5\" wide)")
    lines.append("  - headline_rect: top-of-slide headline text box")
    lines.append('"""')
    lines.append("from __future__ import annotations")
    lines.append("")
    lines.append("")
    lines.append("# Slide canvas (widescreen 16:9)")
    lines.append("SLIDE_WIDTH = 13.333")
    lines.append("SLIDE_HEIGHT = 7.5")
    lines.append("")
    lines.append("# Universal headline position (from headline analysis — median top 0.3\",")
    lines.append("# median width 11.94\", median char count 72)")
    lines.append('HEADLINE_RECT = {"left": 0.2, "top": 0.3, "width": 12.8, "height": 0.9}')
    lines.append("")
    lines.append("# Universal footer position (used for source + footnotes)")
    lines.append('FOOTER_RECT = {"left": 0.2, "top": 7.0, "width": 12.8, "height": 0.4}')
    lines.append("")
    lines.append("")
    lines.append("# ============================================================================")
    lines.append("# LAYOUT PRESETS — by slide signature")
    lines.append("# ============================================================================")
    lines.append("")
    lines.append("LAYOUTS = {")
    lines.append("")

    # Generate from top signatures that have chart AND table
    def rect_str(left: float, top: float, width: float, height: float) -> str:
        return f'{{"left": {left}, "top": {top}, "width": {width}, "height": {height}}}'

    # 1_chart_1_table (145 slides) — single_bar_with_delta
    if "1_chart_1_table" in sig_stats:
        s = sig_stats["1_chart_1_table"]["stats"]
        lines.append("    # ---- single_bar_with_delta (1 chart + 1 table, 145 slides) ----")
        lines.append("    # Most common pattern — label table + bar chart side by side")
        lines.append('    "single_bar_with_delta": {')
        lines.append(f'        "chart_rect": {rect_str(s["chart"]["left"]["median"], s["chart"]["top"]["median"], s["chart"]["width"]["median"], s["chart"]["height"]["median"])},')
        lines.append(f'        "table_rect": {rect_str(s["table"]["left"]["median"], s["table"]["top"]["median"], s["table"]["width"]["median"], s["table"]["height"]["median"])},')
        lines.append('        "delta_col_rect": {"left": 12.52, "top": 2.17, "width": 0.58, "height": 4.5},  # from 40-occurrence cluster')
        lines.append(f'        "_source": "1_chart_1_table signature, {sig_stats["1_chart_1_table"]["slides"]} slides",')
        lines.append('    },')
        lines.append("")

    # 1_chart_2_table (110 slides) — clustered_compare
    if "1_chart_2_table" in sig_stats:
        s = sig_stats["1_chart_2_table"]["stats"]
        lines.append("    # ---- clustered_compare (1 chart + 2 tables, 110 slides) ----")
        lines.append("    # Label table + delta table side by side with chart on right")
        lines.append('    "clustered_compare": {')
        lines.append(f'        "chart_rect": {rect_str(s["chart"]["left"]["median"], s["chart"]["top"]["median"], s["chart"]["width"]["median"], s["chart"]["height"]["median"])},')
        lines.append(f'        "primary_table_rect": {rect_str(s["table"]["left"]["median"], s["table"]["top"]["median"], s["table"]["width"]["median"], s["table"]["height"]["median"])},')
        lines.append(f'        "secondary_table_rect": {rect_str(s["table"]["left"]["median"] + s["table"]["width"]["median"] + 0.1, s["table"]["top"]["median"], s["table"]["width"]["median"], s["table"]["height"]["median"])},')
        lines.append(f'        "_source": "1_chart_2_table signature, {sig_stats["1_chart_2_table"]["slides"]} slides",')
        lines.append('    },')
        lines.append("")

    # 2_chart_2_table (68 slides) — dual_bar_with_delta
    if "2_chart_2_table" in sig_stats:
        s = sig_stats["2_chart_2_table"]["stats"]
        lines.append("    # ---- dual_bar_with_delta (2 charts + 2 tables, 68 slides) ----")
        lines.append("    # Two brand comparison — two bars + two tables side by side")
        lines.append('    "dual_bar_with_delta": {')
        lines.append(f'        "left_chart_rect": {rect_str(s["table"]["left"]["median"], s["table"]["top"]["median"], s["chart"]["width"]["median"], s["chart"]["height"]["median"])},')
        lines.append(f'        "right_chart_rect": {rect_str(s["chart"]["left"]["median"], s["chart"]["top"]["median"], s["chart"]["width"]["median"], s["chart"]["height"]["median"])},')
        lines.append(f'        "left_table_rect": {rect_str(s["table"]["left"]["median"], s["table"]["top"]["median"], s["table"]["width"]["median"], s["table"]["height"]["median"])},')
        lines.append(f'        "right_table_rect": {rect_str(s["chart"]["left"]["median"] - s["table"]["width"]["median"] - 0.1, s["table"]["top"]["median"], s["table"]["width"]["median"], s["table"]["height"]["median"])},')
        lines.append(f'        "_source": "2_chart_2_table signature, {sig_stats["2_chart_2_table"]["slides"]} slides",')
        lines.append('    },')
        lines.append("")

    # 1_table (139 slides) — full-width table
    if "1_table" in sig_stats:
        s = sig_stats["1_table"]["stats"]
        lines.append("    # ---- full_width_table (1 table only, 139 slides) ----")
        lines.append("    # Large single table — Executive Summary, Recommendations")
        lines.append('    "full_width_table": {')
        lines.append(f'        "table_rect": {rect_str(s["table"]["left"]["median"], s["table"]["top"]["median"], s["table"]["width"]["median"], s["table"]["height"]["median"])},')
        lines.append(f'        "_source": "1_table signature, {sig_stats["1_table"]["slides"]} slides",')
        lines.append('    },')
        lines.append("")

    # 3_chart_1_table (45 slides) — scorecard
    if "3_chart_1_table" in sig_stats:
        s = sig_stats["3_chart_1_table"]["stats"]
        lines.append("    # ---- three_metric_scorecard (3 charts + 1 table, 45 slides) ----")
        lines.append("    # Three-panel scorecard comparing metrics side by side")
        lines.append('    "three_metric_scorecard": {')
        lines.append(f'        "label_table_rect": {rect_str(s["table"]["left"]["median"], s["table"]["top"]["median"], s["table"]["width"]["median"], s["table"]["height"]["median"])},')
        lines.append(f'        "chart_panel_template": {rect_str(s["chart"]["left"]["median"], s["chart"]["top"]["median"], s["chart"]["width"]["median"], s["chart"]["height"]["median"])},')
        lines.append('        "chart_panel_count": 3,')
        lines.append('        "chart_panel_gap": 0.1,')
        lines.append(f'        "_source": "3_chart_1_table signature, {sig_stats["3_chart_1_table"]["slides"]} slides",')
        lines.append('    },')
        lines.append("")

    # 2_chart (44 slides) — dual bar no tables
    if "2_chart" in sig_stats:
        s = sig_stats["2_chart"]["stats"]
        lines.append("    # ---- dual_chart_no_table (2 charts, 44 slides) ----")
        lines.append("    # Two charts only — comparison without label tables")
        lines.append('    "dual_chart_no_table": {')
        lines.append(f'        "left_chart_rect": {rect_str(s["chart"]["left"]["median"], s["chart"]["top"]["median"], s["chart"]["width"]["median"], s["chart"]["height"]["median"])},')
        lines.append(f'        "right_chart_rect": {rect_str(s["chart"]["left"]["median"] + s["chart"]["width"]["median"] + 0.3, s["chart"]["top"]["median"], s["chart"]["width"]["median"], s["chart"]["height"]["median"])},')
        lines.append(f'        "_source": "2_chart signature, {sig_stats["2_chart"]["slides"]} slides",')
        lines.append('    },')
        lines.append("")

    lines.append("}")
    lines.append("")
    lines.append("")
    lines.append("# ============================================================================")
    lines.append("# NARROW DELTA COLUMN POSITIONS — top 10 observed positions")
    lines.append("# ============================================================================")
    lines.append("#")
    lines.append("# These are recurring \"narrow column\" positions observed in real decks.")
    lines.append("# Typically ~0.5\" wide × ~4.5\" tall at specific left positions, used as")
    lines.append("# standalone delta columns adjacent to charts.")
    lines.append("")
    lines.append("DELTA_COL_POSITIONS = [")

    # Filter table clusters to narrow tall ones (<=1" wide, >=3" tall)
    narrow_tall = [
        c for c in table_clusters
        if c["width_median"] <= 1.0 and c["height_median"] >= 3.0
    ][:10]
    for c in narrow_tall:
        lines.append(f'    {{"left": {c["left_median"]}, "top": {c["top_median"]}, "width": {c["width_median"]}, "height": {c["height_median"]}, "_occurrences": {c["n"]}}},')

    lines.append("]")
    lines.append("")
    lines.append("")
    lines.append("def get_layout(slide_type: str) -> dict:")
    lines.append('    """Look up a layout preset by slide type."""')
    lines.append("    if slide_type not in LAYOUTS:")
    lines.append('        raise KeyError(')
    lines.append('            f"Unknown slide_type {slide_type!r}. Available: {sorted(LAYOUTS.keys())}"')
    lines.append("        )")
    lines.append("    return LAYOUTS[slide_type]")
    lines.append("")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# lxml_helpers.py generator
# ---------------------------------------------------------------------------


def generate_lxml_helpers_py() -> str:
    """Generate named helpers with OOXML knowledge encoded from the deck analysis."""
    lines: list[str] = []
    lines.append('"""')
    lines.append("OOXML helpers — raw XML manipulation for properties python-pptx does not expose.")
    lines.append("")
    lines.append("Generated from real-deck analysis of 32 PET decks. Each function wraps a")
    lines.append("specific OOXML property that appeared frequently enough to warrant a named helper.")
    lines.append("See experiments/deck_analysis/outputs/ACTIONABLE_FINDINGS.md §2 for the")
    lines.append("observed frequencies that justified each helper.")
    lines.append("")
    lines.append("All functions operate on python-pptx Chart or Series objects and mutate their")
    lines.append("underlying lxml element tree. No new presentation objects are created here.")
    lines.append('"""')
    lines.append("from __future__ import annotations")
    lines.append("")
    lines.append("from lxml import etree")
    lines.append("")
    lines.append("")
    lines.append('NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"')
    lines.append('NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"')
    lines.append('NSMAP = {"c": NS_C, "a": NS_A}')
    lines.append("")
    lines.append("")
    lines.append("def qn_c(local: str) -> str:")
    lines.append('    """Qualified chart namespace tag."""')
    lines.append('    return f"{{{NS_C}}}{local}"')
    lines.append("")
    lines.append("")
    lines.append("def qn_a(local: str) -> str:")
    lines.append('    """Qualified drawingML namespace tag."""')
    lines.append('    return f"{{{NS_A}}}{local}"')
    lines.append("")
    lines.append("")
    lines.append("def _find_or_create(parent, tag: str, insert_before: list[str] | None = None):")
    lines.append('    """Find child by qualified tag, or create it (inserted before listed siblings)."""')
    lines.append("    el = parent.find(tag)")
    lines.append("    if el is not None:")
    lines.append("        return el")
    lines.append("    el = etree.SubElement(parent, tag)")
    lines.append("    return el")
    lines.append("")
    lines.append("")
    lines.append("# ============================================================================")
    lines.append("# BAR / COLUMN GEOMETRY")
    lines.append("# ============================================================================")
    lines.append("")
    lines.append("def set_plot_area_gap(chart, gap_width: int = 80) -> None:")
    lines.append('    """Set <c:gapWidth> on bar/column plot.')
    lines.append("")
    lines.append("    gap_width: percentage of bar thickness; 0 = bars touching, 200 = gap equals")
    lines.append("    bar width. Observed range 40-150 in real decks, with 50/80/100 most common.")
    lines.append("    Default 80 matches the most common PET bar chart.")
    lines.append('    """')
    lines.append("    plot_area = chart._chartSpace.find(f\".//{qn_c('plotArea')}\")")
    lines.append("    if plot_area is None:")
    lines.append("        return")
    lines.append("    for tag in ('barChart', 'bar3DChart'):")
    lines.append("        for bar_chart in plot_area.findall(qn_c(tag)):")
    lines.append("            gw = bar_chart.find(qn_c('gapWidth'))")
    lines.append("            if gw is None:")
    lines.append("                gw = etree.SubElement(bar_chart, qn_c('gapWidth'))")
    lines.append("            gw.set('val', str(gap_width))")
    lines.append("")
    lines.append("")
    lines.append("def set_overlap(chart, overlap_pct: int = 0) -> None:")
    lines.append('    """Set <c:overlap> on bar/column plot.')
    lines.append("")
    lines.append("    overlap_pct: -100 to 100. 100 = fully overlapped (stacked). 0 = clustered")
    lines.append("    side by side. Negative = gap between bars within a cluster. Observed: 100")
    lines.append("    dominates (stacked charts), -20 next most common (slight separation).")
    lines.append('    """')
    lines.append("    plot_area = chart._chartSpace.find(f\".//{qn_c('plotArea')}\")")
    lines.append("    if plot_area is None:")
    lines.append("        return")
    lines.append("    for tag in ('barChart', 'bar3DChart'):")
    lines.append("        for bar_chart in plot_area.findall(qn_c(tag)):")
    lines.append("            ov = bar_chart.find(qn_c('overlap'))")
    lines.append("            if ov is None:")
    lines.append("                ov = etree.SubElement(bar_chart, qn_c('overlap'))")
    lines.append("            ov.set('val', str(overlap_pct))")
    lines.append("")
    lines.append("")
    lines.append("def set_invert_if_negative(chart, value: bool = False) -> None:")
    lines.append('    """Toggle <c:invertIfNegative> on all series. Default False disables the')
    lines.append("    auto-color-flip on negative values (observed: 13,550 series want this off).")
    lines.append('    """')
    lines.append("    for ser in chart._chartSpace.findall(f\".//{qn_c('ser')}\"):")
    lines.append("        iin = ser.find(qn_c('invertIfNegative'))")
    lines.append("        if iin is None:")
    lines.append("            iin = etree.SubElement(ser, qn_c('invertIfNegative'))")
    lines.append("        iin.set('val', '1' if value else '0')")
    lines.append("")
    lines.append("")
    lines.append("# ============================================================================")
    lines.append("# AXIS CONFIGURATION")
    lines.append("# ============================================================================")
    lines.append("")
    lines.append("def invert_cat_axis(chart) -> None:")
    lines.append('    """Set category axis orientation to maxMin (top-down).')
    lines.append("")
    lines.append("    For horizontal bar charts, this makes the first category appear at the top,")
    lines.append("    matching standard PET layout. Observed in 43% of charts (1,865 / 4,354).")
    lines.append('    """')
    lines.append("    _set_axis_orientation(chart, 'catAx', 'maxMin')")
    lines.append("")
    lines.append("")
    lines.append("def invert_val_axis(chart) -> None:")
    lines.append('    """Set value axis orientation to maxMin (rare — only 115 charts)."""')
    lines.append("    _set_axis_orientation(chart, 'valAx', 'maxMin')")
    lines.append("")
    lines.append("")
    lines.append("def _set_axis_orientation(chart, axis_tag: str, orientation: str) -> None:")
    lines.append("    for ax in chart._chartSpace.findall(f\".//{qn_c(axis_tag)}\"):")
    lines.append("        scaling = ax.find(qn_c('scaling'))")
    lines.append("        if scaling is None:")
    lines.append("            scaling = etree.SubElement(ax, qn_c('scaling'))")
    lines.append("        orient = scaling.find(qn_c('orientation'))")
    lines.append("        if orient is None:")
    lines.append("            orient = etree.SubElement(scaling, qn_c('orientation'))")
    lines.append("        orient.set('val', orientation)")
    lines.append("")
    lines.append("")
    lines.append("def hide_cat_labels(chart) -> None:")
    lines.append('    """Set <c:tickLblPos val="none"> on category axis.')
    lines.append("")
    lines.append("    Used when category labels appear in a companion table instead of on the axis.")
    lines.append("    Core clustered_compare pattern. Observed in 371 charts.")
    lines.append('    """')
    lines.append("    _set_tick_lbl_pos(chart, 'catAx', 'none')")
    lines.append("")
    lines.append("")
    lines.append("def hide_val_labels(chart) -> None:")
    lines.append('    """Hide value axis labels (less common, 48 charts)."""')
    lines.append("    _set_tick_lbl_pos(chart, 'valAx', 'none')")
    lines.append("")
    lines.append("")
    lines.append("def _set_tick_lbl_pos(chart, axis_tag: str, pos: str) -> None:")
    lines.append("    for ax in chart._chartSpace.findall(f\".//{qn_c(axis_tag)}\"):")
    lines.append("        tlp = ax.find(qn_c('tickLblPos'))")
    lines.append("        if tlp is None:")
    lines.append("            tlp = etree.SubElement(ax, qn_c('tickLblPos'))")
    lines.append("        tlp.set('val', pos)")
    lines.append("")
    lines.append("")
    lines.append("# ============================================================================")
    lines.append("# DATA LABELS")
    lines.append("# ============================================================================")
    lines.append("")
    lines.append("def set_datalabel_pos_center(chart) -> None:")
    lines.append('    """Center (ctr) — 4,508 occurrences, most common for stacked bars."""')
    lines.append("    _set_datalabel_pos(chart, 'ctr')")
    lines.append("")
    lines.append("")
    lines.append("def set_datalabel_pos_top(chart) -> None:")
    lines.append('    """Top (t) — 2,778 occurrences, common for markers/scatter."""')
    lines.append("    _set_datalabel_pos(chart, 't')")
    lines.append("")
    lines.append("")
    lines.append("def set_datalabel_pos_outside_end(chart) -> None:")
    lines.append('    """Outside end (outEnd) — 2,618 occurrences, common for clustered bars')
    lines.append("    where labels sit outside bar ends.")
    lines.append('    """')
    lines.append("    _set_datalabel_pos(chart, 'outEnd')")
    lines.append("")
    lines.append("")
    lines.append("def set_datalabel_pos_inside_end(chart) -> None:")
    lines.append('    """Inside end (inEnd) — labels inside bar with white text, prevents overflow.')
    lines.append("    143 occurrences, used on single_bar_with_delta patterns.")
    lines.append('    """')
    lines.append("    _set_datalabel_pos(chart, 'inEnd')")
    lines.append("")
    lines.append("")
    lines.append("def _set_datalabel_pos(chart, pos: str) -> None:")
    lines.append("    for ser in chart._chartSpace.findall(f\".//{qn_c('ser')}\"):")
    lines.append("        dlbls = ser.find(qn_c('dLbls'))")
    lines.append("        if dlbls is None:")
    lines.append("            dlbls = etree.SubElement(ser, qn_c('dLbls'))")
    lines.append("        dlp = dlbls.find(qn_c('dLblPos'))")
    lines.append("        if dlp is None:")
    lines.append("            dlp = etree.SubElement(dlbls, qn_c('dLblPos'))")
    lines.append("        dlp.set('val', pos)")
    lines.append("")
    lines.append("")
    lines.append("# ============================================================================")
    lines.append("# NUMBER FORMATS")
    lines.append("# ============================================================================")
    lines.append("")
    lines.append("def set_val_axis_pct_format(chart) -> None:")
    lines.append('    """Apply \\"0%\\" to val axis — 96% of all chart number formats in PET decks."""')
    lines.append("    _set_val_axis_num_format(chart, '0%')")
    lines.append("")
    lines.append("")
    lines.append("def set_val_axis_int_format(chart) -> None:")
    lines.append('    """Apply \\"0\\" (integer) to val axis — 13% of charts."""')
    lines.append("    _set_val_axis_num_format(chart, '0')")
    lines.append("")
    lines.append("")
    lines.append("def _set_val_axis_num_format(chart, fmt: str) -> None:")
    lines.append("    for ax in chart._chartSpace.findall(f\".//{qn_c('valAx')}\"):")
    lines.append("        nf = ax.find(qn_c('numFmt'))")
    lines.append("        if nf is None:")
    lines.append("            nf = etree.SubElement(ax, qn_c('numFmt'))")
    lines.append("        nf.set('formatCode', fmt)")
    lines.append("        nf.set('sourceLinked', '0')")
    lines.append("")
    lines.append("")
    lines.append("def set_datalabel_format(chart, fmt: str = '0%') -> None:")
    lines.append('    """Apply a number format to all series data labels.')
    lines.append("")
    lines.append("    Common formats:")
    lines.append("      '0%'       — integer percent (96% of labels)")
    lines.append("      '0'        — integer")
    lines.append("      '0.0'      — one decimal")
    lines.append("      '0%;\\\\-0%;\\\\ '  — percent with conditional sign (delta columns)")
    lines.append('    """')
    lines.append("    for ser in chart._chartSpace.findall(f\".//{qn_c('ser')}\"):")
    lines.append("        dlbls = ser.find(qn_c('dLbls'))")
    lines.append("        if dlbls is None:")
    lines.append("            dlbls = etree.SubElement(ser, qn_c('dLbls'))")
    lines.append("        nf = dlbls.find(qn_c('numFmt'))")
    lines.append("        if nf is None:")
    lines.append("            nf = etree.SubElement(dlbls, qn_c('numFmt'))")
    lines.append("        nf.set('formatCode', fmt)")
    lines.append("        nf.set('sourceLinked', '0')")
    lines.append("")
    lines.append("")
    lines.append("# ============================================================================")
    lines.append("# SERIES FORMATTING")
    lines.append("# ============================================================================")
    lines.append("")
    lines.append("def set_series_no_border(series) -> None:")
    lines.append('    """Remove series border via <a:ln><a:noFill/></a:ln>. Universal pattern')
    lines.append("    (74,367 noFill tags across all decks).")
    lines.append('    """')
    lines.append("    sp_pr = series._element.find(qn_c('spPr'))")
    lines.append("    if sp_pr is None:")
    lines.append("        sp_pr = etree.SubElement(series._element, qn_c('spPr'))")
    lines.append("    ln = sp_pr.find(qn_a('ln'))")
    lines.append("    if ln is None:")
    lines.append("        ln = etree.SubElement(sp_pr, qn_a('ln'))")
    lines.append("    # Remove existing fills and set noFill")
    lines.append("    for child in list(ln):")
    lines.append("        ln.remove(child)")
    lines.append("    etree.SubElement(ln, qn_a('noFill'))")
    lines.append("")
    lines.append("")
    lines.append("def set_series_fill_rgb(series, hex_color: str) -> None:")
    lines.append('    """Set series fill to solid sRGB color. Prefer this over scheme colors.')
    lines.append("")
    lines.append("    hex_color: 'RRGGBB' or '#RRGGBB' — case-insensitive.")
    lines.append('    """')
    lines.append("    hex_clean = hex_color.lstrip('#').upper()")
    lines.append("    sp_pr = series._element.find(qn_c('spPr'))")
    lines.append("    if sp_pr is None:")
    lines.append("        sp_pr = etree.SubElement(series._element, qn_c('spPr'))")
    lines.append("    # Remove existing fill children")
    lines.append("    for tag in ('solidFill', 'gradFill', 'blipFill', 'pattFill', 'noFill'):")
    lines.append("        existing = sp_pr.find(qn_a(tag))")
    lines.append("        if existing is not None:")
    lines.append("            sp_pr.remove(existing)")
    lines.append("    solid = etree.SubElement(sp_pr, qn_a('solidFill'))")
    lines.append("    rgb = etree.SubElement(solid, qn_a('srgbClr'))")
    lines.append("    rgb.set('val', hex_clean)")
    lines.append("")
    lines.append("")
    lines.append("def set_series_line_width(series, emu: int = 25400) -> None:")
    lines.append('    """Set series line width in EMU. Common values:')
    lines.append("      12700 = 1pt")
    lines.append("      19050 = 1.5pt")
    lines.append("      25400 = 2pt (PET default, 827 occurrences)")
    lines.append("      28575 = 2.25pt (most common overall, 1,223 occurrences)")
    lines.append('    """')
    lines.append("    sp_pr = series._element.find(qn_c('spPr'))")
    lines.append("    if sp_pr is None:")
    lines.append("        sp_pr = etree.SubElement(series._element, qn_c('spPr'))")
    lines.append("    ln = sp_pr.find(qn_a('ln'))")
    lines.append("    if ln is None:")
    lines.append("        ln = etree.SubElement(sp_pr, qn_a('ln'))")
    lines.append("    ln.set('w', str(emu))")
    lines.append("")
    lines.append("")
    lines.append("def set_series_marker_circle(series, size: int = 7) -> None:")
    lines.append('    """Set marker to circle (94% of markers in scatter/line charts).')
    lines.append("")
    lines.append("    size: marker size in points (observed range 3-12, default 7).")
    lines.append('    """')
    lines.append("    _set_series_marker(series, 'circle', size)")
    lines.append("")
    lines.append("")
    lines.append("def _set_series_marker(series, symbol: str, size: int) -> None:")
    lines.append("    marker = series._element.find(qn_c('marker'))")
    lines.append("    if marker is None:")
    lines.append("        marker = etree.SubElement(series._element, qn_c('marker'))")
    lines.append("    sym = marker.find(qn_c('symbol'))")
    lines.append("    if sym is None:")
    lines.append("        sym = etree.SubElement(marker, qn_c('symbol'))")
    lines.append("    sym.set('val', symbol)")
    lines.append("    sz = marker.find(qn_c('size'))")
    lines.append("    if sz is None:")
    lines.append("        sz = etree.SubElement(marker, qn_c('size'))")
    lines.append("    sz.set('val', str(size))")
    lines.append("")
    lines.append("")
    lines.append("# ============================================================================")
    lines.append("# GRIDLINES")
    lines.append("# ============================================================================")
    lines.append("")
    lines.append("def remove_major_gridlines(chart) -> None:")
    lines.append('    """Remove major gridlines from all value axes (84% of PET charts have none)."""')
    lines.append("    for ax in chart._chartSpace.findall(f\".//{qn_c('valAx')}\"):")
    lines.append("        mg = ax.find(qn_c('majorGridlines'))")
    lines.append("        if mg is not None:")
    lines.append("            ax.remove(mg)")
    lines.append("")
    lines.append("")
    lines.append("def add_major_gridlines(chart) -> None:")
    lines.append('    """Add major gridlines (used in 16% of PET charts)."""')
    lines.append("    for ax in chart._chartSpace.findall(f\".//{qn_c('valAx')}\"):")
    lines.append("        if ax.find(qn_c('majorGridlines')) is None:")
    lines.append("            etree.SubElement(ax, qn_c('majorGridlines'))")
    lines.append("")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# CHART_PATTERNS{} generator
# ---------------------------------------------------------------------------


def generate_chart_patterns_py() -> str:
    """Generate CHART_PATTERNS{} with defaults for each chart type, derived from deck observations."""
    lines: list[str] = []
    lines.append('"""')
    lines.append("Named chart patterns — canonical defaults for each chart type.")
    lines.append("")
    lines.append("Generated from real-deck analysis. Each pattern encodes the defaults observed")
    lines.append("in client decks for that chart type. Renderers use these as starting points;")
    lines.append("renderer args can override any field.")
    lines.append('"""')
    lines.append("from __future__ import annotations")
    lines.append("")
    lines.append("from pptx.enum.chart import XL_CHART_TYPE")
    lines.append("")
    lines.append("")
    lines.append("CHART_PATTERNS = {")
    lines.append("")
    lines.append("    # ---- bar_clustered (horizontal) ----")
    lines.append("    # 35% of all PET charts. Horizontal bars with inverted cat axis,")
    lines.append("    # cat labels hidden (shown in companion table), data labels inside-end.")
    lines.append('    "bar_clustered_horizontal": {')
    lines.append('        "chart_type": XL_CHART_TYPE.BAR_CLUSTERED,')
    lines.append('        "bar_dir": "bar",  # horizontal')
    lines.append('        "grouping": "clustered",')
    lines.append('        "gap_width": 80,')
    lines.append('        "overlap": 0,')
    lines.append('        "invert_cat_axis": True,')
    lines.append('        "hide_cat_labels": True,')
    lines.append('        "datalabel_pos": "inEnd",')
    lines.append('        "val_num_format": "0%",')
    lines.append('        "show_legend": False,')
    lines.append('        "show_title": False,')
    lines.append('        "major_gridlines": False,')
    lines.append('        "series_no_border": True,')
    lines.append('        "invert_if_negative": False,')
    lines.append('    },')
    lines.append("")
    lines.append("    # ---- bar_stacked_100 (horizontal) ----")
    lines.append("    # Horizontal 100% stacked, typical for intent distributions and composition.")
    lines.append('    "bar_stacked_100_horizontal": {')
    lines.append('        "chart_type": XL_CHART_TYPE.BAR_STACKED_100,')
    lines.append('        "bar_dir": "bar",')
    lines.append('        "grouping": "percentStacked",')
    lines.append('        "gap_width": 80,')
    lines.append('        "overlap": 100,  # fully stacked')
    lines.append('        "invert_cat_axis": True,')
    lines.append('        "hide_cat_labels": True,')
    lines.append('        "datalabel_pos": "ctr",')
    lines.append('        "val_num_format": "0%",')
    lines.append('        "show_legend": False,')
    lines.append('        "show_title": False,')
    lines.append('        "major_gridlines": False,')
    lines.append('        "series_no_border": True,')
    lines.append('        "invert_if_negative": False,')
    lines.append('    },')
    lines.append("")
    lines.append("    # ---- column_stacked_100 (vertical) ----")
    lines.append("    # 11% of PET charts. Vertical 100% stacked for proportions.")
    lines.append('    "column_stacked_100_vertical": {')
    lines.append('        "chart_type": XL_CHART_TYPE.COLUMN_STACKED_100,')
    lines.append('        "bar_dir": "col",')
    lines.append('        "grouping": "percentStacked",')
    lines.append('        "gap_width": 100,')
    lines.append('        "overlap": 100,')
    lines.append('        "invert_cat_axis": False,')
    lines.append('        "hide_cat_labels": False,')
    lines.append('        "datalabel_pos": "ctr",')
    lines.append('        "val_num_format": "0%",')
    lines.append('        "show_legend": False,')
    lines.append('        "show_title": False,')
    lines.append('        "major_gridlines": False,')
    lines.append('        "series_no_border": True,')
    lines.append('        "invert_if_negative": False,')
    lines.append('    },')
    lines.append("")
    lines.append("    # ---- column_clustered (vertical, QoQ) ----")
    lines.append("    # Side-by-side bars for quarter-over-quarter comparison.")
    lines.append('    "column_clustered_vertical": {')
    lines.append('        "chart_type": XL_CHART_TYPE.COLUMN_CLUSTERED,')
    lines.append('        "bar_dir": "col",')
    lines.append('        "grouping": "clustered",')
    lines.append('        "gap_width": 100,')
    lines.append('        "overlap": -20,  # slight gap between series',)
    lines.append('        "invert_cat_axis": False,')
    lines.append('        "hide_cat_labels": False,')
    lines.append('        "datalabel_pos": "outEnd",')
    lines.append('        "val_num_format": "0%",')
    lines.append('        "show_legend": False,')
    lines.append('        "show_title": False,')
    lines.append('        "major_gridlines": False,')
    lines.append('        "series_no_border": True,')
    lines.append('        "invert_if_negative": False,')
    lines.append('    },')
    lines.append("")
    lines.append("    # ---- line_markers (trended) ----")
    lines.append("    # 14% of PET charts. Multi-wave trend lines with circle markers.")
    lines.append('    "line_markers_trended": {')
    lines.append('        "chart_type": XL_CHART_TYPE.LINE_MARKERS,')
    lines.append('        "grouping": "standard",')
    lines.append('        "datalabel_pos": "t",  # top of marker')
    lines.append('        "val_num_format": "0%",')
    lines.append('        "show_legend": False,')
    lines.append('        "show_title": False,')
    lines.append('        "major_gridlines": False,')
    lines.append('        "marker_symbol": "circle",')
    lines.append('        "marker_size": 7,')
    lines.append('        "line_width_emu": 25400,  # 2pt')
    lines.append('    },')
    lines.append("")
    lines.append("    # ---- xy_scatter (abacus) ----")
    lines.append("    # 17% of PET charts. Scatter for multi-attribute comparisons (abacus, MBD).")
    lines.append('    "xy_scatter_abacus": {')
    lines.append('        "chart_type": XL_CHART_TYPE.XY_SCATTER,')
    lines.append('        "datalabel_pos": "t",')
    lines.append('        "val_num_format": "0%",')
    lines.append('        "show_legend": False,')
    lines.append('        "show_title": False,')
    lines.append('        "major_gridlines": False,')
    lines.append('        "marker_symbol": "circle",')
    lines.append('        "marker_size": 10,')
    lines.append('        "hide_val_labels": False,')
    lines.append('        "hide_cat_labels": False,')
    lines.append('    },')
    lines.append("")
    lines.append("    # ---- doughnut ----")
    lines.append("    # 138 occurrences. Used for segment composition breakdowns.")
    lines.append('    "doughnut_default": {')
    lines.append('        "chart_type": XL_CHART_TYPE.DOUGHNUT,')
    lines.append('        "datalabel_pos": "ctr",')
    lines.append('        "val_num_format": "0%",')
    lines.append('        "show_legend": False,')
    lines.append('        "show_title": False,')
    lines.append('        "hole_size": 50,  # percent')
    lines.append('        "series_no_border": True,')
    lines.append('    },')
    lines.append("")
    lines.append("}")
    lines.append("")
    lines.append("")
    lines.append("def get_pattern(pattern_name: str) -> dict:")
    lines.append('    """Look up chart pattern by name. Raises KeyError with available patterns."""')
    lines.append("    if pattern_name not in CHART_PATTERNS:")
    lines.append('        raise KeyError(')
    lines.append('            f"Unknown pattern {pattern_name!r}. Available: {sorted(CHART_PATTERNS.keys())}"')
    lines.append("        )")
    lines.append("    return CHART_PATTERNS[pattern_name].copy()")
    lines.append("")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> int:
    print("Generating pptx_utils Python modules from deck analysis...")

    brand_src = generate_brand_py()
    (OUTPUTS_DIR / "generated_brand.py").write_text(brand_src, encoding="utf-8")
    print(f"  generated_brand.py           ({len(brand_src.splitlines())} lines)")

    layouts_src = generate_layouts_py()
    (OUTPUTS_DIR / "generated_layouts.py").write_text(layouts_src, encoding="utf-8")
    print(f"  generated_layouts.py         ({len(layouts_src.splitlines())} lines)")

    lxml_src = generate_lxml_helpers_py()
    (OUTPUTS_DIR / "generated_lxml_helpers.py").write_text(lxml_src, encoding="utf-8")
    print(f"  generated_lxml_helpers.py    ({len(lxml_src.splitlines())} lines)")

    patterns_src = generate_chart_patterns_py()
    (OUTPUTS_DIR / "generated_chart_patterns.py").write_text(patterns_src, encoding="utf-8")
    print(f"  generated_chart_patterns.py  ({len(patterns_src.splitlines())} lines)")

    print(f"\nAll outputs in {OUTPUTS_DIR}/")
    print("\nReview, then copy approved content into slidegen/pptx_utils/")
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main())
