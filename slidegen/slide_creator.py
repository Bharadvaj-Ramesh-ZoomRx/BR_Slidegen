"""
slidegen/slide_creator.py — THE atomic unit.

One validated SlideSpec in, one client-ready slide out. Zero intelligence.
Composes pptx_utils primitives deterministically; never writes raw lxml.

Supports all 10 CHART_PATTERNS (88%+ of real PET charts covered by top 6) and
all 7 component types in the spec schema. Full contract in:
    .claude/skills/creation/slide-creator/SKILL.md

Primary API:
    render_slide(spec, prs=None, template_path=None)  → (Presentation, Slide)
    render_deck(specs, output_path, template_path=None)  → Presentation
    render_spec_to_file(spec, output_path, template_path=None)  → Path

Color resolution — 4 sources supported (PRD §6.7):
    1. Explicit hex:       "#F75824"
    2. BRAND{} token:      "{brand.primary_current}"
    3. Context file token: "{context.brand_palette.primary}" (requires context kwarg)
    4. Deck-reader token:  "{deck.slide_N.series_M.color}"  (requires deck_ref kwarg)
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Optional

from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from slidegen.slide_spec import (
    SlideSpec,
    ComponentSpec,
    ChartComponent,
    LabelTableComponent,
    ValueTableComponent,
    DeltaColumnComponent,
    CalloutComponent,
    ImageComponent,
    TextboxComponent,
    Position,
    validate_spec,
)
# Import from submodules directly — pptx_utils/__init__.py doesn't re-export
# everything added in the Apr 15 deck-analysis extension. Direct imports are
# more robust to that and easier to review.
from slidegen.pptx_utils.brand import BRAND, get_brand
from slidegen.pptx_utils.layout import (
    LAYOUTS, slide_header, slide_footer, section_header_bar,
)
from slidegen.pptx_utils.charts import (
    CHART_PATTERNS,
    add_single_bar_chart, add_clustered_bar_chart, add_line_chart,
    add_stacked_column_chart, add_scatter_chart,
)
from slidegen.pptx_utils.tables import add_delta_col, add_value_table
from slidegen.pptx_utils.shapes import textbox, solidrect, callout_box
from slidegen.pptx_utils.images import insert_image
from slidegen.pptx_utils.lxml_helpers import (
    invert_cat_axis, hide_cat_labels, hide_axis,
    set_plot_area_gap, set_overlap,
    set_series_color, set_series_no_border, set_data_label_color,
    set_datalabel_pos_outside_end,
    set_gridlines, set_series_marker, set_series_line_style,
    set_val_axis_scale,
    # Apr 15 extensions — may or may not be re-exported via __init__
    set_invert_if_negative, hide_val_labels,
    set_datalabel_pos_center, set_datalabel_pos_top, set_datalabel_pos_inside_end,
    set_val_axis_pct_format, set_val_axis_int_format,
    set_datalabel_format,
    set_pie_slice_colors, set_donut_hole_size,
)

# Slide dimensions — matches BRAND/layout constants
_SLIDE_W_IN = 13.333
_SLIDE_H_IN = 7.500

# Blank slide layout in the default python-pptx Presentation is index 6
_BLANK_LAYOUT_IDX = 6


# ─────────────────────────────────────────────────────────────────────────────
# Shape naming
# ─────────────────────────────────────────────────────────────────────────────


class ShapeNamer:
    """Assigns zrx_{slide:03d}_{shape:03d} names to shapes in deterministic order."""

    def __init__(self, slide_index: int):
        self.slide_index = slide_index
        self._counter = 0

    def name(self, shape) -> str:
        self._counter += 1
        label = f"zrx_{self.slide_index:03d}_{self._counter:03d}"
        try:
            shape.name = label
        except Exception:
            pass  # some shapes (charts) don't allow rename; registry still tracks
        return label


# ─────────────────────────────────────────────────────────────────────────────
# Color resolution
# ─────────────────────────────────────────────────────────────────────────────


_HEX_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")
_BRAND_TOKEN_RE = re.compile(r"^\{brand\.([a-z_]+)\}$")
_CONTEXT_TOKEN_RE = re.compile(r"^\{context\.([a-z_0-9.]+)\}$")
_DECK_TOKEN_RE = re.compile(r"^\{deck\.([a-z_0-9.]+)\}$")


class ColorResolutionError(ValueError):
    """Raised when a color token cannot be resolved given the available context."""


def resolve_color(
    token: str,
    brand: Optional[dict] = None,
    context: Optional[dict] = None,
    deck_ref: Optional[dict] = None,
) -> RGBColor:
    """Resolve a color token to a python-pptx RGBColor.

    Token grammar (see PRD §6.7):
      - "#RRGGBB"                         — explicit hex
      - "{brand.role}"                    — looks up BRAND[...][role]
      - "{context.path.to.color}"         — looks up context dict by dotted path
      - "{deck.slide_N.series_M.color}"   — looks up deck_ref by dotted path

    Raises ColorResolutionError if a token references missing data.
    """
    if not isinstance(token, str) or not token:
        raise ColorResolutionError(f"Color token must be a non-empty string, got {token!r}")

    # 1. Explicit hex
    if _HEX_RE.match(token):
        h = token[1:]
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

    # 2. BRAND{} token
    m = _BRAND_TOKEN_RE.match(token)
    if m:
        role = m.group(1)
        if brand is None:
            raise ColorResolutionError(
                f"Token {token!r} requires spec.brand to be set"
            )
        if role not in brand:
            raise ColorResolutionError(
                f"BRAND[{brand.get('brand', '?')!r}] has no role {role!r}. "
                f"Available roles: {sorted(k for k in brand.keys() if not k.startswith('_'))}"
            )
        val = brand[role]
        if isinstance(val, RGBColor):
            return val
        if isinstance(val, str) and _HEX_RE.match(val if val.startswith("#") else f"#{val}"):
            h = val.lstrip("#")
            return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
        raise ColorResolutionError(
            f"BRAND[{role!r}] value is not a color: {val!r}"
        )

    # 3. Context file token
    m = _CONTEXT_TOKEN_RE.match(token)
    if m:
        if context is None:
            raise ColorResolutionError(
                f"Token {token!r} requires context kwarg"
            )
        return _lookup_dotted(context, m.group(1).split("."), token)

    # 4. Deck-reader token
    m = _DECK_TOKEN_RE.match(token)
    if m:
        if deck_ref is None:
            raise ColorResolutionError(
                f"Token {token!r} requires deck_ref kwarg"
            )
        return _lookup_dotted(deck_ref, m.group(1).split("."), token)

    raise ColorResolutionError(f"Unrecognized color token: {token!r}")


def _lookup_dotted(root: dict, path: list[str], original_token: str) -> RGBColor:
    """Walk a dict by dotted path, coerce leaf to RGBColor."""
    cur: Any = root
    for key in path:
        if isinstance(cur, dict) and key in cur:
            cur = cur[key]
        else:
            raise ColorResolutionError(
                f"Token {original_token!r} — path {'.'.join(path)!r} not found"
            )
    if isinstance(cur, RGBColor):
        return cur
    if isinstance(cur, str):
        h = cur.lstrip("#")
        if len(h) == 6 and all(c in "0123456789abcdefABCDEF" for c in h):
            return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    raise ColorResolutionError(
        f"Token {original_token!r} resolved to non-color value: {cur!r}"
    )


# ─────────────────────────────────────────────────────────────────────────────
# Position resolution
# ─────────────────────────────────────────────────────────────────────────────


def resolve_position(
    pos: Position,
    layout_cfg: dict,
    spec_layout_name: str,
) -> tuple[float, float, float, float]:
    """Resolve a Position to (left, top, width, height) in inches.

    - Explicit bbox: returned as-is.
    - Preset: looks up `layout_cfg[pos.preset]` or `LAYOUTS[pos.layout][pos.preset]`.
    """
    if pos.is_explicit():
        return (pos.left, pos.top, pos.width, pos.height)

    if pos.is_preset():
        target_layout = pos.layout or spec_layout_name
        if target_layout in LAYOUTS:
            cfg = LAYOUTS[target_layout]
        else:
            raise KeyError(f"Position references layout {target_layout!r} not in LAYOUTS")

        if pos.preset not in cfg:
            available = sorted(k for k in cfg.keys() if "rect" in k or k.startswith("chart") or k.startswith("table") or k.startswith("delta"))
            raise KeyError(
                f"Position preset {pos.preset!r} not in LAYOUTS[{target_layout!r}]. "
                f"Available: {available}"
            )

        rect = cfg[pos.preset]
        if isinstance(rect, dict) and all(k in rect for k in ("left", "top", "width", "height")):
            return (rect["left"], rect["top"], rect["width"], rect["height"])
        raise KeyError(
            f"LAYOUTS[{target_layout!r}][{pos.preset!r}] is not a rect dict: {rect!r}"
        )

    raise ValueError(f"Position has neither explicit bbox nor preset: {pos}")


# ─────────────────────────────────────────────────────────────────────────────
# Chart chrome application (shared across all patterns)
# ─────────────────────────────────────────────────────────────────────────────


# Chart patterns where the value axis MUST stay visible regardless of the
# chrome.value_axis.show default. Line/scatter charts collapse their plot
# area if the value axis is deleted (the line positions depend on axis extent).
# Real-deck analysis confirms: trended scorecards always show the value axis.
_PATTERNS_REQUIRING_VALUE_AXIS: frozenset[str] = frozenset({
    "line_markers_trended",
    "xy_scatter_abacus",
    # Column/clustered charts also benefit from visible axes but don't
    # strictly require them — leaving spec-controlled for now.
})


def _hide_legend_but_keep_element(chart) -> None:
    """Hide the legend visually while keeping the <c:legend> element in XML.

    Why: python-pptx's `chart.has_legend = False` deletes the entire
    <c:legend> element. PowerPoint's strict loader rejects that removal
    specifically for line and doughnut charts, triggering the Repair prompt
    and rendering an empty chart area. Keeping the element present — but
    with a <c:legendPos> + zero-size <c:layout> — satisfies PowerPoint's
    schema expectations AND keeps the legend visually hidden.
    """
    try:
        chart.legend.include_in_layout = False
    except Exception:
        pass

    from lxml import etree

    C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    def c_qn(tag: str) -> str:
        return f"{{{C_NS}}}{tag}"

    legend_el = chart.legend._element

    # Reset: remove any existing legendPos/layout we'll rebuild
    for tag_name in ("legendPos", "layout"):
        for existing in legend_el.findall(c_qn(tag_name)):
            legend_el.remove(existing)

    # Per CT_Legend schema: legendPos?, legendEntry*, layout?, overlay?, spPr?, txPr?
    # Raw python-pptx produces <c:legendPos val="b"/> as the first child.
    # Build legendPos + zero-size layout, prepending so order is schema-correct.
    layout_el = etree.Element(c_qn("layout"))
    manual = etree.SubElement(layout_el, c_qn("manualLayout"))
    for tag, val in (("xMode", "edge"), ("yMode", "edge")):
        el = etree.SubElement(manual, c_qn(tag))
        el.set("val", val)
    for tag in ("x", "y", "w", "h"):
        el = etree.SubElement(manual, c_qn(tag))
        el.set("val", "0")

    legend_pos_el = etree.Element(c_qn("legendPos"))
    legend_pos_el.set("val", "b")

    # Insert legendPos first, then layout after any legendEntry elements
    legend_el.insert(0, legend_pos_el)
    # Find position after any legendEntry to insert layout
    insert_idx = 1  # after legendPos
    for i, child in enumerate(legend_el):
        if _child_local_name(child) == "legendEntry":
            insert_idx = i + 1
    legend_el.insert(insert_idx, layout_el)


def _remove_auto_dLbls_on_chart_type(chart) -> None:
    """Remove any chart-type-level <c:dLbls> that python-pptx auto-creates.

    When we access `series.data_labels` in _apply_data_labels, python-pptx
    can side-effect a chart-type-level (e.g. <c:lineChart>/<c:dLbls>) container
    that raw python-pptx does NOT produce. PowerPoint's strict loader rejects
    this extra dLbls for line charts specifically (causes Repair prompt).

    Our spec's data-label settings live on the SERIES-level <c:dLbls>, so
    removing the chart-type-level container is safe — no visual effect,
    no information loss.

    Keep the chart-type-level dLbls for doughnut/bar (raw python-pptx
    produces it for those chart types — removing would be schema-incorrect).
    """
    if not chart.plots:
        return
    # Only strip for lineChart — raw produces 0 for line, 1 for bar/doughnut
    for plot in chart.plots:
        el = plot._element
        if _child_local_name(el) != "lineChart":
            continue
        for dlbls_el in el.findall("{%s}dLbls" % _C_NS):
            el.remove(dlbls_el)


def _apply_chart_chrome(
    chart,
    component: ChartComponent,
    brand: Optional[dict],
    context: Optional[dict],
    deck_ref: Optional[dict],
) -> None:
    """Apply title / legend / gridlines / data-label format per the spec's chrome.

    Defaults reflect Apr 15 deck-analysis frequencies (99% no title, 99% no
    legend, 84% no gridlines, 96% "0%" labels). Spec overrides those defaults
    when explicitly set.
    """
    chrome = component.chrome

    # Title
    if chrome.title is None:
        chart.has_title = False
    else:
        chart.has_title = True
        try:
            chart.chart_title.text_frame.text = chrome.title
        except Exception:
            pass

    # Legend — NEVER remove the <c:legend> element. python-pptx's
    # `has_legend = False` deletes it entirely, which PowerPoint's strict
    # loader rejects for line and doughnut charts (triggers the Repair
    # prompt and produces an empty-looking slide). Instead, keep the element
    # present and collapse it to zero size when the spec says no legend.
    # Matches the structure python-pptx natively produces + what real PET
    # decks contain (legend element present, visually hidden via layout).
    chart.has_legend = True  # ensures <c:legend> element exists in XML
    pos_map = {
        "top":    XL_LEGEND_POSITION.TOP,
        "bottom": XL_LEGEND_POSITION.BOTTOM,
        "left":   XL_LEGEND_POSITION.LEFT,
        "right":  XL_LEGEND_POSITION.RIGHT,
    }
    try:
        if chrome.legend.show:
            chart.legend.position = pos_map.get(chrome.legend.position, XL_LEGEND_POSITION.BOTTOM)
            chart.legend.include_in_layout = False
        else:
            # Hide legend visually by collapsing it via layout manipulation
            _hide_legend_but_keep_element(chart)
    except Exception:
        pass

    # Gridlines
    try:
        if not chrome.gridlines:
            # Turn off major gridlines on both axes when applicable
            for ax_name in ("value_axis", "category_axis"):
                ax = getattr(chart, ax_name, None)
                if ax is not None:
                    try:
                        ax.has_major_gridlines = False
                    except Exception:
                        pass
    except Exception:
        pass

    # Value axis — hide unless spec says show OR pattern requires visible axis.
    # Line/scatter charts' plot-area layout depends on the value axis being
    # present; deleting it (via <c:delete val="1"/>) causes PowerPoint to flag
    # the file for Repair and show an empty chart area.
    if not chrome.value_axis.show and component.chart_pattern not in _PATTERNS_REQUIRING_VALUE_AXIS:
        try:
            hide_axis(chart, "val")
        except Exception:
            pass


# OOXML dLblPos is constrained per chart type. Invalid values trigger
# PowerPoint's Repair prompt. Per the ECMA-376 schema:
#   Bar/Column:  "ctr" | "inBase" | "inEnd" | "outEnd"
#   Line/Scatter/Stock: "t" (above) | "b" (below) | "l" | "r" | "ctr"
#   Pie/Doughnut: "ctr" | "bestFit" | "inEnd" | "outEnd"
# python-pptx permits cross-type values (library is lenient); PowerPoint
# rejects them (strict loader). Clamp to the valid subset per chart pattern.
_LINE_SCATTER_PATTERNS = {"line_markers_trended", "xy_scatter_abacus"}
_PIE_DOUGHNUT_PATTERNS = {"doughnut_default"}


def _clamp_label_position(position: str, chart_pattern: str) -> str | None:
    """Return a valid dLblPos value for the given chart pattern, or None to skip.

    Input `position` uses our spec vocabulary ("above", "inEnd", etc.) which
    maps 1:1 onto OOXML values in most cases. Positions that don't exist for
    a given chart type are remapped to a sensible equivalent:
      - inEnd / outEnd on line/scatter → "above"
      - above / below on bar/column    → "outEnd"
      - ANY position on pie/doughnut   → None (skip — PowerPoint rejects
        <c:dLblPos> at the series-level dLbls for doughnut; its default
        "bestFit" positioning works correctly without it)
    """
    if chart_pattern in _LINE_SCATTER_PATTERNS:
        if position in ("inEnd", "outEnd"):
            return "above"   # maps to OOXML "t"
        if position in ("above", "below", "ctr"):
            return position
        return "above"       # safe default for line/scatter
    if chart_pattern in _PIE_DOUGHNUT_PATTERNS:
        # PowerPoint rejects <c:dLblPos> inside <c:ser>/<c:dLbls> for
        # doughnut/pie charts — triggers Repair prompt. Return None to
        # skip setting any position; PowerPoint defaults to bestFit.
        return None
    # Bar / column / other: spec vocabulary's inEnd/outEnd/ctr are all valid
    if position in ("above", "below"):
        return "outEnd"      # bar charts use outEnd instead of above
    return position


def _apply_data_labels(
    chart,
    component: ChartComponent,
    brand: Optional[dict],
    context: Optional[dict],
    deck_ref: Optional[dict],
) -> None:
    """Apply data-label settings (show/format/position/color) to all series.

    Color handling nuance: the spec's `font_color` is only applied when labels
    sit ON a colored background (inside-end position on bars, center position
    on stacked segments) — that's where white-on-color makes sense. For labels
    positioned OUTSIDE the colored shape (above / below / outEnd), the series
    color is the right choice and is preserved from `enable_data_labels` (run
    earlier by the per-pattern chart builder).

    Position clamping: `dLblPos` values differ between bar/line/pie chart
    types per OOXML schema. An invalid-for-type position (e.g. "inEnd" on a
    line chart) triggers PowerPoint's Repair prompt. `_clamp_label_position`
    maps the spec's position to a valid one per chart pattern.
    """
    dl = component.chrome.data_labels
    if not dl.show:
        return

    # Clamp position to schema-valid value for this chart type.
    # Returns None for pie/doughnut (PowerPoint rejects dLblPos on series dLbls).
    effective_position = _clamp_label_position(dl.position, component.chart_pattern)

    # Color override: only apply spec's font_color when labels sit on a
    # colored background (inEnd on bars, ctr on stacked segments).
    _ON_COLOR_POSITIONS = {"inEnd", "ctr"}
    apply_color_override = effective_position in _ON_COLOR_POSITIONS

    label_color: Optional[RGBColor] = None
    if dl.font_color and apply_color_override:
        try:
            label_color = resolve_color(dl.font_color, brand=brand, context=context, deck_ref=deck_ref)
        except ColorResolutionError:
            label_color = None

    pos_fn_map = {
        "inEnd":  set_datalabel_pos_inside_end,
        "outEnd": set_datalabel_pos_outside_end,
        "ctr":    set_datalabel_pos_center,
        "above":  set_datalabel_pos_top,
    }
    # None → skip position setting (pie/doughnut uses PowerPoint default bestFit)
    pos_fn = pos_fn_map.get(effective_position) if effective_position else None

    for series in chart.series:
        # Enable + format + (optionally) position
        try:
            series.data_labels.show_value = True
            set_datalabel_format(series, dl.format or "0%")
            if pos_fn is not None:
                pos_fn(series)
            if label_color is not None:
                set_data_label_color(series, label_color)
        except Exception:
            # Some chart types (scatter) expose labels differently; the
            # per-pattern renderer handles those separately
            pass


# ─────────────────────────────────────────────────────────────────────────────
# Per-pattern chart renderers
# ─────────────────────────────────────────────────────────────────────────────


def _series_colors(
    component: ChartComponent,
    brand: Optional[dict],
    context: Optional[dict],
    deck_ref: Optional[dict],
) -> list[RGBColor]:
    """Resolve every series' color token to RGBColor, in series order."""
    return [
        resolve_color(s.color, brand=brand, context=context, deck_ref=deck_ref)
        for s in component.data.series
    ]


def _render_bar_clustered_horizontal(
    slide, component: ChartComponent, rect: tuple, brand, context, deck_ref,
):
    cats = component.data.categories
    series_list = [(s.name, list(s.values)) for s in component.data.series]
    colors = _series_colors(component, brand, context, deck_ref)

    pattern = CHART_PATTERNS["bar_clustered_horizontal"]
    cf, chart = add_clustered_bar_chart(
        slide, cats, series_list,
        rect[0], rect[1], rect[2], rect[3],
        colors=colors,
        legend=component.chrome.legend.show,
        gap=component.chrome.gap_width or pattern["gap"],
        overlap=component.chrome.overlap if component.chrome.overlap is not None else pattern["overlap"],
    )
    invert_cat_axis(chart)
    if component.chrome.hide_category_labels:
        hide_cat_labels(chart)
    for s in chart.series:
        set_series_no_border(s)
        set_invert_if_negative(chart)
    return chart


def _render_column_clustered_vertical(
    slide, component: ChartComponent, rect: tuple, brand, context, deck_ref,
):
    cats = component.data.categories
    series_list = [(s.name, list(s.values)) for s in component.data.series]
    colors = _series_colors(component, brand, context, deck_ref)

    # add_clustered_bar_chart builds BAR_CLUSTERED; for vertical columns we
    # create the chart directly.
    chart_data = CategoryChartData()
    chart_data.categories = cats
    for name, vals in series_list:
        chart_data.add_series(name, vals)
    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(rect[0]), Inches(rect[1]), Inches(rect[2]), Inches(rect[3]),
        chart_data,
    )
    chart = cf.chart

    pattern = CHART_PATTERNS["column_clustered_vertical"]
    set_plot_area_gap(chart, component.chrome.gap_width or pattern["gap"])
    ov = component.chrome.overlap if component.chrome.overlap is not None else pattern.get("overlap")
    if ov:
        set_overlap(chart, ov)
    for i, s in enumerate(chart.series):
        if i < len(colors):
            set_series_color(s, colors[i])
        set_series_no_border(s)
    return chart


def _render_bar_stacked_100_horizontal(
    slide, component: ChartComponent, rect: tuple, brand, context, deck_ref,
):
    cats = component.data.categories
    series_list = [(s.name, list(s.values)) for s in component.data.series]
    colors = _series_colors(component, brand, context, deck_ref)

    chart_data = CategoryChartData()
    chart_data.categories = cats
    for name, vals in series_list:
        chart_data.add_series(name, vals)
    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_STACKED_100,
        Inches(rect[0]), Inches(rect[1]), Inches(rect[2]), Inches(rect[3]),
        chart_data,
    )
    chart = cf.chart

    pattern = CHART_PATTERNS["bar_stacked_100_horizontal"]
    set_plot_area_gap(chart, pattern["gap"])
    set_overlap(chart, pattern["overlap"])
    invert_cat_axis(chart)
    if component.chrome.hide_category_labels:
        hide_cat_labels(chart)
    for i, s in enumerate(chart.series):
        if i < len(colors):
            set_series_color(s, colors[i])
        set_series_no_border(s)
    return chart


def _render_column_stacked_100_vertical(
    slide, component: ChartComponent, rect: tuple, brand, context, deck_ref,
):
    cats = component.data.categories
    series_list = [(s.name, list(s.values)) for s in component.data.series]
    colors = _series_colors(component, brand, context, deck_ref)

    chart_data = CategoryChartData()
    chart_data.categories = cats
    for name, vals in series_list:
        chart_data.add_series(name, vals)
    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_STACKED_100,
        Inches(rect[0]), Inches(rect[1]), Inches(rect[2]), Inches(rect[3]),
        chart_data,
    )
    chart = cf.chart

    pattern = CHART_PATTERNS["column_stacked_100_vertical"]
    set_plot_area_gap(chart, pattern["gap"])
    set_overlap(chart, pattern["overlap"])
    for i, s in enumerate(chart.series):
        if i < len(colors):
            set_series_color(s, colors[i])
        set_series_no_border(s)
    return chart


def _render_line_markers_trended(
    slide, component: ChartComponent, rect: tuple, brand, context, deck_ref,
):
    """Real-deck trended scorecards show data labels above markers (so users
    read exact %), which makes the value-axis tick labels redundant. Default
    behavior: keep the value axis STRUCTURALLY PRESENT (PowerPoint needs it
    for layout; deleting it caused the Repair bug earlier) but HIDE its tick
    labels. Category axis labels (period names) stay visible — users need to
    know which period each point represents.
    """
    cats = component.data.categories
    series_list = [(s.name, list(s.values)) for s in component.data.series]
    colors = _series_colors(component, brand, context, deck_ref)
    pattern = CHART_PATTERNS["line_markers_trended"]

    # Infer scale bounds from the data if the spec doesn't specify
    all_vals = [v for s in component.data.series for v in s.values if v is not None]
    if component.chrome.value_axis.min is not None:
        scale_min = component.chrome.value_axis.min
    else:
        scale_min = 0.0
    if component.chrome.value_axis.max is not None:
        scale_max = component.chrome.value_axis.max
    else:
        scale_max = max(0.1, (int(max(all_vals) * 10) + 2) / 10) if all_vals else 1.0

    cf, chart = add_line_chart(
        slide, cats, series_list,
        rect[0], rect[1], rect[2], rect[3],
        colors=colors,
        marker_size=pattern.get("marker_size", 7),
        line_width_pt=2.25,
        show_labels=component.chrome.data_labels.show,
        num_fmt=component.chrome.data_labels.format or "0%",
        hide_axes=False,          # keep axes structurally present
        scale_min=scale_min,
        scale_max=scale_max,
    )

    # Hide value-axis tick labels (redundant when data labels show exact values)
    # but leave category axis labels visible (period names are meaningful).
    # Only applies when spec has value_axis.show=False (the default); if spec
    # explicitly sets show=True, tick labels remain visible.
    if not component.chrome.value_axis.show:
        from lxml import etree
        C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
        val_ax_el = chart.value_axis._element
        # Set tickLblPos="none" — hides tick labels without deleting the axis
        tick_lbl_pos = val_ax_el.find(f"{{{C_NS}}}tickLblPos")
        if tick_lbl_pos is None:
            tick_lbl_pos = etree.SubElement(val_ax_el, f"{{{C_NS}}}tickLblPos")
        tick_lbl_pos.set("val", "none")

    return chart


def _render_xy_scatter_abacus(
    slide, component: ChartComponent, rect: tuple, brand, context, deck_ref,
):
    """For scatter: series.values is [(x, y), ...] OR parallel [x, x, ...] and
    a second array; our spec convention is `values: list[float]` with one
    value per category, interpreted as Y values with X = category index.
    """
    pattern = CHART_PATTERNS["xy_scatter_abacus"]
    colors = _series_colors(component, brand, context, deck_ref)
    cats = component.data.categories
    n_cats = len(cats)

    # Each series contributes (name, x_vals, y_vals). By convention we treat
    # category index as the X axis and the values as Y.
    series_list: list[tuple[str, list[float], list[float]]] = []
    for s in component.data.series:
        # If values are (x, y) pairs, unpack; else treat as Y with category index
        if s.values and isinstance(s.values[0], (list, tuple)) and len(s.values[0]) == 2:
            x_vals = [float(pair[0]) for pair in s.values]
            y_vals = [float(pair[1]) for pair in s.values]
        else:
            x_vals = [float(i) for i in range(n_cats)]
            y_vals = [float(v) for v in s.values]
        series_list.append((s.name, x_vals, y_vals))

    cf, chart = add_scatter_chart(
        slide, series_list,
        rect[0], rect[1], rect[2], rect[3],
        colors=colors,
        marker_size=pattern.get("marker_size", 10),
        x_min=0, x_max=max(1.0, float(n_cats - 1) if n_cats > 1 else 1.0),
        y_min=0, y_max=1.0,
    )
    return chart


def _render_doughnut_default(
    slide, component: ChartComponent, rect: tuple, brand, context, deck_ref,
):
    """Doughnut takes ONE series of per-category values. Colors apply per SLICE
    (one per category), not per series. If the spec only provides 1 series color,
    pad with the brand's palette so every slice is visible.
    """
    cats = component.data.categories
    if len(component.data.series) != 1:
        raise ValueError(
            f"doughnut_default expects exactly 1 series, got {len(component.data.series)}"
        )
    series = component.data.series[0]
    chart_data = CategoryChartData()
    chart_data.categories = cats
    chart_data.add_series(series.name, list(series.values))
    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.DOUGHNUT,
        Inches(rect[0]), Inches(rect[1]), Inches(rect[2]), Inches(rect[3]),
        chart_data,
    )
    chart = cf.chart

    # Build per-slice colors: spec's series colors first, then brand palette fallback
    colors = _series_colors(component, brand, context, deck_ref)
    n_slices = len(cats)

    if len(colors) < n_slices:
        # Pad from brand palette (primary_current, primary_prior, secondary)
        # then from the observed _observed_palette hex strings
        fallback: list[RGBColor] = []
        if brand is not None:
            for role in ("primary_current", "primary_prior", "secondary"):
                val = brand.get(role)
                if isinstance(val, RGBColor) and val not in colors:
                    fallback.append(val)
            for hex_str in (brand.get("_observed_palette") or []):
                h = hex_str.lstrip("#")
                if len(h) == 6:
                    c = RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
                    if c not in colors + fallback:
                        fallback.append(c)
        # Universal safe defaults if brand didn't supply enough colors
        universal = [
            RGBColor(0x4F, 0x81, 0xBD),  # blue
            RGBColor(0xC0, 0x50, 0x4D),  # red
            RGBColor(0x9B, 0xBB, 0x59),  # green
            RGBColor(0x80, 0x64, 0xA2),  # purple
            RGBColor(0x4B, 0xAC, 0xC6),  # cyan
            RGBColor(0xF7, 0x96, 0x46),  # orange
        ]
        needed = n_slices - len(colors) - len(fallback)
        if needed > 0:
            fallback.extend(universal[:needed])
        colors = colors + fallback[:n_slices - len(colors)]

    if colors:
        set_pie_slice_colors(chart, colors[:n_slices])
        set_donut_hole_size(chart, CHART_PATTERNS["doughnut_default"].get("hole_size", 50))
    return chart


def _render_single_bar(
    slide, component: ChartComponent, rect: tuple, brand, context, deck_ref,
):
    if len(component.data.series) != 1:
        raise ValueError(
            f"single_bar expects exactly 1 series, got {len(component.data.series)}"
        )
    series = component.data.series[0]
    colors = _series_colors(component, brand, context, deck_ref)
    cf, chart = add_single_bar_chart(
        slide, component.data.categories, list(series.values),
        rect[0], rect[1], rect[2], rect[3],
        fill_color=colors[0],
        gap=component.chrome.gap_width or CHART_PATTERNS["single_bar"]["gap"],
    )
    return chart


def _render_clustered_bar_legacy(
    slide, component: ChartComponent, rect: tuple, brand, context, deck_ref,
):
    """Legacy 'clustered_bar' — same as bar_clustered_horizontal with slightly
    different defaults (overlap=-10 vs 0). Kept for backward compat."""
    return _render_bar_clustered_horizontal(slide, component, rect, brand, context, deck_ref)


def _render_stacked_bar_legacy(
    slide, component: ChartComponent, rect: tuple, brand, context, deck_ref,
):
    """Legacy 'stacked_bar' — horizontal stacked (not percent)."""
    cats = component.data.categories
    series_list = [(s.name, list(s.values)) for s in component.data.series]
    colors = _series_colors(component, brand, context, deck_ref)

    chart_data = CategoryChartData()
    chart_data.categories = cats
    for name, vals in series_list:
        chart_data.add_series(name, vals)
    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_STACKED,
        Inches(rect[0]), Inches(rect[1]), Inches(rect[2]), Inches(rect[3]),
        chart_data,
    )
    chart = cf.chart
    set_plot_area_gap(chart, component.chrome.gap_width or CHART_PATTERNS["stacked_bar"]["gap"])
    invert_cat_axis(chart)
    for i, s in enumerate(chart.series):
        if i < len(colors):
            set_series_color(s, colors[i])
        set_series_no_border(s)
    return chart


_CHART_RENDERERS: dict[str, Callable] = {
    "bar_clustered_horizontal":    _render_bar_clustered_horizontal,
    "bar_stacked_100_horizontal":  _render_bar_stacked_100_horizontal,
    "column_stacked_100_vertical": _render_column_stacked_100_vertical,
    "column_clustered_vertical":   _render_column_clustered_vertical,
    "line_markers_trended":        _render_line_markers_trended,
    "xy_scatter_abacus":           _render_xy_scatter_abacus,
    "doughnut_default":            _render_doughnut_default,
    "single_bar":                  _render_single_bar,
    "clustered_bar":               _render_clustered_bar_legacy,
    "stacked_bar":                 _render_stacked_bar_legacy,
}


# ─────────────────────────────────────────────────────────────────────────────
# OOXML strict-mode schema enforcement
# ─────────────────────────────────────────────────────────────────────────────
#
# python-pptx is tolerant of element order within <c:ser> and related elements;
# PowerPoint is NOT. Out-of-order <c:dLbls> (appended after <c:val>) triggers
# PowerPoint's Repair prompt on open. These helpers enforce the CT_*Ser
# element order after the chart is constructed.

_C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def _child_local_name(el) -> str:
    """Return the local tag name (strips namespace)."""
    tag = el.tag
    return tag.split("}")[-1] if "}" in tag else tag


# ── Element order per CT_*Ser (children of <c:ser>) ────────────────────────
# Union order across CT_LineSer, CT_BarSer, CT_PieSer, CT_ScatterSer.
#
# NOTE: dPt is intentionally OMITTED from this order. Per ECMA-376,
# dPt should appear before dLbls — but PowerPoint's doughnut/pie loader
# rejects dPt in that position. python-pptx + set_pie_slice_colors append
# dPt at the end of <c:ser> (after val), and PowerPoint accepts that.
# By omitting dPt from the order tuple, _reorder_children leaves it
# wherever python-pptx put it (default index = len(order) = end).
_SER_CHILD_ORDER: tuple[str, ...] = (
    "idx", "order", "tx", "spPr",
    "invertIfNegative", "pictureOptions", "marker", "explosion",
    # dPt intentionally omitted — stays where python-pptx put it
    "dLbls",
    "trendline", "errBars",
    "xVal", "yVal",       # scatter series
    "cat", "val",
    "smooth", "shape", "bubble3D", "bubbleSize",
    "extLst",
)

# ── Element order per CT_*Chart (children of <c:barChart>, <c:lineChart>, …) ──
# Union order that covers CT_BarChart, CT_LineChart, CT_PieChart, CT_DoughnutChart,
# CT_ScatterChart, CT_Area/Area3DChart, CT_SurfaceChart, etc.
# Based on ECMA-376 Part 1 §21.2.2 — children MUST be in this order for PowerPoint
# to load without triggering the Repair prompt.
_CHART_ELEMENT_CHILD_ORDER: tuple[str, ...] = (
    # Chart-type-specific preambles
    "barDir", "grouping", "scatterStyle", "radarStyle", "ofPieType",
    "varyColors",
    # All series
    "ser",
    # Post-series chrome
    "dLbls", "dropLines", "hiLowLines", "upDownBars",
    "marker", "smooth",
    "gapWidth", "overlap", "serLines",
    "firstSliceAng", "holeSize",
    "bubbleScale", "showNegBubbles", "sizeRepresents",
    "bandFmts", "wireframe", "floor", "sideWall", "backWall",
    # Axis references (always last)
    "axId",
    "extLst",
)

# ── Element order per CT_PlotArea (children of <c:plotArea>) ────────────────
# layout → chart-type elements → axis elements → dTable → spPr
_PLOTAREA_CHILD_ORDER: tuple[str, ...] = (
    "layout",
    # Chart-type elements (any one or combination)
    "areaChart", "area3DChart", "line3DChart", "lineChart", "stockChart",
    "radarChart", "scatterChart", "pieChart", "pie3DChart", "doughnutChart",
    "bar3DChart", "barChart", "ofPieChart", "surface3DChart", "surfaceChart",
    "bubbleChart",
    # Axis elements
    "valAx", "catAx", "dateAx", "serAx",
    "dTable", "spPr",
    "extLst",
)

# ── Element order per CT_DLbls (children of <c:dLbls>) ─────────────────────
# Per ECMA-376 Part 1 §21.2.2.49 — applies to BOTH series-level and
# chart-type-level dLbls. python-pptx creates showVal/showCatName first,
# then our lxml helpers append numFmt/dLblPos at the end — wrong order.
_DLBLS_CHILD_ORDER: tuple[str, ...] = (
    "dLbl",
    "delete",
    "numFmt",
    "spPr",
    "txPr",
    "dLblPos",
    "showLegendKey",
    "showVal",
    "showCatName",
    "showSerName",
    "showPercent",
    "showBubbleSize",
    "separator",
    "showLeaderLines",
    "leaderLines",
    "extLst",
)

# Chart-type element tag names — union of possibilities
_CHART_TYPE_ELEMENTS: frozenset[str] = frozenset({
    "areaChart", "area3DChart", "line3DChart", "lineChart", "stockChart",
    "radarChart", "scatterChart", "pieChart", "pie3DChart", "doughnutChart",
    "bar3DChart", "barChart", "ofPieChart", "surface3DChart", "surfaceChart",
    "bubbleChart",
})


def _reorder_children(parent, order: tuple[str, ...]) -> None:
    """Reorder children of `parent` element according to `order` tuple.
    Unknown children (not in order) are appended at the end in original order.
    Only rewrites if order actually changed.
    """
    order_index = {name: i for i, name in enumerate(order)}
    children = list(parent)

    def _key(el):
        return order_index.get(_child_local_name(el), len(order))

    sorted_children = sorted(children, key=_key)
    if [c.tag for c in children] != [c.tag for c in sorted_children]:
        for c in children:
            parent.remove(c)
        for c in sorted_children:
            parent.append(c)


def _enforce_ser_child_order(chart) -> None:
    """Fix OOXML element order throughout the chart so PowerPoint opens it
    without triggering the Repair prompt.

    python-pptx + lxml helpers append elements rather than inserting in
    schema-declared order. PowerPoint's strict loader rejects out-of-order
    elements in three places:

      1. Children of <c:plotArea> (layout → chart-type → axes → …)
      2. Children of <c:barChart>/<c:lineChart>/<c:doughnutChart>/etc.
         (barDir/grouping → ser → dLbls → marker/smooth/gapWidth/holeSize → axId)
      3. Children of <c:ser> (idx/order/tx/spPr/marker → dLbls → cat/val/smooth)

    Note: python-pptx's `chart.plots[i]._element` IS the chart-type element
    (e.g. <c:lineChart>), NOT the <c:plotArea>. We walk up to the plot area
    via getparent() so reordering is applied to the right element at each level.
    """
    # Find the <c:plotArea> — walk up from the first plot's element
    if not chart.plots:
        return
    plot_area_el = chart.plots[0]._element.getparent()
    if plot_area_el is None or _child_local_name(plot_area_el) != "plotArea":
        return

    # Level 1: reorder plot area children (layout → chart-type elements → axes → …)
    _reorder_children(plot_area_el, _PLOTAREA_CHILD_ORDER)

    # Level 2: reorder each chart-type element's children
    for child in list(plot_area_el):
        if _child_local_name(child) in _CHART_TYPE_ELEMENTS:
            _reorder_children(child, _CHART_ELEMENT_CHILD_ORDER)

    # Level 3: reorder every <c:ser> element's children
    for ser_el in plot_area_el.iter("{%s}ser" % _C_NS):
        _reorder_children(ser_el, _SER_CHILD_ORDER)

    # Level 4: reorder every <c:dLbls> element's children
    # python-pptx creates showVal/showCatName first, then our lxml helpers
    # append numFmt/dLblPos at the end — wrong per CT_DLbls schema.
    for dlbls_el in plot_area_el.iter("{%s}dLbls" % _C_NS):
        _reorder_children(dlbls_el, _DLBLS_CHILD_ORDER)


# ─────────────────────────────────────────────────────────────────────────────
# Component renderers (chart + label_table + value_table + delta_column + ...)
# ─────────────────────────────────────────────────────────────────────────────


def _render_chart_component(
    slide, component: ChartComponent, spec: SlideSpec,
    brand: Optional[dict], context: Optional[dict], deck_ref: Optional[dict],
    namer: ShapeNamer,
):
    rect = resolve_position(component.position, LAYOUTS.get(spec.layout, {}), spec.layout)
    renderer = _CHART_RENDERERS.get(component.chart_pattern)
    if renderer is None:
        raise KeyError(
            f"chart_pattern {component.chart_pattern!r} has no renderer. "
            f"Available: {sorted(_CHART_RENDERERS.keys())}"
        )
    chart = renderer(slide, component, rect, brand, context, deck_ref)
    _apply_chart_chrome(chart, component, brand, context, deck_ref)
    _apply_data_labels(chart, component, brand, context, deck_ref)
    # Remove chart-type-level dLbls that python-pptx auto-created as a side
    # effect of accessing series.data_labels (line charts specifically don't
    # want this per raw python-pptx output).
    _remove_auto_dLbls_on_chart_type(chart)
    # PowerPoint enforces OOXML element order strictly; python-pptx doesn't.
    # After all element mutations, reorder children to match the schema.
    _enforce_ser_child_order(chart)
    # Chart shapes in python-pptx are graphic frames — name the frame
    chart_frame = slide.shapes[-1]
    namer.name(chart_frame)
    return chart


def _render_label_table_component(
    slide, component: LabelTableComponent, spec: SlideSpec,
    brand: Optional[dict], context: Optional[dict], deck_ref: Optional[dict],
    namer: ShapeNamer,
):
    """Render a label table (leftmost column of bar_clustered_horizontal layouts)."""
    rect = resolve_position(component.position, LAYOUTS.get(spec.layout, {}), spec.layout)
    left, top, width, height = rect
    n = len(component.labels)

    if n == 0:
        return None

    # Font: spec override > brand body font > pptx default
    font_name = component.font_name or (brand.get("font_body") if brand else None)

    # Use python-pptx's native table with alternating rows
    tbl_shape = slide.shapes.add_table(n, 1, Inches(left), Inches(top), Inches(width), Inches(height))
    tbl = tbl_shape.table
    row_h = height / n
    for i, row in enumerate(tbl.rows):
        row.height = Inches(row_h)
    tbl.columns[0].width = Inches(width)

    from pptx.dml.color import RGBColor as _RGB
    from slidegen.pptx_utils.lxml_helpers import cell_vcenter
    C_LBGREY = _RGB(0xF4, 0xF4, 0xF4)
    C_WHITE = _RGB(0xFF, 0xFF, 0xFF)
    C_TEXT = _RGB(0x40, 0x40, 0x40)

    for i, label in enumerate(component.labels):
        cell = tbl.cell(i, 0)
        # Alternate rows (grey first per real-deck convention)
        if component.alternating_rows:
            fill_color = C_LBGREY if i % 2 == 0 else C_WHITE
        else:
            fill_color = C_WHITE
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill_color
        cell.margin_left = Inches(0.08)
        cell.margin_right = Inches(0.05)
        cell.margin_top = Inches(0.02)
        cell.margin_bottom = Inches(0.02)

        tf = cell.text_frame
        tf.word_wrap = component.wrap
        tf.text = label
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        for run in p.runs:
            run.font.size = Pt(component.font_size_pt)
            run.font.color.rgb = C_TEXT
            if font_name:
                run.font.name = font_name
        # Vertically center text in the cell (real-deck convention)
        cell_vcenter(cell)

    namer.name(tbl_shape)
    return tbl_shape


def _render_value_table_component(
    slide, component: ValueTableComponent, spec: SlideSpec,
    brand: Optional[dict], context: Optional[dict], deck_ref: Optional[dict],
    namer: ShapeNamer,
):
    rect = resolve_position(component.position, LAYOUTS.get(spec.layout, {}), spec.layout)
    n_rows = len(component.rows) + (1 if component.headers else 0)
    n_cols = max((len(r) for r in component.rows), default=0) or len(component.headers)
    if n_rows == 0 or n_cols == 0:
        return None

    left, top, width, height = rect
    tbl_shape = slide.shapes.add_table(n_rows, n_cols, Inches(left), Inches(top), Inches(width), Inches(height))
    tbl = tbl_shape.table
    row_h = height / n_rows
    col_w = width / n_cols
    for row in tbl.rows:
        row.height = Inches(row_h)
    for col in tbl.columns:
        col.width = Inches(col_w)

    from pptx.dml.color import RGBColor as _RGB
    C_LBGREY = _RGB(0xF4, 0xF4, 0xF4)
    C_WHITE = _RGB(0xFF, 0xFF, 0xFF)
    C_HDRGREY = _RGB(0x40, 0x40, 0x40)
    C_TEXT = _RGB(0x40, 0x40, 0x40)

    row_offset = 0
    if component.headers:
        for c, header in enumerate(component.headers):
            cell = tbl.cell(0, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_HDRGREY
            cell.text_frame.text = header
            for run in cell.text_frame.paragraphs[0].runs:
                run.font.size = Pt(8)
                run.font.bold = True
                run.font.color.rgb = C_WHITE
        row_offset = 1

    for r, row_data in enumerate(component.rows):
        tbl_row = r + row_offset
        for c, cell_val in enumerate(row_data):
            cell = tbl.cell(tbl_row, c)
            if component.alternating_rows:
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_LBGREY if r % 2 == 0 else C_WHITE
            cell.text_frame.text = str(cell_val)
            # Only apply font defaults if not a pass-through table (deck-reader)
            if component.alternating_rows:
                for run in cell.text_frame.paragraphs[0].runs:
                    run.font.size = Pt(8)
                    run.font.color.rgb = C_TEXT

    namer.name(tbl_shape)
    return tbl_shape


def _render_delta_column_component(
    slide, component: DeltaColumnComponent, spec: SlideSpec,
    brand: Optional[dict], context: Optional[dict], deck_ref: Optional[dict],
    namer: ShapeNamer,
):
    rect = resolve_position(component.position, LAYOUTS.get(spec.layout, {}), spec.layout)
    left, top, width, height = rect

    # Convert delta values to the display format expected by add_delta_col
    # (list of floats in same order as chart categories)
    values = [float(v) if v is not None else None for v in component.values]

    # Font sizing: spec values are authoritative. Brand body font is the default
    # when component doesn't override.
    font_name = component.font_name or (brand.get("font_body") if brand else None)

    tbl_shape = add_delta_col(
        slide, values,
        left, top, width, height,
        header=component.header,
        font_size_pt=component.font_size_pt,
        header_font_size_pt=component.header_font_size_pt,
        font_name=font_name,
    )
    namer.name(tbl_shape)
    return tbl_shape


def _render_callout_component(
    slide, component: CalloutComponent, spec: SlideSpec,
    brand: Optional[dict], context: Optional[dict], deck_ref: Optional[dict],
    namer: ShapeNamer,
):
    rect = resolve_position(component.position, LAYOUTS.get(spec.layout, {}), spec.layout)
    text_parts = []
    if component.theme:
        text_parts.append(f"[{component.theme}]")
    text_parts.append(component.text)
    if component.attribution:
        text_parts.append(f"— {component.attribution}")
    full_text = "\n".join(text_parts)

    shape = callout_box(
        slide, rect[0], rect[1], rect[2], rect[3],
        text=full_text,
    )
    namer.name(shape)
    return shape


def _render_image_component(
    slide, component: ImageComponent, spec: SlideSpec,
    brand: Optional[dict], context: Optional[dict], deck_ref: Optional[dict],
    namer: ShapeNamer,
):
    rect = resolve_position(component.position, LAYOUTS.get(spec.layout, {}), spec.layout)
    src = component.source
    if not src or not Path(src).exists():
        # Skip silently with warning — missing images are common in templates
        return None
    shape = insert_image(slide, src, rect[0], rect[1], rect[2], rect[3])
    namer.name(shape)
    return shape


def _render_textbox_component(
    slide, component: TextboxComponent, spec: SlideSpec,
    brand: Optional[dict], context: Optional[dict], deck_ref: Optional[dict],
    namer: ShapeNamer,
):
    rect = resolve_position(component.position, LAYOUTS.get(spec.layout, {}), spec.layout)
    align_map = {
        "left":   PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right":  PP_ALIGN.RIGHT,
    }
    color = None
    if component.font_color:
        try:
            color = resolve_color(component.font_color, brand=brand, context=context, deck_ref=deck_ref)
        except ColorResolutionError:
            color = None
    shape = textbox(
        slide, component.text,
        rect[0], rect[1], rect[2], rect[3],
        fsize=component.font_size_pt,
        bold=component.bold,
        italic=component.italic,
        color=color,
        align=align_map.get(component.alignment, PP_ALIGN.LEFT),
    )
    namer.name(shape)
    return shape


_COMPONENT_RENDERERS: dict[str, Callable] = {
    "chart":         _render_chart_component,
    "label_table":   _render_label_table_component,
    "value_table":   _render_value_table_component,
    "delta_column":  _render_delta_column_component,
    "callout":       _render_callout_component,
    "image":         _render_image_component,
    "textbox":       _render_textbox_component,
}


# ─────────────────────────────────────────────────────────────────────────────
# Slide chrome (headline, subheadline, footer, section bar)
# ─────────────────────────────────────────────────────────────────────────────


def _apply_headline(slide, spec: SlideSpec, brand: Optional[dict]):
    """Apply headline using pptx_utils slide_header (adds accent + module label too)."""
    font = brand.get("font_heading") if brand else None
    slide_header(slide, spec.headline.text, font=font)


def _apply_subheadline(slide, spec: SlideSpec, brand: Optional[dict]):
    """Apply subheadline as a textbox below the headline."""
    if spec.subheadline is None:
        return
    font = brand.get("font_body") if brand else None
    textbox(
        slide, spec.subheadline.text,
        0.30, 0.80, 12.70, 0.40,
        fsize=10, italic=False, font=font,
    )


def _apply_footer(slide, spec: SlideSpec, brand: Optional[dict]):
    if spec.footer is None:
        return
    font = brand.get("font_body") if brand else None
    slide_footer(slide, spec.footer.text, font=font)


def _apply_section_bar(slide, spec: SlideSpec, brand: Optional[dict]):
    if not spec.section:
        return
    section_header_bar(slide, spec.section, top=1.40)


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────


def render_slide(
    spec: SlideSpec,
    prs: Optional[Presentation] = None,
    template_path: Optional[str] = None,
    context: Optional[dict] = None,
    deck_ref: Optional[dict] = None,
) -> tuple[Presentation, Any]:
    """Render one SlideSpec into a new slide.

    Validates spec first (strict mode — raises SpecValidationError on failure).
    Resolves brand via get_brand() if spec.brand is set.
    Resolves colors lazily via resolve_color() during component rendering.

    Returns (prs, slide) so the caller can continue composing or save.
    """
    # 1. Gate — spec must be valid
    validate_spec(spec, strict=True)

    # 2. Create or reuse presentation
    if prs is None:
        prs = Presentation(template_path) if template_path else Presentation()
        prs.slide_width = Inches(_SLIDE_W_IN)
        prs.slide_height = Inches(_SLIDE_H_IN)

    # 3. Resolve brand
    brand: Optional[dict] = None
    if spec.brand:
        try:
            brand = get_brand(spec.brand)
        except KeyError:
            brand = None  # validator already surfaced a warning; spec may use only hex tokens

    # 4. Add blank slide
    blank_layout_idx = min(_BLANK_LAYOUT_IDX, len(prs.slide_layouts) - 1)
    slide = prs.slides.add_slide(prs.slide_layouts[blank_layout_idx])
    namer = ShapeNamer(spec.slide_index)

    # 5. Apply chrome — skip for deck-reader specs (they have their own headline
    # as a component, and adding SlideGen chrome on top creates duplicates)
    is_deck_reader = (spec.metadata and spec.metadata.created_by and
                      "deck-reader" in spec.metadata.created_by)
    if not is_deck_reader:
        _apply_section_bar(slide, spec, brand)
        _apply_headline(slide, spec, brand)
        _apply_subheadline(slide, spec, brand)
        _apply_footer(slide, spec, brand)

    # 6. Render components
    for component in spec.components:
        renderer = _COMPONENT_RENDERERS.get(component.type)
        if renderer is None:
            raise KeyError(
                f"Component type {component.type!r} has no renderer. "
                f"Available: {sorted(_COMPONENT_RENDERERS.keys())}"
            )
        renderer(slide, component, spec, brand, context, deck_ref, namer)

    return prs, slide


def render_deck(
    specs: list[SlideSpec],
    output_path: str | Path,
    template_path: Optional[str] = None,
    context: Optional[dict] = None,
    deck_ref: Optional[dict] = None,
) -> Presentation:
    """Render a list of specs into one PPTX file. Each spec becomes one slide.

    Validates all specs upfront — fails before writing if any is invalid.
    """
    # Pre-validate all specs
    all_errors: dict[str, list[str]] = {}
    for spec in specs:
        errs = validate_spec(spec)
        if errs:
            all_errors[spec.slide_id] = errs
    if all_errors:
        raise ValueError(f"Deck has invalid specs: {all_errors}")

    # Build sequentially
    prs: Optional[Presentation] = None
    for spec in specs:
        prs, _slide = render_slide(spec, prs=prs, template_path=template_path,
                                   context=context, deck_ref=deck_ref)
        template_path = None  # only use template on first slide

    assert prs is not None, "render_deck called with empty specs list"
    prs.save(str(output_path))
    return prs


def render_spec_to_file(
    spec: SlideSpec,
    output_path: str | Path,
    template_path: Optional[str] = None,
    context: Optional[dict] = None,
    deck_ref: Optional[dict] = None,
) -> Path:
    """Convenience: render one spec, save to PPTX file, return path."""
    prs, _ = render_slide(spec, template_path=template_path, context=context, deck_ref=deck_ref)
    prs.save(str(output_path))
    return Path(output_path)
