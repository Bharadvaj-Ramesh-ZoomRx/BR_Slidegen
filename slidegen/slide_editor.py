"""
slide_editor.py — Edit-mode primitive for SlideSpec.

Takes a SlideSpec + an edit instruction dict, produces a modified SlideSpec.
Supports a whitelist of edit actions; rejects unsupported ones.
Spec-validates before returning.

Usage:
    from slidegen.slide_editor import edit_slide

    updated = edit_slide(spec, {"action": "set_headline_text", "value": "New Title"})
"""
from __future__ import annotations

import copy
from typing import Any

from slidegen.slide_spec.schema import (
    SlideSpec, HeadlineSpec, FooterSpec,
    ChartComponent, DeltaColumnComponent, LabelTableComponent,
    ValueTableComponent, TextboxComponent, CalloutComponent,
    ChartChrome, DataLabelsSpec, LegendSpec, AxisSpec,
    SUPPORTED_CHART_PATTERNS,
)
from slidegen.slide_spec.validator import validate_spec, SpecValidationError


class SlideEditError(Exception):
    """Raised when an edit action fails."""
    pass


# ── Supported edit actions ────────────────────────────────────────────────────
# Each action maps to a handler function: f(spec, instruction) -> None
# The handler mutates the spec in-place (caller deep-copies first).

SUPPORTED_ACTIONS = {
    "set_headline_text",
    "set_headline_style",
    "set_headline_color",
    "set_subheadline_text",
    "set_footer_text",
    "set_brand",
    "set_section",
    "set_layout",
    "set_chart_pattern",
    "change_series_color",
    "change_series_name",
    "set_data_labels_format",
    "set_data_labels_position",
    "set_data_labels_show",
    "set_legend_show",
    "set_legend_position",
    "set_gridlines",
    "set_value_axis_show",
    "set_value_axis_format",
    "set_delta_format",
    "set_delta_colors",
    "set_delta_header",
    "reorder_categories",
    "set_slide_id",
    "set_slide_index",
    "set_speaker_notes",
}


def edit_slide(spec: SlideSpec, instruction: dict[str, Any]) -> SlideSpec:
    """Apply an edit instruction to a SlideSpec.

    Args:
        spec: The source SlideSpec.
        instruction: Dict with at minimum {"action": "<action_name>", ...}
            See SUPPORTED_ACTIONS for the full list.

    Returns:
        Modified SlideSpec (deep copy — original is not mutated).

    Raises:
        SlideEditError: If the action is unsupported or arguments are invalid.
        SpecValidationError: If the edited spec fails validation.
    """
    # Validate input spec first
    errors = validate_spec(spec)
    if errors:
        raise SpecValidationError(errors)

    action = instruction.get("action")
    if not action:
        raise SlideEditError("Edit instruction must have an 'action' key")

    if action not in SUPPORTED_ACTIONS:
        raise SlideEditError(
            f"Unsupported edit action: {action!r}. "
            f"Supported: {sorted(SUPPORTED_ACTIONS)}"
        )

    # Deep copy to avoid mutating the original
    edited = copy.deepcopy(spec)

    # Dispatch to handler
    handler = _ACTION_HANDLERS.get(action)
    if handler is None:
        raise SlideEditError(f"No handler registered for action: {action!r}")

    handler(edited, instruction)

    # Validate the edited spec
    validate_spec(edited, strict=True)

    return edited


# ── Handler implementations ──────────────────────────────────────────────────


def _set_headline_text(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    if value is None:
        raise SlideEditError("set_headline_text requires 'value'")
    spec.headline.text = str(value)


def _set_headline_style(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    if value not in ("default", "red_accent", "neutral"):
        raise SlideEditError(
            f"set_headline_style value must be default|red_accent|neutral, got {value!r}"
        )
    spec.headline.style = value


def _set_headline_color(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    if value is None:
        raise SlideEditError("set_headline_color requires 'value' (color token)")
    spec.headline.color = str(value)


def _set_subheadline_text(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    if value is None:
        raise SlideEditError("set_subheadline_text requires 'value'")
    if spec.subheadline is None:
        spec.subheadline = HeadlineSpec(text=str(value))
    else:
        spec.subheadline.text = str(value)


def _set_footer_text(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    if value is None:
        raise SlideEditError("set_footer_text requires 'value'")
    if spec.footer is None:
        spec.footer = FooterSpec(text=str(value))
    else:
        spec.footer.text = str(value)


def _set_brand(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    if value is None:
        raise SlideEditError("set_brand requires 'value'")
    spec.brand = str(value)


def _set_section(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    if value is None:
        raise SlideEditError("set_section requires 'value'")
    spec.section = str(value)


def _set_layout(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    if value is None:
        raise SlideEditError("set_layout requires 'value'")
    spec.layout = str(value)


def _set_chart_pattern(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    if value not in SUPPORTED_CHART_PATTERNS:
        raise SlideEditError(
            f"set_chart_pattern value {value!r} not in SUPPORTED_CHART_PATTERNS. "
            f"Supported: {SUPPORTED_CHART_PATTERNS}"
        )
    comp_idx = instr.get("component_index", 0)
    charts = [c for c in spec.components if isinstance(c, ChartComponent)]
    if comp_idx >= len(charts):
        raise SlideEditError(
            f"component_index={comp_idx} but only {len(charts)} chart(s) on slide"
        )
    charts[comp_idx].chart_pattern = value


def _change_series_color(spec: SlideSpec, instr: dict) -> None:
    series_index = instr.get("series_index")
    value = instr.get("value")
    comp_idx = instr.get("component_index", 0)

    if series_index is None:
        raise SlideEditError("change_series_color requires 'series_index'")
    if value is None:
        raise SlideEditError("change_series_color requires 'value' (color token)")

    charts = [c for c in spec.components if isinstance(c, ChartComponent)]
    if comp_idx >= len(charts):
        raise SlideEditError(f"component_index={comp_idx} out of range")

    chart = charts[comp_idx]
    if series_index >= len(chart.data.series):
        raise SlideEditError(
            f"series_index={series_index} but chart has {len(chart.data.series)} series"
        )
    chart.data.series[series_index].color = str(value)


def _change_series_name(spec: SlideSpec, instr: dict) -> None:
    series_index = instr.get("series_index")
    value = instr.get("value")
    comp_idx = instr.get("component_index", 0)

    if series_index is None:
        raise SlideEditError("change_series_name requires 'series_index'")
    if value is None:
        raise SlideEditError("change_series_name requires 'value'")

    charts = [c for c in spec.components if isinstance(c, ChartComponent)]
    if comp_idx >= len(charts):
        raise SlideEditError(f"component_index={comp_idx} out of range")

    chart = charts[comp_idx]
    if series_index >= len(chart.data.series):
        raise SlideEditError(
            f"series_index={series_index} but chart has {len(chart.data.series)} series"
        )
    chart.data.series[series_index].name = str(value)


def _set_data_labels_format(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    comp_idx = instr.get("component_index", 0)
    if value is None:
        raise SlideEditError("set_data_labels_format requires 'value'")
    charts = [c for c in spec.components if isinstance(c, ChartComponent)]
    if comp_idx >= len(charts):
        raise SlideEditError(f"component_index={comp_idx} out of range")
    charts[comp_idx].chrome.data_labels.format = str(value)


def _set_data_labels_position(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    comp_idx = instr.get("component_index", 0)
    if value not in ("inEnd", "outEnd", "ctr", "above"):
        raise SlideEditError(
            f"set_data_labels_position value must be inEnd|outEnd|ctr|above, got {value!r}"
        )
    charts = [c for c in spec.components if isinstance(c, ChartComponent)]
    if comp_idx >= len(charts):
        raise SlideEditError(f"component_index={comp_idx} out of range")
    charts[comp_idx].chrome.data_labels.position = value


def _set_data_labels_show(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    comp_idx = instr.get("component_index", 0)
    if not isinstance(value, bool):
        raise SlideEditError("set_data_labels_show requires boolean 'value'")
    charts = [c for c in spec.components if isinstance(c, ChartComponent)]
    if comp_idx >= len(charts):
        raise SlideEditError(f"component_index={comp_idx} out of range")
    charts[comp_idx].chrome.data_labels.show = value


def _set_legend_show(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    comp_idx = instr.get("component_index", 0)
    if not isinstance(value, bool):
        raise SlideEditError("set_legend_show requires boolean 'value'")
    charts = [c for c in spec.components if isinstance(c, ChartComponent)]
    if comp_idx >= len(charts):
        raise SlideEditError(f"component_index={comp_idx} out of range")
    charts[comp_idx].chrome.legend.show = value


def _set_legend_position(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    comp_idx = instr.get("component_index", 0)
    if value not in ("top", "bottom", "left", "right"):
        raise SlideEditError(
            f"set_legend_position value must be top|bottom|left|right, got {value!r}"
        )
    charts = [c for c in spec.components if isinstance(c, ChartComponent)]
    if comp_idx >= len(charts):
        raise SlideEditError(f"component_index={comp_idx} out of range")
    charts[comp_idx].chrome.legend.position = value


def _set_gridlines(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    comp_idx = instr.get("component_index", 0)
    if not isinstance(value, bool):
        raise SlideEditError("set_gridlines requires boolean 'value'")
    charts = [c for c in spec.components if isinstance(c, ChartComponent)]
    if comp_idx >= len(charts):
        raise SlideEditError(f"component_index={comp_idx} out of range")
    charts[comp_idx].chrome.gridlines = value


def _set_value_axis_show(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    comp_idx = instr.get("component_index", 0)
    if not isinstance(value, bool):
        raise SlideEditError("set_value_axis_show requires boolean 'value'")
    charts = [c for c in spec.components if isinstance(c, ChartComponent)]
    if comp_idx >= len(charts):
        raise SlideEditError(f"component_index={comp_idx} out of range")
    charts[comp_idx].chrome.value_axis.show = value


def _set_value_axis_format(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    comp_idx = instr.get("component_index", 0)
    if value is None:
        raise SlideEditError("set_value_axis_format requires 'value'")
    charts = [c for c in spec.components if isinstance(c, ChartComponent)]
    if comp_idx >= len(charts):
        raise SlideEditError(f"component_index={comp_idx} out of range")
    charts[comp_idx].chrome.value_axis.format = str(value)


def _set_delta_format(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    comp_idx = instr.get("component_index", 0)
    if value not in ("delta_pp", "delta_pct", "delta_abs"):
        raise SlideEditError(
            f"set_delta_format value must be delta_pp|delta_pct|delta_abs, got {value!r}"
        )
    deltas = [c for c in spec.components if isinstance(c, DeltaColumnComponent)]
    if comp_idx >= len(deltas):
        raise SlideEditError(f"component_index={comp_idx} out of range")
    deltas[comp_idx].format = value


def _set_delta_colors(spec: SlideSpec, instr: dict) -> None:
    comp_idx = instr.get("component_index", 0)
    positive = instr.get("positive_color")
    negative = instr.get("negative_color")
    deltas = [c for c in spec.components if isinstance(c, DeltaColumnComponent)]
    if comp_idx >= len(deltas):
        raise SlideEditError(f"component_index={comp_idx} out of range")
    if positive is not None:
        deltas[comp_idx].positive_color = str(positive)
    if negative is not None:
        deltas[comp_idx].negative_color = str(negative)


def _set_delta_header(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    comp_idx = instr.get("component_index", 0)
    if value is None:
        raise SlideEditError("set_delta_header requires 'value'")
    deltas = [c for c in spec.components if isinstance(c, DeltaColumnComponent)]
    if comp_idx >= len(deltas):
        raise SlideEditError(f"component_index={comp_idx} out of range")
    deltas[comp_idx].header = str(value)


def _reorder_categories(spec: SlideSpec, instr: dict) -> None:
    """Reorder categories (and corresponding series values / deltas / labels)."""
    new_order = instr.get("order")
    if not isinstance(new_order, list):
        raise SlideEditError("reorder_categories requires 'order' (list of indices or names)")

    # Find the chart
    charts = [c for c in spec.components if isinstance(c, ChartComponent)]
    if not charts:
        raise SlideEditError("No chart component to reorder")

    chart = charts[0]
    old_cats = chart.data.categories

    # Resolve order to indices
    if all(isinstance(x, int) for x in new_order):
        indices = new_order
    elif all(isinstance(x, str) for x in new_order):
        cat_to_idx = {c: i for i, c in enumerate(old_cats)}
        indices = []
        for name in new_order:
            if name not in cat_to_idx:
                raise SlideEditError(f"Category {name!r} not found in current categories")
            indices.append(cat_to_idx[name])
    else:
        raise SlideEditError("order must be all ints (indices) or all strings (names)")

    if len(indices) != len(old_cats):
        raise SlideEditError(
            f"order has {len(indices)} items but chart has {len(old_cats)} categories"
        )

    # Reorder categories
    chart.data.categories = [old_cats[i] for i in indices]

    # Reorder series values
    for series in chart.data.series:
        series.values = [series.values[i] for i in indices]

    # Reorder label tables
    for comp in spec.components:
        if isinstance(comp, LabelTableComponent):
            if len(comp.labels) == len(old_cats):
                comp.labels = [comp.labels[i] for i in indices]

    # Reorder delta columns
    for comp in spec.components:
        if isinstance(comp, DeltaColumnComponent):
            if len(comp.values) == len(old_cats):
                comp.values = [comp.values[i] for i in indices]


def _set_slide_id(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    if value is None:
        raise SlideEditError("set_slide_id requires 'value'")
    spec.slide_id = str(value)


def _set_slide_index(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    if value is None or not isinstance(value, int) or value < 0:
        raise SlideEditError("set_slide_index requires non-negative int 'value'")
    spec.slide_index = value


def _set_speaker_notes(spec: SlideSpec, instr: dict) -> None:
    value = instr.get("value")
    if value is None:
        raise SlideEditError("set_speaker_notes requires 'value'")
    if spec.metadata is None:
        from slidegen.slide_spec.schema import SlideMetadata
        spec.metadata = SlideMetadata()
    spec.metadata.speaker_notes = str(value)


# ── Action dispatch table ────────────────────────────────────────────────────

_ACTION_HANDLERS = {
    "set_headline_text": _set_headline_text,
    "set_headline_style": _set_headline_style,
    "set_headline_color": _set_headline_color,
    "set_subheadline_text": _set_subheadline_text,
    "set_footer_text": _set_footer_text,
    "set_brand": _set_brand,
    "set_section": _set_section,
    "set_layout": _set_layout,
    "set_chart_pattern": _set_chart_pattern,
    "change_series_color": _change_series_color,
    "change_series_name": _change_series_name,
    "set_data_labels_format": _set_data_labels_format,
    "set_data_labels_position": _set_data_labels_position,
    "set_data_labels_show": _set_data_labels_show,
    "set_legend_show": _set_legend_show,
    "set_legend_position": _set_legend_position,
    "set_gridlines": _set_gridlines,
    "set_value_axis_show": _set_value_axis_show,
    "set_value_axis_format": _set_value_axis_format,
    "set_delta_format": _set_delta_format,
    "set_delta_colors": _set_delta_colors,
    "set_delta_header": _set_delta_header,
    "reorder_categories": _reorder_categories,
    "set_slide_id": _set_slide_id,
    "set_slide_index": _set_slide_index,
    "set_speaker_notes": _set_speaker_notes,
}
