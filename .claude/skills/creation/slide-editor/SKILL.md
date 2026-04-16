---
name: slide-editor
effort: medium
paths: ["slidegen/slide_editor.py"]
description: "Apply a structural edit to an existing SlideSpec. Takes (spec, instruction) and returns a modified spec — change headline, change colors, change chart pattern, change data labels, reorder categories, etc. Whitelisted actions only; unsupported instructions rejected. Trigger when: workflow 5 (single slide regen) has a targeted change request, workflow 3 (client followup) modifies an existing slide, or any ad-hoc 'Edit Slide N — ...' instruction. Use slide-updater for data-only refresh."
---

# slide-editor

Takes an existing `SlideSpec` + an edit instruction, returns a modified `SlideSpec`. Structural changes only — for data-only refresh use `slide-updater`.

## Cardinal Rules

1. **Whitelisted actions only.** Every supported action is enumerated below. Unknown action → raise `SlideEditError`. Never dispatch free-form instructions.
2. **One action per call.** Caller composes multi-step edits by calling `edit_slide` N times. Keeps each call auditable.
3. **Validate on exit.** `validate_spec(output, strict=True)` before returning. If an edit would produce an invalid spec (e.g. unknown layout, mismatched category counts), reject.
4. **Fresh object.** Never mutate the input in place.
5. **Preserve data lineage.** Edits change structure, not data provenance. `data_lineage` is carried through unchanged unless the action explicitly modifies it.

## Inputs / Outputs

```python
from slidegen.slide_editor import edit_slide

updated_spec = edit_slide(
    spec: SlideSpec,
    instruction: dict,   # {"action": "<whitelisted_action>", ...action-specific fields}
) -> SlideSpec
```

## Supported Actions

| Action | Required fields | Effect |
|---|---|---|
| `set_headline_text` | `value: str` | Replace `spec.headline.text` |
| `set_headline_style` | `value: str` | Change headline style token (`default`, `red_accent`, `neutral`) |
| `set_headline_color` | `value: str` (color token) | Override headline color |
| `set_subheadline_text` | `value: str` | Replace `spec.subheadline.text` (creates subheadline if None) |
| `set_footer_text` | `value: str` | Replace `spec.footer.text` (creates footer if None) |
| `set_brand` | `value: str` (BRAND key) | Swap brand — useful for template migration |
| `set_section` | `value: str` | Change the section label |
| `set_layout` | `value: str` (LAYOUTS key) | Change the layout preset |
| `set_chart_pattern` | `component_index: int`, `value: str` | Change a chart's `chart_pattern` |
| `change_series_color` | `component_index: int`, `series_index: int`, `value: str` (color token) | Change one series' color |
| `change_series_name` | `component_index: int`, `series_index: int`, `value: str` | Rename one series |
| `set_data_labels_format` | `component_index: int`, `value: str` (e.g. `"0%"`, `"0"`) | Change label number format |
| `set_data_labels_position` | `component_index: int`, `value: str` (`inEnd`, `outEnd`, `ctr`, `above`) | Reposition data labels |
| `set_data_labels_show` | `component_index: int`, `value: bool` | Toggle data labels |
| `set_legend_show` | `component_index: int`, `value: bool` | Toggle legend |
| `set_legend_position` | `component_index: int`, `value: str` (`top`, `bottom`, `left`, `right`) | Reposition legend |
| `set_gridlines` | `component_index: int`, `value: bool` | Toggle major gridlines |
| `set_value_axis_show` | `component_index: int`, `value: bool` | Toggle value axis visibility |
| `set_value_axis_format` | `component_index: int`, `value: str` | Change value axis number format |
| `set_delta_format` | `component_index: int`, `value: str` (`delta_pp`, `delta_pct`, `delta_abs`) | Change delta format |
| `set_delta_colors` | `component_index: int`, `positive: str`, `negative: str` (color tokens) | Override delta colors |
| `set_delta_header` | `component_index: int`, `value: str` | Change delta column header |
| `reorder_categories` | `order: list[str]` | Reorder chart categories (and mirror in label_table + delta_column) |
| `set_slide_id` | `value: str` | Rename the slide |

## Decision Rules

| Situation | Response |
|---|---|
| `action` not in whitelist | Raise `SlideEditError` |
| `component_index` out of range | Raise `SlideEditError` |
| Color token has invalid syntax | Raise `SlideEditError` (validator catches it on exit too) |
| `reorder_categories.order` doesn't match the existing category set | Raise `SlideEditError` |
| Edit leaves spec invalid per `validate_spec` | Raise `SlideEditError` with validator's error list |

## References

- `slidegen/slide_editor.py` — implementation (~28 action handlers)
- `.claude/skills/creation/slide-updater/SKILL.md` — for data-only refresh
- PRD §6.7 (spec contract)
