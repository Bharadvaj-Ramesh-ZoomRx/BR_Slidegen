---
name: spec-validator
effort: low
paths: ["slidegen/slide_spec/validator.py", "slidegen/slide_spec/schema.py"]
description: "Gatekeeper between intelligent planning skills and deterministic rendering. Use whenever a SlideSpec is about to be handed to slide-creator, slide-updater, or slide-editor. Produces a list of structured errors (or raises in strict mode) so upstream producers can fix the spec before anything renders. Trigger when: a slide-plan-generator skill has emitted a spec, a deck-reader has reconstructed a spec from an existing deck, or a test is asserting spec completeness. Fails loudly — never patches silently."
---

# spec-validator

Enforces the slide-spec contract. A valid spec is the entry condition for every downstream rendering skill.

## What it checks

1. **Top-level required fields**: `slide_id`, `slide_index ≥ 0`, `layout`, `headline.text`, non-empty `components`
2. **Layout existence**: `spec.layout` must be a key in `LAYOUTS{}` (`slidegen/pptx_utils/layout.py`)
3. **Brand existence** (if `spec.brand` is set): must match `BRAND{}` — tries normalized form before failing
4. **Component types**: every `component.type` must be in `SUPPORTED_COMPONENT_TYPES`
5. **Chart patterns**: every `ChartComponent.chart_pattern` must be in `SUPPORTED_CHART_PATTERNS`
6. **Position integrity**: exactly one of (explicit bbox) or (preset reference) — not both, not neither. Preset references must resolve against the layout's rect keys
7. **Position bounds**: explicit bboxes must fit within `13.333 × 7.500` slide bounds, with positive dimensions
8. **Data shape**: chart categories non-empty, every series has a non-empty `values`, series length matches categories (except scatter patterns where values are `[x, y]` pairs)
9. **Color tokens**: every color field uses valid syntax — `#RRGGBB` or `{namespace.path.field}` (the resolver runs in slide-creator, not here — validator only checks syntax)
10. **Cross-component coherence**: if a slide has a chart + label_table + delta_column, row counts match across all three

## Usage

```python
from slidegen.slide_spec import validate_spec, load_spec, SpecValidationError

spec = load_spec("context/Q1 2026/slide_specs/slide_05.json")

# Collect mode — returns list, caller decides what to do
errors = validate_spec(spec)
if errors:
    print(f"Spec has {len(errors)} problems:")
    for e in errors:
        print(f"  - {e}")

# Strict mode — raises SpecValidationError on first failing check
validate_spec(spec, strict=True)
```

## When to call

| Upstream producer | Call validator before handing to |
|---|---|
| `slide-plan-generator-hypothesis` | `slide-creator` |
| `slide-plan-generator-refresh` | `slide-updater` (validates BOTH the source spec and the refreshed spec) |
| `slide-plan-generator-single` | `slide-creator` |
| `slide-plan-generator-exec-summary` | `slide-creator` |
| `deck-reader` (Tier 1 or Tier 2) | `slide-updater` / `slide-editor` |
| `headline-writer` / `callout-writer` | Next producer in the chain (mutation of an existing spec) |

## What it does NOT check

- **Color resolvability** — whether `{brand.primary_current}` actually exists on the BRAND entry. That's slide-creator's job at render time. Validator only checks syntax.
- **Data plausibility** — whether 45% is a reasonable value for "efficacy recall." Validator is structural, not analytical.
- **Headline accuracy** — whether the headline matches the data. That's a task for upstream narrative skills, not the validator.
- **Position visual collision** — two components can overlap without triggering an error. If you need collision detection, it's a separate check (out of scope for v1.0).

## Error format

Errors are human-readable strings with JSON-path-like breadcrumbs:

```
components[0].data.series[1].values has 3 entries; expected 5 to match categories
components[2].position: position.preset='bogus_rect' not in LAYOUTS['observed_1chart_1table']. Available: ['chart_rect', 'table_rect', 'delta_col_rect', 'chart_top', 'table_top']
layout='oberved_1chart_1table' not in LAYOUTS{}. Available: ['clustered_compare', 'dual_chart_with_delta', 'observed_1chart_1table', ...]
```

Breadcrumbs are deterministic so scripts can parse them to auto-fix common issues (typos in layout names, extra fields, etc.).

## Cardinal rules

1. **Never patch silently.** If a spec is invalid, return the errors. Do not fill in defaults. Do not downgrade strictness. The caller has the context to fix the spec; the validator does not.
2. **Every error is actionable.** Every message names the field path and the expected shape. If an error message is too vague for the caller to act on, improve the validator, not the caller's code.
3. **Fail fast.** In `strict=True` mode, raise on the first problem. This is the mode to use in CI and in production pipelines.
4. **Survive future-compat.** Unknown fields warn but don't fail — a newer spec version may add fields this validator hasn't learned about yet.

## References

- `slidegen/slide_spec/schema.py` — the dataclasses being validated
- `slidegen/slide_spec/validator.py` — the validation implementation
- `slidegen/slide_spec/examples/` — worked example specs
- PRD §6.7 — the spec-as-contract principle
