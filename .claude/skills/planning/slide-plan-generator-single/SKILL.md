---
name: slide-plan-generator-single
effort: low
paths: []
description: "Produce one SlideSpec from one ask (e.g. a client question, a segment analysis request). Takes (ask, data, brand) and emits a single validated SlideSpec ready for slide-creator. Trigger when: Workflow 3 (client followup) fires, Workflow 7/8 (segment analysis as callout/new slide), or any 'Add a slide about X' ad-hoc request. Narrow scope — one ask, one spec."
---

# slide-plan-generator-single

One ask → one `SlideSpec`.

## Cardinal Rules

1. **One ask in, one spec out.** This is not a deck planner. For multi-slide plans, compose this skill N times upstream.
2. **Use the deterministic primitives.** `viz-selector` for `chart_pattern`, `layout-selector` for `layout`. Don't freelance chart/layout choices.
3. **Validate on exit.** `validate_spec(spec, strict=True)`.
4. **Fill metadata.** Set `metadata.created_by="slide-plan-generator-single"`, record the originating `ask`, and carry through any `hypothesis_refs` the caller provides.

## Inputs / Outputs

**Input:**
- `ask: dict` — `{question, metric_tag, q_type, segment, desired_layout_hint}` (all optional except `question`)
- `data`: the data to visualize, in standard `{categories, series}` format
- `brand: str` — BRAND key
- `section: str` — where this slide belongs
- Optional: `headline` (skip headline-writer if pre-written)

**Output:** One valid `SlideSpec`.

## Pipeline

1. `viz-selector.select_chart_pattern(metric_tag, q_type)` → `chart_pattern`
2. Determine component list from the chart pattern (e.g. bar_clustered_horizontal typically gets label_table + delta_column siblings)
3. `layout-selector.select_layout(components)` → `layout`
4. Position each component via layout preset
5. Build `ChartData` from data
6. `headline-writer` if no headline passed in
7. Fill `data_lineage` from the ask's data source
8. `validate_spec(strict=True)` → return

## Python API

```python
from slidegen.slide_plan_single import generate_single_slide, Ask

ask = Ask(
    question="How does efficacy recall differ between Academic and Community prescribers?",
    metric_tag="message_recall",
    segment="Academic vs Community",
    arc="ACT NOW: Segmentation drift",
)
data = {
    "categories": ["Efficacy", "Safety", "Convenience"],
    "series": [
        {"name": "Academic",  "values": [0.55, 0.42, 0.35]},
        {"name": "Community", "values": [0.38, 0.35, 0.30]},
    ],
}
spec = generate_single_slide(
    ask=ask,
    data=data,
    brand="RYBREVANT",
    section="Segment Analysis",
    slide_id="zrx_segment_1",
)
# → validated SlideSpec ready for slide-creator
```

## Status

Landed. viz-selector picks chart_pattern, layout-selector picks layout, headline-writer generates a headline when `ask.suggested_headline` is not provided. Component positions fall back to canonical 1chart_1table coordinates unless overridden via `positions` kwarg.

## References

- PRD §3 Workflow 3/7/8, §5 composition map
- `.claude/skills/planning/viz-selector/SKILL.md`
- `.claude/skills/planning/layout-selector/SKILL.md`
- `.claude/skills/creation/headline-writer/SKILL.md`
