---
name: slide-plan-generator-refresh
effort: medium
paths: []
description: "Generate a diff plan for Workflow 2 (wave refresh). Takes deck-reader output (list[SlideSpec] from wave N-1) + new wave data + any client feedback, emits a structured diff: (a) specs to update (same structure, new data), (b) slides to add (new data warrants a slide not in prior wave), (c) slides to delete (deprecated content). Trigger when: user says 'Refresh wave N of the X PET deck.' Consumer is the wave-refresh-workflow, which then calls slide-updater / slide-creator / deck-assembler per the diff."
---

# slide-plan-generator-refresh

Prior-wave specs + new data → refresh diff plan.

## Cardinal Rules

1. **Preserve narrative continuity by default.** Unless new data clearly requires a structural change, slides stay as-is — only data refreshes.
2. **Every emitted spec is valid.** `validate_spec(strict=True)` on each item in the diff.
3. **Explain every addition and deletion.** The diff plan carries a `rationale` field per change. No silent structural edits.
4. **Deltas are comparative.** Use current-wave vs prior-wave data to justify updates (e.g. "new trend emerged: ...").

## Inputs / Outputs

**Input:**
- `prior_specs: list[SlideSpec]` — from `deck-reader` on the prior wave's PPTX
- `new_data`: new wave data (from `synapse-read` or Excel extraction)
- `feedback` (optional): client feedback on the prior wave
- `narrative_threads` (optional): fresh arcs from `narrative-threads-builder`

**Output:** A structured diff:
```python
{
    "update": [SlideSpec, ...],      # same spec, new data merged (slide-updater will apply)
    "add":    [SlideSpec, ...],      # net-new slides
    "delete": [slide_id, ...],       # slide_ids to drop
    "rationale": {slide_id: "why", ...}
}
```

## Decision Rules

| Situation | Diff action |
|---|---|
| Data for prior slide exists in new wave, structure unchanged | `update` |
| Prior slide's metric no longer tracked | `delete` with rationale |
| New metric appears in new wave with narrative significance | `add` new spec (viz-selector picks pattern) |
| Prior slide's headline no longer matches (data contradicts) | `update` + flag for headline-writer refresh |
| Client feedback requests removal | `delete` with rationale |

## Python API

```python
from slidegen.slide_plan_refresh import generate_refresh_plan, RefreshPlan

plan = generate_refresh_plan(
    prior_specs=[...],                 # from deck-reader on prior wave
    new_data={lineage_key: data_dict, ...},
    feedback={"remove_slides": [...]}, # optional
    narrative_threads=optional_dict,
    new_period_prior="Q4 '25",
    new_period_current="Q1 '26",
)
# plan.update:    list[SlideSpec] — data-refreshed specs ready for slide-updater
# plan.add:       list[SlideSpec] — net-new slides (currently always empty; future work)
# plan.delete:    list[slide_id]  — slides to drop
# plan.rationale: dict[slide_id → reason string]
```

The `lineage_key` matches specs by their `DataLineage`:
- Synapse-canonical keys (preferred): `reporting_plan_id + analysis_ids + segment_ids + deliverables`
- Excel-path fallback: `data_source + extraction_method + question_codes`

## Status

Core planner landed. Narrative-aware `add` pass (propose net-new slides based on narrative_threads arcs) is a future enhancement — currently new lineages are surfaced in `rationale` only and the workflow layer decides whether to build them.

## References

- PRD §3 Workflow 2, §5 composition map
- `.claude/skills/context-data/deck-reader/SKILL.md`
- `.claude/skills/creation/slide-updater/SKILL.md`
- `.claude/skills/planning/viz-selector/SKILL.md`
