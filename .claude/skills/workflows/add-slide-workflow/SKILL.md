---
name: add-slide-workflow
effort: low
paths: []
description: "Triggered by: 'Client asked about X on slide N — create an answer slide' / 'Add a segment comparison slide' / 'Insert a new slide after slide 40'. Inserts ONE new slide into an existing deck. Handles ad-hoc client questions, segment comparisons, and any 'add a slide about X' request. Consolidates client-followup and the Workflow-8 new-slide-for-segment path of segment-analysis."
---

# add-slide-workflow

Insert ONE new slide into an existing deck. Consolidates client-followup + segment-analysis-new-slide into one workflow.

## Trigger phrases

**Ad-hoc client question:**
- "Client asked about Academic vs Community on slide 35 — create an answer slide"
- "Build a followup slide for: 'how does efficacy recall differ by region?'"
- "Add a slide to the deck answering this question"

**Segment comparison (full slide, not annotation):**
- "Add a new slide comparing segments for this brand"
- "Create a side-by-side comparison of Academic vs Community for message recall"
- "Insert a segment-comparison slide after slide 40"

**General-purpose slide insertion:**
- "Insert a slide about X after slide 40"
- "Add a slide for <metric> to the deck"

Detection: if the ask involves segment comparison → route analysis through `segment-comparator` + `stat-sig-annotator`. Otherwise → standard `slide-plan-generator-single`. Output form is the same: ONE new slide inserted.

## Cardinal Rules

1. **One ask, one slide.** For multi-slide mini-decks, escalate to `create-deck-workflow` for the sub-topic.
2. **Referenced slide is context, not content.** If user names a reference slide, `deck-reader` pulls its brand/section/style for context but the new slide has its own lineage + headline.
3. **Data may need a new extraction.** If the ask requires a cut not in existing deck, invoke `synapse-read` with new params — don't derive it from the referenced slide.
4. **User specifies position.** Default insertion is after the referenced slide. User can override with `insert_after_index`.
5. **Statistical rigor for segment cases.** If the ask involves segment comparison, significance testing (`stat-sig-annotator`) is mandatory; n<30 per segment triggers a "low base" annotation.

## Inputs

- `deck_path`: the existing deck
- `question` (or `segment_question` for segment case): the text of the ask
- `reference_slide_index` (optional): slide the user referenced for context
- `insert_after_index` (optional): insertion position, defaults to reference_slide_index
- `data_cuts` (optional): pre-specified extraction params if new data is needed

## Outputs

- Updated PPTX with one inserted slide
- Updated `shape_registry.json`
- If segment analysis: also saves analysis trace via `analysis-trace-store`

## Orchestration

```
1. if reference_slide_index is given:
     deck-reader(deck_path, slide_index=reference_slide_index)
       → reference_spec for context (brand, section, style)

2. BRANCH on ask type:

   ── standard ad-hoc question ──
   2a. if data_cuts needed:
         synapse-read(new extraction params) → data

   ── segment comparison ──
   2b. synapse-read(data_lineage from reference + segment split) → segmented data
   2c. segment-comparator(data) → comparison structure
   2d. stat-sig-annotator(comparison) → with significance markers, low-base flags

3. slide-plan-generator-single(
       ask=question (+ segment detail if applicable),
       data=data or comparison_output,
       brand=reference_spec.brand,
       section=reference_spec.section or "Client Follow-up" or "Segment Analysis"
   )
     → SlideSpec (chart_pattern via viz-selector, layout via layout-selector,
        headline via headline-writer)

4. spec-validator(new_spec, strict=True)

5. slide-creator(new_spec)

6. deck-assembler(existing_deck=deck_path, insert_at=insert_after_index+1, with=[new_spec])

7. (if segment case) analysis-trace-store(comparison + stat-sig results)
```

## Decision Rules

| Situation | Response |
|---|---|
| No `reference_slide_index` given | Skip step 1; viz-selector may fall back to HITL for chart type |
| Question is vague / non-data | viz-selector returns "UNRESOLVED" → surface to user |
| Segment definitions don't exist in Synapse | Halt; instruct user to set up segments first via `synapse-cli` |
| n<30 in any segment (segment case) | Proceed but annotate "low base"; flag in footer |
| No significant differences (segment case) | Still produce slide; headline says "no meaningful difference detected" |
| Insert position is at end of deck | Allowed — `insert_at` = `len(deck)` |
| User wants multiple variants | Run this workflow once per variant, not in the same call |

## References

- PRD §3 Workflow 4
- `.claude/skills/context-data/deck-reader/SKILL.md`
- `.claude/skills/planning/slide-plan-generator-single/SKILL.md`
- `.claude/skills/planning/viz-selector/SKILL.md`
- `.claude/skills/analysis/segment-comparator/` (planned)
- `.claude/skills/analysis/stat-sig-annotator/` (planned)
- `.claude/skills/creation/slide-creator/SKILL.md`
- `.claude/skills/creation/deck-assembler/SKILL.md`
- Sibling: `.claude/skills/workflows/annotate-slide-workflow/SKILL.md` (for callouts on existing slides, not new slides)
