---
name: annotate-slide-workflow
effort: low
paths: []
description: "Triggered by: 'Add a callout to slide N' / 'Highlight the Academic vs Community gap on slide 12 as an annotation' / 'Annotate slide 5 with this quote'. Adds a callout (quote, data annotation, insight tag) INLINE on an existing slide without changing its structure. For adding a whole new slide, use add-slide-workflow instead."
---

# annotate-slide-workflow

Add a callout on an EXISTING slide without changing its structure. Single responsibility — inline annotation only.

## Trigger phrases

**Qualitative quote callout:**
- "Add a quote to slide 10 from the qualitative data"
- "Annotate slide 5 with this quote: '...'"
- "Put a verbatim callout on slide 12"

**Data annotation (insight) callout:**
- "Highlight the Academic vs Community gap on slide 12 as an annotation"
- "Add a callout on slide 20 saying 'Up 8pp among specialists'"
- "Annotate the segment gap on slide 15"

**Ad-hoc insight text:**
- "Add a callout box on slide 5 with: '<text>'"
- "Insert a dashed insight callout after the chart on slide 8"

## Cardinal Rules

1. **Inline only — no new slides.** For adding a full segment-comparison slide, use `add-slide-workflow`. This workflow modifies ONE existing slide.
2. **Real data, real quotes.** Verbatim quotes must be pulled from `qualitative_data.json`. Data annotations must come from a computed finding (e.g. `segment-comparator` output), not invented.
3. **Statistical rigor for segment annotations.** Any "up/down/gap" claim for segments requires `stat-sig-annotator` sign-off. If n<30, annotate with "low base".
4. **Callout position inside existing layout.** Use existing layout's callout rect if defined, else find open whitespace via `layout-selector`-style logic.
5. **Preserve all other components.** The slide's chart, tables, delta columns stay unchanged.

## Inputs

- `deck_path`: deck containing the target slide
- `target_slide_index`: the slide to annotate
- `callout_type`: `"quote"` | `"insight"` | `"freeform"`
- For quote: `quote_theme` (from qualitative themes), optional `quote_pool`
- For insight: the data finding (e.g. segment-comparator output) OR a pre-specified text
- For freeform: the callout text itself
- Optional `position`: explicit bbox or positional hint ("below chart", "right of chart", etc.)

## Outputs

- Updated PPTX with a callout added to the target slide
- Updated `shape_registry.json`
- If segment/data insight: analysis trace saved via `analysis-trace-store`

## Orchestration

```
1. deck-reader(deck_path, slide_index=target_slide_index)
     → existing SlideSpec for the target slide

2. BRANCH on callout_type:

   ── callout_type="quote" ──
   2a. load qualitative_data.json
   2b. callout-writer(
          slide_context=existing_spec,
          theme=quote_theme,
          quote_pool=quote_pool_or_extracted
       )
       → CalloutComponent(theme, text, attribution)

   ── callout_type="insight" (segment/data finding) ──
   2c. if raw data provided: segment-comparator → comparison
       if stat-sig needed: stat-sig-annotator → marked findings
   2d. callout-writer(
          slide_context=existing_spec,
          data_annotation=top_finding,
          style="insight"
       )
       → CalloutComponent(text, style="insight")

   ── callout_type="freeform" ──
   2e. Build CalloutComponent directly from user-provided text

3. Determine callout position:
     if layout has callout_rect: use it
     else: find whitespace right of chart OR below chart per layout preset
     → explicit Position

4. slide-editor(existing_spec, {"action": "add_component", "value": callout_component})
     → updated_spec with callout appended

5. spec-validator(updated_spec, strict=True)

6. slide-creator(updated_spec)
     → re-renders the target slide with the callout

7. deck-assembler(existing_deck=deck_path, replace_at=target_slide_index, with=[updated_spec])

8. (if insight case) analysis-trace-store(comparison + stat-sig results)
```

## Decision Rules

| Situation | Response |
|---|---|
| Target slide has no whitespace for callout | Halt — user must specify explicit position OR use add-slide-workflow for a separate slide |
| Callout_type="quote" but no matching quote in pool | Widen theme search; if still nothing, halt with guidance |
| Quote n<5 | Flag in attribution ("n=4 — low base"); still acceptable |
| Insight has no significant finding (all n.s.) | Callout says "no meaningful difference detected" |
| User provides both raw data AND pre-computed finding | Use pre-computed; ignore raw (trust the user's preprocessing) |

## References

- PRD §3 Workflow 5
- `.claude/skills/context-data/deck-reader/SKILL.md`
- `.claude/skills/creation/callout-writer/SKILL.md`
- `.claude/skills/creation/slide-editor/SKILL.md` (add_component action)
- `.claude/skills/analysis/segment-comparator/` (planned — for insight callouts)
- `.claude/skills/analysis/stat-sig-annotator/` (planned)
- `.claude/skills/creation/slide-creator/SKILL.md`
- Sibling: `.claude/skills/workflows/add-slide-workflow/SKILL.md` (new slides, not callouts)
