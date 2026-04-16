---
name: refresh-deck-workflow
effort: high
paths: []
description: "Triggered by: 'Refresh wave N of the <brand> PET deck' / 'Pull in Q1 '26 data and rebuild the deck' / 'Refresh this deck'. Produces the next period's deck from the prior period's deck + new data. For PET/tracker decks this is the wave refresh flow; for other decks it's a multi-slide data refresh preserving narrative continuity. Primary Q3 demo."
---

# refresh-deck-workflow

Prior-wave deck + new wave data → next wave deck. Primary Q3 demo workflow.

## Trigger phrases

- "Refresh wave 4 of the Rybrevant PET deck"
- "Pull in Q1 '26 data and rebuild the deck"
- "Run the wave refresh on projects/jnj_rybrevant"
- "Create the Q1 '26 PET report using Q4 '25 as the base"

## Cardinal Rules

1. **Preserve narrative continuity, not stale text.** Slide *structure* stays identical unless new data genuinely requires a change. But **headlines MUST be regenerated** to reflect new data — a headline like "dipped 3pp QoQ" from the prior wave becomes a delivery risk when new data says "up 2pp". `slide-updater` handles this via `headline-writer`. Subheadlines with period labels ("Q3 '25 vs Q4 '25") are rewritten to the new period pair.
2. **Refresh the narrative backbone too.** `narrative_threads.md` is regenerated from new data before headlines are rewritten — arcs may shift wave over wave, and headlines inherit from the fresh narrative.
3. **Every refresh is auditable.** Stamp `data_lineage.last_data_pull` on every updated slide. Preserve tag-based lineage in metadata even when the tag itself was rejected (audit trail).
4. **Trust tags that pass health checks, route others through inference.** `deck-reader` applies a 5-step health check on every tagged shape. Only Tier 1 (all checks pass) tags are used for refresh directly. Unhealthy tags → Tier 2 inference. No per-shape user gate on tag failures — gating per-shape in a 40-slide deck is unusable. Blind-trusting a bad tag is how galen-powerpoint's own refresh can destroy a chart's layout — our refresh must not repeat that.
5. **User review gate on low-confidence Tier 2 inferences only.** If `deck-reader` Tier 2 can't confidently reconstruct a slide's lineage (low match score), surface for confirmation before refreshing. This keeps human-in-the-loop narrow — only the genuinely ambiguous cases.
6. **Deletes are explicit.** Slides removed from the diff plan must have a rationale — never silently drop a slide.

## Inputs

- `config_path`: project config.yaml (identifies brand, wave, source paths)
- `prior_wave_deck_path`: path to the previous wave PPTX
- Optional `feedback`: client feedback from prior wave that should inform the refresh
- Synapse credentials (from env or `synapse-auth`)

## Outputs

- New wave PPTX at `projects/{name}/output/{new_wave}/deck.pptx`
- Updated `shape_registry.json` with refreshed lineage
- Diff summary: `{updated: [...], added: [...], deleted: [...]}` printed to terminal

## Orchestration

```
1. deck-reader(prior_wave_deck_path, config_path)
     → list[SlideSpec] for every slide (Tier 1 where tagged, Tier 2 otherwise)
     → summary dict with tagged/untagged counts

2. synapse-read (or excel extraction) (config_path, new wave identifier)
     → dict of {extraction_id: data} for the new wave

3. prior-wave-context-builder (prior_wave_deck_path)
     → updates context/{wave}/prior_wave_context.md (for narrative continuity)

4. slide-plan-generator-refresh (prior_specs, new_data, feedback?, narrative_threads?)
     → diff plan: {update: [SlideSpec], add: [SlideSpec], delete: [slide_id]}
     → each plan entry has a rationale string

5. (optional) trend-analyzer for deltas that warrant callouts

6. Gate: show diff plan to user for review
     → user can approve / reject / edit individual entries

7. for spec in diff.update:
     slide-updater(
        spec,
        new_data[spec.data_lineage],
        period_labels={"prior": new_period_prior, "current": new_period_current}
     )
     → updated_spec with:
        • fresh values + recomputed deltas
        • HEADLINE regenerated via headline-writer (from new data + updated narrative_threads)
        • SUBHEADLINE period labels rewritten ("Q3 '25 vs Q4 '25" → new period pair)
        • audit stamps
     spec-validator(updated_spec, strict=True)

   for spec in diff.add:
     viz-selector + layout-selector + headline-writer
     slide-creator(spec)

   for slide_id in diff.delete:
     (skip in final assembly; log rationale)

8. deck-assembler(specs, output_path=new_wave_deck_path, template_path=?)
     → PPTX + shape_registry.json + backup of any prior output
```

## Decision Rules

| Situation | Response |
|---|---|
| Prior deck has no Connector tags (all untagged) | All slides go through Tier 2 inference; require user confirmation on low-confidence specs |
| New wave data is missing for a tagged slide | Flag in diff plan; skip that slide's update, keep prior data with warning |
| A prior slide's Connector lineage points to a reporting plan that no longer exists | Surface error; user decides whether to delete the slide or swap the lineage |
| Segment IDs in lineage differ from segments available in new wave | Halt; user resolves |
| Diff plan is empty (nothing changed) | Report "no refresh needed" and exit without writing a new deck |

## References

- PRD §3 Workflow 2, §5 composition map, §6.8 (dual-mode lineage)
- `.claude/skills/context-data/deck-reader/SKILL.md`
- `.claude/skills/planning/slide-plan-generator-refresh/SKILL.md`
- `.claude/skills/creation/slide-updater/SKILL.md`
- `.claude/skills/creation/slide-creator/SKILL.md`
- `.claude/skills/creation/deck-assembler/SKILL.md`
- Galen-PowerPoint Connector: `Docs/Export Import Tags - PRD.md` (for Tier 1 reference)
