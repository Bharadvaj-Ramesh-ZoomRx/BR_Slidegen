---
name: executive-summary-workflow
effort: medium
paths: []
description: "Triggered by: 'Generate an executive summary answering these KBQs' / 'Build an ES slide for slides 5-40'. Produces 1-3 ES slides that answer specified KBQs with findings citing back to supporting slide_ids. Every ES claim is sourced from a specific slide; unsourced claims are rejected."
---

# executive-summary-workflow

Full deck + KBQs → 1-3 ES slides with citations. Every claim is traceable to a source slide.

## Trigger phrases

- "Generate an executive summary answering these 3 KBQs"
- "Build an ES slide for the current deck"
- "Write the exec summary for slides 5-40"
- "Create a top-of-funnel summary slide with citations"

## Cardinal Rules

1. **Every claim cites a `slide_id`.** If no slide supports a claim, the claim is dropped. This is non-negotiable — clients audit ES slides first.
2. **KBQs drive ES structure, not the deck structure.** One ES slide (or cluster) per KBQ cluster. If 3 KBQs cluster into 1 arc, that's 1 slide.
3. **≤3 ES slides max.** Longer ES dilutes impact. If 3 KBQs deserve their own ES slide, produce 3 slides — don't pack a mega-slide.
4. **Narrative threads fuel the ES.** Use existing `narrative_threads.md` if present; it provides the arc → claims → citations mapping. Only re-derive if missing.

## Inputs

- `deck_path`: the full deck being summarized
- `kbqs`: list of key business questions (from `kbqs.md` or user input)
- Optional `narrative_threads_path`: pre-built arcs from `sfea-insight-writer`
- `insert_position`: where ES slides go (typically slide 2, after cover)

## Outputs

- Updated PPTX with 1-3 new ES slides inserted
- Each ES slide's `metadata.kbq_refs` + `metadata.citations` populated
- Updated `shape_registry.json`

## Orchestration

```
1. deck-reader(deck_path, all_slides=True)
     → list[SlideSpec] for every slide (used as citation targets)

2. if narrative_threads_path not provided:
     (Optional) sfea-insight-writer on existing hypothesis_bank
     → narrative_threads.md (or use summarization over deck_specs)

3. slide-plan-generator-exec-summary(
       deck_specs=all slides from step 1,
       kbqs=user-provided or from kbqs.md,
       narrative_threads=from step 2,
       brand=first spec.brand,
       section="Executive Summary"
   )
     → 1-3 ES SlideSpec entries, each with:
        • components: [TextboxComponent, ...] for each bullet
        • metadata.kbq_refs: list[str]
        • metadata.citations: list[slide_id]

4. Gate: show ES specs to user for review
     → user can approve / edit individual bullets / cite different slides

5. executive-summary-writer finalizes the bullet text style
     (Not a separate rendering step — it's called inside step 3)

6. for es_spec in plan:
     spec-validator(es_spec, strict=True)
     slide-creator(es_spec)

7. deck-assembler(existing_deck=deck_path, insert_at=insert_position, with=es_specs)
```

## Decision Rules

| Situation | Response |
|---|---|
| A KBQ has NO matching narrative thread and NO supporting slides | Flag; produce an ES slide that explicitly says "open question — not directly addressed in this deck" |
| `slide-plan-generator-exec-summary` can't find ≥3 strong findings for a KBQ | Reduce to 1-2 bullets for that KBQ rather than padding |
| Deck has 5+ KBQs | Cluster KBQs by topic; don't produce 5 ES slides |
| Existing deck already has an ES slide at `insert_position` | Ask user: replace, or insert adjacent? |
| A bullet's citation `slide_id` doesn't exist in deck | Drop the bullet; log warning |

## References

- PRD §3 Workflow 9, §5 composition map
- `.claude/skills/context-data/deck-reader/SKILL.md`
- `.claude/skills/planning/slide-plan-generator-exec-summary/SKILL.md`
- `.claude/skills/creation/executive-summary-writer/SKILL.md` (planned; not yet implemented)
- `.claude/skills/analysis/sfea-insight-writer/SKILL.md` — source of narrative threads
- `.claude/skills/creation/slide-creator/SKILL.md`
