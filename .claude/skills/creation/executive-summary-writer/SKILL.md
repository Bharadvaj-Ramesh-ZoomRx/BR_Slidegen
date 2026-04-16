---
name: executive-summary-writer
effort: medium
paths: ["slidegen/slide_plan_exec_summary.py"]
description: "Produce 1-3 ES slide specs that answer a set of KBQs. Each ES bullet has citations back to the source slide_id that supports the claim. Unsourced claims are rejected. Use via slide-plan-generator-exec-summary (the underlying planner) with optional LLM augmentation for prose style."
---

# executive-summary-writer

ES slide specs with citations. Every bullet traces to a specific slide.

## Cardinal Rules

1. **Every claim cites a `slide_id`.** No ES bullet without at least one `slide_id` in its supporting citations. Clients audit ES slides first; unsourced claims are a hard failure.
2. **≤3 ES slides max.** Longer ES dilutes impact. If more than 3 KBQ clusters matter, negotiate which to drop — don't exceed the cap.
3. **Citations come from the deck itself.** Every cited `slide_id` must exist in the input `deck_specs`.
4. **Text-only layout.** ES slides are TextboxComponents on an `observed_full_width_table` layout. No charts, no tables — just tight prose bullets.
5. **KBQ → arc via narrative_threads when available.** If `narrative_threads.md` is present, prefer arc-based citation clustering over keyword matching (higher-quality pairings).

## API

```python
from slidegen.slide_plan_exec_summary import generate_exec_summary

es_slides = generate_exec_summary(
    deck_specs=[...],                # full deck being summarized (from deck-reader)
    kbqs=["How is recall changing?", "What drives intent?", ...],
    narrative_threads=optional_dict, # from sfea-insight-writer
    brand="RYBREVANT",
    section="Executive Summary",
    insert_position=1,               # where ES slides go (typically slide 2)
    max_es_slides=3,
    max_kbqs_per_slide=3,
)
# → list[SlideSpec], each validated, citations embedded
```

## Pipeline

1. **Cluster KBQs.** Group KBQs by topic or arc (when narrative_threads available). Each cluster becomes one ES slide. Cap clusters at `max_es_slides`.
2. **Find supporting slides** per KBQ:
   - Priority 1: narrative_threads lookup (`kbq_to_arc` → `arc_to_slides`)
   - Priority 2: keyword match against headlines + sections
   - Priority 3: fallback to first N data-driven slides
3. **Distill each finding** from the top supporting slide's headline. LLM augmentation (summarizing across multiple slides) is a future enhancement — current implementation uses the headline directly.
4. **Build textbox components** per bullet, with citations inline.
5. **Validate** every ES spec via `validate_spec(strict=True)`.

## Decision Rules

| Situation | Response |
|---|---|
| KBQ has no matching narrative thread AND no supporting slides | Produce "open question" bullet: "No slides in this deck directly address {kbq}" |
| `slide-plan-generator-exec-summary` can't find ≥3 strong findings | Reduce to 1-2 bullets per slide rather than padding |
| Deck has 5+ KBQs | Cluster by topic; cap at `max_es_slides`. Overflow KBQs surface in log. |
| A bullet's citation `slide_id` doesn't exist in deck | Drop the bullet; log warning (validator will catch it too). |

## Relationship to headline-writer

`executive-summary-writer` emits ES slides whose headline + bullets are *descriptive* (summarizing findings across many slides). Individual data slides still use `headline-writer` for their own headlines. The two skills don't collide — ES bullets are summaries; individual headlines are wave-specific findings.

## References

- `slidegen/slide_plan_exec_summary.py` — implementation
- `.claude/skills/planning/slide-plan-generator-exec-summary/SKILL.md` — the planner layer (calls this)
- `.claude/skills/analysis/sfea-insight-writer/SKILL.md` — source of narrative threads
- `.claude/skills/context-data/deck-reader/SKILL.md` — source of deck_specs
- `.claude/skills/creation/slide-creator/SKILL.md` — renders the ES slides
- PRD §3 Workflow 8, §5 composition map
