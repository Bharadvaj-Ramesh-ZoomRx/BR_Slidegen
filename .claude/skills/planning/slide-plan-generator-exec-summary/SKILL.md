---
name: slide-plan-generator-exec-summary
effort: medium
paths: []
description: "Generate executive summary slide(s) from a full deck + KBQs. Takes list[SlideSpec] (the deck being summarized) + KBQs + narrative threads, emits 1-3 ES slides with findings that cite back to source slide_ids. Every ES claim includes a citation to the slide that supports it. Trigger when: Workflow 9 fires ('Generate an executive summary answering these 3 KBQs')."
---

# slide-plan-generator-exec-summary

Full deck + KBQs → ES slide specs with citations.

## Cardinal Rules

1. **Every claim cites a slide.** No ES bullet without a `slide_id` in its supporting metadata. Clients audit ES slides first; unsourced claims are a hard failure.
2. **KBQ alignment is explicit.** Each ES slide or each bullet maps to one or more KBQs. The mapping lives in `metadata.kbq_refs`.
3. **Three slides max.** Longer ES sections dilute impact. If 3+ KBQs deserve ES treatment, produce one ES slide per KBQ cluster, not a mega-slide.
4. **Consume narrative threads, don't invent arcs.** Use the `narrative-threads-builder` output as the ES backbone. If a KBQ has no matching thread, flag it — don't manufacture one.
5. **Validate on exit.** `validate_spec(strict=True)` on every ES slide.

## Inputs / Outputs

**Input:**
- `deck_specs: list[SlideSpec]` — full deck being summarized
- `kbqs: list[str]` — key business questions to answer
- `narrative_threads: list[dict]` — from `narrative-threads-builder`
- `brand: str`, `section: str` (typically "Executive Summary")

**Output:** `list[SlideSpec]` of 1-3 ES slides, each with:
- `components: [TextboxComponent(...)]` for each insight bullet
- `metadata.kbq_refs: list[str]` — which KBQs this slide answers
- `metadata.citations: list[str]` — slide_ids supporting claims on this slide

## Pipeline

1. Cluster KBQs by topic/arc using narrative_threads
2. For each cluster, pick the 3-5 strongest findings with supporting slide_ids
3. Write each finding as a punchy ES bullet (headline-writer style)
4. Layout via `layout-selector(["textbox", "textbox", ...])` → `observed_full_width_table` (text-only)
5. Validate + return

## Python API

```python
from slidegen.slide_plan_exec_summary import generate_exec_summary

es_slides = generate_exec_summary(
    deck_specs=[...],                # full deck being summarized (from deck-reader)
    kbqs=["How is Rybrevant recall changing?", "Which segments drive intent?"],
    narrative_threads=None,          # optional — from sfea-insight-writer
    brand="RYBREVANT",
    section="Executive Summary",
    insert_position=1,
    max_es_slides=3,
    max_kbqs_per_slide=3,
)
# → list[SlideSpec] with textbox-only ES slides, citations embedded in each bullet
```

## Status

Landed. Citation lookup: narrative_threads arc mapping first, keyword-match fallback, first-N-data-slides as last resort. LLM prose-polish over the citation-gated output is a future enhancement — current implementation uses source slide headlines as-is for bullet text.

## References

- `slidegen/slide_plan_exec_summary.py` — implementation
- PRD §3 Workflow 8
- `.claude/skills/creation/executive-summary-writer/SKILL.md` — the renderer-facing skill that documents the same contract
- `.claude/skills/analysis/sfea-insight-writer/SKILL.md` (PET) / `atu-insight-writer/SKILL.md` (ATU) — source of narrative threads
- `.claude/skills/creation/headline-writer/SKILL.md` — used by generate_exec_summary internally
