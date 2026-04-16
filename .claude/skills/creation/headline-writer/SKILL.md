---
name: headline-writer
effort: medium
paths: ["slidegen/headline_writer.py"]
description: "Generate one slide headline from data + narrative thread + brand + section. Template-based baseline with deterministic patterns (delta-driven, rank-shift, magnitude-threshold, flat, mixed) — can be LLM-augmented for prose variety. Produces client-ready headlines matching real PET deck style (3,569 real headlines in experiments/deck_analysis/outputs/deep_headlines.json as reference). Invoked by slide-updater on every data refresh (headlines must not go stale with new data) and by slide-plan-generator-single for new-slide generation."
---

# headline-writer

Data + narrative context → one client-ready headline string. LLM skill, no Python file backs it.

## Cardinal Rules

1. **Lead with direction.** Every headline names the change: up, down, flat, above, below. Never ambiguous.
2. **Name the driver or segment when possible.** "Efficacy recall dipped 3pp QoQ, driven by Specialists" > "Efficacy recall dropped."
3. **No hedging.** Ban phrases: "may suggest," "appears to indicate," "could potentially," "seems to show." The data says what the data says.
4. **≤120 characters.** Longer headlines wrap awkwardly in the observed layout presets. Count chars before finalizing.
5. **Never invent numbers.** Every delta, count, or rank mentioned must appear in the source data. If you can't justify it from the data, drop it.
6. **Match real PET style.** Reference `experiments/deck_analysis/outputs/deep_headlines.json` for the style pattern: terse, noun-heavy, delta-forward.

## Inputs / Outputs

**Input:**
- `data`: the slide's chart data (categories + series)
- `delta_column.values`: already-computed QoQ deltas
- `narrative_thread`: relevant arc or finding this slide supports
- `brand`: brand name + therapy area for voice
- `section`: the section this slide sits in (for context)
- Optional: `hypothesis_refs`, `prior_headline` (for "rewrite this" requests)

**Output:** A single string fitting `spec.headline.text`. Does NOT return the full SlideSpec — caller assigns the string.

## Style Patterns (from 3,569 real headlines)

| Pattern | Example |
|---|---|
| Delta + driver | "Rybrevant efficacy recall dipped 3pp QoQ, driven by Efficacy-in-1L message" |
| Rank shift | "Safety climbs to #2 recalled message, overtaking Convenience" |
| Magnitude threshold | "Half of NSCLC specialists now recall LITE + EP2 (+8pp vs Q4)" |
| Segment divergence | "Academic HCPs lead Community on Rybrevant awareness by 12pp" |
| Flat with context | "Recall holds at 45% — no wave-on-wave shift after Q4 campaign pulse" |
| Comparative | "Rybrevant vs Tagrisso: same reach, half the unaided recall" |

## Decision Rules

| Situation | Response |
|---|---|
| Data shows no significant change (delta ≤ margin of error) | Use "flat" / "holds" / "steady" pattern, not fabricated movement |
| Small sample (n<30 in any segment cited) | Omit the segment or flag in footer, not headline |
| Narrative thread has a specific arc role (CONVERGENCE, TENSION, etc.) | Tilt headline toward the arc pattern the thread specifies |
| Rewrite request with `prior_headline` | Preserve factual spine; rework wording only |
| Template can't resolve a pattern (no chart, no narrative hint) | Raise `HeadlineGenerationError`; caller falls back to section name or section question |

## Python API

```python
from slidegen.headline_writer import generate_headline, rewrite_period_labels

# Generate a fresh headline from spec's chart data
text = generate_headline(
    spec,                           # SlideSpec with at least one ChartComponent
    narrative_hint="ACT NOW: Efficacy drift",  # optional arc from narrative_threads
    prior_headline=None,            # optional — for rewrite requests
    max_chars=120,
)
# → "Efficacy message recall dipped 3pp QoQ, driven by Specialists"

# Rewrite period references in a subheadline (used by slide-updater when
# wave changes — e.g. "Q3 '25 vs Q4 '25" → "Q4 '25 vs Q1 '26")
new_sub = rewrite_period_labels(
    subheadline_text="Q3 '25 vs Q4 '25 among NSCLC prescribers",
    old_period_prior="Q3 '25", old_period_current="Q4 '25",
    new_period_prior="Q4 '25", new_period_current="Q1 '26",
)
```

## Template patterns (deterministic fallback)

| Shape of the data | Template used | Example |
|---|---|---|
| All categories flat (|Δ|<2pp) | `_template_flat` | "Rybrevant message recall holds steady QoQ — no meaningful wave-on-wave shift" |
| Mixed direction | `_template_mixed` | "Rybrevant efficacy: Safety +4pp, Access -3pp" |
| Clear single direction, same top + driver | `_template_rank_top` | "Efficacy leads Message Recall for Rybrevant at 42% (+3pp QoQ)" |
| Clear single direction, different top vs driver | `_template_delta_driver` | "Rybrevant message recall rose +4pp QoQ, driven by Safety" |

The 120-char limit is enforced as a hard cap with ellipsis truncation.

## References

- `slidegen/headline_writer.py` — implementation (template-based baseline; LLM augmentation is a wrapper layer)
- `experiments/deck_analysis/outputs/deep_headlines.json` — 3,569 real headlines as style corpus
- `.claude/skills/analysis/sfea-insight-writer/SKILL.md` (PET) / `atu-insight-writer/SKILL.md` (ATU) — existing skill that writes narrative arcs + headlines inline (Phase 1 of the 2-phase flow)
- `.claude/skills/creation/slide-updater/SKILL.md` — primary caller; invokes on every data refresh
- PRD §6.7 (spec contract), §6.9 (headline-refresh policy)
