---
name: callout-writer
effort: medium
paths: ["slidegen/callout_writer.py"]
description: "Build a CalloutComponent for a slide — quote box (verbatim + theme + attribution), data annotation (stat-sig-marked insight), or freeform text. Used by annotate-slide-workflow for callouts on existing slides, and by add-slide-workflow for slides with embedded insight boxes. Python API handles default positioning, low-base flagging (n<30), and significance markers (*/**); LLM augmentation is a future wrapper for prose polish."
---

# callout-writer

Data + theme → `CalloutComponent`. LLM skill, no Python file.

## Cardinal Rules

1. **Verbatim quotes must be real.** Pull from qualitative source data (`qualitative_data.json`). Never paraphrase a respondent's words into a quote. If you must edit for brevity, use `[...]` to mark omissions and flag in attribution.
2. **Attribution includes segment + n.** "HCP, Academic (n=18)" is client-ready. "An HCP said" is not.
3. **≥15 words per quote.** Shorter quotes tend to be fragments, not explanations.
4. **Theme tag is a noun phrase.** "Access friction" not "access is hard." Matches the style of existing hypothesis-bank theme tags.
5. **Callout style must match layout.** `"dashed"` for insight callouts, `"default"` for quote boxes, `"insight"` for data-annotation callouts.

## Inputs / Outputs

**Input:**
- `slide_context`: the slide's chart data + headline (so the callout complements, not repeats)
- `theme`: a theme tag from the hypothesis bank or narrative thread
- `quote_pool` (optional): list of candidate verbatim quotes with attributions
- `data_annotation` (optional): for data-driven callouts, the specific finding to annotate
- `style`: `"default"` (quote box), `"dashed"` (insight), `"insight"` (data annotation)
- `position`: where on the slide the callout goes (passed through from layout-selector)

**Output:** A `CalloutComponent` instance ready to append to `spec.components`.

## Callout types

| Type | Fields populated | When |
|---|---|---|
| Qualitative quote | `theme` + `text` (verbatim) + `attribution` | Workflow 4 (storyboarding), qualitative slides |
| Insight / delta | `text` (short data finding, e.g. "↑8pp among Academic") + `theme` (hypothesis tag) | Workflow 7 (add callout), segment analysis |
| Narrative thread tag | `theme` (arc identifier) + `text` (one-sentence summary) | Cross-slide continuity |

## Decision Rules

| Situation | Response |
|---|---|
| `quote_pool` is empty and `data_annotation` is None | Return `None` — surface to planner; callout may not be needed |
| Theme has no matching quote in pool | Try widening to adjacent themes; if still nothing, return `None` |
| Quote n < 5 | Flag in attribution (e.g. "n=4 — low base"); still acceptable for qual slides |
| Multiple candidate quotes | Pick the one with clearest narrative payoff (not shortest, not longest) |
| LLM can't satisfy all rules | Return `None` and surface which rule blocked |

## Python API

```python
from slidegen.callout_writer import (
    Quote, DataInsight,
    build_quote_callout, build_insight_callout, build_freeform_callout,
    pick_quote_from_pool,
)

# Quote callout — pulls from qualitative_data.json
q = Quote(
    text="Rybrevant's efficacy in 1L is a real differentiator against Tagrisso.",
    attribution="Oncologist, Academic (n=42)",
    theme="efficacy_differentiation",
)
component = build_quote_callout(q, theme_override="Efficacy in 1L")

# Insight callout — data finding with significance marker
insight = DataInsight(
    text="Up 8pp among Academic vs Community",
    p_value=0.008,
    n=212,
    theme="segment_divergence",
)
component = build_insight_callout(insight)
# → "Up 8pp among Academic vs Community **"  (p<0.01 marker)

# Freeform callout — user-supplied text
component = build_freeform_callout(
    text="Client noted: 'watch Q2 to see if this holds'",
    style="dashed",
)

# Pool selection — pick best quote from qualitative_data.json candidates
pool = [Quote(...), Quote(...), ...]
best = pick_quote_from_pool(pool, theme="efficacy_differentiation")
```

## References

- `slidegen/callout_writer.py` — implementation
- `slidegen/slide_spec/schema.py` — `CalloutComponent` fields
- `projects/*/context/*/qualitative_data.json` — verbatim source data
- `slidegen/stat_sig.py` — significance tests that populate `DataInsight.p_value`
- `.claude/skills/workflows/annotate-slide-workflow/SKILL.md` — primary consumer
- PRD §6.7 (spec contract)
