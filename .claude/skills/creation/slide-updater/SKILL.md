---
name: slide-updater
effort: medium
paths: ["slidegen/slide_updater.py"]
description: "Refresh a SlideSpec with new data, preserving layout, headline, brand, and metadata. Takes (existing_spec, new_data) and returns an updated spec with fresh chart values + delta column + data labels + last_data_pull timestamp. Trigger when: workflow 2 (wave refresh) has new wave data for an existing slide, workflow 6 (slide update) pulls fresh data for one slide, or any refresh flow where the spec structure stays identical but numbers change. Use edit-slide for structural changes."
---

# slide-updater

Takes an existing `SlideSpec` + fresh data, returns an updated `SlideSpec` with the same structure and new values. Preserves everything the user can see (layout, headline, brand, metadata) except the data itself.

## Cardinal Rules

1. **Data mutation invalidates the old headline.** Change `ChartData.series.values`, `DeltaColumnComponent.values`, and any data labels computed from those. **Also invoke `headline-writer` to regenerate the headline** from the new data — a headline like "Efficacy recall dipped 3pp QoQ" becomes a client-delivery risk when new data says "up 2pp". If the caller wants to preserve the prior headline (e.g. it's a wording-only edit and data happens to be the same), they pass `preserve_headline=True` and acknowledge the mismatch. Layout, chart_pattern, series names, series colors, footer, metadata, and brand are preserved unchanged.
2. **Every output is a fresh object.** Return a new `SlideSpec` — don't mutate the input in place. The caller may want to diff before/after.
3. **Stamp audit fields.** Set `data_lineage.last_data_pull` to now (ISO 8601). Clear `data_lineage.last_refresh_error`. Preserve everything else in `data_lineage`.
4. **Validate on exit.** `validate_spec(output, strict=True)` before returning. If the data mutation breaks coherence (e.g. values-length mismatch), fail loudly.
5. **Recompute deltas from the data.** If both prior-wave and current-wave values are present, compute `delta_column.values` as `current - prior` per row. Never trust pre-existing delta values when the data changed.
6. **Refresh subheadlines that cite period labels.** Subheadlines like "Q4 '25 vs Q3 '25" must be rewritten to reflect the new period labels from the refreshed data.

## Inputs / Outputs

```python
from slidegen.slide_updater import update_slide_data

updated_spec = update_slide_data(
    spec: SlideSpec,                    # existing spec (typically from deck-reader)
    new_data: dict,                     # fresh data keyed by category
    errored: str | None = None,         # set last_refresh_error if non-None
    preserve_headline: bool = False,    # if True, skip headline-writer; keep prior headline as-is
    period_labels: dict | None = None,  # {"prior": "Q3 '25", "current": "Q4 '25"} for subheadline refresh
) -> SlideSpec
```

`new_data` shape follows the standard extraction format:
```python
{
    "categories": ["Efficacy", "Safety", ...],       # must match spec's chart.data.categories
    "series": [
        {"name": "Q4 '25", "values": [0.45, ...]},   # prior wave
        {"name": "Q1 '26", "values": [0.42, ...]},   # current wave
    ],
}
```

## Decision Rules

| Situation | Response |
|---|---|
| `new_data.categories` matches spec's categories exactly | Update values in place; recompute deltas |
| `new_data.categories` order differs but covers the same set | Reorder data to spec's order; update values |
| `new_data.categories` missing some of spec's categories | Raise `SlideUpdateError` — caller must decide (add row or drop row? that's slide-editor, not slide-updater) |
| `new_data` has extra categories not in spec | Raise `SlideUpdateError` — same reason |
| Spec has chrome overrides (custom labels, colors) | Preserve them; don't touch chrome |
| `errored` is non-None | Keep prior data; set `data_lineage.last_refresh_error = errored`; still update `last_data_pull` so the audit trail shows the attempt |
| Headline references a specific delta, direction, or magnitude | Regenerate via `headline-writer` (default behavior). Headline stale-ness is a client-delivery risk — "dipped 3pp" can't stay on a slide whose new data says "up 2pp". |
| `preserve_headline=True` but headline mentions numbers contradicted by new data | Still preserves (caller asked), but emit a warning to log. `deck-audit-workflow` will flag this as a BLOCKER if not reconciled. |
| `period_labels` provided and subheadline text matches the pattern | Rewrite period references in subheadline (e.g. "Q3 '25 vs Q4 '25" → "Q4 '25 vs Q1 '26"). |

## References

- `slidegen/slide_updater.py` — implementation
- `.claude/skills/creation/slide-editor/SKILL.md` — for structural changes (use that, not this)
- `.claude/skills/context-data/deck-reader/SKILL.md` — the typical source of `spec` inputs
- PRD §6.7 (spec contract), §6.8 (lineage audit fields)
