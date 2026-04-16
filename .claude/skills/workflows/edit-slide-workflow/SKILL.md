---
name: edit-slide-workflow
effort: low
paths: []
description: "Triggered by: 'Regenerate slide N' / 'Rebuild slide 12' / 'Update slide 5 with the latest data' / 'Slide 8 is stale — refresh it'. Edits ONE slide in an existing deck: rebuild from current spec (fast path), re-pull fresh data then rebuild (data-refresh path), or apply whitelisted spec edits (edit path)."
---

# edit-slide-workflow

Edit or refresh ONE slide in an existing deck. Routes between fast rebuild and data-refresh based on user intent.

## Trigger phrases

**Rebuild-only (fast path, no data pull):**
- "Regenerate slide N"
- "Rebuild slide 12"
- "Slide 5 looks wrong — rerender it"
- "Apply my spec edits to slide 8 and rebuild"

**Data refresh (slow path, re-queries Synapse/Excel):**
- "Update slide 12 with the latest data"
- "Pull fresh numbers for slide 5"
- "Slide 8 is stale — refresh its data"
- "Re-query Synapse for the data on slide 20"

**Targeted edits (structural change on existing slide):**
- "Change slide 5's headline to 'Efficacy recall holds steady'"
- "Swap slide 12's chart pattern to line_markers_trended"
- "Change series colors on slide 20"

Detection: if user mentions "data", "fresh", "update with", or "pull" → data-refresh mode. If user mentions "regenerate", "rebuild", "rerender" → fast rebuild. If user specifies a concrete edit (headline text, color, pattern) → targeted edit. Default when ambiguous: fast rebuild.

## Cardinal Rules

1. **One slide only.** For multi-slide changes use `refresh-deck-workflow` (data) or `structural-edit-workflow` (adds/removes/reorders).
2. **Preserve everything not explicitly changed OR structurally implied.** Layout, chart pattern, colors — preserved unless user asks otherwise. **Headline is NOT preserved in data_refresh mode by default** — new data invalidates the old headline's numbers/direction; `slide-updater` calls `headline-writer` to regenerate. User can pass `preserve_headline=True` to force-keep.
3. **Subheadlines with period references get rewritten on data refresh.** "Q3 '25 vs Q4 '25" → "Q4 '25 vs Q1 '26" when wave shifts.
4. **PPTX backup before write.** `deck-assembler` handles this automatically.
5. **Spec validates between mutation and render.** Reject invalid edits loudly.
6. **Data-refresh recomputes deltas.** Never copy prior deltas; always derive `current - prior` from fresh data.
7. **Audit stamp on every execution.** Update `last_data_pull` + clear `last_refresh_error` (or set it on failure) — regardless of path taken.

## Inputs

- `deck_path`: existing PPTX
- `slide_index` (0-based) OR `slide_id` (from `shape_registry.json`)
- `mode` (optional): `"rebuild"` | `"data_refresh"` | `"edit"` (default: auto-detect)
- Optional `edits`: list of `slide-editor` instructions (for edit mode)
- Optional `data_override`: pre-fetched data dict (bypass Synapse call)

## Outputs

- Updated PPTX at `deck_path` (backup at `deck_path.bak.{timestamp}`)
- Updated `shape_registry.json` with new `last_refreshed` + `last_data_pull`

## Orchestration

```
1. deck-reader(deck_path, slide_index=N)
     → SlideSpec for slide N (Tier 1 via Connector tags, or Tier 2 inferred)

2. BRANCH on mode:

   ── mode="rebuild" ──
   (skip data re-fetch)
   go to step 4

   ── mode="data_refresh" ──
   2a. if Tier 2 with confidence="low":
         Gate: user confirms inferred lineage
   2b. synapse-read(spec.data_lineage)  [or excel extraction]
         → fresh data
   2c. slide-updater(spec, new_data, period_labels={"prior": "...", "current": "..."})
         → updated spec with:
            • fresh values + recomputed deltas
            • REGENERATED headline via headline-writer (driven by new data + narrative context)
            • REWRITTEN subheadline period labels (if present)
            • stamped audit fields (last_data_pull, last_refresh_error cleared)
   go to step 4

   ── mode="edit" ──
   3. for each edit in edits:
        slide-editor(spec, edit)
      → updated_spec
   go to step 4

4. spec-validator(spec, strict=True)
     → reject incomplete spec

5. slide-creator(spec)
     → rendered Slide

6. deck-assembler(existing_deck=deck_path, replace_at=N, with=[spec])
     → backs up deck_path, replaces slide N, updates registry
```

## Decision Rules

| Situation | Response |
|---|---|
| Slide index out of range | Raise — don't silently extend the deck |
| Tier 2 confidence="low" in data_refresh mode | Halt at 2a; user confirms or corrects lineage |
| Synapse returns empty data in data_refresh mode | Keep prior data; set `last_refresh_error="no data available"`; still stamp `last_data_pull` |
| data_refresh mode but data categories mismatch | Raise `SlideUpdateError` — switch to `edit` mode with `reorder_categories` edit |
| edit mode with unsupported action | Raise via `slide-editor`'s whitelist check |
| Slide has no `data_lineage` (pure text/callout/cover) | Skip data_refresh; fall through to rebuild mode with warning |
| User wants both data refresh + targeted edits | Run data_refresh first, then edit mode on the result |

## References

- PRD §3 Workflow 3, §5 composition map
- `.claude/skills/context-data/deck-reader/SKILL.md`
- `.claude/skills/creation/slide-updater/SKILL.md` (data path)
- `.claude/skills/creation/slide-editor/SKILL.md` (edit path)
- `.claude/skills/creation/slide-creator/SKILL.md` (renderer)
- `.claude/skills/creation/deck-assembler/SKILL.md` (final write)
- `.claude/skills/planning/spec-validator/SKILL.md`
