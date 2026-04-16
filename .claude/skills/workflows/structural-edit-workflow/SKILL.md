---
name: structural-edit-workflow
effort: low
paths: []
description: "Triggered by: 'Remove slide 8' / 'Move slide 10 to position 3' / 'Split slide 5 into two slides' / 'Merge slides 7 and 8'. Performs structural edits to deck ordering: delete, reorder, split, merge. Operates on deck structure without touching individual slide data or design. Uses deck-assembler and spec manipulation."
---

# structural-edit-workflow

Structural edits to deck ordering: delete, reorder, split, merge. Doesn't touch slide content/data.

## Trigger phrases

**Remove:**
- "Remove slide 8"
- "Delete slides 15, 20, 23"
- "Drop the appendix slides at the end"

**Reorder:**
- "Move slide 10 to position 3"
- "Swap slides 5 and 7"
- "Put the segment slides after the recall slides"

**Split:**
- "Slide 5 is too busy — split into two slides"
- "Break slide 12 into separate Academic / Community slides"

**Merge:**
- "Combine slides 7 and 8 into one"
- "Merge the two trend slides at the end"

## Cardinal Rules

1. **Deck-level only.** Changes to an individual slide's content/data/design belong in `edit-slide-workflow`, not here.
2. **PPTX backup before write** — always, even for simple deletes. Structural edits are irreversible at the deck level.
3. **Shape registry stays consistent.** Deleting slide N means registry entries for slide N are removed; reordering means indices update; splitting means one new entry replaces one old; merging means one entry replaces two.
4. **Preserve data lineage.** Splitting a slide: each resulting slide keeps the source slide's `DataLineage`. Merging: the result carries both source lineages (merged as a list if schema allows, else first wins with a note).
5. **User review for non-trivial edits.** Delete is safe. Reorder is safe. Split and merge require user confirmation on the result before writing — easy to lose content.

## Inputs

- `deck_path`: existing PPTX
- `action`: `"delete"` | `"reorder"` | `"split"` | `"merge"`
- Action-specific params:
  - delete: `slide_indices: list[int]`
  - reorder: `new_order: list[int]` (permutation of existing indices) OR `(from_index, to_index)` pair
  - split: `slide_index: int`, `split_by: "category_subset" | "series_subset" | "manual"` + params
  - merge: `slide_indices: list[int]` (≥2 adjacent slides), `strategy: "stack" | "side_by_side"`

## Outputs

- Updated PPTX at `deck_path` (backup at `deck_path.bak.{timestamp}`)
- Updated `shape_registry.json` reflecting new slide structure

## Orchestration

```
1. deck-reader(deck_path, all_slides=True)
     → list[SlideSpec] for every slide

2. BRANCH on action:

   ── delete ──
   2a. new_specs = [spec for i, spec in enumerate(specs) if i not in slide_indices]
   2b. re-number slide_index on remaining specs (0-based)

   ── reorder ──
   2c. new_specs = [specs[i] for i in new_order]
   2d. re-number slide_index

   ── split ──
   2e. source_spec = specs[slide_index]
   2f. build 2+ child specs:
         — split category list in two; assign half to each child chart
         — OR split series list (e.g. prior/current vs segment breakdown)
         — OR user-specified split points
   2g. each child spec copies source's: brand, layout, headline (with "(1 of 2)" / "(2 of 2)" suffix),
         footer, data_lineage
   2h. new_specs = specs[:slide_index] + [child1, child2, ...] + specs[slide_index+1:]
   2i. Gate: show before/after to user for confirmation

   ── merge ──
   2j. source_specs = [specs[i] for i in slide_indices]  (must be adjacent)
   2k. build combined spec:
         — "stack" strategy: two vertical sections (uses observed_1chart_2table or custom layout)
         — "side_by_side" strategy: two charts, two tables (uses observed_2chart_2table)
         — headline merged from both (user can edit)
         — components union, repositioned via layout-selector
   2l. Gate: show result to user for confirmation
   2m. new_specs = specs[:min_idx] + [merged_spec] + specs[max_idx+1:]

3. for each new/modified spec: spec-validator(spec, strict=True)

4. deck-assembler(new_specs, output_path=deck_path, template_path=?)
     → PPTX + updated registry + backup
```

## Decision Rules

| Situation | Response |
|---|---|
| Delete indices out of range | Raise with explicit "slide N doesn't exist" |
| Reorder list not a valid permutation | Raise — every existing index must appear exactly once |
| Split requested but no obvious split dimension (single series, single category) | Halt; ask user for manual split points |
| Merge requested on non-adjacent slides | Halt OR auto-reorder first (confirm with user) |
| Merge produces spec that won't fit any LAYOUT preset | Use explicit positions; warn user about non-canonical layout |
| Referenced slides include a cover/divider | Allow delete/reorder; reject split/merge on covers |
| After edit, deck has 0 slides | Halt — refuse to write empty deck |

## References

- PRD §3 Workflow 6, §5 composition map
- `.claude/skills/context-data/deck-reader/SKILL.md`
- `.claude/skills/planning/layout-selector/SKILL.md`
- `.claude/skills/planning/spec-validator/SKILL.md`
- `.claude/skills/creation/deck-assembler/SKILL.md` (handles insert/replace/reorder/delete mechanics)
- Sibling: `.claude/skills/workflows/edit-slide-workflow/SKILL.md` (for within-slide changes)
