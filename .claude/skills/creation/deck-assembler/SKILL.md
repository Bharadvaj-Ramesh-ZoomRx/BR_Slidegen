---
name: deck-assembler
effort: medium
paths: ["slidegen/deck_assembler.py"]
description: "Compose a list of validated SlideSpec into a final PPTX. Calls slide-creator per spec, handles ordering, insertion, replacement, reordering, and section breaks. Takes either (list[SlideSpec], output_path) for full-deck assembly or (existing_pptx, list[change]) for incremental edits. Updates shape_registry.json. Trigger when: any workflow has finalized its list of SlideSpec and needs a PPTX on disk. The final step of every workflow that produces or modifies a deck."
---

# deck-assembler

Composes specs into a PPTX. Final step of every workflow that writes to disk.

## Cardinal Rules

1. **Never plan.** The spec list is final; deck-assembler doesn't reason about order, headlines, or content. If the list is wrong, fix it upstream.
2. **Validate every spec before rendering.** `validate_spec(spec, strict=True)` on each. Reject the whole deck if any spec is invalid — don't partially assemble.
3. **Deterministic ordering.** Spec order in the input list = slide order in output PPTX. No reordering inside the assembler.
4. **Always back up before overwrite.** If `output_path` exists, copy to `output_path.bak.{timestamp}` first.
5. **Registry is authoritative.** Write `shape_registry.json` alongside the PPTX with data lineage per slide. Downstream tools (slide-updater, refresh workflows) rely on it.

## Inputs / Outputs

```python
from slidegen.deck_assembler import assemble_deck

# Full-deck assembly
assemble_deck(
    specs: list[SlideSpec],
    output_path: str | Path,
    template_path: str | Path | None = None,   # optional master/theme template
) -> dict                                      # {"output_path", "slides", "backup_path"}

# Incremental edit (replace slide N in existing deck)
# Covered by slide-updater + deck-assembler composition in workflow skills
```

CLI (once integrated with pipeline):
```
python -m slidegen.deck_assembler --specs specs/*.json --out deck.pptx [--template tpl.pptx]
```

## Decision Rules

| Situation | Response |
|---|---|
| Any spec fails validation | Raise + halt; do not write partial PPTX |
| `output_path` exists | Backup first, then overwrite |
| `template_path` provided | Start from template; preserve master slides, section breaks, theme |
| No template | Start from a clean blank deck; use `pptx_utils/deck.py` primitives |
| Multiple specs share `slide_id` | Raise — slide_ids must be unique in a deck |
| Spec has `slide_index` but order differs from list position | List position wins; update spec's `slide_index` to match |

## References

- `slidegen/deck_assembler.py` — implementation
- `.claude/skills/creation/slide-creator/SKILL.md` — called per spec
- `.claude/skills/creation/slide-updater/SKILL.md`, `slide-editor/SKILL.md` — produce specs that feed this skill
- `slidegen/pptx_utils/deck.py` — `load_template`, `clear_slide`, `create_sections`
