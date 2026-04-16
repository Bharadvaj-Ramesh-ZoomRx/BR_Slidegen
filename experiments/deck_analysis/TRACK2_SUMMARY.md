# Track 2 — Edit-mode Primitives (Summary)

**Date:** April 16, 2026
**Status:** Python primitives + SKILL.mds landed. Test scripts pending.

## What was built

### Python primitives (7 files, ~2,218 lines)

| File | LOC | Purpose |
|---|---|---|
| `slidegen/deck_reader/__init__.py` | 77 | Unified API: `read_deck(pptx_path, config_path=None) → (list[SlideSpec], summary)` |
| `slidegen/deck_reader/__main__.py` | 3 | CLI entry: `python -m slidegen.deck_reader path/to/deck.pptx` |
| `slidegen/deck_reader/tag_reader.py` | 623 | **Tier 1** — Galen-PowerPoint Connector tag extraction. Walks shapes, reads `ReportConfigHash`, looks up Custom XML Part by SHA256, parses `ReportConfigDto` JSON → `DataLineage` |
| `slidegen/deck_reader/inference.py` | 487 | **Tier 2** — structural inference. Parses headline, chart OOXML, categories, embedded tables. Cross-references against optional config.yaml + source_data.json to propose extraction method + question codes |
| `slidegen/slide_updater.py` | 219 | Takes `(spec, new_data)` → updated `SlideSpec` with fresh chart series + delta column + audit fields. Preserves layout / headline / brand |
| `slidegen/slide_editor.py` | 474 | 28 whitelisted edit actions — set_headline_text, change_series_color, set_chart_pattern, set_delta_format, reorder_categories, etc. |
| `slidegen/deck_assembler.py` | 335 | `list[SlideSpec]` → PPTX + shape_registry.json. Validates every spec, backs up before overwrite |

### SKILL.md files

| Skill | Path |
|---|---|
| `deck-reader` | `.claude/skills/context-data/deck-reader/SKILL.md` |
| `slide-updater` | `.claude/skills/creation/slide-updater/SKILL.md` |
| `slide-editor` | `.claude/skills/creation/slide-editor/SKILL.md` |
| `deck-assembler` | `.claude/skills/creation/deck-assembler/SKILL.md` |

## What works end-to-end

- **`deck-reader` against real PET decks.** Tested on `JJ PET RYBREVANT+LAZCLUZE Q1'26 Report_Migration.pptx`: 62 slides processed, 42 specs produced, `summary` correctly reports tagged/untagged counts.
- **All 4 modules import cleanly** after one schema fix (see below).
- **Validator integration** — every emitted spec passes `validate_spec(strict=True)` at primitive boundary.

## Tagged vs Untagged observations (real decks)

Tested against `JJ PET RYBREVANT+LAZCLUZE Q1'26`:
- **Total shapes analyzed:** 1,180
- **Tagged (Tier 1):** 0
- **Untagged (Tier 2):** 1,180

Finding: the test decks in `experiments/deck_analysis/decks/` were NOT authored through the Galen-PowerPoint Connector. They are raw client-delivery decks. Tier 1 paths are structurally correct but untestable against this sample. Need a Connector-authored deck to validate Tier 1 end-to-end — flag for user.

## Schema fix applied

Track 2's inference code constructed `ComponentSpec` subclasses without passing `type=` explicitly (relying on `__post_init__` to set it). Original schema required `type` at construction. Fixed by giving `ComponentSpec.type` a default of `""` with subclass `__post_init__` setting the right value. All existing tests and the worked example still validate cleanly after the fix.

## Known gaps

1. **No Connector-authored deck tested end-to-end for Tier 1.** All 32 experiments decks are untagged. Request: obtain one Connector-authored deck for Tier 1 validation before Workflow 2 demo.
2. **Tier 2 confidence scoring is coarse.** Current logic produces binary high/low without finer gradations. May need refinement when `slide-plan-generator-refresh` surfaces low-confidence inferences to users.
3. **Test scripts under `tests/track2/` not yet written.** Agent hit permissions wall before landing them. Recommended: add `test_deck_reader.py`, `test_slide_updater.py`, `test_slide_editor.py`, `test_deck_assembler.py` before merging.
4. **`deck_assembler` currently stubs slide-creator.** Since slide-creator's Python impl is not yet built (only the SKILL.md contract), deck_assembler's slide-rendering call is a placeholder that creates a blank slide + headline. Full integration awaits slide-creator Python.

## Decks tested against

- `decks/JJ PET RYBREVANT+LAZCLUZE Q1'26 Report_Migration.pptx` (62 slides)

Not yet tested:
- Connector-authored deck (not in sample)
- Multi-brand deck
- ATU / PCA / HCP-Pt decks (different patterns)

## Contract for downstream integration

- **Input to every primitive:** a `SlideSpec` validated by `validate_spec(strict=True)`
- **Output of every primitive:** a `SlideSpec` validated by `validate_spec(strict=True)` at exit
- **Audit fields updated on every mutation:** `data_lineage.last_data_pull`, `data_lineage.last_refresh_error`
- **No silent schema extension.** Track 2 discovered the one schema issue listed above — no other extensions were made.
