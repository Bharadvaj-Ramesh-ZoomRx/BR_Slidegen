# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Generic pharma consulting report generator — builds PowerPoint slide decks from survey data using a **YAML-driven pipeline**. Each project (brand/product) is defined by a YAML config file; no code changes needed to add new projects.

Currently configured for **Rybrevant (RYB) + Lazcluze** vs **Tagrisso (TAG)** — covers message recall, effectiveness, rep performance, call-to-action metrics, prescription intent, and high-impact interactions across Q3/Q4 2025.

## Folder Structure

```
├── projects/                  # Project folders (one per product/brand)
│   └── jnj_rybrevant/         # J&J Rybrevant PET Q4'25
│       ├── config.yaml        # YAML config — brands, sheets, extractions, asks
│       ├── input/             # ★ USER DROP ZONE — all source files (gitignored)
│       │   └── wave/                         # Update every wave
│       │       └── PET_Q3Q4_2025/           #   One folder per wave
│       │           ├── source_data.xlsx     #     Survey data (Excel)
│       │           ├── call_notes.docx      #     Client call notes
│       │           ├── hypotheses.xlsx      #     Wave hypotheses + client intel
│       │           ├── prior_wave_es.md     #     Prior wave ES findings
│       │           ├── pet_project_kbq.odt  #     Study design + KBQs
│       │           ├── market_context.md    #     Curated market context
│       │           ├── kbqs.md              #     Standing KBQs
│       │           └── survey_context.md    #     Survey instrument + Q codes
│       ├── context/           # ★ SYSTEM GENERATED intermediates (gitignored)
│       │   └── PET_Q3Q4_2025/ # Wave-versioned context outputs
│       │       ├── project_context.md        # Stage 1: /build-project-context
│       │       ├── hypothesis_bank.md        # Stage 2: /hypotheses
│       │       ├── slide_plan.md             # Stage 3: /slide-plan
│       │       └── source_data.json          # Auto-extracted from Excel (JSON cache)
│       ├── templates/         # Template decks — shared across waves (gitignored)
│       ├── output/            # Generated deliverables (gitignored)
│       │   └── PET_Q3Q4_2025/ # Wave-versioned output subfolder
│       │       ├── deck.pptx          # Generated deck
│       │       ├── shape_registry.json # Shape state with data lineage
│       │       └── backups/           # PPTX backups before edits (max 10)
│       └── config_history/    # Timestamped config backups (gitignored)
├── slidegen/                  # SlideGen system
│   ├── pipeline/              # ★ Generic deck generation pipeline
│   │   ├── __init__.py        # Exports: generate_deck(), regenerate_slide()
│   │   ├── project_config.py  # ProjectConfig dataclasses + YAML loader
│   │   ├── data_loaders.py    # 5 generic data extractors + JSON auto-cache
│   │   ├── slide_renderers.py # 9 slide type renderers (RENDERERS registry)
│   │   ├── orchestrator.py    # Pipeline entry + ShapeNamer + per-slide regen + PPTX backup
│   │   └── config_generator.py # Data discovery + config scaffolding helpers
│   ├── pptx_utils/            # ★ Utility package (PRD §4.2)
│   │   ├── __init__.py        # Re-exports everything for backward compat
│   │   ├── brand.py           # BRAND{} dict, color constants, fonts, slide dims
│   │   ├── lxml_helpers.py    # 20 lxml XML manipulation functions
│   │   ├── shapes.py          # 7 shape primitives (textbox, solidrect, etc.)
│   │   ├── layout.py          # LAYOUTS{} dict + 10 slide chrome functions
│   │   ├── charts.py          # CHART_PATTERNS{} dict + 4 chart builders
│   │   ├── tables.py          # 3 table builders (delta col/table, value table)
│   │   ├── com.py             # 7 COM helpers for live editing
│   │   └── registry.py        # 6 registry CRUD functions
│   ├── __init__.py            # Package exports: SlideBuilder, LiveEditor, reconcile
│   ├── __main__.py            # CLI: python -m slidegen <create|edit|reconcile>
│   ├── config.py              # Centralized paths and settings
│   ├── create.py              # SlideBuilder class — python-pptx creation + registry
│   ├── edit.py                # LiveEditor class — win32com live editing + edit log
│   ├── reconcile.py           # Registry reconciliation from live PowerPoint state
│   └── output/                # Generated slides (gitignored)
├── archive/                   # Superseded scripts & old artifacts
│   ├── src/                   # Legacy pipeline (original POC, hardcoded J&J)
│   │   ├── generate_asks.py   # Monolithic 14-slide J&J deck generator
│   │   ├── extract_data.py    # Excel → output/slide_data_v2.pkl
│   │   ├── generate_pptx.py   # pkl → full multi-slide deck
│   │   ├── generate_slide1.py # Single-slide POC generator
│   │   └── validate_data.py   # Cross-check pkl against raw Excel
│   ├── scripts/               # Ad-hoc exploration & discovery
│   │   └── discover_data.py   # Explores Excel sheets
│   ├── explore*.py            # Early data exploration scripts
│   └── pptx.zip               # Old PPTX artifacts
├── docs/                      # Design docs & architecture notes
├── .claude/skills/            # Claude Code skills (auto-discovered)
│   ├── slidegen/              # ★ Primary skill (pipeline + utils + styling + references)
│   │   └── references/        # Consolidated: function-ref, chart-patterns, brand-constants, archetypes
│   └── pptx/                  # General PPTX read/create/edit skill (non-pipeline)
└── .gitignore
```

## Running

```bash
# ── Generic Pipeline (YAML-driven) ──
python -m slidegen.pipeline.orchestrator projects/jnj_rybrevant/config.yaml
# Or from Python:
#   from slidegen.pipeline import generate_deck
#   generate_deck("projects/jnj_rybrevant/config.yaml")

# ── SlideGen System ──
python -m slidegen create                   # Demo slide creation
python -m slidegen edit <filename.pptx>     # Interactive live editor
python -m slidegen reconcile <filename.pptx> # Sync registry from PowerPoint

# ── Legacy (archived) ──
python archive/src/generate_asks.py        # Monolithic 14-slide deck (POC)
python archive/scripts/discover_data.py    # Data exploration
```

## Pipeline Architecture

```
projects/jnj_rybrevant/config.yaml  →  ProjectConfig (dataclasses)
                                      ↓
source_data.json (or Excel)  →  data_loaders.load_all_data()  →  dict[extraction_id → list[dict]]
                                      ↓
orchestrator  →  RENDERERS[slide_type](slide, config, ask, data, namer)
                                      ↓
                      output/{wave}/deck.pptx
                      output/{wave}/shape_registry.json   (shape state + lineage)
                      output/{wave}/backups/               (PPTX backups)
```

### Data Loading (Auto-JSON)

On first run, data is extracted from Excel and saved as `context/{wave}/source_data.json`. Subsequent runs read the JSON directly — no pandas, no column indices, no question-code walking. The JSON auto-invalidates when the Excel file changes (hash check). Delete `source_data.json` to force re-extraction.

### Data Extraction Methods

| Method | Use Case |
|--------|----------|
| `question_code` | Find a code row, walk sub-rows, extract prior/current values |
| `multi_question_code` | One row per code (e.g. CTA: compelling, changed opinion, closing) |
| `row_range` | Fixed row range with column mapping |
| `question_code_multi_col` | Multiple columns per row (e.g. HII: hi vs other) |
| `nested_ordinal` | Grouped ordinal sub-rows (e.g. 1st/2nd/3rd recall order) |

### Slide Type Renderers

| Type | Pattern |
|------|---------|
| `cover` | Title slide |
| `executive_summary` | Bullet-list insights |
| `single_bar_with_delta` | Horizontal bar + QoQ delta column |
| `dual_bar_with_delta` | Two side-by-side bars + deltas |
| `dual_bar_qoq` | Two side-by-side Q4-vs-Q3 clustered bars + deltas |
| `clustered_compare` | Clustered bar comparing two groups + gap/delta columns |
| `qoq_bar_with_delta` | Q4 vs Q3 clustered + delta |
| `two_section_bar` | Two vertically stacked bar sections |
| `stacked_order` | Stacked bar with ordinal breakdown + total column |

## Data Source Layout

All scripts read `source_data.xlsx` (originally "Lung SFEA SB.xlsx") with `header=None` (0-indexed rows/cols):

- **RYB sheet**: col 0=code, col 1=desc, col 7=Q3 Total, col 17=Q4 Total
- **TAG sheet**: col 0=code, col 1=desc, col 7=Q3 Total, col 13=Q4 Total
- **Additonal Analysis sheet** (note: misspelled in source): col 1=metric, col 2=RYB_Q3, col 3=RYB_Q4, col 4=TAG_Q3, col 5=TAG_Q4; message rows 30-39 use cols 3-8 for MR/ME/Believable

## Key Conventions

- **Percentage conversion**: Raw Excel values are decimals (0.0–1.0). `pct()` multiplies by 100 and rounds to 1 decimal. Some rows store whole-number percentages — use `straight()` not `pct()` for those (set `pct_mode: straight` in YAML extraction params).
- **Delta**: Always current minus prior in percentage points.
- **Sorting**: Configurable per-ask via `sort_by` and `sort_desc` in YAML.
- **Brand colors**: Defined per-project in YAML. J&J: RYB orange (`#F75824` current, `#FFC199` prior), TAG violet (`#7030A0` current, `#AD88C8` prior). Delta colors: green positive, red negative.
- **Data format**: All extractors return standardized dicts with `desc`, `prior`, `current` keys (generic — not `q3`/`q4`).
- **Template strings**: Headlines/sections support `{{primary.name}}`, `{{period_current}}`, `{{client}}` etc.

## Dependencies

- `pandas`, `openpyxl` (Excel reading — only on first run per wave)
- `python-pptx` (PowerPoint generation)
- `lxml` (XML manipulation for native PPT charts)
- `pyyaml` (YAML config loading)
- `pywin32` (win32com — live PowerPoint editing via COM, Windows only)

## SlideGen API

```python
# ── Generic Pipeline ──
from slidegen.pipeline import generate_deck
generate_deck("projects/jnj_rybrevant/config.yaml")

# ── Regenerate a single slide ──
from slidegen.pipeline import regenerate_slide
regenerate_slide("projects/jnj_rybrevant/config.yaml", slide_index=4)

# ── Data discovery (for new projects) ──
from slidegen.pipeline.config_generator import discover_excel_structure
summary = discover_excel_structure("projects/new_project/data/source_data.xlsx")

# ── Config backup ──
from slidegen.pipeline.orchestrator import _backup_config
_backup_config("projects/jnj_rybrevant/config.yaml")
# → projects/jnj_rybrevant/config_history/config_20260311_143000.yaml

# ── Edit live (file must be open in PowerPoint) ──
from slidegen.edit import LiveEditor
with LiveEditor("output.pptx") as editor:
    editor.set_text("zrx_001", "New text", color="red")
    editor.set_slide(5)  # switch to slide 5
    editor.find_shapes_by_prefix("zrx_005")  # list all shapes on slide 5
```

## User Input Patterns

All workflows are triggered via **natural language** in the Claude Code terminal:

| Scenario | User Says | Pipeline Call |
|----------|-----------|---------------|
| Edit 1 slide (text/sort/data) | `Edit Slide N — ...` | `regenerate_slide()` |
| Edit slides with new wave data | `Edit slides with new wave data — PET_Q1Q2_2026` | update config → `generate_deck()` |
| Add/remove/reorder slides | `Add slide after N...` / `Remove Slide N` | `generate_deck()` |
| Edit slides with new wave data + new asks | `Edit slides with new wave data + new asks — PET_Q1Q2_2026` | update config + extractions → `generate_deck()` |
| Brand new project | `Create slides for projects/{name}` | full create workflow |

### Example inputs:

```
Create slides for projects/jnj_rybrevant
Create slides for projects/pfizer_ibrance wave PET_Q1_2026

Edit Slide 5 — change headline to "Updated Message Recall"
Edit Slide 3 — sort bars descending by current value
Edit Slide 9 — use data from Q2_15Z instead of Q2_10Z

Edit slides with new wave data — PET_Q1Q2_2026, Q1'26 vs Q2'26
Edit slides with new wave data + new asks — PET_Q1Q2_2026

Add a slide after Slide 6 — clustered_compare for HCP satisfaction
Remove Slide 8
Move Slide 10 before Slide 5
Regenerate all slides
```

## Workflow: Create Slides

When the user says **"Create slides for projects/{name}"** for a new project with data/reference/templates:

1. **Verify folder structure** — Ensure `projects/{name}/input/wave/{wave}/`, `templates/` exist
2. **Discover data** — Run `discover_excel_structure()` on the source Excel to understand sheets, columns, question codes
3. **Read reference docs** — Parse reference files (`.md`, `.docx` via `python -m markitdown`) to understand what slides are needed
4. **Read template** — Run `python -m markitdown template.pptx` to understand available layouts
5. **Generate config scaffold** — Call `generate_config_scaffold()` for a starter YAML
6. **Map asks to pipeline** — For each ask from the reference doc:
   - Match to an existing extraction method (`question_code`, `row_range`, etc.)
   - Match to an existing slide type (`single_bar_with_delta`, `clustered_compare`, etc.)
   - If no existing method/type fits → add a new extractor to `data_loaders.py` or new renderer to `slide_renderers.py` (follow existing patterns, don't modify existing code)
7. **Write config.yaml** — Complete the YAML with all extractions and asks
8. **Generate deck** — Run `generate_deck("projects/{name}/config.yaml")`
9. **Visual QA** — Convert to images, inspect for layout issues, fix and re-run

## Workflow: Edit Slide

When the user says **"Edit Slide N — ..."** (e.g. change headline, update data, re-sort):

1. **Read config** — Load current `config.yaml`, identify the ask at index N-1
2. **Backup config** — Call `_backup_config()` to save timestamped copy to `config_history/`
3. **Update config** — Modify the relevant fields in the YAML (headline, data_key, sort_by, etc.)
4. **Regenerate slide** — Call `regenerate_slide(yaml_path, slide_index=N-1)` — automatically creates PPTX backup in `output/{wave}/backups/`, then clears and re-renders only that slide
5. **Refresh PowerPoint** — Tell user to reopen/refresh the deck in PowerPoint to see changes
6. **Report** — Summarize what changed

## Workflow: Edit Slides with New Wave Data

When the user says **"Edit slides with new wave data — {wave_id}"**:

1. **Backup config** — Call `_backup_config()` to save timestamped copy
2. **Update config** — Change `project.wave`, `period_current`, `period_prior` in config.yaml
3. **Verify data** — Check `input/wave/{wave_id}/source_data.xlsx` exists
4. **Regenerate deck** — Run `generate_deck()` — auto-extracts Excel → saves `source_data.json` → output goes to `output/{wave_id}/deck.pptx`
5. **Report** — Old wave output is preserved, new wave output is in its own folder

## Workflow: Edit Slides with New Wave Data + New Asks

When the user says **"Edit slides with new wave data + new asks — {wave_id}"**:

1. **Backup config** — Call `_backup_config()` to save timestamped copy
2. **Update wave** — Change `project.wave`, `period_current`, `period_prior`
3. **Read new reference** — Parse reference docs for updated asks
4. **Update config** — Modify extractions and asks in config.yaml to match new reference
5. **Gap analysis** — Check if new asks need new extractors or renderers; extend if needed
6. **Regenerate deck** — Run `generate_deck()` — output goes to `output/{wave_id}/deck.pptx`
7. **Report** — Summarize changes from previous wave

## Shape Naming

Pipeline-generated slides assign `zrx_{slide:03d}_{shape:03d}` names to all shapes (e.g. `zrx_001_001` through `zrx_001_008` for slide 1). This enables:
- LiveEditor targeting via COM
- Shape registry saved to `output/{wave}/shape_registry.json` with data lineage per slide
- Per-slide regeneration without affecting other slides
- PPTX backup before each `regenerate_slide()` call (stored in `output/{wave}/backups/`)

## Wave Versioning

Each project supports wave-based folder versioning for input data and output:

```yaml
project:
  wave: "PET_Q3Q4_2025"          # wave identifier

data_source_path: "input/wave/{{wave}}/source_data.xlsx"  # wave-versioned (user drop zone)
template_path: "templates/template.pptx"              # shared (no {{wave}})
output_path: "output/{{wave}}/deck.pptx"              # wave-versioned
context_path: "context/{{wave}}/"                     # system-generated intermediates
```

- `{{wave}}` in paths is interpolated from `project.wave` at config load time
- Template stays flat — shared across waves (typically doesn't change)
- Input, context, and output get wave subfolders — old waves are preserved
- To start a new wave: update `project.wave`, `period_current`, `period_prior` in config.yaml and place new data in `input/wave/{new_wave}/`
- If `wave` is omitted or empty, paths are used as-is (backward compatible)

## Config Versioning

Before any config modification, a timestamped backup is saved:
- Location: `projects/{name}/config_history/config_{YYYYMMDD_HHMMSS}.yaml`
- These are gitignored by default

# currentDate
Today's date is 2026-03-16.

      IMPORTANT: this context may or may not be relevant to your tasks. You should not respond to this context unless it is highly relevant to your task.
