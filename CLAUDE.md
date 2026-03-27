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
│       │           ├── prior_wave_es.md     #     Prior wave ES findings (text)
│       │           ├── [wave_report].pptx   #     Prior wave deck — read by /build-project-context
│       │           ├── pet_project_kbq.odt  #     Study design + KBQs
│       │           ├── kbqs.md              #     Standing KBQs
│       │           └── survey_context.md    #     Survey instrument + Q codes
│       ├── context/           # ★ SYSTEM GENERATED intermediates (gitignored)
│       │   ├── market_context.md             # /market-context — product-level, NOT wave-versioned
│       │   └── PET_Q3Q4_2025/ # Wave-versioned context outputs
│       │       ├── prior_wave_context.md     # Stage 0.5a: /prior-wave-context (if prior files exist)
│       │       ├── survey_context.md         # Stage 0.5b: /survey-context (if survey draft exists)
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
│   │   ├── slide_renderers/   # 20 slide type renderers (RENDERERS registry)
│   │   │   ├── __init__.py   #   Registry + exports
│   │   │   ├── _shared.py    #   Layout constants, helpers, _auto_label_width
│   │   │   ├── bar.py        #   single_bar, qoq_bar, two_section_bar
│   │   │   ├── bar_dual.py   #   dual_bar_with_delta, dual_bar_qoq
│   │   │   ├── compare.py    #   clustered_compare, stacked_order, dual_bar_compare, hii_scorecard, dual_doughnut
│   │   │   ├── dot.py        #   abacus, dual_abacus, followup_rep, message_mbd
│   │   │   ├── line.py       #   trended_scorecard, trended_activity
│   │   │   ├── quadrant.py   #   quadrant_scatter
│   │   │   ├── heatmap.py    #   heatmap_table
│   │   │   └── narrative.py  #   cover, executive_summary
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
| `single_bar_with_delta` | Table-based horizontal bar + QoQ delta column |
| `dual_bar_with_delta` | Two side-by-side bars + deltas (alternate row backgrounds) |
| `dual_bar_qoq` | Two side-by-side Q4-vs-Q3 clustered bars + deltas |
| `clustered_compare` | Table-based clustered bar comparing two groups + gap/delta columns |
| `dual_bar_compare` | Side-by-side dual brand bar comparison + insight callout |
| `qoq_bar_with_delta` | Q4 vs Q3 clustered + delta |
| `two_section_bar` | Two vertically stacked bar sections |
| `stacked_order` | Table-based stacked bar with ordinal breakdown + total column |
| `abacus` | XY scatter abacus with label/value tables + delta column |
| `dual_abacus` | Two side-by-side abacus panels (e.g. Acad vs Comm by brand) |
| `followup_rep` | Template slide 51-style follow-up rep abacus with dual delta columns |
| `hii_scorecard` | Multi-section clustered column chart with section headers + callouts |
| `dual_doughnut` | Side-by-side doughnut pairs comparing patient segments by brand |
| `message_mbd` | Multi-column abacus for Motivation/Believability/Differentiation breakdown |
| `trended_scorecard` | Multi-panel mini line chart grid (QoQ trend scorecard) |
| `trended_activity` | Side-by-side line + stacked column panels (reach/SOV/frequency) |
| `quadrant_scatter` | 2×2 quadrant scatter chart (stated vs derived importance) |
| `heatmap_table` | Heatmap table with green gradient fills + QoQ delta columns |

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
- **Speaker notes**: Every slide automatically gets speaker notes with question codes and question text used to create it.
- **Table-based layouts**: Bar chart renderers (`single_bar_with_delta`, `clustered_compare`, `stacked_order`) use a separate label table + chart (hidden cat labels) + delta column. Label table width is dynamic via `_auto_label_width()` based on longest label text. Labels wrap to 2 lines if needed.
- **Data labels**: Bar charts use `inEnd` position with white text to prevent overflow. Abacus scatter charts show per-point percentage labels above dots with dynamic y-offset based on row count.
- **Abacus extra options**: `hide_val_cols: true` removes value columns (when data labels show values), `current_field`/`prior_field` remap data fields, `color_current`/`color_prior` override brand colors, `legend_current`/`legend_prior` for custom legend text.
- **Alternating row backgrounds**: Label tables and delta/value tables use consistent grey/white alternating rows (grey first).

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

## Workflow: Create Slides (Full Pipeline)

When the user says **"Create slides for projects/{name}"** or **"Run the full create workflow"**:

**Gate structure: Stages 0 through 0.5c run automatically without individual gates — each sub-skill asks only one lightweight file-list confirmation before extracting. The single user validation gate is at the end of Stage 1: the user reviews ALL generated context files before Stage 2 begins. Stages 5–6 are internal — no gate.**

### Stage 0 — Index Excel → JSON (automatic, always)
Converts `source_data.xlsx` into `context/{wave}/source_data.json`. Runs first, no prompt.
```python
from slidegen.pipeline import index_excel
index_excel("projects/{name}/input/wave/{wave}/source_data.xlsx",
            "projects/{name}/context/{wave}/source_data.json")
```

### Stage 0.5a — Market Context (`/market-context`) — auto if context/market_context.md missing
Checks if `context/market_context.md` already exists. If yes, uses it as-is (no re-run). If missing, auto-generates from Claude's clinical/competitive knowledge (+ any enrichment files in the project folder), runs the 3-round adversarial fact-check, and writes `context/market_context.md`. To force regeneration, user must explicitly say "regenerate market context."
**No gate** — runs silently. Fact-check summary shown inline; user only pauses if flagged claims need review.

### Stage 0.5b — Prior Wave Context (`/prior-wave-context`) — auto if prior wave files detected
Scans `input/wave/{wave}/` for prior wave files (`.pptx` reports, `prior_wave_es.md`, readout `.docx`). If found: shows the user one file-list confirmation, extracts all confirmed files, writes `context/{wave}/prior_wave_context.md`. If not found: skipped. Files already consumed here are **excluded** from Stage 1 scanning.
**One lightweight prompt** — file list only. No content gate.

### Stage 0.5c — Survey Context (`/survey-context`) — auto if survey draft detected
Scans `input/wave/{wave}/` for survey draft files (`survey`, `questionnaire`, `draft`, `instrument` in name; any format). If found: one file-list confirmation, extracts + parses question codes + message lists, writes `context/{wave}/survey_context.md`. If not found: uses hand-written `survey_context.md` from input folder if present, otherwise flags as missing. Files consumed here are **excluded** from Stage 1. Survey context does **not** feed Stage 1 — it feeds Stage 2 directly.
**One lightweight prompt** — file list + code/message count. No content gate.

### Stage 1 — Build Project Context (`/build-project-context`)
Reads the **remaining** input files not already consumed by Stages 0.5b/c (call notes, methodology `.odt`, KBQs, any misc docs) **plus** the context files already generated: `context/market_context.md` and `context/{wave}/prior_wave_context.md`. Does NOT re-read prior wave reports or survey drafts — those are handled upstream. Synthesizes `context/{wave}/project_context.md` covering: study design, KBQs, wave hypotheses (from call notes + prior wave context), analytical priorities, methodology notes, message reference.

**⚠ KBQs dependency:** `KBQs.md` must exist in `input/wave/{wave}/` — it is hand-written by the research team and has no auto-generation path. If missing, flag and ask user to provide it before Stage 2.

**→ SINGLE VALIDATION GATE after Stage 1:** Present all generated context files for user review:
- `context/market_context.md` — product-level competitive/clinical context
- `context/{wave}/prior_wave_context.md` — prior wave findings (if generated)
- `context/{wave}/survey_context.md` — question codes + message list (if generated)
- `context/{wave}/project_context.md` — study design + field intel + wave hypotheses

Show a summary table with file name, section count, and key stats per file. Ask user to review and confirm or edit before Stage 2. This is the only content gate before hypotheses.

### Stage 2 — Generate Hypothesis Bank (`/hypotheses`)
Reads five context files. Required: `market_context.md`, `project_context.md`, `KBQs.md`. Strongly recommended: `prior_wave_context.md`, `survey_context.md`.

| File | Provides | If missing |
|------|---------|-----------|
| `context/market_context.md` | Competitive/clinical landscape | Stop |
| `context/{wave}/project_context.md` | Study design + field intel + wave expectations | Stop |
| `input/wave/{wave}/KBQs.md` | Organising structure — all KBQ domains | Stop |
| `context/{wave}/prior_wave_context.md` | Domain-level metrics, recs, gaps from prior delivered report | Proceed, reduced coverage |
| `context/{wave}/survey_context.md` | Question codes + message list ("Test with:" lines) | Proceed, no codes |

`prior_wave_context.md` drives a mandatory hypothesis type: **PRIOR WAVE VALIDATION** — for every domain finding, every carried-forward recommendation, and every unanswered question from the prior report. These close the loop between what was advised and what is being measured.

Produces `context/{wave}/hypothesis_bank.md`.
**→ Pause:** Show hypothesis count by type (new, prior wave validation, action item, methodology artifact), domain breakdown. Confirm before Stage 3.

### Stage 3 — Validate Data + Write Insights & ES (`/sfea-insight-writer`)
Reads `hypothesis_bank.md`, `project_context.md`, and `source_data.json`. Runs in four sub-phases:
- **Phase 0 (auto):** Validates every hypothesis against actual survey data — extracts prior/current/delta per question code, produces `context/{wave}/validated_analysis.md`
- **Phase 1:** Writes data-grounded talking headlines per domain → `context/{wave}/slide_headlines.md` ✋ User confirms
- **Phase 2:** Writes Executive Summary in selected format → `context/{wave}/exec_summary.md` ✋ User confirms
- **Phase 3:** Writes numbered Recommendations → appended to exec_summary.md

### Stage 4 — Build Slide Plan (`/slide-plan`)
Reads `validated_analysis.md`, `slide_headlines.md`, `exec_summary.md`, `hypothesis_bank.md`, `kbqs.md`, and `survey_context.md`. Clusters hypotheses into slides with chart specs, synthesized per-slide insights (drawn from validated analysis), and matched headlines. Always includes Cover (Slide 1), ES (Slide 2), and Recs (Slide 3) before data slides. Produces `context/{wave}/slide_plan.md`.
**→ Pause:** Show slide count, section breakdown, slide titles, and ES/Recs placement. Ask user to confirm before Stage 5.

### Stage 5 — Generate config.yaml (internal — no user gate)
Maps the slide plan to YAML extractions and asks. **Must search `source_data.json` → `_sheets` for every question code in the slide plan** — never copy from existing configs. Each slide becomes an `ask` entry; each data source becomes an `extraction` entry. If a code is not found, flags it explicitly. Proceeds automatically to Stage 6.

### Stage 6 — Run `generate_deck()`
Executes the pipeline: loads config → extracts data (or reads JSON cache) → renders all slides → saves `output/{wave}/deck.pptx`. If `source_data.json` has only the index (from Stage 0), automatically extracts from Excel and saves the full JSON.
**→ Report:** Show success/failure, slide count, output path, and any gaps.

```
INPUT FOLDER: input/wave/{wave}/
  source_data.xlsx          → Stage 0 only (auto)
  *prior wave files*        → Stage 0.5b only
  *survey draft*            → Stage 0.5c only
  KBQs.md                   → Stage 1 + Stage 2 (hand-written, no auto-gen)
  call_notes, odt, misc     → Stage 1 only

─────────────────────────────────────────────────── no gates ──
Stage 0    index_excel()          →  context/{wave}/source_data.json

Stage 0.5a /market-context        →  context/market_context.md
           (skip if already exists;   NOT wave-versioned)
           (one prompt: flagged claims only)

Stage 0.5b /prior-wave-context    →  context/{wave}/prior_wave_context.md
           (skip if no prior files;   one prompt: file list)

Stage 0.5c /survey-context        →  context/{wave}/survey_context.md
           (skip if no survey draft;  one prompt: file list + code count)
─────────────────────────────────────────────────── no gates ──

Stage 1    /build-project-context
           reads: call_notes + odt + KBQs.md + misc input files
                + context/market_context.md
                + context/{wave}/prior_wave_context.md
           does NOT read: prior wave pptx, survey draft, source_data
           writes: context/{wave}/project_context.md

✋ SINGLE VALIDATION GATE — user reviews all 4 context files:
     context/market_context.md          (competitive/clinical)
     context/{wave}/prior_wave_context.md  (prior wave findings)
     context/{wave}/survey_context.md   (question codes + messages)
     context/{wave}/project_context.md  (study design + field intel)

Stage 2    /hypotheses
           reads: market_context.md        (required)
                + project_context.md       (required)
                + KBQs.md                  (required, hand-written)
                + prior_wave_context.md     (strongly recommended)
                + survey_context.md        (strongly recommended)
           hypothesis types generated:
                New · PRIOR WAVE VALIDATION · [ACTION ITEM] · METHODOLOGY ARTIFACT
           writes: context/{wave}/hypothesis_bank.md
    ✋ User confirms count by type + domain coverage

Stage 3    /sfea-insight-writer   →  validated_analysis.md
                                  →  slide_headlines.md      ✋
                                  →  exec_summary.md         ✋

Stage 4    /slide-plan            →  context/{wave}/slide_plan.md
    ✋ User confirms slide count + section breakdown

Stage 5    config.yaml            ← internal, no gate
Stage 6    generate_deck()        →  output/{wave}/deck.pptx
```
    ✋ User confirms
    ↓ Stage 3: /sfea-insight-writer
        Phase 0 (auto): validate vs source_data.json
context/{wave}/validated_analysis.md
        Phase 1: headlines        ✋ User confirms
context/{wave}/slide_headlines.md
        Phase 2–3: ES + Recs      ✋ User confirms
context/{wave}/exec_summary.md
    ↓ Stage 4: /slide-plan
context/{wave}/slide_plan.md      (includes per-slide insights + ES/Recs slides)
    ✋ User confirms
    ↓ Stage 5: config.yaml                    ← internal, searches _sheets
    ↓ Stage 6: generate_deck()
output/{wave}/deck.pptx
```

### Manual / Incremental Workflow

For projects with an existing config, or when making targeted changes:

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
