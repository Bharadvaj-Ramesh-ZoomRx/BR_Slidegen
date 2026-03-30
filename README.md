# R3M Report — Pharma Promotional Effectiveness Tracking

Generic slide deck generator for pharma consulting reports. Builds PowerPoint decks from survey data using a **YAML-driven pipeline**. All interaction happens through natural language in the **Claude Code terminal** — no manual config writing or code changes needed.

Currently configured for **Rybrevant (RYB) + Lazcluze** vs **Tagrisso (TAG)** (Q3/Q4 2025).

## Getting Started

### Prerequisites

```bash
pip install pandas openpyxl python-pptx lxml pyyaml requests pywin32
```

### Create a New Project

1. Create the project folder: `projects/your_project/`
2. Place source Excel in `projects/your_project/input/wave/{wave}/source_data.xlsx`
3. Place reference docs in the same wave folder:
   - `market_context.md` — curated market, disease, product, competitive context
   - `kbqs.md` — standing key business questions
   - `survey_context.md` — survey instrument, question codes, response scales
   - `call_notes.docx` — client call notes (optional)
   - `pet_project_kbq.odt` — study design + KBQs (optional)
   - `prior_wave_es.md` — prior wave executive summary findings (optional)
4. Optionally place a template deck in `projects/your_project/templates/template.pptx`
5. In the Claude Code terminal, say: **`Create slides for projects/your_project`**

Claude Code runs a 6-stage pipeline: indexes Excel → builds project context → generates hypotheses → creates slide plan → writes config.yaml → builds the deck.

### Run an Existing Project

```bash
# Generate from an existing config
python -m slidegen.pipeline.orchestrator projects/jnj_rybrevant/config.yaml

# Or from Python
from slidegen.pipeline import generate_deck
generate_deck("projects/jnj_rybrevant/config.yaml")

# Regenerate a single slide (by 0-based index or ask_id string)
from slidegen.pipeline import regenerate_slide
regenerate_slide("projects/jnj_rybrevant/config.yaml", slide_index=4)
regenerate_slide("projects/jnj_rybrevant/config.yaml", slide_index="ryb_mr")  # by ask_id
```

Output: `projects/jnj_rybrevant/output/PET_Q3Q4_2025/deck.pptx`

## Workflows

All workflows are triggered via natural language in the Claude Code terminal:

| Scenario | User Says | What Happens |
|----------|-----------|--------------|
| Brand new project | `Create slides for projects/{name}` | Discover data, generate config, build deck |
| Edit 1 slide | `Edit Slide N — ...` | `regenerate_slide()` — single slide regen |
| New wave, same asks | `Edit slides with new wave data — PET_Q1Q2_2026` | Update config wave → `generate_deck()` |
| New wave + new asks | `Edit slides with new wave data + new asks — PET_Q1Q2_2026` | Update config + extractions → `generate_deck()` |
| Add/remove/reorder | `Add slide after N...` / `Remove Slide N` | Modify asks → `generate_deck()` |

### Examples

```
# ── Create ──
Create slides for projects/jnj_rybrevant
Create slides for projects/pfizer_ibrance wave PET_Q1_2026

# ── Edit single slide ──
Edit Slide 5 — change headline to "Updated Message Recall"
Edit Slide 3 — sort bars descending by current value
Edit Slide 9 — use data from Q2_15Z instead of Q2_10Z

# ── New wave ──
Edit slides with new wave data — PET_Q1Q2_2026, Q1'26 vs Q2'26
Edit slides with new wave data + new asks — PET_Q1Q2_2026

# ── Add / remove / reorder ──
Add a slide after Slide 6 — clustered_compare for HCP satisfaction
Remove Slide 8
Move Slide 10 before Slide 5
Regenerate all slides
```

## Architecture

### Full Create Pipeline (Stage 0–5)

```
input/wave/{wave}/                       # User drop zone
  source_data.xlsx                       #   Survey data
  market_context.md, kbqs.md, ...        #   Reference docs
        ↓
Stage 0: index_excel()                   # Excel → JSON index (automatic)
        ↓
context/{wave}/source_data.json          #   _sheets: row-level code+desc index
        ↓
Stage 1: /build-project-context          # Synthesize → project_context.md
    ✋ User confirms
Stage 2: /hypotheses                     # Generate → hypothesis_bank.md (can reference _sheets)
    ✋ User confirms
Stage 3: /slide-plan                     # Plan → slide_plan.md (can reference _sheets)
    ✋ User confirms
Stage 4: config.yaml                     # Map slides to extractions (searches _sheets)
Stage 5: generate_deck()                 # Build → deck.pptx
        ↓
context/{wave}/                          # System-generated intermediates
  project_context.md                     #   Stage 1 output
  hypothesis_bank.md                     #   Stage 2 output
  slide_plan.md                          #   Stage 3 output
  source_data.json                       #   Stage 0 index + Stage 5 extracted data
output/{wave}/
  deck.pptx                             #   Generated deck
  shape_registry.json                    #   Shape state + data lineage
  backups/                               #   PPTX backups before edits
```

### Slide Generation Pipeline

```
projects/{name}/config.yaml  →  ProjectConfig  →  validate()  →  errors or proceed
                                      ↓
Track A (JSON-first):  Synapse API  →  fetch_data_as_json()  →  source_data.json
Track B (Excel):       source_data.xlsx  →  data_loaders.load_all_data()  →  dict[extraction_id → list[dict]]
                                      ↓
orchestrator  →  RENDERERS[slide_type](slide, config, ask, data, namer)  →  deck.pptx
```

### Data Loading (Auto-JSON + Two-Track Fetching)

**Two tracks for getting data:**
- **Track A — JSON-First**: For `synapse_report` extractions, calls Synapse `/reports/generate` API directly. Bypasses Excel entirely. **Only activates when `SYNAPSE_API_KEY` env var is set.** If absent, these extractions are skipped.
- **Track B — Excel** (default): Uses local `source_data.xlsx` or downloads via banner plan API. **If no Synapse API key is present, the pipeline proceeds entirely with the provided Excel file.** Supports non-blocking split: `trigger_generation()` returns immediately, `wait_and_download()` blocks later.

On first run, data is extracted and saved as `context/{wave}/source_data.json`. Subsequent runs read the JSON directly — no pandas, no column indices. The JSON auto-invalidates when the Excel file changes (hash check) **or** when extraction params change (extraction hash check). Delete `source_data.json` to force re-extraction.

**Config validation**: `config.validate()` catches mismatched data_keys, unknown slide_types, missing params, and invalid extraction methods before deck generation begins.

### Data Extraction Methods

| Method | Use Case |
|--------|----------|
| `question_code` | Find a code row, walk sub-rows, extract prior/current values |
| `multi_question_code` | One row per code (e.g. CTA: compelling, changed opinion, closing) |
| `row_range` | Fixed row range with column mapping |
| `question_code_multi_col` | Multiple columns per row (e.g. HII: hi vs other) |
| `nested_ordinal` | Grouped ordinal sub-rows (e.g. 1st/2nd/3rd recall order) |
| `synapse_report` | JSON-first: calls Synapse `/reports/generate` API directly (bypasses Excel) |

### Slide Types

20 reusable slide types, each driven by YAML config:

| Slide Type | Description |
|------------|-------------|
| `cover` | Title/cover slide |
| `executive_summary` | Bullet-list insights |
| `single_bar_with_delta` | Horizontal bar + QoQ delta column |
| `dual_bar_with_delta` | Two side-by-side bars + deltas (e.g. MR + ME) |
| `dual_bar_qoq` | Two side-by-side Q4-vs-Q3 clustered bars + deltas |
| `clustered_compare` | Clustered bar comparing two groups + delta/gap columns |
| `dual_bar_compare` | Side-by-side dual brand bar comparison + insight callout |
| `qoq_bar_with_delta` | Q4 vs Q3 clustered bar + delta column |
| `two_section_bar` | Two vertically stacked bar sections (e.g. RYB Rx + TAG Rx) |
| `stacked_order` | Stacked bar with ordinal breakdown + total column |
| `abacus` | XY scatter abacus with label/value tables + delta column |
| `dual_abacus` | Two side-by-side abacus panels (e.g. Acad vs Comm by brand) |
| `followup_rep` | Follow-up rep abacus with dual delta columns |
| `hii_scorecard` | Multi-section clustered column chart with section headers |
| `dual_doughnut` | Side-by-side doughnut pairs comparing patient segments |
| `message_mbd` | Multi-column abacus for Motivation/Believability/Differentiation |
| `trended_scorecard` | Multi-panel mini line chart grid (QoQ trend scorecard) |
| `trended_activity` | Side-by-side line + stacked column panels (reach/SOV/frequency) |
| `quadrant_scatter` | 2×2 quadrant scatter chart (stated vs derived importance) |
| `heatmap_table` | Heatmap table with green gradient fills + QoQ delta columns |

## Wave Versioning

Data, context, and output are versioned by wave (e.g. `PET_Q3Q4_2025`, `PET_Q1Q2_2026`). Templates are shared across waves.

```
projects/jnj_rybrevant/
  config.yaml
  templates/template.pptx              # shared across waves
  input/wave/
    PET_Q3Q4_2025/source_data.xlsx     # wave-versioned input
    PET_Q1Q2_2026/source_data.xlsx
  context/
    PET_Q3Q4_2025/                     # wave-versioned system-generated files
      source_data.json                 #   auto-extracted from Excel
      hypothesis_bank.md               #   analysis intermediates
      slide_plan.md
  output/
    PET_Q3Q4_2025/                     # wave-versioned output
      deck.pptx
      shape_registry.json
      backups/
    PET_Q1Q2_2026/deck.pptx
```

In `config.yaml`, `{{wave}}` in paths is interpolated from `project.wave`:
```yaml
project:
  wave: "PET_Q3Q4_2025"
data_source_path: "input/wave/{{wave}}/source_data.xlsx"
template_path: "templates/template.pptx"           # no {{wave}} — shared
context_path: "context/{{wave}}/"
output_path: "output/{{wave}}/deck.pptx"
```

To start a new wave: update `project.wave`, `period_current`, `period_prior` in config and place new data in `input/wave/{new_wave}/`. Old wave output is preserved.

## Project Structure

```
projects/                     # Project folders (one per product/brand)
  jnj_rybrevant/              # J&J Rybrevant PET Q4'25
    config.yaml                # Auto-generated YAML config
    input/wave/                # Source files — user drop zone (gitignored)
    context/                   # System-generated intermediates (gitignored)
    templates/                 # Template decks (gitignored)
    output/                    # Generated deliverables (gitignored)

slidegen/                     # SlideGen system
  pipeline/                   # Generic deck generation pipeline
    config_generator.py        # Data discovery + config scaffolding + scaffold_config_from_plan()
    project_config.py          # ProjectConfig schema + YAML loader + validate()
    data_loaders.py            # 8 data extractors + JSON auto-cache + _codes rich index
    raw_data_loader.py         # Respondent-level parser + quarter cache
    synapse_fetcher.py         # Synapse API: trigger_generation() + wait_and_download()
    synapse_json_loader.py     # JSON-first: /reports/generate → slidegen format
    slide_renderers/           # 20 slide type renderers (package)
    orchestrator.py            # Pipeline entry + per-slide regen (by index or ask_id) + PPTX backup
  pptx_utils/                 # Utility package (8 modules)
    brand.py                   # BRAND{} dict, colors, fonts, constants
    lxml_helpers.py            # 20 lxml XML chart/axis helpers
    shapes.py                  # 7 shape primitives
    layout.py                  # LAYOUTS{} dict + slide chrome functions
    charts.py                  # CHART_PATTERNS{} + chart builders
    tables.py                  # Delta/value table builders
    com.py                     # COM helpers for live editing
    registry.py                # Registry CRUD operations
  create.py                   # SlideBuilder class
  edit.py                     # LiveEditor class (win32com)
  reconcile.py                # Registry reconciliation

archive/                      # Superseded scripts & old artifacts
docs/                         # Design docs & architecture notes
```

## Key Conventions

- **Percentages**: Raw Excel values are decimals (0.0–1.0). `pct()` converts to percentage. Some rows store whole numbers — use `straight()` for those.
- **Delta**: Always current minus prior in percentage points.
- **Brand colors**: Defined per-project in YAML. J&J: RYB orange (`#F75824`), TAG violet (`#7030A0`).
- **Fonts**: Defined per-project in YAML. J&J: Johnson Display (headers), Johnson Text (body).
- **Shape naming**: Pipeline assigns `zrx_{slide:03d}_{shape:03d}` names for COM targeting and registry tracking.

## Dependencies

| Package | Purpose |
|---------|---------|
| `pandas` | DataFrame operations (Excel extraction only — first run per wave) |
| `openpyxl` | Excel file reading |
| `python-pptx` | PowerPoint generation |
| `lxml` | XML manipulation for chart formatting |
| `pyyaml` | YAML config loading |
| `requests` | Synapse API integration (JSON-first data fetching + banner plan download) |
| `pywin32` | COM automation for live editing (Windows only) |
