# R3M Report — Pharma Promotional Effectiveness Tracking

Builds PowerPoint slide decks from PET survey data using a YAML-driven pipeline. Say what you want in plain English in the Claude Code terminal — no manual config or code changes needed.

**Active project:** Rybrevant (RYB) + Lazcluze vs Tagrisso (TAG) | **Wave:** Q1 2026

---

## For Analysts

### Setup (one-time)

1. Clone this repo
2. Symlink the shared OneDrive `projects/` folder into your clone (see [analyst_setup.md](docs/analyst_setup.md))
3. Install dependencies:
   ```bash
   pip install pandas openpyxl python-pptx lxml pyyaml requests pywin32 python-dotenv
   ```

### Generate a Deck

Open the Claude Code terminal and say:

```
Create slides for projects/jnj_rybrevant
```

That's it. Claude runs the full pipeline and produces `output/{wave}/deck.pptx`.

### Common Tasks

| What you want | What to say |
|---------------|-------------|
| Generate full deck | `Create slides for projects/{name}` |
| Regenerate with new data | `Regenerate with fresh data` |
| Edit one slide | `Edit Slide 5 — change headline to "..."` |
| New wave, same slides | `Edit slides with new wave data — Q2 2026` |
| Add a slide | `Add slide after Slide 6 — clustered_compare for HCP satisfaction` |
| Remove a slide | `Remove Slide 8` |
| Refresh all data | `Refresh this deck` |

### Input Files

Place these in `projects/{name}/input/wave/{wave}/`:

| File | Required? | What it is |
|------|-----------|-----------|
| `source_data.xlsx` | Yes | Aggregated survey data from Synapse |
| `KBQs.md` | Yes | Standing key business questions |
| `Call Notes.docx` | Recommended | Client call notes and action items |
| `[prior wave].pptx` | Recommended | Prior wave report deck |
| `[survey draft].docx` | Recommended | Survey instrument for this wave |
| `PET Project...odt` | Optional | Study design and methodology |

### Output

```
projects/{name}/output/{wave}/
  deck.pptx              # The generated deck
  shape_registry.json    # Shape metadata for live editing
  backups/               # Auto-saved before each edit
```

---

## How It Works

### Pipeline Stages

```
Stage 0    Index Excel              automatic    source_data.json
Stage 0.5  Build context            automatic    market_context.md
           (market, prior wave,                  prior_wave_context.md
            survey structure)                    survey_context.md

Stage 1    Project context          user review  project_context.md
Stage 2    Hypothesis bank          user review  hypothesis_bank.md
Stage 3    Data validation +        user review  validated_analysis.md
           headlines + ES + recs                 slide_headlines.md
                                                 exec_summary.md
Stage 4    Slide plan               user review  slide_plan.md
Stage 5    Config generation        automatic    config.yaml
Stage 6    Deck generation          automatic    deck.pptx
```

Stages marked "user review" pause for your confirmation before continuing.

### Wave Versioning

Each wave gets its own folder for input, context, and output. Templates are shared.

```
projects/jnj_rybrevant/
  config.yaml
  templates/template.pptx            # shared across waves
  input/wave/Q1 2026/                # wave input
  context/
    market_context.md                # shared (not wave-specific)
    Q1 2026/                         # wave context (9 generated files)
  output/Q1 2026/deck.pptx          # wave output
```

To start a new wave: update `project.wave` in config.yaml, place new Excel in the wave folder, and run.

---

## Slide Types

20 chart types available, each driven by YAML config:

| Type | Best for |
|------|----------|
| `single_bar_with_delta` | Ranked list with QoQ change (most common) |
| `abacus` | Attribute ratings — dot plot with value tables |
| `executive_summary` | Text-based insights or recommendations |
| `clustered_compare` | Two groups compared on same metric |
| `dual_bar_with_delta` | Two related metrics side-by-side (e.g. MR + ME) |
| `dual_bar_compare` | Two brands shown in separate chart panels |
| `qoq_bar_with_delta` | Current vs prior clustered bars |
| `two_section_bar` | Two stacked bar sections (e.g. RYB top, TAG bottom) |
| `stacked_order` | Ordinal breakdown (1st / 2nd / 3rd recall) |
| `message_mbd` | Motivation / Believability / Differentiation scatter |
| `hii_scorecard` | Multi-section metric scorecard |
| `dual_doughnut` | Side-by-side doughnut pairs |
| `trended_scorecard` | Mini line chart grid (5-quarter trend) |
| `trended_activity` | Reach / SOV / frequency trend panels |
| `quadrant_scatter` | Stated vs derived importance 2x2 |
| `heatmap_table` | Green gradient table with QoQ deltas |
| `cover` | Title slide |
| `dual_abacus` | Two abacus panels (e.g. Acad vs Comm) |
| `followup_rep` | Follow-up rep abacus with dual deltas |
| `dual_bar_qoq` | Two brands, each with QoQ clustered bars |

---

## Data Extraction Methods

| Method | When to use |
|--------|-------------|
| `question_code` | Standard: find a Q code, walk its sub-rows |
| `multi_question_code` | Multiple Q codes, one row each |
| `row_range` | Specific Excel rows by position |
| `question_code_multi_col` | One code, multiple value columns |
| `nested_ordinal` | Ordinal sub-rows (1st / 2nd / 3rd) |
| `mock` | Hardcoded test data |
| `synapse_report` | Synapse API JSON-first (bypasses Excel) |
| `synapse_raw` | Synapse raw respondent data + local aggregation |

---

## Project Structure

```
projects/                        gitignored — lives on shared OneDrive
  jnj_rybrevant/                 production (symlinked to SharePoint)
  jnj_rybrevant_session/         session variant

slidegen/                        the engine
  pipeline/
    orchestrator.py              entry point: generate_deck(), regenerate_slide()
    project_config.py            YAML schema + loader + validate()
    data_loaders.py              9 extractors + JSON cache + _codes index
    config_generator.py          auto-scaffold config from slide plan
    raw_data_loader.py           respondent-level parser
    synapse_auth.py              API token resolution (key, env var, Azure AD)
    synapse_fetcher.py           banner plan download
    synapse_json_loader.py       JSON-first data fetch
    synapse_raw_fetcher.py       raw survey data fetch
    slide_renderers/             20 renderers (one per chart type)
  pptx_utils/                   PowerPoint helpers (11 modules)
  create.py                     SlideBuilder class
  edit.py                       LiveEditor (COM-based live editing)
  reconcile.py                  registry sync from PowerPoint

.claude/skills/                  Claude Code skills (10 pipeline stages)
archive/                         legacy scripts
docs/                            setup guides + PRD
```

---

## Key Conventions

| Topic | Rule |
|-------|------|
| Percentages | Excel stores 0.0-1.0 decimals. `pct_mode: pct` multiplies by 100. `pct_mode: straight` for values already in %. Auto-detected from `_codes.value_range`. |
| Column layout | Never hardcode column positions. `index_excel()` auto-detects and stores in `_column_layouts`. |
| Deltas | Always current minus prior, in percentage points. |
| Colors | Per-project in YAML. J&J: RYB orange `#F75824`, TAG violet `#7030A0`. |
| Fonts | Per-project in YAML. J&J: Johnson Display (headers), Johnson Text (body). |
| Shape names | `zrx_{slide:03d}_{shape:03d}` for COM targeting and registry. |
| Templates | Prior wave reports are accepted as templates — the pipeline extracts the theme into a clean PPTX. |
| Long labels | Shortened via `label_shortcuts` in config. Max 65 chars (single bar), 55 (dual). |
| Renderer compatibility | Simple `{desc, prior, current}` data works with `single_bar_with_delta`, `abacus`, `executive_summary`. Complex renderers need specific field names. |

---

## CLI

```bash
python -m slidegen.pipeline.orchestrator projects/{name}/config.yaml           # generate deck
python -m slidegen.pipeline.orchestrator projects/{name}/config.yaml --fresh   # force re-extract
python -m slidegen create                                                      # demo slide
python -m slidegen edit <file.pptx>                                            # live editor
python -m slidegen reconcile <file.pptx>                                       # sync registry
python -m slidegen fetch-synapse projects/{name}/config.yaml                   # fetch banner plan
python -m slidegen fetch-raw projects/{name}/config.yaml                       # fetch raw data
```

## Dependencies

| Package | Purpose |
|---------|---------|
| `pandas` | Excel extraction (first run per wave, then cached) |
| `openpyxl` | Excel file reading |
| `python-pptx` | PowerPoint generation |
| `lxml` | XML chart formatting |
| `pyyaml` | YAML config loading |
| `requests` | Synapse API integration |
| `pywin32` | COM live editing (Windows only) |
| `python-dotenv` | `.env` file loading for API keys (optional) |
