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
│       ├── data/              # Source Excel files (gitignored)
│       ├── templates/         # Template decks (gitignored)
│       ├── output/            # Generated deliverables (gitignored)
│       └── reference/         # Client briefs & asks (gitignored)
├── slidegen/                  # SlideGen system
│   ├── pipeline/              # ★ Generic deck generation pipeline
│   │   ├── __init__.py        # Exports: generate_deck()
│   │   ├── project_config.py  # ProjectConfig dataclasses + YAML loader
│   │   ├── data_loaders.py    # 5 generic data extractors + load_all_data()
│   │   ├── slide_renderers.py # 9 slide type renderers (RENDERERS registry)
│   │   └── orchestrator.py    # Pipeline entry: config → data → renderers → PPTX
│   ├── __init__.py            # Package exports: SlideBuilder, LiveEditor, reconcile
│   ├── __main__.py            # CLI: python -m slidegen <create|edit|reconcile>
│   ├── config.py              # Centralized paths and settings
│   ├── pptx_utils.py          # Utility library (50+ functions, lxml/COM/registry)
│   ├── create.py              # SlideBuilder class — python-pptx creation + registry
│   ├── edit.py                # LiveEditor class — win32com live editing + edit log
│   ├── reconcile.py           # Registry reconciliation from live PowerPoint state
│   ├── slide_registry.json    # Shape state (created at runtime, gitignored)
│   └── output/                # Generated slides (gitignored)
├── src/                       # Legacy pipeline (original POC, hardcoded J&J)
│   ├── generate_asks.py       # Monolithic 14-slide J&J deck generator
│   ├── extract_data.py        # Excel → output/slide_data_v2.pkl
│   ├── generate_pptx.py       # pkl → full multi-slide deck
│   └── validate_data.py       # Cross-check pkl against raw Excel
├── scripts/                   # Ad-hoc exploration & discovery
│   └── discover_data.py       # Explores Excel sheets, outputs column_mapping.csv
├── archive/                   # Superseded scripts & old artifacts
├── docs/                      # Design docs & architecture notes
├── .claude/skills/            # Claude Code skills (auto-discovered)
│   ├── pptx/                  # PPTX read/create/edit skill
│   ├── jj-slide-style/        # J&J brand styling skill
│   └── pptx-utils/            # python-pptx helper utilities
├── column_mapping.csv
└── .gitignore
```

## Running

```bash
# ── Generic Pipeline (YAML-driven) ──
python -m slidegen.pipeline.orchestrator projects/jnj_rybrevant/config.yaml
# Or from Python:
#   from slidegen.pipeline import generate_deck
#   generate_deck("projects/jnj_rybrevant/config.yaml")

# ── Legacy Pipeline (hardcoded J&J) ──
python src/generate_asks.py        # Monolithic 14-slide deck

# ── SlideGen System ──
python -m slidegen create                   # Demo slide creation
python -m slidegen edit <filename.pptx>     # Interactive live editor
python -m slidegen reconcile <filename.pptx> # Sync registry from PowerPoint

# ── Data discovery ──
python scripts/discover_data.py
```

## Pipeline Architecture

```
projects/jnj_rybrevant/config.yaml  →  ProjectConfig (dataclasses)
                                      ↓
Excel file  →  data_loaders.load_all_data()  →  dict[extraction_id → list[dict]]
                                      ↓
orchestrator  →  RENDERERS[slide_type](slide, config, ask, data)
                                      ↓
                                 output.pptx
```

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

- `pandas`, `openpyxl` (Excel reading)
- `python-pptx` (PowerPoint generation)
- `lxml` (XML manipulation for native PPT charts)
- `pyyaml` (YAML config loading)
- `pywin32` (win32com — live PowerPoint editing via COM, Windows only)

## SlideGen API

```python
# ── Generic Pipeline ──
from slidegen.pipeline import generate_deck
generate_deck("projects/jnj_rybrevant/config.yaml")

# ── Create a slide manually ──
from slidegen.create import SlideBuilder
builder = SlideBuilder(template="path/to/template.pptx")
slide = builder.add_blank_slide()
builder.add_clustered_bar(slide, categories, series_data, ...)
builder.save("output.pptx")

# ── Edit live (file must be open in PowerPoint) ──
from slidegen.edit import LiveEditor
with LiveEditor("output.pptx") as editor:
    editor.set_text("zrx_001", "New text", color="red")
    editor.move("zrx_002", left=5.0, top=2.0)

# ── Reconcile before editing ──
from slidegen.reconcile import reconcile
report = reconcile("output.pptx")
```
