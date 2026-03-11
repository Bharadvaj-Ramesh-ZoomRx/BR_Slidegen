# R3M Report — Pharma Promotional Effectiveness Tracking

Generic slide deck generator for pharma consulting reports. Builds PowerPoint decks from survey data using a **YAML-driven pipeline** — no code changes needed to add new projects.

Currently configured for **Rybrevant (RYB) + Lazcluze** vs **Tagrisso (TAG)** (Q3/Q4 2025).

## Quick Start

```bash
# Install dependencies
pip install pandas openpyxl python-pptx lxml pyyaml pywin32

# Generate a deck from a project config
python -m slidegen.pipeline.orchestrator projects/jnj_rybrevant/config.yaml

# Or from Python
from slidegen.pipeline import generate_deck
generate_deck("projects/jnj_rybrevant/config.yaml")

# Regenerate a single slide (0-based index)
from slidegen.pipeline import regenerate_slide
regenerate_slide("projects/jnj_rybrevant/config.yaml", slide_index=4)
```

Output: `projects/jnj_rybrevant/output/PET_Q3Q4_2025/deck.pptx`

### Legacy pipeline (archived)

```bash
python archive/src/generate_asks.py   # Hardcoded J&J deck (14 slides, POC)
```

## Architecture

```
projects/jnj_rybrevant/config.yaml       # YAML config: brands, sheets, extractions, asks
        ↓
slidegen/pipeline/
  project_config.py                # Load YAML → ProjectConfig dataclasses
  data_loaders.py                  # Extract data blocks from Excel
  slide_renderers.py               # 9 reusable slide type renderers
  orchestrator.py                  # Ties config + data + renderers → PPTX
        ↓
projects/jnj_rybrevant/output/deck.pptx  # Generated deck
```

### Adding a New Project

1. Create `projects/your_project/config.yaml` — define brands, colors, sheets, extractions, asks
2. Place source Excel as `projects/your_project/data/{wave}/source_data.xlsx`
3. Optionally place a template as `projects/your_project/templates/template.pptx`
4. Run: `python -m slidegen.pipeline.orchestrator projects/your_project/config.yaml`

No Python changes needed — all project-specific config lives in the YAML file.

### Wave Versioning

Data and output are versioned by wave (e.g. `PET_Q3Q4_2025`, `PET_Q1Q2_2026`). Templates are shared across waves.

```
projects/jnj_rybrevant/
  config.yaml
  templates/template.pptx              # shared across waves
  data/
    PET_Q3Q4_2025/source_data.xlsx     # wave-versioned input
    PET_Q1Q2_2026/source_data.xlsx
  output/
    PET_Q3Q4_2025/deck.pptx           # wave-versioned output
    PET_Q1Q2_2026/deck.pptx
  reference/
    PET_Q3Q4_2025/asks.md             # wave-versioned asks
```

In `config.yaml`, `{{wave}}` in paths is interpolated from `project.wave`:
```yaml
project:
  wave: "PET_Q3Q4_2025"
data_source_path: "data/{{wave}}/source_data.xlsx"
template_path: "templates/template.pptx"           # no {{wave}} — shared
output_path: "output/{{wave}}/deck.pptx"
```

## Workflows (Claude Code Terminal)

All workflows are triggered via natural language:

| Scenario | User Says | What Happens |
|----------|-----------|--------------|
| Brand new project | `Create slides for projects/{name}` | Full create workflow |
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

## Data Setup

Source Excel and client assets are **gitignored** (proprietary survey data). To run:

1. Place source Excel as `projects/jnj_rybrevant/data/PET_Q3Q4_2025/source_data.xlsx`
2. Place template deck as `projects/jnj_rybrevant/templates/template.pptx`
3. `projects/jnj_rybrevant/output/PET_Q3Q4_2025/` is auto-created on first run

Paths in `config.yaml` are relative to the project folder. Data and output use `{{wave}}` subfolders (e.g. `data/{{wave}}/source_data.xlsx`); templates are shared across waves.

### Standard File Names

| Folder | File | Purpose |
|--------|------|---------|
| `data/` | `source_data.xlsx` | Survey data (Excel) |
| `data/` | `client_context.xlsx` | Market context (optional) |
| `templates/` | `template.pptx` | Slide template |
| `output/` | `deck.pptx` | Generated deck |
| `reference/` | `asks.md` | Client asks (markdown) |
| `reference/` | `asks_brief.docx` | Client asks (original doc) |

### Excel Structure

All scripts read with `header=None` (0-indexed rows/cols):

| Sheet | Key Columns |
|-------|-------------|
| **RYB** | col 0 = question code, col 1 = description, col 7 = Q3 Total, col 17 = Q4 Total |
| **TAG** | col 0 = question code, col 1 = description, col 7 = Q3 Total, col 13 = Q4 Total |
| **Additonal Analysis** | col 1 = metric, cols 2–5 = RYB/TAG Q3/Q4; rows 30–39 = message M/B/D |

## Slide Types

The pipeline supports 9 reusable slide types, each driven by YAML config:

| Slide Type | Description |
|------------|-------------|
| `cover` | Title/cover slide |
| `executive_summary` | Bullet-list insights |
| `single_bar_with_delta` | Horizontal bar + QoQ delta column |
| `dual_bar_with_delta` | Two side-by-side bars + deltas (e.g. MR + ME) |
| `dual_bar_qoq` | Two side-by-side Q4-vs-Q3 clustered bars + deltas |
| `clustered_compare` | Clustered bar comparing two groups + delta/gap columns |
| `qoq_bar_with_delta` | Q4 vs Q3 clustered bar + delta column |
| `two_section_bar` | Two vertically stacked bar sections (e.g. RYB Rx + TAG Rx) |
| `stacked_order` | Stacked bar with ordinal breakdown + total column |

## J&J Rybrevant Deck Contents

| # | Slide Type | Content |
|---|------------|---------|
| 1 | cover | Cover slide |
| 2 | executive_summary | Executive Summary |
| 3 | dual_bar_with_delta | RYB Message Recall & Effectiveness |
| 4 | single_bar_with_delta | TAG Message Recall |
| 5 | clustered_compare | Rep Performance — J&J vs AZ |
| 6 | dual_bar_qoq | Message Believability & Composite Effectiveness |
| 7 | clustered_compare | Call to Action — J&J vs AZ |
| 8 | single_bar_with_delta | One J&J Vision — Follow-up Reps |
| 9 | two_section_bar | Prescription Intent by Patient Type |
| 10 | qoq_bar_with_delta | Share of Interaction Time by Topic |
| 11 | clustered_compare | High Impact vs Non-High Impact Interactions |
| 12 | single_bar_with_delta | Overall Quality Metrics |
| 13 | clustered_compare | Mariposa 1st Discussed vs NOT |
| 14 | stacked_order | TAG Message Recall Order |

## Project Structure

```
projects/                     # Project folders (one per product/brand)
  jnj_rybrevant/              # J&J Rybrevant PET Q4'25
    config.yaml                # YAML project config (brands, sheets, asks)
    data/                      # Source Excel files (gitignored)
    templates/                 # Template decks (gitignored)
    output/                    # Generated deliverables (gitignored)
    reference/                 # Client briefs & asks (gitignored)

slidegen/                     # SlideGen system
  pipeline/                   # Generic deck generation pipeline
    project_config.py          # ProjectConfig schema + YAML loader
    data_loaders.py            # Generic data extractors (5 methods)
    slide_renderers.py         # 9 slide type renderers
    orchestrator.py            # Pipeline entry + per-slide regen
    config_generator.py        # Data discovery + config scaffolding
  pptx_utils.py               # 50+ helper functions (charts, tables, XML)
  create.py                   # SlideBuilder class
  edit.py                     # LiveEditor class (win32com)
  reconcile.py                # Registry reconciliation

archive/                      # Superseded scripts & old artifacts
  src/                        # Legacy pipeline (original POC)
    generate_asks.py          # Hardcoded 14-slide J&J deck
    extract_data.py           # Excel → pkl
    generate_pptx.py          # pkl → deck
    validate_data.py          # Cross-check pkl vs Excel
  scripts/                    # Ad-hoc data exploration
    discover_data.py          # Excel sheet explorer
docs/                         # Design docs & architecture notes
```

## Key Conventions

- **Percentages**: Raw Excel values are decimals (0.0–1.0). `pct()` converts to percentage. Some rows store whole numbers — use `straight()` for those.
- **Delta**: Always current minus prior in percentage points.
- **Brand colors**: Defined per-project in YAML. J&J: RYB orange (`#F75824`), TAG violet (`#7030A0`).
- **Fonts**: Defined per-project in YAML. J&J: Johnson Display (headers), Johnson Text (body).

## Dependencies

| Package | Purpose |
|---------|---------|
| `pandas` | DataFrame operations |
| `openpyxl` | Excel file reading |
| `python-pptx` | PowerPoint generation |
| `lxml` | XML manipulation for chart formatting |
| `pyyaml` | YAML config loading |
| `pywin32` | COM automation for live editing (Windows only) |
