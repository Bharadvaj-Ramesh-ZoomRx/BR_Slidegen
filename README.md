# R3M Report — Pharma Promotional Effectiveness Tracking

Generic slide deck generator for pharma consulting reports. Builds PowerPoint decks from survey data using a **YAML-driven pipeline**. All interaction happens through natural language in the **Claude Code terminal** — no manual config writing or code changes needed.

Currently configured for **Rybrevant (RYB) + Lazcluze** vs **Tagrisso (TAG)** (Q3/Q4 2025).

## Getting Started

### Prerequisites

```bash
pip install pandas openpyxl python-pptx lxml pyyaml pywin32
```

### Create a New Project

1. Create the project folder: `projects/your_project/`
2. Place source Excel in `projects/your_project/data/{wave}/source_data.xlsx`
3. Place client asks/brief in `projects/your_project/reference/{wave}/` (markdown or docx)
4. Optionally place a template deck in `projects/your_project/templates/template.pptx`
5. In the Claude Code terminal, say: **`Create slides for projects/your_project`**

Claude Code auto-discovers the Excel structure, reads the reference docs, generates `config.yaml`, and builds the deck.

### Run an Existing Project

```bash
# Generate from an existing config
python -m slidegen.pipeline.orchestrator projects/jnj_rybrevant/config.yaml

# Or from Python
from slidegen.pipeline import generate_deck
generate_deck("projects/jnj_rybrevant/config.yaml")

# Regenerate a single slide (0-based index)
from slidegen.pipeline import regenerate_slide
regenerate_slide("projects/jnj_rybrevant/config.yaml", slide_index=4)
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

```
projects/{name}/
  data/{wave}/source_data.xlsx     # Input: survey data (Excel)
  reference/{wave}/asks.md         # Input: client asks
  templates/template.pptx          # Input: slide template (optional)
        ↓
slidegen/pipeline/
  config_generator.py              # Auto-discover Excel → scaffold config.yaml
  project_config.py                # Load YAML → ProjectConfig dataclasses
  data_loaders.py                  # Extract data blocks from Excel (5 methods)
  slide_renderers.py               # Render slides (9 types)
  orchestrator.py                  # Tie config + data + renderers → PPTX
        ↓
projects/{name}/
  config.yaml                      # Generated config (brands, sheets, extractions, asks)
  output/{wave}/deck.pptx          # Generated deck
  output/{wave}/slide_data.json    # Data cache (auto-invalidated)
  output/{wave}/shape_registry.json # Shape state + data lineage
  output/{wave}/backups/           # PPTX backups before edits
```

### Data Extraction Methods

| Method | Use Case |
|--------|----------|
| `question_code` | Find a code row, walk sub-rows, extract prior/current values |
| `multi_question_code` | One row per code (e.g. CTA: compelling, changed opinion, closing) |
| `row_range` | Fixed row range with column mapping |
| `question_code_multi_col` | Multiple columns per row (e.g. HII: hi vs other) |
| `nested_ordinal` | Grouped ordinal sub-rows (e.g. 1st/2nd/3rd recall order) |

### Slide Types

9 reusable slide types, each driven by YAML config:

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

## Wave Versioning

Data and output are versioned by wave (e.g. `PET_Q3Q4_2025`, `PET_Q1Q2_2026`). Templates are shared across waves.

```
projects/jnj_rybrevant/
  config.yaml
  templates/template.pptx              # shared across waves
  data/
    PET_Q3Q4_2025/source_data.xlsx     # wave-versioned input
    PET_Q1Q2_2026/source_data.xlsx
  output/
    PET_Q3Q4_2025/                     # wave-versioned output
      deck.pptx
      slide_data.json
      shape_registry.json
      backups/
    PET_Q1Q2_2026/deck.pptx
  reference/
    PET_Q3Q4_2025/asks.md              # wave-versioned asks
```

In `config.yaml`, `{{wave}}` in paths is interpolated from `project.wave`:
```yaml
project:
  wave: "PET_Q3Q4_2025"
data_source_path: "data/{{wave}}/source_data.xlsx"
template_path: "templates/template.pptx"           # no {{wave}} — shared
output_path: "output/{{wave}}/deck.pptx"
```

To start a new wave: update `project.wave`, `period_current`, `period_prior` in config and place new data in `data/{new_wave}/`. Old wave output is preserved.

## Data Setup

Source Excel and client assets are **gitignored** (proprietary survey data).

### Standard File Names

| Folder | File | Purpose |
|--------|------|---------|
| `data/` | `source_data.xlsx` | Survey data (Excel) |
| `data/` | `client_context.xlsx` | Market context (optional) |
| `templates/` | `template.pptx` | Slide template |
| `output/` | `deck.pptx` | Generated deck |
| `reference/` | `asks.md` | Client asks (markdown) |
| `reference/` | `asks_brief.docx` | Client asks (original doc) |

### Excel Structure (J&J Rybrevant)

All scripts read with `header=None` (0-indexed rows/cols):

| Sheet | Key Columns |
|-------|-------------|
| **RYB** | col 0 = question code, col 1 = description, col 7 = Q3 Total, col 17 = Q4 Total |
| **TAG** | col 0 = question code, col 1 = description, col 7 = Q3 Total, col 13 = Q4 Total |
| **Additonal Analysis** | col 1 = metric, cols 2–5 = RYB/TAG Q3/Q4; rows 30–39 = message M/B/D |

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
    config.yaml                # Auto-generated YAML config
    data/                      # Source Excel files (gitignored)
    templates/                 # Template decks (gitignored)
    output/                    # Generated deliverables (gitignored)
    reference/                 # Client briefs & asks (gitignored)

slidegen/                     # SlideGen system
  pipeline/                   # Generic deck generation pipeline
    config_generator.py        # Data discovery + config scaffolding
    project_config.py          # ProjectConfig schema + YAML loader
    data_loaders.py            # Generic data extractors (5 methods)
    slide_renderers.py         # 9 slide type renderers
    orchestrator.py            # Pipeline entry + per-slide regen + PPTX backup
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
| `pandas` | DataFrame operations |
| `openpyxl` | Excel file reading |
| `python-pptx` | PowerPoint generation |
| `lxml` | XML manipulation for chart formatting |
| `pyyaml` | YAML config loading |
| `pywin32` | COM automation for live editing (Windows only) |
