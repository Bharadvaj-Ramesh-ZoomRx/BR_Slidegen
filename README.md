# R3M Report — Rybrevant Promotional Effectiveness Tracking

Pharma consulting report generator that builds PowerPoint slide decks from survey data. Compares **Rybrevant (RYB) + Lazcluze** vs **Tagrisso (TAG)** across message recall, effectiveness, rep performance, call-to-action metrics, prescription intent, and high-impact interactions (Q3/Q4 2025).

## Quick Start

```bash
# Install dependencies
pip install pandas openpyxl python-pptx lxml pywin32

# Place source data (not committed — see Data Setup below)
# Then generate the 14-slide ask-response deck:
python src/generate_asks.py
```

Output: `J_and_J_project/output/Rybrevant_Asks_Deck_v2.pptx`

## Data Setup

The source Excel file and client assets are **gitignored** because they contain proprietary survey data. To run the pipeline:

1. Place `Lung SFEA SB.xlsx` in `J_and_J_project/data/`
2. Place the template deck in `J_and_J_project/templates/`
3. Create `J_and_J_project/output/` (auto-created on first run)

### Excel Structure

All scripts read with `header=None` (0-indexed rows/cols):

| Sheet | Key Columns |
|-------|-------------|
| **RYB** | col 0 = question code, col 1 = description, col 7 = Q3 Total, col 17 = Q4 Total |
| **TAG** | col 0 = question code, col 1 = description, col 7 = Q3 Total, col 13 = Q4 Total |
| **Additonal Analysis** | col 1 = metric, cols 2–5 = RYB/TAG Q3/Q4; rows 30–39 = message M/B/D |

## Slide Deck Contents

| # | Ask | Content |
|---|-----|---------|
| 1 | — | Cover slide |
| 2 | — | Executive Summary |
| 3 | Ask 1 | RYB Message Recall & Effectiveness |
| 4 | Ask 2 | TAG Message Recall |
| 5 | Ask 3 | Rep Performance — J&J vs AZ (Abacus chart) |
| 6 | Ask 4 | Message Believability & Composite Effectiveness |
| 7 | Ask 5 | Call to Action — J&J vs AZ |
| 8 | Ask 7 | One J&J Vision — Follow-up Reps |
| 9 | Ask 8 | Prescription Intent by Patient Type |
| 10 | Ask 9 | Share of Interaction Time by Topic |
| 11 | Nebulous 2 | High Impact vs Non-High Impact Interactions |
| 12 | — | Overall Quality Metrics |
| 13 | Nebulous 4 | Mariposa 1st Discussed vs NOT — Rep Performance |
| 14 | Nebulous 5 | TAG Message Recall Order (stacked bar) |

## Project Structure

```
src/
  generate_asks.py        # Main deck generator (14 slides from Excel data)
  extract_data.py         # Excel → output/slide_data_v2.pkl
  generate_pptx.py        # pkl → full multi-slide deck (earlier POC)
  validate_data.py        # Cross-check pkl against raw Excel

slidegen/                 # SlideGen system — programmatic slide creation + live editing
  pptx_utils.py           # 50+ helper functions (charts, tables, styling, XML)
  create.py               # SlideBuilder class — python-pptx creation
  edit.py                 # LiveEditor class — win32com live editing
  reconcile.py            # Registry reconciliation from PowerPoint state

scripts/                  # Ad-hoc data exploration
  discover_data.py        # Explores Excel sheets, outputs column_mapping.csv

archive/                  # Superseded scripts
docs/                     # Design docs & architecture notes
J_and_J_project/          # Client data, templates, output (gitignored)
```

## SlideGen API

```python
from slidegen.create import SlideBuilder

builder = SlideBuilder(template="path/to/template.pptx")
slide = builder.add_blank_slide()
builder.add_clustered_bar(slide, categories, series_data, ...)
builder.add_delta_column(slide, deltas, ...)
builder.save("output.pptx")
```

```python
from slidegen.edit import LiveEditor

with LiveEditor("output.pptx") as editor:
    editor.set_text("zrx_001", "New text", color="red")
    editor.move("zrx_002", left=5.0, top=2.0)
```

## Key Conventions

- **Percentages**: Raw Excel values are decimals (0.0–1.0). `pct()` converts to percentage. Some rows store whole numbers — use `sf()` for those.
- **Delta**: Always Q4 minus Q3 in percentage points.
- **Brand colors**: RYB orange (`#F75824`), TAG violet (`#7030A0`). Delta: green = positive, red = negative.
- **Fonts**: Johnson Display (headers), Johnson Text (body). Falls back to Calibri.

## Dependencies

| Package | Purpose |
|---------|---------|
| `pandas` | DataFrame operations |
| `openpyxl` | Excel file reading |
| `python-pptx` | PowerPoint generation |
| `lxml` | XML manipulation for chart data labels |
| `pywin32` | COM automation for live editing & slide export (Windows only) |
