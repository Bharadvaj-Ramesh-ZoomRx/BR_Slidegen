# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Pharma consulting report generator for **Rybrevant (RYB) + Lazcluze** vs **Tagrisso (TAG)** — builds PowerPoint slide decks from survey data in `Lung SFEA SB.xlsx`. The deck covers message recall, effectiveness, rep performance, call-to-action metrics, prescription intent, and high-impact interactions across Q3/Q4 2025.

## Folder Structure

```
├── src/                    # Active pipeline scripts
│   ├── extract_data.py     # Step 1: Excel → output/slide_data_v2.pkl
│   ├── generate_pptx.py    # Step 2: pkl → full multi-slide deck
│   ├── generate_slide1.py  # Step 2b: pkl + template → polished Slide 1
│   └── validate_data.py    # Step 3: cross-check pkl against raw Excel
├── scripts/                # Ad-hoc exploration & discovery
│   ├── discover_data.py    # Explores Excel sheets, outputs column_mapping.csv
│   ├── explore.py          # Quick data inspection
│   └── explore2.py         # Quick data inspection
├── archive/                # Superseded v1 scripts
│   ├── extract_data_v1.py
│   └── generate_pptx_v1.py
├── output/                 # Generated artifacts (gitignored)
│   ├── slide_data_v2.pkl
│   ├── Rybrevant_Analysis_Deck_v2.pptx
│   ├── Slide1_MR_ME_Final_v5.pptx
│   └── validation_report_v2.txt
├── docs/
│   └── session_trace_2026-02-24.md
├── column_mapping.csv
└── .gitignore
```

All scripts resolve paths relative to the project root via `os.path.dirname(os.path.dirname(os.path.abspath(__file__)))`.

## Running

```bash
# Full pipeline (extract → generate → validate)
python src/extract_data.py
python src/generate_pptx.py
python src/validate_data.py

# Single polished slide (requires template PPTX in project root)
python src/generate_slide1.py

# Data discovery
python scripts/discover_data.py
```

## Data Source Layout

All scripts read `Lung SFEA SB.xlsx` (placed in project root, gitignored) with `header=None` (0-indexed rows/cols):

- **RYB sheet**: col 0=code, col 1=desc, col 7=Q3 Total, col 17=Q4 Total
- **TAG sheet**: col 0=code, col 1=desc, col 7=Q3 Total, col 13=Q4 Total
- **Additonal Analysis sheet** (note: misspelled in source): col 1=metric, col 2=RYB_Q3, col 3=RYB_Q4, col 4=TAG_Q3, col 5=TAG_Q4; message rows 30-39 use cols 3-8 for MR/ME/Believable

## Key Conventions

- **Percentage conversion**: Raw Excel values are decimals (0.0–1.0). `pct()` multiplies by 100 and rounds to 1 decimal. Some AA rows (e.g., row 47) store whole-number percentages — use `sf()` not `pct()` for those.
- **Delta**: Always Q4 minus Q3 in percentage points.
- **Sorting**: Slides sort data by Q4 MR descending (highest recall at top).
- **Brand colors**: RYB orange (`#E8692A` Q4, `#F4B98A` Q3), TAG violet (`#7B4FA6` Q4, `#B998D5` Q3). Delta colors: green positive, red negative.
- **Slide numbering**: Slides 3–15 each map to a key in the pickle dict (e.g., `s3_ryb_messaging`, `s5_rep_perf`).

## Dependencies

- `pandas`, `openpyxl` (Excel reading)
- `python-pptx` (PowerPoint generation)
- `matplotlib`, `numpy` (chart rendering in `generate_pptx.py`)
- `lxml` (XML manipulation for native PPT charts in `generate_slide1.py`)
