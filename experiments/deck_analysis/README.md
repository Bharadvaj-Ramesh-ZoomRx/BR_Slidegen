# Deck Analysis — Reverse-Engineering PET Decks

**Purpose:** Reverse-engineer 10-15 PET slide decks across clients to improve `pptx_utils` composition primitives and the skill library for SlideGen.

**Branch:** `vijay-slidegen`

---

## Workflow

1. **Drop decks** into `decks/` — gitignored, not committed
2. **Run parser** (to be built in `parser.py`) across all decks
3. **Review outputs** in `outputs/` — gitignored, not committed
4. **Iterate** on `pptx_utils` and skills based on findings
5. **When stable patterns emerge** — PR to `main`; decks stay in the experiments folder

---

## Folder Structure

```
experiments/deck_analysis/
├── README.md              # This file (committed)
├── parser.py              # Deck parser (committed — to be built)
├── decks/                 # ★ DROP DECKS HERE (gitignored)
│   ├── client_a_pet_q1.pptx
│   ├── client_b_pet_wave3.pptx
│   └── ...
└── outputs/               # Generated analysis (gitignored)
    ├── inventory.json     # Structured per-deck metadata
    ├── report.md          # Human-readable summary
    ├── chart_types.json   # Chart type frequency map
    ├── ooxml_props.json   # Non-python-pptx properties used
    ├── layout_patterns.json  # Recurring multi-element layouts
    └── coverage_gaps.md   # What current pptx_utils doesn't cover
```

---

## What the Parser Will Extract

**Per slide:**
- Layout (slide master reference)
- Chart types and their configurations
- Shape positions, sizes, colors, fonts
- Text content (titles, headlines, callouts)
- Tables and their structures
- OOXML properties beyond python-pptx coverage

**Per deck:**
- Slide order and section structure
- Master slide usage
- Theme colors and fonts
- Template layouts used
- Embedded media

**Cross-deck aggregation:**
- Slide type frequency (which renderers matter for MVP)
- OOXML property inventory (what `lxml_helpers.py` needs)
- Layout pattern catalog (candidates for `LAYOUTS{}`)
- Brand variation analysis (structure of `BRAND{}`)
- Gap analysis vs. current 22 renderers

---

## What This Feeds Into

1. **`pptx_utils` refactor** — make renderers general (not J&J-specific), add missing `lxml_helpers`, build out `CHART_PATTERNS{}` and `LAYOUTS{}`, populate `BRAND{}`
2. **Skill inventory updates** — if new slide types or workflows emerge, update PRD v0.1 → v0.2
3. **Prioritization** — which `pptx_utils` functions to build first, based on frequency in real decks

---

## Confidentiality

Decks contain client-confidential data. They stay in `decks/` (gitignored). Generated outputs in `outputs/` are also gitignored by default. Client names may appear in extracted metadata for internal analysis — not redacted in outputs.

When patterns are ready to move into `pptx_utils` or the PRD, brand-specific details get abstracted (e.g., hex colors become `BRAND["JJ"]["primary"]` constants, not inline client identifiers).
