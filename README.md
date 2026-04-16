# SlideGen & R3M Report

**SlideGen** is ZoomRx Galen-Consulting's deck-generation engine: a library of composable Claude Code skills + a robust `pptx_utils` Python package that produces client-ready PowerPoint decks from survey data.

**R3M Report** is the first application on top of SlideGen — pharma Promotional Effectiveness Tracking (PET) decks, currently active for Rybrevant (RYB) + Lazcluze vs Tagrisso (TAG), Q1 2026 wave.

> 📄 **Strategic PRD (current):** [`docs/SlideGen_PRD_v1.1.md`](docs/SlideGen_PRD_v1.1.md)
> 📄 **Technical blueprint:** [`docs/SlideGen_PRD.md`](docs/SlideGen_PRD.md) (Siva, Mar 2026 — implementation details)
> 📄 **Analyst setup guide:** [`docs/analyst_setup.md`](docs/analyst_setup.md)

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

Open the Claude Code terminal in the repo and say:

```
Create slides for projects/jnj_rybrevant
```

Claude runs the pipeline end-to-end and produces `output/{wave}/deck.pptx`.

### Common Tasks (natural language)

| What you want | What to say |
|---|---|
| Generate a full deck | `Create slides for projects/{name}` |
| Regenerate with new data | `Regenerate with fresh data` |
| Edit one slide | `Edit Slide 5 — change headline to "..."` |
| New wave, same slides | `Refresh wave 4 of the Rybrevant deck` |
| Add a slide | `Add slide after Slide 6 — clustered_compare for HCP satisfaction` |
| Annotate an existing slide | `Add a callout to Slide 10 with this quote: "..."` |
| Remove / reorder slides | `Remove Slide 8` / `Move Slide 10 before Slide 3` |
| Executive summary | `Generate an ES answering these KBQs: ...` |
| Pre-delivery QA | `Audit this deck before delivery` |
| Refresh data across deck | `Refresh this deck with the Q1 '26 data` |

Every command maps to one of 8 MECE workflows (see [§Workflows](#workflows)).

### Inputs

Place these in `projects/{name}/input/wave/{wave}/`:

| File | Required? | What it is |
|---|---|---|
| `source_data.xlsx` | Yes | Aggregated survey data from Synapse |
| `source_raw_data.xlsx` | Optional | Respondent-level data for segment analysis / qualitative callouts |
| `KBQs.md` | Yes | Key business questions (hand-written by research team) |
| `Call Notes.docx` | Recommended | Client call notes + action items |
| `[prior wave].pptx` | Recommended | Prior wave report deck (deck-reader uses it) |
| `[survey draft].docx` | Recommended | Survey instrument for this wave |
| `methodology.odt` | Optional | Study design + methodology notes |

### Outputs

```
projects/{name}/output/{wave}/
  deck.pptx              # generated deck
  shape_registry.json    # shape-level data lineage (for live editing + refresh)
  backups/               # PPTX backups before each edit
```

---

## SlideGen Architecture

SlideGen is built around three co-equal layers:

### 1. Skills ([`.claude/skills/`](.claude/skills))

Claude Code skills = SKILL.md files that tell Claude *what to do* and *when*. Organized into 7 folders matching the building-block taxonomy:

```
.claude/skills/
├── analysis/         6 skills — hypothesis-generator, sfea-insight-writer (PET),
│                                atu-insight-writer (ATU), segment-comparator,
│                                stat-sig-annotator, trend-analyzer
├── context-data/     5 skills — deck-reader + 4 context builders
├── creation/         7 skills — slide-creator, slide-updater, slide-editor,
│                                deck-assembler, headline-writer, callout-writer,
│                                executive-summary-writer
├── planning/         7 skills — viz-selector, layout-selector, spec-validator,
│                                4 slide-plan-generator-* variants
├── projects/         2 skills — pet-deck, atu-deck (HCP-Pt + DT + PCA via project teams)
├── workflows/        8 skills — top-level orchestrators (see §Workflows)
├── pptx/             legacy general-purpose PPTX skill
└── slidegen/         legacy all-in-one skill (being decomposed)
```

### 2. `pptx_utils` Python package ([`slidegen/pptx_utils/`](slidegen/pptx_utils))

Composition primitives — hard-won OOXML knowledge encoded as callable functions. Without these, Claude rediscovers lxml gaps every session.

| Module | Key exports |
|---|---|
| `brand.py` | `BRAND{}` (33 pharma brand palettes), `CLIENT{}` (fonts, heading colors), `get_brand()` |
| `layout.py` | `LAYOUTS{}` (11 coordinate presets including 6 `observed_*` from real-deck clusters) |
| `charts.py` | `CHART_PATTERNS{}` (10 patterns), builders (`add_clustered_bar_chart`, `add_line_chart`, etc.) |
| `lxml_helpers.py` | 20+ OOXML manipulation helpers (axis invert, label positioning, series coloring, gridlines) |
| `shapes.py` | Textbox / solidrect / callout-box / image primitives |
| `tables.py` | Delta columns + value tables (with font sizing, alternating rows) |
| `text.py` | Run-level formatting helpers (delta formatting, wrap, truncate) |
| `registry.py` | Shape registry CRUD (data lineage tracking) |
| `com.py` | win32com helpers for live PPT editing |

### 3. Slide Spec contract ([`slidegen/slide_spec/`](slidegen/slide_spec))

The interface between intelligent skills (planners) and deterministic skills (renderers). Every slide is described as a `SlideSpec` with:

- `brand`, `layout`, `headline`, `subheadline`, `footer`, `section`
- `components`: list of Chart / LabelTable / ValueTable / DeltaColumn / Callout / Image / Textbox
- `data_lineage`: Synapse-canonical IDs (`project_id`, `reporting_plan_id`, `analysis_ids`, `segment_ids`, deliverables) OR Excel-path identifiers
- `metadata`: arc, hypothesis refs, role_in_arc, creation provenance

Every spec producer (planning skills) emits a `SlideSpec`. Every consumer (slide-creator, slide-updater, slide-editor, deck-assembler) accepts a `SlideSpec`. `spec-validator` enforces completeness at every handoff.

20 canonical example specs live in [`slidegen/slide_spec/examples/`](slidegen/slide_spec/examples) covering ~55% of real PET slide compositions.

---

## Workflows

Eight MECE workflows cover the full deck lifecycle. Each is a thin composition of building-block skills:

| # | Workflow | Verb | Triggered by |
|---|---|---|---|
| 1 | `create-deck-workflow` | Create | "Create a new PET deck for {brand}" |
| 2 | `refresh-deck-workflow` | Refresh | "Refresh wave 4 of the Rybrevant deck" |
| 3 | `edit-slide-workflow` | Edit | "Regenerate slide 5" / "Update slide 12 with fresh data" |
| 4 | `add-slide-workflow` | Add | "Add a slide answering this client question..." |
| 5 | `annotate-slide-workflow` | Annotate | "Add a callout to slide 10 with this quote..." |
| 6 | `structural-edit-workflow` | Restructure | "Remove slide 8" / "Reorder slides" |
| 7 | `deck-audit-workflow` | Audit | "Review this deck for issues before delivery" |
| 8 | `executive-summary-workflow` | Summarize | "Generate an ES answering these KBQs..." |

See the strategic PRD ([`docs/SlideGen_PRD_v1.1.md`](docs/SlideGen_PRD_v1.1.md)) §3 for workflow definitions and §5 for composition maps.

---

## Data Access

SlideGen reads from Synapse via four complementary tracks:

| Track | Method | When to use |
|---|---|---|
| **A — JSON-first** | Synapse `/reports/generate` → JSON | `synapse_report` extractions; no local Excel needed |
| **B — Excel (default)** | Local `source_data.xlsx` + pandas | Default path; works without Synapse API key |
| **C — Raw API (legacy)** | Synapse raw-responses + VQs + segments | `synapse_raw` extractions; full respondent-level data |
| **D — Raw-data-first (preferred for respondent-level)** | `synapse-cli` NDJSON stream | `raw_data_first` extractions; 1-2 API calls total, columnar cache |

Synapse auth fallback chain: explicit key → `SYNAPSE_API_KEY` env var → Azure AD auto-acquire → `synapse login` CLI cached token.

---

## Repo Structure

```
galen-consulting-r3m-report/
├── projects/                        # gitignored, shared via OneDrive
│   └── jnj_rybrevant/
│       ├── config.yaml
│       ├── input/wave/{wave}/
│       ├── context/{wave}/          # system-generated
│       └── output/{wave}/           # generated decks + registry
│
├── slidegen/                        # the engine
│   ├── slide_spec/                  # ★ the spec contract
│   │   ├── schema.py                # SlideSpec + components (dataclasses)
│   │   ├── validator.py             # validate_spec()
│   │   └── examples/                # 20 canonical example specs
│   ├── slide_creator.py             # ★ atomic unit: one spec → one slide
│   ├── slide_updater.py             # data refresh + headline regen
│   ├── slide_editor.py              # whitelisted structural edits
│   ├── deck_assembler.py            # list[SlideSpec] → PPTX
│   ├── deck_reader/                 # dual-mode (Connector tags + inference)
│   ├── viz_selector.py              # Metric → Q-type → chart_pattern
│   ├── layout_selector.py           # components → LAYOUTS key
│   ├── headline_writer.py           # template-based headline generator
│   ├── callout_writer.py            # quote / insight / freeform callouts
│   ├── slide_plan_refresh.py        # wave-refresh diff planner
│   ├── slide_plan_single.py         # one ask → one spec
│   ├── slide_plan_exec_summary.py   # KBQs → ES slides with citations
│   ├── segment_comparator.py        # per-segment stats
│   ├── stat_sig.py                  # Welch's t-test, chi-square, annotations
│   ├── trend_analyzer.py            # multi-wave trend detection
│   ├── pptx_utils/                  # composition primitives (see above)
│   └── pipeline/                    # R3M Report YAML-driven orchestrator
│
├── .claude/skills/                  # 40 SKILL.md files across 7 folders
├── experiments/
│   └── deck_analysis/               # Apr 15 analysis of 32 real PET decks
│       ├── decks/                   # source PPTXs (gitignored)
│       ├── outputs/                 # cluster data, headlines, chart positions
│       └── *.py                     # parser, analyzer, brand_mapper, detectors
├── docs/                            # PRDs + setup guides
├── tests/                           # Track 1-3 tests (40+ passing)
├── archive/                         # legacy scripts
└── README.md                        # this file
```

---

## Key Conventions

| Topic | Rule |
|---|---|
| **Percentages** | Excel stores 0.0-1.0 decimals. `pct_mode: pct` multiplies by 100; `pct_mode: straight` for whole-% values. Auto-detected from `_codes.value_range`. |
| **Deltas** | Always current minus prior, in percentage points. |
| **Color tokens** | 4 sources: explicit hex (`"#F75824"`), BRAND token (`"{brand.primary_current}"`), context file (`"{context.brand_palette.primary}"`), deck-reader extraction (`"{deck.slide_N.series_M.color}"`). Resolved at render time by slide-creator. |
| **Shape names** | `zrx_{slide:03d}_{shape:03d}` for COM targeting + registry. |
| **Spec → slide** | Every SlideSpec passes `validate_spec(strict=True)` before hitting any renderer. Fail loudly on incomplete specs. |
| **Headlines must refresh with data** | `slide-updater` regenerates via `headline-writer` on every data refresh. Preserving a stale headline against fresh data is a client-delivery risk (PRD §6.9). |
| **Never hardcode brand colors** | Brand palettes live in `BRAND{}`; renderers resolve color tokens at render time. |

---

## CLI

```bash
# YAML-driven R3M Report pipeline
python -m slidegen.pipeline.orchestrator projects/{name}/config.yaml
python -m slidegen.pipeline.orchestrator projects/{name}/config.yaml --fresh

# Standalone SlideGen ops
python -m slidegen create                   # demo slide
python -m slidegen edit <file.pptx>         # COM live editor
python -m slidegen reconcile <file.pptx>    # sync registry from PPT state
python -m slidegen fetch-synapse projects/{name}/config.yaml
python -m slidegen fetch-raw projects/{name}/config.yaml
python -m slidegen fetch-raw-first projects/{name}/config.yaml

# deck-reader CLI (extracts specs from existing PPTX)
python -m slidegen.deck_reader path/to/deck.pptx [--config config.yaml] [--out specs/]
```

---

## Developing

### Branch model

- `main` — stable
- `vijay-slidegen` — active development (current branch; Apr 2026)

### Tests

```bash
python tests/track1/test_slide_creator_patterns.py           # 10 chart patterns
python tests/track1/render_canonical_examples.py             # 20 canonical examples
python -m pytest tests/track3/test_analysis_skills.py -q     # 30 analysis tests
python -m pytest tests/                                      # full suite
```

### Contributing

- **Every skill** declares `inputs`, `outputs`, `composed_skills` in its SKILL.md frontmatter.
- **Every spec handoff** passes `validate_spec(strict=True)`.
- **Chart patterns** tie to deck-analysis frequencies — see `experiments/deck_analysis/outputs/ACTIONABLE_FINDINGS.md` before proposing new ones.
- **Don't write raw lxml** — use `pptx_utils/lxml_helpers.py`. Missing helper? Add it to `lxml_helpers.py`, don't inline.

### Dependencies

| Package | Purpose |
|---|---|
| `pandas` | Excel extraction (first run per wave, then JSON cached) |
| `openpyxl` | Excel reading |
| `python-pptx` | PowerPoint generation |
| `lxml` | OOXML chart manipulation |
| `pyyaml` | YAML config loading |
| `requests` | Synapse API integration |
| `pywin32` | COM live editing (Windows only) |
| `python-dotenv` | `.env` loading (optional) |
| `scipy` | Statistical tests (optional; pure-Python fallback in `stat_sig.py`) |

---

## Status (Apr 2026)

**Done:**
- Apr 15 deck analysis of 32 real PET decks across 17 clients → populated `BRAND{}`, `LAYOUTS{}`, `CHART_PATTERNS{}`
- Apr 16 deck analysis of 8 ATU decks across 7 clients → ATU section structure, chart frequency comparison vs PET
- Spec contract (`slidegen/slide_spec/`) + validator
- `slide-creator` Python with 10 chart patterns rendering at real-deck fidelity
- 20 canonical example specs (~55% real-slide coverage)
- Edit-mode primitives: `deck-reader` (dual-mode), `slide-updater`, `slide-editor`, `deck-assembler`
- Spec producers: viz-selector, layout-selector, headline-writer, callout-writer, 4 slide-plan-generators
- Analysis skills: segment-comparator, stat-sig-annotator, trend-analyzer
- 8 workflow orchestrator SKILL.mds

**In progress:**
- Line/doughnut PowerPoint Repair bug (line fixed; doughnut still under debug)
- Workflow end-to-end dogfooding against real PET deck

**Next:**
- PET + ATU dogfooding
- HCP-Pt / Digital Tracker / PCA project skills (in parallel via project-team contributors)
- `synapse-cli` integration for data layer
