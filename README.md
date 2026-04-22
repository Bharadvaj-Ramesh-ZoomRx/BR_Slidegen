# SlideGen

**SlideGen** is ZoomRx Galen-Consulting's deck-generation and refresh engine. It reads existing PowerPoint decks, produces auditable slide specs, fetches live data from Synapse, and refreshes charts + tables in place — preserving all formatting.

> **PRD:** [`docs/SlideGen_PRD_v1.1.md`](docs/SlideGen_PRD_v1.1.md) | **Analyst setup:** [`docs/analyst_setup.md`](docs/analyst_setup.md)

---

## What SlideGen Does

```
Existing PPTX deck
    |
    v
[deck-reader] ── extracts layout, components, Connector tags ──> SlideSpec (v1.2)
    |                                                              - positions + chrome (render config)
    |                                                              - data_mapping (fetch + transform instructions)
    |                                                              - data_lineage (project, analysis, segments)
    v                                                              - NO data values (spec is CONFIG, not DATA)
[Synapse API] ── fetches flat records using data_lineage ──>
    |
    v
[pivot engine] ── applies PivotConfig + MappingConfig ──> chart-ready data
    |               (faithful Connector replication)
    v
[clone + refresh] ── replace_data() on cloned deck ──> refreshed PPTX
                      (preserves ALL formatting)
```

**Two paths for spec generation:**
- **Connected slides** (have Synapse Connector tags): raw PivotConfig + MappingConfig extracted from Custom XML Parts and stored in spec. Refresh uses exact Connector transformation logic.
- **Non-connected slides** (no tags): `data_inference.py` matches chart series names from OOXML against Synapse column values to infer the DataTransform. Same refresh pipeline.

Both paths produce the same spec format. The SlideSpec is the single audit layer — audits data lineage, transformation logic, rendering config, and headline provenance.

---

## Quick Start

### Setup

```bash
git clone https://github.com/ZoomRx/galen-consulting-r3m-report.git
cd galen-consulting-r3m-report
pip install pandas openpyxl python-pptx lxml pyyaml requests python-dotenv msal
```

Create `.env` in repo root:
```
SYNAPSE_API_BASE_URL=https://synapse-api.zoomrx.com
SYNAPSE_API_KEY=sk_Ga_ZHWo30zrOGbcnR8hIvOv3WRGz25HKYuhCqGpAe7s
SYNAPSE_API_TOKEN=Bearer sk_Ga_ZHWo30zrOGbcnR8hIvOv3WRGz25HKYuhCqGpAe7s
```

### Generate Config Specs from an Existing Deck

```python
from slidegen.deck_reader.tag_reader import generate_config_specs
from slidegen.slide_spec.schema import dump_spec
import json

specs, summary = generate_config_specs("path/to/deck.pptx")
# Save specs
json_str = json.dumps([json.loads(dump_spec(s)) for s in specs], indent=2)
Path("deck_specs.json").write_text(json_str)
```

### Refresh a Deck from Specs

```bash
python tests/test_spec_refresh_pipeline.py --stage 1   # clone source -> dummy
python tests/test_spec_refresh_pipeline.py --stage 2   # refresh from Synapse via specs
python tests/test_spec_refresh_pipeline.py --stage 3   # verify source ~ refreshed
python tests/test_spec_refresh_pipeline.py --all        # all 3 stages
```

### Refresh a Deck from Spec (Intelligent Pipeline)

```bash
# 1. Embed config into PPTX (scans slides, writes manifest as Custom XML Part)
python -m slidegen.intelligent_refresh embed-config --pptx deck.pptx

# 2. Build spec from Connector configs (connected slides)
python tests/build_full_spec.py

# 3. Read slide context for Claude Code interpretation (non-connected slides)
python -m slidegen.intelligent_refresh read --pptx deck.pptx --slide 0

# 4. Refresh all slides from single spec JSON
python -m slidegen.intelligent_refresh refresh-deck --spec "deck.json"
# Output: Output_deck.pptx (same directory as spec)

# 5. Write data-grounded headline
python -m slidegen.intelligent_refresh headline --pptx Output_deck.pptx --slide 0 --text "..."
```

**File convention:**
```
deck.pptx          # source deck
deck.json          # spec (same name, .json) — data_sources + per-slide mappings
Output_deck.pptx   # refreshed output
```

### R3M Report Pipeline (YAML-driven)

```bash
python -m slidegen.pipeline.orchestrator projects/{name}/config.yaml
python -m slidegen.pipeline.orchestrator projects/{name}/config.yaml --fresh
```

---

## Architecture

### Core Modules

| Module | Purpose |
|---|---|
| **`slidegen/slide_spec/schema.py`** | SlideSpec v1.2 dataclasses: DataTransform, SeriesConfig, ChartDataMapping, TableDataMapping, SegmentRule, split viz fields |
| **`slidegen/slide_spec/data_inference.py`** | Non-connected slide inference: series name matching (exact, fuzzy/abbreviation, partial) against Synapse column values |
| **`slidegen/deck_reader/tag_reader.py`** | `read_tagged_shapes()` for data extraction, `generate_config_specs()` for config-only spec generation with raw Connector configs |
| **`slidegen/synapse_chart_mapper.py`** | Connector-faithful pivot engine: multi-value pivots, applyTranspose, split visualization, selectedRows/Columns, CustomList sort, column alias resolution |
| **`slidegen/synapse_auth.py`** | Auth resolution: API key (sk_*) > env token (JWT) > MSAL cache > device code flow |
| **`slidegen/slide_refresher.py`** | Clone + in-place refresh via `replace_data()` |
| **`slidegen/slide_creator.py`** | Spec-to-slide renderer (deterministic) |
| **`slidegen/pptx_utils/`** | Composition primitives: brand palettes (102 clients, 33 brands), chart patterns (15), layouts (11), lxml helpers (20+) |
| **`slidegen/pipeline/`** | R3M Report YAML-driven orchestrator with 4-tier data loading |

### SlideSpec v1.2 Schema

The spec stores **CONFIG** (fetch instructions + render config), not **DATA** (no chart values, no table text):

```
SlideSpec
  ├── slide_id, slide_index, layout, brand, section
  ├── headline (template or {{headline_writer}} prompt)
  ├── components[]
  │   ├── ChartComponent
  │   │   ├── position, chart_pattern, chrome (gap_width, overlap, data_labels, legend)
  │   │   └── data_mapping: ChartDataMapping
  │   │       ├── transform: DataTransform (row_field, column_field, value_field, filters)
  │   │       ├── series_config: [SeriesConfig (role, color, order)]
  │   │       ├── raw_pivot_config (Connector DataFrameConfigHash JSON)
  │   │       ├── raw_mapping_config (Connector MappingConfig JSON)
  │   │       └── split_order, rows_per_object, top_n_rows
  │   ├── ValueTableComponent (headers, rows, data_mapping)
  │   ├── LabelTableComponent (labels, font config, data_mapping)
  │   └── TextboxComponent, CalloutComponent, ImageComponent
  ├── data_lineage: DataLineage
  │   ├── project_id, project_name, reporting_plan_id
  │   ├── analysis_ids, analysis_names
  │   ├── segments[]: SegmentRule (rule_id, rule_name, values, data_column)
  │   ├── time period config (static IDs or dynamic latest N)
  │   └── Excel-path identifiers (fallback)
  └── metadata: tier, confidence, slide_layout_name
```

### Synapse Data Flow

```
data_lineage (from spec)
    → Synapse POST /api/reports/generate
    → flat records [{y_label, segment_1, time_period_name, percentage, base, ...}]
    → pivot_records_to_chart_data():
        1. Resolve value fields (case-insensitive, fallback chain)
        2. Build compound column keys from ColumnFields (NaN->"(blank)")
        3. Multi-value pivot (separate columns per ValueField)
        4. Filter columns by selectedColumns (last-part priority matching)
        5. Apply topNRows + rowsPerObject/splitOrder (split visualization)
        6. Apply transpose (explicit applyTranspose from MappingConfig)
        7. Apply column aliases from columnDefinitions
        8. Filter/order rows by selectedRows
        9. Apply CustomList sort from sortCriteria
    → ChartRefreshData (categories + series ready for replace_data)
```

### Intelligent Refresh Pipeline

Single spec JSON per deck. Dual-mode engine — raw Connector configs for connected slides, Claude Code interpretation for non-connected slides.

```
Spec JSON (deck.json)
  ├── data_sources: {key: {project_id, reporting_plan_id, analysis_ids, ...}}
  └── slides[]
      ├── Connected: raw_pivot_config + raw_mapping_config per component
      │   → pivot_records_to_chart_data() — exact Connector fidelity
      └── Non-connected: data_mapping {segment_filter, series_column, ...}
          → Claude Code interpretation → pandas pivot
```

**PPTX carries an embedded config** (Custom XML Part) — lightweight manifest of slide IDs and data source assignments. When the deck is edited in PowerPoint (reorder, add, delete slides), the embedded config is used to auto-reconcile spec.json.

```
embed-config → read → interpret → refresh-deck → headline
               ↑                      ↓
           Claude Code           Output_deck.pptx
```

### Auth Resolution

```
SYNAPSE_API_KEY (sk_*)           ← permanent, no expiry (preferred)
  → SYNAPSE_API_TOKEN (Bearer)   ← JWT from browser, ~90min expiry
    → MSAL token cache            ← auto-refresh from prior login
      → Device code flow          ← one-time browser login (requires public client flows)
```

---

## Repo Structure

```
galen-consulting-r3m-report/
├── slidegen/                        # the engine
│   ├── slide_spec/                  # ★ spec schema + inference
│   │   ├── schema.py                # SlideSpec v1.2 dataclasses
│   │   ├── data_inference.py        # Non-connected slide inference
│   │   ├── validator.py             # validate_spec()
│   │   └── examples/                # canonical example specs
│   ├── deck_reader/                 # ★ spec extraction from PPTX
│   │   ├── tag_reader.py            # Tier 1 (Connector tags) + generate_config_specs()
│   │   └── __init__.py              # read_deck() orchestrator (Tier 1 + Tier 2)
│   ├── synapse_chart_mapper.py      # ★ Connector-faithful pivot engine
│   ├── synapse_auth.py              # Auth (API key + JWT + MSAL)
│   ├── slide_refresher.py           # Clone + in-place refresh
│   ├── slide_creator.py             # Spec -> slide renderer
│   ├── intelligent_refresh.py       # ★ Dual-mode refresh engine: embedded config, raw Connector + interpreted mappings, headline writer
│   ├── pptx_utils/                  # composition primitives
│   │   ├── brand.py                 # 102 clients, 33 brand palettes
│   │   ├── charts.py                # 15 chart patterns
│   │   ├── layout.py                # 11 layout presets
│   │   ├── lxml_helpers.py          # 20+ OOXML manipulation helpers
│   │   ├── shapes.py, tables.py, text.py, images.py, deck.py
│   │   ├── com.py                   # win32com live editing
│   │   └── registry.py              # shape registry CRUD
│   └── pipeline/                    # R3M Report YAML-driven pipeline
│       ├── orchestrator.py          # entry point
│       ├── project_config.py        # YAML config dataclasses
│       ├── data_loaders.py          # 4-tier data loading
│       ├── slide_renderers/         # 22 slide type renderers
│       └── synapse_*.py             # Synapse fetchers (4 tracks)
│
├── .claude/skills/                  # 40 SKILL.md files across 7 folders
│   ├── analysis/                    # hypothesis gen, insight writers, stat tests
│   ├── context-data/                # deck-reader, context builders
│   ├── creation/                    # slide-creator, headline-writer, deck-assembler
│   ├── planning/                    # viz-selector, layout-selector, slide-plan generators
│   ├── projects/                    # pet-deck (276 decks), atu-deck (143 decks)
│   └── workflows/                   # 8 top-level orchestrators (incl. refresh-deck-workflow)
│
├── experiments/deck_analysis/       # 905-deck mass grounding exercise
├── tests/
│   ├── test_spec_refresh_pipeline.py  # 3-stage connected refresh test (legacy, uses Connector specs directly)
│   ├── test_intelligent_refresh.py    # Non-connected refresh test harness
│   ├── test_nonconnected_refresh.py   # Non-connected inference prototype (mechanical)
│   ├── build_full_spec.py             # ★ Builds complete spec.json from Connector configs + manual specs
│   ├── [deck].json                    # ★ Spec file (same name as PPTX)
│   └── test_synapse_*.py             # Synapse integration tests
├── projects/                        # gitignored, shared via OneDrive
├── docs/                            # PRDs + setup guides
└── archive/                         # legacy scripts
```

---

## Key Concepts

| Concept | Description |
|---|---|
| **Spec as CONFIG** | SlideSpec stores fetch instructions + render config, never data values. Same spec + new time periods = new deck. |
| **SlideSpec as audit layer** | The spec audits data lineage, transformation logic, rendering config, headline provenance. Not just data. |
| **Clone + refresh** | Clone source PPTX (preserves ALL formatting), update only DATA via `replace_data()`. Never reconstruct slides. |
| **Raw Connector configs** | For connected slides, raw PivotConfig + MappingConfig stored per-component. Refresh uses them directly for exact Connector fidelity. |
| **Series-name inference** | For non-connected slides, chart series names from OOXML matched against Synapse column values to infer DataTransform. |
| **Split visualization** | Connector's `rowsPerObject` + `SPLITORDER` splits one pivot dataset across multiple chart shapes. Each shape gets 1 row. |
| **Two-pass model** | Pass 1 (automatic): layout + components from PPTX. Pass 2 (data lineage): from Connector tags or Synapse catalog matching. |
| **Intelligent interpretation** | Claude Code reads spatial layout to determine data mappings — group shape labels, position-based component grouping, segment filter inference. |
| **Headline grounding** | Headlines written from refreshed data, not copied from source. Ensures no stale headlines survive a data refresh. |

---

## Status (Apr 2026)

**Done:**
- 905-deck mass grounding exercise (102 clients, 33 brands, 15 chart patterns, 11 layouts)
- SlideSpec v1.2 with spec-as-config architecture
- Dual-mode refresh engine: raw Connector configs (connected) + Claude Code interpretation (non-connected)
- Single spec.json per deck with data_sources catalog + per-slide component mappings
- Embedded config in PPTX (Custom XML Part) for structure tracking + spec reconciliation
- 126/126 charts refresh on 29-slide UAT deck (0 errors)
- Headline writer: data-grounded headlines from refreshed chart values
- Non-connected slide inference: spatial label interpretation (group shapes, position proximity)
- Synapse API key auth (no manual token pasting)
- 22 slide type renderers, 40 skills

**In progress:**
- Table refresh from Synapse (currently source-restored; chart refresh is primary)
- Category ordering edge cases (4 slides with custom sort not matching source)
- Column alias formatting (SOV% vs sov casing)

**Next:**
- End-to-end dogfooding on production PET + ATU decks
- config.yaml as thin project-level metadata layer
- Claude Code interpretation loop for remaining non-connected slides
