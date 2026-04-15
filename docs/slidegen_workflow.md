# SlideGen Workflow

End-to-end reference for running the SlideGen pipeline, editing existing decks, and auditing the data trail from survey source → rendered slide.

Audience: analysts and reviewers. Engineering internals live in [SlideGen_PRD.md](SlideGen_PRD.md) and the source under [slidegen/](../slidegen/).

---

## 1. What SlideGen Does

Builds PowerPoint decks from survey data using a **YAML-driven pipeline**. One `config.yaml` per project (brand/product) defines:

- The source Excel(s) or Synapse API endpoints
- The extractions (which question codes → which values)
- The asks (slides, headlines, chart type, data binding)

No code changes to add a new project. Code changes only to add a new **renderer** (new chart type) or **extraction method** (new data shape).

### Data model: four tracks, two granularities

Two axes matter: **how** the data arrives (track) and **what granularity** it has (tier).

| Track | Source | Calls | Tier | Extraction methods | Cache |
|-------|--------|-------|------|--------------------|-------|
| **A — JSON-first** | Synapse `/reports/generate` | 1 per analysis | Aggregated | `synapse_report` | `source_data.json` |
| **B — Excel (default)** | Synapse banner plan → `source_data.xlsx` | Async download | Aggregated | `question_code`, `multi_question_code`, `row_range`, `question_code_multi_col`, `nested_ordinal` | `source_data.json` |
| **C — Raw API (legacy)** | Synapse `/surveys/download-responses` + VQ exports + segments | Multiple | Respondent-level | `synapse_raw` | `source_raw_data.json` |
| **D — Raw-data-first** ⭐ | Synapse NDJSON stream + segment join via `synapse-cli` | **1–2 total** | Respondent-level | `raw_data_first` | `context/{wave}/raw_data_first.json` |

**Track D** (shipped in commit `34266b0`, branch `feat/raw-data-first`) is the current preferred path for respondent-level work. It downloads **every survey response, every VQ column, and every segment assignment** in one NDJSON call plus one segment call, parses into a columnar structure (`columns`, `code_map`, `respondents`, `quarters`), caches the whole thing, and aggregates per-extraction on demand via `aggregate_extraction()`. Tracks A/B/C are untouched — zero regression risk.

Why Track D:
- **1–2 API calls** for an entire deck vs. N calls for Track C
- **Cache invalidation keyed on `synapse` config hash**, not an Excel file hash — nothing to manage on disk
- **Discovery**: `get_available_codes(config)` lists every code in the raw data (useful for skills + config generation)
- **Same 4 aggregation modes** as Tracks C and raw_aggregate: `single`, `multi_code`, `cross_brand`, `by_segment`

All tracks produce identical `{desc, prior, current, delta}` dicts. Renderers don't know which track the data came from.

Requires `pip install -e /path/to/synapse-cli`. Auth resolves via explicit key → `SYNAPSE_API_KEY` env var → Azure AD auto-acquire → **CLI cached token from `synapse login`** (new in Track D).

---

## 2. File Layout

```
projects/{name}/
├── config.yaml                      # the contract (see §6)
├── input/wave/{wave}/               # USER DROP ZONE
│   ├── source_data.xlsx             # Tier 1
│   ├── source_raw_data.xlsx         # Tier 2 (optional)
│   ├── call_notes.docx              # research team field intel
│   ├── prior_wave_es.md             # prior wave findings (text)
│   ├── {prior_wave_report}.pptx     # prior deck (optional)
│   ├── pet_project_kbq.odt          # study design
│   ├── KBQs.md                      # standing key business questions
│   └── {survey_draft}.{docx|pdf}    # survey instrument (optional)
├── context/                         # SYSTEM-GENERATED intermediates
│   ├── market_context.md            # product-level, NOT wave-versioned
│   └── {wave}/
│       ├── prior_wave_context.md    # Stage 0.5b output
│       ├── survey_context.md        # Stage 0.5c output
│       ├── project_context.md       # Stage 1 output
│       ├── hypothesis_bank.md       # Stage 2 output
│       ├── validated_analysis.md    # Stage 3 Phase 0 output
│       ├── narrative_threads.md     # Stage 3 Phase 1 output
│       ├── slide_plan.md            # Stage 4 output
│       ├── source_data.json         # Tier 1 cache (auto)
│       ├── source_raw_data.json     # Tier 2 cache (auto)
│       └── qualitative_data.json    # Stage 0q cache (auto, optional)
├── templates/
│   └── template.pptx                # clean python-pptx template, shared
├── output/{wave}/
│   ├── deck.pptx                    # generated deck
│   ├── shape_registry.json          # shape + data lineage
│   └── backups/                     # PPTX backups (max 10)
└── config_history/
    └── config_{YYYYMMDD_HHMMSS}.yaml # every config edit is backed up
```

`projects/` is **gitignored** — lives on OneDrive, symlinked into the clone. See [analyst_setup.md](analyst_setup.md).

---

## 3. The Full Create Workflow

Triggered by **"Create slides for projects/{name}"**. The pipeline runs 8 stages; most are automatic. Only **three user gates** exist: after Stage 1 (context review), after Stage 2 (hypothesis count), after Stage 3 (narrative threads), after Stage 4 (slide plan).

### Stage –1 — Fetch Synapse Data (optional)

Only if data isn't already in `input/wave/{wave}/`. Four tracks, picked by config:

- **Track A (JSON-first)** — `fetch_data_as_json()` calls `/reports/generate` → JSON directly. Bypasses Excel. Requires `synapse_report` extractions.
- **Track B (Excel, default)** — `trigger_generation()` + `wait_and_download()` downloads `source_data.xlsx`. Non-blocking split lets Stages 0.5–1 run in parallel while the banner plan builds.
- **Track C (Raw API, legacy)** — `fetch_all_raw()` downloads raw Excel + auto-discovers VQs + fetches segment defs. Requires `synapse_raw` extractions. Still supported but superseded by Track D for new work.
- **Track D (Raw-data-first)** ⭐ — `fetch_raw_data_first()` downloads all responses + VQs in one NDJSON call, joins segments in a second call, caches as `context/{wave}/raw_data_first.json`. Requires `raw_data_first` extractions and `synapse-cli` installed. Preferred path for respondent-level analysis.

If no Synapse API key is set, pipeline proceeds with the Excel already in `input/wave/{wave}/`.

### Stage 0 — Index Excel (auto, no gate)

```python
index_excel("input/wave/{wave}/source_data.xlsx",
            "context/{wave}/source_data.json")
```

Builds two critical indexes:
- `_column_layouts` — per-sheet prior/current column positions + segment columns
- `_codes` — per-question-code metadata: `row`, `desc`, `sub_row_count`, `has_sub_codes`, `sample_values`, `value_range` (decimal vs whole)

**This is the source of truth for code → data resolution.** Every later stage uses `_codes` to validate codes exist before referencing them.

### Stage 0q — Index Qualitative (auto, no gate, optional)

Runs automatically if `source_raw_data.xlsx` exists. Extracts verbatim/open-ended responses → `qualitative_data.json`. Hash-cached. Graceful skip if no raw data file.

Downstream: feeds HQ-type hypotheses in Stage 2, theme coding in Stage 3, `qual_theme_analysis` slides + `qual_callout` flags in Stage 4.

### Stages 0.5a / 0.5b / 0.5c — Parallel Context (PARALLEL background subagents)

Three independent context files. Run in parallel; wait for all before Stage 1.

| Skill | Writes | When it runs | Gate |
|-------|--------|--------------|------|
| `/market-context` | `context/market_context.md` | If missing (NOT wave-versioned) | Only if fact-check flags claims |
| `/prior-wave-context` | `context/{wave}/prior_wave_context.md` | If prior wave files detected | One prompt: file list confirmation |
| `/survey-context` | `context/{wave}/survey_context.md` | If survey draft detected | One prompt: file list + code count |

Files consumed by 0.5b and 0.5c are **excluded** from Stage 1 scanning.

### Stage 1 — Build Project Context (`/build-project-context`)

Reads: remaining input files (call notes, methodology ODT, KBQs, misc) + `market_context.md` + `prior_wave_context.md`.
Does NOT read: prior wave PPTX, survey draft, source_data (handled upstream).

Writes: `context/{wave}/project_context.md` — study design, KBQs, wave hypotheses from call notes, analytical priorities, methodology notes, message reference.

**Dependency:** `KBQs.md` must exist (hand-written by research team). Stop if missing.

**✋ GATE — single content validation:** User reviews all 4 context files before Stage 2. Shown as a summary table (file | section count | key stats). Edit or confirm.

### Stage 2 — Hypothesis Bank (`/hypotheses`)

Reads 5 context files:

| File | Required? | Provides |
|------|-----------|----------|
| `market_context.md` | Yes | Competitive/clinical landscape |
| `project_context.md` | Yes | Study design + field intel |
| `KBQs.md` | Yes | Domain structure |
| `prior_wave_context.md` | Strongly recommended | Domain metrics + prior recs |
| `survey_context.md` | Strongly recommended | Question codes + message list |

Produces `hypothesis_bank.md`. Hypothesis types: **New**, **PRIOR WAVE VALIDATION** (mandatory — one per prior finding/rec/gap), **[ACTION ITEM]**, **METHODOLOGY ARTIFACT**, **HQ** (qualitative).

Every hypothesis has a `Test with:` line with the exact question code(s) from `survey_context.md`.

**✋ GATE:** Count by type + domain coverage breakdown. Confirm before Stage 3.

### Stage 3 — Validate + Narrative Threads (`/sfea-insight-writer`)

**Phase 0 (auto):** Every hypothesis gets validated against `source_data.json`. For each hypothesis, extracts prior/current/delta per question code. Writes `validated_analysis.md`. **This is the audit link between hypothesis and data.**

**Phase 1:** Synthesizes 3–5 cross-domain story arcs (CONVERGENCE / TENSION / DIVERGENCE / CLOSURE). Writes arc-informed headlines, arc-organized executive summary, and arc-driven recommendations into one integrated `narrative_threads.md`.

**✋ GATE:** Single review of `narrative_threads.md`.

### Stage 4 — Slide Plan (`/slide-plan`)

Reads `narrative_threads.md`, `validated_analysis.md`, `hypothesis_bank.md`, `kbqs.md`, `survey_context.md`. Clusters hypotheses into slides with:

- Chart specs (renderer type)
- Arc assignment + "Role in arc"
- Headline (copied verbatim from `narrative_threads.md`)
- Question codes

Always includes Cover (Slide 1), ES (Slide 2), Recs (Slide 3). Sequencing: ACT NOW arcs → MONITOR arcs → CELEBRATE arcs.

Writes `context/{wave}/slide_plan.md`.

**✋ GATE:** Slide count + arc distribution + section breakdown. Confirm before Stage 5.

### Stage 5 — Generate config.yaml (internal, no gate)

```python
from slidegen.pipeline.config_generator import scaffold_config_from_plan
yaml_str = scaffold_config_from_plan(
    "context/{wave}/slide_plan.md",
    "context/{wave}/source_data.json",
    base_config_path="projects/{name}/config.yaml",  # merge, don't overwrite
)
```

Every question code in the slide plan is cross-checked against `_codes`. Auto-selects extraction method and `pct_mode` from `value_range`. Flags missing codes explicitly.

**Never copy extractions from another wave's config** — codes move between rows/sheets across waves.

### Stage 6 — `generate_deck()` (internal, no gate)

```python
from slidegen.pipeline import generate_deck
generate_deck("projects/{name}/config.yaml")
```

Loads config → validates → loads data (JSON cache or Excel extract) → renders slides → writes deck + shape_registry.

**Report shown to user:** success/failure, slide count, output path, any gaps.

---

## 4. Incremental Workflows

For projects with an existing config. All modify config.yaml; all back up config to `config_history/` first.

### Edit a single slide

**Trigger:** `Edit Slide N — {change}`

```python
from slidegen.pipeline import regenerate_slide
regenerate_slide("projects/{name}/config.yaml", slide_index=4)       # 0-based
regenerate_slide("projects/{name}/config.yaml", slide_index="ryb_mr") # by ask_id
```

1. Backup config
2. Modify ask N in YAML (headline, data_key, sort_by, extra)
3. `regenerate_slide()` — creates PPTX backup, clears only that slide, re-renders
4. User reopens deck in PowerPoint

### Edit slides with new wave data

**Trigger:** `Edit slides with new wave data — {wave_id}`

1. Backup config
2. Update `project.wave`, `period_current`, `period_prior`
3. Verify `input/wave/{wave_id}/source_data.xlsx` exists
4. `generate_deck()` — auto-extracts Excel → new JSON → new `output/{wave_id}/deck.pptx`

Old wave output preserved in its own folder.

### Edit slides with new wave data + new asks

**Trigger:** `Edit slides with new wave data + new asks — {wave_id}`

1. Backup config
2. Update wave fields
3. Read new reference docs
4. Modify extractions + asks in YAML
5. Gap analysis — new extractors or renderers?
6. `generate_deck()`

### Refresh deck

**Trigger:** `Refresh this deck`

```python
from slidegen.pipeline import refresh_deck
result = refresh_deck("projects/{name}/config.yaml")
# {refreshed: [...], skipped: [...], errors: [...]}
```

1. Force re-extract (delete `source_data.json` cache)
2. Read `shape_registry.json` → find slides with `data_source`
3. Call `regenerate_slide()` for each (with PPTX backup)
4. Update `last_data_pull` and `last_refreshed` in registry

### Force fresh extraction

```bash
python -m slidegen.pipeline.orchestrator projects/{name}/config.yaml --fresh
```

Or `generate_deck(config_path, force_fresh=True)`. Bypasses JSON cache, re-extracts from Excel.

---

## 5. The Audit Chain

The question this doc was written to answer: **how do we prove a rendered number traces back to the source survey?**

Five artifacts, chained:

```
slide_plan.md              →  hypothesis ID + question code
     ↓
validated_analysis.md      →  code + extracted prior/current/delta
     ↓
config.yaml                →  code + extraction method + sheet
     ↓
source_data.json._codes    →  code + row + value_range + sample_values  (+ _meta.excel_hash)
     ↓
shape_registry.json        →  slide # + data_source + last_data_pull + renderer
     ↓
speaker notes              →  slide + code + full question text + extraction method
```

**What's persisted:**
- `source_data.json._meta` has the Excel hash → proves which source file it came from
- `_meta.extracted_at` has the timestamp → proves when
- `shape_registry.json` has `last_data_pull` and `last_refreshed` per slide → proves freshness
- `config_history/` has timestamped YAML backups → proves what changed when
- Every data-driven slide has auto-written speaker notes with codes + question text → human-readable spot check

**Practical audit procedure:**

1. Commit `context/{wave}/` and `output/{wave}/shape_registry.json` so diffs are reviewable across waves.
2. **After Stage 5:** grep every code in `slide_plan.md` against generated `config.yaml`. `scaffold_config_from_plan()` flags missing codes but doesn't block.
3. **After Stage 6:** spot-check 2–3 slides:
   - Open deck → read speaker notes → copy code
   - `grep <code> context/{wave}/source_data.json` → find row in `_codes`
   - Confirm rendered value matches the row
4. **Staleness check:** pipeline prints `[STALE]` if `_meta.extracted_at` > 24h old. Use `--fresh` to bypass cache.

**Not yet mechanized:** a single `audit_{wave}.md` joining slide # | headline | hypothesis ID | code | row | prior | current | extracted_at. ~50 lines to add as a post-Stage-6 step.

---

## 6. config.yaml Reference

Minimum structure:

```yaml
project:
  name: "JNJ Rybrevant"
  client: "Janssen"
  wave: "Q1 2026"
  period_current: "Q1'26"
  period_prior: "Q4'25"

data_source_path: "input/wave/{{wave}}/source_data.xlsx"
raw_data_source_path: "input/wave/{{wave}}/source_raw_data.xlsx"   # optional
template_path: "templates/template.pptx"
output_path: "output/{{wave}}/deck.pptx"
context_path: "context/{{wave}}/"

synapse:                              # optional, for Tracks A/C
  base_url: "https://..."
  survey_id: "..."
  segments:
    - name: "Academic vs Community"
      code: "SEG_AC"
      mode: groupby                   # or: filter

brands:
  primary:
    name: "Rybrevant + Lazcluze"
    short: "RYB"
    color_current: "#F75824"
    color_prior: "#FFC199"
  competitor:
    name: "Tagrisso"
    short: "TAG"
    color_current: "#7030A0"
    color_prior: "#AD88C8"

sheets:
  primary: "Rybrevant"
  competitor: "Tagrisso"

extractions:
  mr_recall:
    method: question_code              # see §7 for all methods
    sheet: primary
    params:
      code: "Q2_10Z"
      pct_mode: pct                    # or: straight
      use_label_shortcuts: true

asks:
  - ask_id: ryb_mr
    slide_type: single_bar_with_delta  # see §8 for all renderers
    headline: "{{primary.name}} message recall gains ground"
    data_key: mr_recall
    sort_by: current
    sort_desc: true
    extra:
      qual_callout:                    # optional on any slide
        quote: "..."
        attribution: "Oncologist, Academic"
        theme: "Efficacy"
```

`{{wave}}`, `{{primary.name}}`, `{{period_current}}`, `{{client}}` are interpolated at load time.

---

## 7. Extraction Methods (complete list)

**Tier 1 — Aggregated Excel / JSON-first**

| Method | Use case |
|--------|----------|
| `question_code` | Find code row, walk sub-rows, extract prior/current |
| `multi_question_code` | One row per code (e.g. CTA: compelling, changed, closing) |
| `row_range` | Fixed row range with column mapping |
| `question_code_multi_col` | Multiple columns per row (e.g. HII: hi vs other) |
| `nested_ordinal` | Grouped ordinal sub-rows (1st/2nd/3rd recall) |
| `mock` | Hardcoded test data from `params.rows` (skips JSON cache) |
| `synapse_report` | JSON-first: calls `/reports/generate` directly |

**Tier 2 — Raw local**

| Method | Use case |
|--------|----------|
| `raw_aggregate` | Parse respondent rows, aggregate via `top2box`/`yes_pct`/`recall_pct`/`mean` |

**Tier 3 — Raw API (legacy)**

| Method | Use case |
|--------|----------|
| `synapse_raw` | Fetch raw Excel from Synapse, auto-merge VQs, apply segment cuts, aggregate locally |

**Tier 4 — Raw-data-first (preferred for respondent-level)**

| Method | Use case |
|--------|----------|
| `raw_data_first` | Download all responses + VQs + segments via `synapse-cli` NDJSON stream (1–2 API calls), cache columnar JSON, aggregate per-extraction on demand |

All raw tiers (2, 3, 4) support modes: `single`, `multi_code`, `cross_brand`, `by_segment`.

**`raw_data_first` params:** `code` or `codes`, `agg_fn` (`recall_pct` / `top2box` / `yes_pct` / `mean`), `mode`, `quarter_current` / `quarter_prior` (default from `project.period_*`), optional `segment_column` + `segment_value` (or `primary_*` / `comp_*` / `segment_values` depending on mode).

**Discovery:** `get_available_codes(config)` returns `[{code, text, column_count, sample_values}, ...]` for every code in the raw data — use it during config scaffolding instead of probing `_codes`.

---

## 8. Renderers (complete list)

22 slide types in [slidegen/pipeline/slide_renderers/](../slidegen/pipeline/slide_renderers/):

**Narrative:** `cover`, `executive_summary`
**Bars:** `single_bar_with_delta`, `dual_bar_with_delta`, `dual_bar_qoq`, `qoq_bar_with_delta`, `two_section_bar`
**Compare:** `clustered_compare`, `dual_bar_compare`, `stacked_order`, `hii_scorecard`, `dual_doughnut`
**Dots:** `abacus`, `dual_abacus`, `followup_rep`, `message_mbd`
**Lines:** `trended_scorecard`, `trended_activity`
**Other:** `quadrant_scatter`, `heatmap_table`, `qual_theme_analysis`

**Renderer-data compatibility:** Most extractions produce `{desc, prior, current}` — only `single_bar_with_delta`, `abacus`, and `executive_summary` work directly with this format. Renderers like `dual_bar_with_delta`, `clustered_compare`, `hii_scorecard`, `heatmap_table` require specific field prefixes or `extra` config — do not assign unless extraction produces matching data.

---

## 9. CLI

```bash
# Generic pipeline
python -m slidegen.pipeline.orchestrator projects/{name}/config.yaml
python -m slidegen.pipeline.orchestrator projects/{name}/config.yaml --fresh

# SlideGen utility commands
python -m slidegen create                        # demo
python -m slidegen edit <file.pptx>              # interactive live editor (Windows, requires PowerPoint open)
python -m slidegen reconcile <file.pptx>         # sync registry from live PPT
python -m slidegen fetch-synapse <config>        # Tracks A + B
python -m slidegen fetch-raw <config>            # Track C
python -m slidegen fetch-raw-first <config>      # Track D (requires synapse-cli)
```

---

## 10. Common Failure Modes

| Symptom | Likely cause | Fix |
|---------|--------------|-----|
| `[STALE]` warning | `extracted_at` > 24h old | Add `--fresh` or delete `source_data.json` |
| PowerPoint repair error on open | Template was a prior wave deck with embedded charts | Use clean python-pptx template |
| Bar labels truncated | Label > `LABEL_MAX` (65 single, 55 dual) | Set `label_shortcuts` + `use_label_shortcuts: true` |
| T2B extraction pulls scale 1–7 | `question_code` grabs all sub-rows | Use `row_range` on the T2B summary block |
| Percentages are 10× too big | Mixed `pct_mode` — row is whole-number, config says `pct` | Set `pct_mode: straight` or check `_codes.value_range` |
| Code in slide plan not in config | `scaffold_config_from_plan` flagged missing code | Search `_codes` for the exact code; fix slide plan or extraction |
| `fetch_all_raw` fails on VQ merge | New VQ not yet in survey | Re-run `fetch_all_raw` after Synapse updates |
| Slide renders but no data | Extraction ran but renderer expects different keys | Check renderer compatibility (§8) |

---

## 11. What Lives Where (cheat sheet)

| Need to... | Look at |
|-----------|---------|
| Add a new project | Create `projects/{name}/config.yaml`; run `discover_excel_structure()` |
| Add a new chart type | [slidegen/pipeline/slide_renderers/](../slidegen/pipeline/slide_renderers/); register in `__init__.py` |
| Add a new extraction method | [slidegen/pipeline/data_loaders.py](../slidegen/pipeline/data_loaders.py) |
| Debug a rendered value | Speaker notes → code → `source_data.json._codes` → Excel row |
| Roll back a config change | `config_history/config_{timestamp}.yaml` |
| Roll back a deck | `output/{wave}/backups/` |
| Check data freshness | `source_data.json._meta.extracted_at` + `shape_registry.json.last_data_pull` |
| Understand a skill's behavior | `.claude/skills/{skill-name}/SKILL.md` |
