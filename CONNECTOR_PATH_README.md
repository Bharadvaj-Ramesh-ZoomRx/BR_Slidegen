# Connector Path — End-to-End Refresh Workflow

This document explains what happens when the connected-refresh pipeline runs against any deck with Galen Connector tags. Read it before refreshing a new deck so it's clear what each step does, what it touches, and how to interpret the output.

## TL;DR — refreshing a deck

```bash
git clone <repo>
cd <repo>
pip install -r requirements.txt
# put credentials in synapse env/.env or docs/.env (see Authentication below)
python -m slidegen.refresh_pipeline path/to/deck.pptx
```

The refreshed deck lands at `output_testing/deck_output/<auto-derived-name>.pptx` with badges on each slide telling you what got refreshed and what didn't. The pipeline is fully automatic — no per-deck configuration, no per-slide intervention. Slides that can't be refreshed (no connector tags, ambiguous tags, or tags pointing at the wrong analysis) are clearly flagged so a human can review them.

## Prerequisites

Before the first run on a new machine:

| Requirement | Where | How |
|---|---|---|
| Python 3.10+ | system | `python --version` |
| Repo dependencies | `requirements.txt` | `pip install -r requirements.txt` |
| Synapse API URL | `synapse env/.env` (preferred) or `docs/.env` or repo-root `.env` | `SYNAPSE_API_URL=https://...` |
| Synapse credentials | same `.env` | `SYNAPSE_API_KEY=...` (preferred) — without it the pipeline falls back to `synapse-cli` cached JWT; run `synapse login` once. |
| LLM credentials (optional) | same `.env` | `LLM_API_KEY=...` enables the headline-rewrite step. Without it, headlines are preserved verbatim. |

Verify with a quick health check:

```bash
python -c "import slidegen.refresh_pipeline; print('OK')"
python -m slidegen.refresh_pipeline --help
```

## What "connector path" means

Decks delivered to Galen clients are PowerPoint files where each chart and table has a **Connector tag** — a chunk of XML stamped onto the shape that records:

- Which Synapse analysis the data came from (`AnalysisIds`, `ProjectId`, `ReportingPlanId`)
- Which time periods to fetch (static IDs + `dynamic_latest_n` + `include_live_wave`)
- How to pivot the data (`PivotConfig`: `RowFields`, `ColumnFields`, `ValueFields`, `Filters`, `columnDefinitions`)
- How to map pivot output to the chart (`MappingConfig`: `selectedColumns`, `selectedRows`, `applyTranspose`)
- Per-shape diagnostic notes (`REFRESH_NOTE`)

A "connected refresh" reads those tags, fetches the latest data from Synapse, and rewrites the chart/table values without changing layout, formatting, or anything else on the slide. Slides without these tags are reported as `non_connected` and left alone.

## Single-command entrypoint

```bash
python -m slidegen.refresh_pipeline <deck.pptx>
python -m slidegen.refresh_pipeline <deck.pptx> --no-cache --workers 8 --out-name MyDeck_v2
```

Outputs land under `output_testing/` by default:

- `output_testing/deck_output/<slug>_<date>_v1.pptx` — the final annotated deck
- `output_testing/json_output/<slug>_<date>_v1_full_spec.json` — regenerated spec
- `..._refresh_status.json` — per-component refresh outcome
- `..._headline_status.json` — per-slide headline rewrites
- `..._label_sync_status.json` — period-banner edits

## Pipeline — five steps

### Step 0 (implicit): Spec regeneration

`refresh_pipeline._regen_spec(source_pptx, out_spec)` walks every chart/table on every slide, reads its connector tag via `slidegen/deck_reader/tag_reader.py`, and writes a `*_full_spec.json` describing every component. This is recomputed on every run so static-pin overrides, split-viz positions, and segment lineages picked up *live* from the deck — never read from a stale checked-in JSON.

Each component records:
- `name` (the physical shape's `name` attribute)
- `shape_id` (the shape's intrinsic XML id — primary key for refresh-time lookup)
- `position` (left/top in inches, used as a tiebreaker if names collide)
- `data_source` (a hash key into the spec's `data_sources` map)
- `raw_pivot_config` + `raw_mapping_config` (verbatim copies of the connector tags)
- `static_time_period_ids` / `static_time_period_names` (per-component pin)
- `chart_pattern` (xy_scatter_abacus, column_stacked, etc.)
- `split_order` / `rows_per_object` / `top_n_rows` (split-viz coordinates)

### Step 1: `refresh_deck_from_spec`

`slidegen.intelligent_refresh.refresh_deck_from_spec()` does the heavy work:

1. **Parallel pre-fetch** — every unique `data_source` in the spec is fetched from Synapse via `ThreadPoolExecutor(max_workers=N)` with on-disk caching at `projects/.synapse_cache/` (cache key = config hash + today's date, daily auto-expiry). Empty-API and HTTP errors degrade to `pd.DataFrame()` rather than failing the run.
2. **Per-component refresh** — for each component:
   - Resolve which physical shape via **shape-id** (primary) → name + position (fallback). Catches dense slides where multiple tables sit within ±0.3" of each other.
   - Decide whether to skip as static. A component is "truly static" only when `static_time_period_ids` is set AND `dynamic_latest_n` is 0/None AND `include_live_wave` is False. Mixed-pin (static + live-wave intent) refreshes normally.
   - Run `pivot_records_to_chart_data()` (the mapper) — see Step 1.5.
   - Write the result via the right path:
     - **Charts** → python-pptx `replace_data()` with formatCode-aware value scaling. Scatter/abacus uses `XyChartData` with X/Y axis-aware scaling.
     - **Tables** → `cell_values` 2D grid (auto-computed by `_auto_compute_cell_values()` when the spec doesn't carry one). Layouts handled: 1×N vertical, N×1 horizontal, M×N matrix, 1×1 templated `(n = X)` substitution.
3. **Notes per shape** — writes a `REFRESH_NOTE` Connector tag onto each shape so Connector itself sees the diagnostic context next time someone opens the deck.

### Step 1.5: The mapper (`synapse_chart_mapper.py`)

`pivot_records_to_chart_data()` is the heart of the refresh logic. It replicates Connector's transformation pipeline against API records:

1. **Value field resolution** — `ValueFields=["mean"]` doesn't match a column? Try the `measure` discriminator column (filter to rows where measure=="Mean", aggregate `value`). Falls through to `percentage` → `decimal` → `value` if neither works.
2. **Filters** — multi-value `;`-joined filters use `isin`; if no value matches the long form ("Specialty - X - CARD"), try compound-suffix matching ("CARD"). Failed filters leave df untouched (don't wipe).
3. **Static time-period pin** — when `static_time_period_names` is set, restrict to those waves.
4. **Latest-wave-only filter** — when `time_period_name` is in the data but NOT in `ColumnFields` and multiple waves are present, restrict to the chronologically latest wave (avoids summing 2 waves of Mean values into 148%).
5. **Pivot** — pandas `pivot_table` with `AggregationType`-aware aggfunc (0=sum, 1=mean, 2=count, 3=min, 4=max).
6. **CustomList row sort** — apply `columnDefinitions[].sortCriteria.CustomList` ordering before split-viz slicing.
7. **selectedColumns filter** — keep only the columns named in `MappingConfig.selectedColumns`. Tracks unmatched entries (drift) and zero-match (`alignment_failed` → preserves source).
8. **Apply transpose** — `applyTranspose: true` swaps rows/columns at this point. Aliases applied via `columnDefinitions[].Alias`.
9. **Formula-column evaluation** — derived columns declared as `<blank:Alias>` with an Excel-style `Formula` (`=B2/100`, `=B2*100`, `=B2+C2`, `=B2-C2`) are evaluated against existing pivot columns. Column letters map to `columnDefinitions` positions; the corresponding series is matched via 4-tier resolution (exact → progressive-suffix → tail-substring → fuzzy 0.85).
10. **Chronology sort** — when categories are wave-shaped (`Nov'25`, `Q1'26`, `Wave 12`) — including segment-suffixed forms (`Nov'25 - Gastro`) — sort by parsed wave token. Direction inferred from source: latest-wave-position-zero → descending; latest-position-last → ascending.
11. **Source-canonical alignment** — when source had distinct `categories` and `series_names`, align refreshed values to source order via the same 4-tier resolution. Dynamic charts get a positional reorder to source order even when alignment isn't strictly necessary, so Wave-7-then-Wave-6 visual order is preserved on refresh.
12. **Tag-mismatch detection** — zero overlap on both axes between source and API → return `tag_mismatch` note with API data faithfully written. User reviews the tag.
13. **Alignment-failed safety net** — every series ends up all-None? Return `alignment_failed` so caller preserves source.

### Step 2: `stamp_refresh_notes`

Per-shape `REFRESH_NOTE` Connector tag stamped onto each component's CustomXML rels: a one-line summary of the refresh outcome (status + note kinds + counts). Visible to anyone opening the deck through Connector.

### Step 3: `refresh_headlines` (LLM)

`slidegen/headline_refresh.py` rewrites slide headlines based on refreshed chart values. Skipped if `LLM_API_KEY` not set.

- `_find_headline_shape(slide)` picks the **topmost wide title shape** (regardless of whether the current text uses narrative verbs). On multi-Title slides it no longer picks a section header below the real headline just because the section header happens to use a verb like "trended."
- `_build_prompt()` adapts to whether the original is narrative-style or a category label. If the original isn't narrative-shaped, the prompt instructs the LLM to *write a fresh* narrative from the refreshed data instead of preserving the original's voice.
- `_call_claude()` uses LiteLLM (default model: `anthropic/claude-sonnet-4-6`).
- Slides whose values didn't change (`_values_eq`) or whose refresh corrupted them (`_all_none`) are skipped.

### Step 4: `sync_period_labels`

`slidegen/label_sync.py` walks every chart on every slide. For each chart whose category axis shifted (e.g. `Oct'25..Mar'26` → `Nov'25..Apr'26`), it builds a per-chart `{old_token: new_token}` shift map by extracting wave tokens from compound labels. Then for each text frame OR table cell whose bbox sits **immediately above or below** the chart (within 0.4" vertically, ≥ 30% horizontal overlap) AND whose text reads as a **label** (short, mostly wave tokens with chart connectors like "vs"/"to") — the wave tokens get rewritten in place. Footnotes, sidebars, headlines, methodology text — left alone.

### Step 5: Slide-level annotation

`scripts/annotate_refresh_and_headline.py` adds two rounded-rect badges to each slide's top-right corner:
- `REFRESH ok 12/12` / `REFRESH 4 ok / 8 static` / `REFRESH tag-mismatch 3/3 - review tag` / `STATIC (no refresh)` / `NON-CONNECTED slide`
- `HEADLINE rewritten` / `HEADLINE preserved` / `HEADLINE skipped (no narrative)`

Each slide's **speaker notes** also get an appended status block (below a divider) listing every component, its status, and any diagnostic notes (tag_mismatch detail, drift, partial). Existing notes (Q text, sample sizes) are preserved via lxml direct paragraph append. Idempotent — re-runs replace the prior block.

## Modules at a glance

| File | What it owns |
|---|---|
| `slidegen/refresh_pipeline.py` | CLI entry + `run_full_pipeline()`. Orchestrates Steps 0-5. Auto-derives output filename via `_slugify()`. |
| `slidegen/intelligent_refresh.py` | `refresh_deck_from_spec` (Step 1), `stamp_refresh_notes` (Step 2). Houses `_claim_shape` (shape-id matching), `_auto_compute_cell_values` (table auto-write), `_format_cell_value`, `walk_shapes_recursive` (group-walk), `_refresh_templated_count_with_records` (1×1 cell substitution). |
| `slidegen/synapse_chart_mapper.py` | `pivot_records_to_chart_data` — the Connector-emulation pivot/map engine. Filter resolver, formula evaluator, chronology sort, source-canonical alignment, fuzzy matching. |
| `slidegen/headline_refresh.py` | `refresh_headlines` (Step 3). Headline-shape detector, narrative classifier, LLM prompt builder, LiteLLM call. |
| `slidegen/label_sync.py` | `sync_period_labels` (Step 4). Per-chart shift map, spatial adjacency filter, run-level token replacement. |
| `slidegen/deck_reader/tag_reader.py` | `generate_config_specs` — extracts every connector tag (PivotConfig, MappingConfig, ReportConfig hashes) from CustomXML parts. |
| `tests/build_full_spec.py` | `build_data_sources`, `build_connected_slide` — flattens connector specs into the format the refresher consumes. |
| `slidegen/synapse_auth.py` | Token resolution: explicit key → env var → Azure AD → cached JWT. |
| `scripts/annotate_refresh_and_headline.py` | Step 5 badge logic. Has `_refresh_badge` and `_headline_badge` classifiers used by the orchestrator. |
| `scripts/refresh_creon_today.py` (and friends) | Per-deck convenience runners. Mostly superseded by `python -m slidegen.refresh_pipeline`. |

## Evals

Tests under `tests/evals/end_to_end/` validate the pipeline against committed goldens:

| Eval | Purpose |
|---|---|
| `test_refresh_execution.py` | Refreshes a fixture deck end-to-end and compares categories, series names, and **values** against a per-deck golden JSON. Catches "mapper succeeded but data didn't move" silently — that class of bug was invisible under categories-only comparison. Mutation test confirms the comparator isn't always-green. |
| `test_propose_pivot_config.py` | Unit tests for the spec-inference logic (`propose_pivot_config`, `propose_raw_configs`) used to bootstrap specs from charts that lack tags. |
| `test_spec_refresh_pipeline.py` | Legacy 3-stage spec dataclass pipeline. Test-only, not the active production path. |

To regenerate goldens after an intentional pipeline change: `python -m tests.evals.end_to_end.generate_golden`.

## Outputs and how to read them

After a run completes:

1. **Open the final deck** — every slide has badges at top right showing refresh + headline status. Open the speaker notes pane to see per-component status.
2. **`*_refresh_status.json`** — per-slide, per-component status. Most useful columns: `status` (`ok`, `static_pinned_skipped`, `ambiguous_tag`, `alignment_failed`, `tag_mismatch`, `not_found`, `empty`), `notes[].kind` (`dynamic_added`, `partial_alignment`, `selectedColumns_drift`, `tag_mismatch`).
3. **`*_headline_status.json`** — `updates[].status`: `updated`, `unchanged`, `no_chart`, `no_headline`, `api_error`.
4. **`*_label_sync_status.json`** — per-slide shift map + every text-frame edit.
5. **`*_full_spec.json`** — complete spec consumed by Step 1. Useful for debugging: every component's resolved tags, position, shape_id, lineage.

## Slide bucket meanings

| Bucket | What it means | Action |
|---|---|---|
| `ok` | All components refreshed cleanly. | None. |
| `ok_with_dynamic_added` | New items flowed in via dynamic tag (extra waves, segments). | Verify they should appear; otherwise narrow the tag. |
| `ok_with_partial` | Static tag, but API has extras the user didn't pin. | Optional: opt those extras in. |
| `ok_with_static_pinned` | Mixed slide — some refreshed, some genuinely pinned. | Per-component status in speaker notes. |
| `non_connected` | No connector tags. | Out of scope; manual edit. |
| `static_review` | Whole slide is static-pinned. | Confirm the pin is still wanted. |
| `tag_mismatch` | Source vs API have zero label overlap on both axes — tag points at wrong analysis. | Repoint the tag, or move slide to non-connected. |
| `ambiguous_tag` | 2+ charts on the slide carry identical tags but had different source values. | Add a per-chart filter or split_order to disambiguate. |
| `alignment_failed` | API returned data but couldn't be aligned (every refreshed value ended up None even after fuzzy matching). | Investigate the tag's `selectedColumns` / `selectedRows`. |

## Authentication

Synapse credentials are loaded from (in order):
1. `.env` in repo root
2. `docs/.env`
3. `synapse env/.env` (shared across machines)

`SYNAPSE_API_URL` is required. `SYNAPSE_API_KEY` is preferred; without it, falls back to `synapse-cli` cached JWT (run `synapse login` to populate).

`LLM_API_KEY` (LiteLLM-compatible) enables Step 3 headline rewrites. Without it, headlines pass through.

## Known limitations

- **1×1 cells with embedded counts**: a wide range of conventions are handled — `(n = X)`, `(N = X)`, `n=X`, `Sample size: X`, `Sample Size = X`, `Base: X`, `Total: X`, `based on X respondents`, `X HCPs/patients/subjects`. Cells using a pattern not in this list will need explicit handling; PRs welcome.
- **Scatter chart series naming**: handled — when a connector tag yields a numeric series name (rare edge case where selectedColumns picks bare value columns without a row-label dimension), the writer falls back to source's series name at the same index, then to `columnDefinitions[].Alias` / `Name`.
- **Group-nested shapes**: chart and table refresh walks groups via `walk_shapes_recursive`. Label-sync also handles grouped charts.
- **Field-date stamps** (`"Q1'26: 01/01/2026 – 25/02/2026"`): the wave token portion gets shifted by label sync, but the date range does not. Reason: the Synapse API doesn't currently return `time_period_start` / `time_period_end` in record responses, so there's no source of truth for the new date range. Requires API-side change before this can be auto-wired.
- **Ambiguous tags**: if a slide has 2+ charts that all resolve to the same analysis with byte-identical tag configs but should display different cuts of data, the pipeline marks them `ambiguous_tag` and preserves source. Add a per-chart filter (`Filters` in PivotConfig) or `split_order` to disambiguate. The pipeline can't auto-resolve this without risking wrong assignments.
- **Period-id reverse-chrono**: time_period_ids are NOT chronological in Synapse. The `dynamic_latest_n` filter parses wave NAMES instead — this is handled, just worth knowing if you debug period-window issues.

## Where to look when something goes wrong

- **Wrong values in chart**: open `*_refresh_status.json`, find the slide and component, check its status + notes.
- **Headline says wrong thing**: check `*_headline_status.json` for the slide; the LLM saw the chart summary in its prompt — if the data is right but the wording is off, that's an LLM issue, not a pipeline issue.
- **Period banner says old wave**: check `*_label_sync_status.json` for the slide. If `slides_with_shift` doesn't include the slide, the chart's category axis didn't shift (so no rewrite was triggered).
- **Whole slide blank or unchanged**: bucket says `non_connected` (no tag) or `static_pinned_skipped` (pinned).

## Adding a new deck

No code changes are required for new decks.

```bash
python -m slidegen.refresh_pipeline path/to/deck.pptx
```

The pipeline auto-derives the output filename from the deck stem (so `"Repatha HCP ATU - Q2'26 Skeleton (1).pptx"` becomes `Repatha_HCP_ATU_Q2_26_Skeleton_<date>_v1.pptx`). For scatter/abacus charts, the badge on each slide reports refresh status — review those slides directly when in doubt.

## Will "refresh this deck" just work?

Short answer: yes for the majority of slides; the rest get clearly flagged and preserved as source.

What "just works" end-to-end after a single command:

- Tagged charts and tables across all 5 layout patterns (vertical label, horizontal label, matrix, formula columns, 1×1 templated cells with `(n = X)` placeholders).
- Charts and tables nested inside group shapes.
- Mixed-pin slides (some charts static-pinned, others dynamic).
- Charts with formula-derived columns (`<blank:Alias>` with `=B2/100` etc.).
- Headlines (rewritten by LLM when chart values changed; left alone otherwise; fresh narrative written when source headline isn't narrative-shaped).
- Period banners adjacent to charts (`Oct'25..Mar'26` rewritten to `Nov'25..Apr'26` etc.).
- Per-slide badges showing exactly what happened.
- Per-component status appended to each slide's speaker notes (existing notes preserved).

What gets flagged instead of refreshed (source values preserved, badge tells you why):

- `non_connected` — no connector tags on the slide. Manual edits only.
- `static_pinned_skipped` — tag explicitly pinned to historical waves.
- `tag_mismatch` — connector tag points at an analysis whose label structure differs from the source. Repoint the tag.
- `ambiguous_tag` — multiple charts on the slide share an identical tag fingerprint but had different source values; pipeline can't disambiguate. Add a per-chart filter or `split_order`.
- `alignment_failed` — refreshed values couldn't be aligned to source structure even after fuzzy matching. Investigate `selectedColumns` / `selectedRows` in the tag.

What is **not** auto-handled (will need manual edit OR future work):

- Field-date stamps like `Q1'26: 01/01/2026 – 25/02/2026` — only the `Q1'26` portion gets shifted, not the date range.
- Templated 1×1 cells using patterns OTHER than `(n = X)` — e.g. `"Sample size: 120"` or `"based on 51 respondents"`.
- Scatter charts where the connector tag selects bare value columns without a row-label dimension — series naming may default to the first numeric value.
- Manual-edit charts that share a tag with sibling charts on the same slide — flagged as `ambiguous_tag` rather than corrupted.

In practice, on a typical deck with valid connector tags and live Synapse data, ~90%+ of slides refresh cleanly with no human intervention. The remaining 10% land in one of the five flagged buckets above with a clear reason.
