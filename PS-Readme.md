# PS-Readme: Intelligent Slide Refresh Pipeline

## Overview

The intelligent refresh pipeline (`slidegen/intelligent_refresh.py`) takes an existing PPTX deck, reads its chart/table structure, connects each component to Synapse data, and refreshes the charts with live data. It handles two types of slides:

- **Connected** — shapes already have Galen Connector tags (PivotConfig + MappingConfig embedded in OOXML)
- **Non-connected** — shapes have no tags; the pipeline infers the data mapping from the DataFrame

Both paths converge at refresh time: once a non-connected chart has its `raw_pivot_config` + `raw_mapping_config` derived, it is refreshed through the same engine as connected charts.

---

## Architecture

```
Source PPTX
    |
    v
+--------------------------------------------------+
|  STEP 1: READ                                    |
|  read_slide_context(pptx, slide_index)           |
|  -> Extracts per-shape: type, position, name,    |
|     chart_pattern, series_names, categories,      |
|     series_values, table headers/rows             |
|  -> Returns dict ready for inspection             |
|                                                   |
|  format_slide_for_interpretation(ctx)             |
|  -> Human-readable text dump of the slide         |
+--------------------------------------------------+
            |
            |   For each shape, check:
            |   does it have Connector tags?
            |
    +-------+--------+
    |                 |
    v                 v
CONNECTED         NON-CONNECTED
(has tags)        (no tags)
    |                 |
    |                 v
    |       STEP 2: USER PROVIDES
    |         - project_id
    |         - reporting_plan_id
    |         - analysis_id (per chart)
    |         - segment_ids (optional)
    |         - dynamic_latest_n (optional)
    |         - static flag per table
    |                 |
    |                 v
    |       STEP 3: FETCH + INFER
    |         fetch_synapse_data(lineage)
    |         -> Returns (records, DataFrame)
    |
    |         propose_pivot_config(df, chart_shape=None)
    |         -> Classifies columns by role:
    |            temporal, categorical, numeric,
    |            identifier, label
    |         -> Assigns: row_field, series_column,
    |            value_field, val_format
    |         -> Returns confidence score
    |
    |         propose_raw_configs(chart_shape_or_None, df)
    |         -> Builds raw_pivot_config +
    |            raw_mapping_config dicts
    |         -> Uses chart_shape for Jaccard validation
    |            when available
    |                 |
    |                 v
    |       STEP 4: USER CONFIRMS
    |         Show proposed: row_field, series_column,
    |         value_field, confidence, derivation reasons
    |         User approves or provides corrections
    |                 |
    |                 v
    |       STEP 5: VERIFY (optional)
    |         verify_analysis_for_chart(
    |           analysis_id, chart_series,
    |           project_id, reporting_plan_id)
    |         -> Fetches minimal data (latest_n=1)
    |         -> Jaccard similarity >= 0.6 = confirmed
    |         -> < 0.6 = rejected, ask user for correct ID
    |                 |
    v                 v
+--------------------------------------------------+
|  STEP 6: BUILD SPEC                              |
|  save_proposed_configs_to_spec(spec_path,         |
|      slide_index, updates)                        |
|  -> Writes raw_pivot_config + raw_mapping_config  |
|     per component into spec JSON                  |
|  -> Creates data_source entries per analysis_id   |
|  -> Component-level data_source override when     |
|     charts on same slide use different analyses   |
|  -> Spec saved to: output/<name>_spec.json        |
+--------------------------------------------------+
            |
            v
+--------------------------------------------------+
|  STEP 7: REFRESH                                 |
|  refresh_deck_from_spec(spec_path)                |
|                                                   |
|  Per component, routes by config presence:        |
|                                                   |
|  raw_pivot_config + raw_mapping_config present?   |
|   YES -> RAW CONNECTOR PATH                       |
|          pivot_records_to_chart_data()             |
|          (synapse_chart_mapper.py)                 |
|          Exact Connector fidelity.                 |
|   NO  -> INTERPRETED MAPPING PATH                 |
|          pandas pivot_table() on data_mapping      |
|          (row_field/series_column/value_field)     |
|                                                   |
|  Both paths produce categories + series -> write  |
|  into chart via replace_data()                    |
|                                                   |
|  Output: output/<name>_refreshed.pptx             |
+--------------------------------------------------+
            |
            v (optional)
+--------------------------------------------------+
|  STEP 8: STAMP CONNECTOR TAGS                    |
|  write_connector_tags(pptx, slide_index,          |
|      shape_configs, data_lineage)                 |
|  -> Writes PivotConfig + MappingConfig as         |
|     p:tagLst XML into the PPTX                    |
|  -> After this, Synapse Connector UI recognises   |
|     the shapes as connected                       |
|  -> Converts non-connected -> connected           |
+--------------------------------------------------+
```

---

## Key Modules

### Production (active pipeline)

| Module | Purpose |
|--------|---------|
| `slidegen/intelligent_refresh.py` | Main orchestrator. CLI entry point. Contains: `read_slide_context`, `fetch_synapse_data`, `propose_pivot_config`, `propose_raw_configs`, `verify_analysis_for_chart`, `save_proposed_configs_to_spec`, `refresh_deck_from_spec`, `write_connector_tags`, embedded config read/write |
| `slidegen/synapse_chart_mapper.py` | Connector-fidelity pivot engine. `pivot_records_to_chart_data()` takes flat records + PivotConfig + MappingConfig and produces categories + series. Replicates Connector's exact transformation logic (compound row index, transpose, filters, CustomList sort, field remapping) |
| `slidegen/synapse_auth.py` | Synapse API token resolution (explicit key, env var, Azure AD, CLI cached token) |

### Supporting (used by production but not part of refresh flow)

| Module | Purpose |
|--------|---------|
| `slidegen/deck_reader/tag_reader.py` | Reads Connector tags from PPTX shapes. `read_tagged_shapes()` for full extraction, `generate_config_specs()` for config-only specs. Used by the SlideSpec pipeline (see Legacy section) |
| `slidegen/deck_reader/inference.py` | Tier 2 structural inference for untagged shapes |
| `slidegen/slide_spec/schema.py` | Core dataclasses: `SlideSpec`, `ChartComponent`, `Position`, `DataLineage`, etc. Deeply used by slide creation/editing modules |
| `slidegen/slide_spec/validator.py` | Spec validation |

### Legacy (SlideSpec-based refresh — parallel system, test-only)

These form a parallel refresh pipeline that uses `SlideSpec` dataclasses instead of JSON specs. They are **not called by `intelligent_refresh.py`** but are imported by test files.

| Module | Purpose |
|--------|---------|
| `slidegen/PS_slide_refresher.py` | `refresh_deck()` using SlideSpec objects |
| `slidegen/slide_refresher.py` | Another SlideSpec-based refresher |
| `tests/test_nonconnected_refresh.py` | End-to-end non-connected test (Repatha) using SlideSpec |
| `tests/test_spec_refresh_pipeline.py` | SlideSpec refresh test with `apply_data_transform()` |
| `tests/PS_test_synapse_refresh.py` | CREON deck round-trip test using PS_slide_refresher |
| `tests/build_full_spec.py` | Generates full spec JSON from `generate_config_specs()` |

---

## Spec JSON Format

The spec is a plain JSON file (not SlideSpec dataclasses). It is the single source of truth for what to refresh and how.

```json
{
  "source_deck": "../Creon - 1 slide.pptx",
  "data_sources": {
    "p523_rp1143_a641211": {
      "project_id": 523,
      "reporting_plan_id": 1143,
      "analysis_ids": [641211],
      "segment_ids": [],
      "dynamic_latest_n": 3
    }
  },
  "slides": [
    {
      "slide_index": 0,
      "data_source": null,
      "components": [
        {
          "type": "chart",
          "name": "PS",
          "chart_pattern": "column_stacked_100_vertical",
          "data_mapping": {
            "row_field": "time_period_name",
            "series_column": "option",
            "value_field": "decimal"
          },
          "raw_pivot_config": { ... },
          "raw_mapping_config": { ... },
          "data_source": "p523_rp1143_a641211"
        },
        {
          "type": "table",
          "name": "Table 17",
          "static": true
        }
      ]
    }
  ]
}
```

### Component routing at refresh time

| Config present | Refresh path | Engine |
|---------------|--------------|--------|
| `raw_pivot_config` + `raw_mapping_config` | RAW CONNECTOR PATH | `pivot_records_to_chart_data()` in `synapse_chart_mapper.py` |
| `data_mapping` only (row_field/series_column/value_field) | INTERPRETED MAPPING PATH | pandas `pivot_table()` in `intelligent_refresh.py` |
| `"static": true` | Skipped | No refresh |

### Data source resolution

- Slide-level `data_source` applies to all components on that slide
- Component-level `data_source` overrides the slide-level (used when charts on the same slide need different analysis_ids)
- Key format: `p{project_id}_rp{reporting_plan_id}_a{analysis_id}`
- Data is cached per data_source key during refresh (fetched once, reused)

---

## Key Functions in intelligent_refresh.py

### Reading

| Function | Purpose |
|----------|---------|
| `read_slide_context(pptx, slide_index)` | Extract all shapes with positions, chart data, table data, text |
| `format_slide_for_interpretation(ctx)` | Human-readable text dump |
| `format_data_for_interpretation(df)` | Human-readable DataFrame summary |

### Data Fetching

| Function | Purpose |
|----------|---------|
| `fetch_synapse_data(lineage)` | Fetch records from Synapse API via `synapse-cli`. Returns `(records, DataFrame)`. Auth resolved automatically via CLI fallback chain |

### Inference (non-connected path)

| Function | Purpose |
|----------|---------|
| `_classify_columns(df)` | Classify every DataFrame column into: temporal, categorical, numeric, identifier, label. Uses column names + value patterns |
| `propose_pivot_config(df, chart_shape=None)` | Infer row_field, series_column, value_field from DataFrame. When chart_shape provided, validates with Jaccard overlap. Returns confidence score + derivation trace |
| `propose_raw_configs(chart_shape_or_None, df)` | Builds full `raw_pivot_config` + `raw_mapping_config` dicts from inference. Delegates to `propose_pivot_config()` internally |

### Verification

| Function | Purpose |
|----------|---------|
| `verify_analysis_for_chart(analysis_id, chart_series, project_id, reporting_plan_id)` | Validates that an analysis_id's data matches a chart's series names. Jaccard >= 0.6 = confirmed |

### Spec Management

| Function | Purpose |
|----------|---------|
| `save_proposed_configs_to_spec(spec_path, slide_index, updates)` | Write derived configs into spec JSON. Creates data_source entries. Supports per-component data_source override |
| `read_embedded_config(pptx)` | Read SlideGen config from PPTX Custom XML Part |
| `write_embedded_config(pptx, config)` | Write/overwrite SlideGen config in PPTX |
| `generate_embedded_config(pptx)` | Scan PPTX and generate config manifest (slide IDs, shape names, component types) |
| `reconcile_spec_with_config(spec, config)` | Reconcile spec against embedded config (handles reorder/add/delete) |

### Refresh

| Function | Purpose |
|----------|---------|
| `refresh_deck_from_spec(spec_path, pptx_path=None, output_path=None)` | Main refresh entry point. Clones source PPTX, fetches data per data_source, refreshes each component. Routes through RAW CONNECTOR or INTERPRETED MAPPING path based on config presence |
| `refresh_slide_from_mapping(pptx, output, slide_index, mapping, lineage)` | Single-slide refresh with explicit mapping + lineage |
| `write_headline(pptx, slide_index, text)` | Write headline text into a slide |

### Connector Tag Stamping

| Function | Purpose |
|----------|---------|
| `write_connector_tags(pptx, slide_index, shape_configs, lineage)` | Stamp non-connected shapes with Synapse Connector tags (in-place PPTX modification). After this, Connector UI recognises them as connected |
| `_build_tag_xml(tags)` | Serialise name/value dict as p:tagLst XML |
| `_inject_custdata(slide_xml, shape_name, cust_tag)` | Inject custDataLst into a shape's nvPr |
| `_shape_tag_rid(slide_xml, shape_name)` | Find existing tag rId for a shape |

---

## CLI Commands

```bash
# Read slide context
python -m slidegen.intelligent_refresh read --pptx deck.pptx --slide 0

# Read slide + fetch data for comparison
python -m slidegen.intelligent_refresh read --pptx deck.pptx --slide 0 --lineage lineage.json

# Embed config into PPTX
python -m slidegen.intelligent_refresh embed-config --pptx deck.pptx

# Show embedded config
python -m slidegen.intelligent_refresh show-config --pptx deck.pptx

# Refresh all slides from spec
python -m slidegen.intelligent_refresh refresh-deck --spec output/spec.json

# Refresh with explicit output path
python -m slidegen.intelligent_refresh refresh-deck --spec output/spec.json --output output/refreshed.pptx

# Setup non-connected: verify + derive configs + stamp tags
python -m slidegen.intelligent_refresh setup-nonconnected \
    --spec output/spec.json \
    --pptx deck.pptx \
    --slide 0 \
    --shapes '[{"shape_name":"PS","analysis_id":641211},{"shape_name":"PMS","analysis_id":641205}]' \
    --project-id 523 \
    --reporting-plan-id 1143

# Write headline
python -m slidegen.intelligent_refresh headline --pptx deck.pptx --slide 0 --text "New Headline"
```

---

## Column Classification Heuristics (propose_pivot_config)

When inferring pivot config from a DataFrame without chart context:

| Role | Detection method | Examples |
|------|-----------------|----------|
| **temporal** | Column name in {time_period_name, period, quarter, wave, date} OR values match patterns: Q1'26, Jan'26, Wave 3, 2026-01 | `time_period_name` |
| **categorical** | String column, 1-100 unique values, not temporal/identifier/label | `option`, `segment_1`, `measure` |
| **numeric** | Numeric dtype. Sub-classified: `decimal` (0-1), `whole` (0-100), `integer` (other) | `decimal`, `percentage`, `base` |
| **identifier** | Column name in {analysis_id, project_id, ...} OR numeric with cardinality > 50 | `analysis_id`, `time_period_id` |
| **label** | Column name in {y_label, product, brand, attribute, message} | `y_label` |

### Field assignment priority

**value_field**: Known name with decimal range (`decimal`, `percentage`) > first 0-1 range column > known name any range > first numeric

**row_field** (no chart context):
- Label column wins over single-value temporal (single wave = degenerate)
- Temporal wins when it has higher cardinality than label (trended chart)
- Falls back to highest-cardinality categorical

**row_field** (with chart context): Jaccard overlap between column values and chart categories

**series_column** (no chart context):
- Known series name (`option`, `measure`, `segment`, `metric`) > lowest-cardinality categorical (2-10 unique = series-like)

**series_column** (with chart context): Jaccard overlap between column values and chart series names

---

## Data Flow: Synapse API -> Chart

```
Synapse API (/reports/generate)
    |
    v
flat records: [{time_period_name, option, decimal, base, ...}, ...]
    |
    v
DataFrame (pandas)
    |
    +--- propose_pivot_config(df) ---> {row_field, series_column, value_field}
    |
    +--- propose_raw_configs(None, df) ---> {raw_pivot_config, raw_mapping_config}
    |
    v
save_proposed_configs_to_spec() ---> spec.json
    |
    v
refresh_deck_from_spec()
    |
    +--- RAW CONNECTOR PATH (raw_pivot_config present):
    |      pivot_records_to_chart_data(records, pivot_config, mapping_config)
    |        -> PivotConfig defines: RowFields, ColumnFields, ValueFields, Filters
    |        -> MappingConfig defines: selectedColumns, selectAllRows, applyTranspose
    |        -> Produces: ChartRefreshData(categories, series)
    |
    +--- INTERPRETED PATH (data_mapping only):
    |      df.pivot_table(index=row_field, columns=series_column, values=value_field)
    |        -> Produces same categories + series shape
    |
    v
chart.replace_data(CategoryChartData or XyChartData)
    |
    v
Restore formatCodes from source chart (replace_data resets them)
    |
    v
prs.save(output_path)
```

---

## Known Decisions and Gotchas

1. **time_period_id is NOT chronological** — for CREON project 523, 2026 IDs are assigned in reverse order. Do not sort client-side. Trust the Synapse CLI to fix ordering via `dynamic_latest_n`.

2. **Minimal user inputs** for non-connected slides: `project_id`, `reporting_plan_id`, `analysis_id(s)`, `segment_ids`, `dynamic_latest_n`, static flag per table. Everything else is inferred from data.

3. **formatCode preservation** — `replace_data()` resets all formatCodes. The pipeline saves them before replacement and restores them after.

4. **Category reordering** — pivot may return categories in alphabetical order, but the source chart has Connector's order. The pipeline reorders to match source when the sets are equal.

5. **Per-component data_source** — when charts on the same slide use different analysis_ids, each component gets its own `data_source` key. The refresh engine fetches and caches per key.

6. **Two parallel systems exist** — the JSON spec pipeline (`intelligent_refresh.py`) is the active one. The SlideSpec dataclass pipeline (`PS_slide_refresher.py`, `test_spec_refresh_pipeline.py`) is legacy/test-only. Do not mix them.

7. **All intermediary outputs go to `projects/{name}/output/`** — spec JSONs, refreshed PPTXs, verification results.

8. **Table Connector tags — simple vs complex**: Simple tables (all visible columns mapped to one Connector pivot) get full Connector tags + `selectedColumns` with only data columns (exclude `row_field`). Complex tables (extra columns outside the mapping, multiple header/row columns) get `cell_values` written + unmapped numeric cells marked `xx` — but NO Connector tags, since Connector can't reproduce partial layouts.

9. **Connector tag minimalism**: Only write the 8 essential tags: `ANALYSISTYPE`, `COLUMNKEYLABELMAP`, `DATAFRAMECONFIGHASH`, `DATAFRAMECONFIGHASH_BACKUP`, `LASTREFRESHTIME`, `MAPPINGCONFIG`, `REPORTCONFIGHASH`, `REPORTCONFIGHASH_BACKUP`. Extra tags (`DARWINVERSION`, `UPDATE`, `VISUALISATION_ID`, `STASSIGID`, `MIGRATIONPHASE`, `FAILEDREFRESHICONID`) cause Connector to mishandle tables.

10. **columnDefinitions sorted alphabetically** — matches Connector convention. Required for computed column formulas to reference correct Excel column letters (A, B, C...).

11. **Computed columns in tables**: Use `<blank:ColName>` in `selectedColumns` and a corresponding entry in `columnDefinitions` with `Alias`, `Formula` (e.g. `=B2+D2+E2`), `IsDefaultAlias: false`. The formula is auto-computed from `source_columns` positions in the sorted `columnDefinitions`. Caller only needs to provide `name` and `source_columns`.

12. **Table selectedColumns**: For charts, include `row_field` (categories axis). For tables, exclude `row_field` — including it adds an unwanted "Deliverable" column when Connector refreshes.
