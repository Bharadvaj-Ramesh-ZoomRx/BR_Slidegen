---
name: refresh-deck-workflow
effort: high
paths: ["slidegen/intelligent_refresh.py", "slidegen/synapse_chart_mapper.py", "slidegen/slide_spec/schema.py", "tests/test_spec_refresh_pipeline.py", "tests/test_intelligent_refresh.py"]
description: "End-to-end slide/deck refresh. Handles BOTH Connector-tagged (connected) and non-tagged (non-connected) slides. Connected: raw PivotConfig + MappingConfig for exact Connector fidelity. Non-connected: Claude Code reads spatial layout, interprets data mappings, executes deterministic refresh. Both paths write data-grounded headlines. Triggered by 'Refresh this slide/deck' or 'Update data on slide N'."
---

# refresh-deck-workflow

Refresh any slide or deck with fresh Synapse data — connected or non-connected.

## Trigger Phrases

- "Refresh this slide / deck"
- "Update the data on slide 6"
- "Refresh Repatha ATU Slide 6.pptx with project 1428, analysis 545991"
- "Pull in Q1 '26 data and rebuild the deck"

## Architecture

```
Source PPTX
    ├── Has Connector tags? ──YES──→ CONNECTED PATH
    │                                 Extract specs (generate_config_specs)
    │                                 Fetch data (Synapse API)
    │                                 Pivot (pivot_records_to_chart_data)
    │                                 Refresh charts + tables (replace_data)
    │
    └── No tags ──────────────→ NON-CONNECTED PATH
                                  Step 1: read_slide_context() → shapes + positions
                                  Step 2: Claude Code interprets spatial layout
                                  Step 3: refresh_slide_from_mapping()
    ↓
BOTH PATHS → Write data-grounded headline → Output PPTX
```

## Cardinal Rules

1. **Two paths, one pipeline.** Connected slides use raw Connector configs for exact fidelity. Non-connected slides use Claude Code interpretation. Both produce the same output: refreshed PPTX with data-grounded headlines.
2. **Headlines MUST be grounded on data.** Never copy headlines from source. Analyze the refreshed chart/table data, write a headline that reflects the current findings. Follow PET deck conventions (see Headline Writing below).
3. **Spatial layout is ground truth for non-connected.** Use group shape labels, nearby text boxes, and table proximity — not headline parsing — to determine which chart shows which segment.
4. **User provides data lineage for non-connected.** The analyst knows their Synapse project/analysis/segments. Claude Code interprets how that data maps to each component.
5. **Preserve formatting, update only data.** Clone + replace_data(). Never reconstruct slides.
6. **Every refresh is auditable.** Stamp data lineage, save mapping JSON, report what changed.

## Connected Path (Connector-Tagged Slides)

For slides with Synapse Connector tags (ReportConfigHash, DataFrameConfigHash, MappingConfig):

```bash
# Full pipeline: clone → dummy → refresh from specs → verify
python tests/test_spec_refresh_pipeline.py --all

# Or stage by stage:
python tests/test_spec_refresh_pipeline.py --stage 1   # clone source → dummy
python tests/test_spec_refresh_pipeline.py --stage 2   # refresh from Synapse via specs
python tests/test_spec_refresh_pipeline.py --stage 3   # verify source ≈ refreshed
```

Key modules:
- `slidegen/deck_reader/tag_reader.py` → `generate_config_specs()` extracts raw configs per component
- `slidegen/synapse_chart_mapper.py` → `pivot_records_to_chart_data()` replicates Connector transformation
- Charts: pivot + replace_data() with formatCode preservation
- Tables: each table's own PivotConfig + MappingConfig → formatted cell values (base sizes, percentages)

## Non-Connected Path (Claude Code Interpretation)

For slides WITHOUT Connector tags. User provides `data_lineage` (project_id, analysis_ids, segment_ids).

### Step 1: Extract slide context

```bash
python -m slidegen.intelligent_refresh read --pptx path/to/slide.pptx --slide 0
```

Output: JSON with all shapes — charts (series, categories, values), tables (headers, sample rows), text boxes, **group shape labels with positions**.

### Step 2: Claude Code interprets

Read the Step 1 output. Determine per-component mappings by spatial proximity:

**What to look for:**
- Group shape labels like "CARDs" at (2.74, 1.69) near Chart 35 at (2.94, 2.44) → Chart 35 shows CARD segment
- Table base sizes to confirm segment (smaller n = one segment, larger = another)
- Series names (L, N, H, IDK) → match to data column values (x_code: L, N, H, 0)
- Table column headers ("Base", "Easy") → map to data fields (base, percentage)

**Write mapping JSON:**
```json
{
  "charts": [
    {"chart_name": "Chart 35", "segment_filter": "CARD",
     "series_column": "x_code",
     "series_name_map": {"L": "L", "N": "N", "H": "H", "IDK": "0"},
     "row_field": "y_label", "value_field": "percentage"}
  ],
  "tables": [
    {"table_name": "Table 25", "segment_filter": "CARD",
     "columns": [
       {"header": "Product", "data_field": "y_label", "format": "string"},
       {"header": "Base", "data_field": "base", "format": "(n = {})"},
       {"header": "", "data_field": null, "format": "spacer"},
       {"header": "Easy", "data_field": "percentage", "format": "{}%",
        "filter": {"x_code": "H"}}
     ]}
  ]
}
```

### Step 3: Refresh

```bash
python -m slidegen.intelligent_refresh refresh \
  --pptx path/to/slide.pptx --output refreshed.pptx \
  --mapping mapping.json --lineage lineage.json
```

### Step 4: Verify

Open the refreshed PPTX, compare visually with source. Check that:
- Chart categories match expected products
- Series values are segment-filtered correctly
- Table base sizes and percentages match the segment
- Formatting preserved (colors, fonts, layout)

## Headline Writing (Both Paths)

After refreshing data, analyze the chart/table values and write a data-grounded headline.

### Process
1. Read the refreshed slide's chart data (categories, series values)
2. Identify the key finding: highest value, biggest delta, segment comparison, rank shift
3. Write headline following PET conventions (≤120 chars, lead with direction, name the driver)
4. Apply:
```bash
python -m slidegen.intelligent_refresh headline --pptx refreshed.pptx --slide 0 --text "headline text"
```

### Headline Patterns (from 3,569 real PET headlines)

| Pattern | Example |
|---|---|
| Delta + driver | "Rybrevant efficacy recall dipped 3pp QoQ, driven by Efficacy-in-1L message" |
| Segment divergence | "Repatha ease of access among CARDs leads PCPs by 20pp (44% vs 24%)" |
| Magnitude threshold | "Half of NSCLC specialists now recall LITE + EP2 (+8pp vs Q4)" |
| Rank shift | "Safety climbs to #2 recalled message, overtaking Convenience" |
| Flat with context | "Recall holds at 45% — no wave-on-wave shift" |
| Comparative | "Repatha vs Tagrisso: same reach, half the unaided recall" |

### Rules
- **Never invent numbers.** Every data point must be verifiable from the slide's chart/table values.
- **Lead with direction.** Up, down, flat, above, below, leads, trails.
- **Name the driver or segment.** "among CARDs" > "among HCPs".
- **No hedging.** Ban: "may suggest," "appears to indicate."
- **≤120 characters.**
- For first iteration (testing on existing decks), ground purely on refreshed data. Client-specific narrative context will be layered in later.

## Decision Rules

| Situation | Response |
|---|---|
| Slide has Connector tags | Connected path — use raw configs |
| Slide has no tags | Non-connected path — Claude Code interprets |
| Mixed deck (some slides tagged, some not) | Per-slide detection; use appropriate path for each |
| User doesn't provide data_lineage for non-connected | Ask: "Which Synapse project/analysis does this slide's data come from?" |
| Fetched data has no matching segment values | Show available segments, ask user to clarify |
| Chart series names don't match any data column | Try fuzzy matching (abbreviations, case-insensitive), report if still unmatched |
| Headline can't be grounded (no clear finding) | Use section title as fallback, flag for manual review |

## Python API

```python
from slidegen.intelligent_refresh import (
    read_slide_context,
    format_slide_for_interpretation,
    format_data_for_interpretation,
    fetch_synapse_data,
    refresh_slide_from_mapping,
    write_headline,
)

# Phase 1: Read
context = read_slide_context("slide.pptx", slide_index=0)
records, df = fetch_synapse_data(data_lineage)
print(format_slide_for_interpretation(context))
print(format_data_for_interpretation(df))

# Phase 2: Refresh (after Claude Code generates mapping)
results = refresh_slide_from_mapping(
    "slide.pptx", "refreshed.pptx", slide_index=0,
    mapping=mapping_dict, data_lineage=data_lineage,
)

# Phase 3: Headline
write_headline("refreshed.pptx", slide_index=0, headline_text="...")
```

## References

- `slidegen/intelligent_refresh.py` — Slide reader + refresh engine
- `slidegen/synapse_chart_mapper.py` — Connector-faithful pivot engine
- `slidegen/deck_reader/tag_reader.py` — Config spec extraction from Connector tags
- `tests/test_spec_refresh_pipeline.py` — Connected path end-to-end test
- `tests/test_intelligent_refresh.py` — Non-connected path test harness
- `.claude/skills/creation/headline-writer/SKILL.md` — Headline conventions reference
