---
name: deck-reader
effort: high
paths: ["slidegen/deck_reader/**/*.py"]
description: "Parse an existing PPTX into a list of SlideSpec. Primary use: retroactive spec generation — the one-time bootstrapping step that makes any existing deck legible to all SlideGen workflows. Dual-mode: Tier 1 reads Galen-PowerPoint Connector tags (ReportConfigHash → Custom XML Part) for canonical Synapse lineage; Tier 2 falls back to structural inference. Trigger when any workflow starts from an existing deck, or explicitly to create the SlideSpec bundle for an existing project (see Pre-condition in refresh-deck-workflow)."
---

# deck-reader

Existing deck → `list[SlideSpec]`. Powers every workflow that starts from a deck (Refresh, Edit, Add, Annotate, Restructure, Audit, Executive Summary).

## Primary Use Case: Retroactive Spec Generation

Before any SlideGen workflow can run on an existing deck, `deck-reader` must be run once to produce the slide specs. The resulting specs + updated `config.yaml` are saved to `projects/{name}/context/{wave}/slide_specs/` and become the persistent representation of the deck in the SlideGen system. This is not a preprocessing step — it's the one-time "register this deck with SlideGen" action.

```
deck-reader("prior_wave.pptx")
→ list[SlideSpec] saved to projects/{name}/context/{wave}/slide_specs/
→ config.yaml updated with extracted extraction params
→ deck is now legible to all 8 SlideGen workflows
```

Run once per deck per wave. Subsequent waves bootstrap from the saved specs, not from re-running deck-reader on the new PPTX.

## Cardinal Rules

1. **Tag existence is necessary but NOT sufficient for Tier 1 trust.** A `ReportConfigHash` shape tag is put through a 5-step health check (below) before being used for refresh. Tags can exist but be wrong — stale lineage after manual chart edits, dead Synapse IDs, structural drift, legacy schema, or corrupted Custom XML pointers. Trusting a bad tag propagates the failure into our refresh and can cause layout-destroying structural alterations (the exact problem galen-powerpoint users hit when refreshing against a misconfigured tag).
2. **Two tiers, not three. No gating on tag failures.** Tag either passes all health checks (Tier 1, use it) OR falls through to Tier 2 inference (structural reconstruction from the chart itself). No user gate on tag issues — gating per-shape in a 40-slide deck is unusable. Unhealthy tag content is preserved in `metadata.original_tag_lineage` for audit but is NOT used for refresh.
3. **Never invent data in Tier 2.** If structural inference can't confidently identify the data source, mark `metadata.confidence="low"` and leave fields empty. The user reviews low-confidence specs before refresh — but that gating is based on inference quality, not on whether a tag was rejected.
4. **Emit valid specs only.** Every spec returned passes `validate_spec(strict=True)`. If a slide can't be reconstructed as a valid spec, skip it with a warning — don't emit a partial spec.
5. **Preserve what's there.** Do not reinterpret headline text, do not rewrite data labels, do not reshape tables. Read what the PPTX says, represent it in the spec.
6. **Tag findings in `metadata`.** Every spec's `metadata` includes `created_by="deck-reader"`, `tier="1"|"2"`, `confidence="high"|"medium"|"low"`. For tagged shapes: `tag_health` records the check results (pass/fail per check); if tag rejected, `original_tag_lineage` preserves what the tag said.

## Inputs / Outputs

**Input:** Path to an existing PPTX.
**Output:** `list[SlideSpec]` — one spec per slide. Plus a summary dict with tagged/untagged counts, and (for Tier 2) a review doc `<deck_name>_inferred_specs.md` listing per-slide inferred values + candidates.

## API

```python
from slidegen.deck_reader import read_deck

specs, summary = read_deck("projects/jnj_rybrevant/input/wave/Q1 2026/prior_wave.pptx",
                           config_path="projects/jnj_rybrevant/config.yaml")  # config is optional

# summary = {
#   "total_slides": 62,
#   "total_shapes": 1180,
#   "tagged_shapes": 0,     # Tier 1 shapes (Connector-authored)
#   "untagged_shapes": 1180, # Tier 2 shapes (structural inference)
#   "by_slide": [...],
# }
```

CLI:
```
python -m slidegen.deck_reader path/to/deck.pptx [--config path/to/config.yaml] [--out specs/]
```

## Tier 1 — Connector tag extraction (with health check)

For shapes carrying Galen-PowerPoint Connector tags:

### Extraction
1. Read shape tag `ReportConfigHash` → SHA256 hex key
2. Look up Custom XML Part entry at `ReportConfig/Entry[{hash}]/Data`
3. Parse `ReportConfigDto` JSON → populate `DataLineage`:
   - `ProjectId` → `data_lineage.project_id`
   - `ReportingPlanId` → `data_lineage.reporting_plan_id`
   - `AnalysisIds[]` → `data_lineage.analysis_ids`
   - `SurveyId` → `data_lineage.survey_id`
   - `SegmentIds[]` → `data_lineage.segment_ids`
   - `StaticTimePeriodIds[]` OR `DynamicTimePeriod{LatestN, IncludeLive}` → deliverable fields
   - `AnalysisType` → `data_lineage.analysis_type`
4. Read `MappingConfig` shape tag → JSON `{field: visual_element}` map (stored in spec's `metadata.mapping_config`)
5. Read `LastRefreshTime`, `RefreshErrorMsgTag`, `ColumnKeyLabelMap` → audit fields

### Tag health check (5 steps)

Each tagged shape goes through these checks. **Any failure → tag rejected, shape routes through Tier 2 inference.** No user gate per shape.

| # | Check | What it catches |
|---|---|---|
| 1 | **Custom XML Part resolves.** Hash points to a valid `ReportConfigDto` JSON. | Corrupted tags / broken relationships |
| 2 | **Synapse IDs still exist.** `reporting_plan_id` + `analysis_ids` + `segment_ids` + `static_time_period_ids` resolve in Synapse today (via `synapse-read` existence check). | Dead references — deleted analyses, removed segments |
| 3 | **Schema compatibility.** DTO parses cleanly into the current Python dataclass — no missing required fields, no unknown critical fields from a legacy Connector schema. | Tags from older Connector versions incompatible with current shape |
| 4 | **Structural consistency.** When data is pulled via the lineage, returned shape (row count, series count, column set) matches the chart's current dimensions within tolerance (±1 row, ±1 column). | Stale lineage after manual chart edits (analyst trimmed rows, renamed series, etc.) |
| 5 | **Mapping plausibility.** `MappingConfig`'s field names appear in the shape's actual data binding; `ColumnKeyLabelMap` keys are present. | Out-of-sync mapping after manual rename / restructure |

### Tier classification

- **Tier 1 (trusted):** all 5 health checks pass → `DataLineage` populated from tag, `metadata.tier="1"`, `confidence="high"`. Refresh proceeds deterministically using the tag.
- **Tier 2 (inference):** any health check fails OR shape is untagged → structural inference from chart content. `metadata.tier="2"`, `confidence` set by inference quality. For failed-tag shapes, `metadata.original_tag_lineage` preserves what the tag said (audit only; not used for refresh). Tier 2's own user-review gate (on `confidence="low"` cases) is the only gating — driven by inference quality, NOT by tag failures.

### User overrides

- `--trust-tags` flag: skip health checks 2-5 (still require check 1 for JSON parseability); use every parseable tag as Tier 1. Fast path when user knows tags are sound and wants to skip the Synapse-existence round-trip.
- `--ignore-tags` flag: skip Tier 1 entirely; every shape goes through Tier 2. Safety path when user knows this deck's tags are broken deck-wide.

## Tier 2 — Structural inference (fallback)

For untagged shapes:

1. Parse slide title / headline textbox → `spec.headline.text`
2. Classify chart OOXML (`barDir` + `grouping` → chart_pattern via heuristics)
3. Read categories from chart `<c:cat>` / label table cells → `ChartData.categories`
4. Read series values + colors → `ChartData.series` (with color tokens `{deck.slide_N.series_M.color}` preserving observed hex)
5. Cross-reference against optional `config.yaml` + `source_data.json` to propose `extraction_method` + `question_codes`
6. Flag confidence:
   - `"high"` — headline clearly matches a known question text + chart pattern is unambiguous
   - `"medium"` — only one of the two matches
   - `"low"` — neither matches; spec has empty `data_source` + `question_codes`

## Decision Rules

| Situation | Response |
|---|---|
| Deck has mixed tagged + untagged shapes | Run health checks per tagged shape; unhealthy tags + untagged shapes all route through Tier 2 |
| Shape has malformed `ReportConfigHash` (no matching XML part) | Health check 1 fails → Tier 2; log rejection reason |
| Shape has tag but Synapse IDs are dead | Health check 2 fails → Tier 2; preserve `metadata.original_tag_lineage` |
| Shape has tag but data shape mismatches chart | Health check 4 fails → Tier 2; preserve tag in audit metadata |
| Shape has tag but MappingConfig references unknown fields | Health check 5 fails → Tier 2; preserve tag in audit metadata |
| Tier 2 inference has 0 confident matches | Emit spec with `confidence="low"`; user review gate (driven by inference quality, not tag status) |
| Slide is a cover/divider/image-only | Emit minimal spec with only `headline` + `metadata`; skip Tier 1/2 chart extraction |
| `validate_spec()` fails on an emitted spec | Drop spec + log warning with which validator errors prevented inclusion |
| User passes `--trust-tags` | Skip health checks 2-5 (still require check 1 for parseability); use tag as Tier 1 |
| User passes `--ignore-tags` | Bypass Tier 1 entirely; every shape goes through Tier 2 |

## References

- `slidegen/deck_reader/tag_reader.py` — Tier 1 implementation
- `slidegen/deck_reader/inference.py` — Tier 2 implementation
- `slidegen/deck_reader/__init__.py` — unified `read_deck()` API
- Galen-PowerPoint Connector: `Docs/Export Import Tags - PRD.md`, `Constants.cs` (shape tag constants), `Services/ShapeConfigurationServiceBase.cs` (hash-based XML part lookup)
- PRD §6.8 — full dual-mode rationale
