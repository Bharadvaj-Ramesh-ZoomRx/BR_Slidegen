---
name: viz-selector
effort: low
paths: ["slidegen/viz_selector.py"]
description: "Deterministically pick a chart_pattern for a slide based on metric tag and/or question type. Implements PRD §6.1 hierarchy: Metric tag → Q-type default → UNRESOLVED (HITL). Pure function, no LLM. Trigger when: a slide-plan-generator needs to set the chart_pattern field for a spec. 36 metric tags are pre-mapped from Apr 15 deck-analysis frequencies (Message Recall → bar_clustered_horizontal at 35% observed, Awareness Trend → line_markers_trended at 14%, etc.)."
---

# viz-selector

Metric + Q-type → chart_pattern. Pure deterministic function.

## Cardinal Rules

1. **Deterministic.** Same inputs → same output, always. No LLM calls, no randomness, no I/O.
2. **Return `UNRESOLVED` when uncertain.** Never guess. The sentinel forces the upstream planner to ask the user or surface the ambiguity explicitly.
3. **Output must be in `SUPPORTED_CHART_PATTERNS`.** The implementation asserts this. Any contributor extending the map must add to `SUPPORTED_CHART_PATTERNS` first (in `slidegen/slide_spec/schema.py`).
4. **Frequencies drive defaults.** The 36 metric mappings come from the Apr 15 deck-analysis of 4,354 real PET charts. Contributors changing a mapping must justify against observation, not personal taste.

## API

```python
from slidegen.viz_selector import select_chart_pattern, get_metric_tags, get_mapping_table

pattern = select_chart_pattern(metric="Message Recall")              # → "bar_clustered_horizontal"
pattern = select_chart_pattern(metric="random_unknown")              # → "UNRESOLVED"
pattern = select_chart_pattern(metric=None, q_type="likert")         # → "bar_clustered_horizontal"
pattern = select_chart_pattern(metric=None, q_type=None)             # → "UNRESOLVED"

tags = get_metric_tags()                                             # list[str] — 36 entries
table = get_mapping_table()                                          # full list of {metric, chart_pattern, observed_pct, rationale}
```

## Priority hierarchy

1. **Metric tag** (priority 1) — 36 pre-mapped pharma research metrics
2. **Question type** (priority 2) — likert, ranking, time_series, categorical_composition, etc. → default pattern
3. **UNRESOLVED** (priority 3) — forces HITL via spec-validator rejection

## Key mappings (from `METRIC_TAG_MAP`)

| Metric tag | Chart pattern | Observed |
|---|---|---|
| `message_recall`, `aided_awareness`, `unaided_awareness`, `topic_recall`, etc. | `bar_clustered_horizontal` | 35% |
| `message_effectiveness`, `motivation`, `believability`, `differentiation`, `mbd` | `xy_scatter_abacus` | 17% |
| `rep_performance`, `call_quality`, `attribute_rating` | `xy_scatter_abacus` | 17% |
| `awareness_trend`, `recall_trend`, `reach`, `share_of_voice`, `frequency` | `line_markers_trended` | 14% |
| `likelihood_to_prescribe`, `prescription_intent`, `patient_allocation`, `share_of_mind` | `column_stacked_100_vertical` | 11% |
| `call_to_action`, `branded_close`, `interaction_format` | `bar_stacked_100_horizontal` | 7% |
| `segment_comparison`, `hii_scorecard`, `brand_comparison` | `column_clustered_vertical` | 7% |

Total: 36 mapped metrics covering 88%+ of real PET chart production.

## Decision Rules

| Situation | Response |
|---|---|
| Both `metric` and `q_type` given, metric known | Return metric's mapping (priority 1) |
| Only `metric` given, unknown | Return `UNRESOLVED` (don't fall through to q_type — that's the caller's choice via a second call) |
| Only `q_type` given, known | Return q_type's default |
| Neither given | Return `UNRESOLVED` |
| Case/whitespace/hyphen variations | Normalized internally via `_normalize()` |

## References

- `slidegen/viz_selector.py` — implementation
- `slidegen/slide_spec/schema.py` — `SUPPORTED_CHART_PATTERNS`
- `experiments/deck_analysis/outputs/ACTIONABLE_FINDINGS.md` — source frequencies
- PRD §6.1 — selection hierarchy rationale
