---
name: layout-selector
effort: low
paths: ["slidegen/layout_selector.py"]
description: "Deterministically pick a LAYOUTS key given the component types on a slide. Pure function. Rules come from Apr 15 deck-analysis clusters: 1 chart + 1 table → observed_1chart_1table (145 slides), 1 chart + 2 tables → observed_1chart_2table (110 slides), etc. Trigger when: a slide-plan-generator has decided which components a slide needs and must pick the layout preset. delta_column is structural (pairs with chart), not counted as a table for selection purposes."
---

# layout-selector

Component-type list → LAYOUTS key. Pure deterministic function.

## Cardinal Rules

1. **Deterministic.** Pure function on component types. No LLM, no I/O.
2. **Output must be a key in `LAYOUTS{}`.** If a new layout is needed, add it to `slidegen/pptx_utils/layout.py` first.
3. **`delta_column` is structural.** It always pairs with a chart as part of observed layout presets. Not counted as a separate table component.
4. **Rules derive from real-deck clusters.** Changes must be justified against Apr 15 cluster counts, not hypothetical symmetry.

## API

```python
from slidegen.layout_selector import select_layout, describe_layout

layout_key = select_layout(["chart", "label_table", "delta_column"])     # → "observed_1chart_1table"
layout_key = select_layout(["chart", "label_table", "value_table"])       # → "observed_1chart_2table"
layout_key = select_layout(["chart", "chart", "label_table", "label_table"])  # → "observed_2chart_2table"

description = describe_layout("observed_1chart_1table")
# → "Single chart + label table + delta column (145 real slides)"
```

## Selection rules

| chart_count | table_count (excl. delta_column) | Layout |
|---|---|---|
| 1 | 1 | `observed_1chart_1table` (145 real slides) |
| 1 | 2+ | `observed_1chart_2table` (110 real slides) |
| 2 | 2 | `observed_2chart_2table` (68 real slides) |
| 2 | 0 | `observed_dual_chart_no_table` (44 real slides) |
| 3+ | any | `observed_three_metric_scorecard` (45 real slides) |
| 0 | 1+ | `observed_full_width_table` (139 real slides) |
| 1 | 0 | `single_chart_with_delta` |

## Decision Rules

| Situation | Response |
|---|---|
| Component list is empty | Return `observed_full_width_table` (text-only cover / ES slide) |
| Unknown component type in list | Ignored (validator will catch it at spec level) |
| 3+ tables with 1 chart | Treat as `observed_1chart_2table` (widest match) |

## References

- `slidegen/layout_selector.py` — implementation
- `slidegen/pptx_utils/layout.py` — `LAYOUTS` dict
- `experiments/deck_analysis/outputs/layout_clusters.json` — cluster signature ground truth
- `experiments/deck_analysis/outputs/layout_clusters.md` — human-readable cluster stats
