---
name: segment-comparator
effort: low
paths: ["slidegen/segment_comparator.py"]
description: "Compute per-segment summaries (mean, stdev, SEM, sample size) and pairwise mean deltas across segments. Output feeds stat-sig-annotator + callout-writer + slide-plan-generator-single for segment-analysis workflows (add-slide-workflow, annotate-slide-workflow). Flags segments with n<30 as low-base for downstream annotation."
---

# segment-comparator

Segment data → structured comparison with means, deltas, low-base flags.

## Cardinal Rules

1. **Segment definitions come from the project.** Never invent segment splits. Use segment names from Synapse segment definitions or `config.yaml` segment specs.
2. **Flag low-base segments** (n < 30). Any downstream claim about a low-base segment should be annotated — stat-sig-annotator + callout-writer honor this flag.
3. **Pairwise deltas only.** For 3+ segments, emit all pairwise deltas; don't pre-choose a "winner."
4. **No significance testing here.** Just summary statistics + deltas. Significance is stat-sig-annotator's job.

## API

```python
from slidegen.segment_comparator import (
    compare_segments, compare_two_segments,
    SegmentSummary, SegmentComparison,
)

comparison = compare_segments(
    data={"Academic": [0.58, 0.62, 0.55, ...], "Community": [0.45, 0.48, 0.50, ...]},
    metric="Efficacy recall",
    low_base_threshold=30,
)
# comparison.segments: list[SegmentSummary(name, n, mean, stdev, sem, values)]
# comparison.overall: SegmentSummary across all values
# comparison.deltas: {(seg_a, seg_b): mean_a - mean_b}
# comparison.low_base_segments: list[str]
```

## Decision Rules

| Situation | Response |
|---|---|
| Segment has n=0 | Skip segment; warn |
| All segments have n<30 | Emit comparison but mark every segment low-base; downstream callouts flag entire comparison as low-base |
| All segments have identical means (delta=0) | Emit zero-deltas; stat-sig will return not-significant |
| Single segment provided | Return comparison with `segments=[seg]`, empty deltas, no overall computation |

## References

- `slidegen/segment_comparator.py` — implementation
- `.claude/skills/analysis/stat-sig-annotator/SKILL.md` — consumes this output
- `.claude/skills/creation/callout-writer/SKILL.md` — uses comparison for insight callouts
- `.claude/skills/workflows/add-slide-workflow/SKILL.md`, `annotate-slide-workflow/SKILL.md` — workflow consumers
- PRD §3 Workflow 4, §4.4 analysis skills inventory
