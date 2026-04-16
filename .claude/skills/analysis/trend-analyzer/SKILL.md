---
name: trend-analyzer
effort: low
paths: ["slidegen/trend_analyzer.py"]
description: "Detect trend direction, acceleration, inflection points across waves. Produces a TrendSummary with a human-readable narrative ('Up 12pp over 4 waves with acceleration in the last 2') suitable for headline-writer input on trend-focused slides (line_markers_trended). Also supports comparing two trends (brand vs competitor, segment A vs B over time)."
---

# trend-analyzer

Multi-wave data → TrendSummary with direction, inflections, narrative.

## Cardinal Rules

1. **Flat threshold is 2pp** by default. Changes below `flat_threshold_pp` are considered "flat" and don't drive narrative. User can tighten for more sensitive tracking.
2. **Narrative is factual, not interpretive.** "Up 12pp over 4 waves, acceleration in last 2" — not "Brand is winning in Q1 '26."
3. **Inflections are sign changes in period-to-period deltas.** A trend going up then flat then up has 2 inflections (up→flat, flat→up). Useful for headlines that want to name the inflection point.
4. **Monotonic doesn't mean smooth.** A trend can be monotonic (always increasing) but have varying slope. The `avg_period_change` captures the mean slope.

## API

```python
from slidegen.trend_analyzer import (
    analyze_trend, compare_trends,
    TrendSummary, TrendPoint,
)

summary = analyze_trend(
    periods=["Q2 '25", "Q3 '25", "Q4 '25", "Q1 '26"],
    values=[0.42, 0.46, 0.50, 0.54],
    metric="Unaided awareness",
    flat_threshold_pp=2.0,
)
# summary.direction: "up" | "down" | "flat" | "mixed"
# summary.total_change: last - first (in original units, e.g. 0.12 for 12pp)
# summary.avg_period_change: mean of period-to-period deltas
# summary.inflections: list of indices where direction changed
# summary.monotonic: True if no inflections
# summary.narrative: "Up 12pp over 4 waves, steady increase"

# Compare two trends
cmp = compare_trends(trend_rybrevant, trend_tagrisso)
# cmp["gap_direction"]: "narrowing" | "widening" | "stable"
# cmp["gap_magnitude"]: current-period gap
# cmp["is_gap_widening"]: bool
# cmp["narrative"]: "Rybrevant gap closed from 16pp to 2pp over 4 waves"
```

## Decision Rules

| Situation | Response |
|---|---|
| Fewer than 2 periods | Raise ValueError — need ≥2 points for trend |
| Exactly 2 periods | Direction + total_change; no inflections possible; narrative simple ("Up/Down Xpp from P1 to P2") |
| All values equal | `direction="flat"`, `total_change=0`, `narrative="Held steady across N waves"` |
| compare_trends with mismatched period sets | Align on common periods; warn about discarded ones |
| compare_trends with trends going opposite directions | `gap_direction` based on absolute gap magnitude change; narrative names both directions |

## References

- `slidegen/trend_analyzer.py` — implementation
- `.claude/skills/creation/headline-writer/SKILL.md` — consumes TrendSummary.narrative for trend-chart headlines
- `.claude/skills/workflows/refresh-deck-workflow/SKILL.md` — invokes for trend callouts during multi-wave refresh
- PRD §4.4 analysis skills inventory
