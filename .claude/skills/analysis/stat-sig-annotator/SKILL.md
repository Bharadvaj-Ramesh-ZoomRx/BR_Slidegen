---
name: stat-sig-annotator
effort: low
paths: ["slidegen/stat_sig.py"]
description: "Run significance tests (Welch's t-test for means, chi-square for proportions) and emit client-ready annotations (*, **, n.s., low base). Used by add-slide-workflow + annotate-slide-workflow for segment comparisons. Provides scipy-backed stats with pure-Python fallback (Welch's t-test + Yates-corrected chi-square) so the module never errors on missing scipy."
---

# stat-sig-annotator

Statistical testing + client-ready annotations. Welch's t-test (means), chi-square (proportions).

## Cardinal Rules

1. **Welch's t-test by default** (unequal variances). Not Student's. Real segment comparisons rarely have equal variance.
2. **Significance markers follow real-deck convention**: `*` = p<0.05, `**` = p<0.01, `n.s.` = not significant, `low base (n<30)` overrides any sig marker when sample is small.
3. **Low base beats significance.** An n<30 claim is flagged as low-base even if p<0.01 — the statistical test may be unreliable at that sample size. Downstream callout-writer + slide-creator honor the low-base flag.
4. **Scipy optional.** Stats functions fall back to pure-Python Welch's t-test + chi-square if scipy isn't installed.
5. **Never run two-tailed + one-tailed mixed.** Caller specifies `two_sided` (default True); we don't mix.

## API

```python
from slidegen.stat_sig import (
    t_test_independent, chi_square_proportions,
    annotate_finding, SigTestResult,
)

# Welch's t-test — for mean comparison (recall %, attribute ratings, etc.)
result = t_test_independent(
    sample_a=[0.60, 0.70, 0.65, 0.62, ...],
    sample_b=[0.40, 0.45, 0.42, 0.38, ...],
    alpha=0.05, two_sided=True,
)
# result.significant, result.p_value, result.annotation ("*" / "**" / "n.s." / "low base")
# result.comparison — human string: "Group A vs Group B: +20pp (p<0.001)"

# Chi-square — for proportion comparison (success rates)
result = chi_square_proportions(
    counts_a=(156, 312),   # (success, total) for group A
    counts_b=(124, 298),   # (success, total) for group B
)

# Compose a client-ready annotation from structured inputs
text = annotate_finding({
    "metric": "Rybrevant recall",
    "segment_a": "Academic", "segment_b": "Community",
    "delta_pp": 8.0, "n_a": 156, "n_b": 203, "p_value": 0.008,
})
# → "Rybrevant recall +8pp higher among Academic (n=156) vs Community (n=203), p<0.01"
```

## Decision Rules

| Situation | Response |
|---|---|
| Either sample has n<30 | `annotation="low base (n<30)"`, surface warning in callout |
| Either sample has n=0 | Raise ValueError — no test possible |
| Both samples have identical values (zero variance, means equal) | Return `SigTestResult(p_value=1.0, significant=False, annotation="n.s.")` — handled before scipy call to avoid NaN |
| scipy not installed | Pure-Python fallback (Welch's t-test + Yates-corrected chi-square). Results are within 1e-6 of scipy's. |

## References

- `slidegen/stat_sig.py` — implementation
- `.claude/skills/analysis/segment-comparator/SKILL.md` — typical input source
- `.claude/skills/creation/callout-writer/SKILL.md` — consumes `SigTestResult` for significance markers
- `.claude/skills/workflows/add-slide-workflow/SKILL.md`, `annotate-slide-workflow/SKILL.md` — workflow consumers
- PRD §3 Workflows 4 & 5, §4.4 analysis skills inventory
