"""
Segment Comparator
==================

Compute per-segment summaries and pairwise mean deltas from raw survey data.
Used by add-slide-workflow and annotate-slide-workflow when segment comparisons
are requested.

Pure Python. No external dependencies.

Example
-------
>>> from slidegen.segment_comparator import compare_segments, compare_two_segments
>>> comparison = compare_segments(
...     data={"Academic": [0.58, 0.62, 0.55], "Community": [0.45, 0.48, 0.50]},
...     metric="Efficacy recall",
... )
>>> comparison.segments[0].mean  # Academic mean
0.5833333333333334
>>> comparison.deltas[("Academic", "Community")]
0.13333333333333336
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class SegmentSummary:
    """Statistical summary for one segment."""
    name: str
    n: int
    mean: float
    stdev: float
    sem: float  # standard error of the mean
    values: list[float]  # raw values if available


@dataclass
class SegmentComparison:
    """Result of comparing multiple segments on a single metric.

    Attributes:
        metric: what is being compared (e.g. "Efficacy recall")
        segments: per-segment statistical summaries
        overall: aggregate summary across all segments (None if no data)
        deltas: pairwise mean differences {(seg_a, seg_b): mean_a - mean_b}
        low_base_segments: segment names with n < threshold
    """
    metric: str
    segments: list[SegmentSummary]
    overall: Optional[SegmentSummary]
    deltas: dict[tuple[str, str], float]
    low_base_segments: list[str]


def _compute_summary(name: str, values: list[float]) -> SegmentSummary:
    """Compute summary statistics for a single segment.

    Handles edge cases: empty list (n=0), single value (stdev=0), identical values.

    Example
    -------
    >>> s = _compute_summary("Acad", [0.5, 0.6, 0.7])
    >>> s.n
    3
    >>> round(s.mean, 4)
    0.6
    """
    n = len(values)
    if n == 0:
        return SegmentSummary(
            name=name, n=0, mean=0.0, stdev=0.0, sem=0.0, values=[]
        )

    mean = sum(values) / n

    if n == 1:
        return SegmentSummary(
            name=name, n=1, mean=mean, stdev=0.0, sem=0.0, values=list(values)
        )

    # Sample standard deviation (ddof=1)
    variance = sum((x - mean) ** 2 for x in values) / (n - 1)
    stdev = math.sqrt(variance)
    sem = stdev / math.sqrt(n)

    return SegmentSummary(
        name=name, n=n, mean=mean, stdev=stdev, sem=sem, values=list(values)
    )


def compare_segments(
    data: dict[str, list[float]],
    metric: str,
    low_base_threshold: int = 30,
) -> SegmentComparison:
    """Compute per-segment summaries + pairwise mean deltas.

    Returns a SegmentComparison ready for stat_sig_annotator to annotate.

    Parameters
    ----------
    data : dict mapping segment name to list of numeric values
    metric : label for what is being compared
    low_base_threshold : segments with n < this are flagged

    Example
    -------
    >>> comp = compare_segments(
    ...     {"Academic": [0.58, 0.62, 0.55], "Community": [0.45, 0.48, 0.50]},
    ...     metric="Efficacy recall",
    ... )
    >>> len(comp.segments)
    2
    >>> comp.low_base_segments  # both n=3 < 30
    ['Academic', 'Community']
    """
    segments: list[SegmentSummary] = []
    for seg_name, seg_values in data.items():
        segments.append(_compute_summary(seg_name, seg_values))

    # Overall aggregate: flatten all values
    all_values: list[float] = []
    for seg_values in data.values():
        all_values.extend(seg_values)
    overall = _compute_summary("Overall", all_values) if all_values else None

    # Low base detection
    low_base_segments = [s.name for s in segments if s.n < low_base_threshold]

    # Pairwise deltas
    deltas: dict[tuple[str, str], float] = {}
    seg_names = list(data.keys())
    for i in range(len(seg_names)):
        for j in range(i + 1, len(seg_names)):
            name_a, name_b = seg_names[i], seg_names[j]
            mean_a = segments[i].mean
            mean_b = segments[j].mean
            deltas[(name_a, name_b)] = mean_a - mean_b

    return SegmentComparison(
        metric=metric,
        segments=segments,
        overall=overall,
        deltas=deltas,
        low_base_segments=low_base_segments,
    )


def compare_two_segments(
    a_values: list[float],
    b_values: list[float],
    metric: str,
    name_a: str = "Segment A",
    name_b: str = "Segment B",
    low_base_threshold: int = 30,
) -> SegmentComparison:
    """Convenience wrapper: compare exactly two segments.

    Example
    -------
    >>> comp = compare_two_segments(
    ...     [0.58, 0.62, 0.55],
    ...     [0.45, 0.48, 0.50],
    ...     metric="Efficacy recall",
    ...     name_a="Academic",
    ...     name_b="Community",
    ... )
    >>> round(comp.deltas[("Academic", "Community")], 4)
    0.1333
    """
    return compare_segments(
        data={name_a: a_values, name_b: b_values},
        metric=metric,
        low_base_threshold=low_base_threshold,
    )
