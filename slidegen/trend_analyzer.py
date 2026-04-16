"""
Trend Analyzer
==============

Detect direction, acceleration, inflections, and produce narrative summaries
from time-series metric data. Used by refresh-deck-workflow and deck-audit-workflow
to characterize trends and compare brand vs competitor trajectories.

Pure Python. No external dependencies.

Example
-------
>>> from slidegen.trend_analyzer import analyze_trend, compare_trends
>>> trend = analyze_trend(
...     periods=["Q2 '25", "Q3 '25", "Q4 '25", "Q1 '26"],
...     values=[0.42, 0.46, 0.50, 0.54],
...     metric="Rybrevant unaided awareness",
... )
>>> trend.direction
'up'
>>> trend.monotonic
True
>>> trend.total_change
0.12
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class TrendPoint:
    """A single data point in a time series."""
    period: str   # e.g. "Q1 '26"
    value: float  # the metric value at that period


@dataclass
class TrendSummary:
    """Summary of a single metric's trajectory over time.

    Attributes:
        metric: what is being tracked
        points: ordered time series data
        direction: "up" | "down" | "flat" | "mixed"
        total_change: last value minus first value
        avg_period_change: mean period-over-period change
        inflections: indices where the direction of change reversed
        monotonic: True if trend never reverses direction
        narrative: human-readable summary suitable for headline-writer input
    """
    metric: str
    points: list[TrendPoint]
    direction: str  # "up" | "down" | "flat" | "mixed"
    total_change: float
    avg_period_change: float
    inflections: list[int]
    monotonic: bool
    narrative: str


def _classify_direction(
    total_change: float,
    period_changes: list[float],
    flat_threshold: float,
) -> str:
    """Classify overall trend direction.

    Example
    -------
    >>> _classify_direction(0.12, [0.04, 0.04, 0.04], 0.02)
    'up'
    >>> _classify_direction(0.01, [0.005, 0.005], 0.02)
    'flat'
    """
    if abs(total_change) < flat_threshold:
        return "flat"

    if not period_changes:
        return "flat"

    ups = sum(1 for c in period_changes if c > flat_threshold)
    downs = sum(1 for c in period_changes if c < -flat_threshold)

    if ups > 0 and downs > 0:
        return "mixed"
    if total_change > 0:
        return "up"
    return "down"


def _find_inflections(
    period_changes: list[float],
    flat_threshold: float,
) -> list[int]:
    """Find indices where the direction of period-over-period change reverses.

    An inflection occurs at index i when the sign of change[i] differs from
    change[i-1] (ignoring flat changes within threshold).

    Example
    -------
    >>> _find_inflections([0.04, 0.04, -0.02, 0.03], 0.01)
    [2, 3]
    """
    inflections: list[int] = []
    prev_dir: Optional[str] = None

    for i, change in enumerate(period_changes):
        if change > flat_threshold:
            curr_dir = "up"
        elif change < -flat_threshold:
            curr_dir = "down"
        else:
            continue  # flat segment, no direction assignment

        if prev_dir is not None and curr_dir != prev_dir:
            inflections.append(i)
        prev_dir = curr_dir

    return inflections


def _build_narrative(
    metric: str,
    direction: str,
    total_change: float,
    avg_period_change: float,
    n_periods: int,
    inflections: list[int],
    monotonic: bool,
    period_changes: list[float],
    flat_threshold: float,
) -> str:
    """Produce a human-readable narrative suitable for headline-writer input.

    Example
    -------
    >>> _build_narrative("Awareness", "up", 0.12, 0.04, 4, [], True, [0.04, 0.04, 0.04], 0.02)
    'Up 12pp over 4 waves, steady increase'
    """
    n_waves = n_periods
    total_pp = abs(total_change) * 100

    if direction == "flat":
        return f"Flat over {n_waves} waves ({total_pp:.0f}pp total change)"

    direction_word = {"up": "Up", "down": "Down", "mixed": "Mixed"}.get(direction, direction)

    parts = [f"{direction_word} {total_pp:.0f}pp over {n_waves} waves"]

    if monotonic and len(period_changes) >= 2:
        # Check for acceleration
        first_half = period_changes[: len(period_changes) // 2]
        second_half = period_changes[len(period_changes) // 2 :]
        avg_first = sum(abs(c) for c in first_half) / len(first_half) if first_half else 0
        avg_second = sum(abs(c) for c in second_half) / len(second_half) if second_half else 0

        if avg_second > avg_first * 1.3:
            n_accel = len(second_half)
            parts.append(f"with acceleration in the last {n_accel}")
        elif avg_first > avg_second * 1.3:
            parts.append("with deceleration in recent periods")
        else:
            parts.append("steady increase" if direction == "up" else "steady decline")
    elif not monotonic:
        n_inflections = len(inflections)
        parts.append(f"{n_inflections} direction change{'s' if n_inflections != 1 else ''}")

    return ", ".join(parts)


def analyze_trend(
    periods: list[str],
    values: list[float],
    metric: str,
    flat_threshold_pp: float = 2.0,
) -> TrendSummary:
    """Detect trend direction, acceleration, inflections.

    Produces a narrative summary suitable for headline-writer input.

    Parameters
    ----------
    periods : ordered period labels (earliest first)
    values : metric values corresponding to each period
    metric : label for the metric being tracked
    flat_threshold_pp : changes below this (in percentage points) are "flat"

    Returns
    -------
    TrendSummary with all computed fields

    Example
    -------
    >>> trend = analyze_trend(
    ...     ["Q2 '25", "Q3 '25", "Q4 '25", "Q1 '26"],
    ...     [0.42, 0.46, 0.50, 0.54],
    ...     "Unaided awareness",
    ... )
    >>> trend.direction
    'up'
    >>> trend.total_change
    0.12
    >>> trend.monotonic
    True
    """
    flat_threshold = flat_threshold_pp / 100.0  # convert pp to decimal

    # Edge cases
    if not periods or not values:
        return TrendSummary(
            metric=metric,
            points=[],
            direction="flat",
            total_change=0.0,
            avg_period_change=0.0,
            inflections=[],
            monotonic=True,
            narrative="No data available",
        )

    if len(periods) != len(values):
        n = min(len(periods), len(values))
        periods = periods[:n]
        values = values[:n]

    points = [TrendPoint(period=p, value=v) for p, v in zip(periods, values)]

    if len(points) == 1:
        return TrendSummary(
            metric=metric,
            points=points,
            direction="flat",
            total_change=0.0,
            avg_period_change=0.0,
            inflections=[],
            monotonic=True,
            narrative=f"Single data point ({values[0]:.0%} in {periods[0]})",
        )

    # Period-over-period changes
    period_changes = [values[i] - values[i - 1] for i in range(1, len(values))]
    total_change = values[-1] - values[0]
    avg_period_change = total_change / len(period_changes) if period_changes else 0.0

    direction = _classify_direction(total_change, period_changes, flat_threshold)
    inflections = _find_inflections(period_changes, flat_threshold)
    monotonic = len(inflections) == 0 and direction in ("up", "down", "flat")

    narrative = _build_narrative(
        metric=metric,
        direction=direction,
        total_change=total_change,
        avg_period_change=avg_period_change,
        n_periods=len(points),
        inflections=inflections,
        monotonic=monotonic,
        period_changes=period_changes,
        flat_threshold=flat_threshold,
    )

    return TrendSummary(
        metric=metric,
        points=points,
        direction=direction,
        total_change=total_change,
        avg_period_change=avg_period_change,
        inflections=inflections,
        monotonic=monotonic,
        narrative=narrative,
    )


def compare_trends(
    trend_a: TrendSummary,
    trend_b: TrendSummary,
) -> dict:
    """Compare two trends -- brand vs competitor, segment A vs B over time.

    Returns a dict with gap analysis suitable for headline-writer.

    Parameters
    ----------
    trend_a : first trend (e.g. focal brand)
    trend_b : second trend (e.g. competitor)

    Returns
    -------
    dict with keys: gap_direction, gap_magnitude, is_gap_widening, narrative

    Example
    -------
    >>> a = analyze_trend(["Q1", "Q2", "Q3"], [0.50, 0.55, 0.60], "Brand A awareness")
    >>> b = analyze_trend(["Q1", "Q2", "Q3"], [0.45, 0.44, 0.43], "Brand B awareness")
    >>> result = compare_trends(a, b)
    >>> result["is_gap_widening"]
    True
    >>> result["gap_direction"]
    'a_leads'
    """
    # Edge case: either trend has no data
    if not trend_a.points or not trend_b.points:
        return {
            "gap_direction": "unknown",
            "gap_magnitude": 0.0,
            "is_gap_widening": False,
            "narrative": "Insufficient data for trend comparison",
        }

    # Use overlapping periods if available, else compare latest values
    last_a = trend_a.points[-1].value
    last_b = trend_b.points[-1].value
    first_a = trend_a.points[0].value
    first_b = trend_b.points[0].value

    current_gap = last_a - last_b
    initial_gap = first_a - first_b
    gap_magnitude = abs(current_gap)
    gap_change = abs(current_gap) - abs(initial_gap)

    # Direction
    if current_gap > 0:
        gap_direction = "a_leads"
    elif current_gap < 0:
        gap_direction = "b_leads"
    else:
        gap_direction = "tied"

    # Is the gap widening?
    is_gap_widening = gap_change > 0

    # Build narrative
    gap_pp = abs(current_gap) * 100
    gap_change_pp = gap_change * 100
    leader = trend_a.metric if current_gap >= 0 else trend_b.metric
    trailer = trend_b.metric if current_gap >= 0 else trend_a.metric

    parts = [f"{leader} leads {trailer} by {gap_pp:.0f}pp"]

    if abs(gap_change_pp) < 1:
        parts.append("gap stable")
    elif is_gap_widening:
        parts.append(f"gap widening (+{abs(gap_change_pp):.0f}pp)")
    else:
        parts.append(f"gap narrowing ({abs(gap_change_pp):.0f}pp)")

    # Trajectory comparison
    if trend_a.direction != trend_b.direction:
        parts.append(f"{trend_a.metric} trending {trend_a.direction} while {trend_b.metric} trending {trend_b.direction}")

    narrative = "; ".join(parts)

    return {
        "gap_direction": gap_direction,
        "gap_magnitude": gap_magnitude,
        "is_gap_widening": is_gap_widening,
        "narrative": narrative,
    }
