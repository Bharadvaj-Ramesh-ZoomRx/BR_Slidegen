"""
Smoke tests for Track 3 analysis modules:
  - slidegen.segment_comparator
  - slidegen.stat_sig
  - slidegen.trend_analyzer

Run: python -m pytest tests/track3/test_analysis_skills.py -v
"""
import math
import sys
import os

# Ensure repo root is on sys.path
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

import pytest

from slidegen.segment_comparator import (
    compare_segments,
    compare_two_segments,
    SegmentComparison,
    SegmentSummary,
)
from slidegen.stat_sig import (
    t_test_independent,
    chi_square_proportions,
    annotate_finding,
    SigTestResult,
)
from slidegen.trend_analyzer import (
    analyze_trend,
    compare_trends,
    TrendSummary,
    TrendPoint,
)


# ─────────────────────────────────────────────────────────────────────────────
# segment_comparator
# ─────────────────────────────────────────────────────────────────────────────


class TestSegmentComparatorEmpty:
    """Edge case: empty input."""

    def test_empty_data_dict(self):
        comp = compare_segments({}, "Efficacy recall")
        assert comp.segments == []
        assert comp.overall is None
        assert comp.deltas == {}
        assert comp.low_base_segments == []

    def test_single_segment_empty_values(self):
        comp = compare_segments({"Acad": []}, "Efficacy recall")
        assert len(comp.segments) == 1
        assert comp.segments[0].n == 0
        assert comp.segments[0].mean == 0.0
        assert comp.deltas == {}


class TestSegmentComparatorSingle:
    """Single segment -- no pairwise deltas."""

    def test_single_segment(self):
        comp = compare_segments({"Acad": [0.5, 0.6, 0.7]}, "Recall")
        assert len(comp.segments) == 1
        assert comp.segments[0].n == 3
        assert abs(comp.segments[0].mean - 0.6) < 1e-9
        assert comp.deltas == {}
        assert comp.low_base_segments == ["Acad"]  # n=3 < 30

    def test_single_value_segment(self):
        comp = compare_segments({"X": [0.42]}, "Metric")
        assert comp.segments[0].n == 1
        assert comp.segments[0].stdev == 0.0
        assert comp.segments[0].sem == 0.0


class TestSegmentComparatorTwoSegments:
    """Two segments with known means."""

    def test_known_means(self):
        comp = compare_two_segments(
            a_values=[0.58, 0.62, 0.55],
            b_values=[0.45, 0.48, 0.50],
            metric="Efficacy recall",
            name_a="Academic",
            name_b="Community",
        )
        assert len(comp.segments) == 2
        acad = comp.segments[0]
        comm = comp.segments[1]
        assert abs(acad.mean - 0.5833333) < 1e-4
        assert abs(comm.mean - 0.4766667) < 1e-4
        delta = comp.deltas[("Academic", "Community")]
        assert abs(delta - (acad.mean - comm.mean)) < 1e-9
        assert delta > 0  # Academic higher

    def test_overall_computed(self):
        comp = compare_segments(
            {"A": [0.5, 0.6], "B": [0.7, 0.8]},
            "Test",
        )
        assert comp.overall is not None
        assert comp.overall.n == 4
        assert abs(comp.overall.mean - 0.65) < 1e-9

    def test_low_base_threshold(self):
        comp = compare_segments(
            {"Big": list(range(50)), "Small": [1, 2, 3]},
            "Test",
            low_base_threshold=30,
        )
        assert "Small" in comp.low_base_segments
        assert "Big" not in comp.low_base_segments


# ─────────────────────────────────────────────────────────────────────────────
# stat_sig
# ─────────────────────────────────────────────────────────────────────────────


class TestTTestEmpty:
    """Edge cases for t_test_independent."""

    def test_empty_samples(self):
        result = t_test_independent([], [])
        assert result.p_value == 1.0
        assert result.significant is False

    def test_one_empty(self):
        result = t_test_independent([0.5, 0.6], [])
        assert result.p_value == 1.0
        assert result.significant is False

    def test_single_values(self):
        result = t_test_independent([0.5], [0.4])
        assert result.annotation == "low base (n<30)"


class TestTTestKnownMeans:
    """Two groups with clear separation."""

    def test_significant_difference(self):
        # Large separation, decent n -- should be significant
        a = [0.6, 0.65, 0.7, 0.62, 0.68, 0.64, 0.66, 0.63]
        b = [0.4, 0.42, 0.38, 0.45, 0.41, 0.39, 0.43, 0.44]
        result = t_test_independent(a, b)
        assert result.test_name == "t-test"
        assert result.p_value < 0.05
        # Still low base (n=8 < 30), so annotation is "low base"
        assert result.annotation == "low base (n<30)"

    def test_identical_means(self):
        a = [0.5, 0.5, 0.5, 0.5]
        b = [0.5, 0.5, 0.5, 0.5]
        result = t_test_independent(a, b)
        assert result.p_value >= 0.05
        assert result.significant is False

    def test_result_fields(self):
        result = t_test_independent([1.0, 2.0, 3.0], [4.0, 5.0, 6.0])
        assert isinstance(result, SigTestResult)
        assert result.test_name == "t-test"
        assert isinstance(result.statistic, float)
        assert isinstance(result.p_value, float)
        assert isinstance(result.annotation, str)
        assert isinstance(result.comparison, str)


class TestChiSquare:
    """Chi-square for proportions."""

    def test_significant_proportions(self):
        # 80/156 vs 60/203 -- different proportions
        result = chi_square_proportions((80, 156), (60, 203))
        assert result.test_name == "chi-square"
        assert result.p_value < 0.05
        assert result.significant is True

    def test_empty_group(self):
        result = chi_square_proportions((0, 0), (5, 10))
        assert result.p_value == 1.0
        assert result.significant is False

    def test_identical_proportions(self):
        result = chi_square_proportions((50, 100), (50, 100))
        assert result.p_value >= 0.05


class TestAnnotateFinding:
    """Client-ready annotation strings."""

    def test_with_pvalue(self):
        text = annotate_finding({
            "metric": "Rybrevant recall",
            "segment_a": "Academic",
            "segment_b": "Community",
            "delta_pp": 8.0,
            "n_a": 156,
            "n_b": 203,
            "p_value": 0.008,
        })
        assert "Rybrevant recall" in text
        assert "+8pp" in text
        assert "Academic (n=156)" in text
        assert "Community (n=203)" in text
        assert "p<0.01" in text

    def test_low_base(self):
        text = annotate_finding({
            "metric": "Safety awareness",
            "segment_a": "Academic",
            "segment_b": "Community",
            "delta_pp": -3.0,
            "n_a": 22,
            "n_b": 203,
        })
        assert "low base" in text
        assert "n=22" in text

    def test_no_pvalue(self):
        text = annotate_finding({
            "metric": "Test",
            "segment_a": "A",
            "segment_b": "B",
            "delta_pp": 5.0,
            "n_a": 100,
            "n_b": 100,
        })
        assert "p<" not in text
        assert "p=" not in text


# ─────────────────────────────────────────────────────────────────────────────
# trend_analyzer
# ─────────────────────────────────────────────────────────────────────────────


class TestTrendAnalyzerEmpty:
    """Edge cases."""

    def test_empty_input(self):
        trend = analyze_trend([], [], "Awareness")
        assert trend.direction == "flat"
        assert trend.total_change == 0.0
        assert trend.narrative == "No data available"

    def test_single_point(self):
        trend = analyze_trend(["Q1 '26"], [0.54], "Awareness")
        assert trend.direction == "flat"
        assert trend.total_change == 0.0
        assert "Single data point" in trend.narrative
        assert "54%" in trend.narrative

    def test_mismatched_lengths(self):
        trend = analyze_trend(
            ["Q1", "Q2", "Q3"],
            [0.5, 0.6],
            "Test",
        )
        assert len(trend.points) == 2


class TestTrendAnalyzerUp:
    """Monotonic upward trend."""

    def test_steady_increase(self):
        trend = analyze_trend(
            ["Q2 '25", "Q3 '25", "Q4 '25", "Q1 '26"],
            [0.42, 0.46, 0.50, 0.54],
            "Unaided awareness",
        )
        assert trend.direction == "up"
        assert abs(trend.total_change - 0.12) < 1e-9
        assert abs(trend.avg_period_change - 0.04) < 1e-9
        assert trend.inflections == []
        assert trend.monotonic is True
        assert "Up" in trend.narrative
        assert "12pp" in trend.narrative

    def test_accelerating(self):
        trend = analyze_trend(
            ["Q1", "Q2", "Q3", "Q4"],
            [0.40, 0.42, 0.47, 0.55],
            "Recall",
        )
        assert trend.direction == "up"
        assert trend.monotonic is True
        assert "acceleration" in trend.narrative.lower()


class TestTrendAnalyzerDown:
    """Monotonic downward trend."""

    def test_steady_decline(self):
        trend = analyze_trend(
            ["Q1", "Q2", "Q3"],
            [0.60, 0.55, 0.50],
            "Competitor share",
        )
        assert trend.direction == "down"
        assert trend.total_change < 0
        assert trend.monotonic is True


class TestTrendAnalyzerMixed:
    """Non-monotonic: direction changes."""

    def test_up_then_down(self):
        trend = analyze_trend(
            ["Q1", "Q2", "Q3", "Q4"],
            [0.40, 0.50, 0.55, 0.42],
            "Awareness",
        )
        assert trend.direction in ("mixed", "flat", "up")  # depends on net change vs threshold
        assert len(trend.inflections) >= 1
        assert trend.monotonic is False


class TestTrendAnalyzerFlat:
    """Flat trend within threshold."""

    def test_flat_within_threshold(self):
        trend = analyze_trend(
            ["Q1", "Q2", "Q3"],
            [0.50, 0.51, 0.50],
            "Metric",
            flat_threshold_pp=2.0,
        )
        assert trend.direction == "flat"
        assert trend.monotonic is True


class TestCompareTrends:
    """Trend comparison."""

    def test_gap_widening(self):
        trend_a = analyze_trend(
            ["Q1", "Q2", "Q3"],
            [0.50, 0.55, 0.60],
            "Brand A",
        )
        trend_b = analyze_trend(
            ["Q1", "Q2", "Q3"],
            [0.45, 0.44, 0.43],
            "Brand B",
        )
        result = compare_trends(trend_a, trend_b)
        assert result["gap_direction"] == "a_leads"
        assert result["is_gap_widening"] is True
        assert result["gap_magnitude"] > 0

    def test_empty_trends(self):
        trend_a = analyze_trend([], [], "A")
        trend_b = analyze_trend([], [], "B")
        result = compare_trends(trend_a, trend_b)
        assert result["gap_direction"] == "unknown"
        assert "Insufficient" in result["narrative"]

    def test_tied(self):
        trend_a = analyze_trend(["Q1", "Q2"], [0.50, 0.55], "A")
        trend_b = analyze_trend(["Q1", "Q2"], [0.50, 0.55], "B")
        result = compare_trends(trend_a, trend_b)
        assert result["gap_direction"] == "tied"
        assert result["gap_magnitude"] == 0.0
