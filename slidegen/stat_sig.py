"""
Statistical Significance Testing
=================================

Welch's t-test, chi-square for proportions, and finding annotation.
Used by add-slide-workflow and annotate-slide-workflow to validate
segment comparisons before they appear on client-ready slides.

Pure Python with optional scipy acceleration. If scipy is not installed,
falls back to manual implementations of Welch's t-test and chi-square.

Example
-------
>>> from slidegen.stat_sig import t_test_independent, annotate_finding
>>> result = t_test_independent([0.6, 0.7, 0.65, 0.62], [0.4, 0.45, 0.42, 0.38])
>>> result.significant
True
>>> annotate_finding({
...     "metric": "Rybrevant recall",
...     "segment_a": "Academic", "segment_b": "Community",
...     "delta_pp": 8.0, "n_a": 156, "n_b": 203, "p_value": 0.008,
... })
'Rybrevant recall +8pp higher among Academic (n=156) vs Community (n=203), p<0.01'
"""
from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Optional


@dataclass
class SigTestResult:
    """Result of a statistical significance test.

    Attributes:
        test_name: "t-test" | "chi-square" | "z-proportion"
        statistic: the test statistic value
        p_value: two-sided (or one-sided) p-value
        significant: True if p < alpha
        alpha: significance level (default 0.05)
        annotation: human marker: "*", "**", "n.s.", or "low base (n<30)"
        comparison: human description of the comparison
    """
    test_name: str
    statistic: float
    p_value: float
    significant: bool
    alpha: float
    annotation: str
    comparison: str


def _mean(values: list[float]) -> float:
    """Compute arithmetic mean. Returns 0.0 for empty list."""
    if not values:
        return 0.0
    return sum(values) / len(values)


def _var(values: list[float], ddof: int = 1) -> float:
    """Compute sample variance with ddof correction. Returns 0.0 for n <= ddof."""
    n = len(values)
    if n <= ddof:
        return 0.0
    m = _mean(values)
    return sum((x - m) ** 2 for x in values) / (n - ddof)


def _t_cdf_approx(t: float, df: float) -> float:
    """Approximate the CDF of Student's t-distribution using the regularized
    incomplete beta function approximation.

    For large df (>100), uses the normal approximation. For smaller df, uses
    a series expansion. Accuracy is within ~0.001 for most practical cases.
    """
    if df <= 0:
        return 0.5

    # For large df, use normal approximation
    if df > 100:
        # Normal CDF via error function
        return 0.5 * (1.0 + math.erf(t / math.sqrt(2.0)))

    # Use the regularized incomplete beta function relationship:
    # CDF(t, df) = 1 - 0.5 * I_x(df/2, 1/2) where x = df/(df+t^2)
    # Approximate via continued fraction or series
    x = df / (df + t * t)
    a = df / 2.0
    b = 0.5

    # Simple beta regularized incomplete function via series
    beta_inc = _beta_regularized(x, a, b)

    if t >= 0:
        return 1.0 - 0.5 * beta_inc
    else:
        return 0.5 * beta_inc


def _beta_regularized(x: float, a: float, b: float) -> float:
    """Regularized incomplete beta function I_x(a, b) via continued fraction.

    Uses Lentz's algorithm for the continued fraction expansion.
    """
    if x <= 0.0:
        return 0.0
    if x >= 1.0:
        return 1.0

    # Use the symmetry relation if x > (a+1)/(a+b+2)
    if x > (a + 1.0) / (a + b + 2.0):
        return 1.0 - _beta_regularized(1.0 - x, b, a)

    # Front factor: x^a * (1-x)^b / (a * B(a,b))
    # ln(B(a,b)) = lgamma(a) + lgamma(b) - lgamma(a+b)
    lbeta = math.lgamma(a) + math.lgamma(b) - math.lgamma(a + b)
    front = math.exp(a * math.log(x) + b * math.log(1.0 - x) - lbeta) / a

    # Continued fraction (Lentz's method)
    TINY = 1e-30
    MAX_ITER = 200

    f = 1.0
    c = 1.0
    d = 1.0 - (a + b) * x / (a + 1.0)
    if abs(d) < TINY:
        d = TINY
    d = 1.0 / d
    f = d

    for m in range(1, MAX_ITER + 1):
        # Even step
        numerator = m * (b - m) * x / ((a + 2 * m - 1) * (a + 2 * m))
        d = 1.0 + numerator * d
        if abs(d) < TINY:
            d = TINY
        c = 1.0 + numerator / c
        if abs(c) < TINY:
            c = TINY
        d = 1.0 / d
        f *= c * d

        # Odd step
        numerator = -(a + m) * (a + b + m) * x / ((a + 2 * m) * (a + 2 * m + 1))
        d = 1.0 + numerator * d
        if abs(d) < TINY:
            d = TINY
        c = 1.0 + numerator / c
        if abs(c) < TINY:
            c = TINY
        d = 1.0 / d
        delta = c * d
        f *= delta

        if abs(delta - 1.0) < 1e-10:
            break

    return front * f


def _welch_df(var_a: float, n_a: int, var_b: float, n_b: int) -> float:
    """Welch-Satterthwaite degrees of freedom for unequal-variance t-test."""
    if n_a <= 1 or n_b <= 1:
        return 1.0
    sa = var_a / n_a
    sb = var_b / n_b
    numerator = (sa + sb) ** 2
    denom = (sa ** 2) / (n_a - 1) + (sb ** 2) / (n_b - 1)
    if denom == 0:
        return 1.0
    return numerator / denom


def _annotation_from_p(p_value: float, alpha: float, low_base: bool) -> str:
    """Produce human-readable annotation marker from p-value.

    Example
    -------
    >>> _annotation_from_p(0.008, 0.05, False)
    '**'
    >>> _annotation_from_p(0.03, 0.05, False)
    '*'
    >>> _annotation_from_p(0.12, 0.05, False)
    'n.s.'
    >>> _annotation_from_p(0.03, 0.05, True)
    'low base (n<30)'
    """
    if low_base:
        return "low base (n<30)"
    if p_value < 0.01:
        return "**"
    if p_value < alpha:
        return "*"
    return "n.s."


def _p_label(p_value: float) -> str:
    """Format p-value for client-ready text.

    Example
    -------
    >>> _p_label(0.003)
    'p<0.01'
    >>> _p_label(0.04)
    'p<0.05'
    >>> _p_label(0.12)
    'p=0.12'
    """
    if p_value < 0.01:
        return "p<0.01"
    if p_value < 0.05:
        return "p<0.05"
    return f"p={p_value:.2f}"


def t_test_independent(
    sample_a: list[float],
    sample_b: list[float],
    alpha: float = 0.05,
    two_sided: bool = True,
) -> SigTestResult:
    """Welch's t-test (unequal variances) for two independent samples.

    Uses scipy.stats.ttest_ind if available; otherwise falls back to a manual
    implementation with an approximate t-distribution CDF.

    Parameters
    ----------
    sample_a : values for group A
    sample_b : values for group B
    alpha : significance level (default 0.05)
    two_sided : if True, two-tailed test; if False, one-tailed

    Returns
    -------
    SigTestResult with test_name="t-test"

    Example
    -------
    >>> result = t_test_independent([0.6, 0.7, 0.65], [0.4, 0.45, 0.42])
    >>> result.test_name
    't-test'
    >>> result.significant  # likely True given the separation
    True
    """
    n_a, n_b = len(sample_a), len(sample_b)
    low_base = n_a < 30 or n_b < 30

    # Edge cases
    if n_a == 0 or n_b == 0:
        return SigTestResult(
            test_name="t-test",
            statistic=0.0,
            p_value=1.0,
            significant=False,
            alpha=alpha,
            annotation="low base (n<30)" if low_base else "n.s.",
            comparison="insufficient data (empty sample)",
        )

    if n_a == 1 and n_b == 1:
        return SigTestResult(
            test_name="t-test",
            statistic=0.0,
            p_value=1.0,
            significant=False,
            alpha=alpha,
            annotation="low base (n<30)",
            comparison="insufficient data (n=1 in both groups)",
        )

    mean_a, mean_b = _mean(sample_a), _mean(sample_b)
    var_a, var_b = _var(sample_a), _var(sample_b)
    delta = mean_a - mean_b

    # Zero variance in both groups: means are identical or not, but no test possible
    if var_a == 0 and var_b == 0:
        sig = delta != 0.0
        return SigTestResult(
            test_name="t-test",
            statistic=0.0,
            p_value=0.0 if sig else 1.0,
            significant=sig,
            alpha=alpha,
            annotation=_annotation_from_p(0.0 if sig else 1.0, alpha, low_base),
            comparison=f"{'+' if delta >= 0 else ''}{delta * 100:.1f}pp difference, zero variance in both groups",
        )

    # Try scipy first
    try:
        from scipy import stats as scipy_stats
        t_result = scipy_stats.ttest_ind(
            sample_a, sample_b, equal_var=False,
            alternative="two-sided" if two_sided else "greater",
        )
        t_stat = float(t_result.statistic)
        p_value = float(t_result.pvalue)
        # scipy returns NaN when variance is zero in one group; fall back
        if math.isnan(t_stat) or math.isnan(p_value):
            raise ImportError("scipy returned NaN; falling back to manual")
    except ImportError:
        # Manual Welch's t-test
        se = math.sqrt(var_a / n_a + var_b / n_b) if (var_a / n_a + var_b / n_b) > 0 else 0.0
        if se == 0:
            t_stat = 0.0
            p_value = 1.0
        else:
            t_stat = delta / se
            df = _welch_df(var_a, n_a, var_b, n_b)
            # Two-sided p-value from t-distribution
            cdf_val = _t_cdf_approx(abs(t_stat), df)
            p_value = 2.0 * (1.0 - cdf_val) if two_sided else (1.0 - cdf_val)
            p_value = max(0.0, min(1.0, p_value))

    annotation = _annotation_from_p(p_value, alpha, low_base)
    delta_pp = delta * 100
    sign = "+" if delta_pp >= 0 else ""
    comparison = f"{sign}{delta_pp:.1f}pp difference, {_p_label(p_value)}"

    return SigTestResult(
        test_name="t-test",
        statistic=t_stat,
        p_value=p_value,
        significant=p_value < alpha,
        alpha=alpha,
        annotation=annotation,
        comparison=comparison,
    )


def chi_square_proportions(
    counts_a: tuple[int, int],
    counts_b: tuple[int, int],
    alpha: float = 0.05,
) -> SigTestResult:
    """Chi-square test for 2x2 contingency table (two proportions).

    Parameters
    ----------
    counts_a : (success_count, total_count) for group A
    counts_b : (success_count, total_count) for group B
    alpha : significance level

    Returns
    -------
    SigTestResult with test_name="chi-square"

    Example
    -------
    >>> result = chi_square_proportions((80, 156), (60, 203))
    >>> result.test_name
    'chi-square'
    """
    success_a, total_a = counts_a
    success_b, total_b = counts_b
    low_base = total_a < 30 or total_b < 30

    # Edge cases
    if total_a == 0 or total_b == 0:
        return SigTestResult(
            test_name="chi-square",
            statistic=0.0,
            p_value=1.0,
            significant=False,
            alpha=alpha,
            annotation="low base (n<30)" if low_base else "n.s.",
            comparison="insufficient data (empty group)",
        )

    fail_a = total_a - success_a
    fail_b = total_b - success_b
    total = total_a + total_b

    # Observed counts: [[success_a, fail_a], [success_b, fail_b]]
    observed = [[success_a, fail_a], [success_b, fail_b]]

    # Try scipy first
    try:
        from scipy import stats as scipy_stats
        chi2_result = scipy_stats.chi2_contingency(observed, correction=True)
        chi2_stat = float(chi2_result[0])
        p_value = float(chi2_result[1])
    except ImportError:
        # Manual chi-square with Yates correction for 2x2
        row_totals = [total_a, total_b]
        col_totals = [success_a + success_b, fail_a + fail_b]

        chi2_stat = 0.0
        for i in range(2):
            for j in range(2):
                expected = row_totals[i] * col_totals[j] / total
                if expected == 0:
                    continue
                obs = observed[i][j]
                # Yates correction
                diff = abs(obs - expected) - 0.5
                if diff < 0:
                    diff = 0
                chi2_stat += (diff ** 2) / expected

        # p-value from chi-square with 1 df using normal approximation
        # chi2(1) ~ (Z)^2, so P(chi2 > x) ≈ 2*(1 - Phi(sqrt(x)))
        if chi2_stat > 0:
            z = math.sqrt(chi2_stat)
            p_value = 2.0 * (1.0 - 0.5 * (1.0 + math.erf(z / math.sqrt(2.0))))
            p_value = max(0.0, min(1.0, p_value))
        else:
            p_value = 1.0

    prop_a = success_a / total_a if total_a > 0 else 0.0
    prop_b = success_b / total_b if total_b > 0 else 0.0
    delta_pp = (prop_a - prop_b) * 100
    sign = "+" if delta_pp >= 0 else ""
    annotation = _annotation_from_p(p_value, alpha, low_base)
    comparison = f"{sign}{delta_pp:.1f}pp ({prop_a:.0%} vs {prop_b:.0%}), {_p_label(p_value)}"

    return SigTestResult(
        test_name="chi-square",
        statistic=chi2_stat,
        p_value=p_value,
        significant=p_value < alpha,
        alpha=alpha,
        annotation=annotation,
        comparison=comparison,
    )


def annotate_finding(finding: dict) -> str:
    """Given a finding dict, return a client-ready annotation string.

    Expected keys: metric, segment_a, segment_b, delta_pp, n_a, n_b.
    Optional: p_value.

    Parameters
    ----------
    finding : dict with keys {metric, segment_a, segment_b, delta_pp, n_a, n_b, p_value?}

    Returns
    -------
    Client-ready annotation string.

    Example
    -------
    >>> annotate_finding({
    ...     "metric": "Rybrevant recall",
    ...     "segment_a": "Academic", "segment_b": "Community",
    ...     "delta_pp": 8.0, "n_a": 156, "n_b": 203, "p_value": 0.008,
    ... })
    'Rybrevant recall +8pp higher among Academic (n=156) vs Community (n=203), p<0.01'

    >>> annotate_finding({
    ...     "metric": "Safety awareness",
    ...     "segment_a": "Academic", "segment_b": "Community",
    ...     "delta_pp": -3.0, "n_a": 22, "n_b": 203,
    ... })
    'Safety awareness -3pp lower among Academic (n=22) vs Community (n=203) (low base: n=22 in Academic)'
    """
    metric = finding.get("metric", "Metric")
    seg_a = finding.get("segment_a", "Group A")
    seg_b = finding.get("segment_b", "Group B")
    delta_pp = finding.get("delta_pp", 0.0)
    n_a = finding.get("n_a", 0)
    n_b = finding.get("n_b", 0)
    p_value = finding.get("p_value")

    # Direction word
    if delta_pp > 0:
        direction = "higher"
        sign = "+"
    elif delta_pp < 0:
        direction = "lower"
        sign = ""  # negative sign is inherent
    else:
        direction = "equal"
        sign = ""

    delta_str = f"{sign}{delta_pp:.0f}pp"

    base = f"{metric} {delta_str} {direction} among {seg_a} (n={n_a}) vs {seg_b} (n={n_b})"

    # Low base check
    low_base_segs = []
    if n_a < 30:
        low_base_segs.append(f"n={n_a} in {seg_a}")
    if n_b < 30:
        low_base_segs.append(f"n={n_b} in {seg_b}")

    if low_base_segs:
        return f"{base} (low base: {'; '.join(low_base_segs)})"

    # p-value annotation
    if p_value is not None:
        return f"{base}, {_p_label(p_value)}"

    return base
