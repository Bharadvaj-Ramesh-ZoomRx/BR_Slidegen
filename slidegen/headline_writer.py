"""
headline_writer.py — data-driven headline generation.

Produces punchy, delta-forward headlines for slides. The PRD (§6.9) requires
that when slide data changes, the headline be regenerated from the new data —
a headline like "dipped 3pp QoQ" becomes a delivery risk when new data says
"up 2pp".

Two implementations:
1. **Template-based** (fast, deterministic) — picks a pattern based on
   metric type + observed change, fills in values. Used as the default.
2. **LLM-augmented** (richer narrative framing) — hook to call an LLM with
   the computed template as scaffolding. Not required for correctness; used
   when richer language is wanted.

Style rules (from deck-analysis of 3,569 real PET headlines):
- ≤120 characters
- Lead with direction (up / down / flat)
- Name the driver or segment when possible
- No hedging language ("may suggest", "appears to")
- Never invent numbers

Contract (match PRD §6.7 spec contract):
    generate_headline(spec, new_data?, narrative=?) -> str
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from slidegen.slide_spec import SlideSpec, ChartComponent


# ─────────────────────────────────────────────────────────────────────────────
# Style constants
# ─────────────────────────────────────────────────────────────────────────────

MAX_HEADLINE_CHARS = 120

# Small deltas below this threshold are considered "flat" / not worth naming
FLAT_DELTA_PP = 2.0

# Direction language
_DIR_UP = ("rose", "climbed", "ticked up", "gained")
_DIR_DOWN = ("dipped", "slipped", "fell", "softened")
_DIR_FLAT = ("held steady", "was flat", "stayed level")


# ─────────────────────────────────────────────────────────────────────────────
# Data extraction helpers
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class ChartSummary:
    """Extracted from the chart component of a spec — the signal a headline needs."""
    metric: str                         # inferred from section or headline hint
    brand: Optional[str]
    categories: list[str]
    prior_values: list[float]           # if 2-series with prior/current semantics
    current_values: list[float]
    deltas_pp: list[float]              # current - prior, in percentage points
    max_delta_category: Optional[str]   # category with largest absolute delta
    max_delta_pp: float                 # the largest absolute delta
    direction: str                      # "up" | "down" | "flat" | "mixed"
    all_flat: bool                      # every category's delta < FLAT_DELTA_PP
    all_same_direction: bool
    top_value: float                    # peak current value
    top_value_category: str


def _summarize_chart(spec: SlideSpec) -> Optional[ChartSummary]:
    """Walk the spec's first ChartComponent and extract headline-relevant signal."""
    chart = next((c for c in spec.components if isinstance(c, ChartComponent)), None)
    if chart is None or not chart.data.series:
        return None

    categories = list(chart.data.categories)
    if len(chart.data.series) >= 2:
        # Convention: series[0] = prior, series[-1] = current
        prior = [float(v) for v in chart.data.series[0].values]
        current = [float(v) for v in chart.data.series[-1].values]
    else:
        current = [float(v) for v in chart.data.series[0].values]
        prior = [0.0] * len(current)

    # Align lengths defensively
    n = min(len(categories), len(prior), len(current))
    categories = categories[:n]
    prior = prior[:n]
    current = current[:n]

    # Values in the spec are typically 0-1 decimals; deltas in percentage points (× 100)
    deltas_pp = [round((current[i] - prior[i]) * 100, 1) for i in range(n)]

    max_delta_idx = max(range(n), key=lambda i: abs(deltas_pp[i])) if n else None
    max_delta_category = categories[max_delta_idx] if max_delta_idx is not None else None
    max_delta_pp = deltas_pp[max_delta_idx] if max_delta_idx is not None else 0.0

    # Aggregate direction
    up_count = sum(1 for d in deltas_pp if d > FLAT_DELTA_PP)
    down_count = sum(1 for d in deltas_pp if d < -FLAT_DELTA_PP)
    if up_count and not down_count:
        direction = "up"
    elif down_count and not up_count:
        direction = "down"
    elif up_count == 0 and down_count == 0:
        direction = "flat"
    else:
        direction = "mixed"

    all_flat = all(abs(d) < FLAT_DELTA_PP for d in deltas_pp)
    all_same_direction = direction in ("up", "down", "flat")

    top_idx = max(range(n), key=lambda i: current[i]) if n else None
    top_value = current[top_idx] if top_idx is not None else 0.0
    top_value_category = categories[top_idx] if top_idx is not None else ""

    # Metric inference: use section if available, else "the metric"
    metric = spec.section or "recall"

    return ChartSummary(
        metric=metric,
        brand=spec.brand,
        categories=categories,
        prior_values=prior,
        current_values=current,
        deltas_pp=deltas_pp,
        max_delta_category=max_delta_category,
        max_delta_pp=max_delta_pp,
        direction=direction,
        all_flat=all_flat,
        all_same_direction=all_same_direction,
        top_value=top_value,
        top_value_category=top_value_category,
    )


# ─────────────────────────────────────────────────────────────────────────────
# Headline templates
# ─────────────────────────────────────────────────────────────────────────────


def _brand_phrase(brand: Optional[str]) -> str:
    """Human-readable brand reference. Uppercase normalized keys are turned into
    Title Case for natural reading."""
    if not brand:
        return ""
    # RYBREVANT -> Rybrevant
    return brand.title().replace("_", " + ")


def _fmt_pp(value_pp: float) -> str:
    """Format a pp delta as '+3pp' or '-2pp'."""
    sign = "+" if value_pp >= 0 else ""
    return f"{sign}{int(round(value_pp))}pp"


def _template_delta_driver(s: ChartSummary) -> str:
    """Pattern: '<Brand> <metric> <direction> <delta>pp QoQ, driven by <category>'."""
    brand = _brand_phrase(s.brand)
    verb = _DIR_UP[0] if s.max_delta_pp > 0 else _DIR_DOWN[0]
    brand_prefix = f"{brand} " if brand else ""
    metric = s.metric.lower()
    delta_str = _fmt_pp(s.max_delta_pp)
    return f"{brand_prefix}{metric} {verb} {delta_str} QoQ, driven by {s.max_delta_category}"


def _template_rank_top(s: ChartSummary) -> str:
    """Pattern: '<Category> tops <metric> recall at <value>%, +<delta>pp QoQ'."""
    brand = _brand_phrase(s.brand)
    top_pct = int(round(s.top_value * 100))
    top_idx = s.categories.index(s.top_value_category) if s.top_value_category in s.categories else 0
    top_delta = s.deltas_pp[top_idx] if 0 <= top_idx < len(s.deltas_pp) else 0.0
    brand_part = f" for {brand}" if brand else ""
    return f"{s.top_value_category} leads {s.metric}{brand_part} at {top_pct}% ({_fmt_pp(top_delta)} QoQ)"


def _template_flat(s: ChartSummary) -> str:
    brand = _brand_phrase(s.brand)
    brand_prefix = f"{brand} " if brand else ""
    return f"{brand_prefix}{s.metric.lower()} holds steady QoQ — no meaningful wave-on-wave shift"


def _template_mixed(s: ChartSummary) -> str:
    """Pattern: '<Brand> <metric> shows mixed movement: <max-up> up, <max-down> down'."""
    brand = _brand_phrase(s.brand)
    ups = [(c, d) for c, d in zip(s.categories, s.deltas_pp) if d > FLAT_DELTA_PP]
    downs = [(c, d) for c, d in zip(s.categories, s.deltas_pp) if d < -FLAT_DELTA_PP]
    if not ups or not downs:
        return _template_delta_driver(s)
    top_up = max(ups, key=lambda kv: kv[1])
    top_down = min(downs, key=lambda kv: kv[1])
    brand_prefix = f"{brand} " if brand else ""
    return (
        f"{brand_prefix}{s.metric.lower()}: {top_up[0]} "
        f"{_fmt_pp(top_up[1])}, {top_down[0]} {_fmt_pp(top_down[1])}"
    )


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────


class HeadlineGenerationError(ValueError):
    """Raised when the spec doesn't provide enough signal to generate a headline."""


def generate_headline(
    spec: SlideSpec,
    narrative_hint: Optional[str] = None,
    prior_headline: Optional[str] = None,
    max_chars: int = MAX_HEADLINE_CHARS,
) -> str:
    """Generate a data-driven headline for a slide.

    Strategy (template-based; LLM augmentation can be added as a later wrapper):
      1. Summarize the chart data (direction, deltas, top category).
      2. Pick a template based on the summary's shape:
         - All flat → "holds steady" template
         - Mixed direction → "mixed movement" template
         - Single direction with clear driver → "<driver>-led" template
         - Outstanding top category → "top value" template
      3. Truncate to max_chars if needed (prefers preserving key data).

    Args:
        spec: the SlideSpec (should have at least one ChartComponent).
        narrative_hint: optional arc/theme from narrative_threads.md (e.g.
            "ACT NOW: Efficacy drift") — used to bias template choice.
        prior_headline: the previous wave's headline — if the data is
            effectively identical, this can be reused. Otherwise regenerated.
        max_chars: max length; default 120 per real-deck style.

    Returns:
        A single-line headline string.

    Raises:
        HeadlineGenerationError: if no chart is present and no narrative_hint.
    """
    summary = _summarize_chart(spec)

    # Fallback: no chart but slide has a section — use narrative hint or simple section-based headline
    if summary is None:
        if narrative_hint:
            return narrative_hint[:max_chars]
        if spec.section:
            return f"{spec.section}"[:max_chars]
        raise HeadlineGenerationError(
            "spec has no chart data and no narrative_hint — cannot generate headline"
        )

    # Choose a template based on the summary shape
    if summary.all_flat:
        headline = _template_flat(summary)
    elif not summary.all_same_direction:
        headline = _template_mixed(summary)
    elif abs(summary.max_delta_pp) >= FLAT_DELTA_PP:
        # Direction is clear and there's a meaningful delta
        # Choose between driver-led and top-value patterns
        if summary.top_value_category == summary.max_delta_category:
            # Same category dominates both — use top-value
            headline = _template_rank_top(summary)
        else:
            headline = _template_delta_driver(summary)
    else:
        # Small changes across the board
        headline = _template_flat(summary)

    # Optional narrative hint prefix
    if narrative_hint and len(headline) + len(narrative_hint) + 3 <= max_chars:
        # Not always included — only when narrative adds context and fits
        pass  # Kept simple: template produces the headline; narrative consumed by LLM-augmented mode

    # Capitalize first letter
    if headline and not headline[0].isupper():
        headline = headline[0].upper() + headline[1:]

    # Hard truncate (shouldn't normally hit with reasonable templates)
    if len(headline) > max_chars:
        headline = headline[: max_chars - 1].rstrip(", ") + "…"

    return headline


def rewrite_period_labels(
    subheadline_text: str,
    old_period_prior: str,
    old_period_current: str,
    new_period_prior: str,
    new_period_current: str,
) -> str:
    """Rewrite period references in a subheadline.

    Example:
        rewrite_period_labels("Q4 '25 vs Q3 '25 among NSCLC prescribers",
                              "Q3 '25", "Q4 '25", "Q4 '25", "Q1 '26")
        → "Q1 '26 vs Q4 '25 among NSCLC prescribers"
    """
    if not subheadline_text:
        return subheadline_text
    result = subheadline_text
    # Replace current first (most specific) then prior, to avoid double-replacement
    if old_period_current and new_period_current:
        result = result.replace(old_period_current, new_period_current)
    if old_period_prior and new_period_prior:
        result = result.replace(old_period_prior, new_period_prior)
    return result
