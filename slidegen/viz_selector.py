"""
viz_selector.py — Deterministic chart-pattern selector.

Implements PRD §6.1 hierarchy:
  1. Metric tag → viz  (highest priority)
  2. Q-type default    (fallback)
  3. UNRESOLVED        (forces HITL)

Pure function — no LLM calls, no I/O, no randomness.
"""
from __future__ import annotations

from slidegen.slide_spec.schema import SUPPORTED_CHART_PATTERNS


# ─────────────────────────────────────────────────────────────────────────────
# METRIC TAG → CHART PATTERN
#
# Built from deck-analysis ground truth (32 PET decks, 4,354 charts).
# Each entry: metric_tag → (chart_pattern, observed_pct, rationale)
#
# The metric_tag is a normalized slug (lowercase, underscores). The caller
# should normalize its input before lookup.
# ─────────────────────────────────────────────────────────────────────────────

METRIC_TAG_MAP: dict[str, tuple[str, float, str]] = {
    # === Message Recall / Awareness / Recognition ===
    # bar_clustered_horizontal is the dominant PET pattern (35% of all charts).
    # 90%+ of recall/recognition slides use it with companion label table.
    "message_recall":         ("bar_clustered_horizontal", 35.0, "Dominant PET pattern for recall metrics — label table + horizontal bars + delta column"),
    "aided_awareness":        ("bar_clustered_horizontal", 35.0, "Same pattern as recall — ranked list of items with QoQ comparison"),
    "unaided_awareness":      ("bar_clustered_horizontal", 35.0, "Same pattern as recall"),
    "brand_awareness":        ("bar_clustered_horizontal", 35.0, "Brand recall follows same pattern"),
    "message_recognition":    ("bar_clustered_horizontal", 35.0, "Recognition slides identical to recall pattern"),
    "recall":                 ("bar_clustered_horizontal", 35.0, "Generic recall metric"),
    "recognition":            ("bar_clustered_horizontal", 35.0, "Generic recognition metric"),
    "top_of_mind":            ("bar_clustered_horizontal", 35.0, "Top-of-mind awareness — ranked bar chart"),
    "topic_recall":           ("bar_clustered_horizontal", 35.0, "Topics discussed — ranked horizontal bars"),

    # === Message Effectiveness / M-B-D ===
    # Abacus (xy_scatter) is used for multi-attribute comparison (17%).
    # ME slides typically show motivation, believability, differentiation as
    # separate dot positions on the same axis — abacus is the canonical format.
    "message_effectiveness":  ("xy_scatter_abacus", 17.0, "ME uses abacus for multi-attribute dot comparison (M/B/D)"),
    "motivation":             ("xy_scatter_abacus", 17.0, "Sub-dimension of ME — dot plot"),
    "believability":          ("xy_scatter_abacus", 17.0, "Sub-dimension of ME — dot plot"),
    "differentiation":        ("xy_scatter_abacus", 17.0, "Sub-dimension of ME — dot plot"),
    "mbd":                    ("xy_scatter_abacus", 17.0, "Combined M/B/D abacus"),

    # === Rep / Call Quality / Attributes ===
    # Abacus for multi-attribute comparison.
    "rep_performance":        ("xy_scatter_abacus", 17.0, "Multi-attribute rep comparison — abacus"),
    "call_quality":           ("xy_scatter_abacus", 17.0, "Quality attributes — abacus dot plot"),
    "rep_attributes":         ("xy_scatter_abacus", 17.0, "Rep attribute ratings — abacus"),
    "attribute_rating":       ("xy_scatter_abacus", 17.0, "Generic attribute comparison — abacus"),

    # === Trended / Time-Series ===
    # line_markers_trended covers 14% of charts. Used for any multi-wave trend.
    "awareness_trend":        ("line_markers_trended", 14.0, "Multi-wave awareness trend — line with markers"),
    "recall_trend":           ("line_markers_trended", 14.0, "Multi-wave recall trend"),
    "reach":                  ("line_markers_trended", 14.0, "Reach over waves — trended line"),
    "share_of_voice":         ("line_markers_trended", 14.0, "SOV over waves — trended line"),
    "frequency":              ("line_markers_trended", 14.0, "Call frequency trend"),
    "trended_scorecard":      ("line_markers_trended", 14.0, "Multi-panel mini line grid"),
    "activity_trend":         ("line_markers_trended", 14.0, "Activity metrics over waves"),

    # === Intent / Prescription / Composition ===
    # column_stacked_100_vertical (11%) for share-of-mind / intent distribution.
    "likelihood_to_prescribe": ("column_stacked_100_vertical", 11.0, "LTIP distribution — 100% stacked column"),
    "prescription_intent":     ("column_stacked_100_vertical", 11.0, "Intent scale distribution"),
    "intent_distribution":     ("column_stacked_100_vertical", 11.0, "Intent category stacking"),
    "patient_allocation":      ("column_stacked_100_vertical", 11.0, "Patient share allocation"),
    "share_of_mind":           ("column_stacked_100_vertical", 11.0, "Mind share as stacked composition"),

    # === Stacked Composition (horizontal) ===
    # bar_stacked_100_horizontal (7%) for horizontal composition views.
    "call_to_action":          ("bar_stacked_100_horizontal", 7.0, "CTA breakdown — horizontal stacked"),
    "branded_close":           ("bar_stacked_100_horizontal", 7.0, "Closing technique composition"),
    "interaction_format":      ("bar_stacked_100_horizontal", 7.0, "In-person / virtual / hybrid breakdown"),

    # === Segment Comparison / Clustered Vertical ===
    # column_clustered_vertical (7%) for side-by-side segment or brand comparison.
    "segment_comparison":      ("column_clustered_vertical", 7.0, "Segment-level comparison — clustered columns"),
    "hii_scorecard":           ("column_clustered_vertical", 7.0, "HII metric scorecard — clustered columns"),
    "brand_comparison":        ("column_clustered_vertical", 7.0, "Brand-to-brand comparison — clustered columns"),
}


# ─────────────────────────────────────────────────────────────────────────────
# Q-TYPE → CHART PATTERN (fallback)
# ─────────────────────────────────────────────────────────────────────────────

Q_TYPE_MAP: dict[str, str] = {
    "likert":                 "bar_clustered_horizontal",
    "ranking":                "xy_scatter_abacus",
    "time_series":            "line_markers_trended",
    "categorical_composition": "column_stacked_100_vertical",
    "ordinal":                "bar_clustered_horizontal",
    "multi_select":           "bar_clustered_horizontal",
    "single_select":          "bar_clustered_horizontal",
    "numeric_scale":          "xy_scatter_abacus",
    "stacked_composition":    "bar_stacked_100_horizontal",
}


# ─────────────────────────────────────────────────────────────────────────────
# Normalization helper
# ─────────────────────────────────────────────────────────────────────────────

def _normalize(tag: str) -> str:
    """Normalize a metric tag or q-type to a lookup key."""
    return tag.strip().lower().replace(" ", "_").replace("-", "_")


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────

def select_chart_pattern(
    metric: str | None = None,
    q_type: str | None = None,
    context: dict | None = None,
) -> str:
    """Select a chart pattern using the PRD §6.1 hierarchy.

    1. If ``metric`` matches a key in METRIC_TAG_MAP, return that pattern.
    2. Else if ``q_type`` matches Q_TYPE_MAP, return that pattern.
    3. Else return ``"UNRESOLVED"`` — spec-validator will reject this,
       forcing the planner to ask the user.

    Args:
        metric:  Metric tag (e.g. "Message Recall", "awareness_trend").
        q_type:  Question type (e.g. "likert", "time_series").
        context: Reserved for future use (e.g. deck-level overrides).

    Returns:
        A key from SUPPORTED_CHART_PATTERNS, or ``"UNRESOLVED"``.
    """
    # Priority 1: metric tag
    if metric:
        key = _normalize(metric)
        if key in METRIC_TAG_MAP:
            pattern = METRIC_TAG_MAP[key][0]
            assert pattern in SUPPORTED_CHART_PATTERNS, (
                f"METRIC_TAG_MAP[{key!r}] → {pattern!r} not in SUPPORTED_CHART_PATTERNS"
            )
            return pattern

    # Priority 2: question type
    if q_type:
        key = _normalize(q_type)
        if key in Q_TYPE_MAP:
            pattern = Q_TYPE_MAP[key]
            assert pattern in SUPPORTED_CHART_PATTERNS, (
                f"Q_TYPE_MAP[{key!r}] → {pattern!r} not in SUPPORTED_CHART_PATTERNS"
            )
            return pattern

    # Priority 3: HITL sentinel
    return "UNRESOLVED"


def get_metric_tags() -> list[str]:
    """Return all recognized metric tags (sorted)."""
    return sorted(METRIC_TAG_MAP.keys())


def get_q_types() -> list[str]:
    """Return all recognized question types (sorted)."""
    return sorted(Q_TYPE_MAP.keys())


def get_mapping_table() -> list[dict]:
    """Return the full metric-to-pattern mapping as a list of dicts.

    Each dict: {metric, chart_pattern, observed_pct, rationale}
    """
    return [
        {
            "metric": tag,
            "chart_pattern": pattern,
            "observed_pct": pct,
            "rationale": rationale,
        }
        for tag, (pattern, pct, rationale) in sorted(METRIC_TAG_MAP.items())
    ]
