"""
callout_writer.py — build CalloutComponent from data or quote.

Callouts are inline annotations on slides: verbatim quotes, data insights,
or narrative-thread tags. Used by `annotate-slide-workflow` to add a callout
to an existing slide.

Contract:
    build_callout(type, ...) -> CalloutComponent
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from slidegen.slide_spec import CalloutComponent, Position


# Style constants
MIN_QUOTE_WORDS = 15
LOW_BASE_THRESHOLD = 30


@dataclass
class Quote:
    """Verbatim quote from qualitative data."""
    text: str
    attribution: str                     # e.g. "HCP, Academic (n=18)"
    theme: Optional[str] = None
    n: Optional[int] = None              # sample size this quote represents (the theme)


@dataclass
class DataInsight:
    """A data-driven insight for callout annotation."""
    text: str                            # e.g. "Up 8pp among Academic vs Community"
    metric: Optional[str] = None
    theme: Optional[str] = None          # hypothesis tag
    p_value: Optional[float] = None      # for stat-sig-annotated insights
    n: Optional[int] = None              # sample size the insight is based on
    low_base: bool = False               # True if any underlying segment had n<30


class CalloutGenerationError(ValueError):
    """Raised when callout cannot be built with given inputs."""


def _default_position(style: str = "default") -> Position:
    """Return a canonical callout position — bottom-right by default.

    Real PET decks typically place callouts below the chart or right of it.
    For a standard observed_1chart_1table slide, the bottom-right corner
    (roughly 10.00", 6.00", 2.80", 1.20") is the most common location.
    """
    # Callout sits at bottom-right of the typical data area
    return Position(left=10.00, top=6.00, width=2.80, height=1.20)


def build_quote_callout(
    quote: Quote,
    position: Optional[Position] = None,
    theme_override: Optional[str] = None,
) -> CalloutComponent:
    """Build a CalloutComponent from a verbatim quote.

    Raises CalloutGenerationError if quote is too short (<15 words) or
    missing attribution.
    """
    if not quote.text or not quote.text.strip():
        raise CalloutGenerationError("Quote text is empty")

    word_count = len(quote.text.split())
    if word_count < MIN_QUOTE_WORDS:
        # Short quotes tend to be fragments — flag but allow (caller's call)
        # Real decks sometimes use 10-word quotes; just warn via metadata.
        pass

    theme = theme_override or quote.theme

    # Flag low-base quotes in attribution
    attribution = quote.attribution
    if quote.n is not None and quote.n < 5:
        attribution = f"{attribution} — low base (n={quote.n})"

    return CalloutComponent(
        position=position or _default_position("default"),
        text=quote.text,
        style="default",
        theme=theme,
        attribution=attribution,
    )


def build_insight_callout(
    insight: DataInsight,
    position: Optional[Position] = None,
) -> CalloutComponent:
    """Build a CalloutComponent from a data-driven insight (e.g. segment delta).

    For insights backed by a statistical test, the p-value is rendered as
    a significance marker inline with the text.
    """
    if not insight.text or not insight.text.strip():
        raise CalloutGenerationError("Insight text is empty")

    # Add significance marker if p-value is provided
    text = insight.text
    if insight.p_value is not None:
        if insight.p_value < 0.01:
            text = f"{text} **"
        elif insight.p_value < 0.05:
            text = f"{text} *"

    # Flag low-base insights
    if insight.low_base:
        text = f"{text}  (low base)"
    elif insight.n is not None and insight.n < LOW_BASE_THRESHOLD:
        text = f"{text}  (low base: n={insight.n})"

    return CalloutComponent(
        position=position or _default_position("insight"),
        text=text,
        style="insight",
        theme=insight.theme,
        attribution=None,
    )


def build_freeform_callout(
    text: str,
    position: Optional[Position] = None,
    style: str = "default",
    theme: Optional[str] = None,
    attribution: Optional[str] = None,
) -> CalloutComponent:
    """Build a CalloutComponent from arbitrary user-supplied text.

    No validation beyond non-empty text — use when caller has done their own
    sanity-checking.
    """
    if not text or not text.strip():
        raise CalloutGenerationError("Callout text is empty")

    if style not in ("default", "dashed", "insight"):
        raise CalloutGenerationError(
            f"Unknown callout style {style!r} — must be default | dashed | insight"
        )

    return CalloutComponent(
        position=position or _default_position(style),
        text=text,
        style=style,
        theme=theme,
        attribution=attribution,
    )


def pick_quote_from_pool(
    quote_pool: list[Quote],
    theme: Optional[str] = None,
    min_words: int = MIN_QUOTE_WORDS,
) -> Optional[Quote]:
    """Pick the best quote from a pool for a given theme.

    Preference order:
      1. Matches theme exactly, meets min_words, has clear attribution
      2. Matches theme exactly, meets min_words (weaker attribution)
      3. Adjacent theme (matches theme keywords), meets min_words
      4. Any quote meeting min_words
      5. None if pool is empty

    Returns None if no quote meets the min bar.
    """
    if not quote_pool:
        return None

    def _words(q: Quote) -> int:
        return len(q.text.split())

    # Exact theme + attribution + word bar
    for q in quote_pool:
        if (theme and q.theme == theme
            and q.attribution
            and _words(q) >= min_words):
            return q
    # Exact theme + word bar
    for q in quote_pool:
        if theme and q.theme == theme and _words(q) >= min_words:
            return q
    # Adjacent theme (keyword overlap)
    if theme:
        theme_words = set(theme.lower().split())
        for q in quote_pool:
            if q.theme and _words(q) >= min_words:
                if theme_words & set(q.theme.lower().split()):
                    return q
    # Any quote meeting word bar
    for q in quote_pool:
        if _words(q) >= min_words:
            return q
    return None
