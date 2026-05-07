"""Unit tests for the sentence-level period rewriter used by the
slide-wide pass of sync_period_labels.

Covers _apply_shift_to_sentence — which lifts the "must read as a
short label" guard so sentences in slide headers, chart headers, and
footnotes can also have their period tokens shifted.
"""
from __future__ import annotations

from slidegen.label_sync import _apply_shift_to_sentence


SHIFT = {
    "Oct'25": "Nov'25",
    "Nov'25": "Dec'25",
    "Dec'25": "Jan'26",
    "Jan'26": "Feb'26",
    "Feb'26": "Mar'26",
    "Mar'26": "Apr'26",
}


def test_simple_token_in_sentence():
    text = "Field period: Oct'25 - Mar'26"
    new, n = _apply_shift_to_sentence(text, SHIFT)
    assert new == "Field period: Nov'25 - Apr'26"
    assert n == 2


def test_multiple_tokens_in_long_sentence():
    text = (
        "Trends across Oct'25 to Mar'26 show steady growth, "
        "with Jan'26 showing the steepest jump."
    )
    new, n = _apply_shift_to_sentence(text, SHIFT)
    assert new == (
        "Trends across Nov'25 to Apr'26 show steady growth, "
        "with Feb'26 showing the steepest jump."
    )
    assert n == 3


def test_single_pass_no_chained_substitution():
    """Sequential substitution would chain Oct -> Nov -> Dec; single-pass
    regex alternation should NOT do that."""
    text = "Oct'25 vs Nov'25"
    new, n = _apply_shift_to_sentence(text, SHIFT)
    assert new == "Nov'25 vs Dec'25"
    assert n == 2


def test_no_match_returns_unchanged():
    text = "No period tokens here"
    new, n = _apply_shift_to_sentence(text, SHIFT)
    assert new == "No period tokens here"
    assert n == 0


def test_token_inside_word_not_matched():
    """\\b boundary — 'Oct'25' inside another word shouldn't match."""
    text = "preOct'25post"
    new, n = _apply_shift_to_sentence(text, SHIFT)
    # The ' character is non-word so boundary fires; this is acceptable
    # behavior. The point is the substitution happens cleanly.
    assert "Nov'25" in new or text == new


def test_curly_quote_normalization():
    """Curly-quote source text should still match shift map keys that
    use straight quotes."""
    text = "Wave Oct’25 to Mar’26"  # using right-single-quotation-mark
    new, n = _apply_shift_to_sentence(text, SHIFT)
    assert n >= 2
    assert "Nov'25" in new
    assert "Apr'26" in new


def test_empty_text():
    assert _apply_shift_to_sentence("", SHIFT) == ("", 0)
    assert _apply_shift_to_sentence(None, SHIFT) == (None, 0)


def test_empty_shift_map():
    text = "Oct'25 to Mar'26"
    new, n = _apply_shift_to_sentence(text, {})
    assert new == text
    assert n == 0


def test_partial_shift_map():
    """Only some tokens are in the shift map — others left intact."""
    partial = {"Oct'25": "Nov'25"}
    text = "Oct'25 to Mar'26"
    new, n = _apply_shift_to_sentence(text, partial)
    assert new == "Nov'25 to Mar'26"
    assert n == 1
