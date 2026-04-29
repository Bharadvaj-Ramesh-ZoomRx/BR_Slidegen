"""Step 3 — Visual formatting preserved across refresh.

Runs compare_formatting() on source vs refreshed decks and asserts:
  - chart_type_match: 100% (every chart type identical)
  - series_colors_match: ≥95% (per-series fill colors preserved within
    a small tolerance for legitimate Connector-driven recolor on
    refresh-time additions)

Step 3 of the 7-step refresh framework — independent of data-level
signals captured in compare_decks.py (Step 2). Catches regressions
that "values match" can't see, e.g. a refresh path that drops the
chart's color pattern or swaps the chart type.

Run:
    pytest tests/evals/end_to_end/test_step3_format_preservation.py -v
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT))

from tests.evals.end_to_end.compare_formatting import compare_formatting  # noqa: E402
from tests.evals.fixtures import FIXTURE_DECKS, REFRESHED_DECKS  # noqa: E402

CHART_TYPE_MIN = 1.00       # chart_type drift is a refusal — must be 100%
SERIES_COLORS_MIN = 0.95    # color drift can have legitimate edge cases


@pytest.mark.parametrize("deck_key", list(REFRESHED_DECKS.keys()))
def test_format_preserved(deck_key: str):
    src = FIXTURE_DECKS.get(deck_key)
    ref = REFRESHED_DECKS.get(deck_key)
    if src is None or not src.exists():
        pytest.skip(f"Source deck not on disk: {src}")
    if ref is None or not ref.exists():
        pytest.skip(f"Refreshed deck not on disk: {ref}")

    report = compare_formatting(src, ref)
    t = report.totals()
    n = t["components_compared"]
    if n == 0:
        pytest.skip(f"No comparable charts on {deck_key}")

    type_pct = t["chart_type_match"] / n
    color_pct = t["series_colors_match"] / n

    msgs = []
    if type_pct < CHART_TYPE_MIN:
        msgs.append(
            f"chart_type {t['chart_type_match']}/{n} = {type_pct:.1%} "
            f"(< floor {CHART_TYPE_MIN:.0%})"
        )
    if color_pct < SERIES_COLORS_MIN:
        msgs.append(
            f"series_colors {t['series_colors_match']}/{n} = {color_pct:.1%} "
            f"(< floor {SERIES_COLORS_MIN:.0%})"
        )
    if t["components_with_errors"] > 0:
        msgs.append(f"{t['components_with_errors']} components had extraction errors")

    assert not msgs, f"[{deck_key}] " + "; ".join(msgs)
