"""Wave-pinned selectedColumns must preserve single-wave intent.

Slides 49/50 of the Testing Deck had selectedColumns like
['value', 'Project Wave 13'] (standalone) or
['title', 'Future_overall - Project Wave 13'] (compound).

The data-side normalization at pivot_records_to_chart_data line 706 strips
"Project " from time_period_name values, leaving the data column as just
"Wave 13". The wave-pinned rewrite (Fix L+M) was then expanding the single
wave entry to ALL fetched waves (Wave 12 AND Wave 13), so:

  - slide 49 charts pinned to ONE wave got both waves on the X-axis (extra data)
  - slide 50 chart with compound 'Future_overall - Project Wave 13' got
    re-pivoted to 4 wave-segment combos × 15 brands instead of 1 cat × 17 series

Fix: when the entry's wave (after normalizing "Project " prefix) IS in the
fetched waves, emit a single rewritten entry. Only fall back to
expand-to-all-fetched when the wave has rolled off the API window.
"""
import sys
from pathlib import Path

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO_ROOT))

from slidegen.synapse_chart_mapper import _rewrite_wave_pinned_selected_columns


def _df(waves):
    """Build a minimal df with time_period_name + time_period_id."""
    return pd.DataFrame({
        "time_period_name": waves,
        "time_period_id": list(range(100, 100 + len(waves))),
    })


def test_standalone_project_wave_kept_when_in_fetched():
    """Slide 49 case: ['value', 'Project Wave 13'] with fetched=['Wave 12','Wave 13'].

    Normalize 'Project Wave 13' → 'Wave 13' (in fetched). KEEP — don't
    expand to all fetched waves."""
    df = _df(["Wave 12", "Wave 13"])
    selected = ["value", "Project Wave 13"]
    rewritten, n = _rewrite_wave_pinned_selected_columns(selected, df)
    assert "Wave 13" in rewritten
    assert "Wave 12" not in rewritten, (
        f"Single-wave pin must not expand to all fetched waves. "
        f"Got {rewritten}"
    )
    assert "value" in rewritten


def test_compound_project_wave_kept_when_in_fetched():
    """Slide 50 case: ['title', 'Future_overall - Project Wave 13'] with
    fetched=['Wave 12','Wave 13']. Normalize compound → keep single entry."""
    df = _df(["Wave 12", "Wave 13"])
    selected = ["title", "Future_overall - Project Wave 13"]
    rewritten, n = _rewrite_wave_pinned_selected_columns(selected, df)
    assert "Future_overall - Wave 13" in rewritten
    assert "Future_overall - Wave 12" not in rewritten, (
        f"Compound single-wave pin must not expand. Got {rewritten}"
    )
    assert "title" in rewritten


def test_already_normalized_wave_kept_unchanged():
    """When selectedColumns already uses 'Wave 13' form (no Project
    prefix), the entry is also a single-pin and stays single."""
    df = _df(["Wave 12", "Wave 13"])
    selected = ["value", "Wave 13"]
    rewritten, _ = _rewrite_wave_pinned_selected_columns(selected, df)
    # Just one wave entry in result
    wave_entries = [e for e in rewritten if "Wave" in e]
    assert wave_entries == ["Wave 13"]


def test_pinned_wave_rolled_off_falls_back_to_expand():
    """Source pinned 'Wave 11' but API now only returns ['Wave 12','Wave 13'].
    Wave 11 isn't in fetched — fall back to expand-to-all so the chart
    advances to whatever the API offers."""
    df = _df(["Wave 12", "Wave 13"])
    selected = ["value", "Wave 11"]
    rewritten, n = _rewrite_wave_pinned_selected_columns(selected, df)
    # Wave 11 dropped; both fetched waves added
    assert "Wave 11" not in rewritten
    assert "Wave 12" in rewritten
    assert "Wave 13" in rewritten


def test_multiple_wave_entries_all_in_fetched_each_kept():
    """Source listed multiple waves, all still in fetched. Each kept
    individually (still single-pin per entry)."""
    df = _df(["Wave 12", "Wave 13"])
    selected = ["L", "Wave 12", "Wave 13"]
    rewritten, _ = _rewrite_wave_pinned_selected_columns(selected, df)
    # Both waves preserved (each entry was a single-pin in fetched)
    assert "L" in rewritten
    assert "Wave 12" in rewritten
    assert "Wave 13" in rewritten


def test_no_time_period_in_df_noop():
    """No-op when the df has no time_period_name column."""
    df = pd.DataFrame({"y_label": ["a", "b"]})
    selected = ["y_label", "Wave 13"]
    rewritten, n = _rewrite_wave_pinned_selected_columns(selected, df)
    assert rewritten == selected
    assert n == 0


def test_no_wave_entries_noop():
    """When no entries are wave-shaped, function returns input unchanged."""
    df = _df(["Wave 12"])
    selected = ["region", "Sum of base"]
    rewritten, n = _rewrite_wave_pinned_selected_columns(selected, df)
    assert rewritten == selected
    assert n == 0
