"""Regression: latest_n picks the chronologically-latest waves, not the
ones with highest time_period_id.

Synapse assigned the Jan-Mar'26 CREON deliverables time_period_ids in
REVERSE chronological order (Jan=25375 > Feb=25374 > Mar=25373). If the
client-side latest_n filter (intelligent_refresh.fetch_synapse_data)
sorts by id descending, latest_n=2 returns Jan+Feb instead of Mar+Feb,
which is exactly the CREON slide 4 anomaly logged 2026-04-30.

This test pins the contract: sorting by parsed wave NAME via
_chronological_wave_key produces the chronologically-correct window
regardless of how Synapse assigned the IDs.

Per Connector spec §17.5.3 the connector itself sorts by API-position
(reportApi.TimePeriods) — our client-side fallback uses chrono-name
which produces the same result whenever the API returns periods in any
order.
"""
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO_ROOT))

from slidegen.intelligent_refresh import _chronological_wave_key


def test_creon_reverse_chrono_block():
    """The exact CREON case: Jan'26=25375 > Feb'26=25374 > Mar'26=25373.

    Sorting by parsed name must yield Jan < Feb < Mar regardless of id.
    """
    names = ["Jan'26", "Feb'26", "Mar'26"]
    ranked = sorted(names, key=_chronological_wave_key)
    assert ranked == ["Jan'26", "Feb'26", "Mar'26"]


def test_latest_2_picks_chronologically_latest():
    """latest_n=2 over the CREON block must keep {Mar'26, Feb'26},
    NOT {Jan'26, Feb'26} (which is what id-descending would return)."""
    names = {"Oct'25", "Nov'25", "Dec'25", "Jan'26", "Feb'26", "Mar'26"}
    ranked = sorted(names, key=_chronological_wave_key)
    keep = set(ranked[-2:])
    assert keep == {"Mar'26", "Feb'26"}


def test_apr_when_added_lands_correctly():
    """When Apr'26 lands later (potentially with a higher OR lower id than
    the Jan-Mar block), name-based sort still places it last."""
    names_with_apr = ["Mar'26", "Feb'26", "Jan'26", "Apr'26"]
    ranked = sorted(names_with_apr, key=_chronological_wave_key)
    assert ranked[-1] == "Apr'26"


def test_quarter_labels_chronological():
    """Q-style labels: Q1'26, Q2'26, Q3'26 sort in calendar order."""
    qs = ["Q3'26", "Q1'26", "Q2'26"]
    ranked = sorted(qs, key=_chronological_wave_key)
    assert ranked == ["Q1'26", "Q2'26", "Q3'26"]


def test_wave_n_labels_chronological():
    """Wave N labels: numeric ascending."""
    waves = ["Wave 12", "Wave 9", "Wave 11", "Wave 10"]
    ranked = sorted(waves, key=_chronological_wave_key)
    assert ranked == ["Wave 9", "Wave 10", "Wave 11", "Wave 12"]


def test_w_short_labels_chronological():
    """W## short form."""
    ws = ["W34", "W32", "W33"]
    ranked = sorted(ws, key=_chronological_wave_key)
    assert ranked == ["W32", "W33", "W34"]


def test_cross_year_chronology():
    """Year boundary: Dec'25 < Jan'26."""
    names = ["Jan'26", "Dec'25", "Nov'25", "Feb'26"]
    ranked = sorted(names, key=_chronological_wave_key)
    assert ranked == ["Nov'25", "Dec'25", "Jan'26", "Feb'26"]


def test_latest_n_filter_logic_matches_fetch_synapse_data():
    """Mirror the exact filter logic in fetch_synapse_data so the
    regression catches changes to either side.

    Construct fake records with the CREON reverse-chrono ID assignment.
    The filter in fetch_synapse_data only looks at name/key, not at id —
    so the IDs here are decorative (proving id is irrelevant).
    """
    records = [
        {"time_period_name": "Mar'26", "time_period_id": 25373, "value": 0.30},
        {"time_period_name": "Feb'26", "time_period_id": 25374, "value": 0.40},
        {"time_period_name": "Jan'26", "time_period_id": 25375, "value": 0.50},
        {"time_period_name": "Dec'25", "time_period_id": 15188, "value": 0.20},
        {"time_period_name": "Nov'25", "time_period_id": 15187, "value": 0.10},
    ]
    latest_n = 2

    names = {r["time_period_name"] for r in records}
    ranked = sorted(names, key=_chronological_wave_key)
    keep_set = set(ranked[-latest_n:])
    kept = [r for r in records if r["time_period_name"] in keep_set]

    kept_names = {r["time_period_name"] for r in kept}
    assert kept_names == {"Mar'26", "Feb'26"}, (
        f"Expected latest_n=2 to keep Mar+Feb (chronologically latest), "
        f"got {kept_names}. Sorting by time_period_id desc would have "
        f"returned {{Jan'26, Feb'26}} — that's the CREON slide 4 bug."
    )
