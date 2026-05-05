"""Deckwide contract — combined refresh + headline check per connected slide.

For every connected slide in each fixture deck, asserts that the
(refresh_status, headline_status) pair is one of the documented valid
combinations. This wires the headline rewrite contract into the broader
eval pass so a single test gives one combined pass/fail per deck instead
of having to read two JSONs separately.

Valid pairs (per `scripts/annotate_refresh_and_headline.py` and Step 7
contract eval):

    refresh                          | headline    | meaning
    ---------------------------------|-------------|----------------------------
    ok (values changed)              | updated     | normal forward refresh
    ok (values unchanged)            | unchanged   | data unchanged; preserve
    ok (values changed, no narrative)| no_headline | refresh ok but slide had no
                                     |             |   headline shape to rewrite
    ok + tag_mismatch note           | unchanged   | tag mismatch: don't rewrite
                                     |             |   from a divergent number
    ok + partial_alignment note      | updated /   | values may have changed for
                                     |  unchanged  |   the kept items only
    ok + dynamic_added note          | updated     | new items flowed in,
                                     |             |   headline reflects shift
    static_pinned_skipped            | unchanged   | frozen by tag
    no_data_for_shifted_window       | unchanged   | no fresh data
    alignment_failed                 | unchanged   | mapper preserved source
    empty / error                    | unchanged   | refresh failed; preserve
    (no chart on slide)              | no_chart    | structural — out of scope

Invalid (assertion failure):
    refresh failed     + headline updated
    ok + tag_mismatch  + headline updated
    ok values unchanged + headline updated  (already caught by Step 7 contract)

Run:
    pytest tests/evals/end_to_end/test_deckwide_contract.py -v
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[3]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from tests.evals.fixtures import REFRESHED_DECKS  # noqa: E402

OUTPUT_DIR = REPO_ROOT / "output" / "step2_test_connected"


def _refresh_status_path(deck_key: str) -> Path:
    return OUTPUT_DIR / f"{deck_key}_refresh_status.json"


def _headline_status_path(deck_key: str) -> Path:
    return OUTPUT_DIR / f"{deck_key}_with_headlines_status.json"


VALID_REFRESH_HEADLINE = {
    # (refresh_outcome, headline_status) pairs that are contractually allowed
    ("ok", "updated"),
    ("ok", "unchanged"),
    ("ok", "no_headline"),
    ("static_pinned_skipped", "unchanged"),
    ("no_data_for_shifted_window", "unchanged"),
    ("alignment_failed", "unchanged"),
    ("empty", "unchanged"),
    ("error", "unchanged"),
    # Headline tests sometimes report "no_chart" for slides without charts —
    # always accept that.
    ("ok", "no_chart"),
    ("static_pinned_skipped", "no_chart"),
    ("no_data_for_shifted_window", "no_chart"),
    ("alignment_failed", "no_chart"),
    ("empty", "no_chart"),
    # Step 7 may also emit extract_error for chart-extraction failures
    ("ok", "extract_error"),
}


def _classify_slide_refresh(slide_status: dict) -> str:
    """Reduce per-shape statuses on one slide to a single refresh-outcome label.

    Priority: alignment_failed > tag_mismatch (note) > error > empty >
    static_pinned_skipped (all) > no_data_for_shifted_window (all) > ok.
    """
    components = (slide_status.get("charts") or []) + (slide_status.get("tables") or [])
    if not components:
        return "ok"  # nothing to refresh on this slide
    statuses = [c.get("status", "") for c in components]
    notes = [n.get("kind") for c in components for n in (c.get("notes") or [])]

    if "alignment_failed" in statuses:
        return "alignment_failed"
    if "error" in statuses:
        return "error"
    if "empty" in statuses:
        return "empty"
    if all(s == "static_pinned_skipped" for s in statuses):
        return "static_pinned_skipped"
    if all(s == "no_data_for_shifted_window" for s in statuses):
        return "no_data_for_shifted_window"
    return "ok"


@pytest.mark.parametrize("deck_key", list(REFRESHED_DECKS.keys()))
def test_refresh_headline_contract(deck_key: str):
    """Every connected slide's (refresh, headline) pair is in the valid set."""
    rs_path = _refresh_status_path(deck_key)
    hl_path = _headline_status_path(deck_key)
    if not rs_path.exists():
        pytest.skip(f"refresh_status not on disk: {rs_path}")
    if not hl_path.exists():
        pytest.skip(f"headline status not on disk: {hl_path}")

    refresh_status = json.loads(rs_path.read_text(encoding="utf-8"))
    headline_status = json.loads(hl_path.read_text(encoding="utf-8"))

    refresh_by_idx = {s["slide_index"]: s for s in refresh_status.get("slides", [])}
    headline_by_idx = {u["slide_idx"]: u for u in headline_status.get("updates", [])}

    invalid: list[str] = []
    note_violations: list[str] = []

    common_slides = set(refresh_by_idx) & set(headline_by_idx)
    if not common_slides:
        pytest.skip(f"No overlapping slides between refresh + headline status for {deck_key}")

    for si in sorted(common_slides):
        rs_slide = refresh_by_idx[si]
        hl_slide = headline_by_idx[si]
        refresh_outcome = _classify_slide_refresh(rs_slide)
        headline_outcome = hl_slide.get("status", "")

        # Tag-mismatch + headline=updated is a contract violation:
        # if the chart's values diverge from source structure, the existing
        # headline likely no longer makes sense and shouldn't be rewritten
        # off potentially wrong numbers.
        had_tag_mismatch = any(
            n.get("kind") == "tag_mismatch"
            for c in (rs_slide.get("charts") or []) + (rs_slide.get("tables") or [])
            for n in (c.get("notes") or [])
        )
        if had_tag_mismatch and headline_outcome == "updated":
            note_violations.append(
                f"slide {si}: tag_mismatch note + headline=updated "
                "(rewriting on divergent numbers — review tag first)"
            )

        if (refresh_outcome, headline_outcome) not in VALID_REFRESH_HEADLINE:
            invalid.append(
                f"slide {si}: refresh={refresh_outcome!r} "
                f"+ headline={headline_outcome!r} not in valid contract set"
            )

    if invalid or note_violations:
        msg = [f"\n[{deck_key}] deckwide contract violations:"]
        for v in invalid[:20]:
            msg.append(f"  - {v}")
        if len(invalid) > 20:
            msg.append(f"  ... and {len(invalid) - 20} more invalid pairs")
        for v in note_violations[:20]:
            msg.append(f"  - {v}")
        if len(note_violations) > 20:
            msg.append(f"  ... and {len(note_violations) - 20} more note violations")
        pytest.fail("\n".join(msg))


@pytest.mark.parametrize("deck_key", list(REFRESHED_DECKS.keys()))
def test_deckwide_summary_counts(deck_key: str):
    """Print a per-deck counts summary so eval runs surface the headline picture
    alongside the refresh picture in one place. Always passes; informational."""
    rs_path = _refresh_status_path(deck_key)
    hl_path = _headline_status_path(deck_key)
    if not rs_path.exists() or not hl_path.exists():
        pytest.skip(f"Status JSONs missing for {deck_key}")

    rs = json.loads(rs_path.read_text(encoding="utf-8"))
    hl = json.loads(hl_path.read_text(encoding="utf-8"))

    refresh_kinds: dict[str, int] = {}
    note_kinds: dict[str, int] = {}
    for slide in rs.get("slides", []):
        refresh_kinds[_classify_slide_refresh(slide)] = (
            refresh_kinds.get(_classify_slide_refresh(slide), 0) + 1
        )
        for c in (slide.get("charts") or []) + (slide.get("tables") or []):
            for n in c.get("notes") or []:
                k = n.get("kind", "?")
                note_kinds[k] = note_kinds.get(k, 0) + 1

    headline_kinds: dict[str, int] = {}
    for u in hl.get("updates", []):
        s = u.get("status", "?")
        headline_kinds[s] = headline_kinds.get(s, 0) + 1

    print(f"\n[{deck_key}] refresh outcomes: {dict(sorted(refresh_kinds.items()))}")
    print(f"[{deck_key}] refresh notes:    {dict(sorted(note_kinds.items()))}")
    print(f"[{deck_key}] headline outcomes:{dict(sorted(headline_kinds.items()))}")


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-v", "-s"]))
