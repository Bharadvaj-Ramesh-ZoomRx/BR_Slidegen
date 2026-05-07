"""ConvertConfigAsNonDisruptive (Connector spec §17.4) — pair old
columnDefinitions to current pivot output via:
  Pass 1: exact match
  Pass 2: TP-template match (TP parts → placeholder, pair templates)
  Pass 3: substring containment fallback (alias rename — old name's
          parts all appear in a new name's parts)

These tests pin the contract used by `pivot_records_to_chart_data` to
rewrite `selectedColumns` against the live pivot output. Closing this
gap is what eliminates the `selectedColumns_drift` REVIEW notes on
Testing Deck slides 39 and 44 (drift on TP names like 'Deliverable'
or alias-renamed segment values).
"""
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO_ROOT))

from slidegen.synapse_chart_mapper import (
    _build_old_to_new_column_map,
    _column_template,
    _rewrite_with_old_to_new,
)


# ── Pass 1: exact match ────────────────────────────────────────────────

def test_exact_match_pairs_to_self():
    m = _build_old_to_new_column_map(
        column_fields=[],
        column_definitions=[{"Name": "A"}, {"Name": "B"}],
        pivot_columns=["A", "B", "C"],
    )
    assert m["A"] == "A"
    assert m["B"] == "B"


def test_no_match_for_genuinely_missing_old():
    m = _build_old_to_new_column_map(
        column_fields=[],
        column_definitions=[{"Name": "Vanished"}],
        pivot_columns=["A", "B"],
    )
    # Vanished doesn't match anything (no fallbacks viable)
    assert "Vanished" not in m


# ── Pass 2: TP-template match ──────────────────────────────────────────

def test_tp_template_rename_picks_latest():
    """Source had 'Q1 2026 @:@ Sum of base', API now returns Q2 + Q3.
    Template '<TP> @:@ Sum of base' matches both new cols; pick the
    LAST in pivot order (chronologically latest)."""
    m = _build_old_to_new_column_map(
        column_fields=["time_period_name"],
        column_definitions=[{"Name": "Q1 2026 @:@ Sum of base"}],
        pivot_columns=["Q2 2026 @:@ Sum of base", "Q3 2026 @:@ Sum of base"],
    )
    assert m["Q1 2026 @:@ Sum of base"] == "Q3 2026 @:@ Sum of base"


def test_tp_template_with_segment_part_preserved():
    """Compound: 'Future_overall @:@ Q1 2026 @:@ Sum of share' →
    pair to 'Future_overall @:@ Q3 2026 @:@ Sum of share' (same segment,
    new TP). The non-TP part 'Future_overall' must be preserved in the
    template match."""
    m = _build_old_to_new_column_map(
        column_fields=["alias_label", "time_period_name"],
        column_definitions=[{"Name": "Future_overall @:@ Q1 2026 @:@ Sum of share"}],
        pivot_columns=[
            "Future_overall @:@ Q3 2026 @:@ Sum of share",
            "Current_2L @:@ Q3 2026 @:@ Sum of share",
        ],
    )
    assert m["Future_overall @:@ Q1 2026 @:@ Sum of share"] == \
        "Future_overall @:@ Q3 2026 @:@ Sum of share"
    # Critically: Current_2L was NOT chosen even though it's also Q3
    assert m["Future_overall @:@ Q1 2026 @:@ Sum of share"] != \
        "Current_2L @:@ Q3 2026 @:@ Sum of share"


def test_tp_template_no_tp_fields_no_pairing():
    """When ColumnFields has no TP entry, Pass 2 doesn't fire; falls to
    Pass 3."""
    m = _build_old_to_new_column_map(
        column_fields=["region"],  # no TP field
        column_definitions=[{"Name": "Q1 2026 @:@ Sum of base"}],
        pivot_columns=["Q3 2026 @:@ Sum of base"],
    )
    # Pass 3 substring would only fire if old_parts ALL appear in new.
    # Here 'Q1 2026' isn't in 'Q3 2026 @:@ Sum of base'. No pairing.
    assert "Q1 2026 @:@ Sum of base" not in m


# ── Pass 3: substring containment (alias rename) ───────────────────────

def test_alias_rename_full_path_to_short_path():
    """Connector default-alias rule trims segment paths. Source had
    'Specialty - CARD - L' but API now emits 'CARD - L' (alias dropped
    the prefix). Pair via substring containment."""
    m = _build_old_to_new_column_map(
        column_fields=[],
        column_definitions=[{"Name": "Specialty - CARD - L"}],
        pivot_columns=["CARD - L", "ONC - L"],
    )
    assert m["Specialty - CARD - L"] == "CARD - L"


def test_substring_picks_best_overlap_when_multiple():
    """When multiple new cols partially overlap, pick highest overlap."""
    m = _build_old_to_new_column_map(
        column_fields=[],
        column_definitions=[{"Name": "Northeast - Cardiology - Tier 3"}],
        pivot_columns=["Cardiology - Tier 3", "Northeast - Cardiology"],
    )
    # 'Cardiology - Tier 3' contains 2/3 old parts (Cardiology, Tier 3).
    # 'Northeast - Cardiology' contains 2/3 old parts (Northeast, Cardiology).
    # Both have shorter=2; pick first match found.
    assert m["Northeast - Cardiology - Tier 3"] in {
        "Cardiology - Tier 3", "Northeast - Cardiology"
    }


# ── Full integration ───────────────────────────────────────────────────

def test_selectedColumns_get_rewritten_after_tp_rename():
    """The whole point: feed selectedColumns from spec, get back rewritten
    list whose entries match the current pivot output."""
    pivot_columns = ["Q3 2026 @:@ Sum of base", "Q4 2026 @:@ Sum of base"]
    m = _build_old_to_new_column_map(
        column_fields=["time_period_name"],
        column_definitions=[
            {"Name": "Q1 2026 @:@ Sum of base"},
            {"Name": "Q2 2026 @:@ Sum of base"},
        ],
        pivot_columns=pivot_columns,
        selected_columns=["region", "Q1 2026 @:@ Sum of base"],
    )
    rewritten = _rewrite_with_old_to_new(
        ["region", "Q1 2026 @:@ Sum of base"], m
    )
    # 'region' has no mapping → passed through unchanged
    # 'Q1 2026...' rewritten to its TP-paired latest TP
    assert "region" in rewritten
    assert "Q4 2026 @:@ Sum of base" in rewritten
    assert "Q1 2026 @:@ Sum of base" not in rewritten


def test_unmapped_entries_pass_through_unchanged():
    """Entries not in oldToNew stay verbatim — caller's responsibility
    to handle them downstream (filter will fail and surface drift)."""
    m = {"A": "A_new"}
    out = _rewrite_with_old_to_new(["A", "B", "C"], m)
    assert out == ["A_new", "B", "C"]


def test_none_mapped_entries_dropped():
    """Old-to-None means 'this entry is gone, drop it'."""
    m = {"A": None, "B": "B_new"}
    out = _rewrite_with_old_to_new(["A", "B", "C"], m)
    assert out == ["B_new", "C"]


def test_empty_pivot_columns_returns_empty_map():
    m = _build_old_to_new_column_map([], [{"Name": "A"}], [])
    assert m == {}


def test_column_template_helper_basic():
    assert _column_template("Q1 2026 @:@ Sum of base", [0]) == \
        "<TP> @:@ Sum of base"


def test_column_template_helper_two_tp_parts():
    assert _column_template("X @:@ Q1 @:@ Y", [1]) == "X @:@ <TP> @:@ Y"


def test_column_template_helper_no_tp_returns_none():
    assert _column_template("Q1 2026 @:@ Sum of base", []) is None
    assert _column_template("not_a_compound", [0]) == "<TP>"
    assert _column_template("Q1", [1]) is None  # tp_index out of range
