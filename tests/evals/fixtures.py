"""Shared fixture paths for evals.

Decks live in projects/ (gitignored, OneDrive). Evals reference them via
FIXTURE_DECKS. If a deck is not found on disk, evals that need it skip
gracefully instead of failing.
"""
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]

FIXTURE_DECKS = {
    "atu_q1_26": REPO_ROOT / "projects" / "J&J Rybrevant PET" / "Template" / "ZoomRx_UC_ATU_Report_Q1_'26.pptx",
    "creon_pet_w33": REPO_ROOT / "projects" / "J&J Rybrevant PET" / "Template" / "CREON Share of Voice Study - W33.pptx",
}

# Refreshed decks — the Stage 2 output (post-refresh, from Vijay's test_spec_refresh_pipeline.py).
# Used by Eval #3 (refresh execution) to compare source vs refreshed.
# If a refreshed fixture is missing, the eval skips with guidance to generate it.
REFRESHED_DECKS = {
    "atu_q1_26": REPO_ROOT / "output" / "atu_refresh" / "ATU_refreshed.pptx",
    "creon_pet_w33": REPO_ROOT / "output" / "creon_pet_w33_refresh" / "creon_pet_w33_refreshed.pptx",
}

GOLDEN_DIR = Path(__file__).parent / "deck_reader" / "golden"
RESULTS_DIR = Path(__file__).parent / "results"


def require_deck(key: str) -> Path:
    """Return path to a fixture deck. Raises FileNotFoundError if missing."""
    path = FIXTURE_DECKS.get(key)
    if path is None:
        raise KeyError(f"Unknown fixture deck: {key}")
    if not path.exists():
        raise FileNotFoundError(
            f"Fixture deck not found: {path}\n"
            "If you don't have the projects/ folder symlinked, this eval will skip."
        )
    return path
