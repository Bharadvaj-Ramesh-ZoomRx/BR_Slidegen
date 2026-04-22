"""Generate a refreshed-deck fixture for Eval #3.

Runs Vijay's test_spec_refresh_pipeline.py Stage 1 + Stage 2 against a
user-specified source deck by monkey-patching the pipeline's module-level
constants. Does NOT modify Vijay's file. Writes output into a per-deck
folder under output/ so fixtures don't collide.

Usage:
    python scripts/gen_refreshed_fixture.py <deck_key>

Where <deck_key> is a key in tests/evals/fixtures.py::FIXTURE_DECKS.
"""
from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))


def main():
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(2)

    deck_key = sys.argv[1]

    from tests.evals.fixtures import FIXTURE_DECKS
    source_pptx = FIXTURE_DECKS.get(deck_key)
    if source_pptx is None:
        print(f"ERROR: unknown deck key '{deck_key}'. Known keys: {list(FIXTURE_DECKS.keys())}")
        sys.exit(2)
    if not source_pptx.exists():
        print(f"ERROR: source deck not on disk: {source_pptx}")
        sys.exit(2)

    work_dir = REPO_ROOT / "output" / f"{deck_key}_refresh"
    work_dir.mkdir(parents=True, exist_ok=True)

    dummy_pptx = work_dir / f"{deck_key}_dummy.pptx"
    refreshed_pptx = work_dir / f"{deck_key}_refreshed.pptx"
    specs_json = work_dir / f"{deck_key}_specs.json"

    # Generate specs up front (Stage 2 can read from disk faster than re-running extraction)
    print(f"Generating specs for {deck_key}...")
    from slidegen.deck_reader.tag_reader import generate_config_specs
    from slidegen.slide_spec.schema import dump_spec

    specs, summary = generate_config_specs(str(source_pptx))
    specs_json.write_text(
        json.dumps([json.loads(dump_spec(s)) for s in specs], indent=2),
        encoding="utf-8",
    )
    print(f"  wrote {len(specs)} specs -> {specs_json.relative_to(REPO_ROOT)}")
    print(f"  tagged shapes: {summary.tagged_shapes}, configs resolved: {summary.report_configs_resolved}")

    # Monkey-patch Vijay's pipeline constants
    import tests.test_spec_refresh_pipeline as pipeline
    pipeline.SOURCE_PPTX = source_pptx
    pipeline.DUMMY_PPTX = dummy_pptx
    pipeline.REFRESHED_PPTX = refreshed_pptx
    pipeline.SPECS_JSON = specs_json

    print(f"\nRunning Stage 1 — create dummy deck")
    pipeline.stage1_create_dummy_deck()

    print(f"\nRunning Stage 2 — refresh from Synapse")
    ok = pipeline.stage2_refresh_from_specs()
    if not ok:
        print("\nStage 2 failed — check token / network. Output may be incomplete.")
        sys.exit(1)

    # Report paths for the caller
    print(f"\n[OK] Refreshed deck ready at: {refreshed_pptx.relative_to(REPO_ROOT)}")
    print(f"  size: {refreshed_pptx.stat().st_size / 1024:.0f} KB")
    print(f"\nNext: register it in tests/evals/fixtures.py::REFRESHED_DECKS and")
    print(f"      run `python -m tests.evals.end_to_end.generate_golden`.")


if __name__ == "__main__":
    main()
