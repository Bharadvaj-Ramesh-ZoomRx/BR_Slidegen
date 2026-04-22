"""
PS_test_synapse_refresh.py — CREON deck -> spec -> deck round-trip test.

Tests the core refresh-deck-workflow using the CREON project:
  Step 1: Deck -> Spec  (read_deck -> 91 SlideSpec JSONs in projects/CREON/specs/)
  Step 2: Spec -> Deck  (PS_slide_refresher.refresh_deck -> CREON_roundtrip_vN.pptx)
  Step 3: Verify        (shape counts + format code preservation)

Step 4 (Synapse live fetch) is excluded until the API key is refreshed (401 on all
22 analyses as of 2026-04-21 — see memory: project_creon_deck_refresh.md).

CREON facts (from prior testing on galen-consulting-r3m-report vijay-slidegen branch):
  - 91 slides, 1606 shapes scanned
  - 380 tagged (Tier 1), 1226 untagged (Tier 2)
  - 42 specs: spec_completeness="complete", 49: "layout_complete_data_missing"
  - 22 unique Synapse analysis IDs

Changes from tests/test_synapse_refresh.py (original — Repatha-hardcoded, untouched):
  - Uses CREON.pptx instead of Repatha deck
  - Uses PS_slide_refresher.refresh_deck() with auto_version=True
  - Saves all 91 specs to projects/CREON/specs/
  - No Synapse API fetch step (Step 4 deferred)
  - verify_round_trip includes format-code check (new in PS_slide_refresher)
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO))

from slidegen.deck_reader import read_deck
from slidegen.slide_spec.schema import dump_spec
from slidegen.PS_slide_refresher import refresh_deck, verify_round_trip, next_output_version

# ── Paths ─────────────────────────────────────────────────────────────────────

CREON_PPTX = REPO / "projects" / "CREON" / "CREON.pptx"
SPECS_DIR = REPO / "projects" / "CREON" / "specs"
OUTPUT_BASE = REPO / "projects" / "CREON" / "output" / "CREON_roundtrip.pptx"


def step1_deck_to_spec() -> tuple[list, dict]:
    """Step 1: Read CREON.pptx -> SlideSpec objects -> save JSON to specs/."""
    print("=" * 60)
    print("STEP 1: Deck -> Spec  (read_deck)")
    print(f"  Source: {CREON_PPTX}")

    if not CREON_PPTX.exists():
        raise FileNotFoundError(f"CREON.pptx not found: {CREON_PPTX}")

    specs, summary = read_deck(str(CREON_PPTX))
    t1 = summary["tier1"]
    t2 = summary["tier2"]

    print(f"  Slides: {summary['total_slides']}")
    print(f"  Shapes scanned: {summary['total_shapes_scanned']}")
    print(f"  Tier 1 — tagged: {t1.tagged_shapes}, "
          f"report configs: {t1.report_configs_resolved}, "
          f"pivot configs: {t1.pivot_configs_resolved}")
    print(f"  Tier 2 — inferred: {getattr(t2, 'inferred_shapes', '?')}")

    # Count completeness levels
    complete = sum(1 for s in specs if s.spec_completeness == "complete")
    layout_only = sum(1 for s in specs
                      if s.spec_completeness == "layout_complete_data_missing")
    partial = sum(1 for s in specs if s.spec_completeness == "partial")
    print(f"  Spec completeness: complete={complete}, "
          f"layout_only={layout_only}, partial={partial}")

    # Save specs to JSON
    SPECS_DIR.mkdir(parents=True, exist_ok=True)
    saved = 0
    for spec in specs:
        spec_path = SPECS_DIR / f"slide_{spec.slide_index:03d}.json"
        spec_json = dump_spec(spec)
        spec_path.write_text(spec_json, encoding="utf-8")
        saved += 1

    print(f"  Saved: {saved} spec JSONs to {SPECS_DIR}")

    return specs, summary


def step2_spec_to_deck(specs: list) -> str:
    """Step 2: Specs -> clone + refresh CREON.pptx -> versioned output PPTX."""
    print("\n" + "=" * 60)
    print("STEP 2: Spec -> Deck  (PS_slide_refresher.refresh_deck)")

    output_path = next_output_version(OUTPUT_BASE)
    print(f"  Output: {output_path}")

    result = refresh_deck(
        source_path=str(CREON_PPTX),
        output_path=str(output_path),
        specs=specs,
        auto_version=False,   # already versioned above
    )

    print(f"  Total slides: {result.total_slides}")
    print(f"  Charts refreshed: {result.charts_refreshed}")
    print(f"  Charts failed: {result.charts_failed}")
    if result.errors:
        for e in result.errors[:5]:
            print(f"    ERROR: {e}")
        if len(result.errors) > 5:
            print(f"    ... and {len(result.errors) - 5} more")

    return str(output_path)


def step3_verify(output_path: str) -> dict:
    """Step 3: Verify round-trip — shape counts + format code preservation."""
    print("\n" + "=" * 60)
    print("STEP 3: Verify round-trip")

    result = verify_round_trip(
        source_path=str(CREON_PPTX),
        output_path=output_path,
        check_formats=True,
    )

    slides_ok = "OK" if result["slides_match"] else "MISMATCH"
    print(f"  Slides: {result['orig_slides']} -> {result['rend_slides']}  [{slides_ok}]")
    print(f"  Shape issues: {len(result['slide_issues'])}")
    print(f"  Format issues: {len(result['format_issues'])}")
    print(f"  Total issues: {result['total_issues']}")

    if result["slide_issues"]:
        for issue in result["slide_issues"][:5]:
            print(f"    Slide {issue['slide']}: {', '.join(issue['issues'])}")

    if result["format_issues"]:
        for fi in result["format_issues"][:3]:
            print(f"    Slide {fi['slide']} @ {fi['position']}: "
                  f"formats changed {fi['src_formats'][:2]} -> {fi['out_formats'][:2]}")

    return result


def main():
    print("\n" + "=" * 60)
    print("CREON DECK -> SPEC -> DECK ROUND-TRIP TEST")
    print("=" * 60)

    # Step 1: Deck -> Spec
    specs, summary = step1_deck_to_spec()

    # Step 2: Spec -> Deck
    output_path = step2_spec_to_deck(specs)

    # Step 3: Verify
    verify = step3_verify(output_path)

    # Summary
    t1 = summary["tier1"]
    print("\n" + "=" * 60)
    print("SUMMARY")
    print(f"  Source: {CREON_PPTX.name}")
    print(f"  Specs generated: {len(specs)} ({t1.tagged_shapes} Tier 1 tagged shapes)")
    print(f"  Output: {Path(output_path).name}")
    print(f"  Shape issues: {len(verify['slide_issues'])}")
    print(f"  Format issues: {len(verify['format_issues'])}")

    passed = (verify["slides_match"] and
              verify["total_issues"] == 0)
    print(f"  Round-trip: {'PASS' if passed else 'FAIL'}")
    print("=" * 60)
    return 0 if passed else 1


if __name__ == "__main__":
    sys.exit(main())
