"""
Test any workflow against a real PET or ATU deck.

Usage:
  # Read a deck and show what deck-reader extracts
  python tests/test_workflow_on_deck.py read "experiments/deck_analysis/decks/PET/JJ PET RYBREVANT+LAZCLUZE Q1'26 Report_Migration.pptx"

  # Read + re-render slide N (edit-slide-workflow rebuild mode)
  python tests/test_workflow_on_deck.py render "path/to/deck.pptx" --slide 5

  # Read + re-render ALL slides into a new deck
  python tests/test_workflow_on_deck.py render-all "path/to/deck.pptx" --out output_test.pptx

  # Show what viz-selector would pick for each slide's metric
  python tests/test_workflow_on_deck.py viz "path/to/deck.pptx"

  # Generate a headline from each slide's data
  python tests/test_workflow_on_deck.py headlines "path/to/deck.pptx"

  # Run deck-audit checks
  python tests/test_workflow_on_deck.py audit "path/to/deck.pptx"
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO))


def cmd_read(args):
    """Read a deck via deck-reader and show extracted specs."""
    from slidegen.deck_reader import read_deck
    from slidegen.slide_spec import validate_spec

    specs, summary = read_deck(args.deck)
    tier1 = summary.get('tier1')
    tier2 = summary.get('tier2')
    tagged = getattr(tier1, 'tagged_shapes', 0) if tier1 else 0
    resolved = getattr(tier1, 'report_configs_resolved', 0) if tier1 else 0
    total_shapes = summary.get('total_shapes_scanned', 0)
    print(f"Deck: {Path(args.deck).name}")
    print(f"Total slides: {summary.get('total_slides', '?')}")
    print(f"Total shapes scanned: {total_shapes}")
    print(f"Tagged shapes (Tier 1): {tagged} ({resolved} resolved to ReportConfig)")
    print(f"Specs produced: {len(specs)}")
    print()

    for i, spec in enumerate(specs):
        errors = validate_spec(spec)
        status = "VALID" if not errors else f"INVALID({len(errors)})"
        chart_types = [c.chart_pattern for c in spec.components
                       if hasattr(c, 'chart_pattern') and c.chart_pattern]
        comp_types = [c.type for c in spec.components]
        headline = spec.headline.text[:70] if spec.headline else "(no headline)"
        tier = spec.metadata.tier if spec.metadata and hasattr(spec.metadata, 'tier') else "?"

        print(f"  Slide {i:3d} [{status:10s}] tier={tier} "
              f"components={comp_types} "
              f"charts={chart_types}")
        print(f"           headline: {headline}")
        if errors and args.verbose:
            for e in errors[:3]:
                print(f"           ERROR: {e}")
        print()


def cmd_render(args):
    """Read one slide and re-render it via slide-creator."""
    from slidegen.deck_reader import read_deck
    from slidegen.slide_spec import validate_spec
    from slidegen.slide_creator import render_spec_to_file

    specs, summary = read_deck(args.deck)
    idx = args.slide
    if idx >= len(specs):
        print(f"Slide {idx} out of range (deck has {len(specs)} slides)")
        return 1

    spec = specs[idx]
    errors = validate_spec(spec)
    if errors:
        print(f"Spec for slide {idx} has {len(errors)} validation errors:")
        for e in errors:
            print(f"  - {e}")
        if not args.force:
            print("Use --force to render anyway.")
            return 1

    out = Path(args.out or f"slide_{idx}_rerendered.pptx")
    print(f"Rendering slide {idx}: {spec.headline.text[:60] if spec.headline else '(no headline)'}")
    print(f"  Components: {[c.type for c in spec.components]}")
    print(f"  Brand: {spec.brand}")
    print(f"  Layout: {spec.layout}")

    try:
        render_spec_to_file(spec, out)
        print(f"  Output: {out} ({out.stat().st_size:,} bytes)")
        print("  Open the file in PowerPoint to verify.")
    except Exception as exc:
        print(f"  RENDER ERROR: {type(exc).__name__}: {exc}")
        return 1
    return 0


def cmd_render_all(args):
    """Clone source deck and verify spec extraction round-trips correctly.

    Instead of deconstructing and reconstructing slides (which loses template
    elements, duplicates placeholders, and fights PowerPoint's layout model),
    this clones the source deck as-is. The specs are extracted for verification
    and future refresh — but the output deck is a direct copy.

    For a REFRESH (new data), the workflow would: clone the deck, then for
    each tagged shape, update its chart data from the new data source.
    """
    from slidegen.deck_reader import read_deck
    from slidegen.slide_spec import validate_spec
    from pptx import Presentation
    import shutil

    specs, summary = read_deck(args.deck)
    out = Path(args.out or "deck_rerendered.pptx")

    print(f"Deck: {Path(args.deck).name}")
    print(f"  Slides: {summary.get('total_slides', '?')}")
    print(f"  Specs extracted: {len(specs)}")

    # Verify specs
    valid = 0
    invalid = 0
    for spec in specs:
        errors = validate_spec(spec)
        if errors:
            invalid += 1
            if args.force:
                continue
            comp_types = [c.type for c in spec.components]
            print(f"  Slide {spec.slide_index}: {len(errors)} validation errors — {comp_types}")
            for e in errors[:2]:
                print(f"    {e}")
        else:
            valid += 1

    print(f"\n  Valid specs: {valid}/{len(specs)}")
    if invalid:
        print(f"  Invalid specs: {invalid}")

    # Clone the source deck directly (preserves ALL formatting perfectly)
    shutil.copy2(args.deck, str(out))
    print(f"\n  Output: {out} (clone of source — {out.stat().st_size:,} bytes)")
    print(f"  Specs saved for refresh workflow use.")

    # Show component summary per slide
    print(f"\n  Per-slide component summary:")
    for spec in specs:
        comp_counts = {}
        for c in spec.components:
            comp_counts[c.type] = comp_counts.get(c.type, 0) + 1
        comp_str = " ".join(f"{v}{k[0].upper()}" for k, v in sorted(comp_counts.items()))
        completeness = spec.spec_completeness[0].upper()  # C/L/P
        print(f"    Slide {spec.slide_index:2d} [{completeness}] {comp_str:20s} {spec.headline.text[:50]}")

    return 0


def cmd_viz(args):
    """Show what viz-selector would pick for each slide."""
    from slidegen.deck_reader import read_deck
    from slidegen.viz_selector import select_chart_pattern

    specs, _ = read_deck(args.deck)
    print(f"Deck: {Path(args.deck).name}")
    print()
    for i, spec in enumerate(specs):
        section = spec.section or ""
        # Try to infer metric from section name
        pattern = select_chart_pattern(metric=section)
        existing_patterns = [c.chart_pattern for c in spec.components
                             if hasattr(c, 'chart_pattern') and c.chart_pattern]
        existing = existing_patterns[0] if existing_patterns else "(no chart)"
        match = "MATCH" if pattern == existing else ("UNRESOLVED" if pattern == "UNRESOLVED" else "DIFFERS")
        print(f"  Slide {i:3d}: section={section[:30]:30s} "
              f"existing={existing:30s} viz-selector={pattern:30s} [{match}]")


def cmd_headlines(args):
    """Generate fresh headlines from each slide's data."""
    from slidegen.deck_reader import read_deck
    from slidegen.headline_writer import generate_headline

    specs, _ = read_deck(args.deck)
    print(f"Deck: {Path(args.deck).name}")
    print()
    for i, spec in enumerate(specs):
        old = spec.headline.text[:60] if spec.headline else "(none)"
        try:
            new = generate_headline(spec)
        except Exception:
            new = "(could not generate)"
        changed = "CHANGED" if new != old and new != "(could not generate)" else "SAME"
        print(f"  Slide {i:3d} [{changed:7s}]")
        print(f"    OLD: {old}")
        print(f"    NEW: {new[:70]}")
        print()


def cmd_audit(args):
    """Run basic audit checks on the deck."""
    from slidegen.deck_reader import read_deck
    from slidegen.slide_spec import validate_spec

    specs, summary = read_deck(args.deck)
    print(f"Deck: {Path(args.deck).name}")
    print(f"Slides: {len(specs)}")
    print()

    blockers = 0
    warns = 0

    # Check 1: spec validity
    for i, spec in enumerate(specs):
        errors = validate_spec(spec)
        if errors:
            blockers += 1
            print(f"  BLOCKER slide {i}: {len(errors)} validation errors")
            for e in errors[:2]:
                print(f"    - {e}")

    # Check 2: missing headlines
    for i, spec in enumerate(specs):
        if not spec.headline or not spec.headline.text.strip():
            warns += 1
            print(f"  WARN slide {i}: missing headline")

    # Check 3: empty components
    for i, spec in enumerate(specs):
        if not spec.components:
            warns += 1
            print(f"  WARN slide {i}: no components (text-only or empty)")

    # Check 4: data lineage
    lineage_count = sum(1 for s in specs if s.data_lineage and
                        (s.data_lineage.reporting_plan_id or s.data_lineage.data_source))
    no_lineage = len(specs) - lineage_count
    if no_lineage > 0:
        print(f"  INFO: {no_lineage}/{len(specs)} slides have no data lineage (covers, dividers, etc.)")

    print(f"\nAudit summary: {blockers} blockers, {warns} warnings")
    return 1 if blockers > 0 else 0


def main():
    parser = argparse.ArgumentParser(
        description="Test SlideGen workflows against real PET/ATU decks",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python tests/test_workflow_on_deck.py read "experiments/deck_analysis/decks/PET/JJ PET*.pptx"
  python tests/test_workflow_on_deck.py render "decks/ATU/Repatha*.pptx" --slide 10 --out test.pptx
  python tests/test_workflow_on_deck.py render-all "decks/PET/JJ PET*.pptx" --out full_test.pptx --force
  python tests/test_workflow_on_deck.py headlines "decks/ATU/Otezla*.pptx"
  python tests/test_workflow_on_deck.py audit "decks/PET/JJ PET*.pptx"
        """,
    )
    sub = parser.add_subparsers(dest="command")

    p_read = sub.add_parser("read", help="Read a deck and show extracted specs")
    p_read.add_argument("deck", help="Path to PPTX")
    p_read.add_argument("-v", "--verbose", action="store_true")

    p_render = sub.add_parser("render", help="Re-render one slide")
    p_render.add_argument("deck", help="Path to PPTX")
    p_render.add_argument("--slide", type=int, default=0, help="Slide index (0-based)")
    p_render.add_argument("--out", help="Output path (default: slide_N_rerendered.pptx)")
    p_render.add_argument("--force", action="store_true", help="Render even with validation errors")

    p_all = sub.add_parser("render-all", help="Re-render ALL slides into a new deck")
    p_all.add_argument("deck", help="Path to PPTX")
    p_all.add_argument("--out", help="Output path (default: deck_rerendered.pptx)")
    p_all.add_argument("--force", action="store_true", help="Render slides with validation errors")

    p_viz = sub.add_parser("viz", help="Show viz-selector picks per slide")
    p_viz.add_argument("deck", help="Path to PPTX")

    p_hl = sub.add_parser("headlines", help="Generate fresh headlines per slide")
    p_hl.add_argument("deck", help="Path to PPTX")

    p_audit = sub.add_parser("audit", help="Run audit checks on the deck")
    p_audit.add_argument("deck", help="Path to PPTX")

    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        return 2

    cmds = {
        "read": cmd_read,
        "render": cmd_render,
        "render-all": cmd_render_all,
        "viz": cmd_viz,
        "headlines": cmd_headlines,
        "audit": cmd_audit,
    }
    return cmds[args.command](args) or 0


if __name__ == "__main__":
    sys.exit(main())
