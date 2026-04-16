"""
Render every canonical SlideSpec example to PPTX for visual review.

Loads every *.json in slidegen/slide_spec/examples/, validates each,
renders to experiments/deck_analysis/outputs/generated/canonical_<pattern>.pptx,
and reports per-file: spec validity, file size, chart count.

Usage:
    python tests/track1/render_canonical_examples.py
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

# Ensure project root is on sys.path
PROJECT_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(PROJECT_ROOT))

from slidegen.slide_spec import load_spec, validate_spec
from slidegen.slide_creator import render_spec_to_file


def main():
    examples_dir = PROJECT_ROOT / "slidegen" / "slide_spec" / "examples"
    output_dir = PROJECT_ROOT / "experiments" / "deck_analysis" / "outputs" / "generated"
    output_dir.mkdir(parents=True, exist_ok=True)

    json_files = sorted(examples_dir.glob("*.json"))
    if not json_files:
        print("ERROR: No *.json files found in", examples_dir)
        sys.exit(1)

    print(f"Found {len(json_files)} canonical example(s) in {examples_dir}\n")
    print(f"{'File':<45} {'Valid?':<8} {'Rendered?':<10} {'Size':>8}  {'Charts':>6}")
    print("-" * 85)

    total = 0
    passed = 0
    failed_names = []

    for json_path in json_files:
        total += 1
        name = json_path.stem
        # Derive pattern name from _pattern field or filename
        try:
            spec = load_spec(json_path)
        except Exception as e:
            print(f"{name + '.json':<45} {'LOAD-ERR':<8} {'--':<10} {'--':>8}  {'--':>6}")
            print(f"    Error: {e}")
            failed_names.append(name)
            continue

        # Validate
        errors = validate_spec(spec, strict=False)
        valid_str = "OK" if not errors else f"FAIL({len(errors)})"

        if errors:
            print(f"{name + '.json':<45} {valid_str:<8} {'--':<10} {'--':>8}  {'--':>6}")
            for err in errors:
                print(f"    {err}")
            failed_names.append(name)
            continue

        # Output name uses the spec filename (stem) so composition-variant
        # specs that share a chart_pattern don't collide on the same output.
        out_path = output_dir / f"canonical_{name}.pptx"

        # Render
        try:
            render_spec_to_file(spec, out_path)
            file_size = out_path.stat().st_size
            size_str = f"{file_size / 1024:.1f} KB"

            # Count charts in spec
            chart_count = sum(
                1 for c in spec.components if hasattr(c, "chart_pattern")
            )

            print(f"{name + '.json':<45} {valid_str:<8} {'OK':<10} {size_str:>8}  {chart_count:>6}")
            passed += 1
        except Exception as e:
            print(f"{name + '.json':<45} {valid_str:<8} {'RENDER-ERR':<10} {'--':>8}  {'--':>6}")
            print(f"    Error: {e}")
            failed_names.append(name)

    print("-" * 85)
    print(f"\nTotal: {total}  |  Passed: {passed}  |  Failed: {total - passed}")
    if failed_names:
        print(f"Failed: {', '.join(failed_names)}")
    print(f"\nOutput dir: {output_dir}")

    sys.exit(0 if not failed_names else 1)


if __name__ == "__main__":
    main()
