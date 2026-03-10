"""
orchestrator.py — Ties config + data + renderers into a complete deck.

Usage:
    from slidegen.pipeline import generate_deck
    generate_deck("projects/jnj_rybrevant.yaml")
"""

from __future__ import annotations

import os
from pptx import Presentation
from pptx.util import Emu

from slidegen.pipeline.project_config import load_project_config, ProjectConfig
from slidegen.pipeline.data_loaders import load_all_data
from slidegen.pipeline.slide_renderers import RENDERERS


def generate_deck(yaml_path: str, output_path: str | None = None) -> str:
    """Generate a full slide deck from a YAML project config.

    Args:
        yaml_path: Path to the project YAML file.
        output_path: Override output path. If None, uses config.output_path.

    Returns:
        Path to the generated PPTX file.
    """
    # 1. Load config
    print(f"Loading config: {yaml_path}")
    config = load_project_config(yaml_path)

    # 2. Load data
    print("Extracting data...")
    data = load_all_data(config)

    # Print extraction summary
    for key, val in data.items():
        if key.startswith("_"):
            continue
        count = len(val) if isinstance(val, list) else "—"
        print(f"  {key}: {count} rows")

    # 3. Create presentation
    print("\nCreating presentation...")
    # Always create a fresh presentation with standard widescreen dimensions.
    # Template layouts could be used but slide removal is fragile in python-pptx.
    prs = Presentation()
    prs.slide_width = Emu(12192000)   # 13.33 inches
    prs.slide_height = Emu(6858000)   # 7.50 inches
    if config.template_path and os.path.exists(config.template_path):
        print(f"  Template available: {os.path.basename(config.template_path)} (not loaded — using blank)")
    else:
        print("  Blank presentation")

    # 4. Build slides
    print("\nBuilding slides...")
    for i, ask in enumerate(config.asks, 1):
        renderer = RENDERERS.get(ask.slide_type)
        if renderer is None:
            print(f"  [{i}] SKIP — unknown slide_type: {ask.slide_type}")
            continue

        # Add blank slide
        layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
        slide = prs.slides.add_slide(layout)

        try:
            renderer(slide, config, ask, data)
            print(f"  [{i}] {ask.id} ({ask.slide_type})")
        except Exception as e:
            print(f"  [{i}] ERROR on {ask.id}: {e}")
            from slidegen.pptx_utils import textbox, C_RED
            textbox(slide, f"Error: {e}", 1, 3, 10, 1, fsize=12, color=C_RED)

    # 5. Save
    out = output_path or config.output_path
    if not out:
        out = os.path.join(os.path.dirname(yaml_path), "output_deck.pptx")

    os.makedirs(os.path.dirname(out), exist_ok=True)
    prs.save(out)
    print(f"\nSaved: {out}")
    print(f"Total slides: {len(prs.slides)}")

    return out


# ── CLI entry point ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python -m slidegen.pipeline.orchestrator <project.yaml> [output.pptx]")
        sys.exit(1)

    yaml_path = sys.argv[1]
    output = sys.argv[2] if len(sys.argv) > 2 else None
    generate_deck(yaml_path, output)
