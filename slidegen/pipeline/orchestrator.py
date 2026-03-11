"""
orchestrator.py — Ties config + data + renderers into a complete deck.

Usage:
    from slidegen.pipeline import generate_deck
    generate_deck("projects/jnj_rybrevant/config.yaml")

    # Regenerate a single slide:
    from slidegen.pipeline.orchestrator import regenerate_slide
    regenerate_slide("projects/jnj_rybrevant/config.yaml", slide_index=4)
"""

from __future__ import annotations

import json
import os
import shutil
from datetime import datetime
from pptx import Presentation
from pptx.util import Emu

from slidegen.pipeline.project_config import load_project_config, ProjectConfig
from slidegen.pipeline.data_loaders import load_all_data
from slidegen.pipeline.slide_renderers import RENDERERS


# ── Shape naming ─────────────────────────────────────────────────────────────

class ShapeNamer:
    """Assigns sequential zrx_ names to shapes on a slide.

    Names follow the pattern zrx_{slide_idx:03d}_{shape_num:03d}, e.g.
    zrx_001_001, zrx_001_002, ... for slide 1.
    """

    def __init__(self, slide_idx: int):
        self._slide_idx = slide_idx
        self._counter = 0
        self._registry: dict[str, dict] = {}

    def name(self, shape, label: str = "") -> str:
        """Assign a zrx_ name to a shape and record it in the registry."""
        self._counter += 1
        name = f"zrx_{self._slide_idx:03d}_{self._counter:03d}"
        shape.name = name
        self._registry[name] = {"label": label}
        return name

    def name_remaining(self, slide):
        """Name any shapes on the slide that don't yet have a zrx_ prefix."""
        for shape in slide.shapes:
            if not shape.name.startswith("zrx_"):
                self.name(shape, label="chrome")

    @property
    def registry(self) -> dict[str, dict]:
        return self._registry


def _save_shape_registry(registries: dict, output_dir: str):
    """Save the combined shape registry for all slides."""
    path = os.path.join(output_dir, "shape_registry.json")
    os.makedirs(output_dir, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump({
            "created": datetime.now().isoformat(),
            "slides": registries,
        }, f, indent=2)


# ── Config backup ────────────────────────────────────────────────────────────

def _backup_config(yaml_path: str) -> str:
    """Copy current config to config_history/ with timestamp."""
    project_dir = os.path.dirname(os.path.abspath(yaml_path))
    history_dir = os.path.join(project_dir, "config_history")
    os.makedirs(history_dir, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.splitext(os.path.basename(yaml_path))[0]
    backup_path = os.path.join(history_dir, f"{base}_{ts}.yaml")
    shutil.copy2(yaml_path, backup_path)
    return backup_path


# ── Slide clearing ───────────────────────────────────────────────────────────

def _clear_slide(slide):
    """Remove all shapes from a slide, preserving the slide itself."""
    sp_tree = slide.shapes._spTree
    removable = [
        '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}sp',
        '{http://schemas.openxmlformats.org/presentationml/2006/main}sp',
    ]
    for child in list(sp_tree):
        tag = child.tag
        if (tag.endswith('}sp') or tag.endswith('}graphicFrame') or
                tag.endswith('}pic') or tag.endswith('}grpSp') or
                tag.endswith('}cxnSp')):
            sp_tree.remove(child)


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
    slide_registries = {}
    print("\nBuilding slides...")
    for i, ask in enumerate(config.asks, 1):
        renderer = RENDERERS.get(ask.slide_type)
        if renderer is None:
            print(f"  [{i}] SKIP — unknown slide_type: {ask.slide_type}")
            continue

        # Add blank slide
        layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
        slide = prs.slides.add_slide(layout)

        namer = ShapeNamer(i)
        try:
            renderer(slide, config, ask, data, namer=namer)
            namer.name_remaining(slide)
            slide_registries[str(i)] = {"ask_id": ask.id, "shapes": namer.registry}
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
    _save_shape_registry(slide_registries, os.path.dirname(out))
    print(f"\nSaved: {out}")
    print(f"Total slides: {len(prs.slides)}")

    return out


def regenerate_slide(yaml_path: str, slide_index: int,
                     output_path: str | None = None) -> str:
    """Regenerate a single slide in an existing deck.

    Clears all shapes from the target slide and re-renders it from
    the current config and data. Other slides are untouched.

    Args:
        yaml_path: Path to project YAML config.
        slide_index: 0-based slide index to regenerate.
        output_path: Path to existing PPTX. If None, uses config.output_path.

    Returns:
        Path to the updated PPTX file.
    """
    config = load_project_config(yaml_path)
    data = load_all_data(config)

    pptx_path = output_path or config.output_path
    if not pptx_path or not os.path.exists(pptx_path):
        raise FileNotFoundError(f"Deck not found: {pptx_path}")

    prs = Presentation(pptx_path)
    if slide_index < 0 or slide_index >= len(prs.slides):
        raise IndexError(f"Slide index {slide_index} out of range (deck has {len(prs.slides)} slides)")

    slide = prs.slides[slide_index]
    _clear_slide(slide)

    ask = config.asks[slide_index]
    renderer = RENDERERS.get(ask.slide_type)
    if renderer is None:
        raise ValueError(f"Unknown slide_type: {ask.slide_type}")

    namer = ShapeNamer(slide_index + 1)
    renderer(slide, config, ask, data, namer=namer)
    namer.name_remaining(slide)

    prs.save(pptx_path)
    print(f"Regenerated slide {slide_index + 1} ({ask.id}) in {pptx_path}")
    return pptx_path


# ── CLI entry point ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python -m slidegen.pipeline.orchestrator <project.yaml> [output.pptx]")
        sys.exit(1)

    yaml_path = sys.argv[1]
    output = sys.argv[2] if len(sys.argv) > 2 else None
    generate_deck(yaml_path, output)
