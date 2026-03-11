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

import hashlib
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

    def __init__(self, slide_idx: int, ask_id: str = "",
                 data_key: str = "", config_hash: str = "",
                 source_file: str = ""):
        self._slide_idx = slide_idx
        self._counter = 0
        self._registry: dict[str, dict] = {}
        self._ask_id = ask_id
        self._data_key = data_key
        self._config_hash = config_hash
        self._source_file = source_file

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

    @property
    def slide_metadata(self) -> dict:
        """Return per-slide metadata including data_source lineage (PRD §6.2)."""
        meta = {
            "ask_id": self._ask_id,
            "generated_at": datetime.now().isoformat(),
            "shapes": self._registry,
        }
        if self._data_key or self._source_file:
            meta["data_source"] = {
                "data_key": self._data_key,
                "config_hash": self._config_hash,
                "source_file": self._source_file,
            }
        return meta


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


# ── PPTX backup (PRD §10.3) ──────────────────────────────────────────────────

_MAX_PPTX_BACKUPS = 10


def _backup_pptx(pptx_path: str) -> str | None:
    """Create a timestamped backup of the deck before editing.

    Stores up to _MAX_PPTX_BACKUPS most recent copies in a backups/ folder
    alongside the deck. Oldest backups are deleted when the cap is exceeded.

    Returns the backup path, or None if the source doesn't exist.
    """
    if not os.path.exists(pptx_path):
        return None

    backup_dir = os.path.join(os.path.dirname(pptx_path), "backups")
    os.makedirs(backup_dir, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.splitext(os.path.basename(pptx_path))[0]
    backup_path = os.path.join(backup_dir, f"{base}_{ts}.pptx")
    shutil.copy2(pptx_path, backup_path)

    # Prune old backups
    backups = sorted(
        [f for f in os.listdir(backup_dir) if f.endswith(".pptx")],
        key=lambda f: os.path.getmtime(os.path.join(backup_dir, f)),
        reverse=True,
    )
    for old in backups[_MAX_PPTX_BACKUPS:]:
        os.remove(os.path.join(backup_dir, old))

    return backup_path


# ── Data cache (JSON) ────────────────────────────────────────────────────────

def _data_cache_path(config: ProjectConfig) -> str:
    """Return the path to the JSON data cache for the current wave."""
    output_dir = os.path.dirname(config.output_path)
    return os.path.join(output_dir, "slide_data.json")


def _config_hash(yaml_path: str) -> str:
    """Fast hash of config file for cache invalidation."""
    with open(yaml_path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()[:12]


def _save_data_cache(data: dict, config: ProjectConfig, yaml_path: str):
    """Save extracted data as JSON cache alongside the deck."""
    cache_path = _data_cache_path(config)
    os.makedirs(os.path.dirname(cache_path), exist_ok=True)

    cache = {
        "_meta": {
            "wave": config.wave,
            "source": config.data_source_path,
            "extracted_at": datetime.now().isoformat(),
            "config_hash": _config_hash(yaml_path),
            "source_mtime": os.path.getmtime(config.data_source_path),
        },
    }
    cache.update(data)

    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(cache, f, indent=2, default=str)

    print(f"  Data cache saved: {cache_path}")


def _load_data_cache(config: ProjectConfig, yaml_path: str) -> dict | None:
    """Load data from JSON cache if valid (config unchanged, Excel not newer).

    Returns None if cache is stale or missing.
    """
    cache_path = _data_cache_path(config)
    if not os.path.exists(cache_path):
        return None

    with open(cache_path, "r", encoding="utf-8") as f:
        cache = json.load(f)

    meta = cache.get("_meta", {})

    # Check config hasn't changed
    if meta.get("config_hash") != _config_hash(yaml_path):
        print("  Data cache stale (config changed) — re-extracting")
        return None

    # Check Excel file hasn't been modified
    if os.path.exists(config.data_source_path):
        excel_mtime = os.path.getmtime(config.data_source_path)
        if excel_mtime > meta.get("source_mtime", 0):
            print("  Data cache stale (Excel updated) — re-extracting")
            return None

    # Cache is valid — strip _meta before returning
    data = {k: v for k, v in cache.items() if k != "_meta"}
    print(f"  Data loaded from cache ({cache_path})")
    return data


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

    # 2. Load data (always re-extract for full deck, then cache)
    print("Extracting data...")
    data = load_all_data(config)

    # Print extraction summary
    for key, val in data.items():
        if key.startswith("_"):
            continue
        count = len(val) if isinstance(val, list) else "—"
        print(f"  {key}: {count} rows")

    # Save data cache for future regenerate_slide() calls
    _save_data_cache(data, config, yaml_path)

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

        cfg_hash = _config_hash(yaml_path)
        namer = ShapeNamer(
            i, ask_id=ask.id,
            data_key=getattr(ask, 'data_key', ''),
            config_hash=cfg_hash,
            source_file=config.data_source_path,
        )
        try:
            renderer(slide, config, ask, data, namer=namer)
            namer.name_remaining(slide)
            slide_registries[str(i)] = namer.slide_metadata
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

    # Try cached data first (fast); fall back to Excel extraction
    data = _load_data_cache(config, yaml_path)
    if data is None:
        print("Extracting data from Excel...")
        data = load_all_data(config)
        _save_data_cache(data, config, yaml_path)

    pptx_path = output_path or config.output_path
    if not pptx_path or not os.path.exists(pptx_path):
        raise FileNotFoundError(f"Deck not found: {pptx_path}")

    # Backup before editing (PRD §10.3)
    backup = _backup_pptx(pptx_path)
    if backup:
        print(f"  Backup saved: {backup}")

    prs = Presentation(pptx_path)
    if slide_index < 0 or slide_index >= len(prs.slides):
        raise IndexError(f"Slide index {slide_index} out of range (deck has {len(prs.slides)} slides)")

    slide = prs.slides[slide_index]
    _clear_slide(slide)

    ask = config.asks[slide_index]
    renderer = RENDERERS.get(ask.slide_type)
    if renderer is None:
        raise ValueError(f"Unknown slide_type: {ask.slide_type}")

    namer = ShapeNamer(
        slide_index + 1, ask_id=ask.id,
        data_key=getattr(ask, 'data_key', ''),
        config_hash=_config_hash(yaml_path),
        source_file=config.data_source_path,
    )
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
