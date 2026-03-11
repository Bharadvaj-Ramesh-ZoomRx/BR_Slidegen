"""
registry.py — Shape registry CRUD operations.

The registry (slide_registry.json) tracks all named shapes in the deck,
enabling targeted COM edits and per-slide regeneration.
"""

from __future__ import annotations

import json
import os
from datetime import datetime
from typing import Optional

from .brand import IN

# Default registry path (fallback)
try:
    from slidegen.config import REGISTRY_PATH
except ImportError:
    REGISTRY_PATH = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "slide_registry.json"
    )


def load_registry(path: Optional[str] = None) -> dict:
    """Load slide_registry.json and return the dict."""
    if path is None:
        path = REGISTRY_PATH
    with open(path) as f:
        return json.load(f)


def save_registry(registry: dict, path: Optional[str] = None) -> None:
    """Save registry dict back to slide_registry.json."""
    if path is None:
        path = REGISTRY_PATH
    with open(path, "w") as f:
        json.dump(registry, f, indent=2)


def registry_get(name: str, path: Optional[str] = None) -> dict:
    """Return the record for a single shape by name. Raises KeyError if not found."""
    reg = load_registry(path)
    if name not in reg["shapes"]:
        raise KeyError(
            f"Shape '{name}' not in registry. "
            f"Run 'python -m slidegen create' or 'python -m slidegen reconcile' first."
        )
    return reg["shapes"][name]


def registry_tag_slide(slide_idx: int, module: str, section: str,
                       data_source: str = "", path: Optional[str] = None) -> None:
    """Store a slide-level metadata record in the registry.

    Enables 'rebuild slide 14 from scratch' without touching other slides.
    Records are stored under reg['slides'][str(slide_idx)].

    Args:
        slide_idx: int — 0-based slide index
        module: module name (e.g. "Personal Promotion Module")
        section: section name (e.g. "Key Findings - Personal Promotions")
        data_source: data source description (e.g. "Lung SFEA SB.xlsx / RYB sheet")
        path: registry file path (default REGISTRY_PATH)
    """
    reg = load_registry(path)
    if "slides" not in reg:
        reg["slides"] = {}
    reg["slides"][str(slide_idx)] = {
        "module":      module,
        "section":     section,
        "data_source": data_source,
        "tagged_at":   datetime.now().isoformat(),
    }
    save_registry(reg, path)


def registry_find_by_type(shape_type: str, slide_idx: Optional[int] = None,
                          path: Optional[str] = None) -> list[tuple[str, dict]]:
    """Query the registry for all shapes matching a type string.

    Enables bulk COM operations such as 'update all footer text on all slides'
    without knowing individual shape names.

    Args:
        shape_type: string — "chart", "textbox", "rect", "image", etc.
        slide_idx: int or None — if set, restrict results to one slide
        path: registry file path (default REGISTRY_PATH)
    Returns: list of (name, record) tuples matching the query
    """
    reg    = load_registry(path)
    shapes = reg.get("shapes", {})
    return [
        (name, record)
        for name, record in shapes.items()
        if record.get("type") == shape_type
        and (slide_idx is None or record.get("slide_idx") == slide_idx)
    ]


def registry_diff_slide(slide_idx: int, com_slide, path: Optional[str] = None) -> list[dict]:
    """Diff the registry snapshot against live COM state for a single slide.

    More ergonomic than reconcile_registry.py when you only care about one slide.

    Args:
        slide_idx: int — 0-based slide index
        com_slide: win32com Slide object (pass prs.Slides(slide_idx + 1))
        path: registry file path (default REGISTRY_PATH)
    Returns: list of dicts with keys: name, field, registry_val, live_val
    """
    reg    = load_registry(path)
    shapes = reg.get("shapes", {})

    live = {}
    for i in range(1, com_slide.Shapes.Count + 1):
        sh = com_slide.Shapes(i)
        live[sh.Name] = {
            "left":   round(sh.Left   / IN, 4),
            "top":    round(sh.Top    / IN, 4),
            "width":  round(sh.Width  / IN, 4),
            "height": round(sh.Height / IN, 4),
        }

    diffs = []
    for name, record in shapes.items():
        if record.get("slide_idx") != slide_idx:
            continue
        if name not in live:
            diffs.append({"name": name, "field": "existence",
                          "registry_val": "present", "live_val": "missing"})
            continue
        for field in ("left", "top", "width", "height"):
            r_val = record.get(field)
            l_val = live[name].get(field)
            if r_val is not None and l_val is not None:
                if abs(r_val - l_val) > 0.01:
                    diffs.append({"name": name, "field": field,
                                  "registry_val": r_val, "live_val": l_val})
    return diffs
