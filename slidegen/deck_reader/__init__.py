"""
deck_reader — Dual-mode PPTX-to-SlideSpec extractor.

Tier 1 (tag_reader): Reads Galen-PowerPoint Connector shape tags + Custom XML Parts.
Tier 2 (inference):  Structural inference for untagged shapes (user-confirmed).

Unified API:
    from slidegen.deck_reader import read_deck
    specs = read_deck("path/to/deck.pptx")
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Optional

from .tag_reader import read_tagged_shapes, TagReaderSummary
from .inference import infer_untagged_shapes, InferenceSummary


def read_deck(
    pptx_path: str | Path,
    config_path: Optional[str | Path] = None,
    source_data_path: Optional[str | Path] = None,
) -> tuple[list, dict]:
    """Unified deck reader: Tier 1 first, Tier 2 for untagged shapes.

    Args:
        pptx_path: Path to the PPTX file.
        config_path: Optional YAML config for Tier 2 inference hints.
        source_data_path: Optional source_data.json for Tier 2 cross-referencing.

    Returns:
        (specs, summary) where:
            specs: list[SlideSpec] — merged from Tier 1 + Tier 2
            summary: dict with keys:
                - tier1: TagReaderSummary (tagged shape counts)
                - tier2: InferenceSummary (inferred shape counts)
                - total_slides: int
                - total_shapes_scanned: int
    """
    pptx_path = Path(pptx_path)
    if not pptx_path.exists():
        raise FileNotFoundError(f"PPTX not found: {pptx_path}")

    # Tier 1: tag-driven extraction
    tier1_specs, tier1_summary, untagged_shapes = read_tagged_shapes(pptx_path)

    # Tier 2: structural inference for untagged shapes
    tier2_specs, tier2_summary = infer_untagged_shapes(
        pptx_path=pptx_path,
        untagged_shapes=untagged_shapes,
        config_path=config_path,
        source_data_path=source_data_path,
    )

    # Merge: combine Tier 1 + Tier 2 components for the same slide.
    # A slide may have both tagged shapes (Tier 1) and untagged shapes (Tier 2).
    # The merged spec gets all components, with Tier 1 lineage taking precedence.
    tier1_by_idx = {s.slide_index: s for s in tier1_specs}
    tier2_by_idx = {s.slide_index: s for s in tier2_specs}

    all_indices = sorted(set(tier1_by_idx.keys()) | set(tier2_by_idx.keys()))
    merged = []
    for idx in all_indices:
        t1 = tier1_by_idx.get(idx)
        t2 = tier2_by_idx.get(idx)
        if t1 and t2:
            # Both tiers have components for this slide — merge components
            t1.components.extend(t2.components)
            merged.append(t1)
        elif t1:
            merged.append(t1)
        else:
            merged.append(t2)

    summary = {
        "tier1": tier1_summary,
        "tier2": tier2_summary,
        "total_slides": tier1_summary.total_slides,
        "total_shapes_scanned": tier1_summary.total_shapes,
    }

    return merged, summary


__all__ = ["read_deck", "read_tagged_shapes", "infer_untagged_shapes"]
