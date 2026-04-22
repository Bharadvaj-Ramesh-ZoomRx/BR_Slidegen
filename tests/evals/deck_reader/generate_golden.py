"""One-time golden file generator for deck-reader evals.

Run this to capture current deck-reader output as ground truth:
    python -m tests.evals.deck_reader.generate_golden

Commit the resulting golden/*.json files. Future eval runs diff against these.
Re-run only when deck-reader intentionally changes (and review the diff before committing).
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT))

from slidegen.deck_reader.tag_reader import generate_config_specs  # noqa: E402
from tests.evals.fixtures import FIXTURE_DECKS, GOLDEN_DIR  # noqa: E402


def build_tag_counts(deck_path: Path) -> dict:
    """Produce the golden tag-count snapshot for a deck."""
    specs, summary = generate_config_specs(str(deck_path))

    per_slide = {}
    for slide_idx, counts in sorted(summary.per_slide.items()):
        per_slide[str(slide_idx + 1)] = {
            "tagged": counts.get("tagged", 0),
            "untagged": counts.get("untagged", 0),
        }

    return {
        "deck": deck_path.name,
        "totals": {
            "total_slides": summary.total_slides,
            "total_shapes": summary.total_shapes,
            "tagged_shapes": summary.tagged_shapes,
            "untagged_shapes": summary.untagged_shapes,
            "custom_xml_parts_found": summary.custom_xml_parts_found,
            "report_configs_resolved": summary.report_configs_resolved,
            "pivot_configs_resolved": summary.pivot_configs_resolved,
            "mapping_configs_resolved": summary.mapping_configs_resolved,
        },
        "per_slide": per_slide,
    }


def main():
    GOLDEN_DIR.mkdir(parents=True, exist_ok=True)

    for key, path in FIXTURE_DECKS.items():
        if not path.exists():
            print(f"[skip] {key}: deck not found at {path}")
            continue
        print(f"[gen]  {key}: reading {path.name}")
        snapshot = build_tag_counts(path)
        out_file = GOLDEN_DIR / f"{key}_tag_counts.json"
        out_file.write_text(json.dumps(snapshot, indent=2), encoding="utf-8")
        print(f"       wrote {out_file.relative_to(REPO_ROOT)}")


if __name__ == "__main__":
    main()
