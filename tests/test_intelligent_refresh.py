"""
test_intelligent_refresh.py — Test harness for intelligent slide refresh.

Imports the core logic from slidegen.intelligent_refresh and provides
a CLI for testing with the Repatha ATU Slide 6 sample deck.

Two-phase pipeline:
  Phase 1 (--read):  Extract slide context + fetch data -> print structured summary
                     Claude Code reads this output and reasons about mappings
  Phase 2 (--refresh mapping.json):  Apply Claude Code's mapping to refresh the deck

Usage:
    # Phase 1: Extract (Claude Code reads the output)
    python tests/test_intelligent_refresh.py --read

    # Phase 2: Refresh (Claude Code provides the mapping)
    python tests/test_intelligent_refresh.py --refresh path/to/mapping.json

    # Or all-in-one with a mapping file:
    python tests/test_intelligent_refresh.py --read --refresh path/to/mapping.json
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

from slidegen.intelligent_refresh import (
    read_slide_context,
    format_slide_for_interpretation,
    format_data_for_interpretation,
    fetch_synapse_data,
    refresh_slide_from_mapping,
)

TESTS_DIR = REPO_ROOT / "tests"
SOURCE_PPTX = TESTS_DIR / "[Vijay] Synapse Connector UAT - Mar 2026.pptx"
OUTPUT_PPTX = TESTS_DIR / "UAT_intelligent_refresh.pptx"

# User-provided data lineage
DATA_LINEAGE = {
    "project_id": 1428,
    "project_name": "Amgen [ATU]: Repatha",
    "reporting_plan_id": 574,
    "analysis_ids": [545991],
    "segment_ids": [14633],
    "dynamic_latest_n": 5,
}


def phase_read(pptx_path: str, data_lineage: dict) -> dict:
    """Extract slide layout + fetch Synapse data. Returns structured context.

    Claude Code reads this output and decides how each chart/table maps to the data.
    """
    context = read_slide_context(pptx_path, slide_index=0)
    records, df = fetch_synapse_data(data_lineage)

    # Build data summary
    data_summary = {
        "record_count": len(records),
        "columns": sorted(df.columns.tolist()) if not df.empty else [],
    }
    if not df.empty:
        for col in df.columns:
            unique = df[col].dropna().unique()
            if len(unique) <= 20:
                data_summary[col] = sorted(str(v) for v in unique)

    return {
        "shapes": context["shapes"],
        "data_summary": data_summary,
        "data_lineage": data_lineage,
    }


def phase_refresh(pptx_path: str, output_path: str, mapping: dict, data_lineage: dict) -> dict:
    """Apply Claude Code's mapping to refresh the deck."""
    return refresh_slide_from_mapping(
        pptx_path, output_path, slide_index=0,
        mapping=mapping, data_lineage=data_lineage,
    )


# ══════════════════════════════════════════════════════════════════════════════
# CLI
# ══════════════════════════════════════════════════════════════════════════════

def main():
    import argparse
    parser = argparse.ArgumentParser(description="Intelligent slide refresh")
    parser.add_argument("--read", action="store_true", help="Phase 1: extract slide context + fetch data")
    parser.add_argument("--refresh", type=str, metavar="MAPPING_JSON", help="Phase 2: refresh with mapping file")
    parser.add_argument("--pptx", type=str, default=str(SOURCE_PPTX), help="Source PPTX path")
    parser.add_argument("--output", type=str, default=str(OUTPUT_PPTX), help="Output PPTX path")
    args = parser.parse_args()

    if args.read:
        print(json.dumps(phase_read(args.pptx, DATA_LINEAGE), indent=2, ensure_ascii=False))

    if args.refresh:
        mapping = json.loads(Path(args.refresh).read_text(encoding="utf-8"))
        results = phase_refresh(args.pptx, args.output, mapping, DATA_LINEAGE)
        print(json.dumps(results, indent=2, ensure_ascii=False))

    if not args.read and not args.refresh:
        parser.print_help()


if __name__ == "__main__":
    main()
