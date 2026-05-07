"""Build a single 'Testing Deck' by copying all connected slides from
multiple source decks via PowerPoint COM automation.

Approach: COM clipboard-paste preserves connector tags, embedded charts,
embedded data, and all relationships exactly as PowerPoint would when a
human copies a slide. python-pptx alone can't reliably cross-deck-copy
shapes with custom XML parts and embedded OLE objects.

Run:
    python scripts/build_testing_deck.py
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

import pythoncom
import win32com.client as win32

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

from slidegen.deck_reader.tag_reader import generate_config_specs


SOURCES = [
    REPO_ROOT / "output_testing/deck_output/AML & MDS PET Q2FY26 Full Report 27MAR2026 - sandbox migration.pptx",
    REPO_ROOT / "output_testing/deck_output/AVEO Wave 5 PET Report v1.0.pptx",
    REPO_ROOT / "output_testing/deck_output/AZN LOKELMA PET Quarterly Report Q1 2026.pptx",
    REPO_ROOT / "output_testing/deck_output/CREON Share of Voice Study - W33 updated source deck.pptx",
    REPO_ROOT / "output_testing/deck_output/DATROWAY EGFRm NSCLC Promotional Effectiveness Tracking (PET) Q1 '26 PP and NPP Final Report_v1 (4).pptx",
    REPO_ROOT / "output_testing/deck_output/Repatha HCP ATU - Q2'26 Skeleton Deck1 (1).pptx",
    REPO_ROOT / "output_testing/deck_output/[ZoomRx] Abilify LAI PET - Full Report (1).pptx",
    REPO_ROOT / "output_testing/deck_output/[ZoomRx] ILAI Q1 '26 - Topline Report.pptx",
    REPO_ROOT / "projects/J&J Rybrevant PET/Template/ZoomRx_UC_ATU_Report_Q1_'26.pptx",
]

TARGET = REPO_ROOT / "output_testing/deck_output/Testing Deck.pptx"


def connected_slide_indices_one_based(deck_path: Path) -> list[int]:
    """Return 1-based PowerPoint slide indices for slides that have at
    least one connector tag."""
    _, summary = generate_config_specs(str(deck_path))
    return sorted(
        si + 1
        for si, info in summary.per_slide.items()
        if info["tagged"] > 0
    )


def main():
    pythoncom.CoInitialize()
    ppt = win32.Dispatch("PowerPoint.Application")
    # PowerPoint COM requires Visible=True on some Office builds —
    # invisible ops sometimes fail with "PresentationsCount" errors.
    ppt.Visible = True

    # If a previous run left it around, delete first.
    if TARGET.exists():
        TARGET.unlink()

    print(f"Creating target: {TARGET.name}")
    target = ppt.Presentations.Add()
    # PowerPoint requires a Save before Add operations on some builds.
    # 24 = ppSaveAsOpenXMLPresentation (.pptx)
    target.SaveAs(str(TARGET), FileFormat=24)

    grand_total_copied = 0
    per_deck_results = []
    for src_path in SOURCES:
        if not src_path.exists():
            print(f"  [skip] missing: {src_path}")
            continue
        try:
            indices = connected_slide_indices_one_based(src_path)
        except Exception as exc:
            print(f"  [error] {src_path.name}: tag_reader failed — {exc}")
            continue
        print(f"\n--- {src_path.name[:55]} ({len(indices)} connected) ---")
        try:
            src = ppt.Presentations.Open(
                str(src_path), ReadOnly=True, WithWindow=False,
            )
        except Exception as exc:
            print(f"  [error] open failed: {exc}")
            continue
        copied = 0
        for one_based_idx in indices:
            try:
                src.Slides(one_based_idx).Copy()
                # Brief pause for clipboard. PowerPoint sometimes
                # races the clipboard.
                time.sleep(0.05)
                target.Slides.Paste()
                copied += 1
                if copied % 10 == 0:
                    print(f"    {copied}/{len(indices)} copied")
            except Exception as exc:
                print(f"    [warn] slide {one_based_idx} copy failed: {exc}")
                continue
        try:
            src.Close()
        except Exception:
            pass
        print(f"  -> {copied} slides copied")
        per_deck_results.append((src_path.name, copied, len(indices)))
        grand_total_copied += copied
        # Save incrementally so a mid-run crash doesn't lose progress.
        target.Save()

    target.Save()
    target.Close()
    ppt.Quit()

    print()
    print("=" * 60)
    print(f"Testing Deck assembled: {grand_total_copied} slides total")
    print("=" * 60)
    for name, copied, total in per_deck_results:
        print(f"  {name[:55]:55s} {copied}/{total}")
    print(f"\nOutput: {TARGET}")


if __name__ == "__main__":
    main()
