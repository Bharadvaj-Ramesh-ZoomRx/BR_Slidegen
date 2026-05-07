"""Surgical swap: replace Testing Deck slides 16 (CREON src 22) and 19
(CREON src 61) with CREON sources 25 and 68. Preserves the other 51
slides intact so user only re-migrates the 2 new slides.

Operation order (positions are 1-based):
  1. Delete TD slide 19  (src 61)        -> 52 slides; pos 16 still src 22
  2. Delete TD slide 16  (src 22)        -> 51 slides; CREON block now
                                           has src 28 at pos 16, src 54
                                           at 17, src 71 at 18, etc.
  3. Copy CREON src 25, paste at pos 16  -> 52 slides; src 25 at 16
  4. Copy CREON src 68, paste at pos 19  -> 53 slides; src 68 at 19

Final CREON block (positions 14-23):
  src 12, 16, *25*, 28, 54, *68*, 71, 74, 78, 100
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

import pythoncom
import win32com.client as win32

REPO_ROOT = Path(__file__).resolve().parents[1]
TARGET = REPO_ROOT / "output_testing/deck_output/Testing Deck.pptx"
CREON_SRC = REPO_ROOT / "output_testing/deck_output/CREON Share of Voice Study - W33 updated source deck.pptx"

# 1-based positions in current Testing Deck.pptx (53 slides)
DELETE_POSITIONS = [19, 16]  # delete from highest to lowest
# (creon_source_slide_idx, target_position_after_deletes)
INSERT_PLAN = [
    (25, 16),
    (68, 19),
]


def main():
    pythoncom.CoInitialize()
    ppt = win32.Dispatch("PowerPoint.Application")
    ppt.Visible = True

    print(f"Opening target: {TARGET.name}")
    target = ppt.Presentations.Open(str(TARGET), ReadOnly=False, WithWindow=False)
    print(f"  initial slide count: {target.Slides.Count}")

    # Step 1+2: delete slides at positions 19 then 16
    for pos in DELETE_POSITIONS:
        try:
            target.Slides(pos).Delete()
            print(f"  deleted slide at position {pos} -> {target.Slides.Count} remain")
        except Exception as exc:
            print(f"  [error] delete pos {pos} failed: {exc}")
            target.Close()
            ppt.Quit()
            return

    # Step 3+4: copy from CREON source, paste at the planned position
    print(f"\nOpening CREON source: {CREON_SRC.name[:60]}")
    src = ppt.Presentations.Open(str(CREON_SRC), ReadOnly=True, WithWindow=False)

    for src_idx, dest_pos in INSERT_PLAN:
        try:
            src.Slides(src_idx).Copy()
            time.sleep(0.1)  # clipboard breathing room
            target.Slides.Paste(dest_pos)
            print(f"  copied CREON src slide {src_idx} -> target pos {dest_pos} "
                  f"(target now has {target.Slides.Count} slides)")
        except Exception as exc:
            print(f"  [error] insert src {src_idx} -> pos {dest_pos} failed: {exc}")

    src.Close()
    target.Save()
    target.Close()
    ppt.Quit()
    print(f"\nFinal slide count: {target.Slides.Count if False else '— see Save above'}")
    print(f"Done. Target written: {TARGET}")


if __name__ == "__main__":
    main()
