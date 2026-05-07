"""Assemble a Headliner test deck from 10 slides across 8 source decks.

Picks one or two slides per deck — all of them have:
  - a talking header (top-left long sentence)
  - at least 1 chart AND 1 table on the slide (likely mix of connected
    + non-connected components)

The user can then open the assembled deck, delete the 10 talking
headers, save, and re-run the headliner via:

    python -m slidegen.headliner_full_workflow \
        output_testing/deck_output/Headliner_Test_Deck.pptx

Uses PowerPoint COM (pywin32) — needed because python-pptx doesn't
copy slides between presentations cleanly (theme/layout/relationships).
"""
from pathlib import Path
import time

import win32com.client


REPO = Path(__file__).resolve().parents[1]
DECK_DIR = REPO / "output_testing" / "deck_output"
OUT_DECK = DECK_DIR / "Headliner_Test_Deck.pptx"

# (source-deck filename, 1-based slide number, label)
# (AML & MDS deck dropped — PowerPoint COM rejects it as Protected
# View / corrupt; replaced with CREON slide 26 for trended-chart diversity.)
PICKS = [
    ("AVEO Wave 5 PET Report v1.0.pptx",
     24, "AVEO — FOTIVDA action requests"),
    ("AZN LOKELMA PET Quarterly Report Q1 2026.pptx",
     8,  "LOKELMA — Nephs/PCPs reach"),
    ("CREON Share of Voice Study - W33 updated source deck.pptx",
     25, "CREON — OSQ rose Mar'26"),
    ("CREON Share of Voice Study - W33 updated source deck.pptx",
     10, "CREON — NP/PAs attribute decline"),
    ("CREON Share of Voice Study - W33 updated source deck.pptx",
     26, "CREON — rep closing trended"),
    ("DATROWAY EGFRm NSCLC Promotional Effectiveness Tracking (PET) Q1 '26 PP and NPP Final Report_v1 (4).pptx",
     36, "DATROWAY — NPP effectiveness Q4'25"),
    ("Repatha HCP ATU - Q2'26 Skeleton Deck1 (1).pptx",
     11, "Repatha — CARDs LDL-C targets"),
    ("Repatha HCP ATU - Q2'26 Skeleton Deck1 (1).pptx",
     81, "Repatha — patient-requested non-statin"),
    ("[ZoomRx] Abilify LAI PET - Full Report (1).pptx",
     15, "Abilify LAI — Interaction Details"),
    ("[ZoomRx] ILAI Q1 '26 - Topline Report.pptx",
     17, "ILAI — Efficacy vs Orals messages"),
]


def main():
    if OUT_DECK.exists():
        OUT_DECK.unlink()
    print(f"Assembling {len(PICKS)} slides into {OUT_DECK.name} ...")

    pp = win32com.client.Dispatch("PowerPoint.Application")
    pp.Visible = 1  # required by COM API
    # Build a fresh empty deck
    out_pres = pp.Presentations.Add()

    # Track open source presentations to close at the end
    open_sources = {}

    try:
        for i, (deck_name, slide_no, label) in enumerate(PICKS, start=1):
            src_path = DECK_DIR / deck_name
            if not src_path.exists():
                print(f"  SKIP missing: {deck_name}")
                continue
            if str(src_path) not in open_sources:
                # ReadOnly=True to avoid lock contention
                src_pres = pp.Presentations.Open(
                    str(src_path), ReadOnly=True, WithWindow=False)
                open_sources[str(src_path)] = src_pres
            src_pres = open_sources[str(src_path)]
            if slide_no > src_pres.Slides.Count:
                print(f"  SKIP {deck_name} slide {slide_no} — only "
                      f"{src_pres.Slides.Count} slides")
                continue
            src_slide = src_pres.Slides(slide_no)
            # Copy + paste at end of out_pres
            src_slide.Copy()
            time.sleep(0.3)  # let clipboard settle
            insert_at = out_pres.Slides.Count + 1
            out_pres.Slides.Paste(insert_at)
            print(f"  {i:2d}. {label}  (from {deck_name[:40]}..., slide {slide_no})")

        # First slide of out_pres is the default blank — remove it
        if out_pres.Slides.Count > len(PICKS):
            out_pres.Slides(1).Delete()

        out_pres.SaveAs(str(OUT_DECK))
        print(f"\nWrote: {OUT_DECK}")
    finally:
        for src in open_sources.values():
            try:
                src.Close()
            except Exception:
                pass
        try:
            out_pres.Close()
        except Exception:
            pass
        try:
            pp.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
