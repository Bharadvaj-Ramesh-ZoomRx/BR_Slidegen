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

import json
import sys
import time
from pathlib import Path

import pythoncom
import win32com.client as win32
from pptx import Presentation

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

from slidegen.deck_reader.tag_reader import generate_config_specs
from slidegen.slide_spec.schema import dump_spec


# (label, path, quota) per source deck.
# Quota = number of slides to pick from that deck.
# Selection rule: spread evenly across the deck's data-slide pool.
# CREON gets 10 (better connection quality); other 8 decks get 5 each → 50 total.
SOURCES = [
    ("AML & MDS",  REPO_ROOT / "output_testing/deck_output/AML & MDS PET Q2FY26 Full Report 27MAR2026 - sandbox migration.pptx", 5),
    ("AVEO",       REPO_ROOT / "output_testing/deck_output/AVEO Wave 5 PET Report v1.0.pptx", 5),
    ("AZN",        REPO_ROOT / "output_testing/deck_output/AZN LOKELMA PET Quarterly Report Q1 2026.pptx", 5),
    ("CREON",      REPO_ROOT / "output_testing/deck_output/CREON Share of Voice Study - W33 updated source deck.pptx", 10),
    ("DATROWAY",   REPO_ROOT / "output_testing/deck_output/DATROWAY EGFRm NSCLC Promotional Effectiveness Tracking (PET) Q1 '26 PP and NPP Final Report_v1 (4).pptx", 5),
    ("Repatha",    REPO_ROOT / "output_testing/deck_output/Repatha HCP ATU - Q2'26 Skeleton Deck1 (1).pptx", 5),
    ("Abilify",    REPO_ROOT / "output_testing/deck_output/[ZoomRx] Abilify LAI PET - Full Report (1).pptx", 5),
    ("ILAI",       REPO_ROOT / "output_testing/deck_output/[ZoomRx] ILAI Q1 '26 - Topline Report.pptx", 5),
    ("UC ATU",     REPO_ROOT / "projects/J&J Rybrevant PET/Template/ZoomRx_UC_ATU_Report_Q1_'26.pptx", 5),
]

TARGET = REPO_ROOT / "output_testing/deck_output/Testing Deck.pptx"


def hidden_slide_indices_one_based(deck_path: Path) -> set[int]:
    """Return 1-based indices of slides marked hidden in PowerPoint
    (<p:sld show="0">). Hidden slides are usually backup/cut content
    and shouldn't appear in the Testing Deck."""
    prs = Presentation(str(deck_path))
    return {
        i + 1 for i, s in enumerate(prs.slides)
        if s.element.get("show") == "0"
    }


def data_slide_pool_one_based(deck_path: Path) -> list[int]:
    """Return 1-based PowerPoint slide indices for DATA SLIDES — slides
    that are fully dynamic (no static-pinned components), have at least
    one dynamic chart or value_table, AND are not hidden. Skips ES /
    text / label-only / hidden slides so every picked slide exercises
    real refresh."""
    hidden = hidden_slide_indices_one_based(deck_path)
    specs, _ = generate_config_specs(str(deck_path))
    indices: list[int] = []
    for spec in specs:
        d = json.loads(dump_spec(spec))
        si_one = d["slide_index"] + 1  # 1-based
        if si_one in hidden:
            continue
        n_dyn_data = 0  # dynamic chart / value_table count
        n_static = 0
        for comp in d.get("components", []):
            t = comp.get("type")
            if t not in ("chart", "value_table", "label_table"):
                continue
            lin = (comp.get("data_mapping") or {}).get("raw_data_lineage") or {}
            if not lin:
                continue
            is_dyn = (lin.get("dynamic_latest_n") or 0) > 0 or bool(lin.get("include_live_wave"))
            if is_dyn:
                if t in ("chart", "value_table"):
                    n_dyn_data += 1
            else:
                if lin.get("static_time_period_ids"):
                    n_static += 1
        if n_dyn_data > 0 and n_static == 0:
            indices.append(si_one)
    return sorted(indices)


def add_section_divider(target, label: str, n_slides: int, total_decks: int) -> None:
    """Append a blank section-divider slide to `target` with project name
    and slide count. Used to separate the per-deck batches so a reviewer
    can see at a glance which slides came from which source."""
    # ppLayoutBlank = 12
    insert_at = target.Slides.Count + 1
    slide = target.Slides.Add(insert_at, 12)
    # Title (project label) — centered horizontally on a 16:9 slide (960pt wide)
    title_tb = slide.Shapes.AddTextbox(1, 80, 200, 800, 80)
    tr = title_tb.TextFrame.TextRange
    tr.Text = label
    tr.Font.Size = 44
    tr.Font.Bold = True
    tr.ParagraphFormat.Alignment = 2  # ppAlignCenter
    # Subtitle: e.g. "5 slides · 1 of 9 source decks"
    sub_tb = slide.Shapes.AddTextbox(1, 80, 300, 800, 50)
    sr = sub_tb.TextFrame.TextRange
    sr.Text = f"{n_slides} slides from this project"
    sr.Font.Size = 22
    sr.Font.Italic = True
    sr.ParagraphFormat.Alignment = 2


def spread_pick(pool: list[int], n: int) -> list[int]:
    """Return n indices spread evenly across the pool (positions at
    (i+0.5)/n). Returns the whole pool if n >= len(pool)."""
    if n >= len(pool):
        return list(pool)
    return [pool[round((i + 0.5) * len(pool) / n - 0.5)] for i in range(n)]


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
    # Force 16:9 — matches the source decks and keeps pasted slides
    # from being rescaled. ppSlideSizeOnScreen16x9 = 15.
    try:
        target.PageSetup.SlideSize = 15
    except Exception as exc:
        print(f"  [warn] could not set 16:9 page size: {exc}")
    # PowerPoint requires a Save before Add operations on some builds.
    # 24 = ppSaveAsOpenXMLPresentation (.pptx)
    target.SaveAs(str(TARGET), FileFormat=24)

    grand_total_copied = 0
    per_deck_results = []
    n_decks = sum(1 for _, p, _ in SOURCES if p.exists())
    for label, src_path, quota in SOURCES:
        if not src_path.exists():
            print(f"  [skip] missing: {src_path}")
            continue
        try:
            pool = data_slide_pool_one_based(src_path)
        except Exception as exc:
            print(f"  [error] {src_path.name}: tag_reader failed — {exc}")
            continue
        indices = spread_pick(pool, quota)
        print(f"\n--- {label} ({src_path.name[:50]}) — pool={len(pool)} pick={len(indices)} ---")
        print(f"    picks: {indices}")
        try:
            src = ppt.Presentations.Open(
                str(src_path), ReadOnly=True, WithWindow=False,
            )
        except Exception as exc:
            print(f"  [error] open failed: {exc}")
            continue
        # Section divider before this deck's batch
        try:
            add_section_divider(target, label, len(indices), n_decks)
        except Exception as exc:
            print(f"    [warn] divider insert failed: {exc}")
        copied = 0
        for one_based_idx in indices:
            try:
                src.Slides(one_based_idx).Copy()
                # Brief pause for clipboard. PowerPoint sometimes
                # races the clipboard.
                time.sleep(0.05)
                target.Slides.Paste()
                copied += 1
            except Exception as exc:
                print(f"    [warn] slide {one_based_idx} copy failed: {exc}")
                continue
        try:
            src.Close()
        except Exception:
            pass
        print(f"  -> {copied} slides copied")
        per_deck_results.append((label, copied, len(indices)))
        grand_total_copied += copied
        # Save incrementally so a mid-run crash doesn't lose progress.
        target.Save()

    target.Save()
    target.Close()
    ppt.Quit()

    print()
    print("=" * 60)
    print(f"Testing Deck assembled: {grand_total_copied} content slides "
          f"+ {len(per_deck_results)} dividers = "
          f"{grand_total_copied + len(per_deck_results)} total")
    print("=" * 60)
    for name, copied, total in per_deck_results:
        print(f"  {name:<12s} {copied}/{total}")
    print(f"\nOutput: {TARGET}")


if __name__ == "__main__":
    main()
