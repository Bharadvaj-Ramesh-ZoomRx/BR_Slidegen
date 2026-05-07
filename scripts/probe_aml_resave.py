"""Probe: round-trip the AML deck via python-pptx and verify
(a) connector tags survive the resave and (b) PowerPoint COM can open
the resaved file. If both pass, we can pre-bake a sanitized copy of
AML before build_testing_deck.py runs."""
import os
import sys
from pathlib import Path

import pythoncom
import win32com.client as w
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from slidegen.deck_reader.tag_reader import generate_config_specs

SRC = Path("output_testing/deck_output/AML & MDS PET Q2FY26 Full Report 27MAR2026 - sandbox migration.pptx").resolve()
OUT = Path(r"C:\Users\BHARAD~1\AppData\Local\Temp\pptest\aml_resaved.pptx")
OUT.parent.mkdir(parents=True, exist_ok=True)

p = Presentation(str(SRC))
print(f"Pre-resave: {len(p.slides)} slides")
p.save(str(OUT))
print(f"Saved to {OUT}")

specs, summary = generate_config_specs(str(OUT))
n_tagged = sum(1 for info in summary.per_slide.values() if info["tagged"] > 0)
print(f"Post-resave tag_reader: {len(specs)} specs, {n_tagged} tagged slides")

pythoncom.CoInitialize()
ppt = w.Dispatch("PowerPoint.Application")
ppt.Visible = True
print("trying COM open:", OUT)
try:
    pp = ppt.Presentations.Open(str(OUT), ReadOnly=True, WithWindow=False)
    print(f"COM OPENED — {pp.Slides.Count} slides")
    pp.Close()
except Exception as e:
    print(f"COM OPEN FAILED: {e}")
ppt.Quit()
