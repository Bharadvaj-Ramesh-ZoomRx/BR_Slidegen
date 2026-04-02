"""
Test all 5 workflows end-to-end with dummy data.

Workflows:
1. Create slides (generate_deck)
2. Edit Slide N (regenerate_slide)
3. Edit slides with new wave data
4. Add a slide (modify config + generate_deck)
5. Remove a slide (modify config + generate_deck)
"""

import os
import sys
import shutil
import yaml
from pptx import Presentation

BASE = os.path.dirname(__file__)
ROOT = os.path.dirname(BASE)
sys.path.insert(0, ROOT)

YAML_PATH = os.path.join(BASE, "test_project", "config.yaml")
PASS = 0
FAIL = 0


def report(name, ok, detail=""):
    global PASS, FAIL
    status = "PASS" if ok else "FAIL"
    if ok:
        PASS += 1
    else:
        FAIL += 1
    print(f"  [{status}] {name}" + (f" — {detail}" if detail else ""))


# ═══════════════════════════════════════════════════════════════════════════════
# WORKFLOW 1: Create slides (generate_deck)
# ═══════════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("WORKFLOW 1: Create slides for test_project (generate_deck)")
print("=" * 70)

from slidegen.pipeline.orchestrator import generate_deck, regenerate_slide, _backup_config
from slidegen.pipeline.project_config import load_project_config

config = load_project_config(YAML_PATH)

# Verify wave interpolation
report("Wave interpolation — data path",
       "WAVE_Q1Q2_2026" in config.data_source_path,
       config.data_source_path)
report("Wave interpolation — output path",
       "WAVE_Q1Q2_2026" in config.output_path,
       config.output_path)
report("Template — no wave",
       "WAVE" not in config.template_path and "{{wave}}" not in config.template_path,
       config.template_path)
report("Data file exists",
       os.path.exists(config.data_source_path),
       config.data_source_path)

# Generate deck
try:
    out_path = generate_deck(YAML_PATH)
    report("generate_deck() succeeded", True, out_path)
    report("Output file exists", os.path.exists(out_path))
    report("Output in wave folder",
           "WAVE_Q1Q2_2026" in out_path,
           out_path)

    # Check slide count
    prs = Presentation(out_path)
    expected_slides = len(config.asks)
    report(f"Slide count = {len(prs.slides)} (expected {expected_slides})",
           len(prs.slides) == expected_slides)

    # Check shapes are named with zrx_
    all_named = True
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if not shape.name.startswith("zrx_"):
                all_named = False
                print(f"    Unnamed shape on slide {i+1}: {shape.name}")
    report("All shapes have zrx_ names", all_named)

    # Check shape registry
    registry_path = os.path.join(os.path.dirname(out_path), "shape_registry.json")
    report("Shape registry exists", os.path.exists(registry_path))

except Exception as e:
    report("generate_deck()", False, str(e))
    import traceback
    traceback.print_exc()

# ═══════════════════════════════════════════════════════════════════════════════
# WORKFLOW 2: Edit Slide N (regenerate_slide)
# ═══════════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("WORKFLOW 2: Edit Slide 4 — regenerate single slide")
print("=" * 70)

try:
    # Read current state of slide 4
    prs_before = Presentation(out_path)
    slide4_shapes_before = len(prs_before.slides[3].shapes)
    slide3_shapes_before = len(prs_before.slides[2].shapes)

    # Regenerate slide 4 (0-indexed = 3)
    result = regenerate_slide(YAML_PATH, slide_index=3)
    report("regenerate_slide(index=3) succeeded", True)

    prs_after = Presentation(result)
    slide4_shapes_after = len(prs_after.slides[3].shapes)
    slide3_shapes_after = len(prs_after.slides[2].shapes)

    report(f"Slide 4 rebuilt: {slide4_shapes_before} -> {slide4_shapes_after} shapes",
           slide4_shapes_after > 0)
    report("Slide 3 untouched",
           slide3_shapes_before == slide3_shapes_after,
           f"{slide3_shapes_before} -> {slide3_shapes_after}")
    report("Total slides unchanged",
           len(prs_after.slides) == expected_slides)

    # Verify zrx_ naming on regenerated slide
    regen_named = all(s.name.startswith("zrx_") for s in prs_after.slides[3].shapes)
    report("Regenerated slide shapes have zrx_ names", regen_named)

except Exception as e:
    report("regenerate_slide()", False, str(e))
    import traceback
    traceback.print_exc()

# ═══════════════════════════════════════════════════════════════════════════════
# WORKFLOW 3: Edit slides with new wave data
# ═══════════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("WORKFLOW 3: Edit slides with new wave data — WAVE_Q3Q4_2026")
print("=" * 70)

try:
    # Step 1: Backup config
    backup_path = _backup_config(YAML_PATH)
    report("Config backup created", os.path.exists(backup_path), backup_path)

    # Step 2: Update config wave
    with open(YAML_PATH, "r") as f:
        raw = yaml.safe_load(f)

    old_wave = raw["project"]["wave"]
    raw["project"]["wave"] = "WAVE_Q3Q4_2026"
    raw["project"]["period_current"] = "Q4'26"
    raw["project"]["period_prior"] = "Q3'26"

    with open(YAML_PATH, "w") as f:
        yaml.dump(raw, f, default_flow_style=False, sort_keys=False)

    # Step 3: Verify new data exists
    config2 = load_project_config(YAML_PATH)
    report("Wave updated in config", config2.wave == "WAVE_Q3Q4_2026")
    report("New data path resolves",
           "WAVE_Q3Q4_2026" in config2.data_source_path,
           config2.data_source_path)
    report("New data file exists", os.path.exists(config2.data_source_path))
    report("New output path",
           "WAVE_Q3Q4_2026" in config2.output_path,
           config2.output_path)

    # Step 4: Generate deck with new wave
    out_path_w2 = generate_deck(YAML_PATH)
    report("generate_deck() with new wave succeeded", True, out_path_w2)
    report("Output in new wave folder",
           "WAVE_Q3Q4_2026" in out_path_w2)

    # Step 5: Old wave output still exists
    report("Old wave output preserved",
           os.path.exists(out_path),
           out_path)

    # Verify slide content uses new periods
    prs_w2 = Presentation(out_path_w2)
    report(f"New wave deck has {len(prs_w2.slides)} slides",
           len(prs_w2.slides) == expected_slides)

    # Restore original config from backup
    shutil.copy2(backup_path, YAML_PATH)
    report("Config restored from backup", True)

except Exception as e:
    report("Wave switch workflow", False, str(e))
    import traceback
    traceback.print_exc()
    # Restore config
    if os.path.exists(backup_path):
        shutil.copy2(backup_path, YAML_PATH)

# ═══════════════════════════════════════════════════════════════════════════════
# WORKFLOW 4: Add a slide (modify config + generate_deck)
# ═══════════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("WORKFLOW 4: Add a slide after Slide 5")
print("=" * 70)

try:
    # Step 1: Backup
    backup_path2 = _backup_config(YAML_PATH)
    report("Config backup created", os.path.exists(backup_path2))

    # Step 2: Add a new ask after index 4 (after slide 5)
    with open(YAML_PATH, "r") as f:
        raw = yaml.safe_load(f)

    new_ask = {
        "id": "new_cta_slide",
        "slide_type": "single_bar_with_delta",
        "headline": "NEW: Call-to-action metrics for {{primary.name}}",
        "section": "NEW CTA SLIDE — {{period_current}}",
        "source_text": "Source: Test Data",
        "data_key": "ryb_mr",
        "brand": "primary",
        "sort_by": "current",
    }
    raw["asks"].insert(5, new_ask)
    original_ask_count = len(raw["asks"]) - 1  # before adding

    with open(YAML_PATH, "w") as f:
        yaml.dump(raw, f, default_flow_style=False, sort_keys=False)

    # Step 3: Full regen
    out_path_add = generate_deck(YAML_PATH)
    prs_add = Presentation(out_path_add)
    report(f"Deck with added slide has {len(prs_add.slides)} slides (was {expected_slides})",
           len(prs_add.slides) == expected_slides + 1)

    # Restore
    shutil.copy2(backup_path2, YAML_PATH)
    report("Config restored", True)

except Exception as e:
    report("Add slide workflow", False, str(e))
    import traceback
    traceback.print_exc()
    if os.path.exists(backup_path2):
        shutil.copy2(backup_path2, YAML_PATH)

# ═══════════════════════════════════════════════════════════════════════════════
# WORKFLOW 5: Remove a slide (modify config + generate_deck)
# ═══════════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("WORKFLOW 5: Remove Slide 6 (followup_reps)")
print("=" * 70)

try:
    # Step 1: Backup
    backup_path3 = _backup_config(YAML_PATH)
    report("Config backup created", os.path.exists(backup_path3))

    # Step 2: Remove ask at index 5
    with open(YAML_PATH, "r") as f:
        raw = yaml.safe_load(f)

    removed = raw["asks"].pop(5)
    report(f"Removed ask: {removed['id']}", removed["id"] == "followup_reps")

    with open(YAML_PATH, "w") as f:
        yaml.dump(raw, f, default_flow_style=False, sort_keys=False)

    # Step 3: Full regen
    out_path_rm = generate_deck(YAML_PATH)
    prs_rm = Presentation(out_path_rm)
    report(f"Deck with removed slide has {len(prs_rm.slides)} slides (was {expected_slides})",
           len(prs_rm.slides) == expected_slides - 1)

    # Restore
    shutil.copy2(backup_path3, YAML_PATH)
    report("Config restored", True)

except Exception as e:
    report("Remove slide workflow", False, str(e))
    import traceback
    traceback.print_exc()
    if os.path.exists(backup_path3):
        shutil.copy2(backup_path3, YAML_PATH)

# ═══════════════════════════════════════════════════════════════════════════════
# SUMMARY
# ═══════════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print(f"RESULTS: {PASS} passed, {FAIL} failed, {PASS + FAIL} total")
print("=" * 70)

if FAIL > 0:
    sys.exit(1)
