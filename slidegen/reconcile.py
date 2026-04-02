"""
Registry reconciliation: read live PowerPoint state and sync slide_registry.json.

Must run before every COM edit session. Ensures the registry reflects the
actual current state of the slide, absorbing any manual edits the analyst
made in PowerPoint.

Usage as module:
    from slidegen.reconcile import reconcile
    report = reconcile("my_deck.pptx", slide_num=1)

Usage as CLI:
    python -m slidegen.reconcile my_deck.pptx [slide_num]
"""

import os
import sys
import json
from datetime import datetime
from collections import defaultdict

from slidegen.config import REGISTRY_PATH, SHAPE_PREFIX, PTS_PER_INCH


def _load_registry(path=None):
    with open(path or REGISTRY_PATH) as f:
        return json.load(f)


def _save_registry(registry, path=None):
    with open(path or REGISTRY_PATH, "w") as f:
        json.dump(registry, f, indent=2)


def reconcile(target_filename, slide_num=1, registry_path=None):
    """Run full reconciliation against a live PowerPoint instance.

    Args:
        target_filename: Bare filename matched against open presentations.
        slide_num: 1-indexed slide number (default 1).
        registry_path: Path to shape_registry.json. If None, uses default
            REGISTRY_PATH from config. For pipeline projects, pass the
            per-wave registry at output/{wave}/shape_registry.json.

    Returns:
        dict with keys:
            found (int): shapes matched in registry and on slide
            missing (list[str]): registered shapes not found on slide
            unregistered (list[str]): zrx_* shapes on slide not in registry
            duplicates (list[str]): zrx_* names with >1 shape (blocks editing)
            changes (list[dict]): position/size deltas detected
            ok (bool): True if no duplicates and reconciliation succeeded
    """
    import win32com.client
    IN = PTS_PER_INCH
    reg_path = registry_path or REGISTRY_PATH

    # ── Load registry ────────────────────────────────────────────────────────
    try:
        registry = _load_registry(reg_path)
    except FileNotFoundError:
        raise FileNotFoundError(
            f"No registry at {reg_path}. "
            "Create a slide first with slidegen.create or generate_deck()."
        )

    # ── Connect to PowerPoint ────────────────────────────────────────────────
    try:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    except Exception as e:
        raise RuntimeError(f"Cannot connect to PowerPoint: {e}")

    try:
        prs = None
        for i in range(1, ppt_app.Presentations.Count + 1):
            p = ppt_app.Presentations(i)
            if target_filename in p.Name:
                prs = p
                break

        if prs is None:
            open_files = [
                ppt_app.Presentations(i).Name
                for i in range(1, ppt_app.Presentations.Count + 1)
            ]
            raise RuntimeError(
                f"'{target_filename}' not found. Open files: {open_files}"
            )

        slide = prs.Slides(slide_num)

        # ── Build shape inventory ────────────────────────────────────────────
        zrx_shapes = defaultdict(list)
        for i in range(1, slide.Shapes.Count + 1):
            sh = slide.Shapes(i)
            if sh.Name.startswith(SHAPE_PREFIX):
                zrx_shapes[sh.Name].append(sh)

        # ── Duplicate detection (hard stop) ──────────────────────────────────
        duplicates = [name for name, refs in zrx_shapes.items() if len(refs) > 1]
        if duplicates:
            return {
                "found": 0,
                "missing": [],
                "unregistered": [],
                "duplicates": duplicates,
                "changes": [],
                "ok": False,
            }

        # ── Reconcile each registered shape ──────────────────────────────────
        reg_shapes = registry.get("shapes", {})
        found = []
        missing = []
        changes = []

        for name, record in reg_shapes.items():
            if name in zrx_shapes:
                sh = zrx_shapes[name][0]
                found.append(name)

                live = {
                    "left":   round(sh.Left / IN, 4),
                    "top":    round(sh.Top / IN, 4),
                    "width":  round(sh.Width / IN, 4),
                    "height": round(sh.Height / IN, 4),
                }

                try:
                    live["text"] = sh.TextFrame.TextRange.Text[:100]
                except AttributeError:
                    pass

                for field in ("left", "top", "width", "height"):
                    old_val = record.get(field)
                    new_val = live[field]
                    if old_val is not None and abs(old_val - new_val) > 0.01:
                        changes.append({
                            "shape": name,
                            "field": field,
                            "before": old_val,
                            "after": new_val,
                        })

                record.update(live)
                record["last_reconciled"] = datetime.now().isoformat(timespec="seconds")
            else:
                missing.append(name)

        # ── Detect unregistered shapes ───────────────────────────────────────
        unregistered = sorted(set(zrx_shapes.keys()) - set(reg_shapes.keys()))

        # ── Update and save ──────────────────────────────────────────────────
        registry.setdefault("meta", {})
        registry["meta"]["last_reconciled"] = datetime.now().isoformat(timespec="seconds")
        registry["meta"]["reconcile_source"] = prs.Name
        _save_registry(registry, reg_path)

        return {
            "found": len(found),
            "missing": missing,
            "unregistered": unregistered,
            "duplicates": [],
            "changes": changes,
            "ok": True,
        }
    finally:
        # Release COM references to prevent orphaned PowerPoint processes
        ppt_app = None


# ── CLI ──────────────────────────────────────────────────────────────────────

def _cli():
    if len(sys.argv) < 2:
        print("Usage: python -m slidegen.reconcile <filename.pptx> [slide_num]")
        sys.exit(1)

    filename = sys.argv[1]
    slide_num = int(sys.argv[2]) if len(sys.argv) > 2 else 1

    try:
        report = reconcile(filename, slide_num)
    except (FileNotFoundError, RuntimeError) as e:
        print(f"[FAIL] {e}")
        sys.exit(1)

    print(f"{'=' * 60}")
    print(f"RECONCILIATION REPORT -- Slide {slide_num}")
    print(f"{'=' * 60}")

    if report["duplicates"]:
        print(f"\n[HARD STOP] Duplicate zrx_* names: {report['duplicates']}")
        print("Fix: rename duplicates in PowerPoint, then re-run.")
        sys.exit(1)

    print(f"  Found: {report['found']}")

    if report["missing"]:
        print(f"\n  [WARN] Missing from slide:")
        for name in report["missing"]:
            print(f"    {name} (deleted manually?)")

    if report["unregistered"]:
        print(f"\n  [WARN] Unregistered zrx_* shapes:")
        for name in report["unregistered"]:
            print(f"    {name} (copy-paste or manual creation?)")

    if report["changes"]:
        print(f"\n  Changes detected ({len(report['changes'])}):")
        for ch in report["changes"]:
            print(f"    {ch['shape']}.{ch['field']}: "
                  f"{ch['before']:.2f}\" -> {ch['after']:.2f}\"")
    else:
        print(f"\n  No changes detected.")

    print(f"\n  Registry updated: {REGISTRY_PATH}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    _cli()
