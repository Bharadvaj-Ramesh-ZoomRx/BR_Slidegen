"""
Phase 0: Validate win32com can edit a live PowerPoint from Claude Code.

This is a disposable spike — no pptx_utils, no registry, no skill library.
It validates the single most critical assumption: that Python running from
Claude Code's bash tool can connect to a running PowerPoint instance via COM
and make live edits visible on the analyst's screen.

Usage:
  1. Open ANY presentation in PowerPoint (at least 1 slide with shapes)
  2. Run: python slidegen/phase0_validate_com.py
  3. Watch: changes appear live on the canvas

What this script does:
  - Connects to the running PowerPoint instance
  - Gets the active presentation and first slide
  - Lists all shapes (name, type, position)
  - Modifies one text shape's content and color
  - Moves one shape slightly to the right
  - Reports success/failure
"""

import sys


def main():
    # ── Step 1: Import win32com ──────────────────────────────────────────────
    try:
        import win32com.client
        print("[OK] win32com imported successfully")
    except ImportError:
        print("[FAIL] pywin32 not installed. Run: pip install pywin32")
        sys.exit(1)

    # ── Step 2: Connect to PowerPoint ────────────────────────────────────────
    try:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        print(f"[OK] Connected to PowerPoint (Version: {ppt_app.Version})")
    except Exception as e:
        print(f"[FAIL] Could not connect to PowerPoint: {e}")
        print("       Make sure PowerPoint is running.")
        sys.exit(1)

    # ── Step 3: Get active presentation ──────────────────────────────────────
    if ppt_app.Presentations.Count == 0:
        print("[FAIL] No presentations open. Open a .pptx file in PowerPoint first.")
        sys.exit(1)

    prs = ppt_app.ActivePresentation
    print(f"[OK] Active presentation: {prs.Name}")
    print(f"     Slides: {prs.Slides.Count}")

    # ── Step 4: Get first slide and list shapes ──────────────────────────────
    slide = prs.Slides(1)
    print(f"\n--- Slide 1: {slide.Shapes.Count} shapes ---")

    text_shape = None
    any_shape = None

    for i in range(1, slide.Shapes.Count + 1):
        sh = slide.Shapes(i)
        pos = f"({sh.Left/72:.2f}\", {sh.Top/72:.2f}\")"
        size = f"{sh.Width/72:.2f}\" x {sh.Height/72:.2f}\""
        has_text = sh.HasTextFrame
        text_preview = ""
        if has_text:
            try:
                raw = sh.TextFrame.TextRange.Text
                text_preview = f' text="{raw[:50]}..."' if len(raw) > 50 else f' text="{raw}"'
            except Exception:
                text_preview = " text=(unreadable)"

        print(f"  [{i}] name=\"{sh.Name}\"  type={sh.Type}  pos={pos}  size={size}{text_preview}")

        # Remember first text shape and any shape for editing
        if has_text and text_shape is None:
            try:
                _ = sh.TextFrame.TextRange.Text
                text_shape = sh
            except Exception:
                pass
        if any_shape is None:
            any_shape = sh

    # ── Step 5: Modify a text shape ──────────────────────────────────────────
    if text_shape is not None:
        original_text = text_shape.TextFrame.TextRange.Text
        test_text = "[Phase 0] COM edit works!"
        print(f"\n--- Editing text shape: \"{text_shape.Name}\" ---")
        print(f"  Before: \"{original_text[:60]}\"")

        text_shape.TextFrame.TextRange.Text = test_text
        text_shape.TextFrame.TextRange.Font.Color.RGB = 0x0000FF  # Red in BGR

        print(f"  After:  \"{test_text}\" (red color)")
        print(f"  [OK] Text + color changed. Check your PowerPoint window!")
    else:
        print("\n[SKIP] No text shapes found on slide 1 to edit.")

    # ── Step 6: Move a shape ─────────────────────────────────────────────────
    if any_shape is not None:
        original_left = any_shape.Left
        original_top = any_shape.Top
        new_left = original_left + 36  # Move 0.5 inch right (36 points)

        print(f"\n--- Moving shape: \"{any_shape.Name}\" ---")
        print(f"  Before: ({original_left/72:.2f}\", {original_top/72:.2f}\")")

        any_shape.Left = new_left

        print(f"  After:  ({new_left/72:.2f}\", {original_top/72:.2f}\")")
        print(f"  [OK] Shape moved 0.5\" right. Check your PowerPoint window!")

    # ── Summary ──────────────────────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("Phase 0 COMPLETE")
    print("=" * 60)
    print("If you saw changes appear live in PowerPoint, Phase 0 is validated.")
    print("The win32com + Claude Code architecture works on this machine.")
    print()
    print("Next: Run phase1_create.py to test the full creation-to-editing handoff.")

    # ── Restore original state ───────────────────────────────────────────────
    restore = input("\nRestore original state? [y/N]: ").strip().lower()
    if restore == "y":
        if text_shape is not None:
            text_shape.TextFrame.TextRange.Text = original_text
            text_shape.TextFrame.TextRange.Font.Color.RGB = 0x000000
        if any_shape is not None:
            any_shape.Left = original_left
        print("[OK] Restored.")


if __name__ == "__main__":
    main()
