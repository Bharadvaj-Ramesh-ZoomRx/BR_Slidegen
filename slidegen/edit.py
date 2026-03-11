"""
Live editing engine via win32com COM automation.

Connects to a running PowerPoint instance and makes surgical edits to
shapes identified by their zrx_ names. All changes appear live on the
analyst's canvas without saving/reopening.

Usage as module:
    from slidegen.edit import LiveEditor
    with LiveEditor("my_deck.pptx") as editor:
        editor.set_text("zrx_003", "New headline", color="red", bold=True)
        editor.move("zrx_004", left=9.5)
        editor.set_fill("zrx_005", "#00B050")

Usage as CLI (interactive):
    python -m slidegen.edit my_deck.pptx
"""

import json
import os
import sys
from datetime import datetime

from slidegen.config import REGISTRY_PATH, EDIT_LOG_PATH, PTS_PER_INCH


# ── BGR color lookup ─────────────────────────────────────────────────────────
_COLOR_MAP = {
    "red":    0x0000FF,
    "green":  0x50B000,
    "blue":   0xFF0000,
    "orange": 0x2458F7,
    "grey":   0x505050,
    "gray":   0x505050,
    "white":  0xFFFFFF,
    "black":  0x000000,
    "purple": 0xA03070,
}


def _resolve_color(color):
    """Convert a color name, hex string, or int to BGR int for COM."""
    if color is None:
        return None
    if isinstance(color, int):
        return color
    if isinstance(color, str):
        if color.lower() in _COLOR_MAP:
            return _COLOR_MAP[color.lower()]
        # Hex string like "#FF0000" -> BGR
        c = color.lstrip("#")
        if len(c) == 6:
            r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
            return b << 16 | g << 8 | r
    raise ValueError(f"Unknown color: {color}")


class LiveEditor:
    """COM-based live editor for an open PowerPoint presentation."""

    def __init__(self, filename, slide_num=1, registry_path=None, log_path=None):
        """Connect to a running PowerPoint and locate the target file.

        Args:
            filename: Bare filename (e.g. 'demo_slide.pptx'), matched
                      against open presentations.
            slide_num: 1-indexed slide number to edit (default 1).
            registry_path: Path to shape_registry.json. If None, uses
                default REGISTRY_PATH. For pipeline projects, pass the
                per-wave path at output/{wave}/shape_registry.json.
            log_path: Path to edit_log.json. If None, uses default
                EDIT_LOG_PATH. For pipeline projects, pass the per-wave
                path at output/{wave}/edit_log.json.
        """
        self.filename = filename
        self.slide_num = slide_num
        self._registry_path = registry_path or REGISTRY_PATH
        self._log_path = log_path or EDIT_LOG_PATH
        self._ppt_app = None
        self._prs = None
        self._slide = None
        self._edits = []

    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, *exc):
        self._save_edit_log()

    def connect(self):
        """Establish COM connection to PowerPoint."""
        import win32com.client
        try:
            self._ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        except Exception as e:
            raise RuntimeError(
                f"Cannot connect to PowerPoint: {e}\n"
                "Make sure PowerPoint is running."
            )

        # Match by exact filename first, then fall back to substring match
        open_files = []
        substring_match = None
        for i in range(1, self._ppt_app.Presentations.Count + 1):
            p = self._ppt_app.Presentations(i)
            open_files.append(p.Name)
            if p.Name == self.filename:
                self._prs = p
                break
            if substring_match is None and self.filename in p.Name:
                substring_match = p

        if self._prs is None and substring_match is not None:
            self._prs = substring_match

        if self._prs is None:
            raise RuntimeError(
                f"'{self.filename}' not found in open presentations.\n"
                f"Open files: {open_files}"
            )

        self._slide = self._prs.Slides(self.slide_num)

    def _find(self, name):
        """Find a shape by zrx_ name on the current slide."""
        for i in range(1, self._slide.Shapes.Count + 1):
            sh = self._slide.Shapes(i)
            if sh.Name == name:
                return sh
        raise RuntimeError(
            f"Shape '{name}' not found on slide {self.slide_num}. "
            "Run reconcile first."
        )

    def _log(self, shape_name, action, changes):
        """Record an edit for the edit log."""
        self._edits.append({
            "id": f"edit_{len(self._edits) + 1:03d}",
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "shape": shape_name,
            "action": action,
            "changes": changes,
        })

    def _save_edit_log(self):
        """Append session edits to the persistent edit log."""
        if not self._edits:
            return
        log_path = self._log_path
        os.makedirs(os.path.dirname(log_path), exist_ok=True)
        existing = []
        if os.path.exists(log_path):
            with open(log_path) as f:
                existing = json.load(f).get("edits", [])
        existing.extend(self._edits)
        with open(log_path, "w") as f:
            json.dump({"edits": existing}, f, indent=2)

    # ── Edit operations ──────────────────────────────────────────────────────

    def list_shapes(self):
        """List all shapes on the current slide.

        Returns:
            list of dicts with keys: name, type, left, top, width, height, text
        """
        IN = PTS_PER_INCH
        result = []
        for i in range(1, self._slide.Shapes.Count + 1):
            sh = self._slide.Shapes(i)
            info = {
                "name": sh.Name,
                "type": sh.Type,
                "left": round(sh.Left / IN, 3),
                "top": round(sh.Top / IN, 3),
                "width": round(sh.Width / IN, 3),
                "height": round(sh.Height / IN, 3),
            }
            try:
                info["text"] = sh.TextFrame.TextRange.Text[:80]
            except Exception:
                info["text"] = None
            result.append(info)
        return result

    def set_text(self, name, text, color=None, size_pt=None, bold=None):
        """Change text content and optional formatting.

        Args:
            name: zrx_ shape name
            text: new text string
            color: color name, hex, or BGR int (optional)
            size_pt: font size in points (optional)
            bold: True/False (optional)
        """
        sh = self._find(name)
        before = sh.TextFrame.TextRange.Text
        tr = sh.TextFrame.TextRange
        tr.Text = text
        changes = {"text": {"before": before[:50], "after": text[:50]}}

        bgr = _resolve_color(color)
        if bgr is not None:
            tr.Font.Color.RGB = bgr
            changes["color"] = str(color)
        if size_pt is not None:
            tr.Font.Size = size_pt
            changes["size_pt"] = size_pt
        if bold is not None:
            tr.Font.Bold = bold
            changes["bold"] = bold

        self._log(name, "set_text", changes)

    def set_fill(self, name, color):
        """Change fill color of a shape.

        Args:
            name: zrx_ shape name
            color: color name, hex string, or BGR int
        """
        sh = self._find(name)
        bgr = _resolve_color(color)
        sh.Fill.ForeColor.RGB = bgr
        self._log(name, "set_fill", {"color": str(color)})

    def move(self, name, left=None, top=None):
        """Move a shape to a new position (inches).

        Args:
            name: zrx_ shape name
            left: new x position in inches (None to keep current)
            top: new y position in inches (None to keep current)
        """
        sh = self._find(name)
        IN = PTS_PER_INCH
        before = {"left": round(sh.Left / IN, 3), "top": round(sh.Top / IN, 3)}

        if left is not None:
            sh.Left = left * IN
        if top is not None:
            sh.Top = top * IN

        after = {"left": round(sh.Left / IN, 3), "top": round(sh.Top / IN, 3)}
        self._log(name, "move", {"before": before, "after": after})

    def resize(self, name, width=None, height=None):
        """Resize a shape (inches).

        Args:
            name: zrx_ shape name
            width: new width in inches (None to keep current)
            height: new height in inches (None to keep current)
        """
        sh = self._find(name)
        IN = PTS_PER_INCH
        before = {"width": round(sh.Width / IN, 3), "height": round(sh.Height / IN, 3)}

        if width is not None:
            sh.Width = width * IN
        if height is not None:
            sh.Height = height * IN

        after = {"width": round(sh.Width / IN, 3), "height": round(sh.Height / IN, 3)}
        self._log(name, "resize", {"before": before, "after": after})

    def get_position(self, name):
        """Get current position and size of a shape.

        Returns:
            dict with keys: left, top, width, height (all in inches)
        """
        sh = self._find(name)
        IN = PTS_PER_INCH
        return {
            "left": round(sh.Left / IN, 4),
            "top": round(sh.Top / IN, 4),
            "width": round(sh.Width / IN, 4),
            "height": round(sh.Height / IN, 4),
        }

    def set_slide(self, slide_num: int):
        """Switch to a different slide number (1-indexed)."""
        self.slide_num = slide_num
        self._slide = self._prs.Slides(slide_num)

    def find_shapes_by_prefix(self, prefix: str) -> list[dict]:
        """Find all shapes whose name starts with the given prefix.

        Useful for finding all shapes on a pipeline-generated slide, e.g.
        find_shapes_by_prefix("zrx_005") returns all shapes on slide 5.

        Returns:
            list of dicts with keys: name, type, left, top, width, height, text
        """
        IN = PTS_PER_INCH
        results = []
        for i in range(1, self._slide.Shapes.Count + 1):
            sh = self._slide.Shapes(i)
            if sh.Name.startswith(prefix):
                info = {
                    "name": sh.Name,
                    "type": sh.Type,
                    "left": round(sh.Left / IN, 3),
                    "top": round(sh.Top / IN, 3),
                    "width": round(sh.Width / IN, 3),
                    "height": round(sh.Height / IN, 3),
                }
                try:
                    info["text"] = sh.TextFrame.TextRange.Text[:80]
                except Exception:
                    info["text"] = None
                results.append(info)
        return results

    def undo_last(self):
        """Undo the most recent edit by re-applying 'before' values.

        Returns:
            The edit that was undone, or None if no edits to undo.
        """
        if not self._edits:
            return None

        edit = self._edits.pop()
        name = edit["shape"]
        action = edit["action"]
        changes = edit["changes"]

        if action == "set_text" and "text" in changes:
            sh = self._find(name)
            sh.TextFrame.TextRange.Text = changes["text"]["before"]

        elif action == "move" and "before" in changes:
            IN = PTS_PER_INCH
            sh = self._find(name)
            sh.Left = changes["before"]["left"] * IN
            sh.Top = changes["before"]["top"] * IN

        elif action == "resize" and "before" in changes:
            IN = PTS_PER_INCH
            sh = self._find(name)
            sh.Width = changes["before"]["width"] * IN
            sh.Height = changes["before"]["height"] * IN

        return edit


# ── CLI ──────────────────────────────────────────────────────────────────────

def _cli():
    """Interactive CLI for making live edits."""
    if len(sys.argv) < 2:
        print("Usage: python -m slidegen.edit <filename.pptx> [slide_num]")
        sys.exit(1)

    filename = sys.argv[1]
    slide_num = int(sys.argv[2]) if len(sys.argv) > 2 else 1

    with LiveEditor(filename, slide_num) as editor:
        print(f"Connected to: {filename} (slide {slide_num})")
        print(f"Shapes on slide:")
        for sh in editor.list_shapes():
            text = f'  text="{sh["text"]}"' if sh["text"] else ""
            print(f"  {sh['name']}  ({sh['left']:.2f}\", {sh['top']:.2f}\"){text}")

        print(f"\nCommands: text <name> <value> | fill <name> <color> | "
              f"move <name> <left> <top> | undo | quit")

        while True:
            try:
                line = input("\n> ").strip()
            except (EOFError, KeyboardInterrupt):
                break
            if not line or line == "quit":
                break

            parts = line.split(maxsplit=2)
            cmd = parts[0].lower()

            try:
                if cmd == "text" and len(parts) >= 3:
                    editor.set_text(parts[1], parts[2])
                    print(f"  [OK] Text updated on {parts[1]}")
                elif cmd == "fill" and len(parts) >= 3:
                    editor.set_fill(parts[1], parts[2])
                    print(f"  [OK] Fill updated on {parts[1]}")
                elif cmd == "move" and len(parts) >= 3:
                    coords = parts[2].split()
                    left = float(coords[0]) if len(coords) > 0 else None
                    top = float(coords[1]) if len(coords) > 1 else None
                    editor.move(parts[1], left, top)
                    print(f"  [OK] Moved {parts[1]}")
                elif cmd == "undo":
                    edit = editor.undo_last()
                    if edit:
                        print(f"  [OK] Undone: {edit['action']} on {edit['shape']}")
                    else:
                        print("  Nothing to undo.")
                elif cmd == "list":
                    for sh in editor.list_shapes():
                        text = f'  "{sh["text"]}"' if sh["text"] else ""
                        print(f"  {sh['name']}  ({sh['left']:.2f}\", {sh['top']:.2f}\"){text}")
                else:
                    print(f"  Unknown command: {cmd}")
            except Exception as e:
                print(f"  [ERROR] {e}")

        print(f"Session complete. {len(editor._edits)} edits logged.")


if __name__ == "__main__":
    _cli()
