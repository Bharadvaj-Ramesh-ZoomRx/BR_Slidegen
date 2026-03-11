---
name: pptx-utils
description: "Reference skill for pptx_utils function signatures, chart pattern checklists, and COM helpers. Use only when you need detailed function-level reference (see references/function-ref.md). For all pipeline work (creating/editing slides), use the slidegen skill instead."
---

# pptx-utils

## Overview

`slidegen.pptx_utils` is the master helper package for all SlideGen PowerPoint generation. It exports 50+ functions across five categories: brand constants, XML chart helpers, shape builders, COM live-edit helpers, and registry helpers.

**Core rule:** `from slidegen.pptx_utils import *` in every slide script. Never write raw lxml or hardcode hex colors.

---

## Workflow Decision Tree

```
Is the file open in PowerPoint right now?
  YES  ->  COM workflow (win32com)
           1. python -m slidegen reconcile file.pptx
           2. prs  = com_connect("file.pptx")
           3. sh   = com_find_shape(slide, "zrx_NNN")
           4. com_set_text / com_move / com_resize / com_set_fill

  NO   ->  Creation workflow (python-pptx)
           1. from slidegen.pptx_utils import *
           2. prs  = Presentation(TEMPLATE)
           3. slide = prs.slides.add_slide(blank_layout)
           4. Build shapes -> name each zrx_NNN -> prs.save() -> save_registry()
```

---

## Quick Start -- New Slide

```python
from pptx import Presentation
from slidegen.pptx_utils import *
from datetime import datetime

prs = Presentation(r"ppt via python\JJ PET RYBREVANT+LAZCLUZE Q4'25 Report.pptx")
blank_layout = next(l for l in prs.slide_layouts if "blank" in l.name.lower())
slide = prs.slides.add_slide(blank_layout)

# Strip template slides -- keep only the new one
sldIdLst = prs.slides._sldIdLst
for sldId in list(sldIdLst)[:-1]:
    sldIdLst.remove(sldId)

slide_header(slide, "Headline text")
tb = textbox(slide, "Body text", left=0.3, top=1.5, width=5, height=0.5, fsize=10)
tb.name = "zrx_001"

prs.save("output.pptx")
save_registry({
    "meta": {"created": datetime.now().isoformat(timespec="seconds"),
              "pptx_file": "output.pptx", "slide_index": 1},
    "shapes": {
        "zrx_001": {"label": "body text", "type": tb.shape_type,
                    "left": 0.3, "top": 1.5, "width": 5, "height": 0.5,
                    "text": "Body text", "last_reconciled": None}
    }
})
```

---

## Rules

- Name every shape `zrx_NNN` (sequential integers, never reuse within a file)
- Inches everywhere in creation code -- pptx_utils converts to EMU/points internally
- Write `slide_registry.json` after every creation script -- never skip
- ASCII-only in print statements (cp1252 terminal encoding)
- Run `python -m slidegen reconcile <file.pptx>` before every COM edit session
- Never mix python-pptx and win32com on the same open file

---

## Key Constants

| Name | Value | Use |
|---|---|---|
| `C_RED` | `#FF0000` | J&J red -- title bar, headline, badge |
| `C_RYB_Q4` | `#F75824` | Deep orange -- Q4 bars, primary accent |
| `C_RYB_Q3` | `#FFC199` | Pale orange -- Q3 bars |
| `C_TAG` | `#7030A0` | Purple -- AstraZeneca / TAGRISSO |
| `C_GREEN` | `#00B050` | Positive delta, NPP module badge |
| `C_GREY` | `#505050` | Body text |
| `FONT_DISPLAY` | `"Johnson Display"` | Headlines |
| `FONT_TEXT` | `"Johnson Text"` | Body text |
| `SLIDE_W_IN` | `13.333` | Slide width |
| `SLIDE_H_IN` | `7.500` | Slide height |
| `IN` | `72` | 1 inch in COM points |

---

## Function Reference

See `references/function-ref.md` for all 50+ function signatures organized by category.
See `references/chart-patterns.md` for chart type checklists and the decision guide.

The library is at `slidegen/pptx_utils/` (a package with brand.py, charts.py, tables.py, etc.). Import via `from slidegen.pptx_utils import *`.
