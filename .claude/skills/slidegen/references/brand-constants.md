# Brand Constants -- Colors, Fonts, Positioning, and Spacing

---

## Color Palette

### python-pptx Colors (RGB hex -- use with RGBColor or pptx_utils constants)

| Constant | Hex | Use |
|---|---|---|
| `C_RED` | `#FF0000` | J&J red -- title bar accent, headline, badge, callout borders (primary brand) |
| `C_RYB_Q4` | `#F75824` | Deep orange -- Q4 bars, primary data accent, dot-plot J&J markers |
| `C_RYB_Q3` | `#FFC199` | Pale orange -- Q3 bars, J&J Q3 markers |
| `C_TAG` | `#7030A0` | Purple -- AstraZeneca, TAGRISSO series, AZ data labels |
| `C_GREEN` | `#00B050` | Positive delta indicators, NPP module badge, positive trend arrow |
| `C_WHITE` | `#FFFFFF` | White fill, white text on red backgrounds |
| `C_GREY` | `#505050` | Body text, industry average markers |
| `C_FTGREY` | `#7F7F7F` | Footer text, faint labels |
| `C_LBGREY` | `#F4F4F4` | Alternating table row background (light) |
| `C_HDRGREY` | `#404040` | Delta table header background (dark) |
| `C_LTGREY` | `#BFBFBF` | Gridlines, borders, flat trend arrow |

### COM Colors (BGR int -- use with com_set_text, com_set_fill)

| Constant | Value | RGB equivalent |
|---|---|---|
| `COM_RED` | `0x0000FF` | Red |
| `COM_ORANGE` | `0x2458F7` | Deep orange (C_RYB_Q4) |
| `COM_GREEN` | `0x50B000` | Green |
| `COM_GREY` | `0x505050` | Grey |

**BGR note:** COM uses byte-reversed RGB. `0x0000FF` = red (FF in blue position, 00 in red). Always use the named constants, never hardcode.

---

## Typography

| Role | Font | Size | Style | Color |
|---|---|---|---|---|
| Slide headline | Johnson Display | 22-26pt | Bold | `C_RED` or white |
| Section divider title | Johnson Display | 36-44pt | Bold | `C_RED` |
| Cover slide title | Johnson Display | 36-44pt | Bold | white |
| Section header bar label | Johnson Text | 9pt | Bold, ALL CAPS | white or dark |
| Chart column title (above chart) | Johnson Text | 9-10pt | Bold | `C_GREY` or `C_HDRGREY` |
| Body text / callout text | Johnson Text | 8-9pt | Regular | `C_GREY` |
| Delta column header | Johnson Text | 7-8pt | Bold | white on `C_HDRGREY` |
| Delta column values | Johnson Text | 7-8pt | Regular | green / red |
| Footer / source | Johnson Text | 6-7pt | Regular | `C_FTGREY` |
| Module badge | Johnson Text | 7pt | Bold, ALL CAPS | white on badge color |
| Breadcrumb | Johnson Text | 7pt | Regular | `C_FTGREY` |

**Font constants:** `FONT_DISPLAY = "Johnson Display"` / `FONT_TEXT = "Johnson Text"`

---

## Slide Dimensions

| Property | Value |
|---|---|
| Width | 13.333 inches |
| Height | 7.500 inches |
| Aspect ratio | 16:9 |

---

## Standard Positioning Grid

### Y-axis (top-to-bottom)

| Element | Top (inches) | Height (inches) |
|---|---|---|
| Accent line (slide_header) | 0.23 | 0.05 |
| Headline text (slide_header) | 0.30 | 0.80 |
| Module badge / breadcrumb | 0.25 | 0.28 |
| section_header_bar | 1.40 | 0.27 |
| Chart area top (standard) | 1.68 | varies |
| slide_footer | 6.93 | 0.35 |

### X-axis (left-to-right) -- Two-chart layout

| Element | Left (inches) | Width (inches) |
|---|---|---|
| MR chart | 0.20 | 7.30 |
| MR delta column | 7.54 | 0.62 |
| Gap between charts | 8.16 | 0.10 |
| ME chart | 8.26 | 4.40 |
| ME delta column | 12.70 | 0.62 |

### X-axis -- Single wide chart

| Element | Left (inches) | Width (inches) |
|---|---|---|
| Single chart | 0.20 | 12.80 |
| Callout box (right panel) | 8.30 | 4.80 |

---

## Module Badge Colors

| Module | Badge color constant |
|---|---|
| Personal Promotion Module | `C_RED` |
| Non-Sales Rep Promotions (NPP) | `C_GREEN` |
| Market Access / Payor | use C_TAG (purple) or custom |

---

## Callout Box Patterns

Callout boxes use `callout_box()` with these border color conventions:

| Border color | Meaning / use |
|---|---|
| `C_RED` | J&J finding, primary insight, concern |
| `C_RYB_Q4` | Q4 highlight, secondary finding |
| `C_GREEN` | Positive result, NPP finding |
| `C_TAG` | AZ-related finding |
| Yellow / `#FFD966` | Warning, caution, notable exception |

Border is dashed (`dashed=True`) by default -- this is the standard across all data slides. Solid borders are rare (methodology slides only).

Fill is auto-computed as 12% tint of the border color (`fill_color` parameter is optional).

---

## Section Header Bar

`section_header_bar(slide, label, top=1.40, icon_path=None)` produces:
- Gray background strip (~0.27" tall) from x=0.20 to x=12.80
- Small 16px icon image at left (optional, skip if `icon_path=None`)
- Bold all-caps label text in `FONT_TEXT`, ~9pt, white or dark

Standard top position: **1.40"** (directly under slide_header separator line).

---

## Delta Column Style

Produced by `add_delta_col(slide, deltas, left, top, width, height, header)`:
- Header row: dark background (`C_HDRGREY`), white bold text
- Alternating row shading: `C_LBGREY` (light gray) for even rows
- Positive delta: `C_GREEN` text
- Negative delta: `C_RED` text
- Zero delta: `C_GREY` text
- N/A (new message with no Q3): gray italic "N/A"
- Delta header label: "D Q4-Q3" (MR column) or "D" (ME column)

---

## Trend Arrow Icons

`trend_arrow_icon(slide, direction, left, top)`:

| Direction | Symbol | Color | Meaning |
|---|---|---|---|
| `"up"` | Triangle up | `C_GREEN` | Positive trend |
| `"down"` | Triangle down | `C_RED` | Negative trend |
| `"flat"` | Double arrow | `C_LTGREY` | No meaningful change |

Typical size: ~0.20" x 0.20". Place to the left or right of a metric value in a scorecard cell.
