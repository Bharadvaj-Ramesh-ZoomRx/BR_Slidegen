---
name: jj-slide-style
description: "Reference skill for detailed slide archetype specs, chart type selection, and brand constants. Use only when the slidegen skill cross-references this skill's detailed reference files. For all pipeline work (creating/editing slides), use the slidegen skill instead."
---

# jj-slide-style

## Overview

This skill codifies the ZoomRx/JnJ slide design system reverse-engineered from production decks. It covers nine slide archetypes, chart type selection, brand constants, and positioning rules -- all generic enough to apply across JnJ projects.

**Companion skill:** `pptx-utils` provides the Python functions that implement these patterns.

---

## Slide Archetypes -- Quick Reference

| Archetype | When to use | Key functions |
|---|---|---|
| Cover | First slide of deck | `cover_slide()` |
| Section divider | Between major sections | `divider_slide()` |
| Standard bar chart | MR/ME reach/preference data | `slide_header`, `add_delta_col`, `manual_legend` |
| Dot-plot / abacus | Message-level attribute ratings | `set_series_marker`, `hide_axis` |
| Line / trend chart | Time-series, R3M rolling | `set_series_line_style`, `set_series_smooth` |
| Scorecard table | Multi-metric summary | `add_delta_col`, `trend_arrow_icon` |
| Callout / insight | Exec summary, qualitative findings | `callout_box`, `textbox` |
| Scatter / quadrant | Importance vs performance | `scatter_quadrant_fills`, `add_chart` |
| Donut / pie | Share-of-wallet, interaction mix | `set_pie_slice_colors`, `set_donut_hole_size` |

See `references/slide-archetypes.md` for full layout specs and pptx_utils call sequences.
See `references/chart-types.md` for chart selection guidance.
See `references/brand-constants.md` for colors, fonts, spacing, and positioning rules.

---

## Universal Slide Elements

Every data slide (non-cover, non-divider) includes all of:

1. **`slide_header(slide, headline)`** -- red accent line + headline + badge + separator at top
2. **`section_header_bar(slide, label, top=1.40)`** -- gray strip with chart/question title
3. **`module_badge(slide, label, color)`** -- colored pill top-right labeling the module
4. **`section_breadcrumb(slide, text)`** -- small gray text top-right for section context
5. **`slide_footer(slide, footer_text)`** -- source + methodology footnote at bottom

These five are always present. Add chart/data content between `section_header_bar` and `slide_footer`.

---

## Positioning Grid

```
Y=0.00  Top of slide
Y=0.23  slide_header accent line (red, full width)
Y=0.30  Module label / headline text
Y=1.40  section_header_bar top  (standard)
Y=1.68  Chart top (standard: header_bar bottom + small gap)
Y=6.55  slide_footer top
Y=7.50  Bottom of slide

X=0.20  Left chart left edge (standard)
X=7.54  MR chart right edge / delta col left edge
X=8.26  ME chart left edge (two-chart layout)
X=12.70 ME delta col right edge
X=13.33 Right of slide
```
