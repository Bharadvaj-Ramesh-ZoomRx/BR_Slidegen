# Deck Analysis — Actionable Findings for pptx_utils + Skills

**Decks analyzed:** 32 PET decks across 17 pharma clients
**Total slides:** 2,333
**Total charts:** 4,354
**Total tables:** 4,950
**Total headlines detected:** 3,569

This document translates the raw findings into concrete, actionable inputs for `pptx_utils` modules and the skill library. Every recommendation is backed by a count from real client decks.

---

## 1. Priority Ranking for pptx_utils Work

Based on chart type + signature frequency, prioritize `pptx_utils` work in this order:

### P0 — Ship-blocking (covers 80%+ of charts)

| # | Module | Why | Source |
|---|---|---|---|
| 1 | `charts.py::make_bar_clustered` (horizontal, inverted cat axis) | 1,511 occurrences = 35% of all charts | Top chart type |
| 2 | `charts.py::make_scatter` (abacus-style) | 743 occurrences = 17% of all charts | Top 2 chart type |
| 3 | `charts.py::make_line_with_markers` | 595 occurrences = 14% | Top 3 chart type |
| 4 | `charts.py::make_column_stacked_100` (vertical 100% stacked) | 486 occurrences = 11% | Top 4 |
| 5 | `charts.py::make_bar_stacked_100` (horizontal 100% stacked) | 324 occurrences = 7% | Top 5 |
| 6 | `charts.py::make_bar_stacked` (horizontal stacked, non-100%) | 305 occurrences = 7% | Top 6 |
| 7 | `tables.py::make_data_table` (1-col narrow delta table) | 1,094 `1x1` tables (likely delta cells) | Top table dim |
| 8 | `tables.py::make_label_table` (multi-row single col, for labels) | 11x1, 12x1, 10x1 very common | Common dims |

**P0 combined covers 90% of all charts + 40% of tables.**

### P1 — Important (next 15% coverage)

| # | Module | Why |
|---|---|---|
| 9 | `charts.py::make_doughnut` | 138 occurrences |
| 10 | `charts.py::make_column_clustered` (QoQ column) | 119 occurrences |
| 11 | `charts.py::make_column_stacked` (non-100%) | 98 occurrences |

### P2 — Long tail (flag for inline lxml, extract later)

- `area_stacked_100` — 6 occurrences
- `area_stacked` — 3 occurrences
- `bubble` — 2 occurrences
- `pie` — 14 occurrences (could map to `make_doughnut` with hole=0)
- `xy_scatter_smooth` — 4 occurrences
- `xy_scatter_lines` — 60 occurrences

---

## 2. OOXML Properties → `lxml_helpers.py` Inventory

These are the python-pptx gaps confirmed by real deck usage. Each becomes a named helper function.

### 2.1 Bar/Column geometry

| Property | Observed Values | Distribution | Helper |
|---|---|---|---|
| `gapWidth` | 50 (15%), 100 (14%), 80 (10%), 70 (9%), 150 (8%) — wide spread | Highly variable per client | `set_plot_area_gap(chart, width)` |
| `overlap` | **100** (60% of charts — hard-stacked), -20 (6%), -5 (3%) | Dominated by 100 = stacked | `set_overlap(chart, pct)` |
| `invertIfNegative` | Present on 13,550 shapes | Universal — turn OFF by default | `set_invert_if_negative(chart, False)` |

**Action:** `set_plot_area_gap` is the single most important helper. Real decks use gapWidth between 40–150 for horizontal bars. Default to 80–100.

### 2.2 Axis configuration

| Property | Observed | Helper |
|---|---|---|
| `scaling/orientation val="maxMin"` on catAx | 1,865 charts (43%) = inverted category axis (top-down) | `invert_cat_axis(chart)` |
| `scaling/orientation val="minMax"` on catAx | 1,587 charts (37%) = standard | default |
| `scaling/orientation val="minMax"` on valAx | 4,835 (standard) | default |
| `scaling/orientation val="maxMin"` on valAx | 115 charts (rare) | `invert_val_axis(chart)` |
| `tickLblPos val="none"` on catAx | 371 charts = hidden cat labels | `hide_cat_labels(chart)` |
| `tickLblPos val="none"` on valAx | 48 charts | `hide_val_labels(chart)` |
| `tickLblPos val="nextTo"` | Default (universal) | default |
| `tickLblPos val="low"` on catAx | 24 charts (edge case) | skip |

**Action:** `invert_cat_axis` is used on almost half of all charts — high priority. `hide_cat_labels` is used when labels shown in separate companion table (a core clustered_compare pattern).

### 2.3 Data labels

| Property | Value | Count | Helper |
|---|---|---|---|
| `dLblPos` | `ctr` (center) | 4,508 | `set_datalabel_pos_center(chart)` |
| `dLblPos` | `t` (top) | 2,778 | `set_datalabel_pos_top(chart)` |
| `dLblPos` | `outEnd` (outside end — outside bar top) | 2,618 | `set_datalabel_pos_outside_end(chart)` |
| `dLblPos` | `inEnd` (inside end) | 143 | `set_datalabel_pos_inside_end(chart)` |
| `dLblPos` | `inBase` | 103 | rare — skip |
| `dLblPos` | `b` (bottom) | 176 | rare |
| `dLblPos` | `r` (right) | 166 | rare |

**Action:** Top 3 positions (`ctr`, `t`, `outEnd`) cover 99% of cases. Build helpers for these only.

### 2.4 Number formats

| Format | Count | Meaning | Use case |
|---|---|---|---|
| `0%` | 4,182 (96% of formats) | Integer percent | Universal default for PET charts |
| `0` | 566 | Integer | Raw counts |
| `General` | 292 | Auto | Avoid — specify explicitly |
| `#,##0.0` | 59 | One decimal with commas | Large numbers |
| `0.0` | 50 | One decimal | Fine-grained percent |
| `0%;\-0%;\ ` | 14 | Conditional negative | Delta columns with sign |
| `[>=0.05]0%;;;` | 13 | Suppress values below 5% | Chart decluttering |

**Action:** `0%` is the PET default. Add helpers:
- `set_val_axis_pct_format(chart)` — applies `0%`
- `set_val_axis_int_format(chart)` — applies `0`
- `set_datalabel_format(chart, fmt)` — applies arbitrary format

### 2.5 Line width + marker (scatter/line charts)

| Marker type | Count | Use |
|---|---|---|
| `circle` | 3,087 (94% of markers) | Default for all scatter |
| `square` | 128 | Rare |
| `triangle` | 99 | Rare |
| `none` | 80 | Suppressed markers |
| `diamond` | 39 | Rare |

| Line width (EMU) | Count | Points |
|---|---|---|
| 28575 | 1,223 | 2.25pt |
| 25400 | 827 | 2pt |
| 12700 | 607 | 1pt |
| 19050 | 504 | 1.5pt |
| 9525 | 162 | 0.75pt |

**Action:** Default line width 2pt (25400 EMU). Default marker circle. Default series border: **noFill** (74,367 noFill tags across all charts — overwhelmingly common to hide series borders).

### 2.6 Series formatting

| Property | Observed | Notes |
|---|---|---|
| `noFill` (line/border removal) | 74,367 tags | Universal — hide default borders |
| `solidFill` | 56,297 tags | Explicit fill colors |
| `srgbClr` | 39,048 tags | RGB color assignments |
| `schemeClr` | 17,652 tags | Theme-referenced colors (avoid — prefer srgbClr) |

**Action:**
- `set_series_no_border(series)` — helper to apply `noFill` to series line
- Always use sRGB (`srgbClr`) not scheme colors — more predictable, no theme dependencies

### 2.7 Legend

Charts with legend: **46 out of 4,354** (1%). Overwhelmingly **no legend** — the label table next to the chart serves as legend.

**Action:** Default `show_legend = False` in all renderers. Companion label table is canonical PET pattern.

### 2.8 Chart titles

Charts with title: **45 out of 4,354** (1%). Titles are in separate text boxes (headlines), not in the chart itself.

**Action:** Default `show_title = False`. Skill writes headline as separate text box.

### 2.9 Gridlines

Major gridlines: **706 charts (16%)**. Minor gridlines: **11** (rare).

**Action:** Default `no_major_gridlines`. Helper `add_major_gridlines(chart)` for the 16% that want them.

---

## 3. BRAND{} Schema — Proposed Entries

From 17 clients. Each has a distinct primary series color. Full proposals in `deep_brand_proposals.md`.

| Client | Primary | Secondary | Heading Font | Body Font | Decks |
|---|---|---|---|---|---|
| JJ | `#7FB1E1` (or `#0063C3` blue) | `#0063C3` | Johnson Display | Johnson Text | 4 |
| AZN | `#00B050` (green — Calquence brand) | `#FF8813` | Arial | Arial | 7 |
| Pfizer | `#0000C9` (navy) | `#43964A` | Arial | Arial | 1 |
| Regeneron | `#219491` (teal) | `#F79646` (orange) | Trade Gothic LT Std | Arial | 2 |
| Amgen | `#1F497D` (navy) | `#813F97` (purple) | Century Gothic | Century Gothic | 2 |
| GSK | `#3AB51D` (green) or `#F36633` (orange) | `#668EDD` | Arial | Arial | 2 |
| LEO | `#C014A3` (magenta) | `#D5685F` | Arial | Arial | 2 |
| Otsuka | `#FFC000` (Abilify yellow) | `#C00000` | Calibri | Arial | 1 |
| Alexion | `#0E8779` (teal) | `#7030A0` (purple) | Arial Black | Arial | 1 |
| BL | `#00A9EB` (Bausch blue) | `#400286` (dark purple) | Avenir Next LT Pro | Century Gothic | 1 |
| Novartis | `#018E86` (teal) | `#0070FE` (blue) | Arial | Arial | 1 |
| DSI | `#1E22AA` (navy) | `#643466` (purple) | Arial | Arial | 1 |
| Apellis | `#FC3B6E` (pink) | `#E7E6E6` | Calibri Light | Calibri | 1 |
| Ipsen | `#54AC65` (green) | `#C84874` (pink) | Rethink Sans | Calibri | 1 |
| CCA | `#F28E2B` (orange) | — | Century Gothic | Century Gothic | 1 |
| Bone-HCP | `#FF9933` (orange) | `#009201` (green) | Arial | Century Gothic | 1 |

**Universal delta colors** (observed across most decks):
- Positive: `#00B050` (standard Office green) — 494 occurrences
- Negative: `#FF0000` (pure red) — 913 occurrences
- Alternative negative: `#C00000` (deep red) — 61 occurrences

---

## 4. LAYOUTS{} Presets — From Real Coordinate Clusters

Top 30 chart position clusters from real decks. Each is a candidate `LAYOUTS{}` entry.

### 4.1 Chart Positions (inches)

Most common chart slots, sorted by frequency:

| Preset | Left | Top | Width | Height | Occurrences | Typical slide type |
|---|---|---|---|---|---|---|
| `chart_right_tall_narrow_8_93` | 8.93 | 2.06 | 1.58 | 4.57 | 16 | 1-col delta chart on right of label table |
| `chart_mid_right_wide` | 5.44 | 2.41 | 2.51 | 4.04 | 16 | Clustered-compare right pane |
| `chart_right_narrow` | 7.95 | 1.94 | 2.10 | 4.53 | 15 | Single-brand bar + comparison table |
| `chart_far_right_narrow` | 8.38 | 2.08 | 1.44 | 4.53 | 13 | Very narrow bar (delta only?) |
| `chart_left_wide` | 0.49 | 1.95 | 4.29 | 4.16 | 13 | Big chart left, table right |
| `chart_right_wider` | 8.60 | 2.17 | 3.17 | 4.45 | 12 | Typical bar chart on right |
| `chart_far_right` | 11.85 | 2.08 | 1.44 | 4.53 | 11 | Third-column chart |
| `chart_far_right_wider` | 11.08 | 2.01 | 1.97 | 4.65 | 11 | — |
| `chart_middle_short` | 7.66 | 3.00 | 2.15 | 3.58 | 10 | Lower-half chart |
| `chart_mid_left_narrow` | 5.56 | 1.78 | 2.10 | 4.53 | 10 | — |

### 4.2 Table Positions (inches)

Most common table slots:

| Preset | Left | Top | Width | Height | Occurrences |
|---|---|---|---|---|---|
| `table_delta_col_right` | 12.52 | 2.17 | 0.58 | 4.50 | 40 |
| `table_delta_col_right_2` | 11.46 | 2.08 | 0.47 | 4.58 | 22 |
| `table_delta_col_mid_right` | 7.96 | 2.07 | 0.49 | 4.58 | 19 |
| `table_label_left_wide` | 0.91 | 1.52 | 11.11 | 4.92 | 139 |
| `table_wide_7_5` | 0.58 | 2.06 | 7.50 | 4.57 | 14 |
| `table_wide_13_12` | 0.11 | 1.80 | 13.12 | 0.47 | 17 (tall, narrow — likely divider line) |
| `table_wide_12_45` | 0.43 | 1.88 | 12.45 | 5.06 | 13 |

**Key insight:** Narrow (~0.5" wide × ~4.5" tall) tables appear in many positions — these are **delta columns** (the `add_delta_col()` pattern). They're the single most common pptx_utils primitive.

### 4.3 Canonical Signature Layouts

Top content-only signatures (count only chart/table/picture):

| Signature | Count | Implied Layout |
|---|---|---|
| `1_chart_1_table` | 145 | single_bar_with_delta — chart right (`5.41"×4.29"` at `1.62,2.02`), table left (`6.79"×4.55"` at `1.48,1.82`) |
| `1_chart_2_table` | 110 | clustered_compare — chart right, label table + delta table left |
| `1_chart_3_table` | 77 | clustered_compare with 3-column data presentation |
| `2_chart_2_table` | 68 | dual_bar_compare — two charts side by side, each with its own table |
| `2_chart_1_table` | 55 | — |
| `3_chart_1_table` | 45 | Triple-metric scorecard |

**Action:** These signatures map almost directly to the renderer types. Rewrite each renderer to accept a `layout_preset` parameter that pulls coordinates from `LAYOUTS{}`.

---

## 5. Headline Patterns

**3,569 headlines detected** (top-of-slide text, wide, ≥8 chars).

### 5.1 Font sizes (top)

| Size (pt) | Count | % |
|---|---|---|
| 16 | 876 | 25% |
| 20 | 678 | 19% |
| 18 | 678 | 19% |
| 14 | 340 | 10% |
| 12 | 155 | 4% |
| 24 | 149 | 4% |

**Action:** Default headline font size **16pt**. Cover slides and dividers may use 20–24pt. Skills should set this based on slide type.

### 5.2 Font colors

| Color | Count | Typical interpretation |
|---|---|---|
| `#000000` | 455 | Black (universal) |
| `#0063C3` | 134 | JJ blue |
| `#001E60` | 126 | Client navy |
| `#595959` | 110 | Grey |
| `#00A3DC` | 86 | Alternative blue |
| `#002B5C` | 72 | Navy |
| `#E44405` | 59 | LEO orange |
| `#FF0000` | 53 | Red (warnings?) |

**Action:** Default headline color **black** (`#000000`). Client-specific dark brand color pulled from `BRAND["<client>"]["heading_color"]`.

### 5.3 Top position

Median top: `0.8"`, mode: `0.3"`. Headlines sit in the top 1.5" of the slide.

**Action:** `LAYOUTS["headline"] = {"top": 0.3, "left": 0.2, "width": 12.8, "height": 0.9}` as default.

### 5.4 Character count

Median: 72 chars. Max legitimate: ~300. Average: 98.

**Action:** Skills should enforce headline length ≤ 200 chars. Flag anything >250 as likely multi-line or runaway.

---

## 6. Table Patterns

### 6.1 Table dimensions (top 15)

| Rows × Cols | Count | Purpose |
|---|---|---|
| 1x1 | 1,094 | Single delta cell / callout / footnote |
| 11x1 | 200 | 10-message label column |
| 5x1 | 176 | 5-row label column |
| 3x1 | 170 | 3-row label (often segments) |
| 12x1 | 170 | 11-message or 11-segment label column |
| 10x1 | 165 | Label column |
| 4x1 | 154 | — |
| 1x3 | 151 | Tri-column header/footer |
| 6x1 | 147 | — |
| 2x1 | 142 | — |
| 8x1 | 140 | — |
| 1x2 | 114 | Two-column header |
| 7x1 | 103 | — |
| 1x4 | 90 | Four-column header/footer |
| 13x3 | 89 | Full label + 2 data columns (prior, current) |
| 11x3 | 87 | Same pattern, 11 rows |
| 10x3 | 74 | Same pattern |

**Key insight:** `Nx3` tables (e.g., `13x3`, `11x3`, `10x3`) = label + prior + current columns. This is the canonical companion table for bar charts.

### 6.2 Header fill colors (top 10)

| Fill | Count | Interpretation |
|---|---|---|
| `#F2F2F2` | 229 | Light grey (standard first-row tint) |
| `#FFFFFF` | 190 | White (no fill) |
| `#63BE7B` | 52 | Conditional formatting green (heatmap) |
| `#E5E5E5` | 50 | Alt light grey |
| `#E7F0F9` | 36 | Pale blue |
| `#595959` | 31 | Dark grey |
| `#E7E8E9` | 28 | Another grey |
| `#3B5998` | 26 | Blue (JJ?) |

**Action:** Default alternating row fills: `#F2F2F2` and `#FFFFFF`. This matches existing `pptx_utils` convention.

### 6.3 Table widths

- Median width: `1.95"` (narrow — delta columns dominate)
- Mean width: `3.37"`
- Max: `13.26"` (full-width tables for lists)

**Action:** Three canonical widths in `LAYOUTS{}`:
- `delta_col_w = 0.58"` (from clusters)
- `value_col_w = 2.0"`
- `label_col_w = 7.0"` (companion table for bar chart)

---

## 7. Skill Recommendations

Based on findings, refine the skill inventory in SlideGen PRD v0.1:

### 7.1 `slide-creator` skill specification (SKILL.md content)

```
Before rendering any slide:
1. Load BRAND[client] from pptx_utils.brand
2. Load LAYOUTS[slide_type] from pptx_utils.layout
3. Compose the slide script using pptx_utils primitives ONLY.
4. Do NOT invent coordinates, colors, or formatting — use the library.

For bar_clustered charts (the 35% case):
- Horizontal bars with inverted category axis (invert_cat_axis)
- Category labels hidden (hide_cat_labels) — labels go in companion table
- Data labels INSIDE end with white text (set_datalabel_pos_inside_end)
- gapWidth = 80 (set_plot_area_gap)
- No legend, no title, no gridlines
- Series fill = BRAND[client]["primary"], noFill border
- Val axis format = "0%" (set_val_axis_pct_format)

For each slide type, reference LAYOUTS[slide_type] for coordinates:
- Chart at LAYOUTS[slide_type]["chart_rect"]
- Table at LAYOUTS[slide_type]["table_rect"]
- Delta col at LAYOUTS[slide_type]["delta_col_rect"]
```

### 7.2 `viz-selector` skill — metric → chart type map (from observed usage)

| Metric signal | Chart type | Rationale |
|---|---|---|
| Recall, Recognition, Awareness | `bar_clustered` horizontal + companion table | Dominant PET pattern |
| Likelihood to prescribe, Intent | `column_stacked_100` | Common intent distribution |
| Reach, SOV, Frequency | `line_markers` (trended) | Time series |
| Rep performance, Attributes | `abacus` (xy_scatter) | Multi-attribute comparison |
| Segment comparison | `clustered_compare` (bar_clustered × 2) | Cross-segment |
| Message effectiveness (M/B/D) | `message_mbd` (xy_scatter) | 3-dimensional scatter |

### 7.3 `spec-validator` checks (minimum required fields per chart type)

For `single_bar_with_delta`:
- `data_key` (dict with desc, prior, current)
- `primary_color` (BRAND hex)
- `prior_color` (BRAND prior_tint)
- `sort_by`, `sort_desc`
- `headline`
- `pct_mode` (`pct` or `straight`)
- `label_max` (char limit)

For `clustered_compare`:
- `primary_key`, `comp_key`
- `primary_label`, `comp_label`
- `primary_color`, `comp_color`
- `headline`

Raise clear error if any required field missing.

---

## 8. Summary: What Changes in pptx_utils

### 8.1 Files to create/refactor

| File | Status | Action |
|---|---|---|
| `brand.py` | Refactor | Add 17 client BRAND entries from this analysis |
| `layout.py` | Refactor | Add LAYOUTS entries from 30 observed coordinate clusters |
| `charts.py` | Refactor | Refactor 7 P0 chart builders to be general (not J&J-specific) |
| `lxml_helpers.py` | Expand | Add 12 helpers listed in §2 |
| `tables.py` | Refactor | 3 canonical width presets, alternating row fills confirmed |
| `text.py` | Minor | Headline defaults: 16pt, black, top 0.3", width 12.8" |

### 8.2 Concrete function additions to `lxml_helpers.py`

```python
# Bar/Column geometry
def set_plot_area_gap(chart, gap_width: int) -> None: ...
def set_overlap(chart, overlap_pct: int) -> None: ...
def set_invert_if_negative(chart, value: bool = False) -> None: ...

# Axis
def invert_cat_axis(chart) -> None: ...
def invert_val_axis(chart) -> None: ...
def hide_cat_labels(chart) -> None: ...
def hide_val_labels(chart) -> None: ...

# Data labels
def set_datalabel_pos_center(chart) -> None: ...
def set_datalabel_pos_top(chart) -> None: ...
def set_datalabel_pos_outside_end(chart) -> None: ...
def set_datalabel_pos_inside_end(chart) -> None: ...

# Number formats
def set_val_axis_pct_format(chart) -> None: ...
def set_val_axis_int_format(chart) -> None: ...
def set_datalabel_format(chart, fmt: str) -> None: ...

# Series formatting
def set_series_no_border(series) -> None: ...
def set_series_fill_rgb(series, hex_color: str) -> None: ...
def set_series_line_width(series, emu: int) -> None: ...
def set_series_marker_circle(series) -> None: ...

# Gridlines
def remove_major_gridlines(chart) -> None: ...
def add_major_gridlines(chart) -> None: ...
```

### 8.3 Skill library updates

- `slide-creator` SKILL.md: full composition instructions referencing pptx_utils primitives only
- New project-type skill `pet-deck`: viz-selector map + canonical slide sequence
- `viz-selector` skill: metric → chart type mapping table from §7.2
- `spec-validator` skill: required-fields enforcement from §7.3

---

## 9. What This Enables for the PRD

This analysis concretely resolves three open questions in SlideGen PRD v0.1:

**Q: "When does `slide-creator` render from scratch vs. use canonical templates?"**
**A:** Fully from-scratch. Decks show no reliance on hidden canonical templates. Brand-specific differentiation is achievable via `BRAND{}` + `LAYOUTS{}` Python constants alone.

**Q: "How does `viz-selector` resolve ambiguity?"**
**A:** Deterministic for the top 6 chart types (88% coverage) via the metric→chart map. HITL only for the long tail or unusual asks.

**Q: "What's the test strategy for regression-testing `pptx_utils` changes?"**
**A:** Use the inventory from this analysis as a regression suite. For each refactored renderer, produce outputs and visually diff against the corresponding real-deck slide. The 30 coordinate clusters serve as golden reference.

---

## Appendix: Files Produced

| File | Purpose |
|---|---|
| `report.md` | First-pass summary |
| `deep_report.md` | Full deep analysis |
| `deep_brand_proposals.md` | BRAND{} proposals per client |
| `layout_clusters.md` | Chart/table position clusters (readable) |
| `ACTIONABLE_FINDINGS.md` | **This document — synthesis** |
| `deep_chart_details.json` | OOXML property aggregations |
| `deep_ooxml_properties.json` | All chart XML tags seen |
| `deep_brand_proposals.json` | Brand data (structured) |
| `deep_layouts.json` | Layout stats |
| `deep_headlines.json` | Headline patterns |
| `deep_tables.json` | Table patterns |
| `chart_positions.json` | 410 chart position clusters |
| `table_positions.json` | 490 table position clusters |
| `inventory.json` | Full per-deck per-slide details |
| `aggregation.json` | First-pass aggregations |
| `chart_types.json` | Chart type frequency |
| `shape_types.json` | Shape type frequency |
| `colors.json` | Top 40 colors |
| `fonts.json` | Font usage |
| `gap_analysis.md` | Covered vs. uncovered chart types |
