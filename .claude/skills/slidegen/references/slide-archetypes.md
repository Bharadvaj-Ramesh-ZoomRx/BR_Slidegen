# Slide Archetypes -- Layout Specs and pptx_utils Call Sequences

Ten distinct slide archetypes cover all slide types in production decks.

---

## Archetype 1: Cover Slide

**When:** First slide of every deck.

**Visual:** Branded background. Large white title (top half). Subtitle smaller below title. Date bottom-left area. Client attribution bottom-right. Logos at bottom.

**Positioning:**
- Title: centered, top ~2.0", display font, ~40pt, white, bold
- Subtitle: centered, top ~3.5", ~20pt, white
- Date: bottom-left ~6.5", ~10pt, white
- Logos: bottom row, ~0.4" from bottom

**Implementation:**
```python
cover_slide(slide, title="Report Title",
            subtitle="Subtitle text",
            date="Q4 2025",
            client_name="Client Name")
```

---

## Archetype 2: Section Divider

**When:** Between major deck sections (methodology, analysis modules, appendix).

**Visual:** Off-white background. Large bold left-aligned section title, vertical center. Client logo bottom-left. Thin accent left-edge stripe (~0.10" wide, full height).

**Implementation:**
```python
divider_slide(slide, title="Section Title",
              logo_path="assets/logo.png")
```

---

## Archetype 3: Standard Horizontal Bar Chart (Dual Panel)

**When:** Message recall, message effectiveness, preference — two metrics side by side with delta columns.

**Visual:** Section header bar at top. Two horizontal bar charts (current + prior period colors). Delta column right of each chart. Manual legend below charts. Module badge top-right. Section breadcrumb top-right. Callout boxes for key findings.

**Layout constants:**
```
CHART_TOP  = 1.47
CHART_H    = 5.05
LEFT_PANEL_LEFT  = 0.20,  LEFT_W = 7.30
RIGHT_PANEL_LEFT = 8.26,  RIGHT_W = 4.40
DELTA_W    = 0.62
```

**pptx_utils call sequence:**
```python
slide_header(slide, "Headline")
section_header_bar(slide, "SECTION LABEL", top=1.40)
module_badge(slide, "MODULE NAME", color=accent_color)
section_breadcrumb(slide, "Key Findings - Section")

# Left chart
left_cs = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, ...)
# Apply: invert_cat_axis, set_series_no_border, set_datalabel_pos_outside_end,
#        set_plot_area_gap(70), set_overlap(-10), set_val_axis_number_format

# Right chart
right_cs = slide.shapes.add_chart(...)
hide_cat_labels(right_cs.chart)  # no Y-axis on right chart

add_delta_col(slide, left_deltas, ...)
add_delta_col(slide, right_deltas, ...)
manual_legend(slide, ...)

slide_footer(slide, "Source: ...")
```

---

## Archetype 4: Dot-Plot / Abacus

**When:** Attribute ratings, multi-brand comparison. Each row is a message/attribute; data points are large colored circles.

**Visual:** Category labels left. Large filled circles per brand/period. Optional gray diamond for benchmark. No Y-axis on right chart. Delta columns rightmost. Optional dashed reference line at benchmark.

**Key differentiators from bar chart:**
- Chart type: `XL_CHART_TYPE.LINE` or `XL_CHART_TYPE.XY_SCATTER`
- Markers large (size=12-16), lines hidden
- No Y-axis (hide_axis)
- Multiple series per brand

**pptx_utils call sequence:**
```python
for series in chart.series:
    set_series_line_style(series, visible=False)

set_series_marker(primary_series, "circle", size=14, fill_color=primary_color)
set_series_marker(comp_series,    "circle", size=14, fill_color=comp_color)
set_series_marker(avg_series,     "diamond", size=10, fill_color=C_GREY)

set_data_label_color(primary_series, primary_color)
set_marker_data_label_pos(primary_series, "r")

hide_axis(chart, "val")

# Benchmark reference line
add_val_axis_reference_line(chart, x_value=35, dash="dash", color=C_GREY)
```

---

## Archetype 5: Line / Trend Chart

**When:** Time-series data (monthly, quarterly). Multiple series tracked over periods. Dots at data points.

**Visual:** Multiple colored lines (one per series/brand). Dashed separators between metric groups. Category axis = time periods; axes usually hidden. Data labels at each point. Section header bar at top.

**Two pipeline renderers:**
- `trended_scorecard` — multiple mini line charts stacked vertically. Each panel has metric label on left, line chart on right, alternating grey row backgrounds.
- `trended_activity` — side-by-side panels mixing line charts + stacked column charts (e.g. Reach, SOV, Frequency).

**pptx_utils call sequence:**
```python
# High-level builder (preferred):
add_line_chart(slide, categories, series_list,
    left, top, width, height, colors,
    marker_size=5, line_width_pt=2.25,
    label_positions=["t", "b"],  # first series above, rest below (avoids overlap)
    num_fmt='0"%"',
    scale_min=25, scale_max=90,
    hide_axes=True)

# Or manual per-series styling:
set_series_line_style(series, width_pt=1.5, dash="solid")
set_series_smooth(series, False)
set_series_marker(series, "circle", size=6, fill_color=series_color)
set_marker_data_label_pos(series, "r")
set_data_label_color(series, series_color)

hide_axis(chart, "val")
hide_axis(chart, "cat")

# Stacked column for share panels:
add_stacked_column_chart(slide, categories, series_list,
    left, top, width, height, colors,
    label_fsize=10, num_fmt='0"%"', gap=100)
```

---

## Archetype 6: Scorecard / Table

**When:** Multi-metric summary comparing items across multiple KPIs. Each cell shows a value; rows are items, columns are metrics.

**Visual:** Table with alternating row shading (`C_LBGREY`). Header row dark (`C_HDRGREY`). Trend arrows inline in cells (green up, red down, yellow flat). Optional mini bar charts. Delta values in parentheses.

**pptx_utils call sequence:**
```python
# Table via python-pptx add_table
# Alternating row colors applied via pptx table cell fill

# Trend arrows inline in scorecard cells
trend_arrow_icon(slide, "up",   left=col_left, top=row_top)
trend_arrow_icon(slide, "down", left=col_left, top=row_top)
trend_arrow_icon(slide, "flat", left=col_left, top=row_top)
```

---

## Archetype 7: Callout / Insight Slide

**When:** Executive summary, qualitative findings, methodology explanation. Text-heavy. Key insights in dashed callout boxes.

**Visual:** Large dashed callout boxes (one per insight). Boxes colored by theme via `border_color`. Text inside boxes in body font. Optional stat callouts (big number + label).

**Layout patterns:**
- Two-column: 2 large boxes side by side (~5.5" wide each)
- Three-column: 3 boxes (~3.8" wide each)
- Single wide: one box spanning most of slide width

**pptx_utils call sequence:**
```python
cb1 = callout_box(slide, left=0.3, top=1.6, width=5.8, height=4.5,
                  text="Key finding text here.", border_color=accent_color)
cb1.name = "zrx_001"

# Optional big stat
stat_callout(slide, value="16", delta=-1, label="metric description",
             left=0.5, top=5.0)
```

---

## Archetype 8: Scatter / Quadrant Map

**When:** Importance vs performance, stated vs derived importance. Four-quadrant grid; points positioned by two numeric dimensions.

**Visual:** Four colored background quadrants (configurable fills). Scatter points positioned within quadrants. Text labels near each point. Dashed divider lines at midpoints. Axis labels: X-axis centered below chart, Y-axis rotated 270° along left edge. Optional average point in different color.

**Pipeline renderer:** `quadrant_scatter`

**pptx_utils call sequence:**
```python
# High-level builder:
add_scatter_chart(slide, series_list,
    left=1.35, top=2.0, width=10.54, height=4.19,
    colors=[marker_color], marker_size=7,
    x_min=0, x_max=0.40, y_min=0.45, y_max=0.75)

# Quadrant fills drawn BEFORE chart (z-order):
solidrect(slide, left, top, left_w, top_h, tl_fill)
solidrect(slide, left + left_w, top, right_w, top_h, tr_fill)
solidrect(slide, left, top + top_h, left_w, bot_h, bl_fill)
solidrect(slide, left + left_w, top + top_h, right_w, bot_h, br_fill)

# Midpoint dividers:
dashed_separator(slide, left + left_w, top, chart_h, vertical=True)
dashed_separator(slide, left, top + top_h, chart_w, vertical=False)

# Y-axis label (rotated 270°): set bodyPr vert="vert270"
# X-axis label: centered textbox below chart
# Data point labels: textboxes positioned near each dot
```

---

## Archetype 9: Donut / Pie Chart

**When:** Share-of-wallet, interaction mix breakdown, category split. Often shown as donut with percentage center stat.

**Visual:** Donut chart (current outer ring = solid fill, prior inner ring = paler shade). Center label: metric description. Legend right of chart. Colors from project brand config.

**pptx_utils call sequence:**
```python
cs = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT,
    Inches(left), Inches(top), Inches(w), Inches(h), cd)
chart = cs.chart

set_pie_slice_colors(chart, [primary_color, comp_color, C_GREEN, C_LTGREY])
set_donut_hole_size(chart, pct=65)
chart.has_legend = False
chart.has_title = False

# Center stat (positioned manually as textbox)
textbox(slide, "% of\nResponses",
        left=center_x - 0.8, top=center_y - 0.5, width=1.6, height=1.0,
        fsize=8, align=PP_ALIGN.CENTER, color=C_GREY)
```

---

## Archetype 10: Heatmap Table

**When:** Multi-dimensional matrix (e.g. message recall by channel). Each cell shows a value with gradient fill intensity proportional to the value. QoQ deltas shown alongside.

**Visual:** Table with gradient fills (lighter = lower, darker = higher). Alternating row backgrounds. Header row with column names + sample sizes. Delta columns integrated next to each value column. Deltas color-coded: green for positive > threshold, red for negative > threshold.

**Pipeline renderer:** `heatmap_table`

**Layout:**
```
Label col (1.54") | Value+Delta (1.54"+0.63") × N columns
```

**Data format:**
```yaml
rows:
  - label: "Row Label"
    values: {"Col A": 47, "Col B": 40, ...}
    deltas: {"Col A": 2, "Col B": -3, ...}
```

**Configurable colors (via ask.extra):**
- `low_color` / `high_color`: heatmap gradient endpoints
- `delta_threshold`: abs value above which deltas get colored
- Row alternation colors are template-matched defaults