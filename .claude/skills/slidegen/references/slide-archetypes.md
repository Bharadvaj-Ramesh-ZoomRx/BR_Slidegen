# Slide Archetypes -- Layout Specs and pptx_utils Call Sequences

Nine distinct slide archetypes cover all slide types in ZoomRx/JnJ production decks.

---

## Archetype 1: Cover Slide

**When:** First slide of every deck.

**Visual:** Full red (`C_RED`) background. Large white title (top half). Subtitle smaller below title. Date bottom-left area. Client attribution bottom-right. J&J Innovative Medicine logo bottom-left. ZoomRx logo inline with subtitle area.

**Positioning:**
- Title: centered, top ~2.0", font Johnson Display, ~40pt, white, bold
- Subtitle: centered, top ~3.5", ~20pt, white
- Date: bottom-left ~6.5", ~10pt, white
- Logos: bottom row, ~0.4" from bottom

**Implementation:**
```python
cover_slide(slide, title="Report Title",
            subtitle="Subtitle text",
            date="Q4 2025",
            client_name="Johnson & Johnson",
            jj_logo_path="assets/jj_logo.png",
            zrx_logo_path="assets/zrx_logo.png")
```

---

## Archetype 2: Section Divider

**When:** Between major deck sections (methodology, analysis modules, appendix).

**Visual:** Off-white (`#F8F8F8`) background. Large bold red left-aligned section title, vertical center. J&J Innovative Medicine logo bottom-left. Thin red left-edge stripe (~0.10" wide, full height).

**Positioning:**
- Red stripe: left=0, top=0, width=0.10, height=7.5
- Title: left=0.5, top~2.5, width=10, font Johnson Display, ~44pt, C_RED, bold
- Logo: bottom-left, ~0.3" from left and bottom

**Implementation:**
```python
divider_slide(slide, title="Section Title",
              logo_path="assets/jj_logo.png")
```

---

## Archetype 3: Standard Horizontal Bar Chart (MR/ME)

**When:** Message recall, message effect, preference -- the core ZoomRx deliverable. Two charts side by side (MR left, ME right) with delta columns.

**Visual:** Section header bar at top. Two horizontal bar charts (Q4 orange, Q3 pale orange). Delta column right of each chart. Manual legend below charts. Module badge top-right. Section breadcrumb top-right. Callout boxes for key findings (dashed red border).

**Layout constants:**
```
CHART_TOP  = 1.47
CHART_H    = 5.05
MR_LEFT    = 0.20,  MR_W = 7.30
ME_LEFT    = 8.26,  ME_W = 4.40
DELTA_W    = 0.62   (for each delta column)
```

**pptx_utils call sequence:**
```python
slide_header(slide, "Headline")
section_header_bar(slide, "QUESTION TEXT", top=1.40)
module_badge(slide, "PERSONAL PROMOTION MODULE", color=C_RED)
section_breadcrumb(slide, "Key Findings - Personal Promotions")

# MR chart
mr_cs = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, ...)
# Apply: invert_cat_axis, set_series_no_border, set_datalabel_pos_outside_end,
#        set_plot_area_gap(70), set_overlap(-10), set_val_axis_number_format

# ME chart
me_cs = slide.shapes.add_chart(...)
hide_cat_labels(me_cs.chart)  # no Y-axis on right chart

add_delta_col(slide, mr_deltas, ...)
add_delta_col(slide, me_deltas, ...)
manual_legend(slide, q4_n=120, q3_n=108, ...)

# Key findings callout (right side or bottom)
cb = callout_box(slide, left=8.3, top=2.0, width=4.8, height=1.2,
                 text="Key finding text", border_color=C_RED)
cb.name = "zrx_NNN"

slide_footer(slide, "Source: ZoomRx SFEA Survey | ...")
```

---

## Archetype 4: Dot-Plot / Abacus

**When:** Attribute ratings by message or rep, multi-brand comparison (J&J vs AZ vs Industry Avg). Each row is a message/attribute; data points are large colored circles.

**Visual:** Category labels left. Large filled circles: orange=J&J Q4, pale orange=J&J Q3, purple=AZ, gray diamond=Industry Average. No Y-axis on right chart. Right panel: ME attribute dots. Delta columns rightmost. Optional dashed reference line at industry average X value.

**Key differentiators from bar chart:**
- Chart type: `XL_CHART_TYPE.LINE` or `XL_CHART_TYPE.XY_SCATTER`
- Markers large (size=12-16), lines hidden
- No Y-axis (hide_axis)
- Multiple series per brand

**pptx_utils call sequence:**
```python
# After adding chart:
for series in chart.series:
    set_series_line_style(series, visible=False)

set_series_marker(jj_q4_series, "circle", size=14, fill_color=C_RYB_Q4)
set_series_marker(jj_q3_series, "circle", size=14, fill_color=C_RYB_Q3)
set_series_marker(az_series,    "circle", size=14, fill_color=C_TAG)
set_series_marker(avg_series,   "diamond", size=10, fill_color=C_GREY)

set_data_label_color(jj_q4_series, C_RYB_Q4)
set_data_label_color(az_series, C_TAG)
set_marker_data_label_pos(jj_q4_series, "r")

hide_axis(chart, "val")    # suppress horizontal value axis

# Industry average reference line
add_val_axis_reference_line(chart, x_value=35, dash="dash", color=C_GREY)
```

---

## Archetype 5: Line / Trend Chart

**When:** Time-series data (monthly, quarterly), rolling 3-month averages (R3M). Multiple messages as separate lines. Dots at key endpoints.

**Visual:** Multiple orange lines (one per message). Dashed horizontal separators between metric groups (Reach / Frequency / SOV). Category axis = months; no axis lines shown. Value at leftmost and rightmost data points labeled. Section header bar at top.

**pptx_utils call sequence:**
```python
# Per series (one per message):
set_series_line_style(series, width_pt=1.5, dash="solid")
set_series_smooth(series, True)
set_series_marker(series, "circle", size=6, fill_color=C_RYB_Q4)
set_marker_data_label_pos(series, "r")   # endpoint label
set_data_label_color(series, C_RYB_Q4)

hide_axis(chart, "val")
hide_axis(chart, "cat")

# Separators between metric groups
dashed_separator(slide, left=0.2, top=3.5, width=12.8)  # between Reach and Frequency
```

---

## Archetype 6: Scorecard / Table

**When:** Multi-metric summary comparing messages or reps across multiple KPIs. Each cell shows a value; rows are messages, columns are metrics.

**Visual:** Table with alternating row shading (`C_LBGREY`). Header row dark (`C_HDRGREY`). Trend arrows inline in cells (green up, red down, yellow flat). Mini bar charts in some columns. Delta values in parentheses.

**pptx_utils call sequence:**
```python
# Table via python-pptx add_table (not in pptx_utils -- use directly)
# Alternating row colors applied via pptx table cell fill

# Trend arrows inline in scorecard cells
trend_arrow_icon(slide, "up",   left=col_left, top=row_top)     # green triangle
trend_arrow_icon(slide, "down", left=col_left, top=row_top)     # red triangle
trend_arrow_icon(slide, "flat", left=col_left, top=row_top)     # yellow double arrow

# Optional mini bar charts alongside table (embedded via add_chart with small dims)
```

---

## Archetype 7: Callout / Insight Slide

**When:** Executive summary, qualitative findings, methodology explanation. Text-heavy. Key insights in dashed callout boxes.

**Visual:** Large dashed callout boxes (one per insight or finding group). Some boxes colored by theme (red=concern, orange=primary, green=positive). Text inside boxes in FONT_TEXT. Optional stat callouts (big number + label) for key metrics.

**Layout patterns:**
- Two-column: 2 large boxes side by side (~5.5" wide each)
- Three-column: 3 boxes (~3.8" wide each)
- Single wide: one box spanning most of slide width
- Top methodology + bottom table: boxes at top, data table below

**pptx_utils call sequence:**
```python
# Primary insight box (red border = J&J finding)
cb1 = callout_box(slide, left=0.3, top=1.6, width=5.8, height=4.5,
                  text="Key finding text here.", border_color=C_RED)
cb1.name = "zrx_001"

# Secondary box (orange = Q4 highlight)
cb2 = callout_box(slide, left=6.5, top=1.6, width=5.8, height=4.5,
                  text="Supporting finding.", border_color=C_RYB_Q4)
cb2.name = "zrx_002"

# Optional big stat
stat_callout(slide, value="16", delta=-1, label="years avg treatment duration",
             left=0.5, top=5.0)
```

---

## Archetype 8: Scatter / Quadrant Map

**When:** Importance vs performance, stated vs derived importance. Four-quadrant grid; bubbles positioned by two numeric dimensions.

**Visual:** Four colored background quadrants: blue top-left, green top-right, gray bottom-left, yellow bottom-right. Scatter points (orange J&J, purple AZ) positioned within quadrants. Text labels near each point. Optional callout box highlighting key findings. Axis labels for each quadrant.

**pptx_utils call sequence:**
```python
# MUST call BEFORE add_chart to get correct z-order
scatter_quadrant_fills(slide, chart_left=1.0, chart_top=1.5, chart_width=10.0, chart_height=5.0,
    tl_color=None,     # default blue
    tr_color=None,     # default green
    bl_color=None,     # default grey
    br_color=None)     # default yellow

cs = slide.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER,
    Inches(1.0), Inches(1.5), Inches(10.0), Inches(5.0), cd)

set_series_marker(jj_series, "circle", size=10, fill_color=C_RYB_Q4)
set_series_marker(az_series, "circle", size=10, fill_color=C_TAG)
set_series_line_style(jj_series, visible=False)
set_series_line_style(az_series, visible=False)

# Quadrant label textboxes (4 corners)
textbox(slide, "High Importance\nLow Performance", left=1.1, top=1.6, ...)
textbox(slide, "High Importance\nHigh Performance", left=7.0, top=1.6, ...)
```

---

## Archetype 9: Donut / Pie Chart

**When:** Share-of-wallet, interaction mix breakdown, category split. Often shown as donut with percentage center stat.

**Visual:** Donut chart (Q4 outer ring = solid fill, Q3 inner ring = paler shade). Center label: metric description. Legend right of chart. Colors match brand (orange=J&J, purple=AZ).

**pptx_utils call sequence:**
```python
cs = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT,
    Inches(left), Inches(top), Inches(w), Inches(h), cd)
chart = cs.chart

set_pie_slice_colors(chart, [C_RYB_Q4, C_TAG, C_GREEN, C_LTGREY])
set_donut_hole_size(chart, pct=65)
chart.has_legend = False
chart.has_title = False

# Center stat (positioned manually as textbox)
textbox(slide, "% of\nInteractions\nRated 6,7",
        left=center_x - 0.8, top=center_y - 0.5, width=1.6, height=1.0,
        fsize=8, align=PP_ALIGN.CENTER, color=C_GREY)
```
