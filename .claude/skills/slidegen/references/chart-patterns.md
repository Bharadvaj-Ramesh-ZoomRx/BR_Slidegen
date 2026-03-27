# Chart Patterns -- Checklists and Examples

---

## Checklist: Every Horizontal Bar Chart

Apply all of these after building any horizontal bar chart:

```python
invert_cat_axis(chart)                          # first item at top
set_series_no_border(series)                    # for each series
set_datalabel_pos_outside_end(series)           # for each series
set_plot_area_gap(chart, 70)                    # standard gap
set_overlap(chart, -10)                         # small gap between bars in cluster
chart.has_legend = False                        # use manual_legend() instead
chart.has_title = False                         # use textbox() column title above
set_val_axis_number_format(chart.value_axis, "0")  # integer labels
```

---

## Checklist: Dot-Plot / Line+Marker Series

```python
set_series_line_style(series, visible=False)    # hide connecting line if dot-only
set_series_marker(series, "circle", size=12, fill_color=primary_color)
set_data_label_color(series, primary_color)
set_marker_data_label_pos(series, "r")          # labels to the right of each dot
```

For two-brand comparison:
```python
# Primary brand series
set_series_marker(primary_series, "circle", size=12, fill_color=primary_color)
set_data_label_color(primary_series, primary_color)
# Competitor series
set_series_marker(comp_series, "circle", size=12, fill_color=comp_color)
set_data_label_color(comp_series, comp_color)
# Benchmark (gray diamonds)
set_series_marker(avg_series, "diamond", size=8, fill_color=C_GREY)
set_series_line_style(avg_series, visible=False)
```

---

## Checklist: Stacked / 100% Stacked Bar

```python
invert_cat_axis(chart)
set_series_no_border(series)                    # for each series
set_stacked_label_pos(series, "ctr")           # labels inside each segment
chart.has_legend = False
chart.has_title = False
```

---

## Checklist: Line Chart (trend over time)

```python
# High-level builder (preferred):
add_line_chart(slide, categories, series_list,
    left, top, width, height, colors,
    label_positions=["t", "b"],  # first above, rest below (avoids overlap)
    marker_size=5, line_width_pt=2.25)

# Or manual styling:
set_series_line_style(series, width_pt=2.0, dash="solid")
set_series_smooth(series, False)
set_series_marker(series, "circle", size=6, fill_color=series_color)
set_marker_data_label_pos(series, "t")         # labels above each point
chart.has_title = False
```

---

## Checklist: Stacked Column Chart (vertical)

```python
add_stacked_column_chart(slide, categories, series_list,
    left, top, width, height, colors,
    label_fsize=10, num_fmt='0"%"', gap=80)
# Data labels centered inside segments (white text)
# Category axis visible at bottom, value axis hidden
```

---

## Checklist: Donut Chart

```python
set_pie_slice_colors(chart, [primary_color, comp_color, C_GREEN, C_LTGREY])
set_donut_hole_size(chart, pct=65)             # standard hole size
chart.has_legend = False
chart.has_title = False
```

---

## Checklist: Scatter / Quadrant Chart

```python
# Draw quadrant fills BEFORE add_chart so rects sit behind
solidrect(slide, ...)  # 4 quadrant background fills

# High-level builder:
add_scatter_chart(slide, series_list,
    left, top, width, height, colors,
    marker_size=7, x_min=0, x_max=1, y_min=0, y_max=1)

# Or manual styling:
set_series_marker(series, "circle", size=10, fill_color=primary_color)
set_series_line_style(series, visible=False)
hide_axis(chart, "val")
```

---

## Checklist: Heatmap Table

```python
# Render via heatmap_table slide type — uses pptx tables, not charts
# Per-cell gradient fills interpolated between low_color and high_color
# Delta columns: green for positive > threshold, red for negative > threshold
# Alternating row backgrounds
# All colors configurable via ask.extra in config.yaml
```

---

## Reference Line (Industry Average / Benchmark)

```python
add_val_axis_reference_line(chart, x_value=35, label="Benchmark",
                             color=C_GREY, dash="dash", width_pt=1.0)
```

Adds a dashed vertical line at a fixed position using a supplementary scatter series.

---

## Dual-Chart Layout (Left + Right panels)

The standard two-chart slide layout:

```python
# Layout constants
CHART_TOP  = 1.47
CHART_H    = 5.05
LEFT_LEFT  = 0.20
LEFT_W     = 7.30
LEFT_DELTA_W = 0.62
RIGHT_LEFT = LEFT_LEFT + LEFT_W + LEFT_DELTA_W + 0.10
RIGHT_W    = 4.40
RIGHT_DELTA_W = 0.62

# Left chart -- shows full Y-axis category labels
left_cs = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED,
    Inches(LEFT_LEFT), Inches(CHART_TOP), Inches(LEFT_W), Inches(CHART_H), left_cd)

# Right chart -- hides Y-axis (left chart already has labels)
right_cs = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED,
    Inches(RIGHT_LEFT), Inches(CHART_TOP), Inches(RIGHT_W), Inches(CHART_H), right_cd)
hide_cat_labels(right_cs.chart)

# Delta columns
add_delta_col(slide, left_deltas, left=LEFT_LEFT+LEFT_W, ...)
add_delta_col(slide, right_deltas, left=RIGHT_LEFT+RIGHT_W, ...)
```

---

## Common Mistakes

| Mistake | Fix |
|---|---|
| Bars in wrong order (bottom-first) | Call `invert_cat_axis(chart)` |
| Data labels overlap bars | Use `set_datalabel_pos_outside_end(series)` |
| Bars too thin | Reduce `set_plot_area_gap(chart, gap_pct)` -- lower = fatter bars |
| Right chart shows duplicate labels | Call `hide_cat_labels(right_chart)` |
| Quadrant fill rects sit on top of scatter | Draw `solidrect()` fills BEFORE `add_chart` |
| Line chart data labels overlap | Use `label_positions=["t", "b"]` in `add_line_chart()` |
| Chart open in PowerPoint + python-pptx edit | Never mix -- use COM for open files, python-pptx for closed |