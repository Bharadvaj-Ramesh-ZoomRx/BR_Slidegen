# Chart Patterns -- Checklists and Examples

---

## Checklist: Every Horizontal Bar Chart

Apply all of these after building any horizontal bar chart:

```python
invert_cat_axis(chart)                          # first item at top
set_series_no_border(series)                    # for each series
set_datalabel_pos_outside_end(series)           # for each series
set_plot_area_gap(chart, 70)                    # standard gap
set_overlap(chart, -10)                         # small gap between Q3/Q4 bars
chart.has_legend = False                        # use manual_legend() instead
chart.has_title = False                         # use textbox() column title above
set_val_axis_number_format(chart.value_axis, "0")  # integer labels
```

---

## Checklist: Dot-Plot / Line+Marker Series

```python
set_series_line_style(series, visible=False)    # hide connecting line if dot-only
set_series_marker(series, "circle", size=12, fill_color=C_RYB_Q4)
set_data_label_color(series, C_RYB_Q4)
set_marker_data_label_pos(series, "r")          # labels to the right of each dot
```

For two-brand comparison (J&J orange vs AZ purple):
```python
# J&J series
set_series_marker(jj_series, "circle", size=12, fill_color=C_RYB_Q4)
set_data_label_color(jj_series, C_RYB_Q4)
# AZ series
set_series_marker(az_series, "circle", size=12, fill_color=C_TAG)
set_data_label_color(az_series, C_TAG)
# Industry average (gray diamonds)
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
set_series_line_style(series, width_pt=2.0, dash="solid")
set_series_smooth(series, True)                 # smooth curves (optional)
set_series_marker(series, "circle", size=6, fill_color=C_RYB_Q4)
set_marker_data_label_pos(series, "t")         # labels above each point
chart.has_title = False
```

---

## Checklist: Donut Chart

```python
set_pie_slice_colors(chart, [C_RYB_Q4, C_TAG, C_GREEN, C_LTGREY])
set_donut_hole_size(chart, pct=65)             # standard hole size
chart.has_legend = False
chart.has_title = False
```

---

## Checklist: Scatter / Quadrant Chart

```python
# Call BEFORE add_chart so rects sit behind
scatter_quadrant_fills(slide, chart_left, chart_top, chart_w, chart_h)
# Then add chart normally
cs = slide.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER,
                             Inches(chart_left), Inches(chart_top),
                             Inches(chart_w), Inches(chart_h), cd)
# Style each series
set_series_marker(series, "circle", size=10, fill_color=C_RYB_Q4)
set_series_line_style(series, visible=False)
hide_axis(chart, "val")                        # hide value axis
```

---

## Reference Line (Industry Average / Benchmark)

```python
add_val_axis_reference_line(chart, x_value=35, label="Industry Avg",
                             color=C_GREY, dash="dash", width_pt=1.0)
```

Adds a dashed vertical line at x=35 using a supplementary scatter series.

---

## MR+ME Two-Chart Layout

The standard ZoomRx two-chart slide layout:

```python
# Layout constants
CHART_TOP  = 1.47    # top of both charts
CHART_H    = 5.05    # height of both charts
MR_LEFT    = 0.20
MR_W       = 7.30
MR_DELTA_W = 0.62
ME_LEFT    = MR_LEFT + MR_W + MR_DELTA_W + 0.10   # = 8.26
ME_W       = 4.40
ME_DELTA_W = 0.62
DELTA_H    = CHART_H - 0.12

# MR chart -- shows full Y-axis category labels
mr_cs = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED,
    Inches(MR_LEFT), Inches(CHART_TOP), Inches(MR_W), Inches(CHART_H), mr_cd)
# Apply bar chart checklist to mr_cs.chart

# ME chart -- hides Y-axis (left chart already has labels)
me_cs = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED,
    Inches(ME_LEFT), Inches(CHART_TOP), Inches(ME_W), Inches(CHART_H), me_cd)
hide_cat_labels(me_cs.chart)   # suppress Y-axis labels on right chart

# Delta columns
add_delta_col(slide, mr_deltas, left=MR_LEFT+MR_W, top=CHART_TOP,
              width=MR_DELTA_W, height=DELTA_H, header="D Q4-Q3")
add_delta_col(slide, me_deltas, left=ME_LEFT+ME_W, top=CHART_TOP,
              width=ME_DELTA_W, height=DELTA_H, header="D")

# Legend
manual_legend(slide, q4_n=120, q3_n=108,
              chart_left=MR_LEFT, chart_right=ME_LEFT+ME_W, chart_bottom=CHART_TOP+CHART_H)
```

---

## Common Mistakes

| Mistake | Fix |
|---|---|
| Bars in wrong order (bottom-first) | Call `invert_cat_axis(chart)` |
| Data labels overlap bars | Use `set_datalabel_pos_outside_end(series)` |
| Bars too thin | Reduce `set_plot_area_gap(chart, gap_pct)` -- lower = fatter bars |
| ME chart shows duplicate labels | Call `hide_cat_labels(me_chart)` |
| Quadrant fill rects sit on top of scatter | Call `scatter_quadrant_fills()` BEFORE `add_chart` |
| Chart open in PowerPoint + python-pptx edit | Never mix -- use COM for open files, python-pptx for closed |
