# Chart Type Selection Guide

Use this to choose the right chart type given the data and communication goal.

---

## Decision Tree

```
What am I showing?

Ranking / comparison of items at one point in time?
  |-- Are items compared between two brands (J&J vs AZ)?
  |     YES --> Clustered horizontal bar (BAR_CLUSTERED)
  |             Two bars per row: Q4 orange + Q3 pale orange
  |             With delta column right of chart
  |
  |-- Is it attribute RATINGS (0-100 scale, not % recall)?
  |     YES --> Dot-plot (LINE or XY_SCATTER with markers, line hidden)
  |             Large circles per brand per row
  |
  |-- Is it 100% composition (shares sum to 100%)?
         YES --> 100% stacked horizontal bar (BAR_STACKED, overlap=100)
                 Labels inside each segment: set_stacked_label_pos(series, "ctr")

Change over time (monthly, quarterly)?
  |-- Multiple messages / attributes tracked over months?
  |     YES --> Multi-line chart (LINE)
  |             One line per message, smooth curves optional
  |             Dots at key endpoints, value labels right-of-line
  |
  |-- Single metric trend (one line)?
        YES --> Line chart (LINE) or simple bar (COLUMN)

Two-dimensional positioning (importance vs performance)?
  --> Scatter chart (XY_SCATTER) with quadrant fills
      scatter_quadrant_fills() BEFORE add_chart

Composition / share breakdown (pie-like)?
  --> Donut chart (DOUGHNUT) with hole size 60-70%
      set_pie_slice_colors() to override slice fills

Single large statistic?
  --> stat_callout() -- not a chart, just styled textboxes
```

---

## Chart Type Reference

| Chart type | python-pptx enum | Typical use in ZoomRx decks |
|---|---|---|
| Clustered horizontal bar | `XL_CHART_TYPE.BAR_CLUSTERED` | MR/ME message recall and effect (core deliverable) |
| Stacked horizontal bar | `XL_CHART_TYPE.BAR_STACKED` | Category composition (reach by channel, SOV) |
| 100% stacked horizontal bar | `XL_CHART_TYPE.BAR_STACKED_100` | Share breakdown where total = 100% |
| Dot-plot (line, markers only) | `XL_CHART_TYPE.LINE` | Attribute ratings, rep performance abacus |
| Scatter | `XL_CHART_TYPE.XY_SCATTER` | Importance vs performance quadrant maps; also used for reference lines |
| Line chart | `XL_CHART_TYPE.LINE` | Time trends (R3M reach, frequency, SOV) |
| Donut | `XL_CHART_TYPE.DOUGHNUT` | Share of interactions, interaction mix |

---

## Series Color Assignment

Apply in this priority order:

1. **J&J / primary brand** -- `C_RYB_Q4` (#F75824 deep orange) for Q4; `C_RYB_Q3` (#FFC199 pale) for Q3
2. **AstraZeneca / TAGRISSO** -- `C_TAG` (#7030A0 purple)
3. **Industry average / benchmark** -- `C_GREY` (#505050) with diamond marker or dashed line
4. **Positive signal** -- `C_GREEN` (#00B050) for NPP module or positive callouts
5. **TAGRISSO + Chemo combo** -- use a distinct blue or teal if needed (define ad-hoc)

Always use `set_series_color(series, color)` -- never hardcode colors in chart XML directly.

---

## Axis Formatting Rules

| Axis | Rule |
|---|---|
| Category (Y) axis on MR chart | Show labels (left chart in two-chart layout) |
| Category (Y) axis on ME chart | Suppress with `hide_cat_labels(chart)` |
| Category (Y) axis on dot-plot right panel | Suppress with `hide_cat_labels(chart)` |
| Value (X) axis | Usually hidden: `hide_axis(chart, "val")` on dot-plots |
| Value (X) axis on bar charts | Keep visible; set number format `"0"` for integers, `"0%"` for percentages |
| Gridlines | Off by default; enabled only on specific slides: `set_gridlines(chart, major=True)` |

---

## Data Label Rules

| Chart type | Data label style |
|---|---|
| Horizontal bar | Outside-end: `set_datalabel_pos_outside_end(series)` |
| Stacked bar | Inside-center: `set_stacked_label_pos(series, "ctr")` |
| Dot-plot (abacus) | Right of marker: `set_marker_data_label_pos(series, "r")` |
| Line chart | Right of endpoint: `set_marker_data_label_pos(series, "r")` |
| Donut/pie | Outside or none -- suppress if crowded |

Color data labels to match series:
- J&J data labels: `set_data_label_color(series, C_RYB_Q4)`
- AZ data labels: `set_data_label_color(series, C_TAG)`

---

## Reference Lines

Use `add_val_axis_reference_line(chart, x_value, ...)` for:
- Industry average benchmark (gray dashed, labeled "Industry Avg")
- Threshold lines (e.g. "35% target" or "Adequate recall threshold")
- Pre/post comparison markers

```python
add_val_axis_reference_line(chart, x_value=35,
                             label="Industry Avg",
                             color=C_GREY, dash="dash", width_pt=1.0)
```

Note: This injects a supplementary scatter series into the chart's plot area. It appears as a clean vertical dashed line on horizontal bar charts.
