# Chart Type Selection Guide

Use this to choose the right chart type given the data and communication goal.

---

## Decision Tree

```
What am I showing?

Ranking / comparison of items at one point in time?
  |-- Are items compared between two brands/groups?
  |     YES --> Clustered horizontal bar (BAR_CLUSTERED)
  |             Two bars per row: current + prior period
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
  |-- Multiple series tracked over periods (scorecard)?
  |     YES --> trended_scorecard: mini line chart grid
  |             Or trended_activity: side-by-side line + stacked column panels
  |
  |-- Single metric trend (one line)?
        YES --> Line chart (LINE) or simple bar (COLUMN)

Two-dimensional positioning (importance vs performance)?
  --> quadrant_scatter: scatter chart with quadrant fills
      Quadrant fills drawn BEFORE add_chart

Composition / share breakdown (pie-like)?
  --> Donut chart (DOUGHNUT) with hole size 60-70%
      set_pie_slice_colors() to override slice fills

Multi-dimensional matrix with values per cell?
  --> heatmap_table: table with gradient fills + delta columns
      No chart — uses pptx tables with per-cell color interpolation

Single large statistic?
  --> stat_callout() -- not a chart, just styled textboxes
```

---

## Chart Type Reference

| Chart type | python-pptx enum | Typical use |
|---|---|---|
| Clustered horizontal bar | `XL_CHART_TYPE.BAR_CLUSTERED` | Message recall, effectiveness (core deliverable) |
| Stacked horizontal bar | `XL_CHART_TYPE.BAR_STACKED` | Category composition (reach by channel, SOV) |
| 100% stacked horizontal bar | `XL_CHART_TYPE.BAR_STACKED_100` | Share breakdown where total = 100% |
| Dot-plot (line, markers only) | `XL_CHART_TYPE.LINE` | Attribute ratings, rep performance abacus |
| Scatter | `XL_CHART_TYPE.XY_SCATTER` | Importance vs performance quadrant maps |
| Line chart | `XL_CHART_TYPE.LINE` | Time trends (reach, frequency, SOV) |
| Stacked column | `XL_CHART_TYPE.COLUMN_STACKED` | Share of voice vertical bars |
| Donut | `XL_CHART_TYPE.DOUGHNUT` | Share of interactions, interaction mix |
| Heatmap table | N/A (pptx table with gradient fills) | Message recall by channel, cross-tabulations |

---

## Series Color Assignment

Apply in this priority order:

1. **Primary brand** — `config.primary.color_current` for current period; `config.primary.color_prior` for prior
2. **Competitor brand** — `config.competitor.color_current`
3. **Benchmark / average** — `C_GREY` with diamond marker or dashed line
4. **Positive signal** — `C_GREEN` for positive callouts
5. **Additional series** — define ad-hoc colors in ask.extra

Always use `set_series_color(series, color)` — never hardcode colors in chart XML directly. Colors should come from the project's YAML config `brands:` section.

---

## Axis Formatting Rules

| Axis | Rule |
|---|---|
| Category (Y) axis on left chart | Show labels |
| Category (Y) axis on right chart | Suppress with `hide_cat_labels(chart)` |
| Value (X) axis on dot-plots | Usually hidden: `hide_axis(chart, "val")` |
| Value (X) axis on bar charts | Keep visible; set number format `"0"` or `"0%"` |
| Both axes on line charts | Usually hidden: `hide_axes=True` in `add_line_chart()` |
| Gridlines | Off by default; enable selectively: `set_gridlines(chart, major=True)` |

---

## Data Label Rules

| Chart type | Data label style |
|---|---|
| Horizontal bar | Outside-end: `set_datalabel_pos_outside_end(series)` |
| Stacked bar | Inside-center: `set_stacked_label_pos(series, "ctr")` |
| Dot-plot (abacus) | Right of marker: `set_marker_data_label_pos(series, "r")` |
| Line chart | First series above, rest below: `label_positions=["t", "b"]` |
| Stacked column | Center: `pos="ctr"` with white text |
| Donut/pie | Outside or none — suppress if crowded |

Color data labels to match series:
- `set_data_label_color(series, series_color)`

---

## Reference Lines

Use `add_val_axis_reference_line(chart, x_value, ...)` for:
- Benchmark lines (gray dashed)
- Threshold lines (e.g. "35% target")

```python
add_val_axis_reference_line(chart, x_value=35,
                             label="Benchmark",
                             color=C_GREY, dash="dash", width_pt=1.0)
```

Note: This injects a supplementary scatter series into the chart's plot area.

---

## Renderer-Data Compatibility Rules

Each slide_type requires specific data field names. Using the wrong data format produces empty charts.

| slide_type | Required data fields | Common mistake |
|---|---|---|
| `single_bar_with_delta` | `desc`, `prior`, `current` (+ optional `short`) | Works with any standard extraction — safest default |
| `abacus` | `desc` (or `short`), `prior`, `current` | Same as single_bar — works with standard data |
| `dual_bar_with_delta` | `{prefix}_current`, `{prefix}_prior` per side | **Fails with simple `prior`/`current`** — needs prefixed fields |
| `clustered_compare` | `primary_current`, `comp_current` (or `extra.primary_key` + `extra.comp_key`) | **Fails with simple `prior`/`current`** — needs two-brand merged data |
| `hii_scorecard` | Needs `extra.sections` defining layout | **Empty without `extra` config** |
| `heatmap_table` | Needs `extra.columns` defining column layout | **Empty without `extra` config** |
| `executive_summary` | `ask.extra.insights` list OR `ask.source_text` path to .md | Fallback reads bullet lines from markdown |

### Safe Defaults for simple `{desc, prior, current}` data:
- **`single_bar_with_delta`** for ranked lists, comparisons, breakdowns
- **`abacus`** for attribute ratings (scatter dot plot with value tables)
- **`executive_summary`** for text-only insights (populate `extra.insights`)

### Label Shortcuts
Long labels truncated at LABEL_MAX (65 chars). Fix: add `label_shortcuts` config with `keywords` -> `short` mappings, set `use_label_shortcuts: true` in extraction params.

### pct_mode
- `pct` (default): 0-1 decimals -> multiply by 100 (0.46 -> 46%)
- `straight`: already whole-number percentages (46 -> 46%)
- Auto-detected from `_codes.value_range`

### Template Rules
- NEVER use a prior wave report PPTX as template — clearing slides with embedded charts corrupts the file
- Always use a clean blank template created by python-pptx with correct slide dimensions

### Multi-Column Field Naming Convention

When using `question_code_multi_col`, the `columns` dict keys become the field names in the extracted data. The renderer expects the pattern `{field}_current` and `{field}_prior`. **Always use this convention:**

```yaml
# CORRECT — renderer finds hii_current, hii_prior, others_current, others_prior
columns:
  prior: 7
  current: 17
  hii_prior: 10
  hii_current: 20
  others_prior: 9
  others_current: 19
  comm_prior: 14
  comm_current: 24
  acad_prior: 15
  acad_current: 25

# WRONG — renderer looks for hii_current but finds current_hii (mismatch)
columns:
  prior_hii: 10
  current_hii: 20
  prior_hii_others: 9
  current_hii_others: 19
```

Then in `extra.series`:
```yaml
extra:
  series:
    - field: hii
      label: "High Impact"
      color: "#F75824"
    - field: others
      label: "Others"
      color: "#999999"
```

The renderer resolves: `r.get(f"{field}_current")` → `r.get("hii_current")` → value.
