# pptx_utils -- Full Function Reference

`from slidegen.pptx_utils import *` in every slide script.

---

## Section 1 -- Constants

| Name | Value | Use |
|---|---|---|
| `C_RED` | `#FF0000` | Accent red |
| `C_GREEN` | `#00B050` | Positive delta |
| `C_WHITE` | `#FFFFFF` | White |
| `C_GREY` | `#505050` | Body text |
| `C_FTGREY` | `#7F7F7F` | Footer / faint text |
| `C_LBGREY` | `#F4F4F4` | Alternating table row bg |
| `C_HDRGREY` | `#404040` | Delta table header bg |
| `C_LTGREY` | `#BFBFBF` | Gridlines / borders |
| `FONT_DISPLAY` | `"Johnson Display"` | Default display font (override per-project via config) |
| `FONT_TEXT` | `"Johnson Text"` | Default body font (override per-project via config) |
| `SLIDE_W_IN` | `13.333` | Slide width in inches |
| `SLIDE_H_IN` | `7.500` | Slide height in inches |
| `IN` | `72` | 1 inch in COM points |
| `EMU_PER_IN` | `914400` | 1 inch in python-pptx EMU |

Brand colors (`C_RYB_Q4`, `C_RYB_Q3`, `C_TAG`, etc.) are defined in `brand.py` but should be accessed via `config.primary.color_current`, `config.competitor.color_current` etc. from the project YAML.

---

## Section 2 -- XML Helpers

Functions that write directly into OOXML because python-pptx has no high-level API.

| Function | Signature | What it does |
|---|---|---|
| `invert_cat_axis` | `(chart)` | First category at top of horizontal bar chart |
| `hide_cat_labels` | `(chart)` | Suppress Y-axis tick labels |
| `set_datalabel_pos_outside_end` | `(series)` | Data labels right of bar |
| `set_series_no_border` | `(series)` | Remove bar outline |
| `set_val_axis_number_format` | `(axis, fmt="0")` | Axis number format: `"0"`, `"0%"`, `"0.0"` |
| `set_plot_area_gap` | `(chart, gap_pct=80)` | Gap between bar clusters as % of bar width |
| `set_overlap` | `(chart, overlap=0)` | Overlap within cluster; negative = gap between bars |
| `set_series_marker` | `(series, marker_type="circle", size=10, fill_color=None, line_color=None)` | Marker symbol, size, fill, border for line/scatter series |
| `set_series_line_style` | `(series, width_pt=1.5, dash="solid", visible=True)` | Line weight, dash style, or hide line |
| `set_series_smooth` | `(series, smooth=True)` | Smooth bezier curves vs straight segments |
| `set_marker_data_label_pos` | `(series, pos="r")` | Label position on line+marker: `"r"`, `"l"`, `"t"`, `"b"`, `"ctr"` |
| `set_data_label_color` | `(series, color_rgb)` | Override data label text color per series |
| `add_val_axis_reference_line` | `(chart, x_value, label="", color=None, dash="dash", width_pt=1.0)` | Vertical reference line at fixed X position |
| `set_stacked_label_pos` | `(series, pos="ctr")` | Label pos inside stacked bar: `"inBase"`, `"inEnd"`, `"ctr"`, `"outEnd"` |
| `set_pie_slice_colors` | `(chart, colors)` | Fill each pie/donut slice by index |
| `set_donut_hole_size` | `(chart, pct=50)` | Inner hole radius as % of diameter |
| `hide_axis` | `(chart, axis="val")` | Fully suppress an axis (line+ticks+labels+gridlines) |
| `set_gridlines` | `(chart, axis="val", major=True, minor=False)` | Enable/disable gridlines |
| `set_series_color` | `(series, fill_color, line_color=None)` | Set bar/line fill and optional border color |
| `set_chart_plot_area` | `(chart, x=0.0, y=0.0, w=1.0, h=1.0)` | Set chart plot area position (fractions of chart frame) |
| `set_val_axis_scale` | `(chart, min_val=0, max_val=100)` | Set explicit min/max scale on the value axis |

---

## Section 2b -- High-Level Chart Builders

| Function | Signature | What it does |
|---|---|---|
| `add_single_bar_chart` | `(slide, categories, values, left, top, width, height, fill_color, ...)` | Single-series horizontal bar chart with data labels. Returns (chart_frame, chart). |
| `add_clustered_bar_chart` | `(slide, categories, series_list, left, top, width, height, colors, ...)` | Multi-series clustered horizontal bar chart. `series_list`: [("Name", [vals...]), ...] |
| `add_line_chart` | `(slide, categories, series_list, left, top, width, height, colors, marker_size=5, line_width_pt=2.25, label_positions=None, ...)` | Line chart with circle markers. `label_positions`: per-series list (e.g. ["t", "b"]) — defaults to first above, rest below. |
| `add_stacked_column_chart` | `(slide, categories, series_list, left, top, width, height, colors, ...)` | Stacked vertical column chart with centered data labels. |
| `add_scatter_chart` | `(slide, series_list, left, top, width, height, colors, marker_size=7, x_min, x_max, y_min, y_max, ...)` | XY scatter chart, no connecting lines. `series_list`: [("Name", [x_vals], [y_vals]), ...] |

---

## Section 3 -- Shape Builders

All positions in **inches**. Returns the shape object; assign `.name = "zrx_NNN"`.

| Function | Signature | What it does |
|---|---|---|
| `textbox` | `(slide, text, left, top, width, height, fsize=9, bold=False, color=None, align=PP_ALIGN.LEFT, italic=False, wrap=True, font=FONT_TEXT)` | Add a text box |
| `solidrect` | `(slide, left, top, width, height, fill, line=None)` | Filled rectangle |
| `horiz_line` | `(slide, left, top, width, color=None, width_pt=1.0)` | Solid horizontal connector line |
| `add_delta_col` | `(slide, deltas, left, top, width, height, header, hdr_h_frac=0.06)` | Delta table with green/red coloring and alternating row shading |
| `slide_header` | `(slide, headline, font=None, module_label="")` | Slide header: accent line + headline |
| `slide_footer` | `(slide, footer_text, font=None)` | Standard small-text footnote at bottom |
| `manual_legend` | `(slide, ...)` | Shared legend below charts |
| `callout_box` | `(slide, left, top, width, height, text=None, border_color=None, dashed=True, fill_color=None, fsize=8, text_color=None)` | Rounded rect with dashed border and light tint fill |
| `section_header_bar` | `(slide, label, top=1.40, icon_path=None, font=None)` | Gray strip with optional icon and bold label |
| `module_badge` | `(slide, label, color=None)` | Colored pill badge top-right for module label |
| `section_breadcrumb` | `(slide, text)` | Small right-aligned breadcrumb text top-right |
| `divider_slide` | `(slide, title, logo_path=None)` | Full section divider slide |
| `cover_slide` | `(slide, title, subtitle, date, client_name, ...)` | Full cover/title slide |
| `stat_callout` | `(slide, value, delta, label, left, top)` | Large stat: big number, grey delta, grey label |
| `insert_image` | `(slide, img_path, left, top, width, height, name=None)` | Place PNG/JPG at inch coordinates |
| `dashed_separator` | `(slide, left, top, length, color=None, width_pt=0.75, dash="dash", vertical=False)` | Dashed line separator (horizontal or vertical) |
| `trend_arrow_icon` | `(slide, direction, left, top)` | Directional arrow: `"up"` (green), `"down"` (red), `"flat"` (yellow) |
| `scatter_quadrant_fills` | `(slide, chart_left, chart_top, chart_width, chart_height, tl_color=None, tr_color=None, bl_color=None, br_color=None)` | Four colored background rects behind a scatter chart |

---

## Section 4 -- COM Helpers

For live editing of a file open in PowerPoint. All positions in **inches**. Colors in **BGR** int order.

| Function | Signature | What it does |
|---|---|---|
| `com_connect` | `(target_filename)` | Connect to running PowerPoint, return Presentation object |
| `com_find_shape` | `(com_slide, name)` | Find shape by `zrx_` name |
| `com_set_text` | `(shape, text, color_bgr=None, size_pt=None, bold=None)` | Set text + optional formatting |
| `com_set_fill` | `(shape, color_bgr)` | Set fill color (BGR int) |
| `com_move` | `(shape, left_in, top_in)` | Move shape to new position |
| `com_resize` | `(shape, width_in, height_in)` | Resize shape |
| `com_get_position` | `(shape)` | Returns `(left, top, width, height)` in inches |

---

## Section 5 -- Registry Helpers

| Function | Signature | What it does |
|---|---|---|
| `load_registry` | `(path=None)` | Load `slide_registry.json` -> dict |
| `save_registry` | `(registry, path=None)` | Write registry dict to JSON |
| `registry_get` | `(name, path=None)` | Get single shape record |
| `registry_tag_slide` | `(slide_idx, module, section, data_source="", path=None)` | Store slide-level metadata |
| `registry_find_by_type` | `(shape_type, slide_idx=None, path=None)` | Query shapes by type: `"chart"`, `"textbox"`, `"rect"`, `"image"` |
| `registry_diff_slide` | `(slide_idx, com_slide, path=None)` | Diff registry vs live COM state for one slide |

---

## Decision Guide

| I need to... | Call |
|---|---|
| Add annotation box with dashed border | `callout_box(slide, left, top, w, h, text="...", border_color=accent_color)` |
| Add the gray chart-title strip | `section_header_bar(slide, "LABEL TEXT", top=1.40)` |
| Color large circle markers on dot-plot | `set_series_marker(series, "circle", size=12, fill_color=color)` |
| Color one bar series | `set_series_color(series, color)` |
| Make colored labels matching series | `set_data_label_color(series, color)` |
| Add a dashed benchmark line | `add_val_axis_reference_line(chart, 35)` |
| Suppress a chart axis completely | `hide_axis(chart, "val")` |
| Position stacked bar labels inside segments | `set_stacked_label_pos(series, "ctr")` |
| Add module label badge top-right | `module_badge(slide, "MODULE NAME", color=accent_color)` |
| Add section context text top-right | `section_breadcrumb(slide, "Key Findings - Section")` |
| Build a cover slide | `cover_slide(slide, title, subtitle, date, client_name)` |
| Build a section divider | `divider_slide(slide, "Section Title")` |
| Show a big stat | `stat_callout(slide, "16", -1, "metric label", left, top)` |
| Place an image/icon | `insert_image(slide, path, left, top, width, height)` |
| Add colored quadrant backgrounds | `solidrect()` × 4 (before add_chart for z-order)` |
| Add a line chart (trend) | `add_line_chart(slide, categories, series_list, left, top, w, h, colors)` |
| Add a stacked column chart | `add_stacked_column_chart(slide, categories, series_list, left, top, w, h, colors)` |
| Add a scatter chart (quadrant) | `add_scatter_chart(slide, series_list, left, top, w, h, colors, x_min, x_max, y_min, y_max)` |
| Add a horizontal divider line (dashed) | `dashed_separator(slide, left, top, length)` |
| Add a vertical divider line (dashed) | `dashed_separator(slide, left, top, length, vertical=True)` |
| Add a trend direction arrow | `trend_arrow_icon(slide, "up", left, top)` |
| Tag a slide for rebuild tracking | `registry_tag_slide(slide_idx, module, section, data_source)` |
| Find all charts on a slide | `registry_find_by_type("chart", slide_idx=2)` |
| Check what moved since last build | `registry_diff_slide(slide_idx, com_slide)` |