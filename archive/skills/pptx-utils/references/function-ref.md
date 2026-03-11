# pptx_utils -- Full Function Reference

`from pptx_utils import *` in every slide script.

---

## Section 1 -- Brand Constants

| Name | Value | Use |
|---|---|---|
| `C_RED` | `#FF0000` | J&J red |
| `C_RYB_Q4` | `#F75824` | Q4 deep orange |
| `C_RYB_Q3` | `#FFC199` | Q3 pale orange |
| `C_TAG` | `#7030A0` | AZ / TAGRISSO purple |
| `C_GREEN` | `#00B050` | Positive delta / NPP |
| `C_WHITE` | `#FFFFFF` | White |
| `C_GREY` | `#505050` | Body text |
| `C_FTGREY` | `#7F7F7F` | Footer / faint text |
| `C_LBGREY` | `#F4F4F4` | Alternating table row bg |
| `C_HDRGREY` | `#404040` | Delta table header bg |
| `C_LTGREY` | `#BFBFBF` | Gridlines / borders |
| `COM_RED` | `0x0000FF` | COM BGR red |
| `COM_ORANGE` | `0x2458F7` | COM BGR orange |
| `COM_GREEN` | `0x50B000` | COM BGR green |
| `FONT_DISPLAY` | `"Johnson Display"` | Headlines |
| `FONT_TEXT` | `"Johnson Text"` | Body text |
| `SLIDE_W_IN` | `13.333` | Slide width in inches |
| `SLIDE_H_IN` | `7.500` | Slide height in inches |
| `IN` | `72` | 1 inch in COM points |
| `EMU_PER_IN` | `914400` | 1 inch in python-pptx EMU |

---

## Section 2 -- XML Helpers

Functions that write directly into OOXML because python-pptx has no high-level API.

### Original 7

| Function | Signature | What it does |
|---|---|---|
| `invert_cat_axis` | `(chart)` | First category at top of horizontal bar chart (sets orientation to maxMin) |
| `hide_cat_labels` | `(chart)` | Suppress Y-axis tick labels (tickLblPos=none) |
| `set_datalabel_pos_outside_end` | `(series)` | Data labels right of bar (dLblPos=outEnd) |
| `set_series_no_border` | `(series)` | Remove bar outline (noFill on spPr) |
| `set_val_axis_number_format` | `(axis, fmt="0")` | Axis number format: `"0"`, `"0%"`, `"0.0"` |
| `set_plot_area_gap` | `(chart, gap_pct=80)` | Gap between bar clusters as % of bar width |
| `set_overlap` | `(chart, overlap=0)` | Overlap within cluster; negative = gap between Q3/Q4 bars |

### New 12

| Function | Signature | What it does |
|---|---|---|
| `set_series_marker` | `(series, marker_type="circle", size=10, fill_color=None, line_color=None)` | Marker symbol, size, fill, border for line/scatter series. Use for dot-plot slides. |
| `set_series_line_style` | `(series, width_pt=1.5, dash="solid", visible=True)` | Line weight, dash style, or hide line. `dash`: `"solid"`, `"dash"`, `"dot"`, `"dashDot"`, `"lgDash"` |
| `set_series_smooth` | `(series, smooth=True)` | Smooth bezier curves vs straight segments (`c:smooth` element) |
| `set_marker_data_label_pos` | `(series, pos="r")` | Label position on line+marker: `"r"`, `"l"`, `"t"`, `"b"`, `"ctr"` |
| `set_data_label_color` | `(series, color_rgb)` | Override data label text color per series (orange J&J vs purple AZ) |
| `add_val_axis_reference_line` | `(chart, x_value, label="", color=None, dash="dash", width_pt=1.0)` | Vertical reference line at fixed X position (injected as scatter series) |
| `set_stacked_label_pos` | `(series, pos="ctr")` | Label pos inside stacked bar: `"inBase"`, `"inEnd"`, `"ctr"`, `"outEnd"` |
| `set_pie_slice_colors` | `(chart, colors)` | Fill each pie/donut slice by index: `colors=[C_RED, C_TAG, C_GREEN, ...]` |
| `set_donut_hole_size` | `(chart, pct=50)` | Inner hole radius as % of diameter (10-90) |
| `hide_axis` | `(chart, axis="val")` | Fully suppress an axis (line+ticks+labels+gridlines). `axis`: `"val"` or `"cat"` |
| `set_gridlines` | `(chart, axis="val", major=True, minor=False)` | Enable/disable major or minor gridlines |
| `set_series_color` | `(series, fill_color, line_color=None)` | Set bar/line fill and optional border color |

---

## Section 3 -- Shape Builders

All positions in **inches**. Returns the shape object; assign `.name = "zrx_NNN"`.

### Original 7

| Function | Signature | What it does |
|---|---|---|
| `textbox` | `(slide, text, left, top, width, height, fsize=9, bold=False, color=None, align=PP_ALIGN.LEFT, italic=False, wrap=True, font=FONT_TEXT)` | Add a text box |
| `solidrect` | `(slide, left, top, width, height, fill, line=None)` | Filled rectangle |
| `horiz_line` | `(slide, left, top, width, color=None, width_pt=1.0)` | Solid horizontal connector line |
| `add_delta_col` | `(slide, deltas, left, top, width, height, header, hdr_h_frac=0.06)` | Q4-Q3 delta table with green/red coloring and alternating row shading |
| `slide_header` | `(slide, headline, module_label="Personal Promotion Module")` | Full ZoomRx slide header: accent line + headline + badge + separator |
| `slide_footer` | `(slide, footer_text)` | Standard small-text footnote at bottom of slide |
| `manual_legend` | `(slide, q4_n, q3_n, chart_left, chart_right, chart_bottom)` | Shared Q3/Q4/delta color legend below a pair of charts |

### New 11

| Function | Signature | What it does |
|---|---|---|
| `callout_box` | `(slide, left, top, width, height, text=None, border_color=None, dashed=True, fill_color=None, fsize=8, text_color=None)` | Rounded rect with dashed border and light tint fill. Most-used annotation shape. `border_color` defaults to `C_RED`; fill auto-computes 12% tint. |
| `section_header_bar` | `(slide, label, top=1.40, icon_path=None)` | Gray strip with optional icon and bold all-caps label. Appears on every data slide as chart title bar. |
| `module_badge` | `(slide, label, color=None)` | Colored pill badge top-right for module label (e.g. red "PERSONAL PROMOTION MODULE") |
| `section_breadcrumb` | `(slide, text)` | Small right-aligned breadcrumb text top-right (section context) |
| `divider_slide` | `(slide, title, logo_path=None)` | Full section divider: off-white bg, red left stripe, large red title, optional logo |
| `cover_slide` | `(slide, title, subtitle, date, client_name, jj_logo_path=None, zrx_logo_path=None)` | Full red-background title slide |
| `stat_callout` | `(slide, value, delta, label, left, top)` | Large stat: big red number, grey delta in parens, grey label. `delta` can be int or None. |
| `insert_image` | `(slide, img_path, left, top, width, height, name=None)` | Place PNG/JPG at inch coordinates. Returns None silently if file missing. |
| `dashed_separator` | `(slide, left, top, width, color=None, width_pt=0.75, dash="dash")` | Horizontal dashed line. `dash`: `"dash"`, `"dot"`, `"dashDot"` |
| `trend_arrow_icon` | `(slide, direction, left, top)` | Directional arrow icon. `direction`: `"up"` (green), `"down"` (red), `"flat"` (yellow) |
| `scatter_quadrant_fills` | `(slide, chart_left, chart_top, chart_width, chart_height, tl_color=None, tr_color=None, bl_color=None, br_color=None)` | Four colored background rects behind a scatter chart. Call BEFORE `add_chart` so rects sit behind. |

---

## Section 4 -- COM Helpers

For live editing of a file open in PowerPoint. All positions in **inches**. Colors in **BGR** int order.

| Function | Signature | What it does |
|---|---|---|
| `com_connect` | `(target_filename)` | Connect to running PowerPoint, return Presentation object |
| `com_find_shape` | `(com_slide, name)` | Find shape by `zrx_` name; raises RuntimeError if missing |
| `com_set_text` | `(shape, text, color_bgr=None, size_pt=None, bold=None)` | Set text + optional formatting |
| `com_set_fill` | `(shape, color_bgr)` | Set fill color (BGR int) |
| `com_move` | `(shape, left_in, top_in)` | Move shape to new position |
| `com_resize` | `(shape, width_in, height_in)` | Resize shape |
| `com_get_position` | `(shape)` | Returns `(left, top, width, height)` in inches |

**BGR color order:** `0x0000FF` = red, `0x2458F7` = orange, `0x50B000` = green

---

## Section 5 -- Registry Helpers

| Function | Signature | What it does |
|---|---|---|
| `load_registry` | `(path=None)` | Load `slide_registry.json` -> dict |
| `save_registry` | `(registry, path=None)` | Write registry dict to JSON |
| `registry_get` | `(name, path=None)` | Get single shape record; raises KeyError if not found |
| `registry_tag_slide` | `(slide_idx, module, section, data_source="", path=None)` | Store slide-level metadata (module, section, source) keyed by slide index |
| `registry_find_by_type` | `(shape_type, slide_idx=None, path=None)` | Query all shapes by type string: `"chart"`, `"textbox"`, `"rect"`, `"image"`. Returns `[(name, record), ...]` |
| `registry_diff_slide` | `(slide_idx, com_slide, path=None)` | Diff registry snapshot vs live COM state for one slide. Returns list of changed shapes. |

---

## Decision Guide

| I need to... | Call |
|---|---|
| Add annotation box with dashed border | `callout_box(slide, left, top, w, h, text="...", border_color=C_RYB_Q4)` |
| Add the gray chart-title strip | `section_header_bar(slide, "LABEL TEXT", top=1.40)` |
| Color large circle markers on dot-plot | `set_series_marker(series, "circle", size=12, fill_color=C_RYB_Q4)` |
| Color one bar series | `set_series_color(series, C_RYB_Q4)` |
| Make colored labels (orange J&J vs purple AZ) | `set_data_label_color(series, C_RYB_Q4)` |
| Add a dashed "industry average" line | `add_val_axis_reference_line(chart, 35)` |
| Suppress a chart axis completely | `hide_axis(chart, "val")` |
| Position stacked bar labels inside segments | `set_stacked_label_pos(series, "ctr")` |
| Add module label badge top-right | `module_badge(slide, "PERSONAL PROMOTION MODULE", color=C_RED)` |
| Add section context text top-right | `section_breadcrumb(slide, "Key Findings - Personal Promotions")` |
| Build a cover slide | `cover_slide(slide, title, subtitle, date, client_name)` |
| Build a section divider | `divider_slide(slide, "Section Title")` |
| Show a big stat (e.g. "16 (-1) years") | `stat_callout(slide, "16", -1, "years avg treatment duration", left, top)` |
| Place an image/icon | `insert_image(slide, path, left, top, width, height, name="zrx_NNN")` |
| Add colored quadrant backgrounds to scatter | `scatter_quadrant_fills(slide, chart_left, chart_top, chart_w, chart_h)` |
| Add a horizontal divider line (solid) | `horiz_line(slide, left, top, width)` |
| Add a horizontal divider line (dashed) | `dashed_separator(slide, left, top, width)` |
| Add a trend direction arrow | `trend_arrow_icon(slide, "up", left, top)` |
| Tag a slide for rebuild tracking | `registry_tag_slide(slide_idx, module, section, data_source)` |
| Find all charts on a slide | `registry_find_by_type("chart", slide_idx=2)` |
| Check what moved since last build | `registry_diff_slide(slide_idx, com_slide)` |
