---
name: slidegen
description: "Use when working with the SlideGen pipeline — creating, editing, or regenerating YAML-driven PowerPoint decks. Trigger when: creating slides for a project, editing slide N, switching wave data, adding/removing slides, writing renderers or extractors, using pptx_utils functions, working with the shape registry, choosing chart types, applying brand colors, or designing slide layouts. This is the primary skill for all SlideGen and pptx_utils work."
---

# SlideGen Skill

## Cardinal Rules

1. **Never write raw lxml** — call `pptx_utils` functions from the package modules
2. **Never hardcode brand colors** — use `BRAND{}` dict keyed by client (e.g. `BRAND["jnj"]`, `BRAND["default"]`)
3. **Always use `pptx_utils/`** for shape creation — `textbox`, `solidrect`, `horiz_line`, etc.
4. **Always name shapes** with `zrx_` prefix via `ShapeNamer`
5. **Always backup config** before modifications — `_backup_config(yaml_path)`
6. **PPTX backup** happens automatically on `regenerate_slide()`
7. **Import from `slidegen.pptx_utils`** — the package `__init__.py` re-exports everything

---

## Standard Workflows

### 1. Create Slides

```
User says: "Create slides for projects/{name}"
```

1. Verify folder structure: `projects/{name}/data/{wave}/`, `reference/{wave}/`, `templates/`
2. Discover data: `discover_excel_structure()` on the source Excel
3. Read reference docs: parse `reference/{wave}/asks.md`
4. Read template: `python -m markitdown template.pptx`
5. Generate config scaffold: `generate_config_scaffold()`
6. Map asks to pipeline — match to existing extraction methods and slide types
7. Write `config.yaml` with all extractions and asks
8. Generate deck: `generate_deck("projects/{name}/config.yaml")`
9. Visual QA: convert to images, inspect, fix, re-run

### 2. Edit Slide N

```
User says: "Edit Slide 5 — change headline to ..."
```

1. Read config, identify ask at index N-1
2. Backup config: `_backup_config(yaml_path)`
3. Update config fields (headline, data_key, sort_by, etc.)
4. Regenerate slide: `regenerate_slide(yaml_path, slide_index=N-1)`
   - PPTX backup created automatically in `output/{wave}/backups/`
5. Tell user to reopen/refresh the deck
6. Report what changed

### 3. Edit Slides with New Wave Data

```
User says: "Edit slides with new wave data — PET_Q1Q2_2026"
```

1. Backup config: `_backup_config(yaml_path)`
2. Update `project.wave`, `period_current`, `period_prior` in config.yaml
3. Verify `data/{wave_id}/source_data.xlsx` exists
4. Regenerate deck: `generate_deck(yaml_path)`
   - Output goes to `output/{wave_id}/deck.pptx` (old wave preserved)
5. Report: old wave output preserved, new wave output in its own folder

### 4. Add/Remove Slides

```
User says: "Add a slide after Slide 6" / "Remove Slide 8"
```

1. Backup config
2. Add/remove/reorder asks in config.yaml
3. Full regen: `generate_deck(yaml_path)`
4. Report new slide count

### 5. Edit Slides with New Wave + New Asks

```
User says: "Edit slides with new wave data + new asks — PET_Q1Q2_2026"
```

1. Backup config
2. Update wave settings
3. Read new reference docs
4. Update extractions and asks in config.yaml
5. Gap analysis: check if new asks need new extractors or renderers
6. Full regen: `generate_deck(yaml_path)`

---

## Decision Rules

| Question | Answer |
|----------|--------|
| Text/sort/data change on one slide? | `regenerate_slide()` |
| New wave, add/remove slides, or structural change? | `generate_deck()` |
| New data shape not fitting existing extractors? | Add new method to `data_loaders.py` |
| New visualization not fitting existing renderers? | Add new renderer to `slide_renderers.py` |
| Live COM edit on open PowerPoint? | Use `LiveEditor` from `slidegen.edit` |

---

## Pipeline Architecture

```
config.yaml  ->  ProjectConfig (dataclasses)
                      |
Excel file  ->  data_loaders.load_all_data()  ->  dict[extraction_id -> list[dict]]
                      |
orchestrator  ->  RENDERERS[slide_type](slide, config, ask, data, namer)
                      |
                 output/{wave}/deck.pptx
                 output/{wave}/slide_data.json      (data cache)
                 output/{wave}/shape_registry.json   (shape state)
                 output/{wave}/backups/              (PPTX backups)
```

---

## pptx_utils/ Package Reference

| Module | Key Exports |
|--------|-------------|
| `brand` | `BRAND{}` (client-keyed), color aliases, font constants, `SLIDE_W_IN`, `SLIDE_H_IN` |
| `lxml_helpers` | `_get_or_add`, `invert_cat_axis`, `hide_cat_labels`, `set_series_color`, `set_plot_area_gap`, `set_overlap`, `hide_axis`, `set_data_label_color` |
| `shapes` | `textbox`, `solidrect`, `horiz_line`, `insert_image`, `dashed_separator`, `stat_callout`, `callout_box` |
| `layout` | `LAYOUTS{}`, `slide_header`, `slide_footer`, `section_header_bar`, `cover_slide`, `divider_slide`, `manual_legend`, `module_badge` |
| `charts` | `CHART_PATTERNS{}`, `enable_data_labels`, `delete_data_label`, `add_single_bar_chart`, `add_clustered_bar_chart` |
| `tables` | `add_delta_col`, `add_delta_table`, `add_value_table` |
| `com` | `com_connect`, `com_find_shape`, `com_set_text`, `com_set_fill`, `com_move`, `com_resize` |
| `registry` | `load_registry`, `save_registry`, `registry_get`, `registry_tag_slide`, `registry_find_by_type`, `registry_diff_slide` |

Import everything via: `from slidegen.pptx_utils import textbox, BRAND, LAYOUTS, CHART_PATTERNS`

---

## Data Extraction Methods

| Method | Use Case |
|--------|----------|
| `question_code` | Find a code row, walk sub-rows, extract prior/current values |
| `multi_question_code` | One row per code (e.g. CTA metrics) |
| `row_range` | Fixed row range with column mapping |
| `question_code_multi_col` | Multiple columns per row (e.g. HII: hi vs other) |
| `nested_ordinal` | Grouped ordinal sub-rows (e.g. 1st/2nd/3rd recall order) |

---

## Slide Type Renderers

| Type | Pattern |
|------|---------|
| `cover` | Title slide |
| `executive_summary` | Bullet-list insights |
| `single_bar_with_delta` | Table-based horizontal bar + QoQ delta column |
| `dual_bar_with_delta` | Two side-by-side bars + deltas (alternate row backgrounds) |
| `dual_bar_qoq` | Two side-by-side Q4-vs-Q3 clustered bars + deltas |
| `clustered_compare` | Table-based clustered bar comparing two groups + gap/delta columns |
| `dual_bar_compare` | Side-by-side dual brand bar comparison + insight callout |
| `qoq_bar_with_delta` | Q4 vs Q3 clustered + delta |
| `two_section_bar` | Two vertically stacked bar sections |
| `stacked_order` | Table-based stacked bar with ordinal breakdown + total column |
| `abacus` | XY scatter abacus with label/value tables + delta column |
| `dual_abacus` | Two side-by-side abacus panels (e.g. Acad vs Comm by brand) |
| `followup_rep` | Template slide 51-style follow-up rep abacus with dual delta columns |
| `hii_scorecard` | Multi-section clustered column chart with section headers + callouts |
| `dual_doughnut` | Side-by-side doughnut pairs comparing patient segments by brand |
| `message_mbd` | Multi-column abacus for Motivation/Believability/Differentiation |

### slide_type → extra Fields

Each `slide_type` expects specific `extra` fields in the ask config:

| slide_type | extra fields | Description |
|------------|-------------|-------------|
| `cover` | `subtitle`, `date`, `client` | Cover slide metadata |
| `executive_summary` | `insights: [str, ...]` | List of bullet-point insights |
| `single_bar_with_delta` | _(none required)_ | Uses `data_key` directly; table-based layout with dynamic label width |
| `dual_bar_with_delta` | `left: {field_prefix, label, delta_header}`, `right: {field_prefix, label, delta_header}`, `callouts: [{text, color}]` | Two side-by-side charts with optional callout boxes |
| `dual_bar_qoq` | `left: {field_prefix, label, delta_header}`, `right: {field_prefix, label, delta_header}` | Two clustered Q4-vs-Q3 charts |
| `clustered_compare` | `series: [{field, label, color}, ...]`, `gap_header` | Table-based two-group comparison with dynamic label width |
| `dual_bar_compare` | `left/right: {data_key, brand, label, delta_header}`, `insight_text`, `highlight_rows` | Side-by-side brand bars with callout |
| `qoq_bar_with_delta` | _(none required)_ | Uses `data_key` directly |
| `two_section_bar` | `top: {data_key, label, brand}`, `bottom: {data_key, label, brand}` | Stacked sections |
| `stacked_order` | `ordinals: ["1st", "2nd", "3rd", "4th"]` | Table-based ordinal breakdown with dynamic label width |
| `abacus` | `scale_min/max`, `scale_ticks`, `current_field/prior_field`, `color_current/prior`, `legend_current/prior`, `hide_val_cols`, `show_data_labels` | XY scatter with customizable fields, colors, and legend |
| `dual_abacus` | `left/right: {data_key, current_field, prior_field, color_current/prior, label}`, `hide_val_cols` | Two side-by-side abacus panels |
| `followup_rep` | `current_field/prior_field`, `delta_current_field/delta_prior_field`, `color_current/prior`, `legend_current/prior` | Template-matched follow-up rep layout (slide 51) |
| `hii_scorecard` | `sections: [{label, subtitle, summary, items}]`, `series: [{field, label, color}]`, `insight_text` | Multi-section clustered column scorecard |
| `dual_doughnut` | `left/right: {label, items: [{brand_label, current, prior, color, color_prior, sample_current/prior}]}` | QoQ doughnut rings per brand per segment |
| `message_mbd` | `series: [{field, label, color}]`, `ce_field`, `ce_prior_field`, `ce_header`, `scale_label` | MBD dot chart with composite effectiveness |

---

## Slide Archetypes

| Archetype | When to use | Key functions |
|---|---|---|
| Cover | First slide of deck | `cover_slide()` |
| Section divider | Between major sections | `divider_slide()` |
| Standard bar chart | MR/ME reach/preference data | `slide_header`, `add_delta_col`, `manual_legend` |
| Dot-plot / abacus | Message-level attribute ratings | `set_series_marker`, `hide_axis` |
| Line / trend chart | Time-series, R3M rolling | `set_series_line_style`, `set_series_smooth` |
| Scorecard table | Multi-metric summary | `add_delta_col`, `trend_arrow_icon` |
| Callout / insight | Exec summary, qualitative findings | `callout_box`, `textbox` |
| Scatter / quadrant | Importance vs performance | `scatter_quadrant_fills` |
| Donut / pie | Share-of-wallet, interaction mix | `set_pie_slice_colors`, `set_donut_hole_size` |

For full layout specs and implementation sequences, see `references/slide-archetypes.md`.
For chart selection guidance, see `references/chart-types.md`.
For brand constants and spacing, see `references/brand-constants.md`.
For function signatures, see `references/function-ref.md`.
For chart pattern checklists, see `references/chart-patterns.md`.

---

## Universal Slide Elements

Every data slide includes:
1. `slide_header(slide, headline)` — accent line + headline + badge + separator
2. `section_header_bar(slide, label, top=1.40)` — gray strip with chart title
3. `slide_footer(slide, footer_text)` — source footnote at bottom
4. **Speaker notes** — auto-generated with question codes and question text used to create the slide

## Table-Based Layout Features

Bar chart renderers (`single_bar_with_delta`, `clustered_compare`, `stacked_order`) use a table+chart layout:
- **Label table** (left): Full message text at 7.5pt with word wrap and alternating grey/white rows
- **Bar chart** (center): Hidden category labels, `inEnd` white data labels
- **Delta column** (right): Consistent alternating rows matching the label table (grey first)
- **Dynamic label width**: `_auto_label_width(labels)` computes optimal width (2.50"–5.50") based on longest label; chart width fills remaining space
- **Insight callout**: `dual_bar_compare` and `hii_scorecard` support `insight_text` displayed in a dashed-border callout box below the legend

## Abacus/Scatter Features

- **Data labels**: Both prior and current series show percentage labels above dots
- **Dynamic y-offset**: Labels scale with row count to avoid overlapping previous row gridlines
- **hide_val_cols**: Set `true` to remove value columns when data labels show the values; chart widens to fill freed space
- **Custom field mapping**: `current_field`/`prior_field` remap data fields (e.g. `primary_current`/`comp_current` for brand comparison)
- **Custom colors/legend**: Override brand colors and legend labels via `extra` config

---

## Layout & Spacing Rules

1. **Minimize white space** — Charts must fill the available vertical space between the section bar (Y≈1.85) and footer (Y≈6.78). Use `MIN_CHART_HEIGHT` (3.0") so slides with few data rows don't leave large empty areas.
2. **Chart height formula**: `chart_h = max(MIN_CHART_HEIGHT, min(MAX_*_HEIGHT, n * per_row))` — always clamp with both a floor and ceiling.
3. **Same scale for side-by-side charts** — When two charts are rendered next to each other (e.g. dual_bar_with_delta), set the same `set_val_axis_scale(ch, 0, axis_max)` on both so equal percentages produce equal bar widths.
4. **Suppress auto chart titles** — Always set `ch.has_title = False` after creating a chart. PowerPoint auto-generates a title from the series name (`autoTitleDeleted="0"`) which overlaps bar content.
5. **Delta table alignment** — The delta table header height (0.28") roughly matches the chart's auto top-padding, keeping data rows aligned with bars. Use `row_height = (chart_h - HEADER_ROW_HEIGHT_IN) / n` so total delta height equals chart height.
6. **Footer avoids master logo** — Footer text starts at `x=2.70` (after the J&J logo which occupies x=0.32–2.61 on the slide master).
7. **No unnecessary decorative lines** — The template's slide master provides all chrome. Don't add accent lines or separators that aren't in the template.

---

## Positioning Grid

```
Y=0.00  Top of slide
Y=0.05  Module label (small grey, right-aligned)
Y=0.16  Headline (bold red, left-aligned)
Y=0.91  Module badge (red pill, top-right)
Y=1.40  section_header_bar / chart_header_row
Y=1.68  Chart top (dual bar, below header row)
Y=1.85  Chart top (standard, below section bar)
Y=6.78  slide_footer (starts at x=2.70 to avoid logo)
Y=7.50  Bottom of slide (SLIDE_H_IN)

X=0.20  Left margin
X=0.30  Chart left edge
X=2.70  Footer text left edge (after master logo)
X=13.33 Right of slide (SLIDE_W_IN)
```

---

## Brand System (BRAND{} dict)

Colors and fonts are defined per-client in `pptx_utils/brand.py`:

```python
from slidegen.pptx_utils import BRAND

# Access by client key
colors = BRAND["jnj"]       # or BRAND["default"] for new clients
primary = colors["primary_current"]
font    = colors["font_display"]
```

Each client entry contains:
| Key | Purpose |
|-----|---------|
| `primary_current` | Current period bars (primary brand) |
| `primary_prior` | Prior period bars (primary brand) |
| `competitor` | Competitor brand bars |
| `competitor_prior` | Competitor prior period bars |
| `accent` | Title bar, headline, badge |
| `positive` / `negative` | Delta colors (green/red) |
| `font_display` / `font_body` | Header and body fonts |
| `template_path` | Default template for this client |

Legacy color aliases (`C_RED`, `C_GREEN`, `C_GREY`, etc.) are available for backward compatibility but new code should prefer `BRAND{}`.

### Structural Constants

| Name | Value | Use |
|------|-------|-----|
| `SLIDE_W_IN` | `13.333` | Slide width (inches) |
| `SLIDE_H_IN` | `7.500` | Slide height (inches) |
| `C_GREEN` | `#00B050` | Positive delta |
| `C_GREY` | `#505050` | Body text |
| `C_FTGREY` | `#7F7F7F` | Footer / faint text |
| `C_LBGREY` | `#F4F4F4` | Alternating table row bg |
| `C_HDRGREY` | `#404040` | Table header bg |

---

## Shape Registry

Shape registry (`shape_registry.json`) tracks all shapes per slide with data lineage:

```json
{
  "slides": {
    "4": {
      "ask_id": "brand_mr",
      "generated_at": "2026-03-11T14:30:00",
      "data_source": {
        "data_key": "brand_mr",
        "config_hash": "a1b2c3d4e5f6",
        "source_file": "data/{wave}/source_data.xlsx"
      },
      "shapes": { "zrx_004_001": {"label": "chrome"}, ... }
    }
  }
}
```

---

## Wave Versioning

```yaml
project:
  wave: "WAVE_Q3Q4_2025"       # any identifier — used in path interpolation
data_source_path: "data/{{wave}}/source_data.xlsx"    # wave-interpolated
template_path: "templates/template.pptx"              # shared (no {{wave}})
output_path: "output/{{wave}}/deck.pptx"              # wave-interpolated
```

Old wave outputs are preserved when switching waves.
