---
name: slidegen
description: "Use when working with the SlideGen pipeline — creating, editing, or regenerating YAML-driven PowerPoint decks. Trigger when: creating slides for a project, editing slide N, switching wave data, adding/removing slides, writing renderers or extractors, using pptx_utils functions, or working with the shape registry. Covers the full pipeline from config to PPTX output."
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
| `single_bar_with_delta` | Horizontal bar + QoQ delta column |
| `dual_bar_with_delta` | Two side-by-side bars + deltas |
| `dual_bar_qoq` | Two side-by-side Q4-vs-Q3 clustered bars + deltas |
| `clustered_compare` | Clustered bar comparing two groups + gap/delta columns |
| `qoq_bar_with_delta` | Q4 vs Q3 clustered + delta |
| `two_section_bar` | Two vertically stacked bar sections |
| `stacked_order` | Stacked bar with ordinal breakdown + total column |

---

## Universal Slide Elements

Every data slide includes:
1. `slide_header(slide, headline)` — accent line + headline + badge + separator
2. `section_header_bar(slide, label, top=1.40)` — gray strip with chart title
3. `slide_footer(slide, footer_text)` — source footnote at bottom

---

## Positioning Grid

```
Y=0.00  Top of slide
Y=0.15  Accent line
Y=0.20  Headline / module badge
Y=1.32  Separator line
Y=1.40  section_header_bar
Y=1.85  Chart top (standard)
Y=7.20  slide_footer
Y=7.50  Bottom of slide (SLIDE_H_IN)

X=0.20  Left margin
X=0.30  Chart left edge
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
