---
name: slide-creator
effort: high
paths: ["slidegen/slide_spec/**/*.py", "slidegen/pptx_utils/**/*.py", ".claude/skills/slide-creator/**"]
description: "THE atomic unit. Use when rendering ONE validated SlideSpec into ONE client-ready PowerPoint slide. Takes a complete spec, composes `pptx_utils` calls, and produces a slide with pixel-level fidelity to real-deck coordinate clusters. Trigger when: a workflow skill (new-deck, wave-refresh, single-slide-regen, slide-update, client-followup, etc.) has produced a valid SlideSpec and needs it rendered. Does NOT plan, does NOT pick visualizations, does NOT guess — requires a validated spec. Pair with spec-validator before calling."
---

# slide-creator

The atomic unit of SlideGen. One validated `SlideSpec` in, one client-ready slide out.

## Cardinal Rules

1. **Never infer.** If the spec is missing a field, FAIL LOUDLY — do not guess. Run `validate_spec(spec, strict=True)` before you start rendering. Incomplete specs are the caller's bug, not yours to paper over.
2. **Never pick a chart type.** The spec's `chart_pattern` is final — look it up in `CHART_PATTERNS` and apply. Visualization choice belongs to `viz-selector`, upstream.
3. **Never hardcode brand colors.** Resolve color tokens at render time (see §Color Resolution below). Brand colors are inputs, not constants.
4. **Defaults follow real decks, not personal taste.** Chrome defaults were measured across 4,354 real PET charts (Apr 15, 2026). Do not flip them without evidence.
5. **Layout fidelity is the acceptance bar.** Shape bboxes must land within 0.1" of the corresponding coordinate cluster median. If you deviate, you've failed — fix the composition, don't fudge the measurement.
6. **Use `pptx_utils` primitives only.** Never write raw lxml. If a needed primitive is missing, add it to `pptx_utils/lxml_helpers.py` — don't inline XML.
7. **Same spec → same slide, always.** Rendering must be deterministic. No randomness, no clock reads, no filesystem state reads beyond what the spec references.

---

## Inputs / Outputs

**Input:** One `SlideSpec` (see `slidegen/slide_spec/schema.py`), already passed through `validate_spec(strict=True)`.

**Output:** One slide appended to (or replacing) a position in a PPTX. Returns the python-pptx `Slide` object so callers can hand off to `deck-assembler` for ordering/insertion.

**Side effects:** Shapes are named `zrx_{slide_index:03d}_{shape_index:03d}` via the project `ShapeNamer`. The shape registry is updated with data lineage from `spec.data_lineage`.

---

## Rendering Process (pseudocode)

```
def render_slide(spec: SlideSpec, prs: Presentation, namer: ShapeNamer) -> Slide:
    # 1. Gate: spec must be valid. Refuse to continue otherwise.
    validate_spec(spec, strict=True)

    # 2. Resolve layout
    layout_cfg = LAYOUTS[spec.layout]          # raises KeyError if absent

    # 3. Resolve brand (if spec.brand set, loads BRAND{} entry — else noop)
    brand = get_brand(spec.brand) if spec.brand else None

    # 4. Add blank slide, apply chrome
    slide = prs.slides.add_slide(prs.slide_layouts[BLANK_LAYOUT_IDX])
    apply_headline(slide, spec.headline, layout_cfg, brand)
    if spec.subheadline: apply_subheadline(slide, spec.subheadline, layout_cfg, brand)
    if spec.footer:      apply_footer(slide, spec.footer, layout_cfg)
    if spec.section:     apply_section_bar(slide, spec.section, layout_cfg, brand)

    # 5. Render each component in spec order
    for component in spec.components:
        rect = resolve_position(component.position, layout_cfg)
        match component.type:
            case "chart":         render_chart(slide, component, rect, brand, namer)
            case "label_table":   render_label_table(slide, component, rect, namer)
            case "value_table":   render_value_table(slide, component, rect, namer)
            case "delta_column":  render_delta_column(slide, component, rect, brand, namer)
            case "callout":       render_callout(slide, component, rect, brand, namer)
            case "image":         render_image(slide, component, rect, namer)
            case "textbox":       render_textbox(slide, component, rect, brand, namer)

    # 6. Stamp shape registry with data_lineage
    register_slide(slide, spec.slide_id, spec.data_lineage, namer)
    return slide
```

---

## Top-6 Chart Pattern Recipes

These 6 patterns cover **88% of real PET charts** (Apr 15 deck analysis). Each maps a `chart_pattern` key → `pptx_utils` composition. The `CHART_PATTERNS` dict in `slidegen/pptx_utils/charts.py` provides the defaults (gap, overlap, label_pos, etc.) — slide-creator applies them plus any spec overrides.

### 1. `bar_clustered_horizontal` (35% of charts)

**Use when:** Comparing prior-vs-current values across multiple categories (message recall, attribute ratings, etc.)

**Composition:**
```python
from slidegen.pptx_utils import (
    add_clustered_bar_chart, CHART_PATTERNS,
    invert_cat_axis, hide_cat_labels, set_plot_area_gap, set_overlap,
    set_series_color, enable_data_labels, set_datalabel_format,
    set_datalabel_pos_inside_end, hide_val_labels,
)

pattern = CHART_PATTERNS["bar_clustered_horizontal"]
chart = add_clustered_bar_chart(slide, rect, categories, series_data, direction="bar")
invert_cat_axis(chart)                       # maxMin — observed in 43% of charts
hide_cat_labels(chart)                       # companion label_table is the legend
set_plot_area_gap(chart, pattern["gap"])     # 80
set_overlap(chart, pattern["overlap"])       # 0
for i, series in enumerate(spec_data.series):
    set_series_color(chart, i, resolve_color(series.color, brand))
enable_data_labels(chart, position=pattern["label_pos"])  # inEnd
set_datalabel_format(chart, pattern["val_num_format"])    # "0%"
set_datalabel_pos_inside_end(chart, font_color=resolve_color(spec.chrome.data_labels.font_color, brand))
if not spec.chrome.value_axis.show: hide_val_labels(chart)
chart.has_title = False     # 99% of charts have no title
chart.has_legend = False    # 99% have no legend
```

**Typical layout:** `observed_1chart_1table` (145 slides) with a sibling `label_table` component on the left and `delta_column` on the right.

### 2. `xy_scatter_abacus` (17%)

**Use when:** Dot plot comparing prior vs current across messages (abacus). Each category is a row; prior marker and current marker on the same horizontal line.

**Composition:**
```python
pattern = CHART_PATTERNS["xy_scatter_abacus"]
chart = add_xy_scatter(slide, rect)
for i, series in enumerate(spec_data.series):
    add_scatter_series(chart, series, color=resolve_color(series.color, brand),
                       marker=series.marker or pattern["marker"])
hide_axis(chart, "val")           # value axis hidden; data labels show the %
enable_data_labels(chart, position="above", format="0%")
set_datalabel_color(chart, resolve_color(spec.chrome.data_labels.font_color, brand))
# y-offset scales with row count so labels don't overlap prior row's gridline
```

**Typical layout:** `observed_1chart_2table` (primary + secondary value tables flanking the scatter).

### 3. `line_markers_trended` (14%)

**Use when:** Trend over waves (rolling window: R3M, quarterly).

**Composition:**
```python
pattern = CHART_PATTERNS["line_markers_trended"]
chart = add_line_chart(slide, rect, categories, series_data)
for i, series in enumerate(spec_data.series):
    set_series_line_style(chart, i, series.line_style or "solid",
                          color=resolve_color(series.color, brand))
    set_series_marker(chart, i, series.marker or "circle",
                      color=resolve_color(series.color, brand))
enable_data_labels(chart, position="above", format=pattern["val_num_format"])  # "0%"
chart.has_title = False
if not spec.chrome.legend.show: chart.has_legend = False
# Gridlines default off (84% of real decks); enable only if chrome.gridlines = True
```

### 4. `column_stacked_100_vertical` (11%)

**Use when:** 100%-stacked composition over time (share of X across categories).

**Composition:**
```python
pattern = CHART_PATTERNS["column_stacked_100_vertical"]
chart = add_column_stacked_100(slide, rect, categories, series_data)
for i, series in enumerate(spec_data.series):
    set_series_color(chart, i, resolve_color(series.color, brand))
enable_data_labels(chart, position="ctr", format="0%")
# White labels inside colored segments — font_color typically "#FFFFFF"
set_datalabel_color(chart, "#FFFFFF")
chart.has_title = False
```

### 5. `bar_stacked_100_horizontal` (7%)

Same as `column_stacked_100_vertical` but `direction="bar"` and `invert_cat_axis=True`.

### 6. `column_clustered_vertical` (7%)

Same as `bar_clustered_horizontal` but `direction="col"`, `invert_cat_axis=False`.

---

## Chrome Defaults (from 4,354 real charts)

| Feature | Default | Observed frequency (real decks) | Source |
|---|---|---|---|
| Chart title | `None` (no title) | 99% no title | Apr 15 deck analysis |
| Legend | Hidden | 99% no legend | ” |
| Major gridlines | Off | 84% no gridlines | ” |
| Data labels — format | `"0%"` | 96% use "0%" | ” |
| Data labels — position (bar) | `inEnd` | dominant | ” |
| Data labels — font color (bar) | white when `inEnd` | dominant | ” |
| Category axis (horizontal bar) | Inverted (maxMin) | 43% | ” |
| Category labels (with companion label table) | Hidden (`tickLblPos="none"`) | 371 charts | ” |
| Plot area gap (clustered) | 80 | dominant | `CHART_PATTERNS["clustered_bar"]["gap"]` |
| Plot area gap (stacked-100) | 50 | dominant | `CHART_PATTERNS["bar_stacked_100_horizontal"]["gap"]` |

**If the spec sets `chrome.{field} = <non-default>`, apply it.** The spec is authoritative. These frequencies are the defaults you start from, not overrides.

---

## Color Resolution

The spec carries color tokens. Resolve at render time, not at spec creation:

```python
def resolve_color(token: str, brand: dict | None, context: dict | None = None,
                  deck_ref: dict | None = None) -> str:
    """Returns a 6-char hex string (no '#')."""
    # 1. Explicit hex
    if token.startswith("#"):
        return token[1:].upper()

    # 2. BRAND{} token: "{brand.primary_current}"
    if token.startswith("{brand."):
        key = token[7:-1]                     # strip "{brand." and "}"
        if brand is None:
            raise ValueError(f"Token {token!r} requires spec.brand to be set")
        return rgb_to_hex(brand[key])

    # 3. Context file token: "{context.brand_palette.primary}"
    if token.startswith("{context."):
        path = token[9:-1].split(".")
        return dotted_get(context, path)

    # 4. Deck-reader token: "{deck.slide_4.series_0.color}"
    if token.startswith("{deck."):
        path = token[6:-1].split(".")
        return dotted_get(deck_ref, path)

    raise ValueError(f"Unrecognized color token: {token!r}")
```

**The four sources** (see PRD §6.7):
1. Explicit hex — `"#F75824"`
2. `BRAND{}` token — `"{brand.primary_current}"`, `"{brand.positive}"`
3. Context file — `"{context.brand_palette.primary}"` (loaded from `project_context.md` / `market_context.md` by the orchestrator)
4. Deck-reader extraction — `"{deck.slide_4.series_0.color}"` (populated by `deck-reader` from prior wave PPTX)

---

## Position Resolution

A `Position` is either explicit or a preset reference. Explicit wins; no mixing:

```python
def resolve_position(pos: Position, layout_cfg: dict) -> tuple[float, float, float, float]:
    if pos.is_explicit():
        return (pos.left, pos.top, pos.width, pos.height)
    if pos.is_preset():
        target = pos.layout or None        # fall back to spec.layout if not set
        rect = layout_cfg[pos.preset] if target is None else LAYOUTS[target][pos.preset]
        return (rect["left"], rect["top"], rect["width"], rect["height"])
    raise ValueError(f"Position has neither explicit bbox nor preset reference: {pos}")
```

Position bounds are validated upstream by `spec-validator` — you can trust the values.

---

## Decision Rules

| Situation | Response |
|---|---|
| `validate_spec(spec)` returns errors | REJECT — raise `SpecValidationError`. Do not render. |
| `spec.chart_pattern` not in `CHART_PATTERNS` | REJECT — raise KeyError with available patterns. |
| `spec.layout` not in `LAYOUTS` | REJECT — raise KeyError. |
| Color token uses `{brand.X}` but `spec.brand` is None | REJECT — raise ValueError naming the problem field. |
| `spec.brand="NEWBRAND"` not in `BRAND{}` | ACCEPT with warning if spec uses no `{brand.X}` tokens; REJECT if it does. |
| Spec has `extra_unknown_field` on a component | WARN, ignore. Forward-compat for future schema versions. |
| Chart has `series.values` length ≠ `categories` length (non-scatter) | REJECTED by validator — shouldn't reach slide-creator. |
| Spec has 0 components | REJECTED by validator. |
| Spec's `spec_version` != "1.0" | Check `__init__.py` migration table; either migrate or reject. |

---

## Acceptance Bar (Visual Regression)

Every slide rendered by slide-creator is measured against the 30 coordinate clusters from Apr 15 deck analysis. See `experiments/deck_analysis/visual_regression.py`.

Passes require:

1. **Layout fidelity**: every shape's `left`, `top`, `width`, `height` within `±0.1"` of the cluster median for shapes of the same type in the matching cluster signature.
2. **Chart pattern fidelity**: the generated chart XML carries the correct `barDir`/`grouping`/`gap`/`overlap`/`invertIfNegative`/axis inversion flags per `CHART_PATTERNS[pattern]`.
3. **Data fidelity**: every data-label string matches the expected format (`"{value:.0%}"` or spec override). No stray trailing zeros, no missing %.
4. **Chrome fidelity**: `has_title` matches spec, `has_legend` matches spec, `majorGridlines` presence matches spec.
5. **Color fidelity** (if spec uses `{brand.X}` tokens and `BRAND{}` entry exists): rendered series colors match `BRAND[brand][key]` exactly.

If any axis fails, the slide is not client-ready. Fix the composition — do not mask the failure by relaxing the harness.

---

## References

- **Spec schema**: `slidegen/slide_spec/schema.py`
- **Validator**: `slidegen/slide_spec/validator.py`
- **Chart patterns**: `slidegen/pptx_utils/charts.py` — `CHART_PATTERNS` dict + builders
- **Layouts**: `slidegen/pptx_utils/layout.py` — `LAYOUTS` dict, `observed_*` keys are real-deck medians
- **Brand**: `slidegen/pptx_utils/brand.py` — `BRAND` + `CLIENT`, lookup via `get_brand()`
- **lxml helpers**: `slidegen/pptx_utils/lxml_helpers.py` — the 20 functions that cover 100% of observed OOXML needs
- **Visual regression harness**: `experiments/deck_analysis/visual_regression.py`
- **Coordinate clusters (golden reference)**: `experiments/deck_analysis/outputs/layout_clusters.json`
- **PRD §6.7** (spec-as-contract) and **§6.8** (data lineage via Connector tags)
- **Deck analysis findings**: `experiments/deck_analysis/outputs/ACTIONABLE_FINDINGS.md`
