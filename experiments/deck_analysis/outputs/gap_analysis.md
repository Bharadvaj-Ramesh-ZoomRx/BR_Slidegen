# Gap Analysis: Decks vs. Existing pptx_utils

This document compares chart types found in real client decks against the 22 existing renderers.

**Caveat:** "Covered" here means the renderer uses a chart of this type — it does NOT mean the renderer matches the exact layout seen in the deck. Deeper comparison of layout signatures, callouts, label placement, etc. is a follow-up pass.

## Covered Chart Types

- **`bar_clustered (57)`** (1511 occurrences) → candidates: `single_bar_with_delta`, `clustered_compare`, `dual_bar_with_delta`
- **`xy_scatter (-4169)`** (674 occurrences) → candidates: `abacus`, `dual_abacus`, `quadrant_scatter`, `followup_rep`, `message_mbd`
- **`line_markers (65)`** (563 occurrences) → candidates: `trended_scorecard`, `trended_activity`
- **`column_stacked_100 (53)`** (486 occurrences) → candidates: `two_section_bar`
- **`bar_stacked_100 (59)`** (324 occurrences) → candidates: `stacked_order`, `two_section_bar`
- **`bar_stacked (58)`** (305 occurrences) → candidates: `stacked_order`, `two_section_bar`
- **`doughnut (-4120)`** (138 occurrences) → candidates: `dual_doughnut`
- **`column_clustered (51)`** (118 occurrences) → candidates: `qoq_bar_with_delta`, `dual_bar_qoq`, `hii_scorecard`
- **`column_stacked (52)`** (98 occurrences) → candidates: `two_section_bar`
- **`xy_scatter_lines (74)`** (60 occurrences) → candidates: `trended_scorecard`, `trended_activity`
- **`line_markers_stacked (66)`** (31 occurrences) → candidates: `trended_scorecard`, `trended_activity`
- **`line (4)`** (14 occurrences) → candidates: `trended_scorecard`, `trended_activity`
- **`pie (5)`** (13 occurrences) → candidates: `dual_doughnut`
- **`xy_scatter_smooth (72)`** (4 occurrences) → candidates: `abacus`, `dual_abacus`, `quadrant_scatter`, `followup_rep`, `message_mbd`

## Uncovered Chart Types

- **`area_stacked_100 (77)`** (6 occurrences) — needs new renderer or lxml helper
- **`area_stacked (76)`** (3 occurrences) — needs new renderer or lxml helper
- **`bubble (15)`** (2 occurrences) — needs new renderer or lxml helper

## Recommended Next Steps

1. **Manual review of high-frequency covered types** — do the existing renderers actually produce what the real decks show? If not, refactor them to be more general.
2. **Uncovered types** — if frequency is high, build new renderers. If low, flag for inline lxml with later extraction.
3. **Layout signature analysis** — the shape signature counts in `report.md` reveal multi-element compositions (e.g., 1 chart + 2 tables = a common clustered_compare pattern). Map these to `LAYOUTS{}` entries in `pptx_utils`.
4. **Color palette** — the top-40 colors feed the `BRAND{}` dict. Client-specific hues (orange, purple, etc.) become named entries.
5. **Font inventory** — if a few fonts dominate, encode defaults in `BRAND{}`.
