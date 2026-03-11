# Session Trace — 2026-02-24
> Rybrevant Associate Asks Project

---

## SESSION 1 — `7025f1f8-1b70-4656-a37f-56c828fa8304`

**Time:** 2026-02-24, 03:43–04:44 UTC (~1 hour)
**Working directory:** `C:\Users\VinothRajapandian\Documents\Claude Apps\Rybrevant - Experiment`

---

### Objective
Build a full v2 9-slide analysis deck using the updated Associate Asks instructions, using the Q4'25 Report as template, with QoQ delta tables, and validate all values against the source Excel.

---

### What Was Done

**1. Reviewed existing project state**
- Read `README.md` to understand prior work
- Confirmed existing deck `Rybrevant_Analysis_Deck.pptx` used matplotlib image charts

**2. Read updated Associate Asks instructions**
- Orange/violet color coding, implications per slide, executive summary slide
- Delta tables show delta only (not repeated message tags)
- Green/red font for deltas
- Survey IDs: RYB IM 738902, RYB NPP 387210, etc.

**3. Explored template and source data**
- Template: `JJ PET RYBREVANT+LAZCLUZE Q4'25 Report.pptx` — 58 slides, 13.33"×7.5"
- Brand colors: Orange #F75824, Red #FF0000, Purple #7638A4
- Fonts: Johnson Display / Johnson Text
- `Lung SFEA SB.xlsx` column mapping confirmed: RYB col7=Q3 Total, col17=Q4 Total; TAG col7=Q3 Total, col13=Q4 Total

**4. Created `extract_data_v2.py`**
- Reads RYB and TAG sheets, extracts all 9 ask datasets, saves `slide_data_v2.pkl`
- Ran successfully

**5. Created `generate_pptx_v2.py`**
- Full 9-slide deck using template, matplotlib charts as images, QoQ delta tables
- Fixed None-value handling in bar/abacus/CTA chart functions
- Output: `Rybrevant_Analysis_Deck_v2.pptx` ✅

**6. Created and ran `validate_data_v2.py`**
- Initial run: 14 FAILs found
- **Root cause 1 (HII FAILs):** Topic-based row matching failed — all AA rows 210–217 shared same `col2="RYB"`. Fix: index-based matching.
- **Root cause 2 (High Impact % bug):** `pct()` multiplied integer values by 100 (57 → 5700%). Fix: use `sf()` for raw value.
- **Root cause 3 (garbled header):** Python string multiplication repeated entire header block.
- After all fixes: **347 checks, 338 PASS, 0 FAIL, 9 SKIP — 100% match** ✅

---

---

## SESSION 2 — `a9943eae-834e-4209-8a9a-a5cc00f55563`

**Time:** 2026-02-24, 07:25–18:27 UTC (~11 hours)
**Working directory:** `C:\Users\VinothRajapandian\Documents\Claude Apps\Rybrevant - Experiment`

---

### Objective
Build a single, perfected, client-delivery slide for Ask #1 (Messaging Recall + Message Effectiveness) using native PPT charts (not images), full message text labels, template branding, and correct delta alignment.

---

### What Was Done

**1. Requirements gathered from user**
- One slide only — Ask #1
- Native PPT charts (not matplotlib images) — editable in PowerPoint
- Delta table: 1 column only (delta value, no repeated message tag column)
- Full message text labels (not coded like "R43")
- All objects aligned; Q4'25 template branding

**2. Research phase**
- Read existing `generate_pptx_v2.py` (slide 3 code for s3_ryb_messaging)
- Extracted 10 full message texts from AA sheet banner plan (rows 30–39, col1):
  - e.g., `"7.1 months improvement in mPFS vs osimertinib (23.7 mo vs 16.6 mo)"` instead of `"Efficacy – mPFS"`
- Confirmed data structure in `slide_data_v2.pkl`: MR Q3/Q4, ME Q3/Q4, delta lists

**3. Created `generate_slide1_final.py` (v1)**
- Template-based: opens Q4'25 report, selects blank layout, strips all original slides
- Native PPT charts via `ChartData` + `XL_CHART_TYPE.BAR_CLUSTERED`
- MR chart: 10 messages (sorted MR Q4 desc, reversed for PPT top-to-bottom rendering)
- ME chart: same message order
- Delta tables: `Table` shape, single Δ column, green/red/grey cell fill, N/A for new Q4 messages
- Output: `Slide1_MR_ME_Final.pptx` ✅

---

### V1 → V5 Iteration History

#### v1 → v2: Fixed delta table reversal bug
**Issue reported:** NCCN Cat1 (top of chart) showed delta -9.2 instead of N/A.
**Root cause:** Chart data fed as reversed list (for PPT top-to-bottom rendering), but delta table was also reversed — misaligning deltas against chart rows.
**Fix:** Delta tables always use original (non-reversed) order. Chart uses reversed, delta table does not.
**Output:** `Slide1_MR_ME_Final_v2.pptx` ✅

| Chart position (top→bottom) | MR value | MR Delta |
|---|---|---|
| 1. NCCN Cat1 | 49.3% | N/A (new Q4 message) |
| 2. Efficacy mPFS | 44.0% | -2.6 |
| ... | ... | ... |
| 10. Safety – ARs | 18.0% | -9.2 |

#### v2 → v3: Adjusted chart widths, clipped delta table bottom
**User request:** Widen message area for legibility, make ME chart uniform width with MR.
**Changes:**
- MR_W: 5.75" → 5.80"; ME_W: 5.45" → 5.80" (now uniform)
- Delta table height trimmed by 0.12" (bottom border moved up)
**Output:** `Slide1_MR_ME_Final_v2.pptx` (saved as v2 due to file lock)

#### v3 → v4 attempt: Unified axis scale, Q3 labels, no gridlines, footer fix
**User request:** Both chart axes should be same scale; show Q3 values; remove gridlines; move footer below J&J logo.
**Changes:**
- Both charts: `minimum_scale=0`, `maximum_scale=100`
- Q3 data labels enabled (lighter grey)
- `has_major_gridlines = False` on both value axes
- Footer pushed from top=6.82" → top=7.20", height trimmed, font reduced to 6pt
**Output:** `Slide1_MR_ME_Final_v3.pptx` ✅

#### v3 → v4: Attempted XML plot area pinning (FAILED — corrupted file)
**User issue:** ME chart bar area still visually larger than MR even with same axis limits.
**Root cause:** MR Y-axis labels consume ~40% of frame width, leaving ~60% for bars. ME has no labels so 100% of frame = bar area.
**Attempted fix:** Added `set_inner_plot_area()` helper using `c:manLayout` XML to pin both bar regions to 50% width.
**Result:** PowerPoint showed "couldn't read some content" — orphaned `externalData rId1` reference caused corruption. Chart pinning was stripped.
**Output:** `Slide1_MR_ME_Final_v4.pptx` ⚠️ (corrupted)

#### v4 → v5: Physical frame sizing (FINAL — correct)
**Decision:** Abandon `c:manLayout` XML entirely. Use physical frame sizing to achieve visual parity.
**Logic:**
- MR Y-axis labels consume ~40% of frame → ~60% = bar area
- If `MR_W = 7.30"`, then bar area ≈ 7.30 × 60% = 4.38"
- Setting `ME_W = 4.40"` (no labels → entire frame = bar area) makes both bar regions physically identical

**Final layout (v5):**

| Object | Left (x) | Width (w) | Top | Height |
|---|---|---|---|---|
| MR Bar Chart | 0.20" | 7.30" | 1.47" | 5.05" |
| MR Delta Table | 7.54" | 0.62" | 1.47" | 5.05" |
| ME Bar Chart | 8.26" | 4.40" | 1.47" | 5.05" |
| ME Delta Table | 12.70" | 0.62" | 1.47" | 5.05" |

**Output:** `Slide1_MR_ME_Final_v5.pptx` ✅ (no corruption, bar areas visually equal)

---

### Additional Work in Session 2

**Message label wrapping**
**User issue:** Messages extending to 3–4 lines.
**Fix:** Rewrote all 10 full text labels to max ~85 chars (from up to 143); raised wrap width to 50 chars. All 10 messages verified at exactly 2 lines.

**Slide rules documentation**
**User request:** Store context on all iterations as elegant rule-based steps for future use.
- Created `pptx_slide_rules.md` in `.claude/memory/` — 12 rule sets covering layout, colors, axes, labels, pitfalls
- Created `SLIDE_RULES.md` in project folder for repo
- Updated `MEMORY.md` to reference the rules file

**Git repository setup**
**User request:** Share trace and scripts with boss; commit to repo.
- Initialized `git init` in project folder
- Created `.gitignore` excluding `.xlsx`, `.pptx`, `.pkl`, `.docx` (sensitive data stays local)
- Staged 14 files (scripts, docs, reports — not data or outputs)
- Committed: `"Add Rybrevant Associate Asks scripts and slide rulebook"`

---

---

## Summary of All Artifacts Created

| File | Session | Purpose |
|---|---|---|
| `extract_data_v2.py` | 1 | Reads xlsx → `slide_data_v2.pkl` for all 9 slides |
| `generate_pptx_v2.py` | 1 | Generates full 9-slide `Rybrevant_Analysis_Deck_v2.pptx` |
| `validate_data_v2.py` | 1 | Cross-validates pkl values against raw xlsx |
| `validation_report_v2.txt` | 1 | 347-row match report (100% pass rate) |
| `generate_slide1_final.py` | 2 | Generates Ask #1 single slide, native PPT charts (v1→v5) |
| `Slide1_MR_ME_Final_v5.pptx` | 2 | Final client-delivery slide |
| `SLIDE_RULES.md` | 2 | 12 rule sets for all future PPTX/chart work |
| `README.md` (updated) | 2 | Full project docs including v1→v5 iteration history |
| `.gitignore` | 2 | Excludes source data/output files from git |

## Key Bugs Fixed

| Bug | Root Cause | Fix |
|---|---|---|
| HII validation FAILs | Topic-based row matching ambiguous when multiple rows share same topic value | Index-based matching |
| High Impact % wrong (5700%) | `pct()` function multiplied integer percentage by 100 | Use `sf()` for raw value |
| Delta table misaligned with chart | Chart data reversed for PPT rendering; delta table also reversed | Delta table always uses original (non-reversed) order |
| c:manLayout corrupted file | Orphaned `externalData rId1` reference in template slides | Abandoned XML approach; used physical frame sizing instead |
| Message labels 3–4 lines | Full texts up to 143 chars; wrap width only 44 | Rewrote texts to max ~85 chars; wrap width raised to 50 |
