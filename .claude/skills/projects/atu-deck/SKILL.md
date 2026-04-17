---
name: atu-deck
effort: high
paths: []
description: "Project-type skill for Awareness, Trial, Usage (ATU) decks. Invoked by create-deck-workflow and refresh-deck-workflow when the project type is ATU. Encodes ATU-specific methodology defaults, universal slide structure, brand funnel framing, competitive benchmarking, and wave-over-wave comparison logic for any pharma client and brand."
---

## Auto-Detected Context
!`python3 -c "
import glob, os
try:
    import yaml
except ImportError:
    yaml = None
configs = sorted(glob.glob('projects/*/config.yaml'), key=os.path.getmtime, reverse=True) if yaml else []
if configs:
    try:
        with open(configs[0]) as f:
            cfg = yaml.safe_load(f)
    except Exception:
        cfg = {}
    if not isinstance(cfg, dict): cfg = {}
    proj = os.path.dirname(configs[0])
    wave = cfg.get('project',{}).get('wave','')
    brand = cfg.get('project',{}).get('brand','')
    client = cfg.get('project',{}).get('client','')
    ta = cfg.get('project',{}).get('therapy_area','')
    print(f'**Active project:** \`{proj}\`')
    print(f'**Brand:** \`{brand}\` | **Client:** \`{client}\` | **TA:** \`{ta}\`')
    print(f'**Active wave:** \`{wave}\`')
    ctx = f'{proj}/context/{wave}' if wave else f'{proj}/context'
    for md in sorted(glob.glob(f'{ctx}/*.md')):
        print(f'  - \`{os.path.basename(md)}\` ({os.path.getsize(md)//1024}KB)')
else:
    print('**No active project detected** -- user must specify project folder')
"
`

# ATU Deck -- Project-Type Skill

You are a project-type skill for **Awareness, Trial, Usage (ATU)** studies. Workflow skills (`create-deck-workflow`, `refresh-deck-workflow`) invoke you to apply ATU-specific methodology defaults before delegating to lower-level planning, analysis, and creation skills.

This skill is **client-agnostic and brand-agnostic**. It works for any pharmaceutical brand that has an ATU study -- the brand, client, therapy area, competitor set, and visual identity are all resolved at runtime from `config.yaml` and `BRAND{}`.

---

## Cardinal Rules

1. **Never hardcode a brand, client, or therapy area.** All identity comes from `config.yaml` via `project.brand`, `project.client`, `project.therapy_area`. Visual identity (colors, fonts, heading color, template) comes from `BRAND{}` / `CLIENT{}` in `slidegen/pptx_utils/brand.py`.
2. **Brand funnel flow is the core ATU value proposition.** Every ATU deck traces the HCP or patient journey from Awareness through Trial to Usage (and often Loyalty). This funnel structure must be preserved in section ordering and narrative arc.
3. **Competitive benchmarking is always present.** ATU studies track the focal brand against named competitors at every funnel stage. All awareness, usage, satisfaction, and perception slides must include competitive context.
4. **The slide plan drives the slide list -- not this skill.** This skill provides ATU methodology defaults and section structure expectations. The actual slide sequence is determined by `slide-plan-generator-hypothesis` (create flow) or `slide-plan-generator-refresh` (refresh flow).
5. **Respect the brand's visual identity.** Use `get_brand(config.project.brand)` for colors, `CLIENT[brand.client]` for fonts and heading color. Never apply one client's typography or palette to another.
6. **Wave-over-wave comparison applies to recurring ATU studies.** Most ATU studies are wave-based (quarterly or semi-annual). Every metric should show current vs. prior wave with deltas where wave data is available.

---

## Inputs

| Input | Source | Required |
|---|---|---|
| `config.yaml` | `projects/{name}/config.yaml` | Yes |
| `project.brand` | config -- `project.brand` (key into `BRAND{}`) | Yes |
| `project.client` | config -- `project.client` (key into `CLIENT{}`) | Yes |
| `project.wave` | config -- `project.wave` (e.g., "Q1 2026") | Yes |
| `project.therapy_area` | config -- `project.therapy_area` | Yes |
| Slide plan | `context/{wave}/slide_plan.md` | Yes (generated upstream) |
| Narrative threads | `context/{wave}/narrative_threads.md` | Yes (generated upstream) |
| `BRAND{}` / `CLIENT{}` | `slidegen/pptx_utils/brand.py` | Yes (runtime lookup) |

---

## What an ATU Deck Universally Contains

Based on analysis of **143 ATU decks** across the full client portfolio (905-deck corpus, 7,226 slides, 11,310 charts). Every ATU deck follows a consistent section structure organized around the brand funnel.

**ATU chart pattern distribution (143 decks):** `column_stacked_100_vertical` 23.0% | `bar_clustered_horizontal` 22.3% | `line_markers_trended` 14.4% | `bar_stacked_100_horizontal` 13.2% | `column_clustered_vertical` 7.1% | `xy_scatter_abacus` 6.9% | `bar_stacked_horizontal` 5.4% | `column_stacked_vertical` 3.8% | `doughnut_default` 2.2%

**Key difference vs PET:** ATU has 2.1x more `column_stacked_100_vertical` (23.0% vs 10.8% in PET) reflecting heavy funnel composition analysis. PET has more `xy_scatter_abacus` (12.5% vs 6.9%) and `line_markers_trended` (16.5% vs 14.4%).

### Universal Sections (present in 6+ of 8 decks)

| Section | Content | Typical Chart Pattern |
|---|---|---|
| **Cover** | Title slide -- brand, wave, audience, date | `cover` |
| **Methodology & Respondent Profile** | Study design, sample size, respondent demographics, specialty mix | `bar_clustered_horizontal` / `stacked_order` / `table` |
| **Executive Summary** | Key findings synthesized across all funnel stages | `executive_summary` |
| **Awareness** | Aided/unaided awareness, brand familiarity, top-of-mind, awareness trends | `bar_clustered_horizontal` / `bar_stacked_100_horizontal` / `line_markers_trended` |
| **Usage / Prescribing** | Current prescribing, patient share, treatment patterns, line of therapy allocation | `column_stacked_100_vertical` / `bar_clustered_horizontal` / `stacked_order` |
| **Competitive Landscape** | Brand-to-brand comparisons at each funnel stage, market share, positioning | `bar_clustered_horizontal` / `column_clustered_vertical` / `dual_bar_compare` |
| **Patient Demographics** | Patient profile, comorbidities, disease severity, diagnosis journey | `bar_clustered_horizontal` / `doughnut` / `pie` |
| **Perception & Attitudes** | Brand perceptions, attribute ratings, satisfaction, unmet needs | `xy_scatter_abacus` / `bar_clustered_horizontal` / `heatmap_table` |
| **Loyalty / Retention** | Switching behavior, reasons for switching, continuation rates | `bar_stacked_100_horizontal` / `stacked_order` / `doughnut` |
| **Satisfaction** | Overall satisfaction, attribute-level satisfaction, dissatisfaction drivers | `xy_scatter_abacus` / `bar_clustered_horizontal` |

### Common Sections (present in 3-5 of 8 decks)

| Section | Content | Typical Chart Pattern |
|---|---|---|
| **Treatment Journey** | Line of therapy sequencing, treatment decision tree, regimen selection | `column_stacked_100_vertical` / `stacked_order` |
| **Barriers to Adoption** | Reasons for not prescribing, concerns, access challenges | `bar_clustered_horizontal` / `stacked_order` |
| **Drivers of Choice** | Reasons for prescribing, triggers, key differentiators | `bar_clustered_horizontal` / `xy_scatter_abacus` |
| **Brand Funnel Visualization** | Funnel stage conversion (Awareness -> Trial -> Usage -> Loyalty) | `bar_stacked_100_horizontal` / `column_stacked_100_vertical` |
| **Recommendations / Implications** | Action-oriented recommendations tied to findings | `executive_summary` |
| **SWOT Analysis** | Brand-level strengths, weaknesses, opportunities, threats | `table` / text-heavy |

### Optional Sections (study-dependent)

| Section | When Present |
|---|---|
| **Scenario Analysis** | When hypothetical treatment profiles or blinded conjoint is tested |
| **Message Testing** | When ATU includes message recall/effectiveness (hybrid ATU+PET) |
| **Cross-Project / Cross-TA Analysis** | When comparing across therapy areas or sub-indications |
| **Segment Deep Dives** | When specific HCP segments (e.g., academic vs. community, high prescribers vs. low) are analyzed separately |
| **Qualitative Themes** | When open-ended responses are captured |
| **Payer / Access** | When reimbursement, formulary status, or prior authorization is tracked |

---

## ATU Methodology Knowledge

### What is ATU?

**Awareness, Trial, Usage** is a recurring (wave-based) market research study that measures the full HCP or patient journey with a pharmaceutical brand -- from initial awareness through trial to ongoing usage and loyalty. ATU studies track how a brand is positioned in the competitive landscape, what drives or inhibits adoption, and how perceptions evolve over time.

ATU studies are distinct from PET (Promotional Effectiveness Tracking) in their focus: PET measures promotional execution (message recall, rep quality, call effectiveness), while ATU measures market dynamics (brand funnel, competitive positioning, treatment patterns).

### Standard ATU Modules

1. **Awareness & Familiarity** -- unaided awareness, aided awareness, familiarity depth (heard of, know something about, know a lot about), top-of-mind awareness
2. **Trial & Adoption** -- ever prescribed/used, first time prescribing, trial triggers, barriers to trial
3. **Usage & Prescribing** -- current prescribing rates, patient share, line of therapy allocation, treatment regimens, frequency of use
4. **Perceptions & Attitudes** -- brand attribute ratings, product perceptions (efficacy, safety, tolerability, convenience), unmet needs, satisfaction
5. **Competitive Positioning** -- brand-to-brand comparisons across attributes, market share dynamics, competitive switching
6. **Loyalty & Retention** -- continuation rates, switching behavior, reasons for switching, loyalty drivers
7. **Patient Profile** -- demographics, comorbidities, disease severity, time since diagnosis, treatment history
8. **Treatment Journey** -- line of therapy, treatment sequencing, referral patterns, diagnostic pathway

### ATU KBQ Framing

ATU Key Business Questions are about:
- **Brand funnel health:** What percentage of HCPs are aware, have trialed, and are currently using the brand? Where does the funnel narrow?
- **Competitive positioning:** How does the brand compare to competitors on awareness, usage, and key attributes?
- **Adoption barriers:** What is preventing aware-but-not-prescribing HCPs from trialing the brand?
- **Usage drivers:** What motivates prescribers to choose this brand over alternatives?
- **Patient selection:** Which patient profiles are most commonly treated with the brand? Are there underserved segments?
- **Perception gaps:** Where does brand perception differ from clinical reality? What attributes need reinforcement?
- **Loyalty dynamics:** What drives switching? What keeps prescribers loyal?
- **Wave-over-wave trajectory:** Are awareness, usage, and satisfaction improving, stable, or declining?

### Wave-Over-Wave Comparison

Like PET, ATU studies are typically recurring:
- Every metric is shown as current wave vs. prior wave with delta (percentage point change)
- Deltas use universal color coding: green for positive change, red for negative
- Period labels are dynamic: `{{period_current}}` vs. `{{period_prior}}` from config
- Multi-wave studies add trended charts showing 3+ waves as line charts
- The funnel itself is compared wave-over-wave: are conversion rates at each stage improving?

### Competitive Context

ATU decks always position the focal brand against its competitive set:
- The primary brand's colors come from `BRAND[config.project.brand]`
- Competitor colors come from `BRAND[competitor_brand]` (if defined) or `BRAND[config.project.brand].secondary`
- Awareness, usage, and perception slides show all tracked brands side by side
- Competitive language: "leads in awareness," "trails [Competitor] in usage share," "closing the familiarity gap," "perception advantage on [attribute]"

---

## ATU Chart Pattern Distribution

Based on analysis of 11,310 charts across 143 ATU decks (7,226 slides):

| Chart Pattern | ATU % | PET % | ATU Emphasis |
|---|---|---|---|
| `bar_clustered` | 30.6% | 34.7% | Similar -- workhorse for ranked comparisons |
| `bar_stacked_100` | 20.6% | 7.4% | **Much higher** -- ATU relies heavily on composition views (patient mix, treatment share, funnel stages) |
| `column_stacked_100` | 16.9% | 11.2% | **Higher** -- share-of-mind, line-of-therapy allocation, brand share |
| `xy_scatter` (abacus) | 9.4% | 15.5% | **Lower** -- fewer multi-attribute dot comparisons than PET |
| `line_markers` | 5.0% | 12.9% | **Lower** -- fewer trended panels (ATU is less wave-intensive than PET) |
| `doughnut` | 4.8% | 3.2% | Slightly higher -- used for patient demographics and single-metric composition |
| `column_stacked` | 4.5% | 2.3% | Slightly higher -- treatment sequence visualization |
| `column_clustered` | 4.1% | 2.7% | Slightly higher -- segment-level comparisons |
| `pie` | 1.3% | 0.3% | Higher -- patient profile breakdowns |

### Key Structural Differences: ATU vs. PET

1. **Composition-heavy:** ATU decks use 2.8x more `bar_stacked_100` charts than PET (20.6% vs 7.4%). This reflects ATU's focus on shares, funnels, and patient mix rather than message-level recall metrics.
2. **Higher chart density:** ATU averages 2.08 charts per slide vs PET's 1.86. ATU slides frequently use multi-chart panel layouts (3-chart and 4-chart slides are common) to show the same metric across brands or segments.
3. **More tables alongside charts:** ATU has 2.37 tables per slide vs PET's implied lower ratio. The `1_chart_2_table` and `1_chart_3_table` signatures are among the most common, reflecting companion label and data tables.
4. **Less abacus, more stacked bars:** PET uses abacus charts heavily for attribute comparisons (15.5%); ATU prefers stacked bars for similar competitive comparisons (20.6% bar_stacked_100).
5. **Less trended content:** ATU has half the line chart usage (5.0% vs 12.9%), reflecting that ATU waves are less focused on temporal trends than PET's wave-over-wave tracking.

---

## ATU Defaults Applied to Downstream Skills

When `atu-deck` is invoked by a workflow, it applies these methodology defaults before delegating:

### To `hypothesis-generator`
- Ensure hypotheses cover all ATU modules (Awareness, Trial, Usage, Perceptions, Competitive, Loyalty, Patient Profile, Treatment Journey)
- Require PRIOR WAVE VALIDATION hypotheses for funnel metrics and competitive position carried forward
- KBQ mapping must span funnel health, competitive positioning, adoption barriers, usage drivers, and patient selection

### To `atu-insight-writer` (NOT sfea-insight-writer — ATU has its own insight skill)
- Uses ATU arc patterns: FUNNEL_LEAKAGE, SHARE_MOMENTUM, COMPETITIVE_CONVERGENCE, BARRIER_CLUSTER, LOYALTY_EROSION, SEGMENT_SPLIT, ADOPTION_CURVE
- Narrative arcs follow the brand funnel: awareness -> consideration -> trial -> usage -> loyalty
- Executive Summary must cover: Awareness position, Usage/prescribing dynamics, Competitive positioning, Key perception gaps, Patient profile insights
- Recommendations use ATU action verbs: ACCELERATE, UNBLOCK, DEFEND, EXPAND, INVESTIGATE

### To `slide-plan-generator-hypothesis`
- Section ordering: Cover, Methodology, Executive Summary, Recommendations, Awareness, Trial/Adoption, Usage/Prescribing, Perceptions/Attitudes, Competitive, Loyalty/Retention, Patient Profile, Treatment Journey, Appendix
- Every data slide must specify prior/current wave columns and delta computation (for recurring studies)
- Multi-brand comparison slides should include all tracked brands
- Arc sequencing: EXPAND funnel arcs first, then COMPETE arcs, then RETAIN arcs

### To `viz-selector` and `layout-selector`
- Awareness (aided/unaided): `bar_clustered_horizontal` or `single_bar_with_delta`
- Brand funnel: `bar_stacked_100_horizontal` or `stacked_order`
- Usage / patient share: `column_stacked_100_vertical` or `stacked_order`
- Treatment journey / line of therapy: `column_stacked_100_vertical`
- Perceptions / attribute ratings: `xy_scatter_abacus` or `heatmap_table`
- Satisfaction: `xy_scatter_abacus` or `bar_clustered_horizontal`
- Competitive comparison: `clustered_compare` or `dual_bar_compare`
- Barriers / drivers: `bar_clustered_horizontal` or `stacked_order`
- Patient demographics: `doughnut` or `bar_clustered_horizontal`
- Trends (multi-wave): `trended_scorecard` or `line_markers_trended`
- Switching behavior: `bar_stacked_100_horizontal` or `stacked_order`
- Segment comparison: `dual_bar_compare` or `heatmap_table`

### To `slide-creator` and `deck-assembler`
- Brand colors from `get_brand(config.project.brand)`
- Client fonts and heading color from `CLIENT[brand.client]`
- Template path from brand config or `config.yaml` `template_path`
- Shape naming: `zrx_{slide:03d}_{shape:03d}` convention
- Speaker notes: include question codes and full question text on every data slide

---

## Executive Summary & Recommendations

ATU Executive Summaries follow a funnel-organized structure regardless of client.

### Language & Tone
- **Data-anchored:** every claim needs a metric, comparison, or directional ("increased," "declined," "led," "trailed")
- **Funnel language:** "awareness-to-trial conversion," "usage penetration," "funnel leakage at [stage]," "loyalty retention rate"
- **Wave-over-wave language:** "improved vs. {{period_prior}}," "declined wave-over-wave," "stable across waves," "reached highest level tracked"
- **Competitive language:** "leads the category in [metric]," "trails [Competitor] in [metric]," "closed the gap on [attribute]," "extended its lead"

### Recommendation Formatting
- Start with a strong **action verb in CAPS**: EXPAND, REINFORCE, ADDRESS, DIFFERENTIATE, TARGET, EDUCATE, LEVERAGE, CONVERT, RETAIN, STRENGTHEN, SIMPLIFY, ACCELERATE
- Tie each recommendation to a specific funnel finding
- Frame as opportunity: "Opportunity to convert aware-but-not-trialing HCPs by..." not "HCPs are not trialing"
- 4-6 recommendations is the sweet spot
- Organize recommendations by funnel stage (Awareness -> Trial -> Usage -> Loyalty)

---

## Orchestration

When invoked by `create-deck-workflow`:

```
1. Read config.yaml -- resolve brand, client, therapy_area, wave, competitor set
2. Validate brand exists in BRAND{} (or project provides brand_palette override)
3. Apply ATU methodology defaults to all downstream skill invocations
4. Delegate to the standard workflow stages:
     hypothesis-generator (with ATU module coverage requirement)
     atu-insight-writer (ATU-specific hypothesis validation + narrative arcs)
     slide-plan-generator-hypothesis (with ATU section ordering)
     viz-selector + layout-selector (with ATU chart pattern defaults)
     slide-creator + deck-assembler (with brand visual identity)
```

When invoked by `refresh-deck-workflow`:

```
1. Read config.yaml -- resolve brand, client, wave, new period labels
2. Apply ATU defaults to refresh-specific skills:
     deck-reader (ATU slide type recognition)
     slide-plan-generator-refresh (ATU section structure preservation)
     slide-updater (wave label rewrite, headline regeneration, delta recomputation)
     deck-assembler (brand visual identity)
```

---

## Decision Rules

| Situation | Response |
|---|---|
| Brand not in `BRAND{}` | Halt -- user must add brand entry or provide `brand_palette` in config |
| No competitor defined | Proceed -- single-brand slides only; skip competitive comparison slides |
| Study has only 1 wave | Skip trended charts and wave comparison; show absolute values only |
| Study has 2 waves | Use QoQ bars and deltas; no trended scorecards |
| Study has 3+ waves | Include trended scorecards for key funnel metrics |
| Patient ATU (not HCP) | Adjust question framing: "patients who are aware" not "HCPs who are aware"; patient-centered language |
| Hybrid ATU+PET study | Merge sections: include Message Recall and Rep Performance from PET alongside ATU funnel sections |
| Multi-indication study | Use segment comparisons; separate funnel metrics by indication |
| Launch tracker (pre/post approval) | Include scenario analysis slides; awareness baseline metrics are central |
| SWOT analysis requested | Add SWOT section; use text-heavy `executive_summary` renderer with structured bullets |
| Segment deep dives in scope | Add segment-specific funnel slides; use `dual_bar_compare` or `heatmap_table` |
| Qualitative data available | Add `qual_theme_analysis` slides; enable `qual_callout` flags on relevant data slides |

---

## New ATU Project Onboarding

When setting up an ATU project for a new brand, the project team must provide:

### Required

1. **Brand entry in `BRAND{}`** -- or a `brand_palette` override in `config.yaml`. Each brand needs: `client` key, `therapy_area`, `primary_current`, `primary_prior`, `secondary` colors, and optionally `competitor_brand`. Extract colors from the client's slide master or brand guidelines.

2. **Client slide master deck** -- a clean `.pptx` template with the client's layout masters, logos, and footer elements. Stored at `projects/{name}/templates/template.pptx`. If unavailable, the pipeline generates from a blank template using `BRAND{}` / `CLIENT{}` definitions.

3. **Synapse project configuration** -- `project_id`, `reporting_plan_id`, and `analysis_ids` for the study. Enables Track A (JSON-first) and Track D (raw-data-first) data fetching. Alternatively, provide `source_data.xlsx` for Track B (Excel) extraction.

4. **KBQs.md** -- hand-written Key Business Questions for the study. Organized by ATU domain (Awareness, Trial, Usage, Perceptions, Competitive, Loyalty, Patient Profile, Treatment Journey). This file has no auto-generation path; the research team authors it.

5. **Competitor set definition** -- list of tracked competitor brands with their colors (if available in `BRAND{}`). ATU studies typically track 3-8 competitor brands. Define in `config.yaml` under `project.competitors`.

### Strongly Recommended

6. **Prior wave deck** -- a `.pptx` from the most recent delivered wave. Enables `deck-reader` to bootstrap slide structure and `prior-wave-context-builder` to extract findings. Place in `input/wave/{wave}/`.

7. **Survey context** -- survey instrument or questionnaire with question codes, awareness lists, brand lists, and response scales. Enables `survey-context-builder` to map codes to extractions.

8. **Call notes / methodology doc** -- client call notes, study design `.odt`, or methodology brief. Feeds `project-context-builder` for field intel and wave expectations.

### Optional

9. **Competitor brand entries in `BRAND{}`** -- for each tracked competitor. Set `competitors` list in `config.yaml` for automatic color resolution in multi-brand comparison slides.

10. **Label shortcuts** -- `label_shortcuts` mapping in `config.yaml` for long brand names or therapy descriptions. Set `use_label_shortcuts: true` in extraction params.

11. **Patient segment definitions** -- if the study segments by HCP type (academic vs. community), prescribing volume (high vs. low), or other cuts. Define in `config.yaml` under `project.segments`.

---

## References

- `slidegen/pptx_utils/brand.py` -- `BRAND{}`, `CLIENT{}`, `get_brand()`, `get_competitor()`
- `slidegen/viz_selector.py` -- `METRIC_TAG_MAP` (includes `aided_awareness`, `unaided_awareness`, `brand_awareness`, `brand_comparison`, `segment_comparison`, `patient_allocation`, `share_of_mind`)
- `.claude/skills/workflows/create-deck-workflow/SKILL.md` -- full create orchestration
- `.claude/skills/workflows/refresh-deck-workflow/SKILL.md` -- full refresh orchestration
- `.claude/skills/analysis/hypothesis-generator/SKILL.md` -- hypothesis generation
- `.claude/skills/analysis/atu-insight-writer/SKILL.md` -- ATU-specific narrative threads + ES + recs (NOT sfea-insight-writer)
- `.claude/skills/planning/slide-plan-generator-hypothesis/SKILL.md` -- slide plan from narrative
- `.claude/skills/planning/viz-selector/SKILL.md`, `layout-selector/SKILL.md` -- chart/layout selection
- `.claude/skills/creation/slide-creator/SKILL.md`, `deck-assembler/SKILL.md` -- rendering + assembly
- `experiments/deck_analysis/outputs/atu_analysis.md` -- ATU deck analysis synthesis (8 decks, 7 clients)
- `experiments/deck_analysis/atu_analyzer.py` -- ATU deck analysis tool
