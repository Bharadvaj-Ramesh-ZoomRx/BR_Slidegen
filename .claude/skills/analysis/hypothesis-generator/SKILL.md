---
name: hypotheses
effort: max
description: "Use when generating a hypothesis bank from context files. Trigger when: user says 'generate hypotheses', 'build hypothesis bank', or needs testable predictions organized by KBQs for a PET study wave. Reads 5 context files: Market Context, Project Context, Prior Wave Context, KBQs, and Survey Context. Prior wave findings are fed directly to generate validation hypotheses — testing whether prior wave findings persist, reverse, or improve in the current wave."
---

## Auto-Detected Context
!`python3 -c "
import glob, json, os
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
    print(f'**Active project:** \`{proj}\`')
    print(f'**Active wave:** \`{wave}\`')
    # Check all 5 required/recommended context files
    mc = f'{proj}/context/market_context.md'
    ctx = f'{proj}/context/{wave}' if wave else f'{proj}/context'
    files_to_check = [
        ('market_context.md', mc, 'Required'),
        ('project_context.md', f'{ctx}/project_context.md', 'Required'),
        ('prior_wave_context.md', f'{ctx}/prior_wave_context.md', 'Recommended'),
        ('survey_context.md', f'{ctx}/survey_context.md', 'Recommended'),
    ]
    for name, path, req in files_to_check:
        if os.path.exists(path):
            print(f'  ✓ {name} ({os.path.getsize(path)//1024}KB) [{req}]')
        else:
            print(f'  ✗ {name} — MISSING [{req}]')
    # Check KBQs in input folder
    for case in ['Wave', 'wave']:
        kbq = f'{proj}/input/{case}/{wave}/KBQs.md'
        if os.path.exists(kbq):
            print(f'  ✓ KBQs.md ({os.path.getsize(kbq)//1024}KB) [Required]')
            break
    else:
        print(f'  ✗ KBQs.md — MISSING [Required]')
else:
    print('**No active project detected** — user must specify project folder')
"
`

# Hypotheses Generator v2 — Storyboard-Integrated

You are a senior market research analyst. Your job is to generate a comprehensive, structured bank of testable hypotheses from context files, and organize them into a 3-part storyboard narrative. You do not need wave data to run — all hypotheses are derived from context and KBQs alone.

---

## STEP 1 — Confirm project and resolve context file paths

Use the auto-detected project, wave, and file status shown above. If correct, proceed. If not, ask the user to specify.

The required files are:

| File | Path | Required? |
|------|------|-----------|
| **Market Context** | `{project}/context/market_context.md` | Required |
| **Project Context** | `{project}/context/{wave}/project_context.md` | Required |
| **Prior Wave Context** | `{project}/context/{wave}/prior_wave_context.md` | Required if exists |
| **KBQs** | `{project}/input/Wave/{wave}/KBQs.md` | Required |
| **Survey Context** | `{project}/context/{wave}/survey_context.md` | Strongly recommended |
| **Qualitative Data** | `{project}/context/{wave}/qualitative_data.json` | Optional — enables QUALITATIVE hypotheses |

If `prior_wave_context.md` is missing: proceed but flag — all prior-wave validation hypotheses will be based on narrative summaries in `project_context.md` only (less specific).

If `survey_context.md` is missing: proceed but flag — "Test with:" lines will be analytical descriptions, not question-code references.

If `qualitative_data.json` is missing: proceed with quantitative hypotheses only — no QUALITATIVE type hypotheses will be generated, and no `qual_callout` flags will be added. This is the normal path when no raw data file was provided.

If `qualitative_data.json` exists: read the `_meta` summary to understand which verbatim questions are available, how many responses each has, and which sheets they come from. This enables QUALITATIVE hypotheses and `qual_callout` flags (see STEP 4).

If either required file (`market_context.md`, `project_context.md`, `KBQs.md`) is missing: stop and tell the user which is absent.

---

## STEP 2 — Read all files in full

Read every available file completely before generating anything. Do not proceed until all are read.

- **Market Context** — clinical evidence, competitive dynamics, physician mindset, product knowledge, pipeline threats
- **Project Context** — study design (TP1/TP2), field intelligence (client-shared), methodology changes, open action items, wave-specific expectations
- **Prior Wave Context** — structured domain-by-domain findings from the prior delivered report: headlines, specific metrics, segment splits (Academic vs Community), open action items, verbatim recommendations carried forward
- **KBQs** — the organising structure; all business questions by domain
- **Survey Context** — question codes, question text, response scales, tracked message list, segment operationalization — used exclusively to write precise "Test with:" lines
- **Qualitative Data** (if exists) — read `_meta` section only at this stage. Note: which verbatim question codes are available (e.g., Q1.53A perception change, Q1.88 interaction improvement, Q2.00 unaided recall, Q3.35 biomarker barriers), how many responses each has, which sheets they come from. Do NOT read individual responses yet — that happens during Stage 3 validation.

---

## STEP 2.5 — Decompose context

Before generating a single hypothesis, extract and organize the following from the files you just read. This decomposition is the grounding layer — every hypothesis must trace back to something here.

### A. Storyboard Configuration

Extract the three storyboard inputs:

**S1 — Period Comparison**
Read the project wave/period fields. Identify: latest period vs. prior period.
Format: `[period_current] vs [period_prior]`
If no explicit period is declared, infer from any wave or time reference in the context.

**S2 — Effectiveness Lens**
Look for a declared `### Effectiveness Lens` under `## Analytical Framework` in the context.
Extract: the split that defines "working vs. not working" for this project (e.g., HII vs Others).
If no Analytical Framework section exists, infer the most logical effectiveness split from the study design (e.g., high impact vs. standard interactions, visual aid used vs. not).
Note any inference with: `[Inferred — no Analytical Framework declared]`

**S3 — Action Segments**
Look for declared `### Action Segments` under `## Analytical Framework` in the context.
Extract: all named segments and their groups (e.g., Practice Setting: Academic / Community).
If no segments are declared, infer from any segmentation described in the study design.
Note any inference with: `[Inferred — no Action Segments declared]`

---

### B. Wave-Level Context Changes

List every documented change since the prior wave:
- Survey / question format changes (reframed scales, new questions, new benchmarks)
- Methodology changes (screener changes, sample composition)
- Field / product changes (new messages, new visual aids, new formulation, sales force changes)
- Competitive changes (new data, competitor moves, market events)

Each item here will generate at least one hypothesis.

---

### C. Client Priorities and Open Action Items

List every concern, priority, or open analytical action item explicitly mentioned in the context.
Each item here MUST generate at least one hypothesis in STEP 4.

---

### D. Prior Wave Findings — Validation Baseline

This section is populated **primarily from `prior_wave_context.md`** — the structured extract of the prior delivered report. If `prior_wave_context.md` exists, use it as the authoritative source for prior wave data; supplement with `project_context.md` Section 3 only where the structured file has gaps.

Extract three layers:

**D1 — Metrics to validate (continuation/reversal hypotheses)**
For every domain in `prior_wave_context.md`, list:
- The headline finding (verbatim from the prior report)
- The specific metric value if available (from the Metrics Snapshot table)
- The trend direction stated (↑/↓/→)
- The segment split if reported (Academic vs Community)

Each finding generates at least one directional hypothesis: will this finding **persist**, **strengthen**, **reverse**, or **diverge further by segment** in the current wave? The rationale must explain the mechanism — what in the field, market, or client response would drive continuity or change.

**D2 — Recommendations to test (action-validation hypotheses)**
For every recommendation in `prior_wave_context.md` Section 5 (Recommendations Carried Forward):
- State the recommendation verbatim
- Generate a hypothesis about whether it was acted on and what the measurable effect would be
- Format: "If reps acted on the recommendation to [X], then [metric Y] will show [direction] vs prior wave, because [mechanism]"
- These hypotheses are inherently ACTION ITEM flagged

**D3 — Anomalies and open questions**
List any prior wave metrics flagged as unexpected, declining, or not yet explained.
From `prior_wave_context.md` Section 6 (Unanswered Questions / Gaps).
Each item MUST generate at least one hypothesis explaining the likely cause or expected resolution.

---

### E. Methodology Artifacts

List every survey or methodology change that could produce a number shift not driven by real field behavior.
Each item here MUST appear as a methodology artifact hypothesis.

---

Display the full decomposition to the user and confirm:
> "Here is the context I've extracted before generating hypotheses. Does this look complete? Anything to add or correct?"

Proceed once confirmed.

---

## STEP 3 — Extract organizing structure

KBQs come from the dedicated KBQs file read in STEP 2. List all KBQs organized by domain and display to the user. Ask:
- "Are there any additional KBQs to add for this wave?"
- "Any to exclude?"

Proceed once confirmed.

---

## STEP 4 — Generate hypotheses

### Core principle

For every context item — wave change, prior wave finding, client action item, competitive move, methodology change — ask: **"What will happen to which survey metric, and why?"** Generate as many hypotheses as the context supports. Do not cap at 2-4 per KBQ. If context supports 6 hypotheses for one KBQ, write 6.

### Coverage requirements (non-negotiable)

- Every methodology change — at least one hypothesis about the metric it affects. Flag: `METHODOLOGY ARTIFACT`
- Every open action item from Project Context or Prior Wave Context — at least one hypothesis. Flag: `[ACTION ITEM]`
- **Every domain finding in Prior Wave Context** — at least one validation hypothesis (persist / strengthen / reverse / diverge). Flag: `PRIOR WAVE VALIDATION`
- **Every recommendation in Prior Wave Context Section 5** — at least one action-validation hypothesis testing whether it was acted on. Flag: `[ACTION ITEM]` + `PRIOR WAVE VALIDATION`
- **Every unanswered question in Prior Wave Context Section 6** — at least one hypothesis proposing a resolution
- Every wave change (new message, new formulation, VA redesign, survey change) — at least one hypothesis
- Every KBQ domain — at least one hypothesis

### Qualitative hypotheses (only when `qualitative_data.json` exists)

QUALITATIVE hypotheses are **separate** from quantitative hypotheses. They produce dedicated qualitative analysis slides (not just callouts on quant slides). Generate them only when:

1. **The "Why" test passes** — the hypothesis asks why something happened, not just what happened. "Did LTIP improve?" is quant. "Why is OS recall flat despite VA placement?" needs qual.
2. **The KBQ test passes** — the linked KBQ requires causal explanation, not just measurement.
3. **The "not already explained" test passes** — no other quant metric explains the finding. If TAG ME explains the call quality gap, no qual needed.

**Expect 5-8 QUALITATIVE hypotheses per wave, not more.** If you find yourself writing more than 10, you are overdoing it. Qual is for explaining the unexplained, not decorating the obvious.

**Format for QUALITATIVE hypotheses:**

```
**HQ[N] — [What HCPs are expected to say about X — stated as a testable theme prediction]**
[Rationale: why do we expect this theme to dominate? Reference the quant finding that creates the "why" question.]
Test with: [verbatim question code] — "[abbreviated question text]" (n=[response count from qualitative_data.json])
Qual analysis: Theme frequency + representative quotes
Dedicated slide: Yes
```

**Additionally**, some quantitative hypotheses may benefit from a `qual_callout` — a verbatim quote on the quant slide that explains the "why" behind the number. Add this flag sparingly:

```
**H[N] — [Standard quant hypothesis]**
[Rationale...]
Test with: [quant question code]
qual_callout: [verbatim question code] — surface dominant theme as callout on this slide
```

**`qual_callout` rules:**
- Maximum 6-8 callouts across the entire hypothesis bank
- Only where the verbatim directly explains the quant finding
- The callout is a single quote + theme stat, not a full analysis
- Never on methodology artifact hypotheses

### Hypothesis generation logic

For each hypothesis, the reasoning chain is:
1. **Context trigger** — what specific thing changed or is known (from any of the 4 files)?
2. **Metric affected** — which survey question/metric will reflect this?
3. **Predicted direction** — will it go up, down, stay flat, diverge by segment?
4. **Mechanism** — why? Name the specific causal link.

### Source material per hypothesis

Draw from all 4 files:
- Market Context: clinical data, competitive moves, physician mindset, AZ strategy
- Project Context: field intel, prior wave ES, open action items, wave changes
- KBQs: which domain this hypothesis addresses
- Survey Context: which question measures this metric; what scale; which segment cuts are available

---

## OUTPUT FORMAT

```
# Hypothesis Bank — [Wave]
**Generated from:** [Market Context] + [Project Context] + [KBQs] + [Survey Context]
**Date:** [today's date]
**Total hypotheses:** [N]

---

## [DOMAIN NAME]
**KBQ:** [Full KBQ text]

**H[N] — [What will happen to which metric — stated as a specific prediction]**
[2-3 sentences of rationale. Name the specific causal mechanism. Reference real specifics: message labels (e.g., NCCN-NSCLC, Efficacy-OS), question codes with their text (e.g., Q2.10 — "Which of these messages do you specifically recall hearing?", Q1.85b — "How likely are you to increase your prescribing?"), prior wave metrics from prior_wave_context.md (e.g., "recall was 45% in Q4"), clinical data points, competitive moves. Write as continuous prose — no bullet points inside rationale.]
Prior wave baseline: [value or finding from prior_wave_context.md this hypothesis validates — omit line if no prior wave data available for this metric]
Test with: [question code — "abbreviated question text"] split by [segment or cut, if applicable]
METHODOLOGY ARTIFACT — verify before client presentation [add only if hypothesis is about a methodology change creating a false signal]
PRIOR WAVE VALIDATION — [persist / strengthen / reverse / diverge] [add only if hypothesis directly tests a prior wave finding or recommendation]
[ACTION ITEM] [add only if hypothesis directly addresses a client open action item or recommendation carried forward]

```

Repeat H[N] blocks continuously across all domains. Number does not restart per domain.

### Qualitative hypotheses section (only if `qualitative_data.json` exists)

After all quantitative domain sections, add a separate section:

```
---

## QUALITATIVE HYPOTHESES
**Available verbatim questions:** [list from qualitative_data.json _meta]

**HQ[N] — [Theme prediction about what HCPs say]**
[Rationale referencing the quant finding that motivates the "why" question]
KBQ: [linked KBQ]
Test with: [verbatim Q code] — "[question text]" (n=[count])
Qual analysis: Theme frequency + representative quotes
Dedicated slide: Yes

```

### Quantitative hypotheses with qual_callout flags

Within the regular quantitative hypothesis sections, add the `qual_callout` line where appropriate:

```
**H[N] — [Standard quant hypothesis]**
[Rationale...]
Test with: [quant Q code]
qual_callout: [verbatim Q code] — surface [expected theme] as callout

```

---

## RULES

1. **No context stored in this skill.** Everything comes from reading the files at runtime.

2. **No cap on hypotheses per domain.** Generate as many as the context supports.

3. **Be specific, not generic.** Name the exact message, question code, segment, trial, or competitive move. A hypothesis that could apply to any pharma study is unacceptable.

4. **New metrics with no baseline cannot use wave-over-wave language.** State expected direction only.

5. **Methodology artifacts are mandatory** wherever a survey or methodology change is documented.

6. **Open action items must become hypotheses.** Every client open action item generates at least one hypothesis.

7. **Prior wave anomalies must generate directional hypotheses** — will the anomaly persist, reverse, or worsen in this wave?

8. **Hypotheses are about field behavior, not survey structure.** A hypothesis whose rationale is primarily "this signal is captured in question X rather than question Y" is a data mapping note, not a testable business insight — do not generate it. The rationale must describe what happens in the field (rep behavior, HCP response, competitive dynamic) and why. Question codes appear only in the "Test with:" line as measurement instruments, not as the subject of the hypothesis itself.

9. **No circular predictors.** Do not write a hypothesis where the predictor variable is a sub-route or derivative of the outcome variable. Example: "OS delivered first -> OS recalled" is invalid because delivery order is only known for interactions where the HCP already recalled OS — the predictor is derived from the outcome. Before writing a driver hypothesis, confirm that the predictor can be measured independently of the outcome.

10. **Driver hypotheses are output-first.** When hypothesising about what drives a high-value outcome (HII, LTIP, close rate), frame the hypothesis from the output metric down — not from the input metric forward. Correct form: "In HII interactions, [input metric] will be significantly higher than in non-HII interactions." Incorrect form: "[Input metric] first -> higher LTIP." This matches the standard analytical approach of splitting on the outcome and comparing inputs across the high vs low groups.

11. **Every empirical claim in a rationale must be traceable to a named source.** If the rationale states a specific number, trend, or prior wave finding (e.g., "~40% of interactions did not lead with 1L", "academic HCPs recalled 3+ messages"), that claim must come from one of the four input files — prior wave ES in Project Context, field intelligence, client action item, or Market Context. Do not invent specificity to make a hypothesis appear grounded. If a plausible structural pattern exists (e.g., academic vs community engagement differences) but no prior wave data supports a specific claim, state the direction only — do not fabricate a number. Hypotheses built on general pharma knowledge with no project-specific signal are disqualified under Rule 3, regardless of how specific they appear.

12. **Do not compare across respondent pools.** This study uses separate survey arms for each brand (TP1/TP2 per brand). RYB respondents and TAG respondents are different HCPs in different call contexts. A hypothesis that compares a RYB metric for RYB respondents against the same metric for TAG respondents (e.g., "HCPs who recall the RYB OS message will rate it higher than HCPs who recall the TAG OS message") is not a within-subject comparison — it is a cross-population comparison confounded by respondent differences, brand exposure, and call context. Only compare across brands when the study design explicitly provides a shared benchmark or cross-brand rating question. Otherwise, test each brand's metrics in absolute terms or vs prior wave.

14. **Prior wave validation hypotheses require a baseline.** A `PRIOR WAVE VALIDATION` hypothesis must cite a specific prior wave finding as its baseline — either a metric value from the Metrics Snapshot table in `prior_wave_context.md`, or a verbatim headline finding. Do not write a validation hypothesis that says "X may have changed" without stating what X was in the prior wave.

15. **Recommendation-validation hypotheses must name the recommendation.** When generating a hypothesis from a carried-forward recommendation (D2), quote or closely paraphrase the recommendation so the client can see their own guidance being tested. These are the most client-valued hypotheses — they close the loop between what was advised and what was measured.

13. **Do not conflate different survey constructs.** Message Effectiveness (ME — motivation/believability rating of a recalled message) and message association (which brand an HCP connects a claim to) measure different things and cannot be used interchangeably to validate the same hypothesis. If a hypothesis requires two different constructs to both be true, split it into two separate hypotheses — one per construct — each with its own clean "Test with:" line.
