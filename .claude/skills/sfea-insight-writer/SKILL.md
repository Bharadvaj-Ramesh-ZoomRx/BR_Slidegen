---
name: sfea-insight-writer
description: "Use when writing data-validated analysis, slide headlines, executive summary, and recommendations for a PET SFEA wave. Trigger when: user says 'write insights', 'build ES', 'run insight writer', or Stage 2 hypothesis bank has been signed off and Stage 3 is ready to begin. Reads hypothesis_bank.md, project_context.md, and source_data.json. Runs in 4 phases: Phase 0 (auto) validates hypotheses vs data → validated_analysis.md; Phase 1 writes talking headlines → slide_headlines.md; Phase 2 writes ES → exec_summary.md; Phase 3 appends Recommendations."
---

# SFEA Insight Writer

Generates data-validated analysis, consulting-grade slide headlines, executive summary, and recommendations for any SFEA study wave. Invoked after the hypothesis bank is signed off.

---

## PIPELINE POSITION

```
Stage 0: index_excel()            →  source_data.json
Stage 1: /build-project-context   →  project_context.md
Stage 2: /hypotheses              →  hypothesis_bank.md
                                        ✋ User signs off
Stage 3: /sfea-insight-writer     ←  YOU ARE HERE
         Phase 0: Data Analysis   →  validated_analysis.md
         Phase 1: Headlines       →  slide_headlines.md       ✋
         Phase 2: ES              →  exec_summary.md          ✋
         Phase 3: Recs            →  (appended to exec_summary.md)
Stage 4: /slide-plan              ←  reads all three outputs
Stage 5: config.yaml + generate_deck()
```

---

## STEP 1 — Resolve file paths

Ask the user for the **project folder** and **wave name** only. Derive all paths:

| File | Path | Required? |
|------|------|-----------|
| **Hypothesis Bank** | `{project}/context/{wave}/hypothesis_bank.md` | Required |
| **Project Context** | `{project}/context/{wave}/project_context.md` | Required |
| **Source Data** | `{project}/context/{wave}/source_data.json` | Required |
| **ES Format** | Ask user — A through H; default **A** | Optional |

If any required file is missing, stop and tell the user which is absent.

Report before reading:
```
Input files:
  ✓ hypothesis_bank.md     (signed-off hypothesis bank)
  ✓ project_context.md     (brand names, wave labels, field intel, methodology flags)
  ✓ source_data.json       (survey data keyed by question code)
ES format: A (Narrative + Recommendations)
```

---

## STEP 2 — Read all files in full

Read all three files completely before doing anything else.

**From `project_context.md` extract and hold:**
- Primary brand and competitor brand(s) — use exact names throughout; never "the competition"
- Wave label (current) and prior wave label — for QoQ framing
- Sample sizes — n= per brand per wave
- Methodology flags — survey changes, screener/benchmark changes affecting QoQ comparability
- Prior wave top findings — for contrast and trend framing
- Study design notes — HII definition, segment cuts used, new questions this wave

**From `hypothesis_bank.md` extract and hold:**
- All hypotheses with their domain, rationale, "Test with:" question codes, segment cuts, and flags
- Expected direction per hypothesis (the prediction)
- METHODOLOGY ARTIFACT and [ACTION ITEM] flags

**From `source_data.json` extract and hold:**
- For each question code referenced in the hypothesis bank: prior value, current value, and delta (current − prior in pp)
- Segment splits where available (e.g., HII vs. non-HII, Community vs. Academic)
- If a code is in the `_sheets` index but not yet fully extracted: note it as "extraction needed" — do not skip it

---

## PHASE 0 — DATA ANALYSIS

For every hypothesis in the bank, validate it against the actual data. This phase is **automatic** — no user gate.

### Process

For each hypothesis:

1. **Extract question codes** from the "Test with:" line only — never from rationale prose
2. **Look up values** in `source_data.json` — prior, current, delta per brand
3. **Compare direction**: does the actual delta confirm, partially confirm, or contradict the hypothesis?
4. **Check segment splits**: if HII/non-HII or setting cuts are specified in "Test with:", look up those values too
5. **Apply methodology flag** if the hypothesis was already flagged — note it affects interpretation

### Validation status

| Status | Criteria |
|--------|----------|
| **CONFIRMED** | Data moves in the predicted direction; delta is meaningful (>3pp or crosses significance threshold) |
| **PARTIALLY CONFIRMED** | Direction correct but magnitude small (<3pp), or confirmed for one brand/segment but not another |
| **NOT CONFIRMED** | Data moves opposite to prediction, or no meaningful change where movement was predicted |
| **INSUFFICIENT DATA** | Question code not found in source_data.json, or n too small to interpret |

### Output format — validated_analysis.md

Phase 0 is a **data extraction stage only** — no interpretation, no narrative, no "so what". Write numbers and validation status. Strategic meaning is added in Phase 1.

Write one block per hypothesis, grouped by domain:

```
## [Domain Name]

### H[N]: [Hypothesis statement — copied from hypothesis bank]
**Status:** CONFIRMED / PARTIALLY CONFIRMED / NOT CONFIRMED / INSUFFICIENT DATA
**Data:**
- [Q code] [Primary Brand]: [prior_wave] = X%, [current_wave] = Y% (Δ +/-Zpp)
- [Q code] [Competitor]:    [prior_wave] = A%, [current_wave] = B% (Δ +/-Cpp)
**Segment data:** [HII: X% vs. non-HII: Y%; or Community: X% vs. Academic: Y%] ← omit if not available
**Data summary:** [One factual sentence — numbers only, no interpretation.
  Good: "RYB recall = 61% (−2pp QoQ); TAG recall = 58% (+3pp QoQ); gap = 3pp (was 8pp in Q4'25)."
  Bad:  "RYB still leads but the gap narrowed — TAG recovered while RYB slipped slightly."]
⚠️ [Methodology flag if applicable — affects QoQ comparability]
[ACTION ITEM] ← if flagged in hypothesis bank
```

If a question code is not found in `source_data.json`:
```
**Status:** INSUFFICIENT DATA
**Note:** Q code [X] not found in source_data.json — extraction may be needed before this hypothesis can be validated.
```

Save as `{project}/context/{wave}/validated_analysis.md`.

**→ No user gate after Phase 0.** Proceed directly to Phase 1.

---

## PHASE 1 — SLIDE HEADLINES

Write one talking headline per domain (and per slide where the hypothesis bank provides slide-level granularity). Headlines are **strategic insights** — data is the evidence, project context is the frame.

### Context Anchoring (do this before writing each headline)

Before writing the headline for a domain, look up the relevant context from `project_context.md`:

| Question | Where to find it |
|---|---|
| *Why might this metric be moving?* | Market context section — competitive events, label updates, formulary changes, campaign launches |
| *What did reps report doing differently?* | Field intelligence / call notes section |
| *Was this flagged as a risk or priority last wave?* | Prior wave ES findings + recommendations |
| *What is the client trying to decide or act on?* | KBQs / client priorities section |
| *Is there a strategic moment (readout, launch, review) this connects to?* | Market context / project notes |

Use the answers to frame the headline around **what is at stake** — not just what moved.

**Example: same data, two different framings**

*Data fact (Phase 0):* RYB recall = 61% (−2pp QoQ); TAG recall = 58% (+3pp QoQ); gap = 3pp (was 8pp)

*Metric-sy headline (wrong — no context):*
> "RYB maintained recall leadership at 61% vs. TAG 58%; the gap narrowed from 8pp to 3pp QoQ"

*Strategic headline (right — uses context):*
> "RYB's recall lead is narrowing as TAG gains ground on NCCN messaging — the argument most tied to 1L prescribing intent; reinforcing NCCN delivery before the OS data readout is now the priority call to action"

The difference: the strategic version tells the client what is at risk, why it matters now, and what to do.

---

### Format

Every headline has two parts:

```
[Part 1 — What happened + data anchor: direction, metric, number]
; [Part 2 — Why it matters / what's at stake / what to do — grounded in context]
```

Use a **semicolon** for additive or parallel findings. Use an **em-dash (—)** when Part 2 is a sharp contrast or pivot.

### Anatomy

```
[Brand] [direction verb] [metric] [data anchor from Phase 0]
; [strategic consequence or action — drawn from project context]
```

**More examples:**

- *"[Primary Brand] LTIP improved to 78% (+7pp) but the gain is concentrated entirely in HII interactions — the 44pp HII vs. non-HII gap confirms that converting more calls to high-quality interactions is the single highest-leverage field action"*
- *"[Primary Brand] reps outperformed [Competitor] on 5 of 6 call quality attributes — but competitor knowledge remains the widest gap (−12pp), a vulnerability as [Competitor] expands its OS data narrative with HCPs"*
- *"[Primary Brand]'s NCCN message recall declined 6pp while [Competitor] gained 5pp — a full reversal in one quarter that threatens [Primary Brand]'s 1L position ahead of the upcoming formulary review"*

### Headline Rules

**Rule 1 — Strategic insight, not data narration**
- ✅ Use data as evidence for a strategic claim
- ✅ Connect the movement to something the client cares about: a competitive threat, a commercial decision, a coaching priority, a timing risk
- ❌ Never write a headline that only describes what the numbers did — that belongs in the chart, not the headline
- ❌ Never write a prediction or hypothesis as a finding
- For PARTIALLY CONFIRMED or NOT CONFIRMED hypotheses: write what the data *actually* showed and why it matters, not what was predicted

**Rule 1a — Every context claim must be file-sourced, not inferred**

Every strategic assertion beyond what the numbers show must be traceable to one of the input files:

| Claim type | Required source |
|---|---|
| A message or metric "drives" prescribing / LTIP | `source_data.json` driver analysis OR `market_context.md` documenting this linkage |
| A competitive event explains a metric movement | `project_context.md` — field intelligence or call notes |
| A strategic deadline or timing risk | `project_context.md` — project timeline or client priorities |
| A clinical or label claim | `market_context.md` — documented product/disease context |
| General pharma inference not in any file | ❌ Not permitted — remove the claim or soften to what the data alone supports |

**Rule 2 — Direction vocabulary**

| Direction | Preferred words |
|---|---|
| Positive | maintained, sustained, strengthened, improved, increased significantly, delivered, exceeded, outperformed, led, held a clear edge |
| Stable/Parity | remained consistent, remained stable, on par with, in line with, broadly similar |
| Negative | declined, dropped significantly, lagged, trailed, fell short, showed a downward trend |
| Opportunity | *"opportunity exists to [verb]"* / *"opportunity remains in [area]"* / *"[metric] remains an area for focus"* |

**Rule 3 — Always anchor to wave and brand**
- Reference the current wave explicitly: *"In Q1'26…"* or *"In the current wave…"*
- Name both brands when comparing — never "the competitor"

**Rule 4 — Numbers are mandatory**
- Every headline must include at least one specific % or pp value from validated_analysis.md
- Use `~` for approximate figures, `%` not "percent", `pp` for percentage points, `vs.` not "versus"
- Limit to 2 numbers — pick the most telling

**Rule 5 — "Directionally" for small n or borderline findings**
- When n is small or delta is borderline: *"Directionally, [Brand] [finding]…"*

**Rule 6 — Slide-type patterns**

| Slide type | Headline pattern |
|---|---|
| Activity / SOV | *"In [wave], [Brand] [led/trailed/maintained] SOV at ~X% vs. [Competitor] ~Y%; [interaction frequency context]"* |
| Message Recall | *"[Message X] was recalled by X% of HCPs — [Brand]'s most-recalled message; [weaker message] remains an opportunity at X%"* |
| Message Effectiveness | *"[Brand]'s [message] motivation score improved +Xpp QoQ; believability held at X% — an area for further reinforcement"* |
| Rep Performance | *"[Brand] rep performance [remained strong/improved/declined] in [wave] — [X of Y] attributes above benchmark; [specific gap]"* |
| LTIP / Closing | *"[Brand] LTIP [improved/declined] to X% in [wave] (+/-Xpp); HII interactions drove X% LTIP vs. X% in non-HII"* |
| Opportunities | *"Opportunities exist to strengthen [Brand] discussions: [action], [action], and [action] each showed <X% performance"* |
| NPP / Omnichannel | *"[Brand] NPP reach [led/trailed] at X%; NPP-exposed HCPs showed X% higher LTIP — reinforcing the omnichannel impact"* |

**Rule 7 — Anti-patterns to reject**

| Anti-pattern | Fix |
|---|---|
| Writing the predicted hypothesis as a finding | Write what the data confirmed, not what was predicted |
| No data anchor | Add a specific % or pp value |
| "Results were mixed" | Pick the dominant direction; flag the exception |
| Missing brand or wave reference | Always name both |

---

### Phase 1 Output

```
## Slide Headlines — [Wave Label]

### [Domain Name]
[N]. [Headline]

### [Domain Name]
[N]. [Headline]

[etc. — one per domain/slide]
```

Save as `{project}/context/{wave}/slide_headlines.md`.

**→ Pause after Phase 1.** Ask: *"Do these headlines reflect the data correctly? Confirm to proceed to the Executive Summary."*

---

## PHASE 2 — EXECUTIVE SUMMARY

Write the ES from validated findings — not from hypothesis predictions. Apply the format specified (default: **A**).

### Format Menu

| Code | Name | Best used when |
|------|------|---------------|
| **A** | Narrative + Recommendations | Full quarterly readout; comprehensive findings |
| **B** | Combined Key Findings + Recs | Tight slide budget |
| **C** | Competitive Scorecard | Strong head-to-head dynamics |
| **D** | Findings \| Considerations Table | Structured, scannable client format |
| **E** | Strategic Questions | Exploratory; hypothesis-generating discussion |
| **F** | Strengths / Opportunities | Coaching-oriented; internal field force |
| **G** | What's Working / Opportunities / Recs | Mid-wave update; action-oriented |
| **H** | Key Takeaways Table | Leadership summary; minimal text |

---

### FORMAT A — Narrative + Recommendations (default)

For each domain (in order: Activity/SOV → Messaging → Rep Performance → Impact/LTIP → NPP → Non-rep roles if applicable):

```
**[DOMAIN NAME]**
• [Primary finding — validated data, specific % and QoQ delta]
  – [Supporting finding or segment nuance with data]
  – [Supporting finding or segment nuance with data]
  – [⚠ Methodology flag if applicable — italicized]
*[IMPLICATION — one interpretive sentence: what this means for field strategy, coaching, or competitive positioning. Not a restatement.]*
```

Close with methodology footnote:
```
*Note: [Primary Brand] [wave] n=[N], [Competitor] [wave] n=[N]. [Caveats: sample directional; SOV screener change; benchmark change; new message baselines.]*
```

---

### FORMAT B — Combined Key Findings + Recs

Left column — Key Findings (max 4–5, one per domain, bold label + finding with data anchor)
Right column — Recommendations (numbered, CAPS verb, one sentence each, mirrors domain order)

---

### FORMAT C — Competitive Scorecard

Table: Metric | [Primary Brand] [wave] | [Competitor] [wave] | Position

Directional indicator: ▲ improved | ▼ declined | → stable | `[NEW]` first wave
Position column: **LEADS** / **PARITY** / **TRAILS**
Bottom row: one-sentence competitive narrative.

---

### FORMAT D — Findings | Considerations

| What we found | What this means |
|---|---|
| [Specific finding with data] | [Implication or action] |

Max 6–8 rows, 1–2 sentences per cell.

---

### FORMAT E — Strategic Questions

3–5 numbered questions:
```
**[N]. [Strategic question]**
Evidence: [2–3 validated data points]
Hypothesis for next wave: [What to watch]
```

---

### FORMAT F — Strengths / Opportunities

**What's working:** CONFIRMED findings where primary brand leads, is stable, or improved
**Where to improve:** NOT CONFIRMED or PARTIALLY CONFIRMED findings; widening competitor gaps

---

### FORMAT G — Three-column

Column 1 — What's working (2–3 bullets with data)
Column 2 — Opportunities (2–3 bullets with data)
Column 3 — Recommendations (2–3 bullets, CAPS verb)

---

### FORMAT H — Key Takeaways

| Domain | Key Takeaway |
|--------|-------------|
| [Domain] | [Validated finding + data + implication in one sentence] |

---

**→ Pause after Phase 2.** Ask: *"Does the ES capture the right stories? Any findings to add, reorder, or drop before moving to recommendations?"*

---

## PHASE 3 — RECOMMENDATIONS

```
[N]. **[CAPS ACTION VERB]** [what to do — specific, named behavior or output]
[One sentence: the exact validated metric and delta that motivates this rec.]
```

**Approved CAPS verbs:** RESTORE, PRIORITIZE, EQUIP, ADDRESS, SUSTAIN, CONVERT, DEPLOY, BUILD, CLARIFY, EXTEND, PROTECT, REINFORCE, AMPLIFY, CLOSE, LEVERAGE, SHARPEN, ACCELERATE

Order by impact priority: highest-leverage first. 4–6 recommendations.

**Example:**
```
1. **RESTORE** [Primary Brand] NCCN message delivery among non-HII reps
[Competitor] recall reversed a Q4 gap, rising from 41% to 46% while [Primary Brand] declined from 49% to 43% — an 11pp swing in one quarter.
```

Append to `{project}/context/{wave}/exec_summary.md`.

---

## WRITING RULES (all phases)

**R1 — Phase 0: data facts only. Phase 1+: strategic insights.**
Phase 0 `Data summary` must be a pure numerical statement — no narrative, no "but", no interpretation.
Phase 1 headlines must answer "so what for the client" — not just "what happened in the data."

**R2 — Every Phase 1 headline must use project context.**
Before writing a headline, look up at least one of: competitive event, prior wave recommendation, client priority, strategic timing, or field intelligence from `project_context.md`. A headline with no context hook is incomplete.

**R3 — Data-anchored, always.** Every finding (Phase 1+) must include at least one specific % or pp value from validated_analysis.md.

**R4 — Wave-over-wave language.** Format: `[metric] [direction] from X% to Y% (+/-Npp)`. If no prior baseline: *"Wave 1 baseline — no prior comparison."*

**R5 — CAPS verbs in recs.** Every rec opens with a CAPS verb. Never passive.

**R6 — Recs motivated by data.** Name the specific metric and delta in the motivating sentence.

**R7 — Methodology flags mandatory.** `⚠` in italics when QoQ comparability is limited. Never buried when it affects a key finding.

**R8 — Always name both brands.** Never "the competitor." Always state both sides when comparing.

**R9 — Implications interpret, not restate.** Say what the finding *means* for the field — what's at risk, what decision it informs, what the client should prioritise.

**R10 — Segment splits add specificity.** Include HII/non-HII or Community/Academic splits when the data is there and when they change the strategic reading.

**R11 — New metrics: "Wave 1 baseline."** State direction vs. external reference only.

**R12 — No filler openers.** Lead with brand, metric, or finding. Never *"It is worth noting…"*, *"Interestingly…"*, *"Looking at…"*

---

## QUALITY CHECK (run silently before each phase output)

- [ ] Every finding sourced from validated_analysis.md — not from hypothesis predictions
- [ ] Every finding has a specific % or pp value
- [ ] Every QoQ comparison has prior + current + delta
- [ ] Every new metric labeled "Wave 1 baseline"
- [ ] Every methodology-affected finding carries ⚠
- [ ] Every rec opens with CAPS verb and names the motivating metric
- [ ] No competitor referred to as "the competition" — always named
- [ ] No filler openers
- [ ] Every context claim in a headline is traceable to a source file — no unverified causal assertions

---

## OUTPUT SEQUENCE

```
Phase 0 → {project}/context/{wave}/validated_analysis.md    (no user gate)
Phase 1 → {project}/context/{wave}/slide_headlines.md       ✋ User confirms
Phase 2 → {project}/context/{wave}/exec_summary.md          ✋ User confirms
Phase 3 → recs appended to exec_summary.md
        → Signal: ready for /slide-plan (Stage 4)
           Passes: validated_analysis.md + slide_headlines.md + exec_summary.md
```
