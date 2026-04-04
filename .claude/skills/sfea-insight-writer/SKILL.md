---
name: sfea-insight-writer
effort: max
description: "Use when writing data-validated analysis, slide headlines, executive summary, and recommendations for a PET SFEA wave. Trigger when: user says 'write insights', 'build ES', 'run insight writer', or Stage 2 hypothesis bank has been signed off and Stage 3 is ready to begin. Reads hypothesis_bank.md, project_context.md, and source_data.json. Runs in 4 phases: Phase 0 (auto) validates hypotheses vs data → validated_analysis.md; Phase 1 writes talking headlines → slide_headlines.md; Phase 2 writes ES → exec_summary.md; Phase 3 appends Recommendations."
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
    ctx = f'{proj}/context/{wave}' if wave else f'{proj}/context'
    for name, req in [('hypothesis_bank.md','Required'), ('project_context.md','Required'), ('source_data.json','Required'), ('validated_analysis.md','Output'), ('slide_headlines.md','Output'), ('exec_summary.md','Output')]:
        path = f'{ctx}/{name}'
        if os.path.exists(path):
            print(f'  ✓ {name} ({os.path.getsize(path)//1024}KB) [{req}]')
        else:
            print(f'  ✗ {name} — MISSING [{req}]')
    sj = f'{ctx}/source_data.json'
    if os.path.exists(sj):
        with open(sj) as f:
            idx = json.load(f)
        codes = idx.get('_codes',{})
        print(f'**source_data.json:** {sum(len(v) for v in codes.values())} codes indexed')
else:
    print('**No active project detected** — user must specify project folder')
"
`

# SFEA Insight Writer

Generates data-validated analysis and narrative threads (story arcs, slide headlines, executive summary, and recommendations) for any SFEA study wave. Invoked after the hypothesis bank is signed off.

---

## PIPELINE POSITION

```
Stage 0: index_excel()            →  source_data.json
Stage 1: /build-project-context   →  project_context.md
Stage 2: /hypotheses              →  hypothesis_bank.md
                                        ✋ User signs off
Stage 3: /sfea-insight-writer     ←  YOU ARE HERE
         Phase 0: Data Validation →  validated_analysis.md       (auto)
         Phase 1: Narrative       →  narrative_threads.md        ✋ SINGLE GATE
Stage 4: /slide-plan              ←  reads both outputs
Stage 5: config.yaml + generate_deck()
```

---

## STEP 1 — Confirm project and resolve file paths

Use the auto-detected project, wave, and file status shown above. If correct, proceed. If not, ask the user to specify.

The required files are:

| File | Path | Required? |
|------|------|-----------|
| **Hypothesis Bank** | `{project}/context/{wave}/hypothesis_bank.md` | Required |
| **Project Context** | `{project}/context/{wave}/project_context.md` | Required |
| **Source Data** | `{project}/context/{wave}/source_data.json` | Required |
| **Market Context** | `{project}/context/market_context.md` | Required |
| **ES Format** | Ask user — A through H; default **A** | Optional |

If any required file is missing (check auto-detected status above), stop and tell the user which is absent.

---

## STEP 2 — Read all files in full

Read all files completely before doing anything else.

**From `project_context.md` extract and hold:**
- Primary brand and competitor brand(s) — use exact names throughout; never "the competition"
- Wave label (current) and prior wave label — for QoQ framing
- Sample sizes — n= per brand per wave
- Methodology flags — survey changes, screener/benchmark changes affecting QoQ comparability
- Prior wave top findings and recommendations — for trend framing and CLOSURE arc detection
- Study design notes — HII definition, segment cuts used, new questions this wave
- Field intelligence / call notes — what reps reported, what the client flagged
- Client priorities / KBQs — what the client is trying to decide or act on

**From `market_context.md` extract and hold:**
- Competitive events — label updates, data readouts, formulary changes, competitor launches
- Clinical context — mechanism of action, treatment landscape, standard of care shifts
- Strategic timing — upcoming readouts, reviews, or decisions that create urgency

**From `hypothesis_bank.md` extract and hold:**
- All hypotheses with their domain, rationale, "Test with:" question codes, segment cuts, and flags
- Expected direction per hypothesis (the prediction)
- METHODOLOGY ARTIFACT and [ACTION ITEM] flags
- PRIOR WAVE VALIDATION hypotheses — these test whether prior recommendations are showing results

**From `source_data.json` extract and hold:**
- For each question code referenced in the hypothesis bank: prior value, current value, and delta (current − prior in pp)
- Segment splits where available (e.g., HII vs. non-HII, Community vs. Academic)
- If a code is in the `_sheets` index but not yet fully extracted: note it as "extraction needed" — do not skip it

---

## PHASE 0 — DATA VALIDATION

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

## PHASE 1 — NARRATIVE THREADS

This is the core editorial phase. It synthesizes validated data into story arcs, then writes headlines, executive summary, and recommendations — all in one integrated document. **This is the single human gate in Stage 3.**

Phase 1 has five steps executed in order:

```
Step A: Build findings matrix           (internal — not shown to user)
Step B: Detect patterns → story arcs    (the editorial backbone)
Step C: Write headlines per arc         (per-slide, arc-informed)
Step D: Write executive summary         (arc-organized, cross-domain)
Step E: Write recommendations           (arc-driven, impact-ordered)
```

---

### STEP A — Build Findings Matrix + KBQ Map

Create an internal working matrix from validated_analysis.md. Tag every CONFIRMED and PARTIALLY CONFIRMED hypothesis on five dimensions:

| Hypothesis | KBQ | Brand affected | Direction | Metric type | Segment pattern |
|---|---|---|---|---|---|
| H[N] | [KBQ #] | [Primary/Competitor/Both] | [Positive/Negative/Stable] | [Activity/Recall/Effectiveness/Quality/Intent/Closing/Tactics/NPP] | [Overall/Academic/Community/HII/non-HII] |

The **KBQ** column is the primary organizing key — it maps every finding back to the business question it answers. Use the hypothesis bank's domain assignment (each hypothesis was generated under a specific KBQ). Additional questions and cross-cutting hypotheses map to the KBQ they most directly inform.

Also tag:
- **Prior wave link**: Does this finding confirm or contradict a prior wave recommendation? (CLOSURE candidate)
- **Cross-domain echo**: Does this finding's direction match findings in other domains for the same brand? (CONVERGENCE candidate)

Do NOT show this matrix to the user. It is an analytical working tool for Step B.

---

### STEP B — Build KBQ-Governed Story Arcs

**Core principle: KBQs govern arc structure. Pattern types inform arc framing.**

The KBQs are the client's questions. The arcs are the answers. Each arc should be recognizable as answering a KBQ — not just describing a statistical pattern that happens to cross domains. Pattern types (CONVERGENCE, TENSION, DIVERGENCE, CLOSURE) are the editorial lens that shapes *how* the answer is told, not *what* the answer is about.

#### Two-pass process

**Pass 1 — KBQ-anchored grouping:**
Group all validated findings (CONFIRMED, PARTIALLY CONFIRMED, and surprising NOT CONFIRMED) by the KBQ they answer. Each KBQ with 3+ validated findings is an arc candidate.

If a KBQ has <3 findings: either merge it with a related KBQ that shares findings, or absorb its findings into a broader arc. Do not drop a KBQ — every KBQ in the hypothesis bank must be addressed, even if its answer is "insufficient data to conclude" or it merges into another arc.

If two KBQs share enough findings that they form a single coherent answer, merge them into one arc. Name both KBQs in the arc title or "KBQ anchor" field.

**Pass 2 — Pattern-type framing:**
For each KBQ-anchored arc, ask: *What is the shape of this KBQ's answer?*

| Pattern | When it applies to a KBQ answer | Signal |
|---|---|---|
| **CONVERGENCE** | Multiple metrics within this KBQ's scope all point the same direction | The answer is clear and reinforcing |
| **TENSION** | The KBQ's answer is "yes, but…" — positive on some metrics, negative on others | The answer reveals a strategic dilemma |
| **DIVERGENCE** | The KBQ's answer differs by segment (Academic vs Community, HII vs non-HII) | The answer is "depends where you look" |
| **CLOSURE** | The KBQ's answer directly validates or invalidates a prior wave recommendation | The loop between advice and outcome |

A single arc can carry multiple pattern types (e.g., TENSION + CLOSURE). Lead with the dominant one.

#### Governing Question as narrative spine

The study's **Governing Question** (from `KBQs.md`) is the spine that connects all arcs. Before writing arcs, state how the arc sequence answers the Governing Question:

```
**Governing Question:** [verbatim from KBQs.md]
**Arc sequence answer:** [1-2 sentences explaining how the arcs, in order, build toward answering the governing question]
```

The arc sequence should read as a logical progression: outcome → message strategy → competitive landscape → field execution → channel strategy. A reader should be able to read just the arc titles in order and understand the wave's story.

#### Rules for arc construction

1. **Each arc must be anchored to 1-2 KBQs.** The KBQ anchor is shown explicitly. An arc without a KBQ anchor is disallowed — if you can't name which business question it answers, it's a pattern observation, not a story arc.
2. **Each arc must be supported by 3+ hypotheses from at least 2 different metric types.** Cross-domain evidence still required within each KBQ answer.
3. **NOT CONFIRMED hypotheses that are surprising belong in the arc of the KBQ they were testing.** They inform the KBQ's answer — "we expected X but found Y" is part of the answer, not a separate story.
4. **Aim for 3-5 arcs.** KBQs can merge; not every KBQ needs its own arc. But every KBQ must be addressed.
5. **Every arc needs a "what's at stake" that connects to the client decision embedded in the KBQ.** The KBQ already names the decision — the "what's at stake" explains what the data means for that decision.
6. **Every CONFIRMED and PARTIALLY CONFIRMED hypothesis must map to at least one arc.** Orphans allowed only for genuinely isolated findings.
7. **Additional Questions / cross-cutting hypotheses** attach to whichever KBQ-anchored arc they most directly inform. Segment cuts (Academic vs Community, HII vs non-HII) are evidence within arcs, not arcs themselves.

#### Arc structure

```
## Thread [N]: [Strategic claim — answers the KBQ, names brands and direction]
**KBQ anchor:** KBQ [N] — [KBQ title from KBQs.md]. Also addresses: [Additional Q #s if any]
**Pattern:** CONVERGENCE / TENSION / DIVERGENCE / CLOSURE
**Urgency:** ACT NOW / MONITOR / CELEBRATE
**Evidence:**
- H[x] ([KBQ domain]): [one-line data summary from validated_analysis.md]
- H[y] ([KBQ domain]): [one-line data summary]
- H[z] ([Additional Q / cross-cutting]): [one-line data summary]
**What's at stake:** [What the data means for the decision embedded in the KBQ — 1-2 sentences.
  Must reference the specific client priority or competitive threat from the KBQ itself,
  grounded in project_context.md or market_context.md.]
**Methodology caveats:** [Any ⚠ flags that affect findings in this arc] ← omit if none
```

#### Example arcs (KBQ-governed)

*KBQ 1 answer — TENSION:*
> **Thread 1: The promotional engine IS converting — but only when interactions reach high-impact quality**
> KBQ anchor: KBQ 1 — Prescribing Conversion. Also addresses: Additional Q6 (2x2 segmentation)
> Pattern: TENSION | Urgency: ACT NOW
> Evidence:
> - H2 (Closing): Branded close +7pp to 53%, overtaking TAG (44%) — conversion IS improving
> - H5 (Intent): HII allocation 4.9/10 vs Others 3.4/10 — but concentrated in HII
> - H35 (Quality): Call quality declined -6pp to 76% — fewer interactions reaching HII threshold
> - H40 (Intent): LTIP flat at 70% while TAG+Chemo surged to 79%
> What's at stake: KBQ 1 asks whether conversion is improving. The answer is "yes, where interactions reach HII quality — but overall quality is declining, so fewer interactions reach that threshold." This is a quality problem, not a conversion problem.

*KBQ 2+3 merged — TENSION:*
> **Thread 2: The message hierarchy is shifting to address real barriers — but OS is stalling and SubQ differentiation lags**
> KBQ anchor: KBQ 2 — Message-to-Barrier Mapping + KBQ 3 — FASPRO SubQ Penetration
> Pattern: TENSION | Urgency: ACT NOW
> Evidence:
> - H10 (Recall): COCOON surged +15pp to #1 — safety barrier being addressed
> - H15 (SubQ): Three SubQ messages debuted at 30-35% — infusion barrier being addressed
> - H7 (Recall): But OS Headline flat at 36% despite being first VA tab — the strongest differentiator isn't gaining traction
> - H16 (Topics): Route of Admin +14pp to 25% — but only 1 in 4 interactions, well below "all"
> What's at stake: KBQ 2 asks whether messages map to barriers. They increasingly do (COCOON → safety, SubQ → infusion). But KBQ 3 asks whether FASPRO is penetrating — at 25% and with low Differentiation scores, the convenience advantage isn't yet a switching reason.

*KBQ 4 answer — CONVERGENCE:*
> **Thread 3: AZ's restructured force is delivering coordinated gains — quality, effectiveness, and intent all moved in TAG's favor**
> KBQ anchor: KBQ 4 — Competitive Resilience
> Pattern: CONVERGENCE | Urgency: ACT NOW
> Evidence:
> - H33 (Quality): TAG call quality +11pp to 84%, overtaking RYB 76%
> - H21/H23 (Recall): TAG NCCN +5pp to #1 at 46%; Tolerability +11pp to 41%
> - Supplementary: TAG ME +7.1pp to parity; TAG+Chemo LTIP +13pp to 79%
> What's at stake: KBQ 4 asks whether J&J can withstand AZ's counter-offensive. AZ added 40 reps, split into 3 teams, and the Q1 data shows this investment paying off across every dimension simultaneously.

*KBQ 5 answer — DIVERGENCE + CLOSURE:*
> **Thread 4: Interaction architecture improvements are landing — but unevenly, with community coaching proving out while academic engagement erodes**
> KBQ anchor: KBQ 5 — Interaction Architecture. Also addresses: Additional Q1 (Acad vs Comm), Q7 (rep-led narrative)
> Pattern: DIVERGENCE + CLOSURE | Urgency: ACT NOW
> Evidence:
> - H27 (Tactics): VA usage surged +12pp to 56% — Rec 3 "LEVERAGE VAs" actioned [CLOSURE]
> - H41 (Segment): Community VA +18pp to 59% vs Academic -2pp to 50% [DIVERGENCE]
> - H11 (Recall): Academic CNS recall dropped -12pp; community held
> - Supplementary: Academic LTIP collapsed -9pp to 69%; community improved +3pp to 71%
> What's at stake: KBQ 5 asks whether reps are structuring for impact. In community, yes — the Q4 VA recommendation produced the largest tactical gain this wave. In academic, the architecture is intact but the content isn't converting — possibly because FLAURA-2's intensification narrative resonates more in clinical-trial-aware academic settings.

---

### STEP C — Write Headlines Per Arc

For each arc, write headlines for the slides that will serve it. Each headline is **specific to the data that will appear on one slide** — it does NOT reference data from other slides or other domains.

**The arc's role is to inform framing, not content.** The arc tells the headline writer:
- What *kind of signal* this data represents (systematic shift? isolated win? coaching proof?)
- What *tone* to strike (urgent? celebratory? diagnostic?)
- What *interpretive lens* to apply (competitive threat? field execution? strategic positioning?)

The headline itself talks only about the metrics, brands, and numbers that belong on its slide.

#### Headline Structure — Insight-First, Data-Light

Headlines lead with the **strategic insight** — what the data means for the client's decision. Data appears as a parenthetical proof point, not as the headline structure. The chart shows the numbers; the headline tells you what they mean.

Every headline has two parts:

```
[Part 1 — Strategic insight: what happened and why it matters, naming brands and direction]
[Part 2 — One parenthetical data anchor as proof — (X%, +Npp) — embedded naturally in the sentence]
```

Use a **semicolon (;)** for additive or parallel findings within the slide.
Use an **em-dash (—)** when Part 2 is a sharp contrast, pivot, or implication.

#### Headline Anatomy

```
[Strategic claim about what the metric movement means for the client]
— [implication or contrast, with one data point as proof]
```

The headline should read like a field strategy observation, not a data readout. A reader who never sees the chart should understand the business implication. A reader who only sees the chart should find the headline adds interpretive value the chart alone doesn't provide.

#### Good vs. Bad Headlines — The Litmus Tests

**Litmus test 1 — Data narration vs. strategic insight:**

Same data, three framings:

*Data narration (wrong — reads like a spreadsheet):*
> "TAG average ME surged from 62.6% to 69.6% (+7.0pp), nearly closing the effectiveness gap with RYB+LAZ (70.0%)"

*Data-reduced but still data-structured (wrong — numbers still drive the sentence):*
> "TAG ME surged +7pp to near-parity with RYB (69.6% vs 70.0%) — broad-based gains across 9 of 10 messages suggest AZ's restructured messaging playbook is resonating"

*Insight-first (right — strategic claim leads, one data anchor as proof):*
> "AZ's restructured messaging playbook is resonating broadly — TAG ME improved across 9 of 10 messages to reach near-parity with RYB, suggesting a systematic messaging shift rather than isolated message improvement"

The difference: the third headline tells you what the data *means* for the competitive landscape. The numbers are in the chart — the headline adds the interpretive layer.

**Litmus test 2 — Arc leakage:**

*Arc leakage (wrong — references data not on this slide):*
> "TAG's ME surge to parity is the messaging dimension of a broader competitive shift — quality, intent, and now effectiveness all moved in AZ's favor in Q1"

*Arc-informed but slide-specific (right):*
> "AZ's restructured messaging playbook is resonating broadly — TAG ME improved across 9 of 10 messages to reach near-parity with RYB, suggesting a systematic messaging shift rather than isolated message improvement"

The difference: the first explicitly names quality and intent data that isn't on the ME slide. The second uses the arc's *interpretive lens* (this is systematic, not random) without importing other slides' data.

**Litmus test 3 — Missing "so what":**

*Missing "so what" (wrong — stops at the number):*
> "RYB+LAZ branded closing improved from 46% to 53% (+7pp) in Q1'26, nearly matching TAG Chemo's 55%"

*Has "so what" (right):*
> "Q4's coaching emphasis on closing technique is paying off — RYB overtook TAG on branded close for the first time, with the gain concentrated in high-impact interactions where reps execute the full interaction sequence"

#### Context Anchoring

Before writing each headline, look up the relevant context:

| Question | Where to find it |
|---|---|
| *Why might this metric be moving?* | `market_context.md` — competitive events, label updates, campaign launches |
| *What did reps report doing differently?* | `project_context.md` — field intelligence / call notes |
| *Was this flagged as a risk or priority last wave?* | `project_context.md` — prior wave findings + recommendations |
| *What is the client trying to decide or act on?* | `project_context.md` — KBQs / client priorities |
| *What arc does this slide serve?* | Step B output — the pattern type and "what's at stake" |

#### Headline Rules

**Rule 1 — Strategic insight, not data narration**
- Use data as evidence for an insight about what this metric movement means
- Connect to something the client cares about: a competitive threat, coaching priority, timing risk
- Never write a headline that only describes what the numbers did
- For NOT CONFIRMED hypotheses: write what the data *actually* showed and why it matters

**Rule 1a — Every context claim must be file-sourced, not inferred**

| Claim type | Required source |
|---|---|
| A metric "drives" prescribing / LTIP | `source_data.json` driver analysis OR `market_context.md` |
| A competitive event explains a movement | `project_context.md` — field intelligence or call notes |
| A strategic deadline or timing risk | `project_context.md` — project timeline or client priorities |
| A clinical or label claim | `market_context.md` — documented product/disease context |
| General pharma inference not in any file | Not permitted — remove or soften to what data alone supports |

**Rule 1b — No arc leakage in headlines**
- A headline must only reference data that will appear on its slide
- The arc informs *how to interpret* the slide's data, not *what other data to mention*
- If you find yourself writing "alongside quality gains" or "compounding the LTIP shift" on a messaging slide, you've leaked — rewrite using only the messaging data

**Rule 2 — Direction vocabulary**

| Direction | Preferred words |
|---|---|
| Positive | maintained, sustained, strengthened, improved, delivered, exceeded, outperformed, led |
| Stable/Parity | remained consistent, remained stable, on par with, broadly similar |
| Negative | declined, dropped, lagged, trailed, fell short, showed a downward trend |
| Opportunity | *"opportunity exists to [verb]"* / *"[metric] remains an area for focus"* |

**Rule 3 — Always anchor to wave and brand**
- Reference the current wave explicitly: *"In Q1'26…"* or *"in Q1'26"*
- Name both brands when comparing — never "the competitor"

**Rule 4 — Data-light: the chart shows the numbers, the headline shows the meaning**
- Headlines are insight-first. Lead with the strategic claim, not the data point.
- Include **at most 1-2 numbers** as parenthetical proof — never more. The chart is right there.
- If you find yourself writing three or more numbers in a headline, you're narrating data, not interpreting it. Remove numbers until the insight leads.
- Use `~` for approximate, `%` not "percent", `pp` for percentage points.
- Acceptable: "RYB overtook TAG on branded close for the first time (+7pp to 53%)" — one proof point.
- Not acceptable: "RYB branded closing surged +7pp to 53% in Q1'26, overtaking TAG (44%, -11pp) — the 18pp swing validates coaching" — five numbers, reads like a ticker tape.

**Rule 5 — "Directionally" for small n**
- When n is small or delta is borderline: *"Directionally, [Brand] [finding]…"*

**Rule 6 — Anti-patterns to reject**

| Anti-pattern | Fix |
|---|---|
| Writing the predicted hypothesis as a finding | Write what the data confirmed, not what was predicted |
| No data anchor at all | Add one specific % or pp value as parenthetical proof |
| **Data-heavy headline (3+ numbers)** | **Strip to 1-2 numbers max. The chart shows the rest. If the headline reads like a data readout, rewrite it as a strategic claim with one proof point.** |
| "Results were mixed" | Pick the dominant direction; flag the exception |
| Missing brand or wave reference | Always name both |
| Headline references data from another slide | Remove — interpret this slide's data through the arc lens instead |
| Headline is just Part 1 (data) with no Part 2 (meaning) | Flip the structure: lead with meaning, use data as proof |
| **Headline structured as "[metric] [went] from X% to Y% (+Zpp)"** | **Restructure: lead with what this movement means for the client; embed one number as evidence** |

---

### STEP D — Write Executive Summary

The ES is organized by **KBQ-anchored story arcs**, not by domains or metric types. Each arc becomes a section, pulling cross-domain evidence that answers the KBQ's business question.

This is where cross-domain connection happens. Unlike headlines (which stay slide-specific), the ES explicitly names evidence from multiple metric types under each KBQ-anchored arc.

**The ES is the narrative layer that connects the slides into a coherent story.** A reader who only reads the ES should understand the 3-5 things that matter this wave and why — and should be able to map each section back to the business question it answers.

Open the ES with the **Governing Question** and a 1-2 sentence framing of how the arcs answer it. This orients the reader before the arc sections begin.

Apply the format specified by the user (default: **A**).

#### Format Menu

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

#### FORMAT A — Narrative + Recommendations (default)

For each arc (ordered by urgency: ACT NOW first, then MONITOR, then CELEBRATE):

```
**[ARC TITLE — the strategic claim, not a domain name]**
• [Primary evidence — cross-domain finding with specific % and QoQ delta]
  – [Supporting evidence from a different domain — with data]
  – [Supporting evidence or segment nuance — with data]
  – [⚠ Methodology flag if applicable — italicized]
*[IMPLICATION — what this pattern means for field strategy, coaching, or competitive positioning.
  This is NOT a restatement of the bullets. It answers: "Given this pattern, what should the client
  prioritize, protect, or change?" Connect to a specific client decision or upcoming event.]*
```

**Critical difference from domain-organized ES:** Each section pulls evidence from whichever domain supports the arc. A section titled "AZ's restructured sales force is delivering" would include a quality metric (rep performance domain), an ME metric (messaging domain), and an LTIP metric (intent domain) — because the *story* is the convergence, not the individual metrics.

Close with methodology footnote:
```
*Note: [Primary Brand] [wave] n=[N], [Competitor] [wave] n=[N]. [Caveats: sample directional; screener change; benchmark change; new message baselines.]*
```

---

#### FORMAT B — Combined Key Findings + Recs

Left column — Key Findings (one per arc, bold arc title + cross-domain evidence with data anchor)
Right column — Recommendations (numbered, CAPS verb, one sentence each, mirrors arc order)

---

#### FORMAT C — Competitive Scorecard

Table: Metric | [Primary Brand] [wave] | [Competitor] [wave] | Position

Directional indicator: ▲ improved | ▼ declined | → stable | `[NEW]` first wave
Position column: **LEADS** / **PARITY** / **TRAILS**
Bottom section: one paragraph per arc explaining the pattern behind the metrics.

---

#### FORMAT D — Findings | Considerations

| What we found (arc-organized) | What this means |
|---|---|
| [Cross-domain evidence under arc title] | [Implication or action] |

Max 6–8 rows (one per arc + key orphans), 1–2 sentences per cell.

---

#### FORMAT E — Strategic Questions

3–5 numbered questions (one per arc):
```
**[N]. [Strategic question — framed by the arc]**
Evidence: [2–3 cross-domain validated data points]
Hypothesis for next wave: [What to watch]
```

---

#### FORMAT F — Strengths / Opportunities

**What's working:** CELEBRATE and positive CLOSURE arcs — where primary brand leads, improved, or coaching landed
**Where to improve:** ACT NOW arcs — where competitive gaps are widening or quality is declining

---

#### FORMAT G — Three-column

Column 1 — What's working (CELEBRATE arcs, 2–3 bullets with cross-domain data)
Column 2 — Opportunities (ACT NOW arcs, 2–3 bullets with cross-domain data)
Column 3 — Recommendations (1 per arc, CAPS verb)

---

#### FORMAT H — Key Takeaways

| Arc | Key Takeaway |
|-----|-------------|
| [Arc title] | [Cross-domain evidence + implication in one sentence] |

---

### STEP E — Write Recommendations

Recommendations are **arc-driven and KBQ-anchored**. Each rec addresses a story arc (which answers a KBQ), not an individual metric. One rec may span findings from multiple metric types within the same KBQ answer.

```
[N]. **[CAPS ACTION VERB]** [what to do — specific, named behavior or output]
[One sentence: the cross-domain evidence pattern that motivates this rec, with specific metrics and deltas.]
```

**Approved CAPS verbs:** RESTORE, PRIORITIZE, EQUIP, ADDRESS, SUSTAIN, CONVERT, DEPLOY, BUILD, CLARIFY, EXTEND, PROTECT, REINFORCE, AMPLIFY, CLOSE, LEVERAGE, SHARPEN, ACCELERATE, INVESTIGATE, MONITOR

**Rules:**
- One recommendation per arc (occasionally two if the arc has distinct action streams)
- Order by urgency: ACT NOW arcs first, then MONITOR, then CELEBRATE/EXTEND
- 4-6 recommendations total
- The motivating sentence must name evidence from at least 2 domains — this is what makes arc-driven recs sharper than metric-specific recs
- CELEBRATE arcs get "SUSTAIN" or "EXTEND" recs, not just congratulations

**Example — arc-driven rec vs. metric-specific rec:**

*Metric-specific (old approach):*
> 1. **RESTORE** RYB+LAZ call quality to regain competitive parity with TAG
> TAG surpassed RYB on overall call quality (84% vs 76%, a +11pp swing) and now leads on 8 of 12 attributes.

*Arc-driven (new approach):*
> 1. **RESTORE** competitive parity before AZ's restructured sales force fully matures
> TAG gained simultaneously on call quality (+11pp to 84%), messaging effectiveness (+7pp to parity), and prescribing intent (+9pp, narrowing the gap to 6pp) — this is a systematic competitive improvement, not an isolated metric shift, and the response must address quality, messaging, and intent together rather than treating them as separate coaching workstreams.

The difference: the arc-driven rec names the *pattern* and prescribes a *coordinated response*. The metric-specific rec addresses one symptom.

---

### ORPHAN FINDINGS

After mapping all hypotheses to arcs, list any CONFIRMED or PARTIALLY CONFIRMED findings that don't fit any arc:

```
## Orphan Findings
These validated findings do not belong to a narrative thread but warrant standalone slide headlines:

### H[N]: [Hypothesis]
**Data:** [key metric from validated_analysis.md]
**Headline:** [standalone headline — still follows all headline rules]
**Note:** [Why this is isolated — e.g., "single metric, no cross-domain echo"]
```

Orphans are legitimate. Not every finding connects to a bigger story. But if you have more than 5 orphans, revisit your arcs — you may have missed a pattern.

---

## OUTPUT — narrative_threads.md

Assemble all Phase 1 output into a single document:

```
# Narrative Threads — [Wave Label]

**Generated:** [date]
**Source:** validated_analysis.md + project_context.md + market_context.md + hypothesis_bank.md
**ES Format:** [A/B/C/D/E/F/G/H]
**Arcs:** [N] threads | **Headlines:** [N] total | **Recommendations:** [N]

---

## GOVERNING QUESTION

**[Governing Question — verbatim from KBQs.md]**
**Arc sequence answer:** [1-2 sentences: how the arcs, in order, answer this question]

---

## STORY ARCS

### Thread [N]: [Strategic claim — answers the KBQ]
**KBQ anchor:** KBQ [N] — [title]. Also addresses: [Additional Q #s if any]
**Pattern:** [CONVERGENCE / TENSION / DIVERGENCE / CLOSURE]
**Urgency:** [ACT NOW / MONITOR / CELEBRATE]
**Evidence:**
- H[x] ([domain]): [one-line data summary]
- H[y] ([domain]): [one-line data summary]
- H[z] ([domain]): [one-line data summary]
**What's at stake:** [1-2 sentences — what data means for the KBQ's embedded decision]
**Methodology caveats:** [⚠ flags if any]

#### Headlines
[N]. **[Headline text]**
    Hypotheses: H[x], H[y]
    Slide data: [which metrics/brands appear on this slide]

[N]. **[Headline text]**
    Hypotheses: H[z]
    Slide data: [which metrics/brands appear on this slide]

[repeat for each headline under this arc]

---

[repeat for each thread]

---

## Orphan Findings

### H[N]: [Hypothesis]
**Data:** [key metric]
**Headline:** [standalone headline]

---

## Executive Summary

[Full ES in the selected format — arc-organized]

---

## Recommendations

[Numbered recs — arc-driven, impact-ordered]

---

*End of Stage 3 — narrative_threads.md*
*Ready for /slide-plan (Stage 4)*
*Outputs: validated_analysis.md + narrative_threads.md*
```

Save as `{project}/context/{wave}/narrative_threads.md`.

**→ SINGLE HUMAN GATE.** Present to the user:

```
Narrative Threads — [Wave] summary:

Governing Question: [abbreviated]

Arcs:
  [N]. [Arc title] ([pattern type] — [urgency])
       KBQ anchor: KBQ [N] — [title]
       Evidence: [count] hypotheses across [count] metric types
       Headlines: [count]

  [repeat]

KBQ coverage: [list each KBQ and which arc addresses it]
Orphans: [count] standalone findings

ES format: [code] — [name]
Recommendations: [count]

Hypothesis coverage: [X] of [Y] CONFIRMED/PARTIALLY CONFIRMED hypotheses mapped to arcs
Unmapped: [list any gaps]
```

Ask: *"Do these story arcs capture the right narrative? Review the threads, headlines, ES, and recs. Confirm to proceed to /slide-plan, or tell me what to reshape."*

---

## WRITING RULES (all phases)

**R1 — Phase 0: data facts only. Phase 1: strategic insights.**
Phase 0 `Data summary` must be a pure numerical statement — no narrative, no "but", no interpretation.
Phase 1 arcs, headlines, ES, and recs must answer "so what for the client" — not just "what happened."

**R2 — Every headline must use project context.**
Before writing a headline, look up at least one of: competitive event, prior wave recommendation, client priority, strategic timing, or field intelligence from `project_context.md` or `market_context.md`. A headline with no context hook is incomplete.

**R3 — Data-anchored, always.** Every finding must include at least one specific % or pp value from validated_analysis.md.

**R4 — Wave-over-wave language.** Format: `[metric] [direction] from X% to Y% (+/-Npp)`. If no prior baseline: *"Wave 1 baseline — no prior comparison."*

**R5 — CAPS verbs in recs.** Every rec opens with a CAPS verb. Never passive.

**R6 — Recs motivated by cross-domain evidence.** Name evidence from at least 2 domains in the motivating sentence.

**R7 — Methodology flags mandatory.** `⚠` in italics when QoQ comparability is limited.

**R8 — Always name both brands.** Never "the competitor."

**R9 — Implications interpret, not restate.** Say what the finding *means* for the field — what's at risk, what decision it informs, what the client should prioritize.

**R10 — Segment splits add specificity.** Include HII/non-HII or Community/Academic splits when they change the strategic reading.

**R11 — New metrics: "Wave 1 baseline."** State direction vs. external reference only.

**R12 — No filler openers.** Lead with brand, metric, or finding. Never *"It is worth noting…"*, *"Interestingly…"*, *"Looking at…"*

**R13 — Headlines stay on their slide.** No referencing data from other slides. The arc informs interpretation, not content.

**R14 — ES is the cross-domain layer.** The ES is where arcs explicitly pull evidence from multiple domains. Headlines do not do this.

---

## QUALITY CHECK (run silently before output)

**Arcs:**
- [ ] 3-5 arcs identified
- [ ] Every arc has a KBQ anchor — no arc exists without naming the KBQ it answers
- [ ] Every KBQ is addressed by at least one arc (merged or standalone)
- [ ] Governing Question stated with arc-sequence answer
- [ ] Each arc has 3+ hypotheses from 2+ metric types
- [ ] Each arc has a "what's at stake" tied to the client decision embedded in its KBQ
- [ ] Every CONFIRMED/PARTIALLY CONFIRMED hypothesis maps to an arc or is listed as orphan
- [ ] No more than 5 orphans (if more, revisit arcs)
- [ ] Arc titles read as answers to the KBQs, not just data pattern descriptions

**Headlines:**
- [ ] Every headline has Part 1 (data) + Part 2 (meaning)
- [ ] No headline references data from another slide (no arc leakage)
- [ ] Every headline has at least one specific % or pp value
- [ ] Every headline uses project/market context for interpretation
- [ ] No filler openers
- [ ] Both brands named when comparing

**ES:**
- [ ] Organized by arcs, not by domains
- [ ] Each section pulls cross-domain evidence
- [ ] Implications interpret the *pattern*, not individual metrics
- [ ] Methodology footnote present

**Recs:**
- [ ] One rec per arc (4-6 total)
- [ ] Each rec opens with CAPS verb
- [ ] Each motivating sentence names evidence from 2+ domains
- [ ] Ordered by urgency (ACT NOW first)
- [ ] CELEBRATE arcs have SUSTAIN/EXTEND recs

---

## OUTPUT SEQUENCE

```
Phase 0 → {project}/context/{wave}/validated_analysis.md       (auto — no gate)
Phase 1 → {project}/context/{wave}/narrative_threads.md        ✋ SINGLE GATE
        → Signal: ready for /slide-plan (Stage 4)
           Passes: validated_analysis.md + narrative_threads.md
```
