---
name: prior-wave-context
effort: high
description: "Pre-Stage 1 sub-skill. Scans the project input wave folder for any files containing prior wave findings (prior wave .pptx reports, prior_wave_es.md, readout decks, executive summaries), confirms with the user which files to use, extracts content from all confirmed files, and writes a structured prior_wave_context.md to the wave context folder. This file is then consumed by /build-project-context (Section 3) and /hypotheses. Run this before /build-project-context whenever prior wave files exist in the input folder. Trigger when: user says 'build prior wave context', 'extract prior wave findings', or when scanning a new wave folder that contains .pptx report files or prior_wave_es files."
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
    for md in sorted(glob.glob(f'{ctx}/*.md')):
        print(f'  - \`{os.path.basename(md)}\` ({os.path.getsize(md)//1024}KB)')
    # Check input folder for prior wave files
    for case in ['Wave', 'wave']:
        inp = f'{proj}/input/{case}/{wave}'
        if os.path.isdir(inp):
            files = [f for f in os.listdir(inp) if not f.startswith('.')]
            print(f'**Input files ({len(files)}):** ' + ', '.join(sorted(files)[:10]))
            break
else:
    print('**No active project detected** — user must specify project folder')
"
`

# Prior Wave Context Extractor

You are a ZoomRx senior analyst. Your job is to find, confirm, and extract all prior wave findings from whatever files exist in the project input folder — regardless of file type or naming convention — and produce a single, structured `prior_wave_context.md` that downstream skills can reliably consume.

This skill runs **before** `/build-project-context`. Its output feeds directly into:
- `/build-project-context` → Section 3 (Wave Hypotheses) and Section 4 (Analytical Priorities)
- `/hypotheses` → Prior wave baseline for each hypothesis

---

## STEP 0 — Confirm project and wave

Use the auto-detected project and wave shown above. If correct, proceed. If not, ask the user to specify.

Derive:
- **INPUT_DIR** = `{project_folder}/input/Wave/{wave}` (check both `Wave` and `wave` casing)
- **OUTPUT_PATH** = `{project_folder}/context/{wave}/prior_wave_context.md`

---

## STEP 1 — Scan and detect candidate prior wave files

Run:
```bash
find "{INPUT_DIR}" -type f | sort
```

For each file, evaluate whether it is a **prior wave source** using this signal table:

| Signal | Examples | Likely prior wave? |
|--------|----------|-------------------|
| `.pptx` + quarter/year in name earlier than current wave | `Q4'25 Report.pptx`, `Q3 2025 Readout.pptx`, `PET Q2 2025 Final.pptx` | **Yes — full report** |
| `.pptx` + `report`, `readout`, `findings`, `ES`, `executive` in name | `RYB Readout Final.pptx` | **Yes — full report** |
| `.md` or `.txt` + `prior`, `previous`, `ES`, `executive`, `summary`, `wave` | `Prior_Wave_ES.md`, `Q4_summary.md` | **Yes — narrative summary** |
| `.docx` + `readout`, `findings`, `debrief`, `summary` | `Q4 Client Debrief.docx` | **Yes — client notes** |
| `source_data.xlsx` | — | **No — skip** |
| `KBQs.md`, `survey_context`, `call_notes` | — | **No — skip** |
| `template.pptx` (in templates folder) | — | **No — skip** |
| Current wave's own `.pptx` output | filename matches current wave name exactly | **No — skip** |

---

## STEP 2 — Confirm with the user (one confirmation only)

Present a table of all detected candidate files. Ask the user to confirm once before any extraction begins:

```
I found the following files that appear to contain prior wave findings:

  File                                           | Type          | Role detected
  ───────────────────────────────────────────────┼───────────────┼────────────────────────
  JJ PET RYBREVANT+LAZCLUZE Q4'25 Report.pptx   | Full deck     | Prior wave report (rich)
  Prior_Wave_ES.md                               | Text summary  | Prior wave ES (narrative)

These files will be used to build prior_wave_context.md for {wave}.

Confirm to proceed, or:
- Remove a file from the list
- Add a file I missed (provide path)
- Tell me which wave the report covers if not obvious from the filename
```

**Wait for user confirmation before extracting anything.** Do not proceed on assumed consent.

Once confirmed, note:
- Which file is the **primary source** (full deck if available, otherwise the richest file)
- Which files are **supplementary** (ES summary, client debrief, etc.)
- What **prior wave name** to use in the output (e.g., "Q4 2025" — infer from filename or ask)

---

## STEP 3 — Extract content from each confirmed file

Use the appropriate extractor per file type. Extract all files before synthesizing.

### `.pptx` — slide-by-slide text extraction

```python
import sys
from pptx import Presentation

path = r"{FILE_PATH}"
prs = Presentation(path)
print(f"Total slides: {len(prs.slides)}")
for i, slide in enumerate(prs.slides, 1):
    texts = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                t = para.text.strip()
                if t:
                    texts.append(t)
    if texts:
        print(f"n--- Slide {i} ---")
        print("n".join(texts))
```

Run as:
```bash
python -c "<above code with FILE_PATH substituted>"
```

**Note on special characters in filenames** (apostrophes, spaces): use a raw string `r"..."` for the path and wrap the entire python command in double quotes on the shell. If the file has an apostrophe (e.g., `Q4'25`), pass the path via a temp variable:

```bash
python -c "
path = r'{FILE_PATH}'
from pptx import Presentation
prs = Presentation(path)
for i, slide in enumerate(prs.slides, 1):
    texts = [p.text.strip() for s in slide.shapes if s.has_text_frame for p in s.text_frame.paragraphs if p.text.strip()]
    if texts:
        print(f'--- Slide {i} ---')
        print('\n'.join(texts))
"
```

When reading a full prior wave `.pptx` deck, pay attention to:
- **Slide titles / section headers** → domain classification (Messaging, Rep Performance, etc.)
- **Headline text boxes** (usually large, top of slide) → key finding per slide
- **Bullet points** → supporting evidence
- **Data labels / callouts** → specific metric values (percentages, QoQ changes)
- **Footer / annotation text** → methodology notes, base sizes, wave labels

### `.md` / `.txt` — read directly

Use the Read tool. No conversion needed.

### `.docx` — extract via python-docx

```bash
python -c "
from docx import Document
doc = Document(r'{FILE_PATH}')
for para in doc.paragraphs:
    if para.text.strip():
        print(para.text)
for table in doc.tables:
    for row in table.rows:
        cells = [c.text.strip() for c in row.cells if c.text.strip()]
        if cells:
            print(' | '.join(cells))
"
```

### `.pdf` — extract via pdfplumber

```bash
python -c "
import pdfplumber
with pdfplumber.open(r'{FILE_PATH}') as pdf:
    for i, page in enumerate(pdf.pages, 1):
        text = page.extract_text()
        if text and text.strip():
            print(f'--- Page {i} ---')
            print(text)
"
```

---

## STEP 4 — Synthesize prior_wave_context.md

Using all extracted content, write `prior_wave_context.md` to `{OUTPUT_PATH}`.

Apply the following structure exactly. Populate every section from the extracted content — do not leave sections empty. If a domain genuinely has no data in the source files, write `No data found in source files for this domain.`

```markdown
# Prior Wave Context — {prior_wave_name}
**Product:** {product/brand name inferred from files}
**Wave covered:** {prior wave name, e.g. Q4 2025}
**Source files:**
- {file 1 name} ({file type} — {role: primary/supplementary})
- {file 2 name} ({file type} — {role})
**Extracted:** {today's date}

---

## 1. Key Findings by Domain

For each domain found in the source files, write one block:

### {Domain Name}  (e.g., Sales Activity & Share of Voice)

**Headline:** {the single most important finding for this domain — as it would appear on a slide headline}

**Supporting findings:**
- {finding with specific value if available, e.g., "J&J SOV: 58% vs AZ 42% (Q4 2025)"}
- {finding}

**Trend vs. prior wave:** {↑ Improving / ↓ Declining / → Stable / ~ Mixed} — {brief rationale}

**Segment split (if available):**
- Academic: {finding}
- Community: {finding}

---

Repeat for each domain. Common domains for PET studies (include only those present in source):
- Sales Activity & Share of Voice
- Message Recall & Delivery
- Message Effectiveness (MBD — Motivation / Believability / Differentiation)
- Rep Performance & Interaction Quality
- Visual Aid Usage & Impact
- Call-to-Action & Branded Close
- Long-Term Intent to Prescribe (LTIP)
- Impact & Prescription Intent
- High Impact Interactions (HII)
- NPP Channel Recall & Effectiveness
- Market Perceptions & Product Attribute Ratings

---

## 2. Metrics Snapshot

Capture every specific numeric value found in the source files:

| Metric | Prior Wave Value | vs. Wave Before | Trend |
|--------|-----------------|-----------------|-------|
| {metric name} | {value, e.g. 42%} | {delta if shown, e.g. +5pp} | ↑/↓/→ |

If values are not available from the source, write "Not reported in source files."

---

## 3. Segment-Specific Highlights

### Academic vs. Community
{Key differences in performance between academic and community HCPs — drawn from source}

### High Impact Interactions
{What drove HII designation; how HII interactions performed vs. standard}

### Other Segments (if reported)
{Rep-led vs. HCP-led, VA users vs. non-users, 1L-first vs. other, etc.}

---

## 4. Open Action Items from Prior Wave

Items the prior readout explicitly flagged for follow-up, monitoring, or action:

- [ ] {Action item — be specific, e.g., "Strengthen community rep narrative control around 1L setting"}
- [ ] {Action item}

If the source files contain a recommendations section, extract every recommendation verbatim or near-verbatim.

---

## 5. Recommendations Carried Forward

Recommendations from the prior wave readout that should be tested or tracked in {current wave}:

- **{PRIORITY VERB}** {recommendation, e.g., "PRIORITIZE 1L setting discussion among community HCPs to strengthen narrative control"}
- **{PRIORITY VERB}** {recommendation}

---

## 6. Unanswered Questions / Gaps

Questions raised in the prior wave that were not conclusively answered, or new questions the data surfaced:

- {question}
- {question}

If none explicitly stated in source files, write "None explicitly flagged in source files."

---

## 7. Methodology Notes from Prior Wave

Any notes about sample size, screener changes, new questions introduced, or comparability caveats mentioned in the prior wave source:

- {note}

If none found: "No methodology notes found in source files."
```

---

## STEP 5 — Show summary and confirm output

After writing the file, show the user:

```
✓ Written: {OUTPUT_PATH}

Prior wave covered: {prior wave name}
Source files used: {N} ({list filenames})
Domains extracted: {list}
Metrics captured: {N} data points
Action items found: {N}

This file is ready for:
  → /build-project-context  (reads it for Section 3 & 4)
  → /hypotheses             (reads it as prior wave baseline)

Run /build-project-context next.
```

---

## RULES

1. **One confirmation only.** Show the file list once and wait for the user to approve. Do not ask again per-file during extraction.
2. **Primary source wins on metrics.** If the full `.pptx` deck and a `.md` ES summary both report the same domain, prefer the deck's data (more granular). Use the ES to fill narrative gaps.
3. **Exact values over paraphrases.** If the source says "42%", write "42%". Do not convert to "approximately 40%".
4. **Flag what's missing, don't invent.** If a domain is present in section headers but has no underlying data visible in the extracted text, write the domain heading and note "Headline only — no supporting data extracted from source."
5. **Prior wave name must be explicit.** Always state the wave the findings come from in the file header. Never leave it ambiguous which wave "prior" refers to.
6. **Output is wave-versioned.** `prior_wave_context.md` lives in `context/{current_wave}/` — each wave setup has its own snapshot of what the prior wave showed.
7. **Do not include current wave survey data.** This file is about what happened before. Current wave data comes from `source_data.xlsx` and is handled by Stage 0 + Stage 5/6.
