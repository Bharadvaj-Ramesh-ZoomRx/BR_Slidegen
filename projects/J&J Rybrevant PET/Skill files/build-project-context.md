# Build Project Context

You are a ZoomRx senior analyst. Your job is to synthesize the study methodology, client-shared field intelligence, and recent client call notes into a clean `Project_Context.md` file.

**Scope of this file:** Study design facts + what the client has shared about field conditions + open action items + wave-specific changes. No market/clinical data (that lives in `Market Context.md`). No KBQs. No hypotheses.

---

## SOURCE FILES

- **Study methodology:** `C:\Users\VinothRajapandian\Documents\Claude Apps\Context\PET, Project, KBQ.odt`
- **Client-shared field intelligence:** `C:\Users\VinothRajapandian\Documents\Claude Apps\Context\Rybrevant - Market, Client context and Business Hypotheses.xlsx` — use ONLY the `Client shared` column from sheet `Market and client context - V2`
- **Client call notes:** `C:\Users\VinothRajapandian\Documents\Claude Apps\Context\RYBREVANT Client Call Notes.docx` — extract last 6 months only
- **Prior wave Executive Summary:** User-provided at runtime — paste of the prior wave ES findings and recommendations (optional but strongly recommended)
- **Output file:** `C:\Users\VinothRajapandian\Documents\Claude Apps\Context\Project_Context.md`

---

## STEP-BY-STEP PROCESS

### Step 1 — Extract study design from ODT

Run this to extract all text from the ODT:

```bash
python3 -c "
import zipfile, re
with zipfile.ZipFile('C:/Users/VinothRajapandian/Documents/Claude Apps/Context/PET, Project, KBQ.odt') as z:
    with z.open('content.xml') as f:
        content = f.read().decode('utf-8')
        text = re.sub(r'<[^>]+>', ' ', content)
        text = re.sub(r'\s+', ' ', text)
        print(text)
"
```

Extract only:
- TP1 methodology (what is captured, who qualifies, base unit)
- TP2 methodology (what is captured, who qualifies, base unit)
- Sample targets and cadence
- HCP segment definitions

Ignore KBQs — those are a separate layer.

---

### Step 2 — Extract client-shared intelligence from Excel

Run this to extract only the `Client shared` column:

```bash
python3 -c "
import openpyxl
wb = openpyxl.load_workbook('C:/Users/VinothRajapandian/Documents/Claude Apps/Context/Rybrevant - Market, Client context and Business Hypotheses.xlsx')
ws = wb['Market and client context - V2']
headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
client_col = headers.index('Client shared') if 'Client shared' in headers else None
topic_col = headers.index('Topics') if 'Topics' in headers else None
area_col = headers.index('Specific areas') if 'Specific areas' in headers else None

if client_col is not None:
    for row in ws.iter_rows(min_row=2, values_only=True):
        topic = str(row[topic_col]).strip() if row[topic_col] else ''
        area = str(row[area_col]).strip() if row[area_col] else ''
        client = str(row[client_col]).strip() if row[client_col] else ''
        if client and client != 'None':
            print(f'[{topic} — {area}]')
            print(client)
            print()
"
```

Do NOT use the `Secondary research` column — that belongs in `Market Context.md`.
Do NOT use the `Hypotheses` or `Questions for the Team` columns — those are separate layers.

---

### Step 3 — Extract recent client call notes from DOCX

Run this to extract all text from the call notes DOCX:

```bash
python3 -c "
from docx import Document
doc = Document('C:/Users/VinothRajapandian/Documents/Claude Apps/Context/RYBREVANT Client Call Notes.docx')
for para in doc.paragraphs:
    if para.text.strip():
        print(para.text)
"
```

From the extracted content:
- Filter to entries from the last 6 months only (calculate cutoff from today's date at runtime)
- Extract: open action items, client decisions, client directions on strategy and reporting
- Do NOT use older entries

---

### Step 4 — Ask the user 5 questions

Ask:

1. **Wave** — What wave is this? (e.g., Q1 2026)
2. **Prior wave ES** — Please paste the key findings and recommendations from the prior wave Executive Summary. (If not available, skip — Section 4 of the output will be left blank.)
3. **Methodology changes** — Any survey or screener changes this wave? (new questions, benchmark shifts, ITP reframing, screener updates)
4. **Message changes** — Any new messages introduced, messages retired, or textual modifications to existing messages?
5. **Sample or segment changes** — Any changes to target sample sizes or HCP segment definitions?

If the user says "none" or "same as before" for questions 3–5, proceed with source file content only.

---

### Step 5 — Confirm before writing

After completing Steps 1–4, show the user a one-line summary per section of what will be written and ask for confirmation before generating the full file.

---

### Step 6 — Write Project_Context.md

Write the file with exactly this structure:

```
# Project Context — RYBREVANT+LAZCLUZE SFEA
**Wave:** [current wave]
**Last Updated:** [today's date]

---

## 1. Study Design

### TP1 — Immediate Recall
[Who qualifies, base unit, what is captured: interaction details / content / rep performance / impact]

### TP2 — Delayed Recall
[Who qualifies, base unit, what is captured: message effectiveness / message association / market perceptions]

### Sample & Cadence
- Quarterly fielding
- RYB target: ~100 interactions/wave; TAG target: ~60 interactions/wave
- TP2 sample smaller (unique HCP level, not interaction level)
- [Any wave-specific changes to sample]

### HCP Segments
[List all active segments with definitions. Flag any new segments this wave.]

---

## 2. Field Intelligence — Client Shared

[Organized by topic area matching the Excel structure. For each topic, summarize what J&J has communicated about brand strategy, rep direction, competitive awareness, and commercial priorities. Write in coherent prose/bullets — not raw cell text.]

Topics to cover (if content exists):
- Brand & product strategy
- Rep strategy and performance
- Messaging priorities
- Commercial challenges and 1L penetration
- Competitive awareness
- Patient journey and usage
- Non-rep promotion (NPP)
- Sales force structure

---

## 3. Open Action Items & Client Directions

[From call notes, last 6 months only. List chronologically, most recent first.]

Format each as:
**[Date] — [Meeting type]**
- Action item or client direction
- Action item or client direction

---

## 4. Prior Wave Summary — [wave name]

[If ES was provided by user: organize findings by domain — Sales Activity, Messaging, Interactions, Rep Performance, Impact/LTIP. Then list all recommendations verbatim.]

[If not provided: leave blank with note: "Prior wave ES not provided — populate before running /hypotheses-v2"]

---

## 5. Wave-Specific Changes

[Document all methodology, survey, and message changes this wave. For each change, flag which metric is affected and whether wave-over-wave comparison is valid.]

| Change | Metric Affected | Comparable QoQ? |
|--------|----------------|----------------|
| [change] | [metric] | [Yes / No / With caution] |

---

## 5. Active Messages This Wave

| Message | Domain | New Q1? | Baseline Available? |
|---------|--------|---------|-------------------|
| [message name] | [Efficacy / Safety / Convenience / NCCN / MoA / Dose Modification] | Y/N | Y/N |
```

---

## RULES

1. **Run all extraction steps before writing.** Do not write from memory.
2. **Client Shared column only from Excel.** Do not use Secondary Research, Hypotheses, or Questions columns.
3. **Call notes: last 6 months only.** Calculate cutoff at runtime from today's date.
4. **No market or clinical data.** Clinical trial results, competitive landscape, and physician mindset live in `Market Context.md`. Do not duplicate them here.
5. **No KBQs. No hypotheses.** Those are separate layers.
6. **Be specific.** Use real message names, real segment names, real dates from call notes.
7. **Confirm before writing.** Show section outline to user first, then write on confirmation.
