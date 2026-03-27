"""
create_process_flow_slide.py
Generates a single-slide PowerPoint process flow for the SlideGen pipeline.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from lxml import etree
from pptx.oxml.ns import qn

# ── Paths ────────────────────────────────────────────────────────────────────
TEMPLATE = (
    r"C:\Users\VinothRajapandian\Documents\Claude Apps\PPT Skills"
    r"\JJ PET RYBREVANT+LAZCLUZE Q4'25 Report.pptx"
)
OUTPUT = (
    r"C:\Users\VinothRajapandian\galen-consulting-r3m-report"
    r"\SlideGen_Process_Flow.pptx"
)

# ── Colours ──────────────────────────────────────────────────────────────────
W    = RGBColor(0xFF, 0xFF, 0xFF)
BLK  = RGBColor(0x20, 0x20, 0x20)
AH   = RGBColor(0x1F, 0x6F, 0xAD)   # auto  header — steel blue
UH   = RGBColor(0xBF, 0x1B, 0x1B)   # user  header — dark red
AB   = RGBColor(0xE9, 0xF3, 0xFB)   # auto  body bg — light blue
UB   = RGBColor(0xFD, 0xED, 0xED)   # user  body bg — light rose
AA   = RGBColor(0x1F, 0x6F, 0xAD)   # auto  accent
UA   = RGBColor(0xF7, 0x58, 0x24)   # user  accent — orange
LBL  = RGBColor(0x88, 0x88, 0x88)   # section-label grey
BDR  = RGBColor(0xCC, 0xCC, 0xCC)   # box border
ARR  = RGBColor(0xAA, 0xAA, 0xAA)   # arrow / connector grey
TBG  = RGBColor(0x1A, 0x1A, 0x1A)   # title bar bg

FD = "Johnson Display"
FB = "Johnson Text"

# ── Stage definitions ────────────────────────────────────────────────────────
STAGES = [
    dict(
        num="0", name="Index Excel", gate="AUTO",
        inputs=["source_data.xlsx"],
        tool="index_excel()", tool_kind="FUNCTION",
        outputs=["source_data.json"],
        note="Hash-cached; auto-invalidates on change",
    ),
    dict(
        num="1", name="Build Project Context", gate="USER GATE",
        inputs=["Call notes (.docx)", "KBQs doc (.odt)",
                "Market Context (.md)", "Prior Wave ES (.md)"],
        tool="/build-project-context", tool_kind="SKILL",
        outputs=["project_context.md"],
        note="Shows summary — pause & confirm",
    ),
    dict(
        num="2", name="Generate Hypotheses", gate="USER GATE",
        inputs=["project_context.md", "KBQs.md", "Survey_Context.md"],
        tool="/hypotheses", tool_kind="SKILL",
        outputs=["hypothesis_bank.md"],
        note="Shows count + domain breakdown",
    ),
    dict(
        num="3", name="Build Slide Plan", gate="USER GATE",
        inputs=["hypothesis_bank.md", "KBQs.md", "Survey_Context.md"],
        tool="/slide-plan", tool_kind="SKILL",
        outputs=["slide_plan.md"],
        note="Shows slide count + sections",
    ),
    dict(
        num="4", name="Generate Config", gate="AUTO",
        inputs=["slide_plan.md", "source_data.json"],
        tool="config_generator", tool_kind="FUNCTION",
        outputs=["config.yaml"],
        note="Maps slide plan → YAML extractions + asks",
    ),
    dict(
        num="5", name="Build the Deck", gate="AUTO",
        inputs=["config.yaml", "template.pptx", "source_data.json"],
        tool="generate_deck()", tool_kind="FUNCTION",
        outputs=["deck.pptx", "shape_registry.json"],
        note="14 renderers; native PPT charts",
    ),
]

# ── Layout constants (all in inches) ─────────────────────────────────────────
SL_W  = 13.33
BW    = 3.85    # box width
BH    = 2.82    # box height (both rows)
GAP   = 0.215   # inter-box gap (arrow zone)
MRG   = 0.695   # left/right margin
HDR_H = 0.40    # header strip height
AW    = 0.10    # left accent bar width

TY  = 0.08      # title bar y
TH  = 0.42      # title bar height
R1Y = 0.60      # row 1 y
R2Y = 3.65      # row 2 y  (0.60+2.82+0.23 gap = 3.65)

XS = [MRG + i * (BW + GAP) for i in range(3)]
# XS ≈ [0.695, 4.760, 8.825]


# ── Shape helpers ────────────────────────────────────────────────────────────
def I(v): return Inches(v)


def rect(slide, x, y, w, h, fill, line=None, lw=0.5):
    s = slide.shapes.add_shape(1, I(x), I(y), I(w), I(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line:
        s.line.color.rgb = line
        s.line.width = Pt(lw)
    else:
        s.line.fill.background()
    return s


def txb(slide, x, y, w, h, text, font, size,
        bold=False, italic=False, color=None,
        align=PP_ALIGN.LEFT, ml=0.03, mt=0.02):
    color = color or BLK
    b = slide.shapes.add_textbox(I(x), I(y), I(w), I(h))
    tf = b.text_frame
    tf.word_wrap = True
    tf.margin_left   = I(ml)
    tf.margin_top    = I(mt)
    tf.margin_right  = I(0.02)
    tf.margin_bottom = I(0)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name   = font
    r.font.size   = Pt(size)
    r.font.bold   = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return b


def txb_lines(slide, x, y, w, h, lines, font, size,
              bold=False, color=None, ml=0.03, ls_pt=None):
    color = color or BLK
    b = slide.shapes.add_textbox(I(x), I(y), I(w), I(h))
    tf = b.text_frame
    tf.word_wrap = True
    tf.margin_left   = I(ml)
    tf.margin_top    = I(0.01)
    tf.margin_right  = I(0.02)
    tf.margin_bottom = I(0)
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if ls_pt:
            pPr = p._p.get_or_add_pPr()
            lnSpc = etree.SubElement(pPr, qn('a:lnSpc'))
            spc   = etree.SubElement(lnSpc, qn('a:spcPts'))
            spc.set('val', str(int(ls_pt * 100)))
        run = p.add_run()
        run.text = line
        run.font.name  = font
        run.font.size  = Pt(size)
        run.font.bold  = bold
        run.font.color.rgb = color
    return b


def arrow_right(slide, x, y, w=0.16, h=0.22):
    # MSO_AUTO_SHAPE_TYPE RIGHT_ARROW = 13
    s = slide.shapes.add_shape(13, I(x), I(y), I(w), I(h))
    s.fill.solid()
    s.fill.fore_color.rgb = ARR
    s.line.fill.background()
    return s


def arrow_down(slide, x, y, w=0.22, h=0.16):
    # MSO_AUTO_SHAPE_TYPE DOWN_ARROW = 36
    s = slide.shapes.add_shape(36, I(x), I(y), I(w), I(h))
    s.fill.solid()
    s.fill.fore_color.rgb = ARR
    s.line.fill.background()
    return s


# ── Stage renderer ───────────────────────────────────────────────────────────
def render_stage(slide, stage, x, y):
    bh   = BH
    user = stage['gate'] == 'USER GATE'
    hc   = UH if user else AH
    bc   = UB if user else AB
    ac   = UA if user else AA

    # Outer box + border
    rect(slide, x, y, BW, bh, bc, BDR, 0.5)
    # Left accent bar
    rect(slide, x, y, AW, bh, ac)
    # Header strip
    rect(slide, x + AW, y, BW - AW, HDR_H, hc)

    # Stage number badge (white pill in header)
    badge_w = 0.30
    bx = x + AW + 0.07
    rect(slide, bx, y + 0.058, badge_w, HDR_H - 0.115, W)
    txb(slide, bx, y + 0.058, badge_w, HDR_H - 0.115,
        f"S{stage['num']}", FD, 9.5, bold=True, color=hc,
        align=PP_ALIGN.CENTER, ml=0, mt=0.03)

    # Stage name
    txb(slide, bx + badge_w + 0.12, y + 0.07,
        BW - AW - badge_w - 0.22, HDR_H - 0.10,
        stage['name'], FD, 9, bold=True, color=W, ml=0, mt=0.02)

    # Body content area
    cx  = x + AW + 0.08
    cw  = BW - AW - 0.12
    cy  = y + HDR_H + 0.07

    # ── INPUTS ──
    txb(slide, cx, cy, cw, 0.14, "INPUTS", FB, 5.5,
        bold=True, color=LBL, ml=0, mt=0)
    cy += 0.13
    inp_lines = [f"\u2022  {i}" for i in stage['inputs']]
    ih = len(inp_lines) * 0.14 + 0.04
    txb_lines(slide, cx, cy, cw, ih, inp_lines, FB, 5.5,
              color=BLK, ml=0.04, ls_pt=7.5)
    cy += ih + 0.06

    # Hairline
    rect(slide, cx, cy, cw, 0.01, BDR)
    cy += 0.05

    # ── TOOL / SKILL ──
    txb(slide, cx, cy, cw, 0.14, stage['tool_kind'], FB, 5.5,
        bold=True, color=LBL, ml=0, mt=0)
    cy += 0.13
    tc = UA if user else AH
    txb(slide, cx, cy, cw, 0.24, stage['tool'], FD, 8.5,
        bold=True, color=tc, ml=0.04, mt=0)
    cy += 0.23

    # Hairline
    rect(slide, cx, cy, cw, 0.01, BDR)
    cy += 0.05

    # ── OUTPUT ──
    txb(slide, cx, cy, cw, 0.14, "OUTPUT", FB, 5.5,
        bold=True, color=LBL, ml=0, mt=0)
    cy += 0.13
    out_lines = [f"\u2192  {o}" for o in stage['outputs']]
    oh = len(out_lines) * 0.155 + 0.04
    txb_lines(slide, cx, cy, cw, oh, out_lines, FD, 6.5,
              bold=True, color=hc, ml=0.04, ls_pt=8.5)

    # ── Note (small italic at bottom) ──
    ny = y + bh - 0.22
    txb(slide, cx, ny, cw, 0.20,
        stage['note'], FB, 5, italic=True, color=LBL, ml=0.04, mt=0)

    # ── Gate badge (bottom-right corner) ──
    badge_text   = "\u270B USER GATE" if user else "\u2699  AUTO"
    badge_color  = UA if user else AH
    badge_w2     = 1.06 if user else 0.74
    badge_x2     = x + BW - badge_w2 - 0.04
    badge_y2     = y + bh - 0.22
    rect(slide, badge_x2, badge_y2, badge_w2, 0.19, badge_color)
    txb(slide, badge_x2, badge_y2, badge_w2, 0.19,
        badge_text, FB, 6, bold=True, color=W,
        align=PP_ALIGN.CENTER, ml=0, mt=0.03)


# ── Main ─────────────────────────────────────────────────────────────────────
def main():
    prs    = Presentation(TEMPLATE)
    blank  = prs.slide_masters[0].slide_layouts[24]   # Blank layout

    # Strip all template slides (proper cleanup — drop_rel to avoid duplicate-name warnings)
    original_count = len(prs.slides)
    sld_id_lst = prs.part._element.find(qn('p:sldIdLst'))
    for _ in range(original_count):
        first = sld_id_lst[0]
        rId   = first.get(qn('r:id'))
        prs.part.drop_rel(rId)
        sld_id_lst.remove(first)

    slide = prs.slides.add_slide(blank)

    # ── Title bar ─────────────────────────────────────────────────────────
    rect(slide, 0, TY, SL_W, TH, TBG)
    txb(slide, 0.20, TY, 9.0, TH,
        "SlideGen Pipeline  \u2014  End-to-End Process Flow",
        FD, 14, bold=True, color=W, ml=0.15, mt=0.08)

    # Legend badges in title bar
    rect(slide, 10.35, TY + 0.09, 0.74, 0.23, AH)
    txb(slide, 10.35, TY + 0.09, 0.74, 0.23, "\u2699  AUTO",
        FB, 6.5, bold=True, color=W, align=PP_ALIGN.CENTER, ml=0, mt=0.04)
    rect(slide, 11.20, TY + 0.09, 1.10, 0.23, UA)
    txb(slide, 11.20, TY + 0.09, 1.10, 0.23, "\u270B USER GATE",
        FB, 6.5, bold=True, color=W, align=PP_ALIGN.CENTER, ml=0, mt=0.04)

    # ── Row 1 — Stages 0, 1, 2 ───────────────────────────────────────────
    for i, stage in enumerate(STAGES[:3]):
        render_stage(slide, stage, XS[i], R1Y)

    # Arrows between stages in row 1
    for i in range(2):
        ax = XS[i] + BW + 0.03
        ay = R1Y + BH / 2 - 0.11
        arrow_right(slide, ax, ay)

    # ── Row 2 — Stages 3, 4, 5 ───────────────────────────────────────────
    for i, stage in enumerate(STAGES[3:]):
        render_stage(slide, stage, XS[i], R2Y)

    # Arrows between stages in row 2
    for i in range(2):
        ax = XS[i] + BW + 0.03
        ay = R2Y + BH / 2 - 0.11
        arrow_right(slide, ax, ay)

    # ── Row-to-row connector ──────────────────────────────────────────────
    # L-path: down from Stage 2 bottom → across → down to Stage 3 top
    r1_bot  = R1Y + BH
    r2_top  = R2Y
    mid_y   = (r1_bot + r2_top) / 2   # horizontal cross-bar y
    cx2     = XS[2] + BW / 2           # x centre of Stage 2
    cx3     = XS[0] + BW / 2           # x centre of Stage 3
    line_t  = 0.018                    # line thickness

    rect(slide, cx2 - line_t/2, r1_bot,      line_t, mid_y - r1_bot,   ARR)  # vert down
    rect(slide, cx3 - line_t/2, mid_y,        cx2 - cx3 + line_t, line_t, ARR)  # horiz across
    rect(slide, cx3 - line_t/2, mid_y,        line_t, r2_top - mid_y,   ARR)  # vert up-to-top

    # Down-arrow indicator at Stage 3 top
    arrow_down(slide, cx3 - 0.11, r2_top - 0.13)

    # Row label
    txb(slide, cx3 - 0.55, mid_y - 0.12, 1.1, 0.18,
        "continues \u2192 \u2193", FB, 5.5, italic=True, color=LBL,
        align=PP_ALIGN.CENTER, ml=0, mt=0)

    # ── Row labels ────────────────────────────────────────────────────────
    # (optional subtle row labels on left margin)
    txb(slide, 0.05, R1Y + BH/2 - 0.10, 0.55, 0.20,
        "Context\nPrep", FB, 5.5, bold=False, color=LBL,
        align=PP_ALIGN.CENTER, ml=0, mt=0)
    txb(slide, 0.05, R2Y + BH/2 - 0.10, 0.55, 0.20,
        "Deck\nBuild", FB, 5.5, bold=False, color=LBL,
        align=PP_ALIGN.CENTER, ml=0, mt=0)

    # ── Footer ────────────────────────────────────────────────────────────
    txb(slide, 0.20, 6.56, 13.0, 0.18,
        "Source: ZoomRx SlideGen system  \u2014  "
        "Each wave: drop source_data.xlsx \u2192 run pipeline \u2192 deck.pptx",
        FB, 5.5, italic=True, color=LBL, ml=0, mt=0)

    prs.save(OUTPUT)
    print("Saved -> " + OUTPUT)


if __name__ == "__main__":
    main()
