"""
Generate Demo One Slider PPTX files:
  - Light version: Column-divider layout (matching Redesign.html)
  - Dark version: Card-based layout (matching Dark.html)
Each has 3 slides: Cover, Pre-filled BA Genie, Blank Template.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# ─── Shared Data ────────────────────────────────────────────

METRICS = [("45%", "Manual Effort Reduction"), ("1.5x", "Throughput Increase"),
           ("~19.5%", "Total Cost Saving"), ("2 wk", "Pilot to Production")]
TMPL_METRICS = [("___%", "Headline metric"), ("___x", "Secondary metric"),
                ("___", "Cost / time saving"), ("___", "Deploy timeline")]

COL1 = [("Requirements are the #1 bottleneck", "in SDLC — 30% of projects delayed by spec gaps"),
        ("64% of failed projects", "trace root cause to errors in requirements gathering"),
        ("Rework costs 15x more", "when defects from requirements surface in production"),
        ("BAs spend most time on low-value tasks", "— formatting & consistency checks"),
        ("No scalable alternative", "— tools lack financial services domain context")]

COL2 = [("Automates generation", "of epics, stories, and test cases from high-level inputs"),
        ("Custom-built AI", "trained on 6 corpuses of regulatory & standards docs"),
        ("Outputs conform to internal standards", "— templates, taxonomy, criteria"),
        ("Model & tech agnostic", "— swap underlying LLMs as capabilities evolve"),
        ("Zero-disruption deployment", "— parallel to existing pipelines; 2 wk pilot")]

COL3 = [("Up to 45% reduction", "in manual effort for requirements authoring"),
        ("1.5x throughput", "— higher-quality requirements in less time"),
        ("Accuracy & consistency", "improved from first draft, fewer review cycles"),
        ("BAs shift to high-value work", "— stakeholder engagement & strategy"),
        ("Sustainable & scalable", "— pilot-to-production proven in 7 weeks")]

T1 = ["What is the business context? What process or function is affected?",
      "What is the core problem? What breaks, fails, or underperforms?",
      "What makes it worse? Volume, complexity, regulatory pressure?",
      "Why isn't it solved already? What have teams tried that fell short?",
      "Cost of inaction? Risk, lost revenue, attrition, compliance exposure?"]

T2 = ["What does the solution do in one sentence?",
      "What are 2-3 key features? What can the user do with it?",
      "What makes it differentiated? Custom model, proprietary data, UX?",
      "How does it integrate? API, plug-in, standalone, embedded?",
      "How fast to deploy? What is the pilot timeline and effort?"]

T3 = ["Headline metric? (e.g., X% reduction in time, cost, or errors)",
      "Secondary metric? (e.g., throughput, capacity, accuracy lift)",
      "Qualitative improvements? Speed, confidence, satisfaction?",
      "How does the team's work change? What do they stop/start?",
      "Is it scalable? Can it expand to other teams or regions?"]


# ─── Helpers ────────────────────────────────────────────────

def rect(slide, l, t, w, h, fill, line=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line:
        s.line.fill.solid(); s.line.fill.fore_color.rgb = line; s.line.width = Pt(1)
    else:
        s.line.fill.background()
    return s

def txt(slide, l, t, w, h, text, sz, color, bold=False, align=PP_ALIGN.LEFT, italic=False):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tb.text_frame.word_wrap = True
    p = tb.text_frame.paragraphs[0]
    p.text = text; p.font.size = Pt(sz); p.font.color.rgb = color
    p.font.bold = bold; p.font.name = "Inter"; p.alignment = align; p.font.italic = italic
    return tb

def rich_txt(slide, l, t, w, h, runs_list, align=PP_ALIGN.LEFT):
    """runs_list: [(text, size, color, bold, italic), ...]"""
    tb = slide.shapes.add_textbox(l, t, w, h)
    tb.text_frame.word_wrap = True
    p = tb.text_frame.paragraphs[0]
    p.alignment = align
    for text, sz, color, bold, italic in runs_list:
        r = p.add_run()
        r.text = text; r.font.size = Pt(sz); r.font.color.rgb = color
        r.font.bold = bold; r.font.name = "Inter"; r.font.italic = italic
    return tb

def bullets_box(slide, l, t, w, items, strong_c, text_c, is_tmpl=False, tmpl_c=None):
    tb = slide.shapes.add_textbox(l, t, w, Inches(3.5))
    tf = tb.text_frame; tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(4); p.space_after = Pt(2)
        if is_tmpl:
            r = p.add_run(); r.text = "•  " + item
            r.font.size = Pt(9.5); r.font.color.rgb = tmpl_c; r.font.italic = True; r.font.name = "Inter"
        else:
            bold_t, rest_t = item
            r1 = p.add_run(); r1.text = "•  " + bold_t + " "
            r1.font.size = Pt(9.5); r1.font.color.rgb = strong_c; r1.font.bold = True; r1.font.name = "Inter"
            r2 = p.add_run(); r2.text = rest_t
            r2.font.size = Pt(9.5); r2.font.color.rgb = text_c; r2.font.name = "Inter"
    return tb

def oval_num(slide, l, t, sz, num, bg, tc=RGBColor(0xFF,0xFF,0xFF)):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, l, t, sz, sz)
    s.fill.solid(); s.fill.fore_color.rgb = bg; s.line.fill.background()
    p = s.text_frame.paragraphs[0]
    p.text = str(num); p.font.size = Pt(10); p.font.color.rgb = tc
    p.font.bold = True; p.font.name = "Inter"; p.alignment = PP_ALIGN.CENTER

def rnd_rect_num(slide, l, t, sz, num, bg, tc=RGBColor(0xFF,0xFF,0xFF)):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, sz, sz)
    s.fill.solid(); s.fill.fore_color.rgb = bg; s.line.fill.background()
    # Reduce corner rounding
    try:
        s.adjustments[0] = 0.2
    except Exception:
        pass
    p = s.text_frame.paragraphs[0]
    p.text = str(num); p.font.size = Pt(10); p.font.color.rgb = tc
    p.font.bold = True; p.font.name = "Inter"; p.alignment = PP_ALIGN.CENTER


def set_bg(slide, color):
    slide.background.fill.solid(); slide.background.fill.fore_color.rgb = color


# ═══════════════════════════════════════════════════════════
# LIGHT VERSION — Column layout
# ═══════════════════════════════════════════════════════════

PINK = RGBColor(0xE6, 0x00, 0x7E)
NAVY = RGBColor(0x1A, 0x1F, 0x3D)
L_BG = RGBColor(0xED, 0xF1, 0xF5)
L_MID = RGBColor(0x4A, 0x50, 0x68)
L_DIM = RGBColor(0x7B, 0x81, 0x98)
L_RIB = RGBColor(0xF0, 0xF3, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
L_DIV = RGBColor(0xD4, 0xDA, 0xE5)
L_FOOT = RGBColor(0xFA, 0xFB, 0xFD)
COL1_C = PINK
COL2_C = NAVY
COL3_C = RGBColor(0x25, 0x63, 0xEB)


def light_cover(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, L_BG)
    txt(sl, Inches(0.8), Inches(0.6), Inches(3), Inches(0.4), "CAPCO", 16, L_DIM, bold=True)
    rect(sl, Inches(0.8), Inches(2.4), Inches(0.6), Inches(0.06), PINK)
    txt(sl, Inches(0.8), Inches(2.65), Inches(9), Inches(1.6), "AI Demo\nOne Slider", 48, NAVY, bold=True)
    txt(sl, Inches(0.8), Inches(4.5), Inches(8), Inches(0.8),
        "Your executive primer before the live demo \u2014 the business problem, the AI solution, and the impact, all on a single slide.",
        16, L_MID)
    # Bottom format hint
    rect(sl, Inches(0.8), Inches(6.5), Inches(0.6), Inches(0.04), PINK)
    txt(sl, Inches(0.8), Inches(6.65), Inches(3), Inches(0.3), "COVER FORMAT", 10, L_DIM, bold=True)
    rich_txt(sl, Inches(0.8), Inches(6.95), Inches(10), Inches(0.4), [
        ("[Solution Name]", 14, NAVY, True, False),
        ("  \u2014  One-line description of what it does and why it matters", 14, L_DIM, False, False),
    ])


def light_content(prs, title, sub, metrics, c1, c2, c3, tmpl=False, fr=""):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, WHITE)

    # Top bar
    rect(sl, Inches(0), Inches(0), SLIDE_W, Inches(0.95), NAVY)
    rect(sl, Inches(0.6), Inches(0.22), Inches(0.05), Inches(0.5), PINK)
    txt(sl, Inches(0.85), Inches(0.2), Inches(8), Inches(0.4), title, 18, WHITE, bold=True)
    txt(sl, Inches(0.85), Inches(0.55), Inches(8), Inches(0.3), sub, 9, RGBColor(0x99,0x9C,0xB0))
    txt(sl, Inches(10.5), Inches(0.3), Inches(2.5), Inches(0.4), "CAPCO", 13,
        RGBColor(0x88,0x8C,0xA8), bold=True, align=PP_ALIGN.RIGHT)

    # Metric ribbon
    rib_top = Inches(0.95); rib_h = Inches(0.85)
    rect(sl, Inches(0), rib_top, SLIDE_W, rib_h, L_RIB)
    mw = SLIDE_W / 4
    for i, (v, d) in enumerate(metrics):
        x = Emu(int(mw * i))
        vc = PINK if i == 0 and not tmpl else NAVY
        if tmpl: vc = L_DIV
        txt(sl, x + Inches(0.5), rib_top + Inches(0.12), Inches(1.5), Inches(0.5),
            v, 24 if not tmpl else 20, vc, bold=True)
        txt(sl, x + Inches(2.0), rib_top + Inches(0.18), Inches(1.5), Inches(0.5),
            d, 8.5, L_MID if not tmpl else L_DIM, italic=tmpl)
        if i < 3:
            rect(sl, Emu(int(mw*(i+1))), rib_top+Inches(0.15), Inches(0.01), Inches(0.55), L_DIV)
    rect(sl, Inches(0), rib_top+rib_h, SLIDE_W, Inches(0.01), L_DIV)

    # 3 columns
    ctop = Inches(1.85); cw = SLIDE_W / 3
    cols = [(c1, "PROBLEM & COMPLICATIONS", COL1_C, 1),
            (c2, "AI SOLUTION & FEATURES", COL2_C, 2),
            (c3, "IMPACT & OUTCOMES", COL3_C, 3)]
    for i, (bul, lbl, acc, num) in enumerate(cols):
        x = Emu(int(cw * i))
        oval_num(sl, x+Inches(0.5), ctop+Inches(0.05), Inches(0.28), num, acc)
        txt(sl, x+Inches(0.85), ctop+Inches(0.07), Inches(3), Inches(0.3), lbl, 10, acc, bold=True)
        rect(sl, x+Inches(0.5), ctop+Inches(0.42), Inches(3.5), Inches(0.035), acc)
        bullets_box(sl, x+Inches(0.5), ctop+Inches(0.55), Inches(3.7), bul,
                    NAVY, L_MID, is_tmpl=tmpl, tmpl_c=L_DIM)
        if i < 2:
            rect(sl, Emu(int(cw*(i+1))), ctop, Inches(0.01), Inches(4.8), L_DIV)

    # Footer
    ft = Inches(6.95)
    rect(sl, Inches(0), ft, SLIDE_W, Inches(0.55), L_FOOT)
    rect(sl, Inches(0), ft, SLIDE_W, Inches(0.01), L_DIV)
    rect(sl, Inches(0.6), ft+Inches(0.2), Inches(0.08), Inches(0.08), PINK)
    txt(sl, Inches(0.8), ft+Inches(0.12), Inches(4), Inches(0.3), "Confidential \u2014 Capco", 8, L_DIM)
    txt(sl, Inches(8), ft+Inches(0.12), Inches(5), Inches(0.3), fr, 8, L_DIM, align=PP_ALIGN.RIGHT)


def gen_light():
    prs = Presentation(); prs.slide_width = SLIDE_W; prs.slide_height = SLIDE_H
    light_cover(prs)
    light_content(prs, "SDLC Transformation with BA Genie",
                  "Financial Services  \u2022  Requirements Engineering  \u2022  Custom AI Solution",
                  METRICS, COL1, COL2, COL3,
                  fr="BA Genie  \u2022  SDLC / Engineering  \u2022  Financial Services")
    light_content(prs, "[Solution Name] \u2014 [Use Case Title]",
                  "[Industry]  \u2022  [Function / Process Area]  \u2022  [Solution Type]",
                  TMPL_METRICS, T1, T2, T3, tmpl=True,
                  fr="[Solution Name]  \u2022  [Domain]")
    prs.save("/home/user/Intake-to-GTM/Demo One Slider - Light.pptx")
    print("Created: Demo One Slider - Light.pptx")


# ═══════════════════════════════════════════════════════════
# DARK VERSION — Card-based layout
# ═══════════════════════════════════════════════════════════

D_BG = RGBColor(0x0C, 0x0E, 0x18)
D_SURF = RGBColor(0x14, 0x16, 0x26)
D_CARD = RGBColor(0x19, 0x1C, 0x30)
D_BORDER = RGBColor(0x25, 0x28, 0x40)
D_TEXT = RGBColor(0xE4, 0xE6, 0xF0)
D_MID = RGBColor(0x8B, 0x90, 0xB0)
D_DIM = RGBColor(0x4A, 0x4F, 0x6E)
VIOLET = RGBColor(0x8B, 0x5C, 0xF6)
BLUE = RGBColor(0x3B, 0x82, 0xF6)


def dark_cover(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, D_BG)
    # Subtle gradient overlay (approximate with semi-transparent shapes)
    rect(sl, Inches(0), Inches(0), Inches(6), Inches(7.5), RGBColor(0x12, 0x0E, 0x28))
    txt(sl, Inches(0.8), Inches(0.6), Inches(3), Inches(0.4), "CAPCO", 18, D_DIM, bold=True)
    # Tag
    tag = rect(sl, Inches(10), Inches(0.5), Inches(2.8), Inches(0.4), RGBColor(0x2A, 0x0E, 0x20))
    txt(sl, Inches(10.2), Inches(0.55), Inches(2.4), Inches(0.3), "AI SOLUTION DEMO", 9, PINK, bold=True, align=PP_ALIGN.CENTER)
    # Gradient line
    rect(sl, Inches(0.8), Inches(2.3), Inches(0.8), Inches(0.05), PINK)
    # Title
    txt(sl, Inches(0.8), Inches(2.6), Inches(9), Inches(1.8), "AI Demo\nOne Slider", 52, WHITE, bold=True)
    # Subtitle
    txt(sl, Inches(0.8), Inches(4.6), Inches(8), Inches(0.8),
        "Business context, AI solution, and measurable impact \u2014 one slide to set the stage before your live demo.",
        17, D_MID)
    # Bottom gradient bar
    rect(sl, Inches(0), Inches(7.44), SLIDE_W, Inches(0.06), PINK)
    # Format hint box
    rect(sl, Inches(0.7), Inches(6.2), Inches(8.5), Inches(0.7), D_CARD, line=D_BORDER)
    rich_txt(sl, Inches(1.2), Inches(6.35), Inches(7.5), Inches(0.4), [
        ("\u25B6  ", 11, PINK, False, False),
        ("[Solution Name]", 13, WHITE, True, False),
        ("  \u2014  One-line description of what it does and why it matters", 13, D_DIM, False, False),
    ])


def dark_content(prs, title, sub, metrics, c1, c2, c3, tmpl=False, fr=""):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, D_BG)

    # Header strip with left gradient accent
    rect(sl, Inches(0), Inches(0), SLIDE_W, Inches(0.9), RGBColor(0x10, 0x12, 0x20))
    rect(sl, Inches(0), Inches(0), Inches(0.05), Inches(0.9), PINK)  # left accent
    tit_c = WHITE if not tmpl else D_DIM
    txt(sl, Inches(0.6), Inches(0.15), Inches(9), Inches(0.4), title, 19, tit_c, bold=True)
    txt(sl, Inches(0.6), Inches(0.5), Inches(9), Inches(0.3), sub, 9, D_DIM)
    txt(sl, Inches(10.5), Inches(0.25), Inches(2.5), Inches(0.4), "CAPCO", 13, D_DIM, bold=True, align=PP_ALIGN.RIGHT)

    # Metric strip — 4 cells with 1px gaps
    ms_top = Inches(0.9); ms_h = Inches(0.95)
    rect(sl, Inches(0), ms_top, SLIDE_W, ms_h, RGBColor(0x1F, 0x22, 0x38))  # fallback bg
    cell_w = SLIDE_W / 4
    for i, (v, d) in enumerate(metrics):
        x = Emu(int(cell_w * i))
        rect(sl, x, ms_top, Emu(int(cell_w) - Inches(0.01).emu if i < 3 else int(cell_w)),
             ms_h, D_SURF)
        vc = PINK if i == 0 and not tmpl else WHITE
        if tmpl: vc = D_BORDER
        txt(sl, x + Inches(0.3), ms_top + Inches(0.1), Inches(2.5), Inches(0.55),
            v, 30 if not tmpl else 24, vc, bold=True, align=PP_ALIGN.CENTER)
        txt(sl, x + Inches(0.3), ms_top + Inches(0.6), Inches(2.5), Inches(0.3),
            d, 8, D_MID if not tmpl else D_DIM, align=PP_ALIGN.CENTER, italic=tmpl)

    # Cards
    card_top = Inches(2.0)
    card_h = Inches(4.7)
    card_w_val = (SLIDE_W.emu - Inches(0.7).emu) / 3
    gap = Inches(0.12).emu
    accents = [PINK, VIOLET, BLUE]
    labels = ["PROBLEM & COMPLICATIONS", "AI SOLUTION & FEATURES", "IMPACT & OUTCOMES"]
    data = [c1, c2, c3]

    for i in range(3):
        x = Emu(int(Inches(0.28).emu + i * (card_w_val + gap)))
        w = Emu(int(card_w_val))
        # Card background
        card = rect(sl, x, card_top, w, card_h, D_CARD, line=D_BORDER)
        # Top accent line
        rect(sl, x, card_top, w, Inches(0.04), accents[i])
        # Number badge
        rnd_rect_num(sl, x + Inches(0.25), card_top + Inches(0.25), Inches(0.28), i+1, accents[i])
        # Title
        txt(sl, x + Inches(0.6), card_top + Inches(0.26), Inches(3), Inches(0.25),
            labels[i], 9.5, accents[i], bold=True)
        # Divider
        rect(sl, x + Inches(0.25), card_top + Inches(0.65), w - Inches(0.5), Inches(0.01), D_BORDER)
        # Bullets
        bullets_box(sl, x + Inches(0.25), card_top + Inches(0.8), w - Inches(0.5),
                    data[i], WHITE, D_MID, is_tmpl=tmpl, tmpl_c=D_DIM)

    # Footer
    ft = Inches(6.95)
    rect(sl, Inches(0), ft, SLIDE_W, Inches(0.55), RGBColor(0x08, 0x0A, 0x12))
    rect(sl, Inches(0), ft, SLIDE_W, Inches(0.01), D_BORDER)
    rect(sl, Inches(0.6), ft + Inches(0.2), Inches(0.07), Inches(0.07), PINK)
    txt(sl, Inches(0.8), ft + Inches(0.12), Inches(4), Inches(0.3), "Confidential \u2014 Capco", 8, D_DIM)
    txt(sl, Inches(8), ft + Inches(0.12), Inches(5), Inches(0.3), fr, 8, D_DIM, align=PP_ALIGN.RIGHT)


def gen_dark():
    prs = Presentation(); prs.slide_width = SLIDE_W; prs.slide_height = SLIDE_H
    dark_cover(prs)
    dark_content(prs, "SDLC Transformation with BA Genie",
                 "Financial Services  \u2022  Requirements Engineering  \u2022  Custom AI Solution",
                 METRICS, COL1, COL2, COL3,
                 fr="BA Genie  \u2022  SDLC / Engineering  \u2022  Financial Services")
    dark_content(prs, "[Solution Name] \u2014 [Use Case Title]",
                 "[Industry]  \u2022  [Function / Process Area]  \u2022  [Solution Type]",
                 TMPL_METRICS, T1, T2, T3, tmpl=True,
                 fr="[Solution Name]  \u2022  [Domain]")
    prs.save("/home/user/Intake-to-GTM/Demo One Slider - Dark.pptx")
    print("Created: Demo One Slider - Dark.pptx")


# ─── Run ────────────────────────────────────────────────────

if __name__ == "__main__":
    gen_light()
    gen_dark()
