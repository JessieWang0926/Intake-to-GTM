#!/usr/bin/env python3
"""Generate Demo One Slider PPTX with 4 slides."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# Colors
PINK = RGBColor(0xE8, 0x43, 0x93)
NAVY = RGBColor(0x0F, 0x1B, 0x2D)
BLUE = RGBColor(0x09, 0x84, 0xE3)
PURPLE = RGBColor(0x6C, 0x5C, 0xE7)
TEAL = RGBColor(0x00, 0xCE, 0xC9)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GREY = RGBColor(0xF4, 0xF5, 0xF7)
MID_GREY = RGBColor(0x6B, 0x7A, 0x8D)
DARK_GREY = RGBColor(0x3A, 0x45, 0x56)
BORDER_GREY = RGBColor(0xE0, 0xE4, 0xEA)
FOOTER_BG = RGBColor(0x1A, 0x1E, 0x2A)
COVER_DARK = RGBColor(0x16, 0x2A, 0x4A)
BODY_TEXT = RGBColor(0x3A, 0x45, 0x56)
MUTED = RGBColor(0x88, 0x92, 0xA6)
PLACEHOLDER_TEXT = RGBColor(0xB0, 0xB8, 0xC4)


def add_shape(slide, left, top, width, height, fill_color=None, line_color=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.line.fill.background()
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.fill.solid()
        shape.line.color.rgb = line_color
        shape.line.width = Pt(1)
    return shape


def add_text_box(slide, left, top, width, height, text, font_size=12,
                 bold=False, color=NAVY, align=PP_ALIGN.LEFT, font_name='Calibri',
                 anchor=MSO_ANCHOR.TOP):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    txBox.text_frame.word_wrap = True
    txBox.text_frame.auto_size = None
    p = txBox.text_frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = font_name
    p.alignment = align
    return txBox


def add_bullet_item(tf, bold_text, rest_text, bullet_color, font_size=10):
    """Add a bullet point with bold lead + normal rest."""
    p = tf.add_paragraph()
    p.space_before = Pt(6)
    p.space_after = Pt(2)
    p.level = 0

    # Bullet character
    run_b = p.add_run()
    run_b.text = "\u2022  "
    run_b.font.size = Pt(font_size)
    run_b.font.color.rgb = bullet_color
    run_b.font.bold = True

    # Bold part
    run_bold = p.add_run()
    run_bold.text = bold_text
    run_bold.font.size = Pt(font_size)
    run_bold.font.color.rgb = NAVY
    run_bold.font.bold = True
    run_bold.font.name = 'Calibri'

    # Rest
    run_rest = p.add_run()
    run_rest.text = " " + rest_text
    run_rest.font.size = Pt(font_size)
    run_rest.font.color.rgb = BODY_TEXT
    run_rest.font.bold = False
    run_rest.font.name = 'Calibri'


def add_copyright(slide):
    """Add dark copyright bar at the bottom."""
    bar = add_shape(slide, Inches(0), Inches(7.1), SLIDE_W, Inches(0.4), FOOTER_BG)
    bar.rotation = 0
    add_text_box(slide, Inches(0), Inches(7.13), SLIDE_W, Inches(0.35),
                 "\u00A9 2026 The Capital Markets Company. Capco Confidential. All rights reserved.",
                 font_size=8, color=RGBColor(0x80, 0x85, 0x95), align=PP_ALIGN.CENTER)


def add_gradient_stripe(slide, top, height=Inches(0.05)):
    """Add a multi-color stripe (approximated with 4 colored rectangles)."""
    seg_w = SLIDE_W // 4
    colors = [PINK, PURPLE, BLUE, TEAL]
    for i, c in enumerate(colors):
        add_shape(slide, Emu(seg_w * i), top, Emu(seg_w), height, c)


def build_cover_slide(prs, filled=True):
    """Build cover slide - split layout."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    # Top stripe
    add_gradient_stripe(slide, Inches(0), Inches(0.05))

    # Left half - dark background
    left_bg = add_shape(slide, Inches(0), Inches(0), Inches(6.667), SLIDE_H, COVER_DARK)

    # CAPCO brand top right
    add_text_box(slide, Inches(11.5), Inches(0.35), Inches(1.5), Inches(0.4),
                 "CAPCO", font_size=12, bold=True,
                 color=NAVY if filled else PLACEHOLDER_TEXT,
                 align=PP_ALIGN.RIGHT)

    if filled:
        # AI Lab badge
        badge = add_shape(slide, Inches(0.7), Inches(2.2), Inches(1.8), Inches(0.35),
                          fill_color=RGBColor(0x1E, 0x34, 0x60))
        add_text_box(slide, Inches(0.7), Inches(2.22), Inches(1.8), Inches(0.35),
                     "\u25CF  CAPCO AI LAB", font_size=8, bold=True, color=TEAL,
                     align=PP_ALIGN.CENTER)

        # Solution name
        add_text_box(slide, Inches(0.7), Inches(2.8), Inches(5.5), Inches(1.2),
                     "BA Genie", font_size=44, bold=True, color=WHITE)

        # Tagline
        add_text_box(slide, Inches(0.7), Inches(4.0), Inches(5.2), Inches(1.5),
                     "AI-powered requirements engineering that transforms the SDLC by automating the generation of accurate, complete epics, features, stories, and test cases.",
                     font_size=13, color=RGBColor(0x8E, 0xA8, 0xC4))
    else:
        add_text_box(slide, Inches(0.7), Inches(2.2), Inches(1.8), Inches(0.35),
                     "\u25CF  [TEAM / LAB]", font_size=8, bold=True,
                     color=RGBColor(0x4A, 0x6A, 0x80), align=PP_ALIGN.CENTER)
        add_text_box(slide, Inches(0.7), Inches(2.8), Inches(5.5), Inches(1.2),
                     "[Solution Name]", font_size=44, bold=True,
                     color=RGBColor(0x5A, 0x7A, 0x96))
        add_text_box(slide, Inches(0.7), Inches(4.0), Inches(5.2), Inches(1.5),
                     "[One sentence describing what this solution does and who it serves. Keep it concise and executive-ready.]",
                     font_size=13, color=RGBColor(0x5A, 0x7A, 0x96))

    # Right half - "At a Glance"
    add_text_box(slide, Inches(7.2), Inches(1.8), Inches(3), Inches(0.3),
                 "AT A GLANCE", font_size=9, bold=True, color=MUTED)

    # Three highlight cards
    highlights = [
        ("The Problem",
         "Manual requirements gathering delays 30% of projects and drives costly rework downstream",
         "[One sentence on the core business problem]", PINK),
        ("The Solution",
         "Custom-built AI that generates complete epics, stories, and test cases from business intent",
         "[One sentence on what the AI solution does]", NAVY),
        ("The Impact",
         "45% reduction in manual effort, 19.5% cost savings, pilot-to-production in 7 weeks",
         "[Key metrics and qualitative outcomes]", BLUE),
    ]

    y_pos = Inches(2.3)
    for title, desc_filled, desc_template, icon_color in highlights:
        # Icon box
        icon = add_shape(slide, Inches(7.2), y_pos, Inches(0.4), Inches(0.4), icon_color)
        icon.shape_style = None
        # Icon symbol
        symbols = {"The Problem": "\u26A0", "The Solution": "\U0001F4A1", "The Impact": "\u2197"}
        sym = symbols.get(title, "")
        add_text_box(slide, Inches(7.2), y_pos, Inches(0.4), Inches(0.4),
                     sym, font_size=14, color=WHITE, align=PP_ALIGN.CENTER)

        title_color = NAVY if filled else PLACEHOLDER_TEXT
        desc_color = MID_GREY if filled else PLACEHOLDER_TEXT
        desc = desc_filled if filled else desc_template

        add_text_box(slide, Inches(7.8), y_pos, Inches(4.5), Inches(0.25),
                     title, font_size=11, bold=True, color=title_color)
        add_text_box(slide, Inches(7.8), y_pos + Inches(0.28), Inches(4.5), Inches(0.5),
                     desc, font_size=10, color=desc_color)

        y_pos += Inches(1.2)

    # Meta pills
    pill_y = Inches(6.0)
    pills = ["Financial Services", "Requirements Engineering", "Custom AI Solution"] if filled else ["[Industry]", "[Domain]", "[Solution Type]"]
    x = Inches(7.2)
    for pill_text in pills:
        pw = Inches(len(pill_text) * 0.09 + 0.3)
        pill = add_shape(slide, x, pill_y, pw, Inches(0.3), LIGHT_GREY, BORDER_GREY)
        pill.fill.solid()
        pill.fill.fore_color.rgb = LIGHT_GREY
        add_text_box(slide, x, pill_y + Inches(0.02), pw, Inches(0.28),
                     pill_text, font_size=8, color=MID_GREY if filled else PLACEHOLDER_TEXT,
                     align=PP_ALIGN.CENTER)
        x += pw + Inches(0.1)

    add_copyright(slide)
    return slide


def build_content_slide(prs, filled=True):
    """Build the main 3-column content slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    # ── Dark header banner ──
    add_shape(slide, Inches(0), Inches(0), SLIDE_W, Inches(0.95), NAVY)

    # Pink accent bar
    add_shape(slide, Inches(0.45), Inches(0.2), Inches(0.04), Inches(0.55), PINK)

    if filled:
        add_text_box(slide, Inches(0.6), Inches(0.2), Inches(8), Inches(0.4),
                     "SDLC Transformation with BA Genie", font_size=18, bold=True, color=WHITE)
        # Subtitle
        add_text_box(slide, Inches(0.6), Inches(0.55), Inches(8), Inches(0.3),
                     "Financial Services  \u2022  Requirements Engineering  \u2022  Custom AI Solution",
                     font_size=10, color=RGBColor(0x80, 0x90, 0xA8))
    else:
        add_text_box(slide, Inches(0.6), Inches(0.2), Inches(8), Inches(0.4),
                     "[Solution Title with Solution Name]", font_size=18, bold=True,
                     color=RGBColor(0x60, 0x70, 0x88))
        add_text_box(slide, Inches(0.6), Inches(0.55), Inches(8), Inches(0.3),
                     "[Industry]  \u2022  [Domain]  \u2022  [Solution Type]",
                     font_size=10, color=RGBColor(0x50, 0x60, 0x78))

    # CAPCO brand
    add_text_box(slide, Inches(11.5), Inches(0.3), Inches(1.5), Inches(0.4),
                 "CAPCO", font_size=11, bold=True,
                 color=RGBColor(0x60, 0x70, 0x88), align=PP_ALIGN.RIGHT)

    # ── Metrics banner (light grey) ──
    add_shape(slide, Inches(0), Inches(0.95), SLIDE_W, Inches(1.05), LIGHT_GREY)
    # Bottom border
    add_shape(slide, Inches(0), Inches(1.98), SLIDE_W, Inches(0.02), BORDER_GREY)

    metrics = [
        ("45%", "Reduction in\nmanual effort", PINK, "[XX%]", "[Primary\nmetric]"),
        ("1.5x", "Throughput\nincrease", NAVY, "[X.Xx]", "[Throughput\nor speed]"),
        ("~19.5%", "Total cost\nsaving", PURPLE, "[~XX%]", "[Cost or\nefficiency]"),
        ("2 wk", "Pilot to\nproduction", BLUE, "[X wk]", "[Time to\nvalue]"),
    ]

    seg_w = Inches(3.333)
    for i, (val, lbl, color, tpl_val, tpl_lbl) in enumerate(metrics):
        x = Emu(seg_w * i)
        display_val = val if filled else tpl_val
        display_lbl = lbl if filled else tpl_lbl
        val_color = color if filled else PLACEHOLDER_TEXT

        add_text_box(slide, x + Inches(0.4), Inches(1.08), Inches(1.6), Inches(0.7),
                     display_val, font_size=32, bold=True, color=val_color,
                     anchor=MSO_ANCHOR.MIDDLE)
        add_text_box(slide, x + Inches(2.0), Inches(1.12), Inches(1.2), Inches(0.7),
                     display_lbl, font_size=10, color=MID_GREY if filled else PLACEHOLDER_TEXT,
                     anchor=MSO_ANCHOR.MIDDLE)

        # Vertical divider
        if i < 3:
            add_shape(slide, Emu(seg_w * (i + 1)), Inches(1.05), Inches(0.01), Inches(0.9), BORDER_GREY)

    # ── Three columns ──
    col_data = [
        {
            "num_color": PINK,
            "heading": "PROBLEM & COMPLICATIONS",
            "heading_color": PINK,
            "bullets": [
                ("Requirements are the #1 bottleneck", "in SDLC \u2014 30% of tech projects delayed by specification gaps"),
                ("64% of failed projects", "trace root cause to errors in requirements gathering and design phases"),
                ("Rework costs 15x more", "when defects from requirements surface in testing or production"),
                ("BAs spend most time on low-value tasks", "\u2014 formatting, consistency checks, and re-writing"),
                ("No scalable alternative", "\u2014 off-the-shelf tools lack financial services domain and regulatory context"),
            ],
            "template_bullets": [
                ("[Core problem]", "\u2014 what is the business context and what process is affected?"),
                ("[Quantified pain]", "\u2014 what does this cost in time, money, or quality?"),
                ("[Compounding effect]", "\u2014 what makes this problem worse over time?"),
                ("[Barrier to solving]", "\u2014 why hasn't this been solved already?"),
                ("[Cost of inaction]", "\u2014 what happens if nothing changes?"),
            ],
        },
        {
            "num_color": NAVY,
            "heading": "AI SOLUTION & FEATURES",
            "heading_color": NAVY,
            "bullets": [
                ("Automates generation", "of epics, features, user stories, and test cases from high-level inputs"),
                ("Custom-built AI", "trained on 6 corpuses of regulatory and standards documents specific to client"),
                ("Outputs conform to internal standards", "\u2014 templates, taxonomy, and acceptance criteria formats"),
                ("Model & tech agnostic architecture", "\u2014 swap underlying LLMs as capabilities evolve"),
                ("Zero-disruption deployment", "\u2014 runs in parallel to existing pipelines; pilot in 2 weeks"),
            ],
            "template_bullets": [
                ("[What it does]", "\u2014 describe the solution at a high level"),
                ("[Key capability #1]", "\u2014 most important feature or function"),
                ("[Key capability #2]", "\u2014 how it integrates or deploys"),
                ("[Differentiation]", "\u2014 what makes this better than alternatives?"),
                ("[Deployment model]", "\u2014 how quickly can it go live?"),
            ],
        },
        {
            "num_color": BLUE,
            "heading": "IMPACT & OUTCOMES",
            "heading_color": BLUE,
            "bullets": [
                ("Up to 45% reduction", "in manual effort for requirements authoring across the team"),
                ("1.5x throughput", "\u2014 more complete, higher-quality requirements delivered in less time"),
                ("Accuracy, completeness, consistency", "improved from first draft \u2014 less review cycles"),
                ("BAs shift to high-value work", "\u2014 stakeholder engagement, edge-case analysis, strategy"),
                ("Sustainable & scalable", "\u2014 designed for growing teams; full pilot-to-production in 7 weeks"),
            ],
            "template_bullets": [
                ("[Primary metric]", "\u2014 e.g., % reduction in effort, cost, or time"),
                ("[Secondary metric]", "\u2014 e.g., throughput increase, error reduction"),
                ("[Quality improvement]", "\u2014 e.g., accuracy, consistency, compliance"),
                ("[Workforce impact]", "\u2014 how does this change how people work?"),
                ("[Scalability]", "\u2014 what does this unlock at enterprise level?"),
            ],
        },
    ]

    col_w = Inches(4.1)
    col_start_y = Inches(2.15)

    for i, col in enumerate(col_data):
        x = Inches(0.35) + Emu(col_w * i + Inches(0.15) * i)

        # Numbered circle
        circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, col_start_y, Inches(0.3), Inches(0.3), )
        circle.fill.solid()
        circle.fill.fore_color.rgb = col["num_color"]
        circle.line.fill.background()
        add_text_box(slide, x, col_start_y + Inches(0.02), Inches(0.3), Inches(0.28),
                     str(i + 1), font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        # Heading
        add_text_box(slide, x + Inches(0.4), col_start_y + Inches(0.02), Inches(3.2), Inches(0.3),
                     col["heading"], font_size=10, bold=True, color=col["heading_color"])

        # Colored rule
        add_shape(slide, x, col_start_y + Inches(0.42), Inches(3.8), Inches(0.03), col["num_color"])

        # Bullets
        txBox = slide.shapes.add_textbox(x, col_start_y + Inches(0.55), Inches(3.8), Inches(3.8))
        txBox.text_frame.word_wrap = True
        # Remove default paragraph
        txBox.text_frame.paragraphs[0].text = ""

        bullets = col["bullets"] if filled else col["template_bullets"]
        bullet_color = col["num_color"]

        for j, (bold_part, rest_part) in enumerate(bullets):
            if j == 0:
                p = txBox.text_frame.paragraphs[0]
            else:
                p = txBox.text_frame.add_paragraph()
            p.space_before = Pt(6)
            p.space_after = Pt(4)

            run_dot = p.add_run()
            run_dot.text = "\u2022  "
            run_dot.font.size = Pt(10)
            run_dot.font.color.rgb = bullet_color
            run_dot.font.bold = True
            run_dot.font.name = 'Calibri'

            run_b = p.add_run()
            run_b.text = bold_part
            run_b.font.size = Pt(10)
            run_b.font.color.rgb = NAVY if filled else MUTED
            run_b.font.bold = True
            run_b.font.name = 'Calibri'

            run_r = p.add_run()
            run_r.text = " " + rest_part
            run_r.font.size = Pt(10)
            run_r.font.color.rgb = BODY_TEXT if filled else MUTED
            run_r.font.bold = False
            run_r.font.name = 'Calibri'

        # Column dividers
        if i < 2:
            div_x = x + Inches(4.15)
            add_shape(slide, div_x, col_start_y, Inches(0.01), Inches(4.3), BORDER_GREY)

    # ── Footer ──
    add_shape(slide, Inches(0), Inches(6.65), SLIDE_W, Inches(0.01), BORDER_GREY)

    if filled:
        # Pink dot + confidential
        add_shape(slide, Inches(0.45), Inches(6.77), Inches(0.08), Inches(0.08), PINK)
        add_text_box(slide, Inches(0.6), Inches(6.72), Inches(2.5), Inches(0.25),
                     "Confidential \u2014 Capco", font_size=8, color=MUTED)
        add_text_box(slide, Inches(8.5), Inches(6.72), Inches(4.5), Inches(0.25),
                     "BA Genie  \u2022  SDLC / Engineering  \u2022  Financial Services",
                     font_size=8, color=MUTED, align=PP_ALIGN.RIGHT)
    else:
        add_shape(slide, Inches(0.45), Inches(6.77), Inches(0.08), Inches(0.08), PINK)
        add_text_box(slide, Inches(0.6), Inches(6.72), Inches(2.5), Inches(0.25),
                     "Confidential \u2014 Capco", font_size=8, color=MUTED)
        add_text_box(slide, Inches(8.5), Inches(6.72), Inches(4.5), Inches(0.25),
                     "[Solution Name]  \u2022  [Domain]  \u2022  [Industry]",
                     font_size=8, color=PLACEHOLDER_TEXT, align=PP_ALIGN.RIGHT)

    add_copyright(slide)
    return slide


def main():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    # Slide 1: Cover (BA Genie filled)
    build_cover_slide(prs, filled=True)

    # Slide 2: Content (BA Genie filled)
    build_content_slide(prs, filled=True)

    # Slide 3: Cover (template)
    build_cover_slide(prs, filled=False)

    # Slide 4: Content (template)
    build_content_slide(prs, filled=False)

    out = "/home/user/Intake-to-GTM/Demo One Slider.pptx"
    prs.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
