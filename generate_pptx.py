"""Generate a professional PPTX: AI Use Case Intake to Go-To-Market.
Clean business design — navy accent, white background, C-level appropriate."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import datetime

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

DARK = RGBColor(0x1A, 0x1A, 0x2E)
NAVY = RGBColor(0x2B, 0x47, 0x7A)
GRAY = RGBColor(0x6B, 0x72, 0x80)
LIGHT_GRAY = RGBColor(0xF3, 0xF4, 0xF6)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BORDER_GRAY = RGBColor(0xE5, 0xE7, 0xEB)

STEP_COLORS = {
    1: RGBColor(0x25, 0x63, 0xEB),
    2: RGBColor(0x05, 0x96, 0x69),
    3: RGBColor(0xD9, 0x77, 0x06),
    4: RGBColor(0xDC, 0x26, 0x26),
    5: RGBColor(0x7C, 0x3A, 0xED),
    6: RGBColor(0x08, 0x91, 0xB2),
    7: RGBColor(0xEA, 0x58, 0x0C),
    8: RGBColor(0x0D, 0x94, 0x88),
}


def add_text(tf, text, size=14, bold=False, color=DARK, alignment=PP_ALIGN.LEFT, space_after=Pt(4)):
    p = tf.add_paragraph() if len(tf.paragraphs) > 0 and tf.paragraphs[0].text != "" else tf.paragraphs[0]
    p.alignment = alignment
    p.space_after = space_after
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = 'Calibri'
    return p


def add_table_slide(prs, step_num, title, phase, tasks):
    """Add a step slide with task/owner table."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
    sc = STEP_COLORS[step_num]

    # Top accent bar
    bar = slide.shapes.add_shape(1, Inches(0), Inches(0), prs.slide_width, Inches(0.06))
    bar.fill.solid()
    bar.fill.fore_color.rgb = sc
    bar.line.fill.background()

    # Step number circle
    left_margin = Inches(0.8)
    circ = slide.shapes.add_shape(9, left_margin, Inches(0.4), Inches(0.55), Inches(0.55))  # Oval
    circ.fill.solid()
    circ.fill.fore_color.rgb = sc
    circ.line.fill.background()
    tf = circ.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = str(step_num)
    run.font.size = Pt(22)
    run.font.bold = True
    run.font.color.rgb = WHITE
    run.font.name = 'Calibri'
    tf.paragraphs[0].space_before = Pt(4)

    # Title
    txBox = slide.shapes.add_textbox(Inches(1.5), Inches(0.42), Inches(8), Inches(0.5))
    tf = txBox.text_frame
    add_text(tf, title, size=22, bold=True, color=DARK)

    # Phase tag
    txBox2 = slide.shapes.add_textbox(Inches(9.8), Inches(0.48), Inches(2), Inches(0.35))
    tf2 = txBox2.text_frame
    add_text(tf2, phase.upper(), size=10, bold=True, color=GRAY, alignment=PP_ALIGN.RIGHT)

    # Table
    rows = len(tasks) + 1
    cols = 2
    tbl_left = left_margin
    tbl_top = Inches(1.2)
    tbl_width = Inches(11.7)
    tbl_height = Inches(0.4) * rows

    table_shape = slide.shapes.add_table(rows, cols, tbl_left, tbl_top, tbl_width, tbl_height)
    table = table_shape.table

    table.columns[0].width = Inches(8.2)
    table.columns[1].width = Inches(3.5)

    # Header
    for i, hdr_text in enumerate(["Task", "Owner"]):
        cell = table.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT if i == 0 else PP_ALIGN.RIGHT
        run = p.add_run()
        run.text = hdr_text
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = WHITE
        run.font.name = 'Calibri'

    # Rows
    for idx, (task, owner) in enumerate(tasks):
        row_idx = idx + 1

        # Task cell
        cell_t = table.cell(row_idx, 0)
        if idx % 2 == 0:
            cell_t.fill.solid()
            cell_t.fill.fore_color.rgb = LIGHT_GRAY
        else:
            cell_t.fill.solid()
            cell_t.fill.fore_color.rgb = WHITE
        p = cell_t.text_frame.paragraphs[0]
        # Parse bold markers
        parts = task.split("**")
        for j, part in enumerate(parts):
            run = p.add_run()
            run.text = part
            run.font.size = Pt(12)
            run.font.name = 'Calibri'
            run.font.color.rgb = DARK
            if j % 2 == 1:
                run.font.bold = True

        # Owner cell
        cell_o = table.cell(row_idx, 1)
        if idx % 2 == 0:
            cell_o.fill.solid()
            cell_o.fill.fore_color.rgb = LIGHT_GRAY
        else:
            cell_o.fill.solid()
            cell_o.fill.fore_color.rgb = WHITE
        p2 = cell_o.text_frame.paragraphs[0]
        p2.alignment = PP_ALIGN.RIGHT
        run = p2.add_run()
        run.text = owner
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = NAVY
        run.font.name = 'Calibri'

    return slide


def add_gate_slide(prs, title, subtitle, is_governance=False):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    color = RGBColor(0xDC, 0x26, 0x26) if is_governance else RGBColor(0xD9, 0x77, 0x06)
    bg_color = RGBColor(0xFE, 0xE2, 0xE2) if is_governance else RGBColor(0xFE, 0xF3, 0xC7)

    # Center box
    box = slide.shapes.add_shape(1, Inches(3.5), Inches(2.2), Inches(6.3), Inches(2.5))
    box.fill.solid()
    box.fill.fore_color.rgb = bg_color
    box.line.color.rgb = color
    box.line.width = Pt(2)

    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.space_after = Pt(12)
    run = p.add_run()
    run.text = title
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = color
    run.font.name = 'Calibri'

    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    run = p2.add_run()
    run.text = subtitle
    run.font.size = Pt(14)
    run.font.color.rgb = GRAY
    run.font.name = 'Calibri'

    return slide


# ============================================================
# SLIDE 1: COVER
# ============================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])

# Navy bottom band
band = slide.shapes.add_shape(1, Inches(0), Inches(5.8), prs.slide_width, Inches(1.7))
band.fill.solid()
band.fill.fore_color.rgb = NAVY
band.line.fill.background()

# Title
txBox = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(11), Inches(1.2))
tf = txBox.text_frame
tf.word_wrap = True
add_text(tf, "AI Use Case Intake", size=40, bold=True, color=DARK, alignment=PP_ALIGN.LEFT)
add_text(tf, "to Go-To-Market", size=40, bold=True, color=NAVY, alignment=PP_ALIGN.LEFT)

# Subtitle
txBox2 = slide.shapes.add_textbox(Inches(1), Inches(3.5), Inches(8), Inches(0.8))
tf2 = txBox2.text_frame
add_text(tf2, "Step-by-Step Guide", size=20, color=GRAY, alignment=PP_ALIGN.LEFT)
add_text(tf2, "8 steps from ideation to launch \u2014 who does what at every stage", size=13, color=GRAY, alignment=PP_ALIGN.LEFT)

# Bottom band text
txBox3 = slide.shapes.add_textbox(Inches(1), Inches(6.1), Inches(6), Inches(0.9))
tf3 = txBox3.text_frame
add_text(tf3, "AI Power Team", size=18, bold=True, color=WHITE, alignment=PP_ALIGN.LEFT)
add_text(tf3, datetime.date.today().strftime("%B %Y"), size=12, color=RGBColor(0xA0, 0xB0, 0xC8), alignment=PP_ALIGN.LEFT)

# ============================================================
# SLIDE 2: TABLE OF CONTENTS
# ============================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])

# Top accent
bar = slide.shapes.add_shape(1, Inches(0), Inches(0), prs.slide_width, Inches(0.06))
bar.fill.solid()
bar.fill.fore_color.rgb = NAVY
bar.line.fill.background()

txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(6), Inches(0.6))
tf = txBox.text_frame
add_text(tf, "Table of Contents", size=28, bold=True, color=DARK)

toc_items = [
    (1, "Identify & Generate AI Use Cases", "Ideation"),
    (2, "Submit Use Case Intake", "Intake"),
    (3, "AI Enablement Team Review", "Assessment"),
    (4, "Project Setup & Resource Allocation", "Setup"),
    (5, "Define the MVP & Build", "Build"),
    (6, "GTM Ready: Package the Solution", "GTM Ready"),
    (7, "GTM Push: Showcase & Sales Activation", "Go-To-Market"),
    (8, "Track Results & Continuous Improvement", "Measure"),
]

for i, (num, t, phase) in enumerate(toc_items):
    y = Inches(1.4) + Inches(0.6) * i

    # Number
    circ = slide.shapes.add_shape(9, Inches(1), y, Inches(0.4), Inches(0.4))
    circ.fill.solid()
    circ.fill.fore_color.rgb = STEP_COLORS[num]
    circ.line.fill.background()
    ctf = circ.text_frame
    p = ctf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = str(num)
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = WHITE
    run.font.name = 'Calibri'

    # Title
    txB = slide.shapes.add_textbox(Inches(1.7), y + Inches(0.02), Inches(7), Inches(0.4))
    ttf = txB.text_frame
    p = ttf.paragraphs[0]
    run = p.add_run()
    run.text = t
    run.font.size = Pt(16)
    run.font.color.rgb = DARK
    run.font.name = 'Calibri'

    # Phase
    txP = slide.shapes.add_textbox(Inches(9), y + Inches(0.06), Inches(2.5), Inches(0.3))
    ptf = txP.text_frame
    p = ptf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = phase
    run.font.size = Pt(11)
    run.font.color.rgb = GRAY
    run.font.name = 'Calibri'


# ============================================================
# STEP SLIDES
# ============================================================

# Step 1
add_table_slide(prs, 1, "Identify & Generate AI Use Cases", "Ideation", [
    ("**Talk to your Practice Lead** about client pain points and market gaps", "Practice Lead"),
    ("Brainstorm from in-flight projects, client conversations & market demand", "AI Champion"),
    ("Check **Demo Catalog** (79+ solutions, 10 domains) to avoid duplication. Interested in a solution? **Request a demo directly with the Product Owner**", "AI Champion"),
])

# Step 2
add_table_slide(prs, 2, "Submit Use Case Intake", "Intake", [
    ("Complete **AI Use-Case Intake Form** (one row per use case)", "AI Champion"),
    ("Answer **4 Core Questions**: What pain? Who feels it? Why it matters? Why unsolved?", "AI Champion"),
    ("Provide: title, practice, persona, business value, revenue potential, target accounts", "AI Champion"),
    ("Assign priority: **High** / **Medium** / **Low**", "AI Champion + Practice Lead"),
])

# Step 3
add_table_slide(prs, 3, "AI Enablement Team Review", "Assessment", [
    ("Determine scope: **POC / MVP / Full Build** based on business value & tech depth", "AI Enablement Team"),
    ("Assess **strategic alignment**, data maturity & regulatory impact", "AI Enablement Team + AI Champion"),
    ("**Confirm resource availability** across AI Forward Engineers", "AI Enablement Team"),
    ("Prioritize **top 1\u20132 use cases** per champion for development", "AI Enablement Team"),
])

# Release Gate
add_gate_slide(prs, "RELEASE GATE", "Approved to proceed?  Yes \u2192 continue to build   |   No \u2192 return to ideation")

# Step 4
add_table_slide(prs, 4, "Project Setup & Resource Allocation", "Setup", [
    ("**Asset Register** \u2014 Model the AI solution into asset register (MP and above)", "Business Sponsor"),
    ("**Set Up Jira** \u2014 Create project, epics, stories & sprint backlog", "AI Enablement Team"),
    ("**Allocate Resources** \u2014 Assign AI Forward Engineers via Retain", "AI Enablement Team"),
    ("**Set up healthy check-in cadence** \u2014 demo \u2192 feedback \u2192 refine cycles (keep AI Enablement in the loop)", "AI Champion + AI Fwd Engineer"),
])

# Step 5
add_table_slide(prs, 5, "Define the MVP & Build", "Build", [
    ("**Define MVP scope** \u2014 features, demo expectations & success criteria", "AI Fwd Engineer + AI Champion"),
    ("Develop through **sprint cycles** with iterative testing & user feedback", "AI Fwd Engineer + AI Enablement Team"),
    ("Provide **sample data & real-life use case scenarios** to guide the build", "AI Champion"),
    ("Build **demo environment** based on provided scenarios & data", "AI Forward Engineer"),
    ("Track: Backlog \u2192 Build In-Progress \u2192 **MVP Demo Complete**", "AI Enablement Team"),
])

# Governance Gate
add_gate_slide(prs, "GOVERNANCE GATE",
    "AI Governance  \u2022  Compliance Checks  \u2022  Data Security Review  \u2022  Model Risk Validation",
    is_governance=True)

# Step 6
add_table_slide(prs, 6, "GTM Ready \u2014 Package the Solution", "GTM Ready", [
    ("Prepare **demo materials**: demo script, mock data & demo video recording", "AI Fwd Engineer + AI Champion"),
    ("Finalize **demo script** and conduct full run-throughs", "AI Fwd Engineer + AI Champion"),
    ('Prepare **"AI vs. Non-AI" talk track** (competitive differentiation)', "AI Champion + Sales"),
    ("Create **one-pager**, pricing template & engagement model", "Domain Leader + Sales"),
    ("Map **target accounts & buyer personas**; register in Demo Catalog", "Sales + AI Enablement Team"),
])

# Step 7
add_table_slide(prs, 7, "GTM Push \u2014 Showcase & Sales Activation", "Go-To-Market", [
    ("**AI Capability Showcases** (bi-weekly) \u2014 all AI Power Team members should attend", "All AI Power Team"),
    ("Run **Sales Activations** targeting specific accounts", "Sales + Account Partner"),
    ("Assign **follow-up actions** with owners and due dates", "AI Champion"),
    ("**Assemble GTM Folder**: demo video, demo script, sales one-slider & sales deck", "AI Enablement Team"),
    ("**Publish to Capco MarketHub & AI Lab** to reach a larger audience", "AI Enablement Team"),
])

# Step 8
add_table_slide(prs, 8, "Track Results & Continuous Improvement", "Measure", [
    ("Track **CTAR** (Client Target Account Revenue): pipeline generated, deals won, revenue attributed to each AI solution", "AI Enablement + AI Champion + Sponsor"),
    ("Monitor **revenue metrics per solution**: total pipeline value, win rate, avg deal size & time-to-close", "AI Enablement + Business Sponsor"),
    ("Monitor maturity: Idea \u2192 In Build \u2192 Demo-Ready \u2192 GTM-Ready \u2192 **In Pursuit**", "AI Enablement Team"),
    ("Capture **client feedback** & feed learnings back into next intake cycle", "AI Champion + Domain Leader"),
])

# ============================================================
# SAVE
# ============================================================
output_path = "/home/user/Intake-to-GTM/AI-Use-Case-to-GTM-Step-Guide.pptx"
prs.save(output_path)
print(f"PPTX saved to: {output_path}")
