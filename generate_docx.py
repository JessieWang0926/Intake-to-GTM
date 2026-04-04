"""Generate a professional DOCX guide: AI Use Case Intake to Go-To-Market.
Clean business design — minimal color, C-level appropriate."""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import datetime

doc = Document()

for section in doc.sections:
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

# Styles
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(10.5)
style.font.color.rgb = RGBColor(0x1A, 0x1A, 0x2E)
style.paragraph_format.space_after = Pt(4)
style.paragraph_format.line_spacing = 1.15

DARK = RGBColor(0x1A, 0x1A, 0x2E)
ACCENT = RGBColor(0x2B, 0x47, 0x7A)  # Muted navy
GRAY = RGBColor(0x6B, 0x72, 0x80)
LIGHT_GRAY = "F3F4F6"
TABLE_HEADER = "2B477A"  # Navy
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

h1 = doc.styles['Heading 1']
h1.font.name = 'Calibri'
h1.font.size = Pt(20)
h1.font.bold = True
h1.font.color.rgb = DARK
h1.paragraph_format.space_before = Pt(0)
h1.paragraph_format.space_after = Pt(4)

h2 = doc.styles['Heading 2']
h2.font.name = 'Calibri'
h2.font.size = Pt(13)
h2.font.bold = True
h2.font.color.rgb = ACCENT
h2.paragraph_format.space_before = Pt(10)
h2.paragraph_format.space_after = Pt(4)


def set_cell_shading(cell, hex_color):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def add_task_table(doc, tasks, step_num):
    table = doc.add_table(rows=1 + len(tasks), cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for row in table.rows:
        row.cells[0].width = Cm(1)
        row.cells[1].width = Cm(10)
        row.cells[2].width = Cm(5)

    # Header
    hdr = table.rows[0]
    for i, text in enumerate(["#", "Task", "Owner"]):
        cell = hdr.cells[i]
        p = cell.paragraphs[0]
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = WHITE
        run.font.name = 'Calibri'
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i == 0 else WD_ALIGN_PARAGRAPH.LEFT
        set_cell_shading(cell, TABLE_HEADER)

    for idx, (task, owner) in enumerate(tasks):
        row = table.rows[idx + 1]
        # Number
        cell_n = row.cells[0]
        p = cell_n.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"{step_num}.{idx+1}")
        run.font.size = Pt(9)
        run.font.color.rgb = GRAY
        run.font.name = 'Calibri'

        # Task
        cell_t = row.cells[1]
        p = cell_t.paragraphs[0]
        parts = task.split("**")
        for j, part in enumerate(parts):
            run = p.add_run(part)
            run.font.size = Pt(10)
            run.font.name = 'Calibri'
            if j % 2 == 1:
                run.bold = True

        # Owner
        cell_o = row.cells[2]
        p2 = cell_o.paragraphs[0]
        run = p2.add_run(owner)
        run.font.size = Pt(9)
        run.font.color.rgb = ACCENT
        run.font.name = 'Calibri'
        run.bold = True

        if idx % 2 == 0:
            set_cell_shading(cell_n, LIGHT_GRAY)
            set_cell_shading(cell_t, LIGHT_GRAY)
            set_cell_shading(cell_o, LIGHT_GRAY)

    doc.add_paragraph()


def add_step_heading(doc, step_num, title, phase):
    p = doc.add_heading(f"Step {step_num}: {title}", level=1)

    p2 = doc.add_paragraph()
    run = p2.add_run(phase.upper())
    run.font.size = Pt(9)
    run.font.color.rgb = GRAY
    run.bold = True
    run.font.name = 'Calibri'

    # Thin rule
    p3 = doc.add_paragraph()
    p3.paragraph_format.space_before = Pt(0)
    p3.paragraph_format.space_after = Pt(6)
    run = p3.add_run("\u2500" * 72)
    run.font.color.rgb = RGBColor(0xD1, 0xD5, 0xDB)
    run.font.size = Pt(5)


def add_gate(doc, title, subtitle):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.rows[0].cells[0]
    set_cell_shading(cell, "FEF3C7")

    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x92, 0x40, 0x0E)
    run.font.name = 'Calibri'

    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p2.add_run(subtitle)
    run.font.size = Pt(9)
    run.font.color.rgb = GRAY
    run.font.name = 'Calibri'

    doc.add_paragraph()


# ============================================================
# COVER PAGE
# ============================================================
for _ in range(7):
    doc.add_paragraph()

title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run("AI Use Case Intake")
run.font.size = Pt(36)
run.font.color.rgb = DARK
run.bold = True
run.font.name = 'Calibri'

title2 = doc.add_paragraph()
title2.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title2.add_run("to Go-To-Market")
run.font.size = Pt(36)
run.font.color.rgb = ACCENT
run.bold = True
run.font.name = 'Calibri'

doc.add_paragraph()
doc.add_paragraph()

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = sub.add_run("Step-by-Step Guide")
run.font.size = Pt(16)
run.font.color.rgb = GRAY
run.font.name = 'Calibri'

doc.add_paragraph()

desc = doc.add_paragraph()
desc.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = desc.add_run("8 steps from ideation to launch\nWho does what at every stage")
run.font.size = Pt(11)
run.font.color.rgb = GRAY
run.font.name = 'Calibri'

for _ in range(6):
    doc.add_paragraph()

# Rule
rule = doc.add_paragraph()
rule.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = rule.add_run("\u2500" * 40)
run.font.color.rgb = RGBColor(0xD1, 0xD5, 0xDB)
run.font.size = Pt(8)

doc.add_paragraph()

footer_p = doc.add_paragraph()
footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = footer_p.add_run("AI Power Team")
run.font.size = Pt(13)
run.font.color.rgb = DARK
run.bold = True

date_p = doc.add_paragraph()
date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = date_p.add_run(datetime.date.today().strftime("%B %Y"))
run.font.size = Pt(10)
run.font.color.rgb = GRAY

doc.add_page_break()

# ============================================================
# TABLE OF CONTENTS
# ============================================================
toc_heading = doc.add_heading("Table of Contents", level=1)
doc.add_paragraph()

toc_items = [
    (1, "Identify & Generate AI Use Cases", "Ideation"),
    (2, "Submit Use Case Intake", "Intake"),
    (3, "AI Enablement Team Review", "Assessment"),
    (0, "Release Gate Decision", ""),
    (4, "Project Setup & Resource Allocation", "Setup"),
    (5, "Define the MVP & Build", "Build"),
    (0, "Governance Gate", ""),
    (6, "GTM Ready: Package the Solution", "GTM Ready"),
    (7, "GTM Push: Showcase & Sales Activation", "Go-To-Market"),
    (8, "Track Results & Continuous Improvement", "Measure"),
]

for num, title_text, phase in toc_items:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    if num:
        run = p.add_run(f"  {num}.  ")
        run.font.size = Pt(12)
        run.font.color.rgb = ACCENT
        run.bold = True
        run.font.name = 'Calibri'
        run = p.add_run(title_text)
        run.font.size = Pt(12)
        run.font.color.rgb = DARK
        run.font.name = 'Calibri'
        if phase:
            run = p.add_run(f"   {phase}")
            run.font.size = Pt(9)
            run.font.color.rgb = GRAY
            run.font.name = 'Calibri'
    else:
        run = p.add_run(f"        {title_text}")
        run.font.size = Pt(10)
        run.font.color.rgb = GRAY
        run.italic = True
        run.font.name = 'Calibri'

doc.add_page_break()

# ============================================================
# STEPS
# ============================================================

# Step 1
add_step_heading(doc, 1, "Identify & Generate AI Use Cases", "Ideation")
add_task_table(doc, [
    ("**Talk to your Practice Lead** about client pain points and market gaps", "Practice Lead"),
    ("Brainstorm from in-flight projects, client conversations & market demand", "AI Champion"),
    ("Check **Demo Catalog** (79+ solutions, 10 domains) to avoid duplication. Interested in a solution? **Request a demo directly with the Product Owner**", "AI Champion"),
], 1)
doc.add_page_break()

# Step 2
add_step_heading(doc, 2, "Submit Use Case Intake", "Intake")
add_task_table(doc, [
    ("Complete **AI Use-Case Intake Form** (one row per use case)", "AI Champion"),
    ("Answer: **What is the pain/problem?**", "AI Champion"),
    ("Answer: **Who feels the pain?** (user, buyer, team)", "AI Champion"),
    ("Answer: **Why does it matter?** (cost, speed, risk, growth)", "AI Champion"),
    ("Answer: **Why is it unsolved today?** (blockers, workarounds)", "AI Champion"),
    ("Provide: title, practice, persona, business value, revenue potential, target accounts", "AI Champion"),
    ("Assign priority: **High** / **Medium** / **Low**", "AI Champion + Practice Lead"),
], 2)
doc.add_page_break()

# Step 3
add_step_heading(doc, 3, "AI Enablement Team Review", "Assessment")
add_task_table(doc, [
    ("Determine scope: **POC / MVP / Full Build** based on business value & tech depth", "AI Enablement Team"),
    ("Assess **strategic alignment**, data maturity & regulatory impact", "AI Enablement Team + AI Champion"),
    ("**Confirm resource availability** across AI Forward Engineers", "AI Enablement Team"),
    ("Prioritize **top 1\u20132 use cases** per champion for development", "AI Enablement Team"),
], 3)
add_gate(doc, "RELEASE GATE", "Approved to proceed?  Yes \u2192 continue to build   |   No \u2192 return to ideation")
doc.add_page_break()

# Step 4
add_step_heading(doc, 4, "Project Setup & Resource Allocation", "Setup")
add_task_table(doc, [
    ("**Asset Register** \u2014 Model the AI solution into asset register (MP and above)", "Business Sponsor"),
    ("**Set Up Jira & Project Folder** \u2014 Create Jira project and project folder structure (epics, stories & backlog defined during build)", "AI Enablement Team"),
    ("**Allocate Resources** \u2014 Assign AI Forward Engineers via Retain", "AI Enablement Team"),
    ("**Set up healthy check-in cadence** \u2014 demo \u2192 feedback \u2192 refine cycles (keep AI Enablement in the loop)", "AI Champion + AI Fwd Engineer"),
], 4)
doc.add_page_break()

# Step 5
add_step_heading(doc, 5, "Define the MVP & Build", "Build")
add_task_table(doc, [
    ("**Define MVP scope** \u2014 features, demo expectations & success criteria", "AI Fwd Engineer + AI Champion"),
    ("Develop through **sprint cycles** with iterative testing & user feedback", "AI Fwd Engineer + AI Enablement Team"),
    ("Provide **sample data & real-life use case scenarios** to guide the build", "AI Champion"),
    ("Build **demo environment** based on provided scenarios & data", "AI Forward Engineer"),
    ("Track: Backlog \u2192 Build In-Progress \u2192 **MVP Demo Complete**", "AI Enablement Team"),
    ("Present solution on **Governance Call** \u2014 explain the solution, walk through architecture & complete risk assessment before launch", "AI Forward Engineer"),
], 5)
add_gate(doc, "GOVERNANCE GATE", "AI Governance  \u2022  Compliance Checks  \u2022  Data Security Review  \u2022  Model Risk Validation")
doc.add_page_break()

# Step 6
add_step_heading(doc, 6, "GTM Ready \u2014 Package the Solution", "GTM Ready")
add_task_table(doc, [
    ("Prepare **demo materials**: demo script, mock data & demo video recording", "AI Fwd Engineer + AI Champion"),
    ("Finalize **demo script** and conduct full run-throughs", "AI Fwd Engineer + AI Champion"),
    ('Prepare **"AI vs. Non-AI" talk track** (competitive differentiation)', "AI Champion + Sales"),
    ("Create **one-pager**, pricing template & engagement model", "Domain Leader + Sales"),
    ("Map **target accounts & buyer personas**; register in Demo Catalog", "Sales + AI Enablement Team"),
], 6)

p = doc.add_paragraph()
run = p.add_run("Deliverables: ")
run.bold = True
run.font.size = Pt(10)
run.font.color.rgb = ACCENT
run = p.add_run("Demo script  |  Mock data  |  Demo video  |  One-pager  |  Talk track  |  Pricing template  |  Target account list")
run.font.size = Pt(9.5)
run.font.color.rgb = GRAY
doc.add_page_break()

# Step 7
add_step_heading(doc, 7, "GTM Push \u2014 Showcase & Sales Activation", "Go-To-Market")
add_task_table(doc, [
    ("**AI Capability Showcases** (bi-weekly) \u2014 all AI Power Team members should attend", "All AI Power Team"),
    ("Run **Sales Activations** targeting specific accounts", "Sales + Account Partner"),
    ("Assign **follow-up actions** with owners and due dates", "AI Champion"),
    ("**Assemble GTM Folder**: demo video, demo script, sales one-slider & sales deck", "AI Enablement Team"),
    ("**Publish to Capco MarketHub & AI Lab** to reach a larger audience", "AI Enablement Team"),
], 7)
doc.add_page_break()

# Step 8
add_step_heading(doc, 8, "Track Results & Continuous Improvement", "Measure")
add_task_table(doc, [
    ("Track **CTAR** (Client Target Account Revenue): pipeline generated, deals won, revenue attributed to each AI solution", "AI Enablement + AI Champion + Sponsor"),
    ("Monitor **revenue metrics per solution**: total pipeline value, win rate, average deal size & time-to-close", "AI Enablement + Business Sponsor"),
    ("Monitor maturity: Idea \u2192 In Build \u2192 Demo-Ready \u2192 GTM-Ready \u2192 **In Pursuit**", "AI Enablement Team"),
    ("Capture **client feedback** & feed learnings back into next intake cycle", "AI Champion + Domain Leader"),
], 8)

# Loop callout
table = doc.add_table(rows=1, cols=1)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
cell = table.rows[0].cells[0]
set_cell_shading(cell, LIGHT_GRAY)
p = cell.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Continuous Loop: ")
run.bold = True
run.font.size = Pt(10)
run.font.color.rgb = ACCENT
run = p.add_run("Results feed back into Step 1. Client feedback generates new ideas, successful demos inspire adjacent solutions, market learnings sharpen future intake.")
run.font.size = Pt(9.5)
run.font.color.rgb = RGBColor(0x37, 0x41, 0x51)

# Save
output_path = "/home/user/Intake-to-GTM/AI-Use-Case-to-GTM-Step-Guide.docx"
doc.save(output_path)
print(f"DOCX saved to: {output_path}")
