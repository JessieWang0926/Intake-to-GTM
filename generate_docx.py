"""Generate a step-by-step DOCX guide: AI Use Case Intake to Go-To-Market."""

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import datetime

doc = Document()

# -- Page margins --
for section in doc.sections:
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

# -- Style setup --
style = doc.styles['Normal']
font = style.font
font.name = 'Calibri'
font.size = Pt(11)
font.color.rgb = RGBColor(0x33, 0x33, 0x33)
style.paragraph_format.space_after = Pt(6)
style.paragraph_format.line_spacing = 1.15

# Heading styles
for level, (size, color) in enumerate([
    (Pt(28), RGBColor(0x4A, 0x15, 0x8D)),  # Heading 1
    (Pt(18), RGBColor(0x5B, 0x21, 0xB6)),  # Heading 2
    (Pt(14), RGBColor(0x6D, 0x28, 0xD9)),  # Heading 3
], start=1):
    h = doc.styles[f'Heading {level}']
    h.font.name = 'Calibri'
    h.font.size = size
    h.font.color.rgb = color
    h.font.bold = True
    h.paragraph_format.space_before = Pt(18 if level > 1 else 0)
    h.paragraph_format.space_after = Pt(8)


def add_purple_bar(doc):
    """Add a thin purple horizontal rule."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run("_" * 80)
    run.font.color.rgb = RGBColor(0x7C, 0x3A, 0xED)
    run.font.size = Pt(2)


def add_role_line(doc, roles):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    run = p.add_run("Key Roles: ")
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0x5B, 0x21, 0xB6)
    run = p.add_run("  |  ".join(roles))
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)


def add_bullet(doc, text, bold_prefix=None):
    p = doc.add_paragraph(style='List Bullet')
    if bold_prefix:
        run = p.add_run(bold_prefix)
        run.bold = True
        run.font.size = Pt(10.5)
        p.add_run(text).font.size = Pt(10.5)
    else:
        p.add_run(text).font.size = Pt(10.5)


# ============================================================
# COVER PAGE
# ============================================================
for _ in range(6):
    doc.add_paragraph()

title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run("AI Use Case Intake\nto Go-To-Market")
run.font.size = Pt(36)
run.font.color.rgb = RGBColor(0x4A, 0x15, 0x8D)
run.bold = True
run.font.name = 'Calibri'

doc.add_paragraph()

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = sub.add_run("Step-by-Step Guide")
run.font.size = Pt(20)
run.font.color.rgb = RGBColor(0x7C, 0x3A, 0xED)
run.font.name = 'Calibri'

doc.add_paragraph()

desc = doc.add_paragraph()
desc.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = desc.add_run(
    "A comprehensive guide for Domain Partners, AI Champions,\n"
    "and AI Lab teams to navigate the full lifecycle\n"
    "from use case ideation through go-to-market launch."
)
run.font.size = Pt(12)
run.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
run.font.name = 'Calibri'

for _ in range(4):
    doc.add_paragraph()

footer_p = doc.add_paragraph()
footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = footer_p.add_run("Global GenAI Lab")
run.font.size = Pt(14)
run.font.color.rgb = RGBColor(0x5B, 0x21, 0xB6)
run.bold = True

date_p = doc.add_paragraph()
date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = date_p.add_run(datetime.date.today().strftime("%B %Y"))
run.font.size = Pt(11)
run.font.color.rgb = RGBColor(0x9C, 0xA3, 0xAF)

# Page break
doc.add_page_break()

# ============================================================
# TABLE OF CONTENTS
# ============================================================
doc.add_heading("Table of Contents", level=1)
doc.add_paragraph()

toc_items = [
    ("1.", "Identify & Generate AI Use Cases", "Ideation"),
    ("2.", "Submit Use Case Intake", "Intake"),
    ("3.", "AI Lab Management Review", "Assessment"),
    ("", "Release Gate Decision", ""),
    ("4.", "Project Setup & Resource Allocation", "Setup"),
    ("5.", "Define the MVP & Build", "Build"),
    ("", "Governance Gate", ""),
    ("6.", "GTM Ready: Package the Solution", "GTM Ready"),
    ("7.", "GTM Push: Showcase & Sales Activation", "Go-To-Market"),
    ("8.", "Track Results & Continuous Improvement", "Measure"),
]

for num, title, phase in toc_items:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    if num:
        run = p.add_run(f"  {num}  ")
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0x5B, 0x21, 0xB6)
        run.bold = True
        run = p.add_run(title)
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
        if phase:
            run = p.add_run(f"  [{phase}]")
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(0x9C, 0xA3, 0xAF)
    else:
        run = p.add_run(f"       {title}")
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(0x7C, 0x3A, 0xED)
        run.italic = True

doc.add_page_break()

# ============================================================
# STEP 1
# ============================================================
doc.add_heading("Step 1: Identify & Generate AI Use Cases", level=1)
p = doc.add_paragraph()
run = p.add_run("PHASE: IDEATION")
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(0x7C, 0x3A, 0xED)
run.bold = True

add_purple_bar(doc)

doc.add_heading("What to Do", level=3)
add_bullet(doc, "Talk to your Practice Lead about client pain points, market gaps, and emerging opportunities in your domain.")
add_bullet(doc, "Each Domain Partner elects an AI Champion to be the point person for AI use case ideation within their practice area.")
add_bullet(doc, "Brainstorm use cases from in-flight projects, client conversations, and competitive market analysis.")
add_bullet(doc, "Check the Global Demo Catalog to see what solutions already exist across 10 domains and 79+ mapped solutions \u2014 avoid duplication.")
add_bullet(doc, "For AI Infusion opportunities, identify repetitive work or manual processes in current fixed-fee projects that could be enhanced with AI.")

doc.add_heading("Requirements", level=3)
add_bullet(doc, "Nominated AI Champion per practice area (elected by Domain Partner)")
add_bullet(doc, "Clear understanding of target client personas and their pain points")
add_bullet(doc, "Awareness of existing solutions in the Global Demo Catalog across 10 domains: Banking/Financial Services, Insurance, SDLC/Engineering, Risk/Compliance/Legal, Knowledge Management, Operations/PMO, Cybersecurity, Customer/Sales/Marketing, Finance/CFO, Cross Domain")
add_bullet(doc, "Use cases must be repeatable across multiple clients or opportunities")
add_bullet(doc, "Relevant to a real client or market pain point")

add_role_line(doc, ["Practice Lead", "AI Champion", "Domain Partner", "Account Partner"])
doc.add_page_break()

# ============================================================
# STEP 2
# ============================================================
doc.add_heading("Step 2: Submit Use Case Intake", level=1)
p = doc.add_paragraph()
run = p.add_run("PHASE: INTAKE")
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(0x7C, 0x3A, 0xED)
run.bold = True

add_purple_bar(doc)

doc.add_heading("What to Do", level=3)
add_bullet(doc, "Complete the AI Use-Case Intake Form (one row per use case).")
p = doc.add_paragraph()
run = p.add_run("Answer the 4 Core Desirability Questions:")
run.bold = True
run.font.size = Pt(11)
add_bullet(doc, "What is the pain/problem? \u2014 Describe the business challenge clearly.")
add_bullet(doc, "Who feels the pain? \u2014 Identify the user, buyer, or team affected.")
add_bullet(doc, "Why does it matter? \u2014 Explain impact on cost, speed, risk, growth, or experience.")
add_bullet(doc, "Why is it unsolved today? \u2014 Note current blockers, manual workarounds, or market gaps.")

doc.add_heading("Intake Form Fields Required", level=3)
add_bullet(doc, "Use Case Title")
add_bullet(doc, "Practice & Sub-Practice / Function")
add_bullet(doc, "Champion Name")
add_bullet(doc, "Description of Use Case")
add_bullet(doc, "Pain Point Addressed & Current Blocker to Resolution")
add_bullet(doc, "Target User (Persona)")
add_bullet(doc, "Expected Business Value (Impact)")
add_bullet(doc, "Market Demand Assessment")
add_bullet(doc, "Priority Ranking: High (strong client value, clear demand) / Medium (promising, needs validation) / Low (interesting, less urgent)")
add_bullet(doc, "Revenue Potential & Targeted Accounts")

doc.add_heading("Tips", level=3)
add_bullet(doc, "Keep answers concise. Submit multiple ideas if relevant.")
add_bullet(doc, "Focus on practical, industry-relevant opportunities.")
add_bullet(doc, "Reference the Global Catalog sheet for existing solutions.")

add_role_line(doc, ["AI Champion", "AI Infusion Champion"])
doc.add_page_break()

# ============================================================
# STEP 3
# ============================================================
doc.add_heading("Step 3: AI Lab Management Review", level=1)
p = doc.add_paragraph()
run = p.add_run("PHASE: ASSESSMENT")
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(0x7C, 0x3A, 0xED)
run.bold = True

add_purple_bar(doc)

doc.add_heading("What to Do", level=3)
add_bullet(doc, "AI Lab Leadership reviews all submitted use cases against strategic priorities.")
add_bullet(doc, "Leadership determines project scope: POC (Proof of Concept), MVP (Minimum Viable Product), or Full Build \u2014 based on the solution\u2019s Business Value and Technology Depth.")
add_bullet(doc, "Evaluate strategic alignment, data maturity, and regulatory/compliance impact.")
add_bullet(doc, "Prioritize top 1\u20132 use cases per champion for development.")

doc.add_heading("Requirements", level=3)
add_bullet(doc, "Strategic alignment with organizational goals confirmed")
add_bullet(doc, "Data readiness and maturity assessment completed")
add_bullet(doc, "Regulatory / compliance impact analysis performed")
add_bullet(doc, "Resource availability confirmed across AI Forward Engineers")
add_bullet(doc, "Business Sponsor identified and sign-off obtained")

doc.add_heading("Release Gate Decision", level=2)
p = doc.add_paragraph(
    "At this point, a gateway decision is made: Is the project approved to proceed? "
    "If YES, the project moves to setup and build. If NO, the use case returns to "
    "ideation for refinement or is deprioritized."
)
p.runs[0].font.size = Pt(11)

add_role_line(doc, ["AI Lab Leadership", "Business Sponsor", "Technology Owner"])
doc.add_page_break()

# ============================================================
# STEP 4
# ============================================================
doc.add_heading("Step 4: Project Setup & Resource Allocation", level=1)
p = doc.add_paragraph()
run = p.add_run("PHASE: SETUP")
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(0x7C, 0x3A, 0xED)
run.bold = True

add_purple_bar(doc)

doc.add_heading("Three Parallel Workstreams", level=3)

p = doc.add_paragraph()
run = p.add_run("Asset Register: ")
run.bold = True
p.add_run("Model the AI solution into the asset register. Document the solution\u2019s scope, technology stack, and expected outcomes.")

p = doc.add_paragraph()
run = p.add_run("Set Up Jira: ")
run.bold = True
p.add_run("Create the Jira project with epics, user stories, and sprint backlog. Establish the sprint cadence and assign the project team.")

p = doc.add_paragraph()
run = p.add_run("Allocate Resources: ")
run.bold = True
p.add_run("Use Retain to assign AI Forward Engineers and supporting team members. Confirm availability and engagement timeline.")

doc.add_heading("Requirements", level=3)
add_bullet(doc, "Jira project board with defined epics, stories, and acceptance criteria")
add_bullet(doc, "Assigned AI Forward Engineer(s) with confirmed availability")
add_bullet(doc, "Asset registration completed in the central register")
add_bullet(doc, "Resource allocation confirmed in Retain system")
add_bullet(doc, "Point of contact and Business Sponsor formally identified")
add_bullet(doc, "Weekly checkpoint cadence established")

add_role_line(doc, ["AI Lab Management", "AI Forward Engineer", "Program Manager"])
doc.add_page_break()

# ============================================================
# STEP 5
# ============================================================
doc.add_heading("Step 5: Define the MVP & Build", level=1)
p = doc.add_paragraph()
run = p.add_run("PHASE: BUILD")
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(0x7C, 0x3A, 0xED)
run.bold = True

add_purple_bar(doc)

doc.add_heading("What to Do", level=3)
add_bullet(doc, "Define the MVP: The AI Lab and the Business collaboratively define the Minimum Viable Product \u2014 what features are in scope, what the demo should show, and what success looks like.")
add_bullet(doc, "Develop the solution through iterative sprint cycles.")
add_bullet(doc, "Build the demo environment with real-world scenarios and sample data.")
add_bullet(doc, "Conduct iterative testing with user feedback at each sprint review.")
add_bullet(doc, "Track demo status progression: Backlog \u2192 Build In-Progress \u2192 MVP Demo Complete \u2192 Additional Features In Progress.")

doc.add_heading("Requirements", level=3)
add_bullet(doc, "MVP scope definition signed off by the Business Sponsor")
add_bullet(doc, "Development environment and tooling ready")
add_bullet(doc, "Sprint cadence with weekly checkpoints")
add_bullet(doc, "Demo script and test scenarios prepared")
add_bullet(doc, "Compliance and data security review initiated in parallel")
add_bullet(doc, "Office hours support available for the build team")

doc.add_heading("Governance Gate", level=2)
p = doc.add_paragraph(
    "Before proceeding to GTM, the solution must pass through a Governance Gate:\n"
)
add_bullet(doc, "AI Governance Evaluation \u2014 assess model fairness, explainability, and ethical considerations")
add_bullet(doc, "Compliance Checks \u2014 regulatory alignment (e.g., ECOA, Basel, NY Reg 126, BCBS 239)")
add_bullet(doc, "Data Security Review \u2014 data handling, privacy, and security controls")
add_bullet(doc, "Model Risk Validation \u2014 model performance, robustness, and risk assessment")

add_role_line(doc, ["AI Forward Engineer", "Business Sponsor", "AI Champion", "CRO", "CDO", "Compliance"])
doc.add_page_break()

# ============================================================
# STEP 6
# ============================================================
doc.add_heading("Step 6: GTM Ready \u2014 Package the Solution", level=1)
p = doc.add_paragraph()
run = p.add_run("PHASE: GTM READY")
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(0x7C, 0x3A, 0xED)
run.bold = True

add_purple_bar(doc)

doc.add_heading("What to Do", level=3)
add_bullet(doc, "Finalize the demo script and conduct full run-throughs.")
add_bullet(doc, 'Prepare the "AI vs. Non-AI" talk track \u2014 1\u20132 lines of competitive differentiation explaining what AI enables vs. the status quo.')
add_bullet(doc, "Create sales enablement materials: solution one-pager, pricing template, engagement model.")
add_bullet(doc, "Identify best-fit buyers and map to target accounts.")
add_bullet(doc, "Register the solution in the Global Demo Catalog with domain mapping and status update.")

doc.add_heading("Required Deliverables", level=3)
add_bullet(doc, "Completed demo with polished user experience")
add_bullet(doc, "One-pager / solution brief for sellers (see Demo One-Slider format)")
add_bullet(doc, 'Competitive talk track (e.g., "With AI, we automate extraction + exceptioning; without AI, manual parsing drives delay + errors.")')
add_bullet(doc, "Pricing template and engagement model")
add_bullet(doc, "Defined target accounts and buyer personas")
add_bullet(doc, "Backlog reference and asset links documented")
add_bullet(doc, "SmartSuite alignment confirmed")

add_role_line(doc, ["AI Champion", "Domain Leader", "Sales Team", "Program Manager"])
doc.add_page_break()

# ============================================================
# STEP 7
# ============================================================
doc.add_heading("Step 7: GTM Push \u2014 Showcase & Sales Activation", level=1)
p = doc.add_paragraph()
run = p.add_run("PHASE: GO-TO-MARKET")
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(0x7C, 0x3A, 0xED)
run.bold = True

add_purple_bar(doc)

doc.add_heading("What to Do", level=3)
add_bullet(doc, "Schedule AI Capability Showcases on the bi-weekly cadence, with rotating Domain Leader facilitation.")
add_bullet(doc, "Run Sales Activations targeting specific accounts identified in the GTM Ready phase.")
add_bullet(doc, "Form Sprint Pods \u2014 small, focused teams for active pursuit opportunities.")
add_bullet(doc, "Present to clients, demonstrate the solution live, and gather real-time feedback.")
add_bullet(doc, "Assign follow-up actions with clear owners and due dates.")
add_bullet(doc, "Send bi-weekly Showcase Communications to the broader organization.")

doc.add_heading("Showcase Planning Requirements", level=3)
add_bullet(doc, "Call Type defined: Sales Activation or AI Capability Showcase")
add_bullet(doc, "Date and Facilitator (rotating Domain Leader) confirmed")
add_bullet(doc, "Topic / Featured Use Case(s) selected")
add_bullet(doc, "Presenter(s) briefed and demo environment tested")
add_bullet(doc, '"AI vs. Non-AI" talk track ready')
add_bullet(doc, "Target accounts and follow-up teams identified")
add_bullet(doc, "Sprint pod members assigned for active pursuits")
add_bullet(doc, "Key decisions, actions, owners, and due dates documented post-showcase")

add_role_line(doc, ["Domain Leader", "Account Partner", "AI Champion", "Business Sponsor"])
doc.add_page_break()

# ============================================================
# STEP 8
# ============================================================
doc.add_heading("Step 8: Track Results & Continuous Improvement", level=1)
p = doc.add_paragraph()
run = p.add_run("PHASE: MEASURE")
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(0x7C, 0x3A, 0xED)
run.bold = True

add_purple_bar(doc)

doc.add_heading("What to Do", level=3)
add_bullet(doc, "Track CTAR results and revenue metrics per solution.")
add_bullet(doc, "Monitor demo maturity progression: Idea \u2192 In Build \u2192 Demo-Ready \u2192 GTM-Ready \u2192 In Pursuit.")
add_bullet(doc, "Capture client feedback, feature requests, and enhancement opportunities.")
add_bullet(doc, "Feed learnings back into the next intake cycle for continuous improvement.")
add_bullet(doc, "Identify AI Infusion opportunities in existing client engagements.")

doc.add_heading("Requirements", level=3)
add_bullet(doc, "Showcase Tracker updated with outcomes after every showcase/activation")
add_bullet(doc, "Domain Metrics Dashboard maintained (solution counts across 10 domains)")
add_bullet(doc, "Revenue and pipeline attribution tracked per solution")
add_bullet(doc, "Continuous feedback loop established back to AI Champions")
add_bullet(doc, "Infusion project enhancements communicated via Account Partners and Program Managers")

doc.add_heading("The Continuous Loop", level=3)
p = doc.add_paragraph(
    "Results from go-to-market activities feed directly back into Step 1, "
    "creating a virtuous cycle. Client feedback generates new use case ideas, "
    "successful demos inspire adjacent solutions, and market learnings sharpen "
    "future intake submissions. The goal is not a one-time launch but a "
    "continuously improving portfolio of AI-powered solutions."
)
p.runs[0].font.size = Pt(11)

add_role_line(doc, ["Program Manager", "AI Lab Leadership", "Domain Leader", "AI Champion"])

# ============================================================
# SAVE
# ============================================================
output_path = "/home/user/Intake-to-GTM/AI-Use-Case-to-GTM-Step-Guide.docx"
doc.save(output_path)
print(f"DOCX saved to: {output_path}")
