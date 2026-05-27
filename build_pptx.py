"""
AI-Driven Embedded Product Lifecycle OrchestratorV3 — Professional PPT Builder
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import pptx.oxml.ns as nsmap
from lxml import etree
import copy

# ── Brand palette ──────────────────────────────────────────────────────────────
NAVY        = RGBColor(0x0A, 0x1F, 0x44)   # deep navy
BLUE        = RGBColor(0x00, 0x72, 0xC6)   # primary accent blue
CYAN        = RGBColor(0x00, 0xB4, 0xD8)   # secondary accent
DARK_GRAY   = RGBColor(0x1E, 0x1E, 0x2E)   # near-black bg
MID_GRAY    = RGBColor(0x2D, 0x3A, 0x4A)   # card / panel bg
LIGHT_GRAY  = RGBColor(0xD0, 0xD8, 0xE8)   # body text on dark
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
ORANGE      = RGBColor(0xFF, 0x7A, 0x00)   # highlight / warning
GREEN       = RGBColor(0x00, 0xC8, 0x8E)   # success / positive
PURPLE      = RGBColor(0x7B, 0x5E, 0xF8)   # AI/ML accent

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H

BLANK_LAYOUT = prs.slide_layouts[6]   # completely blank


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def add_filled_rect(slide, l, t, w, h, fill_color, alpha=None):
    shape = slide.shapes.add_shape(1, l, t, w, h)   # MSO_SHAPE_TYPE.RECTANGLE = 1
    shape.line.fill.background()
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.color.rgb = fill_color
    return shape


def add_text_box(slide, text, l, t, w, h,
                 font_size=18, bold=False, color=WHITE,
                 align=PP_ALIGN.LEFT, italic=False, wrap=True):
    txBox = slide.shapes.add_textbox(l, t, w, h)
    txBox.word_wrap = wrap
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = "Calibri"
    return txBox


def add_multiline_text(slide, lines, l, t, w, h,
                       font_size=16, color=WHITE, bold_first=False,
                       align=PP_ALIGN.LEFT, line_spacing=None):
    """lines: list of (text, bold, font_size_override)"""
    txBox = slide.shapes.add_textbox(l, t, w, h)
    txBox.word_wrap = True
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, item in enumerate(lines):
        if isinstance(item, str):
            text, bold, fs = item, (i == 0 and bold_first), font_size
        else:
            text = item[0]
            bold = item[1] if len(item) > 1 else False
            fs   = item[2] if len(item) > 2 else font_size
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.size = Pt(fs)
        run.font.bold = bold
        run.font.color.rgb = color
        run.font.name = "Calibri"
    return txBox


def full_bg(slide, color):
    add_filled_rect(slide, 0, 0, SLIDE_W, SLIDE_H, color)


def accent_bar(slide, color=BLUE, height=Inches(0.06)):
    add_filled_rect(slide, 0, 0, SLIDE_W, height, color)


def bottom_bar(slide, color=NAVY):
    add_filled_rect(slide, 0, SLIDE_H - Inches(0.45), SLIDE_W, Inches(0.45), color)


def slide_number(slide, num, total, color=LIGHT_GRAY):
    add_text_box(slide, f"{num} / {total}",
                 SLIDE_W - Inches(1.1), SLIDE_H - Inches(0.4),
                 Inches(1.0), Inches(0.35),
                 font_size=11, color=color, align=PP_ALIGN.RIGHT)


def section_tag(slide, label, color=BLUE):
    tb = add_filled_rect(slide, Inches(0.35), Inches(0.55),
                         Inches(2.5), Inches(0.32), color)
    add_text_box(slide, label.upper(),
                 Inches(0.35), Inches(0.55), Inches(2.5), Inches(0.32),
                 font_size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def divider(slide, y, color=BLUE, width_frac=0.92):
    w = SLIDE_W * width_frac
    left = (SLIDE_W - w) / 2
    add_filled_rect(slide, left, y, w, Inches(0.018), color)


# ══════════════════════════════════════════════════════════════════════════════
# TOTAL SLIDES
TOTAL = 18

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — COVER
# ══════════════════════════════════════════════════════════════════════════════
def slide_cover():
    sl = prs.slides.add_slide(BLANK_LAYOUT)
    full_bg(sl, DARK_GRAY)
    # gradient-like layered rects
    add_filled_rect(sl, 0, 0, SLIDE_W * 0.55, SLIDE_H, NAVY)
    add_filled_rect(sl, 0, 0, Inches(0.08), SLIDE_H, BLUE)
    # right-side geometric accent
    add_filled_rect(sl, SLIDE_W * 0.55, 0, SLIDE_W * 0.45, SLIDE_H, DARK_GRAY)
    add_filled_rect(sl, SLIDE_W * 0.6, Inches(0.8), SLIDE_W * 0.38, Inches(0.006), BLUE)
    add_filled_rect(sl, SLIDE_W * 0.6, Inches(5.8), SLIDE_W * 0.38, Inches(0.006), CYAN)

    # Decorative circles (simulate tech feel)
    for cx, cy, sz, clr in [
        (SLIDE_W * 0.82, Inches(1.8), Inches(2.2), MID_GRAY),
        (SLIDE_W * 0.92, Inches(4.5), Inches(1.5), MID_GRAY),
        (SLIDE_W * 0.68, Inches(3.5), Inches(0.9), RGBColor(0x00,0x40,0x80)),
    ]:
        c = sl.shapes.add_shape(9, cx - sz/2, cy - sz/2, sz, sz)   # oval = 9
        c.fill.solid(); c.fill.fore_color.rgb = clr
        c.line.color.rgb = clr

    # Tag line small rect
    add_filled_rect(sl, Inches(0.45), Inches(1.5), Inches(0.06), Inches(1.0), CYAN)

    # Title
    add_text_box(sl, "AI-DRIVEN EMBEDDED",
                 Inches(0.6), Inches(1.3), Inches(7.5), Inches(0.85),
                 font_size=42, bold=True, color=WHITE)
    add_text_box(sl, "PRODUCT LIFECYCLE",
                 Inches(0.6), Inches(2.08), Inches(7.5), Inches(0.85),
                 font_size=42, bold=True, color=WHITE)
    add_text_box(sl, "ORCHESTRATOR",
                 Inches(0.6), Inches(2.85), Inches(7.5), Inches(0.85),
                 font_size=42, bold=True, color=CYAN)

    # Version pill
    add_filled_rect(sl, Inches(0.6), Inches(3.75), Inches(1.3), Inches(0.38), BLUE)
    add_text_box(sl, "VERSION 3.0",
                 Inches(0.6), Inches(3.75), Inches(1.3), Inches(0.38),
                 font_size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Sub-title
    add_text_box(sl,
                 "Autonomous AI orchestration across the full embedded\n"
                 "product lifecycle — from concept to end-of-life.",
                 Inches(0.6), Inches(4.25), Inches(6.8), Inches(0.9),
                 font_size=16, color=LIGHT_GRAY)

    # Date & confidential
    add_text_box(sl, "May 2026  |  CONFIDENTIAL",
                 Inches(0.6), Inches(5.35), Inches(4.0), Inches(0.4),
                 font_size=12, color=RGBColor(0x80,0x90,0xA8))

    bottom_bar(sl)
    slide_number(sl, 1, TOTAL)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — TABLE OF CONTENTS
# ══════════════════════════════════════════════════════════════════════════════
def slide_toc():
    sl = prs.slides.add_slide(BLANK_LAYOUT)
    full_bg(sl, DARK_GRAY)
    accent_bar(sl, BLUE)
    add_text_box(sl, "Table of Contents", Inches(0.5), Inches(0.35), Inches(12), Inches(0.55),
                 font_size=28, bold=True, color=WHITE)
    divider(sl, Inches(1.05))

    sections = [
        ("01", "Executive Summary",              "Vision, problem statement & value proposition"),
        ("02", "Problem & Market Context",        "Pain points in embedded product lifecycles"),
        ("03", "Solution Overview",               "What is AEPLO V3 and how it works"),
        ("04", "System Architecture",             "End-to-end platform design"),
        ("05", "Core AI Modules",                 "Six pillars of AI-driven orchestration"),
        ("06", "Lifecycle Coverage",              "Requirements → Design → Develop → Test → Deploy → EOL"),
        ("07", "Key Features & Capabilities",     "Differentiating capabilities deep-dive"),
        ("08", "Integration Ecosystem",           "Tools, platforms & partner integrations"),
        ("09", "Security & Compliance",           "Functional safety, ITAR, ISO 26262, IEC 61508"),
        ("10", "Performance Metrics & KPIs",      "Quantified business & engineering outcomes"),
        ("11", "Case Studies",                    "Real-world deployments & results"),
        ("12", "Technology Roadmap",              "V3 now → V4 future milestones"),
        ("13", "Team & Governance",               "Core team, advisory board, open-source model"),
        ("14", "Investment & Pricing",            "Engagement model & ROI framework"),
        ("15", "Call to Action",                  "Next steps & pilot programme"),
    ]

    cols = [sections[:8], sections[8:]]
    for ci, col in enumerate(cols):
        lx = Inches(0.5 + ci * 6.4)
        for ri, (num, title, desc) in enumerate(col):
            ty = Inches(1.3 + ri * 0.72)
            add_filled_rect(sl, lx, ty, Inches(0.45), Inches(0.45),
                            BLUE if ci == 0 else PURPLE)
            add_text_box(sl, num, lx, ty, Inches(0.45), Inches(0.45),
                         font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            add_text_box(sl, title,
                         lx + Inches(0.55), ty, Inches(2.6), Inches(0.28),
                         font_size=13, bold=True, color=WHITE)
            add_text_box(sl, desc,
                         lx + Inches(0.55), ty + Inches(0.27), Inches(2.9), Inches(0.28),
                         font_size=10, color=LIGHT_GRAY)

    bottom_bar(sl)
    slide_number(sl, 2, TOTAL)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — EXECUTIVE SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
def slide_exec_summary():
    sl = prs.slides.add_slide(BLANK_LAYOUT)
    full_bg(sl, DARK_GRAY)
    accent_bar(sl, CYAN)
    section_tag(sl, "Executive Summary", CYAN)
    add_text_box(sl, "Vision & Value Proposition", Inches(0.5), Inches(0.9), Inches(12), Inches(0.55),
                 font_size=28, bold=True, color=WHITE)
    divider(sl, Inches(1.52))

    # Vision statement
    add_filled_rect(sl, Inches(0.4), Inches(1.65), Inches(12.5), Inches(1.05), MID_GRAY)
    add_text_box(sl,
                 "\"AEPLO V3 is the first full-stack autonomous AI platform that orchestrates "
                 "every stage of the embedded product lifecycle — cutting development cycles by 40%, "
                 "reducing post-release defects by 60%, and enabling continuous, self-optimising delivery.\"",
                 Inches(0.6), Inches(1.7), Inches(12.1), Inches(0.95),
                 font_size=15, italic=True, color=WHITE)

    # Stat cards
    stats = [
        (GREEN,  "40%",  "Faster Time-\nto-Market"),
        (BLUE,   "60%",  "Defect Reduction\n(post-release)"),
        (ORANGE, "3×",   "Engineering\nProductivity"),
        (PURPLE, "80%",  "Automation of\nRoutine Tasks"),
    ]
    for i, (clr, val, lbl) in enumerate(stats):
        lx = Inches(0.4 + i * 3.15)
        add_filled_rect(sl, lx, Inches(2.9), Inches(2.9), Inches(1.5), MID_GRAY)
        add_filled_rect(sl, lx, Inches(2.9), Inches(2.9), Inches(0.06), clr)
        add_text_box(sl, val, lx, Inches(3.0), Inches(2.9), Inches(0.7),
                     font_size=36, bold=True, color=clr, align=PP_ALIGN.CENTER)
        add_text_box(sl, lbl, lx, Inches(3.68), Inches(2.9), Inches(0.6),
                     font_size=13, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)

    # Bullets
    bullets = [
        "Targets embedded systems across automotive, aerospace, industrial IoT, medical devices & consumer electronics.",
        "Built on a model-agnostic, multi-agent AI framework: integrates LLMs, formal verification, and ML-based testing.",
        "Enterprise-grade: compliant with ISO 26262, IEC 61508, DO-178C, MISRA C, and GDPR/ITAR requirements.",
        "SaaS + on-premise hybrid deployment; plugs directly into existing CI/CD pipelines and tool chains.",
    ]
    for i, b in enumerate(bullets):
        add_filled_rect(sl, Inches(0.4), Inches(4.65 + i * 0.56), Inches(0.12), Inches(0.35), CYAN)
        add_text_box(sl, b, Inches(0.65), Inches(4.65 + i * 0.56), Inches(12.2), Inches(0.45),
                     font_size=13, color=LIGHT_GRAY)

    bottom_bar(sl)
    slide_number(sl, 3, TOTAL)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — PROBLEM & MARKET CONTEXT
# ══════════════════════════════════════════════════════════════════════════════
def slide_problem():
    sl = prs.slides.add_slide(BLANK_LAYOUT)
    full_bg(sl, DARK_GRAY)
    accent_bar(sl, ORANGE)
    section_tag(sl, "Problem & Market", ORANGE)
    add_text_box(sl, "The Embedded Lifecycle Is Broken", Inches(0.5), Inches(0.9), Inches(12), Inches(0.55),
                 font_size=28, bold=True, color=WHITE)
    divider(sl, Inches(1.52), color=ORANGE)

    pains = [
        (ORANGE, "⏱  Schedule Overruns",
         "72% of embedded projects miss their release date.\n"
         "Manual requirement tracing, integration failures & late bug discovery are the top causes."),
        (BLUE,   "🐛  Defect Escape Rate",
         "Embedded defects found post-deployment cost 100× more to fix.\n"
         "Traditional test strategies miss 35–45% of critical edge-case failures."),
        (PURPLE, "🔄  Tool-Chain Fragmentation",
         "Engineering teams juggle 12–20 disparate tools with no unified data model.\n"
         "Context-switching overhead consumes 30% of engineering bandwidth."),
        (GREEN,  "📋  Compliance Burden",
         "Meeting IEC 61508 SIL-3 or ISO 26262 ASIL-D requires thousands of hours of manual artifact generation.\n"
         "Traceability gaps lead to costly re-certifications."),
        (CYAN,   "🔌  Hardware-Software Co-Design Gap",
         "Hardware and software teams operate in silos.\n"
         "Late hardware bring-up causes 60% of integration delays in complex SoC designs."),
        (RGBColor(0xFF,0x40,0x80), "📈  Market Pressure",
         "Embedded software content doubles every 3 years.\n"
         "Competitors shipping AI-enhanced products 2× faster — traditional methods cannot keep pace."),
    ]

    for i, (clr, title, desc) in enumerate(pains):
        col = i % 2; row = i // 2
        lx = Inches(0.4 + col * 6.45)
        ty = Inches(1.72 + row * 1.78)
        add_filled_rect(sl, lx, ty, Inches(6.1), Inches(1.6), MID_GRAY)
        add_filled_rect(sl, lx, ty, Inches(0.07), Inches(1.6), clr)
        add_text_box(sl, title, lx + Inches(0.18), ty + Inches(0.1), Inches(5.7), Inches(0.35),
                     font_size=14, bold=True, color=clr)
        add_text_box(sl, desc, lx + Inches(0.18), ty + Inches(0.42), Inches(5.7), Inches(1.0),
                     font_size=11.5, color=LIGHT_GRAY)

    bottom_bar(sl)
    slide_number(sl, 4, TOTAL)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — SOLUTION OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
def slide_solution():
    sl = prs.slides.add_slide(BLANK_LAYOUT)
    full_bg(sl, DARK_GRAY)
    accent_bar(sl, GREEN)
    section_tag(sl, "Solution Overview", GREEN)
    add_text_box(sl, "What Is AEPLO V3?", Inches(0.5), Inches(0.9), Inches(12), Inches(0.55),
                 font_size=28, bold=True, color=WHITE)
    divider(sl, Inches(1.52), color=GREEN)

    # Central description
    add_filled_rect(sl, Inches(0.4), Inches(1.65), Inches(7.8), Inches(1.2), MID_GRAY)
    add_text_box(sl,
                 "AEPLO V3 is a multi-agent AI orchestration platform that sits at the intersection of "
                 "your embedded development tool chain and a suite of purpose-built AI engines. "
                 "It continuously monitors, analyzes, and acts across the full product lifecycle — "
                 "from requirements ingestion through field telemetry-driven EOL prediction.",
                 Inches(0.6), Inches(1.72), Inches(7.4), Inches(1.08),
                 font_size=13, color=LIGHT_GRAY)

    # Three pillars
    pillars = [
        (BLUE,   "PERCEIVE",  "Unified data ingestion from all tools:\nJIRA, Git, CI/CD, HW simulators,\nfield telemetry, regulatory databases."),
        (PURPLE, "REASON",    "Multi-agent LLM + formal methods:\nrequirement analysis, risk scoring,\nchange-impact prediction, root-cause AI."),
        (GREEN,  "ACT",       "Closed-loop automation:\nauto-generate tests, patches, docs,\ncertification artifacts & release gates."),
    ]
    for i, (clr, title, desc) in enumerate(pillars):
        lx = Inches(0.4 + i * 4.15)
        add_filled_rect(sl, lx, Inches(3.05), Inches(3.9), Inches(2.7), MID_GRAY)
        add_filled_rect(sl, lx, Inches(3.05), Inches(3.9), Inches(0.06), clr)
        add_text_box(sl, title, lx, Inches(3.18), Inches(3.9), Inches(0.45),
                     font_size=18, bold=True, color=clr, align=PP_ALIGN.CENTER)
        add_text_box(sl, desc, lx + Inches(0.2), Inches(3.7), Inches(3.5), Inches(1.9),
                     font_size=13, color=LIGHT_GRAY)

    # Right side — key numbers
    add_filled_rect(sl, Inches(8.5), Inches(1.65), Inches(4.4), Inches(4.15), MID_GRAY)
    add_text_box(sl, "AEPLO V3 AT A GLANCE", Inches(8.6), Inches(1.75), Inches(4.2), Inches(0.35),
                 font_size=12, bold=True, color=CYAN, align=PP_ALIGN.CENTER)
    divider(sl, Inches(2.18), color=BLUE, width_frac=0.32)

    glance = [
        ("6",          "Core AI Agents"),
        ("18+",        "Tool Integrations"),
        ("< 5 min",    "On-boarding time"),
        ("Real-time",  "Lifecycle Visibility"),
        ("ASIL-D",     "Max Safety Integrity"),
        ("Cloud + On-prem", "Deployment Options"),
        ("REST / gRPC","API Surface"),
        ("< 200 ms",   "Agent Response Latency"),
    ]
    for i, (val, lbl) in enumerate(glance):
        ty = Inches(2.35 + i * 0.45)
        add_text_box(sl, val,  Inches(8.6),  ty, Inches(1.9), Inches(0.38),
                     font_size=13, bold=True, color=WHITE)
        add_text_box(sl, lbl,  Inches(10.55), ty, Inches(2.25), Inches(0.38),
                     font_size=12, color=LIGHT_GRAY)

    bottom_bar(sl)
    slide_number(sl, 5, TOTAL)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — SYSTEM ARCHITECTURE
# ══════════════════════════════════════════════════════════════════════════════
def slide_architecture():
    sl = prs.slides.add_slide(BLANK_LAYOUT)
    full_bg(sl, DARK_GRAY)
    accent_bar(sl, BLUE)
    section_tag(sl, "Architecture", BLUE)
    add_text_box(sl, "End-to-End Platform Architecture", Inches(0.5), Inches(0.9), Inches(12), Inches(0.55),
                 font_size=28, bold=True, color=WHITE)
    divider(sl, Inches(1.52))

    # Layer labels
    layers = [
        (NAVY,    "DATA INGESTION LAYER",      "Git · JIRA · CI/CD · HW Sim · Field Telemetry · Regs"),
        (MID_GRAY,"UNIFIED KNOWLEDGE GRAPH",   "Semantic Graph DB · Traceability Engine · Version History"),
        (RGBColor(0x0D,0x2B,0x55), "MULTI-AGENT AI CORE",
                                              "Req Agent · Design Agent · Test Agent · Defect Agent · Cert Agent · EOL Agent"),
        (MID_GRAY,"ORCHESTRATION & AUTOMATION","CI/CD Hooks · Auto-PR · Change Gating · Release Manager"),
        (NAVY,    "DELIVERY & COMPLIANCE",     "Dashboards · Audit Trails · SBOM · Cert Packages · APIs"),
    ]

    for i, (bg, title, sublabel) in enumerate(layers):
        ty = Inches(1.72 + i * 1.0)
        add_filled_rect(sl, Inches(0.35), ty, Inches(12.6), Inches(0.88), bg)

        # Left label
        add_filled_rect(sl, Inches(0.35), ty, Inches(2.8), Inches(0.88),
                        BLUE if i % 2 == 0 else PURPLE)
        add_text_box(sl, title, Inches(0.4), ty + Inches(0.22), Inches(2.7), Inches(0.45),
                     font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text_box(sl, sublabel, Inches(3.3), ty + Inches(0.2), Inches(9.5), Inches(0.5),
                     font_size=12, color=LIGHT_GRAY)

    # Arrows between layers (simulated with small rects)
    for i in range(4):
        ty = Inches(1.72 + (i+1)*1.0) - Inches(0.04)
        add_filled_rect(sl, Inches(6.2), ty, Inches(0.9), Inches(0.08), CYAN)

    bottom_bar(sl)
    slide_number(sl, 6, TOTAL)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — CORE AI MODULES
# ══════════════════════════════════════════════════════════════════════════════
def slide_ai_modules():
    sl = prs.slides.add_slide(BLANK_LAYOUT)
    full_bg(sl, DARK_GRAY)
    accent_bar(sl, PURPLE)
    section_tag(sl, "Core AI Modules", PURPLE)
    add_text_box(sl, "Six Pillars of AI Orchestration", Inches(0.5), Inches(0.9), Inches(12), Inches(0.55),
                 font_size=28, bold=True, color=WHITE)
    divider(sl, Inches(1.52), color=PURPLE)

    modules = [
        (BLUE,   "REQ-AI",
         "Requirement Intelligence Agent",
         "• NLP-driven req parsing & disambiguation\n• Automatic traceability matrix generation\n• Change-impact scoring & conflict detection\n• Supports SysML, DOORS, JIRA, plain text"),
        (GREEN,  "DESIGN-AI",
         "HW/SW Co-Design Assistant",
         "• Architecture pattern recommendation\n• AI-powered memory / power budgeting\n• Interface contract verification\n• RTOS & BSP configuration optimisation"),
        (CYAN,   "TEST-AI",
         "Autonomous Test Orchestrator",
         "• Coverage-guided test case generation\n• Model-in-the-loop & HIL orchestration\n• Regression triage & smart re-run\n• Mutation testing & fault injection"),
        (ORANGE, "DEFECT-AI",
         "Predictive Defect & Root-Cause Engine",
         "• Static + dynamic combined analysis\n• ML-based defect prediction (commit-level)\n• Automated root-cause hypotheses\n• Cross-release defect pattern mining"),
        (PURPLE, "CERT-AI",
         "Certification Autopilot",
         "• ISO 26262 / IEC 61508 artefact generation\n• Gap analysis against standard checklists\n• Audit-ready evidence packaging\n• Regulatory change monitoring"),
        (RGBColor(0xFF,0x40,0x80), "EOL-AI",
         "End-of-Life & Sustainability Agent",
         "• Component obsolescence early-warning\n• Field telemetry → remaining-life prediction\n• SBOM lifecycle management\n• Circular-economy optimisation scoring"),
    ]

    for i, (clr, tag, title, bullets) in enumerate(modules):
        col = i % 3; row = i // 3
        lx = Inches(0.35 + col * 4.32)
        ty = Inches(1.72 + row * 2.65)
        add_filled_rect(sl, lx, ty, Inches(4.1), Inches(2.5), MID_GRAY)
        add_filled_rect(sl, lx, ty, Inches(4.1), Inches(0.06), clr)

        add_filled_rect(sl, lx + Inches(0.15), ty + Inches(0.15),
                        Inches(0.75), Inches(0.35), clr)
        add_text_box(sl, tag, lx + Inches(0.15), ty + Inches(0.15),
                     Inches(0.75), Inches(0.35),
                     font_size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        add_text_box(sl, title, lx + Inches(1.02), ty + Inches(0.15), Inches(2.9), Inches(0.35),
                     font_size=12, bold=True, color=WHITE)
        add_text_box(sl, bullets, lx + Inches(0.15), ty + Inches(0.62), Inches(3.8), Inches(1.75),
                     font_size=11, color=LIGHT_GRAY)

    bottom_bar(sl)
    slide_number(sl, 7, TOTAL)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — LIFECYCLE COVERAGE
# ══════════════════════════════════════════════════════════════════════════════
def slide_lifecycle():
    sl = prs.slides.add_slide(BLANK_LAYOUT)
    full_bg(sl, DARK_GRAY)
    accent_bar(sl, GREEN)
    section_tag(sl, "Lifecycle Coverage", GREEN)
    add_text_box(sl, "Full Embedded Product Lifecycle", Inches(0.5), Inches(0.9), Inches(12), Inches(0.55),
                 font_size=28, bold=True, color=WHITE)
    divider(sl, Inches(1.52), color=GREEN)

    phases = [
        (BLUE,   "1. REQUIREMENTS",    "REQ-AI",     "Ingest · Parse · Trace · Validate"),
        (PURPLE, "2. ARCHITECTURE",    "DESIGN-AI",  "Model · Verify · Optimise · Budget"),
        (GREEN,  "3. DEVELOPMENT",     "DEFECT-AI",  "Predict · Analyse · Suggest · Review"),
        (CYAN,   "4. TESTING",         "TEST-AI",    "Generate · Execute · Triage · Report"),
        (ORANGE, "5. CERTIFICATION",   "CERT-AI",    "Audit · Generate · Gap-fill · Submit"),
        (RGBColor(0xFF,0x40,0x80), "6. DEPLOYMENT", "EOL-AI", "Monitor · Update · SBOM · Secure"),
        (RGBColor(0x90,0x30,0xEF), "7. FIELD OPS",  "EOL-AI", "Telemetry · Predict · Alert · Plan"),
        (ORANGE, "8. END-OF-LIFE",     "EOL-AI",     "Retire · Recycle · Archive · Comply"),
    ]

    # Timeline arrow
    ty_arrow = Inches(4.55)
    add_filled_rect(sl, Inches(0.35), ty_arrow, Inches(12.6), Inches(0.08), MID_GRAY)
    # arrowhead
    add_filled_rect(sl, SLIDE_W - Inches(0.55), ty_arrow - Inches(0.1), Inches(0.2), Inches(0.28), MID_GRAY)

    for i, (clr, phase, agent, actions) in enumerate(phases):
        lx = Inches(0.35 + i * 1.55)
        # dot on timeline
        dot_sz = Inches(0.22)
        c = sl.shapes.add_shape(9, lx + Inches(0.54), ty_arrow - dot_sz/2, dot_sz, dot_sz)
        c.fill.solid(); c.fill.fore_color.rgb = clr
        c.line.color.rgb = clr

        if i % 2 == 0:
            # above timeline
            bty = Inches(1.72)
            add_filled_rect(sl, lx, bty, Inches(1.45), Inches(2.65), MID_GRAY)
            add_filled_rect(sl, lx, bty, Inches(1.45), Inches(0.05), clr)
            add_text_box(sl, phase, lx + Inches(0.06), bty + Inches(0.1), Inches(1.33), Inches(0.52),
                         font_size=9.5, bold=True, color=clr)
            add_filled_rect(sl, lx + Inches(0.06), bty + Inches(0.65), Inches(0.8), Inches(0.26), clr)
            add_text_box(sl, agent, lx + Inches(0.06), bty + Inches(0.65), Inches(0.8), Inches(0.26),
                         font_size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            add_text_box(sl, actions.replace(' · ', '\n'), lx + Inches(0.06), bty + Inches(1.0),
                         Inches(1.33), Inches(1.5), font_size=9, color=LIGHT_GRAY)
            # connector line
            add_filled_rect(sl, lx + Inches(0.65), bty + Inches(2.65), Inches(0.06), Inches(1.83), LIGHT_GRAY)
        else:
            # below timeline
            bty = Inches(4.8)
            add_filled_rect(sl, lx, bty, Inches(1.45), Inches(2.45), MID_GRAY)
            add_filled_rect(sl, lx, bty, Inches(1.45), Inches(0.05), clr)
            add_text_box(sl, phase, lx + Inches(0.06), bty + Inches(0.1), Inches(1.33), Inches(0.52),
                         font_size=9.5, bold=True, color=clr)
            add_filled_rect(sl, lx + Inches(0.06), bty + Inches(0.65), Inches(0.8), Inches(0.26), clr)
            add_text_box(sl, agent, lx + Inches(0.06), bty + Inches(0.65), Inches(0.8), Inches(0.26),
                         font_size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            add_text_box(sl, actions.replace(' · ', '\n'), lx + Inches(0.06), bty + Inches(1.0),
                         Inches(1.33), Inches(1.3), font_size=9, color=LIGHT_GRAY)
            add_filled_rect(sl, lx + Inches(0.65), ty_arrow + Inches(0.08), Inches(0.06), Inches(0.16), LIGHT_GRAY)

    bottom_bar(sl)
    slide_number(sl, 8, TOTAL)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 9 — KEY FEATURES
# ══════════════════════════════════════════════════════════════════════════════
def slide_features():
    sl = prs.slides.add_slide(BLANK_LAYOUT)
    full_bg(sl, DARK_GRAY)
    accent_bar(sl, CYAN)
    section_tag(sl, "Key Features", CYAN)
    add_text_box(sl, "Differentiating Capabilities", Inches(0.5), Inches(0.9), Inches(12), Inches(0.55),
                 font_size=28, bold=True, color=WHITE)
    divider(sl, Inches(1.52))

    features = [
        (BLUE,   "Continuous Traceability",
                 "Bi-directional live traceability from market requirement → silicon → firmware → field data. "
                 "Automatic impact analysis for any change in under 3 seconds."),
        (GREEN,  "AI-Generated Test Suites",
                 "Automatically generates unit, integration and system test cases from requirements and design models. "
                 "Achieves MC/DC coverage targets without manual test authoring."),
        (PURPLE, "Predictive Defect Scoring",
                 "Commit-level defect probability scoring using ensemble ML trained on 50M+ embedded defect records. "
                 "Flags high-risk changes before merge."),
        (CYAN,   "One-Click Certification Packages",
                 "Generates complete IEC 61508 / ISO 26262 safety cases, FMEA, FMEDA, and DFA artefacts automatically. "
                 "Reduces certification prep from weeks to hours."),
        (ORANGE, "Hardware-in-the-Loop Orchestration",
                 "AI-driven HIL scheduling with intelligent test prioritisation. "
                 "Detects hardware anomalies via telemetry signatures and predicts board failures 72 h in advance."),
        (RGBColor(0xFF,0x40,0x80), "Natural Language Interface",
                 "Engineers query the full lifecycle using plain English. "
                 "AEPLO V3 answers: What changed since last sprint that could affect ASIL-B? -- instantly."),
    ]

    for i, (clr, title, desc) in enumerate(features):
        row = i // 2; col = i % 2
        lx = Inches(0.35 + col * 6.45)
        ty = Inches(1.72 + row * 1.85)
        add_filled_rect(sl, lx, ty, Inches(6.1), Inches(1.7), MID_GRAY)
        # accent circle
        c = sl.shapes.add_shape(9, lx + Inches(0.15), ty + Inches(0.55), Inches(0.55), Inches(0.55))
        c.fill.solid(); c.fill.fore_color.rgb = clr; c.line.color.rgb = clr
        add_text_box(sl, "✓", lx + Inches(0.15), ty + Inches(0.55), Inches(0.55), Inches(0.55),
                     font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text_box(sl, title, lx + Inches(0.85), ty + Inches(0.12), Inches(5.0), Inches(0.38),
                     font_size=14, bold=True, color=clr)
        add_text_box(sl, desc, lx + Inches(0.85), ty + Inches(0.5), Inches(5.05), Inches(1.05),
                     font_size=11.5, color=LIGHT_GRAY)

    bottom_bar(sl)
    slide_number(sl, 9, TOTAL)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 10 — INTEGRATION ECOSYSTEM
# ══════════════════════════════════════════════════════════════════════════════
def slide_integrations():
    sl = prs.slides.add_slide(BLANK_LAYOUT)
    full_bg(sl, DARK_GRAY)
    accent_bar(sl, BLUE)
    section_tag(sl, "Integrations", BLUE)
    add_text_box(sl, "Integration Ecosystem", Inches(0.5), Inches(0.9), Inches(12), Inches(0.55),
                 font_size=28, bold=True, color=WHITE)
    divider(sl, Inches(1.52))

    categories = [
        (BLUE,   "REQUIREMENTS & ALM",     "JIRA · IBM DOORS · Polarion · Codebeamer · Azure DevOps"),
        (GREEN,  "SOURCE CONTROL",         "Git / GitHub · GitLab · Bitbucket · Perforce Helix"),
        (PURPLE, "BUILD & CI/CD",          "Jenkins · GitLab CI · GitHub Actions · Bamboo · TeamCity"),
        (CYAN,   "MODELLING & SIMULATION", "MATLAB/Simulink · Enterprise Architect · Rhapsody · ANSYS"),
        (ORANGE, "STATIC ANALYSIS",        "Polyspace · Coverity · SonarQube · MISRA Checker · PRQA"),
        (RGBColor(0xFF,0x40,0x80), "TEST PLATFORMS", "VectorCAST · LDRA · NI TestStand · CAPL · Robot Framework"),
        (BLUE,   "HW DEBUGGERS & JTAG",    "JTAG/SWD · Lauterbach · J-Link · Green Hills Probe"),
        (GREEN,  "SECURITY SCANNING",      "Synopsys Black Duck · FOSSA · Snyk · CycloneDX SBOM"),
        (PURPLE, "CLOUD PLATFORMS",        "AWS · Azure · GCP · Private Cloud · Hybrid Edge"),
        (CYAN,   "COMMUNICATION",          "Slack · MS Teams · PagerDuty · ServiceNow · Email"),
    ]

    for i, (clr, cat, tools) in enumerate(categories):
        col = i % 2; row = i // 2
        lx = Inches(0.35 + col * 6.45)
        ty = Inches(1.72 + row * 1.03)
        add_filled_rect(sl, lx, ty, Inches(6.1), Inches(0.9), MID_GRAY)
        add_filled_rect(sl, lx, ty, Inches(0.06), Inches(0.9), clr)
        add_text_box(sl, cat, lx + Inches(0.2), ty + Inches(0.06), Inches(2.4), Inches(0.32),
                     font_size=11, bold=True, color=clr)
        add_text_box(sl, tools, lx + Inches(0.2), ty + Inches(0.42), Inches(5.7), Inches(0.4),
                     font_size=11, color=LIGHT_GRAY)

    bottom_bar(sl)
    slide_number(sl, 10, TOTAL)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 11 — SECURITY & COMPLIANCE
# ══════════════════════════════════════════════════════════════════════════════
def slide_security():
    sl = prs.slides.add_slide(BLANK_LAYOUT)
    full_bg(sl, DARK_GRAY)
    accent_bar(sl, ORANGE)
    section_tag(sl, "Security & Compliance", ORANGE)
    add_text_box(sl, "Enterprise-Grade Safety & Compliance", Inches(0.5), Inches(0.9), Inches(12), Inches(0.55),
                 font_size=28, bold=True, color=WHITE)
    divider(sl, Inches(1.52), color=ORANGE)

    standards = [
        (ORANGE, "ISO 26262",  "ASIL A–D",  "Automotive functional safety. Full ASIL decomposition\nsupport, FMEA/FTA generation, work-product automation."),
        (BLUE,   "IEC 61508",  "SIL 1–4",   "Industrial functional safety. SIL target verification,\nhardware/software systematic capability assessment."),
        (GREEN,  "DO-178C",    "DAL A–E",   "Avionics software. Requirements traceability, MC/DC\ncoverage, structural coverage analysis automation."),
        (CYAN,   "IEC 62443",  "SL 1–4",    "Industrial cyber-security. Attack surface modelling,\ncyber-FMEA, secure coding enforcement."),
        (PURPLE, "MISRA C/C++","2023",       "Coding standard enforcement with automatic rule\nviolation fix suggestions via DEFECT-AI."),
        (ORANGE, "GDPR / ITAR","Export Ctrl","Data residency controls, audit logging, role-based\naccess, air-gapped deployment option for ITAR work."),
    ]

    for i, (clr, std, level, desc) in enumerate(standards):
        col = i % 3; row = i // 2
        lx = Inches(0.35 + col * 4.3)
        ty = Inches(1.72 + row * 2.35)
        add_filled_rect(sl, lx, ty, Inches(4.05), Inches(2.15), MID_GRAY)
        add_filled_rect(sl, lx, ty, Inches(4.05), Inches(0.05), clr)
        add_text_box(sl, std,   lx + Inches(0.15), ty + Inches(0.12), Inches(2.0), Inches(0.45),
                     font_size=22, bold=True, color=clr)
        add_filled_rect(sl, lx + Inches(2.4), ty + Inches(0.18), Inches(1.35), Inches(0.32), clr)
        add_text_box(sl, level, lx + Inches(2.4), ty + Inches(0.18), Inches(1.35), Inches(0.32),
                     font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text_box(sl, desc,  lx + Inches(0.15), ty + Inches(0.65), Inches(3.7), Inches(1.35),
                     font_size=11, color=LIGHT_GRAY)

    bottom_bar(sl)
    slide_number(sl, 11, TOTAL)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 12 — PERFORMANCE METRICS
# ══════════════════════════════════════════════════════════════════════════════
def slide_metrics():
    sl = prs.slides.add_slide(BLANK_LAYOUT)
    full_bg(sl, DARK_GRAY)
    accent_bar(sl, GREEN)
    section_tag(sl, "Metrics & KPIs", GREEN)
    add_text_box(sl, "Quantified Business & Engineering Outcomes", Inches(0.5), Inches(0.9), Inches(12), Inches(0.55),
                 font_size=28, bold=True, color=WHITE)
    divider(sl, Inches(1.52), color=GREEN)

    # Big metric cards
    big_metrics = [
        (GREEN,  "40%",  "Reduction in\ntime-to-market",     "Validated across 3 automotive OEM pilots"),
        (BLUE,   "60%",  "Fewer post-release\ndefects",      "ML defect predictor + auto-test coverage"),
        (ORANGE, "3×",   "Engineering\nproductivity gain",   "Hours reclaimed from manual artefact work"),
        (PURPLE, "80%",  "Routine task\nautomation rate",    "Docs, traces, test cases, cert artefacts"),
    ]
    for i, (clr, val, lbl, note) in enumerate(big_metrics):
        lx = Inches(0.35 + i * 3.2)
        add_filled_rect(sl, lx, Inches(1.72), Inches(3.0), Inches(2.2), MID_GRAY)
        add_filled_rect(sl, lx, Inches(1.72), Inches(3.0), Inches(0.06), clr)
        add_text_box(sl, val,  lx, Inches(1.85), Inches(3.0), Inches(0.85),
                     font_size=48, bold=True, color=clr, align=PP_ALIGN.CENTER)
        add_text_box(sl, lbl,  lx, Inches(2.72), Inches(3.0), Inches(0.55),
                     font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text_box(sl, note, lx, Inches(3.3), Inches(3.0), Inches(0.52),
                     font_size=10, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)

    # Secondary KPI table
    add_text_box(sl, "Additional KPIs", Inches(0.35), Inches(4.15), Inches(7), Inches(0.38),
                 font_size=14, bold=True, color=CYAN)
    kpis = [
        ("Requirement change turnaround",    "< 2 hours",    "Previously 2–3 days"),
        ("Certification prep time",          "−75%",         "Weeks → hours"),
        ("Test coverage target attainment",  "98%",          "MC/DC for ASIL-C/D"),
        ("Mean-time-to-root-cause",          "−65%",         "DEFECT-AI automated RCA"),
        ("SBOM generation",                  "Real-time",    "Per-commit, every build"),
        ("Component obsolescence alert",     "> 12 months",  "Lead-time warning horizon"),
    ]
    for i, (kpi, val, note) in enumerate(kpis):
        col = i % 2; row = i // 2
        lx = Inches(0.35 + col * 6.5)
        ty = Inches(4.65 + row * 0.7)
        add_filled_rect(sl, lx, ty, Inches(6.1), Inches(0.58), MID_GRAY)
        add_text_box(sl, kpi,  lx + Inches(0.15), ty + Inches(0.1), Inches(3.0), Inches(0.38), font_size=11, color=LIGHT_GRAY)
        add_text_box(sl, val,  lx + Inches(3.2), ty + Inches(0.1),  Inches(1.2), Inches(0.38), font_size=12, bold=True, color=GREEN)
        add_text_box(sl, note, lx + Inches(4.45), ty + Inches(0.1), Inches(1.5), Inches(0.38), font_size=10, color=LIGHT_GRAY)

    bottom_bar(sl)
    slide_number(sl, 12, TOTAL)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 13 — CASE STUDIES
# ══════════════════════════════════════════════════════════════════════════════
def slide_case_studies():
    sl = prs.slides.add_slide(BLANK_LAYOUT)
    full_bg(sl, DARK_GRAY)
    accent_bar(sl, PURPLE)
    section_tag(sl, "Case Studies", PURPLE)
    add_text_box(sl, "Real-World Deployments", Inches(0.5), Inches(0.9), Inches(12), Inches(0.55),
                 font_size=28, bold=True, color=WHITE)
    divider(sl, Inches(1.52), color=PURPLE)

    cases = [
        (BLUE, "Tier-1 Automotive OEM",
         "ADAS ECU Development",
         "Challenge: 14-month schedule overrun on ASIL-D radar controller.\n"
         "Solution: REQ-AI + CERT-AI deployed across 3 supplier teams.",
         [("−40%", "Schedule reduction"), ("ASIL-D", "Achieved on 1st audit"), ("2,400h", "Saved on cert prep")]),
        (GREEN, "Aerospace Avionics OEM",
         "DO-178C DAL-A Flight Management",
         "Challenge: MC/DC coverage was stalling a critical FMS release.\n"
         "Solution: TEST-AI auto-generated 8,000 test cases overnight.",
         [("98.7%", "MC/DC coverage hit"), ("6 weeks", "Saved on test phase"), ("Zero", "Regression escapes")]),
        (ORANGE, "Industrial IoT Platform",
         "IEC 62443 SL-3 Controller",
         "Challenge: Firmware quality issues causing field recalls.\n"
         "Solution: DEFECT-AI deployed in CI; CERT-AI for 62443 gap-fill.",
         [("−62%", "Field defect rate"), ("$4.2M", "Recall cost avoided"), ("3×", "Release cadence")]),
        (PURPLE, "Medical Device Manufacturer",
         "IEC 62304 Class C Software",
         "Challenge: FDA 510(k) submission blocked by traceability gaps.\n"
         "Solution: REQ-AI + CERT-AI rebuilt full traceability in 48 hours.",
         [("48 hrs", "Traceability rebuilt"), ("Approved", "FDA 510(k) outcome"), ("−55%", "QA overhead")]),
    ]

    for i, (clr, org, project, story, stats) in enumerate(cases):
        col = i % 2; row = i // 2
        lx = Inches(0.35 + col * 6.45)
        ty = Inches(1.72 + row * 2.75)
        add_filled_rect(sl, lx, ty, Inches(6.1), Inches(2.55), MID_GRAY)
        add_filled_rect(sl, lx, ty, Inches(6.1), Inches(0.05), clr)
        add_filled_rect(sl, lx, ty + Inches(0.07), Inches(0.06), Inches(2.4), clr)
        add_text_box(sl, org,     lx + Inches(0.2), ty + Inches(0.08), Inches(4.0), Inches(0.35),
                     font_size=13, bold=True, color=WHITE)
        add_text_box(sl, project, lx + Inches(0.2), ty + Inches(0.4),  Inches(5.5), Inches(0.3),
                     font_size=11, italic=True, color=clr)
        add_text_box(sl, story,   lx + Inches(0.2), ty + Inches(0.75), Inches(5.7), Inches(0.75),
                     font_size=11, color=LIGHT_GRAY)
        # stat mini-cards
        for j, (val, lbl) in enumerate(stats):
            sx = lx + Inches(0.2 + j * 1.88)
            add_filled_rect(sl, sx, ty + Inches(1.6), Inches(1.7), Inches(0.75), NAVY)
            add_text_box(sl, val, sx, ty + Inches(1.62), Inches(1.7), Inches(0.35),
                         font_size=16, bold=True, color=clr, align=PP_ALIGN.CENTER)
            add_text_box(sl, lbl, sx, ty + Inches(1.98), Inches(1.7), Inches(0.3),
                         font_size=9, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)

    bottom_bar(sl)
    slide_number(sl, 13, TOTAL)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 14 — TECHNOLOGY ROADMAP
# ══════════════════════════════════════════════════════════════════════════════
def slide_roadmap():
    sl = prs.slides.add_slide(BLANK_LAYOUT)
    full_bg(sl, DARK_GRAY)
    accent_bar(sl, CYAN)
    section_tag(sl, "Technology Roadmap", CYAN)
    add_text_box(sl, "V3 Now → V4 Future Milestones", Inches(0.5), Inches(0.9), Inches(12), Inches(0.55),
                 font_size=28, bold=True, color=WHITE)
    divider(sl, Inches(1.52), color=CYAN)

    roadmap = [
        ("Q1 2026\n(GA)", BLUE,
         ["REQ-AI v3 GA", "CERT-AI ISO 26262 full coverage",
          "18 integrations live", "SaaS EU + US regions"]),
        ("Q2 2026", GREEN,
         ["TEST-AI HIL orchestration v2", "DEFECT-AI commit scoring v3",
          "NL query interface beta", "On-prem k8s packaging"]),
        ("Q3 2026", CYAN,
         ["EOL-AI field telemetry connector", "DESIGN-AI HW co-design beta",
          "DO-178C / DO-254 CERT-AI", "Multi-project portfolio view"]),
        ("Q4 2026", ORANGE,
         ["DESIGN-AI GA with RTOS optimiser", "Federated multi-site support",
          "AI-generated FTA / DFMEA GA", "IEC 62443 full coverage"]),
        ("H1 2027\n(V4 Preview)", PURPLE,
         ["Autonomous on-chip verification agent", "Silicon lifecycle digital twin",
          "Self-healing firmware pipelines", "Quantum-safe crypto toolchain"]),
    ]

    for i, (period, clr, items) in enumerate(roadmap):
        lx = Inches(0.35 + i * 2.55)
        # header
        add_filled_rect(sl, lx, Inches(1.72), Inches(2.38), Inches(0.55), clr)
        add_text_box(sl, period, lx, Inches(1.72), Inches(2.38), Inches(0.55),
                     font_size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        # body
        add_filled_rect(sl, lx, Inches(2.27), Inches(2.38), Inches(4.85), MID_GRAY)
        add_filled_rect(sl, lx, Inches(2.27), Inches(0.06), Inches(4.85), clr)
        for j, item in enumerate(items):
            ty = Inches(2.42 + j * 1.1)
            c = sl.shapes.add_shape(9, lx + Inches(0.12), ty + Inches(0.1),
                                    Inches(0.2), Inches(0.2))
            c.fill.solid(); c.fill.fore_color.rgb = clr; c.line.color.rgb = clr
            add_text_box(sl, item, lx + Inches(0.42), ty, Inches(1.9), Inches(0.9),
                         font_size=11, color=LIGHT_GRAY)

    # "V4 Vision" banner
    add_filled_rect(sl, Inches(0.35), Inches(7.05), Inches(12.6), Inches(0.35), NAVY)
    add_text_box(sl, "V4 VISION: Fully autonomous, self-optimising embedded development with zero-human-in-the-loop for routine engineering tasks.",
                 Inches(0.5), Inches(7.05), Inches(12.3), Inches(0.35),
                 font_size=11, italic=True, color=CYAN)

    bottom_bar(sl)
    slide_number(sl, 14, TOTAL)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 15 — TEAM & GOVERNANCE
# ══════════════════════════════════════════════════════════════════════════════
def slide_team():
    sl = prs.slides.add_slide(BLANK_LAYOUT)
    full_bg(sl, DARK_GRAY)
    accent_bar(sl, BLUE)
    section_tag(sl, "Team & Governance", BLUE)
    add_text_box(sl, "People & Open-Source Model", Inches(0.5), Inches(0.9), Inches(12), Inches(0.55),
                 font_size=28, bold=True, color=WHITE)
    divider(sl, Inches(1.52))

    team = [
        (BLUE,   "CEO / Co-founder",      "15 yrs embedded systems leadership\nFormer VP Engineering, Tier-1 Automotive"),
        (GREEN,  "CTO / Co-founder",      "PhD Computer Science (ML)\nFormer Principal Architect, Aerospace OEM"),
        (PURPLE, "VP Product",            "10 yrs product management in safety-critical\nISO 26262 / DO-178C domain expert"),
        (CYAN,   "Head of AI Research",   "PhD AI & Formal Methods\n12 yrs in autonomous systems verification"),
        (ORANGE, "VP Engineering",        "20 yrs embedded SW engineering\nBuilt 5M+ unit production firmware"),
        (RGBColor(0xFF,0x40,0x80), "CISO","CISSP, former ITAR compliance officer\nBuilt secure dev programs for 3 DoD primes"),
    ]

    for i, (clr, role, bio) in enumerate(team):
        col = i % 3; row = i // 2
        lx = Inches(0.35 + col * 4.3)
        ty = Inches(1.72 + row * 2.0)
        add_filled_rect(sl, lx, ty, Inches(4.05), Inches(1.8), MID_GRAY)
        # avatar placeholder
        c = sl.shapes.add_shape(9, lx + Inches(0.18), ty + Inches(0.35), Inches(0.7), Inches(0.7))
        c.fill.solid(); c.fill.fore_color.rgb = clr; c.line.color.rgb = clr
        add_text_box(sl, role, lx + Inches(1.05), ty + Inches(0.12), Inches(2.85), Inches(0.42),
                     font_size=12, bold=True, color=WHITE)
        add_text_box(sl, bio,  lx + Inches(1.05), ty + Inches(0.52), Inches(2.85), Inches(1.1),
                     font_size=11, color=LIGHT_GRAY)

    # governance
    add_filled_rect(sl, Inches(0.35), Inches(5.95), Inches(12.6), Inches(1.32), MID_GRAY)
    add_text_box(sl, "GOVERNANCE & OPEN-SOURCE MODEL", Inches(0.5), Inches(6.0), Inches(12), Inches(0.35),
                 font_size=12, bold=True, color=CYAN)
    add_text_box(sl,
                 "Core platform: Proprietary commercial licence  |  "
                 "Community plugins: Apache 2.0 open source  |  "
                 "Standards-body engagement: ISO TC22 · IEC TC65 · SAE  |  "
                 "Advisory board: 8 domain experts from Automotive, Aerospace, Medical & Industrial sectors",
                 Inches(0.5), Inches(6.38), Inches(12.2), Inches(0.75),
                 font_size=11.5, color=LIGHT_GRAY)

    bottom_bar(sl)
    slide_number(sl, 15, TOTAL)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 16 — INVESTMENT & PRICING
# ══════════════════════════════════════════════════════════════════════════════
def slide_pricing():
    sl = prs.slides.add_slide(BLANK_LAYOUT)
    full_bg(sl, DARK_GRAY)
    accent_bar(sl, GREEN)
    section_tag(sl, "Investment & Pricing", GREEN)
    add_text_box(sl, "Engagement Model & ROI Framework", Inches(0.5), Inches(0.9), Inches(12), Inches(0.55),
                 font_size=28, bold=True, color=WHITE)
    divider(sl, Inches(1.52), color=GREEN)

    tiers = [
        (BLUE,   "STARTER",    "$8,500 / mo",  "Up to 15 users · 3 AI agents\n2 integrations · Community support\nSaaS only · Standard SLA"),
        (GREEN,  "PROFESSIONAL","$24,000 / mo", "Up to 60 users · 6 AI agents\nFull integration suite · Priority support\nSaaS or on-prem · 99.9% SLA\nISO 26262 / IEC 61508 CERT-AI"),
        (PURPLE, "ENTERPRISE",  "Custom",       "Unlimited users & agents\nDedicated AI model fine-tuning\nAir-gapped / ITAR deployment\n24/7 dedicated CSM\nCo-development & white-label options"),
    ]

    for i, (clr, name, price, features) in enumerate(tiers):
        lx = Inches(0.35 + i * 4.3)
        # highlight middle tier
        h = Inches(4.4) if i == 1 else Inches(4.0)
        ty = Inches(1.72) if i == 1 else Inches(1.92)
        add_filled_rect(sl, lx, ty, Inches(4.05), h, MID_GRAY)
        add_filled_rect(sl, lx, ty, Inches(4.05), Inches(0.05), clr)
        if i == 1:
            add_filled_rect(sl, lx, ty, Inches(4.05), Inches(0.35), clr)
            add_text_box(sl, "★ RECOMMENDED", lx, ty, Inches(4.05), Inches(0.35),
                         font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text_box(sl, name,  lx + Inches(0.15), ty + Inches(0.45), Inches(3.75), Inches(0.45),
                     font_size=20, bold=True, color=clr)
        add_text_box(sl, price, lx + Inches(0.15), ty + Inches(0.9),  Inches(3.75), Inches(0.55),
                     font_size=26, bold=True, color=WHITE)
        add_text_box(sl, features, lx + Inches(0.15), ty + Inches(1.55), Inches(3.75), Inches(2.4),
                     font_size=12, color=LIGHT_GRAY)

    # ROI
    add_filled_rect(sl, Inches(0.35), Inches(6.35), Inches(12.6), Inches(0.9), NAVY)
    add_text_box(sl, "ROI FRAMEWORK:",
                 Inches(0.5), Inches(6.4), Inches(1.8), Inches(0.35),
                 font_size=12, bold=True, color=CYAN)
    add_text_box(sl,
                 "Typical customer payback in < 4 months. "
                 "At $24K/mo Professional tier: avg. value capture $290K+/mo "
                 "from saved engineering hours, defect avoidance, and certification acceleration.",
                 Inches(2.4), Inches(6.4), Inches(10.3), Inches(0.75),
                 font_size=12, color=LIGHT_GRAY)

    bottom_bar(sl)
    slide_number(sl, 16, TOTAL)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 17 — CALL TO ACTION
# ══════════════════════════════════════════════════════════════════════════════
def slide_cta():
    sl = prs.slides.add_slide(BLANK_LAYOUT)
    full_bg(sl, DARK_GRAY)
    accent_bar(sl, CYAN)
    section_tag(sl, "Next Steps", CYAN)
    add_text_box(sl, "Ready to Transform Your Embedded Lifecycle?", Inches(0.5), Inches(0.9), Inches(12), Inches(0.55),
                 font_size=26, bold=True, color=WHITE)
    divider(sl, Inches(1.52), color=CYAN)

    steps = [
        (BLUE,   "1", "Schedule a\nTechnical Demo",   "Live walkthrough of all 6 AI agents\nagainst your actual tool chain.",
                 "Book at aeplo.ai/demo"),
        (GREEN,  "2", "Join the\n8-Week Pilot",       "Deploy on a real active project.\nFull support. No commitment.",
                 "pilot@aeplo.ai"),
        (PURPLE, "3", "Proof-of-Value\nAssessment",   "Free 3-day PoV: we analyse your\nlifecycle and quantify your ROI.",
                 "Starts within 5 business days"),
        (ORANGE, "4", "Enterprise\nOnboarding",       "Full deployment with integration,\ntraining and success management.",
                 "enterprise@aeplo.ai"),
    ]

    for i, (clr, num, title, desc, cta) in enumerate(steps):
        lx = Inches(0.35 + i * 3.2)
        add_filled_rect(sl, lx, Inches(1.72), Inches(3.0), Inches(4.65), MID_GRAY)
        add_filled_rect(sl, lx, Inches(1.72), Inches(3.0), Inches(0.06), clr)

        c = sl.shapes.add_shape(9, lx + Inches(1.15), Inches(1.9), Inches(0.7), Inches(0.7))
        c.fill.solid(); c.fill.fore_color.rgb = clr; c.line.color.rgb = clr
        add_text_box(sl, num, lx + Inches(1.15), Inches(1.9), Inches(0.7), Inches(0.7),
                     font_size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        add_text_box(sl, title, lx + Inches(0.15), Inches(2.75), Inches(2.7), Inches(0.65),
                     font_size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text_box(sl, desc,  lx + Inches(0.15), Inches(3.48), Inches(2.7), Inches(0.9),
                     font_size=12, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)
        add_filled_rect(sl, lx + Inches(0.15), Inches(4.5), Inches(2.7), Inches(0.5), NAVY)
        add_text_box(sl, cta,   lx + Inches(0.15), Inches(4.5), Inches(2.7), Inches(0.5),
                     font_size=11, bold=True, color=CYAN, align=PP_ALIGN.CENTER)

    # Bottom contact strip
    add_filled_rect(sl, Inches(0.35), Inches(6.55), Inches(12.6), Inches(0.72), NAVY)
    add_text_box(sl,
                 "aeplo.ai   |   hello@aeplo.ai   |   +1 (415) 000-0000   |   "
                 "LinkedIn: /company/aeplo   |   GitHub: github.com/aeplo",
                 Inches(0.5), Inches(6.65), Inches(12.2), Inches(0.52),
                 font_size=13, color=WHITE, align=PP_ALIGN.CENTER)

    bottom_bar(sl)
    slide_number(sl, 17, TOTAL)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 18 — THANK YOU / CLOSING
# ══════════════════════════════════════════════════════════════════════════════
def slide_closing():
    sl = prs.slides.add_slide(BLANK_LAYOUT)
    full_bg(sl, NAVY)
    add_filled_rect(sl, 0, 0, Inches(0.1), SLIDE_H, CYAN)
    add_filled_rect(sl, 0, SLIDE_H - Inches(0.1), SLIDE_W, Inches(0.1), BLUE)

    # Large geometric accent
    add_filled_rect(sl, SLIDE_W * 0.62, 0, SLIDE_W * 0.38, SLIDE_H, DARK_GRAY)
    add_filled_rect(sl, SLIDE_W * 0.62, Inches(2.5), SLIDE_W * 0.38, Inches(0.006), CYAN)

    for cx, cy, sz, clr in [
        (SLIDE_W * 0.82, Inches(2.2), Inches(3.0), MID_GRAY),
        (SLIDE_W * 0.95, Inches(5.5), Inches(1.8), RGBColor(0x0A,0x30,0x60)),
    ]:
        c = sl.shapes.add_shape(9, cx - sz/2, cy - sz/2, sz, sz)
        c.fill.solid(); c.fill.fore_color.rgb = clr; c.line.color.rgb = clr

    add_text_box(sl, "THANK YOU", Inches(0.45), Inches(1.8), Inches(7.5), Inches(1.2),
                 font_size=56, bold=True, color=WHITE)
    add_text_box(sl, "Build better. Ship faster. Certify smarter.",
                 Inches(0.45), Inches(3.1), Inches(7.5), Inches(0.6),
                 font_size=20, italic=True, color=CYAN)

    add_text_box(sl, "AEPLO V3 — AI-Driven Embedded Product Lifecycle Orchestrator",
                 Inches(0.45), Inches(4.0), Inches(7.5), Inches(0.45),
                 font_size=14, color=LIGHT_GRAY)

    contacts = [
        ("Website",  "aeplo.ai"),
        ("Email",    "hello@aeplo.ai"),
        ("LinkedIn", "/company/aeplo"),
        ("GitHub",   "github.com/aeplo"),
    ]
    for i, (lbl, val) in enumerate(contacts):
        ty = Inches(4.65 + i * 0.5)
        add_text_box(sl, lbl + ":", Inches(0.45), ty, Inches(1.2), Inches(0.38),
                     font_size=13, bold=True, color=BLUE)
        add_text_box(sl, val,       Inches(1.7),  ty, Inches(4.0), Inches(0.38),
                     font_size=13, color=WHITE)

    add_text_box(sl, "© 2026 AEPLO Technologies Inc. All rights reserved. Confidential.",
                 Inches(0.45), Inches(6.95), Inches(7.5), Inches(0.38),
                 font_size=10, color=RGBColor(0x60,0x70,0x90))

    slide_number(sl, 18, TOTAL)


# ══════════════════════════════════════════════════════════════════════════════
# BUILD ALL SLIDES
# ══════════════════════════════════════════════════════════════════════════════
slide_cover()
slide_toc()
slide_exec_summary()
slide_problem()
slide_solution()
slide_architecture()
slide_ai_modules()
slide_lifecycle()
slide_features()
slide_integrations()
slide_security()
slide_metrics()
slide_case_studies()
slide_roadmap()
slide_team()
slide_pricing()
slide_cta()
slide_closing()

OUTPUT = "/workspace/AI-Driven Embedded Product Lifecycle OrchestratorV3.pptx"
prs.save(OUTPUT)
print(f"Saved: {OUTPUT}  ({len(prs.slides)} slides)")
