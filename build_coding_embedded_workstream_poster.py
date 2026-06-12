"""
One-page poster candidate for the Coding Embedded Workstream common idea.
Generates: Coding Embedded Workstream - Common Idea Poster.pptx
"""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


# Slide size: widescreen 16:9
W = Inches(13.333)
H = Inches(7.5)

prs = Presentation()
prs.slide_width = W
prs.slide_height = H
BLANK = prs.slide_layouts[6]


# Palette aligned with the AEPLO deck.
DARK = RGBColor(0x02, 0x0A, 0x1A)
NAVY = RGBColor(0x05, 0x14, 0x2E)
PANEL = RGBColor(0x0B, 0x1F, 0x3A)
CARD = RGBColor(0x0F, 0x28, 0x45)
BLUE = RGBColor(0x00, 0x72, 0xC6)
CYAN = RGBColor(0x00, 0xC8, 0xF0)
TEAL = RGBColor(0x00, 0xA8, 0x9A)
GREEN = RGBColor(0x00, 0xD4, 0x8A)
ORANGE = RGBColor(0xFF, 0x8C, 0x00)
PURPLE = RGBColor(0x8B, 0x5C, 0xF6)
RED = RGBColor(0xFF, 0x3A, 0x3A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xB8, 0xD0, 0xEC)
DIM = RGBColor(0x60, 0x80, 0xA8)


def rect(slide, l, t, w, h, fill, line_color=None):
    shape = slide.shapes.add_shape(1, l, t, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    return shape


def roundrect(slide, l, t, w, h, fill, line_color=None):
    shape = slide.shapes.add_shape(5, l, t, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    return shape


def oval(slide, l, t, w, h, fill, line_color=None):
    shape = slide.shapes.add_shape(9, l, t, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    return shape


def txt(slide, text, l, t, w, h, size=14, bold=False, italic=False,
        color=WHITE, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(l, t, w, h)
    box.word_wrap = True
    frame = box.text_frame
    frame.word_wrap = True
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = "Segoe UI"
    return box


def bullet_list(slide, items, l, t, w, h, size=10.5, color=LIGHT):
    box = slide.shapes.add_textbox(l, t, w, h)
    frame = box.text_frame
    frame.word_wrap = True
    for idx, item in enumerate(items):
        paragraph = frame.paragraphs[0] if idx == 0 else frame.add_paragraph()
        paragraph.level = 0
        run = paragraph.add_run()
        run.text = "- " + item
        run.font.size = Pt(size)
        run.font.color.rgb = color
        run.font.name = "Segoe UI"
    return box


def arrow(slide, x, y, w, color=CYAN):
    shape = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x, y, w, Inches(0.2))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def stat_card(slide, l, t, w, value, label, color):
    roundrect(slide, l, t, w, Inches(0.82), PANEL)
    rect(slide, l, t, Inches(0.06), Inches(0.82), color)
    txt(slide, value, l + Inches(0.14), t + Inches(0.08), w - Inches(0.28), Inches(0.28),
        size=19, bold=True, color=color, align=PP_ALIGN.CENTER)
    txt(slide, label, l + Inches(0.14), t + Inches(0.42), w - Inches(0.28), Inches(0.28),
        size=8.8, color=LIGHT, align=PP_ALIGN.CENTER)


def build_poster():
    slide = prs.slides.add_slide(BLANK)
    rect(slide, 0, 0, W, H, DARK)
    rect(slide, 0, 0, W, Inches(0.08), CYAN)
    rect(slide, 0, H - Inches(0.08), W, Inches(0.08), GREEN)

    # Template-inspired dark tech backdrop.
    rect(slide, 0, Inches(1.18), W, Inches(0.04), BLUE)
    rect(slide, 0, Inches(1.75), W, Inches(0.02), RGBColor(0x11, 0x3A, 0x66))
    for i in range(19):
        rect(slide, Inches(8.7), Inches(0.15 + i * 0.32),
             Inches(4.35), Inches(0.006), RGBColor(0x06, 0x1C, 0x38))
    for i in range(12):
        rect(slide, Inches(8.72 + i * 0.36), Inches(0.15),
             Inches(0.006), Inches(6.85), RGBColor(0x06, 0x1C, 0x38))
    oval(slide, Inches(11.72), Inches(0.14), Inches(0.5), Inches(0.5), CYAN)
    oval(slide, Inches(11.86), Inches(0.28), Inches(0.22), Inches(0.22), WHITE)

    # Header echoes the supplied template's strong impact statement.
    txt(slide, "CONNECTED EMBEDDED DATA. INTELLIGENT PDLC ACTION.",
        Inches(0.33), Inches(0.17), Inches(8.7), Inches(0.36),
        size=23, bold=True, color=WHITE)
    txt(slide, "REAL NPI IMPACT.",
        Inches(0.33), Inches(0.55), Inches(5.0), Inches(0.42),
        size=25, bold=True, color=CYAN)
    txt(slide,
        "AEPLO unifies Jira, Confluence, GitHub, Copilot, ROVO Studio and JTAG/SWD evidence "
        "to optimise embedded PDLC decisions across PG1, PG3 and PG5.",
        Inches(0.35), Inches(0.98), Inches(7.9), Inches(0.25),
        size=9.8, color=LIGHT)

    roundrect(slide, Inches(8.85), Inches(0.12), Inches(4.15), Inches(0.98), PANEL)
    txt(slide, "ONE AEPLO PLATFORM.", Inches(9.2), Inches(0.27),
        Inches(2.5), Inches(0.25), size=13, bold=True, color=GREEN)
    txt(slide, "EMBEDDED. PDLC. AI OPTIMISATION.", Inches(9.2), Inches(0.55),
        Inches(3.35), Inches(0.24), size=12, bold=True, color=CYAN)
    txt(slide, "Honeywell BA Buildathon | 17 Jun 2026",
        Inches(9.2), Inches(0.83), Inches(3.3), Inches(0.18),
        size=8.3, color=LIGHT)

    # Tool ribbon.
    tools = [
        ("ROVO Studio", PURPLE),
        ("Copilot", GREEN),
        ("Atlassian MCP", ORANGE),
        ("GitHub", WHITE),
        ("Jira", BLUE),
        ("Confluence", CYAN),
    ]
    for i, (name, color) in enumerate(tools):
        x = Inches(0.28 + i * 2.06)
        roundrect(slide, x, Inches(1.34), Inches(1.82), Inches(0.34), NAVY)
        rect(slide, x, Inches(1.34), Inches(0.04), Inches(0.34), color)
        txt(slide, name, x + Inches(0.12), Inches(1.42),
            Inches(1.55), Inches(0.12), size=8.4, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Main connected platform columns.
    columns = [
        (
            GREEN,
            "1. EMBEDDED DATA SOURCES",
            "Discovery + engineering context",
            [
                ("Jira", "NPI epics, defects, decisions"),
                ("Confluence", "PRD, architecture, lessons"),
                ("GitHub", "Code, PRs, branches, reviews"),
                ("JTAG/SWD", "MCU registers, memory, traces"),
            ],
        ),
        (
            BLUE,
            "2. UNIFIED PDLC FOUNDATION",
            "Single source of truth",
            [
                ("Traceability", "Requirements -> code -> test"),
                ("Knowledge Graph", "Board, firmware, defect links"),
                ("Evidence Lake", "Logs, dumps, probe sessions"),
                ("NPI Context", "PG1 / PG3 / PG5 gates"),
            ],
        ),
        (
            ORANGE,
            "3. AEPLO ORCHESTRATION",
            "Work that flows",
            [
                ("Atlassian MCP", "Jira + Confluence actions"),
                ("JTAG Accelerator", "Hardware-backed RCA"),
                ("Release Gate", "READY / NOT_READY verdict"),
                ("GitHub Automation", "Branch, PR, review evidence"),
            ],
        ),
        (
            PURPLE,
            "4. AI-POWERED EXPERIENCES",
            "Actionable intelligence",
            [
                ("ROVO Studio", "NPI demo cockpit"),
                ("Copilot", "Code fix and explanation"),
                ("AI Agents", "Fault, variable, peripheral RCA"),
                ("AI Optimisation", "Risk, tests, fixes, readiness"),
            ],
        ),
        (
            TEAL,
            "5. NPI GATE OUTCOMES",
            "Demo aligned to phases",
            [
                ("PG1 Discovery", "Risk-backed scope and plan"),
                ("PG3 Development", "RCA, patch, PR, evidence"),
                ("PG5 Validation", "Launch readiness action list"),
                ("PDLC Impact", "Faster debug, better decisions"),
            ],
        ),
    ]

    x_positions = [0.24, 2.27, 4.5, 6.73, 8.96]
    widths = [1.82, 2.02, 2.02, 2.02, 1.92]
    for idx, (color, title, subtitle, items) in enumerate(columns):
        x = Inches(x_positions[idx])
        w = Inches(widths[idx])
        y = Inches(1.92)
        roundrect(slide, x, y, w, Inches(3.65), CARD)
        rect(slide, x, y, w, Inches(0.06), color)
        txt(slide, title, x + Inches(0.08), y + Inches(0.12),
            w - Inches(0.16), Inches(0.18), size=7.4, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txt(slide, subtitle, x + Inches(0.12), y + Inches(0.42),
            w - Inches(0.24), Inches(0.18), size=7.3, color=color, align=PP_ALIGN.CENTER)
        for j, (label, body) in enumerate(items):
            iy = y + Inches(0.82 + j * 0.66)
            roundrect(slide, x + Inches(0.12), iy, w - Inches(0.24), Inches(0.48), PANEL)
            oval(slide, x + Inches(0.22), iy + Inches(0.15), Inches(0.18), Inches(0.18), color)
            txt(slide, label, x + Inches(0.46), iy + Inches(0.08),
                w - Inches(0.62), Inches(0.13), size=7.4, bold=True, color=WHITE)
            txt(slide, body, x + Inches(0.46), iy + Inches(0.25),
                w - Inches(0.62), Inches(0.13), size=5.8, color=LIGHT)
        if idx < len(columns) - 1:
            arrow(slide, x + w + Inches(0.03), y + Inches(1.75), Inches(0.3), CYAN)

    # Right value panel from the supplied template.
    roundrect(slide, Inches(11.1), Inches(1.72), Inches(2.02), Inches(3.92), RGBColor(0x12, 0x22, 0x40))
    rect(slide, Inches(11.1), Inches(1.72), Inches(2.02), Inches(0.05), CYAN)
    txt(slide, "BUSINESS VALUE", Inches(11.22), Inches(1.88),
        Inches(1.72), Inches(0.24), size=11.2, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    txt(slide, "THAT EMBEDDED TEAMS CARE ABOUT", Inches(11.24), Inches(2.14),
        Inches(1.68), Inches(0.18), size=6.7, bold=True, color=LIGHT, align=PP_ALIGN.CENTER)
    values = [
        (GREEN, "ACCELERATE NPI", "PG1 risk to PG5 launch evidence"),
        (ORANGE, "REDUCE DEBUG CYCLE", "JTAG-backed RCA and patches"),
        (PURPLE, "IMPROVE QUALITY", "AI-focused tests and gates"),
        (CYAN, "STRENGTHEN GOVERNANCE", "Traceability and proof in Jira"),
    ]
    for i, (color, heading, body) in enumerate(values):
        y = Inches(2.55 + i * 0.68)
        roundrect(slide, Inches(11.25), y, Inches(1.72), Inches(0.5), PANEL)
        oval(slide, Inches(11.38), y + Inches(0.13), Inches(0.24), Inches(0.24), color)
        txt(slide, heading, Inches(11.72), y + Inches(0.08),
            Inches(1.08), Inches(0.11), size=6.5, bold=True, color=color)
        txt(slide, body, Inches(11.72), y + Inches(0.24),
            Inches(1.08), Inches(0.13), size=5.3, color=LIGHT)

    # Bottom band: from data to impact.
    roundrect(slide, Inches(0.28), Inches(5.86), Inches(7.25), Inches(0.78), NAVY)
    txt(slide, "FROM NPI DATA TO EMBEDDED IMPACT", Inches(0.52), Inches(5.96),
        Inches(2.8), Inches(0.18), size=9.5, bold=True, color=WHITE)
    flow = [
        ("Collect", "Jira + Confluence"),
        ("Unify", "AEPLO PDLC graph"),
        ("Orchestrate", "Atlassian MCP"),
        ("Diagnose", "JTAG Accelerator"),
        ("Optimise", "ROVO + Copilot"),
        ("Impact", "PG5 launch"),
    ]
    for i, (label, body) in enumerate(flow):
        cx = Inches(0.75 + i * 1.1)
        oval(slide, cx, Inches(6.22), Inches(0.34), Inches(0.34),
             [BLUE, GREEN, ORANGE, PURPLE, CYAN, GREEN][i])
        txt(slide, str(i + 1), cx, Inches(6.29), Inches(0.34), Inches(0.08),
            size=6.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txt(slide, label, cx - Inches(0.22), Inches(6.61),
            Inches(0.78), Inches(0.11), size=6.3, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txt(slide, body, cx - Inches(0.3), Inches(6.74),
            Inches(0.95), Inches(0.11), size=4.9, color=LIGHT, align=PP_ALIGN.CENTER)
        if i < len(flow) - 1:
            arrow(slide, cx + Inches(0.44), Inches(6.28), Inches(0.35), color=LIGHT)

    # NPI phase banner.
    roundrect(slide, Inches(7.78), Inches(5.86), Inches(5.18), Inches(0.78), PANEL)
    txt(slide, "DEMO PHASES", Inches(7.98), Inches(5.98),
        Inches(1.1), Inches(0.16), size=8.8, bold=True, color=CYAN)
    phases = [("PG1", "Discovery", BLUE), ("PG3", "Development", ORANGE), ("PG5", "Validation & Launch Prep", GREEN)]
    for i, (gate, label, color) in enumerate(phases):
        x = Inches(8.0 + i * 1.62)
        roundrect(slide, x, Inches(6.25), Inches(1.42), Inches(0.28), color)
        txt(slide, f"{gate}  {label}", x, Inches(6.31),
            Inches(1.42), Inches(0.08), size=6.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Footer strip.
    rect(slide, 0, Inches(7.02), W, Inches(0.4), RGBColor(0xD7, 0xF7, 0xFF))
    txt(slide, "Honeywell BA Buildathon. Built for Embedded PDLC Impact.",
        Inches(0.28), Inches(7.12), Inches(3.3), Inches(0.12),
        size=8.5, bold=True, color=NAVY)
    txt(slide,
        "Presenter: Shyam Shrivastava, Project Lead | Team: Atharva, Manas, Sheraaz",
        Inches(3.75), Inches(7.09), Inches(9.15), Inches(0.11),
        size=6.8, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
    txt(slide,
        "ROVO Studio | Copilot | Atlassian MCP | GitHub | Jira | Confluence | AEPLO JTAG Accelerator",
        Inches(3.75), Inches(7.25), Inches(9.15), Inches(0.10),
        size=6.2, color=NAVY, align=PP_ALIGN.RIGHT)


build_poster()

OUTPUT = "/workspace/Coding Embedded Workstream - Common Idea Poster.pptx"
prs.save(OUTPUT)
print(f"Saved: {OUTPUT} ({len(prs.slides)} slide)")
