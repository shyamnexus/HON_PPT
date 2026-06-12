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
    rect(slide, 0, H - Inches(0.08), W, Inches(0.08), BLUE)

    # Decorative right-side circuit grid.
    for i in range(13):
        rect(slide, Inches(9.2), Inches(0.35 + i * 0.48),
             Inches(3.9), Inches(0.01), RGBColor(0x06, 0x1C, 0x38))
    for i in range(8):
        rect(slide, Inches(9.25 + i * 0.52), Inches(0.35),
             Inches(0.01), Inches(6.35), RGBColor(0x06, 0x1C, 0x38))
    oval(slide, Inches(10.3), Inches(1.0), Inches(2.25), Inches(2.25),
         RGBColor(0x06, 0x1C, 0x38))
    oval(slide, Inches(10.65), Inches(1.35), Inches(1.55), Inches(1.55), PANEL)
    oval(slide, Inches(11.02), Inches(1.72), Inches(0.82), Inches(0.82), NAVY)

    # Header.
    txt(slide, "CODING EMBEDDED WORKSTREAM", Inches(0.35), Inches(0.28),
        Inches(5.7), Inches(0.35), size=13, bold=True, color=CYAN)
    txt(slide, "Common Idea: Agentic Firmware Debug-to-Fix Copilot",
        Inches(0.35), Inches(0.68), Inches(8.7), Inches(0.58),
        size=28, bold=True, color=WHITE)
    txt(slide,
        "A unified AEPLO + JTAG/SWD workflow that turns embedded coding defects "
        "into traceable hardware-backed root cause, patch, PR and release evidence.",
        Inches(0.38), Inches(1.27), Inches(8.55), Inches(0.48),
        size=12.5, italic=True, color=LIGHT)

    roundrect(slide, Inches(9.55), Inches(0.58), Inches(3.35), Inches(0.72), BLUE)
    txt(slide, "POSTER CANDIDATE", Inches(9.55), Inches(0.66),
        Inches(3.35), Inches(0.24), size=12, bold=True,
        color=WHITE, align=PP_ALIGN.CENTER)
    txt(slide, "BuildAThon | Embedded Coding", Inches(9.55), Inches(0.94),
        Inches(3.35), Inches(0.22), size=9.5, color=CYAN,
        align=PP_ALIGN.CENTER)

    # Problem and idea blocks.
    roundrect(slide, Inches(0.35), Inches(1.95), Inches(3.9), Inches(1.55), CARD)
    rect(slide, Inches(0.35), Inches(1.95), Inches(0.06), Inches(1.55), RED)
    txt(slide, "WHY IT MATTERS", Inches(0.55), Inches(2.08),
        Inches(3.45), Inches(0.28), size=12, bold=True, color=RED)
    bullet_list(slide, [
        "MCU bugs require manual JTAG setup, register decoding and memory inspection.",
        "Engineers switch between IDE, probe tools, datasheets, Jira and Git.",
        "HardFault knowledge is tribal, so fixes take days and release risk grows.",
    ], Inches(0.55), Inches(2.43), Inches(3.45), Inches(0.86), size=8.8)

    roundrect(slide, Inches(4.55), Inches(1.95), Inches(4.25), Inches(1.55), CARD)
    rect(slide, Inches(4.55), Inches(1.95), Inches(0.06), Inches(1.55), GREEN)
    txt(slide, "COMMON IDEA", Inches(4.75), Inches(2.08),
        Inches(3.8), Inches(0.28), size=12, bold=True, color=GREEN)
    bullet_list(slide, [
        "Attach multi-agent AI to the coding workflow and target hardware.",
        "Use JTAG/SWD evidence as first-class lifecycle data in AEPLO.",
        "Move from bug ticket to diagnosis, controlled fix, PR and release gate.",
    ], Inches(4.75), Inches(2.43), Inches(3.75), Inches(0.86), size=8.8)

    # Workflow centerline.
    txt(slide, "WORKFLOW", Inches(0.35), Inches(3.78),
        Inches(2.0), Inches(0.3), size=12, bold=True, color=CYAN)
    workflow = [
        (BLUE, "1", "Jira Bug\nTicket"),
        (ORANGE, "2", "JTAG\nOrchestrator"),
        (PURPLE, "3", "Debug\nAgents"),
        (TEAL, "4", "RCA +\nEvidence"),
        (GREEN, "5", "Patch +\nPR"),
        (CYAN, "6", "Release\nGate"),
    ]
    for i, (color, num, label) in enumerate(workflow):
        x = Inches(0.55 + i * 2.05)
        oval(slide, x, Inches(4.16), Inches(0.54), Inches(0.54), color)
        txt(slide, num, x, Inches(4.19), Inches(0.54), Inches(0.22),
            size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        roundrect(slide, x - Inches(0.32), Inches(4.82), Inches(1.18), Inches(0.74), PANEL)
        txt(slide, label, x - Inches(0.24), Inches(4.92), Inches(1.02), Inches(0.38),
            size=8.8, bold=True, color=LIGHT, align=PP_ALIGN.CENTER)
        if i < len(workflow) - 1:
            arrow(slide, x + Inches(0.75), Inches(4.33), Inches(1.05), color=CYAN)

    # Agent stack.
    roundrect(slide, Inches(9.25), Inches(1.65), Inches(3.65), Inches(2.2), CARD)
    rect(slide, Inches(9.25), Inches(1.65), Inches(0.06), Inches(2.2), ORANGE)
    txt(slide, "SPECIALIST AGENTS", Inches(9.45), Inches(1.8),
        Inches(3.2), Inches(0.28), size=12, bold=True, color=ORANGE)
    bullet_list(slide, [
        "Fault Analyzer: HardFault, stack and register RCA",
        "Variable Tracer: watchpoints and value history",
        "Peripheral Inspector: GPIO, timers, buses, clocks",
        "Hex/Memory Parser: dumps, maps and symbol context",
        "Probing Agent: J-Link, target MCU and lab evidence",
    ], Inches(9.45), Inches(2.18), Inches(3.2), Inches(1.3), size=8.5)

    # Demo and outcomes.
    roundrect(slide, Inches(0.35), Inches(5.95), Inches(6.25), Inches(1.05), CARD)
    rect(slide, Inches(0.35), Inches(5.95), Inches(0.06), Inches(1.05), CYAN)
    txt(slide, "DEMO STORY", Inches(0.55), Inches(6.05),
        Inches(1.5), Inches(0.25), size=11.5, bold=True, color=CYAN)
    txt(slide,
        "A Cortex-M HardFault ticket opens in Jira. The copilot attaches through "
        "J-Link, captures fault registers and stack frames, explains root cause, "
        "generates a guarded patch, opens a PR and feeds evidence into the release gate.",
        Inches(0.55), Inches(6.35), Inches(5.8), Inches(0.38),
        size=9.2, color=LIGHT)

    txt(slide, "EXPECTED OUTCOMES", Inches(6.95), Inches(5.9),
        Inches(2.2), Inches(0.25), size=11.5, bold=True, color=GREEN)
    stat_card(slide, Inches(6.95), Inches(6.22), Inches(1.42), "<30m", "ticket to RCA", GREEN)
    stat_card(slide, Inches(8.55), Inches(6.22), Inches(1.42), "1 PR", "controlled fix", BLUE)
    stat_card(slide, Inches(10.15), Inches(6.22), Inches(1.42), "100%", "evidence link", ORANGE)
    stat_card(slide, Inches(11.75), Inches(6.22), Inches(1.15), "Ready", "gate input", CYAN)

    # Footer.
    txt(slide, "Team addition: Sheraaz | Common scheme: AEPLO V3 lifecycle orchestration",
        Inches(0.35), Inches(7.18), Inches(8.4), Inches(0.18),
        size=8.5, color=DIM)
    txt(slide, "One-page candidate for Coding Embedded Workstream",
        Inches(9.15), Inches(7.18), Inches(3.8), Inches(0.18),
        size=8.5, color=DIM, align=PP_ALIGN.RIGHT)


build_poster()

OUTPUT = "/workspace/Coding Embedded Workstream - Common Idea Poster.pptx"
prs.save(OUTPUT)
print(f"Saved: {OUTPUT} ({len(prs.slides)} slide)")
