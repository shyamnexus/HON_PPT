"""
AI-Driven Embedded Product Lifecycle Orchestrator V10
3-minute PowerPoint builder.

This version preserves the V9 problem statement and architecture storyline,
then compresses the remaining narrative around PRD-driven Jira work,
PG3 firmware generation, QA release gating, JTAG evidence, and the PG5 STQC
agent.
"""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


# Palette
NAVY = RGBColor(0x05, 0x14, 0x2E)
DARK_NAVY = RGBColor(0x02, 0x0A, 0x1A)
BLUE = RGBColor(0x00, 0x72, 0xC6)
CYAN = RGBColor(0x00, 0xC8, 0xF0)
TEAL = RGBColor(0x00, 0xA8, 0x9A)
PANEL = RGBColor(0x0B, 0x1F, 0x3A)
CARD = RGBColor(0x0F, 0x28, 0x45)
MID = RGBColor(0x0E, 0x2A, 0x4A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xB8, 0xD0, 0xEC)
DIM = RGBColor(0x60, 0x80, 0xA8)
GREEN = RGBColor(0x00, 0xD4, 0x8A)
ORANGE = RGBColor(0xFF, 0x8C, 0x00)
RED = RGBColor(0xFF, 0x3A, 0x3A)
PURPLE = RGBColor(0x8B, 0x5C, 0xF6)
GOLD = RGBColor(0xF5, 0xC5, 0x18)

W = Inches(10.0)
H = Inches(7.5)
TOTAL = 8

prs = Presentation()
prs.slide_width = W
prs.slide_height = H
BLANK = prs.slide_layouts[6]


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
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.04)
    tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = "Segoe UI"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return box


def bg(slide):
    rect(slide, 0, 0, W, H, DARK_NAVY)


def top_bar(slide, color):
    rect(slide, 0, 0, W, Inches(0.06), color)


def bottom_strip(slide, n):
    rect(slide, 0, H - Inches(0.42), W, Inches(0.42), MID)
    txt(slide, f"{n} / {TOTAL}", W - Inches(0.9), H - Inches(0.36),
        Inches(0.75), Inches(0.25), size=9, color=DIM, align=PP_ALIGN.RIGHT)
    txt(slide, "HONEYWELL", Inches(0.35), H - Inches(0.36),
        Inches(1.5), Inches(0.25), size=9, bold=True, color=DIM)


def section_pill(slide, label, color):
    roundrect(slide, Inches(0.35), Inches(0.45), Inches(2.25), Inches(0.32), color)
    txt(slide, label.upper(), Inches(0.35), Inches(0.45),
        Inches(2.25), Inches(0.32), size=8.5, bold=True,
        color=WHITE, align=PP_ALIGN.CENTER)


def title(slide, text):
    txt(slide, text, Inches(0.35), Inches(0.82), Inches(9.2), Inches(0.58),
        size=27, bold=True, color=WHITE)


def divider(slide, y, color):
    rect(slide, Inches(0.35), y, Inches(9.3), Inches(0.018), color)


def card(slide, x, y, w, h, accent, heading, body, heading_size=13, body_size=10.5):
    roundrect(slide, x, y, w, h, CARD)
    rect(slide, x, y, w, Inches(0.05), accent)
    txt(slide, heading, x + Inches(0.15), y + Inches(0.13),
        w - Inches(0.3), Inches(0.45), size=heading_size, bold=True, color=accent)
    txt(slide, body, x + Inches(0.15), y + Inches(0.6),
        w - Inches(0.3), h - Inches(0.72), size=body_size, color=LIGHT)


def slide_01_cover():
    sl = prs.slides.add_slide(BLANK)
    bg(sl)
    rect(sl, 0, 0, W * 0.58, H, NAVY)
    rect(sl, 0, 0, Inches(0.07), H, CYAN)
    rect(sl, W * 0.58, 0, W * 0.42, H, PANEL)

    # Ring accent.
    cx, cy = W * 0.79, H * 0.47
    for r, color in [(Inches(3.1), MID), (Inches(2.35), CARD), (Inches(1.45), NAVY)]:
        oval(sl, cx - r, cy - r, r * 2, r * 2, color)
    for i in range(12):
        rect(sl, W * 0.61 + Inches(i * 0.27), Inches(1.3 + (i % 5) * 0.55),
             Inches(0.08), Inches(0.08), CYAN if i % 3 == 0 else BLUE)

    txt(sl, "AI-DRIVEN EMBEDDED", Inches(0.32), Inches(1.1),
        Inches(6.0), Inches(0.72), size=34, bold=True)
    txt(sl, "PRODUCT LIFECYCLE", Inches(0.32), Inches(1.82),
        Inches(6.0), Inches(0.72), size=34, bold=True)
    txt(sl, "ORCHESTRATOR", Inches(0.32), Inches(2.54),
        Inches(6.0), Inches(0.72), size=34, bold=True, color=CYAN)
    roundrect(sl, Inches(0.32), Inches(3.35), Inches(1.25), Inches(0.34), BLUE)
    txt(sl, "VERSION 10", Inches(0.32), Inches(3.35),
        Inches(1.25), Inches(0.34), size=11, bold=True, align=PP_ALIGN.CENTER)
    txt(sl, "3-minute executive story: PRD -> firmware -> JTAG evidence -> QA READY -> STQC submission",
        Inches(0.32), Inches(4.08), Inches(5.8), Inches(0.68),
        size=13.5, italic=True, color=LIGHT)
    txt(sl, "Atharva | Manas | Sheraaz | Shyam",
        Inches(0.32), Inches(5.35), Inches(4.8), Inches(0.3), size=10, color=DIM)
    bottom_strip(sl, 1)


def slide_02_problem():
    sl = prs.slides.add_slide(BLANK)
    bg(sl)
    rect(sl, 0, 0, W, Inches(1.45), NAVY)
    top_bar(sl, ORANGE)
    txt(sl, "THE PROBLEM", Inches(0.35), Inches(0.15), Inches(9.0), Inches(0.35),
        size=10, bold=True, color=ORANGE)
    txt(sl, "Embedded PDLC Data Is Fragmented", Inches(0.35), Inches(0.5),
        Inches(9.3), Inches(0.75), size=30, bold=True)
    divider(sl, Inches(1.38), ORANGE)
    txt(sl,
        "NPI gates rely on disconnected requirements, Jira tickets, Confluence knowledge, GitHub code, "
        "JTAG/SWD evidence and QA results - slowing decisions from PG1 to PG5.",
        Inches(0.35), Inches(1.52), Inches(9.35), Inches(0.58),
        size=13, italic=True, color=LIGHT)

    silos = [
        (ORANGE, "PG1", "Discovery\nBlind Spots", "Risks and debug hotspots are not visible early"),
        (RED, "BUG", "Late Defect\nDiscovery", "MCU faults surface after coding and integration"),
        (PURPLE, "TOOL", "Tool Chain\nFragmentation", "ROVO, Copilot, Jira, GitHub and JTAG need orchestration"),
        (BLUE, "TRACE", "PDLC\nTrace Gaps", "Requirements, code, tests and evidence drift apart"),
        (TEAL, "JTAG", "Embedded\nHardware Gap", "JTAG/SWD evidence stays outside lifecycle decisions"),
        (GOLD, "PG5", "Launch Gate\nUncertainty", "Readiness depends on manual proof and subjective reviews"),
    ]
    for i, (clr, icon, head, body) in enumerate(silos):
        col, row = i % 3, i // 3
        x = Inches(0.35 + col * 3.15)
        y = Inches(2.25 + row * 2.35)
        roundrect(sl, x, y, Inches(2.95), Inches(2.12), CARD)
        rect(sl, x, y, Inches(2.95), Inches(0.05), clr)
        roundrect(sl, x + Inches(0.17), y + Inches(0.18), Inches(0.56), Inches(0.5), clr)
        txt(sl, icon, x + Inches(0.17), y + Inches(0.18), Inches(0.56), Inches(0.5),
            size=10, bold=True, align=PP_ALIGN.CENTER)
        txt(sl, head, x + Inches(0.85), y + Inches(0.16), Inches(1.9), Inches(0.58),
            size=12.5, bold=True)
        txt(sl, body, x + Inches(0.18), y + Inches(0.84), Inches(2.58), Inches(1.0),
            size=11, color=LIGHT)
    bottom_strip(sl, 2)


def slide_03_architecture():
    sl = prs.slides.add_slide(BLANK)
    bg(sl)
    top_bar(sl, BLUE)
    section_pill(sl, "Architecture", BLUE)
    title(sl, "AEPLO Platform Architecture")
    divider(sl, Inches(1.45), BLUE)
    txt(sl,
        "Connects Jira, Confluence, GitHub, Copilot, ROVO Studio, Atlassian MCP and JTAG/SWD "
        "evidence into one embedded PDLC optimization platform.",
        Inches(0.35), Inches(1.55), Inches(9.35), Inches(0.46),
        size=12, italic=True, color=LIGHT)

    layers = [
        (ORANGE, "CONNECTED INPUTS",
         "Jira | Confluence | GitHub | Copilot | ROVO Studio | JTAG/SWD | QA Results"),
        (BLUE, "DATA INTEGRATION LAYER",
         "Unified Knowledge Graph | Semantic Traceability Engine | Version History"),
        (CYAN, "AEPLO ORCHESTRATION CORE",
         "Atlassian MCP | JTAG Accelerator | Fault Analyzer | Release Agent | QA Agent"),
        (PURPLE, "AUTOMATION & GATING",
         "CI/CD Hooks | Controlled Auto-Fix | Auto-PR | Release Gate | Remediation Actions"),
        (GREEN, "NPI GATE OUTCOMES",
         "PG1 Risk Plan | PG3 RCA + PR | PG5 READY / NOT_READY | Audit Evidence"),
    ]
    for i, (clr, label, detail) in enumerate(layers):
        y = Inches(2.15 + i * 0.96)
        roundrect(sl, Inches(0.35), y, Inches(2.05), Inches(0.82), clr)
        txt(sl, label, Inches(0.35), y, Inches(2.05), Inches(0.82),
            size=9.5, bold=True, align=PP_ALIGN.CENTER)
        roundrect(sl, Inches(2.55), y, Inches(7.1), Inches(0.82), CARD)
        rect(sl, Inches(2.55), y, Inches(0.05), Inches(0.82), clr)
        txt(sl, detail, Inches(2.75), y + Inches(0.19), Inches(6.75), Inches(0.44),
            size=11.5, color=LIGHT)
        if i < len(layers) - 1:
            rect(sl, Inches(1.35), y + Inches(0.82), Inches(0.05), Inches(0.14), clr)
            rect(sl, Inches(5.1), y + Inches(0.82), Inches(0.05), Inches(0.14), clr)
    bottom_strip(sl, 3)


def slide_04_prd_to_backlog():
    sl = prs.slides.add_slide(BLANK)
    bg(sl)
    top_bar(sl, TEAL)
    section_pill(sl, "PRD to Backlog", TEAL)
    title(sl, "Confluence PRD Becomes Pickable Jira Work")
    divider(sl, Inches(1.45), TEAL)
    txt(sl,
        "Epics and stories are created from the PRD hosted on Confluence, guided by SDE Elements "
        "and PSJIRA. The user picks the work item and AEPLO carries the acceptance context forward.",
        Inches(0.35), Inches(1.55), Inches(9.35), Inches(0.55),
        size=12.2, italic=True, color=LIGHT)

    steps = [
        (TEAL, "01", "PRD Source of Truth",
         "Confluence PRD, architecture notes, release intent and regulatory scope are ingested."),
        (BLUE, "02", "SDE Elements + PSJIRA",
         "AI maps requirements into epics, stories, acceptance criteria and trace links."),
        (PURPLE, "03", "User Picks Work",
         "Epics and stories are listed for the developer or product owner to select."),
        (GREEN, "04", "Branch + QA Context",
         "A feature branch is created with PRD clauses, QA inputs and validation expectations attached."),
    ]
    for i, (clr, num, head, body) in enumerate(steps):
        x = Inches(0.35 + i * 2.38)
        roundrect(sl, x, Inches(2.42), Inches(2.15), Inches(3.9), CARD)
        rect(sl, x, Inches(2.42), Inches(2.15), Inches(0.05), clr)
        oval(sl, x + Inches(0.74), Inches(2.78), Inches(0.66), Inches(0.66), clr)
        txt(sl, num, x + Inches(0.74), Inches(2.78), Inches(0.66), Inches(0.66),
            size=14, bold=True, align=PP_ALIGN.CENTER)
        txt(sl, head, x + Inches(0.15), Inches(3.62), Inches(1.85), Inches(0.52),
            size=12.5, bold=True, color=clr, align=PP_ALIGN.CENTER)
        txt(sl, body, x + Inches(0.16), Inches(4.25), Inches(1.82), Inches(1.65),
            size=10.3, color=LIGHT, align=PP_ALIGN.CENTER)
        if i < 3:
            rect(sl, x + Inches(2.14), Inches(4.32), Inches(0.22), Inches(0.05), CYAN)
    roundrect(sl, Inches(0.75), Inches(6.65), Inches(8.5), Inches(0.42), MID)
    txt(sl, "Outcome: every branch starts with traceable PRD intent, Jira work item, QA input and release-gate evidence plan.",
        Inches(0.9), Inches(6.72), Inches(8.2), Inches(0.24),
        size=9.8, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    bottom_strip(sl, 4)


def slide_05_pg3_firmware():
    sl = prs.slides.add_slide(BLANK)
    bg(sl)
    top_bar(sl, ORANGE)
    section_pill(sl, "PG3 Build", ORANGE)
    title(sl, "PG3 Firmware Generation and Validation")
    divider(sl, Inches(1.45), ORANGE)
    txt(sl,
        "At PG3, AEPLO generates firmware/code from the datasheet and PRD. SDK, BSP, OS/RTOS, "
        "compiler and board context are injected before validation.",
        Inches(0.35), Inches(1.55), Inches(9.35), Inches(0.5),
        size=12.2, italic=True, color=LIGHT)

    card(sl, Inches(0.35), Inches(2.25), Inches(2.85), Inches(3.55), BLUE,
         "Context Injection",
         "- Datasheet and register map\n- PRD acceptance criteria\n- SDK/BSP and HAL\n- OS/RTOS constraints\n- Board and memory limits",
         body_size=10.2)
    card(sl, Inches(3.58), Inches(2.25), Inches(2.85), Inches(3.55), ORANGE,
         "Firmware Generation",
         "- Driver skeletons\n- Peripheral init code\n- Configuration tables\n- Unit and HIL test hooks\n- Review-ready feature branch",
         body_size=10.2)
    card(sl, Inches(6.8), Inches(2.25), Inches(2.85), Inches(3.55), GREEN,
         "Evidence Validation",
         "- Datasheet conformance checks\n- PRD trace verification\n- Static and unit checks\n- JTAG/SWD hardware evidence\n- Ready / Not Ready signal",
         body_size=10.2)

    rect(sl, Inches(3.2), Inches(4.0), Inches(0.35), Inches(0.06), CYAN)
    rect(sl, Inches(6.42), Inches(4.0), Inches(0.35), Inches(0.06), CYAN)
    roundrect(sl, Inches(0.75), Inches(6.22), Inches(8.5), Inches(0.62), PANEL)
    txt(sl, "JTAG Agent adds hardware-backed proof: registers, memory dumps, stack frames, peripheral state and probe logs.",
        Inches(0.92), Inches(6.37), Inches(8.15), Inches(0.25),
        size=10.5, bold=True, color=ORANGE, align=PP_ALIGN.CENTER)
    bottom_strip(sl, 5)


def slide_06_qa_gate_loop():
    sl = prs.slides.add_slide(BLANK)
    bg(sl)
    top_bar(sl, PURPLE)
    section_pill(sl, "QA Release Gate", PURPLE)
    title(sl, "Closed Loop Until QA Says READY")
    divider(sl, Inches(1.45), PURPLE)
    txt(sl,
        "Feature branch, PRD criteria, QA inputs and JTAG evidence are validated together. "
        "A NOT_READY decision becomes prioritized Jira work for the next release loop.",
        Inches(0.35), Inches(1.55), Inches(9.35), Inches(0.5),
        size=12.2, italic=True, color=LIGHT)

    loop = [
        (BLUE, "1", "Feature Branch", "Selected story, generated code and linked PR"),
        (ORANGE, "2", "QA Inputs", "Regression, HIL, field issues and acceptance tests"),
        (TEAL, "3", "Validate", "PRD, datasheet, code, tests and JTAG evidence"),
        (RED, "4", "NOT_READY", "Create defects, stories and priority order"),
        (GREEN, "5", "READY", "Release candidate cleared with evidence pack"),
    ]
    positions = [
        (Inches(0.65), Inches(2.42)),
        (Inches(3.55), Inches(2.18)),
        (Inches(6.45), Inches(2.42)),
        (Inches(5.2), Inches(5.15)),
        (Inches(1.85), Inches(5.15)),
    ]
    for (clr, num, head, body), (x, y) in zip(loop, positions):
        roundrect(sl, x, y, Inches(2.5), Inches(1.28), CARD)
        rect(sl, x, y, Inches(0.05), Inches(1.28), clr)
        oval(sl, x + Inches(0.18), y + Inches(0.33), Inches(0.48), Inches(0.48), clr)
        txt(sl, num, x + Inches(0.18), y + Inches(0.33), Inches(0.48), Inches(0.48),
            size=12, bold=True, align=PP_ALIGN.CENTER)
        txt(sl, head, x + Inches(0.78), y + Inches(0.14), Inches(1.55), Inches(0.32),
            size=12, bold=True, color=clr)
        txt(sl, body, x + Inches(0.78), y + Inches(0.5), Inches(1.55), Inches(0.52),
            size=8.8, color=LIGHT)

    # Connector approximation.
    rect(sl, Inches(3.15), Inches(3.0), Inches(0.38), Inches(0.05), CYAN)
    rect(sl, Inches(6.05), Inches(3.0), Inches(0.38), Inches(0.05), CYAN)
    rect(sl, Inches(7.3), Inches(3.7), Inches(0.05), Inches(1.38), CYAN)
    rect(sl, Inches(4.15), Inches(5.78), Inches(1.0), Inches(0.05), CYAN)
    rect(sl, Inches(2.3), Inches(4.02), Inches(0.05), Inches(1.08), CYAN)
    roundrect(sl, Inches(3.52), Inches(4.02), Inches(2.95), Inches(0.65), PANEL)
    txt(sl, "QA gate guides the next-release backlog until READY is achieved.",
        Inches(3.68), Inches(4.18), Inches(2.65), Inches(0.28),
        size=9.6, bold=True, color=PURPLE, align=PP_ALIGN.CENTER)
    bottom_strip(sl, 6)


def slide_07_stqc_agent():
    sl = prs.slides.add_slide(BLANK)
    bg(sl)
    top_bar(sl, GOLD)
    section_pill(sl, "PG5 STQC Agent", GOLD)
    title(sl, "The Star: STQC Agent for BIS Certification")
    divider(sl, Inches(1.45), GOLD)
    txt(sl,
        "All IoT security and surveillance devices need BIS certification governed by STQC guidelines. "
        "This applies to in-house designs and BTS sales.",
        Inches(0.35), Inches(1.55), Inches(9.35), Inches(0.48),
        size=12.2, italic=True, color=LIGHT)

    card(sl, Inches(0.35), Inches(2.2), Inches(2.85), Inches(4.2), GOLD,
         "Why It Matters",
         "Earlier, developers spent months crawling code, creating evidence, validating rules and packaging documents for auditors.\n\n"
         "AEPLO compresses this to days with repeatable evidence generation.",
         heading_size=12.5, body_size=10.5)
    card(sl, Inches(3.58), Inches(2.2), Inches(2.85), Inches(4.2), CYAN,
         "What STQC Agent Does",
         "- Crawls source and configuration\n- Creates audit evidence\n- Enforces hard rules, including no foreign characters\n- Checks STQC ER guideline coverage\n- Validates submission readiness",
         heading_size=12.5, body_size=10.0)
    card(sl, Inches(6.8), Inches(2.2), Inches(2.85), Inches(4.2), GREEN,
         "If Not Ready",
         "- Highlights the exact failing section\n- Explains the missing evidence\n- Creates Jira tickets and stories\n- Sends fixes back through PG3 and QA gate\n- Re-checks until submission-ready",
         heading_size=12.5, body_size=10.0)
    roundrect(sl, Inches(1.25), Inches(6.68), Inches(7.5), Inches(0.34), MID)
    txt(sl, "Output: READY for STQC submission, or NOT_READY with actionable evidence gaps and Jira remediation.",
        Inches(1.42), Inches(6.72), Inches(7.15), Inches(0.18),
        size=9.3, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    bottom_strip(sl, 7)


def slide_08_results():
    sl = prs.slides.add_slide(BLANK)
    bg(sl)
    top_bar(sl, GREEN)
    section_pill(sl, "Results", GREEN)
    title(sl, "Realistic Quantified Results")
    divider(sl, Inches(1.45), GREEN)
    txt(sl,
        "Pilot-ready targets for efficiency, time saved and token optimization. Values are framed as measured/expected ranges for a controlled demo-to-pilot rollout.",
        Inches(0.35), Inches(1.55), Inches(9.35), Inches(0.5),
        size=11.7, italic=True, color=LIGHT)

    metrics = [
        (GOLD, "STQC Evidence Pack", "6-10 weeks -> 2-4 days", "85-92% effort reduction for first submission pack"),
        (ORANGE, "PG3 Firmware Cycle", "3-5 days -> 1-2 days", "45-60% faster first draft and validation loop"),
        (BLUE, "JTAG Root Cause", "1-2 days -> under 2 hours", "70-90% faster hardware-backed diagnosis"),
        (PURPLE, "QA Gate Triage", "2-3 days -> under 4 hours", "60-75% faster release priority decisions"),
    ]
    for i, (clr, head, metric, body) in enumerate(metrics):
        col, row = i % 2, i // 2
        x = Inches(0.35 + col * 4.85)
        y = Inches(2.2 + row * 1.55)
        roundrect(sl, x, y, Inches(4.6), Inches(1.32), CARD)
        rect(sl, x, y, Inches(0.05), Inches(1.32), clr)
        txt(sl, head, x + Inches(0.18), y + Inches(0.08), Inches(2.25), Inches(0.3),
            size=12, bold=True, color=clr)
        txt(sl, metric, x + Inches(2.45), y + Inches(0.08), Inches(1.95), Inches(0.32),
            size=13, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)
        txt(sl, body, x + Inches(0.18), y + Inches(0.52), Inches(4.15), Inches(0.42),
            size=10.2, color=LIGHT)

    roundrect(sl, Inches(0.35), Inches(5.45), Inches(9.3), Inches(1.2), PANEL)
    rect(sl, Inches(0.35), Inches(5.45), Inches(0.05), Inches(1.2), CYAN)
    txt(sl, "Token Optimization", Inches(0.55), Inches(5.58), Inches(2.1), Inches(0.32),
        size=13, bold=True, color=CYAN)
    txt(sl,
        "35-45% lower token use through routed context, PRD/datasheet chunk retrieval, prompt caching for STQC guidelines, "
        "and compact evidence summaries. Repeated STQC runs can reuse 80%+ of guideline and datasheet context.",
        Inches(2.55), Inches(5.58), Inches(6.85), Inches(0.72),
        size=10.5, color=LIGHT)
    txt(sl, "Close: AEPLO keeps looping across PG3, JTAG, QA and STQC until the release is READY.",
        Inches(0.55), Inches(6.86), Inches(8.95), Inches(0.24),
        size=10.5, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    bottom_strip(sl, 8)


slide_01_cover()
slide_02_problem()
slide_03_architecture()
slide_04_prd_to_backlog()
slide_05_pg3_firmware()
slide_06_qa_gate_loop()
slide_07_stqc_agent()
slide_08_results()

OUT = "/workspace/AI-Driven Embedded Product Lifecycle OrchestratorV10.pptx"
prs.save(OUT)
print(f"Saved: {OUT} ({len(prs.slides)} slides)")
