"""
Build V10 by reducing and retitling the actual V9 deck.

This intentionally uses AI-Driven Embedded Product Lifecycle OrchestratorV9.pptx
as the source template so V10 inherits the real V9 theme, shapes, media,
colors, typography, footer treatment, and JTAG workflow slide.
"""

import re
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


WORKSPACE = Path("/workspace")
SOURCE = WORKSPACE / "AI-Driven Embedded Product Lifecycle OrchestratorV9.pptx"
OUTPUT = WORKSPACE / "AI-Driven Embedded Product Lifecycle OrchestratorV10.pptx"

# Keep these V9 slides, in their original order:
# 1 cover, 2 problem, 4 architecture, 6 PG1 flow, 7 PG3 flow,
# 8 JTAG debug, 9 PG5 gate, 10 QA loop, 14 evaluation matrix.
KEEP_SLIDES = {1, 2, 4, 6, 7, 8, 9, 10, 14}
TOTAL_SLIDES = 9

BLUE = RGBColor(0x00, 0x72, 0xC6)
CYAN = RGBColor(0x00, 0xC8, 0xF0)
TEAL = RGBColor(0x00, 0xA8, 0x9A)
CARD = RGBColor(0x0F, 0x28, 0x45)
NAVY = RGBColor(0x05, 0x14, 0x2E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xB8, 0xD0, 0xEC)
GREEN = RGBColor(0x00, 0xD4, 0x8A)
ORANGE = RGBColor(0xFF, 0x8C, 0x00)
PURPLE = RGBColor(0x8B, 0x5C, 0xF6)
GOLD = RGBColor(0xF5, 0xC5, 0x18)
RED = RGBColor(0xFF, 0x3A, 0x3A)


def add_shape(slide, shape_type, x, y, w, h, fill, line=None):
    shape = slide.shapes.add_shape(shape_type, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line:
        shape.line.color.rgb = line
    else:
        shape.line.fill.background()
    return shape


def add_text(slide, text, x, y, w, h, size=7.5, bold=False, color=WHITE, align=PP_ALIGN.CENTER):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.margin_left = Inches(0.02)
    tf.margin_right = Inches(0.02)
    tf.margin_top = Inches(0.0)
    tf.margin_bottom = Inches(0.0)
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = "Segoe UI"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def add_tool_badge(slide, x, y, label, icon, color, width=Inches(1.05), height=Inches(0.32)):
    """Add a compact V9-themed tool icon/badge."""
    add_shape(slide, 5, x, y, width, height, CARD, color)
    icon_size = height - Inches(0.08)
    add_shape(slide, 9, x + Inches(0.05), y + Inches(0.04), icon_size, icon_size, color)
    add_text(
        slide,
        icon,
        x + Inches(0.05),
        y + Inches(0.08),
        icon_size,
        Inches(0.12),
        size=5.8,
        bold=True,
        color=WHITE,
    )
    add_text(
        slide,
        label,
        x + icon_size + Inches(0.1),
        y + Inches(0.09),
        width - icon_size - Inches(0.13),
        Inches(0.12),
        size=5.9 if len(label) > 8 else 6.5,
        bold=True,
        color=LIGHT,
        align=PP_ALIGN.LEFT,
    )


def add_tool_row(slide, specs, x=Inches(0.45), y=Inches(6.82), width=Inches(1.03), gap=Inches(0.08)):
    for idx, (label, icon, color) in enumerate(specs):
        add_tool_badge(slide, x + idx * (width + gap), y, label, icon, color, width=width)


def delete_unwanted_slides(prs):
    """Remove all slides not selected for the 3-minute V10 flow."""
    for idx in reversed(range(len(prs.slides))):
        slide_num = idx + 1
        if slide_num in KEEP_SLIDES:
            continue
        slide_id = prs.slides._sldIdLst[idx]
        prs.part.drop_rel(slide_id.rId)
        del prs.slides._sldIdLst[idx]


def move_slide(prs, old_idx, new_idx):
    """Move a slide within the deck by manipulating the slide id list."""
    slide_id = prs.slides._sldIdLst[old_idx]
    prs.slides._sldIdLst.remove(slide_id)
    prs.slides._sldIdLst.insert(new_idx, slide_id)


def set_text(shape, text):
    """Set text while keeping the shape and most inherited formatting intact."""
    if not hasattr(shape, "text_frame"):
        return
    tf = shape.text_frame
    if not tf.paragraphs:
        shape.text = text
        return
    first_para = tf.paragraphs[0]
    if first_para.runs:
        first_para.runs[0].text = text
        # Clear leftover runs and paragraphs without deleting shape-level style.
        for run in first_para.runs[1:]:
            run.text = ""
        for para in tf.paragraphs[1:]:
            for run in para.runs:
                run.text = ""
    else:
        shape.text = text


def set_shape_text(slide, shape_idx, text):
    set_text(slide.shapes[shape_idx], text)


def update_slide_numbers(prs):
    pattern = re.compile(r"^\d+\s*/\s*\d+$")
    for slide_idx, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if hasattr(shape, "text") and pattern.match(shape.text.strip()):
                set_text(shape, f"{slide_idx}  /  {TOTAL_SLIDES}")


def edit_cover(slide):
    set_shape_text(slide, 223, "VERSION 10")
    set_shape_text(
        slide,
        225,
        "3-Minute Embedded Lifecycle Intelligence\n"
        "PRD -> PG3 Firmware -> JTAG Evidence -> QA Ready -> STQC",
    )
    add_tool_row(
        slide,
        [
            ("Confluence", "C", BLUE),
            ("PSJIRA", "J", BLUE),
            ("PG3", "P3", ORANGE),
            ("JTAG", "JT", CYAN),
            ("QA", "QA", GREEN),
            ("STQC", "ST", GOLD),
        ],
        x=Inches(0.45),
        y=Inches(6.25),
        width=Inches(1.15),
    )


def edit_problem(slide):
    # Problem statement remains the same, only slide number changes.
    add_tool_row(
        slide,
        [
            ("Confluence", "C", BLUE),
            ("JIRA", "J", BLUE),
            ("GitHub", "GH", PURPLE),
            ("JTAG", "JT", CYAN),
            ("QA", "QA", GREEN),
        ],
        x=Inches(0.62),
        y=Inches(6.82),
        width=Inches(1.17),
    )


def edit_architecture(slide):
    # Architecture remains the same, only slide number changes.
    add_tool_row(
        slide,
        [
            ("JIRA", "J", BLUE),
            ("Confluence", "C", BLUE),
            ("GitHub", "GH", PURPLE),
            ("Copilot", "AI", GREEN),
            ("ROVO", "R", ORANGE),
            ("MCP", "M", CYAN),
            ("JTAG", "JT", CYAN),
            ("QA", "QA", GREEN),
        ],
        x=Inches(0.46),
        y=Inches(6.85),
        width=Inches(1.05),
        gap=Inches(0.06),
    )


def edit_prd_to_backlog(slide):
    set_shape_text(slide, 3, "PRD TO BACKLOG")
    set_shape_text(slide, 4, "Confluence PRD -> Pickable Jira Work")
    set_shape_text(
        slide,
        6,
        "Epics and stories are created from the PRD hosted on Confluence, "
        "guided by SDE Elements and PSJIRA. The user picks work items and "
        "AEPLO carries acceptance context into the branch.",
    )
    set_shape_text(slide, 18, "INGEST")
    set_shape_text(slide, 19, "Confluence PRD + Context")
    set_shape_text(
        slide,
        21,
        "Reads PRD, SDE Elements, architecture notes, release intent, "
        "security scope and PSJIRA conventions.",
    )
    set_shape_text(slide, 32, "CREATE")
    set_shape_text(slide, 33, "Epics, Stories and Criteria")
    set_shape_text(
        slide,
        35,
        "Maps requirements into epics, stories, acceptance criteria, "
        "trace links and validation evidence expectations.",
    )
    set_shape_text(slide, 40, "PICK")
    set_shape_text(slide, 41, "User Selects Work")
    set_shape_text(
        slide,
        43,
        "Lists stories for the user to pick, then creates a feature branch "
        "with PRD clauses and QA inputs attached.",
    )
    add_tool_row(
        slide,
        [
            ("Confluence", "C", BLUE),
            ("SDE", "SE", TEAL),
            ("PSJIRA", "PJ", BLUE),
            ("JIRA", "J", BLUE),
            ("Git Branch", "GB", PURPLE),
            ("QA Input", "QA", GREEN),
        ],
        x=Inches(0.58),
        y=Inches(6.83),
        width=Inches(1.22),
    )


def edit_pg3_firmware(slide):
    set_shape_text(slide, 3, "PG3 DEVELOPMENT")
    set_shape_text(slide, 4, "PG3 Firmware Generation and Validation")
    set_shape_text(
        slide,
        6,
        "At PG3, AEPLO generates firmware/code from datasheet and PRD, "
        "injecting SDK, BSP, compiler, OS/RTOS and board context before validation.",
    )
    set_shape_text(slide, 11, "DATASHEET + PRD CONTEXT")
    set_shape_text(
        slide,
        12,
        "Requirement, register map, interface contract and acceptance criteria define the work.",
    )
    set_shape_text(slide, 30, "SDK / OS / RTOS INJECTION")
    set_shape_text(
        slide,
        31,
        "SDK, BSP, HAL, compiler, memory map, OS/RTOS and board constraints are injected.",
    )
    set_shape_text(slide, 49, "FIRMWARE GENERATION")
    set_shape_text(
        slide,
        50,
        "Driver skeletons, peripheral init, configuration tables and test hooks are generated.",
    )
    set_shape_text(slide, 68, "VALIDATED FEATURE BRANCH")
    set_shape_text(
        slide,
        69,
        "Generated code becomes a review-ready branch with tests, evidence and PRD traceability.",
    )
    set_shape_text(slide, 72, "DATASHEET AND HARDWARE EVIDENCE LANE")
    set_shape_text(
        slide,
        73,
        "Checks: datasheet conformance | PRD trace | static/unit validation | "
        "JTAG/SWD registers, stack frames, memory dumps, peripheral state and probe logs.",
    )
    set_shape_text(slide, 76, "Faster Draft")
    set_shape_text(slide, 77, "Firmware first draft moves from days to guided PG3 actions.")
    set_shape_text(slide, 80, "Traceable Validation")
    set_shape_text(slide, 81, "Every generated file links PRD, datasheet, QA and hardware evidence.")
    set_shape_text(slide, 84, "PG3 Ready")
    set_shape_text(slide, 85, "Validated firmware flows into QA release and STQC readiness gates.")
    add_tool_row(
        slide,
        [
            ("Datasheet", "DS", GOLD),
            ("PRD", "PR", BLUE),
            ("SDK", "SK", GREEN),
            ("OS/RTOS", "OS", TEAL),
            ("Git Branch", "GB", PURPLE),
            ("JTAG", "JT", CYAN),
        ],
        x=Inches(0.62),
        y=Inches(6.86),
        width=Inches(1.18),
    )


def edit_jtag(slide):
    # Keep V9 page 8's workflow and use cases intact. Only slide number changes.
    add_tool_row(
        slide,
        [
            ("JTAG", "JT", CYAN),
            ("MCU", "MC", ORANGE),
            ("Fault", "FA", RED),
            ("Trace", "TR", BLUE),
            ("Probe", "PB", GREEN),
            ("GitHub", "GH", PURPLE),
            ("JIRA", "J", BLUE),
        ],
        x=Inches(0.52),
        y=Inches(6.86),
        width=Inches(1.05),
        gap=Inches(0.06),
    )


def edit_stqc(slide):
    set_shape_text(slide, 3, "PG5 STQC AGENT")
    set_shape_text(slide, 4, "STQC Agent for BIS Certification")
    set_shape_text(
        slide,
        6,
        "The star of PG5: IoT security and surveillance devices need BIS certification "
        "governed by STQC guidelines for in-house designs and BTS sales.",
    )
    set_shape_text(slide, 11, "Creates STQC\nAudit Evidence")
    set_shape_text(
        slide,
        12,
        "Crawls code, configuration and evidence sources to prepare auditor-ready STQC artifacts.",
    )
    set_shape_text(slide, 17, "Guideline Verdict:\nREADY / NOT_READY")
    set_shape_text(
        slide,
        18,
        "Checks STQC ER guidance, hard rules such as no foreign characters, and submission completeness.",
    )
    set_shape_text(slide, 23, "Jira Remediation\nActions")
    set_shape_text(
        slide,
        24,
        "Every NOT_READY result highlights the failing section and creates stories or defects to fix it.",
    )
    set_shape_text(slide, 27, "STQC SUBMISSION GATE")
    set_shape_text(slide, 34, "Evidence pack ready.\nCleared for auditor review.")
    set_shape_text(slide, 40, "Guideline gaps found.\nCreate Jira stories.")
    add_tool_row(
        slide,
        [
            ("STQC", "ST", GOLD),
            ("BIS", "BI", ORANGE),
            ("Code", "CD", CYAN),
            ("Rules", "ER", RED),
            ("Audit Pack", "AP", GREEN),
            ("JIRA", "J", BLUE),
        ],
        x=Inches(0.82),
        y=Inches(6.83),
        width=Inches(1.2),
    )


def edit_qa_loop(slide):
    set_shape_text(slide, 3, "QA RELEASE GATE")
    set_shape_text(slide, 4, "QA Release Gate Closed Loop")
    set_shape_text(
        slide,
        6,
        "Feature branch, PRD criteria, QA inputs and JTAG evidence are validated together. "
        "NOT_READY decisions become prioritized next-release work.",
    )
    set_shape_text(slide, 12, "Validate Release Evidence")
    set_shape_text(
        slide,
        13,
        "QA results, PRD acceptance criteria, generated firmware tests and JTAG evidence are mapped together.",
    )
    set_shape_text(slide, 19, "Extract Readiness Gaps")
    set_shape_text(
        slide,
        20,
        "AI correlates failures with commits, Jira issues, Confluence decisions, RTOS timing and hardware evidence.",
    )
    set_shape_text(slide, 25, "Guide Next Release")
    set_shape_text(
        slide,
        26,
        "Automated stories and priorities are created until the release gate returns READY.",
    )
    add_tool_row(
        slide,
        [
            ("QA Gate", "QA", GREEN),
            ("PRD", "PR", BLUE),
            ("JTAG", "JT", CYAN),
            ("JIRA", "J", BLUE),
            ("Confluence", "C", BLUE),
            ("Release", "RL", ORANGE),
        ],
        x=Inches(0.7),
        y=Inches(6.83),
        width=Inches(1.16),
    )


def edit_results(slide):
    set_shape_text(slide, 3, "QUANTIFIED RESULTS")
    set_shape_text(slide, 4, "Efficiency, Time Saved and Token Optimization")
    set_shape_text(
        slide,
        6,
        "Realistic pilot targets for the reduced 3-minute V10 story: measurable STQC acceleration, "
        "PG3 generation, JTAG debugging, QA triage and LLM cost control.",
    )
    set_shape_text(slide, 18, "STQC Evidence")
    set_shape_text(
        slide,
        19,
        "- Submission evidence pack: 8-12 weeks -> 1-3 days\n"
        "- First-pack effort reduction: 90-95%\n"
        "- Hard-rule checks: foreign characters, STQC ER coverage, missing sections",
    )
    set_shape_text(slide, 25, "PG3 Firmware")
    set_shape_text(
        slide,
        26,
        "- Firmware first draft and validation: 3-5 days -> 1-2 days\n"
        "- Faster first loop: 45-60%\n"
        "- Context injected: datasheet, PRD, SDK, BSP, OS/RTOS",
    )
    set_shape_text(slide, 32, "JTAG + QA Gate")
    set_shape_text(
        slide,
        33,
        "- JTAG root cause: 1-2 days -> under 2 hours\n"
        "- QA release triage: 2-3 days -> under 4 hours\n"
        "- NOT_READY creates next-release Jira priorities",
    )
    set_shape_text(slide, 39, "Agentic AI execution pattern")
    set_shape_text(
        slide,
        40,
        "- Reusable agents: PRD ingestion, firmware generation, JTAG debug, QA gate, STQC\n"
        "- Evidence traceability: Confluence -> Jira -> GitHub -> QA -> audit pack\n"
        "- Human choice retained at story selection and release gate",
    )
    set_shape_text(slide, 46, "LLM selection strategies and token optimization")
    set_shape_text(
        slide,
        47,
        "- Token reduction: 35-45% using routed context and retrieval\n"
        "- STQC prompt caching: 80%+ reusable guideline/datasheet context on repeat runs\n"
        "- Compact evidence summaries reduce review and audit-prep cost",
    )
    set_shape_text(
        slide,
        49,
        "FOCUS: AEPLO loops across PRD, PG3 firmware, JTAG, QA and STQC until READY.",
    )
    add_tool_row(
        slide,
        [
            ("STQC", "ST", GOLD),
            ("PG3", "P3", ORANGE),
            ("JTAG", "JT", CYAN),
            ("QA", "QA", GREEN),
            ("Tokens", "TK", PURPLE),
        ],
        x=Inches(1.05),
        y=Inches(6.82),
        width=Inches(1.18),
    )


def main():
    prs = Presentation(SOURCE)
    delete_unwanted_slides(prs)
    # V9 source order has PG5 validation before QA2Release; V10 story needs
    # QA Release Gate before the PG5 STQC focus.
    move_slide(prs, 7, 6)

    edit_cover(prs.slides[0])
    edit_problem(prs.slides[1])
    edit_architecture(prs.slides[2])
    edit_prd_to_backlog(prs.slides[3])
    edit_pg3_firmware(prs.slides[4])
    edit_jtag(prs.slides[5])
    edit_qa_loop(prs.slides[6])
    edit_stqc(prs.slides[7])
    edit_results(prs.slides[8])
    update_slide_numbers(prs)

    prs.save(OUTPUT)
    print(f"Saved: {OUTPUT} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
