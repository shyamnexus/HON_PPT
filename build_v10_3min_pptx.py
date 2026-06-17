"""
Reduce the current V10 deck to the requested 5-slide structure.

The deck is reduced in-place from AI-Driven Embedded Product Lifecycle
OrchestratorV10.pptx:
- keep original page 1 as-is
- replace page 2 with one combined summary slide for pages 2, 3, 4, 5 and 7
- keep original page 6 (JTAG) as-is
- keep original page 8 (STQC) as-is
- keep original page 9 (Results) as-is
"""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


WORKSPACE = Path("/workspace")
OUTPUT = WORKSPACE / "AI-Driven Embedded Product Lifecycle OrchestratorV10.pptx"

NAVY = RGBColor(0x05, 0x14, 0x2E)
DARK_NAVY = RGBColor(0x02, 0x0A, 0x1A)
BLUE = RGBColor(0x00, 0x72, 0xC6)
CYAN = RGBColor(0x00, 0xC8, 0xF0)
TEAL = RGBColor(0x00, 0xA8, 0x9A)
MID = RGBColor(0x0E, 0x2A, 0x4A)
PANEL = RGBColor(0x0B, 0x1F, 0x3A)
CARD = RGBColor(0x0F, 0x28, 0x45)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xB8, 0xD0, 0xEC)
DIM = RGBColor(0x60, 0x80, 0xA8)
GREEN = RGBColor(0x00, 0xD4, 0x8A)
ORANGE = RGBColor(0xFF, 0x8C, 0x00)
RED = RGBColor(0xFF, 0x3A, 0x3A)
PURPLE = RGBColor(0x8B, 0x5C, 0xF6)
GOLD = RGBColor(0xF5, 0xC5, 0x18)


def delete_slides_except(prs, keep_indices):
    for idx in reversed(range(len(prs.slides))):
        if idx in keep_indices:
            continue
        slide_id = prs.slides._sldIdLst[idx]
        prs.part.drop_rel(slide_id.rId)
        del prs.slides._sldIdLst[idx]


def clear_slide(slide):
    for shape in list(slide.shapes):
        element = shape._element
        element.getparent().remove(element)


def rect(slide, x, y, w, h, fill, line=None):
    shape = slide.shapes.add_shape(1, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line:
        shape.line.color.rgb = line
    else:
        shape.line.fill.background()
    return shape


def roundrect(slide, x, y, w, h, fill, line=None):
    shape = slide.shapes.add_shape(5, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line:
        shape.line.color.rgb = line
    else:
        shape.line.fill.background()
    return shape


def oval(slide, x, y, w, h, fill, line=None):
    shape = slide.shapes.add_shape(9, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line:
        shape.line.color.rgb = line
    else:
        shape.line.fill.background()
    return shape


def txt(slide, text, x, y, w, h, size=12, bold=False, italic=False, color=WHITE, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(x, y, w, h)
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


def draw_combined_slide(slide):
    clear_slide(slide)
    rect(slide, 0, 0, Inches(10.0), Inches(7.5), DARK_NAVY)
    rect(slide, 0, 0, Inches(10.0), Inches(0.055), CYAN)

    # V9-style subtle tech grid.
    grid_color = RGBColor(0x08, 0x22, 0x40)
    for i in range(18):
        rect(slide, Inches(0), Inches(0.7 + i * 0.34), Inches(10.0), Inches(0.006), grid_color)
    for i in range(13):
        rect(slide, Inches(0.4 + i * 0.75), Inches(0.7), Inches(0.006), Inches(5.85), grid_color)

    roundrect(slide, Inches(0.38), Inches(0.48), Inches(2.35), Inches(0.3), BLUE)
    txt(slide, "COMBINED FLOW", Inches(0.38), Inches(0.53), Inches(2.35), Inches(0.12), size=8.5, bold=True, align=PP_ALIGN.CENTER)
    txt(slide, "One Slide Summary: Problem -> Architecture -> PRD -> PG3 -> QA Gate", Inches(0.38), Inches(0.86), Inches(9.0), Inches(0.48), size=23, bold=True)
    rect(slide, Inches(0.38), Inches(1.42), Inches(9.2), Inches(0.018), CYAN)
    txt(
        slide,
        "Combines original pages 2, 3, 4, 5 and 7 while preserving the requested V10 pages 1, 6, 8 and 9 unchanged.",
        Inches(0.38), Inches(1.54), Inches(9.25), Inches(0.36), size=10.5, italic=True, color=LIGHT,
    )

    cards = [
        (ORANGE, "1", "Problem", "Fragmented PRD, Jira, Confluence, GitHub, QA and hardware evidence slow PG1-to-PG5 decisions."),
        (BLUE, "2", "Architecture", "AEPLO connects inputs into a knowledge graph, orchestration core, automation and release-gate outcomes."),
        (TEAL, "3", "PRD -> Stories", "Confluence PRD guided by SDE Elements and PSJIRA creates epics/stories for users to pick and branch."),
        (PURPLE, "4", "PG3 Firmware", "Datasheet + PRD + SDK/BSP + OS/RTOS context generate and validate firmware/code."),
        (GREEN, "5", "QA Closed Loop", "QA inputs, PRD and JTAG evidence produce READY / NOT_READY; gaps become next-release priorities."),
    ]
    for i, (color, num, title, body) in enumerate(cards):
        x = Inches(0.42 + i * 1.88)
        y = Inches(2.25)
        roundrect(slide, x, y, Inches(1.7), Inches(3.35), CARD)
        rect(slide, x, y, Inches(1.7), Inches(0.055), color)
        oval(slide, x + Inches(0.55), y + Inches(0.28), Inches(0.58), Inches(0.58), color)
        txt(slide, num, x + Inches(0.55), y + Inches(0.42), Inches(0.58), Inches(0.12), size=12, bold=True, align=PP_ALIGN.CENTER)
        txt(slide, title, x + Inches(0.12), y + Inches(1.04), Inches(1.46), Inches(0.35), size=10.6, bold=True, color=color, align=PP_ALIGN.CENTER)
        txt(slide, body, x + Inches(0.14), y + Inches(1.55), Inches(1.42), Inches(1.25), size=8.4, color=LIGHT, align=PP_ALIGN.CENTER)
        if i < len(cards) - 1:
            rect(slide, x + Inches(1.72), y + Inches(1.66), Inches(0.14), Inches(0.04), CYAN)

    roundrect(slide, Inches(0.72), Inches(6.08), Inches(8.55), Inches(0.5), PANEL)
    txt(
        slide,
        "Output: users pick traceable work, create feature branches, validate PG3 firmware with JTAG evidence, and loop through QA until READY.",
        Inches(0.9), Inches(6.22), Inches(8.2), Inches(0.16), size=9.2, bold=True, color=CYAN, align=PP_ALIGN.CENTER,
    )
    rect(slide, 0, Inches(7.08), Inches(10.0), Inches(0.42), MID)
    txt(slide, "2  /  5", Inches(8.92), Inches(7.18), Inches(0.65), Inches(0.16), size=8.5, color=DIM, align=PP_ALIGN.RIGHT)
    txt(slide, "HONEYWELL", Inches(8.2), Inches(7.18), Inches(0.9), Inches(0.16), size=8.5, bold=True, color=DIM, align=PP_ALIGN.RIGHT)


def main():
    prs = Presentation(OUTPUT)
    if len(prs.slides) >= 9:
        # Keep original pages 1, 2 (converted summary), 6, 8 and 9.
        keep = {0, 1, 5, 7, 8}
        draw_combined_slide(prs.slides[1])
        delete_slides_except(prs, keep)
    elif len(prs.slides) == 5:
        draw_combined_slide(prs.slides[1])
    else:
        raise RuntimeError(f"Expected 9-slide source or 5-slide reduced deck, found {len(prs.slides)} slides")

    prs.save(OUTPUT)
    print(f"Saved: {OUTPUT} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
