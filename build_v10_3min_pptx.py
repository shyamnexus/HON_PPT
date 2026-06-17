"""
Build V10 by reducing and retitling the actual V9 deck.

This intentionally uses AI-Driven Embedded Product Lifecycle OrchestratorV9.pptx
as the source template so V10 inherits the real V9 theme, shapes, media,
colors, typography, footer treatment, and JTAG workflow slide.
"""

import re
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches


WORKSPACE = Path("/workspace")
SOURCE = WORKSPACE / "AI-Driven Embedded Product Lifecycle OrchestratorV9.pptx"
OUTPUT = WORKSPACE / "AI-Driven Embedded Product Lifecycle OrchestratorV10.pptx"
ICON_DIR = WORKSPACE / "v10_tool_icons"

# Keep these V9 slides, in their original order:
# 1 cover, 2 problem, 4 architecture, 6 PG1 flow, 7 PG3 flow,
# 8 JTAG debug, 9 PG5 gate, 10 QA loop, 14 evaluation matrix.
KEEP_SLIDES = {1, 2, 4, 6, 7, 8, 9, 10, 14}
TOTAL_SLIDES = 9


def pil_font(size=42, bold=True):
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    ]
    for candidate in candidates:
        if Path(candidate).exists():
            return ImageFont.truetype(candidate, size)
    return ImageFont.load_default()


def save_icon(name, draw_fn):
    ICON_DIR.mkdir(exist_ok=True)
    path = ICON_DIR / f"{name}.png"
    img = Image.new("RGBA", (512, 512), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw_fn(draw, img)
    img.save(path)
    return path


def draw_text_center(draw, xy, text, fill, size=70):
    draw.text(xy, text, fill=fill, anchor="mm", font=pil_font(size, True))


def rounded_poly(draw, points, fill):
    draw.polygon(points, fill=fill)


def create_tool_icons():
    """Create transparent PNG tool icons used by the V9-derived deck."""

    def jira(draw, _):
        blue = (38, 132, 255, 255)
        dark = (0, 82, 204, 255)
        rounded_poly(draw, [(244, 0), (512, 0), (512, 112), (356, 112), (244, 0)], blue)
        rounded_poly(draw, [(120, 122), (390, 122), (390, 234), (232, 234), (120, 122)], dark)
        rounded_poly(draw, [(0, 246), (278, 246), (278, 358), (156, 358), (156, 512), (0, 512)], blue)
        rounded_poly(draw, [(278, 234), (390, 234), (390, 512), (278, 390)], dark)
        rounded_poly(draw, [(400, 112), (512, 112), (512, 266), (400, 266)], blue)

    def azure_devops(draw, _):
        blue = (0, 120, 212, 255)
        light = (80, 170, 240, 255)
        draw.polygon([(42, 176), (224, 124), (224, 42), (376, 142), (224, 202), (42, 236)], fill=light)
        draw.polygon([(42, 176), (104, 144), (104, 348), (42, 322)], fill=blue)
        draw.polygon([(104, 348), (376, 370), (376, 142), (470, 122), (470, 385), (364, 470)], fill=blue)
        draw.polygon([(196, 410), (196, 512), (104, 348)], fill=(0, 92, 172, 255))

    def rovo(draw, _):
        draw.rounded_rectangle((43, 110, 205, 402), radius=70, fill=(38, 106, 214, 255))
        draw.rounded_rectangle((205, 16, 370, 188), radius=70, fill=(181, 94, 235, 255))
        draw.rounded_rectangle((307, 105, 470, 402), radius=70, fill=(255, 171, 0, 255))
        draw.rounded_rectangle((177, 325, 342, 496), radius=70, fill=(137, 186, 38, 255))
        draw.polygon([(205, 16), (370, 110), (205, 204)], fill=(181, 94, 235, 255))
        draw.polygon([(43, 248), (205, 154), (205, 342)], fill=(38, 106, 214, 255))
        draw.polygon([(307, 105), (470, 198), (307, 496)], fill=(255, 171, 0, 255))

    def copilot(draw, _):
        colors = [(0, 170, 255, 255), (124, 203, 55, 255), (255, 196, 0, 255), (255, 80, 80, 255), (211, 79, 206, 255)]
        for idx, color in enumerate(colors):
            y = 70 + idx * 62
            draw.rounded_rectangle((70 + idx * 18, y, 340 + idx * 18, y + 110), radius=55, fill=color)
        draw.rounded_rectangle((230, 180, 495, 455), radius=70, fill=(213, 76, 202, 235))
        draw.rounded_rectangle((32, 60, 280, 358), radius=70, outline=(20, 145, 230, 255), width=34)

    def github(draw, _):
        draw.ellipse((8, 8, 504, 504), fill=(66, 133, 210, 255))
        draw.ellipse((94, 96, 418, 388), fill=(0, 0, 0, 255))
        draw.polygon([(150, 132), (120, 58), (205, 106)], fill=(0, 0, 0, 255))
        draw.polygon([(362, 132), (392, 58), (307, 106)], fill=(0, 0, 0, 255))
        draw.rectangle((210, 330, 302, 506), fill=(0, 0, 0, 255))
        draw.arc((52, 300, 214, 454), 95, 305, fill=(0, 0, 0, 255), width=38)

    def confluence(draw, _):
        blue = (38, 132, 255, 255)
        dark = (0, 82, 204, 255)
        layer = Image.new("RGBA", (512, 512), (0, 0, 0, 0))
        d = ImageDraw.Draw(layer)
        d.rounded_rectangle((42, 92, 470, 212), radius=38, fill=blue)
        layer = layer.rotate(22, resample=Image.Resampling.BICUBIC, center=(256, 256))
        _.alpha_composite(layer)
        layer = Image.new("RGBA", (512, 512), (0, 0, 0, 0))
        d = ImageDraw.Draw(layer)
        d.rounded_rectangle((42, 300, 470, 420), radius=38, fill=dark)
        layer = layer.rotate(22, resample=Image.Resampling.BICUBIC, center=(256, 256))
        _.alpha_composite(layer)

    def stqc(draw, _):
        draw.polygon([(256, 28), (438, 104), (410, 360), (256, 492), (102, 360), (74, 104)], fill=(245, 197, 24, 255))
        draw.polygon([(256, 70), (394, 128), (372, 338), (256, 448), (140, 338), (118, 128)], fill=(15, 40, 69, 255))
        draw_text_center(draw, (256, 235), "STQC", (255, 255, 255, 255), 58)
        draw_text_center(draw, (256, 302), "ER", (245, 197, 24, 255), 44)

    def jtag(draw, _):
        draw.rounded_rectangle((132, 132, 380, 380), radius=32, fill=(15, 40, 69, 255), outline=(0, 200, 240, 255), width=16)
        for i in range(8):
            x = 96 + i * 46
            draw.rectangle((x, 96, x + 18, 132), fill=(0, 200, 240, 255))
            draw.rectangle((x, 380, x + 18, 416), fill=(0, 200, 240, 255))
        for i in range(5):
            y = 156 + i * 42
            draw.rectangle((72, y, 132, y + 16), fill=(0, 200, 240, 255))
            draw.rectangle((380, y, 440, y + 16), fill=(0, 200, 240, 255))
        draw_text_center(draw, (256, 256), "JTAG", (255, 255, 255, 255), 56)

    def qa(draw, _):
        draw.ellipse((44, 44, 468, 468), fill=(0, 212, 138, 255))
        draw.line([(142, 270), (222, 350), (382, 170)], fill=(255, 255, 255, 255), width=46, joint="curve")

    def doc(draw, _):
        draw.rounded_rectangle((120, 48, 392, 464), radius=24, fill=(255, 255, 255, 255), outline=(0, 114, 198, 255), width=16)
        draw.rectangle((150, 132, 362, 152), fill=(0, 114, 198, 255))
        draw.rectangle((150, 202, 362, 222), fill=(0, 114, 198, 255))
        draw.rectangle((150, 272, 318, 292), fill=(0, 114, 198, 255))
        draw_text_center(draw, (256, 382), "PRD", (0, 114, 198, 255), 58)

    def sdk(draw, _):
        draw.polygon([(256, 42), (438, 148), (256, 254), (74, 148)], fill=(0, 212, 138, 255))
        draw.polygon([(74, 148), (256, 254), (256, 470), (74, 360)], fill=(0, 168, 154, 255))
        draw.polygon([(438, 148), (256, 254), (256, 470), (438, 360)], fill=(0, 114, 198, 255))
        draw_text_center(draw, (256, 296), "SDK", (255, 255, 255, 255), 56)

    icons = {
        "jira": jira,
        "azure_devops": azure_devops,
        "rovo": rovo,
        "copilot": copilot,
        "github": github,
        "confluence": confluence,
        "stqc": stqc,
        "jtag": jtag,
        "qa": qa,
        "prd": doc,
        "sdk": sdk,
    }
    for name, fn in icons.items():
        save_icon(name, fn)


def add_icon(slide, name, x, y, size=0.35):
    path = ICON_DIR / f"{name}.png"
    if path.exists():
        slide.shapes.add_picture(str(path), x, y, width=Inches(size), height=Inches(size))


def add_icon_row(slide, names, x, y, size=0.34, gap=0.12):
    for idx, name in enumerate(names):
        add_icon(slide, name, Inches(x + idx * (size + gap)), Inches(y), size)


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
    add_icon_row(slide, ["confluence", "jira", "github", "jtag", "qa", "stqc"], 0.55, 6.22, 0.36, 0.16)


def edit_problem(slide):
    # Problem statement remains the same, only slide number changes.
    add_icon_row(slide, ["jira", "confluence", "github", "jtag", "qa"], 0.72, 6.62, 0.32, 0.18)


def edit_architecture(slide):
    # Architecture remains the same, only slide number changes.
    add_icon_row(
        slide,
        ["jira", "confluence", "github", "copilot", "rovo", "azure_devops", "jtag", "qa"],
        0.5,
        6.62,
        0.29,
        0.14,
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
    add_icon_row(slide, ["confluence", "rovo", "jira", "github", "qa"], 0.76, 6.62, 0.32, 0.22)


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
    add_icon_row(slide, ["prd", "sdk", "copilot", "github", "azure_devops", "jtag"], 0.72, 6.66, 0.3, 0.18)


def edit_jtag(slide):
    # Keep V9 page 8's workflow and use cases intact. Only slide number changes.
    add_icon_row(slide, ["jtag", "github", "jira"], 0.55, 6.62, 0.32, 0.18)


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
    add_icon_row(slide, ["stqc", "prd", "jira", "confluence"], 0.74, 6.62, 0.32, 0.2)


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
    add_icon_row(slide, ["qa", "prd", "jtag", "jira", "confluence"], 0.74, 6.62, 0.32, 0.2)


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
    add_icon_row(slide, ["stqc", "sdk", "jtag", "qa", "copilot"], 1.1, 6.62, 0.32, 0.2)


def main():
    create_tool_icons()
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
