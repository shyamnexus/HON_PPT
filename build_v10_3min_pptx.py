"""
AI-Driven Embedded Product Lifecycle Orchestrator V10
Honeywell-themed 3-minute PowerPoint builder.

This version preserves the V9 problem statement and architecture storyline,
then compresses the remaining narrative around PRD-driven Jira work,
PG3 firmware generation, QA release gating, JTAG evidence, and the PG5 STQC
agent. It generates embedded Honeywell-style background and illustration
images at build time so the deck is self-contained after creation.
"""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


# Honeywell-oriented palette
HON_RED = RGBColor(0xD7, 0x19, 0x20)
HON_DARK = RGBColor(0x1F, 0x1F, 0x1F)
HON_CHARCOAL = RGBColor(0x2D, 0x2D, 0x2D)
HON_GRAY = RGBColor(0xF3, 0xF4, 0xF6)
HON_LINE = RGBColor(0xD9, 0xDD, 0xE3)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)
TEXT = RGBColor(0x24, 0x24, 0x24)
MUTED = RGBColor(0x68, 0x68, 0x68)
GREEN = RGBColor(0x00, 0x8A, 0x4B)
ORANGE = RGBColor(0xE8, 0x79, 0x00)
BLUE = RGBColor(0x00, 0x66, 0xA4)
PURPLE = RGBColor(0x65, 0x42, 0xA6)
GOLD = RGBColor(0xB8, 0x83, 0x00)
RED_SOFT = RGBColor(0xA6, 0x12, 0x18)

W = Inches(10.0)
H = Inches(7.5)
TOTAL = 9
ASSET_DIR = Path('/workspace/v10_honeywell_assets')

prs = Presentation()
prs.slide_width = W
prs.slide_height = H
BLANK = prs.slide_layouts[6]


def rgb_tuple(color):
    return (color[0], color[1], color[2])


def pil_font(size=28, bold=False):
    candidates = [
        '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf' if bold else '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
        '/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf' if bold else '/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf',
    ]
    for candidate in candidates:
        if Path(candidate).exists():
            return ImageFont.truetype(candidate, size)
    return ImageFont.load_default()


def make_gradient(draw, width, height, start, end):
    for y in range(height):
        ratio = y / max(1, height - 1)
        color = tuple(int(start[i] + (end[i] - start[i]) * ratio) for i in range(3))
        draw.line([(0, y), (width, y)], fill=color)


def draw_circuit(draw, offset_x=0, offset_y=0, scale=1.0, color=(215, 25, 32), muted=(140, 140, 140)):
    nodes = [
        (90, 140), (170, 140), (170, 220), (260, 220), (260, 310),
        (390, 310), (390, 190), (520, 190), (520, 275), (650, 275),
        (735, 170), (865, 170), (865, 255), (1010, 255), (1010, 380),
        (1130, 380), (1130, 500), (950, 500), (950, 610), (760, 610),
        (760, 500), (560, 500), (560, 630), (380, 630), (380, 520),
        (210, 520), (210, 420), (90, 420),
    ]
    pts = [(offset_x + int(x * scale), offset_y + int(y * scale)) for x, y in nodes]
    for a, b in zip(pts, pts[1:]):
        draw.line([a, b], fill=muted, width=max(2, int(3 * scale)))
    for idx, (x, y) in enumerate(pts):
        fill = color if idx % 3 == 0 else (255, 255, 255)
        outline = color if idx % 3 == 0 else muted
        r = max(5, int(8 * scale))
        draw.ellipse((x - r, y - r, x + r, y + r), fill=fill, outline=outline, width=max(1, int(2 * scale)))


def draw_chip(draw, cx, cy, w, h, label, fill=(45, 45, 45), accent=(215, 25, 32)):
    x1, y1 = cx - w // 2, cy - h // 2
    x2, y2 = cx + w // 2, cy + h // 2
    for i in range(12):
        px = x1 - 20 + i * (w + 40) // 11
        draw.rectangle((px, y1 - 16, px + 7, y1), fill=(110, 110, 110))
        draw.rectangle((px, y2, px + 7, y2 + 16), fill=(110, 110, 110))
    for i in range(8):
        py = y1 + i * h // 7
        draw.rectangle((x1 - 16, py, x1, py + 7), fill=(110, 110, 110))
        draw.rectangle((x2, py, x2 + 16, py + 7), fill=(110, 110, 110))
    draw.rounded_rectangle((x1, y1, x2, y2), radius=18, fill=fill, outline=accent, width=4)
    draw.text((cx, cy - 16), label, fill=(255, 255, 255), anchor='mm', font=pil_font(30, True))


def draw_document(draw, x, y, w, h, title, accent=(215, 25, 32)):
    draw.rounded_rectangle((x, y, x + w, y + h), radius=18, fill=(255, 255, 255), outline=(210, 210, 210), width=3)
    draw.rectangle((x, y, x + w, y + 16), fill=accent)
    draw.text((x + 24, y + 42), title, fill=(40, 40, 40), font=pil_font(26, True))
    for i in range(5):
        yy = y + 92 + i * 38
        draw.rectangle((x + 24, yy, x + 52, yy + 18), outline=accent, width=3)
        if i in (0, 2, 3):
            draw.line((x + 30, yy + 9, x + 38, yy + 16, x + 50, yy - 4), fill=accent, width=4)
        draw.line((x + 72, yy + 8, x + w - 32, yy + 8), fill=(180, 180, 180), width=4)


def create_asset(name, theme):
    ASSET_DIR.mkdir(exist_ok=True)
    path = ASSET_DIR / f'{name}.png'
    width, height = 1600, 1200
    img = Image.new('RGB', (width, height), (245, 246, 248))
    draw = ImageDraw.Draw(img)

    if theme == 'cover':
        make_gradient(draw, width, height, (30, 30, 30), (8, 8, 8))
        draw.polygon([(1020, 0), (1600, 0), (1600, 1200), (1160, 1200)], fill=(215, 25, 32))
        draw.polygon([(1160, 0), (1600, 0), (1600, 1200), (1340, 1200)], fill=(166, 18, 24))
        draw_circuit(draw, 600, 80, 0.82, color=(255, 255, 255), muted=(125, 125, 125))
        draw_chip(draw, 1230, 620, 310, 210, 'AEPLO', fill=(34, 34, 34), accent=(255, 255, 255))
    elif theme == 'problem':
        make_gradient(draw, width, height, (255, 255, 255), (235, 237, 240))
        draw.rectangle((0, 0, 1600, 135), fill=(215, 25, 32))
        draw_circuit(draw, 100, 180, 1.05, color=(215, 25, 32), muted=(195, 198, 202))
        for x, y, label in [(250, 410, 'PRD'), (520, 290, 'JIRA'), (850, 370, 'GIT'), (1120, 275, 'JTAG'), (1210, 560, 'QA')]:
            draw.rounded_rectangle((x, y, x + 170, y + 92), radius=14, fill=(255, 255, 255), outline=(215, 25, 32), width=4)
            draw.text((x + 85, y + 46), label, fill=(215, 25, 32), anchor='mm', font=pil_font(30, True))
        draw.text((1180, 930), 'Fragmented signals delay PG1 to PG5', fill=(90, 90, 90), anchor='mm', font=pil_font(38, True))
    elif theme == 'architecture':
        make_gradient(draw, width, height, (248, 249, 250), (229, 232, 236))
        draw.rectangle((0, 0, 1600, 120), fill=(31, 31, 31))
        draw.rectangle((0, 120, 1600, 150), fill=(215, 25, 32))
        for i, label in enumerate(['INPUTS', 'DATA', 'AEPLO CORE', 'GATING', 'OUTCOMES']):
            y = 250 + i * 145
            draw.rounded_rectangle((220, y, 1380, y + 88), radius=16, fill=(255, 255, 255), outline=(215, 25, 32), width=3)
            draw.rectangle((220, y, 430, y + 88), fill=(215, 25, 32))
            draw.text((325, y + 44), label, fill=(255, 255, 255), anchor='mm', font=pil_font(26, True))
            if i < 4:
                draw.line((800, y + 90, 800, y + 140), fill=(215, 25, 32), width=8)
    elif theme == 'prd':
        make_gradient(draw, width, height, (255, 255, 255), (239, 240, 242))
        draw_document(draw, 160, 260, 330, 560, 'Confluence PRD')
        draw_document(draw, 1060, 260, 330, 560, 'PSJIRA')
        for x, text in [(585, 'SDE\nElements'), (775, 'AI\nMapper'), (965, 'Stories')]:
            draw.ellipse((x, 420, x + 150, 570), fill=(215, 25, 32))
            draw.text((x + 75, 495), text, fill=(255, 255, 255), anchor='mm', font=pil_font(25, True), align='center')
        draw.line((495, 535, 585, 495), fill=(215, 25, 32), width=8)
        draw.line((735, 495, 775, 495), fill=(215, 25, 32), width=8)
        draw.line((925, 495, 1060, 535), fill=(215, 25, 32), width=8)
    elif theme == 'firmware':
        make_gradient(draw, width, height, (32, 32, 32), (10, 10, 10))
        draw_circuit(draw, 60, 120, 1.15, color=(215, 25, 32), muted=(95, 95, 95))
        draw_chip(draw, 820, 560, 360, 250, 'MCU', fill=(20, 20, 20), accent=(215, 25, 32))
        for i, text in enumerate(['DATASHEET', 'PRD', 'SDK', 'RTOS']):
            x, y = 180 + i * 285, 870
            draw.rounded_rectangle((x, y, x + 200, y + 64), radius=10, fill=(255, 255, 255), outline=(215, 25, 32), width=3)
            draw.text((x + 100, y + 32), text, fill=(215, 25, 32), anchor='mm', font=pil_font(24, True))
    elif theme == 'jtag':
        make_gradient(draw, width, height, (28, 28, 28), (8, 8, 8))
        draw.rectangle((0, 0, 1600, 135), fill=(215, 25, 32))
        draw_circuit(draw, 40, 160, 1.2, color=(215, 25, 32), muted=(90, 90, 90))
        draw_chip(draw, 1190, 530, 320, 220, 'JTAG', fill=(20, 20, 20), accent=(215, 25, 32))
        labels = [
            (180, 285, 'Bug / Story'),
            (420, 430, 'JTAG\\nOrchestrator'),
            (700, 285, 'Target MCU'),
            (710, 625, 'Evidence'),
        ]
        for x, y, text in labels:
            draw.rounded_rectangle((x, y, x + 190, y + 78), radius=14, fill=(255, 255, 255), outline=(215, 25, 32), width=3)
            draw.text((x + 95, y + 39), text, fill=(215, 25, 32), anchor='mm', font=pil_font(23, True), align='center')
        draw.line((370, 324, 420, 469), fill=(215, 25, 32), width=7)
        draw.line((610, 469, 700, 324), fill=(215, 25, 32), width=7)
        draw.line((800, 365, 805, 625), fill=(215, 25, 32), width=7)
    elif theme == 'qa':
        make_gradient(draw, width, height, (255, 255, 255), (238, 240, 242))
        center = (800, 560)
        radius = 310
        for start, end, color in [(20, 110, (215, 25, 32)), (140, 230, (0, 102, 164)), (260, 350, (0, 138, 75))]:
            for deg in range(start, end, 2):
                import math
                a1 = math.radians(deg)
                a2 = math.radians(deg + 2)
                p1 = (center[0] + int(radius * math.cos(a1)), center[1] + int(radius * math.sin(a1)))
                p2 = (center[0] + int(radius * math.cos(a2)), center[1] + int(radius * math.sin(a2)))
                draw.line((p1, p2), fill=color, width=18)
        draw.ellipse((650, 410, 950, 710), fill=(255, 255, 255), outline=(215, 25, 32), width=6)
        draw.text(center, 'READY?', fill=(215, 25, 32), anchor='mm', font=pil_font(48, True))
        draw.text((800, 930), 'QA Gate closes the next-release loop', fill=(55, 55, 55), anchor='mm', font=pil_font(38, True))
    elif theme == 'stqc':
        make_gradient(draw, width, height, (38, 38, 38), (16, 16, 16))
        draw_document(draw, 180, 210, 440, 650, 'STQC Evidence', accent=(215, 25, 32))
        draw.rounded_rectangle((870, 260, 1360, 780), radius=28, fill=(255, 255, 255), outline=(215, 25, 32), width=6)
        draw.text((1115, 360), 'BIS / STQC', fill=(215, 25, 32), anchor='mm', font=pil_font(46, True))
        draw.ellipse((985, 460, 1245, 720), outline=(215, 25, 32), width=10)
        draw.text((1115, 590), 'READY', fill=(0, 138, 75), anchor='mm', font=pil_font(54, True))
        draw.text((1115, 840), 'Guideline checks | Hard rules | Auditor pack', fill=(245, 245, 245), anchor='mm', font=pil_font(32, True))
    elif theme == 'results':
        make_gradient(draw, width, height, (255, 255, 255), (240, 241, 244))
        draw.rectangle((0, 0, 1600, 135), fill=(215, 25, 32))
        bars = [(260, 760, 'STQC', 0.88), (560, 700, 'PG3', 0.58), (860, 620, 'JTAG', 0.78), (1160, 680, 'TOKENS', 0.42)]
        for x, base, label, pct in bars:
            h = int(520 * pct)
            draw.rectangle((x, base - h, x + 135, base), fill=(215, 25, 32))
            draw.rectangle((x, 240, x + 135, base), outline=(190, 190, 190), width=3)
            draw.text((x + 68, base + 48), label, fill=(70, 70, 70), anchor='mm', font=pil_font(28, True))
        draw.line((180, 780, 1410, 780), fill=(110, 110, 110), width=4)

    path.parent.mkdir(exist_ok=True)
    img.save(path, 'PNG')
    return path


ASSETS = {name: create_asset(name, name) for name in [
    'cover', 'problem', 'architecture', 'prd', 'firmware', 'jtag', 'qa', 'stqc', 'results'
]}


def rect(slide, l, t, w, h, fill, line_color=None, transparency=None):
    shape = slide.shapes.add_shape(1, l, t, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if transparency is not None:
        shape.fill.transparency = transparency
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    return shape


def roundrect(slide, l, t, w, h, fill, line_color=None, transparency=None):
    shape = slide.shapes.add_shape(5, l, t, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if transparency is not None:
        shape.fill.transparency = transparency
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
        color=TEXT, align=PP_ALIGN.LEFT):
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
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return box


def bg(slide, asset_name, overlay=18):
    slide.shapes.add_picture(str(ASSETS[asset_name]), 0, 0, width=W, height=H)
    if overlay:
        rect(slide, 0, 0, W, H, WHITE, transparency=overlay)


def honeywell_header(slide, label, red=True):
    rect(slide, 0, 0, W, Inches(0.56), WHITE)
    rect(slide, 0, Inches(0.56), W, Inches(0.05), HON_RED)
    txt(slide, 'HONEYWELL', Inches(0.35), Inches(0.13), Inches(1.9), Inches(0.28),
        size=14, bold=True, color=HON_RED)
    txt(slide, label.upper(), Inches(7.0), Inches(0.16), Inches(2.65), Inches(0.24),
        size=8.5, bold=True, color=MUTED, align=PP_ALIGN.RIGHT)


def bottom_strip(slide, n):
    rect(slide, 0, H - Inches(0.36), W, Inches(0.36), HON_DARK)
    txt(slide, 'AI-Driven Embedded Product Lifecycle Orchestrator', Inches(0.35), H - Inches(0.27),
        Inches(5.5), Inches(0.18), size=8.5, color=WHITE)
    txt(slide, f'{n} / {TOTAL}', W - Inches(0.82), H - Inches(0.28),
        Inches(0.52), Inches(0.18), size=8.5, color=WHITE, align=PP_ALIGN.RIGHT)


def title(slide, text, color=TEXT):
    txt(slide, text, Inches(0.35), Inches(0.82), Inches(8.95), Inches(0.54),
        size=25, bold=True, color=color)
    rect(slide, Inches(0.35), Inches(1.42), Inches(1.35), Inches(0.06), HON_RED)


def card(slide, x, y, w, h, accent, heading, body, heading_size=12.5, body_size=10.0):
    roundrect(slide, x + Inches(0.04), y + Inches(0.05), w, h, RGBColor(0xC8, 0xCC, 0xD2), transparency=20)
    roundrect(slide, x, y, w, h, WHITE, HON_LINE)
    rect(slide, x, y, Inches(0.08), h, accent)
    txt(slide, heading, x + Inches(0.18), y + Inches(0.13), w - Inches(0.3), Inches(0.42),
        size=heading_size, bold=True, color=accent)
    txt(slide, body, x + Inches(0.18), y + Inches(0.62), w - Inches(0.3), h - Inches(0.76),
        size=body_size, color=TEXT)


def metric_card(slide, x, y, w, h, accent, title_text, value, note):
    roundrect(slide, x, y, w, h, WHITE, HON_LINE)
    rect(slide, x, y, Inches(0.08), h, accent)
    txt(slide, title_text, x + Inches(0.18), y + Inches(0.12), Inches(1.9), Inches(0.28),
        size=11.2, bold=True, color=accent)
    txt(slide, value, x + Inches(2.05), y + Inches(0.08), w - Inches(2.25), Inches(0.33),
        size=12.2, bold=True, color=TEXT, align=PP_ALIGN.RIGHT)
    txt(slide, note, x + Inches(0.18), y + Inches(0.52), w - Inches(0.36), Inches(0.38),
        size=9.2, color=MUTED)


def slide_01_cover():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, 'cover', overlay=0)
    rect(sl, 0, 0, Inches(6.15), H, HON_DARK, transparency=4)
    rect(sl, 0, 0, Inches(0.12), H, HON_RED)
    txt(sl, 'HONEYWELL', Inches(0.35), Inches(0.4), Inches(2.2), Inches(0.35),
        size=18, bold=True, color=HON_RED)
    txt(sl, 'AI-DRIVEN EMBEDDED', Inches(0.35), Inches(1.35), Inches(5.9), Inches(0.62),
        size=31, bold=True, color=WHITE)
    txt(sl, 'PRODUCT LIFECYCLE', Inches(0.35), Inches(2.02), Inches(5.9), Inches(0.62),
        size=31, bold=True, color=WHITE)
    txt(sl, 'ORCHESTRATOR', Inches(0.35), Inches(2.69), Inches(5.9), Inches(0.62),
        size=31, bold=True, color=HON_RED)
    roundrect(sl, Inches(0.35), Inches(3.55), Inches(1.25), Inches(0.34), HON_RED)
    txt(sl, 'VERSION 10', Inches(0.35), Inches(3.59), Inches(1.25), Inches(0.22),
        size=10.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    txt(sl, '3-minute executive story: PRD -> firmware -> JTAG evidence -> QA READY -> STQC submission',
        Inches(0.35), Inches(4.25), Inches(5.45), Inches(0.62), size=13, italic=True, color=WHITE)
    txt(sl, 'Atharva | Manas | Sheraaz | Shyam', Inches(0.35), Inches(5.55), Inches(4.0), Inches(0.25),
        size=10, color=RGBColor(0xDD, 0xDD, 0xDD))
    bottom_strip(sl, 1)


def slide_02_problem():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, 'problem', overlay=12)
    honeywell_header(sl, 'Problem Statement')
    title(sl, 'Embedded PDLC Data Is Fragmented')
    txt(sl,
        'NPI gates rely on disconnected requirements, Jira tickets, Confluence knowledge, GitHub code, JTAG/SWD evidence and QA results - slowing decisions from PG1 to PG5.',
        Inches(0.35), Inches(1.62), Inches(9.2), Inches(0.5), size=12.3, italic=True, color=TEXT)

    silos = [
        (ORANGE, 'PG1', 'Discovery\nBlind Spots', 'Risks and debug hotspots are not visible early'),
        (HON_RED, 'BUG', 'Late Defect\nDiscovery', 'MCU faults surface after coding and integration'),
        (PURPLE, 'TOOL', 'Tool Chain\nFragmentation', 'ROVO, Copilot, Jira, GitHub and JTAG need orchestration'),
        (BLUE, 'TRACE', 'PDLC\nTrace Gaps', 'Requirements, code, tests and evidence drift apart'),
        (GREEN, 'JTAG', 'Embedded\nHardware Gap', 'JTAG/SWD evidence stays outside lifecycle decisions'),
        (GOLD, 'PG5', 'Launch Gate\nUncertainty', 'Readiness depends on manual proof and subjective reviews'),
    ]
    for i, (clr, icon, head, body) in enumerate(silos):
        col, row = i % 3, i // 3
        x = Inches(0.35 + col * 3.15)
        y = Inches(2.38 + row * 2.03)
        roundrect(sl, x, y, Inches(2.95), Inches(1.75), WHITE, HON_LINE)
        rect(sl, x, y, Inches(2.95), Inches(0.07), clr)
        roundrect(sl, x + Inches(0.17), y + Inches(0.22), Inches(0.62), Inches(0.44), clr)
        txt(sl, icon, x + Inches(0.17), y + Inches(0.27), Inches(0.62), Inches(0.18),
            size=8.8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txt(sl, head, x + Inches(0.9), y + Inches(0.16), Inches(1.9), Inches(0.48),
            size=11.7, bold=True, color=TEXT)
        txt(sl, body, x + Inches(0.18), y + Inches(0.78), Inches(2.58), Inches(0.62),
            size=9.7, color=MUTED)
    bottom_strip(sl, 2)


def slide_03_architecture():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, 'architecture', overlay=10)
    honeywell_header(sl, 'Architecture')
    title(sl, 'AEPLO Platform Architecture')
    txt(sl,
        'Connects Jira, Confluence, GitHub, Copilot, ROVO Studio, Atlassian MCP and JTAG/SWD evidence into one embedded PDLC optimization platform.',
        Inches(0.35), Inches(1.62), Inches(9.2), Inches(0.42), size=11.8, italic=True, color=TEXT)

    layers = [
        (HON_RED, 'CONNECTED INPUTS', 'Jira | Confluence | GitHub | Copilot | ROVO Studio | JTAG/SWD | QA Results'),
        (BLUE, 'DATA INTEGRATION LAYER', 'Unified Knowledge Graph | Semantic Traceability Engine | Version History'),
        (HON_RED, 'AEPLO ORCHESTRATION CORE', 'Atlassian MCP | JTAG Accelerator | Fault Analyzer | Release Agent | QA Agent'),
        (PURPLE, 'AUTOMATION & GATING', 'CI/CD Hooks | Controlled Auto-Fix | Auto-PR | Release Gate | Remediation Actions'),
        (GREEN, 'NPI GATE OUTCOMES', 'PG1 Risk Plan | PG3 RCA + PR | PG5 READY / NOT_READY | Audit Evidence'),
    ]
    for i, (clr, label, detail) in enumerate(layers):
        y = Inches(2.17 + i * 0.86)
        roundrect(sl, Inches(0.55), y, Inches(2.1), Inches(0.66), clr)
        txt(sl, label, Inches(0.58), y + Inches(0.18), Inches(2.04), Inches(0.18),
            size=8.2, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        roundrect(sl, Inches(2.85), y, Inches(6.7), Inches(0.66), WHITE, HON_LINE)
        rect(sl, Inches(2.85), y, Inches(0.06), Inches(0.66), clr)
        txt(sl, detail, Inches(3.05), y + Inches(0.19), Inches(6.3), Inches(0.2),
            size=10.2, color=TEXT)
    bottom_strip(sl, 3)


def slide_04_prd_to_backlog():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, 'prd', overlay=20)
    honeywell_header(sl, 'PRD to Backlog')
    title(sl, 'Confluence PRD Becomes Pickable Jira Work')
    txt(sl,
        'Epics and stories are created from the PRD hosted on Confluence, guided by SDE Elements and PSJIRA. The user picks the work item and AEPLO carries the acceptance context forward.',
        Inches(0.35), Inches(1.62), Inches(9.2), Inches(0.5), size=11.8, italic=True, color=TEXT)

    steps = [
        (HON_RED, '01', 'PRD Source of Truth', 'Confluence PRD, architecture notes, release intent and regulatory scope are ingested.'),
        (BLUE, '02', 'SDE Elements + PSJIRA', 'AI maps requirements into epics, stories, acceptance criteria and trace links.'),
        (PURPLE, '03', 'User Picks Work', 'Epics and stories are listed for the developer or product owner to select.'),
        (GREEN, '04', 'Branch + QA Context', 'A feature branch is created with PRD clauses, QA inputs and validation expectations attached.'),
    ]
    for i, (clr, num, head, body) in enumerate(steps):
        x = Inches(0.35 + i * 2.38)
        y = Inches(3.05)
        roundrect(sl, x, y, Inches(2.15), Inches(2.72), WHITE, HON_LINE)
        rect(sl, x, y, Inches(2.15), Inches(0.07), clr)
        oval(sl, x + Inches(0.75), y + Inches(0.32), Inches(0.62), Inches(0.62), clr)
        txt(sl, num, x + Inches(0.75), y + Inches(0.51), Inches(0.62), Inches(0.16),
            size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txt(sl, head, x + Inches(0.15), y + Inches(1.08), Inches(1.85), Inches(0.42),
            size=11.4, bold=True, color=clr, align=PP_ALIGN.CENTER)
        txt(sl, body, x + Inches(0.18), y + Inches(1.58), Inches(1.78), Inches(0.78),
            size=8.9, color=TEXT, align=PP_ALIGN.CENTER)
        if i < 3:
            rect(sl, x + Inches(2.15), y + Inches(1.42), Inches(0.22), Inches(0.04), HON_RED)
    roundrect(sl, Inches(0.75), Inches(6.25), Inches(8.5), Inches(0.42), HON_DARK)
    txt(sl, 'Outcome: every branch starts with traceable PRD intent, Jira work item, QA input and release-gate evidence plan.',
        Inches(0.9), Inches(6.36), Inches(8.2), Inches(0.16), size=8.8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    bottom_strip(sl, 4)


def slide_05_pg3_firmware():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, 'firmware', overlay=0)
    rect(sl, 0, 0, W, H, BLACK, transparency=18)
    honeywell_header(sl, 'PG3 Build')
    title(sl, 'PG3 Firmware Generation and Validation', color=WHITE)
    txt(sl,
        'At PG3, AEPLO generates firmware/code from the datasheet and PRD. SDK, BSP, OS/RTOS, compiler and board context are injected before validation.',
        Inches(0.35), Inches(1.62), Inches(9.2), Inches(0.46), size=11.8, italic=True, color=WHITE)

    card(sl, Inches(0.35), Inches(2.32), Inches(2.85), Inches(3.35), BLUE,
         'Context Injection', '- Datasheet and register map\n- PRD acceptance criteria\n- SDK/BSP and HAL\n- OS/RTOS constraints\n- Board and memory limits', body_size=9.6)
    card(sl, Inches(3.58), Inches(2.32), Inches(2.85), Inches(3.35), HON_RED,
         'Firmware Generation', '- Driver skeletons\n- Peripheral init code\n- Configuration tables\n- Unit and HIL test hooks\n- Review-ready feature branch', body_size=9.6)
    card(sl, Inches(6.8), Inches(2.32), Inches(2.85), Inches(3.35), GREEN,
         'Evidence Validation', '- Datasheet conformance checks\n- PRD trace verification\n- Static and unit checks\n- JTAG/SWD hardware evidence\n- Ready / Not Ready signal', body_size=9.6)
    roundrect(sl, Inches(0.85), Inches(6.1), Inches(8.3), Inches(0.54), HON_RED)
    txt(sl, 'JTAG Agent adds hardware-backed proof: registers, memory dumps, stack frames, peripheral state and probe logs.',
        Inches(1.02), Inches(6.25), Inches(7.95), Inches(0.16), size=9.6, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    bottom_strip(sl, 5)


def slide_06_jtag_debug_copilot():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, 'jtag', overlay=0)
    rect(sl, 0, 0, W, H, BLACK, transparency=22)
    honeywell_header(sl, 'JTAG Debug Copilot')
    title(sl, 'Agent JTAG Debugging: From Failure to Fix', color=WHITE)
    txt(sl,
        'V9 page 8 inserted into the 3-minute flow: the JTAG agent turns live hardware state into root cause, controlled fixes, Jira evidence and release-gate proof.',
        Inches(0.35), Inches(1.62), Inches(9.2), Inches(0.42), size=11.2, italic=True, color=WHITE)

    workflow_path = ASSET_DIR / 'jtag_v9_workflow.png'
    if workflow_path.exists():
        roundrect(sl, Inches(0.35), Inches(2.18), Inches(2.85), Inches(2.85), WHITE, HON_LINE)
        sl.shapes.add_picture(str(workflow_path), Inches(0.48), Inches(2.31), width=Inches(2.58), height=Inches(2.58))
    else:
        card(sl, Inches(0.35), Inches(2.18), Inches(2.85), Inches(2.85), HON_RED,
             'JTAG - Agent Workflow',
             'Bug/Story Ticket -> JTAG Orchestrator -> Target MCU -> Git Provider -> Jira Evidence',
             body_size=10.0)

    roundrect(sl, Inches(0.35), Inches(5.22), Inches(2.85), Inches(0.92), HON_DARK)
    txt(sl, 'JTAG - Agent Workflow', Inches(0.5), Inches(5.38), Inches(2.55), Inches(0.16),
        size=9.0, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    txt(sl, 'Bug/Story -> Orchestrator -> MCU -> Git PR + Jira RCA', Inches(0.5), Inches(5.68),
        Inches(2.55), Inches(0.16), size=7.6, color=WHITE, align=PP_ALIGN.CENTER)

    use_cases = [
        (HON_RED, 'Bootloader -> Application Handoff Failure',
         'SysTick enabled in bootloader, but application vector table had no SysTick handler. Agent correlated PC, vector tables and peripheral state.',
         '45 min RCA vs 2-3 days manual'),
        (BLUE, 'Sub-zero ADC Emulation at Room Temp',
         'ADC stalled at negative temperature. Agent simulated cold ADC state at 25C to validate driver reinit and gain adjustment.',
         'Saved 1 week thermal cycling'),
        (GREEN, 'Boot Failure from Incorrect Configuration',
         'After flashing a new build, all checks looked normal. Agent traced Program Counter and found execution redirected to system ROM instead of flash.',
         'Hardware-backed evidence for fix'),
    ]
    for i, (clr, head, body, result) in enumerate(use_cases):
        y = Inches(2.18 + i * 1.44)
        roundrect(sl, Inches(3.42), y, Inches(6.23), Inches(1.18), WHITE, HON_LINE)
        rect(sl, Inches(3.42), y, Inches(0.08), Inches(1.18), clr)
        txt(sl, head, Inches(3.62), y + Inches(0.1), Inches(3.65), Inches(0.24),
            size=10.8, bold=True, color=clr)
        txt(sl, body, Inches(3.62), y + Inches(0.42), Inches(4.65), Inches(0.42),
            size=8.4, color=TEXT)
        roundrect(sl, Inches(8.26), y + Inches(0.26), Inches(1.18), Inches(0.48), clr)
        txt(sl, result, Inches(8.34), y + Inches(0.35), Inches(1.02), Inches(0.14),
            size=7.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    roundrect(sl, Inches(3.42), Inches(6.55), Inches(6.23), Inches(0.35), HON_RED)
    txt(sl, 'A2A physical interface to hardware: Fault Analyzer | Variable Tracer | Peripheral Inspector | Probing Agent | Controlled Auto Fix',
        Inches(3.58), Inches(6.64), Inches(5.9), Inches(0.12), size=7.8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    bottom_strip(sl, 6)


def slide_06_qa_gate_loop():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, 'qa', overlay=17)
    honeywell_header(sl, 'QA Release Gate')
    title(sl, 'Closed Loop Until QA Says READY')
    txt(sl,
        'Feature branch, PRD criteria, QA inputs and JTAG evidence are validated together. A NOT_READY decision becomes prioritized Jira work for the next release loop.',
        Inches(0.35), Inches(1.62), Inches(9.2), Inches(0.48), size=11.8, italic=True, color=TEXT)

    loop = [
        (BLUE, '1', 'Feature Branch', 'Selected story, generated code and linked PR'),
        (ORANGE, '2', 'QA Inputs', 'Regression, HIL, field issues and acceptance tests'),
        (HON_RED, '3', 'Validate', 'PRD, datasheet, code, tests and JTAG evidence'),
        (RED_SOFT, '4', 'NOT_READY', 'Create defects, stories and priority order'),
        (GREEN, '5', 'READY', 'Release candidate cleared with evidence pack'),
    ]
    positions = [
        (Inches(0.65), Inches(2.48)),
        (Inches(3.55), Inches(2.28)),
        (Inches(6.45), Inches(2.48)),
        (Inches(5.2), Inches(5.05)),
        (Inches(1.85), Inches(5.05)),
    ]
    for (clr, num, head, body), (x, y) in zip(loop, positions):
        roundrect(sl, x, y, Inches(2.5), Inches(1.22), WHITE, HON_LINE)
        rect(sl, x, y, Inches(0.07), Inches(1.22), clr)
        oval(sl, x + Inches(0.2), y + Inches(0.32), Inches(0.46), Inches(0.46), clr)
        txt(sl, num, x + Inches(0.2), y + Inches(0.46), Inches(0.46), Inches(0.14),
            size=10.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txt(sl, head, x + Inches(0.78), y + Inches(0.13), Inches(1.55), Inches(0.28),
            size=11.2, bold=True, color=clr)
        txt(sl, body, x + Inches(0.78), y + Inches(0.48), Inches(1.55), Inches(0.42),
            size=8.3, color=TEXT)
    roundrect(sl, Inches(3.45), Inches(4.05), Inches(3.1), Inches(0.62), HON_DARK)
    txt(sl, 'QA gate guides the next-release backlog until READY is achieved.',
        Inches(3.62), Inches(4.22), Inches(2.76), Inches(0.14), size=8.8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    bottom_strip(sl, 7)


def slide_07_stqc_agent():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, 'stqc', overlay=0)
    rect(sl, 0, 0, W, H, BLACK, transparency=16)
    honeywell_header(sl, 'PG5 STQC Agent')
    title(sl, 'The Star: STQC Agent for BIS Certification', color=WHITE)
    txt(sl,
        'All IoT security and surveillance devices need BIS certification governed by STQC guidelines. This applies to in-house designs and BTS sales.',
        Inches(0.35), Inches(1.62), Inches(9.2), Inches(0.44), size=11.8, italic=True, color=WHITE)

    card(sl, Inches(0.35), Inches(2.22), Inches(2.85), Inches(4.05), GOLD,
         'Why It Matters', 'Earlier, developers spent months crawling code, creating evidence, validating rules and packaging documents for auditors.\n\nAEPLO compresses this to days with repeatable evidence generation.', body_size=9.7)
    card(sl, Inches(3.58), Inches(2.22), Inches(2.85), Inches(4.05), HON_RED,
         'What STQC Agent Does', '- Crawls source and configuration\n- Creates audit evidence\n- Enforces hard rules, including no foreign characters\n- Checks STQC ER guideline coverage\n- Validates submission readiness', body_size=9.2)
    card(sl, Inches(6.8), Inches(2.22), Inches(2.85), Inches(4.05), GREEN,
         'If Not Ready', '- Highlights the exact failing section\n- Explains the missing evidence\n- Creates Jira tickets and stories\n- Sends fixes back through PG3 and QA gate\n- Re-checks until submission-ready', body_size=9.2)
    roundrect(sl, Inches(1.15), Inches(6.56), Inches(7.7), Inches(0.38), HON_RED)
    txt(sl, 'Output: READY for STQC submission, or NOT_READY with actionable evidence gaps and Jira remediation.',
        Inches(1.32), Inches(6.66), Inches(7.35), Inches(0.14), size=8.7, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    bottom_strip(sl, 8)


def slide_08_results():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, 'results', overlay=18)
    honeywell_header(sl, 'Results')
    title(sl, 'Realistic Quantified Results')
    txt(sl,
        'Pilot-ready targets for efficiency, time saved and token optimization. Values are framed as measured/expected ranges for a controlled demo-to-pilot rollout.',
        Inches(0.35), Inches(1.62), Inches(9.2), Inches(0.45), size=11.4, italic=True, color=TEXT)

    metrics = [
        (HON_RED, 'STQC Evidence Pack', '6-10 weeks -> 2-4 days', '85-92% effort reduction for first submission pack'),
        (ORANGE, 'PG3 Firmware Cycle', '3-5 days -> 1-2 days', '45-60% faster first draft and validation loop'),
        (BLUE, 'JTAG Root Cause', '1-2 days -> under 2 hours', '70-90% faster hardware-backed diagnosis'),
        (PURPLE, 'QA Gate Triage', '2-3 days -> under 4 hours', '60-75% faster release priority decisions'),
    ]
    for i, (clr, head, value, note) in enumerate(metrics):
        col, row = i % 2, i // 2
        x = Inches(0.35 + col * 4.85)
        y = Inches(2.25 + row * 1.45)
        metric_card(sl, x, y, Inches(4.6), Inches(1.12), clr, head, value, note)

    roundrect(sl, Inches(0.35), Inches(5.25), Inches(9.3), Inches(1.15), WHITE, HON_LINE)
    rect(sl, Inches(0.35), Inches(5.25), Inches(0.08), Inches(1.15), HON_RED)
    txt(sl, 'Token Optimization', Inches(0.55), Inches(5.4), Inches(2.1), Inches(0.24),
        size=12.4, bold=True, color=HON_RED)
    txt(sl,
        '35-45% lower token use through routed context, PRD/datasheet chunk retrieval, prompt caching for STQC guidelines, and compact evidence summaries. Repeated STQC runs can reuse 80%+ of guideline and datasheet context.',
        Inches(2.55), Inches(5.36), Inches(6.85), Inches(0.6), size=9.6, color=TEXT)
    txt(sl, 'Close: AEPLO keeps looping across PG3, JTAG, QA and STQC until the release is READY.',
        Inches(0.55), Inches(6.72), Inches(8.95), Inches(0.16), size=9.8, bold=True, color=HON_RED, align=PP_ALIGN.CENTER)
    bottom_strip(sl, 9)


slide_01_cover()
slide_02_problem()
slide_03_architecture()
slide_04_prd_to_backlog()
slide_05_pg3_firmware()
slide_06_jtag_debug_copilot()
slide_06_qa_gate_loop()
slide_07_stqc_agent()
slide_08_results()

OUT = '/workspace/AI-Driven Embedded Product Lifecycle OrchestratorV10.pptx'
prs.save(OUT)
print(f'Saved: {OUT} ({len(prs.slides)} slides)')
