"""
AI-Driven Embedded Product Lifecycle OrchestratorV3
Enhanced PPT — original content + fully custom graphics (no raster images needed)
Slide size: 10" x 7.5"
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import math, copy

# ── Palette ────────────────────────────────────────────────────────────────────
NAVY        = RGBColor(0x05, 0x14, 0x2E)
DARK_NAVY   = RGBColor(0x02, 0x0A, 0x1A)
BLUE        = RGBColor(0x00, 0x72, 0xC6)
CYAN        = RGBColor(0x00, 0xC8, 0xF0)
TEAL        = RGBColor(0x00, 0xA8, 0x9A)
MID         = RGBColor(0x0E, 0x2A, 0x4A)
PANEL       = RGBColor(0x0B, 0x1F, 0x3A)
CARD        = RGBColor(0x0F, 0x28, 0x45)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT       = RGBColor(0xB8, 0xD0, 0xEC)
DIM         = RGBColor(0x60, 0x80, 0xA8)
GREEN       = RGBColor(0x00, 0xD4, 0x8A)
ORANGE      = RGBColor(0xFF, 0x8C, 0x00)
RED         = RGBColor(0xFF, 0x3A, 0x3A)
PURPLE      = RGBColor(0x8B, 0x5C, 0xF6)
GOLD        = RGBColor(0xF5, 0xC5, 0x18)

W = Inches(10.0)
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
BLANK = prs.slide_layouts[6]

TOTAL = 12

# ══════════════════════════════════════════════════════════════════════════════
# PRIMITIVE HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def rect(slide, l, t, w, h, fill, line_color=None, line_w=None):
    s = slide.shapes.add_shape(1, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line_color:
        s.line.color.rgb = line_color
        if line_w:
            s.line.width = line_w
    else:
        s.line.fill.background()
    return s

def oval(slide, l, t, w, h, fill, line_color=None):
    s = slide.shapes.add_shape(9, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line_color:
        s.line.color.rgb = line_color
    else:
        s.line.fill.background()
    return s

def roundrect(slide, l, t, w, h, fill, line_color=None, adj=20000):
    """Rounded rectangle - adj controls corner radius (0-50000)."""
    s = slide.shapes.add_shape(5, l, t, w, h)  # 5 = rounded rectangle
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line_color:
        s.line.color.rgb = line_color
    else:
        s.line.fill.background()
    # Adjust corner radius
    try:
        adj_elem = s._element.find('.//' + qn('a:prstGeom'))
        if adj_elem is not None:
            avLst = adj_elem.find(qn('a:avLst'))
            if avLst is not None:
                for gd in avLst.findall(qn('a:gd')):
                    if gd.get('name') == 'adj':
                        gd.set('fmla', f'val {adj}')
    except Exception:
        pass
    return s

def txt(slide, text, l, t, w, h,
        size=14, bold=False, italic=False,
        color=WHITE, align=PP_ALIGN.LEFT, wrap=True):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tb.word_wrap = wrap
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = "Segoe UI"
    return tb

def mtxt(slide, lines, l, t, w, h, default_size=14, default_color=WHITE,
         default_align=PP_ALIGN.LEFT):
    """Multi-paragraph text. lines = list of dicts or tuples."""
    tb = slide.shapes.add_textbox(l, t, w, h)
    tb.word_wrap = True
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(lines):
        if isinstance(item, dict):
            text  = item.get('text', '')
            size  = item.get('size', default_size)
            bold  = item.get('bold', False)
            italic= item.get('italic', False)
            color = item.get('color', default_color)
            align = item.get('align', default_align)
        else:
            text, size, bold, color, align = item[0], default_size, False, default_color, default_align
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color
        run.font.name = "Segoe UI"
    return tb

def bg(slide, color=NAVY):
    rect(slide, 0, 0, W, H, color)

def top_bar(slide, color=BLUE, h=Inches(0.055)):
    rect(slide, 0, 0, W, h, color)

def bottom_strip(slide, color=MID, h=Inches(0.42)):
    rect(slide, 0, H - h, W, h, color)

def slide_num(slide, n):
    txt(slide, f"{n}  /  {TOTAL}",
        W - Inches(0.9), H - Inches(0.38),
        Inches(0.8), Inches(0.3),
        size=9, color=DIM, align=PP_ALIGN.RIGHT)

def section_pill(slide, label, color=BLUE, lx=Inches(0.38), ty=Inches(0.48)):
    rr = roundrect(slide, lx, ty, Inches(2.2), Inches(0.3), color, adj=30000)
    txt(slide, label.upper(), lx, ty, Inches(2.2), Inches(0.3),
        size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

def hdivider(slide, y, color=BLUE, lx_frac=0.04, rx_frac=0.96):
    lx = W * lx_frac
    rect(slide, lx, y, W * (rx_frac - lx_frac), Inches(0.016), color)

def slide_title(slide, title, y=Inches(0.82)):
    txt(slide, title, Inches(0.38), y, Inches(9.0), Inches(0.58),
        size=28, bold=True, color=WHITE)

# ══════════════════════════════════════════════════════════════════════════════
# DECORATIVE HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def hex_dots(slide, cx, cy, radius, n_rings=3, color=BLUE, alpha_factor=1.0):
    """Draw a honeycomb dot-grid accent."""
    dot_r = Inches(0.055)
    spacing = Inches(0.28)
    for ring in range(n_rings + 1):
        if ring == 0:
            pts = [(0, 0)]
        else:
            pts = []
            for k in range(6):
                angle = math.pi / 6 + k * math.pi / 3
                for s in range(ring):
                    fx = ring * math.cos(angle) + s * math.cos(angle + math.pi * 2 / 3)
                    fy = ring * math.sin(angle) + s * math.sin(angle + math.pi * 2 / 3)
                    pts.append((fx, fy))
        for (fx, fy) in pts:
            ox = cx + fx * spacing - dot_r / 2
            oy = cy + fy * spacing - dot_r / 2
            if 0 < ox < W and 0 < oy < H:
                c = oval(slide, ox, oy, dot_r, dot_r, color)

def corner_accent(slide, corner='tr', size=Inches(2.0), color=BLUE):
    """Place a geometric triangle-like accent in a corner."""
    if corner == 'tr':
        rect(slide, W - size, 0, size, size * 0.08, color)
        rect(slide, W - size * 0.08, 0, size * 0.08, size, color)
    elif corner == 'bl':
        rect(slide, 0, H - size * 0.08, size, size * 0.08, color)
        rect(slide, 0, H - size, size * 0.08, size, color)

def glow_line(slide, x1_frac, y, x2_frac, color=CYAN, thick=Inches(0.025)):
    rect(slide, W * x1_frac, y - thick/2, W * (x2_frac - x1_frac), thick, color)

def arrow_right(slide, lx, cy, length, color=CYAN, head=Inches(0.18)):
    """Draw a right-pointing arrow."""
    shaft_h = Inches(0.06)
    rect(slide, lx, cy - shaft_h/2, length - head, shaft_h, color)
    # arrowhead as triangle using a thin tall rect (approximation)
    for i in range(8):
        frac = i / 8
        w = head * (1 - frac)
        rect(slide, lx + length - head + head * frac,
             cy - head/2 * (1-frac),
             w * 0.15, head * (1-frac), color)

def flow_arrow(slide, lx, cy, length, color=CYAN):
    """Simpler flow arrow."""
    shaft_h = Inches(0.055)
    rect(slide, lx, cy - shaft_h/2, length * 0.82, shaft_h, color)
    # triangle head
    tip_x = lx + length
    for i in range(12):
        t = i / 12
        hw = Inches(0.14) * (1 - t)
        rect(slide, tip_x - Inches(0.18) * (1-t), cy - hw, Inches(0.015), hw*2, color)

def numbered_circle(slide, n, cx, cy, r=Inches(0.3), fill=BLUE):
    oval(slide, cx - r, cy - r, r*2, r*2, fill)
    txt(slide, str(n), cx - r, cy - r, r*2, r*2,
        size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

def check_circle(slide, cx, cy, r=Inches(0.22), fill=GREEN):
    oval(slide, cx - r, cy - r, r*2, r*2, fill)
    txt(slide, "✓", cx - r, cy - r - Inches(0.02), r*2, r*2,
        size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

def icon_box(slide, icon, lx, ty, sz=Inches(0.55), fill=BLUE):
    roundrect(slide, lx, ty, sz, sz, fill, adj=25000)
    txt(slide, icon, lx, ty, sz, sz,
        size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

def gradient_panel(slide, lx, ty, w, h, color1, color2=None):
    """Simulate gradient with stacked rects."""
    n = 12
    for i in range(n):
        frac = i / n
        if color2 is None:
            # Fade to darker
            r2 = int(color1.rgb >> 16)
            g2 = int((color1.rgb >> 8) & 0xFF)
            b2 = int(color1.rgb & 0xFF)
            r = int(r2 * (1 - frac * 0.5))
            g = int(g2 * (1 - frac * 0.5))
            b = int(b2 * (1 - frac * 0.5))
            c = RGBColor(r, g, b)
        else:
            r1 = int(color1.rgb >> 16); g1 = int((color1.rgb >> 8) & 0xFF); b1 = int(color1.rgb & 0xFF)
            r2 = int(color2.rgb >> 16); g2 = int((color2.rgb >> 8) & 0xFF); b2 = int(color2.rgb & 0xFF)
            r = int(r1 + (r2-r1)*frac); g = int(g1 + (g2-g1)*frac); b = int(b1 + (b2-b1)*frac)
            c = RGBColor(r, g, b)
        rect(slide, lx, ty + h*frac/n*n + h*frac*(1-1/n),
             w, h/n + 1, c)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — COVER
# ══════════════════════════════════════════════════════════════════════════════
def slide_01_cover():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, DARK_NAVY)

    # Left dark panel
    rect(sl, 0, 0, W * 0.58, H, NAVY)

    # Thin vivid left-edge stripe
    rect(sl, 0, 0, Inches(0.07), H, CYAN)

    # Right side subtle tech grid
    grid_color = RGBColor(0x08, 0x22, 0x40)
    for i in range(18):
        rect(sl, W * 0.58, H * i / 18, W * 0.42, Inches(0.012), grid_color)
    for i in range(12):
        rect(sl, W * 0.58 + W * 0.42 * i / 12, 0, Inches(0.012), H, grid_color)

    # Right side big ring accent
    r_big = Inches(3.2)
    cx = W * 0.79; cy = H * 0.47
    oval(sl, cx - r_big, cy - r_big, r_big*2, r_big*2, RGBColor(0x06, 0x1C, 0x38))
    # inner rings
    for rr, clr in [(Inches(2.7), RGBColor(0x07,0x20,0x3E)),
                    (Inches(2.1), RGBColor(0x08,0x25,0x46)),
                    (Inches(1.4), MID),
                    (Inches(0.75), RGBColor(0x0A,0x30,0x5C))]:
        oval(sl, cx - rr, cy - rr, rr*2, rr*2, clr)

    # Cyan ring border
    for rr, clr, thick in [(Inches(3.2), RGBColor(0x00,0x50,0x80), Inches(0.025)),
                            (Inches(1.4), CYAN, Inches(0.04))]:
        for i in range(80):
            angle = 2 * math.pi * i / 80
            ox = cx + rr * math.cos(angle) - thick/2
            oy = cy + rr * math.sin(angle) - thick/2
            if W * 0.56 < ox < W:
                oval(sl, ox, oy, thick, thick, clr)

    # Hexagonal dot field on right
    hex_dots(sl, cx, cy, Inches(2.5), n_rings=4,
             color=RGBColor(0x00, 0x50, 0x88))

    # Title text
    txt(sl, "AEPLO JTAG",
        Inches(0.28), Inches(1.15), Inches(6.0), Inches(0.82),
        size=38, bold=True, color=WHITE)
    txt(sl, "ACCELERATOR FOR",
        Inches(0.28), Inches(1.92), Inches(6.0), Inches(0.82),
        size=38, bold=True, color=WHITE)
    txt(sl, "EMBEDDED PDLC",
        Inches(0.28), Inches(2.68), Inches(6.0), Inches(0.82),
        size=38, bold=True, color=CYAN)

    # Version badge
    roundrect(sl, Inches(0.28), Inches(3.6), Inches(1.4), Inches(0.38), BLUE, adj=30000)
    txt(sl, "VERSION 3.0",
        Inches(0.28), Inches(3.6), Inches(1.4), Inches(0.38),
        size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Tagline
    rect(sl, Inches(0.28), Inches(4.15), Inches(0.055), Inches(0.9), CYAN)
    txt(sl, "Embedded  |  PDLC  |  AI Optimisation\nPG1 Discovery  ->  PG3 Development  ->  PG5 Launch Prep",
        Inches(0.46), Inches(4.18), Inches(5.4), Inches(0.85),
        size=14.5, italic=True, color=LIGHT)

    # Bottom date/confidential
    txt(sl, "May 2026   |   CONFIDENTIAL",
        Inches(0.28), Inches(5.35), Inches(4.0), Inches(0.38),
        size=11, color=DIM)

    # Three accent lines bottom-left
    for i, clr in enumerate([CYAN, BLUE, RGBColor(0x00,0x40,0x70)]):
        rect(sl, Inches(0.28), Inches(6.35 + i * 0.18), Inches(2.5 - i * 0.5), Inches(0.045), clr)

    bottom_strip(sl)
    slide_num(sl, 1)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — PROBLEM STATEMENT
# ══════════════════════════════════════════════════════════════════════════════
def slide_02_problem():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, DARK_NAVY)

    # Full bleed header strip
    rect(sl, 0, 0, W, Inches(1.45), NAVY)
    rect(sl, 0, 0, W, Inches(0.055), ORANGE)
    txt(sl, "THE PROBLEM", Inches(0.38), Inches(0.15), Inches(9.0), Inches(0.38),
        size=10, bold=True, color=ORANGE)
    txt(sl, "Embedded PDLC Data Is Fragmented",
        Inches(0.38), Inches(0.5), Inches(9.3), Inches(0.75),
        size=30, bold=True, color=WHITE)
    hdivider(sl, Inches(1.38), color=ORANGE)

    # Subtitle note
    txt(sl,
        "NPI gates rely on disconnected requirements, Jira tickets, Confluence knowledge, GitHub code, "
        "JTAG/SWD evidence and QA results — slowing decisions from PG1 to PG5.",
        Inches(0.38), Inches(1.52), Inches(9.3), Inches(0.55),
        size=13, italic=True, color=LIGHT)

    # 6 silo cards
    silos = [
        (ORANGE, "PG1",  "Discovery\nBlind Spots",     "Risks and debug hotspots are not visible early"),
        (RED,    "🐛",   "Late Defect\nDiscovery",     "MCU faults surface after coding and integration"),
        (PURPLE, "🔀",   "Tool Chain\nFragmentation",  "ROVO, Copilot, Jira, GitHub and JTAG need orchestration"),
        (BLUE,   "📋",   "PDLC\nTrace Gaps",           "Requirements, code, tests and evidence drift apart"),
        (TEAL,   "🔌",   "Embedded\nHardware Gap",     "JTAG/SWD evidence stays outside lifecycle decisions"),
        (GOLD,   "PG5",  "Launch Gate\nUncertainty",  "Readiness depends on manual proof and subjective reviews"),
    ]

    for i, (clr, icon, title, desc) in enumerate(silos):
        col = i % 3; row = i // 3
        lx = Inches(0.35 + col * 3.15)
        ty = Inches(2.25 + row * 2.35)

        # Card
        roundrect(sl, lx, ty, Inches(2.95), Inches(2.12), CARD, adj=12000)
        # Top accent line
        rect(sl, lx, ty, Inches(2.95), Inches(0.05), clr)

        # Icon box
        roundrect(sl, lx + Inches(0.18), ty + Inches(0.18),
                  Inches(0.5), Inches(0.5), clr, adj=20000)
        txt(sl, icon, lx + Inches(0.18), ty + Inches(0.18),
            Inches(0.5), Inches(0.5), size=16, align=PP_ALIGN.CENTER)

        txt(sl, title, lx + Inches(0.82), ty + Inches(0.18),
            Inches(1.95), Inches(0.55),
            size=13, bold=True, color=WHITE)
        txt(sl, desc, lx + Inches(0.18), ty + Inches(0.82),
            Inches(2.58), Inches(1.1),
            size=11.5, color=LIGHT)

    bottom_strip(sl)
    slide_num(sl, 2)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — SOLUTION
# ══════════════════════════════════════════════════════════════════════════════
def slide_03_solution():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, DARK_NAVY)
    rect(sl, 0, 0, W, Inches(0.055), CYAN)
    section_pill(sl, "Solution", CYAN)
    slide_title(sl, "AEPLO Platform for Embedded PDLC AI Optimisation")
    hdivider(sl, Inches(1.45), color=CYAN)

    # Central hub
    cx = W / 2; cy = Inches(4.1)
    hub_r = Inches(0.88)
    oval(sl, cx - hub_r, cy - hub_r, hub_r*2, hub_r*2, BLUE)
    oval(sl, cx - hub_r*0.78, cy - hub_r*0.78, hub_r*1.56, hub_r*1.56, NAVY)
    # Pulsing rings
    for rr in [Inches(1.15), Inches(1.42), Inches(1.68)]:
        for i in range(48):
            angle = 2 * math.pi * i / 48
            oval(sl, cx + rr * math.cos(angle) - Inches(0.022),
                 cy + rr * math.sin(angle) - Inches(0.022),
                 Inches(0.044), Inches(0.044), CYAN)

    txt(sl, "AEPLO\nJTAG", cx - hub_r, cy - hub_r, hub_r*2, hub_r*2,
        size=14, bold=True, color=CYAN, align=PP_ALIGN.CENTER)

    # 4 pillars arranged around the hub
    pillars = [
        (CYAN,   "Connected Embedded\nData Foundation",
                 "Unifies Jira, Confluence, GitHub, JTAG/SWD evidence, QA and gate context.",
                 -1, -1),  # top-left
        (GREEN,  "JTAG Accelerator\nInside AEPLO",
                 "Hardware-backed RCA, variable tracing and peripheral inspection become PDLC evidence.",
                 1, -1),   # top-right
        (ORANGE, "NPI Gate\nOrchestration",
                 "PG1 discovery, PG3 development and PG5 launch prep become one demo flow.",
                 -1,  1),  # bottom-left
        (PURPLE, "AI Optimisation\nExperiences",
                 "ROVO Studio, Copilot and agents optimise risks, fixes, tests and readiness.",
                 1,  1),   # bottom-right
    ]

    for clr, title, desc, dx, dy in pillars:
        lx = cx + dx * Inches(3.4) - Inches(1.5)
        ty = cy + dy * Inches(1.95) - Inches(0.65)

        # Connector line to hub
        end_x = cx + dx * hub_r * 0.92
        end_y = cy + dy * hub_r * 0.92
        mid_x = (lx + Inches(1.5) + end_x) / 2
        mid_y = (ty + Inches(0.65) + end_y) / 2
        # Simplified connector: two segments
        rect(sl, min(end_x, mid_x), mid_y - Inches(0.015),
             abs(end_x - mid_x), Inches(0.03), clr)
        rect(sl, mid_x - Inches(0.015), min(end_y, mid_y),
             Inches(0.03), abs(end_y - mid_y), clr)

        # Pillar card
        roundrect(sl, lx, ty, Inches(3.0), Inches(1.3), CARD, adj=10000)
        rect(sl, lx, ty, Inches(3.0), Inches(0.045), clr)

        # Icon dot
        oval(sl, lx + Inches(0.15), ty + Inches(0.18),
             Inches(0.3), Inches(0.3), clr)

        txt(sl, title, lx + Inches(0.58), ty + Inches(0.08),
            Inches(2.28), Inches(0.5),
            size=12, bold=True, color=clr)
        txt(sl, desc, lx + Inches(0.15), ty + Inches(0.6),
            Inches(2.7), Inches(0.62),
            size=10.5, color=LIGHT)

    bottom_strip(sl)
    slide_num(sl, 3)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — ARCHITECTURE
# ══════════════════════════════════════════════════════════════════════════════
def slide_04_architecture():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, DARK_NAVY)
    rect(sl, 0, 0, W, Inches(0.055), BLUE)
    section_pill(sl, "Architecture", BLUE)
    slide_title(sl, "AEPLO Platform Architecture")
    hdivider(sl, Inches(1.45), color=BLUE)

    # Notes text
    txt(sl,
        "Connects Jira, Confluence, GitHub, Copilot, ROVO Studio, Atlassian MCP and JTAG/SWD "
        "evidence into one embedded PDLC optimisation platform.",
        Inches(0.38), Inches(1.55), Inches(9.3), Inches(0.45),
        size=12, italic=True, color=LIGHT)

    # 5-layer architecture diagram
    layers = [
        (ORANGE, "CONNECTED INPUTS",
         "Jira · Confluence · GitHub · Copilot · ROVO Studio · JTAG/SWD · QA Results"),
        (BLUE,   "DATA INTEGRATION LAYER",
         "Unified Knowledge Graph  ·  Semantic Traceability Engine  ·  Version History"),
        (CYAN,   "AEPLO ORCHESTRATION CORE",
         "Atlassian MCP · JTAG Accelerator · Fault Analyzer · Release Agent · QA Agent"),
        (PURPLE, "AUTOMATION & GATING",
         "CI/CD Hooks  ·  Controlled Auto-Fix  ·  Auto-PR  ·  Release Gate  ·  Remediation Actions"),
        (GREEN,  "NPI GATE OUTCOMES",
         "PG1 Risk Plan · PG3 RCA + PR · PG5 READY / NOT_READY · Audit Evidence"),
    ]

    for i, (clr, lbl, detail) in enumerate(layers):
        ty = Inches(2.15 + i * 0.96)
        lw = Inches(2.05)
        rw = Inches(7.12)
        # Left label panel
        roundrect(sl, Inches(0.35), ty, lw, Inches(0.82), clr, adj=8000)
        txt(sl, lbl, Inches(0.35), ty, lw, Inches(0.82),
            size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        # Right detail panel
        roundrect(sl, Inches(2.55), ty, rw, Inches(0.82), CARD, adj=8000)
        rect(sl, Inches(2.55), ty, Inches(0.045), Inches(0.82), clr)
        txt(sl, detail, Inches(2.72), ty + Inches(0.18), rw - Inches(0.3), Inches(0.5),
            size=12, color=LIGHT)

        # Arrow between layers
        if i < 4:
            for ax in [Inches(1.0), Inches(5.0)]:
                ay = ty + Inches(0.82)
                rect(sl, ax, ay, Inches(0.05), Inches(0.14), clr)

    bottom_strip(sl)
    slide_num(sl, 4)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — CONNECTED PLATFORM STACK
# ══════════════════════════════════════════════════════════════════════════════
def slide_05_connected_stack():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, DARK_NAVY)
    rect(sl, 0, 0, W, Inches(0.055), GREEN)
    section_pill(sl, "Connected Platform", GREEN)
    slide_title(sl, "ROVO + Copilot + Atlassian MCP + GitHub + Jira + Confluence")
    hdivider(sl, Inches(1.45), color=GREEN)

    txt(sl,
        "The demo platform brings work management, knowledge, code, AI coding assistance and "
        "hardware evidence into AEPLO's embedded PDLC optimisation loop.",
        Inches(0.38), Inches(1.55), Inches(9.3), Inches(0.45),
        size=12, italic=True, color=LIGHT)

    columns = [
        (BLUE, "PLAN & KNOWLEDGE", [
            ("Jira", "NPI epics, coding defects, actions"),
            ("Confluence", "PRD, architecture, test notes"),
            ("PG1 Discovery", "Risk map and demo scope"),
        ]),
        (PURPLE, "BUILD & CODE", [
            ("GitHub", "Branches, PRs, reviews, evidence"),
            ("Copilot", "Patch suggestions and explanations"),
            ("PG3 Development", "RCA, guarded fix and PR"),
        ]),
        (ORANGE, "ORCHESTRATE", [
            ("Atlassian MCP", "Jira + Confluence actions"),
            ("AEPLO", "Agent routing and PDLC knowledge graph"),
            ("ROVO Studio", "NPI demo cockpit and AI experience"),
        ]),
        (GREEN, "VALIDATE & LAUNCH", [
            ("JTAG Accelerator", "MCU register, memory and probe evidence"),
            ("Release Gate", "READY / NOT_READY with action list"),
            ("PG5 Launch Prep", "Validation proof and governance pack"),
        ]),
    ]

    for i, (clr, title, items) in enumerate(columns):
        lx = Inches(0.35 + i * 2.35)
        roundrect(sl, lx, Inches(2.15), Inches(2.12), Inches(4.55), CARD)
        rect(sl, lx, Inches(2.15), Inches(2.12), Inches(0.06), clr)
        txt(sl, title, lx + Inches(0.12), Inches(2.32),
            Inches(1.88), Inches(0.32), size=10.2, bold=True,
            color=clr, align=PP_ALIGN.CENTER)
        for j, (label, body) in enumerate(items):
            ty = Inches(2.95 + j * 1.1)
            roundrect(sl, lx + Inches(0.15), ty, Inches(1.82), Inches(0.82), PANEL)
            oval(sl, lx + Inches(0.28), ty + Inches(0.18), Inches(0.28), Inches(0.28), clr)
            txt(sl, label, lx + Inches(0.65), ty + Inches(0.12),
                Inches(1.1), Inches(0.18), size=8.8, bold=True, color=WHITE)
            txt(sl, body, lx + Inches(0.28), ty + Inches(0.42),
                Inches(1.42), Inches(0.2), size=6.9, color=LIGHT, align=PP_ALIGN.CENTER)
        if i < len(columns) - 1:
            arrow_right(sl, lx + Inches(2.18), Inches(4.25), Inches(0.35), CYAN)

    roundrect(sl, Inches(0.35), Inches(6.95), Inches(9.3), Inches(0.34), MID)
    txt(sl,
        "COMMON IDEA: AEPLO connects enterprise collaboration tools with embedded hardware evidence to optimise NPI gate decisions.",
        Inches(0.55), Inches(6.99), Inches(8.9), Inches(0.16),
        size=8.7, bold=True, color=GREEN, align=PP_ALIGN.CENTER)

    bottom_strip(sl)
    slide_num(sl, 5)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — CHANGE IMPACT ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
def slide_05_change_impact():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, DARK_NAVY)
    rect(sl, 0, 0, W, Inches(0.055), TEAL)
    section_pill(sl, "PG1 Discovery", TEAL)
    slide_title(sl, "PG1 Discovery — AI Risk & Scope Optimisation")
    hdivider(sl, Inches(1.45), color=TEAL)

    txt(sl,
        "AEPLO starts before coding: it connects PRD content, Jira themes, Confluence context, "
        "GitHub history and target-board data to shape a risk-backed embedded demo plan.",
        Inches(0.38), Inches(1.55), Inches(9.3), Inches(0.45),
        size=12, italic=True, color=LIGHT)

    # 3-step flow (Detect → Analyze → Flag)
    steps = [
        (TEAL,   "01", "COLLECT",
         "NPI + Embedded Context",
         "Ingests PRD, Jira epics, Confluence architecture notes, "
         "GitHub defect history and target-board assumptions."),
        (BLUE,   "02", "MAP",
         "PDLC Risk & Traceability",
         "Builds a lifecycle graph from requirements to code areas, "
         "tests, JTAG debug hotspots and validation evidence needs."),
        (ORANGE, "03", "PLAN",
         "PG1 Demo Readiness",
         "Produces a risk-backed demo scope, board setup checklist, "
         "debug scenario and PG3 development entry criteria."),
    ]

    # Horizontal flow line
    cy_flow = Inches(3.62)
    rect(sl, Inches(0.38), cy_flow - Inches(0.015), Inches(9.24), Inches(0.03),
         RGBColor(0x08, 0x28, 0x48))

    for i, (clr, num, verb, sub, detail) in enumerate(steps):
        lx = Inches(0.38 + i * 3.12)
        card_w = Inches(2.88)
        card_top = Inches(2.2)
        card_h = Inches(4.52)

        # Card
        roundrect(sl, lx, card_top, card_w, card_h, CARD, adj=10000)
        rect(sl, lx, card_top, card_w, Inches(0.05), clr)

        # Number pill
        oval(sl, lx + card_w/2 - Inches(0.42),
             cy_flow - Inches(0.42), Inches(0.84), Inches(0.84), clr)
        txt(sl, num, lx + card_w/2 - Inches(0.42),
            cy_flow - Inches(0.42), Inches(0.84), Inches(0.84),
            size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        # Arrow connector (not on last)
        if i < 2:
            arrow_x = lx + card_w + Inches(0.04)
            for j in range(6):
                rect(sl, arrow_x + j * Inches(0.028), cy_flow - Inches(0.016),
                     Inches(0.02), Inches(0.032), TEAL)

        # Verb
        txt(sl, verb, lx, card_top + Inches(0.15),
            card_w, Inches(0.45),
            size=20, bold=True, color=clr, align=PP_ALIGN.CENTER)

        # Sub-title
        txt(sl, sub, lx, card_top + Inches(0.62),
            card_w, Inches(0.4),
            size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        hdivider(sl, card_top + Inches(1.08), color=clr,
                 lx_frac=(lx.pt + Inches(0.18).pt) / W.pt,
                 rx_frac=(lx.pt + card_w.pt - Inches(0.18).pt) / W.pt)

        # Detail
        txt(sl, detail, lx + Inches(0.18),
            card_top + Inches(1.22), card_w - Inches(0.36), Inches(3.1),
            size=11.5, color=LIGHT)

    bottom_strip(sl)
    slide_num(sl, 6)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — JTAG ACCELERATOR
# ══════════════════════════════════════════════════════════════════════════════
def slide_06_jtag_debug_copilot():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, DARK_NAVY)
    rect(sl, 0, 0, W, Inches(0.055), ORANGE)
    section_pill(sl, "PG3 Development", ORANGE)
    slide_title(sl, "PG3 Development — AEPLO JTAG Accelerator")
    hdivider(sl, Inches(1.45), color=ORANGE)

    txt(sl,
        "Inside AEPLO, the JTAG Accelerator converts embedded coding defects into "
        "hardware-backed RCA, controlled code fixes, GitHub PRs and Jira evidence.",
        Inches(0.38), Inches(1.55), Inches(9.3), Inches(0.48),
        size=12, italic=True, color=LIGHT)

    # End-to-end workflow: ticket -> orchestrator -> hardware agents -> PR/evidence.
    stages = [
        (BLUE, "01", "JIRA / CONFLUENCE CONTEXT",
         "Bug, requirement, architecture note and PG3 acceptance criteria define the work."),
        (ORANGE, "02", "JTAG ORCHESTRATOR",
         "Routes requests across AEPLO, IDE, J-Link probe and target MCU session."),
        (PURPLE, "03", "SPECIALIST AGENTS",
         "Fault Analyzer, Variable Tracer, Peripheral Inspector and Memory Parser collaborate."),
        (GREEN, "04", "CONTROLLED AUTO-FIX",
         "Copilot-assisted patch becomes a GitHub branch, PR and review-ready Jira update."),
    ]

    cy = Inches(3.05)
    for i, (clr, num, title, detail) in enumerate(stages):
        lx = Inches(0.35 + i * 2.38)
        card_w = Inches(2.15)

        roundrect(sl, lx, Inches(2.18), card_w, Inches(2.1), CARD, adj=10000)
        rect(sl, lx, Inches(2.18), card_w, Inches(0.05), clr)
        oval(sl, lx + card_w/2 - Inches(0.28), Inches(2.42),
             Inches(0.56), Inches(0.56), clr)
        txt(sl, num, lx + card_w/2 - Inches(0.28), Inches(2.42),
            Inches(0.56), Inches(0.56), size=12, bold=True,
            color=WHITE, align=PP_ALIGN.CENTER)
        txt(sl, title, lx + Inches(0.12), Inches(3.12),
            card_w - Inches(0.24), Inches(0.42), size=10.5,
            bold=True, color=clr, align=PP_ALIGN.CENTER)
        txt(sl, detail, lx + Inches(0.14), Inches(3.58),
            card_w - Inches(0.28), Inches(0.58), size=8.8,
            color=LIGHT, align=PP_ALIGN.CENTER)

        if i < 3:
            flow_arrow(sl, lx + card_w + Inches(0.08), cy,
                       Inches(0.42), color=ORANGE)

    # Hardware evidence lane.
    roundrect(sl, Inches(0.35), Inches(4.62), Inches(9.3), Inches(1.0), PANEL, adj=10000)
    rect(sl, Inches(0.35), Inches(4.62), Inches(0.05), Inches(1.0), ORANGE)
    txt(sl, "PHYSICAL DEBUG EVIDENCE LANE",
        Inches(0.55), Inches(4.72), Inches(3.0), Inches(0.3),
        size=11, bold=True, color=ORANGE)
    txt(sl,
        "Target MCU: STM32 / Cortex-M  |  Interface: JTAG/SWD  |  Evidence: HardFault registers, "
        "stack frames, memory dumps, peripheral state, probe logs and GitHub PR links.",
        Inches(0.55), Inches(5.08), Inches(8.85), Inches(0.38),
        size=10.5, color=LIGHT)

    outcomes = [
        (TEAL, "Faster RCA", "Days of manual setup become guided ROVO/Copilot actions."),
        (BLUE, "Traceable Fixes", "Every patch links Jira, Confluence, GitHub and hardware evidence."),
        (GREEN, "PG3 Ready", "Validated fixes flow into PG5 launch preparation gates."),
    ]
    for i, (clr, title, detail) in enumerate(outcomes):
        lx = Inches(0.35 + i * 3.15)
        roundrect(sl, lx, Inches(5.9), Inches(2.95), Inches(0.95), CARD, adj=10000)
        rect(sl, lx, Inches(5.9), Inches(0.05), Inches(0.95), clr)
        txt(sl, title, lx + Inches(0.14), Inches(5.98),
            Inches(2.68), Inches(0.28), size=11.5, bold=True, color=clr)
        txt(sl, detail, lx + Inches(0.14), Inches(6.3),
            Inches(2.68), Inches(0.38), size=9.3, color=LIGHT)

    bottom_strip(sl)
    slide_num(sl, 7)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — VALIDATION AND LAUNCH GATEKEEPER
# ══════════════════════════════════════════════════════════════════════════════
def slide_06_gatekeeper():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, DARK_NAVY)
    rect(sl, 0, 0, W, Inches(0.055), BLUE)
    section_pill(sl, "PG5 Validation", BLUE)
    slide_title(sl, "PG5 Validation & Launch Preparation Gatekeeper")
    hdivider(sl, Inches(1.45), color=BLUE)

    txt(sl,
        "AEPLO evaluates launch readiness using code, QA, traceability and JTAG/SWD evidence. "
        "The outcome is an unambiguous gate verdict with action-ready remediation.",
        Inches(0.38), Inches(1.55), Inches(9.3), Inches(0.45),
        size=12, italic=True, color=LIGHT)

    # 3 criteria cards on left
    criteria = [
        (GREEN,  "✓", "Validates PG5\nLaunch Readiness",
                 "Coverage, static analysis, regression status, traceability, "
                 "hardware proof and launch evidence are checked together."),
        (BLUE,   "◈", "Structured Verdict:\nREADY / NOT_READY",
                 "Binary, unambiguous gate decision with confidence score, "
                 "Jira evidence summary and GitHub PR links."),
        (ORANGE, "→", "Clear Remediation\nActions",
                 "Every NOT_READY verdict includes a prioritized action list: "
                 "which test, fix, evidence or launch artefact must be updated."),
    ]

    for i, (clr, icon, title, desc) in enumerate(criteria):
        ty = Inches(2.2 + i * 1.65)
        roundrect(sl, Inches(0.35), ty, Inches(5.5), Inches(1.48), CARD, adj=10000)
        rect(sl, Inches(0.35), ty, Inches(0.055), Inches(1.48), clr)

        # Icon
        roundrect(sl, Inches(0.52), ty + Inches(0.42),
                  Inches(0.55), Inches(0.55), clr, adj=20000)
        txt(sl, icon, Inches(0.52), ty + Inches(0.42),
            Inches(0.55), Inches(0.55),
            size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        txt(sl, title, Inches(1.22), ty + Inches(0.08),
            Inches(4.5), Inches(0.55),
            size=13, bold=True, color=clr)
        txt(sl, desc, Inches(1.22), ty + Inches(0.65),
            Inches(4.45), Inches(0.75),
            size=11.5, color=LIGHT)

    # Right side — READY / NOT READY visual
    gate_lx = Inches(6.15)

    # Gate frame
    roundrect(sl, gate_lx, Inches(2.15), Inches(3.48), Inches(4.92), PANEL, adj=10000)
    rect(sl, gate_lx, Inches(2.15), Inches(3.48), Inches(0.055), BLUE)

    txt(sl, "RELEASE GATE", gate_lx, Inches(2.22),
        Inches(3.48), Inches(0.38),
        size=12, bold=True, color=BLUE, align=PP_ALIGN.CENTER)

    hdivider(sl, Inches(2.65), color=BLUE,
             lx_frac=gate_lx.inches/10, rx_frac=9.63/10)

    # READY card
    roundrect(sl, gate_lx + Inches(0.22), Inches(2.78), Inches(3.04), Inches(1.68),
              RGBColor(0x02, 0x28, 0x18), adj=12000)
    rect(sl, gate_lx + Inches(0.22), Inches(2.78), Inches(3.04), Inches(0.05), GREEN)
    # large checkmark
    oval(sl, gate_lx + Inches(1.1), Inches(2.96), Inches(0.68), Inches(0.68), GREEN)
    txt(sl, "✓", gate_lx + Inches(1.1), Inches(2.96),
        Inches(0.68), Inches(0.68),
        size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    txt(sl, "READY", gate_lx + Inches(1.82), Inches(3.0),
        Inches(1.35), Inches(0.42),
        size=24, bold=True, color=GREEN)
    txt(sl, "All gates passed.\nCleared for release.",
        gate_lx + Inches(0.28), Inches(3.55),
        Inches(2.9), Inches(0.75),
        size=11, color=LIGHT)

    # NOT READY card
    roundrect(sl, gate_lx + Inches(0.22), Inches(4.62), Inches(3.04), Inches(1.68),
              RGBColor(0x28, 0x08, 0x08), adj=12000)
    rect(sl, gate_lx + Inches(0.22), Inches(4.62), Inches(3.04), Inches(0.05), RED)
    oval(sl, gate_lx + Inches(1.1), Inches(4.8), Inches(0.68), Inches(0.68), RED)
    txt(sl, "✕", gate_lx + Inches(1.1), Inches(4.8),
        Inches(0.68), Inches(0.68),
        size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    txt(sl, "NOT READY", gate_lx + Inches(1.82), Inches(4.84),
        Inches(1.55), Inches(0.42),
        size=18, bold=True, color=RED)
    txt(sl, "3 blockers found.\nSee remediation list.",
        gate_lx + Inches(0.28), Inches(5.4),
        Inches(2.9), Inches(0.75),
        size=11, color=LIGHT)

    bottom_strip(sl)
    slide_num(sl, 8)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 9 — QA2RELEASE INTELLIGENCE
# ══════════════════════════════════════════════════════════════════════════════
def slide_07_qa2release():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, DARK_NAVY)
    rect(sl, 0, 0, W, Inches(0.055), PURPLE)
    section_pill(sl, "QA2Release Intelligence", PURPLE)
    slide_title(sl, "QA2Release Intelligence")
    hdivider(sl, Inches(1.45), color=PURPLE)

    txt(sl,
        "Creates a closed-loop embedded PDLC: QA results and JTAG evidence drive root-cause "
        "analysis, test optimisation and next-release learning.",
        Inches(0.38), Inches(1.55), Inches(9.3), Inches(0.48),
        size=12, italic=True, color=LIGHT)

    # 3 numbered steps — vertical flow on left, detail on right
    steps = [
        (PURPLE, "1", "Analyze QA Failures",
                 "All QA failures and hardware evidence are ingested — unit, integration, "
                 "HIL, system tests, JTAG registers and probe logs are mapped together."),
        (BLUE,   "2", "Extract Root Causes",
                 "AI correlates failures with commits, Jira issues, Confluence decisions, "
                 "RTOS timing deltas and JTAG/SWD evidence. Root causes are ranked by confidence."),
        (GREEN,  "3", "Generate Next-Release Improvements",
                 "Automated improvement tickets are raised with code suggestions, test additions, "
                 "coverage targets and launch-prep actions for the next NPI cycle."),
    ]

    for i, (clr, num, title, detail) in enumerate(steps):
        ty = Inches(2.22 + i * 1.65)

        # Vertical connector
        if i < 2:
            rect(sl, Inches(0.72), ty + Inches(0.88),
                 Inches(0.04), Inches(0.77), clr)

        # Number circle
        oval(sl, Inches(0.38), ty + Inches(0.22),
             Inches(0.68), Inches(0.68), clr)
        txt(sl, num, Inches(0.38), ty + Inches(0.22),
            Inches(0.68), Inches(0.68),
            size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        # Card
        roundrect(sl, Inches(1.25), ty, Inches(8.38), Inches(1.48), CARD, adj=10000)
        rect(sl, Inches(1.25), ty, Inches(0.05), Inches(1.48), clr)

        txt(sl, title, Inches(1.42), ty + Inches(0.08),
            Inches(7.8), Inches(0.42),
            size=16, bold=True, color=clr)
        txt(sl, detail, Inches(1.42), ty + Inches(0.54),
            Inches(7.8), Inches(0.84),
            size=12, color=LIGHT)

    # Closed-loop banner at bottom
    roundrect(sl, Inches(0.35), Inches(7.0), Inches(9.3), Inches(0.34), MID, adj=10000)
    rect(sl, Inches(0.35), Inches(7.0), Inches(0.05), Inches(0.34), PURPLE)
    txt(sl, "CLOSED-LOOP LIFECYCLE: QA → Root Cause → Next Release Plan → Better QA",
        Inches(0.55), Inches(7.0), Inches(9.0), Inches(0.34),
        size=11, bold=True, color=PURPLE, align=PP_ALIGN.CENTER)

    bottom_strip(sl)
    slide_num(sl, 9)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 10 — DEMO FLOW
# ══════════════════════════════════════════════════════════════════════════════
def slide_08_demo():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, DARK_NAVY)
    rect(sl, 0, 0, W, Inches(0.055), GOLD)
    section_pill(sl, "Demo Flow", GOLD)
    slide_title(sl, "Demo Flow")
    hdivider(sl, Inches(1.45), color=GOLD)

    txt(sl,
        "Live demonstration of how AEPLO executes the NPI gate pipeline from PG1 discovery "
        "through PG3 JTAG-accelerated development to PG5 validation and launch prep.",
        Inches(0.38), Inches(1.55), Inches(9.3), Inches(0.45),
        size=12, italic=True, color=LIGHT)

    # 3 pipeline stages
    stages = [
        (TEAL,   "⚡", "STAGE 1",
         "PG1 → Discovery",
         "• ROVO Studio cockpit opens NPI scenario\n"
         "• Jira + Confluence + GitHub context ingested\n"
         "• AEPLO maps embedded PDLC risks\n"
         "• Demo scope and board setup are generated"),
        (BLUE,   "◈", "STAGE 2",
         "PG3 → Development",
         "• JTAG Accelerator attaches to target MCU\n"
         "• Fault Analyzer captures register evidence\n"
         "• Copilot proposes guarded fix\n"
         "• GitHub PR links code and hardware proof"),
        (PURPLE, "⟳", "STAGE 3",
         "PG5 → Validation & Launch Prep",
         "• QA and JTAG evidence feed Release Gate\n"
         "• READY / NOT_READY verdict issued\n"
         "• Jira actions and Confluence summary created\n"
         "• Next NPI learning loop is updated"),
    ]

    for i, (clr, icon, stage_lbl, title, bullets) in enumerate(stages):
        lx = Inches(0.35 + i * 3.15)
        card_w = Inches(2.95)

        # Main card
        roundrect(sl, lx, Inches(2.18), card_w, Inches(4.68), CARD, adj=10000)
        rect(sl, lx, Inches(2.18), card_w, Inches(0.05), clr)

        # Stage label
        roundrect(sl, lx + Inches(0.18), Inches(2.3), Inches(1.3), Inches(0.28),
                  clr, adj=25000)
        txt(sl, stage_lbl, lx + Inches(0.18), Inches(2.3),
            Inches(1.3), Inches(0.28),
            size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        # Icon circle
        oval(sl, lx + card_w/2 - Inches(0.42), Inches(2.72),
             Inches(0.84), Inches(0.84), clr)
        txt(sl, icon, lx + card_w/2 - Inches(0.42), Inches(2.72),
            Inches(0.84), Inches(0.84),
            size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        txt(sl, title, lx + Inches(0.15), Inches(3.72),
            card_w - Inches(0.3), Inches(0.5),
            size=13, bold=True, color=clr, align=PP_ALIGN.CENTER)

        hdivider(sl, Inches(4.28), color=clr,
                 lx_frac=(lx.inches + 0.18)/10,
                 rx_frac=(lx.inches + card_w.inches - 0.18)/10)

        txt(sl, bullets, lx + Inches(0.2), Inches(4.38),
            card_w - Inches(0.4), Inches(2.3),
            size=11.5, color=LIGHT)

        # Arrow between stages
        if i < 2:
            ax = lx + card_w + Inches(0.04)
            ay = Inches(4.5)
            for j in range(5):
                rect(sl, ax + j * Inches(0.038), ay,
                     Inches(0.025), Inches(0.025), GOLD)

    bottom_strip(sl)
    slide_num(sl, 10)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 11 — BUSINESS OUTCOMES
# ══════════════════════════════════════════════════════════════════════════════
def slide_09_outcomes():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, DARK_NAVY)
    rect(sl, 0, 0, W, Inches(0.055), GREEN)
    section_pill(sl, "Business Outcomes", GREEN)
    slide_title(sl, "Business Outcomes")
    hdivider(sl, Inches(1.45), color=GREEN)

    txt(sl,
        "Measured outcomes for the Coding Embedded Workstream: connected platform data, "
        "hardware-backed debugging, AI optimisation and NPI gate confidence.",
        Inches(0.38), Inches(1.55), Inches(9.3), Inches(0.45),
        size=12, italic=True, color=LIGHT)

    outcomes = [
        (GREEN,  "⬇",  "Reduced\nDefect Leakage",
                 "60%",  "fewer post-release defects",
                 "ML defect prediction + auto-test generation eliminates "
                 "the most common escape vectors before firmware ships."),
        (BLUE,   "◉",  "Faster Firmware\nRCA",
                 "<30m",  "from ticket to root cause",
                 "JTAG Orchestrator captures HardFault registers, stack frames "
                 "and memory evidence without manual tool switching."),
        (ORANGE, "PG",  "NPI Gate\nConfidence",
                 "PG1/3/5", "demo-ready decisions",
                 "AEPLO turns Discovery, Development and Validation evidence "
                 "into clear decisions and action lists."),
        (PURPLE, "⚡", "Embedded PDLC\nOptimisation",
                 "AI",  "risk, test and fix focus",
                 "ROVO Studio, Copilot and AEPLO agents optimise work across "
                 "planning, coding, validation and launch preparation."),
    ]

    for i, (clr, icon, title, metric, metric_lbl, detail) in enumerate(outcomes):
        col = i % 2; row = i // 2
        lx = Inches(0.35 + col * 4.85)
        ty = Inches(2.18 + row * 2.55)
        card_w = Inches(4.6)
        card_h = Inches(2.3)

        roundrect(sl, lx, ty, card_w, card_h, CARD, adj=10000)
        rect(sl, lx, ty, card_w, Inches(0.05), clr)

        # Icon
        oval(sl, lx + Inches(0.2), ty + Inches(0.2),
             Inches(0.6), Inches(0.6), clr)
        txt(sl, icon, lx + Inches(0.2), ty + Inches(0.2),
            Inches(0.6), Inches(0.6),
            size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        txt(sl, title, lx + Inches(0.95), ty + Inches(0.1),
            Inches(2.5), Inches(0.5),
            size=14, bold=True, color=WHITE)

        # Big metric
        txt(sl, metric, lx + Inches(3.15), ty + Inches(0.08),
            Inches(1.25), Inches(0.62),
            size=38, bold=True, color=clr, align=PP_ALIGN.RIGHT)
        txt(sl, metric_lbl, lx + Inches(3.15), ty + Inches(0.72),
            Inches(1.3), Inches(0.3),
            size=9, color=LIGHT, align=PP_ALIGN.RIGHT)

        hdivider(sl, ty + Inches(0.72), color=clr,
                 lx_frac=(lx.inches + 0.18)/10,
                 rx_frac=(lx.inches + card_w.inches - 0.18)/10)

        txt(sl, detail, lx + Inches(0.2), ty + Inches(0.88),
            card_w - Inches(0.4), Inches(1.28),
            size=11.5, color=LIGHT)

    bottom_strip(sl)
    slide_num(sl, 11)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 12 — CONCLUSION
# ══════════════════════════════════════════════════════════════════════════════
def slide_10_conclusion():
    sl = prs.slides.add_slide(BLANK)
    bg(sl, DARK_NAVY)
    rect(sl, 0, 0, W, Inches(0.055), CYAN)

    # Large background ring
    cx = W * 0.75; cy = H * 0.5
    for rr, clr in [(Inches(3.5), RGBColor(0x06,0x1C,0x38)),
                    (Inches(2.8), RGBColor(0x07,0x20,0x3E)),
                    (Inches(2.0), MID),
                    (Inches(1.2), PANEL)]:
        oval(sl, cx - rr, cy - rr, rr*2, rr*2, clr)
    # Cyan ring
    for i in range(64):
        angle = 2 * math.pi * i / 64
        rr = Inches(3.5)
        ox = cx + rr * math.cos(angle) - Inches(0.022)
        oy = cy + rr * math.sin(angle) - Inches(0.022)
        if ox > Inches(0.1):
            oval(sl, ox, oy, Inches(0.04), Inches(0.04), CYAN)
    hex_dots(sl, cx, cy, Inches(2.5), n_rings=4,
             color=RGBColor(0x00, 0x48, 0x80))

    # Left content
    txt(sl, "CONCLUSION", Inches(0.38), Inches(0.72), Inches(5.5), Inches(0.38),
        size=11, bold=True, color=CYAN)
    txt(sl, "Transform Embedded PDLC\nfrom PG1 to PG5",
        Inches(0.38), Inches(1.08), Inches(5.8), Inches(1.05),
        size=28, bold=True, color=WHITE)

    hdivider(sl, Inches(2.22), color=CYAN, lx_frac=0.038, rx_frac=0.58)

    # 3 conclusion messages
    msgs = [
        (TEAL,   "From fragmented embedded data",
                 "Jira, Confluence, GitHub, QA and JTAG/SWD evidence become one "
                 "connected AEPLO PDLC foundation."),
        (BLUE,   "To intelligent NPI action",
                 "Atlassian MCP, ROVO Studio, Copilot and AEPLO agents drive "
                 "PG1 discovery, PG3 development and PG5 launch readiness."),
        (GREEN,  "Optimise every gate",
                 "The JTAG Accelerator adds hardware-backed RCA and controlled fixes "
                 "to AI-optimised validation and launch decisions."),
    ]

    for i, (clr, heading, body) in enumerate(msgs):
        ty = Inches(2.42 + i * 1.55)
        # Accent dot + line
        oval(sl, Inches(0.38), ty + Inches(0.12),
             Inches(0.28), Inches(0.28), clr)
        rect(sl, Inches(0.68), ty + Inches(0.24),
             Inches(0.55), Inches(0.04), clr)

        txt(sl, heading, Inches(1.38), ty + Inches(0.04),
            Inches(4.5), Inches(0.38),
            size=15, bold=True, color=clr)
        txt(sl, body, Inches(1.38), ty + Inches(0.44),
            Inches(4.45), Inches(0.9),
            size=12, color=LIGHT)

    # CTA box inside ring
    roundrect(sl, cx - Inches(1.55), cy - Inches(0.92),
              Inches(3.1), Inches(1.85), BLUE, adj=14000)
    txt(sl, "READY TO TRANSFORM?",
        cx - Inches(1.45), cy - Inches(0.85),
        Inches(2.9), Inches(0.38),
        size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    txt(sl, "aeplo.ai/demo",
        cx - Inches(1.45), cy - Inches(0.42),
        Inches(2.9), Inches(0.38),
        size=18, bold=True, color=CYAN, align=PP_ALIGN.CENTER)
    txt(sl, "hello@aeplo.ai",
        cx - Inches(1.45), cy + Inches(0.0),
        Inches(2.9), Inches(0.35),
        size=13, color=WHITE, align=PP_ALIGN.CENTER)
    txt(sl, "© 2026 AEPLO Technologies Inc.",
        cx - Inches(1.45), cy + Inches(0.62),
        Inches(2.9), Inches(0.3),
        size=9, color=DIM, align=PP_ALIGN.CENTER)

    bottom_strip(sl)
    slide_num(sl, 12)


# ══════════════════════════════════════════════════════════════════════════════
# BUILD
# ══════════════════════════════════════════════════════════════════════════════
slide_01_cover()
slide_02_problem()
slide_03_solution()
slide_04_architecture()
slide_05_connected_stack()
slide_05_change_impact()
slide_06_jtag_debug_copilot()
slide_06_gatekeeper()
slide_07_qa2release()
slide_08_demo()
slide_09_outcomes()
slide_10_conclusion()

OUT = "/workspace/AI-Driven Embedded Product Lifecycle OrchestratorV3.pptx"
prs.save(OUT)
print(f"Saved: {OUT}  ({len(prs.slides)} slides)")
