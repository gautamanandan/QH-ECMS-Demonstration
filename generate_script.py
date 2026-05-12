from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ── Colours ──────────────────────────────────────────────────────────────────
QH_MAROON   = RGBColor(0x7B, 0x1C, 0x1C)
QH_MAROON_L = RGBColor(0xC0, 0x39, 0x2B)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GREY  = RGBColor(0xF2, 0xF3, 0xF5)
MID_GREY    = RGBColor(0x6B, 0x72, 0x80)
DARK        = RGBColor(0x11, 0x18, 0x27)
YELLOW_BG   = RGBColor(0xFF, 0xFB, 0xEB)
YELLOW_BD   = RGBColor(0xF5, 0x9E, 0x0B)
BLUE_BG     = RGBColor(0xEF, 0xF6, 0xFF)
BLUE_BD     = RGBColor(0x1D, 0x4E, 0xD8)
ROW_GARY    = RGBColor(0xF9, 0xF9, 0xF9)
ROW_TRICIA  = RGBColor(0xFF, 0xFF, 0xFF)
ROW_GAUTAM  = RGBColor(0xF3, 0xF4, 0xF6)
ROW_WOW     = RGBColor(0xFF, 0xFB, 0xEB)
HDR_BG      = RGBColor(0x7B, 0x1C, 0x1C)
SEC_HDR_BG  = RGBColor(0xF0, 0xEE, 0xE8)

doc = Document()

# ── Page margins ─────────────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.2)
    section.bottom_margin = Cm(2.2)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.2)

# ── Helpers ───────────────────────────────────────────────────────────────────
def rgb_hex(rgb: RGBColor) -> str:
    return f'{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}'

def set_cell_bg(cell, rgb: RGBColor):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  rgb_hex(rgb))
    tcPr.append(shd)

def set_cell_borders(cell, sides=('top','bottom','left','right'), color='E2E4E8', sz=4):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBd = OxmlElement('w:tcBorders')
    for side in sides:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'),   'single')
        el.set(qn('w:sz'),    str(sz))
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), color)
        tcBd.append(el)
    tcPr.append(tcBd)

def para_spacing(para, before=0, after=60):
    pPr = para._p.get_or_add_pPr()
    spc = OxmlElement('w:spacing')
    spc.set(qn('w:before'), str(before))
    spc.set(qn('w:after'),  str(after))
    pPr.append(spc)

def add_run(para, text, bold=False, italic=False, color=None, size=None):
    run = para.add_run(text)
    run.bold   = bold
    run.italic = italic
    if color: run.font.color.rgb = color
    if size:  run.font.size      = Pt(size)
    return run

def heading1(doc, text):
    p = doc.add_paragraph()
    para_spacing(p, before=240, after=80)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.bold           = True
    run.font.size      = Pt(16)
    run.font.color.rgb = QH_MAROON
    # bottom border
    pPr  = p._p.get_or_add_pPr()
    pBd  = OxmlElement('w:pBdr')
    bot  = OxmlElement('w:bottom')
    bot.set(qn('w:val'),   'single')
    bot.set(qn('w:sz'),    '12')
    bot.set(qn('w:space'), '4')
    bot.set(qn('w:color'), '7B1C1C')
    pBd.append(bot)
    pPr.append(pBd)
    return p

def heading2(doc, text):
    p = doc.add_paragraph()
    para_spacing(p, before=160, after=60)
    run = p.add_run(text)
    run.bold      = True
    run.font.size = Pt(12)
    run.font.color.rgb = DARK
    return p

def body_para(doc, text, color=None):
    p   = doc.add_paragraph()
    para_spacing(p, before=0, after=60)
    run = p.add_run(text)
    run.font.size = Pt(10)
    run.font.color.rgb = color or MID_GREY
    return p

def callout_box(doc, label, body, bg=YELLOW_BG, label_color=YELLOW_BD, body_color=RGBColor(0x78,0x35,0x0F)):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    cell = tbl.rows[0].cells[0]
    set_cell_bg(cell, bg)
    # label
    lp = cell.add_paragraph()
    para_spacing(lp, before=60, after=30)
    lr = lp.add_run(label.upper())
    lr.bold = True; lr.font.size = Pt(8); lr.font.color.rgb = label_color
    # body
    bp = cell.add_paragraph()
    para_spacing(bp, before=0, after=60)
    br = bp.add_run(body)
    br.font.size = Pt(10); br.font.color.rgb = body_color
    # remove default empty para inside cell
    for p in list(cell.paragraphs):
        if not p.text and p._p != lp._p and p._p != bp._p:
            p._p.getparent().remove(p._p)
    doc.add_paragraph()  # spacer

def page_break(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_break(docx_break_type())

def docx_break_type():
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    return br   # we'll use it differently below

def add_page_break(doc):
    p  = doc.add_paragraph()
    run = p.add_run()
    br  = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    run._r.append(br)

def divider(doc):
    p   = doc.add_paragraph()
    para_spacing(p, before=80, after=80)
    pPr = p._p.get_or_add_pPr()
    pBd = OxmlElement('w:pBdr')
    bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'),   'single')
    bot.set(qn('w:sz'),    '4')
    bot.set(qn('w:space'), '1')
    bot.set(qn('w:color'), 'E2E4E8')
    pBd.append(bot)
    pPr.append(pBd)

# ── Script table builder ──────────────────────────────────────────────────────
def script_table(doc, rows_data):
    """
    rows_data: list of dicts with keys:
        type  : 'header' | 'section' | 'gary' | 'tricia' | 'gautam' | 'wow' | 'all'
        time  : str
        who   : str
        screen: str
        words : str
        notes : str
    """
    COL_WIDTHS = [Cm(1.5), Cm(2.0), Cm(5.5), Cm(8.0), Cm(2.5)]

    tbl = doc.add_table(rows=0, cols=5)
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl.style     = 'Table Grid'

    # header row
    hdr = tbl.add_row()
    for i, label in enumerate(['Time','Who','On Screen / Action','Spoken Words','Notes']):
        cell = hdr.cells[i]
        cell.width = COL_WIDTHS[i]
        set_cell_bg(cell, HDR_BG)
        p   = cell.paragraphs[0]
        run = p.add_run(label.upper())
        run.bold = True; run.font.size = Pt(8); run.font.color.rgb = WHITE
        para_spacing(p, before=40, after=40)

    ROW_COLORS = {
        'gary':   ROW_GARY,
        'tricia': ROW_TRICIA,
        'gautam': ROW_GAUTAM,
        'wow':    ROW_WOW,
        'all':    LIGHT_GREY,
    }

    for rd in rows_data:
        rtype = rd.get('type','gary')

        # section divider row
        if rtype == 'section':
            row = tbl.add_row()
            merged = row.cells[0].merge(row.cells[4])
            set_cell_bg(merged, SEC_HDR_BG)
            p   = merged.paragraphs[0]
            run = p.add_run(rd.get('label','').upper())
            run.bold = True; run.font.size = Pt(8.5); run.font.color.rgb = QH_MAROON
            para_spacing(p, before=40, after=40)
            continue

        row = tbl.add_row()
        bg  = ROW_COLORS.get(rtype, ROW_GARY)

        vals = [
            rd.get('time',''),
            rd.get('who',''),
            rd.get('screen',''),
            rd.get('words',''),
            rd.get('notes',''),
        ]

        for i, (cell, val) in enumerate(zip(row.cells, vals)):
            cell.width = COL_WIDTHS[i]
            set_cell_bg(cell, bg)
            p = cell.paragraphs[0]
            para_spacing(p, before=40, after=40)

            if i == 0:  # Time
                run = p.add_run(val)
                run.bold = True; run.font.size = Pt(9); run.font.color.rgb = QH_MAROON
            elif i == 1:  # Who
                run = p.add_run(val)
                run.bold = True; run.font.size = Pt(9)
                if rtype == 'wow':
                    run.font.color.rgb = RGBColor(0xB4, 0x53, 0x09)
            elif i == 2:  # Screen
                run = p.add_run(val)
                run.italic = True; run.font.size = Pt(9); run.font.color.rgb = MID_GREY
            elif i == 3:  # Words
                run = p.add_run(val)
                run.font.size = Pt(9.5)
                if rtype == 'gautam':
                    run.italic = True; run.font.color.rgb = MID_GREY
                if rtype == 'wow':
                    run.bold = True
            else:  # Notes
                run = p.add_run(val)
                run.font.size = Pt(8.5); run.italic = True; run.font.color.rgb = MID_GREY

    doc.add_paragraph()  # spacer after table

# ─────────────────────────────────────────────────────────────────────────────
# COVER PAGE
# ─────────────────────────────────────────────────────────────────────────────
p = doc.add_paragraph()
para_spacing(p, before=0, after=300)
run = p.add_run('QUEENSLAND HEALTH')
run.bold = True; run.font.size = Pt(11); run.font.color.rgb = QH_MAROON

p = doc.add_paragraph()
para_spacing(p, before=0, after=80)
run = p.add_run('DEMONSTRATION SCRIPT — CONFIDENTIAL')
run.bold = True; run.font.size = Pt(9); run.font.color.rgb = MID_GREY

p = doc.add_paragraph()
para_spacing(p, before=0, after=60)
run = p.add_run('Enterprise Chemical\nManagement System')
run.bold = True; run.font.size = Pt(28); run.font.color.rgb = DARK

p = doc.add_paragraph()
para_spacing(p, before=60, after=300)
run = p.add_run('Presenter guide and click-by-click script for the QH ECMS solution demonstration.')
run.font.size = Pt(13); run.font.color.rgb = MID_GREY

meta_tbl = doc.add_table(rows=2, cols=2)
meta_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
meta_data = [
    ('Duration', 'Approx. 30–45 minutes'),
    ('Format',   'Live interactive prototype'),
    ('Presenters','Gary · Tricia · Gautam'),
    ('Version',  'Draft 1 — May 2026'),
]
for idx, (label, value) in enumerate(meta_data):
    cell = meta_tbl.rows[idx // 2].cells[idx % 2]
    p    = cell.paragraphs[0]
    para_spacing(p, before=30, after=30)
    r1 = p.add_run(label + '\n')
    r1.bold = True; r1.font.size = Pt(9); r1.font.color.rgb = DARK
    r2 = p.add_run(value)
    r2.font.size = Pt(9); r2.font.color.rgb = MID_GREY

doc.add_paragraph()

p = doc.add_paragraph()
para_spacing(p, before=300, after=0)
run = p.add_run('COMMERCIAL IN CONFIDENCE — FOR AUTHORISED RECIPIENTS ONLY')
run.bold = True; run.font.size = Pt(9); run.font.color.rgb = QH_MAROON

add_page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# HOW TO READ THIS SCRIPT
# ─────────────────────────────────────────────────────────────────────────────
heading1(doc, 'How to Read This Script')

body_para(doc, (
    'This document is a presenter guide, not a verbatim transcript. The Spoken Words column captures the '
    'key messages each presenter must land. Deliver them in your own voice. The On Screen column describes '
    'what Gautam clicks or what appears on the prototype.'
))

# Legend table
leg_tbl = doc.add_table(rows=4, cols=2)
leg_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
legend_items = [
    (ROW_WOW,    'Yellow rows = Wow moment. Pause. Let it register.'),
    (ROW_GARY,   'Gary rows = Scene-setting, transitions, Q&A'),
    (ROW_TRICIA, 'Tricia rows = Primary narration as Gautam clicks'),
    (ROW_GAUTAM, 'Gautam rows = Click instruction only, no scripted words'),
]
for i, (color, label) in enumerate(legend_items):
    swatch_cell = leg_tbl.rows[i].cells[0]
    swatch_cell.width = Cm(1.2)
    set_cell_bg(swatch_cell, color)
    set_cell_borders(swatch_cell, color='CCCCCC')
    p = swatch_cell.paragraphs[0]
    para_spacing(p, before=30, after=30)

    label_cell = leg_tbl.rows[i].cells[1]
    p = label_cell.paragraphs[0]
    para_spacing(p, before=30, after=30)
    r = p.add_run(label)
    r.font.size = Pt(9); r.font.color.rgb = MID_GREY
doc.add_paragraph()

divider(doc)
heading2(doc, 'The Presenter Model')

presenter_tbl = doc.add_table(rows=1, cols=3)
presenter_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
presenters = [
    ('Gary',   'Senior Lead',      'Opens the session, introduces the team, sets up each block, lands key transitions, and handles all Q&A.'),
    ('Tricia', 'Primary Narrator', 'Narrates as Gautam clicks — what is happening on screen, why it matters, how it connects to QH operations. Primary voice for ~25–30 minutes.'),
    ('Gautam', 'Demo Driver',      'Drives the prototype. Knows every click, every screen, every data point. No scripted speaking role, but can add commentary where natural.'),
]
for i, (name, role, desc) in enumerate(presenters):
    cell = presenter_tbl.rows[0].cells[i]
    set_cell_bg(cell, LIGHT_GREY)
    set_cell_borders(cell, color='E2E4E8')
    p = cell.paragraphs[0]
    para_spacing(p, before=60, after=20)
    r1 = p.add_run(role.upper() + '\n')
    r1.bold = True; r1.font.size = Pt(8); r1.font.color.rgb = QH_MAROON
    r2 = p.add_run(name + '\n')
    r2.bold = True; r2.font.size = Pt(12)
    r3 = p.add_run(desc)
    r3.font.size = Pt(9); r3.font.color.rgb = MID_GREY
doc.add_paragraph()

divider(doc)
heading2(doc, 'Demo Flow — Four Blocks')

blocks_tbl = doc.add_table(rows=5, cols=5)
blocks_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
blocks_tbl.style = 'Table Grid'
block_hdrs = ['Block','Theme','What the audience sees','Time','Lead']
for i, h in enumerate(block_hdrs):
    cell = blocks_tbl.rows[0].cells[i]
    set_cell_bg(cell, HDR_BG)
    p = cell.paragraphs[0]; para_spacing(p, before=30, after=30)
    r = p.add_run(h.upper())
    r.bold = True; r.font.size = Pt(8); r.font.color.rgb = WHITE

block_rows = [
    ('A','The Scientist\'s World', 'Portal dashboard — a scientist\'s daily view of their lab\'s chemical landscape', '0:00 – 0:08', 'Gary'),
    ('B','Finding What You Need',  'Catalogue search — finding a chemical, reading its SDS, understanding its hazard profile', '0:08 – 0:16', 'Tricia'),
    ('C','Adding a Chemical Safely','Hazardous addition request — incompatibility detection, approval workflow, compliance guardrails', '0:16 – 0:28', 'Tricia'),
    ('D','Capability Summary & Q&A','Wrap-up, broader system capabilities, open questions', '0:28 – 0:35', 'Gary'),
]
for i, row_data in enumerate(block_rows):
    row = blocks_tbl.rows[i+1]
    bg = LIGHT_GREY if i % 2 == 0 else ROW_TRICIA
    for j, val in enumerate(row_data):
        cell = row.cells[j]
        set_cell_bg(cell, bg)
        p = cell.paragraphs[0]; para_spacing(p, before=30, after=30)
        r = p.add_run(val)
        r.font.size = Pt(9)
        if j == 0: r.bold = True; r.font.color.rgb = QH_MAROON
doc.add_paragraph()

divider(doc)
heading2(doc, 'The Three Wow Moments')

callout_box(doc,
    'Wow Moment 1 — First 60 seconds',
    'The dashboard loads and the audience sees Sarah\'s complete chemical world — live inventory, overdue training, pending approvals, recent activity — all in one screen. No hunting. No spreadsheets.',
    bg=RGBColor(0xFE,0xF2,0xF2), label_color=QH_MAROON, body_color=RGBColor(0x7F,0x1D,0x1D)
)
callout_box(doc,
    'Wow Moment 2 — Minute 12',
    'The catalogue finds Acetone in under two seconds. The system already knows it is High risk, DG Class 3, and shows SDS currency. The recall banner for an unrelated product fires automatically — the system is watching.',
    bg=RGBColor(0xFE,0xF2,0xF2), label_color=QH_MAROON, body_color=RGBColor(0x7F,0x1D,0x1D)
)
callout_box(doc,
    'Wow Moment 3 — Minute 20',
    'Sarah tries to add Acetone next to oxidisers. The system fires an incompatibility alert before she can submit — citing the exact AS/NZS 4452 rule being violated and offering three pre-calculated safe alternatives. Compliance built in, not bolted on.',
    bg=RGBColor(0xFE,0xF2,0xF2), label_color=QH_MAROON, body_color=RGBColor(0x7F,0x1D,0x1D)
)

add_page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# BLOCK A
# ─────────────────────────────────────────────────────────────────────────────
heading1(doc, 'Block A — The Scientist\'s World')
p = doc.add_paragraph(); para_spacing(p, before=0, after=80)
r = p.add_run('Duration: ~8 minutes  |  Lead: Gary opens, Tricia takes over  |  Screen: Portal Dashboard')
r.font.size = Pt(9); r.font.color.rgb = MID_GREY

callout_box(doc,
    'Strategic Purpose',
    'Establish the daily reality for a laboratory scientist managing hazardous chemicals. The audience should finish this block understanding who uses the system, what their daily pressures look like, and why a fragmented status quo is a compliance and safety risk.',
    bg=BLUE_BG, label_color=BLUE_BD, body_color=RGBColor(0x1E,0x3A,0x8A)
)

script_table(doc, [
    dict(type='gary',   time='0:00', who='Gary',   screen='Team visible. No screen share yet.',
         words='Good morning — thank you for having us. I\'m Gary, and I\'m going to be guiding you through today. With me is Tricia, who will walk you through the system scenario by scenario, and Gautam, our technical lead who is driving the demo. Let\'s get into it.',
         notes='Warm, brief. Do not over-introduce.'),
    dict(type='gary',   time='0:01', who='Gary',   screen='Still no screen share.',
         words='Before we touch the screen, I want to give you thirty seconds of context. Queensland Health manages thousands of hazardous chemicals across hundreds of laboratories. Right now, that management is fragmented — spreadsheets, paper registers, email approvals. ECMS changes that. What you\'re about to see is a single, integrated platform for every part of the chemical management lifecycle. Let\'s show you what that looks like in practice.',
         notes='Sets the problem before showing the solution.'),
    dict(type='gautam', time='0:02', who='Gautam', screen='Share screen. Navigate to the ECMS Portal. Dashboard lands on Dr Sarah Chen\'s view.',
         words='(no words — click)',
         notes='Let the dashboard fully render before Tricia speaks.'),
    dict(type='wow',    time='0:02', who='Tricia',  screen='Portal dashboard fully visible. PAUSE 3 seconds.',
         words='This is Sarah. She\'s a Senior Scientist at Royal Brisbane and Women\'s Hospital, Lab 312 — Wet Chemistry. This is what she sees the moment she logs in. Not a menu, not a list of modules — her world. Right now. ★ WOW MOMENT 1',
         notes='Pause after "her world." Let it land.'),
    dict(type='tricia', time='0:03', who='Tricia',  screen='Dashboard — left card: inventory panel.',
         words='On the left, she can see every chemical in Lab 312 right now. Five chemicals. Three are High risk. She knows where each one is stored — Oxidisers Locker B, Flammables Cabinet A — without opening a register or calling the lab manager.',
         notes=''),
    dict(type='tricia', time='0:04', who='Tricia',  screen='Dashboard — centre card: tasks panel.',
         words='In the middle, her tasks and approvals. Three items — one due tomorrow, one already overdue, one awaiting WHS sign-off. The system is not just storing data. It\'s actively managing her obligations. She doesn\'t need to remember. The system remembers for her.',
         notes='Emphasise "The system remembers."'),
    dict(type='tricia', time='0:05', who='Tricia',  screen='Dashboard — right card: training panel.',
         words='On the right — her training status. Two certifications current, one overdue. Hazardous Spills — flagged in red. The system surfaces this because it matters. In a lab environment, an untrained operator is a safety and compliance risk. ECMS makes that visible before it becomes an incident.',
         notes='Connect to real-world consequence.'),
    dict(type='tricia', time='0:06', who='Tricia',  screen='Dashboard — bottom row: recent activity + quick actions.',
         words='Below that, recent activity across her lab — who added what, who approved what, what the system flagged automatically. And on the right, her quick actions. Search SDS, add a product, scan a barcode, report an incident. Everything she needs is one click away.',
         notes=''),
    dict(type='gary',   time='0:07', who='Gary',   screen='Dashboard still showing.',
         words='What Tricia has just described is not a report. It is not a read-only view. Everything you see here is live and actionable. Sarah can approve a task, request a chemical, or escalate a training issue without leaving this screen. Let me show you what happens when she needs a new chemical.',
         notes='Transition to Block B.'),
])

add_page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# BLOCK B
# ─────────────────────────────────────────────────────────────────────────────
heading1(doc, 'Block B — Finding What You Need')
p = doc.add_paragraph(); para_spacing(p, before=0, after=80)
r = p.add_run('Duration: ~8 minutes  |  Lead: Tricia  |  Screen: Catalogue Search Results')
r.font.size = Pt(9); r.font.color.rgb = MID_GREY

callout_box(doc,
    'Strategic Purpose',
    'Show that the system knows its chemicals — not just as names in a list but as structured safety data. The catalogue is where regulatory intelligence lives. The recall banner demonstrates that the system is proactive, not just reactive.',
    bg=BLUE_BG, label_color=BLUE_BD, body_color=RGBColor(0x1E,0x3A,0x8A)
)

script_table(doc, [
    dict(type='tricia', time='0:08', who='Tricia',  screen='Dashboard — hover over Search catalogue button.',
         words='Sarah needs to add Acetone to her lab. The first step is finding it in the catalogue. She doesn\'t go to a separate system or open a browser tab. She searches right here.',
         notes=''),
    dict(type='gautam', time='0:08', who='Gautam', screen='Click "Search catalogue". Catalogue results page loads.',
         words='(no words — click)',
         notes='Let the results page fully render.'),
    dict(type='wow',    time='0:09', who='Tricia',  screen='Catalogue results. Acetone result visible. PAUSE 2 seconds.',
         words='One result. Instant. The system searched 51 chemicals on the QH register and cross-referenced 12,483 products in the master catalogue. She has her answer before she has finished reading the screen. ★ WOW MOMENT 2',
         notes='Pause after "finished reading the screen."'),
    dict(type='tricia', time='0:10', who='Tricia',  screen='Point to the product recall banner at top of results.',
         words='Notice what fires at the top of the page before she even looks at the result. An active product recall. Glutaraldehyde 2% — a completely different chemical — has an active recall notice. Sarah didn\'t ask for this. The system surfaced it because it\'s relevant to her lab. That is proactive compliance management.',
         notes='Emphasise that Sarah did not ask for this.'),
    dict(type='tricia', time='0:11', who='Tricia',  screen='Point to filter chips below the recall banner.',
         words='She can refine these results further — filter by hazard class, by chemicals held at her location, by GHS pictogram, by scheduled substance class. The system is already filtered to DG Class 3 — Flammable, which is Acetone\'s classification. Everything is contextual.',
         notes=''),
    dict(type='tricia', time='0:12', who='Tricia',  screen='Point to the Acetone result card — GHS icons, badges, danger statement, SDS date.',
         words='And here is Acetone. PRD-0001. CAS 67-64-1. Sigma-Aldrich Merck, 99.5% purity. The system shows her the GHS hazard pictograms alongside the danger statement: H225, highly flammable liquid and vapour. She can see at a glance that the SDS is current to May 2027. No chasing. No uncertainty.',
         notes='Walk through the card left to right.'),
    dict(type='gautam', time='0:13', who='Gautam', screen='Hover over the "View SDS" button on the Acetone card. Do not click.',
         words='(no words — hover only)',
         notes='Creates visual emphasis without navigating away.'),
    dict(type='tricia', time='0:13', who='Tricia',  screen='SDS button highlighted.',
         words='From here she can pull up the full Safety Data Sheet in one click — the current version, stored and managed within ECMS. She can also go straight to requesting this product for her inventory. That\'s what she\'s going to do next.',
         notes=''),
    dict(type='gary',   time='0:15', who='Gary',   screen='Catalogue results still showing.',
         words='This is the intelligence layer that underpins everything else. The system doesn\'t just hold chemical names. It holds hazard classifications, SDS versions, recall status, and regulatory obligations — and it surfaces the right information at the right moment. Let\'s follow Sarah as she tries to add this chemical to her lab, because this is where the system\'s safety logic really comes to life.',
         notes='Transition to Block C.'),
])

add_page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# BLOCK C
# ─────────────────────────────────────────────────────────────────────────────
heading1(doc, 'Block C — Adding a Chemical Safely')
p = doc.add_paragraph(); para_spacing(p, before=0, after=80)
r = p.add_run('Duration: ~12 minutes  |  Lead: Tricia  |  Screen: Request to Add Hazardous Product')
r.font.size = Pt(9); r.font.color.rgb = MID_GREY

callout_box(doc,
    'Strategic Purpose',
    'This is the centrepiece of the demo. It proves that ECMS does not just record what people do — it actively prevents unsafe decisions. The incompatibility detection is the strongest differentiator. Give this block time and space.',
    bg=BLUE_BG, label_color=BLUE_BD, body_color=RGBColor(0x1E,0x3A,0x8A)
)

script_table(doc, [
    dict(type='gautam', time='0:16', who='Gautam', screen='Click "Request to add" on the Acetone catalogue card. Request form loads.',
         words='(no words — click)',
         notes='Pause before Tricia speaks. Let the form and red alert both settle on screen.'),
    dict(type='wow',    time='0:16', who='Tricia',  screen='Request form. Incompatibility alert dominates the top. PAUSE 4 seconds.',
         words='Sarah has done everything right. She found the product, she is using the right form, she has selected Flammables Cabinet A — her lab\'s designated flammables storage. And the system has stopped her. Before she can submit. Before she can make a mistake. ★ WOW MOMENT 3',
         notes='Biggest pause in the demo. Let the red alert speak for itself.'),
    dict(type='tricia', time='0:17', who='Tricia',  screen='Point to incompatibility alert heading: "Incompatibility detected (BR008, BR009)".',
         words='The system has detected two incompatibility rules — BR008 and BR009. Flammables Cabinet A sits inside Lab 312, which already holds Hydrogen peroxide and Potassium permanganate. Both are oxidisers. Acetone is a flammable solvent — DG Class 3. The system knows that storing a flammable next to an oxidiser violates AS/NZS 4452 segregation requirements. It tells her exactly why — and it tells her before anything is stored, ordered, or approved.',
         notes='Pause on "before anything is stored."'),
    dict(type='tricia', time='0:18', who='Tricia',  screen='Point to the "Why this matters" box within the alert.',
         words='The system doesn\'t just flag the conflict. It explains it in plain language. Acetone and hydrogen peroxide stored together creates a credible fire and explosion hazard. The citation is there. The reason is there. A scientist who has never encountered this incompatibility now understands exactly why the system is intervening.',
         notes=''),
    dict(type='tricia', time='0:19', who='Tricia',  screen='Point to three action buttons: Choose compatible location, Move oxidiser instead, Override.',
         words='And the system gives her three paths forward. She can choose a compatible location — the system has already calculated which ones work. She can move the oxidiser instead. Or she can request an override with WHS sign-off. The system is not blocking her. It is ensuring that whatever she does next is a deliberate, documented decision.',
         notes='Emphasise "deliberate, documented decision."'),
    dict(type='tricia', time='0:20', who='Tricia',  screen='Point to the blue "Suggested compatible locations" panel.',
         words='Below the alert, the system has already done the work. Three compatible locations within Royal Brisbane and Women\'s Hospital — ranked by proximity and suitability. Lab 312\'s Flammables Cabinet A in the same room. The Bulk Chemical Store two floors down. A prep area across the courtyard. Distance, classification, and current holdings — all surfaced automatically.',
         notes=''),
    dict(type='tricia', time='0:21', who='Tricia',  screen='Scroll down to the form fields.',
         words='When she\'s resolved the incompatibility, the rest of the form captures everything the system needs — quantity, batch number, expiry date, custodian, reason for addition. This is the data that drives the audit trail. Every field has a purpose. Nothing is captured for its own sake.',
         notes=''),
    dict(type='tricia', time='0:22', who='Tricia',  screen='Point to the right-hand sidebar — Approval Workflow, SDS Check, Audit History.',
         words='On the right, the approval workflow is already assembled. The system knows this request goes to Susan Mannerheim — the WHS Advisor for Metro North HHS. Two business day SLA. If Susan doesn\'t action it, it escalates automatically to Aisha Rahman after day three. No chasing. No dropped requests. The workflow runs itself.',
         notes=''),
    dict(type='tricia', time='0:23', who='Tricia',  screen='Point to SDS Check and Audit History in the sidebar.',
         words='Below that, SDS currency — reviewed and current. And the audit trail — every action, every timestamp, every system check already recorded. Sarah started the request at 08:18. The incompatibility check ran at 08:32. This is not added to the audit trail at the end. It is written as it happens.',
         notes=''),
    dict(type='tricia', time='0:24', who='Tricia',  screen='Point to the greyed-out "Submit for approval" button and the note below the form.',
         words='And notice — the submit button is disabled. Sarah cannot submit while the incompatibility is active. This is not optional compliance. The system enforces it. The only way to submit is to resolve the conflict. That resolution is recorded. That decision is owned. That is how a safety management system should work.',
         notes='End with conviction. This closes the centrepiece.'),
    dict(type='gary',   time='0:26', who='Gary',   screen='Request form still showing.',
         words='What Tricia has walked you through is the compliance layer working exactly as it should. The system knows the rules. It applies them before any human can make a mistake. And it keeps a complete, tamper-evident record of every decision. That is not a feature. That is the foundation of a system that Queensland Health can trust at scale.',
         notes='Deliberate, measured. Gary takes back the room.'),
])

add_page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# BLOCK D
# ─────────────────────────────────────────────────────────────────────────────
heading1(doc, 'Block D — Capability Summary & Q&A')
p = doc.add_paragraph(); para_spacing(p, before=0, after=80)
r = p.add_run('Duration: ~7 minutes  |  Lead: Gary  |  Screen: Dashboard or blank')
r.font.size = Pt(9); r.font.color.rgb = MID_GREY

callout_box(doc,
    'Strategic Purpose',
    'Land the broader system story — this is not just three screens, it is an enterprise platform. Then open the floor for questions. Gary owns this block entirely.',
    bg=BLUE_BG, label_color=BLUE_BD, body_color=RGBColor(0x1E,0x3A,0x8A)
)

script_table(doc, [
    dict(type='gary', time='0:28', who='Gary', screen='Navigate back to the Portal dashboard.',
         words='Let me bring you back to where we started. Sarah\'s dashboard. Three screens in, and here is what you have seen: a scientist who knows the status of her entire lab at a glance. A catalogue that surfaces the right safety information without being asked. And a system that catches a dangerous storage decision before it can be made. That is the ECMS experience.',
         notes='Reflective tone. Summarise without repeating.'),
    dict(type='gary', time='0:29', who='Gary', screen='Dashboard visible. No clicking required.',
         words='What you have not seen today — and what we are ready to show — is the WHS Advisor\'s approval queue. The inventory management workflow. Risk assessment and SDS version control. Disposal tracking. Reporting dashboards. Barcode scanning for physical stock takes. Incident reporting. Every one of those capabilities is built and ready. Today we showed you the core user journey. The full platform extends well beyond it.',
         notes='Creates appetite for follow-up.'),
    dict(type='gary', time='0:31', who='Gary', screen='Screen can be minimised or left on dashboard.',
         words='Before I open to questions, three things I want you to take away. First — the system enforces compliance before mistakes happen, not after. Second — every action is auditable, timestamped, and owned. Third — the scientist, the WHS advisor, and the lab manager all have exactly the information they need, at the moment they need it, without looking for it. That is what good system design looks like in a safety-critical environment.',
         notes='Deliver the three points slowly. Pause between each.'),
    dict(type='gary', time='0:33', who='Gary', screen='No screen action needed.',
         words='We would love to hear your questions. Gautam and I are both here to go deeper on any part of what you have seen — the workflow logic, the compliance rule engine, the integration approach, or the data model underneath. What would you like to explore?',
         notes='Open, confident. Invite questions, do not deflect.'),
    dict(type='all',  time='0:34+', who='All',  screen='Gautam ready to navigate to any screen on request.',
         words='(Q&A — respond in your own words. Gautam navigates to relevant screens as questions arise.)',
         notes='Gary: commercial/architecture. Tricia: workflow/user. Gautam: navigates on demand.'),
])

divider(doc)
heading2(doc, 'Anticipated Questions')

qa_tbl = doc.add_table(rows=7, cols=3)
qa_tbl.style = 'Table Grid'
qa_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT

qa_hdrs = ['If they ask…','Key points to hit','Who answers']
for i, h in enumerate(qa_hdrs):
    cell = qa_tbl.rows[0].cells[i]
    set_cell_bg(cell, HDR_BG)
    p = cell.paragraphs[0]; para_spacing(p, before=30, after=30)
    r = p.add_run(h.upper())
    r.bold = True; r.font.size = Pt(8); r.font.color.rgb = WHITE

qa_data = [
    ('How does the incompatibility rule engine work?',
     'Rules are configurable by WHS policy owners, not hard-coded. Each rule cites the regulatory reference (AS/NZS 4452, WHS Regulations). New rules can be added without development effort.',
     'Gary + Gautam'),
    ('How does SDS management work?',
     'SDS documents are stored, versioned, and linked to products. Expiry alerts are automated. The system tracks which version was current at the time of each action — important for post-incident review.',
     'Tricia'),
    ('Can different hospitals/labs have different rules?',
     'Yes. The system supports location-specific rule sets and approval hierarchies. Metro North HHS and Gold Coast Health can have different WHS Advisor chains without any configuration conflict.',
     'Gary'),
    ('How does the approval workflow handle leave or absence?',
     'Delegates can be configured per approver. The system also has automatic escalation on SLA breach — as you saw with Susan Mannerheim escalating to Aisha Rahman after 3 days.',
     'Tricia'),
    ('What if someone tries to override the incompatibility?',
     'Overrides require WHS sign-off — they cannot be self-approved. The override request, the justification, and the approver are all written to the audit trail. The system permits it but ensures accountability.',
     'Gary'),
    ('How does barcode/physical stock management work?',
     'Each chemical item has a barcode generated on addition. Mobile-compatible scanning updates inventory in real time. Stock takes can be run from a tablet on the lab floor without a laptop.',
     'Gautam'),
]

for i, (q, a, who) in enumerate(qa_data):
    row = qa_tbl.rows[i+1]
    bg = LIGHT_GREY if i % 2 == 0 else ROW_TRICIA
    for j, val in enumerate([q, a, who]):
        cell = row.cells[j]
        set_cell_bg(cell, bg)
        p = cell.paragraphs[0]; para_spacing(p, before=30, after=30)
        r = p.add_run(val)
        r.font.size = Pt(9)
        if j == 0: r.bold = True

doc.add_paragraph()
divider(doc)
heading2(doc, 'Closing Line — If Asked for a Single Summary Statement')

callout_box(doc,
    'Gary',
    '"ECMS does not just manage chemicals. It manages the decisions people make about chemicals — and it makes sure those decisions are safe, compliant, and owned. That is the difference between a records system and a safety management system."',
    bg=RGBColor(0xFE,0xF2,0xF2), label_color=QH_MAROON, body_color=RGBColor(0x7F,0x1D,0x1D)
)

# ─────────────────────────────────────────────────────────────────────────────
# Save
# ─────────────────────────────────────────────────────────────────────────────
out = '/home/user/QH-ECMS-Demonstration/QH-ECMS-Demo-Script.docx'
doc.save(out)
print(f'Saved: {out}')
