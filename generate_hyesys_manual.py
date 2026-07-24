from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

OUTPUT_PATH = r"C:\Users\JasonOng\OneDrive - Advancer Global Ltd\AST BD\HyESys Dept\6. Documentation\HyESys_Client_User_Manual.docx"

AST_BLUE  = RGBColor(0x00, 0x47, 0xAB)   # corporate blue
DARK_GREY = RGBColor(0x40, 0x40, 0x40)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_BG  = RGBColor(0xF2, 0xF7, 0xFF)   # very light blue for alt rows

doc = Document()

# ── Page margins ─────────────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

# ── Helper: set cell background colour ───────────────────────────────────────
def set_cell_bg(cell, hex_colour):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  hex_colour)
    tcPr.append(shd)

def set_cell_borders(table, border_size=4):
    """Light inner borders on all cells."""
    for row in table.rows:
        for cell in row.cells:
            tc   = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')
            for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
                border = OxmlElement(f'w:{side}')
                border.set(qn('w:val'),  'single')
                border.set(qn('w:sz'),   str(border_size))
                border.set(qn('w:space'),'0')
                border.set(qn('w:color'),'BFCFE8')
                tcBorders.append(border)
            tcPr.append(tcBorders)

def style_para(para, size=10, bold=False, colour=None, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=4):
    para.alignment = align
    para.paragraph_format.space_after = Pt(space_after)
    for run in para.runs:
        run.font.size  = Pt(size)
        run.font.bold  = bold
        if colour:
            run.font.color.rgb = colour

def add_heading(doc, text, level=1):
    """Section heading with AST blue rule."""
    if level == 1:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after  = Pt(4)
        run = p.add_run(text)
        run.font.size  = Pt(13)
        run.font.bold  = True
        run.font.color.rgb = AST_BLUE
        # bottom border = rule line
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'),   'single')
        bottom.set(qn('w:sz'),    '6')
        bottom.set(qn('w:space'), '4')
        bottom.set(qn('w:color'), '0047AB')
        pBdr.append(bottom)
        pPr.append(pBdr)
    else:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after  = Pt(2)
        run = p.add_run(text)
        run.font.size  = Pt(11)
        run.font.bold  = True
        run.font.color.rgb = DARK_GREY
    return p

def add_body(doc, text, size=10, space_after=4):
    p = doc.add_paragraph(text)
    p.paragraph_format.space_after = Pt(space_after)
    for run in p.runs:
        run.font.size = Pt(size)
    return p

def add_bullet(doc, text, size=10):
    p = doc.add_paragraph(text, style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    for run in p.runs:
        run.font.size = Pt(size)
    return p

def add_numbered(doc, text, size=10):
    p = doc.add_paragraph(text, style='List Number')
    p.paragraph_format.space_after = Pt(2)
    for run in p.runs:
        run.font.size = Pt(size)
    return p

def header_row(table, *texts, col_widths=None):
    """Style the first row as a blue header."""
    row = table.rows[0]
    for i, cell in enumerate(row.cells):
        set_cell_bg(cell, '0047AB')
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after  = Pt(2)
        run = p.add_run(texts[i] if i < len(texts) else '')
        run.font.size  = Pt(9)
        run.font.bold  = True
        run.font.color.rgb = WHITE

def body_cell(cell, text, bold=False, centre=False, size=9):
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if centre else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold

def add_note(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Cm(0.5)
    run = p.add_run('⚠  NOTE: ')
    run.font.bold = True
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0xCC, 0x66, 0x00)
    run2 = p.add_run(text)
    run2.font.size = Pt(9)
    run2.font.italic = True

def add_warning(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Cm(0.5)
    run = p.add_run('⛔  WARNING: ')
    run.font.bold = True
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)
    run2 = p.add_run(text)
    run2.font.size = Pt(9)
    run2.font.bold = True

# ═══════════════════════════════════════════════════════════════════════════════
#  COVER PAGE
# ═══════════════════════════════════════════════════════════════════════════════
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(60)
p.paragraph_format.space_after  = Pt(6)
run = p.add_run('HyESys')
run.font.size  = Pt(36)
run.font.bold  = True
run.font.color.rgb = AST_BLUE

p2 = doc.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p2.paragraph_format.space_after = Pt(4)
run2 = p2.add_run('Active Digital Power Compensator')
run2.font.size = Pt(14)
run2.font.color.rgb = DARK_GREY

p3 = doc.add_paragraph()
p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
p3.paragraph_format.space_after = Pt(60)
run3 = p3.add_run('Client User Manual')
run3.font.size = Pt(18)
run3.font.bold = True
run3.font.color.rgb = DARK_GREY

# Cover info table
cover_table = doc.add_table(rows=7, cols=2)
cover_table.style = 'Table Grid'
cover_table.alignment = WD_TABLE_ALIGNMENT.CENTER
fields = [
    ('Project Name',       '[Project Name]'),
    ('Installation Site',  '[Site Address]'),
    ('HyESys Model',       '[H30 / H50 / H125]'),
    ('Serial Number',      '[Serial Number]'),
    ('Document Revision',  'Rev 1.0'),
    ('Issue Date',         '[DD MMM YYYY]'),
    ('Prepared By',        'Advancer Smart Technology Pte Ltd'),
]
for i, (label, value) in enumerate(fields):
    label_cell = cover_table.rows[i].cells[0]
    value_cell = cover_table.rows[i].cells[1]
    set_cell_bg(label_cell, 'E8EFF8')
    body_cell(label_cell, label, bold=True, size=10)
    body_cell(value_cell, value, size=10)
set_cell_borders(cover_table)

doc.add_paragraph()
p_ast = doc.add_paragraph()
p_ast.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_ast = p_ast.add_run('Advancer Smart Technology Pte Ltd  |  www.advancer.sg')
run_ast.font.size  = Pt(9)
run_ast.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
#  SECTION 1 — DOCUMENT CONTROL
# ═══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '1.  Document Control')
add_body(doc, 'This document is issued by Advancer Smart Technology Pte Ltd (AST) for the client\'s reference and ongoing operation of the HyESys unit installed at the above site.')

t = doc.add_table(rows=9, cols=2)
t.style = 'Table Grid'
rows_data = [
    ('Project Name',       '[Project Name]'),
    ('Installation Site',  '[Full site address]'),
    ('HyESys Model',       '[H30 / H50 / H125]'),
    ('Unit Serial Number', '[SN-XXXXXX]'),
    ('MSB / Panel',        '[Panel designation]'),
    ('Commissioning Date', '[DD MMM YYYY]'),
    ('Document Revision',  'Rev 1.0'),
    ('Issue Date',         '[DD MMM YYYY]'),
    ('Approved By',        '[Name, Designation]'),
]
for i, (label, value) in enumerate(rows_data):
    set_cell_bg(t.rows[i].cells[0], 'E8EFF8')
    body_cell(t.rows[i].cells[0], label, bold=True)
    body_cell(t.rows[i].cells[1], value)
set_cell_borders(t)

doc.add_paragraph()

# Revision history
add_heading(doc, 'Revision History', level=2)
rh = doc.add_table(rows=2, cols=4)
rh.style = 'Table Grid'
header_row(rh, 'Rev', 'Date', 'Description', 'Approved By')
data = [('1.0', '[DD MMM YYYY]', 'Initial issue', '[Name]')]
for i, row_data in enumerate(data):
    for j, val in enumerate(row_data):
        body_cell(rh.rows[i+1].cells[j], val, centre=(j==0))
set_cell_borders(rh)

doc.add_paragraph()

# ═══════════════════════════════════════════════════════════════════════════════
#  SECTION 2 — SYSTEM OVERVIEW
# ═══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '2.  System Overview')

add_heading(doc, '2.1  Purpose', level=2)
add_body(doc, 'HyESys is an active digital power compensator installed at the Main Switchboard (MSB) incomer. It simultaneously delivers three functions:')
add_bullet(doc, 'Reactive power compensation — corrects power factor by injecting leading or lagging kVAr, reducing reactive current drawn from the grid.')
add_bullet(doc, 'Three-phase load balancing — redistributes current across phases to eliminate neutral I²R losses.')
add_bullet(doc, 'Energy storage / solar load shaving — stores excess solar generation and releases it to shave peak demand.')
add_body(doc, 'These three functions share the unit\'s rated kVA capacity and cannot all operate at 100% simultaneously. The onboard controller allocates capacity in real time based on site conditions.')

doc.add_paragraph()
add_heading(doc, '2.2  Operating Principle', level=2)
add_body(doc, 'HyESys measures the current waveform at the MSB incomer at high speed. It calculates the reactive, imbalance and harmonic components and injects the exact compensating current through a dedicated circuit breaker connected to the MSB busbars. The result is a cleaner, balanced current seen by the utility meter, reducing kW losses and improving power factor.')

doc.add_paragraph()
add_heading(doc, '2.3  System Capacity', level=2)
cap = doc.add_table(rows=6, cols=6)
cap.style = 'Table Grid'
header_row(cap, 'Model', 'Rated Output (kVA)', 'Max Current (A)', 'Usable Energy (kWh)', 'Weight (kg)', 'Price (SGD)')
cap_data = [
    ('H30',  '30',  '43.5', '69.3',  '1,400', '$100,000'),
    ('H50',  '50',  '72.5', '108.9', '2,200', '$120,000'),
    ('H60',  '60',  '87',   '138.6', '2,800', 'TBD'),
    ('H100', '100', '145',  '217.8', '4,400', 'TBD'),
    ('H125', '125', '181',  '217.8', '4,400', '$100,000'),
]
for i, row in enumerate(cap_data):
    for j, val in enumerate(row):
        body_cell(cap.rows[i+1].cells[j], val, centre=True)
set_cell_borders(cap)

doc.add_paragraph()
add_heading(doc, '2.4  Functional Block Diagram', level=2)
add_body(doc, '[Insert functional block diagram — showing MSB busbars, dedicated circuit breaker, HyESys unit, DC battery string, monitoring gateway, data logger, router and cloud/SCADA connection.]')
add_note(doc, 'Target power factor is 0.98. Unity (1.0) is not the design target due to the law of convergence. The SP penalty threshold is PF < 0.85.')

doc.add_paragraph()

# ═══════════════════════════════════════════════════════════════════════════════
#  SECTION 3 — COMPONENTS HANDED OVER
# ═══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '3.  Components Handed Over to Client')
add_body(doc, 'The following equipment is supplied and handed over to the client upon commissioning. All items should be verified against this list during the handover inspection.')

comp = doc.add_table(rows=9, cols=6)
comp.style = 'Table Grid'
header_row(comp, 'Item', 'Description', 'Qty', 'Serial / Ref No.', 'Location', 'Warranty')
comp_data = [
    ('1', 'HyESys Main Unit',             '1', '[SN-XXXXXX]',   '[MSB Room / Panel Location]', '2 years'),
    ('2', 'HySBatt Battery Pack(s)',       '[x]','[SN-XXXXXX]', '[MSB Room]',                  '2 years'),
    ('3', 'Monitoring Gateway',            '1', '[SN-XXXXXX]',   '[Panel / DIN Rail]',          '1 year'),
    ('4', 'Data Logger (Splitter)',        '1', '[SN-XXXXXX]',   '[Panel / DIN Rail]',          '1 year'),
    ('5', 'Router / SIM Module',           '1', '[SN-XXXXXX]',   '[Panel / Cabinet]',           '1 year'),
    ('6', 'Communication & Power Cables',  '1 set', 'N/A',       'As installed',                'N/A'),
    ('7', 'Keys & Access Items',           '[x]','[Key No.]',   '[Refer to KT]',               'N/A'),
    ('8', 'Manufacturer Docs & Certs',     '1 set','N/A',        'With this manual / digital',  'N/A'),
]
for i, row in enumerate(comp_data):
    for j, val in enumerate(row):
        body_cell(comp.rows[i+1].cells[j], val, centre=(j in [0,2]))
set_cell_borders(comp)

add_note(doc, 'Key handover details to be completed by KT. Manufacturer datasheet, factory acceptance test (FAT) certificates and warranty cards are included in Appendix F.')

doc.add_paragraph()

# ═══════════════════════════════════════════════════════════════════════════════
#  SECTION 4 — KEY WIRING REQUIREMENTS
# ═══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '4.  Key Wiring Requirements')
add_body(doc, 'The following provides a generic overview of the wiring arrangement. All final connections must follow the approved Site Single Line Diagram (SLD) and installation drawings provided in Appendix A.')

add_heading(doc, '4.1  Generic Wiring Overview', level=2)
wiring_items = [
    ('Three-Phase Power Connection',        'L1, L2, L3 connected to HyESys AC terminals via the dedicated circuit breaker. Connections must be correctly phased and torqued to specification.'),
    ('Neutral Connection',                  'Neutral (N) connected to the HyESys neutral terminal. Do not share the neutral with other loads downstream of the breaker.'),
    ('Protective Earth (PE)',               'Earth connection made directly to the MSB earth busbar. A separate dedicated earth conductor is required — do not daisy-chain with other equipment.'),
    ('Dedicated Circuit Breaker',           'HyESys must be connected through a correctly rated dedicated MCCB or MCB as specified in the SLD. (Refer to Jeff for final breaker rating.)'),
    ('Isolation Means',                     'A lockable isolation switch upstream of the dedicated breaker is required to allow safe maintenance isolation. (Jeff to advise on final arrangement.)'),
    ('Communication Wiring',               'CT sensor cables, RS-485/Modbus wiring and Ethernet/SIM connections are routed as shown in the installation drawings. Do not re-route without AST approval.'),
]
wt = doc.add_table(rows=len(wiring_items)+1, cols=2)
wt.style = 'Table Grid'
header_row(wt, 'Connection', 'Requirement')
for i, (label, desc) in enumerate(wiring_items):
    body_cell(wt.rows[i+1].cells[0], label, bold=True)
    body_cell(wt.rows[i+1].cells[1], desc)
set_cell_borders(wt)

doc.add_paragraph()
add_heading(doc, '4.2  Indicative Cable and Breaker Ratings', level=2)
add_body(doc, '[To be confirmed by AST/Jeff based on approved SLD. Final ratings depend on cable run length, ambient temperature and local authority requirements.]')
add_warning(doc, 'No wiring modifications are to be carried out without AST written approval. Incorrect wiring will void warranty and may create electrical hazards.')

doc.add_paragraph()

# ═══════════════════════════════════════════════════════════════════════════════
#  SECTION 5 — POWERING UP SEQUENCE
# ═══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '5.  Powering Up Sequence')
add_note(doc, 'This section covers the full power-up from a de-energised state. For routine daily operation, HyESys starts and stops automatically — no action is required from the client.')

add_heading(doc, '5.1  Power-Up from Cold', level=2)
startup = [
    'Confirm all maintenance work is complete and all personnel are clear of the equipment.',
    'Verify the dedicated circuit breaker is in the OFF position.',
    'Close the upstream isolation switch.',
    'Switch on the auxiliary supply (if a separate auxiliary breaker is provided).',
    'Confirm the HyESys display panel illuminates and shows the startup screen.',
    'Close the dedicated circuit breaker.',
    'Confirm the unit enters STANDBY or AUTO mode on the display.',
    'Verify that power factor and current readings appear normal on the monitoring interface.',
]
for step in startup:
    add_numbered(doc, step)

doc.add_paragraph()
add_heading(doc, '5.2  Safe Isolation for Maintenance (De-energise)', level=2)
shutdown = [
    'Navigate to the HyESys control panel and select MANUAL HOLD or disable automatic operation.',
    'Wait for the unit to ramp down — confirm output current reaches zero on the display.',
    'Open (turn OFF) the dedicated circuit breaker.',
    'Switch off the auxiliary supply.',
    'Apply lockout/tagout on the isolation switch as per your site safety procedure.',
    'Wait at least 5 minutes for DC bus capacitors to discharge before opening the enclosure.',
    'Verify zero voltage with an approved voltage tester before commencing work.',
]
for step in shutdown:
    add_numbered(doc, step)

add_warning(doc, 'Never open the HyESys enclosure without first confirming zero voltage. Lethal DC bus voltages may be present even after AC isolation.')

doc.add_paragraph()

# ═══════════════════════════════════════════════════════════════════════════════
#  SECTION 6 — NORMAL OPERATION
# ═══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '6.  Normal Operation')

add_heading(doc, '6.1  Operating Modes', level=2)
modes = doc.add_table(rows=5, cols=3)
modes.style = 'Table Grid'
header_row(modes, 'Mode', 'Display Indicator', 'Description')
modes_data = [
    ('AUTO',   'Green — AUTO',    'Unit is operating normally. Reactive compensation, balancing and storage functions are active based on site demand.'),
    ('STANDBY','Amber — STANDBY', 'Unit is energised and ready but not actively injecting. Typically seen during low-load or overnight periods.'),
    ('FAULT',  'Red — FAULT',     'A fault condition has been detected. Unit has safely tripped. Check alarm log and contact AST if fault cannot be cleared.'),
    ('MANUAL', 'Blue — MANUAL',   'Unit is under manual control (for commissioning or maintenance only). Not a normal operating mode.'),
]
for i, row in enumerate(modes_data):
    for j, val in enumerate(row):
        body_cell(modes.rows[i+1].cells[j], val, centre=(j==0))
set_cell_borders(modes)

doc.add_paragraph()
add_heading(doc, '6.2  Expected Readings (Normal Conditions)', level=2)
readings = doc.add_table(rows=5, cols=3)
readings.style = 'Table Grid'
header_row(readings, 'Parameter', 'Normal Range', 'Note')
readings_data = [
    ('Power Factor (PF)',    '≥ 0.95 lagging',  'Target is 0.98. Values above 0.85 avoid SP penalty tariff.'),
    ('Phase Current Balance','< 5% imbalance',  'Significant reduction from pre-installation baseline expected.'),
    ('Output Current',       '≤ Rated unit current', 'Must not exceed nameplate rating continuously.'),
    ('DC Bus Voltage',       'Within model range',   'H30: 210–850 V; H50: 350–850 V; H125: 680–900 V.'),
]
for i, row in enumerate(readings_data):
    for j, val in enumerate(row):
        body_cell(readings.rows[i+1].cells[j], val, centre=(j==1))
set_cell_borders(readings)

doc.add_paragraph()
add_heading(doc, '6.3  Confirming Correct Operation', level=2)
add_bullet(doc, 'The monitoring dashboard (accessible via browser or app) should show real-time PF trending ≥ 0.95.')
add_bullet(doc, 'Monthly energy savings reports are generated automatically by the AST cloud platform.')
add_bullet(doc, 'If the display is blank or shows an unexpected state, note the error code and contact AST before taking any action.')

doc.add_paragraph()

# ═══════════════════════════════════════════════════════════════════════════════
#  SECTION 7 — ALARMS AND BASIC TROUBLESHOOTING
# ═══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '7.  Alarms and Basic Troubleshooting')
add_note(doc, 'The client is NOT expected to repair the unit. The actions below are limited to safe observation and basic resets only.')

alarms = doc.add_table(rows=9, cols=4)
alarms.style = 'Table Grid'
header_row(alarms, 'Alarm / Symptom', 'Possible Cause', 'Client Action', 'Contact AST?')
alarms_data = [
    ('FAULT — Overcurrent',           'Load surge or short circuit downstream',       'Check for obvious electrical faults. Do not attempt to reset without AST guidance.',                     'Yes'),
    ('FAULT — Overvoltage',           'Grid voltage spike or improper connection',     'Record the fault code and time. Do not reset. Contact AST.',                                             'Yes'),
    ('FAULT — Overtemperature',       'Poor ventilation or ambient temp too high',     'Check that ventilation openings are unobstructed. Allow unit to cool. Attempt one reset.',               'Yes if fault persists'),
    ('FAULT — Communication Loss',    'Router/SIM issue, cable fault',                 'Check router power and SIM signal. Restart router. If unresolved, contact AST.',                         'Yes if unresolved'),
    ('STANDBY — extended (>24 hrs)',  'Low load, or unit held in manual mode',         'Check operating mode on display. Confirm AUTO mode is selected.',                                        'Yes if AUTO shown but not operating'),
    ('Display blank',                 'Auxiliary supply off, internal fault',           'Check auxiliary circuit breaker. If tripped, reset once. Contact AST if blank persists.',               'Yes if persists'),
    ('PF not improving (dashboard)',  'Unit in STANDBY, or PF already at target',      'Verify unit is in AUTO mode. Review dashboard trend. PF may already be within 0.95–0.98 range.',       'Yes if PF consistently < 0.90'),
    ('Unusual noise or smell',        'Mechanical or electrical fault',                 'Immediately isolate the unit using the dedicated breaker. Do not re-energise. Contact AST urgently.',    'Yes — immediately'),
]
for i, row in enumerate(alarms_data):
    for j, val in enumerate(row):
        contact_yes = (j == 3 and 'Yes' in val)
        body_cell(alarms.rows[i+1].cells[j], val, bold=contact_yes)
set_cell_borders(alarms)

doc.add_paragraph()

# ═══════════════════════════════════════════════════════════════════════════════
#  SECTION 8 — DOS AND DON'TS
# ═══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '8.  Dos and Don\'ts')

dandt = doc.add_table(rows=2, cols=2)
dandt.style = 'Table Grid'
header_row(dandt, '✔  Do', '✘  Do Not')

dos = [
    'Keep ventilation openings unobstructed at all times.',
    'Conduct a visual inspection monthly — check for dust, water ingress, loose labels.',
    'Verify the monitoring dashboard is accessible and data is updating weekly.',
    'Report any unusual alarms, smells or sounds to AST immediately.',
    'Keep the area around the unit clear of combustible materials.',
    'Ensure the dedicated circuit breaker is accessible and labelled.',
    'Follow the isolation procedure before any maintenance near the unit.',
    'Retain all documentation in a safe and accessible location.',
]
donts = [
    'Do not open the enclosure unless authorised and trained to do so.',
    'Do not modify or re-route any wiring without AST written approval.',
    'Do not stack items against or on top of the unit.',
    'Do not attempt to reset more than once without AST guidance.',
    'Do not allow unauthorised personnel to operate the unit controls.',
    'Do not apply water or liquid cleaners to or near the unit.',
    'Do not connect additional loads to the dedicated circuit breaker.',
    'Do not disable alarms or override fault conditions independently.',
]

do_cell = dandt.rows[1].cells[0]
dont_cell = dandt.rows[1].cells[1]

for item in dos:
    p = do_cell.add_paragraph(item, style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    for run in p.runs:
        run.font.size = Pt(9)

for item in donts:
    p = dont_cell.add_paragraph(item, style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    for run in p.runs:
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)

set_cell_borders(dandt)
set_cell_bg(dandt.rows[0].cells[0], '0047AB')
set_cell_bg(dandt.rows[0].cells[1], '0047AB')

doc.add_paragraph()
add_heading(doc, '8.1  Inspection and Maintenance Schedule', level=2)
maint = doc.add_table(rows=5, cols=3)
maint.style = 'Table Grid'
header_row(maint, 'Frequency', 'Activity', 'Responsible')
maint_data = [
    ('Monthly',   'Visual inspection — ventilation, cleanliness, labels, indicator status', 'Client Facility Manager'),
    ('Quarterly', 'Check all cable terminations are secure and undamaged (external visual only)', 'Client Facility Manager'),
    ('Bi-annually','Full inspection and functional test', 'AST Service Engineer'),
    ('Annually',   'Preventive maintenance, firmware check, battery health check', 'AST Service Engineer'),
]
for i, row in enumerate(maint_data):
    for j, val in enumerate(row):
        body_cell(maint.rows[i+1].cells[j], val)
set_cell_borders(maint)

doc.add_paragraph()

# ═══════════════════════════════════════════════════════════════════════════════
#  SECTION 9 — SAFETY AND EMERGENCY RESPONSE
# ═══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '9.  Safety and Emergency Response')

add_heading(doc, '9.1  Electrical Hazards', level=2)
add_bullet(doc, 'HyESys operates at 400 V AC (three-phase) and high DC bus voltages (up to 900 V DC depending on model). Contact with live terminals is potentially fatal.')
add_bullet(doc, 'DC bus capacitors retain charge for up to 5 minutes after AC isolation. Always allow discharge time and verify with a live voltage indicator.')
add_bullet(doc, 'Do not assume the unit is safe because the display is off — internal voltages may still be present.')

doc.add_paragraph()
add_heading(doc, '9.2  Required PPE', level=2)
ppe = ['Insulated safety gloves (minimum Cat III / 1000 V rated)', 'Arc flash-rated face shield and clothing (appropriate to MSB arc flash level)', 'Non-conductive safety footwear', 'Voltage detection / live line indicator']
for item in ppe:
    add_bullet(doc, item)

doc.add_paragraph()
add_heading(doc, '9.3  Emergency Shutdown', level=2)
emer = [
    'In an emergency (fire, smoke, unusual sound, burning smell): immediately open the dedicated circuit breaker.',
    'If safe to do so, open the upstream isolation switch and apply lockout/tagout.',
    'Evacuate the area and follow your site emergency response plan.',
    'Contact AST and your site emergency coordinator immediately.',
    'Do not re-energise the unit until AST has inspected and cleared it.',
]
for step in emer:
    add_numbered(doc, step)

add_warning(doc, 'In the event of fire involving electrical equipment, use a CO₂ or dry powder extinguisher only. Never use water.')

doc.add_paragraph()
add_heading(doc, '9.4  Emergency Contacts', level=2)
ec = doc.add_table(rows=4, cols=3)
ec.style = 'Table Grid'
header_row(ec, 'Contact', 'Name / Organisation', 'Number')
ec_data = [
    ('AST 24-hr Support',        'Advancer Smart Technology Pte Ltd',  '[+65 XXXX XXXX]'),
    ('Site Facility Manager',    '[Name]',                              '[Number]'),
    ('SP PowerGrid Emergency',   'SP PowerGrid Ltd',                    '1800 778 8888'),
]
for i, row in enumerate(ec_data):
    for j, val in enumerate(row):
        body_cell(ec.rows[i+1].cells[j], val)
set_cell_borders(ec)

doc.add_paragraph()

# ═══════════════════════════════════════════════════════════════════════════════
#  SECTION 10 — CLIENT HANDOVER AND TRAINING
# ═══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '10.  Client Handover and Training')

add_heading(doc, '10.1  Training Attendance Record', level=2)
tr = doc.add_table(rows=5, cols=4)
tr.style = 'Table Grid'
header_row(tr, 'Name', 'Designation', 'Organisation', 'Signature')
for i in range(1, 5):
    for j in range(4):
        body_cell(tr.rows[i].cells[j], '')
set_cell_borders(tr)

doc.add_paragraph()
add_heading(doc, '10.2  Handover Checklist', level=2)
hc = doc.add_table(rows=10, cols=3)
hc.style = 'Table Grid'
header_row(hc, 'Item', 'Status', 'Remarks')
hc_data = [
    ('Site commissioning completed and signed off',                 '☐ Complete  ☐ Pending', ''),
    ('Equipment verified against handover list (Section 3)',         '☐ Complete  ☐ Pending', ''),
    ('Keys and access items handed over (KT)',                       '☐ Complete  ☐ Pending', ''),
    ('Client trained on normal operation and monitoring dashboard',  '☐ Complete  ☐ Pending', ''),
    ('Client trained on isolation and emergency shutdown',           '☐ Complete  ☐ Pending', ''),
    ('Monitoring dashboard access credentials provided',             '☐ Complete  ☐ Pending', ''),
    ('All manufacturer documents and certificates handed over',      '☐ Complete  ☐ Pending', ''),
    ('Approved SLD and wiring drawings handed over',                 '☐ Complete  ☐ Pending', ''),
    ('Outstanding items noted and agreed resolution date set',       '☐ Complete  ☐ Pending', ''),
]
for i, row in enumerate(hc_data):
    for j, val in enumerate(row):
        body_cell(hc.rows[i+1].cells[j], val)
set_cell_borders(hc)

doc.add_paragraph()
add_heading(doc, '10.3  Outstanding Items', level=2)
oi = doc.add_table(rows=4, cols=4)
oi.style = 'Table Grid'
header_row(oi, 'No.', 'Description', 'Responsible Party', 'Target Date')
for i in range(1, 4):
    body_cell(oi.rows[i].cells[0], str(i), centre=True)
    for j in range(1, 4):
        body_cell(oi.rows[i].cells[j], '')
set_cell_borders(oi)

doc.add_paragraph()
add_heading(doc, '10.4  Client Acknowledgement', level=2)
add_body(doc, 'We, the undersigned, confirm that the HyESys system has been commissioned, handed over and that we have received training sufficient to operate the system in accordance with this manual.')
doc.add_paragraph()
ack = doc.add_table(rows=4, cols=4)
ack.style = 'Table Grid'
header_row(ack, '', 'AST Representative', 'Client Representative', 'Client Representative 2')
ack_rows = [('Name & Designation', '', '', ''), ('Signature', '', '', ''), ('Date', '', '', '')]
for i, row in enumerate(ack_rows):
    set_cell_bg(ack.rows[i+1].cells[0], 'E8EFF8')
    body_cell(ack.rows[i+1].cells[0], row[0], bold=True)
    for j in range(1, 4):
        body_cell(ack.rows[i+1].cells[j], row[j])
set_cell_borders(ack)

doc.add_paragraph()

# ═══════════════════════════════════════════════════════════════════════════════
#  SECTION 11 — SITE SPECIFIC APPENDICES
# ═══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '11.  Site-Specific Appendices')
add_body(doc, 'The following site-specific documents are attached to this manual. All documents must remain with this manual and be updated if any changes are made to the installation.')

app = doc.add_table(rows=9, cols=3)
app.style = 'Table Grid'
header_row(app, 'Appendix', 'Document', 'Status')
app_data = [
    ('A', 'Approved Single Line Diagram (SLD)',                         '☐ Attached  ☐ Pending'),
    ('B', 'Wiring and Termination Drawings',                            '☐ Attached  ☐ Pending'),
    ('C', 'Equipment Layout Drawing',                                   '☐ Attached  ☐ Pending'),
    ('D', 'Cable Route Drawing',                                        '☐ Attached  ☐ Pending'),
    ('E', 'Commissioning Records and Sign-Off Sheet',                   '☐ Attached  ☐ Pending'),
    ('F', 'Test Reports (FAT, SAT, insulation, CT accuracy)',           '☐ Attached  ☐ Pending'),
    ('G', 'Equipment Datasheets (HyESys, HySBatt, Gateway, Logger)',    '☐ Attached  ☐ Pending'),
    ('H', 'Warranty Cards and Certificates',                            '☐ Attached  ☐ Pending'),
]
for i, row in enumerate(app_data):
    for j, val in enumerate(row):
        body_cell(app.rows[i+1].cells[j], val, bold=(j==0), centre=(j==0))
set_cell_borders(app)

doc.add_paragraph()
add_note(doc, 'This manual should be reviewed and updated whenever there is a change to the installation, a firmware upgrade, or a change in site conditions. Contact AST for the latest version.')

# ── Save ──────────────────────────────────────────────────────────────────────
doc.save(OUTPUT_PATH)
print(f"Saved: {OUTPUT_PATH}")
