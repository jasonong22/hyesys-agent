"""
Builds FEE2026_Mechanical_Equations_Reference.docx
108 equations across ME 101-106 / ME 201-206, each with legend + worked example.
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

DST = r'C:\Users\JasonOng\Desktop\local docs\personal\PE\FEE2026_Mechanical_Equations_Reference.docx'

doc = Document()
for sec in doc.sections:
    sec.top_margin    = Inches(0.85)
    sec.bottom_margin = Inches(0.85)
    sec.left_margin   = Inches(0.95)
    sec.right_margin  = Inches(0.95)

C = {
    'title':      RGBColor(0x1A, 0x1A, 0x5E),
    'topic':      RGBColor(0x00, 0x3D, 0x7A),
    'subtopic':   RGBColor(0x00, 0x5C, 0xA8),
    'eqname':     RGBColor(0x12, 0x40, 0x12),
    'formula':    RGBColor(0x0D, 0x0D, 0x7A),
    'legend_hdr': RGBColor(0xFF, 0xFF, 0xFF),
    'body':       RGBColor(0x1A, 0x1A, 0x1A),
    'note':       RGBColor(0x55, 0x55, 0x55),
    'ex_hdr':     RGBColor(0x7B, 0x4B, 0x00),
    'ex_ans':     RGBColor(0x15, 0x57, 0x24),
}
SHADING = {
    'topic':      'D0E4F5',
    'subtopic':   'EBF5FF',
    'eq':         'F5F5F5',
    'legend_hdr': '003D7A',
    'legend_row': 'F0F6FF',
    'legend_alt': 'FAFAFA',
    'ex_hdr':     'FEF3CD',
    'ex_body':    'FFFEF5',
    'ex_ans':     'D4EDDA',
}

def _shd(hex_fill):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_fill)
    return shd

def ap(text='', bold=False, italic=False, mono=False,
       color=None, shading=None, align=None, sz=10, before=40, after=40):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before / 6)
    p.paragraph_format.space_after  = Pt(after  / 6)
    if align == 'center':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if shading:
        p._p.get_or_add_pPr().append(_shd(shading))
    if text:
        r = p.add_run(text)
        r.bold = bold; r.italic = italic
        if mono: r.font.name = 'Courier New'
        r.font.size = Pt(sz)
        if color:
            r.font.color.rgb = color if isinstance(color, RGBColor) \
                               else RGBColor(*[int(color[i:i+2], 16) for i in (0, 2, 4)])
    return p

def topic_hdr(code, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(4)
    p._p.get_or_add_pPr().append(_shd(SHADING['topic']))
    r = p.add_run(f'  {code}  {title}')
    r.bold = True; r.font.size = Pt(13); r.font.color.rgb = C['topic']

def subtopic_hdr(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(3)
    p._p.get_or_add_pPr().append(_shd(SHADING['subtopic']))
    r = p.add_run(f'  {text}')
    r.bold = True; r.font.size = Pt(11); r.font.color.rgb = C['subtopic']

def ex_block(problem, steps, answer):
    ap('  Worked Example', bold=True, shading=SHADING['ex_hdr'],
       color=C['ex_hdr'], sz=8.5, before=3, after=1)
    ap(f'  Q: {problem}', italic=True, sz=9,
       shading=SHADING['ex_body'], color=C['body'], before=0, after=2)
    for step in steps:
        ap(f'     {step}', mono=True, sz=8.5,
           shading=SHADING['ex_body'], color=C['formula'], before=0, after=0)
    ap(f'  Ans: {answer}', bold=True, sz=9,
       shading=SHADING['ex_ans'], color=C['ex_ans'], before=2, after=8)

def eq_block(number, name, formula, legend_rows, note=None, example=None):
    p_name = doc.add_paragraph()
    p_name.paragraph_format.space_before = Pt(5)
    p_name.paragraph_format.space_after  = Pt(1)
    rn = p_name.add_run(f'  [{number}]  {name}')
    rn.bold = True; rn.font.size = Pt(10); rn.font.color.rgb = C['eqname']

    p_form = doc.add_paragraph()
    p_form.paragraph_format.space_before = Pt(0)
    p_form.paragraph_format.space_after  = Pt(1)
    p_form.paragraph_format.left_indent  = Inches(0.35)
    p_form._p.get_or_add_pPr().append(_shd(SHADING['eq']))
    rf = p_form.add_run(f'  {formula}')
    rf.bold = True; rf.font.name = 'Courier New'
    rf.font.size = Pt(10.5); rf.font.color.rgb = C['formula']

    if note:
        p_note = doc.add_paragraph()
        p_note.paragraph_format.space_before = Pt(0)
        p_note.paragraph_format.space_after  = Pt(1)
        p_note.paragraph_format.left_indent  = Inches(0.35)
        rno = p_note.add_run(f'  > {note}')
        rno.italic = True; rno.font.size = Pt(8.5); rno.font.color.rgb = C['note']

    tbl = doc.add_table(rows=1 + len(legend_rows), cols=3)
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    widths = [Inches(0.85), Inches(4.10), Inches(1.20)]
    for row in tbl.rows:
        for idx, cell in enumerate(row.cells):
            cell.width = widths[idx]

    hdr_texts = ['Symbol', 'Description', 'SI Units']
    for i, cell in enumerate(tbl.rows[0].cells):
        cell._tc.get_or_add_tcPr().append(_shd(SHADING['legend_hdr']))
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(hdr_texts[i])
        run.bold = True; run.font.size = Pt(9); run.font.color.rgb = C['legend_hdr']

    for ri, (sym, desc, unit) in enumerate(legend_rows):
        row = tbl.rows[ri + 1]
        fill = SHADING['legend_row'] if ri % 2 == 0 else SHADING['legend_alt']
        for ci, (cell, val) in enumerate(zip(row.cells, [sym, desc, unit])):
            cell._tc.get_or_add_tcPr().append(_shd(fill))
            p = cell.paragraphs[0]
            if ci == 0: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(val)
            run.font.size = Pt(9)
            if ci == 0:
                run.bold = True; run.font.name = 'Courier New'
                run.font.color.rgb = C['formula']
            else:
                run.font.color.rgb = C['body']

    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    if example:
        ex_block(example['problem'], example['steps'], example['answer'])


# ══════════════════════════════════════════════════════════════════════════════
# DOCUMENT HEADER
# ══════════════════════════════════════════════════════════════════════════════
ap('PROFESSIONAL ENGINEERS EXAMINATION FEE 2026', bold=True, sz=15,
   color=C['title'], align='center', before=0, after=10)
ap('Mechanical Discipline — Complete Equation Reference with Worked Examples',
   bold=True, sz=12, color=C['subtopic'], align='center', before=0, after=6)
ap('Topics: ME 101-106 / ME 201-206  |  All equations web-verified  |  SI units throughout',
   sz=9, color=C['note'], align='center', before=0, after=20)
ap('-' * 105, mono=True, sz=7, color=C['note'], before=0, after=16)

# ══════════════════════════════════════════════════════════════════════════════
# ME 101/201  CONTROL AND INSTRUMENTATION
# ══════════════════════════════════════════════════════════════════════════════
topic_hdr('ME 101 / 201', 'Control and Instrumentation')
subtopic_hdr('Transfer Functions & System Modelling')

eq_block(1, 'Standard Second-Order Transfer Function',
    'G(s) = wn^2 / (s^2 + 2*zeta*wn*s + wn^2)',
    [('G(s)', 'Transfer function (output/input in Laplace domain)', 'dimensionless'),
     ('s',    'Complex frequency variable (Laplace operator)', 'rad/s'),
     ('wn',   'Undamped natural frequency', 'rad/s'),
     ('zeta', 'Damping ratio', 'dimensionless')],
    example={
        'problem': 'A system has G(s) = 100/(s^2 + 6s + 100). Find wn and zeta.',
        'steps': [
            'wn^2 = 100  =>  wn = 10 rad/s',
            '2*zeta*wn = 6  =>  zeta = 6/(2*10) = 0.30',
        ],
        'answer': 'wn = 10 rad/s, zeta = 0.30 (underdamped)'})

eq_block(2, 'Closed-Loop Characteristic Equation',
    '1 + G(s)*H(s) = 0',
    [('G(s)', 'Forward-path (plant) transfer function', 'dimensionless'),
     ('H(s)', 'Feedback transfer function', 'dimensionless')],
    note='Poles of this equation determine stability. Roots must have negative real parts.',
    example={
        'problem': 'G(s) = 10/(s+5), H(s) = 1 (unity feedback). Find closed-loop pole.',
        'steps': [
            '1 + 10/(s+5) = 0  =>  s+5+10 = 0  =>  s = -15',
        ],
        'answer': 'Closed-loop pole at s = -15 (stable, left half-plane)'})

eq_block(3, 'Damping Ratio (from characteristic polynomial as^2 + bs + c = 0)',
    'zeta = b / (2*sqrt(a*c))',
    [('zeta', 'Damping ratio', 'dimensionless'),
     ('a',    'Coefficient of s^2 term', 'varies'),
     ('b',    'Coefficient of s term', 'varies'),
     ('c',    'Constant (s^0) term', 'varies')],
    example={
        'problem': 'Characteristic polynomial: 2s^2 + 8s + 18 = 0. Find zeta.',
        'steps': [
            'a=2, b=8, c=18',
            'zeta = 8 / (2*sqrt(2*18)) = 8 / (2*sqrt(36)) = 8/12 = 0.667',
        ],
        'answer': 'zeta = 0.667 (underdamped)'})

subtopic_hdr('Transient Response (Underdamped, 0 < zeta < 1)')

eq_block(4, 'Damped Natural Frequency',
    'wd = wn * sqrt(1 - zeta^2)',
    [('wd',   'Damped natural frequency', 'rad/s'),
     ('wn',   'Undamped natural frequency', 'rad/s'),
     ('zeta', 'Damping ratio', 'dimensionless')],
    example={
        'problem': 'wn = 10 rad/s, zeta = 0.30. Find wd.',
        'steps': [
            'wd = 10 * sqrt(1 - 0.09) = 10 * sqrt(0.91) = 10 * 0.954 = 9.54 rad/s',
        ],
        'answer': 'wd = 9.54 rad/s'})

eq_block(5, 'Rise Time (10-90%)',
    'tr = (pi - phi) / wd    where phi = arctan(sqrt(1-zeta^2) / zeta)',
    [('tr',   'Rise time', 's'),
     ('phi',  'Phase angle = arctan(sqrt(1-zeta^2)/zeta)', 'rad'),
     ('wd',   'Damped natural frequency', 'rad/s'),
     ('zeta', 'Damping ratio', 'dimensionless')],
    example={
        'problem': 'wn = 10 rad/s, zeta = 0.30. Find tr.',
        'steps': [
            'wd = 9.54 rad/s',
            'phi = arctan(0.954/0.30) = arctan(3.18) = 1.265 rad',
            'tr = (3.1416 - 1.265) / 9.54 = 1.877 / 9.54 = 0.197 s',
        ],
        'answer': 'tr = 0.197 s'})

eq_block(6, 'Peak Time',
    'tp = pi / wd',
    [('tp', 'Time to first peak overshoot', 's'),
     ('wd', 'Damped natural frequency', 'rad/s')],
    example={
        'problem': 'wd = 9.54 rad/s. Find tp.',
        'steps': [
            'tp = pi / 9.54 = 3.1416 / 9.54 = 0.329 s',
        ],
        'answer': 'tp = 0.329 s'})

eq_block(7, 'Percent Overshoot',
    '%OS = exp(-pi*zeta / sqrt(1 - zeta^2)) * 100',
    [('%OS',  'Percent overshoot', '%'),
     ('zeta', 'Damping ratio', 'dimensionless')],
    note='Valid for standard 2nd-order systems without zeros.',
    example={
        'problem': 'zeta = 0.30. Find %OS.',
        'steps': [
            '%OS = exp(-pi*0.30 / sqrt(1-0.09)) * 100',
            '    = exp(-0.9425 / 0.954) * 100 = exp(-0.988) * 100 = 37.2%',
        ],
        'answer': '%OS = 37.2%'})

eq_block(8, 'Settling Time (2% criterion)',
    'ts = 4 / (zeta * wn)',
    [('ts',   'Settling time (response within +/-2% of final value)', 's'),
     ('zeta', 'Damping ratio', 'dimensionless'),
     ('wn',   'Undamped natural frequency', 'rad/s')],
    note='For 5% criterion use ts = 3/(zeta*wn).',
    example={
        'problem': 'wn = 10 rad/s, zeta = 0.30. Find ts (2% criterion).',
        'steps': [
            'ts = 4 / (0.30 * 10) = 4 / 3 = 1.33 s',
        ],
        'answer': 'ts = 1.33 s'})

subtopic_hdr('Steady-State Errors (Unity Feedback)')

eq_block(9, 'Position Error Constant (Type 0 system - step input)',
    'Kp = lim[s->0] G(s)     =>     e_ss = 1 / (1 + Kp)',
    [('Kp',   'Position error constant', 'dimensionless'),
     ('e_ss', 'Steady-state error', 'same units as input'),
     ('G(s)', 'Open-loop transfer function', 'dimensionless')],
    example={
        'problem': 'G(s) = 20/(s+4), unity feedback, unit step input. Find Kp and e_ss.',
        'steps': [
            'Kp = lim[s->0] 20/(s+4) = 20/4 = 5',
            'e_ss = 1/(1+5) = 1/6 = 0.167',
        ],
        'answer': 'Kp = 5, e_ss = 0.167 (16.7% of step magnitude)'})

eq_block(10, 'Velocity Error Constant (Type 1 system - ramp input)',
    'Kv = lim[s->0] s*G(s)     =>     e_ss = 1 / Kv',
    [('Kv',   'Velocity error constant', 's^-1'),
     ('e_ss', 'Steady-state error', 's'),
     ('G(s)', 'Open-loop transfer function', 'dimensionless')],
    example={
        'problem': 'G(s) = 5/(s*(s+2)), unity feedback, ramp input. Find Kv and e_ss.',
        'steps': [
            'Kv = lim[s->0] s * 5/(s*(s+2)) = lim[s->0] 5/(s+2) = 5/2 = 2.5 s^-1',
            'e_ss = 1/Kv = 1/2.5 = 0.40',
        ],
        'answer': 'Kv = 2.5 s^-1, e_ss = 0.40'})

subtopic_hdr('PID Controller')

eq_block(11, 'PID Controller Transfer Function',
    'C(s) = Kp + Ki/s + Kd*s',
    [('C(s)', 'Controller transfer function', 'dimensionless'),
     ('Kp',   'Proportional gain', 'dimensionless'),
     ('Ki',   'Integral gain', 's^-1'),
     ('Kd',   'Derivative gain', 's'),
     ('s',    'Laplace operator', 'rad/s')],
    example={
        'problem': 'Kp=8, Ki=4, Kd=0.5. Write C(s) and find the combined transfer function.',
        'steps': [
            'C(s) = 8 + 4/s + 0.5s',
            'Combined: C(s) = (0.5s^2 + 8s + 4) / s',
        ],
        'answer': 'C(s) = (0.5s^2 + 8s + 4)/s'})

subtopic_hdr('Frequency Domain Stability Margins')

eq_block(12, 'Phase Margin',
    'PM = 180 deg + angle[G(j*wgc)]',
    [('PM',       'Phase margin (system stable if PM > 0 deg)', 'degrees'),
     ('wgc',      'Gain crossover frequency (where |G(jw)| = 1)', 'rad/s'),
     ('angle G',  'Phase angle of open-loop TF at wgc', 'degrees')],
    example={
        'problem': 'At gain crossover frequency, G(jwgc) has phase angle -145 deg. Find PM.',
        'steps': [
            'PM = 180 + (-145) = 35 deg',
        ],
        'answer': 'PM = 35 deg (stable, PM > 0 deg)'})

eq_block(13, 'Gain Margin',
    'GM = 1 / |G(j*wpc)|     or     GM_dB = -20*log10|G(j*wpc)|',
    [('GM',       'Gain margin (stable if GM > 1)', 'dimensionless (or dB)'),
     ('wpc',      'Phase crossover frequency (where angle G = -180 deg)', 'rad/s'),
     ('|G(jw)|',  'Magnitude of open-loop TF at wpc', 'dimensionless')],
    example={
        'problem': '|G(jwpc)| = 0.25 at phase crossover. Find GM in linear and dB.',
        'steps': [
            'GM = 1/0.25 = 4',
            'GM_dB = 20*log10(4) = 12.0 dB',
        ],
        'answer': 'GM = 4 (12.0 dB) — system is stable'})


# ══════════════════════════════════════════════════════════════════════════════
# ME 102/202  DYNAMICS AND VIBRATIONS
# ══════════════════════════════════════════════════════════════════════════════
topic_hdr('ME 102 / 202', 'Dynamics and Vibrations')
subtopic_hdr('Kinematics - Particle (Constant Acceleration)')

eq_block(14, "Newton's Second Law - Linear Motion",
    'F = m * a',
    [('F', 'Net force', 'N'),
     ('m', 'Mass', 'kg'),
     ('a', 'Linear acceleration', 'm/s^2')],
    example={
        'problem': 'A 5 kg block experiences a net force of 30 N. Find acceleration.',
        'steps': ['a = F/m = 30/5 = 6 m/s^2'],
        'answer': 'a = 6 m/s^2'})

eq_block(15, 'SUVAT Equations of Motion (constant acceleration)',
    'v = u + at  |  s = ut + (1/2)*a*t^2  |  v^2 = u^2 + 2as',
    [('v', 'Final velocity', 'm/s'),
     ('u', 'Initial velocity', 'm/s'),
     ('a', 'Constant acceleration', 'm/s^2'),
     ('t', 'Time', 's'),
     ('s', 'Displacement', 'm')],
    example={
        'problem': 'Car starts from rest, accelerates at 3 m/s^2 for 8 s. Find v and s.',
        'steps': [
            'v = 0 + 3*8 = 24 m/s',
            's = 0 + (1/2)*3*8^2 = (1/2)*3*64 = 96 m',
        ],
        'answer': 'v = 24 m/s, s = 96 m'})

eq_block(16, 'Centripetal Acceleration',
    'ac = v^2/r = w^2*r',
    [('ac', 'Centripetal (radial) acceleration, directed toward centre', 'm/s^2'),
     ('v',  'Tangential velocity', 'm/s'),
     ('r',  'Radius of circular path', 'm'),
     ('w',  'Angular velocity', 'rad/s')],
    example={
        'problem': 'A ball on a 1.2 m string rotates at w = 4 rad/s. Find centripetal acceleration.',
        'steps': ['ac = w^2 * r = 4^2 * 1.2 = 16 * 1.2 = 19.2 m/s^2'],
        'answer': 'ac = 19.2 m/s^2 (directed toward centre)'})

subtopic_hdr('Kinetics - Energy and Momentum')

eq_block(17, 'Work-Energy Theorem',
    'W_net = DeltaKE = (1/2)*m*v2^2 - (1/2)*m*v1^2',
    [('W_net', 'Net work done on the body', 'J'),
     ('KE',    'Kinetic energy', 'J'),
     ('m',     'Mass', 'kg'),
     ('v1',    'Initial velocity', 'm/s'),
     ('v2',    'Final velocity', 'm/s')],
    example={
        'problem': 'A 2 kg object accelerates from 3 m/s to 7 m/s. Find net work done.',
        'steps': [
            'W_net = (1/2)*2*7^2 - (1/2)*2*3^2',
            '      = (1/2)*2*49 - (1/2)*2*9 = 49 - 9 = 40 J',
        ],
        'answer': 'W_net = 40 J'})

eq_block(18, 'Conservation of Linear Momentum',
    'm1*v1 + m2*v2 = m1*v1\' + m2*v2\'',
    [('m1, m2',    'Masses of bodies 1 and 2', 'kg'),
     ("v1, v2",    'Velocities before collision', 'm/s'),
     ("v1', v2'",  'Velocities after collision', 'm/s')],
    note='Valid when no external impulse acts on the system.',
    example={
        'problem': '3 kg at 4 m/s hits stationary 1 kg. After: 1 kg moves at 6 m/s. Find v1\'.',
        'steps': [
            '3*4 + 1*0 = 3*v1\' + 1*6',
            '12 = 3*v1\' + 6  =>  v1\' = 2 m/s',
        ],
        'answer': '3 kg body continues at 2 m/s'})

eq_block(19, 'Coefficient of Restitution',
    "e = (v2' - v1') / (v1 - v2)",
    [('e',        'Coefficient of restitution  (0 = plastic, 1 = elastic)', 'dimensionless'),
     ("v1, v2",   'Velocities before impact', 'm/s'),
     ("v1', v2'", 'Velocities after impact', 'm/s')],
    example={
        'problem': "Same collision: v1=4, v2=0, v1'=2, v2'=6. Find e.",
        'steps': [
            "e = (v2' - v1') / (v1 - v2) = (6 - 2) / (4 - 0) = 4/4 = 1.0",
        ],
        'answer': 'e = 1.0 (perfectly elastic collision)'})

subtopic_hdr('Kinematics and Kinetics - Rotation')

eq_block(20, "Newton's Second Law - Rotational Motion",
    'M = I * alpha',
    [('M',     'Net moment (torque)', 'N*m'),
     ('I',     'Mass moment of inertia about rotation axis', 'kg*m^2'),
     ('alpha', 'Angular acceleration', 'rad/s^2')],
    example={
        'problem': 'A flywheel (I = 0.5 kg*m^2) has net torque 4 N*m. Find angular acceleration.',
        'steps': ['alpha = M/I = 4/0.5 = 8 rad/s^2'],
        'answer': 'alpha = 8 rad/s^2'})

eq_block(21, 'Angular Momentum',
    'L = I * w     and     M = dL/dt',
    [('L', 'Angular momentum', 'kg*m^2/s'),
     ('I', 'Mass moment of inertia', 'kg*m^2'),
     ('w', 'Angular velocity', 'rad/s'),
     ('M', 'Applied moment (torque)', 'N*m')],
    example={
        'problem': 'Flywheel I = 0.5 kg*m^2 spinning at w = 100 rad/s. Find L.',
        'steps': ['L = I*w = 0.5 * 100 = 50 kg*m^2/s'],
        'answer': 'L = 50 kg*m^2/s'})

subtopic_hdr('Mechanical Vibrations - SDOF Systems')

eq_block(22, 'Undamped Natural Frequency',
    'wn = sqrt(k/m)',
    [('wn', 'Undamped natural frequency', 'rad/s'),
     ('k',  'Spring stiffness', 'N/m'),
     ('m',  'Mass', 'kg')],
    example={
        'problem': 'Spring-mass system: k = 4000 N/m, m = 10 kg. Find wn and frequency in Hz.',
        'steps': [
            'wn = sqrt(4000/10) = sqrt(400) = 20 rad/s',
            'f = wn/(2*pi) = 20/6.283 = 3.18 Hz',
        ],
        'answer': 'wn = 20 rad/s (3.18 Hz)'})

eq_block(23, 'Equation of Motion - Undamped Free Vibration',
    'm*x_ddot + k*x = F(t)',
    [('m',    'Mass', 'kg'),
     ('x_ddot', 'Acceleration (second derivative of x)', 'm/s^2'),
     ('k',    'Stiffness', 'N/m'),
     ('x',    'Displacement from equilibrium', 'm'),
     ('F(t)', 'External forcing function', 'N')],
    example={
        'problem': 'm = 2 kg, k = 800 N/m, no forcing. Write EOM and state wn.',
        'steps': [
            '2*x_ddot + 800*x = 0',
            'wn = sqrt(800/2) = sqrt(400) = 20 rad/s',
        ],
        'answer': '2*x_ddot + 800*x = 0; wn = 20 rad/s'})

eq_block(24, 'Equation of Motion - Damped Free Vibration',
    'm*x_ddot + c*x_dot + k*x = F(t)',
    [('m',    'Mass', 'kg'),
     ('c',    'Viscous damping coefficient', 'N*s/m'),
     ('x_dot', 'Velocity', 'm/s'),
     ('k',    'Stiffness', 'N/m'),
     ('x',    'Displacement', 'm'),
     ('F(t)', 'External force', 'N')],
    example={
        'problem': 'm=2 kg, c=12 N*s/m, k=800 N/m. Write EOM and classify damping.',
        'steps': [
            '2*x_ddot + 12*x_dot + 800*x = 0',
            'cc = 2*sqrt(k*m) = 2*sqrt(1600) = 80 N*s/m',
            'zeta = c/cc = 12/80 = 0.15  =>  underdamped',
        ],
        'answer': 'EOM: 2*x_ddot + 12*x_dot + 800*x = 0; zeta = 0.15 (underdamped)'})

eq_block(25, 'Damping Ratio',
    'zeta = c / (2*sqrt(k*m)) = c / (2*m*wn)',
    [('zeta', 'Damping ratio', 'dimensionless'),
     ('c',    'Damping coefficient', 'N*s/m'),
     ('k',    'Stiffness', 'N/m'),
     ('m',    'Mass', 'kg'),
     ('wn',   'Natural frequency', 'rad/s')],
    example={
        'problem': 'm = 2 kg, c = 12 N*s/m, k = 800 N/m. Find zeta.',
        'steps': ['zeta = 12 / (2*sqrt(800*2)) = 12 / (2*40) = 12/80 = 0.15'],
        'answer': 'zeta = 0.15 (underdamped system)'})

eq_block(26, 'Critical Damping Coefficient',
    'cc = 2*sqrt(k*m) = 2*m*wn',
    [('cc', 'Critical damping coefficient (boundary: zeta = 1)', 'N*s/m'),
     ('k',  'Stiffness', 'N/m'),
     ('m',  'Mass', 'kg'),
     ('wn', 'Undamped natural frequency', 'rad/s')],
    example={
        'problem': 'm = 5 kg, k = 2000 N/m. Find critical damping coefficient.',
        'steps': ['cc = 2*sqrt(k*m) = 2*sqrt(2000*5) = 2*sqrt(10000) = 2*100 = 200 N*s/m'],
        'answer': 'cc = 200 N*s/m'})

eq_block(27, 'Resonance Frequency (Damped Forced Vibration)',
    'wr = wn * sqrt(1 - 2*zeta^2)',
    [('wr',   'Frequency of maximum amplitude response', 'rad/s'),
     ('wn',   'Undamped natural frequency', 'rad/s'),
     ('zeta', 'Damping ratio', 'dimensionless')],
    note='Valid only when zeta <= 1/sqrt(2) = 0.707.',
    example={
        'problem': 'wn = 20 rad/s, zeta = 0.15. Find wr.',
        'steps': [
            'wr = 20 * sqrt(1 - 2*0.15^2) = 20 * sqrt(1 - 0.045) = 20 * sqrt(0.955)',
            '   = 20 * 0.977 = 19.54 rad/s',
        ],
        'answer': 'wr = 19.54 rad/s'})

eq_block(28, 'Logarithmic Decrement',
    'delta = ln(xn / x(n+1)) = 2*pi*zeta / sqrt(1 - zeta^2)',
    [('delta',    'Logarithmic decrement', 'dimensionless'),
     ('xn',       'Amplitude of nth cycle', 'm'),
     ('x(n+1)',   'Amplitude of (n+1)th cycle', 'm'),
     ('zeta',     'Damping ratio', 'dimensionless')],
    example={
        'problem': 'Consecutive amplitudes: x1 = 18 mm, x2 = 12 mm. Find zeta.',
        'steps': [
            'delta = ln(18/12) = ln(1.5) = 0.405',
            'zeta = delta / sqrt(4*pi^2 + delta^2) = 0.405 / sqrt(39.48 + 0.164) = 0.405/6.296 = 0.0644',
        ],
        'answer': 'zeta = 0.064'})


# ══════════════════════════════════════════════════════════════════════════════
# ME 103/203  FLUID MECHANICS
# ══════════════════════════════════════════════════════════════════════════════
topic_hdr('ME 103 / 203', 'Fluid Mechanics')
subtopic_hdr('Fluid Properties')

eq_block(29, 'Density',
    'rho = m / V',
    [('rho', 'Mass density', 'kg/m^3'),
     ('m',   'Mass', 'kg'),
     ('V',   'Volume', 'm^3')],
    example={
        'problem': 'A 5-litre container holds 4.65 kg of oil. Find its density.',
        'steps': ['rho = m/V = 4.65 / 0.005 = 930 kg/m^3'],
        'answer': 'rho = 930 kg/m^3 (typical light crude oil)'})

eq_block(30, "Newton's Law of Viscosity",
    'tau = mu * (du/dy)',
    [('tau',   'Shear stress in fluid', 'Pa'),
     ('mu',    'Dynamic viscosity', 'Pa*s'),
     ('du/dy', 'Velocity gradient perpendicular to flow', 's^-1')],
    example={
        'problem': 'Oil (mu=0.1 Pa*s) between plates 5 mm apart; upper plate at 2 m/s, lower fixed. Find tau.',
        'steps': [
            'du/dy = 2 / 0.005 = 400 s^-1',
            'tau = 0.1 * 400 = 40 Pa',
        ],
        'answer': 'tau = 40 Pa'})

eq_block(31, 'Kinematic Viscosity',
    'nu = mu / rho',
    [('nu',  'Kinematic viscosity', 'm^2/s'),
     ('mu',  'Dynamic viscosity', 'Pa*s'),
     ('rho', 'Fluid density', 'kg/m^3')],
    example={
        'problem': 'Oil: mu = 0.1 Pa*s, rho = 900 kg/m^3. Find nu.',
        'steps': ['nu = 0.1 / 900 = 1.11e-4 m^2/s = 111 cSt'],
        'answer': 'nu = 1.11e-4 m^2/s'})

subtopic_hdr('Hydrostatics')

eq_block(32, 'Hydrostatic Pressure',
    'P = P0 + rho*g*h',
    [('P',   'Absolute pressure at depth h', 'Pa'),
     ('P0',  'Pressure at free surface', 'Pa'),
     ('rho', 'Fluid density', 'kg/m^3'),
     ('g',   'Gravitational acceleration (9.81)', 'm/s^2'),
     ('h',   'Depth below free surface', 'm')],
    example={
        'problem': 'Find gauge pressure at 15 m depth in seawater (rho = 1025 kg/m^3).',
        'steps': ['P_gauge = rho*g*h = 1025 * 9.81 * 15 = 150,919 Pa = 151 kPa'],
        'answer': 'P_gauge = 151 kPa'})

eq_block(33, "Buoyancy Force - Archimedes' Principle",
    'Fb = rho_fluid * V_sub * g',
    [('Fb',        'Buoyancy (upward) force', 'N'),
     ('rho_fluid', 'Density of surrounding fluid', 'kg/m^3'),
     ('V_sub',     'Volume of body submerged in fluid', 'm^3'),
     ('g',         'Gravitational acceleration', 'm/s^2')],
    example={
        'problem': 'A 0.2 m^3 object submerged in water. Find buoyancy force.',
        'steps': ['Fb = 1000 * 0.2 * 9.81 = 1962 N'],
        'answer': 'Fb = 1962 N (= 1.96 kN)'})

subtopic_hdr('Flow Equations')

eq_block(34, 'Continuity Equation - Incompressible Flow',
    'A1*V1 = A2*V2 = Q',
    [('A1, A2', 'Cross-sectional areas at sections 1 and 2', 'm^2'),
     ('V1, V2', 'Mean flow velocities at sections 1 and 2', 'm/s'),
     ('Q',      'Volumetric flow rate (constant)', 'm^3/s')],
    example={
        'problem': 'Water flows at V1=3 m/s in 200 mm pipe. Pipe reduces to 100 mm. Find V2.',
        'steps': [
            'A1 = pi*(0.2)^2/4 = 0.03142 m^2',
            'A2 = pi*(0.1)^2/4 = 0.007854 m^2',
            'V2 = A1*V1/A2 = 0.03142*3 / 0.007854 = 12 m/s',
        ],
        'answer': 'V2 = 12 m/s'})

eq_block(35, 'Continuity Equation - Compressible Flow',
    'rho1*A1*V1 = rho2*A2*V2 = m_dot',
    [('rho1, rho2', 'Fluid densities at sections 1 and 2', 'kg/m^3'),
     ('A1, A2',     'Cross-sectional areas', 'm^2'),
     ('V1, V2',     'Velocities', 'm/s'),
     ('m_dot',      'Mass flow rate (constant)', 'kg/s')],
    example={
        'problem': 'Air: rho1=1.2 kg/m^3, V1=50 m/s, A1=0.1 m^2. At exit: rho2=0.9, A2=0.08 m^2. Find V2.',
        'steps': [
            'm_dot = rho1*A1*V1 = 1.2*0.1*50 = 6 kg/s',
            'V2 = m_dot/(rho2*A2) = 6/(0.9*0.08) = 83.3 m/s',
        ],
        'answer': 'V2 = 83.3 m/s'})

eq_block(36, "Bernoulli's Equation (steady, inviscid, incompressible)",
    'P1 + (1/2)*rho*V1^2 + rho*g*z1  =  P2 + (1/2)*rho*V2^2 + rho*g*z2',
    [('P',   'Static pressure', 'Pa'),
     ('rho', 'Fluid density', 'kg/m^3'),
     ('V',   'Flow velocity', 'm/s'),
     ('g',   'Gravitational acceleration', 'm/s^2'),
     ('z',   'Elevation above datum', 'm')],
    note='Head form (divide by rho*g): P/rho*g + V^2/2g + z = H (total head, m)',
    example={
        'problem': 'Pipe narrows 200mm->100mm, same elevation. P1=150 kPa, V1=3 m/s. Find P2 (water).',
        'steps': [
            'V2 = 12 m/s (from continuity, eq. 34)',
            'P2 = P1 + (1/2)*rho*(V1^2 - V2^2)',
            'P2 = 150000 + (1/2)*1000*(9 - 144) = 150000 - 67500 = 82,500 Pa',
        ],
        'answer': 'P2 = 82.5 kPa'})

eq_block(37, 'Momentum Equation - Steady 1-D Flow',
    'Sum(F) = rho*Q*(V2 - V1) = m_dot*(V2 - V1)',
    [('Sum(F)',  'Net external force on control volume', 'N'),
     ('rho',    'Fluid density', 'kg/m^3'),
     ('Q',      'Volumetric flow rate', 'm^3/s'),
     ('V1, V2', 'Velocities at inlet and outlet', 'm/s'),
     ('m_dot',  'Mass flow rate', 'kg/s')],
    example={
        'problem': 'Water m_dot=10 kg/s hits a flat plate at 5 m/s and is deflected 90 deg. Find reaction force.',
        'steps': [
            'Fx = m_dot*(0 - 5) = -50 N (x-direction)',
            'Fy = m_dot*(5 - 0) = +50 N (y-direction)',
            '|F| = sqrt(50^2 + 50^2) = 70.7 N',
        ],
        'answer': 'Reaction force = 70.7 N at 45 deg'})

subtopic_hdr('Dimensionless Numbers')

eq_block(38, 'Reynolds Number',
    'Re = rho*V*D/mu = V*D/nu',
    [('Re', 'Reynolds number (Re<2300 laminar; Re>4000 turbulent in pipe)', 'dimensionless'),
     ('rho', 'Fluid density', 'kg/m^3'),
     ('V',   'Characteristic velocity', 'm/s'),
     ('D',   'Characteristic length (pipe diameter)', 'm'),
     ('mu',  'Dynamic viscosity', 'Pa*s'),
     ('nu',  'Kinematic viscosity nu=mu/rho', 'm^2/s')],
    example={
        'problem': 'Water (rho=1000, mu=0.001 Pa*s) at 2 m/s in 50 mm pipe. Find Re.',
        'steps': ['Re = 1000*2*0.05 / 0.001 = 100,000'],
        'answer': 'Re = 100,000 => turbulent flow (Re >> 4000)'})

eq_block(39, 'Froude Number',
    'Fr = V / sqrt(g*L)',
    [('Fr', 'Froude number (Fr<1 subcritical; Fr>1 supercritical)', 'dimensionless'),
     ('V',  'Flow velocity', 'm/s'),
     ('g',  'Gravitational acceleration', 'm/s^2'),
     ('L',  'Characteristic length (e.g. hydraulic depth)', 'm')],
    example={
        'problem': 'River: V = 1.5 m/s, hydraulic depth = 0.8 m. Find Fr.',
        'steps': ['Fr = 1.5 / sqrt(9.81*0.8) = 1.5 / 2.801 = 0.535'],
        'answer': 'Fr = 0.535 < 1 => subcritical (tranquil) flow'})

eq_block(40, 'Mach Number',
    'Ma = V / c     where  c = sqrt(gamma*R*T)',
    [('Ma',    'Mach number (Ma<1 subsonic; Ma>1 supersonic)', 'dimensionless'),
     ('V',     'Flow velocity', 'm/s'),
     ('c',     'Speed of sound in fluid', 'm/s'),
     ('gamma', 'Ratio of specific heats cp/cv (air = 1.4)', 'dimensionless'),
     ('R',     'Specific gas constant (air = 287)', 'J/(kg*K)'),
     ('T',     'Absolute temperature', 'K')],
    example={
        'problem': 'Aircraft at V=340 m/s. Air T=288 K. Find Ma.',
        'steps': [
            'c = sqrt(1.4*287*288) = sqrt(115,834) = 340.3 m/s',
            'Ma = 340/340.3 = 0.999',
        ],
        'answer': 'Ma = 1.0 (transonic)'})

subtopic_hdr('Pipe Flow - Head Losses')

eq_block(41, 'Darcy-Weisbach Equation - Major (Friction) Loss',
    'hf = f * (L/D) * V^2/(2g)',
    [('hf', 'Friction head loss', 'm'),
     ('f',  'Darcy friction factor (dimensionless)', 'dimensionless'),
     ('L',  'Pipe length', 'm'),
     ('D',  'Internal pipe diameter', 'm'),
     ('V',  'Mean flow velocity', 'm/s'),
     ('g',  'Gravitational acceleration', 'm/s^2')],
    note='Fanning form (some texts): hf = 4f(L/D)(V^2/2g). Darcy f = 4 x Fanning f.',
    example={
        'problem': 'Water V=3 m/s, L=500 m, D=150 mm, f=0.015. Find hf.',
        'steps': [
            'hf = 0.015 * (500/0.15) * (3^2/(2*9.81))',
            'hf = 0.015 * 3333 * 0.459 = 22.9 m',
        ],
        'answer': 'hf = 22.9 m'})

eq_block(42, 'Laminar Friction Factor - Hagen-Poiseuille',
    'f = 64 / Re       (valid for Re < 2300)',
    [('f',  'Darcy friction factor', 'dimensionless'),
     ('Re', 'Reynolds number', 'dimensionless')],
    example={
        'problem': 'Viscous oil flows in a pipe with Re = 1200. Find Darcy friction factor.',
        'steps': ['f = 64/Re = 64/1200 = 0.0533'],
        'answer': 'f = 0.053 (laminar flow, Re < 2300)'})

eq_block(43, 'Turbulent Friction Factor - Colebrook-White Equation',
    '1/sqrt(f) = -2*log10[ eps/(3.7*D) + 2.51/(Re*sqrt(f)) ]',
    [('f',   'Darcy friction factor', 'dimensionless'),
     ('eps', 'Absolute pipe roughness', 'm'),
     ('D',   'Pipe internal diameter', 'm'),
     ('Re',  'Reynolds number', 'dimensionless')],
    note='Implicit equation - requires iteration or Moody chart. Valid for turbulent flow (Re > 4000).',
    example={
        'problem': 'Steel pipe: eps=0.046 mm, D=100 mm, Re=100,000. Estimate f (1st iteration, f0=0.018).',
        'steps': [
            '1/sqrt(f) = -2*log10[0.046/(3.7*100) + 2.51/(100000*sqrt(0.018))]',
            '         = -2*log10[1.24e-4 + 1.87e-4] = -2*log10[3.11e-4] = 7.01',
            'f1 = 1/7.01^2 = 0.0203  (iterate to convergence)',
        ],
        'answer': 'f = 0.020 after convergence'})

eq_block(44, 'Minor (Local) Losses - Fittings, Entry, Exit',
    'hm = KL * V^2/(2g)',
    [('hm', 'Minor head loss', 'm'),
     ('KL', 'Loss coefficient (K_entry=0.5, K_exit=1.0)', 'dimensionless'),
     ('V',  'Velocity at the fitting', 'm/s'),
     ('g',  'Gravitational acceleration', 'm/s^2')],
    example={
        'problem': 'Sharp-edged pipe entry (KL=0.5), V=3 m/s. Find hm.',
        'steps': ['hm = 0.5 * 3^2/(2*9.81) = 0.5 * 9/19.62 = 0.229 m'],
        'answer': 'hm = 0.23 m'})

eq_block(45, 'Sudden Expansion Loss - Borda-Carnot',
    'h_exp = (V1 - V2)^2 / (2g)',
    [('h_exp',  'Head loss at sudden expansion', 'm'),
     ('V1',     'Upstream velocity (smaller pipe)', 'm/s'),
     ('V2',     'Downstream velocity (larger pipe)', 'm/s'),
     ('A1, A2', 'Cross-sectional areas upstream and downstream', 'm^2'),
     ('g',      'Gravitational acceleration', 'm/s^2')],
    example={
        'problem': 'Flow at 4 m/s in 80 mm pipe suddenly expands to 150 mm. Find h_exp.',
        'steps': [
            'V2 = V1*(A1/A2) = 4*(80/150)^2 = 4*0.284 = 1.14 m/s',
            'h_exp = (4 - 1.14)^2 / (2*9.81) = (2.86)^2/19.62 = 8.18/19.62 = 0.417 m',
        ],
        'answer': 'h_exp = 0.42 m'})

subtopic_hdr('Fluid Machinery')

eq_block(46, 'Hydraulic Power',
    'P = rho*g*Q*H',
    [('P',   'Hydraulic power delivered to fluid', 'W'),
     ('rho', 'Fluid density', 'kg/m^3'),
     ('g',   'Gravitational acceleration', 'm/s^2'),
     ('Q',   'Volumetric flow rate', 'm^3/s'),
     ('H',   'Total head developed', 'm')],
    example={
        'problem': 'Pump delivers Q=0.05 m^3/s of water at H=25 m total head. Find hydraulic power.',
        'steps': ['P = 1000*9.81*0.05*25 = 12,263 W'],
        'answer': 'P = 12.3 kW'})

eq_block(47, 'Pump Overall Efficiency',
    'eta = P_fluid / P_shaft = rho*g*Q*H / P_shaft',
    [('eta',      'Overall pump efficiency', 'dimensionless (0 to 1)'),
     ('P_fluid',  'Power delivered to fluid', 'W'),
     ('P_shaft',  'Shaft (brake) power input to pump', 'W'),
     ('rho',      'Fluid density', 'kg/m^3'),
     ('Q',        'Flow rate', 'm^3/s'),
     ('H',        'Total head', 'm')],
    example={
        'problem': 'Hydraulic power = 12.3 kW, shaft power input = 15 kW. Find eta.',
        'steps': ['eta = 12263/15000 = 0.818'],
        'answer': 'eta = 81.8%'})


# ══════════════════════════════════════════════════════════════════════════════
# ME 104/204  MECHANICS AND MATERIALS
# ══════════════════════════════════════════════════════════════════════════════
topic_hdr('ME 104 / 204', 'Mechanics and Materials')
subtopic_hdr('Stress, Strain and Elastic Constants')

eq_block(48, 'Direct (Normal) Stress',
    'sigma = F / A',
    [('sigma', 'Normal stress (positive = tension)', 'Pa'),
     ('F',     'Axial force (positive = tensile)', 'N'),
     ('A',     'Cross-sectional area', 'm^2')],
    example={
        'problem': 'A 25 mm diameter steel rod carries 50 kN tensile load. Find sigma.',
        'steps': [
            'A = pi*(0.025)^2/4 = 4.909e-4 m^2',
            'sigma = 50000 / 4.909e-4 = 101.9 MPa',
        ],
        'answer': 'sigma = 102 MPa (tensile)'})

eq_block(49, 'Direct (Normal) Strain',
    'eps = deltaL / L0',
    [('eps',    'Normal strain', 'dimensionless (m/m)'),
     ('deltaL', 'Change in length', 'm'),
     ('L0',     'Original (gauge) length', 'm')],
    example={
        'problem': 'A 500 mm rod elongates by 0.255 mm under load. Find eps.',
        'steps': ['eps = 0.255 / 500 = 5.10e-4 (510 microstrain)'],
        'answer': 'eps = 5.1e-4'})

eq_block(50, "Hooke's Law - Elastic Region",
    'sigma = E * eps',
    [('sigma', 'Normal stress', 'Pa'),
     ('E',     "Young's modulus (steel = 200 GPa)", 'Pa'),
     ('eps',   'Normal strain', 'dimensionless')],
    example={
        'problem': 'Steel rod (E=200 GPa) under sigma=102 MPa. Find eps.',
        'steps': ['eps = sigma/E = 102e6 / 200e9 = 5.1e-4'],
        'answer': 'eps = 5.1e-4 (510 microstrain)'})

eq_block(51, 'Shear Stress and Shear Strain',
    'tau = Fs / A     and     gamma = tau / G',
    [('tau',   'Shear stress', 'Pa'),
     ('Fs',    'Shear force', 'N'),
     ('A',     'Area on which shear acts', 'm^2'),
     ('gamma', 'Shear strain (engineering)', 'dimensionless (rad)'),
     ('G',     'Shear modulus (steel = 79 GPa)', 'Pa')],
    example={
        'problem': '20x20 mm section, 8 kN shear force, G=80 GPa. Find tau and gamma.',
        'steps': [
            'tau = 8000 / (0.020*0.020) = 20 MPa',
            'gamma = tau/G = 20e6 / 80e9 = 2.5e-4 rad',
        ],
        'answer': 'tau = 20 MPa, gamma = 2.5e-4 rad'})

eq_block(52, "Poisson's Ratio",
    'nu = -eps_lateral / eps_axial',
    [('nu',          "Poisson's ratio (steel = 0.30)", 'dimensionless'),
     ('eps_lateral', 'Transverse (lateral) strain', 'dimensionless'),
     ('eps_axial',   'Axial (longitudinal) strain', 'dimensionless')],
    example={
        'problem': 'Steel (nu=0.30) under axial strain eps_axial = 5.1e-4. Find lateral strain.',
        'steps': ['eps_lateral = -nu * eps_axial = -0.30 * 5.1e-4 = -1.53e-4'],
        'answer': 'eps_lateral = -1.53e-4 (contraction)'})

eq_block(53, 'Relationship Between Elastic Constants',
    'G = E / [2*(1 + nu)]     and     K = E / [3*(1 - 2*nu)]',
    [('G',  'Shear modulus', 'Pa'),
     ('E',  "Young's modulus", 'Pa'),
     ('nu', "Poisson's ratio", 'dimensionless'),
     ('K',  'Bulk modulus', 'Pa')],
    example={
        'problem': 'Steel: E = 200 GPa, nu = 0.30. Find G.',
        'steps': ['G = 200e9 / (2*(1+0.30)) = 200e9 / 2.60 = 76.9 GPa'],
        'answer': 'G = 76.9 GPa (literature: ~79 GPa)'})

subtopic_hdr('Thermal Effects')

eq_block(54, 'Free Thermal Strain',
    'eps_T = alpha * DeltaT',
    [('eps_T',   'Thermal strain (free expansion)', 'dimensionless'),
     ('alpha',   'Coefficient of linear thermal expansion (steel = 12e-6)', '/degC'),
     ('DeltaT',  'Temperature change', 'degC (or K)')],
    example={
        'problem': 'Steel rail (alpha=12e-6 /degC) heated by 40 degC. Find eps_T and extension of a 10 m rail.',
        'steps': [
            'eps_T = 12e-6 * 40 = 4.8e-4',
            'Extension = eps_T * L0 = 4.8e-4 * 10 = 4.8 mm',
        ],
        'answer': 'eps_T = 4.8e-4 (4.8 mm per 10 m rail)'})

eq_block(55, 'Constrained Thermal Stress',
    'sigma_T = E * alpha * DeltaT',
    [('sigma_T', 'Thermal stress (compressive if expansion prevented)', 'Pa'),
     ('E',       "Young's modulus", 'Pa'),
     ('alpha',   'Coefficient of thermal expansion', '/degC'),
     ('DeltaT',  'Temperature change', 'degC')],
    example={
        'problem': 'Same steel rail rigidly fixed at both ends. DeltaT = 40 degC. Find sigma_T.',
        'steps': ['sigma_T = 200e9 * 12e-6 * 40 = 96 MPa (compressive)'],
        'answer': 'sigma_T = 96 MPa compressive'})

subtopic_hdr('Bending of Beams')

eq_block(56, 'Second Moment of Area - Rectangle (about centroidal axis)',
    'Ixx = b*h^3 / 12',
    [('Ixx', 'Second moment of area about horizontal centroidal axis', 'm^4'),
     ('b',   'Width of rectangle', 'm'),
     ('h',   'Height (depth) of rectangle', 'm')],
    example={
        'problem': 'Beam cross-section: 60 mm wide x 120 mm deep. Find Ixx.',
        'steps': ['Ixx = 0.060 * (0.120)^3 / 12 = 0.060 * 1.728e-3 / 12 = 8.64e-6 m^4'],
        'answer': 'Ixx = 8.64e-6 m^4'})

eq_block(57, 'Second Moment of Area - Solid Circle',
    'I = pi*d^4 / 64 = pi*r^4 / 4',
    [('I', 'Second moment of area about diameter', 'm^4'),
     ('d', 'Diameter', 'm'),
     ('r', 'Radius', 'm')],
    example={
        'problem': '80 mm diameter solid shaft. Find I.',
        'steps': ['I = pi*(0.080)^4 / 64 = pi*4.096e-6/64 = 2.01e-7 m^4'],
        'answer': 'I = 2.01e-7 m^4'})

eq_block(58, 'Flexure Formula (Bending Stress - Euler-Bernoulli)',
    'sigma = M*y / I     =>     sigma_max = M / Z',
    [('sigma',     'Bending stress at distance y from neutral axis', 'Pa'),
     ('M',         'Bending moment', 'N*m'),
     ('y',         'Distance from neutral axis to point of interest', 'm'),
     ('I',         'Second moment of area about neutral axis', 'm^4'),
     ('sigma_max', 'Maximum bending stress (at extreme fibre)', 'Pa'),
     ('Z',         'Section modulus Z = I/y_max', 'm^3')],
    example={
        'problem': 'Simply supported beam, M_max=12 kN*m, 60x120 mm rectangle. Find sigma_max.',
        'steps': [
            'I = 8.64e-6 m^4 (from eq. 56)',
            'y_max = 120/2 mm = 0.060 m',
            'sigma_max = 12000 * 0.060 / 8.64e-6 = 83.3 MPa',
        ],
        'answer': 'sigma_max = 83.3 MPa (at extreme fibres)'})

eq_block(59, 'Shear Stress in Beams',
    'tau = V*Q / (I*b)',
    [('tau', 'Horizontal shear stress at depth y', 'Pa'),
     ('V',   'Shear force at the cross-section', 'N'),
     ('Q',   'First moment of area of section above y about neutral axis', 'm^3'),
     ('I',   'Second moment of area of full cross-section', 'm^4'),
     ('b',   'Width of cross-section at depth y', 'm')],
    example={
        'problem': '60x120 mm beam, V=20 kN. Find tau at the neutral axis.',
        'steps': [
            'Q_NA = (b * h/2) * (h/4) = 0.060 * 0.060 * 0.030 = 1.08e-4 m^3',
            'tau_NA = 20000*1.08e-4 / (8.64e-6*0.060) = 2.16 / 5.18e-7 = 4.17 MPa',
        ],
        'answer': 'tau_NA = 4.17 MPa'})

eq_block(60, 'Beam Deflection - Simply-Supported, Central Point Load',
    'delta_max = F*L^3 / (48*E*I)',
    [('delta_max', 'Maximum deflection at mid-span', 'm'),
     ('F',         'Point load at mid-span', 'N'),
     ('L',         'Span length', 'm'),
     ('E',         "Young's modulus", 'Pa'),
     ('I',         'Second moment of area', 'm^4')],
    example={
        'problem': '6 m span, central 20 kN load, 60x120 mm section (E=200 GPa). Find delta_max.',
        'steps': [
            'I = 8.64e-6 m^4',
            'delta_max = 20000*6^3 / (48*200e9*8.64e-6)',
            '          = 20000*216 / (82,944,000) = 4,320,000 / 82,944,000 = 0.0521 m',
        ],
        'answer': 'delta_max = 52.1 mm'})

eq_block(61, 'Beam Deflection - Cantilever, Tip Point Load',
    'delta_max = F*L^3 / (3*E*I)',
    [('delta_max', 'Maximum deflection at free end', 'm'),
     ('F',         'Point load at free end', 'N'),
     ('L',         'Cantilever length', 'm'),
     ('E',         "Young's modulus", 'Pa'),
     ('I',         'Second moment of area', 'm^4')],
    example={
        'problem': 'Cantilever 2 m long, 5 kN tip load, 60x120 mm section (E=200 GPa). Find delta_max.',
        'steps': [
            'delta_max = 5000*(2)^3 / (3*200e9*8.64e-6)',
            '          = 5000*8 / 5,184,000 = 40,000/5,184,000 = 0.00772 m',
        ],
        'answer': 'delta_max = 7.7 mm'})

subtopic_hdr('Torsion')

eq_block(62, 'Torsion Formula - Circular Shaft (Coulomb)',
    'tau = T*r / J     =>     tau_max = T*(d/2) / J',
    [('tau',     'Shear stress at radius r', 'Pa'),
     ('T',       'Applied torque', 'N*m'),
     ('r',       'Radial distance from shaft axis', 'm'),
     ('J',       'Polar second moment of area', 'm^4'),
     ('d',       'Shaft diameter', 'm')],
    example={
        'problem': 'Solid shaft d=40 mm, T=500 N*m. Find tau_max.',
        'steps': [
            'J = pi*(0.040)^4/32 = 2.513e-7 m^4',
            'tau_max = T*(d/2)/J = 500*0.020 / 2.513e-7 = 10/2.513e-7 = 39.8 MPa',
        ],
        'answer': 'tau_max = 39.8 MPa'})

eq_block(63, 'Polar Second Moment of Area - Solid Shaft',
    'J = pi*d^4 / 32',
    [('J', 'Polar second moment of area', 'm^4'),
     ('d', 'Shaft diameter', 'm')],
    example={
        'problem': 'Find J for a 40 mm diameter solid shaft.',
        'steps': ['J = pi*(0.040)^4/32 = pi*2.56e-6/32 = 2.513e-7 m^4'],
        'answer': 'J = 2.513e-7 m^4'})

eq_block(64, 'Polar Second Moment of Area - Hollow Shaft',
    'J = pi*(do^4 - di^4) / 32',
    [('J',  'Polar second moment of area', 'm^4'),
     ('do', 'External (outer) diameter', 'm'),
     ('di', 'Internal (inner) diameter', 'm')],
    example={
        'problem': 'Hollow shaft: do=60 mm, di=40 mm. Find J.',
        'steps': [
            'J = pi*((0.060)^4 - (0.040)^4)/32',
            '  = pi*(1.296e-5 - 2.560e-6)/32 = pi*1.040e-5/32 = 1.021e-6 m^4',
        ],
        'answer': 'J = 1.021e-6 m^4'})

eq_block(65, 'Angle of Twist',
    'phi = T*L / (G*J)',
    [('phi', 'Angle of twist', 'rad'),
     ('T',   'Applied torque', 'N*m'),
     ('L',   'Shaft length', 'm'),
     ('G',   'Shear modulus', 'Pa'),
     ('J',   'Polar second moment of area', 'm^4')],
    example={
        'problem': '40 mm solid shaft (G=79 GPa), L=1 m, T=500 N*m. Find phi.',
        'steps': [
            'J = 2.513e-7 m^4',
            'phi = 500*1 / (79e9*2.513e-7) = 500/19,853 = 0.02519 rad = 1.44 deg',
        ],
        'answer': 'phi = 0.0252 rad (1.44 deg)'})

subtopic_hdr("Yield Criteria and Stress Transformations")

eq_block(66, "Principal Stresses - Mohr's Circle",
    'sigma1,2 = (sx+sy)/2 +/- sqrt[((sx-sy)/2)^2 + txy^2]',
    [('sigma1, sigma2', 'Major and minor principal stresses', 'Pa'),
     ('sx',            'Normal stress on x-face', 'Pa'),
     ('sy',            'Normal stress on y-face', 'Pa'),
     ('txy',           'Shear stress on x-face', 'Pa')],
    example={
        'problem': 'sx=80 MPa, sy=20 MPa, txy=30 MPa. Find principal stresses.',
        'steps': [
            'Centre C = (80+20)/2 = 50 MPa',
            'R = sqrt[((80-20)/2)^2 + 30^2] = sqrt[900+900] = sqrt(1800) = 42.4 MPa',
            'sigma1 = 50+42.4 = 92.4 MPa;  sigma2 = 50-42.4 = 7.6 MPa',
        ],
        'answer': 'sigma1 = 92.4 MPa, sigma2 = 7.6 MPa'})

eq_block(67, 'Maximum Shear Stress',
    'tau_max = sqrt[((sx-sy)/2)^2 + txy^2]  =  (sigma1 - sigma2)/2',
    [('tau_max',     'Maximum in-plane shear stress', 'Pa'),
     ('sx, sy',      'Normal stresses on x- and y-faces', 'Pa'),
     ('txy',         'Shear stress', 'Pa'),
     ('sigma1, sigma2', 'Principal stresses', 'Pa')],
    example={
        'problem': 'Using results from eq. 66: sigma1=92.4 MPa, sigma2=7.6 MPa. Find tau_max.',
        'steps': ['tau_max = (sigma1-sigma2)/2 = (92.4-7.6)/2 = 84.8/2 = 42.4 MPa'],
        'answer': 'tau_max = 42.4 MPa'})

eq_block(68, 'Von Mises Yield Criterion (2-D)',
    'sigma_VM = sqrt(sigma1^2 - sigma1*sigma2 + sigma2^2) <= sigma_y',
    [('sigma_VM',      'Von Mises equivalent stress (distortion energy criterion)', 'Pa'),
     ('sigma1, sigma2', 'Principal stresses', 'Pa'),
     ('sigma_y',       'Uniaxial yield strength of material', 'Pa')],
    example={
        'problem': 'sigma1=92.4 MPa, sigma2=7.6 MPa. Material sigma_y=250 MPa. Does it yield?',
        'steps': [
            'sigma_VM = sqrt(92.4^2 - 92.4*7.6 + 7.6^2)',
            '         = sqrt(8537.8 - 702.2 + 57.8) = sqrt(7893) = 88.8 MPa',
            '88.8 MPa < 250 MPa  =>  no yielding',
        ],
        'answer': 'sigma_VM = 88.8 MPa < 250 MPa => safe (FoS = 2.82)'})

eq_block(69, 'Tresca Yield Criterion',
    'tau_max = (sigma1 - sigma2)/2 <= sigma_y/2',
    [('tau_max',       'Maximum shear stress', 'Pa'),
     ('sigma1, sigma2', 'Principal stresses (sigma1 >= sigma2)', 'Pa'),
     ('sigma_y',       'Yield strength', 'Pa')],
    note='Tresca is more conservative (lower predicted failure load) than Von Mises.',
    example={
        'problem': 'tau_max=42.4 MPa (from eq. 67). sigma_y=250 MPa. Check Tresca criterion.',
        'steps': [
            'sigma_y/2 = 250/2 = 125 MPa',
            '42.4 MPa < 125 MPa  =>  no yielding',
        ],
        'answer': 'Safe. Tresca FoS = 125/42.4 = 2.95'})

subtopic_hdr('Pressure Vessels')

eq_block(70, 'Thin-Walled Cylinder - Hoop (Circumferential) Stress',
    'sigma_h = P*d / (2*t)',
    [('sigma_h', 'Hoop (circumferential) stress (= 2 x longitudinal)', 'Pa'),
     ('P',       'Internal gauge pressure', 'Pa'),
     ('d',       'Internal diameter', 'm'),
     ('t',       'Wall thickness', 'm')],
    note='Valid when d/t >= 20 (thin-wall assumption).',
    example={
        'problem': 'Pressure vessel: d=400 mm, t=10 mm, P=2 MPa. Find sigma_h.',
        'steps': ['sigma_h = 2e6*0.4 / (2*0.010) = 800,000/0.020 = 40 MPa'],
        'answer': 'sigma_h = 40 MPa (hoop)'})

eq_block(71, 'Thin-Walled Cylinder - Longitudinal (Axial) Stress',
    'sigma_L = P*d / (4*t)',
    [('sigma_L', 'Longitudinal (axial) stress (sigma_L = sigma_h/2)', 'Pa'),
     ('P',       'Internal gauge pressure', 'Pa'),
     ('d',       'Internal diameter', 'm'),
     ('t',       'Wall thickness', 'm')],
    example={
        'problem': 'Same vessel as eq. 70: d=400 mm, t=10 mm, P=2 MPa. Find sigma_L.',
        'steps': ['sigma_L = 2e6*0.4 / (4*0.010) = 800,000/0.040 = 20 MPa'],
        'answer': 'sigma_L = 20 MPa (= sigma_h/2)'})

eq_block(72, 'Thick-Walled Cylinder - Lame Equations',
    'sigma_theta = A + B/r^2     (hoop)     and     sigma_r = A - B/r^2     (radial)',
    [('sigma_theta', 'Hoop (tangential) stress at radius r', 'Pa'),
     ('sigma_r',     'Radial stress at radius r', 'Pa'),
     ('A, B',        'Lame constants from boundary conditions', 'Pa,  Pa*m^2'),
     ('r',           'Radial position', 'm')],
    example={
        'problem': 'Cylinder ri=50 mm, ro=100 mm, internal P=60 MPa, external P=0. Find sigma_theta at inner wall.',
        'steps': [
            'BC: sigma_r(ri) = A - B/ri^2 = -60 MPa;  sigma_r(ro) = A - B/ro^2 = 0',
            'From 2nd: A = B/ro^2 = B/0.01;  Sub: B/0.01 - B/0.0025 = -60  =>  -300B = -60',
            'B = 0.2 MPa*m^2;  A = 0.2/0.01 = 20 MPa',
            'sigma_theta(ri) = 20 + 0.2/0.0025 = 20 + 80 = 100 MPa',
        ],
        'answer': 'sigma_theta = 100 MPa at inner wall (maximum hoop stress)'})

subtopic_hdr('Buckling')

eq_block(73, 'Euler Critical Buckling Load',
    'P_cr = pi^2*E*I / Le^2',
    [('P_cr', 'Critical (Euler) buckling load', 'N'),
     ('E',    "Young's modulus", 'Pa'),
     ('I',    'Minimum second moment of area of cross-section', 'm^4'),
     ('Le',   'Effective length (Le = K*L, K depends on end conditions)', 'm')],
    note='End condition K: both-pinned=1.0; one-fixed-one-free=2.0; both-fixed=0.5; fixed-pinned=0.7.',
    example={
        'problem': 'Pin-pin column, 50x50 mm square, L=3 m, E=200 GPa. Find P_cr.',
        'steps': [
            'I = 50^4/12 mm^4 = 520,833 mm^4 = 5.208e-7 m^4',
            'Le = K*L = 1.0*3 = 3 m',
            'P_cr = pi^2*200e9*5.208e-7 / 3^2 = 1,027,600/9 = 114,178 N',
        ],
        'answer': 'P_cr = 114 kN'})


# ══════════════════════════════════════════════════════════════════════════════
# ME 105/205  MANUFACTURING TECHNOLOGY
# ══════════════════════════════════════════════════════════════════════════════
topic_hdr('ME 105 / 205', 'Manufacturing Technology')
subtopic_hdr('Metal Cutting - Tool Life')

eq_block(74, "Taylor's Tool Life Equation",
    'V * T^n = C',
    [('V', 'Cutting speed', 'm/min'),
     ('T', 'Tool life', 'min'),
     ('n', "Taylor's exponent (HSS=0.1-0.15; carbide=0.2-0.3; ceramic=0.4-0.5)", 'dimensionless'),
     ('C', "Taylor constant (cutting speed for T=1 min)", 'm/min')],
    note='Rearranged: T = (C/V)^(1/n). Log form: log V + n*log T = log C (straight line on log-log plot).',
    example={
        'problem': 'At V=200 m/min, T=60 min; n=0.25. Find T at V=250 m/min.',
        'steps': [
            'C = V*T^n = 200*60^0.25 = 200*2.783 = 556.6 m/min',
            'T = (C/V)^(1/n) = (556.6/250)^(1/0.25) = (2.226)^4 = 24.5 min',
        ],
        'answer': 'T = 24.5 min at V=250 m/min (41% of original life)'})

subtopic_hdr('Metal Cutting - Cutting Geometry')

eq_block(75, 'Chip Thickness Ratio',
    'rc = t1 / t2',
    [('rc', 'Chip thickness ratio (cutting ratio), 0 < rc < 1', 'dimensionless'),
     ('t1', 'Uncut chip thickness (depth of cut)', 'mm'),
     ('t2', 'Actual (deformed) chip thickness', 'mm')],
    example={
        'problem': 'Uncut chip t1=0.25 mm, measured chip t2=0.75 mm. Find rc.',
        'steps': ['rc = t1/t2 = 0.25/0.75 = 0.333'],
        'answer': 'rc = 0.333'})

eq_block(76, 'Shear Plane Angle',
    'tan(phi) = rc*cos(alpha) / (1 - rc*sin(alpha))',
    [('phi',   'Shear plane angle', 'deg'),
     ('rc',    'Chip thickness ratio', 'dimensionless'),
     ('alpha', 'Tool rake angle (positive if tilted toward workpiece)', 'deg')],
    example={
        'problem': 'rc=0.333, rake angle alpha=10 deg. Find shear plane angle phi.',
        'steps': [
            'tan(phi) = 0.333*cos(10)/(1 - 0.333*sin(10))',
            '         = 0.333*0.9848 / (1 - 0.333*0.1736)',
            '         = 0.3280 / 0.9422 = 0.3481',
            'phi = arctan(0.3481) = 19.2 deg',
        ],
        'answer': 'phi = 19.2 deg'})

eq_block(77, "Merchant's Minimum Energy Equation",
    'phi = 45 + alpha/2 - lambda/2',
    [('phi',    'Shear plane angle', 'deg'),
     ('alpha',  'Tool rake angle', 'deg'),
     ('lambda', 'Friction angle  lambda = arctan(mu)', 'deg')],
    note="Based on minimum energy principle. Predicts optimal shear angle for given rake and friction.",
    example={
        'problem': 'alpha=10 deg, friction coefficient mu=0.5. Find Merchant prediction for phi.',
        'steps': [
            'lambda = arctan(mu) = arctan(0.5) = 26.6 deg',
            'phi = 45 + 10/2 - 26.6/2 = 45 + 5 - 13.3 = 36.7 deg',
        ],
        'answer': "phi = 36.7 deg (Merchant's prediction)"})

subtopic_hdr('Material Removal - Turning')

eq_block(78, 'Cutting Speed',
    'Vc = pi*D*N / 1000',
    [('Vc', 'Peripheral cutting speed', 'm/min'),
     ('D',  'Workpiece diameter', 'mm'),
     ('N',  'Spindle speed', 'rpm')],
    example={
        'problem': 'Turning 80 mm diameter steel bar at N=800 rpm. Find Vc.',
        'steps': ['Vc = pi*80*800/1000 = 201.1 m/min'],
        'answer': 'Vc = 201 m/min'})

eq_block(79, 'Material Removal Rate - Turning',
    'MRR = pi*D*N*f*d = Vc*f*d',
    [('MRR', 'Material removal rate', 'mm^3/min'),
     ('D',   'Workpiece diameter', 'mm'),
     ('N',   'Spindle speed', 'rpm'),
     ('f',   'Feed per revolution', 'mm/rev'),
     ('d',   'Depth of cut (radial)', 'mm'),
     ('Vc',  'Cutting speed', 'm/min (use mm/min = m/min * 1000)')],
    example={
        'problem': 'D=80 mm, N=800 rpm, f=0.25 mm/rev, d=2 mm. Find MRR.',
        'steps': [
            'MRR = Vc*f*d = 201,100 mm/min * 0.25 * 2',
            '    = 100,550 mm^3/min',
        ],
        'answer': 'MRR = 100,550 mm^3/min (= 100.6 cm^3/min)'})

subtopic_hdr('Material Removal - Milling')

eq_block(80, 'Feed Rate - Milling',
    'vf = fz * z * N',
    [('vf', 'Table feed rate', 'mm/min'),
     ('fz', 'Feed per tooth', 'mm/tooth'),
     ('z',  'Number of cutter teeth', 'dimensionless'),
     ('N',  'Spindle speed', 'rpm')],
    example={
        'problem': 'fz=0.08 mm/tooth, z=6 teeth, N=400 rpm. Find table feed rate.',
        'steps': ['vf = 0.08 * 6 * 400 = 192 mm/min'],
        'answer': 'vf = 192 mm/min'})

eq_block(81, 'Material Removal Rate - Face Milling',
    'MRR = vf * d * W',
    [('MRR', 'Material removal rate', 'mm^3/min'),
     ('vf',  'Feed rate', 'mm/min'),
     ('d',   'Axial depth of cut', 'mm'),
     ('W',   'Width of cut (radial engagement)', 'mm')],
    example={
        'problem': 'vf=192 mm/min, axial depth d=3 mm, width W=50 mm. Find MRR.',
        'steps': ['MRR = 192 * 3 * 50 = 28,800 mm^3/min'],
        'answer': 'MRR = 28,800 mm^3/min'})

subtopic_hdr('Machining Power and Energy')

eq_block(82, 'Cutting Power',
    'P = Fc * Vc',
    [('P',  'Cutting power', 'W'),
     ('Fc', 'Principal (tangential) cutting force', 'N'),
     ('Vc', 'Cutting speed', 'm/s')],
    example={
        'problem': 'Fc=800 N, Vc=201 m/min = 3.35 m/s. Find cutting power.',
        'steps': ['P = 800 * 3.35 = 2,680 W'],
        'answer': 'P = 2.68 kW'})

eq_block(83, 'Specific Cutting Energy (Unit Power)',
    'u = P / MRR = Fc / (f*d)',
    [('u',   'Specific cutting energy', 'J/mm^3'),
     ('P',   'Cutting power', 'W'),
     ('MRR', 'Material removal rate', 'mm^3/s'),
     ('Fc',  'Cutting force', 'N'),
     ('f',   'Feed', 'mm/rev'),
     ('d',   'Depth of cut', 'mm')],
    example={
        'problem': 'P=2680 W, MRR=100,550 mm^3/min = 1676 mm^3/s. Find specific cutting energy.',
        'steps': ['u = P/MRR = 2680/1676 = 1.60 J/mm^3'],
        'answer': 'u = 1.60 J/mm^3 (= 1.60 GJ/m^3)'})

subtopic_hdr('Rolling and Forming')

eq_block(84, 'Rolling - Draft',
    'd = t0 - tf',
    [('d',  'Draft (reduction in thickness)', 'mm'),
     ('t0', 'Initial sheet/slab thickness', 'mm'),
     ('tf', 'Final sheet/slab thickness', 'mm')],
    example={
        'problem': 'Steel slab reduced from 25 mm to 18 mm in one pass. Find draft and % reduction.',
        'steps': [
            'd = 25 - 18 = 7 mm',
            '% reduction = 7/25 * 100 = 28%',
        ],
        'answer': 'd = 7 mm (28% reduction per pass)'})

eq_block(85, 'Deep Drawing - Approximate Blank Diameter',
    'Db = sqrt(Dp^2 + 4*Dp*h)',
    [('Db', 'Blank diameter required', 'mm'),
     ('Dp', 'Punch (cup) diameter', 'mm'),
     ('h',  'Cup drawing depth', 'mm')],
    note='Approximate formula based on surface area conservation.',
    example={
        'problem': 'Cup Dp=80 mm, drawing depth h=50 mm. Find required blank diameter.',
        'steps': [
            'Db = sqrt(80^2 + 4*80*50) = sqrt(6400 + 16000) = sqrt(22400) = 149.7 mm',
        ],
        'answer': 'Db = 150 mm'})


# ══════════════════════════════════════════════════════════════════════════════
# ME 106/206  THERMODYNAMICS AND HEAT TRANSFER
# ══════════════════════════════════════════════════════════════════════════════
topic_hdr('ME 106 / 206', 'Thermodynamics and Heat Transfer')
subtopic_hdr('Ideal Gas Laws')

eq_block(86, 'Ideal Gas Law - Molar Form',
    'P*V = n*Rbar*T',
    [('P',    'Absolute pressure', 'Pa'),
     ('V',    'Volume', 'm^3'),
     ('n',    'Amount of substance', 'mol'),
     ('Rbar', 'Universal gas constant = 8.314', 'J/(mol*K)'),
     ('T',    'Absolute temperature', 'K')],
    example={
        'problem': '2 mol of ideal gas at T=300 K, P=200 kPa. Find volume.',
        'steps': ['V = n*Rbar*T/P = 2*8.314*300 / 200,000 = 4988.4/200,000 = 0.0249 m^3'],
        'answer': 'V = 0.0249 m^3 = 24.9 L'})

eq_block(87, 'Ideal Gas Law - Mass Form',
    'P*V = m*R*T     or     P = rho*R*T',
    [('P',   'Absolute pressure', 'Pa'),
     ('V',   'Volume', 'm^3'),
     ('m',   'Mass of gas', 'kg'),
     ('R',   'Specific gas constant R=Rbar/M (air=287)', 'J/(kg*K)'),
     ('T',   'Absolute temperature', 'K'),
     ('rho', 'Gas density', 'kg/m^3')],
    example={
        'problem': '1 kg of air at P=100 kPa, T=20 degC=293 K. Find volume.',
        'steps': ['V = m*R*T/P = 1*287*293/100,000 = 84,091/100,000 = 0.841 m^3'],
        'answer': 'V = 0.841 m^3'})

eq_block(88, 'Specific Heat Relationship - Ideal Gas (Mayer Relation)',
    'cp - cv = R',
    [('cp', 'Specific heat at constant pressure', 'J/(kg*K)'),
     ('cv', 'Specific heat at constant volume', 'J/(kg*K)'),
     ('R',  'Specific gas constant', 'J/(kg*K)')],
    example={
        'problem': 'Air: cp=1005 J/(kg*K), R=287 J/(kg*K). Find cv.',
        'steps': ['cv = cp - R = 1005 - 287 = 718 J/(kg*K)'],
        'answer': 'cv = 718 J/(kg*K)'})

eq_block(89, 'Heat Capacity Ratio',
    'gamma = cp / cv',
    [('gamma', 'Heat capacity ratio (monatomic~1.67; diatomic/air~1.40)', 'dimensionless'),
     ('cp',    'Specific heat at constant pressure', 'J/(kg*K)'),
     ('cv',    'Specific heat at constant volume', 'J/(kg*K)')],
    example={
        'problem': 'Air: cp=1005, cv=718 J/(kg*K). Find gamma.',
        'steps': ['gamma = cp/cv = 1005/718 = 1.40'],
        'answer': 'gamma = 1.40 (diatomic gas, as expected for air)'})

subtopic_hdr('First Law of Thermodynamics')

eq_block(90, 'First Law - Closed System',
    'DeltaU = Q - W',
    [('DeltaU', 'Change in internal energy of the system', 'J'),
     ('Q',      'Heat transferred into the system (Q>0 in; Q<0 out)', 'J'),
     ('W',      'Work done by the system (W>0 out; W<0 in)', 'J')],
    example={
        'problem': 'A gas absorbs Q=500 J heat and does W=200 J work. Find DeltaU.',
        'steps': ['DeltaU = Q - W = 500 - 200 = 300 J'],
        'answer': 'DeltaU = 300 J (internal energy increases)'})

eq_block(91, 'Steady Flow Energy Equation (SFEE) - Open System',
    'Qdot - Wdot = mdot * [(h2-h1) + (1/2)*(V2^2-V1^2) + g*(z2-z1)]',
    [('Qdot',    'Rate of heat transfer into system', 'W'),
     ('Wdot',    'Rate of work output (shaft work)', 'W'),
     ('mdot',    'Mass flow rate', 'kg/s'),
     ('h1, h2',  'Specific enthalpy at inlet and outlet', 'J/kg'),
     ('V1, V2',  'Velocities at inlet and outlet', 'm/s'),
     ('z1, z2',  'Elevations at inlet and outlet', 'm'),
     ('g',       'Gravitational acceleration', 'm/s^2')],
    example={
        'problem': 'Steam turbine: mdot=5 kg/s, h1=3200 kJ/kg, h2=2400 kJ/kg, Qdot=0, KE/PE negligible. Find Wdot.',
        'steps': [
            '0 - Wdot = 5*(2400e3 - 3200e3) = 5*(-800e3) = -4,000,000 W',
            'Wdot = +4,000,000 W',
        ],
        'answer': 'Wdot = 4.0 MW (turbine power output)'})

subtopic_hdr('Second Law and Cycles')

eq_block(92, 'Carnot Efficiency (Maximum Theoretical Efficiency)',
    'eta_c = 1 - TL / TH',
    [('eta_c', 'Carnot (maximum) thermal efficiency', 'dimensionless'),
     ('TL',    'Temperature of cold reservoir (condenser)', 'K (must be absolute)'),
     ('TH',    'Temperature of hot reservoir (boiler)', 'K')],
    note='eta_c = W_net/Q_H = 1 - Q_C/Q_H. Minimum fuel consumption implies Carnot efficiency.',
    example={
        'problem': 'Steam cycle: TH=560 degC=833 K, TL=40 degC=313 K. Find Carnot efficiency.',
        'steps': ['eta_c = 1 - TL/TH = 1 - 313/833 = 1 - 0.376 = 0.624'],
        'answer': 'eta_c = 62.4%'})

eq_block(93, 'Entropy Change (Reversible Process)',
    'dS = dQ_rev / T     =>     DeltaS = integral(dQ/T)_rev',
    [('dS',      'Infinitesimal entropy change', 'J/K'),
     ('dQ_rev',  'Infinitesimal reversible heat transfer', 'J'),
     ('T',       'Absolute temperature at boundary', 'K'),
     ('DeltaS',  'Total entropy change', 'J/K')],
    example={
        'problem': '2 kg of water vaporised reversibly at 100 degC (373 K), h_fg=2257 kJ/kg. Find DeltaS.',
        'steps': [
            'Q = 2 * 2257e3 = 4,514,000 J',
            'DeltaS = Q/T = 4,514,000/373 = 12,104 J/K',
        ],
        'answer': 'DeltaS = 12.1 kJ/K'})

eq_block(94, 'Clausius Inequality',
    'closed_integral( dQ / T ) <= 0',
    [('closed_integral dQ/T', 'Cyclic integral of heat/temperature over a complete cycle', 'J/K')],
    note='= 0 for reversible cycle; < 0 for irreversible (real) cycle. Expression of 2nd Law.',
    example={
        'problem': 'Heat engine: QH=500 J at 800 K, rejects QL=350 J at 300 K. Evaluate cyclic integral.',
        'steps': [
            'sum(dQ/T) = QH/TH - QL/TL = 500/800 - 350/300 = 0.625 - 1.167 = -0.542 J/K',
        ],
        'answer': 'Cyclic integral = -0.54 J/K < 0 => irreversible (real) engine'})

eq_block(95, 'Isentropic Process Relations - Ideal Gas',
    'T2/T1 = (P2/P1)^((gamma-1)/gamma) = (V1/V2)^(gamma-1)',
    [('T1, T2', 'Temperatures before and after isentropic process', 'K'),
     ('P1, P2', 'Pressures before and after', 'Pa'),
     ('V1, V2', 'Specific volumes before and after', 'm^3/kg'),
     ('gamma',  'Heat capacity ratio cp/cv', 'dimensionless')],
    note='Valid for reversible adiabatic (isentropic) processes only.',
    example={
        'problem': 'Air compressed isentropically: P1=100 kPa, T1=300 K to P2=600 kPa. Find T2.',
        'steps': [
            'T2 = T1 * (P2/P1)^((gamma-1)/gamma) = 300 * (6)^(0.4/1.4)',
            '   = 300 * 6^0.2857 = 300 * 1.669 = 500.7 K',
        ],
        'answer': 'T2 = 501 K = 228 degC'})

subtopic_hdr('Heat Transfer - Conduction')

eq_block(96, "Fourier's Law of Heat Conduction",
    'q = -k*A*(dT/dx)     or     q_flux = -k*(dT/dx)',
    [('q',       'Rate of heat transfer', 'W'),
     ('q_flux',  'Heat flux', 'W/m^2'),
     ('k',       'Thermal conductivity of material', 'W/(m*K)'),
     ('A',       'Cross-sectional area perpendicular to heat flow', 'm^2'),
     ('dT/dx',   'Temperature gradient in direction of heat flow', 'K/m')],
    note='Negative sign: heat flows in the direction of decreasing temperature.',
    example={
        'problem': 'Copper wall (k=400 W/(m*K)), A=0.5 m^2, thickness L=10 mm, DeltaT=50 degC. Find q.',
        'steps': ['q = k*A*(DeltaT/L) = 400 * 0.5 * (50/0.010) = 400*0.5*5000 = 1,000,000 W'],
        'answer': 'q = 1.0 MW'})

eq_block(97, 'Thermal Resistance - Conduction (Flat Slab)',
    'R_cond = L / (k*A)',
    [('R_cond', 'Conductive thermal resistance', 'K/W'),
     ('L',      'Slab thickness', 'm'),
     ('k',      'Thermal conductivity', 'W/(m*K)'),
     ('A',      'Area', 'm^2')],
    note='Analogous to Ohm law: Q_dot = DeltaT / R_cond',
    example={
        'problem': 'Copper wall from eq.96: L=10 mm, k=400, A=0.5 m^2. Find R_cond.',
        'steps': ['R_cond = L/(k*A) = 0.010 / (400*0.5) = 0.010/200 = 5.0e-5 K/W'],
        'answer': 'R_cond = 5.0e-5 K/W (very low - copper is excellent conductor)'})

subtopic_hdr('Heat Transfer - Convection')

eq_block(98, "Newton's Law of Cooling (Convection)",
    'q = h*A*(Ts - T_inf)',
    [('q',     'Convective heat transfer rate', 'W'),
     ('h',     'Convective heat transfer coefficient', 'W/(m^2*K)'),
     ('A',     'Surface area in contact with fluid', 'm^2'),
     ('Ts',    'Surface temperature', 'K (or degC)'),
     ('T_inf', 'Free-stream (ambient) fluid temperature', 'K (or degC)')],
    example={
        'problem': 'Flat plate A=0.2 m^2, Ts=80 degC, air at T_inf=25 degC, h=35 W/(m^2*K). Find q.',
        'steps': ['q = 35 * 0.2 * (80 - 25) = 35 * 0.2 * 55 = 385 W'],
        'answer': 'q = 385 W'})

eq_block(99, 'Thermal Resistance - Convection',
    'R_conv = 1 / (h*A)',
    [('R_conv', 'Convective thermal resistance', 'K/W'),
     ('h',      'Convective heat transfer coefficient', 'W/(m^2*K)'),
     ('A',      'Surface area', 'm^2')],
    example={
        'problem': 'Same plate: h=35 W/(m^2*K), A=0.2 m^2. Find R_conv.',
        'steps': ['R_conv = 1/(h*A) = 1/(35*0.2) = 1/7 = 0.143 K/W'],
        'answer': 'R_conv = 0.143 K/W'})

eq_block(100, 'Overall Heat Transfer Coefficient (Plane Wall)',
    '1/(U*A) = 1/(h1*A) + L/(k*A) + 1/(h2*A) = R_total',
    [('U',       'Overall heat transfer coefficient', 'W/(m^2*K)'),
     ('A',       'Heat transfer area', 'm^2'),
     ('h1, h2',  'Convection coefficients on each side of wall', 'W/(m^2*K)'),
     ('L',       'Wall thickness', 'm'),
     ('k',       'Thermal conductivity of wall', 'W/(m*K)'),
     ('R_total', 'Total thermal resistance', 'K/W')],
    example={
        'problem': 'Wall: h1=50, L=20 mm, k=1.0, h2=20 W/(m^2*K), A=1 m^2. Find U.',
        'steps': [
            '1/U = 1/50 + 0.020/1.0 + 1/20 = 0.020 + 0.020 + 0.050 = 0.090 m^2*K/W',
            'U = 1/0.090 = 11.1 W/(m^2*K)',
        ],
        'answer': 'U = 11.1 W/(m^2*K)'})

eq_block(101, 'Nusselt Number',
    'Nu = h*L / k',
    [('Nu', 'Nusselt number (ratio of convective to conductive heat transfer)', 'dimensionless'),
     ('h',  'Convective heat transfer coefficient', 'W/(m^2*K)'),
     ('L',  'Characteristic length', 'm'),
     ('k',  'Fluid thermal conductivity', 'W/(m*K)')],
    example={
        'problem': 'Pipe D=20 mm, h=2000 W/(m^2*K). Air k_fluid=0.026 W/(m*K). Find Nu.',
        'steps': ['Nu = h*L/k = 2000*0.020/0.026 = 40/0.026 = 1538'],
        'answer': 'Nu = 1538 (turbulent forced convection in pipe)'})

eq_block(102, 'Prandtl Number',
    'Pr = mu*cp / k = nu / alpha_th',
    [('Pr',       'Prandtl number (ratio of momentum to thermal diffusivity)', 'dimensionless'),
     ('mu',       'Dynamic viscosity of fluid', 'Pa*s'),
     ('cp',       'Specific heat of fluid at constant pressure', 'J/(kg*K)'),
     ('k',        'Fluid thermal conductivity', 'W/(m*K)'),
     ('nu',       'Kinematic viscosity', 'm^2/s'),
     ('alpha_th', 'Thermal diffusivity alpha=k/(rho*cp)', 'm^2/s')],
    example={
        'problem': 'Engine oil: mu=0.08 Pa*s, cp=1900 J/(kg*K), k=0.14 W/(m*K). Find Pr.',
        'steps': ['Pr = mu*cp/k = 0.08*1900/0.14 = 152/0.14 = 1086'],
        'answer': 'Pr = 1086 (highly viscous - thick thermal boundary layer)'})

eq_block(103, 'Dittus-Boelter Correlation (Turbulent Flow in Pipes)',
    'Nu = 0.023 * Re^0.8 * Pr^n',
    [('Nu', 'Nusselt number', 'dimensionless'),
     ('Re', 'Reynolds number', 'dimensionless'),
     ('Pr', 'Prandtl number', 'dimensionless'),
     ('n',  'n=0.4 for fluid being heated;  n=0.3 for fluid being cooled', 'dimensionless')],
    note='Valid for: Re > 10,000; 0.6 < Pr < 160; L/D > 10 (thermally/hydrodynamically developed).',
    example={
        'problem': 'Water in pipe: Re=50,000, Pr=7.0 (fluid being heated). Find Nu.',
        'steps': [
            'Nu = 0.023 * Re^0.8 * Pr^0.4',
            '   = 0.023 * (50000)^0.8 * 7.0^0.4',
            '   = 0.023 * 4521 * 2.055 = 213.4',
        ],
        'answer': 'Nu = 213'})

eq_block(104, 'Fin Efficiency',
    'eta_f = tanh(m*L) / (m*L)     where  m = sqrt(h*P / (k*Ac))',
    [('eta_f', 'Fin efficiency (actual / maximum possible heat transfer)', 'dimensionless'),
     ('m',     'Fin parameter', 'm^-1'),
     ('L',     'Fin length', 'm'),
     ('h',     'Convective heat transfer coefficient at fin surface', 'W/(m^2*K)'),
     ('P',     'Fin perimeter (cross-sectional)', 'm'),
     ('k',     'Fin material thermal conductivity', 'W/(m*K)'),
     ('Ac',    'Fin cross-sectional area', 'm^2')],
    note='Valid for insulated fin tip. Uniform cross-section.',
    example={
        'problem': 'Aluminium fin (k=200) rectangular 2mm wide x 20mm deep, L=40mm, h=80 W/(m^2*K). Find eta_f.',
        'steps': [
            'Ac = 0.002*0.020 = 4.0e-5 m^2;  P = 2*(0.002+0.020) = 0.044 m',
            'm = sqrt(80*0.044/(200*4.0e-5)) = sqrt(3.52/0.008) = sqrt(440) = 20.98 m^-1',
            'm*L = 20.98*0.040 = 0.839',
            'eta_f = tanh(0.839)/0.839 = 0.685/0.839 = 0.817',
        ],
        'answer': 'eta_f = 81.7%'})

eq_block(105, 'Log Mean Temperature Difference (LMTD) - Heat Exchangers',
    'LMTD = (DeltaT1 - DeltaT2) / ln(DeltaT1/DeltaT2)',
    [('LMTD',   'Log mean temperature difference', 'K (or degC)'),
     ('DeltaT1', 'Temperature difference at one end of heat exchanger', 'K'),
     ('DeltaT2', 'Temperature difference at other end of heat exchanger', 'K')],
    note='Used in Q_dot = U*A*LMTD. For counter-flow HX use counter-flow temperature pairings.',
    example={
        'problem': 'Counter-flow HX: hot fluid 90->50 degC, cold fluid 20->70 degC. Find LMTD.',
        'steps': [
            'DeltaT1 = 90-70 = 20 degC (hot inlet vs cold outlet)',
            'DeltaT2 = 50-20 = 30 degC (hot outlet vs cold inlet)',
            'LMTD = (20-30)/ln(20/30) = -10/(-0.405) = 24.7 degC',
        ],
        'answer': 'LMTD = 24.7 degC'})

subtopic_hdr('Heat Transfer - Radiation')

eq_block(106, 'Stefan-Boltzmann Law - Net Radiation Heat Transfer',
    'q = eps*sigma*A*(Ts^4 - T_surr^4)',
    [('q',      'Net radiation heat transfer rate', 'W'),
     ('eps',    'Emissivity (0 <= eps <= 1; blackbody eps=1)', 'dimensionless'),
     ('sigma',  'Stefan-Boltzmann constant = 5.670e-8', 'W/(m^2*K^4)'),
     ('A',      'Surface area', 'm^2'),
     ('Ts',     'Absolute surface temperature', 'K'),
     ('T_surr', 'Absolute temperature of surroundings', 'K')],
    example={
        'problem': 'Blackbody (eps=1) at 600 degC=873 K, A=0.1 m^2, surroundings at 20 degC=293 K. Find q.',
        'steps': [
            'q = 1 * 5.670e-8 * 0.1 * (873^4 - 293^4)',
            '  = 5.670e-9 * (5.819e11 - 7.370e9)',
            '  = 5.670e-9 * 5.745e11 = 3258 W',
        ],
        'answer': 'q = 3.26 kW'})

eq_block(107, "Wien's Displacement Law",
    'lambda_max * T = 2.898e-3  m*K',
    [('lambda_max',   'Wavelength of peak radiation intensity', 'm'),
     ('T',            'Absolute temperature of blackbody', 'K'),
     ('2.898e-3 m*K', "Wien's displacement constant", 'm*K')],
    example={
        'problem': 'The Sun has lambda_max = 500 nm. Find its surface temperature.',
        'steps': ['T = 2.898e-3 / lambda_max = 2.898e-3 / 500e-9 = 5796 K'],
        'answer': 'T = 5796 K (~5800 K, consistent with known solar surface temperature)'})

eq_block(108, 'View Factor Reciprocity Relation',
    'A1*F12 = A2*F21',
    [('A1, A2', 'Areas of surfaces 1 and 2', 'm^2'),
     ('F12',    'View factor from surface 1 to 2 (fraction of radiation from 1 reaching 2)', 'dimensionless'),
     ('F21',    'View factor from surface 2 to 1', 'dimensionless')],
    note='Summation rule: sum(Fij) = 1 for all j (all radiation from surface i must reach some surface).',
    example={
        'problem': 'Two parallel plates: A1=2 m^2, A2=3 m^2. F12=0.6. Find F21.',
        'steps': [
            'A1*F12 = A2*F21',
            'F21 = A1*F12/A2 = 2*0.6/3 = 0.40',
        ],
        'answer': 'F21 = 0.40'})


# ── footer ────────────────────────────────────────────────────────────────────
ap('-' * 105, mono=True, sz=7, color=C['note'], before=10, after=10)
ap('Total: 108 equations  |  Topics: ME 101-106 / ME 201-206  |  All equations web-verified  |  SI units',
   sz=8, color=C['note'], align='center', before=0, after=6)
ap('-- HyESys Agent', bold=True, sz=9, color=C['title'], align='center', before=0, after=10)

doc.save(DST)
print(f'Saved: {DST}')
print(f'Total equations: 108')
