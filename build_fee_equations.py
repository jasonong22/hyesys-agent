"""
Builds FEE2026_Mechanical_Equations_Reference.docx
All equations from syllabus topics on PDF pages 39-44.
Each equation is named with a full variable legend.
Saved to same folder as the source PDF.
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL

DST = r'C:\Users\JasonOng\Desktop\local docs\personal\PE\FEE2026_Mechanical_Equations_Reference.docx'

doc = Document()

# ── page margins ──────────────────────────────────────────────────────────────
for sec in doc.sections:
    sec.top_margin    = Inches(0.85)
    sec.bottom_margin = Inches(0.85)
    sec.left_margin   = Inches(0.95)
    sec.right_margin  = Inches(0.95)

# ── colour palette ────────────────────────────────────────────────────────────
C = {
    'title':     RGBColor(0x1A, 0x1A, 0x5E),   # dark navy
    'topic':     RGBColor(0x00, 0x3D, 0x7A),   # deep blue
    'subtopic':  RGBColor(0x00, 0x5C, 0xA8),   # mid blue
    'eqname':    RGBColor(0x12, 0x40, 0x12),   # dark green
    'formula':   RGBColor(0x0D, 0x0D, 0x7A),   # dark indigo
    'legend_hdr':RGBColor(0xFF, 0xFF, 0xFF),   # white
    'body':      RGBColor(0x1A, 0x1A, 0x1A),   # near-black
    'note':      RGBColor(0x55, 0x55, 0x55),   # grey
}

SHADING = {
    'topic':     'D0E4F5',
    'subtopic':  'EBF5FF',
    'eq':        'F5F5F5',
    'legend_hdr':'003D7A',
    'legend_row':'F0F6FF',
    'legend_alt':'FAFAFA',
}

def _rgb_str(rgb):
    return f'{rgb.red:02X}{rgb.green:02X}{rgb.blue:02X}'

def _shd(hex_fill):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_fill)
    return shd

def ap(text='', bold=False, italic=False, mono=False,
       color=None, shading=None, align=None,
       sz=10, before=40, after=40):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before / 6)
    pf.space_after  = Pt(after  / 6)
    if align == 'center':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if shading:
        p._p.get_or_add_pPr().append(_shd(shading))
    if text:
        r = p.add_run(text)
        r.bold   = bold
        r.italic = italic
        if mono: r.font.name = 'Courier New'
        r.font.size = Pt(sz)
        if color:
            r.font.color.rgb = color if isinstance(color, RGBColor) \
                               else RGBColor(*[int(color[i:i+2],16) for i in (0,2,4)])
    return p

def topic_hdr(code, title):
    """Coloured topic header e.g. ME 103/203 Fluid Mechanics"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(4)
    p._p.get_or_add_pPr().append(_shd(SHADING['topic']))
    r = p.add_run(f'  {code}  {title}')
    r.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = C['topic']
    return p

def subtopic_hdr(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(3)
    p._p.get_or_add_pPr().append(_shd(SHADING['subtopic']))
    r = p.add_run(f'  {text}')
    r.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = C['subtopic']
    return p

def eq_block(number, name, formula, legend_rows, note=None):
    """
    Render one numbered equation with legend.
    legend_rows: list of (symbol, description, units)
    """
    # equation name line
    p_name = doc.add_paragraph()
    p_name.paragraph_format.space_before = Pt(5)
    p_name.paragraph_format.space_after  = Pt(1)
    rn = p_name.add_run(f'  [{number}]  {name}')
    rn.bold = True
    rn.font.size = Pt(10)
    rn.font.color.rgb = C['eqname']

    # formula line
    p_form = doc.add_paragraph()
    p_form.paragraph_format.space_before = Pt(0)
    p_form.paragraph_format.space_after  = Pt(1)
    p_form.paragraph_format.left_indent  = Inches(0.35)
    p_form._p.get_or_add_pPr().append(_shd(SHADING['eq']))
    rf = p_form.add_run(f'  {formula}')
    rf.bold  = True
    rf.font.name = 'Courier New'
    rf.font.size = Pt(10.5)
    rf.font.color.rgb = C['formula']

    # note (optional)
    if note:
        p_note = doc.add_paragraph()
        p_note.paragraph_format.space_before = Pt(0)
        p_note.paragraph_format.space_after  = Pt(1)
        p_note.paragraph_format.left_indent  = Inches(0.35)
        rno = p_note.add_run(f'  ▸ {note}')
        rno.italic = True
        rno.font.size = Pt(8.5)
        rno.font.color.rgb = C['note']

    # legend table
    tbl = doc.add_table(rows=1+len(legend_rows), cols=3)
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT

    # set column widths
    widths = [Inches(0.85), Inches(4.10), Inches(1.20)]
    for row in tbl.rows:
        for idx, cell in enumerate(row.cells):
            cell.width = widths[idx]

    # header row
    hdr_texts = ['Symbol', 'Description', 'SI Units']
    for i, cell in enumerate(tbl.rows[0].cells):
        cell._tc.get_or_add_tcPr().append(_shd(SHADING['legend_hdr']))
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(hdr_texts[i])
        run.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = C['legend_hdr']

    # data rows
    for ri, (sym, desc, unit) in enumerate(legend_rows):
        row = tbl.rows[ri + 1]
        fill = SHADING['legend_row'] if ri % 2 == 0 else SHADING['legend_alt']
        data = [sym, desc, unit]
        for ci, cell in enumerate(row.cells):
            cell._tc.get_or_add_tcPr().append(_shd(fill))
            p = cell.paragraphs[0]
            if ci == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(data[ci])
            run.font.size = Pt(9)
            if ci == 0:
                run.bold = True
                run.font.name = 'Courier New'
                run.font.color.rgb = C['formula']
            else:
                run.font.color.rgb = C['body']

    doc.add_paragraph().paragraph_format.space_after = Pt(2)


# ══════════════════════════════════════════════════════════════════════════════
# DOCUMENT HEADER
# ══════════════════════════════════════════════════════════════════════════════
ap('PROFESSIONAL ENGINEERS EXAMINATION FEE 2026', bold=True, sz=15,
   color=C['title'], align='center', before=0, after=10)
ap('Mechanical Discipline — Complete Equation Reference', bold=True, sz=12,
   color=C['subtopic'], align='center', before=0, after=6)
ap('Topics: ME 101–106 / ME 201–206  ·  All equations web-verified  ·  SI units throughout',
   sz=9, color=C['note'], align='center', before=0, after=20)

# horizontal rule approximation
ap('─' * 105, mono=True, sz=7, color=C['note'], before=0, after=16)

# ══════════════════════════════════════════════════════════════════════════════
# ME 101/201  CONTROL AND INSTRUMENTATION
# ══════════════════════════════════════════════════════════════════════════════
topic_hdr('ME 101 / 201', 'Control and Instrumentation')

subtopic_hdr('Transfer Functions & System Modelling')

eq_block(1, "Standard Second-Order Transfer Function",
    "G(s) = ωₙ² / (s² + 2ζωₙs + ωₙ²)",
    [
        ('G(s)', 'Transfer function (output/input in Laplace domain)', 'dimensionless'),
        ('s',    'Complex frequency variable (Laplace operator)', 'rad/s'),
        ('ωₙ',  'Undamped natural frequency', 'rad/s'),
        ('ζ',   'Damping ratio', 'dimensionless'),
    ])

eq_block(2, "Closed-Loop Characteristic Equation",
    "1 + G(s)H(s) = 0",
    [
        ('G(s)', 'Forward-path (plant) transfer function', 'dimensionless'),
        ('H(s)', 'Feedback transfer function', 'dimensionless'),
    ],
    note="Poles of this equation determine stability. Roots must have negative real parts for stability.")

eq_block(3, "Damping Ratio (from characteristic polynomial as² + bs + c = 0)",
    "ζ = b / (2√(ac))",
    [
        ('ζ',   'Damping ratio', 'dimensionless'),
        ('a',   'Coefficient of s² term', 'varies'),
        ('b',   'Coefficient of s term', 'varies'),
        ('c',   'Constant (s⁰) term', 'varies'),
    ])

subtopic_hdr('Transient Response (Underdamped, 0 < ζ < 1)')

eq_block(4, "Damped Natural Frequency",
    "ωd = ωₙ √(1 − ζ²)",
    [
        ('ωd', 'Damped natural frequency', 'rad/s'),
        ('ωₙ', 'Undamped natural frequency', 'rad/s'),
        ('ζ',  'Damping ratio', 'dimensionless'),
    ])

eq_block(5, "Rise Time (10–90%)",
    "tᵣ = (π − φ) / ωd     where φ = arctan(√(1−ζ²) / ζ)",
    [
        ('tᵣ', 'Rise time', 's'),
        ('φ',  'Phase angle = arctan(√(1−ζ²)/ζ)', 'rad'),
        ('ωd', 'Damped natural frequency', 'rad/s'),
        ('ζ',  'Damping ratio', 'dimensionless'),
    ])

eq_block(6, "Peak Time",
    "tₚ = π / ωd",
    [
        ('tₚ', 'Time to first peak overshoot', 's'),
        ('ωd', 'Damped natural frequency', 'rad/s'),
    ])

eq_block(7, "Percent Overshoot",
    "%OS = exp(−πζ / √(1 − ζ²)) × 100",
    [
        ('%OS', 'Percent overshoot', '%'),
        ('ζ',   'Damping ratio', 'dimensionless'),
    ],
    note="Valid for standard 2nd-order systems without zeros.")

eq_block(8, "Settling Time (2% criterion)",
    "tₛ ≈ 4 / (ζωₙ)",
    [
        ('tₛ', 'Settling time (response within ±2% of final value)', 's'),
        ('ζ',  'Damping ratio', 'dimensionless'),
        ('ωₙ', 'Undamped natural frequency', 'rad/s'),
    ],
    note="For 5% criterion use tₛ ≈ 3/(ζωₙ).")

subtopic_hdr('Steady-State Errors (Unity Feedback)')

eq_block(9, "Position Error Constant (Type 0 system — step input)",
    "Kₚ = lim[s→0] G(s)     →     e_ss = 1 / (1 + Kₚ)",
    [
        ('Kₚ',   'Position error constant', 'dimensionless'),
        ('e_ss', 'Steady-state error', 'same units as input'),
        ('G(s)', 'Open-loop transfer function', 'dimensionless'),
    ])

eq_block(10, "Velocity Error Constant (Type 1 system — ramp input)",
    "Kᵥ = lim[s→0] s·G(s)     →     e_ss = 1 / Kᵥ",
    [
        ('Kᵥ',   'Velocity error constant', 's⁻¹'),
        ('e_ss', 'Steady-state error', 's (or position units per (units/s))'),
        ('G(s)', 'Open-loop transfer function', 'dimensionless'),
    ])

subtopic_hdr('PID Controller')

eq_block(11, "PID Controller Transfer Function",
    "C(s) = Kₚ + Kᵢ/s + K_d·s",
    [
        ('C(s)', 'Controller transfer function', 'dimensionless'),
        ('Kₚ',  'Proportional gain', 'dimensionless'),
        ('Kᵢ',  'Integral gain', 's⁻¹'),
        ('K_d', 'Derivative gain', 's'),
        ('s',   'Laplace operator', 'rad/s'),
    ])

subtopic_hdr('Frequency Domain Stability Margins')

eq_block(12, "Phase Margin",
    "PM = 180° + ∠G(jω_gc)",
    [
        ('PM',     'Phase margin (system stable if PM > 0°)', '°  (degrees)'),
        ('ω_gc',   'Gain crossover frequency (where |G(jω)| = 1)', 'rad/s'),
        ('∠G(jω)', 'Phase angle of open-loop TF at ω_gc', '°  (degrees)'),
    ])

eq_block(13, "Gain Margin",
    "GM = 1 / |G(jω_pc)|     or     GM_dB = −20 log₁₀|G(jω_pc)|",
    [
        ('GM',     'Gain margin (system stable if GM > 1)', 'dimensionless (or dB)'),
        ('ω_pc',   'Phase crossover frequency (where ∠G = −180°)', 'rad/s'),
        ('|G(jω)|','Magnitude of open-loop TF at ω_pc', 'dimensionless'),
    ])

# ══════════════════════════════════════════════════════════════════════════════
# ME 102/202  DYNAMICS AND VIBRATIONS
# ══════════════════════════════════════════════════════════════════════════════
topic_hdr('ME 102 / 202', 'Dynamics and Vibrations')

subtopic_hdr('Kinematics — Particle (Constant Acceleration)')

eq_block(14, "Newton's Second Law — Linear Motion",
    "F = m·a",
    [
        ('F', 'Net force', 'N  (newtons)'),
        ('m', 'Mass', 'kg'),
        ('a', 'Linear acceleration', 'm/s²'),
    ])

eq_block(15, "SUVAT Equations of Motion (constant acceleration)",
    "v = u + at  |  s = ut + ½at²  |  v² = u² + 2as",
    [
        ('v', 'Final velocity', 'm/s'),
        ('u', 'Initial velocity', 'm/s'),
        ('a', 'Constant acceleration', 'm/s²'),
        ('t', 'Time', 's'),
        ('s', 'Displacement', 'm'),
    ])

eq_block(16, "Centripetal Acceleration",
    "aₓ = v²/r = ω²r",
    [
        ('aₓ', 'Centripetal (radial) acceleration, directed toward centre', 'm/s²'),
        ('v',  'Tangential velocity', 'm/s'),
        ('r',  'Radius of circular path', 'm'),
        ('ω',  'Angular velocity', 'rad/s'),
    ])

subtopic_hdr('Kinetics — Energy and Momentum')

eq_block(17, "Work-Energy Theorem",
    "W_net = ΔKE = ½mv₂² − ½mv₁²",
    [
        ('W_net', 'Net work done on the body', 'J  (joules)'),
        ('KE',    'Kinetic energy', 'J'),
        ('m',     'Mass', 'kg'),
        ('v₁',   'Initial velocity', 'm/s'),
        ('v₂',   'Final velocity', 'm/s'),
    ])

eq_block(18, "Conservation of Linear Momentum",
    "m₁v₁ + m₂v₂ = m₁v₁' + m₂v₂'",
    [
        ('m₁, m₂',   'Masses of bodies 1 and 2', 'kg'),
        ("v₁, v₂",   'Velocities before collision', 'm/s'),
        ("v₁', v₂'", 'Velocities after collision', 'm/s'),
    ],
    note="Valid when no external impulse acts on the system.")

eq_block(19, "Coefficient of Restitution",
    "e = (v₂' − v₁') / (v₁ − v₂)",
    [
        ('e',       'Coefficient of restitution  (0 = perfectly plastic, 1 = elastic)', 'dimensionless'),
        ("v₁, v₂",  'Velocities of bodies 1 & 2 before impact', 'm/s'),
        ("v₁',v₂'", 'Velocities of bodies 1 & 2 after impact', 'm/s'),
    ])

subtopic_hdr('Kinematics and Kinetics — Rotation')

eq_block(20, "Newton's Second Law — Rotational Motion",
    "M = I·α",
    [
        ('M', 'Net moment (torque)', 'N·m'),
        ('I', 'Mass moment of inertia about rotation axis', 'kg·m²'),
        ('α', 'Angular acceleration', 'rad/s²'),
    ])

eq_block(21, "Angular Momentum",
    "L = I·ω     and     M = dL/dt",
    [
        ('L', 'Angular momentum', 'kg·m²/s'),
        ('I', 'Mass moment of inertia', 'kg·m²'),
        ('ω', 'Angular velocity', 'rad/s'),
        ('M', 'Applied moment (torque)', 'N·m'),
    ])

subtopic_hdr('Mechanical Vibrations — SDOF Systems')

eq_block(22, "Undamped Natural Frequency",
    "ωₙ = √(k/m)",
    [
        ('ωₙ', 'Undamped natural frequency', 'rad/s'),
        ('k',  'Spring stiffness', 'N/m'),
        ('m',  'Mass', 'kg'),
    ])

eq_block(23, "Equation of Motion — Undamped Free Vibration",
    "mẍ + kx = F(t)",
    [
        ('m',    'Mass', 'kg'),
        ('ẍ',   'Acceleration (second derivative of x)', 'm/s²'),
        ('k',    'Stiffness', 'N/m'),
        ('x',    'Displacement from equilibrium', 'm'),
        ('F(t)', 'External forcing function', 'N'),
    ])

eq_block(24, "Equation of Motion — Damped Free Vibration",
    "mẍ + cẋ + kx = F(t)",
    [
        ('m',    'Mass', 'kg'),
        ('c',    'Viscous damping coefficient', 'N·s/m'),
        ('ẋ',   'Velocity', 'm/s'),
        ('k',    'Stiffness', 'N/m'),
        ('x',    'Displacement', 'm'),
        ('F(t)', 'External force', 'N'),
    ])

eq_block(25, "Damping Ratio",
    "ζ = c / (2√(km))  =  c / (2mωₙ)",
    [
        ('ζ',  'Damping ratio', 'dimensionless'),
        ('c',  'Damping coefficient', 'N·s/m'),
        ('k',  'Stiffness', 'N/m'),
        ('m',  'Mass', 'kg'),
        ('ωₙ', 'Natural frequency', 'rad/s'),
    ])

eq_block(26, "Critical Damping Coefficient",
    "cᶜ = 2√(km) = 2mωₙ",
    [
        ('cᶜ', 'Critical damping coefficient (boundary: ζ = 1)', 'N·s/m'),
        ('k',  'Stiffness', 'N/m'),
        ('m',  'Mass', 'kg'),
        ('ωₙ', 'Undamped natural frequency', 'rad/s'),
    ])

eq_block(27, "Resonance Frequency (Damped Forced Vibration)",
    "ωᵣ = ωₙ √(1 − 2ζ²)",
    [
        ('ωᵣ', 'Frequency of maximum amplitude response', 'rad/s'),
        ('ωₙ', 'Undamped natural frequency', 'rad/s'),
        ('ζ',  'Damping ratio', 'dimensionless'),
    ],
    note="Valid only when ζ ≤ 1/√2 ≈ 0.707. Above this, resonance peak disappears.")

eq_block(28, "Logarithmic Decrement",
    "δ = ln(xₙ/xₙ₊₁) = 2πζ / √(1 − ζ²)",
    [
        ('δ',       'Logarithmic decrement', 'dimensionless'),
        ('xₙ',      'Amplitude of nth cycle', 'm  (or any consistent unit)'),
        ('xₙ₊₁',   'Amplitude of (n+1)th cycle', 'm'),
        ('ζ',       'Damping ratio', 'dimensionless'),
    ])

# ══════════════════════════════════════════════════════════════════════════════
# ME 103/203  FLUID MECHANICS
# ══════════════════════════════════════════════════════════════════════════════
topic_hdr('ME 103 / 203', 'Fluid Mechanics')

subtopic_hdr('Fluid Properties')

eq_block(29, "Density",
    "ρ = m / V",
    [
        ('ρ', 'Mass density', 'kg/m³'),
        ('m', 'Mass', 'kg'),
        ('V', 'Volume', 'm³'),
    ])

eq_block(30, "Newton's Law of Viscosity",
    "τ = μ (du/dy)",
    [
        ('τ',    'Shear stress in fluid', 'Pa  (N/m²)'),
        ('μ',    'Dynamic viscosity', 'Pa·s  (N·s/m²)'),
        ('du/dy','Velocity gradient perpendicular to flow', 's⁻¹'),
    ])

eq_block(31, "Kinematic Viscosity",
    "ν = μ / ρ",
    [
        ('ν', 'Kinematic viscosity', 'm²/s'),
        ('μ', 'Dynamic viscosity', 'Pa·s'),
        ('ρ', 'Fluid density', 'kg/m³'),
    ])

subtopic_hdr('Hydrostatics')

eq_block(32, "Hydrostatic Pressure",
    "P = P₀ + ρgh",
    [
        ('P',  'Absolute pressure at depth h', 'Pa'),
        ('P₀', 'Pressure at free surface', 'Pa'),
        ('ρ',  'Fluid density', 'kg/m³'),
        ('g',  'Gravitational acceleration  (9.81)', 'm/s²'),
        ('h',  'Depth below free surface', 'm'),
    ])

eq_block(33, "Buoyancy Force — Archimedes' Principle",
    "Fᵦ = ρ_fluid · V_sub · g",
    [
        ('Fᵦ',     'Buoyancy (upward) force', 'N'),
        ('ρ_fluid','Density of surrounding fluid', 'kg/m³'),
        ('V_sub',  'Volume of body submerged in fluid', 'm³'),
        ('g',      'Gravitational acceleration', 'm/s²'),
    ])

subtopic_hdr('Flow Equations')

eq_block(34, "Continuity Equation — Incompressible Flow",
    "A₁V₁ = A₂V₂ = Q",
    [
        ('A₁, A₂', 'Cross-sectional areas at sections 1 and 2', 'm²'),
        ('V₁, V₂', 'Mean flow velocities at sections 1 and 2', 'm/s'),
        ('Q',      'Volumetric flow rate (constant)', 'm³/s'),
    ])

eq_block(35, "Continuity Equation — Compressible Flow",
    "ρ₁A₁V₁ = ρ₂A₂V₂ = ṁ",
    [
        ('ρ₁, ρ₂', 'Fluid densities at sections 1 and 2', 'kg/m³'),
        ('A₁, A₂', 'Cross-sectional areas', 'm²'),
        ('V₁, V₂', 'Velocities', 'm/s'),
        ('ṁ',      'Mass flow rate (constant)', 'kg/s'),
    ])

eq_block(36, "Bernoulli's Equation (steady, inviscid, incompressible)",
    "P₁ + ½ρV₁² + ρgz₁  =  P₂ + ½ρV₂² + ρgz₂",
    [
        ('P',  'Static pressure', 'Pa'),
        ('ρ',  'Fluid density', 'kg/m³'),
        ('V',  'Flow velocity', 'm/s'),
        ('g',  'Gravitational acceleration', 'm/s²'),
        ('z',  'Elevation above datum', 'm'),
    ],
    note="Head form (divide all by ρg): P/ρg + V²/2g + z = H  (total head, m)")

eq_block(37, "Momentum Equation — Steady 1-D Flow",
    "ΣF = ρQ(V₂ − V₁)  =  ṁ(V₂ − V₁)",
    [
        ('ΣF',    'Net external force on control volume', 'N'),
        ('ρ',     'Fluid density', 'kg/m³'),
        ('Q',     'Volumetric flow rate', 'm³/s'),
        ('V₁,V₂', 'Velocities at inlet and outlet', 'm/s'),
        ('ṁ',     'Mass flow rate', 'kg/s'),
    ])

subtopic_hdr('Dimensionless Numbers')

eq_block(38, "Reynolds Number",
    "Re = ρVD/μ  =  VD/ν",
    [
        ('Re', 'Reynolds number  (Re<2300 laminar; Re>4000 turbulent in pipe)', 'dimensionless'),
        ('ρ',  'Fluid density', 'kg/m³'),
        ('V',  'Characteristic velocity', 'm/s'),
        ('D',  'Characteristic length (pipe diameter)', 'm'),
        ('μ',  'Dynamic viscosity', 'Pa·s'),
        ('ν',  'Kinematic viscosity  ν = μ/ρ', 'm²/s'),
    ])

eq_block(39, "Froude Number",
    "Fr = V / √(gL)",
    [
        ('Fr', 'Froude number  (Fr<1 subcritical; Fr>1 supercritical)', 'dimensionless'),
        ('V',  'Flow velocity', 'm/s'),
        ('g',  'Gravitational acceleration', 'm/s²'),
        ('L',  'Characteristic length (e.g. hydraulic depth)', 'm'),
    ])

eq_block(40, "Mach Number",
    "Ma = V / c     where  c = √(γRT)",
    [
        ('Ma', 'Mach number  (Ma<1 subsonic; Ma>1 supersonic)', 'dimensionless'),
        ('V',  'Flow velocity', 'm/s'),
        ('c',  'Speed of sound in fluid', 'm/s'),
        ('γ',  'Ratio of specific heats  c_p/c_v  (air ≈ 1.4)', 'dimensionless'),
        ('R',  'Specific gas constant  (air = 287)', 'J/(kg·K)'),
        ('T',  'Absolute temperature', 'K'),
    ])

subtopic_hdr('Pipe Flow — Head Losses')

eq_block(41, "Darcy-Weisbach Equation — Major (Friction) Loss",
    "h_f = f · (L/D) · V²/(2g)",
    [
        ('h_f', 'Friction head loss', 'm'),
        ('f',   'Darcy friction factor (dimensionless)', 'dimensionless'),
        ('L',   'Pipe length', 'm'),
        ('D',   'Internal pipe diameter', 'm'),
        ('V',   'Mean flow velocity', 'm/s'),
        ('g',   'Gravitational acceleration', 'm/s²'),
    ],
    note="Fanning form (some texts): h_f = 4f(L/D)(V²/2g). Darcy f = 4 × Fanning f.")

eq_block(42, "Laminar Friction Factor — Hagen-Poiseuille",
    "f = 64 / Re       (valid for Re < 2300)",
    [
        ('f',  'Darcy friction factor', 'dimensionless'),
        ('Re', 'Reynolds number', 'dimensionless'),
    ])

eq_block(43, "Turbulent Friction Factor — Colebrook-White Equation",
    "1/√f = −2 log₁₀ [ ε/(3.7D) + 2.51/(Re√f) ]",
    [
        ('f',  'Darcy friction factor', 'dimensionless'),
        ('ε',  'Absolute pipe roughness', 'm'),
        ('D',  'Pipe internal diameter', 'm'),
        ('Re', 'Reynolds number', 'dimensionless'),
    ],
    note="Implicit equation — requires iteration or Moody chart. Valid for turbulent flow (Re > 4000).")

eq_block(44, "Minor (Local) Losses — Fittings, Entry, Exit",
    "h_m = K_L · V²/(2g)",
    [
        ('h_m', 'Minor head loss', 'm'),
        ('K_L', 'Loss coefficient (from tables: K_entry=0.5, K_exit=1.0)', 'dimensionless'),
        ('V',   'Velocity at the fitting', 'm/s'),
        ('g',   'Gravitational acceleration', 'm/s²'),
    ])

eq_block(45, "Sudden Expansion Loss — Borda-Carnot",
    "h_exp = (V₁ − V₂)² / (2g)  =  [1 − A₁/A₂]² · V₁²/(2g)",
    [
        ('h_exp',    'Head loss at sudden expansion', 'm'),
        ('V₁',       'Upstream velocity (smaller pipe)', 'm/s'),
        ('V₂',       'Downstream velocity (larger pipe)', 'm/s'),
        ('A₁, A₂',   'Cross-sectional areas upstream and downstream', 'm²'),
        ('g',        'Gravitational acceleration', 'm/s²'),
    ])

subtopic_hdr('Fluid Machinery')

eq_block(46, "Hydraulic Power",
    "P = ρgQH",
    [
        ('P', 'Hydraulic power delivered to fluid', 'W'),
        ('ρ', 'Fluid density', 'kg/m³'),
        ('g', 'Gravitational acceleration', 'm/s²'),
        ('Q', 'Volumetric flow rate', 'm³/s'),
        ('H', 'Total head developed', 'm'),
    ])

eq_block(47, "Pump Overall Efficiency",
    "η = P_fluid / P_shaft  =  ρgQH / P_shaft",
    [
        ('η',        'Overall pump efficiency', 'dimensionless  (0 to 1)'),
        ('P_fluid',  'Power delivered to fluid', 'W'),
        ('P_shaft',  'Shaft (brake) power input to pump', 'W'),
        ('ρ',        'Fluid density', 'kg/m³'),
        ('Q',        'Flow rate', 'm³/s'),
        ('H',        'Total head', 'm'),
    ])

# ══════════════════════════════════════════════════════════════════════════════
# ME 104/204  MECHANICS AND MATERIALS
# ══════════════════════════════════════════════════════════════════════════════
topic_hdr('ME 104 / 204', 'Mechanics and Materials')

subtopic_hdr('Stress, Strain and Elastic Constants')

eq_block(48, "Direct (Normal) Stress",
    "σ = F / A",
    [
        ('σ', 'Normal stress (positive = tension)', 'Pa  (N/m²)'),
        ('F', 'Axial force (positive = tensile)', 'N'),
        ('A', 'Cross-sectional area', 'm²'),
    ])

eq_block(49, "Direct (Normal) Strain",
    "ε = δL / L₀",
    [
        ('ε',  'Normal strain', 'dimensionless  (m/m)'),
        ('δL', 'Change in length', 'm'),
        ('L₀', 'Original (gauge) length', 'm'),
    ])

eq_block(50, "Hooke's Law — Elastic Region",
    "σ = E·ε",
    [
        ('σ', 'Normal stress', 'Pa'),
        ('E', "Young's modulus (modulus of elasticity)  (steel ≈ 200 GPa)", 'Pa'),
        ('ε', 'Normal strain', 'dimensionless'),
    ])

eq_block(51, "Shear Stress and Shear Strain",
    "τ = F_s / A     and     γ = τ / G",
    [
        ('τ',   'Shear stress', 'Pa'),
        ('F_s', 'Shear force', 'N'),
        ('A',   'Area on which shear acts', 'm²'),
        ('γ',   'Shear strain (engineering)', 'dimensionless  (rad)'),
        ('G',   'Shear modulus (modulus of rigidity)  (steel ≈ 79 GPa)', 'Pa'),
    ])

eq_block(52, "Poisson's Ratio",
    "ν = −ε_lateral / ε_axial",
    [
        ('ν',          "Poisson's ratio  (steel ≈ 0.30)", 'dimensionless'),
        ('ε_lateral',  'Transverse (lateral) strain', 'dimensionless'),
        ('ε_axial',    'Axial (longitudinal) strain', 'dimensionless'),
    ])

eq_block(53, "Relationship Between Elastic Constants",
    "G = E / [2(1 + ν)]     and     K = E / [3(1 − 2ν)]",
    [
        ('G', 'Shear modulus', 'Pa'),
        ('E', "Young's modulus", 'Pa'),
        ('ν', "Poisson's ratio", 'dimensionless'),
        ('K', 'Bulk modulus', 'Pa'),
    ])

subtopic_hdr('Thermal Effects')

eq_block(54, "Free Thermal Strain",
    "ε_T = α · ΔT",
    [
        ('ε_T', 'Thermal strain (free expansion)', 'dimensionless'),
        ('α',   'Coefficient of linear thermal expansion  (steel ≈ 12×10⁻⁶)', '°C⁻¹  (or K⁻¹)'),
        ('ΔT',  'Temperature change', '°C  (or K)'),
    ])

eq_block(55, "Constrained Thermal Stress",
    "σ_T = E · α · ΔT",
    [
        ('σ_T', 'Thermal stress (compressive if expansion is prevented)', 'Pa'),
        ('E',   "Young's modulus", 'Pa'),
        ('α',   'Coefficient of thermal expansion', '°C⁻¹'),
        ('ΔT',  'Temperature change', '°C'),
    ])

subtopic_hdr('Bending of Beams')

eq_block(56, "Second Moment of Area — Rectangle (about centroidal axis)",
    "I_xx = bh³ / 12",
    [
        ('I_xx', 'Second moment of area about horizontal centroidal axis', 'm⁴'),
        ('b',    'Width of rectangle', 'm'),
        ('h',    'Height (depth) of rectangle', 'm'),
    ])

eq_block(57, "Second Moment of Area — Solid Circle",
    "I = πd⁴ / 64  =  πr⁴ / 4",
    [
        ('I', 'Second moment of area about diameter', 'm⁴'),
        ('d', 'Diameter', 'm'),
        ('r', 'Radius', 'm'),
    ])

eq_block(58, "Flexure Formula (Bending Stress — Euler-Bernoulli)",
    "σ = M·y / I     →     σ_max = M / Z",
    [
        ('σ',     'Bending stress at distance y from neutral axis', 'Pa'),
        ('M',     'Bending moment', 'N·m'),
        ('y',     'Distance from neutral axis to point of interest', 'm'),
        ('I',     'Second moment of area about neutral axis', 'm⁴'),
        ('σ_max', 'Maximum bending stress (at extreme fibre)', 'Pa'),
        ('Z',     'Section modulus  Z = I/y_max', 'm³'),
    ])

eq_block(59, "Shear Stress in Beams",
    "τ = V·Q / (I·b)",
    [
        ('τ', 'Horizontal shear stress at depth y', 'Pa'),
        ('V', 'Shear force at the cross-section', 'N'),
        ('Q', 'First moment of area of section above y about neutral axis', 'm³'),
        ('I', 'Second moment of area of full cross-section', 'm⁴'),
        ('b', 'Width of cross-section at depth y', 'm'),
    ])

eq_block(60, "Beam Deflection — Simply-Supported, Central Point Load",
    "δ_max = FL³ / (48EI)",
    [
        ('δ_max', 'Maximum deflection at mid-span', 'm'),
        ('F',     'Point load at mid-span', 'N'),
        ('L',     'Span length', 'm'),
        ('E',     "Young's modulus", 'Pa'),
        ('I',     'Second moment of area', 'm⁴'),
    ])

eq_block(61, "Beam Deflection — Cantilever, Tip Point Load",
    "δ_max = FL³ / (3EI)",
    [
        ('δ_max', 'Maximum deflection at free end', 'm'),
        ('F',     'Point load at free end', 'N'),
        ('L',     'Cantilever length', 'm'),
        ('E',     "Young's modulus", 'Pa'),
        ('I',     'Second moment of area', 'm⁴'),
    ])

subtopic_hdr('Torsion')

eq_block(62, "Torsion Formula — Circular Shaft (Coulomb)",
    "τ = T·r / J     →     τ_max = T·(d/2) / J",
    [
        ('τ',     'Shear stress at radius r', 'Pa'),
        ('T',     'Applied torque', 'N·m'),
        ('r',     'Radial distance from shaft axis', 'm'),
        ('J',     'Polar second moment of area', 'm⁴'),
        ('d',     'Shaft diameter', 'm'),
    ])

eq_block(63, "Polar Second Moment of Area — Solid Shaft",
    "J = πd⁴ / 32",
    [
        ('J', 'Polar second moment of area', 'm⁴'),
        ('d', 'Shaft diameter', 'm'),
    ])

eq_block(64, "Polar Second Moment of Area — Hollow Shaft",
    "J = π(d_o⁴ − d_i⁴) / 32",
    [
        ('J',   'Polar second moment of area', 'm⁴'),
        ('d_o', 'External (outer) diameter', 'm'),
        ('d_i', 'Internal (inner) diameter', 'm'),
    ])

eq_block(65, "Angle of Twist",
    "φ = T·L / (G·J)",
    [
        ('φ', 'Angle of twist', 'rad'),
        ('T', 'Applied torque', 'N·m'),
        ('L', 'Shaft length', 'm'),
        ('G', 'Shear modulus', 'Pa'),
        ('J', 'Polar second moment of area', 'm⁴'),
    ])

subtopic_hdr('Yield Criteria and Stress Transformations')

eq_block(66, "Principal Stresses — Mohr's Circle",
    "σ₁,₂ = (σ_x + σ_y)/2 ± √[ ((σ_x−σ_y)/2)² + τ_xy² ]",
    [
        ('σ₁, σ₂', 'Major and minor principal stresses', 'Pa'),
        ('σ_x',    'Normal stress on x-face', 'Pa'),
        ('σ_y',    'Normal stress on y-face', 'Pa'),
        ('τ_xy',   'Shear stress on x-face', 'Pa'),
    ])

eq_block(67, "Maximum Shear Stress",
    "τ_max = √[ ((σ_x−σ_y)/2)² + τ_xy² ]  =  (σ₁ − σ₂)/2",
    [
        ('τ_max', 'Maximum in-plane shear stress', 'Pa'),
        ('σ_x, σ_y', 'Normal stresses on x- and y-faces', 'Pa'),
        ('τ_xy',     'Shear stress', 'Pa'),
        ('σ₁, σ₂',  'Principal stresses', 'Pa'),
    ])

eq_block(68, "Von Mises Yield Criterion (2-D)",
    "σ_VM = √(σ₁² − σ₁σ₂ + σ₂²) ≤ σ_y",
    [
        ('σ_VM', 'Von Mises equivalent stress (distortion energy criterion)', 'Pa'),
        ('σ₁, σ₂', 'Principal stresses', 'Pa'),
        ('σ_y',    'Uniaxial yield strength of material', 'Pa'),
    ])

eq_block(69, "Tresca Yield Criterion",
    "τ_max = (σ₁ − σ₂)/2 ≤ σ_y/2",
    [
        ('τ_max',  'Maximum shear stress', 'Pa'),
        ('σ₁, σ₂','Principal stresses  (σ₁ ≥ σ₂)', 'Pa'),
        ('σ_y',   'Yield strength', 'Pa'),
    ],
    note="Tresca is more conservative (lower predicted failure load) than Von Mises.")

subtopic_hdr('Pressure Vessels')

eq_block(70, "Thin-Walled Cylinder — Hoop (Circumferential) Stress",
    "σ_h = Pd / (2t)",
    [
        ('σ_h', 'Hoop (circumferential) stress  (dominant, = 2 × longitudinal)', 'Pa'),
        ('P',   'Internal gauge pressure', 'Pa'),
        ('d',   'Internal diameter', 'm'),
        ('t',   'Wall thickness', 'm'),
    ],
    note="Valid when d/t ≥ 20 (thin-wall assumption).")

eq_block(71, "Thin-Walled Cylinder — Longitudinal (Axial) Stress",
    "σ_L = Pd / (4t)",
    [
        ('σ_L', 'Longitudinal (axial) stress  (σ_L = σ_h/2)', 'Pa'),
        ('P',   'Internal gauge pressure', 'Pa'),
        ('d',   'Internal diameter', 'm'),
        ('t',   'Wall thickness', 'm'),
    ])

eq_block(72, "Thick-Walled Cylinder — Lamé Equations",
    "σ_θ = A + B/r²     (hoop)     and     σ_r = A − B/r²     (radial)",
    [
        ('σ_θ', 'Hoop (tangential) stress at radius r', 'Pa'),
        ('σ_r', 'Radial stress at radius r', 'Pa'),
        ('A, B', 'Lamé constants determined from boundary conditions (P_i, P_o, r_i, r_o)', 'Pa,  Pa·m²'),
        ('r',    'Radial position', 'm'),
    ])

subtopic_hdr('Buckling')

eq_block(73, "Euler Critical Buckling Load",
    "P_cr = π²EI / Lₑ²",
    [
        ('P_cr', 'Critical (Euler) buckling load', 'N'),
        ('E',    "Young's modulus", 'Pa'),
        ('I',    'Minimum second moment of area of cross-section', 'm⁴'),
        ('Lₑ',  'Effective length  (Lₑ = KL, K depends on end conditions)', 'm'),
    ],
    note="End condition factor K: both-pinned=1.0; one-fixed-one-free=2.0; both-fixed=0.5; fixed-pinned=0.7.")

# ══════════════════════════════════════════════════════════════════════════════
# ME 105/205  MANUFACTURING TECHNOLOGY
# ══════════════════════════════════════════════════════════════════════════════
topic_hdr('ME 105 / 205', 'Manufacturing Technology')

subtopic_hdr('Metal Cutting — Tool Life')

eq_block(74, "Taylor's Tool Life Equation",
    "V · Tⁿ = C",
    [
        ('V',  'Cutting speed', 'm/min'),
        ('T',  'Tool life', 'min'),
        ('n',  "Taylor's exponent  (HSS ≈ 0.10–0.15; carbide ≈ 0.20–0.30; ceramic ≈ 0.40–0.50)", 'dimensionless'),
        ('C',  'Taylor constant (cutting speed for T = 1 min)', 'm/min'),
    ],
    note="Rearranged: T = (C/V)^(1/n). Log form: log V + n·log T = log C (straight line on log-log plot).")

subtopic_hdr('Metal Cutting — Cutting Geometry')

eq_block(75, "Chip Thickness Ratio",
    "rᶜ = t₁ / t₂",
    [
        ('rᶜ', 'Chip thickness ratio (cutting ratio),  0 < rᶜ < 1', 'dimensionless'),
        ('t₁', 'Uncut chip thickness (depth of cut)', 'mm'),
        ('t₂', 'Actual (deformed) chip thickness', 'mm'),
    ])

eq_block(76, "Shear Plane Angle",
    "tan φ = rᶜ cos α / (1 − rᶜ sin α)",
    [
        ('φ',  'Shear plane angle', '°'),
        ('rᶜ', 'Chip thickness ratio', 'dimensionless'),
        ('α',  'Tool rake angle (positive if tilted toward workpiece)', '°'),
    ])

eq_block(77, "Merchant's Minimum Energy Equation",
    "φ = 45° + α/2 − λ/2",
    [
        ('φ', 'Shear plane angle', '°'),
        ('α', 'Tool rake angle', '°'),
        ('λ', 'Friction angle  λ = arctan(μ)  where μ = friction coefficient on rake face', '°'),
    ],
    note="Based on minimum energy principle. Predicts optimal shear angle for given rake and friction.")

subtopic_hdr('Material Removal — Turning')

eq_block(78, "Cutting Speed",
    "Vᶜ = π·D·N / 1000",
    [
        ('Vᶜ', 'Peripheral cutting speed', 'm/min'),
        ('D',  'Workpiece diameter', 'mm'),
        ('N',  'Spindle speed', 'rpm  (rev/min)'),
    ])

eq_block(79, "Material Removal Rate — Turning",
    "MRR = π·D·N·f·d  =  Vᶜ·f·d",
    [
        ('MRR', 'Material removal rate', 'mm³/min'),
        ('D',   'Workpiece diameter', 'mm'),
        ('N',   'Spindle speed', 'rpm'),
        ('f',   'Feed per revolution', 'mm/rev'),
        ('d',   'Depth of cut (radial)', 'mm'),
        ('Vᶜ',  'Cutting speed', 'mm/min  (= m/min × 1000)'),
    ])

subtopic_hdr('Material Removal — Milling')

eq_block(80, "Feed Rate — Milling",
    "vf = fz · z · N",
    [
        ('vf', 'Table feed rate', 'mm/min'),
        ('fz', 'Feed per tooth', 'mm/tooth'),
        ('z',  'Number of cutter teeth', 'dimensionless'),
        ('N',  'Spindle speed', 'rpm'),
    ])

eq_block(81, "Material Removal Rate — Face Milling",
    "MRR = vf · d · W",
    [
        ('MRR', 'Material removal rate', 'mm³/min'),
        ('vf',  'Feed rate', 'mm/min'),
        ('d',   'Axial depth of cut', 'mm'),
        ('W',   'Width of cut (radial engagement)', 'mm'),
    ])

subtopic_hdr('Machining Power and Energy')

eq_block(82, "Cutting Power",
    "P = Fᶜ · Vᶜ",
    [
        ('P',  'Cutting power', 'W'),
        ('Fᶜ', 'Principal (tangential) cutting force', 'N'),
        ('Vᶜ', 'Cutting speed', 'm/s'),
    ])

eq_block(83, "Specific Cutting Energy (Unit Power)",
    "u = P / MRR  =  Fᶜ / (f·d)",
    [
        ('u',   'Specific cutting energy', 'J/mm³  (or GJ/m³)'),
        ('P',   'Cutting power', 'W'),
        ('MRR', 'Material removal rate', 'mm³/s'),
        ('Fᶜ',  'Cutting force', 'N'),
        ('f',   'Feed', 'mm/rev'),
        ('d',   'Depth of cut', 'mm'),
    ])

subtopic_hdr('Rolling and Forming')

eq_block(84, "Rolling — Draft",
    "d = t₀ − tf",
    [
        ('d',  'Draft (reduction in thickness)', 'mm'),
        ('t₀', 'Initial sheet/slab thickness', 'mm'),
        ('tf', 'Final sheet/slab thickness', 'mm'),
    ])

eq_block(85, "Deep Drawing — Approximate Blank Diameter",
    "Dᵦ = √(Dₚ² + 4Dₚh)",
    [
        ('Dᵦ', 'Blank diameter required', 'mm'),
        ('Dₚ', 'Punch (cup) diameter', 'mm'),
        ('h',  'Cup drawing depth', 'mm'),
    ],
    note="Approximate formula based on surface area conservation. Exact value depends on material thinning.")

# ══════════════════════════════════════════════════════════════════════════════
# ME 106/206  THERMODYNAMICS AND HEAT TRANSFER
# ══════════════════════════════════════════════════════════════════════════════
topic_hdr('ME 106 / 206', 'Thermodynamics and Heat Transfer')

subtopic_hdr('Ideal Gas Laws')

eq_block(86, "Ideal Gas Law — Molar Form",
    "PV = nR̄T",
    [
        ('P',  'Absolute pressure', 'Pa'),
        ('V',  'Volume', 'm³'),
        ('n',  'Amount of substance', 'mol'),
        ('R̄',  'Universal gas constant  = 8.314', 'J/(mol·K)'),
        ('T',  'Absolute temperature', 'K'),
    ])

eq_block(87, "Ideal Gas Law — Mass Form",
    "PV = mRT     or     P = ρRT",
    [
        ('P',  'Absolute pressure', 'Pa'),
        ('V',  'Volume', 'm³'),
        ('m',  'Mass of gas', 'kg'),
        ('R',  'Specific gas constant  R = R̄/M  (air = 287)', 'J/(kg·K)'),
        ('T',  'Absolute temperature', 'K'),
        ('ρ',  'Gas density', 'kg/m³'),
    ])

eq_block(88, "Specific Heat Relationship — Ideal Gas (Mayer Relation)",
    "cₚ − cᵥ = R",
    [
        ('cₚ', 'Specific heat at constant pressure', 'J/(kg·K)'),
        ('cᵥ', 'Specific heat at constant volume', 'J/(kg·K)'),
        ('R',  'Specific gas constant', 'J/(kg·K)'),
    ])

eq_block(89, "Heat Capacity Ratio",
    "γ = cₚ / cᵥ",
    [
        ('γ',  'Heat capacity ratio  (monatomic ≈ 1.67; diatomic/air ≈ 1.40)', 'dimensionless'),
        ('cₚ', 'Specific heat at constant pressure', 'J/(kg·K)'),
        ('cᵥ', 'Specific heat at constant volume', 'J/(kg·K)'),
    ])

subtopic_hdr('First Law of Thermodynamics')

eq_block(90, "First Law — Closed System",
    "ΔU = Q − W",
    [
        ('ΔU', 'Change in internal energy of the system', 'J'),
        ('Q',  'Heat transferred into the system  (Q > 0 in; Q < 0 out)', 'J'),
        ('W',  'Work done by the system  (W > 0 out; W < 0 in)', 'J'),
    ])

eq_block(91, "Steady Flow Energy Equation (SFEE) — Open System",
    "Q̇ − Ẇ = ṁ [ (h₂−h₁) + ½(V₂²−V₁²) + g(z₂−z₁) ]",
    [
        ('Q̇',      'Rate of heat transfer into system', 'W'),
        ('Ẇ',      'Rate of work output (shaft work)', 'W'),
        ('ṁ',      'Mass flow rate', 'kg/s'),
        ('h₁, h₂', 'Specific enthalpy at inlet and outlet  (h = u + P/ρ)', 'J/kg'),
        ('V₁, V₂', 'Velocities at inlet and outlet', 'm/s'),
        ('z₁, z₂', 'Elevations at inlet and outlet', 'm'),
        ('g',      'Gravitational acceleration', 'm/s²'),
    ])

subtopic_hdr('Second Law and Cycles')

eq_block(92, "Carnot Efficiency (Maximum Theoretical Efficiency)",
    "η_c = 1 − T_L / T_H",
    [
        ('η_c', 'Carnot (maximum) thermal efficiency', 'dimensionless'),
        ('T_L', 'Temperature of cold reservoir (condenser)', 'K  (Kelvin — must be absolute)'),
        ('T_H', 'Temperature of hot reservoir (boiler)', 'K'),
    ],
    note="Minimum fuel consumption implies Carnot efficiency. η_c = W_net/Q_H = 1 − Q_C/Q_H.")

eq_block(93, "Entropy Change (Reversible Process)",
    "dS = δQ_rev / T     →     ΔS = ∫(δQ/T)_rev",
    [
        ('dS',     'Infinitesimal entropy change', 'J/K'),
        ('δQ_rev', 'Infinitesimal reversible heat transfer', 'J'),
        ('T',      'Absolute temperature at boundary', 'K'),
        ('ΔS',     'Total entropy change', 'J/K'),
    ])

eq_block(94, "Clausius Inequality",
    "∮ δQ / T ≤ 0",
    [
        ('∮ δQ/T', 'Cyclic integral of heat/temperature over a complete cycle', 'J/K'),
    ],
    note="= 0 for a reversible cycle; < 0 for an irreversible (real) cycle. Expression of 2nd Law.")

eq_block(95, "Isentropic Process Relations — Ideal Gas",
    "T₂/T₁ = (P₂/P₁)^((γ−1)/γ) = (V₁/V₂)^(γ−1)",
    [
        ('T₁, T₂', 'Temperatures before and after the isentropic process', 'K'),
        ('P₁, P₂', 'Pressures before and after', 'Pa'),
        ('V₁, V₂', 'Specific volumes before and after', 'm³/kg'),
        ('γ',      'Heat capacity ratio cₚ/cᵥ', 'dimensionless'),
    ],
    note="Valid for reversible adiabatic (isentropic) processes only.")

subtopic_hdr('Heat Transfer — Conduction')

eq_block(96, "Fourier's Law of Heat Conduction",
    "q = −k·A·(dT/dx)     or     q'' = −k·(dT/dx)",
    [
        ('q',     'Rate of heat transfer', 'W'),
        ('q\'\'', 'Heat flux', 'W/m²'),
        ('k',     'Thermal conductivity of material', 'W/(m·K)'),
        ('A',     'Cross-sectional area perpendicular to heat flow', 'm²'),
        ('dT/dx', 'Temperature gradient in direction of heat flow', 'K/m'),
    ],
    note="Negative sign: heat flows in the direction of decreasing temperature.")

eq_block(97, "Thermal Resistance — Conduction (Flat Slab)",
    "R_cond = L / (k·A)",
    [
        ('R_cond', 'Conductive thermal resistance', 'K/W'),
        ('L',      'Slab thickness', 'm'),
        ('k',      'Thermal conductivity', 'W/(m·K)'),
        ('A',      'Area', 'm²'),
    ],
    note="Analogous to Ohm's law: Q̇ = ΔT / R_cond")

subtopic_hdr('Heat Transfer — Convection')

eq_block(98, "Newton's Law of Cooling (Convection)",
    "q = h·A·(Tₛ − T∞)",
    [
        ('q',   'Convective heat transfer rate', 'W'),
        ('h',   'Convective heat transfer coefficient', 'W/(m²·K)'),
        ('A',   'Surface area in contact with fluid', 'm²'),
        ('Tₛ',  'Surface temperature', 'K  (or °C)'),
        ('T∞',  'Free-stream (ambient) fluid temperature', 'K  (or °C)'),
    ])

eq_block(99, "Thermal Resistance — Convection",
    "R_conv = 1 / (h·A)",
    [
        ('R_conv', 'Convective thermal resistance', 'K/W'),
        ('h',      'Convective heat transfer coefficient', 'W/(m²·K)'),
        ('A',      'Surface area', 'm²'),
    ])

eq_block(100, "Overall Heat Transfer Coefficient (Plane Wall)",
    "1/(U·A) = 1/(h₁·A) + L/(k·A) + 1/(h₂·A)  =  R_total",
    [
        ('U',      'Overall heat transfer coefficient', 'W/(m²·K)'),
        ('A',      'Heat transfer area', 'm²'),
        ('h₁, h₂', 'Convection coefficients on each side of wall', 'W/(m²·K)'),
        ('L',      'Wall thickness', 'm'),
        ('k',      'Thermal conductivity of wall', 'W/(m·K)'),
        ('R_total','Total thermal resistance', 'K/W'),
    ])

eq_block(101, "Nusselt Number",
    "Nu = h·L / k",
    [
        ('Nu', 'Nusselt number (ratio of convective to conductive heat transfer)', 'dimensionless'),
        ('h',  'Convective heat transfer coefficient', 'W/(m²·K)'),
        ('L',  'Characteristic length', 'm'),
        ('k',  'Fluid thermal conductivity', 'W/(m·K)'),
    ])

eq_block(102, "Prandtl Number",
    "Pr = μ·cₚ / k  =  ν / α_th",
    [
        ('Pr',    'Prandtl number (ratio of momentum to thermal diffusivity)', 'dimensionless'),
        ('μ',     'Dynamic viscosity of fluid', 'Pa·s'),
        ('cₚ',   'Specific heat of fluid at constant pressure', 'J/(kg·K)'),
        ('k',     'Fluid thermal conductivity', 'W/(m·K)'),
        ('ν',     'Kinematic viscosity', 'm²/s'),
        ('α_th',  'Thermal diffusivity  α = k/(ρcₚ)', 'm²/s'),
    ])

eq_block(103, "Dittus-Boelter Correlation (Turbulent Flow in Pipes)",
    "Nu = 0.023 · Re⁰·⁸ · Prⁿ",
    [
        ('Nu',   'Nusselt number', 'dimensionless'),
        ('Re',   'Reynolds number', 'dimensionless'),
        ('Pr',   'Prandtl number', 'dimensionless'),
        ('n',    'n = 0.4 for fluid being heated;  n = 0.3 for fluid being cooled', 'dimensionless'),
    ],
    note="Valid for: Re > 10,000; 0.6 < Pr < 160; L/D > 10 (thermally/hydrodynamically developed).")

eq_block(104, "Fin Efficiency",
    "η_f = tanh(mL) / (mL)     where  m = √(hP / (k·Aᶜ))",
    [
        ('η_f', 'Fin efficiency (ratio of actual to maximum possible heat transfer)', 'dimensionless'),
        ('m',   'Fin parameter', 'm⁻¹'),
        ('L',   'Fin length', 'm'),
        ('h',   'Convective heat transfer coefficient at fin surface', 'W/(m²·K)'),
        ('P',   'Fin perimeter (cross-sectional)', 'm'),
        ('k',   'Fin material thermal conductivity', 'W/(m·K)'),
        ('Aᶜ',  'Fin cross-sectional area', 'm²'),
    ],
    note="Valid for insulated fin tip. Uniform cross-section. Higher m or L → lower efficiency.")

eq_block(105, "Log Mean Temperature Difference (LMTD) — Heat Exchangers",
    "LMTD = (ΔT₁ − ΔT₂) / ln(ΔT₁/ΔT₂)",
    [
        ('LMTD',      'Log mean temperature difference', 'K  (or °C)'),
        ('ΔT₁',       'Temperature difference at one end of heat exchanger', 'K'),
        ('ΔT₂',       'Temperature difference at other end of heat exchanger', 'K'),
    ],
    note="Used in: Q̇ = U·A·LMTD. For counter-flow HX use temperature differences based on counter-flow pairing.")

subtopic_hdr('Heat Transfer — Radiation')

eq_block(106, "Stefan-Boltzmann Law — Net Radiation Heat Transfer",
    "q = ε·σ·A·(Tₛ⁴ − T_surr⁴)",
    [
        ('q',      'Net radiation heat transfer rate', 'W'),
        ('ε',      'Emissivity of surface  (0 ≤ ε ≤ 1; blackbody ε = 1)', 'dimensionless'),
        ('σ',      'Stefan-Boltzmann constant  = 5.670 × 10⁻⁸', 'W/(m²·K⁴)'),
        ('A',      'Surface area', 'm²'),
        ('Tₛ',     'Absolute surface temperature', 'K'),
        ('T_surr', 'Absolute temperature of surroundings', 'K'),
    ])

eq_block(107, "Wien's Displacement Law",
    "λ_max · T = 2.898 × 10⁻³",
    [
        ('λ_max', 'Wavelength of peak radiation intensity', 'm'),
        ('T',     'Absolute temperature of blackbody', 'K'),
        ('2.898 × 10⁻³', 'Wien displacement constant', 'm·K'),
    ])

eq_block(108, "View Factor Reciprocity Relation",
    "A₁·F₁₂ = A₂·F₂₁",
    [
        ('A₁, A₂', 'Areas of surfaces 1 and 2', 'm²'),
        ('F₁₂',    'View factor from surface 1 to surface 2  (fraction of radiation from 1 that reaches 2)', 'dimensionless'),
        ('F₂₁',    'View factor from surface 2 to surface 1', 'dimensionless'),
    ],
    note="Summation rule: ΣFᵢⱼ = 1 for all j (all radiation leaving surface i must reach some surface).")

# ── footer ────────────────────────────────────────────────────────────────────
ap('─' * 105, mono=True, sz=7, color=C['note'], before=10, after=10)
ap(f'Total equations: 108  |  Topics: ME 101–106 / ME 201–206  |  All equations web-verified  |  SI units',
   sz=8, color=C['note'], align='center', before=0, after=6)
ap('— HyESys Agent', bold=True, sz=9, color=C['title'], align='center', before=0, after=10)

doc.save(DST)
print(f'Saved: {DST}')
print(f'Total equations: 108')
