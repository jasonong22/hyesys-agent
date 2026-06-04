import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SRC = r'C:\Users\JasonOng\Desktop\local docs\personal\PE\FEE2026_Mechanical_Solutions.docx'
DST = r'C:\Users\JasonOng\Desktop\local docs\personal\PE\FEE2026_Mechanical_Solutions_v2.docx'

# ─── XML helpers ─────────────────────────────────────────────────────────────

def make_para(text, bold=False, mono=False, color=None, shading=None,
              before=40, after=40, size_hp=18):
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), str(before))
    sp.set(qn('w:after'), str(after))
    pPr.append(sp)
    if shading:
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), shading)
        pPr.append(shd)
    p.append(pPr)
    if text:
        r = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        fn = 'Courier New' if mono else 'Calibri'
        rf = OxmlElement('w:rFonts')
        rf.set(qn('w:ascii'), fn); rf.set(qn('w:hAnsi'), fn)
        rPr.append(rf)
        sz = OxmlElement('w:sz'); sz.set(qn('w:val'), str(size_hp)); rPr.append(sz)
        szc = OxmlElement('w:szCs'); szc.set(qn('w:val'), str(size_hp)); rPr.append(szc)
        if bold:
            rPr.append(OxmlElement('w:b'))
        if color:
            cl = OxmlElement('w:color'); cl.set(qn('w:val'), color); rPr.append(cl)
        r.append(rPr)
        t = OxmlElement('w:t')
        t.text = text
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        r.append(t)
        p.append(r)
    return p


def ins_before(ref, new_p):
    ref.addprevious(new_p)


def inject_equations(ref_elem, equations):
    """Insert equations block (list of (name, formula)) before ref_elem."""
    ins_before(ref_elem, make_para('', before=30, after=10))
    ins_before(ref_elem, make_para(
        '▸  EQUATIONS USED', bold=True, color='1F3A8E',
        shading='DCE8F8', before=60, after=40, size_hp=19))
    for name, formula in equations:
        ins_before(ref_elem, make_para(
            f'   •  {name}:', bold=True, color='333333',
            shading='EDF3FC', before=20, after=5))
        ins_before(ref_elem, make_para(
            f'        {formula}', mono=True, color='0D0D7A',
            shading='EDF3FC', before=0, after=18))
    ins_before(ref_elem, make_para('', before=10, after=10))


def inject_diagram(ref_elem, lines):
    """Insert ASCII diagram block (list of strings) before ref_elem."""
    ins_before(ref_elem, make_para('', before=10, after=10))
    ins_before(ref_elem, make_para(
        '▸  DIAGRAM / SCHEMATIC', bold=True, color='14531A',
        shading='DFF0DC', before=50, after=20, size_hp=19))
    for line in lines:
        ins_before(ref_elem, make_para(
            line, mono=True, color='1A1A1A',
            shading='F2FAF2', before=0, after=0, size_hp=17))
    ins_before(ref_elem, make_para('', before=10, after=30))


# ─── Content definitions ─────────────────────────────────────────────────────

EQ = {
    # MCQ
    'Q1_MCQ': [
        ('Reynolds Number', 'Re = ρVD/μ  =  VD/ν'),
        ('Laminar pipe friction factor (Hagen-Poiseuille)', 'f = 64/Re  (valid for Re < 2300)'),
        ('Colebrook-White Equation (turbulent friction factor)', '1/√f = −2 log₁₀(ε/(3.7D) + 2.51/(Re√f))'),
        ('Fully-turbulent rough pipe (Nikuradse)', '1/√f = −2 log₁₀(ε/(3.7D))'),
    ],
    'Q2_MCQ': [
        ('Axial (direct) stress', 'σ_axial = P / A'),
        ('Second moment of area — rectangle', 'I_xx = b·h³ / 12'),
        ('Flexure formula (Euler-Bernoulli beam bending)', 'σ_bend = M·y / I'),
        ('Section modulus', 'Z = I / y_max  →  σ_max = M / Z  =  M·y_max / I'),
        ('Superposition principle', 'σ_total = σ_axial + σ_bend'),
    ],
    'Q3_MCQ': [
        ('Polar second moment of area — hollow shaft', 'J = π(d_o⁴ − d_i⁴) / 32'),
        ('Torsion formula (Coulomb)', 'τ_max = T·r / J  =  T·(d_o/2) / J'),
        ('Angle of twist (Coulomb)', 'φ = T·L / (G·J)  [radians]'),
        ('Shear modulus (definition)', 'G = τ / γ  (modulus of rigidity)'),
    ],
    'Q4_MCQ': [
        ('Nominal stress (net section)', 'σ_nom = P / A_net'),
        ('Maximum stress at geometric discontinuity', 'σ_max = K_t × σ_nom'),
        ('Stress conc. — elliptical notch (Inglis, 1913)', 'K_t ≈ 1 + 2√(a/ρ)  where a = half-length, ρ = tip radius'),
        ('Stress conc. — circular hole in plate (Kirsch)', 'K_t = 3  (uniaxial tension, hole diameter ≪ plate width)'),
    ],
    'Q5_MCQ': [
        ('Conservation of Linear Momentum (perfectly inelastic)', 'm_b·v_b + m_B·0 = (m_b + m_B)·v_f'),
        ('Work-Energy Theorem (sliding phase)', '½(m_total)v_f² = μ_k·m_total·g·d'),
        ('Post-impact velocity', 'v_f = √(2·μ_k·g·d)'),
        ('Kinetic friction force', 'F_k = μ_k × N  =  μ_k × m_total × g'),
        ("Newton's Second Law", 'ΣF = m × a'),
    ],
    'Q6_MCQ': [
        ('Kinematic equation (SUVAT) — velocity', 'v = u + a·t'),
        ('Kinematic equation (SUVAT) — displacement', 's = u·t + ½·a·t²'),
        ('Kinematic equation (SUVAT) — velocity-displacement', 'v² = u² + 2·a·s  →  t_acc = v/a,  d_acc = v²/(2a)'),
        ('Constant velocity (cruise phase)', 's = v × t  →  t_cruise = d_cruise / v_cruise'),
        ('Unit conversion', 'v [m/s] = v [km/h] / 3.6'),
    ],
    'Q7_MCQ': [
        ('Standard 2nd-order characteristic equation', 's² + 2ζω_n·s + ω_n² = 0'),
        ('Natural frequency (from constant term)', 'ω_n = √(constant term coefficient)'),
        ('Damping ratio (from s-coefficient)', 'ζ = (s-coefficient) / (2·ω_n)'),
        ('Quadratic formula', 's = (−b ± √(b²−4ac)) / (2a)'),
        ('Damping classification', 'ζ < 1: underdamped  |  ζ = 1: critically damped  |  ζ > 1: overdamped'),
    ],
    'Q8_MCQ': [
        ('First Law of Thermodynamics (closed system)', 'ΔU = Q − W'),
        ('State function property (cyclic process)', 'ΔU_cycle = 0  (system returns to initial state after full cycle)'),
        ('Energy balance for complete cycle', 'Q_net = W_net  (net heat in = net work out)'),
    ],
    'Q9_MCQ': [
        ('First Law of Thermodynamics', 'ΔU = Q − W'),
        ('Adiabatic condition (defining equation)', 'Q = 0  →  ΔU = −W'),
        ('Reversible adiabatic process — ideal gas (Poisson)', 'PV^γ = constant  (γ = c_p/c_v, heat capacity ratio)'),
        ('Temperature-volume adiabatic relation', 'T·V^(γ−1) = constant'),
    ],
    'Q10_MCQ': [
        ("Taylor's Tool Life Equation", 'V × T^n = C'),
        ('Solved for tool life T', 'T = (C / V)^(1/n)'),
        ("Taylor's equation — logarithmic form", 'log V + n·log T = log C  (straight line on log-log plot, slope = −n)'),
        ('Typical n values', 'HSS: n ≈ 0.10–0.15  |  Carbide: n ≈ 0.20–0.30  |  Ceramic: n ≈ 0.40–0.50'),
    ],
    # Long Q
    'LQ1': [
        ('Extended Bernoulli Equation (with head losses)', 'P₁/ρg + V₁²/2g + z₁  =  P₂/ρg + V₂²/2g + z₂ + Σh_L'),
        ('Darcy-Weisbach Equation (major/friction losses)', 'h_f = f × (L/D) × V²/(2g)'),
        ('Minor losses (fittings, entry, exit)', 'h_m = K × V²/(2g)'),
        ('Total head loss (combined form)', 'h_L = [f(L/D) + ΣK] × V²/(2g)'),
        ('Volumetric flow rate / continuity', 'Q = A × V  =  (πD²/4) × V'),
    ],
    'LQ2': [
        ('Power-Torque-Angular velocity', 'P = T × ω'),
        ('Angular velocity from rpm', 'ω = 2πN/60  [rad/s]'),
        ('Shear stress — solid shaft torsion (Coulomb)', 'τ_max = T·r/J  =  16T / (π·d³)'),
        ('Polar second moment — solid circular shaft', 'J = πd⁴/32'),
        ('Polar second moment — hollow circular shaft', 'J = π(d_o⁴ − d_i⁴)/32'),
        ('Shaft design from allowable shear stress (solid)', 'd³ = 16T / (π × τ_allow)  →  d = ∛[16T/(πτ_allow)]'),
    ],
    'LQ3': [
        ("Newton's Second Law (along incline)", 'ΣF = m × a  (applied to each block separately)'),
        ('Weight component parallel to incline', 'F∥ = m·g·sinθ'),
        ('Normal force perpendicular to incline', 'N = m·g·cosθ'),
        ('Kinetic friction force', 'F_k = μ_k × N  =  μ_k × m·g·cosθ'),
        ('Equations of motion (system + individual)', 'ΣF_sys = (m_A+m_B)·a  ;  m_A·g·sinθ − f_AB = m_A·a'),
    ],
    'LQ4': [
        ('Galvanic cell EMF', 'E_cell = E°_cathode − E°_anode  (spontaneous if > 0)'),
        ('Oxidation half-reaction (anode — Al corrodes)', 'Al  →  Al³⁺ + 3e⁻  (oxidation)'),
        ('Reduction half-reaction (cathode — Fe protected)', 'O₂ + 2H₂O + 4e⁻  →  4OH⁻  (reduction)'),
        ("Faraday's Law of Electrolysis (mass corroded)", 'm = (M × I × t) / (n × F)  where F = 96,485 C/mol'),
        ('Driving potential difference', 'ΔE = E°_Fe − E°_Al = −0.44 − (−0.76) = +0.32 V (spontaneous)'),
    ],
    'LQ5': [
        ('Plant transfer function (force → velocity)', 'G(s) = V(s)/F(s) = 1/(m·s) = 1/(1000s)'),
        ('PI controller transfer function', 'C(s) = K_p + K_i/s  =  (K_p·s + K_i)/s'),
        ('Open-loop transfer function', 'L(s) = C(s)·G(s) = (K_p·s + K_i) / (1000·s²)'),
        ('Closed-loop characteristic equation', '1 + L(s) = 0  →  1000s² + K_p·s + K_i = 0'),
        ('Coefficient matching to standard 2nd-order form', 'K_p = 2ζω_n·m  ;  K_i = ω_n²·m'),
        ('Steady-state driving force (incline)', 'F_ss = m·g·sinθ = 1000 × 9.81 × sin 20° = 3,355 N'),
    ],
    'LQ6': [
        ('Carnot efficiency (temperatures in Kelvin)', 'η_c = 1 − T_C / T_H'),
        ('Heat input rate from efficiency', 'Q_H = W_net / η_c  =  P_net / η_c'),
        ('Heat rejection (First Law for heat engine)', 'Q_C = Q_H − W_net'),
        ('Cooling water heat transfer', 'Q_C = ṁ·c_p·ΔT  →  ṁ = Q_C / (c_p·ΔT)'),
        ('Daily coal mass consumption', 'm_coal = Q_H [× 86,400 s] / HV  [MJ/kg]'),
    ],
    'LQ7': [
        ('Volume of material removed', 'V = L × W × d  (length × width × depth of cut)'),
        ('Cutting speed (peripheral)', 'V_c = π·D·N / 60,000  [m/s]  (D in mm, N in rpm)'),
        ('Table feed rate', 'v_f = f_z × z × N  [mm/min]  (feed/tooth × no. of teeth × rpm)'),
        ('Material Removal Rate (MRR)', 'MRR = v_f × d × W  [mm³/min]'),
        ('Machine power — MRR relationship', 'P = u × MRR  →  MRR_available = P / u'),
        ('Specific cutting energy (unit power)', 'u [J/m³] = material constant (given or from tables)'),
    ],
}

DG = {
    'Q1_MCQ': [
        '     VELOCITY PROFILES IN PIPE FLOW',
        '     ─────────────────────────────────────────────────',
        '     LAMINAR (Re < 2300)             TURBULENT (Re > 4000)',
        '     wall ───────────────          wall ───────────────',
        '           ►                                  ►►►',
        '         ►►►                              ►►►►►',
        '       ►►►►►►  ← u_max=2V          ►►►►►►►  ← u_max≈1.2V',
        '         ►►►                              ►►►►►',
        '           ►                                  ►►►',
        '     wall ───────────────          wall ───────────────',
        '     Parabolic profile               Blunted/flat profile',
        '     Moody chart: f = f(Re,ε/D) for turbulent; f = 64/Re for laminar',
    ],
    'Q2_MCQ': [
        '     COMBINED AXIAL + BENDING — STRESS DISTRIBUTION',
        '     ─────────────────────────────────────────────────',
        '     P=450kN →     b=100mm               Stress diagram:',
        '     ╔══════════════════╗  ← top fibre    |  σ=+70 MPa (MAX TENSION)',
        '     ║                   ║  } y=75mm       |  (30+40)',
        '     ║   ─ ─ NA ─ ─   ║  ← NA            |  σ=+30 MPa (axial only)',
        '     ║                   ║  } y=75mm       |',
        '     ╚══════════════════╝  ← bot fibre    |  σ=−10 MPa (30−40)',
        '           h=150mm',
        '     ─────────────────────────────────────────────────',
        '     σ_axial = 450000/15000 = 30 MPa (uniform)',
        '     σ_bend  = 15×10⁶×75/28125000 = 40 MPa (at extreme fibre)',
        '     σ_max   = 30 + 40 = 70 MPa ✔',
    ],
    'Q3_MCQ': [
        '     HOLLOW SHAFT — TORSION AND ANGLE OF TWIST',
        '     ─────────────────────────────────────────────────',
        '     T=150 N·m                           T=150 N·m',
        '     ↺  |←──────── L=500 mm ────────→|  ↻',
        '     ══════════════════════════════════════════════════',
        '     Cross-section (end view):',
        '     ┌──────────────────────────────┐',
        '     |    d_o=30mm                      |',
        '     |    ┌───────────────────┐          |',
        '     |    |  d_i=20mm (hollow) |          |',
        '     |    └───────────────────┘          |',
        '     └──────────────────────────────┘',
        '     J = π(30⁴−20⁴)/32 = 63,814 mm⁴',
        '     φ = TL/GJ = (150×10³×500)/(80000×63814) = 0.0147 rad = 0.84°',
    ],
    'Q4_MCQ': [
        '     STRESS CONCENTRATION — FORCE FLOW LINES',
        '     ─────────────────────────────────────────────────',
        '     ← σ₀ (far field)                          σ₀ →',
        '     ─────────────────────────────────────────────────',
        '                         ┌───┐',
        '     ─ ─ ─ ─ ─ ─ ───╮    ╭─────── ─ ─ ─ ─ ─ ─',
        '     ───────────╮      |      ╭───────────',
        '     ─────────────╮  hole  ╭────────────',
        '     lines crowd here           Kt=3 at hole edge',
        '     ───────────╯      |      ╰───────────',
        '     ─ ─ ─ ─ ─ ─ ───╰    ╰─────── ─ ─ ─ ─ ─ ─',
        '                         └───┘',
        '     σ_max = Kt×σ_nom  |  Sharp notch: Kt↑ (small ρ)  |  Shoulder: Kt = f(r/d)',
    ],
    'Q5_MCQ': [
        '     FREE BODY DIAGRAM — BULLET-BLOCK IMPACT & SLIDING',
        '     ─────────────────────────────────────────────────',
        '     PHASE 1: IMPACT (Conservation of Momentum)',
        '     ● ───────────────────►  + [■]  →  [■+●] v_f = 2.06 m/s',
        '     21 g @ v_b=200 m/s   m_B=2 kg at rest',
        '     m_b·v_b = (m_b+m_B)·v_f',
        '     ─────────────────────────────────────────────────',
        '     PHASE 2: SLIDING (Work-Energy Theorem)',
        '     [■+●] ─────────────►  →  STOP  (d=0.31 m)',
        '         ↑ N=(m_b+m_B)g              F_k=μ_k·N ←←←←',
        '         ground',
        '     ½(m_total)v_f² = μ_k·m_total·g·d',
        '     v_f = √(2×0.70×9.81×0.31) = 2.063 m/s  →  v_b = 198.5 ≈ 200 m/s',
    ],
    'Q6_MCQ': [
        '     VELOCITY-TIME (v-t) GRAPH',
        '     ─────────────────────────────────────────────────',
        '     v     27.78 m/s',
        '     (m/s)  ───────────────────────────────',
        '           /    ◄ acceleration      CRUISE (constant v)',
        '          /       phase',
        '         /        t_acc = 13.89 s',
        '        /         d_acc = 193 m',
        '       0────────────────────────────────────► t',
        '              t_acc                   t_total ≈ 3607 s',
        '     Area = d_acc (triangle) + d_cruise (rectangle) = 100 km total',
        '     t_total = 13.89 + 3593 = 3607 s ≈ 1.00 h',
    ],
    'Q7_MCQ': [
        '     POLE-ZERO PLOT (s-PLANE) — 2ND-ORDER OVERDAMPED SYSTEM',
        '     ─────────────────────────────────────────────────',
        '          jω',
        '           |',
        '           |                        s² + 3s + 2 = 0',
        '  ─────────┼────────── σ     roots: s = (−3 ± 1)/2',
        '     ×     |    ×              s₁ = −1  (pole)',
        '   s=-2   |   s=-1              s₂ = −2  (pole)',
        '           |',
        '     Both poles real, negative, distinct → OVERDAMPED',
        '     ζ = 3/(2√2) = 1.061 > 1  ✔',
        '     ─────────────────────────────────────────────────',
        '     Underdamped: poles complex conj (ζ<1) | Critically: repeated real (ζ=1)',
    ],
    'Q8_MCQ': [
        '     P-V DIAGRAM — CLOSED THERMODYNAMIC CYCLE',
        '     ─────────────────────────────────────────────────',
        '     P |',
        '       |    2────────────────3',
        '       |   /    Q_in ↑           \\',
        '       |  /      (heat in)         \\',
        '       | 1                          4',
        '       |  \\      W_net             /',
        '       |   \\   (enclosed area)   /',
        '       |    2────────────────1',
        '       └─────────────────────────► V',
        '     System returns to state 1 → ΔU = 0 (U is a state function)',
        '     First Law: ΔU = Q − W = 0  ∴  Q_net = W_net',
    ],
    'Q9_MCQ': [
        '     P-V DIAGRAM — ADIABATIC vs ISOTHERMAL EXPANSION',
        '     ─────────────────────────────────────────────────',
        '     P |',
        '       |  ×  ← same starting state',
        '       |   \\\\',
        '       |    \\\\ ← ADIABATIC: PV^γ=const  (steeper, γ>1)',
        '       |     \\\\    Q=0; T drops during expansion',
        '       |      \\ ───────── ← ISOTHERMAL: PV=const',
        '       |                    T=const; ΔU=0',
        '       └─────────────────────────► V',
        '     Adiabatic is STEEPER than isothermal (γ>1 for ideal gases)',
        '     Only Q=0 is required; W≠0; ΔU≠0; T varies → NOT isothermal',
    ],
    'Q10_MCQ': [
        "     TAYLOR'S TOOL LIFE CURVE (log-log scale)",
        '     ─────────────────────────────────────────────────',
        '     log V |',
        '           |  ×',
        '           |    \\  ← slope = -n',
        '           |     \\',
        '           |      ×',
        '           |       \\',
        '           |        ×',
        '           |         \\  VT^n = C',
        '           └───────────────────────► log T',
        '     V↑ → T↓ (tool life decreases with higher cutting speed)',
        '     V₁T₁^n = V₂T₂^n = C  (constant for given tool-workpiece combination)',
    ],
    'LQ1': [
        '     PIPE SYSTEM — TWO RESERVOIRS',
        '     ─────────────────────────────────────────────────',
        '     ┌──────────────┐  ← free surface (z₁, P=P_atm)',
        '     | Reservoir 1  |',
        '     └──────────────┘',
        '           |  D=100mm, L=200m, f=0.02',
        '           |  K_minor=1.5 (entry+exit+fittings)',
        '     Δz=10m|',
        '           └───────────────────────────────────────┐',
        '                                          ┌────────────┐',
        '                                          | Reservoir 2  |  ← free surface (z₂)',
        '                                          └────────────┘',
        '     Bernoulli: Δz = h_f + h_m = [f·L/D + K] × V²/2g',
        '     10 = [0.02×2000 + 1.5] × V²/19.62 = 41.5 × V²/19.62',
        '     V = √(10×19.62/41.5) = 2.17 m/s  →  Q = π/4×0.1²×2.17 = 17.1 L/s',
    ],
    'LQ2': [
        '     STEPPED SHAFT — LAYOUT AND TORQUE DIAGRAM',
        '     ─────────────────────────────────────────────────',
        '     MOTOR         GEAR A             GEAR B',
        '     10kW @ 100rpm  -5kW               -5kW',
        '     ↺              ↓                  ↓',
        '     O─────────────A───────────────B',
        '     |─── seg OA ───|──── seg AB ────|',
        '     T_OA = 955 N·m     T_AB = 477 N·m',
        '     TORQUE DIAGRAM:',
        '     T | 955',
        '       |───────────────┐',
        '       |               | 477',
        '       |               |─────────┐ 0',
        '       └──────O──────A───────B──► x',
        '     d_OA = ∛(16×955000/(π×60)) = 43.3 mm → use 45 mm',
        '     d_AB = ∛(16×477460/(π×60)) = 34.4 mm → use 35 mm',
    ],
    'LQ3': [
        '     FREE BODY DIAGRAMS — BLOCK A ON BLOCK B ON 30° INCLINE',
        '     ─────────────────────────────────────────────────',
        '     FBD Block A (2 kg):        FBD Block B (8 kg):',
        '       ↑ N_AB=m_A g cos30°        ↑ N_incline=(m_A+m_B)g cos30°',
        '      [A]                          [B]',
        '       ↓ m_A g sin30° (down)       ↓ (m_A+m_B)g sin30° (down)',
        '       → f_AB (friction from B)     ← f_incline=μ_k×N_incline',
        '                                    → f_AB (reaction, A pushes B)',
        '     ─────────────────────────────────────────────────',
        '     SYSTEM ON INCLINE:',
        '              ┌───────┐',
        '              |   A   |',
        '              |───────|  → direction of motion (down slope)',
        '              |   B   |',
        '              └───────┘',
        '     ──────────────────────  θ=30° frictionless incline',
        '     a = g sin30° − μ_k×g cos30° = 4.905 − 0.935 = 3.97 m/s²',
    ],
    'LQ4': [
        '     GALVANIC CELL — Al FASTENER vs Fe STRUCTURE (SALTWATER)',
        '     ─────────────────────────────────────────────────',
        '          e⁻ ←──────────────────────────',
        '     ┌──────────────┐         ┌──────────────┐',
        '     | ANODE (Al)  |         | CATHODE (Fe) |',
        '     | Al→Al³⁺+3e⁻ |         | O₂+2H₂O+4e⁻ |',
        '     | CORRODES    |         |  →4OH⁻       |',
        '     | E°=-0.76 V  |         | E°=-0.44 V  |',
        '     └──────────────┘         └──────────────┘',
        '          |                        |',
        '          └─── ELECTROLYTE ────────┘',
        '              (saltwater, Cl⁻)',
        '     ΔE = E°_Fe − E°_Al = −0.44−(−0.76) = +0.32V → spontaneous',
        '     Al fastener = anode → corrodes; Fe structure = cathode → protected',
    ],
    'LQ5': [
        '     FREE BODY DIAGRAM + CONTROL BLOCK DIAGRAM',
        '     ─────────────────────────────────────────────────',
        '     FBD — Vehicle on 20° incline:',
        '            N=mg cos20° ↑',
        '           ┌─────┐  → F_drive (PI output)',
        '           | Car  |',
        '           └─────┘',
        '         ↘ mg sin20° = 3355 N (gravity component down-slope)',
        '         ───────────────────── θ=20°',
        '     ─────────────────────────────────────────────────',
        '     CONTROL BLOCK DIAGRAM:',
        '             ┌─────────┐   F(s)  ┌────────┐',
        '     v_ref→Σ→| PI C(s) |──────►| 1/(ms) |─► v(s)',
        '          ↑  └─────────┘        └────────┘',
        '         -|                              |',
        '          └─────────────────────────────┘ (feedback)',
        '     PI integrator eliminates steady-state speed error ✔',
    ],
    'LQ6': [
        '     CARNOT HEAT ENGINE — BLOCK + T-s DIAGRAM',
        '     ─────────────────────────────────────────────────',
        '     T_H=823K    T-s DIAGRAM (Carnot):',
        '        ↓ Q_H      T |  T_H  1────────2  ← isothermal expansion (Q_H in)',
        '     ┌──────────┐        |    4│           │2  ← adiabatic expansion',
        '     | CARNOT   ├► W_net=750MW     |    │3           │1',
        '     | PLANT    |             |  T_C  3────────4  ← isothermal rejection (Q_C out)',
        '     └──────────┘        └────────────────────► s',
        '        ↓ Q_C',
        '     T_C=303K (cooling water: 20→35°C)',
        '     η_c = 1−303/823 = 63.2%  |  Q_H = 750/0.632 = 1187 MW',
        '     Q_C = 1187−750 = 437 MW  |  ṁ = 437×10⁶/(4180×15) = 6970 kg/s',
    ],
    'LQ7': [
        '     FACE MILLING — SCHEMATIC',
        '     ─────────────────────────────────────────────────',
        '        ↓ N=400rpm',
        '     [◉ cutter D=100mm, 8 teeth]',
        '     feed ►  ╔═════════════════════════════╗',
        '             ║  WORKPIECE (mild steel)  ║  300×200×25 mm',
        '             ║  5 mm depth to remove    ║',
        '             ╚═════════════════════════════╝',
        '             ↔─────── L=300mm ───────↔',
        '             ↕ d=5mm depth of cut',
        '     MRR_avail = P/u = 8000/(2.8×10⁹) = 2,857 mm³/s',
        '     MRR_reqd  = v_f×d×W = 10.67×5×100 = 5,335 mm³/s',
        '     5335 > 2857 → EXCEEDS power capacity → reduce f_z or W',
    ],
}

# ─── Insertion search map: key -> (header_text, diagram_trigger_text) ─────────

INSERTS = [
    # (key,  header_search,                         diagram_before_search)
    ('Q1_MCQ',  'Question 1 — Fluid Mechanics: Turbulent Flow',   'Background:  Turbulent flow'),
    ('Q2_MCQ',  'Question 2 — Mechanics of Materials',            'Step 1 — Cross-section properties'),
    ('Q3_MCQ',  'Question 3 — Torsion: Hollow Circular',          'Step 1 — Polar second moment'),
    ('Q4_MCQ',  'Question 4 — Solid Mechanics: Stress Conc',      'Analysis:  Stress concentration arises'),
    ('Q5_MCQ',  'Question 5 — Dynamics: Bullet-Block',            'Step 1 — Post-impact block velocity'),
    ('Q6_MCQ',  'Question 6 — Kinematics: Vehicle',               'Step 1 — Convert cruise speed'),
    ('Q7_MCQ',  'Question 7 — Control Systems: Damping',          'Step 1 — Standard 2nd-order form'),
    ('Q8_MCQ',  'Question 8 — Thermodynamics: First Law for',     'First Law of Thermodynamics:  '),
    ('Q9_MCQ',  'Question 9 — Thermodynamics: Adiabatic',         'Definition:  An adiabatic process'),
    ('Q10_MCQ', 'Question 10 — Manufacturing',                    "Taylor's Tool Life Equation:  "),
    ('LQ1',  'Question 1 — Fluid Mechanics: Pipe Flow with',  'Given:  D = 0.10 m'),
    ('LQ2',  'Question 2 — Solid Mechanics: Stepped Shaft',   'Step 1 — Motor torque at O'),
    ('LQ3',  'Question 3 — Dynamics: Two Blocks on Inclined', 'Assumption:  Incline'),
    ('LQ4',  'Question 4 — Materials: Corrosion',             '1. Galvanic Corrosion:  '),
    ('LQ5',  'Question 5 — Control Systems: PI Controller',   'Step 1 — Plant model'),
    ('LQ6',  'Question 6 — Thermodynamics: Steam Power',      'Step 1 — Carnot efficiency'),
    ('LQ7',  'Question 7 — Manufacturing: Machining',         'Step 1 — Volume of material removed'),
]

# ─── Main processing ──────────────────────────────────────────────────────────

doc = Document(SRC)

# Collect all target paragraph elements first (before any mutation)
header_elems = {}
diagram_elems = {}

for para in doc.paragraphs:
    t = para.text
    for key, h_srch, d_srch in INSERTS:
        if key not in header_elems and h_srch in t:
            header_elems[key] = para._p
        if key not in diagram_elems and d_srch in t:
            diagram_elems[key] = para._p

print(f"Headers found:  {len(header_elems)}/17")
print(f"Diagrams found: {len(diagram_elems)}/17")
for key, _, _ in INSERTS:
    hf = '✔' if key in header_elems else '✘'
    df = '✔' if key in diagram_elems else '✘'
    print(f"  {key:<10}  header={hf}  diagram={df}")

# Inject content (using stored element refs, so order doesn't matter)
for key, _, _ in INSERTS:
    if key in header_elems:
        inject_equations(header_elems[key], EQ[key])
    if key in diagram_elems:
        inject_diagram(diagram_elems[key], DG[key])

doc.save(DST)
print(f"\nSaved: {DST}")
