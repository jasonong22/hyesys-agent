"""
FDM6200 Assignment — concise version (<1500 words)
"""
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# Page margins
sec = doc.sections[0]
sec.page_width  = Inches(8.27)
sec.page_height = Inches(11.69)
for attr in ("left_margin","right_margin","top_margin","bottom_margin"):
    setattr(sec, attr, Inches(1.0))

# ── Helpers ───────────────────────────────────────────────────────────────────
def h1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(2)
    r = p.add_run(text)
    r.font.name = "Calibri"; r.font.size = Pt(12); r.bold = True
    r.font.color.rgb = RGBColor(31, 73, 125)

def body(text, indent=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    if indent:
        p.paragraph_format.left_indent = Inches(0.25)
    r = p.add_run(text)
    r.font.name = "Calibri"; r.font.size = Pt(10.5)

def eq(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after  = Pt(1)
    r = p.add_run(text)
    r.font.name = "Courier New"; r.font.size = Pt(10); r.italic = True

def calc(lines):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.cell(0,0)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),"clear"); shd.set(qn("w:color"),"auto"); shd.set(qn("w:fill"),"EFF3FB")
    cell._tc.get_or_add_tcPr().append(shd)
    for i, line in enumerate(lines):
        p = cell.paragraphs[0] if i==0 else cell.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(line)
        r.font.name = "Courier New"; r.font.size = Pt(9.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

def bullet(text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.left_indent = Inches(0.25)
    r = p.add_run(text)
    r.font.name = "Calibri"; r.font.size = Pt(10.5)

def simple_table(headers, rows_data, col_widths=None):
    tbl = doc.add_table(rows=len(rows_data)+1, cols=len(headers))
    tbl.style = "Light Shading Accent 1"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(headers):
        r = tbl.rows[0].cells[j].paragraphs[0].add_run(h)
        r.bold = True; r.font.size = Pt(9.5)
    for i, row in enumerate(rows_data, 1):
        for j, val in enumerate(row):
            tbl.rows[i].cells[j].paragraphs[0].add_run(val).font.size = Pt(9.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ══════════════════════════════════════════════════════════════════════════════
# TITLE
# ══════════════════════════════════════════════════════════════════════════════
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("FDM6200 Fire Dynamics — Take-Home Group Assignment")
r.font.name = "Calibri"; r.font.size = Pt(14); r.bold = True
r.font.color.rgb = RGBColor(31,73,125)

p2 = doc.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
r2 = p2.add_run("Effect of Ventilation Opening Size on Flashover in a Residential Compartment Fire")
r2.font.name = "Calibri"; r2.font.size = Pt(11); r2.italic = True

doc.add_paragraph()

# ══════════════════════════════════════════════════════════════════════════════
# (a) Introduction & Rationale
# ══════════════════════════════════════════════════════════════════════════════
h1("(a)  Introduction and Rationale")
body(
    "Flashover is the rapid transition from a localised fire to full-compartment ignition, "
    "characterised by near-simultaneous ignition of all exposed combustible surfaces. It is the "
    "most critical threshold in compartment fire development, beyond which occupant survival is "
    "near-zero. Workshop 2 defines four onset criteria: upper-layer gas temperature 500–600 °C; "
    "floor radiant heat flux ≥ 20 kW/m²; ignition of upper-layer gases; or flames at the opening."
)
body(
    "This experiment investigates how ventilation opening size affects the critical Heat Release "
    "Rate (HRR) for flashover in a Singapore HDB master bedroom. The ventilation factor Ao√Ho "
    "appears in all three standard flashover correlations (MQH, Babrauskas, Thomas), making it "
    "the key independent variable. Residents who partially close doors/windows directly alter "
    "this factor, with life-safety consequences. The experiment validates theoretical predictions "
    "against bench-scale measurements and identifies which correlation best represents a "
    "gypsum-lined residential compartment."
)

# ══════════════════════════════════════════════════════════════════════════════
# (b) Experimental Design
# ══════════════════════════════════════════════════════════════════════════════
h1("(b)  Experimental Design")
body(
    "A 1:5 Froude-scaled reduced-scale model (0.70 × 0.80 × 0.52 m) is used, representing an "
    "HDB master bedroom (3.5 × 4.0 × 2.6 m). Reduced scale is justified by practical resource "
    "and safety constraints; Froude scaling preserves fluid-dynamic similarity. All surfaces "
    "are lined with 13 mm gypsum plasterboard — the standard HDB wall finish, with well-"
    "characterised thermal properties (k = 0.48×10⁻³ kW/m·K, ρ = 1440 kg/m³, c = 0.84 kJ/kg·K). "
    "A propane burner with mass flow controller reproduces a Fast t² fire (αf = 0.0469 kW/s²), "
    "representing a PU foam mattress — the predominant fatal bedroom fuel."
)
body("The door is fixed open. The window is the independent variable across four conditions:")
simple_table(
    ["Condition", "Window (model)", "Ao,total (m²)", "Ao√Ho (m^5/2)"],
    [
        ["W0 — Door only",    "Closed",       "0.0756", "0.0490"],
        ["W1 — 25% open",     "0.24×0.05 m",  "0.0876", "0.0566"],
        ["W2 — 50% open",     "0.24×0.10 m",  "0.0996", "0.0644"],
        ["W3 — 100% open",    "0.24×0.20 m",  "0.1236", "0.0800"],
    ]
)
body("Instruments: Type-K thermocouple trees (TC1–TC3) at 5 heights; Schmidt-Boelter heat flux "
     "gauge at floor centre; mass loss scale under burner; two video cameras (30 fps).")

# ASCII diagram
tbl_d = doc.add_table(rows=1, cols=1)
tbl_d.style = "Table Grid"
cell_d = tbl_d.cell(0,0)
shd_d = OxmlElement("w:shd")
shd_d.set(qn("w:val"),"clear"); shd_d.set(qn("w:color"),"auto"); shd_d.set(qn("w:fill"),"F5F5F5")
cell_d._tc.get_or_add_tcPr().append(shd_d)
diagram = [
    "  PLAN VIEW — 1:5 Model (0.70 m × 0.80 m)",
    "  ┌─────────────────────────────────────────┐",
    "  │  [Burner/MLS]  TC1   TC2   TC3  [Window]│",
    "  │   (corner)      │     │     │   (variable)│",
    "  │                                           │",
    "  │  [HFG floor]               [CAM-1 side]  │",
    "  ├──┐                                        │",
    "  │Dr│ 0.18×0.42 m (fixed open)  [CAM-2 front]│",
    "  └──┴────────────────────────────────────────┘",
    "  ELEVATION: TC beads at z = 0.10, 0.21, 0.31, 0.42, 0.52 m",
]
for i, line in enumerate(diagram):
    p = cell_d.paragraphs[0] if i==0 else cell_d.add_paragraph()
    r = p.add_run(line); r.font.name = "Courier New"; r.font.size = Pt(9)
doc.add_paragraph().paragraph_format.space_after = Pt(2)
cap = doc.add_paragraph()
cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
cr = cap.add_run("Figure 1: Plan and elevation of 1:5 reduced-scale compartment model.")
cr.font.size = Pt(9); cr.italic = True

# ══════════════════════════════════════════════════════════════════════════════
# (c) Methodology
# ══════════════════════════════════════════════════════════════════════════════
h1("(c)  Methodology")
body("All calculations reference Workshop 2 correlations. Full-scale prototype geometry (W0 door-only ventilation):")

calc([
    "  AT  = 2(3.5×4.0) + 2(3.5×2.6) + 2(4.0×2.6) − (0.9×2.1)  =  65.11 m²",
    "  Ao  = 0.90 × 2.10 = 1.89 m²;   Ho = 2.10 m",
    "  Ao√Ho = 1.89 × √2.10 = 2.739 m^(5/2)",
])

body("Effective heat transfer coefficient for gypsum lining (Workshop 2, hk method):")
calc([
    "  tp  = (ρc/k)(δ/2)² = (1440×0.84/0.48e-3)(0.013/2)² = 106 s",
    "  tc  ≈ t_FO ≈ 183 s  →  tc > tp  →  hk = k/δ = 0.48e-3/0.013 = 0.0369 kW/(m²·K)",
])

body("Flashover HRR — three correlations (Workshop 2):")
eq("MQH:       Q̇_FO = 610(hk · AT · Ao√Ho)^½  =  610(0.0369×65.11×2.739)^½  =  1,565 kW")
eq("Babrauskas: Q̇_FO = 750 · Ao√Ho            =  750 × 2.739               =  2,054 kW")
eq("Thomas:    Q̇_FO = 7.8·AT + 378·Ao√Ho     =  507.9 + 1,035.3           =  1,543 kW")
body(
    "MQH (1,565 kW) is adopted as the design value — it explicitly accounts for lining thermal "
    "properties via hk, making it most physically complete. Thomas gives a near-identical result "
    "(1,543 kW); Babrauskas (2,054 kW) ignores lining properties and is least conservative."
)

body("Time to flashover using Fast t² growth (Workshop 1 & 2):")
eq("t_FO = √(Q̇_FO / αf) = √(1565 / 0.0469) = 183 s ≈ 3 minutes")

body("Verification — MQH pre-flashover temperature formula (Workshop 2):")
eq("ΔTg = 6.85[Q̇² / (hk·AT·Ao√Ho)]^(1/3) = 6.85[1565²/6.584]^(1/3) = 492 °C ≈ 500 °C  ✓")

body("Froude scaling to 1:5 model (Workshop 2):")
calc([
    "  Q̇_model = 1565 × (1/5)^2.5 = 28.0 kW",
    "  t_model  = 183  × √(1/5)   = 81.8 s",
])

body("Experimental steps: (1) Build and verify model dimensions. (2) Install instruments. "
     "(3) Calibrate propane burner HRR (10/20/30 kW reference points). (4) Set window condition. "
     "(5) Log data at 1 Hz; ignite and ramp HRR at model-scale αf trajectory. (6) Declare "
     "flashover per Section (d); shut off burner immediately. (7) Cool; repeat for conditions "
     "W1–W3; perform three replicates per condition.")

# ══════════════════════════════════════════════════════════════════════════════
# (d) Flashover Identification
# ══════════════════════════════════════════════════════════════════════════════
h1("(d)  Flashover Identification Method")
body(
    "Two simultaneous criteria are required (Workshop 2): "
    "(1) Upper-layer temperature ≥ 500 °C at ≥ 2 thermocouple trees (z ≥ 0.42 m model scale); "
    "(2) Floor heat flux ≥ 20 kW/m² at the Schmidt-Boelter gauge. "
    "Flashover time t_FO is recorded when both are met simultaneously. "
    "Video footage is reviewed post-run to confirm tertiary visual indicator: flames appearing "
    "at the door threshold (criterion iv, Workshop 2). Runs where only one criterion is met are "
    "flagged and repeated."
)

# ══════════════════════════════════════════════════════════════════════════════
# (e) Data Analysis Plan
# ══════════════════════════════════════════════════════════════════════════════
h1("(e)  Data Analysis Plan")
body(
    "For each ventilation condition, measured Q̇_FO,exp (from mass flow controller at t_FO) "
    "is compared against all three theoretical predictions; percentage error is calculated. "
    "Q̇_FO is plotted against Ao√Ho to determine whether the MQH (∝ (Ao√Ho)^½) or Babrauskas "
    "(∝ Ao√Ho) scaling best fits observations. Thermocouple data from TC2 is used to plot "
    "temperature profiles at all five heights vs. time, and the two-layer interface height is "
    "estimated per Workshop 2 zone model. Three replicates per condition; results reported as "
    "mean ± SD. One-way ANOVA (α = 0.05) tests statistical significance of ventilation effect."
)

# ══════════════════════════════════════════════════════════════════════════════
# (f) Safety and Ethical Considerations
# ══════════════════════════════════════════════════════════════════════════════
h1("(f)  Safety and Ethical Considerations")
bullet("All tests in a certified fire lab with forced ventilation (≥6 ACH); NEA-compliant.")
bullet("Automatic gas shut-off if upper-layer TC exceeds 600 °C or flashover is declared.")
bullet("Propane per run < 0.5 kg (below SCDF gas quantity threshold).")
bullet("Full PPE: fire-retardant coat, safety glasses, heat-resistant gloves; no synthetic clothing.")
bullet("Two-person rule: one operator, one safety observer at all times.")
bullet("CO monitor at operator station; lab evacuated if CO > 35 ppm (OSHA STEL).")
bullet("Combustion products extracted via activated carbon filter before room exhaust.")
bullet("Protocol approved by SIT Laboratory Safety Committee; no human or animal subjects.")

# ══════════════════════════════════════════════════════════════════════════════
# (g) Possible Experimental Errors
# ══════════════════════════════════════════════════════════════════════════════
h1("(g)  Possible Experimental Errors")
simple_table(
    ["Error", "Type", "Cause", "Mitigation"],
    [
        ["TC radiation over-read",      "Systematic", "Bead absorbs wall radiation → T_indicated > T_gas",               "Apply radiation correction; use aspirated TCs"],
        ["2D vent flow effects",         "Systematic", "Kawagoe Ao√Ho assumes 1D flow; not valid at small scale",        "Apply discharge coefficient Cd correction"],
        ["Stochastic flashover timing",  "Random",     "Gas burner lacks distributed pyrolysis of real fuel",            "3 replicates; report mean ± SD"],
        ["Lining thermal scale mismatch","Systematic", "Fourier conduction does not Froude-scale; hk ratio differs",     "Calculate corrected hk at model scale"],
        ["HFG position bias",            "Systematic", "Single gauge at centre; peak flux is near upper-layer flame",    "Add second gauge at corner position"],
    ]
)

# ══════════════════════════════════════════════════════════════════════════════
# (h) References
# ══════════════════════════════════════════════════════════════════════════════
h1("(h)  References")
refs = [
    "[1] McCaffrey B.J., Quintiere J.G., Harkleroad M.F. (1981). Estimating room temperatures "
    "and flashover likelihood. Fire Technology, 17(2), 98–119. [MQH correlation — Workshop 2]",
    "[2] Babrauskas V. (1980). Estimating room flashover potential. Fire Technology, 16(2), "
    "94–103. [Babrauskas correlation — Workshop 2]",
    "[3] Thomas P.H. (1981). Testing products for contribution to flashover. Fire and Materials, "
    "5(3), 103–111. [Thomas correlation — Workshop 2]",
    "[4] Drysdale D. (2011). An Introduction to Fire Dynamics, 3rd ed. Wiley. "
    "[t² growth rates; compartment fire stages — Workshop 1 & 2]",
    "[5] SFPE Handbook of Fire Protection Engineering, 5th ed. (2016). Springer. "
    "[hk methodology; Froude scaling]",
    "[6] EN 1991-1-2:2002. Eurocode 1: Actions on Structures — Exposure to Fire. CEN. "
    "[Fire load density — Workshop 2]",
]
for ref in refs:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.first_line_indent = Inches(-0.25)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(ref)
    r.font.name = "Calibri"; r.font.size = Pt(10)

# ══════════════════════════════════════════════════════════════════════════════
# Justification
# ══════════════════════════════════════════════════════════════════════════════
h1("Explanation and Justification")
body(
    "The HDB master bedroom was chosen because it represents Singapore's most common fatal fire "
    "scenario (SCDF data) with standardised geometry applicable across hundreds of thousands of "
    "units. A PU foam mattress with Fast t² growth (αf = 0.0469 kW/s²) represents the "
    "predominant residential fuel; Ultra-fast would imply an accelerant, which is outside scope. "
    "Ventilation size was selected as the independent variable because it is occupant-controllable "
    "(closing doors/windows) and appears in all three Workshop 2 correlations, allowing direct "
    "comparison. MQH was selected as the design correlation over Babrauskas because it explicitly "
    "models lining thermal conductivity through hk, which is critical for a gypsum-lined "
    "compartment; Babrauskas ignores lining properties entirely. The 1:5 scale yields a target "
    "model HRR of 28 kW, achievable with standard laboratory propane burners, without requiring "
    "a large-scale burn facility. The two-criterion flashover declaration (T ≥ 500 °C AND "
    "q_floor ≥ 20 kW/m²) cross-validates the thermal and radiative signatures, reducing "
    "false-positive declarations from instrument noise. Gypsum plasterboard was chosen as the "
    "lining because it matches the prototype material, is non-toxic, inexpensive, and its hk "
    "value is directly tabulated in Workshop 2 for calculation validation."
)

# Save
output_path = r"C:\Users\JasonOng\Desktop\FDM6200_Flashover_Assignment.docx"
doc.save(output_path)
print(f"Saved: {output_path}")
