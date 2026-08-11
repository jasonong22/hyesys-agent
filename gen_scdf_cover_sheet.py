"""
Generate SCDF Material Submission Cover Sheet — 5 materials
Project: Boon Lay ST Engineering — Product X Deployment
Based on HiLT (Stanley) 8-point cover sheet guideline, 11 Aug 2026
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUTPUT_PATH = (
    r"C:\Users\JasonOng\OneDrive - Advancer Global Ltd\AST BD"
    r"\HyESys Dept\7. Client Projects\STeng\SCDF"
    r"\HySBatt_SCDF_Material_Cover_Sheet.docx"
)

NAVY = RGBColor(0x1F, 0x49, 0x7D)
RED  = RGBColor(0xCC, 0x00, 0x00)
GREY = RGBColor(0x80, 0x80, 0x80)
HDR_FILL  = "D9E1F2"
RED_FILL  = "FFE0E0"
BOX_FILL  = "F5F5F5"
NOTE_FILL = "FFFDE7"


# ── low-level helpers ────────────────────────────────────────────────────────

def _set_cell_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _hdr_row(tbl, labels, fill=HDR_FILL):
    row = tbl.rows[0]
    for cell, label in zip(row.cells, labels):
        cell.text = label
        cell.paragraphs[0].runs[0].bold = True
        _set_cell_bg(cell, fill)


def _body_rows(tbl, data, start=1):
    for i, row_data in enumerate(data):
        row = tbl.rows[start + i]
        for cell, val in zip(row.cells, row_data):
            cell.text = str(val)


# ── document-level helpers ───────────────────────────────────────────────────

def add_h1(doc, text):
    p = doc.add_heading(text, level=1)
    for run in p.runs:
        run.font.color.rgb = NAVY
    return p


def add_h2(doc, text):
    p = doc.add_heading(text, level=2)
    for run in p.runs:
        run.font.color.rgb = NAVY
    return p


def add_field(doc, label, value="[To be filled]"):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    r = p.add_run(f"{label}: ")
    r.bold = True
    r.font.size = Pt(11)
    rv = p.add_run(value)
    rv.font.size = Pt(11)
    rv.font.color.rgb = GREY


def add_note(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(f"Note: {text}")
    r.font.size = Pt(9)
    r.font.color.rgb = GREY


def add_red_flag(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(f"★  RED CLOUD REQUIRED — {text}")
    r.bold = True
    r.font.color.rgb = RED
    r.font.size = Pt(10)


def add_simple_table(doc, headers, rows, red_rows=None):
    """headers: list[str], rows: list[list[str]], red_rows: set of row indices (0-based body)"""
    red_rows = red_rows or set()
    tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
    tbl.style = "Table Grid"
    _hdr_row(tbl, headers)
    for i, row_data in enumerate(rows):
        row = tbl.rows[1 + i]
        for cell, val in zip(row.cells, row_data):
            cell.text = val
            if i in red_rows:
                _set_cell_bg(cell, RED_FILL)
        if i in red_rows:
            last = row.cells[-1]
            if last.paragraphs[0].runs:
                last.paragraphs[0].runs[0].font.color.rgb = RED
                last.paragraphs[0].runs[0].bold = True
    return tbl


def add_review_box(doc):
    add_h2(doc, "Review Comments")
    tbl = doc.add_table(rows=6, cols=1)
    tbl.style = "Table Grid"
    for row in tbl.rows:
        _set_cell_bg(row.cells[0], NOTE_FILL)
    tbl.rows[0].cells[0].paragraphs[0].add_run(
        "[Space for consultant review comments]"
    ).font.color.rgb = GREY


def add_approval_box(doc):
    add_h2(doc, "Approval Sign-Off")
    tbl = doc.add_table(rows=4, cols=3)
    tbl.style = "Table Grid"
    _hdr_row(tbl, ["Prepared By (Contractor)", "Reviewed By (Consultant)", "Approved By (Consultant)"])
    for label in [("Name:", "Name:", "Name:"),
                  ("Signature:", "Signature:", "Signature:"),
                  ("Date:", "Date:", "Date:")]:
        row = tbl.add_row()
        for cell, val in zip(row.cells, label):
            cell.text = val
            cell.paragraphs[0].runs[0].bold = True


# ── 8-point section builder ──────────────────────────────────────────────────

def section_1(doc, material_name, category, purpose="Material approval prior to installation — SCDF NOA"):
    add_h2(doc, "1. Material Description")
    add_field(doc, "Material Name", material_name)
    add_field(doc, "Material Type / Category", category)
    add_field(doc, "Submission Purpose", purpose)


def section_2(doc, fields=None):
    add_h2(doc, "2. Proposed Material — Dimensions, Sizes & Configuration")
    if fields:
        for label, val in fields:
            add_field(doc, label, val)
    else:
        add_field(doc, "Dimensions")
        add_field(doc, "Sizes / Thickness")
        add_field(doc, "Configuration / Arrangement")


def section_3(doc, system="[To be filled]", location="[To be filled]", interface="[To be filled]"):
    add_h2(doc, "3. System Description")
    add_field(doc, "Proposed System", system)
    add_field(doc, "Location of Installation", location)
    add_field(doc, "Interface with Building / Fire Safety Systems", interface)


def section_4(doc, manufacturer="[To be filled]", origin="[To be filled]", catalogue="[To be filled]"):
    add_h2(doc, "4. Technical Specifications & Product Data")
    add_field(doc, "Manufacturer", manufacturer)
    add_field(doc, "Country of Origin", origin)
    add_field(doc, "Product Catalogue / Datasheet Reference", catalogue)


def section_5(doc, rows=None):
    add_h2(doc, "5. Applicable Product Standards & Codes")
    default_rows = rows or [
        ("[Standard / Code]", "[Title / Description]", "[Compliant / Deviation]"),
        ("[Standard / Code]", "[Title / Description]", "[Compliant / Deviation]"),
        ("[Standard / Code]", "[Title / Description]", "[Compliant / Deviation]"),
    ]
    add_simple_table(doc, ["Standard / Code", "Title / Description", "Compliance Status"], default_rows)


def section_6(doc, sg_code="[To be filled]", deviation_rows=None):
    add_h2(doc, "6. Singapore Standards & Fire Code Compliance")
    add_field(doc, "Singapore Fire Code Clause(s)", sg_code)
    add_field(doc, "SS / CP Reference", "[To be filled]")
    doc.add_paragraph().add_run("Deviations:").bold = True
    deviation_rows = deviation_rows or [
        ("[Clause]", "[Description of Deviation]", "[Justification]"),
        ("[Clause]", "[Description of Deviation]", "[Justification]"),
    ]
    add_simple_table(doc, ["Clause / Requirement", "Description of Deviation", "Justification"], deviation_rows)


def section_7(doc, cert_rows, red_rows=None):
    add_h2(doc, "7. Test Reports, Certificates & Third-Party Approvals")
    add_simple_table(
        doc,
        ["Certificate / Report", "Issuing Body", "Reference / Doc No.", "Highlight"],
        cert_rows,
        red_rows=red_rows,
    )
    doc.add_paragraph()
    for i, row in enumerate(cert_rows):
        if i in (red_rows or set()):
            add_red_flag(doc, row[0])


def section_8(doc, rows=None):
    add_h2(doc, "8. Track Record of Local Projects")
    rows = rows or [
        ("[Project Name]", "[Address]", "[Year]", "[Scope / Application]", "[Reference Contact]"),
        ("[Project Name]", "[Address]", "[Year]", "[Scope / Application]", "[Reference Contact]"),
        ("[Project Name]", "[Address]", "[Year]", "[Scope / Application]", "[Reference Contact]"),
    ]
    add_simple_table(doc, ["Project Name", "Location", "Year", "Scope / Application", "Reference Contact"], rows)


# ── per-material builders ────────────────────────────────────────────────────

def build_h50_ess(doc):
    add_h1(doc, "Material 1: H50 Energy Storage System (HySBatt ESS)")

    section_1(doc,
        material_name="HySBatt H50 Energy Storage System (ESS)",
        category="Battery Energy Storage System — Lithium-based ESS")

    section_2(doc, fields=[
        ("Model", "HySBatt H50"),
        ("Physical Dimensions / Footprint", "[To be filled — refer to HySBatt Datasheet]"),
        ("Number of Battery Packs", "[To be filled]"),
        ("Configuration / Arrangement", "[Describe physical layout — e.g. dedicated ESS room, wall / floor mounted]"),
        ("IP Rating", "IP54"),
    ])

    section_3(doc,
        system="Energy Storage System (ESS) / Active Power Compensation System",
        location="[Specify room / zone — e.g. ESS Room, Level XX]",
        interface="Building MSB; flammable gas detection; fire detection & alarm; mechanical ventilation (fire-rated ductwork)")

    section_4(doc,
        manufacturer="Advancer Smart Technology Pte Ltd",
        origin="[To be filled]",
        catalogue="HySBatt Datasheet — May 2026, Version 2")

    section_5(doc, rows=[
        ("BS 476-20:1987", "Fire tests on building materials — Fire resistance (2HR)", "Compliant — 2HR certified"),
        ("UL 9540A", "Test Method for Evaluating Thermal Runaway Fire Propagation in Battery ESS (Unit Level)", "Compliant — Unit Level certified"),
        ("[Additional Standard]", "[Description]", "[Compliant / Deviation]"),
    ])

    section_6(doc, sg_code="[Cite applicable SCDF Fire Code clause(s) for ESS]")

    section_7(doc,
        cert_rows=[
            ("BS 476-20 2HR Fire Test Certificate", "[Test Lab]", "[Doc No.]", "YES — RED CLOUD"),
            ("UL 9540A Unit Level Test Report",      "[Test Lab]", "[Doc No.]", "YES — RED CLOUD"),
            ("HySBatt Datasheet",                   "Advancer Smart Technology Pte Ltd", "May 2026 v2", "No"),
        ],
        red_rows={0, 1},
    )

    section_8(doc)
    doc.add_paragraph()
    add_review_box(doc)
    doc.add_paragraph()
    add_approval_box(doc)


def build_generic_material(doc, number, name, category, system_hint="[To be filled]"):
    add_h1(doc, f"Material {number}: {name}")

    section_1(doc, material_name=name, category=category)

    section_2(doc)

    section_3(doc,
        system=system_hint,
        location="[To be filled]",
        interface="[Describe interface with fire safety / building systems]")

    section_4(doc)

    section_5(doc)

    section_6(doc)

    section_7(doc,
        cert_rows=[
            ("[Certificate / Test Report]", "[Issuing Body]", "[Doc No.]", "[Yes / No]"),
            ("[Certificate / Test Report]", "[Issuing Body]", "[Doc No.]", "[Yes / No]"),
        ],
        red_rows=set(),
    )

    section_8(doc)
    doc.add_paragraph()
    add_review_box(doc)
    doc.add_paragraph()
    add_approval_box(doc)


# ── main ─────────────────────────────────────────────────────────────────────

def build_doc():
    doc = Document()

    for section in doc.sections:
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # ── Cover page ──────────────────────────────────────────────────────────
    title = doc.add_heading("MATERIAL SUBMISSION COVER SHEET", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.color.rgb = NAVY

    sub = doc.add_paragraph("Boon Lay ST Engineering — Product X Deployment")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in sub.runs:
        run.bold = True
        run.font.size = Pt(13)
        run.font.color.rgb = NAVY

    doc.add_paragraph()

    info_tbl = doc.add_table(rows=5, cols=2)
    info_tbl.style = "Table Grid"
    for label, value in [
        ("Project",                  "Boon Lay ST Engineering — Product X Deployment"),
        ("Contractor / Installer",   "[To be filled]"),
        ("Consultant",               "HiLT Pte Ltd (A Jensen Hughes Company)"),
        ("Submission Date",          "11 August 2026"),
        ("Submission Reference No.", "[To be filled]"),
    ]:
        row = info_tbl.add_row()
        row.cells[0].text = label
        row.cells[0].paragraphs[0].runs[0].bold = True
        row.cells[1].text = value
    # remove the empty first row added by default
    tbl_el = info_tbl._tbl
    tbl_el.remove(tbl_el.tr_lst[0])

    doc.add_paragraph()

    # Materials index
    doc.add_paragraph().add_run("Materials Submitted:").bold = True
    for i, mat in enumerate([
        "H50 Energy Storage System (HySBatt ESS)",
        "Fire Rated Board",
        "Rockwool",
        "Pressure Relief Valve",
        "Activated Carbon Filter",
    ], 1):
        p = doc.add_paragraph(f"{i}.  {mat}", style="List Number")

    doc.add_page_break()

    # ── Material sections ────────────────────────────────────────────────────
    build_h50_ess(doc)
    doc.add_page_break()

    build_generic_material(doc, 2, "Fire Rated Board", "Passive Fire Protection — Fire Rated Board",
                           system_hint="Passive fire protection / fire-rated enclosure / wall lining")
    doc.add_page_break()

    build_generic_material(doc, 3, "Rockwool", "Passive Fire Protection — Mineral Wool Insulation",
                           system_hint="Passive fire protection / thermal and acoustic insulation")
    doc.add_page_break()

    build_generic_material(doc, 4, "Pressure Relief Valve", "Mechanical Safety Device — Pressure Relief Valve",
                           system_hint="ESS / battery room pressure management; mechanical ventilation")
    doc.add_page_break()

    build_generic_material(doc, 5, "Activated Carbon Filter", "Air Filtration — Activated Carbon Filter",
                           system_hint="ESS / battery room ventilation; gas / odour filtration")

    doc.save(OUTPUT_PATH)
    print(f"Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    build_doc()
