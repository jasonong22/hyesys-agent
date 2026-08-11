"""
Generate SCDF Material Submission Cover Sheet for HySBatt ESS
Project: Boon Lay ST Engineering — Product X Deployment
Based on HiLT (Stanley) cover sheet guidelines, 11 Aug 2026
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

OUTPUT_PATH = r"C:\Users\JasonOng\OneDrive - Advancer Global Ltd\AST BD\HyESys Dept\7. Client Projects\STeng\SCDF\HySBatt_SCDF_Material_Cover_Sheet.docx"


def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge, val in kwargs.items():
        child = OxmlElement(f"w:{edge}")
        for attr, v in val.items():
            child.set(qn(f"w:{attr}"), v)
        tcBorders.append(child)
    tcPr.append(tcBorders)


def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def add_heading(doc, text, level=1, color=None):
    para = doc.add_heading(text, level=level)
    if color:
        for run in para.runs:
            run.font.color.rgb = RGBColor(*color)
    return para


def add_field_row(doc, label, placeholder="[To be filled]", indent=True):
    para = doc.add_paragraph()
    if indent:
        para.paragraph_format.left_indent = Cm(0.5)
    run_label = para.add_run(f"{label}: ")
    run_label.bold = True
    run_label.font.size = Pt(11)
    run_ph = para.add_run(placeholder)
    run_ph.font.size = Pt(11)
    run_ph.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
    return para


def add_blank_box(doc, label, height_rows=4):
    """Add a labelled text box as a shaded table cell."""
    doc.add_paragraph().add_run(label).bold = True
    tbl = doc.add_table(rows=height_rows, cols=1)
    tbl.style = "Table Grid"
    for row in tbl.rows:
        for cell in row.cells:
            set_cell_bg(cell, "F5F5F5")
    tbl.rows[0].cells[0].paragraphs[0].add_run("")
    return tbl


def add_red_border_note(doc, text):
    """Add a note paragraph styled to indicate 'box in red cloud'."""
    para = doc.add_paragraph()
    run = para.add_run(f"★ RED CLOUD HIGHLIGHT REQUIRED: {text}")
    run.bold = True
    run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)
    run.font.size = Pt(10)
    return para


def build_doc():
    doc = Document()

    # --- Page margins ---
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    # =========================================================
    # COVER PAGE HEADER
    # =========================================================
    title = doc.add_heading("MATERIAL SUBMISSION COVER SHEET", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

    sub = doc.add_paragraph("Energy Storage System (ESS) — HySBatt")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in sub.runs:
        run.bold = True
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

    doc.add_paragraph()

    # Project info table
    info_tbl = doc.add_table(rows=5, cols=2)
    info_tbl.style = "Table Grid"
    info_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_data = [
        ("Project", "Boon Lay ST Engineering — Product X Deployment"),
        ("Contractor / Installer", "[To be filled]"),
        ("Consultant", "HiLT Pte Ltd (A Jensen Hughes Company)"),
        ("Submission Date", "11 August 2026"),
        ("Submission Reference No.", "[To be filled]"),
    ]
    for i, (k, v) in enumerate(info_data):
        row = info_tbl.rows[i]
        row.cells[0].text = k
        row.cells[0].paragraphs[0].runs[0].bold = True
        row.cells[1].text = v
    doc.add_paragraph()

    # =========================================================
    # SECTION 1 — Material Description (Cover Sheet Header)
    # =========================================================
    add_heading(doc, "Section 1: Material Description", level=1, color=(0x1F, 0x49, 0x7D))
    doc.add_paragraph(
        "This section provides the material description for inclusion on the cover sheet, "
        "together with space for review comments and the approval box."
    )
    add_field_row(doc, "Material Name", "HySBatt Energy Storage System (ESS)")
    add_field_row(doc, "Material Type / Category", "Battery Energy Storage — Lithium-based ESS")
    add_field_row(doc, "Submission Purpose", "Material approval prior to installation — SCDF NOA")
    doc.add_paragraph()

    # =========================================================
    # SECTION 2 — Material Specifications
    # =========================================================
    add_heading(doc, "Section 2: Material Specifications", level=1, color=(0x1F, 0x49, 0x7D))
    doc.add_paragraph(
        "Description of the proposed material, including dimensions, sizes and configuration."
    )
    add_field_row(doc, "Model", "HySBatt — [Specify: H30 / H50 / H60 / H100 / H125]")
    add_field_row(doc, "Rated Power (kVA)", "[To be filled]")
    add_field_row(doc, "Usable Energy (kWh)", "[To be filled]")
    add_field_row(doc, "Number of Battery Packs", "[To be filled]")
    add_field_row(doc, "Pack Dimensions (each)", "1,250 × 500 × 550 mm")
    add_field_row(doc, "Pack Weight (each)", "< 200 kg")
    add_field_row(doc, "Min. Gap Between Packs", "50 mm")
    add_field_row(doc, "IP Rating", "IP54")
    add_field_row(doc, "DC Operating Voltage Range", "[To be filled]")
    add_field_row(doc, "Configuration / Arrangement", "[Describe physical layout — e.g. wall-mounted, floor-standing, in dedicated room]")
    doc.add_paragraph()

    # =========================================================
    # SECTION 3 — System Description
    # =========================================================
    add_heading(doc, "Section 3: System Description", level=1, color=(0x1F, 0x49, 0x7D))
    doc.add_paragraph(
        "Description of the proposed system in which the material will be used."
    )
    add_field_row(doc, "System Type", "Energy Storage System (ESS) / Active Power Compensation System")
    add_field_row(doc, "Integration with Building Systems",
                  "Interface with building electrical MSB; flammable gas detection system; "
                  "fire detection & alarm system; mechanical ventilation (fire-rated ductwork)")
    add_field_row(doc, "Location of Installation",
                  "[Specify room / zone — e.g. Electrical Room, ESS Room, Level XX]")
    add_field_row(doc, "Fire Safety System Interaction",
                  "[Describe interface with fire alarm, suppression, and smoke control systems]")
    doc.add_paragraph()

    # =========================================================
    # SECTION 4 — Technical Specifications & Product Data
    # =========================================================
    add_heading(doc, "Section 4: Technical Specifications & Product Data", level=1, color=(0x1F, 0x49, 0x7D))
    doc.add_paragraph(
        "Including manufacturer, country of origin and product catalogue."
    )
    add_field_row(doc, "Manufacturer", "Advancer Smart Technology Pte Ltd")
    add_field_row(doc, "Brand", "HySBatt / HyESys")
    add_field_row(doc, "Country of Origin", "[To be filled]")
    add_field_row(doc, "Product Catalogue Reference", "HySBatt Datasheet — May 2026, Version 2")
    doc.add_paragraph(
        "★  Attach HySBatt Datasheet (physical and electrical specifications) as supporting document."
    ).runs[0].font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
    doc.add_paragraph()

    # =========================================================
    # SECTION 5 — Applicable Standards & Codes
    # =========================================================
    add_heading(doc, "Section 5: Applicable Product Standards & Codes", level=1, color=(0x1F, 0x49, 0x7D))
    doc.add_paragraph(
        "List all applicable product standards and codes to which the HySBatt ESS complies."
    )
    standards_tbl = doc.add_table(rows=6, cols=3)
    standards_tbl.style = "Table Grid"
    hdr = standards_tbl.rows[0]
    hdr.cells[0].text = "Standard / Code"
    hdr.cells[1].text = "Title / Description"
    hdr.cells[2].text = "Compliance Status"
    for cell in hdr.cells:
        cell.paragraphs[0].runs[0].bold = True
        set_cell_bg(cell, "D9E1F2")
    placeholders = [
        ("BS 476-20:1987", "Fire tests on building materials — Method for determination of fire resistance", "Compliant — 2HR certified"),
        ("UL 9540A", "Test Method for Evaluating Thermal Runaway Fire Propagation in Battery Energy Storage Systems (Unit Level)", "Compliant — Unit Level certified"),
        ("[Standard]", "[Description]", "[Compliant / Partial / Deviation — see Section 6]"),
        ("[Standard]", "[Description]", "[Compliant / Partial / Deviation — see Section 6]"),
        ("[Standard]", "[Description]", "[Compliant / Partial / Deviation — see Section 6]"),
    ]
    for i, (s, d, c) in enumerate(placeholders):
        row = standards_tbl.rows[i + 1]
        row.cells[0].text = s
        row.cells[1].text = d
        row.cells[2].text = c
    doc.add_paragraph()

    # =========================================================
    # SECTION 6 — Singapore Standards & Fire Code Compliance
    # =========================================================
    add_heading(doc, "Section 6: Singapore Standards & Fire Code Compliance", level=1, color=(0x1F, 0x49, 0x7D))
    doc.add_paragraph(
        "Compliance with the relevant Singapore standards and Fire Code requirements, "
        "with any deviations clearly identified and justified."
    )
    add_field_row(doc, "Singapore Fire Code Reference", "[Cite applicable SCDF Fire Code clause(s)]")
    add_field_row(doc, "SS / CP Reference", "[Cite applicable Singapore Standard(s)]")

    dev_heading = doc.add_paragraph()
    dev_heading.add_run("Deviations from Singapore Standards / Fire Code:").bold = True

    dev_tbl = doc.add_table(rows=4, cols=3)
    dev_tbl.style = "Table Grid"
    dev_hdr = dev_tbl.rows[0]
    dev_hdr.cells[0].text = "Clause / Requirement"
    dev_hdr.cells[1].text = "Description of Deviation"
    dev_hdr.cells[2].text = "Justification"
    for cell in dev_hdr.cells:
        cell.paragraphs[0].runs[0].bold = True
        set_cell_bg(cell, "D9E1F2")
    for i in range(1, 4):
        for cell in dev_tbl.rows[i].cells:
            cell.text = "[None / To be filled]"
    doc.add_paragraph()

    # =========================================================
    # SECTION 7 — Test Reports, Certificates & Third-Party Approvals
    # =========================================================
    add_heading(doc, "Section 7: Test Reports, Certificates & Third-Party Approvals", level=1, color=(0x1F, 0x49, 0x7D))
    doc.add_paragraph(
        "List all relevant test reports, certificates and third-party approvals. "
        "For the ESS, compile into one document. The two certificates below must be "
        "highlighted by boxing in RED CLOUD in the compiled submission package."
    )

    cert_tbl = doc.add_table(rows=5, cols=4)
    cert_tbl.style = "Table Grid"
    cert_hdr = cert_tbl.rows[0]
    cert_hdr.cells[0].text = "Certificate / Report"
    cert_hdr.cells[1].text = "Issuing Body"
    cert_hdr.cells[2].text = "Reference / Doc No."
    cert_hdr.cells[3].text = "Highlight Required"
    for cell in cert_hdr.cells:
        cell.paragraphs[0].runs[0].bold = True
        set_cell_bg(cell, "D9E1F2")

    cert_data = [
        ("BS 476-20 2HR Fire Test Certificate", "[Test Lab Name]", "[Doc No.]", "YES — RED CLOUD"),
        ("UL 9540A Unit Level Test Report", "[Test Lab Name]", "[Doc No.]", "YES — RED CLOUD"),
        ("HySBatt Datasheet", "Advancer Smart Technology Pte Ltd", "May 2026 v2", "No"),
        ("[Additional Certificate]", "[Issuing Body]", "[Doc No.]", "[Yes / No]"),
    ]
    for i, row_data in enumerate(cert_data):
        row = cert_tbl.rows[i + 1]
        for j, val in enumerate(row_data):
            row.cells[j].text = val
        if "RED CLOUD" in row_data[3]:
            for cell in row.cells:
                set_cell_bg(cell, "FFE0E0")
            row.cells[3].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xCC, 0x00, 0x00)
            row.cells[3].paragraphs[0].runs[0].bold = True

    doc.add_paragraph()
    add_red_border_note(doc, "BS 476-20 2HR Test Certificate — box this in red cloud in the compiled PDF")
    add_red_border_note(doc, "UL 9540A Unit Level Test Report — box this in red cloud in the compiled PDF")
    doc.add_paragraph()

    # =========================================================
    # SECTION 8 — Track Record of Local Projects
    # =========================================================
    add_heading(doc, "Section 8: Track Record of Local Projects", level=1, color=(0x1F, 0x49, 0x7D))
    doc.add_paragraph(
        "Track record of local projects where the same material or system has been "
        "successfully installed (where available)."
    )
    track_tbl = doc.add_table(rows=5, cols=5)
    track_tbl.style = "Table Grid"
    track_hdr = track_tbl.rows[0]
    track_hdr.cells[0].text = "Project Name"
    track_hdr.cells[1].text = "Location"
    track_hdr.cells[2].text = "Year Completed"
    track_hdr.cells[3].text = "Scope / Application"
    track_hdr.cells[4].text = "Reference Contact"
    for cell in track_hdr.cells:
        cell.paragraphs[0].runs[0].bold = True
        set_cell_bg(cell, "D9E1F2")
    for i in range(1, 5):
        for j, ph in enumerate(["[Project Name]", "[Address / Postal]", "[Year]", "[ESS / Reactive Comp.]", "[Name, Contact]"]):
            track_tbl.rows[i].cells[j].text = ph
    doc.add_paragraph()

    # =========================================================
    # REVIEW COMMENTS BOX
    # =========================================================
    add_heading(doc, "Consultant Review Comments", level=1, color=(0x1F, 0x49, 0x7D))
    review_tbl = doc.add_table(rows=8, cols=1)
    review_tbl.style = "Table Grid"
    for row in review_tbl.rows:
        set_cell_bg(row.cells[0], "FFFDE7")
        row.cells[0].height = Cm(1)
    review_tbl.rows[0].cells[0].paragraphs[0].add_run(
        "[Space for consultant review comments / queries / conditions of approval]"
    ).font.color.rgb = RGBColor(0x80, 0x80, 0x80)
    doc.add_paragraph()

    # =========================================================
    # APPROVAL BOX
    # =========================================================
    add_heading(doc, "Approval Sign-Off", level=1, color=(0x1F, 0x49, 0x7D))
    appr_tbl = doc.add_table(rows=4, cols=3)
    appr_tbl.style = "Table Grid"
    appr_hdr = appr_tbl.rows[0]
    appr_hdr.cells[0].text = "Prepared By (Contractor)"
    appr_hdr.cells[1].text = "Reviewed By (Consultant)"
    appr_hdr.cells[2].text = "Approved By (Consultant)"
    for cell in appr_hdr.cells:
        cell.paragraphs[0].runs[0].bold = True
        set_cell_bg(cell, "D9E1F2")

    appr_rows = [
        ("Name:", "Name:", "Name:"),
        ("Signature:", "Signature:", "Signature:"),
        ("Date:", "Date:", "Date:"),
    ]
    for i, row_data in enumerate(appr_rows):
        row = appr_tbl.rows[i + 1]
        for j, val in enumerate(row_data):
            cell = row.cells[j]
            cell.text = val
            cell.paragraphs[0].runs[0].bold = True

    # =========================================================
    # FOOTER NOTE
    # =========================================================
    doc.add_paragraph()
    footer_para = doc.add_paragraph(
        "Note: This cover sheet and supporting documents form part of the material submission "
        "to SCDF for the Boon Lay ST Engineering — Product X Deployment project. "
        "All deviations from Singapore Standards and Fire Code requirements must be clearly "
        "identified and justified before submission. The same cover sheet procedure applies "
        "to all other materials used for the project, especially fire safety materials."
    )
    footer_para.paragraph_format.left_indent = Cm(0)
    for run in footer_para.runs:
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)

    doc.save(OUTPUT_PATH)
    print(f"Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    build_doc()
