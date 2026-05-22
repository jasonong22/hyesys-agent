import fitz

pdf_path = r"C:\Users\JasonOng\Desktop\local docs\admin\hire\Interview_Evaluation_Form_template.pdf"
doc = fitz.open(pdf_path)

WHITE = (1, 1, 1)

# ══════════════════════════════════════════════════════════
# PAGE 1
# ══════════════════════════════════════════════════════════
p0 = doc[0]

# Candidate info right column (Name / Position / Date / Interviewer)
p0.add_redact_annot(fitz.Rect(181, 173, 525, 263), fill=WHITE)

# Entire Comments column (right-most column of the criteria table)
p0.add_redact_annot(fitz.Rect(410, 305, 525, 756), fill=WHITE)

p0.apply_redactions()

# Tick marks — paint white over each small stroked rect
ticks_p0 = [
    fitz.Rect(361.0, 309.0, 371.5, 320.0),   # Educational Background: 4
    fitz.Rect(301.5, 393.5, 308.5, 400.5),   # Work Experience: 1
    fitz.Rect(301.5, 444.5, 311.5, 454.5),   # Technical Skills: 1
    fitz.Rect(300.0, 503.5, 310.5, 511.0),   # Communication Skills: 1
    fitz.Rect(320.5, 568.5, 328.0, 576.5),   # Problem-Solving Skills: 2
    fitz.Rect(340.0, 624.5, 350.5, 632.5),   # Teamwork / Interpersonal: 3
    fitz.Rect(343.5, 688.0, 351.5, 695.5),   # Attitude & Motivation: 3
]
for r in ticks_p0:
    p0.draw_rect(r, color=None, fill=WHITE)

# ══════════════════════════════════════════════════════════
# PAGE 2
# ══════════════════════════════════════════════════════════
p1 = doc[1]

# Comments column (Adaptability + Appearance rows on page 2)
p1.add_redact_annot(fitz.Rect(410, 115, 525, 245), fill=WHITE)

p1.apply_redactions()

# Tick marks on page 2
ticks_p1 = [
    fitz.Rect(344.5, 123.0, 351.5, 130.0),   # Adaptability: 3
    fitz.Rect(340.0, 192.5, 350.5, 200.0),   # Appearance / Professionalism: 3
    fitz.Rect(307.5, 289.0, 318.0, 297.0),   # Overall Impression: Poor
    fitz.Rect(388.5, 331.0, 399.5, 338.5),   # Recommended Action: Not Suitable
]
for r in ticks_p1:
    p1.draw_rect(r, color=None, fill=WHITE)

# Signature (paint over the handwritten signature area)
p1.draw_rect(fitz.Rect(100, 415, 380, 465), color=None, fill=WHITE)

tmp_path = pdf_path + ".tmp"
doc.save(tmp_path)
doc.close()
import os, shutil
os.replace(tmp_path, pdf_path)
print("Done.")
