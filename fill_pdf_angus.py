import fitz, shutil

template = r"C:\Users\JasonOng\Desktop\local docs\admin\hire\Interview Evaluation Form.pdf"
output   = r"C:\Users\JasonOng\Desktop\local docs\admin\hire\Interview_Evaluation_Form_Angus.pdf"

shutil.copy(template, output)
doc   = fitz.open(output)
BLACK = (0, 0, 0)
FONT  = "helv"

def tick(page, x0, y0, x1, y1, pad=2.0):
    r = fitz.Rect(x0 + pad, y0 + pad, x1 - pad, y1 - pad * 0.3)
    page.draw_rect(r, color=BLACK, fill=BLACK, width=0.5)

# Checkbox x-ranges (same on both pages)
CB = {1:(300.7,310.2), 2:(320.7,330.2), 3:(340.7,350.2),
      4:(360.8,370.3), 5:(380.8,390.3)}

# ══════════════════════════════════════════════════════════
# PAGE 1
# ══════════════════════════════════════════════════════════
p0 = doc[0]

# ── Candidate information ──────────────────────────────────
p0.insert_text((183, 191), "Angus",               fontname=FONT, fontsize=11, color=BLACK)
p0.insert_text((183, 213), "Mechanical Engineer", fontname=FONT, fontsize=11, color=BLACK)
p0.insert_text((183, 236), "18-05-26",            fontname=FONT, fontsize=11, color=BLACK)
p0.insert_text((183, 258), "Jason Ong",           fontname=FONT, fontsize=11, color=BLACK)

# ── Ratings ────────────────────────────────────────────────
# Row: (checkbox y0, checkbox y1) → rating
ROWS_P0 = [
    ((309.4, 323.5), 2),   # Educational Background
    ((390.4, 404.6), 1),   # Work Experience
    ((444.7, 458.8), 2),   # Technical Skills
    ((498.9, 513.0), 3),   # Communication Skills
    ((566.5, 580.7), 3),   # Problem-Solving Skills
    ((619.9, 634.1), 4),   # Teamwork / Interpersonal
    ((686.5, 700.7), 3),   # Attitude & Motivation
]
for (y0, y1), rating in ROWS_P0:
    tick(p0, CB[rating][0], y0, CB[rating][1], y1)

# ── Comments ───────────────────────────────────────────────
COMMENTS_P0 = [
    (fitz.Rect(415, 311, 521, 389),
     "Holds a degree in Building Engineering, which is not directly aligned with the Mechanical Engineer role. Educational foundation is present but lacks the field-specific qualifications required."),
    (fitz.Rect(415, 392, 521, 443),
     "Only part-time work experience since graduating last year with no involvement in the mechanical engineering field. Insufficient industry exposure for this role."),
    (fitz.Rect(415, 446, 521, 497),
     "Shows some awareness of PID controllers and sensors from past academic projects. However, the specific technical skillset required for this role is largely absent."),
    (fitz.Rect(415, 500, 521, 565),
     "Able to communicate project outcomes and experimental findings at a basic level. Technical articulation needs further development to meet the expectations of this role."),
    (fitz.Rect(415, 568, 521, 618),
     "Showed initiative and proactiveness in managing project challenges with subsequent results produced. Gaps and mistakes in the work indicate limited problem-solving maturity."),
    (fitz.Rect(415, 621, 521, 685),
     "Identified as a strong team player who collaborates well with others. Contributed positively to group efforts and maintained good working relationships in past projects."),
    (fitz.Rect(415, 688, 521, 753),
     "Puts in extra effort and displays borderline proactiveness to ensure project quality. Motivation is present but not strongly directed towards the specific demands of this role."),
]
for rect, text in COMMENTS_P0:
    rc = p0.insert_textbox(rect, text, fontname=FONT, fontsize=6.5, color=BLACK, align=0)
    if rc < 0:
        print(f"WARNING: text overflow at {rect}")

# ══════════════════════════════════════════════════════════
# PAGE 2
# ══════════════════════════════════════════════════════════
p1 = doc[1]

# ── Ratings ────────────────────────────────────────────────
ROWS_P1 = [
    ((121.1, 135.2), 3),   # Adaptability
    ((187.7, 201.8), 3),   # Appearance / Professionalism
]
for (y0, y1), rating in ROWS_P1:
    tick(p1, CB[rating][0], y0, CB[rating][1], y1)

# ── Comments ───────────────────────────────────────────────
COMMENTS_P1 = [
    (fitz.Rect(415, 122, 521, 186),
     "Demonstrates some capacity to adapt through experimental project work. Limited professional engineering exposure restricts a fuller assessment of adaptability under real work pressure."),
    (fitz.Rect(415, 189, 521, 240),
     "Presented with an appropriate level of professionalism expected of a junior candidate. Overall conduct during the interview was acceptable."),
]
for rect, text in COMMENTS_P1:
    rc = p1.insert_textbox(rect, text, fontname=FONT, fontsize=6.5, color=BLACK, align=0)
    if rc < 0:
        print(f"WARNING: text overflow at {rect}")

# ── Overall Assessment ─────────────────────────────────────
# Overall Impression: Fair   [269.5,284.3,279.0,298.4]
tick(p1, 269.5, 284.3, 279.0, 298.4)

# Recommended Action: Not Suitable   [389.3,326.0,398.7,340.1]
tick(p1, 389.3, 326.0, 398.7, 340.1)

tmp = output + ".tmp"
doc.save(tmp)
doc.close()
import os
if os.path.exists(output):
    os.remove(output)
os.rename(tmp, output)
print("Done — Interview_Evaluation_Form_Angus.pdf saved.")
