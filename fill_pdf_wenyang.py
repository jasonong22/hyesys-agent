import fitz, shutil, os

template = r"C:\Users\JasonOng\Desktop\local docs\admin\hire\Interview Evaluation Form.pdf"
output   = r"C:\Users\JasonOng\Desktop\local docs\admin\hire\Interview_Evaluation_Form_Wenyang.pdf"

shutil.copy(template, output)
doc   = fitz.open(output)
BLACK = (0, 0, 0)
FONT  = "helv"

def tick(page, x0, y0, x1, y1, pad=2.0):
    r = fitz.Rect(x0 + pad, y0 + pad, x1 - pad, y1 - pad * 0.3)
    page.draw_rect(r, color=BLACK, fill=BLACK, width=0.5)

CB = {1:(300.7,310.2), 2:(320.7,330.2), 3:(340.7,350.2),
      4:(360.8,370.3), 5:(380.8,390.3)}

# ══════════════════════════════════════════════════════════
# PAGE 1
# ══════════════════════════════════════════════════════════
p0 = doc[0]

# ── Candidate information ──────────────────────────────────
p0.insert_text((183, 191), "Wenyang",              fontname=FONT, fontsize=11, color=BLACK)
p0.insert_text((183, 213), "Mechanical Engineer",  fontname=FONT, fontsize=11, color=BLACK)
p0.insert_text((183, 236), "25-04-26",             fontname=FONT, fontsize=11, color=BLACK)
p0.insert_text((183, 258), "Jason Ong",            fontname=FONT, fontsize=11, color=BLACK)

# ── Ratings ────────────────────────────────────────────────
# Criteria row checkbox y-ranges → assigned rating
# 1=Poor, 2=Below Avg, 3=Average, 4=Good, 5=Excellent
ROWS_P0 = [
    ((309.4, 323.5), 3),   # Educational Background — right degree but >5 yrs ago, outdated
    ((390.4, 404.6), 2),   # Work Experience — inconsistent 4 yrs, left 2 jobs, lack of experience
    ((444.7, 458.8), 2),   # Technical Skills — off-touched, steep learning curve
    ((498.9, 513.0), 3),   # Communication Skills — adequately conveyed background and circumstances
    ((566.5, 580.7), 2),   # Problem-Solving Skills — not in touch with industry, limited recent practice
    ((619.9, 634.1), 3),   # Teamwork / Interpersonal — explicitly shows teamwork capabilities
    ((686.5, 700.7), 3),   # Attitude & Motivation — determined to revise knowledge but commitment concern
]
for (y0, y1), rating in ROWS_P0:
    tick(p0, CB[rating][0], y0, CB[rating][1], y1)

# ── Comments ───────────────────────────────────────────────
COMMENTS_P0 = [
    (fitz.Rect(415, 311, 521, 389),
     "Holds the appropriate degree for the role, a positive indicator. However, graduation over 5 years ago combined with inconsistent employment has created a disconnect from current technical standards."),
    (fitz.Rect(415, 392, 521, 443),
     "Work history has been inconsistent for 4 years with two roles left due to family commitments. Lack of continuous experience limits practical depth for this role."),
    (fitz.Rect(415, 446, 521, 497),
     "Technical knowledge has drifted from extended periods away from work. Determination to revise skills is noted, but the gap is significant and the learning curve steep."),
    (fitz.Rect(415, 500, 521, 565),
     "Communicated his background and circumstances clearly during the interview. Verbal communication is adequate, though technical articulation requires further development."),
    (fitz.Rect(415, 568, 521, 618),
     "Employment gaps have reduced exposure to real-world problem-solving. Limited recent industry engagement suggests difficulty meeting technical demands of this role."),
    (fitz.Rect(415, 621, 521, 685),
     "Demonstrates teamwork capabilities and a collaborative disposition. Positive interpersonal conduct noted, though recent industry teamwork exposure is limited."),
    (fitz.Rect(415, 688, 521, 753),
     "Shows genuine determination to update technical knowledge, which is encouraging. Commitment concerns remain given the pattern of leaving previous roles for personal reasons."),
]
for rect, text in COMMENTS_P0:
    rc = p0.insert_textbox(rect, text, fontname=FONT, fontsize=6.0, color=BLACK, align=0)
    if rc < 0:
        print(f"WARNING: text overflow at {rect}")

# ══════════════════════════════════════════════════════════
# PAGE 2
# ══════════════════════════════════════════════════════════
p1 = doc[1]

# ── Ratings ────────────────────────────────────────────────
ROWS_P1 = [
    ((121.1, 135.2), 2),   # Adaptability — steep learning curve, not in touch with industry
    ((187.7, 201.8), 3),   # Appearance / Professionalism — adequate, family maturity noted
]
for (y0, y1), rating in ROWS_P1:
    tick(p1, CB[rating][0], y0, CB[rating][1], y1)

# ── Comments ───────────────────────────────────────────────
COMMENTS_P1 = [
    (fitz.Rect(415, 122, 521, 186),
     "Prolonged absence from consistent work raises concerns about adapting to current industry practices. The steep expected learning curve reflects limited recent exposure to evolving technical demands."),
    (fitz.Rect(415, 189, 521, 240),
     "Presented professionally and conducted himself respectfully during the interview. Personal maturity is evident from his responsibilities as a father of three."),
]
for rect, text in COMMENTS_P1:
    rc = p1.insert_textbox(rect, text, fontname=FONT, fontsize=6.0, color=BLACK, align=0)
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
if os.path.exists(output):
    os.remove(output)
os.rename(tmp, output)
print("Done — Interview_Evaluation_Form_Wenyang.pdf saved.")
