import fitz, shutil, os

template = r"C:\Users\JasonOng\Desktop\local docs\admin\hire\Interview Evaluation Form.pdf"
output   = r"C:\Users\JasonOng\Desktop\local docs\admin\hire\Interview_Evaluation_Form_Louis.pdf"

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
p0.insert_text((183, 191), "Louis",               fontname=FONT, fontsize=11, color=BLACK)
p0.insert_text((183, 213), "Mechanical Engineer", fontname=FONT, fontsize=11, color=BLACK)
p0.insert_text((183, 236), "17-06-26",            fontname=FONT, fontsize=11, color=BLACK)
p0.insert_text((183, 258), "Jason Ong",           fontname=FONT, fontsize=11, color=BLACK)

# ── Ratings ────────────────────────────────────────────────
# 1=Poor, 2=Below Avg, 3=Average, 4=Good, 5=Excellent
ROWS_P0 = [
    ((309.4, 323.5), 3),   # Educational Background — has mech eng fundamentals; baseline degree present
    ((390.4, 404.6), 2),   # Work Experience — SOP-only work, not much experience
    ((444.7, 458.8), 2),   # Technical Skills — fundamentals not practised since graduation
    ((498.9, 513.0), 2),   # Communication Skills — English struggles, hard to follow some points
    ((566.5, 580.7), 1),   # Problem-Solving Skills — purely SOP-based, no independent thinking
    ((619.9, 634.1), 3),   # Teamwork / Interpersonal — Jason said "good average"
    ((686.5, 700.7), 3),   # Attitude & Motivation — Jason said "average 6/10"
]
for (y0, y1), rating in ROWS_P0:
    tick(p0, CB[rating][0], y0, CB[rating][1], y1)

# ── Comments ───────────────────────────────────────────────
COMMENTS_P0 = [
    (fitz.Rect(415, 311, 521, 389),
     "Holds a mechanical engineering degree, providing the necessary academic foundation for the role. Core engineering fundamentals are present at a basic level."),
    (fitz.Rect(415, 392, 521, 443),
     "Current role is limited to executing standard SOPs with no exposure to independent engineering judgment. Limited overall industry experience."),
    (fitz.Rect(415, 446, 521, 497),
     "Academic fundamentals are recognisable but have not been developed since graduation. Lacks the applied depth expected for this role."),
    (fitz.Rect(415, 500, 521, 565),
     "English proficiency is below the level required; struggled to articulate some points clearly during the interview. Communication needs significant improvement for a professional engineering environment."),
    (fitz.Rect(415, 568, 521, 618),
     "Work only requires SOP execution with no need for independent thinking or troubleshooting. No evidence of applied problem-solving ability."),
    (fitz.Rect(415, 621, 521, 685),
     "Functions cooperatively in team settings and maintains an agreeable interpersonal manner. Teamwork is average — contributes without standing out."),
    (fitz.Rect(415, 688, 521, 753),
     "Attitude is acceptable but not particularly driven or proactive. Willing to complete assigned tasks but lacks the self-initiated motivation expected in a dynamic engineering role."),
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
    ((121.1, 135.2), 2),   # Adaptability — very quiet/new, not firm enough for everyday demands
    ((187.7, 201.8), 2),   # Appearance / Professionalism — innocent/quiet, lacks assertiveness for the role
]
for (y0, y1), rating in ROWS_P1:
    tick(p1, CB[rating][0], y0, CB[rating][1], y1)

# ── Comments ───────────────────────────────────────────────
COMMENTS_P1 = [
    (fitz.Rect(415, 122, 521, 186),
     "Very quiet and reserved with limited confidence. May struggle to adapt to the demands of a fast-paced environment without close guidance."),
    (fitz.Rect(415, 189, 521, 240),
     "Innocent and quiet appearance; lacks the assertiveness and firmness needed for day-to-day professional responsibilities in this role."),
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
if os.path.exists(output):
    os.remove(output)
os.rename(tmp, output)
print("Done — Interview_Evaluation_Form_Louis.pdf saved.")
