import fitz, os, shutil

src = r'C:\Users\JasonOng\Desktop\local docs\admin\apprasial\2026 mid\Re_ AST Mid-Year Performance Review 2026\PA1 - Performance Appraisal Form (Non-Executive).pdf'
dst = r'C:\Users\JasonOng\Desktop\local docs\admin\apprasial\2026 mid\Re_ AST Mid-Year Performance Review 2026\PA1_Jason_Ong_Mid_Year_2026.pdf'

shutil.copy2(src, dst)
doc = fitz.open(dst)

font = "helv"
fs = 10
black = (0, 0, 0)

# ====== PAGE 1 ======
p1 = doc[0]

# Header
p1.insert_text((155,  76.8), "Jason Ong Zong Yi", fontname=font, fontsize=fs, color=black)
p1.insert_text((420,  76.8), "System Engineer",   fontname=font, fontsize=fs, color=black)
p1.insert_text((625,  76.8), "HyESys",            fontname=font, fontsize=fs, color=black)
p1.insert_text((655, 105.6), "17 June 2026",      fontname=font, fontsize=fs, color=black)

# Mid-Year Review checkbox tick  (checkbox at x=131.1, y=140.8)
p1.insert_text((133, 140.5), "X", fontname=font, fontsize=9, color=black)

# Section A self-evaluation scores - Self-Evaluation column at x=275
scores = [
    (254.5, 5),   #  1. Communications
    (274.9, 5),   #  2. Effort and Output
    (295.1, 5),   #  3. Initiative
    (315.5, 5),   #  4. Service Focus
    (336.0, 5),   #  5. Quality Focus
    (356.2, 4),   #  6. Teamwork
    (376.6, 5),   #  7. Compliance
    (396.9, 5),   #  8. Discipline and Reliability
    (440.9, 5),   #  9. Job Knowledge
    (461.1, 5),   # 10. Technical Skills
    (484.7, 5),   # 11. Concern for Safety
    (508.2, 5),   # 12. Housekeeping / Professional Image
]
for (y, score) in scores:
    p1.insert_text((275, y), str(score), fontname=font, fontsize=fs, color=black)

# ====== PAGE 2 ======
p2 = doc[1]

# Disciplinary action - tick No  (text "No" at x=380.2, checkbox just before)
p2.insert_text((370, 73.5), "X", fontname=font, fontsize=9, color=black)

# -----------------------------------------------------------------------
# Notable Achievements - 2 active rows; rows 3 and 4 left blank
#
# Column x-boundaries (from text extraction):
#   Targets/KPIs:           x=38  - x=368   (330pt wide)
#   Period of completion:   x=372 - x=458   ( 86pt wide)
#   Measurement:            x=462 - x=613   (151pt wide)
#   Remarks:                x=617 - x=840   (left blank)
#
# Row y-centres (table header ~y=120-133, example row ~y=153, step ~30pt):
#   Row 1 (first blank): y=183
#   Row 2 (second blank): y=213
#   Rows 3 & 4:  empty
# -----------------------------------------------------------------------

# Row 1: Multi-Agent AI System + savings analysis methodology (old rows 1 & 3 merged)
# Period: 6-month future range from assessment date
target1  = ("Design and deploy HyESys Multi-Agent AI System for autonomous energy "
            "optimisation, incorporating validated per-site ML models and energy "
            "savings analysis")
period1  = "Jul 2026 - Dec 2026"
measure1 = "Energy savings methodology documented and deployed"

# Row 2: HySBatt data centre design requirements and validation
# Period: 1-year future range from assessment date
target2  = ("Develop HySBatt data centre design requirements and validate system "
            "for deployment, covering energy storage specifications and HyESys "
            "integration")
period2  = "Jul 2026 - Jun 2027"
measure2 = "Design requirements document completed"

row_data = [
    (183, target1, period1, measure1),
    (213, target2, period2, measure2),
]

# Column x-boundaries with 2pt inset on each side to keep text visibly inside borders
T_L, T_R  = 40,  366   # Targets/KPIs
P_L, P_R  = 374, 456   # Period
M_L, M_R  = 464, 611   # Measurement

for row_y, target, period, measure in row_data:
    # Vertical inset: 2pt top, 2pt bottom from cell boundary
    top = row_y - 13
    bot = row_y + 13

    p2.insert_textbox(fitz.Rect(T_L, top, T_R, bot),
                      target,  fontname=font, fontsize=6.5, color=black, align=0)
    p2.insert_textbox(fitz.Rect(P_L, top, P_R, bot),
                      period,  fontname=font, fontsize=7.0, color=black, align=1)
    p2.insert_textbox(fitz.Rect(M_L, top, M_R, bot),
                      measure, fontname=font, fontsize=6.5, color=black, align=0)

# -----------------------------------------------------------------------
# Employee's Future Goals (left half: x=38-453, y=312-445)
# Revised: keep items 1, 2; edit item 4 (remove Python); remove items 3 & 5
# -----------------------------------------------------------------------
goals = (
    "1. Complete Baoyuan experimental phase and demonstrate measured kW "
    "savings at MSB incomer.\n\n"
    "2. Live deployment of multi-agent system to additional HyESys "
    "customer sites.\n\n"
    "3. Deepen AI/ML skills for advanced predictive optimisation models."
)
p2.insert_textbox(fitz.Rect(38, 312, 453, 445),
                  goals, fontname=font, fontsize=8, color=black, align=0)

# ====== SAVE ======
tmp = dst + ".tmp"
doc.save(tmp)
doc.close()
os.replace(tmp, dst)
print("Saved:", dst)
