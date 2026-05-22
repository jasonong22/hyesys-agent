import fitz

path = r"C:\Users\JasonOng\Desktop\local docs\admin\hire\Interview_Evaluation_Form_template.pdf"
doc = fitz.open(path)

for pno, page in enumerate(doc):
    print(f"\n=== PAGE {pno+1} ANNOTATIONS ===")
    for annot in page.annots():
        print(f"  type={annot.type}, rect={annot.rect}, info={annot.info}")

    print(f"\n=== PAGE {pno+1} DRAWINGS ===")
    drawings = page.get_drawings()
    for d in drawings:
        print(f"  rect={d['rect']}  fill={d['fill']}  color={d['color']}  type={d['type']}")
