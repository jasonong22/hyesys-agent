import pdfplumber
import json

path = r'C:\Users\JasonOng\Desktop\local docs\personal\SIT\Fire Dynamics Workshop 2.pdf'
pages_text = []

with pdfplumber.open(path) as pdf:
    total = len(pdf.pages)
    for i, page in enumerate(pdf.pages):
        text = page.extract_text() or ""
        pages_text.append({"page": i + 1, "text": text})

with open(r'C:\Users\JasonOng\AST_Agent\pdf_output.json', 'w', encoding='utf-8') as f:
    json.dump({"total_pages": total, "pages": pages_text}, f, ensure_ascii=False, indent=2)

print(f"Done. Extracted {total} pages.")
