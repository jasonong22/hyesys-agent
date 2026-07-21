import json
with open("c:/Users/JasonOng/AST_Agent/elsevier_clickables.json", encoding="utf-8") as f:
    items = json.load(f)
toc = [x for x in items if "sc-eBHJIF" in x.get("cls", "")]
print(f"Total TOC buttons: {len(toc)}")
for i, x in enumerate(toc):
    label = x["text"].split("\n")[0]
    print(f"[{i:03d}] {label[:80]}")
