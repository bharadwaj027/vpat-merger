"""Report every platform-label run in a .docx and whether Word will render it bold."""
import re, sys
from docx import Document
from docx.oxml.ns import qn

LABEL = re.compile(r"^(Desktop Web|iOS Mobile Web|Android Mobile Web|Native iOS|Native Android):")

def effective_bold(doc, p, r):
    rpr = r.find(qn("w:rPr"))
    if rpr is not None and rpr.find(qn("w:b")) is not None:
        v = rpr.find(qn("w:b")).get(qn("w:val"))
        return v not in ("0", "false", "off"), "direct"
    if rpr is not None and rpr.find(qn("w:rStyle")) is not None:
        return None, "char-style " + rpr.find(qn("w:rStyle")).get(qn("w:val"))
    return False, "none"

for path in sys.argv[1:]:
    doc = Document(path)
    stats = {}
    for p in doc.element.body.iter(qn("w:p")):
        runs = [r for r in p.iter(qn("w:r")) if "".join(t.text or "" for t in r.iter(qn("w:t"))).strip()]
        if not runs:
            continue
        ptext = "".join(t.text or "" for t in p.iter(qn("w:t")))
        if not LABEL.match(ptext):
            continue
        first = runs[0]
        ftext = "".join(t.text or "" for t in first.iter(qn("w:t")))
        bold, why = effective_bold(doc, p, first)
        whole_label = LABEL.match(ftext) is not None and ftext.strip().endswith(":")
        key = ("bold" if bold else "NOT BOLD") + ("" if whole_label else " (label split/merged into run: %r)" % ftext[:40])
        stats.setdefault(key, []).append(ptext[:70])
    print(path)
    for k, v in stats.items():
        print(f"  {k}: {len(v)}")
        if k != "bold":
            for t in v[:8]:
                print("     ", t)
