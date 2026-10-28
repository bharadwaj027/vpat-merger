"""Check that every run's rPr children follow the CT_RPr schema order (viewers may ignore out-of-order bold)."""
import sys
from docx import Document
from docx.oxml.ns import qn
sys.path.insert(0, "backend")
from vpat_merger.docx_utils import _RPR_ORDER

for path in sys.argv[1:]:
    doc = Document(path)
    bad = {}
    for rpr in doc.element.body.iter(qn("w:rPr")):
        names = [c.tag.split("}")[1] for c in rpr if c.tag.startswith("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}")]
        idx = [_RPR_ORDER.index(n) for n in names if n in _RPR_ORDER]
        if idx != sorted(idx):
            bad.setdefault(" ".join(names), 0)
            bad[" ".join(names)] += 1
    print(path, "out-of-order rPr:", bad or "none")
