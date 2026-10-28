"""Low-level WordprocessingML helpers.

Everything here works on the lxml elements behind python-docx so that tables,
lists, hyperlinks and run formatting survive the merge untouched.
"""
from __future__ import annotations

import copy
import io
import re
from typing import Iterable, Optional

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R_ATTRS = [f"{{{R_NS}}}id", f"{{{R_NS}}}embed", f"{{{R_NS}}}link", f"{{{R_NS}}}pict"]

# Elements that only make sense in their source document; dropped on import.
_STRIP_ON_IMPORT = {
    qn("w:bookmarkStart"),
    qn("w:bookmarkEnd"),
    qn("w:commentRangeStart"),
    qn("w:commentRangeEnd"),
    qn("w:commentReference"),
    qn("w:proofErr"),
}


def norm_space(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").replace(" ", " ")).strip()


def el_text(el) -> str:
    """Visible text of an element (paragraph, cell, table)."""
    parts = []
    for node in el.iter(qn("w:t"), qn("w:tab"), qn("w:br")):
        if node.tag == qn("w:t"):
            parts.append(node.text or "")
        else:
            parts.append(" ")
    return "".join(parts)


def para_style_id(p) -> str:
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        return ""
    ps = ppr.find(qn("w:pStyle"))
    return ps.get(qn("w:val")) if ps is not None else ""


def is_list_para(p) -> bool:
    return p.tag == qn("w:p") and p.find(f"{qn('w:pPr')}/{qn('w:numPr')}") is not None


def has_sectpr(el) -> bool:
    return el.tag == qn("w:p") and el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None


def is_pure_sectpr_para(el) -> bool:
    """An (empty) paragraph whose only job is to carry a section break."""
    return has_sectpr(el) and not norm_space(el_text(el)) and el.find(f".//{qn('w:drawing')}") is None


def cell_blocks(tc) -> list:
    """Block-level children (paragraphs and nested tables) of a table cell."""
    return [c for c in tc if c.tag in (qn("w:p"), qn("w:tbl"))]


def block_texts(blocks: Iterable) -> list[str]:
    """Normalized non-empty text for each block, used for integrity checks."""
    out = []
    for b in blocks:
        if b.tag == qn("w:tbl"):
            for p in b.iter(qn("w:p")):
                t = norm_space(el_text(p))
                if t:
                    out.append(t)
        else:
            t = norm_space(el_text(b))
            if t:
                out.append(t)
    return out


def strip_ids(el):
    """Drop w14:paraId/textId and source-only markers so Word sees no duplicates."""
    for node in el.iter():
        for attr in (f"{{{W14_NS}}}paraId", f"{{{W14_NS}}}textId"):
            if attr in node.attrib:
                del node.attrib[attr]
    for node in list(el.iter(*_STRIP_ON_IMPORT)):
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)
    return el


def first_run_rpr(p):
    for r in p.iter(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        if rpr is not None:
            return rpr
    ppr = p.find(qn("w:pPr"))
    if ppr is not None and ppr.find(qn("w:rPr")) is not None:
        return ppr.find(qn("w:rPr"))
    return None


# Typography carried from a template run onto generated text (labels, markers, titles). Character styles,
# underline, superscript etc. are not inherited so a label never picks up a link style or similar.
_TEMPLATE_KEEP = {"rFonts", "b", "bCs", "i", "iCs", "color", "sz", "szCs", "highlight", "shd", "lang"}


def make_run(text: str, rpr_template=None, bold: Optional[bool] = None, italic: bool = False):
    """A new run with the template's typography; bold=True/False sets bold explicitly."""
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    if rpr_template is not None:
        for child in rpr_template:
            if child.tag.startswith(f"{{{W_NS}}}") and child.tag.split("}")[1] in _TEMPLATE_KEEP:
                rpr.append(copy.deepcopy(child))
    if bold is not None:
        for tag in ("w:b", "w:bCs"):
            for n in rpr.findall(qn(tag)):
                rpr.remove(n)
        if bold:
            rpr.append(OxmlElement("w:b"))
            rpr.append(OxmlElement("w:bCs"))
    if italic and rpr.find(qn("w:i")) is None:
        rpr.append(OxmlElement("w:i"))
    normalize_rpr(rpr)
    if len(rpr):
        r.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    if text != text.strip():
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    return r


# CT_RPr children must appear in schema order; Word rejects out-of-order rPr.
_RPR_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
    "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
    "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
    "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
    "specVanish", "oMath",
]


def _rpr_rank(child) -> int:
    if not child.tag.startswith(f"{{{W_NS}}}"):
        return 900  # w14:/w15: extensions follow the schema elements
    name = child.tag.split("}")[1]
    if name in ("ins", "del", "moveFrom", "moveTo"):  # CT_ParaRPr revision marks come first
        return -1
    if name == "rPrChange":
        return 1000
    return _RPR_ORDER.index(name) if name in _RPR_ORDER else 900


def normalize_rpr(rpr):
    """Put run-property children in schema order and drop invalid empty <w:shd/>.

    The Deque-generated source VPATs write e.g. <w:rFonts> after <w:sz>. Desktop Word tolerates that, but
    stricter readers (Word Online, Office previews, Google Docs, LibreOffice) may discard the whole block,
    which would make the bold platform labels render as plain text.
    """
    for shd in rpr.findall(qn("w:shd")):
        if shd.get(qn("w:val")) is None:
            rpr.remove(shd)
    children = list(rpr)
    ordered = sorted(children, key=_rpr_rank)  # stable: equal ranks keep their order
    if ordered != children:
        for c in children:
            rpr.remove(c)
        for c in ordered:
            rpr.append(c)
    return rpr


def normalize_all_rpr(root):
    for rpr in root.iter(qn("w:rPr")):
        normalize_rpr(rpr)


def insert_runs_at_start(p, runs: list):
    """Insert runs right after the paragraph properties."""
    ppr = p.find(qn("w:pPr"))
    pos = list(p).index(ppr) + 1 if ppr is not None else 0
    for i, r in enumerate(runs):
        p.insert(pos + i, r)


def clear_runs(p):
    for child in list(p):
        if child.tag != qn("w:pPr"):
            p.remove(child)


def empty_paragraph_like(p):
    """A blank paragraph carrying p's paragraph properties (minus numbering / section breaks)."""
    new = OxmlElement("w:p")
    ppr = p.find(qn("w:pPr")) if p is not None else None
    if ppr is not None:
        ppr = copy.deepcopy(ppr)
        for tag in ("w:numPr", "w:sectPr", "w:pStyle"):
            for n in ppr.findall(qn(tag)):
                ppr.remove(n)
        new.append(ppr)
    return new


def paragraph_from_template(template_p, text_runs: list, keep_numbering: bool = True):
    """Clone a paragraph's properties and fill it with the given runs."""
    new = copy.deepcopy(template_p)
    strip_ids(new)
    ppr = new.find(qn("w:pPr"))
    if ppr is not None:
        for n in ppr.findall(qn("w:sectPr")):
            ppr.remove(n)
        if not keep_numbering:
            for n in ppr.findall(qn("w:numPr")):
                ppr.remove(n)
    clear_runs(new)
    for r in text_runs:
        new.append(r)
    return new


class Importer:
    """Copies elements from a source document into the destination document.

    Relationship ids (hyperlinks, images) and numbering ids are only meaningful
    inside their own package, so they are re-created in the destination.
    """

    def __init__(self, src_doc, dst_doc):
        self.src = src_doc
        self.dst = dst_doc
        self.same = src_doc is dst_doc
        self.rid_map: dict[str, Optional[str]] = {}
        self.num_map: dict[str, str] = {}
        self.warnings: list[str] = []
        if not self.same:
            self._build_numbering_map()

    def __call__(self, el):
        new = copy.deepcopy(el)
        strip_ids(new)
        if self.same:
            return new
        for node in new.iter():
            for attr in R_ATTRS:
                rid = node.get(attr)
                if rid is None:
                    continue
                new_rid = self._map_rid(rid)
                if new_rid is None:
                    del node.attrib[attr]
                else:
                    node.set(attr, new_rid)
            if node.tag == qn("w:numId"):
                val = node.get(qn("w:val"))
                if val in self.num_map:
                    node.set(qn("w:val"), self.num_map[val])
        return new

    def _map_rid(self, rid: str) -> Optional[str]:
        if rid in self.rid_map:
            return self.rid_map[rid]
        rel = self.src.part.rels.get(rid)
        new_rid = None
        if rel is None:
            self.warnings.append(f"Dangling relationship {rid} dropped")
        elif rel.is_external:
            new_rid = self.dst.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        elif rel.reltype == RT.IMAGE:
            new_rid, _ = self.dst.part.get_or_add_image(io.BytesIO(rel.target_part.blob))
        else:
            self.warnings.append(f"Unsupported embedded relationship ({rel.reltype}) dropped")
        self.rid_map[rid] = new_rid
        return new_rid

    def _build_numbering_map(self):
        try:
            src_num = self.src.part.numbering_part.element
        except Exception:
            return
        dst_num = self.dst.part.numbering_part.element

        def abstract_xml(numbering, num_id):
            num = next((n for n in numbering.findall(qn("w:num")) if n.get(qn("w:numId")) == num_id), None)
            if num is None:
                return None, None
            abs_id = num.find(qn("w:abstractNumId")).get(qn("w:val"))
            absn = next(
                (a for a in numbering.findall(qn("w:abstractNum")) if a.get(qn("w:abstractNumId")) == abs_id),
                None,
            )
            return num, absn

        def signature(absn):
            if absn is None:
                return None
            clone = copy.deepcopy(absn)
            for attr in list(clone.attrib):
                del clone.attrib[attr]
            for n in clone.findall(qn("w:nsid")) + clone.findall(qn("w:tmpl")):
                clone.remove(n)
            from lxml import etree

            return etree.tostring(clone)

        for num in src_num.findall(qn("w:num")):
            num_id = num.get(qn("w:numId"))
            _, src_abs = abstract_xml(src_num, num_id)
            _, dst_abs = abstract_xml(dst_num, num_id)
            if dst_abs is not None and signature(src_abs) == signature(dst_abs):
                self.num_map[num_id] = num_id
                continue
            # Import the abstract definition under fresh ids.
            existing_abs = [int(a.get(qn("w:abstractNumId"))) for a in dst_num.findall(qn("w:abstractNum"))]
            existing_num = [int(n.get(qn("w:numId"))) for n in dst_num.findall(qn("w:num"))]
            new_abs_id = str(max(existing_abs + [0]) + 1)
            new_num_id = str(max(existing_num + [0]) + 1)
            if src_abs is not None:
                abs_copy = copy.deepcopy(src_abs)
                abs_copy.set(qn("w:abstractNumId"), new_abs_id)
                first_num = dst_num.find(qn("w:num"))
                if first_num is not None:
                    first_num.addprevious(abs_copy)
                else:
                    dst_num.append(abs_copy)
            num_copy = copy.deepcopy(num)
            num_copy.set(qn("w:numId"), new_num_id)
            num_copy.find(qn("w:abstractNumId")).set(qn("w:val"), new_abs_id)
            dst_num.append(num_copy)
            self.num_map[num_id] = new_num_id
