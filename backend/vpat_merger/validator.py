"""Re-open the generated document and verify it against the parsed sources.

Successful file creation is not treated as proof of a correct merge: every
criterion, conformance value, remarks paragraph and scope entry is read back
from the output and compared with the source it came from.
"""
from __future__ import annotations

import io
from dataclasses import dataclass, field

from docx import Document
from docx.oxml.ns import qn

from .docx_utils import R_ATTRS, _rpr_rank, block_texts, cell_blocks, el_text, is_list_para, norm_space
from .matching import MergedTable
from .models import MISSING_MARKER, Platform, SourceVpat, active_platforms
from .parser import SC_RE, is_criteria_header
from .scope import page_key, scope_comment

MAX_DETAILS = 25


@dataclass
class Check:
    name: str
    passed: bool = True
    details: list[str] = field(default_factory=list)

    def fail(self, msg: str):
        self.passed = False
        if len(self.details) < MAX_DETAILS:
            self.details.append(msg)

    def to_dict(self):
        return {"name": self.name, "passed": self.passed, "details": self.details}


def split_by_label(tc, platforms: list[Platform]) -> tuple[dict[Platform, list], list[str]]:
    """Split a merged cell into per-platform block lists using the bold labels."""
    labels = {f"{p.label}:": p for p in platforms}
    parts: dict[Platform, list] = {}
    order: list[str] = []
    current = None
    for b in cell_blocks(tc):
        if b.tag == qn("w:p"):
            runs = [r for r in b.findall(qn("w:r")) if norm_space(el_text(r))]
            if runs:
                first = runs[0]
                rpr = first.find(qn("w:rPr"))
                is_bold = rpr is not None and rpr.find(qn("w:b")) is not None
                text = norm_space(el_text(first))
                if is_bold and text in labels:
                    current = labels[text]
                    order.append(current.value)
                    parts.setdefault(current, []).append(("label", b, first))
                    continue
        if current is not None:
            parts[current].append(("block", b, None))
    return parts, order


def platform_texts(entries) -> list[str]:
    """Visible text of a platform's part, with the label stripped from its first paragraph."""
    out = []
    for kind, b, label_run in entries:
        if kind == "label":
            text = norm_space(el_text(b))
            rest = norm_space(text[len(norm_space(el_text(label_run))):])
            if rest:
                out.append(rest)
        else:
            out.extend(block_texts([b]))
    return out


def list_count(entries) -> int:
    return sum(1 for kind, b, _ in entries if kind == "block" and is_list_para(b) and norm_space(el_text(b)))


def validate_output(data: bytes, sources: dict[Platform, SourceVpat], merged_tables: list[MergedTable],
                    scope_differences=()) -> list[Check]:
    checks: list[Check] = []
    platforms = active_platforms(sources)

    reopen = Check("Generated document reopens as a valid Word file")
    try:
        doc = Document(io.BytesIO(data))
    except Exception as exc:  # pragma: no cover - reported to the user
        reopen.fail(f"Could not reopen output: {exc}")
        return [reopen]
    checks.append(reopen)

    rels = Check("All hyperlinks and images point to valid relationships")
    for node in doc.element.body.iter():
        for attr in R_ATTRS:
            rid = node.get(attr)
            if rid is not None and rid not in doc.part.rels:
                rels.fail(f"Missing relationship {rid}")
    checks.append(rels)

    order = Check("Run formatting is valid for all Word viewers (bold labels display everywhere)")
    for rpr in doc.element.body.iter(qn("w:rPr")):
        ranks = [_rpr_rank(c) for c in rpr]
        if ranks != sorted(ranks):
            order.fail("Out-of-order run properties: " + " ".join(c.tag.split("}")[1] for c in rpr))
    checks.append(order)

    bold = Check("Every platform prefix is bold")
    prefixes = tuple(f"{p.label}:" for p in platforms)
    for p in doc.element.body.iter(qn("w:p")):
        text = norm_space(el_text(p))
        if not text.startswith(prefixes):
            continue
        first = next((r for r in p.iter(qn("w:r")) if norm_space(el_text(r))), None)
        rpr = first.find(qn("w:rPr")) if first is not None else None
        b = rpr.find(qn("w:b")) if rpr is not None else None
        if b is None or b.get(qn("w:val")) in ("0", "false", "off") or not norm_space(el_text(first)) in prefixes:
            bold.fail(f"Prefix not bold: '{text[:60]}'")
    checks.append(bold)

    out_tables = [t for t in doc.element.body.iter(qn("w:tbl")) if is_criteria_header(t)]
    structure = Check("Conformance tables have the three required columns")
    if len(out_tables) != len(merged_tables):
        structure.fail(f"Expected {len(merged_tables)} conformance tables, found {len(out_tables)}")
    for t in out_tables:
        header = [norm_space(el_text(c)) for c in t.findall(qn("w:tr"))[0].findall(qn("w:tc"))]
        if len(header) != 3:
            structure.fail(f"Header has {len(header)} columns: {header}")
    checks.append(structure)

    rows_check = Check("Every criterion appears exactly once, in the right table")
    conf_check = Check("Conformance levels match each platform's source")
    remarks_check = Check("Remarks match each platform's source, under the correct label")
    labels_check = Check("Platform labels are present and consistent in every merged cell")
    lists_check = Check("Bullet/numbered lists in remarks are preserved")
    empty_check = Check("No merged cell is empty")

    seen_keys: set[str] = set()
    for mt, tbl in zip(merged_tables, out_tables):
        rows = tbl.findall(qn("w:tr"))[1:]
        if len(rows) != len(mt.criteria):
            rows_check.fail(f"Table '{mt.key}': expected {len(mt.criteria)} rows, found {len(rows)}")
        for mc, tr in zip(mt.criteria, rows):
            cells = tr.findall(qn("w:tc"))
            name = mc.display_row.title
            if len(cells) != 3:
                rows_check.fail(f"{name}: row has {len(cells)} cells")
                continue
            crit_text = norm_space(el_text(cells[0]))
            m = SC_RE.match(crit_text)
            out_key = m.group(1) if m else "text:" + crit_text.lower()
            if out_key != mc.key:
                rows_check.fail(f"Row order mismatch: expected {name}, found {crit_text}")
            if out_key in seen_keys:
                rows_check.fail(f"{crit_text} appears more than once")
            seen_keys.add(out_key)
            for c in cells:
                if not norm_space(el_text(c)):
                    empty_check.fail(f"{name}: a cell is empty")

            conf_parts, conf_order = split_by_label(cells[1], platforms)
            rem_parts, rem_order = split_by_label(cells[2], platforms)
            expected_order = [p.value for p in platforms]
            if conf_order != expected_order:
                labels_check.fail(f"{name}: conformance labels are {conf_order}")
            if rem_order != expected_order:
                labels_check.fail(f"{name}: remarks labels are {rem_order}")

            for p in platforms:
                src = mc.rows.get(p)
                got_conf = " ".join(platform_texts(conf_parts.get(p, [])))
                want_conf = src.conformance if (src is not None and src.conformance) else MISSING_MARKER
                if norm_space(got_conf) != norm_space(want_conf):
                    conf_check.fail(f"{name} / {p.label}: expected '{want_conf}', found '{got_conf}'")

                got_rem = platform_texts(rem_parts.get(p, []))
                want_rem = src.remarks_text if src is not None else [MISSING_MARKER]
                if got_rem != want_rem:
                    diff = next((i for i, (a, b) in enumerate(zip(got_rem, want_rem)) if a != b),
                                min(len(got_rem), len(want_rem)))
                    remarks_check.fail(
                        f"{name} / {p.label}: paragraph {diff + 1} differs "
                        f"(expected {len(want_rem)} paragraphs, found {len(got_rem)})")
                if src is not None:
                    want_lists = sum(1 for b in cell_blocks(src.remarks_cell)
                                     if is_list_para(b) and norm_space(el_text(b)))
                    got_lists = list_count(rem_parts.get(p, []))
                    if want_lists != got_lists:
                        lists_check.fail(f"{name} / {p.label}: expected {want_lists} list items, found {got_lists}")

    for p in platforms:
        for t in sources[p].tables:
            for r in t.rows:
                if r.key not in seen_keys:
                    rows_check.fail(f"{r.title} from {p.label} is missing from the output")
    checks += [rows_check, conf_check, remarks_check, labels_check, lists_check, empty_check]

    scope = Check("Every evaluated page and component from each source is in the scope section")
    out_text = norm_space(" ".join(el_text(t) for t in doc.element.body.iter(qn("w:tbl"))
                                   if not is_criteria_header(t)))
    body_paras = norm_space(" ".join(el_text(p) for p in doc.element.body.iter(qn("w:p"))))
    for p in platforms:
        sec = next((s for s in sources[p].sections if s.key.startswith("scope of evaluation")), None)
        if sec is None:
            continue
        for b in sec.blocks:
            if b.tag == qn("w:tbl"):
                for tr in b.findall(qn("w:tr"))[1:]:
                    for tc in tr.findall(qn("w:tc")):
                        t = norm_space(el_text(tc))
                        if t and t not in out_text:
                            scope.fail(f"{p.label}: '{t[:80]}' not found in output scope table")
            elif is_list_para(b):
                t = norm_space(el_text(b))
                if t and t not in body_paras:
                    scope.fail(f"{p.label}: component '{t}' not found in output")
    checks.append(scope)

    if scope_differences:
        flags = Check("Pages missing from a platform are highlighted and commented in the scope table")
        comments = {c.comment_id: norm_space(c.text) for c in doc.comments}
        for d in scope_differences:
            key = page_key(d["page"])
            rows = [tr for tbl in doc.element.body.iter(qn("w:tbl")) if not is_criteria_header(tbl)
                    for tr in tbl.findall(qn("w:tr"))[1:]
                    if page_key(el_text(tr.findall(qn("w:tc"))[0])) == key]
            if not rows:
                flags.fail(f"'{d['page']}' is not in the scope table")
            for tr in rows:
                runs = [r for r in tr.iter(qn("w:r")) if norm_space(el_text(r))]
                if not runs or any(r.find(f"{qn('w:rPr')}/{qn('w:highlight')}") is None for r in runs):
                    flags.fail(f"'{d['page']}' row is not fully highlighted")
                if any(tc.find(f"{qn('w:tcPr')}/{qn('w:shd')}") is None for tc in tr.findall(qn("w:tc"))):
                    flags.fail(f"'{d['page']}' row cells are not shaded")
                ids = [int(c.get(qn("w:id"))) for c in tr.iter(qn("w:commentRangeStart"))]
                want = scope_comment(d)
                if not any(comments.get(i) == want for i in ids):
                    flags.fail(f"'{d['page']}' row has no comment naming {', '.join(d['missingFrom'])}")
        checks.append(flags)
    return checks
