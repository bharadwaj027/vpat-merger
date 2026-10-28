"""Compare the evaluated pages/screens (Scope of Evaluation) across the uploaded VPATs.

A page found on one platform but not another (e.g. on Desktop Web but not iOS Mobile Web) is flagged. Pages are
matched by title, not URL: the same page often has a different URL on desktop and mobile.
"""
from __future__ import annotations

import re

from docx.oxml.ns import qn

from .docx_utils import el_text, norm_space
from .models import Issue, Platform, SourceVpat, active_platforms


def page_key(title: str) -> str:
    """'Project-wide' and 'Project wide' are the same page; case and spacing are ignored."""
    return re.sub(r"[\s\-_–—]+", " ", norm_space(title).lower()).strip()


def scope_pages(src: SourceVpat) -> dict[str, dict]:
    """Page key → {title, urls} from the first table in the source's Scope of Evaluation section."""
    sec = next((s for s in src.sections if s.key.startswith("scope of evaluation")), None)
    table = next((b for b in sec.blocks if b.tag == qn("w:tbl")), None) if sec else None
    pages: dict[str, dict] = {}
    if table is None:
        return pages
    for tr in table.findall(qn("w:tr"))[1:]:
        cells = [norm_space(el_text(c)) for c in tr.findall(qn("w:tc"))]
        if not cells or not cells[0]:
            continue
        entry = pages.setdefault(page_key(cells[0]), {"title": cells[0], "urls": []})
        entry["urls"].extend(c for c in cells[1:] if c and c not in entry["urls"])
    return pages


def compare_scope(sources: dict[Platform, SourceVpat], issues: list[Issue]) -> list[dict]:
    """One entry per page missing from at least one uploaded platform; also appends a warning per page."""
    platforms = active_platforms(sources)
    pages = {p: scope_pages(sources[p]) for p in platforms}
    if sum(1 for p in platforms if pages[p]) < 2:
        return []  # nothing to compare (a source has no scope table)

    order: list[str] = []
    for p in platforms:
        order.extend(k for k in pages[p] if k not in order)
    differences = []
    for key in sorted(order, key=lambda k: (k, order.index(k))):
        present = [p for p in platforms if key in pages[p]]
        missing = [p for p in platforms if key not in pages[p] and pages[p]]
        if not missing:
            continue
        title = pages[present[0]][key]["title"]
        differences.append({
            "page": title,
            "presentIn": [p.label for p in present],
            "missingFrom": [p.label for p in missing],
            "urls": {p.label: pages[p][key]["urls"] for p in present},
        })
        issues.append(Issue(
            "warning", "scope-page-missing",
            f"Page '{title}' is evaluated in {_join(present)} but is not listed in the {_join(missing)} "
            f"scope of evaluation.",
            None, title,
        ))
    return differences


def _join(platforms: list) -> str:
    labels = [p.label if isinstance(p, Platform) else p for p in platforms]  # Platform is a str enum: test it first
    return labels[0] if len(labels) == 1 else ", ".join(labels[:-1]) + " and " + labels[-1]


def scope_comment(diff: dict) -> str:
    """Text of the Word comment placed on a highlighted scope row."""
    missing, present = diff["missingFrom"], diff["presentIn"]
    return (f"This page is not present in the {_join(missing)} "
            f"{'VPAT' if len(missing) == 1 else 'VPATs'}. It was evaluated on {_join(present)} only.")
