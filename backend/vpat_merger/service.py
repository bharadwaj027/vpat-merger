"""Orchestrates parse → match → generate → validate, independent of the web layer."""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Callable, Optional

from .generator import Generator, combined_product_name
from .matching import match_criteria
from .models import MIN_SOURCES, Issue, Platform, SourceVpat, active_platforms
from .parser import VpatParseError, guess_platform, parse_vpat
from .scope import compare_scope
from .validator import validate_output

STAGES = [
    ("read", "Reading Word documents"),
    ("extract", "Extracting WCAG criteria"),
    ("match", "Matching criteria across platforms"),
    ("merge", "Merging conformance results and remarks"),
    ("generate", "Generating the combined Word document"),
    ("validate", "Validating the output"),
]


@dataclass
class MergeResult:
    docx: Optional[bytes]
    filename: str
    report: dict


def inspect_upload(data: bytes, filename: str, platform: Platform) -> dict:
    """Basic per-file validation used when a file is dropped into an upload slot."""
    try:
        src = parse_vpat(data, filename, platform)
    except VpatParseError as exc:
        return {"ok": False, "error": str(exc)}
    guessed = guess_platform(src)
    warnings = [i.message for i in src.issues if i.severity != "info"]
    # Only a wrong OS is a slot mistake; native vs mobile web is chosen in the slot (and pre-set from `guessed`).
    if guessed is not None and guessed.os != platform.os:
        warnings.insert(0, f"This file looks like a {guessed.label} VPAT (based on its evaluation methods), "
                           f"but it was added as {platform.label}. Check that it is in the right slot.")
    return {
        "ok": True,
        "productName": src.product_name,
        "criteriaCount": src.criteria_count,
        "tables": [{"heading": t.heading, "count": len(t.rows)} for t in src.tables],
        "guessedPlatform": guessed.value if guessed else None,
        "warnings": warnings,
    }


def safe_filename(name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._ -]+", "", name).strip().replace(" ", "-")
    return cleaned or "Combined"


def merge_vpats(files: dict[Platform, tuple[bytes, str]],
                on_stage: Optional[Callable[[str, str], None]] = None) -> MergeResult:
    stage = on_stage or (lambda _k, _s: None)
    issues: list[Issue] = []

    stage("read", "running")
    sources: dict[Platform, SourceVpat] = {}
    errors = []
    if len(files) < MIN_SOURCES:
        stage("read", "error")
        return MergeResult(None, "", {"ok": False, "errors": [
            f"Upload at least {MIN_SOURCES} VPATs to combine."]})
    for p in active_platforms(files):
        data, name = files[p]
        try:
            sources[p] = parse_vpat(data, name, p)
        except VpatParseError as exc:
            errors.append(f"{p.label}: {exc}")
    if errors:
        stage("read", "error")
        return MergeResult(None, "", {"ok": False, "errors": errors})
    stage("read", "done")

    stage("extract", "running")
    for p, src in sources.items():
        issues.extend(src.issues)
        guessed = guess_platform(src)
        if guessed is not None and guessed != p:
            issues.append(Issue("warning", "platform-slot",
                                f"{src.filename} looks like a {guessed.label} VPAT but was uploaded as {p.label}.", p))
    stage("extract", "done")

    stage("match", "running")
    merged_tables = match_criteria(sources, issues)
    scope_differences = compare_scope(sources, issues)
    stage("match", "done")

    stage("merge", "running")
    generator = Generator(sources, merged_tables, issues, scope_differences)
    stage("merge", "done")

    stage("generate", "running")
    try:
        docx_bytes = generator.build()
    except Exception as exc:
        stage("generate", "error")
        return MergeResult(None, "", {"ok": False, "errors": [f"Document generation failed: {exc}"]})
    stage("generate", "done")

    stage("validate", "running")
    checks = validate_output(docx_bytes, sources, merged_tables, scope_differences)
    stage("validate", "done" if all(c.passed for c in checks) else "error")

    product = combined_product_name(sources)
    filename = f"{safe_filename(product)}-Combined-VPAT.docx"
    return MergeResult(docx_bytes, filename, build_report(sources, merged_tables, issues, checks, product, filename,
                                                         scope_differences))


def build_report(sources, merged_tables, issues, checks, product, filename, scope_differences=()) -> dict:
    criteria = [mc for mt in merged_tables for mc in mt.criteria]
    in_all = [mc for mc in criteria if not mc.missing]
    partial = [mc for mc in criteria if mc.missing]
    platforms = active_platforms(sources)
    severity_rank = {"error": 0, "warning": 1, "info": 2}
    issues = sorted(issues, key=lambda i: severity_rank.get(i.severity, 3))
    return {
        "ok": True,
        "productName": product,
        "filename": filename,
        "platforms": [p.value for p in platforms],
        "platformLabels": {p.value: p.label for p in platforms},
        "sources": [
            {
                "platform": p.value,
                "platformLabel": p.label,
                "filename": sources[p].filename,
                "productName": sources[p].product_name,
                "criteriaCount": sources[p].criteria_count,
                "tables": [{"heading": t.heading, "count": len(t.rows)} for t in sources[p].tables],
            }
            for p in platforms
        ],
        "combined": {
            "uniqueCriteria": len(criteria),
            "inAllSources": len(in_all),
            "missingFromSome": [
                {"key": mc.key, "title": mc.display_row.title, "missing": [p.label for p in mc.missing]}
                for mc in partial
            ],
            "tables": [{"key": mt.key, "count": len(mt.criteria)} for mt in merged_tables],
        },
        "scopeDifferences": list(scope_differences),
        "matrix": [
            {
                "key": mc.key,
                "title": mc.display_row.title,
                "table": mt.key,
                "results": {p.value: (mc.rows[p].conformance if mc.rows.get(p) else None) for p in platforms},
            }
            for mt in merged_tables
            for mc in mt.criteria
        ],
        "issues": [i.to_dict() for i in issues],
        "issueCounts": {s: sum(1 for i in issues if i.severity == s) for s in ("error", "warning", "info")},
        "checks": [c.to_dict() for c in checks],
        "outputValid": all(c.passed for c in checks),
    }
