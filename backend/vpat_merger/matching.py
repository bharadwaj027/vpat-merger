"""Align sections and success criteria across the uploaded platform documents."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional, Sequence, TypeVar

from .docx_utils import norm_space
from .models import CriteriaTable, CriterionRow, Issue, Platform, SourceVpat, active_platforms
from .parser import normalize_conformance

T = TypeVar("T")


def union_order(sequences: Sequence[Sequence[str]]) -> list[str]:
    """Merge ordered key lists, keeping each source's relative order.

    Keys missing from the first list are inserted right after their nearest
    preceding key in the list they came from.
    """
    result: list[str] = []
    for seq in sequences:
        for i, key in enumerate(seq):
            if key in result:
                continue
            anchor = next((seq[j] for j in range(i - 1, -1, -1) if seq[j] in result), None)
            pos = result.index(anchor) + 1 if anchor is not None else 0
            # Keys from earlier sources that this source lacks stay first, unless WCAG numbering says otherwise.
            while pos < len(result) and result[pos] not in seq and not _sc_less(key, result[pos]):
                pos += 1
            result.insert(pos, key)
    return result


def _sc_less(a: str, b: str) -> bool:
    """WCAG numeric ordering for criterion keys (falls back to 'not less')."""
    try:
        return tuple(int(x) for x in a.split(".")) < tuple(int(x) for x in b.split("."))
    except ValueError:
        return False


@dataclass
class MergedCriterion:
    key: str
    rows: dict[Platform, Optional[CriterionRow]]
    table_key: str

    @property
    def present(self) -> list[Platform]:
        return [p for p, r in self.rows.items() if r is not None]

    @property
    def missing(self) -> list[Platform]:
        return [p for p, r in self.rows.items() if r is None]

    @property
    def display_row(self) -> CriterionRow:
        return next(r for r in self.rows.values() if r is not None)


@dataclass
class MergedTable:
    key: str
    tables: dict[Platform, Optional[CriteriaTable]]
    criteria: list[MergedCriterion] = field(default_factory=list)


def match_criteria(sources: dict[Platform, SourceVpat], issues: list[Issue]) -> list[MergedTable]:
    platforms = active_platforms(sources)
    table_keys = union_order([[t.key for t in sources[p].tables] for p in platforms])

    # A criterion belongs to the first table it was found in; flag ones that moved between tables.
    home: dict[str, str] = {}
    for p in platforms:
        for t in sources[p].tables:
            for r in t.rows:
                if r.key in home and home[r.key] != t.key:
                    issues.append(Issue("warning", "table-mismatch",
                                        f"{r.title} is in '{t.heading}' for {p.label} but in a different table in "
                                        f"another source. It is merged into the first table it appeared in.", p, r.key))
                else:
                    home.setdefault(r.key, t.key)

    merged_tables = []
    for tk in table_keys:
        tables = {p: next((t for t in sources[p].tables if t.key == tk), None) for p in platforms}
        for p, t in tables.items():
            if t is None:
                heading = next(t.heading for t in tables.values() if t is not None)
                issues.append(Issue("warning", "missing-table",
                                    f"'{heading}' is not present in the {p.label} VPAT; its rows show "
                                    f"'Not provided in source VPAT' for {p.label}.", p))
        reports = {norm_space(t.report_heading or "") for t in tables.values() if t is not None}
        if len(reports) > 1:
            issues.append(Issue("warning", "standard-mismatch",
                                f"The table '{tk}' sits under different report headings across sources: "
                                f"{', '.join(sorted(reports))}. Check that the same WCAG version was used."))

        def rows_in(t: Optional[CriteriaTable]) -> list[CriterionRow]:
            return [r for r in (t.rows if t else []) if home.get(r.key) == tk]

        keys = union_order([[r.key for r in rows_in(tables[p])] for p in platforms])
        if not keys:
            continue
        # Rows that moved between tables are pulled into their home table.
        lookup: dict[Platform, dict[str, CriterionRow]] = {}
        for p in platforms:
            lookup[p] = {}
            for t in sources[p].tables:
                for r in t.rows:
                    lookup[p].setdefault(r.key, r)
        mt = MergedTable(key=tk, tables=tables)
        for k in keys:
            mt.criteria.append(MergedCriterion(key=k, rows={p: lookup[p].get(k) for p in platforms}, table_key=tk))
        merged_tables.append(mt)
    _check_rows(merged_tables, issues)
    return merged_tables


def _check_rows(tables: list[MergedTable], issues: list[Issue]):
    for mt in tables:
        for mc in mt.criteria:
            name = mc.display_row.title
            for p in mc.missing:
                issues.append(Issue("warning", "missing-criterion",
                                    f"{name} is not in the {p.label} VPAT. The merged row shows "
                                    f"'Not provided in source VPAT' for {p.label} (this is not 'Not Applicable').",
                                    p, mc.key))
            titles = {norm_space(r.title) for r in mc.rows.values() if r is not None}
            if len(titles) > 1:
                issues.append(Issue("warning", "title-mismatch",
                                    f"Criterion {mc.key} has different titles across sources: "
                                    f"{' / '.join(sorted(titles))}. The {mc.present[0].label} title is used.",
                                    None, mc.key))
            for p in mc.present:
                r = mc.rows[p]
                if not r.conformance:
                    issues.append(Issue("error", "missing-conformance",
                                        f"{name}: the {p.label} conformance level is blank in the source.", p, mc.key))
                else:
                    canon = normalize_conformance(r.conformance)
                    if canon is None:
                        issues.append(Issue("error", "invalid-conformance",
                                            f"{name}: '{r.conformance}' for {p.label} is not a standard VPAT "
                                            f"conformance term. It was kept as written; please review.", p, mc.key))
                    elif canon != r.conformance:
                        issues.append(Issue("warning", "conformance-format",
                                            f"{name}: '{r.conformance}' for {p.label} differs in capitalization or "
                                            f"spacing from '{canon}'. It was kept as written.", p, mc.key))
                    elif canon == "Not Evaluated" and r.level and "AAA" not in r.level.upper():
                        issues.append(Issue("warning", "invalid-conformance",
                                            f"{name}: 'Not Evaluated' for {p.label} is only permitted for Level AAA "
                                            f"criteria under VPAT 2.5.", p, mc.key))
                if not r.remarks_text:
                    issues.append(Issue("warning", "empty-remarks",
                                        f"{name}: the {p.label} remarks are empty in the source.", p, mc.key))
