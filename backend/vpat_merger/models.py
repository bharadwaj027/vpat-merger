"""Data model shared by the parsing, matching, generation and validation layers."""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Optional


class Platform(str, Enum):
    DESKTOP = "desktop"
    IOS = "ios"
    ANDROID = "android"
    IOS_NATIVE = "ios_native"
    ANDROID_NATIVE = "android_native"

    @property
    def label(self) -> str:
        return PLATFORM_LABELS[self]

    @property
    def os(self) -> str:
        """Operating-system family: 'desktop', 'ios' or 'android'."""
        return self.value.removesuffix("_native")

    @property
    def is_native(self) -> bool:
        return self.value.endswith("_native")


# These labels prefix each platform's entry in the Conformance Level and Remarks columns.
PLATFORM_LABELS = {
    Platform.DESKTOP: "Desktop Web",
    Platform.IOS: "iOS Mobile Web",
    Platform.ANDROID: "Android Mobile Web",
    Platform.IOS_NATIVE: "Native iOS",
    Platform.ANDROID_NATIVE: "Native Android",
}

PLATFORM_ORDER = [Platform.DESKTOP, Platform.IOS, Platform.ANDROID, Platform.IOS_NATIVE, Platform.ANDROID_NATIVE]
MIN_SOURCES = 2


def active_platforms(sources) -> list[Platform]:
    """The uploaded platforms, in canonical order (web first, then native apps)."""
    return [p for p in PLATFORM_ORDER if p in sources]

VALID_CONFORMANCE = [
    "Supports",
    "Partially Supports",
    "Does Not Support",
    "Not Applicable",
    "Not Evaluated",
]

MISSING_MARKER = "Not provided in source VPAT"


@dataclass
class Issue:
    severity: str  # "error" | "warning" | "info"
    category: str
    message: str
    platform: Optional[Platform] = None
    criterion: Optional[str] = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "severity": self.severity,
            "category": self.category,
            "message": self.message,
            "platform": self.platform.value if self.platform else None,
            "platformLabel": self.platform.label if self.platform else None,
            "criterion": self.criterion,
        }


@dataclass
class CriterionRow:
    """One success-criterion row from a source conformance table."""

    key: str  # normalized identifier, e.g. "1.1.1" (or a text key for malformed rows)
    sc_id: Optional[str]  # "1.1.1" when parseable
    title: str  # full text of the criteria cell
    level: Optional[str]
    conformance: str  # exact text of the conformance cell
    tr: Any  # lxml <w:tr> element
    criteria_cell: Any
    conformance_cell: Any
    remarks_cell: Any
    remarks_text: list[str] = field(default_factory=list)
    malformed: bool = False


@dataclass
class CriteriaTable:
    key: str  # normalized heading of the table section, e.g. "table 1: success criteria, level a"
    heading: str
    report_heading: Optional[str]  # e.g. "WCAG 2.2 Report"
    tbl: Any  # lxml <w:tbl>
    rows: list[CriterionRow] = field(default_factory=list)
    duplicates: list[CriterionRow] = field(default_factory=list)


@dataclass
class Section:
    key: str  # normalized heading text ("__title__" for the Heading 1 block)
    heading: str
    level: int  # 1, 2, 3 (0 for content before any heading)
    heading_el: Any
    blocks: list[Any] = field(default_factory=list)  # body elements after the heading
    criteria_table: Optional[CriteriaTable] = None


@dataclass
class SourceVpat:
    platform: Platform
    filename: str
    document: Any  # python-docx Document
    product_name: str
    sections: list[Section]
    tables: list[CriteriaTable]
    issues: list[Issue] = field(default_factory=list)

    def section(self, key: str) -> Optional[Section]:
        return next((s for s in self.sections if s.key == key), None)

    @property
    def criteria_count(self) -> int:
        return sum(len(t.rows) for t in self.tables)
