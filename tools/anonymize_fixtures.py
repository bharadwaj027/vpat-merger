"""Strip client identifiers from VPAT .docx files so they can be shared as test fixtures.

Rewrites every XML part (body, headers/footers, comments, hyperlink targets in .rels, document properties).
Replacements are deterministic, so the same URL or ID maps to the same value in every file and the
cross-platform matching behaves exactly as with the originals.

Usage: python tools/anonymize_fixtures.py FILE.docx [FILE.docx ...]   (files are rewritten in place)
"""
import hashlib
import json
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

# The real names → neutral names mapping is private, so it lives in a git-ignored JSON file next to this script
# (see anonymize-map.example.json for the format). Names are replaced in all case variants; domains keep their
# subdomain so desktop vs mobile URLs stay distinguishable.
MAP_FILE = Path(__file__).with_name("anonymize-map.local.json")
if not MAP_FILE.exists():
    sys.exit(f"Create {MAP_FILE.name} (copy anonymize-map.example.json) with the names to replace.")
_MAP = json.loads(MAP_FILE.read_text(encoding="utf-8"))
NAMES: dict = _MAP.get("names", {})
DOMAINS: dict = _MAP.get("domains", {})
# Opaque identifiers embedded in URLs (account, session and verification tokens).
ID_PATTERNS = [
    re.compile(r"\bacct_[A-Za-z0-9]{8,}"),
    re.compile(r"\bvs_[A-Za-z0-9]{8,}"),
    re.compile(r"\btest_[A-Za-z0-9]{10,}"),
]


def _fake_id(match: re.Match) -> str:
    original = match.group(0)
    prefix, rest = original.split("_", 1)
    if rest.startswith("TEST"):  # already anonymized: keep it, so re-running is a no-op
        return original
    digest = hashlib.sha256(original.encode()).hexdigest().upper()
    return f"{prefix}_TEST{digest[: max(4, len(rest) - 4)]}"


def _case_variants(src: str, dst: str):
    yield src, dst
    yield src.lower(), dst.lower()
    yield src.upper(), dst.upper()


def anonymize_text(text: str) -> str:
    for src, dst in DOMAINS.items():
        text = re.sub(re.escape(src), dst, text, flags=re.IGNORECASE)
    for pattern in ID_PATTERNS:
        text = pattern.sub(_fake_id, text)
    for src, dst in NAMES.items():
        for a, b in _case_variants(src, dst):
            text = text.replace(a, b)
    return text


def anonymize_docx(path: Path) -> int:
    changed = 0
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / path.name
        with zipfile.ZipFile(path) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.endswith((".xml", ".rels")):
                    text = data.decode("utf-8")
                    new = anonymize_text(text)
                    if new != text:
                        changed += 1
                        data = new.encode("utf-8")
                zout.writestr(item, data)
        shutil.copyfile(out, path)
    return changed


def remaining_identifiers(path: Path) -> list[str]:
    """Any client name, domain or raw ID still present (checks every XML part)."""
    words = [n.lower() for n in NAMES] + [d.lower() for d in DOMAINS]
    found = []
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if name.endswith((".xml", ".rels")):
                text = z.read(name).decode("utf-8", "ignore")
                low = text.lower()
                found += [f"{name}: {w}" for w in words if w in low]
                for pattern in ID_PATTERNS:
                    found += [f"{name}: {m}" for m in pattern.findall(text) if "_TEST" not in m]
    return found


if __name__ == "__main__":
    status = 0
    for arg in sys.argv[1:]:
        p = Path(arg)
        parts = anonymize_docx(p)
        left = remaining_identifiers(p)
        print(f"{p}: {parts} parts rewritten" + (f"; STILL FOUND: {left[:5]}" if left else "; clean"))
        status |= bool(left)
    sys.exit(status)
