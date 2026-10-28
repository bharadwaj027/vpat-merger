"""Command-line entry point: merge two or more VPATs without the web UI.

Usage: python merge_cli.py [--desktop D.docx] [--ios I.docx] [--android A.docx]
                           [--ios-native IN.docx] [--android-native AN.docx] [-o OUTPUT.docx]
At least two inputs are required. --ios/--android are mobile web; --ios-native/--android-native are native apps.
"""
import argparse
import json
import sys
from pathlib import Path

from vpat_merger import Platform, merge_vpats


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--desktop", help="Desktop Web VPAT (.docx)")
    ap.add_argument("--ios", help="iOS Mobile Web VPAT (.docx)")
    ap.add_argument("--android", help="Android Mobile Web VPAT (.docx)")
    ap.add_argument("--ios-native", help="Native iOS app VPAT (.docx)")
    ap.add_argument("--android-native", help="Native Android app VPAT (.docx)")
    ap.add_argument("-o", "--output")
    ap.add_argument("--report", help="Write the validation report as JSON to this path")
    args = ap.parse_args()

    files = {}
    for platform in Platform:
        path = getattr(args, platform.value)
        if path:
            p = Path(path)
            files[platform] = (p.read_bytes(), p.name)
    if len(files) < 2:
        ap.error("give at least two input VPATs")

    result = merge_vpats(files, lambda key, status: print(f"[{status:>7}] {key}", file=sys.stderr))
    if result.docx is None:
        print("\n".join(result.report["errors"]), file=sys.stderr)
        return 1
    out = Path(args.output or result.filename)
    out.write_bytes(result.docx)
    rep = result.report
    if args.report:
        Path(args.report).write_text(json.dumps(rep, indent=2), encoding="utf-8")
    print(f"Wrote {out}")
    print(f"Unique criteria: {rep['combined']['uniqueCriteria']}, in all sources: {rep['combined']['inAllSources']}")
    print(f"Issues: {rep['issueCounts']}")
    for c in rep["checks"]:
        print(f"  [{'PASS' if c['passed'] else 'FAIL'}] {c['name']}")
        for d in c["details"]:
            print(f"         - {d}")
    return 0 if rep["outputValid"] else 2


if __name__ == "__main__":
    sys.exit(main())
