"""Package the Python merge engine for the in-browser (Pyodide) build.

Writes to frontend/public/py/:
  vpat_merger.zip   the backend/vpat_merger package (unpacked into Pyodide's file system at startup)
  python_docx-*.whl the pure-Python python-docx wheel (installed with micropip, no PyPI access needed)
  manifest.json     file names and versions the web worker loads

Run before `npm run build` (the frontend's `build` script does this automatically).
"""
import json
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PACKAGE = ROOT / "backend" / "vpat_merger"
OUT = ROOT / "frontend" / "public" / "py"
PYTHON_DOCX = "python-docx==1.2.0"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for old in OUT.glob("*"):
        old.unlink()

    zip_path = OUT / "vpat_merger.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(PACKAGE.glob("*.py")):
            z.write(f, f"vpat_merger/{f.name}")

    subprocess.run(
        [sys.executable, "-m", "pip", "download", PYTHON_DOCX, "--no-deps", "--only-binary", ":all:",
         "--dest", str(OUT), "--quiet", "--disable-pip-version-check"],
        check=True,
    )
    wheel = next(OUT.glob("python_docx-*.whl"))

    manifest = {"package": zip_path.name, "wheels": [wheel.name]}
    (OUT / "manifest.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    print(f"Wrote {zip_path.relative_to(ROOT)} ({zip_path.stat().st_size // 1024} KB), {wheel.name}")


if __name__ == "__main__":
    main()
