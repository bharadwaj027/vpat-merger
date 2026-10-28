"""FastAPI web layer for the VPAT merger.

Run: uvicorn app:app --port 8000   (from the backend/ folder)
The built React app in ../frontend/dist is served at "/" when present.
"""
from __future__ import annotations

import json
import queue
import threading
import time
import uuid
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles

from vpat_merger import Platform, inspect_upload, merge_vpats
from vpat_merger.models import MIN_SOURCES
from vpat_merger.service import STAGES

MAX_UPLOAD_BYTES = 25 * 1024 * 1024
RESULT_TTL_SECONDS = 60 * 60
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

app = FastAPI(title="VPAT Merger")
_results: dict[str, tuple[float, bytes, str]] = {}
_lock = threading.Lock()


def _platform(value: str) -> Platform:
    try:
        return Platform(value)
    except ValueError:
        raise HTTPException(400, f"Unknown platform '{value}'. Use one of: {', '.join(p.value for p in Platform)}.")


async def _read_upload(upload: UploadFile) -> tuple[bytes, str]:
    name = Path(upload.filename or "upload.docx").name
    if not name.lower().endswith(".docx"):
        raise HTTPException(400, f"{name} is not a .docx file. Upload a Word (.docx) VPAT.")
    data = await upload.read(MAX_UPLOAD_BYTES + 1)
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(413, f"{name} is larger than 25 MB.")
    if not data:
        raise HTTPException(400, f"{name} is empty.")
    return data, name


def _prune():
    now = time.time()
    with _lock:
        for key in [k for k, (ts, _, _) in _results.items() if now - ts > RESULT_TTL_SECONDS]:
            del _results[key]


@app.get("/api/health")
def health():
    return {"ok": True}


@app.get("/api/stages")
def stages():
    return [{"key": k, "label": label} for k, label in STAGES]


@app.post("/api/inspect")
async def inspect(platform: str = Form(...), file: UploadFile = File(...)):
    p = _platform(platform)
    data, name = await _read_upload(file)
    return {"filename": name, **inspect_upload(data, name, p)}


@app.post("/api/merge")
async def merge(desktop: Optional[UploadFile] = File(None), ios: Optional[UploadFile] = File(None),
                android: Optional[UploadFile] = File(None), ios_native: Optional[UploadFile] = File(None),
                android_native: Optional[UploadFile] = File(None)):
    """Streams NDJSON progress events, ending with a {"type": "result"} event.

    Any two or more of the platform files may be supplied; field names are Platform values.
    """
    uploads = {Platform.DESKTOP: desktop, Platform.IOS: ios, Platform.ANDROID: android,
               Platform.IOS_NATIVE: ios_native, Platform.ANDROID_NATIVE: android_native}
    files = {p: await _read_upload(u) for p, u in uploads.items() if u is not None}
    if len(files) < MIN_SOURCES:
        raise HTTPException(400, f"Upload at least {MIN_SOURCES} VPATs to combine.")
    events: queue.Queue = queue.Queue()

    def work():
        try:
            result = merge_vpats(files, lambda key, status: events.put({"type": "stage", "key": key, "status": status}))
            report = dict(result.report)
            if result.docx is not None:
                _prune()
                token = uuid.uuid4().hex
                with _lock:
                    _results[token] = (time.time(), result.docx, result.filename)
                report["downloadToken"] = token
            events.put({"type": "result", "report": report})
        except Exception as exc:  # surface unexpected failures to the UI instead of a dropped stream
            events.put({"type": "result", "report": {"ok": False, "errors": [f"Unexpected error: {exc}"]}})
        finally:
            events.put(None)

    threading.Thread(target=work, daemon=True).start()

    def stream():
        while True:
            item = events.get()
            if item is None:
                break
            yield json.dumps(item) + "\n"

    return StreamingResponse(stream(), media_type="application/x-ndjson")


@app.get("/api/download/{token}")
def download(token: str):
    with _lock:
        entry = _results.get(token)
    if entry is None:
        raise HTTPException(404, "This download has expired. Generate the combined VPAT again.")
    _, data, filename = entry
    return Response(data, media_type=DOCX_MIME,
                    headers={"Content-Disposition": f'attachment; filename="{filename}"'})


_dist = Path(__file__).resolve().parent.parent / "frontend" / "dist"
if _dist.exists():
    app.mount("/assets", StaticFiles(directory=_dist / "assets"), name="assets")

    @app.get("/")
    def index():
        # Always revalidate: after a rebuild the page must point at the new hashed bundle.
        return FileResponse(_dist / "index.html", headers={"Cache-Control": "no-cache"})
