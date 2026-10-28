import json

import conftest
from conftest import FILES
from fastapi.testclient import TestClient

from app import app
from vpat_merger import Platform

client = TestClient(app)
DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def upload(platform):
    path = FILES[platform]
    return (path.name, path.read_bytes(), DOCX)


def test_inspect_valid_file():
    res = client.post("/api/inspect", data={"platform": "ios"}, files={"file": upload(Platform.IOS)})
    body = res.json()
    assert res.status_code == 200 and body["ok"]
    assert body["criteriaCount"] == 55
    assert body["guessedPlatform"] == "ios"
    assert not any("looks like" in w for w in body["warnings"])


def test_inspect_warns_about_wrong_slot():
    res = client.post("/api/inspect", data={"platform": "desktop"}, files={"file": upload(Platform.ANDROID)})
    body = res.json()
    assert body["ok"] and "Android Mobile Web" in body["warnings"][0]


def test_inspect_rejects_non_docx():
    res = client.post("/api/inspect", data={"platform": "ios"}, files={"file": ("x.pdf", b"%PDF", "application/pdf")})
    assert res.status_code == 400 and "not a .docx" in res.json()["detail"]


def test_inspect_rejects_corrupt_docx():
    res = client.post("/api/inspect", data={"platform": "ios"}, files={"file": ("x.docx", b"garbage", DOCX)})
    assert res.status_code == 200 and not res.json()["ok"]


def test_merge_stream_and_download():
    res = client.post("/api/merge", files={
        "desktop": upload(Platform.DESKTOP),
        "ios": upload(Platform.IOS),
        "android": upload(Platform.ANDROID),
    })
    assert res.status_code == 200
    events = [json.loads(line) for line in res.text.splitlines() if line.strip()]
    stages = [(e["key"], e["status"]) for e in events if e["type"] == "stage"]
    assert ("validate", "done") in stages
    report = events[-1]["report"]
    assert report["ok"] and report["outputValid"]
    dl = client.get(f"/api/download/{report['downloadToken']}")
    assert dl.status_code == 200
    assert dl.headers["content-type"] == DOCX
    assert "TestCo-IdentityDashboard-Combined-VPAT.docx" in dl.headers["content-disposition"]
    assert dl.content[:2] == b"PK"


def test_download_unknown_token():
    assert client.get("/api/download/nope").status_code == 404


def test_merge_two_platforms_via_api():
    res = client.post("/api/merge", files={"ios": upload(Platform.IOS), "android": upload(Platform.ANDROID)})
    report = [json.loads(line) for line in res.text.splitlines() if line.strip()][-1]["report"]
    assert report["ok"] and report["outputValid"]
    assert report["platforms"] == ["ios", "android"]


def test_merge_rejects_single_file():
    res = client.post("/api/merge", files={"ios": upload(Platform.IOS)})
    assert res.status_code == 400 and "at least 2" in res.json()["detail"]


def test_native_platforms_via_api():
    native = conftest.FIXTURES / "native"
    files = {"ios_native": ("ios-native.docx", (native / "ios-native.docx").read_bytes(), DOCX),
             "android_native": ("android-native.docx", (native / "android-native.docx").read_bytes(), DOCX)}
    res = client.post("/api/merge", files=files)
    report = [json.loads(line) for line in res.text.splitlines() if line.strip()][-1]["report"]
    assert report["ok"] and report["outputValid"]
    assert report["platformLabels"] == {"ios_native": "Native iOS", "android_native": "Native Android"}


def test_inspect_detects_native_without_slot_warning():
    path = conftest.FIXTURES / "native" / "ios-native.docx"
    res = client.post("/api/inspect", data={"platform": "ios"}, files={"file": (path.name, path.read_bytes(), DOCX)})
    body = res.json()
    assert body["guessedPlatform"] == "ios_native"
    assert not any("looks like" in w for w in body["warnings"])
