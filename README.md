# VPAT Merger

Combines any **two or three** of a **Desktop Web**, an **iOS** and an **Android** VPAT (Word, VPAT 2.5) into one
editable Word report. iOS and Android can each be a **mobile web** or a **native app** VPAT. Each criterion gets one row with per-platform conformance levels and remarks, matching the
layout of a Deque combined multi-platform VPAT (kept locally in `fixtures/reference/`, not in the repository).

## Run it

Requires Python 3.11+ and Node 18+.

```bash
python -m venv ~/.venvs/vpat-merger
~/.venvs/vpat-merger/Scripts/python -m pip install -r backend/requirements.txt
```

```bash
cd frontend && npm install && npm run build
```

```bash
~/.venvs/vpat-merger/Scripts/python -m uvicorn app:app --app-dir backend --port 8000
```

Open http://localhost:8000. (For frontend development, run `npm run dev` in `frontend/` while the API runs on
port 8000; Vite proxies `/api`.)

> On Windows, keep the virtualenv in a short path (as above). `lxml` fails to install under long folder paths
> unless Windows long-path support is enabled.

Command line, no UI:

```bash
cd backend && ~/.venvs/vpat-merger/Scripts/python merge_cli.py --desktop D.docx --ios I.docx --android A.docx -o Combined.docx
# native apps: --ios-native IN.docx --android-native AN.docx --report report.json
```

Tests:

```bash
cd backend && ~/.venvs/vpat-merger/Scripts/python -m pytest -q
```

Most tests use sample VPATs in `fixtures/` (`web.docx`, `ios.docx`, `android.docx`, `native/*.docx`). When those
files are absent the tests that need them are skipped.

## How it works

| Layer | File | Responsibility |
|---|---|---|
| Upload & API | `backend/app.py` | `.docx` checks, size limit, per-slot inspection, NDJSON progress stream, download |
| Parsing | `vpat_merger/parser.py` | Splits each document into heading sections; reads conformance tables row by row from the Word XML |
| Matching | `vpat_merger/matching.py` | Aligns criteria by WCAG id (e.g. `1.4.10`), never by position or wording; flags gaps and bad values |
| Generation | `vpat_merger/generator.py` | Uses the first uploaded document (Desktop when present) as the base package and deep-copies the others into it |
| Word helpers | `vpat_merger/docx_utils.py` | Run/paragraph cloning, hyperlink/image relationship and list-numbering remapping |
| Output validation | `vpat_merger/validator.py` | Re-opens the output and compares every value and remarks paragraph with its source |

### Merge rules

- **Platform labels**: `Desktop Web`, `iOS Mobile Web`, `Android Mobile Web`, `Native iOS`, `Native Android`.
  In the UI, the iOS and Android slots have a *Mobile Web / Native app* switch, pre-set from the file name and
  evaluation methods (e.g. "Native Mobile iOS" → Native iOS); you can override it.
- **Two or three platforms**: only the uploaded platforms get labels; a platform that was not uploaded is
  never shown or reported as missing. The title reads *Desktop and Mobile Web*, or *Mobile Web* for iOS + Android.
- **Conformance Level**: `Desktop Web: …`, `iOS Mobile Web: …`, `Android Mobile Web: …` (bold label, original value,
  blank paragraph between platforms). No overall level is computed and values are never normalised.
- **Remarks**: each platform's original paragraphs, bullets, bold text and links are copied as Word XML under its label.
- **Missing data**: a criterion absent from a source shows *Not provided in source VPAT* for that platform (never
  "Not Applicable") and is listed in the summary. Blank or non-standard conformance values are kept as written and
  flagged as errors. Empty remarks are left empty and flagged.
- **Scope differences**: a page listed in one platform's Scope of Evaluation but not another's (e.g. on Desktop
  Web but not iOS Mobile Web) is flagged as a warning and listed in the summary. Pages are matched by title, so the
  same page with a different URL on mobile is not flagged. In the Word document the page's Scope of Evaluation
  row is highlighted yellow and carries a Word comment (author "VPAT Merger") naming the missing platforms.
- **Duplicate rows** in one source: the first is merged and the duplicate is reported as an error.
- **Other sections**: identical sections are kept once. Sections that differ (for example *Name of Product*) are
  repeated per platform under labels. Evaluation methods keep shared sentences once and list each platform's
  browser/OS/assistive technology separately. The scope table keeps its two columns (Page Title, URL) and lists
  every distinct page from all uploaded sources once.

### Limitations

- Assumes the Deque VPAT 2.5 template layout: Heading 1/2/3 sections and 3-column conformance tables.
  Other templates parse only if their conformance tables use the *Criteria / Conformance Level / Remarks* header.
- Comments, bookmarks and embedded objects other than images inside copied iOS/Android content are dropped.
- Header, footer and page setup come from the first uploaded document (Desktop Web when present).
- Results are kept in server memory for one hour; they are lost on restart.
