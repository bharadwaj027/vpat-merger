import { useEffect, useRef, useState } from "react";
import { inspectFile, mergeFiles, prepareEngine, STAGES } from "./api";
import { Summary } from "./Summary";
import type { Kind, MergeReport, PlatformKey, SlotKey, StageStatus } from "./types";
import { MIN_SOURCES, PLATFORM_LABELS, SLOTS, slotPlatform } from "./types";
import { SlotState, UploadSlot } from "./UploadSlot";

const EMPTY: SlotState = { file: null, checking: false, result: null, kind: "web", kindTouched: false };

export default function App() {
  const [slots, setSlots] = useState<Record<SlotKey, SlotState>>({
    desktop: EMPTY,
    ios: EMPTY,
    android: EMPTY,
  });
  const stages = STAGES;
  // "loading" | "ready" | "error:<message>" — the in-browser merge engine starts downloading on page load.
  const [engine, setEngine] = useState("loading");
  const [progress, setProgress] = useState<Record<string, StageStatus>>({});
  const [running, setRunning] = useState(false);
  const [report, setReport] = useState<MergeReport | null>(null);
  const resultHeading = useRef<HTMLHeadingElement>(null);

  useEffect(() => {
    prepareEngine().then(
      () => setEngine("ready"),
      (e) => setEngine(`error:${(e as Error).message}`),
    );
  }, []);

  useEffect(() => {
    if (report) resultHeading.current?.focus();
  }, [report]);

  // Any two or three slots can be combined; a slot holding an invalid file blocks generation.
  const labelOf = (key: SlotKey) => PLATFORM_LABELS[slotPlatform(key, slots[key].kind)];
  const validSlots = SLOTS.filter((p) => slots[p.key].result?.ok && !slots[p.key].checking);
  const checking = SLOTS.some((p) => slots[p.key].checking);
  const invalid = SLOTS.filter((p) => slots[p.key].file && !slots[p.key].checking && !slots[p.key].result?.ok);
  const ready = validSlots.length >= MIN_SOURCES && !checking && invalid.length === 0;
  const helpText = checking
    ? "Checking documents…"
    : invalid.length
      ? `Fix or remove the ${invalid.map((p) => labelOf(p.key)).join(", ")} file to continue.`
      : validSlots.length < MIN_SOURCES
        ? `Upload at least ${MIN_SOURCES} VPATs (${validSlots.length} ready).`
        : `Ready to combine ${validSlots.map((p) => labelOf(p.key)).join(" + ")}.`;

  async function select(slot: SlotKey, file: File) {
    setReport(null);
    setProgress({});
    const { kind, kindTouched } = slots[slot];
    setSlots((s) => ({ ...s, [slot]: { ...s[slot], file, checking: true, result: null } }));
    let result;
    if (!file.name.toLowerCase().endsWith(".docx")) {
      result = { filename: file.name, ok: false, error: `${file.name} is not a .docx file. Upload a Word (.docx) VPAT.` };
    } else {
      try {
        result = await inspectFile(slotPlatform(slot, kind), file);
      } catch (e) {
        result = { filename: file.name, ok: false, error: `Could not reach the server: ${String(e)}` };
      }
    }
    // Pre-select Mobile Web / Native from the detected type unless the user already chose.
    const guessed = result.ok ? result.guessedPlatform : null;
    const detected: Kind | null = guessed && guessed.split("_")[0] === slot ? (guessed.endsWith("_native") ? "native" : "web") : null;
    setSlots((s) =>
      s[slot].file === file
        ? { ...s, [slot]: { ...s[slot], checking: false, result, kind: !kindTouched && detected ? detected : s[slot].kind } }
        : s,
    );
  }

  function remove(slot: SlotKey) {
    setReport(null);
    setProgress({});
    setSlots((s) => ({ ...s, [slot]: { ...EMPTY, kind: s[slot].kind, kindTouched: s[slot].kindTouched } }));
  }

  function changeKind(slot: SlotKey, kind: Kind) {
    setReport(null);
    setProgress({});
    setSlots((s) => ({ ...s, [slot]: { ...s[slot], kind, kindTouched: true } }));
  }

  async function generate() {
    if (!ready) return;
    setRunning(true);
    setReport(null);
    setProgress({});
    try {
      const files: Partial<Record<PlatformKey, File>> = {};
      for (const p of validSlots) files[slotPlatform(p.key, slots[p.key].kind)] = slots[p.key].file!;
      const r = await mergeFiles(files, (key, status) => setProgress((p) => ({ ...p, [key]: status })));
      setReport(r);
    } catch (e) {
      setReport({ ok: false, errors: [`Could not reach the server: ${String(e)}`] });
    } finally {
      setRunning(false);
    }
  }

  const started = Object.keys(progress).length > 0;

  return (
    <>
      <header className="topbar">
        <div className="wrap">
          <h1>VPAT Merger</h1>
          <p>Combine Desktop Web, mobile web and native app VPATs into one Word report.</p>
        </div>
      </header>

      <main className="wrap">
        <p className={`engine engine--${engine.split(":")[0]}`} role="status">
          {engine === "loading" &&
            "Loading the merge engine… The first visit downloads about 15 MB; later visits load from your browser cache."}
          {engine === "ready" && "Merge engine ready. Your documents are processed in this browser and never uploaded."}
          {engine.startsWith("error:") &&
            `The merge engine could not load (${engine.slice(6)}). Check your internet connection and reload the page.`}
        </p>
        <section aria-labelledby="upload-h" className="panel">
          <h2 id="upload-h">1. Upload two or three VPATs</h2>
          <p className="muted">
            Add any two or all three platforms. For iOS and Android, choose Mobile Web or Native app; this sets the
            prefix used in the report (for example “Native iOS:”). Each file must be a Word (.docx) VPAT 2.5 report.
          </p>
          <div className="slots">
            {SLOTS.map((p) => (
              <UploadSlot
                key={p.key}
                slot={p.key}
                name={p.name}
                hint={p.hint}
                nativeCapable={p.nativeCapable}
                state={slots[p.key]}
                disabled={running}
                onSelect={(f) => select(p.key, f)}
                onRemove={() => remove(p.key)}
                onKindChange={(k) => changeKind(p.key, k)}
              />
            ))}
          </div>
        </section>

        <section aria-labelledby="gen-h" className="panel">
          <h2 id="gen-h">2. Generate</h2>
          <button
            type="button"
            className="btn btn--primary"
            onClick={generate}
            disabled={!ready || running}
            aria-describedby="gen-help"
          >
            {running ? "Generating…" : "Generate Combined VPAT"}
          </button>
          <p id="gen-help" className="muted">
            {helpText}
          </p>
          {started && (
            <ol className="progress" aria-label="Progress" aria-live="polite">
              {stages.map((s) => {
                const st = progress[s.key] ?? "pending";
                return (
                  <li key={s.key} className={`progress__item progress__item--${st}`}>
                    <span className="progress__icon" aria-hidden="true">
                      {st === "done" ? "✓" : st === "running" ? "…" : st === "error" ? "!" : "○"}
                    </span>
                    {s.label}
                    <span className="visually-hidden">
                      {" "}
                      ({st === "done" ? "complete" : st === "running" ? "in progress" : st === "error" ? "problem" : "pending"})
                    </span>
                  </li>
                );
              })}
            </ol>
          )}
        </section>

        {report && (
          <section aria-labelledby="result-h" className="panel">
            <h2 id="result-h" ref={resultHeading} tabIndex={-1}>
              3. Review and download
            </h2>
            {!report.ok ? (
              <div className="badge badge--error" role="alert">
                <p>The combined VPAT could not be generated:</p>
                <ul>
                  {report.errors?.map((e, i) => (
                    <li key={i}>{e}</li>
                  ))}
                </ul>
              </div>
            ) : (
              <>
                <div className="download">
                  <a className="btn btn--primary" href={report.downloadUrl} download={report.filename}>
                    Download {report.filename}
                  </a>
                  <button type="button" className="btn btn--secondary" onClick={generate} disabled={running}>
                    Regenerate
                  </button>
                </div>
                {!report.outputValid && (
                  <p className="badge badge--error" role="alert">
                    Some output checks failed. The document is available, but review the failed checks below first.
                  </p>
                )}
                <Summary report={report} />
              </>
            )}
          </section>
        )}
      </main>
    </>
  );
}
