import type { InspectResult, MergeReport, PlatformKey, StageStatus } from "./types";

// The merge engine runs in a Web Worker (Pyodide) — the VPATs never leave the user's browser,
// so the site can be hosted as static files (GitHub Pages).

const worker = new Worker(new URL("./engine.worker.ts", import.meta.url), { type: "module" });
let nextId = 1;
const pending = new Map<
  number,
  { resolve: (v: unknown) => void; reject: (e: Error) => void; onStage?: (key: string, status: StageStatus) => void }
>();

worker.onmessage = (event: MessageEvent) => {
  const msg = event.data;
  const entry = pending.get(msg.id);
  if (!entry) return;
  if (msg.stage) {
    entry.onStage?.(msg.stage.key, msg.stage.status);
    return;
  }
  pending.delete(msg.id);
  if (msg.ok) entry.resolve(msg.result);
  else entry.reject(new Error(msg.error));
};

function call<T>(message: object, transfer: Transferable[] = [], onStage?: (key: string, status: StageStatus) => void) {
  const id = nextId++;
  return new Promise<T>((resolve, reject) => {
    pending.set(id, { resolve: resolve as (v: unknown) => void, reject, onStage });
    worker.postMessage({ id, ...message }, transfer);
  });
}

let engineReady: Promise<void> | null = null;

/** Downloads and starts the merge engine (first visit ~15 MB, then cached by the browser). */
export function prepareEngine(): Promise<void> {
  engineReady ??= call<void>({ type: "init", baseUrl: document.baseURI });
  return engineReady;
}

export async function inspectFile(platform: PlatformKey, file: File): Promise<InspectResult> {
  await prepareEngine();
  const data = await file.arrayBuffer();
  try {
    return await call<InspectResult>({ type: "inspect", platform, filename: file.name, data }, [data]);
  } catch (e) {
    return { filename: file.name, ok: false, error: (e as Error).message };
  }
}

export const STAGES = [
  { key: "read", label: "Reading Word documents" },
  { key: "extract", label: "Extracting WCAG criteria" },
  { key: "match", label: "Matching criteria across platforms" },
  { key: "merge", label: "Merging conformance results and remarks" },
  { key: "generate", label: "Generating the combined Word document" },
  { key: "validate", label: "Validating the output" },
];

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
let lastDownload: string | null = null;

/** Merges the uploaded files (two or more platforms), reporting progress per stage. */
export async function mergeFiles(
  files: Partial<Record<PlatformKey, File>>,
  onStage: (key: string, status: StageStatus) => void,
): Promise<MergeReport> {
  await prepareEngine();
  const payload: Record<string, { name: string; data: ArrayBuffer }> = {};
  const transfer: ArrayBuffer[] = [];
  for (const [platform, file] of Object.entries(files)) {
    if (!file) continue;
    const data = await file.arrayBuffer();
    payload[platform] = { name: file.name, data };
    transfer.push(data);
  }
  try {
    const { report, docx, filename } = await call<{ report: MergeReport; docx: Uint8Array | null; filename: string }>(
      { type: "merge", files: payload },
      transfer,
      onStage,
    );
    if (docx) {
      if (lastDownload) URL.revokeObjectURL(lastDownload);
      lastDownload = URL.createObjectURL(new Blob([docx], { type: DOCX_MIME }));
      report.downloadUrl = lastDownload;
      report.filename = filename;
    }
    return report;
  } catch (e) {
    return { ok: false, errors: [`The merge engine failed: ${(e as Error).message}`] };
  }
}
