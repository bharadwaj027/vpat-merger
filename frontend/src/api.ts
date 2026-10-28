import type { InspectResult, MergeReport, PlatformKey, StageStatus } from "./types";

async function errorText(res: Response): Promise<string> {
  try {
    const body = await res.json();
    return typeof body.detail === "string" ? body.detail : JSON.stringify(body.detail);
  } catch {
    return `${res.status} ${res.statusText}`;
  }
}

export async function inspectFile(platform: PlatformKey, file: File): Promise<InspectResult> {
  const form = new FormData();
  form.append("platform", platform);
  form.append("file", file);
  const res = await fetch("/api/inspect", { method: "POST", body: form });
  if (!res.ok) return { filename: file.name, ok: false, error: await errorText(res) };
  return res.json();
}

export async function fetchStages(): Promise<{ key: string; label: string }[]> {
  const res = await fetch("/api/stages");
  return res.json();
}

/** Posts the uploaded files (two or three platforms) and reads the NDJSON progress stream. */
export async function mergeFiles(
  files: Partial<Record<PlatformKey, File>>,
  onStage: (key: string, status: StageStatus) => void,
): Promise<MergeReport> {
  const form = new FormData();
  for (const [platform, file] of Object.entries(files)) {
    if (file) form.append(platform, file);
  }
  const res = await fetch("/api/merge", { method: "POST", body: form });
  if (!res.ok || !res.body) return { ok: false, errors: [await errorText(res)] };

  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let report: MergeReport | null = null;
  for (;;) {
    const { value, done } = await reader.read();
    if (value) buffer += decoder.decode(value, { stream: true });
    let nl: number;
    while ((nl = buffer.indexOf("\n")) >= 0) {
      const line = buffer.slice(0, nl).trim();
      buffer = buffer.slice(nl + 1);
      if (!line) continue;
      const event = JSON.parse(line);
      if (event.type === "stage") onStage(event.key, event.status);
      else if (event.type === "result") report = event.report;
    }
    if (done) break;
  }
  return report ?? { ok: false, errors: ["The server closed the connection before finishing."] };
}

export function downloadUrl(token: string): string {
  return `/api/download/${token}`;
}
