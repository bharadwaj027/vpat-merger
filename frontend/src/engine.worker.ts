/// <reference lib="webworker" />
// Runs the Python merge engine (backend/vpat_merger) in the browser with Pyodide.
// Documents are processed entirely on the user's machine; nothing is uploaded.

const PYODIDE_VERSION = "0.26.4";
const PYODIDE_URL = `https://cdn.jsdelivr.net/pyodide/v${PYODIDE_VERSION}/full/`;

// Thin JSON bridge between JavaScript and the engine's service layer.
const GLUE = `
import json
from vpat_merger import Platform, inspect_upload, merge_vpats

def js_inspect(platform, data, filename):
    return json.dumps({"filename": filename, **inspect_upload(bytes(data), filename, Platform(platform))})

def js_merge(files, on_stage):
    # \`files\` arrives as a dict (converted by pyodide.toPy); each "data" is a JS Uint8Array proxy.
    sources = {Platform(k): (bytes(v["data"]), v["name"]) for k, v in files.items()}
    result = merge_vpats(sources, lambda key, status: on_stage(key, status))
    return [json.dumps(result.report), result.docx, result.filename]
`;

type EngineRequest =
  | { id: number; type: "init"; baseUrl: string }
  | { id: number; type: "inspect"; platform: string; filename: string; data: ArrayBuffer }
  | { id: number; type: "merge"; files: Record<string, { name: string; data: ArrayBuffer }> };

let pyodide: any = null;
let ready: Promise<void> | null = null;

function post(msg: unknown, transfer: Transferable[] = []) {
  (self as unknown as DedicatedWorkerGlobalScope).postMessage(msg, transfer);
}

async function init(baseUrl: string) {
  const { loadPyodide } = await import(/* @vite-ignore */ `${PYODIDE_URL}pyodide.mjs`);
  pyodide = await loadPyodide({ indexURL: PYODIDE_URL });
  await pyodide.loadPackage(["lxml", "typing-extensions", "micropip"]);

  const assets = new URL("py/", baseUrl);
  const manifest = await (await fetch(new URL("manifest.json", assets))).json();
  const micropip = pyodide.pyimport("micropip");
  for (const wheel of manifest.wheels as string[]) {
    await micropip.install(new URL(wheel, assets).href, { deps: false });
  }
  const pkg = await (await fetch(new URL(manifest.package, assets))).arrayBuffer();
  pyodide.unpackArchive(pkg, "zip", { extractDir: "/home/pyodide" });
  pyodide.runPython(GLUE);
}

self.onmessage = async (event: MessageEvent<EngineRequest>) => {
  const req = event.data;
  try {
    if (req.type === "init") {
      ready ??= init(req.baseUrl);
      await ready;
      post({ id: req.id, ok: true });
      return;
    }
    if (!ready) throw new Error("Engine not initialised");
    await ready;

    if (req.type === "inspect") {
      const fn = pyodide.globals.get("js_inspect");
      const json = fn(req.platform, new Uint8Array(req.data), req.filename);
      fn.destroy();
      post({ id: req.id, ok: true, result: JSON.parse(json) });
    } else if (req.type === "merge") {
      const files: Record<string, { name: string; data: Uint8Array }> = {};
      for (const [k, v] of Object.entries(req.files)) files[k] = { name: v.name, data: new Uint8Array(v.data) };
      const fn = pyodide.globals.get("js_merge");
      const onStage = (key: string, status: string) => post({ id: req.id, stage: { key, status } });
      const pyFiles = pyodide.toPy(files, { depth: 2 });
      const out = fn(pyFiles, onStage);
      pyFiles.destroy();
      fn.destroy();
      const [reportJson, docxProxy, filename] = out.toJs();
      out.destroy();
      let docx: Uint8Array | null = null;
      if (docxProxy != null) {
        docx = docxProxy instanceof Uint8Array ? docxProxy.slice() : docxProxy.toJs().slice();
        if (typeof docxProxy.destroy === "function") docxProxy.destroy();
      }
      post({ id: req.id, ok: true, result: { report: JSON.parse(reportJson), docx, filename } }, docx ? [docx.buffer] : []);
    }
  } catch (e) {
    post({ id: req.id, ok: false, error: e instanceof Error ? e.message : String(e) });
  }
};
