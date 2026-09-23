import http from "node:http";
import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { readFile, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { extname, join, normalize, resolve, sep } from "node:path";
import assert from "node:assert/strict";

const root = process.cwd();
const dist = resolve(root, "dist-office-windows-native");
const offset = process.pid % 1000;
const port = 7300 + offset;
const debugPort = 14300 + offset;
const origin = `http://127.0.0.1:${port}`;
const profile = join(tmpdir(), `visualtex-p12-office-core-${process.pid}`);
const candidates = [
  process.env.VISUALTEX_CHROME_PATH,
  process.env.PROGRAMFILES && join(process.env.PROGRAMFILES, "Microsoft", "Edge", "Application", "msedge.exe"),
  process.env["PROGRAMFILES(X86)"] && join(process.env["PROGRAMFILES(X86)"], "Microsoft", "Edge", "Application", "msedge.exe"),
  process.env.PROGRAMFILES && join(process.env.PROGRAMFILES, "Google", "Chrome", "Application", "chrome.exe"),
  process.env.LOCALAPPDATA && join(process.env.LOCALAPPDATA, "Google", "Chrome", "Application", "chrome.exe"),
].filter(Boolean);
const browserPath = candidates.find((candidate) => existsSync(candidate));
if (!browserPath) throw new Error("No Chrome/Edge executable found for P12 Office core acceptance.");
if (!existsSync(join(dist, "dialog", "index.html"))) {
  throw new Error("Build the core-enabled Windows Office UI before running P12 acceptance.");
}

const now = Date.now();
function makeSession(id, latex, overrides = {}) {
  return {
    id,
    mode: "edit",
    host: "word",
    formulaId: `${id}-formula`,
    sourceDocumentId: `${id}-document`,
    sourceObjectId: `${id}-object`,
    title: `P12 ${id}`,
    lines: [{ id: `${id}-line`, latex }],
    activeLineId: `${id}-line`,
    codeFormat: "raw",
    displayMode: "inline",
    objectMode: "nativeOle",
    numbered: false,
    mathTypeNumberPosition: "right",
    fontSizePt: 10.5,
    exportWidth: 0,
    exportHeight: 0,
    exportResult: null,
    originalMetadata: null,
    dirty: false,
    status: "created",
    autoCommitOnClose: true,
    explicitCancel: false,
    error: null,
    createdAt: now,
    updatedAt: now,
    expiresAt: now + 60 * 60 * 1000,
    ...overrides,
  };
}

const sessions = new Map([
  ["p12", makeSession("p12", "x+y=1")],
  ["p12b", makeSession("p12b", "q=2")],
  ["p12convert", makeSession("p12convert", "a+b=c", { objectMode: "nativeOle", displayMode: "block" })],
]);
const records = [];
const json = (res, value, statusCode = 200) => {
  const bytes = Buffer.from(JSON.stringify(value));
  res.writeHead(statusCode, {
    "content-type": "application/json; charset=utf-8",
    "content-length": String(bytes.length),
    "cache-control": "no-store",
  });
  res.end(bytes);
};
const mime = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".wasm": "application/wasm",
  ".otf": "font/otf",
  ".woff2": "font/woff2",
};
async function body(req) {
  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  return Buffer.concat(chunks).toString("utf8");
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url ?? "/", origin);
    if (url.pathname === "/api/v1/theme" && req.method === "GET") {
      return json(res, { theme: "light", editorLayout: "standard" });
    }
    if (url.pathname === "/api/v1/preferences" && req.method === "GET") {
      return json(res, {
        powerpointDefaultFontSizePt: 20,
        editorPreferences: {
          settings: {
            language: "cn",
            theme: "light",
            editorLayout: "standard",
            zoom: 0.6,
            formulaLetterFont: "katex",
            formulaChineseFont: "system",
          },
        },
      });
    }
    if (url.pathname === "/api/v1/app/converter/next-batch" && req.method === "GET") {
      return json(res, { sessionIds: [] });
    }
    const sessionMatch = /^\/api\/v1\/sessions\/([^/]+)$/u.exec(url.pathname);
    if (sessionMatch) {
      const id = decodeURIComponent(sessionMatch[1]);
      const current = sessions.get(id);
      if (!current) return json(res, { error: "missing" }, 404);
      if (req.method === "GET") return json(res, current);
      if (req.method === "PATCH") {
        const raw = await body(req);
        const update = raw ? JSON.parse(raw) : {};
        const merged = {
          ...current,
          ...update,
          id: current.id,
          formulaId: current.formulaId,
          createdAt: current.createdAt,
          updatedAt: Date.now(),
        };
        records.push({ at: Date.now(), id, method: "PATCH", update: structuredClone(update), responseStatus: merged.status });
        if (update.status === "committing") {
          sessions.set(id, { ...merged, status: "completed", updatedAt: Date.now() });
          return json(res, merged);
        }
        sessions.set(id, merged);
        return json(res, merged);
      }
    }
    const closeMatch = /^\/api\/v1\/app\/sessions\/([^/]+)\/close$/u.exec(url.pathname);
    if (closeMatch && req.method === "POST") {
      records.push({ at: Date.now(), id: decodeURIComponent(closeMatch[1]), method: "CLOSE" });
      res.writeHead(204); res.end(); return;
    }

    let pathname = url.pathname === "/" ? "/dialog/index.html" : url.pathname;
    pathname = decodeURIComponent(pathname);
    const candidate = normalize(resolve(dist, `.${pathname}`));
    if (!candidate.startsWith(normalize(dist + sep))) {
      res.writeHead(403); res.end("forbidden"); return;
    }
    let info;
    try { info = await stat(candidate); } catch { res.writeHead(404); res.end("not found"); return; }
    if (!info.isFile()) { res.writeHead(404); res.end("not found"); return; }
    const bytes = await readFile(candidate);
    res.writeHead(200, {
      "content-type": mime[extname(candidate)] ?? "application/octet-stream",
      "content-length": String(bytes.length),
      "cache-control": "no-store",
    });
    res.end(bytes);
  } catch (error) {
    res.writeHead(500, { "content-type": "text/plain; charset=utf-8" });
    res.end(error instanceof Error ? error.stack ?? error.message : String(error));
  }
});

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
async function waitHttp(url, timeout = 15_000) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    try { const response = await fetch(url); if (response.ok) return response; } catch {}
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${url}`);
}
class CdpClient {
  constructor(url) { this.url = url; this.id = 1; this.pending = new Map(); }
  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((resolve, reject) => {
      this.socket.addEventListener("open", resolve, { once: true });
      this.socket.addEventListener("error", reject, { once: true });
    });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
      const pending = this.pending.get(message.id);
      if (!pending) return;
      this.pending.delete(message.id);
      message.error ? pending.reject(new Error(message.error.message)) : pending.resolve(message.result);
    });
  }
  send(method, params = {}) {
    const id = this.id++;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }
  close() { this.socket?.close(); }
}
async function connectPage(url) {
  await waitHttp(`http://127.0.0.1:${debugPort}/json/list`);
  const start = Date.now();
  while (Date.now() - start < 15_000) {
    const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
    const page = targets.find((target) => target.type === "page" && target.url === url);
    if (page) {
      const client = new CdpClient(page.webSocketDebuggerUrl);
      await client.connect();
      await client.send("Runtime.enable");
      await client.send("Page.enable");
      return client;
    }
    await sleep(80);
  }
  throw new Error(`No browser page for ${url}`);
}
const evalValue = async (client, expression) => {
  const result = await client.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
  if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description ?? result.exceptionDetails.text);
  return result.result.value;
};
async function waitEval(client, expression, label, timeout = 15_000) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    if (await evalValue(client, expression)) return;
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${label}`);
}
async function textKey(client, text) {
  await client.send("Input.insertText", { text });
  await sleep(40);
}
async function click(client, selector) {
  const point = await evalValue(client, `(()=>{const e=document.querySelector(${JSON.stringify(selector)});if(!e)return null;e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2};})()`);
  if (!point) throw new Error(`Missing selector: ${selector}`);
  for (const type of ["mousePressed", "mouseReleased"]) {
    await client.send("Input.dispatchMouseEvent", { type, ...point, button: "left", clickCount: 1 });
  }
  await sleep(80);
}
function patches(id) { return records.filter((record) => record.method === "PATCH" && record.id === id); }
async function waitRecord(predicate, label, timeout = 12_000) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    const found = records.find(predicate);
    if (found) return found;
    await sleep(80);
  }
  throw new Error(`Timed out waiting for record: ${label}`);
}

await new Promise((resolveListen) => server.listen(port, "127.0.0.1", resolveListen));
let browser;
const clients = [];
try {
  const interactiveUrl = `${origin}/dialog/index.html?sessionId=p12&runtime=vsto-desktop`;
  browser = spawn(browserPath, [
    "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
    `--remote-debugging-port=${debugPort}`, `--user-data-dir=${profile}`, "--window-size=1400,1000", interactiveUrl,
  ], { stdio: "ignore" });
  const interactive = await connectPage(interactiveUrl); clients.push(interactive);
  await waitEval(interactive, `Boolean(document.querySelector('[data-engine="visualtex-core"] .vt-core-input'))`, "interactive Office core editor");
  const mounted = await evalValue(interactive, `({core:Boolean(document.querySelector('[data-engine="visualtex-core"]')),legacy:document.querySelectorAll('math-field').length,coreInput:Boolean(document.querySelector('.vt-core-input')),wasm:[...performance.getEntriesByType('resource')].some(x=>String(x.name).includes('visualtex_engine_wasm_bg')),font:[...performance.getEntriesByType('resource')].some(x=>String(x.name).includes('latinmodern-math'))})`);
  assert.deepEqual(mounted, { core: true, legacy: 0, coreInput: true, wasm: true, font: true });

  await waitRecord((r) => r.id === "p12" && r.method === "PATCH" && r.update.status === "editing", "initial editing receipt");
  const initialPatches = patches("p12");
  assert.ok(initialPatches.length >= 1);
  assert.ok(initialPatches.every((r) => r.update.dirty !== true), "initial load must not mark Session dirty");
  assert.ok(initialPatches.every((r) => !r.update.lines || r.update.lines[0]?.latex === "x+y=1"), "initial load must preserve source");

  await evalValue(interactive, `document.querySelector('.vt-core-input').focus();true`);
  await textKey(interactive, "+z");
  const edited = await waitRecord((r) => r.id === "p12" && r.method === "PATCH" && r.update.lines?.[0]?.latex?.endsWith("+z") && r.update.exportResult, "core autosave export");
  assert.equal(edited.update.dirty, true);
  assert.ok(edited.update.exportResult.svg.includes('fill-opacity="0.001"'));
  assert.ok(!edited.update.exportResult.svg.includes("data-mml-node"));
  assert.ok(!edited.update.exportResult.mathMl.includes("data-visualtex-core-grid-id"));
  assert.ok(!edited.update.exportResult.mathMl.includes("mjx"));
  assert.ok(edited.update.exportResult.pngBase64?.length > 20, "nativeOle core autosave must include PNG");

  await click(interactive, "[data-office-undo-action]");
  await waitRecord((r) => r.id === "p12" && r.method === "PATCH" && r.at > edited.at && r.update.lines?.[0]?.latex === "x+y=1", "core Office undo");
  await click(interactive, "[data-office-redo-action]");
  await waitRecord((r) => r.id === "p12" && r.method === "PATCH" && r.at > edited.at && r.update.lines?.[0]?.latex?.endsWith("+z"), "core Office redo");

  // Start one more export then switch Sessions immediately. Any old export that
  // resolves after this point must fail the session/formula/owner/revision guard.
  await evalValue(interactive, `document.querySelector('.vt-core-input').focus();true`);
  await textKey(interactive, "+s");
  const switchAt = Date.now();
  await evalValue(interactive, `window.dispatchEvent(new CustomEvent('visualtex-office-session',{detail:{sessionId:'p12b'}}));true`);
  await waitEval(interactive, `Boolean(document.querySelector('[data-engine="visualtex-core"] .vt-core-input')) && document.querySelectorAll('math-field').length===0`, "second Session core remount");
  await waitRecord((r) => r.id === "p12b" && r.method === "PATCH" && r.update.status === "editing", "second Session editing receipt");
  await sleep(700);
  const staleExports = records.filter((r) => r.id === "p12" && r.method === "PATCH" && r.at >= switchAt && r.update.exportResult);
  assert.equal(staleExports.length, 0, "old Session export must not be written after Session switch");
  assert.equal(sessions.get("p12b").lines[0].latex, "q=2");
  assert.equal(sessions.get("p12b").dirty, false);

  await evalValue(interactive, `document.querySelector('.vt-core-input').focus();true`);
  await textKey(interactive, "+r");
  const p12bExport = await waitRecord((r) => r.id === "p12b" && r.method === "PATCH" && r.update.lines?.[0]?.latex?.endsWith("+r") && r.update.exportResult, "second Session core export");
  assert.ok(p12bExport.update.exportResult.pngBase64?.length > 20);
  await click(interactive, "[data-office-primary-action]");
  const commit = await waitRecord((r) => r.id === "p12b" && r.method === "PATCH" && r.update.status === "committing", "interactive core commit");
  assert.ok(commit.update.exportResult?.svg?.includes('fill-opacity="0.001"'));
  assert.ok(commit.update.exportResult?.pngBase64?.length > 20);
  await waitRecord((r) => r.id === "p12b" && r.method === "CLOSE", "interactive close after completed commit");

  const convertUrl = `${origin}/dialog/index.html?sessionId=p12convert&runtime=vsto-convert`;
  const newTarget = await interactive.send("Target.createTarget", { url: convertUrl });
  assert.ok(newTarget.targetId);
  const converter = await connectPage(convertUrl); clients.push(converter);
  await waitEval(converter, `document.readyState==='complete' || document.readyState==='interactive'`, "converter document");
  await waitEval(converter, `document.querySelectorAll('math-field').length>0 || Boolean(document.querySelector('.office-dialog-state'))`, "converter legacy surface");
  const converterOwners = await evalValue(converter, `({core:Boolean(document.querySelector('[data-engine="visualtex-core"]')),legacy:document.querySelectorAll('math-field').length})`);
  assert.equal(converterOwners.core, false, "vsto-convert must remain outside the P12 core gate");
  assert.ok(converterOwners.legacy > 0, "vsto-convert should retain the legacy editor/runtime path");
  const converterPatch = await waitRecord((r) => r.id === "p12convert" && r.method === "PATCH" && r.update.exportResult, "legacy converter export", 15_000);
  assert.ok(converterPatch.update.exportResult.svg.includes("data-mml-node") || converterPatch.update.exportResult.svg.includes("mjx"), "hidden converter must still use legacy MathJax export in P12");

  console.log(JSON.stringify({
    status: "PASS",
    actualOfficeDialogBundle: true,
    actualBrowserKeyboard: true,
    interactiveCoreSelected: true,
    interactiveLegacyMathfields: mounted.legacy,
    initialLoadDirty: false,
    coreAutosaveExport: true,
    coreOfficeUndoRedo: true,
    staleSessionExportRejected: true,
    coreCommitUsesExistingSessionFlow: true,
    hiddenConverterCoreSelected: converterOwners.core,
    hiddenConverterLegacyMathfields: converterOwners.legacy,
    hiddenConverterLegacyExport: true,
    sessionPatchCount: records.filter((r) => r.method === "PATCH").length,
    wordAutomation: 0,
    officeProcessActions: 0,
  }, null, 2));
} finally {
  for (const client of clients) client.close();
  browser?.kill();
  await new Promise((resolveClose) => server.close(resolveClose));
}
