import http from "node:http";
import { spawn } from "node:child_process";
import { existsSync, readdirSync } from "node:fs";
import { readFile, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { extname, join, normalize, resolve, sep } from "node:path";
import assert from "node:assert/strict";

const root = process.cwd();
const dist = resolve(root, "dist-office-windows-native");
const assetsDir = join(dist, "assets");
const wasmName = readdirSync(assetsDir).find((name) => /^visualtex_engine_wasm_bg-.*\.wasm$/u.test(name));
const fontName = readdirSync(assetsDir).find((name) => /^latinmodern-math-.*\.otf$/u.test(name));
if (!wasmName || !fontName) throw new Error("P14 resource recovery requires the core WASM and verified math font assets.");

const offset = process.pid % 1000;
const port = 7900 + offset;
const debugPort = 14900 + offset;
const origin = `http://127.0.0.1:${port}`;
const profile = join(tmpdir(), `visualtex-p14-resource-recovery-${process.pid}`);
const candidates = [
  process.env.VISUALTEX_CHROME_PATH,
  process.env.PROGRAMFILES && join(process.env.PROGRAMFILES, "Microsoft", "Edge", "Application", "msedge.exe"),
  process.env["PROGRAMFILES(X86)"] && join(process.env["PROGRAMFILES(X86)"], "Microsoft", "Edge", "Application", "msedge.exe"),
  process.env.PROGRAMFILES && join(process.env.PROGRAMFILES, "Google", "Chrome", "Application", "chrome.exe"),
  process.env.LOCALAPPDATA && join(process.env.LOCALAPPDATA, "Google", "Chrome", "Application", "chrome.exe"),
].filter(Boolean);
const browserPath = candidates.find((candidate) => existsSync(candidate));
if (!browserPath) throw new Error("No Chrome/Edge executable found for P14 resource recovery acceptance.");

const now = Date.now();
function makeSession(id, latex) {
  return {
    id,
    mode: "create",
    host: "word",
    formulaId: `${id}-formula`,
    sourceDocumentId: `${id}-document`,
    sourceObjectId: null,
    title: `P14 ${id}`,
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
    autoCommitOnClose: false,
    explicitCancel: false,
    error: null,
    createdAt: now,
    updatedAt: now,
    expiresAt: now + 60 * 60 * 1000,
  };
}

const ids = ["badwasm", "badfont", "recovered"];
const sessions = new Map([
  ["badwasm", makeSession("badwasm", String.raw`x_1+1`)],
  ["badfont", makeSession("badfont", String.raw`\frac{a}{b}+2`)],
  ["recovered", makeSession("recovered", String.raw`\sqrt{x^2+y^2}+3`)],
]);
const originalSources = new Map([...sessions].map(([id, session]) => [id, session.lines[0].latex]));
const records = [];
let batchServed = false;
let corruptWasmResponses = 1;
let corruptFontResponses = 1;
let wasmRequests = 0;
let fontRequests = 0;

const json = (res, value, statusCode = 200) => {
  const bytes = Buffer.from(JSON.stringify(value));
  res.writeHead(statusCode, {
    "content-type": "application/json; charset=utf-8",
    "content-length": String(bytes.length),
    "cache-control": "no-store",
  });
  res.end(bytes);
};
async function body(req) {
  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  return Buffer.concat(chunks).toString("utf8");
}
const mime = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".wasm": "application/wasm",
  ".otf": "font/otf",
  ".woff2": "font/woff2",
};
function corrupt(bytes) {
  const next = Buffer.from(bytes);
  const index = Math.min(32, Math.max(0, next.length - 1));
  next[index] ^= 0xff;
  return next;
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
        editorPreferences: { settings: { language: "cn", theme: "light", editorLayout: "standard", zoom: 0.6, formulaLetterFont: "katex", formulaChineseFont: "system" } },
      });
    }
    if (url.pathname === "/api/v1/app/converter/next-batch" && req.method === "GET") {
      if (!batchServed) {
        batchServed = true;
        records.push({ method: "BATCH", sessionIds: [...ids], at: Date.now() });
        return json(res, { sessionIds: ids });
      }
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
        const merged = { ...current, ...update, id: current.id, formulaId: current.formulaId, createdAt: current.createdAt, updatedAt: Date.now() };
        sessions.set(id, merged);
        records.push({ method: "PATCH", id, update: structuredClone(update), at: Date.now() });
        return json(res, merged);
      }
    }
    const closeMatch = /^\/api\/v1\/app\/sessions\/([^/]+)\/close$/u.exec(url.pathname);
    if (closeMatch && req.method === "POST") {
      records.push({ method: "CLOSE", id: decodeURIComponent(closeMatch[1]), at: Date.now() });
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
    let bytes = await readFile(candidate);
    const name = candidate.slice(candidate.lastIndexOf(sep) + 1);
    if (name === wasmName) {
      wasmRequests += 1;
      if (corruptWasmResponses > 0) {
        corruptWasmResponses -= 1;
        bytes = corrupt(bytes);
      }
    } else if (name === fontName) {
      fontRequests += 1;
      if (corruptFontResponses > 0) {
        corruptFontResponses -= 1;
        bytes = corrupt(bytes);
      }
    }
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

const sleep = (ms) => new Promise((resolveSleep) => setTimeout(resolveSleep, ms));
async function waitHttp(url, timeout = 15_000) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    try { const response = await fetch(url); if (response.ok) return; } catch {}
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${url}`);
}
class CdpClient {
  constructor(url) { this.url = url; this.id = 1; this.pending = new Map(); }
  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((resolveConnect, reject) => {
      this.socket.addEventListener("open", resolveConnect, { once: true });
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
    return new Promise((resolveSend, reject) => {
      this.pending.set(id, { resolve: resolveSend, reject });
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
async function waitRecord(predicate, label, timeout = 30_000) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    const record = records.find(predicate);
    if (record) return record;
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${label}\n${JSON.stringify(records, null, 2)}`);
}

await new Promise((resolveListen) => server.listen(port, "127.0.0.1", resolveListen));
let browser;
let client;
try {
  const converterUrl = `${origin}/dialog/index.html?runtime=vsto-convert`;
  browser = spawn(browserPath, [
    "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
    `--remote-debugging-port=${debugPort}`, `--user-data-dir=${profile}`, "--window-size=1000,700", converterUrl,
  ], { stdio: "ignore" });
  client = await connectPage(converterUrl);

  const badWasm = await waitRecord((record) => record.id === "badwasm" && record.method === "PATCH" && record.update.status === "failed", "WASM integrity failure");
  assert.match(badWasm.update.error ?? "", /VisualTeX core WASM resource hash mismatch/u);
  assert.equal(sessions.get("badwasm").lines[0].latex, originalSources.get("badwasm"));
  assert.equal(sessions.get("badwasm").exportResult, null);

  const badFont = await waitRecord((record) => record.id === "badfont" && record.method === "PATCH" && record.update.status === "failed", "font integrity failure");
  assert.match(badFont.update.error ?? "", /VisualTeX core font resource hash mismatch/u);
  assert.equal(sessions.get("badfont").lines[0].latex, originalSources.get("badfont"));
  assert.equal(sessions.get("badfont").exportResult, null);

  const recovered = await waitRecord((record) => record.id === "recovered" && record.method === "PATCH" && record.update.status === "committing" && record.update.exportResult, "resource recovery commit");
  assert.equal(recovered.update.lines?.[0]?.latex, originalSources.get("recovered"));
  assert.ok(recovered.update.exportResult.svg?.includes("<svg"));
  assert.ok(recovered.update.exportResult.mathMl?.startsWith("<math"));
  assert.ok(recovered.update.exportResult.pngBase64?.length > 20);
  assert.ok(!/MathJax|data-mml-node|mjx/iu.test(recovered.update.exportResult.svg));

  const terminalOrder = records
    .filter((record) => ids.includes(record.id) && record.method === "PATCH" && ["failed", "committing"].includes(record.update.status))
    .map((record) => record.id);
  assert.deepEqual(terminalOrder, ids, "resource failures must not reorder later batch Sessions");
  assert.ok(wasmRequests >= 3, `Expected runtime retry to refetch verified WASM, saw ${wasmRequests}`);
  assert.ok(fontRequests >= 2, `Expected font retry after corruption, saw ${fontRequests}`);

  const surface = await evalValue(client, `({mathFields:document.querySelectorAll('math-field').length,coreInputs:document.querySelectorAll('.vt-core-input').length,coreEditors:document.querySelectorAll('[data-engine="visualtex-core"]').length})`);
  assert.deepEqual(surface, { mathFields: 0, coreInputs: 0, coreEditors: 0 });

  console.log(JSON.stringify({
    status: "PASS",
    wasmAsset: wasmName,
    fontAsset: fontName,
    wasmHashMismatchFailedClosed: true,
    fontHashMismatchFailedClosed: true,
    failedSessionSourcePreserved: true,
    retryWithoutReloadOrConfigurationDeletion: true,
    recoveredSessionCommitted: true,
    batchOrderPreserved: true,
    wasmRequests,
    fontRequests,
    surface,
    wordAutomation: 0,
    officeProcessActions: 0,
  }, null, 2));
} finally {
  client?.close();
  browser?.kill();
  await new Promise((resolveClose) => server.close(resolveClose));
}
