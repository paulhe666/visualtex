import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { randomUUID } from "node:crypto";
import { existsSync } from "node:fs";
import { readFile, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { extname, join, normalize, resolve, sep } from "node:path";
import http from "node:http";

const root = process.cwd();
const dist = resolve(root, "dist-office-windows-native");
const offset = process.pid % 700;
const port = 7900 + offset;
const debugPort = 14900 + offset;
const origin = `http://127.0.0.1:${port}`;
const profile = join(tmpdir(), `visualtex-p13-office-converter-perf-${process.pid}`);
const browserPath = [
  process.env.VISUALTEX_CHROME_PATH,
  process.env.PROGRAMFILES && join(process.env.PROGRAMFILES, "Microsoft", "Edge", "Application", "msedge.exe"),
  process.env["PROGRAMFILES(X86)"] && join(process.env["PROGRAMFILES(X86)"], "Microsoft", "Edge", "Application", "msedge.exe"),
  process.env.PROGRAMFILES && join(process.env.PROGRAMFILES, "Google", "Chrome", "Application", "chrome.exe"),
  process.env.LOCALAPPDATA && join(process.env.LOCALAPPDATA, "Google", "Chrome", "Application", "chrome.exe"),
].filter(Boolean).find((candidate) => existsSync(candidate));
if (!browserPath) throw new Error("No Chrome/Edge executable found for P13 converter performance acceptance.");
if (!existsSync(join(dist, "dialog", "index.html"))) {
  throw new Error("Build the Windows Office production bundle before P13 performance acceptance.");
}

const now = Date.now();
function makeSession(id, latex, objectMode, index) {
  const picture = objectMode === "crossPlatformPicture";
  return {
    id,
    mode: "create",
    host: picture ? "powerpoint" : "word",
    formulaId: randomUUID(),
    sourceDocumentId: `p13-perf-document-${index}`,
    sourceObjectId: null,
    title: `P13 perf ${id}`,
    lines: [{ id: `${id}-line`, latex }],
    activeLineId: `${id}-line`,
    codeFormat: "latex",
    displayMode: picture || objectMode === "mathTypeOle" ? "block" : (index % 2 ? "inline" : "block"),
    objectMode,
    numbered: false,
    mathTypeNumberPosition: "right",
    fontSizePt: picture ? 20 : 10.5,
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

const sessions = new Map();
const seed = makeSession("seed", String.raw`x+1`, "wordOmml", 0);
sessions.set(seed.id, seed);

const mixedIds = [];
const mixedModes = ["nativeOle", "wordOmml", "mathTypeOle", "crossPlatformPicture"];
for (let index = 0; index < 100; index += 1) {
  const objectMode = mixedModes[index % mixedModes.length];
  const id = `mixed-${String(index).padStart(3, "0")}`;
  let latex;
  if (objectMode === "mathTypeOle") {
    latex = String.raw`\begin{aligned}a_{${index}}&=b+${index}\\c&=d_{${index + 1}}\end{aligned}`;
  } else if (objectMode === "crossPlatformPicture") {
    latex = String.raw`\sqrt{x_{${index}}^2+y^2+1}`;
  } else if (objectMode === "nativeOle") {
    latex = String.raw`\frac{x_{${index}}+1}{y^2+${(index % 7) + 1}}`;
  } else {
    latex = String.raw`x_{${index}}+\sqrt{y^2+${(index % 5) + 1}}`;
  }
  sessions.set(id, makeSession(id, latex, objectMode, index));
  mixedIds.push(id);
}

const ommlIds = [];
for (let index = 0; index < 1000; index += 1) {
  const id = `omml-${String(index).padStart(4, "0")}`;
  const latex = index % 3 === 0
    ? String.raw`\frac{x_{${index}}+1}{y^2+${(index % 11) + 1}}`
    : index % 3 === 1
      ? String.raw`\sqrt{x_{${index}}^2+y^2}+${index % 17}`
      : String.raw`\sum_{k=1}^{${(index % 9) + 2}} k+x_{${index}}`;
  sessions.set(id, makeSession(id, latex, "wordOmml", index + 100));
  ommlIds.push(id);
}
const batchIds = [...mixedIds, ...ommlIds];

let batchServed = false;
let batchStartedAt = 0;
const terminal = [];
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
async function readBody(req) {
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
      if (!batchServed) {
        batchServed = true;
        batchStartedAt = Date.now();
        return json(res, { sessionIds: batchIds });
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
        const raw = await readBody(req);
        const update = raw ? JSON.parse(raw) : {};
        const merged = {
          ...current,
          ...update,
          id: current.id,
          formulaId: current.formulaId,
          createdAt: current.createdAt,
          updatedAt: Date.now(),
        };
        sessions.set(id, merged);
        if (id !== "seed" && (update.status === "committing" || update.status === "failed")) {
          const result = update.exportResult;
          terminal.push({
            id,
            status: update.status,
            at: Date.now(),
            objectMode: current.objectMode,
            hasSvg: Boolean(result?.svg),
            hasPng: Boolean(result?.pngBase64),
            hasMathMl: Boolean(result?.mathMl),
            error: update.error ?? null,
          });
        }
        return json(res, merged);
      }
    }
    const closeMatch = /^\/api\/v1\/app\/sessions\/([^/]+)\/close$/u.exec(url.pathname);
    if (closeMatch && req.method === "POST") {
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
  const started = Date.now();
  while (Date.now() - started < 15_000) {
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
async function evalValue(client, expression) {
  const result = await client.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
  if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description ?? result.exceptionDetails.text);
  return result.result.value;
}
async function waitForTerminals(count, timeout = 240_000) {
  const started = Date.now();
  while (Date.now() - started < timeout) {
    if (terminal.length >= count) return;
    await sleep(100);
  }
  const failures = terminal.filter((item) => item.status === "failed").slice(0, 10);
  throw new Error(`Timed out at ${terminal.length}/${count} terminal conversions: ${JSON.stringify(failures)}`);
}

await new Promise((resolveListen) => server.listen(port, "127.0.0.1", resolveListen));
let browser;
let client;
try {
  const url = `${origin}/dialog/index.html?sessionId=seed&runtime=vsto-convert`;
  browser = spawn(browserPath, [
    "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
    `--remote-debugging-port=${debugPort}`, `--user-data-dir=${profile}`, "--window-size=1000,700", url,
  ], { stdio: "ignore" });
  client = await connectPage(url);
  await waitForTerminals(batchIds.length);

  assert.equal(terminal.length, batchIds.length);
  assert.deepEqual(terminal.map((item) => item.id), batchIds, "P13 converter must preserve batch order across 1100 formulas");
  const failures = terminal.filter((item) => item.status !== "committing");
  assert.deepEqual(failures, [], `P13 performance batch failures: ${JSON.stringify(failures.slice(0, 10))}`);

  const mixed = terminal.slice(0, mixedIds.length);
  const omml = terminal.slice(mixedIds.length);
  assert.equal(mixed.length, 100);
  assert.equal(omml.length, 1000);
  assert.ok(mixed.every((item) => item.hasMathMl));
  assert.ok(omml.every((item) => item.objectMode === "wordOmml" && item.hasMathMl && !item.hasSvg && !item.hasPng));
  assert.ok(mixed.filter((item) => item.objectMode === "nativeOle").every((item) => item.hasSvg && item.hasPng));
  assert.ok(mixed.filter((item) => item.objectMode === "mathTypeOle").every((item) => item.hasSvg && !item.hasPng));
  assert.ok(mixed.filter((item) => item.objectMode === "crossPlatformPicture").every((item) => item.hasSvg && item.hasPng));
  assert.ok(mixed.filter((item) => item.objectMode === "wordOmml").every((item) => !item.hasSvg && !item.hasPng));

  const mixedCompletedAt = mixed.at(-1).at;
  const completedAt = terminal.at(-1).at;
  const resources = await evalValue(client, `({
    wasm:[...performance.getEntriesByType('resource')].filter(x=>String(x.name).includes('visualtex_engine_wasm_bg')).length,
    coreFont:[...performance.getEntriesByType('resource')].filter(x=>String(x.name).includes('latinmodern-math')).length,
    mathFields:document.querySelectorAll('math-field').length,
    coreInputs:document.querySelectorAll('.vt-core-input').length,
    coreEditors:document.querySelectorAll('[data-engine="visualtex-core"]').length
  })`);
  assert.ok(resources.wasm >= 1);
  assert.ok(resources.coreFont >= 1);
  assert.equal(resources.mathFields, 0);
  assert.equal(resources.coreInputs, 0);
  assert.equal(resources.coreEditors, 0);

  const mixedElapsedMs = mixedCompletedAt - batchStartedAt;
  const omml1000ElapsedMs = completedAt - mixedCompletedAt;
  const totalElapsedMs = completedAt - batchStartedAt;
  console.log(JSON.stringify({
    status: "PASS",
    productionOfficeBundle: true,
    mixed100: {
      formulas: mixed.length,
      failures: 0,
      elapsedMs: mixedElapsedMs,
      averageMsPerFormula: mixedElapsedMs / mixed.length,
      counts: Object.fromEntries(mixedModes.map((mode) => [mode, mixed.filter((item) => item.objectMode === mode).length])),
    },
    omml1000: {
      formulas: omml.length,
      failures: 0,
      elapsedMs: omml1000ElapsedMs,
      averageMsPerFormula: omml1000ElapsedMs / omml.length,
      mathMlOnly: true,
    },
    total: {
      formulas: terminal.length,
      failures: 0,
      elapsedMs: totalElapsedMs,
      orderPreserved: true,
      uniqueTerminalSessionIds: new Set(terminal.map((item) => item.id)).size,
    },
    resourceReuse: resources,
    wordAutomation: 0,
    officeProcessActions: 0,
  }, null, 2));
} finally {
  client?.close();
  browser?.kill();
  await new Promise((resolveClose) => server.close(resolveClose));
}
