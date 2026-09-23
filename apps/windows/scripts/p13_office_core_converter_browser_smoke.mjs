import http from "node:http";
import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { randomUUID } from "node:crypto";
import { readFile, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { extname, join, normalize, resolve, sep } from "node:path";
import assert from "node:assert/strict";

const root = process.cwd();
const dist = resolve(root, "dist-office-windows-native");
const offset = process.pid % 1000;
const port = 17600 + offset;
const debugPort = 19600 + offset;
const origin = `http://127.0.0.1:${port}`;
const profile = join(tmpdir(), `visualtex-p13-office-converter-${process.pid}`);
const candidates = [
  process.env.VISUALTEX_CHROME_PATH,
  process.env.PROGRAMFILES && join(process.env.PROGRAMFILES, "Microsoft", "Edge", "Application", "msedge.exe"),
  process.env["PROGRAMFILES(X86)"] && join(process.env["PROGRAMFILES(X86)"], "Microsoft", "Edge", "Application", "msedge.exe"),
  process.env.PROGRAMFILES && join(process.env.PROGRAMFILES, "Google", "Chrome", "Application", "chrome.exe"),
  process.env.LOCALAPPDATA && join(process.env.LOCALAPPDATA, "Google", "Chrome", "Application", "chrome.exe"),
].filter(Boolean);
const browserPath = candidates.find((candidate) => existsSync(candidate));
if (!browserPath) throw new Error("No Chrome/Edge executable found for P13 Office converter acceptance.");
if (!existsSync(join(dist, "dialog", "index.html"))) {
  throw new Error("Build the Windows Office UI before running P13 converter acceptance.");
}

const now = Date.now();
function makeSession(id, latex, overrides = {}) {
  return {
    id,
    mode: "create",
    host: "word",
    formulaId: randomUUID(),
    sourceDocumentId: `${id}-document`,
    sourceObjectId: null,
    title: `P13 ${id}`,
    lines: [{ id: `${id}-line`, latex }],
    activeLineId: `${id}-line`,
    codeFormat: "latex",
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
    ...overrides,
  };
}

function withMetadata(session, extra = {}) {
  return {
    ...session,
    originalMetadata: {
      schema: "visualtex-formula",
      schemaVersion: 1,
      formulaId: session.formulaId,
      title: session.title,
      latex: session.lines.map((line) => line.latex).join("\n"),
      lines: session.lines.map((line) => ({ ...line })),
      codeFormat: session.codeFormat,
      displayMode: session.displayMode,
      numbered: session.numbered,
      fontSizePt: session.fontSizePt,
      renderFontSizePt: session.fontSizePt,
      createdWithVersion: "1.2.7",
      updatedWithVersion: "1.2.7",
      createdAt: new Date(now).toISOString(),
      updatedAt: new Date(now).toISOString(),
      ...extra,
    },
  };
}

const sessions = new Map([
  ["seed", makeSession("seed", String.raw`x+1`)],
  ["native", makeSession("native", String.raw`a_1+b^2`, { objectMode: "nativeOle" })],
  ["omml", withMetadata(makeSession("omml", String.raw`\sqrt{x^2+y^2}`, {
    objectMode: "wordOmml",
    displayMode: "block",
    exportWidth: 321.5,
    exportHeight: 87.25,
  }), { renderWidthPx: 321.5, renderHeightPx: 87.25, baseline: 61.75 })],
  ["mathtype", makeSession("mathtype", String.raw`\begin{aligned}a&=b&c&=d\\u&=v&w&=z\end{aligned}`, {
    objectMode: "mathTypeOle",
    displayMode: "block",
  })],
  ["picture", makeSession("picture", String.raw`\frac{p+q}{r+s}`, {
    host: "powerpoint",
    objectMode: "crossPlatformPicture",
    displayMode: "block",
    fontSizePt: 20,
  })],
  ["bad", makeSession("bad", String.raw`\visualtexDefinitelyUnsupported{q}`, { objectMode: "nativeOle" })],
  ["afterbad", makeSession("afterbad", String.raw`z_9+7`, { objectMode: "wordOmml" })],
  ["tagged", withMetadata(makeSession("tagged", String.raw`x+y`, {
    objectMode: "wordOmml",
    displayMode: "block",
  }), { equationTag: "4.8.4" })],
  ["taggednative", withMetadata(makeSession("taggednative", String.raw`x^2+y^2`, {
    objectMode: "nativeOle",
    displayMode: "block",
  }), { equationTag: "A" })],
  ["taggedmathtype", withMetadata(makeSession("taggedmathtype", String.raw`a+b=c`, {
    objectMode: "mathTypeOle",
    displayMode: "block",
  }), { equationTag: "*", formulaLetterFont: "times", formulaChineseFont: "songti" })],
]);

const batchIds = [
  "native", "omml", "mathtype", "picture", "bad", "afterbad",
  "tagged", "taggednative", "taggedmathtype",
];
let batchServed = false;
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
      if (!batchServed) {
        batchServed = true;
        records.push({ at: Date.now(), method: "BATCH", sessionIds: [...batchIds] });
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
        records.push({ at: Date.now(), id, method: "PATCH", update: structuredClone(update) });
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

const sleep = (ms) => new Promise((resolveSleep) => setTimeout(resolveSleep, ms));
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
async function waitRecord(predicate, label, timeout = 25_000) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    const found = records.find(predicate);
    if (found) return found;
    await sleep(80);
  }
  throw new Error(`Timed out waiting for record: ${label}\n${JSON.stringify(records, null, 2)}`);
}
const commitRecord = (id) => records.find((record) =>
  record.id === id && record.method === "PATCH" && record.update.status === "committing");
const failureRecord = (id) => records.find((record) =>
  record.id === id && record.method === "PATCH" && record.update.status === "failed");

await new Promise((resolveListen) => server.listen(port, "127.0.0.1", resolveListen));
let browser;
let client;
try {
  const converterUrl = `${origin}/dialog/index.html?sessionId=seed&runtime=vsto-convert`;
  browser = spawn(browserPath, [
    "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
    `--remote-debugging-port=${debugPort}`, `--user-data-dir=${profile}`, "--window-size=1000,700", converterUrl,
  ], { stdio: "ignore" });
  client = await connectPage(converterUrl);

  await waitEval(client, `Boolean(document.querySelector('[data-office-headless-converter]'))`, "headless converter surface");
  const surface = await evalValue(client, `({
    headless:Boolean(document.querySelector('[data-office-headless-converter]')),
    mathFields:document.querySelectorAll('math-field').length,
    coreInputs:document.querySelectorAll('.vt-core-input').length,
    coreEditors:document.querySelectorAll('[data-engine="visualtex-core"]').length
  })`);
  assert.deepEqual(surface, { headless: true, mathFields: 0, coreInputs: 0, coreEditors: 0 });

  await waitRecord((record) => commitRecord("seed") === record, "single headless seed export");
  for (const id of batchIds) {
    if (id === "bad") await waitRecord((record) => failureRecord(id) === record, `${id} failure`);
    else await waitRecord((record) => commitRecord(id) === record, `${id} commit`);
  }

  const seed = commitRecord("seed").update.exportResult;
  assert.ok(seed.svg && !/MathJax|data-mml-node|mjx/iu.test(seed.svg));
  assert.ok(seed.mathMl?.startsWith("<math"));
  assert.ok(seed.pngBase64?.length > 20);

  const native = commitRecord("native").update.exportResult;
  assert.ok(native.svg && !/MathJax|data-mml-node|mjx/iu.test(native.svg));
  assert.ok(native.mathMl?.includes("<msub") || native.mathMl?.includes("<msup"));
  assert.ok(native.pngBase64?.length > 20);

  const omml = commitRecord("omml").update.exportResult;
  assert.equal(omml.svg, "");
  assert.equal(omml.svgBase64, "");
  assert.equal(omml.pngBase64, undefined);
  assert.equal(omml.width, 321.5);
  assert.equal(omml.height, 87.25);
  assert.equal(omml.baseline, 61.75);
  assert.ok(omml.mathMl?.startsWith("<math"));

  const mathType = commitRecord("mathtype").update.exportResult;
  assert.ok(mathType.svg && !/MathJax|data-mml-node|mjx/iu.test(mathType.svg));
  assert.equal(mathType.pngBase64, undefined);
  assert.equal(mathType.formulaLetterFont, "times");
  assert.match(mathType.mathMl ?? "", /data-mtef-tabs="true"/u);
  assert.match(mathType.mathMl ?? "", /data-visualtex-mtef-ruler-stops="[0-9,]+"/u);

  const picture = commitRecord("picture").update.exportResult;
  assert.ok(picture.svg && !/MathJax|data-mml-node|mjx/iu.test(picture.svg));
  assert.ok(picture.mathMl?.startsWith("<math"));
  assert.ok(picture.pngBase64?.length > 20);

  const failure = failureRecord("bad");
  assert.ok(failure.update.error && !failure.update.exportResult);
  assert.ok(commitRecord("afterbad"), "batch must continue after an item failure");

  const tagged = commitRecord("tagged").update.exportResult;
  assert.equal(tagged.svg, "");
  assert.equal(tagged.svgBase64, "");
  assert.equal(tagged.pngBase64, undefined);
  assert.match(tagged.mathMl ?? "", /<mlabeledtr>/u);
  assert.match(tagged.mathMl ?? "", /\(4\.8\.4\)/u);

  const taggedNative = commitRecord("taggednative").update.exportResult;
  assert.match(taggedNative.mathMl ?? "", /<mlabeledtr>/u);
  assert.match(taggedNative.mathMl ?? "", /<mtext>\(A\)<\/mtext>/u);
  assert.ok(taggedNative.svg.includes("<g transform=\"translate("));
  assert.ok(!/MathJax|data-mml-node|mjx|<text\b|<image\b|<script\b/iu.test(taggedNative.svg));
  assert.ok(taggedNative.pngBase64?.length > 20);

  const taggedMathType = commitRecord("taggedmathtype").update.exportResult;
  assert.match(taggedMathType.mathMl ?? "", /<mlabeledtr>/u);
  assert.match(taggedMathType.mathMl ?? "", /<mtext>\(\*\)<\/mtext>/u);
  assert.ok(taggedMathType.svg.includes("<g transform=\"translate("));
  assert.ok(!/MathJax|data-mml-node|mjx|<text\b|<image\b|<script\b/iu.test(taggedMathType.svg));
  assert.equal(taggedMathType.pngBase64, undefined);
  assert.equal(taggedMathType.formulaLetterFont, "times");

  const batchTerminalOrder = records
    .filter((record) => batchIds.includes(record.id) && record.method === "PATCH"
      && (record.update.status === "committing" || record.update.status === "failed"))
    .map((record) => record.id);
  assert.deepEqual(batchTerminalOrder, batchIds, "converter batch terminal writes must preserve requested order");
  assert.equal(new Set(batchTerminalOrder).size, batchIds.length, "each batch Session must receive exactly one terminal write");

  const resources = await evalValue(client, `({
    wasm:[...performance.getEntriesByType('resource')].filter(x=>String(x.name).includes('visualtex_engine_wasm_bg')).length,
    coreFont:[...performance.getEntriesByType('resource')].filter(x=>String(x.name).includes('latinmodern-math')).length
  })`);
  assert.ok(resources.wasm >= 1, "headless converter must load the core WASM runtime");
  assert.ok(resources.coreFont >= 1, "headless converter must load the verified core math font");

  console.log(JSON.stringify({
    status: "PASS",
    productionOfficeBundle: true,
    headlessSurface: surface,
    objectModes: {
      nativeOle: "core-svg-mathml-png",
      wordOmml: "core-mathml-only",
      mathTypeOle: "core-svg-mathml-no-png",
      crossPlatformPicture: "core-svg-mathml-png",
    },
    equationTagCoreCompatibility: true,
    batchOrderPreserved: true,
    failureContinues: true,
    identityWrites: batchTerminalOrder,
    coreResources: resources,
    wordAutomation: 0,
    officeProcessActions: 0,
  }, null, 2));
} finally {
  client?.close();
  browser?.kill();
  await new Promise((resolveClose) => server.close(resolveClose));
}
