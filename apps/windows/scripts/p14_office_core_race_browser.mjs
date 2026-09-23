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
const port = 8100 + offset;
const debugPort = 15100 + offset;
const origin = `http://127.0.0.1:${port}`;
const profile = join(tmpdir(), `visualtex-p14-office-races-${process.pid}`);
const candidates = [
  process.env.VISUALTEX_CHROME_PATH,
  process.env.PROGRAMFILES && join(process.env.PROGRAMFILES, "Microsoft", "Edge", "Application", "msedge.exe"),
  process.env["PROGRAMFILES(X86)"] && join(process.env["PROGRAMFILES(X86)"], "Microsoft", "Edge", "Application", "msedge.exe"),
  process.env.PROGRAMFILES && join(process.env.PROGRAMFILES, "Google", "Chrome", "Application", "chrome.exe"),
  process.env.LOCALAPPDATA && join(process.env.LOCALAPPDATA, "Google", "Chrome", "Application", "chrome.exe"),
].filter(Boolean);
const browserPath = candidates.find((candidate) => existsSync(candidate));
if (!browserPath) throw new Error("No Chrome/Edge executable found for P14 Office race acceptance.");
if (!existsSync(join(dist, "dialog", "index.html"))) throw new Error("Build the Windows Office production UI before P14 race acceptance.");

const now = Date.now();
function makeSession(id, latex) {
  return {
    id,
    mode: "edit",
    host: "word",
    formulaId: `${id}-formula`,
    sourceDocumentId: `${id}-document`,
    sourceObjectId: `${id}-object`,
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
    autoCommitOnClose: true,
    explicitCancel: false,
    error: null,
    createdAt: now,
    updatedAt: now,
    expiresAt: now + 60 * 60 * 1000,
  };
}
const sessions = new Map([
  ["raceA", makeSession("raceA", "a=1")],
  ["raceB", makeSession("raceB", "b=2")],
  ["double", makeSession("double", "d=4")],
  ["closeRace", makeSession("closeRace", "c=3")],
]);
const records = [];
const commitDueAt = new Map();

const json = (res, value, statusCode = 200) => {
  const bytes = Buffer.from(JSON.stringify(value));
  res.writeHead(statusCode, { "content-type": "application/json; charset=utf-8", "content-length": String(bytes.length), "cache-control": "no-store" });
  res.end(bytes);
};
const mime = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8", ".wasm": "application/wasm", ".otf": "font/otf", ".woff2": "font/woff2" };
async function body(req) { const chunks = []; for await (const chunk of req) chunks.push(chunk); return Buffer.concat(chunks).toString("utf8"); }

const delayScript = Buffer.from(String.raw`(()=>{
const original=HTMLCanvasElement.prototype.toBlob;
window.__p14DelayNextPngMs=0;
HTMLCanvasElement.prototype.toBlob=function(callback,...args){const delay=Number(window.__p14DelayNextPngMs||0);window.__p14DelayNextPngMs=0;return original.call(this,value=>delay?setTimeout(()=>callback(value),delay):callback(value),...args);};
})();`);

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url ?? "/", origin);
    if (url.pathname === "/p14-delay.js") {
      res.writeHead(200, { "content-type": "text/javascript; charset=utf-8", "content-length": String(delayScript.length), "cache-control": "no-store" });
      res.end(delayScript); return;
    }
    if (url.pathname === "/api/v1/theme" && req.method === "GET") return json(res, { theme: "light", editorLayout: "standard" });
    if (url.pathname === "/api/v1/preferences" && req.method === "GET") return json(res, { powerpointDefaultFontSizePt: 20, editorPreferences: { settings: { language: "cn", theme: "light", editorLayout: "standard", zoom: 0.6, formulaLetterFont: "katex", formulaChineseFont: "system" } } });
    if (url.pathname === "/api/v1/app/converter/next-batch" && req.method === "GET") return json(res, { sessionIds: [] });
    const sessionMatch = /^\/api\/v1\/sessions\/([^/]+)$/u.exec(url.pathname);
    if (sessionMatch) {
      const id = decodeURIComponent(sessionMatch[1]);
      let current = sessions.get(id);
      if (!current) return json(res, { error: "missing" }, 404);
      if (req.method === "GET") {
        const due = commitDueAt.get(id);
        if (due && Date.now() >= due && current.status !== "completed") {
          current = { ...current, status: "completed", updatedAt: Date.now() };
          sessions.set(id, current);
          commitDueAt.delete(id);
        }
        return json(res, current);
      }
      if (req.method === "PATCH") {
        const raw = await body(req);
        const update = raw ? JSON.parse(raw) : {};
        const merged = { ...current, ...update, id: current.id, formulaId: current.formulaId, createdAt: current.createdAt, updatedAt: Date.now() };
        sessions.set(id, merged);
        records.push({ at: Date.now(), id, method: "PATCH", update: structuredClone(update) });
        if (update.status === "committing") commitDueAt.set(id, Date.now() + 250);
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
    if (!candidate.startsWith(normalize(dist + sep))) { res.writeHead(403); res.end("forbidden"); return; }
    let info; try { info = await stat(candidate); } catch { res.writeHead(404); res.end("not found"); return; }
    if (!info.isFile()) { res.writeHead(404); res.end("not found"); return; }
    let bytes = await readFile(candidate);
    if (candidate.endsWith(`${sep}dialog${sep}index.html`)) {
      bytes = Buffer.from(bytes.toString("utf8").replace('<script type="module"', '<script src="/p14-delay.js"></script>\n    <script type="module"'));
    }
    res.writeHead(200, { "content-type": mime[extname(candidate)] ?? "application/octet-stream", "content-length": String(bytes.length), "cache-control": "no-store" });
    res.end(bytes);
  } catch (error) {
    res.writeHead(500, { "content-type": "text/plain; charset=utf-8" });
    res.end(error instanceof Error ? error.stack ?? error.message : String(error));
  }
});

const sleep = (ms) => new Promise((resolveSleep) => setTimeout(resolveSleep, ms));
async function waitHttp(url, timeout = 15_000) { const start = Date.now(); while (Date.now() - start < timeout) { try { const response = await fetch(url); if (response.ok) return; } catch {} await sleep(80); } throw new Error(`Timed out waiting for ${url}`); }
class CdpClient {
  constructor(url) { this.url = url; this.id = 1; this.pending = new Map(); }
  async connect() { this.socket = new WebSocket(this.url); await new Promise((resolveConnect, reject) => { this.socket.addEventListener("open", resolveConnect, { once: true }); this.socket.addEventListener("error", reject, { once: true }); }); this.socket.addEventListener("message", (event) => { const message = JSON.parse(event.data); const pending = this.pending.get(message.id); if (!pending) return; this.pending.delete(message.id); message.error ? pending.reject(new Error(message.error.message)) : pending.resolve(message.result); }); }
  send(method, params = {}) { const id = this.id++; return new Promise((resolveSend, reject) => { this.pending.set(id, { resolve: resolveSend, reject }); this.socket.send(JSON.stringify({ id, method, params })); }); }
  close() { this.socket?.close(); }
}
async function connectPage(url) {
  await waitHttp(`http://127.0.0.1:${debugPort}/json/list`);
  const start = Date.now();
  while (Date.now() - start < 15_000) {
    const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
    const page = targets.find((target) => target.type === "page" && target.url === url);
    if (page) { const client = new CdpClient(page.webSocketDebuggerUrl); await client.connect(); await client.send("Runtime.enable"); await client.send("Page.enable"); return client; }
    await sleep(80);
  }
  throw new Error(`No browser page for ${url}`);
}
const evalValue = async (client, expression) => { const result = await client.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true }); if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description ?? result.exceptionDetails.text); return result.result.value; };
async function waitEval(client, expression, label, timeout = 15_000) { const start = Date.now(); while (Date.now() - start < timeout) { if (await evalValue(client, expression)) return; await sleep(80); } throw new Error(`Timed out waiting for ${label}`); }
async function waitRecord(predicate, label, timeout = 15_000) { const start = Date.now(); while (Date.now() - start < timeout) { const found = records.find(predicate); if (found) return found; await sleep(80); } throw new Error(`Timed out waiting for ${label}\n${JSON.stringify(records, null, 2)}`); }
async function insertText(client, text) { await evalValue(client, `document.querySelector('.vt-core-input').focus();true`); await client.send("Input.insertText", { text }); }
async function point(client, selector) { const p = await evalValue(client, `(()=>{const e=document.querySelector(${JSON.stringify(selector)});if(!e)return null;e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()`); if (!p) throw new Error(`Missing selector ${selector}`); return p; }
async function clickBurst(client, selector, count) { const p = await point(client, selector); for (let index = 0; index < count; index += 1) { await client.send("Input.dispatchMouseEvent", { type: "mousePressed", ...p, button: "left", clickCount: 1 }); await client.send("Input.dispatchMouseEvent", { type: "mouseReleased", ...p, button: "left", clickCount: 1 }); } }

await new Promise((resolveListen) => server.listen(port, "127.0.0.1", resolveListen));
let browser;
const clients = [];
try {
  const urlA = `${origin}/dialog/index.html?sessionId=raceA&runtime=vsto-desktop`;
  browser = spawn(browserPath, ["--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check", `--remote-debugging-port=${debugPort}`, `--user-data-dir=${profile}`, "--window-size=1400,1000", urlA], { stdio: "ignore" });
  const pageA = await connectPage(urlA); clients.push(pageA);
  await waitEval(pageA, `Boolean(document.querySelector('[data-engine="visualtex-core"] .vt-core-input'))`, "race A core editor");
  const urlB = `${origin}/dialog/index.html?sessionId=raceB&runtime=vsto-desktop`;
  const targetB = await pageA.send("Target.createTarget", { url: urlB }); assert.ok(targetB.targetId);
  const pageB = await connectPage(urlB); clients.push(pageB);
  await waitEval(pageB, `Boolean(document.querySelector('[data-engine="visualtex-core"] .vt-core-input'))`, "race B core editor");

  await Promise.all([insertText(pageA, "+A"), insertText(pageB, "+B")]);
  const patchA = await waitRecord((record) => record.id === "raceA" && record.method === "PATCH" && record.update.lines?.[0]?.latex === "a=1+A" && record.update.exportResult, "race A export");
  const patchB = await waitRecord((record) => record.id === "raceB" && record.method === "PATCH" && record.update.lines?.[0]?.latex === "b=2+B" && record.update.exportResult, "race B export");
  assert.equal(sessions.get("raceA").lines[0].latex, "a=1+A");
  assert.equal(sessions.get("raceB").lines[0].latex, "b=2+B");
  assert.ok(!records.some((record) => record.id === "raceA" && record.update?.lines?.[0]?.latex?.includes("+B")));
  assert.ok(!records.some((record) => record.id === "raceB" && record.update?.lines?.[0]?.latex?.includes("+A")));
  assert.ok(patchA.update.exportResult.pngBase64 && patchB.update.exportResult.pngBase64);

  await evalValue(pageA, `window.dispatchEvent(new CustomEvent('visualtex-office-session',{detail:{sessionId:'double'}}));true`);
  await waitEval(pageA, `Boolean(document.querySelector('[data-engine="visualtex-core"] .vt-core-input'))`, "double-submit session");
  await insertText(pageA, "+x");
  await waitRecord((record) => record.id === "double" && record.method === "PATCH" && record.update.lines?.[0]?.latex === "d=4+x" && record.update.exportResult, "double-submit export");
  await clickBurst(pageA, "[data-office-primary-action]", 2);
  await waitRecord((record) => record.id === "double" && record.method === "CLOSE", "double-submit close");
  await sleep(350);
  assert.equal(records.filter((record) => record.id === "double" && record.method === "PATCH" && record.update.status === "committing").length, 1, "double Apply created multiple commits");
  assert.equal(records.filter((record) => record.id === "double" && record.method === "CLOSE").length, 1, "double Apply closed the Session more than once");

  await evalValue(pageA, `window.dispatchEvent(new CustomEvent('visualtex-office-session',{detail:{sessionId:'closeRace'}}));true`);
  await waitEval(pageA, `Boolean(document.querySelector('[data-engine="visualtex-core"] .vt-core-input'))`, "close-during-export session");
  await evalValue(pageA, `window.__p14DelayNextPngMs=700;document.querySelector('.vt-core-input').focus();true`);
  await pageA.send("Input.insertText", { text: "+late" });
  await waitRecord((record) => record.id === "closeRace" && record.method === "PATCH" && record.update.lines?.[0]?.latex === "c=3+late" && !record.update.exportResult, "close-race source autosave");
  const closeRequestedAt = Date.now();
  await evalValue(pageA, `window.dispatchEvent(new CustomEvent('visualtex-office-close-requested'));true`);
  const closeCommit = await waitRecord((record) => record.id === "closeRace" && record.method === "PATCH" && record.update.status === "committing", "close-race commit", 20_000);
  assert.equal(closeCommit.update.lines?.[0]?.latex, "c=3+late");
  assert.ok(closeCommit.update.exportResult?.pngBase64?.length > 20, "native close committed without a complete current export");
  const closeRecord = await waitRecord((record) => record.id === "closeRace" && record.method === "CLOSE", "close-race close", 20_000);
  await sleep(900);
  assert.equal(records.filter((record) => record.id === "closeRace" && record.method === "PATCH" && record.update.status === "committing").length, 1);
  const afterCloseWrites = records.filter((record) => record.id === "closeRace" && record.method === "PATCH" && record.at > closeRecord.at);
  assert.equal(afterCloseWrites.length, 0, `stale export wrote after close: ${JSON.stringify(afterCloseWrites)}`);
  assert.ok(closeCommit.at >= closeRequestedAt);

  console.log(JSON.stringify({
    status: "PASS",
    productionOfficeBundle: true,
    multiWindowIsolation: true,
    parallelSessionSources: { raceA: sessions.get("raceA").lines[0].latex, raceB: sessions.get("raceB").lines[0].latex },
    doubleApplySingleCommit: true,
    closeDuringExportSingleCommit: true,
    staleWriteAfterClose: false,
    wordAutomation: 0,
    officeProcessActions: 0,
  }, null, 2));
} finally {
  for (const client of clients) client.close();
  browser?.kill();
  await new Promise((resolveClose) => server.close(resolveClose));
}
