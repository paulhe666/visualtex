import assert from "node:assert/strict";
import http from "node:http";
import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdtemp, readFile, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { extname, join, normalize, resolve, sep } from "node:path";
import { build } from "vite";

const root = process.cwd();
const outDir = await mkdtemp(join(tmpdir(), "visualtex-p14-runtime-recovery-"));
const entry = resolve(root, "scripts/fixtures/p14_core_runtime_recovery_entry.ts");
const offset = process.pid % 1000;
const debugPort = 15500 + offset;
const profile = join(outDir, "edge-profile");
const browserCandidates = [
  process.env.VISUALTEX_CHROME_PATH,
  process.env.PROGRAMFILES && join(process.env.PROGRAMFILES, "Microsoft", "Edge", "Application", "msedge.exe"),
  process.env["PROGRAMFILES(X86)"] && join(process.env["PROGRAMFILES(X86)"], "Microsoft", "Edge", "Application", "msedge.exe"),
  process.env.PROGRAMFILES && join(process.env.PROGRAMFILES, "Google", "Chrome", "Application", "chrome.exe"),
  process.env.LOCALAPPDATA && join(process.env.LOCALAPPDATA, "Google", "Chrome", "Application", "chrome.exe"),
].filter(Boolean);
const browserPath = browserCandidates.find((candidate) => existsSync(candidate));
if (!browserPath) throw new Error("No Edge/Chrome executable found for P14 runtime recovery acceptance.");

await build({
  configFile: false,
  root,
  logLevel: "error",
  build: {
    target: "chrome105",
    outDir,
    emptyOutDir: true,
    rollupOptions: {
      input: entry,
      output: { entryFileNames: "entry.js" },
    },
  },
});

const mime = {
  ".js": "text/javascript; charset=utf-8",
  ".wasm": "application/wasm",
  ".otf": "font/otf",
  ".woff2": "font/woff2",
  ".css": "text/css; charset=utf-8",
};
const html = Buffer.from('<!doctype html><html><head><meta charset="utf-8"><title>P14 runtime recovery</title></head><body><script type="module" src="/entry.js"></script></body></html>');
const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url ?? "/", "http://127.0.0.1");
    if (url.pathname === "/" || url.pathname === "/index.html") {
      res.writeHead(200, { "content-type": "text/html; charset=utf-8", "content-length": String(html.length), "cache-control": "no-store" });
      res.end(html); return;
    }
    const relative = decodeURIComponent(url.pathname).replace(/^\/+/, "");
    const file = normalize(resolve(outDir, relative));
    if (!file.startsWith(normalize(outDir + sep))) { res.writeHead(403); res.end(); return; }
    const info = await stat(file);
    if (!info.isFile()) { res.writeHead(404); res.end(); return; }
    const bytes = await readFile(file);
    res.writeHead(200, {
      "content-type": mime[extname(file)] ?? "application/octet-stream",
      "content-length": String(bytes.length),
      "cache-control": "no-store",
    });
    res.end(bytes);
  } catch {
    res.writeHead(404); res.end();
  }
});
await new Promise((resolveListen) => server.listen(0, "127.0.0.1", resolveListen));
const address = server.address();
if (!address || typeof address === "string") throw new Error("Unable to determine P14 recovery server port.");
const pageUrl = `http://127.0.0.1:${address.port}/`;

const sleep = (ms) => new Promise((resolveSleep) => setTimeout(resolveSleep, ms));
async function waitJson(url, timeout = 15_000) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    try {
      const response = await fetch(url);
      if (response.ok) return await response.json();
    } catch {}
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${url}`);
}
class Cdp {
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

let browser;
let cdp;
try {
  browser = spawn(browserPath, [
    "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
    `--remote-debugging-port=${debugPort}`, `--user-data-dir=${profile}`, "--window-size=1000,700", pageUrl,
  ], { stdio: "ignore" });
  const start = Date.now();
  let target;
  while (Date.now() - start < 15_000) {
    const targets = await waitJson(`http://127.0.0.1:${debugPort}/json/list`);
    target = targets.find((item) => item.type === "page" && item.url === pageUrl);
    if (target) break;
    await sleep(80);
  }
  if (!target) throw new Error("P14 recovery browser page did not appear.");
  cdp = new Cdp(target.webSocketDebuggerUrl);
  await cdp.connect();
  await cdp.send("Runtime.enable");
  const readyStart = Date.now();
  while (Date.now() - readyStart < 15_000) {
    const check = await cdp.send("Runtime.evaluate", { expression: "Boolean(window.p14RuntimeRecovery)", returnByValue: true });
    if (check.result.value) break;
    await sleep(80);
  }
  const result = await cdp.send("Runtime.evaluate", {
    expression: "window.p14RuntimeRecovery.run()",
    awaitPromise: true,
    returnByValue: true,
  });
  if (result.exceptionDetails) {
    throw new Error(result.exceptionDetails.exception?.description ?? result.exceptionDetails.text);
  }
  const report = result.result.value;
  assert.equal(report.status, "PASS");
  assert.equal(report.businessErrorKeptRuntime, true);
  assert.equal(report.fatalErrorRecognized, true);
  assert.equal(report.checkpointSource, "x+1+q");
  assert.equal(report.recoveredSource, "x+1+q");
  assert.equal(report.runtimeRebuilt, true);
  assert.equal(report.recoveredLayoutSvg, true);
  assert.equal(report.recoveredMathMl, true);
  assert.equal(report.oldOwnerStillReadable, true);
  console.log(JSON.stringify({ ...report, productionRuntimeModule: true, actualBrowserWasm: true, officeProcessActions: 0 }, null, 2));
} finally {
  cdp?.close();
  if (browser && browser.exitCode === null) {
    const exited = new Promise((resolveExit) => browser.once("exit", resolveExit));
    browser.kill();
    await Promise.race([exited, sleep(3_000)]);
  }
  await new Promise((resolveClose) => server.close(resolveClose));
  for (let attempt = 0; attempt < 10; attempt += 1) {
    try {
      await rm(outDir, { recursive: true, force: true });
      break;
    } catch (error) {
      if (!(error && typeof error === "object" && "code" in error && error.code === "EBUSY") || attempt === 9) throw error;
      await sleep(150);
    }
  }
}
