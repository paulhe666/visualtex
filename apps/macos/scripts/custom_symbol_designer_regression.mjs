// The custom symbol designer (apps/shared/custom-symbol-designer) in the built app, headless Chrome with a stub
// Tauri bridge: materials come in at their real size on the baseline, the TeX code follows the drawing, undo and
// redo, registering a command, and the designer surviving a narrow window. Run after `npm run build:desktop`.
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 700;
const port = 23300 + offset;
const debugPort = 30300 + offset;
const base = `http://127.0.0.1:${port}`;
const profile = `/tmp/visualtex-custom-symbol-designer-${process.pid}`;
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url) {
  for (let attempt = 0; attempt < 200; attempt += 1) {
    try {
      if ((await fetch(url)).ok) return;
    } catch {
      // still starting
    }
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

const stub = `(() => {
  const callbacks = new Map(); let next = 1;
  window.__TAURI_INTERNALS__ = {
    transformCallback(cb) { const id = next++; callbacks.set(id, cb); return id; },
    unregisterCallback(id) { callbacks.delete(id); },
    async invoke(cmd, args) { if (cmd === "plugin:event|listen") return args.handler; return null; },
    metadata: { currentWindow: { label: "main" }, currentWebview: { label: "main", windowLabel: "main" } },
    convertFileSrc: (p) => p,
  };
  window.__TAURI_EVENT_PLUGIN_INTERNALS__ = { unregisterListener() {} };
})();`;

const preview = spawn(process.execPath, ["node_modules/vite/bin/vite.js", "preview", "--host", "127.0.0.1", "--port", String(port), "--strictPort"], { stdio: "ignore" });
let chrome;
let socket;
try {
  await waitFor(base);
  chrome = spawn("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", [
    "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
    `--remote-debugging-port=${debugPort}`, `--user-data-dir=${profile}`, "about:blank",
  ], { stdio: "ignore" });
  await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
  const target = (await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json()).find((item) => item.type === "page");
  socket = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((resolve, reject) => { socket.onopen = resolve; socket.onerror = reject; });
  let id = 1;
  const pending = new Map();
  const errors = [];
  socket.onmessage = (event) => {
    const message = JSON.parse(event.data);
    if (message.method === "Runtime.exceptionThrown") errors.push(message.params.exceptionDetails.exception?.description ?? "exception");
    if (message.id && pending.has(message.id)) {
      const { resolve, reject } = pending.get(message.id);
      pending.delete(message.id);
      message.error ? reject(new Error(message.error.message)) : resolve(message.result);
    }
  };
  const send = (method, params = {}) => new Promise((resolve, reject) => {
    const n = id++;
    pending.set(n, { resolve, reject });
    socket.send(JSON.stringify({ id: n, method, params }));
  });
  const evaluate = async (expression) => {
    const result = await send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
    if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description ?? "evaluate failed");
    return result.result.value;
  };
  const viewport = (width, height) => send("Emulation.setDeviceMetricsOverride", { width, height, deviceScaleFactor: 1, mobile: false });
  const click = (selector) => evaluate(`(() => { const el = document.querySelector(${JSON.stringify(selector)}); if (!el) throw new Error("missing ${selector.replace(/"/g, "'")}"); el.click(); })()`);
  const setValue = (selector, value) => evaluate(`(() => {
    const el = document.querySelector(${JSON.stringify(selector)});
    if (!el) throw new Error("missing input");
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value").set.call(el, ${JSON.stringify(value)});
    el.dispatchEvent(new Event("input", { bubbles: true }));
  })()`);
  const addLatex = async (latex) => {
    await click('[data-csd-material-tab="latex"]');
    await setValue("[data-csd-source-input]", latex);
    await evaluate(`document.querySelector("[data-csd-source-input]").form.requestSubmit()`);
    await sleep(700);
  };
  const tex = () => evaluate(`(() => { document.querySelector('[data-csd-code-tab="tex"]').click(); return document.querySelector('[data-csd-code-text="tex"]').textContent; })()`);
  const layerCount = () => evaluate(`document.querySelectorAll("[data-csd-artwork]").length`);

  await send("Page.enable");
  await send("Runtime.enable");
  await viewport(1440, 900);
  await send("Page.addScriptToEvaluateOnNewDocument", { source: stub });
  await send("Page.navigate", { url: `${base}/index.html?transport=tauri&theme=light` });
  await sleep(4000);
  await click('[data-tile-category="custom"]');
  await sleep(400);
  await click("[data-open-custom-symbol-designer]");
  await sleep(1500);
  assert.ok(await evaluate(`Boolean(document.querySelector("[data-csd-dialog]"))`), "the designer opens");

  // ħ comes in at its real size (1000 units per em), its box fitting the glyph, its baseline on the baseline
  await addLatex("\\hbar");
  const hbar = await evaluate(`(() => {
    const points = document.querySelector("[data-csd-selection] polygon").getAttribute("points").split(" ").map((p) => p.split(",").map(Number));
    const ys = points.map((p) => p[1]);
    const xs = points.map((p) => p[0]);
    const baseline = Number(document.querySelector(".csd-guide.is-baseline line").getAttribute("y1"));
    return { height: Math.max(...ys) - Math.min(...ys), width: Math.max(...xs) - Math.min(...xs), bottom: Math.max(...ys), baseline };
  })()`);
  assert.ok(Math.abs(hbar.height - 705.9) < 2, `ħ box height ${hbar.height} (TeX: 0.695 + 0.011 em)`);
  assert.ok(Math.abs(hbar.width - 546.3) < 2, `ħ box width ${hbar.width}`);
  assert.ok(Math.abs(hbar.bottom - (hbar.baseline + 11)) < 1.5, "ħ sits on the baseline");
  assert.ok(await evaluate(`Boolean(document.querySelector("[data-csd-reference] path"))`), "the reference glyph is drawn");
  assert.match(await tex(), /\{\$\\hbar\$\}/, "the TeX code typesets the material");

  // a shape, then undo and redo
  await click('[data-csd-material-tab="shapes"]');
  await click('[data-csd-shape="circle"]');
  await sleep(300);
  assert.equal(await layerCount(), 2);
  assert.match(await tex(), /\\strokepath/, "the circle is a pict2e path in the TeX code");
  assert.match(await tex(), /\\begin\{picture\}/, "the command is a picture");
  assert.doesNotMatch(await tex(), /tikz/i, "no TikZ");
  await click("[data-csd-undo]");
  await sleep(200);
  assert.equal(await layerCount(), 1, "undo removes the circle");
  await click("[data-csd-redo]");
  await sleep(200);
  assert.equal(await layerCount(), 2, "redo brings it back");

  // register as \regtestsym
  await click('[data-csd-side-tab="register"]');
  await setValue("[data-csd-command]", "regtestsym");
  await sleep(200);
  await click("[data-csd-register]");
  await sleep(500);
  assert.equal(await evaluate(`document.querySelector("[data-csd-register-status]")?.getAttribute("data-csd-register-status")`), "success");
  await click("[data-csd-library-button]");
  await sleep(200);
  assert.ok(await evaluate(`Boolean(document.querySelector('[data-csd-symbol="regtestsym"]'))`), "the symbol is in the library");

  // a narrow window keeps the designer and its drawing
  await viewport(620, 760);
  await sleep(800);
  assert.ok(await evaluate(`Boolean(document.querySelector("[data-csd-dialog]"))`), "the designer stays open in a narrow window");
  assert.equal(await layerCount(), 2);
  await viewport(1440, 900);

  assert.deepEqual(errors, [], "no page errors");
  console.log("custom symbol designer: ok");
} finally {
  socket?.close();
  chrome?.kill("SIGKILL");
  preview.kill("SIGKILL");
  await rm(profile, { recursive: true, force: true }).catch(() => undefined);
}
