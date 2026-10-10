import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createServer } from "node:http";
import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { extname, join, normalize, resolve } from "node:path";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

// Responsive layout guards for the Windows desktop and Office editors:
// - the bottom dock and the header rows never paint controls over each other,
//   and the dock never shrinks them with transform: scale, at any window width;
// - panel tabs collapse to icons when the dock itself is narrow;
// - menus do not move after they appear, also while the pointer moves;
// - the formula color picker escapes the scrolling formatting strip;
// - the layout is at rest once a resize is over (no measure/resize loops in
//   tile previews, toolbar rows or the matrix picker);
// - the Office editor keeps the formula canvas wider than the tile rail, also
//   after the rail was widened in a larger window.
//
// Run from apps/windows after `npm run build:all`. VT_DESKTOP_DIST /
// VT_OFFICE_DIST pick other builds (default dist/ and
// dist-office-windows-native/), VT_ONLY=desktop|office runs one window,
// VT_SCAN=1 sweeps every 3-4 px and only reports, VT_REPORT=<file> writes all
// measurements, VT_SHOTS=<dir> saves screenshots.

const sleep = (ms) => new Promise((done) => setTimeout(done, ms));
const offset = process.pid % 500;
const port = 24400 + offset;
const debugPort = 25400 + offset;
const baseUrl = `http://127.0.0.1:${port}`;
const sessionId = "11111111-2222-4333-8444-555555555555";
const desktopDist = resolve(process.env.VT_DESKTOP_DIST || "dist");
const officeDist = resolve(process.env.VT_OFFICE_DIST || "dist-office-windows-native");
const scan = process.env.VT_SCAN === "1";
const reportPath = process.env.VT_REPORT;
const shotsDir = process.env.VT_SHOTS;
const only = process.env.VT_ONLY; // "desktop" or "office"
const chromeProfile = createBrowserProfilePath("visualtex-responsive-layout");
const chromePath = resolveChromiumExecutable();
let distRoot = desktopDist;

const mimeTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".woff2": "font/woff2",
  ".woff": "font/woff",
  ".ttf": "font/ttf",
  ".otf": "font/otf",
  ".wasm": "application/wasm",
};

let session = {
  id: sessionId,
  mode: "create",
  host: "word",
  formulaId: "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee",
  sourceDocumentId: "responsive-layout-regression",
  sourceObjectId: "visualtex-word-vsto-range:0:0",
  title: "Word Formula",
  lines: [{ id: "99999999-8888-4777-8666-555555555555", latex: "x^2+y^2=z^2" }],
  activeLineId: "99999999-8888-4777-8666-555555555555",
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
  status: "editing",
  autoCommitOnClose: true,
  explicitCancel: false,
  error: null,
  createdAt: Date.now(),
  updatedAt: Date.now(),
  expiresAt: Date.now() + 3_600_000,
};

function writeJson(response, status, value) {
  response.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
  });
  response.end(JSON.stringify(value));
}

async function readJsonBody(request) {
  const chunks = [];
  for await (const chunk of request) chunks.push(chunk);
  const source = Buffer.concat(chunks).toString("utf8");
  return source ? JSON.parse(source) : {};
}

const server = createServer(async (request, response) => {
  const url = new URL(request.url ?? "/", baseUrl);
  try {
    if (url.pathname === "/api/v1/theme") {
      writeJson(response, 200, { theme: "light", editorLayout: "classic" });
      return;
    }
    if (url.pathname === "/api/v1/preferences") {
      writeJson(response, 200, {
        powerpointDefaultFontSizePt: 20,
        editorPreferences: {
          settings: {
            editorLayout: "classic",
            sourceOpen: false,
            classicTileWidth: 300,
            classicDockHeight: 240,
          },
        },
      });
      return;
    }
    if (url.pathname === `/api/v1/sessions/${sessionId}`) {
      if (request.method === "PATCH") {
        session = { ...session, ...(await readJsonBody(request)), updatedAt: Date.now() };
      }
      writeJson(response, 200, session);
      return;
    }
    if (url.pathname.startsWith("/api/v1/")) {
      writeJson(response, 503, { error: "unused in responsive layout regression" });
      return;
    }
    if (url.pathname === "/" || url.pathname.startsWith("/dialog/")) {
      const page = url.pathname === "/" ? "index.html" : join("dialog", "index.html");
      const source = await readFile(join(distRoot, page), "utf8");
      const meta = [
        '<meta name="visualtex-install-token" content="responsive-layout-regression" />',
        '<meta name="visualtex-native-powerpoint-commit" content="false" />',
        '<meta name="visualtex-theme" content="light" />',
        '<meta name="visualtex-editor-layout" content="classic" />',
      ].join("\n");
      response.writeHead(200, { "Content-Type": "text/html; charset=utf-8", "Cache-Control": "no-store" });
      response.end(source.replace("</head>", `${meta}\n</head>`));
      return;
    }
    const relative = normalize(decodeURIComponent(url.pathname).slice(1));
    if (relative.startsWith("..")) {
      response.writeHead(403).end();
      return;
    }
    const content = await readFile(join(distRoot, relative));
    response.writeHead(200, {
      "Content-Type": mimeTypes[extname(relative)] ?? "application/octet-stream",
      "Cache-Control": "no-store",
    });
    response.end(content);
  } catch (error) {
    response.writeHead(error?.code === "ENOENT" ? 404 : 500, { "Content-Type": "text/plain; charset=utf-8" });
    response.end(String(error));
  }
});

class CdpClient {
  constructor(url) {
    this.url = url;
    this.nextId = 1;
    this.pending = new Map();
  }

  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((done, fail) => {
      this.socket.addEventListener("open", done, { once: true });
      this.socket.addEventListener("error", fail, { once: true });
    });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
      if (!message.id) return;
      const pending = this.pending.get(message.id);
      if (!pending) return;
      this.pending.delete(message.id);
      if (message.error) pending.reject(new Error(message.error.message));
      else pending.resolve(message.result);
    });
  }

  send(method, params = {}) {
    const id = this.nextId++;
    return new Promise((done, fail) => {
      this.pending.set(id, { resolve: done, reject: fail });
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }

  async evaluate(expression) {
    const result = await this.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
    if (result.exceptionDetails) {
      throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text || "Runtime.evaluate failed");
    }
    return result.result.value;
  }

  close() {
    this.socket?.close();
  }
}

async function waitForUrl(url, timeoutMs = 15_000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    try {
      if ((await fetch(url)).ok) return;
    } catch {}
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

// ---- In-page probes ------------------------------------------------------

// Controls of one row (dock or header): visible part (clipped by scrolling or
// clipping ancestors), pairwise overlaps, and anything shrunk by a transform.
const overlapReport = (rootSelector) => String.raw`(() => {
  const dock = document.querySelector(${JSON.stringify(rootSelector)});
  if (!dock) return null;
  const visible = (element) => {
    let { left, top, right, bottom } = element.getBoundingClientRect();
    for (let parent = element.parentElement; parent && parent !== document.body; parent = parent.parentElement) {
      const style = getComputedStyle(parent);
      if (!/(hidden|clip|auto|scroll)/.test(style.overflowX + style.overflowY)) continue;
      const box = parent.getBoundingClientRect();
      left = Math.max(left, box.left); top = Math.max(top, box.top);
      right = Math.min(right, box.right); bottom = Math.min(bottom, box.bottom);
    }
    return { left, top, right, bottom };
  };
  const name = (button) => button.getAttribute('aria-label') || button.title || button.textContent.trim() || button.className;
  const all = [...dock.querySelectorAll('button, select, input:not([type=hidden])')].filter((button) => button.getClientRects().length);
  const buttons = all.filter((button) => {
    const box = visible(button);
    return box.right - box.left > 2 && box.bottom - box.top > 2;
  });
  const overlaps = [];
  for (let i = 0; i < buttons.length; i++) for (let j = i + 1; j < buttons.length; j++) {
    if (buttons[i].contains(buttons[j]) || buttons[j].contains(buttons[i])) continue;
    const a = visible(buttons[i]); const b = visible(buttons[j]);
    const width = Math.min(a.right, b.right) - Math.max(a.left, b.left);
    const height = Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top);
    if (width > 1 && height > 1) overlaps.push(name(buttons[i]) + ' / ' + name(buttons[j]) + ' (' + Math.round(width) + 'px)');
  }
  const scaled = [];
  for (const element of [dock, ...dock.querySelectorAll('*')]) {
    const transform = getComputedStyle(element).transform;
    const match = /^matrix\(([^,]+),[^,]+,[^,]+,([^,]+),/.exec(transform);
    if (match && (Math.abs(parseFloat(match[1]) - 1) > 0.01 || Math.abs(parseFloat(match[2]) - 1) > 0.01)) {
      scaled.push((element.className || element.tagName) + ' ' + parseFloat(match[1]).toFixed(2));
    }
  }
  // A button hidden by a clipping ancestor that cannot scroll it into view.
  const unreachable = all.filter((button) => {
    const box = visible(button);
    const own = button.getBoundingClientRect();
    if ((box.right - box.left) >= own.width - 1 && (box.bottom - box.top) >= own.height - 1) return false;
    for (let parent = button.parentElement; parent && parent !== dock.parentElement; parent = parent.parentElement) {
      const style = getComputedStyle(parent);
      if (/(auto|scroll)/.test(style.overflowX) && parent.scrollWidth > parent.clientWidth + 1) return false;
    }
    return true;
  }).map(name);
  const label = dock.querySelector('.classic-bottom-tab-label');
  return {
    // Content-box width: what the container query compares against.
    dockWidth: Math.round(dock.clientWidth - parseFloat(getComputedStyle(dock).paddingLeft) - parseFloat(getComputedStyle(dock).paddingRight)),
    buttons: all.length,
    overlaps,
    scaled,
    unreachable,
    labelVisible: Boolean(label && label.getBoundingClientRect().width > 0),
  };
})()`;
const dockReport = overlapReport(".classic-bottom-tabs");
const headerSelector = ".app-shell:not(.office-dialog-shell) > .app-header, .office-dialog-shell .is-office-editor-header";
const headerReport = overlapReport(headerSelector);

// Counts layout activity while nothing should change: style/class mutations,
// ResizeObserver callbacks on the main regions and moved elements.
const restProbe = (durationMs) => String.raw`new Promise((done) => {
  const selectors = ['.workspace', '.classic-bottom-tabs', '.classic-tile-toolbar', '.formula-tile-list',
    '.template-strip', '.matrix-builder', '.mathfield-host', '.formula-tile-preview', '.math-preview'];
  const watched = [...new Set(selectors.flatMap((selector) => [...document.querySelectorAll(selector)]))].slice(0, 400);
  const key = (element) => (element.className && typeof element.className === 'string'
    ? '.' + element.className.trim().split(/\s+/).slice(0, 2).join('.') : element.tagName);
  const snapshot = () => watched.map((element) => {
    const box = element.getBoundingClientRect();
    return [Math.round(box.left * 2), Math.round(box.top * 2), Math.round(box.width * 2), Math.round(box.height * 2)].join(',');
  });
  const before = snapshot();
  let resizeCallbacks = -1;
  const resized = {};
  const observer = new ResizeObserver((entries) => {
    resizeCallbacks += 1;
    if (resizeCallbacks > 0) for (const entry of entries) resized[key(entry.target)] = (resized[key(entry.target)] || 0) + 1;
  });
  for (const element of watched) observer.observe(element);
  const mutated = {};
  let mutations = 0;
  const mutationObserver = new MutationObserver((records) => {
    for (const record of records) {
      if (record.type !== 'attributes') continue;
      const target = record.target;
      if (!(target instanceof Element) || target.closest('math-field, .ML__container, .cm-editor')) continue;
      mutations += 1;
      const name = key(target) + '@' + record.attributeName;
      mutated[name] = (mutated[name] || 0) + 1;
    }
  });
  mutationObserver.observe(document.body, { subtree: true, attributes: true, attributeFilter: ['style', 'class'] });
  setTimeout(() => {
    observer.disconnect();
    mutationObserver.disconnect();
    const after = snapshot();
    const moved = watched.filter((_, index) => before[index] !== after[index]).map(key);
    const top = (record) => Object.entries(record).sort((a, b) => b[1] - a[1]).slice(0, 4);
    const list = document.querySelector('.classic-tile-toolbar .formula-tile-list');
    done({
      resizeCallbacks: Math.max(0, resizeCallbacks),
      mutations,
      moved: [...new Set(moved)].slice(0, 6),
      topResized: top(resized),
      topMutated: top(mutated),
      tileList: list ? { scrollHeight: list.scrollHeight, clientHeight: list.clientHeight, clientWidth: list.clientWidth } : null,
    });
  }, ${durationMs});
})`;

const menuTrace = (triggerSelector, layerSelector) => String.raw`new Promise((done) => {
  const trigger = document.querySelector(${JSON.stringify(triggerSelector)});
  if (!trigger) return done(null);
  const samples = [];
  const started = performance.now();
  const tick = () => {
    const layer = document.querySelector(${JSON.stringify(layerSelector)});
    if (layer && layer.getClientRects().length) {
      const box = layer.getBoundingClientRect();
      samples.push([Math.round(box.left * 100) / 100, Math.round(box.top * 100) / 100, Math.round(box.width)]);
    }
    if (performance.now() - started < 900) requestAnimationFrame(tick);
    else done({ samples, inset: samples.length ? Math.min(...samples.map((s) => s[0])) : null });
  };
  trigger.click();
  requestAnimationFrame(tick);
})`;

const pickerReport = String.raw`(() => {
  const popover = document.querySelector('[data-formula-color-popover]');
  if (!popover) return null;
  const box = popover.getBoundingClientRect();
  let clipped = false;
  for (let parent = popover.parentElement; parent && parent !== document.body; parent = parent.parentElement) {
    const style = getComputedStyle(parent);
    if (!/(hidden|clip|auto|scroll)/.test(style.overflowX + style.overflowY)) continue;
    const clip = parent.getBoundingClientRect();
    if (box.left < clip.left - 1 || box.right > clip.right + 1 || box.top < clip.top - 1 || box.bottom > clip.bottom + 1) clipped = true;
  }
  const buttons = [...popover.querySelectorAll('button')];
  const hit = buttons.filter((button) => {
    const b = button.getBoundingClientRect();
    const target = document.elementFromPoint(b.left + b.width / 2, b.top + b.height / 2);
    return target && (target === button || button.contains(target));
  }).length;
  return {
    inside: box.left >= 0 && box.right <= innerWidth && box.top >= 0 && box.bottom <= innerHeight,
    clipped,
    hittable: buttons.length ? hit / buttons.length : 0,
  };
})()`;

// ---- Driver --------------------------------------------------------------

const report = { desktopDist, officeDist, desktop: {}, office: {} };
const failures = [];
const check = (condition, message) => {
  if (!condition) failures.push(message);
};

function sweep(from, to, step) {
  const values = [];
  for (let value = from; value >= to; value -= step) values.push(value);
  return values;
}

async function main() {
  await new Promise((done, fail) => {
    server.once("error", fail);
    server.listen(port, "127.0.0.1", done);
  });
  let chrome;
  let client;
  try {
    chrome = spawn(chromePath, [
      "--headless=new",
      "--disable-gpu",
      "--no-first-run",
      "--no-default-browser-check",
      "--hide-scrollbars=false",
      `--remote-debugging-port=${debugPort}`,
      `--user-data-dir=${chromeProfile}`,
      "--window-size=1440,900",
      "about:blank",
    ], { stdio: "ignore" });
    await waitForUrl(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
    const target = targets.find((candidate) => candidate.type === "page");
    client = new CdpClient(target.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");

    const viewport = async (width, height, settle = 260) => {
      await client.send("Emulation.setDeviceMetricsOverride", { width, height, deviceScaleFactor: 1, mobile: false });
      await sleep(settle);
    };
    const waitFor = async (expression, label, timeout = 20_000) => {
      const started = Date.now();
      while (Date.now() - started < timeout) {
        try {
          if (await client.evaluate(expression)) return;
        } catch {}
        await sleep(80);
      }
      const state = await client.evaluate("({ url: location.href, classes: [...new Set([...document.querySelectorAll('[class]')].slice(0, 60).map((e) => String(e.className).split(' ')[0]))].join(' '), text: document.body?.innerText.slice(0, 400) })").catch((error) => String(error));
      throw new Error(`Timed out waiting for ${label}: ${JSON.stringify(state)}`);
    };
    const shot = async (name) => {
      if (!shotsDir) return;
      await mkdir(shotsDir, { recursive: true });
      const { data } = await client.send("Page.captureScreenshot", { format: "png" });
      await writeFile(join(shotsDir, `${name}.png`), Buffer.from(data, "base64"));
    };
    const rest = (ms = 600) => client.evaluate(restProbe(ms));
    const stable = (probe) => probe.resizeCallbacks <= 1 && probe.moved.length === 0 && probe.mutations <= 2;
    const dragTileResizer = async (dx) => {
      const p = await client.evaluate(`(() => { const r = document.querySelector('.classic-tile-resizer')?.getBoundingClientRect(); return r ? { x: r.left + r.width / 2, y: r.top + r.height / 2 } : null; })()`);
      if (!p) return false;
      await client.send("Input.dispatchMouseEvent", { type: "mouseMoved", x: p.x, y: p.y });
      await client.send("Input.dispatchMouseEvent", { type: "mousePressed", x: p.x, y: p.y, button: "left", clickCount: 1 });
      await client.send("Input.dispatchMouseEvent", { type: "mouseMoved", x: p.x + dx, y: p.y, button: "left", buttons: 1 });
      await client.send("Input.dispatchMouseEvent", { type: "mouseReleased", x: p.x + dx, y: p.y, button: "left", clickCount: 1 });
      return true;
    };
    // The tile panel starts collapsed in a fresh profile and collapses itself
    // on narrow windows; open it before measuring anything that involves it.
    const ensureTiles = async () => {
      await client.evaluate(`document.querySelector('[aria-label="展开公式磁贴"], [aria-label="Expand formula tiles"]')?.click(), true`);
      await waitFor("Boolean(document.querySelector('.classic-tile-toolbar .formula-tile-list'))", "tile panel");
      await sleep(300);
    };
    const tileWidth = () => client.evaluate("Math.round(document.querySelector('.classic-tile-toolbar')?.getBoundingClientRect().width ?? 0)");

    if (only !== "office") {
      // ---- Desktop window ----
      distRoot = desktopDist;
      await viewport(1280, 820, 0);
      const { identifier: desktopScript } = await client.send("Page.addScriptToEvaluateOnNewDocument", {
        source: `try { if (location.pathname === '/') { localStorage.setItem('visualtex-editor', JSON.stringify({ state: { editorLayout: 'classic', language: 'zh', sidebarOpen: true, lines: [{ id: 'layout', latex: 'x^2+y^2=z^2' }], activeLineId: 'layout' }, version: 0 })); localStorage.setItem('visualtex-onboarding-completed', 'true'); } } catch {}`,
      });
      await client.send("Page.navigate", { url: baseUrl + "/" });
      await waitFor("Boolean(document.querySelector('.workspace.is-classic-layout .classic-bottom-tabs') && document.querySelector('math-field'))", "desktop classic dock");
      await client.evaluate("document.querySelectorAll('.onboarding-backdrop').forEach((node) => node.remove()), true");
      await sleep(600);

      await ensureTiles();
      const dockWidths = scan ? sweep(1440, 500, 4) : [1440, 1280, 1100, 1000, 940, 900, 860, 800, 760, 700, 640, 600, 560, 520, 500];
      report.desktop.dock = [];
      for (const width of dockWidths) {
        await viewport(width, 820);
        const dock = await client.evaluate(dockReport);
        const header = await client.evaluate(headerReport);
        const row = { width, ...dock, header: header && { overlaps: header.overlaps, scaled: header.scaled } };
        check(!header || header.overlaps.length === 0, `desktop ${width}px: header buttons overlap: ${header?.overlaps.join(", ")}`);
        if (!scan || width % 20 === 0) row.rest = await rest(500);
        report.desktop.dock.push(row);
        check(dock.overlaps.length === 0, `desktop ${width}px: dock buttons overlap: ${dock.overlaps.join(", ")}`);
        check(dock.scaled.length === 0, `desktop ${width}px: dock content shrunk by transform: ${dock.scaled.join(", ")}`);
        check(dock.unreachable.length === 0, `desktop ${width}px: dock buttons cut off: ${dock.unreachable.join(", ")}`);
        check(dock.labelVisible === dock.dockWidth >= 940, `desktop ${width}px: tab labels should show only on a dock of 940px or more (dock ${dock.dockWidth}px)`);
        if (row.rest) check(stable(row.rest), `desktop ${width}px: layout still changing at rest ${JSON.stringify(row.rest)}`);
        if ([1280, 1000, 900, 800, 700, 600, 520].includes(width)) await shot(`desktop-${width}`);
      }

      // Menus: the edges must not move once the menu is on screen, also while
      // the pointer moves over its items (hover transitions).
      report.desktop.menus = [];
      for (const width of [1280, 900, 600]) {
        await viewport(width, 820);
        for (const hover of [false, true]) {
          const tracing = client.evaluate(menuTrace(".menu-button", ".app-menu-popover"));
          if (hover) {
            // A real pointer moving over the items (hover transitions).
            await sleep(180);
            const points = await client.evaluate("[...document.querySelectorAll('.app-menu-popover > button')].slice(0, 8).map((b) => { const r = b.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })");
            for (const point of points) {
              await client.send("Input.dispatchMouseEvent", { type: "mouseMoved", ...point });
              await sleep(70);
            }
          }
          const trace = await tracing;
          const lefts = trace?.samples.map((s) => s[0]) ?? [];
          const tops = trace?.samples.map((s) => s[1]) ?? [];
          const drift = lefts.length ? Math.max(...lefts) - Math.min(...lefts) : null;
          const driftY = tops.length ? Math.max(...tops) - Math.min(...tops) : null;
          report.desktop.menus.push({ width, hover, frames: lefts.length, drift, driftY, inset: trace?.inset, first: trace?.samples[0], last: trace?.samples.at(-1) });
          check(lefts.length > 3, `desktop ${width}px: main menu did not open`);
          check(drift !== null && drift <= 0.5, `desktop ${width}px${hover ? " (hover)" : ""}: main menu moved ${drift?.toFixed(2)}px horizontally after opening`);
          check(driftY !== null && driftY <= 0.5, `desktop ${width}px${hover ? " (hover)" : ""}: main menu moved ${driftY?.toFixed(2)}px vertically after opening`);
          check(trace?.inset >= 9.5, `desktop ${width}px: main menu leaves the 10px safe inset (${trace?.inset})`);
          await client.evaluate("document.querySelector('.menu-button').click(), true");
          await sleep(200);
        }
      }

      // Color picker escapes the scrolling strip and stays on screen.
      report.desktop.picker = [];
      for (const width of [1280, 700, 560]) {
        await viewport(width, 820);
        await client.evaluate("document.querySelector('[data-formula-selection-color]')?.click(), true");
        await sleep(250);
        const picker = await client.evaluate(pickerReport);
        report.desktop.picker.push({ width, ...picker });
        check(picker && picker.inside && !picker.clipped && picker.hittable === 1, `desktop ${width}px: color picker clipped or off screen ${JSON.stringify(picker)}`);
        if (width === 700) await shot("desktop-700-color");
        await client.evaluate("document.querySelector('[data-formula-selection-color]')?.click(), true");
        await sleep(150);
      }

      // Tile panel at rest across window widths, window heights and panel widths.
      await viewport(1440, 820);
      await ensureTiles();
      const tileRuns = [];
      const recordTiles = async (axis, value) => {
        const probe = await rest(scan ? 450 : 600);
        tileRuns.push({ axis, value, ...probe });
        check(stable(probe), `desktop tiles ${axis}=${value}: layout still changing at rest ${JSON.stringify(probe)}`);
      };
      for (const width of scan ? sweep(1440, 640, 3) : [1440, 1200, 1000, 860, 760, 700]) {
        await viewport(width, 820, 220);
        await recordTiles("windowWidth", width);
      }
      await viewport(1280, 820);
      await ensureTiles();
      for (const height of scan ? sweep(1000, 460, 3) : [1000, 820, 700, 600, 520]) {
        await viewport(1280, height, 220);
        await recordTiles("windowHeight", height);
      }
      await viewport(1440, 820);
      await ensureTiles();
      await client.evaluate("document.querySelector('.classic-tile-resizer')?.focus(), true");
      await client.send("Input.dispatchKeyEvent", { type: "keyDown", key: "Home", code: "Home", windowsVirtualKeyCode: 36 });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", key: "Home", code: "Home", windowsVirtualKeyCode: 36 });
      await sleep(300);
      for (let step = 0; step < (scan ? 140 : 12); step += 1) {
        await dragTileResizer(scan ? -3 : -24);
        await sleep(200);
        await recordTiles("tileWidth", await tileWidth());
      }
      report.desktop.tiles = tileRuns;
      await client.send("Page.removeScriptToEvaluateOnNewDocument", { identifier: desktopScript });
    }

    if (only !== "desktop") {
      // ---- Office editor window ----
      distRoot = officeDist;
      await viewport(900, 620, 0);
      // The Office editor runs in a Tauri window; give it the minimum of that
      // runtime it touches while rendering (window label, no-op commands).
      await client.send("Page.addScriptToEvaluateOnNewDocument", {
        source: `if (location.pathname.startsWith('/dialog/')) {
          let callback = 1;
          window.__TAURI_INTERNALS__ = {
            metadata: { currentWindow: { label: 'office-editor' }, currentWebview: { label: 'office-editor', windowLabel: 'office-editor' } },
            invoke: async () => null,
            transformCallback: (fn) => { const id = callback++; window['_' + id] = fn; return id; },
            unregisterCallback: () => {},
            convertFileSrc: (path) => path,
          };
          window.__TAURI_EVENT_PLUGIN_INTERNALS__ = { unregisterListener: () => {} };
        }`,
      });
      await client.send("Page.navigate", { url: `${baseUrl}/dialog/index.html?sessionId=${sessionId}&runtime=vsto-desktop` });
      await waitFor("Boolean(document.querySelector('.office-dialog-shell .classic-bottom-tabs') && document.querySelector('math-field'))", "Office dock");
      await sleep(800);
      // The tile rail starts collapsed in Office; open it so the canvas/rail
      // split is measured too.
      await client.evaluate(`document.querySelector('[aria-label="展开公式磁贴"], [aria-label="Expand formula tiles"]')?.click(), true`);
      await waitFor("Boolean(document.querySelector('.office-dialog-shell .classic-tile-toolbar'))", "Office tile rail");
      await sleep(400);
      report.office.dock = [];
      for (const width of scan ? sweep(1100, 480, 4) : [1100, 1000, 900, 800, 760, 700, 640, 600, 560, 520, 500]) {
        await viewport(width, 620);
        const dock = await client.evaluate(dockReport);
        const header = await client.evaluate(headerReport);
        // Visible canvas: the editor pane up to where the tile rail starts.
        const sizes = await client.evaluate(`(() => {
          const pane = document.querySelector('.office-dialog-shell .formula-workspace.editor-pane')?.getBoundingClientRect();
          const rail = document.querySelector('.office-dialog-shell .classic-tile-toolbar')?.getBoundingClientRect();
          if (!pane) return { editor: 0, tiles: 0 };
          const right = rail && rail.width > 0 && rail.left > pane.left ? Math.min(pane.right, rail.left) : pane.right;
          return { editor: Math.round(right - pane.left), tiles: Math.round(rail?.width ?? 0) };
        })()`);
        const row = { width, ...dock, ...sizes, header: header && { overlaps: header.overlaps, scaled: header.scaled } };
        check(!header || header.overlaps.length === 0, `Office ${width}px: header buttons overlap: ${header?.overlaps.join(", ")}`);
        if (!scan || width % 20 === 0) row.rest = await rest(500);
        report.office.dock.push(row);
        check(dock.overlaps.length === 0, `Office ${width}px: dock buttons overlap: ${dock.overlaps.join(", ")}`);
        check(dock.scaled.length === 0, `Office ${width}px: dock content shrunk by transform: ${dock.scaled.join(", ")}`);
        check(dock.unreachable.length === 0, `Office ${width}px: dock buttons cut off: ${dock.unreachable.join(", ")}`);
        check(!sizes.tiles || sizes.editor >= sizes.tiles * 0.9, `Office ${width}px: canvas (${sizes.editor}px) squeezed below the tile rail (${sizes.tiles}px)`);
        if (row.rest) check(stable(row.rest), `Office ${width}px: layout still changing at rest ${JSON.stringify(row.rest)}`);
        if ([1100, 900, 760, 700, 600, 520].includes(width)) await shot(`office-${width}`);
      }
      // A rail the user widened in a large window must give way again when
      // the window shrinks: the canvas stays at least as wide as the rail.
      report.office.widenedRail = [];
      await viewport(1280, 620);
      await dragTileResizer(-700);
      await sleep(300);
      for (const width of [1280, 1100, 1000, 900, 800, 700, 640]) {
        await viewport(width, 620);
        const sizes = await client.evaluate(`(() => {
          const pane = document.querySelector('.office-dialog-shell .formula-workspace.editor-pane')?.getBoundingClientRect();
          const rail = document.querySelector('.office-dialog-shell .classic-tile-toolbar')?.getBoundingClientRect();
          const right = rail && rail.left > pane.left ? Math.min(pane.right, rail.left) : pane.right;
          return { editor: Math.round(right - pane.left), tiles: Math.round(rail?.width ?? 0) };
        })()`);
        report.office.widenedRail.push({ width, ...sizes });
        check(sizes.editor >= sizes.tiles * 0.9, `Office ${width}px with a widened rail: canvas (${sizes.editor}px) squeezed below the tile rail (${sizes.tiles}px)`);
      }
      await viewport(700, 620);
      await client.evaluate("document.querySelector('[data-formula-selection-color]')?.click(), true");
      await sleep(250);
      report.office.picker = await client.evaluate(pickerReport);
      check(report.office.picker && report.office.picker.inside && !report.office.picker.clipped, `Office 700px: color picker clipped or off screen ${JSON.stringify(report.office.picker)}`);
      await shot("office-700-color");
    }

    await client.send("Browser.close").catch(() => {});
  } finally {
    client?.close();
    if (chrome && !chrome.killed) chrome.kill();
    await new Promise((done) => server.close(done));
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => {});
  }

  report.failures = failures;
  if (reportPath) await writeFile(reportPath, JSON.stringify(report, null, 1));
  const unique = [...new Set(failures.map((failure) => failure.replace(/\{.*$/, "").trim()))];
  console.log(`${failures.length} problem(s)` + (unique.length ? ":\n  " + unique.slice(0, 80).join("\n  ") : ""));
  if (!scan) {
    assert.deepEqual(failures, [], "Windows responsive layout regression failed");
    console.log("Windows responsive layout regression passed.");
  }
}

await main();
