// Editing-core harness (docs/editing-core/DESIGN.md §10-11). Drives a running
// VisualTeX editor with real key and mouse events and records where the caret
// goes, so every refactor phase can be checked against the previous build.
//
//   node scripts/editing_core_harness.mjs run --out new.json [--url URL] [--corpus FILE] [--only walk,tab,...] [--filter TEXT] [--letter-font times]
//   node scripts/editing_core_harness.mjs check new.json
//   node scripts/editing_core_harness.mjs compare old.json new.json
//
// Without --url it serves the built dist with vite preview, like the other
// browser regressions. Tests: walk (ArrowRight/ArrowLeft through the formula),
// tab (Tab and Shift+Tab from every stop, then typing after Tab), vertical
// (ArrowUp/ArrowDown from every stop), type (a letter at every stop), delete
// (Backspace and Delete from every stop), click (on the caret, else beside it),
// placeholder (caret against the empty-slot marker, measured in pixels) and
// convert (the caret after input conversions, a fixed list of cases).
import { spawn } from "node:child_process";
import { readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { inflateSync } from "node:zlib";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const here = dirname(fileURLToPath(import.meta.url));
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const ALL_TESTS = ["walk", "tab", "vertical", "type", "delete", "click", "placeholder", "convert"];

// The caret after an input conversion (DESIGN §6.10): [formula, start offset
// (-1: end), typed text, value afterwards, the atom left of the caret]. Typing
// before an unconverted "cos" converts it too; the caret must stay where the
// typing happened.
const CONVERSIONS = [
  ["x+", -1, "sin ", "x+\\sin", "operator sin"],
  ["+y", 0, "sin ", "\\sin+y", "operator sin"],
  ["a+cos", 1, "+sin ", "a+\\sin+\\cos", "operator sin"],
  ["a+cos", 1, "+sin", "a+sin+\\cos", "mord n"],
  ["", 0, "sin x", "\\sin x", "mord x"],
];

function option(args, name, fallback = null) {
  const index = args.indexOf(name);
  return index >= 0 ? args[index + 1] : fallback;
}

// ------------------------------------------------------------------ browser --

class CdpClient {
  constructor(url) {
    this.url = url;
    this.nextId = 1;
    this.pending = new Map();
  }

  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((resolve, reject) => {
      this.socket.addEventListener("open", resolve, { once: true });
      this.socket.addEventListener("error", reject, { once: true });
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
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }

  close() {
    this.socket?.close();
  }
}

async function waitFor(url, timeoutMs = 20000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while the local process starts.
    }
    await sleep(100);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

// In-page helpers. Stops are identified by the path of the atom before them
// (branch:index from the root), which is the same in every build of the model.
const PAGE_LIB = `(() => {
  if (window.__vtxHarness) return true;
  const field = () => document.querySelector("math-field");
  const model = () => field()._mathfield.model;
  const branchName = (b) => (Array.isArray(b) ? b.join(",") : String(b));
  const pathOf = (atom) => {
    const parts = [];
    for (let a = atom; a && a.parent; a = a.parent) {
      const siblings = a.parent.branch ? a.parent.branch(a.parentBranch) : null;
      parts.unshift(branchName(a.parentBranch) + ":" + (siblings ? siblings.indexOf(a) : -1));
    }
    return parts.join("/");
  };
  const describe = (atom) =>
    atom
      ? {
          type: atom.type,
          value: typeof atom.value === "string" ? atom.value : null,
          path: pathOf(atom),
          slot: atom.parent ? pathOf(atom.parent) + "|" + branchName(atom.parentBranch) : "",
        }
      : null;
  const origin = () => field().shadowRoot.querySelector('[part="content"]').getBoundingClientRect();
  const caretElement = () =>
    field().shadowRoot.querySelector(".ML__caret, .ML__text-caret, .ML__latex-caret");
  const frames = () =>
    new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
  const round = (v) => Math.round(v * 100) / 100;
  window.__vtxHarness = {
    frames,
    async load(latex, position) {
      const m = field();
      m.setValue(latex);
      m.focus();
      m.position = position === -1 ? m.lastOffset : position;
      await frames();
      return m.lastOffset;
    },
    state() {
      const m = field();
      const o = origin();
      const c = caretElement();
      let caret = null;
      if (c) {
        const r = c.getBoundingClientRect();
        caret = {
          cls: c.className,
          x: round(r.left - o.left),
          y: round(r.top - o.top),
          h: round(r.height),
          w: round(r.width),
          clientX: round(r.left),
          clientY: round(r.top + r.height / 2),
        };
      }
      return {
        pos: m.position,
        sel: m.selection.ranges.map((r) => [r[0], r[1]]),
        mode: m.mode,
        value: m.value,
        at: describe(model().at(m.position)),
        caret,
      };
    },
    offsets() {
      const out = [];
      for (let i = 0; i <= field().lastOffset; i++) out.push(describe(model().at(i)));
      return out;
    },
    placeholderSnapshot() {
      const sr = field().shadowRoot;
      const ph =
        sr.querySelector(".visualtex-structural-placeholder.ML__selected, .ML__placeholder.ML__selected") ??
        sr.querySelector(".ML__placeholder-selected") ??
        [...sr.querySelectorAll(".visualtex-structural-placeholder, .ML__placeholder")].find((e) =>
          e.querySelector(".ML__caret"),
        ) ??
        null;
      const c = caretElement();
      if (!ph || !c) return null;
      // the marker: the engine's own element (or, in older builds, the
      // placeholder's ::before); the bar: the caret's ::after background or border
      const marker = sr.querySelector(".ML__vtx-marker--selected");
      const after = getComputedStyle(c, "::after");
      const caretColor = parseFloat(after.borderRightWidth) > 0 ? after.borderRightColor : after.backgroundColor;
      const markerColor = marker ? getComputedStyle(marker).backgroundColor : getComputedStyle(ph, "::before").backgroundColor;
      const r = ph.getBoundingClientRect();
      const fs = parseFloat(getComputedStyle(ph).fontSize);
      // The marker is centred on the box; the caret sits at its left. A
      // neighbouring slot's marker (a fraction's other half) lies just
      // outside this clip.
      return {
        clip: { x: r.left - 0.6 * fs, y: r.top + r.height / 2 - 0.58 * fs, width: r.width + 0.8 * fs, height: 1.16 * fs },
        caretColor,
        markerColor,
        fontSize: fs,
      };
    },
    executeCommand(command) {
      return field().executeCommand(command);
    },
  };
  return true;
})()`;

const NO_BLINK = `(() => {
  const s = document.createElement("style");
  s.textContent = ".ML__caret::after,.ML__text-caret::after,.ML__latex-caret::after{animation:none!important;visibility:visible!important;opacity:1!important}";
  document.querySelector("math-field").shadowRoot.append(s);
  return true;
})()`;

const KEYS = {
  ArrowRight: ["ArrowRight", 39],
  ArrowLeft: ["ArrowLeft", 37],
  ArrowUp: ["ArrowUp", 38],
  ArrowDown: ["ArrowDown", 40],
  Tab: ["Tab", 9],
  Home: ["Home", 36],
  End: ["End", 35],
  Backspace: ["Backspace", 8],
  Delete: ["Delete", 46],
};
const SHIFT = 8;

class Editor {
  constructor(client) {
    this.client = client;
  }

  async evaluate(expression) {
    const result = await this.client.send("Runtime.evaluate", {
      expression,
      awaitPromise: true,
      returnByValue: true,
    });
    if (result.exceptionDetails) {
      throw new Error(
        result.exceptionDetails.exception?.description ||
          result.exceptionDetails.text ||
          "Runtime.evaluate failed",
      );
    }
    return result.result.value;
  }

  async settle(ms = 25) {
    await this.evaluate("__vtxHarness.frames()");
    if (ms) await sleep(ms);
  }

  async press(key, modifiers = 0) {
    const [code, vk] = KEYS[key];
    const base = { key, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, modifiers };
    await this.client.send("Input.dispatchKeyEvent", { type: "rawKeyDown", ...base });
    await this.client.send("Input.dispatchKeyEvent", { type: "keyUp", ...base });
    await this.settle();
  }

  async typeText(text) {
    for (const ch of text) {
      if (ch !== " ") await this.type(ch);
      else {
        const base = { key: " ", code: "Space", windowsVirtualKeyCode: 32, nativeVirtualKeyCode: 32 };
        await this.client.send("Input.dispatchKeyEvent", { type: "keyDown", ...base, text: " ", unmodifiedText: " " });
        await this.client.send("Input.dispatchKeyEvent", { type: "keyUp", ...base });
        await this.settle(40);
      }
    }
  }

  async type(ch) {
    const vk = ch.toUpperCase().charCodeAt(0);
    const base = { key: ch, code: `Key${ch.toUpperCase()}`, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk };
    await this.client.send("Input.dispatchKeyEvent", { type: "keyDown", ...base, text: ch, unmodifiedText: ch });
    await this.client.send("Input.dispatchKeyEvent", { type: "keyUp", ...base });
    await this.settle(40);
  }

  async click(x, y) {
    // a click near the previous one within the double-click time would count
    // as a double (triple) click and select instead of placing the caret
    const last = this.lastClick;
    if (last && Math.hypot(x - last.x, y - last.y) < 8) await sleep(Math.max(0, last.time + 550 - Date.now()));
    this.lastClick = { x, y, time: Date.now() };
    await this.client.send("Input.dispatchMouseEvent", { type: "mouseMoved", x, y });
    await this.client.send("Input.dispatchMouseEvent", { type: "mousePressed", x, y, button: "left", clickCount: 1 });
    await this.client.send("Input.dispatchMouseEvent", { type: "mouseReleased", x, y, button: "left", clickCount: 1 });
    await this.settle(40);
  }

  async load(latex, position) {
    const last = await this.evaluate(`__vtxHarness.load(${JSON.stringify(latex)}, ${position})`);
    await sleep(15);
    return last;
  }

  state() {
    return this.evaluate("__vtxHarness.state()");
  }

  async screenshot(clip) {
    const { data } = await this.client.send("Page.captureScreenshot", {
      format: "png",
      clip: { ...clip, scale: 1 },
      captureBeyondViewport: false,
    });
    return Buffer.from(data, "base64");
  }
}

// ---------------------------------------------------------------- png/pixels --

function decodePng(buffer) {
  let offset = 8;
  let width = 0;
  let height = 0;
  let colorType = 0;
  const chunks = [];
  while (offset < buffer.length) {
    const length = buffer.readUInt32BE(offset);
    const type = buffer.toString("ascii", offset + 4, offset + 8);
    const data = buffer.subarray(offset + 8, offset + 8 + length);
    if (type === "IHDR") {
      width = data.readUInt32BE(0);
      height = data.readUInt32BE(4);
      colorType = data[9];
    } else if (type === "IDAT") chunks.push(data);
    else if (type === "IEND") break;
    offset += 12 + length;
  }
  const bpp = colorType === 6 ? 4 : colorType === 2 ? 3 : 0;
  if (!bpp) throw new Error(`Unsupported PNG colour type ${colorType}`);
  const raw = inflateSync(Buffer.concat(chunks));
  const stride = width * bpp;
  const pixels = Buffer.alloc(height * stride);
  for (let y = 0; y < height; y++) {
    const filter = raw[y * (stride + 1)];
    const line = raw.subarray(y * (stride + 1) + 1, (y + 1) * (stride + 1));
    const current = pixels.subarray(y * stride, (y + 1) * stride);
    const previous = y > 0 ? pixels.subarray((y - 1) * stride, y * stride) : null;
    for (let x = 0; x < stride; x++) {
      const a = x >= bpp ? current[x - bpp] : 0;
      const b = previous ? previous[x] : 0;
      const c = previous && x >= bpp ? previous[x - bpp] : 0;
      let value = line[x];
      if (filter === 1) value += a;
      else if (filter === 2) value += b;
      else if (filter === 3) value += (a + b) >> 1;
      else if (filter === 4) {
        const p = a + b - c;
        const pa = Math.abs(p - a);
        const pb = Math.abs(p - b);
        const pc = Math.abs(p - c);
        value += pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
      }
      current[x] = value & 255;
    }
  }
  return { width, height, bpp, pixels };
}

function parseRgb(css) {
  const m = /rgba?\(([^)]+)\)/.exec(css ?? "");
  if (!m) return null;
  const [r, g, b] = m[1].split(",").map((v) => parseFloat(v));
  return [r, g, b];
}

// Bounding box of the largest connected patch of one colour (other patches of
// a similar colour, such as a hover background, are ignored).
function colourBox(image, rgb, tolerance) {
  const { width, height } = image;
  const matches = new Uint8Array(width * height);
  for (let p = 0; p < width * height; p++) {
    const i = p * image.bpp;
    matches[p] =
      Math.abs(image.pixels[i] - rgb[0]) <= tolerance &&
      Math.abs(image.pixels[i + 1] - rgb[1]) <= tolerance &&
      Math.abs(image.pixels[i + 2] - rgb[2]) <= tolerance
        ? 1
        : 0;
  }
  let best = null;
  const seen = new Uint8Array(width * height);
  for (let start = 0; start < width * height; start++) {
    if (!matches[start] || seen[start]) continue;
    const box = { left: Infinity, right: -Infinity, top: Infinity, bottom: -Infinity, size: 0 };
    const stack = [start];
    seen[start] = 1;
    while (stack.length) {
      const p = stack.pop();
      const x = p % width;
      const y = (p - x) / width;
      box.left = Math.min(box.left, x);
      box.right = Math.max(box.right, x);
      box.top = Math.min(box.top, y);
      box.bottom = Math.max(box.bottom, y);
      box.size++;
      for (const q of [x > 0 ? p - 1 : -1, x < width - 1 ? p + 1 : -1, y > 0 ? p - width : -1, y < height - 1 ? p + width : -1]) {
        if (q >= 0 && matches[q] && !seen[q]) {
          seen[q] = 1;
          stack.push(q);
        }
      }
    }
    if (!best || box.size > best.size) best = box;
  }
  return best;
}

// --------------------------------------------------------------------- run --

const compact = (state, original) => ({
  pos: state.pos,
  sel: state.sel,
  mode: state.mode,
  at: state.at ? `${state.at.path} ${state.at.type}${state.at.value ? ` ${state.at.value}` : ""}` : null,
  slot: state.at?.slot ?? null,
  caret: state.caret ? { x: state.caret.x, y: state.caret.y, h: state.caret.h, cls: state.caret.cls } : null,
  ...(original !== undefined && state.value !== original ? { value: state.value } : {}),
});

async function walk(editor, latex, key) {
  const last = await editor.load(latex, key === "ArrowRight" ? 0 : -1);
  const states = [await editor.state()];
  for (let i = 0; i < 2 * last + 6; i++) {
    await editor.press(key);
    const state = await editor.state();
    const previous = states[states.length - 1];
    if (state.pos === previous.pos && JSON.stringify(state.sel) === JSON.stringify(previous.sel)) break;
    states.push(state);
  }
  return states.map((s) => compact(s, latex));
}

async function fromEachStop(editor, latex, stops, action) {
  const out = {};
  for (const stop of stops) {
    await editor.load(latex, stop.pos);
    await action();
    out[stop.at] = compact(await editor.state(), latex);
  }
  return out;
}

async function placeholderProbe(editor, latex, shotDir, index) {
  const count = (latex.match(/\\placeholder\{\}/g) ?? []).length;
  if (!count) return null;
  await editor.load(latex, 0);
  const results = [];
  for (let k = 0; k < count; k++) {
    await editor.evaluate("__vtxHarness.executeCommand('moveToNextPlaceholder')");
    await editor.settle(60);
    const state = await editor.state();
    const snap = await editor.evaluate("__vtxHarness.placeholderSnapshot()");
    const entry = { state: compact(state, latex) };
    if (snap) {
      const png = await editor.screenshot(snap.clip);
      if (shotDir) writeFileSync(join(shotDir, `ph-${index}-${k}.png`), png);
      const image = decodePng(png);
      const scale = image.width / snap.clip.width;
      const caretRgb = parseRgb(snap.caretColor);
      const markerRgb = parseRgb(snap.markerColor);
      const caret = caretRgb ? colourBox(image, caretRgb, 40) : null;
      const marker = markerRgb ? colourBox(image, markerRgb, 10) : null;
      if (caret && marker) {
        entry.pixels = {
          gap: Math.round(((marker.left - caret.right - 1) / scale) * 100) / 100,
          topDelta: Math.round(((caret.top - marker.top) / scale) * 100) / 100,
          bottomDelta: Math.round(((caret.bottom - marker.bottom) / scale) * 100) / 100,
          caretHeight: Math.round(((caret.bottom - caret.top + 1) / scale) * 100) / 100,
          markerHeight: Math.round(((marker.bottom - marker.top + 1) / scale) * 100) / 100,
        };
      } else entry.pixels = { missing: !caret ? "caret" : "marker" };
    }
    results.push(entry);
  }
  return results;
}

async function runFormula(editor, latex, tests, shotDir, index) {
  const result = { latex };
  const walkRight = await walk(editor, latex, "ArrowRight");
  const stops = [];
  for (const s of walkRight) if (!stops.some((t) => t.pos === s.pos)) stops.push(s);
  if (tests.includes("walk")) {
    result.walkRight = walkRight;
    result.walkLeft = await walk(editor, latex, "ArrowLeft");
  }
  if (tests.includes("tab")) {
    result.tab = await fromEachStop(editor, latex, stops, () => editor.press("Tab"));
    result.shiftTab = await fromEachStop(editor, latex, stops, () => editor.press("Tab", SHIFT));
    // whatever Tab lands on must keep what is typed there
    result.typeAfterTab = {};
    for (const stop of stops) {
      await editor.load(latex, stop.pos);
      await editor.press("Tab");
      await editor.type("q");
      result.typeAfterTab[stop.at] = (await editor.state()).value;
    }
  }
  if (tests.includes("vertical")) {
    result.up = await fromEachStop(editor, latex, stops, () => editor.press("ArrowUp"));
    result.down = await fromEachStop(editor, latex, stops, () => editor.press("ArrowDown"));
  }
  if (tests.includes("type")) {
    result.type = {};
    for (const stop of stops) {
      await editor.load(latex, stop.pos);
      await editor.type("q");
      const after = await editor.state();
      result.type[stop.at] = { value: after.value, pos: after.pos };
    }
  }
  if (tests.includes("delete")) {
    result.backspace = await fromEachStop(editor, latex, stops, () => editor.press("Backspace"));
    result.forwardDelete = await fromEachStop(editor, latex, stops, () => editor.press("Delete"));
  }
  if (tests.includes("click")) {
    result.click = {};
    for (const stop of stops) {
      if (!stop.caret) continue;
      await editor.load(latex, stop.pos);
      const state = await editor.state();
      if (!state.caret) continue;
      // On the caret's bar, else just left or right of it: a stop whose caret
      // coincides with another one's (the end of a radicand and the end of
      // the radical) is reached from the item beside it.
      const first = state.at?.type === "first";
      const want = stop.at.split(" ").slice(0, 2).join(" ");
      const centre = state.caret.clientX + (state.caret.w ?? 0) / 2;
      let landed = null;
      for (const dx of [0, -1.5, 1.5]) {
        await editor.load(latex, first ? -1 : 0);
        await editor.click(centre + dx, state.caret.clientY);
        const after = await editor.state();
        const at = after.at ? after.at.path + " " + after.at.type : null;
        if (dx === 0) landed = at;
        if (at === want) {
          landed = at;
          break;
        }
      }
      result.click[stop.at] = landed;
    }
  }
  if (tests.includes("placeholder")) {
    const probe = await placeholderProbe(editor, latex, shotDir, index);
    if (probe) result.placeholder = probe;
  }
  return result;
}

async function run(args) {
  const out = option(args, "--out");
  if (!out) throw new Error("run needs --out FILE");
  const corpusFile = option(args, "--corpus", join(here, "editing_core_corpus.json"));
  const tests = (option(args, "--only", ALL_TESTS.join(",")) ?? "").split(",").filter(Boolean);
  const filter = option(args, "--filter");
  const shotDir = option(args, "--shots");
  let url = option(args, "--url");
  // the formula letter font (Settings); "times" lays formulas out in the newtx profile
  const letterFont = option(args, "--letter-font");
  const corpus = JSON.parse(readFileSync(corpusFile, "utf8")).formulas.filter(
    (latex) => !filter || latex.includes(filter),
  );

  const portOffset = process.pid % 1000;
  const debugPort = 12400 + portOffset;
  let preview = null;
  if (!url) {
    const previewPort = 7400 + portOffset;
    url = `http://127.0.0.1:${previewPort}/`;
    preview = spawn(
      process.execPath,
      ["node_modules/vite/bin/vite.js", "preview", "--host", "127.0.0.1", "--port", String(previewPort), "--strictPort"],
      { cwd: join(here, ".."), stdio: "ignore" },
    );
  }
  const profile = createBrowserProfilePath("visualtex-editing-core-harness");
  let chrome = null;
  let client = null;
  const results = { url, letterFont: letterFont ?? null, date: new Date().toISOString(), tests, formulas: [] };
  try {
    await waitFor(url);
    chrome = spawn(
      resolveChromiumExecutable(),
      [
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        `--remote-debugging-port=${debugPort}`,
        `--user-data-dir=${profile}`,
        "--window-size=1400,1000",
        "about:blank",
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
    const page = targets.find((t) => t.type === "page");
    if (!page) throw new Error("No Chrome page target found");
    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Emulation.setDeviceMetricsOverride", {
      width: 1400,
      height: 1000,
      deviceScaleFactor: 2,
      mobile: false,
    });
    await client.send("Page.navigate", { url });
    const editor = new Editor(client);
    const started = Date.now();
    while (true) {
      const ready = await editor
        .evaluate("Boolean(document.querySelector('math-field')?._mathfield)")
        .catch(() => false);
      if (ready) break;
      if (Date.now() - started > 30000) throw new Error("The editor did not start");
      await sleep(200);
    }
    if (letterFont) {
      await editor.evaluate(
        `localStorage.setItem("visualtex.formula-letter-font", ${JSON.stringify(letterFont)}); window.dispatchEvent(new CustomEvent("visualtex-formula-font-preferences-changed"))`,
      );
    }
    await sleep(1500);
    await editor.evaluate(PAGE_LIB);
    await editor.evaluate(NO_BLINK);
    const box = await editor.evaluate(
      "(() => { const r = document.querySelector('math-field').getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()",
    );
    await editor.click(box.x, box.y);
    for (const [index, latex] of corpus.entries()) {
      try {
        results.formulas.push(await runFormula(editor, latex, tests, shotDir, index));
      } catch (error) {
        results.formulas.push({ latex, error: String(error?.message ?? error) });
      }
      process.stdout.write(`\r${index + 1}/${corpus.length}`);
    }
    process.stdout.write("\n");
    if (tests.includes("convert")) {
      results.conversions = [];
      for (const [latex, position, typed] of CONVERSIONS) {
        await editor.load(latex, position);
        await sleep(300);
        await editor.typeText(typed);
        const after = await editor.state();
        results.conversions.push({ latex, position, typed, value: after.value, at: after.at ? `${after.at.type} ${after.at.value ?? ""}`.trim() : null });
      }
    }
  } finally {
    client?.close();
    chrome?.kill();
    preview?.kill();
    try {
      rmSync(profile, { recursive: true, force: true });
    } catch {
      // The browser may still hold the profile for a moment.
    }
  }
  writeFileSync(out, JSON.stringify(results, null, 1));
  console.log(`wrote ${out}`);
  report(check(results));
}

// --------------------------------------------------------------- invariants --

function uniquePositions(states) {
  const out = [];
  for (const s of states) if (out[out.length - 1] !== s.at) out.push(s.at);
  return out;
}

function check(results) {
  const findings = [];
  const add = (latex, invariant, detail) => findings.push({ latex, invariant, detail });
  for (const c of results.conversions ?? []) {
    const want = CONVERSIONS.find(([latex, position, typed]) => latex === c.latex && position === c.position && typed === c.typed);
    if (want && (c.value !== want[3] || c.at !== want[4]))
      add(c.latex || "(empty)", "caret after a conversion", `typing "${c.typed}" at ${c.position}: ${c.value} (caret after ${c.at}), want ${want[3]} (after ${want[4]})`);
  }
  for (const f of results.formulas) {
    if (f.error) {
      add(f.latex, "error", f.error);
      continue;
    }
    if (f.walkRight && f.walkLeft) {
      const right = uniquePositions(f.walkRight);
      const left = uniquePositions(f.walkLeft).reverse();
      if (JSON.stringify(right) !== JSON.stringify(left))
        add(f.latex, "I1 walks differ", `right ${right.length} stops, left ${left.length} stops`);
      for (const [states, dir] of [
        [f.walkRight, 1],
        [f.walkLeft, -1],
      ]) {
        for (let i = 1; i < states.length; i++) {
          const a = states[i - 1];
          const b = states[i];
          if (!a.caret || !b.caret) {
            if (!b.caret) add(f.latex, "I3 no caret drawn", `${dir > 0 ? "right" : "left"} at ${b.at}`);
            continue;
          }
          const sameLevel = Math.abs(a.caret.y + a.caret.h - (b.caret.y + b.caret.h)) < 1.5;
          const dx = (b.caret.x - a.caret.x) * dir;
          if (a.slot === b.slot && sameLevel && dx < -0.6)
            add(f.latex, "I2 caret goes back", `${dir > 0 ? "right" : "left"} ${a.at} → ${b.at} (${dx.toFixed(1)}px)`);
          if (Math.abs(b.caret.x - a.caret.x) < 0.6 && Math.abs(b.caret.y - a.caret.y) < 0.6)
            add(f.latex, "I3 press shows no movement", `${dir > 0 ? "right" : "left"} ${a.at} → ${b.at}`);
        }
      }
    }
    if (f.type) {
      for (const [stop, r] of Object.entries(f.type)) {
        if (r.value === f.latex || !r.value.includes("q"))
          add(f.latex, "I5 typed letter lost", `at ${stop} → ${r.value}`);
      }
    }
    if (f.typeAfterTab) {
      for (const [stop, value] of Object.entries(f.typeAfterTab)) {
        if (value === f.latex || !value.includes("q"))
          add(f.latex, "I5 typed letter lost after Tab", `Tab from ${stop} → ${value}`);
      }
    }
    if (f.click) {
      for (const [stop, landed] of Object.entries(f.click)) {
        const want = stop.split(" ").slice(0, 2).join(" ");
        if (landed !== want) add(f.latex, "I6 click lands elsewhere", `${want} → ${landed}`);
      }
    }
    if (f.down || f.up) {
      // a rewrite any key press makes (the input conversion: upright e, i, d)
      // is not Up/Down's doing
      const rewrites = new Set([...(f.walkRight ?? []), ...(f.walkLeft ?? [])].map((s) => s.value).filter(Boolean));
      for (const [dir, map] of [
        ["up", f.up],
        ["down", f.down],
      ]) {
        for (const [stop, r] of Object.entries(map ?? {}))
          if (r.value !== undefined && !rewrites.has(r.value))
            add(f.latex, "Up/Down changed the formula", `${dir} at ${stop} → ${r.value}`);
      }
    }
    for (const p of f.placeholder ?? []) {
      const px = p.pixels;
      if (!px) continue;
      if (px.missing) add(f.latex, "I8 placeholder not measured", `${px.missing} missing at ${p.state.at}`);
      else if (Math.abs(px.gap) > 1 || px.topDelta > 1.5 || px.bottomDelta < -1.5)
        add(
          f.latex,
          "I8 caret not on the marker",
          `${p.state.at}: gap ${px.gap}px, top ${px.topDelta}, bottom ${px.bottomDelta} (caret ${px.caretHeight}, marker ${px.markerHeight})`,
        );
    }
  }
  return findings;
}

function report(findings) {
  const byInvariant = new Map();
  for (const f of findings) byInvariant.set(f.invariant, [...(byInvariant.get(f.invariant) ?? []), f]);
  if (!findings.length) {
    console.log("All invariants hold.");
    return;
  }
  for (const [invariant, list] of [...byInvariant].sort()) {
    console.log(`\n${invariant}: ${list.length}`);
    for (const f of list.slice(0, 40)) console.log(`  ${f.latex}  —  ${f.detail}`);
    if (list.length > 40) console.log(`  … ${list.length - 40} more`);
  }
}

// ------------------------------------------------------------------ compare --

function compare(oldRun, newRun) {
  const byLatex = new Map(oldRun.formulas.map((f) => [f.latex, f]));
  let differences = 0;
  for (const f of newRun.formulas) {
    const o = byLatex.get(f.latex);
    if (!o) continue;
    const lines = [];
    if (o.walkRight && f.walkRight) {
      const a = uniquePositions(o.walkRight).join(" | ");
      const b = uniquePositions(f.walkRight).join(" | ");
      if (a !== b) lines.push(`walk right:\n      old ${a}\n      new ${b}`);
    }
    for (const key of ["tab", "shiftTab", "up", "down", "backspace", "forwardDelete"]) {
      if (!o[key] || !f[key]) continue;
      for (const stop of Object.keys(o[key])) {
        const before = o[key][stop];
        const after = f[key][stop];
        if (!after) continue;
        // a placeholder's internal value (older kernels stored "x") is not behaviour
        const describe = (r) => `${String(r.at).replace(/ placeholder x$/, " placeholder")}${r.value !== undefined ? ` value=${r.value}` : ""}`;
        if (describe(before) !== describe(after))
          lines.push(`${key} from ${stop}: ${describe(before)}  →  ${describe(after)}`);
      }
    }
    if (o.type && f.type) {
      for (const stop of Object.keys(o.type)) {
        if (f.type[stop] && o.type[stop].value !== f.type[stop].value)
          lines.push(`type at ${stop}: ${o.type[stop].value}  →  ${f.type[stop].value}`);
      }
    }
    if (o.click && f.click) {
      for (const stop of Object.keys(o.click)) {
        if (f.click[stop] !== undefined && o.click[stop] !== f.click[stop])
          lines.push(`click at ${stop}: ${o.click[stop]}  →  ${f.click[stop]}`);
      }
    }
    if (lines.length) {
      differences += lines.length;
      console.log(`\n== ${f.latex}`);
      for (const line of lines) console.log(`   ${line}`);
    }
  }
  console.log(`\n${differences} behaviour difference(s).`);
}

// --------------------------------------------------------------------- main --

const [command, ...rest] = process.argv.slice(2);
if (command === "run") await run(rest);
else if (command === "check") report(check(JSON.parse(readFileSync(rest[0], "utf8"))));
else if (command === "compare")
  compare(JSON.parse(readFileSync(rest[0], "utf8")), JSON.parse(readFileSync(rest[1], "utf8")));
else {
  console.error("Usage: editing_core_harness.mjs run --out FILE [--url URL] [--only tests] [--filter text] [--shots dir]");
  console.error("       editing_core_harness.mjs check FILE");
  console.error("       editing_core_harness.mjs compare OLD NEW");
  process.exit(2);
}
