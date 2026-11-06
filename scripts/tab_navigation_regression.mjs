// Real-browser regression for Tab in the visual editor: local, one stop per
// press, never selecting or replacing a structure.
//
//   npm run build:desktop && node scripts/tab_navigation_regression.mjs [--verbose]
//
// Auto-exit after one character is turned off for scripts, accents and font
// scopes, so every exit below is made by Tab itself.
import { spawn } from "node:child_process";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const offset = process.pid % 1000;
const previewPort = 19200 + offset;
const debugPort = 24200 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const verbose = process.argv.includes("--verbose");

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
      const pending = message.id && this.pending.get(message.id);
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
}

async function waitFor(url, timeoutMs = 20_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      if ((await fetch(url)).ok) return;
    } catch {
      // still starting
    }
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

// key, code, virtual key, implied modifiers (8 = Shift)
const PRINTABLE = {
  "^": ["^", "Digit6", 54, 8], "_": ["_", "Minus", 189, 8], "(": ["(", "Digit9", 57, 8],
  ")": [")", "Digit0", 48, 8], "+": ["+", "Equal", 187, 8], " ": [" ", "Space", 32, 0],
  "\\": ["\\", "Backslash", 220, 0], "&": ["&", "Digit7", 55, 8],
};
const NAMED = { Tab: 9, Enter: 13, Escape: 27, ArrowRight: 39, ArrowLeft: 37 };
const SHIFT = 8;
const CTRL = 2;

const preview = spawn(
  process.execPath,
  ["node_modules/vite/bin/vite.js", "preview", "--host", "127.0.0.1", "--port", String(previewPort), "--strictPort"],
  { cwd: process.cwd(), stdio: "ignore" },
);
let chrome;
const results = [];
const check = (name, ok, detail) => {
  results.push({ name, ok });
  console.log(`${ok ? "PASS" : "FAIL"} ${name}${ok && !verbose ? "" : " " + JSON.stringify(detail)}`);
};

try {
  await waitFor(baseUrl);
  chrome = spawn(
    resolveChromiumExecutable(),
    [
      "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
      `--remote-debugging-port=${debugPort}`,
      `--user-data-dir=${createBrowserProfilePath("visualtex-tab-navigation")}`,
      "--window-size=1400,1000", baseUrl,
    ],
    { stdio: "ignore" },
  );
  await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
  const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
  const page = targets.find((t) => t.type === "page" && t.url.startsWith(baseUrl)) ?? targets.find((t) => t.type === "page");
  const client = new CdpClient(page.webSocketDebuggerUrl);
  await client.connect();
  await client.send("Runtime.enable");

  const evaluate = async (expression) => {
    const result = await client.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
    if (result.exceptionDetails)
      throw new Error(result.exceptionDetails.exception?.description ?? result.exceptionDetails.text);
    return result.result.value;
  };
  const press = async (name, modifiers = 0) => {
    let key, code, vk, mods = modifiers;
    if (NAMED[name]) [key, code, vk] = [name, name, NAMED[name]];
    else if (PRINTABLE[name]) {
      [key, code, vk] = PRINTABLE[name];
      mods |= PRINTABLE[name][3];
    } else if (/^[0-9]$/.test(name)) [key, code, vk] = [name, `Digit${name}`, name.charCodeAt(0)];
    else [key, code, vk] = [name, `Key${name.toUpperCase()}`, name.toUpperCase().charCodeAt(0)];
    const printable = key.length === 1 && !(mods & CTRL);
    const common = { key, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, modifiers: mods };
    await client.send("Input.dispatchKeyEvent", {
      type: printable ? "keyDown" : "rawKeyDown",
      ...common,
      ...(printable ? { text: key, unmodifiedText: key } : {}),
    });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
    await sleep(70);
  };
  const type = async (text) => {
    for (const ch of text) await press(ch);
  };
  // "a<Tab>b<S-Tab>c<C-Tab><Esc><Enter>" -> keys
  const keys = async (script) => {
    for (const part of script.split(/(<[^>]+>)/).filter(Boolean)) {
      const named = /^<(S-|C-|CS-)?([A-Za-z]+)>$/.exec(part);
      if (!named) {
        await type(part);
        continue;
      }
      const mods = { "S-": SHIFT, "C-": CTRL, "CS-": CTRL | SHIFT }[named[1]] ?? 0;
      await press(named[2] === "Space" ? " " : named[2], mods);
      await sleep(named[2] === "Space" ? 200 : 60);
    }
    await sleep(150);
  };

  await client.send("Page.enable");
  await client.send("Page.navigate", { url: baseUrl });
  for (let i = 0; i < 100; i++) {
    try {
      if ((await evaluate(`document.readyState === "complete" && location.href.startsWith(${JSON.stringify(baseUrl)})`))) break;
    } catch {
      // the first navigation is still replacing the page
    }
    await sleep(100);
  }
  await sleep(500);
  await evaluate(`(() => {
    localStorage.setItem("visualtex.onboarding.v3.completed", "true");
    const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
    const id = crypto.randomUUID();
    const inputBehavior = { ...((persisted.state || {}).inputBehavior || {}),
      autoExitSuperscript: false, autoExitSubscript: false, autoExitAccent: false, autoExitWrapperCommand: false };
    persisted.state = { ...(persisted.state || {}), lines: [{ id, latex: "" }], activeLineId: id,
      autoPairDelimiters: true, inputBehavior, sourceOpen: false };
    localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
  })()`);
  await client.send("Page.reload", { ignoreCache: true });
  for (let i = 0; i < 300; i++) {
    if (await evaluate(`Boolean(document.querySelector("math-field.visual-mathfield")?.shadowRoot)`)) break;
    await sleep(50);
  }
  await evaluate(`document.fonts.ready.then(() => true)`);
  await sleep(400);

  const FIELD = `document.querySelector("math-field.visual-mathfield")`;
  // Empty field, or `latex` with the caret at `position` (default: start).
  const start = async (latex = "", position = 0) => {
    await evaluate(`(() => { const f = ${FIELD};
      f.setValue(${JSON.stringify(latex)}, { mode: "math", format: "latex", insertionMode: "replaceAll", selectionMode: "after", silenceNotifications: true });
      f.dispatchEvent(new InputEvent("input", { bubbles: true, composed: true, inputType: "insertText" }));
      return true; })()`);
    await sleep(100);
    await evaluate(`(() => { const f = ${FIELD}; f.focus();
      f.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      f.position = ${latex ? position : "f.lastOffset"}; return true; })()`);
    await sleep(80);
  };
  const value = () => evaluate(`${FIELD}.value`);
  const normalize = (latex) => latex.replace(/\s+/g, "");
  const expectValue = async (name, script, expected, setup = ["", 0]) => {
    await start(...setup);
    await keys(script);
    const actual = await value();
    check(name, normalize(actual) === normalize(expected), { actual, expected });
    return actual;
  };

  // 1. Leaving a structure from its last slot, one level per press.
  await expectValue("superscript: Tab leaves it", "x^2<Tab>y", "x^2y");
  await expectValue("subscript: Tab leaves it", "x_1<Tab>y", "x_1y");
  await expectValue("brackets: Tab steps past the closing bracket", "(a+b<Tab>c", "\\left(a+b\\right)c");
  await expectValue("\\frac: Tab numerator -> denominator -> after", "\\frac<Space>a<Tab>b<Tab>c", "\\frac{a}{b}c");
  await expectValue("\\sqrt: Tab from the last slot leaves it", "\\sqrt<Space>a<Tab>c", "\\sqrt{a}c");
  await expectValue("nested: one level per press", "\\frac<Space>\\sqrt<Space>a<Tab>z<Tab>b<Tab>w",
    "\\frac{\\sqrt{a}z}{b}w");
  await expectValue("accent: Tab leaves it", "\\hat<Space>a<Tab>b", "\\hat{a}b");
  await expectValue("nested scripts: e^{x^2} Tab twice", "e^x^2<Tab>a<Tab>b", "\\mathrm{e}^{x^2a}b");

  // 2. Shift+Tab mirrors Tab.
  await expectValue("Shift+Tab: back to the numerator's end", "\\frac<Space>a<Tab>b<S-Tab>z", "\\frac{az}{b}");
  await expectValue("Shift+Tab from the first slot goes before the structure", "\\frac<Space>a<S-Tab>w", "w\\frac{a}{\\placeholder{}}");
  await expectValue("Shift+Tab in brackets goes before the opening bracket", "(a<S-Tab>w", "w\\left(a\\right)");

  // 3. Empty slots ahead are visited before leaving; a body slot right after
  // the structure is the next stop.
  await expectValue("top level: Tab enters the next empty slot", "<Tab>a<Tab>b", "x+\\frac{a}{b}",
    ["x+\\frac{\\placeholder{}}{\\placeholder{}}", 1]);
  await expectValue("\\int: lower, upper, then the body slot", "<Tab>a<Tab>b<Tab>f", "\\int_{a}^{b}f",
    ["\\int_{\\placeholder{}}^{\\placeholder{}}\\placeholder{}", 0]);

  // 4. Nothing to leave: Tab keeps the caret and focus where they are.
  const flat = await expectValue("top level with nothing ahead: Tab does nothing", "x+1<Tab>2", "x+12");
  check("focus stays in the formula after Tab", await evaluate(`document.activeElement === ${FIELD}`), flat);

  // 5. Matrices: Tab walks the cells and adds a row after the last one.
  const matrix = "\\begin{pmatrix}\\placeholder{}&\\placeholder{}\\\\\\placeholder{}&\\placeholder{}\\end{pmatrix}";
  await expectValue("matrix: Tab walks the cells in reading order", "<Tab>a<Tab>b<Tab>c<Tab>d",
    "\\begin{pmatrix}a&b\\\\c&d\\end{pmatrix}", [matrix, 0]);
  const grown = await expectValue("matrix: Tab in the last cell adds a row", "<Tab>a<Tab>b<Tab>c<Tab>d<Tab>e",
    "\\begin{pmatrix}a&b\\\\c&d\\\\e&\\placeholder{}\\end{pmatrix}", [matrix, 0]);
  if (verbose) console.log("grown:", grown);
  await expectValue("matrix: Shift+Tab goes back a cell", "<Tab>a<Tab>b<S-Tab>z",
    "\\begin{pmatrix}az&b\\\\\\placeholder{}&\\placeholder{}\\end{pmatrix}", [matrix, 0]);
  const undone = await (async () => {
    await start(matrix, 0);
    await keys("<Tab>a<Tab>b<Tab>c<Tab>d<Tab>");
    await press("z", CTRL);
    await sleep(250);
    return value();
  })();
  check("matrix: Ctrl+Z removes the row Tab added", normalize(undone) === normalize("\\begin{pmatrix}a&b\\\\c&d\\end{pmatrix}"), undone);
  // & is LaTeX's column separator: next cell, or a new column at row end.
  await expectValue("matrix: & moves to the cell on the right", "<Tab>a&b",
    "\\begin{pmatrix}a&b\\\\\\placeholder{}&\\placeholder{}\\end{pmatrix}", [matrix, 0]);
  await expectValue("matrix: & at the end of a row adds a column", "<Tab>a<Tab>b&z",
    // MathLive's own new column: cells start empty, not as placeholders.
    "\\begin{pmatrix}a&b&z\\\\\\placeholder{}&\\placeholder{}&\\end{pmatrix}", [matrix, 0]);
  await expectValue("outside a matrix & stays a character", "x&y", "x\\&y");
  await expectValue("cases: Tab in the last cell adds a row",
    "<Tab>a<Tab>b<Tab>c", "\\begin{cases}a&b\\\\c&\\placeholder{}\\end{cases}",
    ["\\begin{cases}\\placeholder{}&\\placeholder{}\\end{cases}", 0]);

  // 6. Tab while typing a command commits the highlighted candidate exactly
  // like Space.
  const bySpace = await (async () => {
    await start();
    await type("\\fr");
    await sleep(350);
    await press(" ");
    await sleep(250);
    return value();
  })();
  const byTab = await (async () => {
    await start();
    await type("\\fr");
    await sleep(350);
    await press("Tab");
    await sleep(250);
    return value();
  })();
  check("candidate list: Tab commits like Space", byTab === bySpace && byTab.includes("\\frac"), { byTab, bySpace });

  // 7. Font scope: Tab confirms it like Enter, then typing continues outside.
  await start();
  await keys("\\mathbb<Space>RZ<Enter>x");
  const byEnter = await value();
  const scopeByTab = await expectValue("font scope: Tab confirms like Enter", "\\mathbb<Space>RZ<Tab>x", byEnter);
  check("font scope: typing after Tab is outside the scope", /^\\mathbb\{RZ\}/.test(scopeByTab) && !scopeByTab.includes("RZx"), scopeByTab);

  // 8. The way out: Escape opens an empty LaTeX box; Tab leaves the editor.
  await start();
  await type("x");
  await keys("<Escape><Tab>");
  const leftForward = await evaluate(`(() => { const a = document.activeElement;
    return { tag: a?.tagName ?? null, inFormula: Boolean(a?.closest("math-field")), value: ${FIELD}.value, mode: ${FIELD}.mode }; })()`);
  check("Escape, Tab: focus leaves the formula, content untouched",
    leftForward.tag && !leftForward.inFormula && leftForward.value === "x" && leftForward.mode !== "latex", leftForward);
  await start();
  await type("y");
  await keys("<Escape><S-Tab>");
  const leftBackward = await evaluate(`(() => { const a = document.activeElement;
    return { tag: a?.tagName ?? null, inFormula: Boolean(a?.closest("math-field")), value: ${FIELD}.value }; })()`);
  check("Escape, Shift+Tab: focus leaves backwards", leftBackward.tag && !leftBackward.inFormula && leftBackward.value === "y", leftBackward);
} finally {
  chrome?.kill();
  preview.kill();
}

const failed = results.filter((result) => !result.ok);
console.log(`${results.length - failed.length}/${results.length} passed`);
if (failed.length) process.exitCode = 1;
