// Real-browser regression for the input judgment (src/editor/inputJudgment.ts,
// kernel part-021): quick escapes confirmed by a key (Space by default) for the
// name right before the caret, Backspace undoing them, automatic upright
// typography, the two switches and the confirm-escape key setting.
//
//   npm run build:desktop && node scripts/input_judgment_regression.mjs [--verbose]
//
// Each case types keys into an empty formula, then leaves the field (which
// finishes the last word) and compares the formula's LaTeX.
import { spawn } from "node:child_process";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const offset = process.pid % 1000;
const previewPort = 19500 + offset;
const debugPort = 24500 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const verbose = process.argv.includes("--verbose");

// Typed keys -> LaTeX after leaving the field. <Space>, <Tab>, <Backspace>,
// <F2> are keys; everything else is typed character by character. Words
// become symbols only with the confirm-escape key (Space by default), and
// only the name right before the caret; upright typography is automatic.
const BOTH_ON = [
  // The confirm-escape key escapes the name right before the caret only.
  ["alpha<Space>", String.raw`\alpha`],
  ["thetaalpha<Space>", String.raw`theta\alpha`],
  ["alphabeta<Space>", String.raw`alpha\beta`],
  ["theta<Space>alpha<Space>", String.raw`\theta\alpha`],
  ["sintheta<Space>", String.raw`\sin\theta`],
  ["2pi<Space>r^2", String.raw`2\pi r^2`],
  ["pi<Space>r", String.raw`\pi r`],
  ["x in<Space>", String.raw`x\ \in`],
  ["infty<Space>", String.raw`\infty`],
  ["5 deg<Space>", String.raw`5\ \degree`],
  ["int<Space>", String.raw`\int`],
  ["int<Space>xdx", String.raw`\int x\mathrm{d}x`],
  ["sum<Space>_i", String.raw`\sum_{i}`],
  // Without the key nothing escapes; other keys never escape.
  ["alpha", "alpha"],
  ["alpha+", "alpha+"],
  ["thetaalpha", "thetaalpha"],
  ["frac<Space>alpha<Tab>b", String.raw`\frac{alpha}{b}`],
  ["frac<Space>alpha<Space><Tab>b", String.raw`\frac{\alpha}{b}`],
  ["spin", "spin"],
  ["point", "point"],
  ["for", "for"],
  // Function names are set upright, not escaped; a Space after one is used up.
  ["sin<Space>x", String.raw`\sin x`],
  ["sinx", String.raw`\sin x`],
  ["lnx", String.raw`\ln x`],
  ["sup", String.raw`\sup`],
  ["int<Space>sinxdx", String.raw`\int \sin x\mathrm{d}x`],
  // Templates: the caret goes into the first slot.
  ["frac<Space>a<Tab>b", String.raw`\frac{a}{b}`],
  ["sqrt<Space>x", String.raw`\sqrt{x}`],
  ["hat<Space>v", String.raw`\hat{v}`],
  // Backspace right after an escape restores the typed letters, in any structure.
  ["pi<Space><Backspace>", "pi"],
  ["pi<Space><Backspace>+1", "pi+1"],
  ["theta<Space>alpha<Space><Backspace>", String.raw`\theta alpha`],
  ["frac<Space><Backspace>", "frac"],
  ["frac<Space>alpha<Space><Backspace><Tab>b", String.raw`\frac{alpha}{b}`],
  ["sqrt<Space>mu<Space><Backspace>", String.raw`\sqrt{mu}`],
  ["pi<Space>+<Backspace>", String.raw`\pi`],
  // The differential d, units and constants (the upright rule).
  ["dx", String.raw`\mathrm{d}x`],
  ["km/h", String.raw`\mathrm{km}/\mathrm{h}`],
  ["m/s", String.raw`\mathrm{m}/\mathrm{s}`],
  ["v=72 km/h", String.raw`v=72\ \mathrm{km}/\mathrm{h}`],
  ["3 km<Space>", String.raw`3\ \mathrm{km}\ `],
  ["F=mg", "F=mg"],
  ["l=lg", "l=lg"],
  ["lg(x)", String.raw`\lg\left(x\right)`],
  ["q=e^2", "q=e^2"],
  ["ii", String.raw`\mathrm{i}`],
  // Physics products stay italic.
  ["mgh", "mgh"], ["nRT", "nRT"], ["pV", "pV"], ["kT", "kT"], ["hc", "hc"],
  // Symbol sequences convert as typed.
  ["x>=y", String.raw`x\ge y`],
  ["a->b", String.raw`a\to b`],
];

// Subscripts kept open: the escape and its undo inside a subscript.
const SUBSCRIPTS = [
  ["x_alpha<Space>", String.raw`x_{\alpha}`],
  ["x_alpha<Space><Backspace>", String.raw`x_{alpha}`],
  ["x_thetaalpha<Space>", String.raw`x_{theta\alpha}`],
];

const UPRIGHT_OFF = [
  ["sinx+", "sinx+"],
  ["km/h+", "km/h+"],
  ["int<Space>xdx", String.raw`\int xdx`],
  ["dx", "dx"],
  ["pi<Space>+", String.raw`\pi+`],
  ["pi+", "pi+"],
  ["x>=y", String.raw`x\ge y`],
];

const ESCAPES_OFF = [
  ["pi+", "pi+"],
  ["pi<Space>+", String.raw`pi\ +`],
  ["x>=y", "x>=y"],
  ["sinx+", String.raw`\sin x+`],
  ["km/h+", String.raw`\mathrm{km}/\mathrm{h}+`],
  ["sintheta+", String.raw`\sin theta+`],
];

const BOTH_OFF = [
  ["sinx+pi+km/h+dx", "sinx+pi+km/h+dx"],
];

// The confirm-escape key changed to F2: Space is a plain key again.
const F2_KEY = [
  ["alpha<F2>", String.raw`\alpha`],
  ["alpha<Space>x", String.raw`alpha\ x`],
  ["thetaalpha<F2>", String.raw`theta\alpha`],
  ["pi<F2><Backspace>", "pi"],
];

// The confirm-escape key turned off.
const KEY_OFF = [
  ["alpha<Space>", String.raw`alpha\ `],
  ["x>=y", String.raw`x\ge y`],
];

const PRINTABLE = {
  "^": ["^", "Digit6", 54, 8], "_": ["_", "Minus", 189, 8], "(": ["(", "Digit9", 57, 8],
  ")": [")", "Digit0", 48, 8], "+": ["+", "Equal", 187, 8], " ": [" ", "Space", 32, 0],
  "/": ["/", "Slash", 191, 0], "=": ["=", "Equal", 187, 0], ">": [">", "Period", 190, 8],
  "-": ["-", "Minus", 189, 0],
};
const SPACE_CHORD = { code: "Space", key: " ", ctrlKey: false, altKey: false, shiftKey: false, metaKey: false };
const NAMED = { Tab: 9, Backspace: 8, F2: 113 };

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

const preview = spawn(
  process.execPath,
  ["node_modules/vite/bin/vite.js", "preview", "--host", "127.0.0.1", "--port", String(previewPort), "--strictPort"],
  { cwd: process.cwd(), stdio: "ignore" },
);
let chrome;
const failures = [];
let total = 0;

try {
  await waitFor(baseUrl);
  chrome = spawn(
    resolveChromiumExecutable(),
    [
      "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
      `--remote-debugging-port=${debugPort}`,
      `--user-data-dir=${createBrowserProfilePath("visualtex-input-judgment")}`,
      "--window-size=1400,1000", "about:blank",
    ],
    { stdio: "ignore" },
  );
  await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
  const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
  const client = new CdpClient(targets.find((t) => t.type === "page").webSocketDebuggerUrl);
  await client.connect();
  await client.send("Runtime.enable");
  await client.send("Page.enable");

  const evaluate = async (expression) => {
    const result = await client.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
    if (result.exceptionDetails)
      throw new Error(result.exceptionDetails.exception?.description ?? result.exceptionDetails.text);
    return result.result.value;
  };
  const press = async (name) => {
    let key, code, vk, mods = 0;
    if (NAMED[name]) [key, code, vk] = [name, name, NAMED[name]];
    else if (PRINTABLE[name]) [key, code, vk, mods] = PRINTABLE[name];
    else if (/^[0-9]$/.test(name)) [key, code, vk] = [name, `Digit${name}`, name.charCodeAt(0)];
    else [key, code, vk] = [name, `Key${name.toUpperCase()}`, name.toUpperCase().charCodeAt(0)];
    const printable = key.length === 1;
    const common = { key, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, modifiers: mods };
    await client.send("Input.dispatchKeyEvent", {
      type: printable ? "keyDown" : "rawKeyDown",
      ...common,
      ...(printable ? { text: key, unmodifiedText: key } : {}),
    });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
    await sleep(60);
  };
  const FIELD = `document.querySelector("math-field.visual-mathfield")`;

  const load = async (behavior, escapeChord) => {
    await client.send("Page.navigate", { url: baseUrl });
    for (let i = 0; i < 100; i++) {
      try {
        if (await evaluate(`document.readyState === "complete" && location.href.startsWith(${JSON.stringify(baseUrl)})`)) break;
      } catch {
        // navigating
      }
      await sleep(100);
    }
    await evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      const id = crypto.randomUUID();
      persisted.state = { ...(persisted.state || {}), lines: [{ id, latex: "" }], activeLineId: id, sourceOpen: false,
        inputBehavior: { ...((persisted.state || {}).inputBehavior || {}), ...${JSON.stringify(behavior)} } };
      localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
      localStorage.setItem("visualtex-confirm-escape-key", JSON.stringify({ state: { chord: ${JSON.stringify(escapeChord === undefined ? SPACE_CHORD : escapeChord)} }, version: 1 }));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    for (let i = 0; i < 300; i++) {
      try {
        if (await evaluate(`Boolean(${FIELD}?.shadowRoot)`)) break;
      } catch {
        // reloading
      }
      await sleep(50);
    }
    await evaluate(`document.fonts.ready.then(() => true)`);
    await sleep(400);
  };

  const runCase = async (keys) => {
    await evaluate(`(() => { const f = ${FIELD};
      f.setValue("", { mode: "math", format: "latex", insertionMode: "replaceAll", selectionMode: "after", silenceNotifications: true });
      f.dispatchEvent(new InputEvent("input", { bubbles: true, composed: true, inputType: "insertText" }));
      f.focus(); f.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      return true; })()`);
    await sleep(120);
    for (const part of keys.split(/(<[A-Za-z0-9]+>)/).filter(Boolean)) {
      const named = /^<([A-Za-z0-9]+)>$/.exec(part);
      if (named) {
        await press(named[1] === "Space" ? " " : named[1]);
        await sleep(120);
      } else {
        for (const character of part) await press(character);
      }
    }
    await sleep(200);
    await evaluate(`${FIELD}.blur()`);
    await sleep(200);
    return evaluate(`${FIELD}.value`);
  };

  const ON = { autoUprightTypography: true, autoEscapeShortcuts: true };
  for (const [label, behavior, cases, escapeChord] of [
    ["both on", ON, BOTH_ON],
    ["subscripts kept open", { ...ON, autoExitSubscript: false }, SUBSCRIPTS],
    ["upright off", { autoUprightTypography: false, autoEscapeShortcuts: true }, UPRIGHT_OFF],
    ["quick escapes off", { autoUprightTypography: true, autoEscapeShortcuts: false }, ESCAPES_OFF],
    ["both off", { autoUprightTypography: false, autoEscapeShortcuts: false }, BOTH_OFF],
    ["confirm-escape key F2", ON, F2_KEY, { ...SPACE_CHORD, code: "F2", key: "F2" }],
    ["confirm-escape key off", ON, KEY_OFF, null],
  ]) {
    await load(behavior, escapeChord);
    console.log(`--- ${label}`);
    for (const [keys, expected] of cases) {
      total += 1;
      const actual = await runCase(keys);
      const ok = actual === expected;
      if (!ok) failures.push({ label, keys, expected, actual });
      if (!ok || verbose) console.log(`${ok ? "PASS" : "FAIL"} ${keys.padEnd(26)} => ${actual}${ok ? "" : `   (expected ${expected})`}`);
    }
  }
} finally {
  chrome?.kill();
  preview.kill();
}

console.log(`${total - failures.length}/${total} passed`);
if (failures.length) process.exitCode = 1;
