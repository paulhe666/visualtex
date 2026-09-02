import assert from "node:assert/strict";
import { spawn, spawnSync } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 700;
const previewPort = 9700 + offset;
const debugPort = 16000 + offset;
const url = `http://127.0.0.1:${previewPort}`;
const profile = `/tmp/visualtex-physical-ime-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const cleanupTauri = process.argv.includes("--cleanup-tauri");

const tauriAcceptanceBaseline = String.raw`$$
x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}\alpha\beta\oint_{a}b\mathbf{\mathbfit{dgF}}JJ\text{哈哈哈}
$$

$$
x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}\bm{a}\alpha\beta\theta\alpha\alpha
$$
`;

const tauriAcceptanceWithTests = `${tauriAcceptanceBaseline}\n$$\n\\alpha\n$$\n\n$$\n\\int\n$$\n`;

async function waitFor(target, timeout = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeout) {
    try {
      if ((await fetch(target)).ok) return;
    } catch {}
    await sleep(80);
  }
  throw new Error(`Timeout waiting for ${target}`);
}

class Cdp {
  constructor(socketUrl) {
    this.socketUrl = socketUrl;
    this.nextId = 1;
    this.pending = new Map();
  }
  async connect() {
    this.socket = new WebSocket(this.socketUrl);
    await new Promise((resolve, reject) => {
      this.socket.addEventListener("open", resolve, { once: true });
      this.socket.addEventListener("error", reject, { once: true });
    });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
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

function run(command, args) {
  const result = spawnSync(command, args, { encoding: "utf8" });
  if (result.status !== 0) {
    throw new Error(result.stderr || result.stdout || `${command} failed`);
  }
  return result.stdout;
}

function osascript(lines) {
  const args = [];
  for (const line of lines) args.push("-e", line);
  return run("osascript", args);
}

function jxa(script) {
  return run("osascript", ["-l", "JavaScript", "-e", script]);
}

function selectedInputSource() {
  const result = spawnSync(
    "defaults",
    ["read", "com.apple.HIToolbox", "AppleSelectedInputSources"],
    { encoding: "utf8" },
  );
  return `${result.stdout || ""}\n${result.stderr || ""}`;
}

function sourceKind() {
  const source = selectedInputSource();
  if (source.includes("com.apple.inputmethod.SCIM.ITABC")) return "chinese";
  if (source.includes("KeyboardLayout Name") && source.includes("ABC")) return "abc";
  return "unknown";
}

async function toggleInputSource() {
  osascript([
    'tell application "System Events"',
    'key code 49 using {control down}',
    'end tell',
  ]);
  await sleep(350);
}

async function ensureInputSource(target) {
  for (let attempt = 0; attempt < 4; attempt += 1) {
    if (sourceKind() === target) return;
    await toggleInputSource();
  }
  throw new Error(`Could not select ${target}; current=${selectedInputSource()}`);
}

function physical(lines) {
  if (!chrome?.pid) throw new Error("Physical Chrome process is not running");
  osascript([
    'tell application "System Events"',
    `set frontmost of first application process whose unix id is ${chrome.pid} to true`,
    'delay 0.08',
    ...lines,
    'end tell',
  ]);
}

if (cleanupTauri) {
  const readSource = () => {
    const payload = jxa(`(() => {
      const se = Application("System Events");
      const processes = se.applicationProcesses.whose({ name: { _equals: "visualtex" } })();
      const walk = (element, depth = 0) => {
        if (depth > 14) return null;
        try {
          if (element.role() === "AXTextArea") return String(element.value() ?? "");
        } catch {}
        try {
          for (const child of element.uiElements()) {
            const value = walk(child, depth + 1);
            if (value !== null) return value;
          }
        } catch {}
        return null;
      };
      for (const process of processes) {
        let windows = [];
        try { windows = process.windows(); } catch {}
        for (const window of windows) {
          const source = walk(window);
          if (source !== null) return JSON.stringify({ source });
        }
      }
      throw new Error("VisualTeX source editor is not visible");
    })()`);
    return JSON.parse(payload).source;
  };
  const undo = () => osascript([
    'tell application "System Events"',
    'set appProcesses to every application process whose name is "visualtex"',
    'set appProcess to missing value',
    'repeat with candidateProcess in appProcesses',
    'if (count of windows of candidateProcess) > 0 then',
    'set appProcess to candidateProcess',
    'exit repeat',
    'end if',
    'end repeat',
    'if appProcess is missing value then error "VisualTeX main window is not open"',
    'set frontmost of appProcess to true',
    'keystroke "z" using {command down}',
    'end tell',
  ]);

  let source = readSource();
  assert.equal(
    source,
    tauriAcceptanceWithTests,
    `Refusing to clean Tauri acceptance state because the current source is not the exact expected test state: ${JSON.stringify(source)}`,
  );
  for (let attempt = 0; attempt < 12 && source !== tauriAcceptanceBaseline; attempt += 1) {
    undo();
    await sleep(180);
    source = readSource();
  }
  assert.equal(
    source,
    tauriAcceptanceBaseline,
    `Tauri acceptance cleanup did not restore the exact baseline: ${JSON.stringify(source)}`,
  );
  console.log("Tauri IME acceptance cleanup restored the exact pre-test source");
  process.exit(0);
}

const originalInputSourceKind = sourceKind();

const preview = spawn(
  process.execPath,
  ["node_modules/vite/bin/vite.js", "preview", "--host", "127.0.0.1", "--port", String(previewPort), "--strictPort"],
  { cwd: process.cwd(), stdio: "ignore" },
);
let chrome;
let client;

try {
  await waitFor(url);
  chrome = spawn(
    chromePath,
    [
      "--no-first-run",
      "--no-default-browser-check",
      `--remote-debugging-port=${debugPort}`,
      `--user-data-dir=${profile}`,
      "--window-size=1200,800",
      url,
    ],
    { stdio: "ignore" },
  );
  await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
  const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
  const page = targets.find((item) => item.type === "page" && item.url.startsWith(url));
  assert.ok(page, "VisualTeX Chrome page missing");
  client = new Cdp(page.webSocketDebuggerUrl);
  await client.connect();
  await client.send("Runtime.enable");
  await client.send("Page.enable");

  const evaluate = async (expression) => {
    const result = await client.send("Runtime.evaluate", {
      expression,
      awaitPromise: true,
      returnByValue: true,
    });
    if (result.exceptionDetails) {
      throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
    }
    return result.result.value;
  };

  await evaluate(`(() => {
    localStorage.setItem("visualtex.onboarding.v3.completed", "true");
    localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
    localStorage.setItem("visualtex.onboarding.macos.desktop.v1.2.0.completed", "true");
    localStorage.setItem("visualtex.office.macos.native-first-run.v1.2.0.completed", "true");
  })()`);
  await client.send("Page.reload", { ignoreCache: true });
  await sleep(650);

  const reset = async (id) => {
    await evaluate(`(() => {
      const key = "visualtex-editor";
      const persisted = JSON.parse(localStorage.getItem(key) || "{}");
      persisted.state = {
        ...(persisted.state || {}),
        lines: [{ id: ${JSON.stringify(id)}, latex: "" }],
        activeLineId: ${JSON.stringify(id)},
        sourceOpen: false,
      };
      localStorage.setItem(key, JSON.stringify(persisted));
      location.reload();
    })()`);
    await sleep(650);
    await evaluate(`new Promise((resolve) => {
      const poll = () => {
        const field = document.querySelector("math-field");
        const sink = field?.shadowRoot?.querySelector('[part="keyboard-sink"]');
        if (!field || !sink) return setTimeout(poll, 25);
        window.__visualtexPhysicalImeTrace = [];
        const trace = (where, event) => {
          if ((event.type === "keydown" && event.key !== " " && event.code !== "Space") ||
              (event.type !== "keydown" && event.type !== "beforeinput" && event.type !== "input" && !event.type.startsWith("composition"))) return;
          window.__visualtexPhysicalImeTrace.push({
            where,
            type: event.type,
            key: event.key ?? "",
            code: event.code ?? "",
            keyCode: event.keyCode ?? 0,
            isComposing: event.isComposing ?? false,
            inputType: event.inputType ?? "",
            data: event.data ?? null,
            phase: event.eventPhase,
            prevented: event.defaultPrevented,
            value: field.value,
            raw: [...(field.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
              .filter((node) => !node.classList.contains("ML__suggestion"))
              .map((node) => node.textContent ?? "")
              .join(""),
          });
        };
        for (const [where, target] of [["window", window], ["field", field], ["sink", sink]]) {
          for (const type of ["compositionstart", "compositionupdate", "compositionend", "keydown", "beforeinput", "input"]) {
            target.addEventListener(type, (event) => trace(where, event), true);
          }
        }
        const originalExecute = field.executeCommand.bind(field);
        field.executeCommand = (command) => {
          const before = { value: field.value, raw: [...(field.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])].filter((node) => !node.classList.contains("ML__suggestion")).map((node) => node.textContent ?? "").join("") };
          const result = originalExecute(command);
          window.__visualtexPhysicalImeTrace.push({ where: "executeCommand", command: JSON.stringify(command), result, before, after: { value: field.value, raw: [...(field.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])].filter((node) => !node.classList.contains("ML__suggestion")).map((node) => node.textContent ?? "").join("") } });
          return result;
        };
        const originalInsert = field.insert.bind(field);
        field.insert = (latex, options) => {
          const before = { value: field.value, raw: [...(field.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])].filter((node) => !node.classList.contains("ML__suggestion")).map((node) => node.textContent ?? "").join("") };
          const result = originalInsert(latex, options);
          window.__visualtexPhysicalImeTrace.push({ where: "insert", latex, result, before, after: { value: field.value, raw: [...(field.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])].filter((node) => !node.classList.contains("ML__suggestion")).map((node) => node.textContent ?? "").join("") } });
          return result;
        };
        field.focus();
        sink.focus({ preventScroll: true });
        resolve(true);
      };
      poll();
    })`);
    physical([]);
    await sleep(180);
  };

  const state = () => evaluate(`(() => {
    const field = document.querySelector("math-field");
    const panel = document.getElementById("mathlive-suggestion-popover");
    return {
      value: field?.value ?? "",
      raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
        .filter((node) => !node.classList.contains("ML__suggestion"))
        .map((node) => node.textContent ?? "")
        .join(""),
      selected: panel?.querySelector("li.ML__popover__current[data-command]")?.dataset.command ?? "",
      candidates: [...(panel?.querySelectorAll("li[data-command]") ?? [])].map((node) => node.dataset.command ?? ""),
      trace: window.__visualtexPhysicalImeTrace ?? [],
    };
  })()`);

  const chooseNative = async (command) => {
    for (let index = 0; index < 12; index += 1) {
      const current = await state();
      if (current.selected === command) return current;
      physical(['key code 125']);
      await sleep(90);
    }
    throw new Error(`Could not select ${command}: ${JSON.stringify(await state())}`);
  };

  const reproduce = async ({ id, prefix, command }) => {
    await reset(id);
    await ensureInputSource("chinese");
    physical(['keystroke "a"']);
    await sleep(320);
    physical(['key code 51']);
    await sleep(240);
    const afterDelete = await state();
    await ensureInputSource("abc");
    physical(['key code 42', `keystroke ${JSON.stringify(prefix)}`]);
    await sleep(320);
    const beforeSelection = await state();
    const selected = await chooseNative(command);
    physical(['key code 49']);
    await sleep(350);
    const afterSpace = await state();
    return { afterDelete, beforeSelection, selected, afterSpace };
  };

  const alpha = await reproduce({ id: "physical-alpha", prefix: "al", command: "\\alpha" });
  console.log(JSON.stringify({ alpha }, null, 2));
  assert.equal(
    alpha.afterSpace.value.split("\\alpha").length - 1,
    1,
    `Physical Chinese-a/delete -> ABC -> \\al -> alpha -> Space duplicated: ${JSON.stringify(alpha.afterSpace)}`,
  );
  assert.equal(alpha.afterSpace.raw, "", JSON.stringify(alpha.afterSpace));

  const integral = await reproduce({ id: "physical-int", prefix: "in", command: "\\int" });
  console.log(JSON.stringify({ integral }, null, 2));
  assert.equal(
    integral.afterSpace.value.split("\\int").length - 1,
    1,
    `Physical Chinese-a/delete -> ABC -> \\in -> int -> Space duplicated: ${JSON.stringify(integral.afterSpace)}`,
  );
  assert.equal(integral.afterSpace.raw, "", JSON.stringify(integral.afterSpace));

  console.log("macOS physical IME native-command regression passed");
} finally {
  if (
    chrome?.pid &&
    (originalInputSourceKind === "chinese" || originalInputSourceKind === "abc") &&
    sourceKind() !== originalInputSourceKind
  ) {
    try {
      physical([]);
      await ensureInputSource(originalInputSourceKind);
    } catch (error) {
      console.warn(`Could not restore the original macOS input source: ${error}`);
    }
  }
  client?.close();
  chrome?.kill("SIGTERM");
  preview.kill("SIGTERM");
  await sleep(300);
  await rm(profile, { recursive: true, force: true });
}
