import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 900;
const previewPort = 9000 + offset;
const debugPort = 15000 + offset;
const url = `http://127.0.0.1:${previewPort}`;
const profile = `/tmp/visualtex-ime-matrix-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(target, timeout = 15000) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    try {
      if ((await fetch(target)).ok) return;
    } catch {}
    await sleep(80);
  }
  throw new Error(`Timeout: ${target}`);
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
    ["--headless=new", "--disable-gpu", "--no-first-run", `--remote-debugging-port=${debugPort}`, `--user-data-dir=${profile}`, "--window-size=1200,800", url],
    { stdio: "ignore" },
  );
  await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
  const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
  const page = targets.find((item) => item.type === "page" && item.url.startsWith(url));
  assert.ok(page, "VisualTeX page target missing");
  client = new Cdp(page.webSocketDebuggerUrl);
  await client.connect();
  await client.send("Runtime.enable");
  await client.send("Page.enable");
  await client.send("Page.navigate", { url });
  await sleep(600);

  const evaluate = async (expression) => {
    const result = await client.send("Runtime.evaluate", {
      expression,
      awaitPromise: true,
      returnByValue: true,
    });
    if (result.exceptionDetails) throw new Error(result.exceptionDetails.text);
    return result.result.value;
  };

  await evaluate(`(() => {
    localStorage.setItem("visualtex.onboarding.v3.completed", "true");
    localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
    localStorage.setItem("visualtex.onboarding.macos.desktop.v1.2.0.completed", "true");
    localStorage.setItem("visualtex.office.macos.native-first-run.v1.2.0.completed", "true");
  })()`);
  await client.send("Page.reload", { ignoreCache: true });
  await sleep(550);

  let caseId = 0;
  const reset = async () => {
    caseId += 1;
    await evaluate(`(() => {
      const key = "visualtex-editor";
      const persisted = JSON.parse(localStorage.getItem(key) || "{}");
      persisted.state = {
        ...(persisted.state || {}),
        lines: [{ id: "ime-matrix-${caseId}", latex: "" }],
        activeLineId: "ime-matrix-${caseId}",
        sourceOpen: false,
      };
      localStorage.setItem(key, JSON.stringify(persisted));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(480);
    await evaluate(`new Promise((resolve) => {
      const poll = () => {
        const field = document.querySelector("math-field");
        if (!field) return setTimeout(poll, 25);
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        field.dataset.lateCompletionCount = "0";
        field.addEventListener("keydown", (event) => {
          if (event.key === " " || event.code === "Space") {
            field.dataset.lateCompletionCount = String(
              Number(field.dataset.lateCompletionCount || "0") + 1,
            );
          }
        }, true);
        resolve(true);
      };
      poll();
    })`);
  };

  const cancelComposition = async (stale = true) => {
    const state = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      const before = (type, data, composing) => {
        const event = new InputEvent("beforeinput", { inputType: type, data, isComposing: composing, bubbles: true, composed: true, cancelable: true });
        const allowed = field.dispatchEvent(event);
        return { allowed, prevented: event.defaultPrevented };
      };
      const input = (type, data, composing) => field.dispatchEvent(new InputEvent("input", { inputType: type, data, isComposing: composing, bubbles: true, composed: true }));
      field.dispatchEvent(new CompositionEvent("compositionstart", { data: "", bubbles: true, composed: true }));
      const inserted = before("insertCompositionText", "中", true);
      if (inserted.allowed) field.insert("中", { mode: "math", format: "latex", insertionMode: "replaceSelection", selectionMode: "after", focus: true });
      input("insertCompositionText", "中", true);
      const deleted = before("deleteCompositionText", "", true);
      input("deleteCompositionText", "", true);
      if (deleted.allowed) field.executeCommand("deleteBackward");
      field.dispatchEvent(new CompositionEvent("compositionend", { data: "", bubbles: true, composed: true }));
      let replay = null;
      if (${stale}) replay = before("insertText", "\\\\", false);
      return { value: field.value, replay };
    })()`);
    assert.equal(state.value, "", JSON.stringify(state));
    if (stale) assert.equal(state.replay.prevented, true, JSON.stringify(state));
  };

  const key = (value, code, keyCode) => ({ key: value, code, windowsVirtualKeyCode: keyCode, nativeVirtualKeyCode: keyCode });
  const type = async (value, code, keyCode, pause = 65) => {
    const common = key(value, code, keyCode);
    await client.send("Input.dispatchKeyEvent", { type: "keyDown", ...common, text: value, unmodifiedText: value });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
    await sleep(pause);
  };

  const slash = async (pattern) => {
    const common = key("\\", "Backslash", 220);
    if (pattern === "normal") return type("\\", "Backslash", 220, 80);
    if (pattern === "duplicate-char") {
      await client.send("Input.dispatchKeyEvent", { type: "keyDown", ...common });
      await client.send("Input.dispatchKeyEvent", { type: "char", ...common, text: "\\", unmodifiedText: "\\" });
      await client.send("Input.dispatchKeyEvent", { type: "char", ...common, text: "\\", unmodifiedText: "\\" });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      return sleep(90);
    }
    if (pattern === "duplicate-keydown") {
      for (let index = 0; index < 2; index += 1) {
        await client.send("Input.dispatchKeyEvent", { type: "keyDown", ...common });
        await client.send("Input.dispatchKeyEvent", { type: "char", ...common, text: "\\", unmodifiedText: "\\" });
      }
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      return sleep(90);
    }
    const comma = key("、", "Backslash", 220);
    await client.send("Input.dispatchKeyEvent", { type: "keyDown", ...comma, text: "、", unmodifiedText: "、" });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...comma });
    await sleep(210);
    return type("\\", "Backslash", 220, 80);
  };

  const state = () => evaluate(`(() => {
    const field = document.querySelector("math-field");
    return {
      value: field.value,
      raw: Array.from(field.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? []).filter((node) => !node.classList.contains("ML__suggestion")).map((node) => node.textContent ?? "").join(""),
      candidate: field.dataset.pendingNativeSuggestion || document.querySelector("#mathlive-suggestion-popover li.ML__popover__current")?.dataset.command || "",
      lateCompletionCount: Number(field.dataset.lateCompletionCount || "0"),
    };
  })()`);

  const patterns = ["normal", "duplicate-char", "duplicate-keydown", "ideographic-then-latin"];
  const commands = ["pi", "int", "sum", "sqrt", "frac"];
  for (const pattern of patterns) {
    for (const command of commands) {
      await reset();
      await cancelComposition(true);
      await slash(pattern);
      for (const character of command) {
        const upper = character.toUpperCase();
        await type(character, `Key${upper}`, upper.charCodeAt(0));
      }
      const pending = await state();
      if (pending.raw) {
        assert.equal(pending.raw, `\\${command}`, `${pattern}/${command}: ${JSON.stringify(pending)}`);
        if (pending.candidate) assert.equal(pending.candidate, `\\${command}`);
      } else {
        assert.equal(
          pending.value.split(`\\${command}`).length - 1,
          1,
          `${pattern}/${command} completed before Space with the wrong semantic value: ${JSON.stringify(pending)}`,
        );
      }
      await type(" ", "Space", 32, 35);
      const firstCompletion = await state();
      assert.equal(
        firstCompletion.value.split(`\\${command}`).length - 1,
        1,
        `${pattern}/${command} first completion was not semantically unique: ${JSON.stringify(firstCompletion)}`,
      );
      await type(" ", "Space", 32, 150);
      const complete = await state();
      assert.equal(complete.raw, "", JSON.stringify(complete));
      assert.equal(complete.value.split(`\\${command}`).length - 1, 1, `${pattern}/${command}: ${JSON.stringify(complete)}`);
    }
  }

  await reset();
  await cancelComposition(true);
  await slash("normal");
  await type("i", "KeyI", 73);
  await type("n", "KeyN", 78);
  await type("t", "KeyT", 84);
  await evaluate(`(() => {
    const field = document.querySelector("math-field");
    field.dataset.pendingNativeSuggestion = "\\\\int";
    field.dataset.lateCompletionCount = "0";
  })()`);
  await type(" ", "Space", 32, 150);
  const forcedVisualTexCommit = await state();
  assert.equal(
    forcedVisualTexCommit.lateCompletionCount,
    0,
    `VisualTeX-owned completion reached a later listener: ${JSON.stringify(forcedVisualTexCommit)}`,
  );
  assert.equal(forcedVisualTexCommit.value.split("\\int").length - 1, 1);

  await reset();
  for (let index = 0; index < 3; index += 1) await cancelComposition(true);
  await slash("normal");
  await type("p", "KeyP", 80);
  await type("i", "KeyI", 73);
  await type(" ", "Space", 32, 150);
  assert.equal((await state()).value, "\\pi");

  await reset();
  for (let repeat = 0; repeat < 2; repeat += 1) {
    await slash("normal");
    await type("i", "KeyI", 73);
    await type("n", "KeyN", 78);
    await type("t", "KeyT", 84);
    await type(" ", "Space", 32, 450);
  }
  const intentionalRepeat = await state();
  assert.equal(
    intentionalRepeat.value.split("\\int").length - 1,
    2,
    `Intentional repeated integral was suppressed: ${JSON.stringify(intentionalRepeat)}`,
  );

  console.log("IME Backslash and completion transaction matrix passed");
} finally {
  client?.close();
  chrome?.kill("SIGTERM");
  preview.kill("SIGTERM");
  await sleep(250);
  await rm(profile, { recursive: true, force: true });
}
