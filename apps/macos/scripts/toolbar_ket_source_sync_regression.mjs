import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 700;
const previewPort = 14600 + offset;
const debugPort = 16600 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const profile = `/tmp/visualtex-ket-source-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {}
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

class Cdp {
  constructor(url) {
    this.url = url;
    this.id = 1;
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
    const id = this.id++;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }
  close() { this.socket?.close(); }
}

async function main() {
  const preview = spawn(process.execPath, [
    "node_modules/vite/bin/vite.js", "preview", "--host", "127.0.0.1",
    "--port", String(previewPort), "--strictPort",
  ], { cwd: process.cwd(), stdio: "ignore" });
  let chrome;
  let client;
  try {
    await waitFor(baseUrl);
    chrome = spawn(chromePath, [
      "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
      `--remote-debugging-port=${debugPort}`, `--user-data-dir=${profile}`,
      "--window-size=1500,1000", baseUrl,
    ], { stdio: "ignore" });
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
    const page = targets.find((target) => target.type === "page" && target.url.startsWith(baseUrl));
    if (!page) throw new Error("No VisualTeX page target");
    client = new Cdp(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");

    const evaluate = async (expression) => {
      const result = await client.send("Runtime.evaluate", {
        expression, awaitPromise: true, returnByValue: true,
      });
      if (result.exceptionDetails) {
        throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
      }
      return result.result.value;
    };

    await evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.onboarding.macos.desktop.v1.2.0.completed", "true");
      localStorage.setItem("visualtex.office.macos.native-first-run.v1.2.0.completed", "true");
      localStorage.setItem("visualtex.release-welcome.1.2.6.seen", "true");
      let persisted;
      try { persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "null"); } catch { persisted = null; }
      if (!persisted || typeof persisted !== "object") persisted = { state: {}, version: 0 };
      persisted.state = {
        ...(persisted.state || {}),
        title: "Ket Source Sync",
        lines: [
          { id: "ket-1", latex: "a", mode: "display" },
          { id: "ket-2", latex: "b", mode: "display" },
          { id: "ket-3", latex: "", mode: "display" },
        ],
        activeLineId: "ket-3",
        sourceOpen: true,
        editorLayout: "standard",
        latexCodeFormat: "raw",
        checkUpdatesOnStartup: false,
        language: "cn",
      };
      delete persisted.state.latex;
      localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(1000);
    await evaluate(`new Promise((resolve, reject) => {
      const started = performance.now();
      const poll = () => {
        if (document.querySelectorAll("math-field").length === 3 && document.querySelector(".source-panel .cm-content")) return resolve(true);
        if (performance.now() - started > 5000) return reject(new Error("Editor did not mount three fields/source"));
        setTimeout(poll, 30);
      }; poll();
    })`);

    await evaluate(`(() => {
      const field = document.querySelectorAll("math-field")[2];
      field.__visualtexKetRegressionIdentity = "ket-third-field";
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      document.querySelector('button[data-category="physics"]')?.click();
    })()`);
    await sleep(180);
    const toolbarState = await evaluate(`(() => ({
      categories: [...document.querySelectorAll('button[data-category]')].map((button) => ({
        category: button.dataset.category,
        active: button.classList.contains('is-active'),
      })),
      commands: [...document.querySelectorAll('button[data-command-id]')].map((button) => button.dataset.commandId),
    }))()`);
    assert.equal(
      toolbarState.commands.includes("shortcut-ket"),
      true,
      `Physics toolbar must expose shortcut-ket: ${JSON.stringify(toolbarState)}`,
    );
    await evaluate(`document.querySelector('button[data-command-id="shortcut-ket"]').click()`);
    await sleep(220);

    const inserted = await evaluate(`(() => {
      const fields = [...document.querySelectorAll("math-field")];
      return {
        values: fields.map((field) => field.value),
        source: [...document.querySelectorAll(".source-panel .cm-line")].map((line) => line.textContent ?? ""),
        thirdSelection: fields[2].selection,
        thirdSelectedLatex: fields[2].selectionIsCollapsed ? "" : fields[2].getValue(fields[2].selection),
        toolbarPackageMacroName: fields[2].dataset.toolbarPackageMacroName ?? "",
        toolbarPackageMacroPlaceholderIndex: fields[2].dataset.toolbarPackageMacroPlaceholderIndex ?? "",
        toolbarPackageMacroModelStart: fields[2].dataset.toolbarPackageMacroModelStart ?? "",
        toolbarPackageMacroProbe: fields[2].dataset.toolbarPackageMacroProbe ?? "",
        activeElement: document.activeElement?.tagName ?? "",
        thirdHasFocus: fields[2].matches(":focus"),
        sinkHasFocus: fields[2].shadowRoot?.activeElement?.getAttribute("part") ?? "",
      };
    })()`);
    assert.match(inserted.values[2], /\\ket\{/, `Toolbar did not insert ket: ${JSON.stringify(inserted)}`);

    const common = { key: "x", code: "KeyX", windowsVirtualKeyCode: 88, nativeVirtualKeyCode: 88 };
    await client.send("Input.dispatchKeyEvent", { type: "keyDown", ...common, text: "x", unmodifiedText: "x" });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
    await sleep(350);
    const afterXDiagnostic = await evaluate(`(() => {
      const field = document.querySelectorAll("math-field")[2];
      const marker = [...(field.shadowRoot?.querySelectorAll('.ML__caret,.ML__text-caret,.ML__latex-caret') ?? [])]
        .find((node) => node.getBoundingClientRect().height > 0);
      const ancestors = [];
      let current = marker;
      while (current && ancestors.length < 10) {
        ancestors.push(current.className || current.tagName);
        current = current.parentElement;
      }
      return {
        value: field.value,
        position: field.position,
        lastOffset: field.lastOffset,
        selection: field.selection,
        pendingPackageMacroCommand: field.dataset.pendingPackageMacroCommand ?? "",
        pendingPackageMacroContent: field.dataset.pendingPackageMacroContent ?? "",
        pendingPackageMacroClearReason: field.dataset.pendingPackageMacroClearReason ?? "",
        toolbarPackageMacroName: field.dataset.toolbarPackageMacroName ?? "",
        toolbarPackageMacroPlaceholderIndex: field.dataset.toolbarPackageMacroPlaceholderIndex ?? "",
        toolbarPackageMacroModelStart: field.dataset.toolbarPackageMacroModelStart ?? "",
        toolbarPackageMacroProbe: field.dataset.toolbarPackageMacroProbe ?? "",
        sameFieldInstance: field.__visualtexKetRegressionIdentity === "ket-third-field",
        prefix: field.getValue(0, field.position, "latex"),
        elementInfo: Array.from({ length: field.lastOffset + 1 }, (_, offset) => ({
          offset,
          depth: field.getElementInfo(offset)?.depth ?? null,
          latex: field.getElementInfo(offset)?.latex ?? "",
        })),
        caretAncestors: ancestors,
      };
    })()`);

    const commonY = { key: "y", code: "KeyY", windowsVirtualKeyCode: 89, nativeVirtualKeyCode: 89 };
    await client.send("Input.dispatchKeyEvent", { type: "keyDown", ...commonY, text: "y", unmodifiedText: "y" });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...commonY });
    await sleep(250);

    const after = await evaluate(`(() => {
      const fields = [...document.querySelectorAll("math-field")];
      const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      return {
        values: fields.map((field) => field.value),
        source: [...document.querySelectorAll(".source-panel .cm-line")].map((line) => line.textContent ?? ""),
        persistedLines: persisted?.state?.lines ?? [],
        activeLine: document.querySelector(".formula-line.is-active")?.dataset.lineId ?? "",
      };
    })()`);
    assert.match(
      after.values[2],
      /\\ket\{xy\}/,
      `Third MathLive field did not keep accepting ket content: ${JSON.stringify({ inserted, afterXDiagnostic, after })}`,
    );
    assert.match(after.source[2] ?? "", /\\ket\{xy\}/, `Third source line did not live-sync ket content: ${JSON.stringify(after)}`);

    const backspace = { key: "Backspace", code: "Backspace", windowsVirtualKeyCode: 8, nativeVirtualKeyCode: 8 };
    await client.send("Input.dispatchKeyEvent", { type: "rawKeyDown", ...backspace });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...backspace });
    await sleep(180);
    const afterBackspace = await evaluate(`(() => ({
      value: document.querySelectorAll("math-field")[2]?.value ?? "",
      source: [...document.querySelectorAll(".source-panel .cm-line")].map((line) => line.textContent ?? "")[2] ?? "",
    }))()`);
    assert.match(afterBackspace.value, /\\ket\{x\}/, JSON.stringify(afterBackspace));
    assert.match(afterBackspace.source, /\\ket\{x\}/, JSON.stringify(afterBackspace));
    await client.send("Input.dispatchKeyEvent", { type: "keyDown", ...commonY, text: "y", unmodifiedText: "y" });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...commonY });
    await sleep(180);
    const afterRetype = await evaluate(`(() => ({
      value: document.querySelectorAll("math-field")[2]?.value ?? "",
      source: [...document.querySelectorAll(".source-panel .cm-line")].map((line) => line.textContent ?? "")[2] ?? "",
    }))()`);
    assert.match(afterRetype.value, /\\ket\{xy\}/, JSON.stringify(afterRetype));
    assert.match(afterRetype.source, /\\ket\{xy\}/, JSON.stringify(afterRetype));

    await evaluate(`(() => {
      const field = document.querySelectorAll("math-field")[2];
      field.position = field.lastOffset;
      field.selection = { ranges: [[field.lastOffset, field.lastOffset]], direction: "none" };
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      document.querySelector('button[data-category="arrow"]')?.click();
    })()`);
    await sleep(180);
    const arrowButtons = await evaluate(`
      [...document.querySelectorAll('button[data-command-id]')].map((button) => button.dataset.commandId)
    `);
    assert.equal(arrowButtons.includes("reaction-arrow-labeled"), true, "Reaction arrow toolbar button is missing");
    assert.equal(arrowButtons.includes("equilibrium-arrow-labeled"), true, "Equilibrium reaction arrow toolbar button is missing");
    await evaluate(`document.querySelector('button[data-command-id="reaction-arrow-labeled"]').click()`);
    await sleep(220);

    const typeAscii = async (key, code, keyCode) => {
      const commonKey = { key, code, windowsVirtualKeyCode: keyCode, nativeVirtualKeyCode: keyCode };
      await client.send("Input.dispatchKeyEvent", { type: "keyDown", ...commonKey, text: key, unmodifiedText: key });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...commonKey });
      await sleep(180);
    };
    await typeAscii("U", "KeyU", 85);
    const tab = { key: "Tab", code: "Tab", windowsVirtualKeyCode: 9, nativeVirtualKeyCode: 9 };
    await client.send("Input.dispatchKeyEvent", { type: "rawKeyDown", ...tab });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...tab });
    await sleep(160);
    await typeAscii("L", "KeyL", 76);
    await sleep(220);
    const reactionState = await evaluate(`(() => ({
      value: document.querySelectorAll("math-field")[2]?.value ?? "",
      source: [...document.querySelectorAll(".source-panel .cm-line")].map((line) => line.textContent ?? "")[2] ?? "",
    }))()`);
    assert.match(reactionState.value, /\\xrightarrow/, JSON.stringify(reactionState));
    assert.match(reactionState.value, /U/, JSON.stringify(reactionState));
    assert.match(reactionState.value, /L/, JSON.stringify(reactionState));
    assert.match(reactionState.source, /\\xrightarrow/, JSON.stringify(reactionState));

    console.log("Toolbar ket source sync and labeled reaction-arrow regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(250);
    await rm(profile, { recursive: true, force: true, maxRetries: 6, retryDelay: 100 });
  }
}

await main();
