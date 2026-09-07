import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const offset = process.pid % 700;
const previewPort = 15100 + offset;
const debugPort = 17100 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const profile = createBrowserProfilePath("visualtex-nested-ket");
const chromePath = resolveChromiumExecutable();
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

async function main() {
  const preview = spawn(
    process.execPath,
    [
      "node_modules/vite/bin/vite.js",
      "preview",
      "--host",
      "127.0.0.1",
      "--port",
      String(previewPort),
      "--strictPort",
    ],
    { cwd: process.cwd(), stdio: "ignore" },
  );
  let chrome;
  let client;

  try {
    await waitFor(baseUrl);
    chrome = spawn(
      chromePath,
      [
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        `--remote-debugging-port=${debugPort}`,
        `--user-data-dir=${profile}`,
        "--window-size=1500,1000",
        baseUrl,
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json();
    const page = targets.find(
      (target) => target.type === "page" && target.url.startsWith(baseUrl),
    );
    if (!page) throw new Error("No VisualTeX page target found");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(700);

    const evaluate = async (expression) => {
      const result = await client.send("Runtime.evaluate", {
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
    };

    const typeKey = async (key, code, virtualKey, pause = 90) => {
      const text = key.length === 1 ? key : "";
      const common = {
        key,
        code,
        windowsVirtualKeyCode: virtualKey,
        nativeVirtualKeyCode: virtualKey,
      };
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        ...common,
        text,
        unmodifiedText: text,
      });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(pause);
    };

    const typeRawCommand = async (command) => {
      await typeKey("\\", "Backslash", 220, 150);
      for (const character of command) {
        const upper = character.toUpperCase();
        await typeKey(character, `Key${upper}`, upper.charCodeAt(0), 75);
      }
    };

    await evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.onboarding.windows.desktop.v1.1.0.completed", "true");
      localStorage.setItem("visualtex.release-welcome.1.2.6.seen", "true");
      localStorage.setItem("visualtex-desktop-editor-source-open", "false");
      let persisted;
      try { persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "null"); } catch { persisted = null; }
      if (!persisted || typeof persisted !== "object") persisted = { state: {}, version: 0 };
      persisted.state = {
        ...(persisted.state || {}),
        lines: [{ id: "nested-line", latex: "", mode: "display" }],
        activeLineId: "nested-line",
        sourceOpen: false,
        editorLayout: "standard",
        checkUpdatesOnStartup: false,
        inputBehavior: {
          ...(persisted.state?.inputBehavior || {}),
          showStructuredCommandSuggestions: true,
          showOtherCommandSuggestions: true,
          autoExitSuperscript: true,
          autoExitSubscript: true,
          autoExitAccent: true,
          autoExitWrapperCommand: true,
        },
      };
      delete persisted.state.latex;
      localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(850);
    await evaluate(`new Promise((resolve, reject) => {
      const started = performance.now();
      const poll = () => {
        if (document.querySelector("math-field")) return resolve(true);
        if (performance.now() - started > 5000) return reject(new Error("Math field did not mount"));
        setTimeout(poll, 30);
      };
      poll();
    })`);

    const target = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.setValue("\\\\left\\{\\\\left(a+\\\\vec{B}\\\\right)\\\\vec{C}\\\\right\\}", {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
      let target = null;
      for (let offset = 0; offset <= field.lastOffset; offset += 1) {
        const prefix = field.getValue(0, offset, "latex");
        const depth = field.getElementInfo(offset)?.depth ?? -1;
        if (prefix.endsWith("\\\\vec{B}") && (!target || depth > target.depth)) {
          target = { offset, depth };
        }
      }
      if (!target || target.depth < 2) throw new Error("Nested target not found");
      field.position = target.offset;
      field.selection = { ranges: [[target.offset, target.offset]], direction: "none" };
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      return target;
    })()`);
    assert.ok(target.depth >= 2, JSON.stringify(target));
    await sleep(140);
    await typeRawCommand("nabla");
    await typeKey(" ", "Space", 32, 160);
    const afterNabla = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        value: field.value,
        position: field.position,
        depth: field.getElementInfo(field.position)?.depth ?? null,
      };
    })()`);
    assert.match(afterNabla.value, /\\nabla/);
    assert.ok(
      (afterNabla.depth ?? -1) >= 2,
      `Space confirmation escaped the inner delimiter: ${JSON.stringify(afterNabla)}`,
    );
    await typeKey("q", "KeyQ", 81, 140);
    const afterQ = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        value: field.value,
        position: field.position,
        depth: field.getElementInfo(field.position)?.depth ?? null,
      };
    })()`);
    assert.ok(
      (afterQ.depth ?? -1) >= 2,
      `Typing after the accepted command escaped the inner delimiter: ${JSON.stringify(afterQ)}`,
    );
    assert.match(afterQ.value, /\\nabla q?\\right\)/);
    console.log("Nested delimiter command-entry regression passed", JSON.stringify(afterQ));

    await evaluate(`(() => {
      localStorage.setItem("visualtex-desktop-editor-source-open", "true");
      const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      persisted.state = {
        ...(persisted.state || {}),
        lines: [{ id: "ket-line", latex: "", mode: "display" }],
        activeLineId: "ket-line",
        sourceOpen: true,
        editorLayout: "standard",
        latexCodeFormat: "raw",
        checkUpdatesOnStartup: false,
      };
      localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(900);
    await evaluate(`new Promise((resolve, reject) => {
      const started = performance.now();
      const poll = () => {
        if (document.querySelector("math-field") && document.querySelector(".source-panel .cm-content")) return resolve(true);
        if (performance.now() - started > 5000) return reject(new Error("Ket editor/source did not mount"));
        setTimeout(poll, 30);
      };
      poll();
    })`);
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.focus();
      field.position = field.lastOffset;
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      document.querySelector('button[data-category="physics"]')?.click();
    })()`);
    await sleep(180);
    const toolbarCommands = await evaluate(`
      [...document.querySelectorAll('button[data-command-id]')].map((button) => button.dataset.commandId)
    `);
    assert.ok(toolbarCommands.includes("shortcut-ket"), JSON.stringify(toolbarCommands));
    await evaluate(`document.querySelector('button[data-command-id="shortcut-ket"]')?.click()`);
    await sleep(220);
    const inserted = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        value: field.value,
        selection: field.selection,
        selectedLatex: field.selectionIsCollapsed ? "" : field.getValue(field.selection),
        macro: field.dataset.toolbarPackageMacroName ?? "",
        placeholderIndex: field.dataset.toolbarPackageMacroPlaceholderIndex ?? "",
        modelStart: field.dataset.toolbarPackageMacroModelStart ?? "",
      };
    })()`);
    assert.match(inserted.value, /\\ket\{/);
    await typeKey("x", "KeyX", 88, 150);
    await typeKey("y", "KeyY", 89, 180);
    const afterKet = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      return {
        value: field.value,
        source: [...document.querySelectorAll('.source-panel .cm-line')].map((line) => line.textContent ?? "")[0] ?? "",
        persisted: persisted?.state?.lines?.[0]?.latex ?? "",
      };
    })()`);
    assert.match(afterKet.value, /\\ket\{xy\}/, JSON.stringify({ inserted, afterKet }));
    assert.match(afterKet.source, /\\ket\{xy\}/, JSON.stringify(afterKet));
    assert.match(afterKet.persisted, /\\ket\{xy\}/, JSON.stringify(afterKet));
    console.log("Toolbar ket live source synchronization regression passed", JSON.stringify(afterKet));
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(250);
    await rm(profile, {
      recursive: true,
      force: true,
      maxRetries: 6,
      retryDelay: 100,
    }).catch((error) => {
      if (process.platform !== "win32" || error?.code !== "EBUSY") throw error;
    });
  }
}

await main();
