import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const portOffset = process.pid % 700;
const previewPort = 9400 + portOffset;
const debugPort = 14400 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}/editor`;
const chromeProfile = createBrowserProfilePath("visualtex-ideographic-comma");
const chromePath = resolveChromiumExecutable();
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 12000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while the local process starts.
    }
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
        `--user-data-dir=${chromeProfile}`,
        "--window-size=1200,800",
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
    if (!page) throw new Error("No VisualTeX Chrome page target found");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(300);

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

    const reload = async () => {
      await client.send("Page.reload", { ignoreCache: true });
      await sleep(500);
      await evaluate(`new Promise((resolve) => {
        const done = () => document.querySelector("math-field")
          ? resolve(true)
          : setTimeout(done, 25);
        done();
      })`);
      await sleep(120);
    };

    const configure = async () => {
      await evaluate(`(() => {
        localStorage.setItem("visualtex.onboarding.v3.completed", "true");
        const key = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(key) || "{}");
        persisted.state = {
          ...(persisted.state || {}),
          lines: [{ id: "ideographic-comma-line", latex: "" }],
          activeLineId: "ideographic-comma-line",
          sourceOpen: false,
          sidebarOpen: false,
        };
        localStorage.setItem(key, JSON.stringify(persisted));
      })()`);
      await reload();
    };

    const clearAndFocus = async (latex = "") => {
      await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.setValue(${JSON.stringify(latex)}, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.position = field.lastOffset;
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        return true;
      })()`);
      await sleep(80);
    };

    const readField = () =>
      evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          value: field.value,
          raw: Array.from(field.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
        };
      })()`);

    const backslashKey = {
      code: "Backslash",
      windowsVirtualKeyCode: 220,
      nativeVirtualKeyCode: 220,
    };

    // A Chinese input method in punctuation mode reports the physical
    // Backslash key as "、" and commits the character directly.
    const typeDirectIdeographicComma = async () => {
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        key: "、",
        ...backslashKey,
        text: "、",
        unmodifiedText: "、",
      });
      await client.send("Input.dispatchKeyEvent", {
        type: "keyUp",
        key: "、",
        ...backslashKey,
      });
      await sleep(320);
    };

    // Windows input methods report the key as "Process" and commit the text
    // through a separate input event.
    const typeProcessedIdeographicComma = async () => {
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        key: "Process",
        code: "Backslash",
        windowsVirtualKeyCode: 229,
        nativeVirtualKeyCode: 229,
      });
      await client.send("Input.insertText", { text: "、" });
      await client.send("Input.dispatchKeyEvent", {
        type: "keyUp",
        key: "\\",
        ...backslashKey,
      });
      await sleep(320);
    };

    const count = (value, part) => value.split(part).length - 1;

    await configure();

    await clearAndFocus();
    await typeDirectIdeographicComma();
    const direct = await readField();
assert.equal(count(direct.value, "、"), 1, JSON.stringify(direct));
    assert.equal(direct.value.includes("\\\\"), false, JSON.stringify(direct));
    assert.equal(direct.raw, "", JSON.stringify(direct));

    await clearAndFocus("a");
    await typeDirectIdeographicComma();
    await typeDirectIdeographicComma();
    const repeated = await readField();
    assert.equal(count(repeated.value, "、"), 2, JSON.stringify(repeated));
    assert.equal(repeated.value.startsWith("a"), true, JSON.stringify(repeated));

    const typeCharacter = async (key, code, keyCode, pause = 70) => {
      const common = {
        key,
        code,
        windowsVirtualKeyCode: keyCode,
        nativeVirtualKeyCode: keyCode,
      };
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        ...common,
        text: key,
        unmodifiedText: key,
      });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(pause);
    };

    // Switching back to Latin input right after "、" still starts a command.
    await clearAndFocus();
    await typeDirectIdeographicComma();
    await typeCharacter("\\", "Backslash", 220, 80);
    await typeCharacter("p", "KeyP", 80);
    await typeCharacter("i", "KeyI", 73);
    await typeCharacter(" ", "Space", 32, 150);
    const thenCommand = await readField();
    assert.equal(count(thenCommand.value, "、"), 1, JSON.stringify(thenCommand));
    assert.equal(count(thenCommand.value, "\\pi"), 1, JSON.stringify(thenCommand));
    assert.equal(thenCommand.raw, "", JSON.stringify(thenCommand));

    await clearAndFocus();
    await typeProcessedIdeographicComma();
    const processed = await readField();
    assert.equal(count(processed.value, "、"), 1, JSON.stringify(processed));
    assert.equal(processed.raw, "", JSON.stringify(processed));

    console.log("Ideographic comma regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(250);
    await rm(chromeProfile, { recursive: true, force: true });
  }
}

await main();
