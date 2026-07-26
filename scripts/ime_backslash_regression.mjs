import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 1000;
const previewPort = 8800 + offset;
const debugPort = 13800 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}/editor`;
const profile = `/tmp/visualtex-ime-backslash-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while the process starts.
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
        `--user-data-dir=${profile}`,
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
    if (!page) throw new Error("No VisualTeX target found");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(550);

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

    await evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
      localStorage.setItem("visualtex.onboarding.macos.desktop.v1.2.0.completed", "true");
      localStorage.setItem("visualtex.office.macos.native-first-run.v1.2.0.completed", "true");
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(550);
    await evaluate(`new Promise((resolve) => {
      const poll = () => document.querySelector("math-field")
        ? resolve(true)
        : setTimeout(poll, 25);
      poll();
    })`);

    const initial = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      field.executeCommand("selectAll");
      field.executeCommand("deleteBackward");
      field.position = field.lastOffset;
      return { value: field.value, position: field.position };
    })()`);
    assert.equal(initial.value, "");

    const composition = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      const dispatchBeforeInput = (inputType, data, isComposing) => {
        const event = new InputEvent("beforeinput", {
          inputType,
          data,
          isComposing,
          bubbles: true,
          composed: true,
          cancelable: true,
        });
        const allowed = field.dispatchEvent(event);
        return { allowed, defaultPrevented: event.defaultPrevented };
      };
      const dispatchInput = (inputType, data, isComposing) => {
        field.dispatchEvent(new InputEvent("input", {
          inputType,
          data,
          isComposing,
          bubbles: true,
          composed: true,
        }));
      };

      field.dispatchEvent(new CompositionEvent("compositionstart", {
        data: "",
        bubbles: true,
        composed: true,
      }));
      const insert = dispatchBeforeInput("insertCompositionText", "中", true);
      if (insert.allowed) {
        field.insert("中", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceSelection",
          selectionMode: "after",
          focus: true,
          scrollIntoView: false,
        });
      }
      dispatchInput("insertCompositionText", "中", true);

      const remove = dispatchBeforeInput("deleteCompositionText", "", true);
      // WebKit reports the composition deletion before its model cleanup has
      // fully settled. Dispatch the input notification while the guard is
      // unquestionably composing, then remove the temporary model content.
      dispatchInput("deleteCompositionText", "", true);
      if (remove.allowed) field.executeCommand("deleteBackward");
      field.dispatchEvent(new CompositionEvent("compositionend", {
        data: "",
        bubbles: true,
        composed: true,
      }));

      const stale = dispatchBeforeInput("insertText", "\\\\", false);
      if (stale.allowed) {
        field.insert("\\\\", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceSelection",
          selectionMode: "after",
          focus: true,
          scrollIntoView: false,
        });
        dispatchInput("insertText", "\\\\", false);
      }
      return {
        value: field.value,
        stale,
        raw: Array.from(field.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])
          .filter((node) => !node.classList.contains("ML__suggestion"))
          .map((node) => node.textContent ?? "")
          .join(""),
      };
    })()`);
    assert.equal(composition.value, "", JSON.stringify(composition));
    assert.equal(composition.raw, "", JSON.stringify(composition));
    assert.equal(composition.stale.defaultPrevented, true, JSON.stringify(composition));

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

    const backslashCommon = {
      key: "\\",
      code: "Backslash",
      windowsVirtualKeyCode: 220,
      nativeVirtualKeyCode: 220,
    };
    await client.send("Input.dispatchKeyEvent", {
      type: "keyDown",
      ...backslashCommon,
    });
    await client.send("Input.dispatchKeyEvent", {
      type: "char",
      ...backslashCommon,
      text: "\\",
      unmodifiedText: "\\",
    });
    await client.send("Input.dispatchKeyEvent", {
      type: "char",
      ...backslashCommon,
      text: "\\",
      unmodifiedText: "\\",
    });
    await client.send("Input.dispatchKeyEvent", {
      type: "keyUp",
      ...backslashCommon,
    });
    await sleep(80);
    const duplicateProbe = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        raw: Array.from(field.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])
          .filter((node) => !node.classList.contains("ML__suggestion"))
          .map((node) => node.textContent ?? "")
          .join(""),
      };
    })()`);
    assert.equal(duplicateProbe.raw, "\\", JSON.stringify(duplicateProbe));

    await typeCharacter("p", "KeyP", 80);
    await typeCharacter("i", "KeyI", 73);
    const beforeSpace = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        raw: Array.from(field.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])
          .filter((node) => !node.classList.contains("ML__suggestion"))
          .map((node) => node.textContent ?? "")
          .join(""),
        candidate:
          field.dataset.pendingNativeSuggestion ||
          document.querySelector("#mathlive-suggestion-popover li.ML__popover__current")?.dataset.command ||
          "",
      };
    })()`);
    assert.equal(beforeSpace.raw, "\\pi", JSON.stringify(beforeSpace));
    if (beforeSpace.candidate) assert.equal(beforeSpace.candidate, "\\pi");

    await typeCharacter(" ", "Space", 32);
    const finalState = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        value: field.value,
        raw: Array.from(field.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])
          .filter((node) => !node.classList.contains("ML__suggestion"))
          .map((node) => node.textContent ?? "")
          .join(""),
      };
    })()`);
    assert.equal(finalState.value, "\\pi", JSON.stringify(finalState));
    assert.equal(finalState.raw, "", JSON.stringify(finalState));
    assert.equal(finalState.value.split("\\pi").length - 1, 1);

    console.log("IME backslash regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(250);
    await rm(profile, { recursive: true, force: true });
  }
}

await main();
