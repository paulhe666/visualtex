import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const portOffset = process.pid % 1000;
const previewPort = 6400 + portOffset;
const debugPort = 11400 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = `/tmp/visualtex-input-behavior-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
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
        "--window-size=1400,1000",
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
    await sleep(650);

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
      await sleep(650);
      await evaluate(`new Promise((resolve) => {
        const done = () => document.querySelector("math-field") ? resolve(true) : setTimeout(done, 30);
        done();
      })`);
    };

    const typeCharacter = async (value, code, keyCode) => {
      const common = {
        key: value,
        code,
        windowsVirtualKeyCode: keyCode,
        nativeVirtualKeyCode: keyCode,
      };
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        ...common,
        text: value,
        unmodifiedText: value,
      });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(180);
    };

    const configure = async (overrides = {}) => {
      await evaluate(`(() => {
        localStorage.setItem("visualtex.onboarding.v3.completed", "true");
        localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
        const key = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(key) || "{}");
        persisted.state = {
          ...(persisted.state || {}),
          inputBehavior: {
            autoExitSuperscript: true,
            autoExitSubscript: true,
            autoExitAccent: true,
            ...${JSON.stringify(overrides)},
          },
        };
        localStorage.setItem(key, JSON.stringify(persisted));
      })()`);
      await reload();
    };

    const preparePlaceholder = async (latex) => {
      const state = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        field.setValue(${JSON.stringify(latex)}, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.position = field.lastOffset;
        field.executeCommand("moveToPreviousPlaceholder");
        return {
          value: field.value,
          position: field.position,
          lastOffset: field.lastOffset,
          selection: field.selection,
        };
      })()`);
      assert.notEqual(
        state.position,
        state.lastOffset,
        `Placeholder was not selected for ${latex}`,
      );
    };

    const prepareEmptyField = async () => {
      await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        field.executeCommand("selectAll");
        field.executeCommand("deleteBackward");
        field.position = field.lastOffset;
      })()`);
      await sleep(120);
    };

    const readState = () =>
      evaluate(`(() => {
        const field = document.querySelector("math-field");
        const markers = Array.from(field.shadowRoot?.querySelectorAll(
          ".ML__placeholder-selected, .ML__caret, .ML__selected"
        ) || []);
        const marker = markers.find((candidate) =>
          candidate.closest(".ML__msubsup, .ML__op-group")
        );
        const script = marker?.closest(".ML__msubsup, .ML__op-group");
        const markerBox = (marker?.parentElement || marker)?.getBoundingClientRect();
        const scriptBox = script?.getBoundingClientRect();
        return {
          value: field.value,
          position: field.position,
          lastOffset: field.lastOffset,
          inScript: Boolean(marker && script),
          inAccent: markers.some((candidate) => candidate.closest(".ML__accent-body")),
          markerClass: marker?.className || "",
          markerParentClass: marker?.parentElement?.className || "",
          scriptClass: script?.className || "",
          markerCenter: markerBox ? markerBox.top + markerBox.height / 2 : null,
          scriptCenter: scriptBox ? scriptBox.top + scriptBox.height / 2 : null,
        };
      })()`);

    await configure();
    await preparePlaceholder("x^{\\placeholder{}}");
    await typeCharacter("a", "KeyA", 65);
    const superscript = await readState();
    assert.equal(superscript.value, "x^{a}");
    assert.equal(superscript.position, superscript.lastOffset);
    assert.equal(superscript.inScript, false);

    await preparePlaceholder("x_{\\placeholder{}}");
    await typeCharacter("b", "KeyB", 66);
    const subscript = await readState();
    assert.equal(subscript.value, "x_{b}");
    assert.equal(subscript.position, subscript.lastOffset);
    assert.equal(subscript.inScript, false);

    await preparePlaceholder("\\hat{\\placeholder{}}+z");
    await typeCharacter("c", "KeyC", 67);
    await typeCharacter("d", "KeyD", 68);
    const accent = await readState();
    assert.equal(accent.value, "\\hat{c}d+z");

    await preparePlaceholder("\\vec{\\placeholder{}}+z");
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.dispatchEvent(new CompositionEvent("compositionstart", {
        bubbles: true,
        composed: true,
      }));
      field.insert("m", {
        mode: "math",
        format: "latex",
        insertionMode: "replaceSelection",
        selectionMode: "after",
        focus: true,
        scrollIntoView: false,
      });
      field.dispatchEvent(new CompositionEvent("compositionend", {
        data: "m",
        bubbles: true,
        composed: true,
      }));
    })()`);
    await sleep(120);
    await typeCharacter("n", "KeyN", 78);
    const composedAccent = await readState();
    assert.equal(composedAccent.value, "\\vec{m}n+z");

    await configure({ autoExitSuperscript: false, autoExitSubscript: false });
    await reload();
    await prepareEmptyField();
    await typeCharacter("x", "KeyX", 88);
    await typeCharacter("^", "Digit6", 54);
    await typeCharacter("2", "Digit2", 50);
    const bothDisabledSuperscript = await readState();
    assert.match(bothDisabledSuperscript.value, /^x(?:\^2|\^\{2\})$/);
    assert.equal(bothDisabledSuperscript.inScript, true);
    assert.notEqual(
      bothDisabledSuperscript.position,
      bothDisabledSuperscript.lastOffset,
    );

    await configure({ autoExitSuperscript: false, autoExitSubscript: true });
    await preparePlaceholder("x^{\\placeholder{}}");
    await typeCharacter("d", "KeyD", 68);
    const disabled = await readState();
    assert.equal(disabled.value, "x^{d}");
    assert.equal(disabled.inScript, true);
    assert.notEqual(disabled.position, disabled.lastOffset);

    await reload();
    await prepareEmptyField();
    await typeCharacter("x", "KeyX", 88);
    await typeCharacter("^", "Digit6", 54);
    const emptyUpperState = await readState();
    await typeCharacter("2", "Digit2", 50);
    const independentSuperscript = await readState();
    assert.match(independentSuperscript.value, /^x(?:\^2|\^\{2\})$/);
    assert.equal(
      independentSuperscript.inScript,
      true,
      `Upper script incorrectly followed the subscript switch; before=${JSON.stringify(emptyUpperState)} after=${JSON.stringify(independentSuperscript)}`,
    );
    assert.notEqual(independentSuperscript.position, independentSuperscript.lastOffset);

    await configure({ autoExitSuperscript: true, autoExitSubscript: false });
    await reload();
    await prepareEmptyField();
    await typeCharacter("x", "KeyX", 88);
    await typeCharacter("^", "Digit6", 54);
    await typeCharacter("a", "KeyA", 65);
    const enabledSuperscript = await readState();
    assert.equal(enabledSuperscript.value, "x^{a}");
    assert.equal(enabledSuperscript.inScript, false);
    assert.equal(enabledSuperscript.position, enabledSuperscript.lastOffset);

    await reload();
    await prepareEmptyField();
    await typeCharacter("x", "KeyX", 88);
    await typeCharacter("_", "Minus", 189);
    await typeCharacter("b", "KeyB", 66);
    const independentSubscript = await readState();
    assert.equal(independentSubscript.value, "x_{b}");
    assert.equal(
      independentSubscript.inScript,
      true,
      `Lower script incorrectly followed the superscript switch: ${JSON.stringify(independentSubscript)}`,
    );
    assert.notEqual(independentSubscript.position, independentSubscript.lastOffset);

    const menu = await evaluate(`new Promise((resolve) => {
      const trigger = document.querySelector(".canvas-input-behavior-trigger");
      trigger?.click();
      setTimeout(() => resolve({
        triggerText: trigger?.textContent?.trim() ?? "",
        optionCount: document.querySelectorAll(".input-behavior-option").length,
      }), 50);
    })`);
    assert.match(menu.triggerText, /操作逻辑|Input behavior/);
    assert.equal(menu.optionCount, 6);

    console.log("Input behavior regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(300);
    for (let attempt = 0; attempt < 4; attempt += 1) {
      try {
        await rm(chromeProfile, { recursive: true, force: true });
        break;
      } catch (error) {
        if (attempt === 3) throw error;
        await sleep(150);
      }
    }
  }
}

await main();
