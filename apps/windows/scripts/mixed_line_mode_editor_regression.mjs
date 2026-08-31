import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const offset = process.pid % 700;
const previewPort = 22800 + offset;
const debugPort = 23800 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = createBrowserProfilePath("visualtex-mixed-line-mode");
const chromePath = resolveChromiumExecutable();
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15_000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while Vite or Chromium starts.
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

  async evaluate(expression) {
    const result = await this.send("Runtime.evaluate", {
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

  close() {
    this.socket?.close();
  }
}

async function waitForEvaluation(client, expression, description, timeoutMs = 12_000) {
  const startedAt = Date.now();
  let lastValue;
  while (Date.now() - startedAt < timeoutMs) {
    lastValue = await client.evaluate(expression);
    if (lastValue?.ready) return lastValue;
    await sleep(60);
  }
  throw new Error(`Timed out waiting for ${description}: ${JSON.stringify(lastValue)}`);
}

async function dispatchEnter(client, modifiers) {
  const common = {
    key: "Enter",
    code: "Enter",
    modifiers,
    windowsVirtualKeyCode: 13,
    nativeVirtualKeyCode: 13,
  };
  await client.send("Input.dispatchKeyEvent", { type: "keyDown", ...common });
  await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
  await sleep(120);
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
        "--window-size=1400,900",
        baseUrl,
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json();
    const page =
      targets.find(
        (target) => target.type === "page" && target.url.startsWith(baseUrl),
      ) ?? targets.find((target) => target.type === "page");
    assert.ok(page, "VisualTeX page target must exist");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(500);
    await waitForEvaluation(
      client,
      `(() => ({ ready: location.href.startsWith(${JSON.stringify(baseUrl)}) && document.readyState !== "loading" }))()`,
      "VisualTeX same-origin page before localStorage seeding",
    );

    await client.evaluate(`(() => {
      localStorage.clear();
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.onboarding.windows.desktop.v1.1.0.completed", "true");
      localStorage.setItem(
        "visualtex-editor",
        JSON.stringify({
          state: {
            title: "Mixed line mode regression",
            lines: [
              { id: "mixed-inline", latex: "\\\\text{速度}v", mode: "inline" },
              { id: "mixed-display", latex: "\\\\frac{x}{y}", mode: "display" },
            ],
            activeLineId: "mixed-inline",
            latexCodeFormat: "mixed-inline-display",
            formulaAlignment: "left",
            editorLayout: "standard",
            zoom: 0.6,
            language: "cn",
            sourceOpen: false,
          },
          version: 0,
        }),
      );
      return true;
    })()`);
    await client.send("Page.navigate", { url: baseUrl });

    const initial = await waitForEvaluation(
      client,
      `(() => {
        const rows = [...document.querySelectorAll('.formula-line')];
        const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}');
        const modes = persisted?.state?.lines?.map((line) => line.mode) ?? [];
        const toggles = rows.map((row) => ({
          id: row.getAttribute('data-line-id'),
          inlineActive: row.querySelector('.formula-line-mode-toggle button:first-child')?.classList.contains('is-active') ?? false,
          displayActive: row.querySelector('.formula-line-mode-toggle button:last-child')?.classList.contains('is-active') ?? false,
        }));
        return {
          ready: rows.length === 2 && toggles.every((item) => item.inlineActive || item.displayActive),
          modes,
          toggles,
          format: persisted?.state?.latexCodeFormat ?? null,
        };
      })()`,
      "initial mixed row controls",
    );
    assert.equal(initial.format, "mixed-inline-display", JSON.stringify(initial));
    assert.deepEqual(initial.modes, ["inline", "display"], JSON.stringify(initial));
    assert.deepEqual(
      initial.toggles.map((item) => [item.inlineActive, item.displayActive]),
      [[true, false], [false, true]],
      JSON.stringify(initial),
    );

    await client.evaluate(`(() => {
      const row = document.querySelector('.formula-line[data-line-id="mixed-display"]');
      row?.querySelector('.formula-line-mode-toggle button:first-child')?.click();
      return true;
    })()`);
    const toggled = await waitForEvaluation(
      client,
      `(() => {
        const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}');
        const row = document.querySelector('.formula-line[data-line-id="mixed-display"]');
        const mode = persisted?.state?.lines?.find((line) => line.id === 'mixed-display')?.mode;
        return {
          ready: mode === 'inline' && Boolean(row?.querySelector('.formula-line-mode-toggle button:first-child.is-active')),
          mode,
        };
      })()`,
      "manual row mode toggle",
    );
    assert.equal(toggled.mode, "inline");

    await client.evaluate(`(() => {
      const field = document.querySelector('.formula-line[data-line-id="mixed-display"] math-field');
      field.position = field.lastOffset;
      field.selection = { ranges: [[field.lastOffset, field.lastOffset]], direction: 'none' };
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      return true;
    })()`);
    await dispatchEnter(client, 1); // Alt+Enter => display row on Windows.
    const afterAltEnter = await waitForEvaluation(
      client,
      `(() => {
        const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}');
        const lines = persisted?.state?.lines ?? [];
        return {
          ready: lines.length === 3 && lines[2]?.mode === 'display',
          modes: lines.map((line) => line.mode),
          activeLineId: persisted?.state?.activeLineId ?? null,
        };
      })()`,
      "Alt+Enter display row",
    );
    assert.deepEqual(afterAltEnter.modes, ["inline", "inline", "display"]);

    await dispatchEnter(client, 8); // Shift+Enter => inline row.
    const afterShiftEnter = await waitForEvaluation(
      client,
      `(() => {
        const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}');
        const lines = persisted?.state?.lines ?? [];
        return {
          ready: lines.length === 4 && lines[3]?.mode === 'inline',
          modes: lines.map((line) => line.mode),
          activeLineId: persisted?.state?.activeLineId ?? null,
        };
      })()`,
      "Shift+Enter inline row",
    );
    assert.deepEqual(afterShiftEnter.modes, ["inline", "inline", "display", "inline"]);

    await dispatchEnter(client, 0); // Plain Enter inherits inline.
    const inherited = await waitForEvaluation(
      client,
      `(() => {
        const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}');
        const lines = persisted?.state?.lines ?? [];
        return {
          ready: lines.length === 5 && lines[4]?.mode === 'inline',
          modes: lines.map((line) => line.mode),
          activeLineId: persisted?.state?.activeLineId ?? null,
        };
      })()`,
      "plain Enter inherited row mode",
    );
    assert.deepEqual(
      inherited.modes,
      ["inline", "inline", "display", "inline", "inline"],
    );

    await client.evaluate(`(() => {
      const row = document.querySelector('.formula-line[data-line-id="mixed-inline"]');
      row?.querySelector('.formula-line-mode-toggle button:last-child')?.click();
      return true;
    })()`);
    await waitForEvaluation(
      client,
      `(() => {
        const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}');
        return {
          ready: persisted?.state?.lines?.[0]?.mode === 'display',
          modes: persisted?.state?.lines?.map((line) => line.mode) ?? [],
        };
      })()`,
      "display toggle persisted",
    );

    await client.send("Page.navigate", { url: baseUrl });
    const reopened = await waitForEvaluation(
      client,
      `(() => {
        const rows = [...document.querySelectorAll('.formula-line')];
        const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}');
        const modes = persisted?.state?.lines?.map((line) => line.mode) ?? [];
        const activeButtons = rows.map((row) => ({
          inline: row.querySelector('.formula-line-mode-toggle button:first-child')?.classList.contains('is-active') ?? false,
          display: row.querySelector('.formula-line-mode-toggle button:last-child')?.classList.contains('is-active') ?? false,
        }));
        return {
          ready: rows.length === 5 && activeButtons.every((item) => item.inline || item.display),
          modes,
          activeButtons,
        };
      })()`,
      "mixed modes after reload",
    );
    assert.deepEqual(
      reopened.modes,
      ["display", "inline", "display", "inline", "inline"],
      JSON.stringify(reopened),
    );
    assert.deepEqual(
      reopened.activeButtons.map((item) => [item.inline, item.display]),
      [[false, true], [true, false], [false, true], [true, false], [true, false]],
      JSON.stringify(reopened),
    );

    console.log(
      "Windows mixed inline/display editor regression passed: manual toggles, Alt/Shift/plain Enter and reload persistence",
    );
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(180);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => {});
  }
}

await main();
