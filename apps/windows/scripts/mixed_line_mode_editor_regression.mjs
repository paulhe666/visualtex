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
  await dispatchKey(client, "Enter", "Enter", 13, modifiers);
}

async function dispatchKey(client, key, code, virtualKeyCode, modifiers) {
  const common = {
    key,
    code,
    modifiers,
    windowsVirtualKeyCode: virtualKeyCode,
    nativeVirtualKeyCode: virtualKeyCode,
  };
  await client.send("Input.dispatchKeyEvent", { type: "keyDown", ...common });
  await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
  await sleep(120);
}

function focusRowEnd(client, lineId) {
  return client.evaluate(`(() => {
    const field = document.querySelector('.formula-line[data-line-id="${lineId}"] math-field');
    field.position = field.lastOffset;
    field.selection = { ranges: [[field.lastOffset, field.lastOffset]], direction: 'none' };
    field.focus();
    field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
    return true;
  })()`);
}

function rowStateExpression(lineId, ready) {
  return `(() => {
    const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}');
    const lines = persisted?.state?.lines ?? [];
    const row = lines.find((line) => line.id === ${JSON.stringify(lineId)});
    const state = {
      count: lines.length,
      modes: lines.map((line) => line.mode),
      mode: row?.mode ?? null,
      latex: row?.latex ?? null,
      marker:
        document.querySelector('.formula-line[data-line-id="${lineId}"] [data-formula-line-mode-toggle]')?.getAttribute('data-formula-line-mode') ?? null,
    };
    return { ...state, ready: ${ready} };
  })()`;
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
        const toggles = rows.map((row) => {
          const button = row.querySelector('[data-formula-line-mode-toggle]');
          return {
            id: row.getAttribute('data-line-id'),
            mode: button?.getAttribute('data-formula-line-mode') ?? null,
            text: button?.textContent?.trim() ?? '',
          };
        });
        return {
          ready: rows.length === 2 && toggles.every((item) => item.mode === 'inline' || item.mode === 'display'),
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
      initial.toggles.map((item) => [item.mode, item.text]),
      [["inline", "$"], ["display", "$$"]],
      JSON.stringify(initial),
    );

    await client.evaluate(`(() => {
      const row = document.querySelector('.formula-line[data-line-id="mixed-display"]');
      row?.querySelector('[data-formula-line-mode-toggle]')?.click();
      return true;
    })()`);
    const toggled = await waitForEvaluation(
      client,
      `(() => {
        const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}');
        const row = document.querySelector('.formula-line[data-line-id="mixed-display"]');
        const mode = persisted?.state?.lines?.find((line) => line.id === 'mixed-display')?.mode;
        return {
          ready:
            mode === 'inline' &&
            row?.querySelector('[data-formula-line-mode-toggle]')?.getAttribute('data-formula-line-mode') === 'inline',
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
    await dispatchEnter(client, 1); // Alt+Enter forces a display row.
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
      "Alt+Enter forces display row",
    );
    assert.deepEqual(afterAltEnter.modes, ["inline", "inline", "display"]);

    await client.evaluate(`(() => {
      const field = document.querySelector('.formula-line[data-line-id="mixed-inline"] math-field');
      field.position = field.lastOffset;
      field.selection = { ranges: [[field.lastOffset, field.lastOffset]], direction: 'none' };
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      return true;
    })()`);
    await dispatchEnter(client, 8); // Shift+Enter => internal multiline environment.
    const afterShiftEnter = await waitForEvaluation(
      client,
      `(() => {
        const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}');
        const lines = persisted?.state?.lines ?? [];
        const first = lines[0];
        return {
          ready:
            lines.length === 3 &&
            /\\\\begin\\{gathered\\}/.test(first?.latex ?? ''),
          modes: lines.map((line) => line.mode),
          firstLatex: first?.latex ?? '',
          firstMode:
            document.querySelector('.formula-line[data-line-id="mixed-inline"] [data-formula-line-mode-toggle]')?.getAttribute('data-formula-line-mode') ?? null,
          firstToggleDisabled:
            document.querySelector('.formula-line[data-line-id="mixed-inline"] [data-formula-line-mode-toggle]')?.disabled === true,
        };
      })()`,
      "Shift+Enter internal gathered row",
    );
    assert.deepEqual(afterShiftEnter.modes, ["display", "inline", "display"]);
    assert.equal(afterShiftEnter.firstMode, "display", JSON.stringify(afterShiftEnter));
    assert.equal(afterShiftEnter.firstToggleDisabled, true, JSON.stringify(afterShiftEnter));

    await client.evaluate(`(() => {
      const field = document.querySelector('.formula-line[data-line-id="mixed-display"] math-field');
      field.position = field.lastOffset;
      field.selection = { ranges: [[field.lastOffset, field.lastOffset]], direction: 'none' };
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      return true;
    })()`);
    await dispatchEnter(client, 0); // Plain Enter inherits the row mode.
    const inherited = await waitForEvaluation(
      client,
      `(() => {
        const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}');
        const lines = persisted?.state?.lines ?? [];
        return {
          ready: lines.length === 4 && lines[2]?.mode === 'inline',
          modes: lines.map((line) => line.mode),
          activeLineId: persisted?.state?.activeLineId ?? null,
        };
      })()`,
      "plain Enter inherited row mode",
    );
    assert.deepEqual(inherited.modes, ["display", "inline", "inline", "display"]);

    await client.evaluate(`(() => {
      const row = document.querySelector('.formula-line[data-line-id="mixed-display"]');
      row?.querySelector('[data-formula-line-mode-toggle]')?.click();
      return true;
    })()`);
    await waitForEvaluation(
      client,
      `(() => {
        const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}');
        return {
          ready:
            persisted?.state?.lines?.find((line) => line.id === 'mixed-display')?.mode === 'display',
          modes: persisted?.state?.lines?.map((line) => line.mode) ?? [],
        };
      })()`,
      "display toggle persisted",
    );

    // Ctrl+Shift+Enter switches the focused row in place: same row, same LaTeX,
    // one undo step; a multi-line row stays display and nothing is split.
    await focusRowEnd(client, "mixed-display");
    await dispatchEnter(client, 2 | 8);
    const shortcutInline = await waitForEvaluation(
      client,
      rowStateExpression("mixed-display", "state.mode === 'inline' && state.marker === 'inline'"),
      "Ctrl+Shift+Enter switches the row to inline",
    );
    assert.equal(shortcutInline.count, 4, JSON.stringify(shortcutInline));
    assert.equal(shortcutInline.latex, "\\frac{x}{y}", JSON.stringify(shortcutInline));

    await dispatchKey(client, "z", "KeyZ", 90, 2);
    await waitForEvaluation(
      client,
      rowStateExpression("mixed-display", "state.mode === 'display' && state.marker === 'display'"),
      "Ctrl+Z restores the display row",
    );
    await dispatchKey(client, "y", "KeyY", 89, 2);
    await waitForEvaluation(
      client,
      rowStateExpression("mixed-display", "state.mode === 'inline' && state.marker === 'inline'"),
      "Ctrl+Y switches the row to inline again",
    );
    await dispatchEnter(client, 2 | 8);
    const shortcutDisplay = await waitForEvaluation(
      client,
      rowStateExpression("mixed-display", "state.mode === 'display' && state.marker === 'display'"),
      "Ctrl+Shift+Enter switches the row back to display",
    );
    assert.deepEqual(shortcutDisplay.modes, ["display", "display", "inline", "display"], JSON.stringify(shortcutDisplay));

    await focusRowEnd(client, "mixed-inline");
    await dispatchEnter(client, 2 | 8);
    await sleep(300);
    const multilineShortcut = await client.evaluate(rowStateExpression("mixed-inline", "true"));
    assert.equal(multilineShortcut.count, 4, JSON.stringify(multilineShortcut));
    assert.equal(multilineShortcut.mode, "display", JSON.stringify(multilineShortcut));
    assert.equal(multilineShortcut.marker, "display", JSON.stringify(multilineShortcut));
    assert.match(multilineShortcut.latex ?? "", /^\\begin\{gathered\}/, JSON.stringify(multilineShortcut));

    await client.send("Page.navigate", { url: baseUrl });
    const reopened = await waitForEvaluation(
      client,
      `(() => {
        const rows = [...document.querySelectorAll('.formula-line')];
        const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}');
        const modes = persisted?.state?.lines?.map((line) => line.mode) ?? [];
        const toggles = rows.map((row) => ({
          mode: row.querySelector('[data-formula-line-mode-toggle]')?.getAttribute('data-formula-line-mode') ?? null,
          text: row.querySelector('[data-formula-line-mode-toggle]')?.textContent?.trim() ?? '',
        }));
        return {
          ready: rows.length === 4 && toggles.every((item) => item.mode === 'inline' || item.mode === 'display'),
          modes,
          toggles,
        };
      })()`,
      "mixed modes after reload",
    );
    assert.deepEqual(
      reopened.modes,
      ["display", "display", "inline", "display"],
      JSON.stringify(reopened),
    );
    assert.deepEqual(
      reopened.toggles.map((item) => [item.mode, item.text]),
      [["display", "$$"], ["display", "$$"], ["inline", "$"], ["display", "$$"]],
      JSON.stringify(reopened),
    );

    // The three row shortcuts live in Settings -> 快捷键 -> 公式行 and can be
    // re-recorded there. Switch rows with Alt+Shift+Enter instead, check that
    // the old chord no longer switches, then restore the default.
    const openRowHotkeys = async () => {
      await client.evaluate(`document.querySelector("button.settings-toggle").click()`);
      await waitForEvaluation(
        client,
        `(() => ({ ready: Boolean(document.querySelector("[data-shortcut-settings-trigger]")) }))()`,
        "settings dialog",
      );
      await client.evaluate(`document.querySelector("[data-shortcut-settings-trigger]").click()`);
      return waitForEvaluation(
        client,
        `(() => {
          const rows = [...document.querySelectorAll("[data-formula-row-hotkey-action]")].map((row) => ({
            id: row.getAttribute("data-formula-row-hotkey-action"),
            chord: row.querySelector("kbd")?.textContent ?? null,
          }));
          return { ready: rows.length === 3, rows };
        })()`,
        "formula row hotkeys in the hotkey manager",
      );
    };
    const closeHotkeyManager = async () => {
      await client.evaluate(`[...document.querySelectorAll(".formula-hotkey-manager-dialog button")].find((button) => button.textContent.trim() === "完成").click()`);
      await waitForEvaluation(
        client,
        `(() => ({ ready: !document.querySelector(".formula-hotkey-manager-dialog") }))()`,
        "hotkey manager closed",
      );
    };
    const defaults = await openRowHotkeys();
    assert.deepEqual(
      defaults.rows,
      [
        { id: "row-new-inline", chord: "Ctrl+Enter" },
        { id: "row-new-display", chord: "Alt+Enter" },
        { id: "row-toggle-mode", chord: "Ctrl+Shift+Enter" },
      ],
      JSON.stringify(defaults),
    );
    await client.evaluate(`document.querySelector('[data-formula-row-hotkey-action="row-toggle-mode"] .formula-hotkey-binding-actions button').click()`);
    await waitForEvaluation(
      client,
      `(() => ({ ready: Boolean(document.querySelector(".formula-row-hotkey-recorder-dialog")) }))()`,
      "formula row hotkey recorder",
    );
    await dispatchEnter(client, 1 | 8);
    await waitForEvaluation(
      client,
      `(() => ({ ready: document.querySelector(".formula-row-hotkey-recorder-dialog kbd")?.textContent === "Alt+Shift+Enter" }))()`,
      "recorded Alt+Shift+Enter",
    );
    await client.evaluate(`[...document.querySelectorAll(".formula-row-hotkey-recorder-dialog button")].find((button) => button.textContent.trim() === "保存").click()`);
    await waitForEvaluation(
      client,
      `(() => {
        const persisted = JSON.parse(localStorage.getItem("visualtex-formula-row-hotkeys-v1") || "{}");
        const toggle = persisted?.state?.bindings?.find((binding) => binding.actionId === "row-toggle-mode")?.chord;
        return {
          ready: Boolean(toggle && toggle.code === "Enter" && toggle.altKey && toggle.shiftKey && !toggle.ctrlKey && !toggle.metaKey),
          toggle,
        };
      })()`,
      "custom row switch persisted",
    );
    await closeHotkeyManager();

    // Without a row binding, Ctrl+Shift+Enter must no longer switch the row's
    // mode; undo whatever the unbound chord did otherwise.
    await focusRowEnd(client, "mixed-display");
    await dispatchEnter(client, 2 | 8);
    await sleep(300);
    const oldChord = await client.evaluate(rowStateExpression("mixed-display", "true"));
    assert.equal(oldChord.mode, "display", JSON.stringify(oldChord));
    assert.equal(oldChord.marker, "display", JSON.stringify(oldChord));
    if (oldChord.count !== 4) {
      await dispatchKey(client, "z", "KeyZ", 90, 2);
      await waitForEvaluation(
        client,
        rowStateExpression("mixed-display", "state.count === 4 && state.mode === 'display'"),
        "Ctrl+Z removes the row split by the unbound chord",
      );
    }
    await focusRowEnd(client, "mixed-display");
    await dispatchEnter(client, 1 | 8);
    await waitForEvaluation(
      client,
      rowStateExpression("mixed-display", "state.mode === 'inline' && state.marker === 'inline' && state.count === 4"),
      "Alt+Shift+Enter switches the row after re-recording",
    );
    const tooltip = await client.evaluate(`document.querySelector('.formula-line[data-line-id="mixed-display"] [data-formula-line-mode-toggle]')?.getAttribute("title") ?? ""`);
    assert.match(tooltip, /Alt\+Shift\+Enter/, tooltip);

    const restored = await openRowHotkeys();
    assert.equal(restored.rows[2].chord, "Alt+Shift+Enter", JSON.stringify(restored));
    await client.evaluate(`document.querySelector('[data-formula-row-hotkey-action="row-toggle-mode"] button[title="恢复默认"]').click()`);
    await waitForEvaluation(
      client,
      `(() => ({ ready: document.querySelector('[data-formula-row-hotkey-action="row-toggle-mode"] kbd')?.textContent === "Ctrl+Shift+Enter" }))()`,
      "row switch restored to Ctrl+Shift+Enter",
    );
    await closeHotkeyManager();

    console.log(
      "Windows universal row-mode regression passed: single-symbol row toggles, forced Alt+Enter display, Shift+Enter multiline, Enter inheritance, Ctrl+Shift+Enter row switch with undo/redo, reload persistence and re-recorded row shortcuts",
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
