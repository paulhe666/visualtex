import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import process from "node:process";

const portOffset = process.pid % 1000;
const previewPort = 5300 + portOffset;
const debugPort = 10300 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}/editor?visualtex-test-probe=1`;
const chromeProfile = path.join(
  os.tmpdir(),
  `visualtex-editor-smoke-${process.pid}`,
);
const chromeCandidates =
  process.platform === "win32"
    ? [
        "C:/Program Files/Google/Chrome/Application/chrome.exe",
        "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
      ]
    : process.platform === "darwin"
      ? ["/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"]
      : ["/usr/bin/google-chrome", "/usr/bin/chromium", "/usr/bin/chromium-browser"];
const chromePath =
  process.env.VISUALTEX_CHROME_PATH ??
  chromeCandidates.find((candidate) => existsSync(candidate));
if (!chromePath) {
  throw new Error(
    "Google Chrome was not found. Set VISUALTEX_CHROME_PATH to run the editor regression smoke test.",
  );
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return response;
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
    let page;
    const targetStarted = Date.now();
    while (!page && Date.now() - targetStarted < 10000) {
      const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
      page = targets.find(
        (target) => target.type === "page" && target.url.startsWith(baseUrl),
      );
      if (!page) await sleep(100);
    }
    if (!page) throw new Error("No VisualTeX Chrome page target found");

    client = new CdpClient(page.webSocketDebuggerUrl);
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
        const description = result.exceptionDetails.exception?.description;
        throw new Error(
          description || result.exceptionDetails.text || "Runtime.evaluate failed",
        );
      }
      return result.result.value;
    };

    const key = async (value, code, virtualKeyCode) => {
      const common = {
        key: value,
        code,
        windowsVirtualKeyCode: virtualKeyCode,
        nativeVirtualKeyCode: virtualKeyCode,
      };
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        ...common,
        ...(value.length === 1 ? { text: value, unmodifiedText: value } : {}),
      });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(35);
    };

    const imeKey = async (value, code, virtualKeyCode, committedText) => {
      const common = {
        key: value,
        code,
        windowsVirtualKeyCode: virtualKeyCode,
        nativeVirtualKeyCode: virtualKeyCode,
      };
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        ...common,
      });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await client.send("Input.insertText", { text: committedText });
      await sleep(120);
    };

    const waitForEvaluation = async (
      expression,
      description,
      timeoutMs = 5000,
    ) => {
      const started = Date.now();
      let lastValue;
      while (Date.now() - started < timeoutMs) {
        lastValue = await evaluate(expression);
        if (lastValue?.ready) return lastValue;
        await sleep(50);
      }
      throw new Error(
        `Timed out waiting for ${description}: ${JSON.stringify(lastValue)}`,
      );
    };

    const reloadEditor = async () => {
      await client.send("Page.reload", { ignoreCache: true });
      await sleep(700);
      await evaluate(`new Promise((resolve) => {
        const done = () => document.querySelector("math-field") ? resolve(true) : setTimeout(done, 30);
        done();
      })`);
    };

    const resetEditorDocument = async (values = [""], activeIndex = 0) => {
      await evaluate(`(() => {
        const storageKey = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
        const values = ${JSON.stringify(values)};
        const lines = values.map((latex) => ({
          id: crypto.randomUUID(),
          latex,
        }));
        const activeIndex = Math.max(
          0,
          Math.min(${activeIndex}, Math.max(0, lines.length - 1)),
        );
        persisted.state = {
          ...(persisted.state || {}),
          lines,
          activeLineId: lines[activeIndex]?.id ?? null,
        };
        localStorage.setItem(storageKey, JSON.stringify(persisted));
      })()`);
      await reloadEditor();
    };

    const replaceFocusedText = async (text) => {
      await evaluate(`(() => {
        const content = document.querySelector(".source-panel .cm-content");
        if (!content) throw new Error("Source editor is unavailable");
        content.focus();
        const selection = window.getSelection();
        const range = document.createRange();
        range.selectNodeContents(content);
        selection.removeAllRanges();
        selection.addRange(range);
      })()`);
      await client.send("Input.insertText", { text });
      await sleep(220);
    };

    await client.send("Page.navigate", { url: baseUrl });
    await sleep(650);
    await evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
      const storageKey = "visualtex-editor";
      const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
      persisted.state = {
        ...(persisted.state || {}),
        inputBehavior: {
          ...((persisted.state || {}).inputBehavior || {}),
          showStructuredCommandSuggestions: true,
          showOtherCommandSuggestions: true,
        },
      };
      localStorage.setItem(storageKey, JSON.stringify(persisted));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(800);
    await evaluate(`new Promise((resolve) => {
      const done = () => document.querySelector("math-field") ? resolve(true) : setTimeout(done, 30);
      done();
    })`);

    const originalWebGeometry = await evaluate(`(() => {
      const workspace = document.querySelector(".workspace.is-classic-layout");
      const tileToolbar = document.querySelector(".classic-tile-toolbar");
      const editorBody = document.querySelector(".classic-editor-pane-body");
      const bottomDock = document.querySelector(".classic-bottom-dock");
      const tileResizer = document.querySelector(".classic-tile-resizer");
      const dockResizer = document.querySelector(".classic-dock-resizer");
      if (!workspace || !tileToolbar || !editorBody || !bottomDock) {
        return { ready: false };
      }
      const workspaceStyle = getComputedStyle(workspace);
      const editorBodyStyle = getComputedStyle(editorBody);
      const tileStyle = tileResizer ? getComputedStyle(tileResizer) : null;
      const dockStyle = dockResizer ? getComputedStyle(dockResizer) : null;
      return {
        ready: true,
        gridColumns: workspaceStyle.gridTemplateColumns,
        gridRows: editorBodyStyle.gridTemplateRows,
        tileWidth: tileToolbar.getBoundingClientRect().width,
        dockHeight: bottomDock.getBoundingClientRect().height,
        tileResizerPosition: tileStyle?.position ?? "",
        dockResizerPosition: dockStyle?.position ?? "",
        tileResizerWidth: tileResizer?.getBoundingClientRect().width ?? 0,
        dockResizerHeight: dockResizer?.getBoundingClientRect().height ?? 0,
      };
    })()`);
    if (!originalWebGeometry.ready) {
      throw new Error(`Classic Web layout did not mount: ${JSON.stringify(originalWebGeometry)}`);
    }
    const columnTracks = originalWebGeometry.gridColumns
      .split(/\s+/)
      .filter(Boolean);
    const rowTracks = originalWebGeometry.gridRows
      .split(/\s+/)
      .filter(Boolean);
    if (
      columnTracks.length !== 3 ||
      Math.abs(originalWebGeometry.tileWidth - 300) > 1.5 ||
      rowTracks.length !== 3 ||
      Math.abs(originalWebGeometry.dockHeight - 240) > 1.5 ||
      originalWebGeometry.tileResizerPosition !== "relative" ||
      originalWebGeometry.dockResizerPosition !== "relative" ||
      Math.abs(originalWebGeometry.tileResizerWidth - 7) > 1.5 ||
      Math.abs(originalWebGeometry.dockResizerHeight - 7) > 1.5
    ) {
      throw new Error(
        `Classic layout does not match the migrated editor geometry: ${JSON.stringify(originalWebGeometry)}`,
      );
    }

    await evaluate(`document.querySelector(".settings-toggle")?.click()`);
    await waitForEvaluation(
      `(() => ({ ready: Boolean(document.querySelector(".settings-dialog")) }))()`,
      "original Web settings dialog",
    );
    await sleep(220);
    const originalWebSettingsGeometry = await evaluate(`(() => {
      const dialog = document.querySelector(".settings-dialog");
      if (!dialog) return { ready: false };
      const rect = dialog.getBoundingClientRect();
      return {
        ready: true,
        width: rect.width,
        usesOriginalDialog: dialog.classList.contains("settings-dialog"),
        hasWindowsStyleNavigation: Boolean(document.querySelector(".web-settings-nav")),
        sectionCount: dialog.querySelectorAll(".settings-section").length,
        viewportWidth: window.innerWidth,
      };
    })()`);
    if (
      !originalWebSettingsGeometry.usesOriginalDialog ||
      originalWebSettingsGeometry.hasWindowsStyleNavigation ||
      Math.abs(originalWebSettingsGeometry.width - 720) > 1.5 ||
      originalWebSettingsGeometry.sectionCount < 7
    ) {
      throw new Error(
        `Settings dialog does not match the migrated editor: ${JSON.stringify(originalWebSettingsGeometry)}`,
      );
    }
    await evaluate(`document.querySelector(".settings-dialog .dialog-header .icon-button")?.click()`);
    await sleep(100);

    const setFieldAt = async (index, latex) => {
      await evaluate(`(() => {
        const field = document.querySelectorAll("math-field")[${index}];
        if (!field) throw new Error("Formula field ${index} was not found");
        field.focus();
        field.shadowRoot
          ?.querySelector('[part="keyboard-sink"]')
          ?.focus({ preventScroll: true });
        field.setValue(${JSON.stringify(latex)}, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.position = field.lastOffset;
        field.dispatchEvent(new InputEvent("input", {
          bubbles: true,
          composed: true,
          inputType: "insertText",
        }));
        return field.value;
      })()`);
      await sleep(180);
    };

    const setField = async (latex) => setFieldAt(0, latex);

    await setField("\\theta");
    await waitForEvaluation(
      `(() => ({
        ready: Boolean(document.querySelector(".suggestion-popup")),
        candidateVisible: Boolean(document.querySelector(".suggestion-popup")),
      }))()`,
      "custom command candidate for \\theta",
    );
    let thetaState;
    let thetaCommitError;
    for (let attempt = 0; attempt < 3; attempt += 1) {
      await key("Enter", "Enter", 13);
      try {
        thetaState = await waitForEvaluation(`(() => {
      const field = document.querySelector("math-field");
      const surface = document.querySelector(".multi-line-editor");
      const active = document.activeElement;
      const state = {
        value: field.value,
        selection: field.selection,
        position: field.position,
        lineCount: document.querySelectorAll(".formula-line").length,
        candidateVisible: Boolean(document.querySelector(".suggestion-popup")),
        candidateCount: document.querySelectorAll(".suggestion-popup").length,
        candidateConnected: document.querySelector(".suggestion-popup")?.isConnected ?? false,
        candidateParentClass: document.querySelector(".suggestion-popup")?.parentElement?.className ?? "",
        candidateParentQuery: document.querySelector(".suggestion-popup")?.parentElement?.dataset.commandQuery ?? "",
        candidateText: document.querySelector(".suggestion-popup")?.innerText ?? "",
        commandQuery: surface?.dataset.commandQuery ?? "",
        activeLineId: surface?.dataset.activeLineId ?? "",
        activeTag: active?.tagName ?? "",
        activeClass: active?.className ?? "",
      };
      return {
        ...state,
        ready: state.value === "\\\\theta" && !state.candidateVisible,
      };
        })()`, "custom command candidate commit", 900);
        break;
      } catch (error) {
        thetaCommitError = error;
      }
    }
    if (!thetaState) throw thetaCommitError;
    if (thetaState.candidateVisible) {
      throw new Error(`Command candidate remained open after committing \\theta: ${JSON.stringify(thetaState)}`);
    }

    await evaluate(`(() => {
      const storageKey = "visualtex-editor";
      const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
      persisted.state = {
        ...(persisted.state || {}),
        inputBehavior: {
          ...(persisted.state?.inputBehavior || {}),
          showOtherCommandSuggestions: false,
        },
      };
      localStorage.setItem(storageKey, JSON.stringify(persisted));
      location.reload();
    })()`);
    await waitForEvaluation(
      `(() => ({ ready: Boolean(document.querySelector("math-field")) }))()`,
      "formula field after isolating native recommendation test",
    );

    await setField("");
    await key("\\", "Backslash", 220);
    await key("t", "KeyT", 84);
    await key("h", "KeyH", 72);
    await key("e", "KeyE", 69);
    await key("t", "KeyT", 84);
    await key("a", "KeyA", 65);
    const nativePopover = await waitForEvaluation(`(() => {
      const panel = document.getElementById("mathlive-suggestion-popover");
      if (!panel) return { ready: false, visible: false };
      const style = getComputedStyle(panel);
      const visible = panel.classList.contains("is-visible") && style.display !== "none";
      return {
        ready: visible && Boolean(panel.querySelector("li.ML__popover__current")),
        visible,
        background: style.backgroundColor,
        transition: style.transitionDuration,
        animation: style.animationDuration,
      };
    })()`, "MathLive recommendation popover");
    if (!nativePopover?.visible) {
      throw new Error("MathLive recommendation popover is not visible while typing a command");
    }
    if (nativePopover.background === "rgb(97, 97, 97)") {
      throw new Error("MathLive recommendation popover still uses the old black/gray background");
    }

    const nativeBeforeArrow = await waitForEvaluation(`(() => {
      const panel = document.getElementById("mathlive-suggestion-popover");
      const selected = panel?.querySelector("li.ML__popover__current");
      if (!panel || !selected) return { ready: false };
      window.__visualtexStableNativePanel = panel;
      const style = getComputedStyle(selected);
      return {
        ready: true,
        command: selected?.dataset.command,
        value: document.querySelector("math-field").value,
        background: style.backgroundColor,
        border: style.borderColor,
        color: style.color,
      };
    })()`, "selected MathLive recommendation");
    if (nativeBeforeArrow.background === "rgb(31, 99, 142)") {
      throw new Error("Selected native recommendation still uses a solid dark-blue fill");
    }
    if (nativeBeforeArrow.border === "rgba(0, 0, 0, 0)") {
      throw new Error("Selected native recommendation has no visible selection border");
    }

    await key("ArrowDown", "ArrowDown", 40);
    const nativeAfterArrow = await waitForEvaluation(`(() => {
      const panel = document.getElementById("mathlive-suggestion-popover");
      const selected = panel?.querySelector("li.ML__popover__current");
      if (!panel || !selected) return { ready: false };
      const style = getComputedStyle(selected);
      const samePanelNode = panel === window.__visualtexStableNativePanel;
      const command = selected?.dataset.command;
      const visible = panel.classList.contains("is-visible") && style.display !== "none";
      return {
        ready: visible && command === "\\\\thetasym",
        samePanelNode,
        visible,
        command,
        background: style.backgroundColor,
        border: style.borderColor,
        color: style.color,
      };
    })()`, "MathLive recommendation ArrowDown selection");
    if (!nativeAfterArrow.visible) {
      throw new Error("Arrow navigation dismissed the native recommendation panel");
    }
    if (nativeAfterArrow.command !== "\\thetasym") {
      throw new Error(`ArrowDown did not move the native recommendation selection: ${JSON.stringify(nativeAfterArrow)}`);
    }
    if (nativeAfterArrow.background === "rgb(31, 99, 142)") {
      throw new Error("Moved native recommendation still uses a solid dark-blue fill");
    }

    await key("Enter", "Enter", 13);
    const nativeCommitState = await waitForEvaluation(`(() => {
      const state = {
        value: document.querySelector("math-field").value,
        lineCount: document.querySelectorAll(".formula-line").length,
        nativeVisible: document.getElementById("mathlive-suggestion-popover")?.classList.contains("is-visible") ?? false,
        candidateVisible: Boolean(document.querySelector(".suggestion-popup")),
      };
      return {
        ...state,
        ready: state.lineCount === 1 && state.value.endsWith("thetasym") && !state.nativeVisible && !state.candidateVisible,
      };
    })()`, "native MathLive recommendation commit");
    if (nativeCommitState.lineCount !== 1 || !nativeCommitState.value.endsWith("thetasym")) {
      throw new Error(`Enter did not commit the selected native MathLive recommendation: ${JSON.stringify(nativeCommitState)}`);
    }
    if (nativeCommitState.nativeVisible || nativeCommitState.candidateVisible) {
      throw new Error(`Recommendation remained visible after commit: ${JSON.stringify(nativeCommitState)}`);
    }

    await waitForEvaluation(
      `(() => {
        const button = document.querySelector('button[aria-label="撤销"]');
        return { ready: Boolean(button && !button.disabled), disabled: button?.disabled ?? true };
      })()`,
      "native candidate undo button",
    );
    await evaluate(`document.querySelector('button[aria-label="撤销"]').click()`);
    const nativeUndoState = await waitForEvaluation(
      `(() => {
        const value = document.querySelector("math-field").value;
        const redo = document.querySelector('button[aria-label="重做"]');
        return {
          ready: value.trim() === ${JSON.stringify(nativeBeforeArrow.value.trim())} && Boolean(redo && !redo.disabled),
          value,
          expectedValue: ${JSON.stringify(nativeBeforeArrow.value)},
          redoDisabled: redo?.disabled ?? true,
        };
      })()`,
      "native candidate undo replay",
    );
    const nativeUndoValue = nativeUndoState.value;
    if (nativeUndoValue.trim() !== nativeBeforeArrow.value.trim()) {
      throw new Error(`Global undo did not restore the native candidate input: ${JSON.stringify({ nativeBeforeArrow, nativeUndoValue })}`);
    }
    await evaluate(`document.querySelector('button[aria-label="重做"]').click()`);
    const nativeRedoState = await waitForEvaluation(
      `(() => {
        const value = document.querySelector("math-field").value;
        return {
          ready: value.trim() === ${JSON.stringify(nativeCommitState.value.trim())},
          value,
        };
      })()`,
      "native candidate redo replay",
    );
    const nativeRedoValue = nativeRedoState.value;
    if (nativeRedoValue.trim() !== nativeCommitState.value.trim()) {
      throw new Error(`Global redo did not restore the native candidate result: ${JSON.stringify({ nativeCommitState, nativeRedoValue })}`);
    }

    await evaluate(`(() => {
      const storageKey = "visualtex-editor";
      const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
      persisted.state = {
        ...(persisted.state || {}),
        inputBehavior: {
          ...(persisted.state?.inputBehavior || {}),
          showOtherCommandSuggestions: true,
        },
      };
      localStorage.setItem(storageKey, JSON.stringify(persisted));
    })()`);
    await reloadEditor();
    await setField("\\alpha");
    await waitForEvaluation(
      `(() => ({
        ready: Boolean(document.querySelector(".suggestion-popup")),
      }))()`,
      "custom command candidate for \\alpha",
    );
    await key("Enter", "Enter", 13);
    await waitForEvaluation(`(() => ({
      ready:
        document.querySelector("math-field")?.value === "\\\\alpha" &&
        !document.querySelector(".suggestion-popup"),
      value: document.querySelector("math-field")?.value ?? "",
    }))()`, "commit custom alpha candidate");
    await key("Enter", "Enter", 13);
    const enterAfterCandidateState = await waitForEvaluation(`(() => {
      const fields = [...document.querySelectorAll("math-field")];
      return {
        ready: fields.length === 2,
        lineCount: fields.length,
        values: fields.map((field) => field.value),
      };
    })()`, "new line after committed command");
    await key("Backspace", "Backspace", 8);
    await waitForEvaluation(`(() => ({
      ready: document.querySelectorAll("math-field").length === 1,
      lineCount: document.querySelectorAll("math-field").length,
    }))()`, "remove empty line after committed command");

    await reloadEditor();
    await setField("x\\degree");
    const degreeBeforeCommit = await waitForEvaluation(
      `(() => {
        const field = document.querySelector("math-field");
        const selected = document.querySelector(".suggestion-item.is-selected");
        const surface = document.querySelector(".multi-line-editor");
        const state = {
          value: field?.value ?? "",
          position: field?.position ?? -1,
          lastOffset: field?.lastOffset ?? -1,
          selectedCommand: selected?.querySelector(".suggestion-command")?.textContent ?? "",
          commandQuery: surface?.dataset.commandQuery ?? "",
        };
        return {
          ...state,
          ready:
            Boolean(selected) &&
            state.selectedCommand.includes("\\\\circ") &&
            state.commandQuery === "\\\\degree",
        };
      })()`,
      "custom command candidate for \\degree",
    );
    await key("Enter", "Enter", 13);
    const degreeCommitState = await waitForEvaluation(`(() => {
      const field = document.querySelector("math-field");
      const selection = field.selection;
      const state = {
        value: field.value,
        position: field.position,
        lastOffset: field.lastOffset,
        selection,
        candidateVisible: Boolean(document.querySelector(".suggestion-popup")),
      };
      return {
        ...state,
        ready:
          state.value.includes("\\\\circ") &&
          state.position === state.lastOffset &&
          selection.ranges.every(([start, end]) => start === end) &&
          !state.candidateVisible,
      };
    })()`, `degree command caret after commit; before=${JSON.stringify(degreeBeforeCommit)}`);
    await key("Backspace", "Backspace", 8);
    const degreeDeleteState = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        value: field.value,
        position: field.position,
        lastOffset: field.lastOffset,
      };
    })()`);
    if (degreeDeleteState.value === degreeCommitState.value) {
      throw new Error(
        `Backspace did not modify the committed degree command: ${JSON.stringify({ degreeCommitState, degreeDeleteState })}`,
      );
    }

    await resetEditorDocument([""]);
    await setField("");
    await imeKey("、", "Backslash", 220, "、");
    const chineseIdeographicCommaValue = await evaluate(
      `document.querySelector("math-field")?.value ?? ""`,
    );
    if (chineseIdeographicCommaValue !== "、") {
      throw new Error(
        `Chinese Backslash key should insert only one ideographic comma: ${JSON.stringify(chineseIdeographicCommaValue)}`,
      );
    }

    await resetEditorDocument(["\\alpha", "\\beta"], 1);
    await evaluate(`(() => {
      const field = document.querySelectorAll("math-field")[1];
      if (!field) throw new Error("Second formula field was not found");
      field.position = field.lastOffset;
      field.focus();
      field.shadowRoot
        ?.querySelector('[part="keyboard-sink"]')
        ?.focus({ preventScroll: true });
    })()`);
    await key("ArrowUp", "ArrowUp", 38);
    const arrowUpLineState = await waitForEvaluation(`(() => {
      const rows = [...document.querySelectorAll(".formula-line")];
      const fields = [...document.querySelectorAll("math-field")];
      const surface = document.querySelector(".multi-line-editor");
      const firstLineId = rows[0]?.dataset.lineId ?? "";
      return {
        ready:
          rows.length === 2 &&
          rows[0]?.classList.contains("is-active") &&
          fields[0]?.matches(":focus-within") &&
          surface?.dataset.activeLineId === firstLineId,
        activeLineId: surface?.dataset.activeLineId ?? "",
        firstLineId,
        firstFocused: fields[0]?.matches(":focus-within") ?? false,
        candidateVisible: Boolean(document.querySelector(".suggestion-popup")),
      };
    })()`, "ArrowUp switches to previous formula line");

    if (arrowUpLineState.candidateVisible) {
      await key("Escape", "Escape", 27);
      await waitForEvaluation(
        `(() => ({ ready: !document.querySelector(".suggestion-popup") }))()`,
        "dismiss command candidate before formula-line navigation",
      );
    }

    await evaluate(`(() => {
      const field = document.querySelectorAll("math-field")[0];
      if (!field) throw new Error("First formula field was not found");
      field.position = field.lastOffset;
      field.focus();
      field.shadowRoot
        ?.querySelector('[part="keyboard-sink"]')
        ?.focus({ preventScroll: true });
    })()`);
    await sleep(80);
    await key("ArrowDown", "ArrowDown", 40);
    const arrowDownLineState = await waitForEvaluation(`(() => {
      const rows = [...document.querySelectorAll(".formula-line")];
      const fields = [...document.querySelectorAll("math-field")];
      const surface = document.querySelector(".multi-line-editor");
      const secondLineId = rows[1]?.dataset.lineId ?? "";
      return {
        ready:
          rows.length === 2 &&
          rows[1]?.classList.contains("is-active") &&
          fields[1]?.matches(":focus-within") &&
          surface?.dataset.activeLineId === secondLineId,
        activeLineId: surface?.dataset.activeLineId ?? "",
        secondLineId,
        secondFocused: fields[1]?.matches(":focus-within") ?? false,
        candidateVisible: Boolean(document.querySelector(".suggestion-popup")),
        nativeVisible:
          document
            .getElementById("mathlive-suggestion-popover")
            ?.classList.contains("is-visible") ?? false,
      };
    })()`, "ArrowDown switches to next formula line");

    const multiLineAlignment = await evaluate(`(() => {
      const fields = [...document.querySelectorAll(".formula-line math-field")];
      const leftEdges = fields.map((field) => field.getBoundingClientRect().left);
      const textAlignments = fields.map((field) => getComputedStyle(field).textAlign);
      return {
        count: fields.length,
        leftEdges,
        textAlignments,
        aligned:
          fields.length === 2 &&
          Math.abs(leftEdges[0] - leftEdges[1]) <= 1 &&
          textAlignments.every((value) => value !== "right"),
      };
    })()`);
    if (!multiLineAlignment.aligned) {
      throw new Error(
        `Multi-line formulas are not left-aligned: ${JSON.stringify(multiLineAlignment)}`,
      );
    }

    for (const operatorCase of [
      { latex: "\\sum_{i=0}^{n} a_i", command: "\\sum", label: "sum" },
      { latex: "\\prod_{k=1}^{m} b_k", command: "\\prod", label: "product" },
      { latex: "\\int_{0}^{1} f(x)\\,\\mathrm{d}x", command: "\\int", label: "integral" },
    ]) {
      await resetEditorDocument([operatorCase.latex]);
      const navigationStart = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const command = ${JSON.stringify(operatorCase.command)};
        let operatorOffset = -1;
        for (let offset = 0; offset <= field.lastOffset; offset += 1) {
          const latex = field.getElementInfo(offset)?.latex?.trim() ?? "";
          if (latex === command || latex.startsWith(command) || latex.includes(command)) {
            operatorOffset = offset;
            break;
          }
        }
        if (operatorOffset < 0) {
          throw new Error("Structured operator offset was not found: " + command);
        }
        field.position = operatorOffset;
        field.focus();
        field.shadowRoot
          ?.querySelector('[part="keyboard-sink"]')
          ?.focus({ preventScroll: true });
        const beforeLatex = field.value;
        window.__visualtexNavigationInputEvents = 0;
        field.addEventListener("input", () => {
          window.__visualtexNavigationInputEvents += 1;
        });
        const moved = field.executeCommand("moveToSubscript");
        return { moved, operatorOffset, beforeLatex, position: field.position };
      })()`);
      if (!navigationStart.moved) {
        throw new Error(
          `${operatorCase.label} lower-limit navigation could not start: ${JSON.stringify(navigationStart)}`,
        );
      }
      const lowerStart = await waitForEvaluation(
        `(() => {
          const field = document.querySelector("math-field");
          const markers = [...(field?.shadowRoot?.querySelectorAll(
            ".ML__placeholder-selected, .ML__selected, .ML__caret",
          ) ?? [])];
          const caret = markers.find((marker) => marker.closest(".ML__msubsup, .ML__op-group"));
          const script = caret?.closest(".ML__msubsup, .ML__op-group");
          if (!field || !caret || !script) return { ready: false, region: null, value: field?.value ?? "" };
          const caretBounds = (caret.parentElement ?? caret).getBoundingClientRect();
          const scriptBounds = script.getBoundingClientRect();
          const region = caretBounds.top + caretBounds.height / 2 < scriptBounds.top + scriptBounds.height / 2 ? "upper" : "lower";
          return { ready: region === "lower", region, value: field.value };
        })()`,
        `${operatorCase.label} caret starts in lower limit`,
      );
      await evaluate(`(() => {
        const field = document.querySelector("math-field");
        window.__visualtexNavigationBeforeLatex = field.value;
        window.__visualtexNavigationInputEvents = 0;
        return field.value;
      })()`);
      await key("ArrowUp", "ArrowUp", 38);
      const upperState = await waitForEvaluation(
        `(() => {
          const field = document.querySelector("math-field");
          const markers = [...(field?.shadowRoot?.querySelectorAll(
            ".ML__placeholder-selected, .ML__selected, .ML__caret",
          ) ?? [])];
          const caret = markers.find((marker) => marker.closest(".ML__msubsup, .ML__op-group"));
          const script = caret?.closest(".ML__msubsup, .ML__op-group");
          if (!field || !caret || !script) return { ready: false, region: null, value: field?.value ?? "" };
          const caretBounds = (caret.parentElement ?? caret).getBoundingClientRect();
          const scriptBounds = script.getBoundingClientRect();
          const region = caretBounds.top + caretBounds.height / 2 < scriptBounds.top + scriptBounds.height / 2 ? "upper" : "lower";
          return {
            ready:
              region === "upper" &&
              field.value === window.__visualtexNavigationBeforeLatex &&
              window.__visualtexNavigationInputEvents === 0,
            region,
            value: field.value,
            beforeLatex: window.__visualtexNavigationBeforeLatex,
            inputEvents: window.__visualtexNavigationInputEvents,
          };
        })()`,
        `${operatorCase.label} ArrowUp enters upper limit without changing LaTeX`,
      );
      await key("ArrowDown", "ArrowDown", 40);
      const lowerState = await waitForEvaluation(
        `(() => {
          const field = document.querySelector("math-field");
          const markers = [...(field?.shadowRoot?.querySelectorAll(
            ".ML__placeholder-selected, .ML__selected, .ML__caret",
          ) ?? [])];
          const caret = markers.find((marker) => marker.closest(".ML__msubsup, .ML__op-group"));
          const script = caret?.closest(".ML__msubsup, .ML__op-group");
          if (!field || !caret || !script) return { ready: false, region: null, value: field?.value ?? "" };
          const caretBounds = (caret.parentElement ?? caret).getBoundingClientRect();
          const scriptBounds = script.getBoundingClientRect();
          const region = caretBounds.top + caretBounds.height / 2 < scriptBounds.top + scriptBounds.height / 2 ? "upper" : "lower";
          return {
            ready:
              region === "lower" &&
              field.value === window.__visualtexNavigationBeforeLatex &&
              window.__visualtexNavigationInputEvents === 0,
            region,
            value: field.value,
            beforeLatex: window.__visualtexNavigationBeforeLatex,
            inputEvents: window.__visualtexNavigationInputEvents,
          };
        })()`,
        `${operatorCase.label} ArrowDown returns to lower limit without changing LaTeX`,
      );
      if (
        upperState.value !== lowerStart.value ||
        lowerState.value !== lowerStart.value ||
        upperState.inputEvents !== 0 ||
        lowerState.inputEvents !== 0
      ) {
        throw new Error(
          `${operatorCase.label} navigation unexpectedly changed LaTeX`,
        );
      }
    }

    await resetEditorDocument(["\\alpha"]);
    await setField("\\alpha");
    const simpleMetrics = await waitForEvaluation(`(() => {
      const line = document.querySelector(".formula-line");
      const field = document.querySelector("math-field");
      const content = field.shadowRoot.querySelector('[part="content"]');
      const rects = [...content.querySelectorAll("[data-atom-id]")]
        .map((atom) => atom.getBoundingClientRect())
        .filter((rect) => rect.height > 0);
      const top = rects.length ? Math.min(...rects.map((rect) => rect.top)) : 0;
      const bottom = rects.length ? Math.max(...rects.map((rect) => rect.bottom)) : 0;
      const lineRect = line.getBoundingClientRect();
      return {
        ready: field.classList.contains("is-simple-formula") && rects.length > 0,
        lineHeight: lineRect.height,
        fieldHeight: field.getBoundingClientRect().height,
        contentHeight: bottom - top,
        centerDelta: Math.abs((top + bottom) / 2 - (lineRect.top + lineRect.bottom) / 2),
      };
    })()`, "simple formula row sizing");

    await setField("\\frac{a}{b}");
    const tallMetrics = await waitForEvaluation(`(() => {
      const line = document.querySelector(".formula-line");
      const field = document.querySelector("math-field");
      const content = field.shadowRoot.querySelector('[part="content"]');
      const rects = [...content.querySelectorAll("[data-atom-id]")]
        .map((atom) => atom.getBoundingClientRect())
        .filter((rect) => rect.height > 0);
      const top = rects.length ? Math.min(...rects.map((rect) => rect.top)) : 0;
      const bottom = rects.length ? Math.max(...rects.map((rect) => rect.bottom)) : 0;
      const lineRect = line.getBoundingClientRect();
      return {
        ready:
          !field.classList.contains("is-simple-formula") &&
          rects.length > 0 &&
          top >= lineRect.top - 1 &&
          bottom <= lineRect.bottom + 1,
        lineHeight: lineRect.height,
        fieldHeight: field.getBoundingClientRect().height,
        contentHeight: bottom - top,
        topOverflow: Math.max(0, lineRect.top - top),
        bottomOverflow: Math.max(0, bottom - lineRect.bottom),
      };
    })()`, "unclipped tall formula row sizing");
    if (
      tallMetrics.contentHeight > simpleMetrics.contentHeight + 2 &&
      tallMetrics.lineHeight <= simpleMetrics.lineHeight + 2
    ) {
      throw new Error(`Formula row did not expand for taller rendered content: ${JSON.stringify({ simpleMetrics, tallMetrics })}`);
    }
    if (tallMetrics.topOverflow > 1 || tallMetrics.bottomOverflow > 1) {
      throw new Error(`Tall formula is clipped by its row: ${JSON.stringify(tallMetrics)}`);
    }
    if (simpleMetrics.centerDelta > 10) {
      throw new Error(`Simple formula is not vertically centered (delta ${simpleMetrics.centerDelta})`);
    }

    for (let index = 0; index < 8; index += 1) {
      await evaluate(`document.querySelector('button[aria-label="缩小公式"]').click()`);
      await sleep(70);
    }
    await setField("\\alpha");
    const compactZoomMetrics = await waitForEvaluation(`(() => {
      const field = document.querySelector("math-field");
      const line = document.querySelector(".formula-line");
      const content = field.shadowRoot.querySelector('[part="content"]');
      const rects = [...content.querySelectorAll("[data-atom-id]")]
        .map((atom) => atom.getBoundingClientRect())
        .filter((rect) => rect.height > 0);
      const top = rects.length ? Math.min(...rects.map((rect) => rect.top)) : 0;
      const bottom = rects.length ? Math.max(...rects.map((rect) => rect.bottom)) : 0;
      const surface = document.querySelector(".editor-surface");
      const stack = document.querySelector(".mathfield-stack");
      const surfaceRect = surface.getBoundingClientRect();
      const stackRect = stack.getBoundingClientRect();
      return {
        ready: field.classList.contains("is-simple-formula") && rects.length > 0,
        zoomLabel: document.querySelector(".canvas-controls > span")?.textContent?.trim(),
        zoomOutDisabled: document.querySelector('button[aria-label="缩小公式"]')?.disabled,
        fontSize: Number.parseFloat(getComputedStyle(field).fontSize),
        lineHeight: line.getBoundingClientRect().height,
        fieldHeight: field.getBoundingClientRect().height,
        contentHeight: bottom - top,
        surfaceWidth: surfaceRect.width,
        stackWidth: stackRect.width,
        leftGap: stackRect.left - surfaceRect.left,
        rightGap: surfaceRect.right - stackRect.right,
      };
    })()`, "compact simple formula sizing");
    if (
      compactZoomMetrics.zoomLabel !== "20%" ||
      !compactZoomMetrics.zoomOutDisabled ||
      compactZoomMetrics.fontSize > 11.2 ||
      compactZoomMetrics.lineHeight > 36
    ) {
      throw new Error(`20% zoom did not produce a compact formula row: ${JSON.stringify(compactZoomMetrics)}`);
    }
    if (
      compactZoomMetrics.leftGap > 70 ||
      compactZoomMetrics.rightGap > 70 ||
      compactZoomMetrics.stackWidth < compactZoomMetrics.surfaceWidth - 140
    ) {
      throw new Error(`Formula stack did not fill the wide editor surface: ${JSON.stringify(compactZoomMetrics)}`);
    }

    await setField("\\frac{a}{b}");
    const compactTallMetrics = await waitForEvaluation(`(() => {
      const line = document.querySelector(".formula-line");
      const field = document.querySelector("math-field");
      const content = field.shadowRoot.querySelector('[part="content"]');
      const rects = [...content.querySelectorAll("[data-atom-id]")]
        .map((atom) => atom.getBoundingClientRect())
        .filter((rect) => rect.height > 0);
      const top = rects.length ? Math.min(...rects.map((rect) => rect.top)) : 0;
      const bottom = rects.length ? Math.max(...rects.map((rect) => rect.bottom)) : 0;
      const lineRect = line.getBoundingClientRect();
      return {
        ready:
          !field.classList.contains("is-simple-formula") &&
          rects.length > 0 &&
          top >= lineRect.top - 1 &&
          bottom <= lineRect.bottom + 1,
        lineHeight: lineRect.height,
        fieldHeight: field.getBoundingClientRect().height,
        contentHeight: bottom - top,
        topOverflow: Math.max(0, lineRect.top - top),
        bottomOverflow: Math.max(0, bottom - lineRect.bottom),
      };
    })()`, "unclipped compact tall formula sizing");
    if (
      compactTallMetrics.contentHeight > compactZoomMetrics.contentHeight + 2 &&
      compactTallMetrics.lineHeight <= compactZoomMetrics.lineHeight + 2
    ) {
      throw new Error(`Tall formula did not expand at 20% zoom: ${JSON.stringify({ compactZoomMetrics, compactTallMetrics })}`);
    }
    if (compactTallMetrics.topOverflow > 1 || compactTallMetrics.bottomOverflow > 1) {
      throw new Error(`Compact tall formula is clipped: ${JSON.stringify(compactTallMetrics)}`);
    }

    for (let index = 0; index < 8; index += 1) {
      await evaluate(`document.querySelector('button[aria-label="放大公式"]').click()`);
      await sleep(70);
    }

    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.setValue("a\\\\int_{x}^{y}", {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
      field.position = field.lastOffset;
      field.focus();
      field.dispatchEvent(new InputEvent("input", { bubbles: true, composed: true }));
    })()`);
    await sleep(150);
    const deletionStates = [];
    for (let index = 0; index < 12; index += 1) {
      await key("Backspace", "Backspace", 8);
      deletionStates.push(
        await evaluate(`(() => {
          const field = document.querySelector("math-field");
          return { value: field.value, position: field.position };
        })()`),
      );
    }
    const skippedOperator = deletionStates.some(
      (state) => !state.value.includes("a") && /\\\\int/.test(state.value),
    );
    if (skippedOperator) {
      throw new Error(`Backspace skipped the integral and deleted preceding content: ${JSON.stringify(deletionStates)}`);
    }

    const ocrOpenMetrics = null;
    const ocrCenterMetrics = null;

    await setField("a=b");
    await key("Enter", "Enter", 13);
    await sleep(180);
    await setFieldAt(1, "c=d");
    await evaluate(`document.querySelector(".code-format-primary").click()`);
    await sleep(120);
    const formatMenuState = await evaluate(`(() => {
      const menu = document.querySelector(".code-format-menu");
      const rect = menu.getBoundingClientRect();
      const firstOption = menu.querySelector("button[data-format]");
      return {
        count: menu.querySelectorAll("button[data-format]").length,
        visibleTitleCount: menu.querySelectorAll(".code-format-item-title").length,
        descriptionCount: menu.querySelectorAll(".code-format-description").length,
        firstOptionText: firstOption?.innerText.trim() ?? "",
        firstOptionHeight: firstOption?.getBoundingClientRect().height ?? 0,
        top: rect.top,
        right: rect.right,
        bottom: rect.bottom,
        left: rect.left,
        viewportWidth: window.innerWidth,
        viewportHeight: window.innerHeight,
      };
    })()`);
    if (formatMenuState.count !== 18) {
      throw new Error(`Expected 18 LaTeX code formats, found ${formatMenuState.count}`);
    }
    if (
      formatMenuState.visibleTitleCount !== 0 ||
      formatMenuState.descriptionCount !== 0 ||
      formatMenuState.firstOptionText !== "\\frac{x}{y}" ||
      formatMenuState.firstOptionHeight > 50
    ) {
      throw new Error(`LaTeX code-format options are not compact code-only rows: ${JSON.stringify(formatMenuState)}`);
    }
    if (
      formatMenuState.left < 0 ||
      formatMenuState.top < 0 ||
      formatMenuState.right > formatMenuState.viewportWidth + 1 ||
      formatMenuState.bottom > formatMenuState.viewportHeight + 1
    ) {
      throw new Error(`LaTeX code-format menu is outside the viewport: ${JSON.stringify(formatMenuState)}`);
    }

    await evaluate(`document.querySelector('[data-format="align-star"]').click()`);
    await sleep(260);
    const alignFormatState = await evaluate(`(() => {
      const source = document.querySelector(".source-panel .cm-content")?.innerText ?? "";
      const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      return {
        source,
        sourceVisible: Boolean(document.querySelector(".source-panel")),
        persistedFormat: persisted.state?.latexCodeFormat,
      };
    })()`);
    if (!alignFormatState.sourceVisible) {
      throw new Error("Selecting a LaTeX code format did not open the source panel");
    }
    if (
      !alignFormatState.source.includes("\\begin{align*}") ||
      !alignFormatState.source.includes("a=b") ||
      !alignFormatState.source.includes("c=d") ||
      alignFormatState.source.includes("a&=b") ||
      alignFormatState.source.includes("c&=d") ||
      !alignFormatState.source.includes("\\end{align*}")
    ) {
      throw new Error(`align* source was not generated correctly: ${alignFormatState.source}`);
    }
    if (alignFormatState.persistedFormat !== "align-star") {
      throw new Error(`align* selection was not persisted: ${JSON.stringify(alignFormatState)}`);
    }

    await evaluate(`document.querySelector(".code-format-primary").click()`);
    await sleep(100);
    const alignSelectedAfterReopen = await evaluate(
      `document.querySelector('[data-format="align-star"]')?.getAttribute("aria-checked")`,
    );
    if (alignSelectedAfterReopen !== "true") {
      throw new Error("The current align* format was not marked as selected after reopening the menu");
    }
    await evaluate(`document.querySelector('[data-format="equation"]').click()`);
    await sleep(260);
    const equationFormatState = await evaluate(`(() => {
      const source = document.querySelector(".source-panel .cm-content")?.innerText ?? "";
      return {
        source,
        equationCount: (source.match(/\\\\begin\\{equation\\}/g) || []).length,
        endCount: (source.match(/\\\\end\\{equation\\}/g) || []).length,
        hasAlign: source.includes("\\\\begin{align"),
      };
    })()`);
    if (
      equationFormatState.equationCount !== 2 ||
      equationFormatState.endCount !== 2 ||
      equationFormatState.hasAlign
    ) {
      throw new Error(`equation source did not create independent environments: ${JSON.stringify(equationFormatState)}`);
    }
    // The source editor intentionally refreshes a changed format across two
    // animation frames. In a browser preview those frames may be throttled,
    // so wait until the formatted source remains stable before typing a draft.
    await sleep(500);

    const editedEquationSource = [
      "\\begin{equation}",
      "a=q",
      "\\end{equation}",
      "",
      "\\begin{equation}",
      "c=d",
      "\\end{equation}",
    ].join("\n");
    await evaluate(`document.querySelector(".source-panel .cm-content").focus()`);
    await replaceFocusedText(editedEquationSource);
    const dirtyBeforeFormatSwitch = await evaluate(`(() => ({
      dirty: Boolean(document.querySelector(".source-panel .unsaved-chip")),
      source: document.querySelector(".source-panel .cm-content")?.innerText ?? "",
    }))()`);
    if (!dirtyBeforeFormatSwitch.source.includes("a=q")) {
      throw new Error(`CodeMirror draft edit was not registered: ${JSON.stringify(dirtyBeforeFormatSwitch)}`);
    }

    await evaluate(`document.querySelector(".code-format-primary").click()`);
    await sleep(100);
    await evaluate(`document.querySelector('[data-format="align-star"]').click()`);
    await sleep(420);
    const dirtyFormatSwitchState = await evaluate(`(() => ({
      source: document.querySelector(".source-panel .cm-content")?.innerText ?? "",
      dirty: Boolean(document.querySelector(".source-panel .unsaved-chip")),
      formulas: [...document.querySelectorAll("math-field")].map((field) => field.value),
    }))()`);
    if (
      dirtyFormatSwitchState.dirty ||
      !dirtyFormatSwitchState.source.includes("\\begin{align*}") ||
      !dirtyFormatSwitchState.source.includes("a=q") ||
      dirtyFormatSwitchState.source.includes("a&=q") ||
      dirtyFormatSwitchState.formulas[0] !== "a=q" ||
      dirtyFormatSwitchState.formulas[1] !== "c=d"
    ) {
      throw new Error(`Unsynced source edits were lost during format switching: ${JSON.stringify(dirtyFormatSwitchState)}`);
    }

    await evaluate(`(() => {
      localStorage.removeItem("visualtex.onboarding.v3.completed");
      location.reload();
    })()`);
    await sleep(850);
    await evaluate(`new Promise((resolve) => {
      const done = () => document.querySelector(".onboarding-dialog") ? resolve(true) : setTimeout(done, 30);
      done();
    })`);
    for (let index = 0; index < 3; index += 1) {
      await evaluate(`document.querySelector(".onboarding-actions .primary-button").click()`);
      await sleep(100);
    }
    const onboardingHotkeyStep = await evaluate(`(() => ({
      progressCount: document.querySelectorAll(".onboarding-progress > span").length,
      visible: Boolean(document.querySelector(".onboarding-hotkeys-tiles-demo")),
      title: document.querySelector("#onboarding-title")?.textContent ?? "",
      hasCustomTile: Boolean(document.querySelector(".onboarding-custom-tile-guide")),
    }))()`);
    if (
      onboardingHotkeyStep.progressCount !== 9 ||
      !onboardingHotkeyStep.visible ||
      !onboardingHotkeyStep.hasCustomTile ||
      !onboardingHotkeyStep.title.includes("快捷")
    ) {
      throw new Error(`The onboarding hotkey and tile step is incomplete: ${JSON.stringify(onboardingHotkeyStep)}`);
    }

    await evaluate(`document.querySelector(".onboarding-actions .primary-button").click()`);
    await sleep(100);
    const onboardingLayoutStep = await evaluate(`(() => ({
      visible: Boolean(document.querySelector(".onboarding-layout-theme-demo")),
      selectedText: document.querySelector(".onboarding-layout-choice-list article.is-selected strong")?.textContent ?? "",
      themeCount: document.querySelectorAll(".onboarding-theme-swatches > span").length,
    }))()`);
    if (
      !onboardingLayoutStep.visible ||
      !onboardingLayoutStep.selectedText.includes("经典") ||
      onboardingLayoutStep.themeCount !== 5
    ) {
      throw new Error(`The onboarding layout and theme step is incomplete: ${JSON.stringify(onboardingLayoutStep)}`);
    }

    for (let index = 0; index < 3; index += 1) {
      await evaluate(`document.querySelector(".onboarding-actions .primary-button").click()`);
      await sleep(100);
    }
    const onboardingFormatStep = await evaluate(`(() => ({
      progressCount: document.querySelectorAll(".onboarding-progress > span").length,
      visible: Boolean(document.querySelector(".onboarding-code-format-demo")),
      title: document.querySelector("#onboarding-title")?.textContent ?? "",
      source: document.querySelector(".onboarding-code-format-demo pre")?.textContent ?? "",
    }))()`);
    if (
      onboardingFormatStep.progressCount !== 9 ||
      !onboardingFormatStep.visible ||
      !onboardingFormatStep.title.includes("LaTeX") ||
      !onboardingFormatStep.source.includes("\\begin{align*}")
    ) {
      throw new Error(`The onboarding LaTeX format step is incomplete: ${JSON.stringify(onboardingFormatStep)}`);
    }

    console.log(JSON.stringify({
      thetaState,
      enterAfterCandidateState,
      nativePopover,
      nativeCommitState,
      simpleMetrics,
      tallMetrics,
      compactZoomMetrics,
      compactTallMetrics,
      ocrOpenMetrics,
      ocrCenterMetrics,
      deletionStates,
      formatMenuState,
      alignFormatState,
      alignSelectedAfterReopen,
      equationFormatState,
      dirtyBeforeFormatSwitch,
      dirtyFormatSwitchState,
      degreeBeforeCommit,
      degreeCommitState,
      degreeDeleteState,
      chineseIdeographicCommaValue,
      arrowUpLineState,
      arrowDownLineState,
      onboardingFormatStep,
    }, null, 2));
    console.log("Editor regression smoke test passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(300);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => undefined);
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.stack : error);
  process.exitCode = 1;
});
