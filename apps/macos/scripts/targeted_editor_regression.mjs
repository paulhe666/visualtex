import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm, writeFile } from "node:fs/promises";
import process from "node:process";

const scenario = process.argv[2];
if (!new Set(["wrapper", "wrapper-bm", "wrapper-auto", "wrapper-continuous", "wrapper-prefix", "native-input-popover", "native-structure-audit", "native-structure-input-over", "native-structure-input-under", "native-structure-input-multi", "native-structure-input-core", "native-space-selection", "native-space-ime-replay", "candidate-query-reset", "limit-candidate", "raw-placeholder-visual", "placeholder-selection", "pointer-release-stability", "structural-placeholder", "structured-chinese-ime", "accent-placeholder", "caret-probe", "vertical-structure-probe", "vertical-structure-navigation", "scripts", "upright", "formula-formatting", "font-variant-formatting", "context-style", "suggestions", "navigation", "geometry", "source-layout", "source-preview-only", "toolbar-compact", "formula-tiles", "cursor-placement", "font-settings", "output-fonts", "configuration", "ocr-model-selection", "ocr-empty-environment", "settings", "layout", "delete", "export"]).has(scenario)) {
  throw new Error(
    "Usage: node scripts/targeted_editor_regression.mjs <wrapper|wrapper-bm|wrapper-auto|wrapper-continuous|wrapper-prefix|native-input-popover|native-structure-audit|native-structure-input-over|native-structure-input-under|native-structure-input-multi|native-structure-input-core|native-space-selection|native-space-ime-replay|candidate-query-reset|limit-candidate|raw-placeholder-visual|placeholder-selection|pointer-release-stability|structural-placeholder|structured-chinese-ime|accent-placeholder|caret-probe|vertical-structure-probe|vertical-structure-navigation|scripts|upright|formula-formatting|context-style|suggestions|navigation|geometry|source-layout|source-preview-only|toolbar-compact|formula-tiles|cursor-placement|font-settings|output-fonts|configuration|ocr-model-selection|ocr-empty-environment|settings|layout|delete|export>",
  );
}

const offset = process.pid % 1000;
const previewPort = 16400 + offset;
const debugPort = 21600 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = `/tmp/visualtex-targeted-${scenario}-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const canonicalLatex = (value) =>
  value.replace(/\s+/g, "").replace(/\{([A-Za-z0-9])\}/g, "$1");

async function waitFor(url, timeoutMs = 15000) {
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
    this.events = [];
  }

  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((resolve, reject) => {
      this.socket.addEventListener("open", resolve, { once: true });
      this.socket.addEventListener("error", reject, { once: true });
    });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
      if (!message.id) {
        if (
          message.method === "Runtime.exceptionThrown" ||
          message.method === "Runtime.consoleAPICalled"
        ) {
          this.events.push(message);
        }
        return;
      }
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
    let lastTargets = [];
    const targetStartedAt = Date.now();
    while (!page && Date.now() - targetStartedAt < 5000) {
      lastTargets = await (
        await fetch(`http://127.0.0.1:${debugPort}/json/list`)
      ).json();
      page =
        lastTargets.find(
          (target) => target.type === "page" && target.url.startsWith(baseUrl),
        ) ?? lastTargets.find((target) => target.type === "page");
      if (!page) await sleep(80);
    }
    if (!page) {
      throw new Error(
        `No Chrome page target found: ${JSON.stringify(lastTargets)}`,
      );
    }

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
        throw new Error(
          result.exceptionDetails.exception?.description ||
            result.exceptionDetails.text ||
            "Runtime.evaluate failed",
        );
      }
      return result.result.value;
    };

    const waitForEvaluation = async (expression, description, timeoutMs = 12000) => {
      const started = Date.now();
      let lastValue;
      while (Date.now() - started < timeoutMs) {
        lastValue = await evaluate(expression);
        if (lastValue?.ready) return lastValue;
        await sleep(50);
      }
      const runtimeEvents = client.events.map((event) => ({
        method: event.method,
        exception:
          event.params?.exceptionDetails?.exception?.description ??
          event.params?.exceptionDetails?.text ??
          null,
        console:
          event.params?.args?.map((arg) => arg.value ?? arg.description ?? "") ??
          null,
      }));
      throw new Error(
        `Timed out waiting for ${description}: ${JSON.stringify({ lastValue, runtimeEvents })}`,
      );
    };

    const key = async (value, code, virtualKeyCode, modifiers = 0) => {
      const common = {
        key: value,
        code,
        modifiers,
        windowsVirtualKeyCode: virtualKeyCode,
        nativeVirtualKeyCode: virtualKeyCode,
      };
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        ...common,
        ...(value.length === 1 && modifiers === 0
          ? { text: value, unmodifiedText: value }
          : {}),
      });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(45);
    };

    const typeText = async (text) => {
      for (const character of text) {
        const code = character === "\\" ? "Backslash" : `Key${character.toUpperCase()}`;
        const virtualKeyCode = character === "\\" ? 220 : character.toUpperCase().charCodeAt(0);
        await key(character, code, virtualKeyCode);
      }
    };

    if (scenario === "ocr-model-selection") {
      await client.send("Page.addScriptToEvaluateOnNewDocument", {
        source: `(() => {
          const callbacks = new Map();
          let callbackId = 1;
          window.__visualtexOcrModelSelectionProbe = { prewarmed: [] };
          window.__TAURI_INTERNALS__ = {
            metadata: {
              currentWindow: { label: "main" },
              currentWebview: { label: "main" },
            },
            transformCallback(callback, once = false) {
              const id = callbackId++;
              callbacks.set(id, { callback, once });
              return id;
            },
            unregisterCallback(id) {
              callbacks.delete(id);
            },
            async invoke(command, args) {
              if (command === "get_ocr_runtime_status") {
                return {
                  installed: true,
                  pythonPath: "/fake/python",
                  pythonVersion: "3.10.0",
                  paddleVersion: "3.3.1",
                  paddleocrVersion: "3.7.0",
                  runtimePath: "/fake/runtime",
                  offlineBundleAvailable: true,
                  installedModels: ["PP-FormulaNet_plus-M"],
                  defaultModel: "PP-FormulaNet_plus-M",
                  message: "Fake M-only OCR runtime",
                };
              }
              if (command === "prewarm_ocr_model") {
                window.__visualtexOcrModelSelectionProbe.prewarmed.push(args?.model ?? null);
                return null;
              }
              if (command === "configure_silent_ocr") return null;
              if (command === "plugin:event|listen") return 1;
              if (command === "plugin:event|unlisten") return null;
              return null;
            },
          };
        })();`,
      });
    }

    await client.send("Page.navigate", { url: baseUrl });
    await sleep(650);
    await evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
      localStorage.setItem(
        "visualtex.onboarding.macos.desktop.v1.2.0.completed",
        "true",
      );
      localStorage.setItem(
        "visualtex.office.macos.native-first-run.v1.2.0.completed",
        "true",
      );
      const storageKey = "visualtex-editor";
      const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
      persisted.state = {
        ...(persisted.state || {}),
        lines: [{ id: crypto.randomUUID(), latex: "" }],
        activeLineId: null,
        checkUpdatesOnStartup: false,
      };
      persisted.state.activeLineId = persisted.state.lines[0].id;
      delete persisted.state.inputBehavior;
      localStorage.setItem(storageKey, JSON.stringify(persisted));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await waitForEvaluation(
      `(() => ({ ready: Boolean(document.querySelector("math-field")) }))()`,
      "formula field",
    );

    const focusField = async () => {
      await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        if (!field?.isConnected) return { ready: false };
        field.focus();
        field.position = field.lastOffset;
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        return { ready: field.isConnected && field.hasFocus() };
      })()`, "stable focused formula field");
      await sleep(80);
    };

    const clearField = async () => {
      await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        if (!field?.isConnected) return { ready: false };
        field.setValue("", {
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
        return { ready: field.isConnected && field.value === "" };
      })()`, "stable empty formula field");
      await sleep(100);
      await focusField();
    };

    if (scenario === "ocr-empty-environment") {
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector("math-field")),
      }))()`, "OCR empty-environment formula field");
      await evaluate(`(() => {
        const field = document.querySelector("math-field");
        if (!field) return false;
        field.setValue(String.raw\`\\begin{align*}a&=b\\\\c&=d\\end{align*}\`, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.dispatchEvent(new InputEvent("input", {
          bubbles: true,
          composed: true,
          inputType: "insertFromPaste",
        }));
        field.focus();
        field.selection = {
          ranges: [[0, field.lastOffset]],
          direction: "forward",
        };
        return true;
      })()`);
      await sleep(120);
      const beforeDelete = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        return {
          fieldValue: field?.value ?? null,
          lastOffset: field?.lastOffset ?? null,
          selection: field ? structuredClone(field.selection) : null,
          storeLatex: persisted.state?.lines?.[0]?.latex ?? null,
        };
      })()`);
      await key("Backspace", "Backspace", 8);
      await sleep(180);
      const afterDelete = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        return {
          fieldValue: field?.value ?? null,
          visibleValue: field?.getValue("latex-without-placeholders") ?? null,
          lastOffset: field?.lastOffset ?? null,
          selection: field ? structuredClone(field.selection) : null,
          storeLatex: persisted.state?.lines?.[0]?.latex ?? null,
        };
      })()`);
      assert.match(
        beforeDelete.fieldValue ?? "",
        /\\begin\{align\*\}/,
        JSON.stringify(beforeDelete),
      );
      assert.equal(
        afterDelete.fieldValue,
        "",
        `Deleting all OCR multiline content must remount a truly empty MathLive field: ${JSON.stringify(afterDelete)}`,
      );
      assert.equal(
        afterDelete.storeLatex,
        "",
        `Deleting all OCR multiline content must clear the stored formula source: ${JSON.stringify(afterDelete)}`,
      );
      console.log("OCR multiline empty-environment regression passed");
      return;
    }

    if (scenario === "ocr-model-selection") {
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector(".canvas-ocr-model select")),
      }))()`, "OCR model selector");

      await sleep(1500);
      const selectModel = async (model) => {
        await evaluate(`(() => {
          const select = document.querySelector(".canvas-ocr-model select");
          if (!select) return;
          const setter = Object.getOwnPropertyDescriptor(
            HTMLSelectElement.prototype,
            "value",
          )?.set;
          setter?.call(select, ${JSON.stringify(model)});
          select.dispatchEvent(new Event("change", { bubbles: true }));
        })()`);
        await sleep(650);
        return evaluate(`(() => {
          const select = document.querySelector(".canvas-ocr-model select");
          return {
            selected: select?.value ?? "",
            stored: localStorage.getItem("visualtex.ocr.model"),
            prewarmed: window.__visualtexOcrModelSelectionProbe?.prewarmed ?? [],
          };
        })()`);
      };

      const sState = await selectModel("PP-FormulaNet_plus-S");
      assert.equal(sState.selected, "PP-FormulaNet_plus-S", JSON.stringify(sState));
      assert.equal(sState.stored, "PP-FormulaNet_plus-S", JSON.stringify(sState));
      assert.ok(
        sState.prewarmed.includes("PP-FormulaNet_plus-M"),
        JSON.stringify(sState),
      );

      await evaluate(`document.querySelector('button[aria-label="图片公式识别"]')?.click()`);
      const sDialogState = await waitForEvaluation(`(() => {
        const dialog = document.querySelector(".ocr-dialog");
        const select = dialog?.querySelector(".ocr-model-field select");
        const warning = dialog?.querySelector(".ocr-model-warning strong")?.textContent ?? "";
        const importButton = [...(dialog?.querySelectorAll(".ocr-install-actions button") ?? [])]
          .find((button) => button.textContent?.includes("导入模型包"));
        return {
          ready:
            Boolean(dialog) &&
            select?.value === "PP-FormulaNet_plus-S" &&
            warning.includes("尚未安装") &&
            Boolean(importButton),
          selected: select?.value ?? "",
          warning,
          importButton: importButton?.textContent?.trim() ?? "",
        };
      })()`, "S model remains selected in OCR model manager");

      await evaluate(`(() => {
        const select = document.querySelector(".ocr-dialog .ocr-model-field select");
        if (!select) return;
        const setter = Object.getOwnPropertyDescriptor(
          HTMLSelectElement.prototype,
          "value",
        )?.set;
        setter?.call(select, "PP-FormulaNet_plus-L");
        select.dispatchEvent(new Event("change", { bubbles: true }));
      })()`);
      await sleep(650);
      const lState = await evaluate(`(() => ({
        selected: document.querySelector(".ocr-dialog .ocr-model-field select")?.value ?? "",
        toolbarSelected: document.querySelector(".canvas-ocr-model select")?.value ?? "",
        stored: localStorage.getItem("visualtex.ocr.model"),
        warning: document.querySelector(".ocr-dialog .ocr-model-warning strong")?.textContent ?? "",
        prewarmed: window.__visualtexOcrModelSelectionProbe?.prewarmed ?? [],
      }))()`);
      assert.equal(lState.selected, "PP-FormulaNet_plus-L", JSON.stringify(lState));
      assert.equal(lState.toolbarSelected, "PP-FormulaNet_plus-L", JSON.stringify(lState));
      assert.equal(lState.stored, "PP-FormulaNet_plus-L", JSON.stringify(lState));
      assert.match(lState.warning, /尚未安装/);

      console.log(JSON.stringify({ sState, sDialogState, lState }, null, 2));
      console.log("Targeted OCR model selection persistence regression passed");
      return;
    }

    if (scenario === "source-preview-only") {
      await evaluate(`(() => {
        const storageKey = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
        persisted.state = {
          ...(persisted.state || {}),
          editorLayout: "standard",
          sourceOpen: false,
          inputBehavior: {
            ...(persisted.state?.inputBehavior || {}),
            showOtherCommandSuggestions: true,
          },
        };
        localStorage.setItem(storageKey, JSON.stringify(persisted));
        location.reload();
      })()`);
      await waitForEvaluation(`(() => ({
        ready:
          Boolean(document.querySelector("math-field")) &&
          Boolean(document.querySelector(".source-toggle")),
      }))()`, "standard editor with collapsed source");
      await evaluate(`(() => {
        const standardToggle = document.querySelector(".source-toggle");
        const classicToggle = document.querySelector('[data-classic-bottom-view="source"]');
        (standardToggle || classicToggle)?.click();
      })()`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector(".source-panel .cm-content")),
      }))()`, "source editor is visible");

      const sourceFocusPoint = await evaluate(`(() => {
        const lines = [...document.querySelectorAll(".source-panel .cm-line")];
        const target = lines[1] ?? lines[0];
        const rect = target?.getBoundingClientRect();
        return rect ? { x: rect.left + 8, y: rect.top + rect.height / 2 } : null;
      })()`);
      assert.ok(sourceFocusPoint, "source editor has a click target");
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: sourceFocusPoint.x,
        y: sourceFocusPoint.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: sourceFocusPoint.x,
        y: sourceFocusPoint.y,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });
      const focusedState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const surface = document.querySelector(".multi-line-editor");
        const workspace = document.querySelector(".workspace");
        const line = document.querySelector(".formula-line");
        const native = document.getElementById("mathlive-suggestion-popover");
        const stable = document.getElementById("visualtex-native-input-suggestion-popover");
        const isPainted = (node) => {
          if (!node) return false;
          const style = getComputedStyle(node);
          const background = style.backgroundColor;
          return style.display !== "none" &&
            style.visibility !== "hidden" &&
            Number.parseFloat(style.opacity || "1") > 0 &&
            background !== "transparent" &&
            background !== "rgba(0, 0, 0, 0)";
        };
        const selectionNodes = [...(field?.shadowRoot?.querySelectorAll(
          ".ML__selection, .ML__selected, .ML__contains-highlight, .ML__highlight, .ML__placeholder-selected",
        ) ?? [])];
        const caretNodes = [...(field?.shadowRoot?.querySelectorAll(
          ".ML__caret, .ML__text-caret, .ML__latex-caret, .visualtex-structural-placeholder-caret",
        ) ?? [])];
        const lineStyle = line ? getComputedStyle(line) : null;
        return {
          ready:
            document.documentElement.classList.contains("visualtex-source-editor-focused") &&
            workspace?.classList.contains("is-source-editor-focused") &&
            surface?.classList.contains("is-source-preview-only") &&
            field?.classList.contains("visualtex-source-preview-only") &&
            field?.readOnly === true &&
            field?.selectionIsCollapsed === true &&
            !document.querySelector(".suggestion-popup") &&
            !isPainted(native) &&
            !isPainted(stable) &&
            selectionNodes.every((node) => !isPainted(node)) &&
            caretNodes.every((node) => {
              const style = getComputedStyle(node);
              return style.display === "none" || Number.parseFloat(style.opacity || "1") === 0;
            }) &&
            (!lineStyle || lineStyle.backgroundColor === "rgba(0, 0, 0, 0)"),
          rootFocused: document.documentElement.classList.contains("visualtex-source-editor-focused"),
          workspaceFocused: workspace?.classList.contains("is-source-editor-focused") ?? false,
          surfacePreviewOnly: surface?.classList.contains("is-source-preview-only") ?? false,
          fieldPreviewOnly: field?.classList.contains("visualtex-source-preview-only") ?? false,
          readOnly: field?.readOnly ?? false,
          selectionCollapsed: field?.selectionIsCollapsed ?? false,
          customCandidateVisible: Boolean(document.querySelector(".suggestion-popup")),
          nativePainted: isPainted(native),
          stablePainted: isPainted(stable),
          paintedSelectionCount: selectionNodes.filter(isPainted).length,
          visibleCaretCount: caretNodes.filter((node) => {
            const style = getComputedStyle(node);
            return style.display !== "none" && Number.parseFloat(style.opacity || "1") > 0;
          }).length,
          lineBackground: lineStyle?.backgroundColor ?? "",
        };
      })()`, "source focus turns the visual editor into clean live preview");

      const sourceInsertionPoint = await evaluate(`(() => {
        const lines = [...document.querySelectorAll(".source-panel .cm-line")];
        const target =
          lines.find((line, index) => index > 0 && !line.textContent?.trim()) ??
          lines[1] ??
          lines[0];
        const rect = target?.getBoundingClientRect();
        return rect
          ? { x: rect.left + 8, y: rect.top + rect.height / 2 }
          : null;
      })()`);
      assert.ok(sourceInsertionPoint, "source formula line has a click target");
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: sourceInsertionPoint.x,
        y: sourceInsertionPoint.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: sourceInsertionPoint.x,
        y: sourceInsertionPoint.y,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });
      await client.send("Input.insertText", {
        text: "\\theta+\\placeholder{}",
      });
      const liveRenderState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const source = document.querySelector(".source-panel .cm-content")?.innerText ?? "";
        return {
          ready:
            source.includes("\\\\theta+\\\\placeholder{}") &&
            field?.value.includes("\\\\theta") &&
            field?.value.includes("\\\\placeholder{}") &&
            field.readOnly &&
            field.selectionIsCollapsed &&
            document.documentElement.classList.contains("visualtex-source-editor-focused") &&
            !document.querySelector(".suggestion-popup") &&
            !document.querySelector("[data-source-reset]"),
          source,
          value: field?.value ?? "",
          readOnly: field?.readOnly ?? false,
          selectionCollapsed: field?.selectionIsCollapsed ?? false,
          popupVisible: Boolean(document.querySelector(".suggestion-popup")),
          resetVisible: Boolean(document.querySelector("[data-source-reset]")),
          sourceErrorVisible: Boolean(document.querySelector(".source-error-chip")),
          rootFocused: document.documentElement.classList.contains("visualtex-source-editor-focused"),
          cmFocused: document.querySelector(".source-panel .cm-editor")?.classList.contains("cm-focused") ?? false,
          activeTag: document.activeElement?.tagName ?? "",
          activeClass: document.activeElement?.className ?? "",
        };
      })()`, "previewable placeholder source renders live without a reset-only state");

      const visualClickPoint = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const content = field?.shadowRoot?.querySelector('[part="content"]');
        const rect = content?.getBoundingClientRect() ?? field?.getBoundingClientRect();
        return rect
          ? { x: rect.left + Math.max(8, rect.width * 0.65), y: rect.top + rect.height / 2 }
          : null;
      })()`);
      assert.ok(visualClickPoint, "rendered formula has a click target");
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: visualClickPoint.x,
        y: visualClickPoint.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: visualClickPoint.x,
        y: visualClickPoint.y,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });
      const restoredInteractionState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        return {
          ready:
            !document.documentElement.classList.contains("visualtex-source-editor-focused") &&
            !document.querySelector(".workspace")?.classList.contains("is-source-editor-focused") &&
            !document.querySelector(".multi-line-editor")?.classList.contains("is-source-preview-only") &&
            !field?.classList.contains("visualtex-source-preview-only") &&
            field?.readOnly === false &&
            field?.hasFocus() &&
            field?.value.includes("\\\\placeholder{}") &&
            !document.querySelector("[data-source-reset]") &&
            !document.querySelector(".source-error-chip"),
          value: field?.value ?? "",
          readOnly: field?.readOnly ?? true,
          fieldFocused: field?.hasFocus() ?? false,
          fieldPreviewOnly: field?.classList.contains("visualtex-source-preview-only") ?? true,
          resetVisible: Boolean(document.querySelector("[data-source-reset]")),
          sourceErrorVisible: Boolean(document.querySelector(".source-error-chip")),
        };
      })()`, "one visual click accepts the previewable source and enters formula editing");

      await clearField();
      await typeText("\\th");
      const restoredCandidateState = await waitForEvaluation(`(() => {
        const native = document.getElementById("mathlive-suggestion-popover");
        const stable = document.getElementById("visualtex-native-input-suggestion-popover");
        const visible = (node) => {
          if (!node) return false;
          const style = getComputedStyle(node);
          return style.display !== "none" &&
            style.visibility !== "hidden" &&
            Number.parseFloat(style.opacity || "1") > 0;
        };
        const customVisible = Boolean(document.querySelector(".suggestion-popup"));
        const nativeVisible = visible(native);
        const stableVisible = visible(stable);
        return {
          ready: customVisible || nativeVisible || stableVisible,
          customVisible,
          nativeVisible,
          stableVisible,
        };
      })()`, "visual command candidates return after source focus leaves");

      console.log(JSON.stringify({
        focusedState,
        liveRenderState,
        restoredInteractionState,
        restoredCandidateState,
      }, null, 2));
      console.log("Targeted source-only live preview regression passed");
      return;
    }

    if (scenario === "toolbar-compact") {
      await evaluate(`(() => {
        if (!document.querySelector(".formula-toolbar")) {
          document.querySelector(".sidebar-toggle")?.click();
        }
        return true;
      })()`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(
          document.querySelector(".formula-toolbar") &&
            document.querySelector(".template-strip") &&
            document.querySelectorAll(".toolbar-tab").length === 9,
        ),
      }))()`, "formula toolbar");

      const tabLayoutState = await waitForEvaluation(`(() => {
        const expectedOrder = [
          "common",
          "structure",
          "calculus",
          "matrix",
          "relation",
          "greek",
          "arrow",
          "physics",
          "set",
        ];
        const container = document.querySelector(".toolbar-tabs");
        const tabs = [...document.querySelectorAll(".toolbar-tab")];
        const bounds = tabs.map((tab) => tab.getBoundingClientRect());
        const rows = [];
        for (const boundsItem of bounds) {
          const row = rows.find(
            (entry) => Math.abs(entry.top - boundsItem.top) <= 1,
          );
          if (row) row.count += 1;
          else rows.push({ top: boundsItem.top, count: 1 });
        }
        const style = container ? getComputedStyle(container) : null;
        const gridColumnCount =
          style?.gridTemplateColumns.split(" ").filter(Boolean).length ?? 0;
        const actualOrder = tabs.map((tab) => tab.dataset.category ?? "");
        const labelsFit = tabs.every(
          (tab) => tab.scrollWidth <= tab.clientWidth + 1,
        );
        const containerHeight =
          container?.getBoundingClientRect().height ?? Number.POSITIVE_INFINITY;
        return {
          ready:
            JSON.stringify(actualOrder) === JSON.stringify(expectedOrder) &&
            rows.length === 1 &&
            rows[0]?.count === 9 &&
            labelsFit &&
            containerHeight <= 40,
          actualOrder,
          gridColumnCount,
          rows,
          labelsFit,
          containerHeight,
        };
      })()`, "single-row toolbar tabs");

      const categories = [
        "common",
        "structure",
        "calculus",
        "matrix",
        "relation",
        "greek",
        "arrow",
        "physics",
        "set",
      ];
      const categoryStates = [];

      for (const category of categories) {
        await evaluate(`document.querySelector(
          '.toolbar-tab[data-category="${category}"]',
        ).click()`);
        await sleep(100);
        const state = await waitForEvaluation(`(() => {
          const strip = document.querySelector(
            ".template-strip.is-continuous-categories",
          );
          const section = document.querySelector(
            '.toolbar-category-section[data-toolbar-category-section="${category}"]',
          );
          const buttons = [...(section?.querySelectorAll(
            ":scope > .template-button",
          ) ?? [])];
          const bounds = buttons.map((button) =>
            button.getBoundingClientRect(),
          );
          const firstRowTop = bounds[0]?.top ?? -1;
          const firstRow = bounds.filter(
            (rect) => Math.abs(rect.top - firstRowTop) <= 1,
          );
          const directContentOnly = buttons.every(
            (button) =>
              button.children.length === 1 &&
              button.firstElementChild?.classList.contains("math-preview"),
          );
          const previewStates = buttons.map((button) => {
            const preview = button.querySelector(".math-preview");
            const latex = preview?.querySelector(".ML__latex");
            const style = preview ? getComputedStyle(preview) : null;
            const previewBounds = preview?.getBoundingClientRect();
            const latexBounds = latex?.getBoundingClientRect();
            return {
              visible: Boolean(
                preview &&
                  latex &&
                  style?.display !== "none" &&
                  style?.visibility !== "hidden" &&
                  Number.parseFloat(style?.opacity || "1") > 0.1 &&
                  previewBounds &&
                  previewBounds.width > 4 &&
                  previewBounds.height > 4 &&
                  latexBounds &&
                  latexBounds.width > 0 &&
                  latexBounds.height > 0,
              ),
              display: style?.display ?? "missing",
              width: previewBounds?.width ?? 0,
              height: previewBounds?.height ?? 0,
              latexWidth: latexBounds?.width ?? 0,
              latexHeight: latexBounds?.height ?? 0,
            };
          });
          const previewsVisible = previewStates.every((state) => state.visible);
          const ariaLabelsPresent = buttons.every(
            (button) => Boolean(button.getAttribute("aria-label")?.trim()),
          );
          const gradientButtons = buttons.every((button) =>
            getComputedStyle(button).backgroundImage.includes("linear-gradient"),
          );
          const minimumHeight = bounds.length
            ? Math.min(...bounds.map((rect) => rect.height))
            : 0;
          const maximumHeight = bounds.length
            ? Math.max(...bounds.map((rect) => rect.height))
            : 0;
          const equalHeights =
            minimumHeight >= 44 && maximumHeight - minimumHeight <= 1;
          const baseWidth = bounds.length
            ? Math.min(...bounds.map((rect) => rect.width))
            : 0;
          const gridAlignedWidths = bounds.every((rect) => {
            if (baseWidth <= 0) return false;
            const units = rect.width / baseWidth;
            return Math.abs(units - Math.round(units)) <= 0.03;
          });
          const rowGroups = new Map();
          bounds.forEach((rect) => {
            const key = Math.round(rect.top);
            const row = rowGroups.get(key) ?? [];
            row.push(rect);
            rowGroups.set(key, row);
          });
          const seamlessRows = [...rowGroups.values()].every((row) => {
            const sorted = [...row].sort((left, right) => left.left - right.left);
            return sorted.every(
              (rect, index) =>
                index === 0 || Math.abs(rect.left - sorted[index - 1].right) <= 1,
            );
          });
          const stripStyle = strip ? getComputedStyle(strip) : null;
          const sectionStyle = section ? getComputedStyle(section) : null;
          const gridColumnCount = sectionStyle?.gridTemplateColumns
            .split(" ")
            .filter(Boolean).length ?? 0;
          return {
            ready:
              Boolean(strip) &&
              Boolean(section) &&
              buttons.length > 0 &&
              directContentOnly &&
              previewsVisible &&
              ariaLabelsPresent &&
              gradientButtons &&
              equalHeights &&
              gridAlignedWidths &&
              seamlessRows &&
              gridColumnCount > 0,
            category: ${JSON.stringify(category)},
            buttonCount: buttons.length,
            firstRowCount: firstRow.length,
            buttonWidth: bounds[0]?.width ?? 0,
            buttonHeight: bounds[0]?.height ?? 0,
            gridTemplateColumns: sectionStyle?.gridTemplateColumns ?? "",
            gridColumnCount,
            directContentOnly,
            previewsVisible,
            firstPreviewState: previewStates[0] ?? null,
            ariaLabelsPresent,
            gradientButtons,
            firstButtonBackground:
              buttons[0] ? getComputedStyle(buttons[0]).backgroundImage : '',
            equalHeights,
            gridAlignedWidths,
            baseWidth,
            seamlessRows,
            rowCount: rowGroups.size,
            columnGap: sectionStyle?.columnGap ?? '',
            rowGap: sectionStyle?.rowGap ?? '',
            sectionGap: stripStyle?.columnGap ?? '',
            horizontalOverflow:
              strip ? strip.scrollWidth - strip.clientWidth : -1,
          };
        })()`, `dense seamless toolbar category: ${category}`);
        categoryStates.push(state);
      }

      await evaluate(`document.querySelector(
        '.toolbar-tab[data-category="calculus"]',
      ).click()`);
      await sleep(100);
      const partialColumnLayout = await evaluate(`(() => {
        const toolbar = document.querySelector('.formula-toolbar');
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const section = strip?.querySelector(
          '[data-toolbar-category-section="calculus"]',
        );
        const nextSection = strip?.querySelector(
          '[data-toolbar-category-section="matrix"]',
        );
        const buttons = [...(section?.querySelectorAll(
          ':scope > .template-button',
        ) ?? [])];
        const rowCount = Number(toolbar?.dataset.toolbarRowCount ?? 0);
        const remainder = rowCount > 0 ? buttons.length % rowCount : -1;
        const columnCounts = new Map();
        for (const button of buttons) {
          const left = Math.round(button.getBoundingClientRect().left);
          columnCounts.set(left, (columnCounts.get(left) ?? 0) + 1);
        }
        const sortedColumns = [...columnCounts.entries()].sort(
          (left, right) => left[0] - right[0],
        );
        const sectionRect = section?.getBoundingClientRect();
        const nextRect = nextSection?.getBoundingClientRect();
        const stripStyle = strip ? getComputedStyle(strip) : null;
        return {
          category: strip?.dataset.activeCategory ?? '',
          rowCount,
          buttonCount: buttons.length,
          remainder,
          lastColumnCount: sortedColumns.at(-1)?.[1] ?? 0,
          transitionCount:
            strip?.querySelectorAll('.toolbar-category-transition').length ?? -1,
          sectionGap:
            sectionRect && nextRect ? nextRect.left - sectionRect.right : -1,
          configuredGap: stripStyle ? parseFloat(stripStyle.columnGap) : -1,
        };
      })()`);
      assert.equal(
        partialColumnLayout.category,
        'calculus',
        JSON.stringify(partialColumnLayout),
      );
      assert.ok(
        partialColumnLayout.remainder > 0,
        JSON.stringify(partialColumnLayout),
      );
      assert.equal(
        partialColumnLayout.lastColumnCount,
        partialColumnLayout.remainder,
        JSON.stringify(partialColumnLayout),
      );
      assert.equal(
        partialColumnLayout.transitionCount,
        0,
        JSON.stringify(partialColumnLayout),
      );
      assert.ok(
        partialColumnLayout.sectionGap >= 12 &&
          partialColumnLayout.sectionGap <= 16 &&
          Math.abs(
            partialColumnLayout.sectionGap -
              partialColumnLayout.configuredGap,
          ) <= 1,
        JSON.stringify(partialColumnLayout),
      );

      await evaluate(`(() => {
        const toolbar = document.querySelector('.formula-toolbar');
        if (toolbar) {
          toolbar.style.width = '420px';
          toolbar.style.maxWidth = '420px';
          toolbar.style.justifySelf = 'start';
        }
        document.querySelector(
          '.toolbar-tab[data-category="common"]',
        )?.click();
      })()`);
      await sleep(180);
      const internalWheelSetup = await evaluate(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const common = strip?.querySelector(
          '[data-toolbar-category-section="common"]',
        );
        const structure = strip?.querySelector(
          '[data-toolbar-category-section="structure"]',
        );
        if (!strip || !common || !structure) return null;
        const commonStart = common.offsetLeft;
        const commonEnd = Math.max(
          commonStart,
          common.offsetLeft + common.offsetWidth - strip.clientWidth,
        );
        strip.scrollLeft = commonStart;
        strip.dispatchEvent(new WheelEvent('wheel', {
          deltaY: 80,
          bubbles: true,
          cancelable: true,
        }));
        return {
          commonStart,
          commonEnd,
          structureStart: structure.offsetLeft,
          clientWidth: strip.clientWidth,
          commonWidth: common.offsetWidth,
        };
      })()`);
      assert.ok(
        internalWheelSetup &&
          internalWheelSetup.commonWidth > internalWheelSetup.clientWidth,
        JSON.stringify(internalWheelSetup),
      );
      const internalWheelState = await waitForEvaluation(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const active = document.querySelector('.toolbar-tab.is-active');
        const common = strip?.querySelector(
          '[data-toolbar-category-section="common"]',
        );
        const commonStart = common?.offsetLeft ?? -1;
        const commonEnd = common
          ? Math.max(
              commonStart,
              common.offsetLeft + common.offsetWidth - strip.clientWidth,
            )
          : -1;
        return {
          ready:
            strip?.dataset.activeCategory === 'common' &&
            active?.dataset.category === 'common' &&
            (strip?.scrollLeft ?? -1) > commonStart + 1 &&
            (strip?.scrollLeft ?? -1) <= commonEnd + 1,
          category: strip?.dataset.activeCategory ?? '',
          activeTab: active?.dataset.category ?? '',
          scrollLeft: strip?.scrollLeft ?? -1,
          commonStart,
          commonEnd,
        };
      })()`, 'wheel first scrolls inside an overflowing toolbar category');
      await sleep(380);
      await evaluate(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const common = strip?.querySelector(
          '[data-toolbar-category-section="common"]',
        );
        if (!strip || !common) return;
        strip.scrollLeft = Math.max(
          common.offsetLeft,
          common.offsetLeft + common.offsetWidth - strip.clientWidth,
        );
        strip.dispatchEvent(new WheelEvent('wheel', {
          deltaY: 80,
          bubbles: true,
          cancelable: true,
        }));
      })()`);
      const forwardWheelState = await waitForEvaluation(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const active = document.querySelector('.toolbar-tab.is-active');
        const structure = strip?.querySelector(
          '[data-toolbar-category-section="structure"]',
        );
        return {
          ready:
            strip?.dataset.activeCategory === 'structure' &&
            active?.dataset.category === 'structure' &&
            structure &&
            Math.abs((strip?.scrollLeft ?? -1) - structure.offsetLeft) <= 1,
          category: strip?.dataset.activeCategory ?? '',
          activeTab: active?.dataset.category ?? '',
          scrollLeft: strip?.scrollLeft ?? -1,
          structureStart: structure?.offsetLeft ?? -1,
        };
      })()`, 'extra wheel aligns the next category to the left edge');
      await sleep(380);
      await evaluate(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        if (!strip) return;
        strip.dispatchEvent(new WheelEvent('wheel', {
          deltaY: -80,
          bubbles: true,
          cancelable: true,
        }));
      })()`);
      const reverseWheelState = await waitForEvaluation(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const active = document.querySelector('.toolbar-tab.is-active');
        const common = strip?.querySelector(
          '[data-toolbar-category-section="common"]',
        );
        const commonEnd = common
          ? Math.max(
              common.offsetLeft,
              common.offsetLeft + common.offsetWidth - strip.clientWidth,
            )
          : -1;
        return {
          ready:
            strip?.dataset.activeCategory === 'common' &&
            active?.dataset.category === 'common' &&
            Math.abs((strip?.scrollLeft ?? -1) - commonEnd) <= 1,
          category: strip?.dataset.activeCategory ?? '',
          activeTab: active?.dataset.category ?? '',
          scrollLeft: strip?.scrollLeft ?? -1,
          commonEnd,
        };
      })()`, 'reverse wheel returns to the previous category end');
      await sleep(380);
      const underfilledWheelSetup = await evaluate(`(() => {
        const toolbar = document.querySelector('.formula-toolbar');
        if (toolbar) {
          toolbar.style.width = '1200px';
          toolbar.style.maxWidth = '1200px';
        }
        document.querySelector(
          '.toolbar-tab[data-category="greek"]',
        )?.click();
        return true;
      })()`);
      assert.equal(underfilledWheelSetup, true);
      await sleep(180);
      const underfilledBefore = await evaluate(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const greek = strip?.querySelector(
          '[data-toolbar-category-section="greek"]',
        );
        const arrow = strip?.querySelector(
          '[data-toolbar-category-section="arrow"]',
        );
        if (!strip || !greek || !arrow) return null;
        strip.scrollLeft = greek.offsetLeft;
        const result = {
          greekWidth: greek.offsetWidth,
          clientWidth: strip.clientWidth,
          greekStart: greek.offsetLeft,
          arrowStart: arrow.offsetLeft,
        };
        strip.dispatchEvent(new WheelEvent('wheel', {
          deltaY: 80,
          bubbles: true,
          cancelable: true,
        }));
        return result;
      })()`);
      assert.ok(
        underfilledBefore &&
          underfilledBefore.greekWidth <= underfilledBefore.clientWidth,
        JSON.stringify(underfilledBefore),
      );
      const underfilledWheelState = await waitForEvaluation(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const active = document.querySelector('.toolbar-tab.is-active');
        const arrow = strip?.querySelector(
          '[data-toolbar-category-section="arrow"]',
        );
        return {
          ready:
            strip?.dataset.activeCategory === 'arrow' &&
            active?.dataset.category === 'arrow' &&
            arrow &&
            Math.abs((strip?.scrollLeft ?? -1) - arrow.offsetLeft) <= 1,
          category: strip?.dataset.activeCategory ?? '',
          activeTab: active?.dataset.category ?? '',
          scrollLeft: strip?.scrollLeft ?? -1,
          arrowStart: arrow?.offsetLeft ?? -1,
        };
      })()`, 'one wheel aligns the next category when the current category fits');
      await evaluate(`(() => {
        const toolbar = document.querySelector('.formula-toolbar');
        if (toolbar) {
          toolbar.style.removeProperty('width');
          toolbar.style.removeProperty('max-width');
          toolbar.style.removeProperty('justify-self');
        }
        document.querySelector(
          '.toolbar-tab[data-category="matrix"]',
        )?.click();
      })()`);
      const matrixDelimiterState = await waitForEvaluation(`(() => {
        const buttons = [...document.querySelectorAll(
          ".matrix-delimiter-options button",
        )];
        const bounds = buttons.map((button) => button.getBoundingClientRect());
        const previewStates = buttons.map((button) => {
          const preview = button.querySelector(".math-preview");
          const latex = preview?.querySelector(".ML__latex");
          const style = preview ? getComputedStyle(preview) : null;
          const previewBounds = preview?.getBoundingClientRect();
          const latexBounds = latex?.getBoundingClientRect();
          return {
            visible: Boolean(
              preview &&
                latex &&
                style?.display !== "none" &&
                style?.visibility !== "hidden" &&
                Number.parseFloat(style?.opacity || "1") > 0.1 &&
                previewBounds &&
                previewBounds.width > 4 &&
                previewBounds.height > 4 &&
                latexBounds &&
                latexBounds.width > 0 &&
                latexBounds.height > 0,
            ),
            display: style?.display ?? "missing",
            width: previewBounds?.width ?? 0,
            height: previewBounds?.height ?? 0,
            latexWidth: latexBounds?.width ?? 0,
            latexHeight: latexBounds?.height ?? 0,
          };
        });
        const builder = document.querySelector(".matrix-builder");
        const grid = document.querySelector(".matrix-size-grid");
        const builderBounds = builder?.getBoundingClientRect();
        const gridBounds = grid?.getBoundingClientRect();
        const helperTextRemoved =
          !document.querySelector(".matrix-size-picker-label") &&
          !builder?.textContent?.includes("最大 10 × 10") &&
          !builder?.textContent?.includes("Up to 10 × 10") &&
          !builder?.textContent?.includes("单击选择行数和列数") &&
          !builder?.textContent?.includes("Click to select rows and columns");
        return {
          ready:
            buttons.length === 6 &&
            buttons.every(
              (button) =>
                button.children.length === 1 &&
                button.firstElementChild?.classList.contains("math-preview") &&
                Boolean(button.getAttribute("aria-label")?.trim()),
            ) &&
            previewStates.every((state) => state.visible) &&
            bounds.every((rect) => rect.height <= 44) &&
            helperTextRemoved &&
            Boolean(builderBounds && builderBounds.height <= 355) &&
            Boolean(gridBounds && gridBounds.width <= 214),
          count: buttons.length,
          heights: bounds.map((rect) => rect.height),
          contentCounts: buttons.map((button) => button.children.length),
          previewStates,
          ariaLabels: buttons.map(
            (button) => button.getAttribute("aria-label") ?? "",
          ),
          helperTextRemoved,
          builderHeight: builderBounds?.height ?? 0,
          gridWidth: gridBounds?.width ?? 0,
        };
      })()`, "compact matrix delimiter symbols");

      const adaptiveCategoryStates = [];
      for (const category of ["common", "structure", "calculus", "matrix"]) {
        await evaluate(`document.querySelector(
          '.toolbar-tab[data-category="${category}"]',
        ).click()`);
        const state = await waitForEvaluation(`(() => {
          const section = document.querySelector(
            '.toolbar-category-section[data-toolbar-category-section="${category}"]',
          );
          const buttons = [...(section?.querySelectorAll(
            ":scope > .template-button",
          ) ?? [])];
          const previews = buttons.map((button) => {
            const host = button.querySelector(".math-preview");
            const content = host?.querySelector(".math-preview-fit-content");
            const hostBounds = host?.getBoundingClientRect();
            const contentBounds = content?.getBoundingClientRect();
            const widthRatio =
              hostBounds && contentBounds && hostBounds.width > 0
                ? contentBounds.width / hostBounds.width
                : 0;
            const heightRatio =
              hostBounds && contentBounds && hostBounds.height > 0
                ? contentBounds.height / hostBounds.height
                : 0;
            const inside = Boolean(
              hostBounds &&
                contentBounds &&
                contentBounds.left >= hostBounds.left - 0.75 &&
                contentBounds.right <= hostBounds.right + 0.75 &&
                contentBounds.top >= hostBounds.top - 0.75 &&
                contentBounds.bottom <= hostBounds.bottom + 0.75,
            );
            return {
              commandId: button.dataset.commandId ?? "",
              previewLatex: button.dataset.previewLatex ?? "",
              autoFit: button.classList.contains("is-unified-fit"),
              fitReady: host?.dataset.fitReady === "true",
              scale: Number.parseFloat(host?.dataset.fitScale ?? "0"),
              inside,
              widthRatio,
              heightRatio,
              fillRatio: Math.max(widthRatio, heightRatio),
              host: hostBounds
                ? { width: hostBounds.width, height: hostBounds.height }
                : null,
              content: contentBounds
                ? { width: contentBounds.width, height: contentBounds.height }
                : null,
            };
          });
          const invalid = previews.filter(
            (preview) =>
              !preview.autoFit ||
              !preview.fitReady ||
              !preview.inside ||
              preview.fillRatio < 0.7 ||
              !Number.isFinite(preview.scale) ||
              preview.scale <= 0,
          );
          return {
            ready: buttons.length > 0 && invalid.length === 0,
            category: ${JSON.stringify(category)},
            buttonCount: buttons.length,
            minimumFillRatio: Math.min(
              ...previews.map((preview) => preview.fillRatio),
            ),
            maximumScale: Math.max(
              ...previews.map((preview) => preview.scale),
            ),
            invalid,
            previews,
          };
        })()`, `per-command fitted previews: ${category}`);
        adaptiveCategoryStates.push(state);
      }

      await evaluate(`document.querySelector(
        '.toolbar-tab[data-category="physics"]',
      ).click()`);
      const physicsExceptionState = await waitForEvaluation(`(() => {
        const inspect = (commandId) => {
          const button = document.querySelector(
            '.template-button[data-command-id="' + commandId + '"]',
          );
          const host = button?.querySelector(".math-preview");
          const content = host?.querySelector(".math-preview-fit-content");
          const hostBounds = host?.getBoundingClientRect();
          const contentBounds = content?.getBoundingClientRect();
          const widthRatio =
            hostBounds && contentBounds && hostBounds.width > 0
              ? contentBounds.width / hostBounds.width
              : 0;
          const heightRatio =
            hostBounds && contentBounds && hostBounds.height > 0
              ? contentBounds.height / hostBounds.height
              : 0;
          return {
            commandId,
            autoFit: button?.classList.contains("is-unified-fit") ?? false,
            fitReady: host?.dataset.fitReady === "true",
            scale: Number.parseFloat(host?.dataset.fitScale ?? "0"),
            inside: Boolean(
              hostBounds &&
                contentBounds &&
                contentBounds.left >= hostBounds.left - 0.75 &&
                contentBounds.right <= hostBounds.right + 0.75 &&
                contentBounds.top >= hostBounds.top - 0.75 &&
                contentBounds.bottom <= hostBounds.bottom + 0.75,
            ),
            fillRatio: Math.max(widthRatio, heightRatio),
            host: hostBounds
              ? { width: hostBounds.width, height: hostBounds.height }
              : null,
            content: contentBounds
              ? { width: contentBounds.width, height: contentBounds.height }
              : null,
          };
        };
        const commutator = inspect("commutator");
        const anticommutator = inspect("anticommutator");
        const physicsSection = document.querySelector(
          '.toolbar-category-section[data-toolbar-category-section="physics"]',
        );
        const ordinaryButtons = [...(physicsSection?.querySelectorAll(
          ":scope > .template-button:not([data-command-id='commutator']):not([data-command-id='anticommutator'])",
        ) ?? [])];
        const ordinaryUnchanged = ordinaryButtons.every(
          (button) =>
            button.classList.contains("is-unified-fit") &&
            button.querySelector(".math-preview")?.dataset.fitReady === "true",
        );
        return {
          ready:
            [commutator, anticommutator].every(
              (preview) =>
                preview.autoFit &&
                preview.fitReady &&
                preview.inside &&
                preview.fillRatio >= 0.7 &&
                preview.scale > 0,
            ) && ordinaryUnchanged,
          commutator,
          anticommutator,
          ordinaryUnchanged,
        };
      })()`, "fitted physics bracket previews");

      await evaluate(`document.querySelector(
        '.toolbar-tab[data-category="calculus"]',
      ).click()`);
      const calculusPreviewState = await waitForEvaluation(`(() => {
        const preview = (commandId) =>
          document.querySelector(
            '.template-button[data-command-id="' + commandId + '"]',
          )?.dataset.previewLatex ?? "";
        const values = {
          intplain: preview("intplain"),
          definiteIntegral: preview("int"),
          summation: preview("sum"),
          product: preview("prod"),
          limit: preview("lim"),
        };
        return {
          ready:
            values.intplain === "\\\\int" &&
            values.definiteIntegral === "\\\\int_a^b" &&
            values.summation === "\\\\sum_{i=1}^{n}" &&
            values.product === "\\\\prod_{i=1}^{n}" &&
            values.limit === "\\\\lim_{x\\\\to0}" &&
            !Object.values(values).some(
              (value) =>
                value.includes("f(") ||
                value.includes("mathrm") ||
                value.includes("a_i"),
            ),
          values,
        };
      })()`, "simplified calculus toolbar previews");

      await evaluate(`document.querySelector(
        '.toolbar-tab[data-category="common"]',
      ).click()`);
      const commonContentsState = await waitForEvaluation(`(() => {
        const expectedIds = [
          "frac",
          "sqrt",
          "power",
          "subscript",
          "hat",
          "tilde",
          "parentheses",
          "absolute",
          "intplain",
          "int",
          "iint",
          "oint",
          "sum",
          "prod",
          "lim",
          "partial",
          "derivative",
          "nabla",
          "infty",
          "matrix2",
          "cases",
          "vector",
          "alpha",
          "beta",
          "gamma",
          "theta",
          "lambda",
          "mu",
          "pi",
          "sigma",
          "omega",
          "delta",
          "equal",
          "neq",
          "approx",
          "leq",
          "geq",
          "propto",
          "in",
          "subset",
          "rightarrow",
          "notin",
          "forall",
          "exists",
          "leftarrow",
        ];
        const section = document.querySelector(
          '.toolbar-category-section[data-toolbar-category-section="common"]',
        );
        const buttons = [...(section?.querySelectorAll(
          ":scope > .template-button",
        ) ?? [])];
        const actualIds = buttons.map((button) => button.dataset.commandId ?? "");
        const missingIds = expectedIds.filter((id) => !actualIds.includes(id));
        return {
          ready:
            actualIds.length === expectedIds.length &&
            JSON.stringify(actualIds) === JSON.stringify(expectedIds) &&
            missingIds.length === 0,
          count: actualIds.length,
          actualIds,
          missingIds,
        };
      })()`, "expanded common formula collection");

      console.log(
        JSON.stringify({
          categories: categoryStates.map((state) => ({
            category: state.category,
            count: state.buttonCount,
            rows: state.rowCount,
            gradientButtons: state.gradientButtons,
            seamlessRows: state.seamlessRows,
          })),
          matrixDelimiters: matrixDelimiterState.count,
          categoryWheel: {
            internal: internalWheelState,
            forward: forwardWheelState,
            reverse: reverseWheelState,
            underfilled: underfilledWheelState,
            partialColumn: partialColumnLayout,
          },
          fittedCategories: adaptiveCategoryStates.map((state) => ({
            category: state.category,
            count: state.buttonCount,
            minimumFillRatio: state.minimumFillRatio,
          })),
          physicsReady: physicsExceptionState.ready,
          calculusReady: calculusPreviewState.ready,
          commonCount: commonContentsState.count,
          tabRows: tabLayoutState.rows,
        }),
      );
      console.log("Targeted compact formula toolbar regression passed");
      return;
    }

    if (scenario === "formula-tiles") {
      await evaluate(`(() => {
        localStorage.removeItem("visualtex-custom-formula-tiles");
        if (!document.querySelector(".formula-toolbar")) {
          document.querySelector(".sidebar-toggle")?.click();
        }
        return true;
      })()`);
      await waitForEvaluation(`(() => ({
        ready:
          Boolean(document.querySelector('[data-toolbar-view="tools"]')) &&
          Boolean(document.querySelector('[data-toolbar-view="tiles"]')) &&
          !document.querySelector(".formula-toolbar-actions") &&
          !document.querySelector(".add-formula-line"),
      }))()`, "formula tool and tile view tabs without legacy buttons");

      await evaluate(`document.querySelector('[data-toolbar-view="tiles"]').click()`);
      const commonTilesState = await waitForEvaluation(`(() => {
        const panel = document.querySelector(".formula-tiles-panel");
        const buttons = [...document.querySelectorAll(
          '.formula-tile-list > .formula-tile-button',
        )];
        const previews = buttons.map((button) => {
          const host = button.querySelector(".formula-tile-preview");
          const content = host?.querySelector(".math-preview-fit-content");
          const hostBounds = host?.getBoundingClientRect();
          const contentBounds = content?.getBoundingClientRect();
          return {
            id: button.dataset.formulaTileId ?? "",
            fitReady: host?.dataset.fitReady === "true",
            scale: Number.parseFloat(host?.dataset.fitScale ?? "0"),
            buttonHeight: button.getBoundingClientRect().height,
            hostHeight: hostBounds?.height ?? 0,
            inside: Boolean(
              hostBounds &&
                contentBounds &&
                contentBounds.left >= hostBounds.left - 1 &&
                contentBounds.right <= hostBounds.right + 1 &&
                contentBounds.top >= hostBounds.top - 1 &&
                contentBounds.bottom <= hostBounds.bottom + 1,
            ),
          };
        });
        const heights = previews.map((preview) => Math.round(preview.hostHeight));
        return {
          ready:
            Boolean(panel) &&
            buttons.length === 10 &&
            previews.every(
              (preview) =>
                preview.fitReady &&
                preview.scale > 0 &&
                preview.inside &&
                preview.hostHeight >= 52 &&
                preview.buttonHeight >= preview.hostHeight,
            ) &&
            panel.scrollWidth <= panel.clientWidth + 1,
          count: buttons.length,
          heights,
          distinctHeightCount: new Set(heights).size,
          horizontalOverflow: panel
            ? panel.scrollWidth - panel.clientWidth
            : -1,
          previews,
        };
      })()`, "ten fitted common formula tiles");

      await clearField();
      await evaluate(`document.querySelector(
        '[data-formula-tile-id="quadratic-formula"]',
      ).click()`);
      const insertedTileState = await waitForEvaluation(`(() => {
        const field = document.querySelector(
          ".formula-line.is-active math-field",
        );
        return {
          ready:
            Boolean(field) &&
            field.value.includes("\\\\frac{-b\\\\pm\\\\sqrt{b^2-4ac}}{2a}"),
          value: field?.value ?? "",
        };
      })()`, "common formula tile insertion");

      await evaluate(`(() => {
        const field = document.querySelector(
          ".formula-line.is-active math-field",
        );
        if (!field) return false;
        field.setValue("E=mc^2", { silenceNotifications: false });
        field.dispatchEvent(new InputEvent("input", {
          bubbles: true,
          composed: true,
          inputType: "insertText",
          data: null,
        }));
        field.focus();
        return true;
      })()`);
      await waitForEvaluation(`(() => {
        const field = document.querySelector(
          ".formula-line.is-active math-field",
        );
        return { ready: field?.value === "E=mc^2", value: field?.value ?? "" };
      })()`, "selected line replaced for custom tile");
      await evaluate(`document.querySelector(
        '[data-tile-category="custom"]',
      ).click()`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(
          document.querySelector(".save-current-formula-tile:not(:disabled)"),
        ),
      }))()`, "enabled custom tile save button");
      await evaluate(`document.querySelector(
        ".save-current-formula-tile",
      ).click()`);
      const customTileState = await waitForEvaluation(`(() => {
        const button = document.querySelector(
          '[data-formula-tile-id="custom-0"]',
        );
        const stored = JSON.parse(
          localStorage.getItem("visualtex-custom-formula-tiles") || "[]",
        );
        const host = button?.querySelector(".formula-tile-preview");
        return {
          ready:
            Boolean(button) &&
            Array.isArray(stored) &&
            stored.length === 1 &&
            stored[0] === "E=mc^2" &&
            host?.dataset.fitReady === "true",
          stored,
          latex: button?.dataset.formulaTileLatex ?? "",
          fitReady: host?.dataset.fitReady ?? "missing",
        };
      })()`, "persisted selected-line custom formula tile");

      await evaluate(`(() => {
        const button = document.querySelector(
          '[data-formula-tile-id="custom-0"]',
        );
        const field = document.querySelector(
          ".formula-line.is-active math-field",
        );
        if (!button || !field) return false;
        window.__visualtexFieldValueBeforeTileContextMenu = field.value;
        const bounds = button.getBoundingClientRect();
        button.dispatchEvent(new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          button: 2,
          buttons: 2,
          clientX: bounds.left + 24,
          clientY: bounds.top + 20,
        }));
        return true;
      })()`);
      const customTileContextState = await waitForEvaluation(`(() => {
        const menu = document.querySelector(".formula-tile-context-menu");
        const deleteButton = menu?.querySelector('[role="menuitem"]');
        const field = document.querySelector(
          ".formula-line.is-active math-field",
        );
        const stored = JSON.parse(
          localStorage.getItem("visualtex-custom-formula-tiles") || "[]",
        );
        return {
          ready:
            Boolean(menu && deleteButton) &&
            field?.value ===
              window.__visualtexFieldValueBeforeTileContextMenu &&
            stored.length === 1,
          fieldValueBefore:
            window.__visualtexFieldValueBeforeTileContextMenu ?? "",
          fieldValue: field?.value ?? "",
          stored,
          menuText: deleteButton?.textContent?.trim() ?? "",
        };
      })()`, "custom tile right-click menu without insertion conflict");

      await evaluate(`document.querySelector(
        ".formula-tile-context-menu [role=menuitem]",
      ).click()`);
      const deletedCustomTileState = await waitForEvaluation(`(() => {
        const stored = JSON.parse(
          localStorage.getItem("visualtex-custom-formula-tiles") || "[]",
        );
        return {
          ready:
            !document.querySelector('[data-formula-tile-id^="custom-"]') &&
            !document.querySelector(".formula-tile-context-menu") &&
            Boolean(document.querySelector(".formula-tile-empty")) &&
            Array.isArray(stored) &&
            stored.length === 0,
          stored,
          emptyVisible: Boolean(document.querySelector(".formula-tile-empty")),
        };
      })()`, "right-click custom tile deletion persistence");

      console.log(
        JSON.stringify(
          {
            commonTilesState,
            insertedTileState,
            customTileState,
            customTileContextState,
            deletedCustomTileState,
          },
          null,
          2,
        ),
      );
      console.log("Targeted formula tiles regression passed");
      return;
    }

    if (scenario === "source-layout") {
      await evaluate(`(() => {
        const storageKey = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
        const lines = Array.from({ length: 24 }, (_, index) => ({
          id: "source-layout-" + index,
          latex:
            "\\\\frac{a_{" +
            index +
            "}}{b_{" +
            index +
            "}}+\\\\sum_{n=1}^{100}x_n^{" +
            (index + 1) +
            "}",
        }));
        persisted.state = {
          ...(persisted.state || {}),
          lines,
          activeLineId: lines[0].id,
          sourceOpen: false,
          latexCodeFormat: "raw",
          theme: "dark",
        };
        localStorage.setItem(storageKey, JSON.stringify(persisted));
        location.reload();
        return true;
      })()`);
      await waitForEvaluation(`(() => ({
        ready:
          document.querySelectorAll("math-field").length === 24 &&
          Boolean(document.querySelector(".source-toggle")) &&
          !document.querySelector(".source-panel"),
      }))()`, "collapsed source layout");

      const collapsed = await evaluate(`(() => {
        const label = document.querySelector(".source-toggle-label");
        const toggle = document.querySelector(".source-toggle");
        const row = document.querySelector(".source-toggle-row");
        return {
          labelVisible: Boolean(label && label.getBoundingClientRect().width > 0),
          toggleVisible: Boolean(toggle && toggle.getBoundingClientRect().width > 0),
          rowHeight: row?.getBoundingClientRect().height ?? 0,
          pageScrollHeight: document.scrollingElement?.scrollHeight ?? 0,
          pageClientHeight: document.scrollingElement?.clientHeight ?? 0,
        };
      })()`);
      if (
        !collapsed.labelVisible ||
        !collapsed.toggleVisible ||
        collapsed.rowHeight < 30 ||
        collapsed.pageScrollHeight > collapsed.pageClientHeight + 1
      ) {
        throw new Error(`Collapsed source controls are invalid: ${JSON.stringify(collapsed)}`);
      }

      await evaluate(`document.querySelector(".source-toggle").click()`);
      await waitForEvaluation(`(() => ({
        ready:
          Boolean(document.querySelector(".source-panel .cm-scroller")) &&
          Boolean(document.querySelector(".source-collapse-button")) &&
          !document.querySelector(".source-toggle-row"),
      }))()`, "expanded source layout");
      await sleep(180);

      const expanded = await evaluate(`(() => {
        const rect = (element) => {
          const bounds = element?.getBoundingClientRect();
          return bounds
            ? { top: bounds.top, right: bounds.right, bottom: bounds.bottom, left: bounds.left, width: bounds.width, height: bounds.height }
            : null;
        };
        const body = document.querySelector(".editor-pane-body");
        const editorScroll = document.querySelector(".editor-pane-scroll");
        const editorSurface = document.querySelector(".editor-surface");
        const sourceSlot = document.querySelector(".source-pane-slot");
        const sourcePanel = document.querySelector(".source-panel");
        const sourceScroller = document.querySelector(".source-panel .cm-scroller");
        editorScroll.scrollTop = 0;
        sourceScroller.scrollTop = 0;
        editorScroll.scrollTop = Math.min(120, editorScroll.scrollHeight - editorScroll.clientHeight);
        const editorAfterOwnScroll = editorScroll.scrollTop;
        const sourceAfterEditorScroll = sourceScroller.scrollTop;
        sourceScroller.scrollTop = Math.min(120, sourceScroller.scrollHeight - sourceScroller.clientHeight);
        return {
          body: rect(body),
          editorScroll: rect(editorScroll),
          editorSurface: rect(editorSurface),
          sourceSlot: rect(sourceSlot),
          sourcePanel: rect(sourcePanel),
          bodyOverflow: getComputedStyle(body).overflow,
          editorOverflowY: getComputedStyle(editorScroll).overflowY,
          sourceOverflowY: getComputedStyle(sourceScroller).overflowY,
          editorScrollable: editorScroll.scrollHeight > editorScroll.clientHeight + 1,
          sourceScrollable: sourceScroller.scrollHeight > sourceScroller.clientHeight + 1,
          editorAfterOwnScroll,
          editorAfterSourceScroll: editorScroll.scrollTop,
          sourceAfterEditorScroll,
          sourceAfterOwnScroll: sourceScroller.scrollTop,
          syntaxColors: [...document.querySelectorAll(".source-panel .cm-line span")]
            .map((node) => getComputedStyle(node).color)
            .filter((color, index, colors) => colors.indexOf(color) === index),
          pageScrollTop: document.scrollingElement?.scrollTop ?? 0,
          pageScrollHeight: document.scrollingElement?.scrollHeight ?? 0,
          pageClientHeight: document.scrollingElement?.clientHeight ?? 0,
        };
      })()`);
      const boundaryGap =
        expanded.sourcePanel.top - expanded.editorScroll.bottom;
      const rightAlignmentError = Math.abs(
        expanded.editorSurface.right - expanded.sourcePanel.right,
      );
      if (
        !expanded.body ||
        !expanded.sourceSlot ||
        expanded.bodyOverflow !== "hidden" ||
        !["auto", "scroll"].includes(expanded.editorOverflowY) ||
        !["auto", "scroll"].includes(expanded.sourceOverflowY) ||
        !expanded.editorScrollable ||
        !expanded.sourceScrollable ||
        expanded.editorAfterOwnScroll <= 0 ||
        expanded.sourceAfterEditorScroll !== 0 ||
        expanded.sourceAfterOwnScroll <= 0 ||
        !expanded.syntaxColors.includes("rgb(169, 221, 248)") ||
        expanded.editorAfterSourceScroll !== expanded.editorAfterOwnScroll ||
        boundaryGap < 14 ||
        boundaryGap > 18 ||
        rightAlignmentError > 1.5 ||
        expanded.pageScrollTop !== 0 ||
        expanded.pageScrollHeight > expanded.pageClientHeight + 1
      ) {
        throw new Error(`Source split scrolling is invalid: ${JSON.stringify({ ...expanded, boundaryGap, rightAlignmentError })}`);
      }

      await evaluate(`document.querySelector(".source-collapse-button").click()`);
      await waitForEvaluation(`(() => ({
        ready:
          Boolean(document.querySelector(".source-toggle-row")) &&
          !document.querySelector(".source-panel"),
      }))()`, "collapsed source layout after close");
      console.log(
        JSON.stringify(
          { collapsed, expanded, boundaryGap, rightAlignmentError },
          null,
          2,
        ),
      );
      console.log("Targeted source layout regression passed");
      return;
    }

    if (scenario === "cursor-placement") {
      const readPlaceholderCaret = async (
        name,
        expectedPlaceholderCount,
        requireTopmostPlaceholder = false,
      ) =>
        waitForEvaluation(`(() => {
          const field =
            document.querySelector(".formula-line.is-active math-field") ??
            document.querySelector("math-field");
          const root = field?.shadowRoot;
          const placeholders = [...(root?.querySelectorAll(
            ".visualtex-structural-placeholder",
          ) ?? [])];
          const selected = placeholders.find(
            (placeholder) =>
              placeholder.classList.contains("ML__selected") ||
              placeholder.classList.contains("ML__placeholder-selected") ||
              Boolean(placeholder.closest(".ML__selected")),
          );
          const caret = selected?.querySelector(
            ":scope > .visualtex-structural-placeholder-caret",
          );
          const selectedBounds = selected?.getBoundingClientRect();
          const caretBounds = caret?.getBoundingClientRect();
          const placeholderBounds = placeholders
            .map((placeholder) => placeholder.getBoundingClientRect())
            .sort((first, second) => first.top - second.top);
          const nativeCarets = [...(root?.querySelectorAll(
            ".ML__caret, .ML__text-caret, .ML__latex-caret",
          ) ?? [])].filter((node) => {
            const style = getComputedStyle(node);
            const bounds = node.getBoundingClientRect();
            return (
              style.display !== "none" &&
              style.visibility !== "hidden" &&
              Number.parseFloat(style.opacity || "1") > 0.1 &&
              bounds.height > 0
            );
          });
          const caretLeftDelta =
            caretBounds && selectedBounds
              ? caretBounds.left - selectedBounds.left
              : 99;
          const caretRightDelta =
            caretBounds && selectedBounds
              ? caretBounds.right - selectedBounds.left
              : 99;
          const topmostSelected =
            !${requireTopmostPlaceholder} ||
            Boolean(
              selectedBounds &&
                placeholderBounds[0] &&
                Math.abs(selectedBounds.top - placeholderBounds[0].top) <= 1,
            );
          const selectionRange = field?.selection?.ranges?.[0] ?? null;
          const placeholderSelected =
            Boolean(selectionRange) &&
            Math.abs(selectionRange[1] - selectionRange[0]) === 1;
          return {
            ready:
              Boolean(field && selected && caret) &&
              placeholders.length === ${expectedPlaceholderCount} &&
              placeholderSelected &&
              caretLeftDelta < -0.5 &&
              caretRightDelta <= 0.75 &&
              nativeCarets.length === 0 &&
              topmostSelected,
            name: ${JSON.stringify(name)},
            value: field?.value ?? "",
            selection: field?.selection ?? null,
            placeholderCount: placeholders.length,
            selectedBounds: selectedBounds
              ? {
                  left: selectedBounds.left,
                  top: selectedBounds.top,
                  right: selectedBounds.right,
                  bottom: selectedBounds.bottom,
                }
              : null,
            caretBounds: caretBounds
              ? {
                  left: caretBounds.left,
                  top: caretBounds.top,
                  right: caretBounds.right,
                  bottom: caretBounds.bottom,
                }
              : null,
            caretLeftDelta,
            caretRightDelta,
            nativeCaretCount: nativeCarets.length,
            topmostSelected,
            placeholderSelected,
          };
        })()`, `${name} placeholder caret placement`);

      const typeBackslashOverPlaceholder = async (
        name,
        expectedPlaceholderCount,
      ) => {
        await key("\\", "Backslash", 220);
        return waitForEvaluation(`(() => {
          const field = document.querySelector(
            ".formula-line.is-active math-field",
          );
          const placeholders =
            field?.shadowRoot?.querySelectorAll(
              ".visualtex-structural-placeholder",
            ) ?? [];
          const rawLatex = [...(field?.shadowRoot?.querySelectorAll(
            ".ML__raw-latex",
          ) ?? [])]
            .map((node) => node.textContent ?? "")
            .join("");
          return {
            ready:
              placeholders.length === ${expectedPlaceholderCount} &&
              rawLatex.includes("\\\\"),
            value: field?.value ?? "",
            placeholderCount: placeholders.length,
            rawLatex,
            selection: field?.selection ?? null,
          };
        })()`, `${name} direct backslash replaces placeholder`);
      };

      await clearField();
      await typeText("\\frac");
      await waitForEvaluation(`(() => {
        const panel = document.getElementById("mathlive-suggestion-popover");
        return {
          ready:
            panel?.classList.contains("is-visible") &&
            panel.querySelector("li.ML__popover__current")?.dataset.command ===
              "\\\\frac",
          current:
            panel?.querySelector("li.ML__popover__current")?.dataset.command ??
            "",
        };
      })()`, "native fraction suggestion");
      await key(" ", "Space", 32);
      const fractionState = await readPlaceholderCaret(
        "fraction numerator",
        2,
        true,
      );
      const fractionTypedState = await typeBackslashOverPlaceholder(
        "fraction numerator",
        1,
      );

      await key("Enter", "Enter", 13);
      await waitForEvaluation(`(() => {
        const field = document.querySelector(
          ".formula-line.is-active math-field",
        );
        return { ready: Boolean(field?.isConnected && field.value === "") };
      })()`, "empty line for summation insertion");
      await evaluate(`(() => {
        if (!document.querySelector('[data-command-id="sum"]')) {
          document.querySelector(".sidebar-toggle")?.click();
        }
        return true;
      })()`);
      await waitForEvaluation(`(() => ({
        ready:
          Boolean(document.querySelector('[data-command-id="sum"]')) &&
          Boolean(document.querySelector('[data-command-id="int"]')),
      }))()`, "formula toolbar commands");
      await evaluate(
        `document.querySelector('[data-command-id="sum"]').click()`,
      );
      const sumState = await readPlaceholderCaret("summation", 3);
      const sumTypedState = await typeBackslashOverPlaceholder("summation", 2);

      await key("Enter", "Enter", 13);
      await waitForEvaluation(`(() => {
        const field = document.querySelector(
          ".formula-line.is-active math-field",
        );
        return { ready: Boolean(field?.isConnected && field.value === "") };
      })()`, "empty line for integral insertion");
      await evaluate(
        `document.querySelector('[data-command-id="int"]').click()`,
      );
      const integralState = await readPlaceholderCaret("integral", 4);
      const integralTypedState = await typeBackslashOverPlaceholder(
        "integral",
        3,
      );

      console.log(
        JSON.stringify(
          {
            fractionState,
            fractionTypedState,
            sumState,
            sumTypedState,
            integralState,
            integralTypedState,
          },
          null,
          2,
        ),
      );
      console.log("Targeted structural cursor placement regression passed");
      return;
    }

    if (scenario === "geometry") {
      await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        if (!field?.isConnected) return { ready: false };
        field.setValue("p+(z+r)+q+\\\\placeholder{}", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.dispatchEvent(new InputEvent("input", {
          bubbles: true,
          composed: true,
          inputType: "insertText",
        }));
        return { ready: field.isConnected };
      })()`, "stable formula field for geometry");
      await sleep(150);
      const geometry = await evaluate(`(() => {
        const rect = (element) => {
          const bounds = element?.getBoundingClientRect();
          return bounds
            ? { x: bounds.x, y: bounds.y, width: bounds.width, height: bounds.height }
            : null;
        };
        const field = document.querySelector("math-field");
        return {
          viewport: { width: innerWidth, height: innerHeight },
          title: rect(document.querySelector(".document-title-area input")),
          editorScroll: rect(document.querySelector(".editor-pane-scroll")),
          editorSurface: rect(document.querySelector(".editor-surface")),
          formulaLine: rect(document.querySelector(".formula-line")),
          mathfieldHost: rect(document.querySelector(".mathfield-host")),
          mathfield: rect(field),
          value: field?.value ?? "",
        };
      })()`);
      console.log(JSON.stringify(geometry, null, 2));
      console.log("Targeted geometry probe passed");
      return;
    }

    if (scenario === "wrapper-prefix") {
      await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        if (!field?.isConnected) return { ready: false };
        field.setValue("p+\\\\frac{z+n}{d}+q", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        let markerEnd = -1;
        for (let end = 1; end <= field.lastOffset; end += 1) {
          if (
            field.getValue(end - 1, end, "latex").trim() === "z" ||
            field.getElementInfo(end)?.latex?.trim() === "z"
          ) {
            markerEnd = end;
            break;
          }
        }
        if (markerEnd < 0) return { ready: false, value: field.value };
        field.focus();
        field.selection = { ranges: [[markerEnd, markerEnd]], direction: "none" };
        field.position = markerEnd;
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        const host = field.closest(".mathfield-host");
        const hostBounds = host?.getBoundingClientRect();
        const bounds = field.getElementInfo(markerEnd)?.bounds;
        if (host && hostBounds && bounds) {
          host.dataset.testWrapperPrefixAnchorTop = String(
            bounds.top - hostBounds.top + bounds.height / 2,
          );
        }
        return { ready: field.position === markerEnd, value: field.value, markerEnd };
      })()`, "fraction numerator anchor for partial wrapper command");

      await typeText("\\math");
      const selectionState = await waitForEvaluation(`(() => {
        const source = document.getElementById("mathlive-suggestion-popover");
        const stable = document.getElementById("visualtex-stable-native-input-popover");
        const current = source?.querySelector("li.ML__popover__current[data-command]");
        return {
          ready:
            source?.classList.contains("is-visible") &&
            current?.dataset.command === "\\\\mathbb" &&
            !document.querySelector(".suggestion-popup"),
          currentCommand: current?.dataset.command ?? "",
          sourceVisible: source?.classList.contains("is-visible") ?? false,
          stableVisible: stable?.classList.contains("is-visible") ?? false,
          customCandidateVisible: Boolean(document.querySelector(".suggestion-popup")),
          rawLatex: [...(document.querySelector("math-field")?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .map((node) => node.textContent ?? "")
            .join(""),
        };
      })()`, "mathbb selected from partial math input");

      await key(" ", "Space", 32);
      const pendingState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const host = field?.closest(".mathfield-host");
        const expectedTop = Number.parseFloat(
          host?.dataset.testWrapperPrefixAnchorTop ?? "NaN",
        );
        const actualTop = Number.parseFloat(
          host?.style.getPropertyValue("--pending-wrapper-top") ?? "NaN",
        );
        return {
          ready:
            field?.dataset.pendingWrapperCommand === "\\\\mathbb" &&
            host?.classList.contains("has-pending-wrapper-placeholder") &&
            (field?.shadowRoot?.querySelectorAll(".ML__raw-latex").length ?? -1) === 0 &&
            Math.abs(actualTop - expectedTop) <= 1.5,
          value: field?.value ?? "",
          pendingCommand: field?.dataset.pendingWrapperCommand ?? "",
          frameVisible: host?.classList.contains("has-pending-wrapper-placeholder") ?? false,
          expectedTop,
          actualTop,
          rawCount: field?.shadowRoot?.querySelectorAll(".ML__raw-latex").length ?? -1,
        };
      })()`, "partial math selection enters mathbb wrapper input");

      await key("A", "KeyA", 65);
      const insertedState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const probe = document.createElement("math-field");
        probe.style.display = "none";
        document.body.append(probe);
        probe.setValue("p+\\\\frac{z\\\\mathbb{A}+n}{d}+q", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        const expected = probe.value.replaceAll(" ", "");
        probe.remove();
        return {
          ready:
            field?.value.replaceAll(" ", "") === expected &&
            !field.dataset.pendingWrapperCommand,
          value: field?.value ?? "",
          expected,
          pendingCommand: field?.dataset.pendingWrapperCommand ?? "",
        };
      })()`, "partial mathbb wrapper inserts in original numerator slot");

      console.log(JSON.stringify({ selectionState, pendingState, insertedState }, null, 2));
      console.log("Targeted partial wrapper selection regression passed");
      return;
    }

    if (scenario === "native-space-selection") {
      await focusField();
      await typeText("\\the");
      const initialState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const panel = document.getElementById("mathlive-suggestion-popover");
        const items = [...(panel?.querySelectorAll("li[data-command]") ?? [])];
        const current = panel?.querySelector("li.ML__popover__current[data-command]");
        return {
          ready:
            panel?.classList.contains("is-visible") &&
            items.length >= 2 &&
            Boolean(current),
          commands: items.map((item) => item.dataset.command ?? ""),
          firstCommand: items[0]?.dataset.command ?? "",
          selectedCommand: current?.dataset.command ?? "",
          rawLatex: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
        };
      })()`, "native input-selection list for theta prefix");

      await key("ArrowDown", "ArrowDown", 40);
      const movedState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const source = document.getElementById("mathlive-suggestion-popover");
        const stable = document.getElementById("visualtex-native-input-suggestion-popover");
        const sourceCurrent = source?.querySelector("li.ML__popover__current[data-command]");
        const stableCurrent = stable?.querySelector("li.ML__popover__current[data-command]");
        return {
          ready:
            sourceCurrent?.dataset.command === "\\\\theta" &&
            stableCurrent?.dataset.command === "\\\\theta",
          sourceSelected: sourceCurrent?.dataset.command ?? "",
          stableSelected: stableCurrent?.dataset.command ?? "",
          remembered: field?.dataset.pendingNativeSuggestion ?? "",
        };
      })()`, "arrow key selects theta in the native input-selection list");

      const nativeSpaceStartedAt = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        window.__visualtexNativeSpaceTiming = { fieldSpaceCount: 0 };
        field?.addEventListener("keydown", (event) => {
          if (event.key === " " || event.code === "Space") {
            window.__visualtexNativeSpaceTiming.fieldSpaceCount += 1;
          }
        }, { capture: true, once: true });
        return performance.now();
      })()`);
      await key(" ", "Space", 32);
      const committedState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const source = document.getElementById("mathlive-suggestion-popover");
        const stable = document.getElementById(
          "visualtex-native-input-suggestion-popover",
        );
        const normalized = (field?.value ?? "").replaceAll(" ", "");
        return {
          ready:
            normalized === "\\\\theta" &&
            (field?.shadowRoot?.querySelectorAll(".ML__raw-latex").length ?? -1) === 0 &&
            !field?.dataset.pendingNativeSuggestion &&
            !source?.classList.contains("is-visible") &&
            !stable?.classList.contains("is-visible"),
          value: field?.value ?? "",
          normalized,
          pendingNativeSuggestion: field?.dataset.pendingNativeSuggestion ?? "",
          rawCount: field?.shadowRoot?.querySelectorAll(".ML__raw-latex").length ?? -1,
          sourceVisible: source?.classList.contains("is-visible") ?? false,
          stableVisible: stable?.classList.contains("is-visible") ?? false,
          elapsedMs: performance.now() - ${nativeSpaceStartedAt},
          fieldSpaceCount:
            window.__visualtexNativeSpaceTiming?.fieldSpaceCount ?? -1,
          semanticCount:
            normalized.split("\\\\theta").length - 1,
        };
      })()`, "Space commits the arrow-selected theta item");

      if (
        committedState.elapsedMs > 250 ||
        committedState.fieldSpaceCount !== 0 ||
        committedState.semanticCount !== 1
      ) {
        throw new Error(
          `Native Space selection was not uniquely owned by window capture: ${JSON.stringify(committedState)}`,
        );
      }

      if (initialState.firstCommand === "\\theta") {
        throw new Error(`Theta unexpectedly remained the first command: ${JSON.stringify(initialState)}`);
      }
      console.log(JSON.stringify({ initialState, movedState, committedState }, null, 2));
      console.log("Targeted native Space selection regression passed");
      return;
    }

    if (scenario === "native-space-ime-replay") {
      const readAlphaCandidateState = () => evaluate(`(() => {
        const field = document.querySelector("math-field");
        const panel = document.getElementById("mathlive-suggestion-popover");
        const current = panel?.querySelector("li.ML__popover__current[data-command]");
        const raw = [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
          .filter((node) => !node.classList.contains("ML__suggestion"))
          .map((node) => node.textContent ?? "")
          .join("");
        return {
          ready:
            raw === "\\\\al" &&
            [...(panel?.querySelectorAll("li[data-command]") ?? [])]
              .some((item) => item.dataset.command === "\\\\alpha"),
          raw,
          selected: current?.dataset.command ?? "",
          candidates: [...(panel?.querySelectorAll("li[data-command]") ?? [])]
            .map((item) => item.dataset.command ?? ""),
        };
      })()`);

      const prepareAlpha = async () => {
        await clearField();
        await focusField();
        await typeText("\\al");
        let state = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const panel = document.getElementById("mathlive-suggestion-popover");
          const raw = [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join("");
          const candidates = [...(panel?.querySelectorAll("li[data-command]") ?? [])]
            .map((item) => item.dataset.command ?? "");
          return {
            ready: raw === "\\\\al" && candidates.includes("\\\\alpha"),
            raw,
            candidates,
          };
        })()`, "native alpha candidate list before replay probe");
        for (let index = 0; index < 12; index += 1) {
          state = await readAlphaCandidateState();
          if (state.selected === "\\alpha") return state;
          await key("ArrowDown", "ArrowDown", 40);
        }
        throw new Error(`Could not select alpha for replay probe: ${JSON.stringify(state)}`);
      };

      const commitOnce = async () => {
        await key(" ", "Space", 32);
        return waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const normalized = (field?.value ?? "").replaceAll(" ", "");
          return {
            ready: normalized === "\\\\alpha",
            value: field?.value ?? "",
            normalized,
          };
        })()`, "single alpha before replay injection");
      };

      const readReplayState = () => evaluate(`(() => {
        const field = document.querySelector("math-field");
        const normalized = (field?.value ?? "").replaceAll(" ", "");
        return {
          value: field?.value ?? "",
          normalized,
          alphaCount: normalized.split("\\\\alpha").length - 1,
          raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
          mode: field?.mode ?? "",
        };
      })()`);

      const initial = await prepareAlpha();
      const committed = await commitOnce();
      const keypressReplay = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const sink = field?.shadowRoot?.querySelector('[part="keyboard-sink"]');
        if (!sink) return false;
        for (let index = 0; index < 2; index += 1) {
          sink.dispatchEvent(new KeyboardEvent("keypress", {
            key: " ",
            code: "Space",
            keyCode: 32,
            charCode: 32,
            which: 32,
            bubbles: true,
            composed: false,
            cancelable: true,
          }));
        }
        return true;
      })()`);
      await sleep(80);
      const afterKeypressReplay = await readReplayState();

      await prepareAlpha();
      await commitOnce();
      const inputReplay = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const sink = field?.shadowRoot?.querySelector('[part="keyboard-sink"]');
        if (!sink) return false;
        sink.dispatchEvent(new InputEvent("input", {
          inputType: "insertText",
          data: " ",
          bubbles: true,
          composed: false,
          cancelable: true,
        }));
        return true;
      })()`);
      await sleep(80);
      const afterInputReplay = await readReplayState();

      await prepareAlpha();
      await commitOnce();
      const internalSpaceReplay = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const sink = field?.shadowRoot?.querySelector('[part="keyboard-sink"]');
        if (!sink) return false;
        const down = new KeyboardEvent("keydown", {
          key: " ",
          code: "Space",
          keyCode: 32,
          which: 32,
          bubbles: true,
          composed: false,
          cancelable: true,
        });
        const keydownAllowed = sink.dispatchEvent(down);
        const press = new KeyboardEvent("keypress", {
          key: " ",
          code: "Space",
          keyCode: 32,
          charCode: 32,
          which: 32,
          bubbles: true,
          composed: false,
          cancelable: true,
        });
        const keypressAllowed = sink.dispatchEvent(press);
        return {
          keydownAllowed,
          keydownPrevented: down.defaultPrevented,
          keypressAllowed,
          keypressPrevented: press.defaultPrevented,
        };
      })()`);
      await sleep(80);
      const afterInternalSpaceReplay = await readReplayState();

      console.log(JSON.stringify({
        initial,
        committed,
        keypressReplay,
        afterKeypressReplay,
        inputReplay,
        afterInputReplay,
        internalSpaceReplay,
        afterInternalSpaceReplay,
      }, null, 2));
      console.log("Targeted native Space IME replay probe completed");
      return;
    }

    if (scenario === "raw-placeholder-visual") {
      await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        if (!field?.isConnected) return { ready: false };
        field.setValue("\\\\frac{\\\\placeholder{}}{\\\\placeholder{}}", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "placeholder",
          silenceNotifications: true,
        });
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        return {
          ready:
            field.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder").length === 2,
        };
      })()`, "selected fraction numerator placeholder before raw input");
      await typeText("\\the");

      const visualState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const root = field?.shadowRoot;
        const container = root?.querySelector(".ML__container");
        const rawNodes = [...(root?.querySelectorAll(".ML__raw-latex") ?? [])]
          .filter((node) => !node.classList.contains("ML__suggestion"));
        const rawText = rawNodes.map((node) => node.textContent ?? "").join("");
        const isTransparent = (value) =>
          value === "transparent" ||
          value === "rgba(0, 0, 0, 0)" ||
          /rgba\\([^)]*,\\s*0(?:\\.0+)?\\)$/.test(value);
        const inspected = container
          ? [container, ...container.querySelectorAll("*")]
          : [];
        const offenders = inspected.flatMap((node) => {
          if (!(node instanceof HTMLElement)) return [];
          if (
            node.classList.contains("visualtex-structural-placeholder") ||
            node.classList.contains("visualtex-structural-placeholder-caret")
          ) {
            return [];
          }
          const bounds = node.getBoundingClientRect();
          const style = getComputedStyle(node);
          if (
            bounds.width < 8 ||
            bounds.height < 8 ||
            isTransparent(style.backgroundColor)
          ) {
            return [];
          }
          return [{
            classes: node.className,
            tag: node.tagName,
            backgroundColor: style.backgroundColor,
            width: bounds.width,
            height: bounds.height,
          }];
        });
        const latexNodes = [...(root?.querySelectorAll(".ML__latex") ?? [])].map((node) => ({
          backgroundColor: getComputedStyle(node).backgroundColor,
          boxShadow: getComputedStyle(node).boxShadow,
          outlineWidth: getComputedStyle(node).outlineWidth,
        }));
        const selection = root?.querySelector(".ML__selection");
        return {
          ready:
            rawText === "\\\\the" &&
            container?.classList.contains("has-visualtex-raw-latex-command") &&
            offenders.length === 0 &&
            latexNodes.every((item) => isTransparent(item.backgroundColor) && item.boxShadow === "none") &&
            (!selection || getComputedStyle(selection).display === "none"),
          rawText,
          rawClass: container?.classList.contains("has-visualtex-raw-latex-command") ?? false,
          offenders,
          latexNodes,
          selectionDisplay: selection ? getComputedStyle(selection).display : "missing",
          remainingPlaceholderCount: root?.querySelectorAll(".visualtex-structural-placeholder").length ?? -1,
        };
      })()`, "raw LaTeX input has no large gray selection background");

      console.log(JSON.stringify({ visualState }, null, 2));
      console.log("Targeted raw placeholder visual regression passed");
      return;
    }

    if (scenario === "placeholder-selection") {
      await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        if (!field?.isConnected) return { ready: false };
        field.setValue("x+\\\\frac{\\\\alpha}{\\\\placeholder{}}+y", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.focus();
        field.position = field.lastOffset;
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        const placeholder = field.shadowRoot?.querySelector(
          ".visualtex-structural-placeholder",
        );
        const bounds = placeholder?.getBoundingClientRect();
        return {
          ready: Boolean(bounds && bounds.width > 5 && bounds.height > 5),
          left: bounds?.left ?? 0,
          top: bounds?.top ?? 0,
          width: bounds?.width ?? 0,
          height: bounds?.height ?? 0,
        };
      })()`, "alpha numerator followed by denominator placeholder geometry");
      const clickGeometry = await evaluate(`(() => {
        const placeholder = document.querySelector("math-field")?.shadowRoot?.querySelector(
          ".visualtex-structural-placeholder",
        );
        const bounds = placeholder?.getBoundingClientRect();
        return bounds
          ? {
              x: bounds.left + bounds.width / 2,
              y: bounds.top + bounds.height / 2,
            }
          : null;
      })()`);
      if (!clickGeometry) throw new Error("Could not locate denominator placeholder");
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: clickGeometry.x,
        y: clickGeometry.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: clickGeometry.x,
        y: clickGeometry.y,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });
      await sleep(180);

      const alphaPlaceholderState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const root = field?.shadowRoot;
        const container = root?.querySelector(".ML__container");
        const placeholder = root?.querySelector(".visualtex-structural-placeholder");
        const caret = placeholder?.querySelector(
          ":scope > .visualtex-structural-placeholder-caret",
        );
        const isTransparent = (value) =>
          value === "transparent" ||
          value === "rgba(0, 0, 0, 0)" ||
          /rgba\\([^)]*,\\s*0(?:\\.0+)?\\)$/.test(value);
        const highlighted = [
          ...(root?.querySelectorAll(
            ".ML__contains-highlight, .ML__highlight, .ML__selected",
          ) ?? []),
        ];
        const offenders = highlighted.flatMap((node) => {
          if (!(node instanceof HTMLElement)) return [];
          if (
            node === placeholder ||
            node.classList.contains("visualtex-structural-placeholder-caret")
          ) {
            return [];
          }
          const bounds = node.getBoundingClientRect();
          const style = getComputedStyle(node);
          if (
            bounds.width < 8 ||
            bounds.height < 8 ||
            isTransparent(style.backgroundColor)
          ) {
            return [];
          }
          return [{
            classes: node.className,
            backgroundColor: style.backgroundColor,
            width: bounds.width,
            height: bounds.height,
          }];
        });
        const selection = root?.querySelector(".ML__selection");
        return {
          ready:
            field?.value.replaceAll(" ", "") ===
              "x+\\\\frac{\\\\alpha}{\\\\placeholder{}}+y" &&
            offenders.length === 0,
          value: field?.value ?? "",
          placeholderEditingClass: container?.classList.contains(
            "has-visualtex-structural-placeholder-selection",
          ) ?? false,
          caretPresent: Boolean(caret),
          offenders,
          selectionDisplay: selection ? getComputedStyle(selection).display : "missing",
        };
      })()`, "alpha input leaves no gray fraction highlight around the next placeholder");

      const selectedPlaceholderGeometry = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const placeholder = field?.shadowRoot?.querySelector(
          ".visualtex-structural-placeholder.ML__selected, " +
          ".ML__selected .visualtex-structural-placeholder, " +
          ".visualtex-structural-placeholder.ML__placeholder-selected",
        );
        const bounds = placeholder?.getBoundingClientRect();
          return bounds
            ? {
                x: bounds.left + bounds.width / 2,
                y: bounds.top + bounds.height / 2,
                top: bounds.top,
              }
          : null;
      })()`);
      if (!selectedPlaceholderGeometry) {
        throw new Error("Could not locate selected placeholder geometry");
      }
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: selectedPlaceholderGeometry.x,
        y: selectedPlaceholderGeometry.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      await sleep(30);
      const heldSelectedPlaceholderState = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const root = field?.shadowRoot;
        const symbol = field?.placeholderSymbol || "▢";
        const placeholder = [...(root?.querySelectorAll(
          ".ML__cmr[data-atom-id], .ML__placeholder",
        ) ?? [])].find((node) =>
          node.classList.contains("visualtex-structural-placeholder") ||
          node.classList.contains("ML__placeholder") ||
          node.textContent?.trim() === symbol
        );
        const bounds = placeholder?.getBoundingClientRect();
        const style = placeholder ? getComputedStyle(placeholder) : null;
        return {
          pointerSelecting:
            field?.classList.contains("visualtex-pointer-selecting") ?? false,
          top: bounds?.top ?? -1,
            background: style?.backgroundColor ?? "",
            borderTopWidth: style?.borderTopWidth ?? "",
            color: style?.color ?? "",
          };
      })()`);
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: selectedPlaceholderGeometry.x,
        y: selectedPlaceholderGeometry.y,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });
      await sleep(100);
      const releasedSelectedPlaceholderTop = await evaluate(`(() => {
        const placeholder = document.querySelector("math-field")
          ?.shadowRoot?.querySelector(".visualtex-structural-placeholder");
        return placeholder?.getBoundingClientRect().top ?? -1;
      })()`);
      if (
        !heldSelectedPlaceholderState.pointerSelecting ||
        Math.abs(
          heldSelectedPlaceholderState.top - selectedPlaceholderGeometry.top,
        ) > 1 ||
        Math.abs(
          releasedSelectedPlaceholderTop - selectedPlaceholderGeometry.top,
        ) > 1 ||
        heldSelectedPlaceholderState.background !== "rgb(217, 237, 249)" ||
        heldSelectedPlaceholderState.borderTopWidth !== "0px" ||
        heldSelectedPlaceholderState.color !== "rgba(0, 0, 0, 0)"
      ) {
        throw new Error(
          `Selected placeholder shifted during pointer click: ${JSON.stringify({
            selectedPlaceholderGeometry,
            heldSelectedPlaceholderState,
            releasedSelectedPlaceholderTop,
          })}`,
        );
      }

      await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        if (!field?.isConnected) return { ready: false };
        field.setValue("a+b+\\\\frac{\\\\placeholder{}}{\\\\placeholder{}}+c+d", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.focus();
        field.position = field.lastOffset;
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        return { ready: field.value.includes("\\\\frac") };
      })()`, "formula setup for mouse range selection across placeholders");
      await sleep(180);
      const dragGeometry = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const base = field?.shadowRoot?.querySelector(".ML__base");
        const bounds = base?.getBoundingClientRect();
        return {
          ready: Boolean(bounds && bounds.width > 80 && bounds.height > 20),
          left: bounds?.left ?? 0,
          right: bounds?.right ?? 0,
          centerY: bounds ? bounds.top + bounds.height / 2 : 0,
          value: field?.value ?? "",
        };
      })()`, "settled formula geometry for mouse range selection across placeholders");

      await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const host = field?.closest(".mathfield-host");
        if (!field || !host) return;
        const bounds = field.getBoundingClientRect();
        host.dispatchEvent(new PointerEvent("pointerdown", {
          bubbles: true,
          composed: true,
          cancelable: true,
          button: 0,
          buttons: 1,
          clientX: bounds.left + 12,
          clientY: bounds.top + bounds.height / 2,
          pointerId: 1,
          pointerType: "mouse",
          isPrimary: true,
        }));
        field.selection = {
          ranges: [[0, field.lastOffset]],
          direction: "forward",
        };
      })()`);
      await sleep(80);

      const heldPointerSelectionState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const root = field?.shadowRoot;
        const placeholderSymbol = field?.placeholderSymbol || "▢";
        const placeholderAtoms = [
          ...(root?.querySelectorAll(".ML__cmr[data-atom-id], .ML__placeholder") ?? []),
        ].filter((node) =>
          node.classList.contains("ML__placeholder") ||
          node.textContent?.trim() === placeholderSymbol
        );
        const placeholderStyles = placeholderAtoms.map((node) => {
          const style = getComputedStyle(node);
          return {
            classes: node.className,
            backgroundColor: style.backgroundColor,
            borderTopWidth: style.borderTopWidth,
            color: style.color,
          };
        });
        const blackBoxPlaceholders = placeholderStyles.filter((style) =>
          style.backgroundColor === "rgba(0, 0, 0, 0)" ||
          style.backgroundColor === "transparent" ||
          style.borderTopWidth !== "0px" ||
          style.color !== "rgba(0, 0, 0, 0)"
        );
        const selection = root?.querySelector(".ML__selection");
        const selectionBounds = selection?.getBoundingClientRect();
        return {
          ready:
            Boolean(field?.classList.contains("visualtex-pointer-selecting")) &&
            Boolean(field && !field.selectionIsCollapsed) &&
            placeholderAtoms.length >= 2 &&
            blackBoxPlaceholders.length === 0 &&
            Boolean(
              selection &&
              getComputedStyle(selection).display !== "none" &&
              selectionBounds &&
              selectionBounds.width > 5
            ),
          pointerSelectingClass:
            field?.classList.contains("visualtex-pointer-selecting") ?? false,
          selectionCollapsed: field?.selectionIsCollapsed ?? true,
          placeholderCount: placeholderAtoms.length,
          blackBoxPlaceholderCount: blackBoxPlaceholders.length,
          placeholderStyles,
          placeholderCaretCount:
            root?.querySelectorAll(".visualtex-structural-placeholder-caret")
              .length ?? -1,
          selectionDisplay: selection ? getComputedStyle(selection).display : "missing",
          selectionWidth: selectionBounds?.width ?? 0,
        };
      })()`, "held pointer selection keeps placeholders blue and range continuous");

      await evaluate(`(() => {
        const field = document.querySelector("math-field");
        if (!field) return;
        const bounds = field.getBoundingClientRect();
        window.dispatchEvent(new PointerEvent("pointerup", {
          bubbles: true,
          composed: true,
          cancelable: true,
          button: 0,
          buttons: 0,
          clientX: bounds.right - 12,
          clientY: bounds.top + bounds.height / 2,
          pointerId: 1,
          pointerType: "mouse",
          isPrimary: true,
        }));
      })()`);
      await sleep(180);

      const rangeSelectionState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const root = field?.shadowRoot;
        const container = root?.querySelector(".ML__container");
        const ranges = field?.selection.ranges ?? [];
        const [start, end] = ranges[0] ?? [-1, -1];
        const selectedLatex =
          start >= 0 && end >= 0
            ? field.getValue(Math.min(start, end), Math.max(start, end), "latex")
            : "";
        const selection = root?.querySelector(".ML__selection");
        const selectionBounds = selection?.getBoundingClientRect();
        return {
          ready:
            Boolean(field && !field.selectionIsCollapsed) &&
            Math.abs(end - start) > 2 &&
            selectedLatex.includes("\\\\placeholder{}") &&
            !field.classList.contains("visualtex-pointer-selecting") &&
            !container?.classList.contains(
              "has-visualtex-structural-placeholder-selection",
            ) &&
            (root?.querySelectorAll(
              ".visualtex-structural-placeholder-caret",
            ).length ?? -1) === 0 &&
            Boolean(
              selection &&
                getComputedStyle(selection).display !== "none" &&
                selectionBounds &&
                selectionBounds.width > 5,
            ),
          ranges,
          selectionCollapsed: field?.selectionIsCollapsed ?? true,
          selectedLatex,
          pointerSelectingClass: field?.classList.contains(
            "visualtex-pointer-selecting",
          ) ?? false,
          placeholderEditingClass: container?.classList.contains(
            "has-visualtex-structural-placeholder-selection",
          ) ?? false,
          placeholderCaretCount:
            root?.querySelectorAll(".visualtex-structural-placeholder-caret")
              .length ?? -1,
          selectionDisplay: selection ? getComputedStyle(selection).display : "missing",
          selectionWidth: selectionBounds?.width ?? 0,
        };
      })()`, "pointer selection lifecycle preserves a range across structural placeholders");

      console.log(
        JSON.stringify(
          {
            alphaPlaceholderState,
            heldSelectedPlaceholderState,
            releasedSelectedPlaceholderTop,
            dragGeometry,
            heldPointerSelectionState,
            rangeSelectionState,
          },
          null,
          2,
        ),
      );
      console.log("Targeted placeholder range-selection regression passed");
      return;
    }

    if (scenario === "pointer-release-stability") {
      await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        if (!field?.isConnected) return { ready: false };
        field.setValue("a+b+c+d+e+f+g+h+i+j", {
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
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        const base = field.shadowRoot?.querySelector(".ML__base");
        const bounds = base?.getBoundingClientRect();
        return {
          ready: Boolean(bounds && bounds.width > 120 && bounds.height > 12),
          left: bounds?.left ?? 0,
          right: bounds?.right ?? 0,
          centerY: bounds ? bounds.top + bounds.height / 2 : 0,
        };
      })()`, "formula geometry for pointer release stability");
      await sleep(250);
      await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const persisted = JSON.parse(
          localStorage.getItem("visualtex-editor") || "{}",
        );
        return {
          ready:
            field?.value === "a+b+c+d+e+f+g+h+i+j" &&
            persisted.state?.lines?.[0]?.latex === field.value,
          value: field?.value ?? "",
          storeValue: persisted.state?.lines?.[0]?.latex ?? "",
        };
      })()`, "settled formula and store before pointer drag");
      const geometry = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        if (!field) return null;
        const compact = (value) => value.replace(/\\s+/g, "");
        const offsetForPrefix = (prefix) => {
          for (let offset = 0; offset <= field.lastOffset; offset += 1) {
            if (compact(field.getValue(0, offset, "latex")) === compact(prefix)) {
              return offset;
            }
          }
          return -1;
        };
        const startOffset = offsetForPrefix("a+b");
        const endOffset = offsetForPrefix("a+b+c+d+e+f+g+h");
        const contentBounds = field.shadowRoot
          ?.querySelector('[part="content"]')
          ?.getBoundingClientRect();
        if (!contentBounds || startOffset < 0 || endOffset < 0) return null;
        const y = (contentBounds.top + contentBounds.bottom) / 2;
        const samples = [];
        for (
          let x = Math.ceil(contentBounds.left);
          x <= Math.floor(contentBounds.right);
          x += 1
        ) {
          samples.push({
            x,
            offset: field.getOffsetFromPoint(x, y, { bias: 0 }),
          });
        }
        if (!samples.length) return null;
        const sampleForOffset = (targetOffset) => {
          const exact = samples.filter(
            (sample) => sample.offset === targetOffset,
          );
          if (exact.length) return exact[Math.floor(exact.length / 2)];
          return samples.reduce((best, sample) =>
            Math.abs(sample.offset - targetOffset) <
            Math.abs(best.offset - targetOffset)
              ? sample
              : best
          );
        };
        const startSample = sampleForOffset(startOffset);
        const endSample = sampleForOffset(endOffset);
        return {
          startOffset,
          endOffset,
          mappedStartOffset: startSample.offset,
          mappedEndOffset: endSample.offset,
          startX: startSample.x,
          endX: endSample.x,
          leftX: contentBounds.left + 2,
          rightX: contentBounds.right - 2,
          y,
        };
      })()`);
      if (!geometry) throw new Error("Could not resolve pointer release geometry");

      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: geometry.startX,
        y: geometry.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseMoved",
        x: (geometry.startX + geometry.endX) / 2,
        y: geometry.y,
        button: "left",
        buttons: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseMoved",
        x: geometry.endX,
        y: geometry.y,
        button: "left",
        buttons: 1,
      });
      await sleep(100);
      const dragged = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          range: field?.selection?.ranges?.at(-1) ?? null,
          pointerSelecting:
            field?.classList.contains("visualtex-pointer-selecting") ?? false,
          multiLineSelecting:
            field?.classList.contains("has-visualtex-multi-line-selection") ?? false,
          mathLiveTracking:
            Boolean(field?.shadowRoot?.querySelector(".tracking")),
          contentHasPointerCapture:
            field?.shadowRoot
              ?.querySelector('[part="content"]')
              ?.hasPointerCapture?.(1) ?? false,
        };
      })()`);
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: geometry.endX,
        y: geometry.y,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });
      await sleep(120);

      const readSelection = async () => evaluate(`(() => {
        const field = document.querySelector("math-field");
        const range = field?.selection?.ranges?.at(-1) ?? null;
        const start = range ? Math.min(range[0], range[1]) : -1;
        const end = range ? Math.max(range[0], range[1]) : -1;
        return {
          value: field?.value ?? "",
          range,
          selectedLatex:
            field && start >= 0 && end >= 0
              ? field.getValue(start, end, "latex")
              : "",
          pointerSelecting:
            field?.classList.contains("visualtex-pointer-selecting") ?? false,
          mathLiveTracking:
            Boolean(field?.shadowRoot?.querySelector(".tracking")),
          contentHasPointerCapture:
            field?.shadowRoot
              ?.querySelector('[part="content"]')
              ?.hasPointerCapture?.(1) ?? false,
        };
      })()`);
      const released = await readSelection();
      if (
        !released.range ||
        released.range[0] === released.range[1] ||
        released.pointerSelecting ||
        released.mathLiveTracking ||
        released.contentHasPointerCapture
      ) {
        throw new Error(
          `Pointer drag did not finish with a stable range: ${JSON.stringify({ geometry, dragged, released })}`,
        );
      }

      await client.send("Input.dispatchMouseEvent", {
        type: "mouseMoved",
        x: geometry.leftX,
        y: geometry.y,
        button: "none",
        buttons: 0,
      });
      await sleep(100);
      const movedLeft = await readSelection();
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseMoved",
        x: geometry.rightX,
        y: geometry.y,
        button: "none",
        buttons: 0,
      });
      await sleep(100);
      const movedRight = await readSelection();

      const stableSelection =
        JSON.stringify(released.range) === JSON.stringify(movedLeft.range) &&
        JSON.stringify(released.range) === JSON.stringify(movedRight.range) &&
        released.selectedLatex === movedLeft.selectedLatex &&
        released.selectedLatex === movedRight.selectedLatex &&
        !movedLeft.mathLiveTracking &&
        !movedRight.mathLiveTracking &&
        !movedLeft.contentHasPointerCapture &&
        !movedRight.contentHasPointerCapture;
      if (!stableSelection) {
        throw new Error(
          `Selection followed the pointer after mouse release: ${JSON.stringify({ released, movedLeft, movedRight })}`,
        );
      }

      console.log(JSON.stringify({ geometry, dragged, released, movedLeft, movedRight }, null, 2));
      console.log("Targeted pointer release stability regression passed");
      return;
    }

    if (scenario === "candidate-query-reset") {
      await focusField();
      await clearField();
      await typeText("\\int");
      await waitForEvaluation(`(() => {
        const panel = document.getElementById("mathlive-suggestion-popover");
        return {
          ready:
            panel?.classList.contains("is-visible") &&
            panel.querySelector("li.ML__popover__current")?.dataset.command === "\\\\int",
          current: panel?.querySelector("li.ML__popover__current")?.dataset.command ?? "",
        };
      })()`, "integral selected in native input-selection popover");
      await key(" ", "Space", 32);
      const confirmedState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const popup = document.querySelector(".suggestion-popup");
        const query = document.querySelector(".editor-surface")?.dataset.commandQuery ?? "";
        return {
          ready:
            field?.value.includes("\\\\int") &&
            Boolean(popup) &&
            query === "\\\\int" &&
            (field.shadowRoot?.querySelectorAll(".ML__raw-latex").length ?? -1) === 0,
          value: field?.value ?? "",
          query,
          popupVisible: Boolean(popup),
        };
      })()`, "confirmed integral opens VisualTeX command candidate popup");

      const twoStageStates = [];
      await key("Enter", "Enter", 13);
      twoStageStates.push(
        await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const value = field?.value ?? "";
          const selection = field?.selection ?? null;
          return {
            ready:
              value.startsWith("\\\\int_") &&
              (value.match(/\\\\placeholder\{\}/g) ?? []).length === 4 &&
              document.querySelectorAll("math-field").length === 1 &&
              selection?.ranges?.[0]?.[0] !== selection?.ranges?.[0]?.[1],
            command: "\\\\int",
            value,
            selection,
          };
        })()`, "second integral confirmation inserts the VisualTeX structure"),
      );

      for (const [command, expectedPlaceholderCount] of [
        ["\\sum", 3],
        ["\\lim", 3],
      ]) {
        await clearField();
        await typeText(command);
        await waitForEvaluation(`(() => {
          const panel = document.getElementById("mathlive-suggestion-popover");
          return {
            ready:
              panel?.classList.contains("is-visible") &&
              panel.querySelector("li.ML__popover__current")?.dataset.command ===
                ${JSON.stringify(command)},
          };
        })()`, `native ${command} awaits its first confirmation`);
        await key(" ", "Space", 32);
        const firstConfirmation = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          return {
            ready:
              field?.value === ${JSON.stringify(command)} &&
              Boolean(document.querySelector(".suggestion-popup")) &&
              document.querySelectorAll("math-field").length === 1,
            value: field?.value ?? "",
          };
        })()`, `first ${command} confirmation keeps the bare command`);
        await key("Enter", "Enter", 13);
        const secondConfirmation = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const value = field?.value ?? "";
          const selection = field?.selection ?? null;
          return {
            ready:
              value.startsWith(${JSON.stringify(command + "_")}) &&
              (value.match(/\\\\placeholder\{\}/g) ?? []).length ===
                ${expectedPlaceholderCount} &&
              document.querySelectorAll("math-field").length === 1 &&
              selection?.ranges?.[0]?.[0] !== selection?.ranges?.[0]?.[1],
            value,
            selection,
          };
        })()`, `second ${command} confirmation inserts the VisualTeX structure`);
        twoStageStates.push({
          command,
          firstConfirmation,
          secondConfirmation,
        });
      }

      await clearField();
      await typeText("\\int");
      await waitForEvaluation(`(() => {
        const panel = document.getElementById("mathlive-suggestion-popover");
        return {
          ready:
            panel?.classList.contains("is-visible") &&
            panel.querySelector("li.ML__popover__current")?.dataset.command === "\\\\int",
        };
      })()`, "integral selected again before candidate query reset");
      await key(" ", "Space", 32);
      await waitForEvaluation(`(() => ({
        ready:
          document.querySelector("math-field")?.value === "\\\\int" &&
          Boolean(document.querySelector(".suggestion-popup")),
      }))()`, "integral candidate restored before query reset");

      await typeText("\\");
      const resetState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const rawNodes = field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [];
        const rawLatex = [...rawNodes]
          .filter((node) => !node.classList.contains("ML__suggestion"))
          .map((node) => node.textContent ?? "")
          .join("");
        const query = document.querySelector(".editor-surface")?.dataset.commandQuery ?? "";
        const stable = document.getElementById("visualtex-stable-native-input-popover");
        return {
          ready:
            rawLatex === "\\\\" &&
            query === "" &&
            !document.querySelector(".suggestion-popup") &&
            !stable?.classList.contains("is-visible"),
          value: field?.value ?? "",
          rawLatex,
          query,
          customCandidateVisible: Boolean(document.querySelector(".suggestion-popup")),
          nativeInputSelectionVisible: stable?.classList.contains("is-visible") ?? false,
        };
      })()`, "lone backslash clears stale integral command candidate query");

      console.log(
        JSON.stringify({ confirmedState, twoStageStates, resetState }, null, 2),
      );
      console.log("Targeted command candidate query reset regression passed");
      return;
    }

    if (scenario === "structured-chinese-ime") {
      const commitImeText = async (text) => {
        await sleep(120);
        await client.send("Input.imeSetComposition", {
          text,
          selectionStart: Array.from(text).length,
          selectionEnd: Array.from(text).length,
        });
        await sleep(80);
        await client.send("Input.insertText", { text });
        await client.send("Input.imeSetComposition", {
          text: "",
          selectionStart: 0,
          selectionEnd: 0,
        });
        await sleep(260);
      };
      const selectedPlaceholderState = async (description) =>
        waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const range = field?.selection?.ranges?.at(-1) ?? null;
          const selectedLatex = range
            ? field.getValue(
                Math.min(range[0], range[1]),
                Math.max(range[0], range[1]),
                "latex",
              ).trim()
            : "";
          field?.focus();
          field?.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
          return {
            ready:
              Boolean(field) &&
              field.value.includes("\\\\placeholder{}") &&
              selectedLatex === "\\\\placeholder{}",
            value: field?.value ?? "",
            selectedLatex,
            selection: field?.selection ?? null,
          };
        })()`, description);
      const readFieldState = async () =>
        evaluate(`(() => {
          const field = document.querySelector("math-field");
          const range = field?.selection?.ranges?.at(-1) ?? null;
          return {
            value: field?.value ?? "",
            placeholderCount:
              (field?.value.match(/\\\\placeholder\\{\\}/g) ?? []).length,
            selectedLatex: range
              ? field.getValue(
                  Math.min(range[0], range[1]),
                  Math.max(range[0], range[1]),
                  "latex",
                ).trim()
              : "",
            selection: field?.selection ?? null,
          };
        })()`);
      const clickToolbarCommand = async (commandId) => {
        const selector = `[data-command-id="${commandId}"]`;
        await waitForEvaluation(`(() => {
          const button = document.querySelector(${JSON.stringify(selector)});
          if (!(button instanceof HTMLElement)) return { ready: false };
          const rect = button.getBoundingClientRect();
          return { ready: rect.width > 0 && rect.height > 0 };
        })()`, `${commandId} toolbar command`);
        await evaluate(`document.querySelector(${JSON.stringify(selector)})?.click()`);
        return selectedPlaceholderState(`${commandId} selected placeholder`);
      };

      const states = {};
      const toolbarCases = [
        {
          id: "power",
          name: "superscript",
          expected: (value) => value.includes("^{\\text{中文}}"),
          expectedPlaceholderCount: 0,
        },
        {
          id: "subscript",
          name: "subscript",
          expected: (value) => value.includes("_{\\text{中文}}"),
          expectedPlaceholderCount: 0,
        },
        {
          id: "sqrt",
          name: "square root",
          expected: (value) => value.includes("\\sqrt{\\text{中文}}"),
          expectedPlaceholderCount: 0,
        },
        {
          id: "int",
          name: "integral",
          expected: (value) =>
            value.includes("\\int") && value.includes("\\text{中文}"),
          expectedPlaceholderCount: 3,
        },
        {
          id: "matrix2",
          name: "matrix",
          expected: (value) =>
            value.includes("\\begin{bmatrix}") &&
            value.includes("\\text{中文}") &&
            value.includes("\\end{bmatrix}"),
          expectedPlaceholderCount: 3,
        },
      ];
      for (const testCase of toolbarCases) {
        await clearField();
        const before = await clickToolbarCommand(testCase.id);
        await commitImeText("中文");
        const after = await readFieldState();
        if (
          !testCase.expected(after.value) ||
          after.placeholderCount !== testCase.expectedPlaceholderCount
        ) {
          throw new Error(
            `Toolbar ${testCase.name} lost its structure during Chinese IME: ${JSON.stringify({ before, after })}`,
          );
        }
        states[testCase.name] = { before, after };
      }

      await clearField();
      const fractionBefore = await clickToolbarCommand("frac");
      await commitImeText("分子");
      const fractionNumerator = await readFieldState();
      if (
        !fractionNumerator.value.includes(
          "\\frac{\\text{分子}}{\\placeholder{}}",
        ) ||
        fractionNumerator.placeholderCount !== 1
      ) {
        throw new Error(
          `Fraction numerator composition changed another slot: ${JSON.stringify({ fractionBefore, fractionNumerator })}`,
        );
      }
      await evaluate(`(() => {
        const field = document.querySelector("math-field");
        if (!field) return false;
        for (let offset = 1; offset <= field.lastOffset; offset += 1) {
          if (
            field.getValue(offset - 1, offset, "latex").trim() ===
            "\\\\placeholder{}"
          ) {
            field.selection = {
              ranges: [[offset - 1, offset]],
              direction: "none",
            };
            field.focus();
            field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
            return true;
          }
        }
        return false;
      })()`);
      const denominatorSelected = await selectedPlaceholderState(
        "fraction denominator selected placeholder",
      );
      await commitImeText("分母");
      const fractionComplete = await readFieldState();
      if (
        !fractionComplete.value.includes(
          "\\frac{\\text{分子}}{\\text{分母}}",
        ) ||
        fractionComplete.placeholderCount !== 0
      ) {
        throw new Error(
          `Fraction denominator composition lost the fraction: ${JSON.stringify({ denominatorSelected, fractionComplete })}`,
        );
      }
      states.fraction = {
        before: fractionBefore,
        numerator: fractionNumerator,
        denominatorSelected,
        complete: fractionComplete,
      };

      await clearField();
      const undoBefore = await clickToolbarCommand("power");
      await commitImeText("中文");
      const undoAfterComposition = await readFieldState();
      await key("z", "KeyZ", 90, 4);
      const undoState = await readFieldState();
      if (
        !undoState.value.includes("\\placeholder{}") ||
        !undoState.value.includes("^")
      ) {
        throw new Error(
          `Undo did not restore the structured placeholder transaction: ${JSON.stringify({ undoBefore, undoAfterComposition, undoState })}`,
        );
      }
      await key("z", "KeyZ", 90, 12);
      const redoState = await readFieldState();
      if (!redoState.value.includes("^{\\text{中文}}")) {
        throw new Error(
          `Redo did not restore structured Chinese composition: ${JSON.stringify({ undoState, redoState })}`,
        );
      }
      states.history = {
        before: undoBefore,
        composed: undoAfterComposition,
        undo: undoState,
        redo: redoState,
      };

      console.log(JSON.stringify(states, null, 2));
      console.log("Targeted structured Chinese IME regression passed");
      return;
    }

    if (scenario === "structural-placeholder") {
      const placeholderCases = [
        {
          name: "inline selected baseline",
          source: String.raw`\placeholder{}d\placeholder{}`,
          expectedCount: 2,
        },
        {
          name: "fraction",
          source: String.raw`\frac{\placeholder{}}{\placeholder{}}`,
          expectedCount: 2,
        },
        {
          name: "integral limits and integrand",
          source: String.raw`\int_{\placeholder{}}^{\placeholder{}}\placeholder{}\,dx`,
          expectedCount: 3,
        },
        {
          name: "square root",
          source: String.raw`\sqrt{\placeholder{}}`,
          expectedCount: 1,
        },
        {
          name: "superscript",
          source: String.raw`x^{\placeholder{}}`,
          expectedCount: 1,
        },
        {
          name: "subscript",
          source: String.raw`x_{\placeholder{}}`,
          expectedCount: 1,
        },
        {
          name: "matrix cells",
          source: String.raw`\begin{matrix}\placeholder{}&a\\b&\placeholder{}\end{matrix}`,
          expectedCount: 2,
        },
      ];

      const styleStates = [];
      for (const testCase of placeholderCases) {
        await waitForEvaluation(
          `(() => ({ ready: Boolean(document.querySelector("math-field")?.isConnected) }))()`,
          `stable field for placeholder case: ${testCase.name}`,
        );
        await evaluate(`(() => {
          const field = document.querySelector("math-field");
          field.setValue(${JSON.stringify(testCase.source)}, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceAll",
            selectionMode: "placeholder",
            silenceNotifications: true,
          });
          field.focus();
          field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        })()`);
        await sleep(120);
        const state = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const styleNode = field?.shadowRoot?.getElementById(
            "visualtex-structural-placeholder-style",
          );
          const placeholders = [...(field?.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder") ?? [])];
          const styles = placeholders.map((placeholder) => {
            const bounds = placeholder.getBoundingClientRect();
            const style = getComputedStyle(placeholder);
            return {
              top: bounds.top,
              width: bounds.width,
              height: bounds.height,
              ratio: bounds.height > 0 ? bounds.width / bounds.height : 99,
              borderTopWidth: style.borderTopWidth,
              borderRightWidth: style.borderRightWidth,
              borderBottomWidth: style.borderBottomWidth,
              borderLeftWidth: style.borderLeftWidth,
              borderStyle: style.borderStyle,
              borderRadius: style.borderRadius,
              backgroundColor: style.backgroundColor,
              color: style.color,
              opacity: style.opacity,
              boxShadow: style.boxShadow,
              selected:
                placeholder.classList.contains("ML__selected") ||
                Boolean(placeholder.closest(".ML__selected")),
              parentClasses: placeholder.parentElement?.className ?? "",
            };
          });
          const validColors = new Set([
            "rgb(217, 237, 249)",
            "rgb(207, 232, 247)",
          ]);
          const selectedStyle = styles.find((item) => item.selected);
          const unselectedStyle = styles.find((item) => !item.selected);
          const baselineDelta =
            selectedStyle && unselectedStyle
              ? Math.abs(selectedStyle.top - unselectedStyle.top)
              : 0;
          return {
            ready:
              Boolean(styleNode) &&
              placeholders.length === ${testCase.expectedCount} &&
              (${JSON.stringify(testCase.name)} !== "inline selected baseline" ||
                baselineDelta <= 1) &&
              styles.every(
                (item) =>
                  item.width > 7 &&
                  item.height > 6 &&
                  item.ratio > 0.35 &&
                  item.ratio < 0.75 &&
                  item.borderTopWidth === "0px" &&
                  item.borderRightWidth === "0px" &&
                  item.borderBottomWidth === "0px" &&
                  item.borderLeftWidth === "0px" &&
                  validColors.has(item.backgroundColor) &&
                  item.boxShadow === "none" &&
                  Number.parseFloat(item.opacity) >= 0.99,
              ),
            name: ${JSON.stringify(testCase.name)},
            value: field?.value ?? "",
            styleInstalled: Boolean(styleNode),
            count: placeholders.length,
            baselineDelta,
            styles,
          };
        })()`, `AxMath-style structural placeholders: ${testCase.name}`);
        styleStates.push(state);
      }

      await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.setValue("\\\\frac{\\\\placeholder{}}{\\\\placeholder{}}", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "placeholder",
          silenceNotifications: true,
        });
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        field.dataset.testStructuralPlaceholderPosition = String(field.position);
      })()`);
      await sleep(100);
      const emptyState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const placeholders = field?.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder") ?? [];
        const selected = field?.shadowRoot?.querySelector(
          ".visualtex-structural-placeholder.ML__selected, .ML__selected .visualtex-structural-placeholder, .visualtex-structural-placeholder.ML__placeholder-selected",
        );
        const caret = selected?.querySelector(
          ":scope > .visualtex-structural-placeholder-caret",
        );
        const caretStyle = caret ? getComputedStyle(caret) : null;
        const selectionOverlay = field?.shadowRoot?.querySelector(".ML__selection");
        const selectionOverlayDisplay = selectionOverlay
          ? getComputedStyle(selectionOverlay).display
          : "missing";
        const selectedAncestors = [];
        let ancestor = selected?.parentElement ?? null;
        while (ancestor) {
          if (ancestor.classList.contains("ML__selected")) {
            const style = getComputedStyle(ancestor);
            selectedAncestors.push({
              backgroundColor: style.backgroundColor,
              boxShadow: style.boxShadow,
              outlineWidth: style.outlineWidth,
            });
          }
          ancestor = ancestor.parentElement;
        }
        const graySelectionCleared =
          (selectionOverlayDisplay === "missing" || selectionOverlayDisplay === "none") &&
          selectedAncestors.every(
            (style) =>
              style.backgroundColor === "rgba(0, 0, 0, 0)" &&
              style.boxShadow === "none" &&
              style.outlineWidth === "0px",
          );
        return {
          ready:
            placeholders.length === 2 &&
            Boolean(selected) &&
            Boolean(caret) &&
            graySelectionCleared &&
            Number.parseFloat(caretStyle?.borderLeftWidth ?? "0") >= 1 &&
            Number.parseFloat(caretStyle?.left ?? "0") < 0 &&
            caretStyle?.animationName.includes("visualtex-placeholder-caret-blink"),
          value: field?.value ?? "",
          placeholderCount: placeholders.length,
          selected: Boolean(selected),
          caretPresent: Boolean(caret),
          caretBorderWidth: caretStyle?.borderLeftWidth ?? "",
          caretLeft: caretStyle?.left ?? "",
          caretAnimation: caretStyle?.animationName ?? "",
          caretOpacity: caretStyle?.opacity ?? "",
          selectionOverlayDisplay,
          selectedAncestors,
          graySelectionCleared,
          rawLatex: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .map((node) => node.textContent ?? "")
            .join(""),
        };
      })()`, "selected empty fraction numerator placeholder without outer gray selection");

      const hiddenBlinkState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const caret = field?.shadowRoot?.querySelector(
          ".visualtex-structural-placeholder-caret",
        );
        const opacity = Number.parseFloat(
          caret ? getComputedStyle(caret).opacity : "1",
        );
        return {
          ready: Boolean(caret) && opacity <= 0.1,
          opacity,
        };
      })()`, "structural placeholder caret hidden phase", 2500);

      const visibleBlinkState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const caret = field?.shadowRoot?.querySelector(
          ".visualtex-structural-placeholder-caret",
        );
        const opacity = Number.parseFloat(
          caret ? getComputedStyle(caret).opacity : "0",
        );
        return {
          ready: Boolean(caret) && opacity >= 0.9,
          opacity,
        };
      })()`, "structural placeholder caret visible phase", 2500);

      await key("\\", "Backslash", 220);
      const typedState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const placeholders = field?.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder") ?? [];
        const rawLatex = [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
          .map((node) => node.textContent ?? "")
          .join("");
        return {
          ready:
            placeholders.length === 1 &&
            rawLatex.includes("\\\\"),
          value: field?.value ?? "",
          placeholderCount: placeholders.length,
          rawLatex,
          mode: field?.mode ?? "",
        };
      })()`, "typing a backslash replaces the selected structural placeholder");

      await key("Backspace", "Backspace", 8);
      const restoredState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const placeholders = field?.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder") ?? [];
        const rawLatex = [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
          .map((node) => node.textContent ?? "")
          .join("");
        const selected = field?.shadowRoot?.querySelector(
          ".visualtex-structural-placeholder.ML__selected, .ML__selected .visualtex-structural-placeholder",
        );
        const restoredPlaceholder = selected ?? placeholders[0] ?? null;
        const style = restoredPlaceholder
          ? getComputedStyle(restoredPlaceholder)
          : null;
        const caret = selected?.querySelector(
          ":scope > .visualtex-structural-placeholder-caret",
        );
        const caretStyle = caret ? getComputedStyle(caret) : null;
        const savedPosition = Number.parseInt(
          field?.dataset.testStructuralPlaceholderPosition ?? "-1",
          10,
        );
        return {
          ready:
            placeholders.length === 2 &&
            Boolean(selected) &&
            Boolean(caret) &&
            rawLatex === "" &&
            field?.position === savedPosition &&
            style?.borderTopWidth === "0px" &&
            Number.parseFloat(caretStyle?.borderLeftWidth ?? "0") >= 1 &&
            Number.parseFloat(caretStyle?.left ?? "0") < 0 &&
            caretStyle?.animationName.includes("visualtex-placeholder-caret-blink") &&
            ["rgb(217, 237, 249)", "rgb(207, 232, 247)"].includes(
              style?.backgroundColor ?? "",
            ),
          value: field?.value ?? "",
          placeholderCount: placeholders.length,
          rawLatex,
          selectedBackground: style?.backgroundColor ?? "",
          selectedBorder: style?.borderTopWidth ?? "",
          position: field?.position ?? -1,
          savedPosition,
          selectedPlaceholder: Boolean(selected),
          caretPresent: Boolean(caret),
          caretBorderWidth: caretStyle?.borderLeftWidth ?? "",
          caretLeft: caretStyle?.left ?? "",
          caretAnimation: caretStyle?.animationName ?? "",
          mode: field?.mode ?? "",
        };
      })()`, "deleting the backslash restores the empty structural placeholder");

  const persistentPlaceholderCases = [
    {
      name: "fraction",
      source: String.raw`\frac{\placeholder{}}{\placeholder{}}`,
      expectedCount: 2,
    },
    {
      name: "fraction-denominator",
      source: String.raw`\frac{x}{\placeholder{}}`,
      expectedCount: 1,
    },
    {
      name: "square-root",
      source: String.raw`\sqrt{\placeholder{}}`,
      expectedCount: 1,
    },
    {
      name: "superscript",
      source: String.raw`x^{\placeholder{}}`,
      expectedCount: 1,
    },
    {
      name: "subscript",
      source: String.raw`x_{\placeholder{}}`,
      expectedCount: 1,
    },
    {
      name: "integral-limits",
      source: String.raw`\int_{\placeholder{}}^{\placeholder{}}`,
      expectedCount: 2,
    },
    {
      name: "integral-lower-limit",
      source: String.raw`\int_{\placeholder{}}^{n}`,
      expectedCount: 1,
    },
    {
      name: "integral-upper-limit",
      source: String.raw`\int_{0}^{\placeholder{}}`,
      expectedCount: 1,
    },
    {
      name: "summation-limits",
      source: String.raw`\sum_{\placeholder{}}^{\placeholder{}}`,
      expectedCount: 2,
    },
    {
      name: "summation-lower-limit",
      source: String.raw`\sum_{\placeholder{}}^{n}`,
      expectedCount: 1,
    },
    {
      name: "summation-upper-limit",
      source: String.raw`\sum_{i=0}^{\placeholder{}}`,
      expectedCount: 1,
    },
    {
      name: "dot-accent",
      source: String.raw`\dot{\placeholder{}}`,
      expectedCount: 1,
    },
    {
      name: "hat-accent",
      source: String.raw`\hat{\placeholder{}}`,
      expectedCount: 1,
    },
    {
      name: "vector-accent",
      source: String.raw`\vec{\placeholder{}}`,
      expectedCount: 1,
    },
    {
      name: "triple-dot-accent",
      source: String.raw`\dddot{\placeholder{}}`,
      expectedCount: 1,
    },
    {
      name: "quadruple-dot-accent",
      source: String.raw`\ddddot{\placeholder{}}`,
      expectedCount: 1,
    },
    {
      name: "matrix-cell",
      source: String.raw`\begin{matrix}\placeholder{}&\placeholder{}\end{matrix}`,
      expectedCount: 2,
    },
    {
      name: "matrix-second-cell",
      source: String.raw`\begin{matrix}x&\placeholder{}\end{matrix}`,
      expectedCount: 1,
    },
  ];
  const persistentPlaceholderStates = [];
  for (const testCase of persistentPlaceholderCases) {
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.setValue(${JSON.stringify(testCase.source)}, {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "placeholder",
        silenceNotifications: true,
      });
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({
        preventScroll: true,
      });
      return true;
    })()`);
    await sleep(100);

    const initial = await waitForEvaluation(`(() => {
      const field = document.querySelector("math-field");
      const placeholders = Array.from(
        field?.shadowRoot?.querySelectorAll(
          ".visualtex-structural-placeholder",
        ) ?? [],
      );
      const selected = placeholders.find(
        (node) =>
          node.classList.contains("ML__placeholder-selected") ||
          node.classList.contains("ML__selected") ||
          Boolean(node.closest(".ML__selected")),
      );
      const bounds = selected?.getBoundingClientRect();
      const selectedRange = field?.selection?.ranges?.[0];
      const background = selected
        ? getComputedStyle(selected).backgroundColor
        : "";
      const beforeBackground = selected
        ? getComputedStyle(selected, "::before").backgroundColor
        : "";
      return {
        ready:
          placeholders.length === ${testCase.expectedCount} &&
          Boolean(selected) &&
          Boolean(bounds),
        count: placeholders.length,
        value: field?.value ?? "",
        selection: field?.selection ?? null,
        rangeLatex: selectedRange
          ? field.getValue(selectedRange[0], selectedRange[1], "latex")
          : "",
        background,
        beforeBackground,
        top: bounds?.top ?? -1,
        width: bounds?.width ?? -1,
        height: bounds?.height ?? -1,
      };
    })()`, `initial persistent placeholder: ${testCase.name}`);

    const cycles = [];
    for (let cycle = 0; cycle < 2; cycle += 1) {
      await key("x", "KeyX", 88);
      const typed = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const count = field?.shadowRoot?.querySelectorAll(
          ".visualtex-structural-placeholder",
        ).length ?? -1;
        return {
          ready: count === ${testCase.expectedCount - 1},
          count,
          value: field?.value ?? "",
          selection: field?.selection ?? null,
        };
      })()`, `placeholder consumed: ${testCase.name} cycle ${cycle + 1}`);

      await evaluate(`(() => {
        const field = document.querySelector("math-field");
        window.__visualTexPlaceholderRestoreFrames = [];
        field?.addEventListener("input", () => {
          const sample = () => {
            window.__visualTexPlaceholderRestoreFrames.push(
              field.shadowRoot?.querySelectorAll(
                ".visualtex-structural-placeholder",
              ).length ?? -1,
            );
          };
          sample();
          requestAnimationFrame(() => {
            sample();
            requestAnimationFrame(sample);
          });
        }, { once: true });
        return true;
      })()`);
      await key("Backspace", "Backspace", 8);
      const restored = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const placeholders = Array.from(
          field?.shadowRoot?.querySelectorAll(
            ".visualtex-structural-placeholder",
          ) ?? [],
        );
        const selected = placeholders.find(
          (node) =>
            node.classList.contains("ML__placeholder-selected") ||
            node.classList.contains("ML__selected") ||
            Boolean(node.closest(".ML__selected")),
        );
        const caret = selected?.querySelector(
          ".visualtex-structural-placeholder-caret",
        );
        const bounds = selected?.getBoundingClientRect();
        const background = selected ? getComputedStyle(selected).backgroundColor : "";
        const beforeBackground = selected
          ? getComputedStyle(selected, "::before").backgroundColor
          : "";
        const selection = field?.selection ?? null;
        const selectionMatches =
          JSON.stringify(selection) === ${JSON.stringify(JSON.stringify(initial.selection))};
        return {
          ready:
            placeholders.length === ${testCase.expectedCount} &&
            Boolean(selected) &&
            Boolean(caret) &&
            selectionMatches &&
            background === ${JSON.stringify(initial.background)} &&
            beforeBackground === ${JSON.stringify(initial.beforeBackground)} &&
            [background, beforeBackground].some((color) =>
              ["rgb(207, 232, 247)", "rgb(217, 237, 249)"].includes(color)
            ) &&
            (field?.value ?? "")
              .replace(/\\s+/g, "")
              .replace(/\\{([A-Za-z0-9])\\}/g, "$1") ===
              ${JSON.stringify(canonicalLatex(initial.value))} &&
            Math.abs((bounds?.top ?? -99) - ${initial.top}) <= 1 &&
            Math.abs((bounds?.width ?? -99) - ${initial.width}) <= 1 &&
            Math.abs((bounds?.height ?? -99) - ${initial.height}) <= 1,
          count: placeholders.length,
          selected: Boolean(selected),
          caret: Boolean(caret),
          background,
          beforeBackground,
          selection,
          expectedSelection: ${JSON.stringify(initial.selection)},
          expectedRangeLatex: ${JSON.stringify(initial.rangeLatex)},
          expectedInitialValue: ${JSON.stringify(initial.value)},
          typedValue: ${JSON.stringify(typed.value)},
          value: field?.value ?? "",
          selectionMatches,
          top: bounds?.top ?? -1,
          width: bounds?.width ?? -1,
          height: bounds?.height ?? -1,
        };
      })()`, `placeholder restored: ${testCase.name} cycle ${cycle + 1}`);
      await sleep(40);
      const restoreFrameCounts = await evaluate(
        `window.__visualTexPlaceholderRestoreFrames ?? []`,
      );
      if (
        restoreFrameCounts.length < 3 ||
        restoreFrameCounts.slice(1).some(
          (count) => count < testCase.expectedCount,
        )
      ) {
        throw new Error(
          `Placeholder flashed while restoring: ${testCase.name}: ${JSON.stringify(
            restoreFrameCounts,
          )}`,
        );
      }
      cycles.push({ typed, restored, restoreFrameCounts });
    }

    await key("Backspace", "Backspace", 8);
    const removedRestoredPlaceholder = await waitForEvaluation(`(() => {
      const field = document.querySelector("math-field");
      const count = field?.shadowRoot?.querySelectorAll(
        ".visualtex-structural-placeholder",
      ).length ?? -1;
      return {
        ready: count < ${testCase.expectedCount},
        count,
        value: field?.value ?? "",
        selection: field?.selection ?? null,
      };
    })()`, `restored placeholder can be deleted: ${testCase.name}`);
    await sleep(240);
    const stableAfterRemoval = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        count: field?.shadowRoot?.querySelectorAll(
          ".visualtex-structural-placeholder",
        ).length ?? -1,
        value: field?.value ?? "",
      };
    })()`);
    if (
      stableAfterRemoval.count !== removedRestoredPlaceholder.count ||
      stableAfterRemoval.value !== removedRestoredPlaceholder.value
    ) {
      throw new Error(
        `Deleted placeholder reappeared: ${testCase.name}: ${JSON.stringify({
          removedRestoredPlaceholder,
          stableAfterRemoval,
        })}`,
      );
    }
    persistentPlaceholderStates.push({
      name: testCase.name,
      initial,
      cycles,
      removedRestoredPlaceholder,
      stableAfterRemoval,
    });
  }

  const wrapperPlaceholderCases = [
        {
          name: "fraction-numerator",
          source: String.raw`p+\frac{\placeholder{}}{d}+q`,
          command: String.raw`\mathbf`,
        },
        {
          name: "fraction-denominator",
          source: String.raw`p+\frac{n}{\placeholder{}}+q`,
          command: String.raw`\mathcal`,
        },
        {
          name: "integral-upper-limit",
          source: String.raw`p+\int_{l}^{\placeholder{}}f\,dx+q`,
          command: String.raw`\mathfrak`,
        },
        {
          name: "integral-lower-limit",
          source: String.raw`p+\int_{\placeholder{}}^{u}f\,dx+q`,
          command: String.raw`\mathbb`,
        },
        {
          name: "summation-upper-limit",
          source: String.raw`p+\sum_{i=0}^{\placeholder{}}a_{i}+q`,
          command: String.raw`\mathbf`,
        },
        {
          name: "summation-lower-limit",
          source: String.raw`p+\sum_{\placeholder{}}^{n}a_{i}+q`,
          command: String.raw`\mathcal`,
        },
      ];
      const wrapperPlaceholderStates = [];
      for (const testCase of wrapperPlaceholderCases) {
        await evaluate(`(() => {
          const field = document.querySelector("math-field");
          field.setValue(${JSON.stringify(testCase.source)}, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceAll",
            selectionMode: "placeholder",
            silenceNotifications: true,
          });
          field.dispatchEvent(new InputEvent("input", {
            bubbles: true,
            composed: true,
            inputType: "insertText",
          }));
          field.focus();
          field.shadowRoot
            ?.querySelector('[part="keyboard-sink"]')
            ?.focus({ preventScroll: true });
        })()`);
        await sleep(100);
        const caretState = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const host = field?.closest(".mathfield-host");
          const caret = field?.shadowRoot?.querySelector(
            ".visualtex-structural-placeholder-caret",
          );
          if (!field || !host || !caret) return { ready: false };
          const hostBounds = host.getBoundingClientRect();
          const caretBounds = caret.getBoundingClientRect();
          host.dataset.testExpectedWrapperAnchorX = String(
            caretBounds.left - hostBounds.left,
          );
          host.dataset.testExpectedWrapperAnchorY = String(
            caretBounds.top - hostBounds.top + caretBounds.height / 2,
          );
          return {
            ready: caretBounds.height > 0,
            left: caretBounds.left - hostBounds.left,
            centerY:
              caretBounds.top - hostBounds.top + caretBounds.height / 2,
            height: caretBounds.height,
          };
        })()`, `visible structural caret before ${testCase.name}`);
        await typeText(testCase.command);
        await key(" ", "Space", 32);
        const frameState = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const host = field?.closest(".mathfield-host");
          if (!field || !host) return { ready: false };
          const frameCenter = Number.parseFloat(
            host.style.getPropertyValue("--pending-wrapper-left") || "NaN",
          );
          const frameTop = Number.parseFloat(
            host.style.getPropertyValue("--pending-wrapper-top") || "NaN",
          );
          const frameWidth = Number.parseFloat(
            host.style.getPropertyValue("--pending-wrapper-width") || "NaN",
          );
          const frameHeight = Number.parseFloat(
            host.style.getPropertyValue("--pending-wrapper-height") || "NaN",
          );
          const frameLeft = frameCenter - frameWidth / 2;
          const expectedLeft = Number.parseFloat(
            host.dataset.testExpectedWrapperAnchorX || "NaN",
          );
          const expectedTop = Number.parseFloat(
            host.dataset.testExpectedWrapperAnchorY || "NaN",
          );
          const formulaFontSize =
            Number.parseFloat(field.style.fontSize) || 54;
          const minimumFrameHeight = Math.max(
            12,
            formulaFontSize * 0.52,
          );
          const maximumFrameHeight = Math.max(
            minimumFrameHeight,
            formulaFontSize * 1.08,
          );
          const expectedHeight = Math.max(
            minimumFrameHeight,
            Math.min(
              maximumFrameHeight,
              ${caretState.height} + 4,
            ),
          );
          return {
            ready:
              field.dataset.pendingWrapperCommand ===
                ${JSON.stringify(testCase.command)} &&
              host.classList.contains("has-pending-wrapper-placeholder") &&
              Math.abs(frameLeft - expectedLeft) <= 2.5 &&
              Math.abs(frameTop - expectedTop) <= 2.5 &&
              Math.abs(frameHeight - expectedHeight) <= 0.5,
            value: field.value,
            frameLeft,
            frameTop,
            frameWidth,
            frameHeight,
            expectedLeft,
            expectedTop,
            expectedHeight,
          };
        })()`, `wrapper frame at structural caret: ${testCase.name}`);
        wrapperPlaceholderStates.push({
          name: testCase.name,
          caretState,
          frameState,
        });
        await key("Enter", "Enter", 13);
      }

      const heights = Object.fromEntries(
        styleStates.map((state) => [
          state.name,
          state.styles.map((style) => style.height),
        ]),
      );
      const fractionMax = Math.max(...(heights.fraction ?? [0]));
      const superscriptMax = Math.max(...(heights.superscript ?? [99]));
      const subscriptMax = Math.max(...(heights.subscript ?? [99]));
      if (!(superscriptMax < fractionMax && subscriptMax < fractionMax)) {
        throw new Error(
          `Script placeholders did not scale below fraction placeholders: ${JSON.stringify(heights)}`,
        );
      }

      console.log(
        JSON.stringify(
          {
            styleStates,
            emptyState,
            hiddenBlinkState,
            visibleBlinkState,
        typedState,
        restoredState,
        persistentPlaceholderStates,
        wrapperPlaceholderStates,
            heights,
          },
          null,
          2,
        ),
      );
      console.log("Targeted structural placeholder regression passed");
      return;
    }

    if (scenario === "accent-placeholder") {
      const cases = [
        { name: "dot", source: String.raw`a+\dot{\placeholder{}}+b` },
        { name: "hat", source: String.raw`a+\hat{\placeholder{}}+b` },
        { name: "vec", source: String.raw`a+\vec{\placeholder{}}+b` },
        { name: "dddot", source: String.raw`a+\dddot{\placeholder{}}+b` },
        { name: "ddddot", source: String.raw`a+\ddddot{\placeholder{}}+b` },
      ];
      const states = [];
      const screenshotDir =
        process.env.VISUALTEX_ACCENT_SCREENSHOT_DIR?.trim() || "";
      const captureFormulaScreenshot = async (name) => {
        if (!screenshotDir) return;
        const clip = await evaluate(`(() => {
          const field = document.querySelector("math-field");
          const bounds = field.getBoundingClientRect();
          return {
            x: Math.max(0, bounds.left - 12),
            y: Math.max(0, bounds.top - 12),
            width: bounds.width + 24,
            height: bounds.height + 24,
            scale: 1,
          };
        })()`);
        const screenshot = await client.send("Page.captureScreenshot", {
          format: "png",
          fromSurface: true,
          captureBeyondViewport: false,
          clip,
        });
        await writeFile(
          `${screenshotDir}/${name}.png`,
          Buffer.from(screenshot.data, "base64"),
        );
      };
      const readAccentState = () =>
        evaluate(`(() => {
          const field = document.querySelector("math-field");
          const root = field?.shadowRoot;
          const symbol = field?.placeholderSymbol || "▢";
          const placeholder = [...(root?.querySelectorAll(
            ".ML__cmr, .ML__placeholder",
          ) ?? [])].filter((node) =>
            node.classList.contains("visualtex-structural-placeholder") ||
            node.classList.contains("ML__placeholder") ||
            node.textContent?.trim() === symbol
          )[0] ?? null;
          const placeholderState = placeholder ? (() => {
            const node = placeholder;
            const bounds = node.getBoundingClientRect();
            const style = getComputedStyle(node);
            const pseudo = getComputedStyle(node, "::before");
            const accentBody = node.closest(".ML__vlist")?.querySelector(
              ":scope > .ML__center .ML__accent-body",
            );
            const accentBounds = accentBody?.getBoundingClientRect();
            const overlay = node.closest(".ML__vlist")?.querySelector(
              ".visualtex-combining-accent-overlay",
            );
            const overlayBounds = overlay?.getBoundingClientRect();
            const accentAnchor = overlayBounds
              ? overlayBounds.left + overlayBounds.width / 2
              : accentBounds
                ? accentBounds.left + accentBounds.width / 2
                : null;
            return {
              atomId: node.dataset.atomId ?? "",
              classes: node.className,
              parentClasses: node.parentElement?.className ?? "",
              left: bounds.left,
              top: bounds.top,
              centerX: bounds.left + bounds.width / 2,
              width: bounds.width,
              height: bounds.height,
              background: style.backgroundColor,
              borderTopWidth: style.borderTopWidth,
              borderRightWidth: style.borderRightWidth,
              borderBottomWidth: style.borderBottomWidth,
              borderLeftWidth: style.borderLeftWidth,
              color: style.color,
              overflow: style.overflow,
              verticalAlign: style.verticalAlign,
              visualBackground: pseudo.backgroundColor,
              overlayKind: overlay?.dataset.kind ?? "",
              overlayDotCount:
                overlay?.querySelectorAll(
                  ".visualtex-combining-accent-dot",
                ).length ?? 0,
              accentAnchor,
              alignmentDelta:
                accentAnchor === null
                  ? 99
                  : Math.abs(bounds.left + bounds.width / 2 - accentAnchor),
              selected:
                node.classList.contains("ML__selected") ||
                node.classList.contains("ML__placeholder-selected") ||
                Boolean(node.closest(".ML__selected")),
            };
          })() : null;
          return {
            value: field?.value ?? "",
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
            selection: field?.selection ?? null,
            pointerSelecting:
              field?.classList.contains("visualtex-pointer-selecting") ?? false,
            placeholder: placeholderState,
          };
        })()`);

      for (const testCase of cases) {
        await waitForEvaluation(
          `(() => ({ ready: Boolean(document.querySelector("math-field")?.isConnected) }))()`,
          `stable field for accent placeholder case: ${testCase.name}`,
        );
        await evaluate(`(() => {
          const field = document.querySelector("math-field");
          field.setValue(${JSON.stringify(testCase.source)}, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceAll",
            selectionMode: "after",
            silenceNotifications: true,
          });
          field.focus();
          field.position = field.lastOffset;
          field.executeCommand("moveToPreviousPlaceholder");
          field.shadowRoot
            ?.querySelector('[part="keyboard-sink"]')
            ?.focus({ preventScroll: true });
        })()`);
        await sleep(150);
        const initial = await readAccentState();
        await captureFormulaScreenshot(`${testCase.name}-placeholder`);
        await key("ArrowRight", "ArrowRight", 39);
        const afterRight = await readAccentState();
        await key("ArrowLeft", "ArrowLeft", 37);
        const reenteredFromRight = await readAccentState();
        await key("ArrowLeft", "ArrowLeft", 37);
        const afterExitLeft = await readAccentState();
        await key("ArrowRight", "ArrowRight", 39);
        const reenteredFromLeft = await readAccentState();

        const point = {
          x: reenteredFromLeft.placeholder.centerX,
          y:
            reenteredFromLeft.placeholder.top +
            reenteredFromLeft.placeholder.height / 2,
        };
        await client.send("Input.dispatchMouseEvent", {
          type: "mousePressed",
          x: point.x,
          y: point.y,
          button: "left",
          buttons: 1,
          clickCount: 1,
        });
        await sleep(30);
        const held = await readAccentState();
        await client.send("Input.dispatchMouseEvent", {
          type: "mouseReleased",
          x: point.x,
          y: point.y,
          button: "left",
          buttons: 0,
          clickCount: 1,
        });
        await sleep(100);
        const released = await readAccentState();

        const selectedRange = (state) => {
          const [start, end] = state.selection?.ranges?.[0] ?? [-1, -1];
          return Math.abs(end - start) === 1 && state.placeholder?.selected;
        };
        const blueColors = new Set([
          "rgb(217, 237, 249)",
          "rgb(207, 232, 247)",
        ]);
        const stablePlaceholder = (state) =>
          state.placeholder &&
          blueColors.has(state.placeholder.visualBackground) &&
          state.placeholder.color === "rgba(0, 0, 0, 0)" &&
          state.placeholder.borderTopWidth === "0px" &&
          state.placeholder.borderRightWidth === "0px" &&
          state.placeholder.borderBottomWidth === "0px" &&
          state.placeholder.borderLeftWidth === "0px";
        const expectedOverlay = {
          vec: { kind: "vector", dotCount: 0 },
          dddot: { kind: "triple-dot", dotCount: 3 },
          ddddot: { kind: "quadruple-dot", dotCount: 4 },
        }[testCase.name];
        const stableOverlay = (state) =>
          !expectedOverlay ||
          (state.placeholder?.overlayKind === expectedOverlay.kind &&
            state.placeholder?.overlayDotCount ===
              expectedOverlay.dotCount);
        const geometryStable =
          Math.abs(held.placeholder.top - initial.placeholder.top) <= 1 &&
          Math.abs(released.placeholder.top - initial.placeholder.top) <= 1;
        if (
          !selectedRange(initial) ||
          afterRight.position !== 5 ||
          !selectedRange(reenteredFromRight) ||
          afterExitLeft.position !== 2 ||
          !selectedRange(reenteredFromLeft) ||
          initial.placeholder.alignmentDelta > 1 ||
          held.placeholder.alignmentDelta > 1 ||
          !stableOverlay(initial) ||
          !stableOverlay(held) ||
          !stableOverlay(released) ||
          !held.pointerSelecting ||
          !stablePlaceholder(held) ||
          !stablePlaceholder(released) ||
          !geometryStable
        ) {
          throw new Error(
            `Accent placeholder regression failed: ${JSON.stringify({
              name: testCase.name,
              initial,
              afterRight,
              reenteredFromRight,
              afterExitLeft,
              reenteredFromLeft,
              held,
              released,
              geometryStable,
            })}`,
          );
        }
        states.push({
          name: testCase.name,
          alignmentDelta: initial.placeholder.alignmentDelta,
          heldAlignmentDelta: held.placeholder.alignmentDelta,
          topDeltaWhileHeld: held.placeholder.top - initial.placeholder.top,
          topDeltaAfterRelease:
            released.placeholder.top - initial.placeholder.top,
          rightReentrySelection: reenteredFromRight.selection,
          leftReentrySelection: reenteredFromLeft.selection,
          heldVisualBackground: held.placeholder.visualBackground,
        });
      }

      const combiningCharacterCases = [
        {
          name: "vec-character",
          source: String.raw`a+\vec{w}+b`,
          overlayKind: "vector",
          dotCount: 0,
        },
        {
          name: "dddot-character",
          source: String.raw`a+\dddot{w}+b`,
          overlayKind: "triple-dot",
          dotCount: 3,
        },
        {
          name: "ddddot-character",
          source: String.raw`a+\ddddot{w}+b`,
          overlayKind: "quadruple-dot",
          dotCount: 4,
        },
      ];
      for (const testCase of combiningCharacterCases) {
        await evaluate(`(() => {
          const field = document.querySelector("math-field");
          field.setValue(${JSON.stringify(testCase.source)}, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceAll",
            selectionMode: "after",
            silenceNotifications: true,
          });
          field.focus();
        })()`);
        await sleep(150);

        const combiningCharacterState = await evaluate(`(() => {
          const root = document.querySelector("math-field")?.shadowRoot;
          const accent = root?.querySelector(".ML__accent-combining-char");
          const layout = accent?.closest(".ML__vlist");
          const base = [...(layout?.children ?? [])].find(
            (node) => !node.classList.contains("ML__center"),
          );
          const accentBounds = accent?.getBoundingClientRect();
          const baseBounds = base?.getBoundingClientRect();
          const overlay = layout?.querySelector(
            ".visualtex-combining-accent-overlay",
          );
          const overlayBounds = overlay?.getBoundingClientRect();
          const accentVisualCenter =
            overlayBounds
              ? overlayBounds.left + overlayBounds.width / 2
              : -1;
          return {
            ready: Boolean(
              accent &&
                base &&
                accentBounds &&
                baseBounds &&
                overlayBounds,
            ),
            classes: accent?.className ?? "",
            leftOffset: accent ? getComputedStyle(accent).left : "",
            accentOrigin: accentBounds?.left ?? -1,
            accentVisualCenter,
            overlayKind: overlay?.dataset.kind ?? "",
            overlayDotCount:
              overlay?.querySelectorAll(
                ".visualtex-combining-accent-dot",
              ).length ?? 0,
            baseCenter: baseBounds
              ? baseBounds.left + baseBounds.width / 2
              : -1,
            alignmentDelta:
              accentBounds && baseBounds
              ? Math.abs(
                    accentVisualCenter -
                      (baseBounds.left + baseBounds.width / 2),
                  )
                : 99,
          };
        })()`);
        await captureFormulaScreenshot(testCase.name);
        if (
          !combiningCharacterState.ready ||
          !combiningCharacterState.classes.includes(
            "visualtex-combining-accent",
          ) ||
          combiningCharacterState.overlayKind !== testCase.overlayKind ||
          combiningCharacterState.overlayDotCount !== testCase.dotCount ||
          combiningCharacterState.alignmentDelta > 1
        ) {
          throw new Error(
            `Combining accent character alignment regression failed: ${JSON.stringify(
              {
                name: testCase.name,
                state: combiningCharacterState,
              },
            )}`,
          );
        }

        states.push({
          name: testCase.name,
          alignmentDelta: combiningCharacterState.alignmentDelta,
          leftOffset: combiningCharacterState.leftOffset,
        });
      }

    console.log(JSON.stringify(states, null, 2));
      console.log("Targeted accent placeholder regression passed");
      return;
    }

    if (scenario === "caret-probe") {
      const cases = [
        { name: "fraction numerator", source: String.raw`p+\frac{z+n}{d}+q` },
        { name: "square root", source: String.raw`p+\sqrt{z+s}+q` },
      ];
      const states = [];
      for (const testCase of cases) {
        const state = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          if (!field?.isConnected) return { ready: false };
          field.setValue(${JSON.stringify(testCase.source)}, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceAll",
            selectionMode: "after",
            silenceNotifications: true,
          });
          let markerEnd = -1;
          for (let end = 1; end <= field.lastOffset; end += 1) {
            const rangeLatex = field.getValue(end - 1, end, "latex").trim();
            const infoLatex = field.getElementInfo(end)?.latex?.trim() ?? "";
            if (rangeLatex === "z" || infoLatex === "z") {
              markerEnd = end;
              break;
            }
          }
          if (markerEnd < 0) return { ready: false, value: field.value };
          field.focus();
          field.selection = { ranges: [[markerEnd, markerEnd]], direction: "none" };
          field.position = markerEnd;
          field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
          const hostBounds = field.closest(".mathfield-host")?.getBoundingClientRect();
          const markers = [...(field.shadowRoot?.querySelectorAll(
            ".ML__caret, .ML__text-caret, .ML__latex-caret",
          ) ?? [])].map((marker) => {
            const bounds = marker.getBoundingClientRect();
            const chain = [];
            let node = marker;
            while (node && chain.length < 8) {
              chain.push({
                tag: node.tagName,
                className: node.className || "",
                atomId: node.getAttribute?.("data-atom-id") || "",
                text: node.textContent || "",
              });
              node = node.parentElement;
            }
            return {
              left: hostBounds ? bounds.right - hostBounds.left : bounds.right,
              top: hostBounds
                ? bounds.top - hostBounds.top + bounds.height / 2
                : bounds.top + bounds.height / 2,
              width: bounds.width,
              height: bounds.height,
              pseudoVisibility: getComputedStyle(marker, "::after").visibility,
              chain,
            };
          });
          return {
            ready: markers.some((marker) => marker.height > 0),
            name: ${JSON.stringify(testCase.name)},
            value: field.value,
            position: field.position,
            selection: field.selection,
            at: field.getElementInfo(field.position),
            before: field.getElementInfo(Math.max(0, field.position - 1)),
            after: field.getElementInfo(Math.min(field.lastOffset, field.position + 1)),
            markers,
          };
        })()`, `caret probe ${testCase.name}`);
        states.push(state);
      }
      console.log(JSON.stringify(states, null, 2));
      console.log("Targeted caret probe passed");
      return;
    }

    if (scenario === "vertical-structure-probe") {
      const cases = [
        { name: "fraction", latex: String.raw`\frac{N}{D}`, anchor: "N" },
        { name: "integral", latex: String.raw`\int_{L}^{U} f`, anchor: "L" },
        { name: "overset", latex: String.raw`\overset{U}{B}`, anchor: "B" },
        { name: "underset", latex: String.raw`\underset{L}{B}`, anchor: "B" },
        { name: "stackrel", latex: String.raw`\stackrel{U}{B}`, anchor: "B" },
        { name: "stackbin", latex: String.raw`\stackbin{U}{B}`, anchor: "B" },
        {
          name: "overunderset",
          latex: String.raw`\overset{U}{\underset{L}{B}}`,
          anchor: "B",
        },
      ];
      const results = [];
      for (const testCase of cases) {
        await clearField();
        const state = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          if (!field?.isConnected) return { ready: false };
          field.setValue(${JSON.stringify(testCase.latex)}, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceAll",
            selectionMode: "after",
            silenceNotifications: true,
          });
          let anchorOffset = -1;
          for (let offset = 1; offset <= field.lastOffset; offset += 1) {
            const latex = field.getElementInfo(offset)?.latex?.trim() ?? "";
            if (latex === ${JSON.stringify(testCase.anchor)}) {
              anchorOffset = offset;
              break;
            }
          }
          if (anchorOffset < 0) return { ready: false, value: field.value };
          field.focus();
          field.selection = {
            ranges: [[anchorOffset, anchorOffset]],
            direction: "none",
          };
          field.position = anchorOffset;
          field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({
            preventScroll: true,
          });
          const read = () => {
            const marker = [...(field.shadowRoot?.querySelectorAll(
              ".ML__caret, .ML__placeholder-selected, .ML__selected",
            ) ?? [])].find((candidate) => candidate.getBoundingClientRect().height > 0);
            const markerBox = marker?.getBoundingClientRect();
            const chain = [];
            let node = marker;
            while (node && chain.length < 12) {
              const box = node.getBoundingClientRect();
              chain.push({
                tag: node.tagName,
                className: node.className || "",
                atomId: node.getAttribute?.("data-atom-id") || "",
                top: box.top,
                left: box.left,
                width: box.width,
                height: box.height,
              });
              node = node.parentElement;
            }
            return {
              value: field.value,
              position: field.position,
              selection: field.selection,
              marker: markerBox
                ? {
                    top: markerBox.top,
                    left: markerBox.left,
                    width: markerBox.width,
                    height: markerBox.height,
                  }
                : null,
              chain,
            };
          };
          const offsets = [];
          for (let offset = 0; offset <= field.lastOffset; offset += 1) {
            const info = field.getElementInfo(offset);
            const bounds = info?.bounds;
            offsets.push({
              offset,
              latex: info?.latex ?? "",
              depth: info?.depth ?? null,
              bounds: bounds
                ? {
                    top: bounds.top,
                    left: bounds.left,
                    width: bounds.width,
                    height: bounds.height,
                  }
                : null,
            });
          }
          const initial = read();
          const nativeUpChanged = field.executeCommand("moveUp");
          const afterUp = read();
          field.selection = {
            ranges: [[anchorOffset, anchorOffset]],
            direction: "none",
          };
          field.position = anchorOffset;
          const nativeDownChanged = field.executeCommand("moveDown");
          const afterDown = read();
          return {
            ready: true,
            name: ${JSON.stringify(testCase.name)},
            anchorOffset,
            offsets,
            initial,
            nativeUpChanged,
            afterUp,
            nativeDownChanged,
            afterDown,
            contentHtml:
              field.shadowRoot?.querySelector('[part="content"]')?.innerHTML ?? "",
          };
        })()`, `vertical structure probe ${testCase.name}`);
        results.push(state);
      }
      console.log(JSON.stringify(results, null, 2));
      console.log("Vertical structure probe passed");
      return;
    }

    if (scenario === "vertical-structure-navigation") {
      const prepareAt = async (latex, anchor) => {
        await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          if (!field?.isConnected) return { ready: false };
          field.setValue(${JSON.stringify(latex)}, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceAll",
            selectionMode: "after",
            silenceNotifications: true,
          });
          field.dispatchEvent(new InputEvent("input", {
            bubbles: true,
            composed: true,
            inputType: "insertText",
          }));
          return { ready: field.value === ${JSON.stringify(latex)}, value: field.value };
        })()`, `sync vertical structure ${latex}`);
        await sleep(180);
        await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          if (!field?.isConnected) return { ready: false };
          let offset = -1;
          for (let candidate = 1; candidate <= field.lastOffset; candidate += 1) {
            if ((field.getElementInfo(candidate)?.latex ?? "").trim() === ${JSON.stringify(anchor)}) {
              offset = candidate;
              break;
            }
          }
          if (offset < 0) return { ready: false, value: field.value };
          field.focus();
          field.selection = { ranges: [[offset, offset]], direction: "none" };
          field.position = offset;
          field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
          return {
            ready:
              field.position === offset &&
              (field.getElementInfo(field.position)?.latex ?? "").trim() === ${JSON.stringify(anchor)},
            value: field.value,
            position: field.position,
          };
        })()`, `position vertical structure ${latex}`);
        await sleep(80);
      };

      const readPosition = async (expectedLatex, description) =>
        waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const latex = (field?.getElementInfo(field.position)?.latex ?? "").trim();
          return {
            ready: latex === ${JSON.stringify(expectedLatex)},
            value: field?.value ?? "",
            position: field?.position ?? -1,
            latex,
            selection: field?.selection ?? null,
          };
        })()`, description);

      const results = [];
      const runTwoBand = async ({ name, latex, anchor, firstKey, firstExpected, secondKey, secondExpected }) => {
        await prepareAt(latex, anchor);
        await key(firstKey, firstKey, firstKey === "ArrowUp" ? 38 : 40);
        const first = await readPosition(firstExpected, `${name} first vertical move`);
        await key(secondKey, secondKey, secondKey === "ArrowUp" ? 38 : 40);
        const second = await readPosition(secondExpected, `${name} return vertical move`);
        results.push({ name, first, second });
      };

      await runTwoBand({
        name: "overset",
        latex: String.raw`\overset{U}{B}`,
        anchor: "B",
        firstKey: "ArrowUp",
        firstExpected: "U",
        secondKey: "ArrowDown",
        secondExpected: "B",
      });
      await runTwoBand({
        name: "underset",
        latex: String.raw`\underset{L}{B}`,
        anchor: "B",
        firstKey: "ArrowDown",
        firstExpected: "L",
        secondKey: "ArrowUp",
        secondExpected: "B",
      });
      await runTwoBand({
        name: "stackrel",
        latex: String.raw`\stackrel{U}{B}`,
        anchor: "B",
        firstKey: "ArrowUp",
        firstExpected: "U",
        secondKey: "ArrowDown",
        secondExpected: "B",
      });
      await runTwoBand({
        name: "stackbin",
        latex: String.raw`\stackbin{U}{B}`,
        anchor: "B",
        firstKey: "ArrowUp",
        firstExpected: "U",
        secondKey: "ArrowDown",
        secondExpected: "B",
      });

      await prepareAt(String.raw`\overset{U}{\underset{L}{B}}`, "B");
      await key("ArrowUp", "ArrowUp", 38);
      const tripleUpper = await readPosition("U", "overunderset base to upper");
      await key("ArrowDown", "ArrowDown", 40);
      const tripleBaseFromUpper = await readPosition("B", "overunderset upper to base");
      await key("ArrowDown", "ArrowDown", 40);
      const tripleLower = await readPosition("L", "overunderset base to lower");
      await key("ArrowUp", "ArrowUp", 38);
      const tripleBaseFromLower = await readPosition("B", "overunderset lower to base");
      results.push({
        name: "overunderset",
        tripleUpper,
        tripleBaseFromUpper,
        tripleLower,
        tripleBaseFromLower,
      });

      await prepareAt(String.raw`\xrightarrow[L]{U}`, "U");
      await key("ArrowDown", "ArrowDown", 40);
      const xArrowLower = await readPosition("L", "xrightarrow upper to lower");
      await key("ArrowUp", "ArrowUp", 38);
      const xArrowUpper = await readPosition("U", "xrightarrow lower to upper");
      results.push({ name: "xrightarrow", xArrowLower, xArrowUpper });

      const preparePlaceholderBand = async (latex, bandIndex) => {
        await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          if (!field?.isConnected) return { ready: false };
          field.setValue(${JSON.stringify(latex)}, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceAll",
            selectionMode: "after",
            silenceNotifications: true,
          });
          field.dispatchEvent(new InputEvent("input", {
            bubbles: true,
            composed: true,
            inputType: "insertText",
          }));
          return { ready: field.value === ${JSON.stringify(latex)}, value: field.value };
        })()`, `sync placeholder structure ${latex}`);
        await sleep(180);
        await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          if (!field?.isConnected) return { ready: false };
          const placeholders = [];
          for (let offset = 1; offset <= field.lastOffset; offset += 1) {
            const info = field.getElementInfo(offset);
            const bounds = info?.bounds;
            if ((info?.latex ?? "").trim() !== "\\\\placeholder{}" || !bounds) continue;
            placeholders.push({
              offset,
              y: bounds.top + bounds.height / 2,
            });
          }
          placeholders.sort((left, right) => left.y - right.y);
          const target = placeholders[${bandIndex}];
          if (!target) return { ready: false, placeholders, value: field.value };
          field.focus();
          field.selection = {
            ranges: [[target.offset - 1, target.offset]],
            direction: "none",
          };
          field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
          const selectedEnd = field.selection.ranges.at(-1)?.[1] ?? -1;
          return {
            ready: selectedEnd === target.offset,
            placeholders,
            selectedEnd,
            value: field.value,
          };
        })()`, `position placeholder band ${latex}`);
        await sleep(80);
      };

      const readPlaceholderBand = async (expectedBandIndex, description) =>
        waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const placeholders = [];
          for (let offset = 1; offset <= field.lastOffset; offset += 1) {
            const info = field.getElementInfo(offset);
            const bounds = info?.bounds;
            if ((info?.latex ?? "").trim() !== "\\\\placeholder{}" || !bounds) continue;
            placeholders.push({ offset, y: bounds.top + bounds.height / 2 });
          }
          placeholders.sort((left, right) => left.y - right.y);
          const selectedEnd = field?.selection?.ranges?.at(-1)?.[1] ?? -1;
          return {
            ready: selectedEnd === placeholders[${expectedBandIndex}]?.offset,
            selectedEnd,
            placeholders,
            value: field?.value ?? "",
          };
        })()`, description);

      await preparePlaceholderBand(
        String.raw`\overset{\placeholder{}}{\placeholder{}}`,
        0,
      );
      await key("ArrowDown", "ArrowDown", 40);
      const oversetPlaceholderBase = await readPlaceholderBand(
        1,
        "overset upper placeholder to base placeholder",
      );
      await key("ArrowUp", "ArrowUp", 38);
      const oversetPlaceholderUpper = await readPlaceholderBand(
        0,
        "overset base placeholder to upper placeholder",
      );
      results.push({
        name: "overset-placeholders",
        oversetPlaceholderBase,
        oversetPlaceholderUpper,
      });

      await preparePlaceholderBand(
        String.raw`\overset{\placeholder{}}{\underset{\placeholder{}}{\placeholder{}}}`,
        0,
      );
      await key("ArrowDown", "ArrowDown", 40);
      const triplePlaceholderBase = await readPlaceholderBand(
        1,
        "triple upper placeholder to base placeholder",
      );
      await key("ArrowDown", "ArrowDown", 40);
      const triplePlaceholderLower = await readPlaceholderBand(
        2,
        "triple base placeholder to lower placeholder",
      );
      await key("ArrowUp", "ArrowUp", 38);
      const triplePlaceholderBaseAgain = await readPlaceholderBand(
        1,
        "triple lower placeholder to base placeholder",
      );
      results.push({
        name: "overunderset-placeholders",
        triplePlaceholderBase,
        triplePlaceholderLower,
        triplePlaceholderBaseAgain,
      });

      console.log(JSON.stringify(results, null, 2));
      console.log("Vertical structure navigation regression passed");
      return;
    }

    if (scenario === "native-structure-audit") {
      const prefixes = ["\\\\over", "\\\\under", "\\\\stack"];
      const results = [];
      for (const prefix of prefixes) {
        await clearField();
        await typeText(prefix);
        const state = await waitForEvaluation(`(() => {
          const stable = document.getElementById(
            "visualtex-native-input-suggestion-popover",
          );
          const items = [...(stable?.querySelectorAll("li[data-command]") ?? [])]
            .map((item) => ({
              command: item.dataset.command ?? "",
              latex: item.querySelector(".ML__popover__latex")?.textContent?.trim() ?? "",
              previewText: item.querySelector(".ML__popover__command")?.textContent?.trim() ?? "",
              hasVisualPreview: item.classList.contains(
                "has-visualtex-command-preview",
              ),
              previewMarkup:
                item.querySelector(".ML__popover__command")?.innerHTML ?? "",
            }));
          return {
            ready: Boolean(stable?.classList.contains("is-visible")) && items.length > 0,
            items,
          };
        })()`, `${prefix} native structure suggestions`);
        results.push({ prefix, items: state.items });
      }
      await clearField();
      const modelAudit = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const templates = [
          "\\\\overset{\\\\placeholder{}}{\\\\placeholder{}}",
          "\\\\underset{\\\\placeholder{}}{\\\\placeholder{}}",
          "\\\\overunderset{\\\\placeholder{}}{\\\\placeholder{}}{\\\\placeholder{}}",
          "\\\\stackrel{\\\\placeholder{}}{\\\\placeholder{}}",
        ];
        return templates.map((template) => {
          field.setValue("", { mode: "math", format: "latex" });
          field.insert(template, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceSelection",
            selectionMode: "placeholder",
          });
          const initialSelection = field.selection;
          field.insert("ab", {
            mode: "math",
            format: "latex",
            insertionMode: "replaceSelection",
            selectionMode: "after",
          });
          const valueAfterTyping = field.value;
          field.setValue(template, { mode: "math", format: "latex" });
          const offsets = [];
          for (let offset = 0; offset <= field.lastOffset; offset += 1) {
            const info = field.getElementInfo(offset);
            offsets.push({
              offset,
              latex: info?.latex ?? "",
              depth: info?.depth ?? null,
            });
          }
          return {
            template,
            value: field.value,
            initialSelection,
            valueAfterTyping,
            selection: field.selection,
            position: field.position,
            lastOffset: field.lastOffset,
            offsets,
          };
        });
      })()`);
      console.log(JSON.stringify({ results, modelAudit }, null, 2));
      console.log("Native structure suggestion audit passed");
      return;
    }

    if (
      scenario === "native-structure-input-over" ||
      scenario === "native-structure-input-under" ||
      scenario === "native-structure-input-multi" ||
      scenario === "native-structure-input-core"
    ) {
      const cases =
        scenario === "native-structure-input-over"
          ? [
              ["overarc", 1, "\\overarc{a}"],
              ["overline", 1, "\\overline{a}"],
              ["overbrace", 1, "\\overbrace{a}"],
              ["overgroup", 1, "\\overgroup{a}"],
              ["overparen", 1, "\\overparen{a}"],
              ["overleftarrow", 1, "\\overleftarrow{a}"],
              ["overrightarrow", 1, "\\overrightarrow{a}"],
              ["overleftrightarrow", 1, "\\overleftrightarrow{a}"],
              ["overleftharpoon", 1, "\\overleftharpoon{a}"],
              ["overrightharpoon", 1, "\\overrightharpoon{a}"],
              ["overlinesegment", 1, "\\overlinesegment{a}"],
            ]
          : scenario === "native-structure-input-under"
            ? [
                ["underarc", 1, "\\underarc{a}"],
                ["underline", 1, "\\underline{a}"],
                ["underbrace", 1, "\\underbrace{a}"],
                ["undergroup", 1, "\\undergroup{a}"],
                ["underparen", 1, "\\underparen{a}"],
                ["underleftarrow", 1, "\\underleftarrow{a}"],
                ["underrightarrow", 1, "\\underrightarrow{a}"],
                ["underleftrightarrow", 1, "\\underleftrightarrow{a}"],
                ["underlinesegment", 1, "\\underlinesegment{a}"],
              ]
            : scenario === "native-structure-input-core"
              ? [
                  ["sqrt", 1, "\\sqrt{a}"],
                  ["frac", 2, "\\frac{a}{"],
                  ["dfrac", 2, "\\dfrac{a}{"],
                  ["tfrac", 2, "\\tfrac{a}{"],
                  ["binom", 2, "\\binom{a}{"],
                ]
              : [
                ["overset", 2, "\\overset{a}{"],
                ["underset", 2, "\\underset{a}{"],
                ["stackrel", 2, "\\stackrel{a}{"],
                ["stackbin", 2, "\\stackbin{a}{"],
                ["overunderset", 3, "\\overset{a}{\\underset{"],
              ];

      const states = [];
      for (const [command, placeholderCount, expectedFragment] of cases) {
        await clearField();
        await typeText(`\\${command}`);
        await key(" ", "Space", 32);
        const confirmedCommand =
          command === "overunderset" ? "\\overset" : `\\${command}`;
        const pending = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const value = field?.value ?? "";
          return {
            ready:
              Boolean(field?.isConnected) &&
              value.includes(${JSON.stringify(confirmedCommand)}) &&
              !field.shadowRoot?.querySelector(".ML__raw-latex"),
            value,
            selection: field?.selection ?? null,
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, `confirmed \\${command} structure`);
        const actualPlaceholderCount =
          pending.value.match(/\\placeholder\{\}/g)?.length ?? 0;
        if (actualPlaceholderCount !== placeholderCount) {
          throw new Error(
            `\\${command} created ${actualPlaceholderCount} placeholders instead of ${placeholderCount}: ${pending.value}`,
          );
        }

        await typeText("a");
        const typed = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const value = field?.value ?? "";
          return {
            ready: Boolean(field?.isConnected) && value.includes("a"),
            value,
            selection: field?.selection ?? null,
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, `typed content inside \\${command}`);
        if (!typed.value.includes(expectedFragment)) {
          throw new Error(
            `\\${command} did not place the first character in its first LaTeX argument: ${typed.value}`,
          );
        }
        if (typed.value.includes(`\\${command}{}a`)) {
          throw new Error(`\\${command} left an empty argument before typed content`);
        }
        let finalTypedValue = typed.value;
        if (command === "sqrt") {
          await typeText("b");
          const continued = await waitForEvaluation(`(() => {
            const field = document.querySelector("math-field");
            return {
              ready: field?.value === "\\\\sqrt{ab}",
              value: field?.value ?? "",
            };
          })()`, "continued typing inside the radical");
          finalTypedValue = continued.value;
        }
        states.push({ command, pending: pending.value, typed: finalTypedValue });
      }

      const protectedStructureStates = {};
      if (scenario === "native-structure-input-core") {
        await evaluate(`(() => {
          const storageKey = "visualtex-editor";
          const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
          persisted.state = {
            ...(persisted.state || {}),
            inputBehavior: {
              ...(persisted.state?.inputBehavior || {}),
              autoExitSuperscript: true,
              autoExitSubscript: true,
            },
          };
          localStorage.setItem(storageKey, JSON.stringify(persisted));
          location.reload();
        })()`);
        await sleep(650);
        await waitForEvaluation(
          `(() => ({ ready: Boolean(document.querySelector("math-field")) }))()`,
          "formula field after protected-structure setting reload",
        );

        const setPlaceholderValue = async (latex, description) => {
          await waitForEvaluation(`(() => {
            const field = document.querySelector("math-field");
            if (!field?.isConnected) return { ready: false };
            field.setValue(${JSON.stringify(latex)}, {
              mode: "math",
              format: "latex",
              insertionMode: "replaceAll",
              selectionMode: "placeholder",
              silenceNotifications: true,
            });
            field.focus();
            field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
            return {
              ready: !field.selectionIsCollapsed,
              value: field.value,
              selection: field.selection,
            };
          })()`, description);
          await sleep(100);
        };

        const readStructuredCaret = async () =>
          evaluate(`(() => {
            const field = document.querySelector("math-field");
            const markers = [...(field?.shadowRoot?.querySelectorAll(
              ".ML__placeholder-selected, .ML__selected, .ML__caret",
            ) ?? [])];
            const marker = markers.find((candidate) =>
              candidate.closest(".ML__op-group, .ML__mfrac"),
            );
            const operator = marker?.closest(".ML__op-group");
            const fraction = marker?.closest(".ML__mfrac");
            const container = operator || fraction;
            const markerBox = (marker?.parentElement || marker)?.getBoundingClientRect();
            const containerBox = container?.getBoundingClientRect();
            const region =
              markerBox && containerBox
                ? markerBox.top + markerBox.height / 2 <
                  containerBox.top + containerBox.height / 2
                  ? "upper"
                  : "lower"
                : null;
            return {
              value: field?.value ?? "",
              position: field?.position ?? -1,
              lastOffset: field?.lastOffset ?? -1,
              inOperator: Boolean(operator),
              inFraction: Boolean(fraction),
              region,
            };
          })()`);

        await setPlaceholderValue(
          "\\int_{\\placeholder{}}^{n} f",
          "integral lower-limit placeholder",
        );
        await typeText("ab");
        const operatorContinuous = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const normalized = (field?.value ?? "").replaceAll(" ", "");
          const marker = [...(field?.shadowRoot?.querySelectorAll(
            ".ML__placeholder-selected, .ML__selected, .ML__caret",
          ) ?? [])].find((candidate) => candidate.closest(".ML__op-group"));
          return {
            ready: normalized.includes("\\\\int_{ab}^{n}") && Boolean(marker),
            value: field?.value ?? "",
          };
        })()`, "continuous typing inside an integral limit");
        await key("ArrowUp", "ArrowUp", 38);
        const operatorUpper = await readStructuredCaret();
        if (!operatorUpper.inOperator || operatorUpper.region !== "upper") {
          throw new Error(
            `ArrowUp left the integral limits: ${JSON.stringify(operatorUpper)}`,
          );
        }
        await key("ArrowDown", "ArrowDown", 40);
        const operatorLower = await readStructuredCaret();
        if (!operatorLower.inOperator || operatorLower.region !== "lower") {
          throw new Error(
            `ArrowDown left the integral limits: ${JSON.stringify(operatorLower)}`,
          );
        }

        await setPlaceholderValue(
          "\\sqrt{x^{\\placeholder{}}+q}",
          "superscript placeholder inside a radical",
        );
        await typeText("ab");
        const radicalNestedScript = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const normalized = (field?.value ?? "").replaceAll(" ", "");
          return {
            ready: normalized.includes("\\\\sqrt{x^{a}b+q}"),
            value: field?.value ?? "",
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, "ordinary superscript auto-exits inside a radical");

        await setPlaceholderValue(
          "\\sqrt{x_{\\placeholder{}}+q}",
          "subscript placeholder inside a radical",
        );
        await typeText("ab");
        const radicalNestedSubscript = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const normalized = (field?.value ?? "").replaceAll(" ", "");
          return {
            ready: normalized.includes("\\\\sqrt{x_{a}b+q}"),
            value: field?.value ?? "",
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, "ordinary subscript auto-exits inside a radical");

        await setPlaceholderValue(
          "\\int_0^1 x^{\\placeholder{}}\\,\\mathrm{d}x",
          "superscript placeholder inside an integral body",
        );
        await typeText("ab");
        const integralBodyNestedScript = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const normalized = (field?.value ?? "").replaceAll(" ", "");
          return {
            ready: normalized.includes("x^{a}b"),
            value: field?.value ?? "",
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, "ordinary superscript auto-exits inside an integral body");

        await setPlaceholderValue(
          "\\int_{x^{\\placeholder{}}}^{n} f",
          "nested superscript inside an integral limit",
        );
        await typeText("ab");
        const integralLimitNestedScript = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const normalized = (field?.value ?? "").replaceAll(" ", "");
          const marker = [...(field?.shadowRoot?.querySelectorAll(
            ".ML__placeholder-selected, .ML__selected, .ML__caret",
          ) ?? [])].find((candidate) => candidate.closest(".ML__op-group"));
          return {
            ready: normalized.includes("\\\\int_{x^{a}b}^{n}") && Boolean(marker),
            value: field?.value ?? "",
          };
        })()`, "nested superscript auto-exits but remains inside an integral limit");

        await setPlaceholderValue(
          "\\frac{x^{\\placeholder{}}+y}{z}+q",
          "superscript placeholder inside a fraction",
        );
        await typeText("ab");
        const fractionScript = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const normalized = (field?.value ?? "").replaceAll(" ", "");
          const marker = [...(field?.shadowRoot?.querySelectorAll(
            ".ML__placeholder-selected, .ML__selected, .ML__caret",
          ) ?? [])].find((candidate) => candidate.closest(".ML__mfrac"));
          return {
            ready: normalized.includes("x^{a}b") && Boolean(marker),
            value: field?.value ?? "",
          };
        })()`, "ordinary superscript auto-exits inside a fraction");

        await setPlaceholderValue(
          "x^{\\placeholder{}}",
          "one-time superscript auto-exit placeholder",
        );
        await typeText("a");
        const firstScriptExit = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          return {
            ready:
              field?.value === "x^{a}" &&
              field.position === field.lastOffset,
            value: field?.value ?? "",
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, "first character auto-exits its superscript");
        await key("ArrowLeft", "ArrowLeft", 37);
        const manualScriptReentry = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const marker = [...(field?.shadowRoot?.querySelectorAll(
            ".ML__placeholder-selected, .ML__selected, .ML__caret",
          ) ?? [])].find((candidate) => candidate.closest(".ML__msubsup"));
          return {
            ready: Boolean(marker),
            value: field?.value ?? "",
            position: field?.position ?? -1,
          };
        })()`, "manual re-entry into the consumed superscript");
        await typeText("bc");
        const retainedAfterManualReentry = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const marker = [...(field?.shadowRoot?.querySelectorAll(
            ".ML__placeholder-selected, .ML__selected, .ML__caret",
          ) ?? [])].find((candidate) => candidate.closest(".ML__msubsup"));
          return {
            ready:
              field?.value === "x^{abc}" &&
              Boolean(marker) &&
              field.position !== field.lastOffset,
            value: field?.value ?? "",
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, "manual re-entry disables repeated auto-exit for that superscript");

        await setPlaceholderValue(
          "x^{abc}+y^{\\placeholder{}}",
          "a separate superscript after consuming the first one",
        );
        await typeText("d");
        const separateScriptStillAutoExits = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          return {
            ready:
              field?.value === "x^{abc}+y^{d}" &&
              field.position === field.lastOffset,
            value: field?.value ?? "",
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, "a different superscript still performs its first auto-exit");

        await setPlaceholderValue(
          "x_{\\placeholder{}}",
          "one-time subscript auto-exit placeholder",
        );
        await typeText("a");
        const firstSubscriptExit = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          return {
            ready:
              field?.value === "x_{a}" &&
              field.position === field.lastOffset,
            value: field?.value ?? "",
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, "first character auto-exits its subscript");
        await key("ArrowLeft", "ArrowLeft", 37);
        await typeText("bc");
        const retainedSubscriptAfterManualReentry = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const marker = [...(field?.shadowRoot?.querySelectorAll(
            ".ML__placeholder-selected, .ML__selected, .ML__caret",
          ) ?? [])].find((candidate) => candidate.closest(".ML__msubsup"));
          return {
            ready:
              field?.value === "x_{abc}" &&
              Boolean(marker) &&
              field.position !== field.lastOffset,
            value: field?.value ?? "",
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, "manual re-entry disables repeated auto-exit for that subscript");

        await setPlaceholderValue(
          "\\frac{n}{\\placeholder{}}+q",
          "fraction denominator placeholder",
        );
        await key("ArrowUp", "ArrowUp", 38);
        const fractionUpper = await readStructuredCaret();
        if (!fractionUpper.inFraction || fractionUpper.region !== "upper") {
          throw new Error(
            `ArrowUp left the fraction: ${JSON.stringify(fractionUpper)}`,
          );
        }
        await key("ArrowDown", "ArrowDown", 40);
        const fractionLower = await readStructuredCaret();
        if (!fractionLower.inFraction || fractionLower.region !== "lower") {
          throw new Error(
            `ArrowDown left the fraction: ${JSON.stringify(fractionLower)}`,
          );
        }

        Object.assign(protectedStructureStates, {
          operatorContinuous,
          operatorUpper,
          operatorLower,
          radicalNestedScript,
          radicalNestedSubscript,
          integralBodyNestedScript,
          integralLimitNestedScript,
          fractionScript,
          firstScriptExit,
          manualScriptReentry,
          retainedAfterManualReentry,
          separateScriptStillAutoExits,
          firstSubscriptExit,
          retainedSubscriptAfterManualReentry,
          fractionUpper,
          fractionLower,
        });
      }

      console.log(JSON.stringify({ states, protectedStructureStates }, null, 2));
      console.log(`${scenario} passed`);
      return;
    }

    if (scenario === "native-input-popover") {
      await clearField();
      await typeText("\\f");
      const initial = await waitForEvaluation(`(() => {
        const stable = document.getElementById("visualtex-native-input-suggestion-popover");
        const source = document.getElementById("mathlive-suggestion-popover");
        const bounds = stable?.getBoundingClientRect();
        const commands = [...(stable?.querySelectorAll("li[data-command]") ?? [])]
          .map((item) => item.dataset.command ?? "");
        return {
          ready:
            Boolean(stable?.classList.contains("is-visible")) &&
            commands.length >= 2 &&
            source?.dataset.visualtexInputPopoverSource === "true" &&
            !document.querySelector(".suggestion-popup"),
          commands,
          selected: stable?.querySelector("li.ML__popover__current")?.dataset.command ?? "",
          bounds: bounds
            ? { left: bounds.left, top: bounds.top, width: bounds.width, height: bounds.height }
            : null,
          sourceOpacity: source ? getComputedStyle(source).opacity : "",
          customCandidateVisible: Boolean(document.querySelector(".suggestion-popup")),
        };
      })()`, "stable native input-selection popover for \\f");

      await evaluate(`(() => {
        const node = document.getElementById("visualtex-native-input-suggestion-popover");
        const monitor = {
          node,
          removed: 0,
          hiddenTransitions: 0,
          ariaHiddenTransitions: 0,
          outerChildMutations: 0,
        };
        const observer = new MutationObserver((records) => {
          for (const record of records) {
            if (
              record.type === "childList" &&
              [...record.removedNodes].some(
                (removed) => removed === node || removed.contains?.(node),
              )
            ) {
              monitor.removed += 1;
            }
            if (record.target === node && record.type === "childList") {
              monitor.outerChildMutations += 1;
            }
            if (record.target === node && record.attributeName === "class") {
              if (!node.classList.contains("is-visible")) {
                monitor.hiddenTransitions += 1;
              }
            }
            if (record.target === node && record.attributeName === "aria-hidden") {
              if (node.getAttribute("aria-hidden") !== "false") {
                monitor.ariaHiddenTransitions += 1;
              }
            }
          }
        });
        observer.observe(document.body, {
          childList: true,
          subtree: true,
          attributes: true,
          attributeFilter: ["class", "aria-hidden"],
        });
        window.__visualtexNativeInputMonitor = monitor;
        window.__visualtexNativeInputObserver = observer;
      })()`);

      await key("ArrowDown", "ArrowDown", 40);
      const arrowState = await waitForEvaluation(`(() => {
        const monitor = window.__visualtexNativeInputMonitor;
        const stable = document.getElementById("visualtex-native-input-suggestion-popover");
        const bounds = stable?.getBoundingClientRect();
        const selected = stable?.querySelector("li.ML__popover__current")?.dataset.command ?? "";
        return {
          ready:
            stable === monitor?.node &&
            stable?.classList.contains("is-visible") &&
            selected &&
            selected !== ${JSON.stringify(initial.selected)} &&
            monitor.removed === 0 &&
            monitor.hiddenTransitions === 0 &&
            monitor.ariaHiddenTransitions === 0 &&
            Math.abs((bounds?.left ?? 0) - ${initial.bounds.left}) <= 1 &&
            Math.abs((bounds?.top ?? 0) - ${initial.bounds.top}) <= 1 &&
            Math.abs((bounds?.width ?? 0) - ${initial.bounds.width}) <= 1 &&
            Math.abs((bounds?.height ?? 0) - ${initial.bounds.height}) <= 1,
          sameNode: stable === monitor?.node,
          selected,
          bounds: bounds
            ? { left: bounds.left, top: bounds.top, width: bounds.width, height: bounds.height }
            : null,
          removed: monitor?.removed ?? -1,
          hiddenTransitions: monitor?.hiddenTransitions ?? -1,
          ariaHiddenTransitions: monitor?.ariaHiddenTransitions ?? -1,
          customCandidateVisible: Boolean(document.querySelector(".suggestion-popup")),
        };
      })()`, "arrow key moves only the native input-selection highlight");

      await typeText("r");
      const refinedState = await waitForEvaluation(`(() => {
        const monitor = window.__visualtexNativeInputMonitor;
        const stable = document.getElementById("visualtex-native-input-suggestion-popover");
        const commands = [...(stable?.querySelectorAll("li[data-command]") ?? [])]
          .map((item) => item.dataset.command ?? "");
        return {
          ready:
            stable === monitor?.node &&
            stable?.classList.contains("is-visible") &&
            commands.some((command) => command === "\\\\frac") &&
            commands.every((command) => command.startsWith("\\\\fr")) &&
            monitor.removed === 0 &&
            monitor.hiddenTransitions === 0 &&
            monitor.ariaHiddenTransitions === 0 &&
            !document.querySelector(".suggestion-popup"),
          sameNode: stable === monitor?.node,
          commands,
          selected: stable?.querySelector("li.ML__popover__current")?.dataset.command ?? "",
          removed: monitor?.removed ?? -1,
          hiddenTransitions: monitor?.hiddenTransitions ?? -1,
          ariaHiddenTransitions: monitor?.ariaHiddenTransitions ?? -1,
          customCandidateVisible: Boolean(document.querySelector(".suggestion-popup")),
        };
      })()`, "\\f to \\fr updates inside one persistent input-selection frame");

      await key("Backspace", "Backspace", 8);
      const restoredState = await waitForEvaluation(`(() => {
        const monitor = window.__visualtexNativeInputMonitor;
        const stable = document.getElementById("visualtex-native-input-suggestion-popover");
        const commands = [...(stable?.querySelectorAll("li[data-command]") ?? [])]
          .map((item) => item.dataset.command ?? "");
        return {
          ready:
            stable === monitor?.node &&
            stable?.classList.contains("is-visible") &&
            commands.length >= ${initial.commands.length} &&
            monitor.removed === 0 &&
            monitor.hiddenTransitions === 0 &&
            monitor.ariaHiddenTransitions === 0,
          sameNode: stable === monitor?.node,
          commands,
          removed: monitor?.removed ?? -1,
          hiddenTransitions: monitor?.hiddenTransitions ?? -1,
          ariaHiddenTransitions: monitor?.ariaHiddenTransitions ?? -1,
        };
      })()`, "Backspace restores \\f suggestions without remounting the input-selection frame");

      await evaluate(`window.__visualtexNativeInputObserver?.disconnect()`);

      const inspectStructuredPreview = async (
        query,
        command,
        expectedSizeOrder,
      ) => {
        await clearField();
        await typeText(query);
        return waitForEvaluation(`(() => {
          const stable = document.getElementById(
            "visualtex-native-input-suggestion-popover",
          );
          const item = [...(stable?.querySelectorAll("li[data-command]") ?? [])]
            .find((candidate) => candidate.dataset.command === ${JSON.stringify(command)});
          const layers = [...(item?.querySelectorAll(
            ".ML__popover__command .ML__center",
          ) ?? [])]
            .map((node) => {
              const content = node.lastElementChild ?? node;
              const box = content.getBoundingClientRect();
              return {
                top: box.top,
                width: box.width,
                height: box.height,
                text: node.textContent?.trim() ?? "",
              };
            })
            .sort((left, right) => left.top - right.top);
          const top = layers[0];
          const bottom = layers[1];
          const expectedSizeOrder = ${JSON.stringify(expectedSizeOrder)};
          return {
            ready:
              Boolean(stable?.classList.contains("is-visible")) &&
              Boolean(item?.classList.contains("has-visualtex-command-preview")) &&
              layers.length === 2 &&
              (expectedSizeOrder === "small-large"
                ? top.height < bottom.height
                : top.height > bottom.height),
            command: item?.dataset.command ?? "",
            layers,
          };
        })()`, `${command} structured native preview`);
      };

      const oversetPreview = await inspectStructuredPreview(
        "\\overs",
        "\\overset",
        "small-large",
      );
      const undersetPreview = await inspectStructuredPreview(
        "\\unders",
        "\\underset",
        "large-small",
      );

      console.log(JSON.stringify({
        initial,
        arrowState,
        refinedState,
        restoredState,
        oversetPreview,
        undersetPreview,
      }, null, 2));
      console.log("Targeted native input-selection popover regression passed");
      return;
    }

    if (scenario === "export") {
      const exportLatexLines = [
        String.raw`\begin{pmatrix}a&b\\c&d\end{pmatrix}`,
        String.raw`\frac{a}{b}+x^2`,
        String.raw`\int_0^1 x^2\,\mathrm{d}x`,
        String.raw`\symbfit{J}+\bm{\alpha}+\abs{x}+\dv{f}{x}`,
      ];
      await evaluate(`(() => {
        const storageKey = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
        const lines = ${JSON.stringify(exportLatexLines)}.map((latex) => ({
          id: crypto.randomUUID(),
          latex,
        }));
        persisted.state = {
          ...(persisted.state || {}),
          title: "Export Test",
          lines,
          activeLineId: lines[0].id,
        };
        localStorage.setItem(storageKey, JSON.stringify(persisted));
        location.reload();
      })()`);
      await waitForEvaluation(`(() => ({
        ready:
          document.querySelector(".document-title-area input")?.value === "Export Test" &&
          document.querySelectorAll("math-field").length === 4 &&
          [...document.querySelectorAll("math-field")].some((field) => field.value?.includes("\\\\frac")),
        title: document.querySelector(".document-title-area input")?.value ?? "",
        values: [...document.querySelectorAll("math-field")].map((field) => field.value ?? ""),
      }))()`, "multi-row formula document prepared for export");

      await evaluate(`(() => {
        window.__visualtexCapturedExports = [];
        URL.revokeObjectURL = () => {};
        HTMLAnchorElement.prototype.click = function captureVisualTeXExport() {
          window.__visualtexCapturedExports.push({
            filename: this.download,
            href: this.href,
          });
        };
      })()`);

      const clickExportOption = async (label, expectedCount) => {
        await evaluate(`document.querySelector(".workspace-export-trigger")?.click()`);
        await waitForEvaluation(`(() => ({
          ready: Boolean(document.querySelector(".export-dialog")),
        }))()`, `export dialog opened for ${label}`);
        await evaluate(`(() => {
          const button = [...document.querySelectorAll(".export-format-option")]
            .find((candidate) => candidate.querySelector("strong")?.textContent?.trim() === ${JSON.stringify(label)});
          button?.click();
        })()`);
        await waitForEvaluation(`(() => {
          const button = [...document.querySelectorAll(".export-format-option")]
            .find((candidate) => candidate.querySelector("strong")?.textContent?.trim() === ${JSON.stringify(label)});
          return { ready: button?.getAttribute("aria-checked") === "true" };
        })()`, `${label} export format selected`);
        await evaluate(`document.querySelector(".export-confirm-button")?.click()`);
        await waitForEvaluation(`(() => ({
          ready:
            (window.__visualtexCapturedExports?.length ?? 0) >= ${expectedCount} &&
            !document.querySelector(".export-dialog"),
          count: window.__visualtexCapturedExports?.length ?? 0,
        }))()`, `${label} export captured`);
      };

      await clickExportOption("Markdown", 1);
      const markdownState = await waitForEvaluation(`(async () => {
        const item = window.__visualtexCapturedExports?.[0];
        if (!item) return { ready: false };
        const text = await fetch(item.href).then((response) => response.text());
        return {
          ready:
            item.filename.endsWith(".md") &&
            text.includes("Export Test") &&
            text.includes("\\\\begin{pmatrix}") &&
            text.includes("\\\\frac{a}{b}+x^2") &&
            text.includes("\\\\int_0^1") &&
            text.includes("\\\\symbfit{J}+\\\\bm{\\\\alpha}+\\\\abs{x}+\\\\dv{f}{x}"),
          filename: item.filename,
          bytes: new TextEncoder().encode(text).length,
          text,
        };
      })()`, "Markdown Blob content");

      await clickExportOption("SVG", 2);
      const svgState = await waitForEvaluation(`(async () => {
        const item = window.__visualtexCapturedExports?.[1];
        if (!item) return { ready: false };
        const text = await fetch(item.href).then((response) => response.text());
        const opening = text.match(/^<svg\\b[^>]*\\bwidth="([^"]+)"[^>]*\\bheight="([^"]+)"/);
        const width = Number.parseFloat(opening?.[1] ?? "NaN");
        const height = Number.parseFloat(opening?.[2] ?? "NaN");
        return {
          ready:
            item.filename.endsWith(".svg") &&
            text.startsWith("<svg") &&
            !text.includes("<foreignObject") &&
            Number.isFinite(width) &&
            Number.isFinite(height) &&
            height > width,
          filename: item.filename,
          bytes: new TextEncoder().encode(text).length,
          width,
          height,
        };
      })()`, "self-contained SVG Blob content");

      await clickExportOption("PNG", 3);
      const pngState = await waitForEvaluation(`(async () => {
        const item = window.__visualtexCapturedExports?.[2];
        if (!item) return { ready: false };
        const bytes = new Uint8Array(
          await fetch(item.href).then((response) => response.arrayBuffer()),
        );
        const expected = [137, 80, 78, 71, 13, 10, 26, 10];
        const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
        const width = bytes.length >= 24 ? view.getUint32(16) : 0;
        const height = bytes.length >= 24 ? view.getUint32(20) : 0;
        return {
          ready:
            item.filename.endsWith(".png") &&
            bytes.length > expected.length &&
            expected.every((value, index) => bytes[index] === value) &&
            height > width,
          filename: item.filename,
          bytes: bytes.length,
          signature: [...bytes.slice(0, 8)],
          width,
          height,
        };
      })()`, "valid PNG Blob content", 10000);

      console.log(
        JSON.stringify(
          { markdownState, svgState, pngState },
          null,
          2,
        ),
      );
      console.log("Targeted export regression passed");
      return;
    }

    if (scenario === "wrapper-bm") {
      const runWrapperCase = async ({
        command,
        preview = null,
        expected,
        pendingCommand = command,
      }) => {
        await clearField();
        await typeText(command);
        const previewState = preview
          ? await waitForEvaluation(`(() => {
              const item = [...document.querySelectorAll('#mathlive-suggestion-popover li[data-command]')]
                .find((candidate) => candidate.dataset.command === ${JSON.stringify(command)});
              const previewNode = item?.querySelector('[data-visualtex-preview]');
              return {
                ready: Boolean(item && previewNode?.dataset.visualtexPreview === ${JSON.stringify(preview)}),
                command: item?.dataset.command ?? "",
                previewLatex: previewNode?.dataset.visualtexPreview ?? "",
                nativeRegistered: Boolean(item),
              };
            })()`, `${command} visual preview`)
          : await evaluate(`(() => {
              const item = [...document.querySelectorAll('#mathlive-suggestion-popover li[data-command]')]
                .find((candidate) => candidate.dataset.command === ${JSON.stringify(command)});
              const custom = [...document.querySelectorAll('.suggestion-item .suggestion-command')]
                .find((candidate) => candidate.textContent?.trim() === ${JSON.stringify(command)});
              return {
                ready: true,
                skipped: true,
                nativeRegistered: Boolean(item),
                visualTexRegistered: Boolean(custom),
              };
            })()`);
        await key(" ", "Space", 32);
        const emptyState = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const host = field?.closest(".mathfield-host");
          return {
            ready:
              field?.value === "" &&
              field.dataset.pendingWrapperCommand === ${JSON.stringify(pendingCommand)} &&
              host?.classList.contains("has-pending-wrapper-placeholder"),
            value: field?.value ?? "",
            pendingWrapperCommand: field?.dataset.pendingWrapperCommand ?? "",
            placeholderVisible: host?.classList.contains("has-pending-wrapper-placeholder") ?? false,
          };
        })()`, `${command} visual empty wrapper insertion`);
        await key("A", "KeyA", 65);
        const inputState = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const content = field?.shadowRoot?.querySelector('[part="content"]');
          const contentText = content?.textContent ?? "";
          return {
            ready:
              field?.value === ${JSON.stringify(expected)} &&
              contentText.includes("A") &&
              (content?.getBoundingClientRect().width ?? 0) > 0,
            value: field?.value ?? "",
            contentText,
            contentWidth: content?.getBoundingClientRect().width ?? 0,
          };
        })()`, `${command} input remains visible and preserves source command`);
        return { previewState, emptyState, inputState };
      };

      const bmState = await runWrapperCase({
        command: String.raw`\bm`,
        preview: String.raw`\bm{\alpha A}`,
        expected: String.raw`\bm{A}`,
      });
      const mathbfitState = await runWrapperCase({
        command: String.raw`\mathbfit`,
        preview: String.raw`\mathbfit{ABC}`,
        expected: String.raw`\mathbfit{A}`,
      });
      const symbfitState = await runWrapperCase({
        command: String.raw`\symbfit`,
        preview: String.raw`\symbfit{ABC\alpha}`,
        expected: String.raw`\symbfit{A}`,
      });
      const simbfitAliasState = await runWrapperCase({
        command: String.raw`\simbfit`,
        preview: String.raw`\symbfit{ABC\alpha}`,
        pendingCommand: String.raw`\symbfit`,
        expected: String.raw`\symbfit{A}`,
      });
      const boldmathState = await runWrapperCase({
        command: String.raw`\boldmath`,
        preview: String.raw`\mathbfit{A\alpha}`,
        pendingCommand: String.raw`\mathbfit`,
        expected: String.raw`\mathbfit{A}`,
      });
      const bbbAliasState = await runWrapperCase({
        command: String.raw`\Bbb`,
        preview: String.raw`\Bbb{ABC}`,
        pendingCommand: String.raw`\mathbb`,
        expected: String.raw`\mathbb{A}`,
      });
      const absState = await runWrapperCase({
        command: String.raw`\abs`,
        preview: String.raw`\abs{x}`,
        expected: String.raw`\abs{A}`,
      });

      await clearField();
      await typeText(String.raw`\dv`);
      await key(" ", "Space", 32);
      const derivativeShortcutState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const value = field?.value ?? "";
        const placeholderCount = (value.match(/\\\\placeholder\\{\\}/g) ?? []).length;
        return {
          ready: value.includes('\\\\dv{') && placeholderCount === 2 && !field?.selectionIsCollapsed,
          value,
          placeholderCount,
          selection: field ? structuredClone(field.selection) : null,
        };
      })()`, "\\dv two-argument shorthand placeholders");

      const structuredPlaceholderState = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.setValue("p+(z+r)+q+\\\\placeholder{}", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        return {
          ready: field.value.includes("z") && field.lastOffset > 0,
          value: field.value,
          lastOffset: field.lastOffset,
          errors: field.errors,
        };
      })()`);
      await evaluate(`document.querySelector("math-field").dispatchEvent(new InputEvent("input", {
        bubbles: true,
        composed: true,
        inputType: "insertText",
      }))`);
      await sleep(180);
      const structuredPlaceholderAfterInput = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}").state ?? {};
        return {
          ready: field.value.includes("z") && field.lastOffset > 0,
          value: field.value,
          lastOffset: field.lastOffset,
          persistedLatex: persisted.lines?.[0]?.latex ?? "",
        };
      })()`);
      console.log(JSON.stringify({ bmState, mathbfitState, symbfitState, simbfitAliasState, boldmathState, bbbAliasState, absState, derivativeShortcutState, structuredPlaceholderState, structuredPlaceholderAfterInput }));
      console.log("Targeted font/package shorthand wrapper regression passed");
      return;
    }

    if (
      scenario === "wrapper" ||
      scenario === "wrapper-auto" ||
      scenario === "wrapper-continuous"
    ) {
      await clearField();
      await typeText("abcdefghij");
      await typeText("\\mathbb");
      const nativeStructure = await waitForEvaluation(`(() => {
        const panel = document.getElementById("mathlive-suggestion-popover");
        const items = [...(panel?.querySelectorAll("li[data-command]") ?? [])];
        return {
          ready: Boolean(panel?.classList.contains("is-visible") && items.length),
          items: items.map((item) => ({
            command: item.dataset.command ?? "",
            html: item.innerHTML,
            classes: [...item.querySelectorAll("*")].map((node) => node.className).filter(Boolean),
          })),
        };
      })()`, "MathLive mathbb suggestion structure");
      const previewState = await waitForEvaluation(`(() => {
        const item = [...document.querySelectorAll('#mathlive-suggestion-popover li[data-command]')]
          .find((candidate) => candidate.dataset.command === "\\\\mathbb");
        const preview = item?.querySelector('[data-visualtex-preview]');
        return {
          ready: Boolean(item && preview?.dataset.visualtexPreview === "\\\\mathbb{ABC}"),
          nativeVisible: document.getElementById("mathlive-suggestion-popover")?.classList.contains("is-visible") ?? false,
          previewLatex: preview?.dataset.visualtexPreview ?? "",
          previewText: preview?.textContent ?? "",
        };
      })()`, "mathbb visual preview");

      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        key: " ",
        code: "Space",
        windowsVirtualKeyCode: 32,
        nativeVirtualKeyCode: 32,
      });
      await client.send("Input.dispatchKeyEvent", {
        type: "keyUp",
        key: " ",
        code: "Space",
        windowsVirtualKeyCode: 32,
        nativeVirtualKeyCode: 32,
      });
      await sleep(80);
      const insertedState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const host = field?.closest(".mathfield-host");
        const placeholderStyle = host ? getComputedStyle(host, "::after") : null;
        const fakeCaretStyle = host ? getComputedStyle(host, "::before") : null;
        const nativeCaret = field?.shadowRoot?.querySelector(".ML__caret");
        const nativeCaretStyle = nativeCaret
          ? getComputedStyle(nativeCaret, "::after")
          : null;
        const nativeCaretBounds = nativeCaret?.getBoundingClientRect();
        const hostBounds = host?.getBoundingClientRect();
        const modelBounds = field?.getElementInfo(field.position)?.bounds;
        const placeholderLeft = Number.parseFloat(
          host?.style.getPropertyValue("--pending-wrapper-left") ?? "NaN",
        );
      const placeholderTop = Number.parseFloat(
        host?.style.getPropertyValue("--pending-wrapper-top") ?? "NaN",
      );
      const placeholderWidth = Number.parseFloat(
        host?.style.getPropertyValue("--pending-wrapper-width") ?? "NaN",
      );
      const placeholderHeight = Number.parseFloat(
        host?.style.getPropertyValue("--pending-wrapper-height") ?? "NaN",
      );
      const anchorTop = Number.parseFloat(
        host?.dataset.pendingWrapperAnchorY ?? "NaN",
      );
      const frameLeft = placeholderLeft - placeholderWidth / 2;
      const formulaFontSize =
        Number.parseFloat(field?.style.fontSize ?? "") || 54;
      const minimumFrameHeight = Math.max(12, formulaFontSize * 0.52);
      const maximumFrameHeight = Math.max(
        minimumFrameHeight,
        formulaFontSize * 1.08,
      );
        const expectedLeft =
          modelBounds && hostBounds
            ? modelBounds.right - hostBounds.left
            : Number.NaN;
        const expectedTop =
          modelBounds && hostBounds
            ? modelBounds.top - hostBounds.top + modelBounds.height / 2
            : Number.NaN;
        return {
          ready:
            field.value === "abcdefghij" &&
            field.dataset.pendingWrapperCommand === "\\\\mathbb" &&
            host?.classList.contains("has-pending-wrapper-placeholder") &&
            placeholderStyle?.borderStyle === "solid" &&
            Number.parseFloat(placeholderStyle?.borderWidth ?? "0") <= 1.1 &&
            Boolean(nativeCaret) &&
            fakeCaretStyle?.content === "none" &&
            nativeCaretStyle?.visibility === "visible" &&
            nativeCaretStyle?.animationName.includes("caret-blink") &&
          Math.abs(frameLeft - expectedLeft) <= 2 &&
          Math.abs(placeholderTop - anchorTop) <= 1.5 &&
          placeholderHeight >= minimumFrameHeight - 0.5 &&
          placeholderHeight <= maximumFrameHeight + 0.5 &&
            Math.abs(placeholderLeft - (hostBounds?.width ?? 0) / 2) >= 20 &&
            document.querySelectorAll("math-field").length === 1,
          value: field.value,
          pendingWrapperCommand: field.dataset.pendingWrapperCommand ?? "",
          placeholderClass: host?.classList.contains("has-pending-wrapper-placeholder") ?? false,
          placeholderBorderStyle: placeholderStyle?.borderStyle ?? "",
          placeholderBorderWidth: placeholderStyle?.borderWidth ?? "",
        placeholderLeft,
        placeholderTop,
        placeholderWidth,
        placeholderHeight,
        frameLeft,
        anchorTop,
        minimumFrameHeight,
        maximumFrameHeight,
          expectedLeft,
          expectedTop,
          fakeCaretContent: fakeCaretStyle?.content ?? "",
          nativeCaretVisibility: nativeCaretStyle?.visibility ?? "",
          nativeCaretAnimation: nativeCaretStyle?.animationName ?? "",
          nativeCaretBorder: nativeCaretStyle?.borderRightWidth ?? "",
          hostCenter: (hostBounds?.width ?? 0) / 2,
          lineCount: document.querySelectorAll("math-field").length,
        };
      })()`, "mathbb visual empty wrapper insertion");

      await key("A", "KeyA", 65);
      const autoExitState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const host = field?.closest(".mathfield-host");
        return {
          ready:
            field.value === "abcdefghij\\\\mathbb{A}" &&
            !field.dataset.pendingWrapperCommand &&
            !host?.classList.contains("has-pending-wrapper-placeholder"),
          value: field.value,
          pendingWrapperCommand: field.dataset.pendingWrapperCommand ?? "",
          placeholderClass: host?.classList.contains("has-pending-wrapper-placeholder") ?? false,
        };
      })()`, "mathbb default single-character auto exit");
      await key("B", "KeyB", 66);
      const normalFontState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const sink = field?.shadowRoot?.querySelector('[part="keyboard-sink"]');
        return {
          ready: field?.value === "abcdefghij\\\\mathbb{A}B",
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          hasFocus: field?.hasFocus?.() ?? false,
          activeTag: document.activeElement?.tagName ?? "",
          sinkTag: sink?.tagName ?? "",
          sinkValue: sink?.value ?? "",
          pendingWrapperCommand: field?.dataset.pendingWrapperCommand ?? "",
          position: field?.position ?? -1,
          lastOffset: field?.lastOffset ?? -1,
        };
      })()`, "normal font after mathbb auto exit");
      await key("Enter", "Enter", 13);
      const enterState = await waitForEvaluation(`(() => {
        const fields = [...document.querySelectorAll("math-field")];
        return {
          ready: fields.length === 2 && fields[0]?.value === "abcdefghij\\\\mathbb{A}B",
          lineCount: fields.length,
          values: fields.map((field) => field.value),
        };
      })()`, "Enter creates a new formula line after wrapper input");

      await evaluate(`(() => {
        const storageKey = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
        const line = { id: crypto.randomUUID(), latex: "" };
        persisted.state = { ...(persisted.state || {}), lines: [line], activeLineId: line.id };
        localStorage.setItem(storageKey, JSON.stringify(persisted));
        location.reload();
      })()`);
      await waitForEvaluation(`(() => ({ ready: Boolean(document.querySelector("math-field")?.shadowRoot) }))()`, "fresh field for mathcal test");
      await clearField();
      await typeText("\\mathcal");
      await key(" ", "Space", 32);
      await key("g", "KeyG", 71);
      const lowercaseScriptState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        return {
          ready:
            field?.value === "\\\\mathscr{g}" &&
            (field.shadowRoot?.textContent ?? "").includes("ℊ"),
          value: field?.value ?? "",
          shadowText: field?.shadowRoot?.textContent ?? "",
        };
      })()`, "lowercase mathcal compatibility uses mathscr");

      await waitForEvaluation(`(() => ({ ready: Boolean(document.querySelector(".canvas-input-behavior-trigger")) }))()`, "input behavior trigger");
      await evaluate(`document.querySelector(".canvas-input-behavior-trigger").click()`);
      await waitForEvaluation(`(() => ({ ready: Boolean(document.querySelector(".input-behavior-popover")) }))()`, "input behavior menu");
      await evaluate(`(() => {
        const option = [...document.querySelectorAll(".input-behavior-option")]
          .find((label) => label.querySelector("strong")?.textContent?.includes("字体命令输入后跳出"));
        const checkbox = option?.querySelector('input[type="checkbox"]');
        if (!checkbox) throw new Error("Wrapper auto-exit checkbox was not found");
        checkbox.click();
        document.querySelector(".canvas-input-behavior-trigger").click();
      })()`);
      await clearField();
      await typeText("\\mathbb");
      await key(" ", "Space", 32);
      await key("A", "KeyA", 65);
      const persistentOneCharacterState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const host = field?.closest(".mathfield-host");
        const frameStyle = host ? getComputedStyle(host, "::after") : null;
        return {
          ready:
            field?.value === "\\\\mathbb{A}" &&
            field.dataset.pendingWrapperCommand === "\\\\mathbb" &&
            host?.classList.contains("has-pending-wrapper-placeholder") &&
            Number.parseFloat(frameStyle?.width ?? "0") > 18,
          value: field?.value ?? "",
          pendingWrapperCommand: field?.dataset.pendingWrapperCommand ?? "",
          pendingWrapperLength: host?.dataset.pendingWrapperLength ?? "",
          frameWidth: Number.parseFloat(frameStyle?.width ?? "0"),
        };
      })()`, "disabled wrapper auto exit keeps a visible one-character input frame");
      await key("B", "KeyB", 66);
      const continuousState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const host = field?.closest(".mathfield-host");
        const frameStyle = host ? getComputedStyle(host, "::after") : null;
        return {
          ready:
            field?.value === "\\\\mathbb{AB}" &&
            field.dataset.pendingWrapperCommand === "\\\\mathbb" &&
            host?.classList.contains("has-pending-wrapper-placeholder") &&
            host?.dataset.pendingWrapperLength === "2",
          value: field?.value ?? "",
          pendingWrapperCommand: field?.dataset.pendingWrapperCommand ?? "",
          pendingWrapperLength: host?.dataset.pendingWrapperLength ?? "",
          frameWidth: Number.parseFloat(frameStyle?.width ?? "0"),
        };
      })()`, "disabled wrapper auto exit keeps continuous input");
      if (!(continuousState.frameWidth > persistentOneCharacterState.frameWidth + 1)) {
        throw new Error(
          `Wrapper frame did not grow with its content: ${persistentOneCharacterState.frameWidth} -> ${continuousState.frameWidth}`,
        );
      }
      await key("Enter", "Enter", 13);
      const manualConfirmState = await waitForEvaluation(`(() => {
        const fields = [...document.querySelectorAll("math-field")];
        const field = fields[0];
        const host = field?.closest(".mathfield-host");
        return {
          ready:
            fields.length === 1 &&
            field?.value === "\\\\mathbb{AB}" &&
            !field.dataset.pendingWrapperCommand &&
            !host?.classList.contains("has-pending-wrapper-placeholder"),
          lineCount: fields.length,
          value: field?.value ?? "",
          pendingWrapperCommand: field?.dataset.pendingWrapperCommand ?? "",
          frameVisible: host?.classList.contains("has-pending-wrapper-placeholder") ?? false,
        };
      })()`, "Enter confirms a persistent wrapper without adding a line");
      await key("C", "KeyC", 67);
      const postConfirmState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        return {
          ready: field?.value === "\\\\mathbb{AB}C",
          value: field?.value ?? "",
        };
      })()`, "typing after Enter leaves the confirmed font wrapper");

      const nestedWrapperCases = [
        {
          name: "parentheses",
          command: String.raw`\mathbb`,
          source: String.raw`p+(z+r)+q+\placeholder{}`,
        },
        {
          name: "fraction numerator",
          command: String.raw`\mathbf`,
          source: String.raw`p+\frac{z+n}{d}+q+\placeholder{}`,
        },
        {
          name: "fraction denominator",
          command: String.raw`\mathcal`,
          source: String.raw`p+\frac{n}{z+d}+q+\placeholder{}`,
        },
        {
          name: "integral upper limit",
          command: String.raw`\mathfrak`,
          source: String.raw`p+\int_{l}^{z+u}f\,dx+q+\placeholder{}`,
        },
        {
          name: "integral lower limit",
          command: String.raw`\mathbb`,
          source: String.raw`p+\int_{z+l}^{u}f\,dx+q+\placeholder{}`,
        },
        {
          name: "integral integrand",
          command: String.raw`\mathbf`,
          source: String.raw`p+\int_{l}^{u}(z+f)\,dx+q+\placeholder{}`,
        },
        {
          name: "summation upper limit",
          command: String.raw`\mathcal`,
          source: String.raw`p+\sum_{i=0}^{z+n}a_{i}+q+\placeholder{}`,
        },
        {
          name: "summation lower limit",
          command: String.raw`\mathfrak`,
          source: String.raw`p+\sum_{z+i}^{n}a_{i}+q+\placeholder{}`,
        },
        {
          name: "square root",
          command: String.raw`\mathbb`,
          source: String.raw`p+\sqrt{z+s}+q+\placeholder{}`,
        },
        {
          name: "superscript",
          command: String.raw`\mathbf`,
          source: String.raw`p+x^{z+u}+q+\placeholder{}`,
        },
        {
          name: "subscript",
          command: String.raw`\mathcal`,
          source: String.raw`p+x_{z+l}+q+\placeholder{}`,
        },
        {
          name: "matrix cell",
          command: String.raw`\mathfrak`,
          source: String.raw`p+\begin{matrix}a&z+m\\c&d\end{matrix}+q+\placeholder{}`,
        },
      ];

      const setWrapperAutoExit = async (enabled) => {
        await evaluate(`(() => {
          const storageKey = "visualtex-editor";
          const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
          const line = { id: crypto.randomUUID(), latex: "" };
          persisted.state = {
            ...(persisted.state || {}),
            lines: [line],
            activeLineId: line.id,
            inputBehavior: {
              ...(persisted.state?.inputBehavior || {}),
              autoExitWrapperCommand: ${enabled},
            },
          };
          localStorage.setItem(storageKey, JSON.stringify(persisted));
          location.reload();
        })()`);
        await waitForEvaluation(
          `(() => ({ ready: Boolean(document.querySelector("math-field")) }))()`,
          `formula field with wrapper auto-exit ${enabled}`,
        );
        await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const now = performance.now();
          if (!field?.isConnected) return { ready: false };
          if (window.__visualtexWrapperStableField !== field) {
            window.__visualtexWrapperStableField = field;
            window.__visualtexWrapperStableSince = now;
            return { ready: false, stableFor: 0 };
          }
          const stableFor = now - (window.__visualtexWrapperStableSince ?? now);
          return { ready: stableFor >= 180, stableFor };
        })()`, `stable hydrated field with wrapper auto-exit ${enabled}`);
        await focusField();
      };

      const prepareNestedWrapperCase = async ({ name, source }) => {
        await waitForEvaluation(
          `(() => ({ ready: Boolean(document.querySelector("math-field")?.isConnected) }))()`,
          `stable field before nested wrapper case: ${name}`,
        );
        await evaluate(`(() => {
          const field = document.querySelector("math-field");
          field.setValue(${JSON.stringify(source)}, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceAll",
            selectionMode: "after",
            silenceNotifications: true,
          });
          field.dispatchEvent(new InputEvent("input", {
            bubbles: true,
            composed: true,
            inputType: "insertText",
          }));
        })()`);
        await sleep(120);
        await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          if (!field?.isConnected) return { ready: false };
          let markerEnd = -1;
          const candidates = [];
          for (let end = 1; end <= field.lastOffset; end += 1) {
            const rangeLatex = field.getValue(end - 1, end, "latex").trim();
            const infoLatex = field.getElementInfo(end)?.latex?.trim() ?? "";
            if (rangeLatex === "z" || infoLatex === "z") {
              candidates.push({ end, rangeLatex, infoLatex });
              if (markerEnd < 0) markerEnd = end;
            }
          }
          if (markerEnd < 0) {
            return {
              ready: false,
              value: field.value,
              lastOffset: field.lastOffset,
              candidates,
            };
          }
          field.focus();
          field.selection = {
            ranges: [[markerEnd, markerEnd]],
            direction: "none",
          };
          field.position = markerEnd;
          field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
          return {
            ready: field.position === markerEnd,
            name: ${JSON.stringify(name)},
            value: field.value,
            markerEnd,
            lastOffset: field.lastOffset,
            candidates,
          };
        })()`, `nested wrapper model anchor: ${name}`);
        await sleep(80);
        return await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const host = field?.closest(".mathfield-host");
          if (!field || !host) return { ready: false };
          const hostBounds = host.getBoundingClientRect();
          const info = field.getElementInfo(field.position);
          const bounds = info?.bounds;
          const caretMarkers = [
            ...(field.shadowRoot?.querySelectorAll(
              ".visualtex-structural-placeholder-caret, .ML__caret, .ML__text-caret, .ML__latex-caret",
            ) ?? []),
          ]
            .map((marker) => {
              const markerBounds = marker.getBoundingClientRect();
              const style = getComputedStyle(marker);
              return {
                classes: marker.className,
                left: markerBounds.left - hostBounds.left,
                right: markerBounds.right - hostBounds.left,
                centerY:
                  markerBounds.top -
                  hostBounds.top +
                  markerBounds.height / 2,
                width: markerBounds.width,
                height: markerBounds.height,
                visible:
                  style.display !== "none" &&
                  style.visibility !== "hidden" &&
                  Number.parseFloat(style.opacity || "1") > 0 &&
                  markerBounds.height > 0,
              };
            })
            .filter((marker) => marker.visible)
            .sort(
              (first, second) =>
                first.width - second.width ||
                first.height - second.height,
            );
          const caretMarker = caretMarkers[0];
          if (!bounds || bounds.height <= 0) {
            return {
              ready: false,
              value: field.value,
              position: field.position,
              latex: info?.latex ?? "",
            };
          }
          const expectedLeft =
            caretMarker?.left ?? bounds.right - hostBounds.left;
          const expectedTop =
            caretMarker?.centerY ??
            bounds.top - hostBounds.top + bounds.height / 2;
          const expectedHeight = caretMarker?.height ?? bounds.height;
          host.dataset.testExpectedWrapperAnchorX = String(expectedLeft);
          host.dataset.testExpectedWrapperAnchorY = String(expectedTop);
          host.dataset.testExpectedWrapperAnchorHeight =
            String(expectedHeight);
          return {
            ready: true,
            name: ${JSON.stringify(name)},
            expectedLeft,
            expectedTop,
            expectedHeight,
            caretMarkers,
            modelLatex: info?.latex ?? "",
            modelDepth: info?.depth ?? -1,
            value: field.value,
            position: field.position,
          };
        })()`, `rendered nested wrapper model bounds: ${name}`);
      };

      const waitForNestedWrapperState = async ({
        name,
        command,
        expectedSource,
        pending,
      }) => {
        return await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const host = field?.closest(".mathfield-host");
          if (!field || !host) return { ready: false };
          const probe = document.createElement("math-field");
          probe.style.display = "none";
          document.body.append(probe);
          probe.setValue(${JSON.stringify(expectedSource)}, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceAll",
            selectionMode: "after",
            silenceNotifications: true,
          });
          const expected = probe.value;
          const normalizedValue = field.value.replaceAll(" ", "");
          const normalizedExpected = expected.replaceAll(" ", "");
          probe.remove();
          const hostBounds = host.getBoundingClientRect();
          const placeholderLeft = Number.parseFloat(
            host.style.getPropertyValue("--pending-wrapper-left") || "NaN",
          );
          const placeholderTop = Number.parseFloat(
            host.style.getPropertyValue("--pending-wrapper-top") || "NaN",
          );
        const placeholderWidth = Number.parseFloat(
          host.style.getPropertyValue("--pending-wrapper-width") || "NaN",
        );
        const placeholderHeight = Number.parseFloat(
          host.style.getPropertyValue("--pending-wrapper-height") || "NaN",
        );
          const expectedAnchorLeft = Number.parseFloat(
            host.dataset.testExpectedWrapperAnchorX || "NaN",
          );
          const expectedAnchorTop = Number.parseFloat(
            host.dataset.testExpectedWrapperAnchorY || "NaN",
          );
        const productAnchorTop = Number.parseFloat(
          host.dataset.pendingWrapperAnchorY || "NaN",
        );
          const expectedAnchorHeight = Number.parseFloat(
            host.dataset.testExpectedWrapperAnchorHeight || "NaN",
          );
          const currentInfo = field.getElementInfo(field.position);
          const currentBounds = currentInfo?.bounds;
          const currentModelRight = currentBounds
            ? currentBounds.right - hostBounds.left
            : Number.NaN;
          const currentModelTop = currentBounds
            ? currentBounds.top - hostBounds.top + currentBounds.height / 2
            : Number.NaN;
        const frameLeft = placeholderLeft - placeholderWidth / 2;
        const formulaFontSize =
          Number.parseFloat(field.style.fontSize) || 54;
        const minimumFrameHeight = Math.max(12, formulaFontSize * 0.52);
        const maximumFrameHeight = Math.max(
          minimumFrameHeight,
          formulaFontSize * 1.08,
        );
          const hasPending = field.dataset.pendingWrapperCommand === ${JSON.stringify(command)};
          const frameAligned =
            !${pending} ||
            (host.classList.contains("has-pending-wrapper-placeholder") &&
            Number.isFinite(productAnchorTop) &&
            Math.abs(frameLeft - expectedAnchorLeft) <= 2.5 &&
            Math.abs(productAnchorTop - expectedAnchorTop) <= 6 &&
            Math.abs(placeholderTop - productAnchorTop) <= 1.5 &&
            placeholderHeight >= minimumFrameHeight - 0.5 &&
            placeholderHeight <= maximumFrameHeight + 0.5);
          return {
            ready:
              normalizedValue === normalizedExpected &&
              hasPending === ${pending} &&
              host.classList.contains("has-pending-wrapper-placeholder") === ${pending} &&
              frameAligned &&
              field.shadowRoot?.querySelectorAll(".ML__raw-latex").length === 0 &&
              document.querySelectorAll("math-field").length === 1,
            name: ${JSON.stringify(name)},
            value: field.value,
            expected,
            normalizedValue,
            normalizedExpected,
          position: field.position,
          lastOffset: field.lastOffset,
          pendingWrapperCommand: field.dataset.pendingWrapperCommand ?? "",
          frameVisible: host.classList.contains("has-pending-wrapper-placeholder"),
          placeholderLeft,
          placeholderTop,
          placeholderWidth,
          placeholderHeight,
          frameLeft,
          expectedAnchorLeft,
          expectedAnchorTop,
          productAnchorTop,
          minimumFrameHeight,
          maximumFrameHeight,
            expectedAnchorHeight,
            currentModelRight,
            currentModelTop,
            currentModelLatex: currentInfo?.latex ?? "",
            rawLatexCount: field.shadowRoot?.querySelectorAll(".ML__raw-latex").length ?? -1,
            lineCount: document.querySelectorAll("math-field").length,
          };
        })()`, `nested wrapper state: ${name}`);
      };

      const expectedNestedWrapperSource = (
        testCase,
        content,
        trailingContent = "",
      ) =>
        testCase.source.replace(
          "z",
          `z${testCase.command}{${content}}${trailingContent}`,
        );

      const autoExitNestedStates = [];
      if (scenario !== "wrapper-continuous") {
        await setWrapperAutoExit(true);
        for (const testCase of nestedWrapperCases) {
        await prepareNestedWrapperCase(testCase);
        await typeText(testCase.command);
        await key(" ", "Space", 32);
        autoExitNestedStates.push({
          phase: "empty",
          ...(await waitForNestedWrapperState({
            ...testCase,
            expectedSource: testCase.source,
            pending: true,
          })),
        });
        await key("A", "KeyA", 65);
        await key("B", "KeyB", 66);
          autoExitNestedStates.push({
            phase: "auto-exit",
            ...(await waitForNestedWrapperState({
              ...testCase,
            expectedSource: expectedNestedWrapperSource(testCase, "A", "B"),
              pending: false,
            })),
          });
        }
      }

      const continuousNestedStates = [];
      if (scenario !== "wrapper-auto") {
        await setWrapperAutoExit(false);
        for (const testCase of nestedWrapperCases) {
        await prepareNestedWrapperCase(testCase);
        await typeText(testCase.command);
        await key(" ", "Space", 32);
        continuousNestedStates.push({
          phase: "empty",
          ...(await waitForNestedWrapperState({
            ...testCase,
            expectedSource: testCase.source,
            pending: true,
          })),
        });
        await key("A", "KeyA", 65);
        await key("B", "KeyB", 66);
        continuousNestedStates.push({
          phase: "continuous",
          ...(await waitForNestedWrapperState({
            ...testCase,
            expectedSource: expectedNestedWrapperSource(testCase, "AB"),
            pending: true,
          })),
        });
        await key("Backspace", "Backspace", 8);
        continuousNestedStates.push({
          phase: "continuous-backspace",
          ...(await waitForNestedWrapperState({
            ...testCase,
            expectedSource: expectedNestedWrapperSource(testCase, "A"),
            pending: true,
          })),
        });
        await key("B", "KeyB", 66);
        continuousNestedStates.push({
          phase: "continuous-restored",
          ...(await waitForNestedWrapperState({
            ...testCase,
            expectedSource: expectedNestedWrapperSource(testCase, "AB"),
            pending: true,
          })),
        });
        await key("Enter", "Enter", 13);
        await key("C", "KeyC", 67);
          continuousNestedStates.push({
            phase: "confirmed",
            ...(await waitForNestedWrapperState({
              ...testCase,
            expectedSource: expectedNestedWrapperSource(testCase, "AB", "C"),
              pending: false,
            })),
          });
        }
      }

      console.log(JSON.stringify({
        previewState,
        insertedState,
        autoExitState,
        normalFontState,
        enterState,
        lowercaseScriptState,
        persistentOneCharacterState,
        continuousState,
        manualConfirmState,
        postConfirmState,
        autoExitNestedStates,
        continuousNestedStates,
      }, null, 2));
      console.log("Targeted wrapper regression passed");
      return;
    }

    if (scenario === "scripts") {
      const setInputBehavior = async (autoExitSuperscript, autoExitSubscript) => {
        await evaluate(`(() => {
          const storageKey = "visualtex-editor";
          const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
          persisted.state = {
            ...(persisted.state || {}),
            inputBehavior: {
              ...(persisted.state?.inputBehavior || {}),
              autoExitSuperscript: ${autoExitSuperscript},
              autoExitSubscript: ${autoExitSubscript},
            },
          };
          localStorage.setItem(storageKey, JSON.stringify(persisted));
          location.reload();
        })()`);
        await sleep(650);
        await waitForEvaluation(
          `(() => ({ ready: Boolean(document.querySelector("math-field")) }))()`,
          "formula field after script-setting reload",
        );
        await focusField();
        await clearField();
      };

      const scriptKey = async (character, code, virtualKeyCode) => {
        const common = {
          key: character,
          code,
          windowsVirtualKeyCode: virtualKeyCode,
          nativeVirtualKeyCode: virtualKeyCode,
        };
        await client.send("Input.dispatchKeyEvent", {
          type: "keyDown",
          ...common,
          text: character,
          unmodifiedText: character,
        });
        await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
        await sleep(60);
      };

      const runCase = async ({
        name,
        autoExitSuperscript,
        autoExitSubscript,
        scriptCharacter,
        scriptCode,
        scriptVirtualKeyCode,
        expected,
      }) => {
        await setInputBehavior(autoExitSuperscript, autoExitSubscript);
        await key("x", "KeyX", 88);
        await scriptKey(scriptCharacter, scriptCode, scriptVirtualKeyCode);
        await key("a", "KeyA", 65);
        await key("b", "KeyB", 66);
        return await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const markers = [...(field?.shadowRoot?.querySelectorAll(
            ".ML__placeholder-selected, .ML__selected, .ML__caret",
          ) ?? [])];
          return {
            ready: field?.value === ${JSON.stringify(expected)},
            name: ${JSON.stringify(name)},
            value: field?.value ?? "",
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
            markerAncestors: markers.map((marker) => {
              const chain = [];
              let node = marker;
              while (node && chain.length < 7) {
                const bounds = node.getBoundingClientRect();
                chain.push({
                  tag: node.tagName,
                  className: node.className || "",
                  text: node.textContent || "",
                  top: bounds.top,
                  height: bounds.height,
                });
                node = node.parentElement;
              }
              return chain;
            }),
            contentHtml:
              field?.shadowRoot?.querySelector('[part="content"]')?.innerHTML ?? "",
          };
        })()`, `script independence: ${name}`, 3500);
      };

      const cases = [];
      cases.push(await runCase({
        name: "superscript disabled while subscript enabled",
        autoExitSuperscript: false,
        autoExitSubscript: true,
        scriptCharacter: "^",
        scriptCode: "Digit6",
        scriptVirtualKeyCode: 54,
        expected: "x^{ab}",
      }));
      cases.push(await runCase({
        name: "superscript enabled while subscript disabled",
        autoExitSuperscript: true,
        autoExitSubscript: false,
        scriptCharacter: "^",
        scriptCode: "Digit6",
        scriptVirtualKeyCode: 54,
        expected: "x^{a}b",
      }));
      cases.push(await runCase({
        name: "subscript disabled while superscript enabled",
        autoExitSuperscript: true,
        autoExitSubscript: false,
        scriptCharacter: "_",
        scriptCode: "Minus",
        scriptVirtualKeyCode: 189,
        expected: "x_{ab}",
      }));
      cases.push(await runCase({
        name: "subscript enabled while superscript disabled",
        autoExitSuperscript: false,
        autoExitSubscript: true,
        scriptCharacter: "_",
        scriptCode: "Minus",
        scriptVirtualKeyCode: 189,
        expected: "x_{a}b",
      }));

      console.log(JSON.stringify({ cases }, null, 2));
      console.log("Targeted independent script auto-exit regression passed");
      return;
    }

    if (scenario === "formula-formatting" || scenario === "font-variant-formatting") {
      await waitForEvaluation(`(() => ({
        ready: [
          '[data-formula-selection-bold]',
          '[data-formula-selection-italic]',
          '[data-formula-selection-color]',
          '[data-formula-selection-background]',
        ].every((selector) => {
          const button = document.querySelector(selector);
          if (!(button instanceof HTMLElement)) return false;
          const rect = button.getBoundingClientRect();
          return rect.width > 0 && rect.height > 0;
        }),
      }))()`, "desktop formula formatting controls");

      const dispatchFormattingToggle = async (selector) => {
        const encodedSelector = JSON.stringify(selector);
        const target = await evaluate(`(() => {
          const field = document.querySelector('math-field');
          const button = document.querySelector(${encodedSelector});
          if (!field || !(button instanceof HTMLElement)) return null;
          field.focus();
          field.selection = {
            ranges: [[0, field.lastOffset]],
            direction: 'forward',
          };
          globalThis.__visualTexFormattingTrace = { events: [] };
          for (const type of ['pointerenter', 'pointerdown', 'mousedown', 'mouseup', 'click']) {
            button.addEventListener(
              type,
              () => globalThis.__visualTexFormattingTrace.events.push({
                type,
                selection: structuredClone(field.selection),
                activeTag: document.activeElement?.tagName ?? null,
              }),
              { once: true },
            );
          }
          const rect = button.getBoundingClientRect();
          return {
            x: rect.left + rect.width / 2,
            y: rect.top + rect.height / 2,
            selection: structuredClone(field.selection),
          };
        })()`);
        if (!target) throw new Error(`Unable to resolve formatting button ${selector}`);
        await client.send("Input.dispatchMouseEvent", {
          type: "mouseMoved",
          x: target.x,
          y: target.y,
        });
        await client.send("Input.dispatchMouseEvent", {
          type: "mousePressed",
          x: target.x,
          y: target.y,
          button: "left",
          buttons: 1,
          clickCount: 1,
        });
        await client.send("Input.dispatchMouseEvent", {
          type: "mouseReleased",
          x: target.x,
          y: target.y,
          button: "left",
          buttons: 0,
          clickCount: 1,
        });
        await sleep(60);
        const pointerDelivered = await evaluate(`Boolean(globalThis.__visualTexFormattingTrace?.events?.length)`);
        if (!pointerDelivered) {
          await evaluate(`(() => {
            const button = document.querySelector(${encodedSelector});
            if (!(button instanceof HTMLElement)) return false;
            button.dispatchEvent(new PointerEvent('pointerenter', {
              bubbles: false,
              pointerId: 1,
              pointerType: 'mouse',
              button: 0,
              buttons: 0,
            }));
            button.dispatchEvent(new PointerEvent('pointerdown', {
              bubbles: true,
              cancelable: true,
              pointerId: 1,
              pointerType: 'mouse',
              button: 0,
              buttons: 1,
            }));
            return true;
          })()`);
        }
        await sleep(100);
        return evaluate(`(() => {
          const field = document.querySelector('math-field');
          const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}');
          return {
            initialSelection: ${JSON.stringify(target.selection)},
            events: globalThis.__visualTexFormattingTrace?.events ?? [],
            selection: field ? structuredClone(field.selection) : null,
            value: field?.value ?? null,
            storeValue: persisted.state?.lines?.[0]?.latex ?? null,
          };
        })()`);
      };

      const setFormattingValue = async (latex) => {
        const encodedLatex = JSON.stringify(latex);
        await evaluate(`(() => {
          const field = document.querySelector('math-field');
          if (!field) return false;
          field.setValue(${encodedLatex}, {
            mode: 'math',
            format: 'latex',
            insertionMode: 'replaceAll',
            selectionMode: 'after',
            silenceNotifications: true,
          });
          field.dispatchEvent(new InputEvent('input', {
            bubbles: true,
            composed: true,
            inputType: 'insertText',
          }));
          field.focus();
          return true;
        })()`);
        await sleep(60);
        const expectedCanonical = latex.replace(/\s+/g, '');
        const expectedSerialized = JSON.stringify(expectedCanonical);
        await waitForEvaluation(`(() => {
          const field = document.querySelector('math-field');
          const actual = field?.value?.replace(/\\s+/g, '') ?? '';
          return { ready: actual === ${expectedSerialized}, actual };
        })()`, `formula value hydration for ${latex}`);
      };

      const readFormattingValue = () => evaluate(`(() => {
        const field = document.querySelector('math-field');
        return field?.value?.replace(/\\s+/g, '') ?? '';
      })()`);

      await setFormattingValue('abc');
      const boldTrace = await dispatchFormattingToggle('[data-formula-selection-bold]');
      const boldApplied = await readFormattingValue();
      if (boldApplied !== String.raw`\mathbfit{abc}`) {
        throw new Error(
          `Bold toggle must preserve default math italic as \\mathbfit; received ${boldApplied}; trace=${JSON.stringify(boldTrace)}`,
        );
      }
      const boldVisualStyle = await evaluate(`(() => {
        const field = document.querySelector('math-field');
        const node = field?.shadowRoot?.querySelector('.ML__mathbfit, .ML__cmr.ML__bold.ML__it');
        if (!(node instanceof HTMLElement)) return null;
        const style = getComputedStyle(node);
        return { fontStyle: style.fontStyle, fontWeight: style.fontWeight };
      })()`);
      if (
        !boldVisualStyle ||
        boldVisualStyle.fontStyle !== 'italic' ||
        Number.parseInt(boldVisualStyle.fontWeight, 10) < 600
      ) {
        throw new Error(`Bold-italic visual style is not actually bold + italic: ${JSON.stringify(boldVisualStyle)}`);
      }
      await dispatchFormattingToggle('[data-formula-selection-bold]');
      const boldRemoved = await readFormattingValue();
      if (boldRemoved !== 'abc') {
        throw new Error(`Second bold toggle must restore ordinary math; received ${boldRemoved}`);
      }

      await setFormattingValue('xyz');
      await dispatchFormattingToggle('[data-formula-selection-italic]');
      const uprightApplied = await readFormattingValue();
      if (uprightApplied !== String.raw`\mathrm{xyz}`) {
        throw new Error(`Italic toggle must switch default math italic to \\mathrm; received ${uprightApplied}`);
      }
      await dispatchFormattingToggle('[data-formula-selection-italic]');
      const italicRestored = await readFormattingValue();
      if (italicRestored !== 'xyz') {
        throw new Error(`Second italic toggle must restore default math italic; received ${italicRestored}`);
      }

      await setFormattingValue(String.raw`\mathrm{u}`);
      await dispatchFormattingToggle('[data-formula-selection-bold]');
      const uprightBoldApplied = await readFormattingValue();
      if (uprightBoldApplied !== String.raw`\mathbf{u}`) {
        throw new Error(`Bold toggle must keep explicit upright math upright; received ${uprightBoldApplied}`);
      }

      await setFormattingValue(String.raw`\symbfit{J}`);
      const symbfitVisual = await evaluate(`(() => {
        const field = document.querySelector('math-field');
        const node = field?.shadowRoot?.querySelector('.ML__mathbfit, .ML__cmr.ML__bold.ML__it');
        if (!(node instanceof HTMLElement)) return { value: field?.value ?? null, style: null };
        const style = getComputedStyle(node);
        return {
          value: field?.value ?? null,
          style: { fontStyle: style.fontStyle, fontWeight: style.fontWeight },
        };
      })()`);
      if (
        !symbfitVisual?.value?.replace(/\s+/g, '').includes(String.raw`\symbfit{J}`) ||
        symbfitVisual.style?.fontStyle !== 'italic' ||
        Number.parseInt(symbfitVisual.style?.fontWeight ?? '0', 10) < 600
      ) {
        throw new Error(`symbfit must preserve source and render bold italic: ${JSON.stringify(symbfitVisual)}`);
      }

      await setFormattingValue(String.raw`\bm{J}`);
      const bmVisual = await evaluate(`(() => {
        const field = document.querySelector('math-field');
        const node = field?.shadowRoot?.querySelector('.ML__mathbfit, .ML__cmr.ML__bold.ML__it');
        if (!(node instanceof HTMLElement)) return { value: field?.value ?? null, style: null };
        const style = getComputedStyle(node);
        return {
          value: field?.value ?? null,
          style: { fontStyle: style.fontStyle, fontWeight: style.fontWeight },
        };
      })()`);
      if (
        !bmVisual?.value?.replace(/\s+/g, '').includes(String.raw`\bm{J}`) ||
        bmVisual.style?.fontStyle !== 'italic' ||
        Number.parseInt(bmVisual.style?.fontWeight ?? '0', 10) < 600
      ) {
        throw new Error(`bm must preserve source and render bold italic: ${JSON.stringify(bmVisual)}`);
      }

      await setFormattingValue(String.raw`\mathbf{q}`);
      await dispatchFormattingToggle('[data-formula-selection-italic]');
      const boldItalicApplied = await readFormattingValue();
      if (boldItalicApplied !== String.raw`\mathbfit{q}`) {
        throw new Error(`Italic toggle must preserve bold as \\mathbfit; received ${boldItalicApplied}`);
      }
      await dispatchFormattingToggle('[data-formula-selection-italic]');
      const boldUprightRestored = await readFormattingValue();
      if (boldUprightRestored !== String.raw`\mathbf{q}`) {
        throw new Error(`Second italic toggle must restore \\mathbf; received ${boldUprightRestored}`);
      }

      const selectionBold = {
        applied: boldApplied,
        removed: boldRemoved,
      };
      const selectionItalic = {
        upright: uprightApplied,
        restored: italicRestored,
        boldItalic: boldItalicApplied,
        boldRestored: boldUprightRestored,
      };

      if (scenario === "font-variant-formatting") {
        console.log(JSON.stringify({ selectionBold, selectionItalic }, null, 2));
        console.log("Targeted bold/italic font variant regression passed");
        return;
      }

      const removedPersistentControls = await evaluate(`(() => ({
        typingBold: Boolean(document.querySelector('[data-formula-typing-bold]')),
        typingItalic: Boolean(document.querySelector('[data-formula-typing-italic]')),
      }))()`);
      if (removedPersistentControls.typingBold || removedPersistentControls.typingItalic) {
        throw new Error(`Persistent typing controls must stay removed: ${JSON.stringify(removedPersistentControls)}`);
      }

      const exerciseCustomColor = async ({
        selector,
        color,
        storageKey,
        popoverKind,
      }) => {
        await setFormattingValue('custom');
        await dispatchFormattingToggle(selector);
        await waitForEvaluation(`(() => ({
          ready: Boolean(document.querySelector('[data-formula-color-popover="${popoverKind}"] input[type="color"]')),
        }))()`, `${popoverKind} custom color picker`);
        await evaluate(`(() => {
          const input = document.querySelector('[data-formula-color-popover="${popoverKind}"] input[type="color"]');
          if (!(input instanceof HTMLInputElement)) return false;
          const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
          setter?.call(input, '${color}');
          input.dispatchEvent(new Event('change', { bubbles: true }));
          return true;
        })()`);
        await sleep(120);
        await waitForEvaluation(`(() => ({
          ready: Boolean(document.querySelector('[data-formula-custom-color="${color}"]')),
        }))()`, `${popoverKind} saved custom color`);
        const beforeApply = await evaluate(`(() => {
          const field = document.querySelector('math-field');
          return {
            latex: field?.value ?? null,
            selection: field ? structuredClone(field.selection) : null,
            popoverOpen: Boolean(document.querySelector('[data-formula-color-popover="${popoverKind}"]')),
            visibleText: field?.shadowRoot?.querySelector('[part="content"]')?.textContent ?? '',
          };
        })()`);
        const customSwatchClicked = await evaluate(`(() => {
          const item = document.querySelector('[data-formula-custom-color="${color}"]');
          const button = item?.querySelector('.formula-color-swatch');
          if (!(button instanceof HTMLElement)) return false;
          button.click();
          return true;
        })()`);
        if (!customSwatchClicked) throw new Error(`Unable to apply custom color ${color}`);
        await sleep(120);
        const appliedState = await evaluate(`(() => {
          const field = document.querySelector('math-field');
          return {
            latex: field?.value ?? null,
            selection: field ? structuredClone(field.selection) : null,
            popoverOpen: Boolean(document.querySelector('[data-formula-color-popover="${popoverKind}"]')),
            visibleText: field?.shadowRoot?.querySelector('[part="content"]')?.textContent ?? '',
          };
        })()`);
        await dispatchFormattingToggle(selector);
        await waitForEvaluation(`(() => ({
          ready: Boolean(document.querySelector('[data-formula-custom-color="${color}"]')),
        }))()`, `${popoverKind} custom color reopened`);
        const saved = await evaluate(`(() => {
          const presets = document.querySelector('.formula-color-presets')?.getBoundingClientRect();
          const custom = document.querySelector('.formula-custom-colors-panel')?.getBoundingClientRect();
          return {
            stored: JSON.parse(localStorage.getItem('${storageKey}') || '[]'),
            customCount: document.querySelectorAll('[data-formula-custom-color]').length,
            deleteVisible: Boolean(document.querySelector('[data-delete-formula-custom-color="${color}"]')),
            customRightOfPresets: Boolean(
              presets && custom && custom.left >= presets.right - 1,
            ),
          };
        })()`);
        const deleteResult = await evaluate(`(() => {
          const button = document.querySelector('[data-delete-formula-custom-color="${color}"]');
          if (!(button instanceof HTMLElement)) return null;
          button.click();
          return true;
        })()`);
        if (!deleteResult) throw new Error(`Unable to delete custom color ${color}`);
        await sleep(80);
        const removed = await evaluate(`(() => ({
          stored: JSON.parse(localStorage.getItem('${storageKey}') || '[]'),
          stillVisible: Boolean(document.querySelector('[data-formula-custom-color="${color}"]')),
          popoverOpen: Boolean(document.querySelector('[data-formula-color-popover="${popoverKind}"]')),
        }))()`);
        return { beforeApply, appliedState, saved, removed };
      };

      const customTextColor = await exerciseCustomColor({
        selector: '[data-formula-selection-color]',
        color: '#123456',
        storageKey: 'visualtex-custom-formula-text-colors',
        popoverKind: 'color',
      });
      if (
        customTextColor.beforeApply.latex !== 'custom' ||
        !customTextColor.beforeApply.popoverOpen
      ) {
        throw new Error(`Picking a custom text color must not edit or close the formula UI: ${JSON.stringify(customTextColor)}`);
      }
      if (customTextColor.appliedState.latex !== String.raw`\textcolor{#123456}{custom}`) {
        throw new Error(`Saved custom text color did not apply as safe hex LaTeX: ${JSON.stringify(customTextColor)}`);
      }
      if (!customTextColor.saved.stored.includes('#123456')) {
        throw new Error(`Text custom color was not persisted: ${JSON.stringify(customTextColor)}`);
      }
      if (
        !customTextColor.saved.deleteVisible ||
        customTextColor.saved.customCount < 1 ||
        !customTextColor.saved.customRightOfPresets
      ) {
        throw new Error(`Text custom color controls are incomplete: ${JSON.stringify(customTextColor)}`);
      }
      if (customTextColor.removed.stored.includes('#123456') || customTextColor.removed.stillVisible) {
        throw new Error(`Text custom color was not deleted: ${JSON.stringify(customTextColor)}`);
      }
      if (!customTextColor.removed.popoverOpen) {
        throw new Error(`Deleting a custom text color must keep the popover open: ${JSON.stringify(customTextColor)}`);
      }

      const customBackgroundColor = await exerciseCustomColor({
        selector: '[data-formula-selection-background]',
        color: '#654321',
        storageKey: 'visualtex-custom-formula-background-colors',
        popoverKind: 'backgroundColor',
      });
      if (
        customBackgroundColor.beforeApply.latex !== 'custom' ||
        !customBackgroundColor.beforeApply.popoverOpen
      ) {
        throw new Error(`Picking a custom background color must not edit or close the formula UI: ${JSON.stringify(customBackgroundColor)}`);
      }
      if (customBackgroundColor.appliedState.latex !== String.raw`\colorbox{#654321}{$ custom $}`) {
        throw new Error(`Saved custom background color did not apply as safe hex LaTeX: ${JSON.stringify(customBackgroundColor)}`);
      }
      if (
        !customBackgroundColor.saved.stored.includes('#654321') ||
        !customBackgroundColor.saved.customRightOfPresets
      ) {
        throw new Error(`Background custom color was not persisted beside presets: ${JSON.stringify(customBackgroundColor)}`);
      }
      if (customBackgroundColor.removed.stored.includes('#654321') || customBackgroundColor.removed.stillVisible) {
        throw new Error(`Background custom color was not deleted: ${JSON.stringify(customBackgroundColor)}`);
      }
      if (!customBackgroundColor.removed.popoverOpen) {
        throw new Error(`Deleting a custom background color must keep the popover open: ${JSON.stringify(customBackgroundColor)}`);
      }

      console.log(JSON.stringify({
        selectionBold,
        selectionItalic,
        removedPersistentControls,
        customTextColor,
        customBackgroundColor,
      }, null, 2));
      console.log("Targeted desktop formula formatting regression passed");
      return;
    }

    if (scenario === "context-style") {
      await evaluate(`(() => {
        const storageKey = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
        persisted.state = {
          ...(persisted.state || {}),
          pngExportBackground: "#ff00ff",
        };
        localStorage.setItem(storageKey, JSON.stringify(persisted));
        location.reload();
      })()`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector("math-field")),
      }))()`, "formula field with PNG background preference");
      await evaluate(`(() => {
        window.__visualTexPngClipboardProbe = null;
        const clipboard = {
          write: async (items) => {
            const blob = await items[0].getType("image/png");
            const bitmap = await createImageBitmap(blob);
            const canvas = document.createElement("canvas");
            canvas.width = bitmap.width;
            canvas.height = bitmap.height;
            const context = canvas.getContext("2d");
            context.drawImage(bitmap, 0, 0);
            const corner = [...context.getImageData(0, 0, 1, 1).data];
            window.__visualTexPngClipboardProbe = {
              size: blob.size,
              type: blob.type,
              width: bitmap.width,
              height: bitmap.height,
              corner,
            };
            bitmap.close();
          },
        };
        Object.defineProperty(navigator, "clipboard", {
          configurable: true,
          value: clipboard,
        });
      })()`);

      await focusField();
      await clearField();
      await typeText("abc");

      const beforeState = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.selection = {
          ranges: [[0, field.lastOffset]],
          direction: "forward",
        };
        const bounds = field.shadowRoot
          ?.querySelector('[part="content"]')
          ?.getBoundingClientRect();
        return bounds
          ? {
              x: bounds.left + bounds.width / 2,
              y: bounds.top + bounds.height / 2,
              value: field.value,
              selection: field.selection,
              menuItemCount: field.menuItems.length,
            }
          : null;
      })()`);
      if (!beforeState) throw new Error("Unable to resolve formula bounds");
      assert.equal(beforeState.menuItemCount, 0, "MathLive menuItems must be empty");

      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: beforeState.x,
        y: beforeState.y,
        button: "right",
        buttons: 2,
        clickCount: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: beforeState.x,
        y: beforeState.y,
        button: "right",
        buttons: 0,
        clickCount: 1,
      });

      const menuState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const mathLiveMenus = [...(field?.shadowRoot?.querySelectorAll("menu.ui-menu-container") ?? [])];
        const visibleMathLiveMenus = mathLiveMenus.filter((menu) => {
          const style = getComputedStyle(menu);
          return style.display !== "none" && style.visibility !== "hidden";
        });
        const menu = document.querySelector(".formula-editor-context-menu");
        const button = menu?.querySelector("button");
        const bounds = menu?.getBoundingClientRect();
        return {
          ready: Boolean(menu && button && bounds),
          value: field?.value ?? "",
          selection: field?.selection ?? null,
          menuItemCount: field?.menuItems.length ?? -1,
          visibleMathLiveMenuCount: visibleMathLiveMenus.length,
          width: bounds?.width ?? 0,
          height: bounds?.height ?? 0,
          text: button?.textContent?.trim() ?? "",
        };
      })()`, "VisualTeX formula context menu");
      assert.equal(menuState.menuItemCount, 0);
      assert.equal(menuState.visibleMathLiveMenuCount, 0);
      assert.equal(menuState.value, beforeState.value);
      assert.ok(menuState.width >= 160 && menuState.width <= 205, JSON.stringify(menuState));
      assert.ok(menuState.height <= 52, JSON.stringify(menuState));
      assert.match(menuState.text, /PNG/);

      const themeMenuStyles = await evaluate(`(() => {
        const root = document.documentElement;
        const menu = document.querySelector(".formula-editor-context-menu");
        const button = menu?.querySelector("button");
        const original = root.dataset.theme || "light";
        const result = {};
        for (const theme of ["light", "beige", "dark", "purple", "green"]) {
          root.dataset.theme = theme;
          const menuStyle = getComputedStyle(menu);
          const buttonStyle = getComputedStyle(button);
          result[theme] = {
            background: menuStyle.backgroundColor,
            border: menuStyle.borderColor,
            text: buttonStyle.color,
          };
        }
        root.dataset.theme = original;
        return result;
      })()`);
      assert.ok(
        new Set(Object.values(themeMenuStyles).map((style) => style.background)).size >= 4,
        JSON.stringify(themeMenuStyles),
      );

      await evaluate(`document.querySelector(".formula-editor-context-menu button")?.click()`);
      const clipboardState = await waitForEvaluation(`(() => ({
        ready: Boolean(window.__visualTexPngClipboardProbe?.size),
        ...(window.__visualTexPngClipboardProbe || {}),
      }))()`, "PNG clipboard rasterization");
      assert.equal(clipboardState.type, "image/png");
      assert.ok(clipboardState.width > 0 && clipboardState.height > 0);
      assert.deepEqual(clipboardState.corner, [255, 0, 255, 255]);
      assert.equal(
        await evaluate(`Boolean(document.querySelector(".formula-editor-context-menu"))`),
        false,
      );

      await evaluate(`document.querySelector(".workspace-export-trigger")?.click()`);
      const exportCopyLayout = await waitForEvaluation(`(() => {
        const cards = [...document.querySelectorAll(".export-format-option")];
        const png = cards.find((card) => card.querySelector("strong")?.textContent === "PNG");
        const copy = document.querySelector("[data-copy-png-from-export]");
        const pngBounds = png?.getBoundingClientRect();
        const copyBounds = copy?.getBoundingClientRect();
        return {
          ready: Boolean(pngBounds && copyBounds),
          pngBottom: pngBounds?.bottom ?? 0,
          copyTop: copyBounds?.top ?? 0,
          leftDelta: pngBounds && copyBounds ? Math.abs(pngBounds.left - copyBounds.left) : 999,
          widthDelta: pngBounds && copyBounds ? Math.abs(pngBounds.width - copyBounds.width) : 999,
        };
      })()`, "export PNG clipboard action");
      assert.ok(exportCopyLayout.copyTop >= exportCopyLayout.pngBottom, JSON.stringify(exportCopyLayout));
      assert.ok(exportCopyLayout.leftDelta <= 1, JSON.stringify(exportCopyLayout));
      assert.ok(exportCopyLayout.widthDelta <= 1, JSON.stringify(exportCopyLayout));
      await evaluate(`document.querySelector(".export-dialog .icon-button")?.click()`);

      await evaluate(`document.querySelector(".settings-toggle")?.click()`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector("[data-interface-customization-trigger]")),
      }))()`, "settings dialog");
      await evaluate(`document.querySelector("[data-interface-customization-trigger]")?.click()`);
      const backgroundSetting = await waitForEvaluation(`(() => {
        const color = document.querySelector("[data-png-background-color-setting]");
        const transparent = document.querySelector("[data-png-background-transparent]");
        return {
          ready: Boolean(color && transparent),
          color: color?.value ?? "",
          transparentPressed: transparent?.getAttribute("aria-pressed") ?? "",
        };
      })()`, "PNG background customization");
      assert.equal(backgroundSetting.color, "#ff00ff");
      assert.equal(backgroundSetting.transparentPressed, "false");
      await evaluate(`document.querySelector("[data-png-background-transparent]")?.click()`);
      const transparentState = await waitForEvaluation(`(() => {
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        const button = document.querySelector("[data-png-background-transparent]");
        const value = persisted.state?.pngExportBackground;
        return {
          ready: value === "transparent" && button?.getAttribute("aria-pressed") === "true",
          value,
        };
      })()`, "transparent PNG background persistence");

      console.log(JSON.stringify({
        beforeState,
        menuState,
        themeMenuStyles,
        clipboardState,
        exportCopyLayout,
        backgroundSetting,
        transparentState,
      }, null, 2));
      console.log("Targeted VisualTeX PNG context-menu regression passed");
      return;
    }

    if (scenario === "limit-candidate") {
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector(".canvas-input-behavior-trigger")),
      }))()`, "input behavior trigger for plain lim candidate");
      await evaluate(`document.querySelector(".canvas-input-behavior-trigger")?.click()`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector(".input-behavior-popover")),
      }))()`, "input behavior menu for plain lim candidate");
      await evaluate(`(() => {
        const option = [...document.querySelectorAll(".input-behavior-option")].find((label) => {
          const title = label.querySelector("strong")?.textContent ?? "";
          return title.includes("快捷转义") || title.includes("Common math shortcuts");
        });
        const checkbox = option?.querySelector('input[type="checkbox"]');
        if (checkbox?.checked) checkbox.click();
      })()`);
      await evaluate(`document.querySelector(".canvas-input-behavior-trigger")?.click()`);
      await focusField();
      await clearField();
      await typeText("sin");
      await waitForEvaluation(`(() => ({
        ready: document.querySelector("math-field")?.value === "\\\\sin",
      }))()`, "plain function name becomes an upright command");
      await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.setValue("sin", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.position = field.lastOffset;
        field.selection = {
          ranges: [[field.lastOffset, field.lastOffset]],
          direction: "none",
        };
        const zoomButton = [...document.querySelectorAll(".canvas-controls button")]
          .find((button) => !button.disabled);
        zoomButton?.click();
      })()`);
      const semanticRepairState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        return {
          ready:
            field?.value === "\\\\sin" &&
            persisted.state?.lines?.[0]?.latex === "\\\\sin" &&
            field.selection?.ranges?.[0]?.[0] === field.lastOffset &&
            field.selection?.ranges?.[0]?.[1] === field.lastOffset,
          value: field?.value ?? "",
          storeValue: persisted.state?.lines?.[0]?.latex ?? "",
          selection: field?.selection ?? null,
        };
      })()`, "React synchronization repairs a stale MathLive semantic model");

      await clearField();
      await typeText("lim");
      await waitForEvaluation(`(() => {
        const popup = document.querySelector(".suggestion-popup");
        const selected = popup?.querySelector(
          ".suggestion-item.is-selected .suggestion-command",
        )?.textContent?.trim() ?? "";
        return {
          ready: Boolean(popup) && selected === "\\\\lim",
          selected,
          value: document.querySelector("math-field")?.value ?? "",
        };
      })()`, "plain lim opens the structured limit candidate");
      const staleCandidateState = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.setValue("lim", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.position = field.lastOffset;
        field.selection = {
          ranges: [[field.lastOffset, field.lastOffset]],
          direction: "none",
        };
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        return {
          value: field.value,
          popupVisible: Boolean(document.querySelector(".suggestion-popup")),
        };
      })()`);
      await key("Enter", "Enter", 13);
      const result = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const value = field?.value ?? "";
        const selection = field?.selection ?? null;
        return {
          ready:
            value === "\\\\lim_{\\\\placeholder{}\\\\to\\\\placeholder{}} \\\\placeholder{}" &&
            document.querySelectorAll("math-field").length === 1 &&
            selection?.ranges?.[0]?.[0] !== selection?.ranges?.[0]?.[1],
          value,
          selection,
          lineCount: document.querySelectorAll("math-field").length,
          customPopupVisible: Boolean(document.querySelector(".suggestion-popup")),
          nativePopupVisible:
            document.getElementById("mathlive-suggestion-popover")?.classList.contains("is-visible") ?? false,
        };
      })()`, "one Enter accepts the first structured limit candidate");

      await key("z", "KeyZ", 90, 4);
      const undoState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        return {
          ready:
            field?.value === "\\\\lim" &&
            persisted.state?.lines?.[0]?.latex === "\\\\lim" &&
            document.querySelectorAll("math-field").length === 1,
          value: field?.value ?? "",
          storeValue: persisted.state?.lines?.[0]?.latex ?? "",
        };
      })()`, "undo restores the confirmed bare limit command");
      await key("z", "KeyZ", 90, 12);
      const redoState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const value = field?.value ?? "";
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        return {
          ready:
            value === "\\\\lim_{\\\\placeholder{}\\\\to\\\\placeholder{}} \\\\placeholder{}" &&
            persisted.state?.lines?.[0]?.latex === value &&
            document.querySelectorAll("math-field").length === 1,
          value,
          storeValue: persisted.state?.lines?.[0]?.latex ?? "",
          selection: field?.selection ?? null,
        };
      })()`, "redo reapplies the complete limit structure");

      console.log(
        JSON.stringify(
          {
            semanticRepairState,
            staleCandidateState,
            result,
            undoState,
            redoState,
          },
          null,
          2,
        ),
      );
      console.log("Targeted limit candidate Enter regression passed");
      return;
    }

    if (scenario === "upright") {
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector(".canvas-input-behavior-trigger")),
      }))()`, "input behavior trigger for upright detection");
      await evaluate(`document.querySelector(".canvas-input-behavior-trigger")?.click()`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector(".input-behavior-popover")),
      }))()`, "input behavior menu for upright detection");
      await evaluate(`(() => {
        const option = [...document.querySelectorAll(".input-behavior-option")].find((label) => {
          const title = label.querySelector("strong")?.textContent ?? "";
          return title.includes("快捷转义") || title.includes("Common math shortcuts");
        });
        const checkbox = option?.querySelector('input[type="checkbox"]');
        if (checkbox?.checked) checkbox.click();
      })()`);
      await waitForEvaluation(`(() => {
        const option = [...document.querySelectorAll(".input-behavior-option")].find((label) => {
          const title = label.querySelector("strong")?.textContent ?? "";
          return title.includes("快捷转义") || title.includes("Common math shortcuts");
        });
        return {
          ready: option?.querySelector('input[type="checkbox"]')?.checked === false,
        };
      })()`, "common shortcut escaping disabled");
      await evaluate(`document.querySelector(".canvas-input-behavior-trigger")?.click()`);
      await waitForEvaluation(`(() => ({
        ready: !document.querySelector(".input-behavior-popover"),
      }))()`, "input behavior menu closed before upright input");
      await focusField();
      await clearField();
      await typeText("driver");
      const identifierState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        return {
          ready: field?.value === "driver",
          value: field?.value ?? "",
        };
      })()`, "ordinary identifier remains italic variables");

      await clearField();
      await typeText("dx/dy");
      const cartesianDifferentialState = await waitForEvaluation(`(() => {
        const value = document.querySelector("math-field")?.value ?? "";
        return {
          ready: value === "\\\\frac{\\\\mathrm{d}x}{\\\\mathrm{d}y}",
          value,
        };
      })()`, "dx/dy keeps only each differential operator upright");

      await clearField();
      await typeText("d^2y/dx^2");
      const higherDerivativeState = await waitForEvaluation(`(() => {
        const value = document.querySelector("math-field")?.value ?? "";
        return {
          ready:
            value ===
            "\\\\frac{\\\\mathrm{d}^2y}{\\\\mathrm{d}x^2}",
          value,
        };
      })()`, "higher derivative keeps d upright and variables italic");

      await clearField();
      await typeText("\\int");
      await waitForEvaluation(`(() => {
        const panel = document.getElementById("mathlive-suggestion-popover");
        return {
          ready:
            panel?.classList.contains("is-visible") &&
            panel.querySelector("li.ML__popover__current")?.dataset.command === "\\\\int",
        };
      })()`, "integral native command awaits its first confirmation");
      await key(" ", "Space", 32);
      await waitForEvaluation(`(() => ({
        ready:
          document.querySelector("math-field")?.value === "\\\\int" &&
          Boolean(document.querySelector(".suggestion-popup")),
      }))()`, "first integral confirmation preserves the bare command");
      await typeText("f(x)dx");
      const integralMeasureState = await waitForEvaluation(`(() => {
        const value = document.querySelector("math-field")?.value ?? "";
        return {
          ready:
            value.includes("\\\\int") &&
            value.includes("f") &&
            value.includes("x") &&
            value.endsWith("\\\\mathrm{d}x"),
          value,
        };
      })()`, "integral differential measure becomes upright");

      const functionStates = {};
      for (const functionName of ["sin", "cos", "exp"]) {
        await clearField();
        await typeText(functionName);
        functionStates[functionName] = await waitForEvaluation(`(() => {
          const value = document.querySelector("math-field")?.value ?? "";
          return {
            ready: value === ${JSON.stringify("\\" + functionName)},
            value,
          };
        })()`, `plain ${functionName} becomes an upright function command`);
      }

      await clearField();
      await typeText("dr/d");
      await typeText("\\theta");
      await key(" ", "Space", 32);
      await sleep(450);
      const differentialState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const value = field?.value ?? "";
        const uprightCount =
          (value.match(/\\\\differentialD|\\\\mathrm\\{d\\}/g) ?? []).length;
        return {
          ready:
            uprightCount === 2 &&
            /\\\\theta/.test(value),
          value,
          uprightCount,
        };
      })()`, "slash derivative uses two upright differential operators");

      console.log(
        JSON.stringify(
          {
            identifierState,
            cartesianDifferentialState,
            higherDerivativeState,
            integralMeasureState,
            functionStates,
            differentialState,
          },
          null,
          2,
        ),
      );
      console.log("Targeted contextual upright differential regression passed");
      return;
    }

    if (scenario === "suggestions") {
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector(".canvas-input-behavior-trigger")),
      }))()`, "input behavior trigger for other-command suggestions");
      await evaluate(`document.querySelector(".canvas-input-behavior-trigger")?.click()`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector(".input-behavior-popover")),
      }))()`, "input behavior menu for other-command suggestions");
      await evaluate(`(() => {
        const option = [...document.querySelectorAll(".input-behavior-option")].find((label) => {
          const title = label.querySelector("strong")?.textContent ?? "";
          return title.includes("其他命令") || title.includes("Other command suggestions");
        });
        const checkbox = option?.querySelector('input[type="checkbox"]');
        if (checkbox && !checkbox.checked) checkbox.click();
      })()`);
      await waitForEvaluation(`(() => {
        const option = [...document.querySelectorAll(".input-behavior-option")].find((label) => {
          const title = label.querySelector("strong")?.textContent ?? "";
          return title.includes("其他命令") || title.includes("Other command suggestions");
        });
        return {
          ready: option?.querySelector('input[type="checkbox"]')?.checked === true,
        };
      })()`, "other-command suggestion setting enabled");
      await evaluate(`document.querySelector(".canvas-input-behavior-trigger")?.click()`);
      await waitForEvaluation(`(() => ({
        ready: !document.querySelector(".input-behavior-popover"),
      }))()`, "input behavior menu closed before alpha input");
      await focusField();
      await clearField();
      await typeText("alpha");

      const openState = await waitForEvaluation(`(() => {
        const popup = document.querySelector(".suggestion-popup");
        const selected = popup?.querySelector(".suggestion-item.is-selected .suggestion-command");
        const field = document.querySelector("math-field");
        const sink = field?.shadowRoot?.querySelector('[part="keyboard-sink"]');
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        const nativePanel = document.getElementById("mathlive-suggestion-popover");
        return {
          ready: Boolean(popup) && selected?.textContent?.trim() === "\\\\alpha",
          value: field?.value ?? "",
          rawSinkValue: sink?.value ?? "",
          rawSinkText: sink?.textContent ?? "",
          fieldFocused: field?.hasFocus?.() ?? false,
          activeTag: document.activeElement?.tagName ?? "",
          selected: selected?.textContent?.trim() ?? "",
          nativeVisible: nativePanel?.classList.contains("is-visible") ?? false,
          showOther:
            persisted.state?.inputBehavior?.showOtherCommandSuggestions ?? null,
        };
      })()`, "other-command suggestion opens for alpha");

      await key("Enter", "Enter", 13);
      const confirmedState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const nativePanel = document.getElementById("mathlive-suggestion-popover");
        const nativeVisible = nativePanel?.classList.contains("is-visible") ?? false;
        return {
          ready:
            field?.value === "\\\\alpha" &&
            !document.querySelector(".suggestion-popup") &&
            !nativeVisible,
          value: field?.value ?? "",
          popupVisible: Boolean(document.querySelector(".suggestion-popup")),
          nativeVisible,
          popoverPolicy: field?.popoverPolicy ?? "",
          lineCount: document.querySelectorAll("math-field").length,
        };
      })()`, "Enter confirms alpha and dismisses both suggestion panels");

      await key("x", "KeyX", 88);
      const continuedState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const nativePanel = document.getElementById("mathlive-suggestion-popover");
        const nativeVisible = nativePanel?.classList.contains("is-visible") ?? false;
        return {
          ready:
            field?.value === "\\\\alpha x" &&
            !document.querySelector(".suggestion-popup") &&
            !nativeVisible,
          value: field?.value ?? "",
          popupVisible: Boolean(document.querySelector(".suggestion-popup")),
          nativeVisible,
        };
      })()`, "typing after alpha confirmation does not restore either old panel");

      await clearField();
      await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.setValue("\\\\theta", {
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
          data: "a",
        }));
      })()`);
      const navigationOpenState = await waitForEvaluation(`(() => {
        const commands = [...document.querySelectorAll(
          ".suggestion-item .suggestion-command",
        )].map((node) => node.textContent?.trim() ?? "");
        const selected = document.querySelector(
          ".suggestion-item.is-selected .suggestion-command",
        )?.textContent?.trim() ?? "";
        return {
          ready:
            commands.length >= 3 &&
            commands[0] === "\\\\theta" &&
            commands[1] === "\\\\Theta" &&
            commands[2] === "\\\\vartheta" &&
            selected === "\\\\theta",
          commands,
          selected,
        };
      })()`, "other-command candidate list opens with multiple theta variants");

      await evaluate(`(() => {
        const standardToggle = document.querySelector(".source-toggle");
        const classicToggle = document.querySelector('[data-classic-bottom-view="source"]');
        (standardToggle || classicToggle)?.click();
      })()`);
      await waitForEvaluation(`(() => ({
        ready:
          Boolean(document.querySelector(".source-panel")) &&
          Boolean(document.querySelector(".suggestion-popup")),
      }))()`, "source pane opens below command candidates");
      const layerState = await waitForEvaluation(`(() => {
        const popup = document.querySelector(".suggestion-popup");
        const editorScroll = document.querySelector(".editor-pane-scroll");
        const sourcePanel = document.querySelector(".source-panel");
        if (!popup || !editorScroll || !sourcePanel) return { ready: false };
        const editorBounds = editorScroll.getBoundingClientRect();
        popup.style.top = (editorBounds.bottom - 60) + "px";
        const popupBounds = popup.getBoundingClientRect();
        const sourceBounds = sourcePanel.getBoundingClientRect();
        const testX = popupBounds.left + popupBounds.width / 2;
        const testY = Math.min(
          popupBounds.bottom - 2,
          Math.max(sourceBounds.top + 2, editorBounds.bottom + 18),
        );
        const topmostNode = document.elementFromPoint(testX, testY);
        const style = getComputedStyle(popup);
        return {
          ready:
            popup.parentElement === document.body &&
            style.position === "fixed" &&
            Number.parseInt(style.zIndex || "0", 10) >= 300 &&
            popupBounds.bottom > sourceBounds.top &&
            Boolean(topmostNode && popup.contains(topmostNode)),
          parentIsBody: popup.parentElement === document.body,
          position: style.position,
          zIndex: style.zIndex,
          popupBottom: popupBounds.bottom,
          sourceTop: sourceBounds.top,
          topmostClass: topmostNode?.className ?? "",
        };
      })()`, "VisualTeX command candidate stays above the source pane");
      await evaluate(`(() => {
        const standardCollapse = document.querySelector(".source-collapse-button");
        const classicTools = document.querySelector('[data-classic-bottom-view="tools"]');
        (standardCollapse || classicTools)?.click();
      })()`);
      await waitForEvaluation(`(() => ({
        ready:
          !document.querySelector(".source-panel") &&
          Boolean(document.querySelector(".suggestion-popup")),
      }))()`, "source pane closes without dismissing command candidates");

      await key("ArrowDown", "ArrowDown", 40);
      const firstNavigationState = await waitForEvaluation(`(() => {
        const selected = document.querySelector(
          ".suggestion-item.is-selected .suggestion-command",
        )?.textContent?.trim() ?? "";
        return {
          ready: selected === "\\\\Theta",
          selected,
        };
      })()`, "ArrowDown keeps the second other-command candidate selected");
      await sleep(350);
      const stableNavigationState = await waitForEvaluation(`(() => {
        const selected = document.querySelector(
          ".suggestion-item.is-selected .suggestion-command",
        )?.textContent?.trim() ?? "";
        return {
          ready: selected === "\\\\Theta",
          selected,
        };
      })()`, "candidate refresh does not reset selection to the first row");

      await key("ArrowDown", "ArrowDown", 40);
      const secondNavigationState = await waitForEvaluation(`(() => {
        const selected = document.querySelector(
          ".suggestion-item.is-selected .suggestion-command",
        )?.textContent?.trim() ?? "";
        return {
          ready: selected === "\\\\vartheta",
          selected,
        };
      })()`, "second ArrowDown selects the third other-command candidate");
      await key("Enter", "Enter", 13);
      const navigationCommitState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        return {
          ready:
            field?.value === "\\\\vartheta" &&
            !document.querySelector(".suggestion-popup"),
          value: field?.value ?? "",
          popupVisible: Boolean(document.querySelector(".suggestion-popup")),
        };
      })()`, "Enter commits the currently highlighted other-command candidate");

      console.log(JSON.stringify({
        openState,
        confirmedState,
        continuedState,
        navigationOpenState,
        firstNavigationState,
        stableNavigationState,
        secondNavigationState,
        navigationCommitState,
        layerState,
      }, null, 2));
      console.log("Targeted other-command suggestion dismissal and navigation regression passed");
      return;
    }

    if (scenario === "navigation") {
      await evaluate(`(() => {
        const storageKey = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
        const first = { id: crypto.randomUUID(), latex: "\\\\alpha" };
        const second = { id: crypto.randomUUID(), latex: "\\\\beta" };
        persisted.state = {
          ...(persisted.state || {}),
          lines: [first, second],
          activeLineId: second.id,
          inputBehavior: {
            ...(persisted.state?.inputBehavior || {}),
            showOtherCommandSuggestions: true,
          },
        };
        localStorage.setItem(storageKey, JSON.stringify(persisted));
        location.reload();
      })()`);
      await waitForEvaluation(`(() => ({
        ready: document.querySelectorAll("math-field").length === 2,
      }))()`, "two formula fields for navigation");
      await evaluate(`(() => {
        const field = document.querySelectorAll("math-field")[1];
        field.position = field.lastOffset;
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      })()`);
      await key("ArrowUp", "ArrowUp", 38);
      const switchedState = await waitForEvaluation(`(() => {
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
          candidateVisible: Boolean(document.querySelector(".suggestion-popup")),
          query: document.querySelector(".suggestion-popup")?.textContent ?? "",
          activeLineId: surface?.dataset.activeLineId ?? "",
          firstLineId,
          focusedIndex: fields.findIndex((field) => field.matches(":focus-within")),
        };
      })()`, "ArrowUp switches to first formula field");

      await key("Escape", "Escape", 27);
      const dismissedState = await waitForEvaluation(`(() => {
        const popup = document.querySelector(".suggestion-popup");
        const field = document.querySelectorAll("math-field")[0];
        return {
          ready: !popup,
          popupText: popup?.textContent ?? "",
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
          pendingNativeSuggestion: field?.dataset.pendingNativeSuggestion ?? "",
          activeTag: document.activeElement?.tagName ?? "",
          sinkFocused: field?.shadowRoot?.activeElement?.getAttribute?.("part") ?? "",
        };
      })()`, "Escape dismisses formula-line command candidate");
      await sleep(500);
      const stableDismissedState = await waitForEvaluation(`(() => ({
        ready: !document.querySelector(".suggestion-popup"),
        popupText: document.querySelector(".suggestion-popup")?.textContent ?? "",
      }))()`, "dismissed formula-line command candidate stays closed");

      await key("ArrowDown", "ArrowDown", 40);
      const returnedState = await waitForEvaluation(`(() => {
        const rows = [...document.querySelectorAll(".formula-line")];
        const fields = [...document.querySelectorAll("math-field")];
        const surface = document.querySelector(".multi-line-editor");
        const secondLineId = rows[1]?.dataset.lineId ?? "";
        return {
          ready:
            rows[1]?.classList.contains("is-active") &&
            fields[1]?.matches(":focus-within") &&
            surface?.dataset.activeLineId === secondLineId,
          candidateVisible: Boolean(document.querySelector(".suggestion-popup")),
          activeLineId: surface?.dataset.activeLineId ?? "",
          secondLineId,
          focusedIndex: fields.findIndex((field) => field.matches(":focus-within")),
        };
      })()`, "ArrowDown returns to second formula field");

      console.log(JSON.stringify({ switchedState, dismissedState, stableDismissedState, returnedState }, null, 2));
      console.log("Targeted formula-line navigation regression passed");
      return;
    }

    if (scenario === "font-settings") {
      await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        if (!field?.isConnected) return { ready: false };
        field.setValue("x+\\\\alpha+\\\\mathrm{A}+\\\\text{中文}+\\\\sqrt{y}+\\\\sum_{i=1}^{n}i", {
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
        return { ready: field.isConnected && field.value.includes("中文") };
      })()`, "mixed formula font fixture");
      const sourceBefore = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        return {
          ready: Boolean(field?.value?.includes("中文") && persisted.state?.lines?.[0]?.latex?.includes("中文")),
          value: field?.value ?? "",
          stored: persisted.state?.lines?.[0]?.latex ?? "",
        };
      })()`, "mixed formula for font settings");

      await evaluate(`document.querySelector(".settings-toggle")?.click()`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector("[data-interface-customization-trigger]")),
      }))()`, "settings dialog for formula fonts");
      await evaluate(`document.querySelector("[data-interface-customization-trigger]")?.click()`);
      await waitForEvaluation(`(() => ({
        ready:
          Boolean(document.querySelector("[data-formula-letter-font-setting]")) &&
          Boolean(document.querySelector("[data-formula-chinese-font-setting]")),
      }))()`, "formula font controls");

      await evaluate(`(() => {
        const letter = document.querySelector("[data-formula-letter-font-setting]");
        const chinese = document.querySelector("[data-formula-chinese-font-setting]");
        letter.value = "palatino";
        letter.dispatchEvent(new Event("change", { bubbles: true }));
        chinese.value = "songti";
        chinese.dispatchEvent(new Event("change", { bubbles: true }));
      })()`);

      const customState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const shadow = field?.shadowRoot;
        const italic = shadow?.querySelector(".ML__mathit");
        const upright = shadow?.querySelector(".ML__cmr");
        const chinese = shadow?.querySelector(".ML__text");
        const largeOperator = shadow?.querySelector(".ML__op-symbol.ML__large-op");
        const preview = document.querySelector("[data-formula-font-preview]");
        const previewItalic = preview?.querySelector(".ML__mathit");
        const previewChinese = preview?.querySelector(".ML__text");
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        const result = {
          ready:
            persisted.state?.formulaLetterFont === "palatino" &&
            persisted.state?.formulaChineseFont === "songti" &&
            Boolean(italic && upright && chinese && largeOperator && previewItalic && previewChinese),
          letterSetting: persisted.state?.formulaLetterFont ?? "",
          chineseSetting: persisted.state?.formulaChineseFont ?? "",
          italicFont: italic ? getComputedStyle(italic).fontFamily : "",
          uprightFont: upright ? getComputedStyle(upright).fontFamily : "",
          chineseFont: chinese ? getComputedStyle(chinese).fontFamily : "",
          operatorFont: largeOperator ? getComputedStyle(largeOperator).fontFamily : "",
          previewItalicFont: previewItalic ? getComputedStyle(previewItalic).fontFamily : "",
          previewChineseFont: previewChinese ? getComputedStyle(previewChinese).fontFamily : "",
          value: field?.value ?? "",
        };
        result.ready =
          result.ready &&
          result.italicFont.includes("Palatino") &&
          result.uprightFont.includes("Palatino") &&
          result.chineseFont.includes("Songti SC") &&
          result.previewItalicFont.includes("Palatino") &&
          result.previewChineseFont.includes("Songti SC") &&
          result.operatorFont.includes("KaTeX_Size");
        return result;
      })()`, "custom formula fonts applied");
      assert.equal(customState.value, sourceBefore.value, "font changes must not mutate LaTeX");
      assert.ok(customState.operatorFont.includes("KaTeX_Size"), JSON.stringify(customState));

      await evaluate(`(() => {
        const chinese = document.querySelector("[data-formula-chinese-font-setting]");
        chinese.value = "kaiti";
        chinese.dispatchEvent(new Event("change", { bubbles: true }));
      })()`);
      await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.setValue("中文+x", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.dispatchEvent(new InputEvent("input", {
          bubbles: true,
          composed: true,
          inputType: "insertText",
          data: "文",
        }));
      })()`);
      const kaitiState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const nodes = field?.shadowRoot
          ? [...field.shadowRoot.querySelectorAll("*")].map((node) => ({
              className: typeof node.className === "string" ? node.className : "",
              text: node.textContent || "",
              font: getComputedStyle(node).fontFamily,
            }))
          : [];
        const chineseGlyphs = nodes.filter(
          (item) => item.text === "中" || item.text === "文" || item.text === "中文",
        );
        const latinGlyph = nodes.find((item) => item.text === "x") ?? null;
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        const globalLetterFont = localStorage.getItem("visualtex.formula-letter-font");
        const globalChineseFont = localStorage.getItem("visualtex.formula-chinese-font");
        return {
          ready:
            persisted.state?.formulaChineseFont === "kaiti" &&
            globalLetterFont === "palatino" &&
            globalChineseFont === "kaiti" &&
            chineseGlyphs.length > 0 &&
            chineseGlyphs.every((item) => item.font.includes("Kaiti SC")) &&
            Boolean(latinGlyph?.font.includes("Palatino")),
          value: field?.value ?? "",
          globalLetterFont,
          globalChineseFont,
          chineseGlyphs,
          latinGlyph,
          nodes: nodes.filter((item) => item.text.includes("中") || item.text.includes("文")),
          fontAvailable: document.fonts?.check?.('16px "Kaiti SC"', "中文") ?? null,
        };
      })()`, "Kaiti formula font applied to rendered Chinese glyphs");

      const chineseFontSwitches = [];
      for (const [id, family] of [
        ["pingfang", "PingFang SC"],
        ["songti", "Songti SC"],
        ["heiti", "Heiti SC"],
        ["kaiti", "Kaiti SC"],
      ]) {
        await evaluate(`(() => {
          const chinese = document.querySelector("[data-formula-chinese-font-setting]");
          chinese.value = ${JSON.stringify(id)};
          chinese.dispatchEvent(new Event("change", { bubbles: true }));
        })()`);
        chineseFontSwitches.push(
          await waitForEvaluation(`(() => {
            const field = document.querySelector("math-field");
            const glyphs = field?.shadowRoot
              ? [...field.shadowRoot.querySelectorAll(".visualtex-chinese-glyph")]
                  .filter((node) => node.textContent === "中" || node.textContent === "文")
                  .map((node) => ({ text: node.textContent, font: getComputedStyle(node).fontFamily }))
              : [];
            const selected = document.querySelector("[data-formula-chinese-font-setting]")?.value ?? "";
            return {
              ready:
                selected === ${JSON.stringify(id)} &&
                glyphs.length >= 2 &&
                glyphs.every((item) => item.font.includes(${JSON.stringify(family)})),
              selected,
              glyphs,
            };
          })()`, "normal-input Chinese font switch"),
        );
      }

      const beforeResetValue = await evaluate(`document.querySelector("math-field")?.value ?? ""`);
      await evaluate(`document.querySelector("[data-formula-font-reset]")?.click()`);
      const resetState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const shadow = field?.shadowRoot;
        const italic = shadow?.querySelector(".ML__mathit");
        const chinese = shadow?.querySelector(".visualtex-chinese-glyph, .ML__text");
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        const italicFont = italic ? getComputedStyle(italic).fontFamily : "";
        const chineseFont = chinese ? getComputedStyle(chinese).fontFamily : "";
        return {
          ready:
            persisted.state?.formulaLetterFont === "katex" &&
            persisted.state?.formulaChineseFont === "system" &&
            italicFont.includes("KaTeX_Math") &&
            chineseFont.includes("PingFang SC"),
          letterSetting: persisted.state?.formulaLetterFont ?? "",
          chineseSetting: persisted.state?.formulaChineseFont ?? "",
          italicFont,
          chineseFont,
          value: field?.value ?? "",
        };
      })()`, "formula font defaults restored");
      assert.equal(resetState.value, beforeResetValue, "font reset must not mutate LaTeX");

      console.log(JSON.stringify({ sourceBefore, customState, kaitiState, chineseFontSwitches, resetState }, null, 2));
      console.log("Targeted formula font settings regression passed");
      return;
    }

    if (scenario === "output-fonts") {
      await evaluate(`(() => {
        const storageKey = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
        persisted.state = {
          ...(persisted.state || {}),
          formulaLetterFont: "palatino",
          formulaChineseFont: "songti",
          pngExportBackground: "transparent",
        };
        localStorage.setItem(storageKey, JSON.stringify(persisted));
        location.reload();
      })()`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector("math-field")),
      }))()`, "formula field for output font regression");

      await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        if (!field?.isConnected) return { ready: false };
        field.setValue("x+\\\\alpha+\\\\mathrm{A}+\\\\text{中文}+\\\\sqrt{y}+\\\\sum_{i=1}^{n}i+\\\\mathbb{R}+\\\\mathcal{L}+\\\\mathsf{S}", {
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
        return { ready: field.value.includes("中文") };
      })()`, "mixed output font fixture");

      await evaluate(`(() => {
        window.__visualtexOutputFontProbe = { svg: "", png: null };
        const originalCreateObjectURL = URL.createObjectURL.bind(URL);
        URL.createObjectURL = (blob) => {
          if (blob?.type?.startsWith("image/svg+xml")) {
            blob.text().then((text) => {
              window.__visualtexOutputFontProbe.svg = text;
            });
          }
          return originalCreateObjectURL(blob);
        };
        HTMLAnchorElement.prototype.click = function () {};
        const clipboard = {
          write: async (items) => {
            const blob = await items[0].getType("image/png");
            const bitmap = await createImageBitmap(blob);
            window.__visualtexOutputFontProbe.png = {
              type: blob.type,
              size: blob.size,
              width: bitmap.width,
              height: bitmap.height,
            };
            bitmap.close();
          },
        };
        Object.defineProperty(navigator, "clipboard", {
          configurable: true,
          value: clipboard,
        });
      })()`);

      await evaluate(`document.querySelector(".workspace-export-trigger")?.click()`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector(".export-dialog")),
      }))()`, "export dialog for output fonts");
      await evaluate(`(() => {
        const cards = [...document.querySelectorAll(".export-format-option")];
        cards.find((card) => card.querySelector("strong")?.textContent === "SVG")?.click();
      })()`);
      await evaluate(`document.querySelector(".export-dialog .primary-button")?.click()`);
      const svgState = await waitForEvaluation(`(() => {
        const svg = window.__visualtexOutputFontProbe?.svg || "";
        return {
          ready: svg.includes("data-visualtex-output-letter-font=\\\"palatino\\\"") && svg.includes("data-visualtex-output-text-font=\\\"songti\\\""),
          hasPalatino: svg.includes("Palatino"),
          hasSongti: svg.includes("Songti SC"),
          hasLetterMarker: svg.includes("data-visualtex-output-letter-font=\\\"palatino\\\""),
          hasChineseMarker: svg.includes("data-visualtex-output-text-font=\\\"songti\\\""),
          keepsSumPath: /xlink:href=\\\"#MJX-[^\\\"]+-TEX-LO-2211\\\"/.test(svg),
          keepsRootPath: /xlink:href=\\\"#MJX-[^\\\"]+-TEX-N-221A\\\"/.test(svg),
          keepsDoubleStruck: /-TEX-D-211D/.test(svg),
          keepsCalligraphic: /-TEX-C-4C/.test(svg),
          keepsSansSerif: /-TEX-SS-/.test(svg),
        };
      })()`, "SVG output font preferences");
      assert.ok(svgState.hasPalatino, JSON.stringify(svgState));
      assert.ok(svgState.hasSongti, JSON.stringify(svgState));
      assert.ok(svgState.keepsSumPath, JSON.stringify(svgState));
      assert.ok(svgState.keepsRootPath, JSON.stringify(svgState));
      assert.ok(svgState.keepsDoubleStruck, JSON.stringify(svgState));
      assert.ok(svgState.keepsCalligraphic, JSON.stringify(svgState));
      assert.ok(svgState.keepsSansSerif, JSON.stringify(svgState));

      await evaluate(`document.querySelector(".workspace-export-trigger")?.click()`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector("[data-copy-png-from-export]")),
      }))()`, "PNG copy action for output fonts");
      await evaluate(`document.querySelector("[data-copy-png-from-export]")?.click()`);
      const pngState = await waitForEvaluation(`(() => ({
        ready: Boolean(window.__visualtexOutputFontProbe?.png?.size),
        ...(window.__visualtexOutputFontProbe?.png || {}),
      }))()`, "PNG output font rasterization");
      assert.equal(pngState.type, "image/png");
      assert.ok(pngState.width > 0 && pngState.height > 0, JSON.stringify(pngState));

      console.log(JSON.stringify({ svgState, pngState }, null, 2));
      console.log("Targeted SVG/PNG output font regression passed");
      return;
    }

    if (scenario === "configuration") {
      await waitForEvaluation(`(() => ({
        ready:
          Boolean(document.querySelector(".settings-toggle")) &&
          Boolean(document.querySelector("[data-quick-ocr-button]")) &&
          Boolean(document.querySelector("[data-quick-ocr-mode-trigger]")) &&
          Boolean(document.querySelector("[data-silent-ocr-toggle]")),
      }))()`, "settings and quick OCR controls for configuration regression");

      const quickOcrLayouts = [];
      for (const width of [1400, 1040, 900]) {
        await client.send("Emulation.setDeviceMetricsOverride", {
          width,
          height: 1000,
          deviceScaleFactor: 1,
          mobile: false,
        });
        await sleep(90);
        quickOcrLayouts.push(await waitForEvaluation(`(() => {
          const header = document.querySelector(".editor-pane-header");
          const group = document.querySelector(".canvas-tool-group");
          const quick = document.querySelector("[data-quick-ocr-button]");
          const modeTrigger = document.querySelector("[data-quick-ocr-mode-trigger]");
          const silentInput = document.querySelector("[data-silent-ocr-toggle]");
          const silent = silentInput?.closest("label");
          const model = document.querySelector(".canvas-ocr-model");
          if (!header || !group || !quick || !modeTrigger || !silent || !model) return { ready: false };
          const headerRect = header.getBoundingClientRect();
          const groupRect = group.getBoundingClientRect();
          const quickRect = quick.getBoundingClientRect();
          const modeTriggerRect = modeTrigger.getBoundingClientRect();
          const silentRect = silent.getBoundingClientRect();
          const modelRect = model.getBoundingClientRect();
          const contained = [groupRect, quickRect, modeTriggerRect, silentRect, modelRect].every(
            (rect) =>
              rect.left >= headerRect.left - 1 &&
              rect.right <= headerRect.right + 1 &&
              rect.width > 0,
          );
          const silentStyle = getComputedStyle(silent);
          return {
            ready:
              contained &&
              quickRect.width >= 28 &&
              modeTriggerRect.width >= 18 &&
              silentRect.width >= 28 &&
              modelRect.width >= 100 &&
              document.documentElement.scrollWidth <= window.innerWidth + 1 &&
              group.scrollWidth <= group.clientWidth + 1,
            width: window.innerWidth,
            header: { left: headerRect.left, right: headerRect.right },
            group: {
              left: groupRect.left,
              right: groupRect.right,
              scrollWidth: group.scrollWidth,
              clientWidth: group.clientWidth,
            },
            quick: { left: quickRect.left, right: quickRect.right, width: quickRect.width },
            modeTrigger: {
              left: modeTriggerRect.left,
              right: modeTriggerRect.right,
              width: modeTriggerRect.width,
            },
            silent: {
              left: silentRect.left,
              right: silentRect.right,
              width: silentRect.width,
              display: silentStyle.display,
              className: silent.className,
              html: silent.outerHTML,
            },
            model: { left: modelRect.left, right: modelRect.right, width: modelRect.width },
          };
        })()`, `quick OCR toolbar layout at ${width}px`));
      }
      await client.send("Emulation.setDeviceMetricsOverride", {
        width: 1400,
        height: 1000,
        deviceScaleFactor: 1,
        mobile: false,
      });
      await sleep(90);

      await evaluate(`document.querySelector("[data-quick-ocr-mode-trigger]")?.click()`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector("[data-quick-ocr-mode-menu]")),
      }))()`, "Quick OCR capture mode menu");
      await evaluate(`document.querySelector('[data-quick-ocr-mode-option="system-screenshot"]')?.click()`);
      const quickOcrModeState = await waitForEvaluation(`(() => ({
        ready: localStorage.getItem("visualtex.quick-ocr.capture-mode") === "system-screenshot",
        value: localStorage.getItem("visualtex.quick-ocr.capture-mode"),
        menuOpen: Boolean(document.querySelector("[data-quick-ocr-mode-menu]")),
      }))()`, "persisted system-screenshot Quick OCR mode");

      await evaluate(`document.querySelector(".settings-toggle")?.click()`);
      await waitForEvaluation(`(() => ({
        ready:
          Boolean(document.querySelector("[data-save-configuration]")) &&
          Boolean(document.querySelector("[data-import-configuration]")),
      }))()`, "configuration controls");

      await evaluate(`document.querySelector('[data-theme-choice="green"]')?.click()`);
      await evaluate(`document.querySelector("[data-interface-customization-trigger]")?.click()`);
      await waitForEvaluation(`(() => ({
        ready:
          Boolean(document.querySelector("[data-formula-letter-font-setting]")) &&
          Boolean(document.querySelector("[data-formula-inset-left-setting]")),
      }))()`, "interface customization controls for configuration regression");
      await evaluate(`(() => {
        const letter = document.querySelector("[data-formula-letter-font-setting]");
        const chinese = document.querySelector("[data-formula-chinese-font-setting]");
        const left = document.querySelector("[data-formula-inset-left-setting]");
        letter.value = "palatino";
        letter.dispatchEvent(new Event("change", { bubbles: true }));
        chinese.value = "kaiti";
        chinese.dispatchEvent(new Event("change", { bubbles: true }));
        const valueSetter = Object.getOwnPropertyDescriptor(
          HTMLInputElement.prototype,
          "value",
        )?.set;
        valueSetter?.call(left, "19");
        left.dispatchEvent(new Event("input", { bubbles: true }));
      })()`);

      const layoutState = await waitForEvaluation(`(() => {
        const dialog = document.querySelector("[data-interface-customization-dialog]");
        const selectors = [
          "[data-formula-tool-button-size-setting]",
          "[data-formula-tool-button-padding-setting]",
          "[data-formula-inset-left-setting]",
          "[data-formula-inset-right-setting]",
          "[data-formula-row-vertical-inset-setting]",
        ];
        const controls = selectors.map((selector) => {
          const input = document.querySelector(selector);
          const label = input?.closest("label");
          const strong = label?.querySelector("strong");
          const rect = strong?.getBoundingClientRect();
          return {
            selector,
            text: strong?.textContent ?? "",
            width: rect?.width ?? 0,
            height: rect?.height ?? 0,
            whiteSpace: strong ? getComputedStyle(strong).whiteSpace : "",
            labelWidth: label?.getBoundingClientRect().width ?? 0,
          };
        });
        const settingsDialog = document.querySelector(".settings-dialog");
        const dialogRect = dialog?.getBoundingClientRect();
        const settingsRect = settingsDialog?.getBoundingClientRect();
        const rows = [...document.querySelectorAll(
          "[data-interface-customization-dialog] .formula-inset-range",
        )].map((row) => {
          const rect = row.getBoundingClientRect();
          return {
            left: rect.left,
            right: rect.right,
            scrollWidth: row.scrollWidth,
            clientWidth: row.clientWidth,
          };
        });
        return {
          ready:
            Boolean(dialogRect) &&
            Boolean(settingsRect) &&
            dialogRect.left >= settingsRect.left - 1 &&
            dialogRect.right <= settingsRect.right + 1 &&
            dialogRect.left >= -1 &&
            dialogRect.right <= window.innerWidth + 1 &&
            dialog.scrollWidth <= dialog.clientWidth + 1 &&
            rows.every(
              (row) =>
                row.left >= dialogRect.left - 1 &&
                row.right <= dialogRect.right + 1 &&
                row.scrollWidth <= row.clientWidth + 1,
            ) &&
            controls.every(
              (item) =>
                item.width >= 48 &&
                item.height <= 28 &&
                item.whiteSpace === "nowrap" &&
                item.labelWidth >= 300,
            ),
          dialogWidth: dialogRect?.width ?? 0,
          settingsWidth: settingsRect?.width ?? 0,
          dialogBounds: dialogRect
            ? { left: dialogRect.left, right: dialogRect.right }
            : null,
          settingsBounds: settingsRect
            ? { left: settingsRect.left, right: settingsRect.right }
            : null,
          rows,
          controls,
        };
      })()`, "contained interface customization slider layout");

      await evaluate(`document.querySelector("[data-interface-customization-close]")?.click()`);
      await evaluate(`(() => {
        const customTiles = {
          version: 3,
          sections: [{ id: "portable-section", name: "便携分区", createdAt: 1 }],
          tiles: [{
            id: "portable-tile",
            latex: "x^2+y^2",
            sectionId: "portable-section",
            color: "#4d9c8d",
            createdAt: 2,
          }],
        };
        localStorage.setItem("visualtex-custom-formula-tiles", JSON.stringify(customTiles));
        localStorage.setItem("visualtex-common-toolbar-command-ids-v1", JSON.stringify(["frac", "sqrt", "sum"]));
        localStorage.setItem("visualtex-formula-hotkeys-v1", JSON.stringify({ state: { bindings: [] }, version: 2 }));
        localStorage.setItem("visualtex-custom-formula-text-colors", JSON.stringify(["#123456", "#abcdef"]));
        localStorage.setItem("visualtex-custom-formula-background-colors", JSON.stringify(["#ffeeaa"]));
        localStorage.setItem("visualtex-desktop-editor-toolbar-open", "false");
        localStorage.setItem("visualtex-desktop-editor-tiles-open", "true");
        localStorage.setItem("visualtex.ocr.model", "PP-FormulaNet_plus-L");
        localStorage.setItem("visualtex.silent-ocr.enabled", "true");
        localStorage.setItem("visualtex.quick-ocr.capture-mode", "system-screenshot");
        window.__visualtexConfigurationProbe = { text: "", filename: "" };
        const originalCreateObjectURL = URL.createObjectURL.bind(URL);
        URL.createObjectURL = (blob) => {
          if (blob?.type?.includes("json")) {
            blob.text().then((text) => {
              window.__visualtexConfigurationProbe.text = text;
            });
          }
          return originalCreateObjectURL(blob);
        };
        HTMLAnchorElement.prototype.click = function () {
          window.__visualtexConfigurationProbe.filename = this.download || "";
        };
      })()`);
      await evaluate(`document.querySelector("[data-save-configuration]")?.click()`);
      const exported = await waitForEvaluation(`(() => {
        const probe = window.__visualtexConfigurationProbe || {};
        if (!probe.text) return { ready: false };
        const configuration = JSON.parse(probe.text);
        return {
          ready:
            configuration.schema === "visualtex-user-configuration" &&
            configuration.version === 1 &&
            configuration.editorSettings?.theme === "green" &&
            configuration.editorSettings?.formulaLetterFont === "palatino" &&
            configuration.editorSettings?.formulaChineseFont === "kaiti" &&
            configuration.editorSettings?.formulaInsetLeft === 19 &&
            Boolean(configuration.storage?.["visualtex-custom-formula-tiles"]) &&
            Boolean(configuration.storage?.["visualtex-formula-hotkeys-v1"]) &&
            configuration.storage?.["visualtex.ocr.model"] === "PP-FormulaNet_plus-L" &&
            configuration.storage?.["visualtex.silent-ocr.enabled"] === "true" &&
            configuration.storage?.["visualtex.quick-ocr.capture-mode"] === "system-screenshot" &&
            configuration.windows?.main?.width > 0 &&
            configuration.windows?.main?.height > 0 &&
            probe.filename.endsWith(".vtxconfig"),
          text: probe.text,
          filename: probe.filename,
          editorSettings: configuration.editorSettings,
          storage: configuration.storage,
          windows: configuration.windows,
        };
      })()`, "configuration export payload");

      await evaluate(`(() => {
        document.querySelector('[data-theme-choice="light"]')?.click();
        localStorage.removeItem("visualtex-custom-formula-tiles");
        localStorage.removeItem("visualtex-common-toolbar-command-ids-v1");
        localStorage.removeItem("visualtex-formula-hotkeys-v1");
        localStorage.removeItem("visualtex-custom-formula-text-colors");
        localStorage.removeItem("visualtex-custom-formula-background-colors");
        localStorage.setItem("visualtex-desktop-editor-toolbar-open", "true");
        localStorage.setItem("visualtex.ocr.model", "PP-FormulaNet_plus-M");
        localStorage.setItem("visualtex.silent-ocr.enabled", "false");
        localStorage.setItem("visualtex.quick-ocr.capture-mode", "immediate");
      })()`);
      const exportedText = exported.text;
      await evaluate(`(() => {
        const input = document.querySelector(".configuration-file-input");
        const transfer = new DataTransfer();
        transfer.items.add(new File([${JSON.stringify(exportedText)}], "portable.vtxconfig", { type: "application/json" }));
        input.files = transfer.files;
        input.dispatchEvent(new Event("change", { bubbles: true }));
      })()`);
      await sleep(1000);
      const imported = await waitForEvaluation(`(() => {
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        const state = persisted.state || {};
        const customTiles = JSON.parse(localStorage.getItem("visualtex-custom-formula-tiles") || "null");
        const textColors = JSON.parse(localStorage.getItem("visualtex-custom-formula-text-colors") || "[]");
        return {
          ready:
            state.theme === "green" &&
            state.formulaLetterFont === "palatino" &&
            state.formulaChineseFont === "kaiti" &&
            state.formulaInsetLeft === 19 &&
            customTiles?.sections?.[0]?.id === "portable-section" &&
            customTiles?.tiles?.[0]?.id === "portable-tile" &&
            textColors.includes("#123456") &&
            localStorage.getItem("visualtex-desktop-editor-toolbar-open") === "false" &&
            localStorage.getItem("visualtex-desktop-editor-tiles-open") === "true" &&
            localStorage.getItem("visualtex.ocr.model") === "PP-FormulaNet_plus-L" &&
            localStorage.getItem("visualtex.silent-ocr.enabled") === "true" &&
            localStorage.getItem("visualtex.quick-ocr.capture-mode") === "system-screenshot",
          state: {
            theme: state.theme,
            formulaLetterFont: state.formulaLetterFont,
            formulaChineseFont: state.formulaChineseFont,
            formulaInsetLeft: state.formulaInsetLeft,
          },
          customTiles,
          textColors,
          toolbarOpen: localStorage.getItem("visualtex-desktop-editor-toolbar-open"),
          tilesOpen: localStorage.getItem("visualtex-desktop-editor-tiles-open"),
          ocrModel: localStorage.getItem("visualtex.ocr.model"),
          silentOcr: localStorage.getItem("visualtex.silent-ocr.enabled"),
          quickOcrCaptureMode: localStorage.getItem("visualtex.quick-ocr.capture-mode"),
        };
      })()`, "configuration import round trip");

      console.log(JSON.stringify({ quickOcrLayouts, quickOcrModeState, layoutState, exported, imported }, null, 2));
      console.log("Targeted VisualTeX configuration round-trip regression passed");
      return;
    }

    if (scenario === "settings") {
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector(".canvas-input-behavior-trigger")),
      }))()`, "input behavior trigger");
      await evaluate(`document.querySelector(".canvas-input-behavior-trigger").click()`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector(".input-behavior-popover")),
      }))()`, "open input behavior settings");
      const defaults = await waitForEvaluation(`(() => {
        const options = [...document.querySelectorAll(".input-behavior-option")].map((label) => ({
          title: label.querySelector("strong")?.textContent ?? "",
          checked: label.querySelector('input[type="checkbox"]')?.checked ?? false,
        }));
        const structured = options.find(
          (item) =>
            item.title.includes("求和、积分") ||
            item.title.includes("Structured command suggestions"),
        );
        const other = options.find(
          (item) =>
            item.title.includes("其他命令") ||
            item.title.includes("Other command suggestions"),
        );
        return {
          ready: Boolean(structured && other),
          structured,
          other,
          options,
        };
      })()`, "candidate preference defaults");
      if (!defaults.structured.checked || defaults.other.checked) {
        throw new Error(`Unexpected candidate defaults: ${JSON.stringify(defaults)}`);
      }
      await evaluate(`document.querySelector(".canvas-input-behavior-trigger").click()`);
      await waitForEvaluation(`(() => ({
        ready: !document.querySelector(".input-behavior-popover"),
      }))()`, "close input behavior settings");

      await clearField();
      await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.setValue("\\\\sum", {
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
          data: "m",
        }));
      })()`);
      const structuredState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const surface = document.querySelector(".multi-line-editor");
        const nativePanel = document.getElementById("mathlive-suggestion-popover");
        return {
          ready: Boolean(document.querySelector(".suggestion-popup")),
          customVisible: Boolean(document.querySelector(".suggestion-popup")),
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          commandQuery: surface?.dataset.commandQuery ?? "",
          nativeVisible: nativePanel?.classList.contains("is-visible") ?? false,
          nativeCommands: [...(nativePanel?.querySelectorAll("li[data-command]") ?? [])]
            .map((item) => item.dataset.command ?? ""),
          contentHtml: field?.shadowRoot?.querySelector('[part="content"]')?.innerHTML ?? "",
          keyboardSinkValue: field?.shadowRoot?.querySelector('[part="keyboard-sink"]')?.value ?? "",
          shadowText: field?.shadowRoot?.textContent ?? "",
        };
      })()`, "structured VisualTeX suggestion panel");

      await client.send("Page.reload", { ignoreCache: true });
      await waitForEvaluation(
        `(() => ({ ready: Boolean(document.querySelector("math-field")) }))()`,
        "fresh formula field for other-command test",
      );
      await focusField();
      await typeText("\\theta");
      const otherState = await waitForEvaluation(`(() => {
        const nativePanel = document.getElementById("mathlive-suggestion-popover");
        return {
          ready:
            !document.querySelector(".suggestion-popup") &&
            Boolean(nativePanel?.classList.contains("is-visible")),
          customVisible: Boolean(document.querySelector(".suggestion-popup")),
          nativeVisible: nativePanel?.classList.contains("is-visible") ?? false,
          commandQuery: document.querySelector(".multi-line-editor")?.dataset.commandQuery ?? "",
          customText: document.querySelector(".suggestion-popup")?.textContent ?? "",
          value: document.querySelector("math-field")?.value ?? "",
          mode: document.querySelector("math-field")?.mode ?? "",
        };
      })()`, "other command uses only native panel");

      console.log(JSON.stringify({ defaults, structuredState, otherState }, null, 2));
      console.log("Targeted suggestion settings regression passed");
      return;
    }

    if (scenario === "delete") {
      await evaluate(`(() => {
        const storageKey = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
        const first = { id: crypto.randomUUID(), latex: "a" };
        const second = { id: crypto.randomUUID(), latex: "" };
        persisted.state = {
          ...(persisted.state || {}),
          lines: [first, second],
          activeLineId: second.id,
        };
        localStorage.setItem(storageKey, JSON.stringify(persisted));
        location.reload();
      })()`);
      await waitForEvaluation(`(() => ({
        ready: document.querySelectorAll("math-field").length === 2,
      }))()`, "two formula lines for delete test");
      await evaluate(`(() => {
        const field = document.querySelectorAll("math-field")[1];
        field.setValue("", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.mode = "math";
        field.position = field.lastOffset;
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        field.dispatchEvent(new FocusEvent("focus", { bubbles: true, composed: true }));
      })()`);
      await sleep(120);
      await typeText("\\mat");
      const beforeDelete = await waitForEvaluation(`(() => {
        const field = document.querySelectorAll("math-field")[1];
        const caret = field.shadowRoot?.querySelector(".ML__raw-latex.ML__latex-caret");
        const nodes = caret?.parentElement
          ? [...caret.parentElement.querySelectorAll(":scope > .ML__raw-latex")]
          : [];
        const caretIndex = nodes.indexOf(caret);
        const typedRaw = nodes
          .slice(0, caretIndex >= 0 ? caretIndex + 1 : nodes.length)
          .map((node) => node.textContent ?? "")
          .join("");
        const renderedRaw = [...(field.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
          .map((node) => node.textContent ?? "")
          .join("");
        return {
          ready: typedRaw === "\\\\mat" && document.querySelectorAll("math-field").length === 2,
          typedRaw,
          renderedRaw,
          lineCount: document.querySelectorAll("math-field").length,
          contentHtml: field.shadowRoot?.querySelector('[part="content"]')?.innerHTML ?? "",
        };
      })()`, "raw math command before Backspace");
      await key("Backspace", "Backspace", 8);
      const firstDelete = await waitForEvaluation(`(() => {
        const field = document.querySelectorAll("math-field")[1];
        const caret = field.shadowRoot?.querySelector(".ML__raw-latex.ML__latex-caret");
        const nodes = caret?.parentElement
          ? [...caret.parentElement.querySelectorAll(":scope > .ML__raw-latex")]
          : [];
        const caretIndex = nodes.indexOf(caret);
        const typedRaw = nodes
          .slice(0, caretIndex >= 0 ? caretIndex + 1 : nodes.length)
          .map((node) => node.textContent ?? "")
          .join("");
        return {
          ready: typedRaw === "\\\\ma" && document.querySelectorAll("math-field").length === 2,
          typedRaw,
          lineCount: document.querySelectorAll("math-field").length,
          firstValue: document.querySelectorAll("math-field")[0]?.value ?? "",
        };
      })()`, "one-character raw command deletion");
      await key("Backspace", "Backspace", 8);
      const secondDelete = await waitForEvaluation(`(() => {
        const field = document.querySelectorAll("math-field")[1];
        const caret = field.shadowRoot?.querySelector(".ML__raw-latex.ML__latex-caret");
        const nodes = caret?.parentElement
          ? [...caret.parentElement.querySelectorAll(":scope > .ML__raw-latex")]
          : [];
        const caretIndex = nodes.indexOf(caret);
        const typedRaw = nodes
          .slice(0, caretIndex >= 0 ? caretIndex + 1 : nodes.length)
          .map((node) => node.textContent ?? "")
          .join("");
        return {
          ready: typedRaw === "\\\\m" && document.querySelectorAll("math-field").length === 2,
          typedRaw,
          lineCount: document.querySelectorAll("math-field").length,
        };
      })()`, "second one-character raw command deletion");

      await key("Backspace", "Backspace", 8);
      await waitForEvaluation(`(() => {
        const field = document.querySelectorAll("math-field")[1];
        const caret = field.shadowRoot?.querySelector(".ML__raw-latex.ML__latex-caret");
        const nodes = caret?.parentElement
          ? [...caret.parentElement.querySelectorAll(":scope > .ML__raw-latex")]
          : [];
        const caretIndex = nodes.indexOf(caret);
        const typedRaw = nodes
          .slice(0, caretIndex >= 0 ? caretIndex + 1 : nodes.length)
          .map((node) => node.textContent ?? "")
          .join("");
        return { ready: typedRaw === "\\\\", typedRaw };
      })()`, "raw command reduced to backslash");
      await key("Backspace", "Backspace", 8);
      await waitForEvaluation(`(() => {
        const field = document.querySelectorAll("math-field")[1];
        const rawText = [...(field.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
          .map((node) => node.textContent ?? "")
          .join("");
        return {
          ready: rawText === "" && document.querySelectorAll("math-field").length === 2,
          rawText,
          mode: field.mode,
        };
      })()`, "raw command fully cleared without deleting row");
      await key("Backspace", "Backspace", 8);
      const emptyLineDelete = await waitForEvaluation(`(() => {
        const fields = [...document.querySelectorAll("math-field")];
        return {
          ready:
            fields.length === 1 &&
            fields[0]?.value === "a" &&
            fields[0]?.hasFocus?.(),
          lineCount: fields.length,
          values: fields.map((field) => field.value),
          activeTag: document.activeElement?.tagName ?? "",
        };
      })()`, "empty raw-latex row returns to previous formula");

      console.log(JSON.stringify({ beforeDelete, firstDelete, secondDelete, emptyLineDelete }, null, 2));
      console.log("Targeted delete regression passed");
      return;
    }

    const toolbarOrder = await waitForEvaluation(`(() => {
      const group = document.querySelector(".canvas-tool-group");
      const exportMenu = group?.querySelector(".export-menu");
      const behavior = group?.querySelector(".input-behavior-menu");
      const children = [...(group?.children ?? [])];
      const fileActions = document.querySelector(".header-actions .file-actions");
      const editActions = document.querySelector(".header-actions .edit-actions");
      const titleInput = document.querySelector(".document-title-area input");
      const fileStyle = fileActions ? getComputedStyle(fileActions) : null;
      const fileBounds = fileActions?.getBoundingClientRect();
      const editBounds = editActions?.getBoundingClientRect();
      const titleBounds = titleInput?.getBoundingClientRect();
      return {
        ready: Boolean(exportMenu && behavior && fileActions && editActions && titleInput),
        exportIndex: children.indexOf(exportMenu),
        behaviorIndex: children.indexOf(behavior),
        fileBorderWidth: fileStyle?.borderTopWidth ?? "",
        fileBackground: fileStyle?.backgroundColor ?? "",
        fileLeftOffset: fileStyle?.left ?? "",
        titleRight: titleBounds?.right ?? 0,
        fileLeft: fileBounds?.left ?? 0,
        fileRight: fileBounds?.right ?? 0,
        editLeft: editBounds?.left ?? 0,
      };
    })()`, "unified export placement and shifted file actions");
    if (
      toolbarOrder.exportIndex < 0 ||
      toolbarOrder.behaviorIndex < 0 ||
      toolbarOrder.exportIndex >= toolbarOrder.behaviorIndex ||
      toolbarOrder.fileBorderWidth !== "0px" ||
      toolbarOrder.fileLeftOffset !== "6px" ||
      toolbarOrder.titleRight > toolbarOrder.fileLeft ||
      toolbarOrder.editLeft - toolbarOrder.fileRight < 4
    ) {
      throw new Error(`Incorrect export/header placement: ${JSON.stringify(toolbarOrder)}`);
    }

    await evaluate(`document.querySelector(".export-menu-trigger")?.click()`);
    const exportMenuState = await waitForEvaluation(`(() => {
      const popover = document.querySelector(".export-menu-popover");
      const labels = [...document.querySelectorAll(".export-format-options strong")]
        .map((node) => node.textContent?.trim() ?? "");
      const pathSection = document.querySelector(".export-path-section");
      return {
        ready:
          Boolean(popover && pathSection) &&
          labels.join(",") === "Markdown,SVG,PNG",
        labels,
        pathText: pathSection?.textContent?.replace(/\\s+/g, " ").trim() ?? "",
      };
    })()`, "unified export menu options and path selector");
    await evaluate(`document.querySelector(".export-menu-trigger")?.click()`);
    await waitForEvaluation(`(() => ({
      ready: !document.querySelector(".export-menu-popover"),
    }))()`, "export menu closed before matrix test");

    await evaluate(`document.querySelector('button[data-category="matrix"]').click()`);
    const gridState = await waitForEvaluation(`(() => ({
      ready: document.querySelectorAll(".matrix-size-cell").length === 100,
      cellCount: document.querySelectorAll(".matrix-size-cell").length,
    }))()`, "10 by 10 matrix grid");

    await evaluate(`(() => {
      const cell = document.querySelector('.matrix-size-cell[data-matrix-rows="3"][data-matrix-columns="4"]');
      cell?.focus();
      cell?.dispatchEvent(new FocusEvent("focusin", {
        bubbles: true,
        composed: true,
        relatedTarget: null,
      }));
    })()`);
    const hoverState = await waitForEvaluation(`(() => ({
      ready:
        document.querySelector(".matrix-size-badge")?.textContent?.replace(/\\s+/g, " ").trim() === "3 × 4" &&
        document.querySelectorAll(".matrix-size-cell.is-previewed").length === 12,
      badge: document.querySelector(".matrix-size-badge")?.textContent?.replace(/\\s+/g, " ").trim() ?? "",
      previewedCount: document.querySelectorAll(".matrix-size-cell.is-previewed").length,
    }))()`, "matrix hover preview");

    await evaluate(`document.querySelector('.matrix-size-cell[data-matrix-rows="3"][data-matrix-columns="4"]').click()`);
    const selectedState = await waitForEvaluation(`(() => ({
      ready:
        document.querySelector('.matrix-size-cell[data-matrix-rows="3"][data-matrix-columns="4"]')?.classList.contains("is-selected-corner") &&
        document.querySelector(".matrix-insert-button")?.textContent?.includes("3 × 4"),
      selectedCorner: document.querySelector('.matrix-size-cell[data-matrix-rows="3"][data-matrix-columns="4"]')?.classList.contains("is-selected-corner") ?? false,
      insertLabel: document.querySelector(".matrix-insert-button")?.textContent ?? "",
    }))()`, "matrix selection");

    await evaluate(`document.querySelector(".matrix-insert-button").click()`);
    const insertionState = await waitForEvaluation(`(() => {
      const value = document.querySelector("math-field")?.value ?? "";
      const body = value.match(/\\\\begin\\{bmatrix\\}([\\s\\S]*?)\\\\end\\{bmatrix\\}/)?.[1] ?? "";
      return {
        ready:
          value.includes("\\\\begin{bmatrix}") &&
          body.split(/\\\\\\\\/).length === 3 &&
          body.split(/\\\\\\\\/).every((row) => row.split("&").length === 4),
        value,
      };
    })()`, "3 by 4 matrix insertion");

    console.log(
      JSON.stringify(
        {
          toolbarOrder,
          exportMenuState,
          gridState,
          hoverState,
          selectedState,
          insertionState,
        },
        null,
        2,
      ),
    );
    console.log("Targeted layout regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(250);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => undefined);
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.stack : error);
  process.exitCode = 1;
});
