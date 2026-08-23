import { spawn } from "node:child_process";
import { rm, writeFile } from "node:fs/promises";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const scenario = process.argv[2];
if (!new Set(["wrapper", "wrapper-auto", "wrapper-continuous", "wrapper-prefix", "native-input-popover", "usage-ranking", "native-space-selection", "candidate-query-reset", "raw-placeholder-visual", "placeholder-selection", "structural-placeholder", "structured-chinese-ime", "direct-shortcut-placeholder", "toolbar-placeholder-overflow", "horizontal-overflow", "accent-placeholder", "caret-probe", "scripts", "upright", "context-style", "suggestions", "navigation", "geometry", "source-layout", "toolbar-compact", "toolbar-postfix", "classic-panel-resize", "ocr-storage-ui", "formula-tiles", "formula-formatting", "cursor-placement", "settings", "layout", "multi-line-selection", "delete", "export"]).has(scenario)) {
  throw new Error(
    "Usage: node scripts/targeted_editor_regression.mjs <wrapper|wrapper-auto|wrapper-continuous|wrapper-prefix|native-input-popover|usage-ranking|native-space-selection|candidate-query-reset|raw-placeholder-visual|placeholder-selection|structural-placeholder|structured-chinese-ime|direct-shortcut-placeholder|toolbar-placeholder-overflow|horizontal-overflow|accent-placeholder|caret-probe|scripts|upright|context-style|suggestions|navigation|geometry|source-layout|toolbar-compact|toolbar-postfix|classic-panel-resize|ocr-storage-ui|formula-tiles|formula-formatting|cursor-placement|settings|layout|multi-line-selection|delete|export>",
  );
}

const offset = process.pid % 1000;
const previewPort = 6400 + offset;
const debugPort = 11600 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = createBrowserProfilePath(`visualtex-targeted-${scenario}`);
const chromePath = resolveChromiumExecutable();
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
      await sleep(45);
    };

    const typeText = async (text) => {
      for (const character of text) {
        const code = character === "\\" ? "Backslash" : `Key${character.toUpperCase()}`;
        const virtualKeyCode = character === "\\" ? 220 : character.toUpperCase().charCodeAt(0);
        await key(character, code, virtualKeyCode);
      }
    };

    const clickSelectorWithPointer = async (selector) => {
      const point = await waitForEvaluation(`(() => {
        const element = document.querySelector(__SELECTOR__);
        if (!(element instanceof HTMLElement)) return { ready: false };
        const rect = element.getBoundingClientRect();
        return {
          ready: rect.width > 0 && rect.height > 0,
          x: rect.left + rect.width / 2,
          y: rect.top + rect.height / 2,
        };
      })()`.replace("__SELECTOR__", JSON.stringify(selector)), `pointer target ${selector}`);
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: point.x,
        y: point.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: point.x,
        y: point.y,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });
      await sleep(80);
    };

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
      if (${JSON.stringify(scenario)} === "toolbar-compact") {
        localStorage.removeItem("visualtex-common-toolbar-command-ids-v2");
        localStorage.setItem(
          "visualtex-common-toolbar-command-ids-v1",
          JSON.stringify([
            "frac", "sqrt", "power", "subscript", "hat", "tilde",
            "parentheses", "absolute", "intplain", "int", "iint", "oint",
            "sum", "prod", "lim", "partial", "derivative", "nabla", "infty",
            "matrix2", "cases", "vector", "alpha", "beta", "gamma", "theta",
            "lambda", "mu", "pi", "sigma", "omega", "delta", "equal", "neq",
            "approx", "leq", "geq", "propto", "in", "subset", "rightarrow",
            "notin", "forall", "exists", "leftarrow",
          ]),
        );
      }
      const storageKey = "visualtex-editor";
      const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
      persisted.state = {
        ...(persisted.state || {}),
        lines: [{ id: crypto.randomUUID(), latex: "" }],
        activeLineId: null,
        ${scenario === "formula-tiles" || scenario === "usage-ranking" || scenario === "toolbar-placeholder-overflow" ? 'editorLayout: "standard",\r\n        sidebarOpen: true,' : ""}
        ${scenario === "toolbar-postfix" ? 'editorLayout: "classic",' : ""}
        ${scenario === "usage-ranking" ? `personalize: true,
        usage: {
          sqrt: {
            commandId: "sqrt",
            useCount: 18,
            lastUsedAt: 1800,
            recentUses: [],
            acceptedPrefixes: {},
            contextCounts: { toolbar: 18 },
            pinned: false,
          },
          "formula-tile-gaussian-integral": {
            commandId: "formula-tile-gaussian-integral",
            useCount: 12,
            lastUsedAt: 1700,
            recentUses: [],
            acceptedPrefixes: {},
            contextCounts: { toolbar: 12 },
            pinned: false,
          },
          "mathlive-native:\\\\beth": {
            commandId: "mathlive-native:\\\\beth",
            useCount: 25,
            lastUsedAt: 1900,
            recentUses: [],
            acceptedPrefixes: { be: 25 },
            contextCounts: { candidate: 25 },
            pinned: false,
          },
        },` : ""}
      };
      persisted.state.activeLineId = persisted.state.lines[0].id;
      if (
        ${JSON.stringify(scenario)} === "native-input-popover" ||
        ${JSON.stringify(scenario)} === "usage-ranking" ||
        ${JSON.stringify(scenario)} === "upright"
      ) {
        persisted.state.inputBehavior = {
          autoEscapeShortcuts: ${scenario === "upright" ? "false" : "true"},
          autoExitSuperscript: true,
          autoExitSubscript: true,
          autoExitAccent: true,
          autoExitWrapperCommand: true,
          showStructuredCommandSuggestions: false,
          showOtherCommandSuggestions: false,
        };
      } else {
        delete persisted.state.inputBehavior;
      }
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

    if (scenario === "formula-formatting") {
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
        await evaluate(`(() => {
          const field = document.querySelector('math-field');
          if (!field) return false;
          field.focus();
          field.selection = {
            ranges: [[0, field.lastOffset]],
            direction: 'forward',
          };
          return true;
        })()`);
        await clickSelectorWithPointer(selector);
        await sleep(100);
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
          field.focus();
          return true;
        })()`);
        await sleep(60);
      };

      const readFormattingValue = () => evaluate(`(() => {
        const field = document.querySelector('math-field');
        return field?.value?.replace(/\\s+/g, '') ?? '';
      })()`);

      await setFormattingValue('abc');
      await dispatchFormattingToggle('[data-formula-selection-bold]');
      const boldApplied = await readFormattingValue();
      if (boldApplied !== String.raw`\mathbf{abc}`) {
        throw new Error(`Bold toggle must emit \\mathbf exactly; received ${boldApplied}`);
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

      await setFormattingValue('abc');
      await evaluate(`(() => {
        const field = document.querySelector('math-field');
        if (!field) return false;
        field.focus();
        field.selection = { ranges: [[0, field.lastOffset]], direction: 'forward' };
        return true;
      })()`);
      await clickSelectorWithPointer('[data-formula-selection-color]');
      const compactColorPopover = await waitForEvaluation(`(() => {
        const popover = document.querySelector('[data-formula-color-popover="color"]');
        const presets = popover?.querySelector('.formula-color-presets');
        const custom = popover?.querySelector('.formula-custom-colors-panel');
        if (!(popover instanceof HTMLElement) || !(presets instanceof HTMLElement) || !(custom instanceof HTMLElement)) {
          return { ready: false };
        }
        const rect = popover.getBoundingClientRect();
        const presetRect = presets.getBoundingClientRect();
        const customRect = custom.getBoundingClientRect();
        return {
          ready: rect.width <= 200 && customRect.top >= presetRect.bottom - 1,
          width: rect.width,
          presetBottom: presetRect.bottom,
          customTop: customRect.top,
        };
      })()`, 'compact formula color popover');
      if (compactColorPopover.width > 200) {
        throw new Error(`Formula color popover is still oversized: ${JSON.stringify(compactColorPopover)}`);
      }
      await clickSelectorWithPointer('[data-formula-color="#111827"]');
      await sleep(100);
      const textColorApplied = await readFormattingValue();
      if (!textColorApplied.includes('111827') || !textColorApplied.includes('abc')) {
        throw new Error(`Formula text color was not applied to the selected content: ${textColorApplied}`);
      }

      await setFormattingValue('xyz');
      await evaluate(`(() => {
        const field = document.querySelector('math-field');
        if (!field) return false;
        field.focus();
        field.selection = { ranges: [[0, field.lastOffset]], direction: 'forward' };
        return true;
      })()`);
      await clickSelectorWithPointer('[data-formula-selection-background]');
      await clickSelectorWithPointer('[data-formula-color="#fef3c7"]');
      await sleep(100);
      const backgroundColorApplied = await readFormattingValue();
      if (!backgroundColorApplied.toLowerCase().includes('fef3c7') || !backgroundColorApplied.includes('xyz')) {
        throw new Error(`Formula background color was not applied to the selected content: ${backgroundColorApplied}`);
      }

      const selectionColors = {
        text: textColorApplied,
        background: backgroundColorApplied,
        popover: compactColorPopover,
      };

      const removedPersistentControls = await evaluate(`(() => ({
        typingBold: Boolean(document.querySelector('[data-formula-typing-bold]')),
        typingItalic: Boolean(document.querySelector('[data-formula-typing-italic]')),
      }))()`);
      if (removedPersistentControls.typingBold || removedPersistentControls.typingItalic) {
        throw new Error(`Persistent typing controls must stay removed: ${JSON.stringify(removedPersistentControls)}`);
      }

      console.log(JSON.stringify({
        selectionBold,
        selectionItalic,
        selectionColors,
        removedPersistentControls,
      }, null, 2));
      console.log("Targeted desktop formula formatting regression passed");
      return;
    }

    if (scenario === "multi-line-selection") {
      await evaluate(`(() => {
        const storageKey = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
        const lines = [
          { id: crypto.randomUUID(), latex: "abcdefghij" },
          { id: crypto.randomUUID(), latex: "klmnopqrst" },
          { id: crypto.randomUUID(), latex: "uvwxyzabcd" },
          { id: crypto.randomUUID(), latex: "efghijklmn" },
        ];
        persisted.state = {
          ...(persisted.state || {}),
          lines,
          activeLineId: lines[0].id,
          editorLayout: "standard",
          sidebarOpen: false,
        };
        localStorage.setItem(storageKey, JSON.stringify(persisted));
      })()`);
      await client.send("Page.reload", { ignoreCache: true });
      await waitForEvaluation(`(() => ({
        ready: document.querySelectorAll("math-field").length === 4,
      }))()`, "four formula rows for multiline selection");

      const toolbar = await waitForEvaluation(`(() => {
        const selectors = [
          "[data-formula-selection-bold]",
          "[data-formula-selection-italic]",
          "[data-formula-selection-color]",
          "[data-formula-selection-background]",
        ];
        const buttons = selectors.map((selector) => document.querySelector(selector));
        const visible = buttons.every((button) => {
          if (!(button instanceof HTMLElement)) return false;
          const style = getComputedStyle(button);
          const rect = button.getBoundingClientRect();
          return style.display !== "none" && style.visibility !== "hidden" && rect.width > 0 && rect.height > 0;
        });
        return {
          ready: visible,
          count: buttons.filter(Boolean).length,
          visible,
          noHorizontalOverflow:
            document.documentElement.scrollWidth <= document.documentElement.clientWidth + 2,
        };
      })()`, "four selection-only formula formatting controls in the main editor");
      if (!toolbar.noHorizontalOverflow) {
        throw new Error(`Formula formatting controls caused horizontal overflow: ${JSON.stringify(toolbar)}`);
      }

      const drag = await waitForEvaluation(`(() => {
        const fields = [...document.querySelectorAll("math-field")];
        const firstContent = fields[0]?.shadowRoot?.querySelector('[part="content"]');
        const lastContent = fields.at(-1)?.shadowRoot?.querySelector('[part="content"]');
        const first = firstContent?.getBoundingClientRect();
        const last = lastContent?.getBoundingClientRect();
        return {
          ready: Boolean(first && last && first.width > 0 && last.width > 0),
          start: first ? { x: first.left + 2, y: first.top + first.height / 2 } : null,
          end: last ? { x: last.right - 2, y: last.top + last.height / 2 } : null,
        };
      })()`, "multiline selection drag geometry");

      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: drag.start.x,
        y: drag.start.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseMoved",
        x: drag.end.x,
        y: drag.end.y,
        button: "left",
        buttons: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: drag.end.x,
        y: drag.end.y,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });

      const selectionState = await waitForEvaluation(`(() => {
        const rows = [...document.querySelectorAll(".formula-line")];
        const details = rows.map((row) => {
          const field = row.querySelector("math-field");
          const selections = [...(field?.shadowRoot?.querySelectorAll(".ML__selection") ?? [])]
            .filter((node) => {
              const style = getComputedStyle(node);
              const rect = node.getBoundingClientRect();
              return style.display !== "none" && style.visibility !== "hidden" && rect.width > 0 && rect.height > 0;
            });
          const rects = selections.map((node) => node.getBoundingClientRect());
          const union = rects.length
            ? {
                left: Math.min(...rects.map((rect) => rect.left)),
                right: Math.max(...rects.map((rect) => rect.right)),
                top: Math.min(...rects.map((rect) => rect.top)),
                bottom: Math.max(...rects.map((rect) => rect.bottom)),
              }
            : null;
          const fieldRect = field?.getBoundingClientRect();
          const rowRect = row.getBoundingClientRect();
          return {
            selectedClass: row.classList.contains("is-multi-line-selected"),
            selection: field ? JSON.parse(JSON.stringify(field.selection)) : null,
            visibleSelectionCount: selections.length,
            union,
            fieldRect: fieldRect
              ? { left: fieldRect.left, right: fieldRect.right, width: fieldRect.width }
              : null,
            rowWidth: rowRect.width,
            rowBackground: getComputedStyle(row).backgroundColor,
          };
        });
        const ready =
          details.length === 4 &&
          details.every((detail) =>
            detail.selectedClass &&
            detail.visibleSelectionCount > 0 &&
            detail.union &&
            detail.fieldRect &&
            detail.union.right > detail.union.left &&
            detail.union.left >= detail.fieldRect.left - 3 &&
            detail.union.right <= detail.fieldRect.right + 3 &&
            detail.union.right - detail.union.left < detail.rowWidth - 40
          );
        return { ready, details };
      })()`, "natural text-bounded highlight on every selected formula row");

      await sleep(300);
      const beforeMove = await evaluate(`(() =>
        [...document.querySelectorAll("math-field")].map((field) =>
          JSON.parse(JSON.stringify(field.selection)),
        )
      )()`);
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseMoved",
        x: drag.end.x + 160,
        y: drag.end.y,
        button: "none",
        buttons: 0,
      });
      await sleep(120);
      const afterMove = await evaluate(`(() =>
        [...document.querySelectorAll("math-field")].map((field) =>
          JSON.parse(JSON.stringify(field.selection)),
        )
      )()`);
      if (JSON.stringify(afterMove) !== JSON.stringify(beforeMove)) {
        throw new Error(`Multiline selection changed after mouse release: ${JSON.stringify({ beforeMove, afterMove })}`);
      }

      console.log(JSON.stringify({ toolbar, selectionState }, null, 2));
      console.log("Targeted multiline selection visual regression passed");
      return;
    }

    if (scenario === "ocr-storage-ui") {
      await evaluate(`(() => {
        localStorage.setItem("visualtex.onboarding.v3.completed", "true");
        let callbackId = 1;
        const callbacks = new Map();
        const initialRuntime = {
          installed: true,
          pythonPath: "C:\\\\Users\\\\Tester\\\\VisualTeX-OCR\\\\python\\\\python.exe",
          pythonVersion: "3.12.10",
          paddleVersion: "3.3.1",
          paddleocrVersion: "3.7.0",
          runtimePath: "C:\\\\Users\\\\Tester\\\\VisualTeX-OCR",
          storageConfigPath: "C:\\\\Users\\\\Tester\\\\AppData\\\\Roaming\\\\VisualTeX\\\\ocr-storage.json",
          storageSource: "configured",
          storageManaged: true,
          storageAvailableBytes: 68719476736,
          storagePersistentAcrossUninstall: true,
          runtimeBundleAvailable: true,
          offlineBundleAvailable: false,
          installedModels: ["PP-FormulaNet_plus-M"],
          damagedModels: [],
          modelCatalogAvailable: true,
          defaultModel: "PP-FormulaNet_plus-M",
          message: "OCR runtime ready",
        };
        const runtimeAt = (parent, installed) => {
          const runtimePath = parent + "\\\\VisualTeX-OCR";
          return {
            ...initialRuntime,
            installed,
            pythonPath: installed ? runtimePath + "\\\\python\\\\python.exe" : null,
            pythonVersion: installed ? "3.12.10" : null,
            paddleVersion: installed ? "3.3.1" : null,
            paddleocrVersion: installed ? "3.7.0" : null,
            runtimePath,
            storageAvailableBytes: 128849018880,
            installedModels: [],
            message: installed
              ? "OCR runtime reinstalled at the selected location"
              : "OCR runtime is not installed at the selected location",
          };
        };
        window.__visualtexOcrStorageCalls = [];
        window.__visualtexOcrCurrentRuntime = initialRuntime;
        window.__visualtexOcrDialogSelections = [
          "D:\\\\OCR Data",
          "E:\\\\Second OCR",
        ];
        window.__visualtexDelayNextRuntimeStatus = false;
        window.confirm = () => true;
        window.__TAURI_INTERNALS__ = {
          metadata: {
            currentWindow: { label: "main" },
            currentWebview: { windowLabel: "main", label: "main" },
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
            window.__visualtexOcrStorageCalls.push({ command, args });
            if (command === "get_ocr_runtime_status") {
              const snapshot = window.__visualtexOcrCurrentRuntime;
              if (window.__visualtexDelayNextRuntimeStatus) {
                window.__visualtexDelayNextRuntimeStatus = false;
                await new Promise((resolve) => setTimeout(resolve, 450));
                return initialRuntime;
              }
              return snapshot;
            }
            if (command === "get_ocr_install_status") {
              const installed = Boolean(window.__visualtexOcrCurrentRuntime?.installed);
              return {
                schemaVersion: 1,
                state: installed ? "complete" : "notInstalled",
                currentStep: null,
                completedSteps: installed ? ["dependencies", "verify"] : [],
                percent: installed ? 100 : 0,
                message: installed ? "OCR runtime ready" : "OCR runtime is not installed",
                detail: null,
                error: null,
                logPath: "C:\\\\fake\\\\ocr-install.log",
                updatedAtMs: Date.now(),
              };
            }
            if (command === "get_ocr_model_catalog") {
              return {
                schemaVersion: 1,
                platform: "windows",
                architecture: "x64",
                entries: [
                  {
                    model: "PP-FormulaNet_plus-S",
                    url: "https://example.invalid/S",
                    size: 200050659,
                    sha256: "a".repeat(64),
                  },
                  {
                    model: "PP-FormulaNet_plus-M",
                    url: "https://example.invalid/M",
                    size: 425830895,
                    sha256: "b".repeat(64),
                  },
                  {
                    model: "PP-FormulaNet_plus-L",
                    url: "https://example.invalid/L",
                    size: 670293702,
                    sha256: "c".repeat(64),
                  },
                ],
              };
            }
            if (command === "get_ocr_model_download_status") return null;
            if (command === "configure_ocr_storage_location") {
              window.__visualtexOcrCurrentRuntime = runtimeAt(args.selectedDirectory, false);
              return window.__visualtexOcrCurrentRuntime;
            }
            if (command === "install_ocr_runtime") {
              const currentPath = window.__visualtexOcrCurrentRuntime.runtimePath;
              const parent = currentPath.replace(/\\\\VisualTeX-OCR$/i, "");
              window.__visualtexOcrCurrentRuntime = runtimeAt(parent, true);
              return window.__visualtexOcrCurrentRuntime;
            }
            if (command === "reset_ocr_runtime") {
              const currentPath = window.__visualtexOcrCurrentRuntime.runtimePath;
              const parent = currentPath.replace(/\\\\VisualTeX-OCR$/i, "");
              window.__visualtexOcrCurrentRuntime = runtimeAt(parent, false);
              return window.__visualtexOcrCurrentRuntime;
            }
            if (command === "open_ocr_storage_location") return null;
            if (command === "plugin:dialog|open") {
              return window.__visualtexOcrDialogSelections.shift() ?? null;
            }
            if (
              command === "plugin:event|listen" ||
              command === "plugin:event|unlisten"
            ) {
              return 1;
            }
            throw new Error("Unexpected OCR storage UI command: " + command);
          },
        };
      })()`);

      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector(".menu-button")),
      }))()`, "main menu button for OCR storage UI");
      await clickSelectorWithPointer(".menu-button");
      await waitForEvaluation(`(() => {
        const button = document.querySelector(".menu-button");
        return {
          ready: Boolean(document.querySelector("#app-main-menu")),
          expanded: button?.getAttribute("aria-expanded") ?? null,
          buttonConnected: Boolean(button?.isConnected),
          onboardingVisible: Boolean(document.querySelector(".onboarding-overlay")),
        };
      })()`, "main menu for OCR storage UI");
      await evaluate(`(() => {
        const item = [...document.querySelectorAll('#app-main-menu [role="menuitem"]')]
          .find((button) => /图片公式识别|Formula image OCR/.test(button.textContent || ""));
        if (!item) throw new Error("Missing OCR menu item");
        item.click();
      })()`);

      const initialState = await waitForEvaluation(`(() => {
        const dialog = document.querySelector(".ocr-dialog");
        const card = dialog?.querySelector(".ocr-storage-location");
        const code = card?.querySelector("code");
        const buttons = [...(card?.querySelectorAll("button") ?? [])];
        const dialogRect = dialog?.getBoundingClientRect();
        const cardRect = card?.getBoundingClientRect();
        return {
          ready:
            Boolean(dialog && card && code && dialogRect && cardRect) &&
            code.textContent.includes("VisualTeX-OCR") &&
            buttons.length === 2 &&
            buttons.some((button) => /更改位置|Change location/.test(button.textContent || "")) &&
            buttons.some((button) => /打开文件夹|Open folder/.test(button.textContent || "")) &&
            cardRect.left >= dialogRect.left - 1 &&
            cardRect.right <= dialogRect.right + 1 &&
            card.scrollWidth <= card.clientWidth + 1,
          path: code?.textContent ?? "",
          labels: [...(card?.querySelectorAll("span") ?? [])].map((item) => item.textContent),
          buttonLabels: buttons.map((button) => button.textContent?.trim()),
          cardOverflow: card ? card.scrollWidth - card.clientWidth : -1,
        };
      })()`, "independent OCR storage card");

      await evaluate(`(() => {
        const card = document.querySelector(".ocr-storage-location");
        const button = [...(card?.querySelectorAll("button") ?? [])]
          .find((item) => /更改位置|Change location/.test(item.textContent || ""));
        if (!button) throw new Error("Missing OCR storage change button");
        button.click();
      })()`);

      const switchedState = await waitForEvaluation(`(() => {
        const card = document.querySelector(".ocr-storage-location");
        const code = card?.querySelector("code");
        const calls = window.__visualtexOcrStorageCalls || [];
        const configureCall = calls.find(
          (item) => item.command === "configure_ocr_storage_location",
        );
        const installCall = calls.find(
          (item) => item.command === "install_ocr_runtime",
        );
        return {
          ready:
            Boolean(card && code && configureCall && installCall) &&
            code.textContent.includes("D:\\\\OCR Data\\\\VisualTeX-OCR") &&
            configureCall.args?.selectedDirectory === "D:\\\\OCR Data" &&
            card.scrollWidth <= card.clientWidth + 1,
          path: code?.textContent ?? "",
          configureArgs: configureCall?.args ?? null,
          reinstalled: Boolean(installCall),
          calls: calls.map((item) => item.command),
        };
      })()`, "reset and switched OCR storage UI state");

      // Reopen the dialog with a deliberately delayed stale status request,
      // then change the location a second time before that old response returns.
      // The delayed C: result must never overwrite the new E: path.
      await evaluate(`(() => {
        window.__visualtexDelayNextRuntimeStatus = true;
        document.querySelector(".ocr-dialog-header .icon-button")?.click();
      })()`);
      await waitForEvaluation(`(() => ({
        ready: !document.querySelector(".ocr-dialog"),
      }))()`, "closed OCR dialog before second path change");
      await clickSelectorWithPointer(".menu-button");
      await evaluate(`(() => {
        const item = [...document.querySelectorAll('#app-main-menu [role="menuitem"]')]
          .find((button) => /图片公式识别|Formula image OCR/.test(button.textContent || ""));
        if (!item) throw new Error("Missing OCR menu item for second open");
        item.click();
      })()`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector(".ocr-dialog .ocr-storage-location")),
      }))()`, "reopened OCR dialog for stale path race");
      await sleep(60);
      await evaluate(`(() => {
        const button = [...document.querySelectorAll(".ocr-storage-location button")]
          .find((item) => /更改位置|Change location/.test(item.textContent || ""));
        if (!button) throw new Error("Missing second OCR storage change button");
        button.click();
      })()`);
      const secondSwitchState = await waitForEvaluation(`(() => {
        const card = document.querySelector(".ocr-storage-location");
        const code = card?.querySelector("code");
        const calls = window.__visualtexOcrStorageCalls || [];
        const configureCalls = calls.filter(
          (item) => item.command === "configure_ocr_storage_location",
        );
        return {
          ready:
            configureCalls.length === 2 &&
            code?.textContent.includes("E:\\\\Second OCR\\\\VisualTeX-OCR"),
          path: code?.textContent ?? "",
          configureArgs: configureCalls.at(-1)?.args ?? null,
        };
      })()`, "second OCR storage location change");
      await sleep(600);
      const stalePathRaceState = await evaluate(`(() => {
        const card = document.querySelector(".ocr-storage-location");
        const code = card?.querySelector("code");
        return {
          path: code?.textContent ?? "",
          installedReady: /本地 OCR 环境已就绪|Local OCR runtime ready/.test(
            document.querySelector(".ocr-runtime-summary strong")?.textContent || "",
          ),
          stalePathRejected: code?.textContent.includes("E:\\\\Second OCR\\\\VisualTeX-OCR"),
        };
      })()`);
      if (!stalePathRaceState.stalePathRejected || !stalePathRaceState.installedReady) {
        throw new Error(`A delayed stale runtime query overwrote the second path: ${JSON.stringify(stalePathRaceState)}`);
      }

      // Repeat the race for reset: a late old "installed" response must not
      // restore the ready card after reset_ocr_runtime returned uninstalled.
      await evaluate(`(() => {
        window.__visualtexDelayNextRuntimeStatus = true;
        document.querySelector(".ocr-dialog-header .icon-button")?.click();
      })()`);
      await waitForEvaluation(`(() => ({ ready: !document.querySelector(".ocr-dialog") }))()`, "closed OCR dialog before reset race");
      await clickSelectorWithPointer(".menu-button");
      await evaluate(`(() => {
        const item = [...document.querySelectorAll('#app-main-menu [role="menuitem"]')]
          .find((button) => /图片公式识别|Formula image OCR/.test(button.textContent || ""));
        if (!item) throw new Error("Missing OCR menu item for reset race");
        item.click();
      })()`);
      await waitForEvaluation(`(() => ({ ready: Boolean(document.querySelector(".ocr-runtime-details .is-danger")) }))()`, "OCR reset button");
      await sleep(60);
      await evaluate(`(() => {
        const button = document.querySelector(".ocr-runtime-details .is-danger");
        if (!button) throw new Error("Missing OCR reset button");
        button.click();
      })()`);
      const resetState = await waitForEvaluation(`(() => {
        const path = document.querySelector(".ocr-storage-location code")?.textContent ?? "";
        const heading = document.querySelector(".ocr-runtime-summary strong")?.textContent ?? "";
        return {
          ready:
            path.includes("E:\\\\Second OCR\\\\VisualTeX-OCR") &&
            /尚未安装 OCR 运行环境|OCR runtime is not installed/.test(heading) &&
            !document.querySelector(".ocr-runtime-details"),
          path,
          heading,
        };
      })()`, "OCR reset uninstalled state");
      await sleep(600);
      const staleResetRaceState = await evaluate(`(() => {
        const path = document.querySelector(".ocr-storage-location code")?.textContent ?? "";
        const heading = document.querySelector(".ocr-runtime-summary strong")?.textContent ?? "";
        return {
          path,
          heading,
          staleInstalledRejected:
            path.includes("E:\\\\Second OCR\\\\VisualTeX-OCR") &&
            /尚未安装 OCR 运行环境|OCR runtime is not installed/.test(heading) &&
            !document.querySelector(".ocr-runtime-details"),
        };
      })()`);
      if (!staleResetRaceState.staleInstalledRejected) {
        throw new Error(`A delayed stale installed status overwrote reset: ${JSON.stringify(staleResetRaceState)}`);
      }

      await evaluate(`(() => {
        const card = document.querySelector(".ocr-storage-location");
        const button = [...(card?.querySelectorAll("button") ?? [])]
          .find((item) => /打开文件夹|Open folder/.test(item.textContent || ""));
        if (!button) throw new Error("Missing OCR storage open button");
        button.click();
      })()`);
      const openState = await waitForEvaluation(`(() => {
        const calls = window.__visualtexOcrStorageCalls || [];
        return {
          ready: calls.some((item) => item.command === "open_ocr_storage_location"),
          calls: calls.map((item) => item.command),
        };
      })()`, "open OCR storage folder command");

      console.log(
        JSON.stringify({ initialState, switchedState, openState }, null, 2),
      );
      console.log("Targeted OCR independent storage UI regression passed");
      return;
    }

    if (scenario === "classic-panel-resize") {
      await evaluate(`(() => {
        const storageKey = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
        persisted.state = {
          ...(persisted.state || {}),
          editorLayout: "classic",
        };
        localStorage.setItem(storageKey, JSON.stringify(persisted));
        localStorage.setItem("visualtex-classic-tile-width", "300");
        localStorage.setItem("visualtex-classic-dock-height", "240");
      })()`);
      await client.send("Page.reload", { ignoreCache: true });

      const initial = await waitForEvaluation(`(() => {
        const tileHandle = document.querySelector(".classic-tile-resizer");
        const dockHandle = document.querySelector(".classic-dock-resizer");
        const tilePanel = document.querySelector(".classic-tile-toolbar");
        const dock = document.querySelector(".classic-bottom-dock");
        const toolbar = document.querySelector(".classic-bottom-toolbar");
        const strip = toolbar?.querySelector(".template-strip");
        const categorySection = toolbar?.querySelector(".toolbar-category-section");
        const rowCount = Number(toolbar?.dataset.toolbarRowCount || 0);
        const rowGrid = categorySection ?? strip;
        const computedRowCount = rowGrid
          ? getComputedStyle(rowGrid).gridTemplateRows.split(/\\s+/).filter(Boolean).length
          : 0;
        return {
          ready:
            Boolean(tileHandle && dockHandle && tilePanel && dock && toolbar) &&
            rowCount >= 2 &&
            computedRowCount === rowCount,
          tileWidth: tilePanel?.getBoundingClientRect().width ?? 0,
          dockHeight: dock?.getBoundingClientRect().height ?? 0,
          rowCount,
          computedRowCount,
        };
      })()`, "classic resizable panels");

      const tileTabGeometry = await waitForEvaluation(`(() => {
        const panel = document.querySelector('.classic-tile-toolbar');
        const custom = panel?.querySelector('[data-tile-category="custom"]');
        const common = panel?.querySelector('[data-tile-category="common"]');
        const collapse = panel?.querySelector('[data-formula-tile-collapse]');
        if (!(custom instanceof HTMLElement)
          || !(common instanceof HTMLElement)
          || !(collapse instanceof HTMLElement)) return { ready: false };
        const customRect = custom.getBoundingClientRect();
        const commonRect = common.getBoundingClientRect();
        const collapseRect = collapse.getBoundingClientRect();
        const sameRow = Math.max(
          Math.abs(customRect.top - commonRect.top),
          Math.abs(commonRect.top - collapseRect.top),
        ) <= 2;
        return {
          ready: sameRow && collapseRect.left >= commonRect.right - 1,
          sameRow,
          custom: { left: customRect.left, top: customRect.top, right: customRect.right },
          common: { left: commonRect.left, top: commonRect.top, right: commonRect.right },
          collapse: { left: collapseRect.left, top: collapseRect.top, right: collapseRect.right },
        };
      })()`, "classic tile collapse button to the right of Common");
      if (!tileTabGeometry.sameRow) {
        throw new Error(`Classic tile collapse button wrapped onto another row: ${JSON.stringify(tileTabGeometry)}`);
      }

      const handleCenter = async (selector) =>
        evaluate(`(() => {
          const handle = document.querySelector(${JSON.stringify(selector)});
          const rect = handle?.getBoundingClientRect();
          return rect
            ? { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 }
            : null;
        })()`);

      const tileStart = await handleCenter(".classic-tile-resizer");
      if (!tileStart) throw new Error("Missing classic tile resize handle");
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseMoved",
        x: tileStart.x,
        y: tileStart.y,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: tileStart.x,
        y: tileStart.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      for (let step = 1; step <= 6; step += 1) {
        await client.send("Input.dispatchMouseEvent", {
          type: "mouseMoved",
          x: tileStart.x - step * 20,
          y: tileStart.y,
          button: "left",
          buttons: 1,
        });
        await sleep(22);
      }
      const tileDuringDrag = await waitForEvaluation(`(() => {
        const panel = document.querySelector(".classic-tile-toolbar");
        const buttons = [...(panel?.querySelectorAll(".formula-tile-button") ?? [])];
        const width = panel?.getBoundingClientRect().width ?? 0;
        const firstRect = buttons[0]?.getBoundingClientRect();
        const singleColumn = Boolean(firstRect) && buttons.every((button) =>
          Math.abs(button.getBoundingClientRect().left - firstRect.left) < 1,
        );
        const previewsInside = buttons.every((button) => {
          const buttonRect = button.getBoundingClientRect();
          const visual = button.querySelector(".formula-tile-preview .ML__latex");
          const visualRect = visual?.getBoundingClientRect();
          return Boolean(
            visualRect &&
              visualRect.left >= buttonRect.left - 1 &&
              visualRect.right <= buttonRect.right + 1 &&
              visualRect.top >= buttonRect.top - 1 &&
              visualRect.bottom <= buttonRect.bottom + 1,
          );
        });
        return {
          ready:
            width >= ${initial.tileWidth + 90} &&
            buttons.length === 10 &&
            singleColumn &&
            previewsInside,
          width,
          buttonWidth: firstRect?.width ?? 0,
          buttonCount: buttons.length,
          singleColumn,
          previewsInside,
          resizeMode: document.body.dataset.workspaceResize ?? "",
        };
      })()`, "live tile panel resize without formula overflow");
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: tileStart.x - 120,
        y: tileStart.y,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });
      const tileAfterDrag = await waitForEvaluation(`(() => {
        const panel = document.querySelector(".classic-tile-toolbar");
        const width = panel?.getBoundingClientRect().width ?? 0;
        const stored = Number(localStorage.getItem("visualtex-classic-tile-width"));
        return {
          ready:
            width >= ${initial.tileWidth + 90} &&
            Math.abs(stored - width) < 3 &&
            !document.body.dataset.workspaceResize,
          width,
          stored,
        };
      })()`, "persisted tile panel resize");

      const growDockStart = await handleCenter(".classic-dock-resizer");
      if (!growDockStart) throw new Error("Missing classic dock resize handle");
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseMoved",
        x: growDockStart.x,
        y: growDockStart.y,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: growDockStart.x,
        y: growDockStart.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      for (let step = 1; step <= 10; step += 1) {
        await client.send("Input.dispatchMouseEvent", {
          type: "mouseMoved",
          x: growDockStart.x,
          y: growDockStart.y - step * 18,
          button: "left",
          buttons: 1,
        });
        await sleep(28);
      }
      const dockGrowDuringDrag = await waitForEvaluation(`(() => {
        const dock = document.querySelector(".classic-bottom-dock");
        const toolbar = document.querySelector(".classic-bottom-toolbar");
        const strip = toolbar?.querySelector(".template-strip");
        const categorySection = toolbar?.querySelector(".toolbar-category-section");
        const stripRect = strip?.getBoundingClientRect();
        const height = dock?.getBoundingClientRect().height ?? 0;
        const rowCount = Number(toolbar?.dataset.toolbarRowCount || 0);
        const computedRowCount = categorySection
          ? getComputedStyle(categorySection).gridTemplateRows.split(/\\s+/).filter(Boolean).length
          : 0;
        const buttonsVerticallyInside = Boolean(stripRect) &&
          [...(toolbar?.querySelectorAll(".template-button") ?? [])].every((button) => {
            const rect = button.getBoundingClientRect();
            return rect.top >= stripRect.top - 1 && rect.bottom <= stripRect.bottom + 1;
          });
        return {
          ready:
            height >= ${initial.dockHeight + 150} &&
            rowCount >= 6 &&
            computedRowCount === rowCount &&
            buttonsVerticallyInside &&
            document.body.dataset.workspaceResize === "dock",
          height,
          rowCount,
          computedRowCount,
          buttonsVerticallyInside,
          resizeMode: document.body.dataset.workspaceResize ?? "",
        };
      })()`, "live dock growth and category row expansion");
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: growDockStart.x,
        y: growDockStart.y - 180,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });

      const shrinkDockStart = await handleCenter(".classic-dock-resizer");
      if (!shrinkDockStart) throw new Error("Missing dock handle after growth");
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseMoved",
        x: shrinkDockStart.x,
        y: shrinkDockStart.y,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: shrinkDockStart.x,
        y: shrinkDockStart.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      for (let step = 1; step <= 15; step += 1) {
        await client.send("Input.dispatchMouseEvent", {
          type: "mouseMoved",
          x: shrinkDockStart.x,
          y: shrinkDockStart.y + step * 20,
          button: "left",
          buttons: 1,
        });
        await sleep(26);
      }
      const dockShrinkDuringDrag = await waitForEvaluation(`(() => {
        const dock = document.querySelector(".classic-bottom-dock");
        const toolbar = document.querySelector(".classic-bottom-toolbar");
        const strip = toolbar?.querySelector(".template-strip");
        const categorySection = toolbar?.querySelector(".toolbar-category-section");
        const stripRect = strip?.getBoundingClientRect();
        const height = dock?.getBoundingClientRect().height ?? 0;
        const rowCount = Number(toolbar?.dataset.toolbarRowCount || 0);
        const rowGrid = categorySection ?? strip;
        const computedRowCount = rowGrid
          ? getComputedStyle(rowGrid).gridTemplateRows.split(/\\s+/).filter(Boolean).length
          : 0;
        const buttonsVerticallyInside = Boolean(stripRect) &&
          [...(toolbar?.querySelectorAll(".template-button") ?? [])].every((button) => {
            const rect = button.getBoundingClientRect();
            return rect.top >= stripRect.top - 1 && rect.bottom <= stripRect.bottom + 1;
          });
        return {
          ready:
            height <= 134 &&
            rowCount === 1 &&
            computedRowCount === rowCount &&
            buttonsVerticallyInside &&
            document.body.dataset.workspaceResize === "dock",
          height,
          rowCount,
          computedRowCount,
          buttonsVerticallyInside,
        };
      })()`, "live dock shrink to one complete toolbar row");
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: shrinkDockStart.x,
        y: shrinkDockStart.y + 300,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });

      const restoreTwoRowsStart = await handleCenter(".classic-dock-resizer");
      if (!restoreTwoRowsStart) throw new Error("Missing dock handle after one-row shrink");
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseMoved",
        x: restoreTwoRowsStart.x,
        y: restoreTwoRowsStart.y,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: restoreTwoRowsStart.x,
        y: restoreTwoRowsStart.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      for (let step = 1; step <= 4; step += 1) {
        await client.send("Input.dispatchMouseEvent", {
          type: "mouseMoved",
          x: restoreTwoRowsStart.x,
          y: restoreTwoRowsStart.y - step * 18,
          button: "left",
          buttons: 1,
        });
        await sleep(24);
      }
      const dockRestoreTwoRows = await waitForEvaluation(`(() => {
        const toolbar = document.querySelector(".classic-bottom-toolbar");
        const strip = toolbar?.querySelector(".template-strip");
        const categorySection = toolbar?.querySelector(".toolbar-category-section");
        const rowCount = Number(toolbar?.dataset.toolbarRowCount || 0);
        const rowGrid = categorySection ?? strip;
        const computedRowCount = rowGrid
          ? getComputedStyle(rowGrid).gridTemplateRows.split(/\\s+/).filter(Boolean).length
          : 0;
        return {
          ready:
            rowCount === 2 &&
            computedRowCount === rowCount &&
            document.body.dataset.workspaceResize === "dock",
          rowCount,
          computedRowCount,
        };
      })()`, "restored two-row toolbar for matrix tools");
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: restoreTwoRowsStart.x,
        y: restoreTwoRowsStart.y - 72,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });

      await evaluate(`document.querySelector(
        '.classic-bottom-toolbar .toolbar-tab[data-category="matrix"]',
      )?.click()`);
      const compactMatrixState = await waitForEvaluation(`(() => {
        const toolbar = document.querySelector(".classic-bottom-toolbar");
        const strip = toolbar?.querySelector(".template-strip");
        const builder = toolbar?.querySelector(".matrix-builder");
        const grid = toolbar?.querySelector(".matrix-size-grid");
        const delimiterOptions = toolbar?.querySelector(".matrix-delimiter-options");
        const delimiterButtons = [...(delimiterOptions?.querySelectorAll("button") ?? [])];
        const stripRect = strip?.getBoundingClientRect();
        const builderRect = builder?.getBoundingClientRect();
        const gridRect = grid?.getBoundingClientRect();
        const delimiterRect = delimiterOptions?.getBoundingClientRect();
        const maxDelimiterRight = delimiterButtons.reduce(
          (right, button) => Math.max(right, button.getBoundingClientRect().right),
          0,
        );
        const rowCount = Number(toolbar?.dataset.toolbarRowCount || 0);
        return {
          ready:
            rowCount === 2 &&
            Boolean(stripRect && builderRect && gridRect) &&
            builderRect.top >= stripRect.top - 1 &&
            builderRect.bottom <= stripRect.bottom + 1 &&
            gridRect.top >= builderRect.top - 1 &&
            gridRect.bottom <= builderRect.bottom + 1 &&
            delimiterButtons.length === 6 &&
            Boolean(delimiterRect) &&
            maxDelimiterRight <= gridRect.left - 1 &&
            delimiterRect.right <= gridRect.left - 1,
          rowCount,
          stripHeight: stripRect?.height ?? 0,
          builderHeight: builderRect?.height ?? 0,
          gridHeight: gridRect?.height ?? 0,
          delimiterCount: delimiterButtons.length,
          delimiterRight: maxDelimiterRight,
          gridLeft: gridRect?.left ?? 0,
        };
      })()`, "compact two-row matrix toolbar");

      const growMatrixDockStart = await handleCenter(".classic-dock-resizer");
      if (!growMatrixDockStart) throw new Error("Missing dock handle for matrix growth check");
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseMoved",
        x: growMatrixDockStart.x,
        y: growMatrixDockStart.y,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: growMatrixDockStart.x,
        y: growMatrixDockStart.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      for (let step = 1; step <= 4; step += 1) {
        await client.send("Input.dispatchMouseEvent", {
          type: "mouseMoved",
          x: growMatrixDockStart.x,
          y: growMatrixDockStart.y - step * 20,
          button: "left",
          buttons: 1,
        });
        await sleep(26);
      }
      const grownMatrixState = await waitForEvaluation(`(() => {
        const toolbar = document.querySelector(".classic-bottom-toolbar");
        const builder = toolbar?.querySelector(".matrix-builder");
        const grid = toolbar?.querySelector(".matrix-size-grid");
        const delimiterOptions = toolbar?.querySelector(".matrix-delimiter-options");
        const delimiterButtons = [...(delimiterOptions?.querySelectorAll("button") ?? [])];
        const builderRect = builder?.getBoundingClientRect();
        const gridRect = grid?.getBoundingClientRect();
        const maxDelimiterRight = delimiterButtons.reduce(
          (right, button) => Math.max(right, button.getBoundingClientRect().right),
          0,
        );
        const rowCount = Number(toolbar?.dataset.toolbarRowCount || 0);
        return {
          ready:
            rowCount >= 3 &&
            Boolean(builderRect && gridRect) &&
            gridRect.height >= ${compactMatrixState.gridHeight + 24} &&
            gridRect.height <= 153 &&
            gridRect.top >= builderRect.top - 1 &&
            gridRect.bottom <= builderRect.bottom + 1 &&
            delimiterButtons.length === 6 &&
            maxDelimiterRight <= gridRect.left - 1,
          rowCount,
          builderHeight: builderRect?.height ?? 0,
          gridHeight: gridRect?.height ?? 0,
          delimiterCount: delimiterButtons.length,
          delimiterRight: maxDelimiterRight,
          gridLeft: gridRect?.left ?? 0,
        };
      })()`, "matrix picker grows with toolbar height");
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: growMatrixDockStart.x,
        y: growMatrixDockStart.y - 80,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });

      await evaluate(`document.querySelector('[data-classic-bottom-view="source"]')?.click()`);
      const sourceBefore = await waitForEvaluation(`(() => {
        const slot = document.querySelector(".classic-source-pane-slot");
        const editor = slot?.querySelector(".cm-editor");
        const height = slot?.getBoundingClientRect().height ?? 0;
        return {
          ready: Boolean(slot && editor) && height > 80,
          height,
          editorHeight: editor?.getBoundingClientRect().height ?? 0,
        };
      })()`, "classic source panel after toolbar resize");

      const sourceDockStart = await handleCenter(".classic-dock-resizer");
      if (!sourceDockStart) throw new Error("Missing dock handle in source view");
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseMoved",
        x: sourceDockStart.x,
        y: sourceDockStart.y,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: sourceDockStart.x,
        y: sourceDockStart.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      for (let step = 1; step <= 4; step += 1) {
        await client.send("Input.dispatchMouseEvent", {
          type: "mouseMoved",
          x: sourceDockStart.x,
          y: sourceDockStart.y - step * 18,
          button: "left",
          buttons: 1,
        });
        await sleep(28);
      }
      const sourceDuringDrag = await waitForEvaluation(`(() => {
        const slot = document.querySelector(".classic-source-pane-slot");
        const editor = slot?.querySelector(".cm-editor");
        const height = slot?.getBoundingClientRect().height ?? 0;
        return {
          ready:
            height >= ${sourceBefore.height + 50} &&
            (editor?.getBoundingClientRect().height ?? 0) >= ${sourceBefore.editorHeight + 45} &&
            document.body.dataset.workspaceResize === "dock",
          height,
          editorHeight: editor?.getBoundingClientRect().height ?? 0,
        };
      })()`, "live source editor resize");
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: sourceDockStart.x,
        y: sourceDockStart.y - 72,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });

      const persistedState = await waitForEvaluation(`(() => {
        const tile = Number(localStorage.getItem("visualtex-classic-tile-width"));
        const dock = Number(localStorage.getItem("visualtex-classic-dock-height"));
        return {
          ready:
            tile >= ${initial.tileWidth + 90} &&
            dock >= ${sourceBefore.height + 72} &&
            !document.body.dataset.workspaceResize,
          tile,
          dock,
        };
      })()`, "persisted classic panel dimensions");

      console.log(
        JSON.stringify(
          {
            initial,
            tileDuringDrag,
            tileAfterDrag,
            dockGrowDuringDrag,
            dockShrinkDuringDrag,
            dockRestoreTwoRows,
            compactMatrixState,
            grownMatrixState,
            sourceBefore,
            sourceDuringDrag,
            persistedState,
          },
          null,
          2,
        ),
      );
      console.log("Targeted classic panel resize regression passed");
      return;
    }

    if (scenario === "toolbar-postfix") {
      await evaluate(`(() => {
        const storageKey = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
        const first = { id: crypto.randomUUID(), latex: "a" };
        const second = { id: crypto.randomUUID(), latex: "b+c" };
        persisted.state = {
          ...(persisted.state || {}),
          lines: [first, second],
          activeLineId: first.id,
          editorLayout: "classic",
        };
        localStorage.setItem(storageKey, JSON.stringify(persisted));
      })()`);
      await client.send("Page.reload", { ignoreCache: true });
      const startupFocusState = await waitForEvaluation(`(() => {
        const fields = [...document.querySelectorAll("math-field")];
        const field = fields.at(-1);
        const activeLineId = document.querySelector(".multi-line-editor")?.dataset.activeLineId ?? "";
        const lastLineId = document.querySelector(".formula-line:last-child")?.dataset.lineId ?? "";
        return {
          ready:
            fields.length === 2 &&
            Boolean(field?.hasFocus?.()) &&
            field?.position === field?.lastOffset &&
            activeLineId === lastLineId,
          fieldCount: fields.length,
          focused: field?.hasFocus?.() ?? false,
          position: field?.position ?? -1,
          lastOffset: field?.lastOffset ?? -1,
          activeLineId,
          lastLineId,
        };
      })()`, "startup focus on the end of the last formula line");

      await waitForEvaluation(`(() => ({
        ready: Boolean(
          document.querySelector(".classic-bottom-toolbar.is-horizontal") &&
          document.querySelector(".classic-bottom-toolbar .template-strip"),
        ),
      }))()`, "classic horizontal formula toolbar");

      const toolbarPreviewVisibilityState = await waitForEvaluation(`(() => {
        const section = document.querySelector(
          '.classic-bottom-toolbar .toolbar-category-section[data-toolbar-category-section="common"]',
        );
        const buttons = [...(section?.querySelectorAll('.template-button') ?? [])];
        const sectionStyle = section ? getComputedStyle(section) : null;
        const previews = buttons.map((button) => {
          const host = button.querySelector('.math-preview');
          const content = host?.querySelector('.math-preview-fit-content');
          const latex = host?.querySelector('.ML__latex');
          const hostStyle = host ? getComputedStyle(host) : null;
          const contentStyle = content ? getComputedStyle(content) : null;
          const latexStyle = latex ? getComputedStyle(latex) : null;
          const buttonStyle = getComputedStyle(button);
          const buttonBounds = button.getBoundingClientRect();
          const hostBounds = host?.getBoundingClientRect();
          const latexBounds = latex?.getBoundingClientRect();
          const latexColor = latexStyle?.color ?? '';
          return {
            commandId: button.dataset.commandId ?? '',
            visible: Boolean(
              host && content && latex &&
              hostBounds && hostBounds.width > 4 && hostBounds.height > 4 &&
              latexBounds && latexBounds.width > 0 && latexBounds.height > 0 &&
              hostStyle?.display !== 'none' &&
              hostStyle?.visibility !== 'hidden' &&
              contentStyle?.display !== 'none' &&
              contentStyle?.visibility !== 'hidden' &&
              latexStyle?.visibility !== 'hidden' &&
              Number.parseFloat(hostStyle?.opacity || '1') > 0.1 &&
              Number.parseFloat(contentStyle?.opacity || '1') > 0.1 &&
              Number.parseFloat(latexStyle?.opacity || '1') > 0.1 &&
              latexColor !== 'transparent' &&
              latexColor !== 'rgba(0, 0, 0, 0)'
            ),
            button: {
              width: buttonBounds.width,
              height: buttonBounds.height,
              display: buttonStyle.display,
              alignItems: buttonStyle.alignItems,
              justifyContent: buttonStyle.justifyContent,
            },
            host: hostBounds ? { width: hostBounds.width, height: hostBounds.height } : null,
            latex: latexBounds ? { width: latexBounds.width, height: latexBounds.height } : null,
            hostDisplay: hostStyle?.display ?? '',
            hostWidth: hostStyle?.width ?? '',
            hostHeight: hostStyle?.height ?? '',
            hostFlex: hostStyle?.flex ?? '',
            hostColor: hostStyle?.color ?? '',
            latexColor,
            hostContain: hostStyle?.contain ?? '',
            contentTransform: contentStyle?.transform ?? '',
          };
        });
        const invalid = previews.filter((preview) => !preview.visible);
        const windowsSafeCompositing = previews.every(
          (preview) => preview.hostContain === 'none',
        ) && sectionStyle?.transform === 'none' && sectionStyle?.contain === 'none';
        return {
          ready:
            buttons.length > 0 &&
            invalid.length === 0 &&
            windowsSafeCompositing,
          buttonCount: buttons.length,
          invalid: invalid.slice(0, 3),
          invalidCount: invalid.length,
          first: previews[0] ?? null,
          sectionTransform: sectionStyle?.transform ?? '',
          sectionContain: sectionStyle?.contain ?? '',
          windowsSafeCompositing,
        };
      })()`, "visible classic formula tool previews");

      const setActiveField = async (latex) => {
        await waitForEvaluation(`(() => {
          const fields = [...document.querySelectorAll("math-field")];
          const field = fields.at(-1);
          if (!field?.isConnected) return { ready: false };
          field.setValue(${JSON.stringify(latex)}, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceAll",
            selectionMode: "after",
            silenceNotifications: true,
          });
          field.selection = {
            ranges: [[field.lastOffset, field.lastOffset]],
            direction: "none",
          };
          field.position = field.lastOffset;
          field.focus();
          field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
          field.dispatchEvent(new InputEvent("input", {
            bubbles: true,
            composed: true,
            inputType: "insertText",
          }));
          return {
            ready: field.value === ${JSON.stringify(latex)} && field.hasFocus(),
          };
        })()`, `set active formula to ${latex}`);
        await sleep(80);
      };

      const clickToolbarCommand = async (category, commandId) => {
        await evaluate(`document.querySelector(
          '.classic-bottom-toolbar .toolbar-tab[data-category="${category}"]',
        )?.click()`);
        await waitForEvaluation(`(() => ({
          ready: Boolean(document.querySelector(
            '.classic-bottom-toolbar [data-command-id="${commandId}"]',
          )),
        }))()`, `toolbar command ${commandId}`);
        await evaluate(`document.querySelector(
          '.classic-bottom-toolbar [data-command-id="${commandId}"]',
        ).click()`);
        await sleep(120);
      };

      await setActiveField("x");
      await clickToolbarCommand("structure", "upper-script");
      const upperScriptState = await waitForEvaluation(`(() => {
        const field = [...document.querySelectorAll("math-field")].at(-1);
        return {
          ready: field?.value === "x^{\\\\placeholder{}}",
          value: field?.value ?? "",
          selection: field?.selection ?? null,
        };
      })()`, "upper-script targets the preceding character");

      await setActiveField("x");
      await clickToolbarCommand("structure", "scripts");
      const scriptsState = await waitForEvaluation(`(() => {
        const field = [...document.querySelectorAll("math-field")].at(-1);
        return {
          ready:
            field?.value ===
              "x_{\\\\placeholder{}}^{\\\\placeholder{}}",
          value: field?.value ?? "",
          selection: field?.selection ?? null,
        };
      })()`, "combined upper/lower scripts target the preceding character");

      await setActiveField("y");
      await clickToolbarCommand("structure", "lower-script");
      const lowerScriptState = await waitForEvaluation(`(() => {
        const field = [...document.querySelectorAll("math-field")].at(-1);
        return {
          ready: field?.value === "y_{\\\\placeholder{}}",
          value: field?.value ?? "",
          selection: field?.selection ?? null,
        };
      })()`, "lower-script targets the preceding character");

      await setActiveField("a+b");
      await clickToolbarCommand("structure", "dotaccent");
      const dotAccentState = await waitForEvaluation(`(() => {
        const field = [...document.querySelectorAll("math-field")].at(-1);
        return {
          ready: field?.value === "a+\\\\dot{b}",
          value: field?.value ?? "",
        };
      })()`, "dot accent wraps the preceding character");

      await setActiveField("\\alpha");
      await clickToolbarCommand("matrix", "boldsymbol");
      const fontWrapperState = await waitForEvaluation(`(() => {
        const field = [...document.querySelectorAll("math-field")].at(-1);
        return {
          ready: field?.value === "\\\\mathbf{\\\\alpha}",
          value: field?.value ?? "",
        };
      })()`, "font wrapper targets the preceding symbol");

      await setActiveField("x");
      await clickToolbarCommand("structure", "overset");
      const oversetState = await waitForEvaluation(`(() => {
        const field = [...document.querySelectorAll("math-field")].at(-1);
        return {
          ready: field?.value === "\\\\overset{\\\\placeholder{}}{x}",
          value: field?.value ?? "",
        };
      })()`, "overset keeps the preceding character as its base");

      await setActiveField("+");
      await clickToolbarCommand("structure", "dotaccent");
      const operatorFallbackState = await waitForEvaluation(`(() => {
        const field = [...document.querySelectorAll("math-field")].at(-1);
        return {
          ready: field?.value === "+\\\\dot{\\\\placeholder{}}",
          value: field?.value ?? "",
        };
      })()`, "decorator does not consume a preceding operator");

      await evaluate(`document.querySelector(
        '.classic-bottom-toolbar .toolbar-tab[data-category="relation"]',
      )?.click()`);
      const horizontalWheelState = await waitForEvaluation(`(() => {
        const strip = document.querySelector(
          ".classic-bottom-toolbar .template-strip",
        );
        if (!strip) return { ready: false };
        strip.style.width = "320px";
        strip.style.maxWidth = "320px";
        strip.scrollLeft = 0;
        const before = strip.scrollLeft;
        const event = new WheelEvent("wheel", {
          deltaY: 180,
          deltaMode: WheelEvent.DOM_DELTA_PIXEL,
          bubbles: true,
          cancelable: true,
        });
        const dispatchResult = strip.dispatchEvent(event);
        return {
          ready: strip.scrollWidth > strip.clientWidth && strip.scrollLeft > before,
          before,
          after: strip.scrollLeft,
          scrollWidth: strip.scrollWidth,
          clientWidth: strip.clientWidth,
          defaultPrevented: event.defaultPrevented,
          dispatchResult,
        };
      })()`, "mouse wheel scrolls the classic toolbar horizontally");

      console.log(JSON.stringify({
        startupFocusState,
        toolbarPreviewVisibilityState,
        upperScriptState,
        scriptsState,
        lowerScriptState,
        dotAccentState,
        fontWrapperState,
        oversetState,
        operatorFallbackState,
        horizontalWheelState,
      }, null, 2));
      console.log("Targeted postfix toolbar and startup focus regression passed");
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
            rows[0]?.count === expectedOrder.length &&
            labelsFit &&
            containerHeight <= 48,
          actualOrder,
          gridColumnCount,
          rows,
          labelsFit,
          containerHeight,
        };
      })()`, "single-row scrolling toolbar tabs");

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
          const strip = document.querySelector(".template-strip");
          const buttons = [...document.querySelectorAll(
            ".template-strip > .template-button",
          )];
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
          const equalHeights = bounds.every(
            (rect) => Math.abs(rect.height - 54) <= 1,
          );
          const equalWidths = bounds.every(
            (rect) =>
              bounds[0] && Math.abs(rect.width - bounds[0].width) <= 1,
          );
          const fourColumns =
            firstRow.length === Math.min(4, buttons.length) &&
            firstRow.every(
              (rect, index) =>
                index === 0 || rect.left > firstRow[index - 1].left,
            );
          const stripStyle = strip ? getComputedStyle(strip) : null;
          const gridColumnCount = stripStyle?.gridTemplateColumns
            .split(" ")
            .filter(Boolean).length ?? 0;
          return {
            ready:
              Boolean(strip) &&
              buttons.length > 0 &&
              directContentOnly &&
              previewsVisible &&
              ariaLabelsPresent &&
              equalHeights &&
              equalWidths &&
              fourColumns &&
              gridColumnCount === 4 &&
              strip.scrollWidth <= strip.clientWidth + 1,
            category: ${JSON.stringify(category)},
            buttonCount: buttons.length,
            firstRowCount: firstRow.length,
            buttonWidth: bounds[0]?.width ?? 0,
            buttonHeight: bounds[0]?.height ?? 0,
            gridTemplateColumns: stripStyle?.gridTemplateColumns ?? "",
            gridColumnCount,
            directContentOnly,
            previewsVisible,
            firstPreviewState: previewStates[0] ?? null,
            ariaLabelsPresent,
            equalHeights,
            equalWidths,
            horizontalOverflow:
              strip ? strip.scrollWidth - strip.clientWidth : -1,
          };
        })()`, `compact four-column toolbar category: ${category}`);
        categoryStates.push(state);
      }

      await evaluate(`document.querySelector(
        '.toolbar-tab[data-category="matrix"]',
      ).click()`);
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
            buttons.length === 3 &&
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
          const buttons = [...document.querySelectorAll(
            ".template-strip > .template-button",
          )];
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
              autoFit: button.classList.contains("is-auto-fit"),
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
              preview.fillRatio < 0.72 ||
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
            autoFit: button?.classList.contains("is-auto-fit") ?? false,
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
        const ordinaryButtons = [...document.querySelectorAll(
          ".template-strip > .template-button:not([data-command-id='commutator']):not([data-command-id='anticommutator'])",
        )];
        const ordinaryUnchanged = ordinaryButtons.every(
          (button) =>
            !button.classList.contains("is-auto-fit") &&
            button.querySelector(".math-preview")?.dataset.fit === "none",
        );
        return {
          ready:
            [commutator, anticommutator].every(
              (preview) =>
                preview.autoFit &&
                preview.fitReady &&
                preview.inside &&
                preview.fillRatio >= 0.72 &&
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
          "times",
          "div",
          "in",
          "subset",
          "rightarrow",
          "forall",
          "exists",
        ];
        const buttons = [...document.querySelectorAll(
          ".template-strip > .template-button",
        )];
        const actualIds = buttons.map((button) => button.dataset.commandId ?? "");
        const storedIds = JSON.parse(
          localStorage.getItem("visualtex-common-toolbar-command-ids-v2") || "[]",
        );
        const missingIds = expectedIds.filter((id) => !actualIds.includes(id));
        return {
          ready:
            actualIds.length === 45 &&
            JSON.stringify(actualIds) === JSON.stringify(expectedIds) &&
            JSON.stringify(storedIds) === JSON.stringify(expectedIds) &&
            missingIds.length === 0 &&
            !actualIds.includes("notin") &&
            !actualIds.includes("leftarrow"),
          count: actualIds.length,
          actualIds,
          storedIds,
          missingIds,
        };
      })()`, "expanded common formula collection");

      await clearField();
      await evaluate(`(() => {
        document.querySelector('[data-command-id="times"]')?.click();
        document.querySelector('[data-command-id="div"]')?.click();
        return true;
      })()`);
      const arithmeticOperatorState = await waitForEvaluation(`(() => {
        const value = document.querySelector("math-field")?.value ?? "";
        return {
          ready: value.includes("\\\\times") && value.includes("\\\\div"),
          value,
        };
      })()`, "common multiplication and division insertion");

      console.log(
        JSON.stringify(
          {
            tabLayoutState,
            categoryStates,
            matrixDelimiterState,
            adaptiveCategoryStates,
            physicsExceptionState,
            calculusPreviewState,
            commonContentsState,
            arithmeticOperatorState,
          },
          null,
          2,
        ),
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

      await evaluate(`document.querySelector(
        '[data-tile-category="custom"]',
      ).click()`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector('[data-open-custom-symbol-designer]')),
      }))()`, "custom symbol designer trigger");
      await evaluate(`document.querySelector('[data-open-custom-symbol-designer]').click()`);
      const customSymbolDesignerVisualState = await waitForEvaluation(`(() => {
        const dialog = document.querySelector('[data-custom-symbol-designer]');
        const stage = dialog?.querySelector('.custom-symbol-designer-stage');
        const workspace = dialog?.querySelector('.custom-symbol-designer-workspace');
        const paper = dialog?.querySelector('[data-custom-symbol-canvas-paper]');
        const panelHeader = dialog?.querySelector('.custom-symbol-designer-panel > header');
        const sidebars = [...(dialog?.querySelectorAll('.custom-symbol-designer-sidebar') ?? [])];
        const rootStyle = getComputedStyle(document.documentElement);
        const stageStyle = stage ? getComputedStyle(stage) : null;
        const workspaceStyle = workspace ? getComputedStyle(workspace) : null;
        const stageBounds = stage?.getBoundingClientRect();
        const paperBounds = paper?.getBoundingClientRect();
        const dialogBounds = dialog?.getBoundingClientRect();
        const panelHeaderFontSize = Number.parseFloat(
          panelHeader ? getComputedStyle(panelHeader).fontSize : '0',
        );
        const workspaceFill = workspaceStyle?.fill ?? '';
        const sunken = rootStyle.getPropertyValue('--bg-sunken').trim();
        return {
          ready: Boolean(
            dialog && stage && workspace && paper &&
            sunken &&
            workspaceFill &&
            workspaceFill !== 'black' &&
            workspaceFill !== 'rgb(0, 0, 0)' &&
            stageBounds && stageBounds.width > 300 && stageBounds.height > 300 &&
            paperBounds && paperBounds.width >= stageBounds.width * 0.45 &&
            paperBounds.height >= stageBounds.height * 0.45 &&
            dialogBounds && dialogBounds.left >= -1 && dialogBounds.top >= -1 &&
            dialogBounds.right <= innerWidth + 1 && dialogBounds.bottom <= innerHeight + 1 &&
            sidebars.length === 2 &&
            sidebars.every((sidebar) => sidebar.scrollWidth <= sidebar.clientWidth + 1) &&
            panelHeaderFontSize >= 10
          ),
          sunken,
          workspaceFill,
          stageBackground: stageStyle?.backgroundColor ?? '',
          stage: stageBounds ? { width: stageBounds.width, height: stageBounds.height } : null,
          paper: paperBounds ? { width: paperBounds.width, height: paperBounds.height } : null,
          dialog: dialogBounds ? {
            left: dialogBounds.left,
            top: dialogBounds.top,
            right: dialogBounds.right,
            bottom: dialogBounds.bottom,
          } : null,
          panelHeaderFontSize,
          sidebarOverflow: sidebars.map((sidebar) => ({
            clientWidth: sidebar.clientWidth,
            scrollWidth: sidebar.scrollWidth,
            overflow: sidebar.scrollWidth - sidebar.clientWidth,
          })),
        };
      })()`, "balanced custom symbol designer on Windows");
      await evaluate(`document.querySelector('[data-add-custom-symbol-material]')?.click()`);
      const designerLayerBeforeDelete = await waitForEvaluation(`(() => ({
        ready: document.querySelectorAll('[data-custom-symbol-layer]').length > 0,
        count: document.querySelectorAll('[data-custom-symbol-layer]').length,
      }))()`, "custom-symbol designer layer before keyboard delete");
      await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', bubbles: true, cancelable: true }))`);
      const designerLayerAfterDelete = await waitForEvaluation(`(() => ({
        ready: document.querySelectorAll('[data-custom-symbol-layer]').length === 0,
        count: document.querySelectorAll('[data-custom-symbol-layer]').length,
      }))()`, "custom-symbol designer keyboard Delete");
      if (designerLayerBeforeDelete.count <= designerLayerAfterDelete.count) {
        throw new Error(`Delete key did not remove the selected custom-symbol layer: ${JSON.stringify({ designerLayerBeforeDelete, designerLayerAfterDelete })}`);
      }
      await evaluate(`document.querySelector(
        '[data-custom-symbol-designer] button[aria-label="关闭"], [data-custom-symbol-designer] button[aria-label="Close"]',
      )?.click()`);
      await waitForEvaluation(`(() => ({
        ready: !document.querySelector('[data-custom-symbol-designer]'),
      }))()`, "closed custom symbol designer");

      // Registered custom symbols must expose a real, visible delete control in
      // the toolbar. The old top-right glyph could degrade into a '?' and there
      // was no reliable deletion path from the toolbar itself.
      await evaluate(`(() => {
        const key = 'visualtex.custom-symbols.v1';
        const value = {
          version: 1,
          symbols: [{
            id: 'qa-custom-symbol-delete',
            command: 'qaacustomglyph',
            name: 'QA custom glyph',
            role: 'ordinary',
            limitsBehavior: 'auto',
            metrics: { widthEm: 1, ascentEm: 1, descentEm: 0.1 },
            artwork: {
              shapes: [{
                kind: 'line', x1: 80, y1: 500, x2: 920, y2: 500,
                fill: false, strokeWidth: 80, lineCap: 'round',
              }],
            },
            ommlFallback: null,
            createdAt: Date.now(),
            updatedAt: Date.now(),
          }],
        };
        const oldValue = localStorage.getItem(key);
        const newValue = JSON.stringify(value);
        localStorage.setItem(key, newValue);
        window.dispatchEvent(new StorageEvent('storage', { key, oldValue, newValue }));
      })()`);
      const registeredDeleteState = await waitForEvaluation(`(() => {
        const item = document.querySelector('[data-registered-custom-symbol-toolbar-item="qa-custom-symbol-delete"]');
        const button = document.querySelector('[data-delete-registered-custom-symbol-toolbar="qa-custom-symbol-delete"]');
        const label = button?.textContent?.trim() ?? '';
        return {
          ready: Boolean(item && button && /删除|Delete/.test(label)),
          label,
          hasSvgIcon: Boolean(button?.querySelector('svg')),
        };
      })()`, "visible registered custom-symbol delete button");
      if (!registeredDeleteState.hasSvgIcon) {
        throw new Error(`Registered custom-symbol delete control has no SVG trash icon: ${JSON.stringify(registeredDeleteState)}`);
      }
      await evaluate(`document.querySelector('[data-delete-registered-custom-symbol-toolbar="qa-custom-symbol-delete"]')?.click()`);
      const firstDeleteClick = await waitForEvaluation(`(() => ({
        ready: document.querySelector('[data-delete-registered-custom-symbol-toolbar="qa-custom-symbol-delete"]')?.classList.contains('is-confirming') === true,
        stillPresent: Boolean(document.querySelector('[data-registered-custom-symbol-toolbar-item="qa-custom-symbol-delete"]')),
      }))()`, "custom-symbol delete confirmation");
      if (!firstDeleteClick.stillPresent) {
        throw new Error('Registered custom symbol was deleted without the required confirmation click');
      }
      await evaluate(`document.querySelector('[data-delete-registered-custom-symbol-toolbar="qa-custom-symbol-delete"]')?.click()`);
      await waitForEvaluation(`(() => {
        const stored = JSON.parse(localStorage.getItem('visualtex.custom-symbols.v1') || '{"symbols":[]}');
        return {
          ready: !document.querySelector('[data-registered-custom-symbol-toolbar-item="qa-custom-symbol-delete"]') &&
            !stored.symbols?.some?.((symbol) => symbol.id === 'qa-custom-symbol-delete'),
        };
      })()`, "delete registered custom symbol from toolbar");

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
          ".formula-tile-button.is-custom",
        );
        const stored = JSON.parse(
          localStorage.getItem("visualtex-custom-formula-tiles") || "{}",
        );
        const tiles = Array.isArray(stored.tiles) ? stored.tiles : [];
        const host = button?.querySelector(".formula-tile-preview");
        return {
          ready:
            Boolean(button) &&
            tiles.length === 1 &&
            tiles[0].latex.includes("\\\\frac{-b\\\\pm\\\\sqrt{b^2-4ac}}{2a}") &&
            host?.dataset.fitReady === "true",
          stored,
          latex: button?.dataset.formulaTileLatex ?? "",
          fitReady: host?.dataset.fitReady ?? "missing",
        };
      })()`, "persisted selected-line custom formula tile");

      await evaluate(`(() => {
        const button = document.querySelector(
          ".formula-tile-button.is-custom",
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
          localStorage.getItem("visualtex-custom-formula-tiles") || "{}",
        );
        const tiles = Array.isArray(stored.tiles) ? stored.tiles : [];
        return {
          ready:
            Boolean(menu && deleteButton) &&
            field?.value ===
              window.__visualtexFieldValueBeforeTileContextMenu &&
            tiles.length === 1,
          fieldValueBefore:
            window.__visualtexFieldValueBeforeTileContextMenu ?? "",
          fieldValue: field?.value ?? "",
          stored,
          menuText: deleteButton?.textContent?.trim() ?? "",
        };
      })()`, "custom tile right-click menu without insertion conflict");

      await evaluate(`document.querySelector(
        ".formula-tile-context-menu .formula-hotkey-context-action.is-danger",
      ).click()`);
      const deletedCustomTileState = await waitForEvaluation(`(() => {
        const stored = JSON.parse(
          localStorage.getItem("visualtex-custom-formula-tiles") || "{}",
        );
        const tiles = Array.isArray(stored.tiles) ? stored.tiles : [];
        return {
          ready:
            !document.querySelector(".formula-tile-button.is-custom") &&
            !document.querySelector(".formula-tile-context-menu") &&
            Boolean(document.querySelector(".custom-formula-section-empty")) &&
            tiles.length === 0,
          stored,
          emptyVisible: Boolean(document.querySelector(".custom-formula-section-empty")),
        };
      })()`, "right-click custom tile deletion persistence");

      console.log(
        JSON.stringify(
          {
            commonTilesState,
            insertedTileState,
            customSymbolDesignerVisualState,
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
          editorLayout: "standard",
          sidebarOpen: true,
          sourceOpen: false,
          latexCodeFormat: "raw",
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
          const operator = selected?.closest(".ML__op-group");
          const operatorBounds = operator?.getBoundingClientRect();
          const operatorRegion =
            selectedBounds && operatorBounds
              ? selectedBounds.top + selectedBounds.height / 2 <
                  operatorBounds.top + operatorBounds.height / 2
                ? "upper"
                : "lower"
              : null;
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
            operatorRegion,
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
      if (sumState.operatorRegion !== "lower") {
        throw new Error(`Summation did not start at its lower limit: ${JSON.stringify(sumState)}`);
      }
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
      if (integralState.operatorRegion !== "lower") {
        throw new Error(`Integral did not start at its lower limit: ${JSON.stringify(integralState)}`);
      }
      const integralTypedState = await typeBackslashOverPlaceholder(
        "integral",
        3,
      );

      await key("Enter", "Enter", 13);
      await waitForEvaluation(`(() => {
        const field = document.querySelector(
          ".formula-line.is-active math-field",
        );
        return { ready: Boolean(field?.isConnected && field.value === "") };
      })()`, "empty line for product insertion");
      await evaluate(
        `document.querySelector('[data-command-id="prod"]').click()`,
      );
      const productState = await readPlaceholderCaret("product", 3);
      if (productState.operatorRegion !== "lower") {
        throw new Error(`Product did not start at its lower limit: ${JSON.stringify(productState)}`);
      }
      const productTypedState = await typeBackslashOverPlaceholder("product", 2);

      const fillBoundedOperator = async (
        commandId,
        placeholderCount,
        values,
        expectedPattern,
      ) => {
        await clearField();
        await evaluate(
          `document.querySelector('[data-command-id=${JSON.stringify(commandId)}]').click()`,
        );
        const first = await readPlaceholderCaret(
          `${commandId} ordered lower limit`,
          placeholderCount,
        );
        if (first.operatorRegion !== "lower") {
          throw new Error(`${commandId} did not begin at lower limit: ${JSON.stringify(first)}`);
        }
        for (let index = 0; index < values.length; index += 1) {
          await typeText(values[index]);
          if (index < values.length - 1) {
            await key("Tab", "Tab", 9);
          }
        }
        const finalValue = await evaluate(`(() => {
          const field = document.querySelector(
            ".formula-line.is-active math-field",
          );
          return field?.value ?? "";
        })()`);
        const compact = finalValue.replace(/\s+/g, "");
        if (!expectedPattern.test(compact)) {
          throw new Error(
            `${commandId} placeholder order is wrong: ${JSON.stringify({ finalValue, compact })}`,
          );
        }
        return { first, finalValue };
      };

      const orderedSum = await fillBoundedOperator(
        "sum",
        3,
        ["i", "n", "a"],
        /\\sum_\{i\}\^\{n\}a/,
      );
      const orderedProduct = await fillBoundedOperator(
        "prod",
        3,
        ["j", "m", "b"],
        /\\prod_\{j\}\^\{m\}b/,
      );
      const orderedIntegral = await fillBoundedOperator(
        "int",
        4,
        ["a", "b", "f", "x"],
        /\\int_\{a\}\^\{b\}f.*\\mathrm\{d\}x/,
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
            productState,
            productTypedState,
            orderedSum,
            orderedProduct,
            orderedIntegral,
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
        window.__visualtexNativeSpaceTiming = {};
        window.addEventListener("keydown", () => {
          const handlerStartedAt = performance.now();
          queueMicrotask(() => {
            window.__visualtexNativeSpaceTiming.handlerMs =
              performance.now() - handlerStartedAt;
          });
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
          handlerMs: window.__visualtexNativeSpaceTiming?.handlerMs ?? null,
        };
      })()`, "Space commits the arrow-selected theta item");

      if (
        committedState.elapsedMs > 250 ||
        committedState.handlerMs === null ||
        committedState.handlerMs > 32
      ) {
        throw new Error(
          `Native Space selection was delayed: ${JSON.stringify(committedState)}`,
        );
      }

      if (initialState.firstCommand === "\\theta") {
        throw new Error(`Theta unexpectedly remained the first command: ${JSON.stringify(initialState)}`);
      }
      console.log(JSON.stringify({ initialState, movedState, committedState }, null, 2));
      console.log("Targeted native Space selection regression passed");
      return;
    }

    if (scenario === "raw-placeholder-visual") {
      const rawPlaceholderGeometry = await waitForEvaluation(`(() => {
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
        const sink = field.shadowRoot?.querySelector('[part="keyboard-sink"]');
        sink?.focus({ preventScroll: true });
        const placeholder = field.shadowRoot?.querySelector(
          ".visualtex-structural-placeholder",
        );
        const bounds = placeholder?.getBoundingClientRect();
        return {
          ready:
            field.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder").length === 2 &&
            Boolean(bounds && bounds.width > 0 && bounds.height > 0),
          x: bounds ? bounds.left + bounds.width / 2 : 0,
          y: bounds ? bounds.top + bounds.height / 2 : 0,
        };
      })()`, "fraction numerator placeholder before raw input");
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: rawPlaceholderGeometry.x,
        y: rawPlaceholderGeometry.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: rawPlaceholderGeometry.x,
        y: rawPlaceholderGeometry.y,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });
      await sleep(80);
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
          value: field?.value ?? "",
          mode: field?.mode ?? "",
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

    if (scenario === "candidate-query-reset") {
      await focusField();
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

      console.log(JSON.stringify({ confirmedState, resetState }, null, 2));
      console.log("Targeted command candidate query reset regression passed");
      return;
    }

    if (scenario === "direct-shortcut-placeholder") {
      const cases = [
        { input: "frac", command: "\\\\frac", expected: 2 },
        { input: "sqrt", command: "\\\\sqrt", expected: 1 },
        { input: "sum", command: "\\\\sum", expected: 2 },
        { input: "int", command: "\\\\int", expected: 2 },
      ];
      const results = [];
      for (const testCase of cases) {
        await clearField();
        await focusField();
        await typeText(testCase.input);
        const state = await waitForEvaluation(`(() => {
          const field = document.querySelector(".formula-line.is-active math-field");
          const root = field?.shadowRoot;
          const symbol = field?.placeholderSymbol || "▢";
          const placeholders = [...(root?.querySelectorAll(
            ".visualtex-structural-placeholder",
          ) ?? [])];
          const rawLeaves = [...(root?.querySelectorAll("[data-atom-id]") ?? [])]
            .filter(
              (node) =>
                (node.textContent || "").trim() === symbol &&
                !node.querySelector("[data-atom-id]") &&
                !node.classList.contains("visualtex-structural-placeholder"),
            );
          const styles = placeholders.map((node) => {
            const style = getComputedStyle(node);
            return {
              className: node.className,
              backgroundColor: style.backgroundColor,
              color: style.color,
              borderTopWidth: style.borderTopWidth,
            };
          });
          const backgrounds = new Set([
            "rgb(217, 237, 249)",
            "rgb(207, 232, 247)",
          ]);
          return {
            ready:
              Boolean(field?.value.includes(${JSON.stringify("__COMMAND__")})) &&
              !/\\\\(?:mathnormal|mathrm|mathbf|mathit|mathbfit)\\{\\\\[A-Za-z]+/.test(
                field?.value ?? "",
              ) &&
              placeholders.length === ${"__EXPECTED__"} &&
              rawLeaves.length === 0 &&
              styles.every(
                (item) =>
                  backgrounds.has(item.backgroundColor) &&
                  item.color === "rgba(0, 0, 0, 0)" &&
                  item.borderTopWidth === "0px",
              ),
            value: field?.value ?? "",
            placeholderCount: placeholders.length,
            rawPlaceholderCount: rawLeaves.length,
            styles,
          };
        })()`
          .replace("__COMMAND__", testCase.command)
          .replace("__EXPECTED__", String(testCase.expected)),
          `typing ${testCase.input} creates VisualTeX placeholder blocks`,
        );
        results.push({ ...testCase, state });
      }
      console.log(JSON.stringify(results, null, 2));
      console.log("Targeted direct shortcut placeholder regression passed");
      return;
    }

    if (scenario === "horizontal-overflow") {
      await clearField();
      await evaluate(`(() => {
        const field = document.querySelector(".formula-line.is-active math-field");
        const longTerm = Array.from({ length: 70 }, (_, index) =>
          "x_{" + index + "}^{2}+y_{" + index + "}^{2}",
        ).join("+");
        field.setValue(longTerm, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.position = field.lastOffset;
        return field.value;
      })()`);
      const overflowState = await waitForEvaluation(`(() => {
        const field = document.querySelector(".formula-line.is-active math-field");
        const host = field?.closest(".mathfield-host");
        const latex = field?.shadowRoot?.querySelector(".ML__latex");
        if (host && host.scrollWidth > host.clientWidth) {
          host.scrollLeft = host.scrollWidth;
        }
        return {
          ready:
            Boolean(field?.value.includes("x_{69}")) &&
            (latex?.getBoundingClientRect().width ?? 0) > 1000 &&
            getComputedStyle(host).overflowX === "auto" &&
            host.scrollWidth > host.clientWidth + 2 &&
            host.scrollLeft > 0,
          valueLength: field?.value.length ?? 0,
          formulaWidth: latex?.getBoundingClientRect().width ?? -1,
          clientWidth: host?.clientWidth ?? -1,
          scrollWidth: host?.scrollWidth ?? -1,
          scrollLeft: host?.scrollLeft ?? -1,
          overflowX: host ? getComputedStyle(host).overflowX : "",
        };
      })()`, "long formula horizontal scrollbar");
      console.log(JSON.stringify(overflowState, null, 2));
      console.log("Targeted horizontal formula overflow regression passed");
      return;
    }

    if (scenario === "toolbar-placeholder-overflow") {
      await clearField();
      await focusField();
      await typeText("frac");
      const directShortcutState = await waitForEvaluation(`(() => {
        const field = document.querySelector(".formula-line.is-active math-field");
        const root = field?.shadowRoot;
        const symbol = field?.placeholderSymbol || "▢";
        const placeholders = [...(root?.querySelectorAll(
          ".visualtex-structural-placeholder",
        ) ?? [])];
        const rawLeaves = [...(root?.querySelectorAll("[data-atom-id]") ?? [])]
          .filter(
            (node) =>
              (node.textContent || "").trim() === symbol &&
              !node.querySelector("[data-atom-id]") &&
              !node.classList.contains("visualtex-structural-placeholder"),
          );
        const styles = placeholders.map((node) => {
          const style = getComputedStyle(node);
          return {
            className: node.className,
            backgroundColor: style.backgroundColor,
            color: style.color,
            borderTopWidth: style.borderTopWidth,
          };
        });
        const backgrounds = new Set([
          "rgb(217, 237, 249)",
          "rgb(207, 232, 247)",
        ]);
        return {
          ready:
            Boolean(field?.value.includes("\\\\frac")) &&
            !/\\\\(?:mathnormal|mathrm|mathbf|mathit|mathbfit)\\{\\\\[A-Za-z]+/.test(
              field?.value ?? "",
            ) &&
            placeholders.length === 2 &&
            rawLeaves.length === 0 &&
            styles.every(
              (item) =>
                backgrounds.has(item.backgroundColor) &&
                item.color === "rgba(0, 0, 0, 0)" &&
                item.borderTopWidth === "0px",
            ),
          value: field?.value ?? "",
          placeholderCount: placeholders.length,
          rawPlaceholderCount: rawLeaves.length,
          styles,
        };
      })()`, "typing frac creates VisualTeX placeholder blocks");

      await clearField();
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector('[data-command-id="frac"]')),
      }))()`, "fraction toolbar command");
      await clickSelectorWithPointer('[data-command-id=frac]');
      const placeholderState = await waitForEvaluation(`(() => {
        const field = document.querySelector(".formula-line.is-active math-field");
        const root = field?.shadowRoot;
        const placeholderSymbol = field?.placeholderSymbol || "▢";
        const unstyledLeafPlaceholders = [...(root?.querySelectorAll("[data-atom-id]") ?? [])]
          .filter(
            (node) =>
              (node.textContent || "").trim() === placeholderSymbol &&
              !node.querySelector("[data-atom-id]") &&
              !node.classList.contains("visualtex-structural-placeholder"),
          );
        const placeholders = [...(root?.querySelectorAll(
          ".visualtex-structural-placeholder",
        ) ?? [])];
        const styled = placeholders.map((node) => {
          const style = getComputedStyle(node);
          return {
            className: node.className,
            borderTopWidth: style.borderTopWidth,
            backgroundColor: style.backgroundColor,
          };
        });
        const validBackgrounds = new Set([
          "rgb(217, 237, 249)",
          "rgb(207, 232, 247)",
        ]);
        return {
          ready:
            Boolean(field?.value.includes("\\\\frac")) &&
            Boolean(root?.getElementById("visualtex-structural-placeholder-style")) &&
            placeholders.length === 2 &&
            unstyledLeafPlaceholders.length === 0 &&
            styled.every(
              (item) =>
                item.borderTopWidth === "0px" &&
                validBackgrounds.has(item.backgroundColor),
            ),
          value: field?.value ?? "",
          styleInstalled: Boolean(root?.getElementById("visualtex-structural-placeholder-style")),
          placeholderCount: placeholders.length,
          unstyledLeafPlaceholderCount: unstyledLeafPlaceholders.length,
          styled,
        };
      })()`, "toolbar fraction placeholders use VisualTeX blocks");

      const placeholderVariants = [
        { name: "default-math", state: placeholderState },
      ];

      await evaluate(`(() => {
        const button = document.querySelector('[data-command-id="frac"]');
        if (!(button instanceof HTMLElement)) return false;
        const bounds = button.getBoundingClientRect();
        button.dispatchEvent(new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          button: 2,
          buttons: 2,
          clientX: bounds.left + bounds.width / 2,
          clientY: bounds.top + bounds.height / 2,
        }));
        return true;
      })()`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(
          document.querySelector(".formula-hotkey-context-action"),
        ),
      }))()`, "fraction hotkey context action");
      await evaluate(`document.querySelector(
        ".formula-hotkey-context-action",
      ).click()`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector(".formula-hotkey-recorder-dialog")),
      }))()`, "fraction hotkey recorder");
      await evaluate(`document.dispatchEvent(new KeyboardEvent("keydown", {
        bubbles: true,
        cancelable: true,
        code: "KeyF",
        key: "f",
        ctrlKey: true,
        altKey: true,
      }))`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(
          document.querySelector(
            ".formula-hotkey-recorder-footer .primary-button:not(:disabled)",
          ),
        ),
      }))()`, "saveable fraction hotkey");
      await evaluate(`document.querySelector(
        ".formula-hotkey-recorder-footer .primary-button",
      ).click()`);
      await clearField();
      await evaluate(`(() => {
        const field = document.querySelector(".formula-line.is-active math-field");
        field.focus();
        field.shadowRoot
          ?.querySelector('[part="keyboard-sink"]')
          ?.focus({ preventScroll: true });
        return field.dispatchEvent(new KeyboardEvent("keydown", {
          bubbles: true,
          composed: true,
          cancelable: true,
          code: "KeyF",
          key: "f",
          ctrlKey: true,
          altKey: true,
        }));
      })()`);
      const shortcutPlaceholderState = await waitForEvaluation(`(() => {
        const field = document.querySelector(".formula-line.is-active math-field");
        const root = field?.shadowRoot;
        const symbol = field?.placeholderSymbol || "▢";
        const placeholders = [...(root?.querySelectorAll(
          ".visualtex-structural-placeholder",
        ) ?? [])];
        const rawLeaves = [...(root?.querySelectorAll("[data-atom-id]") ?? [])]
          .filter(
            (node) =>
              (node.textContent || "").trim() === symbol &&
              !node.querySelector("[data-atom-id]") &&
              !node.classList.contains("visualtex-structural-placeholder"),
          );
        return {
          ready:
            Boolean(field?.value.includes("\\\\frac")) &&
            placeholders.length === 2 &&
            rawLeaves.length === 0,
          value: field?.value ?? "",
          placeholderCount: placeholders.length,
          rawPlaceholderCount: rawLeaves.length,
          classes: placeholders.map((node) => node.className),
        };
      })()`, "Ctrl+Alt+F fraction placeholders");

      const shortOverflowState = await evaluate(`(() => {
        const field = document.querySelector(".formula-line.is-active math-field");
        const host = field?.closest(".mathfield-host");
        return {
          overflowX: host ? getComputedStyle(host).overflowX : "",
          clientWidth: host?.clientWidth ?? -1,
          scrollWidth: host?.scrollWidth ?? -1,
          unnecessaryOverflow: Boolean(
            host && host.scrollWidth > host.clientWidth + 2,
          ),
        };
      })()`);
      if (
        shortOverflowState.overflowX !== "auto" ||
        shortOverflowState.unnecessaryOverflow
      ) {
        throw new Error(
          `Short formula scrollbar state is invalid: ${JSON.stringify(shortOverflowState)}`,
        );
      }
      console.log(
        JSON.stringify(
          {
            directShortcutState,
            placeholderVariants,
            shortcutPlaceholderState,
            shortOverflowState,
          },
          null,
          2,
        ),
      );
      console.log("Targeted toolbar and shortcut placeholder regression passed");
      return;

      await evaluate(`(() => {
        const field = document.querySelector(".formula-line.is-active math-field");
        const longTerm = Array.from({ length: 70 }, (_, index) =>
          "x_{" + index + "}^{2}+y_{" + index + "}^{2}",
        ).join("+");
        field.setValue(longTerm, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: false,
        });
        field.position = field.lastOffset;
        field.dispatchEvent(new InputEvent("input", {
          bubbles: true,
          composed: true,
          inputType: "insertText",
        }));
        return field.value;
      })()`);
      const overflowState = await waitForEvaluation(`(() => {
        const field = document.querySelector(".formula-line.is-active math-field");
        const host = field?.closest(".mathfield-host");
        const line = field?.closest(".formula-line");
        const stack = field?.closest(".mathfield-stack");
        const latex = field?.shadowRoot?.querySelector(".ML__latex");
        const base = field?.shadowRoot?.querySelector(".ML__base");
        const candidate = [field, host, line, stack].find(
          (node) => node && node.scrollWidth > node.clientWidth + 2,
        );
        if (host && host.scrollWidth > host.clientWidth) {
          host.scrollLeft = host.scrollWidth;
        }
        return {
          ready:
            Boolean(field && host && line && stack) &&
            Boolean(field?.value.includes("x_{69}")) &&
            (latex?.getBoundingClientRect().width ?? 0) > 1000 &&
            getComputedStyle(host).overflowX === "auto" &&
            host.scrollWidth > host.clientWidth + 2 &&
            host.scrollLeft > 0,
          valueLength: field?.value.length ?? 0,
          latexWidth: latex?.getBoundingClientRect().width ?? -1,
          baseWidth: base?.getBoundingClientRect().width ?? -1,
          field: field ? {
            className: field.className,
            clientWidth: field.clientWidth,
            scrollWidth: field.scrollWidth,
            offsetWidth: field.offsetWidth,
            computedWidth: getComputedStyle(field).width,
            display: getComputedStyle(field).display,
            overflowX: getComputedStyle(field).overflowX,
            contentClientWidth: field.shadowRoot?.querySelector('[part="content"]')?.clientWidth ?? -1,
            contentScrollWidth: field.shadowRoot?.querySelector('[part="content"]')?.scrollWidth ?? -1,
          } : null,
          host: host ? {
            className: host.className,
            clientWidth: host.clientWidth,
            scrollWidth: host.scrollWidth,
            computedWidth: getComputedStyle(host).width,
            display: getComputedStyle(host).display,
            justifyContent: getComputedStyle(host).justifyContent,
            overflowX: getComputedStyle(host).overflowX,
            scrollLeft: host.scrollLeft,
            scrollable: host.scrollWidth > host.clientWidth + 2,
          } : null,
          line: line ? {
            clientWidth: line.clientWidth,
            scrollWidth: line.scrollWidth,
            overflowX: getComputedStyle(line).overflowX,
          } : null,
          stack: stack ? {
            clientWidth: stack.clientWidth,
            scrollWidth: stack.scrollWidth,
            overflowX: getComputedStyle(stack).overflowX,
          } : null,
          overflowCandidate: candidate?.className ?? candidate?.tagName ?? "",
        };
      })()`, "long formula overflow diagnostic");

      console.log(
        JSON.stringify(
          {
            directShortcutState,
            placeholderVariants,
            shortcutPlaceholderState,
            shortOverflowState,
            overflowState,
          },
          null,
          2,
        ),
      );
      console.log("Targeted toolbar placeholder and horizontal overflow regression passed");
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
      const pressControlShortcut = async (keyValue, code, virtualKeyCode, shift = false) => {
        const modifiers = 2 | (shift ? 8 : 0);
        await client.send("Input.dispatchKeyEvent", {
          type: "keyDown",
          key: "Control",
          code: "ControlLeft",
          windowsVirtualKeyCode: 17,
          nativeVirtualKeyCode: 17,
          modifiers: 2,
        });
        if (shift) {
          await client.send("Input.dispatchKeyEvent", {
            type: "keyDown",
            key: "Shift",
            code: "ShiftLeft",
            windowsVirtualKeyCode: 16,
            nativeVirtualKeyCode: 16,
            modifiers,
          });
        }
        await client.send("Input.dispatchKeyEvent", {
          type: "keyDown",
          key: keyValue,
          code,
          windowsVirtualKeyCode: virtualKeyCode,
          nativeVirtualKeyCode: virtualKeyCode,
          modifiers,
        });
        await client.send("Input.dispatchKeyEvent", {
          type: "keyUp",
          key: keyValue,
          code,
          windowsVirtualKeyCode: virtualKeyCode,
          nativeVirtualKeyCode: virtualKeyCode,
          modifiers,
        });
        if (shift) {
          await client.send("Input.dispatchKeyEvent", {
            type: "keyUp",
            key: "Shift",
            code: "ShiftLeft",
            windowsVirtualKeyCode: 16,
            nativeVirtualKeyCode: 16,
            modifiers: 2,
          });
        }
        await client.send("Input.dispatchKeyEvent", {
          type: "keyUp",
          key: "Control",
          code: "ControlLeft",
          windowsVirtualKeyCode: 17,
          nativeVirtualKeyCode: 17,
          modifiers: 0,
        });
        await sleep(180);
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
      await pressControlShortcut("z", "KeyZ", 90);
      const undoState = await readFieldState();
      if (
        !undoState.value.includes("\\placeholder{}") ||
        !undoState.value.includes("^")
      ) {
        throw new Error(
          `Undo did not restore the structured placeholder transaction: ${JSON.stringify({ undoBefore, undoAfterComposition, undoState })}`,
        );
      }
      await pressControlShortcut("y", "KeyY", 89);
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
        const placeholderGeometry = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const placeholder = field?.shadowRoot?.querySelector(
            ".visualtex-structural-placeholder",
          );
          const bounds = placeholder?.getBoundingClientRect();
          return {
            ready: Boolean(bounds && bounds.width > 0 && bounds.height > 0),
            x: bounds ? bounds.left + bounds.width / 2 : 0,
            y: bounds ? bounds.top + bounds.height / 2 : 0,
          };
        })()`, `structural placeholder geometry before ${testCase.name}`);
        await client.send("Input.dispatchMouseEvent", {
          type: "mousePressed",
          x: placeholderGeometry.x,
          y: placeholderGeometry.y,
          button: "left",
          buttons: 1,
          clickCount: 1,
        });
        await client.send("Input.dispatchMouseEvent", {
          type: "mouseReleased",
          x: placeholderGeometry.x,
          y: placeholderGeometry.y,
          button: "left",
          buttons: 0,
          clickCount: 1,
        });
        await sleep(80);
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
        { name: "acute", source: String.raw`a+\acute{\placeholder{}}+b` },
        { name: "grave", source: String.raw`a+\grave{\placeholder{}}+b` },
        { name: "dot", source: String.raw`a+\dot{\placeholder{}}+b` },
        { name: "ddot", source: String.raw`a+\ddot{\placeholder{}}+b` },
        { name: "dddot", source: String.raw`a+\dddot{\placeholder{}}+b` },
        { name: "ddddot", source: String.raw`a+\ddddot{\placeholder{}}+b` },
        { name: "tilde", source: String.raw`a+\tilde{\placeholder{}}+b` },
        { name: "bar", source: String.raw`a+\bar{\placeholder{}}+b` },
        { name: "breve", source: String.raw`a+\breve{\placeholder{}}+b` },
        { name: "check", source: String.raw`a+\check{\placeholder{}}+b` },
        { name: "hat", source: String.raw`a+\hat{\placeholder{}}+b` },
        { name: "vec", source: String.raw`a+\vec{\placeholder{}}+b` },
        { name: "widehat", source: String.raw`a+\widehat{\placeholder{}}+b` },
        { name: "widetilde", source: String.raw`a+\widetilde{\placeholder{}}+b` },
        { name: "overline", source: String.raw`a+\overline{\placeholder{}}+b` },
        { name: "mathring", source: String.raw`a+\mathring{\placeholder{}}+b` },
        { name: "overrightarrow", source: String.raw`a+\overrightarrow{\placeholder{}}+b` },
        { name: "overleftarrow", source: String.raw`a+\overleftarrow{\placeholder{}}+b` },
        { name: "overleftrightarrow", source: String.raw`a+\overleftrightarrow{\placeholder{}}+b` },
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
            ".visualtex-accent-structural-placeholder, .ML__vlist .ML__cmr, .ML__placeholder",
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
              ":scope > .ML__center .ML__accent-body, :scope > .ML__center .ML__stretchy",
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

        const blueColors = new Set([
          "rgb(217, 237, 249)",
          "rgb(207, 232, 247)",
        ]);
        const stablePlaceholder = (state) =>
          state.placeholder &&
          (blueColors.has(state.placeholder.visualBackground) ||
            blueColors.has(state.placeholder.background)) &&
          state.placeholder.color === "rgba(0, 0, 0, 0)" &&
          state.placeholder.borderTopWidth === "0px" &&
          state.placeholder.borderRightWidth === "0px" &&
          state.placeholder.borderBottomWidth === "0px" &&
          state.placeholder.borderLeftWidth === "0px";
        const stableAlignment = (state) =>
          !state.placeholder?.classes.includes(
            "visualtex-accent-structural-placeholder",
          ) || state.placeholder.alignmentDelta <= 1;
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
          !stablePlaceholder(initial) ||
          !stablePlaceholder(afterRight) ||
          !stablePlaceholder(reenteredFromRight) ||
          !stablePlaceholder(afterExitLeft) ||
          !stablePlaceholder(reenteredFromLeft) ||
          !stablePlaceholder(held) ||
          !stablePlaceholder(released) ||
          !stableAlignment(initial) ||
          !stableAlignment(held) ||
          !stableOverlay(initial) ||
          !stableOverlay(held) ||
          !stableOverlay(released) ||
          !held.pointerSelecting ||
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

      const readInsertedAccentState = (expectedCommand) =>
        waitForEvaluation(`(() => {
          const field =
            document.querySelector(".formula-line.is-active math-field") ??
            document.querySelector("math-field");
          const root = field?.shadowRoot;
          const symbol = field?.placeholderSymbol || "▢";
          const placeholders = [...(root?.querySelectorAll(
            ".visualtex-accent-structural-placeholder, .visualtex-structural-placeholder",
          ) ?? [])].filter((node) =>
            (node.textContent || "").replace(/\\s+/g, "").startsWith(symbol)
          );
          const rawBlackBoxes = [...(root?.querySelectorAll(
            ".ML__vlist .ML__cmr, .ML__placeholder",
          ) ?? [])].filter((node) =>
            (node.textContent || "").trim() === symbol &&
            !node.classList.contains("visualtex-structural-placeholder")
          );
          const blueColors = new Set([
            "rgb(217, 237, 249)",
            "rgb(207, 232, 247)",
          ]);
          const visuallyStyled = placeholders.every((node) => {
            const style = getComputedStyle(node);
            const pseudo = getComputedStyle(node, "::before");
            return (
              (blueColors.has(style.backgroundColor) ||
                blueColors.has(pseudo.backgroundColor)) &&
              style.color === "rgba(0, 0, 0, 0)" &&
              style.borderTopWidth === "0px"
            );
          });
          return {
            ready:
              Boolean(field?.value.includes(${JSON.stringify(expectedCommand)})) &&
              placeholders.length === 1 &&
              rawBlackBoxes.length === 0 &&
              visuallyStyled,
            value: field?.value ?? "",
            placeholderCount: placeholders.length,
            rawBlackBoxCount: rawBlackBoxes.length,
            classes: placeholders.map((node) => node.className),
          };
        })()`,
          `inserted accent ${expectedCommand}`,
        );

      const shortcutCases = [
        ["acute", "\\acute{"],
        ["grave", "\\grave{"],
        ["hat", "\\hat{"],
        ["widehat", "\\widehat{"],
        ["bar", "\\bar{"],
        ["overline", "\\overline{"],
        ["vec", "\\vec{"],
        ["tilde", "\\tilde{"],
        ["widetilde", "\\widetilde{"],
        ["dot", "\\dot{"],
        ["ddot", "\\ddot{"],
        ["dddot", "\\dddot{"],
        ["breve", "\\breve{"],
        ["check", "\\check{"],
        ["mathring", "\\mathring{"],
      ];
      const shortcutStates = [];
      for (const [input, command] of shortcutCases) {
        await clearField();
        await focusField();
        await typeText(input);
        shortcutStates.push({
          input,
          state: await readInsertedAccentState(command),
        });
      }

      await evaluate(`document.querySelector('[data-category="structure"]')?.click()`);
      const toolbarCases = [
        ["overline", "\\overline{"],
        ["hat", "\\hat{"],
        ["widehat", "\\widehat{"],
        ["tilde", "\\tilde{"],
        ["widetilde", "\\widetilde{"],
        ["dotaccent", "\\dot{"],
        ["ddotaccent", "\\ddot{"],
        ["checkaccent", "\\check{"],
        ["breveaccent", "\\breve{"],
        ["graveaccent", "\\grave{"],
      ];
      const toolbarStates = [];
      for (const [commandId, command] of toolbarCases) {
        await waitForEvaluation(`(() => ({
          ready: Boolean(document.querySelector('[data-command-id="${commandId}"]')),
        }))()`, `toolbar accent ${commandId}`);
        await clearField();
        await clickSelectorWithPointer(`[data-command-id=${commandId}]`);
        toolbarStates.push({
          commandId,
          state: await readInsertedAccentState(command),
        });
      }
      states.push({
        name: "shortcut-entry-matrix",
        count: shortcutStates.length,
        values: shortcutStates.map((item) => item.state.value),
      });
      states.push({
        name: "toolbar-entry-matrix",
        count: toolbarStates.length,
        values: toolbarStates.map((item) => item.state.value),
      });

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

    if (scenario === "usage-ranking") {
      await evaluate(`(() => {
        if (!document.querySelector(".formula-toolbar")) {
          document.querySelector(".sidebar-toggle")?.click();
        }
        return true;
      })()`);
      const commonCommandState = await waitForEvaluation(`(() => {
        const buttons = [...document.querySelectorAll(
          '.template-strip > [data-command-id]',
        )];
        const persisted = JSON.parse(
          localStorage.getItem("visualtex-editor") || "{}",
        );
        return {
          ready:
            buttons.length === 45 &&
            buttons[0]?.dataset.commandId === "frac" &&
            persisted.state?.usage?.sqrt?.contextCounts?.toolbar === 18,
          firstCommandId: buttons[0]?.dataset.commandId ?? "",
          usage: persisted.state?.usage?.sqrt ?? null,
        };
      })()`, "toolbar common commands keep their fixed order despite usage");

      await evaluate(`document.querySelector('[data-toolbar-view="tiles"]').click()`);
      const commonTileState = await waitForEvaluation(`(() => {
        const tiles = [...document.querySelectorAll(
          '.formula-tile-list > .formula-tile-button',
        )];
        return {
          ready:
            tiles.length === 10 &&
            tiles[0]?.dataset.formulaTileId === "quadratic-formula",
          firstTileId: tiles[0]?.dataset.formulaTileId ?? "",
        };
      })()`, "common formula tiles keep their fixed order despite usage");
      await evaluate(`document.querySelector(
        '[data-formula-tile-id="gaussian-integral"]',
      ).click()`);
      const tileUsageState = await waitForEvaluation(`(() => {
        const persisted = JSON.parse(
          localStorage.getItem("visualtex-editor") || "{}",
        );
        const usage = persisted.state?.usage?.["formula-tile-gaussian-integral"];
        return {
          ready:
            usage?.useCount === 13 &&
            usage?.contextCounts?.toolbar === 13,
          usage,
        };
      })()`, "formula tile click frequency persistence");

      await sleep(180);
      await clearField();
      await sleep(120);
      await clearField();
      await typeText("\\be");
      const nativeRankState = await waitForEvaluation(`(() => {
        const source = document.getElementById("mathlive-suggestion-popover");
        const stable = document.getElementById(
          "visualtex-native-input-suggestion-popover",
        );
        const sourceCommands = [...(source?.querySelectorAll(
          "li[data-command]",
        ) ?? [])].map((item) => item.dataset.command ?? "");
        const stableCommands = [...(stable?.querySelectorAll(
          "li[data-command]",
        ) ?? [])].map((item) => item.dataset.command ?? "");
        return {
          ready:
            source?.classList.contains("is-visible") &&
            sourceCommands[0] === "\\\\beth" &&
            stableCommands[0] === "\\\\beth" &&
            source?.querySelector("li.ML__popover__current")?.dataset.command ===
              "\\\\beth",
          sourceCommands,
          stableCommands,
          selected:
            source?.querySelector("li.ML__popover__current")?.dataset.command ??
            "",
        };
      })()`, "MathLive native candidates sorted by lifetime frequency");
      await key("Enter", "Enter", 13);
      const nativeUsageState = await waitForEvaluation(`(() => {
        const persisted = JSON.parse(
          localStorage.getItem("visualtex-editor") || "{}",
        );
        const usage = persisted.state?.usage?.["mathlive-native:\\\\beth"];
        const field = document.querySelector("math-field");
        return {
          ready:
            usage?.useCount === 26 &&
            usage?.contextCounts?.candidate === 26 &&
            field?.value === "\\\\beth",
          usage,
          value: field?.value ?? "",
        };
      })()`, "native candidate usage persisted after commit");

      await client.send("Page.reload", { ignoreCache: true });
      await waitForEvaluation(
        `(() => ({ ready: Boolean(document.querySelector("math-field")) }))()`,
        "reloaded field for persistent ranking",
      );
      await clearField();
      await typeText("\\be");
      const reloadedNativeRankState = await waitForEvaluation(`(() => {
        const source = document.getElementById("mathlive-suggestion-popover");
        const first = source?.querySelector("li[data-command]");
        const persisted = JSON.parse(
          localStorage.getItem("visualtex-editor") || "{}",
        );
        return {
          ready:
            source?.classList.contains("is-visible") &&
            first?.dataset.command === "\\\\beth" &&
            persisted.state?.usage?.["mathlive-native:\\\\beth"]?.useCount ===
              26,
          firstCommand: first?.dataset.command ?? "",
          useCount:
            persisted.state?.usage?.["mathlive-native:\\\\beth"]?.useCount ??
            0,
        };
      })()`, "native frequency ranking survives reload");

      console.log(
        JSON.stringify(
          {
            commonCommandState,
            commonTileState,
            tileUsageState,
            nativeRankState,
            nativeUsageState,
            reloadedNativeRankState,
          },
          null,
          2,
        ),
      );
      console.log("Targeted persistent usage ranking regression passed");
      return;
    }

    if (scenario === "native-input-popover") {
      await clearField();
      await key("Enter", "Enter", 13);
      await waitForEvaluation(`(() => {
        const fields = [...document.querySelectorAll("math-field")];
        return {
          ready: fields.length === 2 && fields[1]?.hasFocus?.(),
          lineCount: fields.length,
          focusedIndex: fields.findIndex((field) => field.hasFocus?.()),
        };
      })()`, "second formula line before native suggestion navigation");
      await key("ArrowUp", "ArrowUp", 38);
      await waitForEvaluation(`(() => {
        const fields = [...document.querySelectorAll("math-field")];
        return {
          ready: fields.length === 2 && fields[0]?.hasFocus?.(),
          focusedIndex: fields.findIndex((field) => field.hasFocus?.()),
        };
      })()`, "first formula line before native suggestion navigation");
      // Row navigation reapplies focus and selection at 0 ms and 80 ms. Wait
      // until those callbacks settle, otherwise the test can type `\\` first,
      // have the caret reset in front of it, and accidentally produce `be\\`.
      await sleep(140);
      await typeText("\\be");
      const initial = await waitForEvaluation(`(() => {
        const stable = document.getElementById("visualtex-native-input-suggestion-popover");
        const source = document.getElementById("mathlive-suggestion-popover");
        const field = document.querySelectorAll("math-field")[0];
        const bounds = stable?.getBoundingClientRect();
        const commands = [...(stable?.querySelectorAll("li[data-command]") ?? [])]
          .map((item) => item.dataset.command ?? "");
        const sourceCommands = [...(source?.querySelectorAll("li[data-command]") ?? [])]
          .map((item) => item.dataset.command ?? "");
        const rawLatex = [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
          .filter((node) => !node.classList.contains("ML__suggestion"))
          .map((node) => node.textContent ?? "")
          .join("");
        return {
          ready:
            Boolean(stable?.classList.contains("is-visible")) &&
            commands.length >= 2 &&
            source?.dataset.visualtexInputPopoverSource === "true" &&
            document.querySelectorAll("math-field").length === 2 &&
            document.querySelectorAll("math-field")[0]?.hasFocus?.() &&
            !document.querySelector(".suggestion-popup"),
          commands,
          sourceCommands,
          rawLatex,
          fieldMode: field?.mode ?? "",
          fieldValue: field?.value ?? "",
          sourceExists: Boolean(source),
          sourceVisible: source?.classList.contains("is-visible") ?? false,
          focusedIndex: [...document.querySelectorAll("math-field")]
            .findIndex((field) => field.hasFocus?.()),
          selected: stable?.querySelector("li.ML__popover__current")?.dataset.command ?? "",
          bounds: bounds
            ? { left: bounds.left, top: bounds.top, width: bounds.width, height: bounds.height }
            : null,
          sourceOpacity: source ? getComputedStyle(source).opacity : "",
          customCandidateVisible: Boolean(document.querySelector(".suggestion-popup")),
        };
      })()`, "stable native input-selection popover for \\be");

      await evaluate(`(() => {
        const source = document.getElementById("mathlive-suggestion-popover");
        const stable = document.getElementById("visualtex-native-input-suggestion-popover");
        const field = document.querySelectorAll("math-field")[0];
        const beforeSelected = stable?.querySelector("li.ML__popover__current")?.dataset.command ?? "";
        const sourceItems = [...(source?.querySelectorAll("li[data-command]") ?? [])];
        const sourceIndex = sourceItems.findIndex((item) =>
          item.classList.contains("ML__popover__current"),
        );
        const downCommand = sourceItems[
          sourceIndex < 0 ? 0 : (sourceIndex + 1) % sourceItems.length
        ]?.dataset.command ?? "";
        const key = downCommand && downCommand !== beforeSelected ? "ArrowDown" : "ArrowUp";
        const direction = key === "ArrowDown" ? 1 : -1;
        const expectedCommand = sourceItems[
          sourceIndex < 0
            ? direction > 0 ? 0 : sourceItems.length - 1
            : (sourceIndex + direction + sourceItems.length) % sourceItems.length
        ]?.dataset.command ?? "";
        source?.classList.remove("is-visible");
        source?.setAttribute("aria-hidden", "true");
        stable?.classList.add("is-visible");
        stable?.setAttribute("aria-hidden", "false");
        field?.dispatchEvent(new KeyboardEvent("keydown", {
          key,
          code: key,
          bubbles: true,
          composed: true,
          cancelable: true,
        }));
        const selected = stable?.querySelector("li.ML__popover__current")?.dataset.command ?? "";
        const latestSource = document.getElementById("mathlive-suggestion-popover");
        latestSource?.classList.add("is-visible");
        latestSource?.setAttribute("aria-hidden", "false");
        window.__visualtexNativePriorityState = {
          beforeSelected,
          selected,
          expectedCommand,
          key,
        };
      })()`);
      const priorityState = await waitForEvaluation(`(() => {
        const state = window.__visualtexNativePriorityState;
        const lines = [...document.querySelectorAll(".formula-line")];
        const activeIndex = lines.findIndex((line) => line.classList.contains("is-active"));
        const source = document.getElementById("mathlive-suggestion-popover");
        const stable = document.getElementById("visualtex-native-input-suggestion-popover");
        return {
          ready: Boolean(
            state?.selected &&
            state.selected === state.expectedCommand &&
            activeIndex === 0
          ),
          beforeSelected: state?.beforeSelected ?? "",
          selected: state?.selected ?? "",
          expectedCommand: state?.expectedCommand ?? "",
          key: state?.key ?? "",
          activeIndex,
          sourceVisible: source?.classList.contains("is-visible") ?? false,
          stableVisible: stable?.classList.contains("is-visible") ?? false,
        };
      })()`, "native suggestion priority while source visibility is transient");

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
        const fields = [...document.querySelectorAll("math-field")];
        return {
          ready:
            stable === monitor?.node &&
            stable?.classList.contains("is-visible") &&
            selected &&
            selected !== ${JSON.stringify(priorityState.selected)} &&
            fields.length === 2 &&
            fields[0]?.hasFocus?.() &&
            monitor.removed === 0 &&
            monitor.hiddenTransitions === 0 &&
            monitor.ariaHiddenTransitions === 0 &&
            Math.abs((bounds?.left ?? 0) - ${initial.bounds.left}) <= 1 &&
            Math.abs((bounds?.top ?? 0) - ${initial.bounds.top}) <= 1 &&
            Math.abs((bounds?.width ?? 0) - ${initial.bounds.width}) <= 1 &&
            Math.abs((bounds?.height ?? 0) - ${initial.bounds.height}) <= 1,
          sameNode: stable === monitor?.node,
          selected,
          focusedIndex: fields.findIndex((field) => field.hasFocus?.()),
          bounds: bounds
            ? { left: bounds.left, top: bounds.top, width: bounds.width, height: bounds.height }
            : null,
          removed: monitor?.removed ?? -1,
          hiddenTransitions: monitor?.hiddenTransitions ?? -1,
          ariaHiddenTransitions: monitor?.ariaHiddenTransitions ?? -1,
          customCandidateVisible: Boolean(document.querySelector(".suggestion-popup")),
        };
      })()`, "arrow key moves only the native input-selection highlight");

      await typeText("t");
      const refinedState = await waitForEvaluation(`(() => {
        const monitor = window.__visualtexNativeInputMonitor;
        const stable = document.getElementById("visualtex-native-input-suggestion-popover");
        const commands = [...(stable?.querySelectorAll("li[data-command]") ?? [])]
          .map((item) => item.dataset.command ?? "");
        return {
          ready:
            stable === monitor?.node &&
            stable?.classList.contains("is-visible") &&
            commands.some((command) => command === "\\\\beta") &&
            commands.every((command) => command.startsWith("\\\\bet")) &&
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
      })()`, "\\be to \\bet updates inside one persistent input-selection frame");

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
      })()`, "Backspace restores \\be suggestions without remounting the input-selection frame");

      await evaluate(`window.__visualtexNativeInputObserver?.disconnect()`);
      console.log(JSON.stringify({ initial, priorityState, arrowState, refinedState, restoredState }, null, 2));
      console.log("Targeted native input-selection popover regression passed");
      return;
    }

    if (scenario === "export") {
      await evaluate(`(() => {
        const storageKey = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
        const line = { id: crypto.randomUUID(), latex: "\\\\frac{a}{b}+x^2" };
        persisted.state = {
          ...(persisted.state || {}),
          title: "Export Test",
          lines: [line],
          activeLineId: line.id,
        };
        localStorage.setItem(storageKey, JSON.stringify(persisted));
        location.reload();
      })()`);
      await waitForEvaluation(`(() => ({
        ready:
          document.querySelector(".document-title-area input")?.value === "Export Test" &&
          document.querySelector("math-field")?.value?.includes("\\\\frac"),
        title: document.querySelector(".document-title-area input")?.value ?? "",
        value: document.querySelector("math-field")?.value ?? "",
      }))()`, "formula document prepared for export");

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
        await evaluate(`document.querySelector(".export-menu-trigger")?.click()`);
        await waitForEvaluation(`(() => ({
          ready: Boolean(document.querySelector(".export-menu-popover")),
        }))()`, `export menu opened for ${label}`);
        await evaluate(`(() => {
          const button = [...document.querySelectorAll(".export-format-options > button")]
            .find((candidate) => candidate.querySelector("strong")?.textContent?.trim() === ${JSON.stringify(label)});
          button?.click();
        })()`);
        await waitForEvaluation(`(() => ({
          ready:
            (window.__visualtexCapturedExports?.length ?? 0) >= ${expectedCount} &&
            !document.querySelector(".export-menu-popover"),
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
            text.includes("\\\\frac{a}{b}+x^2"),
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
        return {
          ready:
            item.filename.endsWith(".svg") &&
            text.startsWith("<svg") &&
            !text.includes("<foreignObject"),
          filename: item.filename,
          bytes: new TextEncoder().encode(text).length,
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
        return {
          ready:
            item.filename.endsWith(".png") &&
            bytes.length > expected.length &&
            expected.every((value, index) => bytes[index] === value),
          filename: item.filename,
          bytes: bytes.length,
          signature: [...bytes.slice(0, 8)],
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

    if (
      scenario === "wrapper" ||
      scenario === "wrapper-auto" ||
      scenario === "wrapper-continuous"
    ) {
      await focusField();
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
      await waitForEvaluation(`(() => ({ ready: Boolean(document.querySelector("math-field")) }))()`, "fresh field for mathcal test");
      await focusField();
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

    if (scenario === "context-style") {
      await focusField();
      await clearField();
      await typeText("abc");

      const contextPoint = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.selection = {
          ranges: [[0, field.lastOffset]],
          direction: "forward",
        };
        const bounds = field.shadowRoot
          ?.querySelector('[part="content"]')
          ?.getBoundingClientRect();
        return bounds
          ? { x: bounds.left + bounds.width / 2, y: bounds.top + bounds.height / 2 }
          : null;
      })()`);
      if (!contextPoint) throw new Error("Unable to resolve formula bounds");

      const openContextMenu = async () => {
        await client.send("Input.dispatchMouseEvent", {
          type: "mousePressed",
          x: contextPoint.x,
          y: contextPoint.y,
          button: "right",
          buttons: 2,
          clickCount: 1,
        });
        await client.send("Input.dispatchMouseEvent", {
          type: "mouseReleased",
          x: contextPoint.x,
          y: contextPoint.y,
          button: "right",
          buttons: 0,
          clickCount: 1,
        });
        await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const menu = field?.shadowRoot?.querySelector("menu.ui-menu-container");
          return { ready: Boolean(menu && getComputedStyle(menu).display !== "none") };
        })()`, "MathLive context menu");
      };

      await openContextMenu();
      const colorMenuPoint = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const root = field?.shadowRoot?.querySelector("menu.ui-menu-container");
        const colorItem = [...(root?.querySelectorAll(":scope > li") ?? [])]
          .find((item) => /^(颜色|Color)$/.test(item.textContent?.trim() ?? ""));
        const bounds = colorItem?.getBoundingClientRect();
        return bounds
          ? { x: bounds.left + bounds.width / 2, y: bounds.top + bounds.height / 2 }
          : null;
      })()`);
      if (!colorMenuPoint) throw new Error("Unable to resolve foreground color menu item");
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseMoved",
        x: colorMenuPoint.x,
        y: colorMenuPoint.y,
        button: "none",
        buttons: 0,
      });
      const redPoint = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const submenu = [...(field?.shadowRoot?.querySelectorAll("menu.swatches-submenu") ?? [])]
          .find((menu) => getComputedStyle(menu).display !== "none");
        const red = [...(submenu?.querySelectorAll(":scope > li") ?? [])]
          .find((item) => item.getAttribute("aria-label") === "red");
        const bounds = red?.getBoundingClientRect();
        return bounds
          ? { ready: true, x: bounds.left + bounds.width / 2, y: bounds.top + bounds.height / 2 }
          : {
              ready: false,
              menuTexts: [...(field?.shadowRoot?.querySelectorAll("menu, menu li") ?? [])]
                .map((node) => ({
                  tag: node.tagName,
                  className: node.className,
                  display: getComputedStyle(node).display,
                  text: node.textContent?.trim() ?? "",
                  aria: node.getAttribute?.("aria-label") ?? "",
                })),
            };
      })()`, "foreground color swatch");
      const collapsedForegroundSelection = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.position = field.lastOffset;
        return {
          collapsed: field.selectionIsCollapsed,
          selection: field.selection,
        };
      })()`);
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: redPoint.x,
        y: redPoint.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: redPoint.x,
        y: redPoint.y,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });
      const foregroundState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        const value = field?.value ?? "";
        return {
          ready:
            field?.queryStyle({ color: "red" }) === "all" &&
            persisted.state?.lines?.[0]?.latex === value,
          value,
          storeValue: persisted.state?.lines?.[0]?.latex ?? "",
          queryStyle: field?.queryStyle({ color: "red" }) ?? "none",
        };
      })()`, "foreground color application");

      await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.selection = {
          ranges: [[0, field.lastOffset]],
          direction: "forward",
        };
      })()`);
      await openContextMenu();
      const backgroundMenuPoint = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const root = field?.shadowRoot?.querySelector("menu.ui-menu-container");
        const backgroundItem = [...(root?.querySelectorAll(":scope > li") ?? [])]
          .find((item) => /^(背景|Background)$/.test(item.textContent?.trim() ?? ""));
        const bounds = backgroundItem?.getBoundingClientRect();
        return bounds
          ? { x: bounds.left + bounds.width / 2, y: bounds.top + bounds.height / 2 }
          : null;
      })()`);
      if (!backgroundMenuPoint) throw new Error("Unable to resolve background color menu item");
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseMoved",
        x: backgroundMenuPoint.x,
        y: backgroundMenuPoint.y,
        button: "none",
        buttons: 0,
      });
      const yellowPoint = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const submenu = [...(field?.shadowRoot?.querySelectorAll("menu.swatches-submenu") ?? [])]
          .find((menu) => getComputedStyle(menu).display !== "none");
        const yellow = [...(submenu?.querySelectorAll(":scope > li") ?? [])]
          .find((item) => item.getAttribute("aria-label") === "yellow");
        const bounds = yellow?.getBoundingClientRect();
        return bounds
          ? { ready: true, x: bounds.left + bounds.width / 2, y: bounds.top + bounds.height / 2 }
          : { ready: false };
      })()`, "background color swatch");
      const collapsedBackgroundSelection = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.position = field.lastOffset;
        return {
          collapsed: field.selectionIsCollapsed,
          selection: field.selection,
        };
      })()`);
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: yellowPoint.x,
        y: yellowPoint.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: yellowPoint.x,
        y: yellowPoint.y,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });
      const backgroundState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        const value = field?.value ?? "";
        return {
          ready:
            field?.queryStyle({ backgroundColor: "yellow" }) === "all" &&
            persisted.state?.lines?.[0]?.latex === value,
          value,
          storeValue: persisted.state?.lines?.[0]?.latex ?? "",
          queryStyle: field?.queryStyle({ backgroundColor: "yellow" }) ?? "none",
        };
      })()`, "background color application");

      if (
        !collapsedForegroundSelection.collapsed ||
        !foregroundState.value.includes("\\textcolor{red}") ||
        !collapsedBackgroundSelection.collapsed ||
        !backgroundState.value.includes("\\textcolor{red}") ||
        !backgroundState.value.includes("\\colorbox{yellow}")
      ) {
        throw new Error(
          `Context style selection restore failed: ${JSON.stringify({
            collapsedForegroundSelection,
            foregroundState,
            collapsedBackgroundSelection,
            backgroundState,
          })}`,
        );
      }

      console.log(
        JSON.stringify(
          { foregroundState, backgroundState },
          null,
          2,
        ),
      );
      console.log("Targeted context style regression passed");
      return;
    }

    if (scenario === "upright") {
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
          shadowText: field?.shadowRoot?.textContent ?? "",
        };
      })()`, "slash derivative uses two upright differential operators");

      await clearField();
      await evaluate(`(() => {
        const field = document.querySelector("math-field");
        if (!field) return false;
        field.setValue("e^{i\\\\theta}", {
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
        return true;
      })()`);
      const exponentialState = await waitForEvaluation(`(() => {
        const value = document.querySelector("math-field")?.value ?? "";
        return {
          ready:
            /\\\\mathrm\\{e\\}/.test(value) &&
            /\\\\mathrm\\{i\\}/.test(value) &&
            /\\\\theta/.test(value),
          value,
        };
      })()`, "Euler constant and imaginary unit remain upright with shortcuts disabled");

      console.log(JSON.stringify({ identifierState, differentialState, exponentialState }, null, 2));
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

      await evaluate(`document.querySelector(".source-toggle")?.click()`);
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
      await evaluate(`document.querySelector(".source-collapse-button")?.click()`);
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

    if (scenario === "settings") {
      const powerPointDefaultInitial = await waitForEvaluation(`(() => {
        let input = document.querySelector(
          '[data-powerpoint-default-font-size]',
        );
        if (!input && !document.querySelector('.settings-dialog')) {
          document.querySelector('.settings-toggle')?.click();
          input = document.querySelector(
            '[data-powerpoint-default-font-size]',
          );
        }
        return {
          ready: input?.value === '20',
          value: input?.value ?? '',
          settingsDialog: Boolean(document.querySelector('.settings-dialog')),
          modalText: document.querySelector('.modal-card')?.textContent?.slice(0, 240) ?? '',
          headerButtons: [...document.querySelectorAll('header button')].map((button) => ({
            className: button.className,
            ariaLabel: button.getAttribute('aria-label') ?? '',
            title: button.getAttribute('title') ?? '',
          })),
        };
      })()`, "PowerPoint default formula font-size setting");
      await evaluate(`(() => {
        const input = document.querySelector(
          '[data-powerpoint-default-font-size]',
        );
        const setter = Object.getOwnPropertyDescriptor(
          HTMLInputElement.prototype,
          'value',
        ).set;
        setter.call(input, '27.5');
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
      })()`);
      const themeChoiceState = await waitForEvaluation(`(() => {
        const rootStyle = getComputedStyle(document.documentElement);
        const swatches = [...document.querySelectorAll('.theme-choice-swatch')];
        const widths = swatches.map((swatch) => swatch.getBoundingClientRect().width);
        const rounded = widths.map((width) => Math.round(width * 10) / 10);
        const distinctWidths = [...new Set(rounded)];
        const surfaceSecondary = rootStyle.getPropertyValue('--surface-secondary').trim();
        const sunken = rootStyle.getPropertyValue('--bg-sunken').trim();
        const paper = rootStyle.getPropertyValue('--bg-paper').trim();
        const canvas = rootStyle.getPropertyValue('--bg-canvas').trim();
        return {
          ready:
            swatches.length >= 10 &&
            distinctWidths.length === 1 &&
            Math.abs((widths[0] ?? 0) - 46) <= 0.6 &&
            Boolean(surfaceSecondary && sunken && paper && canvas),
          count: swatches.length,
          widths: rounded,
          distinctWidths,
          surfaceSecondary,
          sunken,
          paper,
          canvas,
        };
      })()`, "uniform theme swatches and shared theme tokens");

      await evaluate(`document.querySelector('[data-interface-customization-trigger]')?.click()`);
      const interfaceCustomizationLayout = await waitForEvaluation(`(() => {
        const dialog = document.querySelector('[data-interface-customization-dialog]');
        const backdrop = dialog?.parentElement;
        const rect = dialog?.getBoundingClientRect();
        const visible = Boolean(
          rect &&
          rect.left >= -1 &&
          rect.top >= -1 &&
          rect.right <= window.innerWidth + 1 &&
          rect.bottom <= window.innerHeight + 1,
        );
        return {
          ready: Boolean(dialog && backdrop && visible && backdrop.parentElement === document.body),
          visible,
          portalToBody: backdrop?.parentElement === document.body,
          rect: rect ? {
            left: rect.left,
            top: rect.top,
            right: rect.right,
            bottom: rect.bottom,
          } : null,
          settingsOverflow: getComputedStyle(document.querySelector('.settings-dialog')).overflow,
        };
      })()`, "unclipped interface customization portal");
      if (!interfaceCustomizationLayout.visible || !interfaceCustomizationLayout.portalToBody) {
        throw new Error(`Interface customization is still clipped by Settings: ${JSON.stringify(interfaceCustomizationLayout)}`);
      }

      await evaluate(`(() => {
        const input = document.querySelector('[data-formula-inset-left-setting]');
        const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
        setter.call(input, '72');
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
      })()`);
      const insetAndPreviewState = await waitForEvaluation(`(() => {
        const surface = document.querySelector('.editor-surface.multi-line-editor');
        const row = document.querySelector('.formula-line');
        const field = row?.querySelector('math-field');
        const content = field?.shadowRoot?.querySelector('[part="content"]');
        const surfaceRect = surface?.getBoundingClientRect();
        const rowRect = row?.getBoundingClientRect();
        const contentRect = content?.getBoundingClientRect();
        const previewButtons = [...document.querySelectorAll(
          '[data-formula-inset-preview] .formula-inset-preview-toolbar > span',
        )];
        const previewContained = previewButtons.every((button) => {
          const buttonRect = button.getBoundingClientRect();
          const preview = button.querySelector('.math-preview');
          const previewRect = preview?.getBoundingClientRect();
          const visual = preview?.querySelector('.ML__latex');
          const visualRect = visual?.getBoundingClientRect();
          return Boolean(
            previewRect && visualRect &&
            previewRect.left >= buttonRect.left - 1 &&
            previewRect.right <= buttonRect.right + 1 &&
            previewRect.top >= buttonRect.top - 1 &&
            previewRect.bottom <= buttonRect.bottom + 1 &&
            visualRect.left >= buttonRect.left - 1 &&
            visualRect.right <= buttonRect.right + 1 &&
            visualRect.top >= buttonRect.top - 1 &&
            visualRect.bottom <= buttonRect.bottom + 1
          );
        });
        const rowInset = rowRect && surfaceRect ? rowRect.left - surfaceRect.left : -1;
        const contentInset = contentRect && rowRect ? contentRect.left - rowRect.left : -1;
        return {
          ready:
            Boolean(surfaceRect && rowRect && contentRect) &&
            Math.abs(rowInset - 72) <= 2 &&
            contentInset >= -1 && contentInset <= 4 &&
            previewButtons.length === 3 &&
            previewContained,
          rowInset,
          contentInset,
          previewCount: previewButtons.length,
          previewContained,
        };
      })()`, "formula inset controls move MathLive content and preview stays contained");
      if (!insetAndPreviewState.previewContained) {
        throw new Error(`Settings live preview overflowed its toolbar tiles: ${JSON.stringify(insetAndPreviewState)}`);
      }
      await evaluate(`document.querySelector('[data-interface-customization-close]')?.click()`);
      await waitForEvaluation(`(() => ({
        ready: !document.querySelector('[data-interface-customization-dialog]'),
      }))()`, "close interface customization portal");

      await evaluate(`document.querySelector('[data-interface-customization-trigger]')?.click()`);
      await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector('[data-interface-customization-dialog]')),
      }))()`, "reopen interface customization for font preview");
      await evaluate(`(() => {
        const letter = document.querySelector('[data-formula-letter-font-setting]');
        const chinese = document.querySelector('[data-formula-chinese-font-setting]');
        letter.value = 'helvetica';
        letter.dispatchEvent(new Event('change', { bubbles: true }));
        chinese.value = 'kaiti';
        chinese.dispatchEvent(new Event('change', { bubbles: true }));
      })()`);
      const formulaFontPreviewState = await waitForEvaluation(`(() => {
        const preview = document.querySelector('[data-formula-font-preview]');
        const mathLetter = preview?.querySelector('.ML__mathit, .ML__latin');
        const chineseText = [...(preview?.querySelectorAll('.ML__text, .ML__textord') ?? [])]
          .find((node) => (node.textContent ?? '').includes('中文'));
        const letterFamily = mathLetter ? getComputedStyle(mathLetter).fontFamily : '';
        const chineseFamily = chineseText ? getComputedStyle(chineseText).fontFamily : '';
        const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}');
        return {
          ready:
            Boolean(mathLetter && chineseText) &&
            /Arial|Helvetica/i.test(letterFamily) &&
            /KaiTi|Kaiti/i.test(chineseFamily) &&
            persisted.state?.formulaLetterFont === 'helvetica' &&
            persisted.state?.formulaChineseFont === 'kaiti',
          letterFamily,
          chineseFamily,
          letterSetting: persisted.state?.formulaLetterFont ?? '',
          chineseSetting: persisted.state?.formulaChineseFont ?? '',
        };
      })()`, "formula font live preview follows selected western and Chinese fonts");
      if (!/Arial|Helvetica/i.test(formulaFontPreviewState.letterFamily)
        || !/KaiTi|Kaiti/i.test(formulaFontPreviewState.chineseFamily)) {
        throw new Error(`Formula font preview ignored selected fonts: ${JSON.stringify(formulaFontPreviewState)}`);
      }
      await evaluate(`document.querySelector('[data-interface-customization-close]')?.click()`);
      await waitForEvaluation(`(() => ({
        ready: !document.querySelector('[data-interface-customization-dialog]'),
      }))()`, "close interface customization after font preview");

      const powerPointDefaultSaved = await waitForEvaluation(`(() => {
        const input = document.querySelector(
          '[data-powerpoint-default-font-size]',
        );
        const persisted = JSON.parse(
          localStorage.getItem('visualtex-editor') || '{}',
        );
        return {
          ready:
            input?.value === '27.5' &&
            persisted.state?.powerPointDefaultFontSizePt === 27.5,
          value: input?.value ?? '',
          persisted: persisted.state?.powerPointDefaultFontSizePt ?? null,
        };
      })()`, "persisted PowerPoint default formula font size");
      await evaluate(`document.querySelector(
        'button[aria-label="关闭设置"], button[aria-label="Close settings"]',
      ).click()`);
      await waitForEvaluation(`(() => ({
        ready: !document.querySelector('.settings-dialog'),
      }))()`, "closed main settings dialog");

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

      await evaluate(`document.querySelector("[data-open-auto-escape-map]")?.click()`);
      const autoEscapeMapState = await waitForEvaluation(`(() => {
        const read = (shortcut) => {
          const row = document.querySelector(
            '[data-auto-escape-shortcut="' + CSS.escape(shortcut) + '"]',
          );
          return row
            ? {
                shortcut: row.dataset.autoEscapeShortcut ?? '',
                output: row.dataset.autoEscapeOutput ?? '',
                after: row.dataset.autoEscapeAfter ?? '',
                rendered: Boolean(row.querySelector('.auto-escape-map-output-formula .ML__latex')),
              }
            : null;
        };
        const pp = read('pp');
        const relation = read('>=');
        const alpha = read('alpha');
        const hat = read('hat');
        const dx = read('dx');
        const firstInput = document.querySelector('.auto-escape-map-input');
        const firstOutput = document.querySelector('.auto-escape-map-output');
        const firstRow = document.querySelector('.auto-escape-map-row');
        const inputFontSize = Number.parseFloat(
          firstInput ? getComputedStyle(firstInput).fontSize : '0',
        );
        const outputFontSize = Number.parseFloat(
          firstOutput ? getComputedStyle(firstOutput).fontSize : '0',
        );
        const rowHeight = firstRow?.getBoundingClientRect().height ?? 0;
        return {
          ready:
            Boolean(document.querySelector('.input-behavior-popover.is-mapping-view')) &&
            pp?.output === '+' &&
            relation?.output === '\\\\ge' &&
            alpha?.output === '\\\\alpha' &&
            hat?.output === '\\\\hat{#?}' &&
            dx?.output === '\\\\mathrm{d}x' &&
            dx.after.includes('nothing') &&
            [pp, relation, alpha, hat, dx].every((entry) => entry?.rendered) &&
            inputFontSize >= 13 &&
            outputFontSize >= 20 &&
            rowHeight >= 42,
          pp,
          relation,
          alpha,
          hat,
          dx,
          inputFontSize,
          outputFontSize,
          rowHeight,
          count: document.querySelectorAll('[data-auto-escape-shortcut]').length,
        };
      })()`, "source-driven auto-escape mapping list");
      if (autoEscapeMapState.pp.output !== '+') {
        throw new Error(`pp mapping is not the real shortcut value: ${JSON.stringify(autoEscapeMapState)}`);
      }
      await evaluate(`document.querySelector('button[aria-label="返回操作逻辑"], button[aria-label="Back to input behavior"]')?.click()`);
      await waitForEvaluation(`(() => ({
        ready:
          Boolean(document.querySelector('.input-behavior-popover')) &&
          !document.querySelector('.input-behavior-popover.is-mapping-view'),
      }))()`, "return from auto-escape mappings");

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
      await clearField();
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

      const powerPointDefaultReloaded = await waitForEvaluation(`(() => {
        const persisted = JSON.parse(
          localStorage.getItem('visualtex-editor') || '{}',
        );
        return {
          ready: persisted.state?.powerPointDefaultFontSizePt === 27.5,
          persisted: persisted.state?.powerPointDefaultFontSizePt ?? null,
        };
      })()`, "PowerPoint default font size survives reload");

      console.log(JSON.stringify({
        powerPointDefaultInitial,
        themeChoiceState,
        interfaceCustomizationLayout,
        powerPointDefaultSaved,
        powerPointDefaultReloaded,
        defaults,
        autoEscapeMapState,
        structuredState,
        otherState,
      }, null, 2));
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
        ready: Boolean(behavior && fileActions && editActions && titleInput),
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
      toolbarOrder.behaviorIndex < 0 ||
      (toolbarOrder.exportIndex >= 0 &&
        toolbarOrder.exportIndex >= toolbarOrder.behaviorIndex) ||
      toolbarOrder.fileBorderWidth !== "0px" ||
      (toolbarOrder.exportIndex >= 0 && toolbarOrder.fileLeftOffset !== "6px") ||
      toolbarOrder.titleRight > toolbarOrder.fileLeft ||
      toolbarOrder.editLeft - toolbarOrder.fileRight < 4
    ) {
      throw new Error(`Incorrect export/header placement: ${JSON.stringify(toolbarOrder)}`);
    }

    let exportMenuState = null;
    if (toolbarOrder.exportIndex >= 0) {
      await evaluate(`document.querySelector(".export-menu-trigger")?.click()`);
      exportMenuState = await waitForEvaluation(`(() => {
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
    }

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
      const field = document.querySelector(".formula-line.is-active math-field");
      const value = field?.value ?? "";
      const body = value.match(/\\\\begin\\{bmatrix\\}([\\s\\S]*?)\\\\end\\{bmatrix\\}/)?.[1] ?? "";
      const root = field?.shadowRoot;
      const symbol = field?.placeholderSymbol || "▢";
      const placeholders = [...(root?.querySelectorAll(
        ".visualtex-structural-placeholder",
      ) ?? [])];
      const rawLeaves = [...(root?.querySelectorAll("[data-atom-id]") ?? [])]
        .filter(
          (node) =>
            (node.textContent || "").trim() === symbol &&
            !node.querySelector("[data-atom-id]") &&
            !node.classList.contains("visualtex-structural-placeholder"),
        );
      return {
        ready:
          value.includes("\\\\begin{bmatrix}") &&
          body.split(/\\\\\\\\/).length === 3 &&
          body.split(/\\\\\\\\/).every((row) => row.split("&").length === 4) &&
          placeholders.length === 12 &&
          rawLeaves.length === 0,
        value,
        placeholderCount: placeholders.length,
        rawPlaceholderCount: rawLeaves.length,
        placeholderClasses: [...new Set(placeholders.map((node) => node.className))],
      };
    })()`, "3 by 4 matrix insertion with VisualTeX placeholders");

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
