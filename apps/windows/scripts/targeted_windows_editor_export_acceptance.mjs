import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import process from "node:process";

const portOffset = process.pid % 1000;
const previewPort = 6200 + portOffset;
const debugPort = 11200 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = join(tmpdir(), `visualtex-windows-targeted-${process.pid}`);
const artifactRoot = resolve(
  process.cwd(),
  "artifacts",
  `windows-editor-export-${new Date().toISOString().replace(/[:.]/g, "-")}`,
);
const chromeCandidates = [
  process.env.VISUALTEX_CHROME_PATH,
  process.env.PROGRAMFILES && join(process.env.PROGRAMFILES, "Google", "Chrome", "Application", "chrome.exe"),
  process.env["PROGRAMFILES(X86)"] && join(process.env["PROGRAMFILES(X86)"], "Google", "Chrome", "Application", "chrome.exe"),
  process.env.LOCALAPPDATA && join(process.env.LOCALAPPDATA, "Google", "Chrome", "Application", "chrome.exe"),
  process.env.PROGRAMFILES && join(process.env.PROGRAMFILES, "Microsoft", "Edge", "Application", "msedge.exe"),
  process.env["PROGRAMFILES(X86)"] && join(process.env["PROGRAMFILES(X86)"], "Microsoft", "Edge", "Application", "msedge.exe"),
].filter(Boolean);
const chromePath = chromeCandidates.find((candidate) => existsSync(candidate));
if (!chromePath) {
  throw new Error(`No Chrome/Edge executable found. Checked: ${chromeCandidates.join(", ")}`);
}

const sleep = (ms) => new Promise((resolvePromise) => setTimeout(resolvePromise, ms));

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
    await new Promise((resolvePromise, reject) => {
      this.socket.addEventListener("open", resolvePromise, { once: true });
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
    return new Promise((resolvePromise, reject) => {
      this.pending.set(id, { resolve: resolvePromise, reject });
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }

  close() {
    this.socket?.close();
  }
}

async function main() {
  await mkdir(artifactRoot, { recursive: true });
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
        "--window-size=1500,1050",
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

    const waitForEvaluation = async (expression, description, timeoutMs = 6000) => {
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

    const key = async (
      value,
      code,
      virtualKeyCode,
      includeText = value.length === 1,
    ) => {
      const common = {
        key: value,
        code,
        windowsVirtualKeyCode: virtualKeyCode,
        nativeVirtualKeyCode: virtualKeyCode,
      };
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        ...common,
        ...(includeText ? { text: value, unmodifiedText: value } : {}),
      });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(50);
    };

    const typeText = async (text) => {
      for (const character of text) {
        const code = character === "\\" ? "Backslash" : `Key${character.toUpperCase()}`;
        const virtualKeyCode = character === "\\" ? 220 : character.toUpperCase().charCodeAt(0);
        await key(character, code, virtualKeyCode);
      }
    };

    const screenshot = async (name) => {
      const capture = await client.send("Page.captureScreenshot", {
        format: "png",
        captureBeyondViewport: false,
      });
      await writeFile(join(artifactRoot, name), Buffer.from(capture.data, "base64"));
    };

    const resetDocument = async (values = [""], activeIndex = 0, inputBehavior) => {
      await evaluate(`(() => {
        const storageKey = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
        const values = ${JSON.stringify(values)};
        const lines = values.map((latex) => ({ id: crypto.randomUUID(), latex }));
        persisted.state = {
          ...(persisted.state || {}),
          lines,
          activeLineId: lines[${activeIndex}]?.id ?? lines[0]?.id ?? null,
          ...(typeof ${JSON.stringify(inputBehavior)} === "object" && ${JSON.stringify(inputBehavior)} !== null
            ? { inputBehavior: ${JSON.stringify(inputBehavior)} }
            : {}),
        };
        localStorage.setItem(storageKey, JSON.stringify(persisted));
        localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      })()`);
      await client.send("Page.reload", { ignoreCache: true });
      await waitForEvaluation(
        `(() => ({
          ready:
            Boolean(document.querySelector("math-field")) &&
            Boolean(document.querySelector(".canvas-input-behavior-trigger")) &&
            Boolean(document.querySelector('button[data-category="matrix"]')),
        }))()`,
        "formula field and toolbar controls after reset",
      );
      await sleep(150);
    };

    const focusField = async (index = 0) => {
      await waitForEvaluation(`(() => {
        const field = document.querySelectorAll("math-field")[${index}];
        if (!field?.isConnected) return { ready: false };
        field.focus();
        field.position = field.lastOffset;
        const sink = field.shadowRoot?.querySelector('[part="keyboard-sink"]');
        sink?.focus({ preventScroll: true });
        return {
          ready:
            field.matches(":focus-within") &&
            Boolean(sink) &&
            field.shadowRoot?.activeElement === sink,
          activePart: field.shadowRoot?.activeElement?.getAttribute?.("part") ?? "",
        };
      })()`, `focus field ${index}`);
      await sleep(80);
    };

    const insertFieldText = async (text) => {
      await evaluate(`(async () => {
        const field = document.querySelector("math-field");
        if (!field) throw new Error("No math-field is available for text insertion");
        for (const character of Array.from(${JSON.stringify(text)})) {
          const beforeInput = new InputEvent("beforeinput", {
            bubbles: true,
            composed: true,
            cancelable: true,
            inputType: "insertText",
            data: character,
          });
          field.dispatchEvent(beforeInput);
          if (!beforeInput.defaultPrevented) {
            field.insert(character, {
              mode: "math",
              format: "latex",
              insertionMode: "replaceSelection",
              selectionMode: "after",
              focus: true,
              scrollIntoView: false,
            });
            field.dispatchEvent(new InputEvent("input", {
              bubbles: true,
              composed: true,
              inputType: "insertText",
              data: character,
            }));
          }
          await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        }
      })()`);
      await sleep(80);
    };

    const clearField = async () => {
      await evaluate(`(() => {
        const field = document.querySelector("math-field");
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
      })()`);
      await sleep(120);
      await focusField();
    };

    await client.send("Page.navigate", { url: baseUrl });
    await resetDocument();

    const toolbarOrder = await waitForEvaluation(`(() => {
      const group = document.querySelector(".canvas-tool-group");
      const exportButton = group?.querySelector(".workspace-export-trigger");
      const behavior = group?.querySelector(".input-behavior-menu");
      const children = [...(group?.children ?? [])];
      return {
        ready: Boolean(exportButton && behavior),
        exportIndex: children.indexOf(exportButton),
        behaviorIndex: children.indexOf(behavior),
        label: exportButton?.textContent?.trim() ?? "",
      };
    })()`, "Export and input behavior buttons");
    if (toolbarOrder.exportIndex >= toolbarOrder.behaviorIndex) {
      throw new Error(`Export button must precede input behavior: ${JSON.stringify(toolbarOrder)}`);
    }

    await evaluate(`document.querySelector(".workspace-export-trigger").click()`);
    const exportDialog = await waitForEvaluation(`(() => {
      const dialog = document.querySelector(".export-dialog");
      const options = [...document.querySelectorAll(".export-format-option")];
      const input = document.querySelector(".export-path-field input");
      return {
        ready: Boolean(dialog && options.length === 3 && input),
        options: options.map((option) => option.textContent.replace(/\s+/g, " ").trim()),
        inputPlaceholder: input?.placeholder ?? "",
      };
    })()`, "three-format Export dialog");
    await evaluate(`(() => {
      const input = document.querySelector(".export-path-field input");
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value").set;
      setter.call(input, "C:\\\\VisualTeX Acceptance\\\\formula.md");
      input.dispatchEvent(new Event("input", { bubbles: true }));
      document.querySelectorAll(".export-format-option")[1].click();
    })()`);
    const exportPathState = await waitForEvaluation(`(() => {
      const input = document.querySelector(".export-path-field input");
      return {
        ready: input?.value.endsWith("formula.svg"),
        value: input?.value ?? "",
        selected: document.querySelector(".export-format-option.is-active strong")?.textContent ?? "",
      };
    })()`, "manual export path extension update");
    await screenshot("01-export-dialog.png");
    console.log("STAGE export-dialog PASS");
    await evaluate(`document.querySelector('.export-dialog button[aria-label="关闭导出窗口"]').click()`);

    await resetDocument();
    await focusField();
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.setValue("\\\\frac{d\\\\Phi}{d\\\\theta}", {
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
    const liveDifferentialUpright = await waitForEvaluation(`(() => {
      const field = document.querySelector("math-field");
      return {
        ready: field?.value === "\\\\frac{\\\\mathrm{d}\\\\Phi}{\\\\mathrm{d}\\\\theta}",
        value: field?.value ?? "",
      };
    })()`, "immediate generic differential upright write-back");
    console.log("STAGE live-differential-upright PASS");

    await resetDocument();
    await waitForEvaluation(`(() => ({
      ready: Boolean(document.querySelector('button[data-category="matrix"]')),
    }))()`, "matrix category button after document reset");
    await evaluate(`document.querySelector('button[data-category="matrix"]')?.click()`);
    const matrixGrid = await waitForEvaluation(`(() => ({
      ready: document.querySelectorAll(".matrix-size-cell").length === 100,
      count: document.querySelectorAll(".matrix-size-cell").length,
    }))()`, "10 by 10 matrix grid");
    await evaluate(`(() => {
      const cell = document.querySelector('.matrix-size-cell[data-matrix-rows="3"][data-matrix-columns="4"]');
      cell.dispatchEvent(new PointerEvent("pointerenter", { bubbles: true }));
      cell.focus();
    })()`);
    const matrixPreview = await waitForEvaluation(`(() => ({
      ready:
        document.querySelector(".matrix-size-badge")?.textContent?.replace(/\s+/g, " ").trim() === "3 × 4" &&
        document.querySelectorAll(".matrix-size-cell.is-previewed").length === 12,
      badge: document.querySelector(".matrix-size-badge")?.textContent?.replace(/\s+/g, " ").trim() ?? "",
      previewed: document.querySelectorAll(".matrix-size-cell.is-previewed").length,
    }))()`, "3 by 4 matrix preview");
    await evaluate(`document.querySelector('.matrix-size-cell[data-matrix-rows="3"][data-matrix-columns="4"]')?.click()`);
    await waitForEvaluation(`(() => ({
      ready:
        Boolean(document.querySelector(".matrix-insert-button")) &&
        document.querySelector(".matrix-size-badge")?.textContent?.replace(/\\s+/g, " ").trim() === "3 × 4",
    }))()`, "matrix insert button after selecting 3 by 4");
    await evaluate(`document.querySelector(".matrix-insert-button")?.click()`);
    console.log("STAGE matrix-picker PASS");
    const matrixInsertion = await waitForEvaluation(`(() => {
      const value = document.querySelector("math-field")?.value ?? "";
      const body = value.match(/\\\\begin\\{bmatrix\\}([\\s\\S]*?)\\\\end\\{bmatrix\\}/)?.[1] ?? "";
      return {
        ready:
          body.split(/\\\\\\\\/).length === 3 &&
          body.split(/\\\\\\\\/).every((row) => row.split("&").length === 4),
        value,
      };
    })()`, "3 by 4 matrix insertion");

    await resetDocument();
    await waitForEvaluation(`(() => {
      const button = document.querySelector(".canvas-input-behavior-trigger");
      if (!button) return { ready: false };
      button.click();
      return { ready: true };
    })()`, "open input behavior menu after document reset");
    const settingDefaults = await waitForEvaluation(`(() => {
      const options = [...document.querySelectorAll(".input-behavior-option")].map((label) => ({
        title: label.querySelector("strong")?.textContent ?? "",
        checked: label.querySelector('input[type="checkbox"]')?.checked ?? false,
      }));
      const structured = options.find((item) => item.title.includes("求和、积分"));
      const other = options.find((item) => item.title.includes("其他命令"));
      return { ready: Boolean(structured && other), structured, other, options };
    })()`, "input behavior defaults");
    console.log("STAGE matrix-insertion PASS");
    if (!settingDefaults.structured.checked || settingDefaults.other.checked) {
      throw new Error(`Unexpected candidate defaults: ${JSON.stringify(settingDefaults)}`);
    }
    await evaluate(`document.querySelector(".canvas-input-behavior-trigger").click()`);

    await focusField();
    await typeText("\\sum");
    const structuredCandidate = await waitForEvaluation(`(() => ({
      ready: Boolean(document.querySelector(".suggestion-popup")),
      customVisible: Boolean(document.querySelector(".suggestion-popup")),
      query: document.querySelector(".multi-line-editor")?.dataset.commandQuery ?? "",
    }))()`, "structured VisualTeX candidate");

    console.log("STAGE structured-candidate PASS");
    await resetDocument();
    await focusField();
    await typeText("\\theta");
    const otherCandidate = await waitForEvaluation(`(() => {
      const nativePanel = document.getElementById("mathlive-suggestion-popover");
      return {
        ready:
          !document.querySelector(".suggestion-popup") &&
          Boolean(nativePanel?.classList.contains("is-visible")),
        customVisible: Boolean(document.querySelector(".suggestion-popup")),
        nativeVisible: nativePanel?.classList.contains("is-visible") ?? false,
      };
    })()`, "native-only other-command candidate");

    console.log("STAGE native-other-candidate PASS");
    await resetDocument();
    await focusField();
    await typeText("\\mathbb");
    await key(" ", "Space", 32, false);
    const wrapperPlaceholder = await waitForEvaluation(`(() => {
      const field = document.querySelector("math-field");
      const host = field?.closest(".mathfield-host");
      const caret = field?.shadowRoot?.querySelector(".ML__caret");
      const caretBounds = caret?.getBoundingClientRect();
      const hostBounds = host?.getBoundingClientRect();
      const left = Number.parseFloat(host?.style.getPropertyValue("--pending-wrapper-left") ?? "NaN");
      const top = Number.parseFloat(host?.style.getPropertyValue("--pending-wrapper-top") ?? "NaN");
      const expectedLeft = caretBounds && hostBounds ? caretBounds.left - hostBounds.left : Number.NaN;
      const expectedTop = caretBounds && hostBounds
        ? caretBounds.top - hostBounds.top + caretBounds.height / 2
        : Number.NaN;
      const pseudo = host ? getComputedStyle(host, "::after") : null;
      const width = Number.parseFloat(pseudo?.width ?? "NaN");
      const borderWidth = Number.parseFloat(pseudo?.borderTopWidth ?? "NaN");
      return {
        ready:
          !field?.value.includes("\\\\left") &&
          field?.dataset.pendingWrapperCommand === "\\\\mathbb" &&
          host?.classList.contains("has-pending-wrapper-placeholder") &&
          width >= 18 &&
          borderWidth > 0 &&
          Math.abs(left - expectedLeft) <= 1.5 &&
          Math.abs(top - expectedTop) <= 1.5,
        value: field?.value ?? "",
        pending: field?.dataset.pendingWrapperCommand ?? "",
        left,
        top,
        width,
        borderWidth,
        expectedLeft,
        expectedTop,
      };
    })()`, "mathbb placeholder anchored to native caret");
    await screenshot("02-mathbb-placeholder.png");
    console.log("STAGE mathbb-placeholder PASS");
    await key("A", "KeyA", 65);
    console.log("STAGE mathbb-character typed");
    const wrapperAutoExit = await waitForEvaluation(`(() => ({
      ready:
        document.querySelector("math-field")?.value === "\\\\mathbb{A}" &&
        !document.querySelector("math-field")?.dataset.pendingWrapperCommand,
      value: document.querySelector("math-field")?.value ?? "",
    }))()`, "mathbb default one-character auto exit");
    await key("B", "KeyB", 66);
    const wrapperNormalFont = await waitForEvaluation(`(() => ({
      ready: document.querySelector("math-field")?.value === "\\\\mathbb{A}B",
      value: document.querySelector("math-field")?.value ?? "",
    }))()`, "normal font after mathbb auto exit");

    console.log("STAGE mathbb-auto-exit PASS");
    await resetDocument();
    await focusField();
    await typeText("\\mathcal");
    await key(" ", "Space", 32, false);
    await key("g", "KeyG", 71);
    const lowercaseMathcal = await waitForEvaluation(`(() => ({
      ready: document.querySelector("math-field")?.value === "\\\\mathscr{g}",
      value: document.querySelector("math-field")?.value ?? "",
      shadowText: document.querySelector("math-field")?.shadowRoot?.textContent ?? "",
    }))()`, "lowercase mathcal compatibility");

    await evaluate(`document.querySelector(".canvas-input-behavior-trigger").click()`);
    await waitForEvaluation(`(() => ({ ready: Boolean(document.querySelector(".input-behavior-popover")) }))()`, "input behavior menu");
    await evaluate(`(() => {
      const option = [...document.querySelectorAll(".input-behavior-option")]
        .find((label) => label.querySelector("strong")?.textContent?.includes("字体命令输入后跳出"));
      option?.querySelector('input[type="checkbox"]')?.click();
      document.querySelector(".canvas-input-behavior-trigger").click();
    })()`);
    console.log("STAGE lowercase-mathcal PASS");
    await clearField();
    await typeText("\\mathbb");
    await key(" ", "Space", 32, false);
    await key("A", "KeyA", 65);
    await key("B", "KeyB", 66);
    const wrapperContinuous = await waitForEvaluation(`(() => {
      const field = document.querySelector("math-field");
      const host = field?.closest(".mathfield-host");
      const width = Number.parseFloat(host?.style.getPropertyValue("--pending-wrapper-width") ?? "NaN");
      return {
        ready:
          field?.value === "\\\\mathbb{AB}" &&
          field.dataset.pendingWrapperCommand === "\\\\mathbb" &&
          width > 18,
        value: field?.value ?? "",
        pending: field?.dataset.pendingWrapperCommand ?? "",
        width,
      };
    })()`, "continuous wrapper input with content-sized frame when auto-exit is disabled");
    await key("Enter", "Enter", 13);
    const wrapperEnterExit = await waitForEvaluation(`(() => ({
      ready:
        document.querySelector("math-field")?.value === "\\\\mathbb{AB}" &&
        !document.querySelector("math-field")?.dataset.pendingWrapperCommand &&
        !document.querySelector(".mathfield-host")?.classList.contains("has-pending-wrapper-placeholder"),
      value: document.querySelector("math-field")?.value ?? "",
      pending: document.querySelector("math-field")?.dataset.pendingWrapperCommand ?? "",
    }))()`, "Enter exits continuous wrapper input");
    await key("C", "KeyC", 67);
    const wrapperNormalAfterEnter = await waitForEvaluation(`(() => ({
      ready: document.querySelector("math-field")?.value === "\\\\mathbb{AB}C",
      value: document.querySelector("math-field")?.value ?? "",
    }))()`, "normal font input after Enter exits wrapper");

    console.log("STAGE wrapper-continuous-enter-exit PASS");
    await resetDocument();
    await focusField();
    await typeText("\\mat");
    const rawBeforeDelete = await waitForEvaluation(`(() => {
      const field = document.querySelector("math-field");
      const raw = [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
        .map((node) => node.textContent ?? "").join("");
      return { ready: raw.includes("\\\\mat"), raw, rows: document.querySelectorAll("math-field").length };
    })()`, "raw command before delete");
    await key("Backspace", "Backspace", 8);
    const rawFirstDelete = await waitForEvaluation(`(() => {
      const field = document.querySelector("math-field");
      const raw = [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
        .map((node) => node.textContent ?? "").join("");
      return { ready: raw.includes("\\\\ma") && !raw.includes("\\\\mat") && document.querySelectorAll("math-field").length === 1, raw };
    })()`, "one-character raw command deletion");
    await key("Backspace", "Backspace", 8);
    const rawSecondDelete = await waitForEvaluation(`(() => {
      const field = document.querySelector("math-field");
      const raw = [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
        .map((node) => node.textContent ?? "").join("");
      return { ready: raw.includes("\\\\m") && !raw.includes("\\\\ma") && document.querySelectorAll("math-field").length === 1, raw };
    })()`, "second one-character raw command deletion");

    console.log("STAGE raw-delete PASS");
    await resetDocument(["a", ""], 1);
    await focusField(1);
    await key("Backspace", "Backspace", 8);
    const emptyRowDelete = await waitForEvaluation(`(() => {
      const fields = [...document.querySelectorAll("math-field")];
      const surface = document.querySelector(".multi-line-editor");
      const firstRow = document.querySelector(".formula-line");
      const firstField = fields[0];
      return {
        ready:
          fields.length === 1 &&
          firstField?.value === "a" &&
          firstField.matches(":focus-within") &&
          firstField.position === firstField.lastOffset &&
          surface?.dataset.activeLineId === firstRow?.dataset.lineId,
        count: fields.length,
        value: firstField?.value ?? "",
        focused: firstField?.matches(":focus-within") ?? false,
        position: firstField?.position ?? -1,
        lastOffset: firstField?.lastOffset ?? -1,
      };
    })()`, "empty row deletion returns to previous formula");

    console.log("STAGE empty-row-delete PASS");
    // Superscript/subscript auto-exit is validated in the release Tauri/WebView2
    // acceptance script. Synthetic Chrome beforeinput events do not perform the
    // native placeholder replacement consistently and can produce false failures.

    await resetDocument();
    await focusField();
    await evaluate(`document.querySelector('button[data-command-id="hat"]').click()`);
    await waitForEvaluation(`(() => {
      const field = document.querySelector("math-field");
      if (!field?.value.includes("\\\\placeholder")) return { ready: false, value: field?.value ?? "" };
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      return { ready: field.matches(":focus-within"), value: field.value };
    })()`, "focused accent placeholder");
    await sleep(100);
    await insertFieldText("a");
    const accentAfterFirstCharacter = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      return {
        value: field?.value ?? "",
        storedValues: (persisted.state?.lines ?? []).map((line) => line.latex),
        position: field?.position ?? -1,
        lastOffset: field?.lastOffset ?? -1,
        selection: field?.selection ?? null,
      };
    })()`);
    await insertFieldText("b");
    const accentAutoExit = await waitForEvaluation(`(() => ({
      ready: document.querySelector("math-field")?.value === "\\\\hat{a}b",
      value: document.querySelector("math-field")?.value ?? "",
    }))()`, "accent auto exit");
    await screenshot("03-final-editor-state.png");

    const report = {
      chromePath,
      toolbarOrder,
      exportDialog,
      exportPathState,
      matrixGrid,
      matrixPreview,
      matrixInsertion,
      settingDefaults,
      structuredCandidate,
      otherCandidate,
      wrapperPlaceholder,
      wrapperAutoExit,
      wrapperNormalFont,
      lowercaseMathcal,
      wrapperContinuous,
      rawBeforeDelete,
      rawFirstDelete,
      rawSecondDelete,
      emptyRowDelete,
      accentAfterFirstCharacter,
      accentAutoExit,
      artifactRoot,
    };
    await writeFile(
      join(artifactRoot, "acceptance-report.json"),
      `${JSON.stringify(report, null, 2)}\n`,
      "utf8",
    );
    console.log(JSON.stringify(report, null, 2));
    console.log("Targeted Windows editor/export acceptance passed");
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
