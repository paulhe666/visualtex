import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const portOffset = process.pid % 1000;
const previewPort = 4300 + portOffset;
const debugPort = 9300 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}/editor`;
const chromeProfile = `/tmp/visualtex-history-smoke-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return response;
    } catch {
      // Wait for the local process.
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
      const targets = await (
        await fetch(`http://127.0.0.1:${debugPort}/json/list`)
      ).json();
      page = targets.find(
        (target) => target.type === "page" && target.url.startsWith(baseUrl),
      );
      if (!page) await sleep(100);
    }
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

    const key = async (
      value,
      code,
      virtualKeyCode,
      modifiers = 0,
      includeText = value.length === 1 && modifiers === 0,
    ) => {
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
        ...(includeText ? { text: value, unmodifiedText: value } : {}),
      });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(45);
    };

    const waitForFields = async (count = 1) => {
      await evaluate(`new Promise((resolve, reject) => {
        const started = performance.now();
        const done = () => {
          if (document.querySelectorAll("math-field").length >= ${count}) {
            resolve(true);
            return;
          }
          if (performance.now() - started > 5000) {
            reject(new Error("Formula fields did not mount"));
            return;
          }
          setTimeout(done, 30);
        };
        done();
      })`);
    };

    const resetDocument = async ({
      lines,
      activeLineId = lines[0].id,
      sourceOpen = false,
      latexCodeFormat = "raw",
      history = [],
    }) => {
      await evaluate(`(() => {
        localStorage.setItem("visualtex.onboarding.v3.completed", "true");
        localStorage.setItem("visualtex.onboarding.web.v3.completed", "true");
        let persisted;
        try {
          persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "null");
        } catch {
          persisted = null;
        }
        if (!persisted || typeof persisted !== "object") {
          persisted = { state: {}, version: 0 };
        }
        persisted.state = {
          ...(persisted.state || {}),
          title: "History Test",
          lines: ${JSON.stringify(lines)},
          activeLineId: ${JSON.stringify(activeLineId)},
          sourceOpen: ${JSON.stringify(sourceOpen)},
          latexCodeFormat: ${JSON.stringify(latexCodeFormat)},
          language: "cn",
          history: ${JSON.stringify(history)},
        };
        delete persisted.state.latex;
        localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
      })()`);
      await client.send("Page.reload", { ignoreCache: true });
      await sleep(850);
      await waitForFields(lines.length);
      await waitForValues(lines.map((line) => line.latex));
    };

    const installFakeTauri = async () => {
      await evaluate(`(() => {
        let callbackId = 1;
        const callbacks = new Map();
        window.__TAURI_INTERNALS__ = {
          transformCallback(callback, once = false) {
            const id = callbackId++;
            callbacks.set(id, { callback, once });
            return id;
          },
          unregisterCallback(id) {
            callbacks.delete(id);
          },
          async invoke(command) {
            if (command === "get_ocr_runtime_status") {
              return {
                installed: true,
                pythonPath: "/fake/python",
                pythonVersion: "3.13.0",
                paddleVersion: "3.3.1",
                paddleocrVersion: "3.7.0",
                runtimePath: "/fake/runtime",
                message: "Fake OCR runtime ready",
              };
            }
            if (command === "recognize_formula_image") {
              return {
                model: "PP-FormulaNet_plus-S",
                elapsedMs: 5,
                processedWidth: 64,
                processedHeight: 32,
                backgroundInverted: false,
                backgroundLuminance: 255,
                formulas: [{ latex: "\\\\theta" }],
              };
            }
            if (
              command === "plugin:event|listen" ||
              command === "plugin:event|unlisten"
            ) {
              return 1;
            }
            if (
              command === "cancel_ocr_recognition" ||
              command === "restart_ocr_worker"
            ) {
              return null;
            }
            throw new Error("Unexpected fake Tauri command: " + command);
          },
        };
      })()`);
    };

    const focusField = async (index) => {
      await evaluate(`(() => {
        const field = document.querySelectorAll("math-field")[${index}];
        if (!field) throw new Error("Missing formula field ${index}");
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      })()`);
      await sleep(80);
    };

    const click = async (selector) => {
      await evaluate(`(() => {
        const element = document.querySelector(${JSON.stringify(selector)});
        if (!element) {
          throw new Error("Missing element: " + ${JSON.stringify(selector)});
        }
        element.click();
      })()`);
      await sleep(180);
    };

    const values = async () =>
      evaluate(`[...document.querySelectorAll("math-field")].map((field) => field.value)`);

    const waitForValues = async (expected, timeoutMs = 3000) => {
      const expectedJson = JSON.stringify(expected);
      await evaluate(`new Promise((resolve, reject) => {
        const started = performance.now();
        const done = () => {
          const current = [...document.querySelectorAll("math-field")].map(
            (field) => field.value,
          );
          if (JSON.stringify(current) === ${JSON.stringify(expectedJson)}) {
            resolve(true);
            return;
          }
          if (performance.now() - started > ${timeoutMs}) {
            reject(new Error(
              "Timed out waiting for formula values: " + JSON.stringify(current),
            ));
            return;
          }
          setTimeout(done, 30);
        };
        done();
      })`);
    };

    const waitForHistoryAction = async (ariaLabel, timeoutMs = 3000) => {
      const selector = `button[aria-label="${ariaLabel}"]`;
      await evaluate(`new Promise((resolve, reject) => {
        const started = performance.now();
        const done = () => {
          const button = document.querySelector(${JSON.stringify(selector)});
          if (button && !button.disabled) {
            resolve(true);
            return;
          }
          if (performance.now() - started > ${timeoutMs}) {
            reject(new Error(
              "Timed out waiting for history action: " + ${JSON.stringify(ariaLabel)},
            ));
            return;
          }
          setTimeout(done, 30);
        };
        done();
      })`);
    };

    const lineIds = async () =>
      evaluate(`[...document.querySelectorAll(".formula-line")].map((line) => line.dataset.lineId)`);

    const activeLineId = async () =>
      evaluate(`document.querySelector(".formula-line.is-active")?.dataset.lineId ?? ""`);

    const undo = async () => click('button[aria-label="撤销"]');
    const redo = async () => click('button[aria-label="重做"]');

    await client.send("Page.navigate", { url: baseUrl });
    await sleep(700);
    await waitForFields(1);

    await resetDocument({ lines: [{ id: "group-line", latex: "" }] });
    assertEqual(
      await evaluate(`document.querySelector('button[aria-label="撤销"]').disabled`),
      true,
      "Undo should start disabled",
    );
    await focusField(0);
    await key("a", "KeyA", 65);
    await key("b", "KeyB", 66);
    await key("c", "KeyC", 67);
    assertDeepEqual(await values(), ["abc"], "typing abc");
    await undo();
    assertDeepEqual(await values(), [""], "grouped typing should undo once");
    assertEqual(await activeLineId(), "group-line", "undo should restore active line");
    assertEqual(
      await evaluate(`document.querySelector("math-field").position`),
      0,
      "undo should restore initial caret",
    );
    await redo();
    assertDeepEqual(await values(), ["abc"], "redo should restore grouped typing");
    assertEqual(
      await evaluate(`document.querySelector("math-field").position`),
      3,
      "redo should restore final caret",
    );
    await focusField(0);
    await key("Backspace", "Backspace", 8);
    await waitForValues(["ab"]);
    assertDeepEqual(await values(), ["ab"], "Backspace should work after redo");
    await undo();
    assertDeepEqual(await values(), ["abc"], "Backspace should be globally undoable");

    await resetDocument({ lines: [{ id: "shortcut-line", latex: "" }] });
    await focusField(0);
    await key("s", "KeyS", 83);
    await waitForValues(["s"]);
    await waitForHistoryAction("撤销");
    await key("z", "KeyZ", 90, 4, false);
    await waitForValues([""]);
    await waitForHistoryAction("重做");
    assertDeepEqual(await values(), [""], "Cmd+Z should use global history");
    await key("z", "KeyZ", 90, 12, false);
    await waitForValues(["s"]);
    await waitForHistoryAction("撤销");
    assertDeepEqual(await values(), ["s"], "Cmd+Shift+Z should redo globally");
    await key("z", "KeyZ", 90, 2, false);
    await waitForValues([""]);
    await waitForHistoryAction("重做");
    assertDeepEqual(await values(), [""], "Ctrl+Z should use global history");
    await key("y", "KeyY", 89, 2, false);
    await waitForValues(["s"]);
    assertDeepEqual(await values(), ["s"], "Ctrl+Y should redo globally");

    await resetDocument({ lines: [{ id: "title-line", latex: "a" }] });
    await evaluate(`(() => {
      const input = document.querySelector(".document-title-area input");
      input.focus();
      input.setSelectionRange(input.value.length, input.value.length);
    })()`);
    await client.send("Input.insertText", { text: " X" });
    await client.send("Input.insertText", { text: "Y" });
    await sleep(120);
    assertEqual(
      await evaluate(`document.querySelector(".document-title-area input").value`),
      "History Test XY",
      "title typing should update the title",
    );
    await key("z", "KeyZ", 90, 4, false);
    await sleep(180);
    assertEqual(
      await evaluate(`document.querySelector(".document-title-area input").value`),
      "History Test",
      "title Cmd+Z should use global grouped history",
    );
    assertEqual(
      await evaluate(`document.activeElement === document.querySelector(".document-title-area input")`),
      true,
      "title undo should preserve title focus",
    );
    await key("z", "KeyZ", 90, 12, false);
    await sleep(180);
    assertEqual(
      await evaluate(`document.querySelector(".document-title-area input").value`),
      "History Test XY",
      "title Cmd+Shift+Z should redo globally",
    );
    assertDeepEqual(await values(), ["a"], "title history must not modify formulas");

    await resetDocument({ lines: [{ id: "timeout-line", latex: "" }] });
    await focusField(0);
    await key("a", "KeyA", 65);
    await sleep(1200);
    await key("b", "KeyB", 66);
    await waitForValues(["ab"]);
    await waitForHistoryAction("撤销");
    await undo();
    await waitForValues(["a"]);
    assertDeepEqual(await values(), ["a"], "typing after timeout should undo separately");
    await undo();
    assertDeepEqual(await values(), [""], "second undo should remove first group");

    await resetDocument({ lines: [{ id: "selection-line", latex: "" }] });
    await focusField(0);
    await key("a", "KeyA", 65);
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.selection = { ranges: [[0, 0]], direction: "none" };
      field.position = 0;
    })()`);
    await sleep(100);
    await key("b", "KeyB", 66);
    assertDeepEqual(await values(), ["ba"], "typing at moved caret");
    await undo();
    assertDeepEqual(await values(), ["a"], "caret jump must split input transaction");
    await undo();
    assertDeepEqual(await values(), [""], "earlier input should remain separately undoable");

    await resetDocument({ lines: [{ id: "delete-line", latex: "abc" }] });
    await focusField(0);
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.selection = { ranges: [[0, 0]], direction: "none" };
      field.position = 0;
    })()`);
    await key("Delete", "Delete", 46);
    await key("Delete", "Delete", 46);
    assertDeepEqual(await values(), ["c"], "continuous Delete should remove two characters");
    await undo();
    assertDeepEqual(await values(), ["abc"], "continuous Delete should undo as one transaction");

    await resetDocument({ lines: [{ id: "composition-line", latex: "" }] });
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.focus();
      field.dispatchEvent(new CompositionEvent("compositionstart", {
        bubbles: true,
        data: "",
      }));
      field.setValue("\\\\text{中}", {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
      field.dispatchEvent(new InputEvent("input", {
        bubbles: true,
        composed: true,
        inputType: "insertCompositionText",
        data: "中",
      }));
      field.setValue("\\\\text{中文}", {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
      field.dispatchEvent(new InputEvent("input", {
        bubbles: true,
        composed: true,
        inputType: "insertCompositionText",
        data: "中文",
      }));
      field.dispatchEvent(new CompositionEvent("compositionend", {
        bubbles: true,
        data: "中文",
      }));
    })()`);
    await sleep(180);
    const compositionValue = (await values())[0];
    assertMatch(compositionValue, /中文/, "composition final value");
    await undo();
    assertDeepEqual(await values(), [""], "composition should undo in one step");
    await redo();
    assertEqual((await values())[0], compositionValue, "composition redo should restore whole word");

    await resetDocument({ lines: [{ id: "line-1", latex: "" }] });
    await focusField(0);
    await key("x", "KeyX", 88);
    await key("Enter", "Enter", 13);
    await waitForFields(2);
    await sleep(180);
    const idsAfterSecondLine = await lineIds();
    assertEqual(
      await activeLineId(),
      idsAfterSecondLine[1],
      "Enter should activate the second line",
    );
    await key("y", "KeyY", 89);
    await key("Enter", "Enter", 13);
    await waitForFields(3);
    await sleep(180);
    const idsAfterAdd = await lineIds();
    assertEqual(
      await activeLineId(),
      idsAfterAdd[2],
      "Enter should activate the third line",
    );
    const thirdLineId = idsAfterAdd[2];
    await key("z", "KeyZ", 90);
    await focusField(1);
    await undo();
    assertDeepEqual(await values(), ["x", "y", ""], "global undo should target third line");
    assertEqual(await activeLineId(), thirdLineId, "global undo should restore third line focus");
    await undo();
    assertDeepEqual(await lineIds(), idsAfterAdd.slice(0, 2), "undo add should remove third line");
    await redo();
    assertDeepEqual(await lineIds(), idsAfterAdd, "redo add should restore the same line id");
    await focusField(2);
    await key("Backspace", "Backspace", 8);
    assertDeepEqual(await lineIds(), idsAfterAdd.slice(0, 2), "empty Backspace should remove line");
    await undo();
    assertDeepEqual(await lineIds(), idsAfterAdd, "undo remove should restore identical line id");
    await redo();
    assertDeepEqual(await lineIds(), idsAfterAdd.slice(0, 2), "redo remove should delete same id");

    await resetDocument({ lines: [{ id: "toolbar-line", latex: "" }] });
    await focusField(0);
    const compactHeight = await evaluate(
      `document.querySelector(".formula-line").getBoundingClientRect().height`,
    );
    await click('button[data-command-id="frac"]');
    const fractionValue = (await values())[0];
    const fractionHeight = await evaluate(
      `document.querySelector(".formula-line").getBoundingClientRect().height`,
    );
    assertMatch(fractionValue, /\\frac/, "fraction toolbar insertion");
    if (!(fractionHeight > compactHeight)) {
      throw new Error(`Fraction row did not grow (${compactHeight} -> ${fractionHeight})`);
    }
    await undo();
    assertDeepEqual(await values(), [""], "fraction should undo in one step");
    const undoHeight = await evaluate(
      `document.querySelector(".formula-line").getBoundingClientRect().height`,
    );
    if (Math.abs(undoHeight - compactHeight) > 2) {
      throw new Error(`Undo did not restore compact row height (${compactHeight} vs ${undoHeight})`);
    }
    await redo();
    assertEqual((await values())[0], fractionValue, "fraction redo should restore structure");
    const redoHeight = await evaluate(
      `document.querySelector(".formula-line").getBoundingClientRect().height`,
    );
    if (Math.abs(redoHeight - fractionHeight) > 2) {
      throw new Error(`Redo did not restore fraction row height (${fractionHeight} vs ${redoHeight})`);
    }

    await resetDocument({ lines: [{ id: "matrix-line", latex: "" }] });
    await focusField(0);
    await click('button[data-category="matrix"]');
    await click('button[data-command-id="custom-matrix"]');
    const matrixValue = (await values())[0];
    assertMatch(matrixValue, /\\begin\{[bpv]matrix\}/, "matrix toolbar insertion");
    await undo();
    assertDeepEqual(await values(), [""], "matrix should undo in one step");
    await redo();
    assertEqual((await values())[0], matrixValue, "matrix redo should restore structure");

    await resetDocument({ lines: [{ id: "restored-caret", latex: "q" }] });
    await focusField(0);
    await key("x", "KeyX", 88);
    await undo();
    assertDeepEqual(await values(), ["q"], "undo should restore formula before toolbar test");
    assertEqual(
      await evaluate(`document.querySelector("math-field").position`),
      1,
      "undo should restore caret before toolbar insertion",
    );
    await click('button[data-command-id="frac"]');
    const afterUndoToolbarValue = (await values())[0];
    if (!afterUndoToolbarValue.startsWith("q") || !/\\frac/.test(afterUndoToolbarValue)) {
      throw new Error(`Toolbar inserted at wrong restored caret: ${afterUndoToolbarValue}`);
    }
    await undo();
    assertDeepEqual(await values(), ["q"], "toolbar insertion after undo should undo cleanly");

    await resetDocument({ lines: [{ id: "selection-wrap", latex: "x" }] });
    await focusField(0);
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.selection = {
        ranges: [[0, field.lastOffset]],
        direction: "forward",
      };
    })()`);
    await click('button[data-command-id="sqrt"]');
    const wrappedValue = (await values())[0];
    assertMatch(wrappedValue, /\\sqrt/, "sqrt should wrap selected content");
    await undo();
    assertDeepEqual(await values(), ["x"], "wrapped selection should undo in one step");

    await resetDocument({ lines: [{ id: "candidate-line", latex: "" }] });
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.focus();
      field.shadowRoot
        ?.querySelector('[part="keyboard-sink"]')
        ?.focus({ preventScroll: true });
      field.setValue("\\\\the", {
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
    await sleep(220);
    const candidateBefore = (await values())[0];
    assertEqual(
      await evaluate(`Boolean(document.querySelector(".suggestion-popup"))`),
      true,
      "candidate popup should open",
    );
    await key("Enter", "Enter", 13);
    await sleep(220);
    const candidateAfter = (await values())[0];
    assertNotEqual(candidateAfter, candidateBefore, "candidate should replace query");
    await undo();
    assertEqual((await values())[0], candidateBefore, "candidate insertion should undo once");
    await redo();
    assertEqual((await values())[0], candidateAfter, "candidate redo should restore result");

    const hasDesktopOcr = await evaluate(
      `Boolean(document.querySelector('button[aria-label="图片公式识别"]'))`,
    );
    if (hasDesktopOcr) {
      await resetDocument({ lines: [{ id: "modal-line", latex: "a" }] });
      await focusField(0);
      await key("b", "KeyB", 66);
      await installFakeTauri();
      await click('button[aria-label="图片公式识别"]');
      await key("z", "KeyZ", 90, 4, false);
      assertDeepEqual(await values(), ["ab"], "OCR modal must block underlying global undo");
      await click('button[aria-label="关闭 OCR"]');
      await undo();
      assertDeepEqual(await values(), ["a"], "global undo should resume after OCR closes");

      await resetDocument({ lines: [{ id: "ocr-line", latex: "a" }] });
      await focusField(0);
      await installFakeTauri();
      await click('button[aria-label="图片公式识别"]');
      await sleep(220);
      await evaluate(`(() => {
        const input = document.querySelector('.ocr-dialog input[type="file"]');
        if (!input) throw new Error("OCR image input was not found");
        const transfer = new DataTransfer();
        transfer.items.add(
          new File([new Uint8Array([137, 80, 78, 71])], "formula.png", {
            type: "image/png",
          }),
        );
        input.files = transfer.files;
        input.dispatchEvent(new Event("change", { bubbles: true }));
      })()`);
      await sleep(160);
      await evaluate(`(() => {
        const button = [...document.querySelectorAll(".ocr-dialog button")].find(
          (item) => item.textContent?.includes("开始识别"),
        );
        if (!button) throw new Error("OCR recognize button was not found");
        button.click();
      })()`);
      await evaluate(`new Promise((resolve, reject) => {
        const started = performance.now();
        const done = () => {
          if (document.querySelector(".ocr-latex-editor textarea")) {
            resolve(true);
            return;
          }
          if (performance.now() - started > 3000) {
            reject(new Error("Fake OCR result did not appear"));
            return;
          }
          setTimeout(done, 30);
        };
        done();
      })`);
      await evaluate(`(() => {
        const button = [...document.querySelectorAll(".ocr-dialog button")].find(
          (item) => item.textContent?.includes("插入当前光标"),
        );
        if (!button) throw new Error("OCR insert button was not found");
        button.click();
      })()`);
      await sleep(260);
      const ocrValue = (await values())[0];
      if (!ocrValue.startsWith("a") || !/\\theta/.test(ocrValue)) {
        throw new Error(`OCR result was not inserted at the saved caret: ${ocrValue}`);
      }
      await undo();
      assertDeepEqual(await values(), ["a"], "OCR insert should undo in one step");
      assertEqual(
        await evaluate(`document.querySelector("math-field").position`),
        1,
        "OCR undo should restore original caret",
      );
      await redo();
      assertEqual((await values())[0], ocrValue, "OCR redo should restore result");
    }

    await evaluate(`new Promise((resolve) => {
      const request = indexedDB.deleteDatabase("visualtex-history");
      request.onsuccess = () => resolve(true);
      request.onerror = () => resolve(false);
      request.onblocked = () => resolve(false);
    })`);
    await resetDocument({
      lines: [{ id: "source-line", latex: "a=b" }],
      sourceOpen: true,
      latexCodeFormat: "raw",
    });
    await evaluate(`new Promise((resolve) => {
      const done = () => document.querySelector(".cm-content") ? resolve(true) : setTimeout(done, 30);
      done();
    })`);
    await evaluate(`document.querySelector(".cm-content").focus()`);
    await client.send("Input.insertText", { text: "c" });
    await sleep(150);
    assertDeepEqual(await values(), ["a=b"], "CodeMirror draft must not update formulas");
    await key("z", "KeyZ", 90, 4, false);
    await sleep(150);
    assertEqual(
      await evaluate(`document.querySelector(".cm-content").innerText`),
      "a=b",
      "CodeMirror should keep its own draft undo",
    );
    assertEqual(
      await evaluate(`document.querySelector('button[aria-label="撤销"]').disabled`),
      true,
      "CodeMirror draft undo must not create global history",
    );

    await evaluate(`(() => {
      const content = document.querySelector(".cm-content");
      content.focus();
      const selection = window.getSelection();
      const range = document.createRange();
      range.selectNodeContents(content);
      selection.removeAllRanges();
      selection.addRange(range);
    })()`);
    await client.send("Input.insertText", { text: "x=y\nz=w" });
    await sleep(180);
    await click(".source-panel .primary-small-button");
    assertDeepEqual(await values(), ["x=y", "z=w"], "source apply should replace document");
    await sleep(300);
    const checkpointCount = await evaluate(`new Promise((resolve, reject) => {
      const request = indexedDB.open("visualtex-history", 1);
      request.onerror = () => reject(request.error);
      request.onsuccess = () => {
        const database = request.result;
        const transaction = database.transaction("checkpoints", "readonly");
        const countRequest = transaction.objectStore("checkpoints").count();
        countRequest.onerror = () => reject(countRequest.error);
        countRequest.onsuccess = () => {
          const count = countRequest.result;
          database.close();
          resolve(count);
        };
      };
    })`);
    if (checkpointCount < 1) {
      throw new Error(`Source apply did not persist an L3 checkpoint (${checkpointCount})`);
    }
    await undo();
    assertDeepEqual(await values(), ["a=b"], "global undo should restore pre-source document");
    assertEqual(
      await evaluate(`document.querySelector(".cm-content").innerText`),
      "a=b",
      "source panel should follow global undo",
    );
    await redo();
    assertDeepEqual(await values(), ["x=y", "z=w"], "redo should reapply source document");

    const historyEntry = {
      id: "saved-history",
      latex: "p=q",
      createdAt: Date.now(),
    };
    await resetDocument({
      lines: [{ id: "history-line", latex: "r=s" }],
      history: [historyEntry],
    });
    await click('button[aria-label="公式历史"]');
    await click(".history-item");
    assertDeepEqual(await values(), ["p=q"], "history item should restore formula");
    await undo();
    assertDeepEqual(await values(), ["r=s"], "history restore should undo as one document operation");

    await resetDocument({ lines: [{ id: "new-before", latex: "u=v" }] });
    await click('button[aria-label="新建"]');
    assertDeepEqual(await values(), [""], "new document should create blank line");
    await undo();
    assertDeepEqual(await values(), ["u=v"], "new document should be undoable");
    await redo();
    assertDeepEqual(await values(), [""], "new document should be redoable");
    await focusField(0);
    await key("n", "KeyN", 78);
    assertEqual(
      await evaluate(`document.querySelector('button[aria-label="重做"]').disabled`),
      true,
      "new edit after undo/redo branch should clear redo",
    );

    await resetDocument({ lines: [{ id: "before-open", latex: "c=d" }] });
    await evaluate(`(() => {
      const input = document.querySelector('input[type="file"][accept=".json,.visualtex"]');
      if (!input) throw new Error("Document file input was not found");
      const documentValue = {
        version: 3,
        title: "Opened document",
        formulas: [
          { id: "opened-line-1", latex: "m=n" },
          { latex: "k=l" },
        ],
        macros: {},
        settings: {
          theme: "light",
          zoom: 1,
          latexCodeFormat: "raw",
        },
      };
      const transfer = new DataTransfer();
      transfer.items.add(
        new File(
          [JSON.stringify(documentValue)],
          "history-open.visualtex",
          { type: "application/json" },
        ),
      );
      input.files = transfer.files;
      input.dispatchEvent(new Event("change", { bubbles: true }));
    })()`);
    await sleep(450);
    assertDeepEqual(await values(), ["m=n", "k=l"], "opening a document should replace formulas");
    const openedIds = await lineIds();
    assertEqual(openedIds[0], "opened-line-1", "opening should preserve an existing line id");
    if (!openedIds[1] || openedIds[1] === openedIds[0]) {
      throw new Error(`Opening should generate a unique missing line id: ${JSON.stringify(openedIds)}`);
    }
    await undo();
    assertDeepEqual(await values(), ["c=d"], "open document should be undoable");
    assertDeepEqual(await lineIds(), ["before-open"], "undo open should restore old line id");
    await redo();
    assertDeepEqual(await values(), ["m=n", "k=l"], "open document should be redoable");
    assertDeepEqual(await lineIds(), openedIds, "redo open should reuse generated stable ids");

    await resetDocument({
      lines: [
        { id: "export-line-1", latex: "e=f" },
        { id: "export-line-2", latex: "g=h" },
      ],
      activeLineId: "export-line-2",
    });
    await evaluate(`(() => {
      window.__visualtexSavedDocument = null;
      const originalCreateObjectURL = URL.createObjectURL.bind(URL);
      URL.createObjectURL = (blob) => {
        if (blob.type === "application/json") {
          void blob.text().then((text) => {
            window.__visualtexSavedDocument = text;
          });
        }
        return originalCreateObjectURL(blob);
      };
      HTMLAnchorElement.prototype.click = function () {};
    })()`);
    await click('button[aria-label="保存到本地"]');
    await evaluate(`new Promise((resolve, reject) => {
      const started = performance.now();
      const done = () => {
        if (window.__visualtexSavedDocument) {
          resolve(true);
          return;
        }
        if (performance.now() - started > 2000) {
          reject(new Error("Saved document was not captured"));
          return;
        }
        setTimeout(done, 20);
      };
      done();
    })`);
    const exportedIds = await evaluate(`JSON.parse(window.__visualtexSavedDocument).formulas.map((formula) => formula.id)`);
    assertDeepEqual(
      exportedIds,
      ["export-line-1", "export-line-2"],
      "saved document should preserve stable line ids",
    );

    await evaluate(`(() => {
      const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      persisted.state = {
        ...(persisted.state || {}),
        title: "Legacy document",
        latex: "a=b\\nc=d",
      };
      delete persisted.state.lines;
      delete persisted.state.activeLineId;
      localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(850);
    await waitForFields(2);
    assertDeepEqual(await values(), ["a=b", "c=d"], "legacy latex should migrate to stable lines");
    const migratedIds = await lineIds();
    if (
      migratedIds.length !== 2 ||
      !migratedIds[0] ||
      !migratedIds[1] ||
      migratedIds[0] === migratedIds[1]
    ) {
      throw new Error(`Legacy migration did not create unique line ids: ${JSON.stringify(migratedIds)}`);
    }

    console.log("Editor global history regression test passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(300);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => undefined);
  }
}

function assertEqual(actual, expected, label) {
  if (actual !== expected) {
    throw new Error(`${label}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
  }
}

function assertNotEqual(actual, expected, label) {
  if (actual === expected) {
    throw new Error(`${label}: values should differ (${JSON.stringify(actual)})`);
  }
}

function assertDeepEqual(actual, expected, label) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`${label}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
  }
}

function assertMatch(actual, pattern, label) {
  if (!pattern.test(actual)) {
    throw new Error(`${label}: ${JSON.stringify(actual)} does not match ${pattern}`);
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.stack : error);
  process.exitCode = 1;
});
