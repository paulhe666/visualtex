import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const scenario = process.argv[2];
if (!new Set(["wrapper", "scripts", "upright", "suggestions", "settings", "layout", "delete", "export"]).has(scenario)) {
  throw new Error(
    "Usage: node scripts/targeted_editor_regression.mjs <wrapper|scripts|upright|suggestions|settings|layout|delete|export>",
  );
}

const offset = process.pid % 1000;
const previewPort = 6400 + offset;
const debugPort = 11600 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = `/tmp/visualtex-targeted-${scenario}-${process.pid}`;
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

    await client.send("Page.navigate", { url: baseUrl });
    await sleep(650);
    await evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
      const storageKey = "visualtex-editor";
      const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
      persisted.state = {
        ...(persisted.state || {}),
        lines: [{ id: crypto.randomUUID(), latex: "" }],
        activeLineId: null,
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
      await sleep(100);
      await focusField();
    };

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

    if (scenario === "wrapper") {
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
        const placeholderLeft = Number.parseFloat(
          host?.style.getPropertyValue("--pending-wrapper-left") ?? "NaN",
        );
        const placeholderTop = Number.parseFloat(
          host?.style.getPropertyValue("--pending-wrapper-top") ?? "NaN",
        );
        const expectedLeft =
          nativeCaretBounds && hostBounds
            ? nativeCaretBounds.left - hostBounds.left
            : Number.NaN;
        const expectedTop =
          nativeCaretBounds && hostBounds
            ? nativeCaretBounds.top - hostBounds.top + nativeCaretBounds.height / 2
            : Number.NaN;
        return {
          ready:
            field.value === "abcdefghij\\\\mathbb{}" &&
            field.dataset.pendingWrapperCommand === "\\\\mathbb" &&
            host?.classList.contains("has-pending-wrapper-placeholder") &&
            placeholderStyle?.borderStyle === "solid" &&
            Number.parseFloat(placeholderStyle?.borderWidth ?? "0") <= 1.1 &&
            Boolean(nativeCaret) &&
            fakeCaretStyle?.content === "none" &&
            nativeCaretStyle?.visibility === "visible" &&
            nativeCaretStyle?.animationName.includes("caret-blink") &&
            Math.abs(placeholderLeft - expectedLeft) <= 1.5 &&
            Math.abs(placeholderTop - expectedTop) <= 1.5 &&
            Math.abs(placeholderLeft - (hostBounds?.width ?? 0) / 2) >= 20 &&
            document.querySelectorAll("math-field").length === 1,
          value: field.value,
          pendingWrapperCommand: field.dataset.pendingWrapperCommand ?? "",
          placeholderClass: host?.classList.contains("has-pending-wrapper-placeholder") ?? false,
          placeholderBorderStyle: placeholderStyle?.borderStyle ?? "",
          placeholderBorderWidth: placeholderStyle?.borderWidth ?? "",
          placeholderLeft,
          placeholderTop,
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

      console.log(JSON.stringify({ identifierState, differentialState }, null, 2));
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
      }, null, 2));
      console.log("Targeted other-command suggestion dismissal and navigation regression passed");
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
