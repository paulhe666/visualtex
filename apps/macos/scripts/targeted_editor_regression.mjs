import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const scenario = process.argv[2];
if (!new Set(["wrapper", "wrapper-auto", "wrapper-continuous", "wrapper-prefix", "native-input-popover", "native-space-selection", "candidate-query-reset", "raw-placeholder-visual", "placeholder-selection", "structural-placeholder", "caret-probe", "scripts", "upright", "suggestions", "navigation", "geometry", "settings", "layout", "delete", "export"]).has(scenario)) {
  throw new Error(
    "Usage: node scripts/targeted_editor_regression.mjs <wrapper|wrapper-auto|wrapper-continuous|wrapper-prefix|native-input-popover|native-space-selection|candidate-query-reset|raw-placeholder-visual|placeholder-selection|structural-placeholder|caret-probe|scripts|upright|suggestions|navigation|geometry|settings|layout|delete|export>",
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

    if (scenario === "structural-placeholder") {
      const placeholderCases = [
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
          return {
            ready:
              Boolean(styleNode) &&
              placeholders.length === ${testCase.expectedCount} &&
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
            caretStyle?.left === "0px" &&
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
            caretStyle?.left === "0px" &&
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
      console.log(JSON.stringify({ initial, arrowState, refinedState, restoredState }, null, 2));
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
