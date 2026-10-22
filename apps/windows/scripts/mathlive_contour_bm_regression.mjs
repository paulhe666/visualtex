import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const offset = process.pid % 1000;
const previewPort = 7600 + offset;
const debugPort = 17600 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const browserProfile = createBrowserProfilePath(
  "visualtex-mathlive-contour-bm-regression",
);
const browserPath = resolveChromiumExecutable();
const sleep = (milliseconds) =>
  new Promise((resolve) => setTimeout(resolve, milliseconds));

async function waitFor(url, timeoutMs = 15_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while the local preview/browser starts.
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
  let browser;
  let client;

  try {
    await waitFor(baseUrl);
    browser = spawn(
      browserPath,
      [
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        `--remote-debugging-port=${debugPort}`,
        `--user-data-dir=${browserProfile}`,
        "--window-size=1400,1000",
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
    if (!page) throw new Error(`No browser page target: ${JSON.stringify(targets)}`);

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

    const waitForEvaluation = async (
      expression,
      description,
      timeoutMs = 12_000,
    ) => {
      const started = Date.now();
      let lastValue;
      while (Date.now() - started < timeoutMs) {
        lastValue = await evaluate(expression);
        if (lastValue?.ready) return lastValue;
        await sleep(50);
      }
      throw new Error(
        `Timed out waiting for ${description}: ${JSON.stringify({
          lastValue,
          events: client.events,
        })}`,
      );
    };

    const dispatchKey = async (key, code, virtualKeyCode) => {
      const common = {
        key,
        code,
        windowsVirtualKeyCode: virtualKeyCode,
        nativeVirtualKeyCode: virtualKeyCode,
      };
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        ...common,
        ...(key.length === 1 ? { text: key, unmodifiedText: key } : {}),
      });
      await client.send("Input.dispatchKeyEvent", {
        type: "keyUp",
        ...common,
      });
      await sleep(60);
    };

    const typeText = async (text) => {
      for (const character of text) {
        const code =
          character === "\\" ? "Backslash" : `Key${character.toUpperCase()}`;
        const virtualKeyCode =
          character === "\\"
            ? 220
            : character.toUpperCase().charCodeAt(0);
        await dispatchKey(character, code, virtualKeyCode);
      }
    };

    const inspectNativeSuggestionQuery = async (query, expectedCommands) => {
      await evaluate(`(async () => {
        const field = document.querySelector("math-field");
        if (!field?.shadowRoot) return false;
        field.setValue("", {
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
        await new Promise((resolve) =>
          requestAnimationFrame(() => requestAnimationFrame(resolve)),
        );
        field.focus();
        field.position = field.lastOffset;
        field.shadowRoot
          .querySelector('[part="keyboard-sink"]')
          ?.focus({ preventScroll: true });
        return field.hasFocus();
      })()`);
      await sleep(100);
      await typeText(query);
      return waitForEvaluation(
        `(() => {
          const expected = ${JSON.stringify(expectedCommands)};
          const panel = document.getElementById("mathlive-suggestion-popover");
          const items = [...(panel?.querySelectorAll("li[data-command]") ?? [])];
          const byCommand = new Map(
            items.map((item) => [item.dataset.command ?? "", item]),
          );
          const entries = expected.map((command) => {
            const item = byCommand.get(command);
            const preview = item?.querySelector(".ML__popover__command");
            const rendered = preview?.querySelector(".ML__latex") ?? preview;
            const bounds = rendered?.getBoundingClientRect();
            return {
              command: item?.dataset.command ?? "",
              previewLatex: preview?.dataset.visualtexPreview ?? "",
              kind: preview?.dataset.visualtexPreviewKind ?? "native",
              text: rendered?.textContent?.trim() ?? "",
              width: bounds?.width ?? 0,
              height: bounds?.height ?? 0,
              error: Boolean(rendered?.querySelector(".ML__error")),
              linkCount: rendered?.querySelectorAll("a[href]").length ?? 0,
            };
          });
          return {
            ready: Boolean(panel?.classList.contains("is-visible")),
            entries: entries.filter((entry) => entry.command),
          };
        })()`,
        `native suggestion previews for ${query}`,
      );
    };

    await client.send("Page.navigate", { url: baseUrl });
    await sleep(1800);
    await evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem(
        "visualtex.onboarding.windows.desktop.v1.1.0.completed",
        "true",
      );
      const storageKey = "visualtex-editor";
      const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
      const line = { id: crypto.randomUUID(), latex: "" };
      persisted.state = {
        ...(persisted.state || {}),
        lines: [line],
        activeLineId: line.id,
        inputBehavior: {
          autoEscapeShortcuts: true,
          autoExitSuperscript: true,
          autoExitSubscript: true,
          autoExitAccent: true,
          autoExitWrapperCommand: true,
          showStructuredCommandSuggestions: false,
          showOtherCommandSuggestions: false,
        },
      };
      localStorage.setItem(storageKey, JSON.stringify(persisted));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });

    await waitForEvaluation(
      `(() => ({ ready: Boolean(document.querySelector("math-field")?.shadowRoot) }))()`,
      "VisualTeX mathfield",
    );

    const contourLatex = String.raw`\displaystyle\oiint_{a}^{b}+\oiiint_{c}^{d}`;
    const contourState = await waitForEvaluation(
      `(() => {
        const field = document.querySelector("math-field");
        if (!field?.shadowRoot) return { ready: false };
        field.setValue(${JSON.stringify(contourLatex)}, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        const root = field.shadowRoot;
        // The layout engine sets \\oiint in esint10 (as \\usepackage{esint})
        // and draws \\oiiint as amsmath's \\iiint with a contour oval.
        const glyphs = [...root.querySelectorAll(".ML__vtx > span")].filter((node) => node.style.font);
        const esint = glyphs.find((node) => /VTX Esint/.test(node.style.fontFamily));
        const oval = root.querySelector(".ML__vtx-ink ellipse");
        const ovalRect = oval?.getBoundingClientRect();
        return {
          ready: Boolean(esint && oval),
          value: field.value,
          globalStyle: Boolean(document.getElementById("visualtex-mathlive-contour-integral-style")),
          shadowStyle: Boolean(root.getElementById("visualtex-mathlive-contour-integral-shadow-style")),
          oiint: esint ? { font: esint.style.fontFamily } : null,
          oiiint: oval ? { width: ovalRect.width, height: ovalRect.height } : null,
          integralSigns: glyphs.filter((node) => /VTX Math Extension/.test(node.style.fontFamily)).length,
        };
      })()`,
      "contour-integral MathLive rendering",
    );

    assert.equal(contourState.globalStyle, true, "global contour style installed");
    assert.equal(contourState.shadowStyle, true, "shadow contour style installed");
    assert.match(contourState.value, /\\oiint/);
    assert.match(contourState.value, /\\oiiint/);
    assert.match(contourState.oiint.font, /VTX Esint/, "oiint is the esint10 glyph");
    assert.equal(contourState.integralSigns, 3, "oiiint draws amsmath's three integral signs");
    assert.ok(contourState.oiiint.width > 10, "oiiint oval has visible width");
    assert.ok(contourState.oiiint.height > 3, "oiiint oval has visible height");

    const rareIntegralCommands = [
      "intclockwise",
      "varointclockwise",
      "ointctrclockwise",
      "sumint",
      "iiiint",
      "intbar",
      "intBar",
      "fint",
      "cirfnint",
      "awint",
      "intctrclockwise",
      "rppolint",
      "scpolint",
      "npolint",
      "pointint",
      "quatint",
      "intlarhk",
      "intx",
      "intcap",
      "intcup",
      "upint",
      "lowint",
    ];
    const rareIntegralStates = [];
    for (const command of rareIntegralCommands) {
      const latex = `\\displaystyle\\${command}_{a}^{b}`;
      const state = await waitForEvaluation(
        `(async () => {
          const field = document.querySelector("math-field");
          if (!field?.shadowRoot) return { ready: false };
          field.setValue(${JSON.stringify(latex)}, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceAll",
            selectionMode: "after",
            silenceNotifications: true,
          });
          await new Promise((resolve) =>
            requestAnimationFrame(() => requestAnimationFrame(resolve)),
          );
          // drawn as an outline (STIX Two Math), or as a glyph of a TeX font
          // when LaTeX defines the command (amsmath \\iiiint, esint \\fint ...)
          const root = field.shadowRoot;
          const node = root.querySelector(".ML__op-group");
          const paths = [...root.querySelectorAll(".ML__vtx-ink path")];
          const glyph = [...root.querySelectorAll(".ML__vtx > span")].find((span) =>
            /VTX (Esint|Math Extension)/.test(span.style.fontFamily ?? ""),
          );
          const ink = paths[0] ?? null;
          const rect = (ink ?? node)?.getBoundingClientRect();
          return {
            ready: Boolean(node && (ink || glyph)),
            value: field.value,
            classes: node?.className ?? "",
            width: rect?.width ?? 0,
            height: rect?.height ?? 0,
            path: ink?.getAttribute("d") ?? "",
            pathCount: paths.length,
            glyph: Boolean(glyph),
            hasError: Boolean(root.querySelector(".ML__error")),
          };
        })()`,
        `${command} rare-integral rendering`,
      );
      assert.match(state.value, new RegExp(`\\\\${command}`));
      assert.match(state.classes, /ML__op-group/);
      if (state.glyph) {
        assert.equal(state.pathCount, 0, `${command} TeX glyph, no outline`);
      } else {
        assert.equal(state.pathCount, 1, `${command} vector path count`);
        assert.ok(state.path.length > 20, `${command} vector path data`);
      }
      assert.ok(state.width > 8, `${command} visible width`);
      assert.ok(state.height > 40, `${command} display glyph height`);
      assert.equal(state.hasError, false, `${command} has no error atom`);
      rareIntegralStates.push({ command, ...state });
    }

    for (const command of ["intclockwise", "iiiint", "awint", "lowint"]) {
      const latex = `x^{\\${command}_{a}^{b}}`;
      const state = await waitForEvaluation(
        `(async () => {
          const field = document.querySelector("math-field");
          if (!field?.shadowRoot) return { ready: false };
          field.setValue(${JSON.stringify(latex)}, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceAll",
            selectionMode: "after",
            silenceNotifications: true,
          });
          await new Promise((resolve) =>
            requestAnimationFrame(() => requestAnimationFrame(resolve)),
          );
          const root = field.shadowRoot;
          const node = root.querySelector(".ML__op-group");
          const path = root.querySelector(".ML__vtx-ink path");
          const glyph = [...root.querySelectorAll(".ML__vtx > span")].find((span) =>
            /VTX (Esint|Math Extension)/.test(span.style.fontFamily ?? ""),
          );
          const rect = (path ?? node)?.getBoundingClientRect();
          return {
            ready: Boolean(node && (path || glyph)),
            classes: node?.className ?? "",
            width: rect?.width ?? 0,
            height: rect?.height ?? 0,
            path: path?.getAttribute("d") ?? (glyph ? "tex-font-glyph-in-a-tex-font" : ""),
          };
        })()`,
        `${command} small rare-integral rendering`,
      );
      assert.match(state.classes, /ML__op-group/);
      assert.ok(state.path.length > 20, `${command} small vector path data`);
      assert.ok(state.width > 5, `${command} small visible width`);
      assert.ok(state.height > 18, `${command} small glyph height`);
    }

    const limitsState = await waitForEvaluation(
      `(async () => {
        const field = document.querySelector("math-field");
        if (!field?.shadowRoot) return { ready: false };
        field.setValue("\\\\iiiint\\\\limits_{a}^{b}", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        await new Promise((resolve) =>
          requestAnimationFrame(() => requestAnimationFrame(resolve)),
        );
        const operator = field.shadowRoot.querySelector(".ML__op-group");
        const operatorRect = operator?.getBoundingClientRect();
        const scriptBounds = [...field.shadowRoot.querySelectorAll(".ML__mathit")]
          .filter((node) => node.textContent === "a" || node.textContent === "b")
          .map((node) => {
            const bounds = node.getBoundingClientRect();
            return {
              text: node.textContent,
              left: bounds.left,
              top: bounds.top,
              right: bounds.right,
              bottom: bounds.bottom,
            };
          });
        return {
          ready: Boolean(operator),
          value: field.value,
          operatorTop: operatorRect?.top ?? 0,
          operatorBottom: operatorRect?.bottom ?? 0,
          operatorLeft: operatorRect?.left ?? 0,
          operatorRight: operatorRect?.right ?? 0,
          hasAdjacentScripts: Boolean(field.shadowRoot.querySelector(".ML__msubsup")),
          msubsupCount: field.shadowRoot.querySelectorAll(".ML__msubsup").length,
          parentClasses: [operator, operator?.parentElement, operator?.parentElement?.parentElement]
            .filter(Boolean)
            .map((node) => node.className),
          scriptBounds,
        };
      })()`,
      "rare-integral explicit limits",
    );
    assert.match(limitsState.value, /\\iiiint\\limits/);
    assert.equal(
      limitsState.hasAdjacentScripts,
      false,
      "rare integral explicit limits use over-under scripts",
    );

    const noLimitsState = await waitForEvaluation(
      `(async () => {
        const field = document.querySelector("math-field");
        if (!field?.shadowRoot) return { ready: false };
        field.setValue("\\\\iiiint\\\\nolimits_{a}^{b}", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        await new Promise((resolve) =>
          requestAnimationFrame(() => requestAnimationFrame(resolve)),
        );
        const operator = field.shadowRoot.querySelector(".ML__op-group");
        return {
          ready: Boolean(operator),
          value: field.value,
          hasAdjacentScripts: Boolean(field.shadowRoot.querySelector(".ML__msubsup")),
        };
      })()`,
      "rare-integral explicit nolimits",
    );
    assert.match(noLimitsState.value, /\\iiiint\\nolimits/);
    assert.equal(
      noLimitsState.hasAdjacentScripts,
      true,
      "rare integral explicit nolimits use adjacent scripts",
    );

    await waitForEvaluation(
      `(() => {
        const field = document.querySelector("math-field");
        if (!field?.shadowRoot) return { ready: false };
        field.setValue("", {
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
        return { ready: field.value === "" };
      })()`,
      "empty field for bm suggestion",
    );
    await sleep(120);
    await waitForEvaluation(
      `(() => {
        const field = document.querySelector("math-field");
        if (!field?.shadowRoot) return { ready: false };
        field.focus();
        field.position = field.lastOffset;
        field.shadowRoot
          .querySelector('[part="keyboard-sink"]')
          ?.focus({ preventScroll: true });
        return {
          ready: field.isConnected && field.hasFocus(),
          value: field.value,
          activeTag: document.activeElement?.tagName ?? "",
          shadowActivePart:
            field.shadowRoot.activeElement?.getAttribute("part") ?? "",
        };
      })()`,
      "refocused current field for bm suggestion",
    );
    await typeText(String.raw`\b`);

    const bmState = await waitForEvaluation(
      `(() => {
        const panel = document.getElementById("mathlive-suggestion-popover");
        const item = [...(panel?.querySelectorAll("li[data-command]") ?? [])]
          .find((candidate) => candidate.dataset.command === "\\\\bm");
        const preview = item?.querySelector(".ML__popover__command");
        return {
          ready: Boolean(preview?.dataset.visualtexPreview),
          command: item?.dataset.command ?? "",
          previewLatex: preview?.dataset.visualtexPreview ?? "",
          text: preview?.textContent ?? "",
          hasError: Boolean(preview?.querySelector(".ML__error")),
          html: preview?.innerHTML ?? "",
          panelVisible: panel?.classList.contains("is-visible") ?? false,
          commands: [...(panel?.querySelectorAll("li[data-command]") ?? [])]
            .map((candidate) => candidate.dataset.command ?? ""),
          rawLatex: [...(document.querySelector("math-field")?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .map((node) => node.textContent ?? "")
            .join(""),
        };
      })()`,
      "bm native suggestion preview",
    );

    assert.equal(bmState.command, String.raw`\bm`);
    assert.equal(bmState.previewLatex, String.raw`\bm{\alpha A}`);
    assert.equal(bmState.hasError, false, "bm preview has no MathLive error atom");
    assert.ok(bmState.html.length > 20, "bm preview contains rendered markup");
    assert.doesNotMatch(bmState.text, /\\bm/, "bm preview does not show raw command text");

    const uncommonPreviewCases = [
      {
        query: String.raw`\b`,
        expected: {
          "\\bm": "alias",
          "\\bold": "alias",
          "\\biggl": "delimiter",
          "\\bmod": "native",
        },
      },
      {
        query: String.raw`\c`,
        expected: {
          "\\c": "arguments",
          "\\cancel": "arguments",
          "\\ce": "arguments",
          "\\class": "arguments",
          "\\color": "arguments",
        },
      },
      {
        query: String.raw`\math`,
        expected: {
          "\\mathbfit": "state",
          "\\mathbin": "arguments",
          "\\mathchoice": "arguments",
          "\\mathrel": "arguments",
        },
      },
      {
        query: String.raw`\s`,
        expected: {
          "\\scriptstyle": "state",
          "\\sffamily": "state",
          "\\small": "state",
          "\\smash": "arguments",
          "\\space": "spacing",
          "\\strut": "state",
        },
      },
      {
        query: String.raw`\the`,
        expected: {
          "\\the": "fallback",
        },
      },
    ];
    const uncommonPreviewResults = {};
    for (const testCase of uncommonPreviewCases) {
      const result = await inspectNativeSuggestionQuery(
        testCase.query,
        Object.keys(testCase.expected),
      );
      uncommonPreviewResults[testCase.query] = result.entries;
      for (const entry of result.entries) {
        assert.equal(
          entry.kind,
          testCase.expected[entry.command],
          `${entry.command} preview kind`,
        );
        assert.equal(entry.error, false, `${entry.command} preview parse error`);
        assert.equal(entry.linkCount, 0, `${entry.command} preview created a link`);
        assert.ok(
          entry.text.length > 0 || entry.width > 3,
          `${entry.command} preview has no visible content`,
        );
        if (entry.kind === "native") {
          assert.equal(
            entry.previewLatex,
            "",
            `${entry.command} native preview was replaced`,
          );
        } else {
          assert.notEqual(
            entry.previewLatex,
            "",
            `${entry.command} preview was not decorated`,
          );
        }
      }
    }

    console.log(
      JSON.stringify(
        {
          contour: contourState,
          rareIntegralCount: rareIntegralStates.length,
          rareIntegralCommands: rareIntegralStates.map((state) => state.command),
          limits: {
            explicit: limitsState.hasAdjacentScripts,
            explicitNoLimits: noLimitsState.hasAdjacentScripts,
          },
          bm: bmState,
          uncommonPreviews: uncommonPreviewResults,
        },
        null,
        2,
      ),
    );
    console.log("VisualTeX MathLive contour integral and bm preview regression passed");
  } finally {
    client?.close();
    browser?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(250);
    await rm(browserProfile, { recursive: true, force: true }).catch(
      () => undefined,
    );
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.stack : error);
  process.exitCode = 1;
});
