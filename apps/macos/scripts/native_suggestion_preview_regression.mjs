import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const fullAudit = process.argv.includes("--audit");
const sqIntegralAudit = process.argv.includes("--sq-integrals");
const offset = process.pid % 1000;
const previewPort = 18400 + offset;
const debugPort = 23400 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = `/tmp/visualtex-native-suggestion-preview-${process.pid}`;
const chromePath =
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (milliseconds) =>
  new Promise((resolve) => setTimeout(resolve, milliseconds));

async function waitFor(url, timeoutMs = 15_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while the local preview or browser starts.
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
    const targets = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json();
    const page =
      targets.find(
        (target) => target.type === "page" && target.url.startsWith(baseUrl),
      ) ?? targets.find((target) => target.type === "page");
    if (!page) throw new Error("No Chrome page target found.");

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
          result.exceptionDetails.exception?.description ??
            result.exceptionDetails.text ??
            "Runtime.evaluate failed",
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
        text: value,
        unmodifiedText: value,
      });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(40);
    };

    await client.send("Page.navigate", { url: baseUrl });
    await sleep(500);
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
      };
      persisted.state.activeLineId = persisted.state.lines[0].id;
      delete persisted.state.inputBehavior;
      localStorage.setItem(storageKey, JSON.stringify(persisted));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });

    const started = Date.now();
    let fieldReady = false;
    while (Date.now() - started < 12_000) {
      fieldReady = await evaluate(
        `Boolean(document.querySelector("math-field.visual-mathfield")?.shadowRoot)`,
      );
      if (fieldReady) {
        break;
      }
      await sleep(50);
    }
    assert.equal(fieldReady, true, "VisualTeX mathfield mounted");

    const clearAndFocus = async () => {
      const clearStarted = Date.now();
      let cleared = false;
      while (!cleared && Date.now() - clearStarted < 5_000) {
        cleared = await evaluate(`(() => {
          const field = document.querySelector("math-field.visual-mathfield");
          if (!field?.isConnected || !field.shadowRoot) return false;
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
          return field.isConnected && field.value === "";
        })()`);
        if (!cleared) await sleep(40);
      }
      assert.equal(cleared, true, "stable empty VisualTeX mathfield");
      await sleep(100);

      const focusStarted = Date.now();
      let focused = false;
      while (!focused && Date.now() - focusStarted < 5_000) {
        focused = await evaluate(`(() => {
          const field = document.querySelector("math-field.visual-mathfield");
          if (!field?.isConnected || !field.shadowRoot) return false;
          field.focus();
          field.position = field.lastOffset;
          field.shadowRoot
            ?.querySelector('[part="keyboard-sink"]')
            ?.focus({ preventScroll: true });
          return field.isConnected && field.hasFocus();
        })()`);
        if (!focused) await sleep(40);
      }
      assert.equal(focused, true, "stable focused VisualTeX mathfield");
      await sleep(80);
    };

    const typeText = async (text) => {
      for (const character of text) {
        const code =
          character === "\\" ? "Backslash" : `Key${character.toUpperCase()}`;
        const virtualKeyCode =
          character === "\\"
            ? 220
            : character.toUpperCase().charCodeAt(0);
        await key(character, code, virtualKeyCode);
      }
    };

    const inspectQuery = async (query, expectedCommands) => {
      await clearAndFocus();
      await typeText(query);
      const started = Date.now();
      let state;
      while (Date.now() - started < 5_000) {
        state = await evaluate(`(() => {
          const expected = ${JSON.stringify(expectedCommands)};
          const stablePanel = document.getElementById(
            "visualtex-native-input-suggestion-popover",
          );
          const sourcePanel = document.getElementById("mathlive-suggestion-popover");
          const panel = stablePanel?.querySelector("li[data-command]")
            ? stablePanel
            : sourcePanel;
          const items = [...(panel?.querySelectorAll("li[data-command]") ?? [])];
          const byCommand = new Map(
            items.map((item) => [item.dataset.command ?? "", item]),
          );
          const result = expected.map((command) => {
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
          const field = document.querySelector("math-field.visual-mathfield");
          const rawLatex = [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join("");
          return {
            ready:
              Boolean(panel?.classList.contains("is-visible")) &&
              result.every((entry) => entry.command),
            result,
            availableCommands: items.map((item) => item.dataset.command ?? ""),
            fieldValue: field?.value ?? "",
            fieldMode: field?.mode ?? "",
            fieldFocused: field?.hasFocus() ?? false,
            rawLatex,
          };
        })()`);
        if (state?.ready) break;
        await sleep(50);
      }
      assert.equal(
        state?.ready,
        true,
        `native candidates for ${query}: ${JSON.stringify(state)}`,
      );
      await sleep(160);
      return (
        await evaluate(`(() => {
          const expected = ${JSON.stringify(expectedCommands)};
          const stablePanel = document.getElementById(
            "visualtex-native-input-suggestion-popover",
          );
          const sourcePanel = document.getElementById("mathlive-suggestion-popover");
          const panel = stablePanel?.querySelector("li[data-command]")
            ? stablePanel
            : sourcePanel;
          return expected.map((command) => {
            const item = [...(panel?.querySelectorAll("li[data-command]") ?? [])]
              .find((candidate) => candidate.dataset.command === command);
            const preview = item?.querySelector(".ML__popover__command");
            const rendered = preview?.querySelector(".ML__latex") ?? preview;
            const bounds = rendered?.getBoundingClientRect();
            const operator = rendered?.querySelector(".ML__op-symbol");
            const svg = operator?.querySelector("svg");
            const path = svg?.querySelector("path");
            const operatorBounds = operator?.getBoundingClientRect();
            const svgBounds = svg?.getBoundingClientRect();
            const itemBounds = item?.getBoundingClientRect();
            const keybinding = item?.querySelector(".ML__popover__keybinding");
            const keybindingBounds = keybinding?.getBoundingClientRect();
            const keybindingStyle = keybinding ? getComputedStyle(keybinding) : null;
            return {
              command: item?.dataset.command ?? "",
              previewLatex: preview?.dataset.visualtexPreview ?? "",
              kind: preview?.dataset.visualtexPreviewKind ?? "native",
              text: rendered?.textContent?.trim() ?? "",
              width: bounds?.width ?? 0,
              height: bounds?.height ?? 0,
              error: Boolean(rendered?.querySelector(".ML__error")),
              linkCount: rendered?.querySelectorAll("a[href]").length ?? 0,
              keybindingText: item?.querySelector(".ML__popover__keybinding")?.textContent?.trim() ?? "",
              keybindingHtml: keybinding?.innerHTML ?? "",
              keybindingWidth: keybindingBounds?.width ?? 0,
              keybindingHeight: keybindingBounds?.height ?? 0,
              keybindingWhiteSpace: keybindingStyle?.whiteSpace ?? "",
              keybindingFlexBasis: keybindingStyle?.flexBasis ?? "",
              itemHeight: itemBounds?.height ?? 0,
              operatorClass: operator?.className ?? "",
              operatorWidth: operatorBounds?.width ?? 0,
              operatorHeight: operatorBounds?.height ?? 0,
              svgCount: operator?.querySelectorAll("svg").length ?? 0,
              svgWidth: svgBounds?.width ?? 0,
              svgHeight: svgBounds?.height ?? 0,
              svgViewBox: svg?.getAttribute("viewBox") ?? "",
              pathFillRule: path ? getComputedStyle(path).fillRule : "",
            };
          });
        })()`)
      );
    };

    const allCases = [
      {
        query: "\\b",
        expected: {
          "\\bold": "alias",
          "\\biggl": "delimiter",
          "\\biggm": "delimiter",
          "\\biggr": "delimiter",
          "\\bm": "alias",
        },
      },
      {
        query: "\\c",
        expected: {
          "\\c": "arguments",
          "\\cancel": "arguments",
          "\\ce": "arguments",
          "\\class": "arguments",
          "\\color": "arguments",
        },
      },
      {
        query: "\\math",
        expected: {
          "\\mathbfit": "state",
          "\\mathbin": "arguments",
          "\\mathchoice": "arguments",
          "\\mathrel": "arguments",
        },
      },
      {
        query: "\\s",
        expected: {
          "\\scriptstyle": "state",
          "\\sffamily": "state",
          "\\small": "state",
          "\\smash": "arguments",
          "\\space": "spacing",
        },
      },
      {
        query: "\\sq",
        expected: {
          "\\sqiint": "operator",
          "\\sqrt": "arguments",
          "\\sqcap": "native",
          "\\sqcup": "native",
          "\\sqint": "operator",
        },
      },
      {
        query: "\\q",
        expected: {
          "\\quad": "spacing",
          "\\qquad": "spacing",
        },
      },
      {
        query: "\\h",
        expected: {
          "\\href": "arguments",
          "\\hphantom": "state",
          "\\hspace": "spacing",
        },
      },
      {
        query: "\\the",
        expected: {
          "\\the": "fallback",
        },
      },
    ];
    const cases = sqIntegralAudit
      ? allCases.filter((testCase) => testCase.query === "\\sq")
      : allCases;

    const results = {};
    for (const testCase of cases) {
      const entries = await inspectQuery(
        testCase.query,
        Object.keys(testCase.expected),
      );
      results[testCase.query] = entries;
      for (const entry of entries) {
        assert.equal(entry.command in testCase.expected, true);
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
        if (entry.command === "\\sqint" || entry.command === "\\sqiint") {
          assert.match(
            entry.operatorClass,
            /\bML__small-op\b/,
            `${entry.command} completion uses text-style integral geometry`,
          );
          assert.doesNotMatch(
            entry.operatorClass,
            /\bML__large-op\b/,
            `${entry.command} completion must not use display-style integral geometry`,
          );
          assert.equal(entry.svgCount, 1, `${entry.command} completion vector glyph`);
          assert.ok(
            entry.operatorHeight > 20 && entry.operatorHeight < 45,
            `${entry.command} completion height stays within the candidate row scale`,
          );
        }
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
      if (testCase.query === "\\sq") {
        const sqint = entries.find((entry) => entry.command === "\\sqint");
        const sqiint = entries.find((entry) => entry.command === "\\sqiint");
        const sqrt = entries.find((entry) => entry.command === "\\sqrt");
        for (const entry of [sqint, sqiint]) {
          assert.ok(entry, "square-integral candidate exists");
          assert.match(
            entry.operatorClass,
            /\bvisualtex-integral-svg\b/,
            `${entry.command} candidate uses the corrected SVG integral glyph`,
          );
          assert.equal(entry.svgCount, 1, `${entry.command} candidate SVG count`);
          assert.ok(entry.svgWidth > 4, `${entry.command} candidate SVG width`);
          assert.ok(entry.svgHeight > 10, `${entry.command} candidate SVG height`);
          assert.match(
            entry.svgViewBox,
            /^0 -?[\d.]+ [\d.]+ [\d.]+$/,
            `${entry.command} candidate SVG viewBox`,
          );
        }
        assert.ok(
          sqiint.operatorWidth > sqint.operatorWidth,
          "sqiint preview remains wider than sqint after contour repair",
        );
        assert.ok(sqrt, "sqrt candidate exists");
        assert.equal(sqrt.keybindingText, "⌥ V⌃ 2", "sqrt shortcut text");
        assert.equal(sqrt.keybindingWhiteSpace, "nowrap", "sqrt shortcut does not wrap inside a key chord");
        assert.equal(sqrt.keybindingFlexBasis, "46px", "sqrt shortcut column keeps reserved width");
        assert.ok(sqrt.keybindingWidth >= 45, "sqrt shortcut column is not squeezed");
        assert.ok(sqrt.keybindingHeight <= 30, "sqrt shortcuts stay within two compact lines");
        assert.equal(sqrt.itemHeight, 48, "sqrt candidate row keeps the standard height");
      }
    }

    if (fullAudit) {
      const commands = new Map();
      const prefixes =
        "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
      for (const prefix of prefixes) {
        await clearAndFocus();
        const query = `\\${prefix}`;
        await typeText(query);
        const started = Date.now();
        let prefixState;
        while (Date.now() - started < 4_000) {
          prefixState = await evaluate(`(() => {
            const query = ${JSON.stringify(`\\${prefix}`)};
            const field = document.querySelector("math-field.visual-mathfield");
            const stablePanel = document.getElementById(
              "visualtex-native-input-suggestion-popover",
            );
            const sourcePanel = document.getElementById("mathlive-suggestion-popover");
            const panel = stablePanel?.querySelector("li[data-command]")
              ? stablePanel
              : sourcePanel;
            const items = [...(panel?.querySelectorAll("li[data-command]") ?? [])];
            const rawLatex = [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
              .filter((node) => !node.classList.contains("ML__suggestion"))
              .map((node) => node.textContent ?? "")
              .join("");
            const commands = items.map((item) => item.dataset.command ?? "");
            return {
              ready:
                Boolean(panel?.classList.contains("is-visible")) &&
                commands.length > 0 &&
                commands.every((command) => command.startsWith(query)),
              noCandidates:
                rawLatex === query &&
                !panel?.classList.contains("is-visible"),
              rawLatex,
              commands,
            };
          })()`);
          if (prefixState?.ready) break;
          if (
            prefixState?.noCandidates &&
            Date.now() - started >= 500
          ) {
            break;
          }
          await sleep(50);
        }
        if (prefixState?.noCandidates) continue;
        assert.equal(
          prefixState?.ready,
          true,
          `native candidate prefix ${query}: ${JSON.stringify(prefixState)}`,
        );
        await sleep(160);
        const entries = await evaluate(`(() => {
          const query = ${JSON.stringify(`\\${prefix}`)};
          const stablePanel = document.getElementById(
            "visualtex-native-input-suggestion-popover",
          );
          const sourcePanel = document.getElementById("mathlive-suggestion-popover");
          const panel = stablePanel?.querySelector("li[data-command]")
            ? stablePanel
            : sourcePanel;
          const hasVisibleInk = (preview) => {
            const rendered = preview?.querySelector(".ML__latex") ?? preview;
            if (!rendered || rendered.querySelector(".ML__error")) return false;
            const text = (rendered.textContent ?? "")
              .replace(/[\\u200B-\\u200D\\u2060\\uFEFF]/g, "")
              .trim();
            if (text && text !== ".") return true;
            if (rendered.querySelector("svg, canvas, img, .ML__rule")) return true;
            return [rendered, ...rendered.querySelectorAll("*")].some((node) => {
              const style = getComputedStyle(node);
              if (
                style.display === "none" ||
                style.visibility === "hidden" ||
                Number.parseFloat(style.opacity || "1") <= 0
              ) return false;
              if (style.backgroundImage !== "none") return true;
              const hasBorder = [
                style.borderTopWidth,
                style.borderRightWidth,
                style.borderBottomWidth,
                style.borderLeftWidth,
              ].some((width) => Number.parseFloat(width) > 0);
              if (hasBorder && style.borderStyle !== "none") return true;
              return ["::before", "::after"].some((pseudo) => {
                const content = getComputedStyle(node, pseudo).content;
                return content && content !== "none" && content !== "normal" && content !== '\"\"';
              });
            });
          };
          return [...(panel?.querySelectorAll("li[data-command]") ?? [])]
            .filter((item) => (item.dataset.command ?? "").startsWith(query))
            .map((item) => {
              const command = item.dataset.command ?? "";
              const preview = item.querySelector(".ML__popover__command");
              const rendered = preview?.querySelector(".ML__latex") ?? preview;
              const keybinding = item.querySelector(".ML__popover__keybinding");
              const keybindingBounds = keybinding?.getBoundingClientRect();
              const keybindingStyle = keybinding ? getComputedStyle(keybinding) : null;
              return {
                command,
                kind: preview?.dataset.visualtexPreviewKind ?? "native",
                visible: hasVisibleInk(preview),
                error: Boolean(rendered?.querySelector(".ML__error")),
                cacheMismatch: Boolean(
                  preview?.dataset.visualtexPreview &&
                  preview.dataset.visualtexPreviewCommand !== command
                ),
                linkCount: rendered?.querySelectorAll("a[href]").length ?? 0,
                keybindingText: keybinding?.textContent?.trim() ?? "",
                keybindingWidth: keybindingBounds?.width ?? 0,
                keybindingWhiteSpace: keybindingStyle?.whiteSpace ?? "",
                keybindingFlexBasis: keybindingStyle?.flexBasis ?? "",
              };
            });
        })()`);
        for (const entry of entries) commands.set(entry.command, entry);
      }

      const allEntries = [...commands.values()];
      const blank = allEntries.filter((entry) => !entry.visible);
      const errors = allEntries.filter((entry) => entry.error);
      const cacheMismatch = allEntries.filter((entry) => entry.cacheMismatch);
      const links = allEntries.filter((entry) => entry.linkCount > 0);
      const brokenKeybindings = allEntries.filter(
        (entry) =>
          entry.keybindingText &&
          (entry.keybindingWhiteSpace !== "nowrap" ||
            entry.keybindingFlexBasis !== "46px" ||
            entry.keybindingWidth < 45),
      );
      assert.deepEqual(blank, [], "full native preview audit blank commands");
      assert.deepEqual(errors, [], "full native preview audit parse errors");
      assert.deepEqual(
        cacheMismatch,
        [],
        "full native preview audit stale command cache",
      );
      assert.deepEqual(links, [], "full native preview audit links");
      assert.deepEqual(
        brokenKeybindings,
        [],
        "full native preview audit wrapped or squeezed keybindings",
      );
      const categories = Object.fromEntries(
        [...new Set(allEntries.map((entry) => entry.kind))]
          .sort()
          .map((kind) => [
            kind,
            allEntries.filter((entry) => entry.kind === kind).length,
          ]),
      );
      results.fullAudit = {
        commandCount: allEntries.length,
        categories,
        blankCount: blank.length,
        errorCount: errors.length,
        cacheMismatchCount: cacheMismatch.length,
        linkCount: links.length,
        keybindingCount: allEntries.filter((entry) => entry.keybindingText).length,
        brokenKeybindingCount: brokenKeybindings.length,
      };
      console.log(JSON.stringify(results, null, 2));
    }
    console.log("Native suggestion preview regression passed.");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(120);
    await rm(chromeProfile, {
      recursive: true,
      force: true,
      maxRetries: 5,
      retryDelay: 100,
    });
  }
}

await main();
