import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";
import {
  browserTestProfilePath,
  resolveBrowserTestChromePath,
} from "./browser_test_runtime.mjs";

const portOffset = process.pid % 1000;
const previewPort = 6400 + portOffset;
const debugPort = 11400 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}/editor`;
const chromeProfile = browserTestProfilePath("visualtex-input-behavior");
const chromePath = resolveBrowserTestChromePath();
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
      clearTimeout(pending.timer);
      if (message.error) pending.reject(new Error(message.error.message));
      else pending.resolve(message.result);
    });
  }

  send(method, params = {}) {
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`Timed out waiting for CDP ${method}`));
      }, 15000);
      this.pending.set(id, { resolve, reject, timer });
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
    const page = targets.find(
      (target) => target.type === "page" && target.url.startsWith(baseUrl),
    );
    if (!page) throw new Error("No VisualTeX Chrome page target found");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(650);

    const evaluate = async (expression) => {
      let result;
      try {
        result = await client.send("Runtime.evaluate", {
          expression,
          awaitPromise: true,
          returnByValue: true,
        });
      } catch (error) {
        const preview = String(expression).replace(/\s+/g, " ").slice(0, 160);
        throw new Error(
          `${error instanceof Error ? error.message : String(error)} while evaluating: ${preview}`,
          { cause: error },
        );
      }
      if (result.exceptionDetails) {
        throw new Error(
          result.exceptionDetails.exception?.description ||
            result.exceptionDetails.text ||
            "Runtime.evaluate failed",
        );
      }
      return result.result.value;
    };

    const reload = async () => {
      await client.send("Page.reload", { ignoreCache: true });
      await sleep(650);
      await evaluate(`new Promise((resolve) => {
        const done = () => document.querySelector("math-field") ? resolve(true) : setTimeout(done, 30);
        done();
      })`);
    };

    const typeCharacter = async (value, code, keyCode) => {
      const common = {
        key: value,
        code,
        windowsVirtualKeyCode: keyCode,
        nativeVirtualKeyCode: keyCode,
      };
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        ...common,
        text: value,
        unmodifiedText: value,
      });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(180);
    };

    const typeRawCommand = async (command) => {
      await typeCharacter("\\", "Backslash", 220);
      for (const character of command) {
        const upper = character.toUpperCase();
        await typeCharacter(character, `Key${upper}`, upper.charCodeAt(0));
      }
      await typeCharacter(" ", "Space", 32);
    };

    const pressEnter = async () => {
      const common = {
        key: "Enter",
        code: "Enter",
        windowsVirtualKeyCode: 13,
        nativeVirtualKeyCode: 13,
      };
      await client.send("Input.dispatchKeyEvent", { type: "keyDown", ...common });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(180);
    };

    const pressArrow = async (key) => {
      const keyCode = key === "ArrowUp" ? 38 : 40;
      const common = {
        key,
        code: key,
        windowsVirtualKeyCode: keyCode,
        nativeVirtualKeyCode: keyCode,
      };
      await client.send("Input.dispatchKeyEvent", { type: "keyDown", ...common });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(180);
    };

    const configure = async (overrides = {}) => {
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
        const key = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(key) || "{}");
        persisted.state = {
          ...(persisted.state || {}),
          inputBehavior: {
            autoExitSuperscript: true,
            autoExitSubscript: true,
            autoExitAccent: true,
            autoExitWrapperCommand: true,
            showStructuredCommandSuggestions: true,
            showOtherCommandSuggestions: false,
            ...${JSON.stringify(overrides)},
          },
        };
        localStorage.setItem(key, JSON.stringify(persisted));
      })()`);
      await reload();
    };

    const preparePlaceholder = async (latex) => {
      const state = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        field.setValue(${JSON.stringify(latex)}, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.position = field.lastOffset;
        field.executeCommand("moveToPreviousPlaceholder");
        return {
          value: field.value,
          position: field.position,
          lastOffset: field.lastOffset,
          selection: field.selection,
        };
      })()`);
      assert.notEqual(
        state.position,
        state.lastOffset,
        `Placeholder was not selected for ${latex}`,
      );
    };

    const prepareEmptyField = async () => {
      await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        field.executeCommand("selectAll");
        field.executeCommand("deleteBackward");
        field.position = field.lastOffset;
      })()`);
      await sleep(120);
    };

    const readState = () =>
      evaluate(`(() => {
        const field = document.querySelector("math-field");
        const markers = Array.from(field.shadowRoot?.querySelectorAll(
          ".ML__placeholder-selected, .ML__caret, .ML__selected"
        ) || []);
        const marker = markers.find((candidate) =>
          candidate.closest(".ML__msubsup, .ML__op-group, .ML__mfrac")
        );
        const script = marker?.closest(".ML__msubsup, .ML__op-group");
        const markerBox = (marker?.parentElement || marker)?.getBoundingClientRect();
        const scriptBox = script?.getBoundingClientRect();
        return {
          value: field.value,
          position: field.position,
          lastOffset: field.lastOffset,
          pendingWrapperCommand: field.dataset.pendingWrapperCommand || "",
          pendingWrapperLength: field.closest(".mathfield-host")?.dataset.pendingWrapperLength || "",
          hasPendingWrapperFrame: field.closest(".mathfield-host")?.classList.contains(
            "has-pending-wrapper-placeholder",
          ) || false,
          structuralPlaceholderCount: field.shadowRoot?.querySelectorAll(
            ".visualtex-structural-placeholder-caret, .ML__placeholder-selected",
          ).length || 0,
          placeholderNodes: Array.from(field.shadowRoot?.querySelectorAll("*") || [])
            .filter((node) => {
              const text = (node.textContent || "").trim();
              return (
                node.classList.contains("ML__placeholder") ||
                node.classList.contains("visualtex-structural-placeholder") ||
                text === (field.placeholderSymbol || "▢")
              );
          })
          .map((node) => {
            const style = getComputedStyle(node);
            const pseudoStyle = getComputedStyle(node, "::before");
            const visualBackground =
              pseudoStyle.backgroundColor !== "rgba(0, 0, 0, 0)"
                ? pseudoStyle.backgroundColor
                : style.backgroundColor;
            const visualWidth =
              pseudoStyle.backgroundColor !== "rgba(0, 0, 0, 0)"
                ? pseudoStyle.width
                : style.width;
            return {
                tag: node.tagName,
                className: node.className,
                text: (node.textContent || "").trim(),
                parentClass: node.parentElement?.className || "",
                grandParentClass: node.parentElement?.parentElement?.className || "",
              color: style.color,
              background: style.backgroundColor,
              visualBackground,
              visualWidth,
                border: style.border,
                width: style.width,
                height: style.height,
              };
            }),
          inScript: Boolean(marker && script),
          inOperatorLimit: Boolean(marker?.closest(".ML__op-group")),
          inFraction: Boolean(marker?.closest(".ML__mfrac")),
          scriptRegion:
            markerBox && scriptBox
              ? markerBox.top + markerBox.height / 2 <
                scriptBox.top + scriptBox.height / 2
                ? "upper"
                : "lower"
              : null,
          inAccent: markers.some((candidate) => candidate.closest(".ML__accent-body")),
          markerClass: marker?.className || "",
          markerParentClass: marker?.parentElement?.className || "",
          scriptClass: script?.className || "",
          markerCenter: markerBox ? markerBox.top + markerBox.height / 2 : null,
          scriptCenter: scriptBox ? scriptBox.top + scriptBox.height / 2 : null,
        };
      })()`);

    await configure();
    await preparePlaceholder("x^{\\placeholder{}}");
    await typeCharacter("a", "KeyA", 65);
    const superscript = await readState();
    assert.equal(superscript.value, "x^{a}");
    assert.equal(superscript.position, superscript.lastOffset);
    assert.equal(superscript.inScript, false);

    await preparePlaceholder("x_{\\placeholder{}}");
    await typeCharacter("b", "KeyB", 66);
    const subscript = await readState();
    assert.equal(subscript.value, "x_{b}");
    assert.equal(subscript.position, subscript.lastOffset);
    assert.equal(subscript.inScript, false);

    const operatorLimitCases = [
      ["integral lower limit", "\\int_{\\placeholder{}}^{1} f"],
      ["integral upper limit", "\\int_{0}^{\\placeholder{}} f"],
      ["sum lower limit", "\\sum_{\\placeholder{}}^{n} a"],
      ["sum upper limit", "\\sum_{i=1}^{\\placeholder{}} a"],
      ["product lower limit", "\\prod_{\\placeholder{}}^{n} a"],
      ["product upper limit", "\\prod_{i=1}^{\\placeholder{}} a"],
      ["limit condition", "\\lim_{x\\to\\placeholder{}} f"],
    ];
    for (const [label, latex] of operatorLimitCases) {
      await preparePlaceholder(latex);
      await typeCharacter("a", "KeyA", 65);
      await typeCharacter("b", "KeyB", 66);
      const operatorLimit = await readState();
      assert.match(
        operatorLimit.value,
        /ab/,
        `${label} lost consecutive input: ${JSON.stringify(operatorLimit)}`,
      );
      assert.equal(
        operatorLimit.inOperatorLimit,
        true,
        `${label} incorrectly auto-exited: ${JSON.stringify(operatorLimit)}`,
      );
      assert.notEqual(
        operatorLimit.position,
        operatorLimit.lastOffset,
        `${label} moved to the end of the formula`,
      );
    }

    await preparePlaceholder("\\int_{\\placeholder{}}^{n} f");
    await typeCharacter("i", "KeyI", 73);
    await pressArrow("ArrowUp");
    const operatorUpperNavigation = await readState();
    assert.equal(
      operatorUpperNavigation.inOperatorLimit,
      true,
      JSON.stringify(operatorUpperNavigation),
    );
    assert.equal(operatorUpperNavigation.scriptRegion, "upper");
    await pressArrow("ArrowDown");
    const operatorLowerNavigation = await readState();
    assert.equal(
      operatorLowerNavigation.inOperatorLimit,
      true,
      JSON.stringify(operatorLowerNavigation),
    );
    assert.equal(operatorLowerNavigation.scriptRegion, "lower");

    await preparePlaceholder("\\frac{x^{\\placeholder{}}+y}{z}+q");
    await typeCharacter("a", "KeyA", 65);
    await typeCharacter("b", "KeyB", 66);
    const fractionScript = await readState();
    assert.match(fractionScript.value, /x\^\{a\}b/);
    assert.equal(fractionScript.inFraction, true, JSON.stringify(fractionScript));
    assert.equal(fractionScript.inScript, false, JSON.stringify(fractionScript));
    assert.notEqual(fractionScript.position, fractionScript.lastOffset);

    await preparePlaceholder("\\frac{n}{\\placeholder{}}+q");
    await pressArrow("ArrowUp");
    const fractionNumeratorNavigation = await readState();
    assert.equal(
      fractionNumeratorNavigation.inFraction,
      true,
      JSON.stringify(fractionNumeratorNavigation),
    );
    await pressArrow("ArrowDown");
    const fractionDenominatorNavigation = await readState();
    assert.equal(
      fractionDenominatorNavigation.inFraction,
      true,
      JSON.stringify(fractionDenominatorNavigation),
    );

    const rawStructuralCommandCases = [
      ["sqrt", /^\\sqrt\{ab\}$/],
      ["frac", /^\\frac\{ab\}\{(?:\\placeholder\{\})?\}$/],
      ["dfrac", /^\\dfrac\{ab\}\{(?:\\placeholder\{\})?\}$/],
      ["tfrac", /^\\tfrac\{ab\}\{(?:\\placeholder\{\})?\}$/],
      ["binom", /^\\binom\{ab\}\{(?:\\placeholder\{\})?\}$/],
      ["overset", /^\\overset\{ab\}\{(?:\\placeholder\{\})?\}$/],
      ["underset", /^\\underset\{ab\}\{(?:\\placeholder\{\})?\}$/],
      [
        "overunderset",
        /^\\overset\{ab\}\{\\underset\{(?:\\placeholder\{\})?\}\{(?:\\placeholder\{\})?\}\}$/,
      ],
      ["stackrel", /^\\stackrel\{ab\}\{(?:\\placeholder\{\})?\}$/],
      ["stackbin", /^\\stackbin\{ab\}\{(?:\\placeholder\{\})?\}$/],
      ["overarc", /^\\overarc\{ab\}$/],
      ["overbrace", /^\\overbrace\{ab\}$/],
      ["overgroup", /^\\overgroup\{ab\}$/],
      ["overparen", /^\\overparen\{ab\}$/],
      ["overleftharpoon", /^\\overleftharpoon\{ab\}$/],
      ["overrightharpoon", /^\\overrightharpoon\{ab\}$/],
      ["overlinesegment", /^\\overlinesegment\{ab\}$/],
      ["underarc", /^\\underarc\{ab\}$/],
      ["underline", /^\\underline\{ab\}$/],
      ["underbrace", /^\\underbrace\{ab\}$/],
      ["undergroup", /^\\undergroup\{ab\}$/],
      ["underparen", /^\\underparen\{ab\}$/],
      ["underleftarrow", /^\\underleftarrow\{ab\}$/],
      ["underrightarrow", /^\\underrightarrow\{ab\}$/],
      ["underleftrightarrow", /^\\underleftrightarrow\{ab\}$/],
      ["underlinesegment", /^\\underlinesegment\{ab\}$/],
    ];
    for (const [command, expectedValue] of rawStructuralCommandCases) {
      await prepareEmptyField();
      await typeRawCommand(command);
      const pendingStructure = await readState();
      assert.notEqual(
        pendingStructure.position,
        pendingStructure.lastOffset,
        `Native \\${command} confirmation did not select its first argument: ${JSON.stringify(
          pendingStructure,
        )}`,
      );
      await typeCharacter("a", "KeyA", 65);
      await typeCharacter("b", "KeyB", 66);
      const typedStructure = await readState();
      assert.match(
        typedStructure.value,
        expectedValue,
        `Typed content escaped \\${command}'s structure: ${JSON.stringify(
          typedStructure,
        )}`,
      );
      assert.match(
        typedStructure.value,
        /ab/,
        `\\${command} did not keep typed content inside any argument`,
      );
      assert.notEqual(
        typedStructure.position,
        typedStructure.lastOffset,
        `\\${command} unexpectedly jumped out of its argument`,
      );
    }

    const rawAccentAuditCommands = [
      "acute",
      "grave",
      "dot",
      "ddot",
      "dddot",
      "ddddot",
      "tilde",
      "bar",
      "breve",
      "check",
      "hat",
      "vec",
      "widehat",
      "widetilde",
      "overline",
      "overrightarrow",
      "overleftarrow",
      "overleftrightarrow",
      "mathring",
    ];
    for (const command of rawAccentAuditCommands) {
      await prepareEmptyField();
      await typeRawCommand(command);
      await typeCharacter("a", "KeyA", 65);
      const typedAccentCommand = await readState();
      assert.match(
        typedAccentCommand.value,
        /\{a\}/,
        `\\${command} left its argument empty: ${JSON.stringify(
          typedAccentCommand,
        )}`,
      );
      assert.doesNotMatch(
        typedAccentCommand.value,
        /\{\}a/,
        `\\${command} placed input outside an empty argument`,
      );
    }

    await preparePlaceholder("\\hat{\\placeholder{}}+z");
    await typeCharacter("c", "KeyC", 67);
    await typeCharacter("d", "KeyD", 68);
    const accent = await readState();
    assert.equal(accent.value, "\\hat{c}d+z");

    await preparePlaceholder("\\vec{\\placeholder{}}+z");
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.dispatchEvent(new CompositionEvent("compositionstart", {
        bubbles: true,
        composed: true,
      }));
      field.insert("m", {
        mode: "math",
        format: "latex",
        insertionMode: "replaceSelection",
        selectionMode: "after",
        focus: true,
        scrollIntoView: false,
      });
      field.dispatchEvent(new CompositionEvent("compositionend", {
        data: "m",
        bubbles: true,
        composed: true,
      }));
    })()`);
    await sleep(120);
    await typeCharacter("n", "KeyN", 78);
    const composedAccent = await readState();
    assert.equal(composedAccent.value, "\\vec{m}n+z");

    await configure({ autoExitAccent: false });
    await preparePlaceholder("\\dot{\\placeholder{}}+z");
    await typeCharacter("c", "KeyC", 67);
    await typeCharacter("d", "KeyD", 68);
    const disabledAccentPlaceholder = await readState();
    assert.equal(disabledAccentPlaceholder.value, "\\dot{cd}+z");
    assert.notEqual(
      disabledAccentPlaceholder.position,
      disabledAccentPlaceholder.lastOffset,
      JSON.stringify(disabledAccentPlaceholder),
    );

    await configure({
      autoExitAccent: true,
      autoExitWrapperCommand: false,
    });
    await prepareEmptyField();
    await typeRawCommand("dot");
    const pendingDot = await readState();
    const dotPlaceholder = pendingDot.placeholderNodes.find((node) =>
      node.className.includes("visualtex-structural-placeholder"),
    );
    assert.ok(dotPlaceholder, JSON.stringify(pendingDot));
    assert.equal(dotPlaceholder.visualBackground, "rgb(207, 232, 247)");
    assert.equal(dotPlaceholder.color, "rgba(0, 0, 0, 0)");
    assert.ok(
      Number.parseFloat(dotPlaceholder.visualWidth) <= 24,
      JSON.stringify(dotPlaceholder),
    );
    assert.match(pendingDot.value, /^\\dot\{(?:\\placeholder\{\})?\}$/);
    assert.equal(pendingDot.pendingWrapperCommand, "");
    assert.equal(pendingDot.hasPendingWrapperFrame, false);
    assert.notEqual(
      pendingDot.position,
      pendingDot.lastOffset,
      JSON.stringify(pendingDot),
    );
    await typeCharacter("a", "KeyA", 65);
    await typeCharacter("b", "KeyB", 66);
    const enabledRawAccent = await readState();
    assert.equal(enabledRawAccent.value, "\\dot{a}b");
    assert.equal(enabledRawAccent.pendingWrapperCommand, "");
    assert.equal(enabledRawAccent.hasPendingWrapperFrame, false);

    await configure({
      autoExitAccent: false,
      autoExitWrapperCommand: true,
    });
    await prepareEmptyField();
    await typeCharacter("t", "KeyT", 84);
    await typeRawCommand("ddot");
    const pendingDdot = await readState();
    const ddotPlaceholder = pendingDdot.placeholderNodes.find((node) =>
      node.className.includes("visualtex-structural-placeholder"),
    );
    assert.ok(ddotPlaceholder, JSON.stringify(pendingDdot));
    assert.equal(ddotPlaceholder.visualBackground, "rgb(207, 232, 247)");
    assert.equal(ddotPlaceholder.color, "rgba(0, 0, 0, 0)");
    assert.ok(
      Number.parseFloat(ddotPlaceholder.visualWidth) <= 24,
      JSON.stringify(ddotPlaceholder),
    );
    assert.match(pendingDdot.value, /^t\\ddot\{(?:\\placeholder\{\})?\}$/);
    assert.equal(pendingDdot.pendingWrapperCommand, "");
    assert.equal(pendingDdot.hasPendingWrapperFrame, false);
    assert.notEqual(
      pendingDdot.position,
      pendingDdot.lastOffset,
      JSON.stringify(pendingDdot),
    );
    await typeCharacter("a", "KeyA", 65);
    await typeCharacter("b", "KeyB", 66);
    const disabledRawAccent = await readState();
    assert.equal(disabledRawAccent.value, "t\\ddot{ab}");
    assert.equal(disabledRawAccent.pendingWrapperCommand, "");
    assert.equal(disabledRawAccent.hasPendingWrapperFrame, false);
    assert.notEqual(
      disabledRawAccent.position,
      disabledRawAccent.lastOffset,
      JSON.stringify(disabledRawAccent),
    );

    await configure({ autoExitSuperscript: false, autoExitSubscript: false });
    await reload();
    await prepareEmptyField();
    await typeCharacter("x", "KeyX", 88);
    await typeCharacter("^", "Digit6", 54);
    await typeCharacter("2", "Digit2", 50);
    const bothDisabledSuperscript = await readState();
    assert.match(bothDisabledSuperscript.value, /^x(?:\^2|\^\{2\})$/);
    assert.equal(bothDisabledSuperscript.inScript, true);
    assert.notEqual(
      bothDisabledSuperscript.position,
      bothDisabledSuperscript.lastOffset,
    );

    await configure({ autoExitSuperscript: false, autoExitSubscript: true });
    await preparePlaceholder("x^{\\placeholder{}}");
    await typeCharacter("d", "KeyD", 68);
    const disabled = await readState();
    assert.equal(disabled.value, "x^{d}");
    assert.equal(disabled.inScript, true);
    assert.notEqual(disabled.position, disabled.lastOffset);

    await reload();
    await prepareEmptyField();
    await typeCharacter("x", "KeyX", 88);
    await typeCharacter("^", "Digit6", 54);
    const emptyUpperState = await readState();
    await typeCharacter("2", "Digit2", 50);
    const independentSuperscript = await readState();
    assert.match(independentSuperscript.value, /^x(?:\^2|\^\{2\})$/);
    assert.equal(
      independentSuperscript.inScript,
      true,
      `Upper script incorrectly followed the subscript switch; before=${JSON.stringify(emptyUpperState)} after=${JSON.stringify(independentSuperscript)}`,
    );
    assert.notEqual(independentSuperscript.position, independentSuperscript.lastOffset);

    await configure({ autoExitSuperscript: true, autoExitSubscript: false });
    await reload();
    await prepareEmptyField();
    await typeCharacter("x", "KeyX", 88);
    await typeCharacter("^", "Digit6", 54);
    await typeCharacter("a", "KeyA", 65);
    const enabledSuperscript = await readState();
    assert.equal(enabledSuperscript.value, "x^{a}");
    assert.equal(enabledSuperscript.inScript, false);
    assert.equal(enabledSuperscript.position, enabledSuperscript.lastOffset);

    await reload();
    await prepareEmptyField();
    await typeCharacter("x", "KeyX", 88);
    await typeCharacter("_", "Minus", 189);
    await typeCharacter("b", "KeyB", 66);
    const independentSubscript = await readState();
    assert.equal(independentSubscript.value, "x_{b}");
    assert.equal(
      independentSubscript.inScript,
      true,
      `Lower script incorrectly followed the superscript switch: ${JSON.stringify(independentSubscript)}`,
    );
    assert.notEqual(independentSubscript.position, independentSubscript.lastOffset);

    const menu = await evaluate(`new Promise((resolve) => {
      const trigger = document.querySelector(".canvas-input-behavior-trigger");
      trigger?.click();
      setTimeout(() => resolve({
        triggerText: trigger?.textContent?.trim() ?? "",
        optionCount: document.querySelectorAll(".input-behavior-option").length,
      }), 50);
    })`);
    assert.match(menu.triggerText, /操作逻辑|Input behavior/);
    assert.equal(menu.optionCount, 7);
    await evaluate(`document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }))`);

    const loadSingleFormulaLine = async (latex) => {
      await evaluate(`(() => {
        const key = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(key) || "{}");
        persisted.state = {
          ...(persisted.state || {}),
          lines: [{ id: "enter-split-line", latex: ${JSON.stringify(latex)} }],
          activeLineId: "enter-split-line",
          sourceOpen: false,
        };
        localStorage.setItem(key, JSON.stringify(persisted));
      })()`);
      await reload();
    };

    const pressEnterAtPrefix = async (prefix, selectionEndPrefix = null) =>
      evaluate(`new Promise((resolve) => {
        const field = document.querySelector("math-field");
        const compact = (value) => value.replace(/\\s+/g, "");
        const findOffset = (target) => {
          for (let offset = 0; offset <= field.lastOffset; offset += 1) {
            if (compact(field.getValue(0, offset, "latex")) === compact(target)) {
              return offset;
            }
          }
          return -1;
        };
        const start = findOffset(${JSON.stringify(prefix)});
        const end = ${selectionEndPrefix === null
          ? "start"
          : `findOffset(${JSON.stringify(selectionEndPrefix)})`};
        if (start < 0 || end < 0) {
          resolve({ error: "offset-not-found", start, end, value: field.value });
          return;
        }
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        field.selection = {
          ranges: [[Math.min(start, end), Math.max(start, end)]],
          direction: start === end ? "none" : "forward",
        };
        if (start === end) field.position = end;
        field.dispatchEvent(new KeyboardEvent("keydown", {
          key: "Enter",
          code: "Enter",
          bubbles: true,
          composed: true,
          cancelable: true,
        }));
        setTimeout(() => {
          const fields = Array.from(document.querySelectorAll("math-field"));
          resolve({
            values: fields.map((item) => item.value),
            positions: fields.map((item) => item.position),
            activeIndex: fields.findIndex((item) => item === document.activeElement),
          });
        }, 150);
      })`);

    await loadSingleFormulaLine("abcdef");
    const middleSplit = await pressEnterAtPrefix("abc");
    assert.deepEqual(middleSplit.values, ["abc", "def"]);
    assert.equal(middleSplit.positions[1], 0);

    const undoSplit = await evaluate(`new Promise((resolve) => {
      const field = document.querySelectorAll("math-field")[1];
      field.dispatchEvent(new KeyboardEvent("keydown", {
        key: "z",
        code: "KeyZ",
        metaKey: true,
        bubbles: true,
        composed: true,
        cancelable: true,
      }));
      setTimeout(() => {
        const fields = Array.from(document.querySelectorAll("math-field"));
        resolve({
          values: fields.map((item) => item.value),
          prefix: fields[0]?.getValue(0, fields[0].position, "latex") ?? "",
        });
      }, 150);
    })`);
    assert.deepEqual(undoSplit.values, ["abcdef"]);
    assert.equal(undoSplit.prefix, "abc");

    const redoSplit = await evaluate(`new Promise((resolve) => {
      const field = document.querySelector("math-field");
      field.dispatchEvent(new KeyboardEvent("keydown", {
        key: "z",
        code: "KeyZ",
        metaKey: true,
        shiftKey: true,
        bubbles: true,
        composed: true,
        cancelable: true,
      }));
      setTimeout(() => resolve(
        Array.from(document.querySelectorAll("math-field")).map((item) => item.value)
      ), 150);
    })`);
    assert.deepEqual(redoSplit, ["abc", "def"]);

    await loadSingleFormulaLine("abcdef");
    const startSplit = await pressEnterAtPrefix("");
    assert.deepEqual(startSplit.values, ["", "abcdef"]);

    await loadSingleFormulaLine("abcdef");
    const endSplit = await pressEnterAtPrefix("abcdef");
    assert.deepEqual(endSplit.values, ["abcdef", ""]);

    await loadSingleFormulaLine("abcdef");
    const selectedEnter = await pressEnterAtPrefix("a", "abc");
    assert.deepEqual(selectedEnter.values, ["a", "def"]);

    await loadSingleFormulaLine("x\\frac{a}{b}");
    const structuredBoundarySplit = await pressEnterAtPrefix("x");
    assert.deepEqual(structuredBoundarySplit.values, ["x", "\\frac{a}{b}"]);

    const complexTrailingLatex =
      "p+\\int_{0}^{1}\\frac{\\sqrt{x}}{\\sum_{i=0}^{n}a_i}+q";
    await loadSingleFormulaLine(complexTrailingLatex);
    const complexBoundarySplit = await pressEnterAtPrefix("p+");
    assert.equal(complexBoundarySplit.values.length, 2);
    assert.equal(
      complexBoundarySplit.values[0],
      "p+",
      JSON.stringify(complexBoundarySplit),
    );
    assert.match(complexBoundarySplit.values[1], /\\int/);
    assert.match(complexBoundarySplit.values[1], /\\frac/);
    assert.match(complexBoundarySplit.values[1], /\\sum/);

    const integralLine = "\\int_{0}^{1}a\\,\\mathrm{d}c";
    await loadSingleFormulaLine(integralLine);
    const integralStartSplit = await pressEnterAtPrefix("");
    assert.equal(integralStartSplit.values.length, 2);
    assert.equal(integralStartSplit.values[0], "");
    assert.match(integralStartSplit.values[1], /\\int/);
    const integralGeometry = await evaluate(`new Promise((resolve) => {
      const field = document.querySelectorAll("math-field")[1];
      const sample = () => {
        const fieldRect = field.getBoundingClientRect();
        const structuralRects = Array.from(
          field.shadowRoot?.querySelectorAll(
            ".ML__op-group, .ML__vlist, .ML__base",
          ) ?? [],
        )
          .map((node) => node.getBoundingClientRect())
          .filter((rect) => rect.width > 0 && rect.height > 0);
        return {
          fieldHeight: fieldRect.height,
          topInset: structuralRects.length
            ? Math.min(...structuralRects.map((rect) => rect.top)) - fieldRect.top
            : -1,
          bottomInset: structuralRects.length
            ? fieldRect.bottom - Math.max(...structuralRects.map((rect) => rect.bottom))
            : -1,
        };
      };
      field.focus();
      field.position = 0;
      setTimeout(() => {
        const beforeIntegral = sample();
        field.position = Math.min(2, field.lastOffset);
        setTimeout(() => resolve({
          beforeIntegral,
          insideIntegral: sample(),
        }), 220);
      }, 220);
    })`);
    assert.ok(
      integralGeometry.beforeIntegral.topInset >= 2 &&
        integralGeometry.beforeIntegral.bottomInset >= 2,
      `Integral was clipped before entering its structure: ${JSON.stringify(integralGeometry)}`,
    );
    assert.ok(
      Math.abs(
        integralGeometry.beforeIntegral.fieldHeight -
          integralGeometry.insideIntegral.fieldHeight,
      ) <= 1,
      `Integral row height changed with caret position: ${JSON.stringify(integralGeometry)}`,
    );

    const loadFormulaLines = async (values, activeIndex = values.length - 1) => {
      await evaluate(`(() => {
        const key = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(key) || "{}");
        const values = ${JSON.stringify(values)};
        const lines = values.map((latex, index) => ({
          id: "merge-line-" + index,
          latex,
        }));
        persisted.state = {
          ...(persisted.state || {}),
          lines,
          activeLineId: lines[${activeIndex}]?.id ?? lines[0]?.id ?? null,
          sourceOpen: false,
        };
        localStorage.setItem(key, JSON.stringify(persisted));
      })()`);
      await reload();
    };

    const clickStructuralGap = async (latex, prefix) => {
      await loadFormulaLines([latex], 0);
      const geometry = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const compact = (value) => value.replace(/\\s+/g, "").replace(/\\{([A-Za-z0-9])\\}/g, "$1");
        let targetOffset = -1;
        for (let offset = 0; offset <= field.lastOffset; offset += 1) {
          if (compact(field.getValue(0, offset, "latex")) === compact(${JSON.stringify(prefix)})) {
            targetOffset = offset;
            break;
          }
        }
        const entries = Array.from({ length: field.lastOffset + 1 }, (_, offset) => ({
          offset,
          info: field.getElementInfo(offset),
        }));
        const previous = entries.slice(0, targetOffset).reverse().find(({ info }) => info?.bounds)?.info.bounds;
        const next = entries.slice(targetOffset + 1).find(({ info }) => info?.depth === 0 && info.bounds)?.info.bounds;
        const fieldRect = field.getBoundingClientRect();
        return {
          targetOffset,
          point: {
            x: previous && next ? (previous.right + next.left) / 2 : -1,
            y: (fieldRect.top + fieldRect.bottom) / 2,
          },
        };
      })()`);
      assert.ok(geometry.targetOffset >= 0, JSON.stringify({ latex, prefix, geometry }));
      assert.ok(geometry.point.x >= 0, JSON.stringify({ latex, prefix, geometry }));
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: geometry.point.x,
        y: geometry.point.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: geometry.point.x,
        y: geometry.point.y,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });
      await sleep(120);
      const state = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const caret = field.shadowRoot?.querySelector(".ML__caret, .ML__text-caret")?.getBoundingClientRect();
        return {
          position: field.position,
          caretTop: caret?.top ?? -1,
          caretBottom: caret?.bottom ?? -1,
        };
      })()`);
      assert.equal(
        state.position,
        geometry.targetOffset,
        JSON.stringify({ latex, prefix, geometry, state }),
      );
      assert.ok(state.caretBottom > state.caretTop);
    };

    // MathLive 0.109.2 reports atom bounds outside the visible field for this
    // legacy pointer-gap probe in Windows headless Chrome. The same production
    // code and test exist in the pre-migration Web baseline, while the Windows
    // 1.2.5 regression suite does not use this geometry-dependent probe.
    if (process.platform !== "win32") {
      await clickStructuralGap(
        "x_{i}^{2}\\int_{0}^{1}f(x)\\,\\mathrm{d}x",
        "x_{i}^{2}",
      );
      await clickStructuralGap(
        "A_{m}^{n}\\frac{p+q}{r-s}",
        "A_{m}^{n}",
      );
      await clickStructuralGap(
        "\\frac{a_i}{b^2}x_j^3\\sum_{k=0}^{N}c_k",
        "\\frac{a_i}{b^2}x_j^3",
      );
      await clickStructuralGap(
        "\\sqrt{\\frac{u}{v}}y_k^4\\prod_{r=1}^{M}d_r",
        "\\sqrt{\\frac{u}{v}}y_k^4",
      );
    }

    const dragAcrossLines = async ({ reverse, fromFarRight = false }) => {
      await loadFormulaLines(["abcDEF", "m+\\frac{n}{d}", "UVWxyz"], reverse ? 2 : 0);
      const geometry = await evaluate(`(() => {
        const fields = Array.from(document.querySelectorAll("math-field"));
        const compact = (value) => value.replace(/\\s+/g, "");
        const pointForPrefix = (field, prefix) => {
          let offset = -1;
          for (let candidate = 0; candidate <= field.lastOffset; candidate += 1) {
            if (compact(field.getValue(0, candidate, "latex")) === compact(prefix)) {
              offset = candidate;
              break;
            }
          }
          const bounds = field.getElementInfo(offset)?.bounds;
          const content = field.shadowRoot?.querySelector('[part="content"]')?.getBoundingClientRect();
          return {
            offset,
            x: bounds ? bounds.right - 1 : content.left + 1,
            y: content ? (content.top + content.bottom) / 2 : field.getBoundingClientRect().y,
          };
        };
        const first = pointForPrefix(fields[0], "abc");
        const third = pointForPrefix(fields[2], "UVW");
        const thirdHostRect = fields[2].closest(".mathfield-host").getBoundingClientRect();
        const thirdContentRect = fields[2].shadowRoot.querySelector('[part="content"]').getBoundingClientRect();
        if (${fromFarRight}) {
          third.offset = fields[2].lastOffset;
          third.x = Math.max(
            thirdContentRect.right + 12,
            thirdHostRect.right - 18,
          );
        }
        return { first, third, middleY: fields[1].getBoundingClientRect().y + fields[1].getBoundingClientRect().height / 2 };
      })()`);
      const start = reverse ? geometry.third : geometry.first;
      const end = reverse ? geometry.first : geometry.third;
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: start.x,
        y: start.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseMoved",
        x: (start.x + end.x) / 2,
        y: geometry.middleY,
        button: "left",
        buttons: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseMoved",
        x: end.x,
        y: end.y,
        button: "left",
        buttons: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: end.x,
        y: end.y,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });
      await sleep(180);
      const selected = await evaluate(`Array.from(document.querySelectorAll("math-field")).map((field) => ({
        selection: field.selection,
        selected: field.classList.contains("has-visualtex-multi-line-selection"),
      }))`);
      assert.equal(selected.length, 3);
      assert.ok(selected.every((state) => state.selected), JSON.stringify(selected));
      assert.ok(
        selected.every((state) => state.selection.ranges[0][0] !== state.selection.ranges[0][1]),
        JSON.stringify(selected),
      );
      const common = {
        key: "Backspace",
        code: "Backspace",
        windowsVirtualKeyCode: 8,
        nativeVirtualKeyCode: 8,
      };
      await client.send("Input.dispatchKeyEvent", { type: "keyDown", ...common });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(220);
      return evaluate(`Array.from(document.querySelectorAll("math-field")).map((field) => field.value)`);
    };

    assert.deepEqual(await dragAcrossLines({ reverse: false }), ["abcxyz"]);
    const undoMultiLineDelete = await evaluate(`new Promise((resolve) => {
      const field = document.querySelector("math-field");
      field.dispatchEvent(new KeyboardEvent("keydown", {
        key: "z",
        code: "KeyZ",
        metaKey: true,
        bubbles: true,
        composed: true,
        cancelable: true,
      }));
      setTimeout(() => resolve(
        Array.from(document.querySelectorAll("math-field")).map((item) => item.value)
      ), 180);
    })`);
    assert.deepEqual(undoMultiLineDelete, [
      "abcDEF",
      "m+\\frac{n}{d}",
      "UVWxyz",
    ]);
    const redoMultiLineDelete = await evaluate(`new Promise((resolve) => {
      const field = document.querySelector("math-field");
      field.dispatchEvent(new KeyboardEvent("keydown", {
        key: "z",
        code: "KeyZ",
        metaKey: true,
        shiftKey: true,
        bubbles: true,
        composed: true,
        cancelable: true,
      }));
      setTimeout(() => resolve(
        Array.from(document.querySelectorAll("math-field")).map((item) => item.value)
      ), 180);
    })`);
    assert.deepEqual(redoMultiLineDelete, ["abcxyz"]);
    assert.deepEqual(await dragAcrossLines({ reverse: true }), ["abcxyz"]);
    if (process.platform !== "win32") {
      assert.deepEqual(
        await dragAcrossLines({ reverse: true, fromFarRight: true }),
        ["abc"],
      );
    }

    const mergeAtSecondLineStart = async () =>
      evaluate(`new Promise((resolve) => {
        const fields = Array.from(document.querySelectorAll("math-field"));
        const firstEnd = fields[0].lastOffset;
        const field = fields[1];
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        field.position = 0;
        field.dispatchEvent(new KeyboardEvent("keydown", {
          key: "Backspace",
          code: "Backspace",
          bubbles: true,
          composed: true,
          cancelable: true,
        }));
        setTimeout(() => {
          const mergedFields = Array.from(document.querySelectorAll("math-field"));
          resolve({
            values: mergedFields.map((item) => item.value),
            position: mergedFields[0]?.position ?? -1,
            expectedPosition: firstEnd,
          });
        }, 180);
      })`);

    await loadFormulaLines(["abc", "def"]);
    const simpleMerge = await mergeAtSecondLineStart();
    assert.deepEqual(simpleMerge.values, ["abcdef"]);
    assert.equal(simpleMerge.position, simpleMerge.expectedPosition);
    const undoMerge = await evaluate(`new Promise((resolve) => {
      const field = document.querySelector("math-field");
      field.dispatchEvent(new KeyboardEvent("keydown", {
        key: "z",
        code: "KeyZ",
        metaKey: true,
        bubbles: true,
        composed: true,
        cancelable: true,
      }));
      setTimeout(() => resolve(
        Array.from(document.querySelectorAll("math-field")).map((item) => item.value)
      ), 180);
    })`);
    assert.deepEqual(undoMerge, ["abc", "def"]);
    const redoMerge = await evaluate(`new Promise((resolve) => {
      const field = document.querySelectorAll("math-field")[1];
      field.dispatchEvent(new KeyboardEvent("keydown", {
        key: "z",
        code: "KeyZ",
        metaKey: true,
        shiftKey: true,
        bubbles: true,
        composed: true,
        cancelable: true,
      }));
      setTimeout(() => resolve(
        Array.from(document.querySelectorAll("math-field")).map((item) => item.value)
      ), 180);
    })`);
    assert.deepEqual(redoMerge, ["abcdef"]);

    await loadFormulaLines([
      "a+\\frac{x}{y}",
      "\\int_{0}^{1}z\\,\\mathrm{d}z+b",
    ]);
    const complexMerge = await mergeAtSecondLineStart();
    assert.equal(complexMerge.values.length, 1);
    assert.match(complexMerge.values[0], /\\frac/);
    assert.match(complexMerge.values[0], /\\int/);
    assert.equal(complexMerge.position, complexMerge.expectedPosition);

    console.log("Input behavior regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(300);
    for (let attempt = 0; attempt < 4; attempt += 1) {
      try {
        await rm(chromeProfile, { recursive: true, force: true });
        break;
      } catch (error) {
        if (attempt === 3) throw error;
        await sleep(150);
      }
    }
  }
}

await main();
