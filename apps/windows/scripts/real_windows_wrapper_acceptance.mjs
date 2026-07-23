import assert from "node:assert/strict";
import { mkdir, writeFile } from "node:fs/promises";
import { join, resolve } from "node:path";
import process from "node:process";

const debugPort = Number.parseInt(process.env.VISUALTEX_CDP_PORT ?? "19333", 10);
const endpoint = `http://127.0.0.1:${debugPort}`;
const artifactRoot = resolve(
  process.cwd(),
  "artifacts",
  `real-windows-wrapper-${new Date().toISOString().replace(/[:.]/g, "-")}`,
);
const sleep = (ms) => new Promise((resolvePromise) => setTimeout(resolvePromise, ms));

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
  const targets = await (await fetch(`${endpoint}/json/list`)).json();
  const page = targets.find(
    (target) => target.type === "page" && target.url.startsWith("http://tauri.localhost"),
  );
  if (!page) throw new Error(`No VisualTeX WebView2 page found at ${endpoint}`);

  const client = new CdpClient(page.webSocketDebuggerUrl);
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

  const waitForEvaluation = async (expression, description, timeoutMs = 7000) => {
    const started = Date.now();
    let lastValue;
    while (Date.now() - started < timeoutMs) {
      lastValue = await evaluate(expression);
      if (lastValue?.ready) return lastValue;
      await sleep(60);
    }
    throw new Error(`Timed out waiting for ${description}: ${JSON.stringify(lastValue)}`);
  };

  const screenshot = async (name) => {
    const capture = await client.send("Page.captureScreenshot", {
      format: "png",
      captureBeyondViewport: false,
    });
    await writeFile(join(artifactRoot, name), Buffer.from(capture.data, "base64"));
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
    await sleep(90);
  };

  const typeText = async (text) => {
    for (const character of Array.from(text)) {
      const code = character === "\\" ? "Backslash" : `Key${character.toUpperCase()}`;
      const virtualKeyCode = character === "\\"
        ? 220
        : character.toUpperCase().charCodeAt(0);
      await key(character, code, virtualKeyCode);
    }
  };

  const wrapperStateExpression = `(() => {
    const field = document.querySelector("math-field");
    const host = field?.closest(".mathfield-host");
    const pseudo = host ? getComputedStyle(host, "::after") : null;
    const caret = field?.shadowRoot?.querySelector(".ML__caret");
    const placeholders = [...(field?.shadowRoot?.querySelectorAll(".ML__placeholder") ?? [])]
      .map((node) => ({ text: node.textContent ?? "", className: node.className }));
    return {
      ready: Boolean(field && host),
      value: field?.value ?? "",
      position: field?.position ?? -1,
      lastOffset: field?.lastOffset ?? -1,
      selection: field?.selection ?? null,
      pending: field?.dataset.pendingWrapperCommand ?? "",
      hostPending: host?.dataset.pendingWrapperCommand ?? "",
      hasBox: host?.classList.contains("has-pending-wrapper-placeholder") ?? false,
      boxWidth: Number.parseFloat(pseudo?.width ?? "NaN"),
      boxHeight: Number.parseFloat(pseudo?.height ?? "NaN"),
      boxBorder: pseudo?.borderTopWidth ?? "",
      boxContent: pseudo?.content ?? "",
      caretClass: caret?.className ?? "",
      placeholders,
      shadowText: field?.shadowRoot?.textContent ?? "",
    };
  })()`;

  const resetDocument = async (
    autoExitWrapperCommand,
    latex = "",
    autoExitSuperscript = true,
    autoExitSubscript = true,
  ) => {
    await evaluate(`(() => {
      const lines = [{ id: crypto.randomUUID(), latex: ${JSON.stringify(latex)} }];
      localStorage.setItem("visualtex-editor", JSON.stringify({
        state: {
          title: "Real wrapper acceptance",
          lines,
          activeLineId: lines[0].id,
          language: "cn",
          inputBehavior: {
            autoExitSuperscript: ${autoExitSuperscript},
            autoExitSubscript: ${autoExitSubscript},
            autoExitAccent: true,
            autoExitWrapperCommand: ${autoExitWrapperCommand},
            showStructuredCommandSuggestions: true,
            showOtherCommandSuggestions: false
          }
        },
        version: 0
      }));
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.onboarding.windows.desktop.v1.1.0.completed", "true");
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await waitForEvaluation(
      `(() => ({ ready: Boolean(document.querySelector("math-field") && document.querySelector('button[data-category="matrix"]')) }))()`,
      "VisualTeX formula field and toolbar categories",
    );
    await sleep(200);
  };

  const focusField = async (position = null) => {
    await waitForEvaluation(`(() => {
      const field = document.querySelector("math-field");
      if (!field) return { ready: false };
      field.focus();
      field.position = ${position === null ? "field.lastOffset" : position};
      const sink = field.shadowRoot?.querySelector('[part="keyboard-sink"]');
      sink?.focus({ preventScroll: true });
      return {
        ready: field.matches(":focus-within") && field.shadowRoot?.activeElement === sink,
      };
    })()`, "focused real WebView2 math field");
  };

  const clickMathbbToolbar = async (position = null) => {
    await focusField(position);
    await evaluate(`document.querySelector('button[data-category="matrix"]')?.click()`);
    await waitForEvaluation(
      `(() => ({ ready: Boolean(document.querySelector('[data-command-id="blackboard-bold"]')) }))()`,
      "mathbb toolbar button",
    );
    await evaluate(`document.querySelector('[data-command-id="blackboard-bold"]')?.click()`);
    await sleep(250);
  };

  const setInputBehaviorThroughUi = async (title, enabled) => {
    await evaluate(`document.querySelector(".canvas-input-behavior-trigger")?.click()`);
    await waitForEvaluation(`(() => ({
      ready: Boolean(document.querySelector(".input-behavior-popover")),
    }))()`, "input behavior popover");
    const changed = await evaluate(`(() => {
      const title = ${JSON.stringify(title)};
      const option = [...document.querySelectorAll(".input-behavior-option")]
        .find((label) => label.querySelector("strong")?.textContent?.includes(title));
      const input = option?.querySelector('input[type="checkbox"]');
      if (!input) return { ready: false, checked: null, changed: false, title };
      const target = ${enabled};
      const changed = input.checked !== target;
      if (changed) input.click();
      return { ready: true, checked: input.checked, changed, title };
    })()`);
    assert.ok(changed.ready, `${title} is missing from the real UI`);
    await waitForEvaluation(`(() => {
      const title = ${JSON.stringify(title)};
      const option = [...document.querySelectorAll(".input-behavior-option")]
        .find((label) => label.querySelector("strong")?.textContent?.includes(title));
      const input = option?.querySelector('input[type="checkbox"]');
      return { ready: input?.checked === ${enabled}, checked: input?.checked ?? null };
    })()`, `${title} UI set to ${enabled}`);
    await evaluate(`document.querySelector(".canvas-input-behavior-trigger")?.click()`);
    await waitForEvaluation(`(() => ({
      ready: !document.querySelector(".input-behavior-popover"),
    }))()`, "input behavior popover closed");
    await focusField();
    return changed;
  };

  const setWrapperAutoExitThroughUi = (enabled) =>
    setInputBehaviorThroughUi("字体命令输入后跳出", enabled);

  const clickScriptToolbar = async (commandId) => {
    await focusField();
    await evaluate(`document.querySelector('button[data-category="structure"]')?.click()`);
    await waitForEvaluation(
      `(() => ({ ready: Boolean(document.querySelector('[data-command-id="${commandId}"]')) }))()`,
      `${commandId} toolbar button`,
    );
    await evaluate(`document.querySelector('[data-command-id="${commandId}"]')?.click()`);
    await waitForEvaluation(`(() => {
      const field = document.querySelector("math-field");
      if (!field?.value.includes("\\\\placeholder")) {
        return { ready: false, value: field?.value ?? "" };
      }
      field.focus();
      const sink = field.shadowRoot?.querySelector('[part="keyboard-sink"]');
      sink?.focus({ preventScroll: true });
      return {
        ready: field.matches(":focus-within") && field.shadowRoot?.activeElement === sink,
        value: field.value,
      };
    })()`, `${commandId} placeholder focused`);
    await sleep(100);
  };

  const runScriptUiCase = async ({
    settingTitle,
    enabled,
    commandId,
    expectedValue,
    description,
  }) => {
    await resetDocument(true, "x");
    const uiSetting = await setInputBehaviorThroughUi(settingTitle, enabled);
    await clickScriptToolbar(commandId);
    const empty = await evaluate(scriptStateExpression);
    await key("A", "KeyA", 65);
    const afterA = await evaluate(scriptStateExpression);
    await key("B", "KeyB", 66);
    const afterB = await waitForEvaluation(`(() => {
      const field = document.querySelector("math-field");
      const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      return {
        ready: field?.value === ${JSON.stringify(expectedValue)},
        value: field?.value ?? "",
        position: field?.position ?? -1,
        lastOffset: field?.lastOffset ?? -1,
        inputBehavior: persisted.state?.inputBehavior ?? null,
      };
    })()`, description);
    return { settingTitle, enabled, commandId, expectedValue, uiSetting, empty, afterA, afterB };
  };

  const scriptStateExpression = `(() => {
    const field = document.querySelector("math-field");
    const caret = field?.shadowRoot?.querySelector(".ML__caret");
    const script = caret?.closest(".ML__msubsup, .ML__op-group");
    const nucleus = script?.classList.contains("ML__msubsup")
      ? script.previousElementSibling
      : script?.querySelector(".ML__op-symbol");
    const caretRect = (caret?.parentElement ?? caret)?.getBoundingClientRect();
    const nucleusRect = nucleus?.getBoundingClientRect();
    const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
    return {
      ready: Boolean(field),
      value: field?.value ?? "",
      position: field?.position ?? -1,
      lastOffset: field?.lastOffset ?? -1,
      selection: field?.selection ?? null,
      caretCenter: caretRect ? caretRect.top + caretRect.height / 2 : null,
      nucleusCenter: nucleusRect ? nucleusRect.top + nucleusRect.height / 2 : null,
      inputBehavior: persisted.state?.inputBehavior ?? null,
    };
  })()`;

  const report = {
    endpoint,
    autoExit: {},
    continuous: {},
    middle: {},
    rawCommand: {},
    scriptAutoExit: {},
    artifactRoot,
  };

  await resetDocument(true);
  await clickMathbbToolbar();
  report.autoExit.empty = await evaluate(wrapperStateExpression);
  await screenshot("01-auto-exit-empty.png");
  await key("A", "KeyA", 65);
  report.autoExit.afterA = await evaluate(wrapperStateExpression);
  await key("B", "KeyB", 66);
  report.autoExit.afterB = await evaluate(wrapperStateExpression);
  await screenshot("02-auto-exit-after-ab.png");

  await resetDocument(true);
  report.continuous.uiSetting = await setWrapperAutoExitThroughUi(false);
  await clickMathbbToolbar();
  report.continuous.empty = await evaluate(wrapperStateExpression);
  await screenshot("03-continuous-empty.png");
  for (const [index, character] of ["A", "B", "C"].entries()) {
    await key(character, `Key${character}`, character.charCodeAt(0));
    report.continuous[`after${index + 1}`] = await evaluate(wrapperStateExpression);
    await screenshot(`04-continuous-${index + 1}.png`);
  }
  await key("Enter", "Enter", 13);
  report.continuous.afterEnter = await evaluate(wrapperStateExpression);
  await key("D", "KeyD", 68);
  report.continuous.afterD = await evaluate(wrapperStateExpression);
  await screenshot("05-continuous-after-enter-d.png");

  await resetDocument(false, "x+y");
  await clickMathbbToolbar(1);
  report.middle.empty = await evaluate(wrapperStateExpression);
  await key("A", "KeyA", 65);
  await key("B", "KeyB", 66);
  report.middle.afterAB = await evaluate(wrapperStateExpression);
  await key("Enter", "Enter", 13);
  report.middle.afterEnter = await evaluate(wrapperStateExpression);
  await key("C", "KeyC", 67);
  report.middle.afterC = await evaluate(wrapperStateExpression);
  await screenshot("06-middle-insertion.png");

  await resetDocument(true);
  await focusField();
  await typeText("\\mathbb");
  report.rawCommand.beforeConfirm = await evaluate(`(() => {
    const field = document.querySelector("math-field");
    const rawNodes = [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
      .map((node) => node.textContent ?? "");
    return {
      value: field?.value ?? "",
      mode: field?.mode ?? "",
      position: field?.position ?? -1,
      lastOffset: field?.lastOffset ?? -1,
      rawNodes,
      shadowText: field?.shadowRoot?.textContent ?? "",
    };
  })()`);
  await key(" ", "Space", 32, false);
  report.rawCommand.empty = await evaluate(wrapperStateExpression);
  await key("A", "KeyA", 65);
  report.rawCommand.afterA = await evaluate(wrapperStateExpression);
  await key("B", "KeyB", 66);
  report.rawCommand.afterB = await evaluate(wrapperStateExpression);
  await screenshot("07-raw-command-auto-exit.png");

  report.scriptAutoExit.superscriptOffSubscriptOnSuperscript = await runScriptUiCase({
    settingTitle: "上标输入后跳出",
    enabled: false,
    commandId: "power",
    expectedValue: "x^{AB}",
    description: "superscript remains open when only superscript auto-exit is disabled",
  });
  report.scriptAutoExit.superscriptOffSubscriptOnSubscript = await runScriptUiCase({
    settingTitle: "上标输入后跳出",
    enabled: false,
    commandId: "subscript",
    expectedValue: "x_{A}B",
    description: "subscript still exits when only superscript auto-exit is disabled",
  });
  report.scriptAutoExit.superscriptOnSubscriptOffSuperscript = await runScriptUiCase({
    settingTitle: "下标输入后跳出",
    enabled: false,
    commandId: "power",
    expectedValue: "x^{A}B",
    description: "superscript still exits when only subscript auto-exit is disabled",
  });
  report.scriptAutoExit.superscriptOnSubscriptOffSubscript = await runScriptUiCase({
    settingTitle: "下标输入后跳出",
    enabled: false,
    commandId: "subscript",
    expectedValue: "x_{AB}",
    description: "subscript remains open when only subscript auto-exit is disabled",
  });
  await screenshot("08-independent-script-auto-exit.png");

  console.log(JSON.stringify(report, null, 2));

  assert.ok(report.autoExit.empty.hasBox, "auto-exit toolbar insertion has no visible wrapper input box");
  assert.equal(report.autoExit.empty.pending, "\\mathbb");
  assert.equal(report.autoExit.afterA.value, "\\mathbb{A}");
  assert.ok(!report.autoExit.afterA.hasBox, "auto-exit wrapper box did not close after one character");
  assert.equal(report.autoExit.afterB.value, "\\mathbb{A}B");

  assert.ok(report.continuous.empty.hasBox, "continuous toolbar insertion has no visible wrapper input box");
  assert.equal(report.continuous.after3.value, "\\mathbb{ABC}");
  assert.ok(report.continuous.after3.hasBox, "continuous wrapper input closed before Enter");
  assert.ok(
    report.continuous.after3.boxWidth > report.continuous.after1.boxWidth,
    `continuous wrapper box did not grow: ${report.continuous.after1.boxWidth} -> ${report.continuous.after3.boxWidth}`,
  );
  assert.ok(!report.continuous.afterEnter.hasBox, "Enter did not close continuous wrapper input");
  assert.equal(report.continuous.afterD.value, "\\mathbb{ABC}D");
  assert.equal(report.continuous.uiSetting.checked, false);

  assert.ok(report.middle.empty.hasBox, "middle insertion did not open a wrapper input box");
  assert.equal(report.middle.afterAB.value, "x\\mathbb{AB}+y");
  assert.ok(report.middle.afterAB.hasBox, "middle wrapper closed before Enter");
  assert.ok(!report.middle.afterEnter.hasBox, "middle wrapper did not close on Enter");
  assert.equal(report.middle.afterC.value, "x\\mathbb{AB}C+y");

  assert.ok(report.rawCommand.empty.hasBox, "typed raw wrapper command did not open a box");
  assert.equal(report.rawCommand.empty.pending, "\\mathbb");
  assert.equal(report.rawCommand.afterA.value, "\\mathbb{A}");
  assert.ok(!report.rawCommand.afterA.hasBox, "typed raw command did not auto-exit after one character");
  assert.equal(report.rawCommand.afterB.value, "\\mathbb{A}B");

  assert.equal(
    report.scriptAutoExit.superscriptOffSubscriptOnSuperscript.afterB.value,
    "x^{AB}",
  );
  assert.equal(
    report.scriptAutoExit.superscriptOffSubscriptOnSubscript.afterB.value,
    "x_{A}B",
  );
  assert.equal(
    report.scriptAutoExit.superscriptOnSubscriptOffSuperscript.afterB.value,
    "x^{A}B",
  );
  assert.equal(
    report.scriptAutoExit.superscriptOnSubscriptOffSubscript.afterB.value,
    "x_{AB}",
  );
  assert.equal(
    report.scriptAutoExit.superscriptOffSubscriptOnSuperscript.afterB.inputBehavior.autoExitSuperscript,
    false,
  );
  assert.equal(
    report.scriptAutoExit.superscriptOffSubscriptOnSuperscript.afterB.inputBehavior.autoExitSubscript,
    true,
  );
  assert.equal(
    report.scriptAutoExit.superscriptOnSubscriptOffSubscript.afterB.inputBehavior.autoExitSuperscript,
    true,
  );
  assert.equal(
    report.scriptAutoExit.superscriptOnSubscriptOffSubscript.afterB.inputBehavior.autoExitSubscript,
    false,
  );

  client.close();
  console.log("Real Windows Tauri/WebView2 wrapper acceptance passed.");
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
