import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const portOffset = process.pid % 1000;
const previewPort = 7600 + portOffset;
const debugPort = 12600 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}/editor`;
const chromeProfile = `/tmp/visualtex-raw-anchor-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while the process starts.
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
        "--window-size=1400,900",
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
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(600);

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

    const reload = async () => {
      await client.send("Page.reload", { ignoreCache: true });
      await sleep(550);
      await evaluate(`new Promise((resolve) => {
        const poll = () => document.querySelector("math-field")
          ? resolve(true)
          : setTimeout(poll, 25);
        poll();
      })`);
    };

    const configure = async (overrides) => {
      await evaluate(`(() => {
        const key = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(key) || "{}");
        persisted.state = {
          ...(persisted.state || {}),
          inputBehavior: {
            ...(persisted.state?.inputBehavior || {}),
            ...${JSON.stringify(overrides)},
          },
        };
        localStorage.setItem(key, JSON.stringify(persisted));
      })()`);
      await reload();
    };

    await evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
      localStorage.setItem("visualtex.onboarding.macos.desktop.v1.2.0.completed", "true");
      localStorage.setItem("visualtex.office.macos.native-first-run.v1.2.0.completed", "true");
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
          showOtherCommandSuggestions: true,
        },
      };
      localStorage.setItem(key, JSON.stringify(persisted));
    })()`);
    await reload();

    const typeCharacter = async (key, code, keyCode, pause = 105) => {
      const common = {
        key,
        code,
        windowsVirtualKeyCode: keyCode,
        nativeVirtualKeyCode: keyCode,
      };
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        ...common,
        text: key,
        unmodifiedText: key,
      });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(pause);
    };

    const pressKey = async (key, code, keyCode, pause = 140) => {
      const common = {
        key,
        code,
        windowsVirtualKeyCode: keyCode,
        nativeVirtualKeyCode: keyCode,
      };
      await client.send("Input.dispatchKeyEvent", {
        type: "rawKeyDown",
        ...common,
      });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(pause);
    };

    const typeRawCommand = async (command) => {
      await typeCharacter("\\", "Backslash", 220);
      for (const character of command) {
        const upper = character.toUpperCase();
        await typeCharacter(character, `Key${upper}`, upper.charCodeAt(0));
      }
      const beforeSpace = await readState();
      assert.equal(
        beforeSpace.raw,
        `\\${command}`,
        `Raw command changed before completion: ${JSON.stringify(beforeSpace)}`,
      );
      if (beforeSpace.selectedCandidate) {
        assert.equal(beforeSpace.selectedCandidate, `\\${command}`);
      }
      await typeCharacter(" ", "Space", 32, 170);
      return readState();
    };

    const preparePlaceholder = async (latex) => {
      const prepared = await evaluate(`(() => {
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
          selected: field.getValue(field.selection),
        };
      })()`);
      assert.match(
        prepared.selected,
        /\\placeholder\{\}/,
        `Placeholder was not selected: ${latex} ${JSON.stringify(prepared)}`,
      );
    };

    const prepareVerticalPlaceholder = async (latex, region = "upper") => {
      const prepared = await evaluate(`(async () => {
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
        await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        const placeholderCount = (field.value.match(/\\\\placeholder\\{\\}/g) || []).length;
        field.position = field.lastOffset;
        const candidates = [];
        for (let index = 0; index < placeholderCount; index += 1) {
          field.executeCommand('moveToPreviousPlaceholder');
          await new Promise((resolve) => requestAnimationFrame(resolve));
          const selection = JSON.parse(JSON.stringify(field.selection));
          const selected = field.selectionIsCollapsed ? '' : field.getValue(field.selection);
          const marker = Array.from(field.shadowRoot?.querySelectorAll(
            '.visualtex-structural-placeholder-caret, .ML__placeholder-selected, .ML__selected, .ML__caret'
          ) ?? []).find((node) => node.getBoundingClientRect().height > 0);
          const markerRect = (marker?.parentElement ?? marker)?.getBoundingClientRect();
          const placeholderNodes = Array.from(
            field.shadowRoot?.querySelectorAll('.visualtex-structural-placeholder, .ML__placeholder') ?? []
          ).map((node) => {
            const rect = node.getBoundingClientRect();
            return {
              node,
              centerX: rect.left + rect.width / 2,
              centerY: rect.top + rect.height / 2,
              width: rect.width,
              height: rect.height,
            };
          }).filter((item) => item.width > 0 && item.height > 0);
          const markerX = markerRect ? markerRect.left + markerRect.width / 2 : 0;
          const markerY = markerRect ? markerRect.top + markerRect.height / 2 : 0;
          const closest = placeholderNodes.sort((first, second) =>
            Math.hypot(first.centerX - markerX, first.centerY - markerY) -
            Math.hypot(second.centerX - markerX, second.centerY - markerY)
          )[0] ?? null;
          candidates.push({
            selection,
            selected,
            centerY: closest?.centerY ?? Number.POSITIVE_INFINITY,
          });
        }
        candidates.sort((first, second) => first.centerY - second.centerY);
        const target = ${JSON.stringify(region)} === 'upper'
          ? candidates[0]
          : candidates[candidates.length - 1];
        if (target) field.selection = target.selection;
        return {
          value: field.value,
          selection: field.selection,
          selected: field.selectionIsCollapsed ? '' : field.getValue(field.selection),
          placeholders: candidates,
          region: ${JSON.stringify(region)},
        };
      })()`);
      assert.match(
        prepared.selected,
        /\\placeholder\{\}/,
        `Vertical placeholder was not selected: ${latex} ${JSON.stringify(prepared)}`,
      );
      return prepared;
    };

    const readState = () =>
      evaluate(`(() => {
        const field = document.querySelector("math-field");
        const raw = Array.from(field.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])
          .filter((node) => !node.classList.contains("ML__suggestion"))
          .map((node) => node.textContent ?? "")
          .join("");
        const markers = Array.from(field.shadowRoot?.querySelectorAll(
          ".ML__caret, .ML__text-caret, .ML__latex-caret, .ML__placeholder-selected, .visualtex-structural-placeholder-caret"
        ) ?? []).filter((node) => {
          const rect = node.getBoundingClientRect();
          const style = getComputedStyle(node);
          return rect.height > 0 && style.display !== "none" && style.visibility !== "hidden";
        }).sort((first, second) => {
          const priority = (node) =>
            node.classList.contains("ML__placeholder-selected") ? 0 :
            node.classList.contains("ML__caret") || node.classList.contains("ML__text-caret") ? 1 :
            node.classList.contains("ML__latex-caret") ? 2 : 3;
          return priority(first) - priority(second);
        });
        const marker = markers[0] ?? null;
        const markerRect = marker?.getBoundingClientRect();
        const placeholderNodes = Array.from(field.shadowRoot?.querySelectorAll(
          ".ML__placeholder, .visualtex-structural-placeholder, [data-atom-id]"
        ) ?? []).filter((node) => {
          const text = (node.textContent || "").trim();
          return node.classList.contains("ML__placeholder") ||
            node.classList.contains("visualtex-structural-placeholder") ||
            text === (field.placeholderSymbol || "▢");
        });
        const structuralContextNode =
          marker?.classList.contains("visualtex-structural-placeholder-caret") && markerRect
            ? placeholderNodes.sort((first, second) => {
                const distance = (node) => {
                  const rect = node.getBoundingClientRect();
                  return Math.hypot(
                    rect.left + rect.width / 2 - (markerRect.left + markerRect.width / 2),
                    rect.top + rect.height / 2 - (markerRect.top + markerRect.height / 2),
                  );
                };
                return distance(first) - distance(second);
              })[0] ?? marker
            : marker;
        const countAncestors = (className) => {
          let count = 0;
          let current = structuralContextNode;
          while (current) {
            if (current.classList?.contains(className)) count += 1;
            current = current.parentElement;
          }
          return count;
        };
        return {
          value: field.value,
          raw,
          selectedCandidate:
            field.dataset.pendingNativeSuggestion ||
            document.querySelector("#mathlive-suggestion-popover li.ML__popover__current")?.dataset.command ||
            "",
          position: field.position,
          lastOffset: field.lastOffset,
          selection: field.selection,
          selectedLatex: field.selectionIsCollapsed ? "" : field.getValue(field.selection),
          accentDepth: countAncestors("ML__accent-body"),
          scriptDepth: countAncestors("ML__msubsup"),
          operatorDepth: countAncestors("ML__op-group"),
          fractionDepth: countAncestors("ML__mfrac"),
          sqrtDepth: countAncestors("ML__sqrt"),
          markerClass: marker?.className || "",
          prefix: field.getValue(0, field.position, "latex"),
          pendingWrapperCommand:
            field.closest(".mathfield-host")?.dataset.pendingWrapperCommand || "",
          pendingWrapperLength:
            field.closest(".mathfield-host")?.dataset.pendingWrapperLength || "",
        };
      })()`);

    const countCommand = (value, command) =>
      value.split(`\\${command}`).length - 1;

    const scalarCases = [
      {
        label: "accent pi auto-exit",
        latex: "\\hat{\\placeholder{}}+z",
        command: "pi",
        expected: /^\\hat\{\\pi\}\+z$/,
        context: { accentDepth: 0 },
      },
      {
        label: "hat alpha keeps exactly one accent",
        latex: "\\hat{\\placeholder{}}+z",
        command: "alpha",
        expected: /^\\hat\{\\alpha\}\+z$/,
        context: { accentDepth: 0 },
        outerCommand: "hat",
      },
      {
        label: "accent alpha auto-exit",
        latex: "\\vec{\\placeholder{}}+z",
        command: "alpha",
        expected: /^\\vec\{\\alpha\}\+z$/,
        context: { accentDepth: 0 },
        outerCommand: "vec",
      },
      {
        label: "dot alpha keeps exact source",
        latex: "\\dot{\\placeholder{}}",
        command: "alpha",
        expected: /^\\dot\{\\alpha\}$/,
        context: { accentDepth: 0 },
        outerCommand: "dot",
      },
      {
        label: "dot theta keeps exactly one accent",
        latex: "\\dot{\\placeholder{}}+z",
        command: "theta",
        expected: /^\\dot\{\\theta\}\+z$/,
        context: { accentDepth: 0 },
        outerCommand: "dot",
      },
      {
        label: "accent integral auto-exit",
        latex: "\\bar{\\placeholder{}}+z",
        command: "int",
        expected: /^\\bar\{\\int\}\+z$/,
        context: { accentDepth: 0 },
      },
      {
        label: "superscript pi auto-exit",
        latex: "x^{\\placeholder{}}+z",
        command: "pi",
        expected: /^x\^\{?\\pi\}?\+z$/,
        context: { scriptDepth: 0 },
      },
      {
        label: "subscript pi auto-exit",
        latex: "x_{\\placeholder{}}+z",
        command: "pi",
        expected: /^x_\{?\\pi\}?\+z$/,
        context: { scriptDepth: 0 },
      },
      {
        label: "accent inside fraction stays in numerator",
        latex: "\\frac{\\hat{\\placeholder{}}}{b}+z",
        command: "pi",
        expected: /^\\frac\{\\hat\{\\pi\}\}\{b\}\+z$/,
        context: { accentDepth: 0, fractionDepth: 1 },
      },
      {
        label: "operator limit does not auto-exit",
        latex: "\\int_{0}^{\\placeholder{}}f",
        command: "pi",
        expected: /^\\int_(?:0|\{0\})\^\{\\pi\}f$/,
        context: { operatorDepth: 1 },
      },
      {
        label: "nested accent exits only the inner accent",
        latex: "\\hat{\\vec{\\placeholder{}}}+z",
        command: "pi",
        expected: /^\\hat\{\\vec\{\\pi\}\}\+z$/,
        context: { accentDepth: 0 },
      },
    ];

    for (const testCase of scalarCases) {
      await preparePlaceholder(testCase.latex);
      await typeCharacter("a", "KeyA", 65, 170);
      const directBaseline = await readState();

      await preparePlaceholder(testCase.latex);
      const state = await typeRawCommand(testCase.command);
      assert.match(
        state.value,
        testCase.expected,
        `${testCase.label}: ${JSON.stringify(state)}`,
      );
      assert.equal(
        countCommand(state.value, testCase.command),
        1,
        `${testCase.label} duplicated command: ${JSON.stringify(state)}`,
      );
      if (testCase.outerCommand) {
        assert.equal(
          countCommand(state.value, testCase.outerCommand),
          1,
          `${testCase.label} created an extra empty accent: ${JSON.stringify(state)}`,
        );
        assert.doesNotMatch(
          state.value,
          new RegExp(`\\\\${testCase.outerCommand}\\{\\}`),
          `${testCase.label} left an empty accent: ${JSON.stringify(state)}`,
        );
      }
      for (const key of [
        "accentDepth",
        "scriptDepth",
        "operatorDepth",
        "fractionDepth",
        "sqrtDepth",
      ]) {
        assert.equal(
          state[key],
          directBaseline[key],
          `${testCase.label} raw command caret differs from ordinary input (${key}): baseline=${JSON.stringify(directBaseline)} raw=${JSON.stringify(state)}`,
        );
      }
      for (const [key, expected] of Object.entries(testCase.context)) {
        assert.equal(
          state[key],
          expected,
          `${testCase.label} wrong caret context ${key}: ${JSON.stringify(state)}`,
        );
      }
    }

    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      field.setValue("", {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
      field.position = 0;
    })()`);
    const insertedDotTemplate = await typeRawCommand("dot");
    assert.match(
      insertedDotTemplate.value,
      /^\\dot\{\\placeholder\{\}\}$/,
      `typing \\dot did not create one selected placeholder: ${JSON.stringify(insertedDotTemplate)}`,
    );
    assert.match(
      insertedDotTemplate.selectedLatex,
      /\\placeholder\{\}/,
      `typing \\dot did not select its placeholder: ${JSON.stringify(insertedDotTemplate)}`,
    );
    const dotAlphaFromTypedTemplate = await typeRawCommand("alpha");
    assert.equal(
      dotAlphaFromTypedTemplate.value,
      "\\dot{\\alpha}",
      `typing \\alpha inside a typed \\dot duplicated the accent: ${JSON.stringify(dotAlphaFromTypedTemplate)}`,
    );
    assert.equal(
      countCommand(dotAlphaFromTypedTemplate.value, "dot"),
      1,
      `typing \\alpha inside a typed \\dot created another dot: ${JSON.stringify(dotAlphaFromTypedTemplate)}`,
    );

    const verticalPlaceholderCases = [
      {
        label: "fraction numerator keeps raw command",
        latex: "\\frac{\\placeholder{}}{\\placeholder{}}+z",
        expected: /^\\frac\{\\pi\}\{\\placeholder\{\}\}\+z$/,
      },
      {
        label: "integral upper limit keeps raw command",
        latex: "\\int_{\\placeholder{}}^{\\placeholder{}}f",
        expected: /^\\int_\{\\placeholder\{\}\}\^\{\\pi\}f$/,
      },
      {
        label: "sum upper limit keeps raw command",
        latex: "\\sum_{\\placeholder{}}^{\\placeholder{}}x",
        expected: /^\\sum_\{\\placeholder\{\}\}\^\{\\pi\}x$/,
      },
      {
        label: "overset upper slot keeps raw command",
        latex: "\\overset{\\placeholder{}}{\\placeholder{}}+z",
        expected: /^\\overset\{\\pi\}\{\\placeholder\{\}\}\+z$/,
      },
    ];

    for (const testCase of verticalPlaceholderCases) {
      const prepared = await prepareVerticalPlaceholder(testCase.latex, "upper");
      assert.ok(
        prepared.placeholders.length >= 2,
        `${testCase.label} did not expose two placeholders: ${JSON.stringify(prepared)}`,
      );
      const state = await typeRawCommand("pi");
      assert.match(
        state.value,
        testCase.expected,
        `${testCase.label}: ${JSON.stringify({ prepared, state })}`,
      );
    }

    const wrapperCases = [
      {
        label: "mathbb inside accent auto-exit",
        latex: "\\hat{\\placeholder{}}+z",
        command: "mathbb",
        input: "R",
        expected: /^\\hat\{\\mathbb\{R\}\}\+z$/,
        afterFollowup: /^\\hat\{\\mathbb\{R\}\}q\+z$/,
        context: { accentDepth: 0 },
      },
      {
        label: "mathbf inside accent auto-exit",
        latex: "\\vec{\\placeholder{}}+z",
        command: "mathbf",
        input: "x",
        expected: /^\\vec\{\\mathbf\{x\}\}\+z$/,
        afterFollowup: /^\\vec\{\\mathbf\{x\}\}q\+z$/,
        context: { accentDepth: 0 },
      },
      {
        label: "mathcal inside accent auto-exit",
        latex: "\\bar{\\placeholder{}}+z",
        command: "mathcal",
        input: "A",
        expected: /^\\bar\{\\mathcal\{A\}\}\+z$/,
        afterFollowup: /^\\bar\{\\mathcal\{A\}\}q\+z$/,
        context: { accentDepth: 0 },
      },
      {
        label: "mathfrak inside nested accent and fraction",
        latex: "\\frac{\\hat{\\placeholder{}}}{b}+z",
        command: "mathfrak",
        input: "g",
        expected: /^\\frac\{\\hat\{\\mathfrak\{g\}\}\}\{b\}\+z$/,
        afterFollowup: /^\\frac\{\\hat\{\\mathfrak\{g\}\}q\}\{b\}\+z$/,
        context: { accentDepth: 0, fractionDepth: 1 },
      },
      {
        label: "mathbb inside radical stays in radical",
        latex: "\\sqrt{\\placeholder{}}+z",
        command: "mathbb",
        input: "R",
        expected: /^\\sqrt\{\\mathbb\{R\}\}\+z$/,
        afterFollowup: /^\\sqrt\{\\mathbb\{R\}q\}\+z$/,
        context: {},
      },
    ];

    for (const testCase of wrapperCases) {
      await preparePlaceholder(testCase.latex);
      await typeRawCommand(testCase.command);
      const pending = await readState();
      assert.equal(
        pending.pendingWrapperCommand,
        `\\${testCase.command}`,
        `${testCase.label} did not enter wrapper input: ${JSON.stringify(pending)}`,
      );
      await typeCharacter(
        testCase.input,
        `Key${testCase.input.toUpperCase()}`,
        testCase.input.toUpperCase().charCodeAt(0),
        170,
      );
      const state = await readState();
      assert.match(state.value, testCase.expected, `${testCase.label}: ${JSON.stringify(state)}`);
      assert.equal(
        state.pendingWrapperCommand,
        "",
        `${testCase.label} wrapper input did not complete: ${JSON.stringify(state)}`,
      );
      for (const [key, expected] of Object.entries(testCase.context)) {
        assert.equal(
          state[key],
          expected,
          `${testCase.label} wrong caret context ${key}: ${JSON.stringify(state)}`,
        );
      }
      await typeCharacter("q", "KeyQ", 81, 170);
      const followup = await readState();
      assert.match(
        followup.value,
        testCase.afterFollowup,
        `${testCase.label} did not leave the accent before follow-up input: ${JSON.stringify(followup)}`,
      );
    }

    const verticalWrapperCases = [
      {
        label: "mathbb stays in fraction numerator",
        latex: "\\frac{\\placeholder{}}{\\placeholder{}}+z",
        command: "mathbb",
        input: "R",
        expected: /^\\frac\{\\mathbb\{R\}\}\{\\placeholder\{\}\}\+z$/,
        afterFollowup: /^\\frac\{\\mathbb\{R\}q\}\{\\placeholder\{\}\}\+z$/,
      },
      {
        label: "mathbf stays in sum upper limit",
        latex: "\\sum_{\\placeholder{}}^{\\placeholder{}}f",
        command: "mathbf",
        input: "x",
        expected: /^\\sum_\{\\placeholder\{\}\}\^\{\\mathbf\{x\}\}f$/,
        afterFollowup: /^\\sum_\{\\placeholder\{\}\}\^\{\\mathbf\{x\}q\}f$/,
      },
      {
        label: "mathcal stays in overset upper slot",
        latex: "\\overset{\\placeholder{}}{\\placeholder{}}+z",
        command: "mathcal",
        input: "A",
        expected: /^\\overset\{\\mathcal\{A\}\}\{\\placeholder\{\}\}\+z$/,
        afterFollowup: /^\\overset\{\\mathcal\{A\}q\}\{\\placeholder\{\}\}\+z$/,
      },
    ];

    for (const testCase of verticalWrapperCases) {
      await prepareVerticalPlaceholder(testCase.latex, "upper");
      await typeRawCommand(testCase.command);
      const pending = await readState();
      assert.equal(
        pending.pendingWrapperCommand,
        `\\${testCase.command}`,
        `${testCase.label} did not enter wrapper input: ${JSON.stringify(pending)}`,
      );
      await typeCharacter(
        testCase.input,
        `Key${testCase.input.toUpperCase()}`,
        testCase.input.toUpperCase().charCodeAt(0),
        170,
      );
      const state = await readState();
      assert.match(
        state.value,
        testCase.expected,
        `${testCase.label}: ${JSON.stringify(state)}`,
      );
      await typeCharacter("q", "KeyQ", 81, 170);
      const followup = await readState();
      assert.match(
        followup.value,
        testCase.afterFollowup,
        `${testCase.label} moved to another branch: ${JSON.stringify(followup)}`,
      );
    }

    await configure({
      autoExitAccent: false,
      autoExitWrapperCommand: true,
      autoExitSuperscript: true,
      autoExitSubscript: true,
    });
    await preparePlaceholder("\\hat{\\placeholder{}}+z");
    await typeRawCommand("mathbb");
    await typeCharacter("R", "KeyR", 82, 170);
    await typeCharacter("q", "KeyQ", 81, 170);
    const disabledAccentWrapper = await readState();
    assert.equal(
      disabledAccentWrapper.value,
      "\\hat{\\mathbb{R}q}+z",
      `disabled accent auto-exit changed wrapper behavior: ${JSON.stringify(disabledAccentWrapper)}`,
    );

    await configure({
      autoExitAccent: true,
      autoExitWrapperCommand: false,
    });
    await preparePlaceholder("\\vec{\\placeholder{}}+z");
    await typeRawCommand("mathbf");
    await typeCharacter("x", "KeyX", 88, 120);
    await typeCharacter("y", "KeyY", 89, 120);
    const pendingContinuousWrapper = await readState();
    assert.equal(pendingContinuousWrapper.value, "\\vec{\\mathbf{xy}}+z");
    assert.equal(pendingContinuousWrapper.pendingWrapperCommand, "\\mathbf");
    await pressKey("Enter", "Enter", 13, 170);
    const confirmedContinuousWrapper = await readState();
    assert.equal(confirmedContinuousWrapper.pendingWrapperCommand, "");
    await typeCharacter("q", "KeyQ", 81, 170);
    const confirmedAccentExit = await readState();
    assert.equal(
      confirmedAccentExit.value,
      "\\vec{\\mathbf{xy}}q+z",
      `Enter-confirmed wrapper did not exit accent: ${JSON.stringify(confirmedAccentExit)}`,
    );

    await configure({
      autoExitAccent: false,
      autoExitWrapperCommand: false,
    });
    await preparePlaceholder("\\bar{\\placeholder{}}+z");
    await typeRawCommand("mathfrak");
    await typeCharacter("g", "KeyG", 71, 120);
    await typeCharacter("h", "KeyH", 72, 120);
    await pressKey("Enter", "Enter", 13, 170);
    await typeCharacter("q", "KeyQ", 81, 170);
    const bothDisabledWrapper = await readState();
    assert.equal(
      bothDisabledWrapper.value,
      "\\bar{\\mathfrak{gh}q}+z",
      `disabled wrapper/accent settings were ignored: ${JSON.stringify(bothDisabledWrapper)}`,
    );

    await configure({
      autoExitAccent: true,
      autoExitWrapperCommand: true,
      autoExitSuperscript: true,
      autoExitSubscript: true,
    });
    await preparePlaceholder("x^{\\placeholder{}}+z");
    await typeRawCommand("mathbb");
    await typeCharacter("R", "KeyR", 82, 170);
    await typeCharacter("q", "KeyQ", 81, 170);
    const wrapperSuperscriptExit = await readState();
    assert.match(
      wrapperSuperscriptExit.value,
      /^x\^\{?\\mathbb\{R\}\}?q\+z$/,
      `wrapper did not preserve superscript auto-exit: ${JSON.stringify(wrapperSuperscriptExit)}`,
    );

    await preparePlaceholder("\\int_{0}^{\\placeholder{}}f");
    await typeRawCommand("mathbf");
    await typeCharacter("x", "KeyX", 88, 170);
    await typeCharacter("q", "KeyQ", 81, 170);
    const wrapperOperatorLimit = await readState();
    assert.match(
      wrapperOperatorLimit.value,
      /^\\int_(?:0|\{0\})\^\{\\mathbf\{x\}q\}f$/,
      `wrapper incorrectly exited an operator limit: ${JSON.stringify(wrapperOperatorLimit)}`,
    );

    await preparePlaceholder("\\hat{\\placeholder{}}+z");
    await typeRawCommand("mathbb");
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.dispatchEvent(new InputEvent("beforeinput", {
        inputType: "insertText",
        data: "Rq",
        bubbles: true,
        composed: true,
        cancelable: true,
      }));
    })()`);
    await sleep(170);
    const batchedWrapperInput = await readState();
    assert.equal(
      batchedWrapperInput.value,
      "\\hat{\\mathbb{R}}q+z",
      `batched wrapper input left trailing text in accent: ${JSON.stringify(batchedWrapperInput)}`,
    );

    const structuralCases = [
      {
        label: "sqrt inside accent",
        latex: "\\hat{\\placeholder{}}+z",
        command: "sqrt",
        expected: /^\\hat\{\\sqrt\{(?:\\placeholder\{\})?\}\}\+z$/,
        commandCounts: { hat: 1, sqrt: 1 },
      },
      {
        label: "fraction inside accent",
        latex: "\\vec{\\placeholder{}}+z",
        command: "frac",
        expected: /^\\vec\{\\frac\{(?:\\placeholder\{\})?\}\{(?:\\placeholder\{\})?\}\}\+z$/,
        commandCounts: { vec: 1, frac: 1 },
      },
      {
        label: "nested accent command",
        latex: "\\hat{\\placeholder{}}+z",
        command: "vec",
        expected: /^\\hat\{\\vec\{(?:\\placeholder\{\})?\}\}\+z$/,
        commandCounts: { hat: 1, vec: 1 },
      },
      {
        label: "fraction inside superscript",
        latex: "x^{\\placeholder{}}+z",
        command: "frac",
        expected: /^x\^\{\\frac\{(?:\\placeholder\{\})?\}\{(?:\\placeholder\{\})?\}\}\+z$/,
        commandCounts: { frac: 1 },
      },
    ];

    for (const testCase of structuralCases) {
      await preparePlaceholder(testCase.latex);
      const state = await typeRawCommand(testCase.command);
      assert.match(
        state.value,
        testCase.expected,
        `${testCase.label}: ${JSON.stringify(state)}`,
      );
      assert.match(
        state.selectedLatex,
        /\\placeholder\{\}/,
        `${testCase.label} did not select the inner placeholder: ${JSON.stringify(state)}`,
      );
      for (const [command, expectedCount] of Object.entries(testCase.commandCounts)) {
        assert.equal(
          countCommand(state.value, command),
          expectedCount,
          `${testCase.label} duplicated ${command}: ${JSON.stringify(state)}`,
        );
      }
    }

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
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
    })()`);
    await typeCharacter("\\", "Backslash", 220);
    for (const character of "bet") {
      const upper = character.toUpperCase();
      await typeCharacter(character, `Key${upper}`, upper.charCodeAt(0));
    }
    const prefixState = await readState();
    assert.equal(prefixState.raw, "\\bet", JSON.stringify(prefixState));
    assert.equal(
      prefixState.selectedCandidate,
      "\\beta",
      `Default native beta candidate was not selected: ${JSON.stringify(prefixState)}`,
    );
    await typeCharacter(" ", "Space", 32, 170);
    const completedPrefixState = await readState();
    assert.equal(
      completedPrefixState.value.replaceAll(" ", ""),
      "\\beta",
      `Space inserted the raw prefix instead of beta: ${JSON.stringify(completedPrefixState)}`,
    );
    assert.equal(completedPrefixState.raw, "");

    console.log("Raw command anchor and auto-exit regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(250);
    await rm(chromeProfile, {
      recursive: true,
      force: true,
      maxRetries: 6,
      retryDelay: 120,
    });
  }
}

await main();
