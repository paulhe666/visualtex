import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const portOffset = process.pid % 800;
const previewPort = 7700 + portOffset;
const debugPort = 12700 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = createBrowserProfilePath("visualtex-windows-cases");
const chromePath = resolveChromiumExecutable();
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while local processes start.
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
    if (!page) throw new Error("No VisualTeX Chrome page target found");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    // Chromium may expose the page target before its initial navigation has
    // settled. Navigate explicitly so subsequent Runtime.evaluate calls keep a
    // stable same-origin execution context.
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(700);

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

    const typeKey = async (key, code, text = key, modifiers = 0) => {
      const specialVirtualKey =
        key === "Enter"
          ? 13
          : key === "Tab"
            ? 9
            : key === "Escape"
              ? 27
              : key === " "
                ? 32
                : key === "\\"
                  ? 220
                  : 0;
      const virtualKey =
        specialVirtualKey ||
        (key.length === 1 ? key.toUpperCase().charCodeAt(0) : 0);
      const common = {
        key,
        code,
        modifiers,
        windowsVirtualKeyCode: virtualKey,
        nativeVirtualKeyCode: virtualKey,
      };
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        ...common,
        text,
        unmodifiedText: text,
      });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(35);
    };

    await evaluate(`new Promise((resolve) => {
      const done = () => document.querySelector("math-field") ? resolve(true) : setTimeout(done, 30);
      done();
    })`);

    const casesLatex = String.raw`\begin{cases}x & x>0 \\ 0 & x\le 0\end{cases}`;
    const environmentProbe = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.setValue(${JSON.stringify(casesLatex)}, {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
      });
      const positions = [];
      for (let position = 0; position <= field.lastOffset; position += 1) {
        field.position = position;
        positions.push({
          position,
          environment: field._mathfield?.model?.parentEnvironment?.environmentName ?? null,
        });
      }
      return { value: field.value, positions };
    })()`);
    assert.match(environmentProbe.value, /\\begin\{cases\}/);
    assert.ok(
      environmentProbe.positions.some((entry) => entry.environment === "cases"),
      `MathLive did not expose the active cases environment: ${JSON.stringify(environmentProbe.positions)}`,
    );

    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.dispatchEvent(new KeyboardEvent("keydown", {
        key: "Escape",
        code: "Escape",
        bubbles: true,
        composed: true,
      }));
      field.executeCommand(["complete", "reject"]);
      field.mode = "math";
      field.setValue("", {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
      field.position = field.lastOffset;
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus();
      return true;
    })()`);
    // Match the production editor's delayed focus repair window before sending
    // a physical Backslash key, otherwise the 0/80 ms focus callbacks can turn
    // the command introducer into ordinary math text in headless Chromium.
    await sleep(140);

    await typeKey("\\", "Backslash", "\\");
    for (const letter of "beg") {
      await typeKey(letter, `Key${letter.toUpperCase()}`);
    }
    await sleep(120);

    const begSuggestion = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      const panel = document.querySelector('#visualtex-native-input-suggestion-popover');
      return {
        raw: Array.from(field.shadowRoot?.querySelectorAll('.ML__raw-latex') ?? [])
          .filter((node) => !node.classList.contains('ML__suggestion'))
          .map((node) => node.textContent ?? "")
          .join(""),
        visible: panel?.classList.contains('is-visible') ?? false,
        command: panel?.querySelector('li.ML__popover__current')?.dataset.command ?? '',
      };
    })()`);
    assert.equal(begSuggestion.raw.replace(/\s+/g, ""), "\\beg");
    assert.equal(
      begSuggestion.visible,
      true,
      "cases suggestion did not appear at the early \\beg prefix",
    );
    assert.equal(begSuggestion.command, "\\begin{cases}");

    for (const letter of "in") {
      await typeKey(letter, `Key${letter.toUpperCase()}`);
    }
    await sleep(120);

    const casesSuggestion = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      const panel = document.querySelector('#visualtex-native-input-suggestion-popover');
      const current = panel?.querySelector('li.ML__popover__current');
      const preview = current?.querySelector('.ML__popover__command');
      const style = panel ? getComputedStyle(panel) : null;
      return {
        value: field.value,
        raw: Array.from(field.shadowRoot?.querySelectorAll('.ML__raw-latex') ?? [])
          .filter((node) => !node.classList.contains('ML__suggestion'))
          .map((node) => node.textContent ?? "")
          .join(""),
        visible: panel?.classList.contains('is-visible') ?? false,
        command: current?.dataset.command ?? '',
        previewText: preview?.textContent?.trim() ?? '',
        previewHtml: preview?.innerHTML ?? '',
        background: style?.backgroundColor ?? '',
        borderRadius: style?.borderRadius ?? '',
        width: panel?.getBoundingClientRect().width ?? 0,
      };
    })()`);
    assert.equal(casesSuggestion.raw.replace(/\s+/g, ""), "\\begin");
    assert.equal(
      casesSuggestion.visible,
      true,
      "cases did not show the VisualTeX native-style MathLive suggestion popover",
    );
    assert.equal(casesSuggestion.command, "\\begin{cases}");
    assert.ok(casesSuggestion.previewHtml.length > 0, "cases suggestion preview is empty");
    assert.ok(casesSuggestion.width >= 240, `cases suggestion popover is unexpectedly narrow: ${casesSuggestion.width}`);
    assert.doesNotMatch(
      casesSuggestion.value,
      /\\begin\{cases\}/,
      "cases committed before the user accepted the native suggestion",
    );

    await typeKey(" ", "Space", " ");
    await sleep(100);
    const typedCases = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        value: field.value,
        raw: Array.from(field.shadowRoot?.querySelectorAll('.ML__raw-latex') ?? [])
          .filter((node) => !node.classList.contains('ML__suggestion'))
          .map((node) => node.textContent ?? "")
          .join(""),
        environment: field._mathfield?.model?.parentEnvironment?.environmentName ?? null,
      };
    })()`);
    assert.match(
      typedCases.value,
      /\\begin\{cases\}\\placeholder\{\} & \\placeholder\{\}\\end\{cases\}/,
      "Space did not accept cases as a real editable environment",
    );
    assert.equal(typedCases.raw, "", "accepted cases input remained in raw-LaTeX mode");

    // Accepting cases above records one native-candidate use. Verify that the
    // synthetic environment candidate now participates in the exact same
    // frequency-ranked native pool as ordinary MathLive `\\b...` commands.
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.setValue("", { insertionMode: "replaceAll" });
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus();
      return true;
    })()`);
    await typeKey("\\", "Backslash", "\\");
    await typeKey("b", "KeyB");
    await sleep(140);

    const bPool = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      const panel = document.querySelector('#visualtex-native-input-suggestion-popover');
      const items = Array.from(panel?.querySelectorAll('li[data-command]') ?? []);
      return {
        raw: Array.from(field.shadowRoot?.querySelectorAll('.ML__raw-latex') ?? [])
          .filter((node) => !node.classList.contains('ML__suggestion'))
          .map((node) => node.textContent ?? "")
          .join(""),
        visible: panel?.classList.contains('is-visible') ?? false,
        selected: panel?.querySelector('li.ML__popover__current')?.dataset.command ?? '',
        commands: items.map((item) => item.dataset.command ?? ''),
      };
    })()`);
    assert.equal(bPool.raw.replace(/\s+/g, ""), "\\b");
    assert.equal(bPool.visible, true, "native candidate pool is not visible for \\b");
    assert.ok(
      bPool.commands.includes("\\begin{cases}"),
      `cases is missing from the normal \\b candidate pool: ${JSON.stringify(bPool.commands)}`,
    );
    assert.ok(
      bPool.commands.some((command) => command !== "\\begin{cases}" && command.startsWith("\\b")),
      `the \\b pool did not retain ordinary MathLive b-commands: ${JSON.stringify(bPool.commands)}`,
    );
    assert.equal(
      bPool.selected,
      "\\begin{cases}",
      `recent cases usage did not promote it through native frequency ranking: ${JSON.stringify(bPool.commands)}`,
    );

    await typeKey("e", "KeyE");
    await sleep(100);
    const bePool = await evaluate(`(() => {
      const panel = document.querySelector('#visualtex-native-input-suggestion-popover');
      return {
        selected: panel?.querySelector('li.ML__popover__current')?.dataset.command ?? '',
        commands: Array.from(panel?.querySelectorAll('li[data-command]') ?? [])
          .map((item) => item.dataset.command ?? ''),
      };
    })()`);
    assert.ok(bePool.commands.includes("\\begin{cases}"), "cases disappeared at the \\be prefix");
    assert.equal(bePool.selected, "\\begin{cases}");

    await typeKey("g", "KeyG");
    await sleep(100);
    const begPool = await evaluate(`(() => {
      const panel = document.querySelector('#visualtex-native-input-suggestion-popover');
      return {
        selected: panel?.querySelector('li.ML__popover__current')?.dataset.command ?? '',
        commands: Array.from(panel?.querySelectorAll('li[data-command]') ?? [])
          .map((item) => item.dataset.command ?? ''),
      };
    })()`);
    assert.ok(begPool.commands.includes("\\begin{cases}"), "cases disappeared at the \\beg prefix");
    assert.equal(begPool.selected, "\\begin{cases}");

    await typeKey(" ", "Space", " ");
    await sleep(100);
    const casesFromRankedPrefix = await evaluate(`document.querySelector("math-field")?.value ?? ""`);
    assert.match(
      casesFromRankedPrefix,
      /\\begin\{cases\}\\placeholder\{\} & \\placeholder\{\}\\end\{cases\}/,
      "accepting the frequency-ranked cases candidate from a short prefix failed",
    );

    await typeKey("\\", "Backslash", "\\");
    for (const letter of "frac") {
      await typeKey(letter, `Key${letter.toUpperCase()}`);
    }
    await typeKey("Enter", "Enter", "\r");
    const candidateCommit = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        value: field.value,
        rowSeparators: (field.value.match(/\\\\\\\\/g) ?? []).length,
      };
    })()`);
    assert.match(
      candidateCommit.value,
      /\\frac/,
      "Enter did not preserve command-candidate priority inside cases",
    );
    assert.equal(
      candidateCommit.rowSeparators,
      0,
      "confirming a command candidate unexpectedly added a cases row",
    );

    await typeKey("Enter", "Enter", "\r");
    const afterCasesEnter = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        value: field.value,
        rowSeparators: (field.value.match(/\\\\\\\\/g) ?? []).length,
      };
    })()`);
    assert.equal(
      afterCasesEnter.rowSeparators,
      1,
      `Enter inside cases did not add exactly one case row: ${afterCasesEnter.value}`,
    );

    const lineCountBeforeOuterEnter = await evaluate(
      `document.querySelectorAll(".formula-line").length`,
    );
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.position = field.lastOffset;
      field.selection = { ranges: [[field.lastOffset, field.lastOffset]], direction: "none" };
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus();
      return true;
    })()`);
    await typeKey("Enter", "Enter", "\r");
    const lineCountAfterOuterEnter = await evaluate(
      `document.querySelectorAll(".formula-line").length`,
    );
    assert.equal(
      lineCountAfterOuterEnter,
      lineCountBeforeOuterEnter + 1,
      "Enter outside the outermost cases did not create a new VisualTeX formula line",
    );

    const toolbarProbe = await evaluate(`(() => {
      const two = document.querySelector('.template-button[data-command-id="cases"]');
      const three = document.querySelector('.template-button[data-command-id="cases-three"]');
      return {
        two: Boolean(two),
        three: Boolean(three),
        twoEnlarged: two?.classList.contains('is-enlarged-cases-preview') ?? false,
        threeEnlarged: three?.classList.contains('is-enlarged-cases-preview') ?? false,
      };
    })()`);
    assert.deepEqual(toolbarProbe, {
      two: true,
      three: true,
      twoEnlarged: true,
      threeEnlarged: true,
    });

    console.log(
      `Windows cases regression passed: native popover ${Math.round(casesSuggestion.width)}px, Space acceptance, candidate priority, Enter row insertion, toolbar previews`,
    );
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(180);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => {});
  }
}

await main();
