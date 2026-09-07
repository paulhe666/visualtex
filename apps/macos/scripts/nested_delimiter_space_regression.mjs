import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const portOffset = process.pid % 800;
const previewPort = 8400 + portOffset;
const debugPort = 13400 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = `/tmp/visualtex-nested-delimiter-space-${process.pid}`;
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

    await evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
      localStorage.setItem("visualtex.onboarding.macos.desktop.v1.2.0.completed", "true");
      localStorage.setItem("visualtex.office.macos.native-first-run.v1.2.0.completed", "true");
      localStorage.setItem("visualtex.release-welcome.1.2.6.seen", "true");
      const key = "visualtex-editor";
      const persisted = JSON.parse(localStorage.getItem(key) || "{}");
      persisted.state = {
        ...(persisted.state || {}),
        checkUpdatesOnStartup: false,
        autoPairDelimiters: true,
        inputBehavior: {
          ...(persisted.state?.inputBehavior || {}),
          showStructuredCommandSuggestions: true,
          showOtherCommandSuggestions: true,
          autoExitSuperscript: true,
          autoExitSubscript: true,
          autoExitAccent: true,
          autoExitWrapperCommand: true,
        },
      };
      localStorage.setItem(key, JSON.stringify(persisted));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(700);
    await evaluate(`new Promise((resolve) => {
      const poll = () => document.querySelector("math-field") ? resolve(true) : setTimeout(poll, 25);
      poll();
    })`);

    const typeCharacter = async (key, code, keyCode, pause = 120) => {
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

    const typeRawCommand = async (command) => {
      await typeCharacter("\\", "Backslash", 220, 180);
      for (const character of command) {
        const upper = character.toUpperCase();
        await typeCharacter(character, `Key${upper}`, upper.charCodeAt(0), 80);
      }
    };

    const prepareNested = async () =>
      evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.setValue("\\\\left\\{\\\\left(a+\\\\vec{B}\\\\right)\\\\vec{C}\\\\right\\}", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        let target = null;
        for (let offset = 0; offset <= field.lastOffset; offset += 1) {
          const prefix = field.getValue(0, offset, "latex");
          const depth = field.getElementInfo(offset)?.depth ?? -1;
          if (prefix.endsWith("\\\\vec{B}") && (!target || depth > target.depth)) {
            target = { offset, depth };
          }
        }
        if (!target || target.depth < 2) throw new Error("Screenshot-equivalent nested delimiter target was not found");
        field.position = target.offset;
        field.selection = { ranges: [[target.offset, target.offset]], direction: "none" };
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        return target;
      })()`);

    const state = () =>
      evaluate(`(() => {
        const field = document.querySelector("math-field");
        const caret = Array.from(field.shadowRoot?.querySelectorAll('.ML__caret, .ML__placeholder-selected, .ML__selected') ?? [])
          .find((node) => node.getBoundingClientRect().height > 0) || null;
        let delimiterDepth = 0;
        let current = caret;
        while (current) {
          if (current.classList?.contains('ML__left-right')) delimiterDepth += 1;
          current = current.parentElement;
        }
        return {
          value: field.value,
          position: field.position,
          selection: field.selection,
          delimiterDepth,
          modelDepth: field.getElementInfo(field.position)?.depth ?? null,
          pendingNativeSuggestion: field.dataset.pendingNativeSuggestion || "",
          visibleCandidateLabels: Array.from(
            document.querySelectorAll('#visualtex-native-input-suggestion-popover li[data-command] .ML__popover__latex')
          ).map((node) => (node.textContent || '').trim()).filter(Boolean),
          offsets: Array.from({ length: field.lastOffset + 1 }, (_, offset) => ({
            offset,
            depth: field.getElementInfo(offset)?.depth ?? null,
            latex: field.getElementInfo(offset)?.latex ?? "",
            prefix: field.getValue(0, offset, "latex"),
          })),
        };
      })()`);

    const target = await prepareNested();
    assert.ok(target.depth >= 2);
    await typeRawCommand("nabla");
    await typeCharacter(" ", "Space", 32, 180);
    const nablaCommitted = await state();
    assert.ok(
      (nablaCommitted.modelDepth ?? -1) >= 2,
      `Confirming \\nabla escaped the innermost delimiter: ${JSON.stringify(nablaCommitted)}`,
    );
    assert.match(nablaCommitted.value, /\\nabla/);
    await typeCharacter("q", "KeyQ", 81, 160);
    const nablaContinued = await state();
    assert.ok(
      (nablaContinued.modelDepth ?? -1) >= 2,
      `Typing after \\nabla continued outside the innermost delimiter: ${JSON.stringify(nablaContinued)}`,
    );
    assert.match(
      nablaContinued.value,
      /\\nabla q?\\right\)/,
      `The character after \\nabla was not kept inside the inner parentheses: ${JSON.stringify(nablaContinued)}`,
    );

    await evaluate(`(() => {
      const key = "visualtex-editor";
      const persisted = JSON.parse(localStorage.getItem(key) || "{}");
      persisted.state = {
        ...(persisted.state || {}),
        lines: [{ id: "accent-auto-exit-line", latex: "" }],
        activeLineId: "accent-auto-exit-line",
        inputBehavior: {
          ...(persisted.state?.inputBehavior || {}),
          autoExitAccent: true,
        },
      };
      localStorage.setItem(key, JSON.stringify(persisted));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(650);
    await evaluate(`new Promise((resolve) => {
      const poll = () => document.querySelector("math-field") ? resolve(true) : setTimeout(poll, 25);
      poll();
    })`);
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.focus();
      field.position = field.lastOffset;
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
    })()`);
    await typeRawCommand("vec");
    await typeCharacter(" ", "Space", 32, 180);
    await typeCharacter("B", "KeyB", 66, 180);
    const accentAfterB = await state();
    await typeCharacter("q", "KeyQ", 81, 180);
    const accentAutoExit = await state();
    assert.match(
      accentAutoExit.value,
      /^\\vec\{B\}q$/,
      `Accent auto-exit no longer leaves the caret after the accent body: ${JSON.stringify(accentAutoExit)}`,
    );

    await prepareNested();
    await typeRawCommand("dag");
    await sleep(120);
    const daggerCandidates = await state();
    assert.equal(
      daggerCandidates.visibleCandidateLabels.filter(
        (label) => label === "\\dagger",
      ).length,
      1,
      `The visible candidate list should expose exactly one canonical \\dagger: ${JSON.stringify(daggerCandidates)}`,
    );
    assert.ok(
      !daggerCandidates.visibleCandidateLabels.includes("\\dag"),
      `The visible candidate list still exposed legacy \\dag: ${JSON.stringify(daggerCandidates)}`,
    );
    await typeCharacter(" ", "Space", 32, 180);
    const daggerCommitted = await state();
    assert.match(
      daggerCommitted.value,
      /\\vec\{B\}\\dagger/,
      `Legacy \\dag input was not canonicalized: ${JSON.stringify(daggerCommitted)}`,
    );
    assert.ok(
      (daggerCommitted.modelDepth ?? -1) >= 2,
      JSON.stringify(daggerCommitted),
    );

    await prepareNested();
    await typeRawCommand("sqrt");
    await typeCharacter(" ", "Space", 32, 180);
    const sqrtCommitted = await state();
    assert.match(
      sqrtCommitted.value,
      /\\vec\{B\}\\sqrt\{\\placeholder\{\}\}/,
    );
    assert.ok(
      sqrtCommitted.delimiterDepth >= 2 || !sqrtCommitted.selection.ranges.every(([a, b]) => a === b),
      `Structured command left the nested delimiter: ${JSON.stringify(sqrtCommitted)}`,
    );

    console.log("Nested delimiter Space and dagger canonicalization regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(250);
    await rm(chromeProfile, {
      recursive: true,
      force: true,
      maxRetries: 5,
      retryDelay: 100,
    });
  }
}

await main();
