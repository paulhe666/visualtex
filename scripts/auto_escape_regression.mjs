import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const portOffset = process.pid % 700;
const previewPort = 7600 + portOffset;
const debugPort = 12600 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}/editor`;
const chromeProfile = `/tmp/visualtex-auto-escape-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 12000) {
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

function keyInfo(character) {
  if (/^[A-Za-z]$/.test(character)) {
    const upper = character.toUpperCase();
    return { code: `Key${upper}`, keyCode: upper.charCodeAt(0), modifiers: 0 };
  }
  if (character === "=") return { code: "Equal", keyCode: 187, modifiers: 0 };
  if (character === ">") return { code: "Period", keyCode: 190, modifiers: 8 };
  throw new Error(`Unsupported test character: ${character}`);
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
        "--window-size=1200,800",
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
      await evaluate(`new Promise((resolve) => {
        const done = () => document.querySelector("math-field")
          ? resolve(true)
          : setTimeout(done, 25);
        done();
      })`);
      await sleep(120);
    };

    const configure = async (enabled) => {
      await evaluate(`(() => {
        localStorage.setItem("visualtex.onboarding.v3.completed", "true");
        localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
        localStorage.setItem("visualtex.onboarding.macos.desktop.v1.2.0.completed", "true");
        localStorage.setItem("visualtex.office.macos.native-first-run.v1.2.0.completed", "true");
        const key = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(key) || "{}");
        persisted.state = {
          ...(persisted.state || {}),
          lines: [{ id: "auto-escape-line", latex: "" }],
          activeLineId: "auto-escape-line",
          sourceOpen: false,
          editorLayout: "classic",
          classicDockOpen: true,
          sidebarOpen: false,
          inputBehavior: {
            autoEscapeShortcuts: ${enabled},
            autoExitSuperscript: true,
            autoExitSubscript: true,
            autoExitAccent: true,
            autoExitWrapperCommand: true,
            showStructuredCommandSuggestions: true,
            showOtherCommandSuggestions: false,
          },
        };
        localStorage.setItem(key, JSON.stringify(persisted));
      })()`);
      await reload();
    };

    const clearAndFocus = async () => {
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
        return true;
      })()`);
      await sleep(50);
    };

    const typeText = async (text) => {
      await clearAndFocus();
      for (const character of text) {
        const { code, keyCode, modifiers } = keyInfo(character);
        const common = {
          key: character,
          code,
          modifiers,
          windowsVirtualKeyCode: keyCode,
          nativeVirtualKeyCode: keyCode,
        };
        await client.send("Input.dispatchKeyEvent", {
          type: "keyDown",
          ...common,
          text: character,
          unmodifiedText: character,
        });
        await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
        await sleep(55);
      }
      await sleep(160);
      return evaluate(`document.querySelector("math-field").value`);
    };

    await configure(true);

    await evaluate(`document.querySelector(".canvas-input-behavior-trigger").click()`);
    await sleep(80);
    const behaviorUi = await evaluate(`(() => {
      const popover = document.querySelector(".input-behavior-popover");
      return {
        trigger: getComputedStyle(document.querySelector(".canvas-input-behavior-trigger")).fontSize,
        width: getComputedStyle(popover).width,
        heading: getComputedStyle(popover.querySelector(".input-behavior-heading strong")).fontSize,
        description: getComputedStyle(popover.querySelector(".input-behavior-heading span")).fontSize,
        optionTitle: getComputedStyle(popover.querySelector(".input-behavior-option strong")).fontSize,
        optionDescription: getComputedStyle(popover.querySelector(".input-behavior-option small")).fontSize,
        headerZIndex: getComputedStyle(document.querySelector(".formula-workspace.editor-pane > .workspace-heading")).zIndex,
        menuZIndex: getComputedStyle(document.querySelector(".input-behavior-menu")).zIndex,
        popoverZIndex: getComputedStyle(popover).zIndex,
        ...(() => {
          const dock = document.querySelector(".classic-bottom-dock");
          const popoverRect = popover.getBoundingClientRect();
          const dockRect = dock.getBoundingClientRect();
          const overlapTop = Math.max(popoverRect.top, dockRect.top);
          const overlapBottom = Math.min(popoverRect.bottom, dockRect.bottom);
          const overlapHeight = Math.max(0, overlapBottom - overlapTop);
          const x = Math.min(popoverRect.right - 12, Math.max(popoverRect.left + 12, dockRect.left + 12));
          const y = overlapTop + Math.min(24, overlapHeight / 2);
          const topElement = overlapHeight > 0 ? document.elementFromPoint(x, y) : null;
          return {
            overlapHeight,
            popoverOwnsOverlapPoint: Boolean(topElement && popover.contains(topElement)),
          };
        })(),
      };
    })()`);
    assert.equal(behaviorUi.trigger, "13px");
    assert.equal(behaviorUi.width, "420px");
    assert.equal(behaviorUi.heading, "16px");
    assert.equal(behaviorUi.description, "12px");
    assert.equal(behaviorUi.optionTitle, "14px");
    assert.equal(behaviorUi.optionDescription, "12px");
    assert.equal(behaviorUi.headerZIndex, "80");
    assert.equal(behaviorUi.menuZIndex, "90");
    assert.equal(behaviorUi.popoverZIndex, "100");
    assert.ok(
      behaviorUi.overlapHeight > 20,
      `the popover should overlap the classic bottom dock in this regression viewport (actual ${behaviorUi.overlapHeight}px)`,
    );
    assert.equal(
      behaviorUi.popoverOwnsOverlapPoint,
      true,
      "the input-behavior popover must paint above the bottom formula toolbar",
    );

    await evaluate(`document.querySelector(".input-behavior-option input").click()`);
    await sleep(80);
    assert.equal(
      await typeText("alpha"),
      "alpha",
      "the visible automatic-conversion switch must disable shortcuts immediately",
    );
    await evaluate(`document.querySelector(".input-behavior-option input").click()`);
    await sleep(80);
    assert.equal(
      await typeText("alpha"),
      "\\alpha",
      "the visible automatic-conversion switch must restore shortcuts immediately",
    );

    await evaluate(`document.querySelector(".canvas-input-behavior-trigger").click()`);
    await evaluate(`document.querySelector(".export-menu-trigger").click()`);
    await sleep(80);
    const exportUi = await evaluate(`(() => {
      const popover = document.querySelector(".export-menu-popover");
      return {
        width: getComputedStyle(popover).width,
        heading: getComputedStyle(popover.querySelector(".export-menu-heading strong")).fontSize,
        description: getComputedStyle(popover.querySelector(".export-menu-heading span")).fontSize,
        formatTitle: getComputedStyle(popover.querySelector(".export-format-options strong")).fontSize,
        extension: getComputedStyle(popover.querySelector(".export-format-options small")).fontSize,
        pathLabel: getComputedStyle(popover.querySelector(".export-path-copy span")).fontSize,
        path: getComputedStyle(popover.querySelector(".export-path-copy strong")).fontSize,
        chooseButton: getComputedStyle(popover.querySelector(".export-path-button")).fontSize,
      };
    })()`);
    assert.deepEqual(exportUi, {
      width: "420px",
      heading: "16px",
      description: "12px",
      formatTitle: "14px",
      extension: "12px",
      pathLabel: "12px",
      path: "12px",
      chooseButton: "13px",
    });
    await evaluate(`document.querySelector(".export-menu-trigger").click()`);

    assert.match(await typeText(">="), /\\(?:ge|geq)/);
    assert.match(await typeText("geq"), /\\(?:ge|geq)/);
    assert.match(await typeText("leq"), /\\(?:le|leq)/);
    assert.match(await typeText("neq"), /\\(?:ne|neq)/);
    assert.equal(await typeText("pp"), "+");
    assert.equal(await typeText("ss"), "-");
    assert.equal(await typeText("mm"), "\\times");
    assert.equal(await typeText("dd"), "\\div");
    assert.equal(await typeText("eq"), "=");
    assert.match(await typeText("hat"), /\\hat/);
    assert.equal(await typeText("varphi"), "\\varphi");
    assert.equal(
      await typeText("mathbb"),
      "mathbb",
      "font-variant commands must not auto-convert without a backslash",
    );
    assert.equal(await typeText("xx"), "xx", "xx must never become \\times");
    assert.equal(await typeText("dx"), "\\mathrm{d}x");

    await configure(false);
    assert.equal(await typeText("alpha"), "alpha");
    assert.doesNotMatch(await typeText(">="), /\\(?:ge|geq)/);
    assert.equal(await typeText("pp"), "pp");
    assert.equal(await typeText("ss"), "ss");
    assert.equal(await typeText("mm"), "mm");
    assert.equal(await typeText("dd"), "dd");
    assert.equal(await typeText("eq"), "eq");
    assert.equal(await typeText("hat"), "hat");
    assert.equal(await typeText("varphi"), "varphi");
    assert.equal(await typeText("mathbb"), "mathbb");
    assert.equal(await typeText("xx"), "xx");
    assert.equal(await typeText("dx"), "dx");

    console.log("Auto escape regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(200);
    await rm(chromeProfile, { recursive: true, force: true });
  }
}

await main();
