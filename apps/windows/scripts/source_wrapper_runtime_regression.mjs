import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";
import {
  canonicalSourceWrapperCommand,
  sourceWrapperKinds,
} from "../src/editor/sourceWrapperRegistry.ts";

const offset = process.pid % 1000;
const previewPort = 18800 + offset;
const debugPort = 23800 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const browserProfile = createBrowserProfilePath("visualtex-source-wrapper-runtime");
const browserPath = resolveChromiumExecutable();
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return response;
    } catch {}
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
  async evaluate(expression) {
    const result = await this.send("Runtime.evaluate", {
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
  }
  close() {
    this.socket?.close();
  }
}

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
    targets.find((target) => target.type === "page" && target.url.startsWith(baseUrl)) ??
    targets.find((target) => target.type === "page");
  assert.ok(page, "No browser target found");

  client = new CdpClient(page.webSocketDebuggerUrl);
  await client.connect();
  await client.send("Runtime.enable");
  await client.send("Page.enable");
  await client.send("Page.navigate", { url: baseUrl });
  await sleep(1800);

  await client.evaluate(`(() => {
    localStorage.setItem("visualtex.onboarding.v3.completed", "true");
    localStorage.setItem("visualtex.onboarding.windows.desktop.v1.1.0.completed", "true");
    const storageKey = "visualtex-editor";
    const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
    const line = { id: crypto.randomUUID(), latex: "" };
    persisted.state = {
      ...(persisted.state || {}),
      lines: [line],
      activeLineId: line.id,
      inputBehavior: {
        ...(persisted.state?.inputBehavior || {}),
        autoExitWrapperCommand: false,
        autoExitAccent: false,
      },
    };
    localStorage.setItem(storageKey, JSON.stringify(persisted));
  })()`);
  await client.send("Page.reload", { ignoreCache: true });
  await sleep(1800);

  const readyStarted = Date.now();
  while (Date.now() - readyStarted < 12_000) {
    const ready = await client.evaluate(
      'Boolean(document.querySelector("math-field.visual-mathfield")?.shadowRoot)',
    );
    if (ready) break;
    await sleep(60);
  }

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
      text: value.length === 1 ? value : "",
      unmodifiedText: value.length === 1 ? value : "",
    });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
    await sleep(45);
  };

  const typeText = async (text) => {
    for (const character of text) {
      const upper = character.toUpperCase();
      const code =
        character === "\\"
          ? "Backslash"
          : /[A-Z]/.test(upper)
            ? `Key${upper}`
            : "";
      const vk =
        character === "\\"
          ? 220
          : upper.charCodeAt(0);
      await key(character, code, vk);
    }
  };

  const allCases = [...sourceWrapperKinds.entries()]
    .map(([command, kind]) => ({
      command,
      canonical: canonicalSourceWrapperCommand(command),
      kind,
    }))
    .sort((left, right) => left.command.localeCompare(right.command));
  const start = Math.max(
    0,
    Number.parseInt(process.env.VISUALTEX_SOURCE_WRAPPER_START ?? "0", 10) || 0,
  );
  const limit = Math.max(
    1,
    Number.parseInt(
      process.env.VISUALTEX_SOURCE_WRAPPER_LIMIT ?? String(allCases.length),
      10,
    ) || allCases.length,
  );
  const cases = allCases.slice(start, start + limit);
  const results = [];

  for (const testCase of cases) {
    await client.evaluate(`(() => {
      const field = document.querySelector("math-field.visual-mathfield");
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
      field.focus();
      field.position = field.lastOffset;
      field.shadowRoot
        ?.querySelector('[part="keyboard-sink"]')
        ?.focus({ preventScroll: true });
      return true;
    })()`);
    await sleep(100);

    await typeText(testCase.command);
    await key(" ", "Space", 32);
    await sleep(100);

    const empty = await client.evaluate(`(() => {
      const field = document.querySelector("math-field.visual-mathfield");
      return {
        value: field?.value ?? "",
        raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
          .filter((node) => !node.classList.contains("ML__suggestion"))
          .map((node) => node.textContent ?? "")
          .join(""),
        mode: field?.mode ?? "",
        sourceCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
        sourceKind: field?.dataset.visualtexSourceWrapperKind ?? "",
        placeholderCount:
          field?.shadowRoot?.querySelectorAll(
            ".visualtex-structural-placeholder, .ML__placeholder",
          ).length ?? -1,
      };
    })()`);

    assert.equal(empty.value, `${testCase.canonical}{}`, `${testCase.command}: empty source shell value`);
    assert.equal(empty.raw, `${testCase.canonical}{}`, `${testCase.command}: raw source shell`);
    assert.equal(empty.mode, "latex", `${testCase.command}: source mode`);
    assert.equal(empty.sourceCommand, testCase.canonical, `${testCase.command}: source command`);
    assert.equal(empty.sourceKind, testCase.kind, `${testCase.command}: source kind`);
    assert.equal(empty.placeholderCount, 0, `${testCase.command}: placeholder leaked into source-first input`);

    await key("A", "KeyA", 65);
    await sleep(80);
    const typed = await client.evaluate(`(() => {
      const field = document.querySelector("math-field.visual-mathfield");
      return {
        value: field?.value ?? "",
        mode: field?.mode ?? "",
        sourceCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
        placeholderCount:
          field?.shadowRoot?.querySelectorAll(
            ".visualtex-structural-placeholder, .ML__placeholder",
          ).length ?? -1,
      };
    })()`);
    assert.equal(typed.value, `${testCase.canonical}{A}`, `${testCase.command}: typed raw source`);
    assert.equal(typed.mode, "latex", `${testCase.command}: stays source-first before Enter`);
    assert.equal(typed.sourceCommand, testCase.canonical, `${testCase.command}: source wrapper remains active`);
    assert.equal(typed.placeholderCount, 0, `${testCase.command}: typed source gained placeholder`);

    await key("Enter", "Enter", 13);
    await sleep(100);
    const committed = await client.evaluate(`(() => {
      const field = document.querySelector("math-field.visual-mathfield");
      return {
        value: field?.value ?? "",
        mode: field?.mode ?? "",
        sourceCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
        placeholderCount:
          field?.shadowRoot?.querySelectorAll(
            ".visualtex-structural-placeholder, .ML__placeholder",
          ).length ?? -1,
        error: Boolean(field?.shadowRoot?.querySelector(".ML__error")),
      };
    })()`);
    assert.equal(committed.value, `${testCase.canonical}{A}`, `${testCase.command}: committed value`);
    assert.equal(committed.mode, "math", `${testCase.command}: Enter did not apply source`);
    assert.equal(committed.sourceCommand, "", `${testCase.command}: source wrapper remained active after Enter`);
    assert.equal(committed.placeholderCount, 0, `${testCase.command}: committed formula contains placeholder`);
    assert.equal(committed.error, false, `${testCase.command}: committed formula rendered as error`);
    results.push({ ...testCase, empty, typed, committed });
  }

  console.log(
    `Source-wrapper runtime regression passed for ${cases.length} commands [${start}, ${start + cases.length}): ${results.map((item) => item.command).join(", ")}`,
  );
} finally {
  client?.close();
  browser?.kill("SIGTERM");
  preview.kill("SIGTERM");
  await sleep(200);
  await rm(browserProfile, {
    recursive: true,
    force: true,
    maxRetries: 4,
    retryDelay: 150,
  }).catch((error) => {
    if (process.platform !== "win32" || error?.code !== "EBUSY") throw error;
  });
}
