import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 600;
const previewPort = 23800 + offset;
const debugPort = 30800 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = `/tmp/visualtex-custom-symbol-noninterference-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      if ((await fetch(url)).ok) return;
    } catch {
      // Retry during startup.
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
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`CDP ${method} timed out`));
      }, 90000);
      this.pending.set(id, {
        resolve: (value) => {
          clearTimeout(timer);
          resolve(value);
        },
        reject: (error) => {
          clearTimeout(timer);
          reject(error);
        },
      });
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
        result.exceptionDetails.exception?.description ||
          result.exceptionDetails.text ||
          "Runtime.evaluate failed",
      );
    }
    return result.result.value;
  }
  close() {
    this.socket?.close();
  }
}

async function waitUntil(client, expression, timeoutMs = 10000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const value = await client.evaluate(expression);
    if (value) return value;
    await sleep(50);
  }
  throw new Error(`Timed out waiting for ${expression}`);
}

async function dispatchCharacter(client, key, code, keyCode, modifiers = 0) {
  const common = {
    key,
    code,
    windowsVirtualKeyCode: keyCode,
    nativeVirtualKeyCode: keyCode,
    modifiers,
  };
  await client.send("Input.dispatchKeyEvent", { type: "rawKeyDown", ...common });
  await client.send("Input.dispatchKeyEvent", {
    type: "char",
    ...common,
    text: key,
    unmodifiedText: key,
  });
  await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
  await sleep(70);
}

async function dispatchShortcut(client, key, code, keyCode, modifiers) {
  const common = {
    key,
    code,
    windowsVirtualKeyCode: keyCode,
    nativeVirtualKeyCode: keyCode,
    modifiers,
  };
  await client.send("Input.dispatchKeyEvent", { type: "rawKeyDown", ...common });
  await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
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
        "--window-size=1450,980",
        "about:blank",
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
    const page = targets.find(
      (target) =>
        target.type === "page" &&
        (target.url === "about:blank" || target.url.startsWith(baseUrl)),
    );
    assert.ok(page);
    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.addScriptToEvaluateOnNewDocument", {
      source: `(() => {
        try {
          localStorage.setItem("visualtex.onboarding.v3.completed", "true");
          localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
          localStorage.setItem("visualtex.onboarding.macos.desktop.v1.2.0.completed", "true");
          localStorage.setItem("visualtex.office.macos.native-first-run.v1.2.0.completed", "true");
          const key = "visualtex-editor";
          let persisted;
          try { persisted = JSON.parse(localStorage.getItem(key) || "null"); }
          catch { persisted = null; }
          if (!persisted || typeof persisted !== "object") persisted = { state: {}, version: 0 };
          persisted.state = {
            ...(persisted.state || {}),
            checkUpdatesOnStartup: false,
          };
          localStorage.setItem(key, JSON.stringify(persisted));
        } catch {}
      })();`,
    });
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(2500);
    await waitUntil(
      client,
      `document.readyState === "complete" && Boolean(document.querySelector("math-field"))`,
      30000,
    );

    const reset = async ({ title = "History Test", lines, inputBehavior = {} }) => {
      await client.evaluate(`(() => {
        localStorage.setItem("visualtex.onboarding.v3.completed", "true");
        localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
        localStorage.setItem("visualtex.onboarding.macos.desktop.v1.2.0.completed", "true");
        localStorage.setItem("visualtex.office.macos.native-first-run.v1.2.0.completed", "true");
        localStorage.removeItem("visualtex.custom-symbols.v1");
        let persisted;
        try { persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "null"); }
        catch { persisted = null; }
        if (!persisted || typeof persisted !== "object") persisted = { state: {}, version: 0 };
        persisted.state = {
          ...(persisted.state || {}),
          title: ${JSON.stringify(title)},
          lines: ${JSON.stringify(lines)},
          activeLineId: ${JSON.stringify(lines[0].id)},
          sourceOpen: false,
          checkUpdatesOnStartup: false,
          inputBehavior: {
            autoExitSuperscript: true,
            autoExitSubscript: true,
            autoExitAccent: true,
            autoExitWrapperCommand: true,
            showStructuredCommandSuggestions: true,
            showOtherCommandSuggestions: false,
            ...${JSON.stringify(inputBehavior)},
          },
        };
        delete persisted.state.latex;
        localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
      })()`);
      await client.send("Page.reload", { ignoreCache: true });
      await sleep(550);
      await waitUntil(client, `document.querySelectorAll("math-field").length === ${lines.length}`);
    };

    process.stdout.write("[custom-symbol-noninterference] reset script case\n");
    await reset({
      lines: [{ id: "script-line", latex: "" }],
      inputBehavior: { autoExitSuperscript: false, autoExitSubscript: false },
    });
    await client.evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
    })()`);
    await dispatchCharacter(client, "x", "KeyX", 88);
    await dispatchCharacter(client, "^", "Digit6", 54, 8);
    await dispatchCharacter(client, "2", "Digit2", 50);
    const scriptState = await client.evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        value: field.value,
        position: field.position,
        lastOffset: field.lastOffset,
        customStyles: Array.from(field.shadowRoot?.querySelectorAll("style") || [])
          .filter((style) => style.id?.includes("custom-symbol") || style.textContent?.includes("visualtex-custom-symbol-"))
          .map((style) => style.textContent || ""),
      };
    })()`);
    assert.match(scriptState.value, /^x(?:\^2|\^\{2\})$/);
    assert.notEqual(scriptState.position, scriptState.lastOffset);
    assert.equal(
      scriptState.customStyles.some((css) => css.includes("visualtex-custom-symbol-vtxtestsymbol")),
      false,
      "A normal formula must not receive prototype/user symbol CSS",
    );
    process.stdout.write("[custom-symbol-noninterference] ordinary script input verified\n");

    console.log(
      "Custom symbol non-interference regression passed for ordinary superscript input and per-field CSS isolation",
    );
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(220);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => undefined);
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
