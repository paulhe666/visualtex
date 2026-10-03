import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 1000;
const previewPort = 17400 + offset;
const debugPort = 22600 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = `/tmp/visualtex-startup-browser-${process.pid}`;
const chromePath = (process.env.CHROME_PATH || "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome");
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
        "about:blank",
      ],
      { stdio: "ignore" },
    );

    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json();
    const page = targets.find((target) => target.type === "page");
    if (!page) throw new Error(`No Chrome page target found: ${JSON.stringify(targets)}`);

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.addScriptToEvaluateOnNewDocument", {
      source: `(() => {
        const replace = (target, key, value) => {
          try {
            Object.defineProperty(target, key, {
              configurable: true,
              writable: true,
              value,
            });
          } catch {}
        };
        replace(Array.prototype, "at", undefined);
        replace(String.prototype, "replaceAll", undefined);
        replace(globalThis, "queueMicrotask", undefined);
        replace(globalThis, "ResizeObserver", undefined);
        if (globalThis.crypto) replace(globalThis.crypto, "randomUUID", undefined);
        for (const method of ["getItem", "setItem", "removeItem"]) {
          replace(Storage.prototype, method, function () {
            throw new DOMException("Storage disabled by startup regression", "SecurityError");
          });
        }
      })();`,
    });
    await client.send("Page.navigate", { url: baseUrl });

    const started = Date.now();
    let state;
    while (Date.now() - started < 15000) {
      const result = await client.send("Runtime.evaluate", {
        expression: `(() => {
          const root = document.getElementById("root");
          const shell = document.querySelector(".app-shell");
          const text = root?.innerText ?? "";
          return {
            ready: Boolean(shell),
            rootChildren: root?.children.length ?? 0,
            text: text.slice(0, 500),
            compatibility: {
              arrayAt: typeof Array.prototype.at,
              replaceAll: typeof String.prototype.replaceAll,
              queueMicrotask: typeof globalThis.queueMicrotask,
              resizeObserver: typeof globalThis.ResizeObserver,
              nativeUuid: typeof globalThis.crypto?.randomUUID,
            },
          };
        })()`,
        returnByValue: true,
      });
      state = result.result.value;
      if (state?.ready) break;
      await sleep(80);
    }

    if (!state?.ready) {
      const events = client.events.map((event) => ({
        method: event.method,
        exception:
          event.params?.exceptionDetails?.exception?.description ??
          event.params?.exceptionDetails?.text ??
          null,
        console:
          event.params?.args?.map((arg) => arg.value ?? arg.description ?? "") ??
          null,
      }));
      throw new Error(
        `VisualTeX did not mount under degraded WebView APIs: ${JSON.stringify({ state, events })}`,
      );
    }

    for (const name of [
      "arrayAt",
      "replaceAll",
      "queueMicrotask",
      "resizeObserver",
    ]) {
      const value = state.compatibility[name];
      if (value !== "function") {
        throw new Error(`Compatibility API ${name} was not restored: ${value}`);
      }
    }
    if (state.rootChildren < 1 || state.text.includes("界面加载失败")) {
      throw new Error(`Unexpected startup fallback: ${JSON.stringify(state)}`);
    }

    console.log(JSON.stringify(state, null, 2));
    console.log("Startup browser regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(500);
    for (let attempt = 0; attempt < 5; attempt += 1) {
      try {
        await rm(chromeProfile, { recursive: true, force: true });
        break;
      } catch (error) {
        if (attempt === 4) throw error;
        await sleep(200);
      }
    }
  }
}

await main();
