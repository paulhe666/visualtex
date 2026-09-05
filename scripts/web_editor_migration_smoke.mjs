import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";
import {
  browserTestProfilePath,
  resolveBrowserTestChromePath,
} from "./browser_test_runtime.mjs";

const portOffset = process.pid % 1000;
const previewPort = 7300 + portOffset;
const debugPort = 12300 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}/editor`;
const chromeProfile = browserTestProfilePath("visualtex-web-migration");
const chromePath = resolveBrowserTestChromePath();
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 20_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while the local preview or browser starts.
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
      }, 20_000);
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
        "--window-size=1440,1000",
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
    if (!page) throw new Error("No VisualTeX browser target found");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(250);

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
            "Browser evaluation failed",
        );
      }
      return result.result?.value;
    };

    await evaluate(`localStorage.setItem("visualtex.onboarding.web.v3.completed", "true")`);
    await client.send("Page.reload", { ignoreCache: true });
    await evaluate(`new Promise((resolve, reject) => {
      const started = performance.now();
      const done = () => {
        if (document.querySelector("math-field")) return resolve(true);
        if (document.querySelector('[role="alert"]')) {
          return reject(new Error(document.querySelector('[role="alert"]').innerText));
        }
        if (performance.now() - started > 15000) {
          return reject(new Error("Editor did not mount"));
        }
        setTimeout(done, 40);
      };
      done();
    })`);

    const boundary = await evaluate(`(() => ({
      page: document.documentElement.dataset.page,
      app: Boolean(document.querySelector(".app-shell")),
      field: Boolean(document.querySelector("math-field")),
      fileAccept: document.querySelector('input[type="file"]')?.accept ?? "",
      nativeButtons: [...document.querySelectorAll("button")].filter((button) =>
        /OCR|小键盘|keypad|Office|检查更新|check for updates/i.test(
          [button.textContent, button.getAttribute("aria-label"), button.title]
            .filter(Boolean)
            .join(" "),
        ),
      ).length,
      errorBoundary: Boolean(document.querySelector('[role="alert"]')),
    }))()`);
    assert.equal(boundary.page, "editor", JSON.stringify(boundary));
    assert.equal(boundary.app, true, JSON.stringify(boundary));
    assert.equal(boundary.field, true, JSON.stringify(boundary));
    assert.match(boundary.fileAccept, /\.visualtex/);
    assert.equal(boundary.nativeButtons, 0, JSON.stringify(boundary));
    assert.equal(boundary.errorBoundary, false, JSON.stringify(boundary));

    const formulaValue = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.focus();
      field.setValue("\\\\frac{a}{b}", {
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
      return field.value;
    })()`);
    assert.match(formulaValue, /\\frac/);

    await evaluate(`document.querySelector('.menu-button')?.click()`);
    const helpOpened = await evaluate(`(() => {
      const help = [...document.querySelectorAll('[role="menuitem"]')].find((item) =>
        /帮助手册|Help manual/i.test(item.textContent || ""),
      );
      help?.click();
      return Boolean(help);
    })()`);
    assert.equal(helpOpened, true);
    await evaluate(`new Promise((resolve) => {
      const done = () => document.querySelector(".help-dialog")
        ? resolve(true)
        : setTimeout(done, 30);
      done();
    })`);
    const helpState = await evaluate(`(() => ({
      subtitle: document.querySelector(".help-dialog-header span")?.textContent ?? "",
      text: document.querySelector(".help-dialog-content")?.textContent ?? "",
    }))()`);
    assert.equal(helpState.subtitle, "VisualTeX Web");
    assert.match(helpState.text, /浏览器|browser/i);
    await evaluate(`document.querySelector('[aria-label="关闭帮助手册"], [aria-label="Close help manual"]')?.click()`);

    await evaluate(`document.querySelector('[data-classic-bottom-view="source"], .source-toggle')?.click()`);
    await evaluate(`new Promise((resolve, reject) => {
      const started = performance.now();
      const done = () => {
        if (document.querySelector(".cm-content")) return resolve(true);
        if (performance.now() - started > 5000) return reject(new Error("Source editor did not open"));
        setTimeout(done, 30);
      };
      done();
    })`);

    console.log("Web editor migration smoke test passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(200);
    await rm(chromeProfile, { recursive: true, force: true });
  }
}

await main();
