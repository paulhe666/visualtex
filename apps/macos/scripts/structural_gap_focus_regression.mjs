import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";
import { createBrowserProfilePath, resolveChromiumExecutable } from "./browser_test_runtime.mjs";

const offset = process.pid % 1000;
const previewPort = 18800 + offset;
const debugPort = 23800 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const browserProfile = createBrowserProfilePath("visualtex-structural-gap-focus");
const browserPath = resolveChromiumExecutable();
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
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
  close() {
    this.socket?.close();
  }
}

async function main() {
  const preview = spawn(process.execPath, [
    "node_modules/vite/bin/vite.js",
    "preview",
    "--host",
    "127.0.0.1",
    "--port",
    String(previewPort),
    "--strictPort",
  ], { cwd: process.cwd(), stdio: "ignore" });
  let browser;
  let client;
  try {
    await waitFor(baseUrl);
    browser = spawn(browserPath, [
      "--headless=new",
      "--disable-gpu",
      "--no-first-run",
      "--no-default-browser-check",
      `--remote-debugging-port=${debugPort}`,
      `--user-data-dir=${browserProfile}`,
      "--window-size=1400,1000",
      baseUrl,
    ], { stdio: "ignore" });
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
    const page = targets.find((target) => target.type === "page" && target.url.startsWith(baseUrl));
    if (!page) throw new Error("No VisualTeX browser target found");
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
        throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text || "Runtime.evaluate failed");
      }
      return result.result.value;
    };

    await sleep(650);
    await evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
      localStorage.setItem("visualtex.onboarding.macos.desktop.v1.2.0.completed", "true");
      localStorage.setItem("visualtex.office.macos.native-first-run.v1.2.0.completed", "true");
      localStorage.setItem("visualtex.release-welcome.1.2.6.seen", "true");
      const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      persisted.state = { ...(persisted.state || {}), checkUpdatesOnStartup: false, sourceOpen: false };
      localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
    })()`);

    const reload = async () => {
      await client.send("Page.reload", { ignoreCache: true });
      await sleep(700);
      await evaluate(`new Promise((resolve, reject) => {
        const started = performance.now();
        const check = () => {
          const field = document.querySelector("math-field");
          if (field?.shadowRoot) return resolve(true);
          if (performance.now() - started > 8000) return reject(new Error("math-field did not mount"));
          setTimeout(check, 30);
        };
        check();
      })`);
    };
    await reload();

    const loadFormula = async (latex) => {
      await evaluate(`(() => {
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        persisted.state = {
          ...(persisted.state || {}),
          lines: [{ id: "gap-line", latex: ${JSON.stringify(latex)}, mode: "display" }],
          activeLineId: "gap-line",
          sourceOpen: false,
        };
        localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
      })()`);
      await reload();
    };

    const cases = [
      ["x_{i}^{2}\\int_{0}^{1}f(x)\\,\\mathrm{d}x", "x_{i}^{2}"],
      ["A_{m}^{n}\\frac{p+q}{r-s}", "A_{m}^{n}"],
      ["\\frac{a_i}{b^2}x_j^3\\sum_{k=0}^{N}c_k", "\\frac{a_i}{b^2}x_j^3"],
      ["\\sqrt{\\frac{u}{v}}y_k^4\\prod_{r=1}^{M}d_r", "\\sqrt{\\frac{u}{v}}y_k^4"],
    ];
    const results = [];
    for (const [latex, prefix] of cases) {
      await loadFormula(latex);
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
        const entries = Array.from({ length: field.lastOffset + 1 }, (_, offset) => ({ offset, info: field.getElementInfo(offset) }));
        const previous = entries.slice(0, targetOffset).reverse().find(({ info }) => info?.bounds)?.info.bounds;
        const next = entries.slice(targetOffset + 1).find(({ info }) => info?.depth === 0 && info.bounds)?.info.bounds;
        const fieldRect = field.getBoundingClientRect();
        return { targetOffset, point: { x: previous && next ? (previous.right + next.left) / 2 : -1, y: (fieldRect.top + fieldRect.bottom) / 2 } };
      })()`);
      assert.ok(geometry.targetOffset >= 0, JSON.stringify({ latex, prefix, geometry }));
      assert.ok(geometry.point.x >= 0, JSON.stringify({ latex, prefix, geometry }));
      await client.send("Input.dispatchMouseEvent", { type: "mousePressed", x: geometry.point.x, y: geometry.point.y, button: "left", buttons: 1, clickCount: 1 });
      await client.send("Input.dispatchMouseEvent", { type: "mouseReleased", x: geometry.point.x, y: geometry.point.y, button: "left", buttons: 0, clickCount: 1 });
      await sleep(150);
      const state = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const caret = field.shadowRoot?.querySelector(".ML__caret, .ML__text-caret")?.getBoundingClientRect();
        return { position: field.position, caretTop: caret?.top ?? -1, caretBottom: caret?.bottom ?? -1 };
      })()`);
      results.push({ latex, prefix, targetOffset: geometry.targetOffset, actualPosition: state.position });
    }
    console.log(JSON.stringify(results, null, 2));
    const mismatches = results.filter((item) => item.actualPosition !== item.targetOffset);
    if (mismatches.length) {
      console.error(`STRUCTURAL_GAP_MISMATCHES=${mismatches.length}`);
      process.exitCode = 2;
    } else {
      console.log("Structural gap focused regression passed");
    }
  } finally {
    client?.close();
    browser?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(250);
    await rm(browserProfile, { recursive: true, force: true }).catch(() => undefined);
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.stack : error);
  process.exitCode = 1;
});
