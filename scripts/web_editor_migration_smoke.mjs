import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import { createServer } from "node:http";
import process from "node:process";
import {
  browserTestProfilePath,
  resolveBrowserTestChromePath,
} from "./browser_test_runtime.mjs";

const portOffset = process.pid % 1000;
const previewPort = 7300 + portOffset;
const ocrMockPort = 9300 + portOffset;
const debugPort = 12300 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}/editor`;
const ocrMockBaseUrl = `http://127.0.0.1:${ocrMockPort}/v1`;
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
  let openAiStructuredRejections = 0;
  const ocrMock = createServer((request, response) => {
    response.setHeader("access-control-allow-origin", `http://127.0.0.1:${previewPort}`);
    response.setHeader(
      "access-control-allow-headers",
      "authorization, content-type",
    );
    response.setHeader("access-control-allow-methods", "POST, OPTIONS");
    if (request.method === "OPTIONS") {
      response.writeHead(204);
      response.end();
      return;
    }
    if (request.method === "POST" && request.url === "/v1/responses") {
      let body = "";
      request.setEncoding("utf8");
      request.on("data", (chunk) => {
        body += chunk;
      });
      request.on("end", () => {
        const value = JSON.parse(body);
        if (value.text?.format?.type === "json_schema") {
          openAiStructuredRejections += 1;
          response.writeHead(400, { "content-type": "application/json" });
          response.end(
            JSON.stringify({
              error: {
                message: "unknown field text.format json_schema",
              },
            }),
          );
          return;
        }
        response.writeHead(200, { "content-type": "application/json" });
        response.end(
          JSON.stringify({
            output_text: JSON.stringify({
              formulas: [{ latex: "\\sqrt{x}" }],
            }),
          }),
        );
      });
      return;
    }
    response.writeHead(404, { "content-type": "application/json" });
    response.end(JSON.stringify({ error: "Unknown mock OCR route" }));
  });
  await new Promise((resolve, reject) => {
    ocrMock.once("error", reject);
    ocrMock.listen(ocrMockPort, "127.0.0.1", resolve);
  });

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
    await client.send("Browser.grantPermissions", {
      origin: `http://127.0.0.1:${previewPort}`,
      permissions: ["clipboardReadWrite", "clipboardSanitizedWrite"],
    });
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

    await evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.web.v3.completed", "true");
      localStorage.setItem(
        "visualtex.web.ocr.configuration.v1",
        JSON.stringify({
          activeProvider: "openai-compatible",
          openAiCompatible: {
            protocol: "responses",
            baseUrl: ${JSON.stringify(ocrMockBaseUrl)},
            model: "mock-vision",
            prompt: "Recognize the formula",
          },
        }),
      );
      sessionStorage.setItem(
        "visualtex.web.ocr.secrets.v1",
        JSON.stringify({ openAiApiKey: "mock-session-key" }),
      );
    })()`);
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
      desktopOnlyButtons: [...document.querySelectorAll("button")].filter((button) =>
        /小键盘|keypad|Office|检查更新|check for updates/i.test(
          [button.textContent, button.getAttribute("aria-label"), button.title]
            .filter(Boolean)
            .join(" "),
        ),
      ).length,
      ocrButtons: [...document.querySelectorAll("button")].filter((button) =>
        /图片公式识别|Formula image OCR/i.test(
          [button.textContent, button.getAttribute("aria-label"), button.title]
            .filter(Boolean)
            .join(" "),
        ),
      ).length,
      zoom: document.querySelector(".canvas-controls span")?.textContent?.trim() ?? "",
      errorBoundary: Boolean(document.querySelector('[role="alert"]')),
    }))()`);
    assert.equal(boundary.page, "editor", JSON.stringify(boundary));
    assert.equal(boundary.app, true, JSON.stringify(boundary));
    assert.equal(boundary.field, true, JSON.stringify(boundary));
    assert.match(boundary.fileAccept, /\.visualtex/);
    assert.equal(boundary.desktopOnlyButtons, 0, JSON.stringify(boundary));
    assert.ok(boundary.ocrButtons > 0, JSON.stringify(boundary));
    assert.equal(boundary.zoom, "45%", JSON.stringify(boundary));
    assert.equal(boundary.errorBoundary, false, JSON.stringify(boundary));

    const fontRuntimeState = await evaluate(`(() => {
      const fields = [...document.querySelectorAll("math-field")];
      return {
        fieldCount: fields.length,
        editorStyles: fields.filter((field) =>
          field.shadowRoot?.getElementById("visualtex-formula-font-style"),
        ).length,
        conflictingStyles: fields.filter((field) =>
          field.shadowRoot?.getElementById(
            "visualtex-formula-font-runtime-shadow-style",
          ),
        ).length,
        previewUsesBroadLatinOverride:
          document
            .getElementById("visualtex-formula-font-runtime-style")
            ?.textContent?.includes(".ML__latin") ?? false,
      };
    })()`);
    assert.ok(fontRuntimeState.fieldCount > 0, JSON.stringify(fontRuntimeState));
    assert.equal(
      fontRuntimeState.editorStyles,
      fontRuntimeState.fieldCount,
      JSON.stringify(fontRuntimeState),
    );
    assert.equal(
      fontRuntimeState.conflictingStyles,
      0,
      JSON.stringify(fontRuntimeState),
    );
    assert.equal(
      fontRuntimeState.previewUsesBroadLatinOverride,
      false,
      JSON.stringify(fontRuntimeState),
    );

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
      field.position = field.lastOffset;
      field.focus();
      return field.value;
    })()`);
    assert.match(formulaValue, /\\frac/);

    const clipboardState = await evaluate(`(async () => {
      const field = document.querySelector("math-field");
      if (!field) throw new Error("Formula field is unavailable");
      const canvas = document.createElement("canvas");
      canvas.width = 2;
      canvas.height = 2;
      const context = canvas.getContext("2d");
      context.fillStyle = "white";
      context.fillRect(0, 0, canvas.width, canvas.height);
      const blob = await new Promise((resolve, reject) =>
        canvas.toBlob(
          (value) => value ? resolve(value) : reject(new Error("Unable to create clipboard PNG")),
          "image/png",
        ),
      );
      await navigator.clipboard.write([
        new ClipboardItem({ "image/png": blob }),
      ]);
      const clipboardTypes = (await navigator.clipboard.read())
        .flatMap((item) => item.types);
      field.focus();
      field.position = field.lastOffset;
      const keyboardSink = field.shadowRoot?.querySelector('[part="keyboard-sink"]');
      if (!keyboardSink) throw new Error("MathLive keyboard sink is unavailable");
      keyboardSink.focus();
      return { clipboardTypes };
    })()`);
    assert.ok(
      clipboardState.clipboardTypes.includes("image/png"),
      JSON.stringify(clipboardState),
    );
    await client.send("Input.dispatchKeyEvent", {
      type: "keyDown",
      modifiers: 2,
      key: "Control",
      code: "ControlLeft",
      windowsVirtualKeyCode: 17,
      nativeVirtualKeyCode: 17,
    });
    await client.send("Input.dispatchKeyEvent", {
      type: "keyDown",
      modifiers: 2,
      key: "v",
      code: "KeyV",
      windowsVirtualKeyCode: 86,
      nativeVirtualKeyCode: 86,
    });
    await client.send("Input.dispatchKeyEvent", {
      type: "keyUp",
      modifiers: 2,
      key: "v",
      code: "KeyV",
      windowsVirtualKeyCode: 86,
      nativeVirtualKeyCode: 86,
    });
    await client.send("Input.dispatchKeyEvent", {
      type: "keyUp",
      modifiers: 0,
      key: "Control",
      code: "ControlLeft",
      windowsVirtualKeyCode: 17,
      nativeVirtualKeyCode: 17,
    });
    const directPasteFormula = await evaluate(`new Promise((resolve, reject) => {
      const started = performance.now();
      const done = () => {
        const field = document.querySelector("math-field");
        const value = field?.value ?? "";
        if (value.includes("\\\\sqrt{x}")) return resolve(value);
        const toast = document.querySelector(".toast")?.textContent?.trim() ?? "";
        if (/OCR failed|OCR 失败/.test(toast)) return reject(new Error(toast));
        if (performance.now() - started > 5000) {
          return reject(new Error("Pasted image was not recognized and inserted: " + JSON.stringify({
            activeTag: document.activeElement?.tagName,
            value,
            toast,
          })));
        }
        setTimeout(done, 30);
      };
      done();
    })`);
    assert.match(directPasteFormula, /\\frac/);
    assert.match(directPasteFormula, /\\sqrt/);
    assert.equal(
      await evaluate(`Boolean(document.querySelector(".web-ocr-dialog"))`),
      false,
    );
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
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
      field.position = field.lastOffset;
      field.focus();
    })()`);

    await evaluate(`document.querySelector('.menu-button')?.click()`);
    const helpOpened = await evaluate(`(() => {
      const help = [...document.querySelectorAll('[role="menuitem"]')].find((item) =>
        /帮助手册|Help manual/i.test(item.textContent || ""),
      );
      if (!help) return { found: false, hit: false };
      const bounds = help.getBoundingClientRect();
      const hit = document.elementFromPoint(
        bounds.left + bounds.width / 2,
        bounds.top + bounds.height / 2,
      );
      const clickable = hit === help || Boolean(hit?.closest('[role="menuitem"]') === help);
      if (clickable) hit.dispatchEvent(new MouseEvent("click", { bubbles: true }));
      return { found: true, hit: clickable };
    })()`);
    assert.deepEqual(helpOpened, { found: true, hit: true });
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

    await evaluate(`document.querySelector('[aria-label="图片公式识别"], [aria-label="Formula image OCR"]')?.click()`);
    await evaluate(`new Promise((resolve, reject) => {
      const started = performance.now();
      const done = () => {
        if (document.querySelector(".web-ocr-dialog")) return resolve(true);
        if (performance.now() - started > 5000) return reject(new Error("Web OCR dialog did not open"));
        setTimeout(done, 30);
      };
      done();
    })`);
    const ocrState = await evaluate(`(() => ({
      options: [...document.querySelectorAll('.web-ocr-dialog option')].map((option) => option.value),
      activeProvider: document.querySelector(
        '.web-ocr-dialog [aria-label="OCR 提供器"] select, .web-ocr-dialog [aria-label="OCR provider"] select',
      )?.value ?? "",
      privacy: document.querySelector('.web-ocr-dialog .ocr-provider-actions span')?.textContent ?? "",
      localRuntime: /安装 OCR 运行环境|Install OCR runtime/i.test(document.querySelector('.web-ocr-dialog')?.textContent ?? ""),
    }))()`);
    assert.ok(ocrState.options.includes("simpletex"), JSON.stringify(ocrState));
    assert.ok(ocrState.options.includes("paddleocr"), JSON.stringify(ocrState));
    assert.ok(ocrState.options.includes("mathpix"), JSON.stringify(ocrState));
    assert.ok(ocrState.options.includes("openai-compatible"), JSON.stringify(ocrState));
    assert.equal(ocrState.activeProvider, "openai-compatible", JSON.stringify(ocrState));
    assert.match(ocrState.privacy, /直接发送|directly/i);
    assert.equal(ocrState.localRuntime, false, JSON.stringify(ocrState));

    await evaluate(`(() => {
      const input = document.querySelector('.web-ocr-dialog input[type="file"]');
      if (!input) throw new Error("OCR file input is unavailable");
      const file = new File(
        [new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10])],
        "formula.png",
        { type: "image/png" },
      );
      const transfer = new DataTransfer();
      transfer.items.add(file);
      Object.defineProperty(input, "files", {
        configurable: true,
        value: transfer.files,
      });
      input.dispatchEvent(new Event("change", { bubbles: true }));
    })()`);
    await evaluate(`new Promise((resolve, reject) => {
      const started = performance.now();
      const done = () => {
        const recognize = [...document.querySelectorAll(".web-ocr-dialog button")].find(
          (button) => /开始识别|^Recognize$/i.test(button.textContent?.trim() ?? ""),
        );
        if (recognize && !recognize.disabled) {
          recognize.click();
          return resolve(true);
        }
        if (performance.now() - started > 5000) {
          return reject(new Error("OCR image did not become ready"));
        }
        setTimeout(done, 30);
      };
      done();
    })`);
    await evaluate(`new Promise((resolve, reject) => {
      const started = performance.now();
      const done = () => {
        const result = document.querySelector(".ocr-latex-editor textarea")?.value ?? "";
        if (result.includes("\\\\sqrt{x}")) return resolve(true);
        const error = document.querySelector(".ocr-error-box")?.textContent?.trim();
        if (error) return reject(new Error(error));
        if (performance.now() - started > 5000) {
          return reject(new Error("Mock OCR did not return a result"));
        }
        setTimeout(done, 30);
      };
      done();
    })`);
    await evaluate(`(() => {
      const insert = [...document.querySelectorAll(".web-ocr-dialog button")].find(
        (button) => /插入当前光标|Insert at cursor/i.test(button.textContent ?? ""),
      );
      if (!insert || insert.disabled) throw new Error("OCR insert button is unavailable");
      insert.click();
    })()`);
    const insertedFormula = await evaluate(`new Promise((resolve, reject) => {
      const started = performance.now();
      const done = () => {
        const field = document.querySelector("math-field");
        const value = field?.value ?? "";
        if (!document.querySelector(".web-ocr-dialog") && value.includes("\\\\sqrt{x}")) {
          return resolve(value);
        }
        if (performance.now() - started > 5000) {
          return reject(new Error("OCR result was not inserted at the saved cursor"));
        }
        setTimeout(done, 30);
      };
      done();
    })`);
    assert.match(insertedFormula, /\\frac/);
    assert.match(insertedFormula, /\\sqrt/);
    assert.equal(openAiStructuredRejections, 2);

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
    await new Promise((resolve) => ocrMock.close(resolve));
    await sleep(200);
    await rm(chromeProfile, { recursive: true, force: true });
  }
}

await main();
