import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createServer } from "node:http";
import { readFile, rm } from "node:fs/promises";
import { extname, join, normalize } from "node:path";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const offset = process.pid % 500;
const port = 22400 + offset;
const debugPort = 23400 + offset;
const baseUrl = `http://127.0.0.1:${port}`;
const sessionId = "11111111-2222-4333-8444-555555555555";
const officeUrl =
  `${baseUrl}/dialog/index.html?sessionId=${sessionId}&runtime=vsto-desktop`;
const distRoot = join(process.cwd(), "dist-office-windows-native");
const chromeProfile = createBrowserProfilePath("visualtex-office-layout");
const chromePath = resolveChromiumExecutable();

const mimeTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".woff2": "font/woff2",
};

let session = {
  id: sessionId,
  mode: "create",
  host: "word",
  formulaId: "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee",
  sourceDocumentId: "office-layout-regression",
  sourceObjectId: "visualtex-word-vsto-range:0:0",
  title: "Word Formula",
  lines: [
    {
      id: "99999999-8888-4777-8666-555555555555",
      latex: "",
    },
  ],
  activeLineId: "99999999-8888-4777-8666-555555555555",
  codeFormat: "latex",
  displayMode: "inline",
  objectMode: "nativeOle",
  numbered: false,
  mathTypeNumberPosition: "right",
  fontSizePt: 10.5,
  exportWidth: 0,
  exportHeight: 0,
  exportResult: null,
  originalMetadata: null,
  dirty: false,
  status: "editing",
  autoCommitOnClose: true,
  explicitCancel: false,
  error: null,
  createdAt: Date.now(),
  updatedAt: Date.now(),
  expiresAt: Date.now() + 60_000,
};

function writeJson(response, status, value) {
  response.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
  });
  response.end(JSON.stringify(value));
}

async function readJsonBody(request) {
  const chunks = [];
  for await (const chunk of request) chunks.push(chunk);
  const source = Buffer.concat(chunks).toString("utf8");
  return source ? JSON.parse(source) : {};
}

const server = createServer(async (request, response) => {
  const url = new URL(request.url ?? "/", baseUrl);
  try {
    if (url.pathname === "/api/v1/theme") {
      writeJson(response, 200, { theme: "light", editorLayout: "classic" });
      return;
    }
    if (url.pathname === "/api/v1/preferences") {
      writeJson(response, 200, {
        powerpointDefaultFontSizePt: 20,
        editorPreferences: {
          settings: {
            zoom: 0.5,
            sourceOpen: false,
            classicTileWidth: 300,
            classicDockHeight: 240,
          },
        },
      });
      return;
    }
    if (url.pathname === `/api/v1/sessions/${sessionId}`) {
      if (request.method === "PATCH") {
        session = {
          ...session,
          ...(await readJsonBody(request)),
          updatedAt: Date.now(),
        };
      }
      writeJson(response, 200, session);
      return;
    }
    if (url.pathname === `/api/v1/app/sessions/${sessionId}/close`) {
      writeJson(response, 200, { closed: true });
      return;
    }
    if (url.pathname.startsWith("/api/v1/ocr/")) {
      writeJson(response, 503, { error: "OCR unused in office layout regression" });
      return;
    }
    if (url.pathname.startsWith("/dialog/")) {
      const source = await readFile(join(distRoot, "dialog", "index.html"), "utf8");
      const meta = [
        '<meta name="visualtex-install-token" content="office-layout-regression" />',
        '<meta name="visualtex-native-powerpoint-commit" content="false" />',
        '<meta name="visualtex-theme" content="light" />',
      ].join("\n");
      response.writeHead(200, {
        "Content-Type": "text/html; charset=utf-8",
        "Cache-Control": "no-store",
      });
      response.end(source.replace("</head>", `${meta}\n</head>`));
      return;
    }
    if (url.pathname.startsWith("/assets/")) {
      const relative = normalize(url.pathname.slice(1));
      if (relative.startsWith("..")) {
        response.writeHead(403).end();
        return;
      }
      const content = await readFile(join(distRoot, relative));
      response.writeHead(200, {
        "Content-Type": mimeTypes[extname(relative)] ?? "application/octet-stream",
        "Cache-Control": "no-store",
      });
      response.end(content);
      return;
    }
    response.writeHead(404).end();
  } catch (error) {
    response.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" });
    response.end(String(error));
  }
});

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

async function waitFor(url, timeoutMs = 15_000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {}
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

async function waitForPage() {
  const startedAt = Date.now();
  while (Date.now() - startedAt < 15_000) {
    const targets = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json();
    const target = targets.find(
      (candidate) => candidate.type === "page" && candidate.url.startsWith(baseUrl),
    );
    if (target) return target;
    await sleep(80);
  }
  throw new Error("Timed out waiting for Office layout regression page");
}

async function waitForEvaluation(client, expression, description, timeoutMs = 12_000) {
  const startedAt = Date.now();
  let lastValue;
  while (Date.now() - startedAt < timeoutMs) {
    lastValue = await client.evaluate(expression);
    if (lastValue?.ready) return lastValue;
    await sleep(60);
  }
  throw new Error(
    `Timed out waiting for ${description}: ${JSON.stringify(lastValue)}`,
  );
}

async function main() {
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(port, "127.0.0.1", resolve);
  });

  let chrome;
  let client;
  try {
    await waitFor(`${baseUrl}/api/v1/preferences`);
    chrome = spawn(
      chromePath,
      [
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        `--remote-debugging-port=${debugPort}`,
        `--user-data-dir=${chromeProfile}`,
        "--window-size=1760,760",
        officeUrl,
      ],
      { stdio: "ignore" },
    );

    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const target = await waitForPage();
    client = new CdpClient(target.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");

    const layout = await waitForEvaluation(
      client,
      `(() => {
        const shell = document.querySelector(".office-dialog-shell");
        const workspace = document.querySelector(".workspace.is-office-workspace");
        const header = document.querySelector(".editor-pane-header.is-office-editor-header");
        const headerControls = header?.querySelector(".editor-pane-header-controls");
        const editorPane = document.querySelector(".formula-workspace.editor-pane");
        const scroll = document.querySelector(".editor-pane-scroll");
        const formulaLine = document.querySelector(".formula-line");
        const field = document.querySelector("math-field");
        const status = document.querySelector(".editor-statusbar, .editor-status, .classic-status");
        const bottomTabs = document.querySelector(".classic-bottom-tabs");
        const inputBehaviorTrigger = header?.querySelector(
          ".canvas-input-behavior-trigger",
        );
        const inputBehaviorLabel = inputBehaviorTrigger?.querySelector("span");
        const inputBehaviorArrow = inputBehaviorTrigger?.querySelector(
          "svg:last-child",
        );
        const headerRect = header?.getBoundingClientRect();
        const controlsRect = headerControls?.getBoundingClientRect();
        const editorRect = editorPane?.getBoundingClientRect();
        const scrollRect = scroll?.getBoundingClientRect();
        const lineRect = formulaLine?.getBoundingClientRect();
        const fieldRect = field?.getBoundingClientRect();
        const bottomRect = bottomTabs?.getBoundingClientRect();
        const lineStyle = formulaLine ? getComputedStyle(formulaLine) : null;
        const headerChildren = header
          ? [...header.querySelectorAll(
              ":scope .office-inline-options, :scope .office-formatting-mount, :scope .desktop-editor-header-controls, :scope .canvas-tool-group, :scope .office-inline-actions",
            )]
              .map((child) => {
                const rect = child.getBoundingClientRect();
                return {
                  className: child.className,
                  left: rect.left,
                  right: rect.right,
                  top: rect.top,
                  bottom: rect.bottom,
                  width: rect.width,
                  height: rect.height,
                };
              })
              .filter((item) => item.width > 0 && item.height > 0)
          : [];
        const visibleHeaderRows = new Set(
          headerChildren.map((item) => Math.round(item.top)),
        ).size;
        const allHeaderInside = Boolean(headerRect) &&
          headerChildren.every(
            (rect) =>
              rect.top >= headerRect.top - 1 &&
              rect.bottom <= headerRect.bottom + 1 &&
              rect.left >= headerRect.left - 1 &&
              rect.right <= headerRect.right + 1,
          );

        return {
          ready:
            shell instanceof HTMLElement &&
            workspace instanceof HTMLElement &&
            headerRect &&
            controlsRect &&
            editorRect &&
            scrollRect &&
            lineRect &&
            fieldRect &&
            headerRect.height >= 40 &&
            headerRect.height <= 52 &&
            visibleHeaderRows === 1 &&
            allHeaderInside &&
            lineRect.height >= 40 &&
            fieldRect.height >= 30 &&
            !formulaLine?.classList.contains("is-empty") &&
            lineStyle?.boxShadow === "none" &&
            inputBehaviorTrigger instanceof HTMLElement &&
            inputBehaviorLabel instanceof HTMLElement &&
            getComputedStyle(inputBehaviorLabel).display !== "none" &&
            inputBehaviorLabel.textContent?.trim().length > 0 &&
            inputBehaviorArrow instanceof SVGElement &&
            getComputedStyle(inputBehaviorArrow).display !== "none" &&
            lineRect.top >= scrollRect.top - 1 &&
            lineRect.bottom <= scrollRect.bottom + 1 &&
            (!bottomRect || bottomRect.top >= scrollRect.top + 40),
          shellClass: shell?.className ?? "",
          workspaceClass: workspace?.className ?? "",
          header: headerRect
            ? { left: headerRect.left, right: headerRect.right, top: headerRect.top, bottom: headerRect.bottom, width: headerRect.width, height: headerRect.height }
            : null,
          headerControls: controlsRect
            ? { left: controlsRect.left, right: controlsRect.right, top: controlsRect.top, bottom: controlsRect.bottom, width: controlsRect.width, height: controlsRect.height }
            : null,
          visibleHeaderRows,
          headerChildren,
          allHeaderInside,
          editor: editorRect
            ? { top: editorRect.top, bottom: editorRect.bottom, height: editorRect.height }
            : null,
          scroll: scrollRect
            ? { top: scrollRect.top, bottom: scrollRect.bottom, height: scrollRect.height }
            : null,
          formulaLine: lineRect
            ? {
                top: lineRect.top,
                bottom: lineRect.bottom,
                height: lineRect.height,
                background: lineStyle?.backgroundColor ?? "",
                boxShadow: lineStyle?.boxShadow ?? "",
              }
            : null,
          field: fieldRect
            ? { top: fieldRect.top, bottom: fieldRect.bottom, height: fieldRect.height, focused: field?.hasFocus?.() ?? false, position: field?.position ?? -1, lastOffset: field?.lastOffset ?? -1 }
            : null,
          bottomTabs: bottomRect
            ? { top: bottomRect.top, bottom: bottomRect.bottom, height: bottomRect.height }
            : null,
          inputBehavior: {
            text: inputBehaviorLabel?.textContent?.trim() ?? "",
            labelDisplay:
              inputBehaviorLabel instanceof HTMLElement
                ? getComputedStyle(inputBehaviorLabel).display
                : "",
            arrowDisplay:
              inputBehaviorArrow instanceof SVGElement
                ? getComputedStyle(inputBehaviorArrow).display
                : "",
          },
          bodyOverflowX: document.documentElement.scrollWidth - window.innerWidth,
          lineCount: document.querySelectorAll(".formula-line").length,
          emptyVisualClassCount: document.querySelectorAll(".formula-line.is-empty").length,
          modeToggleCount: document.querySelectorAll("[data-formula-line-mode-toggle]").length,
        };
      })()`,
      "single-row Office header and visible empty formula editor",
    );

    assert.equal(layout.visibleHeaderRows, 1);
    assert.equal(layout.allHeaderInside, true);
    assert.equal(layout.lineCount, 1);
    assert.equal(layout.emptyVisualClassCount, 0);
    assert.equal(layout.formulaLine.boxShadow, "none");
    assert.match(layout.inputBehavior.text, /操作逻辑|Input behavior/);
    assert.notEqual(layout.inputBehavior.labelDisplay, "none");
    assert.notEqual(layout.inputBehavior.arrowDisplay, "none");
    assert.equal(layout.modeToggleCount, 0);
    assert.ok(layout.formulaLine.height >= 40);
    assert.ok(layout.field.height >= 30);
    assert.ok(layout.bodyOverflowX <= 1);

    console.log(JSON.stringify(layout, null, 2));
    console.log(
      "Windows Office layout regression passed: single header row, normal formula field geometry, no desktop-shell bleed.",
    );
  } finally {
    client?.close();
    if (chrome && !chrome.killed) chrome.kill();
    await new Promise((resolve) => server.close(resolve));
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => {});
  }
}

await main();
