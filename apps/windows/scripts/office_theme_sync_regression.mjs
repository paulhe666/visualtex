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
const offset = process.pid % 700;
const port = 18100 + offset;
const debugPort = 19100 + offset;
const baseUrl = `http://127.0.0.1:${port}`;
const officeUrl = `${baseUrl}/dialog/theme-regression?runtime=vsto-desktop`;
const distRoot = join(process.cwd(), "dist-office-windows-native");
const chromeProfile = createBrowserProfilePath("visualtex-windows-office-theme");
const chromePath = resolveChromiumExecutable();
let currentTheme = "purple";
let currentEditorLayout = "classic";
let currentSessionHost = "word";
let lastSessionPatch = null;
let currentEditorPreferences = {
  zoom: 0.8,
  formulaInsetLeft: 31,
  formulaInsetRight: 37,
  formulaToolButtonSize: 52,
  formulaToolButtonPadding: 7,
  formulaRowVerticalInset: 9,
  formulaLetterFont: "times",
  formulaChineseFont: "songti",
  keypadMinimizeOnCopy: false,
};

function sessionPayload(update = {}) {
  return {
    id: "theme-regression",
    mode: "edit",
    host: currentSessionHost,
    formulaId: "theme-regression-formula",
    sourceDocumentId: "theme-regression-document",
    sourceObjectId: "theme-regression-object",
    title: "Office compact geometry regression",
    lines: [{ id: "line-1", latex: "E=mc^2,\\quad \\alpha+\\beta=\\gamma,\\quad \\text{中文公式}" }],
    activeLineId: "line-1",
    codeFormat: "raw",
    displayMode: "inline",
    objectMode: "nativeOle",
    numbered: false,
    fontSizePt: 10.5,
    exportWidth: 320,
    exportHeight: 80,
    exportResult: null,
    originalMetadata: null,
    dirty: false,
    status: "editing",
    autoCommitOnClose: true,
    explicitCancel: false,
    error: null,
    createdAt: Date.now(),
    updatedAt: Date.now(),
    expiresAt: Date.now() + 600000,
    ...update,
  };
}

const mimeTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".woff2": "font/woff2",
};

function writeJson(response, status, value) {
  response.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
  });
  response.end(JSON.stringify(value));
}

const server = createServer(async (request, response) => {
  const url = new URL(request.url ?? "/", baseUrl);
  try {
    if (url.pathname === "/api/v1/theme") {
      writeJson(response, 200, {
        theme: currentTheme,
        editorLayout: currentEditorLayout,
      });
      return;
    }
    if (url.pathname === "/api/v1/preferences") {
      writeJson(response, 200, {
        powerpointDefaultFontSizePt: 20,
        editorPreferences: {
          settings: {
            theme: currentTheme,
            editorLayout: currentEditorLayout,
            ...currentEditorPreferences,
          },
        },
      });
      return;
    }
    if (url.pathname === "/api/v1/sessions/theme-regression") {
      if (request.method === "PATCH") {
        let raw = "";
        for await (const chunk of request) raw += chunk;
        lastSessionPatch = JSON.parse(raw || "{}");
        writeJson(response, 200, sessionPayload(lastSessionPatch));
        return;
      }
      writeJson(response, 200, sessionPayload());
      return;
    }
    if (url.pathname.startsWith("/api/v1/sessions/")) {
      writeJson(response, 404, { error: "Unknown theme regression session" });
      return;
    }
    if (url.pathname.startsWith("/dialog/")) {
      const source = await readFile(join(distRoot, "dialog", "index.html"), "utf8");
      const meta = [
        '<meta name="visualtex-install-token" content="theme-regression" />',
        `<meta name="visualtex-theme" content="${currentTheme}" />`,
        `<meta name="visualtex-editor-layout" content="${currentEditorLayout}" />`,
      ].join("\n");
      const html = source.replace("</head>", `${meta}\n</head>`);
      response.writeHead(200, {
        "Content-Type": "text/html; charset=utf-8",
        "Cache-Control": "no-store",
      });
      response.end(html);
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

async function waitFor(url, timeoutMs = 15_000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while the server or browser starts.
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

async function waitForPage() {
  const startedAt = Date.now();
  while (Date.now() - startedAt < 15_000) {
    const targets = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json();
    const target = targets.find(
      (candidate) => candidate.type === "page" && candidate.url.startsWith(officeUrl),
    );
    if (target) return target;
    await sleep(80);
  }
  throw new Error("Timed out waiting for the Office theme regression page");
}

async function waitForFontExportPatch(host, timeoutMs = 8_000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    const exportResult = lastSessionPatch?.exportResult;
    const svg = exportResult?.svg ?? "";
    if (
      exportResult?.formulaLetterFont === "helvetica" &&
      exportResult?.formulaChineseFont === "kaiti" &&
      svg.includes('data-visualtex-output-letter-font="helvetica"') &&
      svg.includes('data-visualtex-output-text-font="kaiti"')
    ) {
      const customLetterTextCount = (svg.match(/data-visualtex-output-letter-font=/g) ?? []).length;
      const customChineseTextCount = (svg.match(/data-visualtex-output-text-font=/g) ?? []).length;
      const remainingMathJaxUses = [...svg.matchAll(/<use\b([^>]*)>/gi)].map((match) => ({
        codePoint: match[1].match(/\bdata-c=["']([0-9A-F]+)["']/i)?.[1] ?? "",
        href: match[1].match(/\bxlink:href=["']#([^"']+)["']/i)?.[1] ?? "",
      }));
      const remainingMathJaxUseCount = remainingMathJaxUses.length;
      return {
        host,
        formulaLetterFont: exportResult.formulaLetterFont,
        formulaChineseFont: exportResult.formulaChineseFont,
        hasLetterMarker: true,
        hasChineseMarker: true,
        customLetterTextCount,
        customChineseTextCount,
        remainingMathJaxUseCount,
        remainingMathJaxUses,
        svgLength: svg.length,
      };
    }
    await sleep(100);
  }
  throw new Error(
    `Timed out waiting for ${host} Office OLE font export: ${JSON.stringify(lastSessionPatch)}`,
  );
}

async function readAppearance(client) {
  return client.evaluate(`(() => {
    let editorState = {};
    try {
      editorState = JSON.parse(localStorage.getItem('visualtex-editor') || '{}')?.state ?? {};
    } catch {}
    return {
      theme: document.documentElement.dataset.theme,
      editorLayout: editorState.editorLayout ?? null,
      zoom: editorState.zoom ?? null,
      formulaInsetLeft: editorState.formulaInsetLeft ?? null,
      formulaInsetRight: editorState.formulaInsetRight ?? null,
      formulaToolButtonSize: editorState.formulaToolButtonSize ?? null,
      formulaToolButtonPadding: editorState.formulaToolButtonPadding ?? null,
      formulaRowVerticalInset: editorState.formulaRowVerticalInset ?? null,
      formulaLetterFont: editorState.formulaLetterFont ?? null,
      formulaChineseFont: editorState.formulaChineseFont ?? null,
      keypadMinimizeOnCopy: editorState.keypadMinimizeOnCopy ?? null,
      background: getComputedStyle(document.documentElement).getPropertyValue('--bg').trim().toLowerCase(),
      surface: getComputedStyle(document.documentElement).getPropertyValue('--surface').trim().toLowerCase(),
      formulaSurface: getComputedStyle(document.documentElement).getPropertyValue('--formula-surface').trim().toLowerCase(),
      caret: getComputedStyle(document.documentElement).getPropertyValue('--formula-caret').trim().toLowerCase(),
    };
  })()`);
}

async function main() {
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(port, "127.0.0.1", resolve);
  });
  let chrome;
  let client;
  try {
    await waitFor(`${baseUrl}/api/v1/theme`);
    chrome = spawn(
      chromePath,
      [
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        `--remote-debugging-port=${debugPort}`,
        `--user-data-dir=${chromeProfile}`,
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

    let initialAppearance;
    for (let attempt = 0; attempt < 40; attempt += 1) {
      await sleep(100);
      initialAppearance = await readAppearance(client);
      if (
        initialAppearance.theme === "purple" &&
        initialAppearance.zoom === 0.8 &&
        initialAppearance.formulaInsetLeft === 31 &&
        initialAppearance.formulaInsetRight === 37 &&
        initialAppearance.formulaLetterFont === "times" &&
        initialAppearance.formulaChineseFont === "songti" &&
        initialAppearance.keypadMinimizeOnCopy === false
      ) break;
    }

    assert.deepEqual(initialAppearance, {
      theme: "purple",
      editorLayout: "classic",
      zoom: 0.8,
      formulaInsetLeft: 31,
      formulaInsetRight: 37,
      formulaToolButtonSize: 52,
      formulaToolButtonPadding: 7,
      formulaRowVerticalInset: 9,
      formulaLetterFont: "times",
      formulaChineseFont: "songti",
      keypadMinimizeOnCopy: false,
      background: "#120e16",
      surface: "#362842",
      formulaSurface: "#433252",
      caret: "#d7c2ff",
    });

    const officeCandidateLayer = await client.evaluate(`(() => {
      const probe = document.createElement("div");
      probe.id = "visualtex-native-input-suggestion-popover";
      probe.className = "is-visible";
      document.body.append(probe);
      const modalProbe = document.createElement("div");
      modalProbe.className = "modal-backdrop";
      document.body.append(modalProbe);
      const candidateZIndex = Number.parseInt(getComputedStyle(probe).zIndex || "0", 10) || 0;
      const modalZIndex = Number.parseInt(getComputedStyle(modalProbe).zIndex || "0", 10) || 0;
      const chromeZIndex = Math.max(
        0,
        ...[".formula-toolbar", ".source-panel", ".editor-pane-header", ".classic-bottom-tabs", ".classic-bottom-formatting-slot"]
          .map((selector) => document.querySelector(selector))
          .filter(Boolean)
          .map((element) => Number.parseInt(getComputedStyle(element).zIndex || "0", 10) || 0),
      );
      probe.remove();
      modalProbe.remove();
      return { candidateZIndex, chromeZIndex, modalZIndex };
    })()`);
    assert.ok(
      officeCandidateLayer.chromeZIndex < 100 &&
        officeCandidateLayer.candidateZIndex >= 190 &&
        officeCandidateLayer.candidateZIndex > officeCandidateLayer.chromeZIndex &&
        officeCandidateLayer.modalZIndex >= 300 &&
        officeCandidateLayer.modalZIndex > officeCandidateLayer.candidateZIndex,
      JSON.stringify(officeCandidateLayer),
    );

    await client.send("Emulation.setDeviceMetricsOverride", {
      width: 1280,
      height: 820,
      deviceScaleFactor: 1,
      mobile: false,
    });
    await sleep(220);
    const compactGeometry = await client.evaluate(`(() => {
      const slot = document.querySelector('.classic-bottom-formatting-slot')?.getBoundingClientRect();
      const tabs = document.querySelector('.classic-bottom-tabs')?.getBoundingClientRect();
      const group = document.querySelector('.classic-bottom-tab-group')?.getBoundingClientRect();
      const actions = document.querySelector('.classic-bottom-actions')?.getBoundingClientRect();
      const editorScroll = document.querySelector('.editor-pane-scroll');
      const editorSurface = document.querySelector('.editor-surface');
      const firstLine = document.querySelector('.formula-line');
      const editorScrollRect = editorScroll?.getBoundingClientRect();
      const editorSurfaceRect = editorSurface?.getBoundingClientRect();
      const firstLineRect = firstLine?.getBoundingClientRect();
      const editorSurfaceStyle = editorSurface ? getComputedStyle(editorSurface) : null;
      const rect = (value) => value ? ({ left: value.left, right: value.right, top: value.top, bottom: value.bottom, width: value.width, height: value.height }) : null;
      return {
        slot: rect(slot), tabs: rect(tabs), group: rect(group), actions: rect(actions), width: innerWidth, height: innerHeight,
        editorTop: {
          surfacePaddingTop: editorSurfaceStyle ? parseFloat(editorSurfaceStyle.paddingTop) : null,
          surfaceTopGap: editorSurfaceRect && editorScrollRect ? editorSurfaceRect.top - editorScrollRect.top : null,
          firstLineTopGap: firstLineRect && editorSurfaceRect ? firstLineRect.top - editorSurfaceRect.top : null,
        },
      };
    })()`);
    assert.ok(
      compactGeometry.slot &&
        compactGeometry.tabs &&
        compactGeometry.slot.left <= compactGeometry.tabs.left + 12 &&
        compactGeometry.slot.top >= compactGeometry.tabs.top - 1 &&
        compactGeometry.slot.bottom <= compactGeometry.tabs.bottom + 1,
      JSON.stringify(compactGeometry),
    );
    assert.ok(
      compactGeometry.editorTop.surfacePaddingTop >= 4 &&
        compactGeometry.editorTop.surfacePaddingTop <= 8 &&
        Math.abs(compactGeometry.editorTop.surfaceTopGap) <= 1 &&
        compactGeometry.editorTop.firstLineTopGap >= 0 &&
        compactGeometry.editorTop.firstLineTopGap <= 10,
      JSON.stringify(compactGeometry),
    );

    // A real OCR modal must cover every ordinary Office workspace layer. Keep
    // the input-behavior popover open first when possible so this also guards the
    // historical "operation logic is above the dialog" regression.
    await client.evaluate(`(() => {
      document.querySelector('.canvas-input-behavior-trigger')?.click();
      return true;
    })()`);
    await sleep(60);
    const operationPopoverWasOpen = await client.evaluate(
      `Boolean(document.querySelector('.input-behavior-popover'))`,
    );
    const ocrOpened = await client.evaluate(`(() => {
      const button = [...document.querySelectorAll('button')]
        .find((candidate) => candidate.textContent?.trim() === 'OCR');
      if (!(button instanceof HTMLButtonElement)) return false;
      button.click();
      return true;
    })()`);
    assert.equal(ocrOpened, true);
    await sleep(180);
    const ocrLayerState = await client.evaluate(`(() => {
      const backdrop = document.querySelector('.ocr-modal-backdrop');
      const dialog = document.querySelector('.ocr-dialog');
      const header = document.querySelector('.editor-pane-header.is-office-editor-header');
      const toolbar = document.querySelector('.formula-toolbar');
      const tabs = document.querySelector('.classic-bottom-tabs');
      const formatting = document.querySelector('.classic-bottom-formatting-slot');
      const operation = document.querySelector('.input-behavior-popover');
      const z = (element) => element
        ? Number.parseInt(getComputedStyle(element).zIndex || '0', 10) || 0
        : 0;
      const covered = (element) => {
        if (!element || !backdrop) return true;
        const bounds = element.getBoundingClientRect();
        if (bounds.width <= 0 || bounds.height <= 0) return true;
        const x = Math.max(1, Math.min(innerWidth - 2, bounds.left + Math.min(bounds.width / 2, 24)));
        const y = Math.max(1, Math.min(innerHeight - 2, bounds.top + Math.min(bounds.height / 2, 18)));
        const top = document.elementFromPoint(x, y);
        return Boolean(top?.closest('.ocr-modal-backdrop'));
      };
      const candidate = document.createElement('div');
      candidate.id = 'visualtex-native-input-suggestion-popover';
      candidate.className = 'is-visible';
      document.body.append(candidate);
      const context = document.createElement('div');
      context.className = 'formula-editor-context-menu';
      document.body.append(context);
      const result = {
        backdropExists: Boolean(backdrop),
        dialogExists: Boolean(dialog),
        backdropZ: z(backdrop),
        headerZ: z(header),
        toolbarZ: z(toolbar),
        tabsZ: z(tabs),
        formattingZ: z(formatting),
        operationZ: z(operation),
        candidateZ: z(candidate),
        contextZ: z(context),
        headerCovered: covered(header),
        toolbarCovered: covered(toolbar),
        tabsCovered: covered(tabs),
        formattingCovered: covered(formatting),
        operationCovered: covered(operation),
      };
      candidate.remove();
      context.remove();
      return result;
    })()`);
    assert.ok(
      operationPopoverWasOpen &&
        ocrLayerState.backdropExists &&
        ocrLayerState.dialogExists &&
        ocrLayerState.backdropZ >= 300 &&
        ocrLayerState.headerZ < 100 &&
        ocrLayerState.tabsZ < 100 &&
        ocrLayerState.formattingZ < 100 &&
        ocrLayerState.candidateZ < ocrLayerState.backdropZ &&
        ocrLayerState.contextZ < ocrLayerState.backdropZ &&
        ocrLayerState.headerCovered &&
        ocrLayerState.toolbarCovered &&
        ocrLayerState.tabsCovered &&
        ocrLayerState.formattingCovered &&
        ocrLayerState.operationCovered,
      JSON.stringify(ocrLayerState),
    );
    const ocrClosed = await client.evaluate(`(() => {
      const button = document.querySelector('button[aria-label="关闭 OCR"], button[aria-label="Close OCR"]');
      if (!(button instanceof HTMLButtonElement)) return false;
      button.click();
      return true;
    })()`);
    assert.equal(ocrClosed, true);
    await sleep(120);
    assert.equal(
      await client.evaluate(`Boolean(document.querySelector('.ocr-modal-backdrop'))`),
      false,
    );

    const zoomAfterUserClick = await client.evaluate(`(() => {
      const button = document.querySelector('button[aria-label="放大公式"], button[aria-label="Zoom in"]');
      if (!(button instanceof HTMLButtonElement)) return null;
      button.click();
      return true;
    })()`);
    assert.equal(zoomAfterUserClick, true);
    await sleep(120);
    const userAdjustedAppearance = await readAppearance(client);
    assert.equal(userAdjustedAppearance.zoom, 0.85);

    currentTheme = "green";
    currentEditorLayout = "standard";
    currentEditorPreferences = {
      ...currentEditorPreferences,
      zoom: 1.2,
      formulaInsetLeft: 18,
      formulaInsetRight: 22,
      formulaToolButtonSize: 60,
      formulaToolButtonPadding: 10,
      formulaRowVerticalInset: 13,
      formulaLetterFont: "helvetica",
      formulaChineseFont: "kaiti",
      keypadMinimizeOnCopy: true,
    };
    let synchronized;
    for (let attempt = 0; attempt < 20; attempt += 1) {
      await sleep(100);
      synchronized = await readAppearance(client);
      if (
        synchronized.theme === "green" &&
        synchronized.editorLayout === "standard" &&
        synchronized.zoom === 0.85 &&
        synchronized.formulaInsetLeft === 18 &&
        synchronized.formulaLetterFont === "helvetica"
      ) break;
    }
    assert.deepEqual(synchronized, {
      theme: "green",
      editorLayout: "standard",
      // Companion changes continue to sync normal editor preferences, but the
      // Office window's user-adjusted zoom must not snap back on the 500 ms poll.
      zoom: 0.85,
      formulaInsetLeft: 18,
      formulaInsetRight: 22,
      formulaToolButtonSize: 60,
      formulaToolButtonPadding: 10,
      formulaRowVerticalInset: 13,
      formulaLetterFont: "helvetica",
      formulaChineseFont: "kaiti",
      keypadMinimizeOnCopy: true,
      background: "#0d120f",
      surface: "#25352d",
      formulaSurface: "#2a3b32",
      caret: "#8bd4ac",
    });

    const wordOleFontExport = await waitForFontExportPatch("word");

    currentSessionHost = "powerpoint";
    lastSessionPatch = null;
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(350);
    const powerPointOleFontExport = await waitForFontExportPatch("powerpoint");

    process.stdout.write(
      `Windows Office theme/editor-layout inheritance and OLE font export regression passed: ${JSON.stringify({ wordOleFontExport, powerPointOleFontExport })}\n`,
    );
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    await new Promise((resolve) => server.close(resolve));
    await sleep(250);
    await rm(chromeProfile, {
      recursive: true,
      force: true,
      maxRetries: 4,
      retryDelay: 150,
    }).catch((error) => {
      if (process.platform !== "win32" || error?.code !== "EBUSY") throw error;
    });
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
