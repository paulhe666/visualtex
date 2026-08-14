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
const offset = process.pid % 600;
const port = 20200 + offset;
const debugPort = 21200 + offset;
const baseUrl = `http://127.0.0.1:${port}`;
const sessionId = "office-font-size-regression-session";
const officeUrl = `${baseUrl}/dialog/${sessionId}?runtime=vsto-desktop`;
const converterUrl = `${baseUrl}/dialog/${sessionId}?runtime=vsto-convert`;
const distRoot = join(process.cwd(), "dist-office-windows-native");
const chromeProfile = createBrowserProfilePath("visualtex-office-font-size");
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

function createSession(host, fontSizePt) {
  const lineId = `${host}-font-line`;
  return {
    id: sessionId,
    mode: "create",
    host,
    formulaId: `${host}-font-formula`,
    sourceDocumentId: `${host}-document`,
    sourceObjectId: null,
    title: host === "word" ? "Word Formula" : "PowerPoint Formula",
    lines: [{ id: lineId, latex: "x+y" }],
    activeLineId: lineId,
    codeFormat: "raw",
    displayMode: host === "word" ? "inline" : "block",
    numbered: false,
    fontSizePt,
    exportWidth: 0,
    exportHeight: 0,
    exportResult: null,
    originalMetadata: null,
    dirty: false,
    status: "created",
    autoCommitOnClose: true,
    explicitCancel: false,
    error: null,
    createdAt: Date.now(),
    updatedAt: Date.now(),
    expiresAt: Date.now() + 60_000,
  };
}

let session = createSession("powerpoint", 20);
const updates = [];
let completeCommits = false;
let closeRequests = 0;

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
      writeJson(response, 200, { theme: "light" });
      return;
    }
    if (url.pathname === "/api/v1/preferences") {
      writeJson(response, 200, {
        powerpointDefaultFontSizePt: 28,
        editorPreferences: {
          settings: {
            sourceOpen: false,
            classicTileWidth: 520,
            classicDockHeight: 420,
          },
        },
      });
      return;
    }
    if (url.pathname === `/api/v1/sessions/${sessionId}`) {
      if (request.method === "PATCH") {
        const update = await readJsonBody(request);
        updates.push(update);
        session = {
          ...session,
          ...update,
          updatedAt: Date.now(),
        };
        if (completeCommits && update.status === "committing") {
          session = {
            ...session,
            status: "completed",
            updatedAt: Date.now(),
          };
        }
      }
      writeJson(response, 200, session);
      return;
    }
    if (url.pathname === `/api/v1/app/sessions/${sessionId}/close`) {
      closeRequests += 1;
      writeJson(response, 200, { closed: true });
      return;
    }
    if (url.pathname.startsWith("/api/v1/ocr/")) {
      writeJson(response, 503, { error: "OCR is not needed in this regression" });
      return;
    }
    if (url.pathname.startsWith("/dialog/")) {
      const source = await readFile(join(distRoot, "dialog", "index.html"), "utf8");
      const meta = [
        '<meta name="visualtex-install-token" content="font-regression" />',
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

async function waitFor(url, timeoutMs = 15_000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while the local server or browser starts.
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
  throw new Error("Timed out waiting for Office font-size regression page");
}

async function waitForEvaluation(client, expression, description, timeoutMs = 10_000) {
  const startedAt = Date.now();
  let lastValue;
  while (Date.now() - startedAt < timeoutMs) {
    lastValue = await client.evaluate(expression);
    if (lastValue?.ready) return lastValue;
    await sleep(60);
  }
  throw new Error(`Timed out waiting for ${description}: ${JSON.stringify(lastValue)}`);
}

async function setFontSize(client, fontSizePt) {
  await client.evaluate(`(() => {
    const select = document.querySelector('[data-office-font-size]');
    if (!(select instanceof HTMLSelectElement)) return false;
    const setter = Object.getOwnPropertyDescriptor(
      HTMLSelectElement.prototype,
      'value',
    ).set;
    setter.call(select, ${JSON.stringify(String(fontSizePt))});
    select.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  })()`);
}

async function clickSelectorWithPointer(client, selector) {
  const point = await waitForEvaluation(
    client,
    `(() => {
      const element = document.querySelector(${JSON.stringify("__SELECTOR__")});
      if (!(element instanceof HTMLElement)) return { ready: false };
      const rect = element.getBoundingClientRect();
      return {
        ready: rect.width > 0 && rect.height > 0,
        x: rect.left + rect.width / 2,
        y: rect.top + rect.height / 2,
      };
    })()`.replace("__SELECTOR__", selector),
    `pointer target ${selector}`,
  );
  await client.send("Input.dispatchMouseEvent", {
    type: "mouseMoved",
    x: point.x,
    y: point.y,
    button: "none",
    buttons: 0,
  });
  await sleep(30);
  await client.send("Input.dispatchMouseEvent", {
    type: "mousePressed",
    x: point.x,
    y: point.y,
    button: "left",
    buttons: 1,
    clickCount: 1,
  });
  await client.send("Input.dispatchMouseEvent", {
    type: "mouseReleased",
    x: point.x,
    y: point.y,
    button: "left",
    buttons: 0,
    clickCount: 1,
  });
  await sleep(80);
}

async function dispatchOfficeShortcut(client, overrides = {}) {
  return client.evaluate(`(() => {
    const target = document.querySelector('math-field') ?? document.body;
    const event = new KeyboardEvent('keydown', {
      key: 's',
      code: 'KeyS',
      ctrlKey: true,
      bubbles: true,
      cancelable: true,
      ...${JSON.stringify(overrides)},
    });
    const dispatchResult = target.dispatchEvent(event);
    return {
      defaultPrevented: event.defaultPrevented,
      dispatchResult,
      lateCaptureCount: window.__visualtexLateSaveShortcutCount ?? 0,
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
        "--window-size=1240,820",
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

    const powerpointDefault = await waitForEvaluation(
      client,
      `(() => {
        const select = document.querySelector('[data-office-font-size]');
        const labels = [...(select?.querySelectorAll('option') ?? [])].map(
          (option) => option.textContent?.trim() ?? '',
        );
        const field = document.querySelector('math-field');
        return {
          ready:
            select instanceof HTMLSelectElement &&
            select.value === '28' &&
            select.querySelectorAll('optgroup').length === 2 &&
            labels.includes('小四（12 磅）') &&
            labels.includes('五号（10.5 磅）') &&
            labels.includes('初号（42 磅）') &&
            field?.position === field?.lastOffset,
          value: select?.value ?? '',
          selectedLabel: select?.selectedOptions[0]?.textContent?.trim() ?? '',
          optionGroupLabels: [...(select?.querySelectorAll('optgroup') ?? [])].map(
            (group) => group.label,
          ),
          host: document.querySelector('.office-dialog-header span')?.textContent ?? '',
          focused: field?.hasFocus?.() ?? false,
          position: field?.position ?? -1,
          lastOffset: field?.lastOffset ?? -1,
          documentFocused: document.hasFocus(),
          activeTag: document.activeElement?.tagName ?? '',
          activeClass: document.activeElement?.className ?? '',
          shadowActivePart:
            field?.shadowRoot?.activeElement?.getAttribute?.('part') ?? '',
        };
      })()`,
      "PowerPoint configured default font size with Chinese size options",
    );
    assert.equal(powerpointDefault.value, "28");
    assert.deepEqual(powerpointDefault.optionGroupLabels, ["中文字号", "磅值"]);

    await setFontSize(client, 12);
    const powerpointChineseSize = await waitForEvaluation(
      client,
      `(() => {
        const select = document.querySelector('[data-office-font-size]');
        return {
          ready:
            select?.value === '12' &&
            select?.selectedOptions[0]?.textContent?.includes('小四'),
          value: select?.value ?? '',
          selectedLabel: select?.selectedOptions[0]?.textContent?.trim() ?? '',
        };
      })()`,
      "PowerPoint Chinese font-size selection",
    );
    for (let attempt = 0; attempt < 80 && session.fontSizePt !== 12; attempt += 1) {
      await sleep(60);
    }
    assert.equal(session.fontSizePt, 12, "PowerPoint should persist 小四 as 12 pt");
    assert.ok(
      updates.some((update) => update.fontSizePt === 12),
      "PowerPoint autosave should include the selected Chinese size",
    );

    await setFontSize(client, 31.5);
    const powerpointSaved = await waitForEvaluation(
      client,
      `(() => ({
        ready: ${JSON.stringify(true)} && window.fetch !== undefined,
        value: document.querySelector('[data-office-font-size]')?.value ?? '',
      }))()`,
      "PowerPoint font input update",
    );
    assert.equal(powerpointSaved.value, "31.5");
    for (let attempt = 0; attempt < 80 && session.fontSizePt !== 31.5; attempt += 1) {
      await sleep(60);
    }
    assert.equal(session.fontSizePt, 31.5, "PowerPoint Session should persist selected size");
    assert.ok(
      updates.some((update) => update.fontSizePt === 31.5),
      "PowerPoint autosave should include selected font size",
    );

    await client.send("Page.reload", { ignoreCache: true });
    const powerpointReloaded = await waitForEvaluation(
      client,
      `(() => {
        const input = document.querySelector('[data-office-font-size]');
        return { ready: input?.value === '31.5', value: input?.value ?? '' };
      })()`,
      "PowerPoint Session font size after reload",
    );
    assert.equal(powerpointReloaded.value, "31.5");

    // Leave the PowerPoint page before replacing the mock Session. Its
    // pagehide autosave is valid for the old Session and must finish before
    // the test starts serving Word data.
    try {
      await client.send("Page.navigate", { url: "about:blank" });
    } catch (error) {
      if (!String(error).includes("navigated or closed")) throw error;
    }
    client.close();
    client = undefined;
    await sleep(250);
    session = createSession("word", 11.5);
    updates.length = 0;
    const wordTarget = await (
      await fetch(
        `http://127.0.0.1:${debugPort}/json/new?${encodeURIComponent(officeUrl)}`,
        { method: "PUT" },
      )
    ).json();
    client = new CdpClient(wordTarget.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    const wordInherited = await waitForEvaluation(
      client,
      `(() => {
        const input = document.querySelector('[data-office-font-size]');
        const field = document.querySelector('math-field');
        return {
          ready:
            input?.value === '11.5' &&
            field?.position === field?.lastOffset,
          value: input?.value ?? '',
          host: document.querySelector('.office-dialog-header span')?.textContent ?? '',
          focused: field?.hasFocus?.() ?? false,
          position: field?.position ?? -1,
          lastOffset: field?.lastOffset ?? -1,
        };
      })()`,
      "Word current paragraph font size",
    );
    assert.equal(wordInherited.value, "11.5");
    assert.ok(
      !updates.some((update) => update.dirty === true),
      "Word initial Session load must not persist a transient dirty state",
    );

    updates.length = 0;
    await client.evaluate(`(() => {
      const field = document.querySelector('math-field');
      if (!field) return false;
      field.setValue('\\\\frac{\\\\placeholder{}}{\\\\placeholder{}}', {
        mode: 'math',
        format: 'latex',
        insertionMode: 'replaceAll',
        selectionMode: 'placeholder',
        silenceNotifications: true,
      });
      field.dispatchEvent(new InputEvent('input', {
        bubbles: true,
        composed: true,
        inputType: 'insertText',
      }));
      return true;
    })()`);
    for (
      let attempt = 0;
      attempt < 100 &&
      !updates.some(
        (update) =>
          update.lines?.[0]?.latex?.includes('\\placeholder{}') &&
          update.exportResult === null &&
          update.error === null,
      );
      attempt += 1
    ) {
      await sleep(60);
    }
    assert.ok(
      updates.some(
        (update) =>
          update.lines?.[0]?.latex?.includes('\\placeholder{}') &&
          update.exportResult === null &&
          update.error === null,
      ),
      "Word placeholder draft should autosave without a stale export result",
    );
    const placeholderDraftState = await client.evaluate(`(() => ({
      value: document.querySelector('math-field')?.value ?? '',
      toast: document.querySelector('.toast')?.textContent?.trim() ?? '',
    }))()`);
    assert.match(placeholderDraftState.value, /\\placeholder/);
    assert.equal(
      placeholderDraftState.toast,
      "",
      "An incomplete placeholder draft must not show a MathJax error toast",
    );

    await client.evaluate(`(() => {
      const field = document.querySelector('math-field');
      if (!field) return false;
      field.setValue('x+y', {
        mode: 'math',
        format: 'latex',
        insertionMode: 'replaceAll',
        selectionMode: 'after',
        silenceNotifications: true,
      });
      field.dispatchEvent(new InputEvent('input', {
        bubbles: true,
        composed: true,
        inputType: 'insertText',
      }));
      return true;
    })()`);
    for (
      let attempt = 0;
      attempt < 100 && session.lines?.[0]?.latex !== "x+y";
      attempt += 1
    ) {
      await sleep(60);
    }
    assert.equal(session.lines?.[0]?.latex, "x+y");

    await setFontSize(client, 10.5);
    const wordChineseSize = await waitForEvaluation(
      client,
      `(() => {
        const select = document.querySelector('[data-office-font-size]');
        return {
          ready:
            select?.value === '10.5' &&
            select?.selectedOptions[0]?.textContent?.includes('五号'),
          value: select?.value ?? '',
          selectedLabel: select?.selectedOptions[0]?.textContent?.trim() ?? '',
        };
      })()`,
      "Word Chinese font-size selection",
    );
    for (let attempt = 0; attempt < 80 && session.fontSizePt !== 10.5; attempt += 1) {
      await sleep(60);
    }
    assert.equal(session.fontSizePt, 10.5, "Word should persist 五号 as 10.5 pt");
    assert.ok(
      updates.some((update) => update.fontSizePt === 10.5),
      "Word autosave should include the selected Chinese size",
    );

    await setFontSize(client, 13);
    for (let attempt = 0; attempt < 80 && session.fontSizePt !== 13; attempt += 1) {
      await sleep(60);
    }
    assert.equal(session.fontSizePt, 13, "Word Session should persist selected size");
    assert.ok(
      updates.some((update) => update.fontSizePt === 13),
      "Word autosave should include selected font size",
    );

    const officeControlBarState = await waitForEvaluation(
      client,
      `(() => {
        const bar = document.querySelector('.editor-pane-header.is-office-editor-header');
        const tileTabs = document.querySelector('.classic-tile-toolbar .formula-tile-tabs');
        const cancel = bar?.querySelector('.office-inline-cancel');
        const primary = bar?.querySelector('.office-inline-primary');
        const size = bar?.querySelector('[data-office-font-size]');
        const display = bar?.querySelector('.office-display-mode-setting');
        const canvasTools = bar?.querySelector('.canvas-tool-group');
        const rects = [cancel, primary, size, display]
          .filter((element) => element instanceof HTMLElement)
          .map((element) => element.getBoundingClientRect());
        const centers = rects.map((rect) => rect.top + rect.height / 2);
        const sameRow = centers.length === 4 &&
          Math.max(...centers) - Math.min(...centers) <= 8;
        const barRect = bar?.getBoundingClientRect();
        const tileRect = tileTabs?.getBoundingClientRect();
        const aboveTiles = Boolean(
          barRect && tileRect && barRect.bottom <= tileRect.top + 1,
        );
        const text = bar?.textContent ?? '';
        const paneHeaderVisible = bar
          ? getComputedStyle(bar).display !== 'none'
          : false;
        return {
          ready: Boolean(
            bar &&
            tileTabs &&
            cancel &&
            primary &&
            size &&
            display &&
            canvasTools &&
            sameRow &&
            aboveTiles &&
            paneHeaderVisible &&
            !text.includes('VisualTeX') &&
            !text.includes('Microsoft Word') &&
            !text.includes('新建 Office 公式') &&
            !text.includes('编辑所选公式')
          ),
          sameRow,
          aboveTiles,
          paneHeaderVisible,
          barHeight: barRect?.height ?? 0,
          barBottom: barRect?.bottom ?? 0,
          tileTop: tileRect?.top ?? 0,
          text,
        };
      })()`,
      "Office editor header above formula tile tabs",
    );
    assert.equal(officeControlBarState.sameRow, true);
    assert.equal(officeControlBarState.aboveTiles, true);
    assert.equal(officeControlBarState.paneHeaderVisible, true);

    await client.evaluate(`(() => {
      document.querySelector('[data-formula-tile-collapse]')?.click();
      document.querySelector('[data-classic-bottom-collapse]')?.click();
      return true;
    })()`);
    const collapsedPreferenceState = await waitForEvaluation(
      client,
      `(() => ({
        ready:
          localStorage.getItem('visualtex-office-editor-tiles-open') === 'false' &&
          localStorage.getItem('visualtex-office-editor-toolbar-open') === 'false',
        tiles: localStorage.getItem('visualtex-office-editor-tiles-open'),
        toolbar: localStorage.getItem('visualtex-office-editor-toolbar-open'),
      }))()`,
      "persist collapsed Office editor panels",
    );
    assert.equal(collapsedPreferenceState.tiles, 'false');
    assert.equal(collapsedPreferenceState.toolbar, 'false');
    await client.evaluate(`location.reload()`);
    const restoredCollapsedLayout = await waitForEvaluation(
      client,
      `(() => ({
        ready: Boolean(
          document.querySelector('.editor-pane-header.is-office-editor-header') &&
          document.querySelector('.classic-tile-expand-button') &&
          document.querySelector('.classic-bottom-dock.is-collapsed') &&
          !document.querySelector('.classic-tile-toolbar')
        ),
        tileToolbarPresent: Boolean(document.querySelector('.classic-tile-toolbar')),
        tileExpandPresent: Boolean(document.querySelector('.classic-tile-expand-button')),
        bottomCollapsed: Boolean(document.querySelector('.classic-bottom-dock.is-collapsed')),
      }))()`,
      "restore collapsed Office editor panels after reload",
    );
    assert.equal(restoredCollapsedLayout.tileToolbarPresent, false);
    assert.equal(restoredCollapsedLayout.tileExpandPresent, true);
    assert.equal(restoredCollapsedLayout.bottomCollapsed, true);
    await client.evaluate(`(() => {
      document.querySelector('.classic-tile-expand-button')?.click();
      document.querySelector('[data-classic-bottom-collapse]')?.click();
      return true;
    })()`);
    await waitForEvaluation(
      client,
      `(() => ({
        ready:
          localStorage.getItem('visualtex-office-editor-tiles-open') === 'true' &&
          localStorage.getItem('visualtex-office-editor-toolbar-open') === 'true' &&
          Boolean(document.querySelector('.classic-tile-toolbar')) &&
          !document.querySelector('.classic-bottom-dock.is-collapsed'),
      }))()`,
      "persist reopened Office editor panels",
    );

    await client.evaluate(`localStorage.removeItem('visualtex-office-editor-source-open')`);
    await client.evaluate(`location.reload()`);
    const defaultOfficeTools = await waitForEvaluation(
      client,
      `(() => ({
        ready: Boolean(document.querySelector('[data-classic-bottom-view="tools"]')),
        toolsSelected: document.querySelector('[data-classic-bottom-view="tools"]')?.getAttribute('aria-selected') === 'true',
        sourceSelected: document.querySelector('[data-classic-bottom-view="source"]')?.getAttribute('aria-selected') === 'true',
      }))()`,
      "default Office editor to formula tools",
    );
    assert.equal(defaultOfficeTools.toolsSelected && !defaultOfficeTools.sourceSelected, true);

    await clickSelectorWithPointer(client, '[data-classic-bottom-view=source]');
    await sleep(900);
    const stableSourceTab = await client.evaluate(`(() => {
      const source = document.querySelector('[data-classic-bottom-view="source"]');
      const tools = document.querySelector('[data-classic-bottom-view="tools"]');
      const dock = document.querySelector('.classic-bottom-dock');
      return {
        sourceSelected: source?.getAttribute('aria-selected') === 'true',
        toolsSelected: tools?.getAttribute('aria-selected') === 'true',
        sourcePanel: dock?.classList.contains('is-source-panel') === true,
        preference: localStorage.getItem('visualtex-office-editor-source-open'),
      };
    })()`);
    assert.equal(
      stableSourceTab.sourceSelected && !stableSourceTab.toolsSelected && stableSourceTab.sourcePanel,
      true,
      'Office Formula tools / LaTeX source tab must remain on the user-selected source view beyond the 500ms companion preference poll',
    );
    assert.equal(stableSourceTab.preference, 'true');
    await client.evaluate(`location.reload()`);
    const restoredOfficeSource = await waitForEvaluation(
      client,
      `(() => ({
        ready: Boolean(document.querySelector('[data-classic-bottom-view="source"]')),
        sourceSelected: document.querySelector('[data-classic-bottom-view="source"]')?.getAttribute('aria-selected') === 'true',
      }))()`,
      "restore Office source tab after reload",
    );
    assert.equal(restoredOfficeSource.sourceSelected, true);
    await clickSelectorWithPointer(client, '[data-classic-bottom-view=tools]');
    await sleep(120);
    assert.equal(
      await client.evaluate(`localStorage.getItem('visualtex-office-editor-source-open')`),
      'false',
    );
    await client.evaluate(`location.reload()`);
    const restoredOfficeTools = await waitForEvaluation(
      client,
      `(() => ({
        ready: Boolean(document.querySelector('[data-classic-bottom-view="tools"]')),
        toolsSelected: document.querySelector('[data-classic-bottom-view="tools"]')?.getAttribute('aria-selected') === 'true',
      }))()`,
      "restore Office formula-tools tab after reload",
    );
    assert.equal(restoredOfficeTools.toolsSelected, true);

    const resizedLayoutPreference = await client.evaluate(`(() => {
      const tile = document.querySelector('.classic-tile-resizer');
      const dock = document.querySelector('.classic-dock-resizer');
      if (!(tile instanceof HTMLElement) || !(dock instanceof HTMLElement)) return null;
      const beforeTile = Number(tile.getAttribute('aria-valuenow'));
      const beforeDock = Number(dock.getAttribute('aria-valuenow'));
      tile.dispatchEvent(new KeyboardEvent('keydown', {
        key: 'ArrowLeft', bubbles: true, cancelable: true,
      }));
      dock.dispatchEvent(new KeyboardEvent('keydown', {
        key: 'ArrowUp', bubbles: true, cancelable: true,
      }));
      return { beforeTile, beforeDock };
    })()`);
    assert.ok(resizedLayoutPreference);
    const persistedPanelSizes = await waitForEvaluation(
      client,
      `(() => {
        const tile = Number(localStorage.getItem('visualtex-classic-tile-width'));
        const dock = Number(localStorage.getItem('visualtex-classic-dock-height'));
        return {
          ready:
            Number.isFinite(tile) &&
            Number.isFinite(dock) &&
            tile > ${resizedLayoutPreference.beforeTile} &&
            dock > ${resizedLayoutPreference.beforeDock},
          tile,
          dock,
        };
      })()`,
      "persist resized Office editor panels",
    );
    await sleep(900);
    const stablePanelSizes = await client.evaluate(`(() => {
      const tile = Number(document.querySelector('.classic-tile-resizer')?.getAttribute('aria-valuenow'));
      const dock = Number(document.querySelector('.classic-dock-resizer')?.getAttribute('aria-valuenow'));
      return { tile, dock };
    })()`);
    assert.ok(
      Math.abs(stablePanelSizes.tile - persistedPanelSizes.tile) < 1 &&
        Math.abs(stablePanelSizes.dock - persistedPanelSizes.dock) < 1,
      `Office panel sizes were overwritten by companion polling: ${JSON.stringify({ persistedPanelSizes, stablePanelSizes })}`,
    );
    await client.evaluate(`location.reload()`);
    const restoredPanelSizes = await waitForEvaluation(
      client,
      `(() => {
        const workspace = document.querySelector('.workspace.is-classic-layout');
        const body = document.querySelector('.classic-editor-pane-body');
        const tile = Number.parseFloat(
          workspace?.style.getPropertyValue('--classic-tile-width') || '0',
        );
        const dock = Number.parseFloat(
          body?.style.getPropertyValue('--classic-dock-height') || '0',
        );
        return {
          ready:
            Math.abs(tile - ${persistedPanelSizes.tile}) < 1 &&
            Math.abs(dock - ${persistedPanelSizes.dock}) < 1,
          tile,
          dock,
        };
      })()`,
      "restore resized Office editor panels after reload",
    );
    assert.ok(Math.abs(restoredPanelSizes.tile - persistedPanelSizes.tile) < 1);
    assert.ok(Math.abs(restoredPanelSizes.dock - persistedPanelSizes.dock) < 1);

    const commonBefore = await waitForEvaluation(
      client,
      `(() => {
        const toolbar = document.querySelector('.formula-toolbar');
        const commonTab = document.querySelector('[data-category="common"]');
        if (!toolbar && document.querySelector('.sidebar-toggle')) {
          document.querySelector('.sidebar-toggle').click();
        }
        if (commonTab && commonTab.getAttribute('aria-pressed') !== 'true') {
          commonTab.click();
        }
        const buttons = [...document.querySelectorAll(
          '[data-toolbar-category-section="common"] > [data-command-id]',
        )];
        const visibility = buttons.map((button) => {
          const host = button.querySelector('.math-preview');
          const fit = button.querySelector('.math-preview-fit-content');
          const latex = button.querySelector('.ML__latex');
          const hostRect = host?.getBoundingClientRect();
          const fitRect = fit?.getBoundingClientRect();
          const latexRect = latex?.getBoundingClientRect();
          const hostStyle = host ? getComputedStyle(host) : null;
          const fitStyle = fit ? getComputedStyle(fit) : null;
          const latexStyle = latex ? getComputedStyle(latex) : null;
          return {
            id: button.dataset.commandId ?? '',
            hostWidth: hostRect?.width ?? 0,
            hostHeight: hostRect?.height ?? 0,
            fitWidth: fitRect?.width ?? 0,
            fitHeight: fitRect?.height ?? 0,
            latexWidth: latexRect?.width ?? 0,
            latexHeight: latexRect?.height ?? 0,
            hostDisplay: hostStyle?.display ?? '',
            hostVisibility: hostStyle?.visibility ?? '',
            hostOpacity: hostStyle?.opacity ?? '',
            fitDisplay: fitStyle?.display ?? '',
            latexDisplay: latexStyle?.display ?? '',
            latexVisibility: latexStyle?.visibility ?? '',
            latexOpacity: latexStyle?.opacity ?? '',
            latexColor: latexStyle?.color ?? '',
          };
        });
        const invalid = visibility.filter((item) =>
          item.hostWidth <= 0 ||
          item.hostHeight <= 0 ||
          item.latexWidth <= 0 ||
          item.latexHeight <= 0 ||
          item.hostDisplay === 'none' ||
          item.hostVisibility === 'hidden' ||
          Number(item.hostOpacity) === 0 ||
          item.latexDisplay === 'none' ||
          item.latexVisibility === 'hidden' ||
          Number(item.latexOpacity) === 0 ||
          item.latexColor === 'rgba(0, 0, 0, 0)'
        );
        return {
          ready:
            buttons.length === 45 &&
            buttons[0]?.dataset.commandId === 'frac' &&
            buttons.at(-1)?.dataset.commandId === 'exists' &&
            buttons.some((button) => button.dataset.commandId === 'times') &&
            buttons.some((button) => button.dataset.commandId === 'div') &&
            !buttons.some((button) => button.dataset.commandId === 'notin') &&
            !buttons.some((button) => button.dataset.commandId === 'leftarrow') &&
            invalid.length === 0,
          count: buttons.length,
          ids: buttons.map((button) => button.dataset.commandId ?? ''),
          invalid,
          firstVisual: visibility[0] ?? null,
        };
      })()`,
      "fixed 45-item common toolbar",
    );
    assert.equal(commonBefore.count, 45);
    assert.equal(commonBefore.ids[0], "frac");
    assert.equal(commonBefore.ids.at(-1), "exists");
    assert.ok(commonBefore.ids.includes("times"));
    assert.ok(commonBefore.ids.includes("div"));
    assert.ok(!commonBefore.ids.includes("notin"));
    assert.ok(!commonBefore.ids.includes("leftarrow"));

    const toolbarPreviewContainment = await waitForEvaluation(
      client,
      `(() => {
        const buttons = [...document.querySelectorAll('.formula-toolbar .template-button[data-command-id]')]
          .filter((button) => {
            const rect = button.getBoundingClientRect();
            const style = getComputedStyle(button);
            return rect.width > 0 && rect.height > 0 && style.display !== 'none' && style.visibility !== 'hidden';
          });
        const tolerance = 1.5;
        const details = buttons.map((button) => {
          const buttonRect = button.getBoundingClientRect();
          const fit = button.querySelector('.math-preview-fit-content');
          const latex = button.querySelector('.ML__latex');
          const fitRect = fit?.getBoundingClientRect();
          const latexRect = latex?.getBoundingClientRect();
          const fits = Boolean(
            fitRect &&
            fitRect.left >= buttonRect.left - tolerance &&
            fitRect.right <= buttonRect.right + tolerance &&
            fitRect.top >= buttonRect.top - tolerance &&
            fitRect.bottom <= buttonRect.bottom + tolerance,
          );
          const latexFits = Boolean(
            latexRect &&
            latexRect.left >= buttonRect.left - tolerance &&
            latexRect.right <= buttonRect.right + tolerance &&
            latexRect.top >= buttonRect.top - tolerance &&
            latexRect.bottom <= buttonRect.bottom + tolerance,
          );
          return {
            id: button.dataset.commandId ?? '',
            fits,
            latexFits,
            button: { left: buttonRect.left, right: buttonRect.right, top: buttonRect.top, bottom: buttonRect.bottom },
            fit: fitRect ? { left: fitRect.left, right: fitRect.right, top: fitRect.top, bottom: fitRect.bottom } : null,
            latex: latexRect ? { left: latexRect.left, right: latexRect.right, top: latexRect.top, bottom: latexRect.bottom } : null,
            scale: fit ? getComputedStyle(fit).transform : '',
          };
        });
        const overflow = details.filter((item) => !item.fits || !item.latexFits);
        return {
          ready: buttons.length > 20 && overflow.length === 0,
          count: buttons.length,
          overflow,
        };
      })()`,
      "Office formula toolbar previews contained by their tiles",
    );
    assert.equal(toolbarPreviewContainment.overflow.length, 0, JSON.stringify(toolbarPreviewContainment.overflow));

    const arithmeticOperatorState = await waitForEvaluation(
      client,
      `(() => {
        const field = document.querySelector("math-field");
        if (!field?.isConnected) return { ready: false, value: "" };
        if (!window.__visualtexArithmeticOperatorsStarted) {
          window.__visualtexArithmeticOperatorsStarted = true;
          field.setValue("", {
            mode: "math",
            format: "latex",
            insertionMode: "replaceAll",
            selectionMode: "after",
            silenceNotifications: true,
          });
          field.position = field.lastOffset;
          document.querySelector('[data-command-id="times"]')?.click();
          document.querySelector('[data-command-id="div"]')?.click();
        }
        const value = field.value;
        return {
          ready: value.includes("\\\\times") && value.includes("\\\\div"),
          value,
        };
      })()`,
      "common multiplication and division insertion",
    );
    assert.match(arithmeticOperatorState.value, /\\times/);
    assert.match(arithmeticOperatorState.value, /\\div/);
    await client.evaluate(`(() => {
      const field = document.querySelector("math-field");
      field?.setValue("abc", {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
      if (field) field.position = field.lastOffset;
      return true;
    })()`);

    await client.evaluate(`(() => {
      document.querySelector('[data-category="matrix"]')?.click();
      return true;
    })()`);
    await waitForEvaluation(
      client,
      `(() => ({
        ready: Boolean(document.querySelector('[data-command-id="blackboard-bold"]')),
      }))()`,
      "matrix toolbar category",
    );
    const contextMenuTriggered = await client.evaluate(`(() => {
      const button = document.querySelector('[data-command-id="blackboard-bold"]');
      if (!(button instanceof HTMLElement)) return false;
      button.dispatchEvent(new MouseEvent('contextmenu', {
        bubbles: true,
        cancelable: true,
        clientX: 320,
        clientY: 240,
        button: 2,
      }));
      return true;
    })()`);
    assert.equal(contextMenuTriggered, true);
    await waitForEvaluation(
      client,
      `(() => ({
        ready: Boolean(
          document.querySelector('[data-add-to-common-command="blackboard-bold"]'),
        ),
      }))()`,
      "Add to Common context-menu action",
    );
    await client.evaluate(`(() => {
      const action = document.querySelector(
        '[data-add-to-common-command="blackboard-bold"]',
      );
      if (!(action instanceof HTMLElement)) return false;
      action.click();
      return true;
    })()`);
    const commonAfter = await waitForEvaluation(
      client,
      `(() => {
        const buttons = [...document.querySelectorAll(
          '[data-toolbar-category-section="common"] > [data-command-id]',
        )];
        const stored = JSON.parse(
          localStorage.getItem('visualtex-common-toolbar-command-ids-v2') || '[]',
        );
        return {
          ready:
            buttons.length === 45 &&
            buttons[0]?.dataset.commandId === 'blackboard-bold' &&
            !buttons.some((button) => button.dataset.commandId === 'exists') &&
            stored.length === 45 &&
            stored[0] === 'blackboard-bold',
          count: buttons.length,
          first: buttons[0]?.dataset.commandId ?? '',
          last: buttons.at(-1)?.dataset.commandId ?? '',
          includesEjectedDefault: buttons.some(
            (button) => button.dataset.commandId === 'exists',
          ),
          stored,
        };
      })()`,
      "manual common command promotion",
    );
    assert.equal(commonAfter.count, 45);
    assert.equal(commonAfter.first, "blackboard-bold");
    assert.equal(commonAfter.includesEjectedDefault, false);
    assert.equal(commonAfter.stored.length, 45);
    await client.evaluate(`(() => {
      document.querySelector(
        '[data-toolbar-category-section="common"] [data-command-id="sqrt"]',
      )?.click();
      return true;
    })()`);
    await sleep(100);
    const commonAfterUse = await client.evaluate(`(() => {
      const buttons = [...document.querySelectorAll(
        '[data-toolbar-category-section="common"] > [data-command-id]',
      )];
      return {
        count: buttons.length,
        first: buttons[0]?.dataset.commandId ?? '',
      };
    })()`);
    assert.equal(commonAfterUse.count, 45);
    assert.equal(
      commonAfterUse.first,
      "blackboard-bold",
      "Using another common command must not reorder the Common category",
    );

    const formattingControls = await waitForEvaluation(
      client,
      `(() => {
        const selectionBold = document.querySelector('[data-formula-selection-bold]');
        const selectionItalic = document.querySelector('[data-formula-selection-italic]');
        const color = document.querySelector('[data-formula-selection-color]');
        const background = document.querySelector(
          '[data-formula-selection-background]',
        );
        const formattingMount = document.querySelector('.classic-bottom-formatting-slot');
        const bottomTabs = document.querySelector('.classic-bottom-tab-group');
        const formattingBounds = formattingMount?.getBoundingClientRect();
        const bottomTabBounds = bottomTabs?.getBoundingClientRect();
        const formattingAtFarLeft = Boolean(
          formattingBounds &&
          bottomTabBounds &&
          formattingBounds.right <= bottomTabBounds.left + 2,
        );
        return {
          ready: Boolean(
            selectionBold &&
            selectionItalic &&
            color &&
            background &&
            formattingMount &&
            bottomTabs &&
            formattingAtFarLeft
          ),
          persistentControlsAbsent:
            !document.querySelector('[data-formula-typing-bold]') &&
            !document.querySelector('[data-formula-typing-italic]'),
          formattingAtFarLeft,
          formattingLeft: formattingBounds?.left ?? 0,
          formattingRight: formattingBounds?.right ?? 0,
          bottomTabsLeft: bottomTabBounds?.left ?? 0,
        };
      })()`,
      "Office formatting controls at the left of Formula tools row",
    );
    assert.equal(formattingControls.persistentControlsAbsent, true);
    assert.equal(formattingControls.formattingAtFarLeft, true);

    await client.send('Emulation.setDeviceMetricsOverride', {
      width: 500,
      height: 760,
      deviceScaleFactor: 1,
      mobile: false,
    });
    await sleep(120);
    await client.evaluate(`(() => {
      const field = document.querySelector('math-field');
      if (!field) return false;
      field.setValue('abc', {
        mode: 'math',
        format: 'latex',
        insertionMode: 'replaceAll',
        selectionMode: 'after',
      });
      field.focus();
      field.selection = { ranges: [[0, field.lastOffset]], direction: 'forward' };
      return true;
    })()`);
    await clickSelectorWithPointer(client, '[data-formula-selection-color]');
    await waitForEvaluation(
      client,
      `(() => ({ ready: Boolean(document.querySelector('[data-formula-color-popover="color"]')) }))()`,
      'Office formula color popover mounted',
    );
    await client.evaluate(`window.dispatchEvent(new Event('resize'))`);
    const colorPopoverLayout = await waitForEvaluation(
      client,
      `(() => {
        const popover = document.querySelector('[data-formula-color-popover="color"]');
        if (!(popover instanceof HTMLElement)) return { ready: false };
        const rect = popover.getBoundingClientRect();
        const dock = document.querySelector('.classic-bottom-dock');
        const body = document.querySelector('.classic-editor-pane-body');
        const tabs = document.querySelector('.classic-bottom-tabs');
        const visible =
          rect.left >= -1 &&
          rect.top >= -1 &&
          rect.right <= window.innerWidth + 1 &&
          rect.bottom <= window.innerHeight + 1;
        return {
          ready: visible && popover.dataset.visualtexAutoAvoidAdjusted === 'true',
          visible,
          autoAdjusted: popover.dataset.visualtexAutoAvoidAdjusted === 'true',
          matchesAutoSelector: popover.matches('[data-visualtex-floating-layer]'),
          rect: { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom },
          dockOverflow: dock ? getComputedStyle(dock).overflow : '',
          bodyOverflow: body ? getComputedStyle(body).overflow : '',
          tabsOverflow: tabs ? getComputedStyle(tabs).overflow : '',
        };
      })()`,
      "unclipped Office formula color popover",
    );
    assert.equal(colorPopoverLayout.visible, true);
    assert.equal(colorPopoverLayout.autoAdjusted, true);
    assert.equal(colorPopoverLayout.dockOverflow, 'visible');
    assert.equal(colorPopoverLayout.bodyOverflow, 'visible');
    assert.equal(colorPopoverLayout.tabsOverflow, 'visible');
    await client.evaluate(`(() => {
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
      document.querySelector('[data-classic-bottom-view="tools"]')?.click();
      return true;
    })()`);
    const compactToolbarLayout = await waitForEvaluation(
      client,
      `(() => {
        const toolbar = document.querySelector('.classic-bottom-toolbar');
        const tile = toolbar?.querySelector('.template-button');
        const tab = toolbar?.querySelector('.toolbar-tab');
        if (!(toolbar instanceof HTMLElement) || !(tile instanceof HTMLElement) || !(tab instanceof HTMLElement)) {
          return { ready: false };
        }
        const tileRect = tile.getBoundingClientRect();
        const tabRect = tab.getBoundingClientRect();
        const tabFontSize = Number.parseFloat(getComputedStyle(tab).fontSize || '0');
        return {
          ready: tileRect.width <= 46 && tabRect.height <= 26 && tabFontSize <= 11.5,
          tileWidth: tileRect.width,
          tileHeight: tileRect.height,
          tabHeight: tabRect.height,
          tabFontSize,
        };
      })()`,
      'compact Office formula toolbar density',
    );
    assert.ok(compactToolbarLayout.tileWidth <= 46);
    assert.ok(compactToolbarLayout.tabHeight <= 26);
    assert.ok(compactToolbarLayout.tabFontSize <= 11.5);

    await clickSelectorWithPointer(client, '.canvas-input-behavior-trigger');
    const inputBehaviorLayout = await waitForEvaluation(
      client,
      `(() => {
        const popover = document.querySelector('.input-behavior-popover');
        const tabs = document.querySelector('.classic-bottom-tabs');
        if (!(popover instanceof HTMLElement) || !(tabs instanceof HTMLElement)) {
          return { ready: false };
        }
        const popoverRect = popover.getBoundingClientRect();
        const tabsRect = tabs.getBoundingClientRect();
        const noDockOverlap = popoverRect.bottom <= tabsRect.top - 4;
        return {
          ready: noDockOverlap,
          noDockOverlap,
          popoverBottom: popoverRect.bottom,
          tabsTop: tabsRect.top,
          popoverClientHeight: popover.clientHeight,
          popoverScrollHeight: popover.scrollHeight,
        };
      })()`,
      'input behavior popover avoids the Office bottom toolbar',
    );
    assert.equal(inputBehaviorLayout.noDockOverlap, true);
    await client.evaluate(`(() => {
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
      return true;
    })()`);
    await client.send('Emulation.clearDeviceMetricsOverride');
    if (process.argv.includes('--layout-only')) {
      console.log(JSON.stringify({
        officeControlBarState,
        restoredCollapsedLayout,
        persistedPanelSizes,
        restoredPanelSizes,
        commonToolbarCount: commonBefore.count,
        firstToolbarVisual: commonBefore.firstVisual,
        formattingControls,
        colorPopoverLayout,
        compactToolbarLayout,
        inputBehaviorLayout,
      }, null, 2));
      console.log('Office editor layout, persistence, toolbar rendering and color-popover regression passed');
      return;
    }
    const noSelectionActions = await client.evaluate(`(() => {
      const field = document.querySelector('math-field');
      if (!field) return null;
      field.selection = {
        ranges: [[field.lastOffset, field.lastOffset]],
        direction: 'none',
      };
      field.position = field.lastOffset;
      const before = field.value;
      document.querySelector('[data-formula-selection-bold]')?.click();
      document.querySelector('[data-formula-selection-italic]')?.click();
      document.querySelector('[data-formula-selection-color]')?.click();
      document.querySelector('[data-formula-selection-background]')?.click();
      return {
        before,
        after: field.value,
        popoverOpen: Boolean(document.querySelector('[data-formula-color-popover]')),
      };
    })()`);
    assert.equal(noSelectionActions?.after, noSelectionActions?.before);
    assert.equal(noSelectionActions?.popoverOpen, false);

    await client.evaluate(`(() => {
      const field = document.querySelector('math-field');
      if (!field) return false;
      field.setValue('abcde', {
        mode: 'math',
        format: 'latex',
        insertionMode: 'replaceAll',
        selectionMode: 'after',
      });
      field.focus();
      field.selection = { ranges: [[0, 3]], direction: 'forward' };
      return true;
    })()`);
    await clickSelectorWithPointer(client, '[data-formula-selection-bold]');
    const selectedBold = await waitForEvaluation(
      client,
      `(() => {
        const field = document.querySelector('math-field');
        if (!field) return { ready: false };
        field.selection = { ranges: [[0, 3]], direction: 'forward' };
        const selectionBold = field.queryStyle({ variantStyle: 'bold' });
        const boldValue = field.value.replace(/\s+/g, '');
        return {
          ready: selectionBold === 'all',
          selectionBold,
          boldValue,
        };
      })()`,
      "selection-only bold formatting",
    );
    assert.equal(selectedBold.selectionBold, "all");
    assert.equal(selectedBold.boldValue, String.raw`\mathbf{abc}de`);
    const laterBold = await client.evaluate(`(() => {
      const field = document.querySelector('math-field');
      if (!field) return 'missing';
      field.selection = {
        ranges: [[field.lastOffset, field.lastOffset]],
        direction: 'none',
      };
      field.position = field.lastOffset;
      field.insert('f', {
        mode: 'math',
        format: 'latex',
        insertionMode: 'insertAfter',
        selectionMode: 'after',
        focus: true,
      });
      const insertedEnd = field.lastOffset;
      field.selection = {
        ranges: [[insertedEnd - 1, insertedEnd]],
        direction: 'forward',
      };
      return field.queryStyle({ variantStyle: 'bold' }) === 'all' ||
        field.queryStyle({ variantStyle: 'bolditalic' }) === 'all'
          ? 'all'
          : 'none';
    })()`);
    assert.notEqual(
      laterBold,
      "all",
      "Selection bold must not make later input persistently bold",
    );

    await client.evaluate(`(() => {
      const field = document.querySelector('math-field');
      if (!field) return false;
      field.setValue('xyz', {
        mode: 'math',
        format: 'latex',
        insertionMode: 'replaceAll',
        selectionMode: 'after',
      });
      field.focus();
      field.selection = { ranges: [[0, field.lastOffset]], direction: 'forward' };
      return true;
    })()`);
    await clickSelectorWithPointer(client, '[data-formula-selection-italic]');
    const selectedUpright = await client.evaluate(`(() => {
      const field = document.querySelector('math-field');
      return field?.value?.replace(/\s+/g, '') ?? '';
    })()`);
    assert.equal(selectedUpright, String.raw`\mathrm{xyz}`);
    await client.evaluate(`(() => {
      const field = document.querySelector('math-field');
      if (!field) return false;
      field.selection = { ranges: [[0, field.lastOffset]], direction: 'forward' };
      return true;
    })()`);
    await clickSelectorWithPointer(client, '[data-formula-selection-italic]');
    const selectedItalic = await client.evaluate(`(() => {
      const field = document.querySelector('math-field');
      return field?.value?.replace(/\s+/g, '') ?? '';
    })()`);
    assert.equal(selectedItalic, "xyz");

    await client.evaluate(`(() => {
      const field = document.querySelector('math-field');
      if (!field) return false;
      field.selection = { ranges: [[0, 2]], direction: 'forward' };
      document.querySelector('[data-formula-selection-color]')?.click();
      return true;
    })()`);
    await waitForEvaluation(
      client,
      `(() => ({
        ready: Boolean(
          document.querySelector(
            '[data-formula-color-popover="color"] [data-formula-color="#dc2626"]',
          ),
        ),
      }))()`,
      "formula text color popover",
    );
    await client.evaluate(`(() => {
      document.querySelector(
        '[data-formula-color-popover="color"] [data-formula-color="#dc2626"]',
      )?.click();
      return true;
    })()`);
    const selectedColor = await client.evaluate(`(() => {
      const field = document.querySelector('math-field');
      if (!field) return { color: 'none' };
      field.selection = { ranges: [[0, 2]], direction: 'forward' };
      return {
        color: field.queryStyle({ color: '#dc2626' }),
        popoverOpen: Boolean(document.querySelector('[data-formula-color-popover]')),
      };
    })()`);
    assert.equal(selectedColor.color, "all");
    assert.equal(selectedColor.popoverOpen, false);

    await client.evaluate(`(() => {
      const field = document.querySelector('math-field');
      if (!field) return false;
      field.selection = { ranges: [[2, 4]], direction: 'forward' };
      document.querySelector('[data-formula-selection-background]')?.click();
      return true;
    })()`);
    await waitForEvaluation(
      client,
      `(() => ({
        ready: Boolean(
          document.querySelector(
            '[data-formula-color-popover="backgroundColor"] [data-formula-color="#fef3c7"]',
          ),
        ),
      }))()`,
      "formula background color popover",
    );
    await client.evaluate(`(() => {
      document.querySelector(
        '[data-formula-color-popover="backgroundColor"] [data-formula-color="#fef3c7"]',
      )?.click();
      return true;
    })()`);
    const selectedBackground = await client.evaluate(`(() => {
      const field = document.querySelector('math-field');
      if (!field) return { background: 'none' };
      field.selection = { ranges: [[2, 4]], direction: 'forward' };
      return {
        background: field.queryStyle({ backgroundColor: '#fef3c7' }),
        popoverOpen: Boolean(document.querySelector('[data-formula-color-popover]')),
      };
    })()`);
    assert.equal(selectedBackground.background, "all");
    assert.equal(selectedBackground.popoverOpen, false);
    const laterColorState = await client.evaluate(`(() => {
      const field = document.querySelector('math-field');
      if (!field) return null;
      field.selection = {
        ranges: [[field.lastOffset, field.lastOffset]],
        direction: 'none',
      };
      field.position = field.lastOffset;
      field.insert('i', {
        mode: 'math',
        format: 'latex',
        insertionMode: 'insertAfter',
        selectionMode: 'after',
        focus: true,
      });
      const end = field.lastOffset;
      field.selection = { ranges: [[end - 1, end]], direction: 'forward' };
      return {
        color: field.queryStyle({ color: '#dc2626' }),
        background: field.queryStyle({ backgroundColor: '#fef3c7' }),
      };
    })()`);
    assert.notEqual(
      laterColorState?.color,
      "all",
      "Selection text color must not affect later input",
    );
    assert.notEqual(
      laterColorState?.background,
      "all",
      "Selection background color must not affect later input",
    );

    await client.evaluate(`(() => {
      const field = document.querySelector('math-field');
      if (!field) return false;
      field.setValue('abcde', {
        mode: 'math',
        format: 'latex',
        insertionMode: 'replaceAll',
        selectionMode: 'after',
      });
      field.focus();
      return true;
    })()`);
    const dragPoints = await waitForEvaluation(
      client,
      `(() => {
        const field = document.querySelector('math-field');
        const first = field?.getElementInfo(1)?.bounds;
        const fourth = field?.getElementInfo(4)?.bounds;
        const last = field?.getElementInfo(field?.lastOffset ?? 0)?.bounds;
        return {
          ready: Boolean(first && fourth && last),
          start: first
            ? { x: first.left + first.width / 2, y: first.top + first.height / 2 }
            : null,
          end: fourth
            ? { x: fourth.left + fourth.width / 2, y: fourth.top + fourth.height / 2 }
            : null,
          after: last
            ? { x: last.right + 80, y: last.top + last.height / 2 }
            : null,
        };
      })()`,
      "formula drag-selection geometry",
    );
    await client.send("Input.dispatchMouseEvent", {
      type: "mousePressed",
      x: dragPoints.start.x,
      y: dragPoints.start.y,
      button: "left",
      buttons: 1,
      clickCount: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseMoved",
      x: dragPoints.end.x,
      y: dragPoints.end.y,
      button: "left",
      buttons: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseReleased",
      x: dragPoints.end.x,
      y: dragPoints.end.y,
      button: "left",
      buttons: 0,
      clickCount: 1,
    });
    await sleep(100);
    const selectionAfterRelease = await client.evaluate(`(() => {
      const field = document.querySelector('math-field');
      return field
        ? JSON.parse(JSON.stringify(field.selection))
        : null;
    })()`);
    assert.ok(
      selectionAfterRelease?.ranges?.some(([start, end]) => start !== end),
      "Mouse drag should create a non-collapsed selection",
    );
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseMoved",
      x: dragPoints.after.x,
      y: dragPoints.after.y,
      button: "none",
      buttons: 0,
    });
    await sleep(120);
    const selectionAfterFreeMove = await client.evaluate(`(() => {
      const field = document.querySelector('math-field');
      return field
        ? JSON.parse(JSON.stringify(field.selection))
        : null;
    })()`);
    assert.deepEqual(
      selectionAfterFreeMove,
      selectionAfterRelease,
      "Selection must remain fixed after pointerup",
    );

    const shortcutMetadata = await client.evaluate(`(() => {
      const primary = document.querySelector('.office-inline-primary');
      window.__visualtexLateSaveShortcutCount = 0;
      window.addEventListener('keydown', (event) => {
        if (event.ctrlKey && (event.code === 'KeyS' || event.key.toLowerCase() === 's')) {
          window.__visualtexLateSaveShortcutCount += 1;
        }
      }, true);
      return {
        ariaKeyShortcuts: primary?.getAttribute('aria-keyshortcuts') ?? '',
        title: primary?.getAttribute('title') ?? '',
        legacyHelperVisible:
          (document.querySelector('.editor-pane-header.is-office-editor-header')?.textContent ?? '')
            .includes('点击完成、按 Ctrl+S') ||
          (document.querySelector('.editor-pane-header.is-office-editor-header')?.textContent ?? '')
            .includes('Finish, press Ctrl+S'),
      };
    })()`);
    assert.equal(shortcutMetadata.ariaKeyShortcuts, "Control+S");
    assert.ok(shortcutMetadata.title.includes("Ctrl+S"));
    assert.equal(shortcutMetadata.legacyHelperVisible, false);

    const committingBeforeShortcut = updates.filter(
      (update) => update.status === "committing",
    ).length;
    const shiftSave = await dispatchOfficeShortcut(client, { shiftKey: true });
    // Ctrl+Shift+S is now a shipped formula shortcut (sum). The Office apply
    // filter must still ignore it, while MathEditor is allowed to consume it.
    assert.equal(shiftSave.defaultPrevented, true);
    assert.equal(shiftSave.dispatchResult, false);
    assert.equal(shiftSave.lateCaptureCount, 1);
    await sleep(120);
    assert.equal(
      updates.filter((update) => update.status === "committing").length,
      committingBeforeShortcut,
      "Ctrl+Shift+S must not start an Office commit",
    );
    // Ctrl+Shift+S legitimately inserts the shipped sum template. Restore a
    // complete formula before testing the independent Office Ctrl+S apply path;
    // otherwise its structural placeholders must (correctly) block commit.
    await client.evaluate(`(() => {
      const field = document.querySelector('math-field');
      if (!field) return false;
      field.setValue('abcde', {
        mode: 'math',
        format: 'latex',
        insertionMode: 'replaceAll',
        selectionMode: 'after',
      });
      field.focus();
      return true;
    })()`);
    await waitForEvaluation(
      client,
      `(() => ({
        ready: document.querySelector('math-field')?.value === 'abcde',
      }))()`,
      "restore valid formula after Ctrl+Shift+S formula shortcut",
    );
    await sleep(160);

    const altSave = await dispatchOfficeShortcut(client, { altKey: true });
    assert.equal(altSave.defaultPrevented, false);
    assert.equal(altSave.lateCaptureCount, 2);

    const metaSave = await dispatchOfficeShortcut(client, {
      ctrlKey: false,
      metaKey: true,
    });
    assert.equal(metaSave.defaultPrevented, false);
    assert.equal(metaSave.lateCaptureCount, 2);

    const composingSave = await dispatchOfficeShortcut(client, { isComposing: true });
    assert.equal(composingSave.defaultPrevented, false);
    assert.equal(composingSave.lateCaptureCount, 3);

    const repeatedSave = await dispatchOfficeShortcut(client, { repeat: true });
    assert.equal(repeatedSave.defaultPrevented, true);
    assert.equal(repeatedSave.dispatchResult, false);
    assert.equal(repeatedSave.lateCaptureCount, 3);
    await sleep(150);
    assert.equal(
      updates.filter((update) => update.status === "committing").length,
      committingBeforeShortcut,
      "Repeated Ctrl+S must be swallowed without starting a commit",
    );

    completeCommits = true;
    const exactSave = await dispatchOfficeShortcut(client);
    assert.equal(exactSave.defaultPrevented, true);
    assert.equal(exactSave.dispatchResult, false);
    assert.equal(exactSave.lateCaptureCount, 3);
    for (
      let attempt = 0;
      attempt < 120 &&
      !(session.status === "completed" && closeRequests === 1);
      attempt += 1
    ) {
      await sleep(60);
    }
    assert.equal(session.status, "completed", "Ctrl+S should commit the Office Session");
    assert.equal(closeRequests, 1, "Ctrl+S should close the Office editor after applying");
    assert.equal(
      updates.filter((update) => update.status === "committing").length,
      committingBeforeShortcut + 1,
      "One Ctrl+S press must enqueue exactly one additional commit",
    );

    // Reproduce the hidden-converter race directly. The previous PowerPoint
    // page has already placed 20/28/31.5 pt values in the component lifecycle;
    // a Word converter must still render from its immutable 11 pt Session,
    // never from the previous React state.
    try {
      await client.send("Page.navigate", { url: "about:blank" });
    } catch (error) {
      if (!String(error).includes("navigated or closed")) throw error;
    }
    client.close();
    client = undefined;
    await sleep(250);
    session = createSession("word", 11);
    session.autoCommitOnClose = false;
    completeCommits = false;
    updates.length = 0;
    const converterTarget = await (
      await fetch(
        `http://127.0.0.1:${debugPort}/json/new?${encodeURIComponent(converterUrl)}`,
        { method: "PUT" },
      )
    ).json();
    client = new CdpClient(converterTarget.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    for (
      let attempt = 0;
      attempt < 200 &&
      !(
        session.status === "committing" &&
        session.exportResult?.svg &&
        session.exportResult?.pngBase64
      );
      attempt += 1
    ) {
      await sleep(60);
    }
    assert.equal(session.status, "committing", "Word converter should commit");
    assert.equal(session.fontSizePt, 11, "Word converter must preserve the paragraph size");
    assert.ok(session.exportResult?.svg, "Word converter should produce SVG");
    assert.ok(session.exportResult?.pngBase64, "Word converter should produce PNG");
    assert.ok(
      Math.abs(session.exportResult.width - 35.504533333333335) < 0.2,
      `Word 11 pt converter width is wrong: ${session.exportResult.width}`,
    );
    assert.ok(
      Math.abs(session.exportResult.height - 13.557333333333332) < 0.2,
      `Word 11 pt converter height is wrong: ${session.exportResult.height}`,
    );
    assert.ok(
      session.exportResult.width < 45,
      "Word converter reused the PowerPoint 20 pt render state",
    );
    const wordConverterUpdate = updates.find(
      (update) => update.status === "committing" && update.exportResult,
    );
    assert.equal(wordConverterUpdate?.fontSizePt, 11);

    console.log(
      JSON.stringify(
        {
          powerpointDefault,
          powerpointChineseSize,
          powerpointSavedSize: 31.5,
          powerpointReloaded,
          wordInherited,
          wordChineseSize,
          wordSavedSize: 13,
          wordConverter: {
            fontSizePt: session.fontSizePt,
            width: session.exportResult.width,
            height: session.exportResult.height,
          },
        },
        null,
        2,
      ),
    );
    console.log("Office formula font-size regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    await new Promise((resolve) => server.close(resolve));
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => undefined);
  }
}

await main();
