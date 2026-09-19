import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { resolveChromiumExecutable } from "./browser_test_runtime.mjs";

// Browser integration only. Both application windows use a NEW temporary
// Chromium profile. The Office Session response is an in-page fixture: this
// test never connects to Word, AppKit, or a user's installed application.
const offset = process.pid % 700;
const previewPort = 19000 + offset;
const debugPort = 21000 + offset;
const origin = `http://127.0.0.1:${previewPort}`;
const sessionId = "17fa9abd-215c-4576-8a7a-394a49747c51";
const profile = {
  inlineWrapper: "paren", inlineTextPolicy: "outside-math",
  displayWrapper: "equation", numbered: false, multilineEnvironment: "align",
};
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
async function waitFor(description, predicate, timeoutMs = 12000) {
  const started = Date.now();
  let lastError;
  while (Date.now() - started < timeoutMs) {
    try { const value = await predicate(); if (value) return value; } catch (error) { lastError = error; }
    await sleep(80);
  }
  throw new Error(`Timed out: ${description}${lastError ? ` (${lastError.message})` : ""}`);
}
class Client {
  nextId = 0;
  pending = new Map();
  exceptions = [];
  async connect(url) {
    this.socket = new WebSocket(url);
    await new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error("CDP connection timeout")), 5000);
      this.socket.addEventListener("open", () => { clearTimeout(timer); resolve(); }, { once: true });
      this.socket.addEventListener("error", (error) => { clearTimeout(timer); reject(error); }, { once: true });
    });
    this.socket.addEventListener("message", ({ data }) => {
      const message = JSON.parse(data);
      if (message.method === "Runtime.exceptionThrown") this.exceptions.push(message.params);
      const task = this.pending.get(message.id);
      if (!task) return;
      this.pending.delete(message.id);
      clearTimeout(task.timer);
      if (message.error) task.reject(new Error(message.error.message)); else task.resolve(message.result);
    });
    await this.send("Runtime.enable");
    await this.send("Page.enable");
  }
  send(method, params = {}) {
    const id = ++this.nextId;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => { this.pending.delete(id); reject(new Error(`CDP timed out: ${method}`)); }, 7000);
      this.pending.set(id, { resolve, reject, timer });
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }
  async evaluate(expression) {
    const result = await this.send("Runtime.evaluate", { expression, returnByValue: true, awaitPromise: true });
    if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description ?? result.exceptionDetails.text);
    return result.result.value;
  }
  close() {
    for (const task of this.pending.values()) { clearTimeout(task.timer); task.reject(new Error("CDP closed")); }
    this.pending.clear();
    this.socket?.close();
  }
}
async function main() {
  const chromeProfile = await mkdtemp(join(tmpdir(), "visualtex-office-profile-browser-"));
  const evidenceRoot = join(process.cwd(), "test-results", "office-profile-browser");
  await mkdir(evidenceRoot, { recursive: true });
  const artifacts = await mkdtemp(join(evidenceRoot, "run-"));
  const preview = spawn(process.execPath, [
    "node_modules/vite/bin/vite.js", "preview", "--host", "127.0.0.1", "--port", String(previewPort), "--strictPort",
  ], { stdio: "ignore", cwd: process.cwd() });
  const office = new Client();
  const settings = new Client();
  let chrome;
  const targets = async () => (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
  try {
    await waitFor("preview server", async () => (await fetch(origin)).ok);
    chrome = spawn(resolveChromiumExecutable(), [
      "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
      `--user-data-dir=${chromeProfile}`, `--remote-debugging-port=${debugPort}`, "--window-size=1440,1000", "about:blank",
    ], { stdio: "ignore" });
    const target = await waitFor("isolated Chrome target", async () => (await targets()).find((item) => item.type === "page"));
    await office.connect(target.webSocketDebuggerUrl);
    await office.send("Emulation.setFocusEmulationEnabled", { enabled: true });
    await office.send("Page.addScriptToEvaluateOnNewDocument", { source: `(() => {
      localStorage.setItem('visualtex-editor', JSON.stringify({ version: 0, state: {
        title: 'main-document-sentinel', lines: [{ id: 'main-sentinel', latex: 'a+b' }],
        activeLineId: 'main-sentinel', language: 'en', editorLayout: 'standard', zoom: 0.6,
        latexFormatProfile: ${JSON.stringify(profile)}
      }}));
      localStorage.setItem('visualtex-office-editor-source-open', 'true');
      localStorage.setItem('visualtex-office-editor-toolbar-open', 'true');
      localStorage.setItem('visualtex-active-theme', 'light');
      window.__copiedLatex = null;
      Object.defineProperty(navigator, 'clipboard', { configurable: true, value: {
        writeText: async (text) => { window.__copiedLatex = text; }
      }});
      const fixture = {
        id: ${JSON.stringify(sessionId)}, mode: 'edit', host: 'word',
        formulaId: 'fixture-formula', sourceDocumentId: 'fixture-document', sourceObjectId: 'fixture-object',
        title: 'Office profile integration', lines: [{ id: 'office-line', latex: 'x+1' }],
        activeLineId: 'office-line', codeFormat: 'raw', displayMode: 'block', numbered: false,
        fontSizePt: 12, exportWidth: 320, exportHeight: 80, exportResult: null, originalMetadata: null,
        dirty: false, status: 'editing', autoCommitOnClose: false, explicitCancel: false, error: null,
        createdAt: Date.now(), updatedAt: Date.now(), expiresAt: Date.now() + 600000
      };
      const originalFetch = window.fetch.bind(window);
      window.fetch = async (input, init = {}) => {
        const url = new URL(typeof input === 'string' ? input : input.url, location.href);
        if (url.pathname === '/api/v1/sessions/${sessionId}') {
          if (String(init.method || 'GET').toUpperCase() === 'PATCH') Object.assign(fixture, JSON.parse(init.body || '{}'));
          return new Response(JSON.stringify(fixture), { status: 200, headers: { 'Content-Type': 'application/json' } });
        }
        return originalFetch(input, init);
      };
    })();` });
    await office.send("Page.navigate", { url: `${origin}/office-native-dialog.html?sessionId=${sessionId}&officeHost=word` });
    await waitFor("Office formula hydration", () => office.evaluate(`document.querySelector('math-field')?.value === 'x+1'`));
    const focus = await waitFor("initial formula focus", () => office.evaluate(`document.activeElement?.tagName === 'MATH-FIELD'`));
    assert.equal(focus, true);
    const before = await office.evaluate(`({
      latex: document.querySelector('math-field').value,
      focused: document.activeElement.tagName,
      primaryPresent: Boolean(document.querySelector('[data-office-primary-action]')),
      sourceCopyPresent: Boolean(document.querySelector('.source-copy-button')),
      inputBehaviorLabel: document.querySelector('.canvas-input-behavior-trigger > span')?.textContent,
      inputBehaviorLabelVisible: (() => {
        const label = document.querySelector('.canvas-input-behavior-trigger > span');
        return Boolean(label && label.getBoundingClientRect().width > 0 && getComputedStyle(label).display !== 'none');
      })()
    })`);
    assert.equal(before.primaryPresent, true);
    assert.equal(before.sourceCopyPresent, true, "Actual source-copy UI must be present");
    assert.equal(before.inputBehaviorLabel, "Input behavior", "The input-behavior trigger must retain its visible label");
    assert.equal(before.inputBehaviorLabelVisible, true);
    const copy = async () => {
      await office.evaluate(`(() => { window.__copiedLatex = null; document.querySelector('.source-copy-button').click(); })()`);
      return waitFor("source-copy clipboard callback", () => office.evaluate("window.__copiedLatex"));
    };
    const initialCopy = await copy();
    assert.match(initialCopy, /\\begin\{equation\*\}/);
    assert.match(initialCopy, /x\+1/);

    // A second same-origin application window emits a REAL storage event to
    // the Office page; the test does not call the production sync function.
    const created = await office.send("Target.createTarget", { url: origin });
    const settingsTarget = await waitFor("settings page target", async () =>
      (await targets()).find((item) => item.id === created.targetId));
    await settings.connect(settingsTarget.webSocketDebuggerUrl);
    await waitFor("settings page origin", () => settings.evaluate(`location.origin === ${JSON.stringify(origin)} && document.readyState === 'complete'`));
    await settings.evaluate(`(() => {
      const stored = JSON.parse(localStorage.getItem('visualtex-editor'));
      stored.state.latexFormatProfile = { ...stored.state.latexFormatProfile, displayWrapper: 'double-dollar', numbered: false };
      localStorage.setItem('visualtex-editor', JSON.stringify(stored));
    })()`);
    await sleep(250);
    const synchronizedCopy = await copy();
    assert.match(synchronizedCopy, /^\$\$\s*x\+1\s*\$\$$/);
    assert.equal(await office.evaluate("document.querySelector('math-field').value"), "x+1");
    const retained = await settings.evaluate(`JSON.parse(localStorage.getItem('visualtex-editor')).state`);
    assert.equal(retained.title, "main-document-sentinel");
    assert.deepEqual(retained.lines.map((line) => line.latex), ["a+b"]);
    assert.equal(retained.latexFormatProfile.displayWrapper, "double-dollar");
    assert.deepEqual(office.exceptions, [], "The Office editor must have no uncaught JavaScript exception");
    const capture = async (name) => {
      const screenshot = await office.send("Page.captureScreenshot", { format: "png", captureBeyondViewport: false });
      const screenshotPath = join(artifacts, name);
      await writeFile(screenshotPath, Buffer.from(screenshot.data, "base64"));
      return screenshotPath;
    };
    const screenshotPath = await capture("office-profile-focus.png");
    const changeLayout = async (layout) => {
      await settings.evaluate(`(() => {
        const stored = JSON.parse(localStorage.getItem('visualtex-editor'));
        stored.state.editorLayout = ${JSON.stringify(layout)};
        localStorage.setItem('visualtex-editor', JSON.stringify(stored));
      })()`);
      await waitFor(`Office ${layout} layout`, () => office.evaluate(
        `Boolean(document.querySelector('.workspace.is-classic-layout')) === ${layout === "classic"}`,
      ));
      await sleep(200);
    };
    const visibleActions = () => office.evaluate(`(() => {
      const visible = (element) => element.getBoundingClientRect().width > 0 && element.getBoundingClientRect().height > 0;
      return {
        primary: [...document.querySelectorAll('[data-office-primary-action]')].filter(visible).length,
        cancel: [...document.querySelectorAll('[data-office-cancel-action]')].filter(visible).length,
        field: [...document.querySelectorAll('math-field')].filter(visible).length,
      };
    })()`);
    await changeLayout("classic");
    const classicActions = await visibleActions();
    assert(classicActions.primary > 0, "Classic layout must expose the real Office primary action");
    assert(classicActions.cancel > 0);
    assert(classicActions.field > 0);
    const classicScreenshotPath = await capture("office-profile-classic.png");
    await changeLayout("standard");
    await office.send("Emulation.setDeviceMetricsOverride", { width: 720, height: 800, deviceScaleFactor: 1, mobile: false });
    await sleep(250);
    const narrowActions = await visibleActions();
    assert(narrowActions.primary > 0);
    assert(narrowActions.cancel > 0);
    assert(narrowActions.field > 0);
    const narrowTrigger = await office.evaluate(`(() => {
      const button = document.querySelector('.canvas-input-behavior-trigger');
      const icon = button?.querySelector('svg');
      const label = button?.querySelector('span');
      return { width: button?.getBoundingClientRect().width,
        iconVisible: Boolean(icon && icon.getBoundingClientRect().width > 0 && getComputedStyle(icon).display !== 'none'),
        labelHidden: label ? getComputedStyle(label).display === 'none' : null,
        title: button?.title };
    })()`);
    assert.equal(narrowTrigger.iconVisible, true, "Compact headers must not hide their sole visible input-behavior icon");
    assert.equal(narrowTrigger.labelHidden, true);
    assert.equal(narrowTrigger.title, "Input behavior");
    const narrowScreenshotPath = await capture("office-profile-narrow.png");
    assert.deepEqual(office.exceptions, []);
    const report = { before, initialCopy, synchronizedCopy, sentinelRetained: true, screenshotPath,
      classicActions, classicScreenshotPath, narrowActions, narrowTrigger, narrowScreenshotPath,
      uncaughtExceptions: office.exceptions.length };
    await writeFile(join(artifacts, "report.json"), JSON.stringify(report, null, 2));
    console.log(JSON.stringify(report, null, 2));
    console.log("macOS Office browser profile/focus regression passed; this is not native Word/WKWebView acceptance");
  } catch (error) {
    try {
      const state = await office.evaluate(`({ url: location.href, active: document.activeElement?.tagName,
        text: document.body.innerText.slice(0, 1800), fields: [...document.querySelectorAll('math-field')].map(x => x.value) })`);
      console.error(JSON.stringify({ state, exceptions: office.exceptions }, null, 2));
      const screenshot = await office.send("Page.captureScreenshot", { format: "png", captureBeyondViewport: false });
      await writeFile(join(artifacts, "failure.png"), Buffer.from(screenshot.data, "base64"));
      console.error(`Failure evidence: ${artifacts}`);
    } catch { /* Retain the original error. */ }
    throw error;
  } finally {
    office.close(); settings.close(); chrome?.kill("SIGTERM"); preview.kill("SIGTERM");
    await sleep(250);
    await rm(chromeProfile, { recursive: true, force: true });
  }
}
main().catch((error) => { console.error(error); process.exitCode = 1; });
