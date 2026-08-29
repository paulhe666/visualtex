import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const portOffset = process.pid % 800;
const devPort = 6100 + portOffset;
const debugPort = 16100 + portOffset;
const baseUrl = `http://127.0.0.1:${devPort}`;
const chromeProfile = createBrowserProfilePath("visualtex-ocr-align-geometry");
const chromePath = resolveChromiumExecutable();

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 20000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return response;
    } catch {
      // Retry while Vite/Chromium starts.
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
  const vite = spawn(
    process.execPath,
    [
      "node_modules/vite/bin/vite.js",
      "--host",
      "127.0.0.1",
      "--port",
      String(devPort),
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
        baseUrl,
      ],
      { stdio: "ignore" },
    );

    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    let page;
    const targetStarted = Date.now();
    while (!page && Date.now() - targetStarted < 10000) {
      const targets = await (
        await fetch(`http://127.0.0.1:${debugPort}/json/list`)
      ).json();
      page = targets.find(
        (target) => target.type === "page" && target.url.startsWith(baseUrl),
      );
      if (!page) await sleep(100);
    }
    if (!page) throw new Error("No VisualTeX browser page target found");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(700);

    const evaluate = async (expression) => {
      const result = await client.send("Runtime.evaluate", {
        expression,
        awaitPromise: true,
        returnByValue: true,
      });
      if (result.exceptionDetails) {
        const description = result.exceptionDetails.exception?.description;
        throw new Error(
          description || result.exceptionDetails.text || "Runtime.evaluate failed",
        );
      }
      return result.result.value;
    };

    const waitForEvaluation = async (
      expression,
      description,
      timeoutMs = 10000,
    ) => {
      const started = Date.now();
      let lastValue;
      while (Date.now() - started < timeoutMs) {
        lastValue = await evaluate(expression);
        if (lastValue?.ready) return lastValue;
        await sleep(100);
      }
      throw new Error(
        `Timed out waiting for ${description}. Last value: ${JSON.stringify(lastValue)}`,
      );
    };

    await evaluate(`(() => {
      localStorage.setItem('visualtex.onboarding.v3.completed', 'true');
      localStorage.setItem('visualtex.office.macos.first-run.v1.completed', 'true');
      localStorage.setItem('visualtex.onboarding.windows.desktop.v1.1.0.completed', 'true');
    })()`);
    await client.send('Page.reload', { ignoreCache: true });
    await sleep(800);
    await waitForEvaluation(
      `(() => ({ ready: document.querySelectorAll('math-field').length >= 1 }))()`,
      "initial MathLive field",
    );

    const rawOcrAlign = String.raw`\\begin{align}
x&=1\\\\
y_{12345}&=\\frac{2}{3}\\\\
z+q+r&s=4
\\end{align}`;
    const ocrSetup = await evaluate(`(async () => {
      const ocr = await import('/src/ocr/ocrService.ts');
      const storeModule = await import('/src/stores/editorStore.ts');
      const clipboard = await import('/src/clipboard/LatexCopyService.ts');
      const markerModule = await import('/src/editor/alignmentMarkers.ts');
      const raw = ${JSON.stringify(rawOcrAlign)};
      const rows = ocr.normalizeOcrFormulaText(raw);
      const marker = markerModule.VISUALTEX_ALIGNMENT_MARKER_LATEX;
      if (rows.length !== 3 || !rows.every((row) => row.includes(marker))) {
        return { ready: false, rows, marker };
      }
      const store = storeModule.useEditorStore.getState();
      const lines = rows.map((latex, index) => ({
        id: 'ocr-align-' + (index + 1),
        latex,
      }));
      store.setLatexCodeFormat('aligned');
      store.replaceDocumentState({
        title: 'OCR align geometry acceptance',
        lines,
        activeLineId: lines[0].id,
        formulaAlignment: 'center',
      });
      window.__visualtexOcrAlignAcceptance = { rows, marker, clipboard };
      window.dispatchEvent(new Event('resize'));
      return { ready: true, rows, marker };
    })()`);
    assert.equal(ocrSetup.ready, true, `OCR rows were not normalized: ${JSON.stringify(ocrSetup)}`);
    assert.equal(ocrSetup.rows.length, 3);

    const geometry = await waitForEvaluation(
      `(() => {
        const fields = [...document.querySelectorAll('.formula-line math-field')];
        if (fields.length !== 3) return { ready: false, fieldCount: fields.length };
        const rows = fields.map((field, index) => {
          const marker = field.shadowRoot?.querySelector('.visualtex-align-marker');
          const host = field.closest('.mathfield-host');
          if (!marker || !host) return { index, missing: true };
          const markerRect = marker.getBoundingClientRect();
          const fieldRect = field.getBoundingClientRect();
          const hostRect = host.getBoundingClientRect();
          return {
            index,
            missing: false,
            markerX: markerRect.left,
            markerWidth: markerRect.width,
            fieldLeft: fieldRect.left,
            fieldRight: fieldRect.right,
            hostLeft: hostRect.left,
            hostRight: hostRect.right,
            marginLeft: parseFloat(getComputedStyle(field).marginLeft) || 0,
            value: field.value,
          };
        });
        if (rows.some((row) => row.missing)) return { ready: false, rows };
        const xs = rows.map((row) => row.markerX);
        const spread = Math.max(...xs) - Math.min(...xs);
        return { ready: spread <= 1, spread, rows };
      })()`,
      "three OCR alignment markers sharing one screen anchor",
      15000,
    );

    assert.equal(geometry.rows.length, 3);
    assert.ok(
      geometry.spread <= 1,
      `OCR align markers are not co-linear: spread=${geometry.spread}px rows=${JSON.stringify(geometry.rows)}`,
    );
    assert.ok(
      new Set(geometry.rows.map((row) => Math.round(row.marginLeft * 10) / 10)).size > 1,
      `Geometry acceptance did not exercise unequal left extents: ${JSON.stringify(geometry.rows)}`,
    );

    const serialized = await evaluate(`(async () => {
      const storeModule = await import('/src/stores/editorStore.ts');
      const clipboard = await import('/src/clipboard/LatexCopyService.ts');
      const markerModule = await import('/src/editor/alignmentMarkers.ts');
      const state = storeModule.useEditorStore.getState();
      const lines = state.lines.map((line) => line.latex);
      const aligned = clipboard.formatLatexLines(lines, 'aligned');
      return {
        format: state.latexCodeFormat,
        lines,
        aligned,
        hasPrivateMarker: aligned.includes(markerModule.VISUALTEX_ALIGNMENT_MARKER_LATEX),
        ampersandCount: (aligned.match(/&/g) || []).length,
        hasAlignedEnvironment: aligned.includes('\\\\begin{aligned}') && aligned.includes('\\\\end{aligned}'),
      };
    })()`);

    assert.equal(serialized.format, "aligned");
    assert.equal(serialized.hasPrivateMarker, false);
    assert.equal(serialized.ampersandCount, 3);
    assert.equal(serialized.hasAlignedEnvironment, true);

    console.log(
      `OCR align editor geometry acceptance passed: anchor spread=${geometry.spread.toFixed(3)}px, ` +
        `margins=${geometry.rows.map((row) => row.marginLeft.toFixed(3)).join('/')}, ` +
        `copy restored ${serialized.ampersandCount} '&' markers inside aligned.`,
    );
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    vite.kill("SIGTERM");
    await rm(chromeProfile, {
      recursive: true,
      force: true,
      maxRetries: 4,
      retryDelay: 150,
    }).catch(() => undefined);
  }
}

await main();
