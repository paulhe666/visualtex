import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";
import { normalizeOcrFormulaText } from "../src/ocr/ocrService.ts";
import { formatLatexLines } from "../src/clipboard/LatexCopyService.ts";
import { VISUALTEX_ALIGNMENT_MARKER_LATEX } from "../src/editor/alignmentMarkers.ts";

const portOffset = process.pid % 800;
const previewPort = 6200 + portOffset;
const debugPort = 16200 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = createBrowserProfilePath("visualtex-ocr-align-geometry");
const chromePath = resolveChromiumExecutable();

const rawOcrAlign = String.raw`\begin{align}
x&=a+b+c\\
y_{12345}&=\frac{p}{q}\\
z+q+r&=s
\end{align}`;
const ocrRows = normalizeOcrFormulaText(rawOcrAlign);
assert.equal(ocrRows.length, 3, `OCR align did not split into three FormulaLines: ${JSON.stringify(ocrRows)}`);
assert.ok(
  ocrRows.every((row) => row.includes(VISUALTEX_ALIGNMENT_MARKER_LATEX)),
  `OCR align did not encode one top-level alignment marker per row: ${JSON.stringify(ocrRows)}`,
);

const serializedAligned = formatLatexLines(ocrRows, "aligned");
assert.ok(serializedAligned.includes("\\begin{aligned}"));
assert.ok(serializedAligned.includes("\\end{aligned}"));
assert.equal(
  (serializedAligned.match(/&/g) ?? []).length,
  3,
  `Aligned copy did not restore three real '&' tokens: ${serializedAligned}`,
);
assert.ok(
  !serializedAligned.includes(VISUALTEX_ALIGNMENT_MARKER_LATEX),
  `Aligned copy leaked VisualTeX private marker syntax: ${serializedAligned}`,
);

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url: string, timeoutMs = 20000) {
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
  url: string;
  nextId = 1;
  pending = new Map<number, { resolve: (value: any) => void; reject: (error: Error) => void }>();
  socket?: WebSocket;

  constructor(url: string) {
    this.url = url;
  }

  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise<void>((resolve, reject) => {
      this.socket!.addEventListener("open", () => resolve(), { once: true });
      this.socket!.addEventListener("error", () => reject(new Error("CDP WebSocket failed")), {
        once: true,
      });
    });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(String(event.data));
      if (!message.id) return;
      const pending = this.pending.get(message.id);
      if (!pending) return;
      this.pending.delete(message.id);
      if (message.error) pending.reject(new Error(message.error.message));
      else pending.resolve(message.result);
    });
  }

  send(method: string, params: Record<string, unknown> = {}) {
    const id = this.nextId++;
    return new Promise<any>((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.socket!.send(JSON.stringify({ id, method, params }));
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
  let chrome: ReturnType<typeof spawn> | undefined;
  let client: CdpClient | undefined;

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
    let page: any;
    const targetStarted = Date.now();
    while (!page && Date.now() - targetStarted < 10000) {
      const targets = await (
        await fetch(`http://127.0.0.1:${debugPort}/json/list`)
      ).json() as any[];
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
    await sleep(650);

    const evaluate = async (expression: string) => {
      const result = await client!.send("Runtime.evaluate", {
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

    await evaluate(`(() => {
      localStorage.setItem('visualtex.onboarding.v3.completed', 'true');
      localStorage.setItem('visualtex.office.macos.first-run.v1.completed', 'true');
      localStorage.setItem('visualtex.onboarding.windows.desktop.v1.1.0.completed', 'true');
      const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}');
      const rows = ${JSON.stringify(ocrRows)};
      const lines = rows.map((latex, index) => ({ id: 'ocr-align-' + (index + 1), latex }));
      persisted.state = {
        ...(persisted.state || {}),
        lines,
        activeLineId: lines[0].id,
        latexCodeFormat: 'aligned',
        formulaAlignment: 'center',
      };
      localStorage.setItem('visualtex-editor', JSON.stringify(persisted));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(800);

    const started = Date.now();
    let state: any = null;
    while (Date.now() - started < 15000) {
      state = await evaluate(`(() => {
        const fields = [...document.querySelectorAll('.formula-line math-field')];
        if (fields.length !== 3) return { ready: false, count: fields.length };
        const entries = fields.map((field) => {
          const host = field.closest('.mathfield-host');
          const marker = field.shadowRoot?.querySelector('.visualtex-align-marker');
          const fieldBounds = field.getBoundingClientRect();
          const markerBounds = marker?.getBoundingClientRect();
          return {
            value: field.value,
            fieldLeft: fieldBounds.left,
            markerLeft: markerBounds?.left ?? null,
            marginLeft: Number.parseFloat(field.style.marginLeft || '0') || 0,
            hostAligned: host?.classList.contains('has-explicit-align-marker') ?? false,
          };
        });
        const markerLefts = entries
          .map((entry) => entry.markerLeft)
          .filter((value) => Number.isFinite(value));
        const markerSpread = markerLefts.length
          ? Math.max(...markerLefts) - Math.min(...markerLefts)
          : Number.POSITIVE_INFINITY;
        return {
          ready:
            fields.length === 3 &&
            markerLefts.length === 3 &&
            entries.every((entry) => entry.hostAligned) &&
            markerSpread <= 1,
          count: fields.length,
          entries,
          markerSpread,
        };
      })()`);
      if (state?.ready) break;
      await sleep(100);
    }

    assert.ok(state?.ready, `OCR alignment geometry did not converge: ${JSON.stringify(state)}`);
    assert.ok(state.markerSpread <= 1, `OCR alignment anchors diverged by ${state.markerSpread}px`);
    assert.ok(
      new Set(state.entries.map((entry: any) => Math.round(entry.marginLeft * 1000))).size > 1,
      `Acceptance did not exercise unequal left extents: ${JSON.stringify(state.entries)}`,
    );

    console.log(
      `OCR align editor geometry acceptance passed: markerSpread=${state.markerSpread}px, ` +
        `margins=${state.entries.map((entry: any) => entry.marginLeft).join('/')}, ` +
        `copy restored 3 '&' tokens in aligned.`,
    );
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await rm(chromeProfile, {
      recursive: true,
      force: true,
      maxRetries: 4,
      retryDelay: 150,
    }).catch(() => undefined);
  }
}

await main();
