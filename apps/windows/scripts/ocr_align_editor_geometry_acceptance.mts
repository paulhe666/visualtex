import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";
import { normalizeOcrFormulaText } from "../src/ocr/ocrService.ts";
import {
  formatLatexLines,
  parseLatexSource,
  parseLatexSourceDraft,
} from "../src/clipboard/LatexCopyService.ts";
import { VISUALTEX_ALIGNMENT_MARKER_LATEX } from "../src/editor/alignmentMarkers.ts";

const portOffset = process.pid % 800;
const previewPort = 6200 + portOffset;
const debugPort = 16200 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = createBrowserProfilePath("visualtex-ocr-align-geometry");
const chromePath = resolveChromiumExecutable();

const exactUserAligned = String.raw`\begin{aligned}
\langle p_1,p_0\rangle &\leftarrow \operatorname{umul}(a,b)=ab
&&\text{Double word product}\\
p_0 &\leftarrow \operatorname{umullo}(a,b)=(ab)\bmod\beta
&&\text{Low word}\\
p_1 &\leftarrow \operatorname{umulhi}(a,b)=\left\lfloor\frac{ab}{\beta}\right\rfloor
&&\text{High word.}
\end{aligned}`;
const exactUserRows = parseLatexSource(exactUserAligned, "aligned");
assert.equal(
  exactUserRows.length,
  3,
  `Exact user aligned source did not parse into three rows: ${JSON.stringify(exactUserRows)}`,
);
assert.deepEqual(
  exactUserRows.map(
    (row) => row.split(VISUALTEX_ALIGNMENT_MARKER_LATEX).length - 1,
  ),
  [3, 3, 3],
  `Exact user aligned source did not preserve '&' + '&&' as three markers per row: ${JSON.stringify(exactUserRows)}`,
);
const exactUserDraft = parseLatexSourceDraft(
  `\\[\n${exactUserAligned}\n\\]`,
  "aligned",
);
assert.ok(
  exactUserDraft.valid,
  `Exact user aligned source failed strict aligned draft parsing: ${JSON.stringify(exactUserDraft)}`,
);
assert.deepEqual(
  exactUserDraft.values,
  exactUserRows,
  `Strict aligned source editor parsing differs from direct aligned import: ${JSON.stringify(exactUserDraft.values)}`,
);
const exactUserSerialized = formatLatexLines(exactUserRows, "aligned");
assert.equal(
  (exactUserSerialized.match(/&/g) ?? []).length,
  9,
  `Exact user aligned serialization did not restore all nine ampersands: ${exactUserSerialized}`,
);
assert.ok(
  exactUserSerialized.includes("&&\\text{Low word}"),
  `Exact user aligned serialization did not preserve the empty && column: ${exactUserSerialized}`,
);

const rawOcrAlign = String.raw`\begin{align}
x&=a+b+c&u&=v\\
y_{12345}&=\frac{p}{q}&&\text{middle note}\\
z+q+r&=s&long&=t
\end{align}`;
const ocrRows = normalizeOcrFormulaText(rawOcrAlign);
assert.equal(ocrRows.length, 3, `OCR align did not split into three FormulaLines: ${JSON.stringify(ocrRows)}`);
assert.ok(
  ocrRows.every(
    (row) => row.split(VISUALTEX_ALIGNMENT_MARKER_LATEX).length - 1 === 3,
  ),
  `OCR align did not preserve all three top-level alignment markers per row: ${JSON.stringify(ocrRows)}`,
);

const serializedAligned = formatLatexLines(ocrRows, "aligned");
assert.ok(serializedAligned.includes("\\begin{aligned}"));
assert.ok(serializedAligned.includes("\\end{aligned}"));
assert.ok(
  serializedAligned.includes("&&\\text{middle note}"),
  `Aligned copy did not preserve an explicit empty column created by '&&': ${serializedAligned}`,
);
assert.equal(
  (serializedAligned.match(/&/g) ?? []).length,
  9,
  `Aligned copy did not restore all nine real '&' tokens: ${serializedAligned}`,
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
          const markers = [...(field.shadowRoot?.querySelectorAll('.visualtex-align-marker') ?? [])];
          const fieldBounds = field.getBoundingClientRect();
          const visualAnchorLeft = (marker) => {
            for (
              let sibling = marker.nextElementSibling;
              sibling;
              sibling = sibling.nextElementSibling
            ) {
              const candidates = sibling.hasAttribute('data-atom-id')
                ? [sibling, ...sibling.querySelectorAll('[data-atom-id]')]
                : [...sibling.querySelectorAll('[data-atom-id]')];
              for (const candidate of candidates) {
                const bounds = candidate.getBoundingClientRect();
                if (bounds.width > 0.01 && bounds.height > 0.01) return bounds.left;
              }
            }
            return marker.getBoundingClientRect().left;
          };
          return {
            value: field.value,
            fieldLeft: fieldBounds.left,
            markerLefts: markers.map((marker) => marker.getBoundingClientRect().left),
            visualAnchorLefts: markers.map(visualAnchorLeft),
            markerMarginLefts: markers.map(
              (marker) => Number.parseFloat(getComputedStyle(marker).marginLeft || '0') || 0,
            ),
            markerMarginRights: markers.map(
              (marker) => Number.parseFloat(getComputedStyle(marker).marginRight || '0') || 0,
            ),
            marginLeft: Number.parseFloat(field.style.marginLeft || '0') || 0,
            hostAligned: host?.classList.contains('has-explicit-align-marker') ?? false,
          };
        });
        // TeX align columns are delimited by markers; a cell after a marker may
        // intentionally be empty (two adjacent ampersands), so marker boundaries are authoritative.
        const markerSpreads = [0, 1, 2].map((markerIndex) => {
          const positions = entries
            .map((entry) => entry.markerLefts[markerIndex])
            .filter((value) => Number.isFinite(value));
          return positions.length === fields.length
            ? Math.max(...positions) - Math.min(...positions)
            : Number.POSITIVE_INFINITY;
        });
        return {
          ready:
            fields.length === 3 &&
            entries.every(
              (entry) => entry.hostAligned && entry.markerLefts.length === 3,
            ) &&
            markerSpreads.every((spread) => spread <= 1),
          count: fields.length,
          entries,
          markerSpreads,
        };
      })()`);
      if (state?.ready) break;
      await sleep(100);
    }

    assert.ok(state?.ready, `OCR alignment geometry did not converge: ${JSON.stringify(state)}`);
    assert.ok(
      state.markerSpreads.every((spread: number) => spread <= 1),
      `OCR marker columns diverged: ${state.markerSpreads.join('/')}px`,
    );
    assert.ok(
      new Set(state.entries.map((entry: any) => Math.round(entry.marginLeft * 1000))).size > 1,
      `Acceptance did not exercise unequal first-column extents: ${JSON.stringify(state.entries)}`,
    );
    assert.ok(
      state.entries.some(
        (entry: any) =>
          (entry.markerMarginLefts[1] ?? 0) > 0 ||
          (entry.markerMarginRights[1] ?? 0) > 0,
      ),
      `Acceptance did not exercise multi-pair marker spacing: ${JSON.stringify(state.entries)}`,
    );

    console.log(
      `OCR multi-align editor geometry acceptance passed: markerSpreads=${state.markerSpreads.join('/')}px, ` +
        `margins=${state.entries.map((entry: any) => entry.marginLeft).join('/')}, ` +
        `copy restored 9 '&' tokens including '&&'.`,
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
