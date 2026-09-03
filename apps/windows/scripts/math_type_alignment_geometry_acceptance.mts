import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const portOffset = process.pid % 700;
const devPort = 7600 + portOffset;
const debugPort = 17600 + portOffset;
const baseUrl = `http://127.0.0.1:${devPort}`;
const chromeProfile = createBrowserProfilePath("visualtex-mathtype-align-geometry");
const chromePath = resolveChromiumExecutable();

const exactUserAligned = String.raw`\begin{aligned}
\langle p_1,p_0\rangle &\leftarrow \operatorname{umul}(a,b)=ab
&&\text{Double word product}\\
p_0 &\leftarrow \operatorname{umullo}(a,b)=(ab)\bmod\beta
&&\text{Low word}\\
p_1 &\leftarrow \operatorname{umulhi}(a,b)=\left\lfloor\frac{ab}{\beta}\right\rfloor
&&\text{High word.}
\end{aligned}`;

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url: string, timeoutMs = 30000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Vite/Chromium startup race.
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
        baseUrl,
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json() as any[];
    const page = targets.find(
      (target) => target.type === "page" && target.url.startsWith(baseUrl),
    );
    assert.ok(page, "No VisualTeX Chromium target found");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(1000);
    const result = await client.send("Runtime.evaluate", {
      expression: `(async () => {
        const runtime = await import('${baseUrl}/src/export/runtime.ts');
        const geometry = await import('${baseUrl}/src/office/shared/mathTypeAlignmentGeometry.ts');
        const source = ${JSON.stringify(exactUserAligned)};
        const svg = runtime.latexToSvg(source, {
          displayMode: true,
          fontSizePt: 12,
          paddingPx: 0,
          paddingXPx: 0,
          paddingYPx: 6,
          background: 'transparent',
          formulaLetterFont: 'times',
        });
        const rawMathMl = runtime.latexToMathMl(source, true);
        const annotatedMathMl = geometry.annotateMathTypeAlignmentGeometry(
          rawMathMl,
          svg.svg,
          svg.width,
        );
        const doc = new DOMParser().parseFromString(annotatedMathMl, 'application/xml');
        const table = [...doc.getElementsByTagName('*')].find(
          (element) => element.localName === 'mtable'
        );
        return {
          stops: table?.getAttribute('data-visualtex-mtef-ruler-stops') ?? '',
          columnalign: table?.getAttribute('columnalign') ?? '',
          rowCount: table ? [...table.children].filter((e) => e.localName === 'mtr').length : 0,
          svgWidth: svg.width,
          annotatedMathMl,
        };
      })()`,
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
    const value = result.result.value as {
      stops: string;
      columnalign: string;
      rowCount: number;
      svgWidth: number;
      annotatedMathMl: string;
    };
    const stops = value.stops
      .split(",")
      .filter(Boolean)
      .map(Number);
    assert.deepEqual(value.columnalign.trim().split(/\s+/), ["right", "left", "right", "left"]);
    assert.equal(value.rowCount, 3);
    assert.equal(stops.length, 2, `Expected two MathType ruler stops, got '${value.stops}'`);
    assert.ok(stops[0] > 0 && stops[1] > stops[0]);
    assert.ok(stops.every((stop) => Number.isInteger(stop) && stop <= 0xffff));
    assert.ok(value.annotatedMathMl.includes('data-visualtex-mtef-ruler-stops='));
    console.log(
      `MathType aligned SVG geometry acceptance passed: rulerStops=${stops.join('/')}, ` +
        `rows=${value.rowCount}, svgWidth=${value.svgWidth.toFixed(3)}px.`,
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
