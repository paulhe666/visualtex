import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 600;
const vitePort = 20300 + offset;
const debugPort = 26300 + offset;
const baseUrl = `http://127.0.0.1:${vitePort}`;
const chromeProfile = `/tmp/visualtex-custom-symbol-png-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while the local process starts.
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
      const timeout = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`CDP ${method} timed out`));
      }, 15000);
      this.pending.set(id, {
        resolve: (value) => {
          clearTimeout(timeout);
          resolve(value);
        },
        reject: (error) => {
          clearTimeout(timeout);
          reject(error);
        },
      });
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
      String(vitePort),
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
    assert.ok(page, "Chrome page target must exist");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(500);

    const result = await client.send("Runtime.evaluate", {
      expression: `(async () => {
        const BS = String.fromCharCode(92);
        const runtime = await import("/src/export/runtime.ts");
        const registry = await import("/src/math/customSymbolRegistry.ts");
        const svg = await runtime.latexToSvg(BS + "vtxtestsymbol", {
          displayMode: false,
          fontSizePt: 18,
          paddingPx: 8,
          background: "transparent",
        });
        const png = await runtime.svgToPng(svg, { scale: 2 });
        registry.replaceCustomSymbolLibrary({ version: 1, symbols: [] });
        const now = Date.now();
        registry.registerCustomSymbol({
          id: "png-erase-symbol",
          command: "pngerase",
          name: "PNG erased symbol",
          role: "ordinary",
          limitsBehavior: "auto",
          metrics: { widthEm: 0.9, ascentEm: 0.72, descentEm: 0.12 },
          artwork: {
            shapes: [
              { kind: "rect", x: 80, y: 80, width: 700, height: 600, fill: true },
              { kind: "path", operation: "erase", d: "M180 380L680 380", fill: false, strokeWidth: 160, lineCap: "round" },
            ],
          },
          ommlFallback: null,
          createdAt: now,
          updatedAt: now,
        });
        const erasedSvg = await runtime.latexToSvg(BS + "pngerase", {
          displayMode: false,
          fontSizePt: 18,
          paddingPx: 8,
          background: "transparent",
        });
        const erasedPng = await runtime.svgToPng(erasedSvg, { scale: 2 });
        return {
          svgHasArtwork: svg.svg.includes('data-visualtex-custom-symbol="vtxtestsymbol"'),
          svgHasCircle: /<circle\\b[^>]*cx="290"[^>]*cy="365"/.test(svg.svg),
          svgHasLine: /<line\\b[^>]*x1="65"[^>]*x2="515"/.test(svg.svg),
          blobType: png.blob.type,
          blobSize: png.blob.size,
          width: png.width,
          height: png.height,
          base64Length: png.base64.length,
          erasedSvgHasMask: /<mask\\b[^>]*id="visualtex-custom-symbol-erase-png-erase-symbol-/.test(erasedSvg.svg),
          erasedBlobType: erasedPng.blob.type,
          erasedBlobSize: erasedPng.blob.size,
          erasedWidth: erasedPng.width,
          erasedHeight: erasedPng.height,
          erasedBase64Length: erasedPng.base64.length,
        };
      })()`,
      awaitPromise: true,
      returnByValue: true,
    });
    if (result.exceptionDetails) {
      throw new Error(
        result.exceptionDetails.exception?.description ||
          result.exceptionDetails.text ||
          "PNG regression Runtime.evaluate failed",
      );
    }
    const value = result.result.value;
    assert.equal(value.svgHasArtwork, true);
    assert.equal(value.svgHasCircle, true);
    assert.equal(value.svgHasLine, true);
    assert.equal(value.blobType, "image/png");
    assert.ok(value.blobSize > 100, "Custom symbol PNG blob must contain raster data");
    assert.ok(value.width > 0 && value.height > 0, "Custom symbol PNG dimensions must be positive");
    assert.ok(value.base64Length > 100, "Custom symbol PNG base64 must be populated");
    assert.equal(value.erasedSvgHasMask, true, "Erased custom symbol SVG must retain its transparent mask before rasterization");
    assert.equal(value.erasedBlobType, "image/png");
    assert.ok(value.erasedBlobSize > 100, "Erased custom symbol PNG blob must contain raster data");
    assert.ok(value.erasedWidth > 0 && value.erasedHeight > 0, "Erased custom symbol PNG dimensions must be positive");
    assert.ok(value.erasedBase64Length > 100, "Erased custom symbol PNG base64 must be populated");

    console.log("Custom symbol prototype and vector-eraser PNG raster regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    vite.kill("SIGTERM");
    await sleep(220);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => undefined);
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
