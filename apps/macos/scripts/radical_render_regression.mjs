import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { mkdir, rm, writeFile } from "node:fs/promises";
import process from "node:process";

const portOffset = process.pid % 1000;
const previewPort = 7600 + portOffset;
const debugPort = 12600 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = `/tmp/visualtex-radical-render-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const screenshotPath = "build-logs/radical-render-regression.png";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while the local process starts.
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
        "--window-size=1400,900",
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
    if (!page) throw new Error("No VisualTeX Chrome page target found");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(650);

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
            "Runtime.evaluate failed",
        );
      }
      return result.result.value;
    };

    await evaluate(`new Promise((resolve) => {
      const done = () => {
        if (location.origin === ${JSON.stringify(baseUrl)}) resolve(true);
        else setTimeout(done, 30);
      };
      done();
    })`);

    await evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
      localStorage.setItem(
        "visualtex.onboarding.macos.desktop.v1.2.0.completed",
        "true",
      );
      localStorage.setItem(
        "visualtex.office.macos.native-first-run.v1.2.0.completed",
        "true",
      );
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(700);
    await evaluate(`new Promise((resolve) => {
      const done = () => document.querySelector("math-field") ? resolve(true) : setTimeout(done, 30);
      done();
    })`);

    const inspection = await evaluate(`new Promise((resolve) => {
      const field = document.querySelector("math-field");
      field.setValue("x=\\\\frac{-b\\\\pm\\\\sqrt{b^2-4ac}}{2a}", {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
      field.blur();
      requestAnimationFrame(() => requestAnimationFrame(() => {
        const root = field.shadowRoot;
        const sign = root?.querySelector(".ML__sqrt-sign");
        const line = root?.querySelector(".ML__sqrt-line");
        const readNode = (node) => {
          if (!node) return null;
          const box = node.getBoundingClientRect();
          const style = getComputedStyle(node);
          return {
            className: node.className,
            text: node.textContent,
            box: {
              left: box.left,
              top: box.top,
              right: box.right,
              bottom: box.bottom,
              width: box.width,
              height: box.height,
            },
            display: style.display,
            position: style.position,
            overflow: style.overflow,
            color: style.color,
            background: style.backgroundColor,
            height: style.height,
            minHeight: style.minHeight,
            width: style.width,
            transform: style.transform,
          };
        };
        const readPseudo = (node, selector) => {
          if (!node) return null;
          const style = getComputedStyle(node, selector);
          return {
            content: style.content,
            display: style.display,
            background: style.backgroundColor,
            minHeight: style.minHeight,
          };
        };
        const fieldBox = field.getBoundingClientRect();
        resolve({
          value: field.value,
          fieldBox: {
            left: fieldBox.left,
            top: fieldBox.top,
            right: fieldBox.right,
            bottom: fieldBox.bottom,
            width: fieldBox.width,
            height: fieldBox.height,
          },
          sign: readNode(sign),
          line: readNode(line),
          lineBefore: readPseudo(line, "::before"),
        });
      }));
    })`);

    const clip = {
      x: Math.max(0, inspection.fieldBox.left),
      y: Math.max(0, inspection.fieldBox.top),
      width: Math.max(1, inspection.fieldBox.width),
      height: Math.max(1, inspection.fieldBox.height),
      scale: 1,
    };
    const screenshot = await client.send("Page.captureScreenshot", {
      format: "png",
      fromSurface: true,
      captureBeyondViewport: true,
      clip,
    });
    await mkdir("build-logs", { recursive: true });
    await writeFile(screenshotPath, Buffer.from(screenshot.data, "base64"));

    const roofHeight = Math.max(1, inspection.line.box.height || 1);
    const sampleRegion = {
      x: inspection.line.box.left - inspection.fieldBox.left + 2,
      // MathLive's nested vlist reports a layout box several pixels below the
      // final raster position. Anchor the pixel check to the radical glyph top,
      // which matches the visible roof in the captured image.
      y: inspection.sign.box.top - inspection.fieldBox.top - 1,
      width: Math.max(1, inspection.line.box.width - 4),
      height: roofHeight + 3,
    };
    const visualCheck = await evaluate(`new Promise((resolve, reject) => {
      const image = new Image();
      image.onload = () => {
        const canvas = document.createElement("canvas");
        canvas.width = image.naturalWidth;
        canvas.height = image.naturalHeight;
        const context = canvas.getContext("2d", { willReadFrequently: true });
        context.drawImage(image, 0, 0);
        const scaleX = image.naturalWidth / ${clip.width};
        const scaleY = image.naturalHeight / ${clip.height};
        const region = ${JSON.stringify(sampleRegion)};
        const startX = Math.max(0, Math.floor(region.x * scaleX));
        const endX = Math.min(
          image.naturalWidth,
          Math.ceil((region.x + region.width) * scaleX),
        );
        const startY = Math.max(0, Math.floor(region.y * scaleY));
        const endY = Math.min(
          image.naturalHeight,
          Math.ceil((region.y + region.height) * scaleY),
        );
        const pixels = context.getImageData(
          startX,
          startY,
          Math.max(1, endX - startX),
          Math.max(1, endY - startY),
        );
        let coveredColumns = 0;
        for (let x = 0; x < pixels.width; x += 1) {
          let hasDarkPixel = false;
          for (let y = 0; y < pixels.height; y += 1) {
            const index = (y * pixels.width + x) * 4;
            const red = pixels.data[index];
            const green = pixels.data[index + 1];
            const blue = pixels.data[index + 2];
            const alpha = pixels.data[index + 3];
            if (alpha > 100 && red < 120 && green < 120 && blue < 120) {
              hasDarkPixel = true;
              break;
            }
          }
          if (hasDarkPixel) coveredColumns += 1;
        }

        resolve({
          imageWidth: image.naturalWidth,
          imageHeight: image.naturalHeight,
          sampledWidth: pixels.width,
          sampledHeight: pixels.height,
          sampledStartY: startY,
          sampledEndY: endY,
          coveredColumns,
          coverage: coveredColumns / Math.max(1, pixels.width),
        });
      };
      image.onerror = () => reject(new Error("Could not decode CDP screenshot"));
      image.src = ${JSON.stringify(`data:image/png;base64,${screenshot.data}`)};
    })`);

    console.log(
      JSON.stringify(
        {
          screenshotPath,
          radicalWidth: inspection.line.box.width,
          radicalHeight: inspection.line.box.height,
          screenshotCoverage: visualCheck.coverage,
        },
        null,
        2,
      ),
    );

    assert.ok(inspection.sign, "MathLive did not render a radical sign");
    assert.ok(inspection.line, "MathLive did not render a radical line node");
    assert.ok(inspection.line.box.width > 1, JSON.stringify(inspection, null, 2));
    assert.notEqual(
      inspection.lineBefore?.background,
      "rgba(0, 0, 0, 0)",
      JSON.stringify(inspection, null, 2),
    );
    assert.ok(
      Number.parseFloat(inspection.lineBefore?.minHeight ?? "0") >= 1,
      JSON.stringify(inspection, null, 2),
    );
    assert.ok(
      visualCheck.coverage >= 0.9,
      `Radical roof is not visually continuous in the screenshot: ${JSON.stringify(
        visualCheck,
      )}`,
    );
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(300);
    await rm(chromeProfile, { recursive: true, force: true });
  }
}

await main();
