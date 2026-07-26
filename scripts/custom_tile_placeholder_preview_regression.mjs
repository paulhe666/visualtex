import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 800;
const previewPort = 8100 + offset;
const debugPort = 15100 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}/editor`;
const chromeProfile = `/tmp/visualtex-custom-tile-placeholder-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while Vite or Chrome starts.
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

  close() {
    this.socket?.close();
  }
}

const tiles = [
  {
    id: "preview-frac",
    latex: String.raw`\frac{\placeholder{}}{\placeholder{}}`,
    expectedPlaceholders: 2,
  },
  {
    id: "preview-indefinite-integral",
    latex: String.raw`\int \placeholder{}\,\mathrm{d}\placeholder{}`,
    expectedPlaceholders: 2,
  },
  {
    id: "preview-definite-integral",
    latex: String.raw`\int_{\placeholder{}}^{\placeholder{}} \placeholder{}\,\mathrm{d}\placeholder{}`,
    expectedPlaceholders: 4,
  },
  {
    id: "preview-limit",
    latex: String.raw`\lim_{\placeholder{}\to\placeholder{}} \placeholder{}`,
    expectedPlaceholders: 3,
  },
  {
    id: "preview-sum",
    latex: String.raw`\sum_{\placeholder{}}^{\placeholder{}} \placeholder{}`,
    expectedPlaceholders: 3,
  },
];

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
        "--window-size=1440,1000",
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
    if (!page) throw new Error("No VisualTeX page target found");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(500);

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

    await evaluate(`(() => {
      localStorage.clear();
      for (const key of [
        "visualtex.onboarding.v3.completed",
        "visualtex.office.macos.first-run.v1.completed",
        "visualtex.onboarding.macos.desktop.v1.2.0.completed",
        "visualtex.office.macos.native-first-run.v1.2.0.completed",
      ]) localStorage.setItem(key, "true");
      localStorage.setItem("visualtex-editor", JSON.stringify({
        state: {
          title: "Placeholder preview regression",
          lines: [{ id: "preview-editor-line", latex: "" }],
          activeLineId: "preview-editor-line",
          language: "cn",
          zoom: 1,
        },
        version: 0,
      }));
      localStorage.setItem("visualtex-custom-formula-tiles", JSON.stringify({
        version: 3,
        sections: [{ id: "preview-section", name: "模板", createdAt: 0 }],
        tiles: ${JSON.stringify(
          tiles.map((tile, index) => ({
            id: tile.id,
            latex: tile.latex,
            sectionId: "preview-section",
            color: null,
            createdAt: index,
          })),
        )},
      }));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(650);

    await evaluate(`new Promise((resolve) => {
      const poll = () => document.querySelector('[data-toolbar-view="tiles"]')
        ? resolve(true)
        : setTimeout(poll, 25);
      poll();
    })`);
    await evaluate(`document.querySelector('[data-toolbar-view="tiles"]')?.click()`);
    await sleep(80);
    await evaluate(`document.querySelector('[data-tile-category="custom"]')?.click()`);
    await evaluate(`new Promise((resolve) => {
      const poll = () => document.querySelectorAll('.formula-tile-button.is-custom').length === ${tiles.length}
        ? resolve(true)
        : setTimeout(poll, 25);
      poll();
    })`);
    await sleep(350);

    const previews = await evaluate(`Array.from(document.querySelectorAll('.formula-tile-button.is-custom')).map((button) => {
      const preview = button.querySelector('.formula-tile-preview');
      const content = preview?.querySelector('.math-preview-fit-content');
      const previewRect = preview?.getBoundingClientRect();
      const contentRect = content?.getBoundingClientRect();
      const placeholders = Array.from(preview?.querySelectorAll('.visualtex-tile-placeholder') ?? []);
      return {
        id: button.dataset.formulaTileId ?? "",
        latex: button.dataset.formulaTileLatex ?? "",
        showPlaceholders: preview?.dataset.showPlaceholders ?? "",
        fitScale: Number(preview?.dataset.fitScale ?? "0"),
        previewOverflow: preview ? getComputedStyle(preview).overflow : "",
        contentFits:
          Boolean(previewRect && contentRect) &&
          contentRect.left >= previewRect.left - 1 &&
          contentRect.right <= previewRect.right + 1 &&
          contentRect.top >= previewRect.top - 1 &&
          contentRect.bottom <= previewRect.bottom + 1,
        previewRect: previewRect
          ? { width: previewRect.width, height: previewRect.height }
          : null,
        contentRect: contentRect
          ? { width: contentRect.width, height: contentRect.height }
          : null,
        placeholderCount: placeholders.length,
        placeholderRects: placeholders.map((node) => {
          const rect = node.getBoundingClientRect();
          const style = getComputedStyle(node);
          const rule = node.querySelector('.ML__rule');
          return {
            width: rect.width,
            height: rect.height,
            borderWidth: parseFloat(style.borderTopWidth),
            borderStyle: style.borderTopStyle,
            borderRadius: parseFloat(style.borderTopLeftRadius),
            backgroundColor: style.backgroundColor,
            ruleVisibility: rule ? getComputedStyle(rule).visibility : '',
          };
        }),
      };
    })`);

    for (const tile of tiles) {
      const previewState = previews.find((item) => item.id === tile.id);
      assert.ok(previewState, `Missing preview for ${tile.id}`);
      assert.equal(previewState.latex, tile.latex, JSON.stringify(previewState));
      assert.equal(previewState.showPlaceholders, "true", JSON.stringify(previewState));
      assert.ok(
        previewState.fitScale >= 0.799 && previewState.fitScale <= 1.201,
        JSON.stringify(previewState),
      );
      assert.equal(previewState.previewOverflow, "hidden", JSON.stringify(previewState));
      assert.equal(previewState.contentFits, true, JSON.stringify(previewState));
      assert.equal(
        previewState.placeholderCount,
        tile.expectedPlaceholders,
        JSON.stringify(previewState),
      );
      assert.ok(
        previewState.placeholderRects.every(
          ({
            width,
            height,
            borderWidth,
            borderStyle,
            borderRadius,
            backgroundColor,
            ruleVisibility,
          }) =>
            width >= 5.2 &&
            height > width * 1.3 &&
            borderWidth > 0 &&
            borderWidth <= 1.2 &&
            borderStyle === "solid" &&
            borderRadius === 0 &&
            backgroundColor === "rgba(0, 0, 0, 0)" &&
            ruleVisibility === "hidden",
        ),
        JSON.stringify(previewState),
      );
    }

    const indefinite = previews.find(
      (item) => item.id === "preview-indefinite-integral",
    );
    const definite = previews.find(
      (item) => item.id === "preview-definite-integral",
    );
    assert.notEqual(
      indefinite.placeholderCount,
      definite.placeholderCount,
      "Definite and indefinite integrals must remain visually distinguishable",
    );

    await evaluate(`document.querySelector('[data-formula-tile-id="preview-definite-integral"]')?.click()`);
    await sleep(100);
    const inserted = await evaluate(`(() => {
      const field = document.querySelector('math-field');
      const stored = JSON.parse(localStorage.getItem('visualtex-custom-formula-tiles') || '{}');
      return {
        value: field?.value ?? "",
        storedLatex: stored.tiles?.find((tile) => tile.id === 'preview-definite-integral')?.latex ?? "",
      };
    })()`);
    assert.ok(inserted.value.includes("\\placeholder"), JSON.stringify(inserted));
    assert.equal(inserted.value.includes("visualtex-tile-placeholder"), false, JSON.stringify(inserted));
    assert.equal(inserted.value.includes("\\rule"), false, JSON.stringify(inserted));
    assert.equal(inserted.storedLatex, tiles[2].latex, JSON.stringify(inserted));

    console.log("Custom tile placeholder preview regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(250);
    for (let attempt = 0; attempt < 4; attempt += 1) {
      try {
        await rm(chromeProfile, { recursive: true, force: true });
        break;
      } catch (error) {
        if (attempt === 3) throw error;
        await sleep(180);
      }
    }
  }
}

await main();
