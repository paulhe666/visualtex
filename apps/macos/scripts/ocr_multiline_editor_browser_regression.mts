import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";
import { VISUALTEX_ALIGNMENT_MARKER_LATEX } from "../src/editor/alignmentMarkers.ts";

const portOffset = process.pid % 700;
const previewPort = 6900 + portOffset;
const debugPort = 16900 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = createBrowserProfilePath("visualtex-macos-ocr-multiline-editor");
const chromePath = resolveChromiumExecutable();
const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url: string, timeoutMs = 20_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while local Vite/Chromium processes start.
    }
    await sleep(100);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

class CdpClient {
  nextId = 1;
  pending = new Map<number, { resolve: (value: any) => void; reject: (error: Error) => void }>();
  socket?: WebSocket;

  constructor(readonly url: string) {}

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
    for (let attempt = 0; attempt < 80 && !page; attempt += 1) {
      const targets = (await (
        await fetch(`http://127.0.0.1:${debugPort}/json/list`)
      ).json()) as any[];
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

    const evaluate = async (expression: string) => {
      const result = await client!.send("Runtime.evaluate", {
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
      localStorage.setItem('visualtex.onboarding.v3.completed', 'true');
      localStorage.setItem('visualtex.onboarding.macos.desktop.v1.2.0.completed', 'true');
      localStorage.removeItem('visualtex-editor');
      return true;
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(700);

    await evaluate(`new Promise((resolve, reject) => {
      const started = Date.now();
      const poll = () => {
        if (document.querySelector('.formula-line math-field')) return resolve(true);
        if (Date.now() - started > 10000) return reject(new Error('math-field did not mount'));
        setTimeout(poll, 30);
      };
      poll();
    })`);

    const dispatch = await evaluate(`(() => {
      const provider = {
        activeProvider: 'openai-compatible',
        openAiCompatible: {
          protocol: 'responses',
          baseUrl: 'https://api.example.test/v1',
          model: 'vision-model',
          prompt: 'Return formula JSON',
          hasApiKey: true,
        },
        ollama: {
          baseUrl: 'http://127.0.0.1:11434',
          model: 'vision-model',
          prompt: 'Return formula JSON',
        },
        mathpix: {
          baseUrl: 'https://api.mathpix.com',
          appId: '',
          hasAppKey: false,
        },
        paddleOcr: {
          model: 'PaddleOCR-VL-1.6',
          hasAccessToken: false,
        }
      };
      const calls = [];
      window.__visualtexOcrMockCalls = calls;
      window.__TAURI_INTERNALS__ = {
        ...(window.__TAURI_INTERNALS__ || {}),
        invoke: async (command, args) => {
          calls.push({ command, args });
          if (command === 'get_ocr_provider_configuration') return provider;
          if (command === 'recognize_formula_image') {
            return {
              provider: 'openai-compatible',
              model: 'vision-model',
              elapsedMs: 19,
              processedWidth: 600,
              processedHeight: 220,
              backgroundInverted: false,
              backgroundLuminance: 0,
              formulas: [{ latex: '\\\\begin{align}x&=1\\\\\\\\longvariable&=2\\\\\\\\z&=3\\\\end{align}' }],
            };
          }
          throw new Error('Unexpected OCR mock command: ' + command);
        },
      };
      const field = document.querySelector('.formula-line math-field');
      field.focus();
      field.position = field.lastOffset;
      const bytes = new Uint8Array([137,80,78,71,13,10,26,10,0,0,0,0]);
      const file = new File([bytes], 'multiline.png', { type: 'image/png' });
      const transfer = new DataTransfer();
      transfer.items.add(file);
      const paste = new Event('paste', { bubbles: true, cancelable: true });
      Object.defineProperty(paste, 'clipboardData', { value: transfer });
      const accepted = field.dispatchEvent(paste);
      return { accepted, defaultPrevented: paste.defaultPrevented };
    })()`);
    assert.equal(dispatch.defaultPrevented, true, `image paste was not captured: ${JSON.stringify(dispatch)}`);

    let state: any = null;
    for (let attempt = 0; attempt < 120; attempt += 1) {
      state = await evaluate(`(() => {
        const fields = [...document.querySelectorAll('.formula-line math-field')];
        const entries = fields.map((field) => {
          const marker = field.shadowRoot?.querySelector('.visualtex-align-marker');
          const markerLeft = marker?.getBoundingClientRect().left ?? null;
          const equalCandidates = [];
          for (let position = 0; position <= field.lastOffset; position += 1) {
            const info = field.getElementInfo(position);
            if (info?.latex?.trim() === '=' && info.bounds && info.depth === 0) {
              equalCandidates.push(info.bounds.left);
            }
          }
          const equalLeft = typeof markerLeft === 'number'
            ? equalCandidates
                .filter((left) => left >= markerLeft - 0.1)
                .sort((a, b) => a - b)[0] ?? null
            : null;
          return {
            value: field.value,
            markerLeft,
            equalLeft,
          };
        });
        const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}');
        const markerLefts = entries.map((entry) => entry.markerLeft).filter(Number.isFinite);
        const equalLefts = entries.map((entry) => entry.equalLeft).filter(Number.isFinite);
        return {
          count: fields.length,
          entries,
          markerSpread: markerLefts.length ? Math.max(...markerLefts) - Math.min(...markerLefts) : null,
          visualAnchorSpread: equalLefts.length ? Math.max(...equalLefts) - Math.min(...equalLefts) : null,
          latexCodeFormat: persisted?.state?.latexCodeFormat ?? null,
          calls: window.__visualtexOcrMockCalls || [],
        };
      })()`);
      if (
        state.count === 3 &&
        state.entries.every((entry: any) => entry.value.includes(VISUALTEX_ALIGNMENT_MARKER_LATEX)) &&
        state.latexCodeFormat === 'aligned'
      ) {
        break;
      }
      await sleep(100);
    }

    assert.equal(state.count, 3, `OCR result was not expanded to three FormulaLines: ${JSON.stringify(state)}`);
    assert.ok(
      state.entries.every((entry: any) => entry.value.includes(VISUALTEX_ALIGNMENT_MARKER_LATEX)),
      `OCR align markers were not preserved in every line: ${JSON.stringify(state.entries)}`,
    );
    assert.equal(state.latexCodeFormat, "aligned", "OCR align did not switch the editor copy format to aligned");
    assert.ok(
      typeof state.visualAnchorSpread === "number" && state.visualAnchorSpread <= 0.75,
      `OCR-derived visible alignment atoms were not aligned: ${JSON.stringify(state)}`,
    );
    assert.ok(
      state.calls.some((call: any) => call.command === "get_ocr_provider_configuration"),
      `OCR provider was not consulted: ${JSON.stringify(state.calls)}`,
    );
    assert.ok(
      state.calls.some((call: any) => call.command === "recognize_formula_image"),
      `OCR recognition command was not invoked: ${JSON.stringify(state.calls)}`,
    );
    assert.ok(
      !state.calls.some((call: any) => call.command === "get_ocr_runtime_status"),
      `remote OCR incorrectly required the local runtime: ${JSON.stringify(state.calls)}`,
    );

    console.log(
      `macOS OCR multiline editor browser regression passed: rows=${state.count}, visualAnchorSpread=${state.visualAnchorSpread}px, provider=remote.`,
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
