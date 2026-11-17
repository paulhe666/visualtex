import assert from "node:assert/strict";
import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { extname, join, normalize } from "node:path";
import process from "node:process";
import { CdpClient, launchBrowser } from "../../shared/visualtex-app/testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const offset = process.pid % 600;
const port = 21400 + offset;
const baseUrl = `http://127.0.0.1:${port}`;
const sessionId = "office-candidate-sync-regression-session";
const officeUrl = `${baseUrl}/dialog/${sessionId}?runtime=vsto-desktop`;
const distRoot = join(process.cwd(), "dist-office-windows-native");

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


// The Office editor's candidates are the main app's: the main window publishes its candidate settings (Settings >
// 命令推荐: look, count, personalized order) and its command usage with the editor preferences, and the Office editor,
// which keeps its own storage on the companion's origin, takes them when it opens. Before, it always showed the
// default list of six whatever the main app had (the user has the AxMath-style row there).
const usageTime = 1_763_300_000_000;
const MAIN_USAGE = {
  frac: { commandId: "frac", useCount: 42, lastUsedAt: usageTime, recentUses: [usageTime], acceptedPrefixes: { "\\fr": 42 }, contextCounts: { typed: 42 }, pinned: false },
  "greek-alpha": { commandId: "greek-alpha", useCount: 7, lastUsedAt: usageTime - 1000, recentUses: [usageTime - 1000], acceptedPrefixes: {}, contextCounts: {}, pinned: true },
};
let session = createSession("word", 12);

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
    if (url.pathname === "/api/v1/theme") return writeJson(response, 200, { theme: "light", editorLayout: "standard" });
    if (url.pathname === "/api/v1/preferences") {
      return writeJson(response, 200, {
        powerpointDefaultFontSizePt: 28,
        editorPreferences: {
          settings: { sourceOpen: false, language: "cn", candidateStyle: "bar", suggestionCount: 9, personalize: false },
          commandUsage: MAIN_USAGE,
        },
      });
    }
    if (url.pathname === `/api/v1/sessions/${sessionId}`) {
      if (request.method === "PATCH") session = { ...session, ...(await readJsonBody(request)), updatedAt: Date.now() };
      return writeJson(response, 200, session);
    }
    if (url.pathname === `/api/v1/app/sessions/${sessionId}/close`) return writeJson(response, 200, { closed: true });
    if (url.pathname.startsWith("/api/v1/ocr/")) return writeJson(response, 503, { error: "OCR is not needed in this regression" });
    if (url.pathname.startsWith("/dialog/")) {
      const source = await readFile(join(distRoot, "dialog", "index.html"), "utf8");
      const meta = [
        '<meta name="visualtex-install-token" content="candidate-regression" />',
        '<meta name="visualtex-native-powerpoint-commit" content="false" />',
        '<meta name="visualtex-theme" content="light" />',
      ].join("\n");
      response.writeHead(200, { "Content-Type": "text/html; charset=utf-8", "Cache-Control": "no-store" });
      response.end(source.replace("</head>", `${meta}\n</head>`));
      return;
    }
    if (url.pathname.startsWith("/assets/")) {
      const relative = normalize(url.pathname.slice(1));
      if (relative.startsWith("..")) return response.writeHead(403).end();
      const content = await readFile(join(distRoot, relative));
      response.writeHead(200, { "Content-Type": mimeTypes[extname(relative)] ?? "application/octet-stream", "Cache-Control": "no-store" });
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


await new Promise((resolve) => server.listen(port, "127.0.0.1", resolve));
let browser;
try {
  await waitFor(`${baseUrl}/api/v1/theme`);
  browser = await launchBrowser(officeUrl, { windowSize: "1100,800" });
  const { client } = browser;
  await client.send("Runtime.enable");
  await client.send("Page.enable");
  // the page reads its window from the Tauri bridge WebView2 gives it; a plain browser needs a stand-in
  await client.send("Page.addScriptToEvaluateOnNewDocument", { source: `
    window.__TAURI_INTERNALS__ = window.__TAURI_INTERNALS__ || {
      metadata: { currentWindow: { label: "office" }, currentWebview: { label: "office", windowLabel: "office" } },
      transformCallback: (cb) => { const id = Math.floor(Math.random() * 1e9); window["_" + id] = cb; return id; },
      unregisterCallback: () => {}, convertFileSrc: (p) => p,
      invoke: async (cmd, args) => (cmd === "plugin:event|listen" ? args.handler : null),
    };
    window.__TAURI_EVENT_PLUGIN_INTERNALS__ = window.__TAURI_EVENT_PLUGIN_INTERNALS__ || { unregisterListener() {} };` });
  await client.send("Page.reload", { ignoreCache: true });
  process.on("exit", () => {
    if (!process.exitCode) return;
    const errors = client.events.map((e) => e.params?.exceptionDetails?.exception?.description ?? e.params?.args?.map((a) => a.value ?? a.description).join(" ")).filter(Boolean);
    if (errors.length) console.error("page errors:\n" + errors.slice(0, 8).join("\n"));
  });
  await waitForEvaluation(client, `(() => ({ ready: Boolean(document.querySelector("math-field.visual-mathfield")?._mathfield) }))()`, "the Office editor's formula field", 30_000);
  await sleep(600);

  // the look: the page carries the main app's style
  const look = await client.evaluate(`document.documentElement.dataset.candidateStyle ?? null`);
  assert.equal(look, "bar", `the Office editor shows the "${look}" candidates, the main app has "bar"`);

  // the count, the personalized order switch and the usage are the editor's own settings now (its storage)
  const stored = await waitForEvaluation(client, `(() => {
    try {
      const state = JSON.parse(localStorage.getItem("visualtex-editor") || "{}").state || {};
      return { ready: state.suggestionCount === 9, suggestionCount: state.suggestionCount, personalize: state.personalize,
        candidateStyle: state.candidateStyle, frac: state.usage?.frac?.useCount ?? null, alphaPinned: state.usage?.["greek-alpha"]?.pinned ?? null };
    } catch (error) { return { ready: false, error: String(error) }; }
  })()`, "the Office editor's stored candidate settings");
  assert.equal(stored.personalize, false, `personalized order: ${JSON.stringify(stored)}`);
  assert.equal(stored.candidateStyle, "bar", `stored look: ${JSON.stringify(stored)}`);
  assert.equal(stored.frac, 42, `the main app's usage of \\frac did not arrive: ${JSON.stringify(stored)}`);
  assert.equal(stored.alphaPinned, true, `the main app's pinned command did not arrive: ${JSON.stringify(stored)}`);

  // typing a command: the list is as long as the main app's count
  await client.evaluate(`(() => { const f = document.querySelector("math-field.visual-mathfield"); f.focus();
    f.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true }); return true; })()`);
  for (const ch of "\\s") {
    const [code, vk] = ch === "\\" ? ["Backslash", 220] : ["Key" + ch.toUpperCase(), ch.toUpperCase().charCodeAt(0)];
    await client.send("Input.dispatchKeyEvent", { type: "keyDown", key: ch, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, text: ch, unmodifiedText: ch });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", key: ch, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk });
    await sleep(60);
  }
  const list = await waitForEvaluation(client, `(() => {
    const popover = document.querySelector("#mathlive-suggestion-popover");
    const items = popover ? [...popover.querySelectorAll("li")].filter((li) => li.getBoundingClientRect().height > 0) : [];
    return { ready: items.length > 0, count: items.length };
  })()`, "the candidate list");
  assert.equal(list.count, 9, `the candidate list has ${list.count} items, the main app shows 9`);
  console.log("Office editor candidates follow the main app: look bar, 9 items, personalized order off, usage merged: passed");
} finally {
  await browser?.close();
  await new Promise((resolve) => server.close(resolve));
}
