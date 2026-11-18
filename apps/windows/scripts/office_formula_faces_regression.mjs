import assert from "node:assert/strict";
import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { extname, join, normalize } from "node:path";
import process from "node:process";
import { launchBrowser } from "../../shared/visualtex-app/testing/browser_test_runtime.mjs";

// The main app's fonts for letters, Greek letters and digits (Settings, over the math font) reach the Office editor
// with the other editor preferences, and the Office images are drawn in them. The Office page keeps its own storage
// (the companion's origin), where the layout engine reads the faces; it took the math font, Chinese font and spacing
// from the main app but not the faces, so with KaTeX as the math font and Times New Roman letters a VisualTeX OLE
// "x+y" came out in KaTeX's letters (user report 2026-11-17). The hidden converter (runtime=vsto-convert) draws the
// image an inserted or converted formula gets, for its session and for the render batches the agent tools and
// conversions send; this checks the SVGs it returns. The preferences come late here: the batches drew the first
// formulas before them, in the page's stored fonts (2026-11-18, in the installed app).

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const port = 22000 + (process.pid % 600);
const baseUrl = `http://127.0.0.1:${port}`;
const sessionId = "office-formula-faces-regression-session";
const converterUrl = `${baseUrl}/dialog/${sessionId}?runtime=vsto-convert`;
const distRoot = join(process.cwd(), "dist-office-windows-native");
const FACES = { letter: "Times New Roman", greek: "Times New Roman", digit: "Times New Roman" };

const mimeTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".woff2": "font/woff2",
};

const lineId = "faces-line";
let session = {
  id: sessionId,
  mode: "create",
  host: "word",
  formulaId: "faces-formula",
  sourceDocumentId: "faces-document",
  sourceObjectId: null,
  title: "Word Formula",
  lines: [{ id: lineId, latex: "x+\\alpha=2" }],
  activeLineId: lineId,
  codeFormat: "raw",
  displayMode: "inline",
  objectMode: "nativeOle",
  numbered: false,
  fontSizePt: 12,
  exportWidth: 0,
  exportHeight: 0,
  exportResult: null,
  originalMetadata: null,
  dirty: false,
  status: "created",
  autoCommitOnClose: false,
  explicitCancel: false,
  error: null,
  createdAt: Date.now(),
  updatedAt: Date.now(),
  expiresAt: Date.now() + 60_000,
};

const batchSession = { ...session, id: "faces-batch-session", formulaId: "faces-batch-formula", lines: [{ id: "batch-line", latex: "y+\\beta=3" }], activeLineId: "batch-line" };
// the visible editor (runtime=vsto-desktop) with a formula open
let editorSession = { ...session, id: "faces-editor-session", formulaId: "faces-editor-formula", lines: [{ id: "editor-line", latex: "x+y" }], activeLineId: "editor-line" };
let batchTaken = false;
let batchResult = null;
const PREFERENCES_DELAY_MS = 1200;

function writeJson(response, status, value) {
  response.writeHead(status, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" });
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
    if (url.pathname === "/api/v1/app/converter/next-render-batch") {
      const id = batchTaken ? null : "faces-batch";
      batchTaken = true;
      return writeJson(response, 200, { id });
    }
    if (url.pathname === "/api/v1/app/converter/render-batches/faces-batch") {
      return writeJson(response, 200, { id: "faces-batch", items: [{ id: "faces-item", session: batchSession }] });
    }
    if (url.pathname.startsWith("/api/v1/app/converter/render-batches/faces-batch/")) {
      batchResult = { kind: url.pathname.endsWith("/complete") ? "complete" : "fail", body: await readJsonBody(request) };
      return writeJson(response, 200, {});
    }
    if (url.pathname === "/api/v1/app/converter/next-batch") return writeJson(response, 200, { sessionIds: [] });
    if (url.pathname === "/api/v1/preferences") {
      await sleep(PREFERENCES_DELAY_MS);
      return writeJson(response, 200, {
        powerpointDefaultFontSizePt: 28,
        editorPreferences: {
          settings: { sourceOpen: false, language: "cn", formulaLetterFont: "katex", formulaFaces: FACES },
        },
      });
    }
    if (url.pathname === `/api/v1/sessions/${sessionId}`) {
      if (request.method === "PATCH") session = { ...session, ...(await readJsonBody(request)), updatedAt: Date.now() };
      return writeJson(response, 200, session);
    }
    if (url.pathname === `/api/v1/sessions/${editorSession.id}`) {
      if (request.method === "PATCH") editorSession = { ...editorSession, ...(await readJsonBody(request)), updatedAt: Date.now() };
      return writeJson(response, 200, editorSession);
    }
    if (url.pathname === `/api/v1/app/sessions/${editorSession.id}/close`) return writeJson(response, 200, { closed: true });
    if (url.pathname === `/api/v1/app/sessions/${sessionId}/close`) return writeJson(response, 200, { closed: true });
    if (url.pathname.startsWith("/api/v1/ocr/")) return writeJson(response, 503, { error: "OCR is not needed in this regression" });
    if (url.pathname.startsWith("/dialog/")) {
      const source = await readFile(join(distRoot, "dialog", "index.html"), "utf8");
      const meta = [
        '<meta name="visualtex-install-token" content="faces-regression" />',
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

// A fresh browser (its own storage, as a new Office window) on an Office page, with the Tauri bridge WebView2 gives
// the page stood in.
async function openOfficePage(url) {
  const opened = await launchBrowser(url, { windowSize: "1100,800" });
  const { client } = opened;
  await client.send("Runtime.enable");
  await client.send("Page.enable");
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
  return opened;
}

await new Promise((resolve) => server.listen(port, "127.0.0.1", resolve));
let browser;
try {
  browser = await openOfficePage(converterUrl);
  const { client } = browser;

  for (let attempt = 0; attempt < 400 && !(session.exportResult?.svg && batchResult); attempt += 1) await sleep(60);
  const svg = session.exportResult?.svg;
  assert.ok(svg, `the converter committed no image (session status ${session.status}, error ${session.error})`);
  assert.equal(batchResult?.kind, "complete", `the render batch: ${JSON.stringify(batchResult)?.slice(0, 300)}`);
  const batchSvg = batchResult.body.items?.[0]?.export?.svg;
  assert.ok(batchSvg, `the render batch returned no image: ${JSON.stringify(batchResult.body).slice(0, 300)}`);

  // the faces arrived in the page's own storage, where the layout engine reads them
  const stored = await client.evaluate(`localStorage.getItem("visualtex.formula-faces")`);
  assert.deepEqual(JSON.parse(stored ?? "null"), FACES, `the Office page stored faces ${stored}`);

  // a letter and a Greek letter in their faces' italic, a digit in the digit face's upright; + and = stay KaTeX's
  const checkFaces = (image, label, [letter, greek, digit], [letterCode, greekCode, digitCode]) => {
    const texts = [...image.matchAll(/<text\b([^>]*)>([^<]*)<\/text>/g)].map(([, attributes, text]) => ({
      text,
      family: /font-family="([^"]*)"/.exec(attributes)?.[1]?.replace(/&quot;/g, '"') ?? "",
      italic: /font-style="italic"/.test(attributes),
    }));
    for (const [character, italic] of [[letter, true], [greek, true], [digit, false]]) {
      const entry = texts.find((item) => item.text === character);
      assert.ok(entry, `${label}: ${character} is not drawn in its face (texts: ${JSON.stringify(texts)})`);
      assert.match(entry.family, /Times New Roman/, `${label}: ${character} is drawn in ${entry.family}`);
      assert.equal(entry.italic, italic, `${label}: ${character} italic: ${entry.italic}`);
    }
    // WPS draws a quoted font name ("Times New Roman", serif) in a fallback font: the names are written bare
    for (const entry of texts) {
      assert.doesNotMatch(entry.family, /["']/, `${label}: ${entry.text} names its font in quotes (${entry.family})`);
    }
    // KaTeX's letters (cmmi10) and digits (cmr10) are not drawn
    assert.doesNotMatch(image, new RegExp(`cmmi10-(${letterCode}|${greekCode})"`), `${label}: ${letter} or ${greek} is drawn from KaTeX's math italic outlines`);
    assert.doesNotMatch(image, new RegExp(`cmr10-${digitCode}"`), `${label}: ${digit} is drawn from KaTeX's roman outlines`);
    assert.match(image, /cmr10-43"/, `${label}: + should stay KaTeX's (cmr10)`);
  };
  checkFaces(svg, "session image", ["x", "α", "2"], [120, 11, 50]);
  checkFaces(batchSvg, "render batch image", ["y", "β", "3"], [121, 12, 51]);
  await browser.close();
  browser = undefined;

  // the visible editor shows the formula in the same fonts: x and y in Times New Roman, + in KaTeX's
  browser = await openOfficePage(`${baseUrl}/dialog/${editorSession.id}?runtime=vsto-desktop`);
  const shown = await (async () => {
    let last;
    for (let attempt = 0; attempt < 300; attempt += 1) {
      last = await browser.client.evaluate(`(() => {
        const field = document.querySelector("math-field.visual-mathfield");
        const spans = [...(field?.shadowRoot?.querySelectorAll('span[style*="font:"]') ?? [])];
        const glyphs = spans.map((span) => ({ text: span.textContent, font: span.style.font || span.getAttribute("style") }));
        return { ready: glyphs.some((g) => g.text === "y"), glyphs };
      })()`);
      if (last?.ready) return last;
      await sleep(60);
    }
    return last;
  })();
  assert.ok(shown?.ready, `the editor did not show the formula: ${JSON.stringify(shown)?.slice(0, 400)}`);
  for (const character of ["x", "y"]) {
    const glyph = shown.glyphs.find((g) => g.text === character);
    assert.ok(glyph && /Times New Roman/.test(glyph.font) && /italic/.test(glyph.font), `the editor shows ${character} in ${glyph?.font}`);
  }
  console.log("Office editor and images (session and render batch, preferences late) use the main app's letter, Greek and digit fonts over KaTeX: passed");
} finally {
  await browser?.close();
  await new Promise((resolve) => server.close(resolve));
}
