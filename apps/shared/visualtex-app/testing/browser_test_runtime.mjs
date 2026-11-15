// The browser side of the regressions both apps run, from the app folder: Chrome (or Edge) located per system, a
// throwaway profile, a CDP client, withAppBrowser (the app's built dist or its Vite dev server in a headless browser)
// and withBrowserTest (a blank page of the dev server, for tests that import the app's modules).
import { existsSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import process from "node:process";
import { spawn } from "node:child_process";
import { mkdtemp, rm } from "node:fs/promises";
import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";

export function createBrowserProfilePath(name) {
  return join(tmpdir(), `${name}-${process.pid}`);
}

export function resolveChromiumExecutable() {
  const configuredPath = (process.env.CHROME_PATH || process.env.VISUALTEX_CHROME_PATH)?.trim();
  if (configuredPath) return configuredPath;

  const candidates =
    process.platform === "win32"
      ? [
          process.env.PROGRAMFILES &&
            join(
              process.env.PROGRAMFILES,
              "Google",
              "Chrome",
              "Application",
              "chrome.exe",
            ),
          process.env["PROGRAMFILES(X86)"] &&
            join(
              process.env["PROGRAMFILES(X86)"],
              "Google",
              "Chrome",
              "Application",
              "chrome.exe",
            ),
          process.env.LOCALAPPDATA &&
            join(
              process.env.LOCALAPPDATA,
              "Google",
              "Chrome",
              "Application",
              "chrome.exe",
            ),
          process.env.PROGRAMFILES &&
            join(
              process.env.PROGRAMFILES,
              "Microsoft",
              "Edge",
              "Application",
              "msedge.exe",
            ),
          process.env["PROGRAMFILES(X86)"] &&
            join(
              process.env["PROGRAMFILES(X86)"],
              "Microsoft",
              "Edge",
              "Application",
              "msedge.exe",
            ),
        ]
      : process.platform === "darwin"
        ? [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            "/Applications/Chromium.app/Contents/MacOS/Chromium",
          ]
        : [
            "/usr/bin/google-chrome",
            "/usr/bin/google-chrome-stable",
            "/usr/bin/chromium",
            "/usr/bin/chromium-browser",
            "/usr/bin/microsoft-edge",
          ];

  const installedPath = candidates.find(
    (candidate) => candidate && existsSync(candidate),
  );
  if (installedPath) return installedPath;

  throw new Error(
    "Chrome, Edge, or Chromium was not found. Set CHROME_PATH to the browser executable.",
  );
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// One CDP connection to a page: send (each command with a timeout), evaluate, and the page's console messages and
// uncaught exceptions in `events` (the raw Runtime.consoleAPICalled / Runtime.exceptionThrown messages).
export class CdpClient {
  constructor(url) {
    this.url = url;
    this.nextId = 1;
    this.pending = new Map();
    this.events = [];
  }

  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((resolve, reject) => {
      this.socket.addEventListener("open", resolve, { once: true });
      this.socket.addEventListener("error", reject, { once: true });
    });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
      if (!message.id) {
        if (message.method === "Runtime.exceptionThrown" || message.method === "Runtime.consoleAPICalled") {
          this.events.push(message);
        }
        return;
      }
      const call = this.pending.get(message.id);
      if (!call) return;
      this.pending.delete(message.id);
      clearTimeout(call.timer);
      if (message.error) call.reject(new Error(message.error.message));
      else call.resolve(message.result);
    });
  }

  // the app's first load in a fresh Vite dev server takes ~20 s on a slow machine
  send(method, params = {}, timeoutMs = 120000) {
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`Browser command timed out: ${method}`));
      }, timeoutMs);
      this.pending.set(id, { resolve, reject, timer });
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }

  async evaluate(expression, timeoutMs) {
    const result = await this.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true }, timeoutMs);
    if (result.exceptionDetails) {
      throw new Error(
        result.exceptionDetails.exception?.description || result.exceptionDetails.text || "Runtime.evaluate failed",
      );
    }
    return result.result.value;
  }

  close() {
    for (const call of this.pending.values()) clearTimeout(call.timer);
    this.pending.clear();
    this.socket?.close();
  }
}

// A headless browser with a throwaway profile, opened at `url`, and a CDP client on that page. The browser picks
// its debugging port and writes it to DevToolsActivePort in the profile, so parallel tests never collide.
// windowSize null keeps the browser's own default size.
export async function launchBrowser(url, { windowSize = "1400,1000", args = [] } = {}) {
  const profile = await mkdtemp(join(tmpdir(), "visualtex-browser-test-"));
  const browser = spawn(resolveChromiumExecutable(), [
    "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
    "--remote-debugging-port=0", `--user-data-dir=${profile}`,
    ...(windowSize === null ? [] : [`--window-size=${windowSize}`]), ...args, url,
  ], { stdio: "ignore" });
  browser.on("error", () => {}); // Report a missing process below with the startup context.
  const close = async () => {
    if (browser.pid && browser.exitCode === null) {
      await new Promise((resolve) => {
        const timer = setTimeout(() => browser.kill("SIGKILL"), 5000);
        browser.once("exit", () => { clearTimeout(timer); resolve(); });
        browser.kill("SIGTERM");
      });
    }
    await rm(profile, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  };
  try {
    let debugPort;
    let target;
    const deadline = Date.now() + 15000;
    while (!target && Date.now() < deadline) {
      try {
        debugPort ??= Number(readFileSync(join(profile, "DevToolsActivePort"), "utf8").split("\n")[0]) || undefined;
        if (debugPort) {
          const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
          const pages = targets.filter((entry) => entry.type === "page");
          target = pages.find((entry) => entry.url.startsWith(url)) ?? (url === "about:blank" ? pages[0] : undefined);
        }
      } catch { /* The browser and its debugging endpoint are still starting. */ }
      if (!target) await sleep(50);
    }
    if (!target) throw new Error(`The browser did not open ${url} within 15 seconds`);
    const client = new CdpClient(target.webSocketDebuggerUrl);
    await client.connect();
    // a page found this way is not the focused one yet (document.hasFocus() is false; element.focus() does nothing)
    await client.send("Page.bringToFront");
    return {
      client,
      debugPort,
      close: async () => {
        client.close();
        await close();
      },
    };
  } catch (error) {
    await close();
    throw error;
  }
}

// Vite as installed in the app the test runs in (apps/shared has no node_modules); its configuration is the app's.
async function appVite() {
  const vitePath = createRequire(join(process.cwd(), "package.json")).resolve("vite");
  return import(pathToFileURL(vitePath).href);
}

// A page of the app's origin without the app: localStorage written there is in place before the app first loads.
export const BLANK_PAGE_PATH = "/__visualtex_test_blank";

// ahead of Vite's own middleware, whose HTML fallback would answer with the app
const blankPagePlugin = {
  name: "visualtex-test-blank-page",
  configureServer(server) {
    server.middlewares.use(serveBlankPage);
  },
  configurePreviewServer(server) {
    server.middlewares.use(serveBlankPage);
  },
};

function serveBlankPage(request, response, next) {
  if (request.url !== BLANK_PAGE_PATH) return next();
  response.setHeader("Content-Type", "text/html");
  response.end("<!doctype html><html><body></body></html>");
}

// The app the test runs in, served on a free port: its built dist by `vite preview` (server "preview"; `outDir`
// previews another build than dist/) or its sources by the Vite dev server (server "dev"). Both also serve
// BLANK_PAGE_PATH.
export async function startAppServer({ server = "preview", outDir } = {}) {
  const vite = await appVite();
  const listen = { host: "127.0.0.1", port: 0, strictPort: false };
  const http =
    server === "dev"
      ? await vite.createServer({ logLevel: "error", server: listen, plugins: [blankPagePlugin] })
      : await vite.preview({
          logLevel: "error",
          preview: listen,
          plugins: [blankPagePlugin],
          ...(outDir ? { build: { outDir } } : {}),
        });
  try {
    if (server === "dev") await http.listen();
  } catch (error) {
    await http.close();
    throw error;
  }
  return { baseUrl: `http://127.0.0.1:${http.httpServer.address().port}`, close: () => http.close() };
}

// The app (startAppServer) in a headless browser driven over CDP. The page opens at `path` of the app ("/" by
// default) or at "about:blank". run gets { baseUrl, client, send, evaluate, debugPort } with Runtime and Page
// enabled (debugPort reaches the browser's other pages).
export async function withAppBrowser(run, { server = "preview", path = "/", outDir, windowSize, args } = {}) {
  const app = await startAppServer({ server, outDir });
  let browser;
  try {
    const { baseUrl } = app;
    browser = await launchBrowser(path === "about:blank" ? path : baseUrl + path, { windowSize, args });
    const { client } = browser;
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await run({
      baseUrl,
      client,
      send: (method, params) => client.send(method, params),
      evaluate: (expression) => client.evaluate(expression),
      debugPort: browser.debugPort,
    });
  } finally {
    await browser?.close();
    await app.close();
  }
}

// Shared by regressions that need the actual application and isolated DOM probes.
export async function withBrowserTest(run) {
  // a dev server of its own: it serves a blank page besides the app
  const vite = await appVite();
  const server = await vite.createServer({
    logLevel: "error",
    server: { host: "127.0.0.1", port: 0, strictPort: false },
    plugins: [blankPagePlugin],
  });
  let browser;
  try {
    await server.listen();
    const baseUrl = `http://127.0.0.1:${server.httpServer.address().port}`;
    browser = await launchBrowser(`${baseUrl}${BLANK_PAGE_PATH}`);
    const { client } = browser;
    const send = (method, params) => client.send(method, params, 60000);
    const evaluate = (expression) => client.evaluate(expression, 60000);
    await send("Page.enable");
    // The target lists the blank page's URL before the page commits; until then the document is about:blank and
    // a test's first import of a /src or /@fs module cannot resolve.
    const committedBy = Date.now() + 15000;
    while (!String(await evaluate("location.href")).startsWith(baseUrl)) {
      if (Date.now() > committedBy) throw new Error("Browser test page did not load within 15 seconds");
      await sleep(50);
    }
    await run({ baseUrl, send, evaluate });
  } finally {
    await browser?.close();
    await server.close();
  }
}
