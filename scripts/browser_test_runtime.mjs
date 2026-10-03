import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import process from "node:process";
import { spawn } from "node:child_process";
import { mkdtemp, rm } from "node:fs/promises";

export function createBrowserProfilePath(name) {
  return join(tmpdir(), `${name}-${process.pid}`);
}

export function resolveChromiumExecutable() {
  const configuredPath = process.env.CHROME_PATH?.trim();
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
            (process.env.CHROME_PATH || "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
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

// Shared by regressions that need the actual application and isolated DOM probes.
export async function withBrowserTest(run) {
  const { createServer } = await import("vite");
  const server = await createServer({
    logLevel: "error",
    server: { host: "127.0.0.1", port: 0, strictPort: false },
  });
  server.middlewares.use((request, response, next) => {
    if (request.url !== "/__visualtex_test_blank") return next();
    response.setHeader("Content-Type", "text/html");
    response.end('<!doctype html><html><body></body></html>');
  });
  const profile = await mkdtemp(join(tmpdir(), "visualtex-browser-test-"));
  const pending = new Map();
  let browser;
  let socket;
  let nextId = 1;
  try {
    await server.listen();
    const baseUrl = `http://127.0.0.1:${server.httpServer.address().port}`;
    const debugPort = 24000 + process.pid % 1000;
    browser = spawn(resolveChromiumExecutable(), [
      "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
      `--remote-debugging-port=${debugPort}`, `--user-data-dir=${profile}`,
      "--window-size=1400,1000", `${baseUrl}/__visualtex_test_blank`,
    ], { stdio: "ignore" });
    browser.on("error", () => {}); // Report a missing process below with the startup context.
    let target;
    const deadline = Date.now() + 15000;
    while (!target && Date.now() < deadline) {
      try {
        const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
        target = targets.find((entry) => entry.type === "page" && entry.url.startsWith(baseUrl));
      } catch { /* Browser debugging endpoint is still starting. */ }
      if (!target) await new Promise((resolve) => setTimeout(resolve, 50));
    }
    if (!target) throw new Error("Browser test did not start within 15 seconds");
    socket = new WebSocket(target.webSocketDebuggerUrl);
    await new Promise((resolve, reject) => {
      socket.addEventListener("open", resolve, { once: true });
      socket.addEventListener("error", reject, { once: true });
    });
    socket.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
      const call = pending.get(message.id);
      if (!call) return;
      pending.delete(message.id);
      clearTimeout(call.timer);
      if (message.error) call.reject(new Error(message.error.message));
      else call.resolve(message.result);
    });
    const send = (method, params = {}) => new Promise((resolve, reject) => {
      const id = nextId++;
      const timer = setTimeout(() => {
        pending.delete(id);
        reject(new Error(`Browser command timed out: ${method}`));
      }, 15000);
      pending.set(id, { resolve, reject, timer });
      socket.send(JSON.stringify({ id, method, params }));
    });
    const evaluate = async (expression) => {
      const result = await send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
      if (result.exceptionDetails) throw new Error(
        result.exceptionDetails.exception?.description ?? result.exceptionDetails.text,
      );
      return result.result.value;
    };
    await send("Page.enable");
    await run({ baseUrl, send, evaluate });
  } finally {
    for (const call of pending.values()) clearTimeout(call.timer);
    socket?.close();
    if (browser?.pid && browser.exitCode === null) {
      await new Promise((resolve) => {
        const timer = setTimeout(() => browser.kill("SIGKILL"), 5000);
        browser.once("exit", () => { clearTimeout(timer); resolve(); });
        browser.kill("SIGTERM");
      });
    }
    await server.close();
    await rm(profile, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  }
}
