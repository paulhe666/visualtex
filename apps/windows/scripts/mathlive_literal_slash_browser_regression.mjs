import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const port = 18000 + (process.pid % 1000);
const debugPort = 23000 + (process.pid % 1000);
const baseUrl = "http://127.0.0.1:" + port;
const chromePath = resolveChromiumExecutable();
const profile = createBrowserProfilePath("visualtex-literal-slash");
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return response;
    } catch {
      // Preview or Chromium is still starting.
    }
    await sleep(80);
  }
  throw new Error("Timed out waiting for " + url);
}

const preview = spawn(
  process.execPath,
  [
    "node_modules/vite/bin/vite.js",
    "preview",
    "--host",
    "127.0.0.1",
    "--port",
    String(port),
    "--strictPort",
  ],
  { cwd: process.cwd(), stdio: "ignore" },
);

let chrome;
let socket;
let nextId = 1;
const pending = new Map();

const send = (method, params = {}) =>
  new Promise((resolve, reject) => {
    const id = nextId++;
    pending.set(id, { resolve, reject });
    socket.send(JSON.stringify({ id, method, params }));
  });

const evaluate = async (expression) => {
  const message = await send("Runtime.evaluate", {
    expression,
    awaitPromise: true,
    returnByValue: true,
  });
  if (message.error || message.result?.exceptionDetails) {
    throw new Error(
      message.error?.message ??
        message.result?.exceptionDetails?.exception?.description ??
        message.result?.exceptionDetails?.text ??
        "Runtime.evaluate failed",
    );
  }
  return message.result?.result?.value;
};

async function dispatchSlash(code, windowsVirtualKeyCode) {
  await send("Input.dispatchKeyEvent", {
    type: "keyDown",
    key: "/",
    code,
    windowsVirtualKeyCode,
    nativeVirtualKeyCode: windowsVirtualKeyCode,
    text: "/",
    unmodifiedText: "/",
  });
  await send("Input.dispatchKeyEvent", {
    type: "keyUp",
    key: "/",
    code,
    windowsVirtualKeyCode,
    nativeVirtualKeyCode: windowsVirtualKeyCode,
  });
  await sleep(60);
}

try {
  await waitFor(baseUrl);
  chrome = spawn(
    chromePath,
    [
      "--headless=new",
      "--disable-gpu",
      "--no-first-run",
      "--no-default-browser-check",
      "--remote-debugging-port=" + debugPort,
      "--user-data-dir=" + profile,
      baseUrl,
    ],
    { stdio: "ignore" },
  );

  const targetResponse = await waitFor(
    "http://127.0.0.1:" + debugPort + "/json/list",
  );
  const targets = await targetResponse.json();
  const target = targets.find(
    (item) => item.type === "page" && item.url.startsWith(baseUrl),
  );
  assert.ok(target, "No VisualTeX preview target was found");

  socket = new WebSocket(target.webSocketDebuggerUrl);
  socket.addEventListener("message", (event) => {
    const message = JSON.parse(event.data);
    if (!message.id || !pending.has(message.id)) return;
    const request = pending.get(message.id);
    pending.delete(message.id);
    if (message.error) request.reject(new Error(message.error.message));
    else request.resolve(message);
  });
  await new Promise((resolve, reject) => {
    socket.addEventListener("open", resolve, { once: true });
    socket.addEventListener("error", reject, { once: true });
  });

  await sleep(1200);
  await evaluate(
    '(async () => {' +
      'await customElements.whenDefined("math-field");' +
      'document.querySelector("[data-visualtex-slash-regression]")?.remove();' +
      'const field=document.createElement("math-field");' +
      'field.dataset.visualtexSlashRegression="1";' +
      'field.style.cssText="position:fixed;left:10px;top:10px;width:400px;height:80px;z-index:100000;background:white";' +
      'document.body.append(field);' +
      'await new Promise((resolve)=>requestAnimationFrame(resolve));' +
      'window.__visualtexSlashRegressionField=field;' +
      'return true;' +
    '})()',
  );

  for (const shortcutsEnabled of [true, false]) {
    const shortcutsSource = shortcutsEnabled
      ? '{ alpha: "\\\\alpha", ">=": "\\\\ge" }'
      : "{}";
    await evaluate(
      "(() => {" +
        "const field=window.__visualtexSlashRegressionField;" +
        'field.setValue("a",{mode:"math",format:"latex"});' +
        "field.inlineShortcuts=" + shortcutsSource + ";" +
        "field.position=field.lastOffset;" +
        "field.focus();" +
        'field.shadowRoot?.querySelector(\'[part="keyboard-sink"]\')?.focus({preventScroll:true});' +
        "return field.value;" +
      "})()",
    );

    await dispatchSlash("Slash", 191);
    const slashValue = await evaluate(
      "window.__visualtexSlashRegressionField.value",
    );
    assert.equal(
      slashValue.replace(/\s+/g, ""),
      "a/",
      "Slash became structural with inline shortcuts " +
        (shortcutsEnabled ? "enabled" : "disabled") +
        ": " +
        slashValue,
    );
    assert.ok(
      !slashValue.includes("\\frac"),
      "Slash unexpectedly produced a fraction with inline shortcuts " +
        (shortcutsEnabled ? "enabled" : "disabled") +
        ": " +
        slashValue,
    );

    await evaluate(
      "(() => {" +
        "const field=window.__visualtexSlashRegressionField;" +
        'field.setValue("a",{mode:"math",format:"latex"});' +
        "field.position=field.lastOffset;" +
        "field.focus();" +
        'field.shadowRoot?.querySelector(\'[part="keyboard-sink"]\')?.focus({preventScroll:true});' +
      "})()",
    );
    await dispatchSlash("NumpadDivide", 111);
    const numpadValue = await evaluate(
      "window.__visualtexSlashRegressionField.value",
    );
    assert.equal(
      numpadValue.replace(/\s+/g, ""),
      "a/",
      "NumpadDivide became structural with inline shortcuts " +
        (shortcutsEnabled ? "enabled" : "disabled") +
        ": " +
        numpadValue,
    );
  }

  const explicitFraction = await evaluate(
    "(() => {" +
      "const field=window.__visualtexSlashRegressionField;" +
      'field.setValue("",{mode:"math",format:"latex"});' +
      'field.insert("\\\\frac{1}{2}",{mode:"math",format:"latex"});' +
      "const model=field._mathfield?.model;" +
      "const atoms=model?.getAtoms?.([0,field.lastOffset],{includeChildren:true})??[];" +
      "return {value:field.value,hasGenfrac:atoms.some((atom)=>atom?.type===\"genfrac\")};" +
    "})()",
  );
  assert.ok(
    explicitFraction?.hasGenfrac,
    "Explicit \\frac insertion lost its fraction structure: " +
      JSON.stringify(explicitFraction),
  );

  console.log("MathLive literal slash browser regression passed.");
} finally {
  for (const request of pending.values()) {
    request.reject(new Error("Browser regression closed before CDP response"));
  }
  pending.clear();
  socket?.close();
  chrome?.kill("SIGTERM");
  preview.kill("SIGTERM");
  for (let attempt = 0; attempt < 5; attempt += 1) {
    try {
      await sleep(120);
      await rm(profile, { recursive: true, force: true });
      break;
    } catch (error) {
      if (attempt === 4) throw error;
    }
  }
}
