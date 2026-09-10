import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const offset = process.pid % 1000;
const previewPort = 8400 + offset;
const debugPort = 18400 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const browserProfile = createBrowserProfilePath(
  "visualtex-mathlive-runtime-safety",
);
const browserPath = resolveChromiumExecutable();
const sleep = (milliseconds) =>
  new Promise((resolve) => setTimeout(resolve, milliseconds));

async function waitFor(url, timeoutMs = 15_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return response;
    } catch {
      // Retry while the preview server or browser starts.
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
        if (
          message.method === "Runtime.exceptionThrown" ||
          message.method === "Runtime.consoleAPICalled"
        ) {
          this.events.push(message);
        }
        return;
      }
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
  let browser;
  let client;

  try {
    await waitFor(baseUrl);
    browser = spawn(
      browserPath,
      [
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        `--remote-debugging-port=${debugPort}`,
        `--user-data-dir=${browserProfile}`,
        "--window-size=1400,1000",
        baseUrl,
      ],
      { stdio: "ignore" },
    );

    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    let page;
    const targetStarted = Date.now();
    while (!page && Date.now() - targetStarted < 10_000) {
      const targets = await (
        await fetch(`http://127.0.0.1:${debugPort}/json/list`)
      ).json();
      page = targets.find(
        (target) => target.type === "page" && target.url.startsWith(baseUrl),
      );
      if (!page) await sleep(80);
    }
    if (!page) throw new Error("No VisualTeX browser target found");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");

    const evaluate = async (expression) => {
      const result = await client.send("Runtime.evaluate", {
        expression,
        awaitPromise: true,
        returnByValue: true,
      });
      if (result.exceptionDetails) {
        const description = result.exceptionDetails.exception?.description;
        throw new Error(
          description || result.exceptionDetails.text || "Runtime.evaluate failed",
        );
      }
      return result.result.value;
    };

    await client.send("Page.navigate", { url: baseUrl });
    await sleep(500);
    await evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(650);
    await evaluate(`new Promise((resolve, reject) => {
      const started = performance.now();
      const check = () => {
        const field = document.querySelector("math-field");
        if (field?.shadowRoot) return resolve(true);
        if (performance.now() - started > 8000) {
          return reject(new Error("Mathfield did not mount"));
        }
        setTimeout(check, 30);
      };
      check();
    })`);

    client.events.length = 0;
    const probe = await evaluate(`(async () => {
      globalThis.__visualtexRuntimeSafetyErrors = [];
      globalThis.__visualtexXssProbe = 0;
      const recordError = (kind, value) => {
        globalThis.__visualtexRuntimeSafetyErrors.push(
          kind + ":" + (value?.message ?? String(value ?? "unknown")),
        );
      };
      window.addEventListener("error", (event) => recordError("error", event.error ?? event.message));
      window.addEventListener("unhandledrejection", (event) => recordError("rejection", event.reason));

      const appField = document.querySelector("math-field");
      const probeHost = document.createElement("div");
      probeHost.dataset.visualtexRuntimeSafetyProbe = "true";
      document.body.append(probeHost);
      const field = new window.MathfieldElement();
      field.value = "";
      probeHost.append(field);
      // Force the transient root shape seen in the reported lifecycle. This
      // invokes the real MathLive setter and would throw at its old mode
      // write, while leaving the rest of the model intact for reparse.
      const modelRoot = field._mathfield.model.root;
      Object.defineProperty(modelRoot, "firstChild", {
        configurable: true,
        get: () => undefined,
      });
      const emptyRootBeforeMacroUpdate = !modelRoot.firstChild;
      const originalMacros = field.macros;
      field.macros = {
        ...originalMacros,
        visualtexruntimeprobe: "\\\\mathrm{VT}",
      };
      const macroCommitted = Boolean(field.macros?.visualtexruntimeprobe);
      const emptyRootAfterMacroUpdate = !modelRoot.firstChild;
      delete modelRoot.firstChild;

      field.setValue("\\\\visualtexruntimeprobe");
      await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      const macroRendered = !field.shadowRoot?.querySelector(".ML__error");
      const macroValue = field.getValue();

      const attack = String.raw\`\\text{<img src=x onerror="globalThis.__visualtexXssProbe=1"> & 'quoted'}\`;
      field.setValue(attack);
      appField.setValue(attack);
      appField.dispatchEvent(new InputEvent("input", {
        bubbles: true,
        composed: true,
        inputType: "insertText",
      }));
      await new Promise((resolve) => setTimeout(resolve, 250));

      const shadowImages =
        (field.shadowRoot?.querySelectorAll("img").length ?? 0) +
        (appField.shadowRoot?.querySelectorAll("img").length ?? 0);
      const documentImages = document.querySelectorAll('img[src="x"]').length;
      const mathMl = field.getValue("math-ml");
      const escapedMathMl =
        !mathMl.includes("<img") &&
        mathMl.includes("&lt;img") &&
        mathMl.includes("&amp;");

      return {
        emptyRootBeforeMacroUpdate,
        emptyRootAfterMacroUpdate,
        macroCommitted,
        macroRendered,
        macroValue,
        shadowImages,
        documentImages,
        xssProbe: globalThis.__visualtexXssProbe,
        escapedMathMl,
        mathMl,
        runtimeErrors: globalThis.__visualtexRuntimeSafetyErrors,
      };
    })()`);
    await sleep(150);
    console.log(JSON.stringify(probe, null, 2));

    assert.equal(
      probe.emptyRootBeforeMacroUpdate,
      true,
      "the regression must exercise an empty mounted MathLive model",
    );
    assert.equal(probe.emptyRootAfterMacroUpdate, true);
    assert.equal(probe.macroCommitted, true);
    assert.equal(probe.macroRendered, true);
    assert.match(probe.macroValue, /visualtexruntimeprobe/);
    assert.equal(probe.shadowImages, 0, "text markup must not create an image node");
    assert.equal(probe.documentImages, 0, "preview markup must not create an image node");
    assert.equal(probe.xssProbe, 0, "text markup must not execute an event handler");
    assert.equal(probe.escapedMathMl, true, `unsafe MathML: ${probe.mathMl}`);
    assert.deepEqual(probe.runtimeErrors, []);

    const exceptions = client.events.filter(
      (event) => event.method === "Runtime.exceptionThrown",
    );
    assert.equal(
      exceptions.length,
      0,
      `unexpected browser exceptions: ${JSON.stringify(exceptions)}`,
    );

    // A persisted pathological formula used to crash the entire application
    // during FormulaField's auxiliary row-height markup pass. Reproduce a real
    // recursive MathLive box overflow and verify that startup remains usable and
    // the original source stays persisted for manual recovery.
    const persistedProbe = await evaluate(`(() => {
      let deepLatex = "x";
      for (let index = 0; index < 900; index += 1) {
        deepLatex = "\\\\begin{array}{c}" + deepLatex + "\\\\end{array}";
      }
      localStorage.setItem("visualtex-editor", JSON.stringify({
        state: {
          title: "Pathological persisted formula",
          lines: [{ id: "deep-array", latex: deepLatex, mode: "display" }],
          activeLineId: "deep-array",
          history: [],
        },
        version: 0,
      }));
      return { length: deepLatex.length };
    })()`);
    client.events.length = 0;
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(1200);
    const startupSafety = await evaluate(`(() => {
      const root = document.getElementById("root");
      const stored = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      const persistedLatex = stored?.state?.lines?.[0]?.latex ?? "";
      return {
        shell: Boolean(document.querySelector(".app-shell")),
        wholeAppFailure: (root?.innerText ?? "").includes("VisualTeX 界面加载失败"),
        isolatedFormulaFailure: Boolean(document.querySelector(".formula-field-render-fallback")),
        mathfieldMounted: Boolean(document.querySelector("math-field")),
        persistedLength: persistedLatex.length,
      };
    })()`);
    console.log(JSON.stringify({ persistedProbe, startupSafety }, null, 2));
    assert.equal(startupSafety.shell, true, "pathological persisted formula must not destroy the app shell");
    assert.equal(startupSafety.wholeAppFailure, false, "pathological persisted formula must not reach the root error boundary");
    assert.equal(
      startupSafety.isolatedFormulaFailure,
      true,
      "structurally unsafe persisted LaTeX must be blocked before entering MathLive",
    );
    assert.equal(
      startupSafety.persistedLength,
      persistedProbe.length,
      "formula isolation must not discard the persisted LaTeX source",
    );

    const customTileProbe = await evaluate(`(() => {
      let deepLatex = "x";
      for (let index = 0; index < 900; index += 1) {
        deepLatex = "\\\\begin{array}{c}" + deepLatex + "\\\\end{array}";
      }
      localStorage.setItem("visualtex-editor", JSON.stringify({
        state: {
          title: "Safe document",
          lines: [{ id: "safe-line", latex: "x+1", mode: "display" }],
          activeLineId: "safe-line",
          history: [],
        },
        version: 0,
      }));
      localStorage.setItem("visualtex-custom-formula-tiles", JSON.stringify({
        version: 3,
        sections: [{ id: "custom", name: "Custom", createdAt: 0 }],
        tiles: [{
          id: "pathological-preview",
          latex: deepLatex,
          sectionId: "custom",
          color: null,
          createdAt: 0,
        }],
      }));
      return { length: deepLatex.length };
    })()`);
    client.events.length = 0;
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(1200);
    const customTileSafety = await evaluate(`(() => {
      const root = document.getElementById("root");
      return {
        shell: Boolean(document.querySelector(".app-shell")),
        wholeAppFailure: (root?.innerText ?? "").includes("VisualTeX 界面加载失败"),
        mathfieldMounted: Boolean(document.querySelector("math-field")),
        customTileStillStored: (localStorage.getItem("visualtex-custom-formula-tiles") ?? "").includes("pathological-preview"),
      };
    })()`);
    console.log(JSON.stringify({ customTileProbe, customTileSafety }, null, 2));
    assert.equal(customTileSafety.shell, true, "a pathological custom-tile preview must not destroy the app shell");
    assert.equal(customTileSafety.wholeAppFailure, false);
    assert.equal(customTileSafety.mathfieldMounted, true);
    assert.equal(customTileSafety.customTileStillStored, true, "preview isolation must not silently delete user toolbar data");

    console.log("VisualTeX MathLive browser runtime safety regression passed");
  } finally {
    client?.close();
    browser?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(250);
    await rm(browserProfile, { recursive: true, force: true }).catch(
      () => undefined,
    );
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.stack : error);
  process.exitCode = 1;
});
