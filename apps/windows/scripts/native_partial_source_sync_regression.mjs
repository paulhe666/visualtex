import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const portOffset = process.pid % 700;
const previewPort = 14900 + portOffset;
const debugPort = 16900 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const profile = createBrowserProfilePath("visualtex-native-partial-source");
const chromePath = resolveChromiumExecutable();
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {}
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
        `--user-data-dir=${profile}`,
        "--window-size=1500,1000",
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
    await sleep(700);

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
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.onboarding.windows.desktop.v1.1.0.completed", "true");
      localStorage.setItem("visualtex.release-welcome.1.2.6.seen", "true");
      localStorage.setItem("visualtex-desktop-editor-source-open", "true");
      let persisted;
      try { persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "null"); } catch { persisted = null; }
      if (!persisted || typeof persisted !== "object") persisted = { state: {}, version: 0 };
      persisted.state = {
        ...(persisted.state || {}),
        lines: [{ id: "partial-source-line", latex: "", mode: "display" }],
        activeLineId: "partial-source-line",
        sourceOpen: true,
        editorLayout: "standard",
        latexCodeFormat: "raw",
        checkUpdatesOnStartup: false,
        language: "cn",
        inputBehavior: {
          ...(persisted.state?.inputBehavior || {}),
          showStructuredCommandSuggestions: true,
          showOtherCommandSuggestions: true,
        },
      };
      delete persisted.state.latex;
      localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(850);
    await evaluate(`new Promise((resolve, reject) => {
      const started = performance.now();
      const poll = () => {
        if (document.querySelector("math-field") && document.querySelector(".source-panel .cm-content")) return resolve(true);
        if (performance.now() - started > 5000) return reject(new Error("Editor/source did not mount"));
        setTimeout(poll, 30);
      };
      poll();
    })`);
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.focus();
      field.position = field.lastOffset;
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
    })()`);
    await sleep(150);

    const typeKey = async (key, code, virtualKey, pause = 80) => {
      const common = {
        key,
        code,
        windowsVirtualKeyCode: virtualKey,
        nativeVirtualKeyCode: virtualKey,
      };
      const text = key.length === 1 ? key : "";
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        ...common,
        text,
        unmodifiedText: text,
      });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(pause);
    };

    // WebView2 SendInput reports an empty KeyboardEvent.code for the real
    // backslash key on this Windows host.
    await typeKey("\\", "", 220, 140);
    for (const letter of "par") {
      await typeKey(
        letter,
        `Key${letter.toUpperCase()}`,
        letter.toUpperCase().charCodeAt(0),
        90,
      );
    }
    await sleep(180);

    const beforeSpace = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      const stable = document.querySelector('#visualtex-native-input-suggestion-popover');
      const source = document.querySelector('#mathlive-suggestion-popover');
      const readItems = (panel) => [...(panel?.querySelectorAll('li[data-command]') ?? [])].map((item) => ({
        command: item.dataset.command ?? "",
        current: item.classList.contains('ML__popover__current'),
        label: item.querySelector('.ML__popover__latex')?.textContent ?? "",
      }));
      return {
        value: field.value,
        mode: field.mode,
        raw: [...(field.shadowRoot?.querySelectorAll('.ML__raw-latex') ?? [])]
          .filter((node) => !node.classList.contains('ML__suggestion'))
          .map((node) => node.textContent ?? "").join(""),
        pendingNativeSuggestion: field.dataset.pendingNativeSuggestion ?? "",
        stableVisible: stable?.classList.contains('is-visible') ?? false,
        stableItems: readItems(stable),
        sourceItems: readItems(source),
        sourceLines: [...document.querySelectorAll('.source-panel .cm-line')].map((line) => line.textContent ?? ""),
      };
    })()`);
    console.log("partial-before-space", JSON.stringify(beforeSpace));
    assert.equal(beforeSpace.raw.replace(/\s+/g, ""), "\\par");
    assert.ok(
      beforeSpace.stableItems.some((item) => item.command === "\\partial"),
      `\\partial is missing from the native candidate pool: ${JSON.stringify(beforeSpace)}`,
    );
    assert.ok(
      !beforeSpace.stableItems.some((item) => item.command === "\\part"),
      `document-only \\part leaked into the math candidate pool: ${JSON.stringify(beforeSpace)}`,
    );

    let selectedCommand = beforeSpace.stableItems.find((item) => item.current)?.command ?? "";
    for (let attempt = 0; selectedCommand !== "\\partial" && attempt < beforeSpace.stableItems.length; attempt += 1) {
      await typeKey("ArrowDown", "ArrowDown", 40, 100);
      selectedCommand = await evaluate(`document.querySelector('#visualtex-native-input-suggestion-popover li.ML__popover__current')?.dataset.command ?? ""`);
    }
    assert.equal(selectedCommand, "\\partial", `Keyboard candidate navigation did not select \\partial: ${selectedCommand}`);
    await typeKey(" ", "Space", 32, 160);
    await sleep(260);

    const afterSpace = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      return {
        value: field.value,
        mode: field.mode,
        raw: [...(field.shadowRoot?.querySelectorAll('.ML__raw-latex') ?? [])]
          .filter((node) => !node.classList.contains('ML__suggestion'))
          .map((node) => node.textContent ?? "").join(""),
        pendingNativeSuggestion: field.dataset.pendingNativeSuggestion ?? "",
        visualTexCandidateVisible: Boolean(document.querySelector(".suggestion-popup")),
        sourceLines: [...document.querySelectorAll('.source-panel .cm-line')].map((line) => line.textContent ?? ""),
        persistedLines: persisted?.state?.lines ?? [],
      };
    })()`);
    console.log("partial-after-space", JSON.stringify(afterSpace));
    assert.equal(afterSpace.raw, "", "accepted partial remained in raw-LaTeX mode");
    assert.match(afterSpace.value, /\\partial/, `visual field did not commit \\partial: ${JSON.stringify(afterSpace)}`);
    assert.match(afterSpace.sourceLines[0] ?? "", /\\partial/, `source pane did not commit \\partial: ${JSON.stringify(afterSpace)}`);
    assert.doesNotMatch(afterSpace.sourceLines[0] ?? "", /^\\par$/);
    assert.equal(
      afterSpace.visualTexCandidateVisible,
      false,
      `VisualTeX opened a second candidate panel after committing \\partial: ${JSON.stringify(afterSpace)}`,
    );

    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.setValue("", { silenceNotifications: false });
      field.focus();
      field.position = field.lastOffset;
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
    })()`);
    await sleep(160);
    await typeKey("\\", "", 220, 140);
    for (const letter of "dag") {
      await typeKey(
        letter,
        `Key${letter.toUpperCase()}`,
        letter.toUpperCase().charCodeAt(0),
        90,
      );
    }
    const daggerCandidates = await evaluate(`(() => {
      const source = document.querySelector('math-field')?.shadowRoot?.querySelector('.ML__popover');
      const stable = document.querySelector('#visualtex-native-input-suggestion-popover');
      const commands = (panel) => [...(panel?.querySelectorAll('li[data-command]') ?? [])]
        .map((item) => item.dataset.command ?? "");
      return { source: commands(source), stable: commands(stable) };
    })()`);
    console.log("dagger-candidates", JSON.stringify(daggerCandidates));
    assert.ok(
      !daggerCandidates.stable.includes("\\dag"),
      `deprecated \\dag leaked into the math candidate pool: ${JSON.stringify(daggerCandidates)}`,
    );
    assert.ok(
      daggerCandidates.stable.includes("\\dagger"),
      `\\dagger is missing after filtering \\dag: ${JSON.stringify(daggerCandidates)}`,
    );

    console.log("Native candidate filtering and partial source synchronization regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(250);
    await rm(profile, {
      recursive: true,
      force: true,
      maxRetries: 6,
      retryDelay: 100,
    }).catch((error) => {
      if (process.platform !== "win32" || error?.code !== "EBUSY") throw error;
    });
  }
}

await main();
