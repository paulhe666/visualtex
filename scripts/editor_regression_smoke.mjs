import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const previewPort = 4173;
const debugPort = 9223;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = `/tmp/visualtex-editor-smoke-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return response;
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
    ["node_modules/vite/bin/vite.js", "preview", "--host", "127.0.0.1", "--port", String(previewPort)],
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
        "--window-size=1400,1000",
        baseUrl,
      ],
      { stdio: "ignore" },
    );

    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    let page;
    const targetStarted = Date.now();
    while (!page && Date.now() - targetStarted < 10000) {
      const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
      page = targets.find(
        (target) => target.type === "page" && target.url.startsWith(baseUrl),
      );
      if (!page) await sleep(100);
    }
    if (!page) throw new Error("No VisualTeX Chrome page target found");

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
        throw new Error(result.exceptionDetails.text || "Runtime.evaluate failed");
      }
      return result.result.value;
    };

    const key = async (value, code, virtualKeyCode) => {
      const common = {
        key: value,
        code,
        windowsVirtualKeyCode: virtualKeyCode,
        nativeVirtualKeyCode: virtualKeyCode,
      };
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        ...common,
        ...(value.length === 1 ? { text: value, unmodifiedText: value } : {}),
      });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(35);
    };

    await evaluate(`localStorage.setItem("visualtex.onboarding.v2.completed", "true")`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(800);
    await evaluate(`new Promise((resolve) => {
      const done = () => document.querySelector("math-field") ? resolve(true) : setTimeout(done, 30);
      done();
    })`);

    const setField = async (latex) => {
      await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.setValue(${JSON.stringify(latex)}, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.position = field.lastOffset;
        field.focus();
        field.dispatchEvent(new InputEvent("input", {
          bubbles: true,
          composed: true,
          inputType: "insertText",
        }));
        return field.value;
      })()`);
      await sleep(180);
    };

    await setField("\\theta");
    const candidateBefore = await evaluate(`Boolean(document.querySelector(".suggestion-popup"))`);
    if (!candidateBefore) throw new Error("Command candidate did not open for \\theta");
    await key("Enter", "Enter", 13);
    await sleep(250);
    const thetaState = await evaluate(`(() => ({
      value: document.querySelector("math-field").value,
      candidateVisible: Boolean(document.querySelector(".suggestion-popup")),
    }))()`);
    if (thetaState.candidateVisible) {
      throw new Error("Command candidate remained open after committing \\theta");
    }

    await setField("");
    await key("\\", "Backslash", 220);
    await key("t", "KeyT", 84);
    await key("h", "KeyH", 72);
    await key("e", "KeyE", 69);
    await key("t", "KeyT", 84);
    await key("a", "KeyA", 65);
    await sleep(220);
    const nativePopover = await evaluate(`(() => {
      const panel = document.getElementById("mathlive-suggestion-popover");
      if (!panel) return null;
      const style = getComputedStyle(panel);
      return {
        visible: panel.classList.contains("is-visible") && style.display !== "none",
        background: style.backgroundColor,
        transition: style.transitionDuration,
        animation: style.animationDuration,
      };
    })()`);
    if (!nativePopover?.visible) {
      throw new Error("MathLive recommendation popover is not visible while typing a command");
    }
    if (nativePopover.background === "rgb(97, 97, 97)") {
      throw new Error("MathLive recommendation popover still uses the old black/gray background");
    }

    const nativeBeforeArrow = await evaluate(`(() => {
      const panel = document.getElementById("mathlive-suggestion-popover");
      const selected = panel.querySelector("li.ML__popover__current");
      window.__visualtexStableNativePanel = panel;
      const style = getComputedStyle(selected);
      return {
        command: selected?.dataset.command,
        background: style.backgroundColor,
        border: style.borderColor,
        color: style.color,
      };
    })()`);
    if (nativeBeforeArrow.background === "rgb(31, 99, 142)") {
      throw new Error("Selected native recommendation still uses a solid dark-blue fill");
    }
    if (nativeBeforeArrow.border === "rgba(0, 0, 0, 0)") {
      throw new Error("Selected native recommendation has no visible selection border");
    }

    await key("ArrowDown", "ArrowDown", 40);
    await sleep(120);
    const nativeAfterArrow = await evaluate(`(() => {
      const panel = document.getElementById("mathlive-suggestion-popover");
      const selected = panel.querySelector("li.ML__popover__current");
      const style = getComputedStyle(selected);
      return {
        samePanelNode: panel === window.__visualtexStableNativePanel,
        command: selected?.dataset.command,
        background: style.backgroundColor,
        border: style.borderColor,
        color: style.color,
      };
    })()`);
    if (!nativeAfterArrow.samePanelNode) {
      throw new Error("Arrow navigation replaced the native recommendation panel and can flicker");
    }
    if (nativeAfterArrow.command !== "\\thetasym") {
      throw new Error(`ArrowDown did not move the native recommendation selection: ${JSON.stringify(nativeAfterArrow)}`);
    }
    if (nativeAfterArrow.background === "rgb(31, 99, 142)") {
      throw new Error("Moved native recommendation still uses a solid dark-blue fill");
    }

    await key("Enter", "Enter", 13);
    await sleep(250);
    const nativeCommitState = await evaluate(`(() => ({
      value: document.querySelector("math-field").value,
      lineCount: document.querySelectorAll(".formula-line").length,
      nativeVisible: document.getElementById("mathlive-suggestion-popover")?.classList.contains("is-visible") ?? false,
      candidateVisible: Boolean(document.querySelector(".suggestion-popup")),
    }))()`);
    if (nativeCommitState.lineCount !== 1 || !nativeCommitState.value.endsWith("thetasym")) {
      throw new Error(`Enter did not commit the selected native MathLive recommendation: ${JSON.stringify(nativeCommitState)}`);
    }
    if (nativeCommitState.nativeVisible || nativeCommitState.candidateVisible) {
      throw new Error(`Recommendation remained visible after commit: ${JSON.stringify(nativeCommitState)}`);
    }

    await setField("\\alpha");
    const simpleMetrics = await evaluate(`(() => {
      const line = document.querySelector(".formula-line");
      const field = document.querySelector("math-field");
      const content = field.shadowRoot.querySelector('[part="content"]');
      const rects = [...content.querySelectorAll("[data-atom-id]")]
        .map((atom) => atom.getBoundingClientRect())
        .filter((rect) => rect.height > 0);
      const top = Math.min(...rects.map((rect) => rect.top));
      const bottom = Math.max(...rects.map((rect) => rect.bottom));
      const lineRect = line.getBoundingClientRect();
      return {
        lineHeight: lineRect.height,
        fieldHeight: field.getBoundingClientRect().height,
        centerDelta: Math.abs((top + bottom) / 2 - (lineRect.top + lineRect.bottom) / 2),
      };
    })()`);

    await setField("\\frac{a}{b}");
    const tallMetrics = await evaluate(`(() => ({
      lineHeight: document.querySelector(".formula-line").getBoundingClientRect().height,
      fieldHeight: document.querySelector("math-field").getBoundingClientRect().height,
    }))()`);
    if (!(simpleMetrics.lineHeight < tallMetrics.lineHeight)) {
      throw new Error(`Simple formula row did not shrink (${simpleMetrics.lineHeight} vs ${tallMetrics.lineHeight})`);
    }
    if (simpleMetrics.centerDelta > 10) {
      throw new Error(`Simple formula is not vertically centered (delta ${simpleMetrics.centerDelta})`);
    }

    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.setValue("a\\\\int_{x}^{y}", {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
      field.position = field.lastOffset;
      field.focus();
      field.dispatchEvent(new InputEvent("input", { bubbles: true, composed: true }));
    })()`);
    await sleep(150);
    const deletionStates = [];
    for (let index = 0; index < 12; index += 1) {
      await key("Backspace", "Backspace", 8);
      deletionStates.push(
        await evaluate(`(() => {
          const field = document.querySelector("math-field");
          return { value: field.value, position: field.position };
        })()`),
      );
    }
    const skippedOperator = deletionStates.some(
      (state) => !state.value.includes("a") && /\\\\int/.test(state.value),
    );
    if (skippedOperator) {
      throw new Error(`Backspace skipped the integral and deleted preceding content: ${JSON.stringify(deletionStates)}`);
    }

    console.log(JSON.stringify({
      thetaState,
      nativePopover,
      nativeCommitState,
      simpleMetrics,
      tallMetrics,
      deletionStates,
    }, null, 2));
    console.log("Editor regression smoke test passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(300);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => undefined);
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.stack : error);
  process.exitCode = 1;
});
