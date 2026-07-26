import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 900;
const previewPort = 7800 + offset;
const debugPort = 13800 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}/editor`;
const chromeProfile = `/tmp/visualtex-formula-alignment-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while the process starts.
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
          title: "Alignment regression",
          lines: [
            { id: "alignment-line-1", latex: "a+b" },
            { id: "alignment-line-2", latex: "\\\\frac{c}{d}" },
          ],
          activeLineId: "alignment-line-1",
          formulaAlignment: "left",
          language: "cn",
          zoom: 1,
        },
        version: 0,
      }));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(650);
    await evaluate(`new Promise((resolve) => {
      const poll = () => document.querySelectorAll("math-field").length === 2 &&
        document.querySelector(".formula-alignment-controls")
        ? resolve(true)
        : setTimeout(poll, 25);
      poll();
    })`);

    const layout = await evaluate(`(() => {
      const title = document.querySelector(".pane-title-copy h1");
      const controls = document.querySelector(".formula-alignment-controls");
      const toolbar = document.querySelector(".formula-toolbar");
      const header = document.querySelector(".formula-toolbar-header");
      const titleRect = title?.getBoundingClientRect();
      const controlsRect = controls?.getBoundingClientRect();
      return {
        title: title?.textContent ?? "",
        titleRight: titleRect?.right ?? -1,
        controlsLeft: controlsRect?.left ?? -1,
        sameVerticalBand: titleRect && controlsRect
          ? Math.abs((titleRect.top + titleRect.bottom) / 2 - (controlsRect.top + controlsRect.bottom) / 2)
          : 999,
        buttons: Array.from(controls?.querySelectorAll("button") ?? []).map(
          (button) => button.dataset.formulaAlignment,
        ),
        firstSidebarChild: toolbar?.firstElementChild?.className ?? "",
        sidebarHeaderTop: header?.getBoundingClientRect().top ?? -1,
        sidebarTop: toolbar?.getBoundingClientRect().top ?? -1,
        hasOldQuickRow: Boolean(document.querySelector(".formula-toolbar-quick-actions")),
      };
    })()`);
    assert.equal(layout.title, "可视化编辑", JSON.stringify(layout));
    assert.deepEqual(layout.buttons, ["left", "center", "right"]);
    assert.ok(layout.controlsLeft >= layout.titleRight, JSON.stringify(layout));
    assert.ok(layout.sameVerticalBand <= 1, JSON.stringify(layout));
    assert.equal(layout.firstSidebarChild, "formula-toolbar-header", JSON.stringify(layout));
    assert.ok(Math.abs(layout.sidebarHeaderTop - layout.sidebarTop) <= 0.5, JSON.stringify(layout));
    assert.equal(layout.hasOldQuickRow, false, JSON.stringify(layout));

    const readState = () => evaluate(`(() => {
      const root = document.querySelector(".multi-line-editor");
      const lines = Array.from(document.querySelectorAll(".formula-line"));
      const buttons = Array.from(document.querySelectorAll(".formula-alignment-button"));
      const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}").state ?? {};
      return {
        rootAlignment: root?.dataset.formulaAlignment ?? "",
        activeLineId: document.querySelector(".formula-line.is-active")?.dataset.lineId ?? "",
        pressed: buttons.find((button) => button.getAttribute("aria-pressed") === "true")?.dataset.formulaAlignment ?? "",
        lines: lines.map((line) => {
          const host = line.querySelector(".mathfield-host");
          const field = line.querySelector("math-field");
          const hostRect = host?.getBoundingClientRect();
          const fieldRect = field?.getBoundingClientRect();
          return {
            id: line.dataset.lineId ?? "",
            latex: field?.value ?? "",
            hostWidth: hostRect?.width ?? 0,
            fieldWidth: fieldRect?.width ?? 0,
            hostJustify: host ? getComputedStyle(host).justifyContent : "",
            hostFieldLeftGap:
              hostRect && fieldRect ? fieldRect.left - hostRect.left : -1,
            hostFieldRightGap:
              hostRect && fieldRect ? hostRect.right - fieldRect.right : -1,
            hasLineAlignment: line.hasAttribute("data-alignment"),
          };
        }),
        persistedAlignment: persisted.formulaAlignment ?? "",
        persistedLines: persisted.lines ?? [],
      };
    })()`);

    const originalLatex = ["a+b", "\\frac{c}{d}"];
    let state = await readState();
    assert.equal(state.rootAlignment, "left", JSON.stringify(state));
    assert.equal(state.pressed, "left", JSON.stringify(state));
    assert.deepEqual(state.lines.map((line) => line.latex), originalLatex);
    assert.ok(state.lines.every((line) => line.hostJustify === "flex-start"), JSON.stringify(state));
    assert.ok(state.lines.every((line) => line.fieldWidth < line.hostWidth - 20), JSON.stringify(state));
    assert.ok(
      state.lines.every(
        (line) =>
          line.hostFieldLeftGap >= -0.5 &&
          line.hostFieldRightGap - line.hostFieldLeftGap > 20,
      ),
      JSON.stringify(state),
    );
    assert.ok(state.lines.every((line) => !line.hasLineAlignment), JSON.stringify(state));

    const clickAlignment = async (alignment) => {
      await evaluate(`document.querySelector('[data-formula-alignment="${alignment}"]')?.click()`);
      await sleep(80);
      return readState();
    };

    const clickBlankArea = async (lineId, side) =>
      evaluate(`(() => {
        const host = document.querySelector('[data-line-id="${lineId}"] .mathfield-host');
        const field = host?.querySelector('math-field');
        if (!host || !field) return { position: -1, lastOffset: -1 };
        const rect = host.getBoundingClientRect();
        const clientX = ${JSON.stringify(side)} === 'left'
          ? rect.left + 4
          : rect.right - 4;
        const clientY = (rect.top + rect.bottom) / 2;
        host.dispatchEvent(new PointerEvent('pointerdown', {
          bubbles: true,
          composed: true,
          cancelable: true,
          button: 0,
          buttons: 1,
          pointerId: 41,
          isPrimary: true,
          clientX,
          clientY,
        }));
        window.dispatchEvent(new PointerEvent('pointerup', {
          bubbles: true,
          composed: true,
          button: 0,
          pointerId: 41,
          isPrimary: true,
          clientX,
          clientY,
        }));
        return { position: field.position, lastOffset: field.lastOffset };
      })()`);

    state = await clickAlignment("center");
    assert.equal(state.rootAlignment, "center", JSON.stringify(state));
    assert.equal(state.pressed, "center", JSON.stringify(state));
    assert.ok(state.lines.every((line) => line.hostJustify === "center"), JSON.stringify(state));
    assert.ok(
      state.lines.every(
        (line) => Math.abs(line.hostFieldLeftGap - line.hostFieldRightGap) <= 1,
      ),
      JSON.stringify(state),
    );
    assert.deepEqual(state.lines.map((line) => line.latex), originalLatex);
    const centeredLeftBlank = await clickBlankArea("alignment-line-1", "left");
    assert.equal(centeredLeftBlank.position, 0, JSON.stringify(centeredLeftBlank));
    const centeredRightBlank = await clickBlankArea("alignment-line-1", "right");
    assert.equal(
      centeredRightBlank.position,
      centeredRightBlank.lastOffset,
      JSON.stringify(centeredRightBlank),
    );
    assert.equal(state.persistedAlignment, "center", JSON.stringify(state));
    assert.ok(
      state.persistedLines.every((line) => !("alignment" in line)),
      JSON.stringify(state),
    );

    await evaluate(`(() => {
      const field = document.querySelector('[data-line-id="alignment-line-2"] math-field');
      field?.focus();
      field?.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
    })()`);
    await sleep(80);
    state = await readState();
    assert.equal(state.activeLineId, "alignment-line-2", JSON.stringify(state));
    assert.equal(state.rootAlignment, "center", JSON.stringify(state));
    assert.equal(state.pressed, "center", JSON.stringify(state));

    state = await clickAlignment("right");
    assert.equal(state.rootAlignment, "right", JSON.stringify(state));
    assert.equal(state.pressed, "right", JSON.stringify(state));
    assert.ok(state.lines.every((line) => line.hostJustify === "flex-end"), JSON.stringify(state));
    assert.ok(
      state.lines.every(
        (line) =>
          line.hostFieldRightGap >= -0.5 &&
          line.hostFieldLeftGap - line.hostFieldRightGap > 20,
      ),
      JSON.stringify(state),
    );
    assert.deepEqual(state.lines.map((line) => line.latex), originalLatex);
    const rightAlignedLeftBlank = await clickBlankArea("alignment-line-1", "left");
    assert.equal(rightAlignedLeftBlank.position, 0, JSON.stringify(rightAlignedLeftBlank));
    const rightAlignedRightBlank = await clickBlankArea("alignment-line-1", "right");
    assert.equal(
      rightAlignedRightBlank.position,
      rightAlignedRightBlank.lastOffset,
      JSON.stringify(rightAlignedRightBlank),
    );
    assert.equal(state.persistedAlignment, "right", JSON.stringify(state));

    await client.send("Page.reload", { ignoreCache: true });
    await sleep(650);
    state = await readState();
    assert.equal(state.rootAlignment, "right", JSON.stringify(state));
    assert.equal(state.pressed, "right", JSON.stringify(state));
    assert.deepEqual(state.lines.map((line) => line.latex), originalLatex);
    assert.ok(state.lines.every((line) => line.hostJustify === "flex-end"), JSON.stringify(state));
    assert.ok(
      state.lines.every(
        (line) =>
          line.hostFieldRightGap >= -0.5 &&
          line.hostFieldLeftGap - line.hostFieldRightGap > 20,
      ),
      JSON.stringify(state),
    );

    const pressEnter = async () => {
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        key: "Enter",
        code: "Enter",
        windowsVirtualKeyCode: 13,
        nativeVirtualKeyCode: 13,
      });
      await client.send("Input.dispatchKeyEvent", {
        type: "keyUp",
        key: "Enter",
        code: "Enter",
        windowsVirtualKeyCode: 13,
        nativeVirtualKeyCode: 13,
      });
      await sleep(100);
    };

    await evaluate(`(() => {
      const field = document.querySelector('[data-line-id="alignment-line-1"] math-field');
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      field.position = field.lastOffset;
    })()`);
    await pressEnter();
    state = await readState();
    assert.equal(state.lines.length, 3, JSON.stringify(state));
    assert.equal(state.rootAlignment, "right", JSON.stringify(state));
    assert.ok(state.lines.every((line) => line.hostJustify === "flex-end"), JSON.stringify(state));
    assert.ok(
      state.lines.every(
        (line) =>
          line.hostFieldRightGap >= -0.5 &&
          line.hostFieldLeftGap - line.hostFieldRightGap > 20,
      ),
      JSON.stringify(state),
    );
    assert.ok(state.persistedLines.every((line) => !("alignment" in line)), JSON.stringify(state));

    await evaluate(`(() => {
      const field = document.querySelector('[data-line-id="alignment-line-1"] math-field');
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      field.position = Math.min(1, field.lastOffset);
    })()`);
    await pressEnter();
    state = await readState();
    assert.equal(state.lines.length, 4, JSON.stringify(state));
    assert.equal(state.rootAlignment, "right", JSON.stringify(state));
    assert.ok(state.lines.every((line) => line.hostJustify === "flex-end"), JSON.stringify(state));
    assert.ok(
      state.lines.every(
        (line) =>
          line.hostFieldRightGap >= -0.5 &&
          line.hostFieldLeftGap - line.hostFieldRightGap > 20,
      ),
      JSON.stringify(state),
    );
    assert.ok(state.persistedLines.every((line) => !("alignment" in line)), JSON.stringify(state));

    console.log("Document-level visual formula alignment regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(250);
    await rm(chromeProfile, { recursive: true, force: true });
  }
}

await main();
