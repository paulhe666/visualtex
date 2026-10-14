import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const portOffset = process.pid % 700;
const previewPort = 8500 + portOffset;
const debugPort = 13500 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = `/tmp/visualtex-keypad-mode-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while local processes start.
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
        "--window-size=619,368",
        baseUrl,
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
    const page = targets.find(
      (target) => target.type === "page" && target.url.startsWith(baseUrl),
    );
    if (!page) throw new Error("No VisualTeX Chrome page target found");
    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
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

    const clickSelector = async (selector) => {
      const rect = await evaluate(`(() => {
        const element = document.querySelector(${JSON.stringify(selector)});
        if (!element) return null;
        const box = element.getBoundingClientRect();
        return { x: box.left + box.width / 2, y: box.top + box.height / 2 };
      })()`);
      if (!rect) throw new Error(`Unable to click missing selector: ${selector}`);
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: rect.x,
        y: rect.y,
        button: "left",
        clickCount: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: rect.x,
        y: rect.y,
        button: "left",
        clickCount: 1,
      });
      await sleep(40);
    };

    await evaluate(`(() => {
      localStorage.setItem('visualtex.onboarding.v3.completed', 'true');
      localStorage.setItem('visualtex.onboarding.macos.desktop.v1.2.0.completed', 'true');
      localStorage.setItem('visualtex.office.macos.native-first-run.v1.2.0.completed', 'true');
      return true;
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(350);
    await evaluate(`new Promise((resolve) => {
      const done = () => document.querySelector("math-field") ? resolve(true) : setTimeout(done, 30);
      done();
    })`);
    await evaluate(`(() => {
      const laterButton = [...document.querySelectorAll('.office-first-run-backdrop button')]
        .find((button) => /Later|稍后处理/.test(button.textContent || ''));
      if (laterButton instanceof HTMLElement) laterButton.click();
      return true;
    })()`);
    await sleep(80);

    const normalProbe = await evaluate(`(() => {
      const header = document.querySelector('.editor-pane-header');
      return {
        width: innerWidth,
        height: innerHeight,
        appHeader: Boolean(document.querySelector('.app-header')),
        editorHeaderFormat: Boolean(document.querySelector('.editor-pane-header .editor-code-format-control')),
        topHeaderFormat: Boolean(document.querySelector('.app-header .code-format-control')),
        editorHeaderKeypadButton: Boolean(document.querySelector('.editor-pane-header [data-keypad-mode-toggle]')),
        topHeaderKeypadButton: Boolean(document.querySelector('.app-header [data-keypad-mode-toggle]')),
        headerClientWidth: header?.clientWidth ?? 0,
        headerScrollWidth: header?.scrollWidth ?? 0,
      };
    })()`);
    assert.ok(normalProbe.width <= 640 && normalProbe.height <= 400, `Unexpected keypad regression viewport: ${JSON.stringify(normalProbe)}`);
    assert.equal(normalProbe.appHeader, true, "normal mode lost the main application header");
    assert.equal(normalProbe.editorHeaderFormat, false, "normal mode must keep the LaTeX format control out of the editor header");
    assert.equal(normalProbe.topHeaderFormat, true, "normal mode did not restore the LaTeX format control to the main application header");
    assert.equal(normalProbe.editorHeaderKeypadButton, false, "normal mode duplicated the keypad button in the editor header");
    assert.equal(normalProbe.topHeaderKeypadButton, true, "keypad mode button is missing from the main application header");
    assert.ok(
      normalProbe.headerScrollWidth <= normalProbe.headerClientWidth + 2,
      `normal editor header overflows at the current compact window size: ${JSON.stringify(normalProbe)}`,
    );

    const assertSharedFloatingLayer = async (
      triggerSelector,
      layerSelector,
      label,
    ) => {
      await clickSelector(triggerSelector);
      await sleep(220);
      const state = await evaluate(`(() => {
        const layer = document.querySelector(${JSON.stringify(layerSelector)});
        if (!(layer instanceof HTMLElement)) return null;
        const rect = layer.getBoundingClientRect();
        return {
          left: rect.left,
          top: rect.top,
          right: rect.right,
          bottom: rect.bottom,
          viewportWidth: innerWidth,
          viewportHeight: innerHeight,
          autoAvoided:
            layer.dataset.visualtexAutoAvoidAdjusted === 'true',
        };
      })()`);
      assert.ok(state, `${label} did not open`);
      assert.equal(
        state.autoAvoided,
        true,
        `${label} did not use the shared floating-layer auto-avoidance manager`,
      );
      assert.ok(
        state.left >= 8 &&
          state.top >= 8 &&
          state.right <= state.viewportWidth - 8 &&
          state.bottom <= state.viewportHeight - 8,
        `${label} escaped the viewport after shared auto-avoidance: ${JSON.stringify(state)}`,
      );
      await evaluate(`document.querySelector(${JSON.stringify(triggerSelector)})?.click()`);
      await sleep(80);
      assert.equal(
        await evaluate(`document.querySelector(${JSON.stringify(layerSelector)}) === null`),
        true,
        `${label} did not close cleanly after the bounds check`,
      );
    };

    await assertSharedFloatingLayer(
      '.menu-button',
      '.app-menu-popover',
      'main application menu',
    );

    await evaluate(`(() => {
      const field = document.querySelector('math-field');
      if (!field) return false;
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      field.position = field.lastOffset;
      return true;
    })()`);
    await client.send("Input.insertText", { text: "abc" });
    await sleep(120);
    assert.match(
      await evaluate(`document.querySelector('math-field')?.value ?? ''`),
      /abc/,
      "unable to prepare a selected formula for color popover bounds regression",
    );

    const colorPopoverBounds = async (selector, kind) => {
      await evaluate(`(() => {
        const field = document.querySelector('math-field');
        if (!field) return false;
        field.focus();
        field.selection = {
          ranges: [[0, field.lastOffset]],
          direction: 'forward',
        };
        return true;
      })()`);
      await sleep(30);
      await clickSelector(selector);
      await sleep(180);
      const bounds = await evaluate(`(() => {
        const popover = document.querySelector('[data-formula-color-popover="${kind}"]');
        const editorPane = popover?.closest('.formula-workspace.editor-pane');
        if (!(popover instanceof HTMLElement)) return null;
        const rect = popover.getBoundingClientRect();
        const editorRect = editorPane instanceof HTMLElement
          ? editorPane.getBoundingClientRect()
          : null;
        return {
          left: rect.left,
          right: rect.right,
          viewportWidth: innerWidth,
          editorLeft: editorRect?.left ?? 0,
          editorRight: editorRect?.right ?? innerWidth,
          editorWidth: editorRect?.width ?? 0,
          autoAvoided:
            popover.dataset.visualtexAutoAvoidAdjusted === 'true',
        };
      })()`);
      assert.ok(bounds, `${kind} color popover did not open`);
      assert.equal(
        bounds.autoAvoided,
        true,
        `${kind} color popover was not processed by the shared floating-layer auto-avoidance manager`,
      );
      const leftBoundary = bounds.editorWidth > 0
        ? Math.max(8, bounds.editorLeft + 8)
        : 8;
      const rightBoundary = bounds.editorWidth > 0
        ? Math.min(bounds.viewportWidth - 8, bounds.editorRight - 8)
        : bounds.viewportWidth - 8;
      assert.ok(
        bounds.left >= leftBoundary && bounds.right <= rightBoundary,
        `${kind} color popover escaped the visible application/editor bounds: ${JSON.stringify(bounds)}`,
      );
      await clickSelector(selector);
      await sleep(40);
      return bounds;
    };

    await colorPopoverBounds('[data-formula-selection-color]', 'color');
    await colorPopoverBounds(
      '[data-formula-selection-background]',
      'backgroundColor',
    );

    // Browser regression covers keypad UI only. Native window-size memory and
    // switching now live entirely in the Rust/Tauri backend and are verified
    // against the real macOS window instead of localStorage/browser bounds.
    await evaluate(`document.querySelector('[data-keypad-mode-toggle]')?.click()`);
    await sleep(120);
    const keypadProbe = await evaluate(`(() => {
      const header = document.querySelector('.editor-pane-header');
      const keypadButton = document.querySelector('.editor-pane-header [data-keypad-mode-toggle]');
      const headerRect = header?.getBoundingClientRect();
      const keypadButtonRect = keypadButton?.getBoundingClientRect();
      return ({
      shell: document.querySelector('.app-shell')?.classList.contains('is-keypad-mode') ?? false,
      workspace: document.querySelector('.workspace')?.classList.contains('is-keypad-mode') ?? false,
      appHeader: Boolean(document.querySelector('.app-header')),
      alignment: Boolean(document.querySelector('.editor-pane-header .formula-alignment-controls')),
      format: Boolean(document.querySelector('.editor-pane-header .editor-code-format-control')),
      keypadButton: Boolean(document.querySelector('.editor-pane-header [data-keypad-mode-toggle]')),
      formulaToolbar: Boolean(document.querySelector('.formula-toolbar')),
      classicDock: Boolean(document.querySelector('.classic-bottom-dock')),
      sourcePane: Boolean(document.querySelector('.source-pane-slot')),
      sourceToggle: Boolean(document.querySelector('.source-toggle-row')),
      canvasTools: Boolean(document.querySelector('.canvas-tool-group')),
      exportAction: Boolean(document.querySelector('.workspace-export-trigger')),
      inputBehavior: Boolean(document.querySelector('.canvas-input-behavior-trigger')),
      quickOcr: Boolean(document.querySelector('[data-quick-ocr-button]')),
      quickOcrMode: Boolean(document.querySelector('[data-quick-ocr-mode-trigger]')),
      silentOcr: Boolean(document.querySelector('[data-silent-ocr-toggle]')),
      ocrModel: Boolean(document.querySelector('.canvas-ocr-model')),
      canvasZoom: Boolean(document.querySelector('.canvas-controls')), 
      editor: Boolean(document.querySelector('.keypad-editor-pane-body math-field')),
      headerVisible: Boolean(headerRect && headerRect.width > 0 && headerRect.height > 0),
      keypadButtonVisible: Boolean(keypadButtonRect && keypadButtonRect.width > 0 && keypadButtonRect.height > 0),
      headerClientWidth: header?.clientWidth ?? 0,
      headerScrollWidth: header?.scrollWidth ?? 0,
      alignmentWidth: document.querySelector('.formula-alignment-controls')?.getBoundingClientRect().width ?? 0,
      desktopControlsWidth: document.querySelector('.desktop-editor-header-controls')?.getBoundingClientRect().width ?? 0,
      canvasToolsWidth: document.querySelector('.canvas-tool-group')?.getBoundingClientRect().width ?? 0,
      paneTitleWidth: document.querySelector('.pane-title-group')?.getBoundingClientRect().width ?? 0,
      });
    })()`);
    assert.equal(keypadProbe.shell, true, "app shell did not enter keypad mode");
    assert.equal(keypadProbe.workspace, true, "workspace did not enter keypad mode");
    assert.equal(keypadProbe.appHeader, false, "main application header is still rendered in keypad mode");
    assert.equal(keypadProbe.alignment, true, "formula alignment controls disappeared in keypad mode");
    assert.equal(keypadProbe.format, true, "LaTeX format selector disappeared in keypad mode");
    assert.equal(keypadProbe.keypadButton, true, "keypad exit button disappeared in keypad mode");
    assert.equal(keypadProbe.headerVisible, true, "keypad editor header is not visible");
    assert.equal(keypadProbe.keypadButtonVisible, true, "keypad exit button is not visible");
    assert.equal(keypadProbe.editor, true, "visual formula editor disappeared in keypad mode");
    assert.ok(
      keypadProbe.headerScrollWidth <= keypadProbe.headerClientWidth + 2,
      `keypad editor header overflows at 619px: ${JSON.stringify(keypadProbe)}`,
    );
    for (const [name, value] of Object.entries({
      formulaToolbar: keypadProbe.formulaToolbar,
      classicDock: keypadProbe.classicDock,
      sourcePane: keypadProbe.sourcePane,
      sourceToggle: keypadProbe.sourceToggle,
      ocrModel: keypadProbe.ocrModel,
      canvasZoom: keypadProbe.canvasZoom,
    })) {
      assert.equal(value, false, `${name} must not be rendered in keypad mode`);
    }
    for (const [name, value] of Object.entries({
      canvasTools: keypadProbe.canvasTools,
      exportAction: keypadProbe.exportAction,
      inputBehavior: keypadProbe.inputBehavior,
      quickOcr: keypadProbe.quickOcr,
      quickOcrMode: keypadProbe.quickOcrMode,
      silentOcr: keypadProbe.silentOcr,
    })) {
      assert.equal(value, true, `${name} must remain available in keypad mode`);
    }

    const codeFormatTriggerHit = await evaluate(`(() => {
      const button = document.querySelector('.editor-code-format-control .code-format-primary');
      if (!(button instanceof HTMLElement)) return { hit: null, reason: 'missing-trigger' };
      const box = button.getBoundingClientRect();
      const hit = document.elementFromPoint(box.left + box.width / 2, box.top + box.height / 2);
      return {
        hit: hit?.closest('.code-format-primary') === button,
        rect: { left: box.left, top: box.top, right: box.right, bottom: box.bottom },
        hitTag: hit?.tagName ?? null,
        hitClass: hit instanceof HTMLElement ? hit.className : null,
      };
    })()`);
    assert.equal(
      codeFormatTriggerHit.hit,
      true,
      `keypad LaTeX format trigger is blocked: ${JSON.stringify(codeFormatTriggerHit)}`,
    );
    await clickSelector('.editor-code-format-control .code-format-primary');
    const visibleFormatHitTarget = await evaluate(`(() => {
      const item = document.querySelector('[data-format="raw"]');
      if (!item) return { format: null, reason: 'missing-item' };
      const box = item.getBoundingClientRect();
      const hit = document.elementFromPoint(
        box.left + box.width / 2,
        box.top + box.height / 2,
      );
      const menu = item.closest('.code-format-menu');
      const menuRect = menu?.getBoundingClientRect();
      return {
        format: hit?.closest('[data-format="raw"]')?.getAttribute('data-format') ?? null,
        itemRect: { left: box.left, top: box.top, right: box.right, bottom: box.bottom },
        menuRect: menuRect
          ? { left: menuRect.left, top: menuRect.top, right: menuRect.right, bottom: menuRect.bottom }
          : null,
        viewport: { width: innerWidth, height: innerHeight },
        autoAvoided:
          menu instanceof HTMLElement &&
          menu.dataset.visualtexAutoAvoidAdjusted === 'true',
        hitTag: hit?.tagName ?? null,
        hitClass: hit instanceof HTMLElement ? hit.className : null,
      };
    })()`);
    assert.equal(
      visibleFormatHitTarget.format,
      "raw",
      `the visible LaTeX format menu is blocked from real pointer input: ${JSON.stringify(visibleFormatHitTarget)}`,
    );
    assert.equal(
      visibleFormatHitTarget.autoAvoided,
      true,
      `the LaTeX format menu did not use the shared floating-layer auto-avoidance manager: ${JSON.stringify(visibleFormatHitTarget)}`,
    );
    assert.ok(
      visibleFormatHitTarget.menuRect &&
        visibleFormatHitTarget.menuRect.left >= 8 &&
        visibleFormatHitTarget.menuRect.top >= 8 &&
        visibleFormatHitTarget.menuRect.right <= visibleFormatHitTarget.viewport.width - 8 &&
        visibleFormatHitTarget.menuRect.bottom <= visibleFormatHitTarget.viewport.height - 8,
      `the LaTeX format menu escaped the viewport after shared auto-avoidance: ${JSON.stringify(visibleFormatHitTarget)}`,
    );
    await clickSelector('[data-format="raw"]');
    assert.equal(
      await evaluate(`document.querySelector('[data-format="raw"]') === null`),
      true,
      "clicking a visible LaTeX format menu item did not close the menu",
    );

    await clickSelector('.editor-code-format-control .code-format-primary');
    await evaluate(`document.querySelector('[data-format="align-star"]')?.scrollIntoView({ block: 'center' })`);
    await sleep(40);
    const alignFormatHitTarget = await evaluate(`(() => {
      const item = document.querySelector('[data-format="align-star"]');
      if (!item) return null;
      const box = item.getBoundingClientRect();
      const hit = document.elementFromPoint(
        box.left + box.width / 2,
        box.top + box.height / 2,
      );
      return hit?.closest('[data-format="align-star"]')?.getAttribute('data-format') ?? null;
    })()`);
    assert.equal(
      alignFormatHitTarget,
      "align-star",
      "a scrolled LaTeX format menu item is blocked from real pointer input",
    );
    await clickSelector('[data-format="align-star"]');

    await evaluate(`(() => {
      Object.defineProperty(navigator, 'clipboard', {
        configurable: true,
        value: { writeText: async (text) => { window.__visualtexCopied = text; } },
      });
      window.__visualtexCopied = '';
      return true;
    })()`);
    await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', {
      key: 's',
      code: 'KeyS',
      ctrlKey: true,
      bubbles: true,
      cancelable: true,
    }))`);
    await sleep(80);
    const copied = await evaluate(`window.__visualtexCopied ?? ''`);
    assert.match(copied, /\\begin\{align\*\}/, "Ctrl+S did not copy the active align-star source format");
    assert.doesNotMatch(copied, /^\$\$/m, "keypad Ctrl+S fell back to a hard-coded double-dollar format");

    await evaluate(`document.querySelector('[data-keypad-mode-toggle]')?.click()`);
    await sleep(120);
    assert.equal(
      await evaluate(`Boolean(document.querySelector('.app-header'))`),
      true,
      "exiting keypad mode did not restore the normal application header",
    );

    console.log("Keypad mode browser regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(180);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => {});
  }
}

await main();
