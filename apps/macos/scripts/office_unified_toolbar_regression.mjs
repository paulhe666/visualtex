import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { mkdir, rm, writeFile } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 600;
const previewPort = 9400 + offset;
const debugPort = 16400 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const sessionId = "11111111-2222-4333-8444-555555555555";
const officeUrl = `${baseUrl}/office-native-dialog.html?sessionId=${sessionId}&officeHost=word`;
const chromeProfile = `/tmp/visualtex-office-unified-toolbar-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while Vite or Chrome starts.
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

  async evaluate(expression) {
    const result = await this.send("Runtime.evaluate", {
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
        "--window-size=1600,900",
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
    assert.ok(page, "VisualTeX page target must exist");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(350);
    await client.evaluate(`(() => {
      localStorage.setItem(
        "visualtex-editor",
        JSON.stringify({
          state: {
            zoom: 0.5,
            title: "Main application sentinel",
            lines: [{ id: "main-sentinel-line", latex: "MAIN_APP_SENTINEL" }],
            activeLineId: "main-sentinel-line",
            formulaAlignment: "right",
            latexCodeFormat: "equation",
            history: [{ id: "main-history", latex: "MAIN_HISTORY_SENTINEL", createdAt: 1 }],
          },
          version: 0,
        }),
      );
      localStorage.removeItem("visualtex-office-editor-zoom-60-migration-v1");
    })()`);
    await client.send("Page.addScriptToEvaluateOnNewDocument", {
      source: `(() => {
        const numberingProbe = new URLSearchParams(location.search).get("numberingProbe");
        let mockSession = {
          id: ${JSON.stringify(sessionId)},
          mode: numberingProbe === "create" ? "create" : "edit",
          host: "word",
          formulaId: "formula-test",
          sourceDocumentId: "document-test",
          sourceObjectId: "object-test",
          title: "Office toolbar regression",
          lines: [{ id: "line-1", latex: "e^{i\\\\pi}+1=0" }],
          activeLineId: "line-1",
          codeFormat: "raw",
          displayMode: numberingProbe ? "block" : "inline",
          numbered: numberingProbe === "edit-numbered",
          fontSizePt: 10.5,
          exportWidth: 320,
          exportHeight: 80,
          exportResult: null,
          originalMetadata: null,
          dirty: false,
          status: "editing",
          autoCommitOnClose: true,
          explicitCancel: false,
          error: null,
          createdAt: Date.now(),
          updatedAt: Date.now(),
          expiresAt: Date.now() + 600000,
        };
        const originalFetch = globalThis.fetch.bind(globalThis);
        globalThis.fetch = async (input, init = {}) => {
          const raw = typeof input === "string" ? input : input?.url;
          const url = new URL(raw, location.href);
          if (url.pathname === "/api/v1/sessions/${sessionId}") {
            const method = String(init.method || "GET").toUpperCase();
            if (method === "PATCH") {
              const patch = init.body ? JSON.parse(String(init.body)) : {};
              mockSession = { ...mockSession, ...patch, updatedAt: Date.now() };
            }
            return new Response(JSON.stringify(mockSession), {
              status: 200,
              headers: { "Content-Type": "application/json" },
            });
          }
          return originalFetch(input, init);
        };
      })();`,
    });
    await client.send("Page.navigate", { url: officeUrl });

    const waitStarted = Date.now();
    while (Date.now() - waitStarted < 12000) {
      const ready = await client.evaluate(`Boolean(
        document.querySelector('.editor-pane-header.is-office-editor-header') &&
        document.querySelector('[data-office-primary-action]') &&
        document.querySelector('math-field')
      )`);
      if (ready) break;
      await sleep(80);
    }

    const inspect = async (width, height) => {
      await client.send("Emulation.setDeviceMetricsOverride", {
        width,
        height,
        deviceScaleFactor: 1,
        mobile: false,
      });
      await sleep(220);
      return client.evaluate(`(() => {
        const header = document.querySelector('.editor-pane-header.is-office-editor-header');
        const headerRect = header?.getBoundingClientRect();
        const selectorEntries = [
          ['displayMode', '.office-display-mode-setting'],
          ['fontSize', '.office-font-size-setting'],
          ['autoCommit', '.office-auto-commit-setting:not(.is-numbering-setting)'],
          ['alignment', '.formula-alignment-controls'],
          ['inputLogic', '.canvas-input-behavior-trigger'],
          ['ocrModel', '.canvas-ocr-model'],
          ['zoom', '.canvas-controls'],
          ['undo', '[data-office-undo-action]'],
          ['redo', '[data-office-redo-action]'],
          ['cancel', '[data-office-cancel-action]'],
          ['primary', '[data-office-primary-action]'],
        ];
        const items = selectorEntries.flatMap(([name, selector]) => {
          const element = document.querySelector(selector);
          if (!(element instanceof HTMLElement)) return [];
          const style = getComputedStyle(element);
          const rect = element.getBoundingClientRect();
          if (style.display === 'none' || style.visibility === 'hidden' || rect.width <= 0 || rect.height <= 0) return [];
          return [{ name, left: rect.left, right: rect.right, top: rect.top, bottom: rect.bottom }];
        });
        const overlaps = [];
        for (let leftIndex = 0; leftIndex < items.length; leftIndex += 1) {
          for (let rightIndex = leftIndex + 1; rightIndex < items.length; rightIndex += 1) {
            const left = items[leftIndex];
            const right = items[rightIndex];
            const intersects =
              Math.min(left.right, right.right) - Math.max(left.left, right.left) > 1 &&
              Math.min(left.bottom, right.bottom) - Math.max(left.top, right.top) > 1;
            if (intersects) overlaps.push([left.name, right.name]);
          }
        }
        const undo = document.querySelector('[data-office-undo-action]');
        const redo = document.querySelector('[data-office-redo-action]');
        const workspace = document.querySelector('.workspace');
        const editorPane = document.querySelector('.formula-workspace.editor-pane');
        const sidebar = document.querySelector('.workspace > .formula-toolbar');
        const workspaceStyle = workspace ? getComputedStyle(workspace) : null;
        const editorPaneStyle = editorPane ? getComputedStyle(editorPane) : null;
        const headerStyle = header ? getComputedStyle(header) : null;
        const sidebarRect = sidebar?.getBoundingClientRect();
        const editorScroll = document.querySelector('.editor-pane-scroll');
        const editorSurface = document.querySelector('.editor-surface');
        const firstFormulaLine = document.querySelector('.formula-line');
        const firstMathfield = firstFormulaLine?.querySelector('.visual-mathfield');
        const activeFormulaLine = document.querySelector('.formula-line.is-active');
        const editorScrollRect = editorScroll?.getBoundingClientRect();
        const editorSurfaceRect = editorSurface?.getBoundingClientRect();
        const firstFormulaLineRect = firstFormulaLine?.getBoundingClientRect();
        const firstMathfieldRect = firstMathfield?.getBoundingClientRect();
        const activeFormulaStyle = activeFormulaLine
          ? getComputedStyle(activeFormulaLine)
          : null;
        const firstFormulaStyle = firstFormulaLine
          ? getComputedStyle(firstFormulaLine)
          : null;
        const editorSurfaceStyle = editorSurface
          ? getComputedStyle(editorSurface)
          : null;
        const formattingSlot = document.querySelector('.classic-bottom-formatting-slot');
        const formattingControls = document.querySelector('.formula-alignment-controls');
        const bottomTabs = document.querySelector('.classic-bottom-tabs');
        const formattingSlotRect = formattingSlot?.getBoundingClientRect();
        const bottomTabsRect = bottomTabs?.getBoundingClientRect();
        const bottomTabGroup = document.querySelector('.classic-bottom-tab-group');
        const bottomActions = document.querySelector('.classic-bottom-actions');
        const collapseButton = document.querySelector('[data-classic-bottom-collapse]');
        const bottomTabGroupRect = bottomTabGroup?.getBoundingClientRect();
        const bottomActionsRect = bottomActions?.getBoundingClientRect();
        const collapseButtonRect = collapseButton?.getBoundingClientRect();
        const bottomSections = [
          ['formatting', formattingSlotRect],
          ['tabs', bottomTabGroupRect],
          ['actions', bottomActionsRect],
        ].filter((entry) => entry[1]);
        const bottomOverlaps = [];
        for (let leftIndex = 0; leftIndex < bottomSections.length; leftIndex += 1) {
          for (let rightIndex = leftIndex + 1; rightIndex < bottomSections.length; rightIndex += 1) {
            const [leftName, leftRect] = bottomSections[leftIndex];
            const [rightName, rightRect] = bottomSections[rightIndex];
            if (Math.min(leftRect.right, rightRect.right) - Math.max(leftRect.left, rightRect.left) > 1) {
              bottomOverlaps.push([leftName, rightName]);
            }
          }
        }
        const templateStrip = document.querySelector('.classic-bottom-toolbar .template-strip');
        const templateStripStyle = templateStrip ? getComputedStyle(templateStrip) : null;
        const formattingButtons = formattingSlot
          ? Array.from(formattingSlot.querySelectorAll('button')).map((button) => {
              const rect = button.getBoundingClientRect();
              return { width: rect.width, height: rect.height };
            })
          : [];
        return {
          viewport: { width: innerWidth, height: innerHeight },
          layoutDebug: {
            workspaceClass: workspace?.className || '',
            workspaceColumns: workspaceStyle?.gridTemplateColumns || '',
            workspacePaddingTop: workspaceStyle?.paddingTop || '',
            editorPaneDisplay: editorPaneStyle?.display || '',
            editorPanePosition: editorPaneStyle?.position || '',
            headerClass: header?.className || '',
            headerPosition: headerStyle?.position || '',
            headerInset: headerStyle ? [headerStyle.top, headerStyle.right, headerStyle.bottom, headerStyle.left] : [],
            headerOffsetParent: header?.offsetParent?.className || header?.offsetParent?.tagName || '',
            sidebar: sidebarRect ? {
              left: sidebarRect.left,
              right: sidebarRect.right,
              top: sidebarRect.top,
              bottom: sidebarRect.bottom,
            } : null,
          },
          header: headerRect ? {
            left: headerRect.left,
            right: headerRect.right,
            top: headerRect.top,
            bottom: headerRect.bottom,
            height: headerRect.height,
            scrollWidth: header.scrollWidth,
            clientWidth: header.clientWidth,
          } : null,
          items,
          overlaps,
          allInsideHeader: Boolean(headerRect) && items
            .filter((item) => item.name !== 'alignment')
            .every((item) =>
              item.left >= headerRect.left - 1 &&
              item.right <= headerRect.right + 1 &&
              item.top >= headerRect.top - 1 &&
              item.bottom <= headerRect.bottom + 1
            ),
          defaultZoomText:
            document.querySelector('.canvas-controls > span')?.textContent?.trim() || '',
          editorTopGeometry: {
            surfacePaddingTop: editorSurfaceStyle
              ? parseFloat(editorSurfaceStyle.paddingTop)
              : -1,
            surfaceTopGap:
              editorSurfaceRect && editorScrollRect
                ? editorSurfaceRect.top - editorScrollRect.top
                : -1,
            firstLineTopGap:
              editorSurfaceRect && firstFormulaLineRect
                ? firstFormulaLineRect.top - editorSurfaceRect.top
                : -1,
          },
          activeLineAppearance: {
            enabled:
              workspace?.classList.contains('has-active-line-highlight') ?? false,
            background: activeFormulaStyle?.backgroundColor ?? '',
            boxShadow: activeFormulaStyle?.boxShadow ?? '',
          },
          firstLineDivider: {
            width: firstFormulaStyle?.borderBottomWidth ?? '',
            style: firstFormulaStyle?.borderBottomStyle ?? '',
            color: firstFormulaStyle?.borderBottomColor ?? '',
          },
          lineNumberAppearance: {
            enabled:
              editorSurface?.classList.contains('has-line-numbers') ?? false,
            count:
              editorSurface?.querySelectorAll('.formula-line-number').length ?? 0,
            leftFormulaGap:
              firstFormulaLineRect && firstMathfieldRect
                ? firstMathfieldRect.left - firstFormulaLineRect.left
                : -1,
          },
          formulaToolbarScroll: templateStrip
            ? {
                overflowX: templateStripStyle?.overflowX || '',
                overflowY: templateStripStyle?.overflowY || '',
                clientHeight: templateStrip.clientHeight,
                scrollHeight: templateStrip.scrollHeight,
              }
            : null,
          bottomPanelLayout: {
            overlaps: bottomOverlaps,
            labelsVisible: Array.from(document.querySelectorAll('.classic-bottom-tab-label')).some(
              (label) => getComputedStyle(label).display !== 'none',
            ),
            collapseVisible: Boolean(
              collapseButtonRect &&
                bottomTabsRect &&
                collapseButtonRect.width > 0 &&
                collapseButtonRect.height > 0 &&
                collapseButtonRect.left >= bottomTabsRect.left - 1 &&
                collapseButtonRect.right <= bottomTabsRect.right + 1,
            ),
            collapse: collapseButtonRect
              ? {
                  left: collapseButtonRect.left,
                  right: collapseButtonRect.right,
                  top: collapseButtonRect.top,
                  bottom: collapseButtonRect.bottom,
                }
              : null,
          },
          formattingPlacement: {
            topHeaderContainsFormatting: Boolean(
              header && formattingControls && header.contains(formattingControls),
            ),
            bottomSlotContainsFormatting: Boolean(
              formattingSlot && formattingControls && formattingSlot.contains(formattingControls),
            ),
            buttonCount: formattingButtons.length,
            buttonSizes: formattingButtons,
            slot: formattingSlotRect
              ? {
                  left: formattingSlotRect.left,
                  right: formattingSlotRect.right,
                  top: formattingSlotRect.top,
                  bottom: formattingSlotRect.bottom,
                }
              : null,
            tabs: bottomTabsRect
              ? {
                  left: bottomTabsRect.left,
                  right: bottomTabsRect.right,
                  top: bottomTabsRect.top,
                  bottom: bottomTabsRect.bottom,
                }
              : null,
          },
          hasLegacyHeader: Boolean(document.querySelector('.office-dialog-header')),
          hasVisualEditorTitle: Boolean(document.querySelector('.editor-pane-header h1, .editor-pane-header .pane-title-copy')),
          hasVisualTexHostTitle: Array.from(document.querySelectorAll('strong, span')).some((element) =>
            /^(VisualTeX|Microsoft Word|Microsoft PowerPoint)$/.test(element.textContent?.trim() || '')
          ),
          undoText: undo?.textContent?.trim() || '',
          redoText: redo?.textContent?.trim() || '',
          undoHasIcon: Boolean(undo?.querySelector('svg')),
          redoHasIcon: Boolean(redo?.querySelector('svg')),
        };
      })()`);
    };

    const dragElement = async (selector, deltaX, deltaY) => {
      const rect = await client.evaluate(`(() => {
        const element = document.querySelector(${JSON.stringify(selector)});
        if (!(element instanceof HTMLElement)) return null;
        const style = getComputedStyle(element);
        const rect = element.getBoundingClientRect();
        if (style.display === 'none' || rect.width <= 0 || rect.height <= 0) return null;
        return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 };
      })()`);
      if (!rect) throw new Error(`Unable to drag missing element: ${selector}`);
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseMoved",
        x: rect.x,
        y: rect.y,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: rect.x,
        y: rect.y,
        button: "left",
        buttons: 1,
        clickCount: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseMoved",
        x: rect.x + deltaX,
        y: rect.y + deltaY,
        button: "left",
        buttons: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: rect.x + deltaX,
        y: rect.y + deltaY,
        button: "left",
        buttons: 0,
        clickCount: 1,
      });
      await sleep(140);
    };

    const wide = await inspect(1600, 900);
    const compact = await inspect(1280, 820);
    await inspect(700, 500);
    const initialCompactTileResize = await client.evaluate(`(() => {
      const tiles = document.querySelector('.classic-tile-toolbar');
      const handle = document.querySelector('.classic-tile-resizer');
      const tileRect = tiles?.getBoundingClientRect();
      const handleRect = handle?.getBoundingClientRect();
      return {
        tileWidth: tileRect?.width ?? 0,
        handleVisible: Boolean(
          handleRect && handleRect.width > 0 && handleRect.height > 0 &&
          getComputedStyle(handle).display !== 'none'
        ),
      };
    })()`);
    await dragElement('.classic-tile-resizer', -72, 0);
    const resizedCompactTile = await client.evaluate(`(() => {
      const tiles = document.querySelector('.classic-tile-toolbar');
      const handle = document.querySelector('.classic-tile-resizer');
      const tileRect = tiles?.getBoundingClientRect();
      const handleRect = handle?.getBoundingClientRect();
      return {
        tileWidth: tileRect?.width ?? 0,
        handleLeft: handleRect?.left ?? 0,
        tileLeft: tileRect?.left ?? 0,
        storedTileWidth: Number(localStorage.getItem('visualtex-classic-tile-width')),
      };
    })()`);
    assert.equal(
      initialCompactTileResize.handleVisible,
      true,
      JSON.stringify(initialCompactTileResize),
    );
    assert.ok(
      resizedCompactTile.tileWidth >= initialCompactTileResize.tileWidth + 60,
      JSON.stringify({ initialCompactTileResize, resizedCompactTile }),
    );
    assert.ok(
      Math.abs(resizedCompactTile.handleLeft - resizedCompactTile.tileLeft) <= 8,
      JSON.stringify(resizedCompactTile),
    );
    assert.ok(
      Math.abs(resizedCompactTile.storedTileWidth - resizedCompactTile.tileWidth) <= 2,
      JSON.stringify(resizedCompactTile),
    );
    await client.evaluate(`document.querySelector('.classic-tile-resizer')?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }))`);
    await sleep(100);
    const officeFormulaLineCenter = await client.evaluate(`(() => {
      const rect = document.querySelector('.formula-line')?.getBoundingClientRect();
      return rect
        ? { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 }
        : null;
    })()`);
    assert.ok(officeFormulaLineCenter, "Office formula row must exist for hover regression");
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseMoved",
      x: officeFormulaLineCenter.x,
      y: officeFormulaLineCenter.y,
    });
    await sleep(80);
    const officeHoverBackgroundWhileDisabled = await client.evaluate(`(() => {
      const line = document.querySelector('.formula-line');
      return line ? getComputedStyle(line).backgroundColor : '';
    })()`);
    assert.equal(
      officeHoverBackgroundWhileDisabled,
      "rgba(0, 0, 0, 0)",
      `Office row hover must stay transparent while highlighting is disabled: ${officeHoverBackgroundWhileDisabled}`,
    );
    const colorPopover = await client.evaluate(`(async () => {
      const field = document.querySelector('math-field');
      const button = document.querySelector('[data-formula-selection-color]');
      if (!(field instanceof HTMLElement) || !(button instanceof HTMLButtonElement)) {
        return null;
      }
      field.focus();
      const lastOffset = Number(field.lastOffset ?? 1);
      field.selection = {
        ranges: [[0, Math.max(1, Math.min(lastOffset, 2))]],
        direction: 'forward',
      };
      button.dispatchEvent(
        new PointerEvent('pointerdown', {
          bubbles: true,
          cancelable: true,
          pointerId: 1,
          pointerType: 'mouse',
          isPrimary: true,
        }),
      );
      button.click();
      await new Promise((resolve) => requestAnimationFrame(() => resolve()));
      const popover = document.querySelector('[data-formula-color-popover="color"]');
      const tabs = document.querySelector('.classic-bottom-tabs');
      if (!(popover instanceof HTMLElement) || !(tabs instanceof HTMLElement)) {
        return null;
      }
      const popoverRect = popover.getBoundingClientRect();
      const tabsRect = tabs.getBoundingClientRect();
      return {
        popover: {
          left: popoverRect.left,
          right: popoverRect.right,
          top: popoverRect.top,
          bottom: popoverRect.bottom,
        },
        tabs: {
          top: tabsRect.top,
          bottom: tabsRect.bottom,
        },
        viewport: { width: innerWidth, height: innerHeight },
      };
    })()`);

    const narrow = await inspect(500, 300);
    const inputBehaviorGeometry = await client.evaluate(`(async () => {
      const trigger = document.querySelector('.canvas-input-behavior-trigger');
      const editor = document.querySelector('.classic-editor-pane-body');
      const sidebar = document.querySelector('.classic-tile-toolbar');
      if (!(trigger instanceof HTMLButtonElement) || !(editor instanceof HTMLElement)) {
        return null;
      }
      trigger.click();
      await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      const popover = document.querySelector('.input-behavior-popover');
      if (!(popover instanceof HTMLElement)) return null;
      const triggerRect = trigger.getBoundingClientRect();
      const editorRect = editor.getBoundingClientRect();
      const sidebarRect = sidebar?.getBoundingClientRect();
      const popoverRect = popover.getBoundingClientRect();
      const settings = {
        trigger: {
          width: triggerRect.width,
          height: triggerRect.height,
          left: triggerRect.left,
          right: triggerRect.right,
        },
        editor: {
          left: editorRect.left,
          right: editorRect.right,
          top: editorRect.top,
          bottom: editorRect.bottom,
        },
        sidebar: sidebarRect
          ? { left: sidebarRect.left, right: sidebarRect.right }
          : null,
        popover: {
          left: popoverRect.left,
          right: popoverRect.right,
          top: popoverRect.top,
          bottom: popoverRect.bottom,
          width: popoverRect.width,
          height: popoverRect.height,
        },
        compact: popover.classList.contains('is-compact'),
      };
      const mappingsButton = popover.querySelector('[data-open-auto-escape-map]');
      if (mappingsButton instanceof HTMLButtonElement) mappingsButton.click();
      await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      const mappings = document.querySelector('.input-behavior-popover.is-mapping-view');
      const mappingsRect = mappings?.getBoundingClientRect();
      return {
        settings,
        mappings: mappingsRect
          ? {
              left: mappingsRect.left,
              right: mappingsRect.right,
              top: mappingsRect.top,
              bottom: mappingsRect.bottom,
              width: mappingsRect.width,
              height: mappingsRect.height,
              compact: mappings.classList.contains('is-compact'),
              scrollHeight: mappings.scrollHeight,
              clientHeight: mappings.clientHeight,
            }
          : null,
        viewport: { width: innerWidth, height: innerHeight },
      };
    })()`);
    const screenshot = await client.send("Page.captureScreenshot", {
      format: "png",
      captureBeyondViewport: false,
      fromSurface: true,
    });
    await mkdir("build-logs", { recursive: true });
    await writeFile(
      "build-logs/office-unified-toolbar-500x300.png",
      Buffer.from(screenshot.data, "base64"),
    );
    await client.evaluate(`(() => {
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    })()`);
    await sleep(80);
    const expandedTileGeometry = await client.evaluate(`(() => {
      const editor = document.querySelector('.classic-editor-pane-body');
      const tiles = document.querySelector('.classic-tile-toolbar');
      return {
        editorRight: editor?.getBoundingClientRect().right ?? 0,
        editorWidth: editor?.getBoundingClientRect().width ?? 0,
        tileWidth: tiles?.getBoundingClientRect().width ?? 0,
        collapseVisible: Boolean(document.querySelector('[data-formula-tile-collapse]')),
      };
    })()`);
    await client.evaluate(`document.querySelector('[data-formula-tile-collapse]')?.click()`);
    await sleep(100);
    const collapsedTileGeometry = await client.evaluate(`(() => {
      const workspace = document.querySelector('.workspace.is-classic-layout');
      const editor = document.querySelector('.classic-editor-pane-body');
      const expand = document.querySelector('[data-formula-tile-expand]');
      const expandRect = expand?.getBoundingClientRect();
      return {
        hasSidebar: workspace?.classList.contains('has-sidebar') ?? false,
        hasTiles: Boolean(document.querySelector('.classic-tile-toolbar')),
        editorRight: editor?.getBoundingClientRect().right ?? 0,
        editorWidth: editor?.getBoundingClientRect().width ?? 0,
        expandVisible: Boolean(
          expandRect && expandRect.width > 0 && expandRect.height > 0 &&
          expandRect.left >= 0 && expandRect.right <= innerWidth + 1 &&
          expandRect.top >= 0 && expandRect.bottom <= innerHeight + 1
        ),
      };
    })()`);
    await client.evaluate(`document.querySelector('[data-formula-tile-expand]')?.click()`);
    await sleep(100);
    const restoredTileGeometry = await client.evaluate(`(() => ({
      hasSidebar: document.querySelector('.workspace.is-classic-layout')?.classList.contains('has-sidebar') ?? false,
      hasTiles: Boolean(document.querySelector('.classic-tile-toolbar')),
      hasCollapse: Boolean(document.querySelector('[data-formula-tile-collapse]')),
    }))()`);

    for (const state of [wide, compact]) {
      assert.ok(state.header, JSON.stringify(state));
      assert.equal(state.hasLegacyHeader, false, JSON.stringify(state));
      assert.equal(state.hasVisualEditorTitle, false, JSON.stringify(state));
      assert.equal(state.hasVisualTexHostTitle, false, JSON.stringify(state));
      assert.equal(state.undoText, "", JSON.stringify(state));
      assert.equal(state.redoText, "", JSON.stringify(state));
      assert.equal(state.undoHasIcon, true, JSON.stringify(state));
      assert.equal(state.redoHasIcon, true, JSON.stringify(state));
      assert.deepEqual(state.overlaps, [], JSON.stringify(state));
      assert.equal(state.allInsideHeader, true, JSON.stringify(state));
      const cancelItem = state.items.find((item) => item.name === "cancel");
      const primaryItem = state.items.find((item) => item.name === "primary");
      assert.ok(cancelItem && primaryItem, JSON.stringify(state));
      assert.ok(
        Math.abs(cancelItem.top - primaryItem.top) <= 0.5 &&
          Math.abs(cancelItem.bottom - primaryItem.bottom) <= 0.5,
        JSON.stringify({ cancelItem, primaryItem }),
      );
      assert.ok(state.header.height <= 48, JSON.stringify(state));
      assert.ok(state.header.scrollWidth <= state.header.clientWidth + 1, JSON.stringify(state));
      assert.ok(
        Math.abs(state.header.right - state.viewport.width) <= 1,
        JSON.stringify(state),
      );
      assert.equal(
        state.items.some((item) => item.name === "autoCommit"),
        false,
        JSON.stringify(state),
      );
      assert.ok(
        state.items
          .filter((item) => item.name !== "alignment")
          .every((item) => Math.abs(item.bottom - item.top - 30) <= 1),
        JSON.stringify(state),
      );
      assert.equal(
        state.formattingPlacement.topHeaderContainsFormatting,
        false,
        JSON.stringify(state),
      );
      assert.equal(
        state.formattingPlacement.bottomSlotContainsFormatting,
        true,
        JSON.stringify(state),
      );
      assert.equal(state.formattingPlacement.buttonCount, 7, JSON.stringify(state));
      assert.ok(
        state.formattingPlacement.buttonSizes.every(
          (button) => button.height >= 28 && button.height <= 30,
        ),
        JSON.stringify(state),
      );
      assert.ok(
        state.formattingPlacement.slot &&
          state.formattingPlacement.tabs &&
          state.formattingPlacement.slot.left <=
            state.formattingPlacement.tabs.left + 12 &&
          state.formattingPlacement.slot.top >=
            state.formattingPlacement.tabs.top - 1 &&
          state.formattingPlacement.slot.bottom <=
            state.formattingPlacement.tabs.bottom + 1,
        JSON.stringify(state),
      );
      assert.equal(state.defaultZoomText, "60%", JSON.stringify(state));
      assert.equal(state.editorTopGeometry.surfacePaddingTop, 6, JSON.stringify(state));
      assert.ok(
        state.editorTopGeometry.surfaceTopGap >= 0 &&
          state.editorTopGeometry.surfaceTopGap <= 1,
        JSON.stringify(state),
      );
      assert.ok(
        state.editorTopGeometry.firstLineTopGap >= 6 &&
          state.editorTopGeometry.firstLineTopGap <= 8,
        JSON.stringify(state),
      );
      assert.equal(state.activeLineAppearance.enabled, false, JSON.stringify(state));
      assert.equal(state.firstLineDivider.width, "1px", JSON.stringify(state));
      assert.equal(state.firstLineDivider.style, "solid", JSON.stringify(state));
      assert.notEqual(
        state.firstLineDivider.color,
        "rgba(0, 0, 0, 0)",
        JSON.stringify(state),
      );
      assert.equal(state.lineNumberAppearance.enabled, false, JSON.stringify(state));
      assert.equal(state.lineNumberAppearance.count, 0, JSON.stringify(state));
      assert.ok(
        state.lineNumberAppearance.leftFormulaGap >= 0 &&
          state.lineNumberAppearance.leftFormulaGap <= 1,
        JSON.stringify(state),
      );
      assert.equal(
        state.activeLineAppearance.background,
        "rgba(0, 0, 0, 0)",
        JSON.stringify(state),
      );
      assert.equal(
        state.activeLineAppearance.boxShadow,
        "none",
        JSON.stringify(state),
      );
      assert.ok(state.formulaToolbarScroll, JSON.stringify(state));
      assert.equal(
        state.formulaToolbarScroll.overflowY,
        "hidden",
        JSON.stringify(state),
      );
      assert.ok(
        state.formulaToolbarScroll.scrollHeight <=
          state.formulaToolbarScroll.clientHeight + 1,
        JSON.stringify(state),
      );
      assert.ok(
        state.layoutDebug.sidebar &&
          state.layoutDebug.sidebar.top >= state.header.bottom - 1 &&
          state.layoutDebug.sidebar.right >= state.viewport.width - 1,
        JSON.stringify(state),
      );
      const itemWidth = (name) => {
        const item = state.items.find((entry) => entry.name === name);
        return item ? item.right - item.left : Infinity;
      };
      assert.ok(itemWidth("fontSize") <= 140, JSON.stringify(state));
      assert.ok(itemWidth("inputLogic") <= 94, JSON.stringify(state));
      assert.ok(itemWidth("ocrModel") <= 138, JSON.stringify(state));
    }

    assert.deepEqual(narrow.bottomPanelLayout.overlaps, [], JSON.stringify(narrow));
    assert.equal(narrow.bottomPanelLayout.labelsVisible, false, JSON.stringify(narrow));
    assert.equal(narrow.bottomPanelLayout.collapseVisible, true, JSON.stringify(narrow));
    assert.ok(narrow.formulaToolbarScroll, JSON.stringify(narrow));
    assert.equal(narrow.formulaToolbarScroll.overflowY, "hidden", JSON.stringify(narrow));
    assert.ok(
      narrow.formulaToolbarScroll.scrollHeight <=
        narrow.formulaToolbarScroll.clientHeight + 1,
      JSON.stringify(narrow),
    );

    assert.ok(inputBehaviorGeometry, "Input behavior popover geometry must be measurable");
    assert.ok(
      Math.abs(
        inputBehaviorGeometry.settings.trigger.width -
          inputBehaviorGeometry.settings.trigger.height,
      ) <= 1,
      JSON.stringify(inputBehaviorGeometry),
    );
    assert.ok(
      inputBehaviorGeometry.settings.trigger.width <= 31,
      JSON.stringify(inputBehaviorGeometry),
    );
    assert.equal(
      inputBehaviorGeometry.settings.compact,
      true,
      JSON.stringify(inputBehaviorGeometry),
    );
    assert.ok(
      inputBehaviorGeometry.settings.popover.left >=
        inputBehaviorGeometry.settings.editor.left + 7 &&
        inputBehaviorGeometry.settings.popover.right <=
          inputBehaviorGeometry.settings.editor.right - 7 &&
        inputBehaviorGeometry.settings.popover.top >= 0 &&
        inputBehaviorGeometry.settings.popover.bottom <=
          inputBehaviorGeometry.viewport.height - 7,
      JSON.stringify(inputBehaviorGeometry),
    );
    assert.ok(
      !inputBehaviorGeometry.settings.sidebar ||
        inputBehaviorGeometry.settings.popover.right <=
          inputBehaviorGeometry.settings.sidebar.left - 7,
      JSON.stringify(inputBehaviorGeometry),
    );
    assert.ok(inputBehaviorGeometry.mappings, JSON.stringify(inputBehaviorGeometry));
    assert.equal(
      inputBehaviorGeometry.mappings.compact,
      true,
      JSON.stringify(inputBehaviorGeometry),
    );
    assert.ok(
      inputBehaviorGeometry.mappings.left >=
        inputBehaviorGeometry.settings.editor.left + 7 &&
        inputBehaviorGeometry.mappings.right <=
          inputBehaviorGeometry.settings.editor.right - 7 &&
        inputBehaviorGeometry.mappings.top >= 0 &&
        inputBehaviorGeometry.mappings.bottom <=
          inputBehaviorGeometry.viewport.height - 7 &&
        inputBehaviorGeometry.mappings.clientHeight <=
          inputBehaviorGeometry.viewport.height - 16,
      JSON.stringify(inputBehaviorGeometry),
    );

    assert.equal(expandedTileGeometry.collapseVisible, true, JSON.stringify(expandedTileGeometry));
    assert.equal(collapsedTileGeometry.hasSidebar, false, JSON.stringify(collapsedTileGeometry));
    assert.equal(collapsedTileGeometry.hasTiles, false, JSON.stringify(collapsedTileGeometry));
    assert.equal(collapsedTileGeometry.expandVisible, true, JSON.stringify(collapsedTileGeometry));
    assert.ok(
      collapsedTileGeometry.editorWidth >=
        expandedTileGeometry.editorWidth + expandedTileGeometry.tileWidth - 2 &&
        collapsedTileGeometry.editorRight >= narrow.viewport.width - 1,
      JSON.stringify({ expandedTileGeometry, collapsedTileGeometry }),
    );
    assert.deepEqual(
      restoredTileGeometry,
      { hasSidebar: true, hasTiles: true, hasCollapse: true },
      JSON.stringify(restoredTileGeometry),
    );

    assert.ok(colorPopover, "The moved font-color button must open its popover");
    assert.ok(
      colorPopover.popover.bottom <= colorPopover.tabs.top + 1,
      JSON.stringify(colorPopover),
    );
    assert.ok(
      colorPopover.popover.top >= 0 &&
        colorPopover.popover.left >= 0 &&
        colorPopover.popover.right <= colorPopover.viewport.width + 1,
      JSON.stringify(colorPopover),
    );

    const migratedZoom = await client.evaluate(`(() => {
      const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      return {
        marker: localStorage.getItem("visualtex-office-editor-zoom-60-migration-v1"),
        zoom: persisted?.state?.zoom,
        title: persisted?.state?.title,
        lineLatex: persisted?.state?.lines?.[0]?.latex,
        activeLineId: persisted?.state?.activeLineId,
        formulaAlignment: persisted?.state?.formulaAlignment,
        latexCodeFormat: persisted?.state?.latexCodeFormat,
        historyLatex: persisted?.state?.history?.[0]?.latex,
        text: document.querySelector('.canvas-controls > span')?.textContent?.trim() || '',
      };
    })()`);
    assert.equal(migratedZoom.marker, "done", JSON.stringify(migratedZoom));
    assert.ok(Math.abs(Number(migratedZoom.zoom) - 0.6) < 0.001, JSON.stringify(migratedZoom));
    assert.equal(migratedZoom.title, "Main application sentinel", JSON.stringify(migratedZoom));
    assert.equal(migratedZoom.lineLatex, "MAIN_APP_SENTINEL", JSON.stringify(migratedZoom));
    assert.equal(migratedZoom.activeLineId, "main-sentinel-line", JSON.stringify(migratedZoom));
    assert.equal(migratedZoom.formulaAlignment, "right", JSON.stringify(migratedZoom));
    assert.equal(migratedZoom.latexCodeFormat, "equation", JSON.stringify(migratedZoom));
    assert.equal(migratedZoom.historyLatex, "MAIN_HISTORY_SENTINEL", JSON.stringify(migratedZoom));
    assert.equal(migratedZoom.text, "60%", JSON.stringify(migratedZoom));

    await client.evaluate(`document.querySelector('.canvas-controls button:last-of-type')?.click()`);
    await sleep(180);
    const adjustedZoom = await client.evaluate(`(() => {
      const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      return {
        zoom: persisted?.state?.zoom,
        text: document.querySelector('.canvas-controls > span')?.textContent?.trim() || '',
      };
    })()`);
    assert.ok(Math.abs(Number(adjustedZoom.zoom) - 0.65) < 0.001, JSON.stringify(adjustedZoom));
    assert.equal(adjustedZoom.text, "65%", JSON.stringify(adjustedZoom));

    const collapsedPanelsBeforeReload = await client.evaluate(`(async () => {
      const bottomCollapse = document.querySelector('[data-classic-bottom-collapse]');
      if (bottomCollapse?.getAttribute('aria-expanded') !== 'false') bottomCollapse?.click();
      document.querySelector('[data-formula-tile-collapse]')?.click();
      await new Promise((resolve) => setTimeout(resolve, 120));
      return {
        toolbarStored: localStorage.getItem('visualtex-office-editor-toolbar-open'),
        tilesStored: localStorage.getItem('visualtex-office-editor-tiles-open'),
        bottomExpanded: document.querySelector('[data-classic-bottom-collapse]')?.getAttribute('aria-expanded'),
        hasBottomContent: Boolean(document.querySelector('.classic-bottom-content')),
        hasTiles: Boolean(document.querySelector('.classic-tile-toolbar')),
        hasTileExpand: Boolean(document.querySelector('[data-formula-tile-expand]')),
      };
    })()`);
    assert.deepEqual(
      collapsedPanelsBeforeReload,
      {
        toolbarStored: "false",
        tilesStored: "false",
        bottomExpanded: "false",
        hasBottomContent: false,
        hasTiles: false,
        hasTileExpand: true,
      },
      JSON.stringify(collapsedPanelsBeforeReload),
    );

    await client.send("Page.navigate", { url: `${officeUrl}&reload=1` });
    const reloadStarted = Date.now();
    while (Date.now() - reloadStarted < 12000) {
      const ready = await client.evaluate(`Boolean(
        document.querySelector('.editor-pane-header.is-office-editor-header') &&
        document.querySelector('[data-office-primary-action]') &&
        document.querySelector('math-field') &&
        document.querySelector('.canvas-controls > span')?.textContent?.trim()
      )`);
      if (ready) break;
      await sleep(80);
    }
    const restoredZoom = await client.evaluate(`(() => {
      const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      return {
        marker: localStorage.getItem("visualtex-office-editor-zoom-60-migration-v1"),
        zoom: persisted?.state?.zoom,
        text: document.querySelector('.canvas-controls > span')?.textContent?.trim() || '',
      };
    })()`);
    assert.equal(restoredZoom.marker, "done", JSON.stringify(restoredZoom));
    assert.ok(Math.abs(Number(restoredZoom.zoom) - 0.65) < 0.001, JSON.stringify(restoredZoom));
    assert.equal(restoredZoom.text, "65%", JSON.stringify(restoredZoom));

    const restoredPanels = await client.evaluate(`(() => ({
      toolbarStored: localStorage.getItem('visualtex-office-editor-toolbar-open'),
      tilesStored: localStorage.getItem('visualtex-office-editor-tiles-open'),
      bottomExpanded: document.querySelector('[data-classic-bottom-collapse]')?.getAttribute('aria-expanded'),
      hasBottomContent: Boolean(document.querySelector('.classic-bottom-content')),
      hasSidebar: document.querySelector('.workspace.is-classic-layout')?.classList.contains('has-sidebar') ?? false,
      hasTiles: Boolean(document.querySelector('.classic-tile-toolbar')),
      hasTileExpand: Boolean(document.querySelector('[data-formula-tile-expand]')),
    }))()`);
    assert.deepEqual(
      restoredPanels,
      {
        toolbarStored: "false",
        tilesStored: "false",
        bottomExpanded: "false",
        hasBottomContent: false,
        hasSidebar: false,
        hasTiles: false,
        hasTileExpand: true,
      },
      JSON.stringify(restoredPanels),
    );

    const numberingPreferenceKey = "visualtex.office.word.create.numbered";
    const numberingProbeUrl = `${officeUrl}&numberingProbe=create`;
    const waitForNumberingCheckbox = async () => {
      const started = Date.now();
      while (Date.now() - started < 12000) {
        const ready = await client.evaluate(`Boolean(
          document.querySelector('.is-numbering-setting input[type="checkbox"]')
        )`);
        if (ready) return;
        await sleep(80);
      }
      throw new Error("Office numbering preference probe did not become ready");
    };

    await client.evaluate(`localStorage.setItem(${JSON.stringify(numberingPreferenceKey)}, "false")`);
    await client.send("Page.navigate", { url: numberingProbeUrl });
    await waitForNumberingCheckbox();
    const numberingInitial = await client.evaluate(`(() => {
      const checkbox = document.querySelector('.is-numbering-setting input[type="checkbox"]');
      return {
        checked: checkbox?.checked ?? null,
        disabled: checkbox?.disabled ?? null,
        stored: localStorage.getItem(${JSON.stringify(numberingPreferenceKey)}),
      };
    })()`);
    assert.deepEqual(
      numberingInitial,
      { checked: false, disabled: false, stored: "false" },
      JSON.stringify(numberingInitial),
    );

    await client.evaluate(`document.querySelector('.is-numbering-setting input[type="checkbox"]')?.click()`);
    await sleep(120);
    const numberingAfterEnable = await client.evaluate(`(() => {
      const checkbox = document.querySelector('.is-numbering-setting input[type="checkbox"]');
      return {
        checked: checkbox?.checked ?? null,
        stored: localStorage.getItem(${JSON.stringify(numberingPreferenceKey)}),
      };
    })()`);
    assert.deepEqual(
      numberingAfterEnable,
      { checked: true, stored: "true" },
      JSON.stringify(numberingAfterEnable),
    );

    await client.send("Page.navigate", { url: numberingProbeUrl });
    await waitForNumberingCheckbox();
    const numberingRestoredEnabled = await client.evaluate(`(() => {
      const checkbox = document.querySelector('.is-numbering-setting input[type="checkbox"]');
      return {
        checked: checkbox?.checked ?? null,
        stored: localStorage.getItem(${JSON.stringify(numberingPreferenceKey)}),
      };
    })()`);
    assert.deepEqual(
      numberingRestoredEnabled,
      { checked: true, stored: "true" },
      JSON.stringify(numberingRestoredEnabled),
    );

    await client.evaluate(`document.querySelector('.is-numbering-setting input[type="checkbox"]')?.click()`);
    await sleep(120);
    await client.send("Page.navigate", { url: numberingProbeUrl });
    await waitForNumberingCheckbox();
    const numberingRestoredDisabled = await client.evaluate(`(() => {
      const checkbox = document.querySelector('.is-numbering-setting input[type="checkbox"]');
      return {
        checked: checkbox?.checked ?? null,
        stored: localStorage.getItem(${JSON.stringify(numberingPreferenceKey)}),
      };
    })()`);
    assert.deepEqual(
      numberingRestoredDisabled,
      { checked: false, stored: "false" },
      JSON.stringify(numberingRestoredDisabled),
    );

    await client.evaluate(`localStorage.setItem(${JSON.stringify(numberingPreferenceKey)}, "true")`);
    await client.send("Page.navigate", { url: `${officeUrl}&numberingProbe=edit-unnumbered` });
    await waitForNumberingCheckbox();
    const editUnnumberedIgnoresPreference = await client.evaluate(`(() => {
      const checkbox = document.querySelector('.is-numbering-setting input[type="checkbox"]');
      return {
        checked: checkbox?.checked ?? null,
        disabled: checkbox?.disabled ?? null,
        stored: localStorage.getItem(${JSON.stringify(numberingPreferenceKey)}),
      };
    })()`);
    assert.deepEqual(
      editUnnumberedIgnoresPreference,
      { checked: false, disabled: true, stored: "true" },
      JSON.stringify(editUnnumberedIgnoresPreference),
    );

    await client.evaluate(`localStorage.setItem(${JSON.stringify(numberingPreferenceKey)}, "false")`);
    await client.send("Page.navigate", { url: `${officeUrl}&numberingProbe=edit-numbered` });
    await waitForNumberingCheckbox();
    const editNumberedIgnoresPreference = await client.evaluate(`(() => {
      const checkbox = document.querySelector('.is-numbering-setting input[type="checkbox"]');
      return {
        checked: checkbox?.checked ?? null,
        disabled: checkbox?.disabled ?? null,
        stored: localStorage.getItem(${JSON.stringify(numberingPreferenceKey)}),
      };
    })()`);
    assert.deepEqual(
      editNumberedIgnoresPreference,
      { checked: true, disabled: true, stored: "false" },
      JSON.stringify(editNumberedIgnoresPreference),
    );

    const numberingPreferenceRegression = {
      initial: numberingInitial,
      afterEnable: numberingAfterEnable,
      restoredEnabled: numberingRestoredEnabled,
      restoredDisabled: numberingRestoredDisabled,
      editUnnumbered: editUnnumberedIgnoresPreference,
      editNumbered: editNumberedIgnoresPreference,
    };

    console.log(JSON.stringify({
      wide,
      compact,
      narrow,
      inputBehaviorGeometry,
      expandedTileGeometry,
      collapsedTileGeometry,
      restoredTileGeometry,
      screenshot: "build-logs/office-unified-toolbar-500x300.png",
      colorPopover,
      migratedZoom,
      adjustedZoom,
      restoredZoom,
      collapsedPanelsBeforeReload,
      restoredPanels,
      numberingPreferenceRegression,
    }, null, 2));
    console.log("Office unified toolbar regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(180);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => undefined);
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
