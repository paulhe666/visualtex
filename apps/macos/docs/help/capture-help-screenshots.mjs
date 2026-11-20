import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { mkdir, rm, writeFile } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 600;
const previewPort = 8200 + offset;
const debugPort = 13200 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = `/tmp/visualtex-help-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const outputDir = new URL("./assets/", import.meta.url);
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
  close() { this.socket?.close(); }
}

async function main() {
  await mkdir(outputDir, { recursive: true });
  const preview = spawn(
    process.execPath,
    ["node_modules/vite/bin/vite.js", "preview", "--host", "127.0.0.1", "--port", String(previewPort), "--strictPort"],
    { cwd: process.cwd(), stdio: "ignore" },
  );
  let chrome;
  let client;
  try {
    await waitFor(baseUrl);
    chrome = spawn(chromePath, [
      "--headless=new",
      "--disable-gpu",
      "--no-first-run",
      "--no-default-browser-check",
      `--remote-debugging-port=${debugPort}`,
      `--user-data-dir=${chromeProfile}`,
      "--window-size=1440,1000",
      baseUrl,
    ], { stdio: "ignore" });
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
    const page = targets.find((target) => target.type === "page" && target.url.startsWith(baseUrl));
    if (!page) throw new Error("VisualTeX page target not found");
    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Emulation.setDeviceMetricsOverride", { width: 1440, height: 1000, deviceScaleFactor: 1, mobile: false });

    const evaluate = async (expression) => {
      const response = await client.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
      if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text || "evaluate failed");
      return response.result.value;
    };
    const clickText = async (needle) => evaluate(`(() => {
      const needle = ${JSON.stringify(needle)};
      const nodes = Array.from(document.querySelectorAll('button,[role="menuitem"],[role="tab"]'));
      const node = nodes.find((el) => [el.textContent, el.getAttribute('aria-label'), el.getAttribute('title')].filter(Boolean).some((value) => value.includes(needle)));
      if (!node) return false;
      node.click();
      return true;
    })()`);
    const dismissOfficeRepairPrompt = async () => {
      const dismissed = await evaluate(`(() => {
        const dialogs = Array.from(document.querySelectorAll('[role="dialog"], .modal, .dialog'));
        const prompt = dialogs.find((node) => /修复 VisualTeX Office 插件|Repair VisualTeX Office Add-ins/.test(node.textContent ?? ''));
        if (!prompt) return false;
        const button = Array.from(prompt.querySelectorAll('button')).find((node) => /稍后处理|Later/.test(node.textContent ?? ''));
        if (!button) throw new Error('Office repair prompt is visible but the Later button was not found');
        button.click();
        return true;
      })()`);
      if (dismissed) await sleep(180);
      return dismissed;
    };
    const assertNoOfficeRepairPrompt = async (name) => {
      const state = await evaluate(`(() => {
        const dialogs = Array.from(document.querySelectorAll('[role="dialog"], .modal, .dialog'));
        const prompt = dialogs.find((node) => /修复 VisualTeX Office 插件|Repair VisualTeX Office Add-ins/.test(node.textContent ?? ''));
        return prompt ? { visible: true, text: (prompt.textContent ?? '').slice(0, 500) } : { visible: false, text: '' };
      })()`);
      assert.equal(state.visible, false, `Office repair prompt blocks ${name}: ${JSON.stringify(state)}`);
    };
    const capture = async (name, selector = null, margin = 14) => {
      await dismissOfficeRepairPrompt();
      await assertNoOfficeRepairPrompt(name);
      let clip;
      if (selector) {
        clip = await evaluate(`(() => {
          const el = document.querySelector(${JSON.stringify(selector)});
          if (!el) return null;
          const r = el.getBoundingClientRect();
          return { x: Math.max(0, r.left-${margin}), y: Math.max(0, r.top-${margin}), width: Math.min(innerWidth-Math.max(0,r.left-${margin}), r.width+${margin*2}), height: Math.min(innerHeight-Math.max(0,r.top-${margin}), r.height+${margin*2}), scale: 1 };
        })()`);
      }
      const shot = await client.send("Page.captureScreenshot", { format: "png", fromSurface: true, captureBeyondViewport: false, ...(clip ? { clip } : {}) });
      await writeFile(new URL(name, outputDir), Buffer.from(shot.data, "base64"));
    };
    const key = async (key, code, text = key, modifiers = 0) => {
      const special = key === "Enter" ? 13 : key === "Tab" ? 9 : key === "Escape" ? 27 : 0;
      const vk = special || (key.length === 1 ? key.toUpperCase().charCodeAt(0) : 0);
      const common = { key, code, modifiers, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk };
      await client.send("Input.dispatchKeyEvent", { type: "keyDown", ...common, text, unmodifiedText: text });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(35);
    };

    await sleep(650);
    await evaluate(`(() => {
      localStorage.setItem('visualtex.onboarding.v3.completed','true');
      localStorage.setItem('visualtex.onboarding.macos.desktop.v1.2.0.completed','true');
      return true;
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(700);
    await dismissOfficeRepairPrompt();
    await assertNoOfficeRepairPrompt('startup');

    // 01 — actual main editor.
    await evaluate(`(() => {
      const field = document.querySelector('math-field');
      if (!field) return false;
      field.setValue('x=\\\\frac{-b\\\\pm\\\\sqrt{b^2-4ac}}{2a}', { mode:'math', format:'latex', insertionMode:'replaceAll', selectionMode:'after' });
      field.blur();
      return true;
    })()`);
    await sleep(180);
    await capture("01-main-editor.png");

    // 02 — align*: type literal & through the real editor event path.
    await evaluate(`document.querySelector('.code-format-primary')?.click()`);
    await sleep(100);
    await evaluate(`document.querySelector('[data-format="align-star"]')?.click()`);
    await sleep(120);
    await evaluate(`(() => { const f=document.querySelector('math-field'); f?.setValue('',{insertionMode:'replaceAll'}); f?.focus(); f?.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus(); return !!f; })()`);
    await key('a','KeyA','a');
    await key('&','Digit7','&',8);
    await key('=','Equal','=');
    await key('b','KeyB','b');
    await key('+','Equal','+',8);
    await key('c','KeyC','c');
    await key('Enter','Enter','\r');
    for (const ch of 'longvariable') await key(ch, `Key${ch.toUpperCase()}`, ch);
    await key('&','Digit7','&',8);
    await key('=','Equal','=');
    await key('d','KeyD','d');
    await key('-','Minus','-');
    await key('e','KeyE','e');
    await sleep(200);
    // Open source pane/tab when present so the literal & is visible beside the rendered alignment.
    await clickText('LaTeX 源码');
    await sleep(180);
    const alignState = await evaluate(`(() => {
      const fields = Array.from(document.querySelectorAll('.formula-line math-field')).filter((field) => field.shadowRoot?.querySelector('.visualtex-align-marker'));
      const anchors = fields.map((field) => field.shadowRoot?.querySelector('.visualtex-align-marker')?.getBoundingClientRect().left ?? null);
      const sourceText = Array.from(document.querySelectorAll('textarea,.cm-content,[contenteditable="true"]')).map((el) => el.value ?? el.textContent ?? '').join(' | ');
      return { rows: fields.length, anchors, sourceHasAmpersand: sourceText.includes('&') };
    })()`);
    assert.equal(alignState.rows, 2, `Expected two explicit align rows: ${JSON.stringify(alignState)}`);
    assert.ok(alignState.anchors.every((value) => typeof value === 'number'), `Missing align anchors: ${JSON.stringify(alignState)}`);
    assert.ok(Math.abs(alignState.anchors[0] - alignState.anchors[1]) <= 2, `Align anchors differ: ${JSON.stringify(alignState)}`);
    await capture("02-align-ampersand.png", ".workspace", 0);

    // Restore tools view and open a real command context menu from the Arrow category.
    await clickText('公式工具');
    await sleep(120);
    const arrowCategorySelected = await evaluate(`(() => {
      const tab = document.querySelector('.formula-toolbar [data-category="arrow"]');
      if (!tab) return false;
      tab.click();
      return true;
    })()`);
    assert.equal(arrowCategorySelected, true, 'Arrow category tab was not found');
    await sleep(220);
    const activeArrowCategory = await evaluate(`document.querySelector('.formula-toolbar .template-strip')?.getAttribute('data-active-category')`);
    assert.equal(activeArrowCategory, 'arrow', `Expected Arrow category, got ${activeArrowCategory}`);
    const contextInfo = await evaluate(`(async () => {
      const scope = document.querySelector('.formula-toolbar [data-toolbar-category-section="arrow"]') || document.querySelector('.formula-toolbar .template-strip');
      const tool = scope?.querySelector('button[data-command-id]');
      if (!tool) return null;
      tool.scrollIntoView({ block:'center', inline:'center' });
      await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      const r=tool.getBoundingClientRect();
      if (r.width <= 0 || r.height <= 0) return null;
      const propKey = Object.keys(tool).find((key) => key.startsWith('__reactProps$'));
      const onContextMenu = propKey ? tool[propKey]?.onContextMenu : null;
      if (typeof onContextMenu !== 'function') return { id:tool.getAttribute('data-command-id'), error:'context handler missing' };
      onContextMenu({
        preventDefault(){},
        stopPropagation(){},
        clientX:Math.max(1, Math.min(innerWidth - 1, r.left+r.width/2)),
        clientY:Math.max(1, Math.min(innerHeight - 1, r.top+r.height/2)),
      });
      return {
        id: tool.getAttribute('data-command-id'),
        label:[tool.textContent,tool.getAttribute('aria-label'),tool.getAttribute('title')].filter(Boolean).join(' | '),
        x:Math.max(1, Math.min(innerWidth - 1, r.left+r.width/2)),
        y:Math.max(1, Math.min(innerHeight - 1, r.top+r.height/2)),
      };
    })()`);
    assert.ok(contextInfo && !contextInfo.error, `No visible formula tool context handler was found in the Arrow category: ${JSON.stringify(contextInfo)}`);
    await sleep(220);
    const contextMenuState = await evaluate(`(() => {
      const menus = Array.from(document.querySelectorAll('[role="menu"], .formula-tile-context-menu, .formula-hotkey-context-menu'));
      const text = menus.map((node) => node.textContent ?? '').join(' | ');
      const bodyText = document.body.textContent ?? '';
      return { count: menus.length, text, hasCommon: text.includes('常用'), hasHotkey: text.includes('快捷键'), bodyHasHotkey: bodyText.includes('设置快捷键') || bodyText.includes('修改快捷键') };
    })()`);
    console.log("context tool debug", contextInfo, contextMenuState);
    assert.equal(contextMenuState.hasCommon, true, `Tool context menu did not expose Common action: ${JSON.stringify({ contextInfo, contextMenuState })}`);
    assert.equal(contextMenuState.hasHotkey, true, `Tool context menu did not expose hotkey action: ${JSON.stringify(contextMenuState)}`);
    console.log("context tool", contextInfo, contextMenuState);
    await capture("03-tool-context-menu.png", ".workspace", 0);

    // 04 — shortcut configuration dialog reached from that exact context menu.
    const shortcutOpened = await evaluate(`(() => {
      const menu = document.querySelector('.formula-hotkey-context-menu');
      if (!menu) return false;
      const button = Array.from(menu.querySelectorAll('button')).find((node) => /快捷键|hotkey/i.test(node.textContent ?? ''));
      if (!button) return false;
      button.click();
      return true;
    })()`);
    await sleep(220);
    assert.equal(shortcutOpened, true, 'Shortcut action could not be opened from the real context menu');
    const shortcutState = await evaluate(`(() => {
      const dialog = document.querySelector('.formula-hotkey-recorder-dialog[role="dialog"]');
      return { visible: !!dialog, text: dialog?.textContent ?? '' };
    })()`);
    assert.equal(shortcutState.visible, true, `Shortcut dialog is not visible: ${JSON.stringify(shortcutState)}`);
    assert.match(shortcutState.text, /快捷键|hotkey/i, `Shortcut dialog content is wrong: ${JSON.stringify(shortcutState)}`);
    console.log("shortcut opened", shortcutOpened);
    await capture("04-custom-shortcut.png");
    await key('Escape','Escape','');
    await sleep(120);

    // 05 — Tiles/custom area.
    const tilesOpened = await evaluate(`(() => {
      const button = document.querySelector('.formula-toolbar [data-toolbar-view="tiles"]');
      if (!button) return false;
      button.click();
      return true;
    })()`);
    assert.equal(tilesOpened, true, 'Tiles view button was not found');
    await sleep(180);
    const customTilesOpened = await evaluate(`(() => {
      const button = document.querySelector('.formula-toolbar [data-tile-category="custom"]');
      if (!button) return false;
      button.click();
      return true;
    })()`);
    assert.equal(customTilesOpened, true, 'Custom tiles category button was not found');
    await sleep(180);
    const customTilesState = await evaluate(`(() => ({
      panel: !!document.querySelector('.formula-tiles-panel'),
      controls: !!document.querySelector('.custom-formula-tile-controls'),
      customActive: document.querySelector('[data-tile-category="custom"]')?.getAttribute('aria-pressed') === 'true',
    }))()`);
    assert.equal(customTilesState.panel && customTilesState.controls && customTilesState.customActive, true, `Custom tiles area is not visible: ${JSON.stringify(customTilesState)}`);
    await capture("05-custom-tiles.png", ".workspace", 0);

    // 10 — actual custom symbol designer.
    const designerOpened = await evaluate(`(() => {
      const button = document.querySelector('[data-open-custom-symbol-designer]');
      if (!button) return false;
      button.click();
      return true;
    })()`);
    await sleep(240);
    assert.equal(designerOpened, true, 'Custom symbol designer could not be opened');
    const designerState = await evaluate(`(() => {
      const dialog = document.querySelector('.custom-symbol-designer-dialog,[role="dialog"]');
      return { visible: !!dialog, text: dialog?.textContent ?? '' };
    })()`);
    assert.equal(designerState.visible, true, 'Custom symbol designer dialog is not visible');
    assert.match(designerState.text, /字符设计器|Custom Symbol Designer/, 'Custom symbol designer dialog content is wrong');
    await capture("10-custom-symbol-designer.png");
    await key('Escape','Escape','');
    await sleep(120);

    // 06 — input behavior popover.
    await clickText('公式工具');
    await sleep(100);
    const behaviorOpened = await clickText('输入行为');
    await sleep(180);
    console.log("input behavior opened", behaviorOpened);
    await capture("06-input-behavior.png");

    // 07 — settings, actual application dialog.
    await key('Escape','Escape','');
    await clickText('设置');
    await sleep(200);
    await capture("07-settings.png");
    const themeStudioOpened = await clickText('界面自定义');
    await sleep(220);
    assert.equal(themeStudioOpened, true, 'Interface customization could not be opened');
    await capture("11-theme-studio.png");
    await key('Escape','Escape','');
    await sleep(100);

    // 08 — actual cases rendering in the current editor.
    await key('Escape','Escape','');
    await sleep(100);
    await evaluate(`(() => {
      const field=document.querySelector('math-field');
      if (!field) return false;
      field.setValue('f(x)=\\\\begin{cases}x^2 & x>0 \\\\\\\\ 0 & x=0 \\\\\\\\ -x & x<0\\\\end{cases}', { mode:'math', format:'latex', insertionMode:'replaceAll', selectionMode:'after' });
      field.blur();
      return true;
    })()`);
    await sleep(180);
    await capture("08-cases.png", ".workspace", 0);

    // 09 — actual matrix tool category/custom matrix UI.
    await clickText('公式工具');
    await sleep(100);
    const matrixOpened = await clickText('矩阵');
    await sleep(180);
    assert.equal(matrixOpened, true, 'Matrix category could not be opened');
    await capture("09-matrix-tools.png", ".workspace", 0);

    await writeFile(new URL("capture-report.json", outputDir), JSON.stringify({ alignState, contextInfo, contextMenuState, shortcutOpened, behaviorOpened, matrixOpened, designerOpened, themeStudioOpened }, null, 2));
    console.log("VisualTeX help screenshots captured");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(180);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => {});
  }
}

await main();
