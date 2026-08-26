import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 650;
const vitePort = 23100 + offset;
const debugPort = 30100 + offset;
const baseUrl = `http://127.0.0.1:${vitePort}`;
const chromeProfile = `/tmp/visualtex-custom-symbol-designer-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      if ((await fetch(url)).ok) return;
    } catch {
      // Retry while local process starts.
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
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`CDP ${method} timed out`));
      }, 15000);
      this.pending.set(id, {
        resolve: (value) => {
          clearTimeout(timer);
          resolve(value);
        },
        reject: (error) => {
          clearTimeout(timer);
          reject(error);
        },
      });
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

async function waitUntil(client, expression, timeoutMs = 12000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const value = await client.evaluate(expression);
    if (value) return value;
    await sleep(60);
  }
  throw new Error(`Timed out waiting for ${expression}`);
}

async function setReactInput(client, selector, value) {
  await client.evaluate(`(() => {
    const input = document.querySelector(${JSON.stringify(selector)});
    if (!(input instanceof HTMLInputElement)) throw new Error(${JSON.stringify(`Input not found: ${selector}`)});
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
    setter?.call(input, ${JSON.stringify(String(value))});
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
    return input.value;
  })()`);
  await sleep(100);
}

async function setReactSelect(client, selector, value) {
  await client.evaluate(`(() => {
    const select = document.querySelector(${JSON.stringify(selector)});
    if (!select || select.tagName !== "SELECT") {
      throw new Error(${JSON.stringify(`Select not found: ${selector}`)} +
        " registrationPanel=" + Boolean(document.querySelector('[data-custom-symbol-registration-panel]')));
    }
    const setter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value")?.set;
    if (setter) setter.call(select, ${JSON.stringify(String(value))});
    else select.value = ${JSON.stringify(String(value))};
    select.dispatchEvent(new Event("input", { bubbles: true }));
    select.dispatchEvent(new Event("change", { bubbles: true }));
    return select.value;
  })()`);
  await sleep(100);
}

async function main() {
  const vite = spawn(
    process.execPath,
    [
      "node_modules/vite/bin/vite.js",
      "--host",
      "127.0.0.1",
      "--port",
      String(vitePort),
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
    assert.ok(page, "VisualTeX browser target must exist");
    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(450);

    await client.evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
      localStorage.setItem("visualtex.onboarding.macos.desktop.v1.2.0.completed", "true");
      localStorage.setItem("visualtex.office.macos.native-first-run.v1.2.0.completed", "true");
      localStorage.removeItem("visualtex.custom-symbols.v1");
      const key = "visualtex-editor";
      const persisted = JSON.parse(localStorage.getItem(key) || "{}");
      persisted.state = {
        ...(persisted.state || {}),
        checkUpdatesOnStartup: false,
        theme: "proof",
      };
      localStorage.setItem(key, JSON.stringify(persisted));
      return true;
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(550);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(550);
    await waitUntil(client, `Boolean(document.querySelector("math-field"))`);

    const baseline = await client.evaluate(`(() => ({
      formula: document.querySelector("math-field")?.value || "",
      storage: localStorage.getItem("visualtex.custom-symbols.v1"),
      theme: document.documentElement.dataset.theme || "",
    }))()`);
    assert.equal(baseline.theme, "proof");
    process.stdout.write("[custom-symbol-designer] Proof editor ready\n");

    await client.evaluate(`(() => {
      const tilesTab = document.querySelector('[data-toolbar-view="tiles"]');
      if (tilesTab instanceof HTMLButtonElement) tilesTab.click();
      return true;
    })()`);
    await sleep(100);
    await waitUntil(client, `Boolean(document.querySelector('[data-tile-category="custom"]'))`);
    await client.evaluate(`document.querySelector('[data-tile-category="custom"]').click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-open-custom-symbol-designer]'))`);
    await client.evaluate(`document.querySelector('[data-open-custom-symbol-designer]').click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-custom-symbol-designer]'))`);

    const themePurity = await client.evaluate(`(() => {
      const selectors = [
        ".custom-symbol-designer-dialog",
        ".custom-symbol-designer-sidebar.is-materials",
        ".custom-symbol-designer-panel",
        ".custom-symbol-designer-canvas",
      ];
      return selectors.map((selector) => {
        const element = document.querySelector(selector);
        return {
          selector,
          background: element ? getComputedStyle(element).backgroundColor : "missing",
        };
      });
    })()`);
    for (const entry of themePurity) {
      assert.notEqual(entry.background, "missing", entry.selector);
      assert.notEqual(
        entry.background,
        "rgb(255, 255, 255)",
        `${entry.selector} leaked a fixed white background under Proof`,
      );
    }
    process.stdout.write("[custom-symbol-designer] Proof theme purity verified\n");

    const redesignedSurface = await client.evaluate(`(() => {
      const materialLibrary = document.querySelector('[data-custom-symbol-material-library]');
      const categories = Array.from(
        document.querySelectorAll('[data-custom-symbol-material-category]'),
      ).map((item) => item.getAttribute('data-custom-symbol-material-category'));
      const fontSize = (selector) => {
        const element = document.querySelector(selector);
        return element ? Number.parseFloat(getComputedStyle(element).fontSize) : 0;
      };
      return {
        materialCount: Number(materialLibrary?.getAttribute('data-material-count') || 0),
        bareMaterialsOnly: materialLibrary?.getAttribute('data-bare-materials-only') || '',
        categories,
        hasCompositeFrac: Boolean(document.querySelector('[data-custom-symbol-material-command="frac"]')),
        hasCompositeSqrt: Boolean(document.querySelector('[data-custom-symbol-material-command="sqrt"]')),
        hasCompositeMatrix: Boolean(document.querySelector('[data-custom-symbol-material-command="matrix2"]')),
        hasFontWrapper: Boolean(document.querySelector('[data-custom-symbol-material-command="math-italic"]')),
        hasInfiniteCanvas: document.querySelector('[data-custom-symbol-canvas-shell]')?.getAttribute('data-custom-symbol-infinite-canvas') === 'true',
        autoInkBounds: document.querySelector('[data-custom-symbol-canvas-shell]')?.getAttribute('data-custom-symbol-auto-ink-bounds') === 'true',
        hasOutputBox: Boolean(document.querySelector('[data-custom-symbol-output-box]')),
        hasCanvasPaper: Boolean(document.querySelector('[data-custom-symbol-canvas-paper]')),
        explanatoryHints: document.querySelectorAll(
          '.custom-symbol-system-font-hint, .custom-symbol-metric-hint, .custom-symbol-effect-hint',
        ).length,
        headerFont: fontSize('.custom-symbol-designer-header strong'),
        panelFont: fontSize('.custom-symbol-designer-panel > header strong'),
        categoryFont: fontSize('[data-custom-symbol-material-category]'),
        inputFont: fontSize('[data-custom-symbol-material-input]'),
        undersizedVisibleText: (() => {
          const dialog = document.querySelector('[data-custom-symbol-designer]');
          if (!dialog) return [{ text: 'missing-dialog', size: 0 }];
          const walker = document.createTreeWalker(dialog, NodeFilter.SHOW_TEXT);
          const undersized = [];
          let node;
          while ((node = walker.nextNode())) {
            const text = node.textContent?.trim() || '';
            const parent = node.parentElement;
            if (!text || !parent) continue;
            if (parent.closest('.math-preview, math-field, svg, [aria-hidden="true"]')) continue;
            if (['SCRIPT', 'STYLE'].includes(parent.tagName)) continue;
            const style = getComputedStyle(parent);
            if (
              style.display === 'none' ||
              style.visibility === 'hidden' ||
              Number(style.opacity) === 0 ||
              parent.getClientRects().length === 0
            ) {
              continue;
            }
            const size = Number.parseFloat(style.fontSize);
            if (Number.isFinite(size) && size < 10) {
              undersized.push({ text: text.slice(0, 36), size, className: parent.className || parent.tagName });
            }
          }
          return undersized;
        })(),
        systemFontOptions: document.querySelectorAll('[data-custom-symbol-system-font-select] option').length,
        systemGlyphCategories: document.querySelectorAll('[data-custom-symbol-system-glyph-category]').length,
        systemGlyphCount: document.querySelectorAll('[data-custom-symbol-system-glyph]').length,
        hasCambria: Array.from(document.querySelectorAll('[data-custom-symbol-system-font-select] option')).some((option) => option.textContent?.includes('Cambria Math')),
        hasItalicMode: Boolean(document.querySelector('[data-custom-symbol-system-font-italic]')),
        hasUprightMode: Boolean(document.querySelector('[data-custom-symbol-system-font-upright]')),
      };
    })()`);
    assert.ok(redesignedSurface.materialCount > 20, "The bare-character library must still expose a substantial glyph set");
    assert.equal(redesignedSurface.bareMaterialsOnly, "true");
    assert.deepEqual(redesignedSurface.categories, [
      "common",
      "calculus",
      "greek",
      "relation",
      "set",
      "arrow",
      "physics",
    ]);
    assert.equal(redesignedSurface.hasCompositeFrac, false);
    assert.equal(redesignedSurface.hasCompositeSqrt, false);
    assert.equal(redesignedSurface.hasCompositeMatrix, false);
    assert.equal(redesignedSurface.hasFontWrapper, false);
    assert.equal(redesignedSurface.hasInfiniteCanvas, true);
    assert.equal(redesignedSurface.autoInkBounds, true);
    assert.equal(redesignedSurface.hasOutputBox, false);
    assert.equal(redesignedSurface.hasCanvasPaper, false);
    assert.equal(redesignedSurface.explanatoryHints, 0);
    assert.ok(redesignedSurface.headerFont >= 14);
    assert.ok(redesignedSurface.panelFont >= 11);
    assert.ok(redesignedSurface.categoryFont >= 10);
    assert.ok(redesignedSurface.inputFont >= 11);
    assert.deepEqual(
      redesignedSurface.undersizedVisibleText,
      [],
      `The redesigned character editor still contains undersized visible text: ${JSON.stringify(redesignedSurface.undersizedVisibleText)}`,
    );
    assert.ok(redesignedSurface.systemFontOptions >= 5, "The extended glyph library must expose multiple system math fonts");
    assert.ok(redesignedSurface.systemGlyphCategories >= 8, "The extended glyph library must be categorized");
    assert.ok(redesignedSurface.systemGlyphCount > 20, "The active extended glyph category must expose a substantial symbol set");
    assert.equal(redesignedSurface.hasCambria, true, "Cambria Math must be an explicit system-font source");
    assert.equal(redesignedSurface.hasItalicMode, true);
    assert.equal(redesignedSurface.hasUprightMode, true);

    const registrationPlaceholders = await client.evaluate(`(() => ({
      command: document.querySelector('[data-custom-symbol-command-input]')?.getAttribute('placeholder') || '',
      fallback: document.querySelector('[data-custom-symbol-omml-fallback-input]')?.getAttribute('placeholder') || '',
    }))()`);
    assert.equal(
      registrationPlaceholders.command,
      String.raw`\selfdefa`,
      `The command placeholder must display exactly one LaTeX backslash: ${JSON.stringify(registrationPlaceholders)}`,
    );
    assert.equal(
      registrationPlaceholders.fallback,
      String.raw`\approx`,
      `The OMML fallback placeholder must display exactly one LaTeX backslash: ${JSON.stringify(registrationPlaceholders)}`,
    );

    await client.evaluate(`document.querySelector('[data-custom-symbol-material-category="calculus"]').click()`);
    await waitUntil(
      client,
      `document.querySelectorAll('[data-custom-symbol-material-latex]').length >= 12 && Array.from(document.querySelectorAll('[data-custom-symbol-material-library] .math-preview')).every((preview) => preview.dataset.fitReady === 'true')`,
      30000,
    );
    const calculusTileFit = await client.evaluate(`(() => {
      const tiles = Array.from(document.querySelectorAll('[data-custom-symbol-material-latex]'));
      return tiles.map((tile) => {
        const preview = tile.querySelector('.math-preview-fit-content');
        const tileRect = tile.getBoundingClientRect();
        const previewRect = preview?.getBoundingClientRect();
        return {
          id: tile.getAttribute('data-custom-symbol-material-command') || '',
          latex: tile.getAttribute('data-custom-symbol-material-latex') || '',
          fitReady: tile.querySelector('.math-preview')?.dataset.fitReady || '',
          tile: { left: tileRect.left, top: tileRect.top, right: tileRect.right, bottom: tileRect.bottom },
          preview: previewRect
            ? { left: previewRect.left, top: previewRect.top, right: previewRect.right, bottom: previewRect.bottom, width: previewRect.width, height: previewRect.height }
            : null,
        };
      });
    })()`);
    assert.ok(calculusTileFit.length >= 12);
    for (const item of calculusTileFit) {
      assert.equal(item.fitReady, 'true', `Material preview did not finish fitting: ${JSON.stringify(item)}`);
      assert.ok(item.preview && item.preview.width > 2 && item.preview.height > 2, `Material glyph is not visible: ${JSON.stringify(item)}`);
      assert.ok(item.preview.left >= item.tile.left + 2, `Material glyph is clipped on the left: ${JSON.stringify(item)}`);
      assert.ok(item.preview.right <= item.tile.right - 2, `Material glyph is clipped on the right: ${JSON.stringify(item)}`);
      assert.ok(item.preview.top >= item.tile.top + 2, `Material glyph is clipped on the top: ${JSON.stringify(item)}`);
      assert.ok(item.preview.bottom <= item.tile.bottom - 2, `Material glyph is clipped on the bottom: ${JSON.stringify(item)}`);
    }
    process.stdout.write("[custom-symbol-designer] registration placeholders and large-operator tile fitting verified\n");

    await client.evaluate(`document.querySelector('[data-custom-symbol-material-category="relation"]').click()`);
    await waitUntil(
      client,
      `document.querySelectorAll('[data-custom-symbol-material-latex]').length >= 20`,
    );
    const designerContrast = await client.evaluate(`(() => {
      const canvas = document.querySelector('[data-custom-symbol-canvas-shell]');
      const materialLibrary = document.querySelector('[data-custom-symbol-material-library]');
      const firstMaterial = document.querySelector('[data-custom-symbol-material-latex]');
      const firstPreview = firstMaterial?.querySelector('.math-preview');
      const equalsTile = document.querySelector('[data-custom-symbol-material-command="equal"]');
      const approxTile = document.querySelector('[data-custom-symbol-material-command="approx"]');
      const visibleGlyph = (tile) => {
        const preview = tile?.querySelector('.math-preview');
        const latex = preview?.querySelector('.ML__latex');
        const candidates = Array.from(latex?.querySelectorAll('*') || []).filter((element) => {
          const text = element.textContent?.trim() || '';
          const rect = element.getBoundingClientRect();
          const style = getComputedStyle(element);
          return Boolean(text) && rect.width > 1 && rect.height > 1 && style.visibility !== 'hidden' && Number(style.opacity) > 0;
        });
        const largest = candidates.sort((left, right) => {
          const a = left.getBoundingClientRect();
          const b = right.getBoundingClientRect();
          return b.width * b.height - a.width * a.height;
        })[0] || latex;
        const rect = largest?.getBoundingClientRect();
        const style = largest ? getComputedStyle(largest) : null;
        return {
          text: latex?.textContent?.trim() || '',
          width: rect?.width || 0,
          height: rect?.height || 0,
          color: style?.color || '',
          opacity: style?.opacity || '',
          visibility: style?.visibility || '',
          display: style?.display || '',
          fontSize: style?.fontSize || '',
          lineHeight: style?.lineHeight || '',
          previewDisplay: preview ? getComputedStyle(preview).display : '',
          previewWidth: preview?.getBoundingClientRect().width || 0,
          previewHeight: preview?.getBoundingClientRect().height || 0,
          htmlLength: preview?.innerHTML.length || 0,
          html: preview?.innerHTML.slice(0, 500) || '',
        };
      };
      const styleOf = (element) => element ? getComputedStyle(element) : null;
      const canvasStyle = styleOf(canvas);
      const libraryStyle = styleOf(materialLibrary);
      const materialStyle = styleOf(firstMaterial);
      const previewStyle = styleOf(firstPreview);
      return {
        relationCount: document.querySelectorAll('[data-custom-symbol-material-latex]').length,
        hasEquals: Boolean(document.querySelector('[data-custom-symbol-material-command="equal"]')),
        hasApprox: Boolean(document.querySelector('[data-custom-symbol-material-command="approx"]')),
        canvasBackground: canvasStyle?.backgroundColor || '',
        canvasColor: canvasStyle?.color || '',
        libraryBackground: libraryStyle?.backgroundColor || '',
        materialBackground: materialStyle?.backgroundColor || '',
        materialBorderColor: materialStyle?.borderTopColor || '',
        materialColor: materialStyle?.color || '',
        previewColor: previewStyle?.color || '',
        equalsGlyph: visibleGlyph(equalsTile),
        approxGlyph: visibleGlyph(approxTile),
      };
    })()`);
    assert.ok(
      designerContrast.relationCount >= 20,
      `Relation character palette unexpectedly empty: ${JSON.stringify(designerContrast)}`,
    );
    assert.equal(designerContrast.hasEquals, true);
    assert.equal(designerContrast.hasApprox, true);
    assert.notEqual(designerContrast.canvasBackground, 'rgba(0, 0, 0, 0)');
    assert.notEqual(designerContrast.canvasBackground, designerContrast.canvasColor);
    assert.notEqual(designerContrast.libraryBackground, 'rgba(0, 0, 0, 0)');
    assert.notEqual(designerContrast.materialBackground, 'rgba(0, 0, 0, 0)');
    assert.notEqual(designerContrast.materialBorderColor, 'rgba(0, 0, 0, 0)');
    assert.notEqual(designerContrast.materialBackground, designerContrast.materialColor);
    assert.equal(designerContrast.previewColor, designerContrast.materialColor);
    assert.match(designerContrast.equalsGlyph.text, /=/);
    assert.match(designerContrast.approxGlyph.text, /≈/);
    for (const glyph of [designerContrast.equalsGlyph, designerContrast.approxGlyph]) {
      assert.ok(
        glyph.htmlLength > 20 && glyph.width >= 8 && glyph.height >= 8,
        `Character tile markup exists but its glyph is not visibly rendered: ${JSON.stringify(glyph)}`,
      );
      assert.equal(glyph.visibility, 'visible');
      assert.notEqual(glyph.opacity, '0');
      assert.equal(glyph.color, designerContrast.materialColor);
    }
    await client.evaluate(`document.querySelector('[data-custom-symbol-material-category="common"]').click()`);
    await waitUntil(
      client,
      `document.querySelector('[data-custom-symbol-material-category="common"]')?.classList.contains('is-active') === true`,
    );
    process.stdout.write("[custom-symbol-designer] readable controls, visible symbol palette and theme-safe canvas verified\n");

    await client.evaluate(`document.querySelector('[data-custom-symbol-system-glyph-category="math-alphanumeric"]').click()`);
    await waitUntil(
      client,
      `document.querySelectorAll('[data-custom-symbol-system-glyph]').length > 900`,
      30000,
    );
    const fullMathematicalAlphabet = await client.evaluate(`(() => ({
      count: document.querySelectorAll('[data-custom-symbol-system-glyph]').length,
      hasLastCodePoint: Boolean(document.querySelector('[data-custom-symbol-system-glyph="U+1D7FF"]')),
    }))()`);
    assert.ok(
      fullMathematicalAlphabet.count > 900,
      `The mathematical alphanumeric browser must expose the complete Unicode block rather than a short prefix: ${JSON.stringify(fullMathematicalAlphabet)}`,
    );
    assert.equal(
      fullMathematicalAlphabet.hasLastCodePoint,
      true,
      "The extended glyph browser must reach U+1D7FF",
    );
    await client.evaluate(`document.querySelector('[data-custom-symbol-system-glyph-category="basic-italic"]').click()`);
    await waitUntil(
      client,
      `Boolean(document.querySelector('[data-custom-symbol-system-glyph="U+0041"]'))`,
    );
    process.stdout.write("[custom-symbol-designer] complete mathematical alphanumeric Unicode block verified\n");

    await client.evaluate(`document.querySelector('[data-custom-symbol-material-category="greek"]').click()`);
    await setReactInput(client, '[data-custom-symbol-material-search]', "approx");
    await waitUntil(
      client,
      `Boolean(document.querySelector('[data-custom-symbol-material-command="approx"]'))`,
    );
    const crossCategorySearch = await client.evaluate(`(() => ({
      activeCategory: document.querySelector('[data-custom-symbol-material-category].is-active')?.getAttribute('data-custom-symbol-material-category') || '',
      foundApprox: Boolean(document.querySelector('[data-custom-symbol-material-command="approx"]')),
      foundCompositeFrac: Boolean(document.querySelector('[data-custom-symbol-material-command="frac"]')),
      visibleSources: Array.from(document.querySelectorAll('[data-custom-symbol-material-latex]')).map(
        (item) => item.getAttribute('data-custom-symbol-material-latex') || '',
      ),
    }))()`);
    assert.equal(crossCategorySearch.activeCategory, "greek");
    assert.equal(
      crossCategorySearch.foundApprox,
      true,
      "A non-empty material search must search bare glyphs across categories",
    );
    assert.equal(crossCategorySearch.foundCompositeFrac, false);
    assert.ok(
      crossCategorySearch.visibleSources.every(
        (source) =>
          (Array.from(source).length === 1 || source.startsWith("\\")) &&
          !/[{}_^&]/.test(source) &&
          !source.includes("\\begin") &&
          !source.includes("\\placeholder"),
      ),
      `The character material browser leaked a composite formula: ${JSON.stringify(crossCategorySearch.visibleSources)}`,
    );
    await setReactInput(client, '[data-custom-symbol-material-search]', "frac");
    await sleep(100);
    assert.equal(
      await client.evaluate(`document.querySelectorAll('[data-custom-symbol-material-latex]').length`),
      0,
      "Composite fraction structures must not appear in the character material browser",
    );
    await setReactInput(client, '[data-custom-symbol-material-search]', "");
    await client.evaluate(`document.querySelector('[data-custom-symbol-material-category="common"]').click()`);
    await waitUntil(
      client,
      `document.querySelector('[data-custom-symbol-material-category="common"]')?.classList.contains('is-active') === true`,
    );
    process.stdout.write("[custom-symbol-designer] cross-category bare-character search and composite-command exclusion verified\n");

    await setReactInput(
      client,
      '[data-custom-symbol-system-glyph-search]',
      "U+0041",
    );
    await waitUntil(
      client,
      `Boolean(document.querySelector('[data-custom-symbol-system-glyph="U+0041"]'))`,
    );
    await client.evaluate(`document.querySelector('[data-custom-symbol-system-glyph="U+0041"]').click()`);
    await waitUntil(
      client,
      `document.querySelectorAll('[data-custom-symbol-layer]').length === 1 && !document.querySelector('[data-custom-symbol-system-glyph][aria-busy="true"]')`,
    );
    const browserSystemGlyph = await client.evaluate(`(() => {
      const selected = document.querySelector('[data-custom-symbol-layer].is-selected');
      const id = selected?.getAttribute('data-custom-symbol-layer') || '';
      const artwork = id
        ? document.querySelector('[data-custom-symbol-artwork-layer="' + id + '"]')
        : null;
      const text = artwork?.querySelector('text');
      return {
        source: selected?.getAttribute('data-layer-source-latex') || '',
        hasTextFallback: Boolean(text),
        fontFamily: text?.getAttribute('font-family') || '',
        fontStyle: text?.getAttribute('font-style') || '',
        status: document.querySelector('[data-custom-symbol-system-font-status]')?.textContent || '',
      };
    })()`);
    assert.equal(browserSystemGlyph.source, "A");
    assert.equal(browserSystemGlyph.hasTextFallback, true);
    assert.match(browserSystemGlyph.fontFamily, /Cambria Math/);
    assert.equal(browserSystemGlyph.fontStyle, "italic");
    assert.match(browserSystemGlyph.status, /browser|浏览器/i);
    await client.evaluate(`document.querySelector('[data-custom-symbol-system-font-upright]').click()`);
    await waitUntil(
      client,
      `document.querySelector('[data-custom-symbol-system-font-upright]')?.getAttribute('aria-pressed') === 'true'`,
    );
    await client.evaluate(`document.querySelector('[data-custom-symbol-system-font-italic]').click()`);
    await waitUntil(
      client,
      `document.querySelector('[data-custom-symbol-system-font-italic]')?.getAttribute('aria-pressed') === 'true'`,
    );
    await client.evaluate(`document.querySelector('[data-delete-custom-symbol-layer]').click()`);
    await waitUntil(
      client,
      `document.querySelectorAll('[data-custom-symbol-layer]').length === 0`,
    );
    await setReactInput(client, '[data-custom-symbol-system-glyph-search]', "");
    process.stdout.write("[custom-symbol-designer] system math glyph insertion and browser fallback verified\n");

    const beforeFirstInsertion = await client.evaluate(`(() => {
      const canvas = document.querySelector('[data-custom-symbol-canvas]');
      const reference = document.querySelector('[data-custom-symbol-reference]');
      const rect = reference?.getBoundingClientRect();
      return {
        viewBox: canvas?.getAttribute('viewBox') || '',
        referenceCenter: rect
          ? { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 }
          : null,
      };
    })()`);
    assert.ok(beforeFirstInsertion.referenceCenter);

    await client.evaluate(`document.querySelector('[data-add-custom-symbol-material]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 1`);
    await waitUntil(
      client,
      `document.querySelector('[data-custom-symbol-ink-size]')?.textContent?.trim()?.endsWith('em') === true`,
    );
    const first = await client.evaluate(`(() => {
      const layer = document.querySelector('[data-custom-symbol-layer]');
      const canvasLayer = document.querySelector('[data-custom-symbol-canvas-layer]');
      const artworkLayer = document.querySelector('[data-custom-symbol-artwork-layer]');
      const canvas = document.querySelector('[data-custom-symbol-canvas]');
      const shell = document.querySelector('[data-custom-symbol-canvas-shell]');
      const artworkRect = artworkLayer?.getBoundingClientRect();
      const shellRect = shell?.getBoundingClientRect();
      const inkSize = document.querySelector('[data-custom-symbol-ink-size]')?.textContent?.trim() || '';
      return {
        layerId: layer?.getAttribute("data-custom-symbol-layer") || "",
        hasCanvasLayer: Boolean(canvasLayer),
        hasPath: Boolean(artworkLayer?.querySelector("path")),
        hasBaseline: Boolean(document.querySelector('[data-custom-symbol-baseline]')),
        hasReferenceAlpha: Boolean(document.querySelector('[data-custom-symbol-reference-alpha]')),
        hasFitControl: Boolean(document.querySelector('[data-custom-symbol-fit-view]')),
        infiniteCanvas: canvas?.getAttribute('data-custom-symbol-infinite-canvas') || '',
        shellInfiniteCanvas: shell?.getAttribute('data-custom-symbol-infinite-canvas') || '',
        hasPaper: Boolean(document.querySelector('[data-custom-symbol-canvas-paper]')),
        hasOutputBox: Boolean(document.querySelector('[data-custom-symbol-output-box]')),
        hasWorkspaceButton: Boolean(document.querySelector('[data-custom-symbol-fit-workspace]')),
        gridDots: document.querySelectorAll('.custom-symbol-designer-grid-dot').length,
        gridPaths: document.querySelectorAll('.custom-symbol-designer-grid-line').length,
        rotationHandles: document.querySelectorAll('[data-custom-symbol-rotation-handle]').length,
        rotationHitTargets: document.querySelectorAll('[data-custom-symbol-rotation-hit-target]').length,
        inkSize,
        viewBox: canvas?.getAttribute("viewBox") || "",
        artworkRect: artworkRect
          ? { left: artworkRect.left, top: artworkRect.top, right: artworkRect.right, bottom: artworkRect.bottom }
          : null,
        shellRect: shellRect
          ? { left: shellRect.left, top: shellRect.top, right: shellRect.right, bottom: shellRect.bottom }
          : null,
      };
    })()`);
    assert.ok(first.layerId);
    assert.equal(first.hasCanvasLayer, true);
    assert.equal(first.hasPath, true);
    assert.equal(first.hasBaseline, true);
    assert.equal(first.hasReferenceAlpha, true, "Reference alpha should be visible by default");
    assert.equal(first.hasFitControl, true);
    assert.equal(first.infiniteCanvas, "true");
    assert.equal(first.shellInfiniteCanvas, "true");
    assert.equal(first.hasPaper, false);
    assert.equal(first.hasOutputBox, false);
    assert.equal(first.hasWorkspaceButton, false);
    assert.equal(first.gridDots, 0);
    assert.equal(first.gridPaths, 2);
    assert.equal(first.rotationHandles, 1);
    assert.equal(first.rotationHitTargets, 1);
    assert.match(first.inkSize, /em$/);
    assert.notEqual(first.inkSize, "—");
    const firstViewBox = first.viewBox.split(/\s+/).map(Number);
    assert.equal(firstViewBox.length, 4);
    assert.ok(firstViewBox.every(Number.isFinite));
    assert.ok(firstViewBox[2] > 0 && firstViewBox[3] > 0);
    assert.ok(first.artworkRect && first.shellRect);
    assert.ok(first.artworkRect.left >= first.shellRect.left - 2);
    assert.ok(first.artworkRect.top >= first.shellRect.top - 2);
    assert.ok(first.artworkRect.right <= first.shellRect.right + 2);
    assert.ok(first.artworkRect.bottom <= first.shellRect.bottom + 2);

    const afterFirstInsertion = await client.evaluate(`(() => {
      const canvas = document.querySelector('[data-custom-symbol-canvas]');
      const reference = document.querySelector('[data-custom-symbol-reference]');
      const rect = reference?.getBoundingClientRect();
      return {
        viewBox: canvas?.getAttribute('viewBox') || '',
        referenceCenter: rect
          ? { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 }
          : null,
      };
    })()`);
    assert.equal(
      afterFirstInsertion.viewBox,
      beforeFirstInsertion.viewBox,
      "Inserting the first character must not implicitly reframe the canvas viewport",
    );
    assert.ok(afterFirstInsertion.referenceCenter);
    assert.ok(
      Math.hypot(
        afterFirstInsertion.referenceCenter.x - beforeFirstInsertion.referenceCenter.x,
        afterFirstInsertion.referenceCenter.y - beforeFirstInsertion.referenceCenter.y,
      ) < 1,
      `The baseline reference must stay visually stationary on first insertion: ${JSON.stringify({ beforeFirstInsertion, afterFirstInsertion })}`,
    );
    process.stdout.write("[custom-symbol-designer] first insertion preserves viewport and baseline reference position\n");

    const backspaceTarget = await client.evaluate(`(() => {
      const layer = document.querySelector('[data-custom-symbol-canvas-layer].is-selected');
      if (!layer) return null;
      const rect = layer.getBoundingClientRect();
      return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 };
    })()`);
    assert.ok(backspaceTarget);
    await client.send("Input.dispatchMouseEvent", {
      type: "mousePressed",
      x: backspaceTarget.x,
      y: backspaceTarget.y,
      button: "left",
      buttons: 1,
      clickCount: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseReleased",
      x: backspaceTarget.x,
      y: backspaceTarget.y,
      button: "left",
      buttons: 0,
      clickCount: 1,
    });
    const canvasFocused = await client.evaluate(
      `document.activeElement === document.querySelector('[data-custom-symbol-canvas]')`,
    );
    assert.equal(canvasFocused, true, "Clicking a character should focus the design canvas for keyboard editing");
    await client.send("Input.dispatchKeyEvent", {
      type: "keyDown",
      key: "Backspace",
      code: "Backspace",
      windowsVirtualKeyCode: 8,
      nativeVirtualKeyCode: 8,
    });
    await client.send("Input.dispatchKeyEvent", {
      type: "keyUp",
      key: "Backspace",
      code: "Backspace",
      windowsVirtualKeyCode: 8,
      nativeVirtualKeyCode: 8,
    });
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 0`);
    await client.evaluate(`document.querySelector('[data-add-custom-symbol-material]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 1`);
    process.stdout.write("[custom-symbol-designer] selected canvas layer Backspace deletion verified\n");

    await client.evaluate(`document.querySelector('[data-toggle-custom-symbol-reference-alpha]').click()`);
    await waitUntil(client, `!document.querySelector('[data-custom-symbol-reference-alpha]')`);
    await client.evaluate(`document.querySelector('[data-toggle-custom-symbol-reference-alpha]').click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-custom-symbol-reference-alpha]'))`);

    const zoomBefore = await client.evaluate(`document.querySelector('[data-custom-symbol-canvas]')?.getAttribute('viewBox') || ''`);
    await client.evaluate(`document.querySelector('[data-custom-symbol-zoom-in]').click()`);
    await sleep(80);
    const zoomAfter = await client.evaluate(`document.querySelector('[data-custom-symbol-canvas]')?.getAttribute('viewBox') || ''`);
    assert.notEqual(zoomAfter, zoomBefore, "Zoom-in control must change the mathematical viewport");
    await client.evaluate(`document.querySelector('[data-custom-symbol-fit-view]').click()`);
    await sleep(80);

    const workspaceState = await client.evaluate(`(() => {
      const canvas = document.querySelector('[data-custom-symbol-canvas]');
      const workspace = (canvas?.getAttribute('data-custom-symbol-workspace') || '').split(/\\s+/).map(Number);
      const minorGrid = document.querySelector('.custom-symbol-designer-grid-layer.is-minor');
      const majorGrid = document.querySelector('.custom-symbol-designer-grid-layer.is-major');
      return {
        workspace,
        hasOutputBox: Boolean(document.querySelector('[data-custom-symbol-output-box]')),
        hasManualCanvasMetrics: Boolean(
          document.querySelector('[data-designer-field="canvas-width"], [data-designer-field="canvas-ascent"], [data-designer-field="canvas-descent"]'),
        ),
        minorOpacity: minorGrid ? Number.parseFloat(getComputedStyle(minorGrid).opacity) : 1,
        majorOpacity: majorGrid ? Number.parseFloat(getComputedStyle(majorGrid).opacity) : 1,
        alphaRect: (() => {
          const reference = document.querySelector('[data-custom-symbol-reference]');
          if (!reference) return null;
          const rect = reference.getBoundingClientRect();
          return { width: rect.width, height: rect.height };
        })(),
      };
    })()`);
    assert.equal(workspaceState.workspace.length, 4);
    assert.ok(workspaceState.workspace[2] >= 1_000_000);
    assert.ok(workspaceState.workspace[3] >= 1_000_000);
    assert.equal(workspaceState.hasOutputBox, false);
    assert.equal(workspaceState.hasManualCanvasMetrics, false);
    assert.ok(workspaceState.minorOpacity >= 0.18 && workspaceState.minorOpacity <= 0.24);
    assert.ok(workspaceState.majorOpacity >= 0.28 && workspaceState.majorOpacity <= 0.36);
    process.stdout.write("[custom-symbol-designer] borderless infinite workspace and softened grid verified\n");

    await setReactSelect(client, '[data-custom-symbol-reference-select]', String.raw`\displaystyle\sum`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-reference]')?.getAttribute('data-custom-symbol-reference-label') === 'Σ'`);
    const sumReference = await client.evaluate(`(() => {
      const reference = document.querySelector('[data-custom-symbol-reference]');
      if (!reference) return null;
      const rect = reference.getBoundingClientRect();
      return { width: rect.width, height: rect.height };
    })()`);
    assert.ok(sumReference && workspaceState.alphaRect);
    assert.ok(
      sumReference.height > workspaceState.alphaRect.height,
      "Large-operator reference should visibly differ from ordinary alpha size",
    );
    const referenceFitCases = [
      [String.raw`\displaystyle\int`, "∫"],
      [String.raw`\displaystyle\oint`, "∮"],
      [String.raw`\displaystyle\sum`, "Σ"],
      [String.raw`\displaystyle\prod`, "Π"],
      [String.raw`\displaystyle\bigcup`, "⋃"],
    ];
    for (const [latex, label] of referenceFitCases) {
      await setReactSelect(client, '[data-custom-symbol-reference-select]', latex);
      await waitUntil(
        client,
        `document.querySelector('[data-custom-symbol-reference]')?.getAttribute('data-custom-symbol-reference-label') === ${JSON.stringify(label)}`,
      );
      await client.evaluate(`document.querySelector('[data-custom-symbol-fit-view]').click()`);
      await sleep(70);
      const referenceState = await client.evaluate(`(() => {
        const reference = document.querySelector('[data-custom-symbol-reference]');
        if (!reference) return null;
        const rect = reference.getBoundingClientRect();
        return {
          label: reference.getAttribute('data-custom-symbol-reference-label') || '',
          width: rect.width,
          height: rect.height,
          visible: getComputedStyle(reference).visibility !== 'hidden',
        };
      })()`);
      assert.ok(referenceState, `${label} reference must exist`);
      assert.equal(referenceState.label, label);
      assert.ok(referenceState.width > 0 && referenceState.height > 0);
      assert.equal(referenceState.visible, true);
    }

    await setReactSelect(client, '[data-custom-symbol-reference-select]', String.raw`\displaystyle\int`);
    await setReactInput(client, '[data-custom-symbol-material-input]', String.raw`\int`);
    await client.evaluate(`document.querySelector('[data-add-custom-symbol-material]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 2`);
    const integralConsistency = await client.evaluate(`(() => {
      const selected = document.querySelector('[data-custom-symbol-layer].is-selected');
      const artwork = selected
        ? document.querySelector('[data-custom-symbol-artwork-layer="' + selected.getAttribute('data-custom-symbol-layer') + '"]')
        : null;
      const reference = document.querySelector('[data-custom-symbol-reference]');
      return {
        source: selected?.getAttribute('data-layer-source-latex') || '',
        artworkPaths: Array.from(artwork?.querySelectorAll('path') || []).map((path) => path.getAttribute('d') || ''),
        referencePaths: Array.from(reference?.querySelectorAll('path') || []).map((path) => path.getAttribute('d') || ''),
      };
    })()`);
    assert.equal(integralConsistency.source, String.raw`\displaystyle\int`);
    assert.deepEqual(
      integralConsistency.artworkPaths,
      integralConsistency.referencePaths,
      "A quick/material \\int must use the exact same displaystyle MathJax glyph outline as the integral reference",
    );
    await client.evaluate(`document.querySelector('[data-delete-custom-symbol-layer]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 1`);
    await client.evaluate(`document.querySelector('[data-custom-symbol-layer]').click()`);
    process.stdout.write("[custom-symbol-designer] material/reference integral outline consistency verified\n");

    await setReactSelect(client, '[data-custom-symbol-reference-select]', String.raw`\alpha`);
    await waitUntil(client, `Boolean(document.querySelector('[data-custom-symbol-reference-alpha]'))`);
    const automaticBounds = await client.evaluate(`(() => ({
      inkSize: document.querySelector('[data-custom-symbol-ink-size]')?.textContent?.trim() || '',
      manualWidth: Boolean(document.querySelector('[data-designer-field="canvas-width"]')),
      manualAscent: Boolean(document.querySelector('[data-designer-field="canvas-ascent"]')),
      manualDescent: Boolean(document.querySelector('[data-designer-field="canvas-descent"]')),
    }))()`);
    assert.match(automaticBounds.inkSize, /em$/);
    assert.equal(automaticBounds.manualWidth, false);
    assert.equal(automaticBounds.manualAscent, false);
    assert.equal(automaticBounds.manualDescent, false);
    process.stdout.write("[custom-symbol-designer] automatic ink bounds, viewport locate and reference glyphs verified\n");

    await setReactInput(client, '[data-designer-field="layer-x"]', 80);
    await setReactInput(client, '[data-designer-field="layer-scale-x"]', 1.25);
    await setReactInput(client, '[data-designer-field="layer-rotation"]', 12);
    const transformed = await client.evaluate(`(() => ({
      transform: document.querySelector('[data-custom-symbol-canvas-layer]')?.getAttribute("transform") || "",
      x: document.querySelector('[data-designer-field="layer-x"]')?.value || "",
      scaleX: document.querySelector('[data-designer-field="layer-scale-x"]')?.value || "",
      rotation: document.querySelector('[data-designer-field="layer-rotation"]')?.value || "",
    }))()`);
    assert.match(transformed.transform, /translate\(80 /);
    assert.match(transformed.transform, /rotate\(12\)/);
    assert.match(transformed.transform, /scale\(1\.25 /);

    const advancedControls = await client.evaluate(`(() => ({
      flipHorizontal: Boolean(document.querySelector('[data-flip-custom-symbol-layer="horizontal"]')),
      flipVertical: Boolean(document.querySelector('[data-flip-custom-symbol-layer="vertical"]')),
      skewX: Boolean(document.querySelector('[data-designer-field="layer-skew-x"]')),
      skewY: Boolean(document.querySelector('[data-designer-field="layer-skew-y"]')),
      outline: Boolean(document.querySelector('[data-custom-symbol-outline-toggle]')),
      perspective: Boolean(document.querySelector('[data-custom-symbol-perspective-toggle]')),
    }))()`);
    assert.deepEqual(advancedControls, {
      flipHorizontal: true,
      flipVertical: true,
      skewX: true,
      skewY: true,
      outline: true,
      perspective: true,
    });

    await client.evaluate(`document.querySelector('[data-flip-custom-symbol-layer="horizontal"]').click()`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-canvas-layer]')?.getAttribute('transform')?.includes('scale(-1.25')`);
    await client.evaluate(`document.querySelector('[data-flip-custom-symbol-layer="horizontal"]').click()`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-canvas-layer]')?.getAttribute('transform')?.includes('scale(1.25')`);
    await client.evaluate(`document.querySelector('[data-custom-symbol-math-italic]').click()`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-canvas-layer]')?.getAttribute('transform')?.includes('skewX(-12)')`);
    await client.evaluate(`document.querySelector('[data-custom-symbol-original-slant]').click()`);
    await waitUntil(client, `!document.querySelector('[data-custom-symbol-canvas-layer]')?.getAttribute('transform')?.includes('skewX(')`);

    await client.evaluate(`document.querySelector('[data-custom-symbol-outline-toggle]').click()`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-artwork-layer] path')?.getAttribute('fill') === 'none'`);
    const outlineState = await client.evaluate(`(() => {
      const path = document.querySelector('[data-custom-symbol-artwork-layer] path');
      return {
        fill: path?.getAttribute('fill') || '',
        stroke: path?.getAttribute('stroke') || '',
        width: Number(path?.getAttribute('stroke-width') || 0),
      };
    })()`);
    assert.equal(outlineState.fill, "none");
    assert.notEqual(outlineState.stroke, "none");
    assert.ok(outlineState.width > 0);

    const shapeCountBeforePerspective = await client.evaluate(`document.querySelectorAll('[data-custom-symbol-artwork-layer] path').length`);
    await client.evaluate(`document.querySelector('[data-custom-symbol-perspective-toggle]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-artwork-layer] path').length > ${shapeCountBeforePerspective}`);
    const perspectiveState = await client.evaluate(`(() => ({
      pathCount: document.querySelectorAll('[data-custom-symbol-artwork-layer] path').length,
      depthField: Boolean(document.querySelector('[data-designer-field="perspective-depth"]')),
      angleField: Boolean(document.querySelector('[data-designer-field="perspective-angle"]')),
      stepsField: Boolean(document.querySelector('[data-designer-field="perspective-steps"]')),
    }))()`);
    assert.ok(perspectiveState.pathCount > shapeCountBeforePerspective);
    assert.equal(perspectiveState.depthField, true);
    assert.equal(perspectiveState.angleField, true);
    assert.equal(perspectiveState.stepsField, true);
    await client.evaluate(`document.querySelector('[data-custom-symbol-perspective-toggle]').click()`);
    await client.evaluate(`document.querySelector('[data-custom-symbol-outline-toggle]').click()`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-artwork-layer] path')?.getAttribute('fill') !== 'none'`);
    process.stdout.write("[custom-symbol-designer] numeric transforms, flips, slant, hollow outline and perspective verified\n");

    await client.evaluate(`document.querySelector('[data-custom-symbol-fit-view]').click()`);
    await sleep(80);

    const scaleBeforeHandle = await client.evaluate(`(() => ({
      x: Number(document.querySelector('[data-designer-field="layer-scale-x"]')?.value || 1),
      y: Number(document.querySelector('[data-designer-field="layer-scale-y"]')?.value || 1),
      rotation: Number(document.querySelector('[data-designer-field="layer-rotation"]')?.value || 0),
      handles: document.querySelectorAll('[data-custom-symbol-resize-handle]').length,
      hitTargets: document.querySelectorAll('[data-custom-symbol-resize-hit-target]').length,
      rotationHandles: document.querySelectorAll('[data-custom-symbol-rotation-handle]').length,
      rotationHitTargets: document.querySelectorAll('[data-custom-symbol-rotation-hit-target]').length,
    }))()`);
    assert.equal(scaleBeforeHandle.handles, 8, "Selected layer should expose eight visual resize handles");
    assert.equal(scaleBeforeHandle.hitTargets, 8, "Selected layer should expose eight forgiving resize hit targets");
    assert.equal(scaleBeforeHandle.rotationHandles, 1, "Selected layer should expose a direct rotation handle");
    assert.equal(scaleBeforeHandle.rotationHitTargets, 1, "The rotation handle should have a forgiving hit target");
    const dragPriority = await client.evaluate(`(() => {
      const selection = document.querySelector('[data-custom-symbol-canvas-layer].is-selected .custom-symbol-designer-selection-box');
      if (!selection) return null;
      const rect = selection.getBoundingClientRect();
      const x = rect.left + rect.width / 2;
      const y = rect.top + rect.height / 2;
      const coveringResizeHandles = Array.from(
        document.querySelectorAll('[data-custom-symbol-resize-hit-target]'),
      ).filter((target) => {
        const hit = target.getBoundingClientRect();
        return x >= hit.left && x <= hit.right && y >= hit.top && y <= hit.bottom;
      }).map((target) => target.getAttribute('data-custom-symbol-resize-hit-target'));
      return { width: rect.width, height: rect.height, coveringResizeHandles };
    })()`);
    assert.ok(dragPriority, "Selected layer must expose a draggable interior");
    if (dragPriority.width > 20 && dragPriority.height > 20) {
      assert.deepEqual(
        dragPriority.coveringResizeHandles,
        [],
        `Resize hit targets must not cover the center drag area: ${JSON.stringify(dragPriority)}`,
      );
    }
    const resizeBox = await client.evaluate(`(() => {
      const handle = document.querySelector('[data-custom-symbol-resize-handle="se"]');
      const hitTarget = document.querySelector('[data-custom-symbol-resize-hit-target="se"]');
      const anchor = document.querySelector('[data-custom-symbol-resize-handle="nw"]');
      if (!handle || !hitTarget || !anchor) return null;
      const handleRect = handle.getBoundingClientRect();
      const hitRect = hitTarget.getBoundingClientRect();
      const anchorRect = anchor.getBoundingClientRect();
      const x = handleRect.right + Math.max(1, (hitRect.right - handleRect.right) * 0.5);
      const y = handleRect.top + handleRect.height / 2;
      return {
        x,
        y,
        anchorX: anchorRect.left + anchorRect.width / 2,
        anchorY: anchorRect.top + anchorRect.height / 2,
        outsideVisualHandle: x > handleRect.right,
        insideHitTarget:
          x >= hitRect.left && x <= hitRect.right && y >= hitRect.top && y <= hitRect.bottom,
        visualWidth: handleRect.width,
        hitWidth: hitRect.width,
      };
    })()`);
    assert.ok(resizeBox, "Bottom-right resize handle must have a forgiving hit target");
    assert.equal(resizeBox.outsideVisualHandle, true, "Resize regression must begin outside the painted handle");
    assert.equal(resizeBox.insideHitTarget, true, "Resize regression must begin inside the expanded hit target");
    assert.ok(
      resizeBox.hitWidth >= resizeBox.visualWidth,
      "Resize hit target must remain at least as large as the painted handle",
    );
    assert.ok(
      resizeBox.hitWidth <= 15.5,
      `Resize hit target must stay compact enough for easy dragging: ${JSON.stringify(resizeBox)}`,
    );
    await client.send("Input.dispatchMouseEvent", {
      type: "mousePressed",
      x: resizeBox.x,
      y: resizeBox.y,
      button: "left",
      buttons: 1,
      clickCount: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseMoved",
      x: resizeBox.x + 35,
      y: resizeBox.y + 30,
      button: "left",
      buttons: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseReleased",
      x: resizeBox.x + 35,
      y: resizeBox.y + 30,
      button: "left",
      buttons: 0,
      clickCount: 1,
    });
    await sleep(100);
    const scaleAfterHandle = await client.evaluate(`(() => {
      const anchor = document.querySelector('[data-custom-symbol-resize-handle="nw"]');
      const anchorRect = anchor?.getBoundingClientRect();
      return {
        x: Number(document.querySelector('[data-designer-field="layer-scale-x"]')?.value || 1),
        y: Number(document.querySelector('[data-designer-field="layer-scale-y"]')?.value || 1),
        anchorX: anchorRect ? anchorRect.left + anchorRect.width / 2 : null,
        anchorY: anchorRect ? anchorRect.top + anchorRect.height / 2 : null,
      };
    })()`);
    assert.ok(
      scaleAfterHandle.x !== scaleBeforeHandle.x || scaleAfterHandle.y !== scaleBeforeHandle.y,
      "Dragging a resize handle must change the selected layer scale",
    );
    const resizeRatioX = Math.abs(scaleAfterHandle.x / scaleBeforeHandle.x);
    const resizeRatioY = Math.abs(scaleAfterHandle.y / scaleBeforeHandle.y);
    assert.ok(
      Math.abs(resizeRatioX - resizeRatioY) < 0.04,
      `Corner resizing should preserve aspect ratio by default: ${JSON.stringify({ scaleBeforeHandle, scaleAfterHandle })}`,
    );
    assert.ok(
      Math.hypot(
        scaleAfterHandle.anchorX - resizeBox.anchorX,
        scaleAfterHandle.anchorY - resizeBox.anchorY,
      ) < 3.5,
      `PowerPoint-style corner resizing must keep the opposite anchor fixed: ${JSON.stringify({ resizeBox, scaleAfterHandle })}`,
    );

    const rotationDrag = await client.evaluate(`(() => {
      const handle = document.querySelector('[data-custom-symbol-rotation-handle]');
      const selection = document.querySelector('[data-custom-symbol-canvas-layer].is-selected .custom-symbol-designer-selection-box');
      if (!handle || !selection) return null;
      const hr = handle.getBoundingClientRect();
      const sr = selection.getBoundingClientRect();
      const centerX = sr.left + sr.width / 2;
      const centerY = sr.top + sr.height / 2;
      const startX = hr.left + hr.width / 2;
      const startY = hr.top + hr.height / 2;
      const dx = startX - centerX;
      const dy = startY - centerY;
      const radians = 55 * Math.PI / 180;
      return {
        centerX,
        centerY,
        startX,
        startY,
        targetX: centerX + dx * Math.cos(radians) - dy * Math.sin(radians),
        targetY: centerY + dx * Math.sin(radians) + dy * Math.cos(radians),
      };
    })()`);
    assert.ok(rotationDrag, "Direct rotation handle must be measurable");
    await client.send("Input.dispatchMouseEvent", {
      type: "mousePressed",
      x: rotationDrag.startX,
      y: rotationDrag.startY,
      button: "left",
      buttons: 1,
      clickCount: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseMoved",
      x: rotationDrag.targetX,
      y: rotationDrag.targetY,
      button: "left",
      buttons: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseReleased",
      x: rotationDrag.targetX,
      y: rotationDrag.targetY,
      button: "left",
      buttons: 0,
      clickCount: 1,
    });
    await sleep(120);
    const directRotation = await client.evaluate(`(() => ({
      rotation: Number(document.querySelector('[data-designer-field="layer-rotation"]')?.value || 0),
      transform: document.querySelector('[data-custom-symbol-canvas-layer].is-selected')?.getAttribute('transform') || '',
    }))()`);
    assert.ok(
      Math.abs(directRotation.rotation - scaleBeforeHandle.rotation) > 35,
      `Dragging the direct rotation handle must change the angle substantially: ${JSON.stringify({ scaleBeforeHandle, directRotation })}`,
    );
    assert.match(directRotation.transform, /rotate\(/);

    await setReactInput(client, '[data-designer-field="layer-rotation"]', 0);
    const snapRotationDrag = await client.evaluate(`(() => {
      const handle = document.querySelector('[data-custom-symbol-rotation-handle]');
      const selection = document.querySelector('[data-custom-symbol-canvas-layer].is-selected .custom-symbol-designer-selection-box');
      if (!handle || !selection) return null;
      const hr = handle.getBoundingClientRect();
      const sr = selection.getBoundingClientRect();
      const centerX = sr.left + sr.width / 2;
      const centerY = sr.top + sr.height / 2;
      const startX = hr.left + hr.width / 2;
      const startY = hr.top + hr.height / 2;
      const dx = startX - centerX;
      const dy = startY - centerY;
      const radians = 85 * Math.PI / 180;
      return {
        startX,
        startY,
        targetX: centerX + dx * Math.cos(radians) - dy * Math.sin(radians),
        targetY: centerY + dx * Math.sin(radians) + dy * Math.cos(radians),
      };
    })()`);
    assert.ok(snapRotationDrag);
    await client.send("Input.dispatchMouseEvent", {
      type: "mousePressed",
      x: snapRotationDrag.startX,
      y: snapRotationDrag.startY,
      button: "left",
      buttons: 1,
      clickCount: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseMoved",
      x: snapRotationDrag.targetX,
      y: snapRotationDrag.targetY,
      button: "left",
      buttons: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseReleased",
      x: snapRotationDrag.targetX,
      y: snapRotationDrag.targetY,
      button: "left",
      buttons: 0,
      clickCount: 1,
    });
    await sleep(100);
    const snappedRotation = Number(
      await client.evaluate(`document.querySelector('[data-designer-field="layer-rotation"]')?.value || 0`),
    );
    assert.equal(
      snappedRotation,
      90,
      `Rotation within 7° of a right angle must snap automatically: ${snappedRotation}`,
    );
    process.stdout.write("[custom-symbol-designer] PowerPoint-style resize, direct rotation and automatic 90-degree snap verified\n");

    const viewBoxBeforePan = await client.evaluate(`document.querySelector('[data-custom-symbol-canvas]')?.getAttribute('viewBox') || ''`);
    const emptyCanvasPoint = await client.evaluate(`(() => {
      const canvas = document.querySelector('[data-custom-symbol-canvas]');
      if (!canvas) return null;
      const rect = canvas.getBoundingClientRect();
      return { x: rect.left + 18, y: rect.top + rect.height / 2 };
    })()`);
    assert.ok(emptyCanvasPoint);
    await client.send("Input.dispatchMouseEvent", {
      type: "mousePressed",
      x: emptyCanvasPoint.x,
      y: emptyCanvasPoint.y,
      button: "left",
      buttons: 1,
      clickCount: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseMoved",
      x: emptyCanvasPoint.x + 45,
      y: emptyCanvasPoint.y + 18,
      button: "left",
      buttons: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseReleased",
      x: emptyCanvasPoint.x + 45,
      y: emptyCanvasPoint.y + 18,
      button: "left",
      buttons: 0,
      clickCount: 1,
    });
    await sleep(80);
    const viewBoxAfterPan = await client.evaluate(`document.querySelector('[data-custom-symbol-canvas]')?.getAttribute('viewBox') || ''`);
    assert.notEqual(viewBoxAfterPan, viewBoxBeforePan, "Dragging empty canvas space must pan the viewport");
    await client.evaluate(`document.querySelector('[data-custom-symbol-fit-view]').click()`);
    await sleep(80);
    await client.evaluate(`document.querySelector('[data-custom-symbol-layer]').click()`);
    process.stdout.write("[custom-symbol-designer] canvas pan verified\n");

    await client.evaluate(`document.querySelector('[data-duplicate-custom-symbol-layer]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 2`);
    const duplicateState = await client.evaluate(`(() => ({
      list: document.querySelectorAll('[data-custom-symbol-layer]').length,
      canvas: document.querySelectorAll('[data-custom-symbol-canvas-layer]').length,
      selected: document.querySelector('[data-custom-symbol-layer].is-selected')?.getAttribute('data-custom-symbol-layer') || "",
    }))()`);
    assert.equal(duplicateState.list, 2);
    assert.equal(duplicateState.canvas, 2);
    assert.ok(duplicateState.selected);

    await client.evaluate(`document.querySelector('[data-custom-symbol-layer].is-selected [data-toggle-custom-symbol-layer-visibility]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-canvas-layer]').length === 1`);
    await client.evaluate(`document.querySelector('[data-custom-symbol-layer].is-selected [data-toggle-custom-symbol-layer-visibility]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-canvas-layer]').length === 2`);
    await client.evaluate(`document.querySelector('[data-custom-symbol-layer].is-selected [data-toggle-custom-symbol-layer-lock]').click()`);
    const locked = await client.evaluate(`document.querySelector('[data-custom-symbol-layer].is-selected')?.getAttribute('data-layer-locked')`);
    assert.equal(locked, "true");
    await client.evaluate(`document.querySelector('[data-custom-symbol-layer].is-selected [data-toggle-custom-symbol-layer-lock]').click()`);
    process.stdout.write("[custom-symbol-designer] duplicate, visibility and lock verified\n");

    const beforeDrag = Number(
      await client.evaluate(`document.querySelector('[data-designer-field="layer-x"]')?.value || "0"`),
    );
    const box = await client.evaluate(`(() => {
      const layer = document.querySelector('[data-custom-symbol-canvas-layer].is-selected') ||
        document.querySelector('[data-custom-symbol-canvas-layer]');
      if (!layer) return null;
      const rect = layer.getBoundingClientRect();
      return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 };
    })()`);
    assert.ok(box, "Selected canvas layer must have a hit box");
    await client.send("Input.dispatchMouseEvent", {
      type: "mousePressed",
      x: box.x,
      y: box.y,
      button: "left",
      buttons: 1,
      clickCount: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseMoved",
      x: box.x + 45,
      y: box.y + 25,
      button: "left",
      buttons: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseReleased",
      x: box.x + 45,
      y: box.y + 25,
      button: "left",
      buttons: 0,
      clickCount: 1,
    });
    await sleep(140);
    const afterDrag = Number(
      await client.evaluate(`document.querySelector('[data-designer-field="layer-x"]')?.value || "0"`),
    );
    assert.notEqual(afterDrag, beforeDrag, "Canvas drag must update mathematical X");
    process.stdout.write("[custom-symbol-designer] CTM drag verified\n");

    await client.evaluate(`document.querySelector('[data-delete-custom-symbol-layer]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 1`);

    await client.evaluate(`(() => {
      const closeButtons = Array.from(document.querySelectorAll('.custom-symbol-designer-footer button'));
      closeButtons[0]?.click();
      return true;
    })()`);
    await waitUntil(client, `!document.querySelector('[data-custom-symbol-designer]')`);
    const afterClose = await client.evaluate(`(() => ({
      formula: document.querySelector("math-field")?.value || "",
      storage: localStorage.getItem("visualtex.custom-symbols.v1"),
    }))()`);
    assert.equal(afterClose.formula, baseline.formula);
    assert.equal(afterClose.storage, baseline.storage);

    await client.evaluate(`document.querySelector('[data-open-custom-symbol-designer]').click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-custom-symbol-designer]'))`);
    const reopened = await client.evaluate(`(() => ({
      layers: document.querySelectorAll('[data-custom-symbol-layer]').length,
      storage: localStorage.getItem("visualtex.custom-symbols.v1"),
      formula: document.querySelector("math-field")?.value || "",
    }))()`);
    assert.equal(reopened.layers, 1, "Designer draft should survive close/reopen in the same session");
    assert.equal(reopened.storage, baseline.storage);
    assert.equal(reopened.formula, baseline.formula);

    await client.evaluate(`document.querySelector('[data-custom-symbol-layer]').click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-custom-symbol-layer].is-selected'))`);
    await client.evaluate(`document.querySelector('[data-crop-preset="top"]').click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-designer-field="crop-height"]'))`);
    const croppedBeforeMove = await client.evaluate(`(() => {
      const interaction = document.querySelector('[data-custom-symbol-canvas-layer].is-selected');
      const id = interaction?.getAttribute('data-custom-symbol-canvas-layer') || '';
      const artwork = id ? document.querySelector('[data-custom-symbol-artwork-layer="' + id + '"]') : null;
      const clipRect = artwork?.querySelector('clipPath rect');
      return {
        transform: interaction?.getAttribute('transform') || "",
        clip: clipRect
          ? [clipRect.getAttribute('x'), clipRect.getAttribute('y'), clipRect.getAttribute('width'), clipRect.getAttribute('height')]
          : [],
        cropHeight: document.querySelector('[data-designer-field="crop-height"]')?.value || "",
      };
    })()`);
    assert.equal(croppedBeforeMove.clip.length, 4);
    assert.ok(Number(croppedBeforeMove.cropHeight) > 0);
    await setReactInput(client, '[data-designer-field="layer-x"]', 140);
    const croppedAfterMove = await client.evaluate(`(() => {
      const interaction = document.querySelector('[data-custom-symbol-canvas-layer].is-selected');
      const id = interaction?.getAttribute('data-custom-symbol-canvas-layer') || '';
      const artwork = id ? document.querySelector('[data-custom-symbol-artwork-layer="' + id + '"]') : null;
      const clipRect = artwork?.querySelector('clipPath rect');
      return {
        transform: interaction?.getAttribute('transform') || "",
        clip: clipRect
          ? [clipRect.getAttribute('x'), clipRect.getAttribute('y'), clipRect.getAttribute('width'), clipRect.getAttribute('height')]
          : [],
      };
    })()`);
    assert.match(croppedAfterMove.transform, /translate\(140 /);
    assert.deepEqual(
      croppedAfterMove.clip,
      croppedBeforeMove.clip,
      "Layer-local crop must move with the layer instead of changing its crop coordinates",
    );
    await setReactInput(client, '[data-designer-field="layer-rotation"]', 17);
    const croppedRotationCenter = await client.evaluate(`(() => {
      const interaction = document.querySelector('[data-custom-symbol-canvas-layer].is-selected');
      const id = interaction?.getAttribute('data-custom-symbol-canvas-layer') || '';
      const artwork = id ? document.querySelector('[data-custom-symbol-artwork-layer="' + id + '"]') : null;
      const clipRect = artwork?.querySelector('clipPath rect');
      const transform = interaction?.getAttribute('transform') || '';
      const origin = transform.match(/translate\\(([-0-9.]+) ([-0-9.]+)\\) rotate\\(17\\)/);
      const x = Number(clipRect?.getAttribute('x') || 0);
      const y = Number(clipRect?.getAttribute('y') || 0);
      const width = Number(clipRect?.getAttribute('width') || 0);
      const height = Number(clipRect?.getAttribute('height') || 0);
      return {
        transform,
        originX: Number(origin?.[1] || NaN),
        originY: Number(origin?.[2] || NaN),
        expectedX: x + width / 2,
        expectedY: y + height / 2,
      };
    })()`);
    assert.ok(Number.isFinite(croppedRotationCenter.originX));
    assert.ok(Number.isFinite(croppedRotationCenter.originY));
    assert.ok(
      Math.abs(croppedRotationCenter.originX - croppedRotationCenter.expectedX) < 0.001,
      `A cropped glyph must rotate around its visible horizontal center: ${JSON.stringify(croppedRotationCenter)}`,
    );
    assert.ok(
      Math.abs(croppedRotationCenter.originY - croppedRotationCenter.expectedY) < 0.001,
      `A cropped glyph must rotate around its visible vertical center: ${JSON.stringify(croppedRotationCenter)}`,
    );
    await setReactInput(client, '[data-designer-field="layer-rotation"]', 0);
    process.stdout.write("[custom-symbol-designer] local crop semantics and visual-center rotation verified\n");

    await client.evaluate(`document.querySelector('[data-split-custom-symbol-glyph="horizontal"]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 4`);
    const sliced = await client.evaluate(`(() => {
      const list = Array.from(document.querySelectorAll('[data-custom-symbol-layer]'));
      const visibleCanvas = Array.from(document.querySelectorAll('[data-custom-symbol-artwork-layer]'));
      const paths = visibleCanvas.map((layer) => layer.querySelector('path')?.getAttribute('d') || "");
      return {
        listCount: list.length,
        hiddenCount: list.filter((layer) => layer.getAttribute('data-layer-visible') === 'false').length,
        visibleCanvasCount: visibleCanvas.length,
        cropCount: visibleCanvas.filter((layer) => layer.querySelector('clipPath rect')).length,
        uniquePaths: Array.from(new Set(paths.filter(Boolean))).length,
      };
    })()`);
    assert.equal(sliced.listCount, 4);
    assert.equal(sliced.hiddenCount, 1, "Original full glyph must remain as a hidden recovery layer");
    assert.equal(sliced.visibleCanvasCount, 3);
    assert.equal(sliced.cropCount, 3);
    assert.equal(
      sliced.uniquePaths,
      1,
      "All slices must keep the same full source path and differ only by clipRect",
    );
    process.stdout.write("[custom-symbol-designer] non-destructive three-way slicing verified\n");

    const geometryPresets = ["line", "circle", "ellipse", "rect", "triangle", "arrow", "arc"];
    for (const preset of geometryPresets) {
      await client.evaluate(`document.querySelector('[data-add-custom-symbol-geometry="${preset}"]').click()`);
      await sleep(45);
    }
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer-kind="geometry"]').length === 7`);
    const geometryState = await client.evaluate(`(() => ({
      listCount: document.querySelectorAll('[data-custom-symbol-layer]').length,
      geometryCanvasCount: document.querySelectorAll('[data-custom-symbol-artwork-kind="geometry"]').length,
      totalCanvasCount: document.querySelectorAll('[data-custom-symbol-artwork-layer]').length,
      geometryTags: Array.from(document.querySelectorAll('[data-custom-symbol-artwork-kind="geometry"]')).map((layer) =>
        Array.from(layer.children).map((child) => child.tagName.toLowerCase()).filter((tag) => tag !== 'defs')
      ).flat(),
    }))()`);
    assert.equal(geometryState.listCount, 11);
    assert.equal(geometryState.geometryCanvasCount, 7);
    assert.equal(geometryState.totalCanvasCount, 10);
    assert.ok(geometryState.geometryTags.includes("line"));
    assert.ok(geometryState.geometryTags.includes("circle"));
    assert.ok(geometryState.geometryTags.includes("ellipse"));
    assert.ok(geometryState.geometryTags.includes("polygon"));
    assert.ok(geometryState.geometryTags.filter((tag) => tag === "path").length >= 2);

    await client.evaluate(`document.querySelector('[data-layer-geometry-preset="rect"]')?.click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-custom-symbol-geometry-properties]'))`);
    await setReactInput(client, '[data-designer-field="geometry-width"]', 520);
    await setReactInput(client, '[data-designer-field="geometry-height"]', 130);
    await setReactInput(client, '[data-designer-field="geometry-stroke-width"]', 10);
    await setReactInput(client, '[data-designer-field="geometry-corner-radius"]', 6);
    const rectGeometry = await client.evaluate(`(() => {
      const rect = document.querySelector('[data-custom-symbol-artwork-preset="rect"] rect');
      return rect ? {
        width: Number(rect.getAttribute('width')),
        height: Number(rect.getAttribute('height')),
        strokeWidth: Number(rect.getAttribute('stroke-width')),
        rx: Number(rect.getAttribute('rx')),
        fill: rect.getAttribute('fill'),
      } : null;
    })()`);
    assert.ok(rectGeometry);
    assert.equal(rectGeometry.width, 520);
    assert.equal(rectGeometry.height, 130);
    assert.equal(rectGeometry.strokeWidth, 10);
    assert.equal(rectGeometry.rx, 6);
    assert.equal(rectGeometry.fill, "none");
    await client.evaluate(`document.querySelector('[data-geometry-fill]')?.click()`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-artwork-preset="rect"] rect')?.getAttribute('fill') !== 'none'`);
    process.stdout.write("[custom-symbol-designer] direct geometry width/height/stroke/fill controls verified\n");

    await client.evaluate(`document.querySelector('[data-custom-symbol-eraser-tool]').click()`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-eraser-tool]')?.getAttribute('aria-pressed') === 'true'`);
    await setReactInput(client, '[data-custom-symbol-eraser-size-number]', 18);
    const eraserCanvas = await client.evaluate(`(() => {
      const canvas = document.querySelector('[data-custom-symbol-canvas]');
      const rect = canvas?.getBoundingClientRect();
      return rect ? { left: rect.left, top: rect.top, width: rect.width, height: rect.height } : null;
    })()`);
    assert.ok(eraserCanvas);
    const eraseStart = { x: eraserCanvas.left + eraserCanvas.width * 0.42, y: eraserCanvas.top + eraserCanvas.height * 0.49 };
    const eraseEnd = { x: eraserCanvas.left + eraserCanvas.width * 0.58, y: eraserCanvas.top + eraserCanvas.height * 0.53 };
    await client.send("Input.dispatchMouseEvent", { type: "mouseMoved", x: eraseStart.x, y: eraseStart.y, button: "none", buttons: 0 });
    await waitUntil(client, `Boolean(document.querySelector('[data-custom-symbol-eraser-cursor]'))`);
    const cursorState = await client.evaluate(`(() => {
      const cursor = document.querySelector('[data-custom-symbol-eraser-cursor] circle');
      return cursor ? { radius: Number(cursor.getAttribute('r')) } : null;
    })()`);
    assert.ok(cursorState);
    assert.ok(Math.abs(cursorState.radius - 9) < 0.001, "Eraser cursor radius must match half the precise 18-unit erase width");
    await client.send("Input.dispatchMouseEvent", { type: "mousePressed", x: eraseStart.x, y: eraseStart.y, button: "left", buttons: 1, clickCount: 1 });
    await client.send("Input.dispatchMouseEvent", { type: "mouseMoved", x: eraseStart.x + (eraseEnd.x - eraseStart.x) * 0.3, y: eraseStart.y + 7, button: "left", buttons: 1 });
    await client.send("Input.dispatchMouseEvent", { type: "mouseMoved", x: eraseStart.x + (eraseEnd.x - eraseStart.x) * 0.62, y: eraseEnd.y - 9, button: "left", buttons: 1 });
    await client.send("Input.dispatchMouseEvent", { type: "mouseMoved", x: eraseEnd.x, y: eraseEnd.y, button: "left", buttons: 1 });
    await client.send("Input.dispatchMouseEvent", { type: "mouseReleased", x: eraseEnd.x, y: eraseEnd.y, button: "left", buttons: 0, clickCount: 1 });
    await waitUntil(client, `Boolean(document.querySelector('[data-layer-geometry-preset="eraser"]'))`);
    const eraserState = await client.evaluate(`(() => {
      const eraserArtwork = document.querySelector('#visualtex-custom-symbol-designer-erase-mask path');
      return {
        layerOperation: document.querySelector('[data-custom-symbol-layer-operation="erase"]')?.getAttribute('data-custom-symbol-layer-operation') || '',
        mask: Boolean(document.querySelector('#visualtex-custom-symbol-designer-erase-mask')),
        eraserPreset: document.querySelector('[data-layer-geometry-preset="eraser"]')?.getAttribute('data-layer-geometry-preset') || '',
        d: eraserArtwork?.getAttribute('d') || '',
        strokeWidth: Number(eraserArtwork?.getAttribute('stroke-width') || 0),
        legacyOverlay: Boolean(document.querySelector('.custom-symbol-designer-eraser-overlay, .custom-symbol-designer-live-eraser')),
        selectedCenterline: Boolean(document.querySelector('[data-custom-symbol-eraser-centerline]')),
      };
    })()`);
    assert.equal(eraserState.layerOperation, "erase");
    assert.equal(eraserState.mask, true);
    assert.equal(eraserState.eraserPreset, "eraser");
    assert.match(eraserState.d, /C/, "Dragged eraser strokes must persist as a smooth cubic path instead of a polyline");
    assert.equal(eraserState.strokeWidth, 18);
    assert.equal(eraserState.legacyOverlay, false, "Completed eraser strokes must not leave the old thick red overlay");
    assert.equal(eraserState.selectedCenterline, false, "Centerline stays hidden while eraser mode is active");

    const clickPoint = { x: eraserCanvas.left + eraserCanvas.width * 0.36, y: eraserCanvas.top + eraserCanvas.height * 0.58 };
    const eraserCountBeforeClick = await client.evaluate(`document.querySelectorAll('[data-layer-geometry-preset="eraser"]').length`);
    await client.send("Input.dispatchMouseEvent", { type: "mousePressed", x: clickPoint.x, y: clickPoint.y, button: "left", buttons: 1, clickCount: 1 });
    await client.send("Input.dispatchMouseEvent", { type: "mouseReleased", x: clickPoint.x, y: clickPoint.y, button: "left", buttons: 0, clickCount: 1 });
    await waitUntil(client, `document.querySelectorAll('[data-layer-geometry-preset="eraser"]').length === ${eraserCountBeforeClick + 1}`);
    await client.evaluate(`document.querySelector('[data-custom-symbol-eraser-tool]').click()`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-eraser-tool]')?.getAttribute('aria-pressed') === 'false'`);
    const selectedCenterlineVisible = await client.evaluate(`Boolean(document.querySelector('[data-custom-symbol-eraser-centerline]'))`);
    assert.equal(selectedCenterlineVisible, true, "Selected erase strokes should expose only a thin editable centerline after leaving eraser mode");
    process.stdout.write("[custom-symbol-designer] smooth precise vector eraser interaction verified\n");

    await setReactInput(client, '[data-designer-field="layer-rotation"]', 27);
    const geometryTransform = await client.evaluate(`document.querySelector('[data-custom-symbol-canvas-layer].is-selected')?.getAttribute('transform') || ""`);
    assert.match(geometryTransform, /rotate\(27\)/);
    process.stdout.write("[custom-symbol-designer] geometry layers verified\n");

    const finalIsolation = await client.evaluate(`(() => ({
      formula: document.querySelector("math-field")?.value || "",
      storage: localStorage.getItem("visualtex.custom-symbols.v1"),
    }))()`);
    assert.equal(finalIsolation.formula, baseline.formula);
    assert.equal(finalIsolation.storage, baseline.storage);
    process.stdout.write("[custom-symbol-designer] pre-registration isolation verified\n");

    await client.evaluate(`document.querySelector('[data-reset-custom-symbol-designer]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 0`);
    await client.evaluate(`document.querySelector('[data-add-custom-symbol-material]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 1`);
    await waitUntil(
      client,
      `document.querySelector('[data-custom-symbol-ink-size]')?.textContent?.trim()?.endsWith('em') === true`,
    );
    const preRegistrationInkSize = await client.evaluate(
      `document.querySelector('[data-custom-symbol-ink-size]')?.textContent?.trim() || ''`,
    );
    assert.match(preRegistrationInkSize, /em$/);
    await setReactInput(client, '[data-custom-symbol-name-input]', "UI registered symbol");
    await setReactSelect(client, '[data-custom-symbol-role-select]', "relation");
    await setReactInput(client, '[data-custom-symbol-omml-fallback-input]', "\\approx");

    await setReactInput(client, '[data-custom-symbol-command-input]', "\\alpha");
    await client.evaluate(`document.querySelector('[data-register-custom-symbol]').click()`);
    await sleep(160);
    const builtinFailure = await client.evaluate(`(() => ({
      status: document.querySelector('[data-custom-symbol-registration-status]')?.getAttribute('data-custom-symbol-registration-status') || "",
      message: document.querySelector('[data-custom-symbol-registration-status]')?.textContent || "",
      storage: localStorage.getItem("visualtex.custom-symbols.v1"),
    }))()`);
    assert.equal(builtinFailure.status, "error");
    assert.match(builtinFailure.message, /alpha/i);
    assert.equal(builtinFailure.storage, baseline.storage);

    await setReactInput(client, '[data-custom-symbol-command-input]', "\\selfdef1");
    await client.evaluate(`document.querySelector('[data-register-custom-symbol]').click()`);
    await sleep(160);
    const numericFailure = await client.evaluate(`(() => ({
      status: document.querySelector('[data-custom-symbol-registration-status]')?.getAttribute('data-custom-symbol-registration-status') || "",
      storage: localStorage.getItem("visualtex.custom-symbols.v1"),
    }))()`);
    assert.equal(numericFailure.status, "error");
    assert.equal(numericFailure.storage, baseline.storage);

    await setReactInput(client, '[data-custom-symbol-command-input]', "\\selfdefa");
    await setReactInput(
      client,
      '[data-custom-symbol-omml-fallback-input]',
      "\\definitelyUnknownVisualTexCommand",
    );
    await client.evaluate(`document.querySelector('[data-register-custom-symbol]').click()`);
    await sleep(160);
    const fallbackFailure = await client.evaluate(`(() => ({
      status: document.querySelector('[data-custom-symbol-registration-status]')?.getAttribute('data-custom-symbol-registration-status') || "",
      storage: localStorage.getItem("visualtex.custom-symbols.v1"),
    }))()`);
    assert.equal(fallbackFailure.status, "error");
    assert.equal(fallbackFailure.storage, baseline.storage);
    process.stdout.write("[custom-symbol-designer] atomic registration failures verified\n");

    await setReactInput(client, '[data-custom-symbol-omml-fallback-input]', "\\approx");
    await client.evaluate(`document.querySelector('[data-register-custom-symbol]').click()`);
    await waitUntil(
      client,
      `document.querySelector('[data-custom-symbol-registration-status]')?.getAttribute('data-custom-symbol-registration-status') === 'success'`,
    );
    const registered = await client.evaluate(`(() => {
      const raw = localStorage.getItem("visualtex.custom-symbols.v1");
      const library = raw ? JSON.parse(raw) : null;
      const symbol = library?.symbols?.[0] || null;
      return {
        symbol,
        currentInkSize: document.querySelector('[data-custom-symbol-ink-size]')?.textContent?.trim() || '',
        hasManualCanvasMetrics: Boolean(
          document.querySelector('[data-designer-field="canvas-width"], [data-designer-field="canvas-ascent"], [data-designer-field="canvas-descent"]'),
        ),
        dirty: document.querySelector('[data-custom-symbol-registration-panel]')?.getAttribute('data-registration-dirty') || "",
        preview: Boolean(document.querySelector('[data-custom-symbol-registered-preview] .math-preview')),
      };
    })()`);
    assert.equal(registered.symbol.command, "selfdefa");
    assert.equal(registered.symbol.name, "UI registered symbol");
    assert.equal(registered.symbol.role, "relation");
    assert.equal(registered.symbol.ommlFallback, "\\approx");
    assert.ok(registered.symbol.artwork.shapes.length > 0);
    assert.match(registered.currentInkSize, /em$/);
    assert.equal(registered.hasManualCanvasMetrics, false);
    assert.ok(registered.symbol.designerSource?.metrics?.widthEm > 0);
    assert.ok(registered.symbol.designerSource?.metrics?.ascentEm > 0);
    assert.ok(registered.symbol.metrics.widthEm > 0);
    assert.ok(
      registered.symbol.metrics.widthEm <=
        registered.symbol.designerSource.metrics.widthEm,
      "Runtime registered width must be derived from visible ink rather than an exposed fixed canvas box",
    );
    assert.equal(registered.dirty, "false");
    assert.equal(registered.preview, true);

    const runtimeRegistration = await client.evaluate(`(async () => {
      const search = await import("/src/autocomplete/CommandSearchEngine.ts");
      const runtime = await import("/src/export/runtime.ts");
      const BS = String.fromCharCode(92);
      const command = BS + "selfdefa";
      const field = document.querySelector("math-field");
      field.setValue("A+" + command + "+B", {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
      field.dispatchEvent(
        new InputEvent("input", {
          bubbles: true,
          inputType: "insertText",
          data: "a",
        }),
      );
      await new Promise((resolve) => setTimeout(resolve, 90));
      const stored = JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1"));
      const symbol = stored.symbols[0];
      const svg = runtime.latexToSvg(command, {
        displayMode: false,
        fontSizePt: 12,
        paddingPx: 0,
        background: "transparent",
      }).svg;
      const mathMl = runtime.latexToMathMl(command, false);
      return {
        value: field.value,
        renderedClass: Boolean(
          field.shadowRoot?.querySelector(".visualtex-custom-symbol-" + symbol.id),
        ),
        shadowStyle: Array.from(field.shadowRoot?.querySelectorAll("style") || []).some(
          (style) => style.textContent?.includes("visualtex-custom-symbol-" + symbol.id),
        ),
        search: search
          .searchCommands(BS + "selfdefa", {}, false, 10)
          .some((entry) => entry.command === command),
        svg: svg.includes('data-visualtex-custom-symbol="' + symbol.id + '"'),
        fallback: /2248/i.test(mathMl),
      };
    })()`);
    assert.equal(runtimeRegistration.value, "A+\\selfdefa+B");
    assert.equal(runtimeRegistration.renderedClass, true);
    assert.equal(runtimeRegistration.shadowStyle, true);
    assert.equal(runtimeRegistration.search, true);
    assert.equal(runtimeRegistration.svg, true);
    assert.equal(runtimeRegistration.fallback, true);
    process.stdout.write("[custom-symbol-designer] successful runtime registration verified\n");

    await setReactInput(client, '[data-custom-symbol-name-input]', "UI registered symbol updated");
    await setReactInput(client, '[data-designer-field="layer-rotation"]', 18);
    const dirtyState = await client.evaluate(`document.querySelector('[data-custom-symbol-registration-panel]')?.getAttribute('data-registration-dirty')`);
    assert.equal(dirtyState, "true");
    await client.evaluate(`document.querySelector('[data-register-custom-symbol]').click()`);
    await waitUntil(
      client,
      `document.querySelector('[data-custom-symbol-registration-status]')?.getAttribute('data-custom-symbol-registration-status') === 'success'`,
    );
    const updated = await client.evaluate(`(() => {
      const symbol = JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1")).symbols[0];
      return {
        name: symbol.name,
        command: symbol.command,
        rotateDeg: symbol.artwork.shapes[0]?.transform?.rotateDeg ?? null,
        dirty: document.querySelector('[data-custom-symbol-registration-panel]')?.getAttribute('data-registration-dirty') || "",
      };
    })()`);
    assert.equal(updated.name, "UI registered symbol updated");
    assert.equal(updated.command, "selfdefa");
    assert.equal(updated.rotateDeg, 18);
    assert.equal(updated.dirty, "false");

    await setReactInput(client, '[data-custom-symbol-command-input]', "\\alpha");
    await client.evaluate(`document.querySelector('[data-register-custom-symbol]').click()`);
    await sleep(160);
    const failedUpdate = await client.evaluate(`(() => {
      const symbol = JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1")).symbols[0];
      return {
        status: document.querySelector('[data-custom-symbol-registration-status]')?.getAttribute('data-custom-symbol-registration-status') || "",
        storedCommand: symbol.command,
        storedName: symbol.name,
        previewCode: document.querySelector('[data-custom-symbol-registered-preview] code')?.textContent || "",
      };
    })()`);
    assert.equal(failedUpdate.status, "error");
    assert.equal(failedUpdate.storedCommand, "selfdefa");
    assert.equal(failedUpdate.storedName, "UI registered symbol updated");
    assert.equal(failedUpdate.previewCode, "\\selfdefa");
    process.stdout.write("[custom-symbol-designer] atomic update failure verified\n");

    await client.evaluate(`document.querySelector('[data-reset-custom-symbol-designer]').click()`);
    await sleep(100);
    const resetSafety = await client.evaluate(`(() => ({
      layers: document.querySelectorAll('[data-custom-symbol-layer]').length,
      symbolCount: JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1")).symbols.length,
      command: JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1")).symbols[0].command,
    }))()`);
    assert.equal(resetSafety.layers, 0);
    assert.equal(resetSafety.symbolCount, 1);
    assert.equal(resetSafety.command, "selfdefa");

    const archiveState = await client.evaluate(`(() => {
      const symbol = JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1")).symbols[0];
      const item = document.querySelector('[data-registered-custom-symbol-command="selfdefa"]');
      return {
        id: symbol.id,
        archiveVersion: symbol.designerSource?.version ?? null,
        assetCount: symbol.designerSource?.assets?.length ?? -1,
        layerCount: symbol.designerSource?.layers?.length ?? -1,
        sourceLatex: symbol.designerSource?.assets?.[0]?.sourceLatex ?? "",
        listItem: Boolean(item),
      };
    })()`);
    assert.equal(archiveState.archiveVersion, 1);
    assert.equal(archiveState.assetCount, 1);
    assert.equal(archiveState.layerCount, 1);
    assert.equal(archiveState.sourceLatex, "\\partial");
    assert.equal(archiveState.listItem, true);

    await client.evaluate(`document.querySelector('[data-registered-custom-symbol-command="selfdefa"] [data-edit-registered-custom-symbol]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 1`);
    const restoredRegistered = await client.evaluate(`(() => {
      const layer = document.querySelector('[data-custom-symbol-layer]');
      const panel = document.querySelector('[data-custom-symbol-registration-panel]');
      return {
        kind: layer?.getAttribute('data-layer-kind') || "",
        sourceLatex: layer?.getAttribute('data-layer-source-latex') || "",
        symbolId: panel?.getAttribute('data-registration-symbol-id') || "",
        command: document.querySelector('[data-custom-symbol-command-input]')?.value || "",
        legacyWarning: Boolean(document.querySelector('[data-custom-symbol-legacy-warning]')),
      };
    })()`);
    assert.equal(restoredRegistered.kind, "glyph");
    assert.equal(restoredRegistered.sourceLatex, "\\partial");
    assert.equal(restoredRegistered.symbolId, archiveState.id);
    assert.equal(restoredRegistered.command, "selfdefa");
    assert.equal(restoredRegistered.legacyWarning, false);
    process.stdout.write("[custom-symbol-designer] editable source archive restored\n");

    await client.evaluate(`document.querySelector('[data-registered-custom-symbol-command="selfdefa"] [data-duplicate-registered-custom-symbol]').click()`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-command-input]')?.value === "selfdefacopy"`);
    const duplicateDraft = await client.evaluate(`(() => ({
      symbolId: document.querySelector('[data-custom-symbol-registration-panel]')?.getAttribute('data-registration-symbol-id') || "",
      command: document.querySelector('[data-custom-symbol-command-input]')?.value || "",
      layers: document.querySelectorAll('[data-custom-symbol-layer]').length,
      symbolCount: JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1")).symbols.length,
    }))()`);
    assert.equal(duplicateDraft.symbolId, "");
    assert.equal(duplicateDraft.command, "selfdefacopy");
    assert.equal(duplicateDraft.layers, 1);
    assert.equal(duplicateDraft.symbolCount, 1);
    process.stdout.write("[custom-symbol-designer] duplicate-as-draft isolation verified\n");

    await client.evaluate(`document.querySelector('[data-reset-custom-symbol-designer]').click()`);
    await client.evaluate(`document.querySelector('[data-registered-custom-symbol-command="selfdefa"] [data-edit-registered-custom-symbol]').click()`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-registration-panel]')?.getAttribute('data-registration-symbol-id') === ${JSON.stringify(archiveState.id)}`);
    const tileStorageBeforeInsert = await client.evaluate(`localStorage.getItem("visualtex-custom-formula-tiles")`);
    await client.evaluate(`document.querySelector('.custom-symbol-designer-footer button').click()`);
    await waitUntil(client, `!document.querySelector('[data-custom-symbol-designer]')`);
    await waitUntil(client, `Boolean(document.querySelector('[data-registered-custom-symbol-command="selfdefa"]'))`);
    const mainToolbarDeleteVisible = await client.evaluate(`(() => {
      const button = document.querySelector('[data-delete-registered-custom-symbol-toolbar]');
      if (!(button instanceof HTMLElement)) return false;
      const rect = button.getBoundingClientRect();
      const style = getComputedStyle(button);
      return rect.width > 0 && rect.height > 0 && style.display !== "none" && style.visibility !== "hidden";
    })()`);
    assert.equal(mainToolbarDeleteVisible, true, "Registered custom symbols must expose a visible delete button in the main toolbar");

    await client.evaluate(`document.querySelector('[data-tile-category="common"]').click()`);
    await sleep(100);
    const commonTileIsolation = await client.evaluate(`(() => ({
      registeredTile: Boolean(document.querySelector('[data-formula-tile-id="registered-symbol-${archiveState.id}"]')),
      registeredCommand: Array.from(document.querySelectorAll('[data-formula-tile-latex]')).some(
        (tile) => tile.getAttribute('data-formula-tile-latex') === "\\\\selfdefa",
      ),
    }))()`);
    assert.deepEqual(
      commonTileIsolation,
      { registeredTile: false, registeredCommand: false },
      "Registered custom symbols must not be injected into the Common tile category",
    );
    await client.evaluate(`document.querySelector('[data-tile-category="custom"]').click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-registered-custom-symbol-command="selfdefa"]'))`);
    process.stdout.write("[custom-symbol-designer] registered symbol Common-category isolation verified\n");

    await client.evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      field.setValue("", {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
      field.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "deleteContentBackward", data: null }));
      field.position = field.lastOffset;
    })()`);
    await sleep(100);
    await client.evaluate(`document.querySelector('[data-registered-custom-symbol-command="selfdefa"] .formula-tile-button.is-registered-custom-symbol').click()`);
    await waitUntil(client, `document.querySelector("math-field")?.value === "\\\\selfdefa"`);
    const toolbarInsert = await client.evaluate(`(() => {
      const field = document.querySelector("math-field");
      const symbol = JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1")).symbols[0];
      return {
        value: field.value,
        rendered: Boolean(field.shadowRoot?.querySelector('.visualtex-custom-symbol-' + symbol.id)),
        tileStorage: localStorage.getItem("visualtex-custom-formula-tiles"),
      };
    })()`);
    assert.equal(toolbarInsert.value, "\\selfdefa");
    assert.equal(toolbarInsert.rendered, true);
    assert.equal(toolbarInsert.tileStorage, tileStorageBeforeInsert);
    process.stdout.write("[custom-symbol-designer] registered toolbar insertion verified\n");

    await client.evaluate(`document.querySelector('[data-open-custom-symbol-designer]').click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-custom-symbol-designer]'))`);
    await client.evaluate(`document.querySelector('[data-registered-custom-symbol-command="selfdefa"] [data-edit-registered-custom-symbol]').click()`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-registration-panel]')?.getAttribute('data-registration-symbol-id') === ${JSON.stringify(archiveState.id)}`);

    await client.evaluate(`document.querySelector('[data-registered-custom-symbol-command="selfdefa"] [data-delete-registered-custom-symbol]').click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-custom-symbol-delete-warning]'))`);
    await client.evaluate(`document.querySelector('[data-cancel-delete-registered-custom-symbol]').click()`);
    await waitUntil(client, `!document.querySelector('[data-custom-symbol-delete-warning]')`);
    const afterCancelDelete = await client.evaluate(`JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1")).symbols.length`);
    assert.equal(afterCancelDelete, 1);

    await client.evaluate(`document.querySelector('[data-registered-custom-symbol-command="selfdefa"] [data-delete-registered-custom-symbol]').click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-confirm-delete-registered-custom-symbol]'))`);
    await client.evaluate(`document.querySelector('[data-confirm-delete-registered-custom-symbol]').click()`);
    await waitUntil(client, `JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1")).symbols.length === 0`);
    await waitUntil(client, `!document.querySelector('[data-registered-custom-symbol-command="selfdefa"]')`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-registration-status="success"]')?.textContent?.includes('selfdefa')`);
    const deletionState = await client.evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        formula: field.value,
        registeredListCount: document.querySelectorAll('[data-registered-custom-symbol]').length,
        toolbarButton: Boolean(document.querySelector('[data-registered-custom-symbol-command="selfdefa"]')),
        customClass: Boolean(field.shadowRoot?.querySelector('.visualtex-custom-symbol-${archiveState.id}')),
        designerLayers: document.querySelectorAll('[data-custom-symbol-layer]').length,
        registrationId: document.querySelector('[data-custom-symbol-registration-panel]')?.getAttribute('data-registration-symbol-id') || "",
      };
    })()`);
    assert.equal(deletionState.formula.trim(), "\\selfdefa");
    assert.equal(deletionState.registeredListCount, 0);
    assert.equal(deletionState.toolbarButton, false);
    assert.equal(deletionState.customClass, false);
    assert.equal(deletionState.designerLayers, 0);
    assert.equal(deletionState.registrationId, "");
    process.stdout.write("[custom-symbol-designer] two-step deletion and unresolved-source preservation verified\n");

    console.log(
      "Custom symbol designer UI composition, geometry, local slicing, transactional registration/update, editable-source restore, toolbar linkage, deletion safety, fallback, and Proof theme purity regression passed",
    );
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    vite.kill("SIGTERM");
    await sleep(240);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => undefined);
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
