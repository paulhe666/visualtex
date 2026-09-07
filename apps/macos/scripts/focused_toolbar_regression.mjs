import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const mode = process.argv[2];
const supportedModes = new Set([
  "row-spacing",
  "continuous-layout",
  "continuous-theme",
  "continuous-wheel",
  "continuous-performance",
  "continuous-fps",
  "piecewise-preview",
]);
if (!supportedModes.has(mode)) {
  throw new Error(
    "Usage: node scripts/focused_toolbar_regression.mjs <row-spacing|continuous-layout|continuous-theme|continuous-wheel|continuous-performance|continuous-fps|piecewise-preview>",
  );
}

const offset = process.pid % 1000;
const previewPort = 17300 + offset;
const debugPort = 22500 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = `/tmp/visualtex-focused-toolbar-${mode}-${process.pid}`;
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
        ...(mode === "continuous-fps" ? [] : ["--disable-gpu"]),
        "--no-first-run",
        "--no-default-browser-check",
        `--remote-debugging-port=${debugPort}`,
        `--user-data-dir=${chromeProfile}`,
        "--window-size=1400,1000",
        "about:blank",
      ],
      { stdio: "ignore" },
    );

    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json();
    const page =
      targets.find(
        (target) => target.type === "page" && target.url.startsWith(baseUrl),
      ) ?? targets.find((target) => target.type === "page");
    if (!page) throw new Error("No Chrome page target found");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.addScriptToEvaluateOnNewDocument", {
      source: `(() => {
        if (location.origin !== ${JSON.stringify(baseUrl)}) return;
        localStorage.setItem("visualtex.onboarding.v3.completed", "true");
        localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
        localStorage.setItem(
          "visualtex.onboarding.macos.desktop.v1.2.0.completed",
          "true",
        );
        localStorage.setItem(
          "visualtex.office.macos.native-first-run.v1.2.0.completed",
          "true",
        );
        localStorage.setItem("visualtex-desktop-editor-toolbar-open", "true");
        const storageKey = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
        const lineId = "focused-toolbar-line";
        persisted.state = {
          ...(persisted.state || {}),
          lines: [{ id: lineId, latex: "x=\\\\frac{-b\\\\pm\\\\sqrt{b^2-4ac}}{2a}" }],
          activeLineId: lineId,
          editorLayout: "classic",
          formulaToolButtonSize: 52,
          formulaToolButtonPadding: 2,
          formulaRowVerticalInset: 5,
          checkUpdatesOnStartup: false,
        };
        localStorage.setItem(storageKey, JSON.stringify(persisted));
      })();`,
    });

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

    const waitForEvaluation = async (
      expression,
      description,
      timeoutMs = 10000,
    ) => {
      const started = Date.now();
      let lastValue;
      while (Date.now() - started < timeoutMs) {
        try {
          lastValue = await evaluate(expression);
          if (lastValue?.ready) return lastValue;
        } catch (error) {
          const message = String(error?.message ?? error);
          if (!/navigated|execution context|target closed/i.test(message)) {
            throw error;
          }
        }
        await sleep(50);
      }
      throw new Error(
        `Timed out waiting for ${description}: ${JSON.stringify(lastValue)}`,
      );
    };

    await client.send("Page.navigate", { url: baseUrl });
    await sleep(800);
    await waitForEvaluation(
      `(() => ({
        ready: Boolean(
          document.querySelector("math-field") &&
          document.querySelector(".formula-toolbar") &&
          document.querySelector(".template-strip.is-continuous-categories")
        ),
      }))()`,
      "focused toolbar workspace",
    );

    if (mode === "row-spacing") {
      await evaluate(`document.querySelector('.settings-toggle')?.click()`);
      await waitForEvaluation(
        `(() => ({
          ready: Boolean(document.querySelector('[data-interface-customization-trigger]')),
        }))()`,
        "settings dialog",
      );
      await evaluate(
        `document.querySelector('[data-interface-customization-trigger]')?.click()`,
      );
      await waitForEvaluation(
        `(() => ({
          ready: Boolean(
            document.querySelector('[data-formula-row-vertical-inset-setting]')
          ),
        }))()`,
        "formula row vertical spacing setting",
      );

      const defaultState = await evaluate(`(() => {
        const rowSlider = document.querySelector(
          '[data-formula-row-vertical-inset-setting]',
        );
        const obsoleteToolbarSlider = document.querySelector(
          '[data-formula-tool-button-vertical-padding-setting]',
        );
        const field = document.querySelector('.formula-line .visual-mathfield');
        const previewRow = document.querySelector('.formula-inset-preview-row');
        const toolbarButton = document.querySelector('.template-button');
        const fieldStyle = field ? getComputedStyle(field) : null;
        const previewStyle = previewRow ? getComputedStyle(previewRow) : null;
        const toolbarStyle = toolbarButton
          ? getComputedStyle(toolbarButton)
          : null;
        return {
          value: Number(rowSlider?.value ?? -1),
          obsoleteToolbarSlider: Boolean(obsoleteToolbarSlider),
          fieldTop: fieldStyle ? parseFloat(fieldStyle.paddingTop) : -1,
          fieldBottom: fieldStyle ? parseFloat(fieldStyle.paddingBottom) : -1,
          previewTop: previewStyle ? parseFloat(previewStyle.paddingTop) : -1,
          previewBottom: previewStyle
            ? parseFloat(previewStyle.paddingBottom)
            : -1,
          toolbarTop: toolbarStyle ? parseFloat(toolbarStyle.paddingTop) : -1,
          toolbarLeft: toolbarStyle ? parseFloat(toolbarStyle.paddingLeft) : -1,
        };
      })()`);
      assert.deepEqual(defaultState, {
        value: 5,
        obsoleteToolbarSlider: false,
        fieldTop: 5,
        fieldBottom: 5,
        previewTop: 5,
        previewBottom: 5,
        toolbarTop: 2,
        toolbarLeft: 2,
      });

      await evaluate(`(() => {
        const slider = document.querySelector(
          '[data-formula-row-vertical-inset-setting]',
        );
        const setter = Object.getOwnPropertyDescriptor(
          HTMLInputElement.prototype,
          'value',
        )?.set;
        setter?.call(slider, '11');
        slider.dispatchEvent(new Event('input', { bubbles: true }));
        slider.dispatchEvent(new Event('change', { bubbles: true }));
      })()`);
      const changedState = await waitForEvaluation(`(() => {
        const field = document.querySelector('.formula-line .visual-mathfield');
        const previewRow = document.querySelector('.formula-inset-preview-row');
        const toolbarButton = document.querySelector('.template-button');
        const fieldStyle = field ? getComputedStyle(field) : null;
        const previewStyle = previewRow ? getComputedStyle(previewRow) : null;
        const toolbarStyle = toolbarButton
          ? getComputedStyle(toolbarButton)
          : null;
        const persisted = JSON.parse(
          localStorage.getItem('visualtex-editor') || '{}',
        ).state ?? {};
        const value = Number(
          document.querySelector(
            '[data-formula-row-vertical-inset-setting]',
          )?.value ?? -1,
        );
        const state = {
          value,
          fieldTop: fieldStyle ? parseFloat(fieldStyle.paddingTop) : -1,
          fieldBottom: fieldStyle ? parseFloat(fieldStyle.paddingBottom) : -1,
          previewTop: previewStyle ? parseFloat(previewStyle.paddingTop) : -1,
          previewBottom: previewStyle
            ? parseFloat(previewStyle.paddingBottom)
            : -1,
          toolbarTop: toolbarStyle ? parseFloat(toolbarStyle.paddingTop) : -1,
          toolbarLeft: toolbarStyle ? parseFloat(toolbarStyle.paddingLeft) : -1,
          persisted: persisted.formulaRowVerticalInset ?? null,
        };
        return {
          ready:
            state.value === 11 &&
            state.fieldTop === 11 &&
            state.fieldBottom === 11 &&
            state.previewTop === 11 &&
            state.previewBottom === 11 &&
            state.toolbarTop === 2 &&
            state.toolbarLeft === 2 &&
            state.persisted === 11,
          ...state,
        };
      })()`, "formula-row spacing changes without altering toolbar padding");
      console.log(JSON.stringify(changedState));
    }

    if (mode === "continuous-layout") {
      const layoutState = await evaluate(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const sections = [...(strip?.querySelectorAll(
          ':scope > .toolbar-category-section',
        ) ?? [])];
        const expected = [
          'common',
          'structure',
          'calculus',
          'matrix',
          'relation',
          'greek',
          'arrow',
          'physics',
          'set',
        ];
        const actual = sections.map(
          (section) => section.dataset.toolbarCategorySection,
        );
        const categoryTints = sections.map((section) =>
          getComputedStyle(section)
            .getPropertyValue('--toolbar-category-tint')
            .trim(),
        );
        const tabs = [...document.querySelectorAll('.toolbar-tab')];
        const tabSeparatorBackgrounds = tabs.slice(1).map((tab) =>
          getComputedStyle(tab, '::before').backgroundColor,
        );
        const categorySectionBackgrounds = sections.map(
          (section) => getComputedStyle(section).backgroundColor,
        );
        const categoryButtonBackgrounds = sections.map((section) => {
          const button = section.querySelector(':scope > .template-button');
          return button ? getComputedStyle(button).backgroundImage : '';
        });
        const gaps = sections.slice(1).map((section, index) => {
          const previous = sections[index];
          return section.getBoundingClientRect().left -
            previous.getBoundingClientRect().right;
        });
        const calculus = sections.find(
          (section) => section.dataset.toolbarCategorySection === 'calculus',
        );
        const common = sections.find(
          (section) => section.dataset.toolbarCategorySection === 'common',
        );
        const rowCount = Number(
          document.querySelector('.formula-toolbar')?.dataset.toolbarRowCount ?? 0,
        );
        const calculusButtons = [...(calculus?.querySelectorAll(
          ':scope > .template-button',
        ) ?? [])];
        const commonButtons = [...(common?.querySelectorAll(
          ':scope > .template-button',
        ) ?? [])];
        const columnCounts = new Map();
        for (const button of calculusButtons) {
          const left = Math.round(button.getBoundingClientRect().left);
          columnCounts.set(left, (columnCounts.get(left) ?? 0) + 1);
        }
        const lastColumnCount = [...columnCounts.entries()]
          .sort((left, right) => left[0] - right[0])
          .at(-1)?.[1] ?? 0;
        const remainder = rowCount > 0
          ? calculusButtons.length % rowCount
          : -1;
        return {
          expected,
          actual,
          categoryTints,
          tabSeparatorBackgrounds,
          categorySectionBackgrounds,
          categoryButtonBackgrounds,
          gaps,
          rowCount,
          calculusCount: calculusButtons.length,
          calculusRemainder: remainder,
          calculusLastColumnCount: lastColumnCount,
          commonCount: commonButtons.length,
          commonRemainder:
            rowCount > 0 ? commonButtons.length % rowCount : -1,
          transitionCards:
            strip?.querySelectorAll('.toolbar-category-transition').length ?? -1,
        };
      })()`);
      assert.deepEqual(layoutState.actual, layoutState.expected);
      assert.equal(layoutState.transitionCards, 0);
      assert.equal(new Set(layoutState.categoryTints).size, layoutState.expected.length);
      assert.equal(layoutState.categoryTints[0], '#d5c5b4');
      assert.ok(layoutState.categoryTints.every(Boolean));
      assert.equal(layoutState.tabSeparatorBackgrounds.length, 8);
      assert.ok(
        layoutState.tabSeparatorBackgrounds.every(
          (background) =>
            background &&
            background !== 'rgba(0, 0, 0, 0)' &&
            background !== 'transparent',
        ),
      );
      assert.equal(
        new Set(layoutState.categorySectionBackgrounds).size,
        layoutState.expected.length,
      );
      assert.equal(
        new Set(layoutState.categoryButtonBackgrounds).size,
        layoutState.expected.length,
      );
      assert.ok(
        layoutState.categoryButtonBackgrounds.every(
          (background) => background && background !== 'none',
        ),
      );
      assert.ok(layoutState.gaps.every((gap) => gap >= 12 && gap <= 16));
      assert.ok(layoutState.calculusRemainder > 0);
      assert.equal(
        layoutState.calculusLastColumnCount,
        layoutState.calculusRemainder,
      );
      console.log(JSON.stringify(layoutState));
    }

    if (mode === "continuous-theme") {
      const themeState = await evaluate(`(async () => {
        const root = document.documentElement;
        const themeIds = ['default', 'beige', 'dark', 'purple', 'green'];
        const categories = [
          'common',
          'structure',
          'calculus',
          'matrix',
          'relation',
          'greek',
          'arrow',
          'physics',
          'set',
        ];
        const parseColor = (value) => {
          const rgb = value.match(/rgba?\\(([^)]+)\\)/i);
          if (rgb) {
            const parts = rgb[1].split(/[\\s,\\/]+/).filter(Boolean).slice(0, 3);
            return parts.map((part) =>
              part.endsWith('%')
                ? Number.parseFloat(part) * 2.55
                : Number.parseFloat(part),
            );
          }
          const srgb = value.match(
            /color\\(srgb\\s+([\\d.]+)\\s+([\\d.]+)\\s+([\\d.]+)/i,
          );
          return srgb
            ? srgb.slice(1, 4).map((part) => Number.parseFloat(part) * 255)
            : null;
        };
        const luminance = (value) => {
          const color = parseColor(value);
          if (!color) return 0;
          const channels = color.map((channel) => {
            const normalized = channel / 255;
            return normalized <= 0.04045
              ? normalized / 12.92
              : ((normalized + 0.055) / 1.055) ** 2.4;
          });
          return channels[0] * 0.2126 + channels[1] * 0.7152 + channels[2] * 0.0722;
        };
        const contrast = (foreground, background) => {
          const first = luminance(foreground);
          const second = luminance(background);
          return (Math.max(first, second) + 0.05) /
            (Math.min(first, second) + 0.05);
        };
        const settle = () => new Promise((resolve) =>
          requestAnimationFrame(() => requestAnimationFrame(resolve)),
        );
        const results = [];
        try {
          for (const theme of themeIds) {
            if (theme === 'default') delete root.dataset.theme;
            else root.dataset.theme = theme;
            await settle();
            const rootStyle = getComputedStyle(root);
            const tabs = categories.map((category) =>
              document.querySelector('.toolbar-tab[data-category="' + category + '"]'),
            );
            const sections = categories.map((category) =>
              document.querySelector(
                '[data-toolbar-category-section="' + category + '"]',
              ),
            );
            const matrixSection = sections[categories.indexOf('matrix')];
            const matrixBuilder = matrixSection?.querySelector('.matrix-builder');
            const activeDelimiter = matrixBuilder?.querySelector(
              '.matrix-delimiter-options button.is-active',
            );
            const sizeGrid = matrixBuilder?.querySelector('.matrix-size-grid');
            const insertButton = matrixBuilder?.querySelector(
              '.matrix-insert-button',
            );
            const tabDetails = tabs.map((tab, index) => {
              const style = getComputedStyle(tab);
              return {
                category: categories[index],
                tint: style.getPropertyValue('--toolbar-category-tint').trim(),
                foreground: style.color,
                background: style.backgroundColor,
                contrast: contrast(style.color, style.backgroundColor),
              };
            });
            const matrixStyle = getComputedStyle(matrixBuilder);
            const delimiterStyle = getComputedStyle(activeDelimiter);
            const sizeGridStyle = getComputedStyle(sizeGrid);
            const insertStyle = getComputedStyle(insertButton);
            results.push({
              theme,
              colorScheme: rootStyle.colorScheme,
              surface: rootStyle.getPropertyValue('--surface').trim(),
              categoryTints: tabDetails.map((item) => item.tint),
              minimumTabContrast: Math.min(
                ...tabDetails.map((item) => item.contrast),
              ),
              matrixToken: rootStyle
                .getPropertyValue('--toolbar-category-matrix')
                .trim(),
              matrixSectionTint: getComputedStyle(matrixSection)
                .getPropertyValue('--toolbar-category-tint')
                .trim(),
              matrixBuilderBackground: matrixStyle.backgroundImage,
              matrixBuilderBorder: matrixStyle.borderColor,
              matrixDelimiterForeground: delimiterStyle.color,
              matrixDelimiterBackground: delimiterStyle.backgroundColor,
              matrixDelimiterContrast: contrast(
                delimiterStyle.color,
                delimiterStyle.backgroundColor,
              ),
              matrixSizeGridBackground: sizeGridStyle.backgroundColor,
              matrixInsertForeground: insertStyle.color,
              matrixInsertBackground: insertStyle.backgroundColor,
              matrixInsertContrast: contrast(
                insertStyle.color,
                insertStyle.backgroundColor,
              ),
            });
          }
        } finally {
          delete root.dataset.theme;
          await settle();
        }
        return results;
      })()`);
      console.log(JSON.stringify(themeState));
      assert.equal(themeState.length, 5);
      assert.equal(
        new Set(themeState.map((state) => state.categoryTints.join('|'))).size,
        themeState.length,
      );
      assert.equal(
        new Set(themeState.map((state) => state.matrixBuilderBackground)).size,
        themeState.length,
      );
      for (const state of themeState) {
        assert.equal(new Set(state.categoryTints).size, 9);
        assert.equal(state.matrixSectionTint, state.matrixToken);
        assert.ok(state.minimumTabContrast >= 4.5);
        assert.ok(
          state.matrixBuilderBackground &&
          state.matrixBuilderBackground !== 'none',
        );
        assert.ok(state.matrixBuilderBorder);
        assert.notEqual(state.matrixDelimiterBackground, state.surface);
        assert.ok(state.matrixDelimiterContrast >= 4.5);
        assert.notEqual(state.matrixSizeGridBackground, state.surface);
        assert.ok(state.matrixInsertBackground);
        assert.ok(state.matrixInsertContrast >= 4.5);
      }
    }

    if (mode === "continuous-wheel") {
      const wheelState = await evaluate(`(async () => {
        const toolbar = document.querySelector('.formula-toolbar');
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        if (!toolbar || !strip) return null;
        toolbar.style.width = '420px';
        toolbar.style.maxWidth = '420px';
        toolbar.style.justifySelf = 'start';
        const waitFrames = (count = 2) => new Promise((resolve) => {
          const step = () => {
            if (count-- <= 0) resolve();
            else requestAnimationFrame(step);
          };
          requestAnimationFrame(step);
        });
        const wheel = async (delta) => {
          strip.dispatchEvent(new WheelEvent('wheel', {
            deltaY: delta,
            bubbles: true,
            cancelable: true,
          }));
          await waitFrames();
          return strip.scrollLeft;
        };
        const sections = [...strip.querySelectorAll(
          ':scope > .toolbar-category-section',
        )];
        const starts = Object.fromEntries(
          sections.map((section) => [
            section.dataset.toolbarCategorySection,
            section.offsetLeft,
          ]),
        );
        const common = sections.find(
          (section) => section.dataset.toolbarCategorySection === 'common',
        );
        const commonEnd = Math.max(
          common.offsetLeft,
          common.offsetLeft + common.offsetWidth - strip.clientWidth,
        );

        strip.scrollLeft = common.offsetLeft;
        await waitFrames();
        const lowStart = strip.scrollLeft;
        const lowEnd = await wheel(36);

        strip.scrollLeft = commonEnd - 20;
        await waitFrames();
        const boundaryStart = strip.scrollLeft;
        const boundaryEnd = await wheel(240);

        strip.scrollLeft = 0;
        await waitFrames();
        const multiEnd = await wheel(6000);
        const crossedCategories = sections.filter(
          (section) => section.offsetLeft < multiEnd,
        ).length;
        const nearestStartDistance = Math.min(
          ...Object.values(starts).map((start) => Math.abs(start - multiEnd)),
        );
        const reverseEnd = await wheel(-720);

        return {
          lowStart,
          lowEnd,
          lowDelta: lowEnd - lowStart,
          boundaryStart,
          boundaryEnd,
          commonEnd,
          crossedBoundary: boundaryEnd > commonEnd + 100,
          multiEnd,
          crossedCategories,
          nearestStartDistance,
          reverseEnd,
          reverseDelta: reverseEnd - multiEnd,
          scrollBehavior: getComputedStyle(strip).scrollBehavior,
          scrollSnapType: getComputedStyle(strip).scrollSnapType,
        };
      })()`);
      assert.ok(wheelState);
      assert.ok(wheelState.lowDelta >= 34 && wheelState.lowDelta <= 38);
      assert.ok(wheelState.crossedBoundary);
      assert.ok(wheelState.crossedCategories >= 3);
      assert.ok(wheelState.nearestStartDistance > 2);
      assert.ok(wheelState.reverseDelta <= -700);
      assert.equal(wheelState.scrollBehavior, 'auto');
      assert.ok(!wheelState.scrollSnapType || wheelState.scrollSnapType === 'none');
      console.log(JSON.stringify(wheelState));
    }

    if (mode === "continuous-wheel-legacy") {
      await evaluate(`(() => {
        const toolbar = document.querySelector('.formula-toolbar');
        toolbar.style.width = '420px';
        toolbar.style.maxWidth = '420px';
        toolbar.style.justifySelf = 'start';
        document.querySelector(
          '.toolbar-tab[data-category="common"]',
        )?.click();
      })()`);
      await sleep(380);

      const insideSetup = await evaluate(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const common = strip?.querySelector(
          '[data-toolbar-category-section="common"]',
        );
        if (!strip || !common) return null;
        const start = common.offsetLeft;
        const end = Math.max(
          start,
          common.offsetLeft + common.offsetWidth - strip.clientWidth,
        );
        strip.scrollLeft = start;
        strip.dispatchEvent(new WheelEvent('wheel', {
          deltaY: 80,
          bubbles: true,
          cancelable: true,
        }));
        return {
          start,
          end,
          width: common.offsetWidth,
          clientWidth: strip.clientWidth,
        };
      })()`);
      assert.ok(insideSetup.width > insideSetup.clientWidth);
      const insideState = await waitForEvaluation(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const common = strip?.querySelector(
          '[data-toolbar-category-section="common"]',
        );
        const start = common?.offsetLeft ?? -1;
        const end = common
          ? Math.max(
              start,
              common.offsetLeft + common.offsetWidth - strip.clientWidth,
            )
          : -1;
        const state = {
          category: strip?.dataset.activeCategory ?? '',
          scrollLeft: strip?.scrollLeft ?? -1,
          start,
          end,
        };
        return {
          ready:
            state.category === 'common' &&
            state.scrollLeft > start + 1 &&
            state.scrollLeft <= end + 1,
          ...state,
        };
      })()`, "wheel scrolls inside an overflowing category");

      await sleep(380);
      const inertiaState = await evaluate(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const common = strip?.querySelector(
          '[data-toolbar-category-section="common"]',
        );
        if (!strip || !common) return null;
        const end = Math.max(
          common.offsetLeft,
          common.offsetLeft + common.offsetWidth - strip.clientWidth,
        );
        strip.scrollLeft = Math.max(common.offsetLeft, end - 24);
        strip.dispatchEvent(new WheelEvent('wheel', {
          deltaY: 80,
          bubbles: true,
          cancelable: true,
        }));
        strip.dispatchEvent(new WheelEvent('wheel', {
          deltaY: 74,
          bubbles: true,
          cancelable: true,
        }));
        strip.dispatchEvent(new WheelEvent('wheel', {
          deltaY: 68,
          bubbles: true,
          cancelable: true,
        }));
        return {
          category: strip.dataset.activeCategory ?? '',
          scrollLeft: strip.scrollLeft,
          end,
        };
      })()`);
      assert.equal(inertiaState.category, 'common');
      assert.ok(Math.abs(inertiaState.scrollLeft - inertiaState.end) <= 1);

      await evaluate(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        strip?.dispatchEvent(new WheelEvent('wheel', {
          deltaY: 80,
          bubbles: true,
          cancelable: true,
        }));
        strip?.dispatchEvent(new WheelEvent('wheel', {
          deltaY: 80,
          bubbles: true,
          cancelable: true,
        }));
      })()`);
      const nextState = await waitForEvaluation(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const structure = strip?.querySelector(
          '[data-toolbar-category-section="structure"]',
        );
        const state = {
          category: strip?.dataset.activeCategory ?? '',
          scrollLeft: strip?.scrollLeft ?? -1,
          structureStart: structure?.offsetLeft ?? -1,
        };
        return {
          ready:
            state.category === 'structure' &&
            Math.abs(state.scrollLeft - state.structureStart) <= 1,
          ...state,
        };
      })()`, "continued deliberate wheel steps align the next category without a pause");

      await sleep(380);
      await evaluate(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        strip?.dispatchEvent(new WheelEvent('wheel', {
          deltaY: -80,
          bubbles: true,
          cancelable: true,
        }));
      })()`);
      const reverseState = await waitForEvaluation(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const common = strip?.querySelector(
          '[data-toolbar-category-section="common"]',
        );
        const commonEnd = common
          ? Math.max(
              common.offsetLeft,
              common.offsetLeft + common.offsetWidth - strip.clientWidth,
            )
          : -1;
        const state = {
          category: strip?.dataset.activeCategory ?? '',
          scrollLeft: strip?.scrollLeft ?? -1,
          commonEnd,
        };
        return {
          ready:
            state.category === 'common' &&
            Math.abs(state.scrollLeft - commonEnd) <= 1,
          ...state,
        };
      })()`, "reverse wheel returns to the previous category end");

      await evaluate(`(() => {
        const toolbar = document.querySelector('.formula-toolbar');
        toolbar.style.width = '1200px';
        toolbar.style.maxWidth = '1200px';
        document.querySelector(
          '.toolbar-tab[data-category="greek"]',
        )?.click();
      })()`);
      await sleep(380);
      const shortSetup = await evaluate(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const greek = strip?.querySelector(
          '[data-toolbar-category-section="greek"]',
        );
        const arrow = strip?.querySelector(
          '[data-toolbar-category-section="arrow"]',
        );
        if (!strip || !greek || !arrow) return null;
        strip.scrollLeft = greek.offsetLeft;
        const state = {
          greekWidth: greek.offsetWidth,
          clientWidth: strip.clientWidth,
          arrowStart: arrow.offsetLeft,
        };
        strip.dispatchEvent(new WheelEvent('wheel', {
          deltaY: 80,
          bubbles: true,
          cancelable: true,
        }));
        return state;
      })()`);
      assert.ok(shortSetup.greekWidth <= shortSetup.clientWidth);
      const shortState = await waitForEvaluation(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const arrow = strip?.querySelector(
          '[data-toolbar-category-section="arrow"]',
        );
        const state = {
          category: strip?.dataset.activeCategory ?? '',
          scrollLeft: strip?.scrollLeft ?? -1,
          arrowStart: arrow?.offsetLeft ?? -1,
        };
        return {
          ready:
            state.category === 'arrow' &&
            Math.abs(state.scrollLeft - state.arrowStart) <= 1,
          ...state,
        };
      })()`, "one wheel aligns the next category when the current one fits");

      console.log(
        JSON.stringify({
          insideState,
          inertiaState,
          nextState,
          reverseState,
          shortState,
        }),
      );
    }

    if (mode === "continuous-performance") {
      const inspectPreviewModes = `(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const sections = [...(strip?.querySelectorAll(
          ':scope > .toolbar-category-section',
        ) ?? [])];
        const modes = Object.fromEntries(
          sections.map((section) => [
            section.dataset.toolbarCategorySection,
            section.dataset.previewMode,
          ]),
        );
        const observedCategories = ['arrow', 'physics', 'set'];
        const staticCategoryDetails = Object.fromEntries(
          observedCategories.map((category) => {
            const section = sections.find(
              (item) => item.dataset.toolbarCategorySection === category,
            );
            const previews = [...(section?.querySelectorAll('.math-preview') ?? [])];
            const insideCount = previews.filter((preview) => {
              const content = preview.querySelector('.math-preview-fit-content');
              const hostBounds = preview.getBoundingClientRect();
              const contentBounds = content?.getBoundingClientRect();
              return Boolean(
                contentBounds &&
                  contentBounds.left >= hostBounds.left - 1 &&
                  contentBounds.right <= hostBounds.right + 1 &&
                  contentBounds.top >= hostBounds.top - 1 &&
                  contentBounds.bottom <= hostBounds.bottom + 1,
              );
            }).length;
            return [category, {
              mode: section?.dataset.previewMode ?? '',
              previewCount: previews.length,
              staticCount: previews.filter(
                (preview) => preview.dataset.staticLayout === 'true',
              ).length,
              measuredCount: previews.filter(
                (preview) => preview.dataset.fitReady === 'true',
              ).length,
              insideCount,
            }];
          }),
        );
        const allPreviews = [...(strip?.querySelectorAll('.math-preview') ?? [])];
        return {
          ready: Boolean(strip) && sections.length === 9,
          modes,
          staticCategoryDetails,
          buttonCount: strip?.querySelectorAll('.template-button').length ?? 0,
          previewCount: allPreviews.length,
          staticPreviewCount: allPreviews.filter(
            (preview) => preview.dataset.staticLayout === 'true',
          ).length,
          measuredPreviewCount: allPreviews.filter(
            (preview) => preview.dataset.fitReady === 'true',
          ).length,
          placeholderCount:
            strip?.querySelectorAll('[data-deferred-preview="true"]').length ?? 0,
        };
      })()`;

      const initialState = await waitForEvaluation(
        inspectPreviewModes,
        "static and measured preview modes",
      );
      assert.equal(initialState.placeholderCount, 0);
      assert.ok(initialState.previewCount >= initialState.buttonCount);
      assert.ok(initialState.staticPreviewCount > 0);
      assert.ok(initialState.measuredPreviewCount < initialState.previewCount);
      for (const category of ['set']) {
        const details = initialState.staticCategoryDetails[category];
        assert.equal(details.mode, 'static');
        assert.ok(details.previewCount > 0);
        assert.equal(details.staticCount, details.previewCount);
        assert.equal(details.measuredCount, 0);
      }

      await evaluate(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const relation = strip?.querySelector(
          '[data-toolbar-category-section="relation"]',
        );
        if (strip && relation) strip.scrollLeft = relation.offsetLeft;
      })()`);
      await sleep(240);
      const shiftedState = await waitForEvaluation(`(() => {
        const state = eval(${JSON.stringify(inspectPreviewModes)});
        return {
          ...state,
          ready:
            state.ready &&
            state.modes.relation === 'full' &&
            state.staticCategoryDetails.set.mode === 'static',
        };
      })()`, "viewport upgrades nearby previews after scroll settles");
      assert.equal(shiftedState.placeholderCount, 0);
      assert.ok(shiftedState.measuredPreviewCount < shiftedState.previewCount);

      await evaluate(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const arrow = strip?.querySelector(
          '[data-toolbar-category-section="arrow"]',
        );
        if (strip && arrow) strip.scrollLeft = arrow.offsetLeft;
      })()`);
      await sleep(240);
      const arrowState = await waitForEvaluation(`(() => {
        const state = eval(${JSON.stringify(inspectPreviewModes)});
        const arrow = state.staticCategoryDetails.arrow;
        return {
          ...state,
          ready:
            state.ready &&
            state.modes.arrow === 'full' &&
            arrow.previewCount > 0 &&
            arrow.staticCount === 0 &&
            arrow.measuredCount === arrow.previewCount &&
            arrow.insideCount === arrow.previewCount,
        };
      })()`, "arrow previews fit inside toolbar cells");
      assert.equal(arrowState.placeholderCount, 0);
      assert.equal(arrowState.staticCategoryDetails.arrow.mode, 'full');
      assert.equal(
        arrowState.staticCategoryDetails.arrow.insideCount,
        arrowState.staticCategoryDetails.arrow.previewCount,
      );

      await evaluate(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const physics = strip?.querySelector(
          '[data-toolbar-category-section="physics"]',
        );
        if (strip && physics) strip.scrollLeft = physics.offsetLeft;
      })()`);
      await sleep(240);
      const physicsState = await waitForEvaluation(`(() => {
        const state = eval(${JSON.stringify(inspectPreviewModes)});
        const physics = state.staticCategoryDetails.physics;
        return {
          ...state,
          ready:
            state.ready &&
            state.modes.physics === 'full' &&
            physics.previewCount > 0 &&
            physics.staticCount === 0 &&
            physics.measuredCount === physics.previewCount &&
            physics.insideCount === physics.previewCount,
        };
      })()`, "physics previews fit inside toolbar cells");
      assert.equal(physicsState.placeholderCount, 0);
      assert.equal(physicsState.staticCategoryDetails.physics.mode, 'full');
      assert.equal(
        physicsState.staticCategoryDetails.physics.insideCount,
        physicsState.staticCategoryDetails.physics.previewCount,
      );
      console.log(JSON.stringify({ initialState, shiftedState, arrowState, physicsState }));
    }

    if (mode === "continuous-performance-legacy") {
      const inspectPreviewWindow = `(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const sections = [...(strip?.querySelectorAll(
          ':scope > .toolbar-category-section',
        ) ?? [])];
        const renderedCategories = sections
          .filter((section) => section.dataset.previewRendered === 'true')
          .map((section) => section.dataset.toolbarCategorySection);
        const deferredCategories = sections
          .filter((section) => section.dataset.previewRendered === 'false')
          .map((section) => section.dataset.toolbarCategorySection);
        const deferredLivePreviews = sections
          .filter((section) => section.dataset.previewRendered === 'false')
          .reduce(
            (count, section) =>
              count + section.querySelectorAll('.math-preview-fit-content').length,
            0,
          );
        return {
          ready: Boolean(strip) && renderedCategories.length > 0,
          renderedCategories,
          deferredCategories,
          buttonCount: strip?.querySelectorAll('.template-button').length ?? 0,
          livePreviewCount:
            strip?.querySelectorAll('.math-preview-fit-content').length ?? 0,
          deferredPlaceholderCount:
            strip?.querySelectorAll('[data-deferred-preview="true"]').length ?? 0,
          deferredLivePreviews,
        };
      })()`;

      const initialState = await waitForEvaluation(
        inspectPreviewWindow,
        "bounded formula preview render window",
      );
      assert.deepEqual(initialState.renderedCategories, ['common', 'structure']);
      assert.ok(initialState.deferredCategories.length >= 6);
      assert.ok(initialState.livePreviewCount < initialState.buttonCount);
      assert.ok(initialState.deferredPlaceholderCount > 0);
      assert.equal(initialState.deferredLivePreviews, 0);

      await evaluate(`document.querySelector(
        '.toolbar-tab[data-category="calculus"]',
      )?.click()`);
      const shiftedState = await waitForEvaluation(`(() => {
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        const sections = [...(strip?.querySelectorAll(
          ':scope > .toolbar-category-section',
        ) ?? [])];
        const renderedCategories = sections
          .filter((section) => section.dataset.previewRendered === 'true')
          .map((section) => section.dataset.toolbarCategorySection);
        const deferredCategories = sections
          .filter((section) => section.dataset.previewRendered === 'false')
          .map((section) => section.dataset.toolbarCategorySection);
        const deferredLivePreviews = sections
          .filter((section) => section.dataset.previewRendered === 'false')
          .reduce(
            (count, section) =>
              count + section.querySelectorAll('.math-preview-fit-content').length,
            0,
          );
        const state = {
          renderedCategories,
          deferredCategories,
          buttonCount: strip?.querySelectorAll('.template-button').length ?? 0,
          livePreviewCount:
            strip?.querySelectorAll('.math-preview-fit-content').length ?? 0,
          deferredPlaceholderCount:
            strip?.querySelectorAll('[data-deferred-preview="true"]').length ?? 0,
          deferredLivePreviews,
        };
        return {
          ready:
            Boolean(strip) &&
            JSON.stringify(renderedCategories) ===
              JSON.stringify(['structure', 'calculus', 'matrix']) &&
            deferredLivePreviews === 0,
          ...state,
        };
      })()`, "preview render window follows the active category");
      assert.ok(shiftedState.livePreviewCount < shiftedState.buttonCount);
      console.log(JSON.stringify({ initialState, shiftedState }));
    }

    if (mode === "continuous-fps") {
      const fpsState = await evaluate(`(async () => {
        const toolbar = document.querySelector('.formula-toolbar');
        const strip = document.querySelector(
          '.template-strip.is-continuous-categories',
        );
        if (!toolbar || !strip) return null;
        toolbar.style.width = '960px';
        toolbar.style.maxWidth = '960px';
        toolbar.style.justifySelf = 'start';
        await document.fonts.ready;

        const percentile = (values, ratio) => {
          const sorted = [...values].sort((left, right) => left - right);
          return sorted[Math.min(
            sorted.length - 1,
            Math.floor(sorted.length * ratio),
          )];
        };
        const run = (speed, measuredFrames = 150) => new Promise((resolve) => {
          const intervals = [];
          let previous = 0;
          let frameCount = 0;
          let direction = 1;
          strip.scrollLeft = 0;
          const step = (now) => {
            if (previous > 0 && frameCount >= 20) intervals.push(now - previous);
            previous = now;
            const maximum = Math.max(0, strip.scrollWidth - strip.clientWidth);
            if (strip.scrollLeft >= maximum - speed * 2) direction = -1;
            if (strip.scrollLeft <= speed * 2) direction = 1;
            strip.dispatchEvent(new WheelEvent('wheel', {
              deltaY: speed * direction,
              bubbles: true,
              cancelable: true,
            }));
            frameCount += 1;
            if (intervals.length < measuredFrames) {
              requestAnimationFrame(step);
              return;
            }
            requestAnimationFrame(() => {
              const averageInterval = intervals.reduce(
                (sum, value) => sum + value,
                0,
              ) / intervals.length;
              const p90Interval = percentile(intervals, 0.9);
              const p95Interval = percentile(intervals, 0.95);
              resolve({
                speed,
                frames: intervals.length,
                averageInterval,
                averageFps: 1000 / averageInterval,
                p90Interval,
                p90Fps: 1000 / p90Interval,
                p95Interval,
                p95Fps: 1000 / p95Interval,
                framesOver20ms: intervals.filter((value) => value > 20).length,
                finalScrollLeft: strip.scrollLeft,
              });
            });
          };
          requestAnimationFrame(step);
        });

        const idle = await run(0);
        await new Promise((resolve) => setTimeout(resolve, 180));
        const low = await run(24);
        await new Promise((resolve) => setTimeout(resolve, 180));
        const medium = await run(96);
        await new Promise((resolve) => setTimeout(resolve, 180));
        const high = await run(360);
        return {
          idle,
          low,
          medium,
          high,
          refreshRateTarget: 60,
          minimumAverageFps: 50,
          minimumP90Fps: 45,
        };
      })()`);
      assert.ok(fpsState);
      console.log(JSON.stringify(fpsState));
      for (const result of [fpsState.low, fpsState.medium, fpsState.high]) {
        assert.ok(
          result.averageFps >= fpsState.minimumAverageFps,
          `Average FPS ${result.averageFps.toFixed(2)} below ${fpsState.minimumAverageFps} at speed ${result.speed}`,
        );
        assert.ok(
          result.p90Fps >= fpsState.minimumP90Fps,
          `P90 FPS ${result.p90Fps.toFixed(2)} below ${fpsState.minimumP90Fps} at speed ${result.speed}`,
        );
      }
    }

    if (mode === "piecewise-preview") {
      const state = await waitForEvaluation(`(() => {
        const commonSection = document.querySelector(
          '[data-toolbar-category-section="common"]',
        );
        const structureSection = document.querySelector(
          '[data-toolbar-category-section="structure"]',
        );
        const twoRow = commonSection?.querySelector(
          '.template-button[data-command-id="cases"]',
        );
        const threeRow = structureSection?.querySelector(
          '.template-button[data-command-id="cases-three"]',
        );
        const inspect = (button) => {
          const host = button?.querySelector('.math-preview');
          const content = host?.querySelector('.math-preview-fit-content');
          const buttonRect = button?.getBoundingClientRect();
          const hostRect = host?.getBoundingClientRect();
          const contentRect = content?.getBoundingClientRect();
          return {
            className: button?.className ?? '',
            buttonHeight: buttonRect?.height ?? 0,
            hostHeight: hostRect?.height ?? 0,
            contentHeight: contentRect?.height ?? 0,
            contentWidth: contentRect?.width ?? 0,
            inside: Boolean(
              hostRect &&
              contentRect &&
              contentRect.left >= hostRect.left - 1 &&
              contentRect.right <= hostRect.right + 1 &&
              contentRect.top >= hostRect.top - 1 &&
              contentRect.bottom <= hostRect.bottom + 1
            ),
            fitReady: host?.dataset.fitReady === 'true',
            scale: Number(host?.dataset.fitScale ?? 0),
          };
        };
        const two = inspect(twoRow);
        const three = inspect(threeRow);
        const normalHeight = structureSection
          ?.querySelector('.template-button:not([data-command-id="cases-three"])')
          ?.getBoundingClientRect().height ?? 0;
        return {
          ready:
            two.fitReady &&
            three.fitReady &&
            two.inside &&
            three.inside &&
            two.className.includes('is-enlarged-cases-preview') &&
            three.className.includes('is-enlarged-cases-preview') &&
            Math.abs(two.buttonHeight - normalHeight) <= 1 &&
            Math.abs(three.buttonHeight - normalHeight) <= 1 &&
            three.contentHeight >= normalHeight * 0.72,
          normalHeight,
          two,
          three,
        };
      })()`, "large piecewise previews");
      console.log(JSON.stringify(state));
    }

    console.log(`Focused toolbar regression passed: ${mode}`);
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(120);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => {});
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
