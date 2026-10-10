import assert from "node:assert/strict";
import { withBrowserTest } from "./browser_test_runtime.mjs";

// Responsive layout guards for the macOS desktop and Office editors:
// - the bottom dock never paints buttons over each other at any window width;
// - panel tabs collapse to icons when the dock itself is narrow;
// - the main menu does not move after it appears;
// - the formula color picker escapes the scrolling formatting strip.

const sessionId = "11111111-2222-4333-8444-555555555555";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const officeFixture = `(() => {
  const fixture = { id: '${sessionId}', mode: 'edit', host: 'powerpoint', formulaId: 'f', sourceDocumentId: 'd',
    sourceObjectId: 'o', title: 'layout', lines: [{ id: 'office-line', latex: 'x^2+y^2=z^2' }], activeLineId: 'office-line',
    codeFormat: 'raw', displayMode: 'block', numbered: false, fontSizePt: 12, exportWidth: 320, exportHeight: 80,
    exportResult: null, originalMetadata: null, dirty: false, status: 'editing', autoCommitOnClose: false,
    explicitCancel: false, error: null, createdAt: Date.now(), updatedAt: Date.now(), expiresAt: Date.now() + 600000 };
  const originalFetch = window.fetch.bind(window);
  window.fetch = async (input, init = {}) => {
    const url = new URL(typeof input === 'string' ? input : input.url, location.href);
    if (url.pathname === '/api/v1/sessions/${sessionId}') {
      if (String(init.method || 'GET').toUpperCase() === 'PATCH') Object.assign(fixture, JSON.parse(init.body || '{}'));
      return new Response(JSON.stringify(fixture), { status: 200, headers: { 'Content-Type': 'application/json' } });
    }
    return originalFetch(input, init);
  };
})();`;
const firstRunDone = `try { localStorage.setItem("visualtex.office.macos.native-first-run.v1.2.0.completed", "true"); } catch {}`;

const dockReport = `(() => {
  const dock = document.querySelector('.classic-bottom-tabs');
  if (!dock) return null;
  const visible = (element) => {
    let { left, top, right, bottom } = element.getBoundingClientRect();
    for (let parent = element.parentElement; parent && parent !== document.body; parent = parent.parentElement) {
      const style = getComputedStyle(parent);
      if (!/(hidden|clip|auto|scroll)/.test(style.overflowX + style.overflowY)) continue;
      const box = parent.getBoundingClientRect();
      left = Math.max(left, box.left); top = Math.max(top, box.top);
      right = Math.min(right, box.right); bottom = Math.min(bottom, box.bottom);
    }
    return { left, top, right, bottom };
  };
  const buttons = [...dock.querySelectorAll('button')].filter((button) => {
    const box = visible(button);
    return box.right - box.left > 2 && box.bottom - box.top > 2;
  });
  const overlaps = [];
  for (let i = 0; i < buttons.length; i++) for (let j = i + 1; j < buttons.length; j++) {
    const a = visible(buttons[i]); const b = visible(buttons[j]);
    const width = Math.min(a.right, b.right) - Math.max(a.left, b.left);
    const height = Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top);
    if (width > 1 && height > 1) overlaps.push((buttons[i].getAttribute('aria-label') || buttons[i].textContent) + ' / ' + (buttons[j].getAttribute('aria-label') || buttons[j].textContent));
  }
  const label = dock.querySelector('.classic-bottom-tab-label');
  return { dockWidth: dock.clientWidth, overlaps, labelVisible: Boolean(label && label.getBoundingClientRect().width > 0) };
})()`;

await withBrowserTest(async ({ baseUrl, send, evaluate }) => {
  const viewport = async (width, height) => {
    await send("Emulation.setDeviceMetricsOverride", { width, height, deviceScaleFactor: 1, mobile: false });
    await sleep(350);
  };
  const waitFor = async (expression, label, timeout = 10000) => {
    const started = Date.now();
    while (Date.now() - started < timeout) {
      if (await evaluate(expression)) return;
      await sleep(60);
    }
    throw new Error(`Timed out waiting for ${label}`);
  };
  const { identifier: firstRunScript } = await send("Page.addScriptToEvaluateOnNewDocument", { source: firstRunDone });

  // Desktop window.
  await viewport(1280, 820);
  await send("Page.navigate", { url: baseUrl });
  await waitFor("Boolean(document.querySelector('.classic-bottom-tabs .classic-bottom-tab-group'))", "desktop dock");
  await evaluate("document.querySelector('.office-first-run-backdrop')?.remove(), true");
  for (const width of [1440, 1280, 1100, 1000, 900, 800, 700, 600, 500]) {
    await viewport(width, 820);
    const report = await evaluate(dockReport);
    assert.deepEqual(report.overlaps, [], `desktop ${width}px: dock buttons overlap`);
    assert.equal(report.labelVisible, report.dockWidth >= 940, `desktop ${width}px: tab labels should collapse below a 940px dock (dock ${report.dockWidth}px)`);
  }

  // Main menu: the left edge must not move while it opens.
  for (const width of [1280, 600]) {
    await viewport(width, 820);
    const trace = await evaluate(`new Promise((resolve) => {
      const lefts = [];
      const started = performance.now();
      const tick = () => {
        const menu = document.querySelector('.app-menu-popover');
        if (menu) lefts.push(menu.getBoundingClientRect().left);
        if (performance.now() - started < 400) requestAnimationFrame(tick);
        else resolve(lefts);
      };
      document.querySelector('.menu-button').click();
      requestAnimationFrame(tick);
    })`);
    assert.ok(trace.length > 3, `desktop ${width}px: menu did not open`);
    const drift = Math.max(...trace) - Math.min(...trace);
    assert.ok(drift <= 0.5, `desktop ${width}px: main menu moved ${drift.toFixed(2)}px after opening`);
    assert.ok(Math.min(...trace) >= 9.5, `desktop ${width}px: main menu leaves the safe inset`);
    await evaluate("document.querySelector('.menu-button').click(), true");
    await sleep(150);
  }

  // Color picker escapes the scrolling strip and stays on screen.
  await viewport(700, 820);
  await evaluate("document.querySelector('[data-formula-selection-color]').click(), true");
  await waitFor("Boolean(document.querySelector('[data-formula-color-popover]'))", "color picker");
  const picker = await evaluate(`(() => {
    const popover = document.querySelector('[data-formula-color-popover]');
    const box = popover.getBoundingClientRect();
    return { parent: popover.parentElement === document.body, position: getComputedStyle(popover).position,
      inside: box.left >= 0 && box.right <= innerWidth && box.top >= 0 && box.bottom <= innerHeight };
  })()`);
  assert.deepEqual(picker, { parent: true, position: "fixed", inside: true });

  // Office editor window.
  await send("Page.removeScriptToEvaluateOnNewDocument", { identifier: firstRunScript });
  await send("Page.addScriptToEvaluateOnNewDocument", { source: firstRunDone + officeFixture });
  await viewport(900, 620);
  await send("Page.navigate", { url: `${baseUrl}/office-native-dialog.html?sessionId=${sessionId}&officeHost=powerpoint` });
  await waitFor("Boolean(document.querySelector('.office-dialog-shell .classic-bottom-tabs'))", "Office dock");
  for (const width of [1100, 900, 800, 700, 600, 520, 500]) {
    await viewport(width, 620);
    const report = await evaluate(dockReport);
    assert.deepEqual(report.overlaps, [], `Office ${width}px: dock buttons overlap`);
    const editor = await evaluate("document.querySelector('.office-dialog-shell .mathfield-host')?.getBoundingClientRect().width ?? 0");
    const tiles = await evaluate("document.querySelector('.office-dialog-shell .classic-tile-toolbar')?.getBoundingClientRect().width ?? 0");
    assert.ok(!tiles || editor >= tiles * 0.9, `Office ${width}px: canvas (${Math.round(editor)}px) squeezed below the tile rail (${Math.round(tiles)}px)`);
  }
});

console.log("macOS responsive layout regression passed.");
