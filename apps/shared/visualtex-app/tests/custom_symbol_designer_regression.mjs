// The custom symbol designer (apps/shared/custom-symbol-designer) in the built app, headless Chrome with a stub
// Tauri bridge: materials come in at their real size on the baseline, the TeX code follows the drawing, undo and
// redo, registering a command, and the designer surviving a narrow window. Run after `npm run build:desktop`.
import assert from "node:assert/strict";
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const stub = `(() => {
  const callbacks = new Map(); let next = 1;
  globalThis.isTauri = true; // set by the Tauri runtime (isTauri())
  window.__TAURI_INTERNALS__ = {
    transformCallback(cb) { const id = next++; callbacks.set(id, cb); return id; },
    unregisterCallback(id) { callbacks.delete(id); },
    async invoke(cmd, args) { if (cmd === "plugin:event|listen") return args.handler; return null; },
    metadata: { currentWindow: { label: "main" }, currentWebview: { label: "main", windowLabel: "main" } },
    convertFileSrc: (p) => p,
  };
  window.__TAURI_EVENT_PLUGIN_INTERNALS__ = { unregisterListener() {} };
})();`;

await withAppBrowser(async ({ baseUrl: base, client, send, evaluate }) => {
  const viewport = (width, height) => send("Emulation.setDeviceMetricsOverride", { width, height, deviceScaleFactor: 1, mobile: false });
  const click = (selector) => evaluate(`(() => { const el = document.querySelector(${JSON.stringify(selector)}); if (!el) throw new Error("missing ${selector.replace(/"/g, "'")}"); el.click(); })()`);
  const setValue = (selector, value) => evaluate(`(() => {
    const el = document.querySelector(${JSON.stringify(selector)});
    if (!el) throw new Error("missing input");
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value").set.call(el, ${JSON.stringify(value)});
    el.dispatchEvent(new Event("input", { bubbles: true }));
  })()`);
  const addLatex = async (latex) => {
    await click('[data-csd-material-tab="latex"]');
    await setValue("[data-csd-source-input]", latex);
    await evaluate(`document.querySelector("[data-csd-source-input]").form.requestSubmit()`);
    await sleep(700);
  };
  const tex = () => evaluate(`(() => { document.querySelector('[data-csd-code-tab="tex"]').click(); return document.querySelector('[data-csd-code-text="tex"]').textContent; })()`);
  const layerCount = () => evaluate(`document.querySelectorAll("[data-csd-artwork]").length`);

  await viewport(1440, 900);
  await send("Page.addScriptToEvaluateOnNewDocument", { source: stub });
  await send("Page.navigate", { url: `${base}/index.html?transport=tauri&theme=light` });
  await sleep(4000);
  await click('[data-tile-category="custom"]');
  await sleep(400);
  await click("[data-open-custom-symbol-designer]");
  await sleep(1500);
  assert.ok(await evaluate(`Boolean(document.querySelector("[data-csd-dialog]"))`), "the designer opens");

  // ħ comes in at its real size (1000 units per em), its box fitting the glyph, its baseline on the baseline
  await addLatex("\\hbar");
  const hbar = await evaluate(`(() => {
    const points = document.querySelector("[data-csd-selection] polygon").getAttribute("points").split(" ").map((p) => p.split(",").map(Number));
    const ys = points.map((p) => p[1]);
    const xs = points.map((p) => p[0]);
    const baseline = Number(document.querySelector(".csd-guide.is-baseline line").getAttribute("y1"));
    return { height: Math.max(...ys) - Math.min(...ys), width: Math.max(...xs) - Math.min(...xs), bottom: Math.max(...ys), baseline };
  })()`);
  assert.ok(Math.abs(hbar.height - 705.9) < 2, `ħ box height ${hbar.height} (TeX: 0.695 + 0.011 em)`);
  assert.ok(Math.abs(hbar.width - 546.3) < 2, `ħ box width ${hbar.width}`);
  assert.ok(Math.abs(hbar.bottom - (hbar.baseline + 11)) < 1.5, "ħ sits on the baseline");
  assert.ok(await evaluate(`Boolean(document.querySelector("[data-csd-reference] path"))`), "the reference glyph is drawn");
  assert.match(await tex(), /\{\$\\hbar\$\}/, "the TeX code typesets the material");

  // a shape, then undo and redo
  await click('[data-csd-material-tab="shapes"]');
  await click('[data-csd-shape="circle"]');
  await sleep(300);
  assert.equal(await layerCount(), 2);
  assert.match(await tex(), /\\strokepath/, "the circle is a pict2e path in the TeX code");
  assert.match(await tex(), /\\begin\{picture\}/, "the command is a picture");
  assert.doesNotMatch(await tex(), /tikz/i, "no TikZ");
  await click("[data-csd-undo]");
  await sleep(200);
  assert.equal(await layerCount(), 1, "undo removes the circle");
  await click("[data-csd-redo]");
  await sleep(200);
  assert.equal(await layerCount(), 2, "redo brings it back");

  // register as \regtestsym
  await click('[data-csd-side-tab="register"]');
  await setValue("[data-csd-command]", "regtestsym");
  await sleep(200);
  await click("[data-csd-register]");
  await sleep(500);
  assert.equal(await evaluate(`document.querySelector("[data-csd-register-status]")?.getAttribute("data-csd-register-status")`), "success");
  await click("[data-csd-library-button]");
  await sleep(200);
  assert.ok(await evaluate(`Boolean(document.querySelector('[data-csd-symbol="regtestsym"]'))`), "the symbol is in the library");

  // a narrow window keeps the designer and its drawing
  await viewport(620, 760);
  await sleep(800);
  assert.ok(await evaluate(`Boolean(document.querySelector("[data-csd-dialog]"))`), "the designer stays open in a narrow window");
  assert.equal(await layerCount(), 2);
  await viewport(1440, 900);

  const errors = client.events
    .filter((event) => event.method === "Runtime.exceptionThrown")
    .map((event) => event.params.exceptionDetails.exception?.description ?? "exception");
  assert.deepEqual(errors, [], "no page errors");
  console.log("custom symbol designer: ok");
}, { path: "about:blank", windowSize: null });
