// The formula input method's shortcut pressed while the main window's document has the keyboard (the backend's event
// visualtex-formula-shortcut, documentFormulaShortcut.ts): the formula opens at the caret and is typed in English (the
// backend asked: anywhere_own_formula_input english true) until no formula is open (english false), an arrow from one
// formula into the next keeping it; the window losing the keyboard gives the input method back too.
import assert from "node:assert/strict";
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// the page's Tauri side: events in, commands recorded; the backend switches (true) whenever asked to
const STAND_IN = `
(() => {
  const handlers = new Map();
  let next = 1;
  window.__calls = [];
  window.__listeners = [];
  window.isTauri = true;
  window.__TAURI_INTERNALS__ = {
    metadata: { currentWindow: { label: "main" }, currentWebview: { label: "main", windowLabel: "main" } },
    transformCallback(callback) { const id = next++; handlers.set(id, callback); return id; },
    unregisterCallback(id) { handlers.delete(id); },
    convertFileSrc: (path) => path,
    async invoke(command, args) {
      if (command === "plugin:event|listen") { window.__listeners.push([args.event, args.handler]); return args.handler; }
      if (command === "anywhere_own_formula_input") { window.__calls.push(args.english ? "english" : "back"); return Boolean(args.english); }
      if (command === "anywhere_english_input") return true;
      return null;
    },
  };
  window.__TAURI_EVENT_PLUGIN_INTERNALS__ = { unregisterListener() {} };
  window.__emit = (event, payload) => { for (const [name, id] of window.__listeners) if (name === event) handlers.get(id)?.({ event, id, payload }); };
})();
`;

const session = (run) => withAppBrowser(async ({ baseUrl, client }) => {
  const evaluate = async (expression) => {
    const result = await client.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
    if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
    return result.result.value;
  };
  await client.send("Page.addScriptToEvaluateOnNewDocument", { source: STAND_IN });
  for (let tries = 0; tries < 300 && !String(await evaluate("location.href")).startsWith(baseUrl); tries += 1) await sleep(50);
  await evaluate(`(() => {
    for (const k of ["visualtex.onboarding.v3.completed", "visualtex.office.macos.first-run.v1.completed", "visualtex.onboarding.macos.desktop.v1.2.0.completed", "visualtex.office.macos.native-first-run.v1.2.0.completed"]) localStorage.setItem(k, "true");
    localStorage.setItem("visualtex.document.tutorial-shown.v1", "1");
    localStorage.setItem("visualtex.workspace-content.v1", JSON.stringify({ state: { content: "document" }, version: 0 }));
    localStorage.setItem("visualtex.document.autosave.v1", JSON.stringify({ doc: { type: "doc", content: [
      { type: "paragraph", content: [{ type: "text", text: "first line" }] },
      { type: "paragraph", content: [{ type: "text", text: "second line" }] } ] } }));
  })()`);
  await client.send("Page.reload", { ignoreCache: true });
  await sleep(800);
  for (let t = 0; t < 200 && !(await evaluate(`Boolean(document.querySelector(".vtx-document p")?.getBoundingClientRect().width)`)); t += 1) await sleep(50);
  await sleep(500);
  const key = async (value, code, vk, modifiers = 0) => {
    const common = { key: value, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, modifiers };
    const text = value.length === 1 ? { text: value, unmodifiedText: value } : {};
    await client.send("Input.dispatchKeyEvent", { type: value.length === 1 ? "keyDown" : "rawKeyDown", ...common, ...text });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
    await sleep(80);
  };
  const clickEndOf = async (index) => {
    const at = await evaluate(`(() => { const p = document.querySelectorAll(".vtx-document p")[${index}]; const r = p.getBoundingClientRect(); return { x: r.right - 4, y: r.top + r.height / 2 }; })()`);
    for (const type of ["mousePressed", "mouseReleased"]) await client.send("Input.dispatchMouseEvent", { type, x: at.x, y: at.y, button: "left", clickCount: 1 });
    await sleep(300);
    await key("End", "End", 35);
  };
  const state = () => evaluate(`({ calls: window.__calls.slice(), open: document.querySelectorAll(".vtx-document .is-open").length,
    inline: document.querySelectorAll(".vtx-document .vtx-math-inline").length, display: document.querySelectorAll(".vtx-document .vtx-math-display").length,
    listening: window.__listeners.some(([name]) => name === "visualtex-formula-shortcut") })`);
  const shortcut = async (display) => {
    await evaluate(`window.__emit("visualtex-formula-shortcut", { display: ${display} })`);
    await sleep(900);
  };
  await run({ evaluate, key, clickEndOf, state, shortcut });
}, { windowSize: "1100,900" });

// the inline shortcut: English while the formula is open, back after Right leaves it
await session(async (t) => {
  await t.clickEndOf(0);
  const before = await t.state();
  assert.ok(before.listening, `the document listens for the shortcut (${JSON.stringify(before)})`);
  await t.shortcut(false);
  const opened = await t.state();
  console.log("inline shortcut:", JSON.stringify(opened));
  assert.equal(opened.open, 1, "the formula opened");
  assert.equal(opened.inline, 1, "an inline formula");
  assert.deepEqual(opened.calls, ["english"], "English asked for");
  await t.key("ArrowRight", "ArrowRight", 39);
  await sleep(500);
  const left = await t.state();
  console.log("after Right:", JSON.stringify(left));
  assert.equal(left.open, 0, "the formula was left");
  assert.deepEqual(left.calls, ["english", "back"], "the input method given back");
});

// the display shortcut, then the window loses the keyboard: given back with the formula still open
await session(async (t) => {
  await t.clickEndOf(0);
  await t.shortcut(true);
  const opened = await t.state();
  console.log("display shortcut:", JSON.stringify(opened));
  assert.equal(opened.display, 1, "a display formula");
  assert.deepEqual(opened.calls, ["english"], "English asked for");
  await t.evaluate(`window.dispatchEvent(new Event("blur"))`);
  await sleep(300);
  const blurred = await t.state();
  console.log("after blur:", JSON.stringify(blurred));
  assert.deepEqual(blurred.calls, ["english", "back"], "the input method given back");
});

// typing "$$" asks for nothing; the shortcut in the open formula does nothing
await session(async (t) => {
  await t.clickEndOf(0);
  for (let i = 0; i < 2; i += 1) await t.key("$", "Digit4", 52, 8);
  await sleep(900);
  const typed = await t.state();
  console.log("typed $$:", JSON.stringify(typed));
  assert.equal(typed.open, 1, "the typed formula opened");
  await t.shortcut(false);
  const again = await t.state();
  assert.equal(again.inline, 1, "no second formula");
  assert.deepEqual(again.calls, [], "no English for a formula typed with dollars, nor for the shortcut in it");
});

console.log("shortcut formulas typed in English until left: passed");
