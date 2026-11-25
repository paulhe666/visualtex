// A formula opened with "$$" (inline) or "$$$$" (display) and left before anything was typed in it stays in the document,
// as its source keeps the dollars, however it is left: the arrow keys, Tab, Enter, Ctrl+Enter, a click elsewhere. (An
// inline formula left empty used to go away: what was typed vanished, 2026-11-24.) Backspace in an empty formula still
// removes it.
import assert from "node:assert/strict";
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// each case in a browser of its own (a page saves its edited document as it unloads, over the one a test stores)
const session = (run) => withAppBrowser(async ({ baseUrl, client }) => {
  const evaluate = async (expression) => {
    const result = await client.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
    if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
    return result.result.value;
  };
  for (let tries = 0; tries < 300 && !String(await evaluate("location.href")).startsWith(baseUrl); tries += 1) await sleep(50);
  const key = async (value, code, vk, modifiers = 0) => {
    const common = { key: value, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, modifiers };
    const text = value.length === 1 ? { text: value, unmodifiedText: value } : {};
    await client.send("Input.dispatchKeyEvent", { type: value.length === 1 ? "keyDown" : "rawKeyDown", ...common, ...text });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
    await sleep(80);
  };
  const KEYS = {
    Right: ["ArrowRight", "ArrowRight", 39], Left: ["ArrowLeft", "ArrowLeft", 37], Up: ["ArrowUp", "ArrowUp", 38],
    Down: ["ArrowDown", "ArrowDown", 40], Tab: ["Tab", "Tab", 9], Enter: ["Enter", "Enter", 13], Backspace: ["Backspace", "Backspace", 8],
  };
  const press = (name, modifiers = 0) => key(...KEYS[name], modifiers);
  const dollars = async (count) => {
    for (let i = 0; i < count; i += 1) await key("$", "Digit4", 52, 8);
    await sleep(900);
  };
  // shown: formulas at rest whose dashed box is on the page (the picture shown, its dashed outline drawn, the resting
  // editor out of the way: an empty formula left after it was opened showed nothing at all)
  const formulas = () => evaluate(`({
    inline: document.querySelectorAll(".vtx-document .vtx-math-inline").length,
    display: document.querySelectorAll(".vtx-document .vtx-math-display").length,
    open: document.querySelectorAll(".vtx-document .is-open").length,
    shown: [...document.querySelectorAll(".vtx-document .vtx-math-inline:not(.is-open), .vtx-document .vtx-math-display:not(.is-open)")].filter((f) => {
      const picture = f.querySelector(".vtx-formula-static");
      const editor = f.querySelector(".vtx-formula-editor");
      const box = picture && getComputedStyle(picture, "::before");
      return picture && getComputedStyle(picture).display !== "none" && box.outlineStyle === "dashed" && parseFloat(box.width) > 0 &&
        (!editor || getComputedStyle(editor).display === "none");
    }).length })`);
  for (let round = 0; round < 2; round += 1) {
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
    for (let t = 0; t < 100 && !(await evaluate(`Boolean(document.querySelector(".vtx-document p")?.getBoundingClientRect().width)`)); t += 1) await sleep(50);
    await sleep(400);
  }
  const clickEndOf = async (index) => {
    const at = await evaluate(`(() => { const p = document.querySelectorAll(".vtx-document p")[${index}]; const r = p.getBoundingClientRect(); return { x: r.right - 4, y: r.top + r.height / 2 }; })()`);
    for (const type of ["mousePressed", "mouseReleased"]) await client.send("Input.dispatchMouseEvent", { type, x: at.x, y: at.y, button: "left", clickCount: 1 });
    await sleep(300);
    await key("End", "End", 35);
  };
  const clickFormula = async () => {
    const at = await evaluate(`(() => { const f = document.querySelector(".vtx-document .vtx-math-inline, .vtx-document .vtx-math-display"); const r = f.querySelector(".vtx-formula-static").getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()`);
    for (const type of ["mousePressed", "mouseReleased"]) await client.send("Input.dispatchMouseEvent", { type, x: at.x, y: at.y, button: "left", clickCount: 1 });
    await sleep(600);
  };
  await run({ press, dollars, formulas, clickEndOf, clickFormula });
}, { windowSize: "1100,900" });

const CTRL = 2;
const leaves = [
  ["Right", (t) => t.press("Right")],
  ["Left", (t) => t.press("Left")],
  ["Up", (t) => t.press("Up")],
  ["Down", (t) => t.press("Down")],
  ["Tab", (t) => t.press("Tab")],
  ["Ctrl+Enter", (t) => t.press("Enter", CTRL)],
  ["a click elsewhere", (t) => t.clickEndOf(1)],
];
const cases = [
  ...leaves.map(([how, leave]) => ({ name: `inline, ${how}`, dollars: 2, leave, expect: { inline: 1, display: 0 } })),
  { name: "inline, Enter", dollars: 2, leave: (t) => t.press("Enter"), expect: { inline: 1, display: 0 } },
  ...leaves.map(([how, leave]) => ({ name: `display, ${how}`, dollars: 4, leave, expect: { inline: 0, display: 1 } })),
];
for (const test of cases) {
  await session(async (t) => {
    await t.clickEndOf(0);
    await t.dollars(test.dollars);
    const opened = await t.formulas();
    assert.equal(opened.open, 1, `${test.name}: the formula opened (${JSON.stringify(opened)})`);
    await test.leave(t);
    await sleep(700);
    const left = await t.formulas();
    console.log(`${test.name}: ${JSON.stringify(left)}`);
    assert.equal(left.open, 0, `${test.name}: the formula was left`);
    assert.equal(left.inline, test.expect.inline, `${test.name}: inline formulas`);
    assert.equal(left.display, test.expect.display, `${test.name}: display formulas`);
    assert.equal(left.shown, 1, `${test.name}: the empty formula shows its dashed box`);
    // a click on the dashed box opens it again
    if (test.name.endsWith("Right")) {
      await t.clickFormula();
      const reopened = await t.formulas();
      assert.equal(reopened.open, 1, `${test.name}: a click on the dashed box opens the formula (${JSON.stringify(reopened)})`);
    }
  });
}

// Backspace in an empty formula still removes it
await session(async (t) => {
  await t.clickEndOf(0);
  await t.dollars(2);
  await t.press("Backspace");
  await sleep(600);
  const removed = await t.formulas();
  assert.equal(removed.inline, 0, `Backspace in the empty formula removes it (${JSON.stringify(removed)})`);
});

console.log("empty formulas left stay; Backspace removes one: passed");
