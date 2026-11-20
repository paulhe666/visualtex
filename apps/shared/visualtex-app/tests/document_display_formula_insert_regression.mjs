// A display formula typed at the end of a line ("$$$$") stands under that line with nothing added after it: no empty
// line between it and the block that follows (it used to add one whenever the next block was not a text line, so a
// formula typed above another formula left a blank line between them). At the document's end, Right out of the
// formula still makes a line to go on writing.
import assert from "node:assert/strict";
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

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
    await sleep(60);
  };
  // the blocks of the document: kind and text
  const blocks = () => evaluate(`[...document.querySelector(".vtx-document").children]
    .filter((n) => !n.classList.contains("ProseMirror-widget") && !n.matches(".vtx-page-spacer, [data-page-spacer]"))
    .map((n) => n.classList.contains("vtx-math-display") ? "display" : n.tagName === "P" ? "p:" + n.textContent.replace(/\\u200b/g, "") : n.tagName.toLowerCase())`);

  // (twice: the page saves its document as it unloads, over the one just stored; the second time that is this one)
  const open = async (content) => {
    for (let round = 0; round < 2; round += 1) {
      await evaluate(`(() => {
        for (const k of ["visualtex.onboarding.v3.completed", "visualtex.office.macos.first-run.v1.completed", "visualtex.onboarding.macos.desktop.v1.2.0.completed", "visualtex.office.macos.native-first-run.v1.2.0.completed"]) localStorage.setItem(k, "true");
        localStorage.setItem("visualtex.document.tutorial-shown.v1", "1");
        localStorage.setItem("visualtex.workspace-content.v1", JSON.stringify({ state: { content: "document" }, version: 0 }));
        localStorage.setItem("visualtex.document.autosave.v1", JSON.stringify({ doc: { type: "doc", content: ${JSON.stringify(content)} } }));
      })()`);
      await client.send("Page.reload", { ignoreCache: true });
      await sleep(800);
      for (let t = 0; t < 100 && !(await evaluate(`Boolean(document.querySelector(".vtx-document p")?.getBoundingClientRect().width)`)); t += 1) await sleep(50);
      await sleep(400);
    }
  };
  // the caret at the end of the first paragraph (a click right of its text)
  const caretAtEndOfFirstLine = async () => {
    const at = await evaluate(`(() => { const p = document.querySelector(".vtx-document p"); const r = p.getBoundingClientRect(); return { x: r.right - 4, y: r.top + r.height / 2 }; })()`);
    for (const type of ["mousePressed", "mouseReleased"]) await client.send("Input.dispatchMouseEvent", { type, x: at.x, y: at.y, button: "left", clickCount: 1 });
    await sleep(250);
    await key("End", "End", 35);
  };
  const typeDollars = async () => {
    for (let i = 0; i < 4; i += 1) await key("$", "Digit4", 52, 8);
    await sleep(900);
  };

  await run({ open, caretAtEndOfFirstLine, typeDollars, blocks, evaluate, key });
}, { windowSize: "1100,900" });

// above another display formula
await session(async ({ open, caretAtEndOfFirstLine, typeDollars, blocks, evaluate }) => {
  await open([
    { type: "paragraph", content: [{ type: "text", text: "积分形式：" }] },
    { type: "math_display", attrs: { latex: "\\frac{\\partial}{\\partial t}\\iiint_V \\rho\\,dV=0" } },
    { type: "paragraph", content: [{ type: "text", text: "after" }] },
  ]);
  await caretAtEndOfFirstLine();
  await typeDollars();
  const above = await blocks();
  console.log("above a formula:", JSON.stringify(above));
  assert.ok(await evaluate(`Boolean(document.querySelector(".vtx-math-display.is-open"))`), "the new display formula is open");
  assert.deepEqual(above.slice(0, 4), ["p:积分形式：", "display", "display", "p:after"], "the new formula stands between the line and the next formula, no empty line added");
});

// at the document's end (a browser of its own: the Mac app takes the stored document into its library once)
await session(async ({ open, caretAtEndOfFirstLine, typeDollars, blocks, evaluate, key }) => {
  await open([{ type: "paragraph", content: [{ type: "text", text: "last line" }] }]);
  await caretAtEndOfFirstLine();
  await typeDollars();
  const last = await blocks();
  console.log("at the end:", JSON.stringify(last));
  assert.deepEqual(last, ["p:last line", "display"], "at the document's end nothing follows the new formula");
  await key("ArrowRight", "ArrowRight", 39);
  await sleep(400);
  const after = await blocks();
  console.log("after Right:", JSON.stringify(after));
  assert.deepEqual(after, ["p:last line", "display", "p:"], "Right out of the last formula makes a line to write on");
  assert.ok(await evaluate(`(() => { const s = getSelection(); const p = [...document.querySelectorAll(".vtx-document > p")].pop(); return !!s.anchorNode && p.contains(s.anchorNode); })()`), "the caret is on that line");
});
console.log("display formula insertion without an empty line: passed");
