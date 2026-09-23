// Attach to a real VisualTeX WebView2 started with a loopback debugging port.
// Usage: node scripts/installed_editor_input_regression.mjs PORT OUTPUT_DIRECTORY
// Uses the open document as a test document; preserve it before running.
// No MathField value/selection setters, React/store hooks, or synthetic DOM
// input events are used: edits go through Chromium's trusted input dispatch.
import assert from "node:assert/strict";
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";

const port = Number(process.argv[2]);
const output = process.argv[3];
assert.ok(Number.isInteger(port) && port > 0 && port < 65536 && output,
  "Usage: node scripts/installed_editor_input_regression.mjs PORT OUTPUT_DIRECTORY");
const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json();
const target = targets.find((item) => item.type === "page" && item.url === "http://tauri.localhost/");
assert.ok(target, "Expected the installed Tauri application, not a browser fixture");
await mkdir(output, { recursive: true });
const ws = new WebSocket(target.webSocketDebuggerUrl);
await new Promise((resolve, reject) => { ws.onopen = resolve; ws.onerror = reject; });
let sequence = 0;
const pending = new Map(), exceptions = [], results = [];
ws.onmessage = ({ data }) => {
  const message = JSON.parse(data);
  if (!message.id) {
    if (message.method === "Runtime.exceptionThrown") exceptions.push(message.params);
    return;
  }
  const operation = pending.get(message.id);
  pending.delete(message.id);
  message.error ? operation.reject(message.error) : operation.resolve(message.result);
};
const send = (method, params = {}) => new Promise((resolve, reject) => {
  const id = ++sequence;
  pending.set(id, { resolve, reject });
  ws.send(JSON.stringify({ id, method, params }));
});
const evaluate = async (expression) => {
  const result = await send("Runtime.evaluate", { expression, returnByValue: true, awaitPromise: true });
  if (result.exceptionDetails) throw Error(JSON.stringify(result.exceptionDetails));
  return result.result.value;
};
const settle = () => new Promise((resolve) => setTimeout(resolve, 180));
const state = () => evaluate(`({
  active: document.activeElement?.tagName,
  source: document.querySelector('.cm-content')?.innerText,
  sourceError: document.querySelector('.source-error-chip')?.textContent,
  fields: [...document.querySelectorAll('math-field')].map(f => ({
    value:f.value, readOnly:f.readOnly, focused:f.hasFocus(), selection:f.selection
  }))
})`);
const click = async (selector) => {
  const r = await evaluate(`document.querySelector(${JSON.stringify(selector)})?.getBoundingClientRect().toJSON()`);
  assert.ok(r?.width && r?.height, `No visible target: ${selector}`);
  const x = r.x + Math.min(r.width / 2, 80), y = r.y + r.height / 2;
  await send("Input.dispatchMouseEvent", { type:"mousePressed", x, y, button:"left", clickCount:1 });
  await send("Input.dispatchMouseEvent", { type:"mouseReleased", x, y, button:"left", clickCount:1 });
  await settle();
};
const key = async (key, code, virtualKey, modifiers = 0, text) => {
  const params = { key, code, windowsVirtualKeyCode: virtualKey, modifiers };
  await send("Input.dispatchKeyEvent", { type:"keyDown", ...params, ...(text ? { text } : {}) });
  await send("Input.dispatchKeyEvent", { type:"keyUp", ...params });
};
const type = async (text) => {
  for (const char of text) await key(char, `Key${char.toUpperCase()}`, char.toUpperCase().charCodeAt(0), 0, char);
  await settle();
};
const source = async (latex) => {
  await click('[data-classic-bottom-view="source"]');
  await click(".cm-content");
  await key("a", "KeyA", 65, 2);
  await send("Input.insertText", { text: `$$\n${latex}\n$$` });
  await settle();
};
const check = async (name, verify) => {
  await settle();
  const current = await state();
  results.push({ name, ...current });
  verify(current);
  assert.deepEqual(exceptions, [], `${name}: runtime exception`);
  console.log(`PASS ${name}`);
};
const editable = (s, latex) => {
  assert.equal(s.fields.length, 1);
  assert.equal(s.fields[0].readOnly, false);
  assert.equal(s.fields[0].focused, true);
  assert.equal(s.fields[0].value, latex);
};

try {
  await send("Runtime.enable");
  await writeFile(path.join(output, "before.json"), JSON.stringify(await state(), null, 2));
  for (const [latex, expected] of [
    [String.raw`a\pm b`, String.raw`a\pm bx`],
    [String.raw`\left\lbrace a\right\rbrace`, String.raw`\left\lbrace a\right\rbrace x`],
  ]) {
    await source(latex);
    await click(".mathfield-host");
    await key("End", "End", 35);
    await type("x");
    await check(`native command source round trip: ${latex}`, (s) => editable(s, expected));
  }
  await source("a");
  await click(".cm-line:nth-child(2)");
  await key("End", "End", 35);
  await key("Backspace", "Backspace", 8);
  await click(".mathfield-host");
  await type("ax");
  await check("delete final source atom, click empty visual field, type ax", (s) => editable(s, "ax"));
  await key("Backspace", "Backspace", 8);
  await check("Backspace", (s) => editable(s, "a"));
  await key("Home", "Home", 36);
  await key("Delete", "Delete", 46);
  await type("x");
  await check("Delete final visual atom then type", (s) => editable(s, "x"));

  for (const draft of ["abc", "\\frac{"]) {
    await source(draft);
    await click('button[aria-label="新建"]');
    await check(`new document replaces focused source: ${draft}`, (s) => {
      editable(s, "");
      assert.equal(s.source.replace(/\s/g, ""), "$$$$");
      assert.ok(!s.sourceError);
    });
    await type("a");
    await check("fresh document receives input", (s) => editable(s, "a"));
  }

  await source("a");
  await click(".mathfield-host");
  await type("b");
  await source("\\frac{");
  await click('button[aria-label="撤销"]');
  await check("global Undo supersedes incomplete source draft", (s) => {
    editable(s, "a");
    assert.equal(s.source.replace(/\s/g, ""), "$$a$$");
    assert.ok(!s.sourceError);
  });
  await type("x");
  await check("typing immediately after replay survives", (s) => editable(s, "ax"));

  await click('button[aria-label="新建"]');
  await type("a");
  await key("Enter", "Enter", 13, 8);
  await type("b");
  await check("Shift+Enter keeps two rows inside one field", (s) => {
    editable(s, "\\begin{gathered}a\\\\ b\\end{gathered}");
  });
  await key("z", "KeyZ", 90, 2);
  await type("x");
  await check("Undo then immediate multiline input", (s) => {
    editable(s, "\\begin{gathered}a\\\\ x\\end{gathered}");
  });

  for (const latex of ["\\vec{}\\vec{}", "\\hat{}", "\\frac{}{}", "\\sqrt{}", "\\left(\\right)"]) {
    await source(latex);
    await click(".mathfield-host");
    await check(`source replacement preserves ${latex}`, (s) => editable(s, latex));
    await send("Page.reload", { ignoreCache:true });
    const started = Date.now();
    while (true) {
      await settle();
      try {
        if (await evaluate("Boolean(document.querySelector('math-field')?.shadowRoot)")) break;
      } catch (error) {
        if (!/context|Cannot find/i.test(JSON.stringify(error))) throw error;
      }
      assert.ok(Date.now() - started < 8000, "Reloaded MathField did not mount");
    }
    await settle();
    await click(".mathfield-host");
    await check(`persisted model reload preserves ${latex}`, (s) => editable(s, latex));
  }
  await type("a");
  await check("restored structure accepts a", (s) => editable(s, "\\left(\\right)a"));
} finally {
  await writeFile(path.join(output, "result.json"), JSON.stringify({ target, results, exceptions }, null, 2));
  const screenshot = await send("Page.captureScreenshot", { format:"png" });
  await writeFile(path.join(output, "final.png"), Buffer.from(screenshot.data, "base64"));
  ws.close();
}
