import assert from "node:assert/strict";
import { withBrowserTest } from "./browser_test_runtime.mjs";

// Editor behaviours shared with the Windows editor (ported 2026-10-14):
// Cmd+J sibling script, Tab focus, Escape on a selection, big-operator
// Backspace, lower-before-upper navigation, empty-slot boxes, slot-tall caret,
// \bm round trip, candidate ranking/numbering and LaTeX source completion.
await withBrowserTest(async ({ baseUrl, send, evaluate }) => {
  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  await send("Page.navigate", { url: baseUrl });
  await evaluate(`(async () => {
    const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
    const deadline = performance.now() + 20000;
    while (!document.querySelector('math-field')?.shadowRoot) {
      if (performance.now() > deadline) throw new Error('Editor did not mount');
      await wait(50);
    }
  })()`);
  const key = async (value, code, keyCode, modifiers = 0) => {
    const common = { key: value, code, modifiers, windowsVirtualKeyCode: keyCode, nativeVirtualKeyCode: keyCode };
    await send("Input.dispatchKeyEvent", {
      type: "keyDown", ...common,
      ...(value.length === 1 && (modifiers & ~8) === 0 ? { text: value, unmodifiedText: value } : {}),
    });
    await send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
    await sleep(60);
  };
  const type = async (text) => {
    for (const character of text) {
      const code = character === "\\" ? "Backslash" : /[0-9]/.test(character) ? `Digit${character}` : `Key${character.toUpperCase()}`;
      await key(character, code, character === "\\" ? 220 : character.toUpperCase().charCodeAt(0));
    }
  };
  const reset = (latex = "") => evaluate(`(async () => {
    const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
    const { useEditorStore: store, createFormulaLine } = await import('/src/stores/editorStore.ts');
    store.setState({ editorLayout: 'standard', lines: [createFormulaLine(${JSON.stringify(latex)}, 'ported')], activeLineId: 'ported' });
    await wait(200);
    const field = document.querySelector('math-field');
    field.focus();
    field.shadowRoot.querySelector('[part="keyboard-sink"]').focus({ preventScroll: true });
    field.position = field.lastOffset;
    await wait(80);
  })()`);
  const field = (expression) => evaluate(`(() => { const f = document.querySelector('math-field'); return (${expression}); })()`);

  await reset();
  await type("A");
  await key("j", "KeyJ", 74, 4);
  await sleep(150);
  await type("1");
  await sleep(120);
  await type("2");
  await sleep(120);
  const scripts = await field(`f.value.replace(/\\s+/g, '')`);
  assert.match(scripts, /^A(_1\^2|\^2_1|_\{1\}\^\{2\}|\^\{2\}_\{1\})$/, "Cmd+J subscript must then fill the empty superscript");

  await reset("x+1");
  await key("Tab", "Tab", 9);
  await sleep(120);
  await type("y");
  const tab = await field(`({ value: f.value, focused: f.hasFocus() })`);
  assert.ok(tab.focused && tab.value.includes("y"), `Tab with no placeholder must keep focus: ${JSON.stringify(tab)}`);

  await reset("c^2");
  await field(`(f.selection = { ranges: [[0, f.lastOffset]], direction: 'forward' }, true)`);
  await key("Escape", "Escape", 27);
  await sleep(120);
  const escape = await field(`({ value: f.value, collapsed: f.selectionIsCollapsed, mode: f.mode })`);
  assert.ok(escape.value.replace(/\s/g, "") === "c^2" && escape.collapsed && escape.mode !== "latex",
    `Escape on a selection must only collapse it: ${JSON.stringify(escape)}`);

  await reset("a");
  await evaluate(`(() => document.querySelector('math-field').insert(String.raw\`\\sum_{\\placeholder{}}^{\\placeholder{}} \\placeholder{}\`,
    { mode: 'math', format: 'latex', selectionMode: 'placeholder', focus: true }))()`);
  await sleep(120);
  assert.equal(await field(`(() => { const m = f._mathfield.model; return m.at(Math.max(f.position, m.anchor))?.parentBranch ?? null; })()`),
    "subscript", "a big-operator template selects its lower limit first");
  await key("Backspace", "Backspace", 8);
  await sleep(150);
  assert.equal((await field(`f.value`)).replace(/\s/g, ""), "a", "Backspace in a blank \\sum removes the operator only");

  await reset(String.raw`x_{a}^{b}`);
  await field(`(f.position = 1, true)`);
  await key("ArrowRight", "ArrowRight", 39);
  assert.equal(await field(`(() => { const m = f._mathfield.model; return m.at(m.position)?.parentBranch ?? null; })()`),
    "subscript", "ArrowRight from the base enters the lower script first");

  await reset(String.raw`\frac{}{b}`);
  assert.ok(await field(`f.shadowRoot.querySelectorAll('.visualtex-empty-slot').length`) >= 1, "an empty slot draws its placeholder box");
  await reset(String.raw`\frac{abc}{d}`);
  await field(`(f.position = 2, true)`);
  await sleep(150);
  assert.match(await field(`f.shadowRoot.querySelector('.ML__caret')?.style.getPropertyValue('--visualtex-caret-height') ?? ''`),
    /px$/, "a caret in a numerator spans the numerator");

  await reset(String.raw`\bm{x}+\bm{\mathbb{R}}`);
  const bold = await field(`f.value`);
  assert.ok(/\\bm\{x\}/.test(bold) && /\\mathbb\{R\}/.test(bold), `\\bm must round-trip: ${bold}`);

  await reset();
  await type("\\int");
  await sleep(250);
  const candidates = await evaluate(`(() => {
    const rows = [...(document.getElementById('mathlive-suggestion-popover')?.querySelectorAll('li[data-command]') ?? [])];
    return { first: rows[0]?.dataset.command, numbers: rows.slice(0, 3).map((row) => row.querySelector('.ML__popover__index')?.textContent) };
  })()`);
  assert.deepEqual(candidates, { first: "\\int", numbers: ["1", "2", "3"] }, "an exact command ranks first and rows are numbered");
  await key("Escape", "Escape", 27);

  const source = await evaluate(`(async () => {
    const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
    const { useEditorStore: store } = await import('/src/stores/editorStore.ts');
    store.setState({ sourceOpen: true });
    await wait(300);
    const { EditorView } = await import('/node_modules/.vite/deps/@codemirror_view.js');
    const view = EditorView.findFromDOM(document.querySelector('.source-panel .cm-content'));
    view.focus();
    view.dispatch({ changes: { from: 0, to: view.state.doc.length, insert: '' } });
    for (const character of String.fromCharCode(92) + 'fra') {
      const head = view.state.selection.main.head;
      view.dispatch(view.state.update({ changes: { from: head, insert: character }, selection: { anchor: head + 1 }, userEvent: 'input.type' }));
      await wait(30);
    }
    await wait(400);
    const tooltip = document.querySelector('.cm-tooltip-autocomplete.cm-vt-completion');
    return tooltip?.querySelector('.cm-completionLabel')?.textContent ?? null;
  })()`);
  assert.equal(source, "\\frac", "the LaTeX source panel completes \\fra with \\frac first");
});

console.log("Ported editor behaviour regression passed");
