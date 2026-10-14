import assert from 'node:assert/strict';
import { withBrowserTest } from './browser_test_runtime.mjs';

await withBrowserTest(async ({ baseUrl, send, evaluate }) => {
  await send('Page.navigate', { url: baseUrl });
  const insertion = await evaluate(`(async () => {
    const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
    const { useEditorStore: store, createFormulaLine } = await import('/src/stores/editorStore.ts');
    const deadline = performance.now() + 10000;
    while (!document.querySelector('.formula-line math-field')?.shadowRoot) {
      if (performance.now() > deadline) throw new Error('Editor did not mount');
      await wait(40);
    }
    store.setState({ editorLayout: 'classic', sourceOpen: false,
      lines: [createFormulaLine('a' + String.fromCharCode(92) + 'int', 'cleanup')], activeLineId: 'cleanup' });
    await wait(200);
    const button = document.querySelector('[data-command-id="int-bare"]');
    if (!button) throw new Error('Integral toolbar command is missing');
    let field = document.querySelector('.formula-line math-field');
    field.focus(); field.selection = { ranges: [[field.lastOffset - 1, field.lastOffset]], direction: 'forward' };
    button.click(); await wait(120);
    const replacement = store.getState().lines[0].latex;

    store.setState({ lines: [createFormulaLine('a', 'cleanup')] }); await wait(120);
    field = document.querySelector('.formula-line math-field'); field.focus(); field.position = field.lastOffset;
    const prototype = Object.getPrototypeOf(field), original = prototype.insert;
    let calls = 0;
    prototype.insert = function() { calls++; return false; };
    try { button.click(); await wait(120); }
    finally { prototype.insert = original; }
    const failedInsertion = store.getState().lines[0].latex;

    store.setState({ lines: [createFormulaLine('abcdefghijklmnopqrstuvwxyz'.repeat(3), 'cleanup')] });
    await wait(150); field = document.querySelector('.formula-line math-field'); field.focus(); await wait(100);
    const measure = Element.prototype.getBoundingClientRect;
    let geometryReads = 0;
    Element.prototype.getBoundingClientRect = function() {
      if (this.getRootNode() === field.shadowRoot) geometryReads++;
      return measure.call(this);
    };
    try { field.position = 5; await wait(120); }
    finally { Element.prototype.getBoundingClientRect = measure; }
    return { replacement, failedInsertion, calls, geometryReads };
  })()`);
  assert.equal(insertion.replacement.replace(/\s+/g, ''), String.raw`a\int`);
  assert.equal(insertion.failedInsertion, 'a');
  assert.equal(insertion.calls, 1, 'A failed insertion must not delete, relocate or retry');
  assert.ok(insertion.geometryReads < 160, `Selection-only geometry work regressed: ${insertion.geometryReads}`);

  const draft = await evaluate(`(async () => {
    const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
    const { useEditorStore: store } = await import('/src/stores/editorStore.ts');
    const { EditorView } = await import('/node_modules/.vite/deps/@codemirror_view.js');
    store.setState({ editorLayout: 'standard', sourceOpen: true, theme: 'light' }); await wait(200);
    const view = EditorView.findFromDOM(document.querySelector('.source-panel .cm-content'));
    const source = String.fromCharCode(92) + '[';
    view.focus(); view.dispatch({ changes: { from: 0, to: view.state.doc.length, insert: source } });
    await wait(50); store.getState().setTheme('dark'); await wait(150);
    const next = EditorView.findFromDOM(document.querySelector('.source-panel .cm-content'));
    return { sameView: next === view, text: next.state.doc.toString() };
  })()`);
  assert.deepEqual(draft, { sameView: true, text: String.raw`\[` });

  // A hidden candidate panel must not own keys dispatched by another input.
  const unrelatedInput = await evaluate(`(() => {
    const source = document.createElement('div'); source.id = 'mathlive-suggestion-popover';
    source.className = 'is-visible'; source.style.display = 'none';
    source.innerHTML = '<ul><li data-command="a">a</li><li data-command="b">b</li></ul>';
    const input = document.createElement('input'); document.body.append(source, input); input.focus();
    try {
      const key = new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true, cancelable: true });
      input.dispatchEvent(key);
      return { prevented: key.defaultPrevented, clonedPanels: document.querySelectorAll('#visualtex-native-input-suggestion-popover').length };
    } finally { source.remove(); input.remove(); }
  })()`);
  assert.deepEqual(unrelatedInput, { prevented: false, clonedPanels: 0 });
  console.log('Editor cleanup regression passed:', { insertion, draft, unrelatedInput });
});
