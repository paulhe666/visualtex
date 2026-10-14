import assert from "node:assert/strict";
import { withBrowserTest } from "./browser_test_runtime.mjs";

await withBrowserTest(async ({ baseUrl, send, evaluate }) => {
  await send("Page.navigate", { url: baseUrl });
  await evaluate(`(async () => {
    const deadline = performance.now() + 10000;
    while (!document.querySelector('math-field')?.shadowRoot) {
      if (performance.now() > deadline) throw new Error('Editor did not mount');
      await new Promise(resolve => setTimeout(resolve, 40));
    }
  })()`);

  const native = await evaluate(`(() => {
    const Field = document.querySelector('math-field').constructor;
    const field = new Field();
    field.mathVirtualKeyboardPolicy = 'manual';
    field.dataset.visualtexAutoExitSuperscript = 'true';
    field.dataset.visualtexAutoExitSubscript = 'true';
    field.dataset.visualtexAutoExitAccent = 'true';
    document.body.append(field);
    const slash = String.fromCharCode(92);
    const seed = () => {
      field.setValue('x^{' + slash + 'placeholder{}}', { selectionMode: 'placeholder' });
      field.focus();
    };
    try {
      seed();
      field.executeCommand('typedText', 'a');
      const keyboard = { latex: field.value, exited: field.position === field.lastOffset };
      seed();
      field.visualTexInsertToolbarTemplate(slash + 'alpha', { selectionMode: 'after' });
      const toolbar = { latex: field.value, exited: field.position === field.lastOffset };
      // Re-entering the same script must not consume auto-exit twice.
      const alpha = Array.from({ length: field.lastOffset + 1 }, (_, i) => i)
        .find(i => field.getElementInfo(i)?.latex === slash + 'alpha');
      field.position = alpha;
      field.visualTexInsertToolbarTemplate('b', { selectionMode: 'after' });
      const reentered = { latex: field.value, stayed: field.position < field.lastOffset };
      field.dataset.visualtexAutoExitSuperscript = 'false';
      seed();
      field.visualTexInsertToolbarTemplate('a', { selectionMode: 'after' });
      const disabled = field.position < field.lastOffset;
      field.setValue(slash + 'vec{' + slash + 'placeholder{}}', { selectionMode: 'placeholder' });
      field.visualTexInsertToolbarTemplate('v', { selectionMode: 'after' });
      const accent = { latex: field.value, exited: field.position === field.lastOffset };
      field.setValue(slash + 'begin{cases}x&y' + slash + 'end{cases}');
      const x = Array.from({ length: field.lastOffset + 1 }, (_, i) => i)
        .find(i => field.getElementInfo(i)?.latex === 'x');
      field.position = x;
      const environment = field.visualTexParentEnvironment;
      return { keyboard, toolbar, reentered, disabled, accent, environment };
    } finally { field.remove(); }
  })()`);
  assert.deepEqual(native, {
    keyboard: { latex: 'x^{a}', exited: true },
    toolbar: { latex: 'x^{\\alpha}', exited: true },
    reentered: { latex: 'x^{\\alpha b}', stayed: true },
    disabled: true,
    accent: { latex: '\\vec{v}', exited: true },
    environment: 'cases',
  });

  const transaction = await evaluate(`(async () => {
    const { useEditorStore: store, createFormulaLine } = await import('/src/stores/editorStore.ts');
    const { HistoryManager } = await import('/src/history/HistoryManager.ts');
    const { commitFormulaEdit, applyHistoryEntryToEditor, getEditorDocumentSnapshot } =
      await import('/src/history/documentHistory.ts');
    store.setState({ lines: [createFormulaLine('a', 'one'), createFormulaLine('b', 'two')], activeLineId: 'one' });
    const history = new HistoryManager();
    history.configure({ getDocumentSnapshot: getEditorDocumentSnapshot, applyEntry: applyHistoryEntryToEditor });
    const seen = [];
    const unsubscribe = store.subscribe(state => seen.push([state.lines[1].latex, state.activeLineId]));
    const selection = { ranges: [[1, 1]], direction: 'none' };
    try {
      commitFormulaEdit(history, { lineId: 'two', beforeLatex: 'b', afterLatex: 'bc',
        beforeSelection: selection, afterSelection: { ranges: [[2, 2]], direction: 'none' },
        beforeActiveLineId: 'one', afterActiveLineId: 'two', editKind: 'insert', source: 'keyboard' });
      await history.undo();
      await history.redo();
      return { seen, undoEntries: history.getState().undoStack.length };
    } finally { unsubscribe(); history.clear(); }
  })()`);
  assert.deepEqual(transaction, {
    seen: [['bc', 'two'], ['b', 'one'], ['bc', 'two']], undoEntries: 1,
  });
});
console.log("Native toolbar semantics and atomic document history passed.");
