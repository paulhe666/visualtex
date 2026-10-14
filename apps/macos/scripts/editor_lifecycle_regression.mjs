import assert from "node:assert/strict";
import { withBrowserTest } from "./browser_test_runtime.mjs";

await withBrowserTest(async ({ baseUrl, send, evaluate }) => {
  const floating = await evaluate(`(async () => {
    const { installFloatingLayerAutoAvoidance } = await import('/src/runtime/floatingLayerAutoAvoidance.ts');
    const timeout = window.setTimeout;
    const wait = (ms) => new Promise((resolve) => timeout(resolve, ms));
    let timers = 0;
    window.setTimeout = (...args) => { timers++; return timeout(...args); };
    const dispose = installFloatingLayerAutoAvoidance();
    const menu = document.createElement('div');
    try {
      for (let i = 0; i < 100; i++) {
        const node = document.createElement('p');
        document.body.append(node);
        await Promise.resolve();
        node.remove();
        await Promise.resolve();
      }
      await wait(40);
      const idleTimers = timers;
      menu.setAttribute('role', 'menu');
      menu.style.cssText = 'position:fixed;width:180px;height:100px;left:' + (innerWidth - 20) + 'px;top:' + (innerHeight - 10) + 'px';
      document.body.append(menu);
      await wait(80);
      const fits = () => {
        const r = menu.getBoundingClientRect();
        return r.right <= innerWidth - 9 && r.bottom <= innerHeight - 9 && r.left >= 9;
      };
      const openedFits = fits();
      menu.style.width = '320px';
      await wait(80);
      const resizedFits = fits();
      menu.hidden = true;
      await wait(40);
      const hiddenRestored = menu.style.translate === '';
      menu.hidden = false;
      await wait(80);
      const reopenedFits = fits();
      dispose();
      return { idleTimers, openedFits, resizedFits, hiddenRestored, reopenedFits,
        disposedRestored: menu.style.translate === '' };
    } finally {
      dispose(); menu.remove(); window.setTimeout = timeout;
    }
  })()`);
  assert.deepEqual(floating, {
    idleTimers: 0, openedFits: true, resizedFits: true,
    hiddenRestored: true, reopenedFits: true, disposedRestored: true,
  });

  await send("Page.navigate", { url: baseUrl });
  // The script runs through the same store and mounted MathEditor as the app.
  const focus = await evaluate(`(async () => {
    const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
    const { useEditorStore, createFormulaLine } = await import('/src/stores/editorStore.ts');
    const readyBy = performance.now() + 10000;
    while (!document.querySelector('math-field')?.shadowRoot) {
      if (performance.now() > readyBy) throw new Error('Editor did not mount');
      await wait(40);
    }
    useEditorStore.setState({
      lines: ['abc', 'def', 'ghi'].map((latex, i) => createFormulaLine(latex, 'race-' + i)),
      activeLineId: 'race-0',
    });
    await wait(200);
    const fields = [...document.querySelectorAll('math-field')];
    const down = () => {
      fields[0].focus(); fields[0].position = 0;
      fields[0].dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', code: 'ArrowDown', bubbles: true, composed: true }));
      return document.activeElement === fields[1];
    };
    const moved = down();
    await wait(20);
    fields[2].focus(); fields[2].position = 2;
    await wait(120);
    const anotherRowPreserved = document.activeElement === fields[2] && fields[2].position === 2;
    const movedAgain = down();
    await wait(20);
    fields[1].position = fields[1].lastOffset;
    await wait(120);
    const caretPreserved = document.activeElement === fields[1] && fields[1].position === fields[1].lastOffset;
    return { moved, anotherRowPreserved, movedAgain, caretPreserved };
  })()`);
  assert.deepEqual(focus, { moved: true, anotherRowPreserved: true, movedAgain: true, caretPreserved: true });

  const formulas = [
    String.raw`\begin{bmatrix}a&b\\c&d\end{bmatrix}`,
    String.raw`\left(\frac{x}{y}\right)`,
    String.raw`\begin{aligned}a&=b\\c&=d\end{aligned}`,
    String.raw`\vec{}+z`,
    "x",
  ];
  const preview = await evaluate(`(async () => {
    const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
    const { useEditorStore, createFormulaLine } = await import('/src/stores/editorStore.ts');
    useEditorStore.setState({
      editorLayout: 'standard', sourceOpen: true,
      lines: [createFormulaLine('x', 'external-sync')], activeLineId: 'external-sync',
    });
    await wait(200);
    const source = document.querySelector('.source-panel .cm-content');
    if (!source) throw new Error('Source editor did not mount');
    source.focus();
    await wait(80);
    const field = document.querySelector('math-field');
    if (!field.readOnly) throw new Error('Source preview did not activate');
    let temporaryFields = 0;
    const observer = new MutationObserver((records) => {
      for (const record of records) for (const node of record.addedNodes) {
        if (node instanceof Element) temporaryFields +=
          Number(node.matches('math-field')) + node.querySelectorAll('math-field').length;
      }
    });
    observer.observe(document.body, { childList: true, subtree: true });
    try {
      const values = [];
      for (const latex of ${JSON.stringify(formulas)}) {
        useEditorStore.setState({ lines: [createFormulaLine(latex, 'external-sync')] });
        await wait(100);
        if (document.querySelector('math-field') !== field) throw new Error('Formula field was remounted');
        values.push(field.value.replace(/\\s+/g, ''));
      }
      return { values, temporaryFields, sourceKeepsFocus: document.activeElement === source };
    } finally { observer.disconnect(); }
  })()`);
  assert.deepEqual(preview, {
    values: formulas.map((latex) => latex.replace(/\s+/g, '')),
    temporaryFields: 0,
    sourceKeepsFocus: true,
  });
  console.log("Editor focus, structural source sync and floating-layer lifecycle passed.");
});
