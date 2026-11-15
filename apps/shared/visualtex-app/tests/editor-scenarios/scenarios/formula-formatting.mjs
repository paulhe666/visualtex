// bold, italic and colour on selected formula content

export const names = ["formula-formatting"];

export async function run({ clearField, clickSelectorWithPointer, evaluate, key, sleep, typeText, waitForEvaluation }) {
  await waitForEvaluation(`(() => ({
        ready: [
          '[data-formula-selection-bold]',
          '[data-formula-selection-italic]',
          '[data-formula-selection-color]',
          '[data-formula-selection-background]',
        ].every((selector) => {
          const button = document.querySelector(selector);
          if (!(button instanceof HTMLElement)) return false;
          const rect = button.getBoundingClientRect();
          return rect.width > 0 && rect.height > 0;
        }),
      }))()`, "desktop formula formatting controls");

  const dispatchFormattingToggle = async (selector) => {
    await evaluate(`(() => {
          const field = document.querySelector('math-field');
          if (!field) return false;
          field.focus();
          field.selection = {
            ranges: [[0, field.lastOffset]],
            direction: 'forward',
          };
          return true;
        })()`);
    await clickSelectorWithPointer(selector);
    await sleep(100);
  };

  const setFormattingValue = async (latex) => {
    const encodedLatex = JSON.stringify(latex);
    await evaluate(`(() => {
          const field = document.querySelector('math-field');
          if (!field) return false;
          field.setValue(${encodedLatex}, {
            mode: 'math',
            format: 'latex',
            insertionMode: 'replaceAll',
            selectionMode: 'after',
            silenceNotifications: true,
          });
          field.focus();
          return true;
        })()`);
    await sleep(60);
  };

  const readFormattingValue = () => evaluate(`(() => {
        const field = document.querySelector('math-field');
        return field?.value?.replace(/\\s+/g, '') ?? '';
      })()`);

  await setFormattingValue('abc');
  await dispatchFormattingToggle('[data-formula-selection-bold]');
  const boldApplied = await readFormattingValue();
  // bold keeps the letters' own (default italic) slant: symbol-wide \bm (editing core, 2026-11-01)
  if (boldApplied !== String.raw`\bm{abc}`) {
    throw new Error(`Bold toggle must keep default math italic as \\bm; received ${boldApplied}`);
  }
  await dispatchFormattingToggle('[data-formula-selection-bold]');
  const boldRemoved = await readFormattingValue();
  if (boldRemoved !== 'abc') {
    throw new Error(`Second bold toggle must restore ordinary math; received ${boldRemoved}`);
  }

  await setFormattingValue('xyz');
  await dispatchFormattingToggle('[data-formula-selection-italic]');
  const uprightApplied = await readFormattingValue();
  if (uprightApplied !== String.raw`\mathrm{xyz}`) {
    throw new Error(`Italic toggle must switch default math italic to \\mathrm; received ${uprightApplied}`);
  }
  await dispatchFormattingToggle('[data-formula-selection-italic]');
  const italicRestored = await readFormattingValue();
  if (italicRestored !== 'xyz') {
    throw new Error(`Second italic toggle must restore default math italic; received ${italicRestored}`);
  }

  await setFormattingValue(String.raw`\mathbf{q}`);
  await dispatchFormattingToggle('[data-formula-selection-italic]');
  const boldItalicApplied = await readFormattingValue();
  // bold italic on a letter whose own slant is italic is the symbol-wide \bm
  if (boldItalicApplied !== String.raw`\bm{q}`) {
    throw new Error(`Italic toggle must keep bold, written \\bm; received ${boldItalicApplied}`);
  }
  await dispatchFormattingToggle('[data-formula-selection-italic]');
  const boldUprightRestored = await readFormattingValue();
  if (boldUprightRestored !== String.raw`\mathbf{q}`) {
    throw new Error(`Second italic toggle must restore \\mathbf; received ${boldUprightRestored}`);
  }

  const selectionBold = {
    applied: boldApplied,
    removed: boldRemoved,
  };
  const selectionItalic = {
    upright: uprightApplied,
    restored: italicRestored,
    boldItalic: boldItalicApplied,
    boldRestored: boldUprightRestored,
  };

  await setFormattingValue(String.raw`\Delta`);
  await dispatchFormattingToggle('[data-formula-selection-bold]');
  const uppercaseGreekBold = await readFormattingValue();
  // \bm keeps the symbol's own upright shape
  if (uppercaseGreekBold !== String.raw`\bm{\Delta}`) {
    throw new Error(`Uppercase Greek bold must stay upright: ${uppercaseGreekBold}`);
  }
  await dispatchFormattingToggle('[data-formula-selection-bold]');
  const uppercaseGreekBoldRestored = await readFormattingValue();
  if (uppercaseGreekBoldRestored !== String.raw`\Delta`) {
    throw new Error(`Uppercase Greek bold toggle did not restore default upright form: ${uppercaseGreekBoldRestored}`);
  }
  await dispatchFormattingToggle('[data-formula-selection-italic]');
  const uppercaseGreekItalic = await readFormattingValue();
  if (uppercaseGreekItalic !== String.raw`\mathit{\Delta}`) {
    throw new Error(`Uppercase Greek italic toggle must visibly italicize: ${uppercaseGreekItalic}`);
  }
  await dispatchFormattingToggle('[data-formula-selection-italic]');
  const uppercaseGreekItalicRestored = await readFormattingValue();
  if (uppercaseGreekItalicRestored !== String.raw`\Delta`) {
    throw new Error(`Uppercase Greek italic toggle did not restore default upright form: ${uppercaseGreekItalicRestored}`);
  }
  const uppercaseGreekFormatting = {
    bold: uppercaseGreekBold,
    boldRestored: uppercaseGreekBoldRestored,
    italic: uppercaseGreekItalic,
    italicRestored: uppercaseGreekItalicRestored,
  };

  // F01: coloring a keyboard-selected superscript must color the script
  // content itself (not the glyph-less ^ attachment), and Escape must
  // never turn the selection into a raw LaTeX group that drops it.
  await setFormattingValue('E=mc^2');
  await key('ArrowLeft', 'ArrowLeft', 37);
  await key('End', 'End', 35);
  // Shift+ArrowLeft from the end selects the superscript (every press
  // changes the selection: the next one would take the base too).
  await key('ArrowLeft', 'ArrowLeft', 37, 8);
  const scriptSelected = await evaluate(`(() => {
        const field = document.querySelector('math-field');
        const [start, end] = field.selection.ranges[0];
        return field.getValue(Math.min(start, end), Math.max(start, end), 'latex');
      })()`);
  if (scriptSelected !== '^2') {
    throw new Error(`Superscript selection setup changed: ${scriptSelected}`);
  }
  await clickSelectorWithPointer('[data-formula-selection-color]');
  await clickSelectorWithPointer('[data-formula-color="#111827"]');
  await sleep(100);
  const scriptColored = await readFormattingValue();
  if (scriptColored !== String.raw`E=mc^{\textcolor{#111827}{2}}`) {
    throw new Error(`Superscript color must wrap the script content: ${scriptColored}`);
  }
  await evaluate(`(() => {
        const field = document.querySelector('math-field');
        field.focus();
        field.selection = { ranges: [[4, field.lastOffset]], direction: 'backward' };
        return true;
      })()`);
  await key('Escape', 'Escape', 27);
  await sleep(100);
  const scriptAfterEscape = await readFormattingValue();
  if (scriptAfterEscape !== scriptColored) {
    throw new Error(`Escape on a selection must not change the formula: ${scriptAfterEscape}`);
  }
  await typeText('x');
  const scriptAfterTyping = await readFormattingValue();
  if (scriptAfterTyping !== String.raw`E=mc^{\textcolor{#111827}{2}}x`) {
    throw new Error(`Typing after Escape must stay outside the colored script: ${scriptAfterTyping}`);
  }
  const scriptColor = { colored: scriptColored, afterEscape: scriptAfterEscape, afterTyping: scriptAfterTyping };

  await setFormattingValue('abc');
  await evaluate(`(() => {
        const field = document.querySelector('math-field');
        if (!field) return false;
        field.focus();
        field.selection = { ranges: [[0, field.lastOffset]], direction: 'forward' };
        return true;
      })()`);
  await clickSelectorWithPointer('[data-formula-selection-color]');
  const compactColorPopover = await waitForEvaluation(`(() => {
        const popover = document.querySelector('[data-formula-color-popover="color"]');
        const presets = popover?.querySelector('.formula-color-presets');
        const custom = popover?.querySelector('.formula-custom-colors-panel');
        if (!(popover instanceof HTMLElement) || !(presets instanceof HTMLElement) || !(custom instanceof HTMLElement)) {
          return { ready: false };
        }
        const rect = popover.getBoundingClientRect();
        const presetRect = presets.getBoundingClientRect();
        const customRect = custom.getBoundingClientRect();
        return {
          ready: rect.width <= 200 && customRect.top >= presetRect.bottom - 1,
          width: rect.width,
          presetBottom: presetRect.bottom,
          customTop: customRect.top,
        };
      })()`, 'compact formula color popover');
  if (compactColorPopover.width > 200) {
    throw new Error(`Formula color popover is still oversized: ${JSON.stringify(compactColorPopover)}`);
  }
  await clickSelectorWithPointer('[data-formula-color="#111827"]');
  await sleep(100);
  const textColorApplied = await readFormattingValue();
  if (!textColorApplied.includes('111827') || !textColorApplied.includes('abc')) {
    throw new Error(`Formula text color was not applied to the selected content: ${textColorApplied}`);
  }

  await setFormattingValue('xyz');
  await evaluate(`(() => {
        const field = document.querySelector('math-field');
        if (!field) return false;
        field.focus();
        field.selection = { ranges: [[0, field.lastOffset]], direction: 'forward' };
        return true;
      })()`);
  await clickSelectorWithPointer('[data-formula-selection-background]');
  await clickSelectorWithPointer('[data-formula-color="#fef3c7"]');
  await sleep(100);
  const backgroundColorApplied = await readFormattingValue();
  if (!backgroundColorApplied.toLowerCase().includes('fef3c7') || !backgroundColorApplied.includes('xyz')) {
    throw new Error(`Formula background color was not applied to the selected content: ${backgroundColorApplied}`);
  }

  await evaluate(`(() => {
        const field = document.querySelector('math-field');
        if (!field) return false;
        field.focus();
        field.selection = { ranges: [[0, field.lastOffset]], direction: 'forward' };
        return true;
      })()`);
  await clickSelectorWithPointer('[data-formula-selection-background]');
  const noBackgroundAction = await evaluate(`(() => {
        const button = document.querySelector('[data-formula-background-none]');
        return {
          exists: Boolean(button),
          text: button?.textContent?.trim() ?? '',
          label: button?.getAttribute('aria-label') ?? '',
        };
      })()`);
  if (!noBackgroundAction.exists || (!noBackgroundAction.text.includes('None') && !noBackgroundAction.text.includes('无背景'))) {
    throw new Error(`No-background action is missing or unclear: ${JSON.stringify(noBackgroundAction)}`);
  }
  await clickSelectorWithPointer('[data-formula-background-none]');
  await sleep(100);
  const backgroundColorCleared = await readFormattingValue();
  if (
    backgroundColorCleared.toLowerCase().includes('fef3c7') ||
    backgroundColorCleared.toLowerCase().includes('background-color') ||
    !backgroundColorCleared.includes('xyz')
  ) {
    throw new Error(`Formula background color was not cleared cleanly: ${backgroundColorCleared}`);
  }

  const selectionColors = {
    text: textColorApplied,
    background: backgroundColorApplied,
    clearedBackground: backgroundColorCleared,
    popover: compactColorPopover,
  };

  const removedPersistentControls = await evaluate(`(() => ({
        typingBold: Boolean(document.querySelector('[data-formula-typing-bold]')),
        typingItalic: Boolean(document.querySelector('[data-formula-typing-italic]')),
      }))()`);
  if (removedPersistentControls.typingBold || removedPersistentControls.typingItalic) {
    throw new Error(`Persistent typing controls must stay removed: ${JSON.stringify(removedPersistentControls)}`);
  }

  await clearField();
  await typeText("abc");
  await key("Enter", "Enter", 13);
  await waitForEvaluation(`(() => ({
        ready: document.querySelectorAll("math-field").length === 2,
      }))()`, "two formula rows for stale selection isolation");
  await typeText("xyz");
  await evaluate(`(() => {
        const fields = [...document.querySelectorAll("math-field")];
        const first = fields[0];
        const second = fields[1];
        if (!first || !second) return false;
        first.selection = { ranges: [[0, first.lastOffset]], direction: "forward" };
        second.focus();
        second.selection = { ranges: [[0, second.lastOffset]], direction: "forward" };
        second.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        return true;
      })()`);
  await sleep(80);
  await clickSelectorWithPointer('[data-formula-selection-bold]');
  const staleSelectionIsolation = await waitForEvaluation(`(() => {
        const fields = [...document.querySelectorAll("math-field")];
        const values = fields.map((field) => field.value.replace(/\s+/g, ""));
        return {
          ready:
            values.length === 2 &&
            values[0] === "abc" &&
            values[1] === "\\\\bm{xyz}",
          values,
          activeIndex: fields.findIndex((field) => field.hasFocus()),
        };
      })()`, "formatting ignores stale selection in an inactive formula row");
  await clickSelectorWithPointer('button[aria-label="撤销"]');
  const staleSelectionUndo = await waitForEvaluation(`(() => {
        const values = [...document.querySelectorAll("math-field")].map((field) =>
          field.value.replace(/\s+/g, ""),
        );
        return { ready: values.length === 2 && values[0] === "abc" && values[1] === "xyz", values };
      })()`, "undo restores only the actively formatted formula row");
  await clickSelectorWithPointer('button[aria-label="重做"]');
  const staleSelectionRedo = await waitForEvaluation(`(() => {
        const values = [...document.querySelectorAll("math-field")].map((field) =>
          field.value.replace(/\s+/g, ""),
        );
        return { ready: values.length === 2 && values[0] === "abc" && values[1] === "\\\\bm{xyz}", values };
      })()`, "redo reapplies formatting only to the active formula row");

  console.log(JSON.stringify({
    selectionBold,
    selectionItalic,
    uppercaseGreekFormatting,
    scriptColor,
    selectionColors,
    removedPersistentControls,
    staleSelectionIsolation,
    staleSelectionUndo,
    staleSelectionRedo,
  }, null, 2));
  console.log("Targeted desktop formula formatting regression passed");
  return;
}
