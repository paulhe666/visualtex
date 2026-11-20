// Chinese input-method composition inside structures

export const names = ["structured-chinese-ime"];

export async function run({ clearField, client, evaluate, key, sleep, typeText, waitForEvaluation }) {
  const commitImeText = async (text) => {
    await sleep(120);
    await client.send("Input.imeSetComposition", {
      text,
      selectionStart: Array.from(text).length,
      selectionEnd: Array.from(text).length,
    });
    await sleep(80);
    await client.send("Input.insertText", { text });
    await client.send("Input.imeSetComposition", {
      text: "",
      selectionStart: 0,
      selectionEnd: 0,
    });
    await sleep(260);
  };
  const selectedPlaceholderState = async (description) =>
    waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const range = field?.selection?.ranges?.at(-1) ?? null;
          const selectedLatex = range
            ? field.getValue(
                Math.min(range[0], range[1]),
                Math.max(range[0], range[1]),
                "latex",
              ).trim()
            : "";
          field?.focus();
          field?.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
          return {
            ready:
              Boolean(field) &&
              field.value.includes("\\\\placeholder{}") &&
              selectedLatex === "\\\\placeholder{}",
            value: field?.value ?? "",
            selectedLatex,
            selection: field?.selection ?? null,
          };
        })()`, description);
  const readFieldState = async () =>
    evaluate(`(() => {
          const field = document.querySelector("math-field");
          const range = field?.selection?.ranges?.at(-1) ?? null;
          const textNodes = [...(field?.shadowRoot?.querySelectorAll(".ML__text") ?? [])];
          return {
            value: field?.value ?? "",
            mode: field?.mode ?? "",
            textNodeCount: textNodes.length,
            textBackgrounds: textNodes.map((node) => getComputedStyle(node).backgroundColor),
            placeholderCount:
              (field?.value.match(/\\\\placeholder\\{\\}/g) ?? []).length,
            selectedLatex: range
              ? field.getValue(
                  Math.min(range[0], range[1]),
                  Math.max(range[0], range[1]),
                  "latex",
                ).trim()
              : "",
            selection: field?.selection ?? null,
          };
        })()`);
  const clickToolbarCommand = async (commandId) => {
    const selector = `[data-command-id="${commandId}"]`;
    await waitForEvaluation(`(() => {
          const button = document.querySelector(${JSON.stringify(selector)});
          if (!(button instanceof HTMLElement)) return { ready: false };
          const rect = button.getBoundingClientRect();
          return { ready: rect.width > 0 && rect.height > 0 };
        })()`, `${commandId} toolbar command`);
    await evaluate(`document.querySelector(${JSON.stringify(selector)})?.click()`);
    return selectedPlaceholderState(`${commandId} selected placeholder`);
  };
  const pressControlShortcut = async (keyValue, code, virtualKeyCode, shift = false) => {
    const modifiers = 2 | (shift ? 8 : 0);
    await client.send("Input.dispatchKeyEvent", {
      type: "keyDown",
      key: "Control",
      code: "ControlLeft",
      windowsVirtualKeyCode: 17,
      nativeVirtualKeyCode: 17,
      modifiers: 2,
    });
    if (shift) {
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        key: "Shift",
        code: "ShiftLeft",
        windowsVirtualKeyCode: 16,
        nativeVirtualKeyCode: 16,
        modifiers,
      });
    }
    await client.send("Input.dispatchKeyEvent", {
      type: "keyDown",
      key: keyValue,
      code,
      windowsVirtualKeyCode: virtualKeyCode,
      nativeVirtualKeyCode: virtualKeyCode,
      modifiers,
    });
    await client.send("Input.dispatchKeyEvent", {
      type: "keyUp",
      key: keyValue,
      code,
      windowsVirtualKeyCode: virtualKeyCode,
      nativeVirtualKeyCode: virtualKeyCode,
      modifiers,
    });
    if (shift) {
      await client.send("Input.dispatchKeyEvent", {
        type: "keyUp",
        key: "Shift",
        code: "ShiftLeft",
        windowsVirtualKeyCode: 16,
        nativeVirtualKeyCode: 16,
        modifiers: 2,
      });
    }
    await client.send("Input.dispatchKeyEvent", {
      type: "keyUp",
      key: "Control",
      code: "ControlLeft",
      windowsVirtualKeyCode: 17,
      nativeVirtualKeyCode: 17,
      modifiers: 0,
    });
    await sleep(180);
  };

  const states = {};

  await clearField();
  await commitImeText("反对法");
  const rootChinese = await readFieldState();
  await key("x", "KeyX", 88);
  const rootAfterPlainMath = await readFieldState();
  await key("Backspace", "Backspace", 8);
  const rootAfterPlainMathCleanup = await readFieldState();
  await key("\\", "Backslash", 220);
  const rootAfterBackslash = await readFieldState();
  await typeText("alpha");
  await key(" ", "Space", 32);
  const rootAfterCommand = await readFieldState();
  states.rootChineseCommand = {
    chinese: rootChinese,
    afterPlainMath: rootAfterPlainMath,
    afterPlainMathCleanup: rootAfterPlainMathCleanup,
    afterBackslash: rootAfterBackslash,
    afterCommand: rootAfterCommand,
  };
  if (
    rootChinese.mode !== "math" ||
    rootAfterPlainMath.value !== "\\text{反对法}x" ||
    rootAfterPlainMathCleanup.value !== "\\text{反对法}" ||
    rootChinese.textBackgrounds.some(
      (background) =>
        background !== "rgba(0, 0, 0, 0)" &&
        background !== "transparent",
    ) ||
    rootAfterBackslash.mode !== "latex" ||
    !/反对法.*(?:\\alpha|α)/u.test(rootAfterCommand.value)
  ) {
    throw new Error(
      `Root Chinese IME must remain a white math editing context and accept a command immediately after text: ${JSON.stringify(states.rootChineseCommand)}`,
    );
  }

  // Reproduce the WebView state shown by the user: a Chinese run remains
  // visibly selected after composition. A physical Backslash must begin a
  // new LaTeX command after that run instead of converting the selection
  // into raw LaTeX and dropping the command prefix.
  await clearField();
  await commitImeText("反对法");
  const selectedChineseBeforeBackslash = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        if (!field) return null;
        field.selection = {
          ranges: [[0, field.lastOffset]],
          direction: "none",
        };
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        return {
          value: field.value,
          mode: field.mode,
          selection: field.selection,
        };
      })()`);
  await key("\\", "Backslash", 220);
  const selectedChineseAfterBackslash = await readFieldState();
  await typeText("gamma");
  await key(" ", "Space", 32);
  const selectedChineseAfterCommand = await readFieldState();
  states.selectedChineseCommand = {
    beforeBackslash: selectedChineseBeforeBackslash,
    afterBackslash: selectedChineseAfterBackslash,
    afterCommand: selectedChineseAfterCommand,
  };
  if (
    selectedChineseAfterBackslash.mode !== "latex" ||
    !selectedChineseAfterBackslash.selection.ranges.every(
      ([start, end]) => start === end,
    ) ||
    !/反对法.*(?:\\gamma|γ)/u.test(selectedChineseAfterCommand.value)
  ) {
    throw new Error(
      `Selected Chinese residue must collapse before Backslash and preserve the Chinese text: ${JSON.stringify(states.selectedChineseCommand)}`,
    );
  }

  // On Windows IMEs the physical Backslash key can still report the
  // ideographic-comma key value while keeping code="Backslash". The
  // physical key must have the same command-prefix semantics.
  await clearField();
  await commitImeText("中文");
  await key("、", "Backslash", 220);
  const imeBackslashAfterKey = await readFieldState();
  await typeText("delta");
  await key(" ", "Space", 32);
  const imeBackslashAfterCommand = await readFieldState();
  states.imePhysicalBackslashCommand = {
    afterBackslash: imeBackslashAfterKey,
    afterCommand: imeBackslashAfterCommand,
  };
  if (
    imeBackslashAfterKey.mode !== "latex" ||
    !/中文.*(?:\\delta|δ)/u.test(imeBackslashAfterCommand.value)
  ) {
    throw new Error(
      `Physical Backslash must start a command after Chinese IME regardless of event.key: ${JSON.stringify(states.imePhysicalBackslashCommand)}`,
    );
  }

  await clearField();
  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        if (!field) return false;
        field.setValue("\\\\text{反对法}", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.position = field.lastOffset;
        field.blur();
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        return true;
      })()`);
  await sleep(180);
  const reopenedChinese = await readFieldState();
  await key("\\", "Backslash", 220);
  const reopenedAfterBackslash = await readFieldState();
  await typeText("beta");
  await key(" ", "Space", 32);
  const reopenedAfterCommand = await readFieldState();
  states.reopenedChineseCommand = {
    reopened: reopenedChinese,
    afterBackslash: reopenedAfterBackslash,
    afterCommand: reopenedAfterCommand,
  };
  if (
    reopenedChinese.mode !== "math" ||
    reopenedChinese.textBackgrounds.some(
      (background) =>
        background !== "rgba(0, 0, 0, 0)" &&
        background !== "transparent",
    ) ||
    reopenedAfterBackslash.mode !== "latex" ||
    !/反对法.*(?:\\beta|β)/u.test(reopenedAfterCommand.value)
  ) {
    throw new Error(
      `Reopened Chinese text must recover a white math editing context and accept a command: ${JSON.stringify(states.reopenedChineseCommand)}`,
    );
  }

  // Reproduce a field whose live editing context is stale "text" even
  // though the content is a completed Chinese text run. The next ordinary
  // character must be inserted in math mode, outside \text{...}.
  await clearField();
  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        if (!field) return false;
        field.setValue("\\\\text{反对法}", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.position = field.lastOffset;
        field.mode = "text";
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        return true;
      })()`);
  await sleep(80);
  const staleTextBeforePlainInput = await readFieldState();
  await key("x", "KeyX", 88);
  const staleTextAfterPlainInput = await readFieldState();
  states.staleTextBoundaryPlainInput = {
    before: staleTextBeforePlainInput,
    after: staleTextAfterPlainInput,
  };
  if (
    staleTextAfterPlainInput.value !== "\\text{反对法}x" ||
    staleTextAfterPlainInput.mode !== "math"
  ) {
    throw new Error(
      `A stale text mode at the right edge of Chinese text must normalize before ordinary input: ${JSON.stringify(states.staleTextBoundaryPlainInput)}`,
    );
  }

  // Same-value synchronization must still repair the live mode. MathLive's
  // upstream fast path used to return before applying mode:"math".
  await clearField();
  const sameValueModeRepair = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        if (!field) return null;
        field.setValue("\\\\text{反对法}", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.position = field.lastOffset;
        field.mode = "text";
        const before = { value: field.value, mode: field.mode, position: field.position };
        field.setValue(field.value, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        return {
          before,
          after: { value: field.value, mode: field.mode, position: field.position },
        };
      })()`);
  states.sameValueModeRepair = sameValueModeRepair;
  if (
    sameValueModeRepair?.before?.mode !== "text" ||
    sameValueModeRepair?.after?.mode !== "math" ||
    sameValueModeRepair?.after?.value !== "\\text{反对法}"
  ) {
    throw new Error(
      `Same-value mode synchronization failed to repair a stale text context: ${JSON.stringify(sameValueModeRepair)}`,
    );
  }

  await clearField();
  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        if (!field) return false;
        field.setValue("\\\\text{abc}", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.position = field.lastOffset;
        field.focus();
        field.mode = "text";
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        return true;
      })()`);
  await sleep(120);
  const regularText = await readFieldState();
  await key("\\", "Backslash", 220);
  const regularTextAfterBackslash = await readFieldState();
  await typeText("gamma");
  await key(" ", "Space", 32);
  const regularTextAfterCommand = await readFieldState();
  states.regularTextCommand = {
    text: regularText,
    afterBackslash: regularTextAfterBackslash,
    afterCommand: regularTextAfterCommand,
  };
  if (
    regularText.mode !== "text" ||
    regularText.textBackgrounds.some(
      (background) =>
        background !== "rgba(0, 0, 0, 0)" &&
        background !== "transparent",
    ) ||
    regularTextAfterBackslash.mode !== "latex" ||
    !/abc.*(?:\\gamma|γ)/u.test(regularTextAfterCommand.value)
  ) {
    throw new Error(
      `Text-mode content must stay visually white and preserve the LaTeX command entry key: ${JSON.stringify(states.regularTextCommand)}`,
    );
  }

  await clearField();
  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        if (!field) return false;
        field.setValue("\\\\text{反对法}", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.position = field.lastOffset;
        field._mathfield.model.mode = "text";
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        return true;
      })()`);
  const staleTextBoundaryBefore = await readFieldState();
  await key("x", "KeyX", 88);
  const staleTextBoundaryAfterX = await readFieldState();
  await key("\\", "Backslash", 220);
  const staleTextBoundaryAfterBackslash = await readFieldState();
  await typeText("alpha");
  await key(" ", "Space", 32);
  const staleTextBoundaryAfterCommand = await readFieldState();
  states.staleTextBoundary = {
    before: staleTextBoundaryBefore,
    afterX: staleTextBoundaryAfterX,
    afterBackslash: staleTextBoundaryAfterBackslash,
    afterCommand: staleTextBoundaryAfterCommand,
  };
  if (
    staleTextBoundaryAfterX.mode !== "math" ||
    staleTextBoundaryAfterX.value !== "\\text{反对法}x" ||
    staleTextBoundaryAfterBackslash.mode !== "latex" ||
    staleTextBoundaryAfterCommand.value !== "\\text{反对法}x\\alpha"
  ) {
    throw new Error(
      `Text-run right boundary must exit stale text mode before the next input: ${JSON.stringify(states.staleTextBoundary)}`,
    );
  }

  const toolbarCases = [
    {
      id: "power",
      name: "superscript",
      expected: (value) => value.includes("^{\\text{中文}}"),
      expectedPlaceholderCount: 0,
    },
    {
      id: "subscript",
      name: "subscript",
      expected: (value) => value.includes("_{\\text{中文}}"),
      expectedPlaceholderCount: 0,
    },
    {
      id: "sqrt",
      name: "square root",
      expected: (value) => value.includes("\\sqrt{\\text{中文}}"),
      expectedPlaceholderCount: 0,
    },
    {
      id: "int",
      name: "integral",
      expected: (value) =>
        value.includes("\\int") && value.includes("\\text{中文}"),
      expectedPlaceholderCount: 3,
    },
    {
      id: "matrix2",
      name: "matrix",
      expected: (value) =>
        value.includes("\\begin{bmatrix}") &&
        value.includes("\\text{中文}") &&
        value.includes("\\end{bmatrix}"),
      expectedPlaceholderCount: 3,
    },
  ];
  for (const testCase of toolbarCases) {
    await clearField();
    const before = await clickToolbarCommand(testCase.id);
    await commitImeText("中文");
    const after = await readFieldState();
    if (
      !testCase.expected(after.value) ||
      after.placeholderCount !== testCase.expectedPlaceholderCount
    ) {
      throw new Error(
        `Toolbar ${testCase.name} lost its structure during Chinese IME: ${JSON.stringify({ before, after })}`,
      );
    }
    states[testCase.name] = { before, after };
  }

  await clearField();
  const fractionBefore = await clickToolbarCommand("frac");
  await commitImeText("分子");
  const fractionNumerator = await readFieldState();
  if (
    !fractionNumerator.value.includes(
      "\\frac{\\text{分子}}{\\placeholder{}}",
    ) ||
    fractionNumerator.placeholderCount !== 1
  ) {
    throw new Error(
      `Fraction numerator composition changed another slot: ${JSON.stringify({ fractionBefore, fractionNumerator })}`,
    );
  }
  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        if (!field) return false;
        for (let offset = 1; offset <= field.lastOffset; offset += 1) {
          if (
            field.getValue(offset - 1, offset, "latex").trim() ===
            "\\\\placeholder{}"
          ) {
            field.selection = {
              ranges: [[offset - 1, offset]],
              direction: "none",
            };
            field.focus();
            field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
            return true;
          }
        }
        return false;
      })()`);
  const denominatorSelected = await selectedPlaceholderState(
    "fraction denominator selected placeholder",
  );
  await commitImeText("分母");
  const fractionComplete = await readFieldState();
  if (
    !fractionComplete.value.includes(
      "\\frac{\\text{分子}}{\\text{分母}}",
    ) ||
    fractionComplete.placeholderCount !== 0
  ) {
    throw new Error(
      `Fraction denominator composition lost the fraction: ${JSON.stringify({ denominatorSelected, fractionComplete })}`,
    );
  }
  states.fraction = {
    before: fractionBefore,
    numerator: fractionNumerator,
    denominatorSelected,
    complete: fractionComplete,
  };

  await clearField();
  const undoBefore = await clickToolbarCommand("power");
  await commitImeText("中文");
  const undoAfterComposition = await readFieldState();
  await pressControlShortcut("z", "KeyZ", 90);
  const undoState = await readFieldState();
  if (
    !undoState.value.includes("\\placeholder{}") ||
    !undoState.value.includes("^")
  ) {
    throw new Error(
      `Undo did not restore the structured placeholder transaction: ${JSON.stringify({ undoBefore, undoAfterComposition, undoState })}`,
    );
  }
  await pressControlShortcut("y", "KeyY", 89);
  const redoState = await readFieldState();
  if (!redoState.value.includes("^{\\text{中文}}")) {
    throw new Error(
      `Redo did not restore structured Chinese composition: ${JSON.stringify({ undoState, redoState })}`,
    );
  }
  states.history = {
    before: undoBefore,
    composed: undoAfterComposition,
    undo: undoState,
    redo: redoState,
  };

  console.log(JSON.stringify(states, null, 2));
  console.log("Targeted structured Chinese IME regression passed");
  return;
}
