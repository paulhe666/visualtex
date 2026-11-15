// every command surface (toolbar, candidates, tiles) inserts what it shows

export const names = ["command-surface-audit"];

export async function run({ evaluate, focusField, key, sleep, typeText, waitForEvaluation }) {
  const cases = [
    { input: "ket", expected: "\\ket{\\placeholder{}}", placeholders: 1, mode: "math" },
    { input: "bra", expected: "\\bra{\\placeholder{}}", placeholders: 1, mode: "math" },
    { input: "braket", expected: "\\braket{\\placeholder{}}{\\placeholder{}}", placeholders: 2, mode: "math" },
    { input: "ketbra", expected: "\\ketbra{\\placeholder{}}{\\placeholder{}}", placeholders: 2, mode: "math" },
    { input: "mel", expected: "\\mel{\\placeholder{}}{\\placeholder{}}{\\placeholder{}}", placeholders: 3, mode: "math" },
    { input: "abs", expected: "\\abs{\\placeholder{}}", placeholders: 1, mode: "math" },
    { input: "norm", expected: "\\norm{\\placeholder{}}", placeholders: 1, mode: "math" },
    { input: "partial", expected: "\\partial", placeholders: 0, mode: "math" },
    { input: "sqrt", expected: "\\sqrt{\\placeholder{}}", placeholders: 1, mode: "math" },
    { input: "frac", expected: "\\frac{\\placeholder{}}{\\placeholder{}}", placeholders: 2, mode: "math" },
    { input: "binom", expected: "\\binom{\\placeholder{}}{\\placeholder{}}", placeholders: 2, mode: "math" },
    { input: "overset", expected: "\\overset{\\placeholder{}}{\\placeholder{}}", placeholders: 2, mode: "math" },
    { input: "underset", expected: "\\underset{\\placeholder{}}{\\placeholder{}}", placeholders: 2, mode: "math" },
    { input: "rangle", expected: "\\rangle", placeholders: 0, mode: "math" },
    { input: "langle", expected: "\\langle", placeholders: 0, mode: "math" },
    { input: "lfloor", expected: "\\lfloor", placeholders: 0, mode: "math" },
    { input: "rfloor", expected: "\\rfloor", placeholders: 0, mode: "math" },
    { input: "lceil", expected: "\\lceil", placeholders: 0, mode: "math" },
    { input: "rceil", expected: "\\rceil", placeholders: 0, mode: "math" },
    { input: "lVert", expected: "\\lVert", placeholders: 0, mode: "math" },
    { input: "rVert", expected: "\\rVert", placeholders: 0, mode: "math" },
    { input: "lvert", expected: "\\lvert", placeholders: 0, mode: "math" },
    { input: "rvert", expected: "\\rvert", placeholders: 0, mode: "math" },
    { input: "Delta", expected: "\\Delta", placeholders: 0, mode: "math" },
    { input: "delta", expected: "\\delta", placeholders: 0, mode: "math" },
    { input: "Box", expected: "\\Box", placeholders: 0, mode: "math" },
    { input: "boxed", expected: "\\boxed{\\placeholder{}}", placeholders: 1, mode: "math" },
    { input: "mathbb", expected: "\\mathbb{}", placeholders: 0, mode: "latex", source: "\\mathbb" },
    { input: "mathbf", expected: "\\mathbf{}", placeholders: 0, mode: "latex", source: "\\mathbf" },
    { input: "mathcal", expected: "\\mathcal{}", placeholders: 0, mode: "latex", source: "\\mathcal" },
    { input: "vec", expected: "\\vec{}", placeholders: 0, mode: "latex", source: "\\vec" },
    { input: "hat", expected: "\\hat{}", placeholders: 0, mode: "latex", source: "\\hat" },
    { input: "bar", expected: "\\bar{}", placeholders: 0, mode: "latex", source: "\\bar" },
  ];

  const results = [];
  for (const testCase of cases) {
    await focusField();
    const previousValue = await evaluate(`(() => {
          const field = document.querySelector("math-field");
          return field?.value ?? "";
        })()`);
    if (previousValue) {
      await evaluate(`(() => {
            const field = document.querySelector("math-field");
            field?.executeCommand("selectAll");
            return true;
          })()`);
      await key("Backspace", "Backspace", 8);
      await waitForEvaluation(`(() => {
            const field = document.querySelector("math-field");
            return {
              ready: Boolean(field?.isConnected && field.value === ""),
              value: field?.value ?? "",
            };
          })()`, `real empty field before \\${testCase.input}`);
      await sleep(80);
    }
    await focusField();
    await typeText("\\" + testCase.input);
    const beforeSpace = await evaluate(`(() => {
          const field = document.querySelector("math-field");
          const completion = field?._mathfield?.visualTexCompletion;
          return {
            suggestions: completion?.suggestions ?? [],
            suggestionIndex: field?._mathfield?.suggestionIndex ?? -1,
            records: (completion?.suggestions ?? []).map((latex) => {
              const record = field?._mathfield?.visualTexCompletionRecords?.get(latex);
              return {
                latex,
                id: record?.id ?? "",
                command: record?.command ?? "",
                score: record?.score ?? null,
              };
            }),
          };
        })()`);
    await key(" ", "Space", 32);
    await sleep(70);
    const state = await evaluate(`(() => {
          const field = document.querySelector("math-field");
          const root = field?.shadowRoot;
          return {
            value: field?.value ?? "",
            mode: field?.mode ?? "",
            selection: field?.selection ?? null,
            selectionIsPlaceholder: field?.selectionIsPlaceholder ?? false,
            selectedPlaceholderCount:
              root?.querySelectorAll(".visualtex-structural-placeholder.ML__selected, .ML__placeholder.ML__selected").length ?? -1,
            placeholderCaretCount:
              root?.querySelectorAll(".visualtex-structural-placeholder-caret").length ?? -1,
            placeholders:
              root?.querySelectorAll(".visualtex-structural-placeholder, .ML__placeholder").length ?? -1,
            sourceCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
            sourceKind: field?.dataset.visualtexSourceWrapperKind ?? "",
            raw: [...(root?.querySelectorAll(".ML__raw-latex") ?? [])]
              .filter((node) => !node.classList.contains("ML__suggestion"))
              .map((node) => node.textContent ?? "")
              .join(""),
          };
        })()`);
    const expectedSource = testCase.source ?? "";
    if (
      state.value !== testCase.expected ||
      state.placeholders !== testCase.placeholders ||
      state.mode !== testCase.mode ||
      state.sourceCommand !== expectedSource
    ) {
      throw new Error(
        `Command surface mismatch for \\${testCase.input}: ${JSON.stringify({ testCase, beforeSpace, state })}`,
      );
    }
    results.push({ input: testCase.input, ...state });
  }

  console.log(JSON.stringify(results, null, 2));
  console.log("Targeted command surface audit passed");
  return;
}
