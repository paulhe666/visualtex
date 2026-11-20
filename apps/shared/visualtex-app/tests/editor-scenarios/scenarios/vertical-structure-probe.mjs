// the caret moves vertically through structures

export const names = ["vertical-structure-probe"];

export async function run({ clearField, waitForEvaluation }) {
  const cases = [
    { name: "fraction", latex: String.raw`\frac{N}{D}`, anchor: "N" },
    { name: "integral", latex: String.raw`\int_{L}^{U} f`, anchor: "L" },
    { name: "overset", latex: String.raw`\overset{U}{B}`, anchor: "B" },
    { name: "underset", latex: String.raw`\underset{L}{B}`, anchor: "B" },
    { name: "stackrel", latex: String.raw`\stackrel{U}{B}`, anchor: "B" },
    { name: "stackbin", latex: String.raw`\stackbin{U}{B}`, anchor: "B" },
    {
      name: "overunderset",
      latex: String.raw`\overset{U}{\underset{L}{B}}`,
      anchor: "B",
    },
    {
      name: "fraction-numerator-superscript",
      latex: String.raw`\frac{x_{L}^{U}}{D}`,
      anchor: "U",
    },
    {
      name: "fraction-numerator-subscript",
      latex: String.raw`\frac{x_{L}^{U}}{D}`,
      anchor: "L",
    },
    {
      name: "nested-fraction",
      latex: String.raw`\frac{\frac{A}{B}}{C}`,
      anchor: "B",
    },
    { name: "scripts", latex: String.raw`x_{L}^{U}`, anchor: "U" },
    { name: "sqrt-index", latex: String.raw`\sqrt[L]{R}`, anchor: "L" },
    { name: "binomial", latex: String.raw`\binom{N}{D}`, anchor: "N" },
    { name: "matrix", latex: String.raw`\begin{matrix}A&B\\C&D\end{matrix}`, anchor: "A" },
    { name: "substack", latex: String.raw`\substack{A\\B}`, anchor: "A" },
    { name: "sum-limits", latex: String.raw`\sum_{L}^{U}`, anchor: "U" },
    { name: "xarrow", latex: String.raw`\xrightarrow[L]{U}`, anchor: "U" },
    { name: "overbrace", latex: String.raw`\overbrace{A}^{U}`, anchor: "U" },
    { name: "underbrace", latex: String.raw`\underbrace{A}_{L}`, anchor: "L" },
    { name: "matrix-in-fraction", latex: String.raw`\frac{\begin{matrix}A\\B\end{matrix}}{C}`, anchor: "B" },
  ];
  const results = [];
  for (const testCase of cases) {
    await clearField();
    const state = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          if (!field?.isConnected) return { ready: false };
          field.setValue(${JSON.stringify(testCase.latex)}, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceAll",
            selectionMode: "after",
            silenceNotifications: true,
          });
          let anchorOffset = -1;
          for (let offset = 1; offset <= field.lastOffset; offset += 1) {
            const latex = field.getElementInfo(offset)?.latex?.trim() ?? "";
            if (latex === ${JSON.stringify(testCase.anchor)}) {
              anchorOffset = offset;
              break;
            }
          }
          if (anchorOffset < 0) return { ready: false, value: field.value };
          field.focus();
          field.selection = {
            ranges: [[anchorOffset, anchorOffset]],
            direction: "none",
          };
          field.position = anchorOffset;
          field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({
            preventScroll: true,
          });
          const read = () => {
            const marker = [...(field.shadowRoot?.querySelectorAll(
              ".ML__caret, .ML__placeholder-selected, .ML__selected",
            ) ?? [])].find((candidate) => candidate.getBoundingClientRect().height > 0);
            const markerBox = marker?.getBoundingClientRect();
            const chain = [];
            let node = marker;
            while (node && chain.length < 12) {
              const box = node.getBoundingClientRect();
              chain.push({
                tag: node.tagName,
                className: node.className || "",
                atomId: node.getAttribute?.("data-atom-id") || "",
                top: box.top,
                left: box.left,
                width: box.width,
                height: box.height,
              });
              node = node.parentElement;
            }
            return {
              value: field.value,
              position: field.position,
              selection: field.selection,
              marker: markerBox
                ? {
                    top: markerBox.top,
                    left: markerBox.left,
                    width: markerBox.width,
                    height: markerBox.height,
                  }
                : null,
              chain,
            };
          };
          const offsets = [];
          for (let offset = 0; offset <= field.lastOffset; offset += 1) {
            const info = field.getElementInfo(offset);
            const bounds = info?.bounds;
            const atom = field._mathfield?.model?.at(offset);
            const ancestors = [];
            for (let child = atom; child?.parent && ancestors.length < 8; child = child.parent) {
              ancestors.push({ type: child.parent.type, branch: child.parentBranch });
            }
            offsets.push({
              offset,
              latex: info?.latex ?? "",
              depth: info?.depth ?? null,
              type: atom?.type ?? null,
              parentBranch: atom?.parentBranch ?? null,
              parentType: atom?.parent?.type ?? null,
              ancestors,
              bounds: bounds
                ? {
                    top: bounds.top,
                    left: bounds.left,
                    width: bounds.width,
                    height: bounds.height,
                  }
                : null,
            });
          }
          const initial = read();
          const nativeUpChanged = field.executeCommand("moveUp");
          const afterUp = read();
          field.selection = {
            ranges: [[anchorOffset, anchorOffset]],
            direction: "none",
          };
          field.position = anchorOffset;
          const nativeDownChanged = field.executeCommand("moveDown");
          const afterDown = read();
          return {
            ready: true,
            name: ${JSON.stringify(testCase.name)},
            anchorOffset,
            offsets,
            initial,
            nativeUpChanged,
            afterUp,
            nativeDownChanged,
            afterDown,
            contentHtml:
              field.shadowRoot?.querySelector('[part="content"]')?.innerHTML ?? "",
          };
        })()`, `vertical structure probe ${testCase.name}`);
    results.push(state);
  }
  console.log(JSON.stringify(results, null, 2));
  console.log("Vertical structure probe passed");
  return;
}
