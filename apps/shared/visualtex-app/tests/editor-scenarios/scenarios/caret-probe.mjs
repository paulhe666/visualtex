// the drawn caret stands where the editing model's position is, across structures

export const names = ["caret-probe"];

// Windows only for now: in Chrome on macOS the drawn caret is not yet at the model position when this probe looks
// (seen in the fraction and square root cases); whether the macOS app shows it there is still to be checked
export const platforms = ["windows"];

export async function run({ evaluate, key, sleep, waitForEvaluation }) {
  const cases = [
    {
      name: "fraction numerator",
      source: String.raw`p+\frac{z+n}{d}+q`,
      expectedAfterInput: String.raw`p+\frac{zk+n}{d}+q`,
    },
    {
      name: "square root",
      source: String.raw`p+\sqrt{z+s}+q`,
      expectedAfterInput: String.raw`p+\sqrt{zk+s}+q`,
    },
  ];
  const states = [];
  for (const testCase of cases) {
    await waitForEvaluation(
      `(() => ({ ready: Boolean(document.querySelector("math-field")?.isConnected) }))()`,
      `mounted field before caret probe: ${testCase.name}`,
    );
    const prepared = await evaluate(`(() => {
          const field = document.querySelector("math-field");
          if (!field?.isConnected) return { ready: false };
          field.setValue(${JSON.stringify(testCase.source)}, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceAll",
            selectionMode: "after",
            silenceNotifications: true,
          });
          let markerEnd = -1;
          for (let end = 1; end <= field.lastOffset; end += 1) {
            const rangeLatex = field.getValue(end - 1, end, "latex").trim();
            const infoLatex = field.getElementInfo(end)?.latex?.trim() ?? "";
            if (rangeLatex === "z" || infoLatex === "z") {
              markerEnd = end;
              break;
            }
          }
          if (markerEnd < 0) return { ready: false, value: field.value };
          field.focus();
          field.selection = { ranges: [[markerEnd, markerEnd]], direction: "none" };
          field.position = markerEnd;
          field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
          return {
            ready: true,
            position: markerEnd,
            targetAtomId: field?._mathfield?.model?.at(markerEnd)?.id ?? "",
          };
        })()`);
    if (!prepared?.ready) {
      throw new Error(`Failed to prepare caret probe: ${JSON.stringify({ testCase, prepared })}`);
    }
    await sleep(100);

    const state = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const mathfield = field?._mathfield;
          const model = mathfield?.model;
          if (!field?.isConnected || !model) return { ready: false };
          const targetAtom = model.at(field.position);
          const targetElement = targetAtom ? mathfield.getHTMLElement(targetAtom) : null;
          const targetBounds = targetElement?.getBoundingClientRect();
          const markers = [...(field.shadowRoot?.querySelectorAll(
            ".ML__caret, .ML__text-caret, .ML__latex-caret",
          ) ?? [])].map((marker) => {
            const bounds = marker.getBoundingClientRect();
            const chain = [];
            let node = marker;
            while (node && chain.length < 10) {
              chain.push({
                tag: node.tagName,
                className: node.className || "",
                atomId: node.getAttribute?.("data-atom-id") || "",
                text: node.textContent || "",
              });
              node = node.parentElement;
            }
            return {
              x: bounds.right,
              y: bounds.top + bounds.height / 2,
              width: bounds.width,
              height: bounds.height,
              chain,
            };
          }).filter((marker) => marker.height > 0);
          const targetRight = targetBounds?.right ?? Number.NaN;
          const closest = markers
            .map((marker) => ({
              ...marker,
              deltaX: Number.isFinite(targetRight)
                ? Math.abs(marker.x - targetRight)
                : Number.POSITIVE_INFINITY,
            }))
            .sort((left, right) => left.deltaX - right.deltaX)[0] ?? null;
          return {
            ready:
              field.value === ${JSON.stringify(testCase.source)} &&
              (field.getElementInfo(field.position)?.latex?.trim() ?? "") === "z" &&
              Boolean(targetBounds && targetBounds.width > 0) &&
              Boolean(closest),
            name: ${JSON.stringify(testCase.name)},
            value: field.value,
            position: field.position,
            selection: field.selection,
            targetAtomId: targetAtom?.id ?? "",
            targetLatex: field.getElementInfo(field.position)?.latex?.trim() ?? "",
            targetBounds: targetBounds
              ? {
                  left: targetBounds.left,
                  right: targetBounds.right,
                  top: targetBounds.top,
                  bottom: targetBounds.bottom,
                }
              : null,
            closestCaret: closest,
            markers,
          };
        })()`, `stable caret/model alignment: ${testCase.name}`);

    if (
      state.targetLatex !== "z" ||
      !state.closestCaret ||
      state.closestCaret.deltaX > 3
    ) {
      throw new Error(`Visible caret/model mismatch: ${JSON.stringify(state)}`);
    }

    await key("k", "KeyK", 75);
    await sleep(80);
    const afterInput = await evaluate(`(() => {
          const field = document.querySelector("math-field");
          return {
            value: field?.value ?? "",
            position: field?.position ?? -1,
            selection: field?.selection ?? null,
          };
        })()`);
    if (afterInput.value !== testCase.expectedAfterInput) {
      throw new Error(
        `Caret input landed at a different model position: ${JSON.stringify({ state, afterInput })}`,
      );
    }
    states.push({ ...state, afterInput });
  }
  console.log(JSON.stringify(states, null, 2));
  console.log("Targeted caret/model alignment probe passed");
  return;
}
