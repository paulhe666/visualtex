// formula row geometry probe

export const names = ["geometry"];

export async function run({ evaluate, sleep, waitForEvaluation }) {
  await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        if (!field?.isConnected) return { ready: false };
        field.setValue("p+(z+r)+q+\\\\placeholder{}", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.dispatchEvent(new InputEvent("input", {
          bubbles: true,
          composed: true,
          inputType: "insertText",
        }));
        return { ready: field.isConnected };
      })()`, "stable formula field for geometry");
  await sleep(150);
  const geometry = await evaluate(`(() => {
        const rect = (element) => {
          const bounds = element?.getBoundingClientRect();
          return bounds
            ? { x: bounds.x, y: bounds.y, width: bounds.width, height: bounds.height }
            : null;
        };
        const field = document.querySelector("math-field");
        return {
          viewport: { width: innerWidth, height: innerHeight },
          title: rect(document.querySelector(".document-title-area input")),
          editorScroll: rect(document.querySelector(".editor-pane-scroll")),
          editorSurface: rect(document.querySelector(".editor-surface")),
          formulaLine: rect(document.querySelector(".formula-line")),
          mathfieldHost: rect(document.querySelector(".mathfield-host")),
          mathfield: rect(field),
          value: field?.value ?? "",
        };
      })()`);
  console.log(JSON.stringify(geometry, null, 2));
  console.log("Targeted geometry probe passed");
  return;
}
