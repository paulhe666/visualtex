// contextual upright differentials

export const names = ["upright"];

// state the page starts with (merged into the editor's persisted state)
export const storage = () => ({ inputBehavior: { autoEscapeShortcuts: false, autoExitSuperscript: true, autoExitSubscript: true, autoExitAccent: true, autoExitWrapperCommand: true, showStructuredCommandSuggestions: false, showOtherCommandSuggestions: false } });

export async function run({ clearField, evaluate, focusField, key, sleep, typeText, waitForEvaluation }) {
  await focusField();
  await clearField();
  await typeText("driver");
  const identifierState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        return {
          ready: field?.value === "driver",
          value: field?.value ?? "",
        };
      })()`, "ordinary identifier remains italic variables");

  await clearField();
  await typeText("dr/d");
  await typeText("\\theta");
  await key(" ", "Space", 32);
  await sleep(450);
  const differentialState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const value = field?.value ?? "";
        const uprightCount =
          (value.match(/\\\\differentialD|\\\\mathrm\\{d\\}/g) ?? []).length;
        return {
          ready:
            uprightCount === 2 &&
            /\\\\theta/.test(value),
          value,
          uprightCount,
          shadowText: field?.shadowRoot?.textContent ?? "",
        };
      })()`, "slash derivative uses two upright differential operators");

  await clearField();
  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        if (!field) return false;
        field.setValue("e^{i\\\\theta}", {
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
        return true;
      })()`);
  const exponentialState = await waitForEvaluation(`(() => {
        const value = document.querySelector("math-field")?.value ?? "";
        return {
          ready:
            /\\\\mathrm\\{e\\}/.test(value) &&
            /\\\\mathrm\\{i\\}/.test(value) &&
            /\\\\theta/.test(value),
          value,
        };
      })()`, "Euler constant and imaginary unit remain upright with shortcuts disabled");

  console.log(JSON.stringify({ identifierState, differentialState, exponentialState }, null, 2));
  console.log("Targeted contextual upright differential regression passed");
  return;
}
