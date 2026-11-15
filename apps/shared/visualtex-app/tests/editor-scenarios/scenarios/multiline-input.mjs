// line-mode shortcuts and multiline input inside a formula

export const names = ["multiline-input"];

export async function run({ clearField, clickSelectorWithPointer, evaluate, key, typeText, waitForEvaluation }) {
  await clearField();
  await typeText("ab");

  await key("Enter", "Enter", 13, 2);
  const forcedInline = await waitForEvaluation(`(() => {
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        const lines = persisted?.state?.lines ?? [];
        const active = document.querySelector(".formula-line.is-active");
        const fields = [...document.querySelectorAll("math-field")];
        return {
          ready:
            lines.length === 2 &&
            lines[1]?.mode === "inline" &&
            active?.querySelector("[data-formula-line-mode-toggle]")?.getAttribute("data-formula-line-mode") === "inline" &&
            fields[1]?.hasFocus?.() === true,
          modes: lines.map((line) => line?.mode ?? null),
          activeMode: active?.querySelector("[data-formula-line-mode-toggle]")?.getAttribute("data-formula-line-mode") ?? null,
          focus: fields.map((field) => field.hasFocus?.() ?? false),
        };
      })()`, "Ctrl+Enter forces a new inline formula row");

  await key("Enter", "Enter", 13, 1);
  const forcedDisplay = await waitForEvaluation(`(() => {
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        const lines = persisted?.state?.lines ?? [];
        const fields = [...document.querySelectorAll("math-field")];
        return {
          ready:
            lines.length === 3 &&
            lines[2]?.mode === "display" &&
            fields[2]?.hasFocus?.() === true &&
            document.activeElement === fields[2] &&
            fields[2]?.shadowRoot?.activeElement?.getAttribute?.("part") === "keyboard-sink",
          modes: lines.map((line) => line?.mode ?? null),
          focus: fields.map((field) => field.hasFocus?.() ?? false),
          activeElementIndex: fields.indexOf(document.activeElement),
          shadowActivePart: fields.map((field) => field.shadowRoot?.activeElement?.getAttribute?.("part") ?? null),
        };
      })()`, "Alt+Enter forces a new display formula row");

  await key("Enter", "Enter", 13);
  const inheritedDisplay = await waitForEvaluation(`(() => {
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        const lines = persisted?.state?.lines ?? [];
        const rows = [...document.querySelectorAll(".formula-line")].map((row) => {
          const field = row.querySelector("math-field");
          return {
            lineId: row.getAttribute("data-line-id"),
            mode: row.querySelector("[data-formula-line-mode-toggle]")?.getAttribute("data-formula-line-mode") ?? null,
            focus: field?.hasFocus?.() ?? false,
            value: field?.value ?? null,
          };
        });
        return {
          ready:
            lines.length === 4 &&
            lines[2]?.mode === "display" &&
            lines[3]?.mode === "display",
          modes: lines.map((line) => line?.mode ?? null),
          lineIds: lines.map((line) => line?.id ?? null),
          activeLineId: persisted?.state?.activeLineId ?? null,
          rows,
        };
      })()`, "plain Enter inherits the previous row mode");

  await key("Backspace", "Backspace", 8);
  await key("Backspace", "Backspace", 8);
  await key("Backspace", "Backspace", 8);
  await waitForEvaluation(`(() => ({
        ready:
          document.querySelectorAll(".formula-line").length === 1 &&
          document.querySelector("math-field")?.value === "ab",
      }))()`, "shortcut regression cleanup");

  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.focus();
        field.position = 1;
        field.selection = { ranges: [[1, 1]], direction: "none" };
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      })()`);

  await key("Enter", "Enter", 13, 8);
  const firstBreak = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const value = field?.value?.replace(/\\s+/g, "") ?? "";
        return {
          ready:
            document.querySelectorAll("math-field").length === 1 &&
            document.querySelectorAll(".formula-line").length === 1 &&
            value.includes("\\\\begin{gathered}") &&
            value.includes("a") && value.includes("b"),
          value: field?.value ?? "",
          selection: field?.selection ?? null,
          fieldCount: document.querySelectorAll("math-field").length,
          lineCount: document.querySelectorAll(".formula-line").length,
        };
      })()`, "Shift+Enter converts one formula row into an internal gathered environment");

  await clickSelectorWithPointer('button[aria-label="撤销"]');
  const undone = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        return {
          ready:
            document.querySelectorAll("math-field").length === 1 &&
            field?.value?.replace(/\\s+/g, "") === "ab",
          value: field?.value ?? "",
          position: field?.position ?? -1,
        };
      })()`, "one undo restores the pre-multiline formula");

  await clickSelectorWithPointer('button[aria-label="重做"]');
  const redone = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const value = field?.value?.replace(/\\s+/g, "") ?? "";
        return {
          ready:
            document.querySelectorAll("math-field").length === 1 &&
            value.includes("\\\\begin{gathered}") &&
            value.includes("a") && value.includes("b") &&
            Boolean(field?.hasFocus()),
          value: field?.value ?? "",
          selection: field?.selection ?? null,
          hasFocus: field?.hasFocus() ?? false,
          selectionIsCollapsed: field?.selectionIsCollapsed ?? null,
        };
      })()`, "redo restores the internal multiline formula");

  await key("c", "KeyC", 67);
  const filledSecondRow = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const value = field?.value?.replace(/\\s+/g, "") ?? "";
        return {
          ready:
            value.includes("\\\\begin{gathered}") &&
            value.includes("a") &&
            value.includes("b") &&
            value.includes("c"),
          value: field?.value ?? "",
          position: field?.position ?? -1,
        };
      })()`, "typing fills the new internal row placeholder");

  await key("Enter", "Enter", 13, 8);
  await key("d", "KeyD", 68);
  const thirdRow = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const value = field?.value ?? "";
        const rowBreaks = (value.match(/\\\\\\\\/g) ?? []).length;
        return {
          ready:
            document.querySelectorAll("math-field").length === 1 &&
            document.querySelectorAll(".formula-line").length === 1 &&
            value.includes("\\\\begin{gathered}") &&
            value.includes("d") &&
            rowBreaks >= 2,
          value,
          rowBreaks,
          fieldCount: document.querySelectorAll("math-field").length,
          lineCount: document.querySelectorAll(".formula-line").length,
        };
      })()`, "Shift+Enter inside gathered adds another row without creating another formula field");

  console.log(JSON.stringify({
    forcedInline,
    forcedDisplay,
    inheritedDisplay,
    firstBreak,
    undone,
    redone,
    filledSecondRow,
    thirdRow,
  }, null, 2));
  console.log("Targeted line-mode shortcuts and in-formula multiline regression passed");
  return;
}
