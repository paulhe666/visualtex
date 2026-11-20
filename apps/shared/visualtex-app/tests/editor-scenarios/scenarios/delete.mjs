// Word-like deletion and line editing

export const names = ["delete"];

export async function run({ evaluate, key, sleep, typeText, waitForEvaluation }) {
  await evaluate(`(() => {
        const storageKey = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
        const first = { id: crypto.randomUUID(), latex: "a" };
        const second = { id: crypto.randomUUID(), latex: "" };
        persisted.state = {
          ...(persisted.state || {}),
          lines: [first, second],
          activeLineId: second.id,
        };
        localStorage.setItem(storageKey, JSON.stringify(persisted));
        location.reload();
      })()`);
  await waitForEvaluation(`(() => ({
        ready: document.querySelectorAll("math-field").length === 2,
      }))()`, "two formula lines for delete test");
  await evaluate(`(() => {
        const field = document.querySelectorAll("math-field")[1];
        field.setValue("", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.mode = "math";
        field.position = field.lastOffset;
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        field.dispatchEvent(new FocusEvent("focus", { bubbles: true, composed: true }));
      })()`);
  await sleep(120);
  await typeText("\\mat");
  const beforeDelete = await waitForEvaluation(`(() => {
        const field = document.querySelectorAll("math-field")[1];
        // the raw LaTeX typed before the caret (the kernel's latex atoms)
        const model = field._mathfield?.model;
        let typedRaw = "";
        for (let offset = 1; model && offset <= field.position; offset++) {
          const atom = model.at(offset);
          if (atom?.type === "latex" && !atom.isSuggestion) typedRaw += atom.value;
        }
        const renderedRaw = [...(field.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
          .map((node) => node.textContent ?? "")
          .join("");
        return {
          ready: typedRaw === "\\\\mat" && document.querySelectorAll("math-field").length === 2,
          typedRaw,
          renderedRaw,
          lineCount: document.querySelectorAll("math-field").length,
          contentHtml: field.shadowRoot?.querySelector('[part="content"]')?.innerHTML ?? "",
        };
      })()`, "raw math command before Backspace");
  await key("Backspace", "Backspace", 8);
  const firstDelete = await waitForEvaluation(`(() => {
        const field = document.querySelectorAll("math-field")[1];
        // the raw LaTeX typed before the caret (the kernel's latex atoms)
        const model = field._mathfield?.model;
        let typedRaw = "";
        for (let offset = 1; model && offset <= field.position; offset++) {
          const atom = model.at(offset);
          if (atom?.type === "latex" && !atom.isSuggestion) typedRaw += atom.value;
        }
        return {
          ready: typedRaw === "\\\\ma" && document.querySelectorAll("math-field").length === 2,
          typedRaw,
          lineCount: document.querySelectorAll("math-field").length,
          firstValue: document.querySelectorAll("math-field")[0]?.value ?? "",
        };
      })()`, "one-character raw command deletion");
  await key("Backspace", "Backspace", 8);
  const secondDelete = await waitForEvaluation(`(() => {
        const field = document.querySelectorAll("math-field")[1];
        // the raw LaTeX typed before the caret (the kernel's latex atoms)
        const model = field._mathfield?.model;
        let typedRaw = "";
        for (let offset = 1; model && offset <= field.position; offset++) {
          const atom = model.at(offset);
          if (atom?.type === "latex" && !atom.isSuggestion) typedRaw += atom.value;
        }
        return {
          ready: typedRaw === "\\\\m" && document.querySelectorAll("math-field").length === 2,
          typedRaw,
          lineCount: document.querySelectorAll("math-field").length,
        };
      })()`, "second one-character raw command deletion");

  await key("Backspace", "Backspace", 8);
  await waitForEvaluation(`(() => {
        const field = document.querySelectorAll("math-field")[1];
        // the raw LaTeX typed before the caret (the kernel's latex atoms)
        const model = field._mathfield?.model;
        let typedRaw = "";
        for (let offset = 1; model && offset <= field.position; offset++) {
          const atom = model.at(offset);
          if (atom?.type === "latex" && !atom.isSuggestion) typedRaw += atom.value;
        }
        return { ready: typedRaw === "\\\\", typedRaw };
      })()`, "raw command reduced to backslash");
  await key("Backspace", "Backspace", 8);
  await waitForEvaluation(`(() => {
        const field = document.querySelectorAll("math-field")[1];
        const rawText = [...(field.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
          .map((node) => node.textContent ?? "")
          .join("");
        return {
          ready: rawText === "" && document.querySelectorAll("math-field").length === 2,
          rawText,
          mode: field.mode,
        };
      })()`, "raw command fully cleared without deleting row");
  await key("Backspace", "Backspace", 8);
  const emptyLineDelete = await waitForEvaluation(`(() => {
        const fields = [...document.querySelectorAll("math-field")];
        return {
          ready:
            fields.length === 1 &&
            fields[0]?.value === "a" &&
            fields[0]?.hasFocus?.(),
          lineCount: fields.length,
          values: fields.map((field) => field.value),
          activeTag: document.activeElement?.tagName ?? "",
        };
      })()`, "empty raw-latex row returns to previous formula");

  console.log(JSON.stringify({ beforeDelete, firstDelete, secondDelete, emptyLineDelete }, null, 2));
  console.log("Targeted delete regression passed");
  return;
}
