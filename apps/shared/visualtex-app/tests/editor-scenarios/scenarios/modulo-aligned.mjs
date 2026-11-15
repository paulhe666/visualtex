// \pmod and \mod inside aligned rows

export const names = ["modulo-aligned"];

export async function run({ clearField, evaluate, key, typeText, waitForEvaluation }) {
  const issue15Source = String.raw`\begin{aligned}
\langle p_1,p_0\rangle &\leftarrow \operatorname{umul}(a,b)=ab
&&\text{Double word product}\\
p_0 &\leftarrow \operatorname{umullo}(a,b)=(ab)\bmod\beta
&&\text{Low word}\\
p_1 &\leftarrow \operatorname{umulhi}(a,b)=\left\lfloor\frac{ab}{\beta}\right\rfloor
&&\text{High word.}
\end{aligned}`;

  await clearField();
  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.setValue(${JSON.stringify(issue15Source)}, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.dispatchEvent(new InputEvent("input", {
          bubbles: true,
          composed: true,
          inputType: "insertFromPaste",
        }));
        return true;
      })()`);
  const issue15State = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const content = field?.shadowRoot?.querySelector('[part="content"]');
        const box = content?.getBoundingClientRect();
        const errors = Array.isArray(field?.errors) ? field.errors : [];
        const value = field?.value ?? "";
        return {
          ready:
            value.includes("\\\\begin{aligned}") &&
            value.includes("\\\\bmod") &&
            errors.length === 0 &&
            !content?.querySelector(".ML__error") &&
            Number.isFinite(box?.width) &&
            Number.isFinite(box?.height) &&
            box.width > 400 &&
            box.height > 70,
          value,
          errors,
          contentText: content?.textContent ?? "",
          width: box?.width ?? 0,
          height: box?.height ?? 0,
          errorAtoms: content?.querySelectorAll(".ML__error").length ?? -1,
        };
      })()`, "Issue #15 aligned formula with \\bmod");

  await clearField();
  await typeText(String.raw`\bmod`);
  const suggestionState = await waitForEvaluation(`(() => {
        const nativeItems = [...document.querySelectorAll(
          '#mathlive-suggestion-popover li[data-command]',
        )];
        const visualItems = [...document.querySelectorAll(
          '.suggestion-item .suggestion-command',
        )];
        return {
          ready:
            nativeItems.some((item) => item.dataset.command === "\\\\bmod") ||
            visualItems.some((item) => item.textContent?.trim() === "\\\\bmod"),
          nativeCommands: nativeItems.map((item) => item.dataset.command ?? ""),
          visualCommands: visualItems.map((item) => item.textContent?.trim() ?? ""),
        };
      })()`, "\\bmod command suggestion");
  await key(" ", "Space", 32);
  await typeText("b");
  const typedState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const content = field?.shadowRoot?.querySelector('[part="content"]');
        const value = field?.value ?? "";
        const errors = Array.isArray(field?.errors) ? field.errors : [];
        return {
          ready:
            value.includes("\\\\bmod") &&
            value.includes("b") &&
            errors.length === 0 &&
            !content?.querySelector(".ML__error"),
          value,
          errors,
          contentText: content?.textContent ?? "",
        };
      })()`, "typed \\bmod operator remains editable");

  console.log(JSON.stringify({ issue15State, suggestionState, typedState }, null, 2));
  console.log("Targeted modulo/aligned editor regression passed");
  return;
}
