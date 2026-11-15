// more source-first wrappers

export const names = ["variant-core-extended"];

export async function run({ clearField, evaluate, focusField, key, sleep, typeText }) {
  await clearField();
  await focusField();
  await typeText("\\mathbb");
  await key(" ", "Space", 32);
  await sleep(80);
  const empty = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const host = field?.closest(".mathfield-host");
        return {
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
          sourceCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
          sourceContent: field?.dataset.visualtexSourceWrapperContent ?? "",
          placeholderCount: field?.shadowRoot?.querySelectorAll(
            ".visualtex-structural-placeholder, .ML__placeholder",
          ).length ?? -1,
          nativeCaretCount: field?.shadowRoot?.querySelectorAll(
            ".ML__caret, .ML__text-caret, .ML__latex-caret",
          ).length ?? -1,
        };
      })()`);
  if (
    empty.value !== "\\mathbb{}" ||
    empty.raw !== "\\mathbb{}" ||
    empty.mode !== "latex" ||
    empty.sourceCommand !== "\\mathbb" ||
    empty.sourceContent !== "" ||
    empty.placeholderCount !== 0 ||
    empty.nativeCaretCount < 1
  ) {
    throw new Error(`Empty variant scope regression: ${JSON.stringify(empty)}`);
  }
  await key("A", "KeyA", 65);
  await sleep(80);
  const single = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          value: field?.value ?? "",
          placeholderCount: field?.shadowRoot?.querySelectorAll(
            ".visualtex-structural-placeholder, .ML__placeholder",
          ).length ?? -1,
        };
      })()`);
  if (
    single.value !== "\\mathbb{A}" ||
    single.placeholderCount !== 0
  ) {
    throw new Error(`Single-character variant regression: ${JSON.stringify(single)}`);
  }

  await clearField();
  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.dataset.visualtexAutoExitWrapperCommand = "false";
      })()`);
  await focusField();
  await typeText("\\mathbb");
  await key(" ", "Space", 32);
  await key("A", "KeyA", 65);
  await key("B", "KeyB", 66);
  await key("Backspace", "Backspace", 8);
  await sleep(80);
  const continuousBeforeEnter = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const host = field?.closest(".mathfield-host");
        return {
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          sourceCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
          content: field?.dataset.visualtexSourceWrapperContent ?? "",
          fieldCount: document.querySelectorAll("math-field").length,
        };
      })()`);
  if (
    continuousBeforeEnter.value !== "\\mathbb{A}" ||
    continuousBeforeEnter.mode !== "latex" ||
    continuousBeforeEnter.sourceCommand !== "\\mathbb" ||
    continuousBeforeEnter.content !== "A"
  ) {
    throw new Error(
      `Continuous variant delete regression: ${JSON.stringify(continuousBeforeEnter)}`,
    );
  }
  await key("Enter", "Enter", 13);
  await sleep(80);
  const continuousAfterEnter = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const host = field?.closest(".mathfield-host");
        return {
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          sourceCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
          fieldCount: document.querySelectorAll("math-field").length,
        };
      })()`);
  if (
    continuousAfterEnter.value !== "\\mathbb{A}" ||
    continuousAfterEnter.mode !== "math" ||
    continuousAfterEnter.sourceCommand ||
    continuousAfterEnter.fieldCount !== continuousBeforeEnter.fieldCount
  ) {
    throw new Error(
      `Variant Enter confirmation regression: ${JSON.stringify({ continuousBeforeEnter, continuousAfterEnter })}`,
    );
  }

  await clearField();
  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.dataset.visualtexAutoExitWrapperCommand = "false";
      })()`);
  await focusField();
  await typeText("\\mathbb");
  await key(" ", "Space", 32);
  await key("A", "KeyA", 65);
  await typeText("\\frac");
  await key(" ", "Space", 32);
  await sleep(80);
  const nested = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const model = field?._mathfield?.model;
        const placeholders = model
          ? Array.from({ length: model.lastOffset + 1 }, (_, index) => model.at(index))
              .filter((atom) => atom?.type === "placeholder")
              .map((atom) => ({
                branch: Array.isArray(atom.parentBranch)
                  ? JSON.stringify(atom.parentBranch)
                  : atom.parentBranch,
                style: atom.style,
              }))
          : [];
        return {
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          sourceCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
          sourceContent: field?.dataset.visualtexSourceWrapperContent ?? "",
          raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
          position: model?.position ?? -1,
          atValue: model?.at(model.position)?.value ?? "",
          nextValue: model?.at(model.position + 1)?.value ?? "",
          placeholders,
        };
      })()`);
  if (
    nested.value !== "\\mathbb{A\\frac{}{}}" ||
    nested.raw !== "\\mathbb{A\\frac{}{}}" ||
    nested.mode !== "latex" ||
    nested.sourceCommand !== "\\mathbb" ||
    nested.sourceContent !== "A\\frac{}{}" ||
    nested.placeholders.length !== 0 ||
    nested.atValue !== "{" ||
    nested.nextValue !== "}"
  ) {
    throw new Error(`Nested variant command regression: ${JSON.stringify(nested)}`);
  }
  console.log(
    JSON.stringify(
      { empty, single, continuousBeforeEnter, continuousAfterEnter, nested },
      null,
      2,
    ),
  );
  console.log("Targeted extended variant core regression passed");
  return;
}
