// source-first wrappers: font and accent commands as they are typed

export const names = ["variant-core"];

export async function run({ clearField, evaluate, focusField, key, sleep, typeText }) {
  await clearField();
  await focusField();
  await typeText("\\mathbb");
  const variantCompletionBeforeSpace = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const completion = field?._mathfield?.visualTexCompletion;
        return {
          mode: field?.mode ?? "",
          raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
          suggestionIndex: field?._mathfield?.suggestionIndex ?? -1,
          suggestions: completion?.suggestions ?? [],
          records: completion?.suggestions?.map((item) => {
            const record = field?._mathfield?.visualTexCompletionRecords?.get(item);
            return { item, command: record?.command ?? "", id: record?.id ?? "" };
          }) ?? [],
        };
      })()`);
  await key(" ", "Space", 32);
  await sleep(80);
  const emptyVariant = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const host = field?.closest(".mathfield-host");
        const nativeCaret = field?.shadowRoot?.querySelector(".ML__caret");
        const latexCaret = field?.shadowRoot?.querySelector(".ML__latex-caret");
        return {
          value: field?.value ?? "",
          raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
          placeholderCount:
            field?.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder").length ?? -1,
          nativeCaret: Boolean(nativeCaret),
          latexCaret: Boolean(latexCaret),
          mode: field?.mode ?? "",
          sourceWrapperCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
          sourceWrapperKind: field?.dataset.visualtexSourceWrapperKind ?? "",
          sourceWrapperContent: field?.dataset.visualtexSourceWrapperContent ?? "",
          atoms: field?._mathfield?.model
            ? Array.from(
                { length: field._mathfield.model.lastOffset + 1 },
                (_, index) => {
                  const atom = field._mathfield.model.at(index);
                  return {
                    index,
                    type: atom?.type ?? null,
                    value: atom?.value ?? "",
                    branch: atom?.parentBranch ?? null,
                    parentType: atom?.parent?.type ?? null,
                  };
                },
              )
            : [],
        };
      })()`);
  if (
    emptyVariant.raw !== "\\mathbb{}" ||
    emptyVariant.mode !== "latex" ||
    emptyVariant.placeholderCount !== 0 ||
    !emptyVariant.latexCaret ||
    emptyVariant.sourceWrapperCommand !== "\\mathbb" ||
    emptyVariant.sourceWrapperKind !== "variant" ||
    emptyVariant.sourceWrapperContent !== ""
  ) {
    throw new Error(`Empty variant scope mismatch: ${JSON.stringify({ variantCompletionBeforeSpace, emptyVariant })}`);
  }

  await key("A", "KeyA", 65);
  await sleep(80);
  const typedVariant = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          value: field?.value ?? "",
          placeholderCount:
            field?.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder").length ?? -1,
        };
      })()`);
  if (
    typedVariant.value !== "\\mathbb{A}" ||
    typedVariant.placeholderCount !== 0
  ) {
    throw new Error(`Typed variant scope mismatch: ${JSON.stringify(typedVariant)}`);
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
  await sleep(80);
  const persistentVariant = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          value: field?.value ?? "",
          raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
          mode: field?.mode ?? "",
          sourceWrapperCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
          sourceWrapperContent: field?.dataset.visualtexSourceWrapperContent ?? "",
          placeholderCount:
            field?.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder").length ?? -1,
        };
      })()`);
  if (
    persistentVariant.raw !== "\\mathbb{AB}" ||
    persistentVariant.mode !== "latex" ||
    persistentVariant.sourceWrapperCommand !== "\\mathbb" ||
    persistentVariant.sourceWrapperContent !== "AB" ||
    persistentVariant.placeholderCount !== 0
  ) {
    throw new Error(
      `Persistent variant scope mismatch: ${JSON.stringify(persistentVariant)}`,
    );
  }
  await key("Backspace", "Backspace", 8);
  const persistentAfterOneDelete = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          value: field?.value ?? "",
          raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
          sourceWrapperCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
          sourceWrapperContent: field?.dataset.visualtexSourceWrapperContent ?? "",
          placeholderCount:
            field?.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder").length ?? -1,
        };
      })()`);
  if (
    persistentAfterOneDelete.raw !== "\\mathbb{A}" ||
    persistentAfterOneDelete.sourceWrapperCommand !== "\\mathbb" ||
    persistentAfterOneDelete.sourceWrapperContent !== "A" ||
    persistentAfterOneDelete.placeholderCount !== 0
  ) {
    throw new Error(
      `Persistent variant delete mismatch: ${JSON.stringify(persistentAfterOneDelete)}`,
    );
  }
  await key("Backspace", "Backspace", 8);
  const persistentEmpty = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          value: field?.value ?? "",
          raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
          sourceWrapperCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
          sourceWrapperContent: field?.dataset.visualtexSourceWrapperContent ?? "",
          placeholderCount:
            field?.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder").length ?? -1,
        };
      })()`);
  if (
    persistentEmpty.raw !== "\\mathbb{}" ||
    persistentEmpty.sourceWrapperCommand !== "\\mathbb" ||
    persistentEmpty.sourceWrapperContent !== "" ||
    persistentEmpty.placeholderCount !== 0
  ) {
    throw new Error(
      `Empty persistent variant mismatch: ${JSON.stringify(persistentEmpty)}`,
    );
  }
  await key("Backspace", "Backspace", 8);
  const persistentExit = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          value: field?.value ?? "",
          sourceWrapperCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
        };
      })()`);
  if (
    persistentExit.value !== "" ||
    persistentExit.sourceWrapperCommand
  ) {
    throw new Error(
      `Empty variant Backspace did not exit: ${JSON.stringify(persistentExit)}`,
    );
  }

  await clearField();
  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.dataset.visualtexAutoExitWrapperCommand = "true";
      })()`);
  await focusField();
  await typeText("\\Bbb");
  await key(" ", "Space", 32);
  const aliasEmpty = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          value: field?.value ?? "",
          raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
          sourceWrapperCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
          placeholderCount:
            field?.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder").length ?? -1,
        };
      })()`);
  await key("R", "KeyR", 82);
  const aliasTyped = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          value: field?.value ?? "",
          sourceWrapperCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
        };
      })()`);
  if (
    aliasEmpty.raw !== "\\mathbb{}" ||
    aliasEmpty.sourceWrapperCommand !== "\\mathbb" ||
    aliasEmpty.placeholderCount !== 0 ||
    aliasTyped.value !== "\\mathbb{R}" ||
    aliasTyped.sourceWrapperCommand
  ) {
    throw new Error(
      `Variant alias mismatch: ${JSON.stringify({ aliasEmpty, aliasTyped })}`,
    );
  }

  await clearField();
  await focusField();
  await typeText("\\mathcal");
  await key(" ", "Space", 32);
  await key("g", "KeyG", 71);
  const lowercaseCalligraphic = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return { value: field?.value ?? "" };
      })()`);
  if (lowercaseCalligraphic.value !== "\\mathscr{g}") {
    throw new Error(
      `Lowercase mathcal compatibility mismatch: ${JSON.stringify(lowercaseCalligraphic)}`,
    );
  }

  await clearField();
  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.dataset.visualtexAutoExitWrapperCommand = "true";
      })()`);
  await focusField();
  await typeText("\\mathbb");
  await key(" ", "Space", 32);
  await typeText("\\alpha");
  await key(" ", "Space", 32);
  await sleep(100);
  const commandAutoExit = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          sourceWrapperCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
          raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
        };
      })()`);
  if (
    commandAutoExit.value !== "\\mathbb{\\alpha}" ||
    commandAutoExit.mode !== "math" ||
    commandAutoExit.sourceWrapperCommand ||
    commandAutoExit.raw
  ) {
    throw new Error(`Variant command auto-exit mismatch: ${JSON.stringify(commandAutoExit)}`);
  }

  await clearField();
  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.dataset.visualtexAutoExitWrapperCommand = "false";
      })()`);
  await focusField();
  await typeText("\\mathbb");
  await key(" ", "Space", 32);
  await typeText("\\alpha");
  await key(" ", "Space", 32);
  await sleep(80);
  const commandBeforeEnter = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          sourceWrapperCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
          sourceWrapperContent: field?.dataset.visualtexSourceWrapperContent ?? "",
          raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
        };
      })()`);
  if (
    commandBeforeEnter.mode !== "latex" ||
    commandBeforeEnter.sourceWrapperCommand !== "\\mathbb" ||
    commandBeforeEnter.sourceWrapperContent !== "\\alpha" ||
    commandBeforeEnter.raw !== "\\mathbb{\\alpha}"
  ) {
    throw new Error(`Persistent variant command mismatch: ${JSON.stringify(commandBeforeEnter)}`);
  }
  await key("Enter", "Enter", 13);
  await sleep(80);
  const commandAfterEnter = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          sourceWrapperCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
        };
      })()`);
  if (
    commandAfterEnter.value !== "\\mathbb{\\alpha}" ||
    commandAfterEnter.mode !== "math" ||
    commandAfterEnter.sourceWrapperCommand
  ) {
    throw new Error(`Persistent variant Enter mismatch: ${JSON.stringify(commandAfterEnter)}`);
  }

  await clearField();
  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.dataset.visualtexAutoExitAccent = "true";
      })()`);
  await focusField();
  await typeText("\\vec");
  await key(" ", "Space", 32);
  await sleep(80);
  const emptyAccent = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          sourceWrapperCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
          sourceWrapperKind: field?.dataset.visualtexSourceWrapperKind ?? "",
          placeholderCount:
            field?.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder").length ?? -1,
          latexCaret: Boolean(field?.shadowRoot?.querySelector(".ML__latex-caret")),
          raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
        };
      })()`);
  if (
    emptyAccent.value !== "\\vec{}" ||
    emptyAccent.raw !== "\\vec{}" ||
    emptyAccent.mode !== "latex" ||
    emptyAccent.sourceWrapperCommand !== "\\vec" ||
    emptyAccent.sourceWrapperKind !== "accent" ||
    emptyAccent.placeholderCount !== 0 ||
    !emptyAccent.latexCaret
  ) {
    throw new Error(`Empty accent source shell mismatch: ${JSON.stringify(emptyAccent)}`);
  }
  await key("x", "KeyX", 88);
  await sleep(80);
  const accentAutoExit = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          sourceWrapperCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
          placeholderCount:
            field?.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder").length ?? -1,
        };
      })()`);
  if (
    accentAutoExit.value !== "\\vec{x}" ||
    accentAutoExit.mode !== "math" ||
    accentAutoExit.sourceWrapperCommand ||
    accentAutoExit.placeholderCount !== 0
  ) {
    throw new Error(`Accent character auto-exit mismatch: ${JSON.stringify(accentAutoExit)}`);
  }

  await clearField();
  await focusField();
  await typeText("\\vec");
  await key(" ", "Space", 32);
  await typeText("\\alpha");
  await key(" ", "Space", 32);
  await sleep(80);
  const accentCommandAutoExit = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          sourceWrapperCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
        };
      })()`);
  if (
    accentCommandAutoExit.value !== "\\vec{\\alpha}" ||
    accentCommandAutoExit.mode !== "math" ||
    accentCommandAutoExit.sourceWrapperCommand
  ) {
    throw new Error(`Accent command auto-exit mismatch: ${JSON.stringify(accentCommandAutoExit)}`);
  }

  await clearField();
  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.dataset.visualtexAutoExitAccent = "false";
      })()`);
  await focusField();
  await typeText("\\hat");
  await key(" ", "Space", 32);
  await key("x", "KeyX", 88);
  await key("y", "KeyY", 89);
  await sleep(80);
  const persistentAccent = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          sourceWrapperCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
          sourceWrapperContent: field?.dataset.visualtexSourceWrapperContent ?? "",
          raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
        };
      })()`);
  if (
    persistentAccent.value !== "\\hat{xy}" ||
    persistentAccent.raw !== "\\hat{xy}" ||
    persistentAccent.mode !== "latex" ||
    persistentAccent.sourceWrapperCommand !== "\\hat" ||
    persistentAccent.sourceWrapperContent !== "xy"
  ) {
    throw new Error(`Persistent accent source shell mismatch: ${JSON.stringify(persistentAccent)}`);
  }
  await key("Enter", "Enter", 13);
  await sleep(80);
  const accentAfterEnter = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          sourceWrapperCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
        };
      })()`);
  if (
    accentAfterEnter.value !== "\\hat{xy}" ||
    accentAfterEnter.mode !== "math" ||
    accentAfterEnter.sourceWrapperCommand
  ) {
    throw new Error(`Persistent accent Enter mismatch: ${JSON.stringify(accentAfterEnter)}`);
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
  await sleep(60);
  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        window.__visualTexSourceStateA = field?._mathfield?.model.getState();
      })()`);
  await key("B", "KeyB", 66);
  await sleep(60);
  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        window.__visualTexSourceStateAB = field?._mathfield?.model.getState();
      })()`);
  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field?._mathfield?.model.setState(window.__visualTexSourceStateA, {
          silenceNotifications: false,
          type: "undo",
        });
      })()`);
  await sleep(80);
  const sourceHistoryA = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const mathfield = field?._mathfield;
        const wrapper = mathfield?.visualTexSourceWrapper;
        const groupOffset =
          wrapper?.group?.parent && mathfield?.model
            ? mathfield.model.offsetOf(wrapper.group)
            : -1;
        return {
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          sourceCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
          sourceContent: field?.dataset.visualtexSourceWrapperContent ?? "",
          groupOffset,
          groupIsCurrent:
            groupOffset >= 0 && mathfield?.model?.at(groupOffset) === wrapper?.group,
        };
      })()`);
  if (
    sourceHistoryA.value !== "\\mathbb{A}" ||
    sourceHistoryA.mode !== "latex" ||
    sourceHistoryA.sourceCommand !== "\\mathbb" ||
    sourceHistoryA.sourceContent !== "A" ||
    sourceHistoryA.groupOffset < 0 ||
    !sourceHistoryA.groupIsCurrent
  ) {
    throw new Error(
      `Source-wrapper history restore A mismatch: ${JSON.stringify(sourceHistoryA)}`,
    );
  }

  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field?._mathfield?.model.setState(window.__visualTexSourceStateAB, {
          silenceNotifications: false,
          type: "redo",
        });
      })()`);
  await sleep(80);
  const sourceHistoryAB = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const mathfield = field?._mathfield;
        const wrapper = mathfield?.visualTexSourceWrapper;
        const groupOffset =
          wrapper?.group?.parent && mathfield?.model
            ? mathfield.model.offsetOf(wrapper.group)
            : -1;
        return {
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          sourceCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
          sourceContent: field?.dataset.visualtexSourceWrapperContent ?? "",
          groupOffset,
          groupIsCurrent:
            groupOffset >= 0 && mathfield?.model?.at(groupOffset) === wrapper?.group,
        };
      })()`);
  if (
    sourceHistoryAB.value !== "\\mathbb{AB}" ||
    sourceHistoryAB.mode !== "latex" ||
    sourceHistoryAB.sourceCommand !== "\\mathbb" ||
    sourceHistoryAB.sourceContent !== "AB" ||
    sourceHistoryAB.groupOffset < 0 ||
    !sourceHistoryAB.groupIsCurrent
  ) {
    throw new Error(
      `Source-wrapper history restore AB mismatch: ${JSON.stringify(sourceHistoryAB)}`,
    );
  }

  console.log(JSON.stringify({
    emptyVariant,
    typedVariant,
    commandAutoExit,
    commandBeforeEnter,
    commandAfterEnter,
    emptyAccent,
    accentAutoExit,
    accentCommandAutoExit,
    persistentAccent,
    accentAfterEnter,
  }, null, 2));
  console.log("Targeted source-wrapper core regression passed");
  return;
}
