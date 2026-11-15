// typing inside nested structures

export const names = ["nested-input-core"];

export async function run({ clearField, client, evaluate, focusField, key, sleep }) {
  await clearField();
  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.setValue("\\\\frac{\\\\placeholder{}}{\\\\placeholder{}}", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "placeholder",
          silenceNotifications: true,
        });
        field.focus();
        field.shadowRoot
          ?.querySelector('[part="keyboard-sink"]')
          ?.focus({ preventScroll: true });
        const cancel = (event) => {
          event.preventDefault();
          field.removeEventListener("beforeinput", cancel, true);
        };
        field.addEventListener("beforeinput", cancel, true);
      })()`);
  await key("x", "KeyX", 88);
  const cancelled = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          value: field?.value ?? "",
          placeholderCount:
            field?.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder").length ?? -1,
          selection: field?.selection ?? null,
        };
      })()`);
  if (
    cancelled.value !== "\\frac{\\placeholder{}}{\\placeholder{}}" ||
    cancelled.placeholderCount !== 2
  ) {
    throw new Error(`Cancelled input mutated placeholder: ${JSON.stringify(cancelled)}`);
  }

  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const model = field?._mathfield?.model;
        field.dataset.visualtexAutoExitAccent = "true";
        field.setValue("\\\\hat{\\\\frac{a}{c}}", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        const a = Array.from(
          { length: model.lastOffset + 1 },
          (_, index) => model.at(index),
        ).find(
          (atom) =>
            atom?.value === "a" &&
            atom?.parentBranch === "above" &&
            atom?.parent?.type === "genfrac",
        );
        field.position = model.offsetOf(a);
        field.focus();
        field.shadowRoot
          ?.querySelector('[part="keyboard-sink"]')
          ?.focus({ preventScroll: true });
      })()`);
  await key("b", "KeyB", 66);
  const nestedAccent = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const model = field?._mathfield?.model;
        const atom = model?.at(model.position);
        return {
          value: field?.value ?? "",
          position: model?.position ?? -1,
          atomValue: atom?.value ?? "",
          branch: atom?.parentBranch ?? null,
          parentType: atom?.parent?.type ?? null,
          grandParentType: atom?.parent?.parent?.type ?? null,
        };
      })()`);
  if (
    nestedAccent.value !== "\\hat{\\frac{ab}{c}}" ||
    nestedAccent.atomValue !== "b" ||
    nestedAccent.branch !== "above" ||
    nestedAccent.parentType !== "genfrac"
  ) {
    throw new Error(`Nested accent auto-exit escaped too far: ${JSON.stringify(nestedAccent)}`);
  }

  await clearField();
  await focusField();
  const fastKey = async (value, code, virtualKeyCode) => {
    const common = {
      key: value,
      code,
      windowsVirtualKeyCode: virtualKeyCode,
      nativeVirtualKeyCode: virtualKeyCode,
    };
    await client.send("Input.dispatchKeyEvent", {
      type: "keyDown",
      ...common,
      ...(value.length === 1 ? { text: value, unmodifiedText: value } : {}),
    });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
  };
  await fastKey("\\", "Backslash", 220);
  await fastKey("f", "KeyF", 70);
  await fastKey("r", "KeyR", 82);
  await fastKey("a", "KeyA", 65);
  await fastKey("c", "KeyC", 67);
  await fastKey(" ", "Space", 32);
  await fastKey("\\", "Backslash", 220);
  await sleep(80);
  const fastNested = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const rawLatex = [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
          .map((node) => node.textContent ?? "")
          .join("");
        return {
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          rawLatex,
          selection: field?.selection ?? null,
        };
      })()`);
  if (
    !fastNested.value.includes("\\frac") ||
    fastNested.mode !== "latex" ||
    !fastNested.rawLatex.includes("\\")
  ) {
    throw new Error(`Fast nested backslash was swallowed: ${JSON.stringify(fastNested)}`);
  }

  console.log(
    JSON.stringify({ cancelled, nestedAccent, fastNested }, null, 2),
  );
  console.log("Targeted nested input core regression passed");
  return;
}
