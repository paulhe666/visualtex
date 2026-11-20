// empty slots: entering, filling and leaving them

export const names = ["placeholder-core"];

export async function run({ clearField, client, evaluate, key, sleep, typeText }) {
  await clearField();
  const state = await evaluate(`(async () => {
        const field = document.querySelector("math-field");
        const sink = field?.shadowRoot?.querySelector('[part="keyboard-sink"]');
        const inspect = (phase) => {
          // The layout engine draws the empty slot: its selected marker and
          // the caret bar standing against the marker's left edge.
          const marker = field?.shadowRoot?.querySelector(".ML__vtx-marker--selected");
          const caret = marker ? field.shadowRoot.querySelector(".ML__caret.ML__vtx-caret--before") : null;
          const caretRect = caret?.getBoundingClientRect() ?? null;
          const markerLeft = marker?.getBoundingClientRect().left ?? null;
          const caretPseudoStyle = caret ? getComputedStyle(caret, "::after") : null;
          return {
            phase,
            value: field?.value ?? "",
            placeholderCount:
              field?.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder").length ?? -1,
            caretPresent: Boolean(caret),
            caretTouchesMarker: Boolean(caretRect && markerLeft !== null && Math.abs(caretRect.right - markerLeft) <= 0.75),
            caretX: caretRect?.left ?? null,
            markerLeft,
            caretAnimation: caretPseudoStyle?.animationName ?? "",
          };
        };
        field.setValue("\\\\frac{\\\\placeholder{}}{\\\\placeholder{}}", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "placeholder",
          silenceNotifications: true,
        });
        field.focus();
        sink?.focus({ preventScroll: true });
        await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        const fraction = inspect("fraction");

        field.setValue("\\\\begin{bmatrix}\\\\placeholder{}&y\\\\\\\\z&t\\\\end{bmatrix}", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "placeholder",
          silenceNotifications: true,
        });
        field.insert("x", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceSelection",
          selectionMode: "after",
          focus: true,
          scrollIntoView: false,
        });
        field.executeCommand("deleteBackward");
        await new Promise((resolve) => requestAnimationFrame(resolve));
        const matrix = inspect("matrix-delete-last");

        field.setValue("\\\\int \\\\placeholder{}\\\\,dx", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "placeholder",
          silenceNotifications: true,
        });
        field.focus();
        sink?.focus({ preventScroll: true });
        return { fraction, matrix, integrandReady: inspect("integrand-ready") };
      })()`);
  await key("q", "KeyQ", 81);
  const integrandFirstInput = await evaluate("(() => { const field = document.querySelector('math-field'); const model = field?._mathfield?.model; const q = model ? Array.from({length:model.lastOffset+1},(_,i)=>model.at(i)).find(atom=>atom?.value==='q') : null; return { value: field?.value ?? '', qHasSlot: Boolean(q?.visualTexStructuralSlot), qSlotParentIsRoot: Boolean(q?.visualTexStructuralSlot?.parent === model?.root), qSlotParentType: q?.visualTexStructuralSlot?.parent?.type ?? null, qSlotBranch: q?.visualTexStructuralSlot?.branch ?? null, qParentType: q?.parent?.type ?? null, qParentBranch: q?.parentBranch ?? null }; })()");
  await key("r", "KeyR", 82);
  await key("Backspace", "Backspace", 8);
  await sleep(60);
  const integrandPartialDelete = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const model = field?._mathfield?.model;
        const q = model
          ? Array.from({ length: model.lastOffset + 1 }, (_, index) => model.at(index))
              .find((atom) => atom?.value === "q")
          : null;
        return {
          value: field?.value ?? "",
          placeholderCount:
            field?.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder").length ?? -1,
          qHasSlot: Boolean(q?.visualTexStructuralSlot),
          qSlotParentIsRoot: Boolean(q?.visualTexStructuralSlot?.parent === model?.root),
          qSlotBranch: q?.visualTexStructuralSlot?.branch ?? null,
        };
      })()`);
  if (
    !integrandPartialDelete.value.includes("q") ||
    integrandPartialDelete.placeholderCount !== 0 ||
    !integrandPartialDelete.qHasSlot ||
    !integrandPartialDelete.qSlotParentIsRoot ||
    integrandPartialDelete.qSlotBranch !== "body"
  ) {
    throw new Error(
      `Structural slot restored too early: ${JSON.stringify(integrandPartialDelete)}`,
    );
  }
  await key("Backspace", "Backspace", 8);
  await sleep(80);
  const integrand = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const caret = field?.shadowRoot?.querySelector(".ML__vtx-marker--selected")
          ? field.shadowRoot.querySelector(".ML__caret.ML__vtx-caret--before")
          : null;
        return {
          value: field?.value ?? "",
          placeholderCount:
            field?.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder").length ?? -1,
          caretPresent: Boolean(caret),
          selection: field?.selection ?? null,
        };
      })()`);
  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.setValue("\\\\placeholder{}d\\\\placeholder{}", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      })()`);
  await sleep(80);
  const geometryBefore = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const placeholders = [...(field?.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder") ?? [])];
        const node = placeholders[0];
        const rect = node?.getBoundingClientRect();
        return {
          count: placeholders.length,
          x: rect?.x ?? 0,
          y: rect?.y ?? 0,
          width: rect?.width ?? 0,
          height: rect?.height ?? 0,
          clickX: rect ? rect.left + rect.width / 2 : 0,
          clickY: rect ? rect.top + rect.height / 2 : 0,
        };
      })()`);
  await client.send("Input.dispatchMouseEvent", {
    type: "mousePressed",
    x: geometryBefore.clickX,
    y: geometryBefore.clickY,
    button: "left",
    clickCount: 1,
  });
  await client.send("Input.dispatchMouseEvent", {
    type: "mouseReleased",
    x: geometryBefore.clickX,
    y: geometryBefore.clickY,
    button: "left",
    clickCount: 1,
  });
  await sleep(80);
  const geometryAfter = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const placeholders = [...(field?.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder") ?? [])];
        const selected = placeholders.find((node) =>
          node.classList.contains("ML__selected") || Boolean(node.closest(".ML__selected")));
        const node = selected ?? placeholders[0];
        const rect = node?.getBoundingClientRect();
        return {
          count: placeholders.length,
          selected: Boolean(selected),
          x: rect?.x ?? 0,
          y: rect?.y ?? 0,
          width: rect?.width ?? 0,
          height: rect?.height ?? 0,
          caret: Boolean(
            field?.shadowRoot?.querySelector(".ML__vtx-marker--selected") &&
            field.shadowRoot.querySelector(".ML__caret.ML__vtx-caret--before"),
          ),
        };
      })()`);
  const geometryDelta = {
    x: Math.abs(geometryAfter.x - geometryBefore.x),
    y: Math.abs(geometryAfter.y - geometryBefore.y),
    width: Math.abs(geometryAfter.width - geometryBefore.width),
    height: Math.abs(geometryAfter.height - geometryBefore.height),
  };

  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.setValue("\\\\int_a^b \\\\placeholder{}\\\\,dx", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "placeholder",
          silenceNotifications: true,
        });
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      })()`);
  await typeText("\\frac");
  await key(" ", "Space", 32);
  await sleep(80);
  const nestedSlotBeforeDelete = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const model = field?._mathfield?.model;
        const fraction = model
          ? Array.from({ length: model.lastOffset + 1 }, (_, index) => model.at(index))
              .find((atom) => atom?.type === "genfrac")
          : null;
        const offset = fraction ? model.offsetOf(fraction) : -1;
        const slot = fraction?.visualTexStructuralSlot ?? null;
        const hasSlot = Boolean(slot);
        const slotParentIsRoot = Boolean(slot?.parent === model?.root);
        const slotAfterOffset = slot?.after ? model?.offsetOf(slot.after) ?? -1 : -1;
        const slotBranch = slot?.branch ?? null;
        const beforeValue = field?.value ?? "";
        let selectionBeforeDelete = field?.selection ?? null;
        let selectedLatex = "";
        let deleteAccepted = false;
        if (fraction && offset > 0) {
          model.setSelection(model.offsetOf(fraction.leftSibling), offset);
          selectionBeforeDelete = field.selection;
          selectedLatex = field.getValue(field.selection, "latex");
          var selectedAtoms = model.getAtoms([
            model.offsetOf(fraction.leftSibling),
            offset,
          ]).map((atom) => ({
            type: atom.type,
            offset: model.offsetOf(atom),
            hasSlot: Boolean(atom.visualTexStructuralSlot),
            parentType: atom.parent?.type ?? null,
          }));
          deleteAccepted = field.executeCommand("deleteBackward");
        }
        return {
          beforeValue,
          foundFraction: Boolean(fraction),
          hasSlot,
          slotParentIsRoot,
          slotAfterOffset,
          slotBranch,
          offset,
          selectionBeforeDelete,
          selectedLatex,
          selectedAtoms: typeof selectedAtoms === "undefined" ? [] : selectedAtoms,
          deleteAccepted,
        };
      })()`);
  await sleep(80);
  const nestedSlotAfterDelete = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          value: field?.value ?? "",
          placeholderCount:
            field?.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder").length ?? -1,
          caretPresent: Boolean(
            field?.shadowRoot?.querySelector(".ML__vtx-marker--selected") &&
            field.shadowRoot.querySelector(".ML__caret.ML__vtx-caret--before"),
          ),
        };
      })()`);

  if (
    !state.fraction.caretPresent ||
    !state.fraction.caretTouchesMarker ||
    state.fraction.caretX == null ||
    state.fraction.markerLeft == null ||
    state.fraction.caretX > state.fraction.markerLeft + 0.75 ||
    !state.fraction.caretAnimation.includes("ML__caret-blink") ||
    state.matrix.placeholderCount !== 1 ||
    !state.matrix.caretPresent ||
    integrand.placeholderCount !== 1 ||
    !integrand.caretPresent ||
    !integrand.value.includes("\\placeholder{}") ||
    geometryBefore.count !== 2 ||
    geometryAfter.count !== 2 ||
    !geometryAfter.selected ||
    !geometryAfter.caret ||
    Math.max(...Object.values(geometryDelta)) > 0.5 ||
    !nestedSlotBeforeDelete.foundFraction ||
    !nestedSlotBeforeDelete.hasSlot ||
    nestedSlotAfterDelete.placeholderCount !== 1 ||
    !nestedSlotAfterDelete.caretPresent
  ) {
    throw new Error(
      `Placeholder core regression: ${JSON.stringify({
            state,
            integrandPartialDelete,
            integrand,
            geometryBefore,
            geometryAfter,
            geometryDelta,
            nestedSlotBeforeDelete,
            nestedSlotAfterDelete,
          })}`,
    );
  }
  console.log(JSON.stringify({
    state,
    integrand,
    geometryBefore,
    geometryAfter,
    geometryDelta,
    nestedSlotBeforeDelete,
    nestedSlotAfterDelete,
  }, null, 2));
  console.log("Targeted placeholder core regression passed");
  return;
}
