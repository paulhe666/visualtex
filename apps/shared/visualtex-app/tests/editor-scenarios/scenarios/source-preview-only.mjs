// the source panel drives the formulas live
import assert from "node:assert/strict";

export const names = ["source-preview-only"];

// state the page starts with (merged into the editor's persisted state)
export const storage = () => ({ editorLayout: "standard", sourceOpen: false, inputBehavior: { showOtherCommandSuggestions: true } });

export async function run({ clearField, client, evaluate, key, sleep, typeText, waitForEvaluation }) {
  await waitForEvaluation(`(() => ({
        ready:
          document.querySelector(".workspace")?.dataset.editorLayout === "standard" &&
          Boolean(document.querySelector("math-field")),
      }))()`, "standard layout from the stored settings");
  await waitForEvaluation(`(() => {
        const workspace = document.querySelector(".workspace");
        const sourcePanel = document.querySelector(".source-panel, .source-pane-slot");
        const sourceToggle = document.querySelector("[data-source-toggle]");
        return {
          ready: Boolean(document.querySelector("math-field")) &&
            Boolean(sourceToggle || sourcePanel),
          workspaceClass: workspace?.className ?? "",
          sourcePanelVisible: Boolean(sourcePanel),
          sourceToggleVisible: Boolean(sourceToggle),
          sourceOpenStored:
            localStorage.getItem("visualtex-desktop-editor-source-open") ?? "",
          editorState:
            JSON.parse(localStorage.getItem("visualtex-editor") || "{}").state
              ?.editorLayout ?? "",
        };
      })()`, "standard editor source controls");
  await evaluate(`(() => {
        const laterButton = [...document.querySelectorAll('.office-first-run-backdrop button')]
          .find((button) => /Later|稍后处理/.test(button.textContent || ''));
        if (laterButton instanceof HTMLElement) laterButton.click();
        return true;
      })()`);
  await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        if (!field?.isConnected) return { ready: false };
        field.setValue("x", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.position = field.lastOffset;
        field.dispatchEvent(new InputEvent("input", {
          bubbles: true,
          composed: true,
          inputType: "insertText",
        }));
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        return { ready: field.value === "x" && field.hasFocus() };
      })()`, "source-preview display fixture");
  await key("Enter", "Enter", 13, 2);
  await waitForEvaluation(`(() => {
        const rows = [...document.querySelectorAll(".formula-line")];
        const fields = [...document.querySelectorAll("math-field")];
        const secondToggle = rows[1]?.querySelector("[data-formula-line-mode-toggle]");
        return {
          ready:
            rows.length === 2 &&
            fields.length === 2 &&
            secondToggle?.getAttribute("data-formula-line-mode") === "inline",
          rowCount: rows.length,
          fieldCount: fields.length,
          secondMode: secondToggle?.getAttribute("data-formula-line-mode") ?? "",
        };
      })()`, "source-preview inline fixture row");
  await waitForEvaluation(`(() => {
        const fields = [...document.querySelectorAll("math-field")];
        const field = fields[1];
        if (!field?.isConnected) return { ready: false };
        field.setValue("y", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.position = field.lastOffset;
        field.dispatchEvent(new InputEvent("input", {
          bubbles: true,
          composed: true,
          inputType: "insertText",
        }));
        const rows = [...document.querySelectorAll(".formula-line")];
        const modes = rows.map((row) =>
          row.querySelector("[data-formula-line-mode-toggle]")?.getAttribute("data-formula-line-mode") ?? "",
        );
        return {
          ready: fields[0]?.value === "x" && field.value === "y" && modes.join(",") === "display,inline",
          values: fields.map((item) => item.value),
          modes,
        };
      })()`, "source-preview mixed line-mode fixture");
  const sourceEditorAlreadyVisible = await evaluate(`Boolean(document.querySelector(".source-panel .cm-content"))`);
  if (!sourceEditorAlreadyVisible) {
    await evaluate(`(() => {
          const standardToggle = document.querySelector("[data-source-toggle]");
          const classicToggle = document.querySelector('[data-classic-bottom-view="source"]');
          (standardToggle || classicToggle)?.click();
        })()`);
  }
  await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector(".source-panel .cm-content")),
      }))()`, "source editor is visible");
  await sleep(450);
  await evaluate(`(() => {
        const laterButton = [...document.querySelectorAll('.office-first-run-backdrop button')]
          .find((button) => /Later|稍后处理/.test(button.textContent || ''));
        if (laterButton instanceof HTMLElement) laterButton.click();
        return true;
      })()`);
  await sleep(80);

  const markerStateBeforeSourceFocus = await evaluate(`(() =>
        [...document.querySelectorAll(".formula-line")].map((line) => {
          const toggle = line.querySelector("[data-formula-line-mode-toggle]");
          return {
            lineId: line.getAttribute("data-line-id") ?? "",
            mode: toggle?.getAttribute("data-formula-line-mode") ?? "",
            text: toggle?.textContent?.trim() ?? "",
          };
        })
      )()`);
  assert.ok(
    markerStateBeforeSourceFocus.length > 0,
    "formula line mode markers exist before source focus",
  );
  assert.ok(
    markerStateBeforeSourceFocus.every(
      (item) =>
        (item.mode === "inline" && item.text === "$") ||
        (item.mode === "display" && item.text === "$$"),
    ),
    `line mode marker text must match the stored mode before source focus: ${JSON.stringify(markerStateBeforeSourceFocus)}`,
  );

  const sourceFocusPoint = await evaluate(`(() => {
        const lines = [...document.querySelectorAll(".source-panel .cm-line")];
        const target = lines[1] ?? lines[0];
        const rect = target?.getBoundingClientRect();
        return rect ? { x: rect.left + 8, y: rect.top + rect.height / 2 } : null;
      })()`);
  assert.ok(sourceFocusPoint, "source editor has a click target");
  await client.send("Input.dispatchMouseEvent", {
    type: "mousePressed",
    x: sourceFocusPoint.x,
    y: sourceFocusPoint.y,
    button: "left",
    buttons: 1,
    clickCount: 1,
  });
  await client.send("Input.dispatchMouseEvent", {
    type: "mouseReleased",
    x: sourceFocusPoint.x,
    y: sourceFocusPoint.y,
    button: "left",
    buttons: 0,
    clickCount: 1,
  });
  const focusedState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const surface = document.querySelector(".multi-line-editor");
        const workspace = document.querySelector(".workspace");
        const line = document.querySelector(".formula-line");
        const native = document.getElementById("mathlive-suggestion-popover");
        const stable = document.getElementById("mathlive-suggestion-popover");
        const isPainted = (node) => {
          if (!node) return false;
          const style = getComputedStyle(node);
          const background = style.backgroundColor;
          return style.display !== "none" &&
            style.visibility !== "hidden" &&
            Number.parseFloat(style.opacity || "1") > 0 &&
            background !== "transparent" &&
            background !== "rgba(0, 0, 0, 0)";
        };
        const selectionNodes = [...(field?.shadowRoot?.querySelectorAll(
          ".ML__selection, .ML__selected, .ML__contains-highlight, .ML__highlight, .ML__placeholder-selected",
        ) ?? [])];
        const caretNodes = [...(field?.shadowRoot?.querySelectorAll(
          ".ML__caret, .ML__text-caret, .ML__latex-caret, .visualtex-structural-placeholder-caret",
        ) ?? [])];
        const lineStyle = line ? getComputedStyle(line) : null;
        const modeToggle = line?.querySelector("[data-formula-line-mode-toggle]");
        const modeToggleStyle = modeToggle ? getComputedStyle(modeToggle) : null;
        return {
          ready:
            document.documentElement.classList.contains("visualtex-source-editor-focused") &&
            workspace?.classList.contains("is-source-editor-focused") &&
            surface?.classList.contains("is-source-preview-only") &&
            surface?.classList.contains("has-mixed-line-modes") &&
            field?.classList.contains("visualtex-source-preview-only") &&
            field?.readOnly === true &&
            field?.selectionIsCollapsed === true &&
            Boolean(modeToggle) &&
            modeToggle?.disabled === true &&
            ["inline", "display"].includes(modeToggle?.getAttribute("data-formula-line-mode") || "") &&
            modeToggleStyle?.display !== "none" &&
            modeToggleStyle?.visibility !== "hidden" &&
            Number.parseFloat(modeToggleStyle?.opacity || "0") > 0 &&
            !document.querySelector(".suggestion-popup") &&
            !isPainted(native) &&
            !isPainted(stable) &&
            selectionNodes.every((node) => !isPainted(node)) &&
            caretNodes.every((node) => {
              const style = getComputedStyle(node);
              return style.display === "none" || Number.parseFloat(style.opacity || "1") === 0;
            }) &&
            (!lineStyle || lineStyle.backgroundColor === "rgba(0, 0, 0, 0)"),
          rootFocused: document.documentElement.classList.contains("visualtex-source-editor-focused"),
          workspaceFocused: workspace?.classList.contains("is-source-editor-focused") ?? false,
          surfacePreviewOnly: surface?.classList.contains("is-source-preview-only") ?? false,
          fieldPreviewOnly: field?.classList.contains("visualtex-source-preview-only") ?? false,
          readOnly: field?.readOnly ?? false,
          selectionCollapsed: field?.selectionIsCollapsed ?? false,
          customCandidateVisible: Boolean(document.querySelector(".suggestion-popup")),
          nativePainted: isPainted(native),
          stablePainted: isPainted(stable),
          paintedSelectionCount: selectionNodes.filter(isPainted).length,
          visibleCaretCount: caretNodes.filter((node) => {
            const style = getComputedStyle(node);
            return style.display !== "none" && Number.parseFloat(style.opacity || "1") > 0;
          }).length,
          lineBackground: lineStyle?.backgroundColor ?? "",
          modeMarkerPresent: Boolean(modeToggle),
          modeMarkerMode: modeToggle?.getAttribute("data-formula-line-mode") ?? "",
          modeMarkerDisabled: modeToggle?.disabled ?? false,
          modeMarkerOpacity: modeToggleStyle?.opacity ?? "",
          modeMarkers: [...document.querySelectorAll(".formula-line")].map((row) => {
            const toggle = row.querySelector("[data-formula-line-mode-toggle]");
            return {
              lineId: row.getAttribute("data-line-id") ?? "",
              mode: toggle?.getAttribute("data-formula-line-mode") ?? "",
              text: toggle?.textContent?.trim() ?? "",
              disabled: toggle?.disabled ?? false,
            };
          }),
        };
      })()`, "source focus turns the visual editor into clean live preview");
  assert.deepEqual(
    focusedState.modeMarkers.map(({ lineId, mode, text }) => ({ lineId, mode, text })),
    markerStateBeforeSourceFocus,
    `source focus must preserve every formula line's $/$$ marker and mode: ${JSON.stringify(focusedState.modeMarkers)}`,
  );
  assert.ok(
    focusedState.modeMarkers.every(
      (item) =>
        item.disabled &&
        ((item.mode === "inline" && item.text === "$") ||
          (item.mode === "display" && item.text === "$$")),
    ),
    `source preview mode markers must stay visible, correct, and read-only: ${JSON.stringify(focusedState.modeMarkers)}`,
  );

  const sourceInsertionPoint = await evaluate(`(() => {
        const lines = [...document.querySelectorAll(".source-panel .cm-line")];
        const targetIndex = lines.findIndex((line, index) =>
          index > 0 &&
          index < lines.length - 1 &&
          lines[index - 1]?.textContent?.trim() === "$$" &&
          lines[index + 1]?.textContent?.trim() === "$$",
        );
        const target = targetIndex >= 0 ? lines[targetIndex] : lines[1] ?? lines[0];
        const rect = target?.getBoundingClientRect();
        return rect
          ? { x: rect.left + 8, y: rect.top + rect.height / 2 }
          : null;
      })()`);
  assert.ok(sourceInsertionPoint, "source formula line has a click target");
  await client.send("Input.dispatchMouseEvent", {
    type: "mousePressed",
    x: sourceInsertionPoint.x,
    y: sourceInsertionPoint.y,
    button: "left",
    buttons: 1,
    clickCount: 1,
  });
  await client.send("Input.dispatchMouseEvent", {
    type: "mouseReleased",
    x: sourceInsertionPoint.x,
    y: sourceInsertionPoint.y,
    button: "left",
    buttons: 0,
    clickCount: 1,
  });
  await client.send("Input.insertText", {
    text: "\\theta+\\placeholder{}",
  });
  const liveRenderState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const source = document.querySelector(".source-panel .cm-content")?.innerText ?? "";
        return {
          ready:
            source.includes("\\\\theta+\\\\placeholder{}") &&
            field?.value.includes("\\\\theta") &&
            field?.value.includes("\\\\placeholder{}") &&
            field.readOnly &&
            field.selectionIsCollapsed &&
            document.documentElement.classList.contains("visualtex-source-editor-focused") &&
            !document.querySelector(".suggestion-popup") &&
            !document.querySelector("[data-source-reset]"),
          source,
          value: field?.value ?? "",
          readOnly: field?.readOnly ?? false,
          selectionCollapsed: field?.selectionIsCollapsed ?? false,
          popupVisible: Boolean(document.querySelector(".suggestion-popup")),
          resetVisible: Boolean(document.querySelector("[data-source-reset]")),
          sourceErrorVisible: Boolean(document.querySelector(".source-error-chip")),
          rootFocused: document.documentElement.classList.contains("visualtex-source-editor-focused"),
          cmFocused: document.querySelector(".source-panel .cm-editor")?.classList.contains("cm-focused") ?? false,
          activeTag: document.activeElement?.tagName ?? "",
          activeClass: document.activeElement?.className ?? "",
        };
      })()`, "previewable placeholder source renders live without a reset-only state");

  const visualClickPoint = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const content = field?.shadowRoot?.querySelector('[part="content"]');
        const rect = content?.getBoundingClientRect() ?? field?.getBoundingClientRect();
        return rect
          ? { x: rect.left + Math.max(8, rect.width * 0.65), y: rect.top + rect.height / 2 }
          : null;
      })()`);
  assert.ok(visualClickPoint, "rendered formula has a click target");
  await client.send("Input.dispatchMouseEvent", {
    type: "mousePressed",
    x: visualClickPoint.x,
    y: visualClickPoint.y,
    button: "left",
    buttons: 1,
    clickCount: 1,
  });
  await client.send("Input.dispatchMouseEvent", {
    type: "mouseReleased",
    x: visualClickPoint.x,
    y: visualClickPoint.y,
    button: "left",
    buttons: 0,
    clickCount: 1,
  });
  const restoredInteractionState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        return {
          ready:
            !document.documentElement.classList.contains("visualtex-source-editor-focused") &&
            !document.querySelector(".workspace")?.classList.contains("is-source-editor-focused") &&
            !document.querySelector(".multi-line-editor")?.classList.contains("is-source-preview-only") &&
            !field?.classList.contains("visualtex-source-preview-only") &&
            field?.readOnly === false &&
            field?.hasFocus() &&
            field?.value.includes("\\\\placeholder{}") &&
            !document.querySelector("[data-source-reset]") &&
            !document.querySelector(".source-error-chip"),
          value: field?.value ?? "",
          readOnly: field?.readOnly ?? true,
          fieldFocused: field?.hasFocus() ?? false,
          fieldPreviewOnly: field?.classList.contains("visualtex-source-preview-only") ?? true,
          resetVisible: Boolean(document.querySelector("[data-source-reset]")),
          sourceErrorVisible: Boolean(document.querySelector(".source-error-chip")),
          modeMarkers: [...document.querySelectorAll(".formula-line")].map((row) => {
            const toggle = row.querySelector("[data-formula-line-mode-toggle]");
            return {
              lineId: row.getAttribute("data-line-id") ?? "",
              mode: toggle?.getAttribute("data-formula-line-mode") ?? "",
              text: toggle?.textContent?.trim() ?? "",
            };
          }),
        };
      })()`, "one visual click accepts the previewable source and enters formula editing");
  assert.deepEqual(
    restoredInteractionState.modeMarkers,
    markerStateBeforeSourceFocus,
    `leaving source focus must not mutate line modes or their $/$$ markers: ${JSON.stringify(restoredInteractionState.modeMarkers)}`,
  );

  await clearField();
  await typeText("\\th");
  const restoredCandidateState = await waitForEvaluation(`(() => {
        const native = document.getElementById("mathlive-suggestion-popover");
        const stable = document.getElementById("mathlive-suggestion-popover");
        const visible = (node) => {
          if (!node) return false;
          const style = getComputedStyle(node);
          return style.display !== "none" &&
            style.visibility !== "hidden" &&
            Number.parseFloat(style.opacity || "1") > 0;
        };
        const customVisible = Boolean(document.querySelector(".suggestion-popup"));
        const nativeVisible = visible(native);
        const stableVisible = visible(stable);
        return {
          ready: customVisible || nativeVisible || stableVisible,
          customVisible,
          nativeVisible,
          stableVisible,
        };
      })()`, "visual command candidates return after source focus leaves");

  console.log(JSON.stringify({
    markerStateBeforeSourceFocus,
    focusedState,
    liveRenderState,
    restoredInteractionState,
    restoredCandidateState,
  }, null, 2));
  console.log("Targeted source-only live preview regression passed");
  return;
}
