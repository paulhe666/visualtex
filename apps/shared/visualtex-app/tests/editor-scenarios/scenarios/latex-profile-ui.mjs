// the universal LaTeX format profile menu

export const names = ["latex-profile-ui"];

export async function run({ clearField, clickSelectorWithPointer, client, evaluate, key, typeText, waitForEvaluation }) {
  await waitForEvaluation(`(() => {
        const button = document.querySelector(".app-header .code-format-primary");
        return { ready: button instanceof HTMLElement };
      })()`, "LaTeX profile trigger");

  await clickSelectorWithPointer(".app-header .code-format-primary");
  await waitForEvaluation(`(() => ({
        ready: [
          '[data-latex-text-policy="text-command"]',
          '[data-latex-text-policy="outside-math"]',
          '[data-latex-inline-wrapper="dollar"]',
          '[data-latex-inline-wrapper="paren"]',
          '[data-latex-display-wrapper="double-dollar"]',
          '[data-latex-display-wrapper="bracket"]',
          '[data-latex-display-wrapper="equation"]',
          '[data-latex-numbered]',
          '[data-latex-multiline="gather"]',
          '[data-latex-multiline="align"]',
        ].every((selector) => document.querySelector(selector) instanceof HTMLElement),
      }))()`, "universal LaTeX profile controls");

  await clickSelectorWithPointer('[data-latex-text-policy="outside-math"]');
  await clickSelectorWithPointer('[data-latex-inline-wrapper="paren"]');
  await clickSelectorWithPointer('[data-latex-display-wrapper="equation"]');
  await clickSelectorWithPointer('[data-latex-numbered]');
  await clickSelectorWithPointer('[data-latex-multiline="align"]');

  const profile = await waitForEvaluation(`(() => {
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        const value = persisted?.state?.latexFormatProfile ?? null;
        return {
          ready:
            value?.inlineTextPolicy === "outside-math" &&
            value?.inlineWrapper === "paren" &&
            value?.displayWrapper === "equation" &&
            value?.numbered === true &&
            value?.multilineEnvironment === "align",
          value,
          summary: document.querySelector("[data-latex-profile-summary]")?.textContent ?? "",
        };
      })()`, "persisted universal LaTeX profile");

  await clickSelectorWithPointer(".app-header .code-format-primary");
  await clearField();
  await typeText("ab");

  const initialRow = await waitForEvaluation(`(() => {
        const row = document.querySelector(".formula-line");
        const toggle = row?.querySelector("[data-formula-line-mode-toggle]");
        const host = row?.querySelector(".mathfield-host");
        const select = row?.querySelector("[data-formula-display-style]");
        const toggleRect = toggle?.getBoundingClientRect();
        const hostRect = host?.getBoundingClientRect();
        return {
          ready:
            toggle instanceof HTMLButtonElement &&
            !select &&
            Boolean(toggleRect && hostRect && toggleRect.right <= hostRect.left + 0.5),
          mode: toggle?.getAttribute("data-formula-line-mode") ?? null,
          text: toggle?.textContent?.trim() ?? "",
          overlap: toggleRect && hostRect
            ? Math.max(0, Math.min(toggleRect.right, hostRect.right) - Math.max(toggleRect.left, hostRect.left))
            : null,
        };
      })()`, "non-overlapping single-symbol per-row mode control");

  if (initialRow.mode !== "inline") {
    await clickSelectorWithPointer("[data-formula-line-mode-toggle]");
  }
  const inlineMode = await waitForEvaluation(`(() => {
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        const line = persisted?.state?.lines?.[0];
        const toggle = document.querySelector("[data-formula-line-mode-toggle]");
        return {
          ready:
            line?.mode === "inline" &&
            toggle?.getAttribute("data-formula-line-mode") === "inline" &&
            toggle?.textContent?.trim() === "$" &&
            !document.querySelector("[data-formula-display-style]"),
          mode: line?.mode ?? null,
          toggleMode: toggle?.getAttribute("data-formula-line-mode") ?? null,
          text: toggle?.textContent?.trim() ?? "",
        };
      })()`, "single per-row control toggles to inline");

  await evaluate(`(() => {
        const field = document.querySelector("math-field");
        if (!field) return false;
        field.focus();
        field.position = 1;
        field.selection = { ranges: [[1, 1]], direction: "none" };
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        return true;
      })()`);
  await key("Enter", "Enter", 13, 8);

  const multiline = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        const line = persisted?.state?.lines?.[0];
        const toggle = document.querySelector("[data-formula-line-mode-toggle]");
        return {
          ready:
            document.querySelectorAll(".formula-line").length === 1 &&
            field?.value?.includes("\\\\begin{aligned}") &&
            toggle instanceof HTMLButtonElement &&
            toggle.disabled === true &&
            toggle.getAttribute("data-formula-line-mode") === "display" &&
            toggle.textContent?.trim() === "$$",
          value: field?.value ?? "",
          mode: line?.mode ?? null,
          toggleMode: toggle?.getAttribute("data-formula-line-mode") ?? null,
          disabled: toggle instanceof HTMLButtonElement ? toggle.disabled : null,
        };
      })()`, "Shift+Enter follows persistent align profile");

  await clickSelectorWithPointer('button[aria-label="撤销"]');
  const multilineUndo = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        const line = persisted?.state?.lines?.[0];
        const toggle = document.querySelector("[data-formula-line-mode-toggle]");
        return {
          ready:
            field?.value === "ab" &&
            line?.mode === "inline" &&
            toggle?.getAttribute("data-formula-line-mode") === "inline" &&
            toggle?.disabled === false,
          value: field?.value ?? "",
          mode: line?.mode ?? null,
          toggleMode: toggle?.getAttribute("data-formula-line-mode") ?? null,
        };
      })()`, "multiline undo restores inline row semantics");

  await clickSelectorWithPointer('button[aria-label="重做"]');
  const multilineRedo = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        const line = persisted?.state?.lines?.[0];
        const toggle = document.querySelector("[data-formula-line-mode-toggle]");
        return {
          ready:
            field?.value?.includes("\\\\begin{aligned}") &&
            line?.mode === "display" &&
            toggle?.getAttribute("data-formula-line-mode") === "display" &&
            toggle?.disabled === true,
          value: field?.value ?? "",
          storedLatex: line?.latex ?? null,
          mode: line?.mode ?? null,
          toggleMode: toggle?.getAttribute("data-formula-line-mode") ?? null,
          disabled: toggle instanceof HTMLButtonElement ? toggle.disabled : null,
        };
      })()`, "multiline redo restores display row semantics");

  const multilineStartGeometry = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const model = field?._mathfield?.model;
        const environment = model?.parentEnvironment;
        if (!field || !model || environment?.environmentName !== "aligned") {
          return { ready: false };
        }
        const first = environment.getCell?.(0, 1)?.find((atom) => atom.type !== "first");
        const second = environment.getCell?.(1, 1)?.find((atom) => atom.type !== "first");
        const firstBounds = first ? field.getElementInfo(model.offsetOf(first))?.bounds : null;
        const secondBounds = second ? field.getElementInfo(model.offsetOf(second))?.bounds : null;
        const leftDelta =
          firstBounds && secondBounds
            ? Math.abs(firstBounds.left - secondBounds.left)
            : null;
        return {
          ready: Boolean(firstBounds && secondBounds && leftDelta !== null && leftDelta <= 2),
          leftDelta,
          first: firstBounds
            ? { left: firstBounds.left, right: firstBounds.right, width: firstBounds.width }
            : null,
          second: secondBounds
            ? { left: secondBounds.left, right: secondBounds.right, width: secondBounds.width }
            : null,
          value: field.value,
        };
      })()`, "Shift+Enter aligned rows begin at the same visual x anchor");

  await key("&", "Digit7", 55, 8);
  const alignmentPoint = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const model = field?._mathfield?.model;
        const environment = model?.parentEnvironment;
        const cursor = model?.at(model.position);
        const branch = cursor?.parentBranch;
        const leftCell = environment?.getCell?.(1, 0) ?? [];
        const rightCell = environment?.getCell?.(1, 1) ?? [];
        const values = (cell) =>
          cell
            .filter((atom) => atom.type !== "first")
            .map((atom) => atom.value ?? atom.command ?? "");
        return {
          ready:
            environment?.environmentName === "aligned" &&
            Array.isArray(branch) &&
            branch[0] === 1 &&
            branch[1] === 1 &&
            values(leftCell).length === 0 &&
            values(rightCell).includes("b"),
          value: field?.value ?? "",
          environment: environment?.environmentName ?? null,
          branch: Array.isArray(branch) ? branch : null,
          leftCell: values(leftCell),
          rightCell: values(rightCell),
        };
      })()`, "ampersand becomes a real aligned column boundary");

  await client.send("Page.reload", { ignoreCache: true });
  const reopened = await waitForEvaluation(`(() => {
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        const profile = persisted?.state?.latexFormatProfile;
        const line = persisted?.state?.lines?.[0];
        return {
          ready:
            profile?.inlineWrapper === "paren" &&
            profile?.displayWrapper === "equation" &&
            profile?.numbered === true &&
            profile?.multilineEnvironment === "align" &&
            /\\\\begin\\{aligned\\}/.test(line?.latex ?? ""),
          profile,
          line,
        };
      })()`, "LaTeX profile and aligned row after reload");

  console.log(JSON.stringify({
    profile,
    initialRow,
    inlineMode,
    multiline,
    multilineUndo,
    multilineRedo,
    multilineStartGeometry,
    alignmentPoint,
    reopened,
  }, null, 2));
  console.log("Targeted universal LaTeX profile UI regression passed");
  return;
}
