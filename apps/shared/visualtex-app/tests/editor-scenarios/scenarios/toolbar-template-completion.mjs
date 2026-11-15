// toolbar templates complete in the formula
import assert from "node:assert/strict";

export const names = ["toolbar-template-completion"];

export async function run({ evaluate, sleep, waitForEvaluation }) {
  await evaluate(`(() => {
        localStorage.setItem("visualtex.onboarding.v3.completed", "true");
        localStorage.setItem(
          "visualtex.onboarding.macos.desktop.v1.2.0.completed",
          "true",
        );
        if (!document.querySelector(".formula-toolbar")) {
          document.querySelector(".sidebar-toggle")?.click();
        }
        return true;
      })()`);
  await waitForEvaluation(`(() => ({
        ready:
          Boolean(document.querySelector(".formula-toolbar")) &&
          Boolean(document.querySelector(".formula-line math-field")) &&
          Boolean(document.querySelector('[data-command-id="left-upper-script"]')) &&
          Boolean(document.querySelector('[data-command-id="int-bare"]')),
      }))()`, "completed toolbar templates");
  await sleep(180);

  const structureIds = [
    "left-upper-script",
    "upper-script",
    "left-scripts",
    "scripts",
    "left-lower-script",
    "lower-script",
    "linear-fraction",
    "skewed-fraction",
  ];
  const calculusIds = [
    "int-bare",
    "iint-bare",
    "iiint-bare",
    "oint-bare",
    "oiint-bare",
    "oiiint-bare",
    "intplain-no-d",
    "int-bounds-no-d",
    "iint-no-d",
    "iint-bounds-no-d",
    "iiint-no-d",
    "iiint-bounds-no-d",
    "oint-no-d",
    "oint-bounds-no-d",
    "oiint-no-d",
    "oiiint-no-d",
  ];
  const toolbarState = await evaluate(`(() => {
        const inspect = (id) => {
          const button = document.querySelector('[data-command-id="' + id + '"]');
          const preview = button?.querySelector('.math-preview');
          const latex = preview?.querySelector('.ML__latex');
          const bounds = latex?.getBoundingClientRect();
          return {
            id,
            present: Boolean(button),
            category: button?.closest('[data-toolbar-category-section]')
              ?.getAttribute('data-toolbar-category-section') ?? null,
            previewLatex: button?.getAttribute('data-preview-latex') ?? null,
            previewVisible: Boolean(
              bounds && bounds.width > 0 && bounds.height > 0 &&
              !preview?.querySelector('.ML__error')
            ),
          };
        };
        return {
          structure: ${JSON.stringify(structureIds)}.map(inspect),
          calculus: ${JSON.stringify(calculusIds)}.map(inspect),
        };
      })()`);
  for (const entry of toolbarState.structure) {
    assert.equal(entry.present, true, JSON.stringify(entry));
    assert.equal(entry.category, "structure", JSON.stringify(entry));
    assert.equal(entry.previewVisible, true, JSON.stringify(entry));
  }
  for (const entry of toolbarState.calculus) {
    assert.equal(entry.present, true, JSON.stringify(entry));
    assert.equal(entry.category, "calculus", JSON.stringify(entry));
    assert.equal(entry.previewVisible, true, JSON.stringify(entry));
  }

  const clickAndRead = async (commandId) => {
    const before = await evaluate(`document.querySelector('.formula-line math-field')?.value ?? ''`);
    const clicked = await evaluate(`(() => {
          const field = document.querySelector('.formula-line math-field');
          const button = document.querySelector('[data-command-id="${commandId}"]');
          if (!field || !button) return false;
          field.focus();
          field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus();
          button.click();
          return true;
        })()`);
    assert.equal(clicked, true, `${commandId}: toolbar button was not clickable`);
    await sleep(120);
    const after = await evaluate(`document.querySelector('.formula-line math-field')?.value ?? ''`);
    assert.notEqual(after, before, `${commandId}: clicking the toolbar did not insert anything`);
    return { commandId, before, after };
  };

  const insertionResults = [];
  for (const id of [
    "left-upper-script",
    "left-scripts",
    "left-lower-script",
    "linear-fraction",
    "skewed-fraction",
    "int-bare",
    "int-bounds-no-d",
    "oiint-bare",
    "oiint-no-d",
  ]) {
    insertionResults.push(await clickAndRead(id));
  }
  const noDifferentialResult = insertionResults.find(
    (entry) => entry.commandId === "int-bounds-no-d",
  );
  assert.ok(noDifferentialResult, "missing no-differential insertion result");
  assert.doesNotMatch(
    noDifferentialResult.after,
    /\\mathrm\{d\}|\\differentialD/,
    JSON.stringify(noDifferentialResult),
  );
  const skewedFractionResult = insertionResults.find(
    (entry) => entry.commandId === "skewed-fraction",
  );
  assert.match(
    skewedFractionResult?.after ?? "",
    /\\nicefrac/,
    JSON.stringify(skewedFractionResult),
  );

  console.log(
    "Toolbar template completion browser regression passed",
    JSON.stringify({ toolbarState, insertionResults }),
  );
  return;
}
