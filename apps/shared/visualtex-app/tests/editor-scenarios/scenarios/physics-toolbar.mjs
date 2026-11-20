// the physics toolbar group inserts its commands
import assert from "node:assert/strict";

export const names = ["physics-toolbar"];

export async function run({ evaluate, waitForEvaluation }) {
  await evaluate(`(() => {
        if (!document.querySelector('.formula-toolbar')) {
          document.querySelector('.sidebar-toggle')?.click();
        }
      })()`);
  await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector('.formula-line math-field')) &&
          Boolean(document.querySelector('[data-command-id="physics-pqty"]')),
      }))()`, "physics toolbar and editable formula");

  const insertionResults = [];
  for (const name of [
    "pqty",
    "vqty",
    "pmqty",
    "order",
    "derivative",
    "functionalderivative",
    "vev",
    "flatfrac",
  ]) {
    const clicked = await evaluate(`(() => {
          const field = document.querySelector('.formula-line math-field');
          const button = document.querySelector('[data-command-id="physics-${name}"]');
          if (!field || !button) return false;
          field.setValue('', { mode: 'math', format: 'latex' });
          field.focus();
          field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus();
          button.click();
          return true;
        })()`);
    assert.equal(clicked, true, `physics-${name}: button not found`);
    const result = await waitForEvaluation(`(() => {
          const field = document.querySelector('.formula-line math-field');
          const value = field?.value ?? '';
          return {
            ready: value.includes(${JSON.stringify(`\\${name}`)}),
            value,
            error: Boolean(field?.shadowRoot?.querySelector('.ML__error')),
          };
        })()`, `physics-${name} insertion`);
    assert.equal(result.error, false, `physics-${name}: invalid insertion`);
    insertionResults.push({ name, ...result });
  }
  console.log("Physics toolbar insertion regression passed", JSON.stringify(insertionResults));
  return;
}
