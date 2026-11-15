import assert from "node:assert/strict";
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

await withAppBrowser(async ({ evaluate }) => {
  // The desktop preview can reload once at startup while persisted editor state is initialized. Probe only after
  // that navigation has settled so a context teardown is not mistaken for a kernel failure.
  await sleep(1800);
  const result = await evaluate(`(async () => {
          await customElements.whenDefined("math-field");
          const samples = [
            "\\\\div", "\\\\divisionsymbol", "\\\\quantity{x}",
            "\\\\pqty{x}", "\\\\derivative{f}{x}",
            "\\\\partialderivative{f}{x}",
            "\\\\pmqty{1&0\\\\\\\\0&1}",
            "\\\\qq{if}", "A\\\\qq{if}B", "\\\\vqty{x}",
            "\\\\absolutevalue{x}", "\\\\vev{A}",
            "\\\\expectationvalue{A}{p}",
            "\\\\matrixelement{p}{A}{q}",
            "\\\\outerproduct{p}{q}",
            "\\\\flatfrac{a}{b}", "\\\\qty{x}"
          ];
          const fields = samples.map((latex) => {
            const field = document.createElement("math-field");
            document.body.append(field);
            field.setValue(latex, { mode: "math", format: "latex" });
            return { latex, field };
          });
          await new Promise((resolve) => setTimeout(resolve, 100));
          return fields.map(({ latex, field }) => {
            const value = field.value;
            let editedValue = "";
            if (latex.startsWith("\\\\pmqty")) {
              field.position = 2;
              field.insert("2", { mode: "math", format: "latex" });
              editedValue = field.value;
            }
            const result = {
              latex,
              value,
              editedValue,
              error: Boolean(field.shadowRoot?.querySelector(".ML__error")),
              array: Boolean(field._mathfield?.model?.atoms?.some(
                (atom) => atom.type === "array",
              )),
              renderedText: field.shadowRoot?.querySelector(".ML__base")?.textContent ?? "",
            };
            field.remove();
            return result;
          });
        })()`);

  const bySource = new Map(result.map((item) => [item.latex, item]));
  for (const item of result) {
    // \qty is physics' \quantity unless it is siunitx's \qty{<number>}{<unit>}
    assert.equal(item.error, false, `${item.latex} failed in the built MathLive kernel`);
    assert.equal(item.value, item.latex, `${item.latex} did not preserve the source command`);
  }
  assert.match(bySource.get(String.raw`\div`)?.renderedText ?? "", /÷/, "\\div lost its division sign");
  assert.match(bySource.get(String.raw`\divisionsymbol`)?.renderedText ?? "", /÷/, "\\divisionsymbol is not division");
  for (const [source, contents] of [
    [String.raw`\vqty{x}`, /x/],
    [String.raw`\absolutevalue{x}`, /x/],
    [String.raw`\vev{A}`, /A/],
    [String.raw`\expectationvalue{A}{p}`, /A/],
    [String.raw`\matrixelement{p}{A}{q}`, /A/],
    [String.raw`\outerproduct{p}{q}`, /pq|p.*q/],
    // the layout engine draws the slash as a stroke, so the text holds the two arguments only
    [String.raw`\flatfrac{a}{b}`, /a\s*\/?\s*b/],
    [String.raw`A\qq{if}B`, /AifB/],
    [String.raw`\qty{x}`, /x/],
  ]) {
    assert.match(
      bySource.get(source)?.renderedText ?? "",
      contents,
      `${source} lost an argument in the built MathLive kernel`,
    );
  }
  assert.equal(
    bySource.get(String.raw`\pmqty{1&0\\0&1}`)?.array,
    true,
    "\\pmqty was not parsed as a real tabular matrix",
  );
  assert.match(
    (bySource.get(String.raw`\pmqty{1&0\\0&1}`)?.editedValue ?? "").replace(/\s+/g, ""),
    /^\\pmqty\{[^&]*2[^&]*&0\\\\0&1\}$/,
    "Editing a matrix cell lost its column or row separators",
  );
  console.log("Built MathLive physics kernel browser regression passed.");
}, { windowSize: null });
