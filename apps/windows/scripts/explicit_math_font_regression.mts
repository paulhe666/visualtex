import assert from "node:assert/strict";
import { DOMParser } from "@xmldom/xmldom";
import { latexToMathMl, latexToSvg } from "../src/export/runtime.ts";
const parse = (xml: string) => new DOMParser().parseFromString(xml, "application/xml");
function tokens(xml: string) {
  return Array.from(parse(xml).getElementsByTagName("*"))
    .filter((n) => n.tagName === "mo" || n.tagName === "mi")
    .map(n => ({ kind:n.tagName, text:n.textContent, variant:n.getAttribute("mathvariant") || (n.tagName === "mo" ? "normal" : "italic") }));
}
let passed = 0;
for (const [command, variant] of [["mathrm","normal"],["mathit","italic"],["mathbf","bold"],["mathbfit","bold-italic"]]) {
  for (const [named, unicode] of [["\\parallel","∥"],["\\nparallel","∦"],["\\angle","∠"],["\\sim","∼"],["\\to","→"],["+","+"]]) {
    for (const input of [named, unicode]) {
      const latex = "\\" + command + "{" + input + "}";
      const actual=tokens(await latexToMathMl(latex,false));
      assert.equal(actual.length,1,latex); assert.equal(actual[0].text,unicode,latex);
      assert.equal(actual[0].variant,variant,latex);
      const svg=await latexToSvg(latex);
      assert.ok(svg.width>0&&svg.height>0&&Number.isFinite(svg.baseline),latex);
      if (unicode !== "+" && (variant==="italic"||variant==="bold-italic"))
        assert.ok(svg.svg.includes('data-visualtex-synthetic-oblique="true"'),latex);
      if (variant==="normal"||variant==="bold")
        assert.ok(!svg.svg.includes("data-visualtex-synthetic-oblique"),latex);
      passed++;
    }
  }
}
for (const [latex, variants] of [
  [String.raw`\mathit{\parallel\mathrm{\parallel}\mathbf{\parallel}\mathnormal{\parallel}}`,["italic","normal","bold","normal"]],
  [String.raw`\mathit{\parallel{\rm\parallel}\parallel}`,["italic","normal","italic"]],
  [String.raw`\mathit{\parallel{\bf\parallel}\parallel}`,["italic","bold","italic"]],
  [String.raw`\boldsymbol{\mathit{\parallel}}`,["bold-italic"]],
  [String.raw`\mathit{\boldsymbol{\parallel}}`,["bold-italic"]],
  [String.raw`\parallel ∥ +`,["normal","normal"]],
] as const) {
  assert.deepEqual(tokens(await latexToMathMl(latex,false)).map(t=>t.variant),variants,latex);passed++;
}
const roman=await latexToSvg(String.raw`\mathrm{\parallel}`,{fontSizePt:12,paddingPx:0,displayMode:false});
const italic=await latexToSvg(String.raw`\mathit{\parallel}`,{fontSizePt:12,paddingPx:0,displayMode:false});
assert.ok(italic.width>roman.width);
assert.equal(italic.height,roman.height); assert.equal(italic.baseline,roman.baseline);passed++;
for (const latex of [String.raw`a\mathit{\parallel}b`,String.raw`\mathit{\parallel}_{i}^{n}`,String.raw`\mathit{\left\Vert\frac{x}{y}\right\Vert}`]) {
  const rendered=await latexToSvg(latex);
  assert.ok(Number.isFinite(rendered.width)&&rendered.width>0);
  assert.ok(!/NaN|Infinity/.test(rendered.svg));passed++;
}
console.log(JSON.stringify({passed,scope:"explicit operator styles, real SVG shear, nested font scope, layout baseline"}));
