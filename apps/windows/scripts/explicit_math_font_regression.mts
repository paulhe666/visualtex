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
// Identifier scope is independent of Unicode symbol categories. The named and
// literal forms must carry the same explicit basic style without erasing an
// intrinsic alphabet or a nested operator-name scope.
for (const [command, variant] of [["mathrm","normal"],["mathit","italic"],["mathbf","bold"],["mathbfit","bold-italic"]]) {
  for (const named of ["hslash","hbar","alpha","Gamma","ell"]) {
    const latex = "\\" + command + "{\\" + named + "}";
    const xml = await latexToMathMl(latex,false);
    const actual = tokens(xml);
    assert.equal(actual.length,1,latex);
    assert.equal(actual[0].kind,"mi",latex);
    assert.equal(actual[0].variant,variant,latex);
    if (named === "hbar") assert.ok(xml.includes('data-mjx-alternate="1"'),latex);
    if (named === "hslash") assert.ok(!xml.includes('data-mjx-alternate="1"'),latex);
    passed++;
  }
}
for (const [latex, expected] of [
  [String.raw`\mathrm{\Bbbk}`,"double-struck"],
  [String.raw`\mathit{\mathbb{R}}`,"double-struck"],
  [String.raw`\mathbf{\mathcal{R}}`,"script"],
  [String.raw`\mathit{\operatorname{x}}`,"normal"],
] as const) {
  assert.equal(tokens(await latexToMathMl(latex,false))[0].variant,expected,latex);passed++;
}
assert.ok((await latexToMathMl(String.raw`\mathit{x}`,false)).includes('data-mjx-variant="-tex-mathit"'));passed++;
const nestedIdentifier=tokens(await latexToMathMl(String.raw`\mathbfit{\hbar\mathrm{\hslash}}`,false));
assert.deepEqual(nestedIdentifier.map(x=>x.variant),["bold-italic","normal"]);passed++;
console.log(JSON.stringify({passed,scope:"explicit operator and identifier styles, intrinsic alphabets, real SVG operator shear, nested font scope, layout baseline"}));
