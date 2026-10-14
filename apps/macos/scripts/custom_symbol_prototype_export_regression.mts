import assert from "node:assert/strict";
import {
  latexToMathMl,
  latexToSvg,
  svgToBase64,
} from "../src/export/runtime.ts";
import {
  CUSTOM_SYMBOL_PROTOTYPE_DEFINITION,
} from "../src/math/customSymbolRegistry.ts";
const CUSTOM_SYMBOL_PROTOTYPE_LATEX = `\\${CUSTOM_SYMBOL_PROTOTYPE_DEFINITION.command}`;

const cases = [
  ["standalone", CUSTOM_SYMBOL_PROTOTYPE_LATEX],
  ["surrounding", `A+${CUSTOM_SYMBOL_PROTOTYPE_LATEX}+B`],
  ["subscript", `x_{${CUSTOM_SYMBOL_PROTOTYPE_LATEX}}`],
  ["scripts", `${CUSTOM_SYMBOL_PROTOTYPE_LATEX}_i^j`],
] as const;

for (const [name, latex] of cases) {
  const result = latexToSvg(latex, {
    displayMode: false,
    fontSizePt: 12,
    paddingPx: 2,
    background: "transparent",
  });
  assert.ok(result.width > 0, `${name} width`);
  assert.ok(result.height > 0, `${name} height`);
  assert.match(
    result.svg,
    /data-visualtex-custom-symbol="vtxtestsymbol"/,
    `${name} registry artwork group`,
  );
  assert.match(result.svg, /<circle\b[^>]*cx="290"[^>]*cy="365"/, `${name} circle geometry`);
  assert.match(result.svg, /<line\b[^>]*x1="65"[^>]*x2="515"/, `${name} line geometry`);
  assert.match(result.svg, /<path\b[^>]*d="M202 508Q179 508/, `${name} compiled partial path`);
  assert.doesNotMatch(
    result.svg,
    /\\vtxtestsymbol/,
    `${name} must not leak the literal command into SVG`,
  );
  assert.doesNotMatch(result.svg, /<foreignObject\b/i, `${name} foreignObject`);
  assert.equal(result.base64, svgToBase64(result.svg), `${name} base64 round trip`);
}

const ordinary = latexToSvg("x+y", {
  displayMode: false,
  fontSizePt: 12,
  paddingPx: 0,
  background: "transparent",
});
assert.doesNotMatch(
  ordinary.svg,
  /data-visualtex-custom-symbol=/,
  "Ordinary SVG must not enter the custom-symbol post-processor",
);

const existingIntegral = latexToSvg("\\oiint_C f", {
  displayMode: true,
  fontSizePt: 12,
  paddingPx: 0,
  background: "transparent",
});
assert.match(
  existingIntegral.svg,
  /data-visualtex-integral="oiint"/,
  "Existing contour-integral vector patch must remain intact",
);

assert.throws(
  () => latexToMathMl(CUSTOM_SYMBOL_PROTOTYPE_LATEX, false),
  /did not resolve LaTeX command \\vtxtestsymbol/,
  "Word/MathML must reject the phase-1 prototype instead of silently degrading it",
);

const normalMathMl = latexToMathMl("x+y", false);
assert.match(normalMathMl, /^<math\b/);

console.log(
  "Custom symbol prototype SVG composition, isolation, existing-integral compatibility, and MathML guard regression passed",
);
