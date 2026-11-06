import assert from "node:assert/strict";
import { latexToMathMl, latexToSvg } from "../src/export/runtime.ts";
import {
  ESINT_INTEGRAL_GLYPHS,
  ESINT_INTEGRAL_GLYPHS_BY_COMMAND,
} from "../src/math/esintGlyphs.ts";

const commands = Array.from(
  new Set(
    ESINT_INTEGRAL_GLYPHS.flatMap((glyph) => [
      glyph.command,
      ...glyph.aliases,
    ]),
  ),
);

// Images are drawn by the editor's layout engine (LaTeX with amsmath + esint):
// each glyph is a <use> of an outline whose id names its TeX font. esint
// operators come from esint10; amsmath's multiple integrals are built from cmex10.
const AMSMATH_MULTIPLE_INTEGRALS = new Set(["iiiint", "idotsint"]);
// \intclockwise (the MathLive/UnicodeMath spelling of esint's \ointclockwise)
// is not an esint command: the engine draws U+2231 from the STIX Two Math outlines.
const STIX_OUTLINE_INTEGRALS = new Set(["intclockwise"]);

function glyphFonts(svg: string) {
  return new Set((svg.match(/id="vtx-([a-z]+\d+)-\d+"/g) ?? []).map((m) => m.slice(8, m.lastIndexOf("-"))));
}

for (const command of commands) {
  const definition = ESINT_INTEGRAL_GLYPHS_BY_COMMAND[command];
  assert.ok(definition, `${command} has an esint glyph definition`);
  const latex = `\\${command}_{a}^{b} f(x)`;
  const expected = AMSMATH_MULTIPLE_INTEGRALS.has(command) ? "cmex10" : "esint10";
  for (const displayMode of [true, false]) {
    const svg = (await latexToSvg(latex, {
      displayMode,
      fontSizePt: 12,
      paddingPx: 0,
      background: "transparent",
    })).svg;
    assert.ok(!svg.includes(`\\${command}`), `${command} SVG contains no unresolved command text`);
    assert.ok(!svg.includes("<text"), `${command} SVG draws no system-font text`);
    if (STIX_OUTLINE_INTEGRALS.has(command)) {
      assert.match(svg, /<path transform="[^"]*" d="/, `${command} SVG draws the STIX operator outline`);
      continue;
    }
    assert.ok(
      glyphFonts(svg).has(expected),
      `${command} ${displayMode ? "display" : "inline"} SVG uses the ${expected} operator glyph`,
    );
  }

  const mathMl = latexToMathMl(latex, true);
  assert.ok(mathMl.startsWith("<math"), `${command} produces MathML`);
  assert.ok(!mathMl.includes(`\\${command}`), `${command} MathML is resolved`);
  assert.ok(!/<merror\b/i.test(mathMl), `${command} MathML has no error node`);
}

const fintSvg = (await latexToSvg(String.raw`\fint_{a}^{b} f(x)`, {
  displayMode: true,
  fontSizePt: 12,
  paddingPx: 0,
  background: "transparent",
})).svg;
assert.ok(glyphFonts(fintSvg).has("esint10"), "fint export uses the esint10 glyph (not the old STIX U+2A0F outline)");

console.log(
  `VisualTeX esint export regression: PASS (${commands.length} commands and aliases)`,
);
