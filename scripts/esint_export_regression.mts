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

function extractPath(svg: string, canonicalCommand: string) {
  const expression = new RegExp(
    `<path\\s+data-visualtex-integral="${canonicalCommand}"\\s+d="([^"]+)"`,
  );
  return svg.match(expression)?.[1] ?? "";
}

for (const command of commands) {
  const definition = ESINT_INTEGRAL_GLYPHS_BY_COMMAND[command];
  assert.ok(definition, `${command} has an esint glyph definition`);
  const latex = `\\${command}_{a}^{b} f(x)`;

  const displaySvg = latexToSvg(latex, {
    displayMode: true,
    fontSizePt: 12,
    paddingPx: 0,
    background: "transparent",
  }).svg;
  const displayPath = extractPath(displaySvg, definition.command);
  assert.equal(
    displayPath,
    definition.large.path,
    `${command} display SVG uses the official esint10 large glyph`,
  );
  assert.ok(
    !displaySvg.includes(`\\${command}`),
    `${command} display SVG contains no unresolved command text`,
  );

  const inlineSvg = latexToSvg(latex, {
    displayMode: false,
    fontSizePt: 12,
    paddingPx: 0,
    background: "transparent",
  }).svg;
  const inlinePath = extractPath(inlineSvg, definition.command);
  assert.equal(
    inlinePath,
    definition.small.path,
    `${command} inline SVG uses the official esint10 small glyph`,
  );

  const mathMl = latexToMathMl(latex, true);
  assert.ok(mathMl.startsWith("<math"), `${command} produces MathML`);
  assert.ok(!mathMl.includes(`\\${command}`), `${command} MathML is resolved`);
  assert.ok(!/<merror\b/i.test(mathMl), `${command} MathML has no error node`);
}

const fint = ESINT_INTEGRAL_GLYPHS_BY_COMMAND.fint;
assert.ok(fint, "fint definition exists");
const fintSvg = latexToSvg(String.raw`\fint_{a}^{b} f(x)`, {
  displayMode: true,
  fontSizePt: 12,
  paddingPx: 0,
  background: "transparent",
}).svg;
assert.equal(
  extractPath(fintSvg, "fint"),
  fint.large.path,
  "fint export does not fall back to the old STIX U+2A0F outline",
);

console.log(
  `VisualTeX esint export regression: PASS (${commands.length} commands and aliases)`,
);
