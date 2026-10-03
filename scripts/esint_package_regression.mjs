import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import {
  ESINT_GLYPHS_JSON_SHA256,
} from "../src/math/esintGlyphs.generatedData.ts";
import {
  ESINT_GLYPH_PAYLOAD,
  ESINT_INTEGRAL_GLYPHS,
  ESINT_INTEGRAL_GLYPHS_BY_COMMAND,
  ESINT_INTEGRAL_GLYPH_UNITS_PER_EM,
} from "../src/math/esintGlyphs.ts";
import {
  RARE_INTEGRAL_GLYPHS_BY_COMMAND,
} from "../src/math/rareIntegralGlyphs.generated.ts";
import {
  EXTENDED_INTEGRAL_COMMANDS,
  EXTENDED_INTEGRAL_MATHML_MACROS,
  EXTENDED_INTEGRAL_SVG_MACROS,
} from "../src/math/extendedIntegralCompatibility.ts";

const execFileAsync = promisify(execFile);

const expectedSlots = new Map([
  ["iiiint", [0o007, 0o010]],
  ["idotsint", [0o011, 0o012]],
  ["sqint", [0o017, 0o020]],
  ["sqiint", [0o021, 0o022]],
  ["ointctrclockwise", [0o027, 0o030]],
  ["ointclockwise", [0o031, 0o032]],
  ["varointclockwise", [0o033, 0o034]],
  ["varointctrclockwise", [0o035, 0o036]],
  ["fint", [0o037, 0o040]],
  ["varoiint", [0o041, 0o042]],
  ["landupint", [0o043, 0o044]],
  ["landdownint", [0o045, 0o046]],
]);

assert.equal(ESINT_INTEGRAL_GLYPH_UNITS_PER_EM, 1000);
assert.equal(ESINT_INTEGRAL_GLYPHS.length, expectedSlots.size);
assert.equal(ESINT_GLYPH_PAYLOAD.source.family, "esint10");
assert.equal(ESINT_GLYPH_PAYLOAD.source.package, "esint-type1");
assert.equal(ESINT_GLYPH_PAYLOAD.source.license, "Public Domain");
assert.equal(
  ESINT_GLYPH_PAYLOAD.source.pfbSha256,
  "3c2c4b9f98b9b741cf7e05155372c53b063fd96596205d5adfe2295ca9c9035e",
);
assert.equal(
  ESINT_GLYPH_PAYLOAD.source.tfmSha256,
  "fc941cd26d2b483f6cc9648d03d28dcc56e1c6621b6d9b6e11435a8cf2de7666",
);
assert.match(ESINT_GLYPHS_JSON_SHA256, /^[a-f0-9]{64}$/);

for (const glyph of ESINT_INTEGRAL_GLYPHS) {
  const slots = expectedSlots.get(glyph.command);
  assert.ok(slots, `unexpected esint glyph ${glyph.command}`);
  assert.deepEqual([glyph.small.slot, glyph.large.slot], slots);
  for (const [variantName, variant] of [
    ["small", glyph.small],
    ["large", glyph.large],
  ]) {
    assert.ok(variant.path.startsWith("M"), `${glyph.command} ${variantName} path`);
    assert.ok(variant.path.length > 80, `${glyph.command} ${variantName} outline`);
    assert.ok(variant.advanceWidth > 0, `${glyph.command} ${variantName} width`);
    assert.ok(
      variant.bounds.xMax > variant.bounds.xMin &&
        variant.bounds.yMax > variant.bounds.yMin,
      `${glyph.command} ${variantName} bounds`,
    );
    // Type 1 CharString closepath keeps the pre-close current point. If the
    // generator incorrectly resets that point to the subpath origin, later
    // relative moveto/hmoveto commands shift secondary contours far outside
    // the font's own advance + italic-correction metrics (sqint/sqiint were
    // the most visible examples in MathLive's native suggestion preview).
    assert.ok(
      variant.bounds.xMax <=
        variant.advanceWidth + variant.italicCorrection + 25,
      `${glyph.command} ${variantName} right outline stays within font metrics`,
    );
    assert.ok(
      variant.bounds.xMin >= -25,
      `${glyph.command} ${variantName} left outline stays within font metrics`,
    );
  }
  assert.ok(
    glyph.large.depth + glyph.large.height >
      glyph.small.depth + glyph.small.height,
    `${glyph.command} display glyph is larger`,
  );
  assert.ok(EXTENDED_INTEGRAL_COMMANDS.includes(glyph.command));
  assert.ok(EXTENDED_INTEGRAL_MATHML_MACROS[glyph.command]);
  assert.match(
    EXTENDED_INTEGRAL_SVG_MACROS[glyph.command],
    new RegExp(`visualtex-integral-export-${glyph.command}`),
  );
}

assert.equal(
  ESINT_INTEGRAL_GLYPHS_BY_COMMAND.dotsint,
  ESINT_INTEGRAL_GLYPHS_BY_COMMAND.idotsint,
);
assert.equal(
  ESINT_INTEGRAL_GLYPHS_BY_COMMAND.intclockwise,
  ESINT_INTEGRAL_GLYPHS_BY_COMMAND.ointclockwise,
);
assert.notEqual(
  ESINT_INTEGRAL_GLYPHS_BY_COMMAND.fint.small.path,
  RARE_INTEGRAL_GLYPHS_BY_COMMAND.fint.small.path,
  "fint must use the official esint10 outline rather than the STIX Unicode approximation",
);
assert.notEqual(
  ESINT_INTEGRAL_GLYPHS_BY_COMMAND.fint.large.path,
  RARE_INTEGRAL_GLYPHS_BY_COMMAND.fint.large.path,
  "display-style fint must use the official esint10 outline",
);

// Regenerate from the installed official TeX Live files and compare the parsed
// payload, so slot mistakes or hand-edited checked-in data fail deterministically.
const { stdout } = await execFileAsync(
  "python3",
  ["scripts/generate_esint_glyph_registry.py", "--json"],
  { cwd: process.cwd(), maxBuffer: 16 * 1024 * 1024 },
);
const regenerated = JSON.parse(stdout);
assert.deepEqual(regenerated, ESINT_GLYPH_PAYLOAD);

console.log(
  `VisualTeX esint package regression: PASS (${expectedSlots.size} official commands)`,
);
