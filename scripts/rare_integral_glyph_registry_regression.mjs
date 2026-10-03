import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  RARE_INTEGRAL_GLYPHS,
  RARE_INTEGRAL_GLYPHS_BY_COMMAND,
  RARE_INTEGRAL_GLYPH_AXIS_HEIGHT,
  RARE_INTEGRAL_GLYPH_SOURCE,
  RARE_INTEGRAL_GLYPH_UNITS_PER_EM,
} from "../src/math/rareIntegralGlyphs.generated.ts";

const expected = [
  ["intclockwise", 0x2231],
  ["varointclockwise", 0x2232],
  ["ointctrclockwise", 0x2233],
  ["sumint", 0x2a0b],
  ["iiiint", 0x2a0c],
  ["intbar", 0x2a0d],
  ["intBar", 0x2a0e],
  ["fint", 0x2a0f],
  ["cirfnint", 0x2a10],
  ["awint", 0x2a11],
  ["rppolint", 0x2a12],
  ["scpolint", 0x2a13],
  ["npolint", 0x2a14],
  ["pointint", 0x2a15],
  ["quatint", 0x2a16],
  ["intlarhk", 0x2a17],
  ["intx", 0x2a18],
  ["intcap", 0x2a19],
  ["intcup", 0x2a1a],
  ["upint", 0x2a1b],
  ["lowint", 0x2a1c],
];

const licenseResourceSource = "../src/math/STIXTwoMath-OFL-1.1.txt";
const licenseResourceTarget = "licenses/STIXTwoMath-OFL-1.1.txt";
const licenseUrl = new URL(
  "../src/math/STIXTwoMath-OFL-1.1.txt",
  import.meta.url,
);
const tauriConfigUrl = new URL("../src-tauri/tauri.conf.json", import.meta.url);
const tauriConfig = JSON.parse(await readFile(tauriConfigUrl, "utf8"));
const bundleResources = tauriConfig.bundle?.resources;

assert.equal(
  bundleResources?.[licenseResourceSource],
  licenseResourceTarget,
  "Tauri bundle must ship the STIX Two Math OFL at the public licenses path",
);
assert.equal(
  new URL(licenseResourceSource, tauriConfigUrl).href,
  licenseUrl.href,
  "Tauri resource mapping must use the canonical source license",
);
assert.equal(
  Object.entries(bundleResources ?? {}).filter(
    ([source, target]) =>
      source === licenseResourceSource || target === licenseResourceTarget,
  ).length,
  1,
  "STIX Two Math OFL must have exactly one bundle mapping",
);

assert.equal(RARE_INTEGRAL_GLYPH_UNITS_PER_EM, 1000);
assert.equal(RARE_INTEGRAL_GLYPH_AXIS_HEIGHT, 250);
assert.equal(RARE_INTEGRAL_GLYPH_SOURCE.family, "STIX Two Math");
assert.equal(RARE_INTEGRAL_GLYPH_SOURCE.version, "2.13 b171");
assert.equal(RARE_INTEGRAL_GLYPH_SOURCE.sourceAxisHeight, 258);
assert.equal(RARE_INTEGRAL_GLYPH_SOURCE.targetAxisHeight, 250);
assert.equal(
  RARE_INTEGRAL_GLYPH_SOURCE.sha256,
  "3a5f3f26f40d5698b3c62dd085d48d6663696a3f80825aab8b553d5097518e8c",
);

assert.deepEqual(
  RARE_INTEGRAL_GLYPHS.map(({ command, codePoint }) => [command, codePoint]),
  expected,
);
assert.equal(
  new Set(RARE_INTEGRAL_GLYPHS.map(({ codePoint }) => codePoint)).size,
  expected.length,
);

for (const glyph of RARE_INTEGRAL_GLYPHS) {
  assert.equal(glyph.character.codePointAt(0), glyph.codePoint, glyph.command);
  assert.equal(RARE_INTEGRAL_GLYPHS_BY_COMMAND[glyph.command], glyph);
  const expectedAliases = glyph.command === "awint" ? ["intctrclockwise"] : [];
  assert.deepEqual(glyph.aliases, expectedAliases, `${glyph.command} aliases`);
  for (const alias of glyph.aliases) {
    assert.equal(RARE_INTEGRAL_GLYPHS_BY_COMMAND[alias], glyph, alias);
  }

  for (const [size, variant, targetAdvance] of [
    ["small", glyph.small, 1111],
    ["large", glyph.large, 2222],
  ]) {
    assert.ok(variant.path.startsWith("M"), `${glyph.command} ${size} SVG M`);
    assert.ok(variant.path.endsWith("Z"), `${glyph.command} ${size} SVG Z`);
    assert.equal(
      variant.mathJaxPath,
      variant.path.slice(1, -1),
      `${glyph.command} ${size} MathJax payload`,
    );
    assert.ok(!variant.mathJaxPath.startsWith("M"));
    assert.ok(!variant.mathJaxPath.endsWith("Z"));
    assert.equal(variant.verticalAdvance, targetAdvance);
    assert.ok(variant.sourceVerticalAdvance > 0);
    assert.ok(
      Math.abs(
        variant.scale - targetAdvance / variant.sourceVerticalAdvance,
      ) < 1e-12,
      `${glyph.command} ${size} normalization scale`,
    );
    assert.ok(variant.advanceWidth > 0);
    assert.ok(variant.italicCorrection >= 0);
    assert.ok(variant.bounds.xMax > variant.bounds.xMin);
    assert.ok(variant.bounds.yMax > variant.bounds.yMin);
    assert.equal(variant.height, Math.max(0, variant.bounds.yMax));
    assert.equal(variant.depth, Math.max(0, -variant.bounds.yMin));
    assert.ok(
      Math.abs(targetAdvance - (variant.height + variant.depth)) <= 3,
      `${glyph.command} ${size} normalized vertical extent`,
    );
  }

  assert.equal(glyph.large.glyphName, `${glyph.small.glyphName}.dsp`);
  assert.ok(glyph.large.height + glyph.large.depth > glyph.small.height + glyph.small.depth);
}

const license = await readFile(
  licenseUrl,
  "utf8",
);
assert.match(license, /Copyright 2001-2021 The STIX Fonts Project Authors/);
assert.match(license, /SIL OPEN FONT LICENSE Version 1\.1/);
assert.match(license, /Reserved Font Name "TM Math"/);

console.log(
  `Rare-integral glyph registry passed: ${RARE_INTEGRAL_GLYPHS.length} code points, ` +
    `${Object.keys(RARE_INTEGRAL_GLYPHS_BY_COMMAND).length} commands/aliases.`,
);
