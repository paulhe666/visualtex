import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { gunzipSync, strFromU8 } from "fflate";
import {
  EXTENDED_INTEGRAL_COMMANDS,
  EXTENDED_INTEGRAL_MATHML_MACROS,
  EXTENDED_INTEGRAL_SVG_MACROS,
} from "../../shared/visualtex-app/src/math/extendedIntegralCompatibility.ts";
import {
  ESINT_GLYPH_PAYLOAD,
  ESINT_INTEGRAL_GLYPHS,
  ESINT_INTEGRAL_GLYPHS_BY_COMMAND,
} from "../src/math/esintGlyphs.ts";
import {
  RARE_INTEGRAL_GLYPHS_GZIP_BASE64,
  RARE_INTEGRAL_GLYPHS_JSON_SHA256,
} from "../../shared/visualtex-app/src/math/rareIntegralGlyphs.generatedData.ts";

const decodedRareJson = strFromU8(
  gunzipSync(Uint8Array.from(Buffer.from(RARE_INTEGRAL_GLYPHS_GZIP_BASE64, "base64"))),
);
assert.equal(
  createHash("sha256").update(decodedRareJson).digest("hex"),
  RARE_INTEGRAL_GLYPHS_JSON_SHA256,
  "checked-in compressed rare-integral payload hash changed",
);
const rare = JSON.parse(decodedRareJson) as {
  unitsPerEm: number;
  axisHeight: number;
  source: {
    family: string;
    version: string;
    sha256: string;
    license: string;
  };
  glyphs: Array<{
    command: string;
    aliases: string[];
    small: { path: string };
    large: { path: string };
  }>;
};
assert.equal(rare.unitsPerEm, 1000);
assert.equal(rare.axisHeight, 250);
assert.equal(rare.source.family, "STIX Two Math");
assert.equal(rare.source.version, "2.13 b171");
assert.equal(
  rare.source.sha256,
  "3a5f3f26f40d5698b3c62dd085d48d6663696a3f80825aab8b553d5097518e8c",
);
assert.equal(rare.source.license, "SIL Open Font License 1.1");
assert.equal(rare.glyphs.length, 21);
for (const glyph of rare.glyphs) {
  assert.ok(glyph.small.path.startsWith("M") && glyph.small.path.endsWith("Z"), glyph.command);
  assert.ok(glyph.large.path.startsWith("M") && glyph.large.path.endsWith("Z"), glyph.command);
  assert.ok(glyph.large.path.length > glyph.small.path.length / 2, glyph.command);
}

assert.equal(ESINT_GLYPH_PAYLOAD.source.family, "esint10");
assert.equal(ESINT_GLYPH_PAYLOAD.source.license, "Public Domain");
assert.equal(ESINT_INTEGRAL_GLYPHS.length, 12);
assert.equal(
  ESINT_INTEGRAL_GLYPHS_BY_COMMAND.dotsint,
  ESINT_INTEGRAL_GLYPHS_BY_COMMAND.idotsint,
);
assert.equal(
  ESINT_INTEGRAL_GLYPHS_BY_COMMAND.intclockwise,
  ESINT_INTEGRAL_GLYPHS_BY_COMMAND.ointclockwise,
);

assert.equal(EXTENDED_INTEGRAL_COMMANDS.length, 33);
for (const command of EXTENDED_INTEGRAL_COMMANDS) {
  assert.ok(EXTENDED_INTEGRAL_MATHML_MACROS[command], `${command} MathML macro`);
  assert.match(
    EXTENDED_INTEGRAL_SVG_MACROS[command],
    new RegExp(`visualtex-integral-export-${command}`),
    `${command} SVG marker`,
  );
}

const tauriConfig = JSON.parse(
  await readFile(new URL("../src-tauri/tauri.conf.json", import.meta.url), "utf8"),
) as {
  bundle?: { resources?: Record<string, string> };
};
assert.equal(
  tauriConfig.bundle?.resources?.["../../shared/visualtex-app/src/math/STIXTwoMath-OFL-1.1.txt"],
  "licenses/STIXTwoMath-OFL-1.1.txt",
  "Windows installer must ship the STIX Two Math OFL next to public licenses",
);
const license = await readFile(
  new URL("../../shared/visualtex-app/src/math/STIXTwoMath-OFL-1.1.txt", import.meta.url),
  "utf8",
);
assert.match(license, /SIL OPEN FONT LICENSE Version 1\.1/);
assert.match(license, /Copyright 2001-2021 The STIX Fonts Project Authors/);

console.log(
  `VisualTeX Windows integral distribution regression: PASS (${EXTENDED_INTEGRAL_COMMANDS.length} commands, ${rare.glyphs.length} STIX rare glyphs, ${ESINT_INTEGRAL_GLYPHS.length} esint overrides)`,
);
