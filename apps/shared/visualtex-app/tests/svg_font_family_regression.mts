// Exported SVGs name their fonts without quotes where CSS allows it: WPS draws an SVG <text> whose font-family is
// quoted ("Times New Roman", serif) in a sans-serif fallback. The whole export (the editor engine, with Times New Roman
// for letters, Greek letters and digits) writes no quoted font name.
// Run from apps/macos: tsx --import ../shared/visualtex-app/testing/node-loader.mjs ../shared/visualtex-app/tests/svg_font_family_regression.mts
import { test } from "node:test";
import assert from "node:assert/strict";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const exportDir = pathToFileURL(join(dirname(fileURLToPath(import.meta.url)), "..", "src", "export")).href;
const { bareFontFamilyList, bareSvgFontFamilies } = await import(`${exportDir}/svgFontFamilies.ts`);

test("names that are runs of identifiers lose their quotes", () => {
  assert.equal(bareFontFamilyList('"Times New Roman", serif'), "Times New Roman, serif");
  assert.equal(bareFontFamilyList("'Times New Roman', serif"), "Times New Roman, serif");
  assert.equal(
    bareFontFamilyList('KaiTi, "KaiTi_GB2312", "Kaiti SC", STKaiti, serif'),
    "KaiTi, KaiTi_GB2312, Kaiti SC, STKaiti, serif",
  );
  assert.equal(bareFontFamilyList('"楷体", serif'), "楷体, serif");
});

test("names CSS cannot read bare keep their quotes", () => {
  // a word starting with a digit, punctuation, and a name that would read as a generic family or a keyword
  assert.equal(bareFontFamilyList('"3D Font", serif'), '"3D Font", serif');
  assert.equal(bareFontFamilyList('"VTX  Roman 10"'), '"VTX Roman 10"');
  assert.equal(bareFontFamilyList('"Font (Pro)", serif'), '"Font (Pro)", serif');
  assert.equal(bareFontFamilyList('"serif"'), '"serif"');
  assert.equal(bareFontFamilyList('"Inherit", serif'), '"Inherit", serif');
  // unbalanced quotes: left alone
  assert.equal(bareFontFamilyList('"Times New Roman, serif'), '"Times New Roman, serif');
  // nothing quoted: unchanged
  assert.equal(bareFontFamilyList("Times New Roman, serif"), "Times New Roman, serif");
});

test("font-family attributes in an SVG are rewritten, everything else is kept", () => {
  const svg =
    '<svg><text x="1" font-family="&quot;Times New Roman&quot;, serif" font-style="italic">x</text>' +
    "<text font-family='\"Kaiti SC\", serif'>圆</text>" +
    '<text font-family="&quot;3D Font&quot;, serif">y</text><path d="M0 0"/></svg>';
  assert.equal(
    bareSvgFontFamilies(svg),
    '<svg><text x="1" font-family="Times New Roman, serif" font-style="italic">x</text>' +
      "<text font-family='Kaiti SC, serif'>圆</text>" +
      '<text font-family="&quot;3D Font&quot;, serif">y</text><path d="M0 0"/></svg>',
  );
});
