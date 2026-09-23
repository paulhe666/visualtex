import assert from "node:assert/strict";
import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { latexToMathMl, latexToSvg } from "../src/export/runtime.ts";

const outputArgument = process.argv[2];
if (!outputArgument) {
  throw new Error(
    "Usage: node --experimental-strip-types scripts/office_oiint_export_fixture.mjs <output.json>",
  );
}

const latex = String.raw`\oiint_{\Sigma}F\,\mathrm{d}S11`;
const exportResult = await latexToSvg(latex, {
  displayMode: true,
  fontSizePt: 14,
  paddingPx: 0,
  background: "transparent",
});
const mathMl = await latexToMathMl(latex, true);

assert.match(
  exportResult.svg,
  /data-visualtex-integral="oiint"/,
  "The Office OLE fixture must use the VisualTeX contour-integral vector glyph.",
);
assert.doesNotMatch(
  exportResult.svg,
  />∯<\/text>/,
  "The Office OLE fixture must not use the old small system-font glyph.",
);
assert.ok(
  exportResult.height > 30,
  `The Office OLE fixture is not display-integral height: ${exportResult.height}`,
);
assert.doesNotMatch(
  mathMl,
  /\\oiint|mathcolor="red"|<merror/i,
  "The Office OLE fixture contains an unresolved LaTeX command.",
);

const outputPath = resolve(outputArgument);
await mkdir(dirname(outputPath), { recursive: true });
await writeFile(
  outputPath,
  JSON.stringify(
    {
      svg: exportResult.svg,
      svgBase64: Buffer.from(exportResult.svg, "utf8").toString("base64"),
      mathMl,
      width: exportResult.width,
      height: exportResult.height,
      baseline: exportResult.baseline,
    },
    null,
    2,
  ),
  "utf8",
);
console.log(`Wrote production Office \\oiint export fixture: ${outputPath}`);
