// Update Equation Format's font step: which formulas were drawn in other fonts (formulaFontsDiffer, the rule the
// Windows add-ins and the macOS app share), the confirmation's font names, and the faces record of the metadata.
// Run from apps/macos: tsx --import ../shared/visualtex-app/testing/node-loader.mjs ../shared/visualtex-app/tests/formula_font_refresh_regression.mts
import { test } from "node:test";
import assert from "node:assert/strict";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const shared = pathToFileURL(join(dirname(fileURLToPath(import.meta.url)), "..", "src", "office", "shared")).href;
const { formulaFontsDiffer, describeFormulaFonts } = await import(`${shared}/formulaFontRefresh.ts`);
const {
  createFormulaMetadata,
  decodeFormulaMetadata,
  encodeFormulaMetadata,
  isVisualTeXFormulaMetadata,
} = await import(`${shared}/formulaMetadata.ts`);

const TNR = "Times New Roman";
const fonts = (letterFont: string, chineseFont: string, faces: [string, string, string]) => ({
  letterFont,
  chineseFont,
  faces: { letter: faces[0], greek: faces[1], digit: faces[2] },
});
const plain = fonts("katex", "system", ["", "", ""]);
const timesFaces = fonts("katex", "system", [TNR, TNR, TNR]);
const record = (latex: string, extra: Record<string, unknown> = {}) => ({ latex, ...extra });

test("a missing record is what VisualTeX drew without it", () => {
  assert.equal(formulaFontsDiffer(record("x+y"), plain), false);
  assert.equal(formulaFontsDiffer(record("x+y"), fonts("times", "system", ["", "", ""])), true);
  // faces of the Settings, formula drawn before faces existed
  assert.equal(formulaFontsDiffer(record("x+y", { formulaLetterFont: "katex", formulaChineseFont: "system" }), timesFaces), true);
  // names compare trimmed and without case
  assert.equal(
    formulaFontsDiffer(
      record("x+y", { formulaLetterFont: "KaTeX", formulaFaces: { letter: "times new roman ", greek: TNR, digit: TNR } }),
      timesFaces,
    ),
    false,
  );
  // one role missing counts as the math font's
  assert.equal(formulaFontsDiffer(record("x+y", { formulaFaces: { letter: TNR } }), timesFaces), true);
  assert.equal(formulaFontsDiffer(record("x", { formulaFaces: { letter: "", greek: "", digit: "" } }), plain), false);
  // an earlier version's letter font is drawn in KaTeX now: drawn again
  assert.equal(formulaFontsDiffer(record("x", { formulaLetterFont: "cambria" }), plain), true);
});

test("the Chinese font counts only for Chinese text", () => {
  const kaiti = fonts("katex", "kaiti", ["", "", ""]);
  assert.equal(formulaFontsDiffer(record("x^2", { formulaChineseFont: "system" }), kaiti), false);
  assert.equal(formulaFontsDiffer(record("\\text{速度}=v", { formulaChineseFont: "system" }), kaiti), true);
  assert.equal(formulaFontsDiffer(record("a，b"), kaiti), true);
  assert.equal(formulaFontsDiffer(record("\\text{速度}", { formulaChineseFont: "kaiti" }), kaiti), false);
});

test("the confirmation names the math font and the faces", () => {
  assert.equal(describeFormulaFonts(plain, true), "KaTeX");
  assert.equal(
    describeFormulaFonts(fonts("times", "system", [TNR, "", "Arial"]), true),
    "newtxmath; letters: Times New Roman, digits: Arial",
  );
  assert.equal(
    describeFormulaFonts(fonts("katex", "system", [TNR, TNR, ""]), false),
    "KaTeX；字母：Times New Roman，希腊字母：Times New Roman",
  );
});

test("the metadata records the faces and keeps them through encoding", () => {
  const line = { id: "12345678-1234-4234-9234-123456789abd", latex: "x+y" };
  const input = { formulaId: "12345678-1234-4234-9234-123456789abc", title: "t", lines: [line], codeFormat: "raw" };
  const recorded = createFormulaMetadata({ ...input, formulaFaces: { letter: TNR, greek: TNR, digit: "" } });
  assert.deepEqual(recorded.formulaFaces, { letter: TNR, greek: TNR, digit: "" });
  const decoded = decodeFormulaMetadata(encodeFormulaMetadata(recorded));
  assert.deepEqual(decoded?.formulaFaces, { letter: TNR, greek: TNR, digit: "" });
  assert.equal(formulaFontsDiffer(decoded!, fonts("katex", "system", [TNR, TNR, ""])), false);
  // without faces given: the ones the layout engine draws in (none stored here)
  assert.deepEqual(createFormulaMetadata(input).formulaFaces, { letter: "", greek: "", digit: "" });
});

test("the validator takes the faces either app writes and nothing else", () => {
  const line = { id: "12345678-1234-4234-9234-123456789abd", latex: "x" };
  const base = createFormulaMetadata({
    formulaId: "12345678-1234-4234-9234-123456789abc",
    title: "t",
    lines: [line],
    codeFormat: "raw",
  });
  assert.equal(isVisualTeXFormulaMetadata({ ...base, formulaFaces: { letter: TNR, greek: "", digit: null } }), true);
  assert.equal(isVisualTeXFormulaMetadata({ ...base, formulaFaces: {} }), true);
  assert.equal(isVisualTeXFormulaMetadata({ ...base, formulaFaces: undefined }), true);
  assert.equal(isVisualTeXFormulaMetadata({ ...base, formulaFaces: "Times" }), false);
  assert.equal(isVisualTeXFormulaMetadata({ ...base, formulaFaces: { letter: 3 } }), false);
  assert.equal(isVisualTeXFormulaMetadata({ ...base, formulaFaces: { digit: "x".repeat(129) } }), false);
});
