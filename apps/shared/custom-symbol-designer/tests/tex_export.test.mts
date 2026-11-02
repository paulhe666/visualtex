// node --import tsx apps/shared/custom-symbol-designer/tests/tex_export.test.mts  (or npx tsx <file>)
// The designer's code export without a browser: SVG path data reduced to what
// pict2e draws, matrices and their rotate-scale-rotate form, the symbol's box,
// and the picture-mode \newcommand of a document (ordinary and operator).
import assert from "node:assert/strict";
import { normalizePathData, pathSegmentsBounds } from "../src/pathData.ts";
import { customSymbolTexExport, decomposeLinear, nativeOperatorBase } from "../src/tex.ts";
import { apply, invert, multiply, rotate, scale, transformMatrix } from "../src/matrix.ts";
import { baseLayerOf, computeSymbolBox, isLatexMaterial } from "../src/symbolBox.ts";
import { isLargeOperatorMaterial, textStyleVariant } from "../src/glyphAsset.ts";
import type { CustomSymbolDesignerDocument, CustomSymbolDesignerGlyphLayer } from "../src/designerTypes.ts";

const near = (actual: number, expected: number, epsilon = 1e-6, label = "") =>
  assert.ok(Math.abs(actual - expected) <= epsilon, `${label} ${actual} != ${expected}`);

// path data: relative commands, H/V, S/T reflection, Q->C, arcs
{
  const segments = normalizePathData("m10 10h20v20H10zM0 0q10 -10 20 0t20 0");
  assert.deepEqual(segments.slice(0, 5).map((s) => s.op), ["M", "L", "L", "L", "Z"]);
  const third = segments[2];
  assert.ok(third.op === "L" && third.x === 30 && third.y === 30);
  const quad = segments[6];
  assert.ok(quad.op === "C");
  near(quad.x1, 20 / 3, 1e-9, "q->c x1");
  near(quad.y1, -20 / 3, 1e-9, "q->c y1");
  const smooth = segments[7];
  assert.ok(smooth.op === "C" && smooth.x === 40 && smooth.y === 0);
  near(smooth.y1, 20 / 3, 1e-9, "t reflection");
}
{
  const bounds = pathSegmentsBounds(normalizePathData("M-100 0A100 100 0 1 0 100 0A100 100 0 1 0 -100 0Z"))!;
  near(bounds.x, -100, 0.01, "circle left");
  near(bounds.width, 200, 0.01, "circle width");
  near(bounds.height, 200, 0.4, "circle height");
}

// matrices, and any 2x2 matrix as rotate * scale * rotate (graphicx's \rotatebox and \scalebox)
{
  const matrix = transformMatrix({ translateX: 10, rotateDeg: 90, originX: 50, originY: 50 });
  const point = apply(matrix, 100, 50);
  near(point.x, 60, 1e-9);
  near(point.y, 100, 1e-9);
  const back = apply(invert(matrix), point.x, point.y);
  near(back.x, 100, 1e-9);
  near(back.y, 50, 1e-9);
  for (const [p, q, r, s] of [
    [-1, 0.26795, 0, 1], // flipped and slanted
    [0.5, 0.2, -0.1, 2],
    [1, 0, 0, -1],
    [0, -1, 1, 0],
    [2, 0, 0, 0.5],
  ]) {
    const { phi, sx, sy, theta } = decomposeLinear(p, q, r, s);
    const m = multiply(multiply(rotate(phi), scale(sx, sy)), rotate(theta));
    [p, r, q, s].forEach((value, index) => near(m[index], value, 1e-9, `decompose [${p},${q},${r},${s}] ${index}`));
  }
}

const marker = [{ kind: "path" as const, d: "M0 0L10 0", fill: true, transform: { matrix: [1, 0, 0, -1, 0, 0] as [number, number, number, number, number, number] } }];
const glyph = (id: string, source: string, metrics: { widthEm: number; ascentEm: number; descentEm: number }, transform: CustomSymbolDesignerGlyphLayer["transform"]): CustomSymbolDesignerGlyphLayer => ({
  id,
  name: source,
  kind: "glyph",
  visible: true,
  locked: false,
  transform,
  clipRect: null,
  asset: { sourceLatex: source, metrics, shapes: marker },
});
const BASELINE = 3000;
const INT = { widthEm: 1, ascentEm: 1.36112, descentEm: 0.86112 };

// an ordinary symbol: alpha (cropped, scaled and turned) over a bar, an eraser left out
{
  const document: CustomSymbolDesignerDocument = {
    version: 1,
    symbolId: null,
    name: "Test",
    command: "\\mysym",
    role: "relation",
    limitsBehavior: "auto",
    metrics: { widthEm: 3.2, ascentEm: 3, descentEm: 1.5 },
    ommlFallback: null,
    baseLayerId: null,
    layers: [
      { ...glyph("a", "\\alpha", { widthEm: 0.6434, ascentEm: 0.442555, descentEm: 0.011 }, { translateX: 100, translateY: BASELINE - 442.555, scaleX: 0.5, scaleY: 0.5, rotateDeg: 20, originX: 321.7, originY: 226.8 }), clipRect: { x: 0, y: 0, width: 300, height: 453.555 } },
      { id: "b", name: "bar", kind: "geometry", geometryPreset: "line", visible: true, locked: false, transform: {}, shape: { kind: "line", x1: 0, y1: 2800, x2: 700, y2: 2800, strokeWidth: 40, lineCap: "round" }, bounds: { x: 0, y: 2780, width: 700, height: 40 } },
      { id: "c", name: "erase", kind: "geometry", geometryPreset: "eraser", visible: true, locked: false, transform: {}, shape: { kind: "path", d: "M0 2900L100 2900", operation: "erase", fill: false, strokeWidth: 30 }, bounds: { x: 0, y: 2885, width: 100, height: 30 } },
      { id: "d", name: "hidden", kind: "geometry", visible: false, locked: false, transform: {}, shape: { kind: "circle", cx: 0, cy: 0, r: 10 }, bounds: { x: -10, y: -10, width: 20, height: 20 } },
    ],
  };
  assert.equal(isLatexMaterial(document.layers[0]), true);
  assert.equal(isLatexMaterial({ ...document.layers[0], asset: { ...(document.layers[0] as CustomSymbolDesignerGlyphLayer).asset, sourceLatex: "ħ" } } as never), false);
  assert.equal(baseLayerOf(document), null, "baseLayerId null: the ink's box");
  const box = computeSymbolBox(document, { left: 0, top: 2550, right: 700, bottom: 3010 });
  near(box.left, -50, 1e-9, "ink box left (margin)");
  near(box.metrics.ascentEm, 0.5, 1e-9);
  const { code, omissions } = customSymbolTexExport(document, box);
  assert.match(code, /^% Preamble: \\usepackage\{amsmath,graphicx,pict2e,trimclip\}$/m);
  assert.match(code, /^\\providecommand\*\\vtxpin\[1\]/m);
  assert.match(code, /^\\newcommand\{\\mysym\}\{\\mathrel\{\\text\{\\setlength\{\\unitlength\}\{1em\}\\raisebox\{-0\.06em\}\{%$/m);
  assert.match(code, /\\begin\{picture\}\(0\.8,0\.56\)\(0,-0\.06\)%/);
  // the turned alpha: \put at its reference point, rotate-scale-rotate around a pinned, cropped box
  assert.match(code, /\\put\([\d.-]+,[\d.-]+\)\{\\rotatebox\{-20\}\{\\scalebox\{0\.5\}\{\\vtxpin\{\\clipbox\*\{0em -0\.011em 0\.3em 0\.44256em\}\{\$\\alpha\$\}\}\}\}\}%/);
  // the bar: a pict2e path in em, the picture's origin at the box's left end of the baseline
  assert.match(code, /\\linethickness\{0\.04em\}\\roundcap\\miterjoin\\moveto\(0\.05,0\.2\)\\lineto\(0\.75,0\.2\)\\strokepath%/);
  assert.ok(!code.includes("hidden"));
  assert.deepEqual(omissions, ["erase: eraser"]);
}

// an operator on \displaystyle\int: the box is the integral's, limits as TeX sets them
{
  const document: CustomSymbolDesignerDocument = {
    version: 1,
    symbolId: null,
    name: "Dot integral",
    command: "dotint",
    role: "operator",
    limitsBehavior: "auto",
    metrics: { widthEm: 3.2, ascentEm: 3, descentEm: 1.5 },
    ommlFallback: null,
    layers: [
      glyph("i", "\\displaystyle\\int", INT, { translateX: 200, translateY: BASELINE - 1361.12 }),
      glyph("dot", "\\cdot", { widthEm: 0.2778, ascentEm: 0.31, descentEm: -0.19 }, { translateX: 400, translateY: BASELINE - 500 }),
    ],
  };
  assert.equal(isLargeOperatorMaterial("\\displaystyle\\int"), true);
  assert.equal(textStyleVariant("\\displaystyle\\int"), "\\textstyle\\int");
  assert.equal(baseLayerOf(document)?.id, "i", "the integral is the base");
  const italics: Record<string, number> = { "\\displaystyle\\int": 0.44444, "\\textstyle\\int": 0.19444 };
  const box = computeSymbolBox(document, { left: 230, top: 1650, right: 1150, bottom: 3850 }, (s) => italics[s]);
  near(box.left, 200, 1e-9, "box = the integral's");
  near(box.metrics.widthEm, 1, 1e-9);
  near(box.metrics.ascentEm, 1.36112, 1e-9);
  near(box.metrics.descentEm, 0.86112, 1e-9);
  near(box.italicEm, 0.44444, 1e-9);
  assert.equal(box.artworkBounds, undefined, "the dot stays inside the box");
  const textStyle = (source: string) =>
    source === "\\displaystyle\\int" ? { metrics: { widthEm: 0.66667, ascentEm: 0.80556, descentEm: 0.30556 }, italicEm: 0.19444 } : undefined;

  // the integral kept as itself: the dot over it, then \int, which TeX sets as always
  assert.equal(nativeOperatorBase(document, box)?.id, "i");
  const native = customSymbolTexExport(document, box, { textStyle, mathFont: "ntx" }).code;
  assert.match(native, /^% Preamble: \\usepackage\{amsmath,graphicx\}$/m);
  assert.match(native, /^% Designed in the Times math font: with \\usepackage\{newtxmath\}/m);
  assert.match(native, /^\\newcommand\{\\dotint\}\{\\mathop\{\\mathchoice$/m);
  // the dot at its place from the integral's reference point; in text style placed by the width and height ratios, sized by the width ratio
  assert.ok(native.includes("    {\\text{\\setlength{\\unitlength}{1em}\\begin{picture}(0,0)\\put(0.2,0.19){$\\cdot$}\\end{picture}}}"));
  assert.ok(native.includes("\\put(0.13333,0.095){\\scalebox{0.66667}{$\\cdot$}}"), "scaling alone needs no pin");
  assert.ok(!native.includes("vtxpin"));
  assert.match(native, /^  \\mskip-\\thinmuskip\\int\\displaylimits\}$/m);
  assert.match(
    customSymbolTexExport({ ...document, limitsBehavior: "nolimits" }, box, { textStyle }).code,
    /\\mskip-\\thinmuskip\\int\\nolimits\}$/m,
  );
  assert.equal(
    customSymbolTexExport({ ...document, layers: [document.layers[0]] }, box).code.split("\n").pop(),
    "\\newcommand{\\dotint}{\\int\\displaylimits}",
  );

  // a base that is not kept as itself (here: with an effect) -> the picture-box form
  const drawn: CustomSymbolDesignerDocument = {
    ...document,
    layers: [{ ...document.layers[0], effects: { outline: { enabled: true, width: 30 } } }, document.layers[1]],
  };
  const drawnBox = computeSymbolBox(drawn, { left: 230, top: 1650, right: 1150, bottom: 3850 }, (s) => italics[s]);
  assert.equal(nativeOperatorBase(drawn, drawnBox), null);
  const { code } = customSymbolTexExport(drawn, drawnBox, { textStyle });
  // \limits or \nolimits typed after the command (#1, #2), then the limits (#3, #4)
  assert.match(code, /^\\NewDocumentCommand\{\\dotint\}\{t\\limits t\\nolimits e\{\^_\}\}\{\\mathchoice$/m);
  const limitsDisplay =
    "\\limits\\IfValueT{#3}{^{\\kern0.22222\\fontdimen6\\textfont2\\relax#3\\kern-0.22222\\fontdimen6\\textfont2\\relax}}" +
    "\\IfValueT{#4}{_{\\kern-0.22222\\fontdimen6\\textfont2\\relax#4\\kern0.22222\\fontdimen6\\textfont2\\relax}}";
  const nolimitsText = "\\nolimits\\IfValueT{#3}{^{#3}}\\IfValueT{#4}{_{\\kern-0.19444\\fontdimen6\\textfont2\\relax#4}}";
  // display: \limits by default, the upper limit half the italic correction right, the lower half left
  assert.ok(code.includes(`\\IfBooleanTF{#1}{${limitsDisplay}}`), "\\limits typed: limits");
  assert.ok(code.includes(`}{${limitsDisplay}}}}`), "display default: limits");
  // text and script sizes: \nolimits by default, the subscript in by the text integral's italic correction
  assert.ok(code.includes(`}{${nolimitsText}}}}`), "text default: nolimits");
  assert.match(code, /_\{\\kern-0\.19444\\fontdimen6\\scriptfont2\\relax#4\}/);
  assert.match(code, /_\{\\kern-0\.19444\\fontdimen6\\scriptscriptfont2\\relax#4\}/);
  // the display branch draws \displaystyle\int at the picture's origin, the text branches \textstyle\int
  assert.match(code, /\\begin\{picture\}\(1,2\.22224\)\(0,-0\.86112\)%\n\s+\\put\(0,0\)\{\$\\displaystyle\\int\$\}%/);
  assert.match(code, /\\begin\{picture\}\(0\.66667,1\.11112\)\(0,-0\.30556\)%\n\s+\\put\(0,0\)\{\$\\textstyle\\int\$\}%/);
  // the dot: as designed in display style, scaled with the integral (0.5) about its reference point in text style
  assert.match(code, /\\put\(0\.2,0\.19\)\{\$\\cdot\$\}%/);
  assert.match(code, /\\put\(0\.13333,0\.095\)\{\\scalebox\{0\.66667\}\{\$\\cdot\$\}\}%/);
}
{
  const plain = customSymbolTexExport(
    { version: 1, symbolId: null, name: "", command: "", role: "operator", limitsBehavior: "limits", metrics: { widthEm: 3.2, ascentEm: 3, descentEm: 1.5 }, ommlFallback: null, layers: [] },
    null,
  );
  assert.equal(plain.code, "% \\mysymbol: nothing to draw yet");
}
console.log("custom symbol tex export: ok");
