import assert from "node:assert/strict";
import { DOMParser } from "@xmldom/xmldom";
import { formatLatexLines, latexCodeFormats, parseLatexSourceDraft } from "../src/clipboard/LatexCopyService.ts";
import { findLatexFormulaSpans, parseLatexMarkdownDocument } from "../src/office/documentImport/documentImportParser.ts";
import { findWindowsWordLatexRedrawSpans } from "../src/office/redraw/wordLatexRedrawParser.ts";
import { renderOfficeFormulaArtifacts } from "../src/office/shared/formulaRenderArtifacts.ts";
import { normalizeMathModeSource } from "../src/math/mathModeSource.ts";
import { stripVisualTexAlignmentMarkers } from "../src/editor/alignmentMarkers.ts";

globalThis.DOMParser ??= DOMParser;
const doc = new DOMParser().parseFromString("<root/>", "application/xml");
Object.getPrototypeOf(doc).querySelector ??= function(name) { return this.getElementsByTagName(name)?.item(0) ?? null; };
if (!("children" in Object.getPrototypeOf(doc.documentElement))) Object.defineProperty(Object.getPrototypeOf(doc.documentElement), "children", { get() { return Array.from(this.childNodes ?? []).filter(n => n.nodeType === 1); } });

const formulas = [
  String.raw`\text{前}\left(a\text{ 和 }b\right)\text{后}`,
  String.raw`\left\{\left(x\text{ 或 }y\right)\middle|z\right.`,
  String.raw`x_\alpha+\sqrt x+\frac12`,
  String.raw`\frac{\text{甲}}{\text{乙}}`,
  String.raw`a\over\text{总数}`,
  String.raw`\text{下标}_1+\text{上标}^2`,
  String.raw`\begin {cases}a&\text{如果 }x>0\\b&\text{否则}\end {cases}`,
  String.raw`\begin{pmatrix}\frac{1}{2}&a\\b&\sqrt[3]{c}\end{pmatrix}`,
  String.raw`\sum_{i=1}^{n}x_i+\int_0^1 \sin x\,\mathrm{d}x`,
  String.raw`\overset{\text{条件}}{=}\operatorname{rank}(A)+\binom{n}{k}`,
  String.raw`\text{价格 \$5 与 }x\%`,
  String.raw`\text{说明}$\max\left(a,\;b\right)$\text{结束}`,
  "x+% a fake delimiter $ and \\text{comment}\n y",
];
let renderCount = 0;
for (const formula of formulas) {
  for (const { id } of latexCodeFormats) {
    const source = formatLatexLines([formula], id);
    const draft = parseLatexSourceDraft(source, id);
    assert.equal(draft.valid, true, `${id}: ${source}: ${draft.error}`);
    assert.deepEqual(draft.values.map(stripVisualTexAlignmentMarkers), [normalizeMathModeSource(formula)], `${id} must retain the complete formula`);
    if (id === "raw") continue;
    const redraw = findWindowsWordLatexRedrawSpans(source);
    const imported = parseLatexMarkdownDocument(source, "auto", 12).filter(b => b.kind === "formula");
    assert.ok(redraw.length > 0, `${id}: scanner lost all formulas`);
    assert.deepEqual(redraw.map(s => s.latex), imported.map(s => s.latex), `${id}: entry points disagree`);
    for (const span of redraw) {
      assert.equal(source.slice(span.start, span.end), span.sourceText);
      let artifacts;
      try { artifacts = renderOfficeFormulaArtifacts({ lines: [{ id: "test", latex: span.latex }], codeFormat: "raw", host: "word", displayMode: span.displayMode }); } catch (error) { throw new Error(`${id}: ${span.latex}`, { cause: error }); }
      assert.ok(artifacts.svg.width > 0 && artifacts.svg.height > 0);
      assert.ok(artifacts.omml?.ommlBase64);
      renderCount += 1;
    }
  }
}

const sources = [
  String.raw`前😀 \$5 $x+1$ \(y+\\)z\) \[a+\\]b\] 后`,
  "```latex\n$x$\n```\n`$y$` and $z$",
  "% $ignored$\n\\[x% \\] ignored\n+y\\]",
  String.raw`\begin { alignat }{2}a&=b&c&=d\end { alignat } \begin{math}x\end{math}`,
  String.raw`\begin{equation}\begin{equation}a\end{equation}+b\end{equation}`,
];
for (const source of sources) {
  const spans = findLatexFormulaSpans(source);
  const redraw = findWindowsWordLatexRedrawSpans(source);
  assert.deepEqual(redraw.map(s => [s.start, s.end, s.latex]), spans.map(s => [s.start, s.end, s.latex]));
  for (const span of redraw) assert.equal(source.slice(span.start, span.end), span.sourceText);
}
assert.equal(findLatexFormulaSpans(sources[1]).length, 1);
assert.equal(findLatexFormulaSpans(sources[2]).length, 1);
assert.equal(findLatexFormulaSpans(sources[3]).length, 2);
assert.equal(findLatexFormulaSpans(sources[4]).length, 1);
assert.equal(normalizeMathModeSource(String.raw`\text{price $x$ and \$5}$y$`), String.raw`\text{price $x$ and \$5}y`);
assert.equal(normalizeMathModeSource("x$"), "x$", "Do not guess an unfinished delimiter");
assert.equal(parseLatexSourceDraft(String.raw`\sqrt\text{中文}`, "raw").valid, false,
  "A forgiving MathLive preview is not proof that TeX can compile the source");
assert.equal(parseLatexSourceDraft(String.raw`\left(a\text{ 和 }b\right)`, "raw").valid, true);
console.log(`Source/render contract PASS: ${formulas.length} formulas × ${latexCodeFormats.length} source formats; ${renderCount} SVG+OMML renders; shared scanner cases`);
