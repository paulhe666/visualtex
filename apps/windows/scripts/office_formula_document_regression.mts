import assert from "node:assert/strict";
import {
  normalizeFormulaEditorDocument,
  serializeFormulaEditorDocument,
  serializeFormulaEditorRenderDocument,
} from "../src/office/shared/formulaEditorDocument.ts";
import {
  createFormulaMetadata,
  decodeFormulaMetadata,
  encodeFormulaMetadata,
} from "../src/office/shared/formulaMetadata.ts";
import { latexToMathMl, latexToSvg } from "../src/export/runtime.ts";
import { isIncompleteLatexDraft } from "../src/math/latexCompatibility.ts";
import {
  canonicalOfficeFingerprintLines,
  isWordOmmlNumberingOnlyEdit,
  persistedOfficeLines,
} from "../src/office/shared/officeEditPersistence.ts";
import { VISUALTEX_ALIGNMENT_MARKER_LATEX } from "../src/editor/alignmentMarkers.ts";

function normalize(source: string, codeFormat: string = "raw") {
  return normalizeFormulaEditorDocument(
    [{ id: "original-line-id", latex: source }],
    codeFormat,
  );
}

const sourceFormattedEquation = normalizeFormulaEditorDocument(
  [
    {
      id: "source-formatted-equation",
      latex: String.raw`\frac{\delta \mathbb{E}[L]}
     {\delta f(\mathbf{x})}
=
2\int
\{f(\mathbf{x})-t\}
p(\mathbf{x},t)\,
\mathrm{d}t
=
0`,
    },
  ],
  "equation",
);
assert.equal(sourceFormattedEquation.lines.length, 1);
assert.equal(sourceFormattedEquation.codeFormat, "equation");
assert.equal(
  sourceFormattedEquation.lines[0].latex,
  String.raw`\frac{\delta \mathbb{E}[L]}{\delta f(\mathbf{x})}=2\int\{f(\mathbf{x})-t\}p(\mathbf{x},t)\,\mathrm{d}t=0`,
);

const multilineCases = [
  {
    name: "align",
    source: String.raw`\begin{align}
a &= b + c \\
d &= e
\end{align}`,
    codeFormat: "align",
    lines: [
      `a ${VISUALTEX_ALIGNMENT_MARKER_LATEX}= b + c`,
      `d ${VISUALTEX_ALIGNMENT_MARKER_LATEX}= e`,
    ],
  },
  {
    name: "align-star",
    source: String.raw`\begin{align*}
x &= y \\
y &= z
\end{align*}`,
    codeFormat: "align-star",
    lines: [
      `x ${VISUALTEX_ALIGNMENT_MARKER_LATEX}= y`,
      `y ${VISUALTEX_ALIGNMENT_MARKER_LATEX}= z`,
    ],
  },
  {
    name: "aligned",
    source: String.raw`\begin{aligned}
p &= q \\
r &= s
\end{aligned}`,
    codeFormat: "aligned",
    lines: [
      `p ${VISUALTEX_ALIGNMENT_MARKER_LATEX}= q`,
      `r ${VISUALTEX_ALIGNMENT_MARKER_LATEX}= s`,
    ],
  },
  {
    name: "gather",
    source: String.raw`\begin{gather}
a=b \\
c=d
\end{gather}`,
    codeFormat: "gather",
    lines: ["a=b", "c=d"],
  },
  {
    name: "gather-star",
    source: String.raw`\begin{gather*}
a=b \\
c=d
\end{gather*}`,
    codeFormat: "gather-star",
    lines: ["a=b", "c=d"],
  },
  {
    name: "multline",
    source: String.raw`\begin{multline}
a+b+c \\
=d+e
\end{multline}`,
    codeFormat: "multline",
    lines: ["a+b+c", "=d+e"],
  },
  {
    name: "multline-star",
    source: String.raw`\begin{multline*}
a+b+c \\
=d+e
\end{multline*}`,
    codeFormat: "multline-star",
    lines: ["a+b+c", "=d+e"],
  },
  {
    name: "equation-split",
    source: String.raw`\begin{equation}
\begin{split}
a &= b \\
c &= d
\end{split}
\end{equation}`,
    codeFormat: "equation-split",
    lines: [
      `a ${VISUALTEX_ALIGNMENT_MARKER_LATEX}= b`,
      `c ${VISUALTEX_ALIGNMENT_MARKER_LATEX}= d`,
    ],
  },
  {
    name: "equation-star-split",
    source: String.raw`\begin{equation*}
\begin{split}
a &= b \\
c &= d
\end{split}
\end{equation*}`,
    codeFormat: "equation-star-split",
    lines: [
      `a ${VISUALTEX_ALIGNMENT_MARKER_LATEX}= b`,
      `c ${VISUALTEX_ALIGNMENT_MARKER_LATEX}= d`,
    ],
  },
] as const;

for (const testCase of multilineCases) {
  const normalized = normalize(testCase.source);
  assert.equal(normalized.codeFormat, testCase.codeFormat, `${testCase.name} format`);
  assert.deepEqual(
    normalized.lines.map((line) => line.latex),
    testCase.lines,
    `${testCase.name} rows`,
  );
  assert.equal(normalized.lines[0].id, "original-line-id");
  assert.equal(new Set(normalized.lines.map((line) => line.id)).size, normalized.lines.length);

  const canonical = serializeFormulaEditorDocument(normalized);
  const secondPass = normalizeFormulaEditorDocument(
    [{ id: "original-line-id", latex: canonical }],
    normalized.codeFormat,
  );
  assert.equal(serializeFormulaEditorDocument(secondPass), canonical, `${testCase.name} round trip`);
  assert.deepEqual(
    secondPass.lines.map((line) => line.latex),
    testCase.lines,
    `${testCase.name} rows after round trip`,
  );

  const metadata = createFormulaMetadata({
    formulaId: crypto.randomUUID(),
    title: testCase.name,
    lines: normalized.lines,
    codeFormat: normalized.codeFormat,
    displayMode: "block",
  });
  const decoded = decodeFormulaMetadata(encodeFormulaMetadata(metadata));
  assert.ok(decoded, `${testCase.name} metadata decoded`);
  const reopened = normalizeFormulaEditorDocument(decoded.lines, decoded.codeFormat);
  assert.equal(
    serializeFormulaEditorDocument(reopened),
    canonical,
    `${testCase.name} canonical source survives Windows Lines + CodeFormat metadata`,
  );

  const renderSource = serializeFormulaEditorRenderDocument(normalized);
  let svg;
  try {
    svg = latexToSvg(renderSource, {
      displayMode: true,
      fontSizePt: 14,
      paddingPx: 10,
      background: "transparent",
    });
  } catch (error) {
    throw new Error(`${testCase.name} canonical SVG failed: ${renderSource}`, { cause: error });
  }
  assert.ok(svg.width > 0 && svg.height > 0, `${testCase.name} SVG`);
  let mathMl: string;
  try {
    mathMl = latexToMathMl(renderSource, true);
  } catch (error) {
    throw new Error(`${testCase.name} canonical MathML failed: ${renderSource}`, { cause: error });
  }
  assert.match(mathMl, /^<math\b/, `${testCase.name} MathML`);
}

const numberingOriginalLines = [
  { id: "numbering-line", latex: String.raw`e^{i\pi}+1=0` },
];
const numberingEditorLines = [
  {
    id: "numbering-line",
    latex: String.raw`\mathrm{e}^{\mathrm{i}\pi}+1=0`,
  },
];
assert.deepEqual(
  canonicalOfficeFingerprintLines(numberingOriginalLines),
  canonicalOfficeFingerprintLines(numberingEditorLines),
  "Office dirty canonicalization should continue treating MathEditor's automatic upright e/i rewrite as equivalent",
);
const numberingCanonicalFingerprint = JSON.stringify({
  lines: canonicalOfficeFingerprintLines(numberingOriginalLines),
  numbered: false,
});
const numberingOnly = isWordOmmlNumberingOnlyEdit({
  mode: "edit",
  objectMode: "wordOmml",
  displayMode: "block",
  numbered: true,
  originalNumbered: false,
  originalFingerprint: numberingCanonicalFingerprint,
  fingerprintAtOriginalNumbering: numberingCanonicalFingerprint,
});
assert.equal(numberingOnly, true, "Word OMML checkbox-only change must be recognized as numbering-only");
const numberingPersistedLines = persistedOfficeLines(
  numberingOnly,
  numberingOriginalLines,
  numberingEditorLines,
);
assert.deepEqual(
  numberingPersistedLines,
  numberingOriginalLines,
  "Numbering-only Word OMML edit must persist the original LaTeX instead of MathEditor's automatic e/i rewrite",
);
const numberingRenderSource = serializeFormulaEditorRenderDocument({
  lines: numberingPersistedLines,
  codeFormat: "raw",
});
const numberingMathMl = latexToMathMl(numberingRenderSource, true);
assert.match(numberingMathMl, /<mi>e<\/mi>/, "Numbering-only MathML should keep e as the original math variable");
assert.match(numberingMathMl, /<mi>i<\/mi>/, "Numbering-only MathML should keep i as the original math variable");
assert.doesNotMatch(
  numberingMathMl,
  /mathvariant=["']normal["']>\s*[ei]\s*</,
  "Numbering-only MathML must not silently persist MathEditor's upright e/i normalization",
);
assert.deepEqual(
  persistedOfficeLines(false, numberingOriginalLines, numberingEditorLines),
  numberingEditorLines,
  "A real content edit must still persist the current editor lines",
);

const equation = normalize(String.raw`\begin{equation}E=mc^2\end{equation}`);
assert.equal(equation.codeFormat, "equation");
assert.deepEqual(equation.lines, [{ id: "original-line-id", latex: "E=mc^2" }]);

const displayMath = normalize(
  String.raw`\begin{displaymath}x^2+y^2=z^2\end{displaymath}`,
);
assert.equal(displayMath.codeFormat, "equation-star");
assert.equal(displayMath.lines[0].latex, "x^2+y^2=z^2");

const alignat = normalize(String.raw`\begin{alignat}{2}
a&=b &\quad c&=d \\
e&=f &\quad g&=h
\end{alignat}`);
assert.equal(alignat.codeFormat, "align");
assert.deepEqual(
  alignat.lines.map((line) => line.latex),
  [
    `a${VISUALTEX_ALIGNMENT_MARKER_LATEX}=b ${VISUALTEX_ALIGNMENT_MARKER_LATEX}\\quad c${VISUALTEX_ALIGNMENT_MARKER_LATEX}=d`,
    `e${VISUALTEX_ALIGNMENT_MARKER_LATEX}=f ${VISUALTEX_ALIGNMENT_MARKER_LATEX}\\quad g${VISUALTEX_ALIGNMENT_MARKER_LATEX}=h`,
  ],
);

const equationWithAlignedSource = String.raw`\begin{equation}
\begin{aligned}
f^{*}(\mathbf{x})
&=
\frac{1}{p(\mathbf{x})}
\int t\,p(\mathbf{x},t)\,\mathrm{d}t \\
&=
\int t\,p(t\mid\mathbf{x})\,\mathrm{d}t
=
\mathbb{E}_{t}[t\mid\mathbf{x}]
\end{aligned}
\end{equation}`;
const equationWithAligned = normalize(equationWithAlignedSource);
assert.equal(equationWithAligned.codeFormat, "equation");
assert.equal(equationWithAligned.lines.length, 1);
assert.ok(equationWithAligned.lines[0].latex.includes("\\begin{aligned}"));
const equationWithAlignedSecondPass = normalizeFormulaEditorDocument(
  equationWithAligned.lines,
  equationWithAligned.codeFormat,
);
assert.deepEqual(equationWithAlignedSecondPass, equationWithAligned);
const equationWithAlignedCanonical = serializeFormulaEditorDocument(equationWithAligned);
assert.ok(equationWithAlignedCanonical.startsWith("\\begin{equation}\n"));
assert.ok(equationWithAlignedCanonical.includes("\\begin{aligned}"));
assert.ok(equationWithAlignedCanonical.includes("\\\\&="));
const equationWithAlignedSvg = latexToSvg(
  serializeFormulaEditorRenderDocument(equationWithAligned), {
  displayMode: true,
  fontSizePt: 14,
  paddingPx: 10,
  background: "transparent",
  },
);
assert.ok(equationWithAlignedSvg.width > 240);
assert.ok(equationWithAlignedSvg.height < 160);

for (const environment of ["equation", "equation*"] as const) {
  const formattedSource = String.raw`\begin{${environment}}
u(x,y)=\sum_{n=1}^{+\infty}c_n,\qquad
f(x,y)=\sum_{m=1}^{+\infty}d_m.
\end{${environment}}`;
  const normalized = normalize(formattedSource);
  assert.equal(normalized.lines.length, 1);
  assert.ok(!normalized.lines[0].latex.includes("\n"));
  assert.ok(normalized.lines[0].latex.includes("\\qquad f(x,y)"));
}

const alreadyNormalized = normalizeFormulaEditorDocument(
  [
    { id: "line-a", latex: "a=b" },
    { id: "line-b", latex: "c=d" },
  ],
  "align-star",
);
assert.equal(alreadyNormalized.codeFormat, "align-star");
assert.deepEqual(alreadyNormalized.lines, [
  { id: "line-a", latex: "a=b" },
  { id: "line-b", latex: "c=d" },
]);

const embeddedEnvironment = normalize(
  String.raw`prefix \begin{align}a&=b\end{align} suffix`,
);
assert.equal(embeddedEnvironment.codeFormat, "raw");
assert.equal(embeddedEnvironment.lines.length, 1);

const placeholder = String.raw`\frac{\placeholder{}}{\placeholder{}}`;
assert.equal(isIncompleteLatexDraft(placeholder), true);
assert.throws(
  () => latexToSvg(placeholder, { displayMode: false, fontSizePt: 14, paddingPx: 1 }),
  /placeholder/i,
);
assert.doesNotThrow(() =>
  latexToSvg(String.raw`\frac{a}{b}`, {
    displayMode: false,
    fontSizePt: 14,
    paddingPx: 1,
  }),
);

console.log("Windows Office canonical formula-document regression passed");
