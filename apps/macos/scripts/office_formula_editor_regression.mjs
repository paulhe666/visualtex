import assert from "node:assert/strict";
import { DOMParser } from "@xmldom/xmldom";
import {
  normalizeFormulaEditorDocument,
  serializeFormulaEditorDocument,
} from "../src/office/shared/formulaEditorDocument.ts";
import {
  createFormulaMetadata,
  decodeFormulaMetadata,
  encodeFormulaMetadata,
  formulaMetadataFromXml,
  formulaMetadataToXml,
} from "../src/office/shared/formulaMetadata.ts";
import {
  renderOfficeFormulaArtifacts,
  tryRenderOfficeFormulaDraftArtifacts,
} from "../src/office/shared/formulaRenderArtifacts.ts";
import { latexToSvg } from "../src/export/latexToSvg.ts";
import { errorMessage } from "../src/runtime/errorMessage.ts";
import { VISUALTEX_ALIGNMENT_MARKER_LATEX } from "../src/editor/alignmentMarkers.ts";
import {
  registerCustomSymbol,
  replaceCustomSymbolLibrary,
} from "../src/math/customSymbolRegistry.ts";
import {
  clearsOfficeEditorActivation,
  isOfficeEditorActivation,
  isOfficeEditorClear,
  shouldAcceptOfficeEditorActivation,
} from "../src/office/dialog/officeEditorActivation.ts";

globalThis.DOMParser ??= DOMParser;
const domProbe = new DOMParser().parseFromString("<root/>", "application/xml");
const documentPrototype = Object.getPrototypeOf(domProbe);
const elementPrototype = Object.getPrototypeOf(domProbe.documentElement);
if (typeof documentPrototype.querySelector !== "function") {
  documentPrototype.querySelector = function querySelector(name) {
    return this.getElementsByTagName(name)?.item(0) ?? null;
  };
}
if (!("children" in elementPrototype)) {
  Object.defineProperty(elementPrototype, "children", {
    configurable: true,
    get() {
      return Array.from(this.childNodes ?? []).filter((node) => node.nodeType === 1);
    },
  });
}

replaceCustomSymbolLibrary({ version: 1, symbols: [] });
const customSymbolNow = Date.now();
registerCustomSymbol({
  id: "office-metadata-selfdefa",
  command: "selfdefa",
  name: "Office metadata custom symbol",
  role: "relation",
  limitsBehavior: "auto",
  metrics: { widthEm: 0.8, ascentEm: 0.64, descentEm: 0.1 },
  artwork: {
    shapes: [
      {
        kind: "circle",
        cx: 400,
        cy: 360,
        r: 245,
        fill: false,
        strokeWidth: 72,
      },
      {
        kind: "line",
        x1: 130,
        y1: 360,
        x2: 670,
        y2: 360,
        fill: false,
        strokeWidth: 72,
        lineCap: "round",
      },
      {
        kind: "path",
        operation: "erase",
        d: "M220 360L580 360",
        fill: false,
        strokeWidth: 110,
        lineCap: "round",
      },
    ],
  },
  ommlFallback: "\\approx",
  createdAt: customSymbolNow,
  updatedAt: customSymbolNow,
});
const customSymbolLines = [{ id: "office-custom-symbol-line", latex: "\\selfdefa" }];
const customSymbolRendered = renderOfficeFormulaArtifacts({
  lines: customSymbolLines,
  codeFormat: "raw",
  displayMode: "inline",
  host: "word",
  includeWordOmml: true,
});
const customSymbolMetadata = createFormulaMetadata({
  formulaId: "12345678-1234-4234-9234-123456789abc",
  title: "Custom symbol metadata round trip",
  lines: customSymbolLines,
  codeFormat: "raw",
  sourceLatex: customSymbolRendered.canonicalLatex,
  displayMode: "inline",
});
const customSymbolEncodedMetadata = encodeFormulaMetadata(customSymbolMetadata);
const customSymbolDecodedMetadata = decodeFormulaMetadata(customSymbolEncodedMetadata);
assert.ok(customSymbolDecodedMetadata);
assert.equal(customSymbolDecodedMetadata.latex, "\\selfdefa");
const customSymbolXml = formulaMetadataToXml(customSymbolMetadata);
const customSymbolXmlMetadata = formulaMetadataFromXml(customSymbolXml);
assert.ok(customSymbolXmlMetadata);
assert.equal(customSymbolXmlMetadata.latex, "\\selfdefa");
const reopenedCustomSymbolDocument = normalizeFormulaEditorDocument(
  customSymbolXmlMetadata.lines,
  customSymbolXmlMetadata.codeFormat,
);
const reopenedCustomSymbolRendered = renderOfficeFormulaArtifacts({
  lines: reopenedCustomSymbolDocument.lines,
  codeFormat: reopenedCustomSymbolDocument.codeFormat,
  displayMode: customSymbolXmlMetadata.displayMode,
  host: "word",
  includeWordOmml: true,
});
assert.equal(reopenedCustomSymbolRendered.canonicalLatex, "\\selfdefa");
assert.match(
  reopenedCustomSymbolRendered.svg.svg,
  /data-visualtex-custom-symbol="office-metadata-selfdefa"/,
);
assert.match(
  reopenedCustomSymbolRendered.svg.svg,
  /<mask\b[^>]*id="visualtex-custom-symbol-erase-office-metadata-selfdefa-/,
  "Word SVG must preserve custom-symbol transparent eraser masks after metadata reopen",
);
const powerpointErasedCustomSymbol = renderOfficeFormulaArtifacts({
  lines: reopenedCustomSymbolDocument.lines,
  codeFormat: reopenedCustomSymbolDocument.codeFormat,
  displayMode: customSymbolXmlMetadata.displayMode,
  host: "powerpoint",
  includeWordOmml: false,
});
assert.match(
  powerpointErasedCustomSymbol.svg.svg,
  /<mask\b[^>]*id="visualtex-custom-symbol-erase-office-metadata-selfdefa-/,
  "PowerPoint SVG must preserve custom-symbol transparent eraser masks",
);
assert.equal(powerpointErasedCustomSymbol.omml, null);
assert.match(reopenedCustomSymbolRendered.omml?.omml ?? "", /≈/);
assert.doesNotMatch(reopenedCustomSymbolRendered.omml?.omml ?? "", /selfdefa/);
replaceCustomSymbolLibrary({ version: 1, symbols: [] });

const warmActivation = {
  sessionId: "12345678-1234-4234-9234-123456789abc",
  host: "word",
  generation: 7,
  receivedEpochMs: 1_750_000_000_000,
};
assert.equal(isOfficeEditorActivation(warmActivation), true);
assert.equal(
  shouldAcceptOfficeEditorActivation(null, warmActivation, "word"),
  true,
);
assert.equal(
  shouldAcceptOfficeEditorActivation(
    warmActivation,
    { ...warmActivation, generation: 6 },
    "word",
  ),
  false,
  "an older WebView event must never restore a stale Session",
);
assert.equal(
  shouldAcceptOfficeEditorActivation(
    warmActivation,
    { ...warmActivation, generation: 8, host: "powerpoint" },
    "word",
  ),
  false,
  "a resident Word WebView must ignore PowerPoint activations",
);
const matchingClear = {
  sessionId: warmActivation.sessionId,
  generation: warmActivation.generation,
};
assert.equal(isOfficeEditorClear(matchingClear), true);
assert.equal(clearsOfficeEditorActivation(warmActivation, matchingClear), true);
assert.equal(isOfficeEditorClear(null), false);
assert.equal(isOfficeEditorClear({ ...matchingClear, generation: "7" }), false);
assert.equal(
  clearsOfficeEditorActivation(warmActivation, null),
  false,
  "a malformed native clear event must be ignored instead of crashing the resident editor",
);

function normalize(source, codeFormat = "raw") {
  return normalizeFormulaEditorDocument(
    [{ id: "original-line-id", latex: source }],
    codeFormat,
  );
}

const wrappedMultiColumnAlignedSource = String.raw`\[
\begin{aligned}
\langle p_1,p_0\rangle &\gets \operatorname{umul}(a,b)=ab
  && \text{Double word product}\\
p_0 &\gets \operatorname{umullo}(a,b)=(ab)\bmod\beta
  && \text{Low word}\\
p_1 &\gets \operatorname{umulhi}(a,b)=\left\lfloor\frac{ab}{\beta}\right\rfloor
  && \text{High word.}
\end{aligned}
\]`;
for (const source of [
  wrappedMultiColumnAlignedSource,
  `$$${wrappedMultiColumnAlignedSource.slice(2, -2)}$$`,
]) {
  const normalized = normalize(source);
  assert.equal(normalized.codeFormat, "aligned");
  assert.equal(normalized.lines.length, 3);
  assert.ok(
    normalized.lines.every(
      (line) =>
        (line.latex.match(new RegExp(VISUALTEX_ALIGNMENT_MARKER_LATEX.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g")) ?? [])
          .length === 3,
    ),
    "all three alignment separators in each row must survive display-wrapper import",
  );
  const rendered = renderOfficeFormulaArtifacts({
    lines: normalized.lines,
    codeFormat: normalized.codeFormat,
    displayMode: "block",
    host: "word",
    includeWordOmml: true,
  });
  assert.doesNotMatch(rendered.svg.svg, /data-mml-node="mtext"[^>]*>[\s\S]*?\\\[/);
  assert.doesNotMatch(rendered.svg.svg, /data-mml-node="mtext"[^>]*>[\s\S]*?\\\]/);
}

const placeholderDraftInput = {
  lines: [
    {
      id: "placeholder-draft",
      latex: String.raw`\frac{\placeholder{}}{\placeholder{}}`,
    },
  ],
  codeFormat: "raw",
  displayMode: "inline",
  host: "word",
};
assert.equal(
  tryRenderOfficeFormulaDraftArtifacts(placeholderDraftInput),
  null,
  "Office autosave must tolerate MathLive placeholder source without producing an invalid Word artifact",
);
assert.throws(
  () => renderOfficeFormulaArtifacts(placeholderDraftInput),
  /placeholder/,
  "explicit Office apply must remain strict while a placeholder is still present",
);
assert.ok(
  tryRenderOfficeFormulaDraftArtifacts({
    ...placeholderDraftInput,
    lines: [{ id: "completed-draft", latex: String.raw`\frac{a}{b}` }],
  }),
  "a completed formula must resume normal draft artifact generation",
);

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
  "source-formatting newlines inside one logical formula row must preserve adjacent TeX arguments without inserting parser-breaking spaces",
);
assert.doesNotThrow(() =>
  renderOfficeFormulaArtifacts({
    lines: sourceFormattedEquation.lines,
    codeFormat: sourceFormattedEquation.codeFormat,
    displayMode: "block",
    includeWordOmml: false,
  }),
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
];

for (const testCase of multilineCases) {
  const normalized = normalize(testCase.source);
  assert.equal(
    normalized.codeFormat,
    testCase.codeFormat,
    `${testCase.name} code format`,
  );
  assert.deepEqual(
    normalized.lines.map((line) => line.latex),
    testCase.lines,
    `${testCase.name} rows`,
  );
  assert.equal(
    normalized.lines[0].id,
    "original-line-id",
    `${testCase.name} must preserve the imported first-line UUID`,
  );
  assert.equal(
    new Set(normalized.lines.map((line) => line.id)).size,
    normalized.lines.length,
    `${testCase.name} row UUIDs must remain unique`,
  );
  const canonicalSource = serializeFormulaEditorDocument(normalized);
  const roundTrip = normalize(canonicalSource, normalized.codeFormat);
  assert.equal(
    serializeFormulaEditorDocument(roundTrip),
    canonicalSource,
    `${testCase.name} canonical source must be stable`,
  );
  assert.deepEqual(
    roundTrip.lines.map((line) => line.latex),
    testCase.lines,
    `${testCase.name} canonical source must preserve every row`,
  );
  const metadata = createFormulaMetadata({
    formulaId: "12345678-1234-4234-9234-123456789abc",
    title: testCase.name,
    lines: normalized.lines,
    codeFormat: normalized.codeFormat,
    sourceLatex: canonicalSource,
    displayMode: "block",
  });
  assert.equal(
    metadata.latex,
    canonicalSource,
    `${testCase.name} metadata must store the canonical serialized source`,
  );
  if (testCase.codeFormat === "align" || testCase.codeFormat === "align-star") {
    const rendered = renderOfficeFormulaArtifacts({
      lines: normalized.lines,
      codeFormat: normalized.codeFormat,
      displayMode: "block",
      includeWordOmml: false,
    });
    const wordRendered = renderOfficeFormulaArtifacts({
      lines: normalized.lines,
      codeFormat: normalized.codeFormat,
      displayMode: "block",
      host: "word",
      includeWordOmml: false,
    });
    const firstImportSvg = latexToSvg(canonicalSource, {
      displayMode: true,
      fontSizePt: 14,
      paddingPx: 10,
      background: "transparent",
    });
    const firstWordImportSvg = latexToSvg(canonicalSource, {
      displayMode: true,
      fontSizePt: 14,
      paddingPx: 2,
      background: "transparent",
      forceExplicitBlack: true,
    });
    assert.equal(
      rendered.canonicalLatex,
      canonicalSource,
      `${testCase.name} edit rendering must rebuild the complete environment`,
    );
    assert.equal(
      rendered.svg.svg.replace(/MJX-\d+-/g, "MJX-N-"),
      firstImportSvg.svg.replace(/MJX-\d+-/g, "MJX-N-"),
      `${testCase.name} first import and edit replacement must share the same SVG`,
    );
    assert.equal(rendered.svg.width, firstImportSvg.width);
    assert.equal(rendered.svg.height, firstImportSvg.height);
    assert.equal(rendered.svg.baseline, firstImportSvg.baseline);
    assert.equal(
      wordRendered.svg.svg.replace(/MJX-\d+-/g, "MJX-N-"),
      firstWordImportSvg.svg.replace(/MJX-\d+-/g, "MJX-N-"),
      `${testCase.name} Word rendering must use the tight 2 px display bounds`,
    );
    assert.equal(wordRendered.svg.width, firstWordImportSvg.width);
    assert.equal(wordRendered.svg.height, firstWordImportSvg.height);
    assert.equal(wordRendered.svg.baseline, firstWordImportSvg.baseline);
    assert.ok(
      /(?:fill|stroke)=["']#000000["']/i.test(wordRendered.svg.svg),
      `${testCase.name} Word SVG must contain explicit black formula paint`,
    );
    assert.ok(
      !/currentColor|var\(|(?:fill|stroke|color)\s*[:=]\s*["']?(?:inherit|white|#fff(?:fff)?)/i.test(
        wordRendered.svg.svg,
      ),
      `${testCase.name} Word SVG must not defer or whiten formula paint`,
    );
    assert.ok(
      wordRendered.svg.height < rendered.svg.height,
      `${testCase.name} Word bounds must be tighter than PowerPoint bounds`,
    );
  }
}

const equation = normalize(String.raw`\begin{equation}E=mc^2\end{equation}`);
assert.equal(equation.codeFormat, "equation");
assert.deepEqual(equation.lines, [
  { id: "original-line-id", latex: "E=mc^2" },
]);

for (const [environment, expectedCodeFormat] of [
  ["equation", "equation"],
  ["equation*", "equation-star"],
]) {
  const formattedSource = String.raw`\begin{${environment}}
u(x,y)=\sum_{n=1}^{+\infty}\sum_{m=1}^{+\infty}c_{nm}\sin\frac{n\pi}{a}x\sin\frac{m\pi}{b}y,\qquad
f(x,y)=\sum_{n=1}^{+\infty}\sum_{m=1}^{+\infty}d_{nm}\sin\frac{n\pi}{a}x\sin\frac{m\pi}{b}y.
\end{${environment}}`;
  const normalized = normalize(formattedSource);
  assert.equal(normalized.codeFormat, expectedCodeFormat);
  assert.equal(
    normalized.lines.length,
    1,
    `${environment} source-formatting newlines must remain one logical formula`,
  );
  assert.ok(!normalized.lines[0].latex.includes("\n"));
  assert.ok(normalized.lines[0].latex.includes("\\qquad f(x,y)"));
  const canonical = serializeFormulaEditorDocument(normalized);
  assert.equal(
    (canonical.match(new RegExp(`\\\\begin\\{${environment.replace("*", "\\*")}\\}`, "g")) ?? []).length,
    1,
    `${environment} serialization must create exactly one opening environment`,
  );
  assert.equal(
    (canonical.match(new RegExp(`\\\\end\\{${environment.replace("*", "\\*")}\\}`, "g")) ?? []).length,
    1,
    `${environment} serialization must create exactly one closing environment`,
  );
  assert.ok(!canonical.includes("&"));
  const rendered = renderOfficeFormulaArtifacts({
    lines: normalized.lines,
    codeFormat: normalized.codeFormat,
    displayMode: "block",
    includeWordOmml: false,
  });
  assert.equal(rendered.canonicalLatex, canonical);
  assert.ok(rendered.svg.width > 0 && rendered.svg.height > 0);
}

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
assert.equal(equationWithAlignedSecondPass.codeFormat, "equation");
assert.deepEqual(
  equationWithAlignedSecondPass.lines,
  equationWithAligned.lines,
  "a normalized outer equation must not be reclassified from its inner aligned environment",
);
const equationWithAlignedRendered = renderOfficeFormulaArtifacts({
  lines: equationWithAligned.lines,
  codeFormat: equationWithAligned.codeFormat,
  displayMode: "block",
  host: "word",
  includeWordOmml: false,
});
assert.equal(equationWithAlignedRendered.codeFormat, "equation");
assert.equal(equationWithAlignedRendered.lines.length, 1);
assert.ok(
  equationWithAlignedRendered.canonicalLatex.startsWith("\\begin{equation}\n"),
);
assert.ok(
  equationWithAlignedRendered.canonicalLatex.includes("\\begin{aligned}"),
);
assert.ok(equationWithAlignedRendered.canonicalLatex.includes("\\\\&="));
assert.ok(equationWithAlignedRendered.svg.width > 240);
assert.ok(
  equationWithAlignedRendered.svg.height < 130,
  "source-formatting newlines inside aligned must not become independent visual rows",
);

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

const alreadyNormalized = normalizeFormulaEditorDocument(
  [
    { id: "line-a", latex: "a=b" },
    { id: "line-b", latex: "c=d" },
  ],
  "align-star",
);
assert.equal(alreadyNormalized.codeFormat, "align-star");
assert.deepEqual(alreadyNormalized.lines, [
  { id: "line-a", latex: `a${VISUALTEX_ALIGNMENT_MARKER_LATEX}=b` },
  { id: "line-b", latex: `c${VISUALTEX_ALIGNMENT_MARKER_LATEX}=d` },
]);
const recoveredEditedAlignment = renderOfficeFormulaArtifacts({
  lines: [
    { id: "edited-line-a", latex: "1 = 22 + 333 + q" },
    { id: "edited-line-b", latex: "44444 = 55 + r" },
  ],
  codeFormat: "align",
  displayMode: "block",
  host: "word",
  includeWordOmml: true,
});
assert.ok(
  recoveredEditedAlignment.lines.every((line) =>
    line.latex.includes(VISUALTEX_ALIGNMENT_MARKER_LATEX),
  ),
  "whole-row align edits must recover their missing relationship markers",
);
assert.match(recoveredEditedAlignment.canonicalLatex, /1 &= 22 \+ 333 \+ q/);
assert.match(recoveredEditedAlignment.omml?.omml ?? "", /<m:eqArr>/);

const embeddedEnvironment = normalize(
  String.raw`prefix \begin{align}a&=b\end{align} suffix`,
);
assert.equal(embeddedEnvironment.codeFormat, "raw");
assert.equal(embeddedEnvironment.lines.length, 1);

assert.equal(errorMessage({ message: "direct message" }, "fallback"), "direct message");
assert.equal(
  errorMessage({ error: { description: "nested description" } }, "fallback"),
  "nested description",
);
assert.equal(
  errorMessage({ details: { code: 7400, host: "word" } }, "fallback"),
  JSON.stringify({ code: 7400, host: "word" }),
);
assert.equal(errorMessage({ status: 500 }, "fallback"), '{"status":500}');
assert.equal(errorMessage(42, "fallback"), "42");
assert.equal(errorMessage("[object Object]", "object fallback"), "object fallback");
assert.equal(
  errorMessage('{"error":{"message":"serialized failure"}}', "fallback"),
  "serialized failure",
);

const cyclic = {};
cyclic.cause = cyclic;
const cyclicMessage = errorMessage(cyclic, "cyclic fallback");
assert.equal(cyclicMessage, "cyclic fallback");
assert.ok(!cyclicMessage.includes("[object Object]"));

for (const reason of [
  { message: "message" },
  { error: "error" },
  { description: "description" },
  { details: "details" },
  { arbitrary: { value: true } },
]) {
  assert.ok(!errorMessage(reason, "fallback").includes("[object Object]"));
}

console.log("Office formula editor regression passed");
