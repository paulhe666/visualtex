#!/usr/bin/env -S npx tsx

import assert from "node:assert/strict";
import { parseLatexSourceDraft } from "../src/clipboard/LatexCopyService.ts";
import { convertVisualTexLatexToMarkup } from "../src/editor/mathLiveIntegralCompatibility.ts";
import { compatibilityRequiredArgumentCounts } from "../src/autocomplete/compatibilityCommands.ts";
import { latexCompletions } from "codemirror-lang-latex";
import {
  createFormulaLine,
  normalizeFormulaLines,
  useEditorStore,
} from "../src/stores/editorStore.ts";

for (const source of ["\\", String.raw`\begin`, String.raw`\begin{}`]) {
  const parsed = parseLatexSourceDraft(source, "raw");
  assert.equal(parsed.valid, false, `${source} must remain an incomplete source draft`);
  assert.deepEqual(parsed.values, [source]);
  assert.equal(parsed.error, "incomplete-environment-command");
}

const structuralDrafts = [
  String.raw`\begin{matrix`,
  String.raw`\begin{matrix}`,
  String.raw`\begin{pmatrix}`,
  String.raw`\begin{bmatrix}`,
  String.raw`\begin{Bmatrix}`,
  String.raw`\begin{vmatrix}`,
  String.raw`\begin{Vmatrix}`,
  String.raw`\begin{smallmatrix}`,
  String.raw`\begin{cases}`,
  String.raw`\begin{align}`,
  String.raw`\begin{align*}`,
  String.raw`\begin{gather}`,
  String.raw`\begin{gather*}`,
  String.raw`\begin{multline}`,
  String.raw`\begin{multline*}`,
  String.raw`\begin{equation}`,
  String.raw`\begin{equation*}`,
  String.raw`\begin{aligned}`,
  String.raw`\begin{gathered}`,
  String.raw`\begin{split}`,
  String.raw`\begin{array}`,
  String.raw`\begin{array}{`,
  String.raw`\begin{array}{c`,
  String.raw`\begin{matrix}\frac{`,
  String.raw`\left`,
  String.raw`\left(`,
  String.raw`\left[`,
  String.raw`\left\langle`,
  String.raw`\frac`,
  String.raw`\frac{a}`,
  String.raw`\dfrac{a}`,
  String.raw`\tfrac`,
  String.raw`\cfrac`,
  String.raw`\binom`,
  String.raw`\sqrt`,
  String.raw`\sqrt[`,
  String.raw`\sqrt[3`,
  String.raw`\overbrace`,
  String.raw`\underbrace`,
  String.raw`\overset`,
  String.raw`\underset`,
  String.raw`\overline`,
  String.raw`\underline`,
  String.raw`\widehat`,
  String.raw`\widetilde`,
  String.raw`\vec`,
  String.raw`\dot`,
  String.raw`\ddot`,
  String.raw`\overrightarrow`,
  String.raw`\underleftarrow`,
  String.raw`\operatorname`,
  String.raw`\text`,
  String.raw`x_`,
  String.raw`x^`,
  String.raw`\sum_`,
  String.raw`\sum^`,
  String.raw`\prod_`,
  String.raw`\int^`,
  String.raw`\lim_`,
];

for (const environment of latexCompletions.environments) {
  const source = `\\begin{${environment}}\n  \n\\end{${environment}}`;
  const created = createFormulaLine(source, `auto-close-${environment}`);
  assert.equal(
    created.latex,
    source,
    `createFormulaLine truncated auto-closed ${environment} environment`,
  );
  const normalized = normalizeFormulaLines([created]);
  assert.equal(normalized.length, 1);
  assert.equal(
    normalized[0]?.latex,
    source,
    `normalizeFormulaLines truncated auto-closed ${environment} environment`,
  );
}

const historyEnvironment = String.raw`\begin{matrix}` + "\n  \n" + String.raw`\end{matrix}`;
const previousStore = useEditorStore.getState();
useEditorStore.setState({
  lines: [createFormulaLine(historyEnvironment, "history-multiline")],
  activeLineId: "history-multiline",
  history: [],
});
useEditorStore.getState().addHistory();
const historyEntry = useEditorStore.getState().history[0];
assert.equal(historyEntry?.lines?.length, 1);
assert.equal(
  historyEntry?.lines?.[0]?.latex,
  historyEnvironment,
  "formula history must preserve one auto-closed multiline environment as one logical FormulaLine",
);
useEditorStore.setState({
  lines: previousStore.lines,
  activeLineId: previousStore.activeLineId,
  history: previousStore.history,
});

for (const source of structuralDrafts) {
  const parsed = parseLatexSourceDraft(source, "raw");
  assert.equal(parsed.valid, false, `draft unexpectedly became valid: ${source}`);
  assert.deepEqual(
    parsed.values,
    [source],
    `preview generation must never rewrite the source value: ${source}`,
  );
  assert.equal(
    parsed.previewValues?.length,
    1,
    `draft has no render-safe preview: ${source} (${parsed.error})`,
  );
  const preview = parsed.previewValues?.[0] ?? "";
  assert.notEqual(preview, source, `draft preview did not complete structure: ${source}`);
  const markup = convertVisualTexLatexToMarkup(preview, { defaultMode: "math" });
  assert.ok(markup.length > 0, `draft preview markup is empty: ${source} -> ${preview}`);
  assert.doesNotMatch(
    markup,
    /ML__error/,
    `draft preview still contains a MathLive error: ${source} -> ${preview}`,
  );
}

for (const [command, argumentCount] of compatibilityRequiredArgumentCounts) {
  const source = String.fromCharCode(92) + command;
  const parsed = parseLatexSourceDraft(source, "raw");
  assert.equal(
    parsed.valid,
    false,
    `compatibility command without ${argumentCount} required argument(s) unexpectedly became valid: ${source}`,
  );
  assert.equal(
    parsed.previewValues?.length,
    1,
    `compatibility command has no render-safe draft preview: ${source} (${parsed.error})`,
  );
  const preview = parsed.previewValues?.[0] ?? "";
  const markup = convertVisualTexLatexToMarkup(preview, { defaultMode: "math" });
  assert.ok(markup.length > 0, `compatibility draft preview markup is empty: ${source}`);
  assert.doesNotMatch(markup, /ML__error/, `compatibility draft preview errored: ${source} -> ${preview}`);
}

const matrixDraft = parseLatexSourceDraft(String.raw`\begin{matrix}`, "raw");
assert.match(matrixDraft.previewValues?.[0] ?? "", /\\placeholder\{\}/);
assert.match(matrixDraft.previewValues?.[0] ?? "", /\\end\{matrix\}$/);

const arrayDraft = parseLatexSourceDraft(String.raw`\begin{array}`, "raw");
assert.match(arrayDraft.previewValues?.[0] ?? "", /\\begin\{array\}\{c\}/);
assert.match(arrayDraft.previewValues?.[0] ?? "", /\\end\{array\}$/);

const leftDraft = parseLatexSourceDraft(String.raw`\left(`, "raw");
assert.match(leftDraft.previewValues?.[0] ?? "", /\\right\.$/);

const fractionDraft = parseLatexSourceDraft(String.raw`\frac`, "raw");
assert.equal(
  fractionDraft.previewValues?.[0],
  String.raw`\frac{\placeholder{}}{\placeholder{}}`,
);

const validStructures = [
  String.raw`[0,1)`,
  String.raw`(0,1]`,
  String.raw`\left[0,1\right)`,
  String.raw`\left(0,1\right]`,
  String.raw`\begin{matrix}a&b\\c&d\end{matrix}`,
  String.raw`\begin{pmatrix}a&b\\c&d\end{pmatrix}`,
  String.raw`\begin{bmatrix}a&b\\c&d\end{bmatrix}`,
  String.raw`\begin{Bmatrix}a&b\\c&d\end{Bmatrix}`,
  String.raw`\begin{vmatrix}a&b\\c&d\end{vmatrix}`,
  String.raw`\begin{Vmatrix}a&b\\c&d\end{Vmatrix}`,
  String.raw`\begin{smallmatrix}a&b\\c&d\end{smallmatrix}`,
  String.raw`\begin{cases}x^2&x>0\\0&x=0\end{cases}`,
  String.raw`\begin{align}a&=b\\c&=d\end{align}`,
  String.raw`\begin{align*}a&=b\\c&=d\end{align*}`,
  String.raw`\begin{gather}a+b\\c+d\end{gather}`,
  String.raw`\begin{gather*}a+b\\c+d\end{gather*}`,
  String.raw`\begin{multline}a+b+c\\d+e+f\end{multline}`,
  String.raw`\begin{multline*}a+b+c\\d+e+f\end{multline*}`,
  String.raw`\begin{equation}a=b\end{equation}`,
  String.raw`\begin{equation*}a=b\end{equation*}`,
  String.raw`\begin{aligned}a&=b\\c&=d\end{aligned}`,
  String.raw`\begin{gathered}a+b\\c+d\end{gathered}`,
  String.raw`\begin{split}a&=b\\c&=d\end{split}`,
  String.raw`\begin{array}{cc}a&b\\c&d\end{array}`,
  String.raw`\left(\frac{a}{b}\right)`,
  String.raw`\overbrace{\frac{a}{b}+c}^{n}`,
  String.raw`\underbrace{a+b+c}_{n}`,
  String.raw`\sum_{i=1}^{n}\frac{1}{i}`,
  String.raw`\prod_{i=1}^{n}x_i`,
  String.raw`\int_{0}^{1}\frac{1}{1+x^2}\,\mathrm{d}x`,
  String.raw`\lim_{x\to 0}\frac{\sin x}{x}`,
];

for (const source of validStructures) {
  const parsed = parseLatexSourceDraft(source, "raw");
  assert.equal(
    parsed.valid,
    true,
    `valid structural source was rejected: ${source} (${parsed.error})`,
  );
  const markup = convertVisualTexLatexToMarkup(source, { defaultMode: "math" });
  assert.ok(markup.length > 0, `valid structural markup is empty: ${source}`);
  assert.doesNotMatch(markup, /ML__error/, `valid structural markup errored: ${source}`);
}

for (const format of [
  "display-dollar",
  "display-bracket",
  "inline-paren",
  "inline-dollar",
  "equation",
  "equation-star",
] as const) {
  const parsed = parseLatexSourceDraft(String.raw`\begin{matrix}`, format);
  assert.equal(parsed.valid, false, `direct matrix draft unexpectedly valid in ${format}`);
  assert.equal(
    parsed.previewValues?.length,
    1,
    `direct matrix draft has no preview in ${format}: ${parsed.error}`,
  );
  assert.match(parsed.previewValues?.[0] ?? "", /\\end\{matrix\}$/);
}

const unknown = parseLatexSourceDraft(String.raw`\definitelyunknown`, "raw");
assert.equal(unknown.valid, false);
assert.equal(
  unknown.previewValues,
  undefined,
  "unknown commands must not be disguised as a valid structural preview",
);

console.log("Source structural draft regression PASS");
