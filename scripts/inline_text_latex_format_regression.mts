import assert from "node:assert/strict";
import {
  formatLatex,
  getLatexCodeFormatDefinition,
  parseLatexSourceDraft,
} from "../src/clipboard/LatexCopyService.ts";

const formatId = "inline-text-double-dollar" as const;
const definition = getLatexCodeFormatDefinition(formatId);
assert.equal(definition.hint, "文字$x^2$文字");
assert.match(definition.descriptionZh, /公式片段使用 \$\.\.\.\$/);

const formula = String.raw`\text{速度}v=\frac{s}{t}+x^{\text{中文}}+y_{\text{下标}}\text{结束}`;
const source = formatLatex(formula, formatId);
assert.equal(
  source,
  String.raw`速度$v=\frac{s}{t}+x^{\text{中文}}+y_{\text{下标}}$结束`,
);
assert.doesNotMatch(source, /\$\$/);
const parsed = parseLatexSourceDraft(source, formatId);
assert.equal(parsed.valid, true);
assert.deepEqual(parsed.values, [formula]);

const legacySource = String.raw`速度$$v=\frac{s}{t}+x^{\text{中文}}+y_{\text{下标}}$$结束`;
const legacyParsed = parseLatexSourceDraft(legacySource, formatId);
assert.equal(legacyParsed.valid, true);
assert.deepEqual(legacyParsed.values, [formula]);


console.log("macOS inline text LaTeX single-dollar format regression passed.");

for (const math of [
  String.raw`\left(a\text{ 和 }b\right)`,
  String.raw`\left\{\left(x\text{ 或 }y\right)\middle|z\right.`,
  String.raw`x_\text{下标}+\sqrt\text{中文}`,
  String.raw`\frac\text{甲}\text{乙}`,
  String.raw`\begin {cases}a&\text{如果 }x>0\\b&\text{否则}\end {cases}`,
]) {
  const formula = String.raw`\text{前}${math}\text{后}`;
  const source = formatLatex(formula, formatId);
  assert.equal(source, `前$${math}$后`, `Do not split a math construct: ${formula}`);
  assert.deepEqual(parseLatexSourceDraft(source, formatId).values, [formula]);
}
