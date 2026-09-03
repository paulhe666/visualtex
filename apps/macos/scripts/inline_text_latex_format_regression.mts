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
