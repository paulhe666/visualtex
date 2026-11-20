import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { webcrypto } from "node:crypto";
import { existsSync, mkdirSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  parseDocumentImport,
  type DocumentImportRun,
} from "../src/office/documentImport/documentImportParser.ts";

if (!globalThis.crypto) {
  Object.defineProperty(globalThis, "crypto", { value: webcrypto });
}

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const windowsRoot = path.resolve(scriptDirectory, "..");
const fixtureDirectory = path.join(scriptDirectory, "fixtures");
const fixturePath = path.join(fixtureDirectory, "document-import-common-coverage.tex");
const outputDirectory = path.join(windowsRoot, "build-logs");
const outputPdf = path.join(outputDirectory, "document-import-common-coverage.pdf");

mkdirSync(outputDirectory, { recursive: true });

const compiler = process.env.VISUALTEX_XELATEX || "xelatex.exe";
const compile = spawnSync(
  compiler,
  [
    "-interaction=nonstopmode",
    "-halt-on-error",
    "-file-line-error",
    `-output-directory=${outputDirectory}`,
    "-jobname=document-import-common-coverage",
    path.basename(fixturePath),
  ],
  {
    cwd: fixtureDirectory,
    encoding: "utf8",
    windowsHide: true,
  },
);

if (compile.error) {
  throw new Error(`Unable to execute ${compiler}: ${compile.error.message}`);
}
assert.equal(
  compile.status,
  0,
  `XeLaTeX rejected the acceptance fixture:\n${compile.stdout}\n${compile.stderr}`,
);
assert.ok(existsSync(outputPdf), `XeLaTeX did not produce ${outputPdf}`);

const source = readFileSync(fixturePath, "utf8");
const parsed = parseDocumentImport(source, "latex");
const runs = parsed.blocks.flatMap((block) => block.runs);
const textRuns = runs.filter(
  (run): run is Extract<DocumentImportRun, { kind: "text" }> => run.kind === "text",
);
const formulaRuns = runs.filter(
  (run): run is Extract<DocumentImportRun, { kind: "formula" }> => run.kind === "formula",
);
const visibleText = textRuns.map((run) => run.text).join(" ");

const structureCounts = {
  headings: parsed.blocks.filter((block) => block.kind === "heading").length,
  bullets: parsed.blocks.filter((block) => block.kind === "bullet").length,
  numbered: parsed.blocks.filter((block) => block.kind === "numbered").length,
  quotes: parsed.blocks.filter((block) => block.kind === "quote").length,
  code: parsed.blocks.filter((block) => block.kind === "code").length,
  formulas: formulaRuns.length,
};
console.log("Compiled fixture parsed structure:", JSON.stringify(structureCounts));

assert.ok(structureCounts.headings >= 7);
assert.ok(structureCounts.bullets >= 4);
assert.ok(structureCounts.numbered >= 2);
assert.ok(structureCounts.quotes >= 2);
assert.equal(structureCounts.code, 2);
assert.ok(formulaRuns.length >= 17, `Expected at least 17 formulas, got ${formulaRuns.length}`);
assert.ok(formulaRuns.some((run) => run.latex.includes("\\mathbb{R}")));
assert.ok(formulaRuns.some((run) => run.latex.includes("begin{cases}")));
assert.ok(formulaRuns.some((run) => run.latex.includes("begin{aligned}")));
assert.ok(formulaRuns.some((run) => run.latex === "u+v"));

for (const expected of [
  "VisualTeX Document Import Coverage",
  "bold",
  "italic",
  "underline",
  "strike",
  "a_b%$",
  "Example",
  "https://example.com",
  "left and right",
  "A description-list definition",
  "Name",
  "Value",
  "Alpha",
  "Beta",
  "图片：example-image-a",
  "A figure placeholder test",
  "Coverage theorem",
  "外部 LaTeX 文件：document_import_input_fragment",
  "外部 LaTeX 文件：document_import_include_fragment",
  "Knuth84",
  "The TeXbook",
]) {
  assert.ok(visibleText.includes(expected), `Missing visible content: ${expected}`);
}

assert.ok(textRuns.some((run) => run.underline && run.text.includes("underline")));
assert.ok(textRuns.some((run) => run.strike && run.text.includes("strike")));
assert.ok(textRuns.some((run) => run.code && run.text.includes("a_b%$")));
assert.ok(parsed.warnings.some((warning) => warning.includes("tabular")));
assert.ok(parsed.warnings.some((warning) => warning.includes("custombox")));
assert.ok(parsed.warnings.some((warning) => warning.includes("自定义宏")));

for (const leaked of [
  "\\documentclass",
  "\\usepackage",
  "\\begin{",
  "\\end{",
  "\\includegraphics",
  "\\input{",
  "\\include{",
  "\\bibitem",
]) {
  assert.ok(!visibleText.includes(leaked), `LaTeX control structure leaked into Word text: ${leaked}`);
}

console.log(
  "Compiled LaTeX document import acceptance passed:",
  JSON.stringify({
    pdf: outputPdf,
    blocks: parsed.blocks.length,
    formulas: parsed.formulaCount,
    inlineFormulas: parsed.inlineFormulaCount,
    displayFormulas: parsed.displayFormulaCount,
    warnings: parsed.warnings.length,
  }),
);
