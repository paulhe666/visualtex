import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { webcrypto } from "node:crypto";
import { readFileSync } from "node:fs";
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
const fixturePath = path.join(
  scriptDirectory,
  "fixtures",
  "document-import-common-coverage.md",
);
const python = process.env.VISUALTEX_PYTHON || "python";
const referenceCompile = spawnSync(
  python,
  [
    "-c",
    [
      "from pathlib import Path",
      "import markdown, sys",
      "source = Path(sys.argv[1]).read_text(encoding='utf-8')",
      "html = markdown.markdown(source, extensions=['extra', 'sane_lists', 'nl2br'])",
      "required = ['<h1', '<h2', '<blockquote>', '<table>', '<pre><code', '<a href=', 'footnote']",
      "missing = [item for item in required if item not in html]",
      "assert not missing, f'Markdown reference compiler missed: {missing}'",
      "print(len(html))",
    ].join("; "),
    fixturePath,
  ],
  { encoding: "utf8", windowsHide: true },
);

if (referenceCompile.error) {
  throw new Error(`Unable to execute ${python}: ${referenceCompile.error.message}`);
}
assert.equal(
  referenceCompile.status,
  0,
  `Python Markdown rejected the acceptance fixture:\n${referenceCompile.stdout}\n${referenceCompile.stderr}`,
);

const source = readFileSync(fixturePath, "utf8");
const parsed = parseDocumentImport(source, "markdown");
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

assert.ok(structureCounts.headings >= 2);
assert.ok(structureCounts.bullets >= 4);
assert.ok(structureCounts.numbered >= 3);
assert.ok(structureCounts.quotes >= 1);
assert.ok(structureCounts.code >= 4);
assert.ok(structureCounts.formulas >= 9);

for (const expected of [
  "VisualTeX Markdown Import",
  "Setext level two",
  "bold",
  "strong",
  "italic",
  "bold italic",
  "strike",
  "inline_code()",
  "code with ` tick",
  "This sentence starts after the hard break",
  "Example",
  "https://example.com",
  "OpenAI",
  "https://openai.com",
  "https://example.org",
  "图片：Diagram（diagram.png）",
  "A & B, α < x >",
  "visible content",
  "☒ Completed task",
  "☐ Pending task",
  "Name",
  "Value",
  "Formula",
  "Alpha",
  "Beta",
  "const formula = \"$not_math$\";",
  "print(\"tilde fence\")",
  "indented_code = True",
  "注 note：",
  "This is the footnote body",
]) {
  assert.ok(visibleText.includes(expected), `Missing visible content: ${expected}`);
}

assert.ok(textRuns.some((run) => run.strike && run.text === "strike"));
assert.ok(textRuns.some((run) => run.code && run.text === "inline_code()"));
assert.ok(formulaRuns.some((run) => run.latex === "E=mc^2"));
assert.ok(formulaRuns.some((run) => run.latex.includes("\\int_0^1")));
assert.ok(formulaRuns.some((run) => run.latex.includes("begin{aligned}")));
assert.ok(formulaRuns.some((run) => run.latex === "n=1"));
assert.ok(parsed.warnings.some((warning) => warning.includes("引用式链接")));
assert.ok(parsed.warnings.some((warning) => warning.includes("脚注")));

for (const leaked of [
  "<!--",
  "-->",
  "<strong>",
  "</strong>",
  "```",
  "~~~",
  "[^note]",
  "|:-----|",
]) {
  assert.ok(!visibleText.includes(leaked), `Markdown syntax leaked into Word text: ${leaked}`);
}

console.log(
  "Compiled Markdown document import acceptance passed:",
  JSON.stringify({
    referenceHtmlCharacters: Number(referenceCompile.stdout.trim()),
    blocks: parsed.blocks.length,
    formulas: parsed.formulaCount,
    inlineFormulas: parsed.inlineFormulaCount,
    displayFormulas: parsed.displayFormulaCount,
    warnings: parsed.warnings.length,
    structureCounts,
  }),
);
