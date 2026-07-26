import assert from "node:assert/strict";
import { buildMarkdownDocument } from "../src/export/markdownExport.ts";

assert.equal(
  buildMarkdownDocument("高斯积分", [
    "\\int_{-\\infty}^{\\infty} e^{-x^2}\\,\\mathrm{d}x = \\sqrt{\\pi}",
    "",
    "E = mc^2",
  ]),
  "# 高斯积分\n\n$$\n\\int_{-\\infty}^{\\infty} e^{-x^2}\\,\\mathrm{d}x = \\sqrt{\\pi}\n$$\n\n$$\nE = mc^2\n$$\n",
);

assert.equal(
  buildMarkdownDocument("Line one\nLine two", ["x+y"]),
  "# Line one Line two\n\n$$\nx+y\n$$\n",
);

assert.equal(
  buildMarkdownDocument("Ignored", ["a=b"], { includeTitle: false }),
  "$$\na=b\n$$\n",
);

assert.equal(buildMarkdownDocument("", ["", "   "]), "");

console.log("Markdown export smoke test passed");
