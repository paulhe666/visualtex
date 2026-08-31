import assert from "node:assert/strict";
import { latexToSvg } from "../src/export/runtime.ts";

const result = latexToSvg(
  String.raw`\arccos x+\operatorname{rank}(A)+\arctan y`,
  {
    displayMode: false,
    formulaLetterFont: "times",
    paddingPx: 2,
    background: "transparent",
  },
);

const replacementTags = [
  ...result.svg.matchAll(
    /<text\b[^>]*data-visualtex-output-letter-font="times"[^>]*>/g,
  ),
].map((match) => match[0]);

assert.ok(
  replacementTags.length >= 3,
  `expected several system-font SVG glyph replacements, received ${replacementTags.length}`,
);

const transforms = replacementTags
  .map((tag) => tag.match(/\btransform="([^"]+)"/)?.[1] ?? "")
  .filter(Boolean);
assert.ok(
  transforms.every((transform) => transform.endsWith("scale(1,-1)")),
  `every replacement must retain the MathJax y-axis flip: ${JSON.stringify(transforms)}`,
);
const positioningSignatures = replacementTags.map((tag) => ({
  transform: tag.match(/\btransform="([^"]+)"/)?.[1] ?? "",
  x: tag.match(/\bx="([^"]+)"/)?.[1] ?? "",
  y: tag.match(/\by="([^"]+)"/)?.[1] ?? "",
}));
assert.ok(
  positioningSignatures.some(
    ({ transform, x, y }) => /translate\(/.test(transform) || Boolean(x) || Boolean(y),
  ),
  `system-font glyphs lost every local positioning attribute: ${JSON.stringify(positioningSignatures)}`,
);
assert.ok(
  new Set(positioningSignatures.map((value) => JSON.stringify(value))).size > 1,
  `multi-letter operator glyphs collapsed onto one origin: ${JSON.stringify(positioningSignatures)}`,
);
assert.ok(result.baseline >= 0 && result.baseline <= result.height);

console.log("SVG system-font local-transform regression passed");
