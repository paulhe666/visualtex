import assert from "node:assert/strict";
import {
  formatFormulaLinesUniversal,
  parseUniversalLatexSourceDraft,
} from "../src/clipboard/LatexCopyService.ts";
import {
  DEFAULT_LATEX_FORMAT_PROFILE,
  normalizeLatexFormatProfile,
} from "../src/clipboard/latexFormatProfile.ts";
import type { FormulaLine } from "../src/types/formula.ts";
import { restoreLatexAlignmentMarkers } from "../src/editor/alignmentMarkers.ts";

const lines: FormulaLine[] = [
  {
    id: "inline",
    latex: String.raw`\text{前}x+1\text{后}`,
    mode: "inline",
    displayStyle: "default",
  },
  {
    id: "display-default",
    latex: String.raw`\frac{x}{y}=z`,
    mode: "display",
    displayStyle: "default",
  },
  {
    id: "display-bracket",
    latex: "E=mc^2",
    mode: "display",
    displayStyle: "bracket",
  },
  {
    id: "display-equation",
    latex: "a=b",
    mode: "display",
    displayStyle: "equation",
  },
  {
    id: "multiline",
    latex: String.raw`\begin{gathered}p=q\\r=s\end{gathered}`,
    mode: "display",
    displayStyle: "default",
  },
];

const defaultProfile = normalizeLatexFormatProfile(DEFAULT_LATEX_FORMAT_PROFILE);
const source = formatFormulaLinesUniversal(lines, defaultProfile);
assert.equal(
  source,
  [
    String.raw`前$x+1$后`,
    "$$\n\\frac{x}{y}=z\n$$",
    "\\[\nE=mc^2\n\\]",
    "\\begin{equation}\na=b\n\\end{equation}",
    "\\begin{gather*}\np=q \\\\\nr=s\n\\end{gather*}",
  ].join("\n\n"),
);

const parsed = parseUniversalLatexSourceDraft(source, defaultProfile);
assert.equal(parsed.valid, true);
assert.deepEqual(parsed.modes, [
  "inline",
  "display",
  "display",
  "display",
  "display",
]);
assert.deepEqual(parsed.displayStyles, [
  "default",
  "double-dollar",
  "bracket",
  "equation",
  "default",
]);
assert.equal(parsed.values?.[0], String.raw`\text{前}x+1\text{后}`);
assert.equal(parsed.values?.[1], String.raw`\frac{x}{y}=z`);
assert.equal(parsed.values?.[2], "E=mc^2");
assert.equal(parsed.values?.[3], "a=b");
assert.equal(
  parsed.values?.[4]?.replace(/\s+/g, ""),
  String.raw`\begin{gathered}p=q\\r=s\end{gathered}`.replace(/\s+/g, ""),
);

const parenProfile = normalizeLatexFormatProfile({
  ...defaultProfile,
  inlineWrapper: "paren",
  inlineTextPolicy: "outside-math",
  displayWrapper: "equation",
  numbered: true,
  multilineEnvironment: "align",
});
const parenSource = formatFormulaLinesUniversal(
  [
    {
      id: "inline",
      latex: String.raw`\text{速度}v\text{恒定}`,
      mode: "inline",
    },
    {
      id: "display",
      latex: "F=ma",
      mode: "display",
      displayStyle: "default",
    },
    {
      id: "aligned",
      latex: String.raw`\begin{aligned}a&=b\\c&=d\end{aligned}`,
      mode: "display",
      displayStyle: "default",
    },
  ],
  parenProfile,
);
assert.equal(
  parenSource,
  [
    String.raw`速度\(v\)恒定`,
    "\\begin{equation}\nF=ma\n\\end{equation}",
    "\\begin{align}\na&=b \\\\\nc&=d\n\\end{align}",
  ].join("\n\n"),
);

const parenParsed = parseUniversalLatexSourceDraft(parenSource, parenProfile);
assert.equal(parenParsed.valid, true);
assert.deepEqual(parenParsed.modes, ["inline", "display", "display"]);
assert.deepEqual(parenParsed.displayStyles, [
  "default",
  "equation",
  "default",
]);
assert.equal(
  restoreLatexAlignmentMarkers(parenParsed.values?.[2] ?? "").replace(/\s+/g, ""),
  String.raw`\begin{aligned}a&=b\\c&=d\end{aligned}`.replace(/\s+/g, ""),
);

const textCommandProfile = normalizeLatexFormatProfile({
  ...defaultProfile,
  inlineWrapper: "paren",
  inlineTextPolicy: "text-command",
});
assert.equal(
  formatFormulaLinesUniversal(
    [
      {
        id: "inline-text-command",
        latex: String.raw`\text{前}x\text{后}`,
        mode: "inline",
      },
    ],
    textCommandProfile,
  ),
  String.raw`\(\text{前}x\text{后}\)`,
);

console.log("universal LaTeX format profile regression: PASS");
