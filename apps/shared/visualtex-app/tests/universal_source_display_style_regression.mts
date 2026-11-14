// The source panel writes a display row that follows the LaTeX format profile with the profile's own wrapper.
// Reading the source back must keep such a row on "follow the profile", so editing one row (or editing and then
// undoing) leaves the other rows' display wrapper alone, and a later profile change still applies to them. Only a
// block written with a different wrapper becomes that row's own choice.
import assert from "node:assert/strict";
import {
  formatFormulaLinesUniversal,
  parseUniversalLatexSourceDraft,
} from "../src/clipboard/LatexCopyService.ts";
import { DEFAULT_LATEX_FORMAT_PROFILE } from "../src/clipboard/latexFormatProfile.ts";
import { reconcileFormulaLines } from "../src/history/documentHistory.ts";
import type { FormulaLine, LatexFormatProfile } from "../src/types/formula.ts";

const profile = (overrides: Partial<LatexFormatProfile>): LatexFormatProfile => ({
  ...DEFAULT_LATEX_FORMAT_PROFILE,
  ...overrides,
});
const styles = (source: string, used: LatexFormatProfile) => {
  const parsed = parseUniversalLatexSourceDraft(source, used);
  assert.equal(parsed.valid, true, `${source}: ${parsed.error}`);
  return parsed.displayStyles;
};

const dollar = profile({ displayWrapper: "double-dollar" });
const bracket = profile({ displayWrapper: "bracket" });
const equation = profile({ displayWrapper: "equation", numbered: false });
const numberedEquation = profile({ displayWrapper: "equation", numbered: true });

// the profile's own wrapper reads back as "default"; another wrapper is the row's choice
assert.deepEqual(styles("$$\na=b\n$$", dollar), ["default"]);
assert.deepEqual(styles("$$a=b$$", dollar), ["default"]);
assert.deepEqual(styles("\\[\na=b\n\\]", dollar), ["bracket"]);
assert.deepEqual(styles("\\begin{equation}\na=b\n\\end{equation}", dollar), ["equation"]);
assert.deepEqual(styles("\\[\na=b\n\\]", bracket), ["default"]);
assert.deepEqual(styles("$$\na=b\n$$", bracket), ["double-dollar"]);
assert.deepEqual(styles("\\begin{equation*}\na=b\n\\end{equation*}", equation), ["default"]);
assert.deepEqual(styles("\\begin{equation}\na=b\n\\end{equation}", equation), ["equation"]);
assert.deepEqual(styles("\\begin{equation}\na=b\n\\end{equation}", numberedEquation), ["default"]);
assert.deepEqual(styles("\\begin{equation*}\na=b\n\\end{equation*}", numberedEquation), ["equation-star"]);
// inline rows have no display wrapper
assert.deepEqual(styles("x $a$ y", dollar), ["default"]);

// writing a document and reading it back changes nothing, whatever the profile
const lines: FormulaLine[] = [
  { id: "a", latex: "a=b", mode: "display", displayStyle: "default" },
  { id: "b", latex: "c=d", mode: "display", displayStyle: "bracket" },
  { id: "c", latex: "e", mode: "inline", displayStyle: "default" },
  { id: "d", latex: "f=g", mode: "display", displayStyle: "equation" },
];
for (const used of [dollar, bracket, equation, numberedEquation]) {
  const source = formatFormulaLinesUniversal(lines, used);
  const parsed = parseUniversalLatexSourceDraft(source, used);
  assert.equal(parsed.valid, true, `${source}: ${parsed.error}`);
  const reread = reconcileFormulaLines(parsed.values, lines, parsed.modes, parsed.displayStyles);
  assert.deepEqual(
    reread,
    lines.map((line) => ({
      ...line,
      // a row whose own choice equals the profile's wrapper reads back as following the profile
      displayStyle:
        line.displayStyle === "bracket" && used.displayWrapper === "bracket"
          ? "default"
          : line.displayStyle === "equation" && used.displayWrapper === "equation" && used.numbered
            ? "default"
            : line.displayStyle,
    })),
    `round trip with ${JSON.stringify(used)}:\n${source}`,
  );
}

// a profile change after a source edit still reaches the rows the edit did not give a wrapper of their own
const edited = parseUniversalLatexSourceDraft("$$\na=q\n$$\n\n$$\nc=d\n$$", dollar);
const editedLines = reconcileFormulaLines(edited.values, lines.slice(0, 2), edited.modes, edited.displayStyles);
assert.match(formatFormulaLinesUniversal(editedLines, bracket), /^\\\[\na=q\n\\\]\n\n\\\[\nc=d\n\\\]$/);

console.log("Universal source display style regression passed");
