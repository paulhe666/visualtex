import assert from "node:assert/strict";
import {
  formatFormulaLines,
  getLatexCodeFormatDefinition,
  parseLatexSourceDraft,
} from "../src/clipboard/LatexCopyService.ts";
import type { FormulaLine } from "../src/types/formula.ts";
import { resolveNewFormulaLineMode } from "../src/editor/formulaLineMode.ts";
import { normalizeFormulaLines, useEditorStore } from "../src/stores/editorStore.ts";

const lines: FormulaLine[] = [
  { id: "inline-text", latex: String.raw`\text{哈哈哈}abc`, mode: "inline" },
  { id: "display", latex: String.raw`\frac{x}{y}=z`, mode: "display" },
  { id: "inline-around", latex: String.raw`\text{前}y\text{后}`, mode: "inline" },
];

const definition = getLatexCodeFormatDefinition("mixed-inline-display");
assert.equal(definition.id, "mixed-inline-display");
assert.equal(
  resolveNewFormulaLineMode("mixed-inline-display", "display", { shiftKey: true, altKey: false }),
  "inline",
  "Shift+Enter must create an inline row in mixed mode",
);
assert.equal(
  resolveNewFormulaLineMode("mixed-inline-display", "inline", { shiftKey: false, altKey: true }),
  "display",
  "Option+Enter must create a display row in mixed mode",
);
assert.equal(
  resolveNewFormulaLineMode("mixed-inline-display", "inline", { shiftKey: false, altKey: false }),
  "inline",
  "plain Enter must inherit the current mixed row mode",
);
assert.equal(
  resolveNewFormulaLineMode("display-dollar", "inline", { shiftKey: true, altKey: false }),
  undefined,
  "legacy formats must retain their existing Enter behavior",
);

const source = formatFormulaLines(lines, "mixed-inline-display");
assert.equal(
  source,
  ["哈哈哈$abc$", String.raw`$$\frac{x}{y}=z$$`, "前$y$后"].join("\n"),
);

const parsed = parseLatexSourceDraft(source, "mixed-inline-display");
assert.equal(parsed.valid, true);
assert.deepEqual(parsed.values, lines.map((line) => line.latex));
assert.deepEqual(parsed.modes, ["inline", "display", "inline"]);

const mathOnlyInline = formatFormulaLines(
  [{ id: "math-only", latex: "a+b", mode: "inline" }],
  "mixed-inline-display",
);
assert.equal(mathOnlyInline, "$a+b$");

const textOnlyInline = formatFormulaLines(
  [{ id: "text-only", latex: String.raw`\text{只有文字}`, mode: "inline" }],
  "mixed-inline-display",
);
assert.equal(textOnlyInline, "只有文字");
const textOnlyParsed = parseLatexSourceDraft(textOnlyInline, "mixed-inline-display");
assert.equal(textOnlyParsed.valid, true);
assert.deepEqual(textOnlyParsed.values, [String.raw`\text{只有文字}`]);
assert.deepEqual(textOnlyParsed.modes, ["inline"]);

const legacyLine = formatFormulaLines(
  [{ id: "legacy", latex: "E=mc^2" }],
  "mixed-inline-display",
);
assert.equal(legacyLine, "$$E=mc^2$$", "legacy rows must default to display mode");

const normalized = normalizeFormulaLines([
  { id: "saved-inline", latex: "a+b", mode: "inline" },
  { id: "legacy-display", latex: "c+d" },
]);
assert.deepEqual(normalized.map((line) => line.mode), ["inline", "display"]);
useEditorStore.setState({
  title: "mixed persistence",
  lines: normalized,
  activeLineId: normalized[0]?.id ?? null,
  latexCodeFormat: "mixed-inline-display",
});
const savedDocument = useEditorStore.getState().toDocument();
assert.deepEqual(
  savedDocument.formulas.map((formula) => formula.displayMode),
  ["inline", "block"],
  "document save must persist per-row inline/display mode",
);
useEditorStore.getState().loadDocument(savedDocument);
assert.deepEqual(
  useEditorStore.getState().lines.map((line) => line.mode),
  ["inline", "display"],
  "document reopen must restore per-row mode",
);

console.log("mixed LaTeX format regression: PASS");
