import assert from "node:assert/strict";
import {
  toggleFormulaSelectionLatex,
  type FormulaSelectionStyleState,
} from "../src/editor/selectionStyleToggle.ts";

const none: FormulaSelectionStyleState = {
  allBold: false,
  allBoldItalic: false,
  allItalic: false,
  allUpright: false,
};

const bold: FormulaSelectionStyleState = {
  ...none,
  allBold: true,
  allUpright: true,
};

const boldItalic: FormulaSelectionStyleState = {
  ...none,
  allBoldItalic: true,
  allItalic: true,
};

const italic: FormulaSelectionStyleState = {
  ...none,
  allItalic: true,
};

const upright: FormulaSelectionStyleState = {
  ...none,
  allUpright: true,
};

assert.equal(
  toggleFormulaSelectionLatex("abc", "bold", none),
  String.raw`\mathbfit{abc}`,
  "bolding ordinary math must preserve its default italic shape",
);
assert.equal(
  toggleFormulaSelectionLatex(String.raw`\mathbf{abc}`, "bold", bold),
  "abc",
);
assert.equal(
  toggleFormulaSelectionLatex(String.raw`\mathrm{abc}`, "bold", upright),
  String.raw`\mathbf{abc}`,
);
assert.equal(
  toggleFormulaSelectionLatex(String.raw`\mathbfit{abc}`, "bold", boldItalic),
  "abc",
);
assert.equal(
  toggleFormulaSelectionLatex(String.raw`\symbfit{abc}`, "bold", none),
  "abc",
  "symbfit must be recognized as an existing bold-italic wrapper",
);
assert.equal(
  toggleFormulaSelectionLatex(String.raw`\bm{\alpha x}`, "bold", none),
  String.raw`\alpha x`,
  "bm must be recognized as an existing bold-italic wrapper",
);
assert.equal(
  toggleFormulaSelectionLatex(
    String.raw`\mathbf{x}+\mathbf{y}`,
    "bold",
    bold,
  ),
  "x+y",
);
assert.equal(
  toggleFormulaSelectionLatex(
    String.raw`\mathrm{x}+\mathit{y}`,
    "bold",
    none,
  ),
  String.raw`\mathbf{x+y}`,
);

assert.equal(
  toggleFormulaSelectionLatex("xyz", "italic", none),
  String.raw`\mathrm{xyz}`,
);
assert.equal(
  toggleFormulaSelectionLatex(String.raw`\mathrm{xyz}`, "italic", upright),
  "xyz",
);
assert.equal(
  toggleFormulaSelectionLatex(String.raw`\mathbf{xyz}`, "italic", bold),
  String.raw`\mathbfit{xyz}`,
);
assert.equal(
  toggleFormulaSelectionLatex(
    String.raw`\mathbfit{xyz}`,
    "italic",
    boldItalic,
  ),
  String.raw`\mathbf{xyz}`,
);
assert.equal(
  toggleFormulaSelectionLatex(
    String.raw`\mathbf{\mathbfit{xyz}}`,
    "italic",
    boldItalic,
  ),
  String.raw`\mathbf{xyz}`,
);
assert.equal(
  toggleFormulaSelectionLatex(String.raw`\mathit{xyz}`, "italic", italic),
  String.raw`\mathrm{xyz}`,
);
assert.equal(
  toggleFormulaSelectionLatex(
    String.raw`\mathbf{x}+\mathrm{y}`,
    "italic",
    none,
  ),
  String.raw`\mathbfit{x}+y`,
);
assert.equal(
  toggleFormulaSelectionLatex(
    String.raw`\mathbfit{x}+y`,
    "italic",
    italic,
  ),
  String.raw`\mathrm{\mathbf{x}+y}`,
);

console.log("VisualTeX formula bold/italic toggle regression passed");
