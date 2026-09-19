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
  toggleFormulaSelectionLatex(String.raw`\Delta`, "bold", none),
  String.raw`\mathbf{\Delta}`,
  "uppercase Greek is upright by default and must become bold upright",
);
assert.equal(
  toggleFormulaSelectionLatex(String.raw`\Delta`, "italic", none),
  String.raw`\mathit{\Delta}`,
  "italic toggle must visibly italicize an upright uppercase Greek letter",
);
for (const decoration of ["hat", "vec", "widehat", "bar", "overline", "dot"]) {
  assert.equal(
    toggleFormulaSelectionLatex(`\\${decoration}{\\Delta}`, "bold", none),
    `\\${decoration}{\\mathbf{\\Delta}}`,
    "decorations must preserve the base symbol's upright Greek semantics",
  );
  assert.equal(
    toggleFormulaSelectionLatex(`\\${decoration}{\\Delta}`, "italic", none),
    `\\${decoration}{\\mathit{\\Delta}}`,
  );
  assert.equal(
    toggleFormulaSelectionLatex(`\\${decoration}{\\mathbf{\\Delta}}`, "bold", bold),
    `\\${decoration}{\\Delta}`,
  );
}
assert.equal(
  toggleFormulaSelectionLatex(String.raw`\vec{\hat{\Delta}}`, "bold", none),
  String.raw`\vec{\hat{\mathbf{\Delta}}}`,
  "nested decorations must not add a bold-italic wrapper around the accent",
);
assert.equal(
  toggleFormulaSelectionLatex(String.raw`\mathit{\Delta}`, "italic", italic),
  String.raw`\Delta`,
  "second italic toggle must restore the default upright uppercase Greek form",
);
assert.equal(
  toggleFormulaSelectionLatex(String.raw`\mathbf{\Delta}`, "italic", bold),
  String.raw`\mathbfit{\Delta}`,
  "italic toggle must preserve bold while italicizing uppercase Greek",
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

assert.equal(
  toggleFormulaSelectionLatex(String.raw`\hat{\bm{\mathbfit{\Delta}}}`, "italic", none),
  String.raw`\hat{\mathbf{\Delta}}`,
  "nested serialized variants must toggle italic off even if the accent reports no uniform style",
);
console.log("VisualTeX formula bold/italic toggle regression passed");
