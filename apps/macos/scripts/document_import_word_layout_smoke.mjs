import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

const wordModule = (name) => readFileSync(
  new URL(`../office/macos-offline/word/${name}.bas`, import.meta.url),
  "utf8",
);
const documentImport = wordModule("VTWordDocumentImport");
const equations = wordModule("VTWordEquations");
// A procedure's start in a module, whichever scope it has (Public when other modules use it).
const procedureStart = (source, kind, name, from = 0) => {
  const match = new RegExp(`^(?:Private|Public) ${kind} ${name}\\b`, "m").exec(source.slice(from));
  return match ? from + match.index : -1;
};
const ribbon = readFileSync(
  new URL("../office/macos-offline/word/customUI14.xml", import.meta.url),
  "utf8",
);

const start = procedureStart(documentImport, "Sub", "VTDocumentImportInsertFormula");
const end = procedureStart(documentImport, "Sub", "VTCancelWordDocumentImportDispatch", start);
assert.ok(start >= 0 && end > start, "Document formula insertion procedure is missing");
const insertion = documentImport.slice(start, end);

for (const required of [
  "VTPrepareWordCreateInsertionRange",
  "VTInsertNativeEquationAtRange",
  "VTPlaceCaretAfterInlineNativeEquation",
  "VTEnsureNativeEquationNumber",
  "VTAddWordFormulaPicture",
  "VTNormalizeUnnumberedDisplayParagraph",
  "VTEnsureImageEquationNumber",
  "VTApplyWordInlineImageBaseline",
  "VTContinuationRangeAfterDisplayFormula",
  "VTSetWordLatexPayload",
  "VTSetWordOmmlPayload",
  "VTSetWordMetadataPayload",
  "VTSetWordFormulaFormat",
  "VTSetWordImageScaleState",
]) {
  assert.ok(
    insertion.includes(required),
    `Document import must reuse the normal Word layout/state routine ${required}`,
  );
}

const nativeInsertionStart = procedureStart(equations, "Function", "VTInsertNativeEquationAtRange");
const nativeInsertionEnd = procedureStart(
  equations,
  "Function",
  "VTFinalizeInlineNativeEquation",
  nativeInsertionStart,
);
assert.ok(
  nativeInsertionStart >= 0 && nativeInsertionEnd > nativeInsertionStart,
  "The shared native-equation insertion routine is missing",
);
const nativeInsertion = equations.slice(
  nativeInsertionStart,
  nativeInsertionEnd,
);
assert.ok(
  nativeInsertion.includes("VTFinalizeInlineNativeEquation"),
  "The shared native-equation insertion routine must finalize inline OMML",
);
for (const required of [
  "nativeEquation.Type = wdOMathDisplay",
  "nativeEquation.Justification = wdOMathJcCenter",
  "VTNormalizeUnnumberedDisplayParagraph",
]) {
  assert.ok(
    nativeInsertion.includes(required),
    `The shared native-equation insertion routine is missing display behavior ${required}`,
  );
}

assert.match(
  insertion,
  /If displayMode = "inline" Then[\s\S]*VTApplyWordInlineImageBaseline/,
  "Inline image formulas must preserve the normal VisualTeX baseline",
);
assert.match(
  insertion,
  /If displayMode = "block" And numbered Then[\s\S]*VTEnsureNativeEquationNumber/,
  "Numbered native formulas must use the normal true-centering number layout",
);
assert.match(
  insertion,
  /If numbered Then[\s\S]*VTEnsureImageEquationNumber/,
  "Numbered image formulas must use the normal true-centering number layout",
);
assert.match(
  documentImport,
  /insertedRange\.Font\.Reset[\s\S]*insertedRange\.Font\.Italic = False/,
  "Imported prose must not inherit italic formatting from the insertion caret",
);
for (const required of [
  "VTPrepareDocumentImportParagraph",
  "VTFinalizeDocumentImportParagraph",
  "wdStyleHeading1",
  "wdStyleHeading2",
  "ApplyBulletDefault",
  "ApplyNumberDefault",
  "paragraphRange.ParagraphFormat.SpaceAfter = 0!",
]) {
  assert.ok(
    documentImport.includes(required),
    `Structured document import is missing Word formatting behavior ${required}`,
  );
}
assert.match(
  documentImport,
  /VTPlaceCaretAfterInlineNativeEquation formulaRange[\s\S]*Set cursorRange = Selection\.Range\.Duplicate/,
  "Batch OMML insertion must leave the first equation before inserting another inline equation",
);
assert.match(
  documentImport,
  /replacesInlineMathAnchor[\s\S]*insertionRange\.Text = plainText/,
  "Text after inline OMML must replace the temporary math-boundary anchor",
);
assert.match(
  ribbon,
  /id="VisualTeX\.Mac\.Word\.DocumentImport"[\s\S]*onAction="VTWordRibbonDocumentImport"/,
  "Word Ribbon must expose the document import command",
);

console.log("Document import Word layout smoke test passed");
