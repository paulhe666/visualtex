import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

const adapter = readFileSync(
  new URL("../office/macos-offline/word/VTWordAdapter.bas", import.meta.url),
  "utf8",
);
const ribbon = readFileSync(
  new URL("../office/macos-offline/word/customUI14.xml", import.meta.url),
  "utf8",
);

const start = adapter.indexOf("Private Sub VTDocumentImportInsertFormula");
const end = adapter.indexOf("Private Sub VTCancelWordDocumentImportDispatch", start);
assert.ok(start >= 0 && end > start, "Document formula insertion procedure is missing");
const insertion = adapter.slice(start, end);

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

const nativeInsertionStart = adapter.indexOf(
  "Private Function VTInsertNativeEquationAtRange",
);
const nativeInsertionEnd = adapter.indexOf(
  "Private Function VTFinalizeInlineNativeEquation",
  nativeInsertionStart,
);
assert.ok(
  nativeInsertionStart >= 0 && nativeInsertionEnd > nativeInsertionStart,
  "The shared native-equation insertion routine is missing",
);
const nativeInsertion = adapter.slice(
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
  adapter,
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
    adapter.includes(required),
    `Structured document import is missing Word formatting behavior ${required}`,
  );
}
assert.match(
  adapter,
  /VTPlaceCaretAfterInlineNativeEquation formulaRange[\s\S]*Set cursorRange = Selection\.Range\.Duplicate/,
  "Batch OMML insertion must leave the first equation before inserting another inline equation",
);
assert.match(
  adapter,
  /replacesInlineMathAnchor[\s\S]*insertionRange\.Text = plainText/,
  "Text after inline OMML must replace the temporary math-boundary anchor",
);
assert.match(
  ribbon,
  /id="VisualTeX\.Mac\.Word\.DocumentImport"[\s\S]*onAction="VTWordRibbonDocumentImport"/,
  "Word Ribbon must expose the document import command",
);

console.log("Document import Word layout smoke test passed");
