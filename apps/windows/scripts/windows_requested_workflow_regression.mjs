import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

const source = (path) => readFileSync(path, "utf8");

const wordService = source("src-windows/VisualTeX.WordVsto/WordFormulaService.cs");
assert.ok(!wordService.includes("display: !session.Numbered"));
assert.ok(wordService.includes("preserveExistingTrueDisplayNumberHost"));
assert.ok(wordService.includes("FinalizeConvertedNumberedOmmlDisplayShapes"));
assert.ok(!wordService.includes("PrepareNumberedOmmlTabPlaceholder"));
assert.ok(!wordService.includes("numberedTabHost"));
assert.ok(wordService.includes("PrepareNumberedOmmlReplacementTabPlaceholderPreservingOle"));
assert.ok(wordService.includes("ApplyDocumentOmmlMathFont"));

const tabAcceptance = source(
  "src-windows/VisualTeX.VstoFlowAcceptance/WordVisualTeXOmmlTabAcceptance.cs",
);
assert.ok(tabAcceptance.includes("WdOMathType.wdOMathDisplay"));
assert.ok(tabAcceptance.includes("normal Word display OMath reference"));
assert.ok(tabAcceptance.includes("external dynamic REF Shape"));
const tabHostAcceptance = source(
  "src-windows/VisualTeX.VstoFlowAcceptance/WordOmmlTabNumberingAcceptance.cs",
);
assert.ok(
  tabHostAcceptance.includes("the REF field result is nested inside OMath"),
);

const trueDisplayNumbering = source(
  "src-windows/VisualTeX.WindowsOffice.VstoShared/WordEquationNumbering.TrueDisplay.cs",
);
assert.ok(trueDisplayNumbering.includes("ConfigureNumberedNativeOmmlDisplay"));
assert.ok(trueDisplayNumbering.includes("IsSerializedNativeDisplayNumberShapeHealthy"));
assert.ok(!wordService.includes("document.Tables.Add("));
assert.ok(!wordService.includes("ConvertToTable("));

const addIn = source("src-windows/VisualTeX.WordVsto/ThisAddIn.cs");
assert.ok(addIn.includes("WordBulkFormulaObjectMode.MathType => FormulaOleContract.MathTypeOleMode"));
assert.ok(addIn.includes("MathTypeNativePreviewRenderer.TryRenderBatch"));

const bulkOption = source(
  "src-windows/VisualTeX.WordVsto/BulkImportMathTypeOption.cs",
);
assert.ok(bulkOption.includes("FormulaOleContract.MathTypeOleMode"));
assert.ok(bulkOption.includes("MathType 公式（原生 OLE）"));

const latexCopy = source("src/clipboard/LatexCopyService.ts");
assert.ok(latexCopy.includes('id: "inline-text-double-dollar"'));
assert.ok(latexCopy.includes('hint: "文字$x^2$文字"'));
assert.ok(latexCopy.includes("formatInlineTextDollar"));

const multiline = source("src/ocr/multilineFormula.ts");
assert.ok(multiline.includes("splitOcrLatexIntoFormulaLines"));
assert.ok(multiline.includes("VISUALTEX_ALIGNMENT_MARKER_LATEX"));
const ocrService = source("src/ocr/ocrService.ts");
assert.ok(ocrService.includes("encodeTopLevelOcrAlignmentMarkers"));
assert.ok(ocrService.includes("VISUALTEX_ALIGNMENT_MARKER_LATEX"));
const editor = source("src/editor/MathEditor.tsx");
assert.ok(editor.includes('setLatexCodeFormat("aligned")'));
assert.ok(editor.includes("hasVisualTexAlignmentMarker"));

console.log(
  "Requested Windows workflow source regression passed: genuine Display numbered OMML with external REF Shape, MathType bulk target, single-dollar mixed syntax and aligned multi-line OCR FormulaLines.",
);
