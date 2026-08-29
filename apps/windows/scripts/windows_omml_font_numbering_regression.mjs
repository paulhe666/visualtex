import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

const source = (path) => readFileSync(path, "utf8");

const service = source("src-windows/VisualTeX.WordVsto/WordFormulaService.cs");
assert.ok(service.includes("ApplyDocumentOmmlMathFont"));
assert.ok(service.includes("document.OMathFontName"));
assert.ok(service.includes("WordOfficeMathFontLoader.LatinModernMathFamily"));
assert.ok(!service.includes("ApplyBulkOmmlRunProperties"));
assert.ok(!service.includes("SplitOmmlFontSegments"));
assert.ok(!service.includes("ApplyOmmlCharacterFonts"));
assert.ok(!service.includes("display: !session.Numbered"));
assert.ok(service.includes("preserveExistingTrueDisplayNumberHost"));
assert.ok(service.includes("FinalizeConvertedNumberedOmmlDisplayShapes"));
assert.ok(!service.includes("PrepareNumberedOmmlTabPlaceholder"));
assert.ok(!service.includes("numberedTabHost"));
assert.ok(service.includes("PrepareNumberedOmmlReplacementTabPlaceholderPreservingOle"));

const fontLoader = source(
  "src-windows/VisualTeX.WordVsto/WordOfficeMathFontLoader.cs",
);
assert.ok(fontLoader.includes('LatinModernMathFamily = "Latin Modern Math"'));
assert.ok(
  fontLoader.includes(
    "6075562B771F8B82F0C179E363389684F2DD09DE30038269E2628E504BD7BE0F",
  ),
);

const packageWxs = source(
  "src-windows/VisualTeX.WindowsOffice.Installer/Package.wxs",
);
assert.ok(packageWxs.includes('FontTitle="Latin Modern Math"'));
assert.ok(packageWxs.includes("latinmodern-math.otf"));

const fontAcceptance = source(
  "src-windows/VisualTeX.VstoFlowAcceptance/WordFormulaFontAcceptance.cs",
);
assert.ok(fontAcceptance.includes("document-level OMML font acceptance passed"));
assert.ok(fontAcceptance.includes("AssertNativeQuadraticOmml"));
assert.ok(fontAcceptance.includes("AssertComparableNativeMathLayout"));
assert.ok(fontAcceptance.includes("AssertTrueTextOmmlTypography"));

const tabAcceptance = source(
  "src-windows/VisualTeX.VstoFlowAcceptance/WordVisualTeXOmmlTabAcceptance.cs",
);
assert.ok(tabAcceptance.includes("WdOMathType.wdOMathDisplay"));
assert.ok(tabAcceptance.includes("pure wdOMathDisplay/m:oMathPara"));
assert.ok(tabAcceptance.includes("external dynamic REF Shape"));
const tabHostAcceptance = source(
  "src-windows/VisualTeX.VstoFlowAcceptance/WordOmmlTabNumberingAcceptance.cs",
);
assert.ok(tabHostAcceptance.includes("FindVisibleEquationNumberRange"));
assert.ok(tabHostAcceptance.includes("the REF field result is nested inside OMath"));
assert.ok(tabHostAcceptance.includes("document.Tables.Count"));
assert.ok(tabHostAcceptance.includes("WdOMathType.wdOMathDisplay"));

const numbering = source(
  "src-windows/VisualTeX.WindowsOffice.VstoShared/WordEquationNumbering.cs",
);
const trueDisplayNumbering = source(
  "src-windows/VisualTeX.WindowsOffice.VstoShared/WordEquationNumbering.TrueDisplay.cs",
);
assert.ok(numbering.includes("CalculateEquationTabStops"));
assert.ok(numbering.includes("FindNumberingOwnerRange"));
assert.ok(numbering.includes("FinalizeConvertedNumberedOmmlDisplayShapes"));
assert.ok(numbering.includes("TryConvertStandardNumberedOmmlTableToStandaloneDisplayParagraph"));
assert.ok(trueDisplayNumbering.includes("ConfigureNumberedNativeOmmlDisplay"));
assert.ok(trueDisplayNumbering.includes("IsSerializedNativeDisplayNumberShapeHealthy"));
assert.ok(trueDisplayNumbering.includes("<m:oMathPara"));
assert.ok(!numbering.includes("EnsureNumberedOmmlUsesInlineTabLayout"));
assert.ok(!numbering.includes("NormalizeNumberedDisplayArgumentSizing"));
assert.ok(!numbering.includes("document.Tables.Add("));
assert.ok(!numbering.includes("ConvertToTable("));
assert.ok(!numbering.includes("columns.Add("));
assert.ok(!service.includes("document.Tables.Add("));
assert.ok(!service.includes("ConvertToTable("));

const scaleAcceptance = source(
  "src-windows/VisualTeX.VstoFlowAcceptance/WordNumberedOmmlTabScaleAcceptance.cs",
);
assert.ok(scaleAcceptance.includes("20"));
assert.ok(scaleAcceptance.includes("tables=0"));
assert.ok(scaleAcceptance.includes("save/reopen"));

const migrationAcceptance = source(
  "src-windows/VisualTeX.VstoFlowAcceptance/WordOmmlNumberingMigrationAcceptance.cs",
);
const emptyRowMigrationAcceptance = source(
  "src-windows/VisualTeX.VstoFlowAcceptance/WordNumberedOmmlEmptyRowAcceptance.cs",
);
assert.ok(
  emptyRowMigrationAcceptance.includes("1x3") ||
    emptyRowMigrationAcceptance.includes("1 x 3"),
);
assert.ok(migrationAcceptance.includes("2x3") || migrationAcceptance.includes("2 x 3"));
assert.ok(migrationAcceptance.includes("pure wdOMathDisplay/m:oMathPara"));
assert.ok(migrationAcceptance.includes("external ordinary REF Shape"));

console.log(
  "Windows OMML document-font, genuine Display numbered OMML, external REF Shape, table-free OLE numbering and legacy migration source regression passed.",
);
