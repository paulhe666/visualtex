import { execFileSync } from "node:child_process";
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";

const root = resolve(new URL("..", import.meta.url).pathname);
const offline = join(root, "office", "macos-offline");
const failures = [];
const notes = [];

function read(relative) {
  return readFileSync(join(root, relative), "utf8");
}

function expect(condition, message) {
  if (!condition) failures.push(message);
}

function expectIncludes(text, value, message) {
  expect(text.includes(value), message ?? `Expected source to contain ${value}`);
}

const requiredFiles = [
  "docs/MACOS_OFFLINE_OFFICE_ARCHITECTURE.md",
  "docs/MACOS_OFFLINE_OFFICE_ACCEPTANCE.md",
  "office/macos-offline/PROTOCOL.md",
  "office/macos-offline/BUILD_ADDINS.md",
  "office/macos-offline/POWERPOINT_INSTALL.md",
  "office/macos-offline/resources/README.md",
  "office/macos-offline/shared/VTProtocol.bas",
  "office/macos-offline/shared/VTOfficePaths.bas",
  "office/macos-offline/shared/VTMetadata.bas",
  "office/macos-offline/shared/VTLauncher.bas",
  "office/macos-offline/shared/VTErrorHandling.bas",
  "office/macos-offline/word/VTRibbonCallbacks.bas",
  "office/macos-offline/word/VTWordAdapter.bas",
  "office/macos-offline/word/VTWordEvents.cls",
  "office/macos-offline/word/customUI14.xml",
  "office/macos-offline/word/VisualTeXWord.scpt",
  "office/macos-offline/powerpoint/VTRibbonCallbacks.bas",
  "office/macos-offline/powerpoint/VTPowerPointAdapter.bas",
  "office/macos-offline/powerpoint/VTPowerPointEvents.cls",
  "office/macos-offline/powerpoint/customUI14.xml",
  "office/macos-offline/powerpoint/VisualTeXPowerPoint.scpt",
  "src-tauri/src/office/macos_offline.rs",
  "src-tauri/src/office/macos_offline_installer.rs",
  "src-tauri/Info.macos.plist",
  "scripts/package_macos_offline_addins.mjs",
  "scripts/register_macos_dev_url_handler.mjs",
  "scripts/tauri_dev.mjs",
  "office-native-dialog.html",
  "src/office/native-dialog-main.tsx",
];

for (const file of requiredFiles) {
  try {
    read(file);
  } catch {
    failures.push(`Missing required macOS offline Office file: ${file}`);
  }
}

const wordRibbon = read("office/macos-offline/word/customUI14.xml");
const powerpointRibbon = read("office/macos-offline/powerpoint/customUI14.xml");
const wordAdapter = read("office/macos-offline/word/VTWordAdapter.bas");
const wordEvents = read("office/macos-offline/word/VTWordEvents.cls");
const powerpointAdapter = read("office/macos-offline/powerpoint/VTPowerPointAdapter.bas");
const powerpointEvents = read("office/macos-offline/powerpoint/VTPowerPointEvents.cls");
const protocol = read("office/macos-offline/shared/VTProtocol.bas");
const officePaths = read("office/macos-offline/shared/VTOfficePaths.bas");
const launcher = read("office/macos-offline/shared/VTLauncher.bas");
const wordScript = read("office/macos-offline/word/VisualTeXWord.scpt");
const powerpointScript = read("office/macos-offline/powerpoint/VisualTeXPowerPoint.scpt");
const rustRuntime = read("src-tauri/src/office/macos_offline.rs");
const nativeInteraction = read("src-tauri/src/office/powerpoint_native.rs");
const appRuntime = read("src-tauri/src/lib.rs");
const installer = read("src-tauri/src/office/macos_offline_installer.rs");
const packager = read("scripts/package_macos_offline_addins.mjs");
const nativeHtml = read("office-native-dialog.html");
const nativeMain = read("src/office/native-dialog-main.tsx");
const dialogApp = read("src/office/dialog/OfficeDialogApp.tsx");
const dialogMessages = read("src/office/dialog/dialogMessages.ts");
const capabilities = read("src-tauri/capabilities/default.json");
const infoPlist = read("src-tauri/Info.macos.plist");
const macSettings = read("src/components/MacOfficeIntegrationSettings.tsx");
const macFirstRun = read("src/components/MacOfficeFirstRunPrompt.tsx");
const macTauriConfig = read("src-tauri/tauri.macos.conf.json");
const platformBundle = read("scripts/build_platform_bundle.mjs");
const lifecycle = read("src-tauri/src/office/lifecycle.rs");

for (const callback of [
  "VTWordRibbonInline",
  "VTWordRibbonDisplay",
  "VTWordRibbonNativeInline",
  "VTWordRibbonNativeDisplay",
  "VTWordRibbonEdit",
  "VTWordRibbonConvertNative",
  "VTWordRibbonNumbering",
  "VTWordRibbonCrossReference",
  "VTWordRibbonOpen",
]) {
  expectIncludes(wordRibbon, `onAction=\"${callback}\"`, `Word Ribbon is missing ${callback}`);
}
for (const callback of [
  "VTPowerPointRibbonNew",
  "VTPowerPointRibbonEdit",
  "VTPowerPointRibbonDelete",
  "VTPowerPointRibbonOpen",
]) {
  expectIncludes(powerpointRibbon, `onAction=\"${callback}\"`, `PowerPoint Ribbon is missing ${callback}`);
}
expectIncludes(wordRibbon, 'id="VisualTeX.Mac.Word.Tab"', "Word must expose a dedicated VisualTeX Ribbon tab");
expectIncludes(wordRibbon, 'label="VisualTeX"', "Word dedicated Ribbon tab must be labelled VisualTeX");
expectIncludes(wordRibbon, 'insertAfterMso="TabInsert"', "Word VisualTeX tab must be placed after Insert");
expectIncludes(wordRibbon, 'onLoad="VTWordRibbonOnLoad"', "Word Ribbon load must initialize the persistent double-click event sink");
expect(!wordRibbon.includes('idMso="TabHome"'), "Word VisualTeX controls must not be injected into Home");
expectIncludes(powerpointRibbon, 'id="VisualTeX.Mac.PowerPoint.Tab"', "PowerPoint must expose a dedicated VisualTeX Ribbon tab");
expectIncludes(powerpointRibbon, 'label="VisualTeX"', "PowerPoint dedicated Ribbon tab must be labelled VisualTeX");
expectIncludes(powerpointRibbon, 'insertAfterMso="TabInsert"', "PowerPoint VisualTeX tab must be placed after Insert");
expectIncludes(powerpointRibbon, 'onLoad="VTPowerPointRibbonOnLoad"', "PowerPoint Ribbon load must initialize the persistent double-click event sink");
expect(!powerpointRibbon.includes('idMso="TabHome"'), "PowerPoint VisualTeX controls must not be injected into Home");

expectIncludes(wordAdapter, "Public Sub AutoExec()", "Word template must publish AutoExec health");
expectIncludes(wordAdapter, "VTInitializeWordEvents", "Word AutoExec must initialize its persistent application event sink");
expectIncludes(wordEvents, "App_WindowBeforeDoubleClick", "Word must use its native application event for double-click editing");
expectIncludes(wordEvents, "Cancel = True", "Word must suppress the default double-click action for a VisualTeX formula");
expectIncludes(wordEvents, "VTVisualTeXInlineShapeAtSelection", "Word double-click editing must resolve a clicked inline formula even when the collapsed selection is adjacent to it");
expectIncludes(wordEvents, "VisualTeX_EditInlineShape", "Word double-click editing must preserve the clicked InlineShape target");
expectIncludes(wordAdapter, "VisualTeX_ApplyPendingResult", "Word template must expose the native result callback");
expectIncludes(wordAdapter, "VisualTeX_DoubleClickEditSelected", "Word must expose a non-modal native double-click macro entry point");
expectIncludes(wordAdapter, "VisualTeX_CreateNativeInline", "Word must expose direct inline OMML insertion");
expectIncludes(wordAdapter, "VisualTeX_CreateNativeDisplay", "Word must expose direct display OMML insertion");
expectIncludes(wordAdapter, "Public Sub VTWordRibbonOnLoad", "The Word Ribbon onLoad callback must initialize the application event sink in an attached isolation template");
expectIncludes(wordAdapter, "Public Sub VTWordRibbonNativeInline", "The visible OMML inline Ribbon button must have a resolvable callback");
expectIncludes(wordAdapter, "Public Sub VTWordRibbonNativeDisplay", "The visible OMML display Ribbon button must have a resolvable callback");
expectIncludes(wordAdapter, "Public Sub VTWordRibbonCrossReference", "The visible Equation-reference Ribbon button must have a resolvable callback");
expectIncludes(wordAdapter, "nativeEquation", "Word requests must preserve the direct native-equation intent");
expectIncludes(wordAdapter, "InlineShapes.AddPicture", "Word formula insertion must create an InlineShape");
expectIncludes(wordAdapter, "placeholder.Width = 1", "Word pending placeholders must remain a one-pixel transaction target");
expectIncludes(wordAdapter, "Selection.SetRange Start:=placeholder.Range.End", "Word must move the caret after the pending formula target");
expectIncludes(wordAdapter, "Private Function VTPrepareWordCreateInsertionRange", "Display creation must isolate a dedicated paragraph before placing its transaction placeholder");
expectIncludes(wordAdapter, "Text = vbCr & vbCr", "A display formula inserted between existing content must preserve both sides around an empty display paragraph");
expectIncludes(wordAdapter, "beforeRange.InlineShapes.Count > 0", "Display paragraph isolation must treat an earlier inline image formula as real surrounding content");
expectIncludes(wordAdapter, 'operationStage = "validate-standalone-image-paragraph"', "Numbered image layout must reject a mixed source paragraph instead of clearing its other content");
expectIncludes(wordAdapter, "surrounding Word content was preserved", "A mixed numbered image paragraph must fail without erasing existing inline formulas or text");
expectIncludes(wordAdapter, "VTEnsureNumberedDisplayTable", "Numbered Word image and OMML formulas must share the stable table layout");
expectIncludes(wordAdapter, "NumRows:=1, NumColumns:=3", "Numbered Word display formulas must use a one-row three-column table");
expectIncludes(wordAdapter, ".AllowAutoFit = False", "The numbered Word display table must not resize itself");
expectIncludes(wordAdapter, ".PreferredWidth = 100!", "The numbered Word display table must fill the text column");
expectIncludes(wordAdapter, "layoutTable.PreferredWidth <> wdUndefined", "The real-host geometry assertion must accept Word for Mac returning wdUndefined after a verified 100-percent table assignment");
expectIncludes(wordAdapter, "layoutTable.Columns(cellIndex).PreferredWidth = 60!", "The formula column must occupy sixty percent of the table");
expectIncludes(wordAdapter, "layoutTable.Columns(cellIndex).PreferredWidth = 20!", "The symmetric side columns must each occupy twenty percent of the table");
expectIncludes(wordAdapter, "wdCellAlignVerticalCenter", "Formula and number cells must use native vertical centering");
expectIncludes(wordAdapter, "numberRange.ParagraphFormat.LineSpacing", "Tall image-number validation must measure the actual number text line height");
expectIncludes(wordAdapter, "formulaCenterY = formulaY + CSng(renderedHeightPoints / 2#)", "Tall image-number validation must compare the rendered formula center rather than its top boundary");
expectIncludes(wordAdapter, "numberCenterY = numberY + numberLineHeight / 2!", "Tall image-number validation must compare the number line visual center");
expectIncludes(wordAdapter, "Application.CaptionLabels(wdCaptionEquation)", "Numbered Word formulas must use Word's built-in Equation caption label");
expectIncludes(wordAdapter, "VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX", "Numbered Word formulas must Bookmark the native SEQ result separately");
expectIncludes(wordAdapter, "Type:=wdFieldEmpty", "Numbered Word formulas must create a real localized Equation SEQ field in a dedicated paragraph");
expectIncludes(wordAdapter, ".Font.Size = 1!", "The native Equation helper paragraph must be compressed to one point");
expectIncludes(wordAdapter, ".Font.Hidden = False", "The native Equation SEQ must remain non-hidden so Word for Mac produces a usable field result");
expect(!wordAdapter.includes(".Font.Hidden = True"), "The native Equation numbering path must never hide the SEQ result with Word's Hidden font property");
expectIncludes(wordAdapter, 'rightContent.Text = "()"', "The right cell must use ordinary parentheses around a native REF field");
expectIncludes(wordAdapter, "VTRefreshEquationNumberMirror", "Numbering must refresh the native SEQ target and remove legacy mirror artifacts");
expectIncludes(wordAdapter, "VTEquationSequenceFieldHasOrdinal", "Numbering must validate the native SEQ through its retained legal ordinal field code");
expectIncludes(wordAdapter, "expectedText = CStr(sequenceOrdinal)", "The native SEQ result must equal the reconciled document-order ordinal");
expect(!wordAdapter.includes("insertionRange.InsertAfter expectedText & vbCr"), "Numbering must not create a second plain-text number paragraph");
expect(!wordAdapter.includes("sequenceParagraph.InsertParagraphAfter"), "Numbering must not create a mirror paragraph beside the native SEQ");
expect(!wordAdapter.includes("sequenceField.Result.Text = expectedText"), "Numbering must not overwrite a Word field result Range");
expectIncludes(wordAdapter, "VTParenthesizedEquationReferenceFieldText( _\n            sequenceBookmarkName)", "The visible right-cell REF must target the native SEQ result Bookmark");
expectIncludes(wordAdapter, "formulaIds = VTValidNumberedFormulaIds(documentObject)", "The Equation picker must enumerate only live VisualTeX numbered formulas");
expectIncludes(wordAdapter, "Text:=VTParenthesizedEquationReferenceFieldText( _", "Body Equation references must target the live VT_N_ Bookmark with a native REF field");
const crossReferenceCommandStart = wordAdapter.indexOf("Public Sub VisualTeX_OpenEquationCrossReference()");
const crossReferenceCommandEnd = wordAdapter.indexOf("Public Sub VisualTeX_OpenApplication()", crossReferenceCommandStart);
const crossReferenceCommandSource = wordAdapter.slice(crossReferenceCommandStart, crossReferenceCommandEnd);
expect(!crossReferenceCommandSource.includes("Selection.InsertCrossReference"), "The VisualTeX picker must not depend on Word's stale global caption inventory");
expectIncludes(wordAdapter, "Selection.InsertCrossReference _", "The reference persistence regression must create a real Word built-in Equation cross-reference");
expectIncludes(wordAdapter, 'regressionStage = "image-cross-reference-refresh-cycle"', "The real-host regression must repeat field refresh before accepting a stable (1) Bookmark");
expectIncludes(wordAdapter, 'regressionStage = "image-cross-reference-later-number"', "The real-host regression must prove a later numbered formula leaves an earlier native REF at (1)");
expectIncludes(wordAdapter, "VTInsertEquationNumberReferenceAtRange", "Word Ribbon must insert a native Equation cross-reference");
expectIncludes(wordAdapter, "VTReconcileEquationNumbers ActiveDocument", "Manual Equation number refresh must rebuild SEQ, visible REF and body REF fields together");
expectIncludes(wordAdapter, "Type:=wdFieldRef", "Visible Equation numbers must remain dynamic REF fields");
expectIncludes(wordAdapter, "VT_WORD_NUMBER_BOOKMARK_PREFIX", "Word must retain a formula-specific Bookmark around the complete visible (n) range");
expect(!wordAdapter.includes("VisualTeXEquation"), "New Word numbering must not use the legacy VisualTeX-only sequence name");
expectIncludes(wordAdapter, "Private Sub VTFormatVisibleEquationReference", "Visible Equation number formatting must be centralized and verified");
expectIncludes(wordAdapter, ".Position = 0\n        .Size = visibleSize", "Visible Equation numbers must use zero baseline shift and a normal visible font size");
expect(!wordAdapter.includes("formulaHeightPoints / 2#"), "Equation-number Font.Position must never be derived from rendered image height");
expectIncludes(wordAdapter, "nativeEquation.Type = wdOMathDisplay", "Numbered OMML must retain Word native display-math layout");
expectIncludes(wordAdapter, "nativeEquation.BuildUp", "Transferred numbered OMML must be rebuilt by Word after it is re-resolved in the center cell");
expectIncludes(wordAdapter, "Private Function VTWrapNativeDisplayParagraphInTable", "Numbered native OMML must wrap its existing display paragraph in place");
expectIncludes(wordAdapter, "Set paragraphRange = nativeEquation.Range.Paragraphs(1).Range.Duplicate", "The in-place native display path must resolve the equation's owning paragraph");
expectIncludes(wordAdapter, "paragraphStart = paragraphRange.Start", "The same-document display path must retain a stable source-paragraph anchor");
expectIncludes(wordAdapter, "nativeEquation.Type = wdOMathDisplay", "The source OMath must be promoted before the complete display paragraph is transferred");
expectIncludes(wordAdapter, "Set paragraphRange = documentObject.Range( _\n        Start:=paragraphStart, End:=paragraphStart).Paragraphs(1).Range.Duplicate", "The same-document display path must re-resolve its source paragraph after BuildUp");
expectIncludes(wordAdapter, "Word lost the unique native OMath while building display math.", "The same-document display path must reject a stale or ambiguous source OMath");
expectIncludes(wordAdapter, "A numbered display OMath must occupy its own paragraph.", "The same-document display path must reject mixed text paragraphs");
expectIncludes(wordAdapter, 'operationStage = "prepare-source-inline"', "The native display wrapper must identify failures while returning the source OMath to inline mode");
expectIncludes(wordAdapter, "nativeEquation.Type = wdOMathInline", "The source OMath must be forced inline before an empty table is created beside it");
expectIncludes(wordAdapter, "Word did not return the source OMath to inline mode.", "The wrapper must verify Mac Word accepted inline mode before table creation");
expectIncludes(wordAdapter, 'operationStage = "create-empty-table"', "The native display wrapper must identify failures while creating the empty table");
expectIncludes(wordAdapter, "Set layoutTable = documentObject.Tables.Add( _\n        Range:=insertionRange, NumRows:=1, NumColumns:=3)", "Numbered native OMML must create its final one-row three-column table directly in the target document");
expectIncludes(wordAdapter, 'operationStage = "promote-source-display"', "The native display wrapper must promote the source only after the table anchor is safe");
const nativeWrapStart = wordAdapter.indexOf("Private Function VTWrapNativeDisplayParagraphInTable");
const nativeWrapEnd = wordAdapter.indexOf("Private Sub VTConfigureNumberedDisplayTable", nativeWrapStart);
const nativeWrapSource = wordAdapter.slice(nativeWrapStart, nativeWrapEnd);
expect(nativeWrapSource.indexOf("nativeEquation.Type = wdOMathInline") < nativeWrapSource.indexOf("Set layoutTable = documentObject.Tables.Add"), "Mac Word must return the source OMath to inline mode before creating the empty table");
expect(nativeWrapSource.indexOf("Set layoutTable = documentObject.Tables.Add") < nativeWrapSource.indexOf("nativeEquation.Type = wdOMathDisplay"), "Mac Word must create the empty table before promoting the source OMath to display math");
expectIncludes(wordAdapter, "centerRange.FormattedText = sourceParagraph.FormattedText", "The freshly re-resolved complete display paragraph must transfer into the center cell without crossing documents");
expectIncludes(wordAdapter, "Word could not re-resolve the source display paragraph after table creation.", "The same-document transfer must reject a stale source Range after inserting the table");
expectIncludes(wordAdapter, "Same-document transfer downgraded the center-cell equation.", "The center-cell OMath must remain native display math after transfer");
expectIncludes(wordAdapter, "sourceParagraph.Delete", "The original display paragraph must be removed only after the center-cell transfer is verified");
expectIncludes(wordAdapter, "Private Sub VTCompactNativeDisplayCellTail", "Numbered display OMML must preserve and compact Word's required empty cell tail");
expectIncludes(wordAdapter, ".LineSpacingRule = wdLineSpaceExactly", "The required display-cell tail must use an exact compact line box");
expectIncludes(wordAdapter, ".LineSpacing = 1!", "The required display-cell tail must contribute only one point of vertical space");
expectIncludes(wordAdapter, 'InStr(1, cellXml, "<m:oMathPara", vbBinaryCompare) = 0', "The production display path must verify real m:oMathPara XML instead of trusting OMath.Type");
expect(!nativeWrapSource.includes("separatorRange.Delete"), "The production display wrapper must not delete the paragraph mark that carries m:oMathPara semantics");
expectIncludes(wordAdapter, "layoutTable.Cell(1, 2).Range.OMaths.Count <> 1", "The transferred native display table must verify exactly one center-cell OMath");
expectIncludes(wordAdapter, '"VTWrapNativeDisplayParagraphInTable/" & operationStage', "Native display failures must expose their exact internal operation stage");
expect(!wordAdapter.includes("paragraphRange.ConvertToTable"), "Numbered native OMML must not call ConvertToTable on a mathematical paragraph");
expectIncludes(wordAdapter, 'operationStage = "reuse-table"', "Existing numbered native formulas must reuse their established one-row three-column table");
expectIncludes(wordAdapter, "Set layoutTable = formulaRange.Tables(1)", "Existing numbered native formulas must not be wrapped into nested tables");
expectIncludes(wordAdapter, 'operationStage = "repair-native-display-cell"', "Editing an older numbered native formula must upgrade a downgraded center cell");
expectIncludes(wordAdapter, "Private Function VTRebuildExistingNativeDisplayCell", "Existing numbered native formulas must rebuild m:oMathPara without replacing their numbering table");
expectIncludes(wordAdapter, 'Const temporaryBookmarkName As String = "VT_TMP_DISPLAY_REPAIR"', "Native display repair must track its same-document temporary source across Range shifts");
expectIncludes(wordAdapter, "centerRange.FormattedText = _\n        documentObject.Bookmarks(temporaryBookmarkName).Range.FormattedText", "Native display repair must replace the complete center-cell content with the same-document display paragraph in one operation");
const nativeRepairStart = wordAdapter.indexOf("Private Function VTRebuildExistingNativeDisplayCell");
const nativeRepairEnd = wordAdapter.indexOf("Private Sub VTCompactNativeDisplayCellTail", nativeRepairStart);
const nativeRepairSource = wordAdapter.slice(nativeRepairStart, nativeRepairEnd);
expect(!nativeRepairSource.includes('operationStage = "clear-center-cell"'), "Image-to-native repair must not delete the center cell before transferring its display paragraph");
expectIncludes(nativeRepairSource, 'operationStage = "append-required-tail"', "Image-to-native repair must restore Word's required second cell paragraph when the imported paragraph mark is merged");
expectIncludes(nativeRepairSource, "centerRange.InsertBefore vbCr", "Image-to-native repair must insert the required tail immediately before the cell-end marker");
expectIncludes(wordAdapter, '"|centerStructure="', "Continuous-insertion invariants must include m:oMathPara and compact-tail geometry");
expect(!wordAdapter.includes("VTMaterializeNativeDisplayCellFromPayload"), "Numbered native OMML must not use the abandoned payload rematerialization path");
expect(!wordAdapter.includes("VTDisplayFlatOpcXml"), "Numbered native OMML must not generate Flat OPC packages");
expect(!wordAdapter.includes(".display.xml"), "Numbered native OMML must not create Flat OPC staging files");
expect(!wordAdapter.includes(".display.docx"), "Numbered native OMML must not create temporary display DOCX files");
expect(!wordAdapter.includes("SaveAs2"), "Numbered native OMML must not serialize a second staging document");
expect(!wordAdapter.includes("InsertFile"), "Numbered native OMML must not import a temporary document into the formula cell");
expectIncludes(wordAdapter, 'centerRange.Text = "AZ"', "The inline OMML size baseline must remain surrounded by ordinary text so Word cannot auto-promote it to display math");
expectIncludes(wordAdapter, "formulaInsert.FormattedText = equationRange.FormattedText", "The inline OMML size baseline must insert only the equation Range between its text anchors");
expect(!wordAdapter.includes("backupInsert.FormattedText = sourceParagraph.FormattedText"), "Display OMML must not use the failed hidden-document OMath.Type copy workaround");
expectIncludes(wordAdapter, "contextRange.Document.Styles(wdStyleNormal).Font.Size", "Display OMML must inherit the document Normal size instead of using an arbitrary font boost");
expect(!wordAdapter.includes("VT_NATIVE_DISPLAY_FONT_SCALE"), "Display OMML must not uniformly scale every glyph instead of using Word display style");
expectIncludes(wordAdapter, "VTNumberedEquationInvariantSnapshot", "The real-host regression must snapshot earlier numbered formulas across later insertions");
expectIncludes(wordAdapter, 'regressionStage = "continuous-insertion-native-numbered"', "The real-host regression must cover consecutive numbered native OMML insertion");
expectIncludes(wordAdapter, "VTEnsureEquationNumberFields layoutTable, formulaId", "Image and OMML numbering must create the same native SEQ and visible REF structure");
expectIncludes(wordAdapter, "VTInsertDedicatedEquationHelperParagraph", "Every numbered formula must own a new helper paragraph instead of reusing the following body paragraph");
expectIncludes(wordAdapter, "insertionRange.Text = vbCr", "A numbered formula must insert a dedicated helper paragraph at its own table boundary");
expectIncludes(wordAdapter, "PreserveFormatting:=False", "Visible right-cell REF fields must not inherit the white 1pt SEQ helper formatting");
expectIncludes(wordAdapter, 'InStr(1, referenceField.Code.Text, "MERGEFORMAT"', "Visible Equation numbers must reject MERGEFORMAT fields that can inherit invisible helper formatting");
expectIncludes(wordAdapter, "VTFormatVisibleEquationReference", "Visible Equation numbers must explicitly restore normal size, automatic color and non-hidden formatting");
expectIncludes(wordAdapter, "VTReconcileEquationNumbers documentObject\n    Set sequenceField = VTResolveEquationSequenceFieldNear", "Word numbering must reconcile native SEQ fields before creating the new visible right-cell REF");
expectIncludes(wordAdapter, "VTVerifyEquationNumberFieldIntegrity", "Word numbering must verify VT_N_, VT_C_, VT_R_ and the visible REF after every real insertion");
expectIncludes(wordAdapter, "complete VT_N_/VT_C_/VT_R_", "Word numbering must reject a missing native sequence Bookmark instead of leaving empty parentheses");
expectIncludes(wordAdapter, "Public Sub VTCleanupOrphanedNumberedDisplaySelection", "Deleting a numbered formula object must have a targeted orphan-table cleanup path");
expectIncludes(wordAdapter, "layoutTable.Cell(1, 2).Range.InlineShapes.Count <> 0", "Orphan cleanup must leave a numbered table untouched while its image formula still exists");
expectIncludes(wordAdapter, "layoutTable.Cell(1, 2).Range.OMaths.Count <> 0", "Orphan cleanup must leave a numbered table untouched while its native OMath still exists");
expectIncludes(wordAdapter, "VTDeleteWordLatexPayload documentObject, formulaId", "Orphan cleanup must delete the removed formula's document payload state");
expectIncludes(wordAdapter, "VTNormalizePlainWordParagraph paragraphRange", "Orphan cleanup must restore an ordinary body-text paragraph and caret");
expectIncludes(wordAdapter, "Public Sub VisualTeX_WatchOrphanedNumberedDisplay()", "Word must run a lightweight adapter-local orphan watcher without adding another modified source module");
expectIncludes(wordAdapter, "Application.OnTime", "The orphan watcher must defer cleanup until after the user's Delete transaction settles");
expectIncludes(wordAdapter, "VTAnyOpenDocumentHasEquationNumbers()", "The orphan watcher must stop scheduling when no VisualTeX numbered formulas remain");
expectIncludes(wordAdapter, "Not VTWordInternalMutationActive()", "The orphan watcher must remain disabled during an internal VisualTeX mutation");
expectIncludes(wordAdapter, "VTCleanupOrphanedNumberedDisplaySelection Selection.Range", "The orphan watcher must inspect only the current selection and its adjacent table");
expectIncludes(wordAdapter, "VTRequireWritableWordDocument\n    VTCleanupOrphanedNumberedDisplaySelection Selection.Range", "Every new formula command must synchronously clean a nearby orphan table before inserting its placeholder");
expectIncludes(wordAdapter, "Public Sub VisualTeX_RunWordUserWorkflowRegression()", "The real host must cover inline-then-display insertion, numbering and direct deletion cleanup");
expectIncludes(wordAdapter, 'regressionStage = "prepare-image-display-after-inline"', "The user workflow regression must preserve an earlier inline formula on the same input line");
expectIncludes(wordAdapter, 'regressionStage = "number-image-display-visible"', "The workflow regression must reject an image Equation number that exists but is visually invisible");
expectIncludes(wordAdapter, 'regressionStage = "insert-native-numbered-after-image"', "The workflow regression must reproduce image-numbered then OMML-numbered insertion order");
expectIncludes(wordAdapter, 'regressionStage = "verify-both-visible-after-second-insertion"', "The workflow regression must re-check both visible numbers after the second numbered formula is inserted");
expectIncludes(wordAdapter, 'regressionStage = "image-display-continuation"', "The workflow regression must type ordinary text after an image display formula");
expectIncludes(wordAdapter, 'regressionStage = "native-display-continuation"', "The workflow regression must type ordinary text after an OMML display formula");
expectIncludes(wordAdapter, 'Selection.TypeText Text:="workflow continuation"', "The image display continuation test must perform a real Word typing operation");
expectIncludes(wordAdapter, 'Selection.TypeText Text:="native continuation"', "The OMML display continuation test must perform a real Word typing operation");
expectIncludes(wordAdapter, 'regressionStage = "delete-image-numbered-display"', "The user workflow regression must delete a numbered image and verify structural cleanup");
expectIncludes(wordAdapter, "Public Sub VisualTeX_RunWordDeletionReferenceRegression()", "The real host must cover four numbered formulas, mixed deletion, renumbering and references");
expectIncludes(wordAdapter, 'regressionStage = "create-four-numbered-formulas"', "Deletion regression must begin with four live numbered image/OMML formulas");
expectIncludes(wordAdapter, 'regressionStage = "delete-native-and-image-formulas"', "Deletion regression must remove both an OMML formula and an image formula");
expectIncludes(wordAdapter, 'regressionStage = "prune-and-renumber-after-deletion"', "Deletion regression must garbage-collect orphan scaffolds before renumbering");
expectIncludes(wordAdapter, 'regressionStage = "verify-picker-items-match-live-formulas"', "Deletion regression must require the picker to match only live formulas and previews");
expectIncludes(wordAdapter, 'regressionStage = "insert-fresh-live-references"', "Deletion regression must insert fresh dynamic references to both surviving formulas");
expectIncludes(wordAdapter, 'regressionStage = "reject-broken-native-reference-results"', "Deletion regression must reject any empty or invalid surviving native Equation reference result");
expectIncludes(wordAdapter, "Public Sub VisualTeX_RunWordReferencePersistenceRegression()", "The real host must verify VisualTeX and Word built-in references before and after native renumbering");
const referenceRegressionStart = wordAdapter.indexOf("Public Sub VisualTeX_RunWordReferencePersistenceRegression()");
const referenceRegressionEnd = wordAdapter.indexOf("Public Sub VisualTeX_RunWordDisplayStrategyProbe()", referenceRegressionStart);
const referenceRegressionSource = wordAdapter.slice(referenceRegressionStart, referenceRegressionEnd);
expectIncludes(wordAdapter, 'regressionStage = "verify-native-caption-inventory"', "Reference persistence regression must begin with exactly two VisualTeX Equation captions");
expect(!referenceRegressionSource.includes('regressionStage = "create-existing-word-equation-caption"'), "Reference visibility acceptance must not add an unrelated ordinary Equation caption to the two-formula user workflow");
expectIncludes(wordAdapter, 'regressionStage = "insert-visualtex-native-reference"', "The VisualTeX picker must insert through Word's native cross-reference API");
expectIncludes(wordAdapter, 'regressionStage = "insert-word-built-in-reference"', "Reference persistence regression must also call Word's built-in InsertCrossReference path directly");
expectIncludes(wordAdapter, '"initial VisualTeX Equation reference"', "The VisualTeX reference must be visible at normal size immediately after insertion");
expectIncludes(wordAdapter, 'regressionStage = "verify-word-built-in-initial-visibility"', "Word built-in acceptance must verify the reference before any VisualTeX refresh or watcher repair runs");
expectIncludes(wordAdapter, '"initial Word built-in Equation reference"', "The first Word-native Equation reference must be visibly formatted immediately after insertion");
expectIncludes(wordAdapter, 'Trim$(builtInReferenceField.Result.Text) <> "2"', "The initial Word built-in reference must contain exactly the selected Equation number, without adjacent caption contamination");
expect(referenceRegressionSource.indexOf('"initial VisualTeX Equation reference"') < referenceRegressionSource.indexOf('regressionStage = "insert-word-built-in-reference"'), "VisualTeX reference visibility must be checked immediately before inserting the Word built-in reference");
expect(referenceRegressionSource.indexOf('"initial Word built-in Equation reference"') < referenceRegressionSource.indexOf("VTReconcileEquationNumbers testDocument", referenceRegressionSource.indexOf('regressionStage = "insert-word-built-in-reference"')), "Word built-in reference visibility must be checked before any VisualTeX refresh");
expectIncludes(wordAdapter, "ReferenceKind:=wdOnlyLabelAndNumber", "Word built-in acceptance must cover a native reference presentation different from the VisualTeX picker");
expectIncludes(wordAdapter, 'regressionStage = "delete-preceding-formula-and-renumber"', "Reference persistence regression must preserve both references when the earlier formula is deleted and numbers are refreshed");
expect(!referenceRegressionSource.includes('regressionStage = "convert-referenced-image-to-native"'), "Reference visibility acceptance must not mix image-to-OMML conversion into the two-reference renumbering workflow");
expectIncludes(wordAdapter, "documentObject.GetCrossReferenceItems(wdCaptionEquation)", "The VisualTeX picker inventory must be grounded in Word's native Equation caption list");
expectIncludes(wordAdapter, "Selection.InsertCrossReference _", "The VisualTeX picker must delegate reference field creation to Word");
expectIncludes(wordAdapter, "ReferenceItem:=nativeReferenceItem", "The VisualTeX picker must pass Word's native Equation item index to InsertCrossReference");
expectIncludes(wordAdapter, "Private Function VTNativeEquationReferenceItemForFormula", "VisualTeX formulas must map their durable formula id to the corresponding native Equation item without touching Word's private _Ref Bookmark");
const sequenceCodeStart = wordAdapter.indexOf("Private Function VTEquationSequenceFieldCodeForOrdinal");
const sequenceCodeEnd = wordAdapter.indexOf("Private Function VTNormalizeEquationFieldCode", sequenceCodeStart);
const sequenceCodeSource = wordAdapter.slice(sequenceCodeStart, sequenceCodeEnd);
expect(!sequenceCodeSource.includes('" \\r "'), "VisualTeX captions must remain flowing native SEQ fields instead of being restarted during every renumber");
const reconcileStart = wordAdapter.indexOf("Private Sub VTReconcileEquationNumbers");
const reconcileEnd = wordAdapter.indexOf("Private Function VTCustomTabStopCount", reconcileStart);
const reconcileSource = wordAdapter.slice(reconcileStart, reconcileEnd);
expect(!reconcileSource.includes("VTCaptureBodyEquationReferenceBindings"), "Native renumbering must not inspect Word's private _Ref Bookmark implementation");
expect(!reconcileSource.includes("VTRestoreBodyEquationReferenceBindings"), "Native renumbering must not rebuild Word-managed _Ref Bookmarks");
const convertStart = wordAdapter.indexOf("Private Sub VTWordConvertInlineShapeToNativeEquation");
const convertEnd = wordAdapter.indexOf("Private Function VTNativeWordDocumentPath", convertStart);
const convertSource = wordAdapter.slice(convertStart, convertEnd);
expect(!convertSource.includes("VTCaptureBodyEquationReferenceBindings"), "Image-to-OMML conversion must preserve native captions instead of snapshotting private REF targets");
expect(!convertSource.includes("VTRestoreBodyEquationReferenceBindings"), "Image-to-OMML conversion must let Word maintain its own cross-reference targets");
expectIncludes(convertSource, "VTBeginWordInternalMutation", "Image-to-OMML conversion must suppress the deletion watcher throughout its structural transaction");
expectIncludes(convertSource, "VTEndWordInternalMutation", "Image-to-OMML conversion must always release its internal mutation guard");
expect(convertSource.indexOf("VTBeginWordInternalMutation") < convertSource.indexOf("DoEvents"), "Image-to-OMML conversion must enter its mutation guard before deferred Word events can run");
expect(convertSource.indexOf("DoEvents") < convertSource.indexOf("VTSetNativeFormulaBookmark _"), "Image-to-OMML conversion must settle deferred events before persisting final VT_F_");
expect(convertSource.indexOf("VTSetNativeFormulaBookmark _") < convertSource.indexOf("VTEndWordInternalMutation"), "Image-to-OMML conversion must retain its mutation guard until final formula identity is verified");
expectIncludes(convertSource, "nativeBookmarkAnchor = equationRange.Start", "Image-to-OMML conversion must retain a post-layout anchor before native field reconciliation");
expectIncludes(convertSource, "Set finalFormulaRange = _\n            numberLayoutRange.OMaths(1).Range.Duplicate", "A numbered image-to-OMML conversion must re-resolve the final formula from its stable numbered table");
expect(convertSource.indexOf("VTReconcileEquationNumbers targetDocument") < convertSource.lastIndexOf("VTSetNativeFormulaBookmark _"), "Image-to-OMML conversion must persist VT_F_ only after native SEQ and REF reconciliation is complete");
expectIncludes(convertSource, "DoEvents\n\n    ' Selection changes and the deferred orphan watcher are allowed to settle", "Image-to-OMML conversion must let Selection and orphan-watcher events settle before final identity persistence");
expectIncludes(convertSource, "VTEnsureEquationNumberFields finalLayoutTable, formulaId", "Image-to-OMML finalization must restore the final VT_R_/VT_N_/VT_C_ scaffold before persisting VT_F_");
const finalNativeBookmarkIndex = convertSource.lastIndexOf("VTSetNativeFormulaBookmark _");
const postFinalNativeBookmarkSource = convertSource.slice(finalNativeBookmarkIndex);
expect(!postFinalNativeBookmarkSource.includes("VTReconcileEquationNumbers"), "No SEQ/REF reconciliation may run after the final converted VT_F_ Bookmark is persisted");
expect(!postFinalNativeBookmarkSource.includes("DoEvents"), "No deferred Word event processing may run after the final converted VT_F_ Bookmark is persisted");
expect(!postFinalNativeBookmarkSource.includes(".Select"), "Image-to-OMML conversion must not change Selection after final identity persistence");
expectIncludes(wordAdapter, "layoutTable.Cell(1, 2).Range.OMaths.Count <> 1", "Image-to-OMML acceptance must inspect the converted formula's own numbered table instead of relying on document-level OMaths.Count");
expect(!wordAdapter.includes("bodyReferenceCount <> 2 Or testDocument.OMaths.Count <> 1"), "Reference acceptance must not use unstable document-level OMath inventory as a proxy for conversion success");
expectIncludes(wordAdapter, "If Len(formulaId) > 0 And _", "Orphan cleanup must delete only native Equation captions that carry a VisualTeX formula identity");
expectIncludes(wordAdapter, "VTPruneOrphanedEquationNumberScaffolds", "Number update and cross-reference commands must prune orphan SEQ/table scaffolds document-wide");
expectIncludes(wordAdapter, 'replacementRange.Text = "(deleted equation)"', "A body reference whose formula was deleted must become an explicit non-field marker");
expectIncludes(wordAdapter, "Private Sub VTFormatBodyEquationReference", "Body Equation reference formatting must not reuse right-cell paragraph formatting");
expectIncludes(wordAdapter, "For Each candidateField In paragraphRange.Fields", "The compact helper paragraph must directly format its native SEQ result as the source for first-time Word cross-reference insertion");
expectIncludes(wordAdapter, ".Size = visibleSize\n                .Hidden = False\n                .Color = wdColorAutomatic", "The native SEQ source must use normal body size and automatic color so Word's first built-in reference is immediately visible");
expectIncludes(wordAdapter, "Private Sub VTNormalizeBodyEquationReferenceVisibility", "Every Equation number refresh must restore user-visible formatting for body references");
expectIncludes(wordAdapter, "VTNormalizeBodyEquationReferenceVisibility documentObject", "Native REF reconciliation must normalize body reference size after Word updates the fields");
expectIncludes(wordAdapter, "VTNormalizeBodyEquationReferenceVisibility ActiveDocument", "The orphan watcher must repair body reference visibility after manual Word field updates");
expectIncludes(wordAdapter, "Private Sub VTAssertBodyEquationReferenceVisible", "The real-host reference regression must reject a one-point body REF even when its result text is correct");
expectIncludes(wordAdapter, '"renumbered native Equation reference"', "Deletion acceptance must verify that surviving native references remain visibly formatted after renumbering");
expectIncludes(wordAdapter, '"renumbered native Equation reference"', "Both surviving references must remain visibly formatted after deletion and renumbering");
expect(referenceRegressionSource.indexOf('regressionStage = "delete-preceding-formula-and-renumber"') < referenceRegressionSource.indexOf('"renumbered native Equation reference"'), "Post-refresh visibility must be checked only after deleting the preceding formula and reconciling numbers");
expectIncludes(wordAdapter, '"renumberedVisualTeXReference=1"', "The reference regression PASS report must record the VisualTeX reference after renumbering");
expectIncludes(wordAdapter, '"renumberedWordBuiltInReference=1"', "The reference regression PASS report must record the Word built-in reference after renumbering");
expect(!wordAdapter.includes('"Image-to-OMML conversion lost the surviving formula identity."'), "Reference acceptance must not fail on an unrelated formula-identity assertion that contradicts hands-on editing");
expectIncludes(wordAdapter, "Private Function VTHelperParagraphOwnsNativeEquationSequence", "Orphan cleanup must verify a helper paragraph owns exactly one native Equation SEQ before deleting it");
expectIncludes(wordAdapter, "Private Sub VTPruneUnbookmarkedEmptyNumberTables", "Document-wide cleanup must remove empty VisualTeX number tables even after VT_R_ is lost");
expectIncludes(wordAdapter, "Private Sub VTRepairLiveNumberedTableScaffolds", "Number refresh must repair missing VT_R_/VT_N_/VT_C_ for a still-live formula instead of deleting it");
expectIncludes(wordAdapter, '"(" & CStr(ordinal) & ")  " & previewText', "The Equation picker must show each live number together with its formula preview");
expectIncludes(wordAdapter, "sourceHeightPoints = target.Height", "Image-to-OMML conversion must preserve the source formula height for number alignment");
expectIncludes(wordAdapter, "VTEnsureNativeEquationNumber", "Image-to-OMML conversion must rebuild the shared numbered table around the native formula");
expectIncludes(wordAdapter, "target.Delete", "Word replacement must delete the old object only after candidate setup");
expectIncludes(wordAdapter, "Public Sub VisualTeX_ConvertSelectedToNativeEquation()", "Word must expose a selected-formula native equation conversion command");
expectIncludes(wordAdapter, "Set targetDocument = target.Range.Document", "Word image-to-native conversion must retain the image's owning document across hidden DOCX staging");
expectIncludes(wordAdapter, "VTSetWordMetadataPayload targetDocument, formulaId, encodedMetadata", "Word image-to-native conversion must store metadata in the owning document rather than a transient ActiveDocument");
expectIncludes(wordAdapter, "insertionAnchor.Collapse wdCollapseEnd", "Word image-to-native conversion must insert after the source picture so deleting the picture shifts the OMath into place");
expectIncludes(wordAdapter, "Set sourceImage = VTFindUniqueInlineShape(encodedMetadata)", "Word image-to-native conversion must resolve a fresh picture object after hidden DOCX staging");
expectIncludes(wordAdapter, "sourceImage.Delete", "Word image-to-native conversion must remove the source picture transactionally before resolving the final OMath Range");
expectIncludes(wordAdapter, "sourceBackupRange.FormattedText", "Word image-to-native conversion must retain an exact source-image rollback copy");
expectIncludes(wordAdapter, "replacementBackupRange.FormattedText", "Word native Range replacement must keep a formatted backup for rollback");
expectIncludes(wordAdapter, "Documents.Open( _", "Word native conversion must open a real DOCX staging package");
expectIncludes(wordAdapter, "FileName:=nativeDocumentPath", "Word native conversion must use the Session's native DOCX path");
expectIncludes(wordAdapter, "insertionRange.FormattedText = stagingEquationRange.FormattedText", "Word native conversion must transfer Word's parsed OMath without flattening it");
expectIncludes(wordAdapter, "Visible:=False", "Word native conversion must keep the DOCX staging document hidden");
expect(!wordAdapter.includes(".InsertXML"), "Word native conversion must avoid Range.InsertXML entirely because Word for Mac raises error 6145");
expect(!wordAdapter.includes("Selection.Paste"), "Word native conversion must not mutate the user's clipboard");
expectIncludes(wordAdapter, "targetDocument.Bookmarks.Exists(VTWordBookmarkName(sessionId))", "Word create commits must recover the pending target through its owning document Bookmark");
expect(wordAdapter.indexOf("targetDocument.Bookmarks.Exists(VTWordBookmarkName(sessionId))") < wordAdapter.indexOf("Set targetImage = VTFindUniqueInlineShape(pendingMarker)"), "Word create commits must try the O(1) pending Bookmark before scanning all InlineShapes");
expectIncludes(wordAdapter, "If target Is Nothing Then Set target = VTFindUniqueInlineShape(pendingMarker)", "Word cancellation must scan InlineShapes only when the pending Bookmark cannot resolve the placeholder");
expectIncludes(wordAdapter, "VTDeletePendingBookmark targetDocument, sessionId", "Word commits must delete pending Bookmarks from the captured owning document");
expectIncludes(wordAdapter, "If documentObject.Bookmarks.Exists(name) Then", "Pending Bookmark deletion must not depend on a transient ActiveDocument");
expectIncludes(wordAdapter, "VTTraceWordSession sessionId", "Word must retain opt-in host-level placeholder identity diagnostics");
expectIncludes(wordAdapter, "Private Const VT_WORD_TRACE_ENABLED As Boolean = False", "Word host tracing must default off to avoid full InlineShapes enumeration and log rewrites on every operation");
expectIncludes(wordAdapter, "If Not VT_WORD_TRACE_ENABLED Then Exit Sub", "Disabled Word tracing must return before touching the document or log");
expectIncludes(wordAdapter, "VTValidateOmmlFragment ommlXml", "Word must validate structural OMML before inserting it");
expect(!wordAdapter.includes("targetRange.Document.OMaths.Add(insertionRange)"), "Word native conversion must not recreate formulas through the broken UnicodeMath linear path");
expectIncludes(wordAdapter, "If displayMode = \"block\" And Not numbered Then", "Every unnumbered display OMML create or edit must begin with a safe inline transaction Range");
expectIncludes(wordAdapter, "Set nativeEquationRange = VTResolveNativeEquationRange", "Word must re-resolve OMath after deleting an adjacent source object");
expectIncludes(wordAdapter, "VTPromoteNativeEquationToDisplay", "Unnumbered display OMML must become display math only after state storage and source-object removal");
expectIncludes(wordAdapter, "Private Function VTOMathTableVisualAdvance", "The real-host OMML size regression must compare inline and display math in identical table structures");
expectIncludes(wordAdapter, "VTConfigureNumberedDisplayTable layoutTable", "The OMML size comparison must reuse the production 20/60/20 table geometry");
expectIncludes(wordAdapter, "inlineTableAdvance = VTOMathTableVisualAdvance", "The real-host regression must measure inline OMML inside the shared table layout");
expectIncludes(wordAdapter, "displayTableAdvance = VTOMathTableVisualAdvance", "The real-host regression must measure display OMML inside the shared table layout");
expectIncludes(wordAdapter, 'regressionStage = "native-display-integral-geometry"', "The real-host regression must cover integral display geometry");
expectIncludes(wordAdapter, 'regressionStage = "native-display-sum-fraction-geometry"', "The real-host regression must cover n-ary and fraction display geometry together");
expectIncludes(wordAdapter, "Private Function VTNativeDocxCompactDisplayAdvance", "Complex display fixtures must reuse the validated compact-tail display strategy without the overflowing nested measurement path");
expectIncludes(wordAdapter, 'resultLine = VTProbeOneDisplayStrategy( _', "Complex display fixtures must execute the real compact-tail strategy inside Word");
expectIncludes(wordAdapter, 'Split(resultLine, "|")', "Complex display fixture metrics must be parsed from the successful batch-probe record");
expectIncludes(wordAdapter, 'formulaAdvance = Val(CStr(resultFields(11)))', "Complex display geometry parsing must use Double-safe Val conversion rather than the overflowing Single path");
expectIncludes(wordAdapter, 'numbered OMML lost its two-paragraph m:oMathPara cell.', "Every numbered OMML layout assertion must verify the authentic two-paragraph display cell");
expectIncludes(wordAdapter, 'numbered OMML display tail is not the compact 1pt paragraph.', "Every numbered OMML layout assertion must verify the compact 1pt tail paragraph");
expectIncludes(wordAdapter, "VTSetWordOmmlPayload testDocument, nativeFormulaId, ommlBase64", "The direct native-number regression fixture must persist the structural OMML payload before display materialization");
expectIncludes(wordAdapter, "VTSetWordOmmlPayload testDocument, conversionFormulaId, ommlBase64", "The image-to-native regression fixture must persist the structural OMML payload before preserving its number");
expectIncludes(wordAdapter, "VTSetWordOmmlPayload testDocument, stabilityNativeFormulaId, ommlBase64", "The continuous native-number regression fixture must persist its own structural OMML payload");
expectIncludes(wordAdapter, "displayTableAdvance <= inlineTableAdvance + 1!", "Display OMML must occupy more vertical space than inline OMML under identical table conditions");
expect(!wordAdapter.includes("Native display OMML did not produce a larger Word line box"), "The OMML size regression must not compare an ordinary paragraph with a table row");
expect(!wordAdapter.includes("hostWindow.GetPoint"), "The Mac Word regression must not call the unsupported Window.GetPoint API");
expectIncludes(wordAdapter, "If pendingPlaceholderRemoved Then", "Failed deferred display insertion must restore its pending transaction target");
expectIncludes(wordAdapter, "VTFinalizeInlineNativeEquation", "Inline OMML must be forced back to wdOMathInline after deleting an adjacent source object");
expectIncludes(wordAdapter, "Start:=exactEquationRange.End, End:=exactEquationRange.End", "Inline OMML caret placement must begin at the exact OMath boundary");
expectIncludes(wordAdapter, "Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove", "Word for Mac inline OMML caret placement must explicitly leave the math zone");
expectIncludes(wordAdapter, "Selection.TypeText Text:=ChrW(8288)", "Inline OMML must create a replaceable ordinary-text anchor after leaving OMath");
expectIncludes(wordAdapter, "anchorRange.OMaths.Count <> 0", "The inline OMML text anchor must be verified outside the math zone");
expectIncludes(wordAdapter, 'regressionStage = "inline-existing-assert"', "The real-host regression must compare empty-paragraph and existing-text inline OMML paths");
expectIncludes(wordAdapter, "nativeEquation.Type = wdOMathInline", "Word must undo its automatic empty-paragraph display promotion before normalizing inline alignment");
expectIncludes(wordAdapter, "VTNormalizeInlineNativeParagraphAlignment", "Inline OMML must normalize an otherwise empty paragraph away from inherited display centering");
expectIncludes(wordAdapter, "Public Sub VisualTeX_RunWordDisplayStrategyProbe()", "Word must expose a batch display-strategy probe before another production implementation is selected");
expectIncludes(wordAdapter, '"/word-display-strategy-probe-result.txt"', "The batch display probe must persist an incrementally readable result file");
expectIncludes(wordAdapter, '"state=RUNNING"', "The batch display probe must publish its running state before testing strategies");
expectIncludes(wordAdapter, '"state=COMPLETE"', "The batch display probe must publish a distinct completion marker");
expectIncludes(wordAdapter, '"formulaAdvance|anchorSpan|fontSize|linearLength|description"', "The batch display probe must report both cell-relative and fixed-anchor geometry");
expectIncludes(wordAdapter, '"fraction", "integral", "sum_fraction"', "The batch display probe must cover simple fractions, integrals and n-ary fraction structures");
expectIncludes(wordAdapter, '"inline-baseline"', "The batch display probe must measure an inline table baseline for every fixture");
expectIncludes(wordAdapter, 'probeStage = "create-before-anchor"', "The batch display probe must create a non-math anchor before every measured table");
expectIncludes(wordAdapter, 'Name:="VTProbeBeforeTable"', "The batch display probe must retain its fixed geometry anchor across source moves");
expectIncludes(wordAdapter, '"formatted-paragraph"', "The batch display probe must retain the current same-document FormattedText route as a measured baseline");
expectIncludes(wordAdapter, '"formatted-paragraph-compact-tail"', "The batch display probe must test the production compact-tail strategy");
expectIncludes(wordAdapter, '"copy-paste-paragraph"', "The batch display probe must test a complete display-paragraph clipboard transfer");
expectIncludes(wordAdapter, '"paste-original-paragraph"', "The batch display probe must test original-format paragraph paste");
expectIncludes(wordAdapter, '"paste-rtf-paragraph"', "The batch display probe must test RTF paragraph transfer");
expectIncludes(wordAdapter, '"cut-paste-paragraph"', "The batch display probe must test moving the complete display paragraph");
expectIncludes(wordAdapter, '"copy-paste-equation"', "The batch display probe must compare paragraph and equation-only clipboard transfer");
expectIncludes(wordAdapter, '"linearize-display-before-build"', "The batch display probe must test rebuilding structural OMML directly inside the empty center cell");
expectIncludes(wordAdapter, '"linearize-build-before-display"', "The batch display probe must test both Word BuildUp and display promotion orders");
expectIncludes(wordAdapter, "sourceXml = VTProbeRangeWordOpenXml(sourceParagraph)", "The batch display probe must inspect the source paragraph's real Word Open XML");
expectIncludes(wordAdapter, "cellXml = VTProbeRangeWordOpenXml", "The batch display probe must inspect the center cell's real Word Open XML");
expectIncludes(wordAdapter, 'VTProbeSubstringCount(cellXml, "<m:oMathPara")', "The batch display probe must detect a true m:oMathPara instead of trusting OMath.Type alone");
expectIncludes(wordAdapter, "VTWriteTextAtomic resultPath, report", "The batch display probe must preserve progress after every strategy");
expectIncludes(wordAdapter, "Public Sub VisualTeX_RunWordNativeRegression()", "The packaged Word add-in must expose a real-host native equation regression entry point");
expectIncludes(wordAdapter, "VTEquationNumberCrossReferenceItems", "The Word cross-reference picker must expose parenthesized items from Word's native Equation list");
expectIncludes(wordAdapter, "GetCrossReferenceItems(wdCaptionEquation)", "The Word cross-reference picker must use the built-in Equation caption registry");
expectIncludes(wordAdapter, "Private Sub VTAssertNumberedEquationLayout", "The real-host regression must measure numbered formula centering and right alignment");
expectIncludes(wordAdapter, "wdHorizontalPositionRelativeToTextBoundary", "The Word numbering regression must measure actual formula and number positions inside the text column");
expectIncludes(wordAdapter, 'regressionStage = "native-numbered-edit"', "The real-host regression must verify numbered OMML editing preserves the layout and number");
expectIncludes(wordAdapter, 'regressionStage = "image-to-native-number-preservation"', "The real-host regression must verify image-to-OMML conversion preserves number geometry");
expectIncludes(wordAdapter, 'assertionName & ": Equation number parentheses are incomplete."', "The Word numbering regression must reject missing number parentheses");
expectIncludes(wordAdapter, "Not VTWordRangeHasMeaningfulText(beforeRange)", "Inline OMML paragraph alignment must change only when there is no surrounding meaningful text");
expectIncludes(wordAdapter, "Set targetDocument = ActiveDocument", "Word commits must capture the owning document before opening hidden staging DOCX files");
expectIncludes(wordAdapter, "targetDocument.Activate", "Word commits must reactivate the owning document after hidden DOCX staging changes ActiveDocument");
expectIncludes(wordAdapter, "VTSetWordLatexPayload targetDocument, formulaId, latexBase64", "Word commits must persist a formula-id keyed LaTeX edit payload in the owning document");
expectIncludes(wordAdapter, "VTSetWordOmmlPayload targetDocument, formulaId, ommlBase64", "Word commits must persist a formula-id keyed structural OMML payload in the owning document");
expectIncludes(wordAdapter, "Set layoutTable = VTEnsureNumberedDisplayTable", "Existing numbered formulas must reuse and refresh the stable display table");
expectIncludes(wordAdapter, "VTNormalizeEquationNumberLayoutWithField _", "New Equation captions must normalize from a stable field-position anchor rather than retaining a stale Word Field object");
expectIncludes(wordAdapter, "Private Function VTResolveEquationSequenceFieldNear", "Equation layout must re-resolve Word SEQ fields after every structural Range edit");
expectIncludes(wordAdapter, "Set sequenceField = VTResolveEquationSequenceFieldNear", "Equation numbering must refresh invalidated Word Field COM objects before reading boundaries or formatting results");
expectIncludes(wordAdapter, "fieldAnchor = VTEquationFieldStart(sequenceField)", "Equation creation must capture a stable field-position anchor after the registered field is moved into place");
expectIncludes(wordAdapter, "insertionParagraphStart = insertionRange.Paragraphs(1).Range.Start", "Equation caption insertion must remember the formula paragraph before Word creates the native field");
expectIncludes(wordAdapter, "Set fieldParagraphRange = _\n        sequenceField.Result.Paragraphs(1).Range.Duplicate", "Equation SEQ insertion must verify that Word retained the native field in its helper paragraph");
expectIncludes(wordAdapter, 'captionStage = "verify-native-seq"', "Equation numbering must validate the direct native SEQ field before returning it");
expect(!wordAdapter.includes("insertionRange.InsertCaption _"), "Numbering must not mutate the formula layout through InsertCaption");
expect(!wordAdapter.includes("destinationRange.FormattedText = fieldOuterRange.FormattedText"), "Equation numbering must not duplicate a native field that Word already inserted at the collapsed formula Range");
expect(!wordAdapter.includes("captionParagraphRange.Delete"), "Equation numbering must not delete the formula paragraph while removing a presumed generated caption paragraph");
expect(!wordAdapter.includes('captionBookmarkName = "VT_TMP_EQ_"'), "Equation field insertion must not use temporary migration bookmarks");
expect(!wordAdapter.includes("fieldBackupDocument"), "Equation numbering must not transfer Word fields through a temporary document");
expect(!wordAdapter.includes("fieldOuterRange.Cut"), "Equation numbering must not rely on clipboard field cut/paste on Word for Mac");
expect(!wordAdapter.includes('operationStage = "remove-caption-prefix"'), "Equation creation must never merge caption paragraphs by deleting across a Word field boundary");
expectIncludes(wordAdapter, "Private Function VTEquationFieldStart", "Equation layout must expose the outer start boundary of a Word field");
expectIncludes(wordAdapter, "Private Function VTEquationFieldEnd", "Equation layout must expose the outer end boundary of a Word field");
expect(!wordAdapter.includes("End:=sequenceField.Result.Start"), "Equation layout must never edit a Range that ends inside a Word field result");
expectIncludes(wordAdapter, "documentObject.Variables", "Word conversion payloads must stay inside the owning Word document");
expectIncludes(wordAdapter, "nativeEquation.Range.Delete", "Word native conversion rollback must remove a partially inserted structural equation");
expectIncludes(wordAdapter, "If replaceTarget Or", "Word native replacement must resolve the OMath at the replaced target boundary");
expectIncludes(wordAdapter, "If Not replaceTarget Then insertionRange.Collapse wdCollapseStart", "Word native edits must replace the exact original OMath Range instead of inserting beside it");
expect(!wordAdapter.includes("originalNativeMath.Range.Delete"), "Word replacement must never delete a stale original OMath COM Range after insertion");
expect(!wordAdapter.includes("candidate.Range.End + originalNativeLength"), "Word replacement must not reconstruct a deletion Range from stale OMath length arithmetic");
expectIncludes(wordAdapter, "VTSetNativeFormulaBookmark", "Word native formulas must retain a persistent VisualTeX identity bookmark");
expectIncludes(wordAdapter, "VTSetWordMetadataPayload", "Word native formulas must retain their complete VisualTeX edit metadata");
expect(!wordAdapter.includes("VTWordConvertNativeBookmarkToImage"), "Word conversion must remain one-way from a VisualTeX image to native OMML");
expectIncludes(wordEvents, "VTIsVisualTeXNativeSelection", "Word native VisualTeX formulas must support double-click editing");
expect(!wordAdapter.includes("VTNativeMathForBookmark(nativeBookmark) Is Nothing"), "Word VBA must assign object-returning functions before testing Is Nothing");
expect(!wordAdapter.includes("If Not VTNativeMathForBookmark(candidate) Is Nothing Then"), "Word VBA must avoid ambiguous Not/function-call/Is Nothing expressions");
expectIncludes(wordAdapter, "Word did not persist the VisualTeX formula properties", "Word must verify the candidate before deleting the old formula");
expectIncludes(wordAdapter, "transactionErrorNumber = Err.Number", "Word rollback must preserve the original transaction error");
expectIncludes(wordAdapter, "VTWriteWordFailureTrace", "Word transaction failures must record their exact stage without enabling expensive full tracing");
expectIncludes(wordAdapter, "errorNumber = Err.Number", "Word creation cleanup must preserve the original error number");
expectIncludes(wordAdapter, "VTShowError \"Word formula creation\", errorNumber, errorDescription", "Word creation errors must survive placeholder cleanup");
expectIncludes(wordAdapter, "If Not insertedNumber Is Nothing Then insertedNumber.Delete", "Word rollback must remove a partially inserted equation number");
expectIncludes(wordAdapter, "VTFindCommittedInlineShape", "Word retries must recognize an already committed Session result");
expectIncludes(wordAdapter, "sourceDocumentId <> VTWordDocumentIdentity()", "Word callback must reject document switching");
expectIncludes(wordAdapter, "Private Function VTWordBookmarkName", "Word pending Bookmarks must use one bounded name generator");
expectIncludes(wordAdapter, "Len(VTWordBookmarkName) > 40", "Word Bookmark names must be guarded by the host length limit");
expectIncludes(powerpointAdapter, "Public Sub Auto_Open()", "PowerPoint add-in must publish Auto_Open health");
expectIncludes(powerpointAdapter, "VisualTeX_DoubleClickEditSelected", "PowerPoint must expose a non-modal native double-click macro entry point");
expectIncludes(powerpointAdapter, "VTInitializePowerPointEvents", "PowerPoint Auto_Open must initialize its persistent application event sink");
expectIncludes(powerpointEvents, "App_WindowBeforeDoubleClick", "PowerPoint must use its native application event for double-click editing");
expectIncludes(powerpointEvents, "Cancel = True", "PowerPoint must suppress the default double-click action for a VisualTeX formula");
expectIncludes(powerpointEvents, "VisualTeX_EditShape", "PowerPoint double-click editing must preserve the clicked Shape target");
expectIncludes(powerpointAdapter, "VisualTeXFormulaId", "PowerPoint add-in must persist formulaId tags");
expectIncludes(powerpointAdapter, "VisualTeXSessionId", "PowerPoint add-in must persist sessionId tags");
expectIncludes(powerpointAdapter, "VisualTeXPending", "PowerPoint add-in must persist pending tags");
expectIncludes(powerpointAdapter, "original.Delete", "PowerPoint replacement must delete the old shape last");
expectIncludes(powerpointAdapter, "candidate.ZOrderPosition <> targetZOrder + 1", "PowerPoint must verify z-order before deleting the old shape");
expectIncludes(powerpointAdapter, 'candidate.Tags("VisualTeXSessionId") <> sessionId', "PowerPoint must verify durable Session tags before deleting the old shape");
expectIncludes(powerpointAdapter, "VTIsCommittedPowerPointShape", "PowerPoint retries must recognize an already committed Session result");
expectIncludes(powerpointAdapter, "VTRestoreZOrder candidate, targetZOrder + 1", "PowerPoint replacement must preserve z-order transactionally");
expectIncludes(powerpointAdapter, "VisualTeX PowerPoint SVG result is missing", "PowerPoint must require the vector SVG export instead of silently rasterizing formulas");
expectIncludes(powerpointAdapter, "PowerPoint could not insert the VisualTeX SVG", "PowerPoint must report an explicit vector insertion failure");
expectIncludes(powerpointAdapter, "fallbackImagePath", "PowerPoint may retain PNG only as a compatibility fallback for Office builds without SVG support");
expect(!wordAdapter.includes('Format$(Now, "yyyy-mm-dd\\Thh:nn:ss") & "Z"'), "Word health must not label local time as UTC");
expect(!powerpointAdapter.includes('Format$(Now, "yyyy-mm-dd\\Thh:nn:ss") & "Z"'), "PowerPoint health must not label local time as UTC");

expectIncludes(officePaths, "Library/Application Scripts/com.microsoft.Word", "Word Session and placeholder files must use Word's Application Scripts directory");
expectIncludes(officePaths, "Library/Application Scripts/com.microsoft.Powerpoint", "PowerPoint Session files must use PowerPoint's Application Scripts directory");
expectIncludes(officePaths, "VT_RUNTIME_DIRECTORY_NAME", "Each Office host must isolate VisualTeX runtime files in a dedicated subdirectory");
expectIncludes(officePaths, 'InStr(1, homePath, "/Library/Containers/", vbTextCompare)', "VBA paths must detect an Office sandbox HOME value");
expectIncludes(officePaths, "homePath = Left$(homePath, sandboxMarker - 1)", "VBA paths must recover the real user home for Application Scripts");
expectIncludes(officePaths, "Application.Name", "VBA runtime paths must select the current Word or PowerPoint host explicitly");
expect(!officePaths.includes("/private/tmp"), "VBA runtime paths must not use a temporary directory blocked by the Office sandbox");
expect(!officePaths.includes("UBF8T346G9.Office/VisualTeX"), "VBA runtime paths must not use Microsoft's protected application-group Data Vault");
expectIncludes(protocol, "VisualTeXPlaceholder.png", "Word must load its transparent placeholder from the persistent Application Scripts directory");
expectIncludes(protocol, "New Collection", "VBA protocol must use the Mac-compatible Collection type");
expect(!protocol.includes("Scripting.Dictionary"), "VBA protocol must not depend on Windows Scripting Runtime");
expectIncludes(protocol, "If Not VT_RANDOM_READY Then", "UUID generation must seed VBA randomness only once per host process");
expectIncludes(protocol, "VT_UUID_COUNTER", "UUID generation must mix a monotonic per-process counter");
expect(!protocol.includes("LenB(StrConv(json, vbFromUnicode))"), "Request sizing must use UTF-8 bytes instead of the host code page");
expect(!protocol.includes("Open temporary For Binary Access Write"), "VBA must not write runtime files directly through the Office sandbox");
expectIncludes(protocol, "VTUtf8Encode", "VBA protocol must provide strict UTF-8 encoding");
expectIncludes(protocol, "VTUtf8Decode", "VBA protocol must provide strict UTF-8 decoding");
expectIncludes(protocol, "Public Function VTBase64UrlDecodeUtf8", "VBA protocol must decode the Word-only LaTeX payload without external dependencies");
expectIncludes(protocol, "VTBase64UrlEncodeUtf8", "VBA must encode runtime file payloads for AppleScriptTask transport");
expectIncludes(protocol, 'VTFileBridgeCall("WriteVisualTeXFile"', "VBA runtime writes must use the fixed AppleScriptTask file bridge");
expectIncludes(protocol, "WriteVisualTeXFile creates the Session parent directory atomically", "Request writes must avoid a redundant directory-creation AppleScriptTask round trip");
expectIncludes(protocol, 'VTFileBridgeCall("ReadVisualTeXFile"', "VBA runtime reads must use the fixed AppleScriptTask file bridge");
expectIncludes(protocol, 'VTFileBridgeCall("EnsureVisualTeXDirectory"', "VBA runtime directory creation must use the fixed AppleScriptTask file bridge");
expectIncludes(protocol, 'VTFileBridgeCall("VisualTeXFileExists"', "VBA runtime existence checks must use the fixed AppleScriptTask file bridge");
expectIncludes(protocol, "VTRuntimeRelativePath", "VBA must reduce every bridged path to a validated runtime-relative path");
expectIncludes(protocol, "Public Function VTProtocolSelfTest() As Boolean", "VBA protocol must expose an actual host-runtime UUID/UTF-8 self-test");
expectIncludes(protocol, "Public Function VTParseInvariantDouble", "VBA protocol must parse dot-decimal dispatch values without depending on an Office host locale API");
expectIncludes(wordAdapter, "VTParseInvariantDouble", "Word must use the shared invariant number parser");
expectIncludes(powerpointAdapter, "VTParseInvariantDouble", "PowerPoint must use the shared invariant number parser");
expect(!wordAdapter.includes("Application.DecimalSeparator"), "Word VBA must not reference the Excel-only Application.DecimalSeparator property");
expect(!powerpointAdapter.includes("Application.DecimalSeparator"), "PowerPoint VBA must not reference the Excel-only Application.DecimalSeparator property");
expectIncludes(launcher, "AppleScriptTask", "VBA launcher must use AppleScriptTask");
expectIncludes(wordScript, "NSWorkspace's sharedWorkspace())'s openURL:targetURL", "Word AppleScriptTask must open the validated Session URL without spawning a shell");
expectIncludes(powerpointScript, "NSWorkspace's sharedWorkspace())'s openURL:targetURL", "PowerPoint AppleScriptTask must open the validated Session URL without spawning a shell");
expect(!wordScript.includes('do shell script "/usr/bin/open " & quoted form of visualTeXURL'), "Word Session launch must not spawn /usr/bin/open");
expect(!powerpointScript.includes('do shell script "/usr/bin/open " & quoted form of visualTeXURL'), "PowerPoint Session launch must not spawn /usr/bin/open");
expect(!wordScript.includes("System Events"), "Word AppleScriptTask must not use UI automation");
expect(!powerpointScript.includes("System Events"), "PowerPoint AppleScriptTask must not use UI automation");
for (const [host, script] of [["Word", wordScript], ["PowerPoint", powerpointScript]]) {
  expectIncludes(script, "validateRelativePath", `${host} file bridge must validate every runtime-relative path`);
  expectIncludes(script, "absoluteRuntimePath", `${host} file bridge must join paths only beneath its fixed runtime root`);
  expectIncludes(script, "writeToFile:targetPath atomically:true", `${host} file bridge must atomically persist decoded runtime data`);
  expectIncludes(script, "createDirectoryAtPath:targetPath withIntermediateDirectories:true", `${host} file bridge must create runtime directories without spawning mkdir`);
  expectIncludes(script, "setAttributes:fileAttributes ofItemAtPath:targetPath", `${host} file bridge must apply runtime file permissions without spawning chmod`);
  expectIncludes(script, "initWithBase64EncodedString", `${host} file bridge must decode Base64URL payloads without shell interpolation`);
  expectIncludes(script, 'candidate contains ".."', `${host} file bridge must reject traversal components`);
  expectIncludes(script, 'quoted form of targetPath', `${host} fixed maintenance commands must shell-quote validated paths`);
  expect(!script.match(/sh -c/i), `${host} AppleScriptTask must not invoke an arbitrary shell program string`);
}

const offlineRuntimeSources = [
  wordAdapter,
  wordEvents,
  powerpointAdapter,
  powerpointEvents,
  protocol,
  officePaths,
  launcher,
  wordScript,
  powerpointScript,
].join("\n").toLowerCase();
for (const forbidden of ["office.js", "https://", "http://", "trusted catalog", "certificate", "webview"]) {
  expect(!offlineRuntimeSources.includes(forbidden), `Offline Office plug-in runtime contains forbidden dependency marker: ${forbidden}`);
}
expect(!wordRibbon.includes("SourceLocation") && !powerpointRibbon.includes("SourceLocation"), "Offline Ribbon XML must not declare a web source location");

expectIncludes(rustRuntime, "visualtex://office/open?session=", "Tauri runtime must accept the fixed Office URL");
expectIncludes(rustRuntime, "create_external", "Tauri runtime must import the VBA-selected Session id");
expectIncludes(rustRuntime, "deny_unknown_fields", "Offline request JSON must reject unknown fields");
expectIncludes(rustRuntime, "run_vba_callback", "Tauri runtime must return results through the VBA callback");
expectIncludes(rustRuntime, 'join("NativeDocuments")', "Tauri must persist native Word staging DOCX files outside ephemeral Session directories");
expectIncludes(rustRuntime, "atomic_write(&native_document_path, &omml_docx, 0o600)?", "Tauri must materialize each formula's durable native Word staging DOCX before dispatch");
expectIncludes(rustRuntime, 'const RESULT_SVG_FILE: &str = "formula.svg"', "Native PowerPoint formulas must be materialized as SVG files");
expectIncludes(rustRuntime, "materialize_powerpoint_svg(session)?", "PowerPoint commits must insert the vector SVG export");
expectIncludes(rustRuntime, "decode_svg", "PowerPoint SVG exports must be validated before Office receives them");
expectIncludes(wordAdapter, 'VTApplicationSupportRoot() & "/NativeDocuments/" & formulaId & ".docx"', "Word image-to-OMML conversion must resolve the same durable formula-scoped staging path");
expectIncludes(rustRuntime, "hide_main_window(app)?", "Office formula requests must hide the main VisualTeX workspace");
expectIncludes(rustRuntime, "open_editor_window(app, &session_id)", "Office formula requests must open the dedicated formula editor");
expectIncludes(rustRuntime, "index.html?view=office-formula", "The native Office editor must reuse the stable desktop entry instead of a blank secondary WebView entry");
expectIncludes(read("src/desktop/main.tsx"), 'view === "office-formula"', "The desktop entry must select the dedicated Office formula view from the window query");
expectIncludes(read("src/desktop/main.tsx"), "<OfficeDialogApp />", "The dedicated desktop window must render the Office formula editor");
expectIncludes(rustRuntime, "window.show()", "The dedicated Office formula editor must be explicitly shown");
expectIncludes(rustRuntime, "window.set_focus()", "The dedicated Office formula editor must receive focus");
expectIncludes(rustRuntime, "focus_open_office_editor", "macOS reopen handling must be able to refocus an existing Office editor");
expectIncludes(nativeInteraction, "focus_open_office_editor", "The PowerPoint compatibility double-click monitor must avoid opening duplicate editor windows");
expect(!nativeInteraction.includes(`if crate::office::macos_offline::focus_open_office_editor(&app) {
                    return;
                }
                match crate::office::macos_offline::run_double_click_edit_macro(
                    crate::office::sessions::OfficeHost::Word`), "An unrelated failed Word editor must not swallow a later double-click on another formula");
expectIncludes(nativeInteraction, "run_double_click_edit_macro", "The native monitor must invoke the DOTM/PPAM edit macro directly instead of depending on an Office.js poller");
expectIncludes(rustRuntime, 'run VB macro macro name "VisualTeX_DoubleClickEditSelected"', "The native runtime must call the fixed Word and PowerPoint double-click macro entry point");
expectIncludes(nativeInteraction, "push_word_edit_selected", "The compatibility monitor must preserve a fallback when the Word macro call fails");
expectIncludes(nativeInteraction, "push_powerpoint_edit_selected", "The compatibility monitor must preserve a fallback when the PowerPoint macro call fails");
expect(!nativeInteraction.includes("native_offline_plugin_loaded"), "The compatibility monitor must not disable itself merely because a native plug-in loaded");
expectIncludes(macSettings, '"install_macos_offline_office_addins"', "macOS Settings must install only the native DOTM/PPAM integration");
expectIncludes(macFirstRun, '"install_macos_offline_office_addins"', "macOS first-run setup must install only the native DOTM/PPAM integration");
for (const obsolete of [
  "install_office_integration",
  "repair_office_integration",
  "regenerate_office_certificate",
  "start_office_companion",
  "stop_office_companion",
]) {
  expect(!macSettings.includes(obsolete), `macOS Settings must not expose the obsolete Office.js action ${obsolete}`);
  expect(!macFirstRun.includes(obsolete), `macOS first-run setup must not expose the obsolete Office.js action ${obsolete}`);
}
expect(!macTauriConfig.includes("dist-office-macos"), "The macOS app bundle must not package the obsolete Office.js web bundle");
expect(!platformBundle.includes('run(npm, ["run", "build:office:macos"])'), "The macOS build must not generate an Office.js web bundle");
expectIncludes(lifecycle, "No Office.js bridge or dialog bundle is required", "macOS startup must not require Office.js UI resources");
expectIncludes(capabilities, '"office-native-*"', "Dedicated native Office windows must receive Tauri core permissions");
expectIncludes(capabilities, '"core:window:allow-close"', "Dedicated native Office windows must be allowed to close after a successful commit or cancel");
expectIncludes(dialogApp, "isMacosOfflineTauriTransport()", "Native Office formula editors must avoid Office.js parent messaging");
expectIncludes(dialogApp, 'import("@tauri-apps/api/window")', "Native Office formula editors must control the actual Tauri window");
expectIncludes(dialogApp, "getCurrentWindow().onCloseRequested", "Closing a native formula window must finalize or cancel its Office transaction");
expectIncludes(dialogApp, "close_macos_offline_office_editor_window", "A successful native Office transaction must destroy the real Tauri editor window");
expectIncludes(rustRuntime, "close_other_editor_windows_for_host", "A new Word Session must retire stale failed Word editor windows before opening the selected formula");
expectIncludes(rustRuntime, "session.host == host", "Stale editor cleanup must remain host-specific and must not close PowerPoint editors for a Word request");
expectIncludes(rustRuntime, ".destroy()", "The native Office editor close command must bypass recursive close interception");
expectIncludes(dialogApp, "公式已经插入，但编辑窗口无法自动关闭", "A close failure after a successful native commit must not be reported as an insertion failure");
expect(!dialogApp.includes("无法插入 PowerPoint 公式"), "The shared Office editor must not mislabel Word failures as PowerPoint insertion failures");
expectIncludes(dialogApp, "latex.trim() && autoCommitOnClose", "Closing a non-empty native editor must commit when auto-apply is enabled");
expectIncludes(dialogApp, ": handleCancel()", "Closing an empty native editor must cancel and remove the pending host object");
expectIncludes(dialogMessages, 'typeof ui.messageParent !== "function"', "Office parent messaging must tolerate native Tauri windows without Office.js");
expectIncludes(appRuntime, "initial_office_url", "Cold Office URL launches must be recognized before the main workspace is revealed");
expectIncludes(appRuntime, "if !office::macos_offline::focus_open_office_editor(app)", "macOS reopen must prefer an Office formula editor over the main workspace");
expectIncludes(rustRuntime, "refresh_health_signal", "Tauri status refresh must ask a running Office host for a fresh health signal");
expectIncludes(rustRuntime, 'macro name "AutoExec"', "Word health refresh must call only the fixed AutoExec macro");
expectIncludes(rustRuntime, 'macro name "Auto_Open"', "PowerPoint health refresh must call only the fixed Auto_Open macro");
expectIncludes(
  read("scripts/register_macos_dev_url_handler.mjs"),
  'property devServer : "http://localhost:1420/"',
  "The macOS development URL launcher must target the configured Vite server instead of opening a blank window",
);
expectIncludes(
  read("scripts/register_macos_dev_url_handler.mjs"),
  "on open location visualTeXURL",
  "The macOS development URL launcher must receive URL AppleEvents instead of relying on shell arguments",
);
expectIncludes(
  read("scripts/register_macos_dev_url_handler.mjs"),
  'CFBundleIdentifier", "-string", "com.visualtex.studio.dev-url-handler"',
  "The macOS development URL launcher must use a distinct bundle identifier from the Tauri application",
);
expectIncludes(
  read("scripts/register_macos_dev_url_handler.mjs"),
  "LSSetDefaultHandlerForURLScheme",
  "The macOS development URL launcher must become the default visualtex URL handler instead of leaving a stale app association",
);
expectIncludes(
  read("scripts/register_macos_dev_url_handler.mjs"),
  "legacyDevApp",
  "The macOS development URL launcher must remove the legacy shell-based handler registration",
);
expectIncludes(
  read("scripts/register_macos_dev_url_handler.mjs"),
  "/usr/bin/curl --silent --fail --max-time 1",
  "The macOS development URL launcher must reject a missing Vite server",
);
expectIncludes(
  read("scripts/register_macos_dev_url_handler.mjs"),
  "VisualTeX 开发服务未运行",
  "The macOS development URL launcher must show a clear missing-server diagnostic",
);
expectIncludes(
  read("scripts/tauri_dev.mjs"),
  "com.visualtex.studio.office",
  "Tauri development startup must pause the existing Office background LaunchAgent",
);
expectIncludes(
  read("scripts/tauri_dev.mjs"),
  "stopStaleDevelopmentProcesses",
  "Tauri development startup must remove stale debug instances before acquiring the single-instance lock",
);
expectIncludes(
  read("scripts/tauri_dev.mjs"),
  'join(repositoryRoot, "node_modules", ".bin", "vite")',
  "Tauri development startup must remove stale Vite servers that occupy the fixed development port",
);
expectIncludes(
  read("scripts/tauri_dev.mjs"),
  "setInterval(pauseMacosOfficeBackground, 400)",
  "Tauri development mode must continuously prevent the Office background process from stealing the single-instance lock during hot reload",
);
expectIncludes(
  appRuntime,
  "#[cfg(not(debug_assertions))]",
  "Debug builds must not resume the installed Office LaunchAgent",
);
expectIncludes(rustRuntime, '("latexBase64", latex_base64)', "Word dispatches must carry a base64url LaTeX payload without changing PowerPoint metadata envelopes");
expectIncludes(rustRuntime, "cleanup_session_files_at", "Completed and cancelled Sessions must remove known local request artifacts");
expectIncludes(rustRuntime, "DirectoryNotEmpty", "Session cleanup must preserve unknown files instead of deleting an entire directory recursively");
expectIncludes(rustRuntime, "com.microsoft.Word/VisualTeXRuntime", "The desktop runtime must read Word's Application Scripts Session root");
expectIncludes(rustRuntime, "com.microsoft.Powerpoint/VisualTeXRuntime", "The desktop runtime must read PowerPoint's Application Scripts Session root");
expectIncludes(rustRuntime, "for host in [OfficeHost::Word, OfficeHost::Powerpoint]", "The desktop runtime must search both host-specific Session roots by UUID");
expectIncludes(rustRuntime, "fs::symlink_metadata", "Offline Session requests must reject symbolic-link substitution before reading");
expectIncludes(rustRuntime, "set_mode(&root, 0o700)", "Each host runtime directory must be private to its owner");
expectIncludes(packager, 'kind === "Word" ? "VTWordEvents" : "VTPowerPointEvents"', "The add-in packager must reject binaries missing the double-click event class module");
expectIncludes(installer, "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ", "Installer must use a transparent Word placeholder image");
expectIncludes(installer, "VisualTeX.dotm", "Installer must preserve the fixed Word filename");
expectIncludes(installer, "VisualTeX.ppam", "Installer must preserve the fixed PowerPoint filename");
expectIncludes(installer, "User Content.localized/Startup.localized/Word/VisualTeX.dotm", "Installer must overwrite Word's active Startup DOTM instead of a detached staging copy");
expectIncludes(installer, "VisualTeX/OfficeAddins/VisualTeX.ppam", "Installer must overwrite the fixed PowerPoint path already registered by Office");
expectIncludes(packager, '"UBF8T346G9.Office"', "The packager's macOS install option must synchronize the active Office files");
expectIncludes(installer, 'home.join("Applications")', "Installer must detect per-user Office application installs");
expectIncludes(installer, "restore_backups(&backups)", "Installer must roll back every staged file after a partial failure");
expectIncludes(installer, 'remove_if_exists(&word_health)', "Installer must clear stale Word health after an update");
expectIncludes(installer, 'remove_if_exists(&powerpoint_health)', "Installer must clear stale PowerPoint health after an update");
expectIncludes(installer, "health_is_current", "Installer must validate exact host/version health records");
expectIncludes(installer, "Err(error) => (false, None, Some(error))", "A corrupt health file must degrade one host instead of failing the whole status view");
expectIncludes(installer, "let word_paths = vec![word_path.clone(), word_script.clone(), placeholder.clone()]", "Word installed status must include its active Startup DOTM, AppleScriptTask and placeholder resources");
expectIncludes(installer, "powerpoint_script.clone()", "PowerPoint installed status must include its AppleScriptTask resource");
expectIncludes(installer, 'health.plugin_version.as_deref() == Some(env!("CARGO_PKG_VERSION"))', "Installer must reject stale plug-in health versions");
expectIncludes(installer, "Library/Application Scripts/com.microsoft.Word", "Installer must use Word's AppleScriptTask directory");
expectIncludes(installer, "Library/Application Scripts/com.microsoft.Powerpoint", "Installer must use PowerPoint's AppleScriptTask directory");
expectIncludes(installer, "addins.json", "Installer must require the compiled add-in checksum manifest");
expectIncludes(installer, "word/vbaProject.bin", "Installer must validate the Word VBA project entry");
expectIncludes(installer, "ppt/vbaProject.bin", "Installer must validate the PowerPoint VBA project entry");
expectIncludes(installer, 'validate_vba_module(path, expected_vba_entry, "VTOfficePaths")', "Installer must reject stale add-ins that predate the shared runtime path module");
expectIncludes(installer, "POWERPOINT_MAIN_CONTENT_TYPE", "Installer must reject files named PPAM whose OOXML main type is not a PowerPoint add-in");
expectIncludes(installer, "validate_main_content_type", "Installer must validate the Word template and PowerPoint add-in main content types");
expectIncludes(packager, "expectedModules", "Packager must verify the reviewed VBA module names");
expectIncludes(packager, "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml", "Packager must require a true PowerPoint add-in OOXML main type");
expectIncludes(packager, 'argument("--powerpoint-shell")', "Packager must support rebuilding a valid PPAM shell around a reviewed VBA project");
expectIncludes(packager, '"VTOfficePaths"', "Packager must require the shared runtime path module");
expectIncludes(packager, "customUI/customUI14.xml", "Packager must inject and verify Ribbon XML");
expect(!installer.includes("Microsoft Word.app\").arg"), "Offline installer must not launch Word as an installation success path");
expect(!installer.includes("Microsoft PowerPoint.app\").arg"), "Offline installer must not launch PowerPoint as an installation success path");

expect(!nativeHtml.toLowerCase().includes("office-js"), "Native editor HTML must not load Office.js");
expectIncludes(nativeMain, "desktopOcrTransport", "Native editor must use Tauri OCR transport instead of HTTP Office transport");
expectIncludes(infoPlist, "<string>visualtex</string>", "macOS bundle must register the visualtex URL scheme");

if (process.platform === "darwin") {
  const temp = mkdtempSync(join(tmpdir(), "visualtex-offline-office-smoke-"));
  try {
    execFileSync("/usr/bin/plutil", ["-lint", join(root, "src-tauri", "Info.macos.plist")], {
      stdio: "pipe",
    });
    for (const [name, source] of [
      ["word", join(offline, "word", "VisualTeXWord.scpt")],
      ["powerpoint", join(offline, "powerpoint", "VisualTeXPowerPoint.scpt")],
    ]) {
      execFileSync("/usr/bin/osacompile", ["-o", join(temp, `${name}.scpt`), source], {
        stdio: "pipe",
      });
    }
    notes.push("macOS plist and both AppleScriptTask sources compiled successfully");
  } catch (error) {
    failures.push(`macOS native source compilation failed: ${error.stderr?.toString().trim() || error.message}`);
  } finally {
    rmSync(temp, { recursive: true, force: true });
  }
} else {
  notes.push("AppleScript/plist compilation skipped on non-macOS host");
}

const logDirectory = join(root, "build-logs", "macos-offline");
mkdirSync(logDirectory, { recursive: true });
const logPath = join(logDirectory, "phase-1-5-smoke.log");
const output = [
  `VisualTeX macOS offline Office smoke: ${failures.length === 0 ? "PASS" : "FAIL"}`,
  ...notes.map((note) => `NOTE ${note}`),
  ...failures.map((failure) => `FAIL ${failure}`),
  "",
].join("\n");
writeFileSync(logPath, output, "utf8");
process.stdout.write(output);

if (failures.length > 0) process.exitCode = 1;
