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
  "src-tauri/src/office/omml_batch.rs",
  "src-tauri/src/office/background.rs",
  "src-tauri/src/office/macos_offline_installer.rs",
  "src-tauri/Info.macos.plist",
  "scripts/package_macos_offline_addins.mjs",
  "scripts/document_import_word_integration.mjs",
  "scripts/register_macos_dev_url_handler.mjs",
  "scripts/tauri_dev.mjs",
  "office-native-dialog.html",
  "src/office/native-dialog-main.tsx",
  "src/office/shared/tauriTransport.ts",
  "src/office/documentImport/OfficeDocumentImportApp.tsx",
  "src/office/redraw/WordLatexRedrawApp.tsx",
  "src/office/redraw/wordLatexRedrawParser.ts",
  "src/office/redraw/wordLatexRedrawRenderer.ts",
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
const metadata = read("office/macos-offline/shared/VTMetadata.bas");
const launcher = read("office/macos-offline/shared/VTLauncher.bas");
const wordScript = read("office/macos-offline/word/VisualTeXWord.scpt");
const powerpointScript = read("office/macos-offline/powerpoint/VisualTeXPowerPoint.scpt");
const rustRuntime = read("src-tauri/src/office/macos_offline.rs");
const ommlBatchRuntime = read("src-tauri/src/office/omml_batch.rs");
const backgroundRuntime = read("src-tauri/src/office/background.rs");
const nativeInteraction = read("src-tauri/src/office/powerpoint_native.rs");
const appRuntime = read("src-tauri/src/lib.rs");
const quickOcrRuntime = read("src-tauri/src/quick_ocr.rs");
const installer = read("src-tauri/src/office/macos_offline_installer.rs");
const packager = read("scripts/package_macos_offline_addins.mjs");
const documentImportWordIntegration = read(
  "scripts/document_import_word_integration.mjs",
);
const nativeHtml = read("office-native-dialog.html");
const nativeMain = read("src/office/native-dialog-main.tsx");
const dialogApp = read("src/office/dialog/OfficeDialogApp.tsx");
const formulaRenderArtifacts = read(
  "src/office/shared/formulaRenderArtifacts.ts",
);
const editorWorkspace = read("src/workspace/EditorWorkspace.tsx");
const mathEditor = read("src/editor/MathEditor.tsx");
const workspacePanelPreferences = read(
  "src/workspace/workspacePanelPreferences.ts",
);
const editorStore = read("src/stores/editorStore.ts");
const settingsDialog = read("src/components/SettingsDialog.tsx");
const inputBehaviorMenu = read("src/components/InputBehaviorMenu.tsx");
const updateDialog = read("src/components/UpdateDialog.tsx");
const applicationConfiguration = read(
  "src/runtime/applicationConfiguration.ts",
);
const formulaToolbar = read("src/toolbar/FormulaToolbar.tsx");
const additionalCommands = read("src/autocomplete/additionalCommands.ts");
const toolbarCommandOmmlRegression = read(
  "scripts/toolbar_command_omml_regression.mjs",
);
const styles = read("src/styles.css");
const dialogMessages = read("src/office/dialog/dialogMessages.ts");
const documentImportApp = read(
  "src/office/documentImport/OfficeDocumentImportApp.tsx",
);
const wordLatexRedrawApp = read("src/office/redraw/WordLatexRedrawApp.tsx");
const wordLatexRedrawParser = read("src/office/redraw/wordLatexRedrawParser.ts");
const wordLatexRedrawRenderer = read("src/office/redraw/wordLatexRedrawRenderer.ts");
const mathMlToLatex = read("src/office/redraw/mathMlToLatex.ts");
const tauriTransport = read("src/office/shared/tauriTransport.ts");
const capabilities = read("src-tauri/capabilities/default.json");
const infoPlist = read("src-tauri/Info.macos.plist");
const macSettings = read("src/components/MacOfficeIntegrationSettings.tsx");
const macFirstRun = read("src/components/MacOfficeFirstRunPrompt.tsx");
const desktopApp = read("src/App.tsx");
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
  "VTWordRibbonConvertImage",
  "VTWordRibbonNumbering",
  "VTWordRibbonCrossReference",
  "VTWordRibbonOpen",
  "VTWordRibbonRestoreSelectionOmmlToLatex",
  "VTWordRibbonRestoreSelectionImageToLatex",
  "VTWordRibbonRestoreDocumentOmmlToLatex",
  "VTWordRibbonRestoreDocumentImageToLatex",
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
expectIncludes(wordRibbon, '<dropDown id="VisualTeX.Mac.Word.ImageFontSize"', "Word must expose an image-formula point-size drop-down");
expectIncludes(wordRibbon, 'getItemLabel="VTWordRibbonGetImageFontSizeItemLabel"', "The Word point-size drop-down must show Chinese names with pt values");
expectIncludes(wordRibbon, 'getSelectedItemIndex="VTWordRibbonGetImageFontSizeSelectedIndex"', "The Word point-size drop-down must report the selected formula size");
expectIncludes(wordRibbon, 'onAction="VTWordRibbonApplyImageFontSizePreset"', "The Word point-size drop-down must apply its selected preset immediately");
expect(!wordRibbon.includes('<editBox id="VisualTeX.Mac.Word.ImageFontSize"'), "Word must not require typing an image formula point size");
expect(wordRibbon.indexOf('id="VisualTeX.Mac.Word.Numbering"') < wordRibbon.indexOf('id="VisualTeX.Mac.Word.ImageFontSize"'), "Word image formula size must appear after Update Equation Numbers");
expect(!wordRibbon.includes('idMso="TabHome"'), "Word VisualTeX controls must not be injected into Home");
expectIncludes(powerpointRibbon, 'id="VisualTeX.Mac.PowerPoint.Tab"', "PowerPoint must expose a dedicated VisualTeX Ribbon tab");
expectIncludes(powerpointRibbon, 'label="VisualTeX"', "PowerPoint dedicated Ribbon tab must be labelled VisualTeX");
expectIncludes(powerpointRibbon, 'insertAfterMso="TabInsert"', "PowerPoint VisualTeX tab must be placed after Insert");
expectIncludes(powerpointRibbon, 'onLoad="VTPowerPointRibbonOnLoad"', "PowerPoint Ribbon load must initialize the persistent double-click event sink");
expectIncludes(powerpointRibbon, '<dropDown id="VisualTeX.Mac.PowerPoint.FormulaFontSize"', "PowerPoint must expose an SVG formula point-size drop-down");
expectIncludes(powerpointRibbon, 'getItemLabel="VTPowerPointRibbonGetFormulaFontSizeItemLabel"', "The PowerPoint point-size drop-down must show Chinese names with pt values");
expectIncludes(powerpointRibbon, 'getSelectedItemIndex="VTPowerPointRibbonGetFormulaFontSizeSelectedIndex"', "The PowerPoint point-size drop-down must report the selected SVG formula size");
expectIncludes(powerpointRibbon, 'onAction="VTPowerPointRibbonApplyFormulaFontSizePreset"', "The PowerPoint point-size drop-down must apply its selected preset immediately");
expect(!powerpointRibbon.includes('<editBox id="VisualTeX.Mac.PowerPoint.FormulaFontSize"'), "PowerPoint must not require typing an SVG formula point size");
expect(powerpointRibbon.indexOf('id="VisualTeX.Mac.PowerPoint.Delete"') < powerpointRibbon.indexOf('id="VisualTeX.Mac.PowerPoint.FormulaFontSize"'), "PowerPoint SVG formula size must appear after Delete Selected Formula instead of between the large buttons");
expect(!powerpointRibbon.includes('idMso="TabHome"'), "PowerPoint VisualTeX controls must not be injected into Home");
expectIncludes(metadata, "Public Function VTUnicodeText", "Dynamic Office Ribbon Chinese labels must be generated at runtime instead of relying on VBE source-file encoding");
expectIncludes(metadata, "If codePoint > 32767 Then codePoint = codePoint - 65536", "Dynamic Office Ribbon labels must convert unsigned UTF-16 code units for Mac VBA");
expectIncludes(metadata, "ChrW(codePoint)", "Dynamic Office Ribbon labels must use Unicode code points");
expectIncludes(metadata, "VTFormulaFontPresetCount = 20", "The shared point-size list must include the four larger presets");
for (const marker of [
  "VTUnicodeText(20843, 21495)",
  "VTUnicodeText(20116, 21495)",
  "VTUnicodeText(23567, 22235)",
  "VTUnicodeText(22235, 21495)",
  "VTUnicodeText(23567, 20108)",
  "VTUnicodeText(23567, 19968)",
  "VTUnicodeText(21021, 21495)",
  "Case 16: VTFormulaFontPresetSize = 48#",
  "Case 17: VTFormulaFontPresetSize = 54#",
  "Case 18: VTFormulaFontPresetSize = 72#",
  "Case 19: VTFormulaFontPresetSize = 96#",
]) {
  expectIncludes(metadata, marker, `The shared Office point-size presets must include ${marker}`);
}
expect(!metadata.includes("八号（"), "VBA source must not embed Chinese point-size labels that Mac VBE can mis-encode");
expectIncludes(metadata, "VTFormulaFontPresetIndex", "Word and PowerPoint must share one point-size preset mapping");
expectIncludes(metadata, "VTUnicodeText(33258, 23450, 20041)", "A non-preset formula size must remain visible as a Unicode custom pt value");
expectIncludes(wordAdapter, "VTWordRibbonGetImageFontSizeItemCount", "Word must expose a dynamic point-size drop-down item count callback");
expectIncludes(wordAdapter, "VTWordRibbonApplyImageFontSizePreset", "Word must apply a selected image point-size preset immediately");
expectIncludes(wordAdapter, "VTUnicodeText(28151, 21512, 23383, 21495)", "Word must report mixed selected image formula sizes without source-encoding corruption");
expectIncludes(powerpointAdapter, "VTPowerPointRibbonGetFormulaFontSizeItemCount", "PowerPoint must expose a dynamic point-size drop-down item count callback");
expectIncludes(powerpointAdapter, "VTPowerPointRibbonApplyFormulaFontSizePreset", "PowerPoint must apply a selected SVG point-size preset immediately");
expectIncludes(powerpointAdapter, "VTUnicodeText(28151, 21512, 23383, 21495)", "PowerPoint must report mixed selected SVG formula sizes without source-encoding corruption");
expectIncludes(powerpointAdapter, "VTPrewarmApplication VT_POWERPOINT_HOST", "PowerPoint must prewarm the resident VisualTeX editor during add-in startup");
expectIncludes(powerpointAdapter, "Call VTWriteAndLaunchSession", "PowerPoint create and edit paths must use the shared write-and-launch bridge");
expect(!powerpointAdapter.includes("VTWriteRequest sessionId, requestJson\n    VTLaunchSession"), "PowerPoint must not retain the two-round-trip request and launch path");
expectIncludes(launcher, "VTTryWriteAndLaunchSessionDirect", "macOS Office formula opens must attempt the sandbox-inbox VBA fast path before AppleScriptTask");
expectIncludes(launcher, 'Environ$("HOME")', "The Office fast path must use the host sandbox HOME instead of the normalized user home");
expectIncludes(launcher, '"/Library/Application Support/VisualTeX/FastOpen/" & normalizedHost', "The Office fast path must write only beneath the host sandbox FastOpen inbox");
expectIncludes(launcher, 'InStr(1, requestJson, """operation"":""formula"""', "Only ordinary formula requests may use the sandbox fast-open path");
expectIncludes(launcher, "VT_FAST_OPEN_MAX_REQUEST_BYTES", "Oversized Office requests must fall back to the established AppleScriptTask bridge");
expectIncludes(launcher, "requestBytes = VTUtf8Encode(requestJson)", "The sandbox fast-open request must be written as strict UTF-8 bytes");
expectIncludes(launcher, "Open temporaryPath For Binary Access Write As #handle", "The Office host must write the fast-open request directly inside its own writable sandbox");
expectIncludes(launcher, "Name temporaryPath As requestPath", "The Office fast-open request must become visible through an atomic same-directory rename");
expectIncludes(launcher, "Private VT_OFFICE_RESIDENT_PREWARMED As Boolean", "The direct Office fast path must cache a confirmed resident prewarm");
expectIncludes(launcher, "VTFastOpenResidentReady(normalizedHost)", "The fast path must recover from lost VBA prewarm state by checking the resident heartbeat in the Office sandbox");
expectIncludes(launcher, '"/resident-ready"', "Word and PowerPoint must verify the resident-ready heartbeat inside their own FastOpen inbox");
expect(!launcher.includes("MacScript("), "The Office inbox fast path must not need a synchronous MacScript launch bridge");
expect(!launcher.includes("/usr/bin/open -gj"), "The Office inbox fast path must not rely on LaunchServices Reopen delivery");
expect(!launcher.includes("nohup"), "The Office fast path must not spawn a sandbox-inherited second VisualTeX process that cannot use Tauri's /tmp socket");
expect(!launcher.includes("Shell(VT_VISUALTEX"), "Mac VBA Shell must not be used because it can trigger Word-to-Terminal permission prompts");
expect(!launcher.includes("open location"), "The rejected MacScript custom-URL fast path must not return");
expect(!launcher.includes("&request="), "The rejected inline custom-URL payload must not return");
expectIncludes(launcher, "VTFastOpenRequestClaimed(requestPath, payloadReadyAt)", "The direct Office fast path must confirm that the resident actually claimed the sandbox request before reporting success");
expectIncludes(launcher, "MAX_CLAIM_WAIT_SECONDS As Double = 0.2", "A stale resident heartbeat must fall back quickly instead of falsely reporting a successful open");
expectIncludes(launcher, '"fastPath=inbox-poll;writeMs="', "The direct Office fast path must report sandbox inbox write and resident-claim timing without a process launch round trip");

expectIncludes(wordAdapter, "Public Sub AutoExec()", "Word template must publish AutoExec health");
expectIncludes(wordAdapter, '"word-office-performance-20260801-r87"', "Word health must identify the painted-centre native Office build");
expectIncludes(wordAdapter, "VTInitializeWordEvents", "Word AutoExec must initialize its persistent application event sink");
expectIncludes(wordEvents, "App_WindowBeforeDoubleClick", "Word must use its native application event for double-click editing");
expectIncludes(wordEvents, "App_WindowSelectionChange", "Word must repair a clicked legacy image-number REF through the native selection-change event");
expectIncludes(wordAdapter, "VisualTeX_StabilizeImageEquationNumberSelection", "Word must expose the narrow image-number selection repair entry point");
expectIncludes(wordEvents, "Cancel = VTHandleWordBeforeDoubleClick(Sel)", "Word must suppress the default double-click action only when the shared handler recognizes a VisualTeX formula");
expectIncludes(wordAdapter, "VTVisualTeXInlineShapeAtSelection", "Word double-click editing must resolve a clicked inline formula even when the collapsed selection is adjacent to it");
expectIncludes(wordAdapter, "Public Function VTHandleWordBeforeDoubleClick", "Word image and native double-click events must share one regression-testable target handler");
expectIncludes(wordAdapter, "VTWordOpenResolvedInlineShape", "Word double-click editing must preserve the once-resolved InlineShape, metadata and geometry target");
expectIncludes(wordAdapter, "The normal committed representation is self-contained", "Word's normal double-click path must avoid redundant document-variable recovery reads");
expectIncludes(wordAdapter, "Private Function VTStoredPayloadChunkCount", "Word Apply must address only the payload chunks that actually exist");
expectIncludes(wordAdapter, "For index = chunkCount + 1 To previousChunkCount", "Word Apply must update payload chunks in place and delete only stale trailing chunks");
expect(!wordAdapter.includes("For index = 1 To VT_WORD_PAYLOAD_MAX_CHUNKS"), "Word Apply must not issue 128 blind Document Variable deletes for every metadata write");
expectIncludes(wordAdapter, "VTPrepareWordImageFormulaState", "Word edit opening must resolve image scale state once on the common path");
expectIncludes(wordAdapter, "Public Sub VisualTeX_EditImageField()", "Word must retain the legacy MacroButton edit entry point for old documents during migration");
expectIncludes(wordAdapter, "Public Function VTEnsureVisualTeXImageMacroButton", "Every image commit must normalize legacy field wrappers to one plain InlineShape");
expectIncludes(wordAdapter, 'VT_WORD_IMAGE_MACRO_SCHEMA_VERSION As String = "7"', "The image migration schema must normalize legacy tables/schema-6 direct SEQ into external helpers, restore painted-centre metadata, and install the Return repair contract");
expectIncludes(wordAdapter, 'VT_WORD_NUMBERED_IMAGE_STYLE_NAME As String = _', "Numbered image formulas must use a dedicated Word paragraph style rather than inheriting direct formula formatting on Return");
expectIncludes(wordAdapter, "equationStyle.NextParagraphStyle = wdStyleNormal", "The numbered-image paragraph style must make Word create a Normal paragraph on Return without a SelectionChange repair");
expect(!wordEvents.includes("VTNormalizeEmptyParagraphAfterNumberedImage"), "Return after a numbered image formula must be native Word style behavior, not an event-time repair");
expectIncludes(wordAdapter, "VTMigrateLegacyNumberedImageTables documentObject", "The one-time image migration must flatten legacy numbered-image tables");
expectIncludes(wordAdapter, "layoutTable.ConvertToText", "Legacy numbered image tables must be flattened through Word's native tab-separated table conversion");
expectIncludes(wordAdapter, "VTFindCommittedInlineShapeNearPosition", "Normal image create/edit resolution must stay local instead of scanning every formula image");
expectIncludes(wordAdapter, "If containingField Is Nothing Then\n        Set VTEnsureVisualTeXImageMacroButton = formulaShape", "A current VisualTeX image must remain a plain InlineShape");
expectIncludes(wordAdapter, "containingField.Delete", "Legacy VisualTeX MacroButtons must be removed before restoring the formula image");
expectIncludes(wordAdapter, "Not containingField Is Nothing Then", "The normalized formula image must not remain inside any Word field");
expectIncludes(wordEvents, "App_WindowBeforeDoubleClick", "VisualTeX image editing must rely on Word's native double-click event rather than a display field");
expect(!wordAdapter.includes('VT_WORD_IMAGE_EDIT_MACRO & " X "'), "No active image path may use a visible X field fallback token");
expectIncludes(documentImportWordIntegration, 'process.argv.includes("--create-image")', "The real Word integration must cover the normal new-image formula transaction");
expectIncludes(documentImportWordIntegration, ': "dfdfdf";', "The new-image regression must preserve the user's literal dfdfdf formula instead of a field fallback glyph");
expectIncludes(wordAdapter, "If mode = \"create\" Then pendingPlaceholderRemoved = True", "A failed image creation must restore its original pending placeholder instead of leaving a half-created field");
expectIncludes(wordAdapter, "Public Sub VisualTeX_WriteSelectedScreenBoundsForRegression()", "The physical Word regression must obtain the selected formula's real screen bounds from Word");
expectIncludes(wordAdapter, "ActiveWindow.GetPoint", "Physical double-click coordinates must come from Word rather than a guessed page transform");
expectIncludes(documentImportWordIntegration, "physicallyDoubleClickSelectedWordFormula", "The physical regression must send a real mouse double-click after selecting the formula");
expectIncludes(documentImportWordIntegration, "macos_physical_double_click.swift", "The physical regression must use Quartz mouse events rather than an edit macro");
expectIncludes(documentImportWordIntegration, '"image-inline"', "The physical Word regression must cover inline image formulas");
expectIncludes(documentImportWordIntegration, '"omml-block"', "The physical Word regression must cover display OMML formulas");
expectIncludes(documentImportWordIntegration, "editorReadyFileName", "The physical Word regression must wait for a visible, hydrated editor rather than only a request file");
expectIncludes(wordAdapter, "Public Sub FormatPicture()", "Word must retain a compatibility FormatPicture command without using it as the VisualTeX image route");
expectIncludes(wordAdapter, "WordBasic.FormatPicture", "Ordinary Word pictures must retain their native formatting command");
expectIncludes(wordAdapter, "VisualTeX_ApplyPendingResult", "Word template must expose the native result callback");
expectIncludes(wordAdapter, "VisualTeX_DoubleClickEditSelected", "Word must expose a non-modal native double-click macro entry point");
expectIncludes(wordAdapter, "VTTryFindNativeFormulaBookmarkLocally", "The shared Word double-click handler must restrict native-equation hit testing to the actual clicked OMath neighborhood");
expectIncludes(wordAdapter, "VTTryVisualTeXMetadataShapeAtDoubleClick", "The shared Word double-click handler must validate the clicked image against its actual InlineShape range");
expectIncludes(wordAdapter, "Public Sub VisualTeX_WriteSelectedDoubleClickScreenBounds()", "The Word compatibility fallback must expose the selected image's real screen bounds for physical hit testing");
expectIncludes(wordAdapter, '"handler-native-not-found"', "A Word double-click without a VisualTeX target must be logged and remain a strict no-op");
expectIncludes(wordAdapter, "VisualTeX_CreateNativeInline", "Word must expose direct inline OMML insertion");
expectIncludes(wordAdapter, "VisualTeX_CreateNativeDisplay", "Word must expose direct display OMML insertion");
expectIncludes(wordAdapter, "Public Sub VTWordRibbonOnLoad", "The Word Ribbon onLoad callback must initialize the application event sink in an attached isolation template");
expectIncludes(wordAdapter, "Public Sub VTWordRibbonNativeInline", "The visible OMML inline Ribbon button must have a resolvable callback");
expectIncludes(wordAdapter, "Public Sub VTWordRibbonNativeDisplay", "The visible OMML display Ribbon button must have a resolvable callback");
for (const callback of [
  "VTWordRibbonRedrawSelectionImage",
  "VTWordRibbonRedrawSelectionOmml",
  "VTWordRibbonRedrawDocumentImage",
  "VTWordRibbonRedrawDocumentOmml",
]) {
  expectIncludes(wordAdapter, `Public Sub ${callback}`, `Word LaTeX redraw must expose ${callback}`);
  expectIncludes(wordRibbon, `onAction="${callback}"`, `The Word Ribbon must bind ${callback}`);
}
expectIncludes(wordAdapter, '"word-latex-redraw-20260802-r1"', "Word must retain an auditable LaTeX redraw source revision");
expectIncludes(wordAdapter, "VTLatexRedrawRequestJson = Replace$(", "Word redraw requests must derive their JSON through the VBA-safe operation replacement path");
expectIncludes(wordAdapter, '"""operation"":""documentImport"""', "Word redraw request construction must replace the document-import operation token");
expectIncludes(wordAdapter, '"""operation"":""latexRedraw"""', "Word redraw requests must replace that token with the dedicated latexRedraw operation");
expectIncludes(wordAdapter, "VTWriteRequest sessionId, requestJson", "Word redraw must create the session before writing its source snapshot");
expectIncludes(wordAdapter, "VTWriteLatexRedrawSource sessionId, sourceText", "Word redraw must write an exact source snapshot before launching");
expectIncludes(wordAdapter, "VTResolveLatexRedrawWordOffsets", "Word redraw must resolve parser offsets against the live Word text before rendering and commit");
expectIncludes(wordAdapter, "VTLatexRedrawWordOffsetsMatch", "Word redraw must verify exact source text before accepting an offset mode");
expectIncludes(wordAdapter, "VTBuildWordOffsetsFromUtf16", "Word redraw must retain a supplementary-Unicode compatibility fallback when direct UTF-16 offsets do not match");
expectIncludes(wordAdapter, "VTFindLatexRedrawWordOffsetsByVisibleText", "Word redraw must fall back to Word Find when hidden Field instructions make visible-text offsets diverge from Range coordinates");
expectIncludes(wordAdapter, "Range coordinates include hidden Field instructions", "Word redraw must document why numbered formulas can invalidate raw parser offsets without any user text edit");
expectIncludes(wordAdapter, 'Replace$(searchPrefix, "^", "^^")', "Word redraw visible-text fallback must escape LaTeX superscript carets for Word Find before exact Range verification");
expectIncludes(wordAdapter, "Start:=targetStart + wordStarts(itemIndex)", "Word redraw must use verified Word offsets when validating and replacing source ranges");
expectIncludes(wordAdapter, "For itemIndex = itemCount - 1 To 0 Step -1", "Word redraw must replace exact source ranges from right to left");
expectIncludes(wordAdapter, "preserveParagraphTopology", "Word redraw must distinguish in-place replacement from interactive display-formula continuation behavior");
expectIncludes(read("scripts/word_latex_redraw_office_integration.mjs"), '"document-image-preserves-paragraph-topology"', "The real Word redraw regression must prove that display formulas do not manufacture blank paragraphs");
expectIncludes(wordAdapter, 'undoRecord.StartCustomRecord "VisualTeX Redraw LaTeX Formulas"', "Word redraw must group every replacement into one undo record without relying on VBA source-code page encoding");
expect(!wordAdapter.includes("是否继续？"), "Word redraw must start immediately after the user chooses an output format");
expect(!wordRibbon.includes("开始前会再次确认"), "The Word Ribbon must not promise an extra redraw confirmation dialog");
expectIncludes(wordAdapter, "Word text changed after LaTeX redraw opened.", "Word redraw must still reject stale Word ranges immediately before mutation");
expectIncludes(wordAdapter, "sourceFontSizes(itemIndex)", "Word redraw must preserve the original Word font size for each formula during insertion");
expectIncludes(wordAdapter, "VTResolveWordLatexRedrawFontsDispatch", "Word redraw must capture per-formula source font sizes before rendering");
expectIncludes(wordAdapter, 'Case "latexRedrawPreflight"', "The Word callback must expose the redraw preflight action");
expectIncludes(wordAdapter, "VTJsonNumber(fontSizePt)", "Word redraw preflight must return locale-independent point sizes");
expectIncludes(protocol, "Public Function VTUtf8ByteLength", "The shared protocol must expose exact UTF-8 byte sizing for the 5 MB redraw limit");
expectIncludes(rustRuntime, 'LATEX_REDRAW_SOURCE_FILE: &str = "latex-redraw-source.txt"', "The native runtime must read the fixed redraw source snapshot");
expectIncludes(rustRuntime, "sourceTextBase64", "The native manifest must preserve the exact source text for Word-side verification");
expectIncludes(rustRuntime, 'open_word_latex_redraw_window(app, &session_id)', "Word redraw must route to its dedicated renderer instead of the document importer");
expectIncludes(rustRuntime, 'view=office-word-latex-redraw', "Word redraw must have a dedicated hidden frontend entry point");
expectIncludes(rustRuntime, ".visible(false)", "The Word redraw renderer must stay invisible during successful automatic runs");
expectIncludes(rustRuntime, ".background_throttling(BackgroundThrottlingPolicy::Disabled)", "Hidden redraw rendering must not be suspended by WebKit background throttling");
expectIncludes(wordLatexRedrawApp, "findWindowsWordLatexRedrawSpans", "The redraw frontend must use the Windows-parity span scanner");
expectIncludes(wordLatexRedrawApp, "prepareWindowsStyleWordLatexRedrawItems", "The redraw frontend must use the direct cached redraw renderer");
expectIncludes(wordLatexRedrawApp, 'focusMacosDocumentImportTarget("latexRedraw")', "The redraw renderer must start automatically without a second user click");
expectIncludes(wordLatexRedrawApp, "resolveMacosLatexRedrawFontSizes", "The redraw renderer must capture Word font sizes before rendering like Windows");
expect(!wordLatexRedrawApp.includes("OfficeDocumentImportApp"), "The Word redraw path must not load the document-import UI");
expect(!wordLatexRedrawParser.includes("documentImportParser"), "The Word redraw parser must be independent of the batch-import parser");
expectIncludes(wordLatexRedrawParser, "Direct TypeScript port of Windows WordBulkImportParser.FindFormulaSpans", "The redraw scanner must document its Windows source of truth");
expectIncludes(wordLatexRedrawRenderer, "const templates = new Map", "The redraw renderer must cache duplicate formulas like Windows");
expectIncludes(wordLatexRedrawRenderer, "String(span.fontSizePt)", "The redraw cache key must include each source formula's Word font size like Windows");
expectIncludes(rustRuntime, 'action", "latexRedrawPreflight"', "The native bridge must dispatch a non-mutating Word redraw preflight");
expectIncludes(rustRuntime, "parse_latex_redraw_font_sizes", "The native bridge must validate the complete Word font-size plan");
expect(!documentImportApp.includes("autoRedrawStartedRef"), "The document importer must not own the automatic Word redraw workflow");
expectIncludes(wordAdapter, "Public Sub VTWordRibbonCrossReference", "The visible Equation-reference Ribbon button must have a resolvable callback");
expectIncludes(wordAdapter, "nativeEquation", "Word requests must preserve the direct native-equation intent");
expectIncludes(wordAdapter, "VT_WORD_IMAGE_SCALE_VARIABLE_PREFIX", "Word must persist formula point size and reference image geometry per formula id");
expectIncludes(wordAdapter, "VTPreferredWordFormulaFontSize(Selection.Range.Duplicate)", "New Word formulas must inherit the current selection point size");
expectIncludes(wordEvents, "VisualTeX_SynchronizeSelectedImageFormulaSize Sel", "Word selection changes must synchronize the point-size drop-down with image geometry");
expect(!wordAdapter.includes("currentWordFontSizePt - normalFontSizePt"), "Word image selection must not depend on the document Normal style when detecting a transient selected-InlineShape font value");
expectIncludes(wordAdapter, "Geometry is", "Word image selection must treat persisted image geometry as the authoritative point-size signal");
expectIncludes(wordAdapter, "Abs(formulaShape.Width - expectedWidthPt) <= 0.5", "Word must ignore a transient selected-InlineShape font value while the stored image geometry is unchanged");
expectIncludes(wordAdapter, 'name:="VisualTeX_WatchSelectedImageFormulaSize"', "Word must run a lightweight selected-image point-size watcher when events are insufficient");
expectIncludes(wordAdapter, "formulaShape.Range.Font.Size = CSng(requestedFontSizePt)", "The shared scaling path must attempt the native Word Range.Font.Size property");
expectIncludes(wordAdapter, "formulaShape.Width = CSng(targetWidth)", "Image formula point sizes must map to proportional width");
expectIncludes(wordAdapter, "formulaShape.Height = CSng(targetHeight)", "Image formula point sizes must map to proportional height");
expectIncludes(wordAdapter, "formulaShape.Range.Font.Position = 0", "Display image formula resizing must clear any inherited inline baseline raise");
expectIncludes(wordAdapter, "VTRefreshNumberedImageFormulaFontLayout", "Every numbered image formula resize must refresh number size, mathematical baseline and tab geometry");
expectIncludes(wordAdapter, "Private Function VTExpectedNumberedImageFormulaPosition", "All numbered-image position calculations must share the signed painted-centre geometry path");
expect(!wordAdapter.includes("Private Function VTValidatedWordPositionValue"), "The Word regression must not synchronously read and convert transient Font.Position values");
expectIncludes(wordAdapter, '"The resized image Equation visual-center position"', "Production numbered-image resizing must align the formula object to the visible number line by visual centre");
expect(!wordAdapter.includes("expectedPosition = CLng(Int( _\n        (CDbl(formulaShape.Height) - requestedFontSizePt)"), "Production numbered-image resizing must not duplicate an unchecked CLng position calculation");
expectIncludes(wordAdapter, 'regressionStage = "numbered-image-enter-creates-normal-line"', "The real-host Word regression must prove Return after a numbered image creates a plain Normal typing paragraph");
expectIncludes(wordAdapter, 'regressionStage = "resize-numbered-image-to-24"', "The real-host regression must exercise the production 24-point resize path");
expectIncludes(wordAdapter, "positionAt24 = expectedPositionAt24", "The 24-point regression report must reuse the position already validated by the complete layout assertion");
expectIncludes(wordAdapter, ".Size = CSng(requestedFontSizePt)", "A numbered image formula must apply the selected point size to its visible number");
expectIncludes(wordAdapter, "VTConfigureNumberedEquationParagraph paragraphRange", "Every numbered image resize must rebuild the exact centre and right tab stops");
expectIncludes(wordAdapter, "expectedPosition = VTApplyNumberedImageFormulaVerticalAlignment", "Every numbered image resize must rerun the shared painted-centre algorithm and keep the number at Position 0");
expectIncludes(wordAdapter, "VTCalculateStaticImageEquationNumberPosition = numberPosition", "Field reconciliation must return the validated position without synchronously rereading Font.Position");
expect(!wordAdapter.includes("appliedPositionValue = CDbl(numberRange.Font.Position)"), "Word must not reread Font.Position immediately after applying the number layout");
expectIncludes(wordAdapter, 'regressionStage = "record-image-vertical-after-field-refresh"', "The real-host regression must rely on the preceding complete layout assertion after field refresh");
expectIncludes(wordAdapter, 'regressionStage = "write-success-report"', "The real-host regression must persist PASS before attempting temporary-document cleanup");
expectIncludes(wordAdapter, 'regressionStage = "cleanup-successful-regression"', "Temporary Word document cleanup must be isolated from the successful regression result");
expectIncludes(wordAdapter, "On Error Resume Next\n    If Not testDocument Is Nothing Then\n        testDocument.Close SaveChanges:=wdDoNotSaveChanges", "A Word cleanup overflow must not replace an already written PASS result");
expectIncludes(wordAdapter, "Set candidate = VTAddWordFormulaPicture", "Word must insert formula artwork through the DOCX-staged SVG compatibility helper");
expectIncludes(wordAdapter, "Set stagingDocument = Documents.Open", "Word must use the validated hidden-document SVG transfer path");
expectIncludes(wordAdapter, "insertionRange.FormattedText = stagingShapeRange.FormattedText", "Word must transfer its parsed SVG InlineShape without clipboard or UI automation");
expectIncludes(wordAdapter, 'VTRequireDispatchValue dispatch, "vectorDocumentPath"', "Word image commits must require the generated SVG staging DOCX path");
expectIncludes(wordAdapter, 'fallbackImagePath = VTDispatchOptional', "Word must retain the PNG preview only as an image-formula compatibility fallback");
expectIncludes(wordAdapter, "VTProbeInlineShapeRangeFontSizeBehavior", "Word must expose a real-host probe for InlineShape.Range.Font.Size behavior");
expectIncludes(wordAdapter, "VisualTeX_ConvertSelectedToImageFormula", "Native OMML formulas must support conversion back to an image formula");
expect(!wordAdapter.includes("ElseIf VTWordConvertNativeBookmarkToImageFast(nativeBookmark) Then"), "OMML-to-image conversion must not reuse stale cached image geometry from older font calibrations");
expectIncludes(wordAdapter, 'VTWordOpenNativeSession nativeBookmark, False, "nativeToImage"', "Managed OMML-to-image conversion must use the dedicated silent re-render Session instead of a visible editor or stale cache");
expectIncludes(wordAdapter, "VTWordNativeSignatureMatches", "The retained cached helper must still reject stale image artifacts before any future fast-path reuse");
expectIncludes(wordAdapter, "Set insertionRange = sourceRange.Duplicate", "Any retained cached OMML-to-image helper must replace the exact OMath Range instead of inserting at OMath.Range.End");
expectIncludes(wordAdapter, "Set candidate = VTDetachWordFormulaPictureFromMath(candidate)", "Any retained cached OMML-to-image helper must detach the picture from Word math before committing it");
expectIncludes(metadata, 'Optional ByVal operationName As String = "formula"', "Formula requests must carry an explicit immutable operation discriminator");
expectIncludes(metadata, '"""operation"":" & VTJsonString(operationName)', "The VBA request must serialize the direct conversion operation for the native runtime");
expectIncludes(rustRuntime, '"formula" | "nativeToImage" | "imageToNative"', "The native runtime must validate both dedicated direct-conversion operations");
expectIncludes(rustRuntime, "fn word_image_cache_paths", "Every Word formula export must have a formula-scoped SVG package cache for sub-second OMML-to-image conversion");
expectIncludes(rustRuntime, "let prepared_image_artifacts = materialize_word_svg_package", "Native OMML commits must refresh their future image-conversion cache without changing the current output kind");
expectIncludes(rustRuntime, 'Some("nativeToImage" | "imageToNative")', "The native URL handler must route both direct conversions through the parked renderer");
expectIncludes(rustRuntime, "if !silent {\n        wake_resident_editor_for_hydration", "Only ordinary Office editing may reveal the resident editor during hydration");
expectIncludes(dialogApp, 'session.operation !== "nativeToImage" &&', "The resident frontend must detect silent native-to-image Sessions");
expectIncludes(dialogApp, 'session.operation !== "imageToNative"', "The resident frontend must detect silent image-to-native Sessions");
expectIncludes(dialogApp, "void handleCommit().then", "A silent direct-conversion Session must commit automatically after hydration");
expectIncludes(dialogApp, "await cancelMacosOfflineOfficeSession(session.id)", "A failed direct conversion must cancel without revealing the formula editor");
expect(!dialogApp.includes('"present_macos_offline_office_editor_window"'), "A failed direct conversion must never reveal the resident editor");
expectIncludes(wordAdapter, 'VTWordEditInlineShape target, True, "imageToNative"', "A missing native staging DOCX must be regenerated by a silent image-to-native Session");
expectIncludes(wordAdapter, "If VTWordConvertInlineShapeToNativeFast(selectedShape) Then Exit Sub", "The Ribbon image-to-OMML command must use the narrow cached fast path before any compatibility workflow");
expectIncludes(wordAdapter, "VTWordConvertInlineShapeToNativeEquation selectedShape", "Numbered or legacy image formulas must retain the established compatible in-Word conversion transaction");
expectIncludes(wordAdapter, "The image remains the rollback", "The cached image-to-OMML fast path must keep the source image live until the new OMath is validated");
expectIncludes(wordAdapter, "Private Function VTInsertCachedNativeEquationFast", "Cached image-to-OMML conversion must use a local native insertion helper instead of the generic full-document scanner");
expectIncludes(wordAdapter, "If probeRange.OMaths.Count <> 1 Then", "Cached native insertion must validate exactly one OMath near the insertion point");
expectIncludes(wordAdapter, 'VTWordEditInlineShape target, True, "imageToNative"', "The cached image-to-OMML fast path must preserve the silent hidden-renderer fallback for older formulas without a native DOCX");
expectIncludes(wordAdapter, "Private Function VTRestoreCachedImageFormulasToLatex", "VisualTeX image formulas with durable LaTeX payloads must restore without launching the hidden renderer");
expectIncludes(wordAdapter, "Private Function VTRestoreCachedNativeFormulasToLatex", "Unmodified VisualTeX OMML formulas with durable LaTeX payloads must restore without launching the hidden renderer");
expectIncludes(wordAdapter, "Private Function VTDocumentHasVisualTeXNativeBookmark", "Pure Word-native documents must skip the VisualTeX signature scan before OMML batch conversion");
expectIncludes(wordAdapter, "If Not VTDocumentHasVisualTeXNativeBookmark(targetDocument) Then", "Cached native restoration must exit before enumerating OMath when no VisualTeX native identity exists");
expectIncludes(wordAdapter, "Private Function VTTryRestoreOmmlBatchSynchronously", "Word-native or modified OMML formulas must use the synchronous no-window batch converter before the hidden fallback");
expectIncludes(wordAdapter, "VTConvertOmmlBatch", "The Word VBA bridge must delegate OMML translation to the VisualTeX native converter");
expectIncludes(wordAdapter, "Private Sub VTCommitSynchronousOmmlLatexResult", "Synchronous OMML conversion must use the lightweight transactional Word text replacement");
expectIncludes(wordAdapter, "Private Function VTTryWriteOmmlBatchManifestDirect", "Word must write the ASCII-safe OMML batch manifest directly before using the AppleScript fallback");
expectIncludes(wordAdapter, "VTParseDocumentImportManifestText", "Synchronous OMML conversion must parse the returned manifest in memory instead of rereading it through AppleScriptTask");
expectIncludes(launcher, "Public Function VTConvertOmmlBatch", "The shared Word launcher must expose the synchronous OMML batch bridge");
expectIncludes(launcher, "Optional ByVal restoreSource As String", "The compatibility OMML bridge must accept an already-written runtime manifest without re-encoding it");
expectIncludes(wordScript, "on ConvertOmmlBatch(argumentText)", "The Word AppleScriptTask source must run the no-window OMML batch command");
expectIncludes(wordScript, '"--office-omml-to-latex-batch"', "The Word AppleScriptTask source must invoke the reviewed native CLI argument");
expectIncludes(wordScript, "base64EncodedStringWithOptions:0", "The Word AppleScriptTask source must return the native LaTeX result in the same bridge call");
expectIncludes(wordScript, 'set executablePath to "/Applications/VisualTeX.app/Contents/MacOS/visualtex"', "The synchronous OMML converter must avoid per-call process discovery when the installed app is available");
expectIncludes(rustRuntime, "pub(crate) fn word_omml_to_mathml", "The native batch converter must reuse the reviewed Word OMML-to-MathML transform");
expectIncludes(ommlBatchRuntime, "pub fn run_cli_if_requested", "The native runtime must expose the no-window OMML batch CLI");
expectIncludes(ommlBatchRuntime, "pub fn mathml_to_latex", "The native runtime must translate MathML to LaTeX without opening the frontend");
expectIncludes(wordAdapter, "VTWordNativeSignatureMatches", "Cached native-to-LaTeX restoration must reject formulas modified by the Word native editor");
expectIncludes(wordAdapter, 'formulaStage = "write-native-signature"', "LaTeX redraw to OMML must persist a native structure signature after the final numbered/display structure is stable so a later restore can return the original LaTeX payload");
expectIncludes(wordAdapter, "For itemIndex = formulaCount To 1 Step -1", "Bulk cached formula-to-LaTeX restoration must replace formulas from the end so Word ranges stay stable");
expectIncludes(wordAdapter, "Private Function VTTryResolveFormulaRestoreNumbering", "Formula-to-LaTeX restore must resolve numbered image and native formulas by exact VisualTeX identity before mutation");
expectIncludes(wordAdapter, "Private Function VTFormulaRestoreNumberingScaffoldOwnsRange", "Formula restore must recognize an exact live numbering scaffold even when an older document lost its matching format Variable");
expectIncludes(rustRuntime, "strip_visualtex_numbered_equation_array", "Numbered OMML fallback conversion must remove only the verified VisualTeX eqArr/#/VT_N_ wrapper before Word OMML-to-MathML conversion");
expectIncludes(wordAdapter, "Private Function VTReplaceFormulaRestoreRangeWithLatex", "Every formula-to-LaTeX path must use the shared numbered-structure-aware replacement helper");
expectIncludes(wordAdapter, "legacyNumberTable.Delete", "Formula-to-LaTeX restore must remove a verified legacy 1x3 numbered table instead of leaving an empty scaffold around LaTeX");
expectIncludes(wordAdapter, "Private Sub VTFinalizeFormulaRestoreNumbering", "Bulk formula restore must defer Equation scaffold cleanup until every source formula range has been replaced");
expectIncludes(wordAdapter, "VTDeleteEquationNumberScaffold documentObject, formulaId, False", "Formula-to-LaTeX restore must delete the exact VT_N_/VT_R_/VT_C_ Equation scaffold and hidden SEQ helper");
expectIncludes(wordAdapter, "Prefer the exact SEQ field result when resolving the hidden helper", "Equation scaffold deletion must resolve the hidden helper from VT_N_ before the collapsed VT_C_ paragraph-end fallback");
expectIncludes(wordAdapter, "VTNormalizePlainWordParagraph paragraphRange", "Restored numbered LaTeX must remove the old Caption/tabbed numbering paragraph layout");
expectIncludes(wordAdapter, "VTWriteFormulaRestoreAndLaunchSession", "Small native OMML restore requests must write both payloads and launch VisualTeX in one AppleScriptTask");
expectIncludes(wordAdapter, "VT_WORD_FORMULA_RESTORE_COMBINED_MAX_BYTES", "Large formula restore manifests must retain the chunked compatibility fallback");
expectIncludes(launcher, "Public Function VTWriteFormulaRestoreAndLaunchSession", "The shared Office launcher must expose the atomic formula-restore write-and-launch bridge");
expectIncludes(wordScript, "on WriteFormulaRestoreAndOpenVisualTeXSession", "The Word AppleScriptTask must atomically write request and restore source before launching VisualTeX");
expectIncludes(wordAdapter, "Set selectedShape = VTVisualTeXInlineShapeAtSelection(Selection)", "The image-to-OMML Ribbon command must resolve the selected formula without unsafe collection indexing");
expect(!wordAdapter.includes("VTWordEditInlineShape target, True\n"), "Image-to-native conversion must not fall back to an ordinary visible edit Session");
expectIncludes(wordAdapter, "VTDetachWordFormulaPictureFromMath", "Native-to-image conversion must detach the final picture from Word's OMath container");
expectIncludes(wordAdapter, "VTDeleteTrailingNativeImageArtifact", "Native-to-image conversion must remove only its invisible trailing boundary characters");
expectIncludes(wordAdapter, "trailingText = ChrW(8288)", "Native-to-image cleanup must explicitly remove Word's word-joiner artifact");
expectIncludes(wordAdapter, "contentEnd = paragraphRange.End", "Native-to-image cleanup must resolve the real Word paragraph content boundary");
expectIncludes(wordAdapter, "terminalSuffixRange.Fields.Count = 0", "Native-to-image cleanup must preserve every suffix containing a Word field");
expectIncludes(wordAdapter, "Not VTWordRangeHasMeaningfulText(terminalSuffixRange)", "Only a whitespace-only terminal suffix may be removed after conversion");
expectIncludes(wordAdapter, 'If displayMode = "inline" And mode = "create" Then', "Only a newly created inline native formula may manufacture an after-formula caret anchor");
expectIncludes(wordAdapter, "Editing or direct conversion must not manufacture a U+2060 caret", "Existing formula conversion must preserve surrounding Word text without a synthetic caret character");
expectIncludes(wordAdapter, "finalFormulaRange.Font.Size = CSng(sourceFontSizePt)", "Direct image-to-OMML conversion must preserve the source point size");
const wordCreateStart = wordAdapter.indexOf("Private Sub VTWordCreate(");
const wordCreateEnd = wordAdapter.indexOf("Private Function VTAddWordFormulaPicture", wordCreateStart);
const wordCreateSource = wordAdapter.slice(wordCreateStart, wordCreateEnd);
expectIncludes(wordCreateSource, "VTAddPendingBookmark insertionRange, sessionId", "Word create must use a lightweight pending Bookmark instead of importing a placeholder picture before the editor opens");
expectIncludes(wordCreateSource, "bookmark-anchor-ready", "Word create performance tracing must expose the lightweight anchor stage");
expect(!wordCreateSource.includes("InlineShapes.AddPicture"), "Word create must not synchronously import a 1×1 placeholder image on the editor-open hot path");
expect(!wordCreateSource.includes("VTRefreshWordHealthQuietly"), "Word create must not synchronously rewrite plug-in health before every editor open");
expectIncludes(wordCreateSource, "VTDeletePendingBookmark ActiveDocument, sessionId", "A failed lightweight Word create must delete its pending Bookmark anchor");
expectIncludes(wordAdapter, "Private Function VTPrepareWordCreateInsertionRange", "Display creation must isolate a dedicated paragraph before placing its transaction Bookmark");
expectIncludes(wordAdapter, "Text = vbCr & vbCr", "A display formula inserted between existing content must preserve both sides around an empty display paragraph");
expectIncludes(wordAdapter, "beforeRange.InlineShapes.Count > 0", "Display paragraph isolation must treat an earlier inline image formula as real surrounding content");
expectIncludes(wordAdapter, "beforeRange.Fields.Count > 0", "Display paragraph isolation must not reuse a visually blank hidden Equation helper field");
expectIncludes(wordAdapter, "VTCreateDedicatedPlainParagraphAt", "A visually empty but structurally occupied caret position must receive a dedicated plain display paragraph");
expect(!wordAdapter.includes("VTSeqHelper_"), "Numbered formulas must never expose an internal helper marker in the Word document");
expectIncludes(wordAdapter, "Private Sub VTConfigureNumberedEquationParagraph", "New numbered formulas must use one ordinary Word paragraph with explicit center and right tab stops");
expectIncludes(wordAdapter, "If paragraphRange.Style <> wdStyleCaption Then", "Caption style must be applied only once so later renumbering cannot reset direct formula/number formatting");
expectIncludes(wordAdapter, "Equation number is not vertically stable", "The real-host geometry assertion must report vertical correction mismatches explicitly");
expectIncludes(wordAdapter, "Position:=textWidth / 2!", "The formula tab stop must remain at the exact text-column center");
expectIncludes(wordAdapter, "Alignment:=wdAlignTabCenter", "The formula must remain center-aligned within the single paragraph");
expectIncludes(wordAdapter, "Position:=textWidth - 1!", "The number tab stop must remain at the right text boundary");
expectIncludes(wordAdapter, "Alignment:=wdAlignTabRight", "The number must remain right-aligned within the single paragraph");
expectIncludes(wordAdapter, "Private Sub VTFinalizeParagraphEquationNumber", "Every new image or OMML number must finalize its SEQ and Bookmarks in the same paragraph");
expectIncludes(wordAdapter, "Private Sub VTRefreshParagraphEquationBookmarks", "Single-paragraph numbering must restore VT_N_, VT_R_ and VT_C_ after field updates");
expectIncludes(wordAdapter, "captionRange.Collapse wdCollapseEnd", "The compatibility VT_C_ Bookmark must be collapsed without creating text or another paragraph mark");
expectIncludes(wordAdapter, "Set VTEnsureImageEquationNumber = VTInsertEquationNumber", "New numbered image formulas must use the single-paragraph numbering path");
expectIncludes(wordAdapter, "VTEnsureNativeEquationArrayNumber(formulaRange, formulaId)", "New numbered OMML formulas must use the native single-paragraph Equation-array path");
expectIncludes(wordAdapter, "If formulaRange.Information(wdWithInTable) Then", "Existing three-cell formulas must retain an explicit legacy compatibility branch");
expectIncludes(wordAdapter, "NumRows:=1, NumColumns:=3", "Legacy numbered-table documents must remain readable and repairable");
expectIncludes(wordAdapter, "Private Function VTCalculateStaticImageEquationNumberPosition", "Numbered image Equations must route reconciliation through the reviewed painted-centre formula");
expectIncludes(wordAdapter, "Private Function VTStaticImageEquationNumberRange", "Image Equation numbers must expose a field-free visible number Range");
expectIncludes(wordAdapter, "Private Function VTWriteStaticImageEquationNumber", "Image Equation renumbering must rewrite the visible number from the external SEQ result");
expectIncludes(wordAdapter, 'insertionRange.Text = vbTab & "(" & normalizedNumber & ")"', "Static image-number creation must write the complete visible suffix without a REF field");
expectIncludes(wordAdapter, "VTParagraphHasSingleVisualTeXImageMacroButton", "Image Equation verification must require exactly one plain VisualTeX InlineShape and no Word field");
expectIncludes(wordAdapter, '"The refreshed image Equation visible REF is incomplete."', "The real-host regression must reject a damaged visible image REF after field refresh");
expect(!wordAdapter.includes("sequenceField.Locked = True"), "Image Equation numbering must not lock native SEQ fields and trigger runtime error 4605");
expectIncludes(wordAdapter, "VT_WORD_EQUATION_NUMBER_INK_CENTER_ABOVE_BASELINE_RATIO", "Image-number alignment must use the measured Cambria Math glyph-path centre");
expectIncludes(wordAdapter, "formulaHeight * (1# - inkCenterYRatio)", "Image-number alignment must derive the picture painted centre from cached PNG alpha bounds");
expectIncludes(wordAdapter, "numberSize * _\n        VT_WORD_EQUATION_NUMBER_INK_CENTER_ABOVE_BASELINE_RATIO", "Image-number alignment must scale the number glyph centre by its actual font size");
expectIncludes(wordAdapter, "-CLng(Int((-rawPosition) + 0.5#))", "The shared image-number helper must round negative picture Position away from zero");
expect(!wordAdapter.includes("VTTryMeasureNumberedImageFormulaPosition"), "Production image-number alignment must not mistake Word GetPoint row boxes for painted bounds");
expect(!wordAdapter.includes("VTTryMeasureCurrentNumberedImagePaintedCenterDelta"), "Regression acceptance must not reuse the invalid GetPoint painted-centre probe");
expectIncludes(wordAdapter, "VT_WORD_EQUATION_NUMBER_FONT_NAME As String = \"Cambria Math\"", "Static image Equation numbers must use the same Western math font family as native OMML");
expectIncludes(wordAdapter, ".NameAscii = VT_WORD_EQUATION_NUMBER_FONT_NAME", "Static image Equation number ASCII glyphs must be forced to Cambria Math");
expectIncludes(wordAdapter, ".NameOther = VT_WORD_EQUATION_NUMBER_FONT_NAME", "Static image Equation number punctuation must be forced to Cambria Math");
expect(!wordAdapter.includes(".NameFarEast = VT_WORD_EQUATION_NUMBER_FONT_NAME"), "Mac Word must not assign Cambria Math to its unsupported Far East font slot");
expectIncludes(wordAdapter, "With paragraphRange.Font\n        .Size = visibleSize\n        .Hidden = False\n        .Color = wdColorAutomatic\n        .Name = VT_WORD_EQUATION_NUMBER_FONT_NAME", "The external SEQ helper's complete character source must survive F9 as visible-size Cambria Math");
expectIncludes(wordAdapter, 'regressionStage = "f9-visible-ref-font"', "The chapter regression must hard-check F9 helper and visible REF fonts");
expectIncludes(wordAdapter, 'regressionStage = "verify-word-built-in-private-ref-span"', "The chapter regression must reject a Word-native cross-reference target that escapes the helper paragraph");
expectIncludes(wordAdapter, "baselineRatio = _\n                        referenceBaselinePt / previousReferenceHeightPt", "Directly resizing an image must preserve its SVG baseline-to-height ratio");
expectIncludes(wordAdapter, "referenceBaselinePt = _\n                        referenceHeightPt * baselineRatio", "Manual image geometry changes must preserve the mathematical baseline ratio instead of rebuilding it from an already-rounded Word Position");
expectIncludes(wordAdapter, "If Abs(formulaShape.Width - expectedWidth) > 0.5 Or _", "Direct image geometry changes must be detected independently of Mac Word's unreliable InlineShape.Range.Font.Size report");
expectIncludes(wordAdapter, "reportedWidth = referenceWidthPt * _", "A changed native Word font-size report must be accepted only when the image geometry independently matches that reported size");
expectIncludes(wordAdapter, "VTApplyWordInlineImageBaseline _\n                    formulaShape, referenceHeightPt, referenceBaselinePt", "Inline image synchronization must repair the persisted mathematical baseline before fragile Word font-size reconciliation");
expectIncludes(wordAdapter, "rawPosition = referenceBaselinePt * actualHeightPt / referenceHeightPt", "Inline image baseline mapping must scale the fractional SVG descent against the final Word image height");
expect(!wordAdapter.includes("VT_WORD_INLINE_IMAGE_MATH_AXIS_OFFSET_PT"), "Inline image alignment must not reintroduce a fixed Word/OMML math-axis offset that shifts already-aligned formulas");
expectIncludes(wordAdapter, 'If displayMode = "inline" Then\n        VTApplyWordInlineImageBaseline _', "Inline image baseline calibration must remain separate from block/display paragraph alignment");
expectIncludes(wordAdapter, 'Else\n        formulaShape.Range.Font.Position = 0', "Block image formulas must keep zero Font.Position instead of reusing the inline baseline correction");
expectIncludes(wordAdapter, "VTRefreshNumberedImageFormulaAfterGeometryChange", "Directly resizing a numbered block image must immediately refresh its mathematical baseline and number font");
expectIncludes(wordAdapter, "Public Sub VisualTeX_RunWordImageNumberVerticalAlignmentRegression()", "Word must expose the dedicated 54.25 by 46.10 point image-number vertical-alignment regression");
expectIncludes(wordAdapter, 'regressionStage = "calibrate-before-field-refresh"', "The dedicated image-number regression must measure actual baselines before field refresh");
expectIncludes(wordAdapter, 'regressionStage = "select-static-visible-number"', "The dedicated image-number regression must select the static visible number and verify that its baseline remains stable");
expectIncludes(wordAdapter, 'regressionStage = "resize-numbered-image-to-24"', "The real-host Word regression must enlarge a numbered image formula and verify number synchronization");
expectIncludes(wordAdapter, 'regressionStage = "resize-numbered-image-back-to-10"', "The real-host Word regression must prove that repeated size changes do not accumulate number offsets");
expectIncludes(wordAdapter, 'regressionStage = "direct-resize-numbered-image"', "The real-host Word regression must directly resize a numbered SVG image and refresh its baseline immediately");
expectIncludes(wordAdapter, 'regressionStage = "restore-direct-resized-image"', "The real-host Word regression must restore direct image geometry without losing the SVG baseline ratio");
expectIncludes(wordAdapter, '"image-number direct-resize baseline regression"', "The direct-resize regression must recompute the formula object shift from the shared painted-centre rule");
expectIncludes(wordAdapter, "Abs(numberSizeAt24 - expectedNumberSizeAt24) > 0.1", "The real-host Word regression must keep the visible Equation number at the document-level number size while a formula is resized");
expectIncludes(wordAdapter, 'positionAt24 = expectedPositionAt24', "The real-host Word regression report must reuse the 24 pt vertical centre already verified by the complete layout assertion");
expectIncludes(wordAdapter, 'regressionStage = "reconcile-image-number-after-field-refresh"', "The dedicated image-number regression must identify failures specifically inside image renumbering after field refresh");
expectIncludes(wordAdapter, '"positionBefore=" & CStr(positionBefore)', "The dedicated image-number PASS report must record the pre-refresh picture-position magnitude");
expectIncludes(wordAdapter, '"positionAfterSelection=" & _', "The dedicated image-number PASS report must record the Position after selecting the REF result");
expectIncludes(wordAdapter, '"alignmentFormula=round(CambriaMathGlyphCenter-PngInkCenterFromBaseline)"', "The dedicated image-number PASS report must record the painted-bbox centre formula");
expectIncludes(wordAdapter, '"numberFont=" & VT_WORD_EQUATION_NUMBER_FONT_NAME', "The dedicated image-number PASS report must record Cambria Math numbering");
expectIncludes(wordAdapter, '"initialVisualCenterShift=" & CStr(positionBefore)', "The dedicated image-number regression must report the initial visual-centre object shift");
expectIncludes(wordAdapter, "numberSize = CDbl(numberRange.Font.Size)", "The image-number regression must promote Word font sentinels to Double before validation");
expect(!wordAdapter.includes("positionValue = CDbl(numberRange.Font.Position)"), "The post-refresh image-number regression must not synchronously reread Word Font.Position");
expectIncludes(wordAdapter, "formulaPosition = VTExpectedNumberedImageFormulaPosition", "The image-number regression must use the shared painted-centre position calculation");
expectIncludes(wordAdapter, 'stageName & "/" & assertionStage', "Image-number regression failures must identify the exact overflow-safe assertion stage");
expectIncludes(wordAdapter, "Application.CaptionLabels(wdCaptionEquation)", "Numbered Word formulas must use Word's built-in Equation caption label");
expectIncludes(wordAdapter, "VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX", "Numbered Word formulas must Bookmark the native SEQ result separately");
expectIncludes(wordAdapter, "Type:=wdFieldEmpty", "Numbered Word formulas must create a real localized Equation SEQ field in the formula paragraph");
expectIncludes(wordAdapter, "numberRange.Text <> \"(\" & expectedText & \")\"", "Single-paragraph verification must reject incomplete visible Equation parentheses");
expectIncludes(wordAdapter, ".Font.Hidden = False", "The native Equation SEQ must remain non-hidden so Word for Mac produces a usable field result");
expect(!wordAdapter.includes(".Font.Hidden = True"), "The native Equation numbering path must never hide the SEQ result with Word's Hidden font property");
expectIncludes(wordAdapter, 'numberRange.Text = "(" & expectedNumberText & ")"', "Image single-paragraph numbering must retain complete ordinary parentheses around the complete formatted Equation number");
expectIncludes(wordAdapter, 'rightContent.Text = "()"', "Legacy right-cell formulas must retain ordinary parentheses around their native REF field");
expectIncludes(wordAdapter, "VTRefreshEquationNumberMirror", "Numbering must refresh the native SEQ target and remove legacy mirror artifacts");
expectIncludes(wordAdapter, "VTEquationSequenceFieldHasOrdinal", "Numbering must validate the native SEQ through its retained legal ordinal field code");
expectIncludes(wordAdapter, "expectedText = CStr(sequenceOrdinal)", "The native SEQ result must equal the reconciled expected ordinal, including a chapter/section-local ordinal");
expect(!wordAdapter.includes("insertionRange.InsertAfter expectedText & vbCr"), "Numbering must not create a second plain-text number paragraph");
expect(!wordAdapter.includes("sequenceParagraph.InsertParagraphAfter"), "Numbering must not create a mirror paragraph beside the native SEQ");
expect(!wordAdapter.includes("sequenceField.Result.Text = expectedText"), "Numbering must not overwrite a Word field result Range");
expectIncludes(wordAdapter, "VTParenthesizedEquationReferenceFieldText( _\n            sequenceBookmarkName)", "Legacy visible REF and body references must target the native SEQ result Bookmark");
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
expectIncludes(wordAdapter, "VTReconcileEquationNumbers documentObject", "The shared Equation number update core must rebuild SEQ, visible REF and body REF fields together");
expectIncludes(wordAdapter, "Set sequenceField = VTEnsureNativeEquationSequenceHelper", "New image Equation numbers must retain a live external native SEQ helper field");
expectIncludes(wordAdapter, "Type:=wdFieldRef", "Inserted body cross-references must remain dynamic REF fields");
expectIncludes(wordAdapter, "VT_WORD_NUMBER_BOOKMARK_PREFIX", "Word must retain a formula-specific Bookmark around the complete visible (n) range");
expect(!wordAdapter.includes("VisualTeXEquation"), "New Word numbering must not use the legacy VisualTeX-only sequence name");
expectIncludes(wordAdapter, "Private Sub VTFormatVisibleEquationReference", "Visible Equation number formatting must be centralized and verified");
expectIncludes(wordAdapter, "VTEquationNumberRaisePoints = 0!", "The temporary legacy image scaffold must remain neutral before static-number replacement");
expectIncludes(wordAdapter, "numberPosition = VTCalculateStaticImageEquationNumberPosition", "Renumbering must reapply the SVG mathematical-baseline formula after every field refresh");
expectIncludes(wordAdapter, "VTApplyStaticImageEquationNumberFormatting", "Static image numbering must format the complete ordinary-text parenthesized number as one range");
expectIncludes(wordAdapter, "VTVisibleEquationReferenceFieldText", "Formula-side numbers must use a non-hyperlinked REF field");
expectIncludes(wordAdapter, ".KeepTogether = False", "Numbered formula paragraphs must disable the pagination flag that Word displays as a black square");
expectIncludes(wordAdapter, ".PageBreakBefore = False", "Numbered formula paragraphs must not carry a forced page-break marker");
expectIncludes(wordAdapter, "Private Function VTNativeEquationArrayMarkerRange", "Native display numbering must identify Word's # Equation-array marker");
expectIncludes(wordAdapter, "Private Function VTNativeEquationNumberIsInsideMath", "Native display numbering must verify that the visible VT_N_ REF remains inside OMath");
expectIncludes(wordAdapter, "Private Function VTEnsureNativeEquationSequenceHelper", "Table-free native numbering must create or reuse a dedicated Equation SEQ paragraph after the formula");
expectIncludes(wordAdapter, "Private Function VTNativeEquationArrayReferenceField", "Table-free native numbering must resolve the unique internal REF that mirrors VT_N_");
expectIncludes(wordAdapter, "The native Equation SEQ was absorbed into OMath.", "Native numbering must hard-fail if Word absorbs the true SEQ into formula math");
expectIncludes(wordAdapter, "helperParagraph.OMaths.Count <> 0", "The external Equation SEQ helper must reject any OMath content");
expectIncludes(wordAdapter, "insertionRange.InsertParagraphAfter", "The external Equation SEQ helper must be created only after the completed formula paragraph");
const nativeArrayStart = wordAdapter.indexOf("Private Function VTEnsureNativeEquationArrayNumber");
const nativeArrayEnd = wordAdapter.indexOf("Private Function VTEnsureNativeEquationNumber", nativeArrayStart);
const nativeArraySource = wordAdapter.slice(nativeArrayStart, nativeArrayEnd);
expectIncludes(nativeArraySource, 'Selection.TypeText Text:="#()"', "Native display numbering must create a complete Word #() Equation-array boundary inside OMath");
expectIncludes(nativeArraySource, "Type:=wdFieldRef", "The OMath array number slot must contain a dynamic REF rather than the true Equation SEQ");
expectIncludes(nativeArraySource, "VTEnsureNativeEquationSequenceHelper", "The true Equation SEQ must be isolated before the internal array REF is created");
expect(!nativeArraySource.includes("VTInsertRegisteredEquationCaption( _\n        numberSlotRange"), "The true Equation SEQ must never be inserted into the OMath number slot");
expectIncludes(nativeArraySource, "nativeEquation.BuildUp", "Native numbered OMML must retain built-up professional mathematics");
expectIncludes(nativeArraySource, "nativeEquation.Type = wdOMathDisplay", "Native numbered OMML must remain authentic Word display math");
expectIncludes(nativeArraySource, "VTFinalizeParagraphEquationNumber", "Native Equation arrays must immediately reconcile their SEQ ordinal and Bookmarks");
expectIncludes(wordAdapter, "Private Function VTNativeEquationNumberBookmarkIsCompatible", "Built-up native OMath must accept Word's safe full-equation Bookmark expansion while retaining exact SEQ identity");
expectIncludes(wordAdapter, "VTWordRangeHasMeaningfulText(beforeRange)", "Expanded native Equation Bookmarks must reject meaningful text outside the OMath");
expectIncludes(wordAdapter, "Private Function VTNativeEquationVisibleHorizontalBounds", "Native Equation geometry must measure visible math characters instead of structural # array boundaries");
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
expect(
  !nativeWrapSource.includes("SaveAs2") &&
    !nativeRepairSource.includes("SaveAs2"),
  "Numbered native OMML must not serialize a second staging document",
);
expect(
  !nativeWrapSource.includes("InsertFile") &&
    !nativeRepairSource.includes("InsertFile"),
  "Numbered native OMML must not import a temporary document into the formula cell",
);
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
expectIncludes(wordAdapter, "VTReconcileEquationNumbers documentObject, helperAnchor\n    Set sequenceField = VTResolveEquationSequenceFieldNear", "Word numbering must incrementally reconcile native SEQ fields before creating the new visible right-cell REF");
expectIncludes(wordAdapter, "VTVerifyEquationNumberFieldIntegrity", "Word numbering must verify VT_N_, VT_C_, VT_R_ and the visible REF after every real insertion");
expectIncludes(wordAdapter, "complete VT_N_/VT_C_/VT_R_", "Word numbering must reject a missing native sequence Bookmark instead of leaving empty parentheses");
expectIncludes(wordAdapter, "Public Sub VTCleanupOrphanedNumberedDisplaySelection", "Deleting a numbered formula object must have a targeted orphan-table cleanup path");
expectIncludes(wordAdapter, "layoutTable.Cell(1, 2).Range.InlineShapes.Count <> 0", "Orphan cleanup must leave a numbered table untouched while its image formula still exists");
expectIncludes(wordAdapter, "layoutTable.Cell(1, 2).Range.OMaths.Count <> 0", "Orphan cleanup must leave a numbered table untouched while its native OMath still exists");
expectIncludes(wordAdapter, "VTDeleteWordLatexPayload documentObject, formulaId", "Orphan cleanup must delete the removed formula's document payload state");
expectIncludes(wordAdapter, "VTIsDetachedVisualTeXNativeSequenceHelper", "Deleted native OMML cleanup must recognize only VisualTeX's exact compact helper layout");
expectIncludes(wordAdapter, "Abs(.LeftIndent + 360!) > 0.2", "Detached native helper cleanup must require VisualTeX's unique off-margin paragraph signature");
expectIncludes(wordAdapter, "VTPruneDetachedVisualTeXNativeSequenceHelpers documentObject", "Refresh numbering must remove detached native SEQ helpers before recounting fields");
expectIncludes(wordAdapter, "sequenceAnchors() As Long", "Equation reconciliation must snapshot stable SEQ anchors before updating fields");
expectIncludes(wordAdapter, "If sequenceAnchors(previousIndex) <= anchorValue Then Exit Do", "Mixed image and OMML Equation helpers must be sorted by document anchor without reading index zero");
expectIncludes(wordAdapter, "sequenceAnchors(sequenceCount) = nativeFormulaRange.Start", "VisualTeX numbering must prefer the visible formula position over a potentially displaced helper-field position");
expectIncludes(wordAdapter, "Private Function VTRepairMixedNumberHelperOrder", "Mixed image and OMML numbering must restore each native SEQ helper beside its visible formula before manual refresh");
expectIncludes(wordAdapter, "helperParagraph.Start <> formulaParagraph.End", "A displaced Equation helper must be detected from the visible formula paragraph boundary");
expectIncludes(wordAdapter, "movedHelpers = VTRepairMixedNumberHelperOrder", "Manual Equation refresh must repair mixed-format helper order before assigning ordinals");
expectIncludes(wordAdapter, "sequenceBookmarkNames(previousIndex + 1) =", "Equation anchor sorting must keep each VT_N_ identity paired with its field position");
expectIncludes(wordAdapter, "referenceAnchors() As Long", "Legacy visible REF reconciliation must snapshot stable anchors before rebuilding fields");
expectIncludes(wordAdapter, "VTNormalizePlainWordParagraph paragraphRange", "Orphan cleanup must restore an ordinary body-text paragraph and caret");
expectIncludes(wordAdapter, "Public Sub VisualTeX_WatchOrphanedNumberedDisplay()", "Word must run a lightweight adapter-local orphan watcher without adding another modified source module");
expectIncludes(wordAdapter, "Application.OnTime", "The orphan watcher must defer cleanup until after the user's Delete transaction settles");
expectIncludes(wordAdapter, "VTAnyOpenDocumentHasEquationNumbers()", "The orphan watcher must stop scheduling when no VisualTeX numbered formulas remain");
expectIncludes(wordAdapter, "Private Function VTImageEquationReferenceField", "Image formulas must display a REF to VT_N_ instead of exposing their true SEQ beside the picture");
expectIncludes(wordAdapter, "Private Function VTMigrateLegacyImageEquationSequenceLayouts", "Refreshing must migrate existing same-paragraph image SEQ layouts exactly once");
expectIncludes(wordAdapter, "VTEnsureNativeEquationSequenceHelper", "Image and OMML formulas must isolate Word's native Equation SEQ in a pure-number helper paragraph");
expectIncludes(wordAdapter, "number as (REF VT_N_)", "The image migration must preserve the visible parenthesized number while isolating the native caption source");
expectIncludes(wordAdapter, "Private Function VTInsertEquationNumberReferenceAtRange", "VisualTeX's own parenthesized reference insertion must remain available and unchanged");
expectIncludes(wordAdapter, "Not VTWordInternalMutationActive()", "The orphan watcher must remain disabled during an internal VisualTeX mutation");
expectIncludes(wordAdapter, "VTCleanupOrphanedNumberedDisplaySelection Selection.Range", "The orphan watcher must inspect only the current selection and its adjacent table");
expectIncludes(wordAdapter, "VTCleanupOrphanedNumberedDisplaySelection Selection.Range", "Every new formula command must synchronously clean a nearby orphan table before inserting its placeholder");
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
expectIncludes(wordAdapter, 'regressionStage = "verify-native-sequence-inventory"', "Reference persistence regression must begin with exactly two live VisualTeX numbered formulas and native SEQ fields");
expect(!referenceRegressionSource.includes('regressionStage = "create-existing-word-equation-caption"'), "Reference visibility acceptance must not add an unrelated ordinary Equation caption to the two-formula user workflow");
expectIncludes(wordAdapter, 'regressionStage = "insert-visualtex-native-reference"', "The VisualTeX picker must insert a direct native REF to the live VT_N_ target");
expectIncludes(wordAdapter, 'regressionStage = "insert-second-visualtex-native-reference"', "Reference persistence regression must create two independent direct native VisualTeX references");
expectIncludes(wordAdapter, '"initial VisualTeX Equation reference"', "The first VisualTeX reference must be visible at normal size immediately after insertion");
expectIncludes(wordAdapter, '"second VisualTeX Equation reference"', "The second VisualTeX reference must be visible at normal size immediately after insertion");
expect(referenceRegressionSource.indexOf('"initial VisualTeX Equation reference"') < referenceRegressionSource.indexOf('regressionStage = "insert-second-visualtex-native-reference"'), "The first direct REF must be verified before the second is inserted");
expectIncludes(wordAdapter, 'regressionStage = "delete-preceding-formula-and-renumber"', "Reference persistence regression must preserve both references when the earlier formula is deleted and numbers are refreshed");
expect(!referenceRegressionSource.includes('regressionStage = "convert-referenced-image-to-native"'), "Reference visibility acceptance must not mix image-to-OMML conversion into the two-reference renumbering workflow");
const pickerItemsStart = wordAdapter.indexOf("Private Function VTEquationNumberCrossReferenceItems");
const pickerItemsEnd = wordAdapter.indexOf("Private Function VTInsertEquationNumberReferenceAtRange", pickerItemsStart);
const pickerItemsSource = wordAdapter.slice(pickerItemsStart, pickerItemsEnd);
const insertReferenceStart = pickerItemsEnd;
const insertReferenceEnd = wordAdapter.indexOf("Public Sub VisualTeX_OpenEquationCrossReference", insertReferenceStart);
const insertReferenceSource = wordAdapter.slice(insertReferenceStart, insertReferenceEnd);
expect(!pickerItemsSource.includes("GetCrossReferenceItems"), "The VisualTeX picker inventory must not depend on Word's whole-paragraph caption list");
expect(!insertReferenceSource.includes("InsertCrossReference"), "Single-paragraph VisualTeX references must not ask Word to reference the entire formula line");
expectIncludes(insertReferenceSource, "Type:=wdFieldRef", "VisualTeX references must remain native dynamic Word REF fields");
expectIncludes(insertReferenceSource, "sequenceBookmarkName", "VisualTeX references must target the exact live VT_N_ sequence Bookmark");
expectIncludes(insertReferenceSource, 'insertedRange.Text = "()"', "VisualTeX references must expose only ordinary parentheses around the native REF result");
expectIncludes(wordAdapter, "Private Function VTNativeEquationReferenceItemForFormula", "VisualTeX formulas must map their durable formula id to the corresponding native Equation item without touching Word's private _Ref Bookmark");
const sequenceCodeStart = wordAdapter.indexOf("Private Function VTEquationSequenceFieldCodeForOrdinal");
const sequenceCodeEnd = wordAdapter.indexOf("Private Function VTNormalizeEquationFieldCode", sequenceCodeStart);
const sequenceCodeSource = wordAdapter.slice(sequenceCodeStart, sequenceCodeEnd);
expect(!sequenceCodeSource.includes('" \\r "'), "VisualTeX captions must remain flowing native SEQ fields instead of being restarted during every renumber");
expectIncludes(sequenceCodeSource, '" \\* ARABIC "', "External native SEQ helpers must retain Word's production ARABIC formatting switch");
const registeredCaptionStart = wordAdapter.indexOf("Private Function VTInsertRegisteredEquationCaption");
const registeredCaptionEnd = wordAdapter.indexOf("Private Function VTIsNativeEquationSequenceField", registeredCaptionStart);
const registeredCaptionSource = wordAdapter.slice(registeredCaptionStart, registeredCaptionEnd);
expectIncludes(registeredCaptionSource, '" \\* ARABIC"', "Fresh native Equation captions must create the ARABIC switch only after the SEQ helper is external to OMath");
expectIncludes(wordAdapter, "sequenceRange.End <= VTEquationFieldEnd(candidate)", "A reopened Mac VT_N_ bookmark may end at the SEQ field boundary while remaining an exact schema-7 helper");
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
expectIncludes(convertSource, "If numberLayoutRange.Information(wdWithInTable) Then", "Image-to-OMML conversion must preserve legacy numbered tables without forcing new formulas into them");
expectIncludes(convertSource, "Set finalFormulaRange = VTResolveNativeEquationRange( _\n                targetDocument, nativeBookmarkAnchor, 128)", "A single-paragraph image-to-OMML conversion must re-resolve the final formula from its stable native anchor");
expectIncludes(convertSource, "VTFinalizeParagraphEquationNumber _", "A converted single-paragraph OMML formula must restore its native SEQ and formula Bookmarks before final identity persistence");
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
expectIncludes(reconcileSource, "VTNormalizeBodyEquationReferenceVisibility _\n        documentObject, changedFrom, True", "Native REF reconciliation must update and normalize only body references whose target number changed");
const orphanWatchStart = wordAdapter.indexOf("Public Sub VisualTeX_WatchOrphanedNumberedDisplay()");
const orphanWatchEnd = wordAdapter.indexOf("Public Sub VisualTeX_CreateInline()", orphanWatchStart);
const orphanWatchSource = wordAdapter.slice(orphanWatchStart, orphanWatchEnd);
expectIncludes(orphanWatchSource, "VTNormalizeSelectedEquationReferences", "The orphan watcher must repair manually updated references only near the current selection");
expectIncludes(wordAdapter, "For Each candidateField In scanRange.Fields", "Selection-based reference repair must not enumerate every field in the active document");
expect(!orphanWatchSource.includes("VTNormalizeBodyEquationReferenceVisibility ActiveDocument"), "The one-second orphan watcher must not rescan every body REF while the user types");
expectIncludes(wordAdapter, "For Each candidateBookmark In localRange.Bookmarks", "Equation reconciliation must resolve VT_N_ from the SEQ result or helper paragraph before falling back to a document-wide scan");
expectIncludes(wordAdapter, "For Each candidate In probeRange.Fields", "Equation field re-resolution must inspect only a bounded local Range instead of rescanning every document field for each formula");
expectIncludes(reconcileSource, "Optional ByVal changedFrom As Long = -1", "Equation reconciliation must support incremental updates from the insertion position");
expectIncludes(wordAdapter, "ElseIf documentObject.Bookmarks.Exists(targetBookmarkName) Then", "Incremental numbering must retain Word-native Equation REF targets instead of updating only VisualTeX references");
expectIncludes(wordAdapter, "targetStart >= changedFrom", "Body REF updates must be limited to targets at or after the changed Equation position");
expectIncludes(wordAdapter, "Private Sub VTAssertBodyEquationReferenceVisible", "The real-host reference regression must reject a one-point body REF even when its result text is correct");
expectIncludes(wordAdapter, '"renumbered native Equation reference"', "Deletion acceptance must verify that surviving native references remain visibly formatted after renumbering");
expectIncludes(wordAdapter, '"renumbered native Equation reference"', "Both surviving references must remain visibly formatted after deletion and renumbering");
expect(referenceRegressionSource.indexOf('regressionStage = "delete-preceding-formula-and-renumber"') < referenceRegressionSource.indexOf('"renumbered native Equation reference"'), "Post-refresh visibility must be checked only after deleting the preceding formula and reconciling numbers");
expectIncludes(wordAdapter, '"renumberedVisualTeXReferenceA=1"', "The reference regression PASS report must record the first direct VisualTeX REF after renumbering");
expectIncludes(wordAdapter, '"renumberedVisualTeXReferenceB=1"', "The reference regression PASS report must record the second direct VisualTeX REF after renumbering");
expect(!wordAdapter.includes('"Image-to-OMML conversion lost the surviving formula identity."'), "Reference acceptance must not fail on an unrelated formula-identity assertion that contradicts hands-on editing");
expectIncludes(wordAdapter, "Private Function VTHelperParagraphOwnsNativeEquationSequence", "Orphan cleanup must verify a helper paragraph owns exactly one native Equation SEQ before deleting it");
expectIncludes(wordAdapter, "Private Sub VTPruneUnbookmarkedEmptyNumberTables", "Document-wide cleanup must remove empty VisualTeX number tables even after VT_R_ is lost");
expectIncludes(wordAdapter, "Private Sub VTRepairLiveNumberedTableScaffolds", "Number refresh must repair missing VT_R_/VT_N_/VT_C_ for a still-live formula instead of deleting it");
expectIncludes(wordAdapter, '"(" & Trim$(numberText) & ")  " & previewText', "The Equation picker must show each complete live number together with its formula preview");
expectIncludes(wordAdapter, "VT_WORD_NUMBERING_MODE_CHAPTER", "Word must support document-level chapter Equation numbering");
expectIncludes(wordAdapter, "VT_WORD_NUMBERING_MODE_SECTION", "Word must support document-level section Equation numbering");
expectIncludes(wordAdapter, '"ReadVisualTeXNumberingPreference"', "Word must read the persistent Equation numbering format through its AppleScriptTask bridge instead of direct sandboxed file access");
expectIncludes(wordAdapter, '"WriteVisualTeXNumberingPreference"', "Word must persist the Equation numbering format through its AppleScriptTask bridge instead of direct sandboxed file access");
expect(!wordAdapter.includes("Open preferencePath For Output"), "Word VBA must never directly write the persistent numbering preference path because macOS returns Error 75 inside the Office sandbox");
expectIncludes(protocol, "Public Function VTFileBridgeCall", "The reviewed Word adapter must be able to call fixed-purpose AppleScriptTask bridge handlers");
expectIncludes(wordScript, 'property numberingPreferenceFileName : "VisualTeXNumberingPreference.txt"', "The numbering preference must remain outside the disposable VisualTeXRuntime directory so Word and VisualTeX restarts do not reset it");
expectIncludes(wordScript, "on WriteVisualTeXNumberingPreference(encodedData)", "The Word AppleScriptTask bridge must expose a fixed-purpose numbering-preference writer");
expectIncludes(wordScript, "on ReadVisualTeXNumberingPreference(ignoredValue)", "The Word AppleScriptTask bridge must expose a fixed-purpose numbering-preference reader");
expectIncludes(wordAdapter, "VTTryReadEquationNumberingPreference", "Documents without their own numbering format must inherit the persistent Word-level default");
expectIncludes(wordAdapter, "Private Sub VTMaterializeDocumentEquationNumberingFormat", "An inherited numbering preference must be materialized into document Variables before live SEQ/REF Range mutation begins");
expectIncludes(wordAdapter, "VT_WORD_NUMBERING_PREFERENCE_CACHE_LOADED", "Word must cache the persistent numbering preference after the first AppleScriptTask read in one host session");
expectIncludes(wordAdapter, "Opening the Equation picker is a read operation.", "Opening the Equation picker must not mutate numbering scaffolds just to build its item list");
expectIncludes(wordAdapter, "Reference insertion is intentionally non-repairing.", "Inserting a body Equation REF must not prune or reconcile numbering as a side effect");
expectIncludes(crossReferenceCommandSource, "Set insertionRange = VTResolveEquationReferenceInsertionRange", "The modal Equation picker must normalize and freeze the user's insertion Range before showing its item list");
expectIncludes(crossReferenceCommandSource, "insertionRange, itemIndex, True", "The modal Equation picker must insert into the frozen body-text Range instead of a post-refresh Selection");
expect(!wordAdapter.includes("VisualTeX_DiagnoseNumberingPerformance"), "Temporary numbering performance probes must never ship in the production Word add-in");
expectIncludes(wordAdapter, "VTWriteEquationNumberingPreference numberingMode, separatorText", "Choosing an Equation numbering format must update the persistent default as well as the active document");
expectIncludes(wordAdapter, 'VTEquationSequenceFieldCodeForOrdinal & " \\r 1"', "Chapter and section Equation sequences must use an explicit native SEQ reset at each VisualTeX-resolved scope boundary");
expect(!wordAdapter.includes('" \\s " & CStr(restartLevel)'), "Chapter and section Equation sequences must not delegate scope detection to Word's Heading-style-only SEQ \\s switch");
expectIncludes(wordAdapter, "sequenceLocalOrdinals(itemIndex)", "Chapter and section numbering must feed the local ordinal from the same heading scan that resolves the visible prefix");
expectIncludes(wordAdapter, "VTEquationSequenceResultText(sequenceField) = CStr(sequenceOrdinal)", "Chapter and section SEQ validation must reject a positive but wrong local ordinal");
expectIncludes(wordAdapter, "VTRefreshFormattedSequenceBookmark", "Visible image, OMath and body references must share one complete formatted number Bookmark");
expectIncludes(wordAdapter, "VTComparableEquationNumberText(expectedNumberText)", "Final native OMath acceptance must compare the visible REF with the complete formatted Equation number rather than the local ordinal only");
expectIncludes(wordAdapter, "VTEquationNumberingDisplaySeparator = ChrW(8208)", "Hyphen-style Equation numbers must use U+2010 HYPHEN so OMML does not add binary-operator spacing around the chapter separator");
expect(!wordAdapter.includes("VTEquationNumberingDisplaySeparator = ChrW(8722)"), "Hyphen-style Equation numbers must not use U+2212 MINUS SIGN because Word OMath adds visible operator spacing around it");
expectIncludes(wordAdapter, "VTFindCommittedInlineShapeInDocument", "A newly numbered image must be recoverable from its exact formulaId + metadata identity when a fresh VT_R_ Range is temporarily unstable");
expectIncludes(wordAdapter, 'If numbered And mode = "create" Then', "Any failed numbered create, image or OMML, must remove its formula-specific numbering scaffold before restoring the source");
expectIncludes(wordAdapter, "activeFormulaId = VTDocumentImportRequired", "Document import must remember the in-flight formula id before mutating Word so a mid-finalize failure can be rolled back completely");
expectIncludes(wordAdapter, "VTDeleteEquationNumberScaffold _\n            targetDocument, activeFormulaId, False", "A failed in-flight document-import formula must remove its half-created VT_N_/VT_C_ numbering scaffold");
expectIncludes(wordAdapter, "Private Sub VTApplyEquationNumberingFormatPreset", "Word must route every numbering-format menu choice through one persistent format setter");
for (const callback of [
  "VTWordRibbonNumberingFormatSequence",
  "VTWordRibbonNumberingFormatChapterDot",
  "VTWordRibbonNumberingFormatChapterDash",
  "VTWordRibbonNumberingFormatSectionDot",
  "VTWordRibbonNumberingFormatSectionDash",
]) {
  expectIncludes(wordAdapter, `Public Sub ${callback}`, `The compact numbering-format menu must expose ${callback}`);
  expectIncludes(wordRibbon, `onAction="${callback}"`, `The compact numbering-format menu must wire ${callback}`);
}
expect(!wordAdapter.includes("VisualTeX_ConfigureEquationNumberingFormat"), "Equation numbering format selection must not fall back to the legacy numeric InputBox dialog");
expectIncludes(wordRibbon, '<menu id="VisualTeX.Mac.Word.NumberingFormat"', "The numbering-format control must keep the old compact button footprint while adding a drop-down arrow");
expect(!wordRibbon.includes('<dropDown id="VisualTeX.Mac.Word.NumberingFormat"'), "The numbering-format control must not use the large Office dropDown that stretches the compact Ribbon rows");
expect(!wordRibbon.includes('sizeString="按节编号 (2.3‐1)"'), "The numbering-format menu must not reserve a wide drop-down text box that distorts the Ribbon layout");
expectIncludes(wordEvents, "VTInvalidateWordEquationNumberingFormatControl", "The numbering-format control must remain refreshable when Word opens, creates, or activates a document");
expectIncludes(wordRibbon, '<group id="VisualTeX.Mac.Word.Group" label="VisualTeX">', "The macOS Word Ribbon must use the Windows-parity primary VisualTeX group");
for (const rowId of [
  "VisualTeX.Mac.Word.ConvertRow",
  "VisualTeX.Mac.Word.NumberingRow",
  "VisualTeX.Mac.Word.ReferenceRow",
]) {
  expectIncludes(wordRibbon, `<box id="${rowId}" boxStyle="horizontal">`, `The compact Word Ribbon must define ${rowId}`);
}
expectIncludes(wordRibbon, 'id="VisualTeX.Mac.Word.NumberingFormat"', "Number format must remain in the dedicated compact numbering row");
const compactConvertRow = wordRibbon.slice(
  wordRibbon.indexOf('<box id="VisualTeX.Mac.Word.ConvertRow"'),
  wordRibbon.indexOf('</box>', wordRibbon.indexOf('<box id="VisualTeX.Mac.Word.ConvertRow"')),
);
expect((compactConvertRow.match(/<button\b/g) ?? []).length === 2, "The first compact Ribbon row must contain only the two direct conversion commands");
expectIncludes(wordRibbon, 'id="VisualTeX.Mac.Word.DocumentImport"\n                  label="批量导入"\n                  size="large"', "Batch import must be an independent large Ribbon button with the same visual weight as the primary formula commands");
expect(!compactConvertRow.includes('id="VisualTeX.Mac.Word.DocumentImport"'), "Batch import must not remain compressed inside the small conversion row");
const redesignedWordIcons = [
  ["VisualTeXWordImageInline", "image-inline.svg"],
  ["VisualTeXWordImageDisplay", "image-display.svg"],
  ["VisualTeXWordOmmlInline", "omml-inline.svg"],
  ["VisualTeXWordOmmlDisplay", "omml-display.svg"],
  ["VisualTeXWordEditSelected", "edit-selected.svg"],
  ["VisualTeXWordConvertOmml", "convert-omml.svg"],
  ["VisualTeXWordConvertImage", "convert-image.svg"],
  ["VisualTeXWordDocumentImport", "document-import.svg"],
  ["VisualTeXWordUpdateNumbers", "update-numbers.svg"],
  ["VisualTeXWordNumberingFormat", "numbering-format.svg"],
  ["VisualTeXWordCrossReference", "cross-reference.svg"],
  ["VisualTeXWordOpenApp", "open-visualtex.svg"],
  ["VisualTeXWordRedrawSelection", "redraw-selection.svg"],
  ["VisualTeXWordRedrawDocument", "redraw-document.svg"],
];
for (const [imageId, svgName] of redesignedWordIcons) {
  expectIncludes(wordRibbon, `image="${imageId}"`, `The redesigned Word Ribbon must reference ${imageId}`);
  expectIncludes(packager, `${imageId}: { svg: "${svgName}" }`, `The macOS packager must render the original VisualTeX SVG ${svgName}`);
  const iconSource = read(`office/macos-offline/word/ribbon-icons/${svgName}`);
  expectIncludes(iconSource, "#1F638E", `${svgName} must use the VisualTeX deep-blue brand color`);
  expect(!iconSource.includes("#41474D") && !iconSource.includes("#7A8793"), `${svgName} must not retain the previous gray palette`);
}
expect(!/image="VisualTeXIcon(?:02|04|06|07|08|09)"/.test(wordRibbon), "The redesigned Word Ribbon must not reuse the previous macOS formula icon set");
expect(!wordRibbon.includes("imageMso="), "Every visible VisualTeX Word command must use its own designed icon rather than a built-in Office icon");
expectIncludes(packager, 'const wordRibbonIconRoot = join(offlineRoot, "word", "ribbon-icons")', "The Word packager must read the reviewed original SVG icon set");
expectIncludes(packager, '"/usr/bin/sips"', "The Word packager must rasterize the original SVG artwork into embedded PNG Ribbon resources");
expect(!packager.includes("windowsRibbonIconBytes"), "The Word packager must not reuse the Windows Ribbon icon data");
expect(!packager.includes("RibbonIconData.cs"), "The Word packager must be independent from the old Windows icon source file");
expectIncludes(wordRibbon, '<group id="VisualTeX.Mac.Word.RedrawGroup" label="LaTeX 重绘">', "LaTeX redraw must remain an independent Ribbon group after the primary tools");
for (const label of [
  "将 OMML 公式转为 LaTeX",
  "将图片公式转为 LaTeX",
  "将全文 OMML 公式转为 LaTeX",
  "将全文图片公式转为 LaTeX",
]) {
  expectIncludes(wordRibbon, `label="${label}"`, `Word redraw menus must expose ${label}`);
}
expectIncludes(wordAdapter, "payload = sourceRange.WordOpenXML", "Native Word formulas must be restored from their current WordOpenXML instead of VisualTeX metadata");
expectIncludes(wordAdapter, 'VTStartWordFormulaRestore "selection", "omml", "image"', "Native OMML without a VisualTeX Bookmark must use the hidden direct image conversion path");
expectIncludes(wordAdapter, "VTTryResolveVisualTeXInlineShapeReference", "Image-to-LaTeX must accept only recoverable VisualTeX image formulas");
expectIncludes(rustRuntime, '"/Applications/Microsoft Word.app/Contents/Resources/omml2mathml.xsl"', "Native OMML restoration must use Microsoft Word's own OMML-to-MathML stylesheet");
expectIncludes(rustRuntime, 'Command::new("/usr/bin/xsltproc")', "Native OMML restoration must execute the reviewed local XSL transform without network access");
expectIncludes(rustRuntime, '"formulaRestore"', "The native Office runtime must support the hidden formula restore operation");
expectIncludes(wordLatexRedrawApp, 'request.operation === "formulaRestore"', "The hidden Word renderer must automatically process formula restore sessions");
expectIncludes(wordLatexRedrawApp, 'displayMode === "block" ? `$$${latex}$$` : `$${latex}$`', "Restored LaTeX must use dollar delimiters for display and inline formulas");
expectIncludes(mathMlToLatex, "export function mathMlToLatex", "The macOS frontend must contain a dedicated MathML-to-LaTeX converter");
expectIncludes(mathMlToLatex, 'case "mfrac"', "The MathML converter must preserve Word fractions");
expectIncludes(mathMlToLatex, 'case "msubsup"', "The MathML converter must preserve combined scripts");
expectIncludes(styles, ".doc-import-primary:disabled", "The batch-import primary action must define a readable disabled state");
expectIncludes(styles, "opacity: 1", "Batch-import disabled actions must not disappear through low global opacity");
expectIncludes(styles, "border-top: 1px solid color-mix", "The batch-import footer must have a distinct high-contrast boundary");
expectIncludes(wordRibbon, '<group id="VisualTeX.Mac.Word.FontSizeGroup" label="公式字号">', "Image formula point sizes must live in an independent Windows-parity Ribbon group");
const numberingFormatStart = wordAdapter.indexOf(
  "Public Sub VTWordRibbonGetNumberingFormatItemCount(",
);
const numberingFormatEnd = wordAdapter.indexOf(
  "Public Sub VTWordRibbonOpen(",
  numberingFormatStart,
);
const numberingFormatSource = wordAdapter.slice(
  numberingFormatStart,
  numberingFormatEnd,
);
expect(
  numberingFormatStart >= 0 && numberingFormatEnd > numberingFormatStart,
  "The Equation numbering-format callbacks must remain discoverable for source-encoding regression checks",
);
expect(
  !/[\u3400-\u9fff]/u.test(numberingFormatSource),
  "The macOS VBA numbering-format compatibility callbacks must generate Chinese labels through VTUnicodeText instead of source-code-page text",
);
expectIncludes(numberingFormatSource, "VTUnicodeText(", "The numbering-format compatibility callbacks must construct their Chinese labels at runtime as Unicode");
expect(!numberingFormatSource.includes("InputBox("), "The numbering-format drop-down must never ask the user to type a numeric option");
expectIncludes(wordAdapter, "numberFontSizePt = VTVisibleEquationNumberFontSize", "Image Equation numbers must use the same document-level number size as native OMath numbers");
expectIncludes(wordAdapter, "sourceHeightPoints = target.Height", "Image-to-OMML conversion must preserve the source formula height for number alignment");
expectIncludes(wordAdapter, "VTEnsureNativeEquationNumber", "Image-to-OMML conversion must rebuild the shared numbered table around the native formula");
expectIncludes(wordAdapter, "target.Delete", "Word replacement must delete the old object only after candidate setup");
expectIncludes(wordAdapter, "Public Sub VisualTeX_ConvertSelectedToNativeEquation()", "Word must expose a selected-formula native equation conversion command");
expectIncludes(wordAdapter, "Set targetDocument = target.Range.Document", "Word image-to-native conversion must retain the image's owning document across hidden DOCX staging");
expectIncludes(wordAdapter, "VTSetWordMetadataPayload targetDocument, formulaId, encodedMetadata", "Word image-to-native conversion must store metadata in the owning document rather than a transient ActiveDocument");
expectIncludes(wordAdapter, "insertionAnchor.Collapse wdCollapseEnd", "Word image-to-native conversion must insert after the source picture so deleting the picture shifts the OMath into place");
expectIncludes(wordAdapter, "Set sourceImage = VTFindUniqueInlineShape(encodedMetadata)", "Word image-to-native conversion must resolve a fresh picture object after hidden DOCX staging");
expectIncludes(wordAdapter, "VTDeleteVisualTeXImageContainer sourceImage", "Word image-to-native conversion must transactionally remove either a current plain image or a legacy field container before resolving the final OMath Range");
expectIncludes(wordAdapter, "sourceBackupRange.FormattedText", "Word image-to-native conversion must retain an exact source-image rollback copy");
expectIncludes(wordAdapter, "replaceSourceContainer = (sourceContainerRange.OMaths.Count = 1)", "An image retained inside Word's empty OMath compatibility container must use an in-place replacement transaction");
expectIncludes(wordAdapter, "If replaceSourceContainer Then", "Image-to-OMML conversion must branch between compatibility-container replacement and standalone-image insertion");
expectIncludes(wordAdapter, "sourceContainerRange, _", "The compatibility-container path must replace the original formula Range rather than nesting a second OMath beside it");
expectIncludes(wordAdapter, "Not VTIsVisualTeXInlineShape", "The image backup must validate the copied VisualTeX image identity without requiring a legacy Word Field wrapper");
expect(!wordAdapter.includes("sourceBackupDocument.OMaths.Count <> 0"), "The image backup must accept a VisualTeX picture that Word keeps inside an empty OMath compatibility container");
expect(!wordAdapter.includes("sourceBackupDocument.Fields.Count <> 1"), "Current plain VisualTeX images must not fail conversion merely because they have no legacy Field wrapper");
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
expectIncludes(wordAdapter, "Public Sub VisualTeX_RunWordMiddleNativeInsertionRegression()", "The packaged Word add-in must expose the 20-plus-formula middle native insertion regression");
expectIncludes(wordAdapter, 'Const insertBeforeOrdinal As Long = 21', "Middle native insertion regression must reproduce insertion at the user's 21st Equation position");
expectIncludes(wordAdapter, '"crossReference=25"', "Middle native insertion regression must prove downstream cross-references flow after renumbering");
expectIncludes(wordAdapter, "Private Function VTCanUseEquationTailFastPath", "Native numbering fast paths must use one strict tail gate");
expectIncludes(wordAdapter, "VTHasManagedEquationNumberAfter", "The strict tail gate must reject a later managed VisualTeX number even when Word's live Field view is stale");
expectIncludes(wordAdapter, "VTReconcileEquationNumbers documentObject", "Middle insertion must be able to replay the full flowing Equation sequence on Word for Mac");
expectIncludes(wordAdapter, "sequenceOrdinal, sequenceHeadingPrefixes(itemIndex), True, _\n                        forceFlowingSequenceReplay", "Chapter/section numbering must honor the same forced SEQ replay used by middle/prepend insertion");
expectIncludes(wordAdapter, "VTReconcileEquationNumbers documentObject, -1, True", "Explicit Update Numbers and sequence-order changes must force a full managed SEQ replay");
expectIncludes(wordAdapter, "Public Sub VisualTeX_RunWordChapterPrependNumberingRegression()", "The packaged Word add-in must expose the chapter-number prepend/update repair regression");
expectIncludes(wordAdapter, "Public Sub VisualTeX_RunWordChapterZeroScopeNumberingRegression()", "The packaged Word add-in must expose the chapter-0 regression before the first Heading 1");
expectIncludes(wordAdapter, '"initial=0-1,0-2,1-1"', "Chapter numbering regression must preserve chapter 0 until the first Heading 1");
expectIncludes(wordAdapter, "Public Sub VisualTeX_RunWordSectionImplicitNumberingRegression()", "The packaged Word add-in must expose the section-0 regression before the first Heading 2");
expectIncludes(wordAdapter, '"initial=0.0.1,1.0.1,1.0.2,2.0.1,2.0.2,2.1.1,2.1.2"', "Section numbering regression must distinguish section 0 from the first real section");
expectIncludes(wordAdapter, '"explicitUpdateRepair=PASS"', "The chapter-number regression must prove explicit Update Numbers repairs a stale cached SEQ result");
expectIncludes(wordAdapter, "Public Sub VisualTeX_RunWordSafeNativeInsertionRegression()", "The packaged Word add-in must expose the three-position safe native insertion regression");
expectIncludes(wordAdapter, 'regressionStage = "body-paragraph-end"', "Safe insertion must cover a normal body paragraph end");
expectIncludes(wordAdapter, 'regressionStage = "empty-line-between-body-and-formula"', "Safe insertion must cover an empty line between body text and an existing formula");
expectIncludes(wordAdapter, 'regressionStage = "empty-line-between-two-formulas"', "Safe insertion must cover an empty line between two existing formulas");
expectIncludes(wordAdapter, "Private Function VTRegressionInsertBlankParagraphBeforeFormula", "Safe insertion regression must create real blank Word paragraphs rather than writing a paragraph mark at an OMath boundary");
expectIncludes(wordAdapter, "formulaParagraph.InsertParagraphBefore", "Safe insertion regression must create the middle blank line through the complete neighboring formula paragraph");
expectIncludes(wordAdapter, "nativeBookmarkName = VTNativeFormulaBookmarkName(formulaId)", "Numbered native formula resolution must recover from a Word-expanded visible-number Bookmark through the same formula's native identity");
expectIncludes(wordAdapter, "Set nativeMath = VTNativeMathForBookmark", "Native identity recovery must resolve the exact OMath from VT_O_ rather than selecting a nearby formula");
expectIncludes(wordAdapter, "documentObject, formulaId, repairedNumberRange", "Native identity recovery must tighten VT_R_ back to the recovered formula's array number range");
expectIncludes(wordAdapter, "VTSetNativeFormulaBookmark documentObject, nativeRange, formulaId", "Native identity recovery must retighten VT_O_ after a blank paragraph shifts the formula");
expectIncludes(wordAdapter, 'ReferenceKind:=wdEntireCaption', "Safe insertion regression must verify Word's entire-caption cross-reference returns only the external SEQ number");
expectIncludes(wordAdapter, '"externalSeq=PASS"', "Safe insertion PASS output must record that every true SEQ remained outside OMath");
expectIncludes(wordAdapter, '"internalRef=PASS"', "Safe insertion PASS output must record that each visible number remained an internal REF");
expectIncludes(wordAdapter, "Public Sub VisualTeX_RunWordSingleParagraphNumberRegression()", "The packaged Word add-in must expose a real-host blank-first-line single-paragraph numbering regression");
expectIncludes(wordAdapter, "Public Sub VisualTeX_RunWordNumberingPerformanceRegression()", "The packaged Word add-in must retain the 30-plus Equation numbering performance regression");
expectIncludes(wordAdapter, 'regressionStage = "blank-first-line-create"', "The real-host regression must create a numbered formula on the first line of a blank document");
expectIncludes(wordAdapter, "testDocument.Tables.Count <> 0", "The blank-first-line regression must reject any new numbered table");
expectIncludes(wordAdapter, "paragraphRange.Paragraphs.Count <> 1 Or tabCount <> 2", "The blank-first-line regression must require one paragraph with exactly two tabs");
expectIncludes(wordAdapter, 'regressionStage = "blank-first-line-cross-reference"', "The blank-first-line regression must verify a live native Equation reference (1)");
expectIncludes(wordAdapter, 'regressionStage = "blank-first-line-convert-to-omml"', "The blank-first-line regression must convert the numbered image through the real image-to-OMML production path");
expectIncludes(wordAdapter, 'regressionStage = "converted-omml-layout"', "The converted OMML must be rechecked for one paragraph, zero tables, two tabs and visual-center alignment");
expectIncludes(wordAdapter, 'regressionStage = "converted-omml-refresh-and-reference"', "The converted OMML must retain visual-center alignment after native SEQ refresh and keep its dynamic reference");
expectIncludes(wordAdapter, '"imageTabs=2"', "The real-host PASS report must record the two-tab image layout");
expectIncludes(wordAdapter, '"convertedTabs=0"', "The real-host PASS report must record the tab-free native Equation-array layout");
expectIncludes(wordAdapter, '"convertedDisplay=PASS"', "The real-host PASS report must record authentic post-conversion display math");
expectIncludes(wordAdapter, '"convertedVisualCenter=PASS"', "The real-host PASS report must record successful post-conversion visual-center verification");
expectIncludes(wordAdapter, '"convertedReference=(1)"', "The real-host PASS report must record the surviving post-conversion dynamic reference");
expectIncludes(wordAdapter, "Public Sub VisualTeX_RunWordNativeRegression()", "The packaged Word add-in must expose a real-host native equation regression entry point");
expectIncludes(wordAdapter, "VTEquationNumberCrossReferenceItems", "The Word cross-reference picker must expose parenthesized items from Word's native Equation list");
expectIncludes(wordAdapter, "Private Function VTEquationNumberCrossReferenceItems", "The Word Ribbon cross-reference picker must enumerate VisualTeX's live formula identities directly");
expectIncludes(wordAdapter, "Private Sub VTAssertNumberedEquationLayout", "The real-host regression must measure numbered formula centering and right alignment");
expectIncludes(wordAdapter, "wdHorizontalPositionRelativeToTextBoundary", "The Word numbering regression must measure actual formula and number positions inside the text column");
expectIncludes(wordAdapter, 'regressionStage = "native-numbered-edit"', "The real-host regression must verify numbered OMML editing preserves the layout and number");
expectIncludes(wordAdapter, 'regressionStage = "image-to-native-number-preservation"', "The real-host regression must verify image-to-OMML conversion preserves number geometry");
expectIncludes(wordAdapter, 'VTEquationNumberTextForFormula( _\n                            documentObject, formulaId)', "Image-to-OMML preflight must compare a visible REF with the complete VT_N_ number, including chapter or section prefixes");
expectIncludes(wordAdapter, 'VT_WORD_NUMBERING_MODE_CHAPTER\n    VTSetDocumentVariable _\n        testDocument, VT_WORD_NUMBERING_SEPARATOR_VARIABLE, "-"', "The real-host image-to-OMML regression must cover a fresh chapter-mode 0-1 formula whose SEQ result is only 1");
expectIncludes(wordAdapter, 'assertionName & ": Equation number parentheses are incomplete."', "The Word numbering regression must reject missing number parentheses");
expectIncludes(wordAdapter, "Not VTWordRangeHasMeaningfulText(beforeRange)", "Inline OMML paragraph alignment must change only when there is no surrounding meaningful text");
expectIncludes(wordAdapter, "Set targetDocument = ActiveDocument", "Word commits must capture the owning document before opening hidden staging DOCX files");
expectIncludes(wordAdapter, "targetDocument.Activate", "Word commits must reactivate the owning document after hidden DOCX staging changes ActiveDocument");
expectIncludes(wordAdapter, "VTSetWordLatexPayload targetDocument, formulaId, latexBase64", "Word commits must persist a formula-id keyed LaTeX edit payload in the owning document");
expectIncludes(wordAdapter, "VTSetWordOmmlPayload targetDocument, formulaId, ommlBase64", "Word commits must persist a formula-id keyed structural OMML payload in the owning document");
expectIncludes(wordAdapter, "Set sequenceField = VTNativeEquationSequenceHelperField", "Existing static image-number formulas must reuse and refresh their external native SEQ helper field");
expectIncludes(wordAdapter, "VTEnsureEquationNumberFields layoutTable, formulaId", "Existing legacy numbered tables must remain repairable without changing their document structure");
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
expectIncludes(wordAdapter, "originalReplacementMath", "Word native replacement must retain the exact replaced OMath identity while resolving the transferred object");
expectIncludes(wordAdapter, "VTOMathHasMeaningfulContent", "Word native transfer must distinguish the real equation from empty dotted OMath shells");
expectIncludes(wordAdapter, "one clean OMath object", "Word native transfer must enforce one clean equation object after replacement");
expectIncludes(wordAdapter, "If Not replaceTarget Then insertionRange.Collapse wdCollapseStart", "Word native edits must replace the exact original OMath Range instead of inserting beside it");
expect(!wordAdapter.includes("originalNativeMath.Range.Delete"), "Word replacement must never delete a stale original OMath COM Range after insertion");
expect(!wordAdapter.includes("candidate.Range.End + originalNativeLength"), "Word replacement must not reconstruct a deletion Range from stale OMath length arithmetic");
expectIncludes(wordAdapter, "VTSetNativeFormulaBookmark", "Word native formulas must retain a persistent VisualTeX identity bookmark");
expectIncludes(wordAdapter, "VTSetWordMetadataPayload", "Word native formulas must retain their complete VisualTeX edit metadata");
expectIncludes(wordAdapter, "Private Function VTWordConvertNativeBookmarkToImageFast", "The formula-scoped cached conversion helper may remain available for future validated reuse");
expectIncludes(wordAdapter, "If numbered Then Exit Function", "The retained cached helper must never handle numbered OMML even if it is re-enabled later");
expectIncludes(wordAdapter, "VTWordEditNativeBookmark nativeBookmark", "Word native VisualTeX formulas must pass the already-resolved Bookmark directly into the edit path");
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
expectIncludes(powerpointAdapter, '"powerpoint-office-performance-20260801-r4"', "PowerPoint health must identify the optimized native Office build");
expectIncludes(powerpointAdapter, "Public Sub VTPowerPointRibbonOnLoad", "PowerPoint Ribbon load must retain its IRibbonUI handle and initialize events");
expectIncludes(powerpointAdapter, "VisualTeX_DoubleClickEditSelected", "PowerPoint must expose a non-modal native double-click macro entry point");
expectIncludes(powerpointAdapter, "VTInitializePowerPointEvents", "PowerPoint Auto_Open must initialize its persistent application event sink");
expectIncludes(powerpointEvents, "App_WindowBeforeDoubleClick", "PowerPoint must use its native application event for double-click editing");
expectIncludes(powerpointEvents, "App_WindowSelectionChange", "PowerPoint selection changes must synchronize SVG formula point-size state");
expectIncludes(powerpointEvents, "VisualTeX_SynchronizeSelectedFormulaSize Sel", "PowerPoint selection changes must refresh the selected SVG formula size");
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
expectIncludes(powerpointAdapter, "VT_POWERPOINT_REFERENCE_FONT_SIZE_PT As Double = 14#", "PowerPoint SVG formulas must use a stable 14 pt reference size");
expectIncludes(powerpointAdapter, "VTPreferredPowerPointFormulaFontSize", "New PowerPoint formulas must inherit a selected text or formula point size");
expectIncludes(powerpointAdapter, "target.Width = CSng(targetWidth)", "PowerPoint SVG point sizes must map to proportional width");
expectIncludes(powerpointAdapter, "target.Height = CSng(targetHeight)", "PowerPoint SVG point sizes must map to proportional height");
expectIncludes(powerpointAdapter, "centerX - targetWidth / 2#", "PowerPoint SVG resizing must preserve the formula center");
expectIncludes(powerpointAdapter, "VisualTeXFontSizePt", "PowerPoint shapes must persist their formula point size as a durable tag");
expectIncludes(powerpointAdapter, "VisualTeXReferenceWidthPt", "PowerPoint shapes must persist the 14 pt SVG reference width");
expectIncludes(powerpointAdapter, "VisualTeXReferenceHeightPt", "PowerPoint shapes must persist the 14 pt SVG reference height");
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
expectIncludes(protocol, "Private Function VTTryReadTextDirect", "VBA runtime reads must use a direct local-file fast path before crossing AppleScriptTask");
expectIncludes(protocol, 'VTFileBridgeCall("ReadVisualTeXFile"', "VBA runtime reads must retain the fixed AppleScriptTask fallback for sandbox compatibility");
expectIncludes(protocol, "For attempt = 1 To 3", "The Office file bridge must retry transient empty AppleScriptTask responses");
expectIncludes(protocol, "If Len(response) > 0 Then Exit For", "The Office file bridge retry must stop immediately after a valid response");
expectIncludes(protocol, "VisualTeX file-existence self-test failed", "The host self-test must exercise the file-existence bridge handler");
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
expectIncludes(wordScript, "my launchVisualTeXURL(visualTeXURL)", "Word AppleScriptTask must route the validated Session URL through the prewarmed executable");
expectIncludes(powerpointScript, "my launchVisualTeXURL(visualTeXURL)", "PowerPoint AppleScriptTask must route the validated Session URL through the prewarmed executable");
expectIncludes(wordScript, 'do shell script "/usr/bin/open -b " & quoted form of "com.visualtex.studio"', "Word must open only the fixed VisualTeX bundle identifier");
expectIncludes(powerpointScript, 'do shell script "/usr/bin/open -b " & quoted form of "com.visualtex.studio"', "PowerPoint must open only the fixed VisualTeX bundle identifier");
const wordPrewarmStart = wordScript.indexOf("on PrewarmVisualTeXApplication(hostName)");
const wordPrewarmEnd = wordScript.indexOf("end PrewarmVisualTeXApplication", wordPrewarmStart);
const wordPrewarmSource = wordScript.slice(wordPrewarmStart, wordPrewarmEnd);
expectIncludes(wordPrewarmSource, "runningVisualTeXExecutable", "Word startup prewarming must resolve one fully launched VisualTeX resident before formula clicks");
expect(!wordPrewarmSource.includes('/usr/bin/open -gj -b '), "Word startup prewarming must delegate cold-launch policy to the shared ready-resident resolver");
expect(!wordScript.includes("System Events"), "Word AppleScriptTask must not use UI automation");
expect(!powerpointScript.includes("System Events"), "PowerPoint AppleScriptTask must not use UI automation");
for (const [host, script] of [["Word", wordScript], ["PowerPoint", powerpointScript]]) {
  expectIncludes(script, "validateRelativePath", `${host} file bridge must validate every runtime-relative path`);
  expectIncludes(script, "absoluteRuntimePath", `${host} file bridge must join paths only beneath its fixed runtime root`);
  expectIncludes(script, "stringByDeletingLastPathComponent", `${host} file bridge must resolve the validated destination parent without spawning dirname`);
  expectIncludes(script, "initWithBase64EncodedString:normalizedData options:0", `${host} file bridge must decode the validated request in-process`);
  expectIncludes(script, "writeToFile:targetPath atomically:true", `${host} file bridge must atomically replace the validated destination through Foundation`);
  expectIncludes(script, 'do shell script "/bin/chmod 600 " & quoted form of targetPath', `${host} file bridge must restrict the final request file permissions`);
  expect(!script.includes("/usr/bin/mktemp"), `${host} request writes must not spawn mktemp on the latency-critical editor-open path`);
  expect(!script.includes("set temporaryPath to do shell script"), `${host} request writes must not materialize a shell-managed temporary file on the latency-critical editor-open path`);
  expect(!script.includes('" && /bin/mv -f " & quoted form of temporaryPath'), `${host} request writes must not spawn a shell move on the latency-critical editor-open path`);
  expectIncludes(script, 'set encodedData to do shell script "/usr/bin/base64 < " & quoted form of targetPath', `${host} file bridge must quote the validated read path`);
  expectIncludes(script, 'do shell script "/bin/mkdir -p " & quoted form of targetPath & " && /bin/chmod 700 " & quoted form of targetPath', `${host} runtime directory creation must quote its validated path`);
  expectIncludes(script, "on launchVisualTeXURL(visualTeXURL)", `${host} Session launch must use the fixed argv-forwarding helper`);
  expectIncludes(script, 'if candidatePath ends with executableSuffix then', `${host} cold resident discovery must reject unrelated executables`);
  const launchStart = script.indexOf("on launchVisualTeXURL(visualTeXURL)");
  const launchEnd = script.indexOf("end launchVisualTeXURL", launchStart);
  const launchSource = script.slice(launchStart, launchEnd);
  expectIncludes(launchSource, "/usr/bin/nohup", `${host} Session launch must preserve the detached forwarding helper that does not steal Office focus`);
  expectIncludes(launchSource, "quoted form of executablePath", `${host} Session launch must execute only the resolved VisualTeX binary`);
  expectIncludes(launchSource, "quoted form of safeURL", `${host} Session launch must put only the validated URL in argv for Tauri single-instance IPC`);
  const resolverStart = script.indexOf("on runningVisualTeXExecutable()");
  const resolverEnd = script.indexOf("end runningVisualTeXExecutable", resolverStart);
  const resolverSource = script.slice(resolverStart, resolverEnd);
  expectIncludes(resolverSource, 'if cachedVisualTeXExecutable is not "" then return cachedVisualTeXExecutable', `${host} hot formula launch must return the prewarmed executable without pgrep/ps validation`);
  expect(
    resolverSource.indexOf('if cachedVisualTeXExecutable is not "" then return cachedVisualTeXExecutable') < resolverSource.indexOf("firstRunningVisualTeXExecutable"),
    `${host} hot formula launch must hit the cached resident before any process scan`,
  );
  expectIncludes(resolverSource, '/usr/bin/open -gj -b ', `${host} cold formula launch must prewarm VisualTeX without activating it`);
  expectIncludes(resolverSource, 'delay 0.5', `${host} cold formula launch must preserve the validated single-instance settling wait`);
  expectIncludes(resolverSource, 'repeat with attemptIndex from 1 to 80', `${host} cold formula launch must bound resident startup waiting`);
  expect(!resolverSource.includes('set standardExecutable to "/Applications/VisualTeX.app/Contents/MacOS/visualtex"'), `${host} cold formula launch must not mistake an installed binary for a ready resident process`);
  expectIncludes(script, "cachedVisualTeXExecutable", `${host} Session launch must cache the executable resolved during prewarm`);
  expect(!script.includes('do shell script "/usr/bin/open " & quoted form of visualTeXURL'), `${host} Session launch must not use a LaunchServices URL AppleEvent that the macOS second instance can lose`);
  expect(!script.includes("openURL:targetURL"), `${host} Session launch must not pass NSURL objects through AppleScriptTask`);
  expectIncludes(script, 'use framework "Foundation"', `${host} AppleScriptTask must use NSProcessInfo only for monotonic timing`);
  expectIncludes(script, "systemUptime", `${host} AppleScriptTask must report monotonic write-and-launch timing without wall-clock ambiguity`);
  expectIncludes(script, 'candidate contains ".."', `${host} file bridge must reject traversal components`);
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
expectIncludes(rustRuntime, "if request.native_equation", "Tauri must materialize only the Word representation required by the current commit mode");
expectIncludes(rustRuntime, "atomic_write(&path, &omml_docx, 0o600)?", "Native Word commits must durably materialize their formula-scoped staging DOCX before dispatch");
expectIncludes(rustRuntime, 'const RESULT_SVG_FILE: &str = "formula.svg"', "Native Office formulas must be materialized as SVG files");
expectIncludes(rustRuntime, 'const RESULT_WORD_SVG_DOCX_FILE: &str = "formula-svg.docx"', "Word must receive SVG through a generated OOXML staging document");
expectIncludes(rustRuntime, "build_word_svg_docx", "Tauri must package the SVG and PNG preview into a minimal Word document");
expectIncludes(rustRuntime, 'xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"', "The Word staging DOCX must use Office SVG blip markup");
expectIncludes(rustRuntime, '"vectorDocumentPath"', "Word dispatches must carry the generated SVG staging DOCX path");
expectIncludes(rustRuntime, '"fallbackImagePath"', "Word and PowerPoint dispatches must retain PNG only as a compatibility fallback");
expectIncludes(rustRuntime, "materialize_powerpoint_svg(session)?", "PowerPoint commits must insert the vector SVG export");
expectIncludes(powerpointAdapter, "Set targetPresentation = VTFindPowerPointPresentation(expectedPresentation)", "PowerPoint Apply must resolve the Session's target presentation without depending on the currently active document");
expectIncludes(powerpointAdapter, "Set currentSlide = targetPresentation.Slides(slideIndex)", "PowerPoint Apply must mutate the slide belonging to the resolved Session target");
expectIncludes(powerpointAdapter, "VTPresentationIdentityFor(candidatePresentation)", "PowerPoint target matching must use the same full-path-or-name identity rule as Session creation");
expect(!powerpointAdapter.includes("VTActivatePowerPointPresentation"), "PowerPoint Apply must not switch document windows as part of Session commit");
expectIncludes(powerpointAdapter, "For Each currentPresentation In Presentations", "PowerPoint create cancellation must locate its pending placeholder across all open presentations");
expect(!powerpointAdapter.includes("For Each currentSlide In ActivePresentation.Slides"), "PowerPoint create cancellation must not depend on the currently active presentation");
expectIncludes(rustRuntime, "decode_svg", "PowerPoint SVG exports must be validated before Office receives them");
expectIncludes(rustRuntime, "POWERPOINT_REFERENCE_FONT_SIZE_PT", "The native runtime must scale PowerPoint SVG geometry from a fixed point-size reference");
expectIncludes(rustRuntime, "previous_reference_height", "PowerPoint edits must infer the current point size from existing SVG geometry");
expectIncludes(rustRuntime, "let committed_font_size = session.font_size_pt", "PowerPoint commits must prefer the font size selected in the editor Session");
expectIncludes(rustRuntime, "let font_size_pt = session", "Word commits must prefer the font size selected in the editor Session");
expect(
  (dialogApp.match(/fontSizePt: officeFontSizePt/g) ?? []).length >= 3,
  "Office explicit apply, autosave, and close-commit drafts must all persist the selected point size",
);
expectIncludes(
  formulaRenderArtifacts,
  "export function tryRenderOfficeFormulaDraftArtifacts",
  "Office draft rendering must provide a non-throwing path for temporary MathLive source",
);
expectIncludes(
  dialogApp,
  "const generateDraftExportResult = useCallback",
  "The Office editor must isolate best-effort draft rendering from strict explicit apply",
);
expectIncludes(
  dialogApp,
  "const exportResult = generateDraftExportResult();",
  "Office autosave must tolerate placeholders and other incomplete LaTeX without showing an export error",
);
expectIncludes(
  dialogApp,
  'status === "editing"\n            ? generateDraftExportResult()\n            : generateSvgExportResult()',
  "Close-time draft persistence must remain tolerant while a committing write remains strict",
);
expectIncludes(
  dialogApp,
  'const update = await buildCurrentSessionUpdate("committing");',
  "Explicit Office apply must continue through the strict artifact generation path",
);
expectIncludes(rustRuntime, '("fontSizePt", format!', "PowerPoint dispatches must carry the resolved point size back to VBA");
expectIncludes(rustRuntime, "metadata.font_size_pt = Some(geometry.font_size_pt)", "PowerPoint metadata must retain the resolved SVG point size");
expectIncludes(wordAdapter, 'VTApplicationSupportRoot() & "/NativeDocuments/" & formulaId & ".docx"', "Word image-to-OMML conversion must resolve the same durable formula-scoped staging path");
const handleOpenUrlStart = rustRuntime.indexOf("pub(crate) fn handle_open_url");
const handleOpenUrlEnd = rustRuntime.indexOf("fn decode_png", handleOpenUrlStart);
const handleOpenUrlSource = rustRuntime.slice(handleOpenUrlStart, handleOpenUrlEnd);
const documentImportWindowStart = rustRuntime.indexOf("fn open_document_import_window");
const documentImportWindowEnd = rustRuntime.indexOf(
  "fn open_word_latex_redraw_window",
  documentImportWindowStart,
);
const documentImportWindowSource = rustRuntime.slice(
  documentImportWindowStart,
  documentImportWindowEnd,
);
expect(handleOpenUrlStart >= 0 && handleOpenUrlEnd > handleOpenUrlStart, "The native Office URL handler source must be discoverable");
expect(
  documentImportWindowStart >= 0 && documentImportWindowEnd > documentImportWindowStart,
  "The Word batch-import window source must be discoverable",
);
expect(!rustRuntime.includes("hide_main_window_for_office_editor"), "Opening an Office formula must not hide an already visible VisualTeX desktop workspace");
expect(!handleOpenUrlSource.includes("reveal_main_window"), "Office URL handling must never reveal the desktop main window");
expect(!handleOpenUrlSource.includes("hide_main_window"), "Office URL handling must never hide the desktop main window");
expectIncludes(rustRuntime, "pub(crate) fn consume_fast_open_request(app: &AppHandle)", "The resident macOS runtime must consume sandbox fast-open requests before revealing the desktop window");
expectIncludes(rustRuntime, "WORD_FAST_OPEN_INBOX_SUFFIX", "The resident runtime must scan the fixed Word sandbox inbox");
expectIncludes(rustRuntime, "POWERPOINT_FAST_OPEN_INBOX_SUFFIX", "The resident runtime must scan the fixed PowerPoint sandbox inbox");
expect(!rustRuntime.includes("GLOBAL_FAST_OPEN_INBOX_SUFFIX"), "The fast-open consumer must not accept a global fallback inbox outside Office containers");
expectIncludes(rustRuntime, "fast_open_session_id(&path)", "Fast-open inbox filenames must be validated as canonical Session UUID request files");
expectIncludes(rustRuntime, "metadata.file_type().is_symlink()", "Fast-open request files and roots must reject symlinks");
expectIncludes(rustRuntime, "FAST_OPEN_MAX_AGE", "Fast-open requests must expire quickly instead of remaining replayable");
expectIncludes(rustRuntime, "persist_fast_open_claim(host, &session_id, &claim_path)", "The resident must claim and validate the inbox request before importing it");
expectIncludes(rustRuntime, "atomic_write_runtime(&request_path(expected_host, session_id)?", "A validated inbox request must be copied into the canonical Application Scripts Session runtime");
expectIncludes(appRuntime, "consume_fast_open_request(app)", "The single-instance callback must check the Office sandbox inbox before revealing the VisualTeX main window");
expectIncludes(appRuntime, "start_fast_open_inbox_watcher(app.handle().clone())", "The resident app must start the low-latency Office inbox watcher after editor prewarming");
expectIncludes(rustRuntime, "FAST_OPEN_POLL_INTERVAL", "The resident Office inbox watcher must use a bounded low-latency polling interval");
expectIncludes(rustRuntime, "FAST_OPEN_MIN_STABLE_AGE", "The inbox watcher must leave a short post-rename stability window before claiming a VBA request");
expectIncludes(rustRuntime, "FAST_OPEN_READY_HEARTBEAT_INTERVAL", "The resident watcher must refresh a bounded readiness heartbeat for Office VBA");
expectIncludes(rustRuntime, "refresh_fast_open_ready_markers", "The resident watcher must publish readiness inside both fixed Office sandbox inboxes");
expectIncludes(rustRuntime, "pub(crate) fn start_fast_open_inbox_watcher", "The macOS Office runtime must expose one guarded resident inbox watcher");
expect(
  appRuntime.indexOf("consume_fast_open_request(app)") < appRuntime.indexOf("reveal_main_window(app)"),
  "A parameterless Office fast-open signal must be consumed before the desktop main window can flash",
);
expect(!rustRuntime.includes("main_was_visible"), "The Office editor lifecycle must derive background mode from current main-window visibility instead of hiding and remembering the workspace");
expect(
  (documentImportWindowSource.match(/order_main_window_behind_office_editor\(app\)\?/g) ?? []).length >= 2,
  "Both reused and newly created Word batch-import windows must send the VisualTeX main workspace behind Word after receiving focus",
);
expectIncludes(rustRuntime, "restore_office_host_focus(host)", "Closing the formula editor must return focus to Word or PowerPoint");
expectIncludes(backgroundRuntime, "prepare_foreground_app", "Office hydration must be able to prepare a regular macOS app without activating every VisualTeX window");
expectIncludes(rustRuntime, "order_main_window_behind_office_editor", "The dedicated Office editor must explicitly keep the desktop main window behind Word or PowerPoint");
expectIncludes(rustRuntime, "native_window.orderBack(None)", "The main VisualTeX workspace must be ordered behind Office instead of being raised with the editor");
expect(!rustRuntime.includes("Duration::from_millis(100)"), "Resident editor prewarming must not use a delayed hide that can race a new Office Session");
expectIncludes(backgroundRuntime, "yieldActivationToApplication", "Closing the Office editor must cooperatively yield activation back to Word or PowerPoint on modern macOS");
expectIncludes(rustRuntime, "fn make_resident_editor_key", "A fully hydrated Office editor must use one dedicated native key-window presentation path");
expectIncludes(rustRuntime, "native_window.orderFrontRegardless();", "The ready Office formula window must rise above Word or PowerPoint after the desktop main window is ordered back");
expectIncludes(rustRuntime, "native_window.makeKeyAndOrderFront(None);", "The ready Office formula window must become the AppKit key window instead of relying only on advisory Tauri focus");
expectIncludes(backgroundRuntime, "NSApplicationActivationOptions::ActivateAllWindows", "Office formula hydration must preserve the proven eb2fcf2a cooperative activation behavior");
expectIncludes(rustRuntime, "MainThreadMarker::new()", "Office focus diagnostics must inspect AppKit state only from the native WebView main-thread callback");
expect(!rustRuntime.includes("activateIgnoringOtherApps"), "The ready Office formula editor must not depend on the macOS 14-deprecated activation override");
expectIncludes(backgroundRuntime, "activate_foreground_app_via_launch_services", "A ready Office formula editor must cross the Word or PowerPoint application boundary through LaunchServices");
expectIncludes(backgroundRuntime, 'Command::new("/usr/bin/open")', "LaunchServices activation must use the fixed macOS open executable");
expectIncludes(backgroundRuntime, '.arg("-b")', "LaunchServices activation must target only the current VisualTeX bundle identifier");
expectIncludes(backgroundRuntime, "for attempt in 0..4", "Accessory-to-Regular activation must use the short retry sequence from eb2fcf2a");
expectIncludes(backgroundRuntime, "Duration::from_millis(5)", "Foreground activation retries must remain within the proven short settling interval");
expectIncludes(rustRuntime, "wake_resident_editor_for_hydration", "An Office request must restore the continuously mounted resident WebView to full native alpha before hydration");
const openEditorWindowStart = rustRuntime.indexOf("fn open_editor_window(");
const openEditorWindowEnd = rustRuntime.indexOf("fn set_word_document_operation_preparing_status", openEditorWindowStart);
const openEditorWindowSource = rustRuntime.slice(openEditorWindowStart, openEditorWindowEnd);
const presentResidentEditorStart = rustRuntime.indexOf("fn present_resident_editor_window(");
const presentResidentEditorEnd = rustRuntime.indexOf(
  "#[cfg(not(target_os = \"macos\"))]",
  presentResidentEditorStart,
);
const presentResidentEditorSource = rustRuntime.slice(
  presentResidentEditorStart,
  presentResidentEditorEnd,
);
expectIncludes(rustRuntime, "make_resident_editor_key(&window)?", "Final ready presentation must combine LaunchServices activation and native key-window ordering");
expect(!openEditorWindowSource.includes("activate_foreground_app_via_launch_services"), "Office hydration must not create a regular Dock tile while the resident editor is still parked");
expect(!openEditorWindowSource.includes("crate::office::background::activate_foreground_app(app)?"), "Office hydration must remain accessory-only until the formula window has normal geometry and opacity");
expect(
  !openEditorWindowSource.includes("set_size(") &&
    !openEditorWindowSource.includes("apply_persisted_office_editor_window_size"),
  "Reusing a resident Office editor must not perform any window-size work in the double-click opening path",
);
expect(
  !presentResidentEditorSource.includes("set_size(") &&
    !presentResidentEditorSource.includes("apply_persisted_office_editor_window_size"),
  "Window-size restoration must never sit in the foreground/key-window critical path",
);
const readyEditorStart = rustRuntime.indexOf("pub fn report_macos_offline_office_editor_ready");
const readyEditorEnd = rustRuntime.indexOf("pub fn close_macos_offline_office_editor_window", readyEditorStart);
const readyEditorSource = rustRuntime.slice(readyEditorStart, readyEditorEnd);
expect(
  readyEditorSource.indexOf("set_resident_editor_parked(&window, false)?") < readyEditorSource.indexOf("crate::office::background::activate_foreground_app(&app)?"),
  "The ready Office editor must restore its real window before macOS creates the regular Dock tile",
);
expectIncludes(rustRuntime, "crate::office::background::activate_foreground_app(app)?", "Ready presentation must force VisualTeX across the Word or PowerPoint application boundary");
expectIncludes(
  rustRuntime,
  "native_window.setLevel(objc2_app_kit::NSNormalWindowLevel - 1)",
  "Opening an Office editor must keep the visible VisualTeX desktop below Word or PowerPoint instead of raising every normal application window",
);
expectIncludes(
  rustRuntime,
  "objc2_app_kit::NSFloatingWindowLevel",
  "A ready Office formula editor must use a floating native level so it stays visually above Word or PowerPoint even when macOS rejects cross-application focus",
);
expectIncludes(
  rustRuntime,
  "if parked {\n                objc2_app_kit::NSNormalWindowLevel\n            } else {\n                objc2_app_kit::NSFloatingWindowLevel",
  "Only the visible Office editor may be promoted; parked resident windows must return to normal level",
);
expectIncludes(
  rustRuntime,
  "restore_main_window_level_after_office_editor(&app)?",
  "Closing the last Office editor must restore the VisualTeX desktop to the normal window level",
);
expect(
  presentResidentEditorSource.indexOf("order_main_window_behind_office_editor(app)?") <
    presentResidentEditorSource.indexOf("crate::office::background::activate_foreground_app(app)?"),
  "The desktop workspace must be lowered before ActivateAllWindows can raise the VisualTeX application window group",
);
expectIncludes(dialogApp, "readPersistedFormulaFontPreferences()", "The resident Office editor must read the independent global formula-font preferences instead of trusting its stale persisted editor store");
expectIncludes(dialogApp, "session?.originalMetadata?.formulaChineseFont ??", "Editing an existing Office formula must restore its saved Chinese font before falling back to the global preference");
expectIncludes(dialogApp, "session?.originalMetadata?.formulaLetterFont ??", "Editing an existing Office formula must restore its saved math-letter font before falling back to the global preference");
expectIncludes(editorStore, "persistFormulaChineseFontPreference(normalized)", "Changing the global Chinese formula font must persist independently of the resident editor document store");
expectIncludes(editorStore, "persistFormulaLetterFontPreference(normalized)", "Changing the global math-letter font must persist independently of the resident editor document store");
expectIncludes(settingsDialog, "data-save-configuration", "Settings must expose a save-current-configuration action");
expectIncludes(settingsDialog, "data-import-configuration", "Settings must expose a configuration import action");
expectIncludes(applicationConfiguration, 'VISUALTEX_CONFIGURATION_SCHEMA = "visualtex-user-configuration"', "Portable configuration files must carry a dedicated schema marker");
expectIncludes(applicationConfiguration, '"visualtex-custom-formula-tiles"', "Portable configuration must include custom formula tiles");
expectIncludes(applicationConfiguration, '"visualtex-formula-hotkeys-v1"', "Portable configuration must include formula hotkeys");
expectIncludes(applicationConfiguration, '"visualtex-custom-formula-text-colors"', "Portable configuration must include custom formula text colours");
expectIncludes(applicationConfiguration, '"visualtex-office-editor-toolbar-open"', "Portable configuration must preserve Office editor panel preferences");
expectIncludes(applicationConfiguration, '"visualtex.ocr.model"', "Portable configuration must preserve the user's preferred OCR model");
expectIncludes(applicationConfiguration, '"get_app_window_configuration"', "Configuration export must read native application window sizes");
expectIncludes(applicationConfiguration, '"apply_app_window_configuration"', "Configuration import must restore native application window sizes");
expectIncludes(appRuntime, 'const MAIN_WINDOW_SIZE_FILE: &str = "main-window-size.json"', "The main VisualTeX window size must persist independently of document content");
expectIncludes(appRuntime, "schedule_persist_main_window_size(app, size.width, size.height)", "Resizing the main VisualTeX window must persist its configuration size");
expectIncludes(appRuntime, "restore_main_window_size(app.handle())", "VisualTeX startup must restore the persisted main window size");
expectIncludes(appRuntime, "get_app_window_configuration", "The Tauri command registry must expose configuration window export");
expectIncludes(appRuntime, "apply_app_window_configuration", "The Tauri command registry must expose configuration window import");
expectIncludes(rustRuntime, "configuration_office_editor_window_size", "Configuration export must include the persisted Office editor size");
expectIncludes(rustRuntime, "apply_configuration_office_editor_window_size", "Configuration import must normalize, persist and apply the Office editor size");
expectIncludes(appRuntime, '"md" | "svg" | "png" | "vtxconfig"', "Native export writing must allow VisualTeX configuration files");
expectIncludes(styles, ".settings-subdialog-content .formula-inset-range strong", "Interface customization sliders must have dedicated label layout rules instead of inheriting the cramped generic range style");
expectIncludes(styles, "white-space: nowrap", "Interface customization parameter labels must stay horizontal");
expectIncludes(appRuntime, "office::background::install_application_icon(app.handle())", "macOS setup must install the VisualTeX application icon before any background-to-foreground transition");
expectIncludes(appRuntime, "Duration::from_secs(2)", "The initial macOS background launch must ignore LaunchServices Reopen before resident Office prewarming settles");
expectIncludes(appRuntime, "Duration::from_millis(150)", "macOS Reopen handling must wait for single-instance Office Session activation before revealing the desktop window");
expectIncludes(appRuntime, "focus_open_office_editor(&app)", "Deferred Reopen handling must suppress the desktop window while an Office editor is hydrating");
expectIncludes(
  rustRuntime,
  "pub(crate) fn has_recent_office_editor_request(max_age: Duration) -> bool",
  "The native runtime must detect an Office request that exists before its editor Session is registered",
);
expectIncludes(
  appRuntime,
  "has_recent_office_editor_request(\n                    std::time::Duration::from_secs(3)",
  "A recent Word or PowerPoint request must suppress the desktop workspace during the Reopen race",
);
expect(
  appRuntime.indexOf("has_recent_office_editor_request(") < appRuntime.indexOf("reveal_main_window(&app)"),
  "The Reopen handler must check for a pending Office request before revealing the desktop workspace",
);
expectIncludes(
  rustRuntime,
  "if !active.ready {\n                let _ = order_main_window_behind_office_editor(app);",
  "A LaunchServices Reopen during hydration must demote the desktop main window without focusing transparent editor content",
);
expectIncludes(backgroundRuntime, "Every Accessory-to-Regular transition must have the real bundle icon", "Every Office foreground transition must preserve the VisualTeX Dock icon");
expectIncludes(backgroundRuntime, 'const DOCK_ICON_MIGRATION_MARKER_FILE: &str = "dock-icon-v5.refreshed"', "The repaired Dock icon lifecycle must refresh stale same-version icon cache once");
expectIncludes(backgroundRuntime, "Install the bundle icon before changing activation policy", "Foreground reveal must install the VisualTeX icon before creating a regular Dock tile");
expectIncludes(rustRuntime, "open_editor_window(\n        app,\n        host,\n        &session_id,\n        received_epoch_ms,\n        received_at,\n        silent,", "Office formula requests must activate the fixed host editor with one generation, timing origin and explicit silent mode");
expectIncludes(rustRuntime, "office-native-dialog.html?transport=tauri", "The resident Office editor must use the direct native-dialog entry so a hidden prewarmed WebView cannot stall on the desktop entry's dynamic import");
expectIncludes(read("src/desktop/main.tsx"), 'view === "office-formula"', "The desktop entry must select the dedicated Office formula view from the window query");
expectIncludes(read("src/desktop/main.tsx"), "<OfficeDialogApp />", "The dedicated desktop window must render the Office formula editor");
expectIncludes(rustRuntime, "window.show()", "The dedicated Office formula editor must be explicitly shown");
expectIncludes(rustRuntime, "window.set_focus()", "The dedicated Office formula editor must receive focus");
expectIncludes(rustRuntime, "focus_open_office_editor", "macOS reopen handling must be able to refocus an existing Office editor");
expectIncludes(nativeInteraction, "focus_open_office_editor", "The Office compatibility double-click monitor must avoid opening duplicate editor windows");
expectIncludes(nativeInteraction, "let Some(formula_id) = formula_id else", "The global PowerPoint monitor must leave ordinary non-VisualTeX shapes to PowerPoint's normal double-click behavior");
expectIncludes(nativeInteraction, "frontmost.as_deref() == Some(WORD_BUNDLE_ID)", "Word image double-clicks must retain an app-side compatibility fallback");
expectIncludes(nativeInteraction, "word_formula_after_double_click(", "The Word compatibility branch must inspect only selected VisualTeX InlineShapes");
expectIncludes(nativeInteraction, "NSEvent::mouseLocation()", "The Word compatibility branch must capture the physical double-click location instead of trusting a stale Word selection");
expectIncludes(nativeInteraction, "selection.screen_bounds.contains(click_x, click_y)", "The Word compatibility branch must reject Ribbon or paragraph double-clicks outside the formula's real screen rectangle");
expectIncludes(nativeInteraction, "VisualTeX_WriteSelectedDoubleClickScreenBounds", "The Word compatibility branch must obtain formula bounds through Word's GetPoint-backed VBA probe");
expectIncludes(nativeInteraction, "run_word_image_double_click_edit_macro", "The Word compatibility branch must invoke the strict image-only VBA entry when an older bare InlineShape is double-clicked");
expect(!nativeInteraction.includes("crate::office::sessions::OfficeHost::Word"), "The Tauri global monitor must never invoke the generic Word double-click macro and duplicate native OMML editing");
expectIncludes(nativeInteraction, "run_double_click_edit_macro", "The native monitor must invoke the PPAM edit macro directly instead of depending on an Office.js poller");
expectIncludes(nativeInteraction, "[25_u64, 35, 60, 100]", "PowerPoint double-click selection retries must stay within the reviewed 220 ms host-settling budget");
expectIncludes(nativeInteraction, "[35_u64, 55, 85]", "Word's compatibility fallback must stay within the reviewed 175 ms wait budget");
expectIncludes(rustRuntime, 'run VB macro macro name "VisualTeX_DoubleClickEditSelected"', "The native runtime must retain the fixed Office double-click macro entry point");
expectIncludes(nativeInteraction, "push_powerpoint_edit_selected", "The compatibility monitor must preserve a fallback when the PowerPoint macro call fails");
expect(!nativeInteraction.includes("native_offline_plugin_loaded"), "The compatibility monitor must not disable itself merely because a native plug-in loaded");
expectIncludes(macSettings, '"install_macos_offline_office_addins"', "macOS Settings must install only the native DOTM/PPAM integration");
expectIncludes(macFirstRun, '"install_macos_offline_office_addins"', "macOS first-run setup must install only the native DOTM/PPAM integration");
expectIncludes(desktopApp, '"get_macos_offline_office_install_status"', "macOS startup must re-check the actual installed DOTM/PPAM files even after an earlier first-run dialog was completed");
expectIncludes(desktopApp, "!status.word.filesInstalled", "A stale or missing Word DOTM must reopen the native Office setup dialog");
expectIncludes(desktopApp, "!status.powerpoint.filesInstalled", "A stale or missing PowerPoint PPAM must reopen the native Office setup dialog");
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
expect(!macTauriConfig.includes('"../office/macos-offline/": "office/macos-offline/"'), "The macOS bundle must not copy the whole Office source tree because Word lock files such as ~$*.dotm carry forbidden Finder metadata");
for (const stableOfficeResource of ["VisualTeX.dotm", "VisualTeX.ppam", "addins.json"]) {
  expectIncludes(macTauriConfig, `../office/macos-offline/resources/${stableOfficeResource}`, `The macOS bundle must explicitly package ${stableOfficeResource}`);
}
expect(!platformBundle.includes('run(npm, ["run", "build:office:macos"])'), "The macOS build must not generate an Office.js web bundle");
expect(!lifecycle.includes("resolve_ui_root"), "macOS startup must not resolve an Office.js UI resource directory");
expectIncludes(lifecycle, "ensure_companion_runtime", "macOS startup must still initialize the private Session/OCR companion runtime");
expectIncludes(capabilities, '"office-native-*"', "Dedicated native Office windows must receive Tauri core permissions");
expectIncludes(capabilities, '"core:window:allow-close"', "Dedicated native Office windows must be allowed to close after a successful commit or cancel");
expectIncludes(capabilities, '"dialog:allow-save"', "Settings configuration backup must be allowed to open the native macOS save dialog");
expectIncludes(capabilities, '"quick-ocr-hud"', "The silent OCR HUD window must receive Tauri core permissions");
expectIncludes(quickOcrRuntime, 'Command::new("/usr/sbin/screencapture")', "Quick OCR must use the native macOS interactive screenshot tool");
expectIncludes(appRuntime, "quick_ocr::wait_for_quick_ocr_system_screenshot", "The Tauri command handler must expose the deferred system-screenshot Quick OCR mode");
expectIncludes(quickOcrRuntime, 'Command::new("/usr/bin/defaults")', "Deferred Quick OCR must read the user's current macOS screenshot destination instead of assuming Desktop");
expectIncludes(quickOcrRuntime, '"com.apple.screencapture"', "Deferred Quick OCR must resolve the macOS screenshot preference domain");
expectIncludes(quickOcrRuntime, "NSPasteboardTypePNG", "Deferred Quick OCR must also detect system screenshots routed to the clipboard");
expectIncludes(quickOcrRuntime, "create_system_screenshot_baseline", "Deferred Quick OCR must snapshot screenshot state before minimizing so fast captures are not missed");
expectIncludes(quickOcrRuntime, "SYSTEM_SCREENSHOT_WAIT_TIMEOUT", "Deferred Quick OCR must have a bounded wait instead of blocking forever");
expectIncludes(quickOcrRuntime, ".minimize()", "Quick OCR must minimize the main window before interactive capture");
expectIncludes(backgroundRuntime, ".unminimize()", "Quick OCR restoration must reuse the proven main-window unminimize path");
expectIncludes(quickOcrRuntime, 'Command::new("/usr/bin/pbcopy")', "Silent OCR must copy recognized LaTeX without revealing the main WebView");
expectIncludes(quickOcrRuntime, "format_silent_ocr_latex", "Silent OCR clipboard output must pass through the selected VisualTeX LaTeX source formatter before pbcopy");
expectIncludes(quickOcrRuntime, "write_text_clipboard(&formatted_latex)", "Silent OCR must copy the formatted LaTeX rather than the raw OCR result");
expectIncludes(quickOcrRuntime, '"display-bracket"', "Silent OCR formatting must support the VisualTeX display-bracket source format");
expectIncludes(quickOcrRuntime, '"equation-star-split"', "Silent OCR formatting must support the full VisualTeX environment format set");
expectIncludes(desktopApp, "configureSilentOcr(silentOcrEnabled, ocrModel, latexCodeFormat)", "The desktop app must synchronize the current LaTeX source format into the native silent OCR runtime");
expectIncludes(desktopApp, "if (!runtime.installedModels.includes(ocrModel))", "The desktop OCR path must preserve an unavailable user-selected model and route the user to model installation instead of silently switching to M");
expect(!desktopApp.includes("setOcrModel(availableModel)"), "Desktop OCR prewarm must not overwrite the user's selected S/L model with an installed fallback");
expectIncludes(dialogApp, "if (!runtime.installedModels.includes(ocrModel))", "The native Office OCR path must preserve an unavailable user-selected model and route the user to model installation instead of silently switching to M");
expect(!dialogApp.includes("setOcrModel(availableModel)"), "Office OCR prewarm must not overwrite the user's selected S/L model with an installed fallback");
expectIncludes(quickOcrRuntime, "RegisterEventHotKey", "Silent OCR must register a true macOS global hotkey instead of relying on foreground DOM keyboard events");
expectIncludes(quickOcrRuntime, "run_on_main_thread", "Carbon hotkey registration must run on the AppKit main thread because RegisterEventHotKey is not thread safe");
expectIncludes(quickOcrRuntime, 'WebviewUrl::App("index.html?view=quick-ocr-hud".into())', "Silent OCR progress must use the dedicated lightweight HUD rather than revealing the main application");
expectIncludes(editorWorkspace, "data-quick-ocr-button", "The desktop editor toolbar must expose a Quick OCR capture button");
expectIncludes(editorWorkspace, "data-quick-ocr-mode-trigger", "Quick OCR must expose a compact capture-mode selector");
expectIncludes(editorWorkspace, 'data-quick-ocr-mode-option="system-screenshot"', "Quick OCR must let the user choose the deferred system-screenshot mode");
expectIncludes(editorWorkspace, 'data-quick-ocr-mode-option="immediate"', "Quick OCR must preserve the immediate-selection mode");
expectIncludes(editorWorkspace, "data-silent-ocr-toggle", "The desktop editor toolbar must expose a silent OCR mode toggle");
expectIncludes(applicationConfiguration, '"visualtex.silent-ocr.enabled"', "Silent OCR enablement must migrate with VisualTeX user configuration backups");
expectIncludes(applicationConfiguration, '"visualtex.quick-ocr.capture-mode"', "Quick OCR capture mode must migrate with VisualTeX user configuration backups");
expect(!inputBehaviorMenu.includes("option.descriptionEn"), "Input behavior cards must not render explanatory subtext below each option title");
expect(!inputBehaviorMenu.includes("控制普通数学输入是否使用快捷转义"), "Input behavior section headings must not render explanatory subtext");
expect(!inputBehaviorMenu.includes("分别选择哪些单槽结构"), "Caret auto-exit heading must not render explanatory subtext");
expect(!inputBehaviorMenu.includes("这里只控制 VisualTeX 的大型命令候选框"), "Command suggestion heading must not render explanatory subtext");
expect(!updateDialog.includes("关闭后不会自动联网检查"), "Update preference must not render the removed explanatory subtext");
expectIncludes(dialogApp, "isMacosOfflineTauriTransport()", "Native Office formula editors must avoid Office.js parent messaging");
expectIncludes(dialogApp, '"visualtex.office.word.create.numbered"', "Word create dialogs must persist the last Add Equation Number choice");
expectIncludes(dialogApp, "readOfficeWordCreateNumberedPreference(sessionNumbered)", "A new Word display formula must restore the remembered numbering choice");
expectIncludes(dialogApp, "session.mode === \"create\" &&\n      session.displayMode === \"block\"", "Only new Word display formulas may inherit the remembered numbering choice");
expectIncludes(dialogApp, "writeOfficeWordCreateNumberedPreference(nextNumbered)", "Changing Add Equation Number must update the remembered create preference immediately");
expectIncludes(dialogApp, "const residentEditorWorkspace", "The native Office window must keep its editor workspace mounted while parked");
expectIncludes(dialogApp, 'className="office-resident-editor-workspace"', "The resident Office editor wrapper must have an explicit layout class instead of becoming an anonymous Grid item");
const officeDialogShellStyles =
  styles.match(/\.office-dialog-shell\s*\{([^}]*)\}/)?.[1] ?? "";
const officeResidentWorkspaceStyles =
  styles.match(/\.office-resident-editor-workspace\s*\{([^}]*)\}/)?.[1] ?? "";
expect(
  /grid-template-rows:\s*minmax\(0,\s*1fr\)\s*;/.test(
    officeDialogShellStyles,
  ),
  "The Office dialog shell must allocate the full height to the unified resident editor",
);
expect(
  /display:\s*grid\s*;/.test(officeResidentWorkspaceStyles) &&
    /grid-template-rows:\s*minmax\(0,\s*1fr\)\s+auto\s*;/.test(
      officeResidentWorkspaceStyles,
    ),
  "The compact resident Office editor must allocate only the flexible editor and bottom status rows",
);
expectIncludes(
  dialogApp,
  "showOfficeActions={false}",
  "The Office formula window must not render the redundant editing-description action row",
);
expectIncludes(
  dialogApp,
  "officeHeaderLeadingControls={officeHeaderLeadingControls}",
  "Office layout, font-size and auto-apply controls must be injected into the unified editor toolbar",
);
expectIncludes(
  dialogApp,
  "officeHeaderTrailingActions={officeHeaderTrailingActions}",
  "Office history and apply controls must be injected into the unified editor toolbar",
);
expectIncludes(
  dialogApp,
  "data-office-cancel-action",
  "The compact Office header must expose Cancel in its unified action row",
);
expectIncludes(
  dialogApp,
  "data-office-primary-action",
  "The compact Office header must expose Update/Insert in its unified action row",
);
expectIncludes(
  dialogApp,
  "data-office-undo-action",
  "Undo must use an icon-only button in the unified Office toolbar",
);
expectIncludes(
  dialogApp,
  "data-office-redo-action",
  "Redo must use an icon-only button in the unified Office toolbar",
);
expectIncludes(dialogApp, "<Undo2", "Undo must render an arrow icon instead of a text label");
expectIncludes(dialogApp, "<Redo2", "Redo must render an arrow icon instead of a text label");
expect(
  !dialogApp.includes('<header className="office-dialog-header"'),
  "The redundant VisualTeX / Office host title bar must no longer render",
);
expectIncludes(
  editorWorkspace,
  "is-office-editor-header",
  "Office controls and formula tools must share the existing editor toolbar row",
);
expect(
  !editorWorkspace.includes('isEn ? "Visual editor" : "可视化编辑"'),
  "The redundant Visual editor title must be omitted from desktop and Office workspaces",
);
expect(
  !editorWorkspace.includes('<span className="pane-icon" aria-hidden="true">'),
  "The redundant Visual editor braces icon must be omitted from the desktop editor toolbar",
);
expect(
  !dialogApp.includes("<ScanLine size={15} />"),
  "The compact Office header must not expose the standalone Image OCR button",
);
expectIncludes(dialogApp, 'editorInstanceKey="resident-office-editor"', "Word and PowerPoint Sessions must reuse one host-resident MathLive editor instead of rebuilding it per formula");
expect(!dialogApp.includes("editorInstanceKey={session.id}"), "A new Office Session must not destroy and remount the resident MathLive editor");
expectIncludes(dialogApp, "const contentReadyDeadlineMs = origin + 5_000", "Cold MathLive mounting must use a bounded readiness deadline instead of a fixed frame count");
expect(!dialogApp.includes("!contentMounted && frame < 12"), "The ready probe must not permanently stop after twelve animation frames");
expectIncludes(dialogApp, "commitMacosOfflineOfficeSession(session.id, update)", "Native Apply must combine its final Session patch and Office commit into one Tauri round trip");
expectIncludes(dialogApp, "completeExportInFlightRef", "Apply must reuse an in-flight PNG export instead of starting duplicate rasterization");
expectIncludes(dialogApp, "getCompleteExportResult(currentFingerprint, exportResult)", "Autosave must reuse the SVG export when producing its PNG compatibility result");
expectIncludes(dialogApp, "await getCompleteExportResult(currentFingerprint)", "Apply must reuse the exact cached export for the current formula fingerprint");
expectIncludes(rustRuntime, "fn atomic_write_runtime", "Ephemeral Office artifacts must avoid crash-durability barriers on the synchronous Apply path");
expectIncludes(rustRuntime, '"apply-backend-complete"', "The native commit path must persist an end-to-end Apply timing stage for the accepted performance budget");
expectIncludes(rustRuntime, "patch: Option<Value>", "Native Apply must accept the final editor state in the same Tauri command as the commit");
expectIncludes(rustRuntime, "Unable to refresh the VisualTeX formula cache after Apply", "Formula-cache maintenance must run after the durable Office callback without blocking Apply");
expectIncludes(tauriTransport, 'from "@tauri-apps/api/window"', "Native Office formula editors must control the actual Tauri window through the shared transport");
expectIncludes(tauriTransport, "getCurrentWindow().onCloseRequested", "Closing a native formula window must listen to the actual Tauri close request");
expectIncludes(dialogApp, "onCurrentTauriWindowCloseRequested", "Closing a native formula window must finalize or cancel its Office transaction before destruction");
expectIncludes(dialogApp, "close_macos_offline_office_editor_window", "A successful native Office transaction must hide and clear the resident Tauri editor window");
expectIncludes(rustRuntime, '"office-native-word-editor"', "Word must use one fixed resident editor window label");
expectIncludes(rustRuntime, '"office-native-powerpoint-editor"', "PowerPoint must keep a separate resident editor window label");
expectIncludes(rustRuntime, "DEFAULT_OFFICE_EDITOR_WIDTH_RATIO", "The Office editor default width must be derived from the measured screen proportion instead of a fixed cross-device pixel size");
expectIncludes(rustRuntime, "DEFAULT_OFFICE_EDITOR_HEIGHT_RATIO", "The Office editor default height must be derived from the measured screen proportion instead of a fixed cross-device pixel size");
expectIncludes(rustRuntime, "width_ratio", "User-resized Office editor geometry must persist as a screen-relative width ratio");
expectIncludes(rustRuntime, "height_ratio", "User-resized Office editor geometry must persist as a screen-relative height ratio");
expectIncludes(rustRuntime, "current_monitor()", "Office editor resizing must use the monitor that actually contains the window");
expectIncludes(rustRuntime, "let saved_size = load_office_editor_window_size(app);", "A resident Office editor must restore the saved shared size only when its native window is first created");
expectIncludes(appRuntime, "WindowEvent::Resized(size)", "Native Office editor resize events must persist the user's screen-relative geometry");
expectIncludes(editorStore, "DEFAULT_EDITOR_ZOOM = 0.45", "A fresh editor profile must start at the reviewed 45% main-editor zoom");
expectIncludes(editorStore, "EDITOR_ZOOM_STEP = 0.05", "Desktop and Office formula zoom must use a shared 5% step");
expectIncludes(editorStore, "zoom / EDITOR_ZOOM_STEP", "Persisted zoom normalization must preserve 5% increments instead of rounding back to 10% steps");
expectIncludes(editorWorkspace, "zoom - EDITOR_ZOOM_STEP", "The shared zoom-out button must decrease formula zoom by 5%");
expectIncludes(editorWorkspace, "zoom + EDITOR_ZOOM_STEP", "The shared zoom-in button must increase formula zoom by 5%");
expectIncludes(desktopApp, "zoom + EDITOR_ZOOM_STEP", "Desktop keyboard zoom-in must use the same 5% step");
expectIncludes(desktopApp, "zoom - EDITOR_ZOOM_STEP", "Desktop keyboard zoom-out must use the same 5% step");
expectIncludes(styles, ".office-inline-actions .primary-button,\n.office-inline-actions .secondary-button {\n  align-self: center;", "Cancel and Update Formula must share one vertical alignment rule in the Office header");
expectIncludes(styles, "min-height: 30px;\n  margin: 0;", "Office header actions must clear the global secondary-button top margin that previously displaced Cancel");
expectIncludes(editorStore, "showLineNumbers: false", "Formula row numbers must default off for both desktop and Office editors");
expectIncludes(editorStore, "showLineNumbers: state.showLineNumbers", "The optional formula row-number preference must persist across application restarts");
expectIncludes(editorStore, "setShowLineNumbers", "The shared editor store must expose an explicit formula row-number setting");
expectIncludes(settingsDialog, 'data-interface-customization-trigger', "Settings must expose Interface customization as a nested-dialog entry point");
expectIncludes(settingsDialog, 'data-interface-customization-dialog', "Interface customization options must render inside their own popup dialog");
expectIncludes(settingsDialog, 'data-show-line-numbers-setting', "The nested Interface customization dialog must expose the formula row-number preference");
expectIncludes(mathEditor, 'showLineNumbers ? " has-line-numbers" : ""', "The formula canvas must reserve its number column only when the setting is enabled");
expectIncludes(mathEditor, "{showLineNumbers ? (", "Formula row labels must not render at all while their preference is disabled");
expectIncludes(styles, ".multi-line-editor.has-line-numbers .formula-line", "The formula number column must be scoped to the optional line-number class");
expectIncludes(styles, ".multi-line-editor:not(.has-line-numbers)[data-formula-alignment=\"left\"]", "A left-aligned formula must use the full row width when line numbers are hidden");
expectIncludes(mathEditor, "scheduleUnconsumedMathSpace", "Math input must have a fallback that inserts spacing only after higher-priority Space handlers decline the key");
expectIncludes(mathEditor, 'field.insert("\\\\ ", {', "The ordinary Space fallback must store an explicit LaTeX control space");
expectIncludes(mathEditor, "beforeNativeSpace.latex !== afterNativeSpace.latex", "Native Space command completions must take priority over visible-space insertion");
expectIncludes(mathEditor, "JSON.stringify(beforeNativeSpace.selection)", "Structural Space navigation must take priority when MathLive changes the selection");
expectIncludes(editorStore, "highlightActiveLine: false", "Active formula-row highlighting must default off for both desktop and Office editors");
expectIncludes(editorStore, "highlightActiveLine: state.highlightActiveLine", "The active-row highlight preference must persist across application restarts");
expectIncludes(editorStore, "setHighlightActiveLine", "The shared editor store must expose an explicit active-row highlight setting");
expectIncludes(editorStore, "DEFAULT_FORMULA_INSET = 6", "Formula-area left and right spacing must have one stable shared default");
expectIncludes(editorStore, "formulaInsetLeft: state.formulaInsetLeft", "The independent left formula-area spacing must persist across application restarts");
expectIncludes(editorStore, "formulaInsetRight: state.formulaInsetRight", "The independent right formula-area spacing must persist across application restarts");
expectIncludes(editorStore, "setFormulaInsetLeft", "The shared editor store must expose an independent left formula-area spacing setter");
expectIncludes(editorStore, "setFormulaInsetRight", "The shared editor store must expose an independent right formula-area spacing setter");
expectIncludes(editorStore, "DEFAULT_FORMULA_TOOL_BUTTON_SIZE = 38", "Formula toolbar buttons must have one stable default cell size");
expectIncludes(editorStore, "DEFAULT_FORMULA_TOOL_BUTTON_PADDING = 2", "Formula toolbar symbols must have one stable content inset");
expectIncludes(editorStore, "DEFAULT_FORMULA_ROW_VERTICAL_INSET = 2", "Formula rows must have one stable default vertical content inset");
expectIncludes(editorStore, "DEFAULT_CLASSIC_TILE_WIDTH = 220", "A fresh classic layout must start with the reviewed tile width");
expectIncludes(editorStore, "DEFAULT_CLASSIC_DOCK_HEIGHT = 245", "A fresh classic layout must start with the reviewed dock height");
expectIncludes(editorStore, "formulaToolButtonSize: state.formulaToolButtonSize", "The formula toolbar button size must persist across application restarts");
expectIncludes(editorStore, "formulaToolButtonPadding: state.formulaToolButtonPadding", "The formula toolbar content inset must persist across application restarts");
expectIncludes(editorStore, "formulaRowVerticalInset: state.formulaRowVerticalInset", "The visual formula-row vertical spacing must persist across application restarts");
expectIncludes(settingsDialog, 'data-highlight-active-line-setting', "The nested Interface customization dialog must expose the active-row highlight preference");
expectIncludes(settingsDialog, 'data-formula-inset-left-setting', "Interface customization must expose an independent left formula-area spacing slider");
expectIncludes(settingsDialog, 'data-formula-inset-right-setting', "Interface customization must expose an independent right formula-area spacing slider");
expectIncludes(settingsDialog, 'data-formula-tool-button-size-setting', "Interface customization must expose a formula toolbar button-size slider");
expectIncludes(settingsDialog, 'data-formula-tool-button-padding-setting', "Interface customization must expose a formula toolbar content-inset slider");
expectIncludes(settingsDialog, 'data-formula-row-vertical-inset-setting', "Interface customization must expose the visual editor formula-row vertical-spacing slider");
expectIncludes(settingsDialog, 'data-formula-inset-preview', "Interface customization must include a live simulated editor preview below the spacing sliders");
expectIncludes(settingsDialog, 'latex={String.raw`x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}`}', "The interface preview must pass single-backslash LaTeX instead of rendering doubled commands as text");
expectIncludes(mathEditor, '"--formula-area-inset-left"', "The real desktop and Office formula canvas must consume the persisted left spacing preference");
expectIncludes(mathEditor, '"--formula-area-inset-right"', "The real desktop and Office formula canvas must consume the persisted right spacing preference");
expectIncludes(mathEditor, '"--formula-row-vertical-inset"', "The real desktop and Office formula rows must consume the persisted top-and-bottom content spacing preference");
expectIncludes(settingsDialog, 'isEn ? "Interface customization" : "界面自定义"', "Interface customization must remain a dedicated popup entry and dialog title");
expectIncludes(editorWorkspace, 'highlightActiveLine ? " has-active-line-highlight" : ""', "Desktop and Office workspaces must opt into active-row highlighting only when the preference is enabled");
expectIncludes(styles, ".workspace.has-active-line-highlight .formula-line.is-active", "Active-row colour and accent styling must be gated by the persisted interface preference");
expectIncludes(styles, ".workspace.has-active-line-highlight .formula-line:hover", "Mouse-hover row feedback must be disabled together with active-row highlighting");
expectIncludes(styles, "border-bottom: 1px solid color-mix", "Every formula row, including a single row, must draw its own lower divider");
expect(
  !styles.includes(".formula-line + .formula-line {"),
  "Formula dividers must not depend on a following row",
);
expectIncludes(styles, ".workspace.is-classic-layout:not(.is-office-workspace)", "Desktop classic layout must reserve a dedicated full-width toolbar row above the editor and tiles");
expectIncludes(styles, "> .editor-pane-header {\n  grid-column: 1 / -1;\n  grid-row: 1;", "The desktop editor toolbar must span across the formula canvas and right-side tiles");
expectIncludes(styles, "> .classic-tile-toolbar {\n  grid-column: 3;\n  grid-row: 2;", "Desktop formula tiles must begin below the full-width editor toolbar");
expectIncludes(styles, ".editor-pane-scroll > .editor-surface.multi-line-editor {\n  padding-top: 6px;", "Desktop and Office editors must keep only a compact top inset above the first formula row");
expectIncludes(dialogApp, "syncOfficeEditorSystemSettings", "The Office editor must inherit the main application's persisted interface settings instead of forcing its own zoom migration");
expectIncludes(dialogApp, "EDITOR_PERSISTENCE_STORAGE_KEY", "The Office editor must read the same persisted editor settings used by the main application");
expectIncludes(workspacePanelPreferences, 'toolbar: "visualtex-office-editor-toolbar-open"', "The Office formula toolbar collapsed state must use a durable Office-specific preference");
expectIncludes(workspacePanelPreferences, 'tiles: "visualtex-office-editor-tiles-open"', "The Office formula tile collapsed state must use a durable Office-specific preference");
expectIncludes(workspacePanelPreferences, 'mode === "desktop" ? "desktop" : "office"', "Office create and edit windows must share one panel preference scope without affecting the desktop workspace");
expectIncludes(dialogApp, 'readWorkspacePanelOpen("office-edit", "tiles")', "The resident Office editor must restore its formula-tile state before the first render");
expectIncludes(dialogApp, 'writeWorkspacePanelOpen("office-edit", "tiles", open)', "Every Office tile expand/collapse action must persist immediately");
expectIncludes(editorWorkspace, 'readWorkspacePanelOpen(mode, "toolbar")', "The shared workspace must restore the bottom toolbar state before rendering");
expectIncludes(editorWorkspace, 'writeWorkspacePanelOpen(mode, "toolbar", resolved)', "Every bottom toolbar expand/collapse action must persist immediately");
expectIncludes(styles, ".formula-toolbar.is-horizontal .template-strip", "The horizontal formula tool strip must have an explicit overflow policy");
expectIncludes(styles, "grid-auto-columns: var(--formula-toolbar-button-size, 52px)", "The horizontal formula toolbar must consume the persisted user-defined cell width");
expectIncludes(styles, "align-content: stretch", "Horizontal formula rows must share the available height instead of leaving a large unused gap below three rows");
expectIncludes(styles, "var(--formula-toolbar-button-padding, 2px)", "Formula command symbols must consume the restored single toolbar content inset");
expectIncludes(styles, "var(--formula-row-vertical-inset, 5px)", "Formula characters must consume the visual formula-row top-and-bottom spacing preference");
expectIncludes(formulaToolbar, "maximumToolbarPreviewInsetRatio - formulaToolButtonPadding * 0.03", "The restored toolbar content inset must affect the real MathPreview fit algorithm");
expectIncludes(formulaToolbar, "formulaToolButtonSize / 42", "Larger formula buttons must allow simple symbols to scale up while complex formulas remain contained");
expectIncludes(formulaToolbar, "target.scrollBy({ left: pixelDelta", "The horizontal formula toolbar must map every wheel delta directly to continuous horizontal movement without category-boundary gating");
expectIncludes(formulaToolbar, "scrollToToolbarCategory", "Clicking a category tab must still provide optional smooth navigation to that continuous section");
expectIncludes(formulaToolbar, 'behavior: "auto"', "Wheel scrolling must remain unsnapped so one gesture can move naturally across multiple categories");
expectIncludes(formulaToolbar, 'className="toolbar-category-section"', "Every formula category must remain rendered in one continuous horizontal strip");
expectIncludes(formulaToolbar, "data-toolbar-category-section={category}", "Continuous toolbar sections must expose stable category boundaries for scrolling and testing");
expect(!formulaToolbar.includes("toolbar-category-transition"), "The obsolete large next-category guide card must not be rendered");
expectIncludes(styles, ".template-strip.is-continuous-categories", "Horizontal formula categories must share one continuous scroll container");
expectIncludes(styles, "gap: 14px", "Adjacent formula categories must use only a compact visual gap");
expectIncludes(formulaToolbar, '"cases-three":', "The three-row piecewise preview must remain registered in the toolbar preview map");
expectIncludes(formulaToolbar, "is-enlarged-cases-preview", "Piecewise previews must receive dedicated enlarged fitting instead of the generic tiny formula scale");
expectIncludes(styles, "gap: 0", "Formula command cells must touch without loose gaps");
expectIncludes(styles, "background: linear-gradient(", "Formula command cells must use a subtle inner highlight gradient");
expectIncludes(formulaToolbar, 'type MatrixDelimiter =', "The custom matrix picker must keep an explicit supported delimiter set");
expectIncludes(formulaToolbar, '| "Bmatrix"', "The matrix picker must include brace-delimited matrices");
expectIncludes(formulaToolbar, '| "Vmatrix"', "The matrix picker must include double-bar matrices");
expectIncludes(additionalCommands, 'makeCommand("matrix-dots"', "The expanded toolbar must include an ellipsis matrix template");
expectIncludes(additionalCommands, '["Uparrow", "\\\\Uparrow"', "The expanded toolbar must include additional double-line arrows");
expectIncludes(additionalCommands, '["plus", "+"', "The expanded relation palette must include the basic plus operator");
expectIncludes(additionalCommands, '["ni", "\\\\ni"', "The expanded set palette must include reverse membership");
expectIncludes(toolbarCommandOmmlRegression, "normalizeMathfieldLatexForRegression", "Every newly added toolbar command must be checked after real MathLive insertion normalization");
expectIncludes(toolbarCommandOmmlRegression, "latexLinesToOmml", "Every newly added toolbar command must be checked through the real Word OMML conversion path");
expectIncludes(styles, "overflow-y: hidden", "The horizontal formula tool strip must never expose a vertical scrollbar");
expect(!rustRuntime.includes("native_window.orderOut(None)"), "The resident Office editor must remain in AppKit ordering so WebKit is never suspended between formula edits");
expect(!rustRuntime.includes("Unable to hide the resident Office editor"), "The macOS resident editor must not use Tauri hide as part of its idle lifecycle");
expectIncludes(rustRuntime, "setAlphaValue(if parked { 0.001 } else { 1.0 })", "Parking must keep a non-zero native alpha without leaving a visible resident Office window on the desktop");
expectIncludes(rustRuntime, "native_window.setAlphaValue(1.0)", "Hydration must restore the resident editor to full native opacity immediately after an Office request");
expectIncludes(rustRuntime, "native_window.setIgnoresMouseEvents(true)", "A parked or hydrating resident editor must never intercept user input");
expectIncludes(rustRuntime, "present_resident_editor_window", "A hydrated resident editor must restore native mouse and focus state before accepting input");
expectIncludes(rustRuntime, "native_window.setIgnoresMouseEvents(parked)", "The resident editor must disable click-through whenever its parked state becomes false after hydration");
expectIncludes(rustRuntime, "window.set_focus()", "A ready resident editor must receive keyboard focus after the second foreground activation");
expect(!rustRuntime.includes("request_office_editor_foreground_activation"), "Office foreground activation must not use the ineffective LaunchServices reopen detour");
expect(!rustRuntime.includes("NSWorkspaceOpenConfiguration::configuration()"), "Office foreground activation must follow the released 1.2.3 AppKit path instead of reopening the application");
expect(!rustRuntime.includes("ActivateIgnoringOtherApps"), "Office foreground activation must not rely on the deprecated macOS ignoringOtherApps option");
expectIncludes(rustRuntime, "native_window.setLevel(objc2_app_kit::NSNormalWindowLevel)", "A visible resident editor must return to the normal macOS window level");
expectIncludes(rustRuntime, "ready: false", "A newly activated resident editor generation must not be focusable before frontend readiness");
expectIncludes(rustRuntime, "runtime.next_generation", "Resident editor activation must reject stale readiness and close callbacks by generation");
expectIncludes(dialogApp, "公式已经插入，但编辑窗口无法自动关闭", "A close failure after a successful native commit must not be reported as an insertion failure");
expect(!dialogApp.includes("无法插入 PowerPoint 公式"), "The shared Office editor must not mislabel Word failures as PowerPoint insertion failures");
expectIncludes(dialogApp, "latex.trim() && autoCommitOnClose", "Closing a non-empty native editor must commit when auto-apply is enabled");
expectIncludes(dialogApp, "await handleCancel();", "Closing an empty native editor must cancel and remove the pending host object");
expectIncludes(dialogMessages, 'typeof ui.messageParent !== "function"', "Office parent messaging must tolerate native Tauri windows without Office.js");
expectIncludes(appRuntime, "initial_office_url", "Cold Office URL launches must be recognized before the main workspace is revealed");
expectIncludes(appRuntime, "if office::macos_offline::focus_open_office_editor(&app)", "macOS reopen must prefer an Office formula editor over the main workspace after the single-instance activation grace period");
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
  "unregisterMacosDevUrlHandler",
  "Tauri development shutdown must release visualtex:// back to the production application",
);
expectIncludes(
  read("scripts/register_macos_dev_url_handler.mjs"),
  '$("com.visualtex.studio")',
  "Development URL cleanup must restore the production VisualTeX bundle identifier",
);
expectIncludes(
  appRuntime,
  "claim_production_visualtex_url_handler",
  "The production application must reclaim visualtex:// from stale development handlers on launch",
);
expectIncludes(
  appRuntime,
  'LSSetDefaultHandlerForURLScheme($("visualtex"), $("com.visualtex.studio"))',
  "The production application must explicitly register its fixed visualtex URL scheme handler",
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
expectIncludes(rustRuntime, '("fontSizePt", format!', "Rust must dispatch the target Word formula point size");
expectIncludes(rustRuntime, '"referenceWidthPt"', "Rust must dispatch the 14 pt reference width");
expectIncludes(rustRuntime, '"referenceHeightPt"', "Rust must dispatch the 14 pt reference height");
expectIncludes(rustRuntime, '"referenceBaselinePt"', "Rust must dispatch the 14 pt reference baseline");
expectIncludes(rustRuntime, "WORD_REFERENCE_FONT_SIZE_PT", "Rust Word geometry must use a stable reference point size");
expectIncludes(rustRuntime, "WORD_TIMES_IMAGE_WIDTH_SCALE: f64 = 1.067", "Times Word images must keep only the measured small horizontal MathJax-layout compensation");
expectIncludes(rustRuntime, "WORD_TIMES_IMAGE_HEIGHT_SCALE: f64 = 1.0", "Times Word images must not inherit the KaTeX 10 percent vertical enlargement");
expectIncludes(rustRuntime, 'Some("times") => (WORD_TIMES_IMAGE_WIDTH_SCALE, WORD_TIMES_IMAGE_HEIGHT_SCALE)', "Word geometry must select the Times-specific two-axis visual calibration from formula metadata");
expectIncludes(rustRuntime, "word_times_image_geometry_matches_native_visual_calibration", "Rust must lock the measured Times image-versus-OMML geometry regression");
expectIncludes(rustRuntime, "source_is_native_word_equation != request.native_equation", "An unchanged edit must still commit when the user requested image/OMML conversion");
expectIncludes(rustRuntime, "cleanup_session_files_at", "Completed and cancelled Sessions must remove known local request artifacts");
expectIncludes(rustRuntime, "DirectoryNotEmpty", "Session cleanup must preserve unknown files instead of deleting an entire directory recursively");
expectIncludes(rustRuntime, "com.microsoft.Word/VisualTeXRuntime", "The desktop runtime must read Word's Application Scripts Session root");
expectIncludes(rustRuntime, "com.microsoft.Powerpoint/VisualTeXRuntime", "The desktop runtime must read PowerPoint's Application Scripts Session root");
expectIncludes(rustRuntime, "for host in [OfficeHost::Word, OfficeHost::Powerpoint]", "The desktop runtime must search both host-specific Session roots by UUID");
expectIncludes(rustRuntime, "fs::symlink_metadata", "Offline Session requests must reject symbolic-link substitution before reading");
expectIncludes(rustRuntime, "set_mode(&root, 0o700)", "Each host runtime directory must be private to its owner");
expectIncludes(packager, 'kind === "Word" ? "VTWordEvents" : "VTPowerPointEvents"', "The add-in packager must reject binaries missing the Office event class module");
expectIncludes(packager, '"App_WindowSelectionChange"', "The Word add-in packager must reject a DOTM missing the image-number selection repair event");
expectIncludes(installer, "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ", "Installer must use a transparent Word placeholder image");
expectIncludes(installer, "VisualTeX.dotm", "Installer must preserve the fixed Word filename");
expectIncludes(installer, "VisualTeX.ppam", "Installer must preserve the fixed PowerPoint filename");
expectIncludes(installer, '["User Content.localized", "User Content"]', "Installer must inspect localized and unlocalized Word user-content roots");
expectIncludes(installer, '["Startup.localized", "Startup"]', "Installer must overwrite every existing Word Startup variant instead of a detached staging copy");
expectIncludes(installer, 'offline_root()?.join("OfficeAddins/VisualTeX.ppam")', "Installer must overwrite the fixed PowerPoint path already registered by Office");
expectIncludes(packager, '"UBF8T346G9.Office"', "The packager's macOS install option must synchronize the active Office files");
expectIncludes(installer, 'home.join("Applications")', "Installer must detect per-user Office application installs");
expectIncludes(installer, "restore_backups(&backups)", "Installer must roll back every staged file after a partial failure");
expectIncludes(installer, 'remove_if_exists(&word_health)', "Installer must clear stale Word health after an update");
expectIncludes(installer, 'remove_if_exists(&powerpoint_health)', "Installer must clear stale PowerPoint health after an update");
expectIncludes(installer, "health_is_current", "Installer must validate exact host/version health records");
expectIncludes(installer, "macro_responsive", "A running current add-in that answers its fixed macro must not be reported as unloaded solely because the optional health file is missing");
expectIncludes(installer, "Err(error) => (false, false, None, Some(error))", "A corrupt health file must degrade one host instead of failing the whole status view");
expectIncludes(installer, "word_paths.extend(word_support_paths)", "Word installed status must include its active Startup DOTM, AppleScriptTask and placeholder resources");
expectIncludes(installer, "addin_installation_matches", "Installed status must require byte-identical add-ins and reject stale loadable copies");
expectIncludes(installer, "files_present", "Installer status must distinguish a stale installed add-in from a first-time missing add-in");
expectIncludes(installer, "if !files_installed", "A stale Office health record must never report a missing or byte-mismatched DOTM/PPAM as currently loaded");
expectIncludes(installer, "if paths.is_empty() && root.is_dir()", "Standard Office installs must not recursively scan unrelated group-container packages before checking the known Word Startup path");
expectIncludes(installer, "read_directory_with_interrupted_retry", "Office add-in discovery must retry transient interrupted directory reads instead of suppressing update detection");
expectIncludes(installer, "request_office_hosts_quit_for_update", "The updater must request a normal Word and PowerPoint quit before replacing loaded VBA add-ins");
expectIncludes(desktopApp, "staleInstalledAddins", "Desktop startup must distinguish an existing stale Office add-in from first-time setup");
expectIncludes(desktopApp, 'setMacOfficePromptMode("repair")', "A previously configured but missing Office add-in must use repair mode instead of reopening first-time setup");
expectIncludes(desktopApp, "setPowerpointRegistrationRequired", "Desktop startup must remember whether PowerPoint genuinely needs first-time PPAM registration before opening the Office prompt");
expectIncludes(desktopApp, 'setMacOfficePromptMode("update")', "A stale DOTM or PPAM must enter the explicit Office add-in update flow");
expectIncludes(desktopApp, "setMacOfficeFirstRunOpen(true)", "A stale Office add-in must always show the update prompt instead of being replaced silently");
expectIncludes(desktopApp, "Office 插件已更新到当前 VisualTeX 版本", "A completed Office add-in update must visibly confirm that the current VisualTeX version is installed");
expectIncludes(macFirstRun, '"request_quit_macos_office_hosts_for_addin_update"', "When Office is running, the update prompt must continue the add-in update without manual file deletion");
expectIncludes(macFirstRun, "无需手动删除旧文件", "The macOS update prompt must explicitly tell users that stale DOTM and PPAM files are replaced automatically");
expectIncludes(macFirstRun, "原有登记保持不变", "A current PowerPoint PPAM must not be mislabeled as requiring first-time registration during a Word repair");
expectIncludes(macFirstRun, "修复 VisualTeX Office 插件", "Missing files after an earlier setup must be presented as repair rather than first-time configuration");
expectIncludes(installer, "powerpoint_script.clone()", "PowerPoint installed status must include its AppleScriptTask resource");
expectIncludes(installer, 'health.plugin_version.as_deref() == Some(env!("CARGO_PKG_VERSION"))', "Installer must reject stale plug-in health versions");
expect(!installer.includes("source_revision_matches"), "Runtime health must not reject a current-version add-in only because an optional sourceRevision field is absent");
expectIncludes(packager, "word-office-performance-20260801-r87", "Packaging must reject a Word DOTM that lacks the current performance revision");
expectIncludes(packager, "powerpoint-office-performance-20260801-r4", "Packaging must reject a PowerPoint PPAM that lacks the current performance revision");
expectIncludes(installer, "POWERPOINT_VBA_SOURCE_REVISION", "Installer validation must reject a stale PowerPoint PPAM without SVG point-size support");
expectIncludes(installer, "Library/Application Scripts/com.microsoft.Word", "Installer must use Word's AppleScriptTask directory");
expectIncludes(installer, "Library/Application Scripts/com.microsoft.Powerpoint", "Installer must use PowerPoint's AppleScriptTask directory");
expectIncludes(installer, "addins.json", "Installer must require the compiled add-in checksum manifest");
expectIncludes(installer, "word/vbaProject.bin", "Installer must validate the Word VBA project entry");
expectIncludes(installer, "ppt/vbaProject.bin", "Installer must validate the PowerPoint VBA project entry");
expectIncludes(installer, 'validate_vba_markers(path, expected_vba_entry, required_vba_markers)', "Installer must reject stale add-ins that lack shared paths, event sinks, double-click handlers or the current plug-in version");
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
