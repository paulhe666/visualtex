import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

async function source(path) {
  return (await readFile(path, "utf8")).replace(/\r\n?/g, "\n");
}

const solution = await source("src-windows/VisualTeX.WindowsOffice.sln");
for (const project of [
  "VisualTeX.WindowsOffice.Contracts",
  "VisualTeX.WindowsOleBridge",
  "VisualTeX.WordVsto",
  "VisualTeX.PowerPointVsto",
  "VisualTeX.WindowsOffice.Tests",
  "VisualTeX.WindowsOffice.Installer",
  "VisualTeX.FormulaOleServer",
  "VisualTeX.FormulaOleServer.Tests",
  "VisualTeX.NativeOfficeOleAcceptance",
]) {
  assert.ok(solution.includes(project), `Solution is missing ${project}`);
}

const contracts = await source(
  "src-windows/VisualTeX.WindowsOffice.Contracts/VisualTeX.WindowsOffice.Contracts.csproj",
);
const oleProject = await source(
  "src-windows/VisualTeX.WindowsOleBridge/VisualTeX.WindowsOleBridge.csproj",
);
const wordProject = await source(
  "src-windows/VisualTeX.WordVsto/VisualTeX.WordVsto.csproj",
);
const powerpointProject = await source(
  "src-windows/VisualTeX.PowerPointVsto/VisualTeX.PowerPointVsto.csproj",
);
const nativeOfficeAcceptanceProject = await source(
  "src-windows/VisualTeX.NativeOfficeOleAcceptance/VisualTeX.NativeOfficeOleAcceptance.csproj",
);
assert.ok(contracts.includes("netstandard2.0"));
assert.ok(oleProject.includes("net8.0-windows"));
assert.ok(oleProject.includes("win-x64"));
assert.ok(oleProject.includes("PublishSingleFile>true"));
assert.ok(oleProject.includes("SelfContained>true"));
assert.ok(wordProject.includes("<TargetFramework>net472</TargetFramework>"));
assert.ok(powerpointProject.includes("<TargetFramework>net472</TargetFramework>"));
assert.ok(wordProject.includes("Microsoft.NETFramework.ReferenceAssemblies.net472"));
assert.ok(powerpointProject.includes("Microsoft.NETFramework.ReferenceAssemblies.net472"));
for (const vstoProject of [wordProject, powerpointProject]) {
  assert.ok(vstoProject.includes("<Platforms>x86;x64</Platforms>"));
  assert.ok(vstoProject.includes("'$(Platform)' == 'x86'"));
  assert.ok(vstoProject.includes(">x86</PlatformTarget>"));
  assert.ok(vstoProject.includes(">x64</PlatformTarget>"));
}
assert.ok(nativeOfficeAcceptanceProject.includes("<TargetFramework>net48</TargetFramework>"));
assert.ok(nativeOfficeAcceptanceProject.includes("<Platforms>x86;x64</Platforms>"));
assert.ok(nativeOfficeAcceptanceProject.includes("Microsoft.Office.Interop.Word"));
assert.ok(nativeOfficeAcceptanceProject.includes("Microsoft.Office.Interop.PowerPoint"));

const customSymbolDialog = await source(
  "src/components/CustomSymbolDesignerDialog.tsx",
);
const customSymbolRegistry = await source("src/math/customSymbolRegistry.ts");
const customSymbolDesignerStyles = await source(
  "src/styles-custom-symbol-designer.css",
);
assert.ok(customSymbolDialog.includes('? { enabled: false, width: 30 }'));
assert.ok(customSymbolRegistry.includes("finiteNumber(value.outline.width, 30)"));
assert.match(
  customSymbolDesignerStyles,
  /\.custom-symbol-designer-viewport-controls\s*\{[^}]*bottom:\s*auto;[^}]*width:\s*max-content;[^}]*height:\s*auto;/s,
);
assert.match(
  customSymbolDesignerStyles,
  /\.custom-symbol-geometry-icon\.is-line::before,[^}]*transform:\s*none;/s,
);
assert.match(
  customSymbolDesignerStyles,
  /\.custom-symbol-geometry-icon\.is-arrow::before\s*\{[^}]*transform:\s*none;/s,
);

const nativeOleProject = await source(
  "src-windows/VisualTeX.FormulaOleServer/VisualTeX.FormulaOleServer.vcxproj",
);
const nativeOleContract = await source(
  "src-windows/VisualTeX.FormulaOleServer/FormulaOleContract.h",
);
const nativeOleIdl = await source(
  "src-windows/VisualTeX.FormulaOleServer/FormulaOleServer.idl",
);
const nativeOleHeader = await source(
  "src-windows/VisualTeX.FormulaOleServer/FormulaOleObject.h",
);
const nativeOleSource = await source(
  "src-windows/VisualTeX.FormulaOleServer/FormulaOleObject.cpp",
);
const nativeOleRegistration = await source(
  "src-windows/VisualTeX.FormulaOleServer/FormulaOleObject.rgs",
);
assert.ok(nativeOleProject.includes("<UseOfAtl>Static</UseOfAtl>"));
assert.ok(nativeOleProject.includes("Release|Win32"));
assert.ok(nativeOleProject.includes("Release|x64"));
for (const identity of [
  "VisualTeX.Formula.1",
  "VisualTeX.Formula",
  "8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B",
  "6C672AF0-7321-4D21-B325-868CB34592C2",
  "A59B7798-6F24-4CF0-B378-E951BFFAFB3A",
  "3C72FF7F-B04A-4FD0-AA7D-61D110D8B3C1",
  "DF66EC66-3B3A-4675-A7BE-30456A04EB96",
  "VisualTeX.Formula.json",
  "VisualTeX.Preview.emf",
  "VisualTeX.Preview.png",
]) {
  assert.ok(nativeOleContract.includes(identity), `Native OLE contract is missing ${identity}`);
}
for (const requiredInterface of [
  "public IOleObject",
  "public IDataObject",
  "public IPersistStorage",
  "public IViewObject2",
  "IDispatchImpl<",
  "IVisualTeXFormulaObject",
  "IVisualTeXFormulaMetadata",
]) {
  assert.ok(
    nativeOleHeader.includes(requiredInterface),
    `Native OLE object is missing ${requiredInterface}`,
  );
}
assert.ok(nativeOleIdl.includes("oleautomation"));
assert.ok(nativeOleIdl.includes("dual"));
assert.ok(nativeOleIdl.includes("IVisualTeXFormulaObject : IDispatch"));
assert.ok(nativeOleIdl.includes("[id(1)] HRESULT InitializeFromFiles"));
assert.ok(nativeOleIdl.includes("[id(2)] HRESULT UpdateFromFiles"));
assert.ok(nativeOleIdl.includes("[id(3)] HRESULT GetFormulaJson"));
assert.ok(nativeOleIdl.includes("IVisualTeXFormulaMetadata : IDispatch"));
assert.ok(nativeOleIdl.includes("[id(1)] HRESULT SetFormulaJson"));
assert.ok(nativeOleSource.includes("CFormulaOleObject::SetFormulaJson"));
assert.ok(nativeOleIdl.includes("DF66EC66-3B3A-4675-A7BE-30456A04EB96"));
assert.ok(nativeOleProject.includes("<Midl Include=\"FormulaOleServer.idl\""));
assert.ok(nativeOleProject.includes("VisualTeX.FormulaOleServer.tlb"));
assert.ok(nativeOleSource.includes("CreateOleAdviseHolder"));
assert.ok(nativeOleSource.includes("CreateDataAdviseHolder"));
assert.ok(nativeOleSource.includes("SendOnDataChange"));
assert.ok(nativeOleSource.includes("PlayEnhMetaFile"));
assert.ok(nativeOleSource.includes("CF_METAFILEPICT"));
assert.ok(nativeOleSource.includes("GetWinMetaFileBits"));
assert.ok(nativeOleSource.includes("TYMED_MFPICT"));
assert.ok(nativeOleSource.includes("Gdiplus::Bitmap"));
assert.ok(nativeOleSource.includes("IsVectorEmf"));
assert.ok(nativeOleSource.includes("EMR_STRETCHDIBITS"));
assert.ok(nativeOleSource.includes("DrawImage / DrawImagePoints"));
assert.ok(nativeOleSource.includes("storage->CreateStream"));
assert.ok(nativeOleSource.includes("SHGetKnownFolderPath(FOLDERID_LocalAppData"));
assert.ok(nativeOleSource.includes("GetFinalPathNameByHandleW"));
assert.ok(nativeOleSource.includes("kPlaceholderMetadataJson"));
assert.ok(nativeOleSource.includes("clientSite_->SaveObject()"));
assert.ok(nativeOleSource.includes("sizeof(DWORD) * 2"));
assert.ok(!nativeOleSource.includes("AddPicture"));
assert.ok(nativeOleRegistration.includes("LocalServer32"));
assert.ok(nativeOleRegistration.includes("ServerExecutable"));
assert.ok(nativeOleRegistration.includes("InprocHandler32 = s 'Ole32.dll'"));
assert.ok(nativeOleRegistration.includes("AuxUserType"));
assert.ok(nativeOleRegistration.includes("DataFormats"));
assert.ok(nativeOleRegistration.includes("3,1,32,1"));
assert.ok(nativeOleRegistration.includes("Insertable"));
assert.ok(nativeOleRegistration.includes("VersionIndependentProgID"));
assert.ok(nativeOleRegistration.includes("ProxyStubClsid32"));
assert.ok(nativeOleRegistration.includes("00020424-0000-0000-C000-000000000046"));
assert.ok(nativeOleRegistration.includes("DF66EC66-3B3A-4675-A7BE-30456A04EB96"));

const nativeOleSmoke = await source(
  "src-windows/VisualTeX.FormulaOleServer.Tests/FormulaOleServerSmoke.cpp",
);
for (const requirement of [
  "RegServerPerUser",
  "CoCreateInstance",
  "IPersistStorage::Save",
  "IPersistStorage::Load",
  "QueryGetData(CF_ENHMETAFILE)",
  "QueryGetData(CF_METAFILEPICT)",
  "GetData(CF_METAFILEPICT)",
  "GetData(PNG)",
  "IViewObject2::Draw",
  "Raster EMF update unexpectedly succeeded",
  "Failed update mutated the formula",
  "VerifyOleCreateProtocol",
  "OLERENDER_NONE",
  "OLERENDER_DRAW",
  "VerifyPlaceholderPersistence",
  "InitializeFromFiles(after placeholder reload)",
]) {
  assert.ok(nativeOleSmoke.includes(requirement), `Native OLE smoke test is missing ${requirement}`);
}

const pipe = await source("src-windows/VisualTeX.WindowsOleBridge/NamedPipeServer.cs");
const program = await source("src-windows/VisualTeX.WindowsOleBridge/Program.cs");
const dispatcher = await source("src-windows/VisualTeX.WindowsOleBridge/OfficeStaDispatcher.cs");
const backend = await source("src-windows/VisualTeX.WindowsOleBridge/WindowsOfficeBackend.cs");
const doubleClickHook = await source("src-windows/VisualTeX.WindowsOleBridge/OfficeDoubleClickHook.cs");
const word = await source("src-windows/VisualTeX.WindowsOleBridge/WordOleService.cs");
const powerpoint = await source("src-windows/VisualTeX.WindowsOleBridge/PowerPointOleService.cs");
assert.ok(pipe.includes("PipeSecurity"));
assert.ok(pipe.includes("WindowsIdentity.GetCurrent"));
assert.ok(pipe.includes("ConstantTimeEquals"));
assert.ok(pipe.includes("MaxLineLength = 1024 * 1024"));
assert.ok(pipe.includes("RequestTimeout = TimeSpan.FromSeconds(30)"));
assert.ok(pipe.includes('"office_operation_timeout"'));
assert.ok(pipe.includes("Environment.Exit(124)"));
assert.ok(program.includes("VisualTeX.OfficeBridge.{sid}"));
assert.ok(program.includes("LocalApplicationData"));
assert.ok(program.includes('"VisualTeX",\n                "office",\n                "temp"'));
assert.ok(dispatcher.includes("SetApartmentState(ApartmentState.STA)"));
assert.ok(dispatcher.includes("Application.Run"));
assert.ok(!word.includes("Task.Run"));
assert.ok(!powerpoint.includes("Task.Run"));
assert.ok(backend.includes("_dispatcher.InvokeAsync"));
assert.ok(word.includes("InlineShapes.AddPicture"));
assert.ok(word.includes("visualtex-word-ole-range:"));
assert.ok(word.includes("EnsureSourceDocument"));
assert.ok(word.includes("WdAlignParagraphCenter"));
assert.ok(!word.includes("ContentControl"));
assert.ok(word.includes("AlternativeText"));
assert.ok(word.includes("Title"));
assert.ok(powerpoint.includes("Shapes.AddPicture"));
assert.ok(powerpoint.includes("visualtex-ppt-ole-slide:"));
assert.ok(powerpoint.includes("ResolveTargetSlide"));
assert.ok(powerpoint.includes("EnsureSourceDocument"));
assert.ok(powerpoint.includes('VisualTeX_{formulaId}'));
assert.ok(powerpoint.includes("AlternativeText"));
assert.ok(powerpoint.includes("Tags"));
assert.ok(powerpoint.includes("CalculateReplacementSize"));
assert.ok(powerpoint.includes("originalMetadata?.RenderHeightPx"));
assert.ok(!powerpoint.includes("FitImage(session.ImagePath, width, height)"));
assert.ok(powerpoint.indexOf("ConfigureShape(newShape") < powerpoint.indexOf("original.Delete()"));
assert.ok(word.indexOf("ConfigureShape(\n                    candidate") < word.indexOf("original.Delete()"));

for (const method of [
  "health",
  "office.detect",
  "powerpoint.getSelection",
  "powerpoint.insertFormula",
  "powerpoint.replaceFormula",
  "powerpoint.markFormula",
  "powerpoint.deleteFormula",
  "word.getSelection",
  "word.insertInlineFormula",
  "word.insertDisplayFormula",
  "word.replaceFormula",
  "word.updateEquationNumbers",
  "office.openWord",
  "office.openPowerPoint",
  "shutdown",
]) {
  assert.ok(backend.includes(`\"${method}\"`), `OLE backend is missing ${method}`);
}

const sessionClient = await source(
  "src-windows/VisualTeX.WindowsOffice.Contracts/VisualTeXSessionClient.cs",
);
const wordVsto = await source("src-windows/VisualTeX.WordVsto/ThisAddIn.cs");
const wordVstoService = await source(
  "src-windows/VisualTeX.WordVsto/WordFormulaService.cs",
);
const officeFormulaSizing = await source(
  "src-windows/VisualTeX.WindowsOffice.Contracts/OfficeFormulaSizing.cs",
);
const wordInlineAlignment = await source(
  "src-windows/VisualTeX.WindowsOffice.Contracts/WordInlineAlignment.cs",
);
const powerpointVsto = await source(
  "src-windows/VisualTeX.PowerPointVsto/ThisAddIn.cs",
);
const powerpointVstoService = await source(
  "src-windows/VisualTeX.PowerPointVsto/PowerPointFormulaService.cs",
);
const vstoDependencyResolver = await source(
  "src-windows/VisualTeX.WindowsOffice.VstoShared/VstoDependencyResolver.cs",
);
const vstoOlePreview = await source(
  "src-windows/VisualTeX.WindowsOffice.VstoShared/OfficeOlePreview.cs",
);
const ribbonIconData = await source(
  "src-windows/VisualTeX.WindowsOffice.VstoShared/RibbonIconData.cs",
);
const ribbonIconProvider = await source(
  "src-windows/VisualTeX.WindowsOffice.VstoShared/RibbonIconProvider.cs",
);
const ribbonVectorIconRenderer = await source(
  "src-windows/VisualTeX.WindowsOffice.VstoShared/RibbonVectorIconRenderer.cs",
);
const wordDoubleClickHook = await source(
  "src-windows/VisualTeX.WordVsto/WordDoubleClickHook.cs",
);
const vstoOlePngExtractor = await source(
  "src-windows/VisualTeX.WindowsOffice.VstoShared/OlePngPreviewExtractor.cs",
);
const wordEquationNumbering = await source(
  "src-windows/VisualTeX.WindowsOffice.VstoShared/WordEquationNumbering.cs",
);
const wordEquationNumberingPerformance = await source(
  "src-windows/VisualTeX.WindowsOffice.VstoShared/WordEquationNumbering.Performance.cs",
);
const wordEquationNumberingTrueDisplay = await source(
  "src-windows/VisualTeX.WindowsOffice.VstoShared/WordEquationNumbering.TrueDisplay.cs",
);
const wordEquationReferenceFields = await source(
  "src-windows/VisualTeX.WordVsto/WordEquationReferenceFields.cs",
);
const wordOmmlConverter = await source(
  "src-windows/VisualTeX.WindowsOffice.VstoShared/WordOmmlConverter.cs",
);
const wordOmmlFormulaStore = await source(
  "src-windows/VisualTeX.WindowsOffice.VstoShared/WordOmmlFormulaStore.cs",
);
const formulaOleInterop = await source(
  "src-windows/VisualTeX.WindowsOffice.Contracts/FormulaOleInterop.cs",
);
const wordFormulaMetadataReader = await source(
  "src-windows/VisualTeX.WindowsOffice.VstoShared/WordFormulaMetadataReader.cs",
);
const wordOleObjectAccessor = await source(
  "src-windows/VisualTeX.WindowsOffice.VstoShared/WordOleObjectAccessor.cs",
);
const nativeOfficeAcceptance = await source(
  "src-windows/VisualTeX.NativeOfficeOleAcceptance/Program.cs",
);
const nativeOfficeAcceptanceScript = await source(
  "scripts/test_windows_native_office_ole.ps1",
);
const vstoFlowAcceptanceProject = await source(
  "src-windows/VisualTeX.VstoFlowAcceptance/VisualTeX.VstoFlowAcceptance.csproj",
);
const vstoFlowAcceptance = await source(
  "src-windows/VisualTeX.VstoFlowAcceptance/Program.cs",
);
const vstoBulkSpacingAcceptance = await source(
  "src-windows/VisualTeX.VstoFlowAcceptance/WordBulkImportLatexSpacingAcceptance.cs",
);
const vstoDisplaySpacingAcceptance = await source(
  "src-windows/VisualTeX.VstoFlowAcceptance/WordDisplaySpacingAcceptance.cs",
);
const vstoEquationNumberFormatAcceptance = await source(
  "src-windows/VisualTeX.VstoFlowAcceptance/WordEquationNumberFormatAcceptance.cs",
);
const vstoFormulaToLatexAcceptance = await source(
  "src-windows/VisualTeX.VstoFlowAcceptance/WordFormulaToLatexAcceptance.cs",
);
const officeServer = await source("src-tauri/src/office/server.rs");
const officeState = await source("src-tauri/src/office/state.rs");
const officeSessions = await source("src-tauri/src/office/sessions.rs");
const officeDialogMain = await source("src/office/dialog/main.tsx");
const officeDialogMessages = await source("src/office/dialog/dialogMessages.ts");
const documentImportApp = await source(
  "src/office/documentImport/DocumentImportApp.tsx",
);
const documentImportParser = await source(
  "src/office/documentImport/documentImportParser.ts",
);
const documentImportFile = await source(
  "src/office/documentImport/documentImportFile.ts",
);
const documentImportCss = await source(
  "src/office/documentImport/documentImport.css",
);
assert.ok(!sessionClient.includes("_installToken = ReadInstallToken()"));
assert.ok(sessionClient.includes("StartVisualTeXCompanion"));
assert.ok(sessionClient.includes("timeout.CancelAfter(TimeSpan.FromSeconds(3))"));
assert.ok(sessionClient.includes("handler.SslProtocols = SslProtocols.Tls12"));
assert.ok(sessionClient.includes("pinnedCertificateMatches && certificateTimeValid"));
assert.ok(sessionClient.includes("_hasValidatedServerCertificate"));
assert.ok(!sessionClient.includes("_lastServerCertificateThumbprint = string.Empty;\r\n        _lastTlsPolicyErrors = string.Empty;"));
assert.ok(sessionClient.includes('Path.Combine(CompanionLogRoot(), "vsto-client.log")'));
assert.ok(sessionClient.includes("assembly.Location"));
assert.ok(sessionClient.includes("tls-certificate-callback"));
assert.ok(!sessionClient.includes("return sslPolicyErrors == SslPolicyErrors.None"));
assert.ok(sessionClient.includes("OpenEditorAsync"));
assert.ok(sessionClient.includes("OpenConverterBatchAsync"));
assert.ok(sessionClient.includes('"/api/v1/app/converter/convert-batch"'));
assert.ok(sessionClient.includes("OpenBulkImportAsync"));
assert.ok(sessionClient.includes("CloseEditorAsync"));
assert.ok(sessionClient.includes("/api/v1/app/sessions/"));
assert.ok(sessionClient.includes('}/close"'));
assert.ok(!sessionClient.includes('new Uri(CompanionOrigin, $"/dialog/'));
assert.ok(wordVsto.includes("await client.OpenEditorAsync"));
assert.ok(powerpointVsto.includes("await client.OpenEditorAsync"));
assert.ok(officeServer.includes("open_desktop_session_window"));
assert.ok(officeServer.includes("bring_session_window_to_front"));
assert.ok(officeServer.includes("set_always_on_top(true)"));
assert.ok(officeServer.includes("is_focused().unwrap_or(false)"));
assert.ok(officeServer.includes("Duration::from_millis(220)"));
assert.ok(!officeServer.includes("UserAttentionType::Informational"));
assert.ok(!officeServer.includes("for delay_ms in [140_u64, 260, 520, 900]"));
assert.ok(officeServer.includes("OFFICE_EDITOR_WINDOW_TRANSITION_LOCK"));
assert.ok(officeServer.includes("OFFICE_EDITOR_PAGE_READY"));
assert.ok(officeServer.includes("switch_office_editor_session"));
assert.ok(officeServer.includes("stale close cleanup skipped"));
assert.ok(officeServer.includes("reveal fallback deferred because page is not ready"));
assert.ok(officeServer.includes("active Office editor WebView is still loading"));
assert.ok(officeServer.includes("WebviewWindowBuilder::new"));
assert.ok(officeServer.includes('"/app/sessions/{session_id}/open"'));
assert.ok(officeServer.includes('"/app/sessions/{session_id}/bulk-import"'));
assert.ok(officeServer.includes('"/app/converter/convert-batch"'));
assert.ok(officeServer.includes('"/app/converter/next-batch"'));
assert.ok(officeServer.includes("queue_desktop_batch_conversion"));
assert.ok(officeServer.includes("OFFICE_BATCH_CONVERSION_QUEUE"));
assert.ok(officeServer.includes("visualtex-office-batch-conversion"));
assert.ok(officeServer.includes("?runtime=vsto-bulk-import"));
assert.ok(officeServer.includes("open_desktop_bulk_import_window"));
assert.ok(officeServer.includes('format!("office-import-{suffix}")'));
assert.ok(officeServer.includes('"/app/sessions/{session_id}/close"'));
assert.ok(officeServer.includes("close_desktop_session"));
assert.ok(officeServer.includes("WebviewUrl::External(url)"));
assert.ok(officeServer.includes("?runtime=vsto-desktop"));
assert.ok(officeServer.includes('visualtex-editor-layout'));
assert.ok(officeServer.includes("current_editor_layout"));
assert.ok(!officeServer.includes("remove_office_js"));
assert.ok(!officeServer.includes('"<script src=\\\"/vendor/office-js/office.js\\\"></script>"'));
assert.ok(officeServer.includes('#[cfg(not(target_os = "windows"))]'));
assert.ok(officeDialogMain.includes("executes the Office.js runtime"));
assert.ok(officeDialogMain.includes("DocumentImportApp"));
assert.ok(officeDialogMain.includes('runtime === "vsto-bulk-import"'));
assert.ok(officeDialogMain.includes('meta[name="visualtex-editor-layout"]'));
assert.ok(!officeDialogMain.includes("Office.onReady"));
assert.ok(officeDialogMessages.includes("no Office.context.ui parent"));
assert.ok(officeDialogMessages.includes("return false"));
assert.ok(!officeDialogMessages.includes("messageParent"));
const officeDialogApp = await source("src/office/dialog/OfficeDialogApp.tsx");
const officeApplyShortcut = await source("src/office/dialog/officeApplyShortcut.ts");
const editorWorkspace = await source("src/workspace/EditorWorkspace.tsx");
const settingsDialog = await source("src/components/SettingsDialog.tsx");
const editorStore = await source("src/stores/editorStore.ts");
const mathEditor = await source("src/editor/MathEditor.tsx");
const formulaToolbar = await source("src/toolbar/FormulaToolbar.tsx");
assert.ok(officeDialogApp.includes("closeOfficeSessionWindow"));
assert.ok(officeDialogApp.includes("IS_VSTO_DESKTOP_RUNTIME"));
assert.ok(officeDialogApp.includes("if (!delivered) window.close()"));
assert.ok(officeDialogApp.includes("const unchangedEdit = session?.mode === \"edit\" && !dirty"));
assert.ok(officeDialogApp.includes("originalFingerprintRef"));
assert.ok(officeDialogApp.includes("originalFingerprintRef.current = loadedFingerprint"));
assert.ok(officeDialogApp.includes("normalizeOfficeCodeFormat"));
assert.ok(officeDialogApp.includes('return "raw"'));
assert.ok(officeDialogApp.includes("setLatexCodeFormat(loadedCodeFormat)"));
assert.ok(officeDialogApp.includes("getOfficePreferences"));
assert.ok(officeDialogApp.includes("applyEditorLayout"));
assert.ok(officeDialogApp.includes("status.editorLayout"));
assert.ok(officeDialogApp.includes("data-office-font-size"));
assert.ok(officeDialogApp.includes("<select"));
assert.ok(officeDialogApp.includes("OFFICE_CHINESE_FONT_SIZE_OPTIONS"));
assert.ok(officeDialogApp.includes('{ name: "小四", fontSizePt: 12 }'));
assert.ok(officeDialogApp.includes('{ name: "五号", fontSizePt: 10.5 }'));
assert.ok(officeDialogApp.includes("officePointFontSizeOptions(officeFontSizePt)"));
assert.ok(officeDialogApp.includes('label={isEn ? "Chinese sizes" : "中文字号"}'));
assert.ok(!officeDialogApp.includes("office-formula-font-size-presets"));
assert.ok(officeDialogApp.includes("fontSizePt: officeFontSizePt"));
assert.ok(officeDialogApp.includes("powerPointDefaultFontSizePt"));
assert.ok(officeDialogApp.includes("registerOfficeApplyShortcut"));
assert.ok(officeDialogApp.includes("commitFromShortcutRef"));
assert.ok(officeDialogApp.includes("IS_VSTO_CONVERT_RUNTIME"));
assert.ok(officeDialogApp.includes("generateSessionExportResult"));
assert.ok(officeDialogApp.includes('sourceSession.objectMode === "wordOmml"'));
assert.ok(officeDialogApp.includes("visualtex-office-batch-conversion"));
assert.ok(officeDialogApp.includes("takeOfficeConverterBatch"));
assert.ok(officeDialogApp.includes("drainCompanionQueue"));
assert.ok(officeDialogApp.includes("batchConversionQueueRef"));
assert.ok(officeDialogApp.includes("!ocrOpen"));
assert.ok(officeDialogApp.includes("!inlineOcrBusyRef.current"));
assert.ok(officeApplyShortcut.includes('window.addEventListener("keydown", handleOfficeApplyShortcut, true)'));
assert.ok(officeApplyShortcut.includes("event.preventDefault()"));
assert.ok(officeApplyShortcut.includes("event.stopImmediatePropagation()"));
assert.ok(officeApplyShortcut.includes("if (event.repeat) return"));
assert.ok(officeApplyShortcut.includes("event.ctrlKey"));
assert.ok(officeApplyShortcut.includes("!event.altKey"));
assert.ok(officeApplyShortcut.includes("!event.shiftKey"));
assert.ok(officeApplyShortcut.includes("!event.metaKey"));
assert.ok(officeApplyShortcut.includes("!event.isComposing"));
assert.ok(officeDialogApp.includes('aria-keyshortcuts="Control+S"'));
assert.ok(officeDialogApp.includes("应用并关闭（Ctrl+S）"));
assert.ok(officeServer.includes('"/preferences"'));
assert.ok(officeServer.includes("powerpoint_default_font_size_pt"));
assert.ok(officeState.includes('office-preferences.json'));
assert.ok(officeState.includes("DEFAULT_POWERPOINT_FORMULA_FONT_SIZE_PT: f64 = 20.0"));
assert.ok(officeState.includes("set_powerpoint_default_font_size_pt"));
assert.ok(officeState.includes("persist_office_preferences"));
assert.ok(officeState.includes("app_editor_layout"));
assert.ok(officeState.includes("normalize_app_editor_layout"));
assert.ok(settingsDialog.includes("data-powerpoint-default-font-size"));
assert.ok(editorStore.includes("powerPointDefaultFontSizePt: 20"));
assert.ok(editorStore.includes("contextCounts"));
assert.ok(mathEditor.includes("rankNativeSuggestionItems"));
assert.ok(mathEditor.includes("recordNativeSuggestionUsage"));
assert.ok(formulaToolbar.includes("commonToolbarCommandLimit = 45"));
assert.ok(formulaToolbar.includes("visualtex-common-toolbar-command-ids-v2"));
assert.ok(formulaToolbar.includes("visualtex-common-toolbar-command-ids-v1"));
assert.ok(formulaToolbar.includes('if (id === "notin") return "times"'));
assert.ok(formulaToolbar.includes('if (id === "leftarrow") return "div"'));
assert.ok(formulaToolbar.includes("addCommandToCommon"));
assert.ok(formulaToolbar.includes("设为常用"));
assert.ok(!formulaToolbar.includes("contextCounts.toolbar"));
assert.ok(!editorWorkspace.includes("data-formula-typing-bold"));
assert.ok(!editorWorkspace.includes("data-formula-typing-italic"));
assert.ok(editorWorkspace.includes("data-formula-selection-bold"));
assert.ok(editorWorkspace.includes("data-formula-selection-italic"));
assert.ok(editorWorkspace.includes("data-formula-selection-color"));
assert.ok(editorWorkspace.includes("data-formula-selection-background"));
assert.ok(!editorWorkspace.includes("typingStyle={typingStyle}"));
assert.ok(!mathEditor.includes("applyTypingStyleAtCollapsedSelection"));
assert.ok(!mathEditor.includes("MathEditorTypingStyle"));
assert.ok(mathEditor.includes("captureSelectionTarget"));
assert.ok(mathEditor.includes("applySelectionStyle"));
assert.ok(mathEditor.includes("MathLive owns pointer capture"));
assert.ok(mathEditor.includes(":host(.has-visualtex-multi-line-selection)"));
assert.ok(mathEditor.includes("multiLineSelectionRef.current && event.buttons === 0"));
assert.ok(documentImportApp.includes("Word 结构预览"));
assert.ok(documentImportApp.includes("doc-import-preview-stage"));
assert.ok(documentImportApp.includes("doc-import-preview-counts"));
assert.ok(documentImportApp.includes("applyDocumentTheme"));
assert.ok(documentImportApp.includes("latexToSvg"));
assert.ok(documentImportApp.includes("fontSizePt: 13"));
assert.ok(documentImportApp.includes("parseDocumentImport"));
assert.ok(documentImportApp.includes("readDocumentImportFile"));
assert.ok(documentImportApp.includes("导入 .tex / .md"));
assert.ok(documentImportApp.includes('accept=".tex,.md,.markdown,text/x-tex,text/markdown"'));
assert.ok(documentImportApp.includes("importedFile.encoding"));
assert.ok(documentImportApp.includes("已编辑"));
assert.ok(documentImportApp.includes("JSON.stringify(preview.parsed)"));
assert.ok(documentImportApp.includes('codeFormat: "visualtex-document-json"'));
assert.ok(documentImportCss.includes("font-size: 13pt;"));
assert.ok(documentImportCss.includes(".doc-import-formula.display"));
assert.ok(documentImportCss.includes("width: 100%;\n  min-width: 0;"));
assert.ok(documentImportParser.includes("findDisplayStart"));
assert.ok(documentImportParser.includes('findUnescaped(text, "\\\\["'));
assert.ok(documentImportParser.includes("normalizeDisplayEnvironment"));
assert.ok(documentImportParser.includes("normalizeMarkdownSource"));
assert.ok(documentImportParser.includes("normalizeLatexExtensions"));
assert.ok(documentImportParser.includes("normalizeLatexInlineBoundaryWhitespace"));
assert.ok(documentImportParser.includes("isTightLatexBoundaryCharacter"));
assert.ok(documentImportParser.includes("visibleMultiArgumentCommands"));
assert.ok(documentImportFile.includes("readDocumentImportFile"));
assert.ok(documentImportFile.includes("documentFormatFromFileName"));
assert.ok(documentImportFile.includes('new TextDecoder("gb18030"'));
assert.ok(documentImportFile.includes('new TextDecoder("utf-16le"'));
assert.ok(documentImportFile.includes("DOCUMENT_IMPORT_MAX_FILE_BYTES"));
assert.ok(documentImportCss.includes(".doc-import-file-button"));
assert.ok(documentImportCss.includes(".doc-import-file-chip"));
assert.ok(wordVsto.includes("ResolveBulkImportDocumentAsync"));
assert.ok(wordVsto.includes("ParseSerialized"));
assert.ok(wordVsto.includes('"visualtex-document-json"'));
assert.ok(wordVsto.includes("OpenBulkImportAsync"));
assert.ok(wordVsto.includes("PrewarmConverterAsync"));
assert.ok(wordVsto.includes("OpenConverterBatchAsync"));
assert.ok(wordVsto.includes("render-batch-complete"));
assert.ok(wordVstoService.includes("screenUpdatingSuspended"));
assert.ok(wordVstoService.includes("InsertBulkOleDocumentTwoPhase"));
assert.ok(wordVstoService.includes("InsertNativeTextRun(document, selection, run)"));
assert.ok(wordVstoService.includes("ResetSelectionTransientFormatting(selection)"));
assert.ok(wordVstoService.includes("trustExportDimensions: bulkImport"));
assert.ok(wordVstoService.includes("NormalizeInlineBaselineBoundary"));
assert.ok(wordVstoService.includes('private const string InlineOleTypingAnchor = "\\u200C";'));
assert.ok(wordVstoService.includes("sentinel.InsertBefore(InlineOleTypingAnchor)"));
assert.ok(wordVstoService.includes("ConfigureInlineOleTypingAnchor"));
assert.ok(wordVstoService.includes("font.Subscript = 0"));
assert.ok(wordVstoService.includes("font.Superscript = 0"));
assert.ok(wordVstoService.includes("RemoveInlineBaselineSentinel"));
assert.ok(wordVstoService.includes("ResolveDisplayInsertionRange"));
assert.ok(
  wordEquationNumberingTrueDisplay.includes("PrepareNumberedNativeOmmlInsertionHost"),
);
assert.ok(wordVstoService.includes("var usePreservedDisplayParagraph ="));
assert.ok(!wordVstoService.includes("var usePreservedOleDisplayParagraph ="));
assert.ok(wordVstoService.includes("preserveNativeOmmlSpacing: nativeOmml"));
assert.ok(wordVstoService.includes("NativeOmmlScreenUpdatingScope.Suspend(_application)"));
assert.ok(wordEquationNumberingPerformance.includes("ScreenUpdatingScope.Suspend(document)"));
assert.ok(
  wordEquationNumbering.includes(
    "TryRefreshHealthyEquationNumbersInPlace(document, out var updated)",
  ),
);
for (const productionNumberingSource of [
  wordVstoService,
  wordEquationNumbering,
  wordEquationNumberingTrueDisplay,
]) {
  assert.ok(!productionNumberingSource.includes("document.Tables.Add(tableAnchor, 1, 3)"));
  assert.ok(!productionNumberingSource.includes(".ConvertToTable("));
  assert.ok(!productionNumberingSource.includes("NormalizeNumberedDisplayArgumentSizing("));
  assert.ok(!productionNumberingSource.includes("WrapVisualTeXNativeEquationNumber("));
}
assert.ok(wordEquationNumberingTrueDisplay.includes("WdOMathType.wdOMathDisplay"));
assert.ok(wordEquationNumberingTrueDisplay.includes("EnsureNativeDisplayNumberShape"));
assert.ok(wordEquationNumberingTrueDisplay.includes("WdFieldType.wdFieldRef"));
assert.ok(wordEquationNumberingTrueDisplay.includes("IsPureTrueDisplayFormulaParagraph"));
assert.ok(wordVstoService.includes('private const string InlineMathGuard = " ";'));
assert.ok(wordVstoService.includes('private const string InlineBaselineSentinel = " ";'));
assert.ok(wordVstoService.includes("LegacyInlineMathGuard"));
assert.ok(wordVstoService.includes("LegacyInlineBaselineSentinel"));
assert.ok(wordVstoService.includes("font.Hidden = -1"));
assert.ok(wordVstoService.includes("RestoreOmmlReplacementRollback"));
const directTableReplacementStart = wordVstoService.indexOf(
  "if (replaceHealthyDirectTableAtomically)\n            {\n                if (reuseHealthyDirectTableForUnnumberingOnly)",
);
const directTableReplacementEnd = wordVstoService.indexOf(
  "else if (replaceHealthyStandaloneDisplayAtomically)",
  directTableReplacementStart,
);
assert.ok(directTableReplacementStart >= 0 && directTableReplacementEnd > directTableReplacementStart);
const directTableReplacement = wordVstoService.slice(
  directTableReplacementStart,
  directTableReplacementEnd,
);
assert.ok(directTableReplacement.includes("WordOmmlConverter.Insert("));
assert.ok(directTableReplacement.includes("replaceTarget: true"));
assert.ok(!directTableReplacement.includes("ReplaceWithPreparedOmml("));
assert.ok(officeSessions.includes("unchanged_edit"));
assert.ok(officeSessions.includes("document_import"));
assert.ok(officeSessions.includes('"visualtex-document-json"'));
assert.ok(officeSessions.includes("document_import_can_commit_source_without_formula_export"));
assert.ok(officeSessions.includes("unchanged_edit_can_complete_without_new_export_result"));
assert.ok(officeSessions.includes("changed_edit_still_requires_a_new_export_result"));
assert.ok(wordVstoService.includes("shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse"));
assert.ok(wordVstoService.includes("shape.Height = height"));
assert.ok(wordVstoService.includes("ApplyInlineBaseline("));
assert.ok(wordVstoService.includes("ReadDefinedShapeFontPosition"));
assert.ok(wordVstoService.includes("CalculateFontPositionWithLegacyFallback"));
assert.ok(wordInlineAlignment.includes("HasValidExportedBaseline"));
assert.ok(wordInlineAlignment.includes("dequantizedMagnitude"));
assert.ok(wordInlineAlignment.includes("LegacyDescentRatio"));
assert.ok(wordVstoService.includes("RestoreTypingBaselineAfter(shape)"));
assert.ok(wordVstoService.includes("font.StrikeThrough = run.Strike"));
assert.ok(wordVstoService.includes("font.Underline = run.Underline"));
assert.ok(wordVstoService.includes("TryResolveWordFontSize"));
assert.ok(wordVstoService.includes("selectionRange.Start - 1"));
assert.ok(wordVstoService.includes("selectionRange.Start + 1"));
assert.ok(wordVstoService.includes("font.Position = 0"));
assert.ok(wordVstoService.includes("OfficeFormulaSizing.EditedSize"));
assert.ok(officeFormulaSizing.includes("Formula height is the visual font-size reference"));
assert.ok(powerpointVstoService.includes("ApplyOleSizeAndRefresh"));
assert.ok(powerpointVstoService.includes("RestoreOlePosition"));
assert.ok(powerpointVstoService.includes("OfficeFormulaSizing.EditedSize"));
assert.ok(powerpointVstoService.includes("return 20f"));
assert.ok(!powerpointVstoService.includes("format.DoVerb(-1)"));
assert.ok(powerpointVstoService.includes("single geometry authority"));
assert.ok(!powerpointVstoService.includes("SetOleServerExtent"));
assert.ok(powerpointVstoService.includes("shape.Type is not MsoShapeType.msoEmbeddedOLEObject"));
assert.ok(powerpointVstoService.includes("var overlay = ReadPictureMetadata(shape)"));
assert.ok(powerpointVstoService.includes("if (overlay is not null) return overlay"));
assert.ok(powerpointVstoService.includes("TryApplyRotation"));
assert.ok(wordVstoService.includes("Marshal.ReleaseComObject(value)"));
assert.ok(powerpointVstoService.includes("Marshal.ReleaseComObject(value)"));
assert.ok(!wordVstoService.includes("Marshal.FinalReleaseComObject(value)"));
assert.ok(!powerpointVstoService.includes("Marshal.FinalReleaseComObject(value)"));
assert.ok(vstoFlowAcceptanceProject.includes("VisualTeX.WordVsto"));
assert.ok(vstoFlowAcceptanceProject.includes("VisualTeX.PowerPointVsto"));
assert.ok(vstoFlowAcceptance.includes("Word formula aspect ratio is distorted"));
assert.ok(vstoFlowAcceptance.includes("Word inline formula baseline is incorrect"));
assert.ok(vstoFlowAcceptance.includes("PowerPoint formula must be an editable picture/graphic"));
assert.ok(vstoFlowAcceptance.includes("Word caret inherited the formula baseline offset"));
assert.ok(vstoFlowAcceptance.includes("Word did not suppress built-in OLE activation"));
assert.ok(vstoFlowAcceptance.includes("PowerPoint edit button did not create an edit Session"));
assert.ok(vstoFlowAcceptance.includes("PowerPoint convert command did not request native OLE"));
assert.ok(vstoFlowAcceptance.includes("Word unchanged edit did not complete after closing the window"));
assert.ok(vstoFlowAcceptance.includes("PowerPoint unchanged edit did not complete after closing the window"));
assert.ok(vstoFlowAcceptance.includes("PowerPoint OLE export still resembles the placeholder cache"));
assert.ok(vstoFlowAcceptance.includes("Local\\VisualTeX.VstoFlowAcceptance"));
assert.ok(vstoFlowAcceptance.includes("acceptance.log"));
assert.ok(vstoFlowAcceptance.includes('"word-create"'));
assert.ok(vstoFlowAcceptance.includes('"word-bulk-import-latex-spacing"'));
assert.ok(vstoFlowAcceptance.includes('"word-display-spacing"'));
assert.ok(vstoFlowAcceptance.includes('"word-equation-number-format"'));
assert.ok(vstoFlowAcceptance.includes('"word-formula-to-latex"'));
assert.ok(vstoEquationNumberFormatAcceptance.includes("Heading1DotId"));
assert.ok(vstoEquationNumberFormatAcceptance.includes("Heading1DashId"));
assert.ok(vstoEquationNumberFormatAcceptance.includes("Heading2DotId"));
assert.ok(vstoEquationNumberFormatAcceptance.includes("Heading2DashId"));
assert.ok(vstoEquationNumberFormatAcceptance.includes("InsertPlainNumberingReference"));
assert.ok(vstoEquationNumberFormatAcceptance.includes("survived save/reopen"));
assert.ok(vstoFormulaToLatexAcceptance.includes("OnRedrawSelectionOleToLatex"));
assert.ok(vstoFormulaToLatexAcceptance.includes("OnRedrawSelectionOmmlToLatex"));
assert.ok(vstoFormulaToLatexAcceptance.includes("OnRedrawDocumentOleToLatex"));
assert.ok(vstoFormulaToLatexAcceptance.includes("OnRedrawDocumentOmmlToLatex"));
assert.ok(vstoFormulaToLatexAcceptance.includes("numbered tables were flattened"));
assert.ok(vstoFormulaToLatexAcceptance.includes("persisted after save/reopen"));
assert.ok(vstoBulkSpacingAcceptance.includes("ordinary CJK boundary spaces"));
assert.ok(vstoBulkSpacingAcceptance.includes("no VTBL bookmark or boundary character remained"));
assert.ok(vstoDisplaySpacingAcceptance.includes("typed prose stayed outside OMath"));
assert.ok(vstoDisplaySpacingAcceptance.includes("no inserted blank paragraph"));
assert.ok(vstoDisplaySpacingAcceptance.includes("markerParagraphRange.End"));
assert.ok(vstoBulkSpacingAcceptance.includes("rawXml.IndexOf('\\u200B') < 0"));
assert.ok(vstoBulkSpacingAcceptance.includes("rawXml.IndexOf('\\u2060') < 0"));
assert.ok(vstoBulkSpacingAcceptance.includes("rawXml.IndexOf('\\uE000') < 0"));
assert.ok(vstoFlowAcceptanceProject.includes("System.IO.Compression"));
assert.ok(vstoFlowAcceptance.includes('oleFormat.DoVerb(0)'));
assert.ok(vstoFlowAcceptance.includes("CloseEditorAsync"));
assert.ok(wordVsto.includes("活动 Word 文档已切换，未写入公式"));
assert.ok(powerpointVsto.includes("活动演示文稿已切换，未写入公式"));
assert.ok(wordVsto.includes("dispatcher.InvokeAsync"));
assert.ok(powerpointVsto.includes("dispatcher.InvokeAsync"));
assert.ok(wordVsto.includes("IDTExtensibility2, Office.IRibbonExtensibility"));
assert.ok(powerpointVsto.includes("IDTExtensibility2, Office.IRibbonExtensibility"));
assert.ok(wordVsto.includes("ClassInterfaceType.None"));
assert.ok(powerpointVsto.includes("ClassInterfaceType.None"));
assert.ok(wordVsto.includes("InterfaceIsIDispatch"));
assert.ok(powerpointVsto.includes("InterfaceIsIDispatch"));
assert.ok(wordVsto.includes("IWordRibbonCallbacks"));
assert.ok(powerpointVsto.includes("IPowerPointRibbonCallbacks"));
assert.ok(wordVsto.includes("ComDefaultInterface(typeof(IWordRibbonCallbacks))"));
assert.ok(powerpointVsto.includes("ComDefaultInterface(typeof(IPowerPointRibbonCallbacks))"));
assert.ok(wordVsto.includes('id="VisualTeX.WordVsto.Tab" label="VisualTeX"'));
assert.ok(powerpointVsto.includes('id="VisualTeX.PowerPointVsto.Tab" label="VisualTeX"'));
assert.ok(wordVsto.includes('insertAfterMso="TabHome"'));
assert.ok(powerpointVsto.includes('insertAfterMso="TabHome"'));
assert.ok(!wordVsto.includes('<tab idMso="TabHome">'));
assert.ok(!powerpointVsto.includes('<tab idMso="TabHome">'));
assert.ok(wordVsto.includes("static ThisAddIn() => VstoDependencyResolver.Install()"));
assert.ok(powerpointVsto.includes("static ThisAddIn() => VstoDependencyResolver.Install()"));
assert.ok(vstoDependencyResolver.includes("AssemblyResolve"));
assert.ok(vstoDependencyResolver.includes("MatchesIdentityIgnoringVersion"));
assert.ok(vstoDependencyResolver.includes('"System.Text.Json"'));
assert.ok(vstoDependencyResolver.includes('"System.Numerics.Vectors"'));
assert.ok(vstoDependencyResolver.includes("Assembly.LoadFrom(candidatePath)"));
assert.ok(!wordVsto.includes("IsSelectedNativeOle() == true"));
assert.ok(wordVsto.includes("cancel = true"));
assert.ok(wordVsto.includes("new WordDoubleClickHook"));
assert.ok(wordVsto.includes("ShouldInterceptNativeOleDoubleClick"));
assert.ok(wordVsto.includes("ClearNativeOleTarget"));
assert.ok(wordVsto.includes("DocumentBeforeSave += OnDocumentBeforeSave"));
assert.ok(
  wordVsto.includes("NormalizeInlineOleParagraphBaselinesBeforeSave"),
);
assert.ok(wordVsto.includes("DocumentBeforeSave -= OnDocumentBeforeSave"));
assert.ok(wordDoubleClickHook.includes("WhMouseLl = 14"));
assert.ok(wordDoubleClickHook.includes("return new IntPtr(1)"));
assert.ok(wordDoubleClickHook.includes('"WINWORD"'));
for (const command of [
  "OnConvertSelected",
  "OnDeleteSelected",
  "OnOpenDesktop",
]) {
  assert.ok(wordVsto.includes(command), `Word Ribbon is missing ${command}`);
  assert.ok(powerpointVsto.includes(command), `PowerPoint Ribbon is missing ${command}`);
}
assert.ok(!wordVsto.includes("OnExportSelectedAsPicture"));
assert.ok(powerpointVsto.includes("OnExportSelectedAsPicture"));
assert.ok(wordVsto.includes("OnUpdateEquationNumbers"));
assert.ok(!wordVsto.includes("OnBatchEquationNumbering"));
assert.ok(!wordVsto.includes("BatchEquationNumberingAsync"));
assert.ok(!wordVsto.includes('id="VisualTeX.WordVsto.BatchNumbers"'));
assert.ok(!wordVsto.includes("VISUALTEX_VSTO_BATCH_NUMBER_REDRAW"));
assert.ok(!wordVsto.includes("VISUALTEX_VSTO_BATCH_NUMBER_FORMAT"));
assert.ok(wordVsto.includes("GetEquationNumberFormatPressed"));
assert.ok(wordVsto.includes("OnEquationNumberFormatChanged"));
assert.ok(wordVsto.includes('id="VisualTeX.WordVsto.NumberFormat"'));
for (const formatTag of [
  "continuous",
  "heading1-dot",
  "heading1-dash",
  "heading2-dot",
  "heading2-dash",
]) {
  assert.ok(wordVsto.includes(`tag="${formatTag}"`));
}
assert.ok(wordVsto.includes("OnInsertEquationReference"));
assert.ok(wordVsto.includes("InsertEquationReferenceAsync"));
assert.ok(wordVsto.includes("FormulaOleContract.NativeOleMode"));
assert.ok(wordVsto.includes("FormulaOleContract.WordOmmlMode"));
assert.ok(wordVsto.includes("OnInsertInlineOmml"));
assert.ok(wordVsto.includes("OnInsertDisplayOmml"));
assert.ok(wordVsto.includes("OnConvertSelectedToOmml"));
for (const callback of [
  "OnRedrawSelectionOleToLatex",
  "OnRedrawSelectionOmmlToLatex",
  "OnRedrawDocumentOleToLatex",
  "OnRedrawDocumentOmmlToLatex",
]) {
  assert.ok(wordVsto.includes(callback), `Word Ribbon is missing ${callback}`);
}
for (const ribbonId of [
  "VisualTeX.WordVsto.RedrawSelectionOleToLatex",
  "VisualTeX.WordVsto.RedrawSelectionOmmlToLatex",
  "VisualTeX.WordVsto.RedrawDocumentOleToLatex",
  "VisualTeX.WordVsto.RedrawDocumentOmmlToLatex",
]) {
  assert.ok(wordVsto.includes(`id="${ribbonId}"`), `Word Ribbon is missing ${ribbonId}`);
}
assert.ok(!wordVsto.includes('id="VisualTeX.WordVsto.Delete"'));
assert.ok(!wordVsto.includes('id="VisualTeX.WordVsto.OpenDesktop"'));
assert.ok(!wordVsto.includes('label="删除所选公式"'));
assert.ok(!wordVsto.includes('label="打开 VisualTeX"'));
assert.ok(wordVstoService.includes("ConvertFormulaObjectsToLatex"));
assert.ok(wordVstoService.includes("WordOmmlNativeSource.RefreshForVisualTeX"));
assert.ok(wordVstoService.includes("BuildFormulaLatexSource"));
assert.ok(wordVstoService.includes("TryGetVisualTeXNumberedTable"));
assert.ok(wordEquationNumbering.includes("FreezeFormulaCrossReferences"));
for (const binding of [
  ['VisualTeX.WordVsto.Inline', 'tag="oleInline"'],
  ['VisualTeX.WordVsto.Display', 'tag="oleDisplay"'],
  ['VisualTeX.WordVsto.InlineOmml', 'tag="ommlInline"'],
  ['VisualTeX.WordVsto.DisplayOmml', 'tag="ommlDisplay"'],
  ['VisualTeX.WordVsto.Edit', 'tag="editSelected"'],
  ['VisualTeX.WordVsto.UpdateNumbers', 'tag="updateNumbers"'],
  ['VisualTeX.WordVsto.BulkImport', 'tag="batchImport"'],
]) {
  assert.ok(wordVsto.includes(`id="${binding[0]}`));
  assert.ok(wordVsto.includes(binding[1]));
}
for (const binding of [
  ['VisualTeX.PowerPointVsto.New', 'tag="insertFormula"'],
  ['VisualTeX.PowerPointVsto.Edit', 'tag="editSelected"'],
  ['VisualTeX.PowerPointVsto.ConvertSelected', 'tag="convertToOle"'],
]) {
  assert.ok(powerpointVsto.includes(`id="${binding[0]}`));
  assert.ok(powerpointVsto.includes(binding[1]));
}
for (const menuId of [
  "VisualTeX.WordVsto.VisualTeXToOmml",
  "VisualTeX.WordVsto.OmmlToVisualTeX",
]) {
  const menuStart = wordVsto.indexOf(`<menu id="${menuId}"`);
  assert.ok(menuStart >= 0, `Word Ribbon is missing ${menuId}`);
  const menuEnd = wordVsto.indexOf(">", menuStart);
  const menuTag = wordVsto.slice(menuStart, menuEnd + 1);
  assert.ok(!menuTag.includes("getImage="), `${menuId} unexpectedly has getImage`);
  assert.ok(!menuTag.includes("imageMso="), `${menuId} unexpectedly has imageMso`);
}
for (const callback of [
  "OnConvertVisualTeXToOmmlSelection",
  "OnConvertVisualTeXToOmmlDocument",
  "OnConvertOmmlToVisualTeXSelection",
  "OnConvertOmmlToVisualTeXDocument",
]) {
  assert.ok(wordVsto.includes(callback), `Word Ribbon is missing ${callback}`);
}
assert.ok(wordVsto.includes('getImage="GetRibbonImage"'));
assert.ok(powerpointVsto.includes('getImage="GetRibbonImage"'));
assert.ok(ribbonIconProvider.includes("GetIPictureDispFromPicture"));
assert.ok(ribbonIconProvider.includes("RibbonVectorIconRenderer.Keys"));
assert.ok(ribbonIconProvider.includes("RibbonVectorIconRenderer.Create"));
assert.ok(ribbonVectorIconRenderer.includes("internal const int PixelSize = 64"));
assert.ok(ribbonVectorIconRenderer.includes("internal const float Dpi = 192f"));
for (const iconKey of [
  "oleDisplay",
  "ommlDisplay",
  "oleInline",
  "ommlInline",
  "insertFormula",
  "updateNumbers",
  "editSelected",
  "convertToOmml",
  "convertToOle",
  "batchImport",
]) {
  assert.ok(
    ribbonVectorIconRenderer.includes(`"${iconKey}"`),
    `Vector Ribbon renderer is missing ${iconKey}`,
  );
}
assert.ok(ribbonIconData.includes("internal const string OleDisplay"));
assert.ok(ribbonIconData.includes("internal const string ConvertToOle"));
assert.ok(wordVsto.includes("service.ReplaceOmml"));
assert.ok(wordVsto.includes("service.InsertOmml"));
assert.ok(wordVsto.includes("mathMl = requiredMathMl"));
assert.ok(wordVsto.includes("targetObjectMode"));
assert.ok(wordVsto.includes("requiresObjectModeChange"));
assert.ok(wordVsto.includes("session.ExportResult is null"));
assert.ok(powerpointVsto.includes('BeginSession("create", "crossPlatformPicture", null)'));
assert.ok(powerpointVsto.includes('BeginSelectedSession("nativeOle", conversionOnly: true)'));
assert.ok(powerpointVsto.includes("capturedSelection"));
assert.ok(powerpointVsto.includes("ResolveFormulaSelection"));
assert.ok(powerpointVsto.includes("targetObjectMode"));
assert.ok(powerpointVsto.includes("requiresObjectModeChange"));
assert.ok(powerpointVsto.includes("new PowerPointDoubleClickHook"));
for (const vstoEntry of [wordVsto, powerpointVsto]) {
  assert.ok(vstoEntry.includes("MaterializeSvg"));
  assert.ok(vstoEntry.includes("OfficeOlePreview.CreateVectorEmfFromSvg"));
  assert.ok(vstoEntry.includes("File.Delete(emfPath)"));
}
assert.ok(wordVsto.includes("FormulaOleContract.NativeOleMode"));
assert.ok(powerpointVsto.includes('session.ObjectMode == "nativeOle"'));
for (const nativeService of [wordVstoService, powerpointVstoService]) {
  assert.ok(nativeService.includes("AddOLEObject"));
  assert.ok(nativeService.includes("FormulaOleContract.ProgId"));
  assert.ok(nativeService.includes("IVisualTeXFormulaObject"));
  assert.ok(nativeService.includes("FormulaOleInterop.Initialize"));
  assert.ok(nativeService.includes("FormulaOleInterop.Update"));
  assert.ok(nativeService.includes("TryUpdateOle"));
  assert.ok(nativeService.includes("TryDelete"));
  assert.ok(nativeService.includes("AddPicture"));
}
assert.ok(vstoOlePreview.includes("EmfType.EmfOnly"));
assert.ok(vstoOlePreview.includes("CreateVectorEmfFromSvg"));
assert.ok(vstoOlePreview.includes("ValidateVectorEmf"));
assert.ok(vstoOlePreview.includes("SVG external references are forbidden"));
assert.ok(vstoOlePreview.includes("Semi-transparent SVG paint cannot be represented"));
assert.ok(vstoOlePreview.includes("EMR_STRETCHDIBITS"));
assert.ok(vstoOlePreview.includes("EMF+ preview contains a raster image draw record"));
assert.ok(!vstoOlePreview.includes("CreateEmfFromPng"));
assert.ok(!vstoOlePreview.includes("DrawImage(image"));
assert.ok(vstoOlePngExtractor.includes("System.Runtime.InteropServices.ComTypes.IDataObject"));
assert.ok(vstoOlePngExtractor.includes("RegisterClipboardFormat(\"PNG\")"));
assert.ok(vstoOlePngExtractor.includes("ReleaseStgMedium"));
assert.ok(vstoOlePngExtractor.includes("MaxPngBytes"));
assert.ok(wordVstoService.includes("ExportSelectedOleAsPicture"));
assert.ok(powerpointVsto.includes("imagePath = client.MaterializeSvg(session)"));
assert.ok(powerpointVsto.includes('BeginSelectedSession("crossPlatformPicture", conversionOnly: true)'));
assert.ok(wordVstoService.includes("WordEquationNumbering.TryReconcile"));
assert.ok(wordVstoService.includes("WordEquationNumbering.Reconcile"));
assert.ok(wordEquationNumbering.includes("SEQ {nativeSequenceName}"));
assert.ok(wordEquationNumbering.includes('LegacyEquationSequenceName = "VisualTeXEquation"'));
assert.ok(!wordEquationNumbering.includes("WdCaptionLabelID.wdCaptionEquation"));
assert.ok(wordEquationNumbering.includes("EquationNumberFormatVariableName"));
assert.ok(wordEquationNumbering.includes("Document.Variables") || wordEquationNumbering.includes("document.Variables"));
assert.ok(wordEquationNumbering.includes("Heading1DotId"));
assert.ok(wordEquationNumbering.includes("Heading2DashId"));
assert.ok(wordEquationNumbering.includes("GetHeadingNumberAnchors"));
assert.ok(wordEquationNumbering.includes("ResolveEquationNumberScope"));
assert.ok(wordEquationNumbering.includes("WordEquationReferenceFields.InsertNavigableReference"));
assert.ok(wordEquationNumbering.includes("NativeNumberBookmarkName(target.FormulaId)"));
assert.ok(wordEquationReferenceFields.includes("WdFieldType.wdFieldGoToButton"));
assert.ok(wordEquationReferenceFields.includes("WdFieldType.wdFieldRef"));
assert.ok(wordEquationReferenceFields.includes("\\\\* CHARFORMAT \\\\!"));
assert.ok(wordEquationNumbering.includes("TryResolveVisualTeXReferenceBookmark"));
assert.ok(wordEquationNumbering.includes("EquationBookmarkPrefix"));
assert.ok(wordEquationNumbering.includes("UpdateNativeCrossReferences"));
assert.ok(!wordEquationNumbering.includes("EquationReferenceBookmarkPrefix"));
assert.ok(wordEquationNumbering.includes("WdTabAlignmentCenter"));
assert.ok(wordEquationNumbering.includes("WdTabAlignmentRight"));
assert.ok(wordEquationNumbering.includes("WordOmmlFormulaStore.BookmarkedFormulaIds"));
assert.ok(wordEquationNumbering.includes("WordOmmlFormulaStore.FindByFormulaId"));
assert.ok(wordEquationNumbering.includes("WordOmmlFormulaStore.GetEquationRange"));
assert.ok(wordEquationNumbering.includes("FormulaFontSize.ResolveSemanticFontSize"));
assert.ok(wordEquationNumbering.includes("ResolveEquationNumberLabelRange"));
assert.ok(wordEquationNumbering.includes('EquationNumberFontName = "Cambria Math"'));
assert.ok(wordEquationNumbering.includes("ApplyEquationNumberFont"));
assert.ok(wordVstoService.includes("paragraphFormat.DisableLineHeightGrid = -1"));
assert.ok(wordEquationNumberingTrueDisplay.includes("StyleNativeDisplayAnchorParagraph"));
assert.ok(wordEquationNumberingTrueDisplay.includes("WdWrapType.wdWrapNone"));
assert.ok(wordOmmlConverter.includes("MML2OMML.XSL"));
assert.ok(wordOmmlConverter.includes("FormattedText"));
assert.ok(wordOmmlConverter.includes("CreateBatchSource"));
assert.ok(wordOmmlConverter.includes("visualtex-omml-batch"));
assert.ok(wordOmmlConverter.includes("WdOMathType.wdOMathDisplay"));
assert.ok(wordOmmlFormulaStore.includes("SaveNewBatch"));
assert.ok(wordOmmlFormulaStore.includes("urn:visualtex:word-omml:1"));
assert.ok(wordOmmlFormulaStore.includes("VTOMML_"));
assert.ok(wordOmmlFormulaStore.includes("BookmarkPrefix"));
assert.ok(wordOmmlFormulaStore.includes("CustomXMLParts"));
assert.ok(wordVstoService.includes("ReplaceOmml"));
assert.ok(wordVstoService.includes("InsertOmml"));
assert.ok(wordVstoService.includes("WordOmmlFormulaStore.Save"));
assert.ok(formulaOleInterop.includes("ThrowIfFailed"));
assert.ok(formulaOleInterop.includes("GetFormulaJson"));
assert.ok(wordFormulaMetadataReader.includes("FormulaOleInterop.ReadMetadata"));
assert.ok(wordOleObjectAccessor.includes("wdOLEVerbShow"));
assert.ok(wordOleObjectAccessor.includes("format.DoVerb"));
assert.ok(nativeOfficeAcceptance.includes("real Word OMML/OLE and PowerPoint native OLE acceptance passed"));
assert.ok(nativeOfficeAcceptance.includes("VerifyWordMixedNumberingScenarios"));
assert.ok(nativeOfficeAcceptance.includes("VerifyPowerPointPictureToOleConversion"));
assert.ok(nativeOfficeAcceptance.includes("WordDoubleClickRouting.ShouldOpenVisualTeX"));
assert.ok(nativeOfficeAcceptance.includes("VerifyWordCachedPreviewOffline"));
assert.ok(nativeOfficeAcceptance.includes("VerifyPowerPointCachedPreviewOffline"));
assert.ok(nativeOfficeAcceptance.includes("UpdateAndVerifyWord"));
assert.ok(nativeOfficeAcceptance.includes("UpdateAndVerifyPowerPoint"));
for (const requirement of [
  "Assert-NoOfficeProcesses",
  "Resolve-OfficePlatform",
  "TimeoutSeconds",
  "TargetFrameworkRootPath",
  "Assert-NoVisualTeXRegistration",
  "ole-server-trace.enabled",
  "VisualTeX real Word/PowerPoint native OLE acceptance passed",
]) {
  assert.ok(
    nativeOfficeAcceptanceScript.includes(requirement),
    `Native Office acceptance script is missing ${requirement}`,
  );
}

const comContracts = await source(
  "src-windows/VisualTeX.WindowsOffice.Contracts/OfficeComInterfaces.cs",
);
assert.ok(!comContracts.includes("interface IOfficeComAddIn"));
assert.ok(!comContracts.includes("interface IOfficeRibbonExtensibility"));

const installOle = await source("scripts/install_windows_ole.ps1");
const installVsto = await source("scripts/install_windows_vsto.ps1");
const installVstoRuntime = await source("scripts/install_windows_vsto_runtime.ps1");
const prepareVstoRuntime = await source("scripts/prepare_windows_vsto_runtime.ps1");
const runtimeVerification = await source("scripts/test_windows_office_runtime.ps1");
const certificateInstaller = await source("scripts/ensure_windows_office_certificate.ps1");
const certificateUninstaller = await source("scripts/remove_windows_office_certificate.ps1");
const uninstallVsto = await source("scripts/uninstall_windows_vsto.ps1");
const buildWindowsOffice = await source("scripts/build_windows_office.ps1");
const ribbonDispatchSmoke = await source(
  "scripts/test_windows_vsto_ribbon_dispatch.ps1",
);
const dependencyLoadingSmoke = await source(
  "scripts/test_windows_vsto_dependency_loading.ps1",
);
const nativeMsi = await source(
  "src-windows/VisualTeX.WindowsOffice.Installer/Package.wxs",
);
const nativeMsiProject = await source(
  "src-windows/VisualTeX.WindowsOffice.Installer/VisualTeX.WindowsOffice.Installer.wixproj",
);
assert.ok(installOle.includes("forwarding to the native Ribbon + OLE LocalServer installer"));
assert.ok(!installVsto.includes("uninstall_windows_ole.ps1"));
assert.ok(!installVsto.includes("ensure_windows_office_certificate.ps1"));
assert.ok(installVsto.includes('"/L*v"'));
assert.ok(installVsto.includes("RelatedProducts"));
assert.ok(installVsto.includes("Get-FileHash"));
assert.ok(installVsto.includes("Wait-ForRelatedProductCount"));
assert.ok(installVsto.includes("product state did not settle"));
assert.ok(installVsto.includes("Assert-NativeOleRegistration"));
assert.ok(installVsto.includes("ServerExecutable"));
assert.ok(installVsto.includes("3,1,32,1"));
assert.ok(installVsto.includes("VisualTeX.FormulaOleServer.exe"));
assert.ok(installVsto.includes('-Name "Mode" -PropertyType String -Value "vsto"'));
assert.ok(installVsto.includes('-Name "NativeOleEnabled"'));
assert.ok(installVsto.includes("hashManifest.dependencies"));
assert.ok(installVsto.includes("Assert-NoOfficeProcesses"));
assert.ok(installVsto.includes("Assert-VstoRuntimeInstalled"));
assert.ok(installVsto.includes("Assert-NetFramework472Installed"));
assert.ok(installVsto.includes("$minimumRelease = 461808"));
assert.ok(installVsto.includes(".NET Framework 4.7.2 or newer is required"));
assert.ok(!installVsto.includes("Assert-NetFramework48Installed"));
assert.ok(!installVsto.includes("expected at least 528040"));
assert.ok(installVsto.includes("Assert-OfficeApplicationsInstalled"));
assert.ok(installVsto.includes("Assert-MsiArchitecture"));
assert.ok(installVsto.includes("Assert-OfficeAddinRegistration"));
assert.ok(installVsto.includes("Resolve-MachineOfficeInstallRoot"));
assert.ok(installVsto.includes("ProgramW6432"));
assert.ok(installVsto.includes("Resolve-PowerShellExecutable"));
assert.ok(installVsto.includes("Sysnative\\WindowsPowerShell"));
assert.ok(installVsto.includes("ArchitectureRelaunched"));
assert.ok(installVsto.includes("TargetProcessPlatform"));
assert.ok(installVsto.includes("Remove-LegacyPerUserOfficeRegistration"));
assert.ok(installVsto.includes("Assert-ManagedComActivation"));
assert.ok(installVsto.includes("Stop-VisualTeXProcessesForRepair"));
assert.ok(installVsto.includes('$startParameters.Verb = "RunAs"'));
assert.ok(installVsto.includes("RegistryHive]::LocalMachine"));
assert.ok(installVsto.includes("FilesAndRegistryVerified"));
assert.ok(installVsto.includes("OfficeRuntimeVerified"));
assert.ok(installVsto.includes("Diagnostic report"));
assert.ok(!installVsto.includes("chain-verified"));
assert.ok(installVsto.includes('"WINWORD", "POWERPNT"'));
assert.ok(installVsto.includes('"MSIRESTARTMANAGERCONTROL=Disable"'));
assert.ok(installVsto.includes('"REBOOT=ReallySuppress"'));
assert.ok(installVsto.includes('"vsto-bootstrap-$stamp.log"'));
for (const required of [
  "VSTORFeature_CLR40",
  "Get-AuthenticodeSignature",
  "Microsoft Corporation",
  'ArgumentList @("/quiet", "/norestart")',
  "-Verb RunAs",
  "1641",
  "3010",
  "CheckOnly",
]) {
  assert.ok(installVstoRuntime.includes(required), `VSTO Runtime installer missing ${required}`);
}
for (const required of [
  "download.microsoft.com",
  "CFE1A40BBE4A50022DB2164ABDB0154984E2CECB761A23CDC81CB5754F6E0A18",
  "10.0.60917.00",
  "Get-AuthenticodeSignature",
  "VISUALTEX_VSTO_RUNTIME_PATH",
]) {
  assert.ok(prepareVstoRuntime.includes(required), `VSTO Runtime preparation missing ${required}`);
}
assert.ok(uninstallVsto.includes("DF66EC66-3B3A-4675-A7BE-30456A04EB96"));
assert.ok(uninstallVsto.includes('Name "NativeOleEnabled"'));
assert.ok(uninstallVsto.includes("Resolve-PowerShellExecutable"));
assert.ok(uninstallVsto.includes("ArchitectureRelaunched"));
assert.ok(uninstallVsto.includes("Sysnative\\WindowsPowerShell"));
assert.ok(uninstallVsto.includes("vsto-uninstall-bootstrap-$stamp.log"));
assert.ok(uninstallVsto.includes("certificate-remove-$stamp.log"));
assert.ok(uninstallVsto.includes("remove_windows_office_certificate.ps1"));
assert.ok(uninstallVsto.includes("Get-Process visualtex"));
assert.ok(uninstallVsto.includes("Stopping VisualTeX process"));
assert.ok(uninstallVsto.includes("$process.WaitForExit()"));
assert.ok(certificateUninstaller.includes("reg.exe"));
assert.ok(certificateUninstaller.includes("SystemCertificates\\Root\\Certificates"));
assert.ok(certificateUninstaller.includes("WaitForExit"));
assert.ok(certificateUninstaller.includes("TimeoutSeconds"));
assert.ok(certificateUninstaller.includes("Test-CertificatePresent"));
assert.ok(!certificateUninstaller.includes("X509Store"));
assert.ok(!certificateUninstaller.includes("certutil.exe"));
assert.ok(!uninstallVsto.includes("Start-Process -FilePath \"powershell.exe\" `\n        -Verb RunAs"));
assert.ok(buildWindowsOffice.includes("test_windows_formula_ole_server.ps1"));
assert.ok(buildWindowsOffice.includes("Stop-BuildOleServerProcesses"));
assert.ok(buildWindowsOffice.includes("acceptance-owned OLE Server"));
assert.ok(buildWindowsOffice.includes("test_windows_vsto_ribbon_dispatch.ps1"));
assert.ok(buildWindowsOffice.includes("test_windows_vsto_dependency_loading.ps1"));
assert.ok(buildWindowsOffice.includes("dependencyEntries"));
assert.ok(buildWindowsOffice.includes("formulaOleServer"));
assert.ok(ribbonDispatchSmoke.includes("ComDefaultInterfaceAttribute"));
assert.ok(ribbonDispatchSmoke.includes("InterfaceIsIDispatch"));
assert.ok(ribbonDispatchSmoke.includes("DispIdAttribute"));
assert.ok(ribbonDispatchSmoke.includes("QueryInterface"));
assert.ok(ribbonDispatchSmoke.includes("VisualTeX.WordVsto.Tab"));
assert.ok(ribbonDispatchSmoke.includes("VisualTeX.PowerPointVsto.Tab"));
assert.ok(ribbonDispatchSmoke.includes("SysWOW64"));
assert.ok(ribbonDispatchSmoke.includes("\\net472\\VisualTeX.WordVsto.dll"));
assert.ok(ribbonDispatchSmoke.includes("\\net472\\VisualTeX.PowerPointVsto.dll"));
assert.ok(dependencyLoadingSmoke.includes("System.Text.Json, Version=8.0.0.0"));
assert.ok(dependencyLoadingSmoke.includes("\\net472\\VisualTeX.WordVsto.dll"));
assert.ok(dependencyLoadingSmoke.includes("\\net472\\VisualTeX.PowerPointVsto.dll"));
assert.ok(nativeMsi.includes("3,1,32,1"));
assert.ok(nativeMsi.includes('Name="CodeBase"'));
assert.ok(nativeMsi.includes('Name="Mode" Type="string" Value="vsto"'));
assert.ok(nativeMsi.includes('Scope="perMachine"'));
assert.ok(nativeMsi.includes('<StandardDirectory Id="ProgramFiles6432Folder">'));
assert.ok(nativeMsi.includes('Root="HKLM"'));
assert.ok(!nativeMsi.includes('Root="HKCU"'));
assert.ok(!nativeMsi.includes('LocalAppDataFolder'));
assert.ok(nativeMsi.includes('Bitness="$(var.ComponentBitness)"'));
assert.ok(nativeMsiProject.includes(">always32</ComponentBitness>"));
assert.ok(nativeMsiProject.includes(">always64</ComponentBitness>"));
assert.ok(nativeMsiProject.includes("ComponentBitness=$(ComponentBitness)"));
assert.ok(nativeMsiProject.includes("VisualTeX.WordVsto\\bin\\$(Platform)\\$(Configuration)\\net472"));
assert.ok(nativeMsiProject.includes("VisualTeX.PowerPointVsto\\bin\\$(Platform)\\$(Configuration)\\net472"));
assert.ok(!nativeMsi.includes("OleManifestEnabled"));
assert.ok(dependencyLoadingSmoke.includes("System.Numerics.Vectors, Version=4.1.4.0"));
assert.ok(dependencyLoadingSmoke.includes("SerializeJson"));
assert.ok(dependencyLoadingSmoke.includes("SysWOW64"));
assert.ok(buildWindowsOffice.includes('StartsWith("8.")'));
assert.ok(buildWindowsOffice.includes('$msbuild $installerProject'));
assert.ok(buildWindowsOffice.includes('PackagePlatform = "x64"'));
assert.ok(buildWindowsOffice.includes('PackagePlatform = "x86"'));
assert.ok(buildWindowsOffice.includes('OlePlatform = "Win32"'));
assert.ok(buildWindowsOffice.includes("VisualTeX-WindowsOffice-VSTO-$packagePlatform.msi"));
assert.ok(buildWindowsOffice.includes('$vstoTargetFramework = "net472"'));
assert.ok(buildWindowsOffice.includes("microsoft.netframework.referenceassemblies.$vstoTargetFramework"));
assert.ok(installVsto.includes("Resolve-OfficePlatform"));
assert.ok(installVsto.includes("RegistryView]::Registry32"));
assert.ok(installVsto.includes("RegistryView]::Registry64"));
assert.ok(installVsto.includes("PackageDirectory"));
assert.ok(installVsto.includes("MSI SHA-256 mismatch"));
assert.ok(runtimeVerification.includes("Get-ManagedComRegistrationState"));
assert.ok(runtimeVerification.includes("discoveredProgIds"));
assert.ok(runtimeVerification.includes("enumerate-com-addins"));
assert.ok(runtimeVerification.includes("connectAttempted"));
assert.ok(!runtimeVerification.includes("COMAddIns.Item($ProgId)"));
for (const msiRequirement of [
  "FormulaOleServerExecutable",
  "LocalServer32",
  "ServerExecutable",
  "InprocHandler32",
  "Ole32.dll",
  "VisualTeX.Formula.1",
  "ProxyStubClsid32",
  "DF66EC66-3B3A-4675-A7BE-30456A04EB96",
  "NativeOleEnabled",
  "FilesAndRegistryVerified",
  "OfficeRuntimeVerified",
  "SystemNumericsVectors",
  "SystemValueTuple",
]) {
  assert.ok(nativeMsi.includes(msiRequirement), `Native MSI is missing ${msiRequirement}`);
}
assert.ok(!nativeMsi.includes("CustomAction"));
assert.ok(!installOle.includes("TrustedCatalog"));
assert.ok(runtimeVerification.includes("Get-ComAddInItem"));
assert.ok(runtimeVerification.includes("Resolve-OfficeExecutablePath"));
assert.ok(runtimeVerification.includes("Resolve-PowerShellExecutable"));
assert.ok(runtimeVerification.includes("Sysnative\\WindowsPowerShell"));
assert.ok(runtimeVerification.includes("ArchitectureRelaunched"));
assert.ok(runtimeVerification.includes("ProgramW6432"));
assert.ok(runtimeVerification.includes("GetActiveObject"));
assert.ok(runtimeVerification.includes("RuntimeVerificationPending"));
assert.ok(runtimeVerification.includes("Start-CompanionAsInteractiveUser"));
assert.ok(runtimeVerification.includes("must run in the interactive user's non-elevated session"));
assert.ok(!runtimeVerification.includes("Shell.Application"));
assert.ok(runtimeVerification.includes('startupMode = "desktop-executable-rot"'));
assert.ok(runtimeVerification.includes("desktop application did not enumerate"));
assert.ok(!runtimeVerification.includes("New-Object -ComObject $comType"));
assert.ok(runtimeVerification.includes('"VisualTeX.WordVsto"'));
assert.ok(runtimeVerification.includes('"VisualTeX.PowerPointVsto"'));
assert.ok(runtimeVerification.includes("Get-DisabledItems"));
assert.ok(runtimeVerification.includes("Get-RecentOfficeLoadEvents"));
assert.ok(runtimeVerification.includes("FilesAndRegistryVerified"));
assert.ok(runtimeVerification.includes('ArgumentList "-Embedding"'));
assert.ok(runtimeVerification.includes("Native Office integration installed and verified successfully"));
assert.ok(certificateInstaller.includes('Split-Path -Parent $PSScriptRoot'));
assert.ok(certificateInstaller.includes("-LiteralPath"));
assert.ok(certificateInstaller.includes("explicitly supplied VisualTeX executable"));

const platformBundle = await source("scripts/build_platform_bundle.mjs");
const tauriBuild = await source("scripts/tauri_build.mjs");
const windowsPowerShell = await source("scripts/windows_powershell.mjs");
const officeLifecycle = await source("src-tauri/src/office/lifecycle.rs");
const windowsBundle = await source("src-tauri/tauri.windows.conf.json");
const installerHooks = await source("src-tauri/windows/hooks.nsh");
const nsisPatch = await source("scripts/patch_generated_nsis.ps1");
assert.ok(platformBundle.includes('"scripts/build_windows_office.ps1"'));
assert.ok(platformBundle.includes("windowsPowerShellPath"));
assert.ok(!platformBundle.includes('run("powershell.exe"'));
assert.ok(tauriBuild.includes("windowsPowerShellPath"));
assert.ok(!tauriBuild.includes('run("powershell.exe"'));
assert.ok(windowsPowerShell.includes('"System32"'));
assert.ok(windowsPowerShell.includes('"WindowsPowerShell"'));
assert.ok(officeLifecycle.includes("windows_powershell_executable"));
assert.ok(!officeLifecycle.includes('hidden_windows_command("powershell.exe")'));
assert.ok(platformBundle.includes('"scripts/prepare_windows_vsto_runtime.ps1"'));
assert.ok(platformBundle.includes('"-SkipTests"'));
assert.ok(!platformBundle.includes('"scripts/build_windows_ole_bridge.ps1"'));
assert.ok(windowsBundle.includes('"../scripts/install_windows_vsto.ps1"'));
assert.ok(windowsBundle.includes('"../scripts/install_windows_vsto_runtime.ps1"'));
assert.ok(!windowsBundle.includes('"../scripts/install_windows_ole.ps1"'));
for (const bundledOfficeResource of [
  "VisualTeX-WindowsOffice-VSTO-x64.msi",
  "VisualTeX-WindowsOffice-VSTO-x64.sha256.json",
  "VisualTeX-WindowsOffice-VSTO-x86.msi",
  "VisualTeX-WindowsOffice-VSTO-x86.sha256.json",
  "vstor_redist.exe",
  "vstor_redist.sha256.json",
]) {
  assert.ok(windowsBundle.includes(bundledOfficeResource));
}
assert.ok(installerHooks.includes("${NSD_Check} $VisualTeXOfficeNativeRadio"));
assert.ok(installerHooks.includes("bundled private Python 3.12.10 x64 runtime"));
assert.ok(!installerHooks.includes("VisualTeXProbeLauncher"));
assert.ok(!installerHooks.includes("VisualTeXProbeCommand"));
assert.ok(!installerHooks.includes('`powershell.exe '));
assert.ok(installerHooks.includes('$WINDIR\\System32\\WindowsPowerShell\\v1.0\\powershell.exe'));
assert.ok(!installerHooks.includes("MUI_CUSTOMFUNCTION_GUIINIT"));
assert.ok(!installerHooks.includes("VisualTeXDefaultMaintenanceUninstall"));
assert.ok(installerHooks.includes("generated Tauri PageReinstall function is patched"));
assert.ok(nsisPatch.includes("Same-version maintenance defaults to the second option"));
assert.ok(nsisPatch.includes('SendMessage $R3 ${BM_SETCHECK} ${BST_CHECKED} 0'));
assert.ok(nsisPatch.includes('StrCpy $ReinstallPageCheck 2'));
assert.ok(nsisPatch.includes('${NSD_SetFocus} $R3'));
assert.ok(installerHooks.includes("VisualTeXRepairMainUninstallRegistration"));
assert.ok(installerHooks.includes('WriteRegStr HKCU "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\VisualTeX"'));
assert.ok(installerHooks.includes('$INSTDIR == "$PROFILE\\AppData\\VisualTeX"'));
assert.ok(installerHooks.includes("vsto-uninstall-bootstrap"));
assert.ok(installerHooks.includes("certificate-remove"));
assert.ok(installerHooks.includes("NSIS_HOOK_POSTUNINSTALL"));
assert.ok(installerHooks.includes('DeleteRegKey HKCU "Software\\visualtex\\VisualTeX"'));
assert.ok(installerHooks.includes('RMDir /r "$INSTDIR"'));
assert.ok(installerHooks.includes("OfficeSessions user data"));
assert.ok(installerHooks.includes("Remove only known legacy application payloads."));
assert.ok(installerHooks.includes("%APPDATA%\\VisualTeX\\ocr-storage.json"));
assert.ok(!installerHooks.includes('RMDir /r "$APPDATA\\VisualTeX"'));
assert.ok(!installerHooks.includes('Delete "$APPDATA\\VisualTeX\\ocr-storage.json"'));
const postUninstallHook = installerHooks.slice(installerHooks.indexOf("!macro NSIS_HOOK_POSTUNINSTALL"));
assert.ok(postUninstallHook.includes("GetCurrentProcessId"));
assert.ok(postUninstallHook.includes("Wait-Process -Id $0"));
assert.ok(postUninstallHook.includes("Remove-Item -LiteralPath '$INSTDIR'"));
assert.ok(!postUninstallHook.includes('RMDir /r "$APPDATA\\VisualTeX"'));
assert.ok(installerHooks.includes("SetErrorLevel 1"));
assert.ok(!installerHooks.includes('ExecToLog `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\\scripts\\remove_windows_office_certificate.ps1"'));
assert.ok(installerHooks.includes('${If} $VisualTeXOfficeChoice == ""'));
assert.ok(installerHooks.includes('StrCpy $VisualTeXOfficeChoice "native"'));
assert.ok(installerHooks.includes("install_windows_vsto.ps1"));
assert.ok(!installerHooks.includes("-CompanionOnly"));
assert.ok(installerHooks.includes("visualtex_office_static_installed"));
assert.ok(installerHooks.includes("RuntimeVerificationPending"));
assert.ok(installerHooks.includes("without leaving a resident VisualTeX process"));
assert.ok(installerHooks.includes("安装阶段不会启动常驻后台进程"));
assert.ok(installerHooks.includes("Companion and Word/PowerPoint connection verification are deferred"));
assert.ok(installerHooks.includes("install_windows_vsto_runtime.ps1"));
assert.ok(installerHooks.includes("vstor_redist.exe"));
assert.ok(installerHooks.includes("vstor_redist.sha256.json"));
assert.ok(installerHooks.includes("visualtex_vsto_runtime_install"));
assert.ok(installerHooks.includes("visualtex_vsto_runtime_declined"));
assert.ok(installerHooks.includes("-CheckOnly"));
assert.ok(installerHooks.includes("UAC"));
assert.ok(installerHooks.includes("ensure_windows_office_certificate.ps1"));
assert.ok(installerHooks.includes("test_windows_office_runtime.ps1"));
assert.ok(installerHooks.includes("VisualTeX-WindowsOffice-VSTO-x64.msi"));
assert.ok(installerHooks.includes("VisualTeX-WindowsOffice-VSTO-x86.msi"));
assert.ok(installerHooks.includes("VisualTeX-WindowsOffice-VSTO-x64.sha256.json"));
assert.ok(installerHooks.includes("VisualTeX-WindowsOffice-VSTO-x86.sha256.json"));
assert.ok(installerHooks.includes('-PackageDirectory "$INSTDIR\\windows-office"'));
assert.ok(installerHooks.includes('-VisualTeXPath "$INSTDIR\\${MAINBINARYNAME}.exe"'));
assert.ok(!installerHooks.includes('-VisualTeXPath "$INSTDIR\\VisualTeX.exe"'));
assert.ok(installerHooks.includes("uninstall_windows_vsto.ps1"));
assert.ok(!installerHooks.includes("uninstall_windows_ole.ps1"));
assert.ok(installerHooks.includes("Get-Process WINWORD,POWERPNT"));
assert.ok(installerHooks.includes("Stop-Process -Force"));
assert.ok(installerHooks.includes("IDYES visualtex_force_close_office"));
assert.ok(installerHooks.includes("未保存的 Office 文档可能丢失"));
assert.ok(installerHooks.includes("选择“否”将返回上一页"));
assert.ok(installerHooks.indexOf("IDYES visualtex_force_close_office") < installerHooks.indexOf("Stop-Process -Force"));
assert.ok(!installerHooks.includes("VisualTeXOfficeOleRadio"));
assert.ok(!installerHooks.includes('VisualTeXOfficeChoice == "ole"'));
assert.ok(
  !installerHooks.includes('-File "$INSTDIR\\scripts\\install_windows_ole.ps1"'),
);
assert.ok(installerHooks.includes("companion"));
assert.ok(installerHooks.includes("Word/PowerPoint connection verification are deferred"));
assert.ok(!installerHooks.includes("visualtex_office_static_runtime_verified"));
assert.ok(installerHooks.includes("IfSilent visualtex_office_done 0"));
assert.ok(installerHooks.includes("Office bootstrap completed without leaving a resident VisualTeX process"));
assert.ok(!installerHooks.includes("Office Add-ins dialogs"));
assert.ok(!installerHooks.includes("Automatically configuring Word and PowerPoint"));

const nativeOfficeVite = await source("vite.office.windows-native.config.ts");
const nativeDialogHtml = await source("office-dialog.html");
const nativeDialogEntry = await source("src/office/dialog/main.tsx");
const officeTsConfig = await source("tsconfig.office.json");
assert.ok(nativeOfficeVite.includes('publicDir: false'));
assert.ok(nativeOfficeVite.includes('dialog: resolve(root, "office-dialog.html")'));
assert.ok(!nativeOfficeVite.includes("office-windows-ole-bridge.html"));
assert.ok(nativeOfficeVite.includes("retired Office.js branch"));
assert.ok(!nativeDialogHtml.includes("office.js"));
assert.ok(nativeDialogEntry.includes("executes the Office.js runtime"));
assert.ok(!nativeDialogEntry.includes("Office.onReady"));
assert.ok(!officeTsConfig.includes('"office-js"'));
assert.ok(!officeTsConfig.includes('src/office/windows-ole'));

const windowsPipe = await source("src-tauri/src/office/windows_pipe.rs");
const windowsBridgeProgram = await source("src-windows/VisualTeX.WindowsOleBridge/Program.cs");
assert.ok(windowsPipe.includes('.arg("--parent-pid")'));
assert.ok(windowsPipe.includes('let bundled_filename = "visualtex-windows-office-bridge.exe"'));
assert.ok(windowsBridgeProgram.includes('Required(options, "parent-pid")'));
assert.ok(windowsBridgeProgram.includes("parent.Exited"));
assert.ok(windowsBridgeProgram.includes("parent.Exited -= parentExited"));
assert.ok(windowsBridgeProgram.includes("catch (ObjectDisposedException)"));
assert.ok(doubleClickHook.includes("WmLButtonDown"));
assert.ok(doubleClickHook.includes("GetDoubleClickTime"));
assert.ok(doubleClickHook.includes("SmCxDoubleClk"));
assert.ok(doubleClickHook.includes("GetOfficeForegroundHost"));
assert.ok(backend.includes("CaptureDoubleClickTargetAsync"));
assert.ok(backend.includes('string.Equals(host, "word"'));
assert.ok(backend.includes("metadata = selection.Metadata"));

const acceptance = await source("scripts/run_windows_office_acceptance.ps1");
for (const requirement of [
  "FormulaCount = 20",
  "word.insertInlineFormula",
  "word.insertDisplayFormula",
  "word.replaceFormula",
  "powerpoint.insertFormula",
  "powerpoint.replaceFormula",
  "powerpoint.deleteFormula",
  "Randomly editing Word formulas",
  "Randomly editing PowerPoint formulas",
  "PowerPoint slide-show",
  "Read-only Word",
  "multiple documents/windows",
  "Word undo",
  "PowerPoint undo",
  "bridge crash/restart",
  "TestModeSwitch",
]) {
  assert.ok(acceptance.includes(requirement), `Windows acceptance harness is missing ${requirement}`);
}

const tests = (
  await Promise.all([
    source("src-windows/VisualTeX.WindowsOffice.Tests/ProtocolSecurityTests.cs"),
    source("src-windows/VisualTeX.WindowsOffice.Tests/StaDispatcherTests.cs"),
    source("src-windows/VisualTeX.WindowsOffice.Tests/ReplacementTransactionTests.cs"),
    source("src-windows/VisualTeX.WindowsOffice.Tests/MetadataAndTempPathTests.cs"),
    source("src-windows/VisualTeX.WindowsOffice.Tests/ComReleaseAndDoubleClickTests.cs"),
    source("src-windows/VisualTeX.WindowsOffice.Tests/PowerPointOleSizingTests.cs"),
  ])
).join("\n");
for (const requirement of [
  "PipeTokenComparisonRejectsMismatch",
  "AllOfficeWorkRunsOnOneStaThread",
  "FailedConfigurationKeepsOriginalAndDeletesCandidate",
  "FormulaMetadataRoundTripsWithPersistentUuid",
  "FinalReleaseInvalidatesARealComObject",
  "FormulaDoubleClickDeduplicatesOnlyTheSamePersistentTarget",
  "LongerReplacementKeepsTheExistingVisualScale",
  "LegacyFormulaUsesItsPhysicalHeightAsTheFontSizeReference",
]) {
  assert.ok(tests.includes(requirement), `Windows test coverage is missing ${requirement}`);
}

console.log("Windows Office architecture smoke test passed");
