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
assert.ok(wordProject.includes("net48"));
assert.ok(powerpointProject.includes("net48"));
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
const wordDoubleClickHook = await source(
  "src-windows/VisualTeX.WordVsto/WordDoubleClickHook.cs",
);
const vstoOlePngExtractor = await source(
  "src-windows/VisualTeX.WindowsOffice.VstoShared/OlePngPreviewExtractor.cs",
);
const wordEquationNumbering = await source(
  "src-windows/VisualTeX.WindowsOffice.VstoShared/WordEquationNumbering.cs",
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
const officeServer = await source("src-tauri/src/office/server.rs");
const officeSessions = await source("src-tauri/src/office/sessions.rs");
const officeDialogMain = await source("src/office/dialog/main.tsx");
const officeDialogMessages = await source("src/office/dialog/dialogMessages.ts");
assert.ok(!sessionClient.includes("_installToken = ReadInstallToken()"));
assert.ok(sessionClient.includes("StartVisualTeXCompanion"));
assert.ok(sessionClient.includes("timeout.CancelAfter(TimeSpan.FromSeconds(2))"));
assert.ok(sessionClient.includes("OpenEditorAsync"));
assert.ok(sessionClient.includes("CloseEditorAsync"));
assert.ok(sessionClient.includes("/api/v1/app/sessions/"));
assert.ok(sessionClient.includes('}/close"'));
assert.ok(!sessionClient.includes('new Uri(CompanionOrigin, $"/dialog/'));
assert.ok(wordVsto.includes("await client.OpenEditorAsync"));
assert.ok(powerpointVsto.includes("await client.OpenEditorAsync"));
assert.ok(officeServer.includes("open_desktop_session_window"));
assert.ok(officeServer.includes("bring_session_window_to_front"));
assert.ok(officeServer.includes("set_always_on_top(true)"));
assert.ok(officeServer.includes("request_user_attention"));
assert.ok(officeServer.includes("WebviewWindowBuilder::new"));
assert.ok(officeServer.includes('"/app/sessions/{session_id}/open"'));
assert.ok(officeServer.includes('"/app/sessions/{session_id}/close"'));
assert.ok(officeServer.includes("close_desktop_session"));
assert.ok(officeServer.includes("WebviewUrl::External(url)"));
assert.ok(officeServer.includes("?runtime=vsto-desktop"));
assert.ok(officeServer.includes("remove_office_js"));
assert.ok(officeServer.includes('"<script src=\\\"/vendor/office-js/office.js\\\"></script>"'));
assert.ok(officeDialogMain.includes('get("runtime") === "vsto-desktop"'));
assert.ok(officeDialogMain.includes("if (isVstoDesktopRuntime)"));
assert.ok(officeDialogMain.includes("else {\n  mount();\n}"));
assert.ok(officeDialogMessages.includes('get("runtime") === "vsto-desktop"'));
assert.ok(officeDialogMessages.includes("return false"));
assert.ok(officeDialogMessages.includes("return true"));
assert.ok(officeDialogMain.includes("if (isVstoDesktopRuntime)"));
const officeDialogApp = await source("src/office/dialog/OfficeDialogApp.tsx");
assert.ok(officeDialogApp.includes("closeOfficeSessionWindow"));
assert.ok(officeDialogApp.includes("IS_VSTO_DESKTOP_RUNTIME"));
assert.ok(officeDialogApp.includes("if (!delivered) window.close()"));
assert.ok(officeDialogApp.includes("const unchangedEdit = session?.mode === \"edit\" && !dirty"));
assert.ok(officeDialogApp.includes("originalFingerprintRef"));
assert.ok(officeDialogApp.includes("originalFingerprintRef.current = loadedFingerprint"));
assert.ok(officeDialogApp.includes("normalizeOfficeCodeFormat"));
assert.ok(officeDialogApp.includes('return "raw"'));
assert.ok(officeDialogApp.includes("setLatexCodeFormat(loadedCodeFormat)"));
assert.ok(officeSessions.includes("unchanged_edit"));
assert.ok(officeSessions.includes("unchanged_edit_can_complete_without_new_export_result"));
assert.ok(officeSessions.includes("changed_edit_still_requires_a_new_export_result"));
assert.ok(wordVstoService.includes("shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse"));
assert.ok(wordVstoService.includes("shape.Height = height"));
assert.ok(wordVstoService.includes("ApplyInlineBaseline(shape, shape.Height"));
assert.ok(wordVstoService.includes("RestoreTypingBaselineAfter(shape)"));
assert.ok(wordVstoService.includes("font.Position = 0"));
assert.ok(wordVstoService.includes("OfficeFormulaSizing.EditedSize"));
assert.ok(officeFormulaSizing.includes("Formula height is the visual font-size reference"));
assert.ok(powerpointVstoService.includes("ApplyOleSizeAndRefresh"));
assert.ok(powerpointVstoService.includes("RestoreOlePosition"));
assert.ok(powerpointVstoService.includes("OfficeFormulaSizing.EditedSize"));
assert.ok(!powerpointVstoService.includes("format.DoVerb(-1)"));
assert.ok(powerpointVstoService.includes("CF_METAFILEPICT"));
assert.ok(powerpointVstoService.includes("shape.Type is not MsoShapeType.msoEmbeddedOLEObject"));
assert.ok(powerpointVstoService.includes("return ReadPictureMetadata(shape)"));
assert.ok(powerpointVstoService.includes("TryApplyRotation"));
assert.ok(wordVstoService.includes("Marshal.ReleaseComObject(value)"));
assert.ok(powerpointVstoService.includes("Marshal.ReleaseComObject(value)"));
assert.ok(!wordVstoService.includes("Marshal.FinalReleaseComObject(value)"));
assert.ok(!powerpointVstoService.includes("Marshal.FinalReleaseComObject(value)"));
assert.ok(vstoFlowAcceptanceProject.includes("VisualTeX.WordVsto"));
assert.ok(vstoFlowAcceptanceProject.includes("VisualTeX.PowerPointVsto"));
assert.ok(vstoFlowAcceptance.includes("Word formula aspect ratio is distorted"));
assert.ok(vstoFlowAcceptance.includes("Word inline formula baseline is incorrect"));
assert.ok(vstoFlowAcceptance.includes("PowerPoint formula must be a picture"));
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
assert.ok(wordDoubleClickHook.includes("WhMouseLl = 14"));
assert.ok(wordDoubleClickHook.includes("return new IntPtr(1)"));
assert.ok(wordDoubleClickHook.includes('"WINWORD"'));
for (const command of [
  "OnConvertSelected",
  "OnDeleteSelected",
  "OnExportSelectedAsPicture",
  "OnOpenDesktop",
]) {
  assert.ok(wordVsto.includes(command), `Word Ribbon is missing ${command}`);
  assert.ok(powerpointVsto.includes(command), `PowerPoint Ribbon is missing ${command}`);
}
assert.ok(wordVsto.includes("OnUpdateEquationNumbers"));
assert.ok(wordVsto.includes("OnInsertEquationReference"));
assert.ok(wordVsto.includes("InsertEquationReferenceAsync"));
assert.ok(wordVsto.includes("FormulaOleContract.NativeOleMode"));
assert.ok(wordVsto.includes("FormulaOleContract.WordOmmlMode"));
assert.ok(wordVsto.includes("OnInsertInlineOmml"));
assert.ok(wordVsto.includes("OnInsertDisplayOmml"));
assert.ok(wordVsto.includes("OnConvertSelectedToOmml"));
for (const binding of [
  ['VisualTeX.WordVsto.Inline', 'tag="oleInline"'],
  ['VisualTeX.WordVsto.Display', 'tag="oleDisplay"'],
  ['VisualTeX.WordVsto.InlineOmml', 'tag="ommlInline"'],
  ['VisualTeX.WordVsto.DisplayOmml', 'tag="ommlDisplay"'],
  ['VisualTeX.WordVsto.Edit', 'tag="editSelected"'],
  ['VisualTeX.WordVsto.ConvertSelected"', 'tag="convertToOle"'],
  ['VisualTeX.WordVsto.ConvertSelectedToOmml', 'tag="convertToOmml"'],
  ['VisualTeX.WordVsto.UpdateNumbers', 'tag="updateNumbers"'],
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
assert.ok(wordVsto.includes('getImage="GetRibbonImage"'));
assert.ok(powerpointVsto.includes('getImage="GetRibbonImage"'));
assert.ok(ribbonIconProvider.includes("GetIPictureDispFromPicture"));
assert.ok(ribbonIconProvider.includes("RibbonIconData.OleDisplay"));
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
assert.ok(powerpointVstoService.includes("ExportSelectedOleAsPicture"));
assert.ok(wordVstoService.includes("WordEquationNumbering.TryReconcile"));
assert.ok(wordVstoService.includes("WordEquationNumbering.Reconcile"));
assert.ok(wordEquationNumbering.includes("SEQ {nativeSequenceName}"));
assert.ok(wordEquationNumbering.includes("WdCaptionLabelID.wdCaptionEquation"));
assert.ok(wordEquationNumbering.includes("selection.InsertCrossReference"));
assert.ok(wordEquationNumbering.includes("WdReferenceKind.wdEntireCaption"));
assert.ok(wordEquationNumbering.includes("EquationBookmarkPrefix"));
assert.ok(wordEquationNumbering.includes("UpdateNativeCrossReferences"));
assert.ok(!wordEquationNumbering.includes("EquationReferenceBookmarkPrefix"));
assert.ok(wordEquationNumbering.includes("WdTabAlignmentCenter"));
assert.ok(wordEquationNumbering.includes("WdTabAlignmentRight"));
assert.ok(wordEquationNumbering.includes("WordOmmlFormulaStore.FormulaIds"));
assert.ok(wordEquationNumbering.includes("WordOmmlFormulaStore.FindByFormulaId"));
assert.ok(wordEquationNumbering.includes("WordOmmlFormulaStore.GetEquationRange"));
assert.ok(wordOmmlConverter.includes("MML2OMML.XSL"));
assert.ok(wordOmmlConverter.includes("FormattedText"));
assert.ok(wordOmmlConverter.includes("WdOMathType.wdOMathDisplay"));
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
assert.ok(installOle.includes('LoadBehavior\" -PropertyType DWord -Value 0'));
assert.ok(installVsto.includes("uninstall_windows_ole.ps1"));
assert.ok(installVsto.includes('"/L*v"'));
assert.ok(installVsto.includes("RelatedProducts"));
assert.ok(installVsto.includes("Get-FileHash"));
assert.ok(installVsto.includes("Wait-ForRelatedProductCount"));
assert.ok(installVsto.includes("product state did not settle"));
assert.ok(installVsto.includes("Assert-NativeOleRegistration"));
assert.ok(installVsto.includes("ServerExecutable"));
assert.ok(installVsto.includes("3,1,32,1"));
assert.ok(installVsto.includes("VisualTeX.FormulaOleServer.exe"));
assert.ok(installVsto.includes('Value "native-vsto-ole"'));
assert.ok(installVsto.includes('Name "NativeOleEnabled"'));
assert.ok(installVsto.includes("hashManifest.dependencies"));
assert.ok(installVsto.includes("Assert-NoOfficeProcesses"));
assert.ok(installVsto.includes('Get-Process WINWORD, POWERPNT'));
assert.ok(installVsto.includes('"MSIRESTARTMANAGERCONTROL=Disable"'));
assert.ok(installVsto.includes('"REBOOT=ReallySuppress"'));
assert.ok(installVsto.includes('"vsto-bootstrap-$bootstrapStamp.log"'));
assert.ok(uninstallVsto.includes("DF66EC66-3B3A-4675-A7BE-30456A04EB96"));
assert.ok(uninstallVsto.includes('Name "NativeOleEnabled"'));
assert.ok(buildWindowsOffice.includes("test_windows_formula_ole_server.ps1"));
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
assert.ok(dependencyLoadingSmoke.includes("System.Text.Json, Version=8.0.0.0"));
assert.ok(nativeMsi.includes("3,1,32,1"));
assert.ok(dependencyLoadingSmoke.includes("System.Numerics.Vectors, Version=4.1.4.0"));
assert.ok(dependencyLoadingSmoke.includes("SerializeJson"));
assert.ok(dependencyLoadingSmoke.includes("SysWOW64"));
assert.ok(buildWindowsOffice.includes('StartsWith("8.")'));
assert.ok(buildWindowsOffice.includes('$msbuild $installerProject'));
assert.ok(buildWindowsOffice.includes('PackagePlatform = "x64"'));
assert.ok(buildWindowsOffice.includes('PackagePlatform = "x86"'));
assert.ok(buildWindowsOffice.includes('OlePlatform = "Win32"'));
assert.ok(buildWindowsOffice.includes("VisualTeX-WindowsOffice-VSTO-$packagePlatform.msi"));
assert.ok(installVsto.includes("Resolve-OfficePlatform"));
assert.ok(installVsto.includes("RegistryView]::Registry32"));
assert.ok(installVsto.includes("RegistryView]::Registry64"));
assert.ok(installVsto.includes("PackageDirectory"));
assert.ok(installVsto.includes("MSI SHA-256 mismatch"));
for (const msiRequirement of [
  "FormulaOleServerExecutable",
  "LocalServer32",
  "ServerExecutable",
  "InprocHandler32",
  "Ole32.dll",
  "VisualTeX.Formula.1",
  "ProxyStubClsid32",
  "DF66EC66-3B3A-4675-A7BE-30456A04EB96",
  "native-vsto-ole",
  "NativeOleEnabled",
  "SystemNumericsVectors",
  "SystemValueTuple",
]) {
  assert.ok(nativeMsi.includes(msiRequirement), `Native MSI is missing ${msiRequirement}`);
}
assert.ok(!nativeMsi.includes("CustomAction"));
assert.ok(installOle.includes("Read-AndValidateManifest"));
assert.ok(installOle.includes("Schannel HTTP 200"));
assert.ok(installOle.includes("Clear-VisualTeXWefCache"));
assert.ok(installOle.includes("Test-VisualTeXOnlyRibbonCache"));
assert.ok(installOle.includes("Preserved shared Office Ribbon cache"));
assert.ok(installOle.includes("$exitDeadline = [DateTimeOffset]::UtcNow.AddSeconds(8)"));
assert.ok(installOle.includes("Get-Process $processName -ErrorAction SilentlyContinue"));
assert.ok(installOle.includes("Stop-Process -Force -ErrorAction SilentlyContinue"));

const platformBundle = await source("scripts/build_platform_bundle.mjs");
const windowsBundle = await source("src-tauri/tauri.windows.conf.json");
const installerHooks = await source("src-tauri/windows/hooks.nsh");
assert.ok(platformBundle.includes('"scripts/build_windows_office.ps1"'));
assert.ok(platformBundle.includes('"-SkipTests"'));
assert.ok(!platformBundle.includes('"scripts/build_windows_ole_bridge.ps1"'));
assert.ok(windowsBundle.includes('"../scripts/install_windows_vsto.ps1"'));
assert.ok(!windowsBundle.includes('"../scripts/install_windows_ole.ps1"'));
for (const bundledOfficeResource of [
  "VisualTeX-WindowsOffice-VSTO-x64.msi",
  "VisualTeX-WindowsOffice-VSTO-x64.sha256.json",
  "VisualTeX-WindowsOffice-VSTO-x86.msi",
  "VisualTeX-WindowsOffice-VSTO-x86.sha256.json",
]) {
  assert.ok(windowsBundle.includes(bundledOfficeResource));
}
assert.ok(installerHooks.includes("${NSD_Check} $VisualTeXOfficeOleRadio"));
assert.ok(installerHooks.includes('${If} $VisualTeXOfficeChoice == ""'));
assert.ok(installerHooks.includes('StrCpy $VisualTeXOfficeChoice "ole"'));
assert.ok(installerHooks.includes("install_windows_vsto.ps1"));
assert.ok(installerHooks.includes("VisualTeX-WindowsOffice-VSTO-x64.msi"));
assert.ok(installerHooks.includes("VisualTeX-WindowsOffice-VSTO-x86.msi"));
assert.ok(installerHooks.includes("VisualTeX-WindowsOffice-VSTO-x64.sha256.json"));
assert.ok(installerHooks.includes("VisualTeX-WindowsOffice-VSTO-x86.sha256.json"));
assert.ok(installerHooks.includes('-PackageDirectory "$INSTDIR\\windows-office"'));
assert.ok(installerHooks.includes("uninstall_windows_vsto.ps1"));
assert.ok(installerHooks.includes("Get-Process WINWORD,POWERPNT"));
assert.ok(installerHooks.includes("Stop-Process -Force"));
assert.ok(installerHooks.includes("IDYES visualtex_force_close_office"));
assert.ok(installerHooks.includes("未保存的 Office 文档可能丢失"));
assert.ok(installerHooks.includes("选择“否”将返回上一页"));
assert.ok(installerHooks.indexOf("IDYES visualtex_force_close_office") < installerHooks.indexOf("Stop-Process -Force"));
assert.ok(!installerHooks.includes("VisualTeXOfficeVstoRadio"));
assert.ok(!installerHooks.includes('VisualTeXOfficeChoice == "vsto"'));
assert.ok(
  !installerHooks.includes('-File "$INSTDIR\\scripts\\install_windows_ole.ps1"'),
);
assert.ok(!installerHooks.includes("Office Add-ins dialogs"));
assert.ok(!installerHooks.includes("Automatically configuring Word and PowerPoint"));

const windowsEntry = await source("src/office/windows-ole/main.ts");
const windowsAdapter = await source("src/office/windows-ole/WindowsOleAdapter.ts");
assert.ok(!windowsEntry.includes("../macos"));
assert.ok(windowsEntry.includes("lastDoubleClickAt"));
assert.ok(windowsEntry.includes("bridge.prepareInteractionTarget"));
assert.ok(windowsEntry.includes("payload.metadata"));
assert.ok(windowsAdapter.includes("pendingInteractionTarget"));
assert.ok(windowsAdapter.includes("capturedTarget.metadata"));
assert.ok(windowsEntry.includes('bridge.run("create", () => event?.completed?.())'));
assert.ok(windowsEntry.includes('bridge.run("edit", () => event?.completed?.())'));
assert.ok(!windowsEntry.includes('bridge.run("create").finally'));
assert.ok(!windowsEntry.includes('bridge.run("edit").finally'));

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
