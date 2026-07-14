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
assert.ok(contracts.includes("netstandard2.0"));
assert.ok(oleProject.includes("net8.0-windows"));
assert.ok(oleProject.includes("win-x64"));
assert.ok(oleProject.includes("PublishSingleFile>true"));
assert.ok(oleProject.includes("SelfContained>true"));
assert.ok(wordProject.includes("net48"));
assert.ok(powerpointProject.includes("net48"));
assert.ok(wordProject.includes("<PlatformTarget>x64</PlatformTarget>"));
assert.ok(powerpointProject.includes("<PlatformTarget>x64</PlatformTarget>"));

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
const powerpointVsto = await source(
  "src-windows/VisualTeX.PowerPointVsto/ThisAddIn.cs",
);
assert.ok(!sessionClient.includes("_installToken = ReadInstallToken()"));
assert.ok(sessionClient.includes("StartVisualTeXCompanion"));
assert.ok(sessionClient.includes("timeout.CancelAfter(TimeSpan.FromSeconds(2))"));
assert.ok(wordVsto.includes("活动 Word 文档已切换，未写入公式"));
assert.ok(powerpointVsto.includes("活动演示文稿已切换，未写入公式"));
assert.ok(wordVsto.includes("dispatcher.InvokeAsync"));
assert.ok(powerpointVsto.includes("dispatcher.InvokeAsync"));
assert.ok(wordVsto.includes("IDTExtensibility2, Office.IRibbonExtensibility"));
assert.ok(powerpointVsto.includes("IDTExtensibility2, Office.IRibbonExtensibility"));
assert.ok(wordVsto.includes("ClassInterfaceType.None"));
assert.ok(powerpointVsto.includes("ClassInterfaceType.None"));

const comContracts = await source(
  "src-windows/VisualTeX.WindowsOffice.Contracts/OfficeComInterfaces.cs",
);
assert.ok(!comContracts.includes("interface IOfficeComAddIn"));
assert.ok(!comContracts.includes("interface IOfficeRibbonExtensibility"));

const installOle = await source("scripts/install_windows_ole.ps1");
const installVsto = await source("scripts/install_windows_vsto.ps1");
assert.ok(installOle.includes('LoadBehavior\" -PropertyType DWord -Value 0'));
assert.ok(installVsto.includes("uninstall_windows_ole.ps1"));
assert.ok(installVsto.includes('"/L*v"'));
assert.ok(installVsto.includes("RelatedProducts"));
assert.ok(installVsto.includes("Get-FileHash"));
assert.ok(installVsto.includes("Expected exactly one VisualTeX MSI product"));
assert.ok(installOle.includes("Read-AndValidateManifest"));
assert.ok(installOle.includes("Schannel HTTP 200"));
assert.ok(installOle.includes("Clear-VisualTeXWefCache"));
assert.ok(installOle.includes("Test-VisualTeXOnlyRibbonCache"));
assert.ok(installOle.includes("Preserved shared Office Ribbon cache"));

const platformBundle = await source("scripts/build_platform_bundle.mjs");
const windowsBundle = await source("src-tauri/tauri.windows.conf.json");
const installerHooks = await source("src-tauri/windows/hooks.nsh");
assert.ok(platformBundle.includes('"scripts/build_windows_ole_bridge.ps1"'));
assert.ok(!platformBundle.includes('"scripts/build_windows_office.ps1"'));
assert.ok(!windowsBundle.includes('"../scripts/install_windows_vsto.ps1"'));
assert.ok(!windowsBundle.includes("VisualTeX-WindowsOffice-VSTO.msi"));
assert.ok(installerHooks.includes("${NSD_Check} $VisualTeXOfficeOleRadio"));
assert.ok(installerHooks.includes("uninstall_windows_vsto.ps1"));
assert.ok(!installerHooks.includes("VisualTeXOfficeVstoRadio"));
assert.ok(!installerHooks.includes('VisualTeXOfficeChoice == "vsto"'));

const windowsEntry = await source("src/office/windows-ole/main.ts");
assert.ok(!windowsEntry.includes("../macos"));
assert.ok(windowsEntry.includes("lastDoubleClickAt"));
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
  ])
).join("\n");
for (const requirement of [
  "PipeTokenComparisonRejectsMismatch",
  "AllOfficeWorkRunsOnOneStaThread",
  "FailedConfigurationKeepsOriginalAndDeletesCandidate",
  "FormulaMetadataRoundTripsWithPersistentUuid",
  "FinalReleaseInvalidatesARealComObject",
  "FormulaDoubleClickDeduplicatesOnlyTheSamePersistentTarget",
]) {
  assert.ok(tests.includes(requirement), `Windows test coverage is missing ${requirement}`);
}

console.log("Windows Office architecture smoke test passed");
