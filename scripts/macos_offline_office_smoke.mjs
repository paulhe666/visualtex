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
  "office/macos-offline/word/customUI14.xml",
  "office/macos-offline/word/VisualTeXWord.scpt",
  "office/macos-offline/powerpoint/VTRibbonCallbacks.bas",
  "office/macos-offline/powerpoint/VTPowerPointAdapter.bas",
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
const powerpointAdapter = read("office/macos-offline/powerpoint/VTPowerPointAdapter.bas");
const protocol = read("office/macos-offline/shared/VTProtocol.bas");
const officePaths = read("office/macos-offline/shared/VTOfficePaths.bas");
const launcher = read("office/macos-offline/shared/VTLauncher.bas");
const wordScript = read("office/macos-offline/word/VisualTeXWord.scpt");
const powerpointScript = read("office/macos-offline/powerpoint/VisualTeXPowerPoint.scpt");
const rustRuntime = read("src-tauri/src/office/macos_offline.rs");
const installer = read("src-tauri/src/office/macos_offline_installer.rs");
const packager = read("scripts/package_macos_offline_addins.mjs");
const nativeHtml = read("office-native-dialog.html");
const nativeMain = read("src/office/native-dialog-main.tsx");
const infoPlist = read("src-tauri/Info.macos.plist");

for (const callback of [
  "VTWordRibbonInline",
  "VTWordRibbonDisplay",
  "VTWordRibbonEdit",
  "VTWordRibbonNumbering",
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

expectIncludes(wordAdapter, "Public Sub AutoExec()", "Word template must publish AutoExec health");
expectIncludes(wordAdapter, "VisualTeX_ApplyPendingResult", "Word template must expose the native result callback");
expectIncludes(wordAdapter, "InlineShapes.AddPicture", "Word formula insertion must create an InlineShape");
expectIncludes(wordAdapter, "target.Delete", "Word replacement must delete the old object only after candidate setup");
expectIncludes(wordAdapter, "Word did not persist the VisualTeX formula properties", "Word must verify the candidate before deleting the old formula");
expectIncludes(wordAdapter, "transactionErrorNumber = Err.Number", "Word rollback must preserve the original transaction error");
expectIncludes(wordAdapter, "errorNumber = Err.Number", "Word creation cleanup must preserve the original error number");
expectIncludes(wordAdapter, "VTShowError \"Word formula creation\", errorNumber, errorDescription", "Word creation errors must survive placeholder cleanup");
expectIncludes(wordAdapter, "If Not insertedNumber Is Nothing Then insertedNumber.Delete", "Word rollback must remove a partially inserted equation number");
expectIncludes(wordAdapter, "VTFindCommittedInlineShape", "Word retries must recognize an already committed Session result");
expectIncludes(wordAdapter, "sourceDocumentId <> VTWordDocumentIdentity()", "Word callback must reject document switching");
expectIncludes(wordAdapter, "Private Function VTWordBookmarkName", "Word pending Bookmarks must use one bounded name generator");
expectIncludes(wordAdapter, "Len(VTWordBookmarkName) > 40", "Word Bookmark names must be guarded by the host length limit");
expectIncludes(powerpointAdapter, "Public Sub Auto_Open()", "PowerPoint add-in must publish Auto_Open health");
expectIncludes(powerpointAdapter, "VisualTeXFormulaId", "PowerPoint add-in must persist formulaId tags");
expectIncludes(powerpointAdapter, "VisualTeXSessionId", "PowerPoint add-in must persist sessionId tags");
expectIncludes(powerpointAdapter, "VisualTeXPending", "PowerPoint add-in must persist pending tags");
expectIncludes(powerpointAdapter, "original.Delete", "PowerPoint replacement must delete the old shape last");
expectIncludes(powerpointAdapter, "candidate.ZOrderPosition <> targetZOrder + 1", "PowerPoint must verify z-order before deleting the old shape");
expectIncludes(powerpointAdapter, 'candidate.Tags("VisualTeXSessionId") <> sessionId', "PowerPoint must verify durable Session tags before deleting the old shape");
expectIncludes(powerpointAdapter, "VTIsCommittedPowerPointShape", "PowerPoint retries must recognize an already committed Session result");
expectIncludes(powerpointAdapter, "VTRestoreZOrder candidate, targetZOrder + 1", "PowerPoint replacement must preserve z-order transactionally");
expect(!wordAdapter.includes('Format$(Now, "yyyy-mm-dd\\Thh:nn:ss") & "Z"'), "Word health must not label local time as UTC");
expect(!powerpointAdapter.includes('Format$(Now, "yyyy-mm-dd\\Thh:nn:ss") & "Z"'), "PowerPoint health must not label local time as UTC");

expectIncludes(officePaths, "UBF8T346G9.Office/VisualTeX", "VBA paths must use the Office application-group container");
expectIncludes(officePaths, 'InStr(1, homePath, "/Library/Containers/", vbTextCompare)', "VBA paths must detect an Office sandbox HOME value");
expectIncludes(officePaths, "homePath = Left$(homePath, sandboxMarker - 1)", "VBA paths must recover the real user home from an Office sandbox HOME value");
expect(!officePaths.includes("Library/Application Support/VisualTeX"), "VBA paths must not use a sandbox-inaccessible user Application Support root");
expectIncludes(protocol, "New Collection", "VBA protocol must use the Mac-compatible Collection type");
expect(!protocol.includes("Scripting.Dictionary"), "VBA protocol must not depend on Windows Scripting Runtime");
expectIncludes(protocol, "If Not VT_RANDOM_READY Then", "UUID generation must seed VBA randomness only once per host process");
expectIncludes(protocol, "VT_UUID_COUNTER", "UUID generation must mix a monotonic per-process counter");
expect(!protocol.includes("LenB(StrConv(json, vbFromUnicode))"), "Request sizing must use UTF-8 bytes instead of the host code page");
expectIncludes(protocol, "Open temporary For Binary Access Write", "VBA request writes must be binary UTF-8 writes");
expectIncludes(protocol, "VTUtf8Encode", "VBA protocol must provide strict UTF-8 encoding");
expectIncludes(protocol, "VTUtf8Decode", "VBA protocol must provide strict UTF-8 decoding");
expectIncludes(protocol, 'If Dir$(sourcePath) = "" Then', "VBA file reads must use the Office for Mac-compatible Dir$ form");
expectIncludes(protocol, "Office for Mac can return an empty Dir$ result", "VBA directory creation must document the Office sandbox behavior");
expectIncludes(protocol, "On Error Resume Next\n    MkDir directoryPath\n    On Error GoTo 0", "VBA directory creation must tolerate sandbox-authorized existing directories");
expectIncludes(protocol, 'VTPathFileExists = (Dir$(value) <> "")', "VBA file existence checks must use the Office for Mac-compatible Dir$ form");
expectIncludes(protocol, "Public Function VTProtocolSelfTest() As Boolean", "VBA protocol must expose an actual host-runtime UUID/UTF-8 self-test");
expectIncludes(launcher, "AppleScriptTask", "VBA launcher must use AppleScriptTask");
expectIncludes(wordScript, "/usr/bin/open ", "Word AppleScriptTask must launch only the fixed open tool");
expectIncludes(powerpointScript, "/usr/bin/open ", "PowerPoint AppleScriptTask must launch only the fixed open tool");
expectIncludes(wordScript, "quoted form of visualTeXURL", "Word AppleScriptTask must quote the URL");
expectIncludes(powerpointScript, "quoted form of visualTeXURL", "PowerPoint AppleScriptTask must quote the URL");
expect(!wordScript.includes("System Events"), "Word AppleScriptTask must not use UI automation");
expect(!powerpointScript.includes("System Events"), "PowerPoint AppleScriptTask must not use UI automation");
expect(!wordScript.match(/\brm\b|delete file|sh -c/i), "Word AppleScriptTask must not expose file deletion or arbitrary shell execution");
expect(!powerpointScript.match(/\brm\b|delete file|sh -c/i), "PowerPoint AppleScriptTask must not expose file deletion or arbitrary shell execution");

const offlineRuntimeSources = [
  wordAdapter,
  powerpointAdapter,
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
expectIncludes(rustRuntime, "cleanup_session_files_at", "Completed and cancelled Sessions must remove known local request artifacts");
expectIncludes(rustRuntime, "DirectoryNotEmpty", "Session cleanup must preserve unknown files instead of deleting an entire directory recursively");
expectIncludes(installer, "VisualTeX.dotm", "Installer must preserve the fixed Word filename");
expectIncludes(installer, "VisualTeX.ppam", "Installer must preserve the fixed PowerPoint filename");
expectIncludes(installer, 'home.join("Applications")', "Installer must detect per-user Office application installs");
expectIncludes(installer, "restore_backups(&backups)", "Installer must roll back every staged file after a partial failure");
expectIncludes(installer, 'remove_if_exists(&word_health)', "Installer must clear stale Word health after an update");
expectIncludes(installer, 'remove_if_exists(&powerpoint_health)', "Installer must clear stale PowerPoint health after an update");
expectIncludes(installer, "health_is_current", "Installer must validate exact host/version health records");
expectIncludes(installer, "Err(error) => (false, None, Some(error))", "A corrupt health file must degrade one host instead of failing the whole status view");
expectIncludes(installer, "word_paths.extend([word_script.clone(), placeholder.clone()])", "Word installed status must include its AppleScriptTask and placeholder resources");
expectIncludes(installer, "powerpoint_script.clone()", "PowerPoint installed status must include its AppleScriptTask resource");
expectIncludes(installer, 'health.plugin_version.as_deref() == Some(env!("CARGO_PKG_VERSION"))', "Installer must reject stale plug-in health versions");
expectIncludes(installer, "Library/Application Scripts/com.microsoft.Word", "Installer must use Word's AppleScriptTask directory");
expectIncludes(installer, "Library/Application Scripts/com.microsoft.Powerpoint", "Installer must use PowerPoint's AppleScriptTask directory");
expectIncludes(installer, "addins.json", "Installer must require the compiled add-in checksum manifest");
expectIncludes(installer, "word/vbaProject.bin", "Installer must validate the Word VBA project entry");
expectIncludes(installer, "ppt/vbaProject.bin", "Installer must validate the PowerPoint VBA project entry");
expectIncludes(installer, 'validate_vba_module(path, expected_vba_entry, "VTOfficePaths")', "Installer must reject stale add-ins that predate the Office shared-container path module");
expectIncludes(packager, "expectedModules", "Packager must verify the reviewed VBA module names");
expectIncludes(packager, '"VTOfficePaths"', "Packager must require the Office shared-container path module");
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
