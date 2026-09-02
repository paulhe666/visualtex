import { execFileSync, spawnSync } from "node:child_process";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  renameSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { homedir } from "node:os";
import { basename, dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

if (process.platform !== "darwin") {
  throw new Error("The Word VBE builder is available only on macOS.");
}

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const argument = (name) => {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
};
const basePath = resolve(
  argument("--base") ??
    join(repositoryRoot, "office", "macos-offline", "resources", "VisualTeX.dotm"),
);
const scratchRoot = join(
  homedir(),
  "Library",
  "Group Containers",
  "UBF8T346G9.Office",
  "VisualTeX",
  "Scratch",
);
const outputPath = resolve(
  argument("--output") ?? join(scratchRoot, "VisualTeXWordBuild.dotm"),
);
const outputDocumentName = basename(outputPath);
const keepWordOpenOnError = process.argv.includes("--keep-word-open-on-error");
const preserveWord = process.argv.includes("--preserve-word");
const buildLockRoot = join(scratchRoot, "VisualTeXWordBuild.lock");
const buildLockOwnerPath = join(buildLockRoot, "pid");
const offlineOfficeRoot = join(repositoryRoot, "office", "macos-offline");
const wordModuleSources = [
  ["VTProtocol", join(offlineOfficeRoot, "shared", "VTProtocol.bas")],
  ["VTOfficePaths", join(offlineOfficeRoot, "shared", "VTOfficePaths.bas")],
  ["VTMetadata", join(offlineOfficeRoot, "shared", "VTMetadata.bas")],
  ["VTLauncher", join(offlineOfficeRoot, "shared", "VTLauncher.bas")],
  ["VTErrorHandling", join(offlineOfficeRoot, "shared", "VTErrorHandling.bas")],
  ["VTWordAdapter", join(offlineOfficeRoot, "word", "VTWordAdapter.bas")],
  ["VTWordEvents", join(offlineOfficeRoot, "word", "VTWordEvents.cls")],
  ["VTRibbonCallbacks", join(offlineOfficeRoot, "word", "VTRibbonCallbacks.bas")],
];
const requestedModuleNames = new Set(
  (argument("--modules") ?? "")
    .split(",")
    .map((value) => value.trim())
    .filter(Boolean),
);
const knownModuleNames = new Set(wordModuleSources.map(([moduleName]) => moduleName));
for (const moduleName of requestedModuleNames) {
  if (!knownModuleNames.has(moduleName)) {
    throw new Error(`Unknown Word VBA module requested by --modules: ${moduleName}`);
  }
}
const incrementalBuild = requestedModuleNames.size > 0;
// Preserve mode may use an incremental copy of the reviewed DOTM. The
// production Startup add-in is temporarily unloaded while that isolated copy is
// edited, then restored before the user's Word document snapshot is checked.
if (preserveWord && outputPath === basePath) {
  throw new Error("--preserve-word requires an isolated --output path");
}
const selectedWordModuleSources = incrementalBuild
  ? wordModuleSources.filter(([moduleName]) => requestedModuleNames.has(moduleName))
  : wordModuleSources;
const officeGroupRoot = join(
  homedir(),
  "Library",
  "Group Containers",
  "UBF8T346G9.Office",
);
const startupRoot = join(
  officeGroupRoot,
  "User Content.localized",
  "Startup.localized",
  "Word",
);
const normalTemplatePath = join(
  officeGroupRoot,
  "User Content.localized",
  "Templates.localized",
  "Normal.dotm",
);
const normalTemplateBackupPath = join(
  scratchRoot,
  "VbeBuildNormalBackup.dotm",
);
const backupRoot = join(scratchRoot, `VbeBuildStartupBackup-${process.pid}`);
const documentName = basename(outputPath).replace(/\.dotm$/i, "");

function run(program, args, options = {}) {
  const encoding = Object.prototype.hasOwnProperty.call(options, "encoding")
    ? options.encoding
    : "utf8";
  return execFileSync(program, args, {
    encoding: encoding === "buffer" ? null : encoding,
    input: options.input,
    stdio: options.stdio ?? ["ignore", "pipe", "pipe"],
    timeout: options.timeout ?? 60_000,
    maxBuffer: 32 * 1024 * 1024,
  });
}

function bestEffort(program, args, options = {}) {
  try {
    return run(program, args, options);
  } catch {
    return "";
  }
}

function processIsAlive(pid) {
  if (!Number.isSafeInteger(pid) || pid <= 0) return false;
  try {
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

function acquireBuildLock() {
  mkdirSync(scratchRoot, { recursive: true });
  try {
    mkdirSync(buildLockRoot);
  } catch (error) {
    if (!(error instanceof Error && "code" in error && error.code === "EEXIST")) {
      throw error;
    }
    const ownerPid = Number.parseInt(
      bestEffort("/bin/cat", [buildLockOwnerPath]).trim(),
      10,
    );
    const ownerCommand = bestEffort("/bin/ps", [
      "-p",
      String(ownerPid),
      "-o",
      "command=",
    ]).trim();
    if (
      ownerPid !== process.pid &&
      processIsAlive(ownerPid) &&
      ownerCommand.includes("rebuild_macos_word_addin.mjs")
    ) {
      throw new Error(
        `Another Word VBE build is already running (pid ${ownerPid}).`,
      );
    }
    rmSync(buildLockRoot, { recursive: true, force: true });
    mkdirSync(buildLockRoot);
  }
  writeFileSync(buildLockOwnerPath, `${process.pid}\n`, "utf8");
}

function releaseBuildLock() {
  const ownerPid = Number.parseInt(
    bestEffort("/bin/cat", [buildLockOwnerPath]).trim(),
    10,
  );
  if (ownerPid === process.pid) {
    rmSync(buildLockRoot, { recursive: true, force: true });
  }
}

function osascript(lines, timeout = 60_000) {
  return run(
    "/usr/bin/osascript",
    lines.flatMap((line) => ["-e", line]),
    { timeout },
  ).trim();
}

function sleep(milliseconds) {
  spawnSync("/bin/sleep", [String(milliseconds / 1000)], { stdio: "ignore" });
}

function readVbaTrust() {
  try {
    return {
      existed: true,
      enabled:
        run("/usr/bin/defaults", [
          "read",
          "com.microsoft.Word",
          "VBAObjectModelIsTrusted",
        ]).trim() === "1",
    };
  } catch {
    return { existed: false, enabled: false };
  }
}

function setVbaTrust(enabled) {
  run("/usr/bin/defaults", [
    "write",
    "com.microsoft.Word",
    "VBAObjectModelIsTrusted",
    "-bool",
    enabled ? "true" : "false",
  ]);
  bestEffort("/usr/bin/killall", ["cfprefsd"]);
  sleep(800);
}

function restoreVbaTrust(state) {
  if (state.existed) {
    setVbaTrust(state.enabled);
  } else {
    bestEffort("/usr/bin/defaults", [
      "delete",
      "com.microsoft.Word",
      "VBAObjectModelIsTrusted",
    ]);
    bestEffort("/usr/bin/killall", ["cfprefsd"]);
  }
}

function recoverInterruptedNormalTemplate() {
  if (!existsSync(normalTemplateBackupPath)) return;
  if (existsSync(normalTemplatePath)) {
    throw new Error(
      `A previous Word VBE build left both the real and backup Normal.dotm in place. Inspect ${normalTemplateBackupPath} before continuing.`,
    );
  }
  mkdirSync(dirname(normalTemplatePath), { recursive: true });
  renameSync(normalTemplateBackupPath, normalTemplatePath);
}

function moveNormalTemplateOut() {
  recoverInterruptedNormalTemplate();
  if (!existsSync(normalTemplatePath)) return;
  mkdirSync(dirname(normalTemplateBackupPath), { recursive: true });
  renameSync(normalTemplatePath, normalTemplateBackupPath);
}

function restoreNormalTemplate() {
  if (!existsSync(normalTemplateBackupPath)) return;
  rmSync(normalTemplatePath, { force: true });
  mkdirSync(dirname(normalTemplatePath), { recursive: true });
  renameSync(normalTemplateBackupPath, normalTemplatePath);
}

function moveStartupTemplatesOut() {
  mkdirSync(startupRoot, { recursive: true });
  mkdirSync(backupRoot, { recursive: true });
  for (const name of readdirSync(startupRoot)) {
    if (/^VisualTeX\.dotm/i.test(name)) {
      renameSync(join(startupRoot, name), join(backupRoot, name));
      continue;
    }
    if (/^~\$.*sualTeX.*\.dotm$/i.test(name)) {
      rmSync(join(startupRoot, name), { force: true });
    }
  }
}

function restoreStartupTemplates() {
  if (!existsSync(backupRoot)) return;
  for (const name of readdirSync(backupRoot)) {
    const destination = join(startupRoot, name);
    rmSync(destination, { force: true });
    renameSync(join(backupRoot, name), destination);
  }
  rmSync(backupRoot, { recursive: true, force: true });
}

function closeWordWithoutSaving() {
  bestEffort("/usr/bin/osascript", [
    "-e",
    'tell application "Microsoft Word" to quit saving no',
  ], { timeout: 20_000 });
  sleep(1_500);
  const pids = bestEffort("/usr/bin/pgrep", ["-x", "Microsoft Word"])
    .trim()
    .split(/\s+/)
    .filter(Boolean);
  for (const pid of pids) bestEffort("/bin/kill", ["-9", pid]);
  sleep(1_500);
}

function wordVisualTeXAddinInstalled() {
  const value = bestEffort("/usr/bin/osascript", [
    "-e",
    'tell application "Microsoft Word"',
    "-e",
    'if not (exists add in "VisualTeX.dotm") then return "MISSING"',
    "-e",
    'return installed of add in "VisualTeX.dotm" as text',
    "-e",
    "end tell",
  ], { timeout: 20_000 }).trim();
  if (value === "true") return true;
  if (value === "false") return false;
  return null;
}

function setWordVisualTeXAddinInstalled(installed) {
  if (installed === null) return;
  run("/usr/bin/osascript", [
    "-e",
    'tell application "Microsoft Word"',
    "-e",
    'if not (exists add in "VisualTeX.dotm") then error "The production VisualTeX add-in is missing"',
    "-e",
    `set installed of add in "VisualTeX.dotm" to ${installed ? "true" : "false"}`,
    "-e",
    "end tell",
  ], { timeout: 20_000 });
}

function wordDocumentSnapshot() {
  const raw = osascript([
    'tell application "Microsoft Word"',
    'set output to ""',
    'set documentCount to count of documents',
    'repeat with documentIndex from 1 to documentCount',
    'set documentObject to document documentIndex',
    'set fullNameText to ""',
    'try',
    'set fullNameText to full name of documentObject as text',
    'end try',
    'set output to output & (name of documentObject as text) & (ASCII character 31) & (saved of documentObject as text) & (ASCII character 31) & fullNameText',
    'if documentIndex is less than documentCount then set output to output & linefeed',
    'end repeat',
    'return output',
    'end tell',
  ], 30_000);
  if (!raw.trim()) return [];
  return raw
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => {
      const [name, saved, fullName = ""] = line.split("\x1f");
      return { name, saved: saved === "true", fullName };
    })
    .sort((left, right) =>
      `${left.name}\x1f${left.fullName}`.localeCompare(
        `${right.name}\x1f${right.fullName}`,
      ),
    );
}

function dismissVbeFileDialogIfOpen() {
  bestEffort("/usr/bin/osascript", [
    "-e",
    'tell application "System Events"',
    "-e",
    'if not (exists process "Microsoft Word") then return "NO_WORD"',
    "-e",
    'tell process "Microsoft Word"',
    "-e",
    'repeat with candidateWindow in windows',
    "-e",
    "try",
    "-e",
    'set candidateName to name of candidateWindow as text',
    "-e",
    'if candidateName starts with "导入文件" or candidateName starts with "Import File" then',
    "-e",
    'perform action "AXRaise" of candidateWindow',
    "-e",
    "delay 0.2",
    "-e",
    "key code 53",
    "-e",
    "delay 0.5",
    "-e",
    'return "CLOSED"',
    "-e",
    "end if",
    "-e",
    "end try",
    "-e",
    "end repeat",
    "-e",
    'return "NO_DIALOG"',
    "-e",
    "end tell",
    "-e",
    "end tell",
  ], { timeout: 15_000 });
}

function dismissVbeWarningDialogsIfOpen() {
  bestEffort("/usr/bin/osascript", [
    "-e",
    'tell application "System Events"',
    "-e",
    'if not (exists process "Microsoft Word") then return "NO_WORD"',
    "-e",
    'tell process "Microsoft Word"',
    "-e",
    "set frontmost to true",
    "-e",
    "repeat 12 times",
    "-e",
    "set warningWindow to missing value",
    "-e",
    "repeat with candidateWindow in windows",
    "-e",
    "try",
    "-e",
    'if description of candidateWindow is "警告" or description of candidateWindow is "Warning" then',
    "-e",
    "set warningWindow to candidateWindow",
    "-e",
    "exit repeat",
    "-e",
    "end if",
    "-e",
    "end try",
    "-e",
    "end repeat",
    "-e",
    "if warningWindow is missing value then exit repeat",
    "-e",
    'perform action "AXRaise" of warningWindow',
    "-e",
    'if exists button "确定" of warningWindow then click button "确定" of warningWindow',
    "-e",
    'if exists button "OK" of warningWindow then click button "OK" of warningWindow',
    "-e",
    "delay 0.25",
    "-e",
    "end repeat",
    "-e",
    'return "DONE"',
    "-e",
    "end tell",
    "-e",
    "end tell",
  ], { timeout: 20_000 });
}

function closeVbeWindowIfOpen() {
  bestEffort("/usr/bin/osascript", [
    "-e",
    'tell application "System Events"',
    "-e",
    'if not (exists process "Microsoft Word") then return "NO_WORD"',
    "-e",
    'tell process "Microsoft Word"',
    "-e",
    'repeat with candidateWindow in windows',
    "-e",
    "try",
    "-e",
    'if name of candidateWindow contains "Microsoft Visual Basic" then',
    "-e",
    'if exists (first button of candidateWindow whose subrole is "AXCloseButton") then click (first button of candidateWindow whose subrole is "AXCloseButton")',
    "-e",
    "delay 0.5",
    "-e",
    'return "CLOSED"',
    "-e",
    "end if",
    "-e",
    "end try",
    "-e",
    "end repeat",
    "-e",
    'return "NO_VBE"',
    "-e",
    "end tell",
    "-e",
    "end tell",
  ], { timeout: 15_000 });
}

function closeBuildDocumentWithoutSaving() {
  bestEffort("/usr/bin/osascript", [
    "-e",
    'tell application "Microsoft Word"',
    "-e",
    `if exists document ${JSON.stringify(outputDocumentName)} then close (document ${JSON.stringify(outputDocumentName)}) saving no`,
    "-e",
    "end tell",
  ], { timeout: 20_000 });
  sleep(500);
}

function waitForWordUiReady() {
  for (let attempt = 0; attempt < 60; attempt += 1) {
    const state = bestEffort("/usr/bin/osascript", [
      "-e",
      'tell application "System Events"',
      "-e",
      'if exists process "Microsoft Word" then',
      "-e",
      'tell process "Microsoft Word"',
      "-e",
      "set visible to true",
      "-e",
      "set frontmost to true",
      "-e",
      'if (count of menu bars) > 0 and (count of windows) > 0 then return "READY"',
      "-e",
      "end tell",
      "-e",
      "end if",
      "-e",
      "end tell",
    ], { timeout: 5_000 }).trim();
    if (state === "READY") return;
    sleep(500);
  }
  throw new Error("Word did not expose a document window and menu bar for VBE automation.");
}

function stopRunningVbaIfNeeded() {
  const state = bestEffort("/usr/bin/osascript", [
    "-e",
    'tell application "System Events"',
    "-e",
    'if not (exists process "Microsoft Word") then return "NO_WORD"',
    "-e",
    'tell process "Microsoft Word"',
    "-e",
    'set vbeWindow to first window whose name contains "Microsoft Visual Basic"',
    "-e",
    'set vbeName to name of vbeWindow as text',
    "-e",
    'if vbeName does not contain "[运行中]" and vbeName does not contain "[running]" then return "READY"',
    "-e",
    'perform action "AXRaise" of vbeWindow',
    "-e",
    'set frontmost to true',
    "-e",
    'key code 47 using {command down}',
    "-e",
    'delay 1',
    "-e",
    'repeat with candidateWindow in windows',
    "-e",
    'repeat with endName in {"结束", "End"}',
    "-e",
    'try',
    "-e",
    'if exists button (endName as text) of candidateWindow then',
    "-e",
    'click button (endName as text) of candidateWindow',
    "-e",
    'delay 1',
    "-e",
    'exit repeat',
    "-e",
    'end if',
    "-e",
    'end try',
    "-e",
    'end repeat',
    "-e",
    'end repeat',
    "-e",
    'return name of vbeWindow as text',
    "-e",
    'end tell',
    "-e",
    'end tell',
  ], { timeout: 15_000 }).trim();
  if (state === "READY" || state === "NO_WORD") return;
  for (let attempt = 0; attempt < 20; attempt += 1) {
    sleep(250);
    const current = bestEffort("/usr/bin/osascript", [
      "-e",
      'tell application "System Events"',
      "-e",
      'tell process "Microsoft Word" to return name of first window whose name contains "Microsoft Visual Basic"',
      "-e",
      'end tell',
    ], { timeout: 5_000 }).trim();
    if (!current.includes("[运行中]") && !current.toLowerCase().includes("[running]")) return;
  }
  throw new Error("Word VBA remained in running state after an explicit stop request.");
}

function openVbeWindow() {
  waitForWordUiReady();
  const existingWindows = bestEffort("/usr/bin/osascript", [
    "-e",
    'tell application "System Events"',
    "-e",
    'if exists process "Microsoft Word" then',
    "-e",
    'tell process "Microsoft Word" to return name of every window',
    "-e",
    "end if",
    "-e",
    "end tell",
  ], { timeout: 5_000 });
  if (existingWindows.includes("Microsoft Visual Basic")) {
    stopRunningVbaIfNeeded();
    return;
  }

  osascript([
    'tell application "Microsoft Word" to activate',
    'tell application "System Events"',
    'tell process "Microsoft Word"',
    "set frontmost to true",
    "set openedEditor to false",
    'repeat with toolsName in {"工具", "Tools"}',
    "if openedEditor is false then",
    "try",
    "set toolsMenu to menu 1 of menu bar item (toolsName as text) of menu bar 1",
    'repeat with macroName in {"宏", "Macro"}',
    "if openedEditor is false then",
    "try",
    "set macroMenu to menu 1 of menu item (macroName as text) of toolsMenu",
    'repeat with editorName in {"Visual Basic 编辑器", "Visual Basic Editor"}',
    "try",
    "click menu item (editorName as text) of macroMenu",
    "set openedEditor to true",
    "exit repeat",
    "end try",
    "end repeat",
    "end try",
    "end if",
    "end repeat",
    "end try",
    "end if",
    "end repeat",
    'if openedEditor is false then error "Unable to find Tools > Macro > Visual Basic Editor in either Chinese or English."',
    "end tell",
    "end tell",
  ]);

  for (let attempt = 0; attempt < 30; attempt += 1) {
    sleep(500);
    const windowNames = bestEffort("/usr/bin/osascript", [
      "-e",
      'tell application "System Events"',
      "-e",
      'if exists process "Microsoft Word" then',
      "-e",
      'tell process "Microsoft Word" to return name of every window',
      "-e",
      "end if",
      "-e",
      "end tell",
    ], { timeout: 5_000 });
    if (windowNames.includes("Microsoft Visual Basic")) {
      stopRunningVbaIfNeeded();
      return;
    }
  }
  throw new Error("Word did not open its Visual Basic Editor window.");
}

function runTransientVbeAutomation(lines, timeout = 60_000) {
  let lastError;
  for (let attempt = 1; attempt <= 3; attempt += 1) {
    try {
      openVbeWindow();
      return osascript(lines, timeout);
    } catch (error) {
      lastError = error;
      if (attempt < 3) sleep(800);
    }
  }
  throw lastError;
}

function openModuleCodeWindow(moduleName) {
  runTransientVbeAutomation([
    'tell application "System Events"',
    'tell process "Microsoft Word"',
    "set frontmost to true",
    'set vbeWindow to first window whose name contains "Microsoft Visual Basic"',
    'perform action "AXRaise" of vbeWindow',
    "set projectOutlineFound to false",
    "try",
    'set projectOutline to first UI element of vbeWindow whose role is "AXOutline"',
    "set projectOutlineFound to true",
    "end try",
    "if projectOutlineFound is false then",
    'repeat with candidateScrollArea in (every UI element of vbeWindow whose role is "AXScrollArea")',
    "try",
    'set projectOutline to first UI element of candidateScrollArea whose role is "AXOutline"',
    "set projectOutlineFound to true",
    "exit repeat",
    "end try",
    "end repeat",
    "end if",
    "if projectOutlineFound is false then",
    'keystroke "r" using {command down}',
    'delay 1',
    'repeat with candidateScrollArea in (every UI element of vbeWindow whose role is "AXScrollArea")',
    "try",
    'set projectOutline to first UI element of candidateScrollArea whose role is "AXOutline"',
    "set projectOutlineFound to true",
    "exit repeat",
    "end try",
    "end repeat",
    "end if",
    'if projectOutlineFound is false then error "Word VBE Project Explorer did not expose a module outline."',
    "set rowIndex to 1",
    "repeat while rowIndex is less than or equal to count of rows of projectOutline",
    "set rowCell to UI element 1 of row rowIndex of projectOutline",
    'set rowNames to ""',
    "repeat with rowElement in UI elements of rowCell",
    "try",
    "set rowName to name of rowElement",
    "if rowName is not missing value then set rowNames to rowNames & (rowName as text)",
    "end try",
    "end repeat",
    "try",
    'set disclosure to first UI element of rowCell whose role is "AXDisclosureTriangle"',
    "if value of disclosure is false then click disclosure",
    "delay 0.3",
    "end try",
    "set rowIndex to rowIndex + 1",
    "end repeat",
    "set moduleRow to 0",
    "repeat with rowIndex from 1 to count of rows of projectOutline",
    "set rowCell to UI element 1 of row rowIndex of projectOutline",
    'set rowNames to ""',
    "repeat with rowElement in UI elements of rowCell",
    "try",
    "set rowName to name of rowElement",
    "if rowName is not missing value then set rowNames to rowNames & (rowName as text)",
    "end try",
    "end repeat",
    `if rowNames contains ${JSON.stringify(moduleName)} then set moduleRow to rowIndex`,
    "end repeat",
    `if moduleRow is 0 then error ${JSON.stringify(`${moduleName} was not found in the Word VBA project`)}`,
    "select row moduleRow of projectOutline",
    "set openedCode to false",
    'repeat with viewName in {"查看", "View"}',
    "if openedCode is false then",
    "try",
    "set viewMenu to menu 1 of menu bar item (viewName as text) of menu bar 1",
    'repeat with codeName in {"代码", "Code"}',
    "try",
    "click menu item (codeName as text) of viewMenu",
    "set openedCode to true",
    "exit repeat",
    "end try",
    "end repeat",
    "end try",
    "end if",
    "end repeat",
    "if openedCode is false then",
    // F7 is VBE's locale-independent View Code command. Some recent Word for
    // Mac builds expose the project row but omit View > Code from AX menus.
    "key code 98",
    "set openedCode to true",
    "end if",
    "delay 0.7",
    "end tell",
    "end tell",
  ]);
}

function removeVbaModule(moduleName) {
  openModuleCodeWindow(moduleName);
  osascript([
    'tell application "System Events"',
    'tell process "Microsoft Word"',
    "set frontmost to true",
    'set vbeWindow to first window whose name contains "Microsoft Visual Basic"',
    'perform action "AXRaise" of vbeWindow',
    "set removedModule to false",
    'repeat with fileName in {"文件", "File"}',
    "if removedModule is false then",
    "try",
    "set fileMenu to menu 1 of menu bar item (fileName as text) of menu bar 1",
    "repeat with candidateItem in menu items of fileMenu",
    "set candidateName to name of candidateItem as text",
    `if candidateName starts with ${JSON.stringify(`删除 ${moduleName}`)} or candidateName starts with ${JSON.stringify(`Remove ${moduleName}`)} then`,
    "click candidateItem",
    "set removedModule to true",
    "exit repeat",
    "end if",
    "end repeat",
    "end try",
    "end if",
    "end repeat",
    `if removedModule is false then error ${JSON.stringify(`Unable to find the Remove ${moduleName} command in either Chinese or English.`)}`,
    "delay 0.8",
    "set dismissedExport to false",
    "repeat with candidateWindow in windows",
    "if dismissedExport is false then",
    "try",
    "repeat with candidateButton in buttons of candidateWindow",
    "set candidateButtonName to name of candidateButton as text",
    'if candidateButtonName starts with "否" or candidateButtonName starts with "No" then',
    "perform action \"AXRaise\" of candidateWindow",
    "click candidateButton",
    "set dismissedExport to true",
    "exit repeat",
    "end if",
    "end repeat",
    "end try",
    "end if",
    "end repeat",
    "delay 0.9",
    "end tell",
    "end tell",
  ]);
}

function replaceVbaModuleSourceText(moduleName, modulePath) {
  const editableSource = readFileSync(modulePath, "utf8")
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .filter((line) => !line.trimStart().startsWith("Attribute "))
    .join("\r");
  if (!editableSource.trim()) {
    throw new Error(`VBA module source is empty: ${modulePath}`);
  }
  openModuleCodeWindow(moduleName);
  const clipboard = spawnSync("/usr/bin/pbcopy", [], {
    input: editableSource,
    encoding: "utf8",
  });
  if (clipboard.status !== 0) {
    throw new Error(
      clipboard.stderr?.trim() || `Unable to stage ${moduleName} source on the clipboard`,
    );
  }
  osascript([
    'tell application "Microsoft Word" to activate',
    'delay 0.3',
    'tell application "System Events"',
    'tell process "Microsoft Word"',
    'set frontmost to true',
    'set vbeWindow to first window whose name contains "Microsoft Visual Basic"',
    'perform action "AXRaise" of vbeWindow',
    'delay 0.2',
    // openModuleCodeWindow() leaves the code pane as the keyboard target. The
    // Mac VBE does not expose that pane as AXTextArea, so use its native editor
    // shortcuts instead of relying on an accessibility role that is absent.
    'keystroke "a" using {command down}',
    'delay 0.1',
    'keystroke "v" using {command down}',
    'delay 2',
    'end tell',
    'end tell',
  ], 60_000);
}

function importVbaModule(modulePath) {
  openVbeWindow();
  osascript([
    'tell application "System Events"',
    'tell process "Microsoft Word"',
    "set frontmost to true",
    'set vbeWindow to first window whose name contains "Microsoft Visual Basic"',
    'perform action "AXRaise" of vbeWindow',
    "set openedImport to false",
    'repeat with fileName in {"文件", "File"}',
    "if openedImport is false then",
    "try",
    "set fileMenu to menu 1 of menu bar item (fileName as text) of menu bar 1",
    "repeat with candidateItem in menu items of fileMenu",
    "set candidateName to name of candidateItem as text",
    'if candidateName starts with "导入文件" or candidateName starts with "Import File" then',
    "click candidateItem",
    "set openedImport to true",
    "exit repeat",
    "end if",
    "end repeat",
    "end try",
    "end if",
    "end repeat",
    'if openedImport is false then error "Unable to find File > Import File in either Chinese or English."',
    "delay 0.4",
    "set importWindow to missing value",
    "repeat with candidateWindow in windows",
    "try",
    "set candidateName to name of candidateWindow as text",
    'if candidateName starts with "导入文件" or candidateName starts with "Import File" then',
    "set importWindow to candidateWindow",
    "exit repeat",
    "end if",
    "end try",
    "end repeat",
    'if importWindow is missing value then error "The VBE Import File dialog did not appear."',
    'perform action "AXRaise" of importWindow',
    "delay 0.3",
    'keystroke "g" using {command down, shift down}',
    "delay 0.6",
    'set pathField to value of attribute "AXFocusedUIElement"',
    `set value of pathField to ${JSON.stringify(modulePath)}`,
    "key code 36",
    "delay 0.9",
    'perform action "AXRaise" of importWindow',
    "key code 36",
    "delay 1.5",
    "set importDialogStillOpen to false",
    "repeat with candidateWindow in windows",
    "try",
    "set candidateName to name of candidateWindow as text",
    'if candidateName starts with "导入文件" or candidateName starts with "Import File" then set importDialogStillOpen to true',
    "end try",
    "end repeat",
    'if importDialogStillOpen then error "The VBE Import File dialog remained open after selecting the module."',
    "end tell",
    "end tell",
  ], 60_000);
}

function compileVbaProject() {
  const compileState = osascript([
    'tell application "System Events"',
    'tell process "Microsoft Word"',
    "set frontmost to true",
    'set vbeWindow to first window whose name contains "Microsoft Visual Basic"',
    'perform action "AXRaise" of vbeWindow',
    "set startedCompile to false",
    'repeat with debugName in {"调试", "Debug"}',
    "if startedCompile is false then",
    "try",
    "set debugMenu to menu 1 of menu bar item (debugName as text) of menu bar 1",
    "repeat with candidateItem in menu items of debugMenu",
    "set candidateName to name of candidateItem as text",
    'if candidateName starts with "编译 " or candidateName starts with "Compile " then',
    "click candidateItem",
    "set startedCompile to true",
    "exit repeat",
    "end if",
    "end repeat",
    "end try",
    "end if",
    "end repeat",
    'if startedCompile is false then error "Unable to find Debug > Compile Project in either Chinese or English."',
    "delay 2",
    "set failureText to \"\"",
    "repeat with candidateWindow in windows",
    'if description of candidateWindow is "警告" or description of candidateWindow is "Warning" then',
    "try",
    "set failureText to value of every static text of candidateWindow as text",
    "end try",
    "end if",
    "end repeat",
    "return failureText",
    "end tell",
    "end tell",
  ]);
  if (!compileState.trim()) return;

  const highlighted = bestEffort("/usr/bin/osascript", [
    "-e",
    'tell application "System Events"',
    "-e",
    'tell process "Microsoft Word"',
    "-e",
    "set frontmost to true",
    "-e",
    "key code 36",
    "-e",
    "delay 0.7",
    "-e",
    'set vbeWindow to first window whose name contains "Microsoft Visual Basic"',
    "-e",
    'set selectedValue to ""',
    "-e",
    "try",
    "-e",
    'set focusedElement to value of attribute "AXFocusedUIElement"',
    "-e",
    'set selectedValue to value of attribute "AXSelectedText" of focusedElement as text',
    "-e",
    "end try",
    "-e",
    'if selectedValue is "" then',
    "-e",
    "repeat with candidateElement in entire contents of vbeWindow",
    "-e",
    "try",
    "-e",
    'if role of candidateElement is "AXTextArea" then',
    "-e",
    'set candidateSelection to value of attribute "AXSelectedText" of candidateElement as text',
    "-e",
    'if candidateSelection is not "" then',
    "-e",
    'set selectedValue to candidateSelection',
    "-e",
    "exit repeat",
    "-e",
    "end if",
    "-e",
    "end if",
    "-e",
    "end try",
    "-e",
    "end repeat",
    "-e",
    "end if",
    "-e",
    'return (name of vbeWindow as text) & "|" & selectedValue',
    "-e",
    "end tell",
    "-e",
    "end tell",
  ], { timeout: 20_000 }).trim();
  let copiedSelection = "";
  if (!highlighted.split("|").at(-1)?.trim()) {
    const clipboardBefore = bestEffort("/usr/bin/pbpaste", []);
    bestEffort("/usr/bin/osascript", [
      "-e",
      'tell application "System Events"',
      "-e",
      'tell process "Microsoft Word"',
      "-e",
      "set frontmost to true",
      "-e",
      'keystroke "c" using {command down}',
      "-e",
      "delay 0.4",
      "-e",
      "end tell",
      "-e",
      "end tell",
    ], { timeout: 10_000 });
    copiedSelection = bestEffort("/usr/bin/pbpaste", []).trim();
    bestEffort("/usr/bin/pbcopy", [], { input: clipboardBefore });
  }
  throw new Error(
    `Word VBE compile failed: ${compileState.trim()}${
      highlighted ? `\nHighlighted statement: ${highlighted}` : ""
    }${copiedSelection ? `\nCopied identifier: ${copiedSelection}` : ""}`,
  );
}

function runIsolatedVbaCompileProbe() {
  closeVbeWindowIfOpen();
  const result = osascript([
    'tell application "Microsoft Word"',
    `if not (exists document ${JSON.stringify(outputDocumentName)}) then error "The isolated Word template is not open"`,
    `activate object document ${JSON.stringify(outputDocumentName)}`,
    'activate',
    `run VB macro macro name ${JSON.stringify(`${outputDocumentName}!VisualTeX_PerformanceNoop`)}`,
    'return "PASS"',
    'end tell',
  ], 60_000).trim();
  if (result !== "PASS") {
    throw new Error(`The isolated Word VBA compile probe returned ${result}`);
  }
}

function replaceAndCompileAdapter() {
  for (const [moduleName, modulePath] of selectedWordModuleSources) {
    if (preserveWord && incrementalBuild) {
      replaceVbaModuleSourceText(moduleName, modulePath);
    } else {
      if (incrementalBuild) removeVbaModule(moduleName);
      importVbaModule(modulePath);
    }
  }
  if (preserveWord && incrementalBuild) {
    runIsolatedVbaCompileProbe();
    return;
  }
  const compileAnchor = requestedModuleNames.has("VTWordAdapter")
    ? "VTWordAdapter"
    : selectedWordModuleSources.at(-1)?.[0] ?? "VTWordAdapter";
  openModuleCodeWindow(compileAnchor);
  compileVbaProject();
}

function baseContainsCurrentVbaSources() {
  const checker = String.raw`
from pathlib import Path
from decimal import Decimal, InvalidOperation
import re
import sys
try:
    from oletools.olevba import VBA_Parser
except Exception:
    print("UNAVAILABLE")
    raise SystemExit(0)

NUMBER_LITERAL = re.compile(
    r"(?<![\w&])"
    r"((?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)"
    r"(?:[#%!@&^])?"
    r"(?![\w])",
    re.IGNORECASE,
)

def normalize_number(match: re.Match[str]) -> str:
    try:
        fixed = format(Decimal(match.group(1)), "f")
    except InvalidOperation:
        return match.group(0)
    if "." in fixed:
        fixed = fixed.rstrip("0").rstrip(".")
    return "0" if fixed in {"", "-0"} else fixed

def strip_vbe_metadata(value: str) -> str:
    lines = value.replace("\r\n", "\n").split("\n")
    if lines and lines[0].lstrip("\ufeff").strip().lower() == "version 1.0 class":
        while lines:
            line = lines.pop(0)
            if line.strip().lower() == "end":
                break
    return "\n".join(
        line for line in lines
        if not line.lstrip().lower().startswith("attribute ")
    )

def normalize_vba(value: str) -> str:
    value = strip_vbe_metadata(value).strip()
    output = []
    non_string = []
    in_string = False
    index = 0

    def flush_non_string() -> None:
        if not non_string:
            return
        segment = "".join(non_string).lower()
        output.append(NUMBER_LITERAL.sub(normalize_number, segment))
        non_string.clear()

    while index < len(value):
        character = value[index]
        if character == '"':
            flush_non_string()
            output.append(character)
            if in_string and index + 1 < len(value) and value[index + 1] == '"':
                output.append('"')
                index += 2
                continue
            in_string = not in_string
            index += 1
            continue
        if in_string:
            output.append(character)
        else:
            non_string.append(character)
        index += 1
    flush_non_string()
    return "".join(output)

base_path, *source_paths = sys.argv[1:]
parser = VBA_Parser(base_path)
try:
    macros = {name: code for _, _, name, code in parser.extract_macros()}
finally:
    parser.close()
checks = [(Path(source_path).name, source_path) for source_path in source_paths]
matched = all(
    macros.get(module_name) is not None
    and normalize_vba(macros[module_name])
        == normalize_vba(Path(source_path).read_text(encoding="utf-8"))
    for module_name, source_path in checks
)
print("MATCH" if matched else "MISMATCH")
`;
  const result = bestEffort("/usr/bin/python3", [
    "-c",
    checker,
    basePath,
    ...wordModuleSources.map(([, modulePath]) => modulePath),
  ], { timeout: 90_000 });
  return result.trim() === "MATCH";
}

function verifyBuiltVba(path) {
  run("/usr/bin/unzip", ["-tqq", path]);
  const vbaProject = run(
    "/usr/bin/unzip",
    ["-p", path, "word/vbaProject.bin"],
    { encoding: "buffer" },
  );
  const required = [
    "VTFileBridgeCall",
    "VTOfficePaths",
    "VTMetadata",
    "VTLauncher",
    "VTErrorHandling",
    "VTRibbonCallbacks",
    "VTAppendText",
    "VTWriteAndLaunchSession",
    "VTPrewarmApplication",
    "VTFinalizeInlineNativeEquation",
    "VTInsertRegisteredEquationCaption",
    "VTWriteWordFailureTrace",
    "VisualTeX_RunWordNativeRegression",
    "AutoExec",
    "word-structured-document-import-20260730-r61",
    "VTWordRibbonDocumentImport",
    "VisualTeX_InsertLatexMarkdownDocument",
    "VTCommitWordDocumentImportDispatch",
    "word-latex-redraw-20260802-r1",
    "VTWordRibbonRedrawSelectionImage",
    "VTWordRibbonRedrawSelectionOmml",
    "VTWordRibbonRedrawDocumentImage",
    "VTWordRibbonRedrawDocumentOmml",
    "VTWordEvents",
    "VTHandleWordBeforeDoubleClick",
    "VTTraceWordDoubleClick",
    "App_WindowBeforeDoubleClick",
    "App_WindowSelectionChange",
    "VTWordRibbonApplyImageFontSizePreset",
    "VTRefreshNumberedImageFormulaFontLayout",
    "VisualTeX_EditImageField",
    "VisualTeX_EditSelectedImageFromNativeMonitor",
    "VTEnsureVisualTeXImageMacroButton",
    "VTNativeMathFastSignature",
    "word-office-performance-20260801-r87",
    "1.2.5",
  ];
  for (const value of required) {
    const utf8 = Buffer.from(value, "utf8");
    const utf16 = Buffer.from(value, "utf16le");
    if (!vbaProject.includes(utf8) && !vbaProject.includes(utf16)) {
      throw new Error(`Built Word VBA project is missing ${value}`);
    }
  }
}

acquireBuildLock();
process.on("exit", releaseBuildLock);
mkdirSync(dirname(outputPath), { recursive: true });
if (!existsSync(basePath)) throw new Error(`Base Word template is missing: ${basePath}`);

const originalTrust = readVbaTrust();
const originalDocumentSnapshot = preserveWord ? wordDocumentSnapshot() : [];
const originalVisualTeXAddinInstalled = preserveWord
  ? wordVisualTeXAddinInstalled()
  : null;
let buildSucceeded = false;
try {
  if (preserveWord) {
    if (!originalTrust.enabled) {
      throw new Error(
        "--preserve-word requires VBAObjectModelIsTrusted to be enabled before the build",
      );
    }
    dismissVbeFileDialogIfOpen();
    dismissVbeWarningDialogsIfOpen();
    closeVbeWindowIfOpen();
    if (incrementalBuild && originalVisualTeXAddinInstalled === true) {
      setWordVisualTeXAddinInstalled(false);
    }
  } else {
    closeWordWithoutSaving();
    moveStartupTemplatesOut();
    if (!incrementalBuild) moveNormalTemplateOut();
    setVbaTrust(true);
  }
  rmSync(outputPath, { force: true });

  if (!preserveWord) {
    run("/usr/bin/open", ["-gj", "-a", "Microsoft Word"]);
  }
  waitForWordUiReady();
  if (incrementalBuild) {
    copyFileSync(basePath, outputPath);
    osascript([
      'tell application "Microsoft Word"',
      `open file name ${JSON.stringify(outputPath)}`,
      "activate",
      "end tell",
    ], 120_000);
  } else {
    osascript([
      'tell application "Microsoft Word"',
      "set buildDocument to make new document",
      `save as buildDocument file name ${JSON.stringify(outputPath)} file format format templateME add to recent files false`,
      "activate",
      "end tell",
    ], 120_000);
  }
  waitForWordUiReady();
  replaceAndCompileAdapter();
  osascript([
    'tell application "Microsoft Word"',
    `save as document ${JSON.stringify(outputDocumentName)} file name ${JSON.stringify(outputPath)}`,
    "end tell",
  ]);
  sleep(2_500);
  if (preserveWord) {
    dismissVbeFileDialogIfOpen();
    dismissVbeWarningDialogsIfOpen();
    closeVbeWindowIfOpen();
    closeBuildDocumentWithoutSaving();
  } else {
    closeWordWithoutSaving();
  }
  verifyBuiltVba(outputPath);

  const size = statSync(outputPath).size;
  if (size < 100_000) {
    throw new Error(`Rebuilt Word template is unexpectedly small: ${size} bytes`);
  }
  process.stdout.write(`Rebuilt and VBE-compiled ${outputPath} (${size} bytes).\n`);
  buildSucceeded = true;
} finally {
  if (preserveWord) {
    dismissVbeFileDialogIfOpen();
    dismissVbeWarningDialogsIfOpen();
    closeVbeWindowIfOpen();
    closeBuildDocumentWithoutSaving();
    setWordVisualTeXAddinInstalled(originalVisualTeXAddinInstalled);
    const finalDocumentSnapshot = wordDocumentSnapshot();
    releaseBuildLock();
    if (
      JSON.stringify(finalDocumentSnapshot) !==
      JSON.stringify(originalDocumentSnapshot)
    ) {
      throw new Error(
        `The isolated DOTM build changed the user's open Word documents: ${JSON.stringify({ originalDocumentSnapshot, finalDocumentSnapshot })}`,
      );
    }
  } else {
    if (buildSucceeded || !keepWordOpenOnError) closeWordWithoutSaving();
    if (!incrementalBuild) restoreNormalTemplate();
    restoreStartupTemplates();
    restoreVbaTrust(originalTrust);
    releaseBuildLock();
  }
}
