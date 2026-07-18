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
const adapterPath = join(
  repositoryRoot,
  "office",
  "macos-offline",
  "word",
  "VTWordAdapter.bas",
);
const startupRoot = join(
  homedir(),
  "Library",
  "Group Containers",
  "UBF8T346G9.Office",
  "User Content.localized",
  "Startup.localized",
  "Word",
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

function moveStartupTemplatesOut() {
  mkdirSync(startupRoot, { recursive: true });
  mkdirSync(backupRoot, { recursive: true });
  for (const name of readdirSync(startupRoot)) {
    if (!/^VisualTeX\.dotm/i.test(name)) continue;
    renameSync(join(startupRoot, name), join(backupRoot, name));
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

function adapterSourceForCodeWindow() {
  const source = readFileSync(adapterPath, "utf8");
  return source.replace(/^Attribute VB_Name = "VTWordAdapter"\r?\n/, "");
}

function openVbeWindow() {
  osascript([
    'tell application "Microsoft Word" to activate',
    'tell application "System Events"',
    'tell process "Microsoft Word"',
    "set frontmost to true",
    'click menu item "Visual Basic 编辑器" of menu 1 of menu item "宏" of menu 1 of menu bar item "工具" of menu bar 1',
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
    if (windowNames.includes("Microsoft Visual Basic")) return;
    if (attempt === 9 || attempt === 19) {
      bestEffort("/usr/bin/osascript", [
        "-e",
        'tell application "Microsoft Word" to activate',
        "-e",
        'tell application "System Events"',
        "-e",
        'tell process "Microsoft Word"',
        "-e",
        "set frontmost to true",
        "-e",
        'click menu item "Visual Basic 编辑器" of menu 1 of menu item "宏" of menu 1 of menu bar item "工具" of menu bar 1',
        "-e",
        "end tell",
        "-e",
        "end tell",
      ], { timeout: 10_000 });
    }
  }
  throw new Error("Word did not open its Visual Basic Editor window.");
}

function openAdapterCodeWindow() {
  openVbeWindow();

  osascript([
    'tell application "System Events"',
    'tell process "Microsoft Word"',
    "set frontmost to true",
    'set vbeWindow to first window whose name contains "Microsoft Visual Basic"',
    'perform action "AXRaise" of vbeWindow',
    'set projectOutline to first UI element of vbeWindow whose role is "AXOutline"',
    "set rowIndex to 1",
    "repeat while rowIndex is less than or equal to count of rows of projectOutline",
    "set rowCell to UI element 1 of row rowIndex of projectOutline",
    "set rowNames to name of every UI element of rowCell as text",
    'if rowNames contains "模块表" or rowNames contains "Modules" then',
    "try",
    'set disclosure to first UI element of rowCell whose role is "AXDisclosureTriangle"',
    "if value of disclosure is false then click disclosure",
    "delay 0.3",
    "end try",
    "end if",
    "set rowIndex to rowIndex + 1",
    "end repeat",
    "set adapterRow to 0",
    "repeat with rowIndex from 1 to count of rows of projectOutline",
    "set rowCell to UI element 1 of row rowIndex of projectOutline",
    "set rowNames to name of every UI element of rowCell as text",
    'if rowNames contains "VTWordAdapter" then set adapterRow to rowIndex',
    "end repeat",
    'if adapterRow is 0 then error "VTWordAdapter was not found in the Word VBA project"',
    "select row adapterRow of projectOutline",
    'click menu item "代码" of menu 1 of menu bar item "查看" of menu bar 1',
    "delay 1",
    "end tell",
    "end tell",
  ]);
}

function replaceAndCompileAdapter() {
  run("/usr/bin/pbcopy", [], { input: adapterSourceForCodeWindow() });
  const compileState = osascript([
    'tell application "System Events"',
    'tell process "Microsoft Word"',
    "set frontmost to true",
    'set vbeWindow to first window whose name contains "Microsoft Visual Basic"',
    'perform action "AXRaise" of vbeWindow',
    "click at {800, 500}",
    'click menu item "全选" of menu 1 of menu bar item "编辑" of menu bar 1',
    "delay 0.5",
    'click menu item "粘贴" of menu 1 of menu bar item "编辑" of menu bar 1',
    "delay 2",
    'click menu item "编译 Project" of menu 1 of menu bar item "调试" of menu bar 1',
    "delay 2",
    "set failureText to \"\"",
    "repeat with candidateWindow in windows",
    'if description of candidateWindow is "警告" then',
    "try",
    "set failureText to value of every static text of candidateWindow as text",
    "end try",
    "end if",
    "end repeat",
    "return failureText",
    "end tell",
    "end tell",
  ]);
  if (compileState.trim()) {
    throw new Error(`Word VBE compile failed: ${compileState.trim()}`);
  }
}

function baseContainsCurrentAdapterSource() {
  const checker = String.raw`
from pathlib import Path
import sys
try:
    from oletools.olevba import VBA_Parser
except Exception:
    print("UNAVAILABLE")
    raise SystemExit(0)

def normalize_vba(value: str) -> str:
    value = value.replace("\r\n", "\n").strip()
    output = []
    in_string = False
    index = 0
    while index < len(value):
        character = value[index]
        if character == '"':
            output.append(character)
            if in_string and index + 1 < len(value) and value[index + 1] == '"':
                output.append('"')
                index += 2
                continue
            in_string = not in_string
            index += 1
            continue
        output.append(character if in_string else character.lower())
        index += 1
    return "".join(output)

base_path, source_path = sys.argv[1:3]
parser = VBA_Parser(base_path)
try:
    macros = {name: code for _, _, name, code in parser.extract_macros()}
finally:
    parser.close()
built = macros.get("VTWordAdapter.bas")
source = Path(source_path).read_text(encoding="utf-8")
print("MATCH" if built is not None and normalize_vba(built) == normalize_vba(source) else "MISMATCH")
`;
  const result = bestEffort("/usr/bin/python3", [
    "-c",
    checker,
    basePath,
    adapterPath,
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
    "VTFinalizeInlineNativeEquation",
    "VTInsertRegisteredEquationCaption",
    "VTWriteWordFailureTrace",
    "VisualTeX_RunWordNativeRegression",
  ];
  for (const value of required) {
    const utf8 = Buffer.from(value, "utf8");
    const utf16 = Buffer.from(value, "utf16le");
    if (!vbaProject.includes(utf8) && !vbaProject.includes(utf16)) {
      throw new Error(`Built Word VBA project is missing ${value}`);
    }
  }
}

mkdirSync(dirname(outputPath), { recursive: true });
if (!existsSync(basePath)) throw new Error(`Base Word template is missing: ${basePath}`);

if (baseContainsCurrentAdapterSource()) {
  copyFileSync(basePath, outputPath);
  verifyBuiltVba(outputPath);
  process.stdout.write(
    `The compiled base already contains the current VTWordAdapter source; copied and verified ${outputPath}.\n`,
  );
  process.exit(0);
}

const originalTrust = readVbaTrust();
try {
  closeWordWithoutSaving();
  moveStartupTemplatesOut();
  setVbaTrust(true);
  copyFileSync(basePath, outputPath);

  osascript([
    'tell application "Microsoft Word"',
    `open file name ${JSON.stringify(outputPath)}`,
    "make new document",
    "activate",
    "end tell",
  ]);
  sleep(3_500);
  openAdapterCodeWindow();
  replaceAndCompileAdapter();
  osascript([
    'tell application "Microsoft Word"',
    `save as document ${JSON.stringify(outputDocumentName)} file name ${JSON.stringify(outputPath)}`,
    "end tell",
  ]);
  sleep(2_500);
  closeWordWithoutSaving();
  verifyBuiltVba(outputPath);

  const size = statSync(outputPath).size;
  if (size < 100_000) {
    throw new Error(`Rebuilt Word template is unexpectedly small: ${size} bytes`);
  }
  process.stdout.write(`Rebuilt and VBE-compiled ${outputPath} (${size} bytes).\n`);
} finally {
  closeWordWithoutSaving();
  restoreStartupTemplates();
  restoreVbaTrust(originalTrust);
}
