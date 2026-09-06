import { execFileSync, spawnSync } from "node:child_process";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  rmSync,
  statSync,
} from "node:fs";
import { homedir } from "node:os";
import { basename, dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

if (process.platform !== "darwin") {
  throw new Error("The PowerPoint VBE compiler is available only on macOS.");
}

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const scratchRoot = join(
  homedir(),
  "Library",
  "Group Containers",
  "UBF8T346G9.Office",
  "VisualTeX",
  "Scratch",
);
const argument = (name) => {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
};
const basePath = resolve(
  argument("--base") ??
    join(scratchRoot, "VisualTeXPowerPointPerformanceBuild4.pptm"),
);
const outputPath = resolve(
  argument("--output") ??
    join(scratchRoot, "VisualTeXPowerPointPerformanceOptimized.pptm"),
);
const keepOpenOnError = process.argv.includes("--keep-open-on-error");
const preservePowerPoint = process.argv.includes("--preserve-powerpoint");
const offlineOfficeRoot = join(repositoryRoot, "office", "macos-offline");
const modules = [
  {
    name: "VTOfficePaths",
    path: join(offlineOfficeRoot, "shared", "VTOfficePaths.bas"),
  },
  {
    name: "VTProtocol",
    path: join(offlineOfficeRoot, "shared", "VTProtocol.bas"),
  },
  {
    name: "VTMetadata",
    path: join(offlineOfficeRoot, "shared", "VTMetadata.bas"),
  },
  {
    name: "VTLauncher",
    path: join(offlineOfficeRoot, "shared", "VTLauncher.bas"),
  },
  {
    name: "VTPowerPointAdapter",
    path: join(
      offlineOfficeRoot,
      "powerpoint",
      "VTPowerPointAdapter.bas",
    ),
  },
];
const outputPresentationName = basename(outputPath, ".pptm");
const lockPath = join(scratchRoot, "VisualTeXPowerPointPerformanceCompile.lock");

function run(program, args, options = {}) {
  return execFileSync(program, args, {
    encoding: options.encoding ?? "utf8",
    input: options.input,
    timeout: options.timeout ?? 60_000,
    maxBuffer: 64 * 1024 * 1024,
    stdio: options.stdio ?? ["ignore", "pipe", "pipe"],
  });
}

function bestEffort(program, args, options = {}) {
  try {
    return run(program, args, options);
  } catch {
    return "";
  }
}

function sleep(milliseconds) {
  spawnSync("/bin/sleep", [String(milliseconds / 1000)], {
    stdio: "ignore",
  });
}

function osascript(lines, timeout = 60_000) {
  return run(
    "/usr/bin/osascript",
    lines.flatMap((line) => ["-e", line]),
    { timeout },
  ).trim();
}

function acquireLock() {
  mkdirSync(scratchRoot, { recursive: true });
  try {
    mkdirSync(lockPath);
  } catch (error) {
    if (!(error instanceof Error && "code" in error && error.code === "EEXIST")) {
      throw error;
    }
    throw new Error(`PowerPoint performance compile lock exists: ${lockPath}`);
  }
}

function releaseLock() {
  rmSync(lockPath, { recursive: true, force: true });
}

function closeOutputPresentationWithoutSaving() {
  bestEffort(
    "/usr/bin/osascript",
    [
      "-e",
      'tell application "Microsoft PowerPoint"',
      "-e",
      `if exists presentation ${JSON.stringify(`${outputPresentationName}.pptm`)} then close presentation ${JSON.stringify(`${outputPresentationName}.pptm`)} saving no`,
      "-e",
      "end tell",
    ],
    { timeout: 20_000 },
  );
  sleep(1_000);
}

function closePowerPointWithoutSaving() {
  bestEffort(
    "/usr/bin/osascript",
    ["-e", 'tell application "Microsoft PowerPoint" to quit saving no'],
    { timeout: 20_000 },
  );
  sleep(1_500);
  for (const pid of bestEffort("/usr/bin/pgrep", ["-x", "Microsoft PowerPoint"])
    .trim()
    .split(/\s+/)
    .filter(Boolean)) {
    bestEffort("/bin/kill", ["-9", pid]);
  }
  sleep(1_000);
}

function waitForPresentation() {
  for (let attempt = 0; attempt < 80; attempt += 1) {
    const state = bestEffort(
      "/usr/bin/osascript",
      [
        "-e",
        'tell application "Microsoft PowerPoint"',
        "-e",
        `if exists presentation ${JSON.stringify(`${outputPresentationName}.pptm`)} then return "READY"`,
        "-e",
        "end tell",
      ],
      { timeout: 5_000 },
    ).trim();
    if (state === "READY") return;
    sleep(500);
  }
  throw new Error(`PowerPoint did not open ${basename(outputPath)}.`);
}

function dismissOpenPrompts() {
  for (let attempt = 0; attempt < 20; attempt += 1) {
    const state = bestEffort(
      "/usr/bin/osascript",
      [
        "-e",
        'tell application "Microsoft PowerPoint" to activate',
        "-e",
        "delay 0.2",
        "-e",
        'tell application "System Events"',
        "-e",
        'tell process "Microsoft PowerPoint"',
        "-e",
        'set handledPrompt to false',
        "-e",
        'repeat with candidateWindow in windows',
        "-e",
        'repeat with buttonName in {"禁用宏", "Disable Macros"}',
        "-e",
        "try",
        "-e",
        'if exists button (buttonName as text) of candidateWindow then',
        "-e",
        'click button (buttonName as text) of candidateWindow',
        "-e",
        'set handledPrompt to true',
        "-e",
        "end if",
        "-e",
        "end try",
        "-e",
        "end repeat",
        "-e",
        'if handledPrompt is false then',
        "-e",
        'set warningText to ""',
        "-e",
        "try",
        "-e",
        'set warningText to value of every static text of candidateWindow as text',
        "-e",
        "end try",
        "-e",
        'if warningText contains "无法加载外接程序" or warningText contains "could not load add-in" then',
        "-e",
        'repeat with buttonName in {"确定", "OK"}',
        "-e",
        "try",
        "-e",
        'if exists button (buttonName as text) of candidateWindow then',
        "-e",
        'click button (buttonName as text) of candidateWindow',
        "-e",
        'set handledPrompt to true',
        "-e",
        "end if",
        "-e",
        "end try",
        "-e",
        "end repeat",
        "-e",
        "end if",
        "-e",
        "end if",
        "-e",
        "end repeat",
        "-e",
        'if handledPrompt then return "HANDLED"',
        "-e",
        'return "CLEAR"',
        "-e",
        "end tell",
        "-e",
        "end tell",
      ],
      { timeout: 5_000 },
    ).trim();
    if (state === "HANDLED") {
      sleep(1_000);
      continue;
    }
    if (state === "CLEAR") return;
    sleep(500);
  }
}

function openVbe() {
  osascript(
    [
      'tell application "Microsoft PowerPoint" to activate',
      "delay 0.8",
      'tell application "System Events"',
      'tell process "Microsoft PowerPoint"',
      "set frontmost to true",
      `set expectedPrefix to ${JSON.stringify(`Microsoft Visual Basic - ${outputPresentationName}`)}`,
      "repeat with candidateWindow in windows",
      "try",
      "if name of candidateWindow starts with expectedPrefix then",
      'perform action "AXRaise" of candidateWindow',
      'return "OPEN"',
      "end if",
      "end try",
      "end repeat",
      "set toolsMenu to menu 1 of menu bar item \"工具\" of menu bar 1",
      "set macroMenu to menu 1 of menu item \"宏\" of toolsMenu",
      'click menu item "Visual Basic 编辑器" of macroMenu',
      "end tell",
      "end tell",
    ],
    30_000,
  );

  for (let attempt = 0; attempt < 40; attempt += 1) {
    const state = bestEffort(
      "/usr/bin/osascript",
      [
        "-e",
        'tell application "Microsoft PowerPoint" to activate',
        "-e",
        "delay 0.2",
        "-e",
        'tell application "System Events"',
        "-e",
        'tell process "Microsoft PowerPoint"',
        "-e",
        `set expectedPrefix to ${JSON.stringify(`Microsoft Visual Basic - ${outputPresentationName}`)}`,
        "-e",
        "repeat with candidateWindow in windows",
        "-e",
        "try",
        "-e",
        "if name of candidateWindow starts with expectedPrefix then",
        "-e",
        'perform action "AXRaise" of candidateWindow',
        "-e",
        'if (count of (every UI element of candidateWindow whose role is "AXOutline")) > 0 then return "READY"',
        "-e",
        'keystroke "r" using {command down}',
        "-e",
        "end if",
        "-e",
        "end try",
        "-e",
        "end repeat",
        "-e",
        'return "WAIT"',
        "-e",
        "end tell",
        "-e",
        "end tell",
      ],
      { timeout: 5_000 },
    ).trim();
    if (state === "READY") return;
    sleep(500);
  }
  throw new Error("PowerPoint VBE did not expose its project outline.");
}

function selectVbaModule(moduleName) {
  osascript(
    [
      'tell application "Microsoft PowerPoint" to activate',
      "delay 0.5",
      'tell application "System Events"',
      'tell process "Microsoft PowerPoint"',
      "set frontmost to true",
      `set expectedPrefix to ${JSON.stringify(`Microsoft Visual Basic - ${outputPresentationName}`)}`,
      "set vbeWindow to first window whose name starts with expectedPrefix",
      'perform action "AXRaise" of vbeWindow',
      'set projectOutline to first UI element of vbeWindow whose role is "AXOutline"',
      "repeat with expansionPass from 1 to 4",
      "set rowIndex to 1",
      "repeat while rowIndex is less than or equal to count of rows of projectOutline",
      "set rowCell to UI element 1 of row rowIndex of projectOutline",
      "try",
      'set disclosure to first UI element of rowCell whose role is "AXDisclosureTriangle"',
      "if value of disclosure is false then click disclosure",
      "end try",
      "set rowIndex to rowIndex + 1",
      "end repeat",
      "end repeat",
      "set moduleRow to 0",
      "repeat with rowIndex from 1 to count of rows of projectOutline",
      "set rowCell to UI element 1 of row rowIndex of projectOutline",
      "set rowNames to name of every UI element of rowCell as text",
      `if rowNames contains ${JSON.stringify(moduleName)} then set moduleRow to rowIndex`,
      "end repeat",
      `if moduleRow is 0 then error ${JSON.stringify(`${moduleName} was not found in the PowerPoint VBA project.`)}`,
      "select row moduleRow of projectOutline",
      "end tell",
      "end tell",
    ],
    60_000,
  );
}

function removeVbaModule(moduleName) {
  selectVbaModule(moduleName);
  osascript(
    [
      'tell application "Microsoft PowerPoint" to activate',
      "delay 0.3",
      'tell application "System Events"',
      'tell process "Microsoft PowerPoint"',
      "set frontmost to true",
      `set expectedPrefix to ${JSON.stringify(`Microsoft Visual Basic - ${outputPresentationName}`)}`,
      "set vbeWindow to first window whose name starts with expectedPrefix",
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
      `if removedModule is false then error ${JSON.stringify(`Unable to remove ${moduleName}.`)}`,
      "delay 0.8",
      "repeat with candidateWindow in windows",
      "try",
      "repeat with candidateButton in buttons of candidateWindow",
      "set buttonName to name of candidateButton as text",
      'if buttonName starts with "否" or buttonName starts with "No" then',
      'perform action "AXRaise" of candidateWindow',
      "click candidateButton",
      "exit repeat",
      "end if",
      "end repeat",
      "end try",
      "end repeat",
      "delay 0.8",
      "end tell",
      "end tell",
    ],
    60_000,
  );
}

function importVbaModule(modulePath) {
  osascript(
    [
      'tell application "Microsoft PowerPoint" to activate',
      "delay 0.3",
      'tell application "System Events"',
      'tell process "Microsoft PowerPoint"',
      "set frontmost to true",
      `set expectedPrefix to ${JSON.stringify(`Microsoft Visual Basic - ${outputPresentationName}`)}`,
      "set vbeWindow to first window whose name starts with expectedPrefix",
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
      'if openedImport is false then error "Unable to open VBE Import File."',
      "delay 0.9",
      'keystroke "g" using {command down, shift down}',
      "delay 0.6",
      'set pathField to value of attribute "AXFocusedUIElement"',
      `set value of pathField to ${JSON.stringify(modulePath)}`,
      "key code 36",
      "delay 0.9",
      "key code 36",
      "delay 1.5",
      "end tell",
      "end tell",
    ],
    60_000,
  );
}

function replaceVbaModule(moduleName, modulePath) {
  removeVbaModule(moduleName);
  importVbaModule(modulePath);
}

function compileVbaProject() {
  const result = osascript(
    [
      'tell application "Microsoft PowerPoint" to activate',
      "delay 0.5",
      'tell application "System Events"',
      'tell process "Microsoft PowerPoint"',
      "set frontmost to true",
      `set expectedPrefix to ${JSON.stringify(`Microsoft Visual Basic - ${outputPresentationName}`)}`,
      "set vbeWindow to first window whose name starts with expectedPrefix",
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
      'if startedCompile is false then return "ALREADY_COMPILED"',
      "delay 2",
      'set failureText to ""',
      "repeat with candidateWindow in windows",
      "try",
      'if description of candidateWindow is "警告" or description of candidateWindow is "Warning" then',
      'set failureText to value of every static text of candidateWindow as text',
      "end if",
      "end try",
      "end repeat",
      'if failureText is not "" then return "ERROR|" & failureText',
      'return "COMPILED"',
      "end tell",
      "end tell",
    ],
    60_000,
  );
  if (result.startsWith("ERROR|")) {
    throw new Error(`PowerPoint VBE compile failed: ${result.slice(6)}`);
  }
  if (!new Set(["COMPILED", "ALREADY_COMPILED"]).has(result)) {
    throw new Error(`Unexpected PowerPoint VBE compile result: ${result}`);
  }
}

function savePresentation() {
  osascript(
    [
      'tell application "Microsoft PowerPoint"',
      `save presentation ${JSON.stringify(`${outputPresentationName}.pptm`)}`,
      "end tell",
    ],
    60_000,
  );
  sleep(3_000);
}

function verifyCompiledModules() {
  const checker = String.raw`
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re
import sys
from oletools.olevba import VBA_Parser

NUMBER_LITERAL = re.compile(
    r"(?<![\w&])((?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)(?:[#%!@&^])?(?![\w])",
    re.IGNORECASE,
)

def normalize_number(match):
    try:
        fixed = format(Decimal(match.group(1)), "f")
    except InvalidOperation:
        return match.group(0)
    if "." in fixed:
        fixed = fixed.rstrip("0").rstrip(".")
    return "0" if fixed in {"", "-0"} else fixed

def normalize_vba(value):
    lines = value.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    lines = [line for line in lines if not line.lstrip().lower().startswith("attribute ")]
    value = "\n".join(lines).strip()
    output = []
    non_string = []
    in_string = False
    index = 0
    def flush():
        if non_string:
            output.append(NUMBER_LITERAL.sub(normalize_number, "".join(non_string).lower()))
            non_string.clear()
    while index < len(value):
        character = value[index]
        if character == '"':
            flush()
            output.append(character)
            if in_string and index + 1 < len(value) and value[index + 1] == '"':
                output.append('"')
                index += 2
                continue
            in_string = not in_string
        elif in_string:
            output.append(character)
        else:
            non_string.append(character)
        index += 1
    flush()
    return "".join(output)

presentation, *source_paths = sys.argv[1:]
parser = VBA_Parser(presentation)
try:
    macros = {Path(name).stem: code for _, _, name, code in parser.extract_macros()}
finally:
    parser.close()
for source_path in source_paths:
    module_name = Path(source_path).stem
    saved = macros.get(module_name)
    if saved is None:
        raise SystemExit(f"MISSING|{module_name}")
    source = Path(source_path).read_text(encoding="utf-8")
    if normalize_vba(saved) != normalize_vba(source):
        raise SystemExit(f"MISMATCH|{module_name}")
print("MATCH")
`;
  const result = run("/usr/bin/python3", [
    "-c",
    checker,
    outputPath,
    ...modules.map((module) => module.path),
  ]).trim();
  if (result !== "MATCH") {
    throw new Error(`Compiled PowerPoint VBA verification failed: ${result}`);
  }
}

acquireLock();
let succeeded = false;
try {
  if (!existsSync(basePath)) throw new Error(`Missing base PPTM: ${basePath}`);
  for (const module of modules) {
    if (!existsSync(module.path)) throw new Error(`Missing VBA source: ${module.path}`);
  }
  if (preservePowerPoint) {
    closeOutputPresentationWithoutSaving();
  } else {
    closePowerPointWithoutSaving();
  }
  rmSync(outputPath, { force: true });
  copyFileSync(basePath, outputPath);
  run("/usr/bin/open", ["-b", "com.microsoft.Powerpoint", outputPath]);
  waitForPresentation();
  dismissOpenPrompts();
  openVbe();
  for (const module of modules) {
    replaceVbaModule(module.name, module.path);
  }
  compileVbaProject();
  savePresentation();
  if (preservePowerPoint) {
    closeOutputPresentationWithoutSaving();
  } else {
    closePowerPointWithoutSaving();
  }
  verifyCompiledModules();
  const size = statSync(outputPath).size;
  if (size < 100_000) throw new Error(`Compiled PPTM is unexpectedly small: ${size}`);
  process.stdout.write(`Compiled and verified ${outputPath} (${size} bytes).\n`);
  succeeded = true;
} finally {
  if (succeeded || !keepOpenOnError) {
    if (preservePowerPoint) {
      closeOutputPresentationWithoutSaving();
    } else {
      closePowerPointWithoutSaving();
    }
  }
  releaseLock();
}
