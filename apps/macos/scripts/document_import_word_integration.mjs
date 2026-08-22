import { spawn, spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { homedir, tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { unzipSync, zipSync, strToU8 } from "fflate";
import { inflateSync } from "node:zlib";
import { latexToSvg } from "../src/export/latexToSvg.ts";
import {
  createFormulaMetadata,
  decodeFormulaMetadata,
  encodeFormulaMetadata,
} from "../src/office/shared/formulaMetadata.ts";
import {
  normalizeFormulaEditorDocument,
  serializeFormulaEditorDocument,
} from "../src/office/shared/formulaEditorDocument.ts";
import { renderOfficeFormulaArtifacts } from "../src/office/shared/formulaRenderArtifacts.ts";

const repositoryRoot = resolve(new URL("..", import.meta.url).pathname);
const templatePath = join(
  repositoryRoot,
  "office/macos-offline/resources/VisualTeX.dotm",
);
const runtimeRoot = join(
  homedir(),
  "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime",
);
const pdfExportRequestPath = join(
  runtimeRoot,
  "document-import-regression-pdf-path.txt",
);
const pdfExportStatusPath = join(
  runtimeRoot,
  "document-import-regression-pdf-status.txt",
);
const imageEditStatusPath = join(
  runtimeRoot,
  "document-import-regression-image-edit-status.txt",
);
const formulaRegressionStatusPath = join(
  runtimeRoot,
  "document-import-regression-formula-status.txt",
);
const physicalScreenBoundsPath = join(
  runtimeRoot,
  "physical-double-click-screen-bounds.txt",
);
const pictureRoutingPerformancePath = join(
  runtimeRoot,
  "picture-routing-performance.txt",
);
const workspaceVisualTeXBinary = join(
  repositoryRoot,
  "src-tauri/target/release/bundle/macos/VisualTeX.app/Contents/MacOS/visualtex",
);
const officeScratchRoot = join(
  homedir(),
  "Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch",
);
const wordStartupRoot = join(
  homedir(),
  "Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word",
);
const installedWordAddinPath = join(wordStartupRoot, "VisualTeX.dotm");
const installedWordAddinBackupPath = join(
  tmpdir(),
  `visualtex-installed-word-backup-${process.pid}.dotm`,
);
const coordinatePdfPath = join(
  officeScratchRoot,
  `document-import-geometry-${process.pid}.pdf`,
);
const reopenedDocumentPath = join(
  officeScratchRoot,
  `document-import-reopen-${process.pid}.docx`,
);
const firstFrameDocumentPath = join(
  officeScratchRoot,
  `word-image-first-frame-${process.pid}.docx`,
);
const firstFramePdfPath = join(
  officeScratchRoot,
  `word-image-first-frame-${process.pid}.pdf`,
);
const firstFrameScrolledPdfPath = join(
  officeScratchRoot,
  `word-image-first-frame-scrolled-${process.pid}.pdf`,
);
const firstFrameReopenedPdfPath = join(
  officeScratchRoot,
  `word-image-first-frame-reopened-${process.pid}.pdf`,
);
const finalBinaryPhysicalStatusPath = join(
  officeScratchRoot,
  "final-binary-physical-double-click.json",
);
const wordPhysicalPerformanceStatusPath = join(
  officeScratchRoot,
  "word-physical-performance.json",
);
const pictureRoutingFixturePng = join(
  officeScratchRoot,
  "picture-routing-fixture.png",
);
const pictureRoutingBrowserArtifactsPath = join(
  officeScratchRoot,
  "word-first-frame-browser-artifacts.json",
);
const sessionsRoot = join(runtimeRoot, "OfficeSessions");
const nativeRoot = join(runtimeRoot, "NativeDocuments");
const physicalDoubleClick = process.argv.includes("--physical-double-click");
const physicalApplyPerformance = process.argv.includes(
  "--physical-apply-performance",
);
const createSourceFormattedEquationRegression = process.argv.includes(
  "--create-source-formatted-equation",
);
const createImageNativeMonitorRegression = process.argv.includes(
  "--create-image-native-monitor-double-click",
);
const physicalHitTestOnly = process.argv.includes("--physical-hit-test-only");
const createImageDisplayRegression = process.argv.includes(
  "--create-image-display",
);
const createImageNumberedRegression = process.argv.includes(
  "--create-image-numbered",
);
const createNativeInlineRegression = process.argv.includes(
  "--create-native-inline",
);
const createNativeDisplayRegression = process.argv.includes(
  "--create-native-display",
);
const createNativeNumberedRegression = process.argv.includes(
  "--create-native-numbered",
);
const createNativeRegression =
  createNativeInlineRegression ||
  createNativeDisplayRegression ||
  createNativeNumberedRegression;
const createImagePhysicalRegression =
  process.argv.includes("--create-image-physical-double-click") ||
  createImageNativeMonitorRegression ||
  createSourceFormattedEquationRegression;
const createImageRegression =
  process.argv.includes("--create-image") ||
  createImageDisplayRegression ||
  createImageNumberedRegression ||
  createImagePhysicalRegression;
const createFormulaRegression = createImageRegression || createNativeRegression;
const physicalTargets = new Set([
  "image-inline",
  "image-block",
  "image-numbered",
  "image-align",
  "image-align-star",
  "omml-inline",
  "omml-block",
  "omml-align",
  "omml-align-star",
]);

function commandLineOption(name) {
  const exactIndex = process.argv.indexOf(name);
  const assigned = process.argv.find((argument) =>
    argument.startsWith(`${name}=`),
  );
  if (exactIndex >= 0 && assigned) {
    throw new Error(`Specify ${name} only once`);
  }
  if (exactIndex >= 0) {
    const value = process.argv[exactIndex + 1];
    if (!value || value.startsWith("--")) {
      throw new Error(`${name} requires a value`);
    }
    return value;
  }
  return assigned?.slice(name.length + 1) ?? "";
}

const physicalTarget = commandLineOption("--physical-target");
const pictureRoutingTarget = commandLineOption("--picture-routing-target");
const pictureRoutingPerformanceValue = commandLineOption(
  "--picture-routing-performance",
);
const pictureRoutingPerformance = pictureRoutingPerformanceValue
  ? Number.parseInt(pictureRoutingPerformanceValue, 10)
  : 0;
if (
  pictureRoutingPerformanceValue &&
  ![1, 100, 1000].includes(pictureRoutingPerformance)
) {
  throw new Error(
    `--picture-routing-performance must be 1, 100 or 1000: ${pictureRoutingPerformanceValue}`,
  );
}
const pictureRoutingTargets = new Set([
  "ordinary-inline",
  "ordinary-floating",
  "forged-prefix",
  "damaged-metadata",
]);
if (pictureRoutingTarget && !pictureRoutingTargets.has(pictureRoutingTarget)) {
  throw new Error(`Unknown --picture-routing-target: ${pictureRoutingTarget}`);
}
const requestedWordAddinPath = commandLineOption("--word-addin");
const activeTemplatePath = requestedWordAddinPath || templatePath;
const createFormulaLatexOption = commandLineOption("--formula-latex");
const createFormulaLetterFontOption = commandLineOption("--formula-letter-font");
const supportedFormulaLetterFonts = new Set([
  "katex",
  "times",
  "cambria",
  "stix",
  "palatino",
  "helvetica",
]);
if (
  createFormulaLetterFontOption &&
  !supportedFormulaLetterFonts.has(createFormulaLetterFontOption)
) {
  throw new Error(
    `--formula-letter-font must be one of ${[...supportedFormulaLetterFonts].join("|")}`,
  );
}
if ((createFormulaLatexOption || createFormulaLetterFontOption) && !createFormulaRegression) {
  throw new Error("--formula-latex/--formula-letter-font require a --create-image/--create-native regression mode");
}
const firstFrameArtifactPath = commandLineOption("--first-frame-artifacts");
const firstFrameImageRegression = Boolean(firstFrameArtifactPath);
const itemLimitOption = commandLineOption("--item-limit");
const diagnosticItemLimit = itemLimitOption ? Number(itemLimitOption) : 17;
if (
  !Number.isInteger(diagnosticItemLimit) ||
  diagnosticItemLimit < 1 ||
  diagnosticItemLimit > 17
) {
  throw new Error("--item-limit must be an integer from 1 through 17");
}
if (physicalDoubleClick && !physicalTargets.has(physicalTarget)) {
  throw new Error(
    "--physical-double-click requires --physical-target " +
      [...physicalTargets].join("|"),
  );
}
if (!physicalDoubleClick && physicalTarget) {
  throw new Error("--physical-target requires --physical-double-click");
}
if (physicalApplyPerformance && !physicalDoubleClick) {
  throw new Error(
    "--physical-apply-performance requires --physical-double-click",
  );
}
if (physicalApplyPerformance && physicalTarget !== "image-inline") {
  throw new Error(
    "--physical-apply-performance currently requires --physical-target image-inline",
  );
}
if (createImageRegression && physicalDoubleClick) {
  throw new Error("--create-image cannot be combined with --physical-double-click");
}
if (physicalHitTestOnly && !createImagePhysicalRegression) {
  throw new Error(
    "--physical-hit-test-only requires --create-image-physical-double-click or --create-image-native-monitor-double-click",
  );
}
if (
  firstFrameImageRegression &&
  (createImageRegression || physicalDoubleClick || pictureRoutingTarget || pictureRoutingPerformance || process.argv.includes("--image"))
) {
  throw new Error(
    "--first-frame-artifacts cannot be combined with --create-image, --image, --physical-double-click or --picture-routing-target",
  );
}
const physicalOutputKind = physicalTarget.split("-", 1)[0];
const outputKind = createImageRegression || firstFrameImageRegression
  ? "image"
  : physicalDoubleClick
    ? physicalOutputKind
    : process.argv.includes("--image")
      ? "image"
      : "omml";
if (
  physicalDoubleClick &&
  process.argv.includes("--image") &&
  outputKind !== "image"
) {
  throw new Error("--image conflicts with the selected OMML physical target");
}
const referenceFontSizePt = 14;
const wordTexImageVisualScale = 1.1;
const wordTimesImageWidthScale = 1.067;
const wordTimesImageHeightScale = 1.0;
const wordDisplayPaddingPx = 2;

function wordImageVisualScalesForFont(formulaLetterFont = "katex") {
  return formulaLetterFont === "times"
    ? { width: wordTimesImageWidthScale, height: wordTimesImageHeightScale }
    : { width: wordTexImageVisualScale, height: wordTexImageVisualScale };
}
const nativeCalibrationWidthPt = 95.71632;
const editorReadyFileName = "editor-ready.json";
const editorPerformanceFileName = "editor-performance.jsonl";
const editorReadySchema = "visualtex-office-editor-ready-v1";
const editorPerformanceSchema = "visualtex-office-editor-performance-v1";
const warmEditorReadyLimitMs = 500;
const diagnosticSuccessPrefix = "VISUALTEX_DOCUMENT_IMPORT_DIAGNOSTIC_PASS:";
const transparentPng = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAEAQH/8l0Z8QAAAABJRU5ErkJggg==",
  "base64",
);
const sleep = (ms) => new Promise((resolvePromise) => setTimeout(resolvePromise, ms));

async function waitForWordAutomationReady(timeoutMs = 30_000) {
  spawnSync("/usr/bin/open", ["-a", "Microsoft Word"], {
    encoding: "utf8",
    timeout: 10_000,
  });
  const deadline = Date.now() + timeoutMs;
  let lastError = "";
  while (Date.now() < deadline) {
    try {
      const state = runAppleScript([
        'tell application "Microsoft Word"',
        "activate",
        "if not (exists active document) then make new document",
        "end tell",
        'tell application "System Events"',
        'tell process "Microsoft Word"',
        "set visible to true",
        "set frontmost to true",
        'if (count of menu bars) = 0 or (count of windows) = 0 then error "Word UI is not ready"',
        'return "READY"',
        "end tell",
        "end tell",
      ], 5_000);
      if (state.trim() === "READY") return;
    } catch (error) {
      lastError = error instanceof Error ? error.message : String(error);
      await sleep(500);
    }
  }
  throw new Error(
    `Microsoft Word did not become automation-ready after startup: ${lastError}`,
  );
}

function wordPictureFormatUiSnapshot() {
  const snapshot = JSON.parse(
    runJxa(
      String.raw`
const systemEvents = Application("System Events");
const word = systemEvents.processes.byName("Microsoft Word");
if (!word.exists()) {
  JSON.stringify({ processExists: false, lines: [] });
} else {
  word.frontmost = true;
  delay(0.45);
  const lines = [];
  let visited = 0;
  const text = (element, property) => {
    try {
      const value = element[property]();
      return value === null || value === undefined ? "" : String(value);
    } catch (_) {
      return "";
    }
  };
  const walk = (element, path, depth) => {
    if (visited >= 5000 || depth > 12) return;
    visited += 1;
    const role = text(element, "role");
    const name = text(element, "name");
    const description = text(element, "description");
    const help = text(element, "help");
    const value = text(element, "value");
    if (name || description || help || value) {
      lines.push([path, role, name, description, help, value].join("|"));
    }
    let children = [];
    try {
      children = element.uiElements();
    } catch (_) {}
    for (let index = 0; index < children.length; index += 1) {
      walk(children[index], path + "/" + index, depth + 1);
    }
  };
  const windows = word.windows();
  for (let index = 0; index < windows.length; index += 1) {
    walk(windows[index], "window" + index, 0);
  }
  JSON.stringify({ processExists: true, visited, lines });
}
`,
      45_000,
    ),
  );
  const lines = snapshot.lines ?? [];
  const raw = lines.join("\n");
  const taskPaneMarkers = [
    "调整图形格式",
    "设置图片格式",
    "adjust shape format",
    "format shape",
    "format picture",
  ];
  const paneMatches = lines.filter((line) => {
    const normalized = line.toLowerCase();
    const fields = line.split("|");
    const role = fields[1] ?? "";
    const description = fields[3] ?? "";
    const ribbonContextTab =
      role === "AXRadioButton" &&
      (description.includes("选项卡") ||
        description.toLowerCase().includes("tab"));
    return (
      !ribbonContextTab &&
      taskPaneMarkers.some((marker) =>
        normalized.includes(marker.toLowerCase()),
      )
    );
  });
  const contextualRibbonMatches = lines.filter((line) => {
    const fields = line.split("|");
    const role = fields[1] ?? "";
    const name = (fields[2] ?? "").toLowerCase();
    const description = (fields[3] ?? "").toLowerCase();
    return (
      role === "AXRadioButton" &&
      (description.includes("选项卡") || description.includes("tab")) &&
      (name.includes("图形格式") ||
        name.includes("图片格式") ||
        name.includes("shape format") ||
        name.includes("picture format"))
    );
  });
  return {
    raw,
    visited: snapshot.visited ?? 0,
    paneMatches,
    contextualRibbonMatches,
    pictureFormatVisible: paneMatches.length > 0,
  };
}

function closeWordPictureFormatTaskPaneViaAccessibility() {
  return JSON.parse(
    runJxa(
      String.raw`
const systemEvents = Application("System Events");
const word = systemEvents.processes.byName("Microsoft Word");
if (!word.exists()) {
  JSON.stringify({ clicked: false, reason: "word-not-running" });
} else {
  let clicked = false;
  let visited = 0;
  const text = (element, property) => {
    try {
      const value = element[property]();
      return value === null || value === undefined ? "" : String(value);
    } catch (_) {
      return "";
    }
  };
  const walk = (element, depth) => {
    if (clicked || visited >= 5000 || depth > 12) return;
    visited += 1;
    const role = text(element, "role");
    const name = text(element, "name").trim().toLowerCase();
    if (
      role === "AXButton" &&
      ((name.startsWith("关闭 ") && name.includes("格式")) ||
        (name.startsWith("close ") && name.includes("format")))
    ) {
      try {
        element.click();
        clicked = true;
        return;
      } catch (_) {}
    }
    let children = [];
    try {
      children = element.uiElements();
    } catch (_) {}
    for (let index = 0; index < children.length; index += 1) {
      walk(children[index], depth + 1);
      if (clicked) return;
    }
  };
  const windows = word.windows();
  for (let index = 0; index < windows.length; index += 1) {
    walk(windows[index], 0);
    if (clicked) break;
  }
  JSON.stringify({ clicked, visited });
}
`,
      45_000,
    ),
  );
}

function visualTeXEditorUiSnapshot() {
  const snapshot = JSON.parse(
    runJxa(
      String.raw`
const systemEvents = Application("System Events");
const visualTeX = systemEvents.processes.byName("visualtex");
if (!visualTeX.exists()) {
  JSON.stringify({ processExists: false, lines: [] });
} else {
  visualTeX.frontmost = true;
  delay(0.35);
  const lines = [];
  let visited = 0;
  const text = (element, property) => {
    try {
      const value = element[property]();
      return value === null || value === undefined ? "" : String(value);
    } catch (_) {
      return "";
    }
  };
  const walk = (element, path, depth) => {
    if (visited >= 5000 || depth > 12) return;
    visited += 1;
    const role = text(element, "role");
    const name = text(element, "name");
    const description = text(element, "description");
    const value = text(element, "value");
    if (name || description || value) {
      lines.push([path, role, name, description, value].join("|"));
    }
    let children = [];
    try {
      children = element.uiElements();
    } catch (_) {}
    for (let index = 0; index < children.length; index += 1) {
      walk(children[index], path + "/" + index, depth + 1);
    }
  };
  const windows = visualTeX.windows();
  for (let index = 0; index < windows.length; index += 1) {
    walk(windows[index], "window" + index, 0);
  }
  JSON.stringify({ processExists: true, windowCount: windows.length, visited, lines });
}
`,
      45_000,
    ),
  );
  const lines = snapshot.lines ?? [];
  const firstWindowLines = lines.filter((line) => line.startsWith("window0/"));
  const staticNames = firstWindowLines
    .filter((line) => line.split("|")[1] === "AXStaticText")
    .map((line) => line.split("|")[2] ?? "");
  const lineLabelIndex = staticNames.indexOf("行");
  const characterLabelIndex = staticNames.indexOf("字符");
  const lineCount =
    lineLabelIndex > 0 ? Number(staticNames[lineLabelIndex - 1]) : Number.NaN;
  const characterCount =
    characterLabelIndex > 0
      ? Number(staticNames[characterLabelIndex - 1])
      : Number.NaN;
  const errorMatches = firstWindowLines.filter((line) =>
    /missing\s*\\end|missing\s*end|无法打开|公式渲染失败|formula rendering failed/i.test(
      line,
    ),
  );
  return {
    windowCount: snapshot.windowCount ?? 0,
    visited: snapshot.visited ?? 0,
    lineCount,
    characterCount,
    errorMatches,
    raw: firstWindowLines.join("\n"),
  };
}

function closeVisualTeXEditorByWindowButton() {
  const status = runAppleScript([
    'tell application "System Events"',
    'if not (exists process "visualtex") then return "NO_PROCESS"',
    'tell process "visualtex"',
    'set frontmost to true',
    'if (count windows) = 0 then return "NO_WINDOW"',
    'if exists (first button of first window whose subrole is "AXCloseButton") then',
    'click (first button of first window whose subrole is "AXCloseButton")',
    'return "PRESSED"',
    'end if',
    'if (count buttons of first window) > 0 then',
    'click first button of first window',
    'return "PRESSED_FALLBACK"',
    'end if',
    'return "NO_BUTTON"',
    'end tell',
    'end tell',
  ], 30_000);
  return {
    pressed: status === "PRESSED" || status === "PRESSED_FALLBACK",
    status,
  };
}

async function waitForVisualTeXEditorToClose(timeoutMs = 45_000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    const state = JSON.parse(
      runJxa(
        String.raw`
const systemEvents = Application("System Events");
const visualTeX = systemEvents.processes.byName("visualtex");
JSON.stringify({
  processExists: visualTeX.exists(),
  windowCount: visualTeX.exists() ? visualTeX.windows().length : 0
});
`,
        15_000,
      ),
    );
    if (!state.processExists || state.windowCount === 0) return state;
    await sleep(250);
  }
  throw new Error("VisualTeX formula editor did not close after its native close button was pressed");
}

async function invokeWordFormatPictureCommand(testDocumentName) {
  const child = spawn(
    "/usr/bin/osascript",
    [
      "-e",
      'tell application "Microsoft Word"',
      "-e",
      `activate object document ${JSON.stringify(testDocumentName)}`,
      "-e",
      'run VB macro macro name "FormatPicture"',
      "-e",
      "end tell",
    ],
    { stdio: "ignore" },
  );
  await sleep(1_200);
  return child;
}

async function stopVisualTeXForManualWordCallback() {
  spawnSync("/usr/bin/killall", ["visualtex"], {
    encoding: "utf8",
    timeout: 10_000,
  });
  await sleep(1_000);
}

async function startVisualTeXForPhysicalRegression() {
  await stopVisualTeXForManualWordCallback();
  if (!existsSync(workspaceVisualTeXBinary)) {
    throw new Error(
      `The workspace VisualTeX validation app is missing: ${workspaceVisualTeXBinary}`,
    );
  }
  const child = spawn(workspaceVisualTeXBinary, ["--office-background"], {
    detached: true,
    stdio: "ignore",
  });
  child.unref();
  await sleep(4_000);
}

function visualTeXEditorWindowBounds() {
  const raw = runAppleScript([
    'tell application "System Events"',
    'tell process "visualtex"',
    'if (count of windows) is 0 then error "VisualTeX editor window is missing"',
    "set windowPosition to position of window 1",
    "set windowSize to size of window 1",
    'return (item 1 of windowPosition as text) & "|" & (item 2 of windowPosition as text) & "|" & (item 1 of windowSize as text) & "|" & (item 2 of windowSize as text)',
    "end tell",
    "end tell",
  ]);
  const [left, top, width, height] = raw.split("|").map(Number);
  if (
    [left, top, width, height].some((value) => !Number.isFinite(value)) ||
    width < 600 ||
    height < 500
  ) {
    throw new Error(`VisualTeX returned invalid editor bounds: ${raw}`);
  }
  return { left, top, width, height };
}

async function replaceActiveVisualTeXFormula(latex) {
  const clipboard = spawnSync("/usr/bin/pbcopy", [], {
    input: latex,
    encoding: "utf8",
    timeout: 5_000,
  });
  if (clipboard.status !== 0) {
    throw new Error(clipboard.stderr || "Unable to prepare formula clipboard");
  }
  const bounds = visualTeXEditorWindowBounds();
  runAppleScript([
    'tell application "System Events"',
    'tell process "visualtex"',
    "set frontmost to true",
    "delay 0.1",
    'keystroke "a" using {command down}',
    "delay 0.05",
    'keystroke "v" using {command down}',
    "end tell",
    "end tell",
  ]);
  await sleep(80);
  return { ...bounds, latex };
}

async function applyActiveVisualTeXFormula(sessionId, timeoutMs = 10_000) {
  const startedEpochMs = Date.now();
  runAppleScript([
    'tell application "System Events"',
    'tell process "visualtex"',
    "set frontmost to true",
    "key code 36 using {command down}",
    "end tell",
    "end tell",
  ]);
  const started = Date.now();
  let backendComplete = null;
  while (Date.now() - started < timeoutMs) {
    const records = editorPerformanceRecords(sessionId);
    backendComplete = records.find(
      (record) => record.stage === "apply-backend-complete",
    );
    if (backendComplete) break;
    await sleep(20);
  }
  if (!backendComplete) {
    throw new Error(
      `VisualTeX Apply did not complete for ${sessionId}: ${JSON.stringify(editorPerformanceRecords(sessionId))}`,
    );
  }
  const completedEpochMs = Number(backendComplete.epochMs);
  const clickToOfficeCompleteMs = completedEpochMs - startedEpochMs;
  if (
    !Number.isFinite(clickToOfficeCompleteMs) ||
    clickToOfficeCompleteMs < 0 ||
    clickToOfficeCompleteMs > timeoutMs
  ) {
    throw new Error(
      `VisualTeX Apply returned invalid timing: ${JSON.stringify({ startedEpochMs, backendComplete })}`,
    );
  }
  return {
    startedEpochMs,
    completedEpochMs,
    clickToOfficeCompleteMs,
    backendElapsedMs: backendComplete.elapsedMs,
    records: editorPerformanceRecords(sessionId).filter((record) =>
      String(record.stage).startsWith("apply-"),
    ),
  };
}

function physicallyDoubleClickAt(x, y, appKitY = true) {
  const argumentsList = [
    join(repositoryRoot, "scripts/macos_physical_double_click.swift"),
    x.toFixed(3),
    y.toFixed(3),
  ];
  if (appKitY) argumentsList.push("--appkit-y");
  const click = spawnSync("/usr/bin/swift", argumentsList, {
    encoding: "utf8",
    timeout: 30_000,
    maxBuffer: 2 * 1024 * 1024,
  });
  if (click.status !== 0) {
    throw new Error(
      click.stderr.trim() || click.stdout.trim() || "Quartz physical double-click failed",
    );
  }
  return click.stdout.trim();
}

function selectedWordFormulaScreenBounds(
  testDocumentName,
  boundsMacro = "VisualTeX_WriteSelectedScreenBoundsForRegression",
) {
  rmSync(physicalScreenBoundsPath, { force: true });
  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(testDocumentName)}`,
    "activate",
    `run VB macro macro name ${JSON.stringify(boundsMacro)}`,
    "end tell",
  ], 30_000);
  if (!existsSync(physicalScreenBoundsPath)) {
    throw new Error("Word did not write physical double-click screen bounds");
  }
  const status = readFileSync(physicalScreenBoundsPath, "utf8").trim();
  const [result, leftText, topText, widthText, heightText] = status.split("|");
  const values = [leftText, topText, widthText, heightText].map(Number);
  if (
    result !== "PASS" ||
    values.some((value) => !Number.isFinite(value)) ||
    values[2] <= 0 ||
    values[3] <= 0
  ) {
    throw new Error(`Word returned invalid physical screen bounds: ${status}`);
  }
  return values;
}

function wordFrontWindowScreenBounds() {
  const raw = runAppleScript([
    'tell application "System Events"',
    'tell process "Microsoft Word"',
    'if (count of windows) is 0 then error "Word window is missing"',
    "set windowPosition to position of window 1",
    "set windowSize to size of window 1",
    'return (item 1 of windowPosition as text) & "|" & (item 2 of windowPosition as text) & "|" & (item 1 of windowSize as text) & "|" & (item 2 of windowSize as text)',
    "end tell",
    "end tell",
  ], 10_000);
  const values = raw.split("|").map(Number);
  if (
    values.length !== 4 ||
    values.some((value) => !Number.isFinite(value)) ||
    values[2] <= 0 ||
    values[3] <= 0
  ) {
    throw new Error(`Word returned invalid window bounds: ${raw}`);
  }
  return values;
}

function physicallyDoubleClickSelectedWordFormula(
  testDocumentName,
  boundsMacro = "VisualTeX_WriteSelectedScreenBoundsForRegression",
) {
  const values = selectedWordFormulaScreenBounds(testDocumentName, boundsMacro);
  const clickX = values[0] + values[2] / 2;
  const clickY = values[1] + values[3] / 2;
  return {
    wordScreenCenterX: clickX,
    wordScreenCenterY: clickY,
    screenBounds: values,
    quartzResult: physicallyDoubleClickAt(clickX, clickY, true),
  };
}
const legacyAlignLatex = String.raw`\begin{align}
1 &= 22 + 333 \\
44444 &= 55
\end{align}`;
const legacyAlignStarLatex = String.raw`\begin{align*}
666 &= 777 + 8 \\
999999 &= 0
\end{align*}`;

function runJxa(source, timeout = 60_000) {
  const result = spawnSync(
    "/usr/bin/osascript",
    ["-l", "JavaScript", "-e", source],
    {
      encoding: "utf8",
      timeout,
      maxBuffer: 16 * 1024 * 1024,
    },
  );
  if (result.status !== 0) {
    const details = [
      result.stderr?.trim(),
      result.stdout?.trim(),
      result.error?.message,
      result.signal ? `signal=${result.signal}` : "",
      `status=${String(result.status)}`,
    ].filter(Boolean);
    throw new Error(details.join("\n") || "JXA failed");
  }
  return result.stdout.trim();
}

function runAppleScript(lines, timeout = 60_000) {
  const args = lines.flatMap((line) => ["-e", line]);
  const result = spawnSync("/usr/bin/osascript", args, {
    encoding: "utf8",
    timeout,
    maxBuffer: 8 * 1024 * 1024,
  });
  if (result.status !== 0) {
    const details = [
      result.stderr?.trim(),
      result.stdout?.trim(),
      result.error?.message,
      result.signal ? `signal=${result.signal}` : "",
      `status=${String(result.status)}`,
    ].filter(Boolean);
    throw new Error(details.join("\n") || "AppleScript failed");
  }
  return result.stdout.trim();
}

function startAppleScript(lines) {
  const args = lines.flatMap((line) => ["-e", line]);
  const child = spawn("/usr/bin/osascript", args, {
    stdio: ["ignore", "pipe", "pipe"],
  });
  let stdout = "";
  let stderr = "";
  let settled = false;
  child.stdout.setEncoding("utf8");
  child.stderr.setEncoding("utf8");
  child.stdout.on("data", (chunk) => {
    stdout += chunk;
  });
  child.stderr.on("data", (chunk) => {
    stderr += chunk;
  });
  const completion = new Promise((resolvePromise, rejectPromise) => {
    child.once("error", rejectPromise);
    child.once("close", (status, signal) => {
      settled = true;
      resolvePromise({ status, signal, stdout, stderr });
    });
  });
  return {
    child,
    completion,
    isSettled: () => settled,
  };
}

function readDocumentImportProgress(path) {
  if (!existsSync(path)) return null;
  const fields = new Map(
    readFileSync(path, "utf8")
      .split(/\r?\n/)
      .filter(Boolean)
      .map((line) => {
        const separator = line.indexOf("=");
        return separator < 0
          ? [line, ""]
          : [line.slice(0, separator), line.slice(separator + 1)];
      }),
  );
  const current = Number(fields.get("current"));
  const total = Number(fields.get("total"));
  const stage = fields.get("stage") ?? "";
  if (!Number.isInteger(current) || !Number.isInteger(total) || !stage) return null;
  return { current, total, stage };
}

function base64Url(value) {
  return Buffer.from(value, "utf8").toString("base64url");
}

const MATH_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/math";
const WORD_NAMESPACE =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function escapeXml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function ommlRun(text, align = false) {
  if (!text) return "";
  const runProperties = align
    ? '<m:rPr><m:scr m:val="roman"/><m:sty m:val="p"/></m:rPr>'
    : "";
  const equationArrayAlignment = align ? "&" : "";
  return `<m:r>${runProperties}<m:t>${escapeXml(`${equationArrayAlignment}${text}`)}</m:t></m:r>`;
}

function ommlBodyForLatex(latex, alignRelation = false) {
  const normalized = latex.replaceAll("&", "").trim();
  if (alignRelation) {
    const relationIndex = normalized.indexOf("=");
    if (relationIndex >= 0) {
      return (
        ommlRun(normalized.slice(0, relationIndex)) +
        ommlRun("=", true) +
        ommlRun(normalized.slice(relationIndex + 1))
      );
    }
  }
  const superscript = latex.match(/^(.*?)([A-Za-z])\^\{?(\d+)\}?$/);
  return superscript
    ? `<m:r><m:t>${escapeXml(superscript[1])}</m:t></m:r>` +
      `<m:sSup><m:e><m:r><m:t>${escapeXml(superscript[2])}</m:t></m:r></m:e>` +
      `<m:sup><m:r><m:t>${escapeXml(superscript[3])}</m:t></m:r></m:sup></m:sSup>`
    : ommlRun(normalized);
}

function ommlForFormula(lines, codeFormat) {
  const relationAligned = [
    "align",
    "align-star",
    "aligned",
    "equation-split",
    "equation-star-split",
  ].includes(codeFormat);
  const converted = lines.map((line) =>
    ommlBodyForLatex(line, relationAligned),
  );
  const body =
    converted.length === 1 && !relationAligned
      ? converted[0]
      : `<m:eqArr><m:eqArrPr><m:baseJc m:val="center"/></m:eqArrPr>${converted
          .map((line) => `<m:e>${line}</m:e>`)
          .join("")}</m:eqArr>`;
  return `<m:oMath xmlns:m="${MATH_NAMESPACE}" xmlns:w="${WORD_NAMESPACE}">${body}</m:oMath>`;
}

function minimalDocxBytes(omml) {
  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    "</Types>";
  const relationships =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    "</Relationships>";
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:m="${MATH_NAMESPACE}">` +
    `<w:body><w:p>${omml}</w:p><w:sectPr/></w:body></w:document>`;
  return zipSync(
    {
      "[Content_Types].xml": strToU8(contentTypes),
      "_rels/.rels": strToU8(relationships),
      "word/document.xml": strToU8(documentXml),
    },
    { level: 6 },
  );
}

function wordSvgDocxBytes(svg, png, widthPoints, heightPoints) {
  const widthEmu = Math.round(widthPoints * 12_700);
  const heightEmu = Math.round(heightPoints * 12_700);
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="svg" ContentType="image/svg+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
  const packageRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
  const documentRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPng" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/formula.png"/>
  <Relationship Id="rIdSvg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/formula.svg"/>
</Relationships>`;
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
  <w:body><w:p><w:r><w:drawing>
    <wp:inline distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="${widthEmu}" cy="${heightEmu}"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:docPr id="1" name="VisualTeX Formula" descr="VisualTeX SVG formula"/>
      <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:nvPicPr><pic:cNvPr id="0" name="formula.svg"/><pic:cNvPicPr/></pic:nvPicPr>
          <pic:blipFill><a:blip r:embed="rIdPng" cstate="print"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
          <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></pic:spPr>
        </pic:pic>
      </a:graphicData></a:graphic>
    </wp:inline>
  </w:drawing></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body>
</w:document>`;
  return zipSync(
    {
      "[Content_Types].xml": strToU8(contentTypes),
      "_rels/.rels": strToU8(packageRelationships),
      "word/document.xml": strToU8(documentXml),
      "word/_rels/document.xml.rels": strToU8(documentRelationships),
      "word/media/formula.png": new Uint8Array(png),
      "word/media/formula.svg": strToU8(svg),
    },
    { level: 6 },
  );
}

function calculateImageGeometry(svg, fontSizePt, formulaLetterFont = "katex") {
  const visualScale = wordImageVisualScalesForFont(formulaLetterFont);
  const naturalWidthPt = svg.width * 0.75 * visualScale.width;
  const naturalHeightPt = svg.height * 0.75 * visualScale.height;
  const referenceScale = Math.min(1, 500 / naturalWidthPt);
  const referenceWidthPt = naturalWidthPt * referenceScale;
  const referenceHeightPt = naturalHeightPt * referenceScale;
  const baselinePx = svg.baseline ?? svg.height;
  const descentRatio = Math.max(0, Math.min(1, (svg.height - baselinePx) / svg.height));
  const referenceBaselinePt = Math.max(
    -256,
    Math.min(0, -Math.max(0, referenceHeightPt * descentRatio)),
  );
  const pointScale = fontSizePt / referenceFontSizePt;
  return {
    widthPoints: referenceWidthPt * pointScale,
    heightPoints: referenceHeightPt * pointScale,
    baseline: Math.max(-256, Math.min(0, Math.round(referenceBaselinePt * pointScale))),
    referenceWidthPt,
    referenceHeightPt,
    referenceBaselinePt,
  };
}

function sha256(value) {
  return createHash("sha256").update(value).digest("hex");
}

function paethPredictor(left, above, upperLeft) {
  const estimate = left + above - upperLeft;
  const leftDistance = Math.abs(estimate - left);
  const aboveDistance = Math.abs(estimate - above);
  const upperLeftDistance = Math.abs(estimate - upperLeft);
  if (leftDistance <= aboveDistance && leftDistance <= upperLeftDistance) return left;
  return aboveDistance <= upperLeftDistance ? above : upperLeft;
}

function decodePngPixels(bytes, label) {
  const png = Buffer.from(bytes);
  if (
    png.length < 33 ||
    !png.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]))
  ) {
    throw new Error(`${label} is not a valid PNG`);
  }
  let offset = 8;
  let width = 0;
  let height = 0;
  let bitDepth = 0;
  let colorType = -1;
  let interlace = -1;
  let palette = null;
  let paletteAlpha = null;
  const compressed = [];
  while (offset + 12 <= png.length) {
    const length = png.readUInt32BE(offset);
    const type = png.subarray(offset + 4, offset + 8).toString("ascii");
    const dataStart = offset + 8;
    const dataEnd = dataStart + length;
    if (dataEnd + 4 > png.length) throw new Error(`${label} has a truncated ${type} chunk`);
    const data = png.subarray(dataStart, dataEnd);
    if (type === "IHDR") {
      width = data.readUInt32BE(0);
      height = data.readUInt32BE(4);
      bitDepth = data[8];
      colorType = data[9];
      interlace = data[12];
    } else if (type === "PLTE") {
      palette = data;
    } else if (type === "tRNS") {
      paletteAlpha = data;
    } else if (type === "IDAT") {
      compressed.push(data);
    } else if (type === "IEND") {
      break;
    }
    offset = dataEnd + 4;
  }
  if (
    width <= 1 ||
    height <= 1 ||
    bitDepth !== 8 ||
    interlace !== 0 ||
    ![0, 2, 3, 4, 6].includes(colorType) ||
    compressed.length === 0
  ) {
    throw new Error(
      `${label} uses an unsupported PNG layout: ${JSON.stringify({ width, height, bitDepth, colorType, interlace })}`,
    );
  }
  const bytesPerPixel = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[colorType];
  const stride = width * bytesPerPixel;
  const inflated = inflateSync(Buffer.concat(compressed));
  if (inflated.length !== (stride + 1) * height) {
    throw new Error(`${label} PNG scanline size is inconsistent`);
  }
  const reconstructed = Buffer.alloc(stride * height);
  let sourceOffset = 0;
  for (let row = 0; row < height; row += 1) {
    const filter = inflated[sourceOffset];
    sourceOffset += 1;
    const rowOffset = row * stride;
    for (let column = 0; column < stride; column += 1) {
      const raw = inflated[sourceOffset + column];
      const left = column >= bytesPerPixel
        ? reconstructed[rowOffset + column - bytesPerPixel]
        : 0;
      const above = row > 0 ? reconstructed[rowOffset - stride + column] : 0;
      const upperLeft = row > 0 && column >= bytesPerPixel
        ? reconstructed[rowOffset - stride + column - bytesPerPixel]
        : 0;
      let value;
      if (filter === 0) value = raw;
      else if (filter === 1) value = raw + left;
      else if (filter === 2) value = raw + above;
      else if (filter === 3) value = raw + Math.floor((left + above) / 2);
      else if (filter === 4) value = raw + paethPredictor(left, above, upperLeft);
      else throw new Error(`${label} uses unknown PNG filter ${filter}`);
      reconstructed[rowOffset + column] = value & 0xff;
    }
    sourceOffset += stride;
  }

  const rgba = Buffer.alloc(width * height * 4);
  let visiblePixels = 0;
  let darkVisiblePixels = 0;
  let minInkX = width;
  let maxInkX = -1;
  let minInkY = height;
  let maxInkY = -1;
  for (let pixel = 0; pixel < width * height; pixel += 1) {
    const source = pixel * bytesPerPixel;
    let red;
    let green;
    let blue;
    let alpha;
    if (colorType === 0) {
      red = green = blue = reconstructed[source];
      alpha = 255;
    } else if (colorType === 2) {
      red = reconstructed[source];
      green = reconstructed[source + 1];
      blue = reconstructed[source + 2];
      alpha = 255;
    } else if (colorType === 3) {
      const paletteIndex = reconstructed[source];
      if (!palette || paletteIndex * 3 + 2 >= palette.length) {
        throw new Error(`${label} references an invalid PNG palette entry`);
      }
      red = palette[paletteIndex * 3];
      green = palette[paletteIndex * 3 + 1];
      blue = palette[paletteIndex * 3 + 2];
      alpha = paletteAlpha?.[paletteIndex] ?? 255;
    } else if (colorType === 4) {
      red = green = blue = reconstructed[source];
      alpha = reconstructed[source + 1];
    } else {
      red = reconstructed[source];
      green = reconstructed[source + 1];
      blue = reconstructed[source + 2];
      alpha = reconstructed[source + 3];
    }
    const target = pixel * 4;
    rgba[target] = red;
    rgba[target + 1] = green;
    rgba[target + 2] = blue;
    rgba[target + 3] = alpha;
    if (alpha >= 16) {
      visiblePixels += 1;
      const x = pixel % width;
      const y = Math.floor(pixel / width);
      minInkX = Math.min(minInkX, x);
      maxInkX = Math.max(maxInkX, x);
      minInkY = Math.min(minInkY, y);
      maxInkY = Math.max(maxInkY, y);
      if (red < 245 || green < 245 || blue < 245) darkVisiblePixels += 1;
    }
  }
  if (darkVisiblePixels === 0) {
    throw new Error(`${label} contains no dark non-transparent formula pixels`);
  }
  return {
    width,
    height,
    visiblePixels,
    darkVisiblePixels,
    coverage: visiblePixels / (width * height),
    inkBounds: {
      minX: minInkX,
      maxX: maxInkX,
      minY: minInkY,
      maxY: maxInkY,
      width: maxInkX - minInkX + 1,
      height: maxInkY - minInkY + 1,
    },
    pixelHash: sha256(rgba),
  };
}

function assertEquivalentPngArtwork(frontend, final, label) {
  if (frontend.pixelHash === final.pixelHash) return { mode: "exact-pixels" };
  const widthScale = final.width / frontend.width;
  const heightScale = final.height / frontend.height;
  const frontendAspect = frontend.inkBounds.width / frontend.inkBounds.height;
  const finalAspect = final.inkBounds.width / final.inkBounds.height;
  const normalizedBounds = {
    minX: Math.abs(frontend.inkBounds.minX / frontend.width - final.inkBounds.minX / final.width),
    maxX: Math.abs(frontend.inkBounds.maxX / frontend.width - final.inkBounds.maxX / final.width),
    minY: Math.abs(frontend.inkBounds.minY / frontend.height - final.inkBounds.minY / final.height),
    maxY: Math.abs(frontend.inkBounds.maxY / frontend.height - final.inkBounds.maxY / final.height),
  };
  const coverageDifference = Math.abs(frontend.coverage - final.coverage);
  const aspectDifference = Math.abs(frontendAspect - finalAspect) / frontendAspect;
  if (
    widthScale < 0.5 ||
    widthScale > 4 ||
    Math.abs(widthScale - heightScale) > 0.08 ||
    Object.values(normalizedBounds).some((difference) => difference > 0.04) ||
    coverageDifference > 0.08 ||
    aspectDifference > 0.08
  ) {
    throw new Error(
      `${label} PNG cache does not preserve the frontend artwork: ${JSON.stringify({ frontend, final, widthScale, heightScale, normalizedBounds, coverageDifference, aspectDifference })}`,
    );
  }
  return {
    mode: "word-regenerated-scaled-cache",
    widthScale,
    heightScale,
    normalizedBounds,
    coverageDifference,
    aspectDifference,
  };
}

function assertWordCompatibleSvgMarkup(svg, label) {
  if (!/(?:fill|stroke)=["']#000000["']/i.test(svg)) {
    throw new Error(`${label} has no explicit #000000 formula paint`);
  }
  if (
    /currentColor|var\(|(?:fill|stroke|color)\s*[:=]\s*["']?(?:inherit|white|#fff(?:fff)?)/i.test(
      svg,
    )
  ) {
    throw new Error(`${label} retains currentColor, CSS variables, inheritance or white paint`);
  }
}

function browserFormulaArtifact(formula, artifactDirectory, index) {
  if (
    !formula?.formulaId ||
    !formula.svgBase64 ||
    !formula.pngBase64 ||
    !formula.ommlBase64 ||
    !formula.ommlDocxBase64 ||
    !formula.metadata ||
    !Number.isFinite(formula.width) ||
    !Number.isFinite(formula.height)
  ) {
    throw new Error(`Browser formula ${index + 1} is incomplete`);
  }
  const svgBytes = Buffer.from(formula.svgBase64, "base64");
  const pngBytes = Buffer.from(formula.pngBase64, "base64");
  const svgMarkup = svgBytes.toString("utf8");
  assertWordCompatibleSvgMarkup(svgMarkup, `Browser formula ${index + 1} SVG`);
  const frontendPng = decodePngPixels(pngBytes, `Browser formula ${index + 1} PNG`);
  const geometry = calculateImageGeometry(
    {
      width: formula.width,
      height: formula.height,
      baseline: formula.baseline ?? formula.height,
    },
    formula.fontSizePt,
    formula.metadata?.formulaLetterFont ?? "katex",
  );
  const stem = `first-frame-${index + 1}-${compactFormulaId(formula.formulaId)}`;
  const imagePath = join(artifactDirectory, `${stem}.svg`);
  const fallbackImagePath = join(artifactDirectory, `${stem}.png`);
  const vectorDocumentPath = join(artifactDirectory, `${stem}-svg.docx`);
  const nativePath = join(nativeRoot, `${formula.formulaId}.docx`);
  writeFileSync(imagePath, svgBytes, { mode: 0o600 });
  writeFileSync(fallbackImagePath, pngBytes, { mode: 0o600 });
  writeFileSync(
    vectorDocumentPath,
    wordSvgDocxBytes(svgMarkup, pngBytes, geometry.widthPoints, geometry.heightPoints),
    { mode: 0o600 },
  );
  writeFileSync(nativePath, Buffer.from(formula.ommlDocxBase64, "base64url"), {
    mode: 0o600,
  });
  return {
    formulaId: formula.formulaId,
    latex: formula.latex,
    metadata: encodeFormulaMetadata(formula.metadata),
    metadataLines: formula.metadata.lines?.map((line) => line.latex) ?? [],
    codeFormat: formula.metadata.codeFormat,
    displayMode: formula.displayMode,
    numbered: Boolean(formula.numbered),
    fontSizePt: Number(formula.fontSizePt),
    ommlBase64: formula.ommlBase64,
    nativePath,
    imagePath,
    vectorDocumentPath,
    fallbackImagePath,
    widthPoints: geometry.widthPoints,
    heightPoints: geometry.heightPoints,
    baseline: geometry.baseline,
    referenceWidthPt: geometry.referenceWidthPt,
    referenceHeightPt: geometry.referenceHeightPt,
    referenceBaselinePt: geometry.referenceBaselinePt,
    frontendPngHash: sha256(pngBytes),
    frontendPngPixelHash: frontendPng.pixelHash,
    frontendSvgHash: sha256(svgBytes),
    frontendPng,
  };
}

function relationshipMap(xml) {
  const relationships = new Map();
  for (const match of xml.matchAll(
    /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/?\s*>/g,
  )) {
    relationships.set(match[1], match[2]);
  }
  return relationships;
}

function inspectSavedWordImagePackage(docxPath, formulas, stage) {
  const archive = unzipSync(readFileSync(docxPath));
  const readEntry = (path) => {
    const value = archive[path];
    if (!value) throw new Error(`${stage} DOCX is missing ${path}`);
    return Buffer.from(value);
  };
  const documentXml = readEntry("word/document.xml").toString("utf8");
  const relationshipsXml = readEntry("word/_rels/document.xml.rels").toString("utf8");
  const relationships = relationshipMap(relationshipsXml);
  const pairs = [];
  for (const block of documentXml.matchAll(/<a:blip\b[\s\S]*?<\/a:blip>/g)) {
    const pngRelationship = block[0].match(/\br:embed="([^"]+)"/i)?.[1];
    const svgRelationship = block[0].match(
      /<asvg:svgBlip\b[^>]*\br:embed="([^"]+)"/i,
    )?.[1];
    if (pngRelationship && svgRelationship) {
      pairs.push({ pngRelationship, svgRelationship });
    }
  }
  if (pairs.length !== formulas.length) {
    throw new Error(
      `${stage} DOCX has ${pairs.length} SVG/PNG relationship pairs for ${formulas.length} formulas`,
    );
  }
  const formulaReports = pairs.map((pair, index) => {
    const formula = formulas[index];
    const pngTarget = relationships.get(pair.pngRelationship);
    const svgTarget = relationships.get(pair.svgRelationship);
    if (!pngTarget || !svgTarget) {
      throw new Error(`${stage} formula ${index + 1} has unresolved image relationships`);
    }
    const pngPath = `word/${pngTarget.replace(/^\.\//, "")}`;
    const svgPath = `word/${svgTarget.replace(/^\.\//, "")}`;
    const finalPngBytes = readEntry(pngPath);
    const finalSvgBytes = readEntry(svgPath);
    const finalPng = decodePngPixels(finalPngBytes, `${stage} formula ${index + 1} PNG`);
    const finalSvg = finalSvgBytes.toString("utf8");
    assertWordCompatibleSvgMarkup(finalSvg, `${stage} formula ${index + 1} SVG`);
    const finalPngHash = sha256(finalPngBytes);
    const finalSvgHash = sha256(finalSvgBytes);
    const pngByteHashMatches = finalPngHash === formula.frontendPngHash;
    const pngPixelHashMatches = finalPng.pixelHash === formula.frontendPngPixelHash;
    const pngArtworkComparison = assertEquivalentPngArtwork(
      formula.frontendPng,
      finalPng,
      `${stage} formula ${index + 1}`,
    );
    const svgByteHashMatches = finalSvgHash === formula.frontendSvgHash;
    return {
      formulaId: formula.formulaId,
      pngPath,
      svgPath,
      pngRelationship: pair.pngRelationship,
      svgRelationship: pair.svgRelationship,
      frontendPngHash: formula.frontendPngHash,
      pngHash: finalPngHash,
      pngByteHashMatches,
      pngPixelHashMatches,
      pngArtworkComparison,
      frontendSvgHash: formula.frontendSvgHash,
      svgHash: finalSvgHash,
      svgByteHashMatches,
      png: finalPng,
    };
  });
  return {
    stage,
    alternateContentCount: (documentXml.match(/<mc:AlternateContent\b/g) ?? []).length,
    blipCount: (documentXml.match(/<a:blip\b/g) ?? []).length,
    svgBlipCount: (documentXml.match(/<asvg:svgBlip\b/g) ?? []).length,
    formulaReports,
  };
}

function svgRelationshipPositions(svgMarkup) {
  const relationshipPattern =
    /<g data-mml-node="mtd" transform="translate\(([-+\d.]+),[-+\d.]+\)">(?:(?!<g data-mml-node="mtd")[\s\S])*?<g data-mml-node="mo" transform="translate\(([-+\d.]+),[-+\d.]+\)"><use data-c="3D"/g;
  return [...svgMarkup.matchAll(relationshipPattern)].map(
    (match) => Number(match[1]) + Number(match[2]),
  );
}

function assertAlignedSvg(svgMarkup, expectedRows, label) {
  const positions = svgRelationshipPositions(svgMarkup);
  if (
    positions.length !== expectedRows ||
    positions.some((position) => !Number.isFinite(position))
  ) {
    throw new Error(
      `${label} did not expose ${expectedRows} SVG relationship positions: ${JSON.stringify(positions)}`,
    );
  }
  const spread = Math.max(...positions) - Math.min(...positions);
  if (spread > 0.01) {
    throw new Error(
      `${label} SVG relationship columns are not aligned: ${JSON.stringify({ positions, spread })}`,
    );
  }
  return { positions, spread };
}

function rasterBounds(band, components) {
  const minX = Math.min(...components.map((component) => component.minX));
  const maxX = Math.max(...components.map((component) => component.maxX));
  return {
    minX,
    maxX,
    width: maxX - minX,
    centerX: (minX + maxX) / 2,
    minY: band.minY,
    maxY: band.maxY,
    height: band.height,
    centerY: band.centerY,
    components,
  };
}

function rasterEntryBounds(entries) {
  if (!entries.length) return null;
  const components = entries.map(({ component }) => component);
  const minX = Math.min(...components.map((component) => component.minX));
  const maxX = Math.max(...components.map((component) => component.maxX));
  const minY = Math.min(...entries.map(({ band }) => band.minY));
  const maxY = Math.max(...entries.map(({ band }) => band.maxY));
  return {
    minX,
    maxX,
    width: maxX - minX,
    centerX: (minX + maxX) / 2,
    minY,
    maxY,
    height: maxY - minY,
    centerY: (minY + maxY) / 2,
    components,
  };
}

function resolveImageRasterGeometry(
  rasterBands,
  textBoundaryCenter,
  wordGeometry = {},
) {
  const requiredGeometry = [
    "displayTop",
    "displayHeight",
    "numberedTop",
    "numberedHeight",
  ];
  if (
    requiredGeometry.some(
      (key) => !Number.isFinite(wordGeometry[key]) || wordGeometry[key] < 0,
    )
  ) {
    throw new Error(
      `Image raster geometry is missing Word formula bounds: ${JSON.stringify(wordGeometry)}`,
    );
  }

  const componentsInVerticalBox = (top, height) => {
    const bottom = top + height;
    const tolerance = 1.5;
    return rasterBands.flatMap((band) =>
      band.maxY >= top - tolerance && band.minY <= bottom + tolerance
        ? band.components.map((component) => ({ band, component }))
        : [],
    );
  };
  const centeredMarker = (entries, label, excludedComponent = null) => {
    const candidates = entries
      .filter(({ component }) => component !== excludedComponent)
      .map((entry) => ({
        ...entry,
        centerError: Math.abs(entry.component.centerX - textBoundaryCenter),
      }))
      .filter(({ centerError }) => centerError <= 8)
      .sort(
        (left, right) =>
          left.centerError - right.centerError ||
          right.component.width - left.component.width,
      );
    const best = candidates[0];
    if (!best) {
      throw new Error(
        `Unable to locate ${label} image formula center marker: ${JSON.stringify({ rasterBands, wordGeometry, textBoundaryCenter })}`,
      );
    }
    return rasterBounds(best.band, [best.component]);
  };

  const numberedEntries = componentsInVerticalBox(
    wordGeometry.numberedTop,
    wordGeometry.numberedHeight,
  );
  const numberCandidates = numberedEntries
    .filter(
      ({ component }) => component.minX > textBoundaryCenter + 100,
    )
    .sort(
      (left, right) => right.component.centerX - left.component.centerX,
    );
  const numberEntry = numberCandidates[0];
  if (!numberEntry) {
    throw new Error(
      `Unable to locate numbered image formula raster number: ${JSON.stringify({ rasterBands, wordGeometry, textBoundaryCenter })}`,
    );
  }

  const unnumberedEntries = componentsInVerticalBox(
    wordGeometry.displayTop,
    wordGeometry.displayHeight,
  );
  const numberedFormulaEntries = numberedEntries.filter(
    ({ component }) => component !== numberEntry.component,
  );
  const unnumbered = centeredMarker(
    unnumberedEntries,
    "unnumbered",
  );
  const numbered = centeredMarker(
    numberedFormulaEntries,
    "numbered",
  );
  return {
    unnumbered,
    numbered,
    unnumberedInk: rasterEntryBounds(unnumberedEntries),
    numberedInk: rasterEntryBounds(numberedFormulaEntries),
    equationNumber: {
      ...numberEntry.component,
      minY: numberEntry.band.minY,
      maxY: numberEntry.band.maxY,
      height: numberEntry.band.height,
      centerY: numberEntry.band.centerY,
    },
  };
}

function manifestText(entries) {
  const seen = new Set();
  return entries
    .map(([key, value]) => {
      if (!/^[A-Za-z0-9]+$/.test(key) || seen.has(key)) {
        throw new Error(`Invalid integration manifest key: ${key}`);
      }
      seen.add(key);
      const text = String(value);
      if (/[\r\n\0]/.test(text)) {
        throw new Error(`Invalid integration manifest value for ${key}`);
      }
      return `${key}=${text}`;
    })
    .join("\n") + "\n";
}

function parseRegressionReport(reportText) {
  const lines = reportText.trim().split(/\r?\n/);
  if (lines[0] !== "PASS") {
    throw new Error(`Word formula regression failed: ${reportText}`);
  }
  return Object.fromEntries(
    lines.slice(1).map((line) => {
      const separator = line.indexOf("=");
      if (separator <= 0) {
        throw new Error(`Invalid Word formula regression line: ${line}`);
      }
      return [line.slice(0, separator), line.slice(separator + 1)];
    }),
  );
}

function numericReportValue(report, key) {
  const value = Number(report[key]);
  if (!Number.isFinite(value)) {
    throw new Error(`Word formula regression omitted ${key}: ${JSON.stringify(report)}`);
  }
  return value;
}

function runFormulaRegressionReport(testDocumentName, formulas) {
  const formulaCount = formulas.length;
  const displayFormulaCount = formulas.filter(
    (formula) => formula.displayMode === "block",
  ).length;
  const alignedFormulaCountExpected = createNativeNumberedRegression
    ? 1
    : formulas.filter((formula) =>
        ["align", "align-star"].includes(formula.codeFormat),
      ).length;
  rmSync(formulaRegressionStatusPath, { force: true });
  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(testDocumentName)}`,
    'run VB macro macro name "VisualTeX_RunDocumentImportFormulaRegression"',
    "end tell",
  ], 60_000);
  if (!existsSync(formulaRegressionStatusPath)) {
    throw new Error("Word did not write the document-import formula regression report");
  }
  const report = parseRegressionReport(
    readFileSync(formulaRegressionStatusPath, "utf8"),
  );
  if (
    report.revision !==
    "word-office-performance-20260801-r77"
  ) {
    throw new Error(`Word loaded the wrong VisualTeX source revision: ${report.revision}`);
  }

  const documentMathCount = numericReportValue(report, "documentMathCount");
  const nativeFormulaCount = numericReportValue(report, "nativeFormulaCount");
  const nativeDisplayCount = numericReportValue(report, "nativeDisplayCount");
  const invalidNativeRangeCount = numericReportValue(
    report,
    "invalidNativeRangeCount",
  );
  const emptyMathCount = numericReportValue(report, "emptyMathCount");
  const alignedFormulaCount = numericReportValue(report, "alignedFormulaCount");
  const imageFormulaCount = numericReportValue(report, "imageFormulaCount");
  const imageDisplayCount = numericReportValue(report, "imageDisplayCount");
  const imageMacroButtonCount = numericReportValue(
    report,
    "imageMacroButtonCount",
  );
  const invalidImageMacroButtonCount = numericReportValue(
    report,
    "invalidImageMacroButtonCount",
  );
  const imageFormulaIds = (report.imageFormulaIds ?? "")
    .split(",")
    .filter(Boolean);
  const maximumImageSpaceBefore = numericReportValue(
    report,
    "maximumImageSpaceBefore",
  );
  const maximumImageSpaceAfter = numericReportValue(
    report,
    "maximumImageSpaceAfter",
  );

  if (outputKind === "omml") {
    if (
      documentMathCount !== formulaCount ||
      nativeFormulaCount !== formulaCount ||
      nativeDisplayCount !== displayFormulaCount ||
      invalidNativeRangeCount !== 0 ||
      emptyMathCount !== 0 ||
      alignedFormulaCount !== alignedFormulaCountExpected ||
      imageFormulaCount !== 0 ||
      imageMacroButtonCount !== 0 ||
      invalidImageMacroButtonCount !== 0
    ) {
      throw new Error(
        `Word OMML formula structure regression failed: ${JSON.stringify(report)}`,
      );
    }
    const nativeSpacingExpectations = createNativeNumberedRegression
      ? [
          ["minimumNativeSpaceBefore", 0],
          ["maximumNativeSpaceBefore", 0],
          ["minimumNativeSpaceAfter", 0],
          ["maximumNativeSpaceAfter", 0],
        ]
      : createNativeDisplayRegression
      ? [
          // A freshly promoted unnumbered native display equation follows the
          // host document's Normal paragraph spacing. The default Word document
          // used by this create regression reports 0 pt before / 8 pt after.
          // Keep this distinct from image displays, which VisualTeX explicitly
          // normalizes to zero because the image already owns its visual height.
          ["minimumNativeSpaceBefore", 0],
          ["maximumNativeSpaceBefore", 0],
          ["minimumNativeSpaceAfter", 8],
          ["maximumNativeSpaceAfter", 8],
        ]
      : createNativeInlineRegression
        ? [
            ["minimumNativeSpaceBefore", 0],
            ["maximumNativeSpaceBefore", 0],
            ["minimumNativeSpaceAfter", 0],
            ["maximumNativeSpaceAfter", 0],
          ]
        : [
          // Numbered native equations retain their existing zero-spaced table
          // layout; unnumbered native equations inherit the configured Normal
          // style spacing. Image-only normalization must not change either case.
          ["minimumNativeSpaceBefore", 0],
          ["maximumNativeSpaceBefore", 6],
          ["minimumNativeSpaceAfter", 0],
          ["maximumNativeSpaceAfter", 9],
        ];
    for (const [key, expected] of nativeSpacingExpectations) {
      if (Math.abs(numericReportValue(report, key) - expected) > 0.05) {
        throw new Error(
          `OMML paragraph spacing changed for ${key}: ${JSON.stringify(report)}`,
        );
      }
    }
  } else if (
    documentMathCount !== 0 ||
    nativeFormulaCount !== 0 ||
    nativeDisplayCount !== 0 ||
    invalidNativeRangeCount !== 0 ||
    emptyMathCount !== 0 ||
    alignedFormulaCount !== 0 ||
    imageFormulaCount !== formulaCount ||
    imageDisplayCount !== displayFormulaCount ||
    imageMacroButtonCount !== 0 ||
    invalidImageMacroButtonCount !== 0 ||
    JSON.stringify(imageFormulaIds) !==
      JSON.stringify(formulas.map((formula) => formula.formulaId)) ||
    maximumImageSpaceBefore > 0.01 ||
    maximumImageSpaceAfter > 0.01
  ) {
    throw new Error(
      `Word image formula structure regression failed: ${JSON.stringify(report)}`,
    );
  }
  return report;
}

function compactFormulaId(id) {
  return id.replaceAll("-", "");
}

function nativeBookmark(id) {
  return `VT_F_${compactFormulaId(id)}`;
}

function currentSessionIds() {
  return new Set(
    existsSync(sessionsRoot)
      ? readdirSync(sessionsRoot, { withFileTypes: true })
          .filter((entry) => entry.isDirectory())
          .map((entry) => entry.name)
      : [],
  );
}

function sessionIdsAddedAfter(before) {
  return [...currentSessionIds()].filter((sessionId) => !before.has(sessionId));
}

async function assertNoNewWordEditSession(before, label, waitMs = 1_200) {
  await sleep(waitMs);
  const added = sessionIdsAddedAfter(before);
  if (added.length > 0) {
    throw new Error(`${label} unexpectedly created VisualTeX Session(s): ${added.join(",")}`);
  }
}

function inspectWordFormulaContainers(testDocumentName, formulas, stage) {
  const report = runFormulaRegressionReport(testDocumentName, formulas);
  if (outputKind === "omml") {
    return {
      stage,
      inlineShapeCount: 0,
      macroButtonCount: 0,
      shapes: [],
    };
  }
  const inspection = runAppleScript([
    'tell application "Microsoft Word"',
    `set documentObject to document ${JSON.stringify(testDocumentName)}`,
    "set unitSeparator to ASCII character 31",
    "set recordSeparator to ASCII character 30",
    "set macroButtonCount to 0",
    "set documentFieldCount to count fields of documentObject",
    "repeat with fieldIndex from 1 to documentFieldCount",
    "set candidateField to field fieldIndex of documentObject",
    "if field type of candidateField is field macro button then set macroButtonCount to macroButtonCount + 1",
    "end repeat",
    "set reportText to (macroButtonCount as text) & unitSeparator & ((count of inline shapes of documentObject) as text)",
    "repeat with shapeIndex from 1 to count of inline shapes of documentObject",
    "set formulaShape to inline shape shapeIndex of documentObject",
    "set shapeRange to text object of formulaShape",
    "set shapeStart to start of content of shapeRange",
    "set shapeEnd to end of content of shapeRange",
    "set formulaParagraph to paragraph 1 of (create range documentObject start shapeStart end shapeStart)",
    "set paragraphRange to text object of formulaParagraph",
    "set paragraphStart to start of content of paragraphRange",
    "set paragraphEnd to end of content of paragraphRange",
    "set paragraphText to content of paragraphRange",
    "set paragraphFieldCount to count of fields of paragraphRange",
    "set metadataText to alternative text of formulaShape",
    "set reportText to reportText & recordSeparator & (shapeIndex as text) & unitSeparator & (shapeStart as text) & unitSeparator & (shapeEnd as text) & unitSeparator & (paragraphStart as text) & unitSeparator & (paragraphEnd as text) & unitSeparator & paragraphText & unitSeparator & (paragraphFieldCount as text) & unitSeparator & metadataText",
    "end repeat",
    "return reportText",
    "end tell",
  ]);
  const [summary, ...recordTexts] = inspection.split("\x1e");
  const [macroButtonCount, inlineShapeCount] = summary
    .split("\x1f")
    .map(Number);
  if (
    !Number.isInteger(macroButtonCount) ||
    macroButtonCount < 0 ||
    !Number.isInteger(inlineShapeCount) ||
    inlineShapeCount < 0
  ) {
    throw new Error(
      `${stage} returned an invalid Word field summary: ${JSON.stringify(inspection)}`,
    );
  }

  if (outputKind === "omml") {
    if (inlineShapeCount !== 0 || macroButtonCount !== 0 || recordTexts.length) {
      throw new Error(
        `${stage} OMML document contains image/MacroButton objects: ${JSON.stringify({
          inlineShapeCount,
          macroButtonCount,
          recordTexts,
        })}`,
      );
    }
    return { stage, inlineShapeCount, macroButtonCount, shapes: [] };
  }

  if (
    inlineShapeCount !== formulas.length ||
    macroButtonCount !== 0 ||
    recordTexts.length !== formulas.length
  ) {
    throw new Error(
      `${stage} did not retain plain field-free VisualTeX images: ${JSON.stringify({
        expected: formulas.length,
        inlineShapeCount,
        macroButtonCount,
        recordCount: recordTexts.length,
      })}`,
    );
  }
  const shapes = recordTexts.map((recordText, index) => {
    const [
      shapeIndexText,
      shapeStartText,
      shapeEndText,
      paragraphStartText,
      paragraphEndText,
      paragraphText,
      paragraphFieldCountText,
      encodedMetadata,
    ] = recordText.split("\x1f");
    const record = {
      shapeIndex: Number(shapeIndexText),
      shapeStart: Number(shapeStartText),
      shapeEnd: Number(shapeEndText),
      paragraphStart: Number(paragraphStartText),
      paragraphEnd: Number(paragraphEndText),
      paragraphText,
      paragraphFieldCount: Number(paragraphFieldCountText),
    };
    const expected = formulas[index];
    const metadata = decodeFormulaMetadata(encodedMetadata ?? "");
    if (
      record.shapeIndex !== index + 1 ||
      !Number.isInteger(record.shapeStart) ||
      !Number.isInteger(record.shapeEnd) ||
      !Number.isInteger(record.paragraphStart) ||
      !Number.isInteger(record.paragraphEnd) ||
      record.shapeStart >= record.shapeEnd ||
      record.paragraphStart > record.shapeStart ||
      record.paragraphEnd < record.shapeEnd ||
      record.paragraphFieldCount !== 0 ||
      metadata?.formulaId !== expected.formulaId
    ) {
      throw new Error(
        `${stage} image ${index + 1} is not one plain field-free ` +
          `VisualTeX InlineShape: ${JSON.stringify({ record, metadata })}`,
      );
    }
    return { ...record, formulaId: metadata.formulaId };
  });

  const paragraphGroups = new Map();
  for (const shape of shapes) {
    const key = `${shape.paragraphStart}:${shape.paragraphEnd}`;
    const group = paragraphGroups.get(key) ?? [];
    group.push(shape);
    paragraphGroups.set(key, group);
  }
  const normalizedParagraphText = (value, imageCount) => {
    let text = String(value ?? "").replace(
      /[\u0001\u0007\u0015\t\r\n\u00a0\u200b\u2060 ]/g,
      "",
    );
    // Word AppleScript exposes each plain InlineShape as one slash in Range.Text.
    // Remove exactly the known image-object placeholders, never arbitrary X/text.
    for (let index = 0; index < imageCount; index += 1) {
      text = text.replace("/", "");
    }
    return text;
  };

  const structuredShapes = shapes.map((shape, index) => {
    const formula = formulas[index];
    const key = `${shape.paragraphStart}:${shape.paragraphEnd}`;
    const paragraphShapes = paragraphGroups.get(key) ?? [];
    const visibleText = normalizedParagraphText(
      shape.paragraphText,
      paragraphShapes.length,
    );
    const paragraphModes = paragraphShapes.map(
      (paragraphShape) => formulas[paragraphShape.shapeIndex - 1]?.displayMode,
    );

    if (formula.displayMode === "block") {
      const validNumberText = /^\([^()]+\)$/.test(visibleText);
      const validDisplayStructure =
        paragraphShapes.length === 1 &&
        paragraphModes.every((mode) => mode === "block") &&
        (formula.numbered
          ? validNumberText
          : visibleText === "");
      if (!validDisplayStructure) {
        throw new Error(
          `${stage} display image ${index + 1} is not isolated in its own ` +
            `Word paragraph: ${JSON.stringify({ shape, paragraphShapes, visibleText })}`,
        );
      }
      return {
        ...shape,
        layoutStructure: formula.numbered
          ? "numbered-display-paragraph"
          : "dedicated-display-paragraph",
        paragraphFormulaCount: paragraphShapes.length,
        visibleParagraphText: visibleText,
      };
    }

    const validInlineStructure =
      visibleText.length > 0 &&
      paragraphModes.every((mode) => mode === "inline");
    if (!validInlineStructure) {
      throw new Error(
        `${stage} inline image ${index + 1} is not embedded in a body-text ` +
          `paragraph: ${JSON.stringify({ shape, paragraphShapes, visibleText })}`,
      );
    }
    return {
      ...shape,
      layoutStructure: "inline-text-flow",
      paragraphFormulaCount: paragraphShapes.length,
      visibleParagraphText: visibleText,
    };
  });
  return {
    stage,
    inlineShapeCount,
    macroButtonCount,
    shapes: structuredShapes,
  };
}

function saveAndReopenWordDocument(testDocumentName) {
  rmSync(reopenedDocumentPath, { force: true });
  try {
    runAppleScript([
      'tell application "Microsoft Word"',
      `set documentObject to document ${JSON.stringify(testDocumentName)}`,
      `save as documentObject file name ${JSON.stringify(reopenedDocumentPath)}`,
      "end tell",
    ], 90_000);
  } catch (error) {
    // Word for Mac can successfully complete SaveAs and then return -128 after
    // the old AppleScript document wrapper becomes invalid. The on-disk DOCX
    // is the source of truth; never retry SaveAs or keep using that wrapper.
    if (!existsSync(reopenedDocumentPath)) throw error;
  }
  runAppleScript(['tell application "Microsoft Word" to quit saving no'], 30_000);
  spawnSync("/bin/sleep", ["2"], { encoding: "utf8" });
  return runAppleScript([
    'tell application "Microsoft Word"',
    `open file name ${JSON.stringify(reopenedDocumentPath)}`,
    "set reopenedDocument to active document",
    "activate object reopenedDocument",
    "activate",
    "return name of reopenedDocument",
    "end tell",
  ], 90_000);
}

function validateFormulaEditSession(
  sessionId,
  formula,
  expectedCodeFormat,
  expectedLines,
) {
  const requestPath = join(sessionsRoot, sessionId, "request.json");
  const request = JSON.parse(readFileSync(requestPath, "utf8"));
  if (
    request.mode !== "edit" ||
    request.host !== "word" ||
    request.formulaId !== formula.formulaId ||
    request.displayMode !== formula.displayMode ||
    Boolean(request.numbered) !== formula.numbered
  ) {
    throw new Error(
      `Unexpected Word formula edit request: ${JSON.stringify(request)}`,
    );
  }
  const metadata = decodeFormulaMetadata(request.encodedMetadata ?? "");
  if (!metadata || metadata.formulaId !== formula.formulaId) {
    throw new Error(`Word edit Session lost formula metadata for ${formula.formulaId}`);
  }
  const normalized = normalizeFormulaEditorDocument(
    metadata.lines,
    metadata.codeFormat,
  );
  if (
    normalized.codeFormat !== expectedCodeFormat ||
    JSON.stringify(normalized.lines.map((line) => line.latex)) !==
      JSON.stringify(expectedLines)
  ) {
    throw new Error(
      `Word edit Session did not restore ${expectedCodeFormat}: ${JSON.stringify({
        metadata,
        normalized,
      })}`,
    );
  }
  if (normalized.lines[0]?.id !== formula.metadataLineId) {
    throw new Error("Word edit normalization did not preserve the imported formula line UUID");
  }
  if (
    Math.abs((request.fontSizePt ?? 0) - formula.fontSizePt) > 0.1
  ) {
    throw new Error(
      `Word edit Session lost formula font size: ${request.fontSizePt}`,
    );
  }
  return { request, metadata, normalized };
}

async function waitForNewSession(before, timeoutMs = 12_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    if (existsSync(sessionsRoot)) {
      const ready = [];
      for (const entry of readdirSync(sessionsRoot, { withFileTypes: true })) {
        if (!entry.isDirectory() || before.has(entry.name)) continue;
        const requestPath = join(sessionsRoot, entry.name, "request.json");
        if (!existsSync(requestPath)) continue;
        try {
          const request = JSON.parse(readFileSync(requestPath, "utf8"));
          if (
            request.operation === "documentImport" &&
            request.sessionId === entry.name &&
            request.host === "word"
          ) {
            ready.push({
              sessionId: entry.name,
              modifiedAt: statSync(requestPath).mtimeMs,
            });
          }
        } catch {
          // The Session directory can become visible before its atomic
          // request.json rename. Wait for a complete, validated request.
        }
      }
      if (ready.length > 0) {
        ready.sort((left, right) => right.modifiedAt - left.modifiedAt);
        return ready[0].sessionId;
      }
    }
    await sleep(100);
  }
  throw new Error("Word did not create a VisualTeX document import Session");
}

async function waitForWordCreateSession(before, timeoutMs = 12_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    for (const sessionId of currentSessionIds()) {
      if (before.has(sessionId)) continue;
      const requestPath = join(sessionsRoot, sessionId, "request.json");
      if (!existsSync(requestPath)) continue;
      try {
        const request = JSON.parse(readFileSync(requestPath, "utf8"));
        if (
          request.mode === "create" &&
          request.host === "word" &&
          request.sessionId === sessionId
        ) {
          return sessionId;
        }
      } catch {
        // The request may still be completing its atomic rename.
      }
    }
    await sleep(100);
  }
  throw new Error("Word did not create a VisualTeX formula creation Session");
}

async function waitForFormulaEditSession(before, formulaId, timeoutMs = 12_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    for (const sessionId of currentSessionIds()) {
      if (before.has(sessionId)) continue;
      const requestPath = join(sessionsRoot, sessionId, "request.json");
      if (!existsSync(requestPath)) continue;
      try {
        const request = JSON.parse(readFileSync(requestPath, "utf8"));
        if (request.mode === "edit" && request.formulaId === formulaId) {
          return sessionId;
        }
      } catch {
        // The request may still be in the middle of its atomic write.
      }
    }
    await sleep(100);
  }
  throw new Error(`Word did not create an edit Session for ${formulaId}`);
}

function editorPerformanceRecords(sessionId) {
  const performancePath = join(
    sessionsRoot,
    sessionId,
    editorPerformanceFileName,
  );
  if (!existsSync(performancePath)) return [];
  return readFileSync(performancePath, "utf8")
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line, index) => {
      try {
        return JSON.parse(line);
      } catch (error) {
        throw new Error(
          `Invalid editor performance record ${index + 1} for ${sessionId}: ${error}`,
        );
      }
    });
}

function validatedPhysicalEditorReadiness(sessionId, formulaId, marker, records) {
  if (
    marker.schema !== editorReadySchema ||
    marker.sessionId !== sessionId ||
    marker.host !== "word" ||
    !Number.isSafeInteger(marker.generation) ||
    marker.generation <= 0
  ) {
    throw new Error(
      `The physical edit wrote an invalid editor-ready marker: ${JSON.stringify(marker)}`,
    );
  }
  for (const key of [
    "epochMs",
    "urlReceivedEpochMs",
    "frontendEpochMs",
  ]) {
    if (!Number.isSafeInteger(marker[key]) || marker[key] <= 0) {
      throw new Error(`The editor-ready marker has an invalid ${key}`);
    }
  }
  for (const key of [
    "hydrateMs",
    "editorMountedMs",
    "contentReadyMs",
    "showFocusMs",
  ]) {
    if (!Number.isFinite(marker[key]) || marker[key] < 0) {
      throw new Error(`The editor-ready marker has an invalid ${key}`);
    }
  }
  if (
    marker.hydrateMs > marker.editorMountedMs ||
    marker.editorMountedMs > marker.contentReadyMs ||
    marker.contentReadyMs > marker.showFocusMs + 10
  ) {
    throw new Error(
      `The physical editor readiness stages are out of order: ${JSON.stringify(marker)}`,
    );
  }
  if (
    marker.urlReceivedEpochMs > marker.frontendEpochMs + 100 ||
    marker.frontendEpochMs > marker.epochMs + 100
  ) {
    throw new Error(
      `The physical editor readiness epochs are out of order: ${JSON.stringify(marker)}`,
    );
  }
  if (
    marker.contentReadyMs > warmEditorReadyLimitMs ||
    marker.showFocusMs > warmEditorReadyLimitMs
  ) {
    throw new Error(
      `The resident Word editor missed the ${warmEditorReadyLimitMs} ms warm target: ` +
        JSON.stringify({ marker, records }),
    );
  }

  const requiredStages = [
    "url-received",
    "request-read",
    "request-imported",
    "window-reused",
    "activation-event-sent",
    "frontend-hydrated",
    "frontend-editor-mounted",
    "frontend-content-ready",
    "window-show-focus",
  ];
  for (const record of records) {
    if (
      record.schema !== editorPerformanceSchema ||
      record.sessionId !== sessionId ||
      record.host !== "word" ||
      !Number.isFinite(record.elapsedMs) ||
      record.elapsedMs < 0
    ) {
      throw new Error(
        `The physical edit wrote an invalid performance record: ${JSON.stringify(record)}`,
      );
    }
  }
  const byStage = Object.fromEntries(
    requiredStages.map((stage) => [
      stage,
      records.filter((record) => record.stage === stage),
    ]),
  );
  for (const stage of requiredStages) {
    if (byStage[stage].length !== 1) {
      throw new Error(
        `The physical edit did not record exactly one ${stage} stage: ${JSON.stringify(records)}`,
      );
    }
  }
  if (records.some((record) => record.stage === "window-created")) {
    throw new Error(
      "The physical Word edit created a new WebView instead of reusing the resident editor",
    );
  }
  for (const stage of [
    "window-reused",
    "activation-event-sent",
    "frontend-hydrated",
    "frontend-editor-mounted",
    "frontend-content-ready",
    "window-show-focus",
  ]) {
    if (byStage[stage][0].generation !== marker.generation) {
      throw new Error(
        `The ${stage} performance record belongs to a stale editor generation`,
      );
    }
  }
  const stageElapsed = Object.fromEntries(
    requiredStages.map((stage) => [stage, byStage[stage][0].elapsedMs]),
  );
  const frontendOriginMs =
    stageElapsed["frontend-content-ready"] - marker.contentReadyMs;
  if (!Number.isFinite(frontendOriginMs) || frontendOriginMs < -1) {
    throw new Error(
      `The physical editor frontend timing origin is invalid: ${JSON.stringify({ frontendOriginMs, stageElapsed, marker })}`,
    );
  }
  for (const [stage, markerKey] of [
    ["frontend-hydrated", "hydrateMs"],
    ["frontend-editor-mounted", "editorMountedMs"],
    ["frontend-content-ready", "contentReadyMs"],
  ]) {
    const frontendRelativeMs = stageElapsed[stage] - frontendOriginMs;
    if (Math.abs(frontendRelativeMs - marker[markerKey]) > 1) {
      throw new Error(
        `The ${stage} timing disagrees with editor-ready.${markerKey}: ${JSON.stringify({ frontendRelativeMs, stageElapsed: stageElapsed[stage], frontendOriginMs, markerValue: marker[markerKey] })}`,
      );
    }
  }
  if (Math.abs(stageElapsed["window-show-focus"] - marker.showFocusMs) > 1) {
    throw new Error(
      "The window-show-focus timing disagrees with editor-ready.showFocusMs",
    );
  }
  const backendOrder = [
    "url-received",
    "request-read",
    "request-imported",
    "window-reused",
    "activation-event-sent",
    "window-show-focus",
  ];
  for (let index = 1; index < backendOrder.length; index += 1) {
    const previous = backendOrder[index - 1];
    const current = backendOrder[index];
    if (stageElapsed[current] + 1 < stageElapsed[previous]) {
      throw new Error(
        `The physical editor backend stages are out of order: ${JSON.stringify(stageElapsed)}`,
      );
    }
  }

  const requestPath = join(sessionsRoot, sessionId, "request.json");
  const requestWrittenEpochMs = statSync(requestPath).mtimeMs;
  const requestToUrlMs = marker.urlReceivedEpochMs - requestWrittenEpochMs;
  const requestToReadyMs = marker.epochMs - requestWrittenEpochMs;
  const urlToReadyEpochMs = marker.epochMs - marker.urlReceivedEpochMs;
  if (
    requestToUrlMs < -250 ||
    requestToUrlMs > 2_000 ||
    requestToReadyMs < -250 ||
    requestToReadyMs > 1_500 ||
    Math.abs(urlToReadyEpochMs - marker.showFocusMs) > 250
  ) {
    throw new Error(
      `The physical editor request/URL/readiness timing is invalid: ${JSON.stringify({
        requestWrittenEpochMs,
        requestToUrlMs,
        requestToReadyMs,
        urlToReadyEpochMs,
        marker,
      })}`,
    );
  }
  return {
    schema: editorReadySchema,
    sessionId,
    formulaId,
    generation: marker.generation,
    requestWrittenEpochMs,
    requestToUrlMs,
    requestToReadyMs,
    urlToReadyEpochMs,
    ...Object.fromEntries(
      [
        "urlReceivedEpochMs",
        "frontendEpochMs",
        "epochMs",
        "hydrateMs",
        "editorMountedMs",
        "contentReadyMs",
        "showFocusMs",
      ].map((key) => [key, marker[key]]),
    ),
    stages: stageElapsed,
  };
}

async function waitForPhysicalEditorVisible(
  sessionId,
  formulaId,
  timeoutMs = 30_000,
) {
  const readyPath = join(sessionsRoot, sessionId, editorReadyFileName);
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    if (existsSync(readyPath)) {
      const marker = JSON.parse(readFileSync(readyPath, "utf8"));
      if (
        marker.schema === editorReadySchema &&
        marker.sessionId === sessionId &&
        marker.host === "word" &&
        marker.windowVisible === true &&
        Number.isFinite(marker.contentReadyMs)
      ) {
        return {
          schema: marker.schema,
          sessionId,
          formulaId,
          generation: marker.generation,
          windowVisible: marker.windowVisible,
          windowFocused: marker.windowFocused,
          contentReadyMs: marker.contentReadyMs,
          showFocusMs: marker.showFocusMs,
        };
      }
    }
    await sleep(50);
  }
  throw new Error(`VisualTeX did not expose a visible Word editor for ${formulaId}`);
}

async function waitForPhysicalEditorReadiness(
  sessionId,
  formulaId,
  timeoutMs = 30_000,
) {
  const readyPath = join(sessionsRoot, sessionId, editorReadyFileName);
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    if (existsSync(readyPath)) {
      const marker = JSON.parse(readFileSync(readyPath, "utf8"));
      const records = editorPerformanceRecords(sessionId);
      const requiredStageNames = new Set(records.map((record) => record.stage));
      if (
        [
          "url-received",
          "request-read",
          "request-imported",
          "window-reused",
          "activation-event-sent",
          "frontend-hydrated",
          "frontend-editor-mounted",
          "frontend-content-ready",
          "window-show-focus",
        ].every((stage) => requiredStageNames.has(stage))
      ) {
        await sleep(100);
        return validatedPhysicalEditorReadiness(
          sessionId,
          formulaId,
          marker,
          editorPerformanceRecords(sessionId),
        );
      }
    }
    await sleep(50);
  }
  const sessionDirectory = join(sessionsRoot, sessionId);
  const files = existsSync(sessionDirectory)
    ? readdirSync(sessionDirectory).sort()
    : [];
  const requestPath = join(sessionDirectory, "request.json");
  const performancePath = join(sessionDirectory, editorPerformanceFileName);
  const request = existsSync(requestPath)
    ? readFileSync(requestPath, "utf8")
    : "<missing>";
  const performance = existsSync(performancePath)
    ? readFileSync(performancePath, "utf8")
    : "<missing>";
  throw new Error(
    `VisualTeX did not write ${editorReadyFileName} and complete performance stages for ${formulaId}: ${JSON.stringify({
      sessionId,
      files,
      request,
      performance,
    })}`,
  );
}

async function assertSinglePhysicalEditSession(before, sessionId, formulaId) {
  await sleep(300);
  const matchingSessionIds = [];
  for (const candidateSessionId of currentSessionIds()) {
    if (before.has(candidateSessionId)) continue;
    const requestPath = join(
      sessionsRoot,
      candidateSessionId,
      "request.json",
    );
    if (!existsSync(requestPath)) continue;
    try {
      const request = JSON.parse(readFileSync(requestPath, "utf8"));
      if (request.mode === "edit" && request.formulaId === formulaId) {
        matchingSessionIds.push(candidateSessionId);
      }
    } catch {
      // A different Session may still be finishing an atomic write.
    }
  }
  if (
    matchingSessionIds.length !== 1 ||
    matchingSessionIds[0] !== sessionId
  ) {
    throw new Error(
      `One physical double-click must create exactly one edit Session: ${JSON.stringify({
        sessionId,
        formulaId,
        matchingSessionIds,
      })}`,
    );
  }
}

function formulaItem({
  formulaId,
  latex,
  metadataLatex = latex,
  expectedCodeFormat,
  displayMode,
  numbered,
  fontSizePt,
  artifactDirectory,
  formulaLetterFont = "katex",
  formulaChineseFont = "system",
}) {
  const normalized = normalizeFormulaEditorDocument(
    [{ id: crypto.randomUUID(), latex: metadataLatex }],
    "raw",
  );
  if (expectedCodeFormat && normalized.codeFormat !== expectedCodeFormat) {
    throw new Error(
      `Formula fixture did not normalize as ${expectedCodeFormat}: ${JSON.stringify(normalized)}`,
    );
  }
  const canonicalLatex = serializeFormulaEditorDocument(normalized);
  const normalizedLines = normalized.lines.map((line) => line.latex);
  const omml = ommlForFormula(normalizedLines, normalized.codeFormat);
  if (["align", "align-star"].includes(normalized.codeFormat)) {
    if (
      (omml.match(/<m:oMath\b/g) ?? []).length !== 1 ||
      (omml.match(/<m:eqArr>/g) ?? []).length !== 1 ||
      (omml.match(/&amp;/g) ?? []).length !== normalizedLines.length
    ) {
      throw new Error(
        `Aligned fixture is not one relationship-aligned OMML equation array: ${omml}`,
      );
    }
  }
  const nativePath = join(nativeRoot, `${formulaId}.docx`);
  writeFileSync(nativePath, minimalDocxBytes(omml));

  let imagePath = "";
  let vectorDocumentPath = "";
  let fallbackImagePath = "";
  let widthPoints = fontSizePt;
  let heightPoints = Math.max(18, fontSizePt * 1.8);
  let baseline = 0;
  let referenceWidthPt = referenceFontSizePt;
  let referenceHeightPt = referenceFontSizePt;
  let referenceBaselinePt = 0;
  let renderWidthPx;
  let renderHeightPx;
  let svgAlignment;

  if (outputKind === "image") {
    const svg = latexToSvg(canonicalLatex, {
      displayMode: displayMode === "block",
      fontSizePt: referenceFontSizePt,
      paddingPx: displayMode === "inline" ? 1 : wordDisplayPaddingPx,
      background: "transparent",
      forceExplicitBlack: true,
      formulaLetterFont,
      formulaChineseFont,
    });
    const geometry = calculateImageGeometry(svg, fontSizePt, formulaLetterFont);
    ({
      widthPoints,
      heightPoints,
      baseline,
      referenceWidthPt,
      referenceHeightPt,
      referenceBaselinePt,
    } = geometry);
    renderWidthPx = svg.width;
    renderHeightPx = svg.height;
    if (["align", "align-star"].includes(normalized.codeFormat)) {
      svgAlignment = assertAlignedSvg(
        svg.svg,
        normalized.lines.length,
        `Initial ${normalized.codeFormat} image`,
      );
    }
    const stem = `document-formula-${compactFormulaId(formulaId)}`;
    imagePath = join(artifactDirectory, `${stem}.svg`);
    fallbackImagePath = join(artifactDirectory, `${stem}.png`);
    vectorDocumentPath = join(artifactDirectory, `${stem}-svg.docx`);
    writeFileSync(imagePath, svg.svg, { mode: 0o600 });
    writeFileSync(fallbackImagePath, transparentPng, { mode: 0o600 });
    writeFileSync(
      vectorDocumentPath,
      wordSvgDocxBytes(svg.svg, transparentPng, widthPoints, heightPoints),
      { mode: 0o600 },
    );
  }

  const metadata = createFormulaMetadata({
    formulaId,
    title: displayMode === "inline" ? "Integration inline formula" : "Integration display formula",
    lines: normalized.lines,
    codeFormat: normalized.codeFormat,
    sourceLatex: canonicalLatex,
    displayMode,
    numbered,
    fontSizePt,
    formulaLetterFont,
    formulaChineseFont,
    referenceWidthPt,
    referenceHeightPt,
    referenceBaselinePt,
    renderWidthPx,
    renderHeightPx,
  });
  return {
    formulaId,
    latex: canonicalLatex,
    metadataLatex: canonicalLatex,
    metadataLineId: normalized.lines[0].id,
    metadataLines: normalizedLines,
    codeFormat: normalized.codeFormat,
    pdfToken: normalizedLines.join(""),
    displayMode,
    numbered,
    fontSizePt,
    formulaLetterFont,
    formulaChineseFont,
    metadata: encodeFormulaMetadata(metadata),
    ommlBase64: Buffer.from(omml, "utf8").toString("base64url"),
    nativePath,
    imagePath,
    vectorDocumentPath,
    fallbackImagePath,
    widthPoints,
    heightPoints,
    baseline,
    referenceWidthPt,
    referenceHeightPt,
    referenceBaselinePt,
    svgAlignment,
  };
}

function editedImageFormulaArtifacts(
  formula,
  editSession,
  updatedLineLatex,
  artifactDirectory,
) {
  const updatedLines = editSession.normalized.lines.map((line, index) => ({
    ...line,
    latex: updatedLineLatex[index] ?? line.latex,
  }));
  if (updatedLines.length !== updatedLineLatex.length) {
    throw new Error(
      `Edited ${formula.codeFormat} fixture changed its row count unexpectedly`,
    );
  }
  const rendered = renderOfficeFormulaArtifacts({
    lines: updatedLines,
    codeFormat: editSession.normalized.codeFormat,
    displayMode: formula.displayMode,
    host: "word",
    includeWordOmml: false,
  });
  const svgAlignment = assertAlignedSvg(
    rendered.svg.svg,
    rendered.lines.length,
    `Edited ${rendered.codeFormat} image`,
  );
  const geometry = calculateImageGeometry(
    rendered.svg,
    formula.fontSizePt,
    formula.formulaLetterFont ?? "katex",
  );
  const stem = `edited-${compactFormulaId(formula.formulaId)}`;
  const imagePath = join(artifactDirectory, `${stem}.svg`);
  const fallbackImagePath = join(artifactDirectory, `${stem}.png`);
  const vectorDocumentPath = join(artifactDirectory, `${stem}-svg.docx`);
  writeFileSync(imagePath, rendered.svg.svg, { mode: 0o600 });
  writeFileSync(fallbackImagePath, transparentPng, { mode: 0o600 });
  writeFileSync(
    vectorDocumentPath,
    wordSvgDocxBytes(
      rendered.svg.svg,
      transparentPng,
      geometry.widthPoints,
      geometry.heightPoints,
    ),
    { mode: 0o600 },
  );

  const omml = ommlForFormula(
    rendered.lines.map((line) => line.latex),
    rendered.codeFormat,
  );
  writeFileSync(formula.nativePath, minimalDocxBytes(omml), { mode: 0o600 });
  const metadata = createFormulaMetadata({
    formulaId: formula.formulaId,
    title: editSession.metadata.title,
    lines: rendered.lines,
    codeFormat: rendered.codeFormat,
    sourceLatex: rendered.canonicalLatex,
    displayMode: formula.displayMode,
    numbered: formula.numbered,
    fontSizePt: formula.fontSizePt,
    referenceWidthPt: geometry.referenceWidthPt,
    referenceHeightPt: geometry.referenceHeightPt,
    referenceBaselinePt: geometry.referenceBaselinePt,
    renderWidthPx: rendered.svg.width,
    renderHeightPx: rendered.svg.height,
    original: editSession.metadata,
  });
  return {
    lines: rendered.lines,
    codeFormat: rendered.codeFormat,
    canonicalLatex: rendered.canonicalLatex,
    metadata: encodeFormulaMetadata(metadata),
    ommlBase64: Buffer.from(omml, "utf8").toString("base64url"),
    imagePath,
    fallbackImagePath,
    vectorDocumentPath,
    ...geometry,
    renderWidthPx: rendered.svg.width,
    renderHeightPx: rendered.svg.height,
    svgAlignment,
  };
}

function commitEditedImageFormula(
  testDocumentName,
  sessionId,
  formula,
  editSession,
  updatedLineLatex,
) {
  const sessionDirectory = join(sessionsRoot, sessionId);
  const artifacts = editedImageFormulaArtifacts(
    formula,
    editSession,
    updatedLineLatex,
    sessionDirectory,
  );
  const request = editSession.request;
  const dispatch = manifestText([
    ["protocolVersion", "1"],
    ["sessionId", sessionId],
    ["action", "commit"],
    ["host", "word"],
    ["mode", "edit"],
    ["formulaId", formula.formulaId],
    ["displayMode", formula.displayMode],
    ["numbered", formula.numbered ? "1" : "0"],
    ["nativeEquation", "0"],
    ["imagePath", artifacts.imagePath],
    ["vectorDocumentPath", artifacts.vectorDocumentPath],
    ["fallbackImagePath", artifacts.fallbackImagePath],
    ["metadata", artifacts.metadata],
    ["latexBase64", base64Url(artifacts.canonicalLatex)],
    ["ommlBase64", artifacts.ommlBase64],
    ["nativeDocumentPath", formula.nativePath],
    ["pendingMarker", request.pendingMarker ?? ""],
    [
      "sourceMarker",
      request.sourceObjectId ?? request.encodedMetadata ?? "",
    ],
    ["sourceDocumentId", request.sourceDocumentId ?? ""],
    ["widthPoints", artifacts.widthPoints.toFixed(6)],
    ["heightPoints", artifacts.heightPoints.toFixed(6)],
    ["baseline", artifacts.baseline.toFixed(6)],
    ["fontSizePt", formula.fontSizePt.toFixed(6)],
    ["referenceWidthPt", artifacts.referenceWidthPt.toFixed(6)],
    ["referenceHeightPt", artifacts.referenceHeightPt.toFixed(6)],
    ["referenceBaselinePt", artifacts.referenceBaselinePt.toFixed(6)],
  ]);
  writeFileSync(join(sessionDirectory, "dispatch.txt"), dispatch, { mode: 0o600 });
  writeFileSync(join(sessionsRoot, "word-active-session.txt"), sessionId, {
    mode: 0o600,
  });
  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(testDocumentName)}`,
    'run VB macro macro name "VisualTeX_ApplyPendingResult"',
    "end tell",
  ], 90_000);

  formula.latex = artifacts.canonicalLatex;
  formula.metadataLatex = artifacts.canonicalLatex;
  formula.metadataLines = artifacts.lines.map((line) => line.latex);
  formula.codeFormat = artifacts.codeFormat;
  formula.metadata = artifacts.metadata;
  formula.ommlBase64 = artifacts.ommlBase64;
  formula.imagePath = artifacts.imagePath;
  formula.vectorDocumentPath = artifacts.vectorDocumentPath;
  formula.fallbackImagePath = artifacts.fallbackImagePath;
  formula.widthPoints = artifacts.widthPoints;
  formula.heightPoints = artifacts.heightPoints;
  formula.baseline = artifacts.baseline;
  formula.referenceWidthPt = artifacts.referenceWidthPt;
  formula.referenceHeightPt = artifacts.referenceHeightPt;
  formula.referenceBaselinePt = artifacts.referenceBaselinePt;
  formula.svgAlignment = artifacts.svgAlignment;
  return artifacts;
}

function appendParagraphMetadata(entries, index, paragraph) {
  if (!paragraph) return;
  const prefix = `item${index}`;
  entries.push([`${prefix}paragraphId`, paragraph.id]);
  entries.push([`${prefix}paragraphStyle`, paragraph.style ?? "normal"]);
  entries.push([`${prefix}paragraphAlignment`, paragraph.alignment ?? "left"]);
  entries.push([`${prefix}listKind`, paragraph.listKind ?? "none"]);
  entries.push([`${prefix}listLevel`, String(paragraph.listLevel ?? 0)]);
  entries.push([`${prefix}paragraphStart`, paragraph.start ? "1" : "0"]);
  entries.push([`${prefix}paragraphEnd`, paragraph.end ? "1" : "0"]);
}

function appendText(entries, index, text, paragraph) {
  entries.push([`item${index}kind`, "text"]);
  entries.push([`item${index}textBase64`, base64Url(text)]);
  appendParagraphMetadata(entries, index, paragraph);
}

function appendFormula(entries, index, formula, paragraph) {
  const prefix = `item${index}`;
  entries.push([`${prefix}kind`, "formula"]);
  entries.push([`${prefix}formulaId`, formula.formulaId]);
  entries.push([`${prefix}latexBase64`, base64Url(formula.latex)]);
  entries.push([`${prefix}displayMode`, formula.displayMode]);
  entries.push([`${prefix}numbered`, formula.numbered ? "1" : "0"]);
  entries.push([`${prefix}fontSizePt`, formula.fontSizePt.toFixed(6)]);
  entries.push([`${prefix}metadata`, formula.metadata]);
  entries.push([`${prefix}ommlBase64`, formula.ommlBase64]);
  entries.push([`${prefix}nativeDocumentPath`, formula.nativePath]);
  entries.push([`${prefix}imagePath`, formula.imagePath]);
  entries.push([`${prefix}vectorDocumentPath`, formula.vectorDocumentPath]);
  entries.push([`${prefix}fallbackImagePath`, formula.fallbackImagePath]);
  entries.push([`${prefix}widthPoints`, formula.widthPoints.toFixed(6)]);
  entries.push([`${prefix}heightPoints`, formula.heightPoints.toFixed(6)]);
  entries.push([`${prefix}baseline`, formula.baseline.toFixed(6)]);
  entries.push([`${prefix}referenceWidthPt`, formula.referenceWidthPt.toFixed(6)]);
  entries.push([`${prefix}referenceHeightPt`, formula.referenceHeightPt.toFixed(6)]);
  entries.push([`${prefix}referenceBaselinePt`, formula.referenceBaselinePt.toFixed(6)]);
  appendParagraphMetadata(entries, index, paragraph);
}

function exportWordPdfWithoutSelectingFormula(
  testDocumentName,
  outputPath,
  label,
) {
  rmSync(outputPath, { force: true });
  rmSync(pdfExportStatusPath, { force: true });
  writeFileSync(pdfExportRequestPath, outputPath, { mode: 0o600 });
  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(testDocumentName)}`,
    'run VB macro macro name "VisualTeX_ExportActiveDocumentPdfForRegression"',
    "end tell",
  ], 90_000);
  const exportStatus = existsSync(pdfExportStatusPath)
    ? readFileSync(pdfExportStatusPath, "utf8").trim()
    : "missing-status";
  if (!exportStatus.startsWith("ok|") || !existsSync(outputPath)) {
    throw new Error(`${label} PDF export failed: ${exportStatus}`);
  }
}

function pdfRasterSummary(pdfPath, label, minimumBands = 1) {
  const result = spawnSync(
    "/usr/bin/swift",
    [
      join(repositoryRoot, "scripts/pdf_formula_geometry.swift"),
      pdfPath,
      "--raster-only",
    ],
    {
      encoding: "utf8",
      timeout: 120_000,
      maxBuffer: 8 * 1024 * 1024,
    },
  );
  if (result.status !== 0) {
    throw new Error(result.stderr.trim() || `${label} PDF raster extraction failed`);
  }
  const geometry = JSON.parse(result.stdout);
  const bands = geometry.rasterBands ?? [];
  const components = bands.flatMap((band) => band.components ?? []);
  if (bands.length < minimumBands || components.length === 0) {
    throw new Error(
      `${label} PDF contains insufficient visible formula ink: ${JSON.stringify({ bands: bands.length, components: components.length })}`,
    );
  }
  const minX = Math.min(...components.map((component) => component.minX));
  const maxX = Math.max(...components.map((component) => component.maxX));
  const minY = Math.min(...bands.map((band) => band.minY));
  const maxY = Math.max(...bands.map((band) => band.maxY));
  return {
    pageWidth: geometry.pageWidth,
    pageHeight: geometry.pageHeight,
    bandCount: bands.length,
    componentCount: components.length,
    minX,
    maxX,
    minY,
    maxY,
    width: maxX - minX,
    height: maxY - minY,
  };
}

function assertEquivalentPdfInk(reference, candidate, label) {
  const differences = {
    bandCount: Math.abs(reference.bandCount - candidate.bandCount),
    componentCount: Math.abs(reference.componentCount - candidate.componentCount),
    minX: Math.abs(reference.minX - candidate.minX),
    maxX: Math.abs(reference.maxX - candidate.maxX),
    minY: Math.abs(reference.minY - candidate.minY),
    maxY: Math.abs(reference.maxY - candidate.maxY),
  };
  if (
    differences.bandCount !== 0 ||
    differences.componentCount !== 0 ||
    differences.minX > 1 ||
    differences.maxX > 1 ||
    differences.minY > 1 ||
    differences.maxY > 1
  ) {
    throw new Error(
      `${label} changed visible PDF ink: ${JSON.stringify({ reference, candidate, differences })}`,
    );
  }
}

function saveUntouchedWordDocument(testDocumentName, outputPath) {
  rmSync(outputPath, { force: true });
  try {
    return runAppleScript([
      'tell application "Microsoft Word"',
      `set documentObject to document ${JSON.stringify(testDocumentName)}`,
      `save as documentObject file name ${JSON.stringify(outputPath)}`,
      "return name of active document",
      "end tell",
    ], 90_000);
  } catch (error) {
    if (!existsSync(outputPath)) throw error;
    return runAppleScript([
      'tell application "Microsoft Word"',
      "return name of active document",
      "end tell",
    ], 30_000);
  }
}

function scrollAwayAndBackWithoutSelectingFormula(testDocumentName) {
  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(testDocumentName)}`,
    "activate",
    "end tell",
    'tell application "System Events"',
    'tell process "Microsoft Word"',
    "key code 125 using {command down}",
    "delay 0.7",
    "key code 126 using {command down}",
    "delay 0.7",
    "end tell",
    "end tell",
  ], 30_000);
}

function coldReopenWordDocument(documentPath) {
  runAppleScript(['tell application "Microsoft Word" to quit saving no'], 30_000);
  spawnSync("/usr/bin/killall", ["Microsoft Word"], {
    encoding: "utf8",
    timeout: 10_000,
  });
  spawnSync("/bin/sleep", ["2"], { encoding: "utf8" });
  return runAppleScript([
    'tell application "Microsoft Word"',
    `open file name ${JSON.stringify(documentPath)}`,
    "set reopenedDocument to active document",
    "activate object reopenedDocument",
    "activate",
    "return name of reopenedDocument",
    "end tell",
  ], 90_000);
}

function createdImagePdfInkBounds(testDocumentName, label) {
  rmSync(coordinatePdfPath, { force: true });
  rmSync(pdfExportStatusPath, { force: true });
  writeFileSync(pdfExportRequestPath, coordinatePdfPath, { mode: 0o600 });
  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(testDocumentName)}`,
    'run VB macro macro name "VisualTeX_ExportActiveDocumentPdfForRegression"',
    "end tell",
  ], 90_000);
  const exportStatus = existsSync(pdfExportStatusPath)
    ? readFileSync(pdfExportStatusPath, "utf8").trim()
    : "missing-status";
  if (!exportStatus.startsWith("ok|") || !existsSync(coordinatePdfPath)) {
    throw new Error(`${label} PDF export failed: ${exportStatus}`);
  }
  const swiftGeometry = spawnSync(
    "/usr/bin/swift",
    [
      join(repositoryRoot, "scripts/pdf_formula_geometry.swift"),
      coordinatePdfPath,
      "--raster-only",
    ],
    {
      encoding: "utf8",
      timeout: 120_000,
      maxBuffer: 8 * 1024 * 1024,
    },
  );
  if (swiftGeometry.status !== 0) {
    throw new Error(
      swiftGeometry.stderr.trim() || `${label} PDF raster extraction failed`,
    );
  }
  const geometry = JSON.parse(swiftGeometry.stdout);
  const components = (geometry.rasterBands ?? []).flatMap(
    (band) => band.components ?? [],
  );
  if (components.length === 0) {
    throw new Error(`${label} PDF contains no visible formula ink`);
  }
  const minX = Math.min(...components.map((component) => component.minX));
  const maxX = Math.max(...components.map((component) => component.maxX));
  return {
    minX,
    maxX,
    width: maxX - minX,
    componentCount: components.length,
    rasterBands: geometry.rasterBands,
  };
}

function assertCreatedImageFormulaInk(bounds, formula, label) {
  const minimumVisibleWidth = Math.max(18, formula.widthPoints * 0.5);
  if (!Number.isFinite(bounds.width) || bounds.width < minimumVisibleWidth) {
    throw new Error(
      `${label} rendered as a fallback glyph instead of ${formula.latex}: ` +
        JSON.stringify({ bounds, expectedWidthPoints: formula.widthPoints }),
    );
  }
}

function writePictureRoutingFixture() {
  if (!existsSync(pictureRoutingBrowserArtifactsPath)) {
    throw new Error(
      `The browser-rendered picture fixture bundle is missing: ${pictureRoutingBrowserArtifactsPath}`,
    );
  }
  const fixtureBundle = JSON.parse(
    readFileSync(pictureRoutingBrowserArtifactsPath, "utf8"),
  );
  const fixturePngBase64 = fixtureBundle.formulas?.[0]?.pngBase64;
  if (!fixturePngBase64) {
    throw new Error("The browser-rendered picture fixture has no PNG payload");
  }
  writeFileSync(pictureRoutingFixturePng, Buffer.from(fixturePngBase64, "base64"), {
    mode: 0o600,
  });
}

async function runPictureRoutingPerformanceRegression() {
  writePictureRoutingFixture();
  const macroName = {
    1: "VisualTeX_RunPictureRoutingPerformance1",
    100: "VisualTeX_RunPictureRoutingPerformance100",
    1000: "VisualTeX_RunPictureRoutingPerformance1000",
  }[pictureRoutingPerformance];
  rmSync(pictureRoutingPerformancePath, { force: true });
  runAppleScript([
    'tell application "Microsoft Word"',
    "activate",
    `run VB macro macro name ${JSON.stringify(macroName)}`,
    "end tell",
  ], 300_000);
  if (!existsSync(pictureRoutingPerformancePath)) {
    throw new Error("Word did not write the picture-routing performance report");
  }
  const raw = readFileSync(pictureRoutingPerformancePath, "utf8").trim();
  const fields = raw.split("|");
  if (fields[0] !== "PASS" || fields.length !== 6) {
    throw new Error(`Picture-routing performance regression failed: ${raw}`);
  }
  const report = {
    requestedFormulaCount: Number(fields[1]),
    actualFormulaCount: Number(fields[2]),
    iterations: Number(fields[3]),
    visualTeXPerCallMs: Number(fields[4]),
    ordinaryPerCallMs: Number(fields[5]),
  };
  if (
    report.requestedFormulaCount !== pictureRoutingPerformance ||
    report.actualFormulaCount !== pictureRoutingPerformance ||
    report.iterations !== 5000 ||
    !Number.isFinite(report.visualTeXPerCallMs) ||
    !Number.isFinite(report.ordinaryPerCallMs) ||
    report.visualTeXPerCallMs > 10 ||
    report.ordinaryPerCallMs > 10 ||
    report.visualTeXPerCallMs > report.ordinaryPerCallMs + 1
  ) {
    throw new Error(
      `Picture-routing O(1) benchmark missed its target: ${JSON.stringify(report)}`,
    );
  }
  console.log(
    JSON.stringify(
      {
        kind: "word-picture-routing-performance",
        ...report,
        visualTeXTargetMs: 10,
        ordinaryOverheadToleranceMs: 1,
      },
      null,
      2,
    ),
  );
  console.log("Word picture routing performance regression passed");
}

async function runPictureRoutingNativeRegression(beforeSessions) {
  if (!existsSync(pictureRoutingBrowserArtifactsPath)) {
    throw new Error(
      `The browser-rendered picture fixture bundle is missing: ${pictureRoutingBrowserArtifactsPath}`,
    );
  }
  writePictureRoutingFixture();
  const macroName = {
    "ordinary-inline": "VisualTeX_CreateOrdinaryInlinePictureRegression",
    "ordinary-floating": "VisualTeX_CreateOrdinaryFloatingPictureRegression",
    "forged-prefix": "VisualTeX_CreateForgedPrefixPictureRegression",
    "damaged-metadata": "VisualTeX_CreateDamagedMetadataPictureRegression",
  }[pictureRoutingTarget];
  const testDocumentName = runAppleScript([
    'tell application "Microsoft Word"',
    "make new document",
    "set testDocument to active document",
    "activate object testDocument",
    "activate",
    `run VB macro macro name ${JSON.stringify(macroName)}`,
    "return name of testDocument",
    "end tell",
  ], 60_000);
  const sessionsBeforeClick = currentSessionIds();
  let physicalClick = {
    skipped: "damaged picture routing is validated through FormatPicture",
  };
  let physicalDoubleClickUi = {
    raw: "",
    visited: 0,
    paneMatches: [],
    contextualRibbonMatches: [],
    pictureFormatVisible: false,
  };
  let paneAfterCloseUi = null;
  if (pictureRoutingTarget === "ordinary-inline") {
    const physicalClickAttempts = [];
    for (let attempt = 1; attempt <= 3; attempt += 1) {
      physicalClickAttempts.push(
        physicallyDoubleClickSelectedWordFormula(
          testDocumentName,
          "VisualTeX_WriteSelectedPictureScreenBoundsRegression",
        ),
      );
      await sleep(1_750);
      physicalDoubleClickUi = wordPictureFormatUiSnapshot();
      if (physicalDoubleClickUi.pictureFormatVisible) break;
    }
    physicalClick = { attempts: physicalClickAttempts };
    if (!physicalDoubleClickUi.pictureFormatVisible) {
      throw new Error(
        `The pane detector missed Word's native Format Shape task pane after three ordinary-picture double-clicks: ${physicalDoubleClickUi.raw}`,
      );
    }
    const sessionsAfterPhysicalClick = [...currentSessionIds()].filter(
      (sessionId) => !sessionsBeforeClick.has(sessionId),
    );
    if (sessionsAfterPhysicalClick.length > 0) {
      throw new Error(
        `An ordinary picture double-click incorrectly created VisualTeX Session(s): ${sessionsAfterPhysicalClick.join(",")}`,
      );
    }
    const closePaneResult = closeWordPictureFormatTaskPaneViaAccessibility();
    if (!closePaneResult.clicked) {
      throw new Error(
        `The accessibility close helper did not click Word's Format Shape close button: ${JSON.stringify(closePaneResult)}`,
      );
    }
    await sleep(750);
    paneAfterCloseUi = wordPictureFormatUiSnapshot();
    if (paneAfterCloseUi.pictureFormatVisible) {
      throw new Error(
        `VisualTeX ClosePane fallback did not close Word's Format Shape task pane: ${paneAfterCloseUi.raw}`,
      );
    }
  }
  const nativeCommandProcess = await invokeWordFormatPictureCommand(testDocumentName);
  const nativeCommandUi = wordPictureFormatUiSnapshot();
  const nativeCommandPending =
    nativeCommandProcess.exitCode === null && nativeCommandProcess.signalCode === null;
  const sessionsAfterNativeCommand = [...currentSessionIds()].filter(
    (sessionId) => !sessionsBeforeClick.has(sessionId),
  );
  if (sessionsAfterNativeCommand.length > 0) {
    nativeCommandProcess.kill("SIGTERM");
    throw new Error(
      `${pictureRoutingTarget} FormatPicture incorrectly created VisualTeX Session(s): ${sessionsAfterNativeCommand.join(",")}`,
    );
  }
  // Word for Mac may return from WordBasic.FormatPicture immediately and its
  // non-modal Picture Format surface is not consistently exposed through the
  // accessibility tree. The durable contract here is that an unmarked picture
  // creates no VisualTeX Session and the VBA override falls through to the
  // unchanged native WordBasic command.
  runAppleScript([
    'tell application "System Events"',
    'tell process "Microsoft Word" to key code 53',
    'end tell',
  ], 15_000);
  nativeCommandProcess.kill("SIGTERM");
  console.log(
    JSON.stringify(
      {
        kind: "word-native-picture-routing",
        target: pictureRoutingTarget,
        testDocumentName,
        physicalClick,
        physicalDoubleClickUi,
        paneAfterCloseUi,
        nativeCommandUi,
        nativeCommandPending,
        nativeCommandDispatchedWithoutVisualTeXSession: true,
        visualTeXSessionsCreated: sessionsAfterNativeCommand,
        sessionsBeforeHarness: beforeSessions.size,
      },
      null,
      2,
    ),
  );
  console.log("Word native picture routing regression passed");
}

async function runFirstFrameImageRegression(beforeSessions) {
  const browserBundle = JSON.parse(readFileSync(firstFrameArtifactPath, "utf8"));
  if (
    browserBundle.schema !== "visualtex-word-browser-artifacts-v1" ||
    browserBundle.outputKind !== "image" ||
    !Array.isArray(browserBundle.formulas) ||
    browserBundle.formulas.length !== 4
  ) {
    throw new Error(
      `Unexpected browser artifact bundle: ${JSON.stringify({ schema: browserBundle.schema, outputKind: browserBundle.outputKind, formulaCount: browserBundle.formulas?.length })}`,
    );
  }

  runAppleScript([
    'tell application "Microsoft Word"',
    "activate",
    "end tell",
  ], 30_000);
  await sleep(3_000);
  let testDocumentName = runAppleScript([
    'tell application "Microsoft Word"',
    "make new document",
    "set testDocument to active document",
    "activate object testDocument",
    "activate",
    'run VB macro macro name "VisualTeX_InsertLatexMarkdownDocument"',
    "return name of testDocument",
    "end tell",
  ], 60_000);
  const sessionId = await waitForNewSession(beforeSessions);
  sessionDirectory = join(sessionsRoot, sessionId);
  const request = JSON.parse(
    readFileSync(join(sessionDirectory, "request.json"), "utf8"),
  );
  if (
    request.operation !== "documentImport" ||
    request.host !== "word" ||
    request.sessionId !== sessionId ||
    !request.sourceDocumentId ||
    !request.documentImport?.bookmarkName
  ) {
    throw new Error(
      `Unexpected first-frame Word document import request: ${JSON.stringify(request)}`,
    );
  }
  await stopVisualTeXForManualWordCallback();

  const formulas = browserBundle.formulas.map((formula, index) =>
    browserFormulaArtifact(formula, sessionDirectory, index),
  );
  nativeFiles.push(...formulas.map((formula) => formula.nativePath));
  const inlineParagraphId = crypto.randomUUID();
  const endingParagraphId = crypto.randomUUID();
  const items = [];
  appendText(items, 0, "首帧行内公式：", {
    id: inlineParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "none",
    listLevel: 0,
    start: true,
    end: false,
  });
  appendFormula(items, 1, formulas[0], {
    id: inlineParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "none",
    listLevel: 0,
    start: false,
    end: false,
  });
  appendText(items, 2, "，正文继续。", {
    id: inlineParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "none",
    listLevel: 0,
    start: false,
    end: true,
  });
  for (let index = 1; index < formulas.length; index += 1) {
    // Display formulas already create and finalize their own Word paragraph.
    // Supplying text-paragraph boundary metadata here double-finalizes it.
    appendFormula(items, index + 2, formulas[index]);
  }
  appendText(items, 6, "首帧验证结束。", {
    id: endingParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "none",
    listLevel: 0,
    start: true,
    end: true,
  });

  const manifestPath = join(sessionDirectory, "document-import.txt");
  writeFileSync(
    manifestPath,
    manifestText([
      ["protocolVersion", "1"],
      ["sessionId", sessionId],
      ["outputKind", "image"],
      ["sourceDocumentId", request.sourceDocumentId],
      ["bookmarkName", request.documentImport.bookmarkName],
      ["itemCount", "7"],
      ...items,
    ]),
    { mode: 0o600 },
  );
  writeFileSync(
    join(sessionDirectory, "dispatch.txt"),
    manifestText([
      ["protocolVersion", "1"],
      ["sessionId", sessionId],
      ["action", "documentCommit"],
      ["host", "word"],
      ["sourceDocumentId", request.sourceDocumentId],
      ["bookmarkName", request.documentImport.bookmarkName],
      ["documentImportPath", manifestPath],
    ]),
    { mode: 0o600 },
  );
  writeFileSync(join(sessionsRoot, "word-active-session.txt"), sessionId, {
    mode: 0o600,
  });

  const callbackStatusPath = join(sessionDirectory, "word-callback-status.txt");
  rmSync(callbackStatusPath, { force: true });
  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(testDocumentName)}`,
    'run VB macro macro name "VisualTeX_ApplyPendingResultForRegression"',
    "end tell",
  ], 90_000);
  const callbackStatus = existsSync(callbackStatusPath)
    ? readFileSync(callbackStatusPath, "utf8")
    : "missing-status";
  if (!callbackStatus.startsWith("PASS")) {
    throw new Error(`First-frame Word callback failed:\n${callbackStatus}`);
  }

  // Critical ordering: do not inspect, select or click an InlineShape before
  // the first PDF and DOCX are produced. These are the untouched Word frame.
  exportWordPdfWithoutSelectingFormula(
    testDocumentName,
    firstFramePdfPath,
    "Untouched first-frame",
  );
  const untouchedPdf = pdfRasterSummary(
    firstFramePdfPath,
    "Untouched first-frame",
    formulas.length,
  );
  testDocumentName = saveUntouchedWordDocument(
    testDocumentName,
    firstFrameDocumentPath,
  );
  const untouchedPackage = inspectSavedWordImagePackage(
    firstFrameDocumentPath,
    formulas,
    "untouched-first-frame",
  );

  scrollAwayAndBackWithoutSelectingFormula(testDocumentName);
  exportWordPdfWithoutSelectingFormula(
    testDocumentName,
    firstFrameScrolledPdfPath,
    "Scrolled-away-and-back first-frame",
  );
  const scrolledPdf = pdfRasterSummary(
    firstFrameScrolledPdfPath,
    "Scrolled-away-and-back first-frame",
    formulas.length,
  );
  assertEquivalentPdfInk(
    untouchedPdf,
    scrolledPdf,
    "Scrolling away and back",
  );

  testDocumentName = coldReopenWordDocument(firstFrameDocumentPath);
  exportWordPdfWithoutSelectingFormula(
    testDocumentName,
    firstFrameReopenedPdfPath,
    "Cold-reopened first-frame",
  );
  const reopenedPdf = pdfRasterSummary(
    firstFrameReopenedPdfPath,
    "Cold-reopened first-frame",
    formulas.length,
  );
  assertEquivalentPdfInk(untouchedPdf, reopenedPdf, "Cold Word restart");
  const reopenedPackage = inspectSavedWordImagePackage(
    firstFrameDocumentPath,
    formulas,
    "cold-reopened",
  );

  // Only after every untouched-frame acceptance has passed may the regression
  // read formula metadata and shape geometry through VBA.
  const formulaReport = runFormulaRegressionReport(testDocumentName, formulas);
  console.log(
    JSON.stringify(
      {
        kind: "word-image-untouched-first-frame",
        sessionId,
        documentPath: firstFrameDocumentPath,
        pdfPaths: {
          untouched: firstFramePdfPath,
          scrolled: firstFrameScrolledPdfPath,
          coldReopened: firstFrameReopenedPdfPath,
        },
        formulas: formulas.map((formula) => ({
          formulaId: formula.formulaId,
          displayMode: formula.displayMode,
          numbered: formula.numbered,
          codeFormat: formula.codeFormat,
          frontendPngHash: formula.frontendPngHash,
          frontendPngPixelHash: formula.frontendPngPixelHash,
          frontendSvgHash: formula.frontendSvgHash,
        })),
        pdf: { untouched: untouchedPdf, scrolled: scrolledPdf, reopened: reopenedPdf },
        package: { untouched: untouchedPackage, reopened: reopenedPackage },
        formulaReport,
      },
      null,
      2,
    ),
  );
  console.log("Word untouched image first-frame integration passed");
}

async function runCreatedImageFormulaRegression(
  beforeSessions,
  runPhysicalDoubleClickAfterReopen = false,
) {
  const createdNumbered =
    createImageNumberedRegression || createNativeNumberedRegression;
  const requestedNumbered =
    createNativeNumberedRegression ? false : createdNumbered;
  const createdNativeEquation = createNativeRegression;
  const createdDisplayMode =
    createSourceFormattedEquationRegression ||
    createImageDisplayRegression ||
    createImageNumberedRegression ||
    createNativeDisplayRegression ||
    createNativeNumberedRegression
      ? "block"
      : "inline";
  const defaultCreatedLatex = createSourceFormattedEquationRegression
    ? String.raw`\begin{equation}
\frac{\delta \mathbb{E}[L]}
     {\delta f(\mathbf{x})}
=
2\int
\{f(\mathbf{x})-t\}
p(\mathbf{x},t)\,
\mathrm{d}t
=
0
\end{equation}`
    : createNativeNumberedRegression
      ? String.raw`(a+b)^{n}=\sum_{k=0}^{n}\binom{n}{k}a^{n-k}b^{k}`
      : "dfdfdf";
  const createdLatex = createFormulaLatexOption || defaultCreatedLatex;
  const createdFormulaLetterFont = createFormulaLetterFontOption || "katex";
  const createMacroName = createdNativeEquation
    ? createdDisplayMode === "block"
      ? "VisualTeX_CreateNativeDisplay"
      : "VisualTeX_CreateNativeInline"
    : createdNumbered
      ? "VisualTeX_CreateNumberedDisplay"
      : createdDisplayMode === "block"
        ? "VisualTeX_CreateDisplay"
        : "VisualTeX_CreateInline";
  const requestedDisplayMode = createdDisplayMode;
  if (runPhysicalDoubleClickAfterReopen) {
    rmSync(finalBinaryPhysicalStatusPath, { force: true });
  }
  runAppleScript([
    'tell application "Microsoft Word"',
    "activate",
    "end tell",
  ], 30_000);
  await sleep(3_000);
  let testDocumentName = runAppleScript([
    'tell application "Microsoft Word"',
    "make new document",
    "set testDocument to active document",
    "set testDocumentName to name of testDocument",
    "repeat 3 times",
    "activate object testDocument",
    "delay 0.2",
    "end repeat",
    "activate",
    `run VB macro macro name ${JSON.stringify(createMacroName)}`,
    "return testDocumentName",
    "end tell",
  ], 60_000);

  const sessionId = await waitForWordCreateSession(beforeSessions, 30_000);
  sessionDirectory = join(sessionsRoot, sessionId);
  const request = JSON.parse(
    readFileSync(join(sessionDirectory, "request.json"), "utf8"),
  );
  await stopVisualTeXForManualWordCallback();
  const pendingMarker = request.pendingMarker ?? request.sourceObjectId ?? "";
  const fontSizePt = Number(request.fontSizePt ?? 11);
  if (
    request.mode !== "create" ||
    request.host !== "word" ||
    request.sessionId !== sessionId ||
    request.displayMode !== requestedDisplayMode ||
    request.numbered !== requestedNumbered ||
    request.nativeEquation !== createdNativeEquation ||
    !request.formulaId ||
    !request.sourceDocumentId ||
    !pendingMarker ||
    !Number.isFinite(fontSizePt)
  ) {
    throw new Error(
      `Unexpected Word formula creation request: ${JSON.stringify(request)}`,
    );
  }

  const formula = formulaItem({
    formulaId: request.formulaId,
    latex: createdLatex,
    expectedCodeFormat: createSourceFormattedEquationRegression
      ? "equation"
      : "raw",
    displayMode: createdDisplayMode,
    numbered: createdNumbered,
    fontSizePt,
    artifactDirectory: sessionDirectory,
    formulaLetterFont: createdFormulaLetterFont,
  });
  nativeFiles.push(formula.nativePath);
  const dispatch = manifestText([
    ["protocolVersion", "1"],
    ["sessionId", sessionId],
    ["action", "commit"],
    ["host", "word"],
    ["mode", "create"],
    ["formulaId", formula.formulaId],
    ["displayMode", formula.displayMode],
    ["numbered", createdNumbered ? "1" : "0"],
    ["nativeEquation", createdNativeEquation ? "1" : "0"],
    ["imagePath", formula.imagePath],
    ["vectorDocumentPath", formula.vectorDocumentPath],
    ["fallbackImagePath", formula.fallbackImagePath],
    ["metadata", formula.metadata],
    ["latexBase64", base64Url(formula.latex)],
    ["ommlBase64", formula.ommlBase64],
    ["nativeDocumentPath", formula.nativePath],
    ["pendingMarker", pendingMarker],
    ["sourceMarker", request.sourceObjectId ?? pendingMarker],
    ["sourceDocumentId", request.sourceDocumentId],
    ["widthPoints", formula.widthPoints.toFixed(6)],
    ["heightPoints", formula.heightPoints.toFixed(6)],
    ["baseline", formula.baseline.toFixed(6)],
    ["fontSizePt", formula.fontSizePt.toFixed(6)],
    ["referenceWidthPt", formula.referenceWidthPt.toFixed(6)],
    ["referenceHeightPt", formula.referenceHeightPt.toFixed(6)],
    ["referenceBaselinePt", formula.referenceBaselinePt.toFixed(6)],
  ]);
  writeFileSync(join(sessionDirectory, "dispatch.txt"), dispatch, {
    mode: 0o600,
  });
  writeFileSync(join(sessionsRoot, "word-active-session.txt"), sessionId, {
    mode: 0o600,
  });

  const callbackStatusPath = join(
    sessionDirectory,
    "word-callback-status.txt",
  );
  rmSync(callbackStatusPath, { force: true });
  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(testDocumentName)}`,
    'run VB macro macro name "VisualTeX_ApplyPendingResultForRegression"',
    "end tell",
  ], 90_000);
  if (!existsSync(callbackStatusPath)) {
    throw new Error("Word did not write the formula-create callback status file");
  }
  const callbackStatus = readFileSync(callbackStatusPath, "utf8");
  if (!callbackStatus.startsWith("PASS")) {
    throw new Error(`Word formula-create callback failed:\n${callbackStatus}`);
  }

  let afterCommit = { skipped: "final-binary-physical-focus" };
  let afterCommitInk = { skipped: "final-binary-physical-focus" };
  let afterNativeNormalization = afterCommit;
  let afterNativeNormalizationInk = afterCommitInk;
  let afterSaveReopen = afterCommit;
  let afterSaveReopenInk = afterCommitInk;
  if (!runPhysicalDoubleClickAfterReopen) {
    afterCommit = runFormulaRegressionReport(testDocumentName, [formula]);
    if (!createdNativeEquation) {
      afterCommitInk = createdImagePdfInkBounds(
        testDocumentName,
        "Created dfdfdf formula after commit",
      );
      assertCreatedImageFormulaInk(
        afterCommitInk,
        formula,
        "Created dfdfdf formula after commit",
      );
      runAppleScript([
        'tell application "Microsoft Word"',
        `activate object document ${JSON.stringify(testDocumentName)}`,
        'run VB macro macro name "VisualTeX_MigrateImageMacroButtons"',
        "end tell",
      ], 60_000);
      afterNativeNormalization = runFormulaRegressionReport(
        testDocumentName,
        [formula],
      );
      afterNativeNormalizationInk = createdImagePdfInkBounds(
        testDocumentName,
        "Created dfdfdf formula after native normalization",
      );
      assertCreatedImageFormulaInk(
        afterNativeNormalizationInk,
        formula,
        "Created dfdfdf formula after native normalization",
      );
    } else {
      afterCommitInk = { skipped: "native-equation" };
      afterNativeNormalization = afterCommit;
      afterNativeNormalizationInk = afterCommitInk;
    }
    afterSaveReopen = afterNativeNormalization;
    afterSaveReopenInk = afterNativeNormalizationInk;
    testDocumentName = saveAndReopenWordDocument(testDocumentName);
    await sleep(2_500);
    let afterSaveReopenError;
    for (let attempt = 0; attempt < 3; attempt += 1) {
      try {
        afterSaveReopen = runFormulaRegressionReport(
          testDocumentName,
          [formula],
        );
        afterSaveReopenError = undefined;
        break;
      } catch (reason) {
        afterSaveReopenError = reason;
        await sleep(1_000);
      }
    }
    if (!afterSaveReopen) throw afterSaveReopenError;
    if (!createdNativeEquation) {
      afterSaveReopenInk = createdImagePdfInkBounds(
        testDocumentName,
        "Created dfdfdf formula after save and reopen",
      );
      assertCreatedImageFormulaInk(
        afterSaveReopenInk,
        formula,
        "Created dfdfdf formula after save and reopen",
      );
    }
  }

  let physicalDoubleClickResult = null;
  if (runPhysicalDoubleClickAfterReopen) {
    // Do not click immediately after the callback. Real Word can asynchronously
    // clear an SVG InlineShape.Title after commit, which was the exact state the
    // previous fixture missed. Wait for that host work to drain, then force the
    // observed metadata-only state and require a physical double-click to repair
    // and edit it through AlternativeText.
    await sleep(3_000);
    const documentBeforeAppRestart = runAppleScript([
      'tell application "Microsoft Word"',
      'if not (exists active document) then return "no-active-document"',
      'return name of active document as text',
      'end tell',
    ]);
    await startVisualTeXForPhysicalRegression();
    const documentAfterAppRestart = runAppleScript([
      'tell application "Microsoft Word"',
      'if not (exists active document) then return "no-active-document"',
      'return name of active document as text',
      'end tell',
    ]);
    if (
      documentBeforeAppRestart !== testDocumentName ||
      documentAfterAppRestart !== testDocumentName
    ) {
      throw new Error(
        `Word lost the physical-regression document around the VisualTeX restart: before=${documentBeforeAppRestart} after=${documentAfterAppRestart}`,
      );
    }
    const physicalSelection = runAppleScript([
      'tell application "Microsoft Word"',
      `set documentObject to document ${JSON.stringify(testDocumentName)}`,
      "activate object documentObject",
      "activate",
      'run VB macro macro name "VisualTeX_AssertWordHostSelfTest"',
      ...(createImageNativeMonitorRegression
        ? ['run VB macro macro name "VisualTeX_DisableWordEventsForRegression"']
        : []),
      "set formulaShape to inline shape 1 of documentObject",
      'set title of formulaShape to ""',
      "set formulaRange to text object of formulaShape",
      "select formulaRange",
      'return "image-metadata-only" & (ASCII character 31) & (start of content of formulaRange as text) & (ASCII character 31) & (end of content of formulaRange as text) & (ASCII character 31) & (width of formulaShape as text) & (ASCII character 31) & (height of formulaShape as text) & (ASCII character 31) & (length of (title of formulaShape as text) as text) & (ASCII character 31) & (length of (alternative text of formulaShape as text) as text)',
      "end tell",
    ]);
    const physicalSelectionFields = physicalSelection.split("\x1f");
    if (
      physicalSelectionFields[0] !== "image-metadata-only" ||
      physicalSelectionFields[5] !== "0" ||
      !(Number(physicalSelectionFields[6]) > 0)
    ) {
      throw new Error(
        `Word did not expose the required metadata-only image fixture: ${physicalSelection}`,
      );
    }

    let negativeHitTests = null;
    if (physicalHitTestOnly) {
      const formulaBounds = selectedWordFormulaScreenBounds(testDocumentName);
      const wordBounds = wordFrontWindowScreenBounds();
      const formulaLeft = formulaBounds[0];
      const formulaTop = formulaBounds[1];
      const formulaWidth = formulaBounds[2];
      const formulaHeight = formulaBounds[3];
      const wordLeft = wordBounds[0];
      const wordWidth = wordBounds[2];
      let blankX = formulaLeft + formulaWidth + Math.max(80, formulaWidth);
      if (blankX >= wordLeft + wordWidth - 40) {
        blankX = formulaLeft - Math.max(80, formulaWidth);
      }
      if (blankX <= wordLeft + 20 || blankX >= wordLeft + wordWidth - 20) {
        throw new Error(
          `Unable to choose same-line blank-space click outside formula: ${JSON.stringify({ formulaBounds, wordBounds, blankX })}`,
        );
      }
      const blankY = formulaTop + formulaHeight / 2;
      const beforeBlankClick = currentSessionIds();
      const blankQuartzResult = physicallyDoubleClickAt(blankX, blankY, true);
      await assertNoNewWordEditSession(
        beforeBlankClick,
        "A same-line blank-space double-click",
      );

      runAppleScript([
        'tell application "Microsoft Word"',
        `set documentObject to document ${JSON.stringify(testDocumentName)}`,
        "activate object documentObject",
        "activate",
        "set formulaShape to inline shape 1 of documentObject",
        "select text object of formulaShape",
        "end tell",
      ], 30_000);
      const refreshedWordBounds = wordFrontWindowScreenBounds();
      const ribbonX = refreshedWordBounds[0] + refreshedWordBounds[2] * 0.55;
      const ribbonY = refreshedWordBounds[1] + Math.min(120, refreshedWordBounds[3] * 0.16);
      const beforeRibbonClick = currentSessionIds();
      const ribbonQuartzResult = physicallyDoubleClickAt(ribbonX, ribbonY, false);
      await assertNoNewWordEditSession(
        beforeRibbonClick,
        "A Word Ribbon double-click with a stale formula selection",
      );

      runAppleScript([
        'tell application "Microsoft Word"',
        `set documentObject to document ${JSON.stringify(testDocumentName)}`,
        "activate object documentObject",
        "activate",
        "set formulaShape to inline shape 1 of documentObject",
        "select text object of formulaShape",
        "end tell",
      ], 30_000);
      negativeHitTests = {
        formulaBounds,
        blank: { x: blankX, y: blankY, quartzResult: blankQuartzResult },
        ribbon: { x: ribbonX, y: ribbonY, quartzResult: ribbonQuartzResult },
      };
    }

    const sessionsBeforePhysicalEdit = currentSessionIds();
    const physicalClick = createSourceFormattedEquationRegression
      ? (() => {
          runAppleScript([
            'tell application "Microsoft Word"',
            `activate object document ${JSON.stringify(testDocumentName)}`,
            "activate",
            'run VB macro macro name "VisualTeX_DoubleClickEditSelected"',
            "end tell",
          ], 60_000);
          return {
            skipped:
              "source-formatted equation content uses the same VBA edit entry; physical pane routing is covered separately",
          };
        })()
      : physicallyDoubleClickSelectedWordFormula(testDocumentName);
    const physicalEditSessionId = await waitForFormulaEditSession(
      sessionsBeforePhysicalEdit,
      formula.formulaId,
      120_000,
    );
    editSessionDirectories.push(join(sessionsRoot, physicalEditSessionId));
    const physicalEditSession = validateFormulaEditSession(
      physicalEditSessionId,
      formula,
      formula.codeFormat,
      formula.metadataLines,
    );
    const editorReadiness = physicalHitTestOnly
      ? await waitForPhysicalEditorVisible(
          physicalEditSessionId,
          formula.formulaId,
        )
      : await waitForPhysicalEditorReadiness(
          physicalEditSessionId,
          formula.formulaId,
        );
    await assertSinglePhysicalEditSession(
      sessionsBeforePhysicalEdit,
      physicalEditSessionId,
      formula.formulaId,
    );
    if (physicalHitTestOnly) {
      physicalDoubleClickResult = {
        sessionId: physicalEditSessionId,
        formulaId: formula.formulaId,
        selection: physicalSelectionFields,
        physicalClick,
        negativeHitTests,
        editorReadiness,
      };
      writeFileSync(
        finalBinaryPhysicalStatusPath,
        JSON.stringify(
          {
            status: "PASS",
            revision: "word-double-click-hit-test-20260809-r1",
            ...physicalDoubleClickResult,
          },
          null,
          2,
        ),
        { mode: 0o600 },
      );
      console.log(JSON.stringify(physicalDoubleClickResult, null, 2));
      console.log("Word physical double-click hit-test integration passed");
      return;
    }
    const pictureFormatUi = wordPictureFormatUiSnapshot();
    if (pictureFormatUi.pictureFormatVisible) {
      throw new Error(
        `Word displayed its native picture-format UI for the final binary formula double-click: ${pictureFormatUi.raw}`,
      );
    }
    let editorUi = null;
    let editorClose = null;
    if (createSourceFormattedEquationRegression) {
      editorUi = visualTeXEditorUiSnapshot();
      const expectedCharacterCount = formula.metadataLines[0].length;
      if (
        editorUi.lineCount !== 1 ||
        editorUi.characterCount !== expectedCharacterCount ||
        editorUi.errorMatches.length > 0
      ) {
        throw new Error(
          `The source-formatted equation did not load as one complete editable formula: ${JSON.stringify({ expectedCharacterCount, editorUi })}`,
        );
      }
      const closeRequest = closeVisualTeXEditorByWindowButton();
      if (!closeRequest.pressed) {
        throw new Error(
          `The VisualTeX formula editor close button was not pressed: ${JSON.stringify(closeRequest)}`,
        );
      }
      const closeState = await waitForVisualTeXEditorToClose();
      editorClose = { closeRequest, closeState };
    }
    physicalDoubleClickResult = {
      sessionId: physicalEditSessionId,
      formulaId: formula.formulaId,
      selection: physicalSelectionFields,
      documentStage: "after-image-commit-metadata-only",
      codeFormat: physicalEditSession.normalized.codeFormat,
      lines: physicalEditSession.normalized.lines.map((line) => line.latex),
      physicalClick,
      editorReadiness,
      pictureFormatUi,
      editorUi,
      editorClose,
    };
    writeFileSync(
      finalBinaryPhysicalStatusPath,
      JSON.stringify(
        {
          status: "PASS",
          revision: "word-office-performance-20260801-r77",
          ...physicalDoubleClickResult,
        },
        null,
        2,
      ),
      { mode: 0o600 },
    );
  }

  console.log(
    JSON.stringify(
      {
        sessionId,
        formulaId: formula.formulaId,
        latex: formula.latex,
        reports: {
          afterCommit,
          afterNativeNormalization,
          afterSaveReopen,
        },
        visibleInk: {
          afterCommit: afterCommitInk,
          afterNativeNormalization: afterNativeNormalizationInk,
          afterSaveReopen: afterSaveReopenInk,
        },
        ...(physicalDoubleClickResult
          ? { physicalDoubleClick: physicalDoubleClickResult }
          : {}),
      },
      null,
      2,
    ),
  );
  console.log(
    runPhysicalDoubleClickAfterReopen
      ? "Word final binary image physical double-click integration passed"
      : "Word image formula creation integration passed",
  );
}

const before = currentSessionIds();
let sessionDirectory = "";
let installedWordAddinBackedUp = false;
const nativeFiles = [];
const editSessionDirectories = [];

try {
  mkdirSync(sessionsRoot, { recursive: true });
  mkdirSync(nativeRoot, { recursive: true });
  mkdirSync(officeScratchRoot, { recursive: true });
  mkdirSync(wordStartupRoot, { recursive: true });
  try {
    runAppleScript([
      'tell application "Microsoft Word" to quit saving no',
    ], 20_000);
  } catch {
    // Continue with a hard process cleanup below.
  }
  spawnSync("/usr/bin/killall", ["Microsoft Word"], {
    encoding: "utf8",
    timeout: 10_000,
  });
  await sleep(2_000);
  if (existsSync(installedWordAddinPath)) {
    copyFileSync(installedWordAddinPath, installedWordAddinBackupPath);
    rmSync(installedWordAddinPath, { force: true });
    installedWordAddinBackedUp = true;
  }
  // Use the reviewed DOTM as a real global template for every integration
  // run. This is required for Word's native double-click event sink and legacy
  // image-field migration to remain available after DOCX save/reopen; the
  // user's previous Startup add-in is restored in finally.
  if (!existsSync(activeTemplatePath)) {
    throw new Error(`The requested Word add-in does not exist: ${activeTemplatePath}`);
  }
  copyFileSync(activeTemplatePath, installedWordAddinPath);
  await waitForWordAutomationReady();
  if (pictureRoutingPerformance) {
    await runPictureRoutingPerformanceRegression();
  } else if (pictureRoutingTarget) {
    await runPictureRoutingNativeRegression(before);
  } else if (firstFrameImageRegression) {
    await runFirstFrameImageRegression(before);
  } else if (createFormulaRegression) {
    await runCreatedImageFormulaRegression(
      before,
      createImagePhysicalRegression,
    );
  } else {
  let testDocumentName = runAppleScript([
    'tell application "Microsoft Word"',
    "make new document",
    "set testDocument to active document",
    "activate",
    'run VB macro macro name "VisualTeX_InsertLatexMarkdownDocument"',
    "return name of testDocument",
    "end tell",
  ], 60_000);

  const sessionId = await waitForNewSession(before);
  sessionDirectory = join(sessionsRoot, sessionId);
  const request = JSON.parse(
    readFileSync(join(sessionDirectory, "request.json"), "utf8"),
  );
  if (
    request.operation !== "documentImport" ||
    request.sessionId !== sessionId ||
    request.host !== "word"
  ) {
    throw new Error(`Unexpected Word document import request: ${JSON.stringify(request)}`);
  }
  await stopVisualTeXForManualWordCallback();

  const formulas = [
    formulaItem({
      formulaId: crypto.randomUUID(),
      latex: "101=202",
      displayMode: "inline",
      numbered: false,
      fontSizePt: 11,
      artifactDirectory: sessionDirectory,
    }),
    formulaItem({
      formulaId: crypto.randomUUID(),
      latex: "12345=67890",
      displayMode: "block",
      numbered: false,
      fontSizePt: 14,
      artifactDirectory: sessionDirectory,
    }),
    formulaItem({
      formulaId: crypto.randomUUID(),
      latex: "24680=13579",
      displayMode: "block",
      numbered: true,
      fontSizePt: 18,
      artifactDirectory: sessionDirectory,
    }),
    formulaItem({
      formulaId: crypto.randomUUID(),
      latex: String.raw`\sum_{i=1}^{\infty}{a_k\left( x-x_0 \right) ^k}`,
      displayMode: "inline",
      numbered: false,
      fontSizePt: 12,
      artifactDirectory: sessionDirectory,
    }),
    formulaItem({
      formulaId: crypto.randomUUID(),
      latex: String.raw`\sum_{i=1}^{\infty}{a_kP_k\left( x \right)}`,
      displayMode: "inline",
      numbered: false,
      fontSizePt: 12,
      artifactDirectory: sessionDirectory,
    }),
    formulaItem({
      formulaId: crypto.randomUUID(),
      latex: legacyAlignLatex,
      expectedCodeFormat: "align",
      displayMode: "block",
      numbered: false,
      fontSizePt: 14,
      artifactDirectory: sessionDirectory,
    }),
    formulaItem({
      formulaId: crypto.randomUUID(),
      latex: legacyAlignStarLatex,
      expectedCodeFormat: "align-star",
      displayMode: "block",
      numbered: false,
      fontSizePt: 14,
      artifactDirectory: sessionDirectory,
    }),
  ];
  nativeFiles.push(...formulas.map((formula) => formula.nativePath));

  const bodyParagraphId = crypto.randomUUID();
  const followingParagraphId = crypto.randomUUID();
  const headingParagraphId = crypto.randomUUID();
  const bulletParagraphId = crypto.randomUUID();
  const bulletParagraph2Id = crypto.randomUUID();
  const bulletFormulaParagraphId = crypto.randomUUID();
  const numberParagraphId = crypto.randomUUID();
  const endingParagraphId = crypto.randomUUID();
  const items = [];
  appendText(items, 0, "结构化测试", {
    id: headingParagraphId,
    style: "heading1",
    alignment: "left",
    listKind: "none",
    listLevel: 0,
    start: true,
    end: true,
  });
  appendText(items, 1, "开头文字：", {
    id: bodyParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "none",
    listLevel: 0,
    start: true,
    end: false,
  });
  appendFormula(items, 2, formulas[0], {
    id: bodyParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "none",
    listLevel: 0,
    start: false,
    end: false,
  });
  appendText(items, 3, "，行内公式之后。", {
    id: bodyParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "none",
    listLevel: 0,
    start: false,
    end: true,
  });
  appendFormula(items, 4, formulas[1]);
  appendText(items, 5, "未编号行间公式之后。", {
    id: followingParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "none",
    listLevel: 0,
    start: true,
    end: true,
  });
  appendFormula(items, 6, formulas[2]);
  appendText(items, 7, "多项式逼近的逼近系数和原函数原则上没有硬性关系。", {
    id: bulletParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "bullet",
    listLevel: 1,
    start: true,
    end: true,
  });
  appendText(items, 8, "多项式逼近是全局性的，而幂级数逼近有收敛半径。", {
    id: bulletParagraph2Id,
    style: "normal",
    alignment: "left",
    listKind: "bullet",
    listLevel: 1,
    start: true,
    end: true,
  });
  appendText(items, 9, "形式上的区别：幂级数的形式是", {
    id: bulletFormulaParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "bullet",
    listLevel: 1,
    start: true,
    end: false,
  });
  appendFormula(items, 10, formulas[3], {
    id: bulletFormulaParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "bullet",
    listLevel: 1,
    start: false,
    end: false,
  });
  appendText(items, 11, "而多项式级数的形式是", {
    id: bulletFormulaParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "bullet",
    listLevel: 1,
    start: false,
    end: false,
  });
  appendFormula(items, 12, formulas[4], {
    id: bulletFormulaParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "bullet",
    listLevel: 1,
    start: false,
    end: true,
  });
  appendText(items, 13, "编号列表正文", {
    id: numberParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "number",
    listLevel: 1,
    start: true,
    end: true,
  });
  appendFormula(items, 14, formulas[5]);
  appendFormula(items, 15, formulas[6]);
  appendText(items, 16, "结尾文字。", {
    id: endingParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "none",
    listLevel: 0,
    start: true,
    end: true,
  });

  const manifestPath = join(sessionDirectory, "document-import.txt");
  const entries = [
    ["protocolVersion", "1"],
    ["sessionId", sessionId],
    ["operation", "documentImport"],
    ["outputKind", outputKind],
    ["sourceDocumentId", request.sourceDocumentId],
    ["bookmarkName", request.documentImport.bookmarkName],
    ["itemCount", String(diagnosticItemLimit)],
    ...items,
  ];
  writeFileSync(manifestPath, manifestText(entries), { mode: 0o600 });

  const dispatch = manifestText([
    ["protocolVersion", "1"],
    ["sessionId", sessionId],
    ["action", "documentCommit"],
    ["host", "word"],
    ["sourceDocumentId", request.sourceDocumentId],
    ["bookmarkName", request.documentImport.bookmarkName],
    ["documentImportPath", manifestPath],
  ]);
  writeFileSync(join(sessionDirectory, "dispatch.txt"), dispatch, { mode: 0o600 });
  writeFileSync(join(sessionsRoot, "word-active-session.txt"), sessionId, { mode: 0o600 });

  const bookmarkPreflight = runAppleScript([
    'tell application "Microsoft Word"',
    `set documentObject to document ${JSON.stringify(testDocumentName)}`,
    "activate object documentObject",
    `set targetExists to exists bookmark ${JSON.stringify(request.documentImport.bookmarkName)} of documentObject`,
    "set bookmarkNames to name of every bookmark of documentObject",
    'return (targetExists as text) & (ASCII character 31) & (bookmarkNames as text)',
    "end tell",
  ]);
  if (!bookmarkPreflight.startsWith("true\x1f")) {
    throw new Error(
      `The Word document-import bookmark disappeared before callback: ${bookmarkPreflight}`,
    );
  }

  const callbackStatusPath = join(
    sessionDirectory,
    "word-callback-status.txt",
  );
  const documentImportProgressPath = join(
    sessionDirectory,
    "document-import-progress.txt",
  );
  rmSync(callbackStatusPath, { force: true });
  rmSync(documentImportProgressPath, { force: true });
  const observedDocumentImportProgress = [];
  const callbackProcess = startAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(testDocumentName)}`,
    'run VB macro macro name "VisualTeX_ConfigureDocumentImportParagraphSpacingRegression"',
    'run VB macro macro name "VisualTeX_ApplyPendingResultForRegression"',
    "end tell",
  ]);
  const callbackDeadline = Date.now() + 90_000;
  while (!callbackProcess.isSettled() && Date.now() < callbackDeadline) {
    const progress = readDocumentImportProgress(documentImportProgressPath);
    const previous = observedDocumentImportProgress.at(-1);
    if (
      progress &&
      (!previous ||
        previous.current !== progress.current ||
        previous.total !== progress.total ||
        previous.stage !== progress.stage)
    ) {
      observedDocumentImportProgress.push(progress);
    }
    await sleep(12);
  }
  if (!callbackProcess.isSettled()) {
    callbackProcess.child.kill("SIGTERM");
    throw new Error("Word document-import callback timed out after 90 seconds");
  }
  const callbackResult = await callbackProcess.completion;
  if (callbackResult.status !== 0) {
    throw new Error(
      [
        callbackResult.stderr.trim(),
        callbackResult.stdout.trim(),
        callbackResult.signal ? `signal=${callbackResult.signal}` : "",
        `status=${String(callbackResult.status)}`,
      ]
        .filter(Boolean)
        .join("\n") || "Word document-import AppleScript failed",
    );
  }
  const finalDocumentImportProgress = readDocumentImportProgress(
    documentImportProgressPath,
  );
  if (
    !finalDocumentImportProgress ||
    finalDocumentImportProgress.current !== diagnosticItemLimit ||
    finalDocumentImportProgress.total !== diagnosticItemLimit ||
    finalDocumentImportProgress.stage !== "complete"
  ) {
    throw new Error(
      `Word document-import progress did not finish correctly: ${JSON.stringify({
        finalDocumentImportProgress,
        observedDocumentImportProgress,
      })}`,
    );
  }
  if (
    !observedDocumentImportProgress.some(
      (progress) =>
        progress.stage === "inserting" &&
        progress.current > 0 &&
        progress.current < progress.total,
    )
  ) {
    throw new Error(
      `Word document-import progress did not expose an intermediate insertion count: ${JSON.stringify(observedDocumentImportProgress)}`,
    );
  }
  for (let index = 1; index < observedDocumentImportProgress.length; index += 1) {
    if (
      observedDocumentImportProgress[index].current <
      observedDocumentImportProgress[index - 1].current
    ) {
      throw new Error(
        `Word document-import progress moved backwards: ${JSON.stringify(observedDocumentImportProgress)}`,
      );
    }
  }
  if (!existsSync(callbackStatusPath)) {
    throw new Error("Word did not write the regression callback status file");
  }
  const callbackStatus = readFileSync(callbackStatusPath, "utf8");
  if (!callbackStatus.startsWith("PASS")) {
    throw new Error(`Word document-import callback failed:\n${callbackStatus}`);
  }
  if (diagnosticItemLimit < 17) {
    throw new Error(`${diagnosticSuccessPrefix}${diagnosticItemLimit}`);
  }
  let formulaRegressionReport = runFormulaRegressionReport(
    testDocumentName,
    formulas,
  );
  const initialFormulaContainerReport = inspectWordFormulaContainers(
    testDocumentName,
    formulas,
    "after-import",
  );

  const pdfPath = coordinatePdfPath;
  rmSync(pdfPath, { force: true });
  const bookmarkNames = formulas.map((formula) => nativeBookmark(formula.formulaId));
  const numberedCompactId = compactFormulaId(formulas[2].formulaId);
  const numberBookmarkName = `VT_R_${numberedCompactId}`;
  const inspectionLines = [
    'tell application "Microsoft Word"',
    `set documentObject to document ${JSON.stringify(testDocumentName)}`,
    "activate object documentObject",
    "activate",
    "set documentText to content of text object of documentObject",
    "set bookmarkNames to name of every bookmark of documentObject",
    "set tableCount to count of tables of documentObject",
    "set shapeCount to count of inline shapes of documentObject",
    "set pageSetupObject to page setup of section 1 of documentObject",
    "set pageWidthValue to page width of pageSetupObject",
    "set pageHeightValue to page height of pageSetupObject",
    "set leftMarginValue to left margin of pageSetupObject",
    "set rightMarginValue to right margin of pageSetupObject",
  ];
  if (outputKind === "omml") {
    inspectionLines.push(
      'set alternativeTexts to ""',
      `set inlineSize to font size of font object of text object of bookmark ${JSON.stringify(bookmarkNames[0])} of documentObject`,
      `set displaySize to font size of font object of text object of bookmark ${JSON.stringify(bookmarkNames[1])} of documentObject`,
      `set numberedSize to font size of font object of text object of bookmark ${JSON.stringify(bookmarkNames[2])} of documentObject`,
      `set displayFormulaRange to text object of bookmark ${JSON.stringify(bookmarkNames[1])} of documentObject`,
      "set displayStartPosition to start of content of displayFormulaRange",
      "set displayEndPosition to end of content of displayFormulaRange",
      "set displayStartRange to create range documentObject start displayStartPosition end displayStartPosition",
      "set displayEndRange to create range documentObject start displayEndPosition end displayEndPosition",
      "set displayLeft to get range information displayStartRange information type horizontal position relative to page",
      "set displayRight to get range information displayEndRange information type horizontal position relative to page",
      "set displayTop to get range information displayStartRange information type vertical position relative to page",
      "set inlineWidth to 0",
      "set inlineHeight to 0",
      "set displayHeight to 0",
      `set numberedFormulaRange to text object of bookmark ${JSON.stringify(bookmarkNames[2])} of documentObject`,
      "set numberedStartPosition to start of content of numberedFormulaRange",
      "set numberedEndPosition to end of content of numberedFormulaRange",
      "set numberedStartRange to create range documentObject start numberedStartPosition end numberedStartPosition",
      "set numberedEndRange to create range documentObject start numberedEndPosition end numberedEndPosition",
      "set numberedLeft to get range information numberedStartRange information type horizontal position relative to page",
      "set numberedRight to get range information numberedEndRange information type horizontal position relative to page",
      "set numberedTop to get range information numberedStartRange information type vertical position relative to page",
      "set numberedHeight to 0",
    );
  } else {
    inspectionLines.push(
      "set inlineShapeObject to inline shape 1 of documentObject",
      "set displayShapeObject to inline shape 2 of documentObject",
      "set numberedShapeObject to inline shape 3 of documentObject",
      "set listFormulaShapeObject1 to inline shape 4 of documentObject",
      "set listFormulaShapeObject2 to inline shape 5 of documentObject",
      "set alignShapeObject to inline shape 6 of documentObject",
      "set alignStarShapeObject to inline shape 7 of documentObject",
      'set alternativeTexts to (alternative text of inlineShapeObject) & "|" & (alternative text of displayShapeObject) & "|" & (alternative text of numberedShapeObject) & "|" & (alternative text of listFormulaShapeObject1) & "|" & (alternative text of listFormulaShapeObject2) & "|" & (alternative text of alignShapeObject) & "|" & (alternative text of alignStarShapeObject)',
      "set inlineSize to font size of font object of text object of inlineShapeObject",
      "set inlineWidth to width of inlineShapeObject",
      "set inlineHeight to height of inlineShapeObject",
      "set displaySize to font size of font object of text object of displayShapeObject",
      "set numberedSize to font size of font object of text object of numberedShapeObject",
      "set displayFormulaRange to text object of displayShapeObject",
      "set displayStartPosition to start of content of displayFormulaRange",
      "set displayStartRange to create range documentObject start displayStartPosition end displayStartPosition",
      "set displayLeft to get range information displayStartRange information type horizontal position relative to page",
      "set displayRight to displayLeft + (width of displayShapeObject)",
      "set displayTop to get range information displayStartRange information type vertical position relative to page",
      "set displayHeight to height of displayShapeObject",
      "set numberedFormulaRange to text object of numberedShapeObject",
      "set numberedStartPosition to start of content of numberedFormulaRange",
      "set numberedStartRange to create range documentObject start numberedStartPosition end numberedStartPosition",
      "set numberedLeft to get range information numberedStartRange information type horizontal position relative to page",
      "set numberedRight to numberedLeft + (width of numberedShapeObject)",
      "set numberedTop to get range information numberedStartRange information type vertical position relative to page",
      "set numberedHeight to height of numberedShapeObject",
    );
  }
  inspectionLines.push(
    `set numberRange to text object of bookmark ${JSON.stringify(numberBookmarkName)} of documentObject`,
    "set numberStartPosition to start of content of numberRange",
    "set numberEndPosition to end of content of numberRange",
    "set numberStartRange to create range documentObject start numberStartPosition end numberStartPosition",
    "set numberEndRange to create range documentObject start numberEndPosition end numberEndPosition",
    "set numberLeft to get range information numberStartRange information type horizontal position relative to page",
    "set numberRight to get range information numberEndRange information type horizontal position relative to page",
    "set numberTop to get range information numberStartRange information type vertical position relative to page",
    'return documentText & "\n---VT---\n" & (bookmarkNames as text) & "\n---VT---\n" & alternativeTexts & "\n---VT---\n" & shapeCount & "," & tableCount & "," & inlineSize & "," & displaySize & "," & numberedSize & "," & inlineWidth & "," & inlineHeight & "," & pageWidthValue & "," & pageHeightValue & "," & leftMarginValue & "," & rightMarginValue & "," & displayLeft & "," & displayRight & "," & displayTop & "," & displayHeight & "," & numberedLeft & "," & numberedRight & "," & numberedTop & "," & numberedHeight & "," & numberLeft & "," & numberRight & "," & numberTop',
    "end tell",
  );
  const inspection = runAppleScript(inspectionLines);
  rmSync(pdfExportStatusPath, { force: true });
  writeFileSync(pdfExportRequestPath, pdfPath, { mode: 0o600 });
  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(testDocumentName)}`,
    'run VB macro macro name "VisualTeX_ExportActiveDocumentPdfForRegression"',
    "end tell",
  ], 90_000);
  const exportStatus = existsSync(pdfExportStatusPath)
    ? readFileSync(pdfExportStatusPath, "utf8").trim()
    : "missing-status";
  if (!exportStatus.startsWith("ok|")) {
    throw new Error(`Word PDF regression export failed: ${exportStatus}`);
  }
  if (!existsSync(pdfPath)) {
    throw new Error(`Word did not export the coordinate verification PDF: ${pdfPath}`);
  }
  const swiftGeometryArguments = [
    join(repositoryRoot, "scripts/pdf_formula_geometry.swift"),
    pdfPath,
    ...(outputKind === "omml"
      ? [
          formulas[1].pdfToken,
          formulas[2].pdfToken,
          formulas[5].pdfToken,
          formulas[6].pdfToken,
        ]
      : ["--number-only"]),
  ];
  const swiftGeometry = spawnSync(
    "/usr/bin/swift",
    swiftGeometryArguments,
    {
      encoding: "utf8",
      timeout: 120_000,
      maxBuffer: 8 * 1024 * 1024,
    },
  );
  if (swiftGeometry.status !== 0) {
    throw new Error(
      swiftGeometry.stderr.trim() || "PDFKit formula geometry extraction failed",
    );
  }
  const renderedGeometry = JSON.parse(swiftGeometry.stdout);
  const ommlAlignmentGeometry = [];
  if (outputKind === "omml") {
    if (!Array.isArray(renderedGeometry.aligned) || renderedGeometry.aligned.length !== 2) {
      throw new Error(
        `PDF regression did not return both aligned OMML formulas: ${swiftGeometry.stdout}`,
      );
    }
    for (const [index, alignedFormula] of renderedGeometry.aligned.entries()) {
      const positions = alignedFormula.relationshipXs;
      if (
        !Array.isArray(positions) ||
        positions.length !== 2 ||
        positions.some((position) => !Number.isFinite(position))
      ) {
        throw new Error(
          `Aligned OMML formula ${index + 1} did not expose both PDF relationship positions: ${JSON.stringify(alignedFormula)}`,
        );
      }
      const spread = Math.max(...positions) - Math.min(...positions);
      if (spread > 0.5) {
        throw new Error(
          `Aligned OMML formula ${index + 1} has a misaligned relationship column: ${JSON.stringify({ positions, spread })}`,
        );
      }
      ommlAlignmentGeometry.push({ positions, spread });
    }
  }

  const [documentText, bookmarkText, alternativeText, numericText] =
    inspection.split("\n---VT---\n");
  const [
    shapeCount,
    tableCount,
    inlineSize,
    displaySize,
    numberedSize,
    inlineWidth,
    inlineHeight,
    pageWidth,
    pageHeight,
    leftMargin,
    rightMargin,
    displayLeft,
    displayRight,
    displayTop,
    displayHeight,
    numberedLeft,
    numberedRight,
    numberedTop,
    numberedHeight,
    numberLeft,
    numberRight,
    numberTop,
  ] = numericText.split(",").map(Number);

  for (const expected of ["开头文字：", "行内公式之后。", "未编号行间公式之后。", "结尾文字。"]) {
    if (!documentText.includes(expected)) {
      throw new Error(`Word import text is missing ${expected}: ${JSON.stringify(documentText)}`);
    }
  }
  if (outputKind === "omml") {
    for (const bookmarkName of bookmarkNames) {
      if (!bookmarkText.includes(bookmarkName)) {
        throw new Error(`Word import is missing formula bookmark ${bookmarkName}`);
      }
    }
    if (shapeCount !== 0) {
      throw new Error(`OMML import unexpectedly created ${shapeCount} inline shapes`);
    }
  } else {
    if (shapeCount !== formulas.length) {
      throw new Error(
        `Image import created ${shapeCount} inline shapes instead of ${formulas.length}`,
      );
    }
    const metadataPayloads = alternativeText.split("|");
    if (
      metadataPayloads.length !== formulas.length ||
      metadataPayloads.some((value) => !value.startsWith("visualtex:v1:deflate:"))
    ) {
      throw new Error(
        "Image formulas did not retain independent VisualTeX metadata payloads",
      );
    }
    metadataPayloads.forEach((payload, index) => {
      const metadata = decodeFormulaMetadata(payload);
      const expected = formulas[index];
      if (
        !metadata ||
        metadata.formulaId !== expected.formulaId ||
        metadata.displayMode !== expected.displayMode ||
        Boolean(metadata.numbered) !== expected.numbered ||
        Math.abs((metadata.fontSizePt ?? 0) - expected.fontSizePt) > 0.001
      ) {
        throw new Error(
          `Image formula ${index + 1} did not retain its independent identity, mode, numbering and font size`,
        );
      }
    });
  }
  for (const numberBookmark of [
    `VT_R_${numberedCompactId}`,
    `VT_N_${numberedCompactId}`,
    `VT_C_${numberedCompactId}`,
  ]) {
    if (!bookmarkText.includes(numberBookmark)) {
      throw new Error(`Numbered display formula is missing ${numberBookmark}`);
    }
  }
  const geometryValues = [
    pageWidth,
    pageHeight,
    leftMargin,
    rightMargin,
    inlineWidth,
    inlineHeight,
    displayLeft,
    displayRight,
    displayTop,
    displayHeight,
    numberedLeft,
    numberedRight,
    numberedTop,
    numberedHeight,
    numberLeft,
    numberRight,
    numberTop,
  ];
  if (geometryValues.some((value) => !Number.isFinite(value) || value < 0)) {
    throw new Error(`Word returned invalid formula geometry: ${JSON.stringify(geometryValues)}`);
  }
  const textBoundaryLeft = leftMargin;
  const textBoundaryRight = pageWidth - rightMargin;
  const textBoundaryCenter = (textBoundaryLeft + textBoundaryRight) / 2;
  const displayCenter = (displayLeft + displayRight) / 2;
  const numberedCenter = (numberedLeft + numberedRight) / 2;
  const displayCenterError = Math.abs(displayCenter - textBoundaryCenter);
  const numberedCenterError = Math.abs(numberedCenter - textBoundaryCenter);
  const displayToNumberedCenterError = Math.abs(displayCenter - numberedCenter);
  const imageRasterGeometry =
    outputKind === "image"
      ? resolveImageRasterGeometry(
          renderedGeometry.rasterBands ?? [],
          textBoundaryCenter,
          {
            pageHeight,
            displayTop,
            displayHeight,
            numberedTop,
            numberedHeight,
            numberTop,
          },
        )
      : null;
  const measuredUnnumberedCenter =
    outputKind === "omml"
      ? renderedGeometry.unnumbered.centerX
      : imageRasterGeometry.unnumbered.centerX;
  const measuredNumberedCenter =
    outputKind === "omml"
      ? renderedGeometry.numbered.centerX
      : imageRasterGeometry.numbered.centerX;
  const renderedUnnumberedCenterError = Math.abs(
    measuredUnnumberedCenter - textBoundaryCenter,
  );
  const renderedNumberedCenterError = Math.abs(
    measuredNumberedCenter - textBoundaryCenter,
  );
  const renderedFormulaCenterDifference = Math.abs(
    measuredUnnumberedCenter - measuredNumberedCenter,
  );
  let imageVisualCalibration = null;

  const sizes = [inlineSize, displaySize, numberedSize];
  const expectedSizes = [11, 14, 18];
  if (outputKind === "omml") {
    sizes.forEach((size, index) => {
      if (!Number.isFinite(size) || Math.abs(size - expectedSizes[index]) > 0.1) {
        throw new Error(`Formula ${index + 1} font size mismatch: ${size}`);
      }
    });
  } else {
    if (sizes.some((size) => !Number.isFinite(size) || size <= 0)) {
      throw new Error(`Word returned invalid image Range.Font.Size values: ${sizes.join(",")}`);
    }
    const actualDimensions = [
      [inlineWidth, inlineHeight],
      [displayRight - displayLeft, displayHeight],
      [numberedRight - numberedLeft, numberedHeight],
    ];
    actualDimensions.forEach(([width, height], index) => {
      const expected = formulas[index];
      if (
        Math.abs(width - expected.widthPoints) > 0.15 ||
        Math.abs(height - expected.heightPoints) > 0.15
      ) {
        throw new Error(
          `Image formula ${index + 1} visual geometry does not match its independent font size: ` +
            `${JSON.stringify({ width, height, expectedWidth: expected.widthPoints, expectedHeight: expected.heightPoints })}`,
        );
      }
    });
    for (const [label, inkBounds, wordShapeWidth] of [
      [
        "unnumbered display",
        imageRasterGeometry.unnumberedInk,
        displayRight - displayLeft,
      ],
      [
        "numbered display",
        imageRasterGeometry.numberedInk,
        numberedRight - numberedLeft,
      ],
    ]) {
      const minimumInkWidth = Math.max(20, wordShapeWidth * 0.3);
      if (!inkBounds || inkBounds.width < minimumInkWidth) {
        throw new Error(
          `${label} formula rendered as a narrow fallback glyph instead of ` +
            `its complete SVG: ${JSON.stringify({ inkBounds, wordShapeWidth, minimumInkWidth })}`,
        );
      }
    }
    const calibrationInk = imageRasterGeometry.unnumberedInk;
    const nativeWidthRatio = calibrationInk.width / nativeCalibrationWidthPt;
    const boxToInkHeightRatio = displayHeight / calibrationInk.height;
    if (nativeWidthRatio < 0.95 || nativeWidthRatio > 1.05) {
      throw new Error(
        `The 14 pt Word image formula does not visually match the native ` +
          `Cambria Math calibration: ${JSON.stringify({ calibrationInk, nativeCalibrationWidthPt, nativeWidthRatio })}`,
      );
    }
    if (boxToInkHeightRatio > 1.8) {
      throw new Error(
        `The Word image formula retains excessive transparent vertical padding: ` +
          `${JSON.stringify({ displayHeight, calibrationInk, boxToInkHeightRatio })}`,
      );
    }
    imageVisualCalibration = {
      nativeCalibrationWidthPt,
      imageInkWidthPt: calibrationInk.width,
      nativeWidthRatio,
      wordShapeHeightPt: displayHeight,
      imageInkHeightPt: calibrationInk.height,
      boxToInkHeightRatio,
      visualScale: wordImageVisualScalesForFont("katex"),
      displayPaddingPx: wordDisplayPaddingPx,
    };
  }

  const centerTolerancePt = outputKind === "image" ? 0.5 : 0.25;
  if (Math.abs(renderedGeometry.pageWidth - pageWidth) > 0.25) {
    throw new Error(
      `Word/PDF page width mismatch: Word=${pageWidth}, PDF=${renderedGeometry.pageWidth}`,
    );
  }
  if (outputKind === "omml") {
    if (renderedUnnumberedCenterError > centerTolerancePt) {
      throw new Error(
        `Unnumbered display formula is not centered: error=${renderedUnnumberedCenterError} pt`,
      );
    }
    if (renderedNumberedCenterError > centerTolerancePt) {
      throw new Error(
        `Numbered display formula is not centered: error=${renderedNumberedCenterError} pt`,
      );
    }
  } else {
    // Word's horizontal position for an InlineShape Range is a paragraph/text
    // anchor, not the centered visual image edge. The symmetric fixture places
    // its relationship sign at the image center, so the PDF raster marker must
    // cross the text-area center. Its glyph ink center can be a few points off
    // because the '=' outline and SVG padding are not geometrically symmetric.
    for (const [label, geometry] of [
      ["unnumbered", imageRasterGeometry.unnumbered],
      ["numbered", imageRasterGeometry.numbered],
    ]) {
      if (
        textBoundaryCenter < geometry.minX - centerTolerancePt ||
        textBoundaryCenter > geometry.maxX + centerTolerancePt
      ) {
        throw new Error(
          `${label} image formula center marker does not cross the text-area center: ${JSON.stringify({ geometry, textBoundaryCenter, centerTolerancePt })}`,
        );
      }
    }
  }
  if (renderedFormulaCenterDifference > centerTolerancePt) {
    throw new Error(
      `Numbering shifted the formula center marker: difference=${renderedFormulaCenterDifference} pt`,
    );
  }
  const equationNumberGeometry =
    outputKind === "omml"
      ? renderedGeometry.equationNumber
      : imageRasterGeometry.equationNumber;
  const numberedFormulaPdfCenterY =
    outputKind === "omml"
      ? renderedGeometry.numbered.centerY
      : imageRasterGeometry.numbered.centerY;
  const equationNumberInkCenterDifference = Math.abs(
    equationNumberGeometry.centerY - numberedFormulaPdfCenterY,
  );
  const equationNumberVerticalError =
    outputKind === "omml"
      ? equationNumberInkCenterDifference
      : Math.abs(numberTop - numberedTop);
  if (outputKind === "omml") {
    if (equationNumberVerticalError > 0.25) {
      throw new Error(
        `Equation number is not vertically centered with its formula: error=${equationNumberVerticalError} pt`,
      );
    }
  } else {
    const rasterTolerancePt = 1.5;
    if (equationNumberVerticalError > 0.25) {
      throw new Error(
        `Image equation number and formula outer boxes do not share a top edge: ${JSON.stringify({ numberedTop, numberTop, equationNumberVerticalError })}`,
      );
    }
    for (const [label, geometry] of [
      ["formula ink", imageRasterGeometry.numberedInk],
      ["equation number", equationNumberGeometry],
    ]) {
      if (!geometry || geometry.height > numberedHeight + rasterTolerancePt) {
        throw new Error(
          `Numbered image ${label} is taller than the Word image outer box: ${JSON.stringify({ geometry, numberedHeight, rasterTolerancePt })}`,
        );
      }
    }
    if (equationNumberInkCenterDifference > 0.5) {
      throw new Error(
        `Image equation number is not vertically centered with the formula ink: ${JSON.stringify({ equationNumberInkCenterDifference, numberedFormula: imageRasterGeometry.numberedInk, equationNumber: equationNumberGeometry })}`,
      );
    }
  }
  const measuredNumberedRight =
    outputKind === "omml"
      ? renderedGeometry.numbered.maxX
      : textBoundaryCenter + (numberedRight - numberedLeft) / 2;
  if (
    equationNumberGeometry.minX <= measuredNumberedRight + 4 ||
    equationNumberGeometry.maxX > textBoundaryRight + 0.5
  ) {
    throw new Error(
      `Equation number is outside the expected right-side region: ${JSON.stringify({
        formulaRight: measuredNumberedRight,
        numberLeft: equationNumberGeometry.minX,
        numberRight: equationNumberGeometry.maxX,
        textBoundaryRight,
      })}`,
    );
  }

  const editRegressions = [];
  if (outputKind === "image") {
    const imageEditCases = [
      {
        formula: formulas[0],
        shapeIndex: 1,
        codeFormat: "raw",
        expectedLines: formulas[0].metadataLines,
        recovery: true,
      },
      {
        formula: formulas[5],
        shapeIndex: 6,
        codeFormat: "align",
        expectedLines: formulas[5].metadataLines,
        updatedLines: ["1 = 22 + 333 + q", "44444 = 55 + r"],
      },
      {
        formula: formulas[6],
        shapeIndex: 7,
        codeFormat: "align-star",
        expectedLines: formulas[6].metadataLines,
        updatedLines: ["666 = 777 + 8 + s", "999999 = 0 + t"],
      },
    ];

    for (const editCase of imageEditCases) {
      const sessionsBeforeEdit = currentSessionIds();
      if (editCase.recovery) rmSync(imageEditStatusPath, { force: true });
      runAppleScript([
        'tell application "Microsoft Word"',
        `set documentObject to document ${JSON.stringify(testDocumentName)}`,
        "activate object documentObject",
        "activate",
        `set formulaShape to inline shape ${editCase.shapeIndex} of documentObject`,
        "select text object of formulaShape",
        editCase.recovery
          ? 'run VB macro macro name "VisualTeX_RunSelectedImageEditRecoveryRegression"'
          : 'run VB macro macro name "VisualTeX_DoubleClickEditSelected"',
        "end tell",
      ], 60_000);
      const editSessionId = await waitForFormulaEditSession(
        sessionsBeforeEdit,
        editCase.formula.formulaId,
      );
      editSessionDirectories.push(join(sessionsRoot, editSessionId));
      const editSession = validateFormulaEditSession(
        editSessionId,
        editCase.formula,
        editCase.codeFormat,
        editCase.expectedLines,
      );

      let restoredReference;
      if (editCase.recovery) {
        const imageEditStatus = existsSync(imageEditStatusPath)
          ? readFileSync(imageEditStatusPath, "utf8").trim()
          : "missing-status";
        if (!imageEditStatus.startsWith("ok|")) {
          throw new Error(
            `Word image edit recovery regression failed: ${imageEditStatus}`,
          );
        }
        restoredReference = imageEditStatus.slice(3);
        const expectedReference =
          `visualtex:formula-ref:v1:${editCase.formula.formulaId}:` +
          `${editCase.formula.displayMode}:${editCase.formula.numbered ? "1" : "0"}`;
        if (restoredReference !== expectedReference) {
          throw new Error(
            `Word did not restore the image formula Title before editing: ${JSON.stringify({
              restoredReference,
              expectedReference,
            })}`,
          );
        }
      }

      const regression = {
        kind: editCase.recovery
          ? "image-metadata-title-recovery"
          : "image-batch-edit-session",
        sessionId: editSessionId,
        formulaId: editCase.formula.formulaId,
        codeFormat: editSession.normalized.codeFormat,
        lines: editSession.normalized.lines.map((line) => line.latex),
        ...(restoredReference ? { restoredReference } : {}),
      };

      if (editCase.updatedLines) {
        const replacement = commitEditedImageFormula(
          testDocumentName,
          editSessionId,
          editCase.formula,
          editSession,
          editCase.updatedLines,
        );
        const sessionsBeforeReplacementEdit = currentSessionIds();
        runAppleScript([
          'tell application "Microsoft Word"',
          `set documentObject to document ${JSON.stringify(testDocumentName)}`,
          "activate object documentObject",
          "activate",
          `set formulaShape to inline shape ${editCase.shapeIndex} of documentObject`,
          "select text object of formulaShape",
          'run VB macro macro name "VisualTeX_DoubleClickEditSelected"',
          "end tell",
        ], 60_000);
        const replacementEditSessionId = await waitForFormulaEditSession(
          sessionsBeforeReplacementEdit,
          editCase.formula.formulaId,
        );
        editSessionDirectories.push(
          join(sessionsRoot, replacementEditSessionId),
        );
        const replacementEditSession = validateFormulaEditSession(
          replacementEditSessionId,
          editCase.formula,
          editCase.codeFormat,
          editCase.updatedLines,
        );
        Object.assign(regression, {
          kind: "image-align-edit-replacement",
          replacementSessionId: replacementEditSessionId,
          replacementLatex: replacement.canonicalLatex,
          replacementLines: replacementEditSession.normalized.lines.map(
            (line) => line.latex,
          ),
          svgRelationshipPositions: replacement.svgAlignment.positions,
          svgRelationshipSpread: replacement.svgAlignment.spread,
        });
      }
      editRegressions.push(regression);
    }
    formulaRegressionReport = runFormulaRegressionReport(
      testDocumentName,
      formulas,
    );
  } else {
    const nativeEditCases = [
      {
        formula: formulas[5],
        codeFormat: "align",
        lines: ["1 = 22 + 333", "44444 = 55"],
      },
      {
        formula: formulas[6],
        codeFormat: "align-star",
        lines: ["666 = 777 + 8", "999999 = 0"],
      },
    ];
    for (const editCase of nativeEditCases) {
      const sessionsBeforeEdit = currentSessionIds();
      runAppleScript([
        'tell application "Microsoft Word"',
        `set documentObject to document ${JSON.stringify(testDocumentName)}`,
        "activate object documentObject",
        "activate",
        `select text object of bookmark ${JSON.stringify(nativeBookmark(editCase.formula.formulaId))} of documentObject`,
        'run VB macro macro name "VisualTeX_DoubleClickEditSelected"',
        "end tell",
      ], 60_000);
      const editSessionId = await waitForFormulaEditSession(
        sessionsBeforeEdit,
        editCase.formula.formulaId,
      );
      editSessionDirectories.push(join(sessionsRoot, editSessionId));
      const editSession = validateFormulaEditSession(
        editSessionId,
        editCase.formula,
        editCase.codeFormat,
        editCase.lines,
      );
      editRegressions.push({
        kind: "omml-multiline-edit",
        sessionId: editSessionId,
        formulaId: editCase.formula.formulaId,
        codeFormat: editSession.normalized.codeFormat,
        lines: editSession.normalized.lines.map((line) => line.latex),
        displayMode: editSession.request.displayMode,
        numbered: editSession.request.numbered,
        fontSizePt: editSession.request.fontSizePt,
      });
    }
  }

  const postEditFormulaContainerReport = inspectWordFormulaContainers(
    testDocumentName,
    formulas,
    "after-edit",
  );
  let reopenedFormulaContainerReport;
  if (physicalDoubleClick && outputKind === "image") {
    // The metadata-routing physical test targets the exact freshly imported
    // batch image. Keep it independent from Word's unrelated SaveAs wrapper,
    // which can return -128 after a successful save on some Mac builds.
    reopenedFormulaContainerReport = {
      ...postEditFormulaContainerReport,
      stage: "save-reopen-skipped-for-physical-image-routing",
    };
  } else {
    testDocumentName = saveAndReopenWordDocument(testDocumentName);
    reopenedFormulaContainerReport = inspectWordFormulaContainers(
      testDocumentName,
      formulas,
      "after-save-reopen",
    );
    formulaRegressionReport = runFormulaRegressionReport(
      testDocumentName,
      formulas,
    );
  }

  if (physicalDoubleClick) {
    const physicalFormulaIndex = {
      "image-inline": 0,
      "image-block": 1,
      "image-numbered": 2,
      "image-align": 5,
      "image-align-star": 6,
      "omml-inline": 0,
      "omml-block": 1,
      "omml-align": 5,
      "omml-align-star": 6,
    }[physicalTarget];
    const physicalFormula = formulas[physicalFormulaIndex];
    await startVisualTeXForPhysicalRegression();
    const sessionsBeforePhysicalEdit = currentSessionIds();
    const physicalSelection = runAppleScript([
      'tell application "Microsoft Word"',
      `set documentObject to document ${JSON.stringify(testDocumentName)}`,
      "activate object documentObject",
      "activate",
      'run VB macro macro name "VisualTeX_AssertWordHostSelfTest"',
      ...(outputKind === "image"
        ? [
            `set formulaShape to inline shape ${physicalFormulaIndex + 1} of documentObject`,
            "set formulaRange to text object of formulaShape",
            "select formulaRange",
            'return "image" & (ASCII character 31) & (start of content of formulaRange as text) & (ASCII character 31) & (end of content of formulaRange as text) & (ASCII character 31) & (width of formulaShape as text) & (ASCII character 31) & (height of formulaShape as text)',
          ]
        : [
            `set formulaRange to text object of bookmark ${JSON.stringify(nativeBookmark(physicalFormula.formulaId))} of documentObject`,
            "select formulaRange",
            'return "omml" & (ASCII character 31) & (start of content of formulaRange as text) & (ASCII character 31) & (end of content of formulaRange as text)',
          ]),
      "end tell",
    ]);
    console.log(
      `WORD_PHYSICAL_DOUBLE_CLICK_READY|${JSON.stringify({
        documentName: testDocumentName,
        target: physicalTarget,
        outputKind,
        formulaId: physicalFormula.formulaId,
        selection: physicalSelection.split("\x1f"),
      })}`,
    );
    const physicalClick = physicallyDoubleClickSelectedWordFormula(
      testDocumentName,
    );

    const physicalEditSessionId = await waitForFormulaEditSession(
      sessionsBeforePhysicalEdit,
      physicalFormula.formulaId,
      600_000,
    );
    editSessionDirectories.push(join(sessionsRoot, physicalEditSessionId));
    const physicalEditSession = validateFormulaEditSession(
      physicalEditSessionId,
      physicalFormula,
      physicalFormula.codeFormat,
      physicalFormula.metadataLines,
    );
    const editorReadiness = await waitForPhysicalEditorReadiness(
      physicalEditSessionId,
      physicalFormula.formulaId,
    );
    await assertSinglePhysicalEditSession(
      sessionsBeforePhysicalEdit,
      physicalEditSessionId,
      physicalFormula.formulaId,
    );
    let physicalApply = null;
    if (physicalApplyPerformance) {
      rmSync(wordPhysicalPerformanceStatusPath, { force: true });
      const replacementLatex = "101=303";
      const editInteraction = await replaceActiveVisualTeXFormula(
        replacementLatex,
      );
      const timing = await applyActiveVisualTeXFormula(
        physicalEditSessionId,
      );
      if (timing.clickToOfficeCompleteMs > 1_500) {
        throw new Error(
          `Word physical Apply missed the accepted 1500 ms limit: ${JSON.stringify(timing)}`,
        );
      }
      await sleep(100);
      const encodedMetadata = runAppleScript([
        'tell application "Microsoft Word"',
        `set documentObject to document ${JSON.stringify(testDocumentName)}`,
        `return alternative text of inline shape ${physicalFormulaIndex + 1} of documentObject`,
        "end tell",
      ]);
      const updatedMetadata = decodeFormulaMetadata(encodedMetadata);
      const updatedDocument = updatedMetadata
        ? normalizeFormulaEditorDocument(
            updatedMetadata.lines,
            updatedMetadata.codeFormat,
          )
        : null;
      if (
        updatedMetadata?.formulaId !== physicalFormula.formulaId ||
        updatedDocument?.lines[0]?.latex !== replacementLatex
      ) {
        throw new Error(
          `Word physical Apply did not persist the edited formula: ${JSON.stringify({ replacementLatex, updatedMetadata, updatedDocument })}`,
        );
      }
      physicalApply = {
        replacementLatex,
        editInteraction,
        timing,
        persistedLatex: updatedDocument.lines[0].latex,
      };
      writeFileSync(
        wordPhysicalPerformanceStatusPath,
        JSON.stringify(
          {
            status: "PASS",
            revision: "word-office-performance-20260801-r77",
            sessionId: physicalEditSessionId,
            formulaId: physicalFormula.formulaId,
            editorReadiness,
            physicalApply,
          },
          null,
          2,
        ),
        { mode: 0o600 },
      );
    }
    const pictureFormatUi = wordPictureFormatUiSnapshot();
    if (outputKind === "image" && pictureFormatUi.pictureFormatVisible) {
      throw new Error(
        `VisualTeX physical double-click also opened Word Picture Format UI: ${pictureFormatUi.raw}`,
      );
    }
    editRegressions.push({
      kind: `${physicalTarget}-physical-double-click`,
      target: physicalTarget,
      sessionId: physicalEditSessionId,
      formulaId: physicalFormula.formulaId,
      codeFormat: physicalEditSession.normalized.codeFormat,
      lines: physicalEditSession.normalized.lines.map((line) => line.latex),
      physicalClick,
      editorReadiness,
      physicalApply,
      pictureFormatUi,
    });
  }

  console.log(
    JSON.stringify(
      {
        sessionId,
        outputKind,
        formulas: formulas.map((formula, index) => ({
          formulaId: formula.formulaId,
          displayMode: formula.displayMode,
          numbered: formula.numbered,
          fontSizePt: formula.fontSizePt,
          codeFormat: formula.codeFormat,
          lines: formula.metadataLines,
          ...(formula.svgAlignment
            ? { svgAlignment: formula.svgAlignment }
            : {}),
          ...(outputKind === "omml"
            ? { bookmark: nativeBookmark(formula.formulaId) }
            : {
                shapeIndex: index + 1,
                wordObjectType: "InlineShape",
                layoutStructure:
                  formula.displayMode === "block"
                    ? formula.numbered
                      ? "numbered-display-paragraph"
                      : "dedicated-display-paragraph"
                    : "inline-text-flow",
              }),
        })),
        shapeCount,
        tableCount,
        documentImportProgress: {
          observed: observedDocumentImportProgress,
          final: finalDocumentImportProgress,
        },
        geometry: {
          pageWidth,
          leftMargin,
          rightMargin,
          textBoundaryLeft,
          textBoundaryRight,
          textBoundaryCenter,
          unnumberedDisplay: {
            left: displayLeft,
            right: displayRight,
            center: displayCenter,
            centerError: displayCenterError,
          },
          numberedDisplay: {
            left: numberedLeft,
            right: numberedRight,
            center: numberedCenter,
            centerError: numberedCenterError,
          },
          equationNumber: {
            left: numberLeft,
            right: numberRight,
          },
          displayToNumberedCenterError,
          renderedPdf: {
            pageWidth: renderedGeometry.pageWidth,
            pageHeight: renderedGeometry.pageHeight,
            unnumberedDisplay:
              outputKind === "omml"
                ? {
                    ...renderedGeometry.unnumbered,
                    centerError: renderedUnnumberedCenterError,
                  }
                : {
                    ...imageRasterGeometry.unnumbered,
                    wordShapeWidth: displayRight - displayLeft,
                    wordShapeHeight: displayHeight,
                    inkBounds: imageRasterGeometry.unnumberedInk,
                    centerError: renderedUnnumberedCenterError,
                  },
            numberedDisplay:
              outputKind === "omml"
                ? {
                    ...renderedGeometry.numbered,
                    centerError: renderedNumberedCenterError,
                  }
                : {
                    ...imageRasterGeometry.numbered,
                    wordShapeWidth: numberedRight - numberedLeft,
                    wordShapeHeight: numberedHeight,
                    inkBounds: imageRasterGeometry.numberedInk,
                    centerError: renderedNumberedCenterError,
                  },
            equationNumber: {
              ...equationNumberGeometry,
              verticalCenterError: equationNumberVerticalError,
              inkCenterDifference: equationNumberInkCenterDifference,
              rightBoundaryInset: textBoundaryRight - equationNumberGeometry.maxX,
            },
            formulaCenterDifference: renderedFormulaCenterDifference,
            centerTolerancePt,
            imageVisualCalibration,
            ommlAlignmentGeometry,
          },
        },
        documentText,
        formulaRegressionReport,
        formulaContainerReports: {
          afterImport: initialFormulaContainerReport,
          afterEdit: postEditFormulaContainerReport,
          afterSaveReopen: reopenedFormulaContainerReport,
        },
        editRegressions,
      },
      null,
      2,
    ),
  );
  console.log("Word document import integration passed");
  }
} catch (error) {
  if (
    error instanceof Error &&
    error.message.startsWith(diagnosticSuccessPrefix)
  ) {
    console.log(
      `Word document-import diagnostic passed ${error.message.slice(diagnosticSuccessPrefix.length)} items`,
    );
  } else {
    try {
      const wordState = runAppleScript([
        'tell application "Microsoft Word"',
        'if not (exists active document) then return "no-active-document"',
        "set documentObject to active document",
        "set unitSeparator to ASCII character 31",
        "set bookmarkNames to name of every bookmark of documentObject",
        "set documentText to content of text object of documentObject",
        "set paragraphCount to count paragraphs of documentObject",
        'return (name of documentObject as text) & unitSeparator & (paragraphCount as text) & unitSeparator & (bookmarkNames as text) & unitSeparator & documentText',
        "end tell",
      ], 15_000);
      console.error(`Word state after callback failure:\n${wordState}`);
    } catch (stateError) {
      console.error(
        `Unable to inspect Word after callback failure: ${stateError instanceof Error ? stateError.message : String(stateError)}`,
      );
    }
    if (createImagePhysicalRegression) {
      writeFileSync(
        finalBinaryPhysicalStatusPath,
        JSON.stringify(
          {
            status: "FAIL",
            revision: "word-office-performance-20260801-r77",
            error: error instanceof Error ? error.message : String(error),
          },
          null,
          2,
        ),
        { mode: 0o600 },
      );
    }
    if (sessionDirectory) {
    const stagePath = join(sessionDirectory, "document-import-stage.txt");
    if (existsSync(stagePath)) {
      console.error(`Last Word document-import stage:\n${readFileSync(stagePath, "utf8")}`);
    }
    const failurePath = join(sessionDirectory, "word-failure.log");
    if (existsSync(failurePath)) {
      console.error(`Word document-import failure:\n${readFileSync(failurePath, "utf8")}`);
    }
    }
    throw error;
  }
} finally {
  try {
    runAppleScript(['tell application "Microsoft Word" to quit saving no'], 20_000);
  } catch {
    // Continue with a hard process cleanup below.
  }
  spawnSync("/usr/bin/killall", ["Microsoft Word"], {
    encoding: "utf8",
    timeout: 10_000,
  });
  rmSync(pdfExportRequestPath, { force: true });
  rmSync(pdfExportStatusPath, { force: true });
  rmSync(imageEditStatusPath, { force: true });
  rmSync(formulaRegressionStatusPath, { force: true });
  rmSync(physicalScreenBoundsPath, { force: true });
  rmSync(coordinatePdfPath, { force: true });
  rmSync(installedWordAddinPath, { force: true });
  if (installedWordAddinBackedUp && existsSync(installedWordAddinBackupPath)) {
    copyFileSync(installedWordAddinBackupPath, installedWordAddinPath);
  }
  rmSync(installedWordAddinBackupPath, { force: true });
  if (sessionDirectory) rmSync(sessionDirectory, { recursive: true, force: true });
  for (const editSessionDirectory of editSessionDirectories) {
    rmSync(editSessionDirectory, { recursive: true, force: true });
  }
  rmSync(join(sessionsRoot, "word-active-session.txt"), { force: true });
  for (const nativeFile of nativeFiles) rmSync(nativeFile, { force: true });
}
