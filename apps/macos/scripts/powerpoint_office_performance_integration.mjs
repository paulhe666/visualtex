import { execFileSync, spawnSync } from "node:child_process";
import {
  existsSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { decodeFormulaMetadata } from "../src/office/shared/formulaMetadata.ts";
import { normalizeFormulaEditorDocument } from "../src/office/shared/formulaEditorDocument.ts";

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const runtimeRoot = join(
  homedir(),
  "Library/Application Scripts/com.microsoft.Powerpoint/VisualTeXRuntime",
);
const sessionsRoot = join(runtimeRoot, "OfficeSessions");
const installedAddinPath = join(
  homedir(),
  "Library/Group Containers/UBF8T346G9.Office/VisualTeX/OfficeAddins/VisualTeX.ppam",
);
const resultPath = join(
  homedir(),
  "Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch/powerpoint-office-performance.json",
);
const movedCopyResultPath = join(
  runtimeRoot,
  "Tests/powerpoint-moved-copy-regression-result.txt",
);
const hangTracePath = "/tmp/visualtex-powerpoint-hang.txt";
const sourceVisualTeXAppPath = join(
  repositoryRoot,
  "src-tauri/target/release/bundle/macos/VisualTeX.app",
);
const installedVisualTeXAppPath = "/Applications/VisualTeX.app";
const installedVisualTeXAppBackup = `/tmp/VisualTeX-PowerPointPerformance-${process.pid}.app`;
const hadInstalledVisualTeXApp = existsSync(installedVisualTeXAppPath);
let installedVisualTeXAppBackedUp = false;
const editorReadyFile = "editor-ready.json";
const editorPerformanceFile = "editor-performance.jsonl";
const sleep = (ms) => new Promise((resolvePromise) => setTimeout(resolvePromise, ms));

function run(program, args, timeout = 60_000) {
  return execFileSync(program, args, {
    encoding: "utf8",
    timeout,
    maxBuffer: 32 * 1024 * 1024,
  }).trim();
}

function bestEffort(program, args, timeout = 20_000) {
  try {
    return run(program, args, timeout);
  } catch {
    return "";
  }
}

function runAppleScript(lines, timeout = 60_000) {
  return run(
    "/usr/bin/osascript",
    lines.flatMap((line) => ["-e", line]),
    timeout,
  );
}

function currentSessionIds() {
  if (!existsSync(sessionsRoot)) return new Set();
  return new Set(
    readdirSync(sessionsRoot, { withFileTypes: true })
      .filter((entry) => entry.isDirectory())
      .map((entry) => entry.name),
  );
}

function requestForSession(sessionId) {
  const requestPath = join(sessionsRoot, sessionId, "request.json");
  if (!existsSync(requestPath)) return null;
  try {
    return JSON.parse(readFileSync(requestPath, "utf8"));
  } catch {
    return null;
  }
}

async function waitForNewSession(before, operation, formulaId, timeoutMs = 20_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    for (const sessionId of currentSessionIds()) {
      if (before.has(sessionId)) continue;
      const request = requestForSession(sessionId);
      if (
        request?.host === "powerpoint" &&
        request?.mode === operation &&
        (!formulaId || request.formulaId === formulaId)
      ) {
        return { sessionId, request };
      }
    }
    await sleep(25);
  }
  throw new Error(`PowerPoint did not create a ${operation} VisualTeX Session`);
}

async function waitForEditorReady(sessionId, timeoutMs = 10_000) {
  const readyPath = join(sessionsRoot, sessionId, editorReadyFile);
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (existsSync(readyPath)) {
      const marker = JSON.parse(readFileSync(readyPath, "utf8"));
      if (marker.sessionId === sessionId) return marker;
    }
    await sleep(20);
  }
  throw new Error(`PowerPoint editor did not become ready for ${sessionId}`);
}

function editorPerformanceRecords(sessionId) {
  const performancePath = join(
    sessionsRoot,
    sessionId,
    editorPerformanceFile,
  );
  if (!existsSync(performancePath)) return [];
  return readFileSync(performancePath, "utf8")
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => JSON.parse(line));
}

function visualTeXEditorWindowBounds() {
  const raw = runAppleScript([
    'tell application "System Events"',
    'tell process "visualtex"',
    'if (count of windows) is 0 then error "VisualTeX editor window is missing"',
    'set editorWindow to window 1',
    'set windowPosition to position of editorWindow',
    'set windowSize to size of editorWindow',
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

function focusedVisualTeXElement() {
  return runAppleScript([
    'tell application "System Events"',
    'tell process "visualtex"',
    'set focusedElement to value of attribute "AXFocusedUIElement"',
    'set roleValue to ""',
    'set subroleValue to ""',
    'set nameValue to ""',
    'set descriptionValue to ""',
    'try',
    'set roleValue to role of focusedElement as text',
    'end try',
    'try',
    'set subroleValue to subrole of focusedElement as text',
    'end try',
    'try',
    'set nameValue to name of focusedElement as text',
    'end try',
    'try',
    'set descriptionValue to description of focusedElement as text',
    'end try',
    'return roleValue & "|" & subroleValue & "|" & nameValue & "|" & descriptionValue',
    "end tell",
    "end tell",
  ]);
}

function visualTeXAccessibilitySummary(limit = 220) {
  return runAppleScript([
    'tell application "System Events"',
    'tell process "visualtex"',
    'set editorWindow to missing value',
    'repeat with candidateWindow in windows',
    'try',
    'if (name of candidateWindow as text) is "VisualTeX Office Formula" then',
    'set editorWindow to candidateWindow',
    'exit repeat',
    'end if',
    'end try',
    'end repeat',
    'if editorWindow is missing value then return "NO_EDITOR_WINDOW"',
    'set editorItems to entire contents of editorWindow',
    'set summaryText to ""',
    'set itemIndex to 0',
    'repeat with candidateElement in editorItems',
    'set itemIndex to itemIndex + 1',
    `if itemIndex > ${Number(limit)} then exit repeat`,
    'set roleValue to ""',
    'set subroleValue to ""',
    'set nameValue to ""',
    'set descriptionValue to ""',
    'set valueText to ""',
    'try\nset roleValue to role of candidateElement as text\nend try',
    'try\nset subroleValue to subrole of candidateElement as text\nend try',
    'try\nset nameValue to name of candidateElement as text\nend try',
    'try\nset descriptionValue to description of candidateElement as text\nend try',
    'try\nset valueText to value of candidateElement as text\nend try',
    'if (length of valueText) > 100 then set valueText to text 1 thru 100 of valueText',
    'set summaryText to summaryText & itemIndex & "|" & roleValue & "|" & subroleValue & "|" & nameValue & "|" & descriptionValue & "|" & valueText & linefeed',
    'end repeat',
    'return summaryText',
    'end tell',
    'end tell',
  ]);
}

async function replaceActiveFormula(latex) {
  const clipboard = spawnSync("/usr/bin/pbcopy", [], {
    input: latex,
    encoding: "utf8",
    timeout: 5_000,
  });
  if (clipboard.status !== 0) {
    throw new Error(clipboard.stderr || "Unable to prepare the formula clipboard");
  }
  console.log(`POWERPOINT_FOCUS_BEFORE|${focusedVisualTeXElement()}`);
  const result = runAppleScript([
    'tell application "System Events"',
    'tell process "visualtex"',
    "set visible to true",
    "set frontmost to true",
    "delay 0.1",
    "set editorWindow to missing value",
    "repeat with candidateWindow in windows",
    "try",
    'if (name of candidateWindow as text) is "VisualTeX Office Formula" then',
    "set editorWindow to candidateWindow",
    "exit repeat",
    "end if",
    "end try",
    "end repeat",
    'if editorWindow is missing value then error "The active VisualTeX Office editor was not found."',
    "set editorItems to entire contents of editorWindow",
    "set formulaField to missing value",
    "repeat with candidateElement in editorItems",
    "try",
    'if (role of candidateElement as text) is "AXTextField" then',
    "set formulaField to candidateElement",
    "exit repeat",
    "end if",
    "end try",
    "end repeat",
    'if formulaField is missing value then error "The VisualTeX formula input field was not found."',
    "set fieldPosition to position of formulaField",
    "set fieldSize to size of formulaField",
    "set clickX to (item 1 of fieldPosition) + ((item 1 of fieldSize) / 2)",
    "set clickY to (item 2 of fieldPosition) + ((item 2 of fieldSize) / 2)",
    "click at {clickX, clickY}",
    "delay 0.15",
    'keystroke "a" using {command down}',
    "delay 0.05",
    'keystroke "v" using {command down}',
    "delay 0.2",
    'return "REPLACED"',
    "end tell",
    "end tell",
  ]);
  if (result.trim() !== "REPLACED") {
    throw new Error(`Unexpected VisualTeX formula replacement result: ${result}`);
  }
  await sleep(100);
  console.log(`POWERPOINT_FOCUS_AFTER|${focusedVisualTeXElement()}`);
}

function selectFormulaShape(formulaId, slideIndex) {
  const shapeName = `VisualTeX_${formulaId}`;
  runAppleScript([
    'tell application "Microsoft PowerPoint"',
    `set currentSlide to slide ${Number(slideIndex)} of active presentation`,
    "go to slide (view of active window) number (slide index of currentSlide)",
    `if not (exists shape ${JSON.stringify(shapeName)} of currentSlide) then error "The VisualTeX formula shape is missing"`,
    `select shape ${JSON.stringify(shapeName)} of currentSlide`,
    "end tell",
  ]);
}

function processIds(processName) {
  return bestEffort("/usr/bin/pgrep", ["-x", processName])
    .split(/\s+/)
    .map((value) => Number.parseInt(value, 10))
    .filter((value) => Number.isInteger(value) && value > 0);
}

function forceQuitPowerPoint() {
  bestEffort("/usr/bin/osascript", [
    "-e",
    'tell application "Microsoft PowerPoint" to quit saving no',
  ], 10_000);
  bestEffort("/usr/bin/killall", ["Microsoft PowerPoint"], 10_000);
  for (const pid of processIds("Microsoft PowerPoint")) {
    bestEffort("/bin/kill", ["-9", String(pid)], 5_000);
  }
}

function sampleApplyHang(sessionId) {
  const sections = [`sessionId=${sessionId}`, `capturedEpochMs=${Date.now()}`];
  for (const processName of ["visualtex", "Microsoft PowerPoint", "osascript"]) {
    for (const pid of processIds(processName)) {
      let sample = "";
      try {
        sample = run("/usr/bin/sample", [String(pid), "1", "1"], 15_000);
      } catch (error) {
        sample = error instanceof Error ? error.message : String(error);
      }
      sections.push(`\n===== ${processName} pid=${pid} =====\n${sample}`);
    }
  }
  const combined = `${sections.join("\n")}\n`;
  writeFileSync(hangTracePath, combined, { mode: 0o600 });
  const symbols = combined
    .split(/\r?\n/)
    .map((line) =>
      line.match(/visualtex_lib::office::macos_offline::([^\s]+)/)?.[1],
    )
    .filter(Boolean);
  writeFileSync(
    "/tmp/visualtex-powerpoint-hang-symbols.txt",
    `${[...new Set(symbols)].join("\n")}\n`,
    { mode: 0o600 },
  );
}

async function applyActiveFormula(sessionId, timeoutMs = 30_000) {
  // Resolve/focus the editor before the measured Apply press. The first
  // Accessibility tree walk on a fresh macOS host can take seconds and is not
  // part of the user's actual Cmd+Enter-to-commit latency.
  const editorReady = runAppleScript([
    'tell application "System Events"',
    'tell process "visualtex"',
    "set frontmost to true",
    "delay 0.05",
    "set editorWindow to missing value",
    "repeat with candidateWindow in windows",
    "try",
    'if (name of candidateWindow as text) is "VisualTeX Office Formula" then',
    "set editorWindow to candidateWindow",
    "exit repeat",
    "end if",
    "end try",
    "end repeat",
    'if editorWindow is missing value then error "The active VisualTeX Office editor was not found."',
    'return "READY"',
    "end tell",
    "end tell",
  ]);
  if (editorReady.trim() !== "READY") {
    throw new Error(`Unexpected VisualTeX editor readiness result: ${editorReady}`);
  }

  const startedEpochMs = Date.now();
  const vbaStagePath = join(sessionsRoot, sessionId, "vba-commit-stage.txt");
  const vbaStages = [];
  let lastVbaStage = "";
  const clickResult = runAppleScript([
    'tell application "System Events"',
    'tell process "visualtex"',
    'keystroke return using {command down}',
    'return "PRESSED"',
    "end tell",
    "end tell",
  ]);
  if (clickResult.trim() !== "PRESSED") {
    throw new Error(`Unexpected VisualTeX Apply press result: ${clickResult}`);
  }
  const deadline = Date.now() + timeoutMs;
  let hangSampled = false;
  while (Date.now() < deadline) {
    if (existsSync(vbaStagePath)) {
      try {
        const currentVbaStage = readFileSync(vbaStagePath, "utf8")
          .split(/\r?\n/)[0]
          .trim();
        if (currentVbaStage && currentVbaStage !== lastVbaStage) {
          lastVbaStage = currentVbaStage;
          vbaStages.push({
            stage: currentVbaStage,
            epochMs: Date.now(),
            elapsedFromClickMs: Date.now() - startedEpochMs,
          });
        }
      } catch {
        // Retry while the atomic stage marker is being replaced.
      }
    }
    const record = editorPerformanceRecords(sessionId).find(
      (candidate) => candidate.stage === "apply-backend-complete",
    );
    if (record) {
      const completedEpochMs = Number(record.epochMs);
      return {
        startedEpochMs,
        completedEpochMs,
        clickToOfficeCompleteMs: completedEpochMs - startedEpochMs,
        backendElapsedMs: Number(record.elapsedMs),
        records: editorPerformanceRecords(sessionId).filter((candidate) =>
          String(candidate.stage).startsWith("apply-"),
        ),
        vbaStages,
      };
    }
    if (!hangSampled && Date.now() - startedEpochMs >= 5_000) {
      hangSampled = true;
      sampleApplyHang(sessionId);
    }
    await sleep(20);
  }
  throw new Error(`PowerPoint Apply did not complete for ${sessionId}`);
}

function enableMacrosIfPrompted() {
  bestEffort("/usr/bin/osascript", [
    "-e",
    'tell application "System Events"',
    "-e",
    'if exists process "Microsoft PowerPoint" then',
    "-e",
    'tell process "Microsoft PowerPoint"',
    "-e",
    "repeat with candidateWindow in windows",
    "-e",
    'repeat with buttonName in {"启用宏", "Enable Macros"}',
    "-e",
    "try",
    "-e",
    'if exists button (buttonName as text) of candidateWindow then click button (buttonName as text) of candidateWindow',
    "-e",
    "end try",
    "-e",
    "end repeat",
    "-e",
    "end repeat",
    "-e",
    "end tell",
    "-e",
    "end if",
    "-e",
    "end tell",
  ]);
}

async function waitForPowerPointUi(timeoutMs = 30_000) {
  const deadline = Date.now() + timeoutMs;
  let lastError = "";
  while (Date.now() < deadline) {
    try {
      const state = runAppleScript([
        'tell application "Microsoft PowerPoint" to activate',
        'tell application "System Events"',
        'tell process "Microsoft PowerPoint"',
        "set visible to true",
        "set frontmost to true",
        'if (count of menu bars) = 0 then error "PowerPoint UI is not ready"',
        'return "READY"',
        "end tell",
        "end tell",
      ], 5_000);
      if (state === "READY") return;
    } catch (error) {
      lastError = error instanceof Error ? error.message : String(error);
    }
    await sleep(500);
  }
  throw new Error(`PowerPoint did not become UI-ready: ${lastError}`);
}

function createTestPresentation() {
  return runAppleScript([
    'tell application "Microsoft PowerPoint"',
    "activate",
    "set testPresentation to make new presentation",
    "if (count of slides of testPresentation) is 0 then make new slide at end of testPresentation",
    "return name of testPresentation as text",
    "end tell",
  ]);
}

function runPowerPointMacro(name) {
  return runAppleScript([
    'tell application "Microsoft PowerPoint"',
    `run VB macro macro name ${JSON.stringify(name)} list of parameters {}`,
    "end tell",
  ]);
}

function formulaDocumentFromEditRequest(request) {
  const metadata = decodeFormulaMetadata(request?.encodedMetadata ?? "");
  const document = metadata
    ? normalizeFormulaEditorDocument(metadata.lines, metadata.codeFormat)
    : null;
  if (!metadata || !document) {
    throw new Error("PowerPoint edit request did not contain valid formula metadata");
  }
  return { metadata, document };
}

async function ensureAddinLoaded() {
  try {
    runPowerPointMacro("Auto_Open");
    return;
  } catch {
    if (!existsSync(installedAddinPath)) {
      throw new Error(`Installed PowerPoint add-in is missing: ${installedAddinPath}`);
    }
    run("/usr/bin/open", ["-b", "com.microsoft.Powerpoint", installedAddinPath]);
    await sleep(2_000);
    enableMacrosIfPrompted();
    await sleep(500);
    runPowerPointMacro("Auto_Open");
  }
}

function stageLatestVisualTeXApplication() {
  bestEffort("/usr/bin/killall", ["visualtex"]);
  rmSync(installedVisualTeXAppBackup, { recursive: true, force: true });
  if (hadInstalledVisualTeXApp) {
    run("/usr/bin/ditto", [installedVisualTeXAppPath, installedVisualTeXAppBackup], 120_000);
    installedVisualTeXAppBackedUp = true;
  }
  rmSync(installedVisualTeXAppPath, { recursive: true, force: true });
  run("/usr/bin/ditto", [sourceVisualTeXAppPath, installedVisualTeXAppPath], 120_000);
  // LaunchAgent may relaunch the resident while /Applications/VisualTeX.app is
  // being replaced. Kill once more after the staged app is complete so the
  // acceptance run cannot inherit a process that started from mixed bytes.
  bestEffort("/usr/bin/killall", ["visualtex"]);
}

function restoreInstalledVisualTeXApplication() {
  bestEffort("/usr/bin/killall", ["visualtex"]);
  rmSync(installedVisualTeXAppPath, { recursive: true, force: true });
  if (installedVisualTeXAppBackedUp && existsSync(installedVisualTeXAppBackup)) {
    bestEffort("/usr/bin/ditto", [installedVisualTeXAppBackup, installedVisualTeXAppPath], 120_000);
  }
  rmSync(installedVisualTeXAppBackup, { recursive: true, force: true });
  if (hadInstalledVisualTeXApp && existsSync(installedVisualTeXAppPath)) {
    bestEffort("/usr/bin/open", ["-gj", installedVisualTeXAppPath, "--args", "--office-background"], 20_000);
  }
}

async function main() {
  rmSync(resultPath, { force: true });
  forceQuitPowerPoint();
  await sleep(750);
  stageLatestVisualTeXApplication();
  await sleep(750);
  run("/usr/bin/open", ["-gj", installedVisualTeXAppPath, "--args", "--office-background"]);
  await sleep(8_000);
  run("/usr/bin/open", ["-b", "com.microsoft.Powerpoint"]);
  await waitForPowerPointUi();
  const powerPointPid = processIds("Microsoft PowerPoint")[0];
  if (!Number.isInteger(powerPointPid)) {
    throw new Error("PowerPoint fresh-host PID was not available after launch");
  }
  await ensureAddinLoaded();
  // A freshly launched PowerPoint can report the add-in as loaded slightly
  // before its VBA/add-in event loop is ready to accept the first external
  // macro call. Give the host one short, bounded settle window so clean-start
  // acceptance reflects normal user timing instead of racing Office startup.
  await sleep(3_000);
  const presentationName = createTestPresentation();

  const beforeCreate = currentSessionIds();
  runPowerPointMacro("VisualTeX_NewFormula");
  const created = await waitForNewSession(beforeCreate, "create");
  const createReady = await waitForEditorReady(created.sessionId);
  const createLatex = String.raw`E=mc^2`;
  await replaceActiveFormula(createLatex);
  const createApply = await applyActiveFormula(created.sessionId);
  if (createApply.clickToOfficeCompleteMs > 1_500) {
    throw new Error(
      `PowerPoint create Apply exceeded 1500 ms: ${JSON.stringify(createApply)}`,
    );
  }
  await sleep(150);
  selectFormulaShape(
    created.request.formulaId,
    created.request.powerPoint.slideIndex,
  );

  rmSync(movedCopyResultPath, { force: true });
  runPowerPointMacro("VisualTeX_PreparePowerPointMovedCopyRegression");
  if (!existsSync(movedCopyResultPath)) {
    throw new Error("PowerPoint moved-copy regression did not write its prepared state");
  }
  const movedCopyPrepared = readFileSync(movedCopyResultPath, "utf8");
  if (!movedCopyPrepared.startsWith("PREPARED")) {
    throw new Error(`PowerPoint moved-copy preparation failed: ${movedCopyPrepared}`);
  }
  const copiedIdMatch = movedCopyPrepared.match(/^copiedId=(\d+)$/m);
  const sourceLeftMatch = movedCopyPrepared.match(/^sourceLeft=([-+0-9.]+)$/m);
  const sourceTopMatch = movedCopyPrepared.match(/^sourceTop=([-+0-9.]+)$/m);
  const copiedLeftMatch = movedCopyPrepared.match(/^copiedLeft=([-+0-9.]+)$/m);
  const copiedTopMatch = movedCopyPrepared.match(/^copiedTop=([-+0-9.]+)$/m);
  if (!copiedIdMatch || !sourceLeftMatch || !sourceTopMatch || !copiedLeftMatch || !copiedTopMatch) {
    throw new Error(`PowerPoint moved-copy state is incomplete: ${movedCopyPrepared}`);
  }
  const sourceLeft = Number(sourceLeftMatch[1]);
  const sourceTop = Number(sourceTopMatch[1]);
  const copiedLeft = Number(copiedLeftMatch[1]);
  const copiedTop = Number(copiedTopMatch[1]);
  const beforeMovedCopyEdit = currentSessionIds();
  runPowerPointMacro("VisualTeX_DoubleClickEditSelected");
  const movedCopyEdit = await waitForNewSession(beforeMovedCopyEdit, "edit");
  const movedCopyReady = await waitForEditorReady(movedCopyEdit.sessionId);
  if (movedCopyEdit.request.forkCopiedFormula !== true) {
    throw new Error("The moved PowerPoint copy was not classified as an independent copied formula");
  }
  if (Number(movedCopyEdit.request.powerPoint?.shapeId) !== Number(copiedIdMatch[1])) {
    throw new Error(
      `PowerPoint moved-copy edit targeted Shape.ID ${movedCopyEdit.request.powerPoint?.shapeId}, expected ${copiedIdMatch[1]}`,
    );
  }
  if (movedCopyEdit.request.formulaId === created.request.formulaId) {
    throw new Error("The moved PowerPoint copy reused the source formulaId");
  }
  const movedCopyLatex = String.raw`E=mc^2+1`;
  await replaceActiveFormula(movedCopyLatex);
  const movedCopyApply = await applyActiveFormula(movedCopyEdit.sessionId);
  if (movedCopyApply.clickToOfficeCompleteMs > 1_500) {
    throw new Error(
      `PowerPoint moved-copy Apply exceeded 1500 ms: ${JSON.stringify(movedCopyApply)}`,
    );
  }
  const movedCopyVerification = movedCopyPrepared.trim();

  selectFormulaShape(
    created.request.formulaId,
    created.request.powerPoint.slideIndex,
  );
  const beforeEdit = currentSessionIds();
  const editInvokedEpochMs = Date.now();
  runPowerPointMacro("VisualTeX_DoubleClickEditSelected");
  const edited = await waitForNewSession(
    beforeEdit,
    "edit",
    created.request.formulaId,
  );
  const editReady = await waitForEditorReady(edited.sessionId);
  if (
    Math.abs(Number(edited.request.powerPoint?.left) - sourceLeft) > 0.1 ||
    Math.abs(Number(edited.request.powerPoint?.top) - sourceTop) > 0.1
  ) {
    throw new Error(
      `Editing the moved copy changed the source formula position: expected ${sourceLeft},${sourceTop}, got ${edited.request.powerPoint?.left},${edited.request.powerPoint?.top}`,
    );
  }
  const editInvokeToReadyMs = Number(editReady.epochMs) - editInvokedEpochMs;
  const editRequestToReadyMs =
    Number(editReady.epochMs) - Number(editReady.urlReceivedEpochMs);
  if (!Number.isFinite(editRequestToReadyMs) || editRequestToReadyMs > 1_000) {
    throw new Error(
      `PowerPoint editor request-to-ready exceeded 1000 ms: ${editRequestToReadyMs}`,
    );
  }
  const createdState = formulaDocumentFromEditRequest(edited.request);
  if (
    createdState.metadata.formulaId !== created.request.formulaId ||
    createdState.document.lines[0]?.latex !== createLatex
  ) {
    throw new Error("PowerPoint did not persist the created formula metadata");
  }

  const editLatex = String.raw`E^2=p^2c^2+m^2c^4`;
  await replaceActiveFormula(editLatex);
  const editApply = await applyActiveFormula(edited.sessionId);
  if (editApply.clickToOfficeCompleteMs > 1_500) {
    throw new Error(
      `PowerPoint edit Apply exceeded 1500 ms: ${JSON.stringify(editApply)}`,
    );
  }
  await sleep(150);
  selectFormulaShape(
    created.request.formulaId,
    created.request.powerPoint.slideIndex,
  );
  const beforeVerify = currentSessionIds();
  runPowerPointMacro("VisualTeX_DoubleClickEditSelected");
  const verified = await waitForNewSession(
    beforeVerify,
    "edit",
    created.request.formulaId,
  );
  const verifyReady = await waitForEditorReady(verified.sessionId);
  const editedState = formulaDocumentFromEditRequest(verified.request);
  if (
    editedState.metadata.formulaId !== created.request.formulaId ||
    editedState.document.lines[0]?.latex !== editLatex
  ) {
    throw new Error("PowerPoint did not persist the edited formula metadata");
  }

  selectFormulaShape(
    movedCopyEdit.request.formulaId,
    movedCopyEdit.request.powerPoint.slideIndex,
  );
  const beforeMovedCopyVerify = currentSessionIds();
  runPowerPointMacro("VisualTeX_DoubleClickEditSelected");
  const movedCopyVerified = await waitForNewSession(
    beforeMovedCopyVerify,
    "edit",
    movedCopyEdit.request.formulaId,
  );
  const movedCopyVerifyReady = await waitForEditorReady(movedCopyVerified.sessionId);
  const movedCopyCenterBefore = {
    x:
      Number(movedCopyEdit.request.powerPoint?.left) +
      Number(movedCopyEdit.request.powerPoint?.width) / 2,
    y:
      Number(movedCopyEdit.request.powerPoint?.top) +
      Number(movedCopyEdit.request.powerPoint?.height) / 2,
  };
  const movedCopyCenterAfter = {
    x:
      Number(movedCopyVerified.request.powerPoint?.left) +
      Number(movedCopyVerified.request.powerPoint?.width) / 2,
    y:
      Number(movedCopyVerified.request.powerPoint?.top) +
      Number(movedCopyVerified.request.powerPoint?.height) / 2,
  };
  if (
    Math.abs(movedCopyCenterAfter.x - movedCopyCenterBefore.x) > 0.1 ||
    Math.abs(movedCopyCenterAfter.y - movedCopyCenterBefore.y) > 0.1
  ) {
    throw new Error(
      `The edited PowerPoint copy did not preserve its moved center: expected ${movedCopyCenterBefore.x},${movedCopyCenterBefore.y}, got ${movedCopyCenterAfter.x},${movedCopyCenterAfter.y}`,
    );
  }
  const movedCopyState = formulaDocumentFromEditRequest(movedCopyVerified.request);
  if (
    movedCopyState.metadata.formulaId !== movedCopyEdit.request.formulaId ||
    movedCopyState.document.lines[0]?.latex !== movedCopyLatex
  ) {
    throw new Error("PowerPoint did not persist the independently edited moved copy");
  }

  const result = {
    status: "PASS",
    revision: "powerpoint-office-performance-20260801-r4",
    powerPointPid,
    presentationName,
    formulaId: created.request.formulaId,
    create: {
      sessionId: created.sessionId,
      requestToReadyMs: Number(createReady.requestToReadyMs),
      showFocusMs: Number(createReady.showFocusMs),
      apply: createApply,
      persistedLatex: createdState.document.lines[0].latex,
    },
    movedCopy: {
      sessionId: movedCopyEdit.sessionId,
      requestToReadyMs:
        Number(movedCopyReady.epochMs) - Number(movedCopyReady.urlReceivedEpochMs),
      showFocusMs: Number(movedCopyReady.showFocusMs),
      apply: movedCopyApply,
      verification: movedCopyVerification,
      sourcePosition: { left: sourceLeft, top: sourceTop },
      movedPosition: { left: copiedLeft, top: copiedTop },
      verificationSessionId: movedCopyVerified.sessionId,
      verificationRequestToReadyMs:
        Number(movedCopyVerifyReady.epochMs) - Number(movedCopyVerifyReady.urlReceivedEpochMs),
    },
    edit: {
      sessionId: edited.sessionId,
      requestToReadyMs: editRequestToReadyMs,
      invokeToReadyMs: editInvokeToReadyMs,
      showFocusMs: Number(editReady.showFocusMs),
      apply: editApply,
      persistedLatex: editedState.document.lines[0].latex,
      verificationSessionId: verified.sessionId,
      verificationRequestToReadyMs: Number(verifyReady.requestToReadyMs),
    },
  };
  writeFileSync(resultPath, `${JSON.stringify(result, null, 2)}\n`, {
    mode: 0o600,
  });
  console.log(JSON.stringify(result, null, 2));
}

try {
  await main();
} finally {
  forceQuitPowerPoint();
  restoreInstalledVisualTeXApplication();
}
