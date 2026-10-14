import { spawn, spawnSync } from "node:child_process";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { homedir, tmpdir } from "node:os";
import { basename, join, resolve } from "node:path";
import {
  createFormulaMetadata,
  encodeFormulaMetadata,
} from "../src/office/shared/formulaMetadata.ts";

function option(name: string) {
  const index = process.argv.indexOf(name);
  if (index < 0) return "";
  const value = process.argv[index + 1];
  if (!value || value.startsWith("--")) {
    throw new Error(`${name} requires a value`);
  }
  return value;
}

const repositoryRoot = resolve(new URL("..", import.meta.url).pathname);
const wordAddinPath = resolve(option("--word-addin"));
const sourceDocumentPath = option("--document");
const snapshotPrefix = `word-global-format-${Date.now()}`;
const useRibbon = process.argv.includes("--use-ribbon");
if (!wordAddinPath || !existsSync(wordAddinPath)) {
  throw new Error(`Missing compiled Word add-in: ${wordAddinPath}`);
}

const home = homedir();
const runtimeRoot = join(
  home,
  "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime",
);
const sessionsRoot = join(runtimeRoot, "OfficeSessions");
const testsRoot = join(runtimeRoot, "Tests");
const startupRoot = join(
  home,
  "Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word",
);
const installedAddinPath = join(startupRoot, "VisualTeX.dotm");
const installedBackupPath = join(
  tmpdir(),
  `visualtex-word-global-format-startup-${process.pid}.dotm`,
);
const wordApplicationScriptsRoot = join(
  home,
  "Library/Application Scripts/com.microsoft.Word",
);
const installedWordScriptPath = join(
  wordApplicationScriptsRoot,
  "VisualTeXWord.scpt",
);
const installedWordScriptBackupPath = join(
  tmpdir(),
  `visualtex-word-global-format-script-${process.pid}.scpt`,
);
const wordScriptSourcePath = join(
  repositoryRoot,
  "office/macos-offline/word/VisualTeXWord.scpt",
);
const workspaceVisualTeXApp = join(
  repositoryRoot,
  "src-tauri/target/release/bundle/macos/VisualTeX.app",
);
const workspaceVisualTeXExecutable = join(
  workspaceVisualTeXApp,
  "Contents/MacOS/visualtex",
);
const installedVisualTeXApp = "/Applications/VisualTeX.app";
const installedVisualTeXExecutable = join(
  installedVisualTeXApp,
  "Contents/MacOS/visualtex",
);
const installedVisualTeXBackupApp = join(
  tmpdir(),
  `VisualTeX-global-format-backup-${process.pid}.app`,
);
const wordFastOpenReadyPath = join(
  home,
  "Library/Containers/com.microsoft.Word/Data/Library/Application Support/VisualTeX/FastOpen/word/resident-ready",
);
const resultPath = join(
  testsRoot,
  "word-global-format-conversion-result.txt",
);

const formulas = [
  {
    formulaId: "61616161-6161-4616-8616-616161616161",
    lineId: "71717171-7171-4717-8717-717171717171",
    latex: String.raw`x^2+1`,
    displayMode: "inline" as const,
    numbered: false,
  },
  {
    formulaId: "62626262-6262-4626-8626-626262626262",
    lineId: "72727272-7272-4727-8727-727272727272",
    latex: String.raw`\frac{a}{b}`,
    displayMode: "block" as const,
    numbered: false,
  },
  {
    formulaId: "63636363-6363-4636-8636-636363636363",
    lineId: "73737373-7373-4737-8737-737373737373",
    latex: String.raw`\sum_{i=1}^{n}\left(\binom{n}{i}\right)`,
    displayMode: "block" as const,
    numbered: true,
  },
];

const metadataPaths = formulas.map((_, index) =>
  join(testsRoot, `word-global-format-metadata-${index + 1}.txt`),
);
const sleep = (milliseconds: number) =>
  new Promise((resolveSleep) => setTimeout(resolveSleep, milliseconds));

function run(program: string, args: string[], timeout = 30_000) {
  const result = spawnSync(program, args, {
    encoding: "utf8",
    timeout,
    killSignal: "SIGKILL",
    maxBuffer: 32 * 1024 * 1024,
  });
  if (result.status !== 0) {
    const detail = [
      result.stderr.trim(),
      result.stdout.trim(),
      result.error instanceof Error ? result.error.message : "",
      result.signal ? `signal=${result.signal}` : "",
      `status=${String(result.status)}`,
    ].filter(Boolean).join("\n");
    throw new Error(detail || `${program} failed`);
  }
  return result.stdout.trim();
}

function bestEffort(program: string, args: string[], timeout = 30_000) {
  try {
    return run(program, args, timeout);
  } catch {
    return "";
  }
}

function appleScript(lines: string[], timeout = 30_000) {
  return run(
    "/usr/bin/osascript",
    lines.flatMap((line) => ["-e", line]),
    timeout,
  );
}

function quitWord() {
  bestEffort(
    "/usr/bin/osascript",
    ["-e", 'tell application "Microsoft Word" to quit saving no'],
    20_000,
  );
  bestEffort("/usr/bin/killall", ["Microsoft Word"], 10_000);
}

function killVisualTeX() {
  bestEffort("/usr/bin/killall", ["visualtex"], 10_000);
}

async function waitForResidentBinding(
  pid: number,
  executablePath: string,
  timeoutMs = 25_000,
) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (existsSync(wordFastOpenReadyPath)) {
      const lines = readFileSync(wordFastOpenReadyPath, "utf8")
        .trim()
        .split(/\r?\n/);
      if (
        lines[0] === "visualtex-fast-open-ready-v2" &&
        lines[2] === String(pid) &&
        lines[3] === executablePath
      ) {
        return;
      }
    }
    await sleep(200);
  }
  throw new Error(
    `Workspace VisualTeX did not publish its v2 Word resident binding: ${wordFastOpenReadyPath}`,
  );
}

async function waitForWordReady(timeoutMs = 30_000) {
  spawnSync("/usr/bin/open", ["-a", "Microsoft Word"], {
    encoding: "utf8",
    timeout: 10_000,
  });
  const deadline = Date.now() + timeoutMs;
  let consecutiveReady = 0;
  while (Date.now() < deadline) {
    try {
      const version = appleScript([
        'tell application "Microsoft Word"',
        "activate",
        "return version",
        "end tell",
      ], 5_000);
      if (version) {
        consecutiveReady += 1;
        if (consecutiveReady >= 2) return;
        await sleep(500);
        continue;
      }
    } catch {
      // Retry while Word loads Startup add-ins.
    }
    consecutiveReady = 0;
    await sleep(500);
  }
  throw new Error("Microsoft Word did not become automation-ready");
}

function runWordMacro(name: string, timeout = 90_000) {
  return appleScript([
    'tell application "Microsoft Word"',
    "activate",
    `run VB macro macro name ${JSON.stringify(name)}`,
    "end tell",
  ], timeout);
}

function runWordSessionMacro(name: string, timeout = 90_000) {
  try {
    return runWordMacro(name, timeout);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (message.includes("-609") || message.includes("连接无效")) return "";
    throw error;
  }
}

async function pressGlobalConversionRibbonItem(label: string) {
  const deadline = Date.now() + 20_000;
  let lastError = "";
  while (Date.now() < deadline) {
    try {
      appleScript([
        'tell application "Microsoft Word" to activate',
        'tell application "System Events"',
        'tell application process "Microsoft Word"',
        'set frontmost to true',
        'set standardWindow to first window whose subrole is "AXStandardWindow"',
        'set ribbonTabs to tab group 1 of standardWindow',
        'set visualTeXTab to first radio button of ribbonTabs whose name is "VisualTeX"',
        'perform action "AXPress" of visualTeXTab',
        'delay 0.35',
        'set formatMenu to missing value',
        'repeat with candidate in (entire contents of ribbonTabs)',
        'try',
        'if role of candidate is "AXMenuButton" and name of candidate is "全文格式转换" then',
        'set formatMenu to contents of candidate',
        'exit repeat',
        'end if',
        'end try',
        'end repeat',
        'if formatMenu is missing value then',
        'set ribbonSummary to ""',
        'repeat with candidate in (entire contents of ribbonTabs)',
        'try',
        'set candidateRole to role of candidate as text',
        'set candidateName to name of candidate as text',
        'if candidateRole is "AXButton" or candidateRole is "AXMenuButton" then set ribbonSummary to ribbonSummary & candidateRole & ":" & candidateName & " | "',
        'end try',
        'end repeat',
        'error "VisualTeX global conversion menu not found; rendered=" & ribbonSummary',
        'end if',
        'perform action "AXPress" of formatMenu',
        'delay 0.25',
        'set menuWindow to first window whose name is "全文格式转换"',
        'set targetItem to missing value',
        'repeat with candidate in (entire contents of menuWindow)',
        'try',
        `if role of candidate is "AXMenuButton" and name of candidate is ${JSON.stringify(label)} then`,
        'set targetItem to contents of candidate',
        'exit repeat',
        'end if',
        'end try',
        'end repeat',
        'if targetItem is missing value then error "VisualTeX global conversion menu item not found"',
        'perform action "AXPress" of targetItem',
        'end tell',
        'end tell',
      ], 10_000);
      return;
    } catch (error) {
      lastError = error instanceof Error ? error.message : String(error);
      await sleep(350);
    }
  }
  throw new Error(`Timed out pressing Word Ribbon item ${label}: ${lastError}`);
}

function sessionIds() {
  if (!existsSync(sessionsRoot)) return new Set<string>();
  return new Set(
    readdirSync(sessionsRoot, { withFileTypes: true })
      .filter((entry) => entry.isDirectory())
      .map((entry) => entry.name),
  );
}

function sessionRequest(sessionId: string) {
  const requestPath = join(sessionsRoot, sessionId, "request.json");
  if (!existsSync(requestPath)) return null;
  try {
    return JSON.parse(readFileSync(requestPath, "utf8")) as {
      operation?: string;
      host?: string;
      documentImport?: {
        sourceKind?: string;
        outputKind?: string;
        redrawScope?: string;
      };
    };
  } catch {
    return null;
  }
}

async function waitForFormulaRestoreSession(
  before: Set<string>,
  sourceKind: "image" | "omml",
  outputKind: "omml" | "image",
  timeoutMs = 30_000,
) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    for (const sessionId of sessionIds()) {
      if (before.has(sessionId)) continue;
      const request = sessionRequest(sessionId);
      if (
        request?.operation === "formulaRestore" &&
        request.host === "word" &&
        request.documentImport?.sourceKind === sourceKind &&
        request.documentImport?.outputKind === outputKind &&
        request.documentImport?.redrawScope === "document"
      ) {
        return sessionId;
      }
    }
    await sleep(100);
  }
  throw new Error(
    `Timed out waiting for formulaRestore ${sourceKind}->${outputKind} Session`,
  );
}

async function waitForDocumentImportComplete(
  sessionId: string,
  timeoutMs = 120_000,
) {
  const progressPath = join(
    sessionsRoot,
    sessionId,
    "document-import-progress.txt",
  );
  const deadline = Date.now() + timeoutMs;
  let lastProgress = "";
  while (Date.now() < deadline) {
    if (existsSync(progressPath)) {
      lastProgress = readFileSync(progressPath, "utf8");
      if (/stage=complete(?:\r?\n|$)/.test(lastProgress)) return;
      if (/stage=error(?:\r?\n|$)/.test(lastProgress)) {
        throw new Error(
          `Word global format conversion reported an error:\n${lastProgress}`,
        );
      }
    }
    await sleep(150);
  }
  throw new Error(
    `Timed out waiting for global format conversion ${sessionId}: ${lastProgress || "no progress"}`,
  );
}

function inspectWordState() {
  rmSync(resultPath, { force: true });
  runWordMacro("VisualTeX_InspectGlobalFormatConversionRegression", 30_000);
  if (!existsSync(resultPath)) {
    throw new Error(`Word inspection did not write ${resultPath}`);
  }
  const text = readFileSync(resultPath, "utf8").trim();
  const values = new Map<string, string>();
  for (const line of text.split(/\r?\n/)) {
    const separator = line.indexOf("=");
    if (separator > 0) values.set(line.slice(0, separator), line.slice(separator + 1));
  }
  return { text, values };
}

function requireState(
  state: ReturnType<typeof inspectWordState>,
  expectedKind: "image" | "omml",
) {
  for (let index = 1; index <= formulas.length; index += 1) {
    if (state.values.get(`kind${index}`) !== expectedKind) {
      throw new Error(
        `Formula ${index} identity was not preserved as ${expectedKind}:\n${state.text}`,
      );
    }
  }
  const expectedImages = expectedKind === "image" ? "3" : "0";
  const expectedOmml = expectedKind === "omml" ? "3" : "0";
  if (
    state.values.get("images") !== expectedImages ||
    state.values.get("omml") !== expectedOmml ||
    state.values.get("tables") !== "0" ||
    state.values.get("sequences") !== "1" ||
    state.values.get("number3") !== "1"
  ) {
    throw new Error(`Global format conversion structure mismatch:\n${state.text}`);
  }
}

type DocumentSnapshot = {
  path: string;
  images: number;
  omml: number;
  tables: number;
  formulas: Array<{ formulaId: string; kind: string; numbered: boolean; number: string;
    fontSizePt: number; width: number | null; height: number | null;
    referenceWidthPt: number; referenceHeightPt: number; latex: string; mathText: string | null }>;
  sequences: Array<{ code: string; result: string; owner: string | null }>;
};

function saveAndInspectDocument(stage: string): DocumentSnapshot {
  const path = join(testsRoot, `${snapshotPrefix}-${stage}.docx`);
  appleScript([
    'tell application "Microsoft Word"',
    "set documentObject to active document",
    `save as documentObject file name ${JSON.stringify(path)} file format format document default add to recent files false`,
    "end tell",
  ], 30_000);
  const result = run("python3", [join(repositoryRoot, "scripts/inspect_word_conversion_docx.py"), path]);
  writeFileSync(path.replace(/\.docx$/, ".json"), result);
  return JSON.parse(result);
}

function requireDocumentSnapshot(state: DocumentSnapshot, baseline: DocumentSnapshot, kind: "image" | "omml") {
  if (state.formulas.length !== baseline.formulas.length || state.tables !== baseline.tables ||
      state.sequences.length !== baseline.sequences.length) {
    throw new Error(`Document inventory changed unexpectedly: ${state.path}`);
  }
  for (const original of baseline.formulas) {
    const current = state.formulas.find(f => f.formulaId === original.formulaId);
    if (!current || current.kind !== kind || current.numbered !== original.numbered ||
        current.number !== original.number || Math.abs(current.fontSizePt - original.fontSizePt) > 0.1) {
      throw new Error(`Formula identity/number/font mismatch: ${original.formulaId}; ${JSON.stringify(current)}`);
    }
    if (kind === "image") {
      const widthScale = current.width! / current.referenceWidthPt;
      const heightScale = current.height! / current.referenceHeightPt;
      const expectedScale = current.fontSizePt / 14;
      if (Math.abs(widthScale / heightScale - 1) > 0.02 ||
          Math.abs(heightScale / expectedScale - 1) > 0.03) {
        throw new Error(`Formula dimensions were distorted: ${JSON.stringify(current)}`);
      }
    }
  }
  const ordinary = (snapshot: DocumentSnapshot) => snapshot.sequences.filter(f => !f.owner)
    .map(f => [f.code.trim(), f.result.trim()]);
  if (JSON.stringify(ordinary(state)) !== JSON.stringify(ordinary(baseline))) {
    throw new Error(`Ordinary Word SEQ fields changed: ${state.path}`);
  }
}

if (!existsSync(workspaceVisualTeXExecutable)) {
  throw new Error(`Workspace VisualTeX app bundle is missing: ${workspaceVisualTeXExecutable}`);
}

mkdirSync(startupRoot, { recursive: true });
mkdirSync(testsRoot, { recursive: true });
mkdirSync(wordApplicationScriptsRoot, { recursive: true });
let backedUpAddin = false;
let backedUpWordScript = false;
let backedUpInstalledApp = false;
let addinChanged = false;
let wordScriptChanged = false;
let installedAppChanged = false;
const existingVisualTeX = bestEffort("/usr/bin/pgrep", ["-x", "visualtex"])
  .split(/\s+/)
  .filter(Boolean);

try {
  quitWord();
  killVisualTeX();
  await sleep(1_000);

  rmSync(installedBackupPath, { force: true });
  if (existsSync(installedAddinPath)) {
    run("/bin/cp", [installedAddinPath, installedBackupPath], 15_000);
    backedUpAddin = true;
    addinChanged = true;
    rmSync(installedAddinPath, { force: true });
  }
  addinChanged = true;
  copyFileSync(wordAddinPath, installedAddinPath);

  rmSync(installedWordScriptBackupPath, { force: true });
  if (existsSync(installedWordScriptPath)) {
    copyFileSync(installedWordScriptPath, installedWordScriptBackupPath);
    backedUpWordScript = true;
    wordScriptChanged = true;
    rmSync(installedWordScriptPath, { force: true });
  }
  wordScriptChanged = true;
  run("/usr/bin/osacompile", [
    "-o",
    installedWordScriptPath,
    wordScriptSourcePath,
  ], 30_000);

  formulas.forEach((formula, index) => {
    const metadata = createFormulaMetadata({
      formulaId: formula.formulaId,
      title: `Global format regression ${index + 1}`,
      lines: [{ id: formula.lineId, latex: formula.latex }],
      codeFormat: "raw",
      sourceLatex: formula.latex,
      displayMode: formula.displayMode,
      numbered: formula.numbered,
      fontSizePt: 14,
      referenceWidthPt: index === 2 ? 150 : index === 1 ? 76 : 60,
      referenceHeightPt: index === 2 ? 42 : index === 1 ? 28 : 20,
      referenceBaselinePt: 0,
      appVersion: "1.2.6",
    });
    writeFileSync(metadataPaths[index], encodeFormulaMetadata(metadata), {
      mode: 0o600,
    });
  });

  rmSync(installedVisualTeXBackupApp, { recursive: true, force: true });
  if (existsSync(installedVisualTeXApp)) {
    run("/usr/bin/ditto", [installedVisualTeXApp, installedVisualTeXBackupApp], 120_000);
    backedUpInstalledApp = true;
    installedAppChanged = true;
    rmSync(installedVisualTeXApp, { recursive: true, force: true });
  }
  installedAppChanged = true;
  run("/usr/bin/ditto", [workspaceVisualTeXApp, installedVisualTeXApp], 120_000);
  if (!existsSync(installedVisualTeXExecutable)) {
    throw new Error(`Installed test VisualTeX app is missing: ${installedVisualTeXExecutable}`);
  }

  const visualTeX = spawn(
    installedVisualTeXExecutable,
    ["-ApplePersistenceIgnoreState", "YES", "--office-background"],
    { detached: true, stdio: "ignore" },
  );
  visualTeX.unref();
  await waitForResidentBinding(visualTeX.pid!, installedVisualTeXExecutable);

  await waitForWordReady();
  appleScript([
    'tell application "Microsoft Word"',
    'if not (exists add in "VisualTeX.dotm") then error "VisualTeX.dotm was not registered from Startup"',
    'set installed of add in "VisualTeX.dotm" to true',
    "end tell",
  ]);
  runWordMacro("VisualTeX_PerformanceNoop", 30_000);
  if (sourceDocumentPath) {
    appleScript(['tell application "Microsoft Word"',
      `open file name ${JSON.stringify(resolve(sourceDocumentPath))} add to recent files false`, "end tell"]);
    const openDeadline = Date.now() + 30_000;
    let documentReady = false;
    while (Date.now() < openDeadline) {
      const documentName = bestEffort("/usr/bin/osascript", ["-e",
        'tell application "Microsoft Word" to get name of active document'], 5_000);
      if (documentName === basename(sourceDocumentPath) ||
          documentName === basename(sourceDocumentPath, ".docx")) {
        documentReady = true;
        break;
      }
      await sleep(200);
    }
    if (!documentReady) throw new Error(`Word did not finish opening ${sourceDocumentPath}`);
  } else {
    runWordMacro("VisualTeX_SetupGlobalFormatConversionRegression", 60_000);
    requireState(inspectWordState(), "image");
  }
  const baseline = saveAndInspectDocument("before");

  let before = sessionIds();
  const imageToOmmlStarted = Date.now();
  if (useRibbon) {
    await pressGlobalConversionRibbonItem("全文图片公式 → Word OMML");
  } else {
    runWordSessionMacro("VisualTeX_ConvertDocumentImagesToOmml", 60_000);
  }
  const imageToOmmlSession = await waitForFormulaRestoreSession(
    before,
    "image",
    "omml",
  );
  await waitForDocumentImportComplete(imageToOmmlSession);
  const imageToOmmlMs = Date.now() - imageToOmmlStarted;
  if (!sourceDocumentPath) requireState(inspectWordState(), "omml");
  const nativeSnapshot = saveAndInspectDocument("omml");
  requireDocumentSnapshot(nativeSnapshot, baseline, "omml");

  before = sessionIds();
  const ommlToImageStarted = Date.now();
  if (useRibbon) {
    await pressGlobalConversionRibbonItem("全文 OMML → 图片公式");
  } else {
    runWordSessionMacro("VisualTeX_ConvertDocumentOmmlToImages", 60_000);
  }
  const ommlToImageSession = await waitForFormulaRestoreSession(
    before,
    "omml",
    "image",
  );
  await waitForDocumentImportComplete(ommlToImageSession);
  const ommlToImageMs = Date.now() - ommlToImageStarted;
  if (!sourceDocumentPath) requireState(inspectWordState(), "image");
  const imageSnapshot = saveAndInspectDocument("image");
  requireDocumentSnapshot(imageSnapshot, baseline, "image");

  process.stdout.write(
    [
      "Word global format conversion E2E: PASS",
      `imageToOmmlSessionId=${imageToOmmlSession}`,
      `imageToOmmlMs=${imageToOmmlMs}`,
      `ommlToImageSessionId=${ommlToImageSession}`,
      `ommlToImageMs=${ommlToImageMs}`,
      `formulaIds=${baseline.formulas.map((formula) => formula.formulaId).join(",")}`,
      `numberedFormulas=${baseline.formulas.filter(f => f.numbered).length}`,
      `tables=${imageSnapshot.tables}`,
      `sequenceCount=${imageSnapshot.sequences.length}`,
      `snapshot=${imageSnapshot.path}`,
    ].join("\n") + "\n",
  );
} finally {
  quitWord();
  killVisualTeX();
  if (addinChanged) {
    rmSync(installedAddinPath, { force: true });
    if (backedUpAddin && existsSync(installedBackupPath)) {
      copyFileSync(installedBackupPath, installedAddinPath);
    }
  }
  rmSync(installedBackupPath, { force: true });
  if (wordScriptChanged) {
    rmSync(installedWordScriptPath, { force: true });
    if (backedUpWordScript && existsSync(installedWordScriptBackupPath)) {
      copyFileSync(installedWordScriptBackupPath, installedWordScriptPath);
    }
  }
  rmSync(installedWordScriptBackupPath, { force: true });
  if (installedAppChanged) {
    rmSync(installedVisualTeXApp, { recursive: true, force: true });
    if (backedUpInstalledApp && existsSync(installedVisualTeXBackupApp)) {
      run("/usr/bin/ditto", [installedVisualTeXBackupApp, installedVisualTeXApp], 120_000);
    }
  }
  rmSync(installedVisualTeXBackupApp, { recursive: true, force: true });
  if (existingVisualTeX.length > 0) {
    bestEffort(
      "/usr/bin/open",
      ["-gj", "-b", "com.visualtex.studio", "--args", "--office-background"],
      20_000,
    );
  }
}
