import { spawn, spawnSync } from "node:child_process";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { homedir, tmpdir } from "node:os";
import { basename, join, resolve } from "node:path";
import {
  createFormulaMetadata,
  decodeFormulaMetadata,
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
if (!wordAddinPath || !existsSync(wordAddinPath)) {
  throw new Error(`Missing compiled Word add-in: ${wordAddinPath}`);
}

const home = homedir();
const sourceFormulaId = "68686868-6868-4868-8868-686868686868";
const sourceLineId = "78787878-7878-4878-8878-787878787878";
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
  `visualtex-word-copy-conversion-startup-${process.pid}.dotm`,
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
  `visualtex-word-copy-conversion-script-${process.pid}.scpt`,
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
const stagedVisualTeXRoot = join(
  "/Applications",
  `VisualTeX-E2E-${process.pid}`,
);
const stagedVisualTeXApp = join(stagedVisualTeXRoot, "VisualTeX.app");
const stagedVisualTeXExecutable = join(
  stagedVisualTeXApp,
  "Contents/MacOS/visualtex",
);
const wordFastOpenReadyPath = join(
  home,
  "Library/Containers/com.microsoft.Word/Data/Library/Application Support/VisualTeX/FastOpen/word/resident-ready",
);
const metadataPath = join(testsRoot, "word-copy-conversion-metadata.txt");
const documentPath = join(testsRoot, "word-copy-conversion-document.txt");
const imageStartPath = join(testsRoot, "word-copy-image-to-native-start.txt");
const imageResultPath = join(testsRoot, "word-copy-image-to-native-result.txt");
const nativeStartPath = join(testsRoot, "word-copy-native-to-image-start.txt");
const nativeResultPath = join(testsRoot, "word-copy-native-to-image-result.txt");
const fixtureDocumentPath = join(
  home,
  "Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch/word-copy-conversion-fixture.docx",
);

const resultPaths = [
  documentPath,
  imageStartPath,
  imageResultPath,
  nativeStartPath,
  nativeResultPath,
];
const sleep = (milliseconds: number) =>
  new Promise((resolveSleep) => setTimeout(resolveSleep, milliseconds));

function run(program: string, args: string[], timeout = 30_000) {
  const result = spawnSync(program, args, {
    encoding: "utf8",
    timeout,
    maxBuffer: 32 * 1024 * 1024,
  });
  if (result.status !== 0) {
    throw new Error(
      result.stderr.trim() || result.stdout.trim() || `${program} failed`,
    );
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
  timeoutMs = 20_000,
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

function parseKeyValueFile(path: string) {
  const text = readFileSync(path, "utf8").trim();
  const lines = text.split(/\r?\n/).filter(Boolean);
  const values = new Map<string, string>();
  for (const line of lines.slice(1)) {
    const separator = line.indexOf("=");
    if (separator > 0) {
      values.set(line.slice(0, separator), line.slice(separator + 1));
    }
  }
  return { status: lines[0] ?? "", values, text };
}

async function waitForFile(path: string, prefix: string, timeoutMs = 30_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (existsSync(path)) {
      const text = readFileSync(path, "utf8").trim();
      if (text.startsWith(prefix)) return text;
      if (text.startsWith("FAIL")) throw new Error(`${basename(path)}:\n${text}`);
    }
    await sleep(200);
  }
  throw new Error(`Timed out waiting for ${path} to start with ${prefix}`);
}

async function waitForSessionPerformanceStage(
  sessionId: string,
  stage: string,
  timeoutMs = 60_000,
) {
  const performancePath = join(
    sessionsRoot,
    sessionId,
    "editor-performance.jsonl",
  );
  const needle = `\"stage\":\"${stage}\"`;
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (existsSync(performancePath)) {
      const text = readFileSync(performancePath, "utf8");
      if (text.includes(needle)) return;
    }
    await sleep(200);
  }
  throw new Error(
    `Timed out waiting for Session ${sessionId} performance stage ${stage}`,
  );
}
function assertForkedMetadata(
  encoded: string,
  expectedFormulaId: string,
  forbiddenLineIds: Set<string>,
) {
  const metadata = decodeFormulaMetadata(encoded);
  if (!metadata) throw new Error("Committed copied formula metadata did not decode");
  if (metadata.formulaId !== expectedFormulaId) {
    throw new Error(
      `Committed metadata formulaId mismatch: ${metadata.formulaId} != ${expectedFormulaId}`,
    );
  }
  for (const line of metadata.lines) {
    if (forbiddenLineIds.has(line.id)) {
      throw new Error(`Copied formula retained source lineId ${line.id}`);
    }
  }
  return metadata;
}

if (!existsSync(workspaceVisualTeXExecutable)) {
  throw new Error("Workspace VisualTeX app bundle is missing");
}

const sourceMetadata = createFormulaMetadata({
  formulaId: sourceFormulaId,
  title: "Word copy conversion regression",
  lines: [{ id: sourceLineId, latex: String.raw`x^2+1` }],
  codeFormat: "latex",
  displayMode: "inline",
  numbered: false,
  fontSizePt: 14,
  referenceWidthPt: 36,
  referenceHeightPt: 16,
  referenceBaselinePt: 0,
  appVersion: "1.2.6",
});
const encodedSourceMetadata = encodeFormulaMetadata(sourceMetadata);

mkdirSync(startupRoot, { recursive: true });
mkdirSync(testsRoot, { recursive: true });
mkdirSync(wordApplicationScriptsRoot, { recursive: true });
let backedUpAddin = false;
let backedUpWordScript = false;
const existingVisualTeX = bestEffort("/usr/bin/pgrep", ["-x", "visualtex"])
  .split(/\s+/)
  .filter(Boolean);

try {
  quitWord();
  killVisualTeX();
  await sleep(1_000);

  rmSync(installedBackupPath, { force: true });
  if (existsSync(installedAddinPath)) {
    copyFileSync(installedAddinPath, installedBackupPath);
    backedUpAddin = true;
    rmSync(installedAddinPath, { force: true });
  }
  copyFileSync(wordAddinPath, installedAddinPath);

  rmSync(installedWordScriptBackupPath, { force: true });
  if (existsSync(installedWordScriptPath)) {
    copyFileSync(installedWordScriptPath, installedWordScriptBackupPath);
    backedUpWordScript = true;
    rmSync(installedWordScriptPath, { force: true });
  }
  run("/usr/bin/osacompile", [
    "-o",
    installedWordScriptPath,
    wordScriptSourcePath,
  ], 30_000);
  const installedScriptSource = run(
    "/usr/bin/osadecompile",
    [installedWordScriptPath],
    30_000,
  );
  if (!installedScriptSource.includes("visualtex-fast-open-ready-v2")) {
    throw new Error("Temporary Word AppleScriptTask does not contain v2 resident binding");
  }

  for (const path of resultPaths) rmSync(path, { force: true });
  writeFileSync(metadataPath, encodedSourceMetadata, { mode: 0o600 });

  rmSync(stagedVisualTeXRoot, { recursive: true, force: true });
  mkdirSync(stagedVisualTeXRoot, { recursive: true });
  run("/usr/bin/ditto", [workspaceVisualTeXApp, stagedVisualTeXApp], 120_000);
  if (!existsSync(stagedVisualTeXExecutable)) {
    throw new Error(`Staged VisualTeX app is missing: ${stagedVisualTeXExecutable}`);
  }

  const visualTeX = spawn(
    stagedVisualTeXExecutable,
    ["-ApplePersistenceIgnoreState", "YES", "--office-background"],
    {
    detached: true,
      stdio: "ignore",
    },
  );
  visualTeX.unref();
  await waitForResidentBinding(visualTeX.pid!, stagedVisualTeXExecutable);

  const runningExecutable = bestEffort("/bin/ps", [
    "-p",
    String(visualTeX.pid),
    "-o",
    "command=",
  ]);
  if (!runningExecutable.includes(stagedVisualTeXExecutable)) {
    throw new Error(`Workspace VisualTeX resident did not stay running: ${runningExecutable}`);
  }

  await waitForWordReady();
  appleScript([
    'tell application "Microsoft Word"',
    'if not (exists add in "VisualTeX.dotm") then error "VisualTeX.dotm was not registered from Startup"',
    'set installed of add in "VisualTeX.dotm" to true',
    "end tell",
  ]);
  runWordMacro("VisualTeX_PerformanceNoop", 30_000);

  runWordMacro("VisualTeX_RunWordCopiedImageToNativeConversionRegression");
  await waitForFile(imageStartPath, "START");
  const imageStart = parseKeyValueFile(imageStartPath);
  const imageSessionId = imageStart.values.get("sessionId") ?? "";
  const documentName = imageStart.values.get("documentName") ?? "";
  if (!imageSessionId || !documentName) {
    throw new Error(`Image conversion start result is incomplete:\n${imageStart.text}`);
  }
  await waitForSessionPerformanceStage(
    imageSessionId,
    "apply-backend-complete",
    90_000,
  );
  runWordMacro("VisualTeX_InspectWordCopiedImageToNativeConversionRegression");
  await waitForFile(imageResultPath, "PASS");
  const imageResult = parseKeyValueFile(imageResultPath);
  const nativeFormulaId = imageResult.values.get("newFormulaId") ?? "";
  const nativeEncodedMetadata = imageResult.values.get("metadata") ?? "";
  if (!nativeFormulaId || nativeFormulaId === sourceFormulaId) {
    throw new Error("Image→OMML copied conversion retained the source formulaId");
  }
  const nativeMetadata = assertForkedMetadata(
    nativeEncodedMetadata,
    nativeFormulaId,
    new Set(sourceMetadata.lines.map((line) => line.id)),
  );

  rmSync(nativeStartPath, { force: true });
  rmSync(nativeResultPath, { force: true });
  runWordMacro("VisualTeX_RunWordCopiedNativeToImageConversionRegression");
  await waitForFile(nativeStartPath, "START");
  const nativeStart = parseKeyValueFile(nativeStartPath);
  const nativeSessionId = nativeStart.values.get("sessionId") ?? "";
  if (!nativeSessionId) {
    throw new Error(`Native conversion start result is incomplete:\n${nativeStart.text}`);
  }
  await waitForSessionPerformanceStage(
    nativeSessionId,
    "apply-backend-complete",
    90_000,
  );
  runWordMacro("VisualTeX_InspectWordCopiedNativeToImageConversionRegression");
  await waitForFile(nativeResultPath, "PASS");
  const nativeResult = parseKeyValueFile(nativeResultPath);
  const imageFormulaId = nativeResult.values.get("newImageFormulaId") ?? "";
  const imageEncodedMetadata = nativeResult.values.get("metadata") ?? "";
  if (
    !imageFormulaId ||
    imageFormulaId === sourceFormulaId ||
    imageFormulaId === nativeFormulaId
  ) {
    throw new Error("OMML→image copied conversion retained an earlier formulaId");
  }
  const imageMetadata = assertForkedMetadata(
    imageEncodedMetadata,
    imageFormulaId,
    new Set([
      ...sourceMetadata.lines.map((line) => line.id),
      ...nativeMetadata.lines.map((line) => line.id),
    ]),
  );

  process.stdout.write(
    [
      "Word copied formula conversion E2E: PASS",
      `sourceFormulaId=${sourceFormulaId}`,
      `sourceLineId=${sourceLineId}`,
      `imageToNativeSessionId=${imageSessionId}`,
      `nativeFormulaId=${nativeFormulaId}`,
      `nativeLineIds=${nativeMetadata.lines.map((line) => line.id).join(",")}`,
      `nativeToImageSessionId=${nativeSessionId}`,
      `imageFormulaId=${imageFormulaId}`,
      `imageLineIds=${imageMetadata.lines.map((line) => line.id).join(",")}`,
    ].join("\n") + "\n",
  );
} finally {
  quitWord();
  killVisualTeX();
  rmSync(installedAddinPath, { force: true });
  if (backedUpAddin && existsSync(installedBackupPath)) {
    copyFileSync(installedBackupPath, installedAddinPath);
  }
  rmSync(installedBackupPath, { force: true });
  rmSync(fixtureDocumentPath, { force: true });
  rmSync(installedWordScriptPath, { force: true });
  if (backedUpWordScript && existsSync(installedWordScriptBackupPath)) {
    copyFileSync(installedWordScriptBackupPath, installedWordScriptPath);
  }
  rmSync(installedWordScriptBackupPath, { force: true });
  rmSync(stagedVisualTeXRoot, { recursive: true, force: true });
  if (existingVisualTeX.length > 0) {
    bestEffort(
      "/usr/bin/open",
      ["-gj", "-b", "com.visualtex.studio", "--args", "--office-background"],
      20_000,
    );
  }
}
