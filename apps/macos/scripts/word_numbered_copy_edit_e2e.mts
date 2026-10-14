import { spawn, spawnSync } from "node:child_process";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
} from "node:fs";
import { homedir, tmpdir } from "node:os";
import { basename, join, resolve } from "node:path";
import { decodeFormulaMetadata } from "../src/office/shared/formulaMetadata.ts";

function option(name: string) {
  const index = process.argv.indexOf(name);
  if (index < 0) return "";
  const value = process.argv[index + 1];
  if (!value || value.startsWith("--")) throw new Error(`${name} requires a value`);
  return value;
}

const repositoryRoot = resolve(new URL("..", import.meta.url).pathname);
const wordAddinPath = resolve(option("--word-addin"));
const reproDocumentPath = resolve(option("--repro-doc"));
const explicitVisualTeXAppPath = option("--visualtex-app");
if (!existsSync(wordAddinPath)) throw new Error(`Missing Word add-in: ${wordAddinPath}`);
if (!existsSync(reproDocumentPath)) throw new Error(`Missing repro document: ${reproDocumentPath}`);

const home = homedir();
const startupRoot = join(
  home,
  "Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word",
);
const installedAddinPath = join(startupRoot, "VisualTeX.dotm");
const installedBackupPath = join(tmpdir(), `visualtex-numbered-edit-startup-${process.pid}.dotm`);
const wordApplicationScriptsRoot = join(home, "Library/Application Scripts/com.microsoft.Word");
const installedWordScriptPath = join(wordApplicationScriptsRoot, "VisualTeXWord.scpt");
const installedWordScriptBackupPath = join(tmpdir(), `visualtex-numbered-edit-script-${process.pid}.scpt`);
const wordScriptSourcePath = join(repositoryRoot, "office/macos-offline/word/VisualTeXWord.scpt");
const workspaceVisualTeXApp = explicitVisualTeXAppPath
  ? resolve(explicitVisualTeXAppPath)
  : join(repositoryRoot, "src-tauri/target/release/bundle/macos/VisualTeX.app");
const stagedVisualTeXRoot = join("/Applications", `VisualTeX-NumberedEditE2E-${process.pid}`);
const stagedVisualTeXApp = join(stagedVisualTeXRoot, "VisualTeX.app");
const stagedVisualTeXExecutable = join(stagedVisualTeXApp, "Contents/MacOS/visualtex");
const wordFastOpenReadyPath = join(
  home,
  "Library/Containers/com.microsoft.Word/Data/Library/Application Support/VisualTeX/FastOpen/word/resident-ready",
);
const sessionsRoot = join(home, "Library/Application Support/com.visualtex.studio/office/sessions");
const reproDocumentName = basename(reproDocumentPath);
const sleep = (milliseconds: number) => new Promise((resolveSleep) => setTimeout(resolveSleep, milliseconds));
const clipboardBefore = bestEffort("/usr/bin/pbpaste", []);

function run(program: string, args: string[], timeout = 30_000, input?: string) {
  const result = spawnSync(program, args, { encoding: "utf8", timeout, maxBuffer: 32 * 1024 * 1024, input });
  if (result.status !== 0) {
    throw new Error(result.stderr.trim() || result.stdout.trim() || `${program} failed`);
  }
  return result.stdout.trim();
}

function bestEffort(program: string, args: string[], timeout = 30_000) {
  try { return run(program, args, timeout); } catch { return ""; }
}

function appleScript(lines: string[], timeout = 30_000) {
  return run("/usr/bin/osascript", lines.flatMap((line) => ["-e", line]), timeout);
}

function quitWord() {
  bestEffort("/usr/bin/osascript", ["-e", 'tell application "Microsoft Word" to quit saving no'], 20_000);
  bestEffort("/usr/bin/killall", ["Microsoft Word"], 10_000);
}

function killVisualTeX() {
  bestEffort("/usr/bin/killall", ["visualtex"], 10_000);
}

async function waitForWordReady(timeoutMs = 30_000) {
  spawnSync("/usr/bin/open", ["-a", "Microsoft Word"], { encoding: "utf8", timeout: 10_000 });
  const deadline = Date.now() + timeoutMs;
  let readyCount = 0;
  while (Date.now() < deadline) {
    try {
      const version = appleScript(['tell application "Microsoft Word"', "activate", "return version", "end tell"], 5_000);
      if (version) {
        readyCount += 1;
        if (readyCount >= 2) return;
        await sleep(400);
        continue;
      }
    } catch {}
    readyCount = 0;
    await sleep(400);
  }
  throw new Error("Microsoft Word did not become automation-ready");
}

async function waitForResidentBinding(pid: number, timeoutMs = 20_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (existsSync(wordFastOpenReadyPath)) {
      const lines = readFileSync(wordFastOpenReadyPath, "utf8").trim().split(/\r?\n/);
      if (lines[0] === "visualtex-fast-open-ready-v2" && lines[2] === String(pid) && lines[3] === stagedVisualTeXExecutable) return;
    }
    await sleep(200);
  }
  throw new Error("VisualTeX did not publish the expected Word resident binding");
}

function sessionIds() {
  if (!existsSync(sessionsRoot)) return new Set<string>();
  return new Set(readdirSync(sessionsRoot, { withFileTypes: true }).filter((entry) => entry.isDirectory()).map((entry) => entry.name));
}

async function waitForNewWordEditSession(previous: Set<string>, timeoutMs = 30_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (existsSync(sessionsRoot)) {
      for (const entry of readdirSync(sessionsRoot, { withFileTypes: true })) {
        if (!entry.isDirectory() || previous.has(entry.name)) continue;
        const path = join(sessionsRoot, entry.name, "session.json");
        if (!existsSync(path)) continue;
        try {
          const session = JSON.parse(readFileSync(path, "utf8"));
          if (session.host === "word" && session.mode === "edit" && session.sourceDocumentId) {
            return { path, session };
          }
        } catch {}
      }
    }
    await sleep(150);
  }
  throw new Error("Timed out waiting for a new Word edit session");
}

async function waitForSessionDirty(sessionPath: string, timeoutMs = 15_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const session = JSON.parse(readFileSync(sessionPath, "utf8"));
      if (session.dirty === true) return session;
    } catch {}
    await sleep(150);
  }
  throw new Error("The real VisualTeX editor did not mark the copied formula session dirty");
}

async function waitForSessionCompleted(sessionPath: string, timeoutMs = 60_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const session = JSON.parse(readFileSync(sessionPath, "utf8"));
      if (session.status === "completed") return session;
      if (session.status === "failed") {
        throw new Error(`Office session failed: ${session.error ?? "unknown error"}`);
      }
    } catch (error) {
      if (error instanceof Error && error.message.startsWith("Office session failed:")) throw error;
    }
    await sleep(150);
  }
  throw new Error("Timed out waiting for the copied formula Apply to complete");
}

function readWordFormulaState() {
  const countText = appleScript([
    'tell application "Microsoft Word"',
    `activate document ${JSON.stringify(reproDocumentName)}`,
    "return count of inline shapes of active document",
    "end tell",
  ]);
  const count = Number.parseInt(countText, 10);
  if (!Number.isFinite(count) || count < 1) throw new Error(`Invalid Word inline shape count: ${countText}`);
  const shapes: Array<{ index: number; title: string; metadataEnvelope: string; metadata: ReturnType<typeof decodeFormulaMetadata> }> = [];
  for (let index = 1; index <= count; index += 1) {
    const title = appleScript([
      'tell application "Microsoft Word"',
      `activate document ${JSON.stringify(reproDocumentName)}`,
      `return title of inline shape ${index} of active document`,
      "end tell",
    ]);
    const metadataEnvelope = appleScript([
      'tell application "Microsoft Word"',
      `activate document ${JSON.stringify(reproDocumentName)}`,
      `return alternative text of inline shape ${index} of active document`,
      "end tell",
    ]);
    shapes.push({ index, title, metadataEnvelope, metadata: decodeFormulaMetadata(metadataEnvelope) });
  }
  const bookmarks = appleScript([
    'tell application "Microsoft Word"',
    `activate document ${JSON.stringify(reproDocumentName)}`,
    "return name of every bookmark of active document",
    "end tell",
  ]);
  return { shapes, bookmarks };
}

if (!existsSync(workspaceVisualTeXApp)) throw new Error("Release VisualTeX.app is missing");
mkdirSync(startupRoot, { recursive: true });
mkdirSync(wordApplicationScriptsRoot, { recursive: true });
let backedUpAddin = false;
let backedUpWordScript = false;
const preexistingVisualTeX = bestEffort("/usr/bin/pgrep", ["-x", "visualtex"]).split(/\s+/).filter(Boolean);

try {
  quitWord();
  killVisualTeX();
  await sleep(800);

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
  run("/usr/bin/osacompile", ["-o", installedWordScriptPath, wordScriptSourcePath]);

  rmSync(stagedVisualTeXRoot, { recursive: true, force: true });
  mkdirSync(stagedVisualTeXRoot, { recursive: true });
  run("/usr/bin/ditto", [workspaceVisualTeXApp, stagedVisualTeXApp], 120_000);
  const visualTeX = spawn(stagedVisualTeXExecutable, ["-ApplePersistenceIgnoreState", "YES", "--office-background"], {
    detached: true,
    stdio: "ignore",
  });
  visualTeX.unref();
  if (!visualTeX.pid) throw new Error("VisualTeX resident did not start");
  await waitForResidentBinding(visualTeX.pid);

  const priorSessions = sessionIds();
  await waitForWordReady();
  appleScript([
    'tell application "Microsoft Word"',
    'if not (exists add in "VisualTeX.dotm") then error "VisualTeX.dotm was not registered from Startup"',
    'set installed of add in "VisualTeX.dotm" to true',
    `open file name ${JSON.stringify(reproDocumentPath)}`,
    `activate document ${JSON.stringify(reproDocumentName)}`,
    "if (count of inline shapes of active document) is not 2 then error \"Expected exactly two copied formula images\"",
    "end tell",
  ], 40_000);

  const sourceTitleBefore = appleScript([
    'tell application "Microsoft Word"',
    `activate document ${JSON.stringify(reproDocumentName)}`,
    "return title of inline shape 1 of active document",
    "end tell",
  ]);
  const copyTitleBefore = appleScript([
    'tell application "Microsoft Word"',
    `activate document ${JSON.stringify(reproDocumentName)}`,
    "return title of inline shape 2 of active document",
    "end tell",
  ]);
  const sourceMetadataEnvelopeBefore = appleScript([
    'tell application "Microsoft Word"',
    `activate document ${JSON.stringify(reproDocumentName)}`,
    "return alternative text of inline shape 1 of active document",
    "end tell",
  ]);
  const copyMetadataEnvelopeBefore = appleScript([
    'tell application "Microsoft Word"',
    `activate document ${JSON.stringify(reproDocumentName)}`,
    "return alternative text of inline shape 2 of active document",
    "end tell",
  ]);
  const sourceMetadataBefore = decodeFormulaMetadata(sourceMetadataEnvelopeBefore);
  const copyMetadataBefore = decodeFormulaMetadata(copyMetadataEnvelopeBefore);
  if (!sourceMetadataBefore || !copyMetadataBefore) {
    throw new Error("The pre-edit source/copy metadata did not decode");
  }
  if (sourceMetadataBefore.formulaId !== copyMetadataBefore.formulaId) {
    throw new Error("The repro document no longer contains an unforked Word copy");
  }
  if (sourceTitleBefore !== copyTitleBefore) {
    throw new Error("The repro document copy identity changed before the edit test");
  }
  const sourceLineIds = new Set<string>(
    sourceMetadataBefore.lines.map((line) => line.id),
  );

  appleScript([
    'tell application "Microsoft Word"',
    `activate document ${JSON.stringify(reproDocumentName)}`,
    "select text object of inline shape 2 of active document",
    'run VB macro macro name "VisualTeX_EditSelected"',
    "end tell",
  ], 40_000);

  const { path: sessionPath, session } = await waitForNewWordEditSession(priorSessions);
  const forkedMetadata = session.originalMetadata;
  if (!forkedMetadata) throw new Error("Copied edit session has no original metadata");
  if (!session.sourceObjectId || !String(session.sourceObjectId).startsWith("VT_E_")) {
    throw new Error(`Copied edit did not use a VT_E_ target: ${session.sourceObjectId}`);
  }
  if (session.formulaId === sourceMetadataBefore.formulaId) {
    throw new Error("Copied edit retained the pre-edit source formulaId");
  }
  if (forkedMetadata.formulaId !== session.formulaId) {
    throw new Error("Forked session metadata was not rekeyed to the copied formulaId");
  }
  for (const line of session.lines ?? []) {
    if (sourceLineIds.has(line.id)) throw new Error(`Copied edit retained source lineId ${line.id}`);
  }
  const bookmarks = appleScript([
    'tell application "Microsoft Word"',
    `activate document ${JSON.stringify(reproDocumentName)}`,
    "return name of every bookmark of active document",
    "end tell",
  ]);
  if (!bookmarks.includes(session.sourceObjectId)) {
    throw new Error(`Word does not contain the edit target bookmark ${session.sourceObjectId}`);
  }

  run("/usr/bin/pbcopy", [], 5_000, "x^2+2");
  appleScript([
    'tell application "System Events"',
    'tell process "visualtex"',
    "set frontmost to true",
    'keystroke "a" using {command down}',
    'keystroke "v" using {command down}',
    "end tell",
    "end tell",
  ], 10_000);
  await waitForSessionDirty(sessionPath, 15_000);

  appleScript([
    'tell application "System Events"',
    'tell process "visualtex"',
    "set frontmost to true",
    "keystroke return using {command down}",
    "end tell",
    "end tell",
  ], 10_000);

  await waitForSessionCompleted(sessionPath, 90_000);
  await sleep(500);
  const wordState = readWordFormulaState();
  const sourceShapeAfter = wordState.shapes.find(
    (shape) => shape.metadata?.formulaId === sourceMetadataBefore.formulaId,
  );
  const copyShapeAfter = wordState.shapes.find(
    (shape) => shape.metadata?.formulaId === session.formulaId,
  );
  if (!sourceShapeAfter || !copyShapeAfter || sourceShapeAfter.index === copyShapeAfter.index) {
    throw new Error(
      `Post-Apply Word identity mismatch: ${wordState.shapes
        .map((shape) => `#${shape.index}:${shape.metadata?.formulaId ?? "invalid"}:${shape.title}`)
        .join(" | ")}; bookmarks=${wordState.bookmarks}`,
    );
  }
  const copyMetadata = copyShapeAfter.metadata;
  if (!copyMetadata) throw new Error("Committed copied formula metadata did not decode");
  for (const line of copyMetadata.lines) {
    if (sourceLineIds.has(line.id)) throw new Error(`Committed copied formula retained source lineId ${line.id}`);
  }
  const vbaWindows = bestEffort("/usr/bin/osascript", [
    "-e", 'tell application "System Events" to tell process "Microsoft Word" to get name of every window',
  ]);
  if (vbaWindows.includes("Visual Basic for Applications")) {
    throw new Error("Word displayed a Visual Basic for Applications error after Apply");
  }

  process.stdout.write([
    "Word numbered copied formula edit→Apply E2E: PASS",
    `sessionId=${session.id}`,
    `sourceFormulaId=${sourceMetadataBefore.formulaId}`,
    `copiedFormulaId=${session.formulaId}`,
    `sourceLineIds=${[...sourceLineIds].join(",")}`,
    `copiedLineIds=${copyMetadata.lines.map((line) => line.id).join(",")}`,
    `sourceObjectId=${session.sourceObjectId}`,
    `sourceIndex=${sourceShapeAfter.index}`,
    `copyIndex=${copyShapeAfter.index}`,
    `sourceTitle=${sourceShapeAfter.title}`,
    `copyTitle=${copyShapeAfter.title}`,
    `sessionPath=${sessionPath}`,
  ].join("\n") + "\n");
} finally {
  quitWord();
  killVisualTeX();
  rmSync(installedAddinPath, { force: true });
  if (backedUpAddin && existsSync(installedBackupPath)) copyFileSync(installedBackupPath, installedAddinPath);
  rmSync(installedBackupPath, { force: true });
  rmSync(installedWordScriptPath, { force: true });
  if (backedUpWordScript && existsSync(installedWordScriptBackupPath)) copyFileSync(installedWordScriptBackupPath, installedWordScriptPath);
  rmSync(installedWordScriptBackupPath, { force: true });
  rmSync(stagedVisualTeXRoot, { recursive: true, force: true });
  bestEffort("/usr/bin/pbcopy", [], 5_000);
  spawnSync("/usr/bin/pbcopy", [], { input: clipboardBefore, encoding: "utf8", timeout: 5_000 });
  if (preexistingVisualTeX.length > 0) {
    bestEffort("/usr/bin/open", ["-gj", "-b", "com.visualtex.studio", "--args", "--office-background"], 20_000);
  }
}
