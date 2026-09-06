import { spawn, spawnSync } from "node:child_process";
import { existsSync, readFileSync, readdirSync, rmSync } from "node:fs";
import { homedir, tmpdir } from "node:os";
import { join, resolve } from "node:path";

const appRoot = resolve(new URL("..", import.meta.url).pathname);
const sourceAppBundle = resolve(
  process.argv.includes("--app")
    ? process.argv[process.argv.indexOf("--app") + 1]
    : join(appRoot, "src-tauri/target/release/bundle/macos/VisualTeX.app"),
);
const sourceAppExecutable = join(sourceAppBundle, "Contents/MacOS/visualtex");
if (!existsSync(sourceAppExecutable)) {
  throw new Error(`Missing VisualTeX executable: ${sourceAppExecutable}`);
}
const installedApp = "/Applications/VisualTeX.app";
const installedAppBackup = join(
  tmpdir(),
  `VisualTeX-PowerPointFastOpenE2E-backup-${process.pid}.app`,
);
const appBundle = installedApp;
const appExecutable = join(appBundle, "Contents/MacOS/visualtex");

const home = homedir();
const appSessionsRoot = join(home, "Library/Application Support/com.visualtex.studio/office/sessions");
const runtimeSessionsRoot = join(
  home,
  "Library/Application Scripts/com.microsoft.Powerpoint/VisualTeXRuntime/OfficeSessions",
);
const readyPath = join(
  home,
  "Library/Containers/com.microsoft.Powerpoint/Data/Library/Application Support/VisualTeX/FastOpen/powerpoint/resident-ready",
);
const sleep = (ms) => new Promise((resolveSleep) => setTimeout(resolveSleep, ms));

function run(program, args, timeout = 30_000) {
  const result = spawnSync(program, args, {
    encoding: "utf8",
    timeout,
    maxBuffer: 32 * 1024 * 1024,
  });
  if (result.status !== 0) {
    throw new Error(result.stderr.trim() || result.stdout.trim() || `${program} failed`);
  }
  return result.stdout.trim();
}

function bestEffort(program, args, timeout = 30_000) {
  try {
    return run(program, args, timeout);
  } catch {
    return "";
  }
}

function appleScript(lines, timeout = 30_000) {
  return run(
    "/usr/bin/osascript",
    lines.flatMap((line) => ["-e", line]),
    timeout,
  );
}

function sessionIds() {
  if (!existsSync(appSessionsRoot)) return new Set();
  return new Set(
    readdirSync(appSessionsRoot, { withFileTypes: true })
      .filter((entry) => entry.isDirectory())
      .map((entry) => entry.name),
  );
}

async function waitForHeartbeat(pid, timeoutMs = 15_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (existsSync(readyPath)) {
      const lines = readFileSync(readyPath, "utf8").trim().split(/\r?\n/);
      if (
        lines[0] === "visualtex-fast-open-ready-v2" &&
        lines[2] === String(pid) &&
        lines[3] === appExecutable
      ) {
        return;
      }
    }
    await sleep(150);
  }
  throw new Error("VisualTeX did not publish the expected PowerPoint resident heartbeat");
}

async function waitForPowerPointSession(previousIds, presentationName, timeoutMs = 20_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (existsSync(appSessionsRoot)) {
      for (const entry of readdirSync(appSessionsRoot, { withFileTypes: true })) {
        if (!entry.isDirectory() || previousIds.has(entry.name)) continue;
        const sessionPath = join(appSessionsRoot, entry.name, "session.json");
        if (!existsSync(sessionPath)) continue;
        try {
          const session = JSON.parse(readFileSync(sessionPath, "utf8"));
          if (
            session.host === "powerpoint" &&
            session.mode === "create" &&
            new Set([
              presentationName,
              `visualtex-ppt-native-presentation:${presentationName}`,
            ]).has(session.sourceDocumentId)
          ) {
            return { session, sessionPath };
          }
        } catch {
          // Retry until the atomic write has finished.
        }
      }
    }
    await sleep(100);
  }
  throw new Error("PowerPoint fast-open request was claimed but no create Session was imported");
}

async function waitForEditorReady(sessionId, timeoutMs = 20_000) {
  const path = join(runtimeSessionsRoot, sessionId, "editor-ready.json");
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (existsSync(path)) {
      try {
        const ready = JSON.parse(readFileSync(path, "utf8"));
        if (
          ready.host === "powerpoint" &&
          ready.sessionId === sessionId &&
          ready.windowVisible === true &&
          ready.windowFocused === true &&
          ready.appActive === true
        ) {
          return { ready, path };
        }
      } catch {
        // Retry until the JSON write is complete.
      }
    }
    await sleep(100);
  }
  throw new Error(`PowerPoint editor never became visible/focused for Session ${sessionId}`);
}

function createTemporaryPresentation() {
  return appleScript([
    'tell application "Microsoft PowerPoint"',
    "activate",
    "set testPresentation to make new presentation",
    "if (count of slides of testPresentation) is 0 then make new slide at end of testPresentation",
    "return name of testPresentation as text",
    "end tell",
  ]);
}

function closePresentationWithoutSaving(name) {
  bestEffort("/usr/bin/osascript", [
    "-e", 'tell application "Microsoft PowerPoint"',
    "-e", `if exists presentation ${JSON.stringify(name)} then close presentation ${JSON.stringify(name)} saving no`,
    "-e", "end tell",
  ]);
}

let testPresentationName = "";
let testSessionId = "";
let visualTeX;
const hadInstalledVisualTeX = existsSync(installedApp);
let installedAppBackedUp = false;

try {
  bestEffort("/usr/bin/killall", ["visualtex"], 10_000);
  await sleep(700);

  rmSync(installedAppBackup, { recursive: true, force: true });
  if (hadInstalledVisualTeX) {
    run("/usr/bin/ditto", [installedApp, installedAppBackup], 120_000);
    installedAppBackedUp = true;
  }
  rmSync(installedApp, { recursive: true, force: true });
  run("/usr/bin/ditto", [sourceAppBundle, installedApp], 120_000);
  if (!existsSync(appExecutable)) {
    throw new Error(`Installed VisualTeX executable is missing: ${appExecutable}`);
  }

  visualTeX = spawn(appExecutable, ["--office-background"], {
    detached: true,
    stdio: "ignore",
  });
  visualTeX.unref();
  if (!visualTeX.pid) throw new Error("VisualTeX resident failed to start");
  await waitForHeartbeat(visualTeX.pid);

  testPresentationName = createTemporaryPresentation();
  const priorSessions = sessionIds();

  appleScript([
    'tell application "Microsoft PowerPoint"',
    `activate presentation ${JSON.stringify(testPresentationName)}`,
    'run VB macro macro name "VisualTeX_NewFormula" list of parameters {}',
    "end tell",
  ], 30_000);

  const { session, sessionPath } = await waitForPowerPointSession(
    priorSessions,
    testPresentationName,
  );
  testSessionId = session.id;
  if (!session.formulaId) throw new Error("PowerPoint create Session has no formulaId");
  const { ready, path: readyFile } = await waitForEditorReady(session.id);

  const shapeCount = Number(
    appleScript([
      'tell application "Microsoft PowerPoint"',
      `activate presentation ${JSON.stringify(testPresentationName)}`,
      "return count of shapes of slide 1 of active presentation",
      "end tell",
    ]),
  );
  if (!Number.isFinite(shapeCount) || shapeCount < 1) {
    throw new Error("PowerPoint create did not leave its pending formula placeholder");
  }

  process.stdout.write(
    [
      "PowerPoint fast-open create E2E: PASS",
      `presentation=${testPresentationName}`,
      `sessionId=${session.id}`,
      `formulaId=${session.formulaId}`,
      `sessionPath=${sessionPath}`,
      `readyPath=${readyFile}`,
      `showFocusMs=${ready.showFocusMs}`,
      `windowVisible=${ready.windowVisible}`,
      `windowFocused=${ready.windowFocused}`,
      `appActive=${ready.appActive}`,
    ].join("\n") + "\n",
  );
} finally {
  if (testPresentationName) closePresentationWithoutSaving(testPresentationName);
  bestEffort("/usr/bin/killall", ["visualtex"], 10_000);
  if (testSessionId) {
    rmSync(join(appSessionsRoot, testSessionId), { recursive: true, force: true });
    rmSync(join(runtimeSessionsRoot, testSessionId), { recursive: true, force: true });
  }
  rmSync(installedApp, { recursive: true, force: true });
  if (installedAppBackedUp && existsSync(installedAppBackup)) {
    bestEffort("/usr/bin/ditto", [installedAppBackup, installedApp], 120_000);
  }
  rmSync(installedAppBackup, { recursive: true, force: true });
  if (hadInstalledVisualTeX && existsSync(installedApp)) {
    bestEffort("/usr/bin/open", ["-gj", installedApp, "--args", "--office-background"], 20_000);
  }
}
