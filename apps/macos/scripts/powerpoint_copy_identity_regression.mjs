import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, rmSync } from "node:fs";
import { homedir } from "node:os";
import { basename, join, resolve } from "node:path";

function option(name) {
  const index = process.argv.indexOf(name);
  if (index < 0) return "";
  const value = process.argv[index + 1];
  if (!value || value.startsWith("--")) {
    throw new Error(`${name} requires a value`);
  }
  return value;
}

const presentationPath = resolve(option("--presentation"));
if (!presentationPath || !existsSync(presentationPath)) {
  throw new Error(`Missing compiled PowerPoint presentation: ${presentationPath}`);
}
const presentationName = basename(presentationPath);
const home = homedir();
const resultPath = join(
  home,
  "Library/Application Scripts/com.microsoft.Powerpoint/VisualTeXRuntime/Tests/powerpoint-copy-identity-regression-result.txt",
);
const sleep = (ms) => new Promise((resolveSleep) => setTimeout(resolveSleep, ms));

function run(program, args, options = {}) {
  const result = spawnSync(program, args, {
    encoding: "utf8",
    timeout: options.timeout ?? 30_000,
  });
  if (result.status !== 0) {
    throw new Error(
      result.stderr.trim() || result.stdout.trim() || `${program} failed`,
    );
  }
  return result.stdout.trim();
}

function bestEffort(program, args, options = {}) {
  try {
    return run(program, args, options);
  } catch {
    return "";
  }
}

function runAppleScript(lines, timeout = 30_000) {
  const args = [];
  for (const line of lines) args.push("-e", line);
  return run("/usr/bin/osascript", args, { timeout });
}

function quitPowerPoint() {
  bestEffort("/usr/bin/osascript", [
    "-e",
    'tell application "Microsoft PowerPoint" to quit saving no',
  ], { timeout: 20_000 });
  bestEffort("/usr/bin/killall", ["Microsoft PowerPoint"], {
    timeout: 10_000,
  });
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
  while (Date.now() < deadline) {
    enableMacrosIfPrompted();
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
    } catch {
      // Retry while PowerPoint opens the PPTM and finishes macro security UI.
    }
    await sleep(500);
  }
  throw new Error("PowerPoint did not become UI-ready");
}

mkdirSync(
  join(
    home,
    "Library/Application Scripts/com.microsoft.Powerpoint/VisualTeXRuntime/Tests",
  ),
  { recursive: true },
);

try {
  quitPowerPoint();
  await sleep(2_500);
  rmSync(resultPath, { force: true });
  run("/usr/bin/open", ["-b", "com.microsoft.Powerpoint", presentationPath], {
    timeout: 10_000,
  });
  enableMacrosIfPrompted();
  await sleep(800);
  let presentationReady = false;
  const presentationDeadline = Date.now() + 30_000;
  while (Date.now() < presentationDeadline) {
    enableMacrosIfPrompted();
    try {
      const state = runAppleScript([
        'tell application "Microsoft PowerPoint"',
        "activate",
        `if exists presentation ${JSON.stringify(presentationName)} then return "READY"`,
        'return "NOT_READY"',
        "end tell",
      ], 5_000);
      if (state === "READY") {
        presentationReady = true;
        break;
      }
    } catch {
      // Macro/security UI can temporarily block PowerPoint AppleEvents.
    }
    await sleep(500);
  }
  if (!presentationReady) {
    throw new Error(
      `PowerPoint did not register the regression presentation ${presentationName}`,
    );
  }
  enableMacrosIfPrompted();
  await waitForPowerPointUi(10_000);
  runAppleScript([
    'tell application "Microsoft PowerPoint"',
    'run VB macro macro name "VisualTeX_RunPowerPointCopyIdentityRegression" list of parameters {}',
    "end tell",
  ], 90_000);
  if (!existsSync(resultPath)) {
    throw new Error(`PowerPoint copy-identity regression did not write ${resultPath}`);
  }
  const result = readFileSync(resultPath, "utf8").trim();
  if (!result.startsWith("PASS")) {
    throw new Error(`PowerPoint copy-identity regression failed:\n${result}`);
  }
  process.stdout.write(`${result}\nPowerPoint copy identity regression: PASS\n`);
} finally {
  quitPowerPoint();
}
