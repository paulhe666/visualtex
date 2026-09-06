import { spawnSync } from "node:child_process";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  rmSync,
} from "node:fs";
import { homedir } from "node:os";
import { join, resolve } from "node:path";

function option(name) {
  const index = process.argv.indexOf(name);
  if (index < 0) return "";
  const value = process.argv[index + 1];
  if (!value || value.startsWith("--")) {
    throw new Error(`${name} requires a value`);
  }
  return value;
}

const wordAddinPath = resolve(option("--word-addin"));
if (!wordAddinPath || !existsSync(wordAddinPath)) {
  throw new Error(`Missing compiled Word add-in: ${wordAddinPath}`);
}

const home = homedir();
const startupRoot = join(
  home,
  "Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word",
);
const installedAddinPath = join(startupRoot, "VisualTeX.dotm");
const installedBackupPath = join(
  home,
  "Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch/VisualTeX-startup-before-copy-identity-regression.dotm",
);
const resultPath = join(
  home,
  "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime/Tests/word-copy-identity-regression-result.txt",
);
const nativeProbeResultPath = join(
  home,
  "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime/Tests/word-native-copy-identity-probe-result.txt",
);
const numberedResultPath = join(
  home,
  "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime/Tests/word-numbered-copy-identity-regression-result.txt",
);

const sleep = (ms) => new Promise((resolveSleep) => setTimeout(resolveSleep, ms));

function runAppleScript(lines, timeout = 30_000) {
  const args = [];
  for (const line of lines) args.push("-e", line);
  const result = spawnSync("/usr/bin/osascript", args, {
    encoding: "utf8",
    timeout,
  });
  if (result.status !== 0) {
    throw new Error(
      result.stderr.trim() || result.stdout.trim() || "AppleScript failed",
    );
  }
  return result.stdout.trim();
}

function quitWord() {
  try {
    runAppleScript(['tell application "Microsoft Word" to quit saving no'], 20_000);
  } catch {
    // Word can already be closed or can invalidate the first AppleEvent while quitting.
  }
  spawnSync("/usr/bin/killall", ["Microsoft Word"], {
    encoding: "utf8",
    timeout: 10_000,
  });
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
      const version = runAppleScript([
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
      // Retry while Word finishes loading Startup templates.
    }
    consecutiveReady = 0;
    await sleep(500);
  }
  throw new Error("Microsoft Word did not become stably automation-ready");
}

function runWordMacro(name, timeout = 60_000) {
  return runAppleScript([
    'tell application "Microsoft Word"',
    "activate",
    `run VB macro macro name ${JSON.stringify(name)}`,
    "end tell",
  ], timeout);
}

mkdirSync(startupRoot, { recursive: true });
mkdirSync(
  join(home, "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime/Tests"),
  { recursive: true },
);
let backedUp = false;
const clipboardBefore = spawnSync("/usr/bin/pbpaste", [], {
  encoding: "utf8",
}).stdout ?? "";

try {
  quitWord();
  rmSync(installedBackupPath, { force: true });
  if (existsSync(installedAddinPath)) {
    copyFileSync(installedAddinPath, installedBackupPath);
    backedUp = true;
    rmSync(installedAddinPath, { force: true });
  }
  copyFileSync(wordAddinPath, installedAddinPath);
  rmSync(resultPath, { force: true });
  rmSync(nativeProbeResultPath, { force: true });
  rmSync(numberedResultPath, { force: true });

  await waitForWordReady();
  runAppleScript([
    'tell application "Microsoft Word"',
    'if not (exists add in "VisualTeX.dotm") then error "VisualTeX.dotm was not registered from Startup"',
    'set installed of add in "VisualTeX.dotm" to true',
    'end tell',
  ], 30_000);
  runWordMacro("VisualTeX_PerformanceNoop", 30_000);
  runWordMacro("VisualTeX_RunWordCopyIdentityRegression", 90_000);
  if (!existsSync(resultPath)) {
    throw new Error(`Word copy-identity regression did not write ${resultPath}`);
  }
  const result = readFileSync(resultPath, "utf8").trim();
  if (!result.startsWith("PASS")) {
    throw new Error(`Word copy-identity regression failed:\n${result}`);
  }
  process.stdout.write(`${result}\nWord image copy identity regression: PASS\n`);

  runWordMacro("VisualTeX_RunWordNumberedCopyIdentityRegression", 90_000);
  if (!existsSync(numberedResultPath)) {
    throw new Error(`Word numbered copy-identity regression did not write ${numberedResultPath}`);
  }
  const numberedResult = readFileSync(numberedResultPath, "utf8").trim();
  if (!numberedResult.startsWith("PASS")) {
    throw new Error(`Word numbered copy-identity regression failed:\n${numberedResult}`);
  }
  process.stdout.write(`${numberedResult}\nWord numbered image copy identity regression: PASS\n`);

  runWordMacro("VisualTeX_RunWordNativeCopyIdentityProbe", 60_000);
  if (!existsSync(nativeProbeResultPath)) {
    throw new Error(`Word native-copy probe did not write ${nativeProbeResultPath}`);
  }
  const nativeProbeResult = readFileSync(nativeProbeResultPath, "utf8").trim();
  if (!nativeProbeResult.startsWith("PASS")) {
    throw new Error(`Word native-copy probe failed:\n${nativeProbeResult}`);
  }
  process.stdout.write(`${nativeProbeResult}\nWord native copy identity probe: PASS\n`);
} finally {
  quitWord();
  rmSync(installedAddinPath, { force: true });
  if (backedUp && existsSync(installedBackupPath)) {
    copyFileSync(installedBackupPath, installedAddinPath);
  }
  rmSync(installedBackupPath, { force: true });
  spawnSync("/usr/bin/pbcopy", [], {
    input: clipboardBefore,
    encoding: "utf8",
  });
}
