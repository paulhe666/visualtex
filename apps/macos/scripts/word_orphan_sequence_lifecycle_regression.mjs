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
import { fileURLToPath } from "node:url";

function option(name) {
  const index = process.argv.indexOf(name);
  if (index < 0) return "";
  const value = process.argv[index + 1];
  if (!value || value.startsWith("--")) {
    throw new Error(`${name} requires a value`);
  }
  return value;
}

const wordAddinPath = resolve(
  option("--word-addin") ||
    fileURLToPath(
      new URL(
        "../office/macos-offline/resources/VisualTeX.dotm",
        import.meta.url,
      ),
    ),
);
if (!wordAddinPath || !existsSync(wordAddinPath)) {
  throw new Error(`Missing compiled Word add-in: ${wordAddinPath}`);
}

const home = homedir();
const startupRoot = join(
  home,
  "Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word",
);
const installedAddinPath = join(startupRoot, "VisualTeX.dotm");
const scratchRoot = join(
  home,
  "Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch",
);
const installedBackupPath = join(
  scratchRoot,
  "VisualTeX-startup-before-orphan-sequence-lifecycle-regression.dotm",
);
const testsRoot = join(
  home,
  "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime/Tests",
);
const resultPath = join(
  testsRoot,
  "word-orphan-sequence-lifecycle-regression-result.txt",
);
const managedInterleaveResultPath = join(
  testsRoot,
  "word-managed-sequence-interleave-regression-result.txt",
);
const performanceResultPath = join(
  testsRoot,
  "word-image-to-native-fast-performance-result.txt",
);
const conversionRollbackResultPath = join(
  testsRoot,
  "word-image-native-fallback-result.txt",
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
    // Word can already be closed or invalidate the first AppleEvent while quitting.
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
      // Retry while Word finishes loading its Startup templates.
    }
    consecutiveReady = 0;
    await sleep(500);
  }
  throw new Error("Microsoft Word did not become stably automation-ready");
}

function runWordMacro(name, timeout = 120_000) {
  const appleEventTimeoutSeconds = Math.max(30, Math.ceil(timeout / 1000));
  return runAppleScript([
    `with timeout of ${appleEventTimeoutSeconds} seconds`,
    'tell application "Microsoft Word"',
    "activate",
    `run VB macro macro name ${JSON.stringify(name)}`,
    "end tell",
    "end timeout",
  ], timeout);
}

mkdirSync(startupRoot, { recursive: true });
mkdirSync(scratchRoot, { recursive: true });
mkdirSync(testsRoot, { recursive: true });
let backedUp = false;
const clipboardBefore = spawnSync("/usr/bin/pbpaste", [], {
  encoding: "utf8",
}).stdout ?? "";

try {
  const fixtureBuild = spawnSync(
    process.execPath,
    [fileURLToPath(new URL("./word_omml_regression.mjs", import.meta.url))],
    {
      encoding: "utf8",
      timeout: 120_000,
      env: {
        ...process.env,
        VISUALTEX_WORD_REGRESSION_ROOT: join(
          home,
          "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime",
        ),
      },
    },
  );
  if (fixtureBuild.status !== 0) {
    throw new Error(
      fixtureBuild.stderr.trim() ||
        fixtureBuild.stdout.trim() ||
        "Could not prepare the Word OMML regression fixtures",
    );
  }
  quitWord();
  rmSync(installedBackupPath, { force: true });
  if (existsSync(installedAddinPath)) {
    copyFileSync(installedAddinPath, installedBackupPath);
    backedUp = true;
    rmSync(installedAddinPath, { force: true });
  }
  copyFileSync(wordAddinPath, installedAddinPath);
  rmSync(resultPath, { force: true });
  rmSync(managedInterleaveResultPath, { force: true });
  rmSync(performanceResultPath, { force: true });
  rmSync(conversionRollbackResultPath, { force: true });

  await waitForWordReady();
  runAppleScript([
    'tell application "Microsoft Word"',
    'if not (exists add in "VisualTeX.dotm") then error "VisualTeX.dotm was not registered from Startup"',
    'set installed of add in "VisualTeX.dotm" to true',
    'end tell',
  ], 30_000);
  runWordMacro("VisualTeX_PerformanceNoop", 30_000);
  runWordMacro("VisualTeX_RunWordManagedSequenceInterleaveRegression", 240_000);
  if (!existsSync(managedInterleaveResultPath)) {
    throw new Error(
      `Word managed SEQ interleave regression did not write ${managedInterleaveResultPath}`,
    );
  }
  const managedInterleaveResult = readFileSync(
    managedInterleaveResultPath,
    "utf8",
  ).trim();
  if (!managedInterleaveResult.startsWith("PASS")) {
    throw new Error(
      `Word managed SEQ interleave regression failed:\n${managedInterleaveResult}`,
    );
  }
  for (const marker of [
    "managedHelpers=7",
    "validFormulaIds=7",
    "ordinaryEquationSeq=1",
    "legacyExpected=7",
    "wordNativeResult=8",
    "reconciledExpected=8",
    "orphanHelpers=0",
    "restartSwitchResult=20",
    "repeatSwitchResult=20",
    "headingSwitchResult=1",
    "switchFlowManagedResult=22",
    "identifierBoundary=PASS",
  ]) {
    if (!managedInterleaveResult.includes(marker)) {
      throw new Error(
        `Word managed SEQ interleave regression omitted ${marker}:\n${managedInterleaveResult}`,
      );
    }
  }
  process.stdout.write(
    `${managedInterleaveResult}\nWord managed SEQ interleave regression: PASS\n`,
  );

  runWordMacro("VisualTeX_RunWordOrphanSequenceLifecycleRegression");
  if (!existsSync(resultPath)) {
    throw new Error(`Word lifecycle regression did not write ${resultPath}`);
  }
  const result = readFileSync(resultPath, "utf8").trim();
  if (!result.startsWith("PASS")) {
    throw new Error(`Word lifecycle regression failed:\n${result}`);
  }
  for (const marker of [
    "directDeleteMiddle=PASS",
    "cutFormula=PASS",
    "failureImmediateRetry=PASS",
    "ordinaryWordEquationSeqPreserved=PASS",
    "orphanBeforeResult=8",
    "orphanAfterExpected=7",
    "managedHelpers=7",
  ]) {
    if (!result.includes(marker)) {
      throw new Error(`Word lifecycle regression omitted ${marker}:\n${result}`);
    }
  }
  process.stdout.write(`${result}\nWord orphan sequence lifecycle regression: PASS\n`);

  runWordMacro("VisualTeX_RunWordImageNativeFallbackRegression");
  if (!existsSync(conversionRollbackResultPath)) {
    throw new Error(`Word conversion rollback regression did not write ${conversionRollbackResultPath}`);
  }
  const conversionRollbackResult = readFileSync(
    conversionRollbackResultPath,
    "utf8",
  ).trim();
  if (!conversionRollbackResult.startsWith("PASS")) {
    throw new Error(`Word conversion rollback regression failed:\n${conversionRollbackResult}`);
  }
  for (const marker of [
    "sourceDeletionRollback=PASS",
    "complexArrayNumbering=PASS",
    "directComplexConversion=PASS",
  ]) {
    if (!conversionRollbackResult.includes(marker)) {
      throw new Error(`Word conversion rollback regression omitted ${marker}:\n${conversionRollbackResult}`);
    }
  }
  process.stdout.write(`${conversionRollbackResult}\nWord numbered conversion rollback/retry regression: PASS\n`);

  runWordMacro("VisualTeX_RunImageToNativeFastPerformanceRegression");
  if (!existsSync(performanceResultPath)) {
    throw new Error(`Word performance regression did not write ${performanceResultPath}`);
  }
  const performanceResult = readFileSync(performanceResultPath, "utf8").trim();
  if (!performanceResult.startsWith("PASS")) {
    throw new Error(`Word image-to-OMML performance regression failed:\n${performanceResult}`);
  }
  if (!performanceResult.includes("conversionPath=direct-cache-insert")) {
    throw new Error(`Word image-to-OMML conversion did not use direct cache insertion:\n${performanceResult}`);
  }
  process.stdout.write(`${performanceResult}\nWord single image-to-OMML performance regression: PASS\n`);
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
