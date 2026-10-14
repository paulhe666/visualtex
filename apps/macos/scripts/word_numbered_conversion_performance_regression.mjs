import { spawnSync } from "node:child_process";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readdirSync,
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
      new URL("../office/macos-offline/resources/VisualTeX.dotm", import.meta.url),
    ),
);
if (!existsSync(wordAddinPath)) {
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
  "VisualTeX-startup-before-numbered-conversion-performance.dotm",
);
const testsRoot = join(
  home,
  "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime/Tests",
);
const runtimeRoot = join(
  home,
  "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime",
);
const formulaId = "0274750e-c676-4f7b-8c04-a9d9ef7336a0";
const resultPath = join(
  testsRoot,
  "word-numbered-image-omml-fast-performance-result.txt",
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
    // Word can already be closed while the regression restores Startup.
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
      // Retry while Word loads its Startup template.
    }
    consecutiveReady = 0;
    await sleep(500);
  }
  throw new Error("Microsoft Word did not become automation-ready");
}

function runWordMacro(name, timeout = 180_000) {
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

try {
  const nativeDocumentsRoot = join(runtimeRoot, "NativeDocuments");
  const imageDocumentsRoot = join(runtimeRoot, "ImageDocuments");
  const nativeFixturePath = join(nativeDocumentsRoot, `${formulaId}.docx`);
  const numberedNativeSourcePath = join(
    nativeDocumentsRoot,
    "22222222-2222-4222-8222-222222222222.docx",
  );
  const imageFixturePaths = ["docx", "svg", "png"].map((extension) =>
    join(imageDocumentsRoot, `${formulaId}.${extension}`),
  );
  const ommlFixturePath = join(testsRoot, "word-native-regression-omml.txt");
  const fixturesReady =
    existsSync(numberedNativeSourcePath) &&
    existsSync(ommlFixturePath) &&
    imageFixturePaths.every((path) => existsSync(path));
  if (!fixturesReady) {
    const fixtureBuild = spawnSync(
      process.execPath,
      [fileURLToPath(new URL("./word_omml_regression.mjs", import.meta.url))],
      {
        encoding: "utf8",
        timeout: 120_000,
        env: {
          ...process.env,
          VISUALTEX_WORD_REGRESSION_ROOT: runtimeRoot,
        },
      },
    );
    if (fixtureBuild.status !== 0) {
      throw new Error(
        fixtureBuild.stderr.trim() ||
          fixtureBuild.stdout.trim() ||
          "Could not prepare the Word conversion fixtures",
      );
    }
  }

  mkdirSync(imageDocumentsRoot, { recursive: true });
  copyFileSync(numberedNativeSourcePath, nativeFixturePath);
  if (!imageFixturePaths.every((path) => existsSync(path))) {
    const reusableImageStem = readdirSync(imageDocumentsRoot)
      .filter((name) => name.endsWith(".docx"))
      .map((name) => name.slice(0, -5))
      .find(
        (stem) =>
          stem !== formulaId &&
          existsSync(join(imageDocumentsRoot, `${stem}.svg`)) &&
          existsSync(join(imageDocumentsRoot, `${stem}.png`)),
      );
    if (!reusableImageStem) {
      throw new Error("No complete cached Word image fixture is available");
    }
    for (const extension of ["docx", "svg", "png"]) {
      copyFileSync(
        join(imageDocumentsRoot, `${reusableImageStem}.${extension}`),
        join(imageDocumentsRoot, `${formulaId}.${extension}`),
      );
    }
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

  await waitForWordReady();
  runAppleScript([
    'tell application "Microsoft Word"',
    'if not (exists add in "VisualTeX.dotm") then error "VisualTeX.dotm was not registered from Startup"',
    'set installed of add in "VisualTeX.dotm" to true',
    "end tell",
  ]);
  runWordMacro("VisualTeX_RunNumberedImageOmmlFastPerformanceRegression");
  if (!existsSync(resultPath)) {
    throw new Error(`Word regression did not write ${resultPath}`);
  }
  const result = readFileSync(resultPath, "utf8").trim();
  if (!result.startsWith("PASS")) {
    throw new Error(`Word numbered conversion regression failed:\n${result}`);
  }
  const metric = (name) => {
    const match = result.match(new RegExp(`^${name}=(\\d+)$`, "m"));
    if (!match) throw new Error(`Word numbered conversion omitted ${name}:\n${result}`);
    return Number.parseInt(match[1], 10);
  };
  const imageToNativeMs = metric("imageToNativeMs");
  const nativeToImageMs = metric("nativeToImageMs");
  const imageNumberMs = metric("imageNumberMs");
  if (imageToNativeMs > 750 || nativeToImageMs > 750 || imageNumberMs > 500) {
    throw new Error(
      `Word numbered conversion exceeded its interactive budget:\n${result}`,
    );
  }
  process.stdout.write(`${result}\nWord numbered image/OMML performance: PASS\n`);
} finally {
  quitWord();
  rmSync(installedAddinPath, { force: true });
  if (backedUp && existsSync(installedBackupPath)) {
    copyFileSync(installedBackupPath, installedAddinPath);
  }
  rmSync(installedBackupPath, { force: true });
}
