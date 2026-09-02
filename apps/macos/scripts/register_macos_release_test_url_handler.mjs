import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const releaseApp = join(
  repositoryRoot,
  "src-tauri/target/release/bundle/macos/VisualTeX.app",
);
const releaseExecutable = join(releaseApp, "Contents/MacOS/visualtex");
const bundleRoot = join(repositoryRoot, "src-tauri/target/release/bundle/macos");
const handlerApp = join(bundleRoot, "VisualTeX Performance URL Handler.app");
const launcherSource = join(bundleRoot, "VisualTeXPerformanceURLHandler.applescript");
const infoPlist = join(handlerApp, "Contents/Info.plist");
const launchServices =
  "/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister";
const handlerBundleId = "com.visualtex.studio.performance-url-handler";
const productionBundleId = "com.visualtex.studio";

function appleScriptString(value) {
  return value.replaceAll("\\", "\\\\").replaceAll('"', '\\"');
}

function run(program, args) {
  return execFileSync(program, args, {
    encoding: "utf8",
    stdio: ["ignore", "pipe", "pipe"],
    timeout: 30_000,
  });
}

function launcherAppleScript() {
  const executable = appleScriptString(releaseExecutable);
  return `property releaseExecutable : "${executable}"
property officeURLPrefix : "visualtex://office/open?session="

on run
    my launchVisualTeX("")
end run

on open location visualTeXURL
    my launchVisualTeX(visualTeXURL as text)
end open location

on launchVisualTeX(visualTeXURL)
    if visualTeXURL is not "" then
        if visualTeXURL does not start with officeURLPrefix then
            display alert "VisualTeX Office URL invalid" as critical
            return
        end if
    end if
    set commandText to quoted form of releaseExecutable
    if visualTeXURL is not "" then
        set commandText to commandText & " " & quoted form of visualTeXURL
    end if
    do shell script commandText & " >/tmp/visualtex-release-test-url.log 2>&1 &"
end launchVisualTeX
`;
}

function registerHandler() {
  if (!existsSync(releaseExecutable)) {
    throw new Error(`Release VisualTeX executable is missing: ${releaseExecutable}`);
  }
  mkdirSync(bundleRoot, { recursive: true });
  rmSync(handlerApp, { recursive: true, force: true });
  writeFileSync(launcherSource, launcherAppleScript(), "utf8");
  run("/usr/bin/osacompile", ["-o", handlerApp, launcherSource]);
  const urlTypes = JSON.stringify([
    {
      CFBundleTypeRole: "Editor",
      CFBundleURLName: "VisualTeX Performance Office Session",
      CFBundleURLSchemes: ["visualtex"],
    },
  ]);
  for (const [key, type, value] of [
    ["CFBundleIdentifier", "-string", handlerBundleId],
    ["CFBundleDisplayName", "-string", "VisualTeX Performance URL Handler"],
    ["CFBundleName", "-string", "VisualTeX Performance URL Handler"],
    ["CFBundleShortVersionString", "-string", "1.2.5-test"],
    ["CFBundleVersion", "-string", "1"],
    ["LSUIElement", "-bool", "YES"],
  ]) {
    run("/usr/bin/plutil", ["-replace", key, type, value, infoPlist]);
  }
  run("/usr/bin/plutil", [
    "-replace",
    "CFBundleURLTypes",
    "-json",
    urlTypes,
    infoPlist,
  ]);
  run("/usr/bin/codesign", ["--force", "--deep", "--sign", "-", handlerApp]);
  run(launchServices, ["-f", handlerApp]);
  run("/usr/bin/osascript", [
    "-l",
    "JavaScript",
    "-e",
    `ObjC.import("CoreServices"); $.LSSetDefaultHandlerForURLScheme($("visualtex"), $("${handlerBundleId}"));`,
  ]);
  process.stdout.write(`Registered visualtex:// to ${handlerApp}\n`);
}

function restoreProduction() {
  if (existsSync(handlerApp)) {
    try {
      run(launchServices, ["-u", handlerApp]);
    } catch {
      // Best-effort unregister.
    }
  }
  for (const productionApp of [
    "/Applications/VisualTeX.app",
    join(process.env.HOME ?? "", "Applications", "VisualTeX.app"),
  ]) {
    if (!productionApp || !existsSync(productionApp)) continue;
    try {
      run(launchServices, ["-f", productionApp]);
      break;
    } catch {
      // Continue.
    }
  }
  run("/usr/bin/osascript", [
    "-l",
    "JavaScript",
    "-e",
    `ObjC.import("CoreServices"); $.LSSetDefaultHandlerForURLScheme($("visualtex"), $("${productionBundleId}"));`,
  ]);
  process.stdout.write("Restored production visualtex:// handler\n");
}

if (process.argv.includes("--restore")) {
  restoreProduction();
} else {
  registerHandler();
}
