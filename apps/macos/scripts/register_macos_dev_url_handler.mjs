import { execFileSync } from "node:child_process";
import {
  existsSync,
  mkdirSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const debugExecutable = join(repositoryRoot, "src-tauri", "target", "debug", "visualtex");
const devBundleRoot = join(
  repositoryRoot,
  "src-tauri",
  "target",
  "debug",
  "bundle",
  "macos",
);
const devApp = join(devBundleRoot, "VisualTeX Dev URL Handler.app");
const legacyDevApp = join(devBundleRoot, "VisualTeX Dev.app");
const launcherSource = join(devBundleRoot, "VisualTeXDevURLHandler.applescript");
const infoPlist = join(devApp, "Contents", "Info.plist");
const launchServices =
  "/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister";

function appleScriptString(value) {
  return value.replaceAll("\\", "\\\\").replaceAll('"', '\\"');
}

function launcherAppleScript() {
  const executable = appleScriptString(debugExecutable);
  return `property debugExecutable : "${executable}"
property devServer : "http://localhost:1420/"
property officeURLPrefix : "visualtex://office/open?session="

on run
    my launchVisualTeX("")
end run

on open location visualTeXURL
    my launchVisualTeX(visualTeXURL as text)
end open location

on launchVisualTeX(visualTeXURL)
    try
        do shell script "/usr/bin/curl --silent --fail --max-time 1 " & quoted form of devServer & " >/dev/null 2>&1"
    on error
        display alert "VisualTeX 开发服务未运行" message "请先在项目目录执行 npm run tauri:dev，并保持终端窗口运行。" as critical
        return
    end try

    if visualTeXURL is not "" then
        if visualTeXURL does not start with officeURLPrefix then
            display alert "VisualTeX Office 链接无效" message "开发启动器只接受固定的 VisualTeX Office Session 链接。" as critical
            return
        end if
    end if

    set commandText to quoted form of debugExecutable
    if visualTeXURL is not "" then
        set commandText to commandText & " " & quoted form of visualTeXURL
    end if
    do shell script commandText & " >/tmp/visualtex-dev-url.log 2>&1 &"
end launchVisualTeX
`;
}

function updateInfoPlist() {
  const urlTypes = JSON.stringify([
    {
      CFBundleTypeRole: "Editor",
      CFBundleURLName: "VisualTeX Offline Office Session",
      CFBundleURLSchemes: ["visualtex"],
    },
  ]);
  for (const [key, type, value] of [
    ["CFBundleIdentifier", "-string", "com.visualtex.studio.dev-url-handler"],
    ["CFBundleDisplayName", "-string", "VisualTeX Dev URL Handler"],
    ["CFBundleName", "-string", "VisualTeX Dev URL Handler"],
    ["CFBundleShortVersionString", "-string", "1.2.2-dev"],
    ["CFBundleVersion", "-string", "1"],
    ["LSUIElement", "-bool", "YES"],
  ]) {
    execFileSync("/usr/bin/plutil", ["-replace", key, type, value, infoPlist], {
      stdio: "pipe",
    });
  }
  execFileSync(
    "/usr/bin/plutil",
    ["-replace", "CFBundleURLTypes", "-json", urlTypes, infoPlist],
    { stdio: "pipe" },
  );
}

export function registerMacosDevUrlHandler() {
  if (process.platform !== "darwin") return null;
  if (!existsSync(debugExecutable)) {
    throw new Error(
      `VisualTeX debug executable is missing: ${debugExecutable}. Start \"tauri dev\" first.`,
    );
  }

  mkdirSync(devBundleRoot, { recursive: true });
  if (existsSync(legacyDevApp)) {
    try {
      execFileSync(launchServices, ["-u", legacyDevApp], { stdio: "pipe" });
    } catch {
      // The legacy shell-based handler may already be absent from LaunchServices.
    }
    rmSync(legacyDevApp, { recursive: true, force: true });
  }
  rmSync(devApp, { recursive: true, force: true });
  writeFileSync(launcherSource, launcherAppleScript(), "utf8");
  execFileSync("/usr/bin/osacompile", ["-o", devApp, launcherSource], {
    stdio: "pipe",
  });
  updateInfoPlist();
  execFileSync("/usr/bin/plutil", ["-lint", infoPlist], { stdio: "pipe" });
  execFileSync("/usr/bin/codesign", ["--force", "--deep", "--sign", "-", devApp], {
    stdio: "pipe",
  });
  execFileSync(launchServices, ["-f", devApp], { stdio: "pipe" });
  execFileSync(
    "/usr/bin/osascript",
    [
      "-l",
      "JavaScript",
      "-e",
      'ObjC.import("CoreServices"); $.LSSetDefaultHandlerForURLScheme($("visualtex"), $("com.visualtex.studio.dev-url-handler"));',
    ],
    { stdio: "pipe" },
  );
  return devApp;
}

export function unregisterMacosDevUrlHandler() {
  if (process.platform !== "darwin" || !existsSync(devApp)) return;
  try {
    execFileSync(launchServices, ["-u", devApp], { stdio: "pipe" });
  } catch {
    // LaunchServices cleanup is best-effort during development shutdown.
  }
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    const registered = registerMacosDevUrlHandler();
    if (registered) {
      process.stdout.write(`Registered visualtex:// for ${registered}\n`);
    }
  } catch (error) {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
    process.exitCode = 1;
  }
}
