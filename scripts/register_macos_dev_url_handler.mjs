import { execFileSync } from "node:child_process";
import {
  existsSync,
  lstatSync,
  mkdirSync,
  readlinkSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from "node:fs";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const debugExecutable = join(repositoryRoot, "src-tauri", "target", "debug", "visualtex");
const devApp = join(
  repositoryRoot,
  "src-tauri",
  "target",
  "debug",
  "bundle",
  "macos",
  "VisualTeX Dev.app",
);
const appExecutable = join(devApp, "Contents", "MacOS", "visualtex");
const infoPlist = join(devApp, "Contents", "Info.plist");
const launchServices =
  "/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister";

function xmlEscape(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function ensureExecutableLink() {
  mkdirSync(dirname(appExecutable), { recursive: true });
  if (existsSync(appExecutable) || lstatExists(appExecutable)) {
    const currentTarget = lstatSync(appExecutable).isSymbolicLink()
      ? resolve(dirname(appExecutable), readlinkSync(appExecutable))
      : null;
    if (currentTarget === debugExecutable) return;
    rmSync(appExecutable, { force: true });
  }
  symlinkSync(relative(dirname(appExecutable), debugExecutable), appExecutable);
}

function lstatExists(path) {
  try {
    lstatSync(path);
    return true;
  } catch {
    return false;
  }
}

export function registerMacosDevUrlHandler() {
  if (process.platform !== "darwin") return null;
  if (!existsSync(debugExecutable)) {
    throw new Error(
      `VisualTeX debug executable is missing: ${debugExecutable}. Start \"tauri dev\" first.`,
    );
  }

  ensureExecutableLink();
  writeFileSync(
    infoPlist,
    `<?xml version="1.0" encoding="UTF-8"?>\n` +
      `<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n` +
      `<plist version="1.0">\n<dict>\n` +
      `  <key>CFBundleDevelopmentRegion</key><string>English</string>\n` +
      `  <key>CFBundleDisplayName</key><string>VisualTeX Dev</string>\n` +
      `  <key>CFBundleExecutable</key><string>visualtex</string>\n` +
      `  <key>CFBundleIdentifier</key><string>com.visualtex.studio</string>\n` +
      `  <key>CFBundleInfoDictionaryVersion</key><string>6.0</string>\n` +
      `  <key>CFBundleName</key><string>VisualTeX Dev</string>\n` +
      `  <key>CFBundlePackageType</key><string>APPL</string>\n` +
      `  <key>CFBundleShortVersionString</key><string>1.1.0-dev</string>\n` +
      `  <key>CFBundleVersion</key><string>1</string>\n` +
      `  <key>CFBundleURLTypes</key>\n` +
      `  <array><dict>\n` +
      `    <key>CFBundleTypeRole</key><string>Editor</string>\n` +
      `    <key>CFBundleURLName</key><string>${xmlEscape("VisualTeX Offline Office Session")}</string>\n` +
      `    <key>CFBundleURLSchemes</key><array><string>visualtex</string></array>\n` +
      `  </dict></array>\n` +
      `</dict>\n</plist>\n`,
    "utf8",
  );

  execFileSync("/usr/bin/plutil", ["-lint", infoPlist], { stdio: "pipe" });
  execFileSync(launchServices, ["-f", devApp], { stdio: "pipe" });
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
