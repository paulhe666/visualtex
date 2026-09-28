import { spawnSync } from "node:child_process";
import {
  windowsPowerShellEnvironment,
  windowsPowerShellPath,
} from "./windows_powershell.mjs";

function run(command, args, env = process.env) {
  const isWindowsCmd =
    process.platform === "win32" && command.toLowerCase().endsWith(".cmd");
  const executable = isWindowsCmd ? (process.env.ComSpec ?? "cmd.exe") : command;
  const executableArgs = isWindowsCmd
    ? ["/d", "/s", "/c", command, ...args]
    : args;
  const result = spawnSync(executable, executableArgs, {
    stdio: "inherit",
    shell: false,
    env,
  });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    throw new Error(`${command} ${args.join(" ")} failed with ${result.status}`);
  }
}

const npm = process.platform === "win32" ? "npm.cmd" : "npm";
const powershell = process.platform === "win32" ? windowsPowerShellPath() : "powershell";
const powershellEnvironment =
  process.platform === "win32"
    ? windowsPowerShellEnvironment()
    : process.env;

// Prepare only native/resource inputs consumed by Tauri. The main desktop
// frontend is built exactly once by tauri_build.mjs before Tauri codegen.
if (process.platform === "win32") {
  const noOcrBundle = process.env.VISUALTEX_NO_OCR_BUNDLE === "1";
  const useFixedOfficeBundle =
    process.env.VISUALTEX_USE_FIXED_OFFICE_BUNDLE === "1";
  if (noOcrBundle) {
    console.log("Skipping bundled OCR Python preparation for the no-OCR installer flavor.");
  } else {
    run(powershell, [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      "scripts/prepare_windows_ocr_python.ps1",
    ], powershellEnvironment);
  }
  if (useFixedOfficeBundle) {
    // P16 final-candidate mode: Office/VSTO/native-preview binaries are frozen
    // inputs. Verify exact accepted identities and never rebuild or restage
    // them. The Office WebView UI is still rebuilt from the current core source.
    run(powershell, [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      "scripts/verify_fixed_windows_office_inputs.ps1",
    ], powershellEnvironment);
    run(npm, ["run", "build:office:windows-native"]);
  } else {
    run(powershell, [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      "scripts/prepare_windows_vsto_runtime.ps1",
    ], powershellEnvironment);
    // MathType native preview is a required Windows bundle resource. Reuse a
    // previously staged, validated 7.9.1 runtime or rebuild it only when the
    // caller explicitly supplies VISUALTEX_MATHTYPE_INSTALLER. Never silently
    // ship a package that falls back to the non-native preview path.
    run(powershell, [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      "scripts/prepare_private_mathtype_runtime.ps1",
    ], powershellEnvironment);
    run(npm, ["run", "build:office:windows-native"]);
    run(powershell, [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      "scripts/build_windows_office.ps1",
      "-Configuration",
      "Release",
      "-SkipTests",
    ], powershellEnvironment);
  }
  run(powershell, [
    "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
    "scripts/verify_windows_native_render_package.ps1",
  ], powershellEnvironment);
}
