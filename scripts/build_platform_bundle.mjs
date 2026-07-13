import { spawnSync } from "node:child_process";

function run(command, args) {
  const isWindowsCmd =
    process.platform === "win32" && command.toLowerCase().endsWith(".cmd");
  const executable = isWindowsCmd ? (process.env.ComSpec ?? "cmd.exe") : command;
  const executableArgs = isWindowsCmd
    ? ["/d", "/s", "/c", command, ...args]
    : args;
  const result = spawnSync(executable, executableArgs, {
    stdio: "inherit",
    shell: false,
  });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    throw new Error(`${command} ${args.join(" ")} failed with ${result.status}`);
  }
}

const npm = process.platform === "win32" ? "npm.cmd" : "npm";
run(npm, ["run", "build:desktop"]);

if (process.platform === "darwin") {
  run(npm, ["run", "build:office:macos"]);
  run(npm, ["run", "prepare:ocr-offline"]);
} else if (process.platform === "win32") {
  run(npm, ["run", "build:office:windows-ole"]);
  // Tauri opens externalBin before running beforeBuildCommand. The top-level
  // tauri_build wrapper prepares native artifacts first, then sets this flag
  // so the nested build cannot try to overwrite an executable Tauri holds.
  if (process.env.VISUALTEX_TAURI_NATIVE_PREBUILT !== "1") {
    run("powershell.exe", [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      "scripts/build_windows_ole_bridge.ps1",
    ]);
  }
}
