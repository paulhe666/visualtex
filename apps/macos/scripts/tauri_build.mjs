import { spawnSync } from "node:child_process";

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
  if (result.status !== 0) process.exit(result.status ?? 1);
}

const npm = process.platform === "win32" ? "npm.cmd" : "npm";
const tauri = process.platform === "win32" ? "tauri.cmd" : "tauri";

// Native Office artifacts must exist before Tauri opens externalBin files.
run(npm, ["run", "build:bundle"]);
const forwardedArgs = process.argv.slice(2);
const hasExplicitFeatures = forwardedArgs.some(
  (argument) => argument === "--features" || argument === "-f",
);
const releaseFeatures = hasExplicitFeatures
  ? []
  : ["--features", "tauri/custom-protocol"];
run(tauri, ["build", ...releaseFeatures, ...forwardedArgs], {
  ...process.env,
  VISUALTEX_TAURI_NATIVE_PREBUILT: "1",
});
