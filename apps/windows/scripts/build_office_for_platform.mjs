import { spawnSync } from "node:child_process";

const npm = process.platform === "win32" ? "npm.cmd" : "npm";
const script = process.platform === "win32"
  ? "build:office:windows-ole"
  : "build:office:macos";
const result = spawnSync(npm, ["run", script], { stdio: "inherit" });
if (result.error) throw result.error;
if (result.status !== 0) process.exit(result.status ?? 1);
