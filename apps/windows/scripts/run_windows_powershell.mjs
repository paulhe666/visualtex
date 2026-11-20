import { spawnSync } from "node:child_process";
import {
  windowsPowerShellEnvironment,
  windowsPowerShellPath,
} from "./windows_powershell.mjs";

const [script, ...args] = process.argv.slice(2);
if (!script) {
  throw new Error("Usage: node scripts/run_windows_powershell.mjs <script.ps1> [arguments...]");
}

const result = spawnSync(
  windowsPowerShellPath(),
  ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script, ...args],
  {
    stdio: "inherit",
    shell: false,
    env: windowsPowerShellEnvironment(),
  },
);
if (result.error) throw result.error;
process.exit(result.status ?? 1);
