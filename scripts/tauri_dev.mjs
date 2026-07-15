import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import {
  registerMacosDevUrlHandler,
  unregisterMacosDevUrlHandler,
} from "./register_macos_dev_url_handler.mjs";

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const debugExecutable = join(repositoryRoot, "src-tauri", "target", "debug", "visualtex");
const tauri = process.platform === "win32" ? "tauri.cmd" : "tauri";
const child = spawn(tauri, ["dev", ...process.argv.slice(2)], {
  cwd: repositoryRoot,
  env: process.env,
  stdio: "inherit",
  shell: false,
});

let registered = false;
let shuttingDown = false;
const registrationTimer =
  process.platform === "darwin"
    ? setInterval(() => {
        if (registered || !existsSync(debugExecutable)) return;
        try {
          const app = registerMacosDevUrlHandler();
          registered = true;
          process.stdout.write(`Registered visualtex:// for macOS development: ${app}\n`);
        } catch (error) {
          process.stderr.write(
            `Unable to register the VisualTeX development URL handler: ${
              error instanceof Error ? error.message : String(error)
            }\n`,
          );
        }
      }, 500)
    : null;

function cleanup() {
  if (registrationTimer) clearInterval(registrationTimer);
  if (registered) unregisterMacosDevUrlHandler();
}

function forwardSignal(signal) {
  if (shuttingDown) return;
  shuttingDown = true;
  cleanup();
  if (!child.killed) child.kill(signal);
}

for (const signal of ["SIGINT", "SIGTERM", "SIGHUP"]) {
  process.on(signal, () => forwardSignal(signal));
}

child.on("error", (error) => {
  cleanup();
  process.stderr.write(`Unable to start Tauri development mode: ${error.message}\n`);
  process.exitCode = 1;
});

child.on("exit", (code, signal) => {
  cleanup();
  if (signal) {
    process.kill(process.pid, signal);
    return;
  }
  process.exitCode = code ?? 1;
});
