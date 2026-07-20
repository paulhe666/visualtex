import { execFileSync, spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { registerMacosDevUrlHandler } from "./register_macos_dev_url_handler.mjs";

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const debugExecutable = join(repositoryRoot, "src-tauri", "target", "debug", "visualtex");
const tauri = process.platform === "win32" ? "tauri.cmd" : "tauri";

function pauseMacosOfficeBackground() {
  if (process.platform !== "darwin") return;
  const userId = typeof process.getuid === "function" ? process.getuid() : null;
  if (userId !== null) {
    try {
      execFileSync(
        "/bin/launchctl",
        ["bootout", `gui/${userId}/com.visualtex.studio.office`],
        { stdio: "ignore" },
      );
    } catch {
      // The background LaunchAgent may not currently be loaded.
    }
  }
  try {
    execFileSync(
      "/usr/bin/pkill",
      ["-f", `${debugExecutable} --office-background`],
      { stdio: "ignore" },
    );
  } catch {
    // No competing background development process is a valid state.
  }
}

function stopStaleDevelopmentProcesses() {
  if (process.platform !== "darwin") return;
  for (const pattern of [
    debugExecutable,
    join(repositoryRoot, "node_modules", ".bin", "vite"),
  ]) {
    try {
      execFileSync("/usr/bin/pkill", ["-f", pattern], { stdio: "ignore" });
    } catch {
      // No stale process matching this project path is a valid state.
    }
  }
}

function prepareMacosForegroundDevelopment() {
  if (process.platform !== "darwin") return;
  pauseMacosOfficeBackground();
  stopStaleDevelopmentProcesses();
}

prepareMacosForegroundDevelopment();

const child = spawn(tauri, ["dev", ...process.argv.slice(2)], {
  cwd: repositoryRoot,
  env: process.env,
  stdio: "inherit",
  shell: false,
});

let registered = false;
let shuttingDown = false;
const backgroundGuardTimer =
  process.platform === "darwin"
    ? setInterval(pauseMacosOfficeBackground, 400)
    : null;
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
  if (backgroundGuardTimer) clearInterval(backgroundGuardTimer);
  if (registrationTimer) clearInterval(registrationTimer);
  stopStaleDevelopmentProcesses();
  // Keep the development URL handler registered. When Vite is not running,
  // the handler shows a precise diagnostic instead of leaving Word with
  // kLSApplicationNotFoundErr and a generic formula-creation failure.
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
