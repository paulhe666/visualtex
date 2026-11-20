// Every check of one app: its builds once, then each test:*, verify:* and check:* script of its package.json, without
// the build that script would run first. Left out: scripts that only chain other scripts, and the ones the app's
// package.json lists under "checks.skip" with the reason (they need Office, an installed or release build, the OCR
// model packs, or they use the signed-in session's clipboard). "checks.timeouts" gives long steps more than 20 minutes.
//
//   node tools/check_app.mjs [windows|macos] [--list] [--only name,name] [--logs dir]
//
// Without an app name it checks the app of this system. Each step's output goes to <logs>/<step>.log (default
// apps/<app>/.check-logs), the results to <logs>/summary.json; the exit code is 1 when a step failed.
import { spawn, spawnSync } from "node:child_process";
import { createWriteStream, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const args = process.argv.slice(2);
const option = (name) => {
  const index = args.indexOf(name);
  return index >= 0 ? args[index + 1] : undefined;
};
const app =
  args.find((value) => value === "windows" || value === "macos") ??
  { win32: "windows", darwin: "macos" }[process.platform];
if (!app) {
  console.error("Name the app (windows or macos): this system builds neither.");
  process.exit(2);
}
const cwd = join(root, "apps", app);
const pkg = JSON.parse(readFileSync(join(cwd, "package.json"), "utf8"));
const checks = pkg.checks ?? {};
const skip = checks.skip ?? {};
const timeouts = checks.timeouts ?? {};

// "npm run build:desktop && x" -> "x": the builds run once, first
const builds = checks.build ?? ["build:desktop"];
const buildPrefix = new RegExp(`^npm run (${builds.map((name) => name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|")}) && `);
const onlyChains = (command) => command.split("&&").every((part) => part.trim().startsWith("npm run "));

let steps = builds.map((name) => ({ name, cmd: `npm run ${name}`, timeout: timeouts[name] ?? 1200 }));
for (const [name, command] of Object.entries(pkg.scripts)) {
  if (!/^(test|verify|check)\b/.test(name) || Object.hasOwn(skip, name)) continue;
  let body = command;
  while (buildPrefix.test(body)) body = body.replace(buildPrefix, "");
  if (onlyChains(body)) continue;
  steps.push({ name, cmd: body, timeout: timeouts[name] ?? 1200 });
}
for (const name of Object.keys(skip)) {
  if (!Object.hasOwn(pkg.scripts, name)) console.warn(`checks.skip names a script that does not exist: ${name}`);
}

if (args.includes("--list")) {
  for (const step of steps) console.log(`${step.name}\t${step.cmd}`);
  for (const [name, reason] of Object.entries(skip)) console.log(`(left out) ${name}: ${reason}`);
  process.exit(0);
}
const only = option("--only");
if (only) {
  const wanted = new Set(only.split(","));
  steps = steps.filter((step) => wanted.has(step.name));
}

const logs = resolve(option("--logs") ?? join(cwd, ".check-logs"));
mkdirSync(logs, { recursive: true });
const head = spawnSync("git", ["rev-parse", "HEAD"], { cwd, encoding: "utf8" }).stdout.trim();
const summary = { app, head, node: process.version, platform: process.platform, started: new Date().toISOString(), steps: {} };

// as npm does for its scripts: the app's node_modules/.bin first on the path (Windows spells the key "Path"); the
// tests' local servers never go through a proxy
const pathKey = Object.keys(process.env).find((key) => key.toUpperCase() === "PATH") ?? "PATH";
const noProxy = [process.env.NO_PROXY ?? process.env.no_proxy, "127.0.0.1", "localhost"].filter(Boolean).join(",");
const env = {
  ...process.env,
  [pathKey]: [join(cwd, "node_modules", ".bin"), process.env[pathKey]].join(process.platform === "win32" ? ";" : ":"),
  NO_PROXY: noProxy,
  no_proxy: noProxy,
};

function run(step) {
  return new Promise((done) => {
    const log = createWriteStream(join(logs, `${step.name.replace(/[^\w.-]+/g, "_")}.log`));
    const started = Date.now();
    const windows = process.platform === "win32";
    const child = windows
      ? spawn(step.cmd, { cwd, shell: true, env, windowsHide: true })
      : spawn("/bin/sh", ["-c", step.cmd], { cwd, env, detached: true });
    const tail = [];
    let partial = "";
    const take = (chunk) => {
      log.write(chunk);
      const lines = (partial + chunk.toString("utf8")).split(/\r?\n/);
      partial = lines.pop();
      for (const line of lines) {
        tail.push(line);
        if (tail.length > 40) tail.shift();
      }
    };
    child.stdout.on("data", take);
    child.stderr.on("data", take);
    let timedOut = false;
    const timer = setTimeout(() => {
      timedOut = true;
      if (windows) spawnSync("taskkill", ["/pid", String(child.pid), "/T", "/F"]);
      else {
        try {
          process.kill(-child.pid, "SIGKILL");
        } catch {
          // already gone
        }
      }
    }, step.timeout * 1000);
    child.on("close", (code, signal) => {
      clearTimeout(timer);
      if (partial) tail.push(partial);
      log.end();
      done({ code, signal, timedOut, seconds: Math.round((Date.now() - started) / 1000), tail: tail.slice(-25) });
    });
  });
}

for (const step of steps) {
  process.stdout.write(`[${new Date().toISOString()}] ${step.name} ... `);
  const result = await run(step);
  result.ok = result.code === 0 && !result.timedOut;
  summary.steps[step.name] = { cmd: step.cmd, ...result };
  summary.finished = new Date().toISOString();
  writeFileSync(join(logs, "summary.json"), JSON.stringify(summary, null, 1));
  console.log(result.ok ? `ok ${result.seconds}s` : `FAIL code=${result.code}${result.timedOut ? " timeout" : ""} ${result.seconds}s`);
  // without its build nothing after it means anything
  if (!result.ok && builds.includes(step.name)) break;
}
const failed = Object.entries(summary.steps).filter(([, step]) => !step.ok).map(([name]) => name);
console.log(`${app}: ${Object.keys(summary.steps).length} steps, ${failed.length} failed${failed.length ? `: ${failed.join(", ")}` : ""}`);
process.exitCode = failed.length ? 1 : 0;
