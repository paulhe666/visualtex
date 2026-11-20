import { spawnSync } from "node:child_process";
import {
  closeSync,
  mkdirSync,
  openSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { join } from "node:path";
import process from "node:process";

const lockPath = "/tmp/visualtex-document-import-browser-suite.lock";
const logRoot = join(process.cwd(), "build-logs", "document-import");
const stages = [
  { name: "default", args: [] },
  { name: "literal-fallback", args: ["--literal-fallback"] },
  { name: "theorem-structure", args: ["--theorem-structure"] },
  { name: "edge-structures", args: ["--edge-structures"] },
  { name: "long-physics", args: ["--long-physics"] },
  { name: "boundary-value", args: ["--boundary-value"] },
];

function processExists(pid) {
  if (!Number.isInteger(pid) || pid <= 0) return false;
  try {
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

function acquireLock() {
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      const descriptor = openSync(lockPath, "wx", 0o600);
      writeFileSync(descriptor, `${process.pid}\n`, "utf8");
      closeSync(descriptor);
      return;
    } catch (error) {
      if (error?.code !== "EEXIST") throw error;
      const existingPid = Number.parseInt(readFileSync(lockPath, "utf8").trim(), 10);
      if (processExists(existingPid)) {
        throw new Error(
          `Document-import browser acceptance is already running as PID ${existingPid}`,
        );
      }
      rmSync(lockPath, { force: true });
    }
  }
  throw new Error("Unable to acquire the document-import browser acceptance lock");
}

function timestamp() {
  return new Date().toISOString().replace(/[:.]/g, "-");
}

function tail(value, lineCount = 60) {
  return String(value ?? "")
    .split(/\r?\n/)
    .slice(-lineCount)
    .join("\n");
}

acquireLock();
mkdirSync(logRoot, { recursive: true });
const runStamp = timestamp();

try {
  for (const stage of stages) {
    const started = Date.now();
    const result = spawnSync(
      process.execPath,
      ["scripts/document_import_browser_regression.mjs", ...stage.args],
      {
        cwd: process.cwd(),
        encoding: "utf8",
        timeout: 90_000,
        maxBuffer: 16 * 1024 * 1024,
        killSignal: "SIGTERM",
      },
    );
    const durationMs = Date.now() - started;
    const logPath = join(logRoot, `${runStamp}-browser-${stage.name}.log`);
    const output = [
      `stage=${stage.name}`,
      `durationMs=${durationMs}`,
      `status=${result.status}`,
      `signal=${result.signal ?? ""}`,
      "--- stdout ---",
      result.stdout ?? "",
      "--- stderr ---",
      result.stderr ?? "",
      result.error ? `--- spawn error ---\n${result.error.stack ?? result.error}` : "",
    ].join("\n");
    writeFileSync(logPath, output, { encoding: "utf8", mode: 0o600 });

    if (result.error || result.status !== 0) {
      const reason = result.error?.code === "ETIMEDOUT"
        ? `timed out after ${durationMs} ms`
        : `exited with status ${result.status ?? "unknown"}`;
      console.error(`FAIL document-import browser ${stage.name}: ${reason}`);
      console.error(tail(`${result.stdout ?? ""}\n${result.stderr ?? ""}`));
      console.error(`Full log: ${logPath}`);
      process.exitCode = 1;
      break;
    }
    console.log(
      `PASS document-import browser ${stage.name} (${durationMs} ms) log=${logPath}`,
    );
  }
} finally {
  rmSync(lockPath, { force: true });
}

if (!process.exitCode) {
  console.log("Document import browser suite passed serially with no overlapping processes");
}
