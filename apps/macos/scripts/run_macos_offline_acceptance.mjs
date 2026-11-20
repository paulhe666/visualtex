import { spawnSync } from "node:child_process";
import { mkdirSync, writeFileSync } from "node:fs";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const repositoryRoot = resolve(fileURLToPath(new URL("..", import.meta.url)));
const logRoot = join(repositoryRoot, "build-logs", "macos-offline");
mkdirSync(logRoot, { recursive: true });

const commands = [
  {
    id: "01-source-smoke",
    program: "npm",
    args: ["run", "test:macos-offline-office"],
  },
  {
    id: "02-rust-regression",
    program: "cargo",
    args: ["test", "--manifest-path", "src-tauri/Cargo.toml", "--lib", "--no-fail-fast"],
  },
  {
    id: "03-desktop-build",
    program: "npm",
    args: ["run", "build:desktop"],
  },
  {
    id: "04-word-omml",
    program: "npm",
    args: ["run", "test:word-omml"],
  },
  {
    id: "05-platform-onboarding",
    program: "npm",
    args: ["run", "test:platform-onboarding"],
  },
];

const results = [];
for (const command of commands) {
  const startedAt = new Date().toISOString();
  const result = spawnSync(command.program, command.args, {
    cwd: repositoryRoot,
    env: process.env,
    encoding: "utf8",
    maxBuffer: 64 * 1024 * 1024,
  });
  const status = result.status ?? 1;
  const log = [
    `startedAt=${startedAt}`,
    `command=${command.program} ${command.args.join(" ")}`,
    `exitCode=${status}`,
    "",
    "--- stdout ---",
    result.stdout ?? "",
    "",
    "--- stderr ---",
    result.stderr ?? "",
    "",
  ].join("\n");
  writeFileSync(join(logRoot, `${command.id}.log`), log, "utf8");
  results.push({ id: command.id, status });
  process.stdout.write(`${status === 0 ? "PASS" : "FAIL"} ${command.id}\n`);
}

const summary = [
  `VisualTeX macOS offline Office acceptance: ${results.every((item) => item.status === 0) ? "PASS" : "FAIL"}`,
  `completedAt=${new Date().toISOString()}`,
  ...results.map((item) => `${item.id}=${item.status}`),
  "",
].join("\n");
writeFileSync(join(logRoot, "acceptance-summary.log"), summary, "utf8");
process.stdout.write(summary);

if (results.some((item) => item.status !== 0)) process.exitCode = 1;
