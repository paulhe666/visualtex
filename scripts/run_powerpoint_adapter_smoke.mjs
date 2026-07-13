import { spawnSync } from "node:child_process";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const tsxCli = join(root, "node_modules", "tsx", "dist", "cli.mjs");
const testFile = join(root, "scripts", "powerpoint_adapter_smoke.mjs");
const variableNames = [
  "POWERPOINT_API_LEVEL",
  "POWERPOINT_FAIL_TAGS",
  "POWERPOINT_REJECT_PNG",
  "POWERPOINT_IN_PLACE_EDIT",
];
const baseEnvironment = { ...process.env };
for (const name of variableNames) delete baseEnvironment[name];

const scenarios = [
  { POWERPOINT_API_LEVEL: "1.1" },
  { POWERPOINT_API_LEVEL: "1.5" },
  { POWERPOINT_API_LEVEL: "1.8" },
  { POWERPOINT_API_LEVEL: "1.10" },
  { POWERPOINT_API_LEVEL: "1.5", POWERPOINT_FAIL_TAGS: "1" },
  { POWERPOINT_API_LEVEL: "1.5", POWERPOINT_REJECT_PNG: "1" },
  { POWERPOINT_API_LEVEL: "1.1", POWERPOINT_IN_PLACE_EDIT: "1" },
];

for (const scenario of scenarios) {
  console.log(`PowerPoint adapter scenario: ${JSON.stringify(scenario)}`);
  const result = spawnSync(process.execPath, [tsxCli, testFile], {
    cwd: root,
    env: { ...baseEnvironment, ...scenario },
    stdio: "inherit",
  });
  if (result.error) throw result.error;
  if (result.status !== 0) process.exit(result.status ?? 1);
}
