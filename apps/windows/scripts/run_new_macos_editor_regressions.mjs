import { spawnSync } from "node:child_process";
import process from "node:process";

const scenarios = [
  "vertical-structure-probe",
  "vertical-structure-navigation",
  "native-structure-audit",
  "native-structure-input-under",
  "native-structure-input-multi",
  "native-structure-input-core",
];

for (const scenario of scenarios) {
  console.log(`=== ${scenario} ===`);
  const result = spawnSync(
    process.execPath,
    [
      "scripts/run_upstream_macos_regression.mjs",
      "../macos/scripts/targeted_editor_regression.mjs",
      scenario,
    ],
    {
      cwd: process.cwd(),
      stdio: "inherit",
      shell: false,
    },
  );
  if (result.error) throw result.error;
  if (result.status !== 0) process.exit(result.status ?? 1);
}

console.log("All new macOS editor regression scenarios passed on Windows");
