import { spawnSync } from "node:child_process";
import process from "node:process";

const upstreamRunner = "scripts/run_upstream_macos_regression.mjs";

const checks = [
  {
    label: "source-first wrapper registry coverage",
    command: [
      "--import",
      "tsx",
      "scripts/source_wrapper_registry_regression.mts",
    ],
  },
  {
    label: "core LaTeX compatibility commands",
    command: [
      "--import",
      "tsx",
      "--import",
      "./scripts/kernel-node.mts",
      "scripts/core_latex_command_compatibility_regression.mts",
    ],
  },
  {
    label: "source-first wrapper browser matrix",
    command: [
      "--import",
      "tsx",
      "scripts/source_wrapper_browser_regression.mts",
    ],
  },
  {
    label: "physics kernel compatibility commands",
    command: [
      "--import",
      "tsx",
      "--import",
      "./scripts/kernel-node.mts",
      "scripts/physics_kernel_compatibility_regression.mts",
    ],
  },
  {
    label: "built MathLive physics kernel",
    command: ["scripts/physics_kernel_browser_regression.mjs"],
  },
  {
    label: "Office editor persistence isolation",
    command: ["scripts/editor_persistence_isolation_regression.mjs"],
  },
  {
    label: "current Windows formula tiles",
    command: ["scripts/targeted_editor_regression.mjs", "formula-tiles"],
  },
  {
    label: "current Windows editor layout and matrix builder",
    command: ["scripts/targeted_editor_regression.mjs", "layout"],
  },
  {
    label: "current Windows formula formatting",
    command: ["scripts/targeted_editor_regression.mjs", "formula-formatting"],
  },
  {
    label: "current Windows horizontal formula overflow",
    command: ["scripts/targeted_editor_regression.mjs", "horizontal-overflow"],
  },
  {
    label: "existing vertical structures and native placeholders",
    command: ["scripts/run_new_macos_editor_regressions.mjs"],
  },
  {
    label: "current structural placeholder core",
    command: ["scripts/targeted_editor_regression.mjs", "placeholder-core"],
  },
  {
    label: "Word-like delete and line editing",
    command: ["scripts/targeted_editor_regression.mjs", "delete"],
  },
  {
    label: "Enter transitions",
    command: [
      upstreamRunner,
      "../macos/scripts/enter_transition_regression.mjs",
    ],
  },
  {
    label: "current Windows structured Chinese IME",
    command: ["scripts/targeted_editor_regression.mjs", "structured-chinese-ime"],
  },
  {
    label: "current Windows stale IME recovery",
    command: ["scripts/targeted_editor_regression.mjs", "ime-stale-recovery"],
  },
  {
    label: "line scrolling, deletion, merge and clipping",
    command: [
      upstreamRunner,
      "../macos/scripts/line_scroll_clipping_regression.mjs",
    ],
  },
  {
    label: "complete current Windows input behaviour",
    command: ["scripts/input_behavior_regression.mjs"],
  },
];

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

for (const check of checks) {
  console.log(`\n=== ${check.label} ===`);
  const result = spawnSync(process.execPath, check.command, {
    cwd: process.cwd(),
    stdio: "inherit",
    shell: false,
  });
  if (result.error) throw result.error;
  if (result.status !== 0) process.exit(result.status ?? 1);
  // Each upstream scenario owns a separate Chromium profile and CDP port. Give
  // Windows enough time to release the previous browser process before the next
  // scenario starts, otherwise Chrome can briefly expose an opaque startup page
  // whose localStorage is unavailable even though the target URL is correct.
  await sleep(700);
}

console.log("\nAll current Windows editor acceptance checks passed");
