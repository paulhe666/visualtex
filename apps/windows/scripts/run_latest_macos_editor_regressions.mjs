import { spawnSync } from "node:child_process";
import process from "node:process";

const upstreamRunner = "scripts/run_upstream_macos_regression.mjs";

const latestMacScript = (name) =>
  `origin-main:apps/macos/scripts/${name}`;

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
      "scripts/core_latex_command_compatibility_regression.mts",
    ],
  },
  {
    label: "source-first wrapper registry",
    command: [
      "--import",
      "tsx",
      "scripts/source_wrapper_registry_regression.mts",
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
    label: "custom tile placeholder previews",
    command: [
      upstreamRunner,
      latestMacScript("custom_tile_placeholder_preview_regression.mjs"),
    ],
  },
  {
    label: "standard and classic editor layouts",
    command: [
      upstreamRunner,
      latestMacScript("editor_layout_switch_regression.mjs"),
    ],
  },
  {
    label: "document-level visual formula alignment",
    command: [
      upstreamRunner,
      latestMacScript("quick_format_toolbar_regression.mjs"),
    ],
  },
  {
    label: "existing vertical structures and native placeholders",
    command: ["scripts/run_new_macos_editor_regressions.mjs"],
  },
  {
    label: "structural placeholder deletion and restoration",
    command: [
      upstreamRunner,
      "../macos/scripts/targeted_editor_regression.mjs",
      "structural-placeholder",
    ],
  },
  {
    label: "placeholder range selection",
    command: [
      upstreamRunner,
      "../macos/scripts/targeted_editor_regression.mjs",
      "placeholder-selection",
    ],
  },
  {
    label: "Word-like delete and line editing",
    command: [
      upstreamRunner,
      "../macos/scripts/targeted_editor_regression.mjs",
      "delete",
    ],
  },
  {
    label: "Enter transitions",
    command: [
      upstreamRunner,
      "../macos/scripts/enter_transition_regression.mjs",
    ],
  },
  {
    label: "IME backslash transaction matrix",
    command: [
      upstreamRunner,
      "../macos/scripts/ime_backslash_matrix.mjs",
    ],
  },
  {
    label: "IME backslash input",
    command: [
      upstreamRunner,
      "../macos/scripts/ime_backslash_regression.mjs",
    ],
  },
  {
    label: "line scrolling, deletion, merge and clipping",
    command: [
      upstreamRunner,
      "../macos/scripts/line_scroll_clipping_regression.mjs",
    ],
  },
  {
    label: "raw command anchors and wrapper auto-exit",
    command: [
      upstreamRunner,
      latestMacScript("raw_command_anchor_regression.mjs"),
    ],
  },
  {
    label: "complete upstream input behaviour",
    command: [
      upstreamRunner,
      "../macos/scripts/input_behavior_regression.mjs",
    ],
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

console.log("\nAll latest macOS editor regressions passed on Windows");
