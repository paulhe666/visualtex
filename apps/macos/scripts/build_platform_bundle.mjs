import { spawnSync } from "node:child_process";

function run(command, args) {
  const result = spawnSync(command, args, {
    stdio: "inherit",
    shell: false,
  });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    throw new Error(`${command} ${args.join(" ")} failed with ${result.status}`);
  }
}

run("node", ["scripts/verify_macos_offline_addins.mjs"]);
if (process.platform === "darwin") {
  // Use the same validator as the install/status commands. The standalone
  // package checker can pass while Rust still expects an older VBA revision.
  run("cargo", [
    "test", "--manifest-path", "src-tauri/Cargo.toml", "--lib",
    "packaged_office_artifacts_pass_the_real_installer_validator",
  ]);
}
run("npm", ["run", "build:desktop"]);
if (process.env.VISUALTEX_API_ONLY_OCR !== "1") {
  run("npm", ["run", "prepare:ocr-offline"]);
} else {
  process.stdout.write("Skipping bundled offline OCR runtime for API-only build.\n");
}
