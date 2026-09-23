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
run("npm", ["run", "build:desktop"]);
if (process.env.VISUALTEX_API_ONLY_OCR !== "1") {
  run("npm", ["run", "prepare:ocr-offline"]);
} else {
  process.stdout.write("Skipping bundled offline OCR runtime for API-only build.\n");
}
