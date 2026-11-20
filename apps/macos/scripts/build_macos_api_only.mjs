import { spawnSync } from "node:child_process";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  renameSync,
  rmSync,
} from "node:fs";
import { dirname, join } from "node:path";

if (process.platform !== "darwin") {
  throw new Error("The macOS API-only bundle can be built only on macOS.");
}

const version = JSON.parse(readFileSync("package.json", "utf8")).version;
const dmgRoot = join("src-tauri", "target", "release", "bundle", "dmg");
const bundleRoot = dirname(dmgRoot);
const standardDmg = join(dmgRoot, `VisualTeX_${version}_aarch64.dmg`);
const apiOnlyDmg = join(dmgRoot, `VisualTeX_${version}_aarch64-no-ocr.dmg`);
const fullBackup = join(
  bundleRoot,
  `.VisualTeX_${version}_aarch64-full-backup-${process.pid}.dmg`,
);

mkdirSync(bundleRoot, { recursive: true });
let preservedFullDmg = false;
if (existsSync(standardDmg)) {
  copyFileSync(standardDmg, fullBackup);
  preservedFullDmg = true;
}

try {
  const result = spawnSync(process.execPath, ["scripts/tauri_build.mjs"], {
    stdio: "inherit",
    env: {
      ...process.env,
      VISUALTEX_API_ONLY_OCR: "1",
    },
  });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    throw new Error(`API-only Tauri build failed with status ${result.status ?? "unknown"}`);
  }
  if (!existsSync(standardDmg)) {
    throw new Error(`Tauri did not produce ${standardDmg}`);
  }

  rmSync(apiOnlyDmg, { force: true });
  renameSync(standardDmg, apiOnlyDmg);
  process.stdout.write(`API-only DMG: ${apiOnlyDmg}\n`);
} finally {
  if (preservedFullDmg && existsSync(fullBackup)) {
    mkdirSync(dmgRoot, { recursive: true });
    copyFileSync(fullBackup, standardDmg);
  }
  rmSync(fullBackup, { force: true });
}
