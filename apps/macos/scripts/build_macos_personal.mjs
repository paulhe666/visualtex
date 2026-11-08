// The author's own build (src/edition.ts: document mode with LaTeX): the full app, its DMG named
// VisualTeX_<version>_aarch64-personal.dmg. tauri_build.mjs empties bundle/dmg first, so the public DMG of an
// earlier build is kept aside and put back, as build_macos_api_only.mjs does.
import { spawnSync } from "node:child_process";
import { copyFileSync, existsSync, mkdirSync, readFileSync, renameSync, rmSync } from "node:fs";
import { dirname, join } from "node:path";

if (process.platform !== "darwin") {
  throw new Error("The macOS personal bundle can be built only on macOS.");
}

const version = JSON.parse(readFileSync("package.json", "utf8")).version;
const dmgRoot = join("src-tauri", "target", "release", "bundle", "dmg");
const bundleRoot = dirname(dmgRoot);
const standardDmg = join(dmgRoot, `VisualTeX_${version}_aarch64.dmg`);
const personalDmg = join(dmgRoot, `VisualTeX_${version}_aarch64-personal.dmg`);
const publicBackup = join(bundleRoot, `.VisualTeX_${version}_aarch64-public-backup-${process.pid}.dmg`);

mkdirSync(bundleRoot, { recursive: true });
let preservedPublicDmg = false;
if (existsSync(standardDmg)) {
  copyFileSync(standardDmg, publicBackup);
  preservedPublicDmg = true;
}

try {
  const result = spawnSync(process.execPath, ["scripts/tauri_build.mjs", ...process.argv.slice(2)], {
    stdio: "inherit",
    env: { ...process.env, VISUALTEX_PERSONAL_BUILD: "1" },
  });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    throw new Error(`Personal Tauri build failed with status ${result.status ?? "unknown"}`);
  }
  if (!existsSync(standardDmg)) {
    throw new Error(`Tauri did not produce ${standardDmg}`);
  }
  rmSync(personalDmg, { force: true });
  // moved, not copied: no personal build is left under the public name
  renameSync(standardDmg, personalDmg);
  process.stdout.write(`Personal DMG: ${personalDmg}\n`);
} finally {
  if (preservedPublicDmg && existsSync(publicBackup)) {
    mkdirSync(dmgRoot, { recursive: true });
    copyFileSync(publicBackup, standardDmg);
  }
  rmSync(publicBackup, { force: true });
}
