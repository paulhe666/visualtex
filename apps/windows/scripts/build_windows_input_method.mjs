// Builds VisualTeX's Windows input method (input-method/) into input-method/build/VisualTeX Input:
// visualtex_tip.dll, the TSF text input processor that types formulas (cargo, release; registered
// by the installer, elevated), and a build id (what the input method does: the installer compares
// it to decide whether the installed copy is current).
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { cpSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

if (process.platform !== "win32") {
  process.stdout.write("The Windows input method is built on Windows only.\n");
  process.exit(0);
}

const appDirectory = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(appDirectory, "input-method");
const output = join(source, "build", "VisualTeX Input");

rmSync(output, { recursive: true, force: true });
mkdirSync(output, { recursive: true });
const target = join(source, "build", "target");
execFileSync("cargo", ["build", "--release", "--manifest-path", join(source, "Cargo.toml"), "--target-dir", target], {
  stdio: "inherit",
});
cpSync(join(target, "release", "visualtex_tip.dll"), join(output, "visualtex_tip.dll"));

const digest = createHash("sha256");
digest.update(readFileSync(join(output, "visualtex_tip.dll")));
writeFileSync(join(output, "build-id"), `${digest.digest("hex").slice(0, 16)}\n`);
process.stdout.write(`Built ${output}\n`);
