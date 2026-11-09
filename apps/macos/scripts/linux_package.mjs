// After `tauri build` on Linux: the .deb's package name is "visualtex" (Tauri derives
// "visual-te-x" from the product name) and its dependencies are listed once.
import { execFileSync } from "node:child_process";
import { existsSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const appDirectory = join(dirname(fileURLToPath(import.meta.url)), "..");
const target = process.env.CARGO_TARGET_DIR ?? join(appDirectory, "src-tauri", "target");
const debDirectory = join(target, "release", "bundle", "deb");
if (!existsSync(debDirectory)) {
  process.stdout.write("No .deb to fix.\n");
  process.exit(0);
}
for (const name of readdirSync(debDirectory).filter((file) => file.endsWith(".deb"))) {
  const deb = join(debDirectory, name);
  const work = mkdtempSync(join(tmpdir(), "visualtex-deb-"));
  try {
    execFileSync("dpkg-deb", ["-R", deb, work]);
    const controlFile = join(work, "DEBIAN", "control");
    const control = readFileSync(controlFile, "utf8")
      .replace(/^Package: .*$/m, "Package: visualtex")
      .replace(/^Depends: (.*)$/m, (_, list) => `Depends: ${[...new Set(list.split(",").map((item) => item.trim()))].join(", ")}`);
    writeFileSync(controlFile, control);
    execFileSync("dpkg-deb", ["--root-owner-group", "-Zxz", "-b", work, deb], { stdio: ["ignore", "ignore", "inherit"] });
    process.stdout.write(`${name}: package visualtex\n`);
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
}
