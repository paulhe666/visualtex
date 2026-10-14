import { copyFileSync, existsSync, mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const source = resolve(
  process.argv.includes("--source")
    ? process.argv[process.argv.indexOf("--source") + 1]
    : join(repositoryRoot, "office/macos-offline/resources/VisualTeX.ppam"),
);
const installed = join(
  homedir(),
  "Library/Group Containers/UBF8T346G9.Office/VisualTeX/OfficeAddins/VisualTeX.ppam",
);
const backup = join(
  homedir(),
  "Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch/VisualTeX.ppam.before-fast-open-fix",
);

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

if (!existsSync(source)) throw new Error(`Missing candidate PowerPoint add-in: ${source}`);
mkdirSync(dirname(installed), { recursive: true });
mkdirSync(dirname(backup), { recursive: true });
if (existsSync(installed) && !existsSync(backup)) copyFileSync(installed, backup);
const staged = `${installed}.staged.${process.pid}`;
rmSync(staged, { force: true });
copyFileSync(source, staged);
renameSync(staged, installed);
if (sha256(source) !== sha256(installed)) {
  throw new Error("Installed PowerPoint add-in differs from the candidate bytes");
}
process.stdout.write(`${JSON.stringify({ source, installed, backup, sha256: sha256(installed) }, null, 2)}\n`);
