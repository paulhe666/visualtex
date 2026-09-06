import { copyFileSync, existsSync, readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { homedir } from "node:os";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { dirname } from "node:path";

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const resourcesRoot = join(repositoryRoot, "office/macos-offline/resources");
const verifiedInstalledWord = join(
  homedir(),
  "Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/VisualTeX.dotm",
);
const resourceWord = join(resourcesRoot, "VisualTeX.dotm");
const resourcePowerPoint = join(resourcesRoot, "VisualTeX.ppam");
const manifestPath = join(resourcesRoot, "addins.json");
const packageVersion = JSON.parse(readFileSync(join(repositoryRoot, "package.json"), "utf8")).version;

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

for (const [path, label] of [
  [verifiedInstalledWord, "verified installed Word add-in"],
  [resourcePowerPoint, "PowerPoint resource"],
]) {
  if (!existsSync(path)) throw new Error(`Missing ${label}: ${path}`);
}

copyFileSync(verifiedInstalledWord, resourceWord);
const manifest = {
  schemaVersion: 1,
  pluginVersion: packageVersion,
  files: {
    "VisualTeX.dotm": { sha256: sha256(resourceWord) },
    "VisualTeX.ppam": { sha256: sha256(resourcePowerPoint) },
  },
};
writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, "utf8");
process.stdout.write(`${JSON.stringify({
  wordSha256: manifest.files["VisualTeX.dotm"].sha256,
  powerpointSha256: manifest.files["VisualTeX.ppam"].sha256,
  manifestPath,
}, null, 2)}\n`);
