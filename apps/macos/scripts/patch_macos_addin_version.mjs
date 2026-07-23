import { createHash } from "node:crypto";
import {
  readFileSync,
  renameSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { unzipSync, zipSync } from "fflate";

const [oldVersion, newVersion] = process.argv.slice(2);
if (!oldVersion || !newVersion) {
  throw new Error(
    "Usage: node scripts/patch_macos_addin_version.mjs <old-version> <new-version>",
  );
}

const oldBytes = Buffer.from(oldVersion, "utf8");
const newBytes = Buffer.from(newVersion, "utf8");
if (oldBytes.length !== newBytes.length) {
  throw new Error(
    `Compiled VBA version patch requires equal UTF-8 byte lengths: ${oldVersion} (${oldBytes.length}) vs ${newVersion} (${newBytes.length})`,
  );
}

const appRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const resourcesRoot = join(appRoot, "office", "macos-offline", "resources");
const packageVersion = JSON.parse(
  readFileSync(join(appRoot, "package.json"), "utf8"),
).version;
if (packageVersion !== newVersion) {
  throw new Error(
    `Target add-in version ${newVersion} does not match package.json ${packageVersion}`,
  );
}

const protocolSource = readFileSync(
  join(appRoot, "office", "macos-offline", "shared", "VTProtocol.bas"),
  "utf8",
);
if (!protocolSource.includes(`VT_PLUGIN_VERSION As String = "${newVersion}"`)) {
  throw new Error("VTProtocol.bas has not been updated to the target version");
}

const targets = [
  { name: "VisualTeX.dotm", vbaEntry: "word/vbaProject.bin" },
  { name: "VisualTeX.ppam", vbaEntry: "ppt/vbaProject.bin" },
];

function countOccurrences(bytes, needle) {
  const source = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let count = 0;
  let offset = 0;
  while (offset <= source.length - needle.length) {
    const found = source.indexOf(needle, offset);
    if (found < 0) break;
    count += 1;
    offset = found + needle.length;
  }
  return count;
}

function replaceOccurrences(bytes, before, after) {
  const output = Buffer.from(bytes);
  let count = 0;
  let offset = 0;
  while (offset <= output.length - before.length) {
    const found = output.indexOf(before, offset);
    if (found < 0) break;
    after.copy(output, found);
    count += 1;
    offset = found + after.length;
  }
  return { bytes: new Uint8Array(output), count };
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

for (const target of targets) {
  const path = join(resourcesRoot, target.name);
  const archive = unzipSync(new Uint8Array(readFileSync(path)));
  const project = archive[target.vbaEntry];
  if (!project) throw new Error(`${target.name} is missing ${target.vbaEntry}`);

  const oldCount = countOccurrences(project, oldBytes);
  const newCountBefore = countOccurrences(project, newBytes);
  if (oldCount === 0 && newCountBefore === 2) {
    process.stdout.write(`${target.name}: already contains ${newVersion}\n`);
    continue;
  }
  if (oldCount !== 2 || newCountBefore !== 0) {
    throw new Error(
      `${target.name} expected exactly two ${oldVersion} markers and no ${newVersion} markers; found old=${oldCount}, new=${newCountBefore}`,
    );
  }

  const patched = replaceOccurrences(project, oldBytes, newBytes);
  if (patched.count !== 2) {
    throw new Error(`${target.name} patched an unexpected number of version markers`);
  }
  if (
    countOccurrences(patched.bytes, oldBytes) !== 0 ||
    countOccurrences(patched.bytes, newBytes) !== 2
  ) {
    throw new Error(`${target.name} version marker verification failed after patch`);
  }

  archive[target.vbaEntry] = patched.bytes;
  const temporary = `${path}.version-patch.tmp`;
  rmSync(temporary, { force: true });
  writeFileSync(temporary, Buffer.from(zipSync(archive, { level: 9 })));
  renameSync(temporary, path);
  process.stdout.write(
    `${target.name}: patched ${oldVersion} -> ${newVersion} (${patched.count} markers)\n`,
  );
}

const manifestPath = join(resourcesRoot, "addins.json");
const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));
manifest.pluginVersion = newVersion;
manifest.files ??= {};
for (const target of targets) {
  manifest.files[target.name] = {
    sha256: sha256(join(resourcesRoot, target.name)),
  };
}
writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);
process.stdout.write(`addins.json: updated for ${newVersion}\n`);
