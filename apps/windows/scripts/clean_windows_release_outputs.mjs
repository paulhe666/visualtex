import { readdir, rm } from "node:fs/promises";
import { join, resolve, toNamespacedPath } from "node:path";
import process from "node:process";
import { setTimeout as delay } from "node:timers/promises";

if (process.platform !== "win32") {
  throw new Error("Windows release cleanup must run on Windows");
}

const root = process.cwd();
const repositoryRoot = resolve(root, "..", "..");
const argumentsList = process.argv.slice(2);
const cleanAcceptanceOutputs = argumentsList.includes("--acceptance");
const dryRun = argumentsList.includes("--dry-run");
const releaseRoot = resolve(root, "src-tauri/target/release");
const fixedPaths = [
  resolve(root, "dist"),
  join(releaseRoot, "visualtex.exe"),
  join(releaseRoot, "frontend-dist-manifest.json"),
  join(releaseRoot, "nsis"),
  join(releaseRoot, "bundle", "nsis"),
];

async function remove(path, label = "release output") {
  if (dryRun) {
    console.log(`Would clean ${label}: ${path}`);
    return;
  }
  const retryable = new Set(["EBUSY", "EPERM", "ENOTEMPTY"]);
  for (let attempt = 1; attempt <= 30; attempt += 1) {
    try {
      await rm(path, {
        recursive: true,
        force: true,
        maxRetries: 3,
        retryDelay: 250,
      });
      console.log(`Cleaned ${label}: ${path}`);
      return;
    } catch (error) {
      if (
        attempt === 30 ||
        !error ||
        typeof error !== "object" ||
        !("code" in error) ||
        !retryable.has(String(error.code))
      ) {
        throw error;
      }
      const waitMs = Math.min(5000, attempt * 400);
      console.warn(
        `Release output is temporarily locked; retrying in ${waitMs} ms (${attempt}/30): ${path}`,
      );
      await delay(waitMs);
    }
  }
}

for (const path of fixedPaths) await remove(path);

if (cleanAcceptanceOutputs) {
  const acceptancePaths = [
    resolve(root, "artifacts"),
    resolve(root, "build-logs"),
    resolve(root, "%TEMP%"),
    resolve(root, "src-tauri/target-perf"),
    resolve(root, "src-tauri/src-tauri"),
    resolve(repositoryRoot, "artifacts"),
    resolve(repositoryRoot, "%TEMP%"),
  ];
  for (const path of acceptancePaths) {
    await remove(path, "local acceptance output");
  }

  // NUL is a reserved DOS device name. Historical Git-Bash runs can still
  // materialize an NTFS entry with that spelling; the extended-length form is
  // required to address it as a filesystem path rather than as the device.
  for (const nulPath of [resolve(root, "NUL"), resolve(repositoryRoot, "NUL")]) {
    await remove(toNamespacedPath(nulPath), "reserved-name acceptance artifact");
  }

  for (const parent of [root, repositoryRoot]) {
    const rootEntries = await readdir(parent, { withFileTypes: true }).catch(() => []);
    for (const entry of rootEntries) {
      if (entry.isDirectory() && entry.name.startsWith("Tempvisualtex-")) {
        await remove(join(parent, entry.name), "local acceptance output");
      }
    }
  }
}

for (const parent of [join(releaseRoot, "build"), join(releaseRoot, ".fingerprint")]) {
  const entries = await readdir(parent, { withFileTypes: true }).catch(() => []);
  for (const entry of entries) {
    if (entry.name.startsWith("visualtex-")) {
      await remove(join(parent, entry.name));
    }
  }
}

console.log(
  dryRun
    ? "Dry run complete; no files were removed."
    : cleanAcceptanceOutputs
      ? "Windows release and known local acceptance outputs cleaned. User data, Office build outputs and unrelated Cargo dependencies were not touched."
      : "Windows release outputs cleaned without touching artifacts/, build-logs/, user data, Office build outputs or unrelated Cargo dependencies. Pass --acceptance to remove only the known local acceptance-output roots as well.",
);
