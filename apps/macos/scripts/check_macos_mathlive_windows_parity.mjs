import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";

// Read-only migration gate. Never copies files, changes refs, or uses an npm
// kernel as a replacement for the committed Windows kernel.
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const reference = process.argv[2] ?? "7e9c8050e5b47405a4f7e5ad214375451bfa28be";
if (!/^[0-9a-f]{7,40}$/i.test(reference)) {
  throw new Error("Pass a commit hash, not a branch name or a shell expression.");
}
const git = (...args) => execFileSync("git", args, {
  cwd: root,
  maxBuffer: 16 * 1024 * 1024,
  timeout: 30000,
  stdio: ["ignore", "pipe", "pipe"],
});
const sha256 = (data) => createHash("sha256").update(data).digest("hex");
const blobHash = (data) => createHash("sha1")
  .update(Buffer.from(`blob ${data.length}\0`)).update(data).digest("hex");
const commit = git("rev-parse", `${reference}^{commit}`).toString().trim();
const repositoryRoot = git("rev-parse", "--show-toplevel").toString().trim();
const sourcePrefix = "apps/windows/";
const kernelDirectory = "vendor/mathlive/kernel-parts/";
const entries = execFileSync("git", [
  "ls-tree", "-rz", "--full-tree", commit, "--", sourcePrefix + kernelDirectory,
], { cwd: repositoryRoot, maxBuffer: 1024 * 1024, timeout: 30000 })
  .toString().split("\0").filter(Boolean).map((entry) => {
    const [metadata, sourcePath] = entry.split("\t");
    const [mode, type, hash] = metadata.split(" ");
    if (mode !== "100644" || type !== "blob" || !sourcePath) {
      throw new Error(`Unexpected kernel tree entry: ${entry}`);
    }
    return { path: sourcePath.slice(sourcePrefix.length), hash };
  }).sort((a, b) => a.path.localeCompare(b.path));
if (entries.length === 0) throw new Error("Reference has no kernel source parts.");
for (let i = 0; i < entries.length; i += 1) {
  const expected = `${kernelDirectory}part-${String(i).padStart(3, "0")}.mjsfrag`;
  if (entries[i].path !== expected) {
    throw new Error(`Non-contiguous reference kernel: expected ${expected}`);
  }
}

const expectedParts = [];
const actualParts = [];
const missing = [];
const different = [];
let matched = 0;
for (const entry of entries) {
  const expected = git("cat-file", "blob", entry.hash);
  expectedParts.push(expected);
  const filename = path.join(root, entry.path);
  if (!fs.existsSync(filename)) {
    missing.push(entry.path);
    continue;
  }
  const actual = fs.readFileSync(filename);
  actualParts.push(actual);
  if (blobHash(actual) === entry.hash) matched += 1;
  else different.push({ path: entry.path, expectedBlob: entry.hash, actualBlob: blobHash(actual) });
}
const expectedPaths = new Set(entries.map((entry) => entry.path));
const directory = path.join(root, kernelDirectory);
const unexpected = fs.existsSync(directory)
  ? fs.readdirSync(directory).filter((name) => /^part-\d{3}\.mjsfrag$/.test(name))
    .map((name) => kernelDirectory + name).filter((name) => !expectedPaths.has(name))
  : [];
const modulePaths = [
  "vite.mathliveSelection.ts",
  "vite.mathliveSemanticCompletion.ts",
  "vite.mathliveEditingKernel.ts",
  "vite.visualtexMathLiveBehaviorFork.ts",
  "vite.mathliveRuntimeSafety.ts",
  "vite.mathliveIntegralCompatibility.ts",
  "src/math/rareIntegralGlyphs.generatedData.ts",
];
const modules = modulePaths.map((relativePath) => {
  const expected = git("show", `${commit}:${sourcePrefix}${relativePath}`);
  const filename = path.join(root, relativePath);
  const actual = fs.existsSync(filename) ? fs.readFileSync(filename) : null;
  return {
    path: relativePath,
    status: actual === null ? "missing" : actual.equals(expected) ? "match" : "different",
    expectedBlob: blobHash(expected),
    actualBlob: actual === null ? null : blobHash(actual),
  };
});
const config = fs.readFileSync(path.join(root, "vite.config.ts"), "utf8");
const usesFormalLoader = /from\s+["']\.\/vite\.mathliveIntegralCompatibility["']/.test(config)
  && /visualTexMathLiveContourIntegralCompatibility\(\)/.test(config)
  && !/node_modules\/mathlive\/mathlive\.mjs/.test(config);
const referenceKernel = Buffer.concat(expectedParts);
const complete = matched === entries.length && unexpected.length === 0;
const npmFilename = path.join(root, "node_modules/mathlive/mathlive.mjs");
const npmKernel = fs.existsSync(npmFilename) ? fs.readFileSync(npmFilename) : null;
const report = {
  reference: commit,
  workingHead: git("rev-parse", "HEAD").toString().trim(),
  workingBranch: git("branch", "--show-current").toString().trim(),
  sourceParts: {
    expected: entries.length, matched, missing, different, unexpected,
    expectedBytes: referenceKernel.length,
    expectedSHA256: sha256(referenceKernel),
    actualSHA256: complete ? sha256(Buffer.concat(actualParts)) : null,
  },
  modules,
  usesFormalLoader,
  installedNpmKernel: npmKernel === null ? null : {
    bytes: npmKernel.length,
    sha256: sha256(npmKernel),
    equalToWindowsSource: npmKernel.equals(referenceKernel),
  },
  readyForRegression: complete && usesFormalLoader && modules.every((entry) => entry.status === "match"),
};
console.log(JSON.stringify(report, null, 2));
if (!report.readyForRegression) process.exitCode = 1;
