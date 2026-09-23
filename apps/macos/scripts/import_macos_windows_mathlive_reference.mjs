import assert from "node:assert/strict";
import {
  existsSync,
  lstatSync,
  mkdirSync,
  readFileSync,
  realpathSync,
  writeFileSync,
} from "node:fs";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";

/**
 * One-time, opt-in import of the committed Windows MathLive reference.
 * Default / --check is READ ONLY. --apply creates only missing allowlisted
 * files. Existing different files, symlinks, and a different checkout layout
 * are errors; no overwrite, deletion, Git mutation, network, or Office access.
 * A partial successful import can be resumed by running --apply again.
 */
const REFERENCE = "7e9c8050e5b47405a4f7e5ad214375451bfa28be";
const KERNEL_SHA256 = "806b39764e226095921a0ffeba60dcdd3397bc4f53cd64f1feb9ebe5287ca530";
const root = realpathSync(path.resolve(path.dirname(fileURLToPath(import.meta.url)), ".."));
const args = process.argv.slice(2);
if (args.length > 1 || (args.length === 1 && !["--check", "--apply"].includes(args[0]))) {
  throw new Error("Usage: node scripts/import_macos_windows_mathlive_reference.mjs [--check|--apply]");
}
const apply = args[0] === "--apply";
const git = (...arguments_) => execFileSync("git", arguments_, {
  cwd: root,
  stdio: ["ignore", "pipe", "pipe"],
  timeout: 30000,
  maxBuffer: 8 * 1024 * 1024,
});
const repository = realpathSync(git("rev-parse", "--show-toplevel").toString().trim());
assert.equal(root, path.join(repository, "apps", "macos"), "Run only from this repository's apps/macos checkout");
assert.equal(git("rev-parse", `${REFERENCE}^{commit}`).toString().trim(), REFERENCE);
const originalHead = git("rev-parse", "HEAD").toString().trim();
const originalBranch = git("branch", "--show-current").toString().trim();
const blobHash = (bytes) => createHash("sha1")
  .update(Buffer.from(`blob ${bytes.length}\0`)).update(bytes).digest("hex");
const sha256 = (bytes) => createHash("sha256").update(bytes).digest("hex");
const paths = Array.from({ length: 30 }, (_, index) => {
  const filename = `vendor/mathlive/kernel-parts/part-${String(index).padStart(3, "0")}.mjsfrag`;
  return { source: filename, target: filename, kernel: true };
});
paths.push({
  source: "scripts/targeted_editor_regression.mjs",
  target: "scripts/windows_reference_editor_regression.mjs",
  kernel: false,
});

function validateTarget(relativePath) {
  const target = path.join(root, relativePath);
  assert(target.startsWith(root + path.sep), "Import target escaped the macOS application");
  let cursor = root;
  const components = relativePath.split("/");
  for (let index = 0; index < components.length; index += 1) {
    cursor = path.join(cursor, components[index]);
    let stat;
    try { stat = lstatSync(cursor); } catch (error) {
      if (error.code === "ENOENT") break;
      throw error;
    }
    assert(!stat.isSymbolicLink(), `Refusing a symlink: ${cursor}`);
    assert(index === components.length - 1 ? stat.isFile() : stat.isDirectory(),
      `Unexpected filesystem entry: ${cursor}`);
  }
  return target;
}

// Read and verify ALL inputs before creating any destination.
const plan = paths.map((entry) => {
  const source = `apps/windows/${entry.source}`;
  const expectedBlob = git("rev-parse", `${REFERENCE}:${source}`).toString().trim();
  const bytes = git("cat-file", "blob", expectedBlob);
  assert.equal(blobHash(bytes), expectedBlob, `Git blob validation failed: ${source}`);
  const destination = validateTarget(entry.target);
  const present = existsSync(destination);
  if (present) {
    assert(readFileSync(destination).equals(bytes),
      `Existing file differs; left untouched: ${entry.target}`);
  }
  return { ...entry, bytes, expectedBlob, present };
});
const kernel = Buffer.concat(plan.filter((entry) => entry.kernel).map((entry) => entry.bytes));
assert.equal(kernel.length, 1494717, "Unexpected reference kernel length");
assert.equal(sha256(kernel), KERNEL_SHA256, "Unexpected reference kernel checksum");
assert(existsSync(path.join(root, "scripts/browser_test_runtime.mjs")),
  "The reference regression requires the existing macOS browser_test_runtime.mjs helper");
const missing = plan.filter((entry) => !entry.present);
console.log(JSON.stringify({
  mode: apply ? "apply-missing-only" : "read-only-check",
  reference: REFERENCE,
  workingHead: originalHead,
  workingBranch: originalBranch,
  kernelSHA256: KERNEL_SHA256,
  matchingKernelParts: plan.filter((entry) => entry.kernel && entry.present).length,
  missingKernelParts: missing.filter((entry) => entry.kernel).length,
  filesToCreate: missing.map((entry) => ({ path: entry.target, bytes: entry.bytes.length })),
}, null, 2));

if (apply) {
  for (const entry of missing) {
    const destination = validateTarget(entry.target);
    mkdirSync(path.dirname(destination), { recursive: true });
    // Exclusive creation also refuses a file created concurrently after the
    // preflight; this command never replaces another process's work.
    writeFileSync(destination, entry.bytes, { flag: "wx", mode: 0o644 });
    assert.equal(blobHash(readFileSync(destination)), entry.expectedBlob,
      `Read-back verification failed: ${entry.target}`);
    console.log(`Created and verified ${entry.target}`);
  }
  for (const entry of plan) {
    assert.equal(blobHash(readFileSync(validateTarget(entry.target))), entry.expectedBlob);
  }
  assert.equal(git("rev-parse", "HEAD").toString().trim(), originalHead, "HEAD changed concurrently");
  assert.equal(git("branch", "--show-current").toString().trim(), originalBranch, "Branch changed concurrently");
  console.log("Verified all 30 canonical kernel parts and the separate Windows reference regression runner.");
  console.log("The Vite entry has NOT been changed. Build and runtime acceptance are still required.");
} else {
  console.log("Read-only check complete. No files or Git state were changed.");
}
