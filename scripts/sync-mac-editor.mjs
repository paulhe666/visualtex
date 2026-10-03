#!/usr/bin/env node
// Copies the macOS editor into the web build and re-applies the web patches.
//
//   node scripts/sync-mac-editor.mjs [commit]   copy files from apps/macos at <commit>
//   node scripts/sync-mac-editor.mjs --check    verify the tree matches manifest + patches
//   node scripts/sync-mac-editor.mjs --save-patch <path>
//                                               record local edits of a synced file
//
// Every file listed in web-sync/manifest.json is a byte-for-byte copy of
// apps/macos/<path>, except for the unified diffs kept in web-sync/patches/.
import { execFileSync } from "node:child_process";
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, relative } from "node:path";

const root = join(dirname(new URL(import.meta.url).pathname), "..");
const manifestPath = join(root, "web-sync/manifest.json");
const patchDir = join(root, "web-sync/patches");
const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));

const git = (args, options = {}) =>
  execFileSync("git", args, { cwd: root, maxBuffer: 1 << 28, ...options });

function listSourceFiles(commit) {
  const files = [];
  for (const entry of manifest.files) {
    if (!entry.endsWith("/")) {
      files.push(entry);
      continue;
    }
    const listed = git([
      "ls-tree", "-r", "--name-only", commit, `${manifest.sourceRoot}/${entry}`,
    ]).toString().trim().split("\n").filter(Boolean);
    files.push(...listed.map((path) => path.slice(manifest.sourceRoot.length + 1)));
  }
  return files;
}

function readSource(commit, path) {
  return git(["show", `${commit}:${manifest.sourceRoot}/${path}`]);
}

// Mechanical rewrites listed in the manifest (for example the macOS Chrome
// path in browser tests). Patches are recorded on top of the rewritten text.
function sourceContent(commit, path) {
  const raw = readSource(commit, path);
  const rules = (manifest.transforms ?? []).filter((rule) => path.startsWith(rule.prefix));
  if (!rules.length) return raw;
  let text = raw.toString();
  for (const rule of rules) text = text.split(rule.from).join(rule.to);
  return Buffer.from(text);
}

function patchFiles() {
  if (!existsSync(patchDir)) return [];
  const result = [];
  const walk = (dir) => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const full = join(dir, entry.name);
      if (entry.isDirectory()) walk(full);
      else if (entry.name.endsWith(".patch")) result.push(full);
    }
  };
  walk(patchDir);
  return result.sort();
}

// Writes pristine sources into `target`, then applies every patch there.
function materialize(commit, target) {
  const files = listSourceFiles(commit);
  for (const path of files) {
    const destination = join(target, path);
    mkdirSync(dirname(destination), { recursive: true });
    writeFileSync(destination, sourceContent(commit, path));
  }
  for (const patch of patchFiles()) {
    execFileSync("git", ["apply", "--whitespace=nowarn", patch], { cwd: target });
  }
  return files;
}

const [mode, argument] = process.argv.slice(2);

if (mode === "--save-patch") {
  if (!argument) throw new Error("--save-patch needs a file path");
  const path = relative(root, join(root, argument));
  const pristine = sourceContent(manifest.commit, path).toString();
  const current = readFileSync(join(root, path), "utf8");
  const patch = join(patchDir, `${path}.patch`);
  if (pristine === current) {
    rmSync(patch, { force: true });
    console.log(`${path} matches macOS; patch removed`);
  } else {
    const scratch = mkdtempSync(join(tmpdir(), "visualtex-sync-"));
    mkdirSync(join(scratch, "a", dirname(path)), { recursive: true });
    mkdirSync(join(scratch, "b", dirname(path)), { recursive: true });
    writeFileSync(join(scratch, "a", path), pristine);
    writeFileSync(join(scratch, "b", path), current);
    let diff = "";
    try {
      execFileSync("git", ["diff", "--no-index", "--src-prefix=", "--dst-prefix=", "--", `a/${path}`, `b/${path}`], { cwd: scratch });
    } catch (error) {
      diff = error.stdout.toString();
    }
    rmSync(scratch, { recursive: true, force: true });
    mkdirSync(dirname(patch), { recursive: true });
    writeFileSync(patch, diff);
    console.log(`saved ${relative(root, patch)}`);
  }
} else if (mode === "--check") {
  const scratch = mkdtempSync(join(tmpdir(), "visualtex-sync-"));
  execFileSync("git", ["init", "-q"], { cwd: scratch });
  const files = materialize(manifest.commit, scratch);
  const drift = files.filter((path) => {
    const local = join(root, path);
    return !existsSync(local) ||
      !readFileSync(local).equals(readFileSync(join(scratch, path)));
  });
  rmSync(scratch, { recursive: true, force: true });
  if (drift.length) {
    console.error("Files differ from macOS + web patches:\n  " + drift.join("\n  "));
    console.error("Record intended edits with --save-patch <path>.");
    process.exit(1);
  }
  console.log(`${files.length} files match ${manifest.sourceRoot}@${manifest.commit.slice(0, 7)} + ${patchFiles().length} patches`);
} else {
  const commit = git(["rev-parse", mode || manifest.commit]).toString().trim();
  const files = materialize(commit, root);
  manifest.commit = commit;
  writeFileSync(manifestPath, JSON.stringify(manifest, null, 2) + "\n");
  console.log(`synced ${files.length} files from ${manifest.sourceRoot}@${commit.slice(0, 7)}`);
}
