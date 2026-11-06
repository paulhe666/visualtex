#!/usr/bin/env node
// Copies the macOS app (apps/macos) and the packages it shares (apps/shared) into the web build and re-applies the
// web patches.
//
//   node scripts/sync-mac-editor.mjs [commit]   copy files from <commit>
//   node scripts/sync-mac-editor.mjs --check    verify the tree matches manifest + patches
//   node scripts/sync-mac-editor.mjs --save-patch <path>
//                                               record local edits of a synced file
//
// Every file listed in web-sync/manifest.json is a copy of apps/macos/<path> (and every file of apps/shared a copy
// at shared/<path>), except for the unified diffs kept in web-sync/patches/. Two mechanical rewrites happen before
// the patches: the manifest's transforms, and imports of apps/shared, which point at shared/ in this repository.
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
import { dirname, join, posix, relative } from "node:path";

const root = join(dirname(new URL(import.meta.url).pathname), "..");
const manifestPath = join(root, "web-sync/manifest.json");
const patchDir = join(root, "web-sync/patches");
const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));
const shared = manifest.shared ?? null;

const git = (args, options = {}) =>
  execFileSync("git", args, { cwd: root, maxBuffer: 1 << 28, ...options });

const excluded = (path) =>
  (manifest.exclude ?? []).some((rule) => (rule.endsWith("/") ? path.startsWith(rule) : path === rule || (rule.startsWith("*") && path.endsWith(rule.slice(1)))));

function listTree(commit, sourcePath) {
  return git(["ls-tree", "-r", "--name-only", commit, sourcePath]).toString().trim().split("\n").filter(Boolean);
}

/** Every synced file: its path here and its path in the main repository. */
function listSourceFiles(commit) {
  const files = [];
  for (const entry of manifest.files) {
    if (!entry.endsWith("/")) {
      files.push({ path: entry, source: `${manifest.sourceRoot}/${entry}` });
      continue;
    }
    for (const source of listTree(commit, `${manifest.sourceRoot}/${entry}`)) {
      const path = source.slice(manifest.sourceRoot.length + 1);
      if (!excluded(path)) files.push({ path, source });
    }
  }
  if (shared) {
    for (const source of listTree(commit, shared.sourceRoot)) {
      const path = `${shared.dest}/${source.slice(shared.sourceRoot.length + 1)}`;
      if (!excluded(path)) files.push({ path, source });
    }
  }
  return files;
}

function sourceOf(path) {
  if (shared && path.startsWith(`${shared.dest}/`)) return `${shared.sourceRoot}/${path.slice(shared.dest.length + 1)}`;
  return `${manifest.sourceRoot}/${path}`;
}

const TEXT = /\.(?:[cm]?[jt]sx?|css|json|html|md)$/;

// A relative import from apps/macos into apps/shared ("../../../shared/...") becomes the same import of shared/ here.
function rewriteSharedImports(text, source, path) {
  if (!shared) return text;
  const sharedParent = posix.dirname(shared.sourceRoot);
  return text.replace(/((?:\.\.\/)+)shared\//g, (match, dots, offset) => {
    if (posix.normalize(posix.join(posix.dirname(source), dots)).replace(/\/$/, "") !== sharedParent) return match;
    void offset;
    const target = posix.relative(posix.dirname(path), shared.dest) || ".";
    return `${target.startsWith(".") ? target : `./${target}`}/`;
  });
}

// Mechanical rewrites (the manifest's transforms, imports of apps/shared). Patches are recorded on top of them.
function sourceContent(commit, path) {
  const source = sourceOf(path);
  const raw = git(["show", `${commit}:${source}`]);
  if (!TEXT.test(path)) return raw;
  let text = raw.toString();
  if (source.startsWith(`${manifest.sourceRoot}/`)) text = rewriteSharedImports(text, source, path);
  for (const rule of (manifest.transforms ?? []).filter((rule) => path.startsWith(rule.prefix))) {
    text = text.split(rule.from).join(rule.to);
  }
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

// Writes pristine sources into `target`, then applies every patch there (unless `pristine`).
function materialize(commit, target, { pristine = false } = {}) {
  const files = listSourceFiles(commit);
  for (const { path } of files) {
    const destination = join(target, path);
    mkdirSync(dirname(destination), { recursive: true });
    writeFileSync(destination, sourceContent(commit, path));
  }
  if (!pristine) {
    for (const patch of patchFiles()) {
      execFileSync("git", ["apply", "--whitespace=nowarn", patch], { cwd: target });
    }
  }
  return files.map((file) => file.path);
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
  // --pristine <commit>: copy without the patches (to port them onto a new base)
  const pristine = mode === "--pristine";
  const commitArg = pristine ? argument : mode;
  const commit = git(["rev-parse", commitArg || manifest.commit]).toString().trim();
  const files = materialize(commit, root, { pristine });
  manifest.commit = commit;
  writeFileSync(manifestPath, JSON.stringify(manifest, null, 2) + "\n");
  console.log(`synced ${files.length} files from ${manifest.sourceRoot}@${commit.slice(0, 7)}${pristine ? " (no patches)" : ""}`);
}
