#!/usr/bin/env node
// Keeps macOS's copy of Windows' LaTeX -> MathML step byte-identical.
//
// Word equations must come out of VisualTeX identically on Windows and macOS
// (docs/equation-format-v2.md). The OMML writer is the Rust port of Windows'
// C# writer; its input MathML must therefore come from exactly the same
// TypeScript that Windows' Office dialog runs: apps/windows/src/export/runtime.ts
// latexToMathMl and every module it imports. This script copies that import
// closure into apps/macos/src/office/wordMathMl/windows/ (same relative paths,
// unchanged bytes).
//
//   node scripts/sync_word_mathml_mirror.mjs           copy (run after changing those Windows files)
//   node scripts/sync_word_mathml_mirror.mjs --check   fail if the mirror differs (tests/CI)
//
// --check passes without comparing when apps/windows is absent (web branch).
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from "node:fs";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const macRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const windowsSrc = resolve(macRoot, "../windows/src");
const mirrorRoot = join(macRoot, "src/office/wordMathMl/windows");
const entry = join(windowsSrc, "export/runtime.ts");
const check = process.argv.includes("--check");

if (!existsSync(entry)) {
  if (check) {
    console.log("Word MathML mirror: apps/windows is not present; nothing to compare.");
    process.exit(0);
  }
  console.error(`Missing ${entry}`);
  process.exit(1);
}

const importPattern = /(?:import|export)\s+(?:type\s+)?(?:[^'";]*?\sfrom\s+)?["']([^"']+)["']/g;
function resolveImport(from, specifier) {
  if (!specifier.startsWith(".")) return null;
  const base = resolve(dirname(from), specifier);
  for (const candidate of [base, `${base}.ts`, `${base}.tsx`, join(base, "index.ts")]) {
    if (existsSync(candidate) && statSync(candidate).isFile()) return candidate;
  }
  throw new Error(`${relative(windowsSrc, from)} imports missing ${specifier}`);
}
const closure = [];
const seen = new Set();
(function walk(file) {
  if (seen.has(file)) return;
  seen.add(file);
  closure.push(file);
  for (const match of readFileSync(file, "utf8").matchAll(importPattern)) {
    const target = resolveImport(file, match[1]);
    if (target) walk(target);
  }
})(entry);

const expected = new Map(closure.map((file) => [relative(windowsSrc, file), readFileSync(file)]));
function mirrorFiles(dir = mirrorRoot, found = new Map()) {
  if (!existsSync(dir)) return found;
  for (const name of readdirSync(dir)) {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) mirrorFiles(path, found);
    else found.set(relative(mirrorRoot, path), readFileSync(path));
  }
  return found;
}
const actual = mirrorFiles();

if (check) {
  const problems = [];
  for (const [path, bytes] of expected) {
    if (!actual.has(path)) problems.push(`missing ${path}`);
    else if (!actual.get(path).equals(bytes)) problems.push(`differs ${path}`);
  }
  for (const path of actual.keys()) if (!expected.has(path)) problems.push(`extra ${path}`);
  if (problems.length) {
    console.error(`Word MathML mirror is out of date (run node scripts/sync_word_mathml_mirror.mjs):\n  ${problems.join("\n  ")}`);
    process.exit(1);
  }
  console.log(`Word MathML mirror matches Windows (${expected.size} files).`);
  process.exit(0);
}

rmSync(mirrorRoot, { recursive: true, force: true });
for (const [path, bytes] of expected) {
  const target = join(mirrorRoot, path);
  mkdirSync(dirname(target), { recursive: true });
  writeFileSync(target, bytes);
}
console.log(`Copied ${expected.size} files to ${relative(macRoot, mirrorRoot)}.`);
