import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";
import postcss from "postcss";

// Read-only declaration audit in the actual stylesheet import order. This is
// not a computed-style/browser test: specificity and shorthand interactions
// still require the UI regressions. Never writes or replaces macOS styles.
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const reference = "7e9c8050e5b47405a4f7e5ad214375451bfa28be";
const readWindows = (relativePath) => execFileSync("git", [
  "show", `${reference}:apps/windows/${relativePath}`,
], { cwd: root, encoding: "utf8", maxBuffer: 4 * 1024 * 1024, timeout: 30000 });
const readMacos = (relativePath) => readFileSync(path.join(root, relativePath), "utf8");
const scope = new RegExp(process.argv[2] ??
  "latex-profile|mathlive-suggestion|native-input-suggestion|formula-line-mode|formula-line-format|formula-display-style|is-persistent-action|is-office-editor-header");
const normalize = (value) => value.replace(/\s+/g, " ").trim();
function stylesheetPaths(entry, read) {
  return [...read(entry).matchAll(/import\s+["']([^"']+\.css)["']/g)]
    .map((match) => match[1]).filter((value) => value.startsWith("."))
    .map((value) => path.posix.normalize(path.posix.join(path.posix.dirname(entry), value)));
}
function parentChain(node) {
  const chain = [];
  for (let parent = node.parent; parent; parent = parent.parent) {
    if (parent.type === "atrule") chain.unshift(`@${parent.name} ${normalize(parent.params)}`);
  }
  return chain;
}
function rules(paths, read) {
  const result = new Map();
  for (const filename of paths) {
    const tree = postcss.parse(read(filename), { from: filename });
    tree.walkRules((rule) => {
      if (!scope.test(rule.selector)) return;
      const selector = normalize(rule.selector);
      const parents = parentChain(rule);
      const key = JSON.stringify([parents, selector]);
      const current = result.get(key) ?? { selector, parents, declarations: new Map() };
      for (const node of rule.nodes) {
        if (node.type !== "decl") continue;
        const previous = current.declarations.get(node.prop);
        if (previous?.important && !node.important) continue;
        current.declarations.set(node.prop, {
          value: normalize(node.value), important: Boolean(node.important),
          location: `${filename}:${node.source.start.line}`,
        });
      }
      result.set(key, current);
    });
  }
  return result;
}
const windowsPaths = stylesheetPaths("src/main.tsx", readWindows);
const macosPaths = stylesheetPaths("src/desktop/main.tsx", readMacos);
assert(windowsPaths.length > 0 && macosPaths.length > 0, "No stylesheet imports found");
const expected = rules(windowsPaths, readWindows);
const actual = rules(macosPaths, readMacos);
assert(expected.size > 0, "Reference editor CSS scope unexpectedly empty");
const differences = [];
for (const [key, rule] of expected) {
  const match = actual.get(key);
  const declarations = [];
  for (const [property, value] of rule.declarations) {
    const candidate = match?.declarations.get(property);
    if (candidate?.value !== value.value || candidate?.important !== value.important) {
      declarations.push({ property, expected: value, actual: candidate ?? null });
    }
  }
  if (declarations.length) differences.push({ selector: rule.selector, parents: rule.parents, declarations });
}
const extra = [...actual].filter(([key]) => !expected.has(key)).map(([, rule]) => ({
  selector: rule.selector, parents: rule.parents,
  location: [...rule.declarations.values()][0]?.location,
}));
console.log(JSON.stringify({
  reference, scope: scope.source, windowsPaths, macosPaths,
  expectedRules: expected.size, actualRules: actual.size,
  differences, extra,
}, null, 2));
if (differences.length) process.exitCode = 1;
