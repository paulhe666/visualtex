import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { build, loadConfigFromFile } from "vite";
import {
  mathLiveBrowserEntry,
  visualTexMathLiveContourIntegralCompatibility,
} from "../vite.mathliveIntegralCompatibility.ts";

/**
 * Diagnostic build only. Read the immutable, complete Windows kernel directly
 * from Git while its local fragment import is pending. Never writes source
 * files or changes Git state. Output is isolated from dist and the DMG input.
 * This is NOT a replacement for the normal source-complete build gate.
 */
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const reference = "7e9c8050e5b47405a4f7e5ad214375451bfa28be";
const expectedSHA256 = "806b39764e226095921a0ffeba60dcdd3397bc4f53cd64f1feb9ebe5287ca530";
const git = (...args: string[]) => execFileSync("git", args, {
  cwd: root, maxBuffer: 8 * 1024 * 1024, timeout: 30000,
  stdio: ["ignore", "pipe", "pipe"],
});
const modules = [
  "vite.mathliveIntegralCompatibility.ts",
  "vite.mathliveRuntimeSafety.ts",
  "vite.visualtexMathLiveBehaviorFork.ts",
  "vite.mathliveEditingKernel.ts",
  "vite.mathliveSelection.ts",
  "vite.mathliveSemanticCompletion.ts",
  "src/math/rareIntegralGlyphs.generatedData.ts",
];
for (const relativePath of modules) {
  assert(readFileSync(path.join(root, relativePath)).equals(
    git("show", `${reference}:apps/windows/${relativePath}`),
  ), `Diagnostic build requires the unmodified Windows module: ${relativePath}`);
}
const parts = Array.from({ length: 30 }, (_, index) => git("show",
  `${reference}:apps/windows/vendor/mathlive/kernel-parts/part-${String(index).padStart(3, "0")}.mjsfrag`,
));
const kernel = Buffer.concat(parts);
assert.equal(kernel.length, 1494717);
assert.equal(createHash("sha256").update(kernel).digest("hex"), expectedSHA256);
const config = await loadConfigFromFile(
  { command: "build", mode: "production" }, path.join(root, "vite.config.ts"), root,
);
assert(config, "macOS Vite configuration could not be read");
const plugin = visualTexMathLiveContourIntegralCompatibility();
// The canonical transform is unchanged. Only the diagnostic source provider
// reads immutable Git blobs instead of the not-yet-imported local fragments.
let sourceLoads = 0;
plugin.load = (id) => {
  if (id !== "\0visualtex-mathlive-kernel") return null;
  sourceLoads += 1;
  return kernel.toString("utf8");
};
const flattenPlugins = (await Promise.all(config.config.plugins ?? []))
  .flat(Infinity).filter(Boolean);
const otherPlugins = flattenPlugins.filter((entry: any) =>
  entry.name !== "visualtex-mathlive-integral-compatibility",
);
const aliases = config.config.resolve?.alias;
assert(Array.isArray(aliases), "Expected the macOS alias list");
const outDir = path.join(root, "build", "macos-committed-kernel-audit", "dist");
console.log(JSON.stringify({
  purpose: "diagnostic-only-not-DMG-input", reference,
  kernelBytes: kernel.length, kernelSHA256: expectedSHA256,
  verifiedCanonicalModules: modules.length, outDir,
}, null, 2));
await build({
  ...config.config,
  configFile: false,
  root,
  plugins: [plugin, ...otherPlugins],
  resolve: {
    ...config.config.resolve,
    alias: [
      { find: /^mathlive$/, replacement: mathLiveBrowserEntry },
      ...aliases.filter((entry) => String(entry.find) !== String(/^mathlive$/)),
    ],
  },
  build: {
    ...config.config.build,
    outDir,
    // Preserve all pre-existing audit artifacts; never empty an unknown dir.
    emptyOutDir: false,
  },
});
assert.equal(sourceLoads, 1, "Diagnostic build did not consume exactly one canonical kernel");
console.log("Canonical Windows kernel diagnostic build passed; local source-complete and DMG gates remain separate.");
