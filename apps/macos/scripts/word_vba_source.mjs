import { readFileSync, writeFileSync } from "node:fs";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const offlineRoot = fileURLToPath(new URL("../office/macos-offline/", import.meta.url));
const runtimePath = join(offlineRoot, "word", "VTWordAdapter.bas");
const diagnosticsPath = fileURLToPath(new URL("../tests/office/word/VTWordAdapterDiagnostics.bas.inc", import.meta.url));
const flag = "#Const VT_WORD_DIAGNOSTICS = False";

// The Word add-in's VBA modules in import order: the modules both add-ins share, then Word's own. VTWordAdapter
// holds every macro Word, the Ribbon, the app and AppleScript call; the other VTWord* modules hold their code and
// are private to the project (Option Private Module).
export const WORD_VBA_MODULES = [
  ["VTProtocol", join(offlineRoot, "shared", "VTProtocol.bas")],
  ["VTOfficePaths", join(offlineRoot, "shared", "VTOfficePaths.bas")],
  ["VTMetadata", join(offlineRoot, "shared", "VTMetadata.bas")],
  ["VTLauncher", join(offlineRoot, "shared", "VTLauncher.bas")],
  ["VTErrorHandling", join(offlineRoot, "shared", "VTErrorHandling.bas")],
  ...[
    "VTWordCore",
    "VTWordFormulaData",
    "VTWordSessions",
    "VTWordImages",
    "VTWordEquations",
    "VTWordNumbering",
    "VTWordNumberFormat",
    "VTWordImageNumbers",
    "VTWordDisplayNumbering",
    "VTWordReferences",
    "VTWordDocumentImport",
    "VTWordRedraw",
    "VTWordConversion",
    "VTWordOmmlFont",
    "VTWordAgent",
    "VTWordAdapter",
  ].map((name) => [name, join(offlineRoot, "word", `${name}.bas`)]),
  ["VTWordEvents", join(offlineRoot, "word", "VTWordEvents.cls")],
  ["VTRibbonCallbacks", join(offlineRoot, "word", "VTRibbonCallbacks.bas")],
];

export function wordAdapterSource({ diagnostics = false } = {}) {
  const runtime = readFileSync(runtimePath, "utf8");
  if (!diagnostics) return runtime;

  const tests = readFileSync(diagnosticsPath, "utf8");
  const firstProcedure = tests.search(/^(?:Public|Private) (?:Sub|Function) /m);
  if (firstProcedure < 0 || !runtime.includes(flag)) {
    throw new Error("Word diagnostic source is missing its declarations or build flag");
  }
  // VBA has no partial modules. Compose tests into an isolated build so they
  // can exercise private functions without expanding the production API.
  return runtime.replace(flag, "#Const VT_WORD_DIAGNOSTICS = True\n\n" + tests.slice(0, firstProcedure).trim())
    + "\n" + tests.slice(firstProcedure);
}

// The modules a diagnostic build replaces: VTWordAdapter with the tests composed in, and every other module with
// diagnostic hooks with its flag switched on.
export function wordDiagnosticModuleSources() {
  const sources = { VTWordAdapter: wordAdapterSource({ diagnostics: true }) };
  for (const [name, path] of WORD_VBA_MODULES) {
    if (name === "VTWordAdapter") continue;
    const source = readFileSync(path, "utf8");
    if (source.includes(flag)) sources[name] = source.replace(flag, "#Const VT_WORD_DIAGNOSTICS = True");
  }
  return sources;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const option = (name) => {
    const index = process.argv.indexOf(name);
    const value = index >= 0 ? process.argv[index + 1] : undefined;
    return value && !value.startsWith("--") ? value : undefined;
  };
  const output = option("--output");
  const outputDirectory = option("--output-dir");
  if (!output && !outputDirectory) {
    throw new Error(
      "Usage: node scripts/word_vba_source.mjs [--diagnostics] --output /path/VTWordAdapter.bas\n"
        + "       node scripts/word_vba_source.mjs --diagnostics --output-dir /path (every module a diagnostic build replaces)",
    );
  }
  if (output) {
    if (resolve(output) === runtimePath) {
      throw new Error("Generated VBA must not overwrite the production source");
    }
    writeFileSync(output, wordAdapterSource({ diagnostics: process.argv.includes("--diagnostics") }));
  }
  if (outputDirectory) {
    if (!process.argv.includes("--diagnostics")) throw new Error("--output-dir writes the diagnostic modules");
    if (resolve(outputDirectory) === join(offlineRoot, "word")) {
      throw new Error("Generated VBA must not overwrite the production source");
    }
    for (const [name, source] of Object.entries(wordDiagnosticModuleSources())) {
      writeFileSync(join(outputDirectory, `${name}.bas`), source);
    }
  }
}
