import { readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

const runtimePath = new URL("../office/macos-offline/word/VTWordAdapter.bas", import.meta.url);
const diagnosticsPath = new URL("../tests/office/word/VTWordAdapterDiagnostics.bas.inc", import.meta.url);

export function wordAdapterSource({ diagnostics = false } = {}) {
  const runtime = readFileSync(runtimePath, "utf8");
  if (!diagnostics) return runtime;

  const tests = readFileSync(diagnosticsPath, "utf8");
  const firstProcedure = tests.search(/^(?:Public|Private) (?:Sub|Function) /m);
  const flag = "#Const VT_WORD_DIAGNOSTICS = False";
  if (firstProcedure < 0 || !runtime.includes(flag)) {
    throw new Error("Word diagnostic source is missing its declarations or build flag");
  }
  // VBA has no partial modules. Compose tests into an isolated build so they
  // can exercise private functions without expanding the production API.
  return runtime.replace(flag, "#Const VT_WORD_DIAGNOSTICS = True\n\n" + tests.slice(0, firstProcedure).trim())
    + "\n" + tests.slice(firstProcedure);
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const output = process.argv[process.argv.indexOf("--output") + 1];
  if (!process.argv.includes("--output") || !output || output.startsWith("--")) {
    throw new Error("Usage: node scripts/word_vba_source.mjs [--diagnostics] --output /path/VTWordAdapter.bas");
  }
  if (resolve(output) === fileURLToPath(runtimePath)) {
    throw new Error("Generated VBA must not overwrite the production source");
  }
  writeFileSync(output, wordAdapterSource({ diagnostics: process.argv.includes("--diagnostics") }));
}
