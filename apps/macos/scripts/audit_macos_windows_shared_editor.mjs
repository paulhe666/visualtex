import assert from "node:assert/strict";
import { execFileSync, spawnSync } from "node:child_process";
import { existsSync, readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

// Read-only source comparison against the pinned Windows migration target.
// Prints differences without changing source, the index, history, or branches.
const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const reference = "7e9c8050e5b47405a4f7e5ad214375451bfa28be";
const paths = [
  "src/editor/MathEditor.tsx",
  "src/editor/mathLiveKernelCommands.ts",
  "src/editor/mathLiveIntegralCompatibility.ts",
  "src/editor/mathLiveOptionCompatibility.ts",
  "src/editor/nativeSuggestionPreviews.ts",
  "src/editor/normalizeChineseLatex.ts",
  "src/editor/formulaLineMode.ts",
  "src/autocomplete/CommandSearchEngine.ts",
  "src/autocomplete/compatibilityCommands.ts",
  "src/autocomplete/additionalCommands.ts",
  "src/autocomplete/commandRegistry.ts",
  "src/autocomplete/runtimeCommandRegistry.ts",
  "src/history/HistoryManager.ts",
  "src/history/documentHistory.ts",
  "src/workspace/EditorWorkspace.tsx",
  "src/components/InputBehaviorMenu.tsx",
];
const requested = process.argv[2];
if (requested) assert(paths.includes(requested), "Choose an allowlisted shared editor path");
for (const relative of requested ? [requested] : paths) {
  const expected = execFileSync("git", ["show", `${reference}:apps/windows/${relative}`], {
    cwd: root, maxBuffer: 4 * 1024 * 1024, timeout: 30000,
  });
  const filename = resolve(root, relative);
  if (!existsSync(filename)) { console.log(`MISSING ${relative} (${expected.length} reference bytes)`); continue; }
  const actual = readFileSync(filename);
  if (actual.equals(expected)) { console.log(`MATCH ${relative}`); continue; }
  const result = spawnSync("diff", ["-u", "--label", `Windows ${reference.slice(0, 8)}/${relative}`, "--label", `macOS/${relative}`, "-", filename], {
    cwd: root, input: expected, encoding: "utf8", maxBuffer: 8 * 1024 * 1024, timeout: 30000,
  });
  if (result.error) throw result.error;
  assert([0, 1].includes(result.status), result.stderr || "diff failed");
  if (requested) console.log(result.stdout);
  else {
    const lines = result.stdout.split("\n");
    const added = lines.filter((line) => line.startsWith("+") && !line.startsWith("+++")).length;
    const removed = lines.filter((line) => line.startsWith("-") && !line.startsWith("---")).length;
    console.log(`DIFFERENT ${relative}: +${added} -${removed} (macOS ${actual.length}, Windows ${expected.length} bytes)`);
  }
}
