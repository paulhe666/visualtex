import { commandRegistry } from "../src/autocomplete/commandRegistry.ts";
import { validateLatex } from "mathlive/ssr";

type Failure = {
  id: string;
  command: string;
  previewLatex: string;
  reason: string;
};

const expectedNativeGaps = new Set<string>();

const failures: Failure[] = [];
for (const command of commandRegistry) {
  try {
    const errors = validateLatex(command.previewLatex);
    const relevantErrors = errors.filter((error) =>
      ["unknown-command", "invalid-command", "unexpected-command-in-math-mode"].includes(
        error.code,
      ),
    );
    if (relevantErrors.length > 0) {
      failures.push({
        id: command.id,
        command: command.command,
        previewLatex: command.previewLatex,
        reason: relevantErrors
          .map((error) => `${error.code}:${error.arg ?? ""}`)
          .join(", "),
      });
    }
  } catch (error) {
    failures.push({
      id: command.id,
      command: command.command,
      previewLatex: command.previewLatex,
      reason: error instanceof Error ? error.message : String(error),
    });
  }
}

const unexpectedFailures = failures.filter(
  (failure) => !expectedNativeGaps.has(failure.id),
);
const missingExpectedCoverage: string[] = [];

console.log(
  JSON.stringify(
    {
      totalCommands: commandRegistry.length,
      nativeGapCount: failures.length,
      expectedNativeGaps: failures,
      unexpectedFailures,
      missingExpectedCoverage,
    },
    null,
    2,
  ),
);

if (unexpectedFailures.length > 0 || missingExpectedCoverage.length > 0) {
  process.exitCode = 1;
}
