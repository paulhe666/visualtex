import { readFileSync } from "node:fs";
import { resolve } from "node:path";

const SCHEMA = "visualtex-office-performance-budget-v1";
const MIN_MEASURED_SAMPLES = 10;
const OPERATIONS = Object.freeze({
  "edit-open": 300,
  "create-apply": 700,
  "edit-apply": 700,
});
const REQUIRED_CASES = Object.freeze([
  ["word-image", "inline"],
  ["word-image", "display"],
  ["word-omml", "inline"],
  ["word-omml", "display"],
  ["powerpoint-svg", "display"],
]);
const DENSE_TARGETS = Object.freeze(["first", "middle", "last"]);

function requiredGroups() {
  const groups = [];
  for (const [scenario, variant] of REQUIRED_CASES) {
    groups.push(
      {
        scenario,
        variant,
        densityProfile: "sparse",
        targetPosition: "new",
        operation: "create-apply",
      },
      {
        scenario,
        variant,
        densityProfile: "sparse",
        targetPosition: "single",
        operation: "edit-open",
      },
      {
        scenario,
        variant,
        densityProfile: "sparse",
        targetPosition: "single",
        operation: "edit-apply",
      },
      {
        scenario,
        variant,
        densityProfile: "dense-50",
        targetPosition: "new",
        operation: "create-apply",
      },
    );
    for (const targetPosition of DENSE_TARGETS) {
      groups.push(
        {
          scenario,
          variant,
          densityProfile: "dense-50",
          targetPosition,
          operation: "edit-open",
        },
        {
          scenario,
          variant,
          densityProfile: "dense-50",
          targetPosition,
          operation: "edit-apply",
        },
      );
    }
  }
  return groups;
}

function argument(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

function usage() {
  process.stderr.write(
    "Usage: node scripts/verify_office_performance_budget.mjs --input /path/result.json [--scope word-image|word-omml|powerpoint-svg]\n",
  );
}

function fail(message) {
  throw new Error(message);
}

function percentile(sorted, percentileValue) {
  if (sorted.length === 0) return null;
  const rank = Math.max(1, Math.ceil((percentileValue / 100) * sorted.length));
  return sorted[Math.min(sorted.length - 1, rank - 1)];
}

function median(sorted) {
  if (sorted.length === 0) return null;
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0
    ? (sorted[middle - 1] + sorted[middle]) / 2
    : sorted[middle];
}

function fixed(value) {
  return Number(value.toFixed(3));
}

const input = argument("--input");
const scope = argument("--scope");
if (!input || (scope && !REQUIRED_CASES.some(([scenario]) => scenario === scope))) {
  usage();
  process.exit(2);
}

const payload = JSON.parse(readFileSync(resolve(input), "utf8"));
if (payload?.schema !== SCHEMA) {
  fail(`Unexpected performance result schema: ${String(payload?.schema)}`);
}
if (!Array.isArray(payload.samples)) {
  fail("Performance result must contain a samples array");
}
if (typeof payload.revision !== "string" || payload.revision.trim().length === 0) {
  fail("Performance result must identify the tested Git revision");
}

const selectedGroups = requiredGroups().filter(
  ({ scenario }) => !scope || scenario === scope,
);
const summaries = [];
const failures = [];

for (const group of selectedGroups) {
  const {
    scenario,
    variant,
    densityProfile,
    targetPosition,
    operation,
  } = group;
  const budgetMs = OPERATIONS[operation];
  const matching = payload.samples.filter(
    (sample) =>
      sample?.scenario === scenario &&
      sample?.variant === variant &&
      sample?.densityProfile === densityProfile &&
      sample?.targetPosition === targetPosition &&
      sample?.operation === operation,
  );
  const warmups = matching.filter((sample) => sample?.warmup === true);
  const measured = matching.filter((sample) => sample?.warmup === false);
  const key = `${scenario}/${variant}/${densityProfile}/${targetPosition}/${operation}`;

  if (warmups.length < 1) {
    failures.push(`${key}: missing the required warm-up record`);
  }
  if (measured.length < MIN_MEASURED_SAMPLES) {
    failures.push(
      `${key}: expected at least ${MIN_MEASURED_SAMPLES} measured samples, found ${measured.length}`,
    );
  }

  const durations = [];
  measured.forEach((sample, index) => {
    if (sample?.success !== true) {
      failures.push(`${key}: measured sample ${index + 1} did not complete successfully`);
    }
    const existingFormulaCount = Number(sample?.existingFormulaCount);
    const editableFormulaCount = Number(sample?.editableFormulaCount);
    if (
      !Number.isInteger(existingFormulaCount) ||
      existingFormulaCount < 0 ||
      (densityProfile === "dense-50" && existingFormulaCount < 50)
    ) {
      failures.push(
        `${key}: measured sample ${index + 1} has invalid existingFormulaCount`,
      );
    }
    if (densityProfile === "dense-50") {
      if (sample?.fixtureSource !== "visualtex-create-apply") {
        failures.push(
          `${key}: measured sample ${index + 1} was not built through VisualTeX create/apply`,
        );
      }
      if (
        !Number.isInteger(editableFormulaCount) ||
        editableFormulaCount < 50 ||
        editableFormulaCount > existingFormulaCount
      ) {
        failures.push(
          `${key}: measured sample ${index + 1} has invalid editableFormulaCount`,
        );
      }
    }
    const durationMs = Number(sample?.durationMs);
    if (!Number.isFinite(durationMs) || durationMs < 0 || durationMs > 60_000) {
      failures.push(`${key}: measured sample ${index + 1} has invalid durationMs`);
      return;
    }
    durations.push(durationMs);
    if (durationMs > budgetMs) {
      failures.push(
        `${key}: sample ${index + 1} took ${fixed(durationMs)} ms, exceeding ${budgetMs} ms`,
      );
    }
  });

  const sorted = [...durations].sort((left, right) => left - right);
  summaries.push({
    scenario,
    variant,
    densityProfile,
    targetPosition,
    operation,
    budgetMs,
    warmups: warmups.length,
    count: sorted.length,
    minMs: sorted.length ? fixed(sorted[0]) : null,
    medianMs: sorted.length ? fixed(median(sorted)) : null,
    p95Ms: sorted.length ? fixed(percentile(sorted, 95)) : null,
    maxMs: sorted.length ? fixed(sorted.at(-1)) : null,
    status:
      sorted.length >= MIN_MEASURED_SAMPLES &&
      measured.every((sample) => {
        const existingFormulaCount = Number(sample?.existingFormulaCount);
        const editableFormulaCount = Number(sample?.editableFormulaCount);
        return (
          sample?.success === true &&
          Number.isInteger(existingFormulaCount) &&
          existingFormulaCount >= 0 &&
          (densityProfile !== "dense-50" ||
            (existingFormulaCount >= 50 &&
              sample?.fixtureSource === "visualtex-create-apply" &&
              Number.isInteger(editableFormulaCount) &&
              editableFormulaCount >= 50 &&
              editableFormulaCount <= existingFormulaCount)) &&
          Number.isFinite(Number(sample?.durationMs)) &&
          Number(sample.durationMs) <= budgetMs
        );
      })
        ? "PASS"
        : "FAIL",
  });
}

const report = {
  schema: "visualtex-office-performance-budget-report-v1",
  status: failures.length === 0 ? "PASS" : "FAIL",
  revision: payload.revision,
  scope: scope ?? "release",
  minimumMeasuredSamples: MIN_MEASURED_SAMPLES,
  summaries,
  failures,
};
process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
if (failures.length > 0) process.exit(1);
