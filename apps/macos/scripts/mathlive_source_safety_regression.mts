import assert from "node:assert/strict";
import {
  inspectMathLiveSourceSafety,
  MAX_MATHLIVE_ENVIRONMENT_DEPTH,
  MAX_MATHLIVE_GROUP_DEPTH,
  MAX_MATHLIVE_SOURCE_LENGTH,
} from "../src/editor/mathLiveSourceSafety.ts";

assert.equal(
  inspectMathLiveSourceSafety(
    String.raw`\frac{\sum_{i=1}^{n} i^2}{\sqrt{1+x^2}}`,
  ),
  null,
);

// Long but shallow formulas are legitimate and must not be rejected merely for
// being large. The crash guard targets structural recursion, not ordinary size.
assert.equal(inspectMathLiveSourceSafety("x+".repeat(20_000)), null);

const nestedGroups =
  "{".repeat(MAX_MATHLIVE_GROUP_DEPTH + 1) +
  "x" +
  "}".repeat(MAX_MATHLIVE_GROUP_DEPTH + 1);
assert.deepEqual(inspectMathLiveSourceSafety(nestedGroups), {
  kind: "group-depth",
  value: MAX_MATHLIVE_GROUP_DEPTH + 1,
  limit: MAX_MATHLIVE_GROUP_DEPTH,
});

let nestedEnvironments = "x";
for (let index = 0; index < MAX_MATHLIVE_ENVIRONMENT_DEPTH + 1; index += 1) {
  nestedEnvironments = String.raw`\begin{array}{c}${nestedEnvironments}\end{array}`;
}
assert.deepEqual(inspectMathLiveSourceSafety(nestedEnvironments), {
  kind: "environment-depth",
  value: MAX_MATHLIVE_ENVIRONMENT_DEPTH + 1,
  limit: MAX_MATHLIVE_ENVIRONMENT_DEPTH,
});

// A literal escaped \\begin token must not be mistaken for an environment.
const escapedBegins = String.raw`\\begin{array}{c}`.repeat(
  MAX_MATHLIVE_ENVIRONMENT_DEPTH + 50,
);
assert.equal(inspectMathLiveSourceSafety(escapedBegins), null);

const oversized = "x".repeat(MAX_MATHLIVE_SOURCE_LENGTH + 1);
assert.deepEqual(inspectMathLiveSourceSafety(oversized), {
  kind: "source-length",
  value: MAX_MATHLIVE_SOURCE_LENGTH + 1,
  limit: MAX_MATHLIVE_SOURCE_LENGTH,
});

console.log("VisualTeX MathLive source safety regression passed");
