import assert from "node:assert/strict";
import {
  DEFAULT_SUGGESTION_COUNT,
  MAX_SUGGESTION_COUNT,
  MIN_SUGGESTION_COUNT,
  normalizeSuggestionCount,
} from "../src/stores/settingNormalizers.ts";

assert.equal(normalizeSuggestionCount(undefined), DEFAULT_SUGGESTION_COUNT);
assert.equal(normalizeSuggestionCount(null), DEFAULT_SUGGESTION_COUNT);
assert.equal(normalizeSuggestionCount("6"), DEFAULT_SUGGESTION_COUNT);
assert.equal(normalizeSuggestionCount(Number.NaN), DEFAULT_SUGGESTION_COUNT);
assert.equal(normalizeSuggestionCount(0), MIN_SUGGESTION_COUNT);
assert.equal(normalizeSuggestionCount(-20), MIN_SUGGESTION_COUNT);
assert.equal(normalizeSuggestionCount(4.6), 5);
assert.equal(normalizeSuggestionCount(100), MAX_SUGGESTION_COUNT);

console.log("Suggestion settings migration smoke test passed");
