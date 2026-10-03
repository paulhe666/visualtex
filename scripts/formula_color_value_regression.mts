import assert from "node:assert/strict";
import {
  isSafeFormulaStyleColor,
  normalizeCustomFormulaColor,
} from "../src/workspace/formulaColor";

assert.equal(normalizeCustomFormulaColor("#F7103A"), "#f7103a");
assert.equal(normalizeCustomFormulaColor("#f13"), "#ff1133");
assert.equal(normalizeCustomFormulaColor("rgb(247, 16, 58)"), "#f7103a");
assert.equal(normalizeCustomFormulaColor("rgb(247 16 58)"), "#f7103a");
assert.equal(normalizeCustomFormulaColor("rgb(100%, 0%, 50%)"), "#ff0080");
assert.equal(normalizeCustomFormulaColor("rgba(247, 16, 58, 1)"), "#f7103a");
assert.equal(normalizeCustomFormulaColor("rgb(not-a-color)"), null);

assert.equal(isSafeFormulaStyleColor("#f7103a"), true);
assert.equal(isSafeFormulaStyleColor("red"), true);
assert.equal(isSafeFormulaStyleColor("rgb(247, 16, 58)"), false);
assert.equal(isSafeFormulaStyleColor("color(display-p3 1 0 0)"), false);

console.log("Formula color value regression passed");
