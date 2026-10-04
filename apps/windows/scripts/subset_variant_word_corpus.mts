import { existsSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";
import { createHash } from "node:crypto";
import { latexToMathMl } from "../src/export/runtime.ts";

const output = resolve(process.argv[2] ?? "");
if (!process.argv[2] || existsSync(output)) throw new Error("Supply a fresh corpus output path");
const families = [
  ["subsetneqq", "varsubsetneqq"],
  ["supsetneqq", "varsupsetneqq"],
  ["nsubseteq", "nsubseteqq"],
  ["nsupseteq", "nsupseteqq"],
] as const;
const formulas: object[] = [];
for (const [ordinary, alternate] of families) {
  const ordinaryRelation = "a\\" + ordinary + " b";
  const alternateRelation = "c\\" + alternate + " d";
  const cases = [
    ["normal", ordinaryRelation + "+\\quad " + alternateRelation, "inline"],
    ["italic", "\\mathit{" + ordinaryRelation + "}+\\quad \\mathit{" + alternateRelation + "}", "inline"],
    ["bold-italic", "\\mathbfit{" + ordinaryRelation + "}+\\quad \\mathbfit{" + alternateRelation + "}", "block"],
    ["scripts", "x^{" + ordinaryRelation + "}_{i}+y^{" + alternateRelation + "}_{j}", "block"],
    ["fraction-fence", "\\frac{" + ordinaryRelation + "}{" + alternateRelation + "}+\\left(x\\" + alternate + " y\\right)^2", "block"],
  ] as const;
  for (const [context, latex, displayMode] of cases) {
    const mathMl = await latexToMathMl(latex, displayMode === "block");
    formulas.push({index: formulas.length + 1, id: "SV" + String(formulas.length + 1).padStart(2, "0"),
      latex, displayMode, mathMl, category: context, ordinary, alternate,
      mathMlSha256: createHash("sha256").update(mathMl).digest("hex")});
  }
}
writeFileSync(output, JSON.stringify({schemaVersion: 1, acceptanceCorpus: true,
  scope: "Twenty compound ordinary/alternate subset-family cases for isolated VT/MathType lifecycle; no Word pixel or OMML fidelity claim.",
  formulas}, null, 2) + "\n");
console.log(JSON.stringify({output, cases: formulas.length, families: families.length}));
