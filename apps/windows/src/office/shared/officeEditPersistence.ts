import { normalizeChineseLatex } from "../../editor/normalizeChineseLatex";
import type { FormulaEditorLine } from "./formulaEditorDocument";

export function canonicalOfficeFingerprintLines(
  lines: ReadonlyArray<Pick<FormulaEditorLine, "latex">>,
) {
  return lines.map((line) => normalizeChineseLatex(line.latex));
}
