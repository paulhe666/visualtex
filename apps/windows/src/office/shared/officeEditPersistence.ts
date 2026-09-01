import { normalizeChineseLatex } from "../../editor/normalizeChineseLatex";
import type { FormulaEditorLine } from "./formulaEditorDocument";

export function canonicalOfficeFingerprintLines(
  lines: ReadonlyArray<Pick<FormulaEditorLine, "latex">>,
) {
  return lines.map((line) => normalizeChineseLatex(line.latex));
}

export function isWordOmmlNumberingOnlyEdit(input: {
  mode?: string | null;
  objectMode?: string | null;
  displayMode?: string | null;
  numbered: boolean;
  originalNumbered: boolean;
  originalFingerprint: string;
  fingerprintAtOriginalNumbering: string;
}) {
  return (
    input.mode === "edit" &&
    input.objectMode === "wordOmml" &&
    input.displayMode === "block" &&
    input.numbered !== input.originalNumbered &&
    Boolean(input.originalFingerprint) &&
    input.fingerprintAtOriginalNumbering === input.originalFingerprint
  );
}

export function persistedOfficeLines(
  numberingOnlyEdit: boolean,
  originalLines: FormulaEditorLine[] | undefined,
  currentLines: FormulaEditorLine[],
) {
  return numberingOnlyEdit && originalLines?.length
    ? originalLines
    : currentLines;
}
