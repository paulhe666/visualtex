import { formatLatexLines } from "../../../../shared/visualtex-app/src/clipboard/LatexCopyService";
import { stripVisualTexAlignmentMarkers } from "../../../../shared/visualtex-app/src/editor/alignmentMarkers";
import type { LatexCodeFormat } from "../../../../shared/visualtex-app/src/types/formula";
import {
  serializeFormulaEditorDocument,
  type FormulaEditorDocument,
} from "../shared/formulaEditorDocument";
import { latexToMathMl } from "../../../../shared/visualtex-app/src/export/runtime";

// Word equations are written identically on Windows and macOS
// (docs/equation-format-v2.md). The OMML writer is the Rust port of Windows'
// C# writer, so its MathML input must be the MathML Windows' Office dialog
// produces: the same render LaTeX, converted by the same shared TypeScript
// (export/runtime.ts latexToMathMl).

// Windows OfficeDialogApp: serializeFormulaEditorRenderDocument.
const structuredRenderFormats = new Set<LatexCodeFormat>([
  "equation",
  "equation-star",
  "align",
  "align-star",
  "aligned",
  "gather",
  "gather-star",
  "multline",
  "multline-star",
  "equation-split",
  "equation-star-split",
]);

/**
 * The LaTeX Windows renders for an Office formula: wrapper-only source formats
 * (`$…$`, `\(…\)`, `\[…\]`, raw lines) render their body lines; structured
 * environments render their canonical environment, `aligned` without `\[…\]`.
 */
export function wordEquationRenderLatex(document: FormulaEditorDocument) {
  if (!structuredRenderFormats.has(document.codeFormat)) {
    return formatLatexLines(
      document.lines.map((line) => stripVisualTexAlignmentMarkers(line.latex)),
      "raw",
    );
  }
  const canonical = serializeFormulaEditorDocument(document);
  if (document.codeFormat !== "aligned") return canonical;
  return canonical
    .replace(/^\s*\\\[\s*/, "")
    .replace(/\s*\\\]\s*$/, "")
    .trim();
}

/** Windows' MathML for a Word formula (input of the shared OMML writer). */
export function wordEquationMathMl(
  document: FormulaEditorDocument,
  displayMode: "inline" | "block",
) {
  return latexToMathMl(
    wordEquationRenderLatex(document),
    displayMode === "block",
  );
}
