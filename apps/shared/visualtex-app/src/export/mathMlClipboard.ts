// MathML on the clipboard (the formula rows' and the document's right-click menus): formulas as the Office add-ins
// insert them (latexToMathMl: MathJax with VisualTeX's commands and custom symbols), written as standard
// Presentation MathML that Word, MathType, OneNote and web editors take — one <math> per formula, its LaTeX kept as
// the TeX annotation (pasted back into a LaTeX tool, it is the source again). The clipboard gets the MathML as text:
// Word turns a pasted <math> into an equation.

import { restoreLatexAlignmentMarkers, stripVisualTexAlignmentMarkers } from "../editor/alignmentMarkers";

export interface MathMlFormula {
  /** the formula's LaTeX as the editor stores it (VisualTeX's commands, custom symbols) */
  latex: string;
  /** a display formula (block) rather than an inline one */
  display: boolean;
  /** the LaTeX the annotation carries (default: `latex`) */
  tex?: string;
}

const MATHML_NS = "http://www.w3.org/1998/Math/MathML";

const escapeXml = (text: string) => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

/**
 * MathJax's MathML as plain Presentation MathML: none of its own data attributes, the namespace and the display
 * attribute on <math>, the content in <semantics> with the TeX annotation.
 */
export function cleanMathMl(mathMl: string, display: boolean, tex: string): string {
  const out = mathMl.trim().replace(/\s+data-[\w-]+="[^"]*"/g, "").replace(/\s+class="MJX-[^"]*"/g, "");
  const open = /^<math\b[^>]*>/.exec(out);
  const close = out.lastIndexOf("</math>");
  if (!open || close < open[0].length) throw new Error("not MathML");
  const inner = out.slice(open[0].length, close).replace(/^\s*\n|\n\s*$/g, "");
  const attrs = `xmlns="${MATHML_NS}"${display ? ' display="block"' : ""}`;
  // (the presentation as one row, <semantics> taking one element before its annotations; MathJax's indentation kept,
  // two levels deeper)
  const body = inner
    .split("\n")
    .map((line) => (line.trim() ? `    ${/^\s/.test(line) ? line : `  ${line}`}` : line))
    .join("\n");
  const annotation = `<annotation encoding="application/x-tex">${escapeXml(tex.trim())}</annotation>`;
  return `<math ${attrs}>\n  <semantics>\n    <mrow>\n${body}\n    </mrow>\n    ${annotation}\n  </semantics>\n</math>`;
}

/** The formulas as MathML text, one <math> after another (a blank line between). */
export async function formulasToMathMl(formulas: readonly MathMlFormula[]): Promise<string> {
  const { latexToMathMl } = await import("./runtime");
  const parts: string[] = [];
  // the editor's alignment points: the rows' "&" in aligned rows, nothing elsewhere (as Copy LaTeX writes them)
  const aligned = (latex: string) =>
    /^\s*\\begin\{aligned\}/.test(latex) ? restoreLatexAlignmentMarkers(latex) : stripVisualTexAlignmentMarkers(latex);
  for (const formula of formulas) {
    const latex = aligned(formula.latex).trim();
    if (!latex) continue;
    parts.push(cleanMathMl(latexToMathMl(latex, formula.display), formula.display, aligned(formula.tex ?? formula.latex)));
  }
  return parts.join("\n\n");
}

/** The formulas copied as MathML text; how many there were (0: nothing copied). */
export async function copyFormulasAsMathMl(formulas: readonly MathMlFormula[]): Promise<number> {
  const text = await formulasToMathMl(formulas);
  if (!text) return 0;
  await navigator.clipboard.writeText(text);
  return text.split("<math ").length - 1;
}
