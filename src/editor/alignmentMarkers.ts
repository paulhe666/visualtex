import type { LatexCodeFormat } from "../types/formula";

/**
 * MathLive renders a bare `&` as a visible ampersand outside a tabular
 * environment. VisualTeX keeps explicit align points as a zero-width classed
 * atom instead, then converts them back to real `&` tokens when serializing an
 * align-like LaTeX environment.
 */
export const VISUALTEX_ALIGNMENT_MARKER_LATEX =
  "\\class{visualtex-align-marker}{\\kern0pt}";
export const VISUALTEX_ALIGNMENT_MARKER_CLASS = "visualtex-align-marker";

const alignmentFormats = new Set<LatexCodeFormat>([
  "align",
  "align-star",
  "aligned",
  "equation-split",
  "equation-star-split",
]);

export function usesExplicitAlignmentPoints(format: LatexCodeFormat) {
  return alignmentFormats.has(format);
}

export function hasVisualTexAlignmentMarker(latex: string) {
  return latex.includes(VISUALTEX_ALIGNMENT_MARKER_LATEX);
}

export function stripVisualTexAlignmentMarkers(latex: string) {
  return latex.split(VISUALTEX_ALIGNMENT_MARKER_LATEX).join("");
}

export function restoreLatexAlignmentMarkers(latex: string) {
  return latex.split(VISUALTEX_ALIGNMENT_MARKER_LATEX).join("&");
}
