/**
 * A FormulaLine is one logical editor row. Physical newlines can still appear
 * inside a complete TeX environment copied from document import metadata.
 * TeX treats those newlines as whitespace; explicit `\\` commands carry the
 * actual mathematical row structure. Never truncate at the first newline.
 */
export function normalizeFormulaLinePhysicalWhitespace(latex: string) {
  return latex
    .replace(/\r\n?/g, "\n")
    .replace(/[ \t]*\n+[ \t]*/g, " ");
}
