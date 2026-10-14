import { findLatexFormulaSpans } from "../documentImport/documentImportParser";

export type WordLatexRedrawDisplayMode = "inline" | "block";

export interface WordLatexRedrawSpan {
  start: number;
  end: number;
  sourceText: string;
  latex: string;
  displayMode: WordLatexRedrawDisplayMode;
}

function isWordMathematicalAlphanumeric(character: string) {
  const codePoint = character.codePointAt(0) ?? 0;
  if (
    (codePoint >= 0x1d400 && codePoint <= 0x1d7ff) ||
    (codePoint >= 0x2100 && codePoint <= 0x214f)
  ) {
    const normalized = character.normalize("NFKC");
    return /^[A-Za-z0-9]$/.test(normalized);
  }
  return false;
}

/**
 * Word can return selected LaTeX using Unicode mathematical-alphanumeric
 * glyphs and duplicate the command backslash before those glyphs. Keep the
 * exact Word text for range validation, but restore a standard LaTeX body for
 * MathJax and OMML conversion.
 */
function normalizeWordLatexBody(body: string) {
  const characters = Array.from(body);
  let normalized = "";
  for (let index = 0; index < characters.length; index += 1) {
    const character = characters[index];
    if (
      character === "\\" &&
      characters[index + 1] === "\\" &&
      isWordMathematicalAlphanumeric(characters[index + 2] ?? "")
    ) {
      normalized += "\\";
      index += 1;
      continue;
    }
    if (isWordMathematicalAlphanumeric(character)) {
      normalized += character.normalize("NFKC");
      continue;
    }
    normalized += character === "−" ? "-" : character;
  }
  return normalized;
}

/** Redraw and bulk import share delimiter, environment, comment and code rules.
 * Keep original UTF-16 offsets for Word even when its styled text is normalized.
 */
export function findWindowsWordLatexRedrawSpans(source: string): WordLatexRedrawSpan[] {
  return findLatexFormulaSpans(source).map((span) => ({
    start: span.start,
    end: span.end,
    sourceText: span.sourceText,
    latex: normalizeWordLatexBody(span.latex),
    displayMode: span.displayMode,
  }));
}
