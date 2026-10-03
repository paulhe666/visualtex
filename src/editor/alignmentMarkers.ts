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

const inferredRelationCommands = new Set([
  "approx",
  "asymp",
  "coloneqq",
  "cong",
  "equiv",
  "ge",
  "geq",
  "gets",
  "iff",
  "implies",
  "in",
  "le",
  "leftarrow",
  "Leftarrow",
  "leftrightarrow",
  "Leftrightarrow",
  "leq",
  "longleftarrow",
  "Longleftarrow",
  "longleftrightarrow",
  "Longleftrightarrow",
  "longmapsto",
  "longrightarrow",
  "Longrightarrow",
  "mapsto",
  "mid",
  "ne",
  "neq",
  "notin",
  "parallel",
  "perp",
  "propto",
  "rightarrow",
  "Rightarrow",
  "simeq",
  "sim",
  "subset",
  "subseteq",
  "supset",
  "supseteq",
  "to",
]);

function markerCharacterIsEscaped(source: string, index: number) {
  let slashCount = 0;
  for (let cursor = index - 1; cursor >= 0 && source[cursor] === "\\"; cursor -= 1) {
    slashCount += 1;
  }
  return slashCount % 2 === 1;
}

/**
 * Normalizes editable align rows to VisualTeX's zero-width marker form.
 * Explicit top-level `&` tokens are always preserved. If an edit operation
 * replaced a whole row and thereby removed its marker, the first top-level
 * relation is used as the conservative recovery point.
 */
export function ensureVisualTexAlignmentMarkers(
  latex: string,
  inferRelation = true,
) {
  const source = String(latex ?? "");
  let result = "";
  let braceDepth = 0;
  const environments: string[] = [];
  let hasTopLevelMarker = false;
  let inferredRelationIndex: number | null = null;

  for (let index = 0; index < source.length; index += 1) {
    const topLevel = braceDepth === 0 && environments.length === 0;
    if (source.startsWith(VISUALTEX_ALIGNMENT_MARKER_LATEX, index)) {
      result += VISUALTEX_ALIGNMENT_MARKER_LATEX;
      if (topLevel) hasTopLevelMarker = true;
      index += VISUALTEX_ALIGNMENT_MARKER_LATEX.length - 1;
      continue;
    }

    if (source[index] === "\\") {
      const environment = source
        .slice(index)
        .match(/^\\(begin|end)\{([A-Za-z]+\*?)\}/);
      if (environment) {
        result += environment[0];
        if (environment[1] === "begin") {
          environments.push(environment[2]);
        } else {
          const matchingIndex = environments.lastIndexOf(environment[2]);
          if (matchingIndex >= 0) environments.splice(matchingIndex, 1);
        }
        index += environment[0].length - 1;
        continue;
      }
      if (topLevel && inferredRelationIndex === null) {
        const command = source.slice(index + 1).match(/^[A-Za-z]+/)?.[0];
        if (command && inferredRelationCommands.has(command)) {
          inferredRelationIndex = result.length;
        }
      }
    }

    const character = source[index];
    if (
      character === "&" &&
      topLevel &&
      !markerCharacterIsEscaped(source, index)
    ) {
      result += VISUALTEX_ALIGNMENT_MARKER_LATEX;
      hasTopLevelMarker = true;
      continue;
    }
    if (
      topLevel &&
      inferredRelationIndex === null &&
      (character === "=" || character === "<" || character === ">") &&
      !markerCharacterIsEscaped(source, index)
    ) {
      inferredRelationIndex = result.length;
    }

    result += character;
    if (character === "{" && !markerCharacterIsEscaped(source, index)) {
      braceDepth += 1;
    } else if (
      character === "}" &&
      !markerCharacterIsEscaped(source, index)
    ) {
      braceDepth = Math.max(0, braceDepth - 1);
    }
  }

  if (hasTopLevelMarker || !inferRelation || inferredRelationIndex === null) {
    return result;
  }
  return (
    result.slice(0, inferredRelationIndex) +
    VISUALTEX_ALIGNMENT_MARKER_LATEX +
    result.slice(inferredRelationIndex)
  );
}
