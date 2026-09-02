import {
  formatLatexLines,
  isLatexCodeFormat,
  parseLatexSource,
} from "../../clipboard/LatexCopyService";
import { createUuid } from "../../runtime/browserCompatibility";
import { unwrapSingleLatexDisplayMath } from "../../math/latexEnvironment";
import {
  ensureVisualTexAlignmentMarkers,
  usesExplicitAlignmentPoints,
} from "../../editor/alignmentMarkers";
import type { LatexCodeFormat } from "../../types/formula";

export interface FormulaEditorLine {
  id: string;
  latex: string;
}

export interface FormulaEditorDocument {
  lines: FormulaEditorLine[];
  codeFormat: LatexCodeFormat;
}

function normalizeAlignmentLines(
  lines: FormulaEditorLine[],
  codeFormat: LatexCodeFormat,
) {
  if (!usesExplicitAlignmentPoints(codeFormat)) return lines;
  return lines.map((line) => ({
    ...line,
    latex: ensureVisualTexAlignmentMarkers(line.latex),
  }));
}

function normalizeLogicalFormulaLineWhitespace(value: string) {
  const source = value.replace(/\r\n?/g, "\n");
  return source
    .replace(/[ \t]*\n[ \t]*/g, (match, offset: number) => {
      const before = source.slice(0, offset);
      const after = source.slice(offset + match.length);
      const previous = before.at(-1) ?? "";
      const next = after[0] ?? "";
      if (!previous || !next) return "";

      // MathLive does not reliably skip a literal space between consecutive
      // mandatory arguments. Turning `}\\n{` into `} {` can therefore parse a
      // valid `\\frac{numerator}{denominator}` as a fraction with an empty
      // denominator. Preserve a separator only when removing it would merge
      // two lexical words or extend a TeX control word.
      const trailingControlWord = /\\[A-Za-z@]+$/.test(before);
      const nextStartsControlWordCharacter = /^[A-Za-z@]/.test(after);
      const mergesTextWords =
        /[\p{L}\p{N}]$/u.test(before) && /^[\p{L}\p{N}]/u.test(after);
      return trailingControlWord && nextStartsControlWordCharacter
        ? " "
        : mergesTextWords
          ? " "
          : "";
    })
    .trim();
}

export function serializeFormulaEditorDocument(document: FormulaEditorDocument) {
  return formatLatexLines(
    document.lines.map((line) => line.latex),
    document.codeFormat,
  );
}

interface DetectedFormulaEnvironment {
  codeFormat: LatexCodeFormat;
  source: string;
}

function escapeRegExp(value: string) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function replaceEnvironment(
  source: string,
  original: string,
  replacement: string,
) {
  const escaped = escapeRegExp(original);
  return source
    .replace(
      new RegExp(
        `\\\\begin\\s*\\{${escaped}\\}(?:\\s*\\{[^{}]*\\})?`,
      ),
      `\\begin{${replacement}}`,
    )
    .replace(
      new RegExp(`\\\\end\\s*\\{${escaped}\\}`),
      `\\end{${replacement}}`,
    );
}

function detectFormulaEnvironment(source: string): DetectedFormulaEnvironment | null {
  const normalized = source
    .replace(/\r\n?/g, "\n")
    .replace(/\\(begin|end)\s*\{\s*([^{}]+?)\s*\}/g, "\\$1{$2}")
    .trim();
  if (!normalized) return null;

  const displayBody = unwrapSingleLatexDisplayMath(normalized);
  if (displayBody !== null) {
    const detectedBody = detectFormulaEnvironment(displayBody);
    if (detectedBody) return detectedBody;
    return {
      codeFormat: "equation-star",
      source: `\\begin{equation*}${displayBody}\\end{equation*}`,
    };
  }

  const equation = normalized.match(
    /^\\begin\s*\{(equation\*?)\}([\s\S]*)\\end\s*\{\1\}$/,
  );
  if (
    equation &&
    /\\begin\s*\{split\}[\s\S]*\\end\s*\{split\}/.test(equation[2])
  ) {
    return {
      codeFormat:
        equation[1] === "equation*"
          ? "equation-star-split"
          : "equation-split",
      source: normalized,
    };
  }

  const environment = normalized.match(
    /^\\begin\s*\{(align\*?|alignat\*?|flalign\*?|eqnarray\*?|aligned|alignedat|gather\*?|multline\*?|equation\*?|displaymath)\}(?:\s*\{[^{}]*\})?[\s\S]*\\end\s*\{\1\}$/,
  )?.[1];
  if (!environment) return null;

  switch (environment) {
    case "align":
      return { codeFormat: "align", source: normalized };
    case "align*":
      return { codeFormat: "align-star", source: normalized };
    case "alignat":
      return {
        codeFormat: "align",
        source: replaceEnvironment(normalized, "alignat", "align"),
      };
    case "alignat*":
      return {
        codeFormat: "align-star",
        source: replaceEnvironment(normalized, "alignat*", "align*"),
      };
    case "flalign":
    case "eqnarray":
      return {
        codeFormat: "align",
        source: replaceEnvironment(normalized, environment, "align"),
      };
    case "flalign*":
    case "eqnarray*":
      return {
        codeFormat: "align-star",
        source: replaceEnvironment(normalized, environment, "align*"),
      };
    case "aligned":
      return { codeFormat: "aligned", source: normalized };
    case "alignedat":
      return {
        codeFormat: "aligned",
        source: replaceEnvironment(normalized, "alignedat", "aligned"),
      };
    case "gather":
      return { codeFormat: "gather", source: normalized };
    case "gather*":
      return { codeFormat: "gather-star", source: normalized };
    case "multline":
      return { codeFormat: "multline", source: normalized };
    case "multline*":
      return { codeFormat: "multline-star", source: normalized };
    case "equation":
      return { codeFormat: "equation", source: normalized };
    case "equation*":
    case "displaymath":
      return {
        codeFormat: "equation-star",
        source:
          environment === "displaymath"
            ? replaceEnvironment(normalized, "displaymath", "equation*")
            : normalized,
      };
    default:
      return null;
  }
}

export function normalizeFormulaEditorDocument(
  lines: FormulaEditorLine[],
  codeFormat: unknown,
): FormulaEditorDocument {
  const fallbackFormat: LatexCodeFormat = isLatexCodeFormat(codeFormat)
    ? codeFormat
    : "raw";
  const sourceLines = Array.isArray(lines) ? lines : [];
  const safeLines = sourceLines.length
    ? sourceLines.map((line) => ({
        id: line.id || createUuid(),
        latex:
          typeof line.latex === "string"
            ? normalizeLogicalFormulaLineWhitespace(line.latex)
            : "",
      }))
    : [{ id: createUuid(), latex: "" }];

  if (safeLines.length !== 1) {
    return {
      lines: normalizeAlignmentLines(safeLines, fallbackFormat),
      codeFormat: fallbackFormat,
    };
  }

  const detected = detectFormulaEnvironment(safeLines[0].latex);
  if (!detected) {
    return {
      lines: normalizeAlignmentLines(safeLines, fallbackFormat),
      codeFormat: fallbackFormat,
    };
  }
  // A caller may already own a normalized logical document. For example, an
  // imported `equation` can legitimately contain one `aligned` environment as
  // its body. Re-detecting that inner environment on a second normalization
  // pass would silently replace the outer `equation`, lose its numbering
  // semantics, and then split source-formatting newlines into visual rows.
  // Only unwrap a detected environment when the caller supplied raw source or
  // when the detected wrapper agrees with the caller's existing code format.
  if (fallbackFormat !== "raw" && detected.codeFormat !== fallbackFormat) {
    return {
      lines: normalizeAlignmentLines(safeLines, fallbackFormat),
      codeFormat: fallbackFormat,
    };
  }

  const parsed = parseLatexSource(detected.source, detected.codeFormat)
    .map((latex) => latex.trim())
    .filter(Boolean);
  if (!parsed.length) {
    return { lines: safeLines, codeFormat: fallbackFormat };
  }

  return {
    codeFormat: detected.codeFormat,
    lines: normalizeAlignmentLines(
      parsed.map((latex, index) => ({
        id: index === 0 ? safeLines[0].id : createUuid(),
        latex,
      })),
      detected.codeFormat,
    ),
  };
}
