import { VISUALTEX_ALIGNMENT_MARKER_LATEX } from "../editor/alignmentMarkers.ts";

const MULTILINE_ENVIRONMENTS = new Set([
  "align",
  "align*",
  "aligned",
  "alignedat",
  "alignedat*",
  "gather",
  "gather*",
  "gathered",
  "multline",
  "multline*",
  "split",
]);

const PROTECTED_ROW_ENVIRONMENTS = new Set([
  "array",
  "cases",
  "matrix",
  "matrix*",
  "pmatrix",
  "bmatrix",
  "Bmatrix",
  "vmatrix",
  "Vmatrix",
  "smallmatrix",
]);

const PENDING_OCR_KEY = "__visualtexPendingMultilineOcrFormula";
const PENDING_LIFETIME_MS = 30_000;

export interface PendingMultilineOcrFormula {
  joinedLatex: string;
  lines: string[];
  expiresAt: number;
}

type OcrGlobal = typeof globalThis & {
  [PENDING_OCR_KEY]?: PendingMultilineOcrFormula;
};

function stripWholeMathDelimiter(value: string): string {
  let result = value.trim();
  if (result.startsWith("\\[") && result.endsWith("\\]")) {
    result = result.slice(2, -2).trim();
  } else if (result.startsWith("$$") && result.endsWith("$$")) {
    result = result.slice(2, -2).trim();
  } else if (
    result.length >= 2 &&
    result.startsWith("$") &&
    result.endsWith("$") &&
    !result.startsWith("$$")
  ) {
    result = result.slice(1, -1).trim();
  }
  return result;
}

function unwrapWholeEnvironment(value: string): string {
  let result = value.trim();
  for (;;) {
    const match = result.match(/^\\begin\{([^}]+)\}([\s\S]*)\\end\{\1\}$/);
    if (!match || !MULTILINE_ENVIRONMENTS.has(match[1])) return result;
    result = match[2].trim();
  }
}

function readEnvironmentToken(
  source: string,
  offset: number,
): { kind: "begin" | "end"; name: string; end: number } | null {
  const match = source.slice(offset).match(/^\\(begin|end)\{([^}]+)\}/);
  if (!match) return null;
  return {
    kind: match[1] as "begin" | "end",
    name: match[2],
    end: offset + match[0].length,
  };
}

function isEscaped(source: string, offset: number): boolean {
  let slashes = 0;
  for (let index = offset - 1; index >= 0 && source[index] === "\\"; index -= 1) {
    slashes += 1;
  }
  return slashes % 2 === 1;
}

function splitTopLevelRows(source: string): string[] {
  const rows: string[] = [];
  const environments: string[] = [];
  let braceDepth = 0;
  let rowStart = 0;

  const protectedEnvironmentActive = () =>
    environments.some((name) => PROTECTED_ROW_ENVIRONMENTS.has(name));

  const pushRow = (end: number) => {
    const row = source.slice(rowStart, end).trim();
    if (row) rows.push(row);
  };

  for (let index = 0; index < source.length; index += 1) {
    const environment =
      source[index] === "\\" ? readEnvironmentToken(source, index) : null;
    if (environment) {
      if (environment.kind === "begin") {
        environments.push(environment.name);
      } else {
        const reverseIndex = environments.lastIndexOf(environment.name);
        if (reverseIndex >= 0) environments.splice(reverseIndex, 1);
      }
      index = environment.end - 1;
      continue;
    }

    const character = source[index];
    if (character === "{" && !isEscaped(source, index)) {
      braceDepth += 1;
      continue;
    }
    if (character === "}" && !isEscaped(source, index)) {
      braceDepth = Math.max(0, braceDepth - 1);
      continue;
    }

    if (braceDepth !== 0 || protectedEnvironmentActive()) continue;

    if (character === "\n" || character === "\r") {
      pushRow(index);
      if (character === "\r" && source[index + 1] === "\n") index += 1;
      rowStart = index + 1;
      continue;
    }

    if (character === "\\" && source[index + 1] === "\\") {
      pushRow(index);
      index += 1;
      if (source[index + 1] === "[") {
        const optionEnd = source.indexOf("]", index + 2);
        if (optionEnd >= 0) index = optionEnd;
      }
      rowStart = index + 1;
    }
  }

  pushRow(source.length);
  return rows;
}

function encodeTopLevelAlignmentMarkers(value: string): string {
  let result = "";
  let braceDepth = 0;
  const environments: string[] = [];
  for (let index = 0; index < value.length; index += 1) {
    const environment =
      value[index] === "\\" ? readEnvironmentToken(value, index) : null;
    if (environment) {
      const token = value.slice(index, environment.end);
      result += token;
      if (environment.kind === "begin") environments.push(environment.name);
      else {
        const position = environments.lastIndexOf(environment.name);
        if (position >= 0) environments.splice(position, 1);
      }
      index = environment.end - 1;
      continue;
    }
    const character = value[index];
    if (character === "{" && !isEscaped(value, index)) braceDepth += 1;
    else if (character === "}" && !isEscaped(value, index)) {
      braceDepth = Math.max(0, braceDepth - 1);
    }
    if (
      character === "&" &&
      braceDepth === 0 &&
      environments.length === 0 &&
      !isEscaped(value, index)
    ) {
      result += VISUALTEX_ALIGNMENT_MARKER_LATEX;
      continue;
    }
    result += character;
  }
  return result.trim();
}

/**
 * Converts a multi-row OCR result into independently editable VisualTeX formula
 * lines. Matrix/cases/array rows stay inside one formula because their row breaks
 * are structural, while align/gather/split wrappers are unwrapped and split.
 */
export function splitOcrLatexIntoFormulaLines(rawLatex: string): string[] {
  const stripped = unwrapWholeEnvironment(stripWholeMathDelimiter(rawLatex));
  const rows = splitTopLevelRows(stripped)
    .map(encodeTopLevelAlignmentMarkers)
    .filter(Boolean);
  return rows.length > 0 ? rows : [stripped.trim()].filter(Boolean);
}

/**
 * Normalizes the OCR payload shown/inserted by existing consumers and records the
 * exact multi-line result for the editor-store bridge. The visible value uses real
 * newlines, so previews and copied OCR text remain readable rather than containing
 * a private sentinel.
 */
export function prepareOcrLatexForEditor(rawLatex: string): string {
  const lines = splitOcrLatexIntoFormulaLines(rawLatex);
  const joinedLatex = lines.join("\n");
  const target = globalThis as OcrGlobal;
  if (lines.length > 1) {
    target[PENDING_OCR_KEY] = {
      joinedLatex,
      lines,
      expiresAt: Date.now() + PENDING_LIFETIME_MS,
    };
  }
  return joinedLatex;
}

export function peekPendingMultilineOcrFormula(): PendingMultilineOcrFormula | null {
  const target = globalThis as OcrGlobal;
  const pending = target[PENDING_OCR_KEY];
  if (!pending) return null;
  if (pending.expiresAt < Date.now()) {
    delete target[PENDING_OCR_KEY];
    return null;
  }
  return pending;
}

export function clearPendingMultilineOcrFormula(): void {
  delete (globalThis as OcrGlobal)[PENDING_OCR_KEY];
}
