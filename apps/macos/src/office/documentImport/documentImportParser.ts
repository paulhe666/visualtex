import { createUuid } from "../../runtime/browserCompatibility.ts";

export type DocumentImportSourceKind = "auto" | "markdown" | "latex";
export type DocumentFormulaDisplayMode = "inline" | "block";
export type DocumentFormulaOutputKind = "omml" | "image";
export type DocumentParagraphStyle =
  | "normal"
  | "heading1"
  | "heading2"
  | "heading3"
  | "heading4"
  | "quote"
  | "code";
export type DocumentListKind = "none" | "bullet" | "number";
export type DocumentParagraphAlignment = "left" | "center" | "right" | "justify";

export interface DocumentParagraphMetadata {
  paragraphId?: string;
  paragraphStyle?: DocumentParagraphStyle;
  paragraphAlignment?: DocumentParagraphAlignment;
  listKind?: DocumentListKind;
  listLevel?: number;
  paragraphStart?: boolean;
  paragraphEnd?: boolean;
}

export interface DocumentTextBlock extends DocumentParagraphMetadata {
  id: string;
  kind: "text";
  text: string;
}

export interface DocumentFormulaBlock extends DocumentParagraphMetadata {
  id: string;
  kind: "formula";
  latex: string;
  sourceText?: string;
  sourceStart?: number;
  sourceEnd?: number;
  displayMode: DocumentFormulaDisplayMode;
  numbered: boolean;
  fontSizePt: number;
}

export type DocumentImportBlock = DocumentTextBlock | DocumentFormulaBlock;

interface ExtractedFormula {
  token: string;
  block: DocumentFormulaBlock;
}

interface ExtractedLiteral {
  token: string;
  text: string;
}

interface ParagraphContext {
  style: DocumentParagraphStyle;
  alignment: DocumentParagraphAlignment;
  listKind: DocumentListKind;
  listLevel: number;
}

interface TheoremEnvironmentDefinition {
  environment: string;
  label: string;
  numbered: boolean;
  counterName: string;
  bodyStyle: DocumentParagraphStyle;
}

interface TheoremMarkerPayload {
  label: string;
  note: string;
  numbered: boolean;
  counterName: string;
  bodyStyle: DocumentParagraphStyle;
}

const blockEnvironmentNames = new Set([
  "equation",
  "equation*",
  "align",
  "align*",
  "alignat",
  "alignat*",
  "flalign",
  "flalign*",
  "eqnarray",
  "eqnarray*",
  "gather",
  "gather*",
  "multline",
  "multline*",
  "displaymath",
]);
const inlineEnvironmentNames = new Set(["math"]);
const protectedLatexEnvironmentNames = new Set([
  "verbatim",
  "verbatim*",
  "lstlisting",
  "minted",
  "comment",
]);
const structuredLatexEnvironmentNames = new Set([
  "document",
  "itemize",
  "enumerate",
  "quote",
  "quotation",
  "center",
  "flushleft",
  "flushright",
  "abstract",
  "description",
]);
const builtInTheoremDefinitions: TheoremEnvironmentDefinition[] = [
  ["theorem", "定理", true, "theorem", "quote"],
  ["lemma", "引理", true, "lemma", "quote"],
  ["proposition", "命题", true, "proposition", "quote"],
  ["corollary", "推论", true, "corollary", "quote"],
  ["definition", "定义", true, "definition", "quote"],
  ["axiom", "公理", true, "axiom", "quote"],
  ["conjecture", "猜想", true, "conjecture", "quote"],
  ["claim", "断言", true, "claim", "quote"],
  ["criterion", "判据", true, "criterion", "quote"],
  ["property", "性质", true, "property", "quote"],
  ["fact", "事实", true, "fact", "quote"],
  ["observation", "观察", true, "observation", "quote"],
  ["example", "例", true, "example", "quote"],
  ["exercise", "练习", true, "exercise", "quote"],
  ["problem", "问题", true, "problem", "quote"],
  ["question", "问题", true, "question", "quote"],
  ["remark", "注", false, "remark", "quote"],
  ["note", "注", false, "note", "quote"],
  ["notation", "记号", false, "notation", "quote"],
  ["case", "情形", false, "case", "quote"],
  ["proof", "证明", false, "proof", "normal"],
  ["solution", "解答", false, "solution", "normal"],
].map(([environment, label, numbered, counterName, bodyStyle]) => ({
  environment: environment as string,
  label: label as string,
  numbered: numbered as boolean,
  counterName: counterName as string,
  bodyStyle: bodyStyle as DocumentParagraphStyle,
}));
const literalLatexCommandPattern = /\\(?:documentclass|usepackage|RequirePackage|PassOptionsToPackage|newcommand|renewcommand|providecommand|DeclareRobustCommand|DeclareMathOperator\*?|newenvironment|renewenvironment|provideenvironment|newtheoremstyle|newtheorem|theoremstyle|swapnumbers|def|gdef|edef|xdef|let|newlength|setlength|addtolength|definecolor|colorlet|newcolumntype|DeclarePairedDelimiter\*?|DeclareDocumentCommand|NewDocumentCommand|RenewDocumentCommand|ProvideDocumentCommand)\b/g;
const formulaTokenPrefix = "\uE000VT_FORMULA_";
const formulaTokenSuffix = "\uE001";
const formulaTokenPattern = /\uE000VT_FORMULA_(\d+)\uE001/g;
const literalTokenPrefix = "\uE200VT_LITERAL_";
const literalTokenSuffix = "\uE201";
const markerPrefix = "\uE100VT_";
const markerSuffix = "\uE101";

function marker(value: string) {
  return `${markerPrefix}${value}${markerSuffix}`;
}

function leadingMarker(line: string) {
  const trimmed = line.trim();
  if (!trimmed.startsWith(markerPrefix)) return null;
  const end = trimmed.indexOf(markerSuffix, markerPrefix.length);
  if (end < 0) return null;
  return {
    value: trimmed.slice(markerPrefix.length, end),
    rest: trimmed.slice(end + markerSuffix.length),
  };
}

function theoremDefinitionsFromSource(source: string) {
  const definitions = new Map<string, TheoremEnvironmentDefinition>();
  for (const definition of builtInTheoremDefinitions) {
    definitions.set(definition.environment, { ...definition });
    definitions.set(`${definition.environment}*`, {
      ...definition,
      environment: `${definition.environment}*`,
      numbered: false,
    });
  }

  const declarationPattern =
    /\\newtheorem(\*)?\s*\{([^{}]+)\}\s*(?:\[([^\]]+)\]\s*)?\{([^{}]*)\}\s*(?:\[([^\]]+)\])?/g;
  for (
    let match = declarationPattern.exec(source);
    match;
    match = declarationPattern.exec(source)
  ) {
    const environment = match[2].trim();
    const sharedCounter = match[3]?.trim();
    const label = match[4].trim() || environment;
    if (!environment) continue;
    const builtIn = definitions.get(environment);
    definitions.set(environment, {
      environment,
      label,
      numbered: !match[1],
      counterName: sharedCounter || environment,
      bodyStyle: builtIn?.bodyStyle ?? "quote",
    });
  }
  return definitions;
}

function encodeTheoremMarker(payload: TheoremMarkerPayload) {
  return encodeURIComponent(JSON.stringify(payload));
}

function decodeTheoremMarker(value: string): TheoremMarkerPayload | null {
  try {
    const payload = JSON.parse(decodeURIComponent(value)) as Partial<TheoremMarkerPayload>;
    if (
      typeof payload.label !== "string" ||
      typeof payload.note !== "string" ||
      typeof payload.numbered !== "boolean" ||
      typeof payload.counterName !== "string" ||
      !["normal", "quote"].includes(payload.bodyStyle ?? "")
    ) {
      return null;
    }
    return payload as TheoremMarkerPayload;
  } catch {
    return null;
  }
}

function isEscaped(source: string, index: number) {
  let slashCount = 0;
  for (let cursor = index - 1; cursor >= 0 && source[cursor] === "\\"; cursor -= 1) {
    slashCount += 1;
  }
  return slashCount % 2 === 1;
}

interface ProtectedSourceRange {
  start: number;
  end: number;
}

function escapeRegExp(value: string) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function mergeProtectedRanges(ranges: ProtectedSourceRange[]) {
  const sorted = ranges
    .filter((range) => range.end > range.start)
    .sort((left, right) => left.start - right.start || left.end - right.end);
  const merged: ProtectedSourceRange[] = [];
  for (const range of sorted) {
    const previous = merged.at(-1);
    if (!previous || range.start > previous.end) {
      merged.push({ ...range });
    } else {
      previous.end = Math.max(previous.end, range.end);
    }
  }
  return merged;
}

function sourceIndexInsideRanges(ranges: ProtectedSourceRange[], index: number) {
  return ranges.some((range) => index >= range.start && index < range.end);
}

function findLatexCommandEnd(source: string, start: number) {
  let braceDepth = 0;
  let bracketDepth = 0;
  for (let cursor = start; cursor < source.length; cursor += 1) {
    const character = source[cursor];
    if (character === "\\" && cursor + 1 < source.length) {
      cursor += 1;
      continue;
    }
    if (character === "{") braceDepth += 1;
    else if (character === "}" && braceDepth > 0) braceDepth -= 1;
    else if (character === "[") bracketDepth += 1;
    else if (character === "]" && bracketDepth > 0) bracketDepth -= 1;
    else if (
      (character === "\n" || character === "\r") &&
      braceDepth === 0 &&
      bracketDepth === 0
    ) {
      let continuation = cursor + 1;
      while (
        continuation < source.length &&
        /[ \t\r\n]/.test(source[continuation])
      ) {
        continuation += 1;
      }
      if (source[continuation] === "{" || source[continuation] === "[") {
        continue;
      }
      return cursor;
    }
  }
  return source.length;
}

function findLatexEnvironmentEnd(source: string, start: number, environment: string) {
  const tokenPattern = new RegExp(
    `\\\\(begin|end)\\s*\\{${escapeRegExp(environment)}\\}`,
    "g",
  );
  tokenPattern.lastIndex = start;
  let depth = 0;
  for (let match = tokenPattern.exec(source); match; match = tokenPattern.exec(source)) {
    if (match[1] === "begin") depth += 1;
    else depth -= 1;
    if (depth === 0) return tokenPattern.lastIndex;
  }
  return source.length;
}

function latexMathSourceRanges(source: string) {
  const ranges: ProtectedSourceRange[] = [];
  let cursor = 0;
  while (cursor < source.length) {
    const delimiter = delimiterAt(source, cursor);
    if (!delimiter) {
      cursor += 1;
      continue;
    }
    const closing = findClosingDelimiter(source, delimiter);
    if (!closing) {
      cursor += Math.max(1, delimiter.opening.length);
      continue;
    }
    ranges.push({ start: cursor, end: closing.end });
    cursor = closing.end;
  }
  return mergeProtectedRanges(ranges);
}

function latexLiteralFallbackRanges(
  source: string,
  theoremEnvironmentNames: ReadonlySet<string> = new Set(),
) {
  const ranges: ProtectedSourceRange[] = [];
  const mathRanges = latexMathSourceRanges(source);

  literalLatexCommandPattern.lastIndex = 0;
  for (
    let match = literalLatexCommandPattern.exec(source);
    match;
    match = literalLatexCommandPattern.exec(source)
  ) {
    const end = findLatexCommandEnd(source, match.index);
    ranges.push({ start: match.index, end });
    literalLatexCommandPattern.lastIndex = Math.max(
      literalLatexCommandPattern.lastIndex,
      end,
    );
  }

  for (const [opening, closing] of [
    ["\\makeatletter", "\\makeatother"],
    ["\\ExplSyntaxOn", "\\ExplSyntaxOff"],
  ] as const) {
    let start = source.indexOf(opening);
    while (start >= 0) {
      const closingIndex = source.indexOf(closing, start + opening.length);
      const end = closingIndex < 0 ? source.length : closingIndex + closing.length;
      ranges.push({ start, end });
      start = source.indexOf(opening, end);
    }
  }

  const beginPattern = /\\begin\s*\{([^{}]+)\}/g;
  for (let match = beginPattern.exec(source); match; match = beginPattern.exec(source)) {
    const environment = match[1].trim();
    const supported =
      blockEnvironmentNames.has(environment) ||
      inlineEnvironmentNames.has(environment) ||
      structuredLatexEnvironmentNames.has(environment) ||
      theoremEnvironmentNames.has(environment);
    const literalEnvironment =
      protectedLatexEnvironmentNames.has(environment) || !supported;
    if (!literalEnvironment || sourceIndexInsideRanges(mathRanges, match.index)) continue;
    const end = findLatexEnvironmentEnd(source, match.index, environment);
    ranges.push({ start: match.index, end });
    beginPattern.lastIndex = end;
  }
  return mergeProtectedRanges(ranges);
}

function extractLatexLiteralFallbacks(
  source: string,
  theoremEnvironmentNames: ReadonlySet<string>,
) {
  const literals: ExtractedLiteral[] = [];
  const ranges = latexLiteralFallbackRanges(source, theoremEnvironmentNames);
  let output = "";
  let cursor = 0;
  for (const range of ranges) {
    output += source.slice(cursor, range.start);
    const token = `${literalTokenPrefix}${literals.length}${literalTokenSuffix}`;
    literals.push({ token, text: source.slice(range.start, range.end) });
    output += `\n${token}\n`;
    cursor = range.end;
  }
  output += source.slice(cursor);
  return { text: output, literals };
}

function latexProtectedRanges(
  source: string,
  theoremEnvironmentNames: ReadonlySet<string> = new Set(),
  includeComments = true,
) {
  const ranges = latexLiteralFallbackRanges(source, theoremEnvironmentNames);
  const mathRanges = latexMathSourceRanges(source);
  const verbPattern = /\\verb\*?([^\sA-Za-z0-9])[^\r\n]*?\1/g;
  for (let match = verbPattern.exec(source); match; match = verbPattern.exec(source)) {
    ranges.push({ start: match.index, end: verbPattern.lastIndex });
  }
  if (!includeComments) return mergeProtectedRanges(ranges);
  for (let cursor = 0; cursor < source.length; cursor += 1) {
    if (
      source[cursor] !== "%" ||
      isEscaped(source, cursor) ||
      sourceIndexInsideRanges(ranges, cursor) ||
      sourceIndexInsideRanges(mathRanges, cursor)
    ) {
      continue;
    }
    let end = cursor + 1;
    while (end < source.length && source[end] !== "\r" && source[end] !== "\n") {
      end += 1;
    }
    ranges.push({ start: cursor, end });
    cursor = end - 1;
  }
  return mergeProtectedRanges(ranges);
}

function markdownProtectedRanges(source: string) {
  const ranges: ProtectedSourceRange[] = [];
  let offset = 0;
  let fenceStart = -1;
  let fenceCharacter = "";
  let fenceLength = 0;

  while (offset < source.length) {
    let lineEnd = offset;
    while (
      lineEnd < source.length &&
      source[lineEnd] !== "\r" &&
      source[lineEnd] !== "\n"
    ) {
      lineEnd += 1;
    }
    let nextOffset = lineEnd;
    if (nextOffset < source.length) {
      nextOffset +=
        source[nextOffset] === "\r" && source[nextOffset + 1] === "\n"
          ? 2
          : 1;
    }
    const line = source.slice(offset, lineEnd);
    const fence = line.match(/^ {0,3}(`{3,}|~{3,})/);
    if (fenceStart < 0 && fence) {
      fenceStart = offset;
      fenceCharacter = fence[1][0];
      fenceLength = fence[1].length;
    } else if (
      fenceStart >= 0 &&
      new RegExp(`^ {0,3}${escapeRegExp(fenceCharacter)}{${fenceLength},}\\s*$`).test(
        line,
      )
    ) {
      ranges.push({ start: fenceStart, end: nextOffset });
      fenceStart = -1;
      fenceCharacter = "";
      fenceLength = 0;
    } else if (
      fenceStart < 0 &&
      /^(?:\t| {4})/.test(line) &&
      !/^\s*(?:[-+*]|\d+[.)])\s+/.test(line)
    ) {
      ranges.push({ start: offset, end: nextOffset });
    }
    offset = nextOffset;
  }
  if (fenceStart >= 0) ranges.push({ start: fenceStart, end: source.length });

  const fenced = mergeProtectedRanges(ranges);
  const inlineCodePattern = /(`+)([^\r\n]*?)\1/g;
  for (
    let match = inlineCodePattern.exec(source);
    match;
    match = inlineCodePattern.exec(source)
  ) {
    const insideBlock = fenced.some(
      (range) => match.index >= range.start && match.index < range.end,
    );
    if (!insideBlock) {
      ranges.push({ start: match.index, end: inlineCodePattern.lastIndex });
    }
  }
  return mergeProtectedRanges(ranges);
}

function protectedSourceRanges(
  source: string,
  sourceKind: DocumentImportSourceKind,
  theoremEnvironmentNames: ReadonlySet<string> = new Set(),
) {
  return sourceKind === "latex"
    ? latexProtectedRanges(source, theoremEnvironmentNames)
    : markdownProtectedRanges(source);
}

function latexCommentIndex(
  line: string,
  lineOffset: number,
  protectedRanges: ProtectedSourceRange[],
) {
  for (let index = 0; index < line.length; index += 1) {
    if (
      line[index] === "%" &&
      !isEscaped(line, index) &&
      !sourceIndexInsideRanges(protectedRanges, lineOffset + index)
    ) {
      return index;
    }
  }
  return -1;
}

function hasLatexCommentSignal(source: string) {
  return source.split(/\r\n|\r|\n/).some((line) => {
    const index = line.indexOf("%");
    return index >= 0 && !isEscaped(line, index) && line.slice(0, index).trim() === "";
  });
}

function stripLatexComments(
  source: string,
  theoremEnvironmentNames: ReadonlySet<string>,
) {
  const protectedRanges = latexProtectedRanges(source, theoremEnvironmentNames, false);
  const records: Array<{
    text: string;
    commentOnly: boolean;
    decorativeComment: boolean;
  }> = [];
  let lineOffset = 0;
  for (const line of source.split("\n")) {
    const commentIndex = latexCommentIndex(line, lineOffset, protectedRanges);
    const commentOnly =
      commentIndex >= 0 && line.slice(0, commentIndex).trim().length === 0;
    const commentBody = commentIndex >= 0 ? line.slice(commentIndex + 1) : "";
    records.push({
      text: commentIndex >= 0 ? line.slice(0, commentIndex) : line,
      commentOnly,
      decorativeComment:
        commentOnly && /(?:={3,}|-{3,}|_{3,}|\*{3,}|#{3,})/.test(commentBody),
    });
    lineOffset += line.length + 1;
  }

  // Section-divider comments are often wrapped manually as two physical
  // lines, for example `% ===== title %` followed by `=====`. The second line
  // is not meaningful document content; remove it only when it is directly
  // attached to a decorative comment so ordinary equations/text using `=` are
  // unaffected.
  const decorativeLinePattern = /^\s*[=\-_*#]{3,}\s*$/;
  return records
    .map((record, index) => {
      if (!decorativeLinePattern.test(record.text)) return record.text;
      const previous = records[index - 1];
      const next = records[index + 1];
      return previous?.decorativeComment || next?.decorativeComment ? "" : record.text;
    })
    .join("\n");
}

function unwrapSimpleLatexCommands(source: string) {
  let value = source;
  value = value
    .replace(/\\footnote\s*\{([^{}]*)\}/g, "（注：$1）")
    .replace(/\\url\s*\{([^{}]*)\}/g, "$1")
    .replace(/\\(?:eqref|ref|pageref)\s*\{([^{}]*)\}/g, "〔引用：$1〕")
    .replace(
      /\\cite(?:\[[^\]]*\])?\s*\{([^{}]*)\}/g,
      "〔引用：$1〕",
    )
    .replace(/\\colorbox\s*\{[^{}]*\}\s*\{([^{}]*)\}/g, "$1")
    .replace(/\\fcolorbox\s*\{[^{}]*\}\s*\{[^{}]*\}\s*\{([^{}]*)\}/g, "$1")
    .replace(/\\resizebox\s*\{[^{}]*\}\s*\{[^{}]*\}\s*\{([^{}]*)\}/g, "$1")
    .replace(/\\(?:makebox|framebox|parbox)\s*(?:\[[^\]]*\])?\s*\{([^{}]*)\}/g, "$1");
  const oneArgumentCommands = [
    "textbf",
    "textit",
    "emph",
    "underline",
    "sout",
    "st",
    "strike",
    "texttt",
    "textrm",
    "textsf",
    "textnormal",
    "mbox",
    "caption",
    "fbox",
  ];
  const commandPattern = new RegExp(
    `\\\\(?:${oneArgumentCommands.join("|")})\\s*\\{([^{}]*)\\}`,
    "g",
  );
  for (let pass = 0; pass < 8; pass += 1) {
    const next = value.replace(commandPattern, "$1");
    if (next === value) break;
    value = next;
  }
  return value.replace(
    /\\href\s*\{([^{}]*)\}\s*\{([^{}]*)\}/g,
    (_match, target: string, label: string) =>
      `${label}${target ? `（${target}）` : ""}`,
  );
}

function decodeMarkdownHtmlEntities(value: string) {
  const named: Record<string, string> = {
    amp: "&",
    apos: "'",
    gt: ">",
    lt: "<",
    nbsp: " ",
    quot: '"',
  };
  return value.replace(
    /&(?:#(\d+)|#x([0-9a-f]+)|([a-z][a-z0-9]+));/gi,
    (match, decimal: string | undefined, hexadecimal: string | undefined, name: string | undefined) => {
      if (decimal) {
        const codePoint = Number(decimal);
        return Number.isInteger(codePoint) && codePoint >= 0 && codePoint <= 0x10ffff
          ? String.fromCodePoint(codePoint)
          : match;
      }
      if (hexadecimal) {
        const codePoint = Number.parseInt(hexadecimal, 16);
        return Number.isInteger(codePoint) && codePoint >= 0 && codePoint <= 0x10ffff
          ? String.fromCodePoint(codePoint)
          : match;
      }
      return name ? named[name.toLowerCase()] ?? match : match;
    },
  );
}

function markdownTableCells(line: string) {
  const normalized = line.trim().replace(/^\|/, "").replace(/\|$/, "");
  const cells: string[] = [];
  let buffer = "";
  for (let index = 0; index < normalized.length; index += 1) {
    const character = normalized[index];
    if (character === "\\" && index + 1 < normalized.length) {
      buffer += character + normalized[index + 1];
      index += 1;
      continue;
    }
    if (character === "|") {
      cells.push(buffer.trim());
      buffer = "";
    } else {
      buffer += character;
    }
  }
  cells.push(buffer.trim());
  return cells;
}

function isMarkdownTableSeparator(line: string) {
  const cells = markdownTableCells(line);
  return (
    cells.length > 0 &&
    cells.every((cell) => /^:?-{3,}:?$/.test(cell.replace(/\s+/g, "")))
  );
}

/**
 * Normalize Markdown extensions into the flat paragraph/list/code protocol
 * understood by the native macOS Word writer. Formula delimiters are preserved
 * verbatim so the normal formula extraction pass still owns every math span.
 */
function normalizeMarkdownStructure(source: string) {
  const input = source.replace(/\r\n?/g, "\n");
  const originalLines = input.split("\n");
  const references = new Map<string, string>();
  const footnotes = new Map<string, string>();
  const retained: string[] = [];

  for (let index = 0; index < originalLines.length; index += 1) {
    const line = originalLines[index];
    const footnote = line.match(/^\s*\[\^([^\]]+)\]:\s*(.*)$/);
    if (footnote) {
      const body = [footnote[2]];
      while (
        index + 1 < originalLines.length &&
        /^(?: {2,}|\t)\S/.test(originalLines[index + 1])
      ) {
        index += 1;
        body.push(originalLines[index].trim());
      }
      footnotes.set(footnote[1].trim().toLowerCase(), body.join(" ").trim());
      continue;
    }
    const reference = line.match(
      /^\s*\[([^\]^][^\]]*)\]:\s*(\S+)(?:\s+["'(].*["')])?\s*$/,
    );
    if (reference) {
      references.set(reference[1].trim().toLowerCase(), reference[2]);
      continue;
    }
    retained.push(line);
  }

  const replaceReferences = (line: string) => {
    let next = line.replace(
      /!\[([^\]]*)\]\[([^\]]*)\]/g,
      (match, alt: string, key: string) => {
        const target = references.get((key || alt).trim().toLowerCase());
        return target ? `![${alt}](${target})` : match;
      },
    );
    next = next.replace(
      /\[([^\]]+)\]\[([^\]]*)\]/g,
      (match, label: string, key: string) => {
        const target = references.get((key || label).trim().toLowerCase());
        return target ? `[${label}](${target})` : match;
      },
    );
    next = next.replace(/\[\^([^\]]+)\]/g, (_match, key: string) => `〔注 ${key}〕`);
    next = next.replace(
      /\[(?!\^)([^\]]+)\](?![\[(])/g,
      (match, label: string, offset: number, whole: string) => {
        if (whole[offset - 1] === "!" || isEscaped(whole, offset)) return match;
        const target = references.get(label.trim().toLowerCase());
        return target ? `[${label}](${target})` : match;
      },
    );
    return next;
  };

  const output: string[] = [];
  let start = 0;
  if (retained[0]?.trim() === "---") {
    const end = retained
      .slice(1)
      .findIndex((line) => /^(?:---|\.\.\.)\s*$/.test(line.trim()));
    if (end >= 0) {
      output.push("**文档元数据**", "```yaml", ...retained.slice(1, end + 1), "```", "");
      start = end + 2;
    }
  }

  for (let index = start; index < retained.length; index += 1) {
    const raw = retained[index];
    const trimmed = raw.trim();

    const tildeFence = raw.match(/^(\s*)~{3,}(.*)$/);
    if (tildeFence) {
      output.push(`${tildeFence[1]}\`\`\`${tildeFence[2]}`);
      continue;
    }

    if (
      /^(?: {4}|\t)/.test(raw) &&
      !/^\s*(?:[-+*]|\d+[.)])\s+/.test(raw)
    ) {
      const code: string[] = [];
      while (index < retained.length) {
        const candidate = retained[index];
        if (!candidate.trim()) {
          code.push("");
          index += 1;
          continue;
        }
        if (!/^(?: {4}|\t)/.test(candidate)) break;
        code.push(candidate.startsWith("\t") ? candidate.slice(1) : candidate.slice(4));
        index += 1;
      }
      index -= 1;
      output.push("```", ...code, "```");
      continue;
    }

    if (
      index + 1 < retained.length &&
      trimmed &&
      /^ {0,3}(?:=+|-+)\s*$/.test(retained[index + 1])
    ) {
      const level = retained[index + 1].includes("=") ? "#" : "##";
      output.push(`${level} ${replaceReferences(trimmed)}`);
      index += 1;
      continue;
    }

    if (
      raw.includes("|") &&
      index + 1 < retained.length &&
      isMarkdownTableSeparator(retained[index + 1])
    ) {
      const header = markdownTableCells(raw);
      output.push(
        header.map((cell) => `**${replaceReferences(cell)}**`).join(" ｜ "),
        "",
      );
      index += 2;
      while (
        index < retained.length &&
        retained[index].includes("|") &&
        retained[index].trim()
      ) {
        output.push(
          markdownTableCells(retained[index])
            .map((cell) => replaceReferences(cell))
            .join(" ｜ "),
          "",
        );
        index += 1;
      }
      index -= 1;
      continue;
    }

    if (
      /^ {0,3}(?:\*\s*){3,}$/.test(raw) ||
      /^ {0,3}(?:-\s*){3,}$/.test(raw) ||
      /^ {0,3}(?:_\s*){3,}$/.test(raw)
    ) {
      output.push("────────────────────");
      continue;
    }

    const task = raw.match(/^(\s*[-+*]\s+)\[([ xX])\]\s+(.*)$/);
    if (task) {
      output.push(
        `${task[1]}${task[2].toLowerCase() === "x" ? "☒" : "☐"} ${replaceReferences(task[3])}`,
      );
      continue;
    }

    output.push(replaceReferences(raw));
  }

  if (footnotes.size) {
    output.push("");
    for (const [key, value] of footnotes) {
      output.push(`**注 ${key}：** ${replaceReferences(value)}`, "");
    }
  }
  return output.join("\n").replace(/<!--[\s\S]*?-->/g, "");
}

function cleanInlineMarkup(raw: string, sourceKind: DocumentImportSourceKind) {
  let value = raw.replace(/\r\n?/g, "\n");
  if (sourceKind === "latex") {
    value = unwrapSimpleLatexCommands(value)
      .replace(/\\(?:newline|linebreak)\b/g, "\n")
      .replace(/\\(?:qedhere|qed)\b/g, "□")
      .replace(/\\label\s*\{[^{}]*\}/g, "")
      .replace(/\\\\(?:\[[^\]]*\])?/g, "\n")
      .replace(/\\%/g, "%")
      .replace(/\\&/g, "&")
      .replace(/\\#/g, "#")
      .replace(/\\_/g, "_")
      .replace(/\\\$/g, "$")
      .replace(/\\\{/g, "{")
      .replace(/\\\}/g, "}")
      .replace(/~/g, " ");
  } else {
    value = value
      .replace(
        /!\[([^\]]*)\]\(([^)]*)\)/g,
        (_match, alt: string, target: string) =>
          `图片：${alt || "未命名"}${target ? `（${target}）` : ""}`,
      )
      .replace(
        /\[([^\]]+)\]\(([^)]*)\)/g,
        (_match, label: string, target: string) =>
          `${label}${target ? `（${target}）` : ""}`,
      )
      .replace(/`([^`]+)`/g, "$1")
      .replace(/\*\*([^*]+)\*\*/g, "$1")
      .replace(/__([^_]+)__/g, "$1")
      .replace(/(?<!\*)\*([^*\n]+)\*(?!\*)/g, "$1")
      .replace(/(?<!_)_([^_\n]+)_(?!_)/g, "$1")
      .replace(/~~([^~]+)~~/g, "$1");
  }
  if (sourceKind === "markdown") {
    value = decodeMarkdownHtmlEntities(value).replace(/<[^>]+>/g, "");
  }
  return value
    .replace(/[ \t]*\n[ \t]*/g, " ")
    .replace(/[ \t]{2,}/g, " ");
}

interface MathDelimiter {
  opening: string;
  closing: string;
  displayMode: DocumentFormulaDisplayMode;
  numbered: boolean;
  contentStart: number;
  environment?: string;
}

function delimiterAt(source: string, index: number): MathDelimiter | null {
  if (source.startsWith("$$", index) && !isEscaped(source, index)) {
    return {
      opening: "$$",
      closing: "$$",
      displayMode: "block",
      numbered: false,
      contentStart: index + 2,
    };
  }
  if (source.startsWith("\\[", index) && !isEscaped(source, index)) {
    return {
      opening: "\\[",
      closing: "\\]",
      displayMode: "block",
      numbered: false,
      contentStart: index + 2,
    };
  }
  if (source.startsWith("\\(", index) && !isEscaped(source, index)) {
    return {
      opening: "\\(",
      closing: "\\)",
      displayMode: "inline",
      numbered: false,
      contentStart: index + 2,
    };
  }
  if (source[index] === "$" && !isEscaped(source, index)) {
    return {
      opening: "$",
      closing: "$",
      displayMode: "inline",
      numbered: false,
      contentStart: index + 1,
    };
  }
  if (source.startsWith("\\begin", index) && !isEscaped(source, index)) {
    const match = source.slice(index).match(/^\\begin\s*\{([^{}]+)\}/);
    const environment = match?.[1]?.trim();
    if (
      match &&
      environment &&
      (blockEnvironmentNames.has(environment) || inlineEnvironmentNames.has(environment))
    ) {
      const displayMode = inlineEnvironmentNames.has(environment) ? "inline" : "block";
      return {
        opening: match[0],
        closing: `\\end{${environment}}`,
        displayMode,
        numbered:
          displayMode === "block" &&
          !environment.endsWith("*") &&
          environment !== "displaymath",
        contentStart: index + match[0].length,
        environment,
      };
    }
  }
  return null;
}

interface ClosingDelimiterMatch {
  start: number;
  end: number;
}

function findClosingDelimiter(
  source: string,
  delimiter: MathDelimiter,
): ClosingDelimiterMatch | null {
  let depth = 1;
  for (let cursor = delimiter.contentStart; cursor < source.length; cursor += 1) {
    if (isEscaped(source, cursor)) continue;
    // Delimiters in TeX comments are not field boundaries.
    if (source[cursor] === "%") {
      const end = source.slice(cursor).search(/[\r\n]/);
      if (end < 0) return null;
      cursor += end;
      continue;
    }
    if (delimiter.environment) {
      if (source[cursor] !== "\\") continue;
      const token = source.slice(cursor).match(/^\\(begin|end)\s*\{\s*([^{}]+?)\s*\}/);
      if (!token || token[2] !== delimiter.environment) continue;
      depth += token[1] === "begin" ? 1 : -1;
      if (depth === 0) return { start: cursor, end: cursor + token[0].length };
      cursor += token[0].length - 1;
      continue;
    }
    if (!source.startsWith(delimiter.closing, cursor)) continue;
    if (delimiter.closing === "$" && source.startsWith("$$", cursor)) {
      cursor += 1;
      continue;
    }
    return { start: cursor, end: cursor + delimiter.closing.length };
  }
  return null;
}


export interface DocumentLatexFormulaSpan {
  id: string;
  start: number;
  end: number;
  sourceText: string;
  latex: string;
  displayMode: DocumentFormulaDisplayMode;
  numbered: boolean;
  fontSizePt: number;
}

export function findLatexFormulaSpans(
  source: string,
  defaultFontSizePt = 12,
  requestedKind: DocumentImportSourceKind = "auto",
): DocumentLatexFormulaSpan[] {
  if (new TextEncoder().encode(source).byteLength > 5 * 1024 * 1024) {
    throw new Error("Word LaTeX redraw source exceeds the 5 MB limit.");
  }
  return scanLatexFormulaSpans(source, defaultFontSizePt, resolvedSourceKind(source, requestedKind));
}

function scanLatexFormulaSpans(
  source: string,
  defaultFontSizePt: number,
  sourceKind: DocumentImportSourceKind,
  theoremEnvironmentNames: ReadonlySet<string> = new Set(),
): DocumentLatexFormulaSpan[] {
  const protectedRanges = protectedSourceRanges(source, sourceKind, theoremEnvironmentNames);
  const fontSizePt = Math.min(512, Math.max(1, defaultFontSizePt));
  const spans: DocumentLatexFormulaSpan[] = [];
  let protectedIndex = 0;
  let cursor = 0;

  while (cursor < source.length) {
    while (
      protectedIndex < protectedRanges.length &&
      protectedRanges[protectedIndex].end <= cursor
    ) {
      protectedIndex += 1;
    }
    const protectedRange = protectedRanges[protectedIndex];
    if (
      protectedRange &&
      cursor >= protectedRange.start &&
      cursor < protectedRange.end
    ) {
      cursor = protectedRange.end;
      continue;
    }

    const delimiter = delimiterAt(source, cursor);
    if (!delimiter) {
      cursor += 1;
      continue;
    }
    const closing = findClosingDelimiter(source, delimiter);
    if (!closing) {
      cursor += delimiter.opening.length;
      continue;
    }
    if (
      protectedRange &&
      closing.start > protectedRange.start &&
      cursor < protectedRange.start
    ) {
      cursor = protectedRange.end;
      continue;
    }

    const sourceText = source.slice(cursor, closing.end);
    const latex = (delimiter.environment && delimiter.displayMode === "block"
      ? sourceText
      : source.slice(delimiter.contentStart, closing.start)
    )
      .trim()
      .replace(/^\s*\\displaystyle\s*/, "")
      .replace(/\\label\s*\{[^{}]*\}/g, "")
      .replace(/\\tag\*?\s*\{[^{}]*\}/g, "")
      .trim();
    if (latex) {
      spans.push({
        id: createUuid(),
        start: cursor,
        end: closing.end,
        sourceText,
        latex,
        displayMode: delimiter.displayMode,
        numbered: delimiter.displayMode === "block" && delimiter.numbered,
        fontSizePt,
      });
      if (spans.length > 1000) {
        throw new Error("Word LaTeX redraw supports at most 1000 formulas.");
      }
    }
    cursor = closing.end;
  }

  return joinSplitMathScopes(source, spans);
}

/** Recover legacy editor output that closed/reopened math inside an unfinished
 * TeX scope. Only whitespace/TeX spacing may separate the fragments; ordinary
 * prose, complete formulas, unmatched closers and code blocks are never joined.
 * Offsets still refer to the original Word text, including both delimiters.
 */
function joinSplitMathScopes(source: string, spans: DocumentLatexFormulaSpan[]) {
  const scopeState = (latex: string) => {
    const stack: string[] = [];
    for (let i = 0; i < latex.length; i += 1) {
      if (isEscaped(latex, i)) continue;
      if (latex[i] === "%") {
        while (i < latex.length && !/[\r\n]/.test(latex[i])) i += 1;
        continue;
      }
      if (latex[i] === "{") stack.push("group");
      if (latex[i] === "}" && stack.pop() !== "group") return null;
      if (latex[i] !== "\\") continue;
      const token = latex.slice(i).match(/^\\(left|right|mleft|mright)\b|^\\(begin|end)\s*\{\s*([^{}]+?)\s*\}/);
      if (!token) continue;
      const opening = token[1] === "left" || token[1] === "mleft" || token[2] === "begin";
      const name = token[2] ? `env:${token[3]}` : "fence";
      if (opening) stack.push(name);
      else if (stack.pop() !== name) return null;
      i += token[0].length - 1;
    }
    return stack;
  };
  const result: DocumentLatexFormulaSpan[] = [];
  for (let index = 0; index < spans.length; index += 1) {
    const first = spans[index];
    let latex = first.latex;
    let endIndex = index;
    let state = scopeState(latex);
    while (state?.length && endIndex + 1 < spans.length) {
      const next = spans[endIndex + 1];
      const gap = source.slice(spans[endIndex].end, next.start);
      if (next.displayMode !== first.displayMode ||
          !/^(?:[ \t]|\\(?:[ ,;:!]|quad\b|qquad\b))*$/.test(gap) ||
          !/^(?:\$|\\\()/u.test(first.sourceText) ||
          !/^(?:\$|\\\()/u.test(next.sourceText)) break;
      latex += gap + next.latex;
      endIndex += 1;
      state = scopeState(latex);
    }
    if (endIndex > index && state?.length === 0) {
      const end = spans[endIndex].end;
      result.push({ ...first, latex, end, sourceText: source.slice(first.start, end) });
      index = endIndex;
    } else result.push(first);
  }
  return result;
}

function extractFormulas(
  source: string,
  defaultFontSizePt: number,
  sourceKind: DocumentImportSourceKind,
  theoremEnvironmentNames: ReadonlySet<string> = new Set(),
) {
  const formulas: ExtractedFormula[] = [];
  let output = "";
  let cursor = 0;
  for (const span of scanLatexFormulaSpans(source, defaultFontSizePt, sourceKind, theoremEnvironmentNames)) {
    output += source.slice(cursor, span.start);
    const token = `${formulaTokenPrefix}${formulas.length}${formulaTokenSuffix}`;
    formulas.push({
      token,
      block: {
        id: span.id,
        kind: "formula",
        latex: span.latex,
        sourceText: span.sourceText,
        displayMode: span.displayMode,
        numbered: span.numbered,
        fontSizePt: span.fontSizePt,
      },
    });
    output += span.displayMode === "block" ? `\n${token}\n` : token;
    cursor = span.end;
  }
  output += source.slice(cursor);
  return { text: output, formulas };
}

function formulaFromToken(token: string, formulas: ExtractedFormula[]) {
  const match = token.match(/^\uE000VT_FORMULA_(\d+)\uE001$/);
  const index = match ? Number(match[1]) : -1;
  return Number.isInteger(index) ? formulas[index]?.block : undefined;
}

function literalFromToken(token: string, literals: ExtractedLiteral[]) {
  const match = token.match(/^\uE200VT_LITERAL_(\d+)\uE201$/);
  const index = match ? Number(match[1]) : -1;
  return Number.isInteger(index) ? literals[index] : undefined;
}

function defaultContext(): ParagraphContext {
  return { style: "normal", alignment: "left", listKind: "none", listLevel: 0 };
}

function headingStyle(level: number): DocumentParagraphStyle {
  if (level <= 1) return "heading1";
  if (level === 2) return "heading2";
  if (level === 3) return "heading3";
  return "heading4";
}

function replaceLatexTableEnvironmentFlat(source: string) {
  return source.replace(
    /\\begin\s*\{tabular\}\s*\{[^{}]*\}([\s\S]*?)\\end\s*\{tabular\}/gi,
    (_match, tableBody: string) => {
      const normalized = tableBody
        .replace(/\\(?:toprule|midrule|bottomrule|hline)\b/g, "")
        .replace(/\\multicolumn\{[^{}]*\}\{[^{}]*\}\{([^{}]*)\}/g, "$1")
        .replace(/\\multirow\{[^{}]*\}\{[^{}]*\}\{([^{}]*)\}/g, "$1");
      const rows = normalized
        .split(/\\\\(?:\[[^\]]*\])?/)
        .map((row) => row.trim())
        .filter(Boolean)
        .map((row) => row.replace(/(?<!\\)&/g, " ｜ ").replace(/\\&/g, "&"));
      return rows.length ? `\n${rows.join("\n\n")}\n\n` : "";
    },
  );
}

/** Expand only the LaTeX extensions that can be represented by the flat native
 * macOS Word protocol. Unknown commands/environments are intentionally left for
 * the existing literal-source fallback instead of being silently discarded. */
function normalizeLatexExtensionsFlat(source: string) {
  let body = source.replace(/\r\n?/g, "\n");
  const macros = new Map<string, { argumentCount: number; replacement: string }>();
  const preservedDefinitions: string[] = [];
  const preserveDefinition = (definition: string) => {
    const index = preservedDefinitions.push(definition) - 1;
    return `\uE300VT_MACRO_DEFINITION_${index}\uE301`;
  };
  const commandDefinition =
    /\\(?:newcommand|renewcommand|providecommand)\*?\s*\{\\([A-Za-z@]+)\}\s*(?:\[(\d+)\])?\s*\{((?:[^{}]|\{[^{}]*\})*)\}/g;
  body = body.replace(
    commandDefinition,
    (match, name: string, count: string | undefined, replacement: string) => {
      macros.set(name, {
        argumentCount: Math.max(0, Math.min(2, Number(count ?? "0") || 0)),
        replacement,
      });
      return preserveDefinition(match);
    },
  );
  body = body.replace(
    /\\def\\([A-Za-z@]+)\s*\{((?:[^{}]|\{[^{}]*\})*)\}/g,
    (match, name: string, replacement: string) => {
      macros.set(name, { argumentCount: 0, replacement });
      return preserveDefinition(match);
    },
  );

  for (let pass = 0; pass < 8; pass += 1) {
    let changed = false;
    for (const [name, macro] of macros) {
      const escaped = escapeRegExp(name);
      if (macro.argumentCount === 0) {
        body = body.replace(new RegExp(`\\\\${escaped}(?![A-Za-z@])`, "g"), () => {
          changed = true;
          return macro.replacement;
        });
      } else if (macro.argumentCount === 1) {
        body = body.replace(
          new RegExp(
            `\\\\${escaped}\\s*\\{((?:[^{}]|\\{[^{}]*\\})*)\\}`,
            "g",
          ),
          (_match, argument: string) => {
            changed = true;
            return macro.replacement.replace(/#1/g, argument);
          },
        );
      } else {
        body = body.replace(
          new RegExp(
            `\\\\${escaped}\\s*\\{((?:[^{}]|\\{[^{}]*\\})*)\\}\\s*\\{((?:[^{}]|\\{[^{}]*\\})*)\\}`,
            "g",
          ),
          (_match, first: string, second: string) => {
            changed = true;
            return macro.replacement.replace(/#1/g, first).replace(/#2/g, second);
          },
        );
      }
    }
    if (!changed) break;
  }

  body = body.replace(
    /\uE300VT_MACRO_DEFINITION_(\d+)\uE301/g,
    (_match, index: string) => preservedDefinitions[Number(index)] ?? "",
  );

  const preservedFloatEnvironments: string[] = [];
  const preserveFloatEnvironment = (environment: string) => {
    const index = preservedFloatEnvironments.push(environment) - 1;
    return `\uE302VT_FLOAT_ENVIRONMENT_${index}\uE303`;
  };
  body = body.replace(
    /\\begin\s*\{(figure\*?|table\*?)\}(?:\s*\[[^\]]*\])?([\s\S]*?)\\end\s*\{\1\}/gi,
    (match, environment: string, content: string) => {
      const isFigure = environment.toLowerCase().startsWith("figure");
      const containsMath =
        /\$|\\\(|\\\[|\\begin\s*\{(?:equation|align|gather|multline|math|displaymath)/.test(
          content,
        );
      if (isFigure && containsMath) return preserveFloatEnvironment(match);
      if (!isFigure && !/\\begin\s*\{tabular\}/i.test(content)) {
        return preserveFloatEnvironment(match);
      }
      return `\n${content}\n`;
    },
  );

  body = body
    .replace(
      /\\(part|chapter|section|subsection|subsubsection|paragraph|subparagraph)(\*?)\s*\[[^\]]*\]\s*\{/gi,
      "\\$1$2{",
    )
    .replace(/\\begin\s*\{(?:itemize|enumerate)\}\s*\[[^\]]*\]/gi, (match) =>
      match.replace(/\s*\[[^\]]*\]$/, ""),
    )
    .replace(/\\begin\s*\{minipage\}(?:\[[^\]]*\])?\s*\{[^{}]*\}/gi, "")
    .replace(/\\end\s*\{minipage\}/gi, "")
    .replace(
      /\\begin\s*\{thebibliography\}\s*\{[^{}]*\}/gi,
      "\\section*{参考文献}\n\\begin{itemize}\n",
    )
    .replace(/\\end\s*\{thebibliography\}/gi, "\n\\end{itemize}")
    .replace(
      /\\bibitem(?:\[([^\]]+)\])?\{([^{}]+)\}/gi,
      (_match, label: string | undefined, key: string) =>
        `\\item \\textbf{[${label || key}]} `,
    )
    .replace(/\\begin\s*\{description\}/gi, "\\begin{itemize}\n")
    .replace(/\\end\s*\{description\}/gi, "\n\\end{itemize}")
    .replace(/\\item\s*\[([^\]]+)\]/g, "\\item \\textbf{$1}：")
    .replace(/\\centering\b/gi, "")
    .replace(
      /\\includegraphics(?:\[[^\]]*\])?\{([^{}]+)\}/gi,
      "【图片：$1】",
    )
    .replace(
      /\\caption\s*\{((?:[^{}]|\{[^{}]*\})*)\}/gi,
      "\\textbf{图注：} $1",
    )
    .replace(/\\input\s*\{([^{}]+)\}/gi, "【外部文件：$1】")
    .replace(/\\include\s*\{([^{}]+)\}/gi, "【外部文件：$1】")
    .replace(/\\bibliography\s*\{([^{}]+)\}/gi, "【参考文献库：$1】");

  body = replaceLatexTableEnvironmentFlat(body);
  return body.replace(
    /\uE302VT_FLOAT_ENVIRONMENT_(\d+)\uE303/g,
    (_match, index: string) => preservedFloatEnvironments[Number(index)] ?? "",
  );
}

function normalizeTheoremEnvironments(
  source: string,
  definitions: ReadonlyMap<string, TheoremEnvironmentDefinition>,
) {
  const withOpenings = source.replace(
    /\\begin\s*\{([^{}]+)\}\s*(?:\[([^\]]*)\])?/g,
    (match, rawEnvironment: string, note = "") => {
      const environment = rawEnvironment.trim();
      const definition = definitions.get(environment);
      if (!definition) return match;
      const payload: TheoremMarkerPayload = {
        label: definition.label,
        note: note.trim(),
        numbered: definition.numbered,
        counterName: definition.counterName,
        bodyStyle: definition.bodyStyle,
      };
      return `\n${marker(`THEOREM_START:${encodeTheoremMarker(payload)}`)}\n`;
    },
  );
  return withOpenings.replace(
    /\\end\s*\{([^{}]+)\}/g,
    (match, rawEnvironment: string) =>
      definitions.has(rawEnvironment.trim())
        ? `\n${marker("THEOREM_END")}\n`
        : match,
  );
}

function normalizeLatexStructure(
  source: string,
  theoremDefinitions: ReadonlyMap<string, TheoremEnvironmentDefinition>,
) {
  return normalizeTheoremEnvironments(source, theoremDefinitions)
    .replace(/\\begin\s*\{document\}|\\end\s*\{document\}/g, "")
    .replace(/\\maketitle\b/g, "")
    .replace(/\\title\s*\{([^{}]*)\}/g, `\n${marker("HEADING:1")}$1\n`)
    .replace(/\\author\s*\{([^{}]*)\}/g, `\n${marker("CENTER")}$1\n`)
    .replace(/\\date\s*\{([^{}]*)\}/g, `\n${marker("CENTER")}$1\n`)
    .replace(/\\(?:part|chapter|section)\*?\s*\{([^{}]*)\}/g, `\n${marker("HEADING:1")}$1\n`)
    .replace(/\\subsection\*?\s*\{([^{}]*)\}/g, `\n${marker("HEADING:2")}$1\n`)
    .replace(/\\subsubsection\*?\s*\{([^{}]*)\}/g, `\n${marker("HEADING:3")}$1\n`)
    .replace(/\\(?:paragraph|subparagraph)\*?\s*\{([^{}]*)\}/g, `\n${marker("HEADING:4")}$1\n`)
    .replace(/\\begin\s*\{itemize\}/g, `\n${marker("LIST_START:bullet")}\n`)
    .replace(/\\end\s*\{itemize\}/g, `\n${marker("LIST_END")}\n`)
    .replace(/\\begin\s*\{enumerate\}/g, `\n${marker("LIST_START:number")}\n`)
    .replace(/\\end\s*\{enumerate\}/g, `\n${marker("LIST_END")}\n`)
    .replace(/\\item(?:\[[^\]]*\])?\s*/g, `\n${marker("ITEM")}`)
    .replace(/\\begin\s*\{(?:quote|quotation)\}/g, `\n${marker("QUOTE_START")}\n`)
    .replace(/\\end\s*\{(?:quote|quotation)\}/g, `\n${marker("QUOTE_END")}\n`)
    .replace(/\\begin\s*\{center\}/g, `\n${marker("ALIGN:center")}\n`)
    .replace(/\\end\s*\{center\}/g, `\n${marker("ALIGN:left")}\n`)
    .replace(/\\begin\s*\{flushright\}/g, `\n${marker("ALIGN:right")}\n`)
    .replace(/\\end\s*\{flushright\}/g, `\n${marker("ALIGN:left")}\n`)
    .replace(/\\begin\s*\{flushleft\}/g, `\n${marker("ALIGN:left")}\n`)
    .replace(/\\end\s*\{flushleft\}/g, `\n${marker("ALIGN:left")}\n`)
    .replace(/\\begin\s*\{(?:abstract|description)\}/g, "\n")
    .replace(/\\end\s*\{(?:abstract|description)\}/g, "\n")
    .replace(/\\par\b/g, "\n\n");
}

const cjkInlineBoundaryCharacter =
  /[\u2E80-\u2EFF\u3000-\u303F\u3040-\u30FF\u31F0-\u31FF\u3400-\u4DBF\u4E00-\u9FFF\uAC00-\uD7AF\uF900-\uFAFF\uFE30-\uFE4F\uFF00-\uFFEF]/;

function textEndsAtCjkBoundary(value: string) {
  const character = value.match(/(\S)\s*$/)?.[1] ?? "";
  return cjkInlineBoundaryCharacter.test(character);
}

function textStartsAtCjkBoundary(value: string) {
  const character = value.match(/^\s*(\S)/)?.[1] ?? "";
  return cjkInlineBoundaryCharacter.test(character);
}

function normalizeInlineFormulaBoundarySpacing(content: DocumentImportBlock[]) {
  for (let index = 0; index < content.length; index += 1) {
    if (content[index].kind !== "formula") continue;
    const previous = content[index - 1];
    const next = content[index + 1];
    if (
      previous?.kind === "text" &&
      /[ \t]$/.test(previous.text) &&
      textEndsAtCjkBoundary(previous.text)
    ) {
      previous.text = previous.text.replace(/[ \t]+$/g, "");
    }
    if (
      next?.kind === "text" &&
      /^[ \t]/.test(next.text) &&
      textStartsAtCjkBoundary(next.text)
    ) {
      next.text = next.text.replace(/^[ \t]+/g, "");
    }
  }
}

function emitParagraph(
  blocks: DocumentImportBlock[],
  raw: string,
  context: ParagraphContext,
  formulas: ExtractedFormula[],
  sourceKind: DocumentImportSourceKind,
) {
  if (!raw.trim()) return;
  const paragraphId = createUuid();
  const content: DocumentImportBlock[] = [];
  let cursor = 0;
  formulaTokenPattern.lastIndex = 0;
  for (let match = formulaTokenPattern.exec(raw); match; match = formulaTokenPattern.exec(raw)) {
    const text = cleanInlineMarkup(raw.slice(cursor, match.index), sourceKind);
    if (text) content.push({ id: createUuid(), kind: "text", text });
    const formula = formulaFromToken(match[0], formulas);
    if (formula) content.push({ ...formula });
    cursor = match.index + match[0].length;
  }
  const tail = cleanInlineMarkup(raw.slice(cursor), sourceKind);
  if (tail) content.push({ id: createUuid(), kind: "text", text: tail });
  if (!content.length) return;
  normalizeInlineFormulaBoundarySpacing(content);

  const firstText = content.find((block): block is DocumentTextBlock => block.kind === "text");
  const lastText = [...content]
    .reverse()
    .find((block): block is DocumentTextBlock => block.kind === "text");
  if (firstText) firstText.text = firstText.text.replace(/^\s+/, "");
  if (lastText) lastText.text = lastText.text.replace(/\s+$/, "");

  const visible = content.filter((block) => block.kind === "formula" || block.text.length > 0);
  visible.forEach((block, index) => {
    Object.assign(block, {
      paragraphId,
      paragraphStyle: context.style,
      paragraphAlignment: context.alignment,
      listKind: context.listKind,
      listLevel: context.listLevel,
      paragraphStart: index === 0,
      paragraphEnd: index === visible.length - 1,
    });
  });
  blocks.push(...visible);
}

function emitDisplayFormula(
  blocks: DocumentImportBlock[],
  token: string,
  formulas: ExtractedFormula[],
) {
  const formula = formulaFromToken(token, formulas);
  if (formula) blocks.push({ ...formula });
}

function emitLiteralFallback(
  blocks: DocumentImportBlock[],
  literal: ExtractedLiteral,
) {
  const lines = literal.text.replace(/\r\n?/g, "\n").split("\n");
  for (const line of lines) {
    if (!line.length) continue;
    const paragraphId = createUuid();
    blocks.push({
      id: createUuid(),
      kind: "text",
      text: line,
      paragraphId,
      paragraphStyle: "code",
      paragraphAlignment: "left",
      listKind: "none",
      listLevel: 0,
      paragraphStart: true,
      paragraphEnd: true,
    });
  }
}

function parseStructuredLines(
  source: string,
  formulas: ExtractedFormula[],
  literals: ExtractedLiteral[],
  sourceKind: DocumentImportSourceKind,
) {
  const blocks: DocumentImportBlock[] = [];
  const listStack: DocumentListKind[] = [];
  const theoremBodyStyleStack: DocumentParagraphStyle[] = [];
  const theoremCounters = new Map<string, number>();
  let alignment: DocumentParagraphAlignment = "left";
  let quoteDepth = 0;
  let current = "";
  let currentContext = defaultContext();
  let codeFence = false;

  const activeParagraphStyle = (): DocumentParagraphStyle =>
    codeFence
      ? "code"
      : quoteDepth > 0
        ? "quote"
        : theoremBodyStyleStack.at(-1) ?? "normal";

  const flush = () => {
    emitParagraph(blocks, current, currentContext, formulas, sourceKind);
    current = "";
    currentContext = {
      style: activeParagraphStyle(),
      alignment,
      listKind: "none",
      listLevel: 0,
    };
  };

  const beginParagraph = (context: ParagraphContext, text: string) => {
    flush();
    currentContext = context;
    current = text;
  };

  for (const originalLine of source.split("\n")) {
    const line = originalLine.replace(/[ \t]+$/g, "");
    const trimmed = line.trim();
    const leading = leadingMarker(trimmed);
    const control = leading?.value;

    if (control) {
      if (control.startsWith("THEOREM_START:")) {
        flush();
        const payload = decodeTheoremMarker(
          control.slice("THEOREM_START:".length),
        );
        if (payload) {
          let title = payload.label;
          if (payload.numbered) {
            const number = (theoremCounters.get(payload.counterName) ?? 0) + 1;
            theoremCounters.set(payload.counterName, number);
            title = `${title} ${number}`;
          }
          if (payload.note) title += `（${payload.note}）`;
          beginParagraph(
            {
              style: "heading4",
              alignment: "left",
              listKind: "none",
              listLevel: 0,
            },
            title,
          );
          flush();
          theoremBodyStyleStack.push(payload.bodyStyle);
        }
      } else if (control === "THEOREM_END") {
        flush();
        theoremBodyStyleStack.pop();
      } else if (control.startsWith("HEADING:")) {
        const title = leading?.rest ?? "";
        beginParagraph(
          {
            style: headingStyle(Number(control.split(":")[1])),
            alignment: "left",
            listKind: "none",
            listLevel: 0,
          },
          title,
        );
        flush();
      } else if (control.startsWith("LIST_START:")) {
        flush();
        listStack.push(control.endsWith("number") ? "number" : "bullet");
      } else if (control === "LIST_END") {
        flush();
        listStack.pop();
      } else if (control === "ITEM") {
        beginParagraph(
          {
            style: activeParagraphStyle(),
            alignment,
            listKind: listStack.at(-1) ?? "bullet",
            listLevel: Math.max(1, listStack.length),
          },
          leading?.rest ?? "",
        );
      } else if (control === "QUOTE_START") {
        flush();
        quoteDepth += 1;
      } else if (control === "QUOTE_END") {
        flush();
        quoteDepth = Math.max(0, quoteDepth - 1);
      } else if (control.startsWith("ALIGN:")) {
        flush();
        const value = control.split(":")[1];
        alignment = value === "center" || value === "right" ? value : "left";
      } else if (control === "CENTER") {
        beginParagraph(
          { style: "normal", alignment: "center", listKind: "none", listLevel: 0 },
          leading?.rest ?? "",
        );
        flush();
      }
      continue;
    }

    const literal = literalFromToken(trimmed, literals);
    if (literal) {
      flush();
      emitLiteralFallback(blocks, literal);
      continue;
    }

    const displayFormula = formulaFromToken(trimmed, formulas);
    if (displayFormula?.displayMode === "block") {
      flush();
      emitDisplayFormula(blocks, trimmed, formulas);
      continue;
    }

    if (sourceKind === "markdown") {
      if (/^```/.test(trimmed)) {
        flush();
        codeFence = !codeFence;
        continue;
      }
      const heading = trimmed.match(/^(#{1,6})\s+(.+)$/);
      if (heading) {
        beginParagraph(
          {
            style: headingStyle(heading[1].length),
            alignment: "left",
            listKind: "none",
            listLevel: 0,
          },
          heading[2],
        );
        flush();
        continue;
      }
      const bullet = line.match(/^(\s*)[-+*]\s+(.+)$/);
      if (bullet) {
        beginParagraph(
          {
            style: "normal",
            alignment: "left",
            listKind: "bullet",
            listLevel: Math.max(1, Math.floor(bullet[1].length / 2) + 1),
          },
          bullet[2],
        );
        continue;
      }
      const numbered = line.match(/^(\s*)\d+[.)]\s+(.+)$/);
      if (numbered) {
        beginParagraph(
          {
            style: "normal",
            alignment: "left",
            listKind: "number",
            listLevel: Math.max(1, Math.floor(numbered[1].length / 2) + 1),
          },
          numbered[2],
        );
        continue;
      }
      const quote = trimmed.match(/^>\s?(.*)$/);
      if (quote) {
        beginParagraph(
          { style: "quote", alignment: "left", listKind: "none", listLevel: 0 },
          quote[1],
        );
        continue;
      }
    }

    if (!trimmed) {
      flush();
      continue;
    }
    if (!current) {
      currentContext = {
        style: activeParagraphStyle(),
        alignment,
        listKind: "none",
        listLevel: 0,
      };
      current = trimmed;
    } else {
      current += ` ${trimmed}`;
    }
  }
  flush();
  return blocks;
}

function resolvedSourceKind(source: string, requested: DocumentImportSourceKind) {
  if (requested !== "auto") return requested;
  if (hasLatexCommentSignal(source)) return "latex";
  literalLatexCommandPattern.lastIndex = 0;
  if (literalLatexCommandPattern.test(source)) return "latex";
  const hasLiteralEnvironmentOutsideMath = latexLiteralFallbackRanges(source).some(
    (range) => /^\\begin\s*\{/.test(source.slice(range.start, range.end).trim()),
  );
  if (hasLiteralEnvironmentOutsideMath) return "latex";
  // LaTeX fragments pasted from notes often omit both \documentclass and the
  // document environment. Structural commands and environments must still
  // select the LaTeX parser. Do not end this pattern with \b: environment
  // matches end in `}`, which is not a word character, so that boundary made
  // valid `\begin{itemize}` fragments fail auto-detection.
  return /\\(?:documentclass|usepackage|title\s*\{|author\s*\{|date\s*\{|maketitle\b|(?:part|chapter|section|subsection|subsubsection|paragraph|subparagraph)\*?\s*\{|item(?:\[[^\]]*\])?\s|(?:textbf|textit|emph|underline|texttt|textrm|textsf|textnormal|mbox|caption|footnote|href)\s*\{|(?:newline|linebreak|par)\b|(?:begin|end)\s*\{(?:document|itemize|enumerate|quote|quotation|center|flushleft|flushright|abstract|description|verbatim\*?|lstlisting|minted|comment)\})/.test(
    source,
  )
    ? "latex"
    : "markdown";
}

export function parseLatexMarkdownDocument(
  source: string,
  requestedKind: DocumentImportSourceKind = "auto",
  defaultFontSizePt = 12,
): DocumentImportBlock[] {
  let normalized = source.replace(/^\uFEFF/, "").replace(/\r\n?/g, "\n");
  const sourceKind = resolvedSourceKind(normalized, requestedKind);
  let literals: ExtractedLiteral[] = [];
  let theoremDefinitions = new Map<string, TheoremEnvironmentDefinition>();
  let theoremEnvironmentNames = new Set<string>();

  if (sourceKind === "markdown") {
    // Protect math before Markdown structural normalization. Setext headings,
    // tables and list syntax may use characters such as `=`, `-` and `|` that
    // are also perfectly valid inside multi-line LaTeX. Extracting formulas to
    // private-use tokens first prevents Markdown from rewriting a display
    // equation before the formula parser gets a chance to see it.
    const extracted = extractFormulas(
      normalized,
      defaultFontSizePt,
      sourceKind,
      theoremEnvironmentNames,
    );
    const structured = normalizeMarkdownStructure(extracted.text);
    return parseStructuredLines(
      structured,
      extracted.formulas,
      literals,
      sourceKind,
    );
  }

  normalized = normalizeLatexExtensionsFlat(normalized);
  theoremDefinitions = theoremDefinitionsFromSource(normalized);
  theoremEnvironmentNames = new Set(theoremDefinitions.keys());
  const literalized = extractLatexLiteralFallbacks(
    normalized,
    theoremEnvironmentNames,
  );
  normalized = stripLatexComments(
    literalized.text,
    theoremEnvironmentNames,
  );
  literals = literalized.literals;

  const extracted = extractFormulas(
    normalized,
    defaultFontSizePt,
    sourceKind,
    theoremEnvironmentNames,
  );
  const structured = normalizeLatexStructure(
    extracted.text,
    theoremDefinitions,
  );
  return parseStructuredLines(
    structured,
    extracted.formulas,
    literals,
    sourceKind,
  );
}

export function mergeDocumentImportBlocks(
  previous: DocumentImportBlock[],
  next: DocumentImportBlock[],
): DocumentImportBlock[] {
  const previousFormulas = previous.filter(
    (block): block is DocumentFormulaBlock => block.kind === "formula",
  );
  let formulaIndex = 0;
  return next.map((block) => {
    if (block.kind === "text") return block;
    const prior = previousFormulas[formulaIndex];
    formulaIndex += 1;
    if (!prior) return block;
    return {
      ...block,
      id: prior.id,
      fontSizePt: prior.fontSizePt,
      displayMode: prior.displayMode,
      numbered: prior.displayMode === "block" && prior.numbered,
    };
  });
}
