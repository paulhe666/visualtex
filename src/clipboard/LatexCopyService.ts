import { validateLatex } from "mathlive/ssr";
import { normalizeMathLiveCanonicalUprightCommands } from "../editor/normalizeChineseLatex.ts";
import type { LatexCodeFormat } from "../types/formula";

export type LatexCodeFormatGroup = "single" | "multi";

export interface LatexCodeFormatDefinition {
  id: LatexCodeFormat;
  group: LatexCodeFormatGroup;
  titleZh: string;
  titleEn: string;
  hint: string;
  descriptionZh: string;
  descriptionEn: string;
  numbered?: boolean;
}

export const DEFAULT_LATEX_CODE_FORMAT: LatexCodeFormat = "display-dollar";

export const latexCodeFormats: readonly LatexCodeFormatDefinition[] = [
  {
    id: "raw",
    group: "single",
    titleZh: "纯 LaTeX 源码",
    titleEn: "Raw LaTeX",
    hint: "\\frac{x}{y}",
    descriptionZh: "每个公式占一行，不添加环境",
    descriptionEn: "One formula per line without wrappers",
  },
  {
    id: "inline-dollar",
    group: "single",
    titleZh: "行内公式 · 美元符号",
    titleEn: "Inline math · dollar signs",
    hint: "$ ... $",
    descriptionZh: "每个公式分别使用 $...$",
    descriptionEn: "Wrap every formula with $...$",
  },
  {
    id: "inline-text-double-dollar",
    group: "single",
    titleZh: "行内公式 · 文字在公式外",
    titleEn: "Inline formula · text outside math",
    hint: "文字$$x^2$$文字",
    descriptionZh:
      "顶层文字直接放在公式环境外，公式片段使用 $$...$$；上下标等公式结构中的中文仍保留 \\text{}",
    descriptionEn:
      "Keep top-level text outside math and wrap formula fragments with $$...$$; Chinese inside scripts and other math structures remains in \\text{}",
  },
  {
    id: "inline-paren",
    group: "single",
    titleZh: "行内公式 · 圆括号",
    titleEn: "Inline math · parentheses",
    hint: "\\( ... \\)",
    descriptionZh: "每个公式分别使用 \\( ... \\)",
    descriptionEn: "Wrap every formula with \\( ... \\)",
  },
  {
    id: "display-dollar",
    group: "single",
    titleZh: "行间公式 · 双美元符号",
    titleEn: "Display math · double dollars",
    hint: "$$ ... $$",
    descriptionZh: "每个公式分别使用 $$...$$",
    descriptionEn: "Wrap every formula with $$...$$",
  },
  {
    id: "display-bracket",
    group: "single",
    titleZh: "行间公式 · 方括号",
    titleEn: "Display math · brackets",
    hint: "\\[ ... \\]",
    descriptionZh: "每个公式分别使用 \\[ ... \\]",
    descriptionEn: "Wrap every formula with \\[ ... \\]",
  },
  {
    id: "equation",
    group: "single",
    titleZh: "equation · 自动编号",
    titleEn: "equation · numbered",
    hint: "\\begin{equation}",
    descriptionZh: "每个公式一个 equation 环境并自动编号",
    descriptionEn: "One numbered equation environment per formula",
    numbered: true,
  },
  {
    id: "equation-star",
    group: "single",
    titleZh: "equation* · 不编号",
    titleEn: "equation* · unnumbered",
    hint: "\\begin{equation*}",
    descriptionZh: "每个公式一个 equation* 环境，不显示编号",
    descriptionEn: "One unnumbered equation* environment per formula",
    numbered: false,
  },
  {
    id: "align",
    group: "multi",
    titleZh: "align · 多行自动编号",
    titleEn: "align · numbered rows",
    hint: "\\begin{align}",
    descriptionZh: "所有公式合并为一个 align 环境，每行自动编号",
    descriptionEn: "Combine all formulas into one numbered align environment",
    numbered: true,
  },
  {
    id: "align-star",
    group: "multi",
    titleZh: "align* · 多行不编号",
    titleEn: "align* · unnumbered rows",
    hint: "\\begin{align*}",
    descriptionZh: "所有公式合并为一个 align* 环境，不显示编号",
    descriptionEn: "Combine all formulas into one unnumbered align* environment",
    numbered: false,
  },
  {
    id: "aligned",
    group: "multi",
    titleZh: "aligned · 方括号行间公式",
    titleEn: "aligned · bracket display",
    hint: "\\[ \\begin{aligned}",
    descriptionZh: "所有公式合并到 \\[...\\] 内的 aligned 环境",
    descriptionEn: "Combine all formulas in an aligned environment inside \\[...\\]",
    numbered: false,
  },
  {
    id: "gather",
    group: "multi",
    titleZh: "gather · 多行自动编号",
    titleEn: "gather · numbered rows",
    hint: "\\begin{gather}",
    descriptionZh: "所有公式居中排列，每行自动编号",
    descriptionEn: "Center all formulas and number every row",
    numbered: true,
  },
  {
    id: "gather-star",
    group: "multi",
    titleZh: "gather* · 多行不编号",
    titleEn: "gather* · unnumbered rows",
    hint: "\\begin{gather*}",
    descriptionZh: "所有公式居中排列，不显示编号",
    descriptionEn: "Center all formulas without row numbers",
    numbered: false,
  },
  {
    id: "multline",
    group: "multi",
    titleZh: "multline · 长公式自动编号",
    titleEn: "multline · numbered",
    hint: "\\begin{multline}",
    descriptionZh: "把多行内容视为一个长公式并生成一个编号",
    descriptionEn: "Treat the rows as one long equation with one number",
    numbered: true,
  },
  {
    id: "multline-star",
    group: "multi",
    titleZh: "multline* · 长公式不编号",
    titleEn: "multline* · unnumbered",
    hint: "\\begin{multline*}",
    descriptionZh: "把多行内容视为一个长公式，不显示编号",
    descriptionEn: "Treat the rows as one long equation without a number",
    numbered: false,
  },
  {
    id: "equation-split",
    group: "multi",
    titleZh: "equation + split · 单一编号",
    titleEn: "equation + split · one number",
    hint: "\\begin{equation} \\begin{split}",
    descriptionZh: "所有公式放入 split，并由外层 equation 生成一个编号",
    descriptionEn: "Put all formulas in split with one outer equation number",
    numbered: true,
  },
  {
    id: "equation-star-split",
    group: "multi",
    titleZh: "equation* + split · 不编号",
    titleEn: "equation* + split · unnumbered",
    hint: "\\begin{equation*} \\begin{split}",
    descriptionZh: "所有公式放入 split，外层 equation* 不显示编号",
    descriptionEn: "Put all formulas in split inside an unnumbered equation*",
    numbered: false,
  },
] as const;

export function isLatexCodeFormat(value: unknown): value is LatexCodeFormat {
  return latexCodeFormats.some((format) => format.id === value);
}

export function getLatexCodeFormatDefinition(
  format: LatexCodeFormat,
): LatexCodeFormatDefinition {
  return (
    latexCodeFormats.find((definition) => definition.id === format) ??
    latexCodeFormats.find(
      (definition) => definition.id === DEFAULT_LATEX_CODE_FORMAT,
    )!
  );
}

export function splitLatexLines(latex: string): string[] {
  const lines = latex.replace(/\r\n?/g, "\n").split("\n");
  return lines.length ? lines : [""];
}

function filledFormulaLines(latex: string): string[] {
  const lines = splitLatexLines(
    normalizeMathLiveCanonicalUprightCommands(latex),
  )
    .map((line) => line.trim())
    .filter(Boolean);
  return lines.length ? lines : [""];
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function wrapEnvironment(name: string, body: string): string {
  return `\\begin{${name}}\n${body}\n\\end{${name}}`;
}

function isEscaped(source: string, index: number): boolean {
  let slashCount = 0;
  for (let cursor = index - 1; cursor >= 0 && source[cursor] === "\\"; cursor -= 1) {
    slashCount += 1;
  }
  return slashCount % 2 === 1;
}

interface EnvironmentToken {
  kind: "begin" | "end";
  name: string;
  end: number;
}

function readEnvironmentToken(source: string, index: number): EnvironmentToken | null {
  if (source[index] !== "\\") return null;
  const match = source.slice(index).match(/^\\(begin|end)\{([A-Za-z]+\*?)\}/);
  if (!match) return null;
  return {
    kind: match[1] as EnvironmentToken["kind"],
    name: match[2],
    end: index + match[0].length,
  };
}

function updateEnvironmentStack(stack: string[], token: EnvironmentToken) {
  if (token.kind === "begin") {
    stack.push(token.name);
    return;
  }
  const matchingIndex = stack.lastIndexOf(token.name);
  if (matchingIndex >= 0) stack.splice(matchingIndex, 1);
}

function hasTopLevelAlignmentMarker(latex: string): boolean {
  let braceDepth = 0;
  const environments: string[] = [];

  for (let index = 0; index < latex.length; index += 1) {
    const token = readEnvironmentToken(latex, index);
    if (token) {
      updateEnvironmentStack(environments, token);
      index = token.end - 1;
      continue;
    }

    const character = latex[index];
    if (character === "{" && !isEscaped(latex, index)) braceDepth += 1;
    else if (character === "}" && !isEscaped(latex, index)) {
      braceDepth = Math.max(0, braceDepth - 1);
    } else if (
      character === "&" &&
      !isEscaped(latex, index) &&
      braceDepth === 0 &&
      environments.length === 0
    ) {
      return true;
    }
  }

  return false;
}

const relationCommands = [
  "\\Longleftrightarrow",
  "\\Longrightarrow",
  "\\Leftrightarrow",
  "\\Rightarrow",
  "\\leftrightarrow",
  "\\rightarrow",
  "\\leftarrow",
  "\\subseteq",
  "\\supseteq",
  "\\notin",
  "\\approx",
  "\\equiv",
  "\\simeq",
  "\\propto",
  "\\mapsto",
  "\\subset",
  "\\supset",
  "\\cong",
  "\\neq",
  "\\leq",
  "\\geq",
  "\\sim",
  "\\to",
  "\\ne",
  "\\le",
  "\\ge",
  "\\in",
] as const;

function findTopLevelRelationIndex(latex: string): number {
  let braceDepth = 0;
  const environments: string[] = [];

  for (let index = 0; index < latex.length; index += 1) {
    const token = readEnvironmentToken(latex, index);
    if (token) {
      updateEnvironmentStack(environments, token);
      index = token.end - 1;
      continue;
    }

    const character = latex[index];
    if (character === "{" && !isEscaped(latex, index)) {
      braceDepth += 1;
      continue;
    }
    if (character === "}" && !isEscaped(latex, index)) {
      braceDepth = Math.max(0, braceDepth - 1);
      continue;
    }
    if (braceDepth !== 0 || environments.length !== 0) continue;

    if (character === "=" || character === "<" || character === ">") {
      return index;
    }

    if (character !== "\\") continue;
    for (const command of relationCommands) {
      if (!latex.startsWith(command, index)) continue;
      const nextCharacter = latex[index + command.length];
      if (nextCharacter && /[A-Za-z]/.test(nextCharacter)) continue;
      return index;
    }
  }

  return -1;
}

function addAlignmentMarker(latex: string): string {
  if (!latex || hasTopLevelAlignmentMarker(latex)) return latex;
  const relationIndex = findTopLevelRelationIndex(latex);
  if (relationIndex < 0) return latex;
  return `${latex.slice(0, relationIndex)}&${latex.slice(relationIndex)}`;
}

function stripTopLevelAlignmentMarkers(latex: string): string {
  let result = "";
  let braceDepth = 0;
  const environments: string[] = [];

  for (let index = 0; index < latex.length; index += 1) {
    const token = readEnvironmentToken(latex, index);
    if (token) {
      result += latex.slice(index, token.end);
      updateEnvironmentStack(environments, token);
      index = token.end - 1;
      continue;
    }

    const character = latex[index];
    if (character === "{" && !isEscaped(latex, index)) braceDepth += 1;
    else if (character === "}" && !isEscaped(latex, index)) {
      braceDepth = Math.max(0, braceDepth - 1);
    }

    if (
      character === "&" &&
      !isEscaped(latex, index) &&
      braceDepth === 0 &&
      environments.length === 0
    ) {
      continue;
    }
    result += character;
  }

  return result.trim();
}

function formatRows(lines: string[], alignRelations: boolean): string {
  return lines
    .map((line, index) => {
      const content = alignRelations ? addAlignmentMarker(line) : line;
      return index < lines.length - 1 ? `${content} \\\\` : content;
    })
    .join("\n");
}

type InlineTextSegment = {
  kind: "math" | "text";
  value: string;
};

function readBalancedGroupEnd(
  source: string,
  openingIndex: number,
  opening = "{",
  closing = "}",
): number | null {
  if (source[openingIndex] !== opening) return null;
  let depth = 0;
  for (let index = openingIndex; index < source.length; index += 1) {
    if (source[index] === "%" && !isEscaped(source, index)) {
      const lineEnd = source.indexOf("\n", index);
      if (lineEnd < 0) return null;
      index = lineEnd;
      continue;
    }
    if (source[index] === opening && !isEscaped(source, index)) depth += 1;
    else if (source[index] === closing && !isEscaped(source, index)) {
      depth -= 1;
      if (depth === 0) return index + 1;
      if (depth < 0) return null;
    }
  }
  return null;
}

function appendInlineTextSegment(
  segments: InlineTextSegment[],
  kind: InlineTextSegment["kind"],
  value: string,
) {
  if (!value) return;
  const previous = segments.at(-1);
  if (previous?.kind === kind) previous.value += value;
  else segments.push({ kind, value });
}

function splitTopLevelTextSegments(latex: string): InlineTextSegment[] {
  const segments: InlineTextSegment[] = [];
  const environments: string[] = [];
  let math = "";
  let braceDepth = 0;

  const flushMath = () => {
    appendInlineTextSegment(segments, "math", math);
    math = "";
  };

  for (let index = 0; index < latex.length; index += 1) {
    const token = readEnvironmentToken(latex, index);
    if (token) {
      math += latex.slice(index, token.end);
      updateEnvironmentStack(environments, token);
      index = token.end - 1;
      continue;
    }

    if (
      braceDepth === 0 &&
      environments.length === 0 &&
      latex.startsWith("\\text{", index)
    ) {
      const openingBrace = index + "\\text".length;
      const end = readBalancedGroupEnd(latex, openingBrace);
      if (end !== null) {
        flushMath();
        appendInlineTextSegment(
          segments,
          "text",
          latex.slice(openingBrace + 1, end - 1),
        );
        index = end - 1;
        continue;
      }
    }

    const character = latex[index];
    if (character === "{" && !isEscaped(latex, index)) braceDepth += 1;
    else if (character === "}" && !isEscaped(latex, index)) {
      braceDepth = Math.max(0, braceDepth - 1);
    }
    math += character;
  }

  flushMath();
  return segments;
}

function formatInlineTextDoubleDollar(latex: string): string {
  const segments = splitTopLevelTextSegments(latex);
  if (!segments.length) return "";
  return segments
    .map((segment) => {
      if (segment.kind === "text") return segment.value;
      const math = segment.value.trim();
      return math ? `$$${math}$$` : "";
    })
    .join("");
}

function escapeOutsideTextForMath(value: string): string {
  return value.replace(/(?<!\\)([{}])/g, "\\$1");
}

function parseInlineTextDoubleDollarLine(line: string): string | null {
  let result = "";
  let cursor = 0;
  while (cursor < line.length) {
    const opening = line.indexOf("$$", cursor);
    if (opening < 0) {
      const text = line.slice(cursor);
      if (text) result += `\\text{${escapeOutsideTextForMath(text)}}`;
      break;
    }
    const text = line.slice(cursor, opening);
    if (text) result += `\\text{${escapeOutsideTextForMath(text)}}`;
    const closing = line.indexOf("$$", opening + 2);
    if (closing < 0) return null;
    const math = line.slice(opening + 2, closing).trim();
    if (!math) return null;
    result += math;
    cursor = closing + 2;
  }
  return result || "";
}

function parseInlineTextDoubleDollarLines(source: string): string[] {
  const values: string[] = [];
  for (const line of source.split("\n")) {
    if (!line.trim()) continue;
    const value = parseInlineTextDoubleDollarLine(line);
    if (value === null) return [];
    values.push(value);
  }
  return values;
}

export function formatLatex(latex: string, format: LatexCodeFormat): string {
  const lines = filledFormulaLines(latex);

  switch (format) {
    case "raw":
      return lines.join("\n");
    case "inline-dollar":
      return lines.map((line) => `$${line}$`).join("\n");
    case "inline-text-double-dollar":
      return lines.map(formatInlineTextDoubleDollar).join("\n");
    case "inline-paren":
      return lines.map((line) => `\\(${line}\\)`).join("\n");
    case "display-dollar":
      return lines.map((line) => `$$\n${line}\n$$`).join("\n\n");
    case "display-bracket":
      return lines.map((line) => `\\[\n${line}\n\\]`).join("\n\n");
    case "equation":
      return lines.map((line) => wrapEnvironment("equation", line)).join("\n\n");
    case "equation-star":
      return lines.map((line) => wrapEnvironment("equation*", line)).join("\n\n");
    case "align":
      return wrapEnvironment("align", formatRows(lines, true));
    case "align-star":
      return wrapEnvironment("align*", formatRows(lines, true));
    case "aligned":
      return `\\[\n${wrapEnvironment("aligned", formatRows(lines, true))}\n\\]`;
    case "gather":
      return wrapEnvironment("gather", formatRows(lines, false));
    case "gather-star":
      return wrapEnvironment("gather*", formatRows(lines, false));
    case "multline":
      return wrapEnvironment("multline", formatRows(lines, false));
    case "multline-star":
      return wrapEnvironment("multline*", formatRows(lines, false));
    case "equation-split":
      return wrapEnvironment(
        "equation",
        wrapEnvironment("split", formatRows(lines, true)),
      );
    case "equation-star-split":
      return wrapEnvironment(
        "equation*",
        wrapEnvironment("split", formatRows(lines, true)),
      );
    default:
      return formatLatex(latex, DEFAULT_LATEX_CODE_FORMAT);
  }
}

function extractEnvironmentBodies(source: string, name: string): string[] {
  const escapedName = escapeRegExp(name);
  const pattern = new RegExp(
    `\\\\begin\\{${escapedName}\\}([\\s\\S]*?)\\\\end\\{${escapedName}\\}`,
    "g",
  );
  return [...source.matchAll(pattern)].map((match) => match[1].trim());
}

function splitTopLevelRows(body: string): string[] {
  const rows: string[] = [];
  let current = "";
  let braceDepth = 0;
  const environments: string[] = [];

  for (let index = 0; index < body.length; index += 1) {
    const token = readEnvironmentToken(body, index);
    if (token) {
      current += body.slice(index, token.end);
      updateEnvironmentStack(environments, token);
      index = token.end - 1;
      continue;
    }

    const character = body[index];
    if (character === "%" && !isEscaped(body, index)) {
      const lineEnd = body.indexOf("\n", index);
      if (lineEnd < 0) {
        current += body.slice(index);
        break;
      }
      current += body.slice(index, lineEnd + 1);
      index = lineEnd;
      continue;
    }

    if (character === "{" && !isEscaped(body, index)) braceDepth += 1;
    else if (character === "}" && !isEscaped(body, index)) {
      braceDepth = Math.max(0, braceDepth - 1);
    }

    if (
      character === "\\" &&
      body[index + 1] === "\\" &&
      braceDepth === 0 &&
      environments.length === 0
    ) {
      rows.push(current.trim());
      current = "";
      index += 1;

      let cursor = index + 1;
      while (/\s/.test(body[cursor] ?? "")) cursor += 1;
      if (body[cursor] === "[") {
        const closingBracket = body.indexOf("]", cursor + 1);
        if (closingBracket >= 0) cursor = closingBracket + 1;
      }
      while (/\s/.test(body[cursor] ?? "")) cursor += 1;
      index = cursor - 1;
      continue;
    }

    current += character;
  }

  if (current.trim() || rows.length === 0) rows.push(current.trim());
  return rows.filter((row) => row.length > 0);
}

function parseWrappedBlocks(source: string, pattern: RegExp): string[] {
  return [...source.matchAll(pattern)]
    .map((match) => match[1].trim())
    .filter(Boolean);
}

function parseInlineDollarLines(source: string): string[] {
  return source
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line.startsWith("$") && !line.startsWith("$$"))
    .map((line) =>
      line.endsWith("$") && !line.endsWith("$$")
        ? line.slice(1, -1).trim()
        : line.slice(1).trim(),
    )
    .filter(Boolean);
}

function parseMultilineEnvironment(source: string, name: string): string[] {
  const body = extractEnvironmentBodies(source, name)[0];
  if (body === undefined) return [];
  return splitTopLevelRows(body).map(stripTopLevelAlignmentMarkers);
}

function parseByFormat(source: string, format: LatexCodeFormat): string[] {
  switch (format) {
    case "raw":
      return source
        .split("\n")
        .map((line) => line.trim())
        .filter(Boolean);
    case "inline-dollar":
      return parseInlineDollarLines(source);
    case "inline-text-double-dollar":
      return parseInlineTextDoubleDollarLines(source);
    case "inline-paren":
      return parseWrappedBlocks(source, /\\\(([\s\S]*?)\\\)/g);
    case "display-dollar":
      return parseWrappedBlocks(source, /\$\$([\s\S]*?)\$\$/g);
    case "display-bracket":
      return parseWrappedBlocks(source, /\\\[([\s\S]*?)\\\]/g);
    case "equation":
      return extractEnvironmentBodies(source, "equation");
    case "equation-star":
      return extractEnvironmentBodies(source, "equation*");
    case "align":
      return parseMultilineEnvironment(source, "align");
    case "align-star":
      return parseMultilineEnvironment(source, "align*");
    case "aligned":
      return parseMultilineEnvironment(source, "aligned");
    case "gather":
      return parseMultilineEnvironment(source, "gather");
    case "gather-star":
      return parseMultilineEnvironment(source, "gather*");
    case "multline":
      return parseMultilineEnvironment(source, "multline");
    case "multline-star":
      return parseMultilineEnvironment(source, "multline*");
    case "equation-split":
    case "equation-star-split":
      return parseMultilineEnvironment(source, "split");
    default:
      return [];
  }
}

function parseCoveredBlocksStrict(
  source: string,
  pattern: RegExp,
): string[] | null {
  pattern.lastIndex = 0;
  const values: string[] = [];
  let cursor = 0;
  for (const match of source.matchAll(pattern)) {
    const index = match.index ?? 0;
    if (source.slice(cursor, index).trim()) return null;
    const value = match[1]?.trim() ?? "";
    if (!value) return null;
    values.push(value);
    cursor = index + match[0].length;
  }
  if (source.slice(cursor).trim()) return null;
  return values.length ? values : null;
}

function parseEnvironmentBlocksStrict(
  source: string,
  name: string,
): string[] | null {
  const escapedName = escapeRegExp(name);
  return parseCoveredBlocksStrict(
    source,
    new RegExp(
      `\\\\begin\\{${escapedName}\\}([\\s\\S]*?)\\\\end\\{${escapedName}\\}`,
      "g",
    ),
  );
}

function parseSingleMultilineEnvironmentStrict(
  source: string,
  name: string,
): string[] | null {
  const bodies = parseEnvironmentBlocksStrict(source, name);
  if (!bodies || bodies.length !== 1) return null;
  const rows = splitTopLevelRows(bodies[0]).map(stripTopLevelAlignmentMarkers);
  return rows.length ? rows : null;
}

function parseInlineDollarLinesStrict(source: string): string[] | null {
  const values: string[] = [];
  for (const rawLine of source.split("\n")) {
    const line = rawLine.trim();
    if (!line) continue;
    if (
      !line.startsWith("$") ||
      line.startsWith("$$") ||
      !line.endsWith("$") ||
      line.endsWith("$$")
    ) {
      return null;
    }
    const value = line.slice(1, -1).trim();
    if (!value) return null;
    values.push(value);
  }
  return values.length ? values : null;
}

function parseInlineParenLinesStrict(source: string): string[] | null {
  const values: string[] = [];
  for (const rawLine of source.split("\n")) {
    const line = rawLine.trim();
    if (!line) continue;
    if (!line.startsWith("\\(") || !line.endsWith("\\)")) return null;
    const value = line.slice(2, -2).trim();
    if (!value) return null;
    values.push(value);
  }
  return values.length ? values : null;
}

function parseInlineTextDoubleDollarLinesStrict(
  source: string,
): string[] | null {
  const values: string[] = [];
  for (const line of source.split("\n")) {
    if (!line.trim()) continue;
    const value = parseInlineTextDoubleDollarLine(line);
    if (value === null || !value.trim()) return null;
    values.push(value);
  }
  return values.length ? values : null;
}

function parseByFormatStrict(
  source: string,
  format: LatexCodeFormat,
): string[] | null {
  switch (format) {
    case "raw": {
      const values = source
        .split("\n")
        .map((line) => line.trim())
        .filter(Boolean);
      return values.length ? values : null;
    }
    case "inline-dollar":
      return parseInlineDollarLinesStrict(source);
    case "inline-text-double-dollar":
      return parseInlineTextDoubleDollarLinesStrict(source);
    case "inline-paren":
      return parseInlineParenLinesStrict(source);
    case "display-dollar":
      return parseCoveredBlocksStrict(source, /\$\$([\s\S]*?)\$\$/g);
    case "display-bracket":
      return parseCoveredBlocksStrict(source, /\\\[([\s\S]*?)\\\]/g);
    case "equation":
      return parseEnvironmentBlocksStrict(source, "equation");
    case "equation-star":
      return parseEnvironmentBlocksStrict(source, "equation*");
    case "align":
      return parseSingleMultilineEnvironmentStrict(source, "align");
    case "align-star":
      return parseSingleMultilineEnvironmentStrict(source, "align*");
    case "gather":
      return parseSingleMultilineEnvironmentStrict(source, "gather");
    case "gather-star":
      return parseSingleMultilineEnvironmentStrict(source, "gather*");
    case "multline":
      return parseSingleMultilineEnvironmentStrict(source, "multline");
    case "multline-star":
      return parseSingleMultilineEnvironmentStrict(source, "multline*");
    case "aligned": {
      const outer = parseCoveredBlocksStrict(source, /\\\[([\s\S]*?)\\\]/g);
      if (!outer || outer.length !== 1) return null;
      return parseSingleMultilineEnvironmentStrict(outer[0], "aligned");
    }
    case "equation-split":
    case "equation-star-split": {
      const outerName = format === "equation-split" ? "equation" : "equation*";
      const outer = parseEnvironmentBlocksStrict(source, outerName);
      if (!outer || outer.length !== 1) return null;
      return parseSingleMultilineEnvironmentStrict(outer[0], "split");
    }
    default:
      return null;
  }
}

const requiredCommandArgumentCount = new Map<string, number>([
  ["frac", 2],
  ["dfrac", 2],
  ["tfrac", 2],
  ["cfrac", 2],
  ["binom", 2],
  ["overset", 2],
  ["underset", 2],
  ["stackrel", 2],
  ["stackbin", 2],
  ["sqrt", 1],
  ["text", 1],
  ["textrm", 1],
  ["mathrm", 1],
  ["mathbf", 1],
  ["mathit", 1],
  ["mathsf", 1],
  ["mathtt", 1],
  ["mathbb", 1],
  ["mathcal", 1],
  ["mathfrak", 1],
  ["operatorname", 1],
  ["overline", 1],
  ["underline", 1],
  ["hat", 1],
  ["widehat", 1],
  ["tilde", 1],
  ["widetilde", 1],
  ["vec", 1],
  ["dot", 1],
  ["ddot", 1],
  ["dddot", 1],
  ["ddddot", 1],
  ["check", 1],
  ["breve", 1],
  ["acute", 1],
  ["grave", 1],
  ["mathring", 1],
  ["overbrace", 1],
  ["underbrace", 1],
  ["overarc", 1],
  ["underarc", 1],
  ["overparen", 1],
  ["underparen", 1],
  ["overgroup", 1],
  ["undergroup", 1],
  ["overrightarrow", 1],
  ["overleftarrow", 1],
  ["overleftrightarrow", 1],
  ["underleftarrow", 1],
  ["underrightarrow", 1],
  ["underleftrightarrow", 1],
]);

function skipLatexWhitespace(source: string, start: number): number {
  let index = start;
  while (/\s/.test(source[index] ?? "")) index += 1;
  return index;
}

function readLatexCommandEnd(source: string, start: number): number {
  if (source[start] !== "\\") return start;
  let index = start + 1;
  if (/[A-Za-z]/.test(source[index] ?? "")) {
    while (/[A-Za-z]/.test(source[index] ?? "")) index += 1;
    return index;
  }
  return Math.min(source.length, index + 1);
}

function readLatexArgumentEnd(source: string, start: number): number | null {
  const index = skipLatexWhitespace(source, start);
  if (index >= source.length || source[index] === "}") return null;
  if (source[index] === "{") return readBalancedGroupEnd(source, index);
  if (source[index] === "\\") return readLatexCommandEnd(source, index);
  return index + 1;
}

function hasBalancedLatexGroups(source: string): boolean {
  const stack: string[] = [];
  for (let index = 0; index < source.length; index += 1) {
    const character = source[index];
    if (character === "%" && !isEscaped(source, index)) {
      const lineEnd = source.indexOf("\n", index);
      if (lineEnd < 0) break;
      index = lineEnd;
      continue;
    }
    if (isEscaped(source, index)) continue;
    if (character === "{" || character === "[") stack.push(character);
    else if (character === "}" || character === "]") {
      const expected = character === "}" ? "{" : "[";
      if (stack.pop() !== expected) return false;
    }
  }
  return stack.length === 0;
}

function hasCompleteRequiredCommandArguments(source: string): boolean {
  for (let index = 0; index < source.length; index += 1) {
    if (source[index] === "%" && !isEscaped(source, index)) {
      const lineEnd = source.indexOf("\n", index);
      if (lineEnd < 0) break;
      index = lineEnd;
      continue;
    }
    if (source[index] !== "\\" || isEscaped(source, index)) continue;
    const commandEnd = readLatexCommandEnd(source, index);
    const command = source.slice(index + 1, commandEnd);
    const requiredArguments = requiredCommandArgumentCount.get(command);
    if (!requiredArguments) {
      index = commandEnd - 1;
      continue;
    }

    let cursor = skipLatexWhitespace(source, commandEnd);
    if (source[cursor] === "*") cursor = skipLatexWhitespace(source, cursor + 1);
    if (command === "sqrt" && source[cursor] === "[") {
      const optionalEnd = readBalancedGroupEnd(source, cursor, "[", "]");
      if (optionalEnd === null) return false;
      cursor = optionalEnd;
    }
    for (let argument = 0; argument < requiredArguments; argument += 1) {
      const argumentEnd = readLatexArgumentEnd(source, cursor);
      if (argumentEnd === null) return false;
      cursor = argumentEnd;
    }
    index = commandEnd - 1;
  }
  return true;
}

function validateFormulaDraft(latex: string): string | null {
  if (!latex.trim()) return "empty-formula";
  if (!hasBalancedLatexGroups(latex)) return "unbalanced-group";
  if (/(^|[^\\])[_^]\s*$/.test(latex)) return "incomplete-script";
  if (!hasCompleteRequiredCommandArguments(latex)) {
    return "incomplete-command-arguments";
  }
  const errors = validateLatex(latex);
  if (errors.length) return errors[0]?.code ?? "invalid-latex";
  return null;
}

export interface LatexSourceDraftResult {
  valid: boolean;
  values: string[];
  error?: string;
}

export function parseLatexSourceDraft(
  source: string,
  format: LatexCodeFormat = DEFAULT_LATEX_CODE_FORMAT,
): LatexSourceDraftResult {
  const normalized = source.replace(/\r\n?/g, "\n");
  if (!normalized.trim()) return { valid: true, values: [""] };
  const values = parseByFormatStrict(normalized.trim(), format);
  if (!values?.length) {
    return { valid: false, values: [], error: "incomplete-format-wrapper" };
  }
  for (const value of values) {
    const error = validateFormulaDraft(value);
    if (error) return { valid: false, values, error };
  }
  return { valid: true, values };
}

export function parseLatexSource(
  source: string,
  preferredFormat: LatexCodeFormat = DEFAULT_LATEX_CODE_FORMAT,
): string[] {
  const normalized = source.replace(/\r\n?/g, "\n").trim();
  if (!normalized) return [""];

  const preferred = parseByFormat(normalized, preferredFormat);
  if (preferred.length) return preferred;

  const fallbackOrder: LatexCodeFormat[] = [
    "equation-split",
    "equation-star-split",
    "align",
    "align-star",
    "aligned",
    "gather",
    "gather-star",
    "multline",
    "multline-star",
    "equation",
    "equation-star",
    "display-dollar",
    "display-bracket",
    "inline-paren",
    "inline-text-double-dollar",
    "inline-dollar",
    "raw",
  ];

  for (const format of fallbackOrder) {
    if (format === preferredFormat) continue;
    const parsed = parseByFormat(normalized, format);
    if (parsed.length) return parsed;
  }

  return [normalized];
}

export async function copyLatex(
  latex: string,
  format: LatexCodeFormat = DEFAULT_LATEX_CODE_FORMAT,
) {
  await navigator.clipboard.writeText(formatLatex(latex, format));
}
