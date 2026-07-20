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
  const lines = splitLatexLines(latex)
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

export function formatLatex(latex: string, format: LatexCodeFormat): string {
  const lines = filledFormulaLines(latex);

  switch (format) {
    case "raw":
      return lines.join("\n");
    case "inline-dollar":
      return lines.map((line) => `$${line}$`).join("\n");
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
