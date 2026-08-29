import { createUuid } from "../../runtime/browserCompatibility";
import { splitFormulaEquationTag } from "../shared/formulaEquationTag.ts";

export type DocumentSourceFormat = "auto" | "markdown" | "latex";
export type ResolvedDocumentSourceFormat = Exclude<DocumentSourceFormat, "auto">;
export type DocumentObjectMode = "wordOmml" | "nativeOle" | "mathTypeOle";

export type DocumentImportRun =
  | {
      kind: "text";
      text: string;
      bold?: boolean;
      italic?: boolean;
      code?: boolean;
      strike?: boolean;
      underline?: boolean;
    }
  | {
      kind: "formula";
      latex: string;
      display: boolean;
      equationTag?: string;
    };

export type DocumentImportBlockKind =
  | "paragraph"
  | "heading"
  | "bullet"
  | "numbered"
  | "quote"
  | "code"
  | "display";

export interface DocumentImportBlock {
  id: string;
  kind: DocumentImportBlockKind;
  level: number;
  runs: DocumentImportRun[];
}

export interface ParsedDocumentImport {
  format: ResolvedDocumentSourceFormat;
  blocks: DocumentImportBlock[];
  warnings: string[];
  formulaCount: number;
  inlineFormulaCount: number;
  displayFormulaCount: number;
  textCharacterCount: number;
}

const displayEnvironmentPattern = /\\begin\{(equation\*?|align\*?|gather\*?|multline\*?|displaymath)\}/gi;

type TheoremBodyKind = "quote" | "normal";

interface TheoremEnvironmentDefinition {
  label: string;
  numbered: boolean;
  counterName: string;
  bodyKind: TheoremBodyKind;
}

interface TheoremMarkerPayload extends TheoremEnvironmentDefinition {
  note: string;
}

const theoremStartMarkerPrefix = "\uE410VT_THEOREM_START:";
const theoremStartMarkerSuffix = "\uE411";
const theoremEndMarker = "\uE412VT_THEOREM_END\uE413";

function encodeTheoremMarker(payload: TheoremMarkerPayload) {
  const encoded = encodeURIComponent(JSON.stringify(payload)).replace(/%/g, "§");
  return `${theoremStartMarkerPrefix}${encoded}${theoremStartMarkerSuffix}`;
}

function decodeTheoremMarker(value: string): TheoremMarkerPayload | null {
  if (!value.startsWith(theoremStartMarkerPrefix) || !value.endsWith(theoremStartMarkerSuffix)) {
    return null;
  }
  try {
    const payload = JSON.parse(
      decodeURIComponent(
        value
          .slice(theoremStartMarkerPrefix.length, -theoremStartMarkerSuffix.length)
          .replace(/§/g, "%"),
      ),
    ) as Partial<TheoremMarkerPayload>;
    if (
      typeof payload.label !== "string" ||
      typeof payload.numbered !== "boolean" ||
      typeof payload.counterName !== "string" ||
      (payload.bodyKind !== "quote" && payload.bodyKind !== "normal") ||
      typeof payload.note !== "string"
    ) {
      return null;
    }
    return payload as TheoremMarkerPayload;
  } catch {
    return null;
  }
}

function id() {
  return createUuid();
}

function isEscaped(text: string, index: number) {
  let slashes = 0;
  for (let cursor = index - 1; cursor >= 0 && text[cursor] === "\\"; cursor -= 1) {
    slashes += 1;
  }
  return slashes % 2 === 1;
}

function findUnescaped(text: string, token: string, start = 0) {
  for (let index = Math.max(0, start); index <= text.length - token.length; index += 1) {
    if (text.slice(index, index + token.length) === token && !isEscaped(text, index)) {
      return index;
    }
  }
  return -1;
}

function decodeText(text: string, format: ResolvedDocumentSourceFormat) {
  if (format === "markdown") {
    return text
      .replace(/\\([\\`*_{}\[\]()#+\-.!$])/g, "$1")
      .replace(/\uE000[ \t]*/g, "\n")
      .replace(/&nbsp;/gi, "\u00a0")
      .replace(/&amp;/gi, "&")
      .replace(/&lt;/gi, "<")
      .replace(/&gt;/gi, ">")
      .replace(/&quot;/gi, '"')
      .replace(/&#39;|&apos;/gi, "'")
      .replace(/&#(\d+);/g, (_match, value: string) =>
        String.fromCodePoint(Math.max(0, Math.min(0x10ffff, Number(value)))),
      )
      .replace(/&#x([0-9a-f]+);/gi, (_match, value: string) =>
        String.fromCodePoint(Math.max(0, Math.min(0x10ffff, Number.parseInt(value, 16)))),
      );
  }
  return text
    .replace(/~/g, "\u00a0")
    .replace(/\\ /g, "\u00a0")
    .replace(/\\%/g, "%")
    .replace(/\\_/g, "_")
    .replace(/\\&/g, "&")
    .replace(/\\#/g, "#")
    .replace(/\\\$/g, "$")
    .replace(/\\\{/g, "{")
    .replace(/\\\}/g, "}")
    .replace(/\\textbackslash\{\}/g, "\\")
    .replace(/\\newline/g, "\n")
    .replace(/\\\\/g, "\n");
}

function findMatchingBrace(text: string, open: number) {
  let depth = 0;
  for (let index = open; index < text.length; index += 1) {
    if (text[index] === "{" && !isEscaped(text, index)) depth += 1;
    if (text[index] === "}" && !isEscaped(text, index)) {
      depth -= 1;
      if (depth === 0) return index;
    }
  }
  return -1;
}

function findMarkdownClosingDelimiter(
  text: string,
  delimiter: "**" | "__" | "***" | "___",
  start: number,
) {
  let cursor = Math.max(0, start);
  while (cursor <= text.length - delimiter.length) {
    const found = text.indexOf(delimiter, cursor);
    if (found < 0) return -1;
    if (!isEscaped(text, found)) {
      if (
        delimiter.length === 2 &&
        text[found + delimiter.length] === delimiter[0]
      ) {
        return found + 1;
      }
      return found;
    }
    cursor = found + delimiter.length;
  }
  return -1;
}

function mergeTextRuns(runs: DocumentImportRun[]) {
  const merged: DocumentImportRun[] = [];
  for (const run of runs) {
    const previous = merged.at(-1);
    if (
      run.kind === "text" &&
      previous?.kind === "text" &&
      Boolean(previous.bold) === Boolean(run.bold) &&
      Boolean(previous.italic) === Boolean(run.italic) &&
      Boolean(previous.code) === Boolean(run.code) &&
      Boolean(previous.strike) === Boolean(run.strike) &&
      Boolean(previous.underline) === Boolean(run.underline)
    ) {
      previous.text += run.text;
    } else {
      merged.push(run);
    }
  }
  return merged;
}

function isTightLatexBoundaryCharacter(character: string | undefined) {
  if (!character) return false;
  return /[\u2e80-\u2fff\u3000-\u303f\u3040-\u30ff\u31c0-\u31ef\u3400-\u4dbf\u4e00-\u9fff\uac00-\ud7af\uf900-\ufaff\uff00-\uffef，。！？；：、）》】」』”’…,.!?;:()[\]{}]/u.test(
    character,
  );
}

function normalizeLatexInlineBoundaryWhitespace(runs: DocumentImportRun[]) {
  for (let index = 0; index < runs.length; index += 1) {
    const run = runs[index];
    if (run.kind !== "formula" || run.display) continue;

    const previous = runs[index - 1];
    if (previous?.kind === "text" && /[ \t]+$/.test(previous.text)) {
      const visible = previous.text.replace(/[ \t]+$/, "");
      if (isTightLatexBoundaryCharacter(visible.at(-1))) previous.text = visible;
    }

    const next = runs[index + 1];
    if (next?.kind === "text" && /^[ \t]+/.test(next.text)) {
      const visible = next.text.replace(/^[ \t]+/, "");
      if (isTightLatexBoundaryCharacter(visible[0])) next.text = visible;
    }
  }
  return runs.filter((run) => run.kind === "formula" || run.text.length > 0);
}

function parseInline(
  text: string,
  format: ResolvedDocumentSourceFormat,
  inherited: {
    bold?: boolean;
    italic?: boolean;
    code?: boolean;
    strike?: boolean;
    underline?: boolean;
  } = {},
): DocumentImportRun[] {
  const runs: DocumentImportRun[] = [];
  let buffer = "";
  const flush = () => {
    if (!buffer) return;
    runs.push({ kind: "text", text: decodeText(buffer, format), ...inherited });
    buffer = "";
  };

  for (let index = 0; index < text.length; ) {
    if (text[index] === "$" && !isEscaped(text, index) && text[index + 1] !== "$") {
      const end = findUnescaped(text, "$", index + 1);
      if (end > index + 1) {
        flush();
        runs.push({ kind: "formula", latex: text.slice(index + 1, end).trim(), display: false });
        index = end + 1;
        continue;
      }
    }
    if (text.startsWith("\\(", index) && !isEscaped(text, index)) {
      const end = findUnescaped(text, "\\)", index + 2);
      if (end > index + 2) {
        flush();
        runs.push({ kind: "formula", latex: text.slice(index + 2, end).trim(), display: false });
        index = end + 2;
        continue;
      }
    }

    if (format === "markdown") {
      const rest = text.slice(index);
      const image = rest.match(/^!\[([^\]]*)\]\((\S+?)(?:\s+["'][^"']*["'])?\)/);
      if (image) {
        flush();
        const alt = decodeText(image[1], format).trim() || "未命名图片";
        runs.push({
          kind: "text",
          text: `【图片：${alt}（${image[2]}）】`,
          ...inherited,
        });
        index += image[0].length;
        continue;
      }
      const link = rest.match(/^\[([^\]]+)\]\((\S+?)(?:\s+["'][^"']*["'])?\)/);
      if (link) {
        flush();
        runs.push(...parseInline(link[1], format, inherited));
        runs.push({ kind: "text", text: `（${link[2]}）`, ...inherited });
        index += link[0].length;
        continue;
      }
      const autoLink = rest.match(/^<(https?:\/\/[^>]+|mailto:[^>]+)>/i);
      if (autoLink) {
        flush();
        runs.push({ kind: "text", text: autoLink[1], ...inherited });
        index += autoLink[0].length;
        continue;
      }
      if (rest.startsWith("<!--")) {
        const end = rest.indexOf("-->", 4);
        flush();
        index += end >= 0 ? end + 3 : rest.length;
        continue;
      }
      const htmlBreak = rest.match(/^<br\s*\/?\s*>/i);
      if (htmlBreak) {
        flush();
        runs.push({ kind: "text", text: "\n", ...inherited });
        index += htmlBreak[0].length;
        continue;
      }
      const htmlTag = rest.match(/^<\/?[A-Za-z][^>]*>/);
      if (htmlTag) {
        flush();
        index += htmlTag[0].length;
        continue;
      }
      if (text.startsWith("~~", index) && !isEscaped(text, index)) {
        const end = findUnescaped(text, "~~", index + 2);
        if (end > index + 2) {
          flush();
          runs.push(
            ...parseInline(text.slice(index + 2, end), format, {
              ...inherited,
              strike: true,
            }),
          );
          index = end + 2;
          continue;
        }
      }
      const tripleDelimiter = text.startsWith("***", index)
        ? "***"
        : text.startsWith("___", index)
          ? "___"
          : null;
      if (tripleDelimiter && !isEscaped(text, index)) {
        const end = findMarkdownClosingDelimiter(
          text,
          tripleDelimiter,
          index + tripleDelimiter.length,
        );
        if (end > index + tripleDelimiter.length) {
          flush();
          runs.push(
            ...parseInline(text.slice(index + tripleDelimiter.length, end), format, {
              ...inherited,
              bold: true,
              italic: true,
            }),
          );
          index = end + tripleDelimiter.length;
          continue;
        }
      }
      const strongDelimiter = text.startsWith("**", index)
        ? "**"
        : text.startsWith("__", index)
          ? "__"
          : null;
      if (strongDelimiter && !isEscaped(text, index)) {
        const end = findMarkdownClosingDelimiter(
          text,
          strongDelimiter,
          index + strongDelimiter.length,
        );
        if (end > index + strongDelimiter.length) {
          flush();
          runs.push(
            ...parseInline(text.slice(index + strongDelimiter.length, end), format, {
              ...inherited,
              bold: true,
            }),
          );
          index = end + strongDelimiter.length;
          continue;
        }
      }
      if ((text[index] === "*" || text[index] === "_") && !isEscaped(text, index)) {
        const end = findUnescaped(text, text[index], index + 1);
        if (end > index + 1) {
          flush();
          runs.push(...parseInline(text.slice(index + 1, end), format, { ...inherited, italic: true }));
          index = end + 1;
          continue;
        }
      }
      if (text[index] === "`") {
        let delimiterLength = 1;
        while (text[index + delimiterLength] === "`") delimiterLength += 1;
        const delimiter = "`".repeat(delimiterLength);
        const end = text.indexOf(delimiter, index + delimiterLength);
        if (end >= index + delimiterLength) {
          flush();
          let codeText = text.slice(index + delimiterLength, end);
          if (codeText.startsWith(" ") && codeText.endsWith(" ") && codeText.trim()) {
            codeText = codeText.slice(1, -1);
          }
          runs.push({ kind: "text", text: codeText, ...inherited, code: true });
          index = end + delimiterLength;
          continue;
        }
      }
    } else if (text[index] === "\\") {
      if (text.startsWith("\\verb", index)) {
        let delimiterIndex = index + "\\verb".length;
        if (text[delimiterIndex] === "*") delimiterIndex += 1;
        const delimiter = text[delimiterIndex];
        if (delimiter && !/[A-Za-z0-9\s]/.test(delimiter)) {
          const end = text.indexOf(delimiter, delimiterIndex + 1);
          if (end > delimiterIndex) {
            flush();
            runs.push({
              kind: "text",
              text: text.slice(delimiterIndex + 1, end),
              ...inherited,
              code: true,
            });
            index = end + 1;
            continue;
          }
        }
      }
      const commands = [
        { name: "\\textbf{", style: { bold: true } },
        { name: "\\textit{", style: { italic: true } },
        { name: "\\emph{", style: { italic: true } },
        { name: "\\texttt{", style: { code: true } },
        { name: "\\underline{", style: { underline: true } },
        { name: "\\uline{", style: { underline: true } },
        { name: "\\sout{", style: { strike: true } },
      ];
      const command = commands.find((candidate) => text.startsWith(candidate.name, index));
      if (command) {
        const open = index + command.name.length - 1;
        const close = findMatchingBrace(text, open);
        if (close > open) {
          flush();
          runs.push(
            ...parseInline(text.slice(open + 1, close), format, {
              ...inherited,
              ...command.style,
            }),
          );
          index = close + 1;
          continue;
        }
      }

      const latexCommand = text.slice(index).match(/^\\([A-Za-z@]+)\*?/);
      if (latexCommand) {
        const name = latexCommand[1].toLowerCase();
        let cursor = index + latexCommand[0].length;
        while (text[cursor] === " " || text[cursor] === "\t") cursor += 1;

        const skipOptionalArguments = (start: number) => {
          let position = start;
          while (text[position] === " " || text[position] === "\t") position += 1;
          while (text[position] === "[") {
            let depth = 0;
            let close = -1;
            for (let scan = position; scan < text.length; scan += 1) {
              if (text[scan] === "[" && !isEscaped(text, scan)) depth += 1;
              if (text[scan] === "]" && !isEscaped(text, scan)) {
                depth -= 1;
                if (depth === 0) {
                  close = scan;
                  break;
                }
              }
            }
            if (close < 0) break;
            position = close + 1;
            while (text[position] === " " || text[position] === "\t") position += 1;
          }
          return position;
        };
        const readMandatoryArguments = (start: number, count: number) => {
          const arguments_: string[] = [];
          let position = skipOptionalArguments(start);
          while (arguments_.length < count && text[position] === "{") {
            const close = findMatchingBrace(text, position);
            if (close <= position) break;
            arguments_.push(text.slice(position + 1, close));
            position = skipOptionalArguments(close + 1);
          }
          return { arguments_, nextIndex: position };
        };

        const visibleMultiArgumentCommands: Record<string, { count: number; visible: number }> = {
          textcolor: { count: 2, visible: 1 },
          colorbox: { count: 2, visible: 1 },
          fcolorbox: { count: 3, visible: 2 },
          scalebox: { count: 2, visible: 1 },
          resizebox: { count: 3, visible: 2 },
          rotatebox: { count: 2, visible: 1 },
          parbox: { count: 2, visible: 1 },
          raisebox: { count: 2, visible: 1 },
          makebox: { count: 1, visible: 0 },
          framebox: { count: 1, visible: 0 },
        };
        const visibleMultiArgument = visibleMultiArgumentCommands[name];
        if (visibleMultiArgument) {
          const parsedArguments = readMandatoryArguments(
            cursor,
            visibleMultiArgument.count,
          );
          const visible = parsedArguments.arguments_[visibleMultiArgument.visible];
          if (visible !== undefined) {
            flush();
            runs.push(...parseInline(visible, format, inherited));
            index = parsedArguments.nextIndex;
            continue;
          }
        }

        if (name === "href" && text[cursor] === "{") {
          const urlClose = findMatchingBrace(text, cursor);
          const labelOpen = urlClose >= 0 ? urlClose + 1 : -1;
          if (urlClose > cursor && text[labelOpen] === "{") {
            const labelClose = findMatchingBrace(text, labelOpen);
            if (labelClose > labelOpen) {
              flush();
              runs.push(...parseInline(text.slice(labelOpen + 1, labelClose), format, inherited));
              runs.push({
                kind: "text",
                text: `（${decodeText(text.slice(cursor + 1, urlClose), format)}）`,
                ...inherited,
              });
              index = labelClose + 1;
              continue;
            }
          }
        }

        if (text[cursor] === "{") {
          const close = findMatchingBrace(text, cursor);
          if (close > cursor) {
            const argument = text.slice(cursor + 1, close);
            flush();
            if (name === "textbackslash") {
              runs.push({ kind: "text", text: "\\", ...inherited });
            } else if (name === "url") {
              runs.push({ kind: "text", text: decodeText(argument, format), ...inherited });
            } else if (name === "footnote" || name === "thanks") {
              runs.push({ kind: "text", text: "（注：", ...inherited });
              runs.push(...parseInline(argument, format, inherited));
              runs.push({ kind: "text", text: "）", ...inherited });
            } else if (name === "cite" || name === "citep" || name === "citet") {
              runs.push({ kind: "text", text: `[${decodeText(argument, format)}]`, ...inherited });
            } else if (name === "ref" || name === "eqref" || name === "pageref" || name === "autoref") {
              runs.push({ kind: "text", text: `（${decodeText(argument, format)}）`, ...inherited });
            } else if (name === "input" || name === "include" || name === "subfile") {
              runs.push({
                kind: "text",
                text: `【外部 LaTeX 文件：${decodeText(argument, format)}】`,
                ...inherited,
              });
            } else if (name === "bibliography" || name === "addbibresource") {
              runs.push({
                kind: "text",
                text: `【参考文献数据：${decodeText(argument, format)}】`,
                ...inherited,
              });
            } else if (name === "bibliographystyle") {
              // Bibliography styling has no visible body in Word.
            } else if (name === "label" || name === "index" || name === "glossary") {
              // Structural metadata has no visible body in Word.
            } else if (name === "includegraphics") {
              runs.push({ kind: "text", text: `【图片：${decodeText(argument, format)}】`, ...inherited });
            } else if (name === "caption") {
              runs.push({ kind: "text", text: "图注：", ...inherited, bold: true });
              runs.push(...parseInline(argument, format, inherited));
            } else {
              // Unknown formatting commands commonly wrap their visible text in
              // the first mandatory argument. Preserve that body instead of
              // leaking the control sequence or dropping user content.
              runs.push(...parseInline(argument, format, inherited));
            }
            index = close + 1;
            continue;
          }
        }

        const visibleCommands: Record<string, string> = {
          latex: "LaTeX",
          tex: "TeX",
          ldots: "…",
          dots: "…",
          textemdash: "—",
          textendash: "–",
          textquotedblleft: "“",
          textquotedblright: "”",
          textquoteleft: "‘",
          textquoteright: "’",
          quad: "\u00a0",
          qquad: "\u00a0\u00a0",
          smallskip: "\n",
          medskip: "\n",
          bigskip: "\n",
          par: "\n",
          tiny: "",
          scriptsize: "",
          footnotesize: "",
          small: "",
          normalsize: "",
          large: "",
          huge: "",
          centering: "",
          raggedright: "",
          raggedleft: "",
        };
        if (Object.prototype.hasOwnProperty.call(visibleCommands, name)) {
          flush();
          runs.push({ kind: "text", text: visibleCommands[name], ...inherited });
          index = cursor;
          continue;
        }
      }

      const accent = text.slice(index).match(/^\\(["'`^~=.uvHckbdtr])\s*\{/);
      if (accent) {
        const open = index + accent[0].length - 1;
        const close = findMatchingBrace(text, open);
        if (close > open) {
          flush();
          runs.push(...parseInline(text.slice(open + 1, close), format, inherited));
          index = close + 1;
          continue;
        }
      }
    }

    buffer += text[index];
    index += 1;
  }
  flush();
  const merged = mergeTextRuns(runs);
  return format === "latex"
    ? normalizeLatexInlineBoundaryWhitespace(merged)
    : merged;
}

function normalizeDisplayFormulaBody(body: string) {
  return body
    .replace(/\\label\s*\{[^{}]*\}/gi, "")
    .replace(/\\(?:notag|nonumber)\b/gi, "");
}

function normalizeDisplayEnvironment(environment: string, body: string) {
  // Physical source line breaks inside TeX math are ordinary whitespace.
  // Preserve only explicit mathematical row breaks such as `\\\\`; otherwise
  // a prettified equation/equation* source is incorrectly reopened as several
  // VisualTeX rows, and a complete aligned environment can be truncated by a
  // single-line editor field.
  const normalized = normalizeDisplayFormulaBody(body)
    .replace(/\r\n?/g, "\n")
    .replace(/[ \t]*\n+[ \t]*/g, " ")
    .trim();
  switch (environment.replace(/\*$/, "").toLowerCase()) {
    case "align":
      return `\\begin{aligned}${normalized}\\end{aligned}`;
    case "gather":
    case "multline":
      return `\\begin{gathered}${normalized}\\end{gathered}`;
    default:
      return normalized;
  }
}

function normalizeDelimitedDisplay(body: string) {
  // Newlines inside $$...$$ and \\[...\\] are TeX whitespace, not
  // VisualTeX row boundaries. Collapsing them keeps paired \\left/\\right
  // delimiters in one expression while explicit \\\\ row breaks remain.
  return normalizeDisplayFormulaBody(body)
    .replace(/\r\n?/g, "\n")
    .replace(/[ \t]*\n+[ \t]*/g, " ")
    .trim();
}

interface DisplayStart {
  position: number;
  startToken: string;
  endToken: string;
  environment?: string;
}

function isLikelyMarkdownBracketMath(text: string, start: number) {
  const contentStart = start + 2;
  const end = findUnescaped(text, "\\]", contentStart);
  if (end < 0) return true;
  const body = text.slice(contentStart, end).trim();
  if (!body) return false;
  return (
    body.includes("\n") ||
    /[\\=+*/_^{}<>0-9&]/.test(body) ||
    !/\s/.test(body)
  );
}

function findDisplayStart(text: string, format: ResolvedDocumentSourceFormat, from: number): DisplayStart | null {
  const candidates: DisplayStart[] = [];
  const dollars = findUnescaped(text, "$$", from);
  if (dollars >= 0) candidates.push({ position: dollars, startToken: "$$", endToken: "$$" });
  const bracket = findUnescaped(text, "\\[", from);
  if (
    bracket >= 0 &&
    (format === "latex" || isLikelyMarkdownBracketMath(text, bracket))
  ) {
    candidates.push({ position: bracket, startToken: "\\[", endToken: "\\]" });
  }
  if (format === "latex") {
    displayEnvironmentPattern.lastIndex = from;
    const match = displayEnvironmentPattern.exec(text);
    if (match && !isEscaped(text, match.index)) {
      candidates.push({
        position: match.index,
        startToken: match[0],
        endToken: `\\end{${match[1]}}`,
        environment: match[1],
      });
    }
  }
  return candidates.sort((left, right) => left.position - right.position)[0] ?? null;
}

function textBlock(
  kind: Exclude<DocumentImportBlockKind, "display" | "code">,
  text: string,
  format: ResolvedDocumentSourceFormat,
  level = 0,
): DocumentImportBlock | null {
  const normalized = text.replace(/\s*\n\s*/g, " ").trim();
  if (!normalized) return null;
  return { id: id(), kind, level, runs: parseInline(normalized, format) };
}

function appendMixedBlocks(
  blocks: DocumentImportBlock[],
  text: string,
  format: ResolvedDocumentSourceFormat,
  warnings: string[],
  textKind: Exclude<DocumentImportBlockKind, "display" | "code"> = "paragraph",
  level = 0,
) {
  let cursor = 0;
  while (cursor < text.length) {
    const start = findDisplayStart(text, format, cursor);
    if (!start) {
      const block = textBlock(textKind, text.slice(cursor), format, level);
      if (block) blocks.push(block);
      return;
    }
    const before = textBlock(textKind, text.slice(cursor, start.position), format, level);
    if (before) blocks.push(before);
    const contentStart = start.position + start.startToken.length;
    const end = findUnescaped(text, start.endToken, contentStart);
    if (end < 0) {
      const body = text.slice(contentStart);
      const split = splitFormulaEquationTag(
        start.environment
          ? normalizeDisplayEnvironment(start.environment, body)
          : normalizeDelimitedDisplay(body),
      );
      blocks.push({
        id: id(),
        kind: "display",
        level: 0,
        runs: [
          {
            kind: "formula",
            latex: split.latex,
            display: true,
            ...(split.equationTag ? { equationTag: split.equationTag } : {}),
          },
        ],
      });
      warnings.push(
        start.environment
          ? `LaTeX 环境 ${start.environment} 未闭合，预览已读取到文末。`
          : `行间公式缺少结束标记 ${start.endToken}，预览已读取到文末。`,
      );
      return;
    }
    const body = text.slice(contentStart, end);
    const split = splitFormulaEquationTag(
      start.environment
        ? normalizeDisplayEnvironment(start.environment, body)
        : normalizeDelimitedDisplay(body),
    );
    blocks.push({
      id: id(),
      kind: "display",
      level: 0,
      runs: [
        {
          kind: "formula",
          latex: split.latex,
          display: true,
          ...(split.equationTag ? { equationTag: split.equationTag } : {}),
        },
      ],
    });
    cursor = end + start.endToken.length;
  }
}

function detectFormat(source: string): ResolvedDocumentSourceFormat {
  // A LaTeX fragment may consist only of a proof/theorem/custom environment or
  // a preamble macro definition. Do not require one of a tiny fixed set of
  // environments before enabling the LaTeX parser.
  return /\\(?:documentclass|usepackage|RequirePackage|newcommand|renewcommand|providecommand|DeclareMathOperator\*?|DeclarePairedDelimiter\w*|newtheorem\*?|begin\{[A-Za-z@*]+\}|\[|\(|text(?:bf|it|tt)\{|emph\{|item(?:\s|\[)|(?:part|chapter|section|subsection|subsubsection|paragraph|subparagraph)\*?\{)/i.test(
    source,
  )
    ? "latex"
    : "markdown";
}

function markdownTableCells(line: string) {
  let normalized = line.trim();
  if (normalized.startsWith("|")) normalized = normalized.slice(1);
  if (normalized.endsWith("|")) normalized = normalized.slice(0, -1);
  const cells: string[] = [];
  let buffer = "";
  for (let index = 0; index < normalized.length; index += 1) {
    if (normalized[index] === "|" && !isEscaped(normalized, index)) {
      cells.push(buffer.trim());
      buffer = "";
    } else {
      buffer += normalized[index];
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

function normalizeMarkdownSource(source: string, warnings: string[]) {
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
    const reference = line.match(/^\s*\[([^\]^][^\]]*)\]:\s*(\S+)(?:\s+["'(].*["')])?\s*$/);
    if (reference) {
      references.set(reference[1].trim().toLowerCase(), reference[2]);
      continue;
    }
    retained.push(line);
  }

  const replaceReferences = (line: string) => {
    let next = line.replace(/!\[([^\]]*)\]\[([^\]]*)\]/g, (match, alt: string, key: string) => {
      const target = references.get((key || alt).trim().toLowerCase());
      return target ? `![${alt}](${target})` : match;
    });
    next = next.replace(/\[([^\]]+)\]\[([^\]]*)\]/g, (match, label: string, key: string) => {
      const target = references.get((key || label).trim().toLowerCase());
      return target ? `[${label}](${target})` : match;
    });
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
    const end = retained.slice(1).findIndex((line) => /^(?:---|\.\.\.)\s*$/.test(line.trim()));
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
      output.push(header.map((cell) => `**${replaceReferences(cell)}**`).join(" ｜ "));
      index += 2;
      while (index < retained.length && retained[index].includes("|") && retained[index].trim()) {
        output.push(
          markdownTableCells(retained[index])
            .map((cell) => replaceReferences(cell))
            .join(" ｜ "),
        );
        index += 1;
      }
      output.push("");
      index -= 1;
      continue;
    }

    if (/^ {0,3}(?:\*\s*){3,}$/.test(raw) || /^ {0,3}(?:-\s*){3,}$/.test(raw) || /^ {0,3}(?:_\s*){3,}$/.test(raw)) {
      output.push("────────────────────");
      continue;
    }

    const task = raw.match(/^(\s*[-+*]\s+)\[([ xX])\]\s+(.*)$/);
    if (task) {
      output.push(`${task[1]}${task[2].toLowerCase() === "x" ? "☒" : "☐"} ${replaceReferences(task[3])}`);
      continue;
    }

    const hardBreak = / {2,}$/.test(raw);
    const normalizedLine = replaceReferences(hardBreak ? raw.replace(/ {2,}$/, "") : raw);
    output.push(hardBreak ? `${normalizedLine}\uE000` : normalizedLine);
  }

  if (footnotes.size) {
    output.push("");
    for (const [key, value] of footnotes) {
      output.push(`**注 ${key}：** ${replaceReferences(value)}`, "");
    }
  }
  if (references.size) {
    warnings.push(`已解析 ${references.size} 个 Markdown 引用式链接。`);
  }
  if (footnotes.size) {
    warnings.push(`已将 ${footnotes.size} 个 Markdown 脚注转换为 Word 注释段落。`);
  }
  return output.join("\n");
}

function replaceLatexTableEnvironment(body: string, warnings: string[]) {
  let converted = 0;
  const result = body.replace(
    /\\begin\{tabular\}\s*\{[^{}]*\}([\s\S]*?)\\end\{tabular\}/gi,
    (_match, tableBody: string) => {
      converted += 1;
      const normalized = tableBody
        .replace(/\\(?:toprule|midrule|bottomrule|hline)\b/g, "")
        .replace(/\\multicolumn\{[^{}]*\}\{[^{}]*\}\{([^{}]*)\}/g, "$1")
        .replace(/\\multirow\{[^{}]*\}\{[^{}]*\}\{([^{}]*)\}/g, "$1");
      return normalized
        .split(/\\\\(?:\[[^\]]*\])?/)
        .map((row) => row.trim())
        .filter(Boolean)
        .map((row) => row.replace(/(?<!\\)&/g, " ｜ ").replace(/\\&/g, "&"))
        .join("\n\n");
    },
  );
  if (converted) {
    warnings.push(`已将 ${converted} 个 LaTeX tabular 表格转换为可编辑的 Word 文本行。`);
  }
  return result;
}

type LatexDocumentMacro = {
  argumentCount: number;
  replacement: string;
  optionalDefault?: string;
};

function skipLatexSpaces(source: string, index: number) {
  while (index < source.length && /\s/.test(source[index])) index += 1;
  return index;
}

function readLatexDelimitedGroup(
  source: string,
  start: number,
  open: "{" | "[",
  close: "}" | "]",
) {
  if (source[start] !== open) return null;
  let depth = 0;
  for (let index = start; index < source.length; index += 1) {
    if (source[index] === "\\") {
      index += 1;
      continue;
    }
    if (source[index] === open) depth += 1;
    else if (source[index] === close) {
      depth -= 1;
      if (depth === 0) {
        return {
          content: source.slice(start + 1, index),
          end: index + 1,
        };
      }
    }
  }
  return null;
}

function readLatexMacroArgument(source: string, start: number) {
  const position = skipLatexSpaces(source, start);
  if (source[position] === "{") {
    return readLatexDelimitedGroup(source, position, "{", "}");
  }
  if (source[position] === "\\") {
    const command = source.slice(position).match(/^\\(?:[A-Za-z@]+|.)/);
    return command
      ? { content: command[0], end: position + command[0].length }
      : null;
  }
  return position < source.length
    ? { content: source[position], end: position + 1 }
    : null;
}

function collectLatexDocumentMacros(
  source: string,
  macros: Map<string, LatexDocumentMacro>,
) {
  const definitionStart = /\\(?:newcommand|renewcommand|providecommand|def|DeclareMathOperator|DeclarePairedDelimiter(?:X)?)(?:\*?)(?![A-Za-z@])/g;
  let output = "";
  let cursor = 0;
  while (cursor < source.length) {
    definitionStart.lastIndex = cursor;
    const candidate = definitionStart.exec(source);
    if (!candidate) {
      output += source.slice(cursor);
      break;
    }
    output += source.slice(cursor, candidate.index);
    const start = candidate.index;
    const command = candidate[0].replace(/\*$/, "");
    let position = start + candidate[0].length;
    let parsed = false;

    if (/^\\(?:newcommand|renewcommand|providecommand)$/.test(command)) {
      position = skipLatexSpaces(source, position);
      const nameGroup = readLatexDelimitedGroup(source, position, "{", "}");
      const nameMatch = nameGroup?.content.match(/^\\([A-Za-z@]+)$/);
      if (nameGroup && nameMatch) {
        position = skipLatexSpaces(source, nameGroup.end);
        let argumentCount = 0;
        let optionalDefault: string | undefined;
        if (source[position] === "[") {
          const countGroup = readLatexDelimitedGroup(source, position, "[", "]");
          if (countGroup && /^\d+$/.test(countGroup.content.trim())) {
            argumentCount = Math.max(0, Math.min(9, Number(countGroup.content.trim())));
            position = skipLatexSpaces(source, countGroup.end);
            if (source[position] === "[") {
              const defaultGroup = readLatexDelimitedGroup(source, position, "[", "]");
              if (defaultGroup) {
                optionalDefault = defaultGroup.content;
                argumentCount = Math.max(1, argumentCount);
                position = skipLatexSpaces(source, defaultGroup.end);
              }
            }
          }
        }
        const replacement = readLatexDelimitedGroup(source, position, "{", "}");
        if (replacement) {
          macros.set(nameMatch[1], {
            argumentCount,
            replacement: replacement.content,
            ...(optionalDefault !== undefined ? { optionalDefault } : {}),
          });
          cursor = replacement.end;
          parsed = true;
        }
      }
    } else if (command === "\\def") {
      position = skipLatexSpaces(source, position);
      const nameMatch = source.slice(position).match(/^\\([A-Za-z@]+)/);
      if (nameMatch) {
        position += nameMatch[0].length;
        let argumentCount = 0;
        while (true) {
          position = skipLatexSpaces(source, position);
          const parameter = source.slice(position).match(/^#([1-9])/);
          if (!parameter) break;
          argumentCount = Math.max(argumentCount, Number(parameter[1]));
          position += parameter[0].length;
        }
        position = skipLatexSpaces(source, position);
        const replacement = readLatexDelimitedGroup(source, position, "{", "}");
        if (replacement) {
          macros.set(nameMatch[1], {
            argumentCount,
            replacement: replacement.content,
          });
          cursor = replacement.end;
          parsed = true;
        }
      }
    } else if (command === "\\DeclareMathOperator") {
      position = skipLatexSpaces(source, position);
      const nameGroup = readLatexDelimitedGroup(source, position, "{", "}");
      const nameMatch = nameGroup?.content.match(/^\\([A-Za-z@]+)$/);
      if (nameGroup && nameMatch) {
        position = skipLatexSpaces(source, nameGroup.end);
        const label = readLatexDelimitedGroup(source, position, "{", "}");
        if (label) {
          macros.set(nameMatch[1], {
            argumentCount: 0,
            replacement: `${candidate[0].endsWith("*") ? "\\operatorname*" : "\\operatorname"}{${label.content}}`,
          });
          cursor = label.end;
          parsed = true;
        }
      }
    } else if (command === "\\DeclarePairedDelimiter" || command === "\\DeclarePairedDelimiterX") {
      position = skipLatexSpaces(source, position);
      const nameGroup = readLatexDelimitedGroup(source, position, "{", "}");
      const nameMatch = nameGroup?.content.match(/^\\([A-Za-z@]+)$/);
      if (nameGroup && nameMatch) {
        position = skipLatexSpaces(source, nameGroup.end);
        let argumentCount = command === "\\DeclarePairedDelimiterX" ? 1 : 1;
        if (source[position] === "[") {
          const countGroup = readLatexDelimitedGroup(source, position, "[", "]");
          if (countGroup && /^\d+$/.test(countGroup.content.trim())) {
            argumentCount = Math.max(1, Math.min(9, Number(countGroup.content.trim())));
            position = skipLatexSpaces(source, countGroup.end);
          }
        }
        const left = readLatexDelimitedGroup(source, position, "{", "}");
        position = left ? skipLatexSpaces(source, left.end) : position;
        const right = left ? readLatexDelimitedGroup(source, position, "{", "}") : null;
        position = right ? skipLatexSpaces(source, right.end) : position;
        const bodyGroup = command === "\\DeclarePairedDelimiterX" && right
          ? readLatexDelimitedGroup(source, position, "{", "}")
          : null;
        if (left && right && (command !== "\\DeclarePairedDelimiterX" || bodyGroup)) {
          macros.set(nameMatch[1], {
            argumentCount,
            replacement: `\\left${left.content}${bodyGroup?.content ?? "#1"}\\right${right.content}`,
          });
          cursor = bodyGroup?.end ?? right.end;
          parsed = true;
        }
      }
    }

    if (!parsed) {
      output += source[start];
      cursor = start + 1;
    }
  }
  return output;
}

function expandLatexDocumentMacro(
  source: string,
  name: string,
  macro: LatexDocumentMacro,
) {
  const token = `\\${name}`;
  let output = "";
  let cursor = 0;
  let changed = false;
  while (cursor < source.length) {
    const start = source.indexOf(token, cursor);
    if (start < 0) {
      output += source.slice(cursor);
      break;
    }
    output += source.slice(cursor, start);
    const boundary = source[start + token.length];
    if (isEscaped(source, start) || /[A-Za-z@]/.test(boundary ?? "")) {
      output += token;
      cursor = start + token.length;
      continue;
    }

    let position = start + token.length;
    if (source[position] === "*") position += 1;
    const arguments_: string[] = [];
    if (macro.optionalDefault !== undefined) {
      position = skipLatexSpaces(source, position);
      if (source[position] === "[") {
        const optional = readLatexDelimitedGroup(source, position, "[", "]");
        if (!optional) {
          output += token;
          cursor = start + token.length;
          continue;
        }
        arguments_.push(optional.content);
        position = optional.end;
      } else {
        arguments_.push(macro.optionalDefault);
      }
    }

    let valid = true;
    while (arguments_.length < macro.argumentCount) {
      const argument = readLatexMacroArgument(source, position);
      if (!argument) {
        valid = false;
        break;
      }
      arguments_.push(argument.content);
      position = argument.end;
    }
    if (!valid) {
      output += token;
      cursor = start + token.length;
      continue;
    }

    output += macro.replacement.replace(/#([1-9])/g, (_match, index: string) =>
      arguments_[Number(index) - 1] ?? "",
    );
    cursor = position;
    changed = true;
  }
  return { source: output, changed };
}

function normalizeLatexExtensions(source: string, warnings: string[]) {
  let body = source.replace(/\r\n?/g, "\n");
  const macros = new Map<string, LatexDocumentMacro>();
  body = collectLatexDocumentMacros(body, macros);

  for (let pass = 0; pass < 12; pass += 1) {
    let changed = false;
    for (const [name, macro] of macros) {
      const expanded = expandLatexDocumentMacro(body, name, macro);
      body = expanded.source;
      changed ||= expanded.changed;
    }
    if (!changed) break;
  }
  if (macros.size) {
    warnings.push(`已展开 ${macros.size} 个 LaTeX 自定义宏（支持嵌套内容、默认参数和最多九个参数）。`);
  }

  const takeCommand = (name: string) => {
    let value = "";
    const pattern = new RegExp(
      `\\\\${name}\\s*\\{((?:[^{}]|\\{[^{}]*\\})*)\\}`,
      "i",
    );
    body = body.replace(pattern, (_match, content: string) => {
      value = content.trim();
      return "";
    });
    return value;
  };
  body = body
    .replace(/\\(?:documentclass|usepackage|RequirePackage)(?:\s*\[[^\]]*\])?\s*\{[^{}]*\}/gi, "")
    .replace(/\\(?:geometry|hypersetup|graphicspath)\s*\{(?:[^{}]|\{[^{}]*\})*\}/gi, "")
    .replace(/\\(?:pagestyle|thispagestyle|bibliographystyle|setcounter)\s*\{[^{}]*\}(?:\s*\{[^{}]*\})?/gi, "")
    .replace(/\\(?:setlength|addtolength)\s*\{[^{}]*\}\s*\{[^{}]*\}/gi, "")
    .replace(/\\(?:allowdisplaybreaks|sloppy|fussy)\b(?:\s*\[[^\]]*\])?/gi, "");

  const title = takeCommand("title");
  const author = takeCommand("author");
  const date = takeCommand("date");
  body = body.replace(/\\maketitle\b/g, () => {
    const parts: string[] = [];
    if (title) parts.push(`\\section*{${title}}`);
    if (author) parts.push(`\\textit{${author}}`);
    if (date) parts.push(date);
    return parts.join("\n\n");
  });

  body = body
    .replace(
      /\\(part|chapter|section|subsection|subsubsection|paragraph|subparagraph)(\*?)\s*\[[^\]]*\]\s*\{/gi,
      "\\$1$2{",
    )
    .replace(/\\begin\{(?:itemize|enumerate)\}\s*\[[^\]]*\]/gi, (match) =>
      match.replace(/\s*\[[^\]]*\]$/, ""),
    )
    .replace(/\\begin\{minipage\}(?:\[[^\]]*\])?\s*\{[^{}]*\}/gi, "")
    .replace(/\\end\{minipage\}/gi, "")
    .replace(/\\begin\{thebibliography\}\s*\{[^{}]*\}/gi, "\\section*{参考文献}\n\\begin{itemize}\n")
    .replace(/\\end\{thebibliography\}/gi, "\n\\end{itemize}")
    .replace(/\\bibitem(?:\[([^\]]+)\])?\{([^{}]+)\}/gi, (_match, label: string | undefined, key: string) =>
      `\\item \\textbf{[${label || key}]} `,
    )
    .replace(/\\begin\{flalign(\*?)\}/gi, "\\begin{align$1}")
    .replace(/\\end\{flalign(\*?)\}/gi, "\\end{align$1}")
    .replace(/\\begin\{alignat(\*?)\}\s*\{[^{}]*\}/gi, "\\begin{align$1}")
    .replace(/\\end\{alignat(\*?)\}/gi, "\\end{align$1}")
    .replace(/\\begin\{eqnarray(\*?)\}/gi, "\\begin{align$1}")
    .replace(/\\end\{eqnarray(\*?)\}/gi, "\\end{align$1}")
    .replace(/\\begin\{math\}/gi, "\\(")
    .replace(/\\end\{math\}/gi, "\\)")
    .replace(/\\begin\{description\}/gi, "\\begin{itemize}\n")
    .replace(/\\end\{description\}/gi, "\n\\end{itemize}")
    .replace(/\\item\s*\[([^\]]+)\]/g, "\\item \\textbf{$1}：")
    .replace(/\\begin\{(?:center|flushleft|flushright|figure\*?|table\*?)\}/gi, "")
    .replace(/\\end\{(?:center|flushleft|flushright|figure\*?|table\*?)\}/gi, "")
    .replace(/\\includegraphics(?:\[[^\]]*\])?\{([^{}]+)\}/gi, "【图片：$1】")
    .replace(/\\caption\s*\{((?:[^{}]|\{[^{}]*\})*)\}/gi, "\\textbf{图注：} $1")
    .replace(/\\begin\{abstract\}/gi, "\\section*{摘要}")
    .replace(/\\end\{abstract\}/gi, "");

  const theoremDefinitions = new Map<string, TheoremEnvironmentDefinition>();
  const builtInTheorems: Array<
    [string, string, boolean, string, TheoremBodyKind]
  > = [
    ["theorem", "定理", true, "theorem", "quote"],
    ["lemma", "引理", true, "lemma", "quote"],
    ["proposition", "命题", true, "proposition", "quote"],
    ["corollary", "推论", true, "corollary", "quote"],
    ["definition", "定义", true, "definition", "quote"],
    ["axiom", "公理", true, "axiom", "quote"],
    ["assumption", "假设", true, "assumption", "quote"],
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
  ];
  for (const [environment, label, numbered, counterName, bodyKind] of builtInTheorems) {
    const definition = { label, numbered, counterName, bodyKind };
    theoremDefinitions.set(environment, definition);
    theoremDefinitions.set(`${environment}*`, { ...definition, numbered: false });
  }

  body = body.replace(
    /\\newtheorem(\*)?\s*\{([^{}]+)\}\s*(?:\[([^\]]+)\]\s*)?\{([^{}]*)\}\s*(?:\[([^\]]+)\])?/gi,
    (
      _match,
      starred: string | undefined,
      rawEnvironment: string,
      sharedCounter: string | undefined,
      rawLabel: string,
    ) => {
      const environment = rawEnvironment.trim();
      if (!environment) return "";
      const existing = theoremDefinitions.get(environment);
      theoremDefinitions.set(environment, {
        label: rawLabel.trim() || environment,
        numbered: !starred,
        counterName: sharedCounter?.trim() || environment,
        bodyKind: existing?.bodyKind ?? "quote",
      });
      return "";
    },
  );

  body = body.replace(
    /\\begin\{([^{}]+)\}\s*(?:\[([^\]]*)\])?/g,
    (match, rawEnvironment: string, note = "") => {
      const definition = theoremDefinitions.get(rawEnvironment.trim());
      if (!definition) return match;
      return `\n${encodeTheoremMarker({
        ...definition,
        note: note.trim(),
      })}\n`;
    },
  );
  body = body.replace(
    /\\end\{([^{}]+)\}/g,
    (match, rawEnvironment: string) =>
      theoremDefinitions.has(rawEnvironment.trim())
        ? `\n${theoremEndMarker}\n`
        : match,
  );
  body = body
    .replace(/\\qedhere\b/g, " □")
    .replace(/\\qed\b/g, " □");

  body = replaceLatexTableEnvironment(body, warnings);

  const preservedEnvironments = new Set([
    "document",
    "equation",
    "equation*",
    "align",
    "align*",
    "gather",
    "gather*",
    "multline",
    "multline*",
    "displaymath",
    "itemize",
    "enumerate",
    "quote",
    "quotation",
    "verbatim",
    "lstlisting",
    "aligned",
    "gathered",
    "cases",
    "matrix",
    "pmatrix",
    "bmatrix",
    "Bmatrix",
    "vmatrix",
    "Vmatrix",
    "array",
    "split",
  ]);
  const unknown = new Set<string>();
  body = body.replace(/\\(begin|end)\{([A-Za-z@*]+)\}(?:\[[^\]]*\])?/g, (match, _edge: string, name: string) => {
    if (preservedEnvironments.has(name)) return match;
    unknown.add(name);
    return "";
  });
  if (unknown.size) {
    warnings.push(
      `以下 LaTeX 环境没有对应的 Word 原生结构，已保留其中可见内容：${[...unknown].join("、")}。`,
    );
  }
  return body;
}

function findLatexInlineLiteralEnd(text: string, index: number) {
  if (text[index] !== "\\" || isEscaped(text, index)) return -1;
  const match = text
    .slice(index)
    .match(/^\\(?:verb|lstinline)\*?(?![A-Za-z@])(?:\[[^\]\n]*\])?/i);
  if (!match) return -1;
  const delimiterIndex = index + match[0].length;
  const delimiter = text[delimiterIndex];
  if (!delimiter || /[A-Za-z0-9\s]/.test(delimiter)) return -1;
  const close = text.indexOf(delimiter, delimiterIndex + 1);
  if (close >= 0) return close + 1;
  const lineEnd = text.indexOf("\n", delimiterIndex + 1);
  return lineEnd >= 0 ? lineEnd : text.length;
}

function findLatexDocumentToken(source: string, token: string, startIndex = 0) {
  const lowered = source.toLowerCase();
  const target = token.toLowerCase();
  for (let index = Math.max(0, startIndex); index < source.length; ) {
    const literalEnd = findLatexInlineLiteralEnd(source, index);
    if (literalEnd > index) {
      index = literalEnd;
      continue;
    }
    if (source[index] === "%" && !isEscaped(source, index)) {
      const lineEnd = source.indexOf("\n", index + 1);
      index = lineEnd >= 0 ? lineEnd + 1 : source.length;
      continue;
    }
    if (source[index] === "\\" && !isEscaped(source, index)) {
      const literalEnvironment = source
        .slice(index)
        .match(/^\\begin\{(verbatim\*?|lstlisting\*?)\}(?:\[[^\]\n]*\])?/i);
      if (literalEnvironment) {
        const endToken = `\\end{${literalEnvironment[1]}}`;
        const environmentEnd = lowered.indexOf(
          endToken.toLowerCase(),
          index + literalEnvironment[0].length,
        );
        index = environmentEnd >= 0 ? environmentEnd + endToken.length : source.length;
        continue;
      }
      if (lowered.startsWith(target, index)) return index;
    }
    index += 1;
  }
  return -1;
}

function findLatexCommentStart(line: string) {
  for (let index = 0; index < line.length; index += 1) {
    const literalEnd = findLatexInlineLiteralEnd(line, index);
    if (literalEnd > index) {
      index = literalEnd - 1;
      continue;
    }
    if (line[index] === "%" && !isEscaped(line, index)) return index;
  }
  return -1;
}

function normalizeLatexSource(source: string, warnings: string[]) {
  const normalizedSource = source.replace(/\r\n?/g, "\n");
  const beginToken = "\\begin{document}";
  const endToken = "\\end{document}";
  const begin = findLatexDocumentToken(normalizedSource, beginToken);
  const end = begin >= 0
    ? findLatexDocumentToken(normalizedSource, endToken, begin + beginToken.length)
    : -1;
  const beginSentinel = "\uE100VISUALTEX_LATEX_DOCUMENT_BEGIN\uE101";
  const endSentinel = "\uE102VISUALTEX_LATEX_DOCUMENT_END\uE103";
  const markedSource = begin < 0
    ? normalizedSource
    : end >= 0
      ? `${normalizedSource.slice(0, begin)}${beginSentinel}${normalizedSource.slice(begin + beginToken.length, end)}${endSentinel}${normalizedSource.slice(end + endToken.length)}`
      : `${normalizedSource.slice(0, begin)}${beginSentinel}${normalizedSource.slice(begin + beginToken.length)}`;

  let body = normalizeLatexExtensions(markedSource, warnings);
  if (begin >= 0) {
    const markedBegin = body.indexOf(beginSentinel);
    const contentStart = markedBegin >= 0 ? markedBegin + beginSentinel.length : 0;
    const markedEnd = body.indexOf(endSentinel, contentStart);
    body = markedEnd >= 0 ? body.slice(contentStart, markedEnd) : body.slice(contentStart);
    if (end < 0) warnings.push("LaTeX 文档缺少 \\end{document}，预览已读取其余内容。");
  }

  const result: string[] = [];
  let literalEnd = "";
  for (const line of body.split("\n")) {
    const trimmed = line.trim();
    if (literalEnd) {
      result.push(line);
      if (trimmed.toLowerCase() === literalEnd.toLowerCase()) literalEnd = "";
      continue;
    }
    const literal = trimmed.match(/^\\begin\{(verbatim|lstlisting)\}(?:\[[^\]]*\])?\s*$/i);
    if (literal) {
      literalEnd = `\\end{${literal[1]}}`;
      result.push(line);
      continue;
    }
    const comment = findLatexCommentStart(line);
    result.push(comment >= 0 ? line.slice(0, comment) : line);
  }
  return result.join("\n").trim();
}

function listLevel(indentation: string) {
  const columns = [...indentation].reduce((total, character) => total + (character === "\t" ? 4 : 1), 0);
  return Math.min(8, Math.max(0, Math.floor(columns / 2)));
}

export function parseDocumentImport(
  source: string,
  requestedFormat: DocumentSourceFormat,
): ParsedDocumentImport {
  if (!source.trim()) throw new Error("请输入需要导入的 LaTeX 或 Markdown 内容。");
  if (source.length > 5_000_000) throw new Error("批量导入内容不能超过 5 MB。");
  const warnings: string[] = [];
  const format = requestedFormat === "auto" ? detectFormat(source) : requestedFormat;
  const normalized = (format === "latex"
    ? normalizeLatexSource(source, warnings)
    : normalizeMarkdownSource(source, warnings))
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .trim();
  const blocks: DocumentImportBlock[] = [];
  const paragraph: string[] = [];
  const quote: string[] = [];
  const listModes: Array<"bullet" | "numbered"> = [];
  const theoremBodyKinds: TheoremBodyKind[] = [];
  const theoremCounters = new Map<string, number>();
  let latexQuoteDepth = 0;
  let inCode = false;
  let codeEnd = "";
  let codeDescription = "";
  const code: string[] = [];

  const flushParagraph = () => {
    if (!paragraph.length) return;
    appendMixedBlocks(blocks, paragraph.join("\n"), format, warnings);
    paragraph.length = 0;
  };
  const flushQuote = () => {
    if (!quote.length) return;
    appendMixedBlocks(blocks, quote.join("\n"), format, warnings, "quote");
    quote.length = 0;
  };
  const finishCode = (warning?: string) => {
    blocks.push({
      id: id(),
      kind: "code",
      level: 0,
      runs: [{ kind: "text", text: code.join("\n").replace(/[\r\n]+$/, ""), code: true }],
    });
    code.length = 0;
    inCode = false;
    codeEnd = "";
    codeDescription = "";
    if (warning) warnings.push(warning);
  };

  for (const raw of normalized.split("\n")) {
    const trimmed = raw.trim();
    if (format === "markdown" && trimmed.startsWith("```")) {
      flushParagraph();
      flushQuote();
      if (inCode && codeEnd === "```") finishCode();
      else if (!inCode) {
        inCode = true;
        codeEnd = "```";
        codeDescription = "Markdown 代码块";
      } else code.push(raw);
      continue;
    }
    if (format === "latex" && !inCode) {
      const start = trimmed.match(/^\\begin\{(verbatim|lstlisting)\}(?:\[[^\]]*\])?\s*$/i);
      if (start) {
        flushParagraph();
        flushQuote();
        inCode = true;
        codeEnd = `\\end{${start[1]}}`;
        codeDescription = `LaTeX ${start[1]} 环境`;
        continue;
      }
    }
    if (inCode) {
      if (trimmed.toLowerCase() === codeEnd.toLowerCase()) finishCode();
      else code.push(raw);
      continue;
    }

    if (format === "latex") {
      const theoremStart = decodeTheoremMarker(trimmed);
      if (theoremStart) {
        flushParagraph();
        flushQuote();
        let theoremTitle = theoremStart.label;
        if (theoremStart.numbered) {
          const nextNumber = (theoremCounters.get(theoremStart.counterName) ?? 0) + 1;
          theoremCounters.set(theoremStart.counterName, nextNumber);
          theoremTitle += ` ${nextNumber}`;
        }
        if (theoremStart.note) theoremTitle += `（${theoremStart.note}）`;
        blocks.push({
          id: id(),
          kind: "heading",
          level: 4,
          runs: parseInline(theoremTitle, format),
        });
        theoremBodyKinds.push(theoremStart.bodyKind);
        continue;
      }
      if (trimmed === theoremEndMarker) {
        flushParagraph();
        flushQuote();
        if (theoremBodyKinds.length) theoremBodyKinds.pop();
        else warnings.push("忽略了没有对应开始标记的 LaTeX 定理环境结束标记。");
        continue;
      }
      if (/^\\begin\{(?:quote|quotation)\}\s*$/i.test(trimmed)) {
        flushParagraph();
        flushQuote();
        latexQuoteDepth += 1;
        continue;
      }
      if (/^\\end\{(?:quote|quotation)\}\s*$/i.test(trimmed)) {
        flushParagraph();
        flushQuote();
        latexQuoteDepth = Math.max(0, latexQuoteDepth - 1);
        continue;
      }
      const listStart = trimmed.match(/^\\begin\{(itemize|enumerate)\}\s*$/i);
      if (listStart) {
        flushParagraph();
        flushQuote();
        listModes.push(listStart[1].toLowerCase() === "enumerate" ? "numbered" : "bullet");
        continue;
      }
      if (/^\\end\{(?:itemize|enumerate)\}\s*$/i.test(trimmed)) {
        flushParagraph();
        flushQuote();
        if (listModes.length) listModes.pop();
        else warnings.push(`忽略了没有对应开始标记的 ${trimmed}。`);
        continue;
      }
      const heading = trimmed.match(/^\\(part|chapter|section|subsection|subsubsection|paragraph|subparagraph)\*?\{(.*)\}\s*$/i);
      if (heading) {
        flushParagraph();
        flushQuote();
        const levels: Record<string, number> = {
          part: 1,
          chapter: 1,
          section: 1,
          subsection: 2,
          subsubsection: 3,
          paragraph: 4,
          subparagraph: 5,
        };
        blocks.push({
          id: id(),
          kind: "heading",
          level: levels[heading[1].toLowerCase()] ?? 4,
          runs: parseInline(heading[2], format),
        });
        continue;
      }
      const item = trimmed.match(/^\\item(?:\s*\[[^\]]*\])?\s*(.*)$/);
      if (item) {
        flushParagraph();
        flushQuote();
        appendMixedBlocks(
          blocks,
          item[1],
          format,
          warnings,
          listModes.at(-1) === "numbered" ? "numbered" : "bullet",
          Math.max(0, listModes.length - 1),
        );
        continue;
      }
    } else {
      if (trimmed.startsWith(">")) {
        flushParagraph();
        quote.push(trimmed.replace(/^>+\s?/, ""));
        continue;
      }
      flushQuote();
      const heading = raw.match(/^(#{1,6})\s+(.+?)\s*#*\s*$/);
      if (heading) {
        flushParagraph();
        blocks.push({ id: id(), kind: "heading", level: heading[1].length, runs: parseInline(heading[2], format) });
        continue;
      }
      const bullet = raw.match(/^(\s*)[-+*]\s+(.+)$/);
      if (bullet) {
        flushParagraph();
        appendMixedBlocks(blocks, bullet[2], format, warnings, "bullet", listLevel(bullet[1]));
        continue;
      }
      const numbered = raw.match(/^(\s*)\d+[.)]\s+(.+)$/);
      if (numbered) {
        flushParagraph();
        appendMixedBlocks(blocks, numbered[2], format, warnings, "numbered", listLevel(numbered[1]));
        continue;
      }
    }

    if (!trimmed) {
      flushParagraph();
      flushQuote();
      continue;
    }
    const theoremBodyKind = theoremBodyKinds.at(-1);
    if (latexQuoteDepth > 0 || theoremBodyKind === "quote") quote.push(trimmed);
    else paragraph.push(trimmed);
  }

  if (inCode) finishCode(`${codeDescription}未闭合，预览已读取到文末。`);
  flushParagraph();
  flushQuote();
  if (latexQuoteDepth > 0) warnings.push("LaTeX quote/quotation 环境未闭合，预览已读取到文末。");
  if (theoremBodyKinds.length) {
    warnings.push(`LaTeX 文档有 ${theoremBodyKinds.length} 个定理/证明环境未闭合。`);
  }
  if (listModes.length) warnings.push(`LaTeX 文档有 ${listModes.length} 个列表环境未闭合。`);
  if (!blocks.length) throw new Error("没有找到可以插入 Word 的文字或公式。");

  const runs = blocks.flatMap((block) => block.runs);
  const formulaCount = runs.filter((run) => run.kind === "formula").length;
  const displayFormulaCount = runs.filter((run) => run.kind === "formula" && run.display).length;
  const textCharacterCount = runs.reduce(
    (total, run) => total + (run.kind === "text" ? run.text.length : 0),
    0,
  );
  return {
    format,
    blocks,
    warnings,
    formulaCount,
    inlineFormulaCount: formulaCount - displayFormulaCount,
    displayFormulaCount,
    textCharacterCount,
  };
}
