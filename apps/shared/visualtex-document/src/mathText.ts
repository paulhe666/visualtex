// Text with LaTeX math delimiters (pasted LLM/OCR output, Markdown, LaTeX
// source) -> paragraphs with inline formulas and display formulas.
// Inline: \( \) and $ $ with Pandoc's rule (the opening $ has a non-space
// right after it, the closing $ a non-space right before it and no digit
// right after it), so prices like "$5 and $10" stay text.
// Display: \[ \], $$ $$ and the amsmath display environments.

export type MathTextInline = { kind: "text"; text: string } | { kind: "math"; latex: string };
export type MathTextBlock =
  | { kind: "paragraph"; parts: MathTextInline[] }
  | { kind: "display"; latex: string; numbered: boolean };

const DISPLAY_ENVIRONMENTS: Record<string, { inner: string | null; numbered: boolean }> = {
  equation: { inner: null, numbered: true },
  "equation*": { inner: null, numbered: false },
  displaymath: { inner: null, numbered: false },
  align: { inner: "aligned", numbered: true },
  "align*": { inner: "aligned", numbered: false },
  flalign: { inner: "aligned", numbered: true },
  "flalign*": { inner: "aligned", numbered: false },
  eqnarray: { inner: "aligned", numbered: true },
  "eqnarray*": { inner: "aligned", numbered: false },
  gather: { inner: "gathered", numbered: true },
  "gather*": { inner: "gathered", numbered: false },
  multline: { inner: "gathered", numbered: true },
  "multline*": { inner: "gathered", numbered: false },
};

/** Whether the text holds any math delimiter worth converting. */
export function hasMathDelimiters(text: string): boolean {
  return /\\\(|\\\[|\$|\\begin\{/.test(text);
}

function escapedAt(text: string, index: number): boolean {
  let backslashes = 0;
  for (let i = index - 1; i >= 0 && text[i] === "\\"; i--) backslashes++;
  return backslashes % 2 === 1;
}

/** \label{...} and \nonumber/\notag are the document's business, not the formula's. */
function cleanDisplayLatex(latex: string): string {
  return latex
    .replace(/\\label\s*\{[^{}]*\}/g, "")
    .replace(/\\(?:nonumber|notag)\b/g, "")
    .trim();
}

function displayFromEnvironment(name: string, body: string): { latex: string; numbered: boolean } {
  const env = DISPLAY_ENVIRONMENTS[name];
  const cleaned = cleanDisplayLatex(body);
  if (!env || !env.inner) return { latex: cleaned, numbered: env?.numbered ?? false };
  const normalized = name.startsWith("eqnarray") ? cleaned.replace(/&\s*([=<>]|\\le|\\ge)\s*&/g, "&$1 ") : cleaned;
  return { latex: `\\begin{${env.inner}}${normalized}\\end{${env.inner}}`, numbered: env.numbered };
}

/** Split paragraph text into text and inline formulas. */
export function splitInlineMath(text: string): MathTextInline[] {
  const parts: MathTextInline[] = [];
  let buffer = "";
  const flush = () => {
    if (buffer) parts.push({ kind: "text", text: buffer });
    buffer = "";
  };
  let i = 0;
  while (i < text.length) {
    const ch = text[i];
    if (ch === "\\" && text[i + 1] === "(" && !escapedAt(text, i)) {
      const close = text.indexOf("\\)", i + 2);
      if (close > 0) {
        flush();
        parts.push({ kind: "math", latex: text.slice(i + 2, close).trim() });
        i = close + 2;
        continue;
      }
    }
    if (ch === "\\" && text[i + 1] === "$") {
      buffer += "$";
      i += 2;
      continue;
    }
    if (ch === "$" && text[i + 1] !== "$" && text[i - 1] !== "$" && !escapedAt(text, i)) {
      const open = i;
      const first = text[open + 1];
      if (first && !/\s/.test(first)) {
        let close = -1;
        for (let j = open + 1; j < text.length; j++) {
          if (text[j] === "\n" && text[j + 1] === "\n") break;
          if (text[j] === "$" && !escapedAt(text, j)) {
            if (text[j + 1] === "$") break;
            if (!/\s/.test(text[j - 1]) && !/[0-9]/.test(text[j + 1] ?? "")) close = j;
            break;
          }
        }
        if (close > open + 1) {
          flush();
          parts.push({ kind: "math", latex: text.slice(open + 1, close).trim() });
          i = close + 1;
          continue;
        }
      }
    }
    buffer += ch;
    i++;
  }
  flush();
  return parts;
}

/** Text -> blocks. Blank lines separate paragraphs; single line breaks inside a paragraph become spaces. */
export function parseMathText(text: string): MathTextBlock[] {
  const source = text.replace(/\r\n?/g, "\n");
  const blocks: MathTextBlock[] = [];
  let paragraph = "";
  const flushParagraph = () => {
    const chunks = paragraph.split(/\n\s*\n/);
    for (const chunk of chunks) {
      const line = chunk.replace(/\s*\n\s*/g, " ").trim();
      if (line) blocks.push({ kind: "paragraph", parts: splitInlineMath(line) });
    }
    paragraph = "";
  };
  let i = 0;
  while (i < source.length) {
    if (!escapedAt(source, i)) {
      let match: { end: number; latex: string; numbered: boolean } | null = null;
      if (source.startsWith("\\[", i)) {
        const close = source.indexOf("\\]", i + 2);
        if (close > 0) match = { end: close + 2, latex: cleanDisplayLatex(source.slice(i + 2, close)), numbered: false };
      } else if (source.startsWith("$$", i)) {
        const close = source.indexOf("$$", i + 2);
        if (close > 0) match = { end: close + 2, latex: cleanDisplayLatex(source.slice(i + 2, close)), numbered: false };
      } else if (source.startsWith("\\begin{", i)) {
        const name = /^\\begin\{([a-zA-Z*]+)\}/.exec(source.slice(i))?.[1];
        if (name && DISPLAY_ENVIRONMENTS[name]) {
          const openEnd = i + "\\begin{".length + name.length + 1;
          const closeTag = `\\end{${name}}`;
          const close = source.indexOf(closeTag, openEnd);
          if (close > 0) {
            const display = displayFromEnvironment(name, source.slice(openEnd, close));
            match = { end: close + closeTag.length, ...display };
          }
        }
      }
      if (match) {
        flushParagraph();
        if (match.latex) blocks.push({ kind: "display", latex: match.latex, numbered: match.numbered });
        i = match.end;
        continue;
      }
    }
    paragraph += source[i];
    i++;
  }
  flushParagraph();
  return blocks;
}
