// The document edited as its source (the app's source tab): LaTeX or Markdown read back into the document as one
// change — only the top-level blocks that differ are replaced, so the others keep their ids, their formulas and
// the caret; formulas' labels and reference list keys are matched to the ids they had, so that references in the
// blocks kept still find them. A LaTeX preamble's page settings (what settingsPreamble writes) are read too.
// Code that is not finished (an open brace, environment or formula) is reported, not read; what reading would drop
// (commands and environments the document has no place for, comments, preamble lines) is reported before it goes.
import { Fragment, type Node, type Schema } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import { closeHistory } from "prosemirror-history";
import { latexToMarkdown } from "./latexImport.ts";
import { markdownSourceKey, readMarkdown, rebaseMarkdownSource, type MarkdownSource } from "./markdown.ts";
import { withoutMarkdownSyntax } from "./markdownLive.ts";
import { closeFormula, openFormulaId } from "./formulas.ts";
import { DEFAULT_DOCUMENT_SETTINGS, PAPER_SIZES, SINGLE_LINE_HEIGHT, CHINESE_FONT_SIZES, documentSettings, type DocumentSettings, type PaperName } from "./settings.ts";
import { imageFiles, setImageFileResolver } from "./images.ts";
import { separateParagraphs, sourceTextToDocument, toSourceDocument } from "./sourceMarkdown.ts";

export type SourceFormat = "latex" | "markdown";

/** Unfinished code: where (offsets in the source, its 1-based line) and what. */
export interface SourceProblem {
  from: number;
  to: number;
  line: number;
  kind: "brace" | "extra-brace" | "environment" | "end" | "math" | "fence";
  /** the environment's name, the delimiter */
  detail: string;
}

/** Something reading the source drops: a command, an environment, a comment, a preamble line; the lines it is on. */
export interface SourceLoss {
  kind: "command" | "environment" | "comment" | "preamble" | "line";
  /** \vspace, figure*, the line's text */
  name: string;
  lines: number[];
}

const lineAt = (text: string, offset: number) => {
  let line = 1;
  for (let i = 0; i < offset && i < text.length; i++) if (text[i] === "\n") line++;
  return line;
};

const VERBATIM = /^(lstlisting|verbatim\*?|minted|comment|Verbatim)$/;

/** The first unfinished piece of LaTeX: an unclosed brace, environment or formula; an \end without its \begin. */
export function latexProblem(text: string): SourceProblem | null {
  const braces: number[] = [];
  const environments: Array<{ name: string; at: number; end: number }> = [];
  let math: { at: number; close: string } | null = null;
  const problem = (kind: SourceProblem["kind"], from: number, to: number, detail: string): SourceProblem => ({ kind, from, to, line: lineAt(text, from), detail });
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === "%") {
      while (i < text.length && text[i] !== "\n") i++;
      continue;
    }
    if (c === "\\") {
      const rest = text.slice(i, i + 200);
      const env = /^\\(begin|end)\s*\{([^{}\n]*)\}/.exec(rest);
      if (env) {
        const name = env[2].trim();
        if (env[1] === "begin") {
          if (VERBATIM.test(name)) {
            // its text is taken as it is, to its \end
            const close = text.indexOf(`\\end{${name}}`, i + env[0].length);
            if (close < 0) return problem("environment", i, i + env[0].length, name);
            i = close + `\\end{${name}}`.length - 1;
            continue;
          }
          environments.push({ name, at: i, end: i + env[0].length });
        } else {
          const open = environments.pop();
          if (!open) return problem("end", i, i + env[0].length, name);
          if (open.name !== name) return problem("environment", open.at, open.end, open.name);
        }
        i += env[0].length - 1;
        continue;
      }
      const verb = /^\\verb\*?([^A-Za-z\s*])/.exec(rest);
      if (verb) {
        const close = text.indexOf(verb[1], i + verb[0].length);
        if (close < 0 || text.slice(i, close).includes("\n")) return problem("brace", i, i + verb[0].length, "\\verb");
        i = close;
        continue;
      }
      const next = text[i + 1];
      if (next === "[" || next === "(") {
        if (!math) math = { at: i, close: next === "[" ? "\\]" : "\\)" };
        i++;
        continue;
      }
      if (next === "]" || next === ")") {
        if (math && math.close === `\\${next}`) math = null;
        i++;
        continue;
      }
      i++;
      continue;
    }
    if (c === "$") {
      const display = text[i + 1] === "$";
      const delimiter = display ? "$$" : "$";
      if (math && math.close === delimiter) math = null;
      else if (!math) math = { at: i, close: delimiter };
      if (display) i++;
      continue;
    }
    if (c === "{") braces.push(i);
    else if (c === "}") {
      if (!braces.length) return problem("extra-brace", i, i + 1, "}");
      braces.pop();
    }
  }
  if (math) {
    const opener = math.close === "\\]" ? "\\[" : math.close === "\\)" ? "\\(" : math.close;
    return problem("math", math.at, math.at + opener.length, opener);
  }
  if (environments.length) {
    const open = environments[environments.length - 1];
    return problem("environment", open.at, open.end, open.name);
  }
  if (braces.length) return problem("brace", braces[braces.length - 1], braces[braces.length - 1] + 1, "{");
  return null;
}

/** The first unfinished piece of Markdown: an open code fence or $$ block (they would take in all that follows). */
export function markdownProblem(text: string): SourceProblem | null {
  const lines = text.split("\n");
  let fence: { mark: string; at: number; line: number; length: number } | null = null;
  let math: { at: number; line: number } | null = null;
  let offset = 0;
  for (let n = 0; n < lines.length; n++) {
    const line = lines[n].replace(/^(?:\s*>)*\s*(?:(?:[-*+]|\d+[.)])\s+)?/, "");
    const lead = lines[n].length - line.length;
    const marker = /^(`{3,}|~{3,})/.exec(line);
    if (fence) {
      if (marker && marker[1][0] === fence.mark[0] && marker[1].length >= fence.mark.length && !line.slice(marker[1].length).trim()) fence = null;
    } else if (marker) fence = { mark: marker[1], at: offset + lead, line: n + 1, length: marker[1].length };
    else if (/^\$\$/.test(line)) {
      // $$ on its own line opens (or closes) a block; $$…$$ on one line is a whole formula
      const rest = line.slice(2);
      const closesHere = /\$\$\s*$/.test(rest);
      if (math) {
        math = null;
      } else if (!closesHere) math = { at: offset + lead, line: n + 1 };
    } else if (math && /\$\$\s*$/.test(line)) math = null;
    offset += lines[n].length + 1;
  }
  if (fence) return { kind: "fence", from: fence.at, to: fence.at + fence.length, line: fence.line, detail: fence.mark };
  if (math) return { kind: "math", from: math.at, to: math.at + 2, line: math.line, detail: "$$" };
  return null;
}

export const sourceProblem = (text: string, format: SourceFormat): SourceProblem | null =>
  format === "latex" ? latexProblem(text) : markdownProblem(text);

// ------------------------------------------------------------ settings --

type PreambleSettings = Partial<DocumentSettings>;

const MM: Record<string, number> = { mm: 1, cm: 10, in: 25.4, pt: 25.4 / 72.27, bp: 25.4 / 72 };

function length(value: string | undefined): number | null {
  const match = /^\s*([\d.]+)\s*(mm|cm|in|pt|bp)\s*$/.exec(value ?? "");
  return match ? Math.round(Number(match[1]) * MM[match[2]] * 10) / 10 : null;
}

/** The page settings a preamble sets in the forms settingsPreamble writes (and a few common others). */
export function preambleSettings(preamble: string): PreambleSettings {
  const text = preamble.replace(/(^|[^\\])%.*$/gm, "$1");
  const out: PreambleSettings = {};
  const documentClass = /\\documentclass\s*(?:\[([^\]]*)\])?\s*\{([^{}]*)\}/.exec(text);
  const classOptions = (documentClass?.[1] ?? "").split(",").map((option) => option.trim());
  for (const option of classOptions) {
    const points = /^(\d+(?:\.\d+)?)pt$/.exec(option);
    if (points) out.fontSize = Number(points[1]);
    const zihao = /^zihao\s*=\s*(-?\d)$/.exec(option);
    const named = zihao && CHINESE_FONT_SIZES.find((size) => size.zihao === zihao[1]);
    if (named) out.fontSize = named.pt;
    const spread = /^linespread\s*=\s*([\d.]+)$/.exec(option);
    if (spread) out.lineSpacing = Math.round((Number(spread[1]) / SINGLE_LINE_HEIGHT) * 100) / 100;
  }
  const scrextend = /\\usepackage\s*\[\s*fontsize\s*=\s*([\d.]+)pt\s*\]\s*\{scrextend\}/.exec(text);
  if (scrextend) out.fontSize = Number(scrextend[1]);
  const linespread = /\\linespread\s*\{([\d.]+)\}/.exec(text);
  if (linespread) out.lineSpacing = Math.round(((Number(linespread[1]) * 1.2) / SINGLE_LINE_HEIGHT) * 100) / 100;
  // the page: \usepackage[…]{geometry} or \geometry{…}
  const geometry = /\\usepackage\s*\[([^\]]*)\]\s*\{geometry\}/.exec(text)?.[1] ?? /\\geometry\s*\{([^{}]*)\}/.exec(text)?.[1];
  if (geometry !== undefined) {
    const options = new Map<string, string>();
    for (const part of geometry.split(",")) {
      const [key, value] = part.split("=").map((s) => s.trim());
      if (key) options.set(key, value ?? "");
    }
    const width = length(options.get("paperwidth"));
    const height = length(options.get("paperheight"));
    const named = ([...options.keys()].map((key) => /^(a4|a5|b5|letter)paper$/.exec(key)?.[1]).find(Boolean) ?? null) as PaperName | null;
    const paper = named ?? (Object.keys(PAPER_SIZES) as PaperName[]).find((name) => width !== null && height !== null && Math.abs(PAPER_SIZES[name].width - width) < 0.6 && Math.abs(PAPER_SIZES[name].height - height) < 0.6);
    if (paper) out.paper = paper;
    const margin = length(options.get("margin"));
    const side = (key: string, ...also: string[]) => length(options.get(key)) ?? also.map((k) => length(options.get(k))).find((v) => v !== null) ?? margin;
    const top = side("top", "vmargin");
    const bottom = side("bottom", "vmargin");
    const left = side("left", "hmargin", "inner");
    const right = side("right", "hmargin", "outer");
    if (top !== null) out.marginTop = top;
    if (bottom !== null) out.marginBottom = bottom;
    if (left !== null) out.marginLeft = left;
    if (right !== null) out.marginRight = right;
  }
  if (/\\numberwithin\s*\{equation\}\s*\{section\}/.test(text)) out.equationNumbering = /\\renewcommand\s*\{?\\theequation\}?\s*\{\\thesection-/.test(text) ? "1-1" : "1.1";
  if (/\\tagform@\}\s*\[1\]\s*\{\\maketag@@@\{\[/.test(text)) out.equationBrackets = "[]";
  if (/\\setcounter\s*\{secnumdepth\}\s*\{\s*0\s*\}/.test(text)) out.headingNumbering = "none";
  const latin = /\\setmainfont\s*(?:\[[^\]]*\])?\s*\{([^{}]+)\}/.exec(text);
  if (latin) out.latinFont = latin[1].trim();
  const cjkFont = /\\setCJKmainfont\s*(?:\[[^\]]*\])?\s*\{([^{}]+)\}/.exec(text);
  if (cjkFont) out.cjkFont = cjkFont[1].trim();
  return out;
}

/** The preamble of a whole LaTeX document (before \begin{document}), or "" for a body alone. */
export function latexPreamble(text: string): string {
  const at = text.search(/\\begin\s*\{document\}/);
  return at < 0 ? "" : text.slice(0, at);
}

/**
 * The settings the user changed in the preamble: each setting the user's preamble gives differently from the one
 * written for the document now (`written`); one the user took out goes back to its default (null).
 */
export function preambleSettingsChange(user: string, written: string, current: DocumentSettings): Partial<Record<keyof DocumentSettings, DocumentSettings[keyof DocumentSettings] | null>> {
  if (!user.trim()) return {};
  const mine = preambleSettings(user);
  const theirs = preambleSettings(written);
  const change: Partial<Record<keyof DocumentSettings, DocumentSettings[keyof DocumentSettings] | null>> = {};
  for (const key of Object.keys(DEFAULT_DOCUMENT_SETTINGS) as Array<keyof DocumentSettings>) {
    const a = mine[key];
    const b = theirs[key];
    const same = typeof a === "number" && typeof b === "number" ? Math.abs(a - b) < 0.051 : a === b;
    if (same) continue;
    // a heading numbering written "1." or "1" is not in the preamble: only "none" is
    if (key === "headingNumbering" && a === undefined && current.headingNumbering !== "none") continue;
    change[key] = a ?? null;
  }
  return change;
}

// ---------------------------------------------------------------- reading --

/**
 * The source as a document (and, for Markdown, what keeps its unchanged blocks as they were written). The images'
 * paths the exports write (figures/fig-1.png, assets/fig-1.png: imageFiles) are the document's images again
 * (`images`: the document whose images they are).
 */
export function readSource(schema: Schema, text: string, format: SourceFormat, images: Node | null = null): { doc: Node; source: MarkdownSource | null } {
  const paths = new Map<string, string>();
  // (LaTeX writes them in figures/, Markdown in assets/)
  if (images) for (const folder of ["figures", "assets"]) for (const [key, path] of imageFiles(images, folder).paths) paths.set(path, key);
  if (paths.size) setImageFileResolver((path) => paths.get(path) ?? null);
  try {
    // (the source editor's document is the Markdown's lines)
    if (format === "markdown") return { doc: sourceTextToDocument(schema, text), source: null };
    return { doc: toSourceDocument(separateParagraphs(readMarkdown(schema, latexToMarkdown(text)).doc)), source: null };
  } finally {
    if (paths.size) setImageFileResolver(null);
  }
}

/**
 * Ids in a document read from the source made the ones the document has: a formula's label (eq:<id>, or the label
 * it was read with) gives back its id, a reference list entry's key its entry's id; references follow.
 */
export function matchIds(read: Node, current: Node): Node {
  const formulas = new Map<string, string>();
  current.descendants((node) => {
    if (node.type.name === "math_display" && node.attrs.id) {
      const id = node.attrs.id as string;
      formulas.set(id, id);
      if (!id.startsWith("eq:")) formulas.set(`eq:${id}`, id);
    }
    return true;
  });
  const keys = new Map<string, string>();
  current.descendants((node) => {
    if (node.type.name === "bib_entry" && node.attrs.id && node.attrs.key) keys.set(node.attrs.key, node.attrs.id);
    return true;
  });
  const ids = new Map<string, string>();
  const used = new Set<string>();
  read.descendants((node) => {
    const id = node.attrs?.id as string | null | undefined;
    if (!id) return true;
    const old = node.type.name === "math_display" ? formulas.get(id) : node.type.name === "bib_entry" && node.attrs.key ? keys.get(node.attrs.key) : undefined;
    if (old && !used.has(old)) {
      used.add(old);
      if (old !== id) ids.set(id, old);
    }
    return true;
  });
  if (!ids.size) return read;
  const json = read.toJSON() as JsonNode;
  const walk = (node: JsonNode) => {
    const attrs = node.attrs;
    if (attrs) {
      if ((node.type === "math_display" || node.type === "bib_entry") && typeof attrs.id === "string" && ids.has(attrs.id)) attrs.id = ids.get(attrs.id);
      if (node.type === "eqref" && typeof attrs.target === "string" && ids.has(attrs.target)) attrs.target = ids.get(attrs.target);
      if (node.type === "cite" && Array.isArray(attrs.targets)) attrs.targets = (attrs.targets as string[]).map((t) => ids.get(t) ?? t);
    }
    node.content?.forEach(walk);
  };
  walk(json);
  return read.type.schema.nodeFromJSON(json);
}

interface JsonNode {
  type: string;
  attrs?: Record<string, unknown>;
  content?: JsonNode[];
}

/** A block as it compares: its content without ids. */
const blockKey = (node: Node) => JSON.stringify(node.toJSON(), (key, value) => (key === "id" ? undefined : value));

/**
 * The change that makes the document what the source says: the top-level blocks from the first that differs to the
 * last that differs replaced (one step in the history of its own); the preamble's settings; for Markdown, the file
 * kept for writing back. A block differs when it reads otherwise than the document's block and than what the
 * document's own source (`written`, what the source tab shows for it) reads as — a block the user did not touch
 * stays as it is even when its source does not read back exactly (an image whose file is elsewhere). Null: the
 * source says what the document is.
 */
export function sourceTransaction(state: EditorState, text: string, format: SourceFormat, written: string): Transaction | null {
  const { schema } = state;
  const current = state.doc;
  const read = readSource(schema, text, format, current);
  // (the document's text is its Markdown source: what was read, as source)
  let doc = toSourceDocument(matchIds(read.doc, current));
  if (!doc.childCount) doc = doc.type.create(doc.attrs, schema.nodes.paragraph.create());
  const source = read.source ? rebaseMarkdownSource(read.source, read.doc, doc) : null;
  const shown = withoutMarkdownSyntax(current);
  const compare = shown.childCount === current.childCount ? shown : current;
  const keys = (node: Node) => {
    const list: string[] = [];
    node.forEach((child) => list.push(blockKey(child)));
    return list;
  };
  const oldKeys = keys(compare);
  const newKeys = keys(doc);
  const base = written === text ? newKeys : keys(matchIds(readSource(schema, written, format, current).doc, current));
  // what the document's source reads as, block for block (when it reads as as many blocks)
  const baseKeys = base.length === oldKeys.length ? base : null;
  const same = (oldIndex: number, newIndex: number) => oldKeys[oldIndex] === newKeys[newIndex] || (!!baseKeys && baseKeys[oldIndex] === newKeys[newIndex]);
  let start = 0;
  while (start < oldKeys.length && start < newKeys.length && same(start, start)) start++;
  let end = 0;
  while (end < oldKeys.length - start && end < newKeys.length - start && same(oldKeys.length - 1 - end, newKeys.length - 1 - end)) end++;
  const tr = state.tr;
  if (start < oldKeys.length - end || start < newKeys.length - end) {
    let from = 0;
    for (let i = 0; i < start; i++) from += current.child(i).nodeSize;
    let to = from;
    for (let i = start; i < oldKeys.length - end; i++) to += current.child(i).nodeSize;
    const blocks: Node[] = [];
    for (let i = start; i < newKeys.length - end; i++) blocks.push(doc.child(i));
    tr.replaceWith(from, to, Fragment.from(blocks));
  }
  if (format === "latex") {
    const change = preambleSettingsChange(latexPreamble(text), latexPreamble(written), documentSettings(current));
    for (const [name, value] of Object.entries(change)) if (current.attrs[name] !== value) tr.setDocAttribute(name, value);
  }
  if (format === "markdown") tr.setMeta(markdownSourceKey, source);
  if (!tr.docChanged) {
    // the file kept for writing back may still change (its spacing): not a step in the history
    return format === "markdown" ? tr.setMeta("addToHistory", false) : null;
  }
  if (openFormulaId(state)) closeFormula(tr);
  return closeHistory(tr);
}

// ----------------------------------------------------------------- losses --

/** Commands the reader takes in as something the document writes otherwise (counted as that). */
const SAME_AS: Record<string, string> = {
  emph: "textit",
  textsl: "textit",
  uline: "underline",
  ref: "eqref",
  autoref: "eqref",
  cref: "eqref",
  Cref: "eqref",
  citep: "cite",
  citet: "cite",
  citealp: "cite",
  citealt: "cite",
  parencite: "cite",
  textcite: "cite",
  autocite: "cite",
  supercite: "cite",
  upcite: "cite",
  thanks: "footnote",
  newline: "\\",
  linebreak: "\\",
  "(": "$",
  ")": "$",
  "[": "$$",
  "]": "$$",
  bf: "textbf",
  bfseries: "textbf",
  it: "textit",
  itshape: "textit",
  em: "textit",
  tt: "texttt",
  ttfamily: "texttt",
};

/** Commands and environments the reader takes in (their meaning is kept even when the export writes it otherwise). */
const READ_COMMANDS = new Set(
  (
    "textbf textit underline texttt sout url href caption subcaption label eqref cite footnote textcolor colorbox textsuperscript " +
    "textsubscript chapter section subsection subsubsection paragraph maketitle title author date and today tableofcontents item " +
    "begin end centering raggedleft raggedright par ldots dots mbox textrm textnormal text bibitem newblock hline " +
    "toprule midrule bottomrule cline cmidrule multicolumn multirow newpage clearpage cleardoublepage pagebreak pagenumbering " +
    "pagestyle thispagestyle setcounter noindent hfill quad qquad enspace thinspace endfirsthead endhead endfoot endlastfoot " +
    "tabularnewline arraystretch appendix TeX LaTeX LaTeXe BibTeX XeTeX XeLaTeX LuaTeX LuaLaTeX pdfTeX pdfLaTeX AmS AmSTeX " +
    "copyright textcopyright textregistered texttrademark dag ddag textdagger textdaggerdbl S P pounds textsterling euro texteuro " +
    "textdegree textbullet textperiodcentered textbackslash textendash textemdash textellipsis textbar textless textgreater " +
    "textquotedblleft textquotedblright textquoteleft textquoteright guillemotleft guillemotright slash textasciicircum " +
    "textasciitilde ss ae AE oe OE o O aa AA l L i j c v u H r k resizebox scalebox small footnotesize scriptsize normalsize " +
    "graphicspath vtxwrapclear setlength unitlength put linewidth textwidth"
  ).split(" "),
);
const READ_ENVIRONMENTS = new Set(
  (
    "document itemize enumerate description quote quotation center flushleft flushright abstract figure figure* table table* " +
    "tabular tabular* tabularx tabulary longtable xtabular supertabular tabu lstlisting verbatim minted thebibliography " +
    "wrapfigure threeparttable tablenotes subfigure minipage picture adjustbox"
  ).split(" "),
);

const ACCENT_COMMANDS = new Set(["'", '"', "`", "^", "~", "=", ".", ",", "%", "&", "#", "_", "$", "{", "}", "\\", " "]);

interface Tokens {
  counts: Map<string, number>;
  lines: Map<string, number[]>;
}

/** The commands, environments and comments in a LaTeX body, outside formulas and code, with their lines. */
function bodyTokens(text: string, firstLine = 1): Tokens {
  const counts = new Map<string, number>();
  const lines = new Map<string, number[]>();
  const add = (key: string, line: number) => {
    counts.set(key, (counts.get(key) ?? 0) + 1);
    const list = lines.get(key) ?? [];
    if (!list.includes(line)) list.push(line);
    lines.set(key, list);
  };
  let line = firstLine;
  let math: string | null = null;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === "\n") {
      line++;
      continue;
    }
    if (c === "%") {
      const close = text.indexOf("\n", i);
      const comment = text.slice(i + 1, close < 0 ? text.length : close).trim();
      // "%" ending a line of the export's own (\setlength{\unitlength}{\linewidth}%) says nothing
      if (comment) add(`%${comment}`, line);
      i = (close < 0 ? text.length : close) - 1;
      continue;
    }
    if (c === "$") {
      const display = text[i + 1] === "$";
      const delimiter = display ? "$$" : "$";
      if (math === null) math = delimiter;
      else if (math === delimiter) math = null;
      if (display) i++;
      continue;
    }
    if (c !== "\\") continue;
    const env = /^\\(begin|end)\s*\{([^{}\n]*)\}/.exec(text.slice(i, i + 200));
    if (env) {
      const name = env[2].trim();
      if (env[1] === "begin") {
        if (/^(equation|align|gather|multline|eqnarray|flalign|alignat)\*?$/.test(name) && math === null) math = `\\end{${name}}`;
        else if (VERBATIM.test(name)) {
          const close = text.indexOf(`\\end{${name}}`, i);
          const skipped = text.slice(i, close < 0 ? text.length : close);
          line += skipped.split("\n").length - 1;
          i = (close < 0 ? text.length : close + `\\end{${name}}`.length) - 1;
          if (!READ_ENVIRONMENTS.has(name)) add(`env:${name}`, line);
          continue;
        }
        if (math === null) add(`env:${name}`, line);
      } else if (math === `\\end{${name}}`) math = null;
      i += env[0].length - 1;
      continue;
    }
    const name = /^\\([A-Za-z]+|.)/.exec(text.slice(i, i + 64))?.[1] ?? "";
    if (name === "[" || name === "(") {
      if (math === null) math = name === "[" ? "\\]" : "\\)";
    } else if (name === "]" || name === ")") {
      if (math === `\\${name}`) math = null;
    } else if (math === null && name && !ACCENT_COMMANDS.has(name)) add(`cmd:${SAME_AS[name] ?? name}`, line);
    i += name.length;
  }
  return { counts, lines };
}

/**
 * What reading this LaTeX would not keep: the commands and environments the document's LaTeX (`written`, the
 * export of what it was read as) has fewer of, that the reader does not take in; comments; preamble lines the
 * export does not write (the page settings it reads are written back).
 */
export function latexLosses(text: string, written: string): SourceLoss[] {
  const losses: SourceLoss[] = [];
  const bodyAt = (source: string) => {
    const begin = /\\begin\s*\{document\}/.exec(source);
    if (!begin) return { at: 0, end: source.length };
    const end = source.search(/\\end\s*\{document\}/);
    return { at: begin.index + begin[0].length, end: end < 0 ? source.length : end };
  };
  const mine = bodyAt(text);
  const theirs = bodyAt(written);
  const firstLine = lineAt(text, mine.at);
  const user = bodyTokens(text.slice(mine.at, mine.end), firstLine);
  const exported = bodyTokens(written.slice(theirs.at, theirs.end));
  for (const [key, count] of user.counts) {
    if (count <= (exported.counts.get(key) ?? 0)) continue;
    const [kind, name] = key.startsWith("%") ? ["comment", `% ${key.slice(1)}`] : [key.slice(0, key.indexOf(":")), key.slice(key.indexOf(":") + 1)];
    if (kind === "cmd" && READ_COMMANDS.has(name)) continue;
    if (kind === "env" && READ_ENVIRONMENTS.has(name)) continue;
    losses.push({ kind: kind === "cmd" ? "command" : kind === "env" ? "environment" : "comment", name: kind === "cmd" ? `\\${name}` : name, lines: user.lines.get(key) ?? [] });
  }
  // the preamble: lines the export does not write (\title, \author, \date are read; the page settings read and written)
  const preamble = text.slice(0, mine.at > 0 ? text.search(/\\begin\s*\{document\}/) : 0);
  if (preamble) {
    const writtenLines = new Set(latexPreamble(written).split("\n").map((l) => l.trim()));
    preamble.split("\n").forEach((raw, index) => {
      const line = raw.trim();
      if (!line || writtenLines.has(line)) return;
      if (/^\\(title|author|date)\b/.test(line)) return;
      // a setting read from it (written back in the export's form)
      if (Object.keys(preambleSettings(line)).length && !/^\\documentclass/.test(line)) return;
      if (/^\\documentclass/.test(line) && /\\documentclass\s*(?:\[[^\]]*\])?\s*\{(article|ctexart)\}/.test(line)) return;
      if (/^%/.test(line)) losses.push({ kind: "comment", name: line, lines: [index + 1] });
      else losses.push({ kind: "preamble", name: line, lines: [index + 1] });
    });
  }
  return losses;
}

/**
 * What reading this Markdown would not keep: the lines the Markdown written for what it was read as (`written`)
 * has not (the file's unchanged blocks are written back as they were, so these are lines read differently).
 */
export function markdownLosses(text: string, written: string): SourceLoss[] {
  if (text.replace(/\s+$/, "") === written.replace(/\s+$/, "")) return [];
  const have = new Map<string, number>();
  for (const line of written.split("\n")) {
    const key = line.trim();
    if (key) have.set(key, (have.get(key) ?? 0) + 1);
  }
  const losses: SourceLoss[] = [];
  text.split("\n").forEach((raw, index) => {
    const key = raw.trim();
    if (!key) return;
    const left = have.get(key) ?? 0;
    if (left > 0) {
      have.set(key, left - 1);
      return;
    }
    losses.push({ kind: "line", name: key, lines: [index + 1] });
  });
  return losses;
}

/**
 * Where a top-level block of the document is in its source, roughly (for showing it): the place of a piece of its
 * text, the one nearest its share of the document if the piece is there more than once. Null: not found.
 */
export function sourcePlaceOf(doc: Node, index: number, source: string): number | null {
  if (index < 0 || index >= doc.childCount) return null;
  const text = doc.child(index).textContent;
  // a piece of plain text (no characters the source escapes)
  const piece = /[\p{L}\p{N}][\p{L}\p{N} ,.]*[\p{L}\p{N}]/u.exec(text)?.[0]?.slice(0, 40);
  if (!piece) return null;
  const share = source.length * (index / Math.max(1, doc.childCount));
  let best: number | null = null;
  for (let at = source.indexOf(piece); at >= 0; at = source.indexOf(piece, at + 1)) {
    if (best === null || Math.abs(at - share) < Math.abs(best - share)) best = at;
  }
  return best;
}
