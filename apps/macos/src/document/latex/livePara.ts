// Live paragraphs (the personal edition's LaTeX documents). While TeX makes the pages anew (half a second to a
// few), the paragraph being edited is typeset at once by the document's live LuaLaTeX server (tex-bridge's
// live.rs: the same TeX, the preamble already read) and drawn over its old lines on the page; what follows it on
// the page moves with it. Only a paragraph of the body made of text, in-line formulas and the usual text
// commands; and only once its text as TeX last typeset it has come out of the server exactly where the page has
// it — the same lines, every glyph within a twentieth of a point — so that what is drawn is what TeX will draw.
import type { SyncRecord, SyncTex } from "./synctex";
import type { LiveResult } from "./texService";

export const SP_PER_BP = 65781.76;

export interface ParagraphChange {
  /** the paragraph in the text TeX last typeset, its first and last lines there (1-based) */
  oldText: string;
  oldFirst: number;
  oldLast: number;
  /** where it starts in both texts (the text before it is the same) */
  start: number;
  /** the paragraph now */
  newText: string;
}

/** The document's own commands that matter here (its preamble's definitions): those that make a heading, a list
 *  item, a display (`structural`: their lines are not text), and those that begin a paragraph (`starters`: \par in
 *  their definition — a line starting with one starts a paragraph). */
export interface DocMacros {
  structural: RegExp | null;
  starters: RegExp | null;
}

export const NO_MACROS: DocMacros = { structural: null, starters: null };

const STRUCTURAL_BODY =
  /\\(?:part|chapter|section|subsection|subsubsection|paragraph|subparagraph|item|begin|end|caption|footnote|marginpar|clearpage|newpage|cleardoublepage|maketitle|includegraphics)(?![A-Za-z@])|\\\[|\\\]|\$\$/;
const STARTER_BODY = /\\(?:par|medskip|bigskip|smallskip|vspace|vskip|noindent)(?![A-Za-z@])/;

/** The definition bodies of the commands a preamble defines (\newcommand and friends, \def). */
function definitions(preamble: string): Map<string, string> {
  const out = new Map<string, string>();
  const code = preamble.replace(/(^|[^\\])%[^\n]*/g, "$1");
  const heads = [
    /\\(?:newcommand|renewcommand|providecommand|DeclareRobustCommand)\*?\s*\{?\s*\\([A-Za-z@]+)\s*\}?\s*(?:\[\d\]\s*)?(?:\[[^\]]*\]\s*)?\{/g,
    /\\(?:NewDocumentCommand|RenewDocumentCommand|DeclareDocumentCommand|ProvideDocumentCommand)\s*\{?\s*\\([A-Za-z@]+)\s*\}?\s*\{[^{}]*\}\s*\{/g,
    /\\[gex]?def\s*\\([A-Za-z@]+)[^{]*\{/g,
  ];
  for (const re of heads) {
    for (let m = re.exec(code); m; m = re.exec(code)) {
      // the body: up to the brace that closes the one the match ends with
      let depth = 1;
      let i = re.lastIndex;
      for (; i < code.length && depth > 0; i++) {
        const c = code[i];
        if (c === "\\") i++;
        else if (c === "{") depth++;
        else if (c === "}") depth--;
      }
      out.set(m[1], code.slice(re.lastIndex, i - 1));
    }
  }
  return out;
}

const commandsRe = (names: string[]) =>
  names.length ? new RegExp(`\\\\(?:${names.map((n) => n.replace(/[@]/g, "@")).join("|")})(?![A-Za-z@])`) : null;

/** What the preamble's own commands are (see DocMacros); one command calling another counts as it. */
export function docMacros(preamble: string): DocMacros {
  const defs = definitions(preamble);
  const structural = new Set<string>();
  const starters = new Set<string>();
  for (let round = 0; round < 3; round++) {
    const structuralRe = commandsRe([...structural]);
    const startersRe = commandsRe([...starters]);
    for (const [name, body] of defs) {
      if (STRUCTURAL_BODY.test(body) || structuralRe?.test(body)) {
        structural.add(name);
        starters.delete(name);
      } else if (STARTER_BODY.test(body) || startersRe?.test(body)) starters.add(name);
    }
  }
  return { structural: commandsRe([...structural]), starters: commandsRe([...starters]) };
}

const blank = (line: string) => /^[ \t\f\r]*$/.test(line);
/** A list item's first line (\item, \item[label]): it starts a paragraph, typeset in its list. */
const ITEM = /^\s*\\item(?![A-Za-z@])/;
const itemLess = (line: string) => line.replace(ITEM, "");
/** A line of a paragraph's text: not blank, and no structure, display or command that makes it more than text
 *  (DENIED below, the document's own structural commands). Such lines (a heading, \begin{equation}, \item) end
 *  the run of lines typeset together. */
const textLine = (line: string, m: DocMacros) => {
  const rest = itemLess(line);
  return !blank(line) && !/\\item(?![A-Za-z@])/.test(rest) && !DENIED.test(rest) && !DISPLAY.test(rest) && !m.structural?.test(rest);
};
/** A line that starts a paragraph of its own (it begins with one of the document's paragraph commands). */
const startsParagraph = (line: string, m: DocMacros) => {
  if (ITEM.test(line)) return true;
  if (!m.starters) return false;
  const head = /^\s*(\\[A-Za-z@]+)/.exec(line);
  return !!head && m.starters.test(head[1]);
};

function lineStart(text: string, pos: number): number {
  return text.lastIndexOf("\n", pos - 1) + 1;
}

function lineEnd(text: string, pos: number): number {
  const i = text.indexOf("\n", pos);
  return i < 0 ? text.length : i;
}

/**
 * The text lines around [from, to] (TeX's paragraph in the source, or the part of one between a display and the
 * paragraph's end): up to a blank line or a line that is more than text, before and after. Null when a line of
 * [from, to] is itself more than text.
 */
function paragraphAround(text: string, from: number, to: number, m: DocMacros): [number, number] | null {
  let start = lineStart(text, from);
  let end = lineEnd(text, to);
  if (!allText(text.slice(start, end), m, true)) return null;
  // up to the line that starts the paragraph (a blank line, a heading or display before it, or a paragraph command)
  while (start > 0 && !startsParagraph(text.slice(start, lineEnd(text, start)), m)) {
    const prev = lineStart(text, start - 1);
    if (!textLine(text.slice(prev, start - 1), m)) break;
    start = prev;
  }
  while (end < text.length) {
    const next = lineEnd(text, end + 1);
    const line = text.slice(end + 1, next);
    if (!textLine(line, m) || startsParagraph(line, m)) break;
    end = next;
  }
  return [start, end];
}

/** Text lines, a paragraph command only on the first (`anyFirst`: on any — the lines an edit touched). */
const allText = (text: string, m: DocMacros, anyFirst = false) =>
  text.split("\n").every((line, i) => textLine(line, m) && (i === 0 || anyFirst || !startsParagraph(line, m)));
const count = (text: string, c: string) => {
  let n = 0;
  for (let i = text.indexOf(c); i >= 0; i = text.indexOf(c, i + 1)) n++;
  return n;
};

/** The one paragraph an edit changed (old: the text TeX last typeset), else null (no change, several). */
export function changedParagraph(old: string, now: string, m: DocMacros = NO_MACROS): ParagraphChange | null {
  if (old === now) return null;
  const max = Math.min(old.length, now.length);
  let p = 0;
  while (p < max && old.charCodeAt(p) === now.charCodeAt(p)) p++;
  let s = 0;
  while (s < max - p && old.charCodeAt(old.length - 1 - s) === now.charCodeAt(now.length - 1 - s)) s++;
  const around = paragraphAround(now, p, Math.max(p, now.length - s), m);
  if (!around) return null;
  const [start, end] = around;
  const oldEnd = end - (now.length - old.length);
  if (oldEnd < old.length - s || oldEnd <= start) return null;
  const oldText = old.slice(start, oldEnd);
  const newText = now.slice(start, end);
  // the same run of lines before the edit (it did not join two, nor take a heading or a display away)
  if (!allText(oldText, m) || !allText(newText, m)) return null;
  const oldFirst = count(old.slice(0, start), "\n") + 1;
  return { oldText, oldFirst, oldLast: oldFirst + count(oldText, "\n"), start, newText };
}

const DISPLAY = /\\\[|\\\]|\$\$/;

// Commands that make a paragraph more than a paragraph: structure, floats and notes, pages, definitions and
// counters (they would stay in the server), colour (drawn black here), what TeX reads from files.
const DENIED = new RegExp(
  "\\\\(?:" +
    [
      "begin", "end", "bibitem", "part", "chapter", "section", "subsection", "subsubsection", "paragraph", "subparagraph",
      "footnote", "footnotemark", "footnotetext", "marginpar", "marginnote", "todo", "caption", "includegraphics",
      "includepdf", "input", "include", "InputIfFileExists", "newpage", "clearpage", "cleardoublepage", "pagebreak",
      "enlargethispage", "maketitle", "tableofcontents", "listoffigures", "listoftables", "bibliography",
      "printbibliography", "appendix", "frontmatter", "mainmatter", "backmatter", "def", "gdef", "edef", "xdef", "let",
      "newcommand", "renewcommand", "providecommand", "DeclareRobustCommand", "NewDocumentCommand",
      "RenewDocumentCommand", "DeclareDocumentCommand", "newenvironment", "renewenvironment", "setlength",
      "addtolength", "settowidth", "setcounter", "addtocounter", "stepcounter", "refstepcounter", "newcounter",
      "newlength", "color", "textcolor", "colorbox", "fcolorbox", "pagecolor", "definecolor", "hl", "endinput", "stop",
      "bye", "documentclass", "usepackage", "RequirePackage", "makeatletter", "makeatother", "catcode", "special",
      "directlua", "luaexec", "write", "immediate", "openout", "tikz", "pgf", "draw", "node",
    ].join("|") +
    ")(?![A-Za-z@])",
);
// environments whose content is not text (code, pictures, tables, displays, floats): a paragraph inside one is not
// typeset alone
const DENIED_ENVIRONMENTS = /^(?:verbatim\*?|Verbatim\*?|BVerbatim|LVerbatim|lstlisting|minted|comment|filecontents\*?|alltt|tikzpicture|pgfpicture|picture|tabular[x*]?|tabulary|longtable|array|align\*?|alignat\*?|flalign\*?|equation\*?|eqnarray\*?|gather\*?|multline\*?|displaymath|math|figure\*?|table\*?|wrapfigure|thebibliography|lstinputlisting|tcolorbox|mdframed|framed|minipage|center|flushleft|flushright|titlepage)$/;

/** Braces balance (escaped ones and comments aside), never closing more than they opened. */
function balanced(text: string): boolean {
  let depth = 0;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === "\\") {
      i++;
      continue;
    }
    if (c === "%") {
      const nl = text.indexOf("\n", i);
      if (nl < 0) break;
      i = nl;
      continue;
    }
    if (c === "{") depth++;
    else if (c === "}" && --depth < 0) return false;
  }
  return depth === 0;
}

/** The environments open at `pos` (comments aside), innermost last. */
function environmentsAt(text: string, pos: number): string[] {
  const stack: string[] = [];
  const re = /\\(begin|end)\s*\{([^}]*)\}|%[^\n]*/g;
  re.lastIndex = 0;
  for (let m = re.exec(text); m && m.index < pos; m = re.exec(text)) {
    if (!m[1]) continue;
    if (m.index > 0 && text[m.index - 1] === "\\") continue;
    if (m[1] === "begin") stack.push(m[2].trim());
    else {
      const at = stack.lastIndexOf(m[2].trim());
      if (at >= 0) stack.length = at;
    }
  }
  return stack;
}

/** The lists a paragraph starting with \item is in (from the body's start; comments aside): each open
 *  environment's \begin as written, outermost first, and how many items come before it in the innermost one. */
export interface ListContext {
  begins: string[];
  ends: string[];
  before: number;
  /** the innermost environment's name */
  name: string;
}

export function listContext(text: string, pos: number, main: boolean): ListContext | null {
  type Open = { name: string; begin: string; items: number };
  const stack: Open[] = [];
  let from = 0;
  if (main) {
    const at = text.search(/\\begin\s*\{document\}/);
    if (at < 0 || at > pos) return null;
    from = text.indexOf("\n", at) + 1;
  }
  const code = text.slice(from, pos).replace(/(^|[^\\])%[^\n]*/g, "$1");
  const re = /\\(begin|end)\s*\{([^}]*)\}((?:\s*\[[^\]]*\])?(?:\s*\{[^{}]*\})*)|\\item(?![A-Za-z@])/g;
  for (let m = re.exec(code); m; m = re.exec(code)) {
    if (m.index > 0 && code[m.index - 1] === "\\") continue;
    if (!m[1]) {
      if (stack.length) stack[stack.length - 1].items++;
    } else if (m[1] === "begin") stack.push({ name: m[2].trim(), begin: m[0], items: 0 });
    else {
      const at = stack.map((s) => s.name).lastIndexOf(m[2].trim());
      if (at >= 0) stack.length = at;
    }
  }
  if (!stack.length) return null;
  const inner = stack[stack.length - 1];
  // the lists only (an environment around them that is not one, a theorem's or a proof's, would bring its head):
  // an outer list holds the inner one in an item of its own, an item without a label (no bullet drawn)
  const lists = stack.filter((s, i) => i === stack.length - 1 || s.items > 0);
  const begins = lists.map((s, i) => (i < lists.length - 1 ? `${s.begin}\\item[]` : s.begin));
  return { begins, ends: lists.map((s) => `\\end{${s.name}}`).reverse(), before: inner.items, name: inner.name };
}

/** The proof the paragraph opens (\begin{proof} as written right before it, \label lines aside), else null: the
 *  proof's head starts the paragraph, set in the proof. */
export function proofOpening(text: string, change: ParagraphChange): string | null {
  let before = text.slice(0, change.start).replace(/(^|[^\\])%[^\n]*/g, "$1");
  for (;;) {
    const trimmed = before.replace(/\s+$/, "");
    const label = /\\label\s*\{[^{}]*\}$/.exec(trimmed);
    if (!label) {
      const open = /\\begin\s*\{proof\}(\s*\[[^\]]*\])?$/.exec(trimmed);
      return open ? open[0] : null;
    }
    before = trimmed.slice(0, label.index);
  }
}

/** What the live server is sent for a proof's first paragraph: the proof opened as written, the paragraph, the proof
 *  closed — its QED symbol only when the paragraph is the proof's last. */
export function proofText(prefix: string, opening: string, paragraph: string, last: boolean): string {
  return (prefix ? prefix + "\n" : "") + opening + "\n" + paragraph + "\n" + (last ? "" : "\\renewcommand{\\qedsymbol}{}") + "\\end{proof}";
}

/** The paragraph ends a proof (\end{proof} follows it): it ends with the QED symbol as \end{proof} puts it. */
export function endsProof(text: string, change: ParagraphChange): boolean {
  const after = text.slice(change.start + change.newText.length);
  return /^\s*\n\s*\\end\s*\{proof\}/.test(after);
}

/** The paragraph is a list item (it starts with \item). */
export const isListItem = (paragraph: string) => ITEM.test(paragraph);

/** What the live server is sent for a list item: its lists opened as the source opens them (an enumerated list
 *  counting the items before it), the item, the lists closed. */
export function listText(prefix: string, ctx: ListContext, enumerated: boolean, paragraph: string): string {
  const count = enumerated && ctx.before > 0 ? `\n\\makeatletter\\addtocounter{\\@enumctr}{${ctx.before}}\\makeatother` : "";
  return (prefix ? prefix + "\n" : "") + ctx.begins.join("\n") + count + "\n" + paragraph + "\n" + ctx.ends.join("\n");
}

/** An environment that numbers its items (enumerate, or one of the document's defined in terms of one). */
export function enumerated(name: string, preamble: string): boolean {
  if (/enum/i.test(name)) return true;
  const def = new RegExp(`\\\\(?:newenvironment|renewenvironment)\\*?\\s*\\{${name.replace(/[^A-Za-z0-9*]/g, "")}\\}[^\\n]*?\\{([^\\n]*)`).exec(preamble);
  return !!def && /\\begin\s*\{enumerate\*?\}/.test(def[1]);
}

/** The text without its formulas (what is inside them — matrices, cases, \begin{aligned} — is TeX's math, not the
 *  paragraph's structure): each one a space. */
export function withoutFormulas(text: string): string {
  let out = "";
  let i = 0;
  const close = (open: string) => (open === "$" ? "$" : open === "$$" ? "$$" : open === "\\(" ? "\\)" : "\\]");
  while (i < text.length) {
    const c = text[i];
    if (c === "\\" && (text[i + 1] === "(" || text[i + 1] === "[")) {
      const open = text.slice(i, i + 2);
      const end = text.indexOf(close(open), i + 2);
      if (end < 0) return out + text.slice(i);
      out += " ";
      i = end + 2;
      continue;
    }
    if (c === "\\") {
      const env = /^\\begin\s*\{((?:equation|align|alignat|gather|multline|flalign|eqnarray|displaymath|math)\*?)\}/.exec(text.slice(i, i + 40));
      if (env) {
        const end = text.indexOf(`\\end{${env[1]}}`, i);
        if (end < 0) return out + text.slice(i);
        out += " ";
        i = end + `\\end{${env[1]}}`.length;
        continue;
      }
      out += text.slice(i, i + 2);
      i += 2;
      continue;
    }
    if (c === "$") {
      const open = text.startsWith("$$", i) ? "$$" : "$";
      let j = i + open.length;
      while (j < text.length && !(text.startsWith(open, j) && text[j - 1] !== "\\")) j++;
      if (j >= text.length) return out + text.slice(i);
      out += " ";
      i = j + open.length;
      continue;
    }
    out += c;
    i++;
  }
  return out;
}

/** A paragraph the live server may typeset: `text` the file's text now, `main` whether it is the main file;
 *  `displays`: its formulas shown apart (\[…\], $$…$$, unnumbered display environments) are part of it. */
export function liveEligible(change: ParagraphChange, text: string, main: boolean, displays = false): boolean {
  for (const t of [change.newText, change.oldText]) {
    // (an \item only as the paragraph's first word)
    let rest = itemLess(t);
    if (displays) rest = withoutFormulas(rest.replace(MATH_ENV_LINE, ""));
    if (!/\S/.test(t) || /\\item(?![A-Za-z@])/.test(rest) || DENIED.test(rest) || (!displays && DISPLAY.test(rest)) || !balanced(t)) return false;
  }
  const envs = environmentsAt(text, change.start);
  if (main) {
    // the body only
    if (envs[0] !== "document") return false;
    envs.shift();
  }
  return !envs.some((e) => DENIED_ENVIRONMENTS.test(e));
}

// Declarations of the body that hold for what follows them (a size, a family, paragraph shape and spacing): the
// ones before a paragraph at the body's top level go before it in the live server too.
const DECLARATION = new RegExp(
  "^\\s*(?:" +
    [
      "\\\\(?:tiny|scriptsize|footnotesize|small|normalsize|large|Large|LARGE|huge|Huge|normalfont|rmfamily|sffamily|ttfamily|mdseries|bfseries|upshape|itshape|slshape|scshape|sloppy|fussy|raggedright|raggedleft|centering|selectfont)(?![A-Za-z@])\\s*",
      "\\\\(?:setlength|addtolength)\\s*\\{?\\s*\\\\(?:parindent|parskip|baselineskip|lineskip|lineskiplimit|emergencystretch|tolerance)\\s*\\}?\\s*\\{[^{}]*\\}\\s*",
      "\\\\linespread\\s*\\{[^{}]*\\}\\s*(?:\\\\selectfont\\s*)?",
      "\\\\(?:tolerance|pretolerance|emergencystretch|parindent|parskip|hyphenpenalty|exhyphenpenalty|looseness)\\s*=?\\s*-?[\\d.]+\\s*(?:pt|em|ex|bp|mm|cm|in)?\\s*",
    ].join("|") +
    ")+(?:%.*)?$",
);

/** The body's declarations at its top level before `start` (the main file's body; another file's text before it). */
export function bodyPrefix(text: string, start: number, main: boolean): string {
  let from = 0;
  if (main) {
    const at = text.search(/\\begin\s*\{document\}/);
    if (at < 0 || at > start) return "";
    from = text.indexOf("\n", at) + 1;
  }
  const out: string[] = [];
  let depth = 0;
  for (const raw of text.slice(from, start).split("\n")) {
    const line = raw.replace(/(^|[^\\])%.*$/, "$1");
    if (depth === 0 && DECLARATION.test(line)) out.push(line.trim());
    for (const m of line.matchAll(/\\(begin|end)\s*\{[^}]*\}|(?<!\\)[{}]/g)) {
      if (m[1] === "begin" || m[0] === "{") depth++;
      else if (m[1] === "end" || m[0] === "}") depth = Math.max(0, depth - 1);
    }
  }
  return out.join("\n");
}

/** What the live server is sent for a paragraph: the body's declarations, then the indent as asked
 *  (-1: \indent, 0: \noindent, else that much in sp — a list's paragraphs are indented by its own amount), then
 *  the paragraph. */
export function liveText(prefix: string, indent: number, paragraph: string): string {
  const head = indent < 0 ? "\\indent " : indent === 0 ? "\\noindent " : `\\noindent\\kern${Math.round(indent)}sp\\relax `;
  return (prefix ? prefix + "\n" : "") + head + paragraph;
}

/** The paragraph begins with one of the document's paragraph commands (it sets its own indent). */
export function beginsParagraphCommand(paragraph: string, m: DocMacros): boolean {
  return startsParagraph(paragraph.split("\n")[0] ?? "", m);
}

// ---- the units the pages are edited in ----

/** \begin and \end of an unnumbered display environment (a formula shown apart, in a paragraph). */
const MATH_ENV_LINE = /\\(?:begin|end)\s*\{(?:equation\*|align\*|gather\*|multline\*|flalign\*|alignat\*|eqnarray\*|displaymath)\}/g;
/** Lines that are a unit of their own (not part of a paragraph): structure, floats' parts, files, page breaks. */
const UNIT_BREAK =
  /^\s*\\(?:part|chapter|section|subsection|subsubsection|paragraph|subparagraph|maketitle|tableofcontents|listoffigures|listoftables|newpage|clearpage|cleardoublepage|pagebreak|appendix|frontmatter|mainmatter|backmatter|bibliography|printbibliography|bibliographystyle|includegraphics|caption|centering|input|include|documentclass|usepackage|title|author|date|hline|toprule|midrule|bottomrule)(?![A-Za-z@])/;
const ENV_LINE = /^\s*\\(begin|end)\s*\{([^}]*)\}/;
const MATH_ENV_NAME = /^(?:equation|align|alignat|gather|multline|flalign|eqnarray|displaymath|math|split|gathered|aligned|cases|array|matrix|pmatrix|bmatrix|vmatrix|Vmatrix)\*?$/;

/** Display environments that number their lines (their numbers are the document's counters: a unit of their own,
 *  not typeset with the paragraph around them). */
const NUMBERED_ENV = /\\(begin|end)\s*\{(equation|align|alignat|gather|multline|flalign|eqnarray)\}/;

/** A line that is not part of a paragraph: blank, structure, an environment's begin or end (an unnumbered
 *  formula's aside). */
function breaksUnit(line: string, m: DocMacros): boolean {
  if (blank(line) || UNIT_BREAK.test(line) || NUMBERED_ENV.test(line)) return true;
  const env = ENV_LINE.exec(line);
  if (env && !MATH_ENV_NAME.test(env[2].trim())) return true;
  const head = /^\s*(\\[A-Za-z@]+)/.exec(line);
  return !!head && !!m.structural?.test(head[1]);
}

/** The numbered display environment `pos` is in (from its \begin line to its \end), else null. */
function numberedBlock(text: string, pos: number): [number, number] | null {
  const re = new RegExp(NUMBERED_ENV.source, "g");
  let open: { name: string; at: number } | null = null;
  for (let m = re.exec(text); m; m = re.exec(text)) {
    if (m.index > pos && !open) return null;
    if (m[1] === "begin" && !open) open = { name: m[2], at: m.index };
    else if (m[1] === "end" && open && m[2] === open.name) {
      const end = m.index + m[0].length;
      if (pos >= open.at && pos <= end) return [lineStart(text, open.at), lineEnd(text, end)];
      open = null;
    }
  }
  return null;
}

/**
 * The unit of the source the caret at `pos` edits on the pages: its paragraph — the lines between blank lines and
 * structure, formulas shown apart included; a list item from its \item; a paragraph command's paragraph from it —
 * or the structure line it is on (a heading). [from, to): from its first line's first non-blank character.
 */
export function visualUnit(text: string, pos: number, m: DocMacros = NO_MACROS, splitDisplays = false): [number, number] {
  let start = lineStart(text, pos);
  let end = lineEnd(text, pos);
  const indent = (from: number) => {
    let i = from;
    while (i < text.length && (text[i] === " " || text[i] === "\t")) i++;
    return i;
  };
  // (`splitDisplays`: a line with a formula shown apart is a unit of its own, as the source view's live paragraphs
  // have them — for a paragraph that the live server cannot typeset with its displays)
  const breaks = (line: string) => breaksUnit(line, m) || (splitDisplays && (DISPLAY.test(line) || /\\(?:begin|end)\s*\{/.test(line)));
  const block = numberedBlock(text, pos);
  if (block) return [indent(block[0]), block[1]];
  if (breaks(text.slice(start, end))) return [indent(start), end];
  while (start > 0 && !startsParagraph(text.slice(start, lineEnd(text, start)), m)) {
    const prev = lineStart(text, start - 1);
    if (breaks(text.slice(prev, start - 1))) break;
    start = prev;
  }
  while (end < text.length) {
    const next = lineEnd(text, end + 1);
    const line = text.slice(end + 1, next);
    if (breaks(line) || startsParagraph(line, m)) break;
    end = next;
  }
  return [indent(start), end];
}

/** The 1-based line of offset `pos`. */
export function lineNumberAt(text: string, pos: number): number {
  return count(text.slice(0, pos), "\n") + 1;
}

// ---- where the paragraph is on the pages ----

export interface LineBox {
  h: number;
  v: number;
  width: number;
  height: number;
  depth: number;
}

export interface ParagraphPlace {
  page: number;
  boxes: LineBox[];
  /** where SyncTeX says each line's glues end (bp) */
  glues: number[][];
}

/** The lines of the page's text around a paragraph: the line boxes of its vbox in order, and the run of them whose
 *  records are the paragraph's (`core`, first and last index). */
export interface ParagraphArea {
  page: number;
  lines: LineBox[];
  glues: number[][];
  core: [number, number];
  /** the width of the vbox the lines are in (the text's, a column's: a list item's lines are narrower) */
  width: number;
}

/**
 * Where a paragraph's lines are (source lines first..last of SyncTeX files `files`): the hboxes of the page's text
 * whose glyph records are mostly the paragraph's — on one page, next to each other. (LuaTeX tags some of a line's
 * records with the lines before it: a paragraph's first line may not be among them; see placesOf.)
 */
export function paragraphArea(tex: SyncTex, files: number[], first: number, last: number): ParagraphArea | null {
  let found: ParagraphArea | null = null;
  for (const [n, page] of tex.pages) {
    const recs = page.records;
    // a record's line box: the hbox it is in that sits in a vbox (the page's text, a minipage's)
    const lineOf = (i: number): number => {
      let at = recs[i].parent;
      while (at >= 0) {
        const r = recs[at];
        if (r.kind === "(" && (r.parent < 0 || recs[r.parent].kind === "[")) return at;
        at = r.parent;
      }
      return -1;
    };
    const inside = new Map<number, number>();
    const outside = new Map<number, number>();
    const glueEnds = new Map<number, number[]>();
    recs.forEach((r: SyncRecord, i) => {
      // kerns, glue, maths carry the line TeX read them on ("x" records the line its paragraph ended on: a list
      // item's \item line on the item before)
      if (!"kg$".includes(r.kind)) return;
      const line = lineOf(i);
      if (line < 0) return;
      if (r.kind === "g") glueEnds.set(line, [...(glueEnds.get(line) ?? []), r.h]);
      // what ends a paragraph sits at its last line's right edge, made on the line that ended it (the next
      // item's): it does not tell the line's own
      if (r.h >= recs[line].h + recs[line].width - 0.5) return;
      const mine = files.includes(r.file) && r.line >= first && r.line <= last;
      const map = mine ? inside : outside;
      map.set(line, (map.get(line) ?? 0) + 1);
    });
    const chosen = [...inside.keys()].filter((i) => (inside.get(i) ?? 0) >= (outside.get(i) ?? 0));
    if (!chosen.length) continue;
    if (found) return null; // on two pages
    const parent = recs[chosen[0]].parent;
    if (chosen.some((i) => recs[i].parent !== parent)) return null;
    // the vbox's lines in order; the chosen ones must be a run of them
    const all = recs.map((_, i) => i).filter((i) => recs[i].kind === "(" && recs[i].parent === parent).sort((x, y) => recs[x].v - recs[y].v);
    const at = chosen.map((i) => all.indexOf(i)).sort((x, y) => x - y);
    if (at[at.length - 1] - at[0] !== at.length - 1) return null;
    found = {
      page: n,
      lines: all.map((i) => ({ h: recs[i].h, v: recs[i].v, width: recs[i].width, height: recs[i].height, depth: recs[i].depth })),
      glues: all.map((i) => glueEnds.get(i) ?? []),
      core: [at[0], at[at.length - 1]],
      width: parent >= 0 ? recs[parent].width : recs[chosen[0]].width,
    };
  }
  return found;
}

/** The places of a paragraph of `count` lines in its area: each run of that many lines of one width holding the
 *  core (one when the core is all of it). */
export function placesOf(area: ParagraphArea, count: number): ParagraphPlace[] {
  const out: ParagraphPlace[] = [];
  const [c0, c1] = area.core;
  for (let start = Math.max(0, c1 - count + 1); start <= c0 && start + count <= area.lines.length; start++) {
    const boxes = area.lines.slice(start, start + count);
    // (a display's line is as wide as its formula: the check compares each line's width)
    out.push({ page: area.page, boxes, glues: area.glues.slice(start, start + count) });
  }
  return out;
}

// ---- the page's glyphs (its SVG) and the check ----

export interface PageGlyph {
  x: number;
  y: number;
}

/** A glyph of a page's SVG: where TeX put it (`x`, `y`: its <text>'s own coordinates — SyncTeX's and the live
 *  server's) and how far the groups around it move it on the page (`tx`, `ty`: a box drawn with a transform, a
 *  tcolorbox's content, shows elsewhere than its coordinates say), bp from the page's top left. */
export interface PageGlyph {
  x: number;
  y: number;
  tx: number;
  ty: number;
}

export interface PageChar extends PageGlyph {
  text: string;
  size: number;
  /** its font is a formula's (true), a text font (false), not known (null) */
  math: boolean | null;
}

type Matrix = [number, number, number, number, number, number];
const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];
const multiply = (m: Matrix, n: Matrix): Matrix => [
  m[0] * n[0] + m[2] * n[1],
  m[1] * n[0] + m[3] * n[1],
  m[0] * n[2] + m[2] * n[3],
  m[1] * n[2] + m[3] * n[3],
  m[0] * n[4] + m[2] * n[5] + m[4],
  m[1] * n[4] + m[3] * n[5] + m[5],
];
/** An SVG transform attribute (matrix, translate, scale; several in a row). */
function parseTransform(attr: string): Matrix {
  let m = IDENTITY;
  for (const t of attr.matchAll(/(matrix|translate|scale)\s*\(([^)]*)\)/g)) {
    const v = t[2].split(/[\s,]+/).filter(Boolean).map(Number);
    const n: Matrix =
      t[1] === "matrix" && v.length === 6
        ? (v as Matrix)
        : t[1] === "translate"
          ? [1, 0, 0, 1, v[0] ?? 0, v[1] ?? 0]
          : t[1] === "scale"
            ? [v[0] ?? 1, 0, 0, v[1] ?? v[0] ?? 1, 0, 0]
            : IDENTITY;
    m = multiply(m, n);
  }
  return m;
}

/** TFM math fonts by their names (Computer Modern's, AMS's, the usual math packages'): their glyphs are a
 *  formula's. */
const MATH_FONT = /^(?:cmmi|cmsy|cmex|cmbsy|cmmib|msam|msbm|eufm|eusm|eurm|rsfs|stmary|wasy|lasy|esint|txsy|txmi|txex|txsym|pxsy|pxmi|pxex|zxmi|zxsy|lmmi|lmsy|lmex|mathkerncmssi|ntxmi|ntxsy|ntxexx|newtxmi|bbold|dsrom|futm|fourier)/i;

/** The page's glyphs (dvisvgm: one <text> a glyph) with the groups' transforms around them, their characters and
 *  fonts' sizes; `page` false: a live paragraph's SVG (its own user units, from its top left). */
export function pageChars(svg: string, page = true): PageChar[] {
  const box = page ? /viewBox=['"]([-\d.e]+)[ ,]+([-\d.e]+)/.exec(svg) : null;
  const dx = box ? -Number(box[1]) : 0;
  const dy = box ? -Number(box[2]) : 0;
  const sizes = new Map<string, number>();
  const families = new Map<string, string>();
  for (const m of svg.matchAll(/text\.([\w-]+)\s*\{([^}]*)\}/g)) {
    const size = /font-size:\s*([\d.]+)px/.exec(m[2]);
    const family = /font-family:\s*([^;]+)/.exec(m[2]);
    if (size) sizes.set(m[1], Number(size[1]));
    if (family) families.set(m[1], family[1].trim());
  }
  const decode = (t: string) =>
    t.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (_, e: string) =>
      e[0] === "#" ? String.fromCodePoint(e[1] === "x" ? parseInt(e.slice(2), 16) : Number(e.slice(1))) : ({ amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" } as Record<string, string>)[e],
    );
  const out: PageChar[] = [];
  const stack: Matrix[] = [IDENTITY];
  const re = /<g\b([^>]*?)(\/?)>|<\/g>|<text\b([^>]*)>([^<]*)<\/text>/g;
  for (let m = re.exec(svg); m; m = re.exec(svg)) {
    if (m[0].startsWith("</g")) {
      if (stack.length > 1) stack.pop();
      continue;
    }
    if (m[0].startsWith("<g")) {
      if (m[2]) continue; // <g/>
      const t = /\btransform=['"]([^'"]*)['"]/.exec(m[1] ?? "");
      stack.push(t ? multiply(stack[stack.length - 1], parseTransform(t[1])) : stack[stack.length - 1]);
      continue;
    }
    const attrs = m[3] ?? "";
    const x = /\bx=['"]([-\d.e]+)['"]/.exec(attrs);
    const y = /\by=['"]([-\d.e]+)['"]/.exec(attrs);
    if (!x || !y) continue;
    let mt = stack[stack.length - 1];
    const own = /\btransform=['"]([^'"]*)['"]/.exec(attrs);
    if (own) mt = multiply(mt, parseTransform(own[1]));
    const rx = Number(x[1]);
    const ry = Number(y[1]);
    const cls = /\bclass=['"]([^'"]+)['"]/.exec(attrs);
    const family = (cls && families.get(cls[1])) || "";
    // a TFM math font: a formula's; another TFM font (its name a TeX font's): text; a native font (nf…): not known
    const math = MATH_FONT.test(family) ? true : /^nf\d+$/.test(family) || !family ? null : false;
    out.push({
      x: rx + dx,
      y: ry + dy,
      tx: mt[0] * rx + mt[2] * ry + mt[4] - rx,
      ty: mt[1] * rx + mt[3] * ry + mt[5] - ry,
      text: decode(m[4] ?? ""),
      size: (cls && sizes.get(cls[1])) || 10,
      math,
    });
  }
  return out;
}

/** Where the page's SVG has each glyph (see PageGlyph). */
export function pageGlyphs(svg: string, page = true): PageGlyph[] {
  return pageChars(svg, page).map(({ x, y, tx, ty }) => ({ x, y, tx, ty }));
}

const TOLERANCE = 0.05;
/** How far a glyph may be from the page's (bp): TeX distributions differ by a tenth of a point in a CJK
 *  punctuation's spacing (xeCJK versions) — invisible; the lines themselves (their baselines, widths, breaks) must
 *  be TeX's exactly. */
const GLYPH_TOLERANCE = 0.35;

/**
 * The live server's paragraph (`result`, its text as TeX last typeset it) against the page: the same lines at
 * the same baselines and, where the page's glyphs are known, every glyph where the page has one (and no more on
 * the page in the paragraph's place). Returns where the paragraph's box starts on the page (its top left).
 */
export function verify(
  result: LiveResult,
  place: ParagraphPlace,
  glyphs: PageGlyph[] | null,
  report?: { matched: number; onPage: number; misses: [number, number][] },
): { x0: number; y0: number; tx: number; ty: number } | null {
  if (result.status !== "ok" || result.undrawable > 0 || result.lines.length !== place.boxes.length || !result.lines.length) return null;
  // the paragraph's left edge: its first line's, less that line's shift (a list item's lines are moved in)
  const x0 = place.boxes[0].h - result.lines[0][4];
  const y0 = place.boxes[0].v - result.lines[0][0];
  for (let i = 0; i < result.lines.length; i++) {
    if (Math.abs(y0 + result.lines[i][0] - place.boxes[i].v) > TOLERANCE) return null;
    if (Math.abs(result.lines[i][3] - place.boxes[i].width) > TOLERANCE) return null;
  }
  if (!glyphs) {
    // a picture of the page: each line's glues end where SyncTeX recorded them (LuaTeX's records place a line by
    // its glue, not its glyphs)
    let recorded = 0;
    let matched = 0;
    for (let i = 0; i < place.boxes.length; i++) {
      const baseline = place.boxes[i].v;
      const ends = result.glues.filter(([, gy]) => Math.abs(y0 + gy - baseline) <= TOLERANCE).map(([gx]) => x0 + gx);
      for (const h of place.glues[i]) {
        recorded++;
        if (ends.some((e) => Math.abs(e - h) <= TOLERANCE)) matched++;
      }
    }
    if (report) {
      report.matched = matched;
      report.onPage = recorded;
    }
    return recorded >= place.boxes.length && matched === recorded ? { x0, y0, tx: 0, ty: 0 } : null;
  }
  const last = place.boxes[place.boxes.length - 1];
  const top = place.boxes[0].v - place.boxes[0].height - 0.5;
  const bottom = last.v + last.depth + 0.5;
  const left = x0 - 1;
  const right = Math.max(...place.boxes.map((b) => b.h + b.width)) + 1;
  // the page's glyphs in the paragraph's place, by baseline
  const byLine = new Map<number, PageGlyph[]>();
  let onPage = 0;
  for (const g of glyphs) {
    if (g.y < top || g.y > bottom || g.x < left || g.x > right) continue;
    onPage++;
    const key = Math.round(g.y * 10);
    for (const k of [key - 1, key, key + 1]) {
      const list = byLine.get(k);
      if (list) list.push(g);
      else byLine.set(k, [g]);
    }
  }
  let matched = 0;
  // how far the page draws the matched glyphs from their coordinates (a box with a transform): the most common
  const moves = new Map<string, { tx: number; ty: number; n: number }>();
  const own: [number, number][] = result.svg ? pageGlyphs(result.svg, false).map((g) => [g.x, g.y]) : result.glyphs.map(([, gx, gy]) => [gx, gy]);
  for (const [gx, gy] of own) {
    const x = x0 + gx;
    const y = y0 + gy;
    const list = byLine.get(Math.round(y * 10));
    const hit = list?.find((g) => Math.abs(g.x - x) <= GLYPH_TOLERANCE && Math.abs(g.y - y) <= TOLERANCE);
    if (hit) {
      matched++;
      const k = `${hit.tx.toFixed(2)},${hit.ty.toFixed(2)}`;
      const e = moves.get(k);
      if (e) e.n++;
      else moves.set(k, { tx: hit.tx, ty: hit.ty, n: 1 });
    } else report?.misses.push([x, y]);
  }
  const n = own.length;
  if (report) {
    report.matched = matched;
    report.onPage = onPage;
  }
  if (matched < n - Math.floor(n * 0.01) || Math.abs(onPage - n) > Math.max(1, Math.floor(n * 0.01))) return null;
  const move = [...moves.values()].sort((a, b) => b.n - a.n)[0] ?? { tx: 0, ty: 0 };
  return { x0, y0, tx: move.tx, ty: move.ty };
}

// ---- the drawing ----

export interface LiveOverlay {
  page: number;
  /** bp: the old lines' top and bottom (the page is shown above and below them), how far what follows moves */
  top: number;
  bottom: number;
  shift: number;
  /** the new lines: SVG content in the page's bp */
  svg: string;
  /** places where the page itself stays shown over the new lines (a list's bullet drawn by a form: bp rects) */
  keep?: Array<{ x: number; y: number; w: number; h: number }>;
}

const pathId = (key: string) => "vtxlg-" + key.replace(/[^A-Za-z0-9_-]/g, "_");
const fmt = (v: number) => (Math.round(v * 1000) / 1000).toString();

/** The paragraph's glyphs and rules at (x0, y0) on the page; `paths` the outlines known so far. */
export function overlayOf(result: LiveResult, place: ParagraphPlace, x0: number, y0: number, paths: Map<string, string>, tx = 0, ty = 0): LiveOverlay {
  const drawn = overlayAt(result, place, x0, y0, paths);
  if (!tx && !ty) return drawn;
  // the page draws the paragraph moved (its box's transform): so is the overlay, and the part of the page it covers
  return {
    ...drawn,
    top: drawn.top + ty,
    bottom: drawn.bottom + ty,
    keep: drawn.keep?.map((r) => ({ ...r, x: r.x + tx, y: r.y + ty })),
    svg: `<g transform="translate(${fmt(tx)} ${fmt(ty)})">${drawn.svg}</g>`,
  };
}

function overlayAt(result: LiveResult, place: ParagraphPlace, x0: number, y0: number, paths: Map<string, string>): LiveOverlay {
  const last = place.boxes[place.boxes.length - 1];
  const newLast = y0 + result.lines[result.lines.length - 1][0];
  // forms drawn by reference (a list's bullet) are not in the answer: the page shows them where they are, up to the
  // first glyph after them on their line
  const keep = (result.marks ?? [])
    .filter((m) => m[0].startsWith("xobj"))
    .map(([, mx, my]) => {
      const line = result.lines.reduce((best, l) => (Math.abs(l[0] - my) < Math.abs(best[0] - my) ? l : best), result.lines[0]);
      const next = (result.map ?? []).filter((g) => Math.abs(g[1] - line[0]) < 0.5 && g[0] > mx + 0.2).reduce((m, g) => Math.min(m, g[0]), mx + 14);
      return { x: x0 + mx - 0.5, y: y0 + line[0] - line[1] - 0.5, w: Math.max(1, Math.min(next, mx + 30) - mx), h: line[1] + line[2] + 1 };
    });
  const frame = { page: place.page, top: place.boxes[0].v - place.boxes[0].height, bottom: last.v + last.depth, shift: newLast - last.v, keep };
  if (result.svg) {
    // the engine's own page of the paragraph (dvisvgm: its ids and classes are its own), moved to the paragraph
    const open = result.svg.indexOf(">", result.svg.indexOf("<svg"));
    const close = result.svg.lastIndexOf("</svg>");
    const inner = open >= 0 && close > open ? result.svg.slice(open + 1, close) : "";
    return { ...frame, svg: `<g transform="translate(${fmt(x0)} ${fmt(y0)})">${inner}</g>` };
  }
  const used = new Set<string>();
  let uses = "";
  const fill = (color: number | undefined) => (color && result.colors?.[color] ? ` fill="${result.colors[color]}"` : "");
  for (const [fi, gx, gy, code, color] of result.glyphs) {
    const f = result.fonts[fi];
    if (!f?.key) continue;
    const key = `${f.key}:${code}`;
    const d = paths.get(key);
    if (!d) continue;
    used.add(key);
    const a = f.sx * f.extend;
    const c = f.sx * f.slant;
    uses += `<use href="#${pathId(key)}" transform="matrix(${fmt(a)} 0 ${fmt(c)} ${fmt(-f.sy)} ${fmt(x0 + gx)} ${fmt(y0 + gy)})"${fill(color)}/>`;
  }
  let defs = "";
  for (const key of used) defs += `<path id="${pathId(key)}" d="${paths.get(key)}"/>`;
  let rules = "";
  for (const [rx, ry, rw, rh, color] of result.rules) rules += `<rect x="${fmt(x0 + rx)}" y="${fmt(y0 + ry)}" width="${fmt(rw)}" height="${fmt(rh)}"${fill(color)}/>`;
  return { ...frame, svg: `<defs>${defs}</defs><g fill="#000">${uses}${rules}</g>` };
}
