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
/** A line of a paragraph's text: not blank, and no structure, display or command that makes it more than text
 *  (DENIED below, the document's own structural commands). Such lines (a heading, \begin{equation}, \item) end
 *  the run of lines typeset together. */
const textLine = (line: string, m: DocMacros) => !blank(line) && !DENIED.test(line) && !DISPLAY.test(line) && !m.structural?.test(line);
/** A line that starts a paragraph of its own (it begins with one of the document's paragraph commands). */
const startsParagraph = (line: string, m: DocMacros) => {
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
      "begin", "end", "item", "bibitem", "part", "chapter", "section", "subsection", "subsubsection", "paragraph", "subparagraph",
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

/** A paragraph the live server may typeset: `text` the file's text now, `main` whether it is the main file. */
export function liveEligible(change: ParagraphChange, text: string, main: boolean): boolean {
  for (const t of [change.newText, change.oldText]) {
    if (!/\S/.test(t) || DENIED.test(t) || DISPLAY.test(t) || !balanced(t)) return false;
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
 *  (-1: \indent, 0: \noindent), then the paragraph. */
export function liveText(prefix: string, indent: number, paragraph: string): string {
  return (prefix ? prefix + "\n" : "") + (indent < 0 ? "\\indent " : "\\noindent ") + paragraph;
}

/** The paragraph begins with one of the document's paragraph commands (it sets its own indent). */
export function beginsParagraphCommand(paragraph: string, m: DocMacros): boolean {
  return startsParagraph(paragraph.split("\n")[0] ?? "", m);
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
      if (!"xkg$".includes(r.kind)) return;
      const line = lineOf(i);
      if (line < 0) return;
      if (r.kind === "g") glueEnds.set(line, [...(glueEnds.get(line) ?? []), r.h]);
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
    if (boxes.every((b) => Math.abs(b.width - boxes[0].width) <= 0.01)) out.push({ page: area.page, boxes, glues: area.glues.slice(start, start + count) });
  }
  return out;
}

// ---- the page's glyphs (its SVG) and the check ----

export interface PageGlyph {
  x: number;
  y: number;
}

/** Where the page's SVG (dvisvgm: one <text> a glyph) draws each glyph, bp from the page's top left (`page`
 *  false: in the SVG's own user units, the DVI's — a live paragraph's, from its top left). */
export function pageGlyphs(svg: string, page = true): PageGlyph[] {
  const box = page ? /viewBox=['"]([-\d.e]+)[ ,]+([-\d.e]+)/.exec(svg) : null;
  const dx = box ? -Number(box[1]) : 0;
  const dy = box ? -Number(box[2]) : 0;
  const out: PageGlyph[] = [];
  const re = /<text\b([^>]*)>/g;
  for (let m = re.exec(svg); m; m = re.exec(svg)) {
    const x = /\bx=['"]([-\d.e]+)['"]/.exec(m[1]);
    const y = /\by=['"]([-\d.e]+)['"]/.exec(m[1]);
    if (x && y) out.push({ x: Number(x[1]) + dx, y: Number(y[1]) + dy });
  }
  return out;
}

const TOLERANCE = 0.05;

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
): { x0: number; y0: number } | null {
  if (result.status !== "ok" || result.undrawable > 0 || result.lines.length !== place.boxes.length || !result.lines.length) return null;
  const x0 = place.boxes[0].h;
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
    return recorded >= place.boxes.length && matched === recorded ? { x0, y0 } : null;
  }
  const last = place.boxes[place.boxes.length - 1];
  const top = place.boxes[0].v - place.boxes[0].height - 0.5;
  const bottom = last.v + last.depth + 0.5;
  const left = x0 - 1;
  const right = x0 + place.boxes[0].width + 1;
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
  const own: [number, number][] = result.svg ? pageGlyphs(result.svg, false).map((g) => [g.x, g.y]) : result.glyphs.map(([, gx, gy]) => [gx, gy]);
  for (const [gx, gy] of own) {
    const x = x0 + gx;
    const y = y0 + gy;
    const list = byLine.get(Math.round(y * 10));
    if (list?.some((g) => Math.abs(g.x - x) <= TOLERANCE && Math.abs(g.y - y) <= TOLERANCE)) matched++;
    else report?.misses.push([x, y]);
  }
  const n = own.length;
  if (report) {
    report.matched = matched;
    report.onPage = onPage;
  }
  if (matched < n - Math.floor(n * 0.01) || Math.abs(onPage - n) > Math.max(1, Math.floor(n * 0.01))) return null;
  return { x0, y0 };
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
}

const pathId = (key: string) => "vtxlg-" + key.replace(/[^A-Za-z0-9_-]/g, "_");
const fmt = (v: number) => (Math.round(v * 1000) / 1000).toString();

/** The paragraph's glyphs and rules at (x0, y0) on the page; `paths` the outlines known so far. */
export function overlayOf(result: LiveResult, place: ParagraphPlace, x0: number, y0: number, paths: Map<string, string>): LiveOverlay {
  const last = place.boxes[place.boxes.length - 1];
  const newLast = y0 + result.lines[result.lines.length - 1][0];
  const frame = { page: place.page, top: place.boxes[0].v - place.boxes[0].height, bottom: last.v + last.depth, shift: newLast - last.v };
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
