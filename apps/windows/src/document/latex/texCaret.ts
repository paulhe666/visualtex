// The caret on a LaTeX document's pages (personal edition): a paragraph's source cut into the units the caret
// knows (characters, spaces, formulas, commands that print something else, markup that prints nothing), matched to
// the glyphs TeX set for it (the live server's map of them, or the page's own SVG), so that a click goes to the
// source offset between the two glyphs clicked between, and every source offset has its place on the page.

/** What a stretch of the source is to the caret. */
export type UnitKind =
  /** one character shown (an escape \%, a TeX ligature -- `` as the one character it makes) */
  | "char"
  /** spaces, a line end, ~, \  and the like: the gap between two words */
  | "space"
  /** a blank line: a new paragraph */
  | "par"
  /** \\ and \newline: a line break */
  | "break"
  /** a formula ($…$, \(…\), \[…\], $$…$$, a display environment) */
  | "math"
  /** a command that prints something other than its source (\ref, \cite, \item, the document's own macros) */
  | "atom"
  /** what prints nothing: \textbf{ and its }, braces, comments, declarations, \label{…} */
  | "mark";

export interface SourceUnit {
  kind: UnitKind;
  /** offsets in the file's text */
  from: number;
  to: number;
  /** char: what it shows, normalized (see norm) */
  text?: string;
  /** math: shown as a display; the formula's own LaTeX (inside its delimiters) */
  display?: boolean;
  inner?: [number, number];
  /** atom, mark: the command */
  name?: string;
}

/** A glyph TeX set, as the caret sees it. */
export interface TypesetItem {
  /** the one character it stands for, normalized ("" unknown) */
  text: string;
  /** page bp: left edge, baseline, advance */
  x: number;
  y: number;
  adv: number;
  /** index of its line */
  line: number;
  /** in a formula (true), in a word (false), not known (null: a TFM font's character of a pdfTeX document) */
  math: boolean | null;
  /** its font's size (bp) */
  size: number;
}

export interface LayoutLine {
  /** page bp */
  y: number;
  top: number;
  bottom: number;
  left: number;
  right: number;
}

/** Where the caret can be: a source offset, its line and x (page bp). */
export interface CaretStop {
  offset: number;
  line: number;
  x: number;
}

export interface CaretLayout {
  page: number;
  /** the source range the layout covers */
  from: number;
  to: number;
  units: SourceUnit[];
  items: TypesetItem[];
  /** for each unit, the items it took: [first, end) */
  spans: Array<[number, number]>;
  lines: LayoutLine[];
  stops: CaretStop[];
}

// ---- the source ----

/** The text a character is compared by: compatibility forms folded (⾔ 言, ， ,, ﬁ fi), quotes and dashes as one. */
export function norm(s: string): string {
  return s
    .normalize("NFKC")
    .replace(/---/g, "—")
    .replace(/--/g, "–")
    .replace(/[‘’‚‛`'′]/g, "'")
    .replace(/[“”„‟"]|''|``/g, '"')
    .replace(/[‐‑]/g, "-")
    .replace(/\s+/g, "");
}

/** Commands whose arguments are shown text ("t"), or not shown ("m", "o" optional): \textbf{…}, \textcolor{c}{…}. */
const TEXT_ARGS: Record<string, string> = {
  textbf: "t", textit: "t", emph: "t", textrm: "t", textsf: "t", texttt: "t", textsc: "t", textup: "t", textsl: "t",
  textmd: "t", textnormal: "t", underline: "t", uline: "t", uuline: "t", uwave: "t", sout: "t", xout: "t", mbox: "t",
  text: "t", textsuperscript: "t", textsubscript: "t", hl: "t", fbox: "t", ul: "t", st: "t",
  textcolor: "omt", colorbox: "omt", fcolorbox: "ommt", href: "omt", hyperlink: "mt", hypertarget: "mt",
  foreignlanguage: "omt", makebox: "oot", framebox: "oot", raisebox: "moot",
  chapter: "ot", section: "ot", subsection: "ot", subsubsection: "ot", part: "ot", paragraph: "ot",
  subparagraph: "ot", caption: "ot", title: "ot", author: "t", date: "t", textbfit: "t", CJKunderline: "t",
};
/** Commands that print nothing (with their arguments). */
const MARKS: Record<string, string> = {
  label: "m", index: "m", vspace: "m", nopagebreak: "o", pagebreak: "o", phantomsection: "", noindent: "", indent: "",
  par: "", relax: "", protect: "", nobreak: "", allowbreak: "", leavevmode: "", smallskip: "", medskip: "", bigskip: "",
  bfseries: "", itshape: "", em: "", rmfamily: "", sffamily: "", ttfamily: "", upshape: "", slshape: "", scshape: "",
  mdseries: "", normalfont: "", tiny: "", scriptsize: "", footnotesize: "", small: "", normalsize: "", large: "",
  Large: "", LARGE: "", huge: "", Huge: "", selectfont: "", centering: "", raggedright: "", raggedleft: "",
  color: "om", addcontentsline: "mmm", markboth: "mm", markright: "m", hypersetup: "m", setlength: "mm",
  bf: "", it: "", rm: "", sf: "", tt: "", sc: "", sl: "", ignorespaces: "", unskip: "", null: "",
};
/** Commands that make space. */
const SPACES: Record<string, string> = {
  quad: "", qquad: "", enspace: "", enskip: "", thinspace: "", negthinspace: "", hfill: "", hfil: "", hss: "",
  hspace: "m", kern: "", hskip: "", space: "", nobreakspace: "",
};
/** Commands that print one character. */
const CHARS: Record<string, string> = {
  textbackslash: "\\", textasciitilde: "~", textasciicircum: "^", textunderscore: "_", textbar: "|",
  textless: "<", textgreater: ">", textbraceleft: "{", textbraceright: "}", textendash: "–",
  textemdash: "—", textbullet: "•", textperiodcentered: "·", textquoteleft: "'",
  textquoteright: "'", textquotedblleft: '"', textquotedblright: '"', S: "§", P: "¶", dag: "†",
  ddag: "‡", copyright: "©", textregistered: "®", texttrademark: "™", textdegree: "°",
  ss: "ß", ae: "æ", AE: "Æ", oe: "œ", OE: "Œ", o: "ø", O: "Ø", aa: "å",
  AA: "Å", l: "ł", L: "Ł", i: "ı", j: "ȷ",
};
const BREAKS = new Set(["newline", "linebreak", "\\"]);
const MATH_ENVS = /^(?:equation|align|alignat|gather|multline|flalign|eqnarray|displaymath|math|split|dmath|dgroup)\*?$/;
const ACCENTS: Record<string, string> = {
  "'": "́", "`": "̀", "^": "̂", '"': "̈", "~": "̃", "=": "̄", ".": "̇",
  u: "̆", v: "̌", H: "̋", c: "̧", d: "̣", b: "̱", r: "̊", k: "̨",
};

/** The end of the group opening at `open` (the offset after its closing brace); comments and escapes skipped. */
export function groupEnd(text: string, open: number, limit = text.length): number {
  let depth = 0;
  for (let i = open; i < limit; i++) {
    const c = text[i];
    if (c === "\\") i++;
    else if (c === "%") {
      const nl = text.indexOf("\n", i);
      i = nl < 0 || nl >= limit ? limit : nl;
    } else if (c === "{") depth++;
    else if (c === "}" && --depth === 0) return i + 1;
  }
  return limit;
}

/** An optional argument [..] starting at `at` (spaces before it allowed): its end, else -1. */
function optionalEnd(text: string, at: number, limit: number): number {
  let i = at;
  while (i < limit && (text[i] === " " || text[i] === "\t")) i++;
  if (text[i] !== "[") return -1;
  let depth = 0;
  for (; i < limit; i++) {
    const c = text[i];
    if (c === "\\") i++;
    else if (c === "{") depth++;
    else if (c === "}") depth--;
    else if (c === "]" && depth === 0) return i + 1;
  }
  return -1;
}

/** A mandatory argument starting at `at` (spaces, one line end before it allowed): [start, end) of the group or
 *  the one token, else null. */
function mandatory(text: string, at: number, limit: number): [number, number] | null {
  let i = at;
  let lines = 0;
  while (i < limit && /\s/.test(text[i])) {
    if (text[i] === "\n" && ++lines > 1) return null;
    i++;
  }
  if (i >= limit) return null;
  if (text[i] === "{") return [i, groupEnd(text, i, limit)];
  if (text[i] === "\\") {
    const m = /^\\(?:[A-Za-z@]+|.)/.exec(text.slice(i, Math.min(limit, i + 64)));
    return m ? [i, i + m[0].length] : null;
  }
  return [i, i + 1];
}

/** The end of the formula starting at `at` with `open` (its closing delimiter's end), or -1. */
function mathEnd(text: string, at: number, open: string, limit: number): number {
  const close = open === "$" ? "$" : open === "$$" ? "$$" : open === "\\(" ? "\\)" : "\\]";
  let depth = 0;
  for (let i = at + open.length; i < limit; i++) {
    const c = text[i];
    if (c === "%") {
      const nl = text.indexOf("\n", i);
      i = nl < 0 ? limit : nl;
      continue;
    }
    if (c === "{") depth++;
    else if (c === "}") depth--;
    if (depth > 0) {
      if (c === "\\") i++;
      continue;
    }
    if (text.startsWith(close, i)) {
      // ($ ends at $, not at the first of a $$ when the formula is $…$: TeX reads $$ inside $…$ as end and start)
      return i + close.length;
    }
    if (c === "\\") i++;
  }
  return -1;
}

/**
 * The units of text[from, to), in order, covering all of it. `macros`: the document's own commands and their
 * arguments (as TEXT_ARGS says them: an argument is taken for shown text; one without arguments prints
 * something: an atom).
 */
export function tokenize(text: string, from: number, to: number, macros: Map<string, string> = new Map()): SourceUnit[] {
  const out: SourceUnit[] = [];
  const push = (u: SourceUnit) => {
    if (u.to > u.from) out.push(u);
  };
  // the argument list of a command (spec letters): shown ones go through `inside`, the rest are marks
  const args = (spec: string, at: number, inside: (a: number, b: number) => void): number => {
    let i = at;
    for (const kind of spec) {
      if (kind === "o") {
        const e = optionalEnd(text, i, to);
        if (e > 0) {
          push({ kind: "mark", from: i, to: e });
          i = e;
        }
        continue;
      }
      const m = mandatory(text, i, to);
      if (!m) break;
      if (kind === "t" && text[m[0]] === "{") {
        push({ kind: "mark", from: i, to: m[0] + 1 });
        inside(m[0] + 1, m[1] - 1);
        push({ kind: "mark", from: m[1] - 1, to: m[1] });
      } else if (kind === "t") {
        push({ kind: "mark", from: i, to: m[0] });
        inside(m[0], m[1]);
      } else push({ kind: "mark", from: i, to: m[1] });
      i = m[1];
    }
    return i;
  };
  const scan = (a: number, b: number) => {
    let i = a;
    while (i < b) {
      const c = text[i];
      // comments: the line end goes with them (TeX drops it)
      if (c === "%") {
        const nl = text.indexOf("\n", i);
        const e = nl < 0 || nl >= b ? b : nl + 1;
        push({ kind: "mark", from: i, to: e });
        i = e;
        continue;
      }
      if (c === " " || c === "\t" || c === "\n" || c === "\r" || c === "~") {
        let e = i;
        let newlines = 0;
        while (e < b && /[ \t\r\n~]/.test(text[e])) {
          if (text[e] === "\n") newlines++;
          e++;
        }
        push({ kind: newlines >= 2 ? "par" : "space", from: i, to: e });
        i = e;
        continue;
      }
      if (c === "{" || c === "}") {
        push({ kind: "mark", from: i, to: i + 1 });
        i++;
        continue;
      }
      if (c === "$") {
        const open = text.startsWith("$$", i) ? "$$" : "$";
        const e = mathEnd(text, i, open, b);
        if (e < 0) {
          push({ kind: "atom", from: i, to: b, name: "$" });
          i = b;
          continue;
        }
        push({ kind: "math", from: i, to: e, display: open === "$$", inner: [i + open.length, e - open.length] });
        i = e;
        continue;
      }
      if (c === "-" && text.startsWith("--", i)) {
        const n = text.startsWith("---", i) ? 3 : 2;
        push({ kind: "char", from: i, to: i + n, text: n === 3 ? "—" : "–" });
        i += n;
        continue;
      }
      if ((c === "`" || c === "'") && text[i + 1] === c) {
        push({ kind: "char", from: i, to: i + 2, text: '"' });
        i += 2;
        continue;
      }
      if (c === "\\") {
        const m = /^\\(?:([A-Za-z@]+)\*?|(.))/s.exec(text.slice(i, Math.min(b, i + 80)));
        if (!m) {
          push({ kind: "mark", from: i, to: i + 1 });
          i++;
          continue;
        }
        const name = m[1] ?? m[2];
        let e = i + m[0].length;
        if (m[2] !== undefined) {
          // control symbols
          if (name === "(" || name === "[") {
            const end = mathEnd(text, i, "\\" + name, b);
            if (end < 0) {
              push({ kind: "atom", from: i, to: b, name });
              i = b;
            } else {
              push({ kind: "math", from: i, to: end, display: name === "[", inner: [i + 2, end - 2] });
              i = end;
            }
            continue;
          }
          if (name === "\\") {
            // \\, \\*, \\[len]
            if (text[e] === "*") e++;
            const o = optionalEnd(text, e, b);
            push({ kind: "break", from: i, to: o > 0 ? o : e, name });
            i = o > 0 ? o : e;
            continue;
          }
          if ("%&#$_{}".includes(name)) {
            push({ kind: "char", from: i, to: e, text: name });
            i = e;
            continue;
          }
          if (ACCENTS[name] !== undefined) {
            const arg = mandatory(text, e, b);
            if (arg) {
              const base = text.slice(arg[0], arg[1]).replace(/[{}]/g, "");
              push({ kind: "char", from: i, to: arg[1], text: norm((base + ACCENTS[name]).normalize("NFC")) });
              i = arg[1];
              continue;
            }
          }
          // \  \, \; \: \! \/ \- \@
          push({ kind: name === " " || name === "," || name === ";" || name === ":" || name === "\n" ? "space" : "mark", from: i, to: e, name });
          i = e;
          continue;
        }
        if (name === "begin") {
          const env = /^\s*\{([^{}]*)\}/.exec(text.slice(e, Math.min(b, e + 80)));
          const envName = env?.[1].trim() ?? "";
          const close = text.indexOf(`\\end{${envName}}`, e);
          const end = close < 0 || close >= b ? b : close + `\\end{${envName}}`.length;
          if (MATH_ENVS.test(envName)) {
            const inner0 = e + (env ? env[0].length : 0);
            push({ kind: "math", from: i, to: end, display: envName !== "math", inner: [inner0, close < 0 ? end : close] });
          } else push({ kind: "atom", from: i, to: end, name: envName });
          i = end;
          continue;
        }
        if (BREAKS.has(name)) {
          const o = optionalEnd(text, e, b);
          push({ kind: "break", from: i, to: o > 0 ? o : e, name });
          i = o > 0 ? o : e;
          continue;
        }
        if (TEXT_ARGS[name] !== undefined) {
          push({ kind: "mark", from: i, to: e, name });
          i = args(TEXT_ARGS[name], e, (x, y) => scan(x, y));
          continue;
        }
        if (CHARS[name] !== undefined) {
          // (a control word eats the spaces after it; {} after it is part of it)
          let end = e;
          if (text.startsWith("{}", end)) end += 2;
          push({ kind: "char", from: i, to: end, text: norm(CHARS[name]) });
          i = end;
          continue;
        }
        if (SPACES[name] !== undefined) {
          let end = e;
          if (SPACES[name]) {
            if (text[end] === "*") end++;
            const m2 = mandatory(text, end, b);
            if (m2) end = m2[1];
          }
          push({ kind: "space", from: i, to: end, name });
          i = end;
          continue;
        }
        if (MARKS[name] !== undefined) {
          const end = args(MARKS[name], e, (x, y) => push({ kind: "mark", from: x, to: y }));
          push({ kind: "mark", from: i, to: Math.max(end, e), name });
          // (args pushed their own marks: keep the order by sorting at the end)
          i = Math.max(end, e);
          continue;
        }
        const own = macros.get(name);
        if (own) {
          // the document's own command with arguments: taken for text it shows
          push({ kind: "mark", from: i, to: e, name });
          i = args(own, e, (x, y) => scan(x, y));
          continue;
        }
        // something it prints: \ref{…}, \cite[…]{…}, \item[…], \LaTeX, the document's own without arguments (an
        // unknown command: with the groups right after it)
        let end = e;
        if (own === undefined) {
          for (;;) {
            if (text[end] === "[") {
              const o = optionalEnd(text, end, b);
              if (o < 0) break;
              end = o;
            } else if (text[end] === "{") end = groupEnd(text, end, b);
            else break;
          }
        }
        push({ kind: "atom", from: i, to: end, name });
        i = end;
        continue;
      }
      // a character (a surrogate pair is one)
      const cp = text.codePointAt(i) ?? 0;
      const len = cp > 0xffff ? 2 : 1;
      const shown = norm(text.slice(i, i + len));
      push({ kind: "char", from: i, to: i + len, text: shown });
      i += len;
    }
  };
  scan(from, to);
  out.sort((x, y) => x.from - y.from || x.to - y.to);
  // (marks pushed inside a mark's arguments come first: one unit for the whole)
  const merged: SourceUnit[] = [];
  for (const u of out) {
    const last = merged[merged.length - 1];
    if (last && u.from < last.to) {
      if (u.to > last.to) last.to = u.to;
      continue;
    }
    merged.push(u);
  }
  return merged;
}

/** The document's own commands and their arguments as tokenize takes them ("" none: it prints something; "ot":
 *  an optional argument, then one taken for text): \newcommand{\x}[2], \NewDocumentCommand, \def\x#1#2. */
export function documentCommands(preamble: string): Map<string, string> {
  const out = new Map<string, string>();
  const code = preamble.replace(/(^|[^\\])%[^\n]*/g, "$1");
  for (const m of code.matchAll(/\\(?:newcommand|renewcommand|providecommand|DeclareRobustCommand)\*?\s*\{?\s*\\([A-Za-z@]+)\s*\}?\s*(?:\[(\d)\])?(\s*\[)?/g)) {
    const n = Number(m[2] ?? 0);
    out.set(m[1], n === 0 ? "" : m[3] ? "o" + "t".repeat(n - 1) : "t".repeat(n));
  }
  for (const m of code.matchAll(/\\(?:NewDocumentCommand|RenewDocumentCommand|DeclareDocumentCommand|ProvideDocumentCommand)\s*\{?\s*\\([A-Za-z@]+)\s*\}?\s*\{([^{}]*)\}/g))
    out.set(m[1], [...m[2].replace(/\{[^{}]*\}/g, "")].filter((c) => /[mrRvbOos]/.test(c)).map((c) => (/[Oo]/.test(c) ? "o" : c === "s" ? "" : "t")).join(""));
  for (const m of code.matchAll(/\\[gex]?def\s*\\([A-Za-z@]+)((?:#\d)*)/g)) out.set(m[1], "t".repeat(m[2].length / 2));
  return out;
}

// ---- matching the source to the glyphs ----

const BIG = 1e9;

/** What it costs to leave a glyph unmatched (one TeX made that is not a character of the source: a list's label, a
 *  heading's number), to leave a source character without a glyph, for a formula or command to take a glyph. */
function skipItem(it: TypesetItem): number {
  if (!it.text) return 0.5;
  return it.math === true ? 0.6 : it.math === null ? 0.8 : 1;
}
function mathTakes(it: TypesetItem): number {
  if (!it.text) return 0;
  // (a word's glyph costs more taken than left unmatched: text never becomes part of a formula's box)
  return it.math === true ? 0 : it.math === null ? 0.1 : 1.6;
}
const ATOM_TAKES = 0.2;
const SKIP_CHAR = 1;

/**
 * Which glyphs each unit made: the alignment of the units' characters with the glyphs' (in the order TeX set them),
 * formulas and printing commands taking the glyphs between. Returns per unit [first, end) of `items`.
 */
export function align(units: SourceUnit[], items: TypesetItem[]): Array<[number, number]> {
  const n = units.length;
  const m = items.length;
  const W = m + 1;
  if ((n + 1) * W > 6e6) return alignGreedy(units, items);
  const cost = new Float32Array((n + 1) * W).fill(BIG);
  // how a cell was reached: 1 char matched, 2 char skipped, 3 item skipped, 4 a unit that takes nothing, 5 out of a
  // formula/command (its cells in `inside`)
  const how = new Uint8Array((n + 1) * W);
  const inside = new Float32Array(W);
  const insideHow = new Uint8Array((n + 1) * W); // 1 opened here, 2 took item j-1
  cost[0] = 0;
  for (let i = 0; i <= n; i++) {
    const row = i * W;
    const u = units[i];
    const taking = u && (u.kind === "math" || u.kind === "atom");
    if (taking) inside.fill(BIG);
    for (let j = 0; j <= m; j++) {
      // an item skipped
      if (j > 0) {
        const c = cost[row + j - 1] + skipItem(items[j - 1]);
        if (c < cost[row + j]) {
          cost[row + j] = c;
          how[row + j] = 3;
        }
      }
      const here = cost[row + j];
      if (!u) continue;
      const next = row + W;
      if (taking) {
        // in the formula: opened at this item, or taking item j-1
        let v = here;
        let h = 1;
        if (j > 0) {
          const t = inside[j - 1] + (u.kind === "math" ? mathTakes(items[j - 1]) : ATOM_TAKES);
          if (t < v) {
            v = t;
            h = 2;
          }
        }
        inside[j] = v;
        insideHow[row + j] = h;
        if (v < cost[next + j]) {
          cost[next + j] = v;
          how[next + j] = 5;
        }
        continue;
      }
      if (here >= BIG) continue;
      if (u.kind === "char") {
        if (j < m) {
          const c = here + (u.text && u.text === items[j].text ? 0 : 3);
          if (c < cost[next + j + 1]) {
            cost[next + j + 1] = c;
            how[next + j + 1] = 1;
          }
        }
        const c = here + SKIP_CHAR;
        if (c < cost[next + j]) {
          cost[next + j] = c;
          how[next + j] = 2;
        }
      } else if (here < cost[next + j]) {
        cost[next + j] = here;
        how[next + j] = 4;
      }
    }
  }
  // back from the end
  const spans: Array<[number, number]> = new Array(n);
  let i = n;
  let j = m;
  while (i > 0 || j > 0) {
    const h = how[i * W + j];
    if (h === 3) {
      j--;
      continue;
    }
    if (i === 0) {
      j--;
      continue;
    }
    if (h === 1) {
      spans[i - 1] = [j - 1, j];
      i--;
      j--;
    } else if (h === 2 || h === 4) {
      spans[i - 1] = [j, j];
      i--;
    } else if (h === 5) {
      const end = j;
      let k = j;
      const row = (i - 1) * W;
      while (insideHow[row + k] === 2 && k > 0) k--;
      spans[i - 1] = [k, end];
      i--;
      j = k;
    } else {
      // (unreachable: every reached cell says how)
      spans[i - 1] = [j, j];
      i--;
    }
  }
  return spans;
}

/** For very long paragraphs: characters matched in order where they agree, formulas taking what is between. */
function alignGreedy(units: SourceUnit[], items: TypesetItem[]): Array<[number, number]> {
  const spans: Array<[number, number]> = [];
  let j = 0;
  for (let i = 0; i < units.length; i++) {
    const u = units[i];
    if (u.kind === "char") {
      let k = j;
      while (k < items.length && k < j + 8 && items[k].text !== u.text) k++;
      if (k < items.length && items[k].text === u.text) {
        spans.push([k, k + 1]);
        j = k + 1;
      } else spans.push([j, j]);
    } else if (u.kind === "math" || u.kind === "atom") {
      // up to the next character's glyph
      let next = i + 1;
      while (next < units.length && units[next].kind !== "char") next++;
      const want = units[next]?.text;
      let k = j;
      while (k < items.length && (want === undefined || items[k].text !== want || items[k].math === true)) k++;
      spans.push([j, k]);
      j = k;
    } else spans.push([j, j]);
  }
  return spans;
}

// ---- the caret's places ----

/** The line an item is on: the one whose baseline it sits on, else the one whose height and depth hold it. */
export function lineOf(lines: LayoutLine[], y: number): number {
  let best = 0;
  let dist = Infinity;
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];
    const d = y < l.top ? l.top - y : y > l.bottom ? y - l.bottom : Math.abs(y - l.y) * 0.01;
    if (d < dist) {
      dist = d;
      best = i;
    }
  }
  return best;
}

interface UnitBox {
  firstLine: number;
  lastLine: number;
  left: number;
  right: number;
}

function boxOf(items: TypesetItem[], span: [number, number]): UnitBox | null {
  const [a, b] = span;
  if (b <= a) return null;
  const firstLine = items[a].line;
  const lastLine = items[b - 1].line;
  let left = Infinity;
  let right = -Infinity;
  for (let k = a; k < b; k++) {
    const it = items[k];
    if (it.line === firstLine) left = Math.min(left, it.x);
    if (it.line === lastLine) right = Math.max(right, it.x + it.adv);
  }
  return { firstLine, lastLine, left, right };
}

/** The caret's places in order: before and after each unit that shows something (a space's place is where the
 *  next one starts — on the next line when the line breaks there). Text typed at a place goes where the
 *  character before it is (inside its \textbf{…}); after a space, before what follows the space. */
function stopsOf(units: SourceUnit[], items: TypesetItem[], spans: Array<[number, number]>): CaretStop[] {
  const stops: CaretStop[] = [];
  /** the last unit before the one at hand that is not a mark: what the caret's offset follows */
  let before: SourceUnit | null = null;
  let lastAfter: CaretStop | null = null;
  for (let i = 0; i < units.length; i++) {
    const u = units[i];
    const box = boxOf(items, spans[i]);
    if (box) {
      let offset = u.from;
      if (before && (before.kind === "space" || before.kind === "par" || before.kind === "break")) offset = before.to;
      else if (before) offset = before.to;
      const start = { offset, line: box.firstLine, x: box.left };
      const same = lastAfter && lastAfter.offset === before?.to && before?.kind !== "space" && before?.kind !== "par" && before?.kind !== "break";
      if (same && lastAfter && lastAfter.line === start.line && Math.abs(lastAfter.x - start.x) < 0.5) {
        // the place between two characters: one stop (the first's offset, at the second's start — a kern between
        // them is before the caret, as browsers draw it)
        lastAfter.x = start.x;
      } else stops.push(start);
      lastAfter = { offset: u.to, line: box.lastLine, x: box.right };
      stops.push(lastAfter);
    }
    if (u.kind !== "mark") before = u;
  }
  return stops;
}

/** The layout of a unit range: its units matched to TeX's glyphs, its lines and the caret's places. */
export function layoutOf(
  page: number,
  text: string,
  from: number,
  to: number,
  items: TypesetItem[],
  lines: LayoutLine[],
  macros?: Map<string, string>,
): CaretLayout {
  const units = tokenize(text, from, to, macros);
  const spans = align(units, items);
  return { page, from, to, units, items, spans, lines, stops: stopsOf(units, items, spans) };
}

// ---- items from a live answer or from the page ----

/** Glyphs with several characters (a ligature: ffi) as one item a character, its advance shared. */
function pushItem(out: TypesetItem[], text: string, x: number, y: number, adv: number, line: number, math: boolean | null, size: number) {
  const chars = [...norm(text)];
  if (chars.length <= 1) {
    out.push({ text: chars[0] ?? "", x, y, adv, line, math, size });
    return;
  }
  const w = adv / chars.length;
  chars.forEach((c, k) => out.push({ text: c, x: x + k * w, y, adv: w, line, math, size }));
}

/** The items and lines of a live answer drawn with its box's top left at (x0, y0) on the page. */
export function liveItems(
  map: Array<[number, number, number, string, number, number?]>,
  resultLines: Array<[number, number, number, number, number]>,
  x0: number,
  y0: number,
): { items: TypesetItem[]; lines: LayoutLine[] } {
  const lines: LayoutLine[] = resultLines.map(([baseline, h, d, w, shift]) => ({
    y: y0 + baseline,
    top: y0 + baseline - h,
    bottom: y0 + baseline + d,
    left: x0 + shift,
    right: x0 + shift + w,
  }));
  const words = map.some((g) => g[4] === 1);
  const items: TypesetItem[] = [];
  for (const [x, y, adv, text, kind, size] of map) {
    const math = kind === 3 || kind === 5 ? true : kind === 1 || kind === 4 ? false : words ? true : null;
    pushItem(items, text, x0 + x, y0 + y, adv, lineOf(lines, y0 + y), math, size ?? 10);
  }
  return { items, lines };
}

/** A glyph of a page's SVG (dvisvgm: one <text> a glyph): page bp, its character, its font size. */
export interface PageChar {
  /** where TeX put it (SyncTeX's coordinates), bp */
  x: number;
  y: number;
  /** how far the page draws it from there (a box's transform) */
  tx?: number;
  ty?: number;
  text: string;
  size: number;
  /** its font is a formula's (true), a text font (false), not known (null) */
  math?: boolean | null;
}

const CJK = /[⺀-鿿豈-﫿＀-￯　-〿]/;

/** The items of the page's glyphs in a paragraph's lines (`lines` in TeX's coordinates; each glyph's advance: to
 *  the next on its line, else from its size), placed where the page draws them; `tx`, `ty`: how far the page draws
 *  them from TeX's coordinates (the lines are to be moved as much). */
export function pageItems(chars: PageChar[], lines: LayoutLine[]): { items: TypesetItem[]; tx: number; ty: number } {
  const inLines = chars.filter((c) => lines.some((l) => c.y >= l.top - 0.5 && c.y <= l.bottom + 0.5 && c.x >= l.left - 2 && c.x <= l.right + 2));
  const moves = new Map<string, { tx: number; ty: number; n: number }>();
  for (const c of inLines) {
    const k = `${(c.tx ?? 0).toFixed(2)},${(c.ty ?? 0).toFixed(2)}`;
    const e = moves.get(k);
    if (e) e.n++;
    else moves.set(k, { tx: c.tx ?? 0, ty: c.ty ?? 0, n: 1 });
  }
  const move = [...moves.values()].sort((a, b) => b.n - a.n)[0] ?? { tx: 0, ty: 0 };
  const items: TypesetItem[] = [];
  for (let k = 0; k < inLines.length; k++) {
    const c = inLines[k];
    const next = inLines[k + 1];
    const guess = c.size * (CJK.test(c.text) ? 1 : 0.5);
    const adv = next && Math.abs(next.y - c.y) < 0.1 && next.x > c.x && next.x - c.x <= guess * 1.6 ? next.x - c.x : guess;
    pushItem(items, c.text, c.x + move.tx, c.y + move.ty, adv, lineOf(lines, c.y), c.math ?? null, c.size);
  }
  return { items, tx: move.tx, ty: move.ty };
}

/** Lines moved by (tx, ty). */
export function movedLines(lines: LayoutLine[], tx: number, ty: number): LayoutLine[] {
  if (!tx && !ty) return lines;
  return lines.map((l) => ({ y: l.y + ty, top: l.top + ty, bottom: l.bottom + ty, left: l.left + tx, right: l.right + tx }));
}

// ---- the caret ----

/** The stop of a source offset (`line`: the one on that line when the offset has two, a line's end and the next's
 *  start; an offset between stops: the next one; past the end: the last). */
export function stopOf(layout: CaretLayout, offset: number, line?: number): CaretStop | null {
  const s = layout.stops;
  if (!s.length) return null;
  let found: CaretStop | null = null;
  let after: CaretStop | null = null;
  for (const stop of s) {
    if (stop.offset === offset) {
      if (line === undefined || stop.line === line) return stop;
      found ??= stop;
    }
    if (stop.offset > offset && !after) after = stop;
  }
  return found ?? after ?? s[s.length - 1];
}

/** The offset a page point goes to: the place nearest it on the line it is on (between the glyphs it is between). */
export function offsetAt(layout: CaretLayout, x: number, y: number): number | null {
  if (!layout.stops.length || !layout.lines.length) return null;
  const line = lineAt(layout, y);
  let best: CaretStop | null = null;
  for (const s of layout.stops) {
    if (s.line !== line) continue;
    if (!best || Math.abs(s.x - x) < Math.abs(best.x - x)) best = s;
  }
  return best?.offset ?? null;
}

/** The line a page y is on (the nearest when between lines). */
export function lineAt(layout: CaretLayout, y: number): number {
  let best = 0;
  let dist = Infinity;
  layout.lines.forEach((l, i) => {
    const d = y < l.top ? l.top - y : y > l.bottom ? y - l.bottom : 0;
    if (d < dist || (d === dist && Math.abs(y - l.y) < Math.abs(y - layout.lines[best].y))) {
      dist = d;
      best = i;
    }
  });
  return best;
}

/** The unit a page point is on when it is a formula or a printing command (the box of its glyphs), else null. */
export function atomAt(layout: CaretLayout, x: number, y: number): number | null {
  for (let i = 0; i < layout.units.length; i++) {
    const u = layout.units[i];
    if (u.kind !== "math" && u.kind !== "atom") continue;
    const [a, b] = layout.spans[i];
    for (let k = a; k < b; k++) {
      const it = layout.items[k];
      const line = layout.lines[it.line];
      if (x >= it.x - 0.5 && x <= it.x + it.adv + 0.5 && y >= line.top - 1 && y <= line.bottom + 1) return i;
    }
  }
  return null;
}

/** The stop before/after the caret's (left/right arrow); null past the layout's first/last. */
export function stepStop(layout: CaretLayout, offset: number, line: number | undefined, dir: -1 | 1): CaretStop | null {
  const s = layout.stops;
  const exact = s.findIndex((stop) => stop.offset === offset && (line === undefined || stop.line === line));
  let here = exact >= 0 ? exact : s.findIndex((stop) => stop.offset === offset);
  if (here < 0) {
    // between stops: the one after it is the first step right, the one before it the first step left
    const next = s.findIndex((stop) => stop.offset > offset);
    if (next < 0) return dir < 0 ? s[s.length - 1] ?? null : null;
    return dir > 0 ? s[next] : next > 0 ? s[next - 1] : null;
  }
  for (let k = here + dir; k >= 0 && k < s.length; k += dir) {
    // (a stop at the same place: no step)
    if (s[k].line !== s[here].line || Math.abs(s[k].x - s[here].x) > 0.01 || s[k].offset !== s[here].offset) return s[k];
    here = k;
  }
  return null;
}

/** The place on the line above/below nearest `x`, null past the first/last line. */
export function lineStop(layout: CaretLayout, line: number, x: number): CaretStop | null {
  let best: CaretStop | null = null;
  for (const s of layout.stops) {
    if (s.line !== line) continue;
    if (!best || Math.abs(s.x - x) < Math.abs(best.x - x)) best = s;
  }
  return best;
}

/** The first and last place of a line (Home, End). */
export function lineEdges(layout: CaretLayout, line: number): [CaretStop, CaretStop] | null {
  const on = layout.stops.filter((s) => s.line === line);
  return on.length ? [on[0], on[on.length - 1]] : null;
}

/** The selection's rectangles (page bp) from offset a to b: per line, from where it starts to where it ends. */
export function selectionRects(layout: CaretLayout, a: number, b: number): Array<{ x: number; y: number; w: number; h: number }> {
  if (a === b) return [];
  const [lo, hi] = a < b ? [a, b] : [b, a];
  const out: Array<{ x: number; y: number; w: number; h: number }> = [];
  const byLine = new Map<number, [number, number]>();
  for (let i = 0; i < layout.units.length; i++) {
    const u = layout.units[i];
    if (u.to <= lo || u.from >= hi) continue;
    const box = boxOf(layout.items, layout.spans[i]);
    if (!box) continue;
    for (let line = box.firstLine; line <= box.lastLine; line++) {
      const l = layout.lines[line];
      const left = line === box.firstLine ? box.left : l.left;
      const right = line === box.lastLine ? box.right : l.right;
      const old = byLine.get(line);
      byLine.set(line, old ? [Math.min(old[0], left), Math.max(old[1], right)] : [left, right]);
    }
  }
  for (const [line, [x0, x1]] of [...byLine].sort((p, q) => p[0] - q[0])) {
    const l = layout.lines[line];
    out.push({ x: x0, y: l.top, w: Math.max(0.5, x1 - x0), h: l.bottom - l.top });
  }
  return out;
}

/** The source ranges that deleting [a, b) removes: what shows (characters, spaces, formulas, commands), not the
 *  markup around it (\textbf{ and its } stay), so the source stays balanced. */
export function deletionRanges(units: SourceUnit[], a: number, b: number): Array<[number, number]> {
  const out: Array<[number, number]> = [];
  for (const u of units) {
    if (u.kind === "mark" || u.to <= a || u.from >= b) continue;
    // a unit partly inside: whole (a formula, a command) or its inside part (spaces)
    const from = u.kind === "space" || u.kind === "par" ? Math.max(u.from, a) : u.from;
    const to = u.kind === "space" || u.kind === "par" ? Math.min(u.to, b) : u.to;
    const last = out[out.length - 1];
    if (last && last[1] === from) last[1] = to;
    else out.push([from, to]);
  }
  return out;
}

/** The empty text commands left around `at` after a deletion (\textbf{} \emph{}): their ranges. */
export function emptyGroupsAround(text: string, at: number): Array<[number, number]> {
  const out: Array<[number, number]> = [];
  const re = /\\(?:textbf|textit|emph|textrm|textsf|texttt|textsc|textup|textsl|textmd|underline|uline|sout|mbox)\{\}/g;
  const from = Math.max(0, at - 40);
  re.lastIndex = from;
  for (let m = re.exec(text); m && m.index <= at + 1; m = re.exec(text)) {
    if (m.index + m[0].length >= at - 1) out.push([m.index, m.index + m[0].length]);
  }
  return out;
}

/** What typing a character puts in the source (LaTeX's special characters escaped). */
export function escapeTyped(s: string): string {
  return s.replace(/[\\%#&_{}~^]/g, (c) =>
    c === "\\" ? "\\textbackslash{}" : c === "~" ? "\\textasciitilde{}" : c === "^" ? "\\textasciicircum{}" : "\\" + c,
  );
}
