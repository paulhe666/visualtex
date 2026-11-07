// Inline Markdown read the way Typora reads it, character by character (markdownLive.ts shows a formatted span's
// Markdown while the caret is in it and reads it again after every change). Each character of the text is either a
// marker ("**", "`", "[", "](url)", "<u>", "\") or text, and the text carries the formatting of the spans around it.
//
// CommonMark's rules with Typora's leniency, as tested in Typora 1.x: a closing "**" may follow a space ("**x **" is
// bold), punctuation next to a delimiter does not stop it ("前**“引号”**后", "**中文。**后" are bold), "__" closes
// before a letter ("__b__c"), but an opening delimiter never has a space after it ("** x**" and "a * b * c" stay
// text) and "_" inside a word never opens ("snake_case_word"); "**a *b** c*" is bold "a *b" (a "**" closes the
// nearest "**", not a lone "*" in between) and a lone "*" is not closed by "**" ("*it**" stays text). "~~" strikes, "==" highlights, backticks make code, [text](url "title"),
// <url>, www.… and https://… link, <u> <sup> <sub> <kbd> <mark> <s> <b> <i> are their formatting, "\" escapes.

/** A span's kind; "escape" is "\" with the character after it. */
export type InlineKind =
  | "strong"
  | "em"
  | "strike"
  | "highlight"
  | "code"
  | "link"
  | "underline"
  | "superscript"
  | "subscript"
  | "kbd"
  | "escape";

export interface InlineSpan {
  kind: InlineKind;
  /** the span with its markers, in characters: [from, to) */
  from: number;
  to: number;
  /** its text: [contentFrom, contentTo) */
  contentFrom: number;
  contentTo: number;
  /** a link's address and title */
  href?: string;
  title?: string | null;
  /** a link written as its address (https://…, www.…): no markers */
  bare?: boolean;
}

export interface InlineInput {
  /** one character per position; an inline object (a formula, an image) is OBJECT */
  text: string;
  /** characters that are never Markdown (typed escaped, or formatted text read as one piece) */
  literal?: (index: number) => boolean;
  /** spans kept with no text between their markers (Ctrl+B's "****"; a bold run whose text was deleted) */
  empty?: ReadonlyArray<{ from: number; to: number; kind: InlineKind }>;
}

export interface InlineParse {
  /** every span, outer ones first */
  spans: InlineSpan[];
  /** 1 for a marker character, 2 for a link's address or title (shown as the address) */
  syntax: Uint8Array;
}

export const OBJECT = "￼";

/** CommonMark's escapable characters: ASCII punctuation. */
export const ESCAPABLE = "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~";

/** The markers of a span with nothing in it. */
export const EMPTY_MARKERS: Partial<Record<InlineKind, string>> = {
  strong: "****",
  em: "**",
  strike: "~~~~",
  highlight: "====",
  code: "``",
  underline: "<u></u>",
  superscript: "<sup></sup>",
  subscript: "<sub></sub>",
  kbd: "<kbd></kbd>",
};

/** HTML tags read as formatting (<b>x</b> is bold; written back as **x**). */
const TAGS: Record<string, InlineKind> = {
  u: "underline",
  ins: "underline",
  sup: "superscript",
  sub: "subscript",
  kbd: "kbd",
  mark: "highlight",
  s: "strike",
  del: "strike",
  strike: "strike",
  b: "strong",
  strong: "strong",
  i: "em",
  em: "em",
};

const isSpace = (c: string) => c === " " || c === "\t" || c === "\n" || c === "\r" || c === " " || c === "　" || c === "";
const isPunct = (c: string) => c !== OBJECT && /[\p{P}\p{S}]/u.test(c);
const isWord = (c: string) => c === OBJECT || /[\p{L}\p{N}]/u.test(c);
/** where an address stops in Chinese text (no space before the next word) */
const isCjk = (c: string) => /[⺀-鿿豈-﫿＀-￯　-〿]/.test(c);

interface Delimiter {
  /** the run's first character */
  pos: number;
  char: string;
  length: number;
  /** characters still free: [pos + used, pos + used + remaining) — openers give theirs from the end, closers from the start */
  remaining: number;
  used: number;
  canOpen: boolean;
  canClose: boolean;
  /** CommonMark's own flanking: the run could both open and close (the rule of three) */
  both: boolean;
}

export function parseInline(input: InlineInput): InlineParse {
  const s = input.text;
  const n = s.length;
  const spans: InlineSpan[] = [];
  const syntax = new Uint8Array(n);
  const literal = (i: number) => input.literal?.(i) ?? false;
  const empties = new Map<number, { from: number; to: number; kind: InlineKind }>();
  for (const e of input.empty ?? []) {
    const markers = EMPTY_MARKERS[e.kind];
    if (markers && e.to - e.from === markers.length && s.slice(e.from, e.to) === markers) empties.set(e.from, e);
  }
  const at = (i: number) => (i >= 0 && i < n ? s[i] : "");
  const mark = (from: number, to: number) => {
    for (let i = from; i < to; i++) syntax[i] = 1;
  };

  /** A code span from the backticks at `i`: its end, or -1 (the backticks are text). */
  const codeEnd = (i: number, end: number, record: boolean): number => {
    let j = i;
    while (j < end && s[j] === "`" && !literal(j)) j++;
    const ticks = j - i;
    let k = j;
    while (k < end) {
      if (s[k] !== "`" || literal(k)) {
        k++;
        continue;
      }
      let m = k;
      while (m < end && s[m] === "`" && !literal(m)) m++;
      if (m - k === ticks) {
        if (record) {
          let contentFrom = j;
          let contentTo = k;
          // one space each side is padding when both are there (not all spaces)
          if (contentTo - contentFrom >= 2 && s[contentFrom] === " " && s[contentTo - 1] === " " && s.slice(contentFrom, contentTo).trim()) {
            contentFrom++;
            contentTo--;
          }
          spans.push({ kind: "code", from: i, to: m, contentFrom, contentTo });
          mark(i, contentFrom);
          mark(contentTo, m);
        }
        return m;
      }
      k = m;
    }
    return -1;
  };

  /** "]" closing the "[" at `i` (brackets nest; code spans and escapes inside are skipped), or -1. */
  const closingBracket = (i: number, end: number): number => {
    let depth = 0;
    for (let k = i; k < end; k++) {
      if (literal(k)) continue;
      const c = s[k];
      if (c === "\\" && k + 1 < end && ESCAPABLE.includes(s[k + 1])) {
        k++;
        continue;
      }
      if (c === "`") {
        const e = codeEnd(k, end, false);
        if (e > 0) {
          k = e - 1;
          continue;
        }
        while (k + 1 < end && s[k + 1] === "`") k++;
        continue;
      }
      if (c === "[") depth++;
      else if (c === "]" && --depth === 0) return k;
    }
    return -1;
  };

  /** "(destination "title")" after a link's "]" at `close`: its end, address and title, or null. */
  const destination = (close: number, end: number): { to: number; href: string; title: string | null } | null => {
    let p = close + 1;
    if (at(p) !== "(" || literal(p)) return null;
    p++;
    while (p < end && (s[p] === " " || s[p] === "\t")) p++;
    let href = "";
    if (at(p) === "<") {
      let q = p + 1;
      while (q < end && s[q] !== ">" && s[q] !== "<" && s[q] !== "\n") q += s[q] === "\\" ? 2 : 1;
      if (at(q) !== ">") return null;
      href = s.slice(p + 1, q);
      p = q + 1;
    } else {
      let depth = 0;
      const start = p;
      while (p < end) {
        const c = s[p];
        if (c === "\\" && p + 1 < end && ESCAPABLE.includes(s[p + 1])) {
          p += 2;
          continue;
        }
        if (isSpace(c) || c === OBJECT) break;
        if (c === "(") depth++;
        else if (c === ")") {
          if (depth === 0) break;
          depth--;
        }
        p++;
      }
      if (depth) return null;
      href = s.slice(start, p);
    }
    let title: string | null = null;
    const gap = p;
    while (p < end && (s[p] === " " || s[p] === "\t")) p++;
    const quote = at(p);
    if (p > gap && (quote === '"' || quote === "'" || quote === "(")) {
      const closeQuote = quote === "(" ? ")" : quote;
      let q = p + 1;
      while (q < end && s[q] !== closeQuote) q += s[q] === "\\" ? 2 : 1;
      if (at(q) !== closeQuote) return null;
      title = s.slice(p + 1, q).replace(/\\(.)/g, "$1");
      p = q + 1;
      while (p < end && (s[p] === " " || s[p] === "\t")) p++;
    }
    if (at(p) !== ")" || literal(p)) return null;
    return { to: p + 1, href: href.replace(/\\([!-/:-@[-`{-~])/g, "$1"), title };
  };

  /** <scheme:address> or <name@domain>: its end, or -1. */
  const autolinkEnd = (i: number, end: number): number => {
    const rest = s.slice(i, Math.min(end, i + 2048));
    const url = /^<([A-Za-z][A-Za-z0-9+.-]{1,31}:[^\s<>￼]*)>/.exec(rest);
    const mail = url ? null : /^<([A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)*)>/.exec(rest);
    const match = url ?? mail;
    if (!match) return -1;
    const to = i + match[0].length;
    for (let k = i; k < to; k++) if (literal(k)) return -1;
    spans.push({ kind: "link", from: i, to, contentFrom: i + 1, contentTo: to - 1, href: mail ? `mailto:${match[1]}` : match[1], title: null });
    mark(i, i + 1);
    mark(to - 1, to);
    return to;
  };

  /** <u>…</u> and the other formatting tags: the end of the pair, or -1. */
  const tagEnd = (i: number, end: number, inLink: boolean): number => {
    const open = /^<([A-Za-z]+)>/.exec(s.slice(i, Math.min(end, i + 12)));
    const kind = open ? TAGS[open[1].toLowerCase()] : undefined;
    if (!open || !kind) return -1;
    const name = open[1].toLowerCase();
    const contentFrom = i + open[0].length;
    for (let k = i; k < contentFrom; k++) if (literal(k)) return -1;
    let depth = 1;
    for (let k = contentFrom; k < end; k++) {
      if (s[k] !== "<" || literal(k)) continue;
      const tag = /^<(\/?)([A-Za-z]+)>/.exec(s.slice(k, Math.min(end, k + 14)));
      if (!tag || tag[2].toLowerCase() !== name) continue;
      if (!tag[1]) depth++;
      else if (--depth === 0) {
        const to = k + tag[0].length;
        spans.push({ kind, from: i, to, contentFrom, contentTo: k });
        mark(i, contentFrom);
        mark(k, to);
        scan(contentFrom, k, inLink);
        return to;
      }
    }
    return -1;
  };

  /** https://… and www.… (GitHub's extended autolinks): the end, or -1. */
  const bareEnd = (i: number, end: number): number => {
    const before = at(i - 1);
    if (before && !isSpace(before) && !"*_~(".includes(before) && !isCjk(before)) return -1;
    const rest = s.slice(i, Math.min(end, i + 2048));
    const head = /^(?:https?:\/\/|www\.)/i.exec(rest);
    if (!head) return -1;
    let len = 0;
    while (len < rest.length && !isSpace(rest[len]) && rest[len] !== "<" && rest[len] !== OBJECT && !isCjk(rest[len])) len++;
    let url = rest.slice(0, len);
    // trailing punctuation is the sentence's; a ")" only when it closes nothing
    for (;;) {
      const last = url.slice(-1);
      if ("?!.,:*_~'\"".includes(last) && last) url = url.slice(0, -1);
      else if (last === ")" && (url.match(/\(/g)?.length ?? 0) < (url.match(/\)/g)?.length ?? 0)) url = url.slice(0, -1);
      else break;
    }
    const domain = url.slice(head[0].length).split(/[/?#]/)[0];
    const parts = domain.split(".");
    if (!domain || parts.some((p) => !p || !/^[\w-]+$/.test(p)) || (head[0].toLowerCase() === "www." && parts.length < 2) || parts.slice(-2).some((p) => p.includes("_"))) return -1;
    if (/^https?:\/\/$/i.test(url) || url.length <= head[0].length) return -1;
    const to = i + url.length;
    for (let k = i; k < to; k++) if (literal(k)) return -1;
    spans.push({ kind: "link", from: i, to, contentFrom: i, contentTo: to, href: /^www\./i.test(url) ? `https://${url}` : url, title: null, bare: true });
    return to;
  };

  /** "*", "_", "~~" and "==" runs: which can open and close (Typora's leniency, see the top). */
  const delimiter = (i: number, j: number, c: string): Delimiter => {
    const prev = at(i - 1);
    const next = at(j);
    const leftFlanking = !isSpace(next) && (!isPunct(next) || isSpace(prev) || isPunct(prev));
    const rightFlanking = !isSpace(prev) && (!isPunct(prev) || isSpace(next) || isPunct(next));
    let canOpen = !isSpace(next);
    if (c === "_" && isWord(prev)) canOpen = false;
    // a closer may follow a space only where it could not open ("**x ** y")
    const canClose = i > 0 && (!isSpace(prev) || isSpace(next));
    return { pos: i, char: c, length: j - i, remaining: j - i, used: 0, canOpen, canClose, both: leftFlanking && rightFlanking };
  };

  const emphasis = (delimiters: Delimiter[]) => {
    const opener = (c: number): number => {
      const closer = delimiters[c];
      let nearest = -1;
      for (let k = c - 1; k >= 0; k--) {
        const d = delimiters[k];
        if (d.remaining <= 0 || !d.canOpen || d.char !== closer.char) continue;
        if (closer.char === "~" || closer.char === "=") {
          if (d.remaining === 2 && closer.remaining === 2) return k;
          continue;
        }
        if ((d.both || closer.both) && (d.length + closer.length) % 3 === 0 && !(d.length % 3 === 0 && closer.length % 3 === 0)) continue;
        // a lone "*" is not closed by a longer run ("*it**" stays text, as in Typora)
        if (d.length === 1 && closer.length >= 2) continue;
        // a "**" closes the nearest "**" rather than a lone "*" in between
        if (nearest < 0) {
          nearest = k;
          if (!(closer.remaining >= 2 && d.remaining === 1)) return k;
        } else if (d.remaining >= 2) return k;
      }
      return nearest;
    };
    for (let c = 0; c < delimiters.length; c++) {
      const closer = delimiters[c];
      while (closer.remaining > 0 && closer.canClose) {
        const o = opener(c);
        if (o < 0) break;
        const open = delimiters[o];
        const use = closer.char === "~" || closer.char === "=" ? 2 : open.remaining >= 2 && closer.remaining >= 2 ? 2 : 1;
        const contentFrom = open.pos + open.remaining;
        const contentTo = closer.pos + closer.used;
        const kind: InlineKind = closer.char === "~" ? "strike" : closer.char === "=" ? "highlight" : use === 2 ? "strong" : "em";
        spans.push({ kind, from: contentFrom - use, to: contentTo + use, contentFrom, contentTo });
        mark(contentFrom - use, contentFrom);
        mark(contentTo, contentTo + use);
        open.remaining -= use;
        closer.remaining -= use;
        closer.used += use;
        for (let k = o + 1; k < c; k++) delimiters[k].remaining = 0;
      }
    }
  };

  function scan(start: number, end: number, inLink: boolean) {
    const delimiters: Delimiter[] = [];
    let i = start;
    while (i < end) {
      const kept = empties.get(i);
      if (kept && kept.to <= end) {
        const half = EMPTY_MARKERS[kept.kind]!.length / 2;
        spans.push({ kind: kept.kind, from: kept.from, to: kept.to, contentFrom: kept.from + half, contentTo: kept.from + half });
        mark(kept.from, kept.to);
        i = kept.to;
        continue;
      }
      const c = s[i];
      if (literal(i)) {
        i++;
        continue;
      }
      if (c === "\\" && i + 1 < end && ESCAPABLE.includes(s[i + 1]) && !literal(i + 1)) {
        spans.push({ kind: "escape", from: i, to: i + 2, contentFrom: i + 1, contentTo: i + 2 });
        mark(i, i + 1);
        i += 2;
        continue;
      }
      if (c === "`") {
        const e = codeEnd(i, end, true);
        if (e > 0) i = e;
        else while (i < end && s[i] === "`") i++;
        continue;
      }
      if (c === "<") {
        let e = autolinkEnd(i, end);
        if (e < 0) e = tagEnd(i, end, inLink);
        i = e > 0 ? e : i + 1;
        continue;
      }
      if (c === "!" && at(i + 1) === "[") {
        // an image: left as it is written (pictures come from the image commands)
        const close = closingBracket(i + 1, end);
        const dest = close > 0 ? destination(close, end) : null;
        i = dest ? dest.to : i + 1;
        continue;
      }
      if (c === "[" && !inLink) {
        const close = closingBracket(i, end);
        const dest = close > 0 ? destination(close, end) : null;
        if (dest) {
          spans.push({ kind: "link", from: i, to: dest.to, contentFrom: i + 1, contentTo: close, href: dest.href, title: dest.title });
          mark(i, i + 1);
          mark(close, dest.to);
          // the address and title between "](" and ")"
          for (let k = close + 2; k < dest.to - 1; k++) syntax[k] = 2;
          scan(i + 1, close, true);
          i = dest.to;
          continue;
        }
        i++;
        continue;
      }
      if (!inLink && (c === "h" || c === "H" || c === "w" || c === "W")) {
        const e = bareEnd(i, end);
        if (e > 0) {
          i = e;
          continue;
        }
      }
      if (c === "*" || c === "_" || c === "~" || c === "=") {
        let j = i;
        while (j < end && s[j] === c && !literal(j)) j++;
        if (c === "*" || c === "_" || j - i === 2) delimiters.push(delimiter(i, j, c));
        i = j;
        continue;
      }
      i++;
    }
    emphasis(delimiters);
  }

  scan(0, n, false);
  // outer spans first: by start, then the longer one
  spans.sort((a, b) => a.from - b.from || b.to - a.to);
  return { spans, syntax };
}

/** The formatting a character gets from the spans around it (markers get their enclosing spans' only). */
export function marksOfParse(parse: InlineParse, length: number): Array<InlineSpan[]> {
  const out: InlineSpan[][] = Array.from({ length }, () => []);
  for (const span of parse.spans) {
    if (span.kind === "escape") continue;
    for (let i = span.contentFrom; i < span.contentTo; i++) out[i].push(span);
  }
  return out;
}

// ------------------------------------------------------------- writing back --

/** A span's formatting as a character carries it. */
export interface InlineMark {
  kind: Exclude<InlineKind, "escape">;
  href?: string;
  title?: string | null;
}

export const sameMark = (a: InlineMark, b: InlineMark) =>
  a.kind === b.kind && (a.kind !== "link" || (a.href === b.href && (a.title ?? null) === (b.title ?? null)));

/** What markdown writes, outer to inner (a link outside, code inside). */
const ORDER: InlineMark["kind"][] = ["link", "strong", "em", "strike", "highlight", "underline", "superscript", "subscript", "kbd", "code"];

export interface WriteUnit {
  /** the character (OBJECT for an inline object) */
  ch: string;
  marks: InlineMark[];
  /** typed escaped: written as it is, never read as Markdown */
  literal?: boolean;
}

export interface Written {
  text: string;
  /** where each unit went in `text` */
  at: number[];
  /** the marks each unit was written with (whitespace at a span's ends is left outside it) */
  marks: InlineMark[][];
  /** before each unit (and at the end, index n): where the markers between it and the unit before start, and where
   *  the closing ones end and the opening ones begin */
  gapStart: number[];
  split: number[];
}

/** Characters that could be read as Markdown where they stand (escaped with "\" when asked). */
const RISKY = "\\*_~=`[]<>";

/** The Markdown of a run of formatted text: markers put around it, so that reading it gives the formatting back. */
export function writeInline(units: WriteUnit[], options: { em?: "*" | "_"; escape?: boolean } = {}): Written {
  const n = units.length;
  const marks = units.map((u) => [...u.marks]);
  const has = (i: number, m: InlineMark) => i >= 0 && i < n && marks[i].some((x) => sameMark(x, m));
  // code runs keep the formatting on both sides of them open (**a `c` b** stays one bold span)
  for (let i = 0; i < n; ) {
    if (!marks[i].some((m) => m.kind === "code")) {
      i++;
      continue;
    }
    let j = i;
    while (j < n && marks[j].some((m) => m.kind === "code")) j++;
    if (i > 0 && j < n) {
      const around = marks[i - 1].filter((m) => m.kind !== "code" && marks[j].some((x) => sameMark(x, m)));
      for (let k = i; k < j; k++) for (const m of around) if (!marks[k].some((x) => sameMark(x, m))) marks[k].push(m);
    }
    i = j;
  }
  // whitespace at the ends of a span goes outside it (an opening "**" cannot have a space after it)
  const space = (i: number) => /^\s$/u.test(units[i].ch);
  for (const kind of ["strong", "em", "strike", "highlight", "underline", "superscript", "subscript", "kbd"] as const) {
    for (let i = 0; i < n; ) {
      const m = marks[i].find((x) => x.kind === kind);
      if (!m) {
        i++;
        continue;
      }
      let j = i;
      while (j < n && has(j, m)) j++;
      let a = i;
      let b = j;
      while (a < b && space(a)) a++;
      while (b > a && space(b - 1)) b--;
      for (let k = i; k < j; k++) if (k < a || k >= b) marks[k] = marks[k].filter((x) => !sameMark(x, m));
      i = j;
    }
  }
  const extent = (i: number, m: InlineMark) => {
    let j = i;
    while (j < n && has(j, m)) j++;
    return j - i;
  };
  const runText = (i: number, m: InlineMark) => {
    let out = "";
    for (let j = i; j < n && has(j, m); j++) out += units[j].ch;
    return out;
  };
  const em = options.em ?? "*";
  let text = "";
  const at: number[] = new Array(n);
  const gapStart: number[] = new Array(n + 1);
  const split: number[] = new Array(n + 1);
  const stack: Array<{ mark: InlineMark; close: string }> = [];
  const open = (i: number, m: InlineMark) => {
    let before = "";
    let after = "";
    switch (m.kind) {
      case "strong":
        before = after = "**";
        break;
      case "em":
        before = after = em;
        break;
      case "strike":
        before = after = "~~";
        break;
      case "highlight":
        before = after = "==";
        break;
      case "underline":
      case "superscript":
      case "subscript":
      case "kbd": {
        const tag = { underline: "u", superscript: "sup", subscript: "sub", kbd: "kbd" }[m.kind];
        before = `<${tag}>`;
        after = `</${tag}>`;
        break;
      }
      case "code": {
        const content = runText(i, m);
        const longest = Math.max(0, ...(content.match(/`+/g) ?? []).map((r) => r.length));
        const ticks = "`".repeat(longest + 1);
        const pad = content.startsWith("`") || content.endsWith("`") || (content.length >= 2 && content.startsWith(" ") && content.endsWith(" ") && content.trim()) ? " " : "";
        before = ticks + pad;
        after = pad + ticks;
        break;
      }
      case "link": {
        const content = runText(i, m);
        const href = m.href ?? "";
        const bare = !m.title && (content === href || (/^www\./i.test(content) && href === `https://${content}`)) && /^(?:https?:\/\/|www\.)\S+$/i.test(content);
        const auto = !m.title && content === href && /^[A-Za-z][A-Za-z0-9+.-]{1,31}:[^\s<>]*$/.test(href);
        if (bare) break;
        if (auto) {
          before = "<";
          after = ">";
          break;
        }
        const balanced = (href.match(/\(/g)?.length ?? 0) === (href.match(/\)/g)?.length ?? 0);
        const dest = /[\s<>]/.test(href) || !balanced ? `<${href}>` : href;
        before = "[";
        after = `](${dest}${m.title ? ` "${m.title.replace(/"/g, '\\"')}"` : ""})`;
        break;
      }
    }
    text += before;
    stack.push({ mark: m, close: after });
  };
  for (let i = 0; i < n; i++) {
    const want = marks[i];
    gapStart[i] = text.length;
    const k = stack.findIndex((e) => !want.some((x) => sameMark(x, e.mark)));
    // what is inside a span that ends closes with it (and opens again after it)
    if (k >= 0) while (stack.length > k) text += stack.pop()!.close;
    split[i] = text.length;
    const missing = want.filter((m) => !stack.some((e) => sameMark(e.mark, m)));
    missing.sort((a, b) => extent(i, b) - extent(i, a) || ORDER.indexOf(a.kind) - ORDER.indexOf(b.kind));
    for (const m of missing) open(i, m);
    const inCode = want.some((m) => m.kind === "code");
    const ch = units[i].ch;
    if (options.escape && !inCode && !units[i].literal && RISKY.includes(ch)) text += "\\";
    at[i] = text.length;
    text += ch;
  }
  gapStart[n] = text.length;
  while (stack.length) text += stack.pop()!.close;
  split[n] = text.length;
  return { text, at, marks, gapStart, split };
}

/** Whether two characters' formatting is the same (code excludes the rest, as in the document). */
export function sameMarks(a: readonly InlineMark[], b: readonly InlineMark[]): boolean {
  const norm = (list: readonly InlineMark[]) => (list.some((m) => m.kind === "code") ? list.filter((m) => m.kind === "code") : list);
  const x = norm(a);
  const y = norm(b);
  return x.length === y.length && x.every((m) => y.some((o) => sameMark(m, o)));
}
