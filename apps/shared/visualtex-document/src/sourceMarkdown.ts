// The document's text is its Markdown source, line by line, as in Obsidian's editor (2026-11-10, second form): every
// line of the document is a paragraph (a heading while its text starts with "#"), and lists ("- ", "1. ", "- [ ] " with
// tabs for nesting), quotes ("> "), code fences, rules ("---") are those lines' text. What the text means is read with
// the tokenizer Obsidian's editor uses (CodeMirror 5's Markdown mode, cm5Markdown.ts), line after line with the state
// the lines before left, and drawn as Obsidian's Live Preview draws it (its rules read from Obsidian 1.15, see
// ~/claude/obsidian-study/FINDINGS.md): tokens in a row form a group (a whole `**bold**` or `[text](url)`, a bullet's
// "-" alone), and a group's markers hide unless the selection touches it, both ends counted; a heading's "#", a quote's
// ">" and an escape's "\" while the selection is not in their line; a bullet's "-" is drawn as a bullet, a task's
// "[ ]" as a box; nothing is drawn again while an input method composes or the mouse button is down. Formulas,
// tables, pictures, the title and page commands stay objects of the document between the lines (formulas.ts etc.).

import { Fragment, Slice, type Mark, type Node, type ResolvedPos, type Schema } from "prosemirror-model";
import { Plugin, PluginKey, TextSelection, type Command, type EditorState, type Transaction } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";
import { OBJECT, writeInline, type InlineMark, type WriteUnit } from "./inlineMarkdown.ts";
import { blockLanguage, documentToMarkdown, NO_CAPTION, readMarkdown } from "./markdown.ts";
import { captionLabelLength } from "./figures.ts";
import { CODE_LANGUAGES } from "./codeBlocks.ts";
import { homeInLine, renumberLists } from "./sourceKeys.ts";
import { blankLine as tokenizerBlankLine, copyState, startState, stateKey, tokenizeLine, type LineFacts, type MarkdownState, type MarkdownToken } from "./cm5Markdown.ts";

// ------------------------------------------------------------------------------------------------- a line's syntax

/** A formatted span: its range in the line's text and its markers (hidden while the selection does not touch it). */
export interface SourceUnit {
  from: number;
  to: number;
  /** its markers */
  marks: Array<[number, number]>;
  /** text hidden while the selection is away and shown as text otherwise (an escape's "\") */
  plain?: Array<[number, number]>;
  /** what it is (strong, em, strike, highlight, code, link, image, escape, break, html, underline, …, heading) */
  kind: string;
  /** a link's address */
  href?: string;
  /** a link's or a picture's text (between "[" and "]") */
  text?: [number, number];
  /** "line": the markers show while the selection is in the span's line */
  reveal?: "line";
  /** drawn after the span (a picture), or in its place while it is hidden (a line break for <br>) */
  widget?: { kind: "image"; src: string; alt: string } | { kind: "break" };
}

/** The style of a stretch of the line's text. */
export interface SourceStyle {
  from: number;
  to: number;
  kind: "strong" | "em" | "strike" | "highlight" | "code" | "link" | "url" | "underline" | "superscript" | "subscript" | "kbd" | "color";
  href?: string;
  color?: string;
  background?: string;
}

export interface HeadingSyntax {
  level: number;
  /** the end of "#… " (the markers and the spaces after them; 0 for an underlined heading) */
  prefixEnd: number;
  close: [number, number] | null;
  /** Pandoc's " {-}" at the end: an unnumbered section */
  attributes: [number, number] | null;
  unnumbered: boolean;
}

export interface BlockSyntax {
  text: string;
  tokens: MarkdownToken[];
  facts: LineFacts;
  units: SourceUnit[];
  styles: SourceStyle[];
  heading: HeadingSyntax | null;
}

/** An inline object (a formula, a footnote) in a line's text: reads as a letter would. */
const OBJECT_CHAR = "";

/** A textblock's text as Markdown reads it: each inline object one character (a line break a newline). */
export function blockText(node: Node): string {
  let text = "";
  node.forEach((child) => {
    if (child.isText) text += child.text ?? "";
    else if (child.type.name === "hard_break") text += "\n";
    else text += OBJECT_CHAR;
  });
  return text;
}

/** "#" to "######" at a line's start, then spaces or the line's end: a heading (with its closing run and " {-}"). */
export function headingSyntax(text: string): HeadingSyntax | null {
  const m = /^(#{1,6})(?=[ \t]|$)[ \t]*/.exec(text);
  if (!m) return null;
  let end = text.length;
  let attributes: [number, number] | null = null;
  let unnumbered = false;
  const a = /[ \t]+\{([^{}]*)\}[ \t]*$/.exec(text.slice(m[0].length));
  if (a && a[1].trim().split(/\s+/).some((part) => part === "-" || part === ".unnumbered")) {
    attributes = [m[0].length + a.index, text.length];
    unnumbered = true;
    end = attributes[0];
  }
  let close: [number, number] | null = null;
  const rest = text.slice(m[0].length, end);
  const c = /[ \t](#+)[ \t]*$/.exec(rest);
  if (c) {
    const start = m[0].length + c.index + 1;
    close = [start, start + c[1].length];
  } else if (/^#+[ \t]*$/.test(rest) && rest.length) {
    close = [m[0].length, m[0].length + rest.replace(/[ \t]+$/, "").length];
  }
  return { level: m[1].length, prefixEnd: m[0].length, close, attributes, unnumbered };
}

const HTML_STYLES: Record<string, SourceStyle["kind"]> = {
  u: "underline", ins: "underline", sup: "superscript", sub: "subscript", kbd: "kbd", mark: "highlight", b: "strong",
  strong: "strong", i: "em", em: "em", s: "strike", del: "strike", strike: "strike",
};
const VOID_TAGS = new Set(["area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr"]);
const SAFE_COLOR = /^(#[0-9a-f]{3,8}|[a-z]+|rgba?\([\d\s.,%]+\)|hsla?\([\d\s.,%a-z]+\))$/i;
/** Pictures drawn from their address as it stands (a web address or data). */
const SHOWN_PICTURE = /^(https?:\/\/|data:image\/)/i;
const BARE_URL = /(?:https?:\/\/|www\.)[^\s<]*[^\s<?!.,:*_~)'"]/gi;

function tagColors(attrs: string): { color?: string; background?: string } {
  const style = /\bstyle\s*=\s*(["'])(.*?)\1/is.exec(attrs)?.[2] ?? "";
  const out: { color?: string; background?: string } = {};
  for (const declaration of style.split(";")) {
    const at = declaration.indexOf(":");
    if (at < 0) continue;
    const property = declaration.slice(0, at).trim().toLowerCase();
    const value = declaration.slice(at + 1).trim();
    if (!SAFE_COLOR.test(value)) continue;
    if (property === "color") out.color = value;
    else if (property === "background-color" || property === "background") out.background = value;
  }
  return out;
}

const has = (token: MarkdownToken, cls: string) => (" " + token.type + " ").includes(" " + cls + " ");

const SPAN_KINDS: Array<[string, SourceUnit["kind"], SourceStyle["kind"]]> = [
  ["strong", "strong", "strong"],
  ["em", "em", "em"],
  ["strikethrough", "strike", "strike"],
  ["highlight", "highlight", "highlight"],
  ["code", "code", "code"],
];

/** The spans and styles a line's tokens make (for the commands, exports and the text as it reads). */
function spansOf(text: string, tokens: MarkdownToken[]): { units: SourceUnit[]; styles: SourceStyle[] } {
  const units: SourceUnit[] = [];
  const styles: SourceStyle[] = [];
  const open = new Map<string, [number, number]>();
  let link: { from: number; textFrom: number; textTo: number; urlFrom: number; urlTo: number; image: boolean } | null = null;
  for (const token of tokens) {
    if (has(token, "formatting-escape")) {
      units.push({ from: token.from, to: token.to, marks: [], plain: [[token.from, token.from + 1]], kind: "escape", reveal: "line" });
      continue;
    }
    for (const [name, kind, style] of SPAN_KINDS) {
      if (!has(token, "formatting-" + name)) continue;
      // "***": the outer "*" is the em's, the inner "**" the strong's
      let [from, to] = [token.from, token.to];
      if (has(token, "formatting-strong") && has(token, "formatting-em") && to - from === 3) {
        const opening = !open.has(kind);
        if (kind === "em") [from, to] = opening ? [from, from + 1] : [to - 1, to];
        else [from, to] = opening ? [from + 1, to] : [from, to - 1];
      }
      const start = open.get(kind);
      if (!start) open.set(kind, [from, to]);
      else {
        open.delete(kind);
        units.push({ from: start[0], to, marks: [start, [from, to]], kind });
        if (from > start[1]) styles.push({ from: start[1], to: from, kind: style });
      }
    }
    // links and pictures: "[" / "![" … "]" "(" address ")"
    const piece = text.slice(token.from, token.to);
    if ((has(token, "formatting-link") || has(token, "formatting-image")) && !has(token, "formatting-link-string")) {
      if (piece === "[" || piece === "!" || piece === "<") {
        if (!link) link = { from: token.from, textFrom: token.to, textTo: token.to, urlFrom: -1, urlTo: -1, image: has(token, "formatting-image") };
        else if (piece === "[") link.textFrom = token.to;
      } else if ((piece === "]" || piece === ">") && link) {
        link.textTo = token.from;
        if (piece === ">") {
          // <address>
          const href = text.slice(link.textFrom, link.textTo);
          units.push({ from: link.from, to: token.to, marks: [[link.from, link.textFrom], [token.from, token.to]], kind: "link", href });
          styles.push({ from: link.textFrom, to: link.textTo, kind: "link", href });
          link = null;
        } else if (!/^\s?[([]/.test(text.slice(token.to, token.to + 2))) link = null;
      }
    } else if (has(token, "formatting-link-string") && link) {
      if (piece === "(" || piece === "[") link.urlFrom = token.to;
      else {
        link.urlTo = token.from;
        const href = text.slice(link.urlFrom, link.urlTo).trim().split(/\s+/)[0] ?? "";
        const kind = link.image ? "image" : "link";
        const shownPicture = link.image && SHOWN_PICTURE.test(href);
        units.push({
          from: link.from,
          to: token.to,
          marks: [[link.from, link.textFrom], [link.textTo, link.urlFrom], [link.urlTo, token.to]],
          plain: [[link.urlFrom, link.urlTo], ...(shownPicture ? [[link.textFrom, link.textTo] as [number, number]] : [])],
          kind,
          href,
          text: [link.textFrom, link.textTo],
          ...(shownPicture ? { widget: { kind: "image" as const, src: href, alt: text.slice(link.textFrom, link.textTo) } } : {}),
        });
        if (link.textTo > link.textFrom) styles.push({ from: link.textFrom, to: link.textTo, kind: "link", href });
        if (link.urlTo > link.urlFrom) styles.push({ from: link.urlFrom, to: link.urlTo, kind: "url", href });
        link = null;
      }
    }
  }
  // spans left open at the line's end format the rest of it (the tokenizer reads them so, as Obsidian shows them)
  for (const [kind, start] of open) {
    units.push({ from: start[0], to: text.length, marks: [start], kind });
    const style = SPAN_KINDS.find((k) => k[1] === kind)![2];
    if (text.length > start[1]) styles.push({ from: start[1], to: text.length, kind: style });
  }
  // code spans and links hold what follows literally
  const literal = (at: number) => units.some((u) => (u.kind === "code" || u.kind === "link" || u.kind === "image") && at >= u.from && at < u.to);
  for (const m of text.matchAll(BARE_URL)) {
    if (literal(m.index!)) continue;
    const href = m[0];
    styles.push({ from: m.index!, to: m.index! + href.length, kind: "link", href: /^www\./i.test(href) ? `https://${href}` : href });
  }
  // inline HTML: a tag and its closing tag around text style it (<u>, <sup>, <kbd>, <span style="color: …">), <br>
  const tags: Array<{ tag: string; from: number; to: number; attrs: string }> = [];
  for (const m of text.matchAll(/<(\/?)([a-zA-Z][\w-]*)([^<>]*)>/g)) {
    const from = m.index!;
    const to = from + m[0].length;
    if (literal(from)) continue;
    const tag = m[2].toLowerCase();
    if (tag === "br" && !m[1]) {
      units.push({ from, to, marks: [[from, to]], kind: "break", widget: { kind: "break" } });
      continue;
    }
    if (VOID_TAGS.has(tag) || /\/\s*$/.test(m[3])) continue;
    if (!m[1]) {
      tags.push({ tag, from, to, attrs: m[3] });
      continue;
    }
    const at = tags.map((t) => t.tag).lastIndexOf(tag);
    if (at < 0) continue;
    const start = tags[at];
    tags.splice(at);
    const kind = HTML_STYLES[tag];
    units.push({ from: start.from, to, marks: [[start.from, start.to], [from, to]], kind: kind ?? "html" });
    if (from > start.to) {
      if (kind) styles.push({ from: start.to, to: from, kind });
      const colors = tagColors(start.attrs);
      if (colors.color || colors.background) styles.push({ from: start.to, to: from, kind: "color", ...colors });
    }
  }
  return { units, styles };
}

function makeSyntax(text: string, tokens: MarkdownToken[], facts: LineFacts): BlockSyntax {
  const { units, styles } = spansOf(text, tokens);
  let heading: HeadingSyntax | null = null;
  if (facts.header && !facts.list && !facts.quote && !facts.codeInside && !facts.codeBegin) {
    const atx = headingSyntax(text);
    if (atx) heading = atx;
    else if (!/^\s*(=+|-+)\s*$/.test(text)) heading = { level: facts.header, prefixEnd: 0, close: null, attributes: null, unnumbered: false };
  }
  return { text, tokens, facts, units, styles, heading };
}

const frontMatterReads = new WeakMap<BlockSyntax, BlockSyntax>();

/** A line of front matter: YAML text (not a heading, a rule or Markdown). */
function frontMatterSyntax(syntax: BlockSyntax, tokens: MarkdownToken[]): BlockSyntax {
  let read = frontMatterReads.get(syntax);
  if (!read) {
    read = { text: syntax.text, tokens, facts: { header: 0, list: 0, quote: 0, codeBegin: false, codeEnd: false, codeInside: false, indentedCode: false, codeLang: "", hr: false, blank: !syntax.text, frontMatter: true }, units: [], styles: [], heading: null };
    frontMatterReads.set(syntax, read);
  }
  return read;
}

/** Paragraphs and headings are the document's lines (a title, a caption or a table cell is read on its own). */
function headingCapable(node: Node): boolean {
  return node.type.name === "paragraph" || node.type.name === "heading";
}

interface LineRead {
  syntax: BlockSyntax;
  out: MarkdownState;
}

/** Lines read before: per line node, per state it was read with. */
const reads = new WeakMap<Node, Map<string, LineRead>>();
/** Each textblock's syntax as last read in its document (blockSyntax). */
const lastRead = new WeakMap<Node, BlockSyntax>();

const SETEXT = /^ {0,3}(?:=+|-{2,})\s*$/;

/** A textblock read alone (a title, a caption, a table cell; a line not read in its document yet). */
function readAlone(node: Node): BlockSyntax {
  const text = blockText(node);
  const { tokens, facts } = tokenizeLine(text, startState(), () => undefined);
  const syntax = makeSyntax(text, tokens, headingCapable(node) ? facts : { ...facts, header: 0 });
  lastRead.set(node, syntax);
  return syntax;
}

/** A textblock's Markdown as its document reads it. */
export function blockSyntax(node: Node): BlockSyntax {
  return lastRead.get(node) ?? readAlone(node);
}

export interface SourceLine {
  node: Node;
  pos: number;
  syntax: BlockSyntax;
}

/** A fenced code block: its first and last line (indices into the lines) and language. */
interface CodeBlockRange {
  first: number;
  last: number;
  lang: string;
}

interface DocumentLines {
  lines: SourceLine[];
  code: CodeBlockRange[];
  /** textblocks that are not lines (in tables, captions, the title) */
  others: Array<{ node: Node; pos: number; syntax: BlockSyntax }>;
}

/** Every line of a document read in order (unchanged lines with the same state before them are not read again). */
function readDocument(doc: Node): DocumentLines {
  const lines: SourceLine[] = [];
  const others: DocumentLines["others"] = [];
  const children: Array<{ node: Node; pos: number }> = [];
  doc.forEach((node, pos) => children.push({ node, pos }));
  let state = startState();
  const textOf = (i: number) => {
    const child = children[i];
    return child && child.node.isTextblock && headingCapable(child.node) ? blockText(child.node) : undefined;
  };
  children.forEach(({ node, pos }, i) => {
    if (node.isTextblock && headingCapable(node)) {
      const next = textOf(i + 1);
      const key = stateKey(state) + (next !== undefined && SETEXT.test(next) ? "|s" : "");
      let byState = reads.get(node);
      let read = byState?.get(key);
      if (!read) {
        const text = blockText(node);
        const working = copyState(state);
        const { tokens, facts } = tokenizeLine(text, working, (n) => textOf(i + n));
        read = { syntax: makeSyntax(text, tokens, facts), out: working };
        if (!byState) reads.set(node, (byState = new Map()));
        byState.set(key, read);
      }
      state = copyState(read.out);
      lastRead.set(node, read.syntax);
      lines.push({ node, pos, syntax: read.syntax });
      return;
    }
    // an object between lines (a formula, a table, a picture, the title): what was open ends there
    tokenizerBlankLine(state);
    state.code = 0;
    state.fencedEndRE = null;
    if (node.isTextblock) others.push({ node, pos, syntax: readAlone(node) });
    else
      node.descendants((child, offset) => {
        if (!child.isTextblock) return true;
        if (!child.type.spec.code) others.push({ node: child, pos: pos + 1 + offset, syntax: readAlone(child) });
        return false;
      });
  });
  // front matter: the document's first line "---" to the next "---" (or "..."): YAML lines (Obsidian's properties)
  if (lines.length && lines[0].pos === 0 && lines[0].syntax.text === "---") {
    const end = lines.findIndex((l, i) => i > 0 && (l.syntax.text === "---" || l.syntax.text === "..."));
    if (end > 0) {
      for (let i = 0; i <= end; i++) {
        const syntax = lines[i].syntax;
        const tokens = syntax.text ? [{ from: 0, to: syntax.text.length, type: i === 0 || i === end ? "formatting hmd-frontmatter" : "hmd-frontmatter" }] : [];
        lines[i] = { ...lines[i], syntax: frontMatterSyntax(syntax, tokens) };
        lastRead.set(lines[i].node, lines[i].syntax);
      }
    }
  }
  // fenced code blocks
  const code: CodeBlockRange[] = [];
  let first = -1;
  lines.forEach((line, i) => {
    const f = line.syntax.facts;
    if (f.codeBegin) first = i;
    else if (f.codeEnd && first >= 0) {
      code.push({ first, last: i, lang: lines[first].syntax.facts.codeLang });
      first = -1;
    }
  });
  if (first >= 0) code.push({ first, last: lines.length - 1, lang: lines[first].syntax.facts.codeLang });
  return { lines, code, others };
}

// ------------------------------------------------------------------------------------------------- drawing it

const TOKEN_CLASSES: Array<[string, string]> = [
  ["strong", "md-strong"],
  ["em", "md-em"],
  ["strikethrough", "md-strike"],
  ["highlight", "md-highlight"],
  ["inline-code", "md-code"],
];

const STYLE_CLASS: Record<SourceStyle["kind"], string> = {
  strong: "md-strong", em: "md-em", strike: "md-strike", highlight: "md-highlight", code: "md-code", link: "md-link",
  url: "md-url", underline: "md-underline", superscript: "md-sup", subscript: "md-sub", kbd: "md-kbd", color: "md-color",
};

/** The selection as the drawing uses it: its ranges (none while the editor does not have the focus, as in Obsidian). */
type Ranges = ReadonlyArray<{ from: number; to: number }>;

/** Inclusive overlap (Obsidian's test): a range touches [from, to] when it reaches either end. */
function touched(ranges: Ranges, from: number, to: number): boolean {
  return ranges.some((r) => r.from <= to && r.to >= from);
}

function rangesOf(state: EditorState): Ranges {
  return state.selection.ranges.map((r) => ({ from: r.$from.pos, to: r.$to.pos }));
}

function checkboxDOM(checked: boolean) {
  return (view: EditorView, getPos: () => number | undefined) => {
    const label = document.createElement("span");
    label.className = "md-task";
    label.contentEditable = "false";
    const box = document.createElement("input");
    box.type = "checkbox";
    box.checked = checked;
    box.tabIndex = -1;
    label.appendChild(box);
    label.addEventListener("mousedown", (event) => event.preventDefault());
    box.addEventListener("click", (event) => {
      event.preventDefault();
      const pos = getPos();
      if (pos === undefined) return;
      // the "[ ]" / "[x]" the box stands for starts at the widget's place
      const $pos = view.state.doc.resolve(pos);
      const text = $pos.parent.textBetween($pos.parentOffset, $pos.parent.content.size, undefined, OBJECT_CHAR);
      if (!/^\[.\]/.test(text)) return;
      const mark = pos + 1;
      const now = view.state.doc.textBetween(mark, mark + 1);
      view.dispatch(view.state.tr.insertText(now === " " ? "x" : " ", mark, mark + 1));
    });
    return label;
  };
}

function quoteBorder() {
  const border = document.createElement("span");
  border.className = "md-quote-border";
  border.textContent = ">";
  return border;
}

function pictureWidget(src: string, alt: string) {
  return () => {
    const box = document.createElement("span");
    box.className = "md-picture";
    box.contentEditable = "false";
    const img = document.createElement("img");
    img.src = src;
    img.alt = alt;
    img.draggable = false;
    box.appendChild(img);
    return box;
  };
}

const lineBreakWidget = () => document.createElement("br");

/** A zero-width character that keeps a line whose text is hidden one line high. */
function lineHolder() {
  const holder = document.createElement("span");
  holder.className = "md-line-holder";
  holder.textContent = "\u200b";
  return holder;
}

function ruleWidget() {
  const rule = document.createElement("span");
  rule.className = "md-rule";
  rule.contentEditable = "false";
  return rule;
}

function codeLabel(lang: string) {
  return () => {
    const label = document.createElement("span");
    label.className = "md-code-language";
    label.contentEditable = "false";
    label.textContent = lang;
    return label;
  };
}

function calloutTitle(type: string, title: string) {
  return () => {
    const label = document.createElement("span");
    label.className = "md-callout-title";
    label.contentEditable = "false";
    label.textContent = title || type.charAt(0).toUpperCase() + type.slice(1).toLowerCase();
    return label;
  };
}

interface LineContext {
  /** the selection's ranges (empty: drawn with nothing revealed) */
  ranges: Ranges;
  /** the fenced code block the line is in */
  block: CodeBlockRange | null;
  /** the document range of that block (first line's start to last line's end) */
  blockRange: [number, number] | null;
  /** a callout the line starts ("> [!NOTE] Title") */
  callout: { type: string; title: string } | null;
  /** whether the line is in a callout */
  inCallout: string | null;
  /** a plain line right after another line of the same paragraph (Markdown joins them; no first-line indent) */
  continuation: boolean;
}

const hide = (from: number, to: number) => Decoration.inline(from, to, { class: "md-hidden" }, { md: "hide" });

/**
 * A line's decorations, as Obsidian's Live Preview draws its tokens (FINDINGS.md "The reveal rule"): styles for every
 * token; formatting tokens grouped with the span they belong to, hidden unless the selection touches the group.
 */
function lineDecorations(line: SourceLine, cx: LineContext): Decoration[] {
  const { node, pos, syntax } = line;
  const start = pos + 1;
  const end = start + node.content.size;
  const { text, tokens, facts } = syntax;
  const out: Decoration[] = [];
  const lineTouched = touched(cx.ranges, start, end);

  // the line itself: a list's hanging indent, a quote's border, code, a rule, a task done
  const lineClasses: string[] = [];
  const lineStyle: string[] = [];
  if (facts.codeBegin || facts.codeInside || facts.codeEnd) {
    lineClasses.push("md-codeline");
    if (facts.codeBegin) lineClasses.push("md-codeline-begin");
    if (facts.codeEnd) lineClasses.push("md-codeline-end");
  } else if (facts.list) {
    const prefix = /^([>\s]*)(([*+-] |(\d+)([.)] ))(?:\[(.)\] )?)?/.exec(text);
    const indent = prefix ? prefix[0].replace(/\t/g, "    ").length : 0;
    lineClasses.push("md-listline");
    if (prefix?.[6] && prefix[6] !== " ") lineClasses.push("md-task-done");
    lineStyle.push(`padding-left:${indent}ch`, `text-indent:-${indent}ch`);
  }
  if (facts.quote) {
    lineClasses.push("md-quoteline", `md-quote-${Math.min(facts.quote, 6)}`);
    if (cx.inCallout) lineClasses.push("md-calloutline", `md-callout-${cx.inCallout}`);
  }
  if (facts.hr) lineClasses.push("md-ruleline");
  if (cx.continuation) lineClasses.push("md-cont");
  if (lineClasses.length || lineStyle.length) {
    out.push(Decoration.node(pos, pos + node.nodeSize, { class: lineClasses.join(" "), ...(lineStyle.length ? { style: lineStyle.join(";") } : {}) }));
  }

  // a fenced code block's fences: shown while the selection is in the block, else hidden (the first with its language)
  if ((facts.codeBegin || facts.codeEnd) && cx.blockRange) {
    const inBlock = touched(cx.ranges, cx.blockRange[0], cx.blockRange[1]);
    for (const token of tokens) out.push(Decoration.inline(start + token.from, start + token.to, { class: "md-mark" }));
    if (!inBlock && text.length) {
      out.push(hide(start, end));
      // (the line keeps its height: Up / Down stop on it, as on Obsidian's)
      out.push(Decoration.widget(end, lineHolder, { side: 1, md: "holder", key: "holder" }));
      if (facts.codeBegin && cx.block?.lang) {
        // the language's name as the code block's header shows it ("js" -> "JavaScript")
        const id = blockLanguage(cx.block.lang);
        const name = CODE_LANGUAGES.find((l) => l.id === id)?.label ?? cx.block.lang;
        out.push(Decoration.widget(end, codeLabel(name), { side: 1, md: "widget", key: `lang:${name}` }));
      }
    }
    return out;
  }
  if (facts.frontMatter) {
    out.push(Decoration.node(pos, pos + node.nodeSize, { class: "md-frontmatter" }));
    return out;
  }
  if (facts.codeInside) return out;
  // an indented code line (four spaces or a tab): its text styled as code, in a line of its own (Obsidian)
  if (facts.indentedCode) {
    const at = text.search(/\S/);
    if (at >= 0) out.push(Decoration.inline(start + at, end, { class: "md-code" }));
    return out;
  }

  // a rule: drawn as a rule unless the selection touches its line
  if (facts.hr) {
    // (a rule's "---" shown is text, as Obsidian shows it)
    if (!lineTouched) {
      out.push(hide(start, end));
      out.push(Decoration.widget(end, ruleWidget, { side: 1, md: "widget", key: "rule" }));
    }
    return out;
  }

  // every token's style
  for (const token of tokens) {
    const classes: string[] = [];
    if (has(token, "formatting")) classes.push("md-mark");
    for (const [cls, name] of TOKEN_CLASSES) if (has(token, cls)) classes.push(name);
    if (has(token, "link") && !has(token, "formatting") && !has(token, "url")) classes.push("md-link");
    if (has(token, "url") && !has(token, "formatting")) classes.push(has(token, "string") ? "md-url" : "md-link");
    if (classes.length) out.push(Decoration.inline(start + token.from, start + token.to, { class: classes.join(" ") }));
  }

  // the groups (Obsidian's walk): a group runs while tokens follow one another and belong to a span or a link
  type Pending = { from: number; to: number; deco: Decoration; always?: boolean };
  let groupFrom = 0;
  let groupTo = 0;
  let pending: Pending[] | null = null;
  const flush = (at: number) => {
    if (pending) {
      if (touched(cx.ranges, start + groupFrom, start + groupTo)) {
        for (const p of pending) if (p.always) out.push(p.deco);
      } else for (const p of pending) out.push(p.deco);
      pending = null;
    }
    groupFrom = groupTo = at;
  };
  const add = (p: Pending) => (pending ??= []).push(p);
  const taskOnLine = /^([>\s]*)([*+-] |\d+[.)] )\[(.)\] /.exec(text);
  // brackets with no address or label after them ("[ ]", "[word]") stay as written (Obsidian hides only a link's)
  const keptBrackets = new Set<number>();
  {
    const stack: number[] = [];
    tokens.forEach((token, k) => {
      if (!has(token, "formatting-link") || has(token, "formatting-link-string")) return;
      const piece = text.slice(token.from, token.to);
      if (piece === "[") stack.push(k);
      else if (piece === "]" && stack.length) {
        const open = stack.pop()!;
        if (!/^\s?[([]/.test(text.slice(token.to, token.to + 2))) {
          keptBrackets.add(open);
          keptBrackets.add(k);
        }
      }
    });
  }
  // a callout's "[!TYPE]": markup of the quote, not a link
  const calloutTag = cx.callout ? /^>\s*(\[!\w+\][+-]?)/.exec(text) : null;
  const calloutFrom = calloutTag ? text.indexOf(calloutTag[1]) : -1;
  const calloutTo = calloutTag ? calloutFrom + calloutTag[1].length : -1;
  tokens.forEach((token, tokenIndex) => {
    const x = token.from;
    const e = token.to;
    const classed = token.type !== "";
    if (calloutTag && x >= calloutFrom && e <= calloutTo) {
      out.push(Decoration.inline(start + x, start + e, { class: "md-mark" }));
      return;
    }
    if (keptBrackets.has(tokenIndex)) {
      if (x !== groupTo) flush(x);
      return;
    }
    if (x !== groupTo || !classed) flush(x);
    if (!classed) return;
    if (has(token, "formatting-header")) {
      if (!lineTouched && text.trim() !== text.slice(x, e).trim()) out.push(hide(start + x, start + e));
      return;
    }
    if (has(token, "formatting-quote")) {
      if (!lineTouched) {
        if (has(token, "formatting-quote-1")) out.push(Decoration.inline(start + x, start + x + 1, { class: "md-transparent" }, { md: "mark" }));
        else {
          out.push(hide(start + x, start + x + 1));
          out.push(Decoration.widget(start + x, quoteBorder, { side: 1, md: "widget", key: "quote" }));
        }
        // a callout's "[!TYPE] Title": its type drawn as the title
        if (cx.callout && has(token, "formatting-quote-1")) {
          const m = /^>\s*(\[!(\w+)\][+-]?)\s*/.exec(text);
          if (m) {
            const from = text.indexOf(m[1]);
            out.push(hide(start + from, start + from + m[1].length));
            if (!cx.callout.title) out.push(Decoration.widget(start + from + m[1].length, calloutTitle(cx.callout.type, ""), { side: -1, md: "widget", key: `callout:${cx.callout.type}` }));
          }
        }
      }
      return;
    }
    if (has(token, "formatting-escape")) {
      if (!lineTouched) out.push(hide(start + x, start + x + 1));
    }
    const span = TOKEN_CLASSES.some(([cls]) => has(token, cls));
    const linkish = span || has(token, "link") || has(token, "image") || has(token, "formatting-link") || has(token, "formatting-image");
    const formatting = has(token, "formatting");
    const urlPart = has(token, "string") && has(token, "url");
    const task = has(token, "formatting-task");
    const bullet = has(token, "formatting-list-ul");
    const autolinkText = has(token, "url") && !has(token, "string");
    const continues = urlPart || task || bullet || autolinkText || has(token, "formatting-link") || has(token, "formatting-image");
    if (linkish || continues) {
      groupTo = e;
      if (formatting || continues) {
        if (task) {
          const checked = text.charAt(x + 1) !== " ";
          add({ from: x, to: e, deco: Decoration.inline(start + x, start + e, { class: "md-hidden" }, { md: "hide" }) });
          add({ from: x, to: e, deco: Decoration.widget(start + x, checkboxDOM(checked), { side: 1, md: "widget", key: `task:${checked}` }) });
        } else if (bullet) {
          if (!taskOnLine) {
            groupTo = x + 1;
            add({ from: x, to: x + 1, deco: Decoration.inline(start + x, start + x + 1, { class: "md-bullet" }, { md: "mark" }) });
          } else add({ from: x, to: e, deco: hide(start + x, start + e) });
        } else if (autolinkText) {
          // an address in "<…>": shown, underlined (its "<" ">" are the group's markers)
        } else add({ from: x, to: e, deco: hide(start + x, start + e) });
      }
    }
    if (has(token, "formatting-list-ol")) {
      groupTo = e;
      add({ from: x, to: e, deco: Decoration.inline(start + x, start + e, { class: "md-number" }, { md: "mark" }) });
    }
  });
  flush(text.length);

  // inline HTML, <br>, pictures and bare addresses: their own spans
  for (const unit of syntax.units) {
    if (unit.kind === "escape" || SPAN_KINDS.some(([, kind]) => kind === unit.kind) || unit.kind === "link" || unit.kind === "image") {
      if (unit.widget?.kind === "image") out.push(Decoration.widget(start + unit.to, pictureWidget(unit.widget.src, unit.widget.alt), { side: -1, md: "widget", key: `image:${unit.widget.src}` }));
      continue;
    }
    const shown = touched(cx.ranges, start + unit.from, start + unit.to);
    for (const [from, to] of unit.marks) {
      if (to > from) out.push(Decoration.inline(start + from, start + to, { class: shown ? "md-mark" : "md-mark md-hidden" }, { md: shown ? "mark" : "hide" }));
    }
    if (unit.widget?.kind === "break" && !shown) out.push(Decoration.widget(start + unit.to, lineBreakWidget, { side: -1, md: "widget", key: "break" }));
  }
  for (const style of syntax.styles) {
    if (style.kind === "color" || style.kind === "underline" || style.kind === "superscript" || style.kind === "subscript" || style.kind === "kbd" || (style.kind === "link" && !tokens.some((t) => t.from <= style.from && t.to >= style.to && (has(t, "link") || has(t, "url"))))) {
      const attrs: Record<string, string> = { class: STYLE_CLASS[style.kind] };
      if (style.href) attrs["data-href"] = style.href;
      const css = [style.color ? `color:${style.color}` : "", style.background ? `background-color:${style.background}` : ""].filter(Boolean).join(";");
      if (css) attrs.style = css;
      out.push(Decoration.inline(start + style.from, start + style.to, attrs));
    }
  }
  return out;
}

/** Decorations of a textblock that is not a line (a table cell, a caption, the title): inline Markdown only. */
function alonePieceDecorations(node: Node, pos: number, syntax: BlockSyntax, ranges: Ranges): Decoration[] {
  return lineDecorations({ node, pos, syntax: { ...syntax, facts: { ...syntax.facts, header: 0, list: 0, quote: 0, hr: false, codeBegin: false, codeEnd: false, codeInside: false, indentedCode: false } } }, { ranges, block: null, blockRange: null, callout: null, inCallout: null, continuation: false });
}

interface SourceLiveState {
  read: DocumentLines;
  /** every line's decorations with nothing revealed */
  base: DecorationSet;
  /** what the view draws: the selection's lines (and code block) as the selection shows them */
  shown: DecorationSet;
  focused: boolean;
  composing: boolean;
  pointer: boolean;
  /** every marker hidden whatever the selection (printing) until the selection or the text changes */
  away?: boolean;
}

type LiveMeta = { focused: boolean } | { composing: boolean } | { pointer: boolean } | "away" | "caret" | "headings" | "renumber";

export const sourceLiveKey = new PluginKey<SourceLiveState>("visualtex-document-source");

/** Every marker hidden, the selection's too (the page view printed as it reads), until the selection moves. */
export function hideSourceMarkers(state: EditorState): Transaction {
  return state.tr.setMeta(sourceLiveKey, "away");
}

/** The callouts of a document: per line index, the callout it is in (its type), and the lines that start one. */
function calloutsOf(lines: SourceLine[]): { inside: Map<number, string>; starts: Map<number, { type: string; title: string }> } {
  const inside = new Map<number, string>();
  const starts = new Map<number, { type: string; title: string }>();
  let current: string | null = null;
  lines.forEach((line, i) => {
    const m = /^>\s*\[!(\w+)\][+-]?\s*(.*)$/.exec(line.syntax.text);
    if (m && line.syntax.facts.quote) {
      current = m[1].toLowerCase();
      starts.set(i, { type: current, title: m[2] });
    } else if (!line.syntax.facts.quote) current = null;
    if (current) inside.set(i, current);
  });
  return { inside, starts };
}

function contextOf(read: DocumentLines, index: number, ranges: Ranges, callouts: ReturnType<typeof calloutsOf>): LineContext {
  const block = read.code.find((b) => index >= b.first && index <= b.last) ?? null;
  const blockRange: [number, number] | null = block ? [read.lines[block.first].pos + 1, read.lines[block.last].pos + 1 + read.lines[block.last].node.content.size] : null;
  const line = read.lines[index];
  const previous = index > 0 ? read.lines[index - 1] : null;
  const plain = (f: LineFacts) => !f.blank && !f.list && !f.quote && !f.header && !f.hr && !f.codeBegin && !f.codeInside && !f.codeEnd && !f.indentedCode;
  const continuation = !!previous && previous.pos + previous.node.nodeSize === line.pos && plain(line.syntax.facts) && plain(previous.syntax.facts);
  return { ranges, block, blockRange, callout: callouts.starts.get(index) ?? null, inCallout: callouts.inside.get(index) ?? null, continuation };
}

function allDecorations(doc: Node, read: DocumentLines): DecorationSet {
  const decorations: Decoration[] = [];
  const callouts = calloutsOf(read.lines);
  read.lines.forEach((line, i) => decorations.push(...lineDecorations(line, contextOf(read, i, [], callouts))));
  for (const other of read.others) decorations.push(...alonePieceDecorations(other.node, other.pos, other.syntax, []));
  return DecorationSet.create(doc, decorations);
}

/** The decorations a line's range holds or overlaps (one mapped across a split or a join too), not its neighbours'
 *  line decorations, which only touch its ends. */
function ownDecorations(set: DecorationSet, from: number, to: number): Decoration[] {
  return set.find(from, to).filter((d) => (d.from < to && d.to > from) || (d.from === d.to && d.from > from && d.from < to));
}

/** The base set with the lines the selection reaches drawn as it shows them. */
function shownDecorations(state: EditorState, value: { read: DocumentLines; base: DecorationSet }): DecorationSet {
  const ranges = rangesOf(state);
  const { read } = value;
  let decorations = value.base;
  const redo = new Set<number>();
  read.lines.forEach((line, i) => {
    if (touched(ranges, line.pos + 1, line.pos + 1 + line.node.content.size)) redo.add(i);
  });
  // a code block the selection is in: all its lines (its fences show)
  for (const block of read.code) {
    const from = read.lines[block.first].pos + 1;
    const to = read.lines[block.last].pos + 1 + read.lines[block.last].node.content.size;
    if (touched(ranges, from, to)) {
      redo.add(block.first);
      redo.add(block.last);
    }
  }
  if (redo.size) {
    const callouts = calloutsOf(read.lines);
    for (const i of redo) {
      const line = read.lines[i];
      decorations = decorations.remove(ownDecorations(decorations, line.pos, line.pos + line.node.nodeSize)).add(state.doc, lineDecorations(line, contextOf(read, i, ranges, callouts)));
    }
  }
  for (const other of read.others) {
    if (!touched(ranges, other.pos + 1, other.pos + 1 + other.node.content.size)) continue;
    decorations = decorations.remove(ownDecorations(decorations, other.pos, other.pos + other.node.nodeSize)).add(state.doc, alonePieceDecorations(other.node, other.pos, other.syntax, ranges));
  }
  return decorations;
}

/** The base decorations after a change: lines read otherwise than before drawn again, the others mapped. */
function rebaseDecorations(tr: Transaction, old: SourceLiveState, read: DocumentLines): DecorationSet {
  let base = old.base.map(tr.mapping, tr.doc);
  const before = new Set<BlockSyntax>(old.read.lines.map((l) => l.syntax));
  // a callout or code block's context comes from other lines: when the block structure moved, draw all again
  const shape = (r: DocumentLines) => r.code.map((b) => `${b.first}-${b.last}`).join(",") + "|" + [...calloutsOf(r.lines).inside.keys()].join(",");
  if (shape(old.read) !== shape(read)) return allDecorations(tr.doc, read);
  const callouts = calloutsOf(read.lines);
  read.lines.forEach((line, i) => {
    if (before.has(line.syntax)) return;
    base = base.remove(ownDecorations(base, line.pos, line.pos + line.node.nodeSize)).add(tr.doc, lineDecorations(line, contextOf(read, i, [], callouts)));
  });
  const oldOthers = new Set<BlockSyntax>(old.read.others.map((o) => o.syntax));
  for (const other of read.others) {
    if (oldOthers.has(other.syntax)) continue;
    base = base.remove(ownDecorations(base, other.pos, other.pos + other.node.nodeSize)).add(tr.doc, alonePieceDecorations(other.node, other.pos, other.syntax, []));
  }
  return base;
}

/** A paragraph whose text begins "# " (or underlined) is a heading of that level, a heading without it a paragraph. */
function normalizeHeadings(state: EditorState, lines: SourceLine[]): Transaction | null {
  const { heading, paragraph } = state.schema.nodes;
  if (!heading || !paragraph) return null;
  let tr: Transaction | null = null;
  for (const line of lines) {
    const node = line.node;
    const syntax = line.syntax.heading;
    const pos = tr ? tr.mapping.map(line.pos) : line.pos;
    if (syntax) {
      if (node.type === heading && node.attrs.level === syntax.level && !!node.attrs.unnumbered === syntax.unnumbered) continue;
      tr = (tr ?? state.tr).setNodeMarkup(pos, heading, { ...(node.type === heading ? node.attrs : {}), level: syntax.level, unnumbered: syntax.unnumbered });
    } else if (node.type === heading) {
      tr = (tr ?? state.tr).setNodeMarkup(pos, paragraph, { align: node.attrs.align ?? null, indent: node.attrs.indent ?? 0 });
    }
  }
  return tr;
}

/** Hebrew, Arabic and the scripts near them: their caret moves as the browser lays them out. */
const RTL = /[\u0590-\u08ff\ufb1d-\ufdff\ufe70-\ufeff]/;
let graphemes: Intl.Segmenter | null | undefined;

/** The length of the user-perceived character at the end (or start) of a text. */
function graphemeLength(text: string, atEnd: boolean): number {
  if (graphemes === undefined) graphemes = typeof Intl !== "undefined" && "Segmenter" in Intl ? new Intl.Segmenter(undefined, { granularity: "grapheme" }) : null;
  if (graphemes) {
    const pieces = [...graphemes.segment(atEnd ? text.slice(-32) : text.slice(0, 32))];
    const piece = atEnd ? pieces[pieces.length - 1] : pieces[0];
    return piece ? piece.segment.length : 1;
  }
  const code = text.charCodeAt(atEnd ? text.length - 1 : 0);
  return text.length > 1 && (atEnd ? code >= 0xdc00 && code <= 0xdfff : code >= 0xd800 && code <= 0xdbff) ? 2 : 1;
}

/**
 * Left / Right: one character in the text, markers included (the browser's own step passes over hidden markers, and the
 * markers come into view as the caret reaches them). At a line's end, beside a formula or other object and in
 * right-to-left text the editor's own keys and the browser do it.
 */
function stepCharacter(view: EditorView, dir: -1 | 1, extend: boolean): boolean {
  const { state } = view;
  const selection = state.selection;
  if (!(selection instanceof TextSelection) || (!selection.empty && !extend)) return false;
  const $head = selection.$head;
  if (!$head.parent.isTextblock || $head.parent.type.spec.code) return false;
  const next = dir < 0 ? $head.nodeBefore : $head.nodeAfter;
  if (!next || !next.isText || !next.text) return false;
  const size = graphemeLength(next.text, dir < 0);
  if (RTL.test(dir < 0 ? next.text.slice(-size) : next.text.slice(0, size))) return false;
  const pos = $head.pos + dir * size;
  view.dispatch(state.tr.setSelection(extend ? TextSelection.create(state.doc, selection.anchor, pos) : TextSelection.create(state.doc, pos)).setMeta(sourceLiveKey, "caret").scrollIntoView());
  return true;
}

/**
 * Home / End: to the start / end of the caret's line, a selection collapsed there (Shift: extended), markers included —
 * on a line's first / last row its start / end; on a wrapped line's other rows where the row is drawn to start / end.
 */
function stepLineEdge(view: EditorView, dir: -1 | 1, extend: boolean): boolean {
  const { state } = view;
  const selection = state.selection;
  if (!(selection instanceof TextSelection)) return false;
  const $head = selection.$head;
  if (!$head.parent.isTextblock || $head.parent.type.spec.code) return false;
  let pos: number;
  if (view.endOfTextblock(dir < 0 ? "up" : "down")) pos = dir < 0 ? $head.start() : $head.end();
  else {
    const dom = view.nodeDOM($head.before());
    if (!(dom instanceof HTMLElement)) return false;
    const box = dom.getBoundingClientRect();
    const caret = view.coordsAtPos($head.pos);
    const found = view.posAtCoords({ left: dir < 0 ? box.left + 1 : box.right - 1, top: (caret.top + caret.bottom) / 2 });
    if (!found || found.pos < $head.start() || found.pos > $head.end()) return false;
    pos = found.pos;
  }
  view.dispatch(state.tr.setSelection(extend ? TextSelection.create(state.doc, selection.anchor, pos) : TextSelection.create(state.doc, pos)).setMeta(sourceLiveKey, "caret").scrollIntoView());
  return true;
}

/** Which characters of a textblock show nothing as the view draws it now. */
function hiddenAt(view: EditorView | null, state: EditorState, pos: number, node: Node): Uint8Array {
  const hidden = new Uint8Array(node.content.size);
  const set = sourceLiveKey.getState(state)?.shown;
  if (!set) return hidden;
  for (const d of set.find(pos + 1, pos + 1 + node.content.size)) {
    if ((d.spec as { md?: string }).md !== "hide") continue;
    hidden.fill(1, Math.max(0, d.from - pos - 1), Math.min(node.content.size, d.to - pos - 1));
  }
  void view;
  return hidden;
}

/** Up / Down: the column kept through a run of presses (the caret's x when the run began). */
const goals = new WeakMap<EditorView, { pos: number; x: number }>();

/**
 * Up / Down from a line's first or last row into the line before or after it: to the place under (over) the caret's
 * column, as that line is drawn before its markers show (Obsidian's editor measures the line as drawn), the column kept
 * through a run of presses; a line that shows only a rule or hidden fences: its start when the column is at the left
 * edge, else its end. Within a wrapped line, and next to objects, the browser does it.
 */
function stepLine(view: EditorView, dir: -1 | 1, extend: boolean): boolean {
  const { state } = view;
  const selection = state.selection;
  if (!(selection instanceof TextSelection) || (!selection.empty && !extend)) return false;
  const $head = selection.$head;
  if (!$head.parent.isTextblock || $head.parent.type.spec.code || $head.depth < 1) return false;
  if (!view.endOfTextblock(dir < 0 ? "up" : "down")) return false;
  const container = $head.node($head.depth - 1);
  const index = $head.index($head.depth - 1) + dir;
  const sibling = index >= 0 && index < container.childCount ? container.child(index) : null;
  if (!sibling || !sibling.isTextblock || sibling.type.spec.code) return false;
  const siblingPos = dir < 0 ? $head.before() - sibling.nodeSize : $head.after();
  const dom = view.nodeDOM(siblingPos);
  if (!(dom instanceof HTMLElement)) return false;
  const goal = goals.get(view);
  const caretBox = view.coordsAtPos($head.pos);
  const x = goal && goal.pos === $head.pos ? goal.x : caretBox.left;
  const lineStart = siblingPos + 1;
  const lineEnd = siblingPos + 1 + sibling.content.size;
  let target: number;
  const hidden = hiddenAt(view, state, siblingPos, sibling);
  if (sibling.content.size && hidden.every(Boolean)) {
    // nothing of its text shows: a rule (its line's start, as Obsidian's block rule) or a code block's fence (its
    // start when the column is at the left edge, else its end)
    const box = dom.getBoundingClientRect();
    const left = box.left + parseFloat(getComputedStyle(dom).paddingLeft || "0");
    target = blockSyntax(sibling).facts.hr || x <= left + 2 ? lineStart : lineEnd;
  } else {
    const range = document.createRange();
    range.selectNodeContents(dom);
    const rects = Array.from(range.getClientRects()).filter((rect) => rect.height > 0);
    if (!rects.length) target = lineStart;
    else {
      const row = dir < 0 ? rects[rects.length - 1] : rects[0];
      const found = view.posAtCoords({ left: x, top: (row.top + row.bottom) / 2 });
      if (!found || found.pos < lineStart || found.pos > lineEnd) return false;
      target = found.pos;
    }
  }
  view.dispatch(state.tr.setSelection(extend ? TextSelection.create(state.doc, selection.anchor, target) : TextSelection.create(state.doc, target)).setMeta(sourceLiveKey, "caret").scrollIntoView());
  // (on a rule the column starts again from the line's start, as Obsidian's block rule does)
  const rule = sibling.content.size > 0 && hidden.every(Boolean) && blockSyntax(sibling).facts.hr;
  goals.set(view, { pos: view.state.selection.head, x: rule ? view.coordsAtPos(view.state.selection.head).left : x });
  return true;
}

/**
 * A click, End or Up / Down that put the caret where only hidden markers follow it in its line ("**bold|**" with the
 * markers hidden): the caret goes to the line's end, where it looks to be (Obsidian moves a caret out of a hidden closer).
 */
function caretPastHidden(old: EditorState, state: EditorState): Transaction | null {
  const selection = state.selection;
  if (!(selection instanceof TextSelection) || !selection.empty || old.doc !== state.doc) return null;
  const $head = selection.$head;
  const node = $head.parent;
  if (!node.isTextblock || node.type.spec.code) return null;
  const offset = $head.parentOffset;
  if (offset >= node.content.size) return null;
  const hidden = hiddenAt(null, old, $head.before(), node);
  for (let i = offset; i < node.content.size; i++) if (!hidden[i]) return null;
  return state.tr.setSelection(TextSelection.create(state.doc, $head.end())).setMeta(sourceLiveKey, "caret");
}

/** The source editor: each line's Markdown read and drawn as Obsidian draws it, headings from their "#". */
export function sourceLivePlugin(options: { enabled: () => boolean } = { enabled: () => true }): Plugin<SourceLiveState> {
  return new Plugin<SourceLiveState>({
    key: sourceLiveKey,
    state: {
      init: (_config, state) => {
        const read = readDocument(state.doc);
        const base = allDecorations(state.doc, read);
        return { read, base, shown: base, focused: false, composing: false, pointer: false };
      },
      apply(tr, value, _old, state) {
        const meta = tr.getMeta(sourceLiveKey) as LiveMeta | undefined;
        let { focused, composing, pointer } = value;
        if (meta && typeof meta === "object") {
          if ("focused" in meta) focused = meta.focused;
          if ("composing" in meta) composing = meta.composing;
          if ("pointer" in meta) pointer = meta.pointer;
        }
        const away = meta === "away" ? true : tr.docChanged || tr.selectionSet ? false : value.away;
        let { read, base } = value;
        if (tr.docChanged) {
          read = readDocument(tr.doc);
          base = rebaseDecorations(tr, value, read);
        }
        let shown: DecorationSet;
        // (Obsidian draws nothing again while an input method composes or the mouse button is down: only mapped)
        if ((composing || pointer) && !(meta && typeof meta === "object" && ("composing" in meta || "pointer" in meta))) shown = value.shown.map(tr.mapping, tr.doc);
        else if (away || !focused) shown = base;
        else shown = shownDecorations(state, { read, base });
        return { read, base, shown, focused, composing, pointer, away };
      },
    },
    appendTransaction(transactions, old, state) {
      if (!options.enabled()) return null;
      const value = sourceLiveKey.getState(state);
      if (!value) return null;
      if (transactions.some((tr) => tr.getMeta(sourceLiveKey) === "headings")) return normalizeHeadings(state, value.read.lines);
      if (!transactions.some((tr) => tr.docChanged)) {
        if (!transactions.some((tr) => tr.selectionSet && !tr.getMeta(sourceLiveKey))) return null;
        return caretPastHidden(old, state);
      }
      // (not under an input method's composition: a line changing its type would end it)
      if (value.composing) return null;
      const headings = normalizeHeadings(state, value.read.lines);
      // ordered lists numbered on (not after this plugin's own renumbering)
      if (transactions.some((tr) => tr.getMeta(sourceLiveKey) === "renumber")) return headings;
      const changed: number[] = [];
      for (const tr of transactions) {
        if (!tr.docChanged) continue;
        tr.mapping.maps.forEach((map, k) => {
          const rest = tr.mapping.slice(k + 1);
          map.forEach((_a, _b, from, to) => {
            const $from = state.doc.resolve(Math.min(state.doc.content.size, rest.map(from, -1)));
            const $to = state.doc.resolve(Math.min(state.doc.content.size, rest.map(to, 1)));
            for (let i = $from.index(0); i <= $to.index(0); i++) changed.push(i);
          });
        });
      }
      const base = headings ? state.apply(headings) : state;
      const renumber = renumberLists(old, base, changed);
      if (!renumber) return headings;
      if (!headings) return renumber.setMeta(sourceLiveKey, "renumber");
      for (const step of renumber.steps) headings.step(step);
      return headings.setMeta(sourceLiveKey, "renumber");
    },
    props: {
      decorations(state) {
        if (!options.enabled()) return DecorationSet.empty;
        return sourceLiveKey.getState(state)?.shown ?? DecorationSet.empty;
      },
      handleKeyDown(view, event) {
        if (!options.enabled() || view.composing || event.isComposing) return false;
        if (event.altKey || event.metaKey || event.ctrlKey) return false;
        if (event.key === "ArrowLeft" || event.key === "ArrowRight") return stepCharacter(view, event.key === "ArrowLeft" ? -1 : 1, event.shiftKey);
        if (event.key === "ArrowUp" || event.key === "ArrowDown") return stepLine(view, event.key === "ArrowUp" ? -1 : 1, event.shiftKey);
        // Home in an item: to its text first (Obsidian)
        if (event.key === "Home" && !event.shiftKey && homeInLine(view.state, view.dispatch, view)) return true;
        if (event.key === "Home" || event.key === "End") return stepLineEdge(view, event.key === "Home" ? -1 : 1, event.shiftKey);
        return false;
      },
      // what is pasted (from a web page, a word processor, the app's own copy) comes in as source lines
      transformPasted: (slice) => (options.enabled() ? toSourceSlice(slice) : slice),
      // copied text: the lines as they stand, one per line (as Obsidian copies), objects as their Markdown
      clipboardTextSerializer: (slice, view) => (options.enabled() ? sourceSliceText(slice, view.state.schema) : ""),
      // pasted text: its lines as lines, blank ones too (formulas, tables, pictures as objects)
      clipboardTextParser: (text, $context, plain, view) => {
        if (!options.enabled() || $context.parent.type.spec.code) return null as unknown as Slice;
        void plain;
        const doc = sourceTextToDocument(view.state.schema, text);
        const lone = doc.childCount === 1 && doc.firstChild!.isTextblock;
        return lone ? new Slice(doc.firstChild!.content, 0, 0) : new Slice(doc.content, doc.firstChild!.isTextblock ? 1 : 0, doc.lastChild!.isTextblock ? 1 : 0);
      },
      handleDOMEvents: {
        focus: (view) => {
          view.dispatch(view.state.tr.setMeta(sourceLiveKey, { focused: true } satisfies LiveMeta).setMeta("addToHistory", false));
          return false;
        },
        blur: (view) => {
          view.dispatch(view.state.tr.setMeta(sourceLiveKey, { focused: false } satisfies LiveMeta).setMeta("addToHistory", false));
          return false;
        },
        mousedown: (view, event) => {
          if (event.button !== 0) return false;
          view.dispatch(view.state.tr.setMeta(sourceLiveKey, { pointer: true } satisfies LiveMeta).setMeta("addToHistory", false));
          const up = () => {
            window.removeEventListener("mouseup", up, true);
            if (!view.isDestroyed) view.dispatch(view.state.tr.setMeta(sourceLiveKey, { pointer: false } satisfies LiveMeta).setMeta("addToHistory", false));
          };
          window.addEventListener("mouseup", up, true);
          return false;
        },
        compositionstart: (view) => {
          view.dispatch(view.state.tr.setMeta(sourceLiveKey, { composing: true } satisfies LiveMeta).setMeta("addToHistory", false));
          return false;
        },
        compositionend: (view) => {
          // (the committed text is read a moment after the event)
          setTimeout(() => {
            if (view.isDestroyed) return;
            view.dispatch(view.state.tr.setMeta(sourceLiveKey, { composing: false } satisfies LiveMeta).setMeta("addToHistory", false));
            if (options.enabled()) view.dispatch(view.state.tr.setMeta(sourceLiveKey, "headings").setMeta("addToHistory", false));
          }, 30);
          return false;
        },
      },
    },
    view: (view) => {
      // the editor may have the focus before the plugin sees a focus event
      if (view.hasFocus()) queueMicrotask(() => !view.isDestroyed && view.dispatch(view.state.tr.setMeta(sourceLiveKey, { focused: true } satisfies LiveMeta).setMeta("addToHistory", false)));
      return {};
    },
  });
}

/** The document's lines as the source editor read them (for the keys that edit lines). */
export function sourceLines(state: EditorState): SourceLine[] {
  return sourceLiveKey.getState(state)?.read.lines ?? [];
}

/** Whether a position is in a code span's text or a code block line: what is typed there stays as typed. */
export function inSourceCode($pos: ResolvedPos): boolean {
  const node = $pos.parent;
  if (!node.isTextblock || node.type.spec.code) return false;
  const syntax = blockSyntax(node);
  if (syntax.facts.codeInside || syntax.facts.indentedCode) return true;
  const offset = $pos.parentOffset;
  return syntax.units.some((unit) => unit.kind === "code" && unit.marks.length > 1 && offset >= unit.marks[0][1] && offset <= unit.marks[unit.marks.length - 1][0]);
}

// ------------------------------------------------------------------------------------------------- commands

const MARKERS: Record<"strong" | "em" | "strike" | "highlight" | "code" | "underline", [string, string]> = {
  strong: ["**", "**"],
  em: ["*", "*"],
  strike: ["~~", "~~"],
  highlight: ["==", "=="],
  code: ["`", "`"],
  underline: ["<u>", "</u>"],
};

/** The smallest span of a kind around [from, to] (its markers included) in a block. */
function unitAround(node: Node, from: number, to: number, kind: string): SourceUnit | null {
  let best: SourceUnit | null = null;
  for (const unit of blockSyntax(node).units) {
    if (unit.kind !== kind || from < unit.from || to > unit.to || unit.marks.length < 2) continue;
    if (!best || unit.to - unit.from < best.to - best.from) best = unit;
  }
  return best;
}

/** Whether the selection is in a span of a kind (the toolbar's pressed buttons). */
export function sourceFormatActive(state: EditorState, kind: string): boolean {
  const { $from, $to } = state.selection;
  if (!$from.parent.isTextblock || !$from.sameParent($to)) return false;
  return !!unitAround($from.parent, $from.parentOffset, $to.parentOffset, kind);
}

const WORD = /[\p{L}\p{N}_]/u;

/** The word the caret is in or at an end of (letters, digits, "_"; a run of Chinese text is one), or null. */
function wordAround(text: string, offset: number): [number, number] | null {
  let from = offset;
  while (from > 0 && WORD.test(text[from - 1])) from--;
  let to = offset;
  while (to < text.length && WORD.test(text[to])) to++;
  return to > from ? [from, to] : null;
}

/**
 * Ctrl+B, Ctrl+I, …, as Obsidian's: with a selection, its text between the markers (the spaces at its ends left out,
 * the selection kept on the text), or the markers taken off the span of that kind it is in. With the caret: in a span
 * of that kind, past its closing marker when the caret is just before it, else the span's markers taken off; in a word,
 * the word between the markers; elsewhere the markers with the caret between them.
 */
export function toggleSourceFormat(kind: keyof typeof MARKERS): Command {
  return (state, dispatch) => {
    const { $from, $to, empty } = state.selection;
    if (!$from.parent.isTextblock || !$from.sameParent($to) || $from.parent.type.spec.code) return false;
    const node = $from.parent;
    const start = $from.start();
    const [open, close] = MARKERS[kind];
    const from = $from.pos - start;
    const to = $to.pos - start;
    const unit = unitAround(node, from, to, kind);
    if (unit) {
      if (!dispatch) return true;
      const [o1, o2] = unit.marks[0];
      const [c1, c2] = unit.marks[unit.marks.length - 1];
      if (empty && from === c1 && c1 > o2) {
        dispatch(state.tr.setSelection(TextSelection.create(state.doc, start + c2)).scrollIntoView());
        return true;
      }
      // taken off: the markers go, the selection stays on the same text
      dispatch(state.tr.delete(start + c1, start + c2).delete(start + o1, start + o2).scrollIntoView());
      return true;
    }
    if (!dispatch) return true;
    const text = blockSyntax(node).text;
    if (empty) {
      const word = wordAround(text, from);
      const tr = word ? state.tr.insertText(close, start + word[1]).insertText(open, start + word[0]) : state.tr.insertText(open + close, $from.pos);
      tr.setSelection(TextSelection.create(tr.doc, $from.pos + open.length));
      dispatch(tr.scrollIntoView());
      return true;
    }
    // the selected text without the spaces at its ends
    const selected = text.slice(from, to);
    const lead = selected.length - selected.trimStart().length;
    const trail = selected.length - selected.trimEnd().length;
    const a = start + from + lead;
    const b = start + to - trail;
    if (b <= a) return false;
    const tr = state.tr.insertText(close, b).insertText(open, a);
    const forward = state.selection.head >= state.selection.anchor;
    const inner = [a + open.length, b + open.length];
    tr.setSelection(TextSelection.create(tr.doc, forward ? inner[0] : inner[1], forward ? inner[1] : inner[0]));
    dispatch(tr.scrollIntoView());
    return true;
  };
}

const STYLE_UNITS = new Set(["strong", "em", "strike", "highlight", "code", "underline", "superscript", "subscript", "kbd", "html"]);

/**
 * Ctrl+\ (clear format): the markers of the formatted spans the selection takes or touches (the caret: the spans it is
 * in) taken out, their text kept; links, headings and escapes stay.
 */
export const clearSourceFormat: Command = (state, dispatch) => {
  const { from, to } = state.selection;
  const ranges: Array<[number, number]> = [];
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isTextblock) return true;
    if (node.type.spec.code) return false;
    const start = pos + 1;
    for (const unit of blockSyntax(node).units) {
      if (!STYLE_UNITS.has(unit.kind) || start + unit.to < from || start + unit.from > to) continue;
      for (const [a, b] of unit.marks) if (b > a) ranges.push([start + a, start + b]);
    }
    return false;
  });
  if (!ranges.length) return false;
  if (dispatch) {
    const tr = state.tr;
    for (const [a, b] of ranges.sort((x, y) => y[0] - x[0])) tr.delete(a, b);
    dispatch(tr.scrollIntoView());
  }
  return true;
};

/** Ctrl+K (Obsidian's "Insert Markdown link"): "[selection](|)", or "[|]()" with no selection. */
export const insertSourceLink: Command = (state, dispatch) => {
  const { $from, $to, empty } = state.selection;
  if (!$from.parent.isTextblock || !$from.sameParent($to) || $from.parent.type.spec.code) return false;
  if (!dispatch) return true;
  if (empty) {
    const tr = state.tr.insertText("[]()", $from.pos);
    dispatch(tr.setSelection(TextSelection.create(tr.doc, $from.pos + 1)).scrollIntoView());
    return true;
  }
  const tr = state.tr.insertText("]()", $to.pos).insertText("[", $from.pos);
  dispatch(tr.setSelection(TextSelection.create(tr.doc, $to.pos + 3)).scrollIntoView());
  return true;
};

/** Ctrl+1 … Ctrl+6 (Ctrl+0: body text): the blocks' "#" prefixes set to that level (taken off at that level already). */
export function setSourceHeading(level: number): Command {
  return (state, dispatch) => {
    const blocks: Array<{ node: Node; pos: number }> = [];
    state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
      if (node.isTextblock) {
        if (headingCapable(node)) blocks.push({ node, pos });
        return false;
      }
      return true;
    });
    if (!blocks.length) return false;
    if (!dispatch) return true;
    let tr = state.tr;
    for (const { node, pos } of blocks.reverse()) {
      const syntax = headingSyntax(blockText(node));
      const at = tr.mapping.map(pos) + 1;
      const current = syntax ? syntax.prefixEnd : 0;
      // the level it has already: body text again (as Ctrl+1 on a level-1 section)
      const target = syntax?.level === level ? 0 : level;
      const prefix = target > 0 ? "#".repeat(target) + " " : "";
      tr = prefix ? tr.insertText(prefix, at, at + current) : current ? tr.delete(at, at + current) : tr;
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Tab / Shift+Tab on a heading: one level deeper / higher, its "#" prefix rewritten (levels 1 to 6). */
export function shiftSourceHeading(delta: 1 | -1): Command {
  return (state, dispatch) => {
    const { $from } = state.selection;
    const node = $from.parent;
    if (!node.isTextblock || !headingCapable(node)) return false;
    const syntax = headingSyntax(blockText(node));
    if (!syntax) return false;
    const level = Math.max(1, Math.min(6, syntax.level + delta));
    if (dispatch && level !== syntax.level) {
      const start = $from.start();
      dispatch(state.tr.insertText("#".repeat(level), start, start + syntax.level).scrollIntoView());
    }
    return true;
  };
}

// ------------------------------------------------------------------------------------------------- old documents, exports

/** The default highlight (==text==): other colours are the document's own marks. */
const MARKDOWN_HIGHLIGHT = "#fff59d";
const SOURCE_MARKS = ["link", "strong", "em", "strike", "code", "underline", "superscript", "subscript", "kbd", "highlight"];

/** A document mark as Markdown source writes it, or null (a colour, a size, another highlight: marks that stay). */
function inlineMarkOf(mark: Mark): InlineMark | null {
  const name = mark.type.name;
  if (name === "link") return { kind: "link", href: mark.attrs.href, title: mark.attrs.title ?? null };
  if (name === "highlight") return mark.attrs.color === MARKDOWN_HIGHLIGHT ? { kind: "highlight" } : null;
  return SOURCE_MARKS.includes(name) ? { kind: name as InlineMark["kind"] } : null;
}

/**
 * A document as source lines: each textblock's formatting written as markers in its text (other marks — colour, size —
 * stay marks), each heading given its "#" prefix, and the blocks Obsidian keeps as text made lines: a list's items
 * "- " / "1. " / "- [ ] " (nested ones a tab deeper), a quote's lines "> " (an alert's first "> [!NOTE]"), a code block
 * between "```lang" fences, a rule "---", line breaks new lines. Objects (formulas, tables, pictures, the title, page
 * commands) stay. Already source lines: unchanged.
 */
export function toSourceDocument(doc: Node): Node {
  const schema = doc.type.schema;
  const out: Node[] = [];
  let previous: Node | null = null;
  doc.forEach((block) => {
    const lines = linesOf(block, schema, "", "");
    // what a converted list, quote or code block ends: a blank line before the next text (else it would be read as
    // part of them)
    if (previous && CONTAINER.has(previous.type.name) && lines[0]?.isTextblock && lines[0].content.size) out.push(schema.nodes.paragraph.create());
    out.push(...lines);
    previous = block;
  });
  if (!out.length) out.push(schema.nodes.paragraph.create());
  const converted = doc.type.create(doc.attrs, out, doc.marks);
  return converted.eq(doc) ? doc : converted;
}

/** Blocks that become lines of text. */
const CONTAINER = new Set(["bullet_list", "ordered_list", "blockquote", "code_block", "horizontal_rule"]);

/** A document's lines given the node type their text says (a "#" line a heading; read as the document reads them). */
export function withHeadings(doc: Node): Node {
  const { heading, paragraph } = doc.type.schema.nodes;
  const read = readDocument(doc);
  const byNode = new Map(read.lines.map((l) => [l.node, l.syntax.heading]));
  let changed = false;
  const children: Node[] = [];
  doc.forEach((node) => {
    if (!byNode.has(node)) {
      children.push(node);
      return;
    }
    const h = byNode.get(node);
    if (h && (node.type !== heading || node.attrs.level !== h.level || !!node.attrs.unnumbered !== h.unnumbered)) {
      children.push(heading.create({ ...(node.type === heading ? node.attrs : {}), level: h.level, unnumbered: h.unnumbered }, node.content, node.marks));
      changed = true;
    } else if (!h && node.type === heading) {
      children.push(paragraph.create(null, node.content, node.marks));
      changed = true;
    } else children.push(node);
  });
  return changed ? doc.type.create(doc.attrs, children, doc.marks) : doc;
}

/** Formulas typed in a line of Markdown text as `$…$` (Pandoc's rule: no space inside the dollars, no digit after). */
const INLINE_MATH = /(?<![\\$])\$(?!\$)([^\s$](?:[^$\n]*?[^\s$\\])?)\$(?![$\d])/g;

/**
 * Markdown text as the source editor's document: a line of text per line (headings typed per their "#"); display
 * formulas (`$$` lines), tables (a header and a delimiter row) and lines of only pictures become the document's
 * objects as `readBlock` (Markdown read into the document's nodes) gives them; `$…$` in a line a formula.
 */
export function sourceTextToDocument(schema: Schema, text: string, readBlock: (markdown: string) => Node = (markdown) => readMarkdown(schema, markdown).doc): Node {
  const rows = text.replace(/\r\n?/g, "\n").split("\n");
  if (rows.length > 1 && rows[rows.length - 1] === "") rows.pop();
  const out: Node[] = [];
  let fence: string | null = null;
  const objects = (markdown: string) => {
    const read = readBlock(markdown);
    const nodes: Node[] = [];
    read.forEach((node) => nodes.push(node));
    return nodes;
  };
  const lineNode = (row: string) => {
    if (!row) return schema.nodes.paragraph.create();
    const children: Node[] = [];
    let at = 0;
    if (!fence) {
      for (const m of row.matchAll(INLINE_MATH)) {
        // (not in a code span)
        const before = row.slice(0, m.index!);
        if (((before.match(/`/g) ?? []).length % 2) === 1) continue;
        if (m.index! > at) children.push(schema.text(row.slice(at, m.index!)));
        children.push(schema.nodes.math_inline.create({ latex: m[1] }));
        at = m.index! + m[0].length;
      }
    }
    if (at < row.length) children.push(schema.text(row.slice(at)));
    return schema.nodes.paragraph.create(null, children);
  };
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const fenceMark = /^\s*(`{3,}|~{3,})/.exec(row)?.[1];
    if (fence) {
      if (fenceMark && fenceMark[0] === fence[0] && fenceMark.length >= fence.length && !row.trim().slice(fenceMark.length).trim()) fence = null;
      out.push(lineNode(row));
      continue;
    }
    if (fenceMark) {
      fence = fenceMark;
      out.push(lineNode(row));
      continue;
    }
    // a display formula: "$$" … "$$" (or one line "$$…$$")
    if (/^\s*\$\$/.test(row)) {
      let j = i;
      if (!(row.trim().length > 4 && row.trim().endsWith("$$"))) {
        j = i + 1;
        while (j < rows.length && !rows[j].trim().endsWith("$$")) j++;
      }
      if (j < rows.length) {
        const nodes = objects(rows.slice(i, j + 1).join("\n"));
        if (nodes.length === 1 && nodes[0].type.name === "math_display") {
          out.push(nodes[0]);
          i = j;
          continue;
        }
      }
    }
    // a table: a row of cells and a delimiter row under it, then its rows; its caption above it ("Table: …", Pandoc's,
    // a blank line between or not), or below it, or the note that it has none (markdown.ts: a caption between two tables
    // is the second's unless that one has its own after it)
    const captionAbove = captionLabelLength(row, "table") > 0 ? tableAfter(rows, i + 1) : -1;
    const tableStart = tableAt(rows, i) ? i : captionAbove;
    if (tableStart >= 0) {
      let j = tableEnd(rows, tableStart);
      if (captionAbove < 0) {
        const k = rows[j] === "" ? j + 1 : j;
        const following = k < rows.length ? tableAfter(rows, k + 1) : -1;
        const below = k < rows.length && captionLabelLength(rows[k], "table") > 0 && (following < 0 || captionAfter(rows, tableEnd(rows, following)));
        if (below || (k < rows.length && NO_CAPTION.test(rows[k]))) j = k + 1;
      }
      const nodes = objects(rows.slice(i, j).join("\n"));
      if (nodes.length && nodes.every((n) => !n.isTextblock)) {
        out.push(...nodes);
        i = j - 1;
        continue;
      }
    }
    // a line of pictures
    if (/^\s*(!\[[^\]]*\]\([^)]*\)\s*)+$/.test(row)) {
      const nodes = objects(row);
      if (nodes.length === 1 && (!nodes[0].isTextblock || nodes[0].childCount && [...Array(nodes[0].childCount).keys()].every((k) => nodes[0].child(k).type.name === "image"))) {
        out.push(nodes[0]);
        continue;
      }
    }
    out.push(lineNode(row));
  }
  if (!out.length) out.push(schema.nodes.paragraph.create());
  return withHeadings(schema.topNodeType.create(null, out));
}

/**
 * A copied slice's text in the source editor: its lines as they stand, one per line, objects written as the source
 * export writes them (a blank line on each side); "" when the slice is not lines (cells of a table: the editor's own
 * copy then).
 */
function sourceSliceText(slice: Slice, schema: Schema): string {
  const blocks: Node[] = [];
  let run: Node[] = [];
  const flush = () => {
    if (run.length) blocks.push(schema.nodes.paragraph.create(null, run));
    run = [];
  };
  slice.content.forEach((node) => {
    if (node.isInline) run.push(node);
    else {
      flush();
      blocks.push(node);
    }
  });
  flush();
  const content = Fragment.from(blocks);
  if (!blocks.length || !schema.topNodeType.validContent(content)) return "";
  return documentToMarkdown(schema.topNodeType.create(null, content), { sourceText: true }).replace(/\n$/, "");
}

/** A table starts at row `i` of a text: a row of cells and a delimiter row under it. */
function tableAt(rows: string[], i: number): boolean {
  return i + 1 < rows.length && /^\s*\|/.test(rows[i]) && /^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$/.test(rows[i + 1]) && rows[i + 1].includes("-");
}

/** Where a table starts at row `i` or after one blank row there, or -1. */
function tableAfter(rows: string[], i: number): number {
  return tableAt(rows, i) ? i : rows[i] === "" && tableAt(rows, i + 1) ? i + 1 : -1;
}

/** The row after a table that starts at row `start`. */
function tableEnd(rows: string[], start: number): number {
  let j = start + 2;
  while (j < rows.length && /^\s*\|/.test(rows[j])) j++;
  return j;
}

/** A table's own caption or the note that it has none follows at row `j` (a blank row between or not). */
function captionAfter(rows: string[], j: number): boolean {
  const k = rows[j] === "" ? j + 1 : j;
  return k < rows.length && (captionLabelLength(rows[k], "table") > 0 || NO_CAPTION.test(rows[k]));
}

/** A paragraph-based document (saved before the line model, read from Markdown): a blank line between paragraphs. */
export function separateParagraphs(doc: Node): Node {
  const schema = doc.type.schema;
  const out: Node[] = [];
  let previous: Node | null = null;
  doc.forEach((block) => {
    if (previous && previous.type.name === "paragraph" && block.type.name === "paragraph" && previous.content.size && block.content.size) out.push(schema.nodes.paragraph.create());
    out.push(block);
    previous = block;
  });
  return doc.type.create(doc.attrs, out, doc.marks);
}

/** A line: a paragraph (or heading) of `prefix` and the inline source of `content`. */
function line(schema: Schema, prefix: string, content: Fragment, attrs: Record<string, unknown> | null = null): Node {
  const children: Node[] = [];
  if (prefix) children.push(schema.text(prefix));
  content.forEach((child) => children.push(child));
  return schema.nodes.paragraph.create(attrs, children);
}

/** A block as lines; `first` is the first line's prefix (a list item's marker), `rest` the others'. */
function linesOf(block: Node, schema: Schema, rest: string, first: string): Node[] {
  const name = block.type.name;
  if (name === "paragraph" || name === "heading") {
    const source = sourceBlock(block, schema);
    if (name === "heading" && !first) return [source];
    // line breaks: new lines
    const out: Node[] = [];
    let piece: Node[] = [];
    let prefix = first;
    const keep = name === "paragraph" ? { align: block.attrs.align ?? null, indent: block.attrs.indent ?? 0 } : null;
    const flush = () => {
      out.push(line(schema, prefix, Fragment.fromArray(piece), keep));
      piece = [];
      prefix = rest;
    };
    source.forEach((child) => {
      if (child.type.name === "hard_break") flush();
      else piece.push(child);
    });
    flush();
    return out;
  }
  if (name === "bullet_list" || name === "ordered_list") {
    const out: Node[] = [];
    const startAt = Number(block.attrs.order ?? 1) || 1;
    block.forEach((item, _offset, index) => {
      const marker = name === "bullet_list" ? "- " : `${startAt + index}. `;
      const checked = item.attrs.checked;
      const task = checked === null || checked === undefined ? "" : checked ? "[x] " : "[ ] ";
      item.forEach((child, _o, k) => {
        out.push(...linesOf(child, schema, rest + "\t", k === 0 ? (index === 0 ? first : rest) + marker + task : rest + "\t"));
      });
      if (!item.childCount) out.push(line(schema, (index === 0 ? first : rest) + marker + task, Fragment.empty));
    });
    return out;
  }
  if (name === "blockquote") {
    const out: Node[] = [];
    const quote = rest + "> ";
    if (block.attrs.alert) out.push(line(schema, (first || rest) + "> ", Fragment.from(schema.text(`[!${String(block.attrs.alert).toUpperCase()}]`))));
    block.forEach((child, _o, k) => {
      if (k > 0 && child.type.name === "paragraph") out.push(line(schema, rest + ">", Fragment.empty));
      out.push(...linesOf(child, schema, quote, k === 0 && !block.attrs.alert ? (first || rest) + "> " : quote));
    });
    return out;
  }
  if (name === "code_block") {
    const lang = String(block.attrs.language ?? "");
    const out = [line(schema, first + "```" + lang, Fragment.empty)];
    for (const text of block.textContent.split("\n")) out.push(line(schema, rest, text ? Fragment.from(schema.text(text)) : Fragment.empty));
    out.push(line(schema, rest + "```", Fragment.empty));
    return out;
  }
  if (name === "horizontal_rule") return [line(schema, first + "---", Fragment.empty)];
  // an object: its own textblocks (table cells, captions) as source; itself stays
  if (block.isLeaf) return [block];
  if (block.isTextblock) return [sourceBlock(block, schema)];
  const convert = (node: Node): Node => {
    if (node.isTextblock) return sourceBlock(node, schema);
    if (node.isLeaf) return node;
    const children: Node[] = [];
    node.forEach((child) => children.push(convert(child)));
    return node.copy(Fragment.fromArray(children));
  };
  return [convert(block)];
}

function sourceBlock(node: Node, schema: Schema): Node {
  if (node.type.spec.code) return node;
  let marked = false;
  node.forEach((child) => {
    if (child.marks.some((mark) => inlineMarkOf(mark) || mark.type.name === "md_syntax")) marked = true;
  });
  const isHeading = node.type.name === "heading";
  const syntax = isHeading ? headingSyntax(blockText(node)) : null;
  const headingPrefix = isHeading && !syntax ? "#".repeat(Math.min(6, Math.max(1, node.attrs.level ?? 1))) + " " : "";
  // an unnumbered section says so in its text
  const headingSuffix = isHeading && node.attrs.unnumbered && !syntax?.unnumbered ? " {-}" : "";
  if (!marked && !headingPrefix && !headingSuffix) return node;
  // the block's units: each character (an object one) with its source marks, the others kept aside
  const units: WriteUnit[] = [];
  const objects: Node[] = [];
  const others: Array<readonly Mark[]> = [];
  node.forEach((child) => {
    // (the markers the old editor showed around the caret: never part of the text)
    if (child.marks.some((mark) => mark.type.name === "md_syntax")) return;
    const sourceMarks: InlineMark[] = [];
    const rest: Mark[] = [];
    for (const mark of child.marks) {
      const inline = inlineMarkOf(mark);
      if (inline) sourceMarks.push(inline);
      else rest.push(mark);
    }
    if (child.isText) {
      for (const ch of child.text ?? "") {
        units.push({ ch, marks: sourceMarks });
        others.push(rest);
      }
    } else {
      units.push({ ch: OBJECT, marks: sourceMarks });
      objects.push(child.mark(rest));
      others.push(rest);
    }
  });
  const written = writeInline(units, { escape: true });
  // the written text, back into nodes: objects where their characters went, other marks on their characters
  const content: Node[] = [];
  if (headingPrefix) content.push(schema.text(headingPrefix));
  let objectIndex = 0;
  const atOf = new Map<number, number>();
  written.at.forEach((at, index) => atOf.set(at, index));
  let run = "";
  let runMarks: readonly Mark[] = [];
  const flush = () => {
    if (run) content.push(schema.text(run, runMarks));
    run = "";
  };
  for (let i = 0; i < written.text.length; i++) {
    const index = atOf.get(i);
    const ch = written.text[i];
    const marks = index !== undefined ? others[index] : [];
    if (index !== undefined && units[index].ch === OBJECT) {
      flush();
      content.push(objects[objectIndex++]);
      continue;
    }
    if (!sameMarkSet(marks, runMarks)) flush();
    runMarks = marks;
    run += ch;
  }
  flush();
  if (headingSuffix) content.push(schema.text(headingSuffix));
  return node.type.create(node.attrs, content, node.marks);
}

/** Pasted content as source lines (toSourceDocument for a slice: lists, quotes, code become lines). */
export function toSourceSlice(slice: Slice): Slice {
  const first = slice.content.firstChild;
  if (!first) return slice;
  const schema = first.type.schema;
  // inline content alone: one line's
  if (slice.content.firstChild?.isInline) {
    const lineNode = sourceBlock(schema.nodes.paragraph.create(null, slice.content), schema);
    return new Slice(lineNode.content, 0, 0);
  }
  const out: Node[] = [];
  slice.content.forEach((block) => out.push(...linesOf(block, schema, "", "")));
  // (a slice opened into a list item or quote: its lines are whole blocks now)
  const flat = out.every((n) => n.isTextblock);
  return new Slice(Fragment.from(out), flat && slice.openStart > 0 ? 1 : 0, flat && slice.openEnd > 0 ? 1 : 0);
}

function sameMarkSet(a: readonly Mark[], b: readonly Mark[]): boolean {
  return a.length === b.length && a.every((mark, i) => mark.eq(b[i]));
}

/**
 * The document as its formatting reads (for exports, printing, copying as rich text, the outline): each textblock's
 * markers taken out and its formatting made marks, headings without their "#".
 */
export function renderedDocument(doc: Node): Node {
  const schema = doc.type.schema;
  const convert = (node: Node): Node => {
    if (node.isTextblock) return renderedBlock(node, schema);
    if (node.isLeaf) return node;
    const children: Node[] = [];
    node.forEach((child) => children.push(convert(child)));
    return node.copy(Fragment.fromArray(children));
  };
  return convert(doc);
}

/** The characters of a block that are not part of what it reads as (markers, addresses), and where a <br> was. */
function renderedHidden(syntax: BlockSyntax): { hidden: Uint8Array; breaks: Set<number> } {
  const hidden = new Uint8Array(syntax.text.length);
  const breaks = new Set<number>();
  for (const unit of syntax.units) {
    for (const [from, to] of unit.marks) hidden.fill(1, from, to);
    // a link's address goes; a picture's description stays (the picture itself is not the document's)
    for (const [from, to] of unit.plain ?? []) if (!(unit.text && unit.text[0] === from && unit.text[1] === to)) hidden.fill(1, from, to);
    if (unit.widget?.kind === "break") breaks.add(unit.from);
  }
  return { hidden, breaks };
}

/** A textblock's content as it reads, in order: its text without the markers, and its objects (at their offsets). */
export function shownContent(node: Node): Array<{ offset: number; text: string } | { offset: number; node: Node }> {
  const { hidden } = renderedHidden(blockSyntax(node));
  const out: Array<{ offset: number; text: string } | { offset: number; node: Node }> = [];
  let offset = 0;
  node.forEach((child) => {
    if (child.isText) {
      const text = child.text ?? "";
      let shown = "";
      for (let i = 0; i < text.length; i++) if (!hidden[offset + i]) shown += text[i];
      if (shown) out.push({ offset, text: shown });
      offset += text.length;
    } else {
      if (!hidden[offset]) out.push({ offset, node: child });
      offset += child.nodeSize;
    }
  });
  return out;
}

export function renderedBlock(node: Node, schema: Schema): Node {
  const syntax = blockSyntax(node);
  if (!syntax.units.length && !syntax.styles.length) return node;
  // where a <br> was: a line break
  const { hidden, breaks } = renderedHidden(syntax);
  const marksOf = (style: SourceStyle): Mark[] => {
    if (style.kind === "color") {
      const out: Mark[] = [];
      if (style.color && schema.marks.textColor) out.push(schema.marks.textColor.create({ color: style.color }));
      if (style.background && schema.marks.highlight) out.push(schema.marks.highlight.create({ color: style.background }));
      return out;
    }
    const type = style.kind === "url" ? null : schema.marks[style.kind];
    if (!type) return [];
    if (style.kind === "link") return [type.create({ href: style.href ?? "", title: null })];
    if (style.kind === "highlight") return [type.create({ color: MARKDOWN_HIGHLIGHT })];
    return [type.create()];
  };
  const content: Node[] = [];
  let offset = 0;
  node.forEach((child) => {
    const size = child.isText ? (child.text ?? "").length : 1;
    if (!child.isText) {
      if (!hidden[offset]) content.push(child);
      offset += size;
      return;
    }
    const text = child.text ?? "";
    let run = "";
    let runMarks: readonly Mark[] | null = null;
    const flush = () => {
      if (run) content.push(schema.text(run, runMarks ?? child.marks));
      run = "";
    };
    for (let i = 0; i < text.length; i++) {
      const at = offset + i;
      if (breaks.has(at) && schema.nodes.hard_break) {
        flush();
        content.push(schema.nodes.hard_break.create());
      }
      if (hidden[at]) continue;
      let marks = child.marks;
      for (const style of syntax.styles) {
        if (at >= style.from && at < style.to) for (const mark of marksOf(style)) marks = mark.addToSet(marks);
      }
      if (runMarks && !sameMarkSet(marks, runMarks)) flush();
      runMarks = marks;
      run += text[i];
    }
    flush();
    offset += size;
  });
  return node.type.create(node.attrs, content, node.marks);
}

/** A block's text as it reads (markers out): the outline, numbering and search. */
export function renderedText(node: Node): string {
  return node.isTextblock ? renderedBlock(node, node.type.schema).textContent : node.textContent;
}
