// The document's text is its Markdown source (2026-11-10): a source editor as Obsidian's Live Preview is one. The text
// keeps every marker ("**", "*", "~~", "==", "`", "[…](…)", "\", a heading's "# "): what the text means is read from it
// after every change, with the parser Obsidian's editor family uses (@lezer/markdown), and drawn with decorations. A
// formatted span shows its markers while the selection touches it (its ends included) and hides them otherwise; a heading
// shows its "#" while the selection is in its line; a paragraph whose text begins "#" is a heading, and stops being one
// when the "#" goes. Nothing is converted one way: "**" text "**", "****" with text typed between, "**x**" edited into
// "*x*" all read as they stand. Formulas stay the document's own (inline atoms; "$$" opens one, formulas.ts), and so do
// the blocks that are nodes (lists, quotes, code, tables, figures).
//
// Emphasis follows CommonMark with one change for Chinese, Japanese and Korean text: next to a CJK character, punctuation
// does not stop a run of "*" or "_" from opening or closing ("前**粗体。**后" is bold; CommonMark and Obsidian's editor
// both get it wrong).

import { parser as baseParser, InlineContext, type DelimiterType, type Element as MdElement, type MarkdownConfig } from "@lezer/markdown";
import { Fragment, Slice, type Mark, type Node, type ResolvedPos, type Schema } from "prosemirror-model";
import { Plugin, PluginKey, TextSelection, type Command, type EditorState, type Transaction } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";
import { OBJECT, writeInline, type InlineMark, type WriteUnit } from "./inlineMarkdown.ts";

// ------------------------------------------------------------------------------------------------- the parser

const PUNCTUATION = (() => {
  try {
    return new RegExp("[\\p{S}\\p{P}]", "u");
  } catch {
    return /[!"#$%&'()*+,\-./:;<=>?@[\\\]^_`{|}~\xA1‐-‧]/;
  }
})();
const CJK = /[ᄀ-ᇿ⺀-⿿　-〿぀-ヿ㄀-ㇿ㈀-㏿㐀-䶿一-鿿ꥠ-꥿가-퟿豈-﫿︰-﹏＀-￯]|[\ud840-\ud87f]/;

/** Whether a run of delimiter characters at [from, to) may open and close (CommonMark's flanking, CJK-friendly). */
function flanking(cx: InlineContext, from: number, to: number) {
  const before = cx.slice(from - 1, from);
  const after = cx.slice(to, to + 1);
  // next to CJK text, punctuation counts as a letter would
  const cjk = CJK.test(before) || CJK.test(after);
  const pBefore = !cjk && PUNCTUATION.test(before);
  const pAfter = !cjk && PUNCTUATION.test(after);
  const sBefore = /\s|^$/.test(before);
  const sAfter = /\s|^$/.test(after);
  const left = !sAfter && (!pAfter || sBefore || pBefore);
  const right = !sBefore && (!pBefore || sAfter || pAfter);
  return { left, right, pBefore, pAfter };
}

const Emphasis: MarkdownConfig = {
  remove: ["Emphasis"],
  parseInline: [
    {
      name: "SourceEmphasis",
      parse(cx, next, start) {
        if (next !== 95 && next !== 42) return -1;
        let pos = start + 1;
        while (cx.char(pos) === next) pos++;
        const { left, right, pBefore, pAfter } = flanking(cx, start, pos);
        const canOpen = left && (next === 42 || !right || pBefore);
        const canClose = right && (next === 42 || !left || pAfter);
        return cx.addDelimiter(next === 95 ? InlineContext.emphasisUnderscore : InlineContext.emphasisAsterisk, start, pos, canOpen, canClose);
      },
    },
  ],
};

/** A pair of two-character delimiters ("~~", "=="), as GFM's strikethrough reads them (CJK-friendly). */
function pairedDelimiter(name: string, code: number): MarkdownConfig {
  const type: DelimiterType = { resolve: name, mark: `${name}Mark` };
  return {
    defineNodes: [{ name }, { name: `${name}Mark` }],
    parseInline: [
      {
        name,
        parse(cx, next, pos) {
          if (next !== code || cx.char(pos + 1) !== code || cx.char(pos + 2) === code) return -1;
          const { left, right } = flanking(cx, pos, pos + 2);
          return cx.addDelimiter(type, pos, pos + 2, left, right);
        },
      },
    ],
  };
}

/** Bare addresses ("https://…", "www.…") read as links, as GFM and Obsidian do. */
const BareUrl: MarkdownConfig = {
  defineNodes: [{ name: "BareURL" }],
  parseInline: [
    {
      name: "BareURL",
      parse(cx, next, start) {
        if (next !== 104 && next !== 119) return -1; // h, w
        const before = cx.slice(start - 1, start);
        if (before && !/[\s*_~(]/.test(before)) return -1;
        const rest = cx.slice(start, Math.min(cx.end, start + 2048));
        const m = /^(?:https?:\/\/|www\.)[^\s<]*[^\s<?!.,:*_~)'"]/i.exec(rest);
        if (!m) return -1;
        return cx.addElement(cx.elt("BareURL", start, start + m[0].length));
      },
      before: "Link",
    },
  ],
};

export const sourceParser = baseParser.configure([Emphasis, pairedDelimiter("Strikethrough", 126), pairedDelimiter("Highlight", 61), BareUrl]);
const nodeName = (type: number) => sourceParser.nodeSet.types[type]?.name ?? "";
const childrenOf = (element: MdElement) => (element as unknown as { children?: MdElement[] }).children ?? [];

// ------------------------------------------------------------------------------------------------- a block's syntax

/** A formatted span: its range in the block's text and its markers (hidden while the selection does not touch it). */
export interface SourceUnit {
  from: number;
  to: number;
  /** its markers: hidden while the selection is away from the span, shown in the markers' colour while it touches it */
  marks: Array<[number, number]>;
  /** text hidden while the selection is away and shown as text otherwise (a link's address, a picture's description) */
  plain?: Array<[number, number]>;
  /** what it is (strong, em, strike, highlight, code, link, image, escape, break, html, underline, …, heading) */
  kind: string;
  /** a link's address */
  href?: string;
  /** a link's or a picture's text (between "[" and "]") */
  text?: [number, number];
  /** "line": the markers show while the selection is in the span's line (a heading's "#", an escape's "\") */
  reveal?: "line";
  /** drawn after the span (a picture), or in its place while it is hidden (a line break for <br>) */
  widget?: { kind: "image"; src: string; alt: string } | { kind: "break" };
}

/** The style of a stretch of the block's text. */
export interface SourceStyle {
  from: number;
  to: number;
  kind: "strong" | "em" | "strike" | "highlight" | "code" | "link" | "url" | "underline" | "superscript" | "subscript" | "kbd" | "color";
  href?: string;
  /** inline HTML's style="color: …" / "background-color: …" */
  color?: string;
  background?: string;
}

export interface HeadingSyntax {
  level: number;
  /** the end of "#… " (the markers and the spaces after them) */
  prefixEnd: number;
  /** a closing run of "#" at the line's end */
  close: [number, number] | null;
  /** Pandoc's attributes at the end (" {-}", " {.unnumbered}"): an unnumbered section */
  attributes: [number, number] | null;
  unnumbered: boolean;
}

export interface BlockSyntax {
  text: string;
  units: SourceUnit[];
  styles: SourceStyle[];
  heading: HeadingSyntax | null;
}

/** An inline object (a formula, an image) in a block's text: reads as a letter would. */
const OBJECT_CHAR = "\uE000";

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

/** "#" to "######" at a line's start, then spaces or the line's end: a heading (CommonMark, as Obsidian reads it). */
export function headingSyntax(text: string): HeadingSyntax | null {
  const m = /^(#{1,6})(?=[ \t]|$)[ \t]*/.exec(text);
  if (!m) return null;
  // Pandoc's "# Title {-}" (or {.unnumbered}, {#id .unnumbered}): an unnumbered section
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
    // "# #": the line holds the closing run only
    close = [m[0].length, m[0].length + rest.replace(/[ \t]+$/, "").length];
  }
  return { level: m[1].length, prefixEnd: m[0].length, close, attributes, unnumbered };
}

const STYLE_OF: Record<string, SourceStyle["kind"]> = { StrongEmphasis: "strong", Emphasis: "em", Strikethrough: "strike", Highlight: "highlight" };
const HTML_STYLES: Record<string, SourceStyle["kind"]> = {
  u: "underline", ins: "underline", sup: "superscript", sub: "subscript", kbd: "kbd", mark: "highlight", b: "strong",
  strong: "strong", i: "em", em: "em", s: "strike", del: "strike", strike: "strike",
};
/** HTML elements that have no closing tag. */
const VOID_TAGS = new Set(["area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr"]);
const SAFE_COLOR = /^(#[0-9a-f]{3,8}|[a-z]+|rgba?\([\d\s.,%]+\)|hsla?\([\d\s.,%a-z]+\))$/i;
/** Pictures drawn from their address as it stands (a web address or data); others are the document's own objects. */
const SHOWN_PICTURE = /^(https?:\/\/|data:image\/)/i;

/** An HTML tag's colours (style="color: …; background-color: …"). */
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

/** The formatted spans and styles of inline Markdown text (offsets into `text`). */
export function parseSourceInline(text: string, offset = 0): { units: SourceUnit[]; styles: SourceStyle[] } {
  const units: SourceUnit[] = [];
  const styles: SourceStyle[] = [];
  const slice = (from: number, to: number) => text.slice(from - offset, to - offset);
  const walk = (elements: MdElement[]) => {
    // inline HTML: a tag and its closing tag around text style it (<u>, <sup>, <kbd>, <span style="color: …">)
    const open: Array<{ tag: string; element: MdElement; attrs: string }> = [];
    for (const element of elements) {
      const name = nodeName(element.type);
      const children = childrenOf(element);
      const marks = children.filter((child) => nodeName(child.type).endsWith("Mark"));
      if (STYLE_OF[name] && marks.length >= 2) {
        const first = marks[0];
        const last = marks[marks.length - 1];
        units.push({ from: element.from, to: element.to, marks: [[first.from, first.to], [last.from, last.to]], kind: STYLE_OF[name] });
        if (last.from > first.to) styles.push({ from: first.to, to: last.from, kind: STYLE_OF[name] });
        walk(children.filter((child) => !nodeName(child.type).endsWith("Mark")));
      } else if (name === "InlineCode" && marks.length >= 2) {
        const first = marks[0];
        const last = marks[marks.length - 1];
        units.push({ from: element.from, to: element.to, marks: [[first.from, first.to], [last.from, last.to]], kind: "code" });
        if (last.from > first.to) styles.push({ from: first.to, to: last.from, kind: "code" });
      } else if (name === "Link" || name === "Image") {
        // [text](address "title"), [text][label]: the text shows; "](", ")" and the label are markers, the address
        // shows as a link while the selection is at it
        const linkMarks = children.filter((child) => nodeName(child.type) === "LinkMark");
        const url = children.find((child) => nodeName(child.type) === "URL");
        const label = children.find((child) => nodeName(child.type) === "LinkLabel");
        if (linkMarks.length < 2) continue;
        const textFrom = linkMarks[0].to;
        const textTo = linkMarks[1].from;
        const inline = linkMarks.length >= 4 && slice(linkMarks[2].from, linkMarks[2].to) === "(";
        // "[text]" alone (no address, no label): text
        if (!inline && !label) continue;
        const href = url ? slice(url.from, url.to) : undefined;
        const unitMarks: Array<[number, number]> = [[element.from, textFrom]];
        const plain: Array<[number, number]> = [];
        if (url) {
          unitMarks.push([textTo, url.from], [url.to, element.to]);
          plain.push([url.from, url.to]);
        } else unitMarks.push([textTo, element.to]);
        const picture = name === "Image" && href && SHOWN_PICTURE.test(href) ? href : null;
        // a picture drawn from its address: drawn alone while the selection is away from its source
        if (picture) plain.push([textFrom, textTo]);
        units.push({
          from: element.from,
          to: element.to,
          marks: unitMarks,
          plain,
          kind: name === "Image" ? "image" : "link",
          href,
          text: [textFrom, textTo],
          ...(picture ? { widget: { kind: "image" as const, src: picture, alt: slice(textFrom, textTo) } } : {}),
        });
        if (textTo > textFrom) styles.push({ from: textFrom, to: textTo, kind: "link", href });
        if (url && url.to > url.from) styles.push({ from: url.from, to: url.to, kind: "url", href });
        walk(children.filter((child) => child.from >= textFrom && child.to <= textTo && nodeName(child.type) !== "LinkMark"));
      } else if (name === "Autolink") {
        const url = children.find((child) => nodeName(child.type) === "URL");
        if (url) {
          units.push({ from: element.from, to: element.to, marks: [[element.from, url.from], [url.to, element.to]], kind: "link", href: slice(url.from, url.to) });
          styles.push({ from: url.from, to: url.to, kind: "link", href: slice(url.from, url.to) });
        }
      } else if (name === "BareURL" || name === "URL") {
        const href = slice(element.from, element.to);
        styles.push({ from: element.from, to: element.to, kind: "link", href: /^www\./i.test(href) ? `https://${href}` : href });
      } else if (name === "Escape") {
        // its "\" shows (as text) while the selection is in the line (Obsidian)
        units.push({ from: element.from, to: element.to, marks: [], plain: [[element.from, element.from + 1]], kind: "escape", reveal: "line" });
      } else if (name === "HTMLTag") {
        const tag = /^<(\/?)\s*([a-zA-Z][\w-]*)([^>]*)>$/s.exec(slice(element.from, element.to));
        if (!tag) continue;
        const tagName = tag[2].toLowerCase();
        const closing = tag[1] === "/";
        if (tagName === "br" && !closing) {
          units.push({ from: element.from, to: element.to, marks: [[element.from, element.to]], kind: "break", widget: { kind: "break" } });
          continue;
        }
        if (VOID_TAGS.has(tagName) || /\/\s*$/.test(tag[3])) continue;
        if (!closing) {
          open.push({ tag: tagName, element, attrs: tag[3] });
          continue;
        }
        const at = open.map((o) => o.tag).lastIndexOf(tagName);
        if (at < 0) continue;
        const start = open[at];
        open.splice(at);
        const kind = HTML_STYLES[tagName];
        units.push({ from: start.element.from, to: element.to, marks: [[start.element.from, start.element.to], [element.from, element.to]], kind: kind ?? "html" });
        if (element.from > start.element.to) {
          if (kind) styles.push({ from: start.element.to, to: element.from, kind });
          const colors = tagColors(start.attrs);
          if (colors.color || colors.background) styles.push({ from: start.element.to, to: element.from, kind: "color", ...colors });
        }
      }
    }
  };
  walk(sourceParser.parseInline(text, offset));
  return { units, styles };
}

const syntaxCache = new WeakMap<Node, BlockSyntax>();

/** A textblock's Markdown (cached: an unchanged block is the same node). */
export function blockSyntax(node: Node): BlockSyntax {
  const known = syntaxCache.get(node);
  if (known) return known;
  const text = blockText(node);
  const heading = headingCapable(node) ? headingSyntax(text) : null;
  const start = heading ? heading.prefixEnd : 0;
  const end = heading?.close ? heading.close[0] : heading?.attributes ? heading.attributes[0] : text.length;
  const inline = parseSourceInline(text.slice(0, end), 0);
  // a heading's markers ("# ", a closing run, " {-}") are a unit of their own: they show while the selection is in the line
  const units = inline.units.filter((unit) => unit.from >= start);
  if (heading) {
    const marks: Array<[number, number]> = [[0, heading.prefixEnd]];
    if (heading.close) marks.push(heading.close);
    if (heading.attributes) marks.push(heading.attributes);
    units.unshift({ from: 0, to: text.length, marks, kind: "heading", reveal: "line" });
  }
  const syntax: BlockSyntax = { text, units, styles: inline.styles.filter((style) => style.from >= start), heading };
  syntaxCache.set(node, syntax);
  return syntax;
}

/** Paragraphs and headings take a heading's syntax (a title, a caption or a code line does not). */
function headingCapable(node: Node): boolean {
  return node.type.name === "paragraph" || node.type.name === "heading";
}

// ------------------------------------------------------------------------------------------------- drawing it

const STYLE_CLASS: Record<SourceStyle["kind"], string> = {
  strong: "md-strong", em: "md-em", strike: "md-strike", highlight: "md-highlight", code: "md-code", link: "md-link",
  url: "md-url", underline: "md-underline", superscript: "md-sup", subscript: "md-sub", kbd: "md-kbd", color: "md-color",
};

function styleAttrs(style: SourceStyle): Record<string, string> {
  const attrs: Record<string, string> = { class: STYLE_CLASS[style.kind] };
  if (style.href) attrs["data-href"] = style.href;
  const css = [style.color ? `color:${style.color}` : "", style.background ? `background-color:${style.background}` : ""].filter(Boolean).join(";");
  if (css) attrs.style = css;
  return attrs;
}

/** Whether the selection touches [from, to] (both ends counted, as Obsidian shows a span's markers at its edges). */
function touches(state: EditorState, from: number, to: number): boolean {
  return state.selection.ranges.some((range) => range.$from.pos <= to && range.$to.pos >= from);
}

/** The line (between line breaks) of [from, to] in a block's text. */
function lineAround(text: string, from: number, to: number): [number, number] {
  const start = from === 0 ? 0 : text.lastIndexOf("\n", from - 1) + 1;
  const end = text.indexOf("\n", to);
  return [start, end < 0 ? text.length : end];
}

/** Whether a span's markers show: the selection at the span (its ends counted), or in its line for a heading. */
function unitShown(state: EditorState, text: string, start: number, unit: SourceUnit): boolean {
  if (unit.reveal !== "line") return touches(state, start + unit.from, start + unit.to);
  const [from, to] = lineAround(text, unit.from, unit.to);
  return touches(state, start + from, start + to);
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

function blockDecorations(node: Node, pos: number, state: EditorState | null): Decoration[] {
  const syntax = blockSyntax(node);
  const start = pos + 1;
  const decorations: Decoration[] = [];
  const blockEnd = start + node.content.size;
  const inBlock = state ? touches(state, start, blockEnd) : false;
  for (const style of syntax.styles) decorations.push(Decoration.inline(start + style.from, start + style.to, styleAttrs(style)));
  for (const unit of syntax.units) {
    const shown = !!state && inBlock && unitShown(state, syntax.text, start, unit);
    for (const [from, to] of unit.marks) {
      if (to > from) decorations.push(Decoration.inline(start + from, start + to, { class: shown ? "md-mark" : "md-mark md-hidden" }, { unit: unit.kind }));
    }
    if (!shown) for (const [from, to] of unit.plain ?? []) if (to > from) decorations.push(Decoration.inline(start + from, start + to, { class: "md-hidden" }, { unit: unit.kind }));
    if (unit.widget?.kind === "image") {
      decorations.push(Decoration.widget(start + unit.to, pictureWidget(unit.widget.src, unit.widget.alt), { side: -1, widget: "image", key: `image:${unit.widget.src}` }));
    } else if (unit.widget?.kind === "break" && !shown) {
      decorations.push(Decoration.widget(start + unit.to, lineBreakWidget, { side: -1, widget: "break", key: "break" }));
    }
  }
  if (state && inBlock) decorations.push(...typingDecorations(syntax, start, state));
  return decorations;
}

/** The quote the selection's head is in (the outermost): its "> " before each of its lines, as Obsidian shows them. */
function quoteMarks(state: EditorState): Decoration[] {
  const $head = state.selection.$head;
  let depth = 0;
  for (let d = 1; d <= $head.depth && !depth; d++) if ($head.node(d).type.name === "blockquote") depth = d;
  if (!depth) return [];
  const quotePos = $head.before(depth);
  const widgets: Decoration[] = [];
  $head.node(depth).descendants((node, offset) => {
    if (!node.isTextblock) return true;
    const pos = quotePos + 1 + offset;
    const $pos = state.doc.resolve(pos);
    let quotes = 0;
    for (let d = $pos.depth; d > 0; d--) if ($pos.node(d).type.name === "blockquote") quotes++;
    const text = "> ".repeat(quotes);
    widgets.push(Decoration.widget(pos + 1, quoteMarkWidget(text), { side: -1, widget: "quote", text, key: `quote:${quotes}` }));
    return false;
  });
  return widgets;
}

function quoteMarkWidget(text: string) {
  return () => {
    const mark = document.createElement("span");
    mark.className = "md-mark md-quote-mark";
    mark.contentEditable = "false";
    mark.textContent = text;
    return mark;
  };
}

/** Which characters of a block show nothing as the selection stands (markers of spans it is away from). */
function hiddenAt(node: Node, start: number, state: EditorState): Uint8Array {
  const syntax = blockSyntax(node);
  const hidden = new Uint8Array(syntax.text.length);
  for (const unit of syntax.units) {
    if (unitShown(state, syntax.text, start, unit)) continue;
    for (const [from, to] of unit.marks) hidden.fill(1, from, to);
    for (const [from, to] of unit.plain ?? []) hidden.fill(1, from, to);
  }
  return hidden;
}

/** CommonMark's flanking of a run of delimiter characters at [from, to) of a text (CJK-friendly, as the parser's). */
function textFlanking(text: string, from: number, to: number) {
  const before = text.slice(Math.max(0, from - 1), from);
  const after = text.slice(to, to + 1);
  const cjk = CJK.test(before) || CJK.test(after);
  const pBefore = !cjk && PUNCTUATION.test(before);
  const pAfter = !cjk && PUNCTUATION.test(after);
  const sBefore = /\s|^$/.test(before);
  const sAfter = /\s|^$/.test(after);
  return { left: !sAfter && (!pAfter || sBefore || pBefore), right: !sBefore && (!pBefore || sAfter || pAfter), pBefore };
}

/**
 * Feedback while formatting is typed (Obsidian's editor shows it): an empty pair around the caret ("**|**", "`|`",
 * "[|]") shows as markers, and an opening marker before the caret with no closing one yet ("a **bold|") shows as a
 * marker with the rest of its line in its formatting.
 */
function typingDecorations(syntax: BlockSyntax, start: number, state: EditorState): Decoration[] {
  if (!state.selection.empty) return [];
  const caret = state.selection.head - start;
  const text = syntax.text;
  if (caret < 0 || caret > text.length) return [];
  const [lineFrom, lineTo] = lineAround(text, caret, caret);
  const unitAt = (i: number) => syntax.units.find((unit) => unit.kind !== "heading" && i >= unit.from && i < unit.to);
  const marker = (from: number, to: number) => Decoration.inline(start + from, start + to, { class: "md-mark" }, { unit: "typing" });
  const before = text.charAt(caret - 1);
  const after = text.charAt(caret);
  if (before && after && !unitAt(caret - 1) && !unitAt(caret)) {
    if (before === after && "*_~=`".includes(before)) {
      let from = caret;
      while (from > lineFrom && text[from - 1] === before) from--;
      let to = caret;
      while (to < lineTo && text[to] === before) to++;
      const n = caret - from;
      // ("*|*" stays as it is: Obsidian)
      if (n === to - caret && (before === "`" ? n <= 3 : "*_".includes(before) ? n === 2 || n === 3 : n === 2)) return [marker(from, caret), marker(caret, to)];
    }
    if (before === "[" && after === "]") return [marker(caret - 1, caret), marker(caret, caret + 1)];
  }
  for (let i = caret - 1; i >= lineFrom; ) {
    const unit = unitAt(i);
    if (unit) {
      i = unit.from - 1;
      continue;
    }
    const ch = text[i];
    if (ch !== "*" && ch !== "_" && ch !== "~" && ch !== "=" && ch !== "`") {
      i--;
      continue;
    }
    let from = i;
    while (from > lineFrom && text[from - 1] === ch && !unitAt(from - 1)) from--;
    const run = i + 1 - from;
    const { left, right, pBefore } = textFlanking(text, from, i + 1);
    // (a backtick opens code whatever stands around it)
    const opens = ch === "`" ? i + 1 < lineTo : ch === "~" || ch === "=" ? run === 2 && left : left && (ch === "*" || !right || pBefore);
    if (opens) {
      const kinds: SourceStyle["kind"][] =
        ch === "`" ? ["code"] : ch === "~" ? ["strike"] : ch === "=" ? ["highlight"] : run >= 3 ? ["strong", "em"] : run === 2 ? ["strong"] : ["em"];
      return [
        marker(from, i + 1),
        Decoration.inline(start + i + 1, start + lineTo, { class: kinds.map((kind) => STYLE_CLASS[kind]).join(" ") }, { unit: "typing" }),
      ];
    }
    i = from - 1;
  }
  return [];
}

interface SourceLiveState {
  /** every block's decorations with its markers hidden */
  base: DecorationSet;
  /** every marker hidden whatever the selection (printing) until the selection or the text changes */
  away?: boolean;
}

/** Every marker hidden, the selection's too (the page view printed as it reads), until the selection moves. */
export function hideSourceMarkers(state: EditorState): Transaction {
  return state.tr.setMeta(sourceLiveKey, "away");
}

export const sourceLiveKey = new PluginKey<SourceLiveState>("visualtex-document-source");

function allDecorations(doc: Node): DecorationSet {
  const decorations: Decoration[] = [];
  doc.descendants((node, pos) => {
    if (node.isTextblock) {
      decorations.push(...blockDecorations(node, pos, null));
      return false;
    }
    return true;
  });
  return DecorationSet.create(doc, decorations);
}

/** The textblocks a transaction's steps touched (their positions in the new document). */
function changedBlocks(tr: Transaction): Array<{ node: Node; pos: number }> {
  const ranges: Array<[number, number]> = [];
  tr.mapping.maps.forEach((map, index) => {
    map.forEach((_oldStart, _oldEnd, newStart, newEnd) => {
      // where this step's change is in the final document
      const rest = tr.mapping.slice(index + 1);
      ranges.push([rest.map(newStart, -1), rest.map(newEnd, 1)]);
    });
  });
  const seen = new Set<number>();
  const blocks: Array<{ node: Node; pos: number }> = [];
  for (const [from, to] of ranges) {
    tr.doc.nodesBetween(Math.max(0, from), Math.min(tr.doc.content.size, to), (node, pos) => {
      if (node.isTextblock) {
        if (!seen.has(pos)) {
          seen.add(pos);
          blocks.push({ node, pos });
        }
        return false;
      }
      return true;
    });
  }
  return blocks;
}

/** A paragraph whose text begins "# " is a heading of that level, a heading without it a paragraph. */
function normalizeHeadings(state: EditorState, blocks: Array<{ node: Node; pos: number }>): Transaction | null {
  const { heading, paragraph } = state.schema.nodes;
  if (!heading || !paragraph) return null;
  let tr: Transaction | null = null;
  for (const { pos } of blocks) {
    const node = (tr ?? state.tr).doc.nodeAt(pos);
    if (!node || !headingCapable(node)) continue;
    const syntax = headingSyntax(blockText(node));
    const $pos = (tr ?? state.tr).doc.resolve(pos);
    const parent = $pos.parent;
    const index = $pos.index();
    if (syntax) {
      if (node.type === heading && node.attrs.level === syntax.level && !!node.attrs.unnumbered === syntax.unnumbered) continue;
      if (!parent.canReplaceWith(index, index + 1, heading)) continue;
      tr = (tr ?? state.tr).setNodeMarkup(pos, heading, { ...(node.type === heading ? node.attrs : {}), level: syntax.level, unnumbered: syntax.unnumbered });
    } else if (node.type === heading) {
      if (!parent.canReplaceWith(index, index + 1, paragraph)) continue;
      tr = (tr ?? state.tr).setNodeMarkup(pos, paragraph, null);
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
 * markers come into view as the caret reaches them). At a block's end, beside a formula or other object, in code and
 * in right-to-left text the editor's own keys and the browser do it.
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
  view.dispatch(state.tr.setSelection(extend ? TextSelection.create(state.doc, selection.anchor, pos) : TextSelection.create(state.doc, pos)).scrollIntoView());
  return true;
}

/**
 * Home / End on a block's first / last line: to the block's start / end, markers included (the browser's own key stops
 * at hidden markers, and reads the caret back a moment later). On the other lines of a wrapped block the browser does it.
 */
function stepLineEdge(view: EditorView, dir: -1 | 1, extend: boolean): boolean {
  const { state } = view;
  const selection = state.selection;
  if (!(selection instanceof TextSelection) || (!selection.empty && !extend)) return false;
  const $head = selection.$head;
  if (!$head.parent.isTextblock || $head.parent.type.spec.code) return false;
  if (!view.endOfTextblock(dir < 0 ? "up" : "down")) return false;
  const pos = dir < 0 ? $head.start() : $head.end();
  view.dispatch(state.tr.setSelection(extend ? TextSelection.create(state.doc, selection.anchor, pos) : TextSelection.create(state.doc, pos)).scrollIntoView());
  return true;
}

/** Up / Down: the column kept through a run of presses (the caret's x when the run began). */
const goals = new WeakMap<EditorView, { pos: number; x: number }>();

/**
 * Up / Down from the first or last line of a block into the block before or after it, when that block shows less than
 * its text (hidden markers; a heading's "#" with the number floated before it, where the browser's own step puts the
 * caret at the line's start or end): to the place under (over) the caret as the line is drawn before its markers show,
 * as Obsidian's editor does. Within a block, and next to other blocks, the browser does it.
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
  if (!hiddenAt(sibling, siblingPos + 1, state).some(Boolean)) return false;
  const dom = view.nodeDOM(siblingPos);
  if (!(dom instanceof HTMLElement)) return false;
  const range = document.createRange();
  range.selectNodeContents(dom);
  const rects = Array.from(range.getClientRects()).filter((rect) => rect.height > 0);
  if (!rects.length) return false;
  const line = dir < 0 ? rects[rects.length - 1] : rects[0];
  const goal = goals.get(view);
  const x = goal && goal.pos === $head.pos ? goal.x : view.coordsAtPos($head.pos).left;
  const found = view.posAtCoords({ left: x, top: (line.top + line.bottom) / 2 });
  if (!found || found.pos < siblingPos + 1 || found.pos > siblingPos + 1 + sibling.content.size) return false;
  view.dispatch(state.tr.setSelection(extend ? TextSelection.create(state.doc, selection.anchor, found.pos) : TextSelection.create(state.doc, found.pos)).scrollIntoView());
  goals.set(view, { pos: view.state.selection.head, x });
  return true;
}

/**
 * A click, End or Up / Down that put the caret where only hidden markers follow it in its line ("**bold|**" with the
 * markers hidden): the caret goes to the line's end, where it looks to be.
 */
function caretPastHidden(old: EditorState, state: EditorState): Transaction | null {
  const selection = state.selection;
  if (!(selection instanceof TextSelection) || !selection.empty || old.doc !== state.doc) return null;
  const $head = selection.$head;
  const node = $head.parent;
  if (!node.isTextblock || node.type.spec.code) return null;
  const offset = $head.parentOffset;
  const syntax = blockSyntax(node);
  const [, lineTo] = lineAround(syntax.text, offset, offset);
  if (lineTo <= offset) return null;
  const hidden = hiddenAt(node, $head.start(), old);
  for (let i = offset; i < lineTo; i++) if (!hidden[i]) return null;
  return state.tr.setSelection(TextSelection.create(state.doc, $head.start() + lineTo)).setMeta(sourceLiveKey, "caret");
}

/** Every paragraph and heading of a document (headings follow their "#" again after an input method's composition). */
function headingBlocks(doc: Node): Array<{ node: Node; pos: number }> {
  const blocks: Array<{ node: Node; pos: number }> = [];
  doc.descendants((node, pos) => {
    if (node.isTextblock) {
      if (headingCapable(node)) blocks.push({ node, pos });
      return false;
    }
    return true;
  });
  return blocks;
}

/** The source editor: decorations from each block's Markdown, headings from their "#". */
export function sourceLivePlugin(options: { enabled: () => boolean } = { enabled: () => true }): Plugin<SourceLiveState> {
  // An input method composing: a block never changes its type under it (the composition would end, its first letters
  // left in the text); the headings are made right when it ends.
  let composing = false;
  return new Plugin<SourceLiveState>({
    key: sourceLiveKey,
    state: {
      init: (_config, state) => ({ base: allDecorations(state.doc) }),
      apply(tr, value) {
        const away = tr.getMeta(sourceLiveKey) === "away" ? true : tr.docChanged || tr.selectionSet ? false : value.away;
        if (!tr.docChanged) return away === value.away ? value : { ...value, away };
        let base = value.base.map(tr.mapping, tr.doc);
        for (const { node, pos } of changedBlocks(tr)) {
          // the block's own decorations (inline ones: its range) made again from its text
          const own = base.find(pos + 1, pos + 1 + node.content.size);
          if (own.length) base = base.remove(own);
          base = base.add(tr.doc, blockDecorations(node, pos, null));
        }
        return { base, away };
      },
    },
    appendTransaction(transactions, old, state) {
      if (!options.enabled()) return null;
      if (transactions.some((tr) => tr.getMeta(sourceLiveKey) === "headings")) return normalizeHeadings(state, headingBlocks(state.doc));
      if (!transactions.some((tr) => tr.docChanged)) {
        // the caret put in by a click or the browser's keys (not by this plugin)
        if (!transactions.some((tr) => tr.selectionSet && !tr.getMeta(sourceLiveKey))) return null;
        return caretPastHidden(old, state);
      }
      if (composing) return null;
      const blocks: Array<{ node: Node; pos: number }> = [];
      transactions.forEach((tr, index) => {
        if (!tr.docChanged) return;
        // positions in the final state
        const later = transactions.slice(index + 1).filter((t) => t.docChanged);
        for (const block of changedBlocks(tr)) {
          let pos = block.pos;
          for (const t of later) pos = t.mapping.map(pos, 1);
          blocks.push({ node: block.node, pos });
        }
      });
      return normalizeHeadings(state, blocks);
    },
    props: {
      decorations(state) {
        if (!options.enabled()) return DecorationSet.empty;
        const value = sourceLiveKey.getState(state);
        const base = value?.base ?? DecorationSet.empty;
        if (value?.away) return base;
        // the blocks the selection is in: their markers as the selection shows them
        let decorations = base;
        const done = new Set<number>();
        for (const range of state.selection.ranges) {
          state.doc.nodesBetween(range.$from.pos, range.$to.pos, (node, pos) => {
            if (node.isTextblock) {
              if (!done.has(pos)) {
                done.add(pos);
                const own = decorations.find(pos + 1, pos + 1 + node.content.size);
                decorations = decorations.remove(own).add(state.doc, blockDecorations(node, pos, state));
              }
              return false;
            }
            return true;
          });
        }
        const quotes = quoteMarks(state);
        return quotes.length ? decorations.add(state.doc, quotes) : decorations;
      },
      handleKeyDown(view, event) {
        if (!options.enabled() || view.composing || event.isComposing) return false;
        if (event.altKey || event.metaKey || event.ctrlKey) return false;
        if (event.key === "ArrowLeft" || event.key === "ArrowRight") return stepCharacter(view, event.key === "ArrowLeft" ? -1 : 1, event.shiftKey);
        if (event.key === "ArrowUp" || event.key === "ArrowDown") return stepLine(view, event.key === "ArrowUp" ? -1 : 1, event.shiftKey);
        if (event.key === "Home" || event.key === "End") return stepLineEdge(view, event.key === "Home" ? -1 : 1, event.shiftKey);
        return false;
      },
      // what is pasted (from a web page, a word processor, the app's own copy) comes in as source
      transformPasted: (slice) => (options.enabled() ? toSourceSlice(slice) : slice),
      handleDOMEvents: {
        compositionstart: () => {
          composing = true;
          return false;
        },
        compositionend: (view) => {
          // (the committed text is read a moment after the event)
          setTimeout(() => {
            composing = false;
            if (!view.isDestroyed && options.enabled()) view.dispatch(view.state.tr.setMeta(sourceLiveKey, "headings").setMeta("addToHistory", false));
          }, 30);
          return false;
        },
      },
    },
  });
}

/** Whether a position is in a code span's text (between its backticks): what is typed there stays as typed. */
export function inSourceCode($pos: ResolvedPos): boolean {
  const node = $pos.parent;
  if (!node.isTextblock || node.type.spec.code) return false;
  const offset = $pos.parentOffset;
  return blockSyntax(node).units.some((unit) => unit.kind === "code" && offset >= unit.marks[0][1] && offset <= unit.marks[unit.marks.length - 1][0]);
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
 * A document whose formatting is marks (a document saved before the source editor, Markdown read by markdown.ts, a
 * paste) as source: each textblock's formatting written as markers in its text, each heading given its "#" prefix.
 * Other marks (underline, colour, size, …) stay marks.
 */
export function toSourceDocument(doc: Node): Node {
  const schema = doc.type.schema;
  const convert = (node: Node): Node => {
    if (node.isTextblock) return sourceBlock(node, schema);
    if (node.isLeaf) return node;
    const children: Node[] = [];
    node.forEach((child) => children.push(convert(child)));
    return node.copy(Fragment.fromArray(children));
  };
  return convert(doc);
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

/** Pasted content as source: its textblocks' formatting written as markers (toSourceDocument for a slice). */
export function toSourceSlice(slice: Slice): Slice {
  const first = slice.content.firstChild;
  if (!first) return slice;
  const schema = first.type.schema;
  const convert = (fragment: Fragment): Fragment => {
    const out: Node[] = [];
    let inline: Node[] = [];
    const flushInline = () => {
      // inline content at the top of a slice: as a paragraph's
      if (!inline.length) return;
      sourceBlock(schema.nodes.paragraph.create(null, inline), schema).forEach((child) => out.push(child));
      inline = [];
    };
    fragment.forEach((node) => {
      if (node.isInline) {
        inline.push(node);
        return;
      }
      flushInline();
      if (node.isTextblock) out.push(sourceBlock(node, schema));
      else if (node.isLeaf) out.push(node);
      else out.push(node.copy(convert(node.content)));
    });
    flushInline();
    return Fragment.from(out);
  };
  return new Slice(convert(slice.content), slice.openStart, slice.openEnd);
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
