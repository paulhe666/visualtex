// Markdown input (DESIGN.md §7): in Markdown input mode, Markdown typed in the text formats it as in Typora. Inline
// Markdown (**bold**, *italic*, `code`, ~~strike~~, ==highlight==, [text](url), <u>…</u>) is markdownLive.ts's: it
// formats as soon as it is closed and shows its markers while the caret is in it. At the start of a line "- ", "* ",
// "+ ", "1. " make a list, "> " a quote, "[ ] " in a list item a task — at once, on the space. A line of "# Title",
// "```lang", "---" / "***" / "___" becomes a section, a code block or a rule when Enter ends it or the caret leaves it;
// "| a | b |" makes a table and "[toc]" a table of contents on Enter, "[!NOTE]" in a quote makes it an alert. "\"
// before a punctuation character types that character (Markdown's escape), which then never formats anything. Brackets
// and quotes come in pairs. Typora's keys: Ctrl+1…6 and Ctrl+0 (heading levels, body text), Ctrl+= / Ctrl+- (a level
// up / down), Ctrl+B/I/U, Ctrl+Shift+` (code), Alt+Shift+5 (strike), Ctrl+K (link), Ctrl+\ (clear format), Ctrl+T
// (table), Ctrl+Shift+K (code block), Ctrl+Shift+M (formula), Ctrl+Shift+Q (quote), Ctrl+Shift+[ / ] (lists), Ctrl+[ / ]
// (list level), Ctrl+L / E / D (select the line, the styled span, the word).
import { Fragment, type Node, type ResolvedPos } from "prosemirror-model";
import { Plugin, PluginKey, Selection, TextSelection, type Command, type EditorState, type Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { closeHistory, isHistoryTransaction } from "prosemirror-history";
import { ReplaceAroundStep, ReplaceStep, canJoin, findWrapping } from "prosemirror-transform";
import { joinBackward, lift, setBlockType, wrapIn } from "prosemirror-commands";
import { keymap } from "prosemirror-keymap";
import { liftListItem, sinkListItem } from "prosemirror-schema-list";
import { insertTable, toggleList } from "./blocks.ts";
import { insertDisplayLatex } from "./commands.ts";
import { insertCodeBlock } from "./codeBlocks.ts";
import { requestImage, resolveImageSource } from "./images.ts";
import { insertFootnote, setFootnoteText } from "./notes.ts";
import { newNodeId } from "./ids.ts";
import { blockLanguage } from "./markdown.ts";
import { ALERTS } from "./schema.ts";
import { clearFormat, isEscaped, markEscaped, markdownLiveKey, shownMarkdown, typoraFormat, typoraLink } from "./markdownLive.ts";

export type InputMode = "latex" | "markdown";

/** Groups of Markdown rules that can be turned off. */
export interface MarkdownRuleSwitches {
  /** line-start markers and the Enter blocks */
  blocks: boolean;
  /** inline marks */
  marks: boolean;
  /** a Chinese input method's full-width variants (》 1。 ··· 、 ／, the full-width space) */
  chinese: boolean;
}

export const ALL_RULES: MarkdownRuleSwitches = { blocks: true, marks: true, chinese: true };

interface MarkdownInputState {
  /** the last change was a conversion: the next one starts an undo step of its own */
  converted: boolean;
  /** the line typed in last (its start): leaving it reads its Markdown ("# Title" becomes a section) */
  edited: number | null;
  /** closing brackets and quotes put in with their opening ones: typing that character steps over them */
  pairs: number[];
}

type InputMeta = { converted: true; left?: true } | { pairs: number[] };

export const markdownInputKey = new PluginKey<MarkdownInputState>("visualtex-document-markdown-input");

/** CommonMark's escapable characters: ASCII punctuation. */
const ESCAPABLE = "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~";
/** The space that ends a line-start marker: a space, a no-break space or an input method's full-width space. */
const SPACE = "[ \\u00a0\\u3000]";
const OBJECT = "￼";

/** How far back the line-start markers are looked for. */
const MAX_MATCH = 500;

/** The textblock's text before `$pos` (formulas and other inline objects one character each, as in positions). */
function textBefore($pos: ResolvedPos): string {
  return $pos.parent.textBetween(Math.max(0, $pos.parentOffset - MAX_MATCH), $pos.parentOffset, undefined, OBJECT);
}

const lineText = (node: Node) => node.textBetween(0, node.content.size, undefined, OBJECT);

// ------------------------------------------------------- line-start markers --

/** The line-start markers; with `chinese` also what a Chinese input method types for them ("1。", "》", the
 * full-width space). */
function lineStartRules(chinese: boolean) {
  const space = chinese ? SPACE : "[ \\u00a0]";
  return {
    bullet: new RegExp(`^[-*+]${space}$`),
    ordered: new RegExp(`^(\\d{1,9})${chinese ? "[.)。．）]" : "[.)]"}${space}$`),
    quote: new RegExp(`^(${chinese ? "[>》]+" : ">+"})${space}$`),
  };
}
const LINE_START = { chinese: lineStartRules(true), plain: lineStartRules(false) };

/** The block wrapped in `type` (a list or a quote), joined to the same kind of block right before it. */
function wrapBlock(state: EditorState, $cursor: ResolvedPos, type: Node["type"], attrs: Record<string, unknown> | null): Transaction | null {
  const start = $cursor.start();
  const tr = state.tr.delete(start, $cursor.pos);
  const range = tr.doc.resolve(start).blockRange();
  const wrapping = range && findWrapping(range, type, attrs);
  if (!range || !wrapping) return null;
  tr.wrap(range, wrapping);
  const before = tr.doc.resolve(start - 1).nodeBefore;
  const joins =
    before?.type === type &&
    canJoin(tr.doc, start - 1) &&
    (type.name !== "ordered_list" || before.attrs.start + before.childCount === (attrs?.start ?? 1));
  if (joins) tr.join(start - 1);
  return tr;
}

/** A paragraph after the block at `inside`, when nothing follows it (the text goes on after a code block). */
function paragraphAfter(tr: Transaction, inside: number) {
  const after = tr.doc.resolve(inside).after();
  if (after >= tr.doc.content.size || !tr.doc.resolve(after).nodeAfter) tr.insert(after, tr.doc.type.schema.nodes.paragraph.create());
}

const ALERT = /^\[!(\w+)\]$/;
const TASK_PLAIN = /^\[([ xX])\][  ]$/;
const TASK_CHINESE = /^[\[【]([ xX])[\]】][  　]$/;

function lineStartConversion(state: EditorState, $cursor: ResolvedPos, before: string, rules: MarkdownRuleSwitches): Transaction | null {
  if ($cursor.parentOffset > 16 || !rules.blocks) return null;
  const block = $cursor.parent;
  const { nodes } = state.schema;
  const start = $cursor.start();
  if (isEscaped(state, start, $cursor.pos)) return null;
  if (block.type !== nodes.paragraph) return null;
  const { bullet: BULLET, ordered: ORDERED, quote: QUOTE } = rules.chinese ? LINE_START.chinese : LINE_START.plain;
  // "[!NOTE]" as a quote's first line: the quote is an alert of that kind (the line is its title)
  const alert = ALERT.exec(before);
  const inQuote = $cursor.depth >= 2 && $cursor.node(-1).type === nodes.blockquote && $cursor.index(-1) === 0;
  if (alert && inQuote && $cursor.parentOffset === block.content.size && ALERTS.includes(alert[1].toLowerCase()) && !$cursor.node(-1).attrs.alert) {
    return state.tr.delete(start, $cursor.pos).setNodeAttribute($cursor.before(-1), "alert", alert[1].toLowerCase());
  }
  // "- " or "1. " at the start of a list item after the first: the item one level deeper, in a list of that kind
  const inItem = $cursor.depth >= 3 && $cursor.node(-1).type === nodes.list_item && $cursor.index(-1) === 0;
  const ordered = ORDERED.exec(before);
  if (inItem && (BULLET.test(before) || ordered)) return nestItem(state, $cursor, ordered ? nodes.ordered_list : nodes.bullet_list, ordered ? Number(ordered[1]) : 1);
  // "[ ] " / "[x] " opening a list item: a task (a Chinese input method types 【 】)
  const task = (rules.chinese ? TASK_CHINESE : TASK_PLAIN).exec(before);
  if (task && $cursor.depth >= 2 && $cursor.node(-1).type === nodes.list_item && $cursor.index(-1) === 0 && $cursor.node(-1).attrs.checked === null) {
    return state.tr.delete(start, $cursor.pos).setNodeAttribute($cursor.before(-1), "checked", task[1] !== " ");
  }
  let match: RegExpExecArray | null;
  if (BULLET.test(before)) return wrapBlock(state, $cursor, nodes.bullet_list, null);
  if ((match = ORDERED.exec(before))) return wrapBlock(state, $cursor, nodes.ordered_list, { start: Number(match[1]) });
  if ((match = QUOTE.exec(before))) {
    // ">> " nests: one quote for each ">"
    const tr = wrapBlock(state, $cursor, nodes.blockquote, null);
    for (let level = 1; tr && level < match[1].length; level++) {
      const range = tr.doc.resolve(tr.mapping.map(start)).blockRange();
      const wrapping = range && findWrapping(range, nodes.blockquote);
      if (range && wrapping) tr.wrap(range, wrapping);
    }
    return tr;
  }
  return null;
}

/** The list item at the caret one level deeper, in a list of `type` (its marker deleted). */
function nestItem(state: EditorState, $cursor: ResolvedPos, type: Node["type"], start: number): Transaction | null {
  const tr = state.tr.delete($cursor.start(), $cursor.pos);
  let sunk: Transaction | null = null;
  sinkListItem(state.schema.nodes.list_item)(state.apply(tr), (t) => {
    sunk = t;
  });
  if (!sunk) return null;
  for (const step of (sunk as Transaction).steps) tr.step(step);
  const $in = tr.doc.resolve(tr.mapping.map($cursor.start()));
  for (let d = $in.depth; d > 0; d--) {
    const list = $in.node(d);
    if (list.type !== state.schema.nodes.bullet_list && list.type !== state.schema.nodes.ordered_list) continue;
    if (list.type !== type) tr.setNodeMarkup($in.before(d), type, { ...list.attrs, start });
    break;
  }
  return tr;
}

/** The line-start conversion the text before the caret calls for, if any. */
function conversionAtCaret(state: EditorState, rules: MarkdownRuleSwitches): Transaction | null {
  const { $cursor } = state.selection as TextSelection;
  if (!$cursor || !$cursor.parent.isTextblock || $cursor.parent.type.spec.code) return null;
  return lineStartConversion(state, $cursor, textBefore($cursor), rules);
}

/** Whether a line-start conversion could follow the text `before` (a quick test before the typed text goes in). */
function mayConvert(before: string): boolean {
  return before.length <= 17 && (Object.values(LINE_START.chinese).some((rule) => rule.test(before)) || TASK_CHINESE.test(before) || ALERT.test(before));
}

function asConversion(tr: Transaction): Transaction {
  return closeHistory(tr).setMeta(markdownInputKey, { converted: true } satisfies InputMeta).scrollIntoView();
}

function convert(view: EditorView, rules: MarkdownRuleSwitches): boolean {
  const tr = conversionAtCaret(view.state, rules);
  if (!tr) return false;
  view.dispatch(asConversion(tr));
  return true;
}

// ------------------------------------------------------ lines read on Enter --

const HEADING = /^(#{1,6})[ \t ](.*)$/;
const HEADING_CHINESE = /^(#{1,6})[ \t 　](.*)$/;
const ENTER_FENCE = /^(?:```|~~~|···)[ \t]*([^\s`]*)[ \t]*$/;
const ENTER_FENCE_PLAIN = /^(?:```|~~~)[ \t]*([^\s`]*)[ \t]*$/;
const ENTER_RULE = /^(?:(?:-[ \t]*){3,}|(?:\*[ \t]*){3,}|(?:_[ \t]*){3,})$/;
const ENTER_TABLE = /^\|(.*)\|$/;
const ENTER_TOC = /^\[toc\]$/i;
/** a line of HTML that starts one of CommonMark's HTML blocks (a <div>, <details>, <table>…) or is a comment */
const ENTER_HTML =
  /^(?:<(address|article|aside|blockquote|center|details|dialog|dir|div|dl|fieldset|figcaption|figure|footer|form|h[1-6]|header|hr|iframe|main|menu|nav|ol|p|section|summary|table|ul|video|audio|picture)(?:\s[^>]*)?\/?>.*|<!--.*-->)$/i;
/** "[^1]: text": a footnote's text */
const ENTER_FOOTNOTE = /^\[\^([^\]\s]+)\]:\s*(.+)$/;

/** "| a | b |" -> ["a", "b"] ("\|" is a character of its cell). */
function tableCells(inner: string): string[] {
  return inner.split(/(?<!\\)\|/).map((cell) => cell.trim().replace(/\\\|/g, "|"));
}

/** "# Title" (with closing #s or not; Pandoc's "{-}" after it: unnumbered): the heading's level and where its text
 *  starts and ends in the line. */
function headingOf(line: string, chinese: boolean): { level: number; start: number; end: number; unnumbered: boolean } | null {
  const match = (chinese ? HEADING_CHINESE : HEADING).exec(line);
  if (!match) return null;
  const start = match[1].length + 1;
  let end = line.length;
  const attributes = /[ \t]+\{([^{}]*)\}[ \t]*$/.exec(line.slice(start));
  const unnumbered = !!attributes && attributes[1].trim().split(/\s+/).some((a) => a === "-" || a === ".unnumbered");
  if (unnumbered) end = start + attributes!.index;
  const closing = /[ \t]+#+[ \t]*$/.exec(line.slice(start, end));
  if (closing) end = start + closing.index;
  else if (/^#+[ \t]*$/.test(line.slice(start, end))) end = start;
  return { level: match[1].length, start, end, unnumbered };
}

/** The line at `pos` (a paragraph) made a section: its "# " and closing "#"s go, its text and formatting stay. */
function headingConversion(tr: Transaction, pos: number, node: Node, heading: { level: number; start: number; end: number; unnumbered: boolean }): boolean {
  const $pos = tr.doc.resolve(pos);
  const type = tr.doc.type.schema.nodes.heading;
  if (!$pos.parent.canReplaceWith($pos.index(), $pos.index() + 1, type)) return false;
  const start = pos + 1;
  if (heading.end < node.content.size) tr.delete(start + heading.end, start + node.content.size);
  tr.delete(start, start + heading.start);
  tr.setNodeMarkup(pos, type, { id: node.attrs.id, level: heading.level, unnumbered: heading.unnumbered, align: node.attrs.align ?? null, indent: 0 });
  return true;
}

/** "```lang": the line at `pos` a code block (empty, in that language). */
function fenceConversion(tr: Transaction, pos: number, node: Node, language: string): boolean {
  const $pos = tr.doc.resolve(pos);
  const { code_block } = tr.doc.type.schema.nodes;
  if (!$pos.parent.canReplaceWith($pos.index(), $pos.index() + 1, code_block)) return false;
  tr.delete(pos + 1, pos + 1 + node.content.size);
  tr.setNodeMarkup(pos, code_block, { id: newNodeId(), language: blockLanguage(language) });
  return true;
}

/** "---": the line at `pos` a rule. */
function ruleConversion(tr: Transaction, pos: number, node: Node): boolean {
  const $pos = tr.doc.resolve(pos);
  const { horizontal_rule } = tr.doc.type.schema.nodes;
  if (!$pos.parent.canReplaceWith($pos.index(), $pos.index() + 1, horizontal_rule)) return false;
  tr.replaceWith(pos, pos + node.nodeSize, horizontal_rule.create());
  return true;
}

function enterConversion(state: EditorState, $cursor: ResolvedPos, line: string, rules: MarkdownRuleSwitches): Transaction | null {
  const { nodes } = state.schema;
  const parent = $cursor.node(-1);
  const index = $cursor.index(-1);
  const before = $cursor.before();
  const after = $cursor.after();
  const atEnd = $cursor.parentOffset === $cursor.parent.content.size;
  // "# Title": a section; Enter goes on with body text (the line is split where the caret is)
  const heading = $cursor.parentOffset > 0 ? headingOf(line, rules.chinese) : null;
  if (heading) {
    const tr = state.tr;
    if (!headingConversion(tr, before, $cursor.parent, heading)) return null;
    const caret = Math.min(Math.max(tr.mapping.map($cursor.pos, -1), before + 1), tr.doc.resolve(before + 1).end());
    tr.split(caret, 1, [{ type: nodes.paragraph, attrs: { align: $cursor.parent.attrs.align ?? null } }]);
    return tr.setSelection(TextSelection.create(tr.doc, caret + 2));
  }
  if (!atEnd) return null;
  let match: RegExpExecArray | null;
  if ((match = (rules.chinese ? ENTER_FENCE : ENTER_FENCE_PLAIN).exec(line))) {
    const tr = state.tr;
    if (!fenceConversion(tr, before, $cursor.parent, match[1])) return null;
    paragraphAfter(tr, before + 1);
    return tr.setSelection(TextSelection.create(tr.doc, before + 1));
  }
  if (ENTER_RULE.test(line)) {
    const content = Fragment.from([nodes.horizontal_rule.create(), nodes.paragraph.create()]);
    if (!parent.canReplace(index, index + 1, content)) return null;
    const tr = state.tr.replaceWith(before, after, content);
    return tr.setSelection(TextSelection.create(tr.doc, before + 2));
  }
  if ((match = ENTER_TABLE.exec(line))) {
    const cells = tableCells(match[1]);
    if (cells.length < 2) return null;
    const cell = (text: string) => nodes.table_cell.create(null, nodes.paragraph.create(null, text ? state.schema.text(text) : null));
    const header = nodes.table_row.create(null, cells.map(cell));
    const body = nodes.table_row.create(null, cells.map(() => cell("")));
    const table = nodes.table.create({ id: newNodeId() }, [header, body]);
    const content = Fragment.from([table, nodes.paragraph.create()]);
    if (!parent.canReplace(index, index + 1, content)) return null;
    const tr = state.tr.replaceWith(before, after, content);
    // table > row > cell > paragraph: the caret in the first cell of the second row
    return tr.setSelection(TextSelection.create(tr.doc, before + 1 + header.nodeSize + 3));
  }
  if (ENTER_TOC.test(line) && nodes.toc) {
    const content = Fragment.from([nodes.toc.create({ id: newNodeId() }), nodes.paragraph.create()]);
    if (!parent.canReplace(index, index + 1, content)) return null;
    const tr = state.tr.replaceWith(before, after, content);
    return tr.setSelection(TextSelection.create(tr.doc, before + 2));
  }
  if ((match = ENTER_HTML.exec(line)) && nodes.html_block) {
    const content = Fragment.from([nodes.html_block.create({ id: newNodeId(), html: line }), nodes.paragraph.create()]);
    if (!parent.canReplace(index, index + 1, content)) return null;
    const tr = state.tr.replaceWith(before, after, content);
    return tr.setSelection(TextSelection.create(tr.doc, before + 2));
  }
  // "[^1]: text" gives its text to an empty footnote (the first one, or the one of that number)
  if ((match = ENTER_FOOTNOTE.exec(line)) && nodes.footnote) {
    const empty: number[] = [];
    state.doc.descendants((node, pos) => {
      if (node.type === nodes.footnote && !String(node.attrs.text ?? "").trim()) empty.push(pos);
      return true;
    });
    const all: number[] = [];
    state.doc.descendants((node, pos) => {
      if (node.type === nodes.footnote) all.push(pos);
      return true;
    });
    const numbered = /^\d+$/.test(match[1]) ? all[Number(match[1]) - 1] : undefined;
    const target = numbered !== undefined && empty.includes(numbered) ? numbered : empty[0];
    if (target === undefined) return null;
    const tr = state.tr;
    setFootnoteText(target, match[2].trim())(state, (t) => {
      for (const step of t.steps) tr.step(step);
    });
    tr.delete(tr.mapping.map(before), tr.mapping.map(after));
    if (!tr.doc.content.size || !tr.doc.lastChild?.isTextblock) tr.insert(tr.doc.content.size, nodes.paragraph.create());
    const at = Math.min(tr.mapping.map(before), tr.doc.content.size - 1);
    return tr.setSelection(Selection.near(tr.doc.resolve(at)));
  }
  return null;
}

/** Enter on a line of Markdown that makes a block ("# Title", "```lang", "---", "| a | b |", "[toc]"): that block. */
export function markdownEnter(mode: () => InputMode, rules: () => MarkdownRuleSwitches = () => ALL_RULES): Command {
  return (state, dispatch) => {
    if (mode() !== "markdown" || !rules().blocks) return false;
    const { $cursor } = state.selection as TextSelection;
    if (!$cursor || $cursor.parent.type !== state.schema.nodes.paragraph) return false;
    if (isEscaped(state, $cursor.start(), $cursor.end())) return false;
    const line = lineText($cursor.parent);
    const tr = enterConversion(state, $cursor, line, rules());
    if (!tr) return false;
    if (dispatch) dispatch(asConversion(tr));
    return true;
  };
}

/** Enter in a section (Typora): the text after the caret is body text; at the section's start, an empty line goes
 *  before it. */
export function markdownHeadingEnter(mode: () => InputMode): Command {
  return (state, dispatch) => {
    if (mode() !== "markdown") return false;
    const { $cursor } = state.selection as TextSelection;
    const { heading, paragraph } = state.schema.nodes;
    if (!$cursor || $cursor.parent.type !== heading) return false;
    if (!dispatch) return true;
    const tr = state.tr;
    if ($cursor.parentOffset === 0 && $cursor.parent.content.size > 0) {
      dispatch(tr.insert($cursor.before(), paragraph.create()).scrollIntoView());
      return true;
    }
    tr.split($cursor.pos, 1, [{ type: paragraph, attrs: { align: $cursor.parent.attrs.align ?? null } }]);
    dispatch(tr.setSelection(TextSelection.create(tr.doc, $cursor.pos + 2)).scrollIntoView());
    return true;
  };
}

/** The Markdown of a line typed and left (Typora): "# Title" a section, "```lang" a code block, "---" a rule. */
function leftLineConversion(state: EditorState, pos: number, rules: MarkdownRuleSwitches): Transaction | null {
  const node = state.doc.nodeAt(pos);
  if (!node || node.type !== state.schema.nodes.paragraph || isEscaped(state, pos + 1, pos + node.nodeSize - 1)) return null;
  const line = lineText(node);
  const tr = state.tr;
  const heading = headingOf(line, rules.chinese);
  if (heading && heading.end > heading.start && headingConversion(tr, pos, node, heading)) return tr;
  const fence = (rules.chinese ? ENTER_FENCE : ENTER_FENCE_PLAIN).exec(line);
  if (fence && fenceConversion(tr, pos, node, fence[1])) return tr;
  if (ENTER_RULE.test(line) && ruleConversion(tr, pos, node)) return tr;
  return null;
}

// --------------------------------------------------------------- the plugin --

/** Brackets and quotes typed in pairs (Typora's "Auto pair brackets and quotes"). */
const PAIRS: Record<string, string> = { "(": ")", "[": "]", "{": "}", '"': '"', "'": "'" };

/** The typed character `text` at `from..to`: a pair, a step over a closing character, a wrapped selection. */
function pairInput(view: EditorView, from: number, to: number, text: string): boolean {
  const state = view.state;
  const $from = state.doc.resolve(from);
  const pairs = markdownInputKey.getState(state)?.pairs ?? [];
  const next = $from.parent.textBetween($from.parentOffset, Math.min($from.parent.content.size, $from.parentOffset + 1), undefined, OBJECT);
  const prev = $from.parent.textBetween(Math.max(0, $from.parentOffset - 1), $from.parentOffset, undefined, OBJECT);
  // over the character put in with its opening one (a bracket, a quote, a tag's name)
  if (from === to && next === text && pairs.includes(from)) {
    view.dispatch(state.tr.setSelection(TextSelection.create(state.doc, from + 1)).scrollIntoView());
    return true;
  }
  const close = PAIRS[text];
  if (!close) return false;
  const quote = text === close;
  if (from !== to) {
    // a selection in one line: wrapped
    if (!state.doc.resolve(to).sameParent($from)) return false;
    const tr = state.tr.insertText(close, to).insertText(text, from);
    tr.setSelection(TextSelection.create(tr.doc, to + 1));
    view.dispatch(tr.scrollIntoView());
    return true;
  }
  // not before a word; a quote not after a word either (it's an apostrophe: don't, it's)
  if (/[\p{L}\p{N}_]/u.test(next) || (quote && /[\p{L}\p{N}_\\]/u.test(prev))) return false;
  if (prev === "\\") return false;
  const tr = state.tr.insertText(text + close, from, to);
  tr.setSelection(TextSelection.create(tr.doc, from + 1));
  view.dispatch(tr.setMeta(markdownInputKey, { pairs: [from + 1] } satisfies InputMeta).scrollIntoView());
  return true;
}

/** "</" in a line with an unclosed <u>, <sup>, <sub>, <kbd>…: the tag's name and ">" (Typora). */
function closeTag(view: EditorView, from: number, to: number, text: string): boolean {
  if (text !== "/" || from !== to) return false;
  const $from = view.state.doc.resolve(from);
  const before = textBefore($from);
  if (!before.endsWith("<")) return false;
  const open: string[] = [];
  for (const match of before.slice(0, -1).matchAll(/<(\/?)(u|ins|sup|sub|kbd|mark|s|del|strike|b|strong|i|em)>/gi)) {
    const name = match[2].toLowerCase();
    if (!match[1]) open.push(name);
    else if (open[open.length - 1] === name) open.pop();
  }
  const name = open.pop();
  if (!name) return false;
  const tr = view.state.tr.insertText(`/${name}>`, from, to);
  tr.setSelection(TextSelection.create(tr.doc, from + 1));
  // typed on, the name and ">" are stepped over
  const pairs = Array.from({ length: name.length + 1 }, (_, i) => from + 1 + i);
  view.dispatch(tr.setMeta(markdownInputKey, { pairs } satisfies InputMeta).scrollIntoView());
  return true;
}

/** "![alt](picture)" closed: the picture, when it can be found (a file next to the document, an image already in it). */
function typedImage(view: EditorView, from: number, to: number, text: string): boolean {
  if (text !== ")" || from !== to || !view.state.schema.nodes.image) return false;
  const $from = view.state.doc.resolve(from);
  const match = /!\[([^\]\n]*)\]\(([^()\s]+)$/.exec(textBefore($from));
  if (!match) return false;
  const key = resolveImageSource(match[2]);
  if (!key) return false;
  const start = from - match[0].length;
  view.dispatch(view.state.tr.replaceWith(start, from + pairedAfter(view.state, from, ")"), view.state.schema.nodes.image.create({ src: key, alt: match[1] })).scrollIntoView());
  return true;
}

/** 1 when the character after `pos` is `ch` put in with its opening one (it goes with what closes it). */
function pairedAfter(state: EditorState, pos: number, ch: string): number {
  const $pos = state.doc.resolve(pos);
  const next = $pos.parent.textBetween($pos.parentOffset, Math.min($pos.parent.content.size, $pos.parentOffset + 1));
  return next === ch && (markdownInputKey.getState(state)?.pairs ?? []).includes(pos) ? 1 : 0;
}

/** "[^1]" closed: a footnote there, its text asked for at once (Typora's footnote reference). */
function typedFootnote(view: EditorView, from: number, to: number, text: string): boolean {
  if (text !== "]" || from !== to || !view.state.schema.nodes.footnote) return false;
  const $from = view.state.doc.resolve(from);
  const match = /\[\^([^\]\s]+)$/.exec(textBefore($from));
  if (!match || isEscaped(view.state, from - match[0].length, from)) return false;
  const start = from - match[0].length;
  const tr = view.state.tr.delete(start, from + pairedAfter(view.state, from, "]"));
  view.dispatch(tr.setSelection(TextSelection.create(tr.doc, start)));
  return insertFootnote()(view.state, view.dispatch, view);
}

export function markdownInputPlugin(mode: () => InputMode, rules: () => MarkdownRuleSwitches = () => ALL_RULES): Plugin<MarkdownInputState> {
  return new Plugin<MarkdownInputState>({
    key: markdownInputKey,
    state: {
      init: () => ({ converted: false, edited: null, pairs: [] }),
      apply(tr, value, _old, state) {
        const meta = tr.getMeta(markdownInputKey) as InputMeta | undefined;
        let { edited, pairs } = value;
        if (tr.docChanged) {
          if (edited !== null) {
            const mapped = tr.mapping.mapResult(edited, 1);
            edited = mapped.deleted ? null : mapped.pos;
          }
          // a closing character stays one while it is there and the caret is in its line
          pairs = pairs.map((pos) => ({ from: tr.mapping.map(pos, 1), to: tr.mapping.map(pos + 1, -1) })).filter((r) => r.to - r.from === 1).map((r) => r.from);
          // the line typed in (not by undo, not by a conversion or another plugin's follow-up)
          const { $head } = state.selection;
          const typing = !meta && !isHistoryTransaction(tr) && !tr.getMeta("appendedTransaction") && !tr.getMeta(markdownLiveKey);
          if (typing && $head.parent.type === state.schema.nodes.paragraph) edited = $head.before();
        }
        if (meta && "left" in meta) edited = null;
        if (meta && "pairs" in meta) pairs = [...pairs, ...meta.pairs].slice(-24);
        const $head = state.selection.$head;
        pairs = pairs.filter((pos) => pos >= $head.start() && pos <= $head.end());
        const converted = meta && "converted" in meta ? true : tr.docChanged || tr.selectionSet ? false : value.converted;
        return converted === value.converted && edited === value.edited && pairs === value.pairs ? value : { converted, edited, pairs };
      },
    },
    // The change after a conversion starts an undo step of its own: Ctrl+Z then takes back just that change. (Not
    // another plugin's follow-up: ids given, the Markdown shown around the caret — an appended transaction is not
    // marked as one yet when it is filtered.)
    filterTransaction(tr, state) {
      const edit = tr.steps.some((step) => step instanceof ReplaceStep || step instanceof ReplaceAroundStep);
      const followUp = tr.getMeta("appendedTransaction") || tr.getMeta("addToHistory") === false || tr.getMeta(markdownLiveKey);
      if (tr.docChanged && edit && !followUp && !tr.getMeta(markdownInputKey) && markdownInputKey.getState(state)?.converted) closeHistory(tr);
      return true;
    },
    // Leaving a line of "# Title", "```lang" or "---" that was typed makes it its block (Typora).
    appendTransaction(trs, _old, state) {
      if (mode() !== "markdown" || !rules().blocks) return null;
      const value = markdownInputKey.getState(state);
      if (value?.edited === null || value?.edited === undefined) return null;
      if (!trs.some((tr) => tr.selectionSet || tr.docChanged)) return null;
      const { $head } = state.selection;
      const inLine = $head.parent.isTextblock && $head.depth > 0 && $head.before() === value.edited;
      if (inLine) return null;
      // the line is no longer the one typed in
      const tr = leftLineConversion(state, value.edited, rules());
      if (!tr) return state.tr.setMeta(markdownInputKey, { converted: true, left: true } satisfies InputMeta).setMeta("addToHistory", false);
      return closeHistory(tr).setMeta(markdownInputKey, { converted: true, left: true } satisfies InputMeta);
    },
    props: {
      handleTextInput(view, from, to, text, deflt) {
        if (mode() !== "markdown" || view.composing) return false;
        const state = view.state;
        const $from = state.doc.resolve(from);
        if (!$from.parent.isTextblock || $from.parent.type.spec.code) return false;
        if ($from.marks().some((m) => m.type.spec.code)) return false;
        // Markdown's escape: "\" and a punctuation character type the character, for good.
        if (from === to && text.length === 1 && ESCAPABLE.includes(text) && $from.parentOffset > 0) {
          const before = textBefore($from);
          const backslash = from - 1;
          if (before.endsWith("\\") && !isEscaped(state, backslash, from)) {
            const tr = markEscaped(state.tr.insertText(text, backslash, from), backslash);
            view.dispatch(tr.scrollIntoView());
            return true;
          }
        }
        if (text.length === 1 && (typedImage(view, from, to, text) || typedFootnote(view, from, to, text) || closeTag(view, from, to, text) || pairInput(view, from, to, text))) {
          // stepping over a "]" can close "[!NOTE]" too
          if (from === to && text === "]" && mayConvert(textBefore($from) + text)) convert(view, rules());
          return true;
        }
        if (!mayConvert(textBefore($from) + text)) return false;
        // The text goes in as typed (an ordinary step), then the conversion (a step of its own).
        view.dispatch(deflt());
        convert(view, rules());
        return true;
      },
      handleDOMEvents: {
        // What an input method commits is in the text by now: the same rules apply to it.
        compositionend(view) {
          if (mode() !== "markdown") return false;
          setTimeout(() => {
            if (!view.composing && mode() === "markdown" && !view.isDestroyed) convert(view, rules());
          });
          return false;
        },
      },
    },
  });
}

// ------------------------------------------------------------ Typora's keys --

/** Ctrl+1…6: the line a section of that level, or body text again (Typora). */
function headingKey(level: number): Command {
  return (state, dispatch) => {
    const block = state.selection.$from.parent;
    const { heading, paragraph } = state.schema.nodes;
    if (block.type !== paragraph && block.type !== heading) return false;
    const keep = { align: block.attrs.align ?? null, indent: 0 };
    if (level === 0 || (block.type === heading && block.attrs.level === level)) return block.type === paragraph ? true : setBlockType(paragraph, keep)(state, dispatch);
    return setBlockType(heading, { ...keep, level })(state, dispatch);
  };
}

/** Ctrl+= / Ctrl+-: a section a level up (body text becomes level 6) / down (level 6 becomes body text). */
function headingStep(up: boolean): Command {
  return (state, dispatch) => {
    const block = state.selection.$from.parent;
    const { heading, paragraph } = state.schema.nodes;
    if (block.type !== paragraph && block.type !== heading) return false;
    const level = block.type === heading ? block.attrs.level : 7;
    const next = up ? Math.max(1, level - 1) : level + 1;
    if (next === level) return true;
    return headingKey(next > 6 ? 0 : next)(state, dispatch);
  };
}

/** Ctrl+Shift+Q: a quote, or out of the quote. */
const quoteKey: Command = (state, dispatch) => {
  const { $from } = state.selection;
  for (let d = $from.depth; d > 0; d--) if ($from.node(d).type === state.schema.nodes.blockquote) return lift(state, dispatch);
  return wrapIn(state.schema.nodes.blockquote)(state, dispatch);
};

/** Ctrl+] / Ctrl+[: a list item a level deeper / higher. */
const listLevel = (deeper: boolean): Command => (state, dispatch) => {
  const item = state.schema.nodes.list_item;
  return (deeper ? sinkListItem(item) : liftListItem(item))(state, dispatch);
};

/** Backspace right after a closing character that came with its opening one: both go. */
const pairBackspace: Command = (state, dispatch) => {
  const { $cursor } = state.selection as TextSelection;
  if (!$cursor || !$cursor.parent.isTextblock) return false;
  const pairs = markdownInputKey.getState(state)?.pairs ?? [];
  if (!pairs.includes($cursor.pos)) return false;
  const prev = $cursor.parent.textBetween(Math.max(0, $cursor.parentOffset - 1), $cursor.parentOffset);
  const next = $cursor.parent.textBetween($cursor.parentOffset, Math.min($cursor.parent.content.size, $cursor.parentOffset + 1));
  if (!prev || PAIRS[prev] !== next) return false;
  if (dispatch) dispatch(state.tr.delete($cursor.pos - 1, $cursor.pos + 1).scrollIntoView());
  return true;
};

/**
 * Backspace at the start of a list item after the first: the item's text joins the item before as a paragraph of its
 * own (Typora: "- a" "- b" -> "- a" + "  b"); the next Backspace joins the two lines.
 */
const joinListItemBackward: Command = (state, dispatch, view) => {
  const { $cursor } = state.selection as TextSelection;
  const { list_item } = state.schema.nodes;
  if (!$cursor || $cursor.parentOffset > 0 || $cursor.depth < 3) return false;
  if ($cursor.node(-1).type !== list_item || $cursor.index(-1) !== 0 || $cursor.index(-2) === 0) return false;
  return joinBackward(state, dispatch, view);
};

/**
 * Backspace at the start of a line right after a list, a quote or the abstract: the line joins the last line in it (Typora: "- a"
 * then "p" -> "- ap").
 */
const joinIntoContainerBefore: Command = (state, dispatch) => {
  const { $cursor } = state.selection as TextSelection;
  if (!$cursor || $cursor.parentOffset > 0 || !$cursor.parent.isTextblock || $cursor.depth < 1) return false;
  const { bullet_list, ordered_list, blockquote, abstract } = state.schema.nodes;
  const index = $cursor.index(-1);
  if (index === 0) return false;
  const container = $cursor.node(-1).child(index - 1);
  if (container.type !== bullet_list && container.type !== ordered_list && container.type !== blockquote && container.type !== abstract) return false;
  // the last line inside it
  let last: Node = container;
  let end = $cursor.before() - 1;
  while (!last.isTextblock) {
    const child = last.lastChild;
    if (!child) return false;
    end -= 1;
    last = child;
  }
  if (last.type.spec.code) return false;
  if (!dispatch) return true;
  const block = $cursor.parent;
  const tr = state.tr.delete($cursor.before(), $cursor.after());
  tr.insert(end, block.content);
  tr.setSelection(TextSelection.create(tr.doc, end));
  dispatch(tr.scrollIntoView());
  return true;
};

/** Backspace at the start of a table cell: the end of the cell before it (Typora); in the first cell, nothing. */
const cellBackspace: Command = (state, dispatch) => {
  const { $cursor } = state.selection as TextSelection;
  if (!$cursor || $cursor.parentOffset > 0) return false;
  for (let d = $cursor.depth - 1; d > 0; d--) {
    const name = $cursor.node(d).type.name;
    if (name !== "table_cell" && name !== "table_header") continue;
    // only from the cell's first line
    if ($cursor.index(d) !== 0) return false;
    const before = Selection.findFrom(state.doc.resolve($cursor.before(d)), -1, true);
    if (dispatch && before && before.from > $cursor.before(d - 2)) dispatch(state.tr.setSelection(before).scrollIntoView());
    return true;
  }
  return false;
};

/** Backspace in an empty first line of the document with more after it: the line goes (Typora). */
const emptyFirstLineBackspace: Command = (state, dispatch) => {
  const { $cursor } = state.selection as TextSelection;
  if (!$cursor || $cursor.depth !== 1 || $cursor.index(0) !== 0 || $cursor.parent.content.size > 0 || state.doc.childCount < 2) return false;
  if ($cursor.parent.type !== state.schema.nodes.paragraph) return false;
  if (dispatch) {
    const tr = state.tr.delete(0, $cursor.parent.nodeSize);
    dispatch(tr.setSelection(Selection.atStart(tr.doc)).scrollIntoView());
  }
  return true;
};

/** Ctrl+L: the line's text selected; Ctrl+D: the word at the caret; Ctrl+E: the formatted span (or the cell). */
const selectLine: Command = (state, dispatch) => {
  const { $from } = state.selection;
  if (!$from.parent.isTextblock) return false;
  if (dispatch) dispatch(state.tr.setSelection(TextSelection.create(state.doc, $from.start(), $from.end())));
  return true;
};

const selectWord: Command = (state, dispatch) => {
  const { $from, empty } = state.selection;
  if (!$from.parent.isTextblock || !empty) return false;
  const text = lineText($from.parent);
  let a = $from.parentOffset;
  let b = a;
  const word = /[\p{L}\p{N}_]/u;
  while (a > 0 && word.test(text[a - 1])) a--;
  while (b < text.length && word.test(text[b])) b++;
  if (a === b) return false;
  if (dispatch) dispatch(state.tr.setSelection(TextSelection.create(state.doc, $from.start() + a, $from.start() + b)));
  return true;
};

const selectScope: Command = (state, dispatch) => {
  const { $from } = state.selection;
  const shown = shownMarkdown(state);
  if (shown && $from.pos >= shown.from && $from.pos <= shown.to) {
    if (dispatch) dispatch(state.tr.setSelection(TextSelection.create(state.doc, shown.from, shown.to)));
    return true;
  }
  for (let d = $from.depth; d > 0; d--) {
    const name = $from.node(d).type.name;
    if (name === "table_cell" || name === "table_header") {
      const sel = TextSelection.between(state.doc.resolve($from.start(d)), state.doc.resolve($from.end(d)));
      if (dispatch) dispatch(state.tr.setSelection(sel));
      return true;
    }
  }
  return selectLine(state, dispatch);
};

/** ⌥⌘X (Typora on the Mac): the line a task, or a task no more. */
const taskKey: Command = (state, dispatch) => {
  const { $from } = state.selection;
  const { list_item, bullet_list } = state.schema.nodes;
  for (let d = $from.depth; d > 0; d--) {
    const node = $from.node(d);
    if (node.type !== list_item) continue;
    if (dispatch) dispatch(state.tr.setNodeAttribute($from.before(d), "checked", node.attrs.checked === null ? false : null));
    return true;
  }
  let wrapped: Transaction | null = null;
  toggleList("bullet_list")(state, (tr) => {
    wrapped = tr;
  });
  if (!wrapped) return false;
  const tr = wrapped as Transaction;
  const $in = tr.selection.$from;
  for (let d = $in.depth; d > 0; d--) {
    if ($in.node(d).type !== list_item || $in.node(d - 1).type !== bullet_list) continue;
    tr.setNodeAttribute($in.before(d), "checked", false);
    break;
  }
  if (dispatch) dispatch(tr);
  return true;
};

const MAC = typeof navigator !== "undefined" && /Mac|iPhone|iPad|iPod/.test(navigator.platform);

/** Typora's keys, in Markdown input (its macOS keys on the Mac: ⌥⌘ for blocks, ⌃` for code). */
export function markdownKeymap(mode: () => InputMode, keys: { bold: Command; italic: Command; underline: Command; code: Command; strike: Command; link: Command }): Plugin {
  const only = (command: Command): Command => (state, dispatch, view) => mode() === "markdown" && command(state, dispatch, view);
  const mac: Record<string, Command> = MAC
    ? {
        "Alt-Mod-q": only(quoteKey),
        "Alt-Mod-o": only(toggleList("ordered_list")),
        "Alt-Mod-u": only(toggleList("bullet_list")),
        "Alt-Mod-x": only(taskKey),
        "Alt-Mod-c": only(insertCodeBlock("")),
        "Alt-Mod-b": only(insertDisplayLatex("")),
        "Alt-Mod-t": only(insertTable(3, 3)),
        "Ctrl-`": only(typoraFormat("code", keys.code)),
        "Shift-Ctrl-`": only(typoraFormat("strike", keys.strike)),
        "Ctrl-Mod-i": only((_state, _dispatch, view) => !!view && requestImage(view)),
      }
    : {};
  return keymap({
    ...mac,
    "Mod-0": only(headingKey(0)),
    "Mod-1": only(headingKey(1)),
    "Mod-2": only(headingKey(2)),
    "Mod-3": only(headingKey(3)),
    "Mod-4": only(headingKey(4)),
    "Mod-5": only(headingKey(5)),
    "Mod-6": only(headingKey(6)),
    "Mod-=": only(headingStep(true)),
    "Mod-+": only(headingStep(true)),
    "Mod--": only(headingStep(false)),
    "Mod-t": only(insertTable(3, 3)),
    // on the Mac ⌘⇧K stays the keypad's (Typora's code block there is ⌥⌘C)
    ...(MAC ? {} : { "Shift-Mod-k": only(insertCodeBlock("")) }),
    "Shift-Mod-m": only(insertDisplayLatex("")),
    "Shift-Mod-q": only(quoteKey),
    "Shift-Mod-[": only(toggleList("ordered_list")),
    "Shift-Mod-]": only(toggleList("bullet_list")),
    "Mod-]": only(listLevel(true)),
    "Mod-[": only(listLevel(false)),
    "Shift-Mod-i": only((_state, _dispatch, view) => !!view && requestImage(view)),
    "Mod-b": only(typoraFormat("strong", keys.bold)),
    "Mod-i": only(typoraFormat("em", keys.italic)),
    "Mod-u": only(typoraFormat("underline", keys.underline)),
    "Shift-Mod-`": only(typoraFormat("code", keys.code)),
    "Shift-Alt-5": only(typoraFormat("strike", keys.strike)),
    "Mod-k": only(typoraLink(keys.link)),
    "Mod-\\": only(clearFormat),
    "Mod-l": only(selectLine),
    "Mod-d": only(selectWord),
    "Mod-e": only(selectScope),
    Backspace: only(
      (state, dispatch, view) =>
        pairBackspace(state, dispatch, view) ||
        cellBackspace(state, dispatch, view) ||
        joinListItemBackward(state, dispatch, view) ||
        joinIntoContainerBefore(state, dispatch, view) ||
        emptyFirstLineBackspace(state, dispatch, view),
    ),
  });
}

