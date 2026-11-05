// Markdown input (DESIGN.md §7): in Markdown input mode, Markdown typed in the text formats it, as in Typora and the
// note apps. At the start of a line "# " makes a section, "- " / "1. " a list, "> " a quote, "```lang " a code
// block; a line of "```lang", "---" or "| a | b |" becomes a code block, a rule or a table on Enter; **bold**,
// *italic*, `code`, ~~strike~~, ==highlight== and [text](url) become marks once their closing delimiter is typed;
// "\" before a punctuation character types that character (Markdown's escape), which then never formats anything.
// What was typed stays in the history: each conversion is an undo step of its own, so Ctrl+Z gives the typed
// text back.
import { Fragment, type MarkType, type Node, type ResolvedPos } from "prosemirror-model";
import { Plugin, PluginKey, TextSelection, type Command, type EditorState, type Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { closeHistory } from "prosemirror-history";
import { canJoin, findWrapping } from "prosemirror-transform";
import { lift, setBlockType, toggleMark, wrapIn } from "prosemirror-commands";
import { keymap } from "prosemirror-keymap";
import { insertTable, toggleList } from "./blocks.ts";
import { insertDisplayLatex } from "./commands.ts";
import { newNodeId } from "./ids.ts";
import { MARKDOWN_HIGHLIGHT, blockLanguage } from "./markdown.ts";

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
  /** characters typed escaped ("\*"): positions of literal characters that never act as Markdown */
  escaped: number[];
}

type InputMeta = { converted: true } | { escaped: number };

export const markdownInputKey = new PluginKey<MarkdownInputState>("visualtex-document-markdown-input");

/** CommonMark's escapable characters: ASCII punctuation. */
const ESCAPABLE = "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~";
/** The space that ends a line-start marker: a space, a no-break space or an input method's full-width space. */
const SPACE = "[ \\u00a0\\u3000]";
const OBJECT = "\ufffc";

const isEscaped = (state: EditorState, from: number, to: number) =>
  (markdownInputKey.getState(state)?.escaped ?? []).some((pos) => pos >= from && pos < to);

/** How far back a mark's delimiters are looked for. */
const MAX_MATCH = 500;

/** The textblock's text before `$pos` (formulas and other inline objects one character each, as in positions). */
function textBefore($pos: ResolvedPos): string {
  return $pos.parent.textBetween(Math.max(0, $pos.parentOffset - MAX_MATCH), $pos.parentOffset, undefined, OBJECT);
}

/** The marks typing at `$pos` gives (a mark just closed is not among them). */
const marksAt = (state: EditorState, $pos: ResolvedPos) => state.storedMarks ?? $pos.marks();

// ------------------------------------------------------- line-start markers --

/** The line-start markers; with `chinese` also what a Chinese input method types for them ("1。", "》", "···",
 * the full-width space). */
function lineStartRules(chinese: boolean) {
  const space = chinese ? SPACE : "[ \\u00a0]";
  return {
    heading: new RegExp(`^(#{1,6})${space}$`),
    bullet: new RegExp(`^[-*+]${space}$`),
    ordered: new RegExp(`^(\\d{1,9})${chinese ? "[.)。．）]" : "[.)]"}${space}$`),
    quote: new RegExp(`^${chinese ? "[>》]" : ">"}${space}$`),
    fence: new RegExp(`^(?:${chinese ? "\`\`\`|~~~|···" : "\`\`\`|~~~"})([^\\s\`]*)${space}$`),
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

/** The textblock becomes `type` (its text after the marker stays), when its parent allows that. */
function retype(state: EditorState, $cursor: ResolvedPos, type: Node["type"], attrs: Record<string, unknown>): Transaction | null {
  if (!$cursor.node(-1).canReplaceWith($cursor.index(-1), $cursor.index(-1) + 1, type)) return null;
  const start = $cursor.start();
  const tr = state.tr.delete(start, $cursor.pos);
  tr.setBlockType(start, start, type, attrs);
  return tr;
}

/** A paragraph after the block holding `inside`, when nothing follows it (the text goes on after a code block). */
function paragraphAfter(tr: Transaction, inside: number) {
  const after = tr.doc.resolve(inside).after();
  if (after >= tr.doc.content.size || !tr.doc.resolve(after).nodeAfter) tr.insert(after, tr.doc.type.schema.nodes.paragraph.create());
}

const TASK_PLAIN = /^\[([ xX])\][ \u00a0]$/;
const TASK_CHINESE = /^[\[【]([ xX])[\]】][ \u00a0\u3000]$/;

function lineStartConversion(state: EditorState, $cursor: ResolvedPos, before: string, rules: MarkdownRuleSwitches): Transaction | null {
  if ($cursor.parentOffset > 16 || !rules.blocks) return null;
  const block = $cursor.parent;
  const { nodes } = state.schema;
  const start = $cursor.start();
  if (isEscaped(state, start, $cursor.pos)) return null;
  const { heading: HEADING, bullet: BULLET, ordered: ORDERED, quote: QUOTE, fence: FENCE } = rules.chinese ? LINE_START.chinese : LINE_START.plain;
  let match: RegExpExecArray | null;
  if ((match = HEADING.exec(before))) {
    if (block.type !== nodes.paragraph && block.type !== nodes.heading) return null;
    const level = Math.min(6, match[1].length);
    return retype(state, $cursor, nodes.heading, { id: block.attrs.id, level, align: block.attrs.align ?? null, indent: 0 });
  }
  if (block.type !== nodes.paragraph) return null;
  // "[ ] " / "[x] " opening a list item: a task (a Chinese input method types 【 】)
  const task = (rules.chinese ? TASK_CHINESE : TASK_PLAIN).exec(before);
  if (task && $cursor.depth >= 2 && $cursor.node(-1).type === nodes.list_item && $cursor.index(-1) === 0 && $cursor.node(-1).attrs.checked === null) {
    return state.tr.delete(start, $cursor.pos).setNodeAttribute($cursor.before(-1), "checked", task[1] !== " ");
  }
  if (BULLET.test(before)) return wrapBlock(state, $cursor, nodes.bullet_list, null);
  if ((match = ORDERED.exec(before))) return wrapBlock(state, $cursor, nodes.ordered_list, { start: Number(match[1]) });
  if (QUOTE.test(before)) return wrapBlock(state, $cursor, nodes.blockquote, null);
  if ((match = FENCE.exec(before))) {
    const tr = retype(state, $cursor, nodes.code_block, { id: newNodeId(), language: blockLanguage(match[1]) });
    if (tr) paragraphAfter(tr, start);
    return tr;
  }
  return null;
}

// ------------------------------------------------------------- inline marks --

interface MarkRule {
  /** group 1: what precedes the opening delimiter (kept); group 2: the text */
  pattern: RegExp;
  delimiter: number;
  mark: (state: EditorState) => MarkType | undefined;
  attrs?: Record<string, unknown>;
}

// Not right after a Latin letter, a digit, ":" or "/" (Milkdown's guard): "2*3*" and addresses stay text; next to
// Chinese text the marks still work.
const MARK_RULES: MarkRule[] = [
  { pattern: /(^|[^*\\A-Za-z0-9:/])\*\*([^*\s](?:[^*]*[^*\s])?)\*\*$/, delimiter: 2, mark: (s) => s.schema.marks.strong },
  { pattern: /(^|[^\w\\])__([^_\s](?:[^_]*[^_\s])?)__$/, delimiter: 2, mark: (s) => s.schema.marks.strong },
  { pattern: /(^|[^*\\A-Za-z0-9:/])\*([^*\s](?:[^*]*[^*\s])?)\*$/, delimiter: 1, mark: (s) => s.schema.marks.em },
  { pattern: /(^|[^\w\\])_([^_\s](?:[^_]*[^_\s])?)_$/, delimiter: 1, mark: (s) => s.schema.marks.em },
  { pattern: /(^|[^~\\A-Za-z0-9:/])~~([^~\s](?:[^~]*[^~\s])?)~~$/, delimiter: 2, mark: (s) => s.schema.marks.strike },
  { pattern: /(^|[^=\\A-Za-z0-9:/])==([^=\s](?:[^=]*[^=\s])?)==$/, delimiter: 2, mark: (s) => s.schema.marks.highlight, attrs: { color: MARKDOWN_HIGHLIGHT } },
  { pattern: /(^|[^`\\])`([^`]+)`$/, delimiter: 1, mark: (s) => s.schema.marks.code },
];

const LINK = /(^|[^!\\])\[([^\]]+)\]\(([^()\s]+)(?:\s+"([^"]*)")?\)$/;
/** inside a backtick span still being typed, or a link's address: the characters are literal so far */
const OPEN_CODE = (text: string) => ((text.match(/`/g)?.length ?? 0) % 2 === 1);
const IN_LINK_ADDRESS = /\]\([^()\s]*$/;

/** The text between the delimiters takes the mark; typing after it is plain again. */
function markConversion(state: EditorState, $cursor: ResolvedPos, before: string, rules: MarkdownRuleSwitches): Transaction | null {
  if (!rules.marks || marksAt(state, $cursor).some((m) => m.type.spec.code)) return null;
  const caret = $cursor.pos;
  const link = LINK.exec(before);
  if (link && state.schema.marks.link) {
    const open = caret - (before.length - link.index - link[1].length);
    if (isEscaped(state, open, open + 1)) return null;
    const textEnd = open + 1 + link[2].length;
    const tr = state.tr.delete(textEnd, caret).delete(open, open + 1);
    const type = state.schema.marks.link;
    tr.addMark(open, textEnd - 1, type.create({ href: link[3], title: link[4] ?? null }));
    return tr.removeStoredMark(type);
  }
  for (const rule of MARK_RULES) {
    const match = rule.pattern.exec(before);
    const type = match && rule.mark(state);
    if (!match || !type) continue;
    const open = caret - (before.length - match.index - match[1].length);
    const head = before.slice(0, match.index + match[1].length);
    // Typed inside a code span or a link address, the delimiters are characters.
    if (type !== state.schema.marks.code && (OPEN_CODE(head) || IN_LINK_ADDRESS.test(head))) continue;
    if (isEscaped(state, open, open + rule.delimiter) || isEscaped(state, caret - rule.delimiter, caret)) continue;
    if (!$cursor.parent.type.allowsMarkType(type)) continue;
    const tr = state.tr.delete(caret - rule.delimiter, caret).delete(open, open + rule.delimiter);
    tr.addMark(open, caret - 2 * rule.delimiter, type.create(rule.attrs ?? null));
    return tr.removeStoredMark(type);
  }
  return null;
}

// --------------------------------------------------------------- the plugin --

/** The conversion the text before the caret calls for, if any. */
function conversionAtCaret(state: EditorState, rules: MarkdownRuleSwitches): Transaction | null {
  const { $cursor } = state.selection as TextSelection;
  if (!$cursor || !$cursor.parent.isTextblock || $cursor.parent.type.spec.code) return null;
  const before = textBefore($cursor);
  return lineStartConversion(state, $cursor, before, rules) ?? markConversion(state, $cursor, before, rules);
}

/** Whether a conversion could follow the text `before` (a quick test before the typed text goes in). */
function mayConvert(before: string): boolean {
  if (before.length <= 17 && (Object.values(LINE_START.chinese).some((rule) => rule.test(before)) || TASK_CHINESE.test(before))) return true;
  return LINK.test(before) || MARK_RULES.some((rule) => rule.pattern.test(before));
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

export function markdownInputPlugin(mode: () => InputMode, rules: () => MarkdownRuleSwitches = () => ALL_RULES): Plugin<MarkdownInputState> {
  return new Plugin<MarkdownInputState>({
    key: markdownInputKey,
    state: {
      init: () => ({ converted: false, escaped: [] }),
      apply(tr, value) {
        const meta = tr.getMeta(markdownInputKey) as InputMeta | undefined;
        let escaped = value.escaped;
        if (tr.docChanged && escaped.length) {
          // A literal character keeps its place while it is there.
          escaped = escaped
            .map((pos) => ({ from: tr.mapping.map(pos, 1), to: tr.mapping.map(pos + 1, -1) }))
            .filter((range) => range.to - range.from === 1)
            .map((range) => range.from);
        }
        if (meta && "escaped" in meta) escaped = [...escaped, meta.escaped].slice(-64);
        const converted = meta && "converted" in meta ? true : tr.docChanged || tr.selectionSet ? false : value.converted;
        return converted === value.converted && escaped === value.escaped ? value : { converted, escaped };
      },
    },
    // The change after a conversion starts an undo step of its own: Ctrl+Z then takes back just that change.
    filterTransaction(tr, state) {
      if (tr.docChanged && !tr.getMeta(markdownInputKey) && !tr.getMeta("appendedTransaction") && markdownInputKey.getState(state)?.converted) {
        closeHistory(tr);
      }
      return true;
    },
    props: {
      handleTextInput(view, from, to, text, deflt) {
        if (mode() !== "markdown" || view.composing) return false;
        const state = view.state;
        const $from = state.doc.resolve(from);
        if (!$from.parent.isTextblock || $from.parent.type.spec.code) return false;
        if (marksAt(state, $from).some((m) => m.type.spec.code)) return false;
        // Markdown's escape: "\" and a punctuation character type the character, for good.
        if (from === to && text.length === 1 && ESCAPABLE.includes(text) && $from.parentOffset > 0) {
          const before = textBefore($from);
          const backslash = from - 1;
          if (before.endsWith("\\") && !isEscaped(state, backslash, from)) {
            const tr = state.tr.insertText(text, backslash, from).setMeta(markdownInputKey, { escaped: backslash } satisfies InputMeta);
            view.dispatch(tr.scrollIntoView());
            return true;
          }
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

// ------------------------------------------------------------------- Enter --

const ENTER_FENCE = /^(?:```|~~~|···)[ \t]*([^\s`]*)[ \t]*$/;
const ENTER_FENCE_PLAIN = /^(?:```|~~~)[ \t]*([^\s`]*)[ \t]*$/;
const ENTER_RULE = /^(?:(?:-[ \t]*){3,}|(?:\*[ \t]*){3,}|(?:_[ \t]*){3,})$/;
const ENTER_TABLE = /^\|(.*)\|$/;

/** "| a | b |" -> ["a", "b"] ("\|" is a character of its cell). */
function tableCells(inner: string): string[] {
  return inner.split(/(?<!\\)\|/).map((cell) => cell.trim().replace(/\\\|/g, "|"));
}

function enterConversion(state: EditorState, $cursor: ResolvedPos, line: string, rules: MarkdownRuleSwitches): Transaction | null {
  const { nodes } = state.schema;
  const parent = $cursor.node(-1);
  const index = $cursor.index(-1);
  const before = $cursor.before();
  const after = $cursor.after();
  let match: RegExpExecArray | null;
  if ((match = (rules.chinese ? ENTER_FENCE : ENTER_FENCE_PLAIN).exec(line))) {
    if (!parent.canReplaceWith(index, index + 1, nodes.code_block)) return null;
    const tr = state.tr.delete($cursor.start(), $cursor.pos);
    tr.setBlockType(before + 1, before + 1, nodes.code_block, { id: newNodeId(), language: blockLanguage(match[1]) });
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
  return null;
}

/** Enter at the end of a line of Markdown that makes a block ("```lang", "---", "| a | b |"): that block. */
export function markdownEnter(mode: () => InputMode, rules: () => MarkdownRuleSwitches = () => ALL_RULES): Command {
  return (state, dispatch) => {
    if (mode() !== "markdown" || !rules().blocks) return false;
    const { $cursor } = state.selection as TextSelection;
    if (!$cursor || $cursor.parent.type !== state.schema.nodes.paragraph) return false;
    if ($cursor.parentOffset !== $cursor.parent.content.size) return false;
    if (isEscaped(state, $cursor.start(), $cursor.pos)) return false;
    const line = $cursor.parent.textBetween(0, $cursor.parent.content.size, undefined, OBJECT);
    const tr = enterConversion(state, $cursor, line, rules());
    if (!tr) return false;
    if (dispatch) dispatch(asConversion(tr));
    return true;
  };
}

// -------------------------------------------------------------- Typora's keys --

/** Ctrl+1/2/3: the line a section of that level, or body text again (Typora). */
function headingKey(level: number): Command {
  return (state, dispatch) => {
    const block = state.selection.$from.parent;
    const { heading, paragraph } = state.schema.nodes;
    if (block.type !== paragraph && block.type !== heading) return false;
    const keep = { align: block.attrs.align ?? null, indent: 0 };
    if (block.type === heading && block.attrs.level === level) return setBlockType(paragraph, keep)(state, dispatch);
    return setBlockType(heading, { ...keep, level })(state, dispatch);
  };
}

/** Ctrl+Shift+Q: a quote, or out of the quote. */
const quoteKey: Command = (state, dispatch) => {
  const { $from } = state.selection;
  for (let d = $from.depth; d > 0; d--) if ($from.node(d).type === state.schema.nodes.blockquote) return lift(state, dispatch);
  return wrapIn(state.schema.nodes.blockquote)(state, dispatch);
};

/** Typora's formatting keys, in Markdown input (Ctrl+0 and Ctrl+Shift+K stay the app's zoom and keypad keys). */
export function markdownKeymap(mode: () => InputMode): Plugin {
  const only = (command: Command): Command => (state, dispatch, view) => mode() === "markdown" && command(state, dispatch, view);
  return keymap({
    "Mod-1": only(headingKey(1)),
    "Mod-2": only(headingKey(2)),
    "Mod-3": only(headingKey(3)),
    "Mod-t": only(insertTable(3, 3)),
    "Shift-Mod-m": only(insertDisplayLatex("")),
    "Shift-Mod-q": only(quoteKey),
    "Shift-Mod-[": only(toggleList("ordered_list")),
    "Shift-Mod-]": only(toggleList("bullet_list")),
    "Shift-Mod-`": only((state, dispatch) => toggleMark(state.schema.marks.code)(state, dispatch)),
    "Shift-Alt-5": only((state, dispatch) => toggleMark(state.schema.marks.strike)(state, dispatch)),
  });
}
