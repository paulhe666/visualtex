// The document editor's plugins in one place (MATHCHA-SPEC.md): Mathcha's triggers ($, the "\" box) and keys,
// paste, ids, formulas, numbering, tables, gap cursor, history.
import { mixedSelectionPlugin, toggleDocumentBold } from "./mixedSelection.ts";
import { Fragment, Slice, type Node, type Schema } from "prosemirror-model";
import { EditorState, NodeSelection, Plugin, PluginKey, TextSelection, type Command, type Transaction } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";
import { closeHistory, history, redo, undo } from "prosemirror-history";
import { keymap } from "prosemirror-keymap";
import { baseKeymap, chainCommands, toggleMark } from "prosemirror-commands";
import { gapCursor } from "prosemirror-gapcursor";
import { dropCursor } from "prosemirror-dropcursor";
import { tableEditing } from "prosemirror-tables";
import { tableColumnsPlugin } from "./tableColumns.ts";
import {
  deleteFormulaBackward,
  deleteFormulaForward,
  enterFormulaBackward,
  enterFormulaDown,
  enterFormulaForward,
  enterFormulaUp,
  insertDisplayLatex,
  insertInlineFormula,
  openSelectedFormula,
} from "./commands.ts";
import { CLOSED_BY_DOLLAR, findNodeById, formulaPlugin, openFormula, openFormulaId, type FormulaHost } from "./formulas.ts";
import { assignIds, idsPlugin, newNodeId } from "./ids.ts";
import { hasMathDelimiters, parseMathText, type MathTextBlock } from "./mathText.ts";
import {
  changeSectionLevel,
  clearLineStyle,
  currentListKind,
  goToCell,
  indent,
  insertTabCharacter,
  leaveEmptyQuoteLine,
  sectionLevel,
  selectAdjacentBlock,
  setAlign,
  splitKeepingStyle,
  splitListItemOrLeave,
  tableEnter,
  tableRowBelow,
  toggleList,
  toggleSection,
} from "./blocks.ts";
import { commandMenuPlugin, type MathCommandProvider } from "./suggest.ts";
import { codeBlockKeys, codeBlockPlugin, type CodeBlockOptions, type CodeHighlighter } from "./codeBlocks.ts";
import { richBlocksPlugin } from "./richBlocks.ts";
import { looksLikeMarkdown, looksLikeMarkdownDocument, markdownSourceKey, markdownSourcePlugin, markdownToSlice } from "./markdown.ts";
import { ALL_RULES, markdownEnter, markdownHeadingEnter, markdownInputPlugin, markdownKeymap, type InputMode, type MarkdownRuleSwitches } from "./markdownInput.ts";
import { markdownRedo, markdownUndo, sliceWithoutMarkdownSyntax } from "./markdownLive.ts";
import { blockSyntax, clearSourceFormat, inSourceCode, insertSourceLink, setSourceHeading, sourceLiveKey, sourceLivePlugin, toggleSourceFormat, toSourceSlice } from "./sourceMarkdown.ts";
import { codeLineEnter, indentLines, outdentLines, sourceEnter, sourcePlainEnter, sourceShiftEnter } from "./sourceKeys.ts";
import { documentDragPlugin } from "./documentDrag.ts";
import { latexToMarkdown, looksLikeLatexSource } from "./latexImport.ts";
import { requestLink } from "./links.ts";
import { bibliographyEntries, insertFootnote, noteText, notesPlugin } from "./notes.ts";
import { bracketLabel, equationLabels, headingNumbers, ownEquationTag } from "./settings.ts";
import { paginationPlugin, type PaginationOptions } from "./pagination.ts";
import { compositionGuardPlugin } from "./composition.ts";
import { imagesPlugin, type ImageLabels } from "./images.ts";
import { boardFixes, DEFAULT_BOARD_LABELS, imageBoardsPlugin, type BoardLabels } from "./imageBoards.ts";
import { DEFAULT_CAPTION_LABELS, figureFixes, figuresPlugin, type CaptionLabels } from "./figures.ts";
import { pageCommandLatex } from "./pageCommands.ts";
import { abstractLabelPlugin, leaveEmptyAuthorLine } from "./titleBlock.ts";

// --------------------------------------------------------------- triggers --

interface DollarState {
  /** "$" just closed a formula: how many "$" typed next are the rest of its closing "$$" or "$$$$" */
  closing: number;
  /** Markdown input: the position of a "$" typed as a character, waiting for what comes after it */
  pending: number | null;
  /** Markdown input: "$…$" an edit made in a line (not by typing its closing "$"), a formula once the caret leaves it */
  waiting: ReadonlyArray<{ from: number; to: number }>;
}

type DollarMeta = { closing: number } | { pending: number | null };

/** "$…$" in a line's text (Pandoc's rule, as Markdown is read: sourceMarkdown.ts INLINE_MATH); objects are "￼". */
const DOLLAR_SPAN = /(?<![\\$])\$(?!\$)([^\s$￼](?:[^$\n￼]*?[^\s$\\￼])?)\$(?![$\d])/g;

/** An edit the user made (typing, a deletion, a cut): not undo, a paste or drop, a composition, the source tab. */
function userEdit(tr: Transaction): boolean {
  if (!tr.docChanged || tr.getMeta("history$") || tr.getMeta("composition") !== undefined) return false;
  if (tr.getMeta("appendedTransaction") || tr.getMeta("addToHistory") === false || tr.getMeta(markdownSourceKey) !== undefined) return false;
  const ui = tr.getMeta("uiEvent");
  return ui !== "paste" && ui !== "drop";
}

/** The ranges a transaction changed, in its resulting document, and whether each was only a deletion. */
function changedRanges(tr: Transaction): Array<{ from: number; to: number; deleted: boolean }> {
  const ranges: Array<{ from: number; to: number; deleted: boolean }> = [];
  tr.mapping.maps.forEach((map, index) => {
    map.forEach((oldStart, oldEnd, newStart, newEnd) => {
      // into the transaction's final document
      const rest = tr.mapping.slice(index + 1);
      ranges.push({ from: rest.map(newStart, -1), to: rest.map(newEnd, 1), deleted: newEnd === newStart && oldEnd > oldStart });
    });
  });
  return ranges;
}

/** Whether a position of a textblock is where "$" are characters: code (a fence, a code span, the code mark). */
function inDollarCode(state: EditorState, pos: number): boolean {
  const $pos = state.doc.resolve(pos);
  if (!$pos.parent.isTextblock || $pos.parent.type.spec.code || inSourceCode($pos)) return true;
  return $pos.marks().some((m) => m.type.spec.code || m.type.name === "code");
}

/** The "$…$" spans of the textblock around `pos` that reach into [from, to] (document positions). */
function dollarSpansNear(state: EditorState, pos: number, from: number, to: number): Array<{ from: number; to: number }> {
  const $pos = state.doc.resolve(pos);
  const block = $pos.parent;
  if (!block.isTextblock || inDollarCode(state, pos)) return [];
  const start = $pos.start();
  const text = block.textBetween(0, block.content.size, undefined, "￼");
  const spans: Array<{ from: number; to: number }> = [];
  for (const m of text.matchAll(DOLLAR_SPAN)) {
    const a = start + m.index!;
    const b = a + m[0].length;
    if (b < from || a > to) continue;
    // (not in a code span)
    if ((text.slice(0, m.index!).match(/`/g) ?? []).length % 2 === 1) continue;
    spans.push({ from: a, to: b });
  }
  return spans;
}

/** A span still "$…$" by the rule, whole, in a line where "$" makes formulas. */
function stillDollarSpan(state: EditorState, span: { from: number; to: number }): string | null {
  if (span.to > state.doc.content.size || span.from >= span.to) return null;
  const $from = state.doc.resolve(span.from);
  if (!$from.sameParent(state.doc.resolve(span.to)) || inDollarCode(state, span.from)) return null;
  const found = dollarSpansNear(state, span.from, span.from, span.to).find((s) => s.from === span.from && s.to === span.to);
  return found ? state.doc.textBetween(span.from + 1, span.to - 1) : null;
}

/**
 * The caret between two dollars a deletion left side by side ("$ $" with its space deleted, "$x$" with its x): their
 * position, the "$$" that opens an inline formula when typed. Not "\$", a "$" of "$$$", code.
 */
function caretBetweenDollars(state: EditorState): number | null {
  const sel = state.selection;
  if (!(sel instanceof TextSelection) || !sel.empty || inDollarCode(state, sel.head)) return null;
  const $head = sel.$head;
  const text = $head.parent.textBetween(0, $head.parent.content.size, undefined, "￼");
  const before = text.slice(0, $head.parentOffset);
  const after = text.slice($head.parentOffset);
  if (!/[$￥]$/.test(before) || !/^[$￥]/.test(after) || /[$￥]{2}$/.test(before) || /^[$￥]{2}/.test(after)) return null;
  if (/(?:^|[^\\])(?:\\\\)*\\[$￥]$/.test(before) || (before.match(/`/g) ?? []).length % 2 === 1) return null;
  return sel.head;
}

/**
 * Formulas typed with dollars. In Markdown input "$$" opens an inline formula and "$$$$" a display formula: the
 * first "$" is a character until the second one comes right after it, and the third and fourth reach the new inline
 * formula, which then gives way to a display formula (DocumentFormulaField). A single "$" stays text ("$5"), "\$"
 * too; "￥", the same key (Shift+4) with a Chinese input method's punctuation, counts as "$". Inside code they stay
 * characters. A "$" that closes "$…$" typed the LaTeX way (text after an earlier "$" in the line, no space just inside
 * the dollars: the rule Markdown is read by) makes that text a formula. (It stayed text, and in "$a$$b$" the first
 * formula's closing "$" and the second's opening one were taken for "$$": "$a" was left as text, a formula opened
 * there, and what was typed next, the sentence's words too, went into it, 2026-11-11.) Right after "$" closed a formula, the rest of its closing delimiter does nothing: one more "$" after an
 * inline formula, up to three after a display formula. (LaTeX input: "$" opens an inline formula at once.)
 */
function dollarPlugin(inputMode: () => InputMode) {
  const key = new PluginKey<DollarState>("visualtex-document-dollar");
  const isDollar = (text: string) => text === "$" || text === "￥";
  return new Plugin<DollarState>({
    key,
    state: {
      init: () => ({ closing: 0, pending: null, waiting: [] }),
      apply: (tr, value, _old, state) => {
        // "$…$" made by an edit wait for the caret to leave them (mapped; ones the edit reaches added)
        let waiting = value.waiting;
        if (inputMode() !== "markdown") waiting = waiting.length ? [] : waiting;
        else if (tr.docChanged) {
          if (waiting.length) {
            waiting = waiting
              .map((s) => ({ from: tr.mapping.map(s.from, 1), to: tr.mapping.map(s.to, -1) }))
              .filter((s) => stillDollarSpan(state, s) !== null);
          }
          if (userEdit(tr)) {
            for (const range of changedRanges(tr)) {
              for (const span of dollarSpansNear(state, range.from, range.from, range.to)) {
                if (!waiting.some((s) => s.from === span.from && s.to === span.to)) waiting = [...waiting, span];
              }
            }
          }
        }
        const next = (closing: number, pending: number | null): DollarState =>
          closing === value.closing && pending === value.pending && waiting === value.waiting ? value : { closing, pending, waiting };
        const meta = tr.getMeta(key) as DollarMeta | undefined;
        const closed = tr.getMeta(CLOSED_BY_DOLLAR) as number | undefined;
        if (closed) return next(closed, null);
        if (meta && "closing" in meta) return next(meta.closing, value.pending);
        if (meta && "pending" in meta) return next(0, meta.pending);
        if (tr.getMeta("appendedTransaction") || !(tr.docChanged || tr.selectionSet)) return next(value.closing, value.pending);
        return next(0, null);
      },
    },
    /**
     * Dollars an edit leaves in a line that are not typed as "$" (only a typed "$" was looked at: "$ $" with its space
     * then deleted stayed text, and a "$x$" made so stayed text until the file was read again, 2026-11-22): a deletion
     * that leaves the caret between two "$" opens an inline formula there, as typing "$$" does; "$…$" an edit makes
     * becomes a formula once the caret is outside it, as the closing "$" typed would make it.
     */
    appendTransaction(transactions, _oldState, newState) {
      if (inputMode() !== "markdown" || openFormulaId(newState)) return null;
      // (never while an input method composes: the line's text is not touched under it)
      // (nor while the mouse button is held: a selection being dragged; the release converts)
      const live = sourceLiveKey.getState(newState);
      if (transactions.some((tr) => tr.getMeta("composition") !== undefined) || live?.composing || live?.pointer) return null;
      const edits = transactions.filter(userEdit);
      const caret = edits.length ? caretBetweenDollars(newState) : null;
      if (caret !== null && edits.some((tr) => changedRanges(tr).some((r) => r.deleted && r.from === caret))) {
        const tr = newState.tr.delete(caret - 1, caret + 1);
        tr.insert(caret - 1, newState.schema.nodes.math_inline.create({ id: newNodeId(), latex: "" }));
        return openFormula(tr.setMeta(key, { pending: null } satisfies DollarMeta), caret - 1, { kind: "end" }).scrollIntoView();
      }
      const { waiting } = key.getState(newState)!;
      if (!waiting.length) return null;
      const head = newState.selection.head;
      const ready = waiting
        .filter((span) => head < span.from || head > span.to)
        .map((span) => ({ span, latex: stillDollarSpan(newState, span) }))
        .filter((item): item is { span: { from: number; to: number }; latex: string } => item.latex !== null)
        .sort((a, b) => b.span.from - a.span.from);
      if (!ready.length) return null;
      const tr = newState.tr;
      let limit = Infinity;
      for (const { span, latex } of ready) {
        if (span.to > limit) continue;
        tr.replaceWith(span.from, span.to, newState.schema.nodes.math_inline.create({ id: newNodeId(), latex }));
        limit = span.from;
      }
      return tr;
    },
    props: {
      handleTextInput(view, from, to, text) {
        const state = view.state;
        if (openFormulaId(state)) return false;
        const { closing, pending } = key.getState(state)!;
        if (isDollar(text) && closing > 0) {
          view.dispatch(state.tr.setMeta(key, { closing: closing - 1 } satisfies DollarMeta));
          return true;
        }
        const $from = state.doc.resolve(from);
        if (!$from.parent.isTextblock || $from.parent.type.spec.code) return false;
        if ($from.marks().some((m) => m.type.spec.code || m.type.name === "code") || inSourceCode($from)) return false;
        // Markdown input: a "$" right after a "$" typed as a character, "$$": an inline formula in their place.
        if (pending !== null && from === to && from === pending + 1 && isDollar(text)) {
          view.dispatch(state.tr.delete(pending, pending + 1).setMeta(key, { pending: null } satisfies DollarMeta));
          return insertInlineFormula(view.state, view.dispatch, view);
        }
        if (!isDollar(text)) return false;
        if (inputMode() !== "markdown") return insertInlineFormula(state, view.dispatch, view);
        // "$" typed over text selected in a line: that text a formula (Obsidian puts "$" around it)
        if (from !== to && state.doc.resolve(to).sameParent($from)) {
          let objects = false;
          state.doc.nodesBetween(from, to, (node) => {
            if (node.isInline && !node.isText) objects = true;
          });
          if (!objects) return insertInlineFormula(state, view.dispatch, view);
        }
        // "\$" is Markdown's escape (the Markdown input rules make it the character).
        const before = $from.parent.textBetween(Math.max(0, $from.parentOffset - 2), $from.parentOffset, undefined, "￼");
        if (before.endsWith("\\") && !before.endsWith("\\\\")) return false;
        // "$…$": this "$" closes a formula typed after an earlier "$" in the line
        if (from === to) {
          const line = $from.parent.textBetween(0, $from.parentOffset, undefined, "￼");
          const open = Math.max(line.lastIndexOf("$"), line.lastIndexOf("￥"));
          const content = open >= 0 ? line.slice(open + 1) : "";
          const lead = open >= 0 ? line.slice(0, open) : "";
          // (not "$$", an escaped "\$", a "$" inside a code span, or text with an object or a space at an end)
          const opens =
            content !== "" &&
            !/^\s|\s$|￼/.test(content) &&
            !/[$￥]$/.test(lead) &&
            !/(?:^|[^\\])(?:\\\\)*\\$/.test(lead) &&
            (lead.match(/`/g) ?? []).length % 2 === 0;
          if (opens) {
            const node = state.schema.nodes.math_inline.create({ id: newNodeId(), latex: content });
            view.dispatch(
              state.tr
                .replaceWith($from.start() + open, from, node)
                .setMeta(key, { pending: null } satisfies DollarMeta)
                .scrollIntoView(),
            );
            return true;
          }
        }
        view.dispatch(state.tr.insertText(text, from, to).setMeta(key, { pending: from } satisfies DollarMeta));
        return true;
      },
    },
  });
}

// ------------------------------------------------------------------ paste --

export function mathTextToSlice(schema: Schema, blocks: MathTextBlock[]): Slice {
  const nodes: Node[] = blocks.map((block) => {
    if (block.kind === "display") {
      return schema.nodes.math_display.create({ id: newNodeId(), latex: block.latex, numbered: block.numbered });
    }
    const inline = block.parts
      .filter((part) => part.kind === "math" || part.text)
      .map((part) =>
        part.kind === "math"
          ? schema.nodes.math_inline.create({ id: newNodeId(), latex: part.latex })
          : schema.text(part.text),
      );
    return schema.nodes.paragraph.create(null, inline);
  });
  const fragment = Fragment.from(nodes);
  // One paragraph pastes into the current line; more are blocks.
  const openStart = nodes[0]?.type.name === "paragraph" ? 1 : 0;
  const openEnd = nodes.at(-1)?.type.name === "paragraph" ? 1 : 0;
  return new Slice(fragment, openStart, openEnd);
}

/**
 * Whether the HTML that came with pasted text only shows that text as it is (a code editor such as VS Code, a <pre>,
 * a chat message): then the text's Markdown or LaTeX is what was meant, not the HTML's look. HTML that renders it
 * (headings as <h1>, no "#" left in its text) is a formatted copy and pastes as one.
 */
function htmlShowsSource(html: string, text: string, types: readonly string[]): boolean {
  if (!html) return true;
  if (types.includes("vscode-editor-data") || /<pre[\s>]|white-space:\s*pre/i.test(html)) return true;
  const shown = new DOMParser().parseFromString(html, "text/html").body.textContent?.replace(/\s+/g, " ") ?? "";
  const marked = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .find((line) => /^(#{1,6}\s|[-*+]\s|\d+[.)]\s|>\s|```|\|)/.test(line) || /\*\*|__|~~|==|`|\\(section|begin|textbf|item)\b/.test(line));
  return !!marked && shown.includes(marked.replace(/\s+/g, " "));
}

/** A line of Markdown that is a block the document builds as one (a list item, a quote, code, a table, a rule, a
 *  display formula); headings and paragraphs are their text in the source editor. */
const MARKDOWN_BLOCKS = /^[ \t]*(?:[-*+][ \t]|\d+[.)][ \t]|>|```|~~~|\||\$\$|(?:[-*_][ \t]*){3,}$)/m;

/** Whether text is one display formula and nothing else ("$$…$$" or "\[…\]", on one line or several). */
function isDisplayFormulaOnly(text: string): boolean {
  const whole = text.trim();
  const [open, close] = whole.startsWith("$$") ? ["$$", "$$"] : whole.startsWith("\\[") ? ["\\[", "\\]"] : ["", ""];
  if (!open || whole.length <= open.length + close.length || !whole.endsWith(close)) return false;
  const body = whole.slice(open.length, -close.length);
  return !!body.trim() && !body.includes(close) && !body.startsWith("$");
}

/**
 * A paste that leaves an object last in the document (a display formula, a table pasted into the last line): a text line
 * below it for the caret, as leaving a display formula makes one. The object stayed selected, so the next paste or key
 * went over it (of two formulas pasted one after the other one was lost, 2026-11-24).
 */
function withLineAfterObject(tr: Transaction): Transaction {
  const sel = tr.selection;
  if (!(sel instanceof NodeSelection) || !sel.node.isBlock || sel.node.isTextblock) return tr;
  const $end = tr.doc.resolve(sel.to);
  const paragraph = tr.doc.type.schema.nodes.paragraph;
  if ($end.nodeAfter || !$end.parent.canReplaceWith($end.index(), $end.index(), paragraph)) return tr;
  const at = sel.to;
  tr.insert(at, paragraph.create());
  return tr.setSelection(TextSelection.create(tr.doc, at + 1));
}

/** Pasted plain text with math delimiters becomes text with formulas. */
function mathPastePlugin(schema: Schema, inputMode: () => InputMode, latexSource: boolean) {
  return new Plugin({
    key: new PluginKey("visualtex-document-math-paste"),
    // (the editor's own paste: what it reads from the clipboard; the pastes below do the same themselves)
    appendTransaction(transactions, _oldState, state) {
      if (!transactions.some((tr) => tr.getMeta("paste"))) return null;
      const tr = withLineAfterObject(state.tr);
      return tr.docChanged ? tr : null;
    },
    props: {
      handlePaste(view, event) {
        const data = event.clipboardData;
        if (!data || data.types.includes("application/x-visualtex-document")) return false;
        // In code, text is text.
        if (view.state.selection.$from.parent.type.spec.code) return false;
        const text = data.getData("text/plain");
        if (!text) return false;
        const html = data.types.includes("text/html") ? data.getData("text/html") : "";
        const source = htmlShowsSource(html, text, data.types);
        // LaTeX source (\section, itemize, \textbf, equation…) is read into the document's structure, in either input.
        // (in the source editor what is read comes in as source: its formatting written as markers)
        const asSource = (slice: Slice) => (inputMode() === "markdown" ? toSourceSlice(slice) : slice);
        if (latexSource && source && looksLikeLatexSource(text)) {
          const keys = new Map(bibliographyEntries(view.state.doc).map((entry) => [entry.key, entry.id]));
          view.dispatch(withLineAfterObject(view.state.tr.replaceSelection(asSource(markdownToSlice(schema, latexToMarkdown(text), keys)))).scrollIntoView());
          return true;
        }
        // In the source editor pasted text is its lines (sourceMarkdown.ts clipboardTextParser: formulas, tables and
        // pictures as objects; a formatted copy — HTML — comes in through transformPasted): text alone, and text whose
        // HTML only shows it as it is (a chat message, a code editor), which ProseMirror would read from the HTML (a
        // chat message's "$$…$$" stayed text, 2026-11-24). A block of code on a page (<pre>) stays code, unless all it
        // holds is a display formula.
        if (inputMode() === "markdown" && !html) return false;
        if (inputMode() === "markdown" && source) {
          const codeBlock = /<pre[\s>]/i.test(html) && !data.types.includes("vscode-editor-data") && !isDisplayFormulaOnly(text);
          return codeBlock ? false : view.pasteText(text);
        }
        // Pasted Markdown keeps its structure (headings, lists, tables, math, marks): in Markdown input whenever it reads
        // as Markdown, in LaTeX input when it is unmistakably a Markdown document (an AI answer, a README).
        // (the source editor takes Markdown text as it is, unless it has blocks to build: lists, quotes, code, tables)
        const markdown =
          inputMode() === "markdown" ? looksLikeMarkdown(text) && MARKDOWN_BLOCKS.test(text) : looksLikeMarkdownDocument(text);
        if (markdown && source) {
          view.dispatch(withLineAfterObject(view.state.tr.replaceSelection(asSource(markdownToSlice(schema, text)))).scrollIntoView());
          return true;
        }
        if (!hasMathDelimiters(text)) return false;
        const blocks = parseMathText(text);
        if (!blocks.some((b) => b.kind === "display" || b.parts.some((p) => p.kind === "math"))) return false;
        const slice = mathTextToSlice(schema, blocks);
        view.dispatch(withLineAfterObject(view.state.tr.replaceSelection(slice)).scrollIntoView());
        return true;
      },
    },
  });
}

// ------------------------------------------------------------------ links --

/** A click on a link (Ctrl, Shift, the middle button) never takes the window to it: the editor stays. */
function linkClickPlugin() {
  const stay = (_view: EditorView, event: Event) => {
    if ((event.target as Element | null)?.closest?.("a.vtx-link")) event.preventDefault();
    return false;
  };
  return new Plugin({ props: { handleDOMEvents: { click: stay, auxclick: stay } } });
}

// -------------------------------------------------------------- numbering --

export const numberingKey = new PluginKey<DecorationSet>("visualtex-document-numbering");

/** An equation reference: "(n)" of its target; Ctrl+click goes to the equation. */
class EqrefView {
  dom: HTMLElement;
  constructor(
    private node: Node,
    private view: EditorView,
    _getPos: () => number | undefined,
    decorations: readonly Decoration[],
  ) {
    this.dom = document.createElement("span");
    this.dom.className = "vtx-eqref";
    this.dom.contentEditable = "false";
    this.dom.addEventListener("mousedown", this.onMouseDown);
    this.render(decorations);
  }
  private onMouseDown = (event: MouseEvent) => {
    if (!(event.ctrlKey || event.metaKey) || !this.node.attrs.target) return;
    event.preventDefault();
    const target = findNodeById(this.view.state.doc, this.node.attrs.target);
    if (!target) return;
    this.view.dispatch(this.view.state.tr.setSelection(NodeSelection.create(this.view.state.doc, target.pos)).scrollIntoView());
    this.view.focus();
  };
  private render(decorations: readonly Decoration[]) {
    let number: string | null = null;
    for (const d of decorations) {
      const spec = d.spec as { refNumber?: string | null };
      if (spec.refNumber !== undefined) number = spec.refNumber;
    }
    this.dom.textContent = number === null ? "(??)" : number;
    const zh = /^zh/i.test(document.documentElement.lang || navigator.language);
    this.dom.title =
      number === null
        ? zh ? "引用的公式已不存在" : "The referenced equation is gone"
        : zh ? `Ctrl+单击：转到公式 ${number}` : `Ctrl+click: go to equation ${number}`;
    this.dom.classList.toggle("is-broken", number === null);
  }
  update(node: Node, decorations: readonly Decoration[]) {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.render(decorations);
    return true;
  }
  ignoreMutation() {
    return true;
  }
  stopEvent(event: Event) {
    return event.type === "mousedown" && ((event as MouseEvent).ctrlKey || (event as MouseEvent).metaKey);
  }
  destroy() {
    this.dom.removeEventListener("mousedown", this.onMouseDown);
  }
}

/** Insert a reference to the display formula `target` at the selection. */
export function insertEquationReference(target: string): Command {
  return (state, dispatch) => {
    const type = state.schema.nodes.eqref;
    const $from = state.selection.$from;
    if (!$from.parent.inlineContent) return false;
    if (dispatch) dispatch(state.tr.replaceSelectionWith(type.create({ target }), false).scrollIntoView());
    return true;
  };
}

/** The numbered display formulas, in order (for the reference picker). */
export function numberedEquations(doc: Node): Array<{ id: string; number: number; latex: string }> {
  const list: Array<{ id: string; number: number; latex: string }> = [];
  doc.descendants((node) => {
    if (node.type.name === "math_display" && node.attrs.numbered && node.attrs.id && !ownEquationTag(node.attrs.latex)) {
      list.push({ id: node.attrs.id, number: list.length + 1, latex: node.attrs.latex });
    }
    return true;
  });
  return list;
}


/** Display formulas are numbered (1), (2), ... in document order; the node views draw the number. */
export function equationNumbers(doc: Node): Map<string, number> {
  const numbers = new Map<string, number>();
  let n = 0;
  doc.descendants((node) => {
    if (node.type.name === "math_display" && node.attrs.numbered && node.attrs.id && !ownEquationTag(node.attrs.latex)) numbers.set(node.attrs.id, ++n);
    return true;
  });
  return numbers;
}

function numberingPlugin() {
  const build = (doc: Node) => {
    const decorations: Decoration[] = [];
    // the labels in the document's numbering format: (3), (2.3), [2-3] (settings.ts)
    const labels = equationLabels(doc);
    // the sections' numbers, drawn by the style sheet from the attribute (h1[data-number]::before)
    // with its width (digits 0.5em, a period 0.25em: Times' and Latin Modern's bold figures, about) for a centred
    // heading, which centres number and title together as LaTeX does
    for (const [pos, number] of headingNumbers(doc)) {
      const width = number.replace(/[^.]/g, "").length * 0.25 + number.replace(/\./g, "").length * 0.5;
      decorations.push(Decoration.node(pos, pos + doc.nodeAt(pos)!.nodeSize, { "data-number": number, style: `--vtx-number-width:${width}em` }));
    }
    doc.descendants((node, pos) => {
      // a formula with its own \tag draws it itself (it takes no number)
      if (node.type.name === "math_display" && node.attrs.numbered && !ownEquationTag(node.attrs.latex)) {
        const label = node.attrs.id ? labels.get(node.attrs.id) : undefined;
        decorations.push(Decoration.node(pos, pos + node.nodeSize, {}, { equationNumber: label ? bracketLabel(doc, label) : "(?)" }));
      } else if (node.type.name === "eqref") {
        const label = node.attrs.target ? labels.get(node.attrs.target) : undefined;
        decorations.push(Decoration.node(pos, pos + node.nodeSize, {}, { refNumber: label ? bracketLabel(doc, label) : null }));
      }
      return true;
    });
    return DecorationSet.create(doc, decorations);
  };
  return new Plugin<DecorationSet>({
    key: numberingKey,
    state: {
      init: (_config, state) => build(state.doc),
      apply: (tr, value) => (tr.docChanged ? build(tr.doc) : value),
    },
    props: {
      decorations: (state) => numberingKey.getState(state),
      nodeViews: {
        eqref: (node, view, getPos, decorations) => new EqrefView(node, view, getPos, decorations),
      },
    },
  });
}

// ------------------------------------------------------------------- keys --

/** In a list only: one level deeper / higher (elsewhere Tab is a tab character, not an indent). */
const listIndent = (delta: 1 | -1): Command => (state, dispatch, view) =>
  currentListKind(state) ? indent(delta)(state, dispatch, view) : false;

/** Tab: a deeper section, the next table cell, a deeper list item, or a tab character (Mathcha); in the source editor
 *  a section's "#" one more. */
const tabKey = (inputMode: () => InputMode): Command => (state, dispatch, view) =>
  inputMode() === "markdown"
    ? chainCommands(goToCell(1), indentLines)(state, dispatch, view)
    : sectionLevel(state)
      ? changeSectionLevel(1)(state, dispatch, view)
      : chainCommands(goToCell(1), listIndent(1), insertTabCharacter)(state, dispatch, view);

/** Shift+Tab: a higher section, the previous table cell, a higher list item; in the source editor (Obsidian) Tab and
 *  Shift+Tab indent and outdent whole lines. */
const shiftTabKey = (inputMode: () => InputMode): Command => (state, dispatch, view) =>
  inputMode() === "markdown"
    ? chainCommands(goToCell(-1), outdentLines)(state, dispatch, view)
    : sectionLevel(state)
      ? changeSectionLevel(-1)(state, dispatch, view)
      : chainCommands(goToCell(-1), listIndent(-1))(state, dispatch, view);

/** Mathcha's keys (its toolbar tooltips name them); in Markdown input, Enter also ends a line of Markdown that makes a
 * block ("```lang", "---", "| a | b |"). */
function documentKeymap(schema: Schema, inputMode: () => InputMode, rules: () => MarkdownRuleSwitches) {
  const { nodes, marks } = schema;
  return keymap({
    // the Markdown shown in a formatted span is hidden before the history steps (markdownLive.ts)
    "Mod-z": markdownUndo,
    "Mod-y": markdownRedo,
    "Shift-Mod-z": markdownRedo,
    Backspace: chainCommands(clearLineStyle, deleteFormulaBackward, selectAdjacentBlock(-1)),
    Delete: chainCommands(deleteFormulaForward, selectAdjacentBlock(1)),
    ArrowRight: enterFormulaForward,
    ArrowLeft: enterFormulaBackward,
    ArrowDown: enterFormulaDown,
    ArrowUp: enterFormulaUp,
    // (in the source editor each line typed is an undo step of its own, as in Obsidian)
    Enter: lineUndoStep(inputMode, chainCommands(
      openSelectedFormula,
      markdownEnter(inputMode, rules),
      // the source editor's lines (sourceKeys.ts): in code the line's indentation, a list or quote continued
      (state, dispatch, view) => inputMode() === "markdown" && codeLineEnter(inCodeLine)(state, dispatch, view),
      (state, dispatch, view) => inputMode() === "markdown" && sourceEnter(state, dispatch, view),
      (state, dispatch, view) => inputMode() === "markdown" && sourcePlainEnter(state, dispatch, view),
      (state, dispatch, view) => inputMode() === "markdown" && tableEnter(state, dispatch, view),
      splitListItemOrLeave,
      leaveEmptyQuoteLine,
      leaveEmptyAuthorLine,
      // Typora: after a section comes body text (Mathcha keeps the section's style); in the source editor a line
      // split off a heading has no "#" and is body text by itself
      (state, dispatch, view) => inputMode() !== "markdown" && markdownHeadingEnter(inputMode)(state, dispatch, view),
      splitKeepingStyle,
    )),
    "Mod-Enter": tableRowBelow,
    Tab: tabKey(inputMode),
    "Shift-Tab": shiftTabKey(inputMode),
    "Mod-b": toggleDocumentBold,
    "Mod-i": toggleMark(marks.em),
    "Mod-u": toggleMark(marks.underline),
    // the app's link dialog (insert, change or remove a link)
    "Mod-k": requestLink,
    // (in the source editor a heading is its "#")
    "Mod-Alt-1": (state, dispatch, view) =>
      inputMode() === "markdown" ? setSourceHeading(sectionLevel(state) || 1)(state, dispatch, view) : toggleSection(state, dispatch, view),
    // Word's footnote key
    "Mod-Alt-f": insertFootnote(),
    "Shift-Mod-7": toggleList("ordered_list"),
    "Shift-Mod-8": toggleList("bullet_list"),
    "Shift-Mod-l": setAlign("left"),
    "Shift-Mod-e": setAlign("center"),
    "Shift-Mod-r": setAlign("right"),
    "Shift-Mod-j": setAlign("justify"),
    "Shift-Enter": (state, dispatch, view) => {
      // the source editor: a new line under the item's text (Obsidian)
      if (inputMode() === "markdown" && sourceShiftEnter(state, dispatch, view)) return true;
      if (dispatch) dispatch(state.tr.replaceSelectionWith(nodes.hard_break.create()).scrollIntoView());
      return true;
    },
  });
}

/** Whether the caret is in a fenced code block's line (the source editor reads it so). */
const inCodeLine = (state: EditorState) => {
  const facts = blockSyntax(state.selection.$head.parent).facts;
  return facts.codeInside || facts.codeBegin;
};

/** A command whose change starts an undo step of its own in Markdown input (Enter: Obsidian undoes a line at a time). */
function lineUndoStep(inputMode: () => InputMode, command: Command): Command {
  return (state, dispatch, view) => command(state, dispatch && inputMode() === "markdown" ? (tr) => dispatch(closeHistory(tr)) : dispatch, view);
}

export interface DocumentPluginOptions {
  host: FormulaHost;
  /** A formula's LaTeX as plain text leaves the document (copy); default: as stored. */
  formulaText?: (latex: string) => string;
  /** The formula editor's commands, offered when a LaTeX command is typed in the text (\alpha, \frac, ...). */
  mathCommands?: MathCommandProvider | null;
  /** Code block highlighting (the app's: VS Code's grammars and colours). */
  codeHighlighter?: CodeHighlighter | null;
  /** The code block header's words in the app's language. */
  codeLabels?: CodeBlockOptions["labels"];
  /** diagrams (Mermaid) drawn under their code */
  codeDiagram?: CodeBlockOptions["diagram"];
  /** LaTeX or Markdown input (what typing converts, what the candidate menu offers, how Markdown pastes). */
  inputMode?: () => InputMode;
  /** The Markdown rules that are on (all by default). */
  markdownRules?: () => MarkdownRuleSwitches;
  /** Pasted LaTeX source is read into the document's structure (default); false: it pastes as text. */
  latexPaste?: boolean;
  /** The page view: the document divided into the paper's pages (none: one long page). */
  pagination?: PaginationOptions | null;
  /** The image bar's button names in the app's language. */
  imageLabels?: () => ImageLabels;
  /** the image board's bar */
  boardLabels?: () => BoardLabels;
  /** the hint in an empty caption */
  captionLabels?: () => CaptionLabels;
}

const DEFAULT_IMAGE_LABELS: ImageLabels = { left: "Align left", center: "Centre", right: "Align right", wrapLeft: "Wrap text, image left", wrapRight: "Wrap text, image right" };

const DEFAULT_CODE_LABELS: CodeBlockOptions["labels"] = { plainText: "Plain Text", copy: "Copy", copied: "Copied", language: "Language" };

/** Plain text of copied content: text as it is, formulas as $...$ and $$...$$ (what the paste rule reads back). */
export function slicePlainText(slice: Slice, formulaText: (latex: string) => string = (l) => l): string {
  const schema = slice.content.firstChild?.type.schema;
  if (schema) slice = sliceWithoutMarkdownSyntax(slice, schema);
  const blocks: string[] = [];
  const inline = (node: Node) => {
    let out = "";
    node.forEach((child) => {
      if (child.isText) out += child.text;
      else if (child.type.name === "math_inline") out += `$${formulaText(child.attrs.latex)}$`;
      else if (child.type.name === "hard_break") out += "\n";
      else if (child.type.name === "eqref") out += `\\eqref{eq:${child.attrs.target ?? ""}}`;
      else if (child.type.name === "image") out += child.attrs.alt ? `[${child.attrs.alt}]` : "";
      else if (child.type.name === "footnote" || child.type.name === "cite") out += noteText(child, null);
    });
    return out;
  };
  const walk = (node: Node) => {
    if (node.type.name === "math_display") blocks.push(`$$${formulaText(node.attrs.latex)}$$`);
    else if (node.type.name === "page_command") blocks.push(pageCommandLatex(node.attrs.command, node.attrs.value));
    else if (node.isTextblock) blocks.push(inline(node));
    else if (node.type.name === "table_row") {
      const cells: string[] = [];
      node.forEach((cell) => {
        const parts: string[] = [];
        cell.forEach((b) => parts.push(b.type.name === "math_display" ? `$$${formulaText(b.attrs.latex)}$$` : inline(b)));
        cells.push(parts.join(" "));
      });
      blocks.push(cells.join("\t"));
    } else node.forEach(walk);
  };
  // Inline nodes at the top of a slice (part of one paragraph) are one line.
  let run: Node[] = [];
  const flush = () => {
    if (run.length) blocks.push(inline(run[0].type.schema.nodes.paragraph.create(null, run)));
    run = [];
  };
  slice.content.forEach((node) => {
    if (node.isInline) run.push(node);
    else {
      flush();
      walk(node);
    }
  });
  flush();
  return blocks.join("\n\n");
}

export function documentPlugins(schema: Schema, options: DocumentPluginOptions): Plugin[] {
  const inputMode = options.inputMode ?? (() => "latex");
  const allRules = options.markdownRules ?? (() => ALL_RULES);
  const markdown = () => inputMode() === "markdown";
  // Markdown input is the source editor (sourceMarkdown.ts): a heading is its text's "#", never converted, and what is
  // typed stays in the text
  const rules = () => ({ ...allRules(), headings: !markdown(), source: markdown() });
  // Ctrl+B: in an open formula its own bold, in the text the source's "**"
  const sourceBold: Command = (state, dispatch, view) =>
    openFormulaId(state) ? toggleDocumentBold(state, dispatch, view) : toggleSourceFormat("strong")(state, dispatch, view);
  return [
    // first: a composition ProseMirror missed the end of is ended before anything asks whether one is going on
    compositionGuardPlugin(),
    idsPlugin(),
    markdownSourcePlugin(),
    // before the drag and selection gestures: a press on a task's box ticks it
    richBlocksPlugin(),
    documentDragPlugin(),
    // a selection that takes parts of formulas: its deletion, typing over it, its clipboard (before the rules)
    mixedSelectionPlugin(),
    commandMenuPlugin(options.mathCommands ?? null, inputMode, rules),
    // captions (figures.ts): their Enter / Backspace / Delete before the Markdown rules and the keymaps
    figuresPlugin({ labels: options.captionLabels ?? (() => DEFAULT_CAPTION_LABELS) }),
    // before the Markdown rules: what follows a "$" is decided first ("\$" it leaves to them)
    dollarPlugin(inputMode),
    markdownInputPlugin(inputMode, rules),
    // the source editor: each block's Markdown read from its text after every change and drawn (sourceMarkdown.ts)
    sourceLivePlugin({ enabled: markdown }),
    mathPastePlugin(schema, inputMode, options.latexPaste ?? true),
    linkClickPlugin(),
    formulaPlugin(options.host, { undo: markdownUndo, redo: markdownRedo, displayFormula: insertDisplayLatex("") }),
    new Plugin({
      key: new PluginKey("visualtex-document-clipboard-text"),
      props: {
        // a copy leaves out the Markdown shown in a formatted span (its formatting goes with it)
        transformCopied: (slice, view) => sliceWithoutMarkdownSyntax(slice, view.state.schema),
        clipboardTextSerializer: (slice) => slicePlainText(slice, options.formulaText),
      },
    }),
    numberingPlugin(),
    abstractLabelPlugin(),
    notesPlugin(),
    imagesPlugin({ labels: options.imageLabels ?? (() => DEFAULT_IMAGE_LABELS) }),
    imageBoardsPlugin({ labels: options.boardLabels ?? (() => DEFAULT_BOARD_LABELS) }),
    ...(options.pagination ? [paginationPlugin(options.pagination)] : []),
    codeBlockPlugin({ highlighter: options.codeHighlighter ?? null, labels: options.codeLabels ?? DEFAULT_CODE_LABELS, diagram: options.codeDiagram ?? null }),
    keymap(codeBlockKeys),
    markdownKeymap(
      inputMode,
      {
        bold: sourceBold,
        italic: toggleSourceFormat("em"),
        underline: toggleSourceFormat("underline"),
        code: toggleSourceFormat("code"),
        strike: toggleSourceFormat("strike"),
        link: requestLink,
      },
      { heading: setSourceHeading, highlight: toggleSourceFormat("highlight"), link: insertSourceLink, clear: clearSourceFormat },
    ),
    documentKeymap(schema, inputMode, rules),
    keymap(baseKeymap),
    tableColumnsPlugin(),
    tableEditing(),
    gapCursor(),
    dropCursor(),
    history({ newGroupDelay: 600 }),
  ];
}

export interface FootnoteEditorOptions {
  host: FormulaHost;
  mathCommands?: MathCommandProvider | null;
  inputMode?: () => InputMode;
  /** Enter (outside a formula's menu): the note is done */
  submit: Command;
  /** Escape */
  cancel: Command;
}

/** Entries a footnote's text can take: formulas (\alpha, \frac… into a formula; \( … \)) and a line break. */
const inFootnote = (entry: { id: string }) => entry.id.startsWith("math-") || entry.id === "inline-math" || entry.id === "md-inline-math" || entry.id === "newline";

/**
 * A footnote's own editor (the footnote panel): text and inline formulas, typed as in the document — "$$" or "\("
 * opens a formula, a formula command (\frac) starts one, the formula editor's keys move in and out of formulas.
 */
export function footnoteEditorPlugins(schema: Schema, options: FootnoteEditorOptions): Plugin[] {
  const inputMode = options.inputMode ?? (() => "latex");
  return [
    idsPlugin(),
    mixedSelectionPlugin(),
    commandMenuPlugin(options.mathCommands ?? null, inputMode, () => ALL_RULES, inFootnote),
    dollarPlugin(inputMode),
    formulaPlugin(options.host, { undo, redo, displayFormula: () => false }),
    keymap({
      "Mod-z": undo,
      "Mod-y": redo,
      "Shift-Mod-z": redo,
      Backspace: deleteFormulaBackward,
      Delete: deleteFormulaForward,
      ArrowRight: enterFormulaForward,
      ArrowLeft: enterFormulaBackward,
      ArrowDown: enterFormulaDown,
      ArrowUp: enterFormulaUp,
      Enter: chainCommands(openSelectedFormula, options.submit),
      Escape: options.cancel,
      "Shift-Enter": (state, dispatch) => {
        if (dispatch) dispatch(state.tr.replaceSelectionWith(schema.nodes.hard_break.create()).scrollIntoView());
        return true;
      },
    }),
    keymap(baseKeymap),
    history({ newGroupDelay: 600 }),
  ];
}

/** The editor state of a document; nodes without an id (older files, pasted JSON) get one, outside the history. */
export function createDocumentState(schema: Schema, doc: Node, options: DocumentPluginOptions): EditorState {
  let state = EditorState.create({ doc, plugins: documentPlugins(schema, options) });
  // image lines of two or more images are boards (a document from before boards)
  const boards = boardFixes(state);
  if (boards) state = state.apply(boards.setMeta("addToHistory", false));
  // image lines, boards and tables in their figures with a caption slot (a document from before captions)
  const figures = figureFixes(state);
  if (figures) state = state.apply(figures.setMeta("addToHistory", false));
  const tr = assignIds(state.doc, null, () => state.tr);
  return tr ? state.apply(tr.setMeta("addToHistory", false)) : state;
}

export function isDocumentEmpty(state: EditorState): boolean {
  const doc = state.doc;
  return doc.childCount === 1 && doc.firstChild!.type.name === "paragraph" && doc.firstChild!.content.size === 0;
}
