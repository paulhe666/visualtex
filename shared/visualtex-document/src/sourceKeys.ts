// The source editor's keys for lines, as Obsidian's (read from Obsidian 1.15, ~/claude/obsidian-study/FINDINGS.md
// "Editing keys"): Enter continues a list item ("- ", "1. " counted on, "- [ ] ") or a quote ("> ") and ends an empty
// one; Shift+Enter a new line indented under the item's text; Tab / Shift+Tab indent / outdent whole lines; and the
// toolbar's list, quote, task and code block buttons as edits of the lines' text.

import { TextSelection, type Command, type EditorState, type Transaction } from "prosemirror-state";
import type { Node } from "prosemirror-model";

/** A line's prefix: quotes and indentation, a bullet or number, a task's box (Obsidian's pattern). */
export const LINE_PREFIX = /^([>\s]*)(([*+-] |(\d+)([.)] ))(?:\[(.)\] )?)?/;

const OBJECT_CHAR = "";

interface CaretLine {
  node: Node;
  /** position of the line's text start */
  start: number;
  text: string;
  /** the caret's offset in the line */
  ch: number;
  /** the line's index among the document's children, and the document */
  index: number;
}

/** The line (a paragraph or heading at the document's top) the caret is in, or null. */
function caretLine(state: EditorState): CaretLine | null {
  const { $head, empty } = state.selection;
  if (!empty || $head.depth !== 1 || !$head.parent.isTextblock || $head.parent.type.spec.code) return null;
  const node = $head.parent;
  if (node.type.name !== "paragraph" && node.type.name !== "heading") return null;
  return { node, start: $head.start(), text: node.textBetween(0, node.content.size, undefined, OBJECT_CHAR), ch: $head.parentOffset, index: $head.index(0) };
}

function lineText(doc: Node, index: number): string | null {
  if (index < 0 || index >= doc.childCount) return null;
  const node = doc.child(index);
  if (node.type.name !== "paragraph" && node.type.name !== "heading") return null;
  return node.textBetween(0, node.content.size, undefined, OBJECT_CHAR);
}

/** A new line at the caret holding `prefix` (Obsidian's "\n" + prefix), `skip` characters after the caret dropped. */
function newLine(state: EditorState, line: CaretLine, prefix: string, skip = 0): Transaction {
  const at = line.start + line.ch;
  const tr = state.tr;
  if (skip) tr.delete(at, at + skip);
  tr.split(at, 1, [{ type: state.schema.nodes.paragraph }]);
  const lineStart = at + 2;
  if (prefix) tr.insertText(prefix, lineStart);
  return tr.setSelection(TextSelection.create(tr.doc, lineStart + prefix.length)).scrollIntoView();
}

/** Lines whose ranges `edits` change (offsets in each line), applied from the last. */
function editLines(state: EditorState, edits: Array<{ index: number; from: number; to: number; text: string }>): Transaction {
  const tr = state.tr;
  const starts: number[] = [];
  state.doc.forEach((_node, offset) => starts.push(offset));
  for (const edit of [...edits].sort((a, b) => b.index - a.index || b.from - a.from)) {
    const start = starts[edit.index] + 1;
    tr.insertText(edit.text, start + edit.from, start + edit.to);
  }
  return tr;
}

/**
 * Enter (Obsidian's newlineAndIndentContinueMarkdownList): in a list item or quote, a new line with the same prefix
 * (the number counted on, a task's box empty); an empty item: its marker taken off (indented: one level out); in a
 * code block a new line with the line's indentation. Elsewhere: false (the editor's own Enter).
 */
export const sourceEnter: Command = (state, dispatch) => {
  const line = caretLine(state);
  if (!line) return false;
  const doc = state.doc;
  const s = LINE_PREFIX.exec(line.text);
  if (!s || !s[0]) return false;
  let [l, c, d, h, u, p, f] = s;
  const m = l.length;
  if (line.ch < l.length || (!d && !c)) return false;
  if (!d) {
    // a line indented under an item: that item's marker (an item in as many quotes: Obsidian takes any item whose
    // prefix is as long, so a "> " line under "- a" went on as "- ")
    const quotes = (prefix: string) => prefix.replace(/[^>]/g, "").length;
    for (let v = line.index - 1; v >= 0; v--) {
      const t = lineText(doc, v);
      if (t === null) break;
      const k = LINE_PREFIX.exec(t)!;
      if (k[2] && k[0].length === m && quotes(k[1]) === quotes(l)) {
        [l, c, d, h, u, p, f] = k;
        break;
      }
      if (k[0].length < m) break;
    }
  }
  if (!dispatch) return true;
  if (line.text.slice(l.length).trim() === "") {
    const edit = (from: number, to: number, text = "", index = line.index) => ({ index, from, to, text });
    if (!c) {
      dispatch(editLines(state, [edit(0, line.ch)]).scrollIntoView());
      return true;
    }
    if (c.endsWith(">") || c.endsWith("> ")) {
      if (d) {
        dispatch(editLines(state, [edit(c.length, line.ch)]).scrollIntoView());
        return true;
      }
      const w = c.endsWith(">") ? 1 : 2;
      if (line.index > 0 && lineText(doc, line.index - 1) === c) {
        dispatch(editLines(state, [edit(c.length - w, c.length, "", line.index - 1), edit(c.length - w, c.length)]).scrollIntoView());
        return true;
      }
      // "\n" + the quote one level less, in place of the last ">"
      const tr = editLines(state, [edit(c.length - w, c.length)]);
      const at = tr.mapping.map(line.start + c.length - w);
      tr.split(at, 1, [{ type: state.schema.nodes.paragraph }]);
      const outer = c.slice(0, c.length - w);
      if (outer) tr.insertText(outer, at + 2);
      dispatch(tr.setSelection(TextSelection.create(tr.doc, at + 2 + outer.length)).scrollIntoView());
      return true;
    }
    if (c.endsWith("\t")) {
      dispatch(editLines(state, [edit(c.length - 1, c.length)]).scrollIntoView());
      return true;
    }
    let v = 0;
    while (v < 4 && c.charAt(c.length - v - 1) === " ") v++;
    dispatch(editLines(state, [v > 0 ? edit(c.length - v, c.length) : edit(0, c.length)]).scrollIntoView());
    return true;
  }
  if (!d) {
    dispatch(newLine(state, line, c));
    return true;
  }
  let g = h;
  if (u) g = `${parseInt(u, 10) + 1}${p}`;
  if (f !== undefined) g += "[ ] ";
  const after = LINE_PREFIX.exec(line.text.slice(line.ch));
  if (after && after[3]) g = "";
  dispatch(newLine(state, line, c + g, after && after[1] ? after[1].length : 0));
  return true;
};

/** Shift+Enter (Obsidian's newlineAndIndentOnly): a new line indented under the item's text, its quotes kept. */
export const sourceShiftEnter: Command = (state, dispatch) => {
  const line = caretLine(state);
  if (!line) return false;
  const s = LINE_PREFIX.exec(line.text);
  if (!s || !s[0] || line.ch < s[0].length) {
    if (dispatch) dispatch(newLine(state, line, ""));
    return true;
  }
  const [, c, d] = s;
  if (dispatch) dispatch(newLine(state, line, c + (d ? d.replace(/\t/g, "    ").replace(/./g, " ") : "")));
  return true;
};

/** Enter in a fenced code block's line: a new line with the line's indentation (Obsidian: newlineAndIndentOnly). */
export function codeLineEnter(inCode: (state: EditorState) => boolean): Command {
  return (state, dispatch) => {
    const line = caretLine(state);
    if (!line || !inCode(state)) return false;
    const indent = /^[ \t]*/.exec(line.text)?.[0] ?? "";
    if (dispatch) dispatch(newLine(state, line, line.ch >= indent.length ? indent : ""));
    return true;
  };
}

/** The top-level lines the selection takes (their indices). */
function selectedLines(state: EditorState): number[] {
  const { $from, $to } = state.selection;
  if ($from.depth < 1 || $to.depth < 1) return [];
  const out: number[] = [];
  for (let i = $from.index(0); i <= $to.index(0); i++) {
    const node = state.doc.child(i);
    if (node.type.name === "paragraph" || node.type.name === "heading") out.push(i);
  }
  return out;
}

/** Tab (Obsidian's indentMore): a tab before each line the selection takes, wherever the caret is in it. */
export const indentLines: Command = (state, dispatch) => {
  const lines = selectedLines(state);
  if (!lines.length) return false;
  if (dispatch) dispatch(editLines(state, lines.map((index) => ({ index, from: 0, to: 0, text: "\t" }))).scrollIntoView());
  return true;
};

/** Shift+Tab (indentLess): a tab (or up to four spaces) taken from the start of each line the selection takes. */
export const outdentLines: Command = (state, dispatch) => {
  const lines = selectedLines(state);
  if (!lines.length) return false;
  const edits: Array<{ index: number; from: number; to: number; text: string }> = [];
  for (const index of lines) {
    const text = lineText(state.doc, index) ?? "";
    if (text.startsWith("\t")) edits.push({ index, from: 0, to: 1, text: "" });
    else {
      const spaces = /^ {1,4}/.exec(text)?.[0].length ?? 0;
      if (spaces) edits.push({ index, from: 0, to: spaces, text: "" });
    }
  }
  if (dispatch && edits.length) dispatch(editLines(state, edits).scrollIntoView());
  return true;
};

/**
 * A list on the selected lines (the toolbar's list buttons, Ctrl+Shift+[ / ]): each line given the marker ("- ", or
 * "1. " counted on), another marker replaced; lines that all have it already lose it.
 */
export function toggleLineList(kind: "bullet" | "ordered"): Command {
  return (state, dispatch) => {
    const lines = selectedLines(state);
    if (!lines.length) return false;
    const parsed = lines.map((index) => ({ index, text: lineText(state.doc, index) ?? "" })).map((l) => ({ ...l, m: LINE_PREFIX.exec(l.text)! }));
    const isKind = (m: RegExpExecArray) => (kind === "bullet" ? !!m[3] && !m[4] : !!m[4]);
    const all = parsed.every((l) => l.m[3] && isKind(l.m));
    if (!dispatch) return true;
    let n = 1;
    const edits = parsed.map(({ index, m }) => {
      const from = m[1].length;
      const to = from + (m[3]?.length ?? 0);
      return { index, from, to, text: all ? "" : kind === "bullet" ? "- " : `${n++}. ` };
    });
    dispatch(editLines(state, edits).scrollIntoView());
    return true;
  };
}

/** A quote on the selected lines ("> " before each), or taken off when they all are quoted. */
export const toggleLineQuote: Command = (state, dispatch) => {
  const lines = selectedLines(state);
  if (!lines.length) return false;
  const texts = lines.map((index) => ({ index, text: lineText(state.doc, index) ?? "" }));
  const all = texts.every((t) => /^\s*>/.test(t.text));
  if (!dispatch) return true;
  dispatch(
    editLines(
      state,
      texts.map(({ index, text }) => {
        if (!all) return { index, from: 0, to: 0, text: "> " };
        const m = /^(\s*)> ?/.exec(text)!;
        return { index, from: m[1].length, to: m[0].length, text: "" };
      }),
    ).scrollIntoView(),
  );
  return true;
};

/** Ctrl+L (Obsidian's checklist key): a line a task "- [ ] ", a task ticked, a ticked one unticked. */
export const cycleLineTask: Command = (state, dispatch) => {
  const lines = selectedLines(state);
  if (!lines.length) return false;
  if (!dispatch) return true;
  const edits = lines.map((index) => {
    const text = lineText(state.doc, index) ?? "";
    const m = LINE_PREFIX.exec(text)!;
    const at = m[1].length;
    if (m[6] !== undefined) {
      const box = at + m[3].length + 1;
      return { index, from: box, to: box + 1, text: m[6] === " " ? "x" : " " };
    }
    if (m[3]) return { index, from: at + m[3].length, to: at + m[3].length, text: "[ ] " };
    return { index, from: at, to: at, text: "- [ ] " };
  });
  dispatch(editLines(state, edits).scrollIntoView());
  return true;
};

/** The toolbar's code block: "```" lines around the selected lines (an empty block after the caret's line). */
export const insertCodeFence: Command = (state, dispatch) => {
  const lines = selectedLines(state);
  if (!lines.length) return false;
  if (!dispatch) return true;
  const { paragraph } = state.schema.nodes;
  const first = lines[0];
  const last = lines[lines.length - 1];
  let start = 0;
  state.doc.forEach((_n, offset, index) => {
    if (index === first) start = offset;
  });
  let end = 0;
  state.doc.forEach((n, offset, index) => {
    if (index === last) end = offset + n.nodeSize;
  });
  const empty = !state.selection.empty ? false : !(lineText(state.doc, first) ?? "").trim();
  const tr = state.tr;
  if (empty) {
    // the empty line becomes the fence, an empty line in it and the closing fence after
    tr.insertText("```", start + 1);
    const after = tr.mapping.map(end);
    tr.insert(after, [paragraph.create(), paragraph.create(null, state.schema.text("```"))]);
    tr.setSelection(TextSelection.create(tr.doc, start + 4));
  } else {
    tr.insert(end, paragraph.create(null, state.schema.text("```")));
    tr.insert(start, paragraph.create(null, state.schema.text("```")));
    tr.setSelection(TextSelection.create(tr.doc, start + 4));
  }
  dispatch(tr.scrollIntoView());
  return true;
};

/**
 * "```" typed at a line's start (the line now only that): the closing fence on a line after it, the caret staying after
 * the opening one, where the language goes (Obsidian).
 */
export function closeFenceAfterTyping(state: EditorState): Transaction | null {
  const line = caretLine(state);
  if (!line) return null;
  if (!/^\s*```$/.test(line.text) || line.ch !== line.text.length) return null;
  // not when it closes an open block (the lines above hold an opening fence with no closing one)
  let open = false;
  for (let v = 0; v < line.index; v++) {
    const t = lineText(state.doc, v);
    if (t !== null && /^\s*```/.test(t)) open = !open;
  }
  if (open) return null;
  const tr = state.tr;
  const end = line.start + line.node.content.size + 1;
  const indent = /^\s*/.exec(line.text)![0];
  tr.insert(end, state.schema.nodes.paragraph.create(null, state.schema.text(indent + "```")));
  return tr.setSelection(TextSelection.create(tr.doc, line.start + line.ch));
}

/** The list the caret's line is an item of, by its marker (the toolbar's pressed list button). */
export function lineListKind(state: EditorState): "bullet" | "ordered" | null {
  const { $head } = state.selection;
  if (!$head.parent.isTextblock) return null;
  const m = LINE_PREFIX.exec($head.parent.textBetween(0, $head.parent.content.size, undefined, OBJECT_CHAR));
  return m?.[4] ? "ordered" : m?.[3] ? "bullet" : null;
}

/** A line prefix's indentation (columns, a tab four) and quote depth. */
function levelOf(quoteIndent: string): { indent: number; quotes: number } {
  return { indent: quoteIndent.replace(/>/g, "").replace(/\t/g, "    ").length, quotes: (quoteIndent.match(/>/g) ?? []).length };
}

/**
 * Ordered lists numbered again after a change (Obsidian's renumbering with smartIndentList): an item changed or after a
 * change continues the item before it at its level (a new level starts at 1, a number typed on a new item stays), and
 * the items after it at that level follow on.
 */
export function renumberLists(old: EditorState, state: EditorState, changed: number[]): Transaction | null {
  const doc = state.doc;
  const count = doc.childCount;
  const textAt = (i: number) => lineText(doc, i);
  const starts: number[] = [];
  doc.forEach((_n, offset) => starts.push(offset));
  const edits: Array<{ index: number; from: number; to: number; text: string }> = [];
  const done = new Set<number>();
  const oldText = (i: number) => {
    // the same line before the change (by position)
    const pos = starts[i];
    if (pos >= old.doc.content.size) return null;
    const $pos = old.doc.resolve(Math.min(pos + 1, old.doc.content.size));
    if ($pos.depth < 1) return null;
    const node = $pos.node(1);
    return node.isTextblock ? node.textBetween(0, node.content.size, undefined, OBJECT_CHAR) : null;
  };
  const previousNumber = (i: number, indent: number, quotes: number): number | null => {
    for (let v = i - 1; v >= 0; v--) {
      const t = textAt(v);
      if (t === null) return null;
      if (!t.trim()) continue;
      const m = LINE_PREFIX.exec(t)!;
      const level = levelOf(m[1]);
      if (level.quotes !== quotes) return null;
      // (a line of an item's text, not an item, is passed over)
      if (!m[2]) continue;
      if (level.indent > indent) continue;
      if (level.indent < indent) return null;
      return m[4] ? parseInt(m[4], 10) : null;
    }
    return null;
  };
  const visit = (i: number) => {
    if (done.has(i)) return;
    done.add(i);
    const t = textAt(i);
    if (t === null) return;
    const m = LINE_PREFIX.exec(t);
    if (!m || !m[4]) return;
    const { indent, quotes } = levelOf(m[1]);
    let n = previousNumber(i, indent, quotes);
    if (n === null) {
      // the first item at its level: as typed when the marker is new or its level changed; else 1
      const before = oldText(i);
      const was = before !== null ? LINE_PREFIX.exec(before) : null;
      if (!was || was[2] !== m[2]) n = parseInt(m[4], 10);
      else n = levelOf(was[1]).indent === indent ? parseInt(m[4], 10) : 1;
    } else n++;
    if (String(n) !== m[4]) edits.push({ index: i, from: m[1].length, to: m[1].length + m[3].length, text: `${n}${m[5]}` });
    let blank = false;
    for (let w = i + 1; w < count && !done.has(w); w++) {
      const next = textAt(w);
      if (next === null) break;
      if (!next.trim()) {
        blank = true;
        continue;
      }
      const k = LINE_PREFIX.exec(next)!;
      if (!k[0]) {
        if (blank) break;
        continue;
      }
      const level = levelOf(k[1]);
      if (level.indent === indent && level.quotes === quotes) {
        if (!k[4]) break;
        n++;
        done.add(w);
        if (String(n) === k[4]) break;
        edits.push({ index: w, from: k[1].length, to: k[1].length + k[3].length, text: `${n}${k[5]}` });
      }
      if (level.indent < indent) break;
    }
  };
  for (const i of [...new Set(changed.flatMap((i) => [i, i + 1]))].sort((a, b) => a - b)) if (i < count) visit(i);
  return edits.length ? editLines(state, edits) : null;
}

/** Home: to the start of an item's text (after its indentation, quote, marker and box), from there to the line's start. */
export const homeInLine: Command = (state, dispatch) => {
  const { $head, empty } = state.selection;
  if ($head.depth !== 1 || !$head.parent.isTextblock) return false;
  const text = $head.parent.textBetween(0, $head.parent.content.size, undefined, OBJECT_CHAR);
  const m = LINE_PREFIX.exec(text);
  if (!m || !m[0] || !m[2] && !/>/.test(m[1])) return false;
  const prefixEnd = m[0].length;
  if ($head.parentOffset <= prefixEnd) return false;
  if (dispatch) {
    const pos = $head.start() + prefixEnd;
    dispatch(state.tr.setSelection(empty ? TextSelection.create(state.doc, pos) : TextSelection.create(state.doc, state.selection.anchor, pos)).scrollIntoView());
  }
  return true;
};

/** Enter elsewhere (CodeMirror's insertNewlineAndIndent, as Obsidian has it): a new line, the spaces after the caret
 *  dropped. */
export const sourcePlainEnter: Command = (state, dispatch) => {
  const line = caretLine(state);
  if (!line) return false;
  const skip = /^[ \t]*/.exec(line.text.slice(line.ch))?.[0].length ?? 0;
  if (dispatch) dispatch(newLine(state, line, "", skip));
  return true;
};
