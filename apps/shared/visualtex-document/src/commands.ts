// Document commands: entering formulas from the text (DESIGN.md §3) and
// creating them. The rules for leaving a formula are the editing core's
// (horizontalExit/verticalExit) and the formula session's.
import { NodeSelection, Selection, TextSelection, type Command, type EditorState, type Transaction } from "prosemirror-state";
import type { Node } from "prosemirror-model";
import { formulaKey, openFormula, openFormulaId } from "./formulas.ts";
import { isFormulaNodeName } from "./schema.ts";
import { newNodeId } from "./ids.ts";

const CJK = /[⺀-鿿豈-﫿＀-￯　-〿]/;

/** Text selected in the document, as formula LaTeX (CJK text stays text inside \text{}). */
export function textToFormulaLatex(text: string): string {
  const trimmed = text.replace(/\s+/g, " ").trim();
  if (!trimmed) return "";
  if (!CJK.test(trimmed)) return trimmed;
  return trimmed.replace(/[⺀-鿿豈-﫿＀-￯　-〿]+/g, (run) => `\\text{${run}}`);
}

function isFormula(node: Node | null | undefined): node is Node {
  return !!node && isFormulaNodeName(node.type.name);
}

/** Ctrl+M: an inline formula at the caret (the selected text becomes its LaTeX), opened. */
export const insertInlineFormula: Command = (state, dispatch) => {
  const type = state.schema.nodes.math_inline;
  const { $from, from, to, empty } = state.selection;
  if (!$from.parent.inlineContent || !$from.parent.type.contentMatch.matchType(type)) {
    if (!$from.parent.isTextblock) return false;
  }
  if (state.selection instanceof NodeSelection && isFormula(state.selection.node)) {
    // A selected formula opens instead.
    if (dispatch) dispatch(openFormula(state.tr, from, { kind: "end" }));
    return true;
  }
  const latex = empty ? "" : textToFormulaLatex(state.doc.textBetween(from, to, " "));
  if (dispatch) {
    const node = type.create({ id: newNodeId(), latex });
    const tr = state.tr.replaceSelectionWith(node, false);
    const pos = tr.mapping.map(from, -1);
    dispatch(openFormula(tr, pos, { kind: "end" }).scrollIntoView());
  }
  return true;
};

/** A display formula (an empty paragraph becomes one; otherwise one after the block), opened. */
export const insertDisplayFormula: Command = (state, dispatch) => insertDisplayLatex("")(state, dispatch);

/** A display formula holding `initial` (\begin{align}, \begin{gather} start with their environment), opened. */
export function insertDisplayLatex(initial: string, numbered = false): Command {
  return (state, dispatch) => displayFormula(state, dispatch, initial, numbered);
}

/** A closed inline formula holding `latex` at the caret; the text caret goes on after it. */
export function insertInlineLatex(latex: string): Command {
  return (state, dispatch) => {
    const { $from } = state.selection;
    if (!$from.parent.inlineContent) return false;
    if (dispatch) {
      const tr = state.tr.replaceSelectionWith(state.schema.nodes.math_inline.create({ id: newNodeId(), latex }), false);
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

function displayFormula(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  initial: string,
  numbered = false,
): boolean {
  const type = state.schema.nodes.math_display;
  const sel = state.selection;
  if (sel instanceof NodeSelection && sel.node.type.name === "math_inline") {
    // An inline formula becomes a display formula of its own.
    if (dispatch) {
      const latex = sel.node.attrs.latex;
      const tr = state.tr.delete(sel.from, sel.to);
      const $pos = tr.doc.resolve(tr.mapping.map(sel.from));
      const node = type.create({ id: newNodeId(), latex });
      const blockEnd = $pos.after($pos.depth);
      tr.insert(blockEnd, node);
      dispatch(openFormula(tr, blockEnd, { kind: "end" }).scrollIntoView());
    }
    return true;
  }
  const { $from, from, to, empty } = sel;
  if (!$from.parent.isTextblock) return false;
  const latex = initial || (empty ? "" : textToFormulaLatex(state.doc.textBetween(from, to, " ")));
  if (!dispatch) return true;
  const node = type.create({ id: newNodeId(), latex, numbered });
  const tr = state.tr;
  if (!empty) tr.deleteSelection();
  const $at = tr.doc.resolve(tr.mapping.map(from));
  const parent = $at.parent;
  let pos: number;
  if (parent.type.name === "paragraph" && parent.content.size === 0) {
    pos = $at.before();
    tr.replaceWith(pos, $at.after(), node);
  } else if ($at.parentOffset === 0) {
    pos = $at.before();
    tr.insert(pos, node);
  } else if ($at.parentOffset === parent.content.size) {
    pos = $at.after();
    tr.insert(pos, node);
  } else {
    // In the middle of a paragraph: split it and put the formula between the halves.
    tr.split($at.pos);
    pos = tr.mapping.map($at.pos) - 1;
    tr.insert(pos, node);
  }
  // As in Mathcha, a line always follows a display formula, so the text can go on below it.
  const after = pos + node.nodeSize;
  if (!tr.doc.resolve(after).nodeAfter?.isTextblock) tr.insert(after, state.schema.nodes.paragraph.create());
  dispatch(openFormula(tr, pos, { kind: initial ? "start" : "end" }).scrollIntoView());
  return true;
}

/** Right/Left (Down/Up) from the text into an adjacent formula. */
function enterAdjacent(direction: -1 | 1, vertical: boolean): Command {
  return (state, dispatch, view) => {
    if (openFormulaId(state)) return false;
    const sel = state.selection;
    if (sel instanceof NodeSelection && isFormula(sel.node)) {
      if (vertical) return false;
      // A selected (not open) formula: the arrow puts the text caret beside it, as for any selected object.
      if (dispatch) {
        const $side = state.doc.resolve(direction > 0 ? sel.to : sel.from);
        dispatch(state.tr.setSelection(Selection.near($side, direction)).scrollIntoView());
      }
      return true;
    }
    if (!(sel instanceof TextSelection) || !sel.empty) return false;
    const $c = sel.$cursor;
    if (!$c) return false;
    if (!vertical) {
      const inline = direction > 0 ? $c.nodeAfter : $c.nodeBefore;
      if (isFormula(inline)) {
        const pos = direction > 0 ? $c.pos : $c.pos - inline.nodeSize;
        if (dispatch) dispatch(openFormula(state.tr, pos, { kind: direction > 0 ? "start" : "end" }));
        return true;
      }
    }
    // At the end (start) of a textblock with a display formula right after (before) it.
    const atEdge = direction > 0 ? $c.parentOffset === $c.parent.content.size : $c.parentOffset === 0;
    if (!atEdge) return false;
    if (vertical && view && !view.endOfTextblock(direction > 0 ? "down" : "up")) return false;
    const blockPos = direction > 0 ? $c.after() : $c.before();
    const $block = state.doc.resolve(blockPos);
    const neighbour = direction > 0 ? $block.nodeAfter : $block.nodeBefore;
    if (neighbour?.type.name !== "math_display") return false;
    const pos = direction > 0 ? blockPos : blockPos - neighbour.nodeSize;
    if (dispatch) dispatch(openFormula(state.tr, pos, { kind: direction > 0 ? "start" : "end" }));
    return true;
  };
}

export const enterFormulaForward = enterAdjacent(1, false);
export const enterFormulaBackward = enterAdjacent(-1, false);
export const enterFormulaDown = enterAdjacent(1, true);
export const enterFormulaUp = enterAdjacent(-1, true);

/** Backspace right after a formula / Delete right before one: the whole formula goes (Mathcha). */
function deleteAdjacentFormula(direction: -1 | 1): Command {
  return (state, dispatch) => {
    const sel = state.selection;
    if (!(sel instanceof TextSelection) || !sel.empty || !sel.$cursor) return false;
    const $c = sel.$cursor;
    let node = direction < 0 ? $c.nodeBefore : $c.nodeAfter;
    let pos = direction < 0 ? $c.pos - (node?.nodeSize ?? 0) : $c.pos;
    if (!isFormula(node)) {
      // At the start (end) of a textblock right after (before) a display formula: the formula is selected first, a
      // second press deletes it (a whole display formula does not vanish with one key; Typora's users lose work so).
      const atEdge = direction < 0 ? $c.parentOffset === 0 : $c.parentOffset === $c.parent.content.size;
      if (!atEdge || $c.depth < 1) return false;
      const $block = state.doc.resolve(direction < 0 ? $c.before() : $c.after());
      node = direction < 0 ? $block.nodeBefore : $block.nodeAfter;
      if (!isFormula(node)) return false;
      pos = direction < 0 ? $block.pos - node.nodeSize : $block.pos;
      // An empty line next to it goes instead, and the caret goes into the formula (its end, or its start for
      // Delete): the line was what the key was pressed on, the formula is kept.
      if ($c.parent.type === state.schema.nodes.paragraph && $c.parent.content.size === 0) {
        if (dispatch) {
          const tr = state.tr.delete($c.before(), $c.after());
          const at = direction < 0 ? pos : $c.before();
          dispatch(openFormula(tr, at, { kind: direction < 0 ? "end" : "start" }).scrollIntoView());
        }
        return true;
      }
      if (dispatch) dispatch(state.tr.setSelection(NodeSelection.create(state.doc, pos)).scrollIntoView());
      return true;
    }
    if (dispatch) dispatch(state.tr.delete(pos, pos + node!.nodeSize).scrollIntoView());
    return true;
  };
}

export const deleteFormulaBackward = deleteAdjacentFormula(-1);
export const deleteFormulaForward = deleteAdjacentFormula(1);

/** Enter on a selected formula opens it. */
export const openSelectedFormula: Command = (state, dispatch) => {
  const sel = state.selection;
  if (!(sel instanceof NodeSelection) || !isFormula(sel.node) || openFormulaId(state)) return false;
  if (dispatch) dispatch(openFormula(state.tr, sel.from, { kind: "end" }));
  return true;
};

/** Toggle the selected (or open) display formula's equation number. */
export const toggleEquationNumber: Command = (state, dispatch) => {
  const sel = state.selection;
  if (!(sel instanceof NodeSelection) || sel.node.type.name !== "math_display") return false;
  if (dispatch) {
    const tr = state.tr.setNodeAttribute(sel.from, "numbered", !sel.node.attrs.numbered);
    const open = formulaKey.getState(state)?.openId;
    if (open) tr.setMeta(formulaKey, { open, entry: { kind: "end" }, keep: true });
    dispatch(tr);
  }
  return true;
};

/** The formula the caret is on (selected or open), if any. */
export function selectedFormula(state: EditorState): { node: Node; pos: number } | null {
  const sel = state.selection;
  return sel instanceof NodeSelection && isFormula(sel.node) ? { node: sel.node, pos: sel.from } : null;
}
