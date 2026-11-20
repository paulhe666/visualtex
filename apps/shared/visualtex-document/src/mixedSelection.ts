// A selection through the text that takes parts of formulas (formulas.ts SelectGesture, DESIGN.md §3.5): what it
// holds, and what Delete, typing, the clipboard and Bold do with it — the text and each formula's part together, as
// one edit, the way the formula rows treat a selection across rows.
import type { Node, Slice } from "prosemirror-model";
import { Plugin, PluginKey, type Command, type EditorState, type Transaction } from "prosemirror-state";
import { toggleMark } from "prosemirror-commands";
import type { EditorView } from "prosemirror-view";
import {
  FORMULA_PARTS,
  findNodeById,
  formulaParts,
  formulaViewOf,
  openFormula,
  openFormulaId,
  typeIntoOpenFormula,
  type FormulaPart,
  type FormulaView,
} from "./formulas.ts";
import { isFormulaNodeName } from "./schema.ts";

interface ResolvedPart {
  part: FormulaPart;
  range: [number, number];
  pos: number;
  node: Node;
  formula: FormulaView;
}

/** The selection's formula parts that hold something, with their formulas. */
function resolvedParts(view: EditorView): ResolvedPart[] {
  return formulaParts(view.state).flatMap((part) => {
    if (!part.range || part.range[0] === part.range[1]) return [];
    const at = findNodeById(view.state.doc, part.id);
    const formula = formulaViewOf(view, part.id);
    return at && formula ? [{ part, range: part.range, pos: at.pos, node: at.node, formula }] : [];
  });
}

/** The selection with its formulas' parts: the slice the clipboard and a drag take, and the range it spans. */
export function selectionSliceWithParts(view: EditorView): { slice: Slice; from: number; to: number } | null {
  if (!formulaParts(view.state).length) return null;
  const parts = resolvedParts(view);
  const sel = view.state.selection;
  const tr = view.state.tr;
  let from = sel.from;
  let to = sel.to;
  for (const p of parts) {
    const latex = p.formula.latexOf(p.range);
    if (latex === null) continue;
    tr.setNodeAttribute(p.pos, "latex", latex);
    from = Math.min(from, p.pos);
    to = Math.max(to, p.pos + p.node.nodeSize);
  }
  return { slice: tr.doc.slice(from, to), from, to };
}

/**
 * The selection's range when no formula is only partly in it (a part from a formula's start to its end is the
 * formula): the text and those formulas whole. Null when a formula is partly selected.
 */
export function wholeFormulasRange(view: EditorView): { from: number; to: number } | null {
  const sel = view.state.selection;
  let from = sel.from;
  let to = sel.to;
  for (const part of formulaParts(view.state)) {
    const at = findNodeById(view.state.doc, part.id);
    const formula = formulaViewOf(view, part.id);
    if (!at || !formula) continue;
    const last = formula.lastOffset();
    if (!part.range || last === null || part.range[0] !== 0 || part.range[1] !== last) return null;
    from = Math.min(from, at.pos);
    to = Math.max(to, at.pos + at.node.nodeSize);
  }
  return from < to ? { from, to } : null;
}

/** The selection and its formulas' parts deleted (a formula left empty goes); null without parts. */
export function deleteSelectionWithParts(view: EditorView): Transaction | null {
  if (!formulaParts(view.state).length) return null;
  const tr = view.state.tr;
  const gone: string[] = [];
  for (const p of resolvedParts(view)) {
    const latex = p.formula.withoutRange(p.range);
    if (latex === null) continue;
    if (!latex.trim()) gone.push(p.part.id);
    else tr.setNodeAttribute(p.pos, "latex", latex);
  }
  tr.deleteSelection();
  for (const id of gone) {
    const at = findNodeById(tr.doc, id);
    if (at) tr.delete(at.pos, at.pos + at.node.nodeSize);
  }
  return tr.setMeta(FORMULA_PARTS, []);
}

/**
 * Delete (Backspace, typing over it) a selection with formula parts. Begun in a formula, the caret goes back into
 * it, at the end of what is left (where the deletion began, as in the formula rows); else where the text was.
 */
function removeWithParts(view: EditorView, text?: string): boolean {
  const sel = view.state.selection;
  const startId = resolvedParts(view).find((p) => p.pos < sel.from || (p.pos === sel.from && p.node.isBlock && sel.empty))?.part.id;
  const tr = deleteSelectionWithParts(view);
  if (!tr) return false;
  const start = startId ? findNodeById(tr.doc, startId) : null;
  if (start && isFormulaNodeName(start.node.type.name)) openFormula(tr, start.pos, { kind: "end" });
  view.dispatch(tr.scrollIntoView());
  if (text) {
    if (openFormulaId(view.state)) typeIntoOpenFormula(view, text);
    else view.dispatch(view.state.tr.insertText(text).scrollIntoView());
  }
  return true;
}

// ---------------------------------------------------------------- bold --

/** `\bm{…}` (or `\boldsymbol{…}`) around the whole formula: its content. */
function boldContent(latex: string): string | null {
  const match = /^\s*\\(?:bm|boldsymbol)\s*\{/.exec(latex);
  if (!match) return null;
  let depth = 0;
  const open = match[0].length - 1;
  for (let i = open; i < latex.length; i++) {
    const ch = latex[i];
    if (ch === "\\") {
      i++;
      continue;
    }
    if (ch === "{") depth++;
    else if (ch === "}" && --depth === 0) return latex.slice(i + 1).trim() ? null : latex.slice(open + 1, i);
  }
  return null;
}

/**
 * Bold (Ctrl+B, the toolbar's B): an open formula's own selection; else the selection's text, the formulas in it
 * and the formulas' parts — every symbol bold (\bm), or none of it when all of it already is. A caret alone sets
 * the bold of what is typed next.
 */
export const toggleDocumentBold: Command = (state, dispatch, view) => {
  const openId = openFormulaId(state);
  if (openId) {
    const formula = view ? formulaViewOf(view, openId) : null;
    return !!formula && (!dispatch || formula.toggleBold());
  }
  const strong = state.schema.marks.strong;
  const sel = state.selection;
  const parts = view ? resolvedParts(view) : [];
  if (sel.empty && !parts.length) return toggleMark(strong)(state, dispatch);
  // what is bold now
  let textBold = true;
  let anyText = false;
  const formulas: { pos: number; node: Node }[] = [];
  state.doc.nodesBetween(sel.from, sel.to, (node, pos) => {
    if (node.isText) {
      anyText = true;
      if (!strong.isInSet(node.marks)) textBold = false;
    } else if (isFormulaNodeName(node.type.name) && pos >= sel.from && pos + node.nodeSize <= sel.to) {
      formulas.push({ pos, node });
      return false;
    }
    return true;
  });
  const allBold =
    (!anyText || textBold) &&
    formulas.every((f) => boldContent(f.node.attrs.latex) !== null) &&
    parts.every((p) => p.formula.rangeBold(p.range) === true);
  const enable = !allBold;
  if (!dispatch) return true;
  const tr = state.tr;
  if (anyText) {
    if (enable) tr.addMark(sel.from, sel.to, strong.create());
    else tr.removeMark(sel.from, sel.to, strong);
  }
  for (const f of formulas) {
    const latex = String(f.node.attrs.latex);
    const inner = boldContent(latex);
    if (enable && inner === null && latex.trim()) tr.setNodeAttribute(f.pos, "latex", `\\bm{${latex}}`);
    if (!enable && inner !== null) tr.setNodeAttribute(f.pos, "latex", inner);
  }
  for (const p of parts) {
    const latex = p.formula.styled(p.range, enable);
    if (latex !== null && latex !== p.node.attrs.latex) tr.setNodeAttribute(p.pos, "latex", latex);
  }
  // the parts stay selected (bold changes no offsets)
  tr.setMeta(FORMULA_PARTS, formulaParts(state));
  dispatch(tr);
  return true;
};

// ---------------------------------------------------------------- plugin --

/** Keys, typing and the clipboard on a selection that takes parts of formulas. */
export function mixedSelectionPlugin(): Plugin {
  return new Plugin({
    key: new PluginKey("visualtex-document-mixed-selection"),
    props: {
      handleKeyDown(view, event) {
        if (!formulaParts(view.state).length || event.isComposing) return false;
        const plain = !event.ctrlKey && !event.metaKey && !event.altKey;
        if ((event.key === "Backspace" || event.key === "Delete") && plain) {
          event.preventDefault();
          return removeWithParts(view);
        }
        if (event.key === "Enter" && plain && !event.shiftKey) {
          // the selection goes; the line is not split (the caret may be in a formula)
          event.preventDefault();
          return removeWithParts(view);
        }
        return false;
      },
      handleTextInput(view, _from, _to, text) {
        if (!formulaParts(view.state).length) return false;
        return removeWithParts(view, text);
      },
      handleDOMEvents: {
        copy(view, event) {
          const mixed = selectionSliceWithParts(view);
          if (!mixed || !event.clipboardData) return false;
          const { dom, text } = view.serializeForClipboard(mixed.slice);
          event.clipboardData.clearData();
          event.clipboardData.setData("text/html", dom.innerHTML);
          event.clipboardData.setData("text/plain", text);
          event.preventDefault();
          return true;
        },
        cut(view, event) {
          const mixed = selectionSliceWithParts(view);
          if (!mixed || !event.clipboardData) return false;
          const { dom, text } = view.serializeForClipboard(mixed.slice);
          event.clipboardData.clearData();
          event.clipboardData.setData("text/html", dom.innerHTML);
          event.clipboardData.setData("text/plain", text);
          event.preventDefault();
          removeWithParts(view);
          return true;
        },
      },
    },
  });
}

export function selectionHasFormulaParts(state: EditorState): boolean {
  return formulaParts(state).length > 0;
}
