// The title block and the abstract (LaTeX's \title, \author, \date + \maketitle, \begin{abstract}): the lines under
// the title and the abstract's commands, and the abstract's heading in the document's language ("摘要" when the
// document has Chinese, as ctexart prints it — the export's class follows the same text).
import type { Node, NodeType } from "prosemirror-model";
import { Plugin, PluginKey, TextSelection, type Command } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import { setBlockType, splitBlockAs, wrapIn } from "prosemirror-commands";

/** Chinese, Japanese and full-width text: the export writes ctexart for it (latex.ts). */
export const CJK_TEXT = /[⺀-鿿豈-﫿＀-￯　-〿]/;

/** Whether the document has text the export writes ctexart for (text, formulas, footnotes). */
export function documentHasCjk(doc: Node): boolean {
  let found = false;
  doc.descendants((node) => {
    if (found) return false;
    if (node.isText) found = CJK_TEXT.test(node.text ?? "");
    else if (typeof node.attrs.latex === "string") found = CJK_TEXT.test(node.attrs.latex);
    else if (node.type.name === "footnote") found = CJK_TEXT.test(String(node.attrs.text ?? ""));
    return !found;
  });
  return found;
}

/** The types of the title block's lines. */
export function isTitleBlockLine(node: Node): boolean {
  const name = node.type.name;
  return name === "title" || name === "author" || name === "date";
}

/** \and in an author line: the next author (the text after the caret goes with it); elsewhere the line an author. */
export const nextAuthor: Command = (state, dispatch, view) => {
  const { $from } = state.selection;
  const author = state.schema.nodes.author;
  if ($from.parent.type !== author) return setBlockType(author)(state, dispatch);
  return splitBlockAs(() => ({ type: author }))(state, dispatch, view);
};

/** Enter on an empty author line: body text (as Enter leaves an empty list item). */
export const leaveEmptyAuthorLine: Command = (state, dispatch) => {
  const { $cursor } = state.selection as TextSelection;
  if (!$cursor || $cursor.parent.type !== state.schema.nodes.author || $cursor.parent.content.size > 0) return false;
  return setBlockType(state.schema.nodes.paragraph)(state, dispatch);
};

/** \begin{abstract}: the line goes into an abstract; on a line that cannot (a section, the title), an empty abstract
 *  after it; in an abstract, nothing. */
export const insertAbstract: Command = (state, dispatch) => {
  const abstract = state.schema.nodes.abstract;
  const { $from } = state.selection;
  for (let d = $from.depth; d > 0; d--) if ($from.node(d).type === abstract) return true;
  if ($from.parent.type === state.schema.nodes.paragraph && wrapIn(abstract)(state, dispatch)) return true;
  if ($from.depth < 1) return false;
  const after = $from.after(1);
  const node = abstract.create(null, state.schema.nodes.paragraph.create());
  if (!state.doc.resolve(after).parent.canReplaceWith($from.index(0) + 1, $from.index(0) + 1, abstract)) return false;
  if (dispatch) {
    const tr = state.tr.insert(after, node);
    dispatch(tr.setSelection(TextSelection.create(tr.doc, after + 2)).scrollIntoView());
  }
  return true;
};

/** Whether `type` is a container a line leaves on an empty Enter (a quote, the abstract). */
export function isQuoteLike(type: NodeType): boolean {
  return type.name === "blockquote" || type.name === "abstract";
}

/** The abstract's heading follows the document's language (its class in the export). */
export const abstractLabelKey = new PluginKey<DecorationSet>("visualtex-document-abstract-label");

export function abstractLabelPlugin(): Plugin<DecorationSet> {
  const build = (doc: Node) => {
    const abstracts: Array<{ pos: number; size: number }> = [];
    doc.descendants((node, pos) => {
      if (node.type.name === "abstract") abstracts.push({ pos, size: node.nodeSize });
      return node.isBlock && !node.isTextblock;
    });
    if (!abstracts.length || !documentHasCjk(doc)) return DecorationSet.empty;
    return DecorationSet.create(
      doc,
      abstracts.map(({ pos, size }) => Decoration.node(pos, pos + size, { class: "is-zh" })),
    );
  };
  return new Plugin<DecorationSet>({
    key: abstractLabelKey,
    state: {
      init: (_config, state) => build(state.doc),
      apply: (tr, value) => (tr.docChanged ? build(tr.doc) : value),
    },
    props: { decorations: (state) => abstractLabelKey.getState(state) },
  });
}
