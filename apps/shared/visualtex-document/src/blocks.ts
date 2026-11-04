// Paragraph, section, list and table commands with Mathcha's behaviour (MATHCHA-SPEC.md): sections are a
// toggled style with three numbered levels, Enter keeps the line's style, Tab is a tab character in text, the
// next cell in a table, an indent in a list and a demotion on a section.
import { Fragment, type Node, type NodeType } from "prosemirror-model";
import { liftTarget } from "prosemirror-transform";
import { NodeSelection, Selection, TextSelection, type Command, type EditorState, type Transaction } from "prosemirror-state";
import { setBlockType, splitBlockAs } from "prosemirror-commands";
import { liftListItem, sinkListItem, splitListItem, wrapInList } from "prosemirror-schema-list";
import {
  TableMap,
  addColumnAfter,
  addColumnBefore,
  addRowAfter,
  addRowBefore,
  deleteColumn,
  deleteRow,
  deleteTable,
  isInTable,
  mergeCells,
  setCellAttr,
  splitCell,
} from "prosemirror-tables";

/** The top-level block holding the selection, with its position. */
export function currentBlock(state: EditorState): { node: Node; pos: number } | null {
  const $from = state.selection.$from;
  if ($from.depth < 1) {
    const node = state.doc.nodeAt(state.selection.from);
    return node ? { node, pos: state.selection.from } : null;
  }
  return { node: $from.node(1), pos: $from.before(1) };
}

/** The textblock holding the selection's start (a paragraph or a section), with its position. */
function textblockAt(state: EditorState): { node: Node; pos: number } | null {
  const $from = state.selection.$from;
  return $from.parent.isTextblock && $from.depth > 0 ? { node: $from.parent, pos: $from.before() } : null;
}

function listDepth(state: EditorState): { depth: number; type: NodeType } | null {
  const { bullet_list, ordered_list } = state.schema.nodes;
  const $from = state.selection.$from;
  for (let d = $from.depth; d > 0; d--) {
    const type = $from.node(d).type;
    if (type === bullet_list || type === ordered_list) return { depth: d, type };
  }
  return null;
}

/** The list kind holding the caret, if any. */
export function currentListKind(state: EditorState): "ordered_list" | "bullet_list" | null {
  return (listDepth(state)?.type.name as "ordered_list" | "bullet_list" | undefined) ?? null;
}

/** The section level (1-3) of the caret's line, or 0 for body text. */
export function sectionLevel(state: EditorState): number {
  const block = textblockAt(state);
  return block?.node.type.name === "heading" ? block.node.attrs.level : 0;
}

/** Ctrl+Alt+1: the line becomes a level-1 section, or body text again. */
export const toggleSection: Command = (state, dispatch) => {
  const block = textblockAt(state);
  if (!block) return false;
  const { heading, paragraph } = state.schema.nodes;
  const keep = { align: block.node.attrs.align ?? null, indent: block.node.attrs.indent ?? 0 };
  if (block.node.type === heading) return setBlockType(paragraph, keep)(state, dispatch);
  return setBlockType(heading, { ...keep, level: 1 })(state, dispatch);
};

/** Whether the caret's line is the document's title. */
export function isTitleLine(state: EditorState): boolean {
  return textblockAt(state)?.node.type.name === "title";
}

/** Ctrl+Alt+0: the line becomes the document's title, or body text again. */
export const toggleTitle: Command = (state, dispatch) => {
  const block = textblockAt(state);
  if (!block || block.node.type.spec.code) return false;
  const { title, paragraph } = state.schema.nodes;
  return setBlockType(block.node.type === title ? paragraph : title)(state, dispatch);
};

/** Tab / Shift+Tab on a section: one level deeper / higher (1 ↔ 2 ↔ 3). */
export function changeSectionLevel(delta: 1 | -1): Command {
  return (state, dispatch) => {
    const block = textblockAt(state);
    if (!block || block.node.type.name !== "heading") return false;
    const level = Math.max(1, Math.min(3, block.node.attrs.level + delta));
    if (dispatch && level !== block.node.attrs.level) dispatch(state.tr.setNodeAttribute(block.pos, "level", level));
    return true;
  };
}

/** Ctrl+Shift+7 / 8: an ordered / unordered list; again: back to body text; the other kind: switch. */
export function toggleList(kind: "ordered_list" | "bullet_list"): Command {
  return (state, dispatch) => {
    const type = state.schema.nodes[kind];
    const list = listDepth(state);
    if (list?.type === type) return liftListItem(state.schema.nodes.list_item)(state, dispatch);
    if (list) {
      if (dispatch) {
        const $from = state.selection.$from;
        dispatch(state.tr.setNodeMarkup($from.before(list.depth), type, { ...$from.node(list.depth).attrs, start: 1 }));
      }
      return true;
    }
    // A section (or the title) becomes body text first: list items hold paragraphs.
    const block = textblockAt(state);
    if (block?.node.type.name !== "heading" && block?.node.type.name !== "title") return wrapInList(type)(state, dispatch);
    const tr = state.tr.setBlockType(block.pos, block.pos + block.node.nodeSize, state.schema.nodes.paragraph, {
      align: block.node.attrs.align ?? null,
      indent: 0,
    });
    const next = state.apply(tr);
    let wrapped = false;
    wrapInList(type)(next, (wrap) => {
      for (const step of wrap.steps) tr.step(step);
      tr.setSelection(Selection.fromJSON(tr.doc, next.apply(wrap).selection.toJSON()));
      wrapped = true;
    });
    if (wrapped && dispatch) dispatch(tr);
    return wrapped;
  };
}

/** Ctrl+Shift+L / E / R / J: the paragraphs and sections in the selection are aligned. */
export function setAlign(align: "left" | "center" | "right" | "justify"): Command {
  return (state, dispatch) => {
    const { from, to } = state.selection;
    const tr = state.tr;
    state.doc.nodesBetween(from, to, (node, pos) => {
      if (node.isTextblock && "align" in node.attrs) {
        tr.setNodeAttribute(pos, "align", align === "left" ? null : align);
        return false;
      }
      return true;
    });
    if (!tr.docChanged) return false;
    if (dispatch) dispatch(tr);
    return true;
  };
}

/** The alignment of the caret's line. */
export function currentAlign(state: EditorState): "left" | "center" | "right" | "justify" {
  const block = textblockAt(state);
  return (block?.node.attrs.align as "center" | "right" | "justify" | null) ?? "left";
}

/** Mathcha's Indent / Outdent: a list item nests one level deeper / higher, a paragraph moves one step. */
export function indent(delta: 1 | -1): Command {
  return (state, dispatch) => {
    const { list_item } = state.schema.nodes;
    if (listDepth(state)) return (delta > 0 ? sinkListItem(list_item) : liftListItem(list_item))(state, dispatch);
    const block = textblockAt(state);
    if (!block || !("indent" in block.node.attrs)) return false;
    const next = Math.max(0, Math.min(8, (block.node.attrs.indent ?? 0) + delta));
    if (next === block.node.attrs.indent) return false;
    if (dispatch) dispatch(state.tr.setNodeAttribute(block.pos, "indent", next));
    return true;
  };
}

/** Tab in body text: a tab character (Mathcha). */
export const insertTabCharacter: Command = (state, dispatch) => {
  if (!state.selection.$from.parent.isTextblock) return false;
  if (dispatch) dispatch(state.tr.insertText("\t").scrollIntoView());
  return true;
};

/**
 * Backspace at the start of a styled line, as in Word: the style goes first — a section becomes body text, a list
 * item leaves its list, a quote ends, an indent steps back. Only a plain line joins the one before it.
 */
export const clearLineStyle: Command = (state, dispatch) => {
  const { $cursor } = state.selection as TextSelection;
  if (!$cursor || $cursor.parentOffset > 0 || !$cursor.parent.isTextblock) return false;
  const block = $cursor.parent;
  const { heading, title, paragraph, list_item, blockquote } = state.schema.nodes;
  if (block.type === heading || block.type === title) {
    if (dispatch) {
      const attrs = { align: block.attrs.align ?? null, indent: block.attrs.indent ?? 0 };
      dispatch(state.tr.setBlockType($cursor.pos, $cursor.pos, paragraph, attrs).scrollIntoView());
    }
    return true;
  }
  // The first line of a list item or a quote: out of the list / quote.
  const container = $cursor.depth >= 2 && $cursor.index(-1) === 0 ? $cursor.node(-1).type : null;
  if (container === list_item) return liftListItem(list_item)(state, dispatch);
  if (container === blockquote) {
    const range = $cursor.blockRange();
    const target = range && liftTarget(range);
    if (range && target !== null && target !== undefined) {
      if (dispatch) dispatch(state.tr.lift(range, target).scrollIntoView());
      return true;
    }
  }
  if (block.type === paragraph && (block.attrs.indent ?? 0) > 0) {
    if (dispatch) dispatch(state.tr.setNodeAttribute($cursor.before(), "indent", block.attrs.indent - 1));
    return true;
  }
  return false;
};

/** Enter: a new line in the same style (a section stays a section of the same level, as in Mathcha); after the
 *  title, body text. */
export const splitKeepingStyle: Command = (state, dispatch, view) => {
  const { $from } = state.selection;
  if (!$from.parent.isTextblock || $from.parent.type.spec.code) return false;
  const { title, paragraph } = state.schema.nodes;
  return splitBlockAs((node) => (node.type === title ? { type: paragraph } : { type: node.type, attrs: { ...node.attrs, id: null } }))(
    state,
    dispatch,
    view,
  );
};

/** Enter in a list item: a new item; on an empty item: out of the list (Mathcha). */
export const splitListItemOrLeave: Command = (state, dispatch, view) => {
  const { $from, empty } = state.selection;
  const itemType = state.schema.nodes.list_item;
  const inItem = $from.depth >= 2 && $from.node(-1).type === itemType;
  if (inItem && empty && $from.parent.content.size === 0) return liftListItem(itemType)(state, dispatch, view);
  return splitListItem(itemType)(state, dispatch, view);
};

/**
 * Backspace at the start of a line right after a display formula or a rule, Delete at the end of a line right before
 * one: the block is selected first (a second press deletes it) instead of vanishing with one key — Typora deletes
 * it silently, which its users report as lost work. Next to a code block the caret goes into it instead of merging
 * the two (Typora).
 */
export function selectAdjacentBlock(direction: -1 | 1): Command {
  return (state, dispatch, view) => {
    const { $cursor } = state.selection as TextSelection;
    if (!$cursor || !$cursor.parent.isTextblock || $cursor.depth < 1) return false;
    const atEdge = direction < 0 ? $cursor.parentOffset === 0 : $cursor.parentOffset === $cursor.parent.content.size;
    if (!atEdge || (view && !view.endOfTextblock(direction < 0 ? "backward" : "forward"))) return false;
    const index = $cursor.index(-1);
    const parent = $cursor.node(-1);
    const neighbourIndex = index + direction;
    if (neighbourIndex < 0 || neighbourIndex >= parent.childCount) return false;
    const neighbour = parent.child(neighbourIndex);
    const name = neighbour.type.name;
    const pos = direction < 0 ? $cursor.before() - neighbour.nodeSize : $cursor.after();
    if (name === "math_display" || name === "horizontal_rule") {
      if (dispatch) dispatch(state.tr.setSelection(NodeSelection.create(state.doc, pos)).scrollIntoView());
      return true;
    }
    if (name === "code_block") {
      const caret = direction < 0 ? pos + neighbour.nodeSize - 1 : pos + 1;
      if (dispatch) dispatch(state.tr.setSelection(TextSelection.create(state.doc, caret)).scrollIntoView());
      return true;
    }
    return false;
  };
}

/** Enter on an empty line in a quote: out of the quote (as on an empty list item; Typora, Obsidian). */
export const leaveEmptyQuoteLine: Command = (state, dispatch) => {
  const { $cursor } = state.selection as TextSelection;
  const { paragraph, blockquote } = state.schema.nodes;
  if (!$cursor || $cursor.parent.type !== paragraph || $cursor.parent.content.size > 0) return false;
  if ($cursor.depth < 2 || $cursor.node(-1).type !== blockquote) return false;
  const range = $cursor.blockRange();
  const target = range && liftTarget(range);
  if (!range || target === null || target === undefined) return false;
  if (dispatch) dispatch(state.tr.lift(range, target).scrollIntoView());
  return true;
};

/** The table cell holding the caret: its table, row and column. */
function tableCellAt(state: EditorState) {
  const $from = state.selection.$from;
  for (let d = $from.depth; d > 2; d--) {
    const name = $from.node(d).type.name;
    if (name !== "table_cell" && name !== "table_header") continue;
    const table = $from.node(d - 2);
    const tableStart = $from.start(d - 2);
    const map = TableMap.get(table);
    const rect = map.findCell($from.before(d) - tableStart);
    return { table, tableStart, tablePos: $from.before(d - 2), map, row: rect.top, col: rect.left };
  }
  return null;
}

/** The caret at the end of the text of the cell at (row, col). */
function caretInCell(tr: Transaction, tableStart: number, map: TableMap, row: number, col: number) {
  const cellPos = map.map[row * map.width + col];
  const cell = tr.doc.nodeAt(tableStart + cellPos);
  if (!cell) return;
  const sel = Selection.findFrom(tr.doc.resolve(tableStart + cellPos + cell.nodeSize - 1), -1, true);
  if (sel) tr.setSelection(sel);
}

/** Enter in a table (Markdown input, as in Typora): the cell below; in the last row, out of the table. */
export const tableEnter: Command = (state, dispatch) => {
  const at = tableCellAt(state);
  if (!at || !state.selection.empty) return false;
  if (!dispatch) return true;
  const tr = state.tr;
  if (at.row < at.map.height - 1) caretInCell(tr, at.tableStart, at.map, at.row + 1, at.col);
  else {
    const after = at.tablePos + at.table.nodeSize;
    if (!tr.doc.resolve(after).nodeAfter?.isTextblock) tr.insert(after, state.schema.nodes.paragraph.create());
    tr.setSelection(TextSelection.create(tr.doc, after + 1));
  }
  dispatch(tr.scrollIntoView());
  return true;
};

/** Ctrl+Enter in a table: an empty row below this one, the caret in it (Typora). */
export const tableRowBelow: Command = (state, dispatch) => {
  const at = tableCellAt(state);
  if (!at) return false;
  if (!dispatch) return true;
  let added = false;
  addRowAfter(state, (tr) => {
    const table = tr.doc.nodeAt(at.tablePos);
    if (table) caretInCell(tr, at.tablePos + 1, TableMap.get(table), at.row + 1, at.col);
    dispatch(tr.scrollIntoView());
    added = true;
  });
  return added;
};

/** \horizontal-line: a rule after the current block. */
export const insertDivider: Command = (state, dispatch) => {
  const block = currentBlock(state);
  if (!block) return false;
  if (dispatch) {
    const at = block.pos + block.node.nodeSize;
    const tr = state.tr.insert(at, [state.schema.nodes.horizontal_rule.create(), state.schema.nodes.paragraph.create()]);
    dispatch(tr.setSelection(TextSelection.create(tr.doc, at + 2)).scrollIntoView());
  }
  return true;
};

/**
 * A page command (\newpage, \pagenumbering{…}…) where the caret is: in an empty line it takes the line's place, at a
 * line's start it goes before it, at its end after it, in its middle the line is split there (TeX ends the paragraph).
 */
export function insertPageCommand(command: string, value = ""): Command {
  return (state, dispatch) => {
    const type = state.schema.nodes.page_command;
    const $from = state.selection.$from;
    if (!type || !$from.parent.isTextblock || $from.parent.type.spec.code) return false;
    if (!dispatch) return true;
    const node = type.create({ command, value });
    const paragraph = $from.parent;
    const start = $from.before();
    const tr = state.tr;
    if (paragraph.content.size === 0) {
      tr.insert(start, node);
      tr.setSelection(TextSelection.create(tr.doc, start + node.nodeSize + 1));
    } else if ($from.parentOffset === 0) {
      tr.insert(start, node);
    } else if ($from.parentOffset === paragraph.content.size) {
      const after = $from.after();
      tr.insert(after, [node, state.schema.nodes.paragraph.create()]);
      tr.setSelection(TextSelection.create(tr.doc, after + node.nodeSize + 1));
    } else {
      tr.split($from.pos);
      const at = tr.mapping.map($from.pos) - 1;
      tr.insert(at, node);
      tr.setSelection(TextSelection.create(tr.doc, at + node.nodeSize + 1));
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** \table: a rows x cols table after the current block, or in place of an empty paragraph; caret in its first cell. */
export function insertTable(rows = 3, cols = 3): Command {
  return (state, dispatch) => {
    if (isInTable(state)) return false;
    const { nodes } = state.schema;
    if (!dispatch) return true;
    const cell = () => nodes.table_cell.create(null, nodes.paragraph.create());
    const rowNodes: Node[] = [];
    for (let r = 0; r < rows; r++) rowNodes.push(nodes.table_row.create(null, Fragment.from(Array.from({ length: cols }, cell))));
    const table = nodes.table.create(null, rowNodes);
    const $from = state.selection.$from;
    const tr = state.tr;
    let at: number;
    if ($from.depth === 1 && $from.parent.type.name === "paragraph" && $from.parent.content.size === 0) {
      at = $from.before();
      tr.replaceWith(at, $from.after(), table);
    } else {
      const block = currentBlock(state);
      at = block ? block.pos + block.node.nodeSize : state.doc.content.size;
      tr.insert(at, table);
    }
    // A paragraph after the table, so the text can go on.
    const after = at + table.nodeSize;
    if (after >= tr.doc.content.size || tr.doc.resolve(after).nodeAfter?.type.name !== "paragraph") {
      tr.insert(after, nodes.paragraph.create());
    }
    // table > row > cell > paragraph: the caret in the first cell's text
    dispatch(tr.setSelection(TextSelection.create(tr.doc, at + 4)).scrollIntoView());
    return true;
  };
}

/** Tab / Shift+Tab in a table: the end of the next (previous) cell's text; Tab in the last cell adds a row. */
export function goToCell(direction: -1 | 1): Command {
  return (state, dispatch) => {
    const $from = state.selection.$from;
    let cellDepth = -1;
    for (let d = $from.depth; d > 0; d--) {
      const name = $from.node(d).type.name;
      if (name === "table_cell" || name === "table_header") {
        cellDepth = d;
        break;
      }
    }
    if (cellDepth < 0) return false;
    const table = $from.node(cellDepth - 2);
    const tableStart = $from.start(cellDepth - 2);
    const map = TableMap.get(table);
    const cellPos = $from.before(cellDepth) - tableStart;
    const cells = [...new Set(map.map)];
    const target = cells[cells.indexOf(cellPos) + direction];
    if (!dispatch) return true;
    if (target === undefined) {
      if (direction < 0) return true;
      let added = false;
      addRowAfter(state, (tr) => {
        const after = tr.doc.resolve(tr.mapping.map(tableStart + cellPos + table.nodeAt(cellPos)!.nodeSize));
        const next = Selection.findFrom(tr.doc.resolve(Math.min(after.end(after.depth) + 1, tr.doc.content.size)), 1, true);
        if (next) tr.setSelection(next);
        dispatch(tr.scrollIntoView());
        added = true;
      });
      return added;
    }
    const cell = table.nodeAt(target)!;
    const sel = Selection.findFrom(state.doc.resolve(tableStart + target + cell.nodeSize - 1), -1, true);
    if (sel) dispatch(state.tr.setSelection(sel).scrollIntoView());
    return true;
  };
}

/** The table commands of Mathcha's context menu. */
export const tableCommands = {
  addRowBefore,
  addRowAfter,
  addColumnBefore,
  addColumnAfter,
  deleteRow,
  deleteColumn,
  deleteTable,
  mergeCells,
  splitCell,
  alignLeft: setCellAttr("align", null),
  alignCenter: setCellAttr("align", "center"),
  alignRight: setCellAttr("align", "right"),
} satisfies Record<string, Command>;

export { isInTable };
