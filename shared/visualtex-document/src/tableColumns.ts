// The document's tables as LaTeX sets the export's tabular (latex.ts): a column is as wide as its widest cell (l, c,
// r: one line per cell), or a p{} column of a set width: one whose cells hold display math or several paragraphs
// (\linewidth shared out, as before), or one given a width by dragging its right edge (pt, kept as the cells'
// colwidth). The page view draws them so: a node view for the table (no width of its own: as wide as its columns,
// centred by the style sheet) and the p{} columns' cells marked with their width; the drag that sets a width
// (Mathcha's column resizing, in pt rather than screen pixels, so that the export writes what the page shows). A
// double click on a column's right edge gives it its natural width back.
import type { Node } from "prosemirror-model";
import { Plugin, PluginKey, type EditorState } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView, type NodeView } from "prosemirror-view";
import { TableMap } from "prosemirror-tables";

export interface TableColumn {
  /** the column's alignment: most of its cells' (the first one's on a tie; null: left) */
  align: "left" | "center" | "right" | null;
  /** a cell starting in it holds display math or several paragraphs */
  block: boolean;
  /** a width set by dragging (pt), or null */
  width: number | null;
}

/** The narrowest a dragged column gets (pt). */
const MIN_WIDTH = 12;

/** A table's columns: alignment, block content, set width (latex.ts writes them as l c r / p{}). */
export function tableColumns(table: Node): TableColumn[] {
  const map = TableMap.get(table);
  const columns: TableColumn[] = Array.from({ length: map.width }, () => ({ align: null, block: false, width: null }));
  const counts = Array.from({ length: map.width }, () => new Map<string, number>());
  for (let row = 0; row < map.height; row++) {
    for (let col = 0; col < map.width; col++) {
      const offset = map.map[row * map.width + col];
      const rect = map.findCell(offset);
      if (rect.left !== col || rect.top !== row) continue;
      const cell = table.nodeAt(offset)!;
      if (cell.childCount > 1 || cell.firstChild?.type.name === "math_display") columns[col].block = true;
      if (rect.right - rect.left === 1 && cell.textContent.trim()) {
        const align = cell.attrs.align === "center" || cell.attrs.align === "right" ? cell.attrs.align : "left";
        counts[col].set(align, (counts[col].get(align) ?? 0) + 1);
      }
      const widths = cell.attrs.colwidth as number[] | null;
      for (let k = 0; k < rect.right - rect.left; k++) {
        const width = widths?.[k];
        if (width && width > 0 && columns[col + k].width === null) columns[col + k].width = width;
      }
    }
  }
  counts.forEach((count, col) => {
    let best = 0;
    for (const [align, n] of count) {
      if (n > best) {
        best = n;
        columns[col].align = align === "left" ? null : (align as "center" | "right");
      }
    }
  });
  return columns;
}

/** A p{} column's width: its set width (pt), or its share of \linewidth (a column of blocks). */
export function columnWidth(column: TableColumn, columns: number): { pt: number } | { share: number } | null {
  if (column.width) return { pt: column.width };
  if (column.block) return { share: Number((1 / Math.max(1, columns)).toFixed(3)) };
  return null;
}

/** The table's DOM: a wrapper (the style sheet's .tableWrapper) around the table, its rows in the tbody. */
class TableView implements NodeView {
  dom: HTMLElement;
  contentDOM: HTMLElement;
  private table: HTMLTableElement;

  constructor(private node: Node) {
    this.dom = document.createElement("div");
    this.dom.className = "tableWrapper";
    this.table = this.dom.appendChild(document.createElement("table"));
    this.contentDOM = this.table.appendChild(document.createElement("tbody"));
    this.setId();
  }

  private setId() {
    if (this.node.attrs.id) this.table.setAttribute("data-id", this.node.attrs.id);
    else this.table.removeAttribute("data-id");
  }

  update(node: Node) {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.setId();
    return true;
  }

  ignoreMutation(record: MutationRecord | { type: "selection"; target: globalThis.Node }) {
    return record.type === "attributes" && (record.target === this.table || record.target === this.dom);
  }
}

interface ResizeState {
  /** the cell whose right edge the pointer is on (its position), or -1 */
  handle: number;
  /** a drag: the column's width when it started (pt), the pointer's x then, the width now */
  dragging: { startX: number; startWidth: number; width: number } | null;
}

export const tableColumnsKey = new PluginKey<ResizeState>("visualtex-table-columns");

/** Screen pixels per TeX point (the page view's zoom). */
function pxPerPt(view: EditorView): number {
  const value = parseFloat(getComputedStyle(view.dom).getPropertyValue("--vtx-pt"));
  return Number.isFinite(value) && value > 0 ? value : 96 / 72;
}

/** The column at a cell's right edge: the table's start, its map and the column's index. */
function columnAt(state: EditorState, cellPos: number): { table: Node; start: number; map: TableMap; col: number } | null {
  const $cell = state.doc.resolve(cellPos);
  const table = $cell.node(-1);
  const cell = $cell.nodeAfter;
  if (!table || table.type.name !== "table" || !cell) return null;
  const start = $cell.start(-1);
  const map = TableMap.get(table);
  return { table, start, map, col: map.colCount(cellPos - start) + (cell.attrs.colspan ?? 1) - 1 };
}

/** The cells of a column (those starting in it or spanning it): positions in the document. */
function columnCells(start: number, map: TableMap, col: number): number[] {
  const cells: number[] = [];
  for (let row = 0; row < map.height; row++) {
    const offset = map.map[row * map.width + col];
    if (row && offset === map.map[(row - 1) * map.width + col]) continue;
    cells.push(start + offset);
  }
  return cells;
}

/** The cell whose right edge is under the pointer (within 4 px), or -1. */
function handleAt(view: EditorView, event: MouseEvent): number {
  const target = (event.target as Element | null)?.closest?.("td, th");
  if (!target || !view.dom.contains(target)) return -1;
  const r = target.getBoundingClientRect();
  let cell: Element | null = null;
  if (Math.abs(event.clientX - r.right) <= 4) cell = target;
  else if (Math.abs(event.clientX - r.left) <= 4) cell = target.previousElementSibling?.matches("td, th") ? target.previousElementSibling : null;
  if (!cell) return -1;
  const pos = view.posAtDOM(cell, 0);
  const $pos = view.state.doc.resolve(pos);
  for (let d = $pos.depth; d > 0; d--) {
    const name = $pos.node(d).type.name;
    if (name === "table_cell" || name === "table_header") return $pos.before(d);
  }
  return -1;
}

function setColumnWidth(view: EditorView, cellPos: number, width: number | null) {
  const column = columnAt(view.state, cellPos);
  if (!column) return;
  const { start, map, col } = column;
  const tr = view.state.tr;
  for (const pos of columnCells(start, map, col)) {
    const cell = view.state.doc.nodeAt(pos)!;
    const colspan = cell.attrs.colspan ?? 1;
    const index = col - map.colCount(pos - start);
    const widths: number[] = cell.attrs.colwidth ? (cell.attrs.colwidth as number[]).slice() : new Array(colspan).fill(0);
    widths[index] = width ?? 0;
    tr.setNodeMarkup(pos, null, { ...cell.attrs, colwidth: widths.some((w) => w > 0) ? widths : null });
  }
  if (tr.docChanged) view.dispatch(tr);
}

/** The p{} columns' cells: marked, with their width (and a dragged column's width while it is dragged). */
function columnDecorations(state: EditorState): DecorationSet {
  const resize = tableColumnsKey.getState(state);
  const drag = resize?.dragging && resize.handle >= 0 ? columnAt(state, resize.handle) : null;
  const decorations: Decoration[] = [];
  state.doc.descendants((node, pos) => {
    if (node.type.name !== "table") return !node.isTextblock;
    const columns = tableColumns(node);
    const map = TableMap.get(node);
    const start = pos + 1;
    for (let row = 0; row < map.height; row++) {
      for (let col = 0; col < map.width; col++) {
        const offset = map.map[row * map.width + col];
        const rect = map.findCell(offset);
        if (rect.left !== col || rect.top !== row) continue;
        const cell = node.nodeAt(offset)!;
        let width: string | null = null;
        if (rect.right - rect.left === 1) {
          const dragged = drag && drag.start === start && drag.col === col ? resize!.dragging!.width : null;
          const p = dragged ? { pt: dragged } : columnWidth(columns[col], map.width);
          if (p) width = "pt" in p ? `calc(${p.pt} * var(--vtx-pt, 1.333px))` : `calc(${p.share} * var(--vtx-linewidth, 100%))`;
        }
        if (width) decorations.push(Decoration.node(start + offset, start + offset + cell.nodeSize, { class: "is-p", style: `width:${width}` }));
      }
    }
    // the dragged edge's handle (the style sheet's .column-resize-handle) on the column's cells
    if (resize && resize.handle >= 0) {
      const at = columnAt(state, resize.handle);
      if (at && at.start === start) {
        for (const cellPos of columnCells(start, map, at.col)) {
          const cell = state.doc.nodeAt(cellPos)!;
          const offset = cellPos - start;
          const rect = map.findCell(offset);
          if (rect.right - 1 !== at.col) continue;
          const handle = document.createElement("div");
          handle.className = "column-resize-handle";
          decorations.push(Decoration.widget(cellPos + cell.nodeSize - 1, handle));
        }
      }
    }
    return false;
  });
  return DecorationSet.create(state.doc, decorations);
}

/** The tables' node view, their p{} columns, and the drag that sets a column's width. */
export function tableColumnsPlugin(): Plugin<ResizeState> {
  let cache: { doc: Node; resize: ResizeState | undefined; set: DecorationSet } | null = null;
  return new Plugin<ResizeState>({
    key: tableColumnsKey,
    state: {
      init: () => ({ handle: -1, dragging: null }),
      apply: (tr, value) => {
        const meta = tr.getMeta(tableColumnsKey) as Partial<ResizeState> | undefined;
        let next = meta ? { ...value, ...meta } : value;
        if (tr.docChanged && next.handle >= 0 && !next.dragging) {
          const handle = tr.mapping.mapResult(next.handle, 1);
          next = { ...next, handle: handle.deleted ? -1 : handle.pos };
        }
        return next;
      },
    },
    props: {
      nodeViews: { table: (node) => new TableView(node) },
      decorations: (state) => {
        // once per document and drag state
        const resize = tableColumnsKey.getState(state);
        if (cache && cache.doc === state.doc && cache.resize === resize) return cache.set;
        cache = { doc: state.doc, resize, set: columnDecorations(state) };
        return cache.set;
      },
      attributes: (state): Record<string, string> => ((tableColumnsKey.getState(state)?.handle ?? -1) >= 0 ? { class: "resize-cursor" } : {}),
      handleDOMEvents: {
        mousemove: (view, event) => {
          const state = tableColumnsKey.getState(view.state);
          if (!state || state.dragging || !view.editable) return false;
          const handle = event.buttons ? -1 : handleAt(view, event);
          if (handle !== state.handle) view.dispatch(view.state.tr.setMeta(tableColumnsKey, { handle }));
          return false;
        },
        mouseleave: (view) => {
          const state = tableColumnsKey.getState(view.state);
          if (state && state.handle >= 0 && !state.dragging) view.dispatch(view.state.tr.setMeta(tableColumnsKey, { handle: -1 }));
          return false;
        },
        dblclick: (view, event) => {
          const state = tableColumnsKey.getState(view.state);
          if (!state || state.handle < 0) return false;
          event.preventDefault();
          setColumnWidth(view, state.handle, null);
          return true;
        },
        mousedown: (view, event) => {
          const state = tableColumnsKey.getState(view.state);
          if (!state || state.handle < 0 || event.button !== 0 || !view.editable || event.detail > 1) return false;
          const column = columnAt(view.state, state.handle);
          if (!column) return false;
          event.preventDefault();
          const scale = pxPerPt(view);
          // the column's width now: a cell of its own, without its \tabcolsep on each side
          const own = columnCells(column.start, column.map, column.col).find((pos) => (view.state.doc.nodeAt(pos)?.attrs.colspan ?? 1) === 1);
          const dom = own !== undefined ? view.nodeDOM(own) : null;
          const startWidth = dom instanceof HTMLElement ? Math.max(MIN_WIDTH, dom.getBoundingClientRect().width / scale - 12) : 60;
          const handle = state.handle;
          view.dispatch(view.state.tr.setMeta(tableColumnsKey, { dragging: { startX: event.clientX, startWidth, width: startWidth } }));
          const widthAt = (e: MouseEvent) => Math.round(Math.max(MIN_WIDTH, startWidth + (e.clientX - event.clientX) / scale) * 2) / 2;
          const move = (e: MouseEvent) => {
            if (!e.buttons) return finish(e);
            const dragging = tableColumnsKey.getState(view.state)?.dragging;
            if (dragging) view.dispatch(view.state.tr.setMeta(tableColumnsKey, { dragging: { ...dragging, width: widthAt(e) } }));
          };
          const finish = (e: MouseEvent) => {
            window.removeEventListener("mousemove", move);
            window.removeEventListener("mouseup", finish);
            const moved = Math.abs(e.clientX - event.clientX) > 1;
            view.dispatch(view.state.tr.setMeta(tableColumnsKey, { dragging: null }));
            if (moved) setColumnWidth(view, handle, widthAt(e));
          };
          window.addEventListener("mousemove", move);
          window.addEventListener("mouseup", finish);
          return true;
        },
      },
    },
  });
}
