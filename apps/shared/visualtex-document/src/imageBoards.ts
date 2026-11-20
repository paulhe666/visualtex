// Image boards (DESIGN.md §11.6; research ~/claude/document-images/research/free-layout.md): two or more images side
// by side are one block, as wide as the line, in which each image has a free place - its left edge `x` and its top
// edge `top`, fractions of the line width (\linewidth), as its width is. The board is as high as its lowest image;
// it never splits across pages. Images move in it as in PowerPoint: guides for edges, centres and equal gaps while
// dragging (Alt: none, Shift: one axis, Ctrl: a copy, Esc: back), corner handles keep the aspect, the arrow keys
// nudge, and the bar aligns, distributes and tidies. Dragged out of the board an image goes where an image drop goes.
// LaTeX: a plain row is written as an image line; otherwise one picture environment with \unitlength = \linewidth
// and a \put per image (kernel LaTeX, exact). Markdown: rows of <img> with margins (or, when images overlap, an
// absolutely placed group), each image carrying data-x/data-top so VisualTeX reads the board back exactly.
import { Fragment, type Node, type NodeSpec, type Schema } from "prosemirror-model";
import { EditorState, NodeSelection, Plugin, PluginKey, TextSelection, type Transaction } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView, type NodeView } from "prosemirror-view";
import { closeHistory } from "prosemirror-history";
import { DEFAULT_IMAGE_WIDTH, IMAGE_GAP, imageAspect, imageDataUrl, imageEntry, imageUrl, registerDataUrl, isImage, isImageLine, lineMetrics, roundWidth, type ImageDropTarget } from "./images.ts";

// --------------------------------------------------------------------- model --

const id = { id: { default: null as string | null } };

export const imageBoardNodes: Record<string, NodeSpec> = {
  image_board: {
    group: "block",
    content: "board_image+",
    isolating: true,
    selectable: true,
    draggable: false,
    attrs: id,
    parseDOM: [{ tag: "div.vtx-image-board", getAttrs: (dom: HTMLElement | string) => (typeof dom === "string" ? {} : { id: dom.getAttribute("data-id") }) }],
    toDOM: (node) => ["div", { class: "vtx-image-board", ...(node.attrs.id ? { "data-id": node.attrs.id } : {}) }, 0],
  },
  board_image: {
    atom: true,
    selectable: true,
    draggable: false,
    attrs: { ...id, src: { default: "" }, alt: { default: "" }, width: { default: DEFAULT_IMAGE_WIDTH }, x: { default: 0 }, top: { default: 0 } },
    parseDOM: [
      {
        tag: "img[data-board-x]",
        getAttrs: (dom: HTMLElement | string) => {
          if (typeof dom === "string") return false;
          const stored = dom.getAttribute("data-key");
          const key = stored && imageEntry(stored) ? stored : registerDataUrl(dom.getAttribute("src") ?? "");
          if (!key) return false;
          return {
            src: key,
            alt: dom.getAttribute("alt") ?? "",
            width: Number(dom.getAttribute("data-width")) || DEFAULT_IMAGE_WIDTH,
            x: Number(dom.getAttribute("data-board-x")) || 0,
            top: Number(dom.getAttribute("data-board-top")) || 0,
          };
        },
      },
    ],
    // what a copy carries: the bytes (another document or app gets the image)
    toDOM: (node) => [
      "img",
      {
        src: imageDataUrl(node.attrs.src),
        "data-key": node.attrs.src,
        alt: node.attrs.alt,
        "data-width": String(node.attrs.width),
        "data-board-x": String(node.attrs.x),
        "data-board-top": String(node.attrs.top),
      },
    ],
  },
};

export const isBoard = (node: Node | null | undefined): boolean => node?.type.name === "image_board";
export const isBoardImage = (node: Node | null | undefined): boolean => node?.type.name === "board_image";

const round3 = (value: number) => Math.round(value * 1000) / 1000;

/** An image's height on the board (a fraction of the line width). */
export function boardImageHeight(node: Node): number {
  return node.attrs.width / imageAspect(node.attrs.src);
}

export interface BoardItem {
  node: Node;
  index: number;
  x: number;
  top: number;
  width: number;
  height: number;
}

export function boardItems(board: Node): BoardItem[] {
  const items: BoardItem[] = [];
  board.forEach((node, _offset, index) => {
    items.push({ node, index, x: node.attrs.x, top: node.attrs.top, width: node.attrs.width, height: boardImageHeight(node) });
  });
  return items;
}

/** The board's height: its lowest image's bottom. */
export function boardHeight(board: Node): number {
  let height = 0;
  board.forEach((node) => {
    height = Math.max(height, node.attrs.top + boardImageHeight(node));
  });
  return height;
}

/**
 * The board as a plain row - one bottom line, the standard gap between neighbours, the row at the left, centre or
 * right - or null. Such a board is written as an image line (and an image line is shown as such a board).
 */
export function canonicalRow(board: Node): { align: "left" | "center" | "right"; items: BoardItem[] } | null {
  const items = boardItems(board).sort((a, b) => a.x - b.x);
  if (!items.length) return null;
  const bottom = items[0].top + items[0].height;
  for (const item of items) if (Math.abs(item.top + item.height - bottom) > 0.0015) return null;
  for (let i = 1; i < items.length; i++) {
    if (Math.abs(items[i].x - (items[i - 1].x + items[i - 1].width) - IMAGE_GAP) > 0.0015) return null;
  }
  const total = items[items.length - 1].x + items[items.length - 1].width - items[0].x;
  const start = items[0].x;
  if (Math.abs(start) <= 0.0015) return { align: "left", items };
  if (Math.abs(start - (1 - total) / 2) <= 0.0015) return { align: "center", items };
  if (Math.abs(start + total - 1) <= 0.0015) return { align: "right", items };
  return null;
}

/** An image line's images as board images, where the line shows them (bottoms on one line). */
export function lineAsBoard(schema: Schema, line: Node): Node {
  const images: Node[] = [];
  line.forEach((child) => {
    if (isImage(child)) images.push(child);
  });
  const total = images.reduce((sum, image) => sum + image.attrs.width, 0) + (images.length - 1) * IMAGE_GAP;
  const align = line.attrs.align as string | null;
  let x = align === "center" ? (1 - total) / 2 : align === "right" ? 1 - total : 0;
  x = Math.max(0, x);
  const heights = images.map((image) => image.attrs.width / imageAspect(image.attrs.src));
  const row = Math.max(...heights);
  const children = images.map((image, i) => {
    const node = schema.nodes.board_image.create({ id: image.attrs.id, src: image.attrs.src, alt: image.attrs.alt, width: image.attrs.width, x: round3(x), top: round3(row - heights[i]) });
    x += image.attrs.width + IMAGE_GAP;
    return node;
  });
  return schema.nodes.image_board.create({ id: line.attrs.id }, children);
}

/** A board of one image at the left, centre or right edge (its top 0) as an image line, or null. */
function boardAsLine(schema: Schema, board: Node): Node | null {
  if (board.childCount !== 1) return null;
  const image = board.firstChild!;
  const { x, width } = image.attrs;
  const align = Math.abs(x) <= 0.002 ? null : Math.abs(x - (1 - width) / 2) <= 0.002 ? "center" : Math.abs(x + width - 1) <= 0.002 ? "right" : undefined;
  if (align === undefined) return null;
  return schema.nodes.paragraph.create({ id: board.attrs.id, align }, schema.nodes.image.create({ id: image.attrs.id, src: image.attrs.src, alt: image.attrs.alt, width: image.attrs.width }));
}

/**
 * What a board keeps: images two or more side by side are a board (an image line of them becomes one, where it
 * shows them); a board's top image starts at its top; images stay inside the line; a board of one image at an edge
 * is an image line again; an empty board goes.
 */
export function boardFixes(state: EditorState): Transaction | null {
  const schema = state.schema;
  if (!schema.nodes.image_board) return null;
  const changes: Array<{ pos: number; node: Node; replacement: Node | null }> = [];
  state.doc.descendants((node, pos) => {
    if (node.type.name === "paragraph") {
      if (isImageLine(node) && !node.attrs.indent) {
        let images = 0;
        node.forEach((child) => {
          if (isImage(child)) images++;
        });
        if (images >= 2) changes.push({ pos, node, replacement: lineAsBoard(schema, node) });
      }
      return false;
    }
    if (!isBoard(node)) return true;
    if (node.childCount === 0) {
      changes.push({ pos, node, replacement: null });
      return false;
    }
    const line = boardAsLine(schema, node);
    if (line) {
      changes.push({ pos, node, replacement: line });
      return false;
    }
    let minTop = Infinity;
    node.forEach((child) => {
      minTop = Math.min(minTop, child.attrs.top);
    });
    const children: Node[] = [];
    let changed = false;
    node.forEach((child) => {
      const width = Math.min(1, child.attrs.width);
      const x = round3(Math.max(0, Math.min(1 - width, child.attrs.x)));
      const top = round3(child.attrs.top - (minTop > 0.0005 || minTop < 0 ? minTop : 0));
      if (x !== child.attrs.x || top !== child.attrs.top || width !== child.attrs.width) {
        changed = true;
        children.push(child.type.create({ ...child.attrs, x, top, width }));
      } else children.push(child);
    });
    if (changed) changes.push({ pos, node, replacement: node.copy(Fragment.from(children)) });
    return false;
  });
  if (!changes.length) return null;
  const tr = state.tr;
  for (const change of changes.reverse()) {
    if (change.replacement) tr.replaceWith(change.pos, change.pos + change.node.nodeSize, change.replacement);
    else tr.delete(change.pos, change.pos + change.node.nodeSize);
  }
  // a selected image stays selected (the same place in the replaced block)
  const sel = state.selection;
  if (sel instanceof NodeSelection && (isImage(sel.node) || isBoardImage(sel.node))) {
    const id = sel.node.attrs.id;
    let found = -1;
    tr.doc.descendants((node, pos) => {
      if (found < 0 && id && (isImage(node) || isBoardImage(node)) && node.attrs.id === id) found = pos;
      return found < 0;
    });
    if (found >= 0) tr.setSelection(NodeSelection.create(tr.doc, found));
  }
  return tr;
}

// ---------------------------------------------------------------- selection --

export const boardsKey = new PluginKey<BoardSelection>("visualtex-document-image-boards");

/** The images selected on a board besides the selected one (Shift/Ctrl+click): their ids. */
interface BoardSelection {
  extra: string[];
}

/** The selected board images: the node selection and the ones added to it. Positions in the document, in order. */
export function selectedBoardImages(state: EditorState): { boardPos: number; positions: number[] } | null {
  const sel = state.selection;
  if (!(sel instanceof NodeSelection) || !isBoardImage(sel.node)) return null;
  const $pos = state.doc.resolve(sel.from);
  const board = $pos.parent;
  const boardPos = $pos.before();
  const extra = new Set(boardsKey.getState(state)?.extra ?? []);
  const positions: number[] = [];
  board.forEach((child, offset) => {
    const pos = boardPos + 1 + offset;
    if (pos === sel.from || (child.attrs.id && extra.has(child.attrs.id))) positions.push(pos);
  });
  return { boardPos, positions };
}

// ------------------------------------------------------------------ commands --

/** Each moved image's new place (and width), as one undo step; the moved ones stay selected. */
function placeImages(view: EditorView, places: Map<number, { x?: number; top?: number; width?: number }>, select?: number) {
  const state = view.state;
  const tr = closeHistory(state.tr);
  for (const [pos, place] of places) {
    const node = state.doc.nodeAt(pos);
    if (!node || !isBoardImage(node)) continue;
    const width = place.width !== undefined ? roundWidth(place.width) : node.attrs.width;
    tr.setNodeMarkup(pos, undefined, {
      ...node.attrs,
      width,
      x: round3(place.x ?? node.attrs.x),
      top: round3(place.top ?? node.attrs.top),
    });
  }
  const sel = state.selection;
  if (select !== undefined) tr.setSelection(NodeSelection.create(tr.doc, select));
  else if (sel instanceof NodeSelection) tr.setSelection(NodeSelection.create(tr.doc, sel.from));
  view.dispatch(tr);
}

export type BoardCommand =
  | "alignLeft"
  | "alignCenter"
  | "alignRight"
  | "alignTop"
  | "alignMiddle"
  | "alignBottom"
  | "distributeH"
  | "distributeV"
  | "sameWidth"
  | "tidyRow"
  | "tidyGrid"
  | "front"
  | "back";

/** A bar command on the selected board images (one of them: against the board). */
export function boardCommand(view: EditorView, command: BoardCommand): boolean {
  const selected = selectedBoardImages(view.state);
  if (!selected) return false;
  const state = view.state;
  const board = state.doc.nodeAt(selected.boardPos)!;
  const all = boardItems(board).map((item) => ({ ...item, pos: selected.boardPos + 1 + offsetOf(board, item.index) }));
  const chosen = all.filter((item) => selected.positions.includes(item.pos));
  const places = new Map<number, { x?: number; top?: number; width?: number }>();
  const H = boardHeight(board);
  const left = chosen.length > 1 ? Math.min(...chosen.map((i) => i.x)) : 0;
  const right = chosen.length > 1 ? Math.max(...chosen.map((i) => i.x + i.width)) : 1;
  const top = chosen.length > 1 ? Math.min(...chosen.map((i) => i.top)) : 0;
  const bottom = chosen.length > 1 ? Math.max(...chosen.map((i) => i.top + i.height)) : H;
  switch (command) {
    case "alignLeft":
      for (const item of chosen) places.set(item.pos, { x: left });
      break;
    case "alignCenter":
      for (const item of chosen) places.set(item.pos, { x: (left + right) / 2 - item.width / 2 });
      break;
    case "alignRight":
      for (const item of chosen) places.set(item.pos, { x: right - item.width });
      break;
    case "alignTop":
      for (const item of chosen) places.set(item.pos, { top });
      break;
    case "alignMiddle":
      for (const item of chosen) places.set(item.pos, { top: (top + bottom) / 2 - item.height / 2 });
      break;
    case "alignBottom":
      for (const item of chosen) places.set(item.pos, { top: bottom - item.height });
      break;
    case "distributeH":
    case "distributeV": {
      const group = chosen.length >= 3 ? chosen : all;
      if (group.length < 3) return false;
      const horizontal = command === "distributeH";
      const sorted = [...group].sort((a, b) => (horizontal ? a.x - b.x : a.top - b.top));
      const first = sorted[0];
      const last = sorted[sorted.length - 1];
      const span = horizontal ? last.x + last.width - first.x : last.top + last.height - first.top;
      const sizes = sorted.reduce((sum, item) => sum + (horizontal ? item.width : item.height), 0);
      const gap = (span - sizes) / (sorted.length - 1);
      let at = horizontal ? first.x : first.top;
      for (const item of sorted) {
        places.set(item.pos, horizontal ? { x: at } : { top: at });
        at += (horizontal ? item.width : item.height) + gap;
      }
      break;
    }
    case "sameWidth": {
      const sel = state.selection as NodeSelection;
      const width = sel.node.attrs.width;
      for (const item of chosen) places.set(item.pos, { width });
      break;
    }
    case "tidyRow": {
      // all of the board's images on one bottom line, the standard gap apart, centred (fitted when too wide)
      const sorted = [...all].sort((a, b) => a.x - b.x || a.top - b.top);
      let widths = sorted.map((item) => item.width);
      const room = 0.99 - (sorted.length - 1) * IMAGE_GAP;
      const sum = widths.reduce((a, b) => a + b, 0);
      if (sum > room) widths = widths.map((w) => Math.floor(((w * room) / sum) * 1000) / 1000);
      const heights = sorted.map((item, i) => widths[i] / imageAspect(item.node.attrs.src));
      const row = Math.max(...heights);
      let x = (1 - (widths.reduce((a, b) => a + b, 0) + (sorted.length - 1) * IMAGE_GAP)) / 2;
      sorted.forEach((item, i) => {
        places.set(item.pos, { x, top: row - heights[i], width: widths[i] });
        x += widths[i] + IMAGE_GAP;
      });
      break;
    }
    case "tidyGrid": {
      // rows as they are (images overlapping in height share a row), columns of one width, centred in the line
      const rows = rowsOf(all);
      const columns = Math.max(...rows.map((row) => row.length));
      const cell = (1 - (columns - 1) * IMAGE_GAP) / columns;
      let y = 0;
      for (const row of rows) {
        const widths = row.map((item) => Math.min(item.width, cell));
        const heights = row.map((item, i) => widths[i] / imageAspect(item.node.attrs.src));
        const rowHeight = Math.max(...heights);
        const used = row.length * cell + (row.length - 1) * IMAGE_GAP;
        let x = (1 - used) / 2;
        row.forEach((item, i) => {
          places.set(item.pos, { x: x + (cell - widths[i]) / 2, top: y + (rowHeight - heights[i]) / 2, width: widths[i] });
          x += cell + IMAGE_GAP;
        });
        y += rowHeight + IMAGE_GAP;
      }
      break;
    }
    case "front":
    case "back": {
      // the paint order is the node order
      const sel = state.selection as NodeSelection;
      const node = sel.node;
      const tr = closeHistory(state.tr).delete(sel.from, sel.to);
      const boardNow = tr.doc.nodeAt(selected.boardPos)!;
      const at = command === "front" ? selected.boardPos + boardNow.nodeSize - 1 : selected.boardPos + 1;
      tr.insert(at, node);
      tr.setSelection(NodeSelection.create(tr.doc, at));
      view.dispatch(tr);
      return true;
    }
  }
  if (!places.size) return false;
  placeImages(view, places);
  return true;
}

function offsetOf(board: Node, index: number): number {
  let offset = 0;
  for (let i = 0; i < index; i++) offset += board.child(i).nodeSize;
  return offset;
}

/** Images grouped in rows: those whose heights overlap (one after another) share a row; rows top to bottom. */
function rowsOf<T extends { x: number; top: number; height: number }>(items: T[]): T[][] {
  const sorted = [...items].sort((a, b) => a.top - b.top);
  const rows: T[][] = [];
  let bottom = -Infinity;
  for (const item of sorted) {
    if (rows.length && item.top < bottom - 0.002) {
      rows[rows.length - 1].push(item);
      bottom = Math.max(bottom, item.top + item.height);
    } else {
      rows.push([item]);
      bottom = item.top + item.height;
    }
  }
  return rows.map((row) => row.sort((a, b) => a.x - b.x));
}

/** The rows of a board for Markdown: null when images overlap (in a row, side by side). */
export function boardRows(board: Node): BoardItem[][] | null {
  const rows = rowsOf(boardItems(board));
  for (const row of rows) {
    for (let i = 1; i < row.length; i++) if (row[i].x < row[i - 1].x + row[i - 1].width - 0.0005) return null;
  }
  return rows;
}

// ------------------------------------------------------------------- drops --

/** A board under a page point: an image dropped there lands at that place on it (centred on the pointer). */
export function boardDropTarget(view: EditorView, x: number, y: number, width: number, aspect: number, exceptPos = -1): ImageDropTarget | null {
  const element = document.elementFromPoint(x, y)?.closest?.(".vtx-image-board");
  if (!element || !view.dom.contains(element)) return null;
  const inside = view.posAtDOM(element, 0);
  const boardPos = inside - 1;
  const board = boardPos >= 0 ? view.state.doc.nodeAt(boardPos) : null;
  if (!board || !isBoard(board) || boardPos === exceptPos) return null;
  const rect = element.getBoundingClientRect();
  const w = Math.min(1, width);
  const left = (x - rect.left) / rect.width - w / 2;
  const top = (y - rect.top) / rect.width - w / aspect / 2;
  return { kind: "board", pos: boardPos, x: round3(Math.max(0, Math.min(1 - w, left))), top: round3(top) };
}

/** Where a board drop lands: the image's outline there. */
export function boardDropMarker(view: EditorView, target: Extract<ImageDropTarget, { kind: "board" }>, width: number, aspect: number, element: HTMLElement): boolean {
  const dom = view.nodeDOM(target.pos) as HTMLElement | null;
  if (!dom) return false;
  const rect = dom.getBoundingClientRect();
  Object.assign(element.style, {
    left: `${rect.left + target.x * rect.width}px`,
    top: `${rect.top + target.top * rect.width}px`,
    width: `${Math.min(1, width) * rect.width}px`,
    height: `${(Math.min(1, width) / aspect) * rect.width}px`,
  });
  return true;
}

/** An image node (from a line, a sentence or another board) as a board image of the board at `boardPos`. */
export function addToBoard(tr: Transaction, boardPos: number, image: Node, x: number, top: number): number {
  const board = tr.doc.nodeAt(boardPos)!;
  const at = boardPos + board.nodeSize - 1;
  tr.insert(at, tr.doc.type.schema.nodes.board_image.create({ id: null, src: image.attrs.src, alt: image.attrs.alt, width: Math.min(1, image.attrs.width), x, top }));
  return at;
}

/** A board image taken off its board to an image drop target (a line, a place in text or a line, a wrap). */
export function moveBoardImageOut(view: EditorView, target: ImageDropTarget, from: number, copy: boolean): Transaction | null {
  const state = view.state;
  const node = state.doc.nodeAt(from);
  if (!node || !isBoardImage(node)) return null;
  const schema = state.schema;
  const tr = closeHistory(state.tr);
  if (!copy) tr.delete(from, from + node.nodeSize);
  const attrs = { src: node.attrs.src, alt: node.attrs.alt, width: node.attrs.width, wrap: null as string | null };
  if (target.kind === "board") {
    const at = addToBoard(tr, tr.mapping.map(target.pos), node, target.x, target.top);
    tr.setSelection(NodeSelection.create(tr.doc, at));
    return tr;
  }
  if (target.kind === "line") {
    const pos = tr.mapping.map(target.pos);
    tr.insert(pos, schema.nodes.paragraph.create({ align: "center" }, schema.nodes.image.create(attrs)));
    tr.setSelection(NodeSelection.create(tr.doc, pos + 1));
    return tr;
  }
  if (target.kind === "wrap") {
    const pos = tr.mapping.map(target.pos);
    tr.insert(pos + 1, schema.nodes.image.create({ ...attrs, wrap: target.side, width: Math.min(0.7, attrs.width) }));
    tr.setSelection(NodeSelection.create(tr.doc, pos + 1));
    return tr;
  }
  // into a sentence: as high as the text (measured before the change)
  const metrics = lineMetrics(view, target.pos);
  const $target = state.doc.resolve(target.pos);
  if (metrics && !isImageLine($target.parent) && $target.parent.content.size > 0) {
    const aspect = imageAspect(node.attrs.src);
    if ((attrs.width * metrics.width) / aspect > metrics.font * 2) attrs.width = roundWidth((metrics.font * aspect) / metrics.width);
  }
  const pos = tr.mapping.map(target.pos);
  tr.insert(pos, schema.nodes.image.create(attrs));
  tr.setSelection(NodeSelection.create(tr.doc, pos));
  return tr;
}

// --------------------------------------------------------------------- drag --

const SNAP_PX = 6;
const LEAVE_PX = 32;
const SNAPS = [0.25, 1 / 3, 0.5, 2 / 3, 0.75, 1];

interface Box {
  x: number;
  top: number;
  width: number;
  height: number;
}

/** Guides drawn over the page while an image moves: lines through aligned edges, gaps of equal size. */
class Guides {
  private elements: HTMLElement[] = [];

  show(lines: Array<{ x1: number; y1: number; x2: number; y2: number; gap?: boolean }>) {
    this.clear();
    for (const line of lines) {
      const el = document.createElement("div");
      el.className = line.gap ? "vtx-board-guide is-gap" : "vtx-board-guide";
      const vertical = Math.abs(line.x1 - line.x2) < 0.5;
      Object.assign(el.style, vertical
        ? { left: `${line.x1}px`, top: `${Math.min(line.y1, line.y2)}px`, width: "1px", height: `${Math.abs(line.y2 - line.y1)}px` }
        : { left: `${Math.min(line.x1, line.x2)}px`, top: `${line.y1}px`, width: `${Math.abs(line.x2 - line.x1)}px`, height: "1px" });
      document.body.append(el);
      this.elements.push(el);
    }
  }

  clear() {
    for (const el of this.elements) el.remove();
    this.elements = [];
  }
}

interface Snap {
  delta: number;
  lines: Array<{ x1: number; y1: number; x2: number; y2: number; gap?: boolean }>;
}

/**
 * The nearest snap of a moving box on one axis: its edges and centre to the other images' edges and centres and to
 * the line's edges and centre; its gap to a neighbour to a gap already on the board (or the standard gap).
 */
function snapAxis(box: Box, others: Box[], horizontal: boolean, tolerance: number, rect: DOMRect, height: number): Snap | null {
  const start = horizontal ? box.x : box.top;
  const size = horizontal ? box.width : box.height;
  const ends = [start, start + size / 2, start + size];
  const targets: Array<{ value: number; kind: 0 | 1 | 2 | 3; from?: Box }> = [];
  for (const other of others) {
    const s = horizontal ? other.x : other.top;
    const z = horizontal ? other.width : other.height;
    targets.push({ value: s, kind: 0, from: other }, { value: s + z / 2, kind: 1, from: other }, { value: s + z, kind: 2, from: other });
  }
  if (horizontal) targets.push({ value: 0, kind: 3 }, { value: 0.5, kind: 3 }, { value: 1, kind: 3 });
  let best: { delta: number; value: number; from?: Box } | null = null;
  for (const target of targets) {
    for (const end of ends) {
      const delta = target.value - end;
      if (Math.abs(delta) <= tolerance && (!best || Math.abs(delta) < Math.abs(best.delta))) best = { delta, value: target.value, from: target.from };
    }
  }
  // gaps: to the nearest neighbour on each side, matched to the board's other gaps and the standard gap
  const gaps = new Set<number>([IMAGE_GAP]);
  for (const a of others) {
    for (const b of others) {
      if (a === b) continue;
      const overlapOther = horizontal ? a.top < b.top + b.height && b.top < a.top + a.height : a.x < b.x + b.width && b.x < a.x + a.width;
      const gap = horizontal ? b.x - (a.x + a.width) : b.top - (a.top + a.height);
      if (overlapOther && gap > 0.002) gaps.add(round3(gap));
    }
  }
  let gapSnap: { delta: number; a: number; b: number; across: number } | null = null;
  for (const other of others) {
    const overlap = horizontal ? box.top < other.top + other.height && other.top < box.top + box.height : box.x < other.x + other.width && other.x < box.x + box.width;
    if (!overlap) continue;
    const s = horizontal ? other.x : other.top;
    const z = horizontal ? other.width : other.height;
    const across = horizontal ? Math.max(box.top, other.top) + 0.01 : Math.max(box.x, other.x) + 0.01;
    for (const gap of gaps) {
      // this box after the other, or before it
      for (const [delta, a, b] of [
        [s + z + gap - start, s + z, s + z + gap],
        [s - gap - size - start, s - gap, s],
      ] as Array<[number, number, number]>) {
        if (Math.abs(delta) <= tolerance && (!gapSnap || Math.abs(delta) < Math.abs(gapSnap.delta)) && (!best || Math.abs(delta) < Math.abs(best.delta))) {
          gapSnap = { delta, a, b, across };
        }
      }
    }
  }
  const W = rect.width;
  if (gapSnap) {
    const lines = horizontal
      ? [{ x1: rect.left + gapSnap.a * W, y1: rect.top + gapSnap.across * W, x2: rect.left + gapSnap.b * W, y2: rect.top + gapSnap.across * W, gap: true }]
      : [{ x1: rect.left + gapSnap.across * W, y1: rect.top + gapSnap.a * W, x2: rect.left + gapSnap.across * W, y2: rect.top + gapSnap.b * W, gap: true }];
    return { delta: gapSnap.delta, lines };
  }
  if (!best) return null;
  const v = best.value;
  const lines = horizontal
    ? [{ x1: rect.left + v * W, y1: rect.top - 8, x2: rect.left + v * W, y2: rect.top + Math.max(height, box.top + box.height) * W + 8 }]
    : [{ x1: rect.left - 8, y1: rect.top + v * W, x2: rect.right + 8, y2: rect.top + v * W }];
  return { delta: best.delta, lines };
}

/**
 * A press on a board image: it (or the selected ones) follow the pointer with guides; released, they are placed as
 * one undo step. Far outside the board the image is dropped where an image drop goes (a line, text, another board).
 */
export class BoardDrag {
  private moved = false;
  private cancelled = false;
  private out = false;
  private readonly guides = new Guides();
  private marker: HTMLElement | null = null;
  private readonly boardDom: HTMLElement;
  private readonly items: Array<{ pos: number; dom: HTMLElement; box: Box }>;
  private readonly primary: { pos: number; dom: HTMLElement; box: Box };
  private readonly others: Box[];
  private delta = { x: 0, y: 0 };
  /** the board's box when the drag started (it grows while an image moves down) */
  private startRect: DOMRect | null = null;
  private outTarget: ImageDropTarget | null = null;

  constructor(
    private readonly view: EditorView,
    private readonly start: PointerEvent,
    private readonly boardPos: number,
    pressedPos: number,
    private readonly outTargetAt: (x: number, y: number, image: Node, from: number) => ImageDropTarget | null,
    private readonly outMarker: (target: ImageDropTarget | null, element: HTMLElement) => boolean,
  ) {
    const state = view.state;
    const board = state.doc.nodeAt(boardPos)!;
    this.boardDom = view.nodeDOM(boardPos) as HTMLElement;
    const selected = selectedBoardImages(state)?.positions ?? [pressedPos];
    const moving = selected.includes(pressedPos) ? selected : [pressedPos];
    this.items = [];
    const others: Box[] = [];
    board.forEach((child, offset) => {
      const pos = boardPos + 1 + offset;
      const box = { x: child.attrs.x, top: child.attrs.top, width: child.attrs.width, height: boardImageHeight(child) };
      if (moving.includes(pos)) this.items.push({ pos, dom: view.nodeDOM(pos) as HTMLElement, box });
      else others.push(box);
    });
    this.others = others;
    this.primary = this.items.find((item) => item.pos === pressedPos) ?? this.items[0];
    window.addEventListener("pointermove", this.onMove, true);
    window.addEventListener("pointerup", this.onUp, true);
    window.addEventListener("pointercancel", this.onCancel, true);
    window.addEventListener("keydown", this.onKey, true);
  }

  private finish() {
    window.removeEventListener("pointermove", this.onMove, true);
    window.removeEventListener("pointerup", this.onUp, true);
    window.removeEventListener("pointercancel", this.onCancel, true);
    window.removeEventListener("keydown", this.onKey, true);
    this.guides.clear();
    this.marker?.remove();
    this.boardDom.classList.remove("is-arranging");
    for (const item of this.items) {
      item.dom.style.removeProperty("transform");
      item.dom.classList.remove("is-leaving");
    }
    document.documentElement.style.removeProperty("cursor");
  }

  private onMove = (event: PointerEvent) => {
    if (event.pointerId !== this.start.pointerId || this.cancelled) return;
    const dxPx = event.clientX - this.start.clientX;
    const dyPx = event.clientY - this.start.clientY;
    if (!this.moved) {
      if (Math.abs(dxPx) < 3 && Math.abs(dyPx) < 3) return;
      this.moved = true;
      this.boardDom.classList.add("is-arranging");
    }
    event.preventDefault();
    const rect = this.boardDom.getBoundingClientRect();
    const start = (this.startRect ??= rect);
    document.documentElement.style.setProperty("cursor", event.ctrlKey ? "copy" : "move", "important");
    // out of the board: the pointer beside it or above it, or the image a clear gap below its bottom (a new row
    // right under the others grows the board) - then an image drop elsewhere (one image only)
    const imageTop = (this.primary.box.top + dyPx / start.width) * start.width;
    const outside =
      event.clientX < start.left - LEAVE_PX ||
      event.clientX > start.right + LEAVE_PX ||
      event.clientY < start.top - LEAVE_PX ||
      imageTop > start.height + LEAVE_PX;
    if (outside && this.items.length === 1) {
      this.out = true;
      this.guides.clear();
      for (const item of this.items) {
        item.dom.style.removeProperty("transform");
        item.dom.classList.add("is-leaving");
      }
      const node = this.view.state.doc.nodeAt(this.primary.pos)!;
      this.outTarget = this.outTargetAt(event.clientX, event.clientY, node, this.primary.pos);
      if (!this.marker) {
        this.marker = document.createElement("div");
        document.body.append(this.marker);
      }
      this.marker.className = `vtx-drop-marker is-caret${this.outTarget?.kind === "board" ? " is-box" : this.outTarget?.kind === "wrap" ? " is-wrap" : ""}`;
      this.marker.hidden = !this.outMarker(this.outTarget, this.marker);
      return;
    }
    this.out = false;
    this.marker?.remove();
    this.marker = null;
    for (const item of this.items) item.dom.classList.remove("is-leaving");
    const W = rect.width;
    let dx = dxPx / W;
    let dy = dyPx / W;
    if (event.shiftKey) {
      if (Math.abs(dx) > Math.abs(dy)) dy = 0;
      else dx = 0;
    }
    const lines: Snap["lines"] = [];
    if (!event.altKey) {
      const box = { ...this.primary.box, x: this.primary.box.x + dx, top: this.primary.box.top + dy };
      const tolerance = SNAP_PX / W;
      const height = Math.max(0, ...this.others.map((o) => o.top + o.height));
      const sx = event.shiftKey && dx === 0 ? null : snapAxis(box, this.others, true, tolerance, rect, height);
      const sy = event.shiftKey && dy === 0 ? null : snapAxis(box, this.others, false, tolerance, rect, height);
      if (sx) {
        dx += sx.delta;
        lines.push(...sx.lines);
      }
      if (sy) {
        dy += sy.delta;
        lines.push(...sy.lines);
      }
    }
    // inside the line's width
    const minX = Math.max(...this.items.map((item) => -item.box.x));
    const maxX = Math.min(...this.items.map((item) => 1 - item.box.width - item.box.x));
    dx = Math.max(minX, Math.min(maxX, dx));
    this.delta = { x: dx, y: dy };
    for (const item of this.items) item.dom.style.transform = `translate(${dx * W}px, ${dy * W}px)`;
    // the board grows below as an image moves down
    const bottom = Math.max(...this.others.map((o) => o.top + o.height), ...this.items.map((item) => item.box.top + dy + item.box.height));
    this.boardDom.style.paddingBottom = `${Math.max(bottom, 0) * 100}%`;
    this.guides.show(lines);
  };

  private onUp = (event: PointerEvent) => {
    if (event.pointerId !== this.start.pointerId) return;
    this.finish();
    const view = this.view;
    if (this.cancelled) {
      restoreBoard(view, this.boardPos);
      return;
    }
    if (!this.moved) {
      view.focus();
      return;
    }
    event.preventDefault();
    const copy = event.ctrlKey;
    if (this.out) {
      const tr = this.outTarget ? moveBoardImageOut(view, this.outTarget, this.primary.pos, copy) : null;
      if (tr) view.dispatch(tr.setMeta("uiEvent", "drop").scrollIntoView());
      else restoreBoard(view, this.boardPos);
      view.focus();
      return;
    }
    const { x: dx, y: dy } = this.delta;
    const state = view.state;
    const tr = closeHistory(state.tr);
    if (copy) {
      for (const item of this.items) {
        const node = state.doc.nodeAt(item.pos)!;
        addToBoard(tr, this.boardPos, node, round3(item.box.x + dx), round3(item.box.top + dy));
      }
      tr.setMeta(boardsKey, { extra: [] });
    } else {
      for (const item of this.items) {
        const node = state.doc.nodeAt(item.pos)!;
        tr.setNodeMarkup(item.pos, undefined, { ...node.attrs, x: round3(item.box.x + dx), top: round3(item.box.top + dy) });
      }
      const sel = state.selection;
      if (sel instanceof NodeSelection) tr.setSelection(NodeSelection.create(tr.doc, sel.from));
    }
    view.dispatch(tr);
    restoreBoard(view, this.boardPos);
    view.focus();
  };

  private onCancel = (event: PointerEvent) => {
    if (event.pointerId === this.start.pointerId) this.cancel();
  };

  private onKey = (event: KeyboardEvent) => {
    if (event.key !== "Escape") return;
    event.preventDefault();
    event.stopImmediatePropagation();
    this.cancel();
  };

  private cancel() {
    this.cancelled = true;
    this.finish();
    restoreBoard(this.view, this.boardPos);
  }
}

/** The board's drawn height back to its images' (after a live move). */
function restoreBoard(view: EditorView, boardPos: number) {
  const board = view.state.doc.nodeAt(boardPos);
  const dom = view.nodeDOM(boardPos) as HTMLElement | null;
  if (board && dom && isBoard(board)) dom.style.paddingBottom = `${boardHeight(board) * 100}%`;
}

// --------------------------------------------------------------------- views --

class BoardView implements NodeView {
  dom: HTMLElement;
  contentDOM: HTMLElement;

  constructor(private node: Node) {
    this.dom = document.createElement("div");
    this.dom.className = "vtx-image-board";
    this.dom.setAttribute("role", "group");
    this.contentDOM = this.dom;
    this.render();
  }

  private render() {
    if (this.node.attrs.id) this.dom.dataset.id = this.node.attrs.id;
    this.dom.style.paddingBottom = `${boardHeight(this.node) * 100}%`;
  }

  update(node: Node): boolean {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.render();
    return true;
  }

  ignoreMutation(mutation: { type: string; target: globalThis.Node }): boolean {
    // the board's own style (its height); its images are their views'
    return mutation.type === "attributes" && mutation.target === this.dom;
  }
}

class BoardImageView implements NodeView {
  dom: HTMLElement;
  private img: HTMLImageElement;
  private label: HTMLElement | null = null;

  constructor(
    private node: Node,
    private view: EditorView,
    private getPos: () => number | undefined,
  ) {
    this.dom = document.createElement("div");
    this.dom.className = "vtx-board-image";
    this.dom.contentEditable = "false";
    this.img = document.createElement("img");
    this.img.draggable = false;
    this.dom.append(this.img);
    for (const corner of ["nw", "ne", "sw", "se"] as const) {
      const handle = document.createElement("span");
      handle.className = `vtx-board-handle is-${corner}`;
      handle.addEventListener("pointerdown", (event) => this.resize(event, corner));
      this.dom.append(handle);
    }
    this.render();
  }

  private render() {
    const { src, width, x, top, alt } = this.node.attrs;
    const url = imageUrl(src);
    if (this.img.getAttribute("src") !== url) this.img.src = url;
    this.img.alt = alt;
    this.dom.style.left = `${x * 100}%`;
    this.dom.style.top = `${top * 100}cqw`;
    this.dom.style.width = `${width * 100}%`;
    this.dom.style.aspectRatio = String(imageAspect(src));
    this.dom.classList.toggle("is-missing", !url);
  }

  update(node: Node): boolean {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.render();
    return true;
  }

  selectNode() {
    this.dom.classList.add("is-selected");
  }

  deselectNode() {
    this.dom.classList.remove("is-selected");
  }

  stopEvent(event: Event): boolean {
    return (event.target as Element | null)?.classList?.contains("vtx-board-handle") ?? false;
  }

  ignoreMutation(): boolean {
    return true;
  }

  /** A corner dragged: the width follows (the aspect kept, the opposite corner where it is), snapping. */
  private resize(event: PointerEvent, corner: "nw" | "ne" | "sw" | "se") {
    if (event.button !== 0) return;
    event.preventDefault();
    event.stopPropagation();
    const pos = this.getPos();
    if (pos === undefined) return;
    const view = this.view;
    const boardDom = this.dom.parentElement;
    if (!boardDom) return;
    const W = boardDom.getBoundingClientRect().width;
    const aspect = imageAspect(this.node.attrs.src);
    const { x, top, width } = this.node.attrs as { x: number; top: number; width: number };
    const height = width / aspect;
    const west = corner === "nw" || corner === "sw";
    const north = corner === "nw" || corner === "ne";
    const others = new Set<number>();
    view.state.doc.descendants((node, at) => {
      if ((isImage(node) || isBoardImage(node)) && at !== pos) others.add(round3(node.attrs.width));
      return true;
    });
    const label = document.createElement("div");
    label.className = "vtx-image-size";
    document.body.append(label);
    this.label = label;
    let next = { x, top, width };
    const move = (e: PointerEvent) => {
      const dx = (e.clientX - event.clientX) / W;
      let w = Math.max(0.03, width + (west ? -dx : dx));
      // the opposite edge stays: the room up to the line's edge on the dragged side
      w = Math.min(w, west ? x + width : 1 - x);
      if (!e.altKey) {
        for (const snap of [...SNAPS, ...others]) {
          if (Math.abs(w - snap) <= 0.01 && snap <= (west ? x + width : 1 - x)) {
            w = snap;
            break;
          }
        }
      }
      const h = w / aspect;
      next = { x: west ? x + width - w : x, top: north ? top + height - h : top, width: w };
      this.dom.style.left = `${next.x * 100}%`;
      this.dom.style.top = `${next.top * 100}cqw`;
      this.dom.style.width = `${w * 100}%`;
      label.textContent = `${Math.round(w * 100)}%`;
      const r = this.dom.getBoundingClientRect();
      label.style.left = `${west ? r.left : r.right}px`;
      label.style.top = `${r.top - 26}px`;
    };
    const up = () => {
      window.removeEventListener("pointermove", move, true);
      window.removeEventListener("pointerup", up, true);
      label.remove();
      this.label = null;
      const now = this.getPos();
      if (now !== undefined && Math.abs(next.width - width) > 0.0005) {
        placeImages(view, new Map([[now, next]]), now);
      } else this.render();
    };
    window.addEventListener("pointermove", move, true);
    window.addEventListener("pointerup", up, true);
    move(event);
  }

  destroy() {
    this.label?.remove();
  }
}

// ----------------------------------------------------------------------- bar --

export interface BoardLabels {
  alignLeft: string;
  alignCenter: string;
  alignRight: string;
  alignTop: string;
  alignMiddle: string;
  alignBottom: string;
  distributeH: string;
  distributeV: string;
  sameWidth: string;
  tidyRow: string;
  tidyGrid: string;
  front: string;
  back: string;
}

export const DEFAULT_BOARD_LABELS: BoardLabels = {
  alignLeft: "Align left",
  alignCenter: "Align centre",
  alignRight: "Align right",
  alignTop: "Align top",
  alignMiddle: "Align middle",
  alignBottom: "Align bottom",
  distributeH: "Distribute horizontally",
  distributeV: "Distribute vertically",
  sameWidth: "Same width",
  tidyRow: "Tidy into a row",
  tidyGrid: "Tidy into a grid",
  front: "Bring to front",
  back: "Send to back",
};

const BOARD_ICONS: Record<keyof BoardLabels, string> = {
  alignLeft: '<path d="M4 3v18"/><rect x="7" y="6" width="10" height="4" rx="1"/><rect x="7" y="14" width="6" height="4" rx="1"/>',
  alignCenter: '<path d="M12 3v18"/><rect x="6" y="6" width="12" height="4" rx="1"/><rect x="8" y="14" width="8" height="4" rx="1"/>',
  alignRight: '<path d="M20 3v18"/><rect x="7" y="6" width="10" height="4" rx="1"/><rect x="11" y="14" width="6" height="4" rx="1"/>',
  alignTop: '<path d="M3 4h18"/><rect x="6" y="7" width="4" height="10" rx="1"/><rect x="14" y="7" width="4" height="6" rx="1"/>',
  alignMiddle: '<path d="M3 12h18"/><rect x="6" y="6" width="4" height="12" rx="1"/><rect x="14" y="8" width="4" height="8" rx="1"/>',
  alignBottom: '<path d="M3 20h18"/><rect x="6" y="7" width="4" height="10" rx="1"/><rect x="14" y="11" width="4" height="6" rx="1"/>',
  distributeH: '<path d="M3 4v16M21 4v16"/><rect x="9.5" y="7" width="5" height="10" rx="1"/>',
  distributeV: '<path d="M4 3h16M4 21h16"/><rect x="7" y="9.5" width="10" height="5" rx="1"/>',
  sameWidth: '<rect x="5" y="4" width="14" height="6" rx="1"/><rect x="5" y="14" width="14" height="6" rx="1"/>',
  tidyRow: '<rect x="3" y="8" width="5" height="8" rx="1"/><rect x="9.5" y="8" width="5" height="8" rx="1"/><rect x="16" y="8" width="5" height="8" rx="1"/>',
  tidyGrid: '<rect x="4" y="4" width="7" height="7" rx="1"/><rect x="13" y="4" width="7" height="7" rx="1"/><rect x="4" y="13" width="7" height="7" rx="1"/><rect x="13" y="13" width="7" height="7" rx="1"/>',
  front: '<rect x="4" y="4" width="10" height="10" rx="1"/><rect x="10" y="10" width="10" height="10" rx="1" fill="currentColor"/>',
  back: '<rect x="10" y="10" width="10" height="10" rx="1"/><rect x="4" y="4" width="10" height="10" rx="1" fill="currentColor"/>',
};

/** The bar above a selected board image: align, distribute, tidy, order. */
class BoardBar {
  private element: HTMLElement;
  private buttons = new Map<keyof BoardLabels, HTMLButtonElement>();

  constructor(
    private view: EditorView,
    labels: BoardLabels,
  ) {
    this.element = document.createElement("div");
    this.element.className = "vtx-image-bar is-board";
    this.element.setAttribute("role", "toolbar");
    const groups: Array<Array<keyof BoardLabels>> = [
      ["alignLeft", "alignCenter", "alignRight", "alignTop", "alignMiddle", "alignBottom"],
      ["distributeH", "distributeV", "sameWidth"],
      ["tidyRow", "tidyGrid"],
      ["front", "back"],
    ];
    groups.forEach((group, index) => {
      if (index) {
        const gap = document.createElement("span");
        gap.className = "vtx-image-bar-gap";
        this.element.append(gap);
      }
      for (const name of group) {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "vtx-image-bar-button";
        button.title = labels[name];
        button.setAttribute("aria-label", labels[name]);
        button.innerHTML = `<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">${BOARD_ICONS[name]}</svg>`;
        button.addEventListener("pointerdown", (event) => event.preventDefault());
        button.addEventListener("click", (event) => {
          event.preventDefault();
          boardCommand(this.view, name);
          this.view.focus();
        });
        this.element.append(button);
        this.buttons.set(name, button);
      }
    });
    this.element.hidden = true;
    document.body.append(this.element);
    window.addEventListener("scroll", this.place, true);
    window.addEventListener("resize", this.place);
  }

  place = () => {
    const view = this.view;
    const selected = selectedBoardImages(view.state);
    if (!selected || (!view.hasFocus() && !this.element.contains(document.activeElement))) {
      this.element.hidden = true;
      return;
    }
    const board = view.nodeDOM(selected.boardPos) as HTMLElement | null;
    if (!board) {
      this.element.hidden = true;
      return;
    }
    const count = view.state.doc.nodeAt(selected.boardPos)!.childCount;
    const several = selected.positions.length;
    for (const [name, button] of this.buttons) {
      button.disabled = (name === "distributeH" || name === "distributeV") ? (several >= 3 ? false : count < 3) : name === "sameWidth" ? several < 2 : false;
    }
    this.element.hidden = false;
    const dom = view.nodeDOM((view.state.selection as NodeSelection).from) as HTMLElement | null;
    const r = (dom ?? board).getBoundingClientRect();
    const rb = board.getBoundingClientRect();
    const bar = this.element.getBoundingClientRect();
    const left = Math.max(4, Math.min(window.innerWidth - bar.width - 4, (rb.left + rb.right) / 2 - bar.width / 2));
    const above = Math.min(r.top, rb.top);
    const top = above - bar.height - 8 < 4 ? Math.max(r.bottom, rb.bottom) + 8 : above - bar.height - 8;
    this.element.style.left = `${left}px`;
    this.element.style.top = `${top}px`;
  };

  destroy() {
    window.removeEventListener("scroll", this.place, true);
    window.removeEventListener("resize", this.place);
    this.element.remove();
  }
}

// -------------------------------------------------------------------- plugin --

const NUDGE = { plain: 0.005, fine: 0.001, coarse: IMAGE_GAP };

export interface BoardsPluginOptions {
  labels: () => BoardLabels;
}

export function imageBoardsPlugin(options: BoardsPluginOptions): Plugin<BoardSelection> {
  return new Plugin<BoardSelection>({
    key: boardsKey,
    state: {
      init: () => ({ extra: [] }),
      apply(tr, value, _old, state) {
        const meta = tr.getMeta(boardsKey) as BoardSelection | undefined;
        if (meta) return meta;
        // an extra selection lasts while one of its board's images is selected
        if (!value.extra.length) return value;
        const sel = state.selection;
        if (!(sel instanceof NodeSelection) || !isBoardImage(sel.node)) return { extra: [] };
        return value;
      },
    },
    appendTransaction(trs, _old, state) {
      if (!trs.some((tr) => tr.docChanged)) return null;
      return boardFixes(state);
    },
    view(view) {
      const bar = new BoardBar(view, options.labels());
      const raf = () => requestAnimationFrame(bar.place);
      raf();
      return { update: raf, destroy: () => bar.destroy() };
    },
    props: {
      nodeViews: {
        image_board: (node) => new BoardView(node),
        board_image: (node, view, getPos) => new BoardImageView(node, view, getPos),
      },
      decorations(state) {
        const extra = boardsKey.getState(state)?.extra ?? [];
        if (!extra.length) return null;
        const decorations: Decoration[] = [];
        state.doc.descendants((node, pos) => {
          if (isBoardImage(node) && node.attrs.id && extra.includes(node.attrs.id)) decorations.push(Decoration.node(pos, pos + node.nodeSize, { class: "is-also-selected" }));
          return node.isBlock;
        });
        return DecorationSet.create(state.doc, decorations);
      },
      handleKeyDown(view, event) {
        const state = view.state;
        const sel = state.selection;
        // after the board's figure (figures.ts), not between the board and its caption
        const afterFigure = (pos: number) => {
          const $pos = state.doc.resolve(pos);
          return $pos.parent.type.name === "figure" ? $pos.after() : pos;
        };
        if (sel instanceof NodeSelection && isBoard(sel.node) && event.key === "Enter" && !event.shiftKey && !event.ctrlKey && !event.metaKey) {
          const after = afterFigure(sel.to);
          const tr = state.tr.insert(after, state.schema.nodes.paragraph.create());
          tr.setSelection(TextSelection.create(tr.doc, after + 1));
          view.dispatch(tr.scrollIntoView());
          return true;
        }
        const selected = selectedBoardImages(state);
        if (!selected) return false;
        if (event.key === "Enter" && !event.shiftKey && !event.ctrlKey && !event.metaKey && !event.altKey) {
          // a paragraph after the board
          const after = afterFigure(selected.boardPos + state.doc.nodeAt(selected.boardPos)!.nodeSize);
          const tr = state.tr.insert(after, state.schema.nodes.paragraph.create());
          tr.setSelection(TextSelection.create(tr.doc, after + 1)).setMeta(boardsKey, { extra: [] });
          view.dispatch(tr.scrollIntoView());
          return true;
        }
        const arrows: Record<string, [number, number]> = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] };
        if (arrows[event.key] && !event.altKey && !event.metaKey) {
          const step = event.ctrlKey ? NUDGE.fine : event.shiftKey ? NUDGE.coarse : NUDGE.plain;
          const [sx, sy] = arrows[event.key];
          const places = new Map<number, { x?: number; top?: number }>();
          for (const pos of selected.positions) {
            const node = state.doc.nodeAt(pos)!;
            places.set(pos, { x: Math.max(0, Math.min(1 - node.attrs.width, node.attrs.x + sx * step)), top: node.attrs.top + sy * step });
          }
          placeImages(view, places);
          return true;
        }
        if (event.key === "Tab" && !event.ctrlKey && !event.altKey && !event.metaKey) {
          // the board's images in reading order: rows top to bottom, left to right
          const board = state.doc.nodeAt(selected.boardPos)!;
          const order = rowsOf(boardItems(board).map((item) => ({ ...item, pos: selected.boardPos + 1 + offsetOf(board, item.index) }))).flat();
          const at = order.findIndex((item) => item.pos === sel.from);
          const next = order[(at + (event.shiftKey ? order.length - 1 : 1)) % order.length];
          view.dispatch(state.tr.setSelection(NodeSelection.create(state.doc, next.pos)).setMeta(boardsKey, { extra: [] }));
          return true;
        }
        if (event.key === "Escape") {
          view.dispatch(state.tr.setSelection(NodeSelection.create(state.doc, selected.boardPos)).setMeta(boardsKey, { extra: [] }));
          return true;
        }
        if (event.key === "Delete" || event.key === "Backspace") {
          const tr = state.tr;
          for (const pos of [...selected.positions].sort((a, b) => b - a)) tr.delete(pos, pos + state.doc.nodeAt(pos)!.nodeSize);
          tr.setMeta(boardsKey, { extra: [] });
          view.dispatch(tr);
          return true;
        }
        return false;
      },
      // typing does not replace a selected image
      handleTextInput(view) {
        const sel = view.state.selection;
        return sel instanceof NodeSelection && (isBoardImage(sel.node) || isBoard(sel.node));
      },
    },
  });
}

/** Shift/Ctrl+click on a board image: added to (or taken from) the selection; plain click: it alone. */
export function selectBoardImage(view: EditorView, pos: number, add: boolean): void {
  const state = view.state;
  const node = state.doc.nodeAt(pos);
  if (!node || !isBoardImage(node)) return;
  const sel = state.selection;
  const extra = new Set(boardsKey.getState(state)?.extra ?? []);
  if (add && sel instanceof NodeSelection && isBoardImage(sel.node) && state.doc.resolve(sel.from).before() === state.doc.resolve(pos).before()) {
    if (pos === sel.from) return;
    const id = node.attrs.id as string | null;
    if (!id) return;
    if (extra.has(id)) extra.delete(id);
    else extra.add(id);
    // the node selection stays where it is; the others are marked
    view.dispatch(state.tr.setMeta(boardsKey, { extra: [...extra] }));
    return;
  }
  if (sel instanceof NodeSelection && sel.from === pos && extra.size) return;
  view.dispatch(state.tr.setSelection(NodeSelection.create(state.doc, pos)).setMeta(boardsKey, { extra: [] }));
}
