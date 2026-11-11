// Figures and tables with their captions (DESIGN.md §11): an image line, an image board or a table is the
// body of a float, LaTeX's figure / table. A figure's caption slot is always there, written or left empty; a table's
// caption is above it (as a new table's is), below it, or not there (the table then has no number, as a tabular
// without \caption). The caption is text (marks, inline formulas) numbered as \caption numbers it: "Figure 1: …"
// under a figure, "Table 1: …" by a table ("图 1: …" / "表 1: …" in a document ctex sets, i.e. one with Chinese text,
// as latex.ts decides); an empty caption still numbers ("Figure 1", the caption package drops the colon). A wrapped
// image's figure floats beside the paragraph after it with its caption under the image, as wrapfigure does.
import { Fragment, Slice, type Node, type NodeSpec, type NodeType, type Schema } from "prosemirror-model";
import { EditorState, NodeSelection, Plugin, PluginKey, Selection, TextSelection, type Command, type Transaction } from "prosemirror-state";
import { ReplaceAroundStep } from "prosemirror-transform";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";
import { bodyLayout, documentSettings, PAPER_SIZES } from "./settings.ts";

// --------------------------------------------------------------------- model --

const id = { id: { default: null as string | null } };

/** A table's rules: Mathcha's ruled grid, or booktabs' three rules (\toprule, \midrule under the first row, \bottomrule). */
export type TableRules = "grid" | "booktabs";

export const figureNodes: Record<string, NodeSpec> = {
  /** A figure: an image line or an image board, and its caption under it. */
  figure: {
    group: "block",
    // the body may be missing for a moment (its image deleted): the figure then goes (figureFixes)
    content: "(paragraph | image_board)? caption",
    isolating: true,
    attrs: id,
    parseDOM: [{ tag: "figure", getAttrs: (dom: HTMLElement | string) => (typeof dom === "string" ? {} : { id: dom.getAttribute("data-id") }) }],
    toDOM: (node) => ["figure", { class: "vtx-figure", ...(node.attrs.id ? { "data-id": node.attrs.id } : {}) }, 0],
  },
  /** A table and its caption, above or below it, or a table without one (LaTeX's table / longtable). */
  table_figure: {
    group: "block",
    // (the table may be missing for a moment, its caption above: the float then goes, figureFixes)
    content: "caption table? | table caption?",
    isolating: true,
    attrs: { ...id, rules: { default: "grid" as TableRules } },
    parseDOM: [
      {
        tag: "div.vtx-table-figure",
        getAttrs: (dom: HTMLElement | string) =>
          typeof dom === "string" ? {} : { id: dom.getAttribute("data-id"), rules: dom.getAttribute("data-rules") === "booktabs" ? "booktabs" : "grid" },
      },
    ],
    toDOM: (node) => ["div", { class: "vtx-table-figure", "data-rules": node.attrs.rules, ...(node.attrs.id ? { "data-id": node.attrs.id } : {}) }, 0],
  },
  /** A figure's or a table's caption: text and inline formulas. */
  caption: {
    content: "inline*",
    defining: true,
    parseDOM: [{ tag: "figcaption" }],
    toDOM: () => ["figcaption", { class: "vtx-caption" }, 0],
  },
};

export const isFigure = (node: Node | null | undefined): boolean => node?.type.name === "figure";
export const isTableFigure = (node: Node | null | undefined): boolean => node?.type.name === "table_figure";
/** A figure or a table with its caption. */
export const isCaptioned = (node: Node | null | undefined): boolean => isFigure(node) || isTableFigure(node);

const isImageNode = (node: Node | null | undefined) => node?.type.name === "image";
const isBlank = (node: Node) => node.isText && !node.text!.trim();

/** A paragraph that holds images (not wrapped) and nothing else: an image line (images.ts). */
function imageLine(node: Node): boolean {
  if (node.type.name !== "paragraph" || !node.childCount) return false;
  let images = 0;
  let other = false;
  node.forEach((child) => {
    if (isImageNode(child) && !child.attrs.wrap) images++;
    else if (!isBlank(child)) other = true;
  });
  return images > 0 && !other;
}

/** A paragraph that holds one wrapped image and nothing else: the body of a figure text wraps. */
export function isWrapBody(node: Node | null | undefined): boolean {
  if (!node || node.type.name !== "paragraph") return false;
  let images = 0;
  let other = false;
  node.forEach((child) => {
    if (isImageNode(child) && child.attrs.wrap) images++;
    else if (!isBlank(child)) other = true;
  });
  return images === 1 && !other;
}

/** What a figure's body may be: an image line, a wrapped image, an image board. */
function figureBody(node: Node | null | undefined): boolean {
  return !!node && (node.type.name === "image_board" ? node.childCount > 0 : imageLine(node) || isWrapBody(node));
}

/** The figure's body (null: none) and caption. */
export function figureParts(figure: Node): { body: Node | null; caption: Node } {
  return { body: figure.childCount > 1 ? figure.firstChild : null, caption: figure.lastChild! };
}

/** Where a table's caption is: above the table, below it, or none (the table has no caption and no number). */
export type CaptionSide = "above" | "below" | "none";

/** The table figure's caption (null: none), its table (null: none) and where the caption is. */
export function tableFigureParts(node: Node): { caption: Node | null; table: Node | null; side: CaptionSide } {
  let caption: Node | null = null;
  let table: Node | null = null;
  let side: CaptionSide = "none";
  for (let i = 0; i < node.childCount; i++) {
    const child = node.child(i);
    if (child.type.name === "caption") {
      caption = child;
      side = table ? "below" : "above";
    } else table = child;
  }
  return { caption, table, side };
}

/** The side a figure is wrapped on (its image's), or null. */
export function figureWrap(figure: Node): "left" | "right" | null {
  const body = figureParts(figure).body;
  if (!body || !isWrapBody(body)) return null;
  let side: "left" | "right" | null = null;
  body.forEach((child) => {
    if (isImageNode(child)) side = child.attrs.wrap;
  });
  return side;
}

/** The ids of the images in a figure's body. */
function bodyImageIds(body: Node | null): string[] {
  const ids: string[] = [];
  body?.forEach((child) => {
    if ((child.type.name === "image" || child.type.name === "board_image") && child.attrs.id) ids.push(child.attrs.id);
  });
  return ids;
}

// ------------------------------------------------------------------ numbers --

const CJK = /[⺀-鿿豈-﫿＀-￯　-〿]/;

/** Whether the document is Chinese for LaTeX: it has CJK text, so the export sets it with ctexart (latex.ts). */
export function documentIsChinese(doc: Node): boolean {
  let found = false;
  doc.descendants((node) => {
    if (found) return false;
    if (node.isText) found = CJK.test(node.text!);
    else if (node.type.name === "math_inline" || node.type.name === "math_display") found = CJK.test(node.attrs.latex ?? "");
    else if (node.type.name === "footnote") found = CJK.test(node.attrs.text ?? "");
    return !found;
  });
  return found;
}

export interface CaptionLabel {
  kind: "figure" | "table";
  number: number;
  /** "Figure 3", "表 2" */
  name: string;
}

/** Each figure's and table's number and name (keyed by its node), counted in document order as LaTeX counts them. */
export function captionLabels(doc: Node, chinese = documentIsChinese(doc)): Map<Node, CaptionLabel> {
  const labels = new Map<Node, CaptionLabel>();
  let figures = 0;
  let tables = 0;
  doc.descendants((node) => {
    if (isFigure(node)) {
      figures++;
      labels.set(node, { kind: "figure", number: figures, name: `${chinese ? "图" : "Figure"} ${figures}` });
      return false;
    }
    if (isTableFigure(node)) {
      // (a table without a caption has no number, as a tabular without \caption)
      if (!tableFigureParts(node).caption) return false;
      tables++;
      labels.set(node, { kind: "table", number: tables, name: `${chinese ? "表" : "Table"} ${tables}` });
      return false;
    }
    return node.isBlock;
  });
  return labels;
}

/** What precedes a caption's text: "Figure 1: " (the name alone when the caption is empty, as LaTeX's caption package). */
export function captionPrefix(label: CaptionLabel, caption: Node): string {
  return caption.content.size ? `${label.name}: ` : label.name;
}

/** A caption label at the start of a text read from a file ("Figure 2:", "Fig. 2.", "图 2", "Table:" …): its length. */
export function captionLabelLength(text: string, kind: "figure" | "table"): number {
  const pattern =
    kind === "figure"
      ? /^\s*(?:Figure|Fig\.|图)\s*\d+(?:[.-]\d+)?\s*(?:[:：.]\s*|\s+|$)/i
      : /^\s*(?:Table|表)(?:\s*\d+(?:[.-]\d+)?)?\s*(?:[:：.]\s*|\s+(?=\S)|$)/i;
  const match = pattern.exec(text);
  if (!match) return 0;
  // "Table" alone is a word, not a label: it needs a number or a colon
  if (kind === "table" && !/\d|[:：]/.test(match[0])) return 0;
  return match[0].length;
}

// ------------------------------------------------------------- long tables --

/**
 * A table's height in LaTeX, estimated (pt): a row is as high as its most lines (\baselineskip each, the class's:
 * bodyLayout), a cell's lines its paragraphs, its line breaks and, in a p{} column (a cell with a display formula or several
 * paragraphs, latex.ts), its text over the column's width; plus the rules.
 */
export function tableHeightPt(doc: Node, table: Node): number {
  const settings = documentSettings(doc);
  const body = bodyLayout(doc);
  const size = body.fontSize;
  const baseline = size * body.lineHeight;
  const paper = PAPER_SIZES[settings.paper];
  const lineWidth = ((paper.width - settings.marginLeft - settings.marginRight) * 72.27) / 25.4;
  let columns = 0;
  table.firstChild?.forEach((cell) => (columns += cell.attrs.colspan ?? 1));
  const columnWidth = lineWidth / Math.max(1, columns);
  let height = 0;
  table.forEach((row) => {
    let lines = 1;
    row.forEach((cell) => {
      let cellLines = 0;
      const block = cell.childCount > 1 || cell.firstChild?.type.name === "math_display";
      cell.forEach((child) => {
        if (child.type.name === "math_display") {
          cellLines += 2;
          return;
        }
        let breaks = 0;
        let chars = 0;
        child.forEach((inline) => {
          if (inline.type.name === "hard_break") breaks++;
          else if (inline.isText) for (const ch of inline.text!) chars += CJK.test(ch) ? 2 : 1;
          else chars += 4;
        });
        const perLine = Math.max(4, Math.floor((columnWidth * (cell.attrs.colspan ?? 1)) / (size * 0.5)));
        cellLines += (block ? Math.max(1, Math.ceil(chars / perLine)) : 1) + breaks;
      });
      lines = Math.max(lines, Math.ceil(cellLines / Math.max(1, cell.attrs.rowspan ?? 1)));
    });
    height += lines * baseline + 0.4;
  });
  return height;
}

/** The text area's height (pt). */
export function textHeightPt(doc: Node): number {
  const settings = documentSettings(doc);
  const paper = PAPER_SIZES[settings.paper];
  return ((paper.height - settings.marginTop - settings.marginBottom) * 72.27) / 25.4;
}

/**
 * Whether a table is written as a longtable: when it cannot be on one page, caption and all (a table float cannot
 * break; longtable can, and LaTeX then continues it on the next pages as the page view does).
 */
export function isLongTable(doc: Node, table: Node): boolean {
  return tableHeightPt(doc, table) + 3 * documentSettings(doc).fontSize * 1.2 > textHeightPt(doc);
}

// --------------------------------------------------------------- the rules --

/** What the normalising pass found to do at a place (applied from the document's end, so places stay valid). */
type Fix =
  | { kind: "wrap"; pos: number; node: Node; table: boolean; caption: Fragment | null; side?: CaptionSide }
  | { kind: "delete"; pos: number; node: Node }
  | { kind: "unwrap"; pos: number; node: Node }
  | { kind: "unwrapImage"; pos: number; bodyPos: number; body: Node }
  | { kind: "lift"; pos: number; node: Node; caption: Fragment | null };

/** Captions with text by the body they captioned: "i:<image id>" (a figure of one image), "b:<board id>"; and every
 *  table's caption with where it was, by the table's id. */
function captionsByBody(doc: Node): { captions: Map<string, Fragment>; tables: Map<string, { content: Fragment; side: CaptionSide }> } {
  const captions = new Map<string, Fragment>();
  const tables = new Map<string, { content: Fragment; side: CaptionSide }>();
  doc.descendants((node) => {
    if (isFigure(node)) {
      const { body, caption } = figureParts(node);
      const ids = bodyImageIds(body);
      if (ids.length === 1 && caption.content.size) captions.set(`i:${ids[0]}`, caption.content);
      if (body?.type.name === "image_board" && body.attrs.id && caption.content.size) captions.set(`b:${body.attrs.id}`, caption.content);
      return false;
    }
    if (isTableFigure(node)) {
      const { caption, table, side } = tableFigureParts(node);
      if (table?.attrs.id) tables.set(table.attrs.id, { content: caption?.content ?? Fragment.empty, side });
      return false;
    }
    return node.isBlock;
  });
  return { captions, tables };
}

/**
 * What keeps every figure and table captioned (an appended transaction, and once when a document is opened):
 * - an image line, an image board or a table that is not in its float yet gets one, its caption empty (a table's above
 *   it) — or, when it was moved out of a figure or table of its own (a drag, a wrap, an alignment), the caption it had
 *   there (a table's where it was, or none);
 * - a figure whose image is gone goes, its caption with it (so does a table float without a table); a figure whose
 *   line got text is a paragraph again (its caption, if written, a paragraph after it);
 * - a wrapped image at the start of a body paragraph (an older document, a wrap gesture, a file read) becomes the
 *   wrapped figure before that paragraph; a wrapped figure that has no body paragraph after it stops wrapping.
 * An image line that is a list item's first paragraph stays as it is (a list item starts with a paragraph).
 */
export function figureFixes(state: EditorState, previous: EditorState | null = null): Transaction | null {
  const schema = state.schema;
  const { figure, table_figure: tableFigure, caption: captionType, paragraph } = schema.nodes;
  if (!figure || !tableFigure || !captionType) return null;
  const fixes: Fix[] = [];
  let carried: ReturnType<typeof captionsByBody> | null = null;
  const carry = (key: string | null): Fragment | null => {
    if (!key || !previous) return null;
    carried ??= captionsByBody(previous.doc);
    return carried.captions.get(key) ?? null;
  };
  // a table moved out of its float: its caption, where it was (above, below, none)
  const carryTable = (id: string | null): { content: Fragment; side: CaptionSide } | null => {
    if (!id || !previous) return null;
    carried ??= captionsByBody(previous.doc);
    return carried.tables.get(id) ?? null;
  };
  const canHold = (parent: Node, index: number, type: NodeType) => parent.canReplaceWith(index, index + 1, type);
  state.doc.descendants((node, pos, parent, index) => {
    if (!parent) return false;
    const name = node.type.name;
    if (name === "figure") {
      const { body } = figureParts(node);
      if (!body || !figureBody(body)) {
        // its line got text: a paragraph again; its images gone: the figure goes
        fixes.push(body && body.type.name === "paragraph" && countImages(body) ? { kind: "unwrap", pos, node } : { kind: "delete", pos, node });
        return false;
      }
      // a wrapped figure wraps the body paragraph after it (wrapfig wraps paragraphs only)
      if (isWrapBody(body)) {
        const next = index + 1 < parent.childCount ? parent.child(index + 1) : null;
        const wraps = parent === state.doc && !!next && next.type.name === "paragraph" && !imageLine(next) && next.textContent.trim().length > 0;
        if (!wraps) fixes.push({ kind: "unwrapImage", pos: pos + 1, bodyPos: pos + 1, body });
      }
      return false;
    }
    if (name === "table_figure") {
      if (!tableFigureParts(node).table) fixes.push({ kind: "delete", pos, node });
      return false;
    }
    if (name === "table") {
      if (canHold(parent, index, tableFigure)) {
        const had = carryTable(node.attrs.id ?? null);
        fixes.push({ kind: "wrap", pos, node, table: true, caption: had?.content ?? null, side: had?.side ?? "above" });
      }
      return false;
    }
    if (name === "image_board") {
      if (node.childCount && canHold(parent, index, figure)) {
        const ids = bodyImageIds(node);
        fixes.push({ kind: "wrap", pos, node, table: false, caption: carry(node.attrs.id ? `b:${node.attrs.id}` : null) ?? (ids.length === 1 ? carry(`i:${ids[0]}`) : null) });
      }
      return false;
    }
    if (name === "paragraph") {
      if (imageLine(node)) {
        if (canHold(parent, index, figure)) {
          const ids = bodyImageIds(node);
          fixes.push({ kind: "wrap", pos, node, table: false, caption: ids.length === 1 ? carry(`i:${ids[0]}`) : null });
        }
      } else if (parent === state.doc && isImageNode(node.firstChild) && node.firstChild!.attrs.wrap && node.textContent.trim()) {
        // a wrapped image opening a body paragraph: its figure before the paragraph
        const image = node.firstChild!;
        fixes.push({ kind: "lift", pos, node, caption: carry(image.attrs.id ? `i:${image.attrs.id}` : null) });
      }
      return false;
    }
    // into lists, quotes and the like; not into textblocks, tables' cells or atoms
    return !node.isTextblock && !node.isAtom;
  });
  if (!fixes.length) return null;
  const tr = state.tr;
  const makeCaption = (content: Fragment | null) => captionType.create(null, content ?? Fragment.empty);
  for (const fix of fixes.reverse()) {
    if (fix.kind === "delete") tr.delete(fix.pos, fix.pos + fix.node.nodeSize);
    else if (fix.kind === "wrap") {
      const end = fix.pos + fix.node.nodeSize;
      const caption = makeCaption(fix.caption);
      // a table's caption above it (the gap, its table, after the caption), below it (the table before it) or none
      const side = fix.table ? (fix.side ?? "above") : "below";
      const wrapper = fix.table ? tableFigure.create(null, side === "none" ? Fragment.empty : caption) : figure.create(null, caption);
      // not a "structure" step: its undo takes the caption away with the float
      tr.step(new ReplaceAroundStep(fix.pos, end, fix.pos, end, new Slice(Fragment.from(wrapper), 0, 0), side === "above" ? 1 + caption.nodeSize : 1, false));
    } else if (fix.kind === "unwrap") {
      // the line with its text back as a paragraph; a written caption after it as a paragraph of its own
      const { body, caption } = figureParts(fix.node);
      const end = fix.pos + fix.node.nodeSize;
      const after = caption.content.size ? Fragment.from(paragraph.create({ align: "center" }, caption.content)) : Fragment.empty;
      if (body) tr.step(new ReplaceAroundStep(fix.pos, end, fix.pos + 1, fix.pos + 1 + body.nodeSize, new Slice(after, 0, 0), 0, false));
      else tr.replaceWith(fix.pos, end, after);
    } else if (fix.kind === "unwrapImage") {
      fix.body.forEach((child, offset) => {
        if (isImageNode(child) && child.attrs.wrap) tr.setNodeMarkup(fix.bodyPos + 1 + offset, undefined, { ...child.attrs, wrap: null });
      });
      tr.setNodeMarkup(fix.bodyPos, undefined, { ...fix.body.attrs, align: "center" });
    } else if (fix.kind === "lift") {
      const image = fix.node.firstChild!;
      tr.delete(fix.pos + 1, fix.pos + 1 + image.nodeSize);
      // the space that followed the image on the line is not text
      const rest = tr.doc.nodeAt(fix.pos);
      const first = rest?.firstChild;
      if (first?.isText && /^\s+/.test(first.text!)) tr.delete(fix.pos + 1, fix.pos + 1 + /^\s+/.exec(first.text!)![0].length);
      tr.insert(fix.pos, figure.create(null, [paragraph.create(null, image), makeCaption(fix.caption)]));
    }
  }
  // a selected image (or table) stays selected where it went
  const sel = state.selection;
  if (sel instanceof NodeSelection && sel.node.attrs.id && (isImageNode(sel.node) || sel.node.type.name === "board_image")) {
    let found = -1;
    tr.doc.descendants((node, pos) => {
      if (found < 0 && node.type === sel.node.type && node.attrs.id === sel.node.attrs.id) found = pos;
      return found < 0;
    });
    if (found >= 0) tr.setSelection(NodeSelection.create(tr.doc, found));
  }
  return tr;
}

function countImages(node: Node): number {
  let count = 0;
  node.descendants((child) => {
    if (child.type.name === "image" || child.type.name === "board_image") count++;
    return true;
  });
  return count;
}

// -------------------------------------------------------------- positions --

/** The figure or table float a position is in (the innermost), with its position, or null. */
export function captionedAt(doc: Node, pos: number): { node: Node; pos: number; depth: number } | null {
  const $pos = doc.resolve(pos);
  for (let d = $pos.depth; d > 0; d--) {
    const node = $pos.node(d);
    if (isCaptioned(node)) return { node, pos: $pos.before(d), depth: d };
  }
  return null;
}

/** The depth to insert a block after: the float's around a figure's line or a table (not between its parts). */
export function outsideCaptioned(doc: Node, pos: number, depth: number): number {
  const $pos = doc.resolve(pos);
  if (depth > 1 && isCaptioned($pos.node(depth - 1))) return depth - 1;
  return depth;
}

/** The table float the selection is in (in its table or its caption, or the float selected), or null. */
function tableFigureAt(state: EditorState): { node: Node; pos: number } | null {
  const sel = state.selection;
  if (sel instanceof NodeSelection && isTableFigure(sel.node)) return { node: sel.node, pos: sel.from };
  const found = captionedAt(state.doc, sel.from);
  return found && isTableFigure(found.node) ? { node: found.node, pos: found.pos } : null;
}

/** Where the caption is of the table the selection is in (above, below, none), or null outside a table. */
export function tableCaptionSide(state: EditorState): CaptionSide | null {
  const at = tableFigureAt(state);
  return at ? tableFigureParts(at.node).side : null;
}

/**
 * The caption of the table the selection is in put above the table, below it, or taken away (the table then has no
 * number). A caption moved keeps its text; one put back is empty. An empty caption gets the caret (its hint shows).
 */
export function setTableCaption(side: CaptionSide): Command {
  return (state, dispatch) => {
    const at = tableFigureAt(state);
    if (!at) return false;
    const parts = tableFigureParts(at.node);
    if (!parts.table || parts.side === side) return false;
    if (!dispatch) return true;
    const tr = state.tr;
    let captionFrom = -1;
    at.node.forEach((child, offset) => {
      if (child.type.name === "caption") captionFrom = at.pos + 1 + offset;
    });
    const sel = state.selection;
    // the caret's offset in the caption it is in (-1: elsewhere)
    const inCaption = parts.caption && sel instanceof TextSelection && sel.from > captionFrom && sel.to < captionFrom + parts.caption.nodeSize ? sel.from - captionFrom - 1 : -1;
    if (parts.caption) tr.delete(captionFrom, captionFrom + parts.caption.nodeSize);
    if (side !== "none") {
      const caption = parts.caption ?? state.schema.nodes.caption.create();
      const figure = tr.doc.nodeAt(at.pos)!;
      const place = side === "above" ? at.pos + 1 : at.pos + 1 + figure.firstChild!.nodeSize;
      tr.insert(place, caption);
      // the caret into it when it was there, or when it is empty (its hint shows: it is to be written)
      if (inCaption >= 0 || !caption.content.size) tr.setSelection(TextSelection.create(tr.doc, place + 1 + Math.max(0, inCaption)));
    } else if (inCaption >= 0) {
      // the caption gone with the caret: the caret into the table
      const near = Selection.findFrom(tr.doc.resolve(at.pos + 1), 1, true);
      if (near) tr.setSelection(near);
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

// ------------------------------------------------------------------- view --

export interface CaptionLabels {
  /** the hint in an empty caption while the caret is in it */
  caption: string;
}

export const DEFAULT_CAPTION_LABELS: CaptionLabels = { caption: "Caption" };

export const figuresKey = new PluginKey<DecorationSet>("visualtex-document-figures");

function labelWidget(text: string, hint: string): () => HTMLElement {
  return () => {
    const span = document.createElement("span");
    span.className = "vtx-caption-label";
    span.contentEditable = "false";
    span.textContent = text;
    // the hint an empty caption shows while the caret is in it (document.css)
    span.setAttribute("data-hint", hint);
    return span;
  };
}

/** The captions' numbers (a widget before the text), empty captions marked, wrapped figures placed. */
function captionDecorations(doc: Node, hint: string): DecorationSet {
  const decorations: Decoration[] = [];
  const labels = captionLabels(doc);
  doc.descendants((node, pos) => {
    if (!isCaptioned(node)) return node.isBlock;
    const label = labels.get(node);
    let captionPos = pos + 1;
    let caption = node.firstChild!;
    if (isTableFigure(node)) {
      // above or below its table, or none (nothing to number)
      const parts = tableFigureParts(node);
      if (!parts.caption) return false;
      caption = parts.caption;
      if (parts.side === "below") captionPos = pos + node.nodeSize - 1 - caption.nodeSize;
    }
    if (isFigure(node)) {
      caption = node.lastChild!;
      captionPos = pos + node.nodeSize - 1 - caption.nodeSize;
      const side = figureWrap(node);
      if (side) {
        let width = 0.4;
        figureParts(node).body!.forEach((child) => {
          if (isImageNode(child)) width = child.attrs.width;
        });
        decorations.push(Decoration.node(pos, pos + node.nodeSize, { class: `is-wrap-${side}`, style: `width:${(width * 100).toFixed(2)}%` }));
      }
    }
    if (label) {
      const text = captionPrefix(label, caption);
      decorations.push(Decoration.widget(captionPos + 1, labelWidget(text, hint), { side: -1, key: `caption-label:${text}:${hint}`, ignoreSelection: true }));
    }
    decorations.push(Decoration.node(captionPos, captionPos + caption.nodeSize, { class: caption.content.size ? "has-text" : "is-empty" }));
    return false;
  });
  return DecorationSet.create(doc, decorations);
}

/** The caption the caret is in when it is empty (its hint shows): a decoration, or null. */
function activeCaption(state: EditorState): Decoration | null {
  const sel = state.selection;
  if (!(sel instanceof TextSelection) || !sel.empty) return null;
  const $pos = sel.$from;
  if ($pos.parent.type.name !== "caption" || $pos.parent.content.size) return null;
  return Decoration.node($pos.before(), $pos.after(), { class: "is-active" });
}

/** The end (or start) of the textblock nearest a position, in a direction: a caret there. */
function caretNear(doc: Node, pos: number, dir: -1 | 1): Selection | null {
  return Selection.findFrom(doc.resolve(pos), dir, true);
}

/** The last image of a figure (to select it from its caption): its position, or -1. */
function lastImageOf(figure: Node, figurePos: number): number {
  const body = figureParts(figure).body;
  if (!body) return -1;
  let found = -1;
  body.forEach((child, offset) => {
    if (child.type.name === "image" || child.type.name === "board_image") found = figurePos + 2 + offset;
  });
  return found;
}

function captionKeys(view: EditorView, event: KeyboardEvent): boolean {
  if (event.ctrlKey || event.metaKey || event.altKey || event.isComposing) return false;
  const key = event.key;
  if (key !== "Enter" && key !== "Backspace" && key !== "Delete") return false;
  const state = view.state;
  const sel = state.selection;
  if (!(sel instanceof TextSelection) || !sel.empty) return false;
  const $pos = sel.$from;
  const parent = $pos.parent;
  const depth = $pos.depth;
  const doc = state.doc;
  const dispatch = (tr: Transaction) => {
    view.dispatch(tr.scrollIntoView());
    return true;
  };
  // in a caption
  if (parent.type.name === "caption" && depth > 1) {
    const wrapper = $pos.node(depth - 1);
    const wrapperPos = $pos.before(depth - 1);
    const wrapperEnd = wrapperPos + wrapper.nodeSize;
    const tableSide = isTableFigure(wrapper) ? tableFigureParts(wrapper) : null;
    if (key === "Enter") {
      if (event.shiftKey) return false;
      // a table's caption above it: on into its first cell; a figure's, or a table's below it: a new paragraph after
      // the float
      if (tableSide?.table && tableSide.side === "above") {
        const near = caretNear(doc, $pos.after(), 1);
        if (near) return dispatch(state.tr.setSelection(near));
      }
      const tr = state.tr.insert(wrapperEnd, state.schema.nodes.paragraph.create());
      return dispatch(tr.setSelection(TextSelection.create(tr.doc, wrapperEnd + 1)));
    }
    if (key === "Backspace" && $pos.parentOffset === 0) {
      // the caption stays: a figure's image is selected; under a table the caret goes into the table's last cell,
      // above one to the text before
      if (isFigure(wrapper)) {
        const image = lastImageOf(wrapper, wrapperPos);
        if (image >= 0) return dispatch(state.tr.setSelection(NodeSelection.create(doc, image)));
        return true;
      }
      if (tableSide?.side === "below") {
        const inTable = caretNear(doc, $pos.before(), -1);
        if (inTable) dispatch(state.tr.setSelection(inTable));
        return true;
      }
      const before = wrapperPos > 0 ? caretNear(doc, wrapperPos, -1) : null;
      if (before) dispatch(state.tr.setSelection(before));
      return true;
    }
    // Delete at its end: nothing joins into the caption or out of it
    if (key === "Delete" && $pos.parentOffset === parent.content.size) return true;
    return false;
  }
  // in a figure's line (the caret beside its images)
  if (depth > 1 && isFigure($pos.node(depth - 1)) && $pos.index(depth - 1) === 0) {
    const figurePos = $pos.before(depth - 1);
    if (key === "Enter") {
      if (event.shiftKey) return false;
      // at the line's start a paragraph before the figure; else one after it (as on a selected image)
      const figure = $pos.node(depth - 1);
      const at = $pos.parentOffset === 0 ? figurePos : figurePos + figure.nodeSize;
      const tr = state.tr.insert(at, state.schema.nodes.paragraph.create());
      if (at > figurePos) tr.setSelection(TextSelection.create(tr.doc, at + 1));
      return dispatch(tr);
    }
    if (key === "Delete" && $pos.parentOffset === parent.content.size) {
      return dispatch(state.tr.setSelection(TextSelection.create(doc, $pos.after() + 1)));
    }
    if (key === "Backspace" && $pos.parentOffset === 0) {
      const $figure = doc.resolve(figurePos);
      const before = $figure.nodeBefore;
      if (before && before.type.name === "paragraph" && !before.content.size) return dispatch(state.tr.delete(figurePos - before.nodeSize, figurePos));
      const near = figurePos > 0 ? caretNear(doc, figurePos, -1) : null;
      if (near) dispatch(state.tr.setSelection(near));
      return true;
    }
    return false;
  }
  if (!parent.isTextblock || depth < 1) return false;
  // next to a figure or table: the caret goes into it, never a join (an empty paragraph goes)
  if (key === "Backspace" && $pos.parentOffset === 0) {
    const $block = doc.resolve($pos.before());
    const before = $block.nodeBefore;
    if (!isCaptioned(before)) return false;
    const end = $pos.before();
    if (parent.type.name === "paragraph" && !parent.content.size) {
      const tr = state.tr.delete(end, $pos.after());
      const near = caretNear(tr.doc, end, -1);
      if (near) tr.setSelection(near);
      return dispatch(tr);
    }
    const near = caretNear(doc, end, -1);
    if (near) dispatch(state.tr.setSelection(near));
    return true;
  }
  if (key === "Delete" && $pos.parentOffset === parent.content.size) {
    const $block = doc.resolve($pos.after());
    const after = $block.nodeAfter;
    if (!isCaptioned(after)) return false;
    const start = $pos.after();
    if (parent.type.name === "paragraph" && !parent.content.size) {
      const tr = state.tr.delete($pos.before(), start);
      const near = caretNear(tr.doc, $pos.before(), 1);
      if (near) tr.setSelection(near);
      return dispatch(tr);
    }
    const near = caretNear(doc, start, 1);
    if (near) dispatch(state.tr.setSelection(near));
    return true;
  }
  return false;
}

export interface FiguresPluginOptions {
  labels: () => CaptionLabels;
}

export function figuresPlugin(options: FiguresPluginOptions): Plugin<DecorationSet> {
  return new Plugin<DecorationSet>({
    key: figuresKey,
    state: {
      init: (_config, state) => captionDecorations(state.doc, options.labels().caption),
      apply: (tr, value, _old, state) => (tr.docChanged ? captionDecorations(state.doc, options.labels().caption) : value),
    },
    appendTransaction(trs, old, state) {
      if (!trs.some((tr) => tr.docChanged)) return null;
      return figureFixes(state, old);
    },
    props: {
      decorations(state) {
        const set = figuresKey.getState(state);
        const active = activeCaption(state);
        return active && set ? set.add(state.doc, [active]) : set;
      },
      handleKeyDown: captionKeys,
      handleDOMEvents: {
        // a press beside a caption (it is as wide as its text, centred): the caret into it, at the nearer end
        mousedown(view, event) {
          const target = event.target as HTMLElement | null;
          if (event.button !== 0 || !target?.matches?.("figure.vtx-figure, .vtx-table-figure")) return false;
          const caption = target.querySelector<HTMLElement>(":scope > .vtx-caption");
          const rect = caption?.getBoundingClientRect();
          if (!caption || !rect || event.clientY < rect.top || event.clientY > rect.bottom) return false;
          const start = view.posAtDOM(caption, 0);
          const node = view.state.doc.resolve(start).parent;
          if (node.type.name !== "caption") return false;
          event.preventDefault();
          const at = event.clientX < (rect.left + rect.right) / 2 ? start : start + node.content.size;
          view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at)));
          view.focus();
          return true;
        },
      },
      // a selected image that is its figure's only one, or its selected board, is copied (and cut) with its caption:
      // the whole figure
      transformCopied(slice, view) {
        const sel = view.state.selection;
        if (!(sel instanceof NodeSelection) || slice.content.childCount !== 1) return slice;
        const board = sel.node.type.name === "image_board";
        if (!isImageNode(sel.node) && !board) return slice;
        const found = captionedAt(view.state.doc, sel.from);
        if (!found || !isFigure(found.node) || (!board && countImages(found.node) !== 1)) return slice;
        return new Slice(Fragment.from(found.node.type.create(null, found.node.content)), 0, 0);
      },
    },
  });
}

/** A figure's or table's caption as plain inline content for a schema (text read from a file). */
export function captionNode(schema: Schema, content: Fragment | Node[] | null = null): Node {
  return schema.nodes.caption.create(null, content);
}
