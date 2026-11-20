// Vertical lists and \halign (tex.web parts 37 and 679), for the boxes that
// LaTeX builds with alignments: arrays, matrices, cases, \overbrace, ...
import { NULL_FLAG } from "./arith.ts";
import { GLUE_NODE, HLIST_NODE, NORMAL, RULE_NODE, type BoxNode, type GlueSpec, type Node, type RuleNode, type Tag } from "./nodes.ts";
import { EXACTLY, hpack } from "./pack.ts";

export const IGNORE_DEPTH = -65536000; // -1000pt

export interface VParams {
  baselineskip: GlueSpec;
  lineskip: GlueSpec;
  lineskiplimit: number;
}

export function glue(width: number, stretch = 0, stretchOrder = 0, shrink = 0, shrinkOrder = 0): GlueSpec {
  return { width, stretch, stretchOrder, shrink, shrinkOrder };
}

/** An internal vertical list being built, with \prevdepth. */
export class VList {
  readonly items: Node[] = [];
  prevDepth = IGNORE_DEPTH;
  params: VParams;
  tag: Tag;

  constructor(params: VParams, tag?: Tag) {
    this.params = params;
    this.tag = tag;
  }

  /** append_to_vlist (tex.web 679) */
  appendBox(b: BoxNode): void {
    if (this.prevDepth > IGNORE_DEPTH) {
      const d = this.params.baselineskip.width - this.prevDepth - b.height;
      const spec = d < this.params.lineskiplimit ? { ...this.params.lineskip } : { ...this.params.baselineskip, width: d };
      this.items.push({ type: GLUE_NODE, spec, subtype: NORMAL, tag: this.tag });
    }
    this.items.push(b);
    this.prevDepth = b.depth;
  }

  /** \hrule in vertical mode resets \prevdepth. */
  appendRule(r: RuleNode): void {
    this.items.push(r);
    this.prevDepth = IGNORE_DEPTH;
  }

  /** kerns and \vskip glue leave \prevdepth alone */
  append(n: Node): void {
    this.items.push(n);
  }

  noInterlineSkip(): void {
    this.prevDepth = IGNORE_DEPTH;
  }
}

/** One alignment row: the finished horizontal list of each cell (template applied). */
export interface AlignRow {
  kind: "row";
  cells: Node[][];
  /** the columns each cell spans (\span, \multispan); 1 if absent */
  spans?: number[];
  tag?: Tag;
}

/** Material between rows (\noalign{...}). */
export interface AlignNoalign {
  kind: "noalign";
  apply(v: VList): void;
}

interface PendingRow {
  box: BoxNode; // the row (unset until fin_align)
  cells: { list: Node[]; natural: BoxNode; column: number; span: number }[];
  tag?: Tag;
}

/**
 * \halign in internal vertical mode, natural width (no "to"/"spread").
 * `tabskips[j]` is the tabskip glue before column j+1 (index 0 = left edge).
 * Rows and noalign material are appended to `v`; returns the rows' boxes.
 */
export function halign(items: (AlignRow | AlignNoalign)[], tabskips: GlueSpec[], v: VList, tag?: Tag): BoxNode[] {
  const spanOf = (it: AlignRow, j: number) => Math.max(1, it.spans?.[j] ?? 1);
  const ncols = Math.max(0, ...items.map((it) => (it.kind === "row" ? it.cells.reduce((n, _c, j) => n + spanOf(it, j), 0) : 0)));
  const colW: number[] = new Array(ncols).fill(NULL_FLAG);
  // span nodes (tex.web 798): the widest spanned box from column `from` to column `to`
  let spanW: { from: number; to: number; width: number }[] = [];
  const tab = (j: number) => tabskips[j] ?? tabskips[tabskips.length - 1] ?? glue(0);
  const rows: PendingRow[] = [];
  const startIndex = v.items.length;
  for (const it of items) {
    if (it.kind === "noalign") {
      it.apply(v);
      continue;
    }
    // fin_col for each cell, fin_row for the row
    const rowList: Node[] = [{ type: GLUE_NODE, spec: tab(0), subtype: 12, tag }];
    const cells: PendingRow["cells"] = [];
    let column = 0;
    it.cells.forEach((cellList, j) => {
      const natural = hpack(cellList, 0, undefined, it.tag ?? tag);
      const span = spanOf(it, j);
      if (span === 1) {
        if (natural.width > colW[column]) colW[column] = natural.width;
      } else spanW.push({ from: column, to: column + span - 1, width: natural.width });
      cells.push({ list: cellList, natural, column, span });
      column += span;
      rowList.push(natural);
      rowList.push({ type: GLUE_NODE, spec: tab(column), subtype: 12, tag });
    });
    const rowBox = hpack(rowList, 0, undefined, it.tag ?? tag);
    v.appendBox(rowBox);
    rows.push({ box: rowBox, cells, tag: it.tag });
  }
  // fin_align: final column widths; a column nobody used gets width 0 and its
  // following tabskip is zeroed.
  const tabs: GlueSpec[] = [];
  for (let j = 0; j <= ncols; j++) tabs.push(tab(j));
  for (let j = 0; j < ncols; j++) {
    if (colW[j] === NULL_FLAG) {
      colW[j] = 0;
      tabs[j + 1] = glue(0);
    }
    // tex.web 801-803: a span from this column needs what this column and its
    // tabskip leave of its width from the next column on (the last column
    // spanned takes the rest)
    const rest: typeof spanW = [];
    for (const sp of spanW) {
      if (sp.from !== j) {
        rest.push(sp);
        continue;
      }
      const width = sp.width - colW[j] - tabs[j + 1].width;
      if (sp.to === j + 1) {
        if (width > colW[j + 1]) colW[j + 1] = width;
      } else rest.push({ from: j + 1, to: sp.to, width });
    }
    spanW = rest;
  }
  const proto: Node[] = [{ type: GLUE_NODE, spec: tabs[0], subtype: 12 }];
  for (let j = 0; j < ncols; j++) {
    const b: BoxNode = { type: HLIST_NODE, width: colW[j], height: 0, depth: 0, shift: 0, list: [], glueSet: 0, glueSign: 0, glueOrder: 0 };
    proto.push(b, { type: GLUE_NODE, spec: tabs[j + 1], subtype: 12 });
  }
  const p = hpack(proto);
  // Set the glue in all the unset boxes (rows) of the list
  for (const row of rows) {
    const q = row.box;
    const newList: Node[] = [{ type: GLUE_NODE, spec: tabs[0], subtype: 12, tag }];
    for (const c of row.cells) {
      // a spanned cell is as wide as its columns and the tabskips between them
      let w = colW[c.column];
      for (let k = 1; k < c.span; k++) w += tabs[c.column + k].width + colW[c.column + k];
      const r = hpack(c.list, w, EXACTLY, c.natural.tag);
      r.height = q.height;
      r.depth = q.depth;
      newList.push(r, { type: GLUE_NODE, spec: tabs[c.column + c.span], subtype: 12, tag });
    }
    q.list = newList;
    q.width = p.width;
    q.glueOrder = p.glueOrder;
    q.glueSign = p.glueSign;
    q.glueSet = p.glueSet;
    q.shift = 0;
  }
  // Rules with running dimensions extend to the boundaries of the alignment
  for (let k = startIndex; k < v.items.length; k++) {
    const n = v.items[k];
    if (n.type === RULE_NODE) {
      if (n.width === NULL_FLAG) n.width = p.width;
    }
  }
  // the rows' boxes, in the order of the rows
  return rows.map((row) => row.box);
}
