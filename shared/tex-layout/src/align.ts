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
  tag?: Tag;
}

/** Material between rows (\noalign{...}). */
export interface AlignNoalign {
  kind: "noalign";
  apply(v: VList): void;
}

interface PendingRow {
  box: BoxNode; // the row (unset until fin_align)
  cells: { list: Node[]; natural: BoxNode }[];
  tag?: Tag;
}

/**
 * \halign in internal vertical mode, natural width (no "to"/"spread").
 * `tabskips[j]` is the tabskip glue before column j+1 (index 0 = left edge).
 * Rows and noalign material are appended to `v`.
 */
export function halign(items: (AlignRow | AlignNoalign)[], tabskips: GlueSpec[], v: VList, tag?: Tag): void {
  const ncols = Math.max(0, ...items.map((it) => (it.kind === "row" ? it.cells.length : 0)));
  const colW: number[] = new Array(ncols).fill(NULL_FLAG);
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
    const cells: { list: Node[]; natural: BoxNode }[] = [];
    it.cells.forEach((cellList, j) => {
      const natural = hpack(cellList, 0, undefined, it.tag ?? tag);
      if (natural.width > colW[j]) colW[j] = natural.width;
      cells.push({ list: cellList, natural });
      rowList.push(natural);
      rowList.push({ type: GLUE_NODE, spec: tab(j + 1), subtype: 12, tag });
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
    row.cells.forEach((c, j) => {
      const r = hpack(c.list, colW[j], EXACTLY, c.natural.tag);
      r.height = q.height;
      r.depth = q.depth;
      newList.push(r, { type: GLUE_NODE, spec: tabs[j + 1], subtype: 12, tag });
    });
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
}
