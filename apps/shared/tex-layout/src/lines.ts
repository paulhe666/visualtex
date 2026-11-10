// A formula set on several lines, as TeX breaks an in-line formula in a
// paragraph (tex.web 761, 813-890): only after a relation or a binary
// operator of the outermost list (the penalties mlist_to_hlist puts there),
// the operator staying at the end of its line and the glue after it dropped.
//
// The formula is typeset once on one line; it is then cut into pieces at
// those places and each piece is moved where its line puts it. Who decides
// where the lines are is the caller's (VisualTeX Anywhere: the host field
// lays out space held for each piece and reports where each one went), so
// this module only cuts and moves; `fillLines` is the plain first-fit rule
// for callers that know the lines' widths.
import { INF_PENALTY } from "./mlist.ts";
import { CHAR_NODE, GLUE_NODE, HLIST_NODE, KERN_NODE, MATH_NODE, PENALTY_NODE, RULE_NODE, VLIST_NODE, type BoxNode, type Node, type Tag } from "./nodes.ts";
import type { ShipResult } from "./shipout.ts";

const EM = 10 * 65536;

/** A piece of the formula between two break places (sp, on the one-line layout). */
export interface FormulaPiece {
  /** where the piece starts and ends on the one-line formula (its end: where the next piece starts) */
  start: number;
  end: number;
  /** its material's height above and depth below the baseline */
  height: number;
  depth: number;
}

/** Where a piece goes: its start, from the first piece's start on the first line's baseline (em; y down). */
export interface PiecePlace {
  x: number;
  y: number;
}

/**
 * Places the pieces of a formula, given their widths in em: one place per
 * piece (a missing or unusable one continues the line of the piece before).
 */
export type PiecePlacer = (widths: number[]) => (PiecePlace | null | undefined)[];

/** The formula's lines as drawn (sp): every piece with its place, and where the first one starts. */
export interface FormulaLines {
  /** the first piece's start in the drawn box (x from its left edge; y = 0 is the first line's baseline) */
  originX: number;
  pieces: (FormulaPiece & { width: number; x: number; y: number })[];
}

function nodeWidth(n: Node): number {
  switch (n.type) {
    case CHAR_NODE:
      return n.font.charWidth(n.font.charInfo(n.char));
    case HLIST_NODE:
    case VLIST_NODE:
    case RULE_NODE:
    case KERN_NODE:
    case MATH_NODE:
      return n.width;
    case GLUE_NODE:
      return n.spec.width;
    default:
      return 0;
  }
}

function nodeExtent(n: Node): [number, number] {
  switch (n.type) {
    case CHAR_NODE: {
      const i = n.font.charInfo(n.char);
      return [n.font.charHeight(i), n.font.charDepth(i)];
    }
    case HLIST_NODE:
    case VLIST_NODE:
      return [n.height - n.shift, n.depth + n.shift];
    case RULE_NODE:
      return [Math.max(0, n.height), Math.max(0, n.depth)];
    default:
      return [0, 0];
  }
}

/** TeX drops glue, kerns, penalties and math nodes at the start of a line (tex.web 879). */
const discardable = (n: Node) => n.type === GLUE_NODE || n.type === KERN_NODE || n.type === PENALTY_NODE || n.type === MATH_NODE;

/**
 * A piece wider than this (em) is cut further, between atoms, where TeX
 * leaves no break (a long run of symbols with no operator would run past the
 * field's edge as one piece, 2026-11-09); narrower ones keep TeX's breaks only.
 */
export const LONGEST_PIECE = 8;

/**
 * The parts such a piece is cut into are at most this wide (em): a field
 * breaks its line between parts only, so a long run ends its lines within so
 * much of the field's edge, about as its own words do (parts of 8 em left up
 * to 8 em of a line empty, 2026-11-09).
 */
export const LONG_PIECE_PART = 2;

/** Where a node starts on the one-line formula and how far it reaches up and down. */
interface Placed {
  x: number;
  height: number;
  depth: number;
  /** an atom starts here: a character or a box of its own (not the shifted box of the scripts of the atom before it) */
  atom: boolean;
}

/**
 * The pieces of a one-line formula (its outermost hlist, natural widths):
 * cut at every penalty below 10000 of the outermost list. A piece ends at
 * where the next starts, the first material after its penalty (after the
 * operator and the space after it). A piece wider than `longest` em is cut
 * again at atom starts, into parts of at most `part` em where the atoms
 * allow (splitLong).
 */
export function formulaPieces(box: BoxNode, longest = LONGEST_PIECE, part = Math.min(longest, LONG_PIECE_PART)): FormulaPiece[] {
  const pieces: FormulaPiece[] = [];
  const placed: Placed[] = [];
  let x = 0;
  let current: FormulaPiece = { start: 0, end: 0, height: 0, depth: 0 };
  let pendingEnd: number | null = null;
  for (const n of box.list) {
    if (pendingEnd !== null && !discardable(n)) {
      // the space after the operator stays with its piece: pieces one after
      // the other on a line are spaced as on the one-line formula
      current.end = x;
      pieces.push(current);
      current = { start: x, end: 0, height: 0, depth: 0 };
      pendingEnd = null;
    }
    if (n.type === PENALTY_NODE && n.penalty < INF_PENALTY && pendingEnd === null) pendingEnd = x;
    const [h, d] = nodeExtent(n);
    if (pendingEnd === null) {
      current.height = Math.max(current.height, h);
      current.depth = Math.max(current.depth, d);
    }
    const atom = n.type === CHAR_NODE || ((n.type === HLIST_NODE || n.type === VLIST_NODE) && n.shift === 0);
    placed.push({ x, height: h, depth: d, atom });
    x += nodeWidth(n);
  }
  current.end = x;
  pieces.push(current);
  return pieces.flatMap((piece) => splitLong(piece, placed, longest * EM, part * EM));
}

/**
 * A piece wider than `longest` (sp) in parts at atom starts: each part ends at
 * the last atom start that keeps it within `part` (sp); an atom wider than
 * that is a part of its own.
 */
function splitLong(piece: FormulaPiece, placed: Placed[], longest: number, part: number): FormulaPiece[] {
  if (piece.end - piece.start <= longest) return [piece];
  const starts = placed.filter((p) => p.atom && p.x > piece.start && p.x < piece.end).map((p) => p.x);
  const cuts: number[] = [];
  let from = piece.start;
  let last: number | null = null;
  for (const at of starts) {
    if (at <= from) continue;
    if (at - from > part && last !== null) {
      cuts.push(last);
      from = last;
      last = null;
    }
    if (at - from > part) {
      cuts.push(at);
      from = at;
      continue;
    }
    last = at;
  }
  if (piece.end - from > part && last !== null) cuts.push(last);
  if (cuts.length === 0) return [piece];
  const edges = [piece.start, ...cuts, piece.end];
  return edges.slice(0, -1).map((start, k) => {
    const end = edges[k + 1];
    let height = 0;
    let depth = 0;
    for (const p of placed) {
      if (p.x < start || p.x >= end) continue;
      height = Math.max(height, p.height);
      depth = Math.max(depth, p.depth);
    }
    return { start, end, height, depth };
  });
}

/**
 * The pieces at their places: every glyph, rule, glue span and box of the
 * shipped one-line formula moves with the piece it is in, and the result's
 * box is the lines' bounding box. Null when everything stays on one line as
 * typeset (nothing to move).
 */
export function placePieces(
  box: BoxNode,
  ship: ShipResult,
  pieces: FormulaPiece[],
  placer: PiecePlacer,
): { box: BoxNode; ship: ShipResult; lines: FormulaLines } | null {
  if (pieces.length === 0) return null;
  const widths = pieces.map((p) => (p.end - p.start) / EM);
  let given: (PiecePlace | null | undefined)[] = [];
  try {
    given = placer(widths) ?? [];
  } catch {
    given = [];
  }
  // each piece's place (sp, from the first piece's start), continuing the line of the one before where none is given
  const places: { x: number; y: number }[] = [];
  for (let k = 0; k < pieces.length; k++) {
    const g = given[k];
    if (k === 0) {
      places.push({ x: 0, y: 0 });
      continue;
    }
    if (g && Number.isFinite(g.x) && Number.isFinite(g.y) && Math.abs(g.x) < 1e4 && Math.abs(g.y) < 1e4) {
      places.push({ x: Math.round(g.x * EM), y: Math.round(g.y * EM) });
    } else {
      const before = places[k - 1];
      places.push({ x: before.x + (pieces[k].start - pieces[k - 1].start), y: before.y });
    }
  }
  const natural = places.every((p, k) => p.y === 0 && Math.abs(p.x - pieces[k].start) <= 1);
  const originX = Math.max(0, -Math.min(...places.map((p) => p.x)));
  const lines: FormulaLines = {
    originX,
    pieces: pieces.map((p, k) => ({ ...p, width: p.end - p.start, x: places[k].x, y: places[k].y })),
  };
  if (natural) return { box, ship, lines };
  const dx = pieces.map((p, k) => originX + places[k].x - p.start);
  const dy = places.map((p) => p.y);
  const starts = pieces.map((p) => p.start);
  const pieceAt = (x: number) => {
    let k = 0;
    while (k + 1 < starts.length && starts[k + 1] <= x) k++;
    return k;
  };
  const tagPiece = new Map<Tag, number>();
  const note = (tag: Tag, k: number) => {
    if (tag !== undefined && tag !== null && !tagPiece.has(tag)) tagPiece.set(tag, k);
  };
  const glyphs = ship.glyphs.map((g) => {
    const k = pieceAt(g.x);
    note(g.tag, k);
    return { ...g, x: g.x + dx[k], y: g.y + dy[k] };
  });
  const rules = ship.rules.map((r) => {
    const k = pieceAt(r.x);
    note(r.tag, k);
    return { ...r, x: r.x + dx[k], y: r.y + dy[k] };
  });
  const spans = ship.spans.map((s) => {
    const k = pieceAt(s.x);
    note(s.tag, k);
    return { ...s, x: s.x + dx[k], y: s.y + dy[k] };
  });
  // the outermost box (the whole formula) is the drawn box itself
  const boxes = ship.boxes.map((b, i) => {
    if (i === 0) return b;
    const k = pieceAt(b.x);
    note(b.box.tag, k);
    return { ...b, x: b.x + dx[k], y: b.y + dy[k] };
  });
  const lineOf = new Map<Tag, number>();
  for (const [tag, y] of ship.lineOf) lineOf.set(tag, y + (dy[tagPiece.get(tag) ?? 0] ?? 0));
  let width = 0;
  let height = 0;
  let depth = 0;
  lines.pieces.forEach((p) => {
    width = Math.max(width, originX + p.x + p.width);
    height = Math.max(height, p.height - p.y);
    depth = Math.max(depth, p.depth + p.y);
  });
  return {
    box: { ...box, width, height, depth },
    ship: { glyphs, rules, spans, boxes, lineOf },
    lines,
  };
}

/**
 * First fit (what a host field's word wrap does): pieces go on the line while
 * they fit, the first line `first` em wide from the formula's start, the
 * following ones `width` em wide from `left` em (from the formula's start),
 * `skip` em apart.
 */
export function fillLines(first: number, width: number, left: number, skip: number): PiecePlacer {
  return (widths) => {
    const out: PiecePlace[] = [];
    let x = 0;
    let y = 0;
    let room = first;
    widths.forEach((w, k) => {
      if (k > 0 && w > room + 1e-6) {
        y += skip;
        x = left;
        room = width;
      }
      out.push({ x, y });
      x += w;
      room -= w;
    });
    return out;
  };
}
