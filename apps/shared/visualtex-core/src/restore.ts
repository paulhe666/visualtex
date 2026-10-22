// Restore after a rewrite (docs/editing-core/DESIGN.md §6.10). When the
// formula is replaced (an input conversion, a reparse, symbol re-rendering),
// the caret is carried through an alignment of the old and new atoms: if the
// atom before the caret survived, the caret follows it; if it was rewritten,
// the caret goes after its replacement. Not by distance from the end, which
// fails as soon as anything after the caret changes too.
import { goTo, goToStop } from "./navigation.ts";
import { allowedStops, type CoreModel } from "./stops.ts";
import type { CoreAtom } from "./structures.ts";

/** The caret and the formula it was in, to carry through a rewrite. */
export interface CaretSnapshot {
  atoms: string[];
  position: number;
  anchor: number;
  collapsed: boolean;
}

/** An atom as the alignment compares it: what it is and where in the structure. */
function describe(atom: CoreAtom | undefined): string {
  if (!atom) return "";
  let depth = 0;
  for (let p = atom.parent; p; p = p.parent) depth++;
  const branch = Array.isArray(atom.parentBranch) ? atom.parentBranch.join(",") : atom.parentBranch ?? "";
  return `${atom.type ?? ""}|${atom.value ?? ""}|${atom.command ?? ""}|${branch}|${depth}`;
}

function describeAll(model: CoreModel): string[] {
  const out: string[] = [];
  for (let offset = 0; offset <= model.lastOffset; offset++) out.push(describe(model.at(offset)));
  return out;
}

export function captureCaret(model: CoreModel): CaretSnapshot {
  return { atoms: describeAll(model), position: model.position, anchor: model.anchor, collapsed: model.selectionIsCollapsed };
}

/** For each old offset, the new offset of the same atom (-1: rewritten), by a longest common subsequence. */
function align(a: string[], b: string[]): number[] {
  const n = a.length;
  const m = b.length;
  // suffix LCS lengths: lcs[i][j] for a[i..], b[j..]
  const lcs = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));
  for (let i = n - 1; i >= 0; i--)
    for (let j = m - 1; j >= 0; j--) lcs[i][j] = a[i] === b[j] ? lcs[i + 1][j + 1] + 1 : Math.max(lcs[i + 1][j], lcs[i][j + 1]);
  const map = new Array<number>(n).fill(-1);
  for (let i = 0, j = 0; i < n && j < m; ) {
    if (a[i] === b[j]) map[i++] = j++;
    else if (lcs[i + 1][j] >= lcs[i][j + 1]) i++;
    else j++;
  }
  return map;
}

/** Where an old offset is in the new formula. */
function carry(map: number[], offset: number, newLast: number): number {
  if (map[offset] >= 0) return map[offset];
  // rewritten: after the replacement, i.e. just before the next atom that survived
  for (let i = offset + 1; i < map.length; i++) if (map[i] >= 0) return Math.max(0, map[i] - 1);
  return newLast;
}

/** Put the caret (a selection collapses to its moving end) where the snapshot's caret went. */
export function restoreCaret(model: CoreModel, snapshot: CaretSnapshot): boolean {
  const map = align(snapshot.atoms, describeAll(model));
  let target = Math.min(carry(map, snapshot.position, model.lastOffset), model.lastOffset);
  const stops = allowedStops(model);
  if (!stops.includes(target)) {
    // the nearest stop at or after it (else before it)
    target = stops.find((s) => s >= target) ?? stops[stops.length - 1] ?? 0;
    return goTo(model, target);
  }
  return goToStop(model, target);
}
