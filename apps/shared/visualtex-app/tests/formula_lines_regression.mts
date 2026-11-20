// A formula set on several lines (shared/tex-layout lines.ts: VisualTeX Anywhere wraps a formula as
// the host field wraps its text), on synthetic boxes. Run: node --test ../shared/visualtex-app/tests/formula_lines_regression.mts
import { test } from "node:test";
import assert from "node:assert/strict";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
// apps/shared/tex-layout, from apps/<platform>/scripts
const root = pathToFileURL(join(dirname(fileURLToPath(import.meta.url)), "..", "..", "tex-layout", "src")).href;
const { formulaPieces, placePieces, fillLines } = await import(`${root}/lines.ts`);
const { GLUE_NODE, KERN_NODE, PENALTY_NODE, HLIST_NODE, MATH_NODE } = await import(`${root}/nodes.ts`);
const EM = 655360;
const box = (w: number, tag: string, h = 0.7, d = 0.2) => ({ type: HLIST_NODE, width: w * EM, height: h * EM, depth: d * EM, shift: 0, list: [], glueSet: 0, glueSign: 0, glueOrder: 0, tag });
const glue = (w: number) => ({ type: GLUE_NODE, spec: { width: w * EM, stretch: 0, stretchOrder: 0, shrink: 0, shrinkOrder: 0 }, subtype: 0 });
const pen = (p: number) => ({ type: PENALTY_NODE, penalty: p });
const on = { type: MATH_NODE, width: 0, subtype: 0 };
// a + b = c  : [a][g][+][pen700][g][b][g][=][pen500][g][c]
const list = [on, box(0.5, "a"), glue(0.22), box(0.78, "+"), pen(700), glue(0.22), box(0.43, "b"), glue(0.28), box(0.78, "="), pen(500), glue(0.28), box(0.43, "c"), { ...on }];
const total = list.reduce((x: number, n: any) => x + (n.width ?? n.spec?.width ?? 0), 0);
const outer = { type: HLIST_NODE, width: total, height: 0.7 * EM, depth: 0.2 * EM, shift: 0, list, glueSet: 0, glueSign: 0, glueOrder: 0 };

test("pieces end after each operator with the space after it", () => {
  const pieces = formulaPieces(outer as any);
  assert.equal(pieces.length, 3);
  assert.equal(pieces[0].start, 0);
  // a + (space) | b = (space) | c
  assert.equal(pieces[1].start, (0.5 + 0.22 + 0.78 + 0.22) * EM);
  assert.equal(pieces[0].end, pieces[1].start);
  assert.equal(pieces[2].end, total);
});

test("penalties of 10000 are no break places", () => {
  const l = list.map((n: any) => (n.type === PENALTY_NODE ? pen(10000) : n));
  assert.equal(formulaPieces({ ...outer, list: l } as any).length, 1);
});

const ship = () => ({
  glyphs: [],
  rules: [],
  spans: [],
  boxes: [{ box: outer, x: 0, y: 0 }, ...(() => { let x = 0; const out: any[] = []; for (const n of list as any[]) { if (n.type === HLIST_NODE) out.push({ box: n, x, y: 0 }); x += n.width ?? n.spec?.width ?? 0; } return out; })()],
  lineOf: new Map(["a", "+", "b", "=", "c"].map((t) => [t, 0])),
});

test("on one line the layout is untouched", () => {
  const s = ship();
  const placed = placePieces(outer as any, s as any, formulaPieces(outer as any), () => []);
  assert.ok(placed);
  assert.equal(placed!.box, outer);
  assert.equal(placed!.ship, s);
});

test("pieces go where the placer puts them and the box covers the lines", () => {
  const pieces = formulaPieces(outer as any);
  // second piece on a line 1.5em lower, starting 1em left of the formula's start
  const placed = placePieces(outer as any, ship() as any, pieces, (w: number[]) => [{ x: 0, y: 0 }, { x: -1, y: 1.5 }, null])!;
  assert.equal(placed.lines.originX, EM);
  const b = placed.ship.boxes.find((x: any) => x.box.tag === "b");
  assert.equal(b.x, 0); // the line's start, the box's left edge
  assert.equal(b.y, 1.5 * EM);
  const c = placed.ship.boxes.find((x: any) => x.box.tag === "c");
  // the third piece (no place given) continues the second's line, spaced as in the formula
  assert.equal(c.y, 1.5 * EM);
  assert.equal(c.x - b.x, pieces[2].start - pieces[1].start);
  assert.equal(placed.ship.lineOf.get("c"), 1.5 * EM);
  assert.ok(placed.box.depth >= 1.5 * EM + 0.2 * EM - 1);
  assert.ok(placed.box.width >= EM + 0.5 * EM);
});

test("first fit breaks where the line is full", () => {
  const place = fillLines(2, 10, -1, 1.2)([1.5, 1.5, 1.5]);
  assert.deepEqual(place[0], { x: 0, y: 0 });
  assert.deepEqual(place[1], { x: -1, y: 1.2 });
  assert.deepEqual(place[2], { x: 0.5, y: 1.2 });
});

// a long run of atoms with no operator (x x x … : twenty 1em boxes)
const run = (n: number, script = false) => {
  const items: any[] = [on];
  for (let k = 0; k < n; k++) {
    items.push(box(1, `x${k}`));
    // a script box hangs on its atom: shifted
    if (script && k === 4) items.push({ ...box(0.4, "sup"), shift: -0.3 * EM });
  }
  items.push({ ...on });
  const width = items.reduce((x: number, n: any) => x + (n.width ?? n.spec?.width ?? 0), 0);
  return { type: HLIST_NODE, width, height: 0.7 * EM, depth: 0.2 * EM, shift: 0, list: items, glueSet: 0, glueSign: 0, glueOrder: 0 };
};

test("a piece wider than the longest is cut at atom starts", () => {
  const pieces = formulaPieces(run(20) as any);
  assert.ok(pieces.length >= 3);
  for (const p of pieces) assert.ok(p.end - p.start <= 8 * EM, `piece ${p.start / EM}-${p.end / EM} em`);
  assert.equal(pieces[0].start, 0);
  assert.equal(pieces.at(-1)!.end, 20 * EM);
  for (let k = 1; k < pieces.length; k++) assert.equal(pieces[k].start, pieces[k - 1].end);
});

test("the parts of a long piece are short: the field breaks its line between them", () => {
  const pieces = formulaPieces(run(20) as any);
  assert.equal(pieces.length, 10);
  for (const p of pieces) assert.ok(p.end - p.start <= 2 * EM, `piece ${p.start / EM}-${p.end / EM} em`);
});

test("a script box stays with its atom", () => {
  const pieces = formulaPieces(run(12, true) as any, 4.2);
  // the shifted box after x4 starts at 5em: no piece starts there
  assert.ok(!pieces.some((p: any) => p.start === 5 * EM), JSON.stringify(pieces.map((p: any) => p.start / EM)));
});

test("pieces within the longest keep TeX's breaks only", () => {
  assert.equal(formulaPieces(run(6) as any).length, 1);
});
