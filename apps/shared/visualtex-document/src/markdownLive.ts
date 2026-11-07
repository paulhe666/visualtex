// Typora's live Markdown in Markdown input (DESIGN.md §7). A formatted span shows its Markdown while the caret is in it
// or at its edges — "**bold**", "`code`", "[text](https://…)", "<u>…</u>" — as grey characters the caret steps through
// and Backspace deletes; the span's text is read as Markdown again after every change (delete one "*" of "**bold**"
// and it is italic, delete the other and it is text), and when the caret leaves, the characters go and the formatting
// stays. Typing the delimiter that closes (or opens) a span formats it the same way, its Markdown still showing until
// the caret moves on; what is typed after it is plain.
//
// The document stores formatting as marks; the characters of the span being shown are real text with the `md_syntax`
// mark (one region at a time, in one textblock, always shown as Markdown that reads back to its marks). Showing and
// hiding them stays out of the undo history; what the user types and deletes is in it, as usual. Exports, saving and
// the clipboard leave the shown characters out (withoutMarkdownSyntax).
import { Fragment, Slice, type Mark, type MarkType, type Node, type Schema } from "prosemirror-model";
import { Plugin, PluginKey, TextSelection, type Command, type EditorState, type Transaction } from "prosemirror-state";
import { AddMarkStep, RemoveMarkStep, ReplaceStep, type Mappable } from "prosemirror-transform";
import { isHistoryTransaction, redo, undo } from "prosemirror-history";
import type { EditorView } from "prosemirror-view";
import {
  EMPTY_MARKERS,
  OBJECT,
  marksOfParse,
  parseInline,
  sameMarks,
  writeInline,
  type InlineKind,
  type InlineMark,
  type InlineParse,
} from "./inlineMarkdown.ts";

/** The highlight Markdown's ==text== gives (markdown.ts's; kept here to avoid an import cycle). */
const MARKDOWN_HIGHLIGHT = "#fff59d";

type Mode = () => boolean;

/** A span of the region, in document positions. */
interface DocSpan {
  kind: InlineKind;
  from: number;
  to: number;
  contentFrom: number;
  contentTo: number;
  bare?: boolean;
}

/** The text shown as Markdown. */
interface Region {
  from: number;
  to: number;
  spans: DocSpan[];
  /** its marker characters */
  syntax: Array<[number, number]>;
}

interface LiveState {
  region: Region | null;
  /** characters typed escaped ("\*"): never Markdown */
  escaped: number[];
  /** text typed since the last reading, where a span may have been closed (or opened) */
  typed: [number, number] | null;
  /** the region's text changed since it was last read */
  edited: boolean;
  /** formatting changed in the region from outside (a toolbar button, undo): its markers go, the marks stay */
  restyled: boolean;
  /** an undo or redo changed these ranges: markers it brought back are read again as typed text */
  restored: Array<[number, number]>;
  /** the caret goes inside the run at its edge rather than outside it (text was just typed into the run; a line was
   *  split before it) */
  inside: "left" | "right" | null;
}

type LiveMeta =
  /** the state a transaction of this plugin leaves */
  | { state: LiveState }
  /** read again now (after a composition or a press) */
  | { settle: true }
  /** a character typed escaped */
  | { escaped: number }
  /** the shown Markdown was hidden by a command: read the caret's place afresh (the caret inside the run at its edge) */
  | { reset: true; inside?: "left" | "right" };

export const markdownLiveKey = new PluginKey<LiveState>("visualtex-document-markdown-live");

const EMPTY_STATE: LiveState = { region: null, escaped: [], typed: null, edited: false, restyled: false, restored: [], inside: null };

// ------------------------------------------------------------------ marks --

const MD_MARKS = ["link", "strong", "em", "strike", "code", "underline", "superscript", "subscript", "kbd", "highlight"];

/** A document mark as Markdown formatting (a highlight only in Markdown's colour; colours and sizes are not). */
function inlineMarkOf(mark: Mark): InlineMark | null {
  const name = mark.type.name;
  if (name === "link") return { kind: "link", href: mark.attrs.href, title: mark.attrs.title ?? null };
  if (name === "highlight") return mark.attrs.color === MARKDOWN_HIGHLIGHT ? { kind: "highlight" } : null;
  if (MD_MARKS.includes(name)) return { kind: name as InlineMark["kind"] };
  return null;
}

function documentMark(schema: Schema, m: InlineMark): Mark | null {
  const type = schema.marks[m.kind];
  if (!type) return null;
  if (m.kind === "link") return type.create({ href: m.href ?? "", title: m.title ?? null });
  if (m.kind === "highlight") return type.create({ color: MARKDOWN_HIGHLIGHT });
  return type.create();
}

const isMarkdownMark = (mark: Mark) => mark.type.name === "md_syntax" || inlineMarkOf(mark) !== null;

const syntaxType = (schema: Schema): MarkType => schema.marks.md_syntax;

// ------------------------------------------------------------------ units --

/** One position of a textblock's inline content: a character, or an inline object (a formula, an image, a break). */
interface Unit {
  pos: number;
  ch: string;
  marks: readonly Mark[];
  syntax: boolean;
  literal: boolean;
}

function unitsOf(doc: Node, from: number, to: number, escaped: readonly number[]): Unit[] {
  const units: Unit[] = [];
  doc.nodesBetween(from, to, (node, pos) => {
    if (node.isText) {
      const text = node.text!;
      const syntax = node.marks.some((m) => m.type.name === "md_syntax");
      for (let p = Math.max(from, pos); p < Math.min(to, pos + text.length); p++) {
        units.push({ pos: p, ch: text[p - pos], marks: node.marks, syntax, literal: escaped.includes(p) });
      }
      return false;
    }
    if (node.isInline) {
      if (pos >= from && pos + node.nodeSize <= to) {
        units.push({ pos, ch: node.type.name === "hard_break" ? "\n" : OBJECT, marks: node.marks, syntax: false, literal: true });
      }
      return false;
    }
    return true;
  });
  return units;
}

const unitMarks = (unit: Unit): InlineMark[] => unit.marks.map(inlineMarkOf).filter((m): m is InlineMark => m !== null);

/** The marks a position should have: its colours and sizes, the formatting read, the marker mark (`syntax`: 1 a
 *  marker, 2 a link's address). */
function wantedMarks(schema: Schema, unit: Unit | null, read: InlineMark[], syntax: number): readonly Mark[] {
  let set: readonly Mark[] = (unit?.marks ?? []).filter((m) => !isMarkdownMark(m));
  for (const m of read) {
    const mark = documentMark(schema, m);
    if (mark) set = mark.addToSet(set);
  }
  if (syntax) set = syntaxType(schema).create({ role: syntax === 2 ? "url" : null }).addToSet(set);
  return set;
}

/** Sets the marks of [from, from + wanted.length): only what differs is changed. */
function setMarks(tr: Transaction, from: number, wanted: ReadonlyArray<readonly Mark[]>) {
  const schema = tr.doc.type.schema;
  const types = Object.values(schema.marks).filter((type) => type.name === "md_syntax" || MD_MARKS.includes(type.name));
  let i = 0;
  while (i < wanted.length) {
    const node = tr.doc.nodeAt(from + i);
    if (!node || sameSet(node.marks, wanted[i])) {
      i++;
      continue;
    }
    let j = i + 1;
    while (j < wanted.length && sameSet(wanted[j], wanted[i])) j++;
    for (const type of types) tr.removeMark(from + i, from + j, type);
    for (const mark of wanted[i]) if (types.includes(mark.type)) tr.addMark(from + i, from + j, mark);
    i = j;
  }
}

const sameSet = (a: readonly Mark[], b: readonly Mark[]) => a.length === b.length && a.every((m, k) => m.eq(b[k]));

// ---------------------------------------------------------------- reading --

/** The spans kept with no text: the region's spans whose text is gone, their markers still there. */
function keptEmpty(region: Region, units: Unit[]) {
  const out: Array<{ from: number; to: number; kind: InlineKind }> = [];
  for (const span of region.spans) {
    if (span.contentFrom !== span.contentTo || !EMPTY_MARKERS[span.kind]) continue;
    const from = span.from - region.from;
    const to = span.to - region.from;
    if (from < 0 || to > units.length || !units.slice(from, to).every((u) => u.syntax)) continue;
    out.push({ from, to, kind: span.kind });
  }
  return out;
}

function regionOf(from: number, parse: InlineParse, length: number): Region {
  const syntax: Array<[number, number]> = [];
  for (let i = 0; i < length; i++) {
    if (!parse.syntax[i]) continue;
    const last = syntax[syntax.length - 1];
    if (last && last[1] === from + i) last[1]++;
    else syntax.push([from + i, from + i + 1]);
  }
  return {
    from,
    to: from + length,
    spans: parse.spans.map((s) => ({ kind: s.kind, from: from + s.from, to: from + s.to, contentFrom: from + s.contentFrom, contentTo: from + s.contentTo, bare: s.bare })),
    syntax,
  };
}

const toMark = (span: { kind: InlineKind; href?: string; title?: string | null }): InlineMark =>
  span.kind === "link" ? { kind: "link", href: span.href ?? "", title: span.title ?? null } : { kind: span.kind as InlineMark["kind"] };

// -------------------------------------------------------------- expanding --

/**
 * Shows [from, to) as Markdown: markers go in around its formatted runs, so that reading the text gives every run its
 * formatting back (tried with "*", with escapes, with "_"; when nothing reads back, it is not shown). `source`: units
 * that are Markdown as they stand (text just typed, the region's own text with its markers): written as they are, and
 * read for what they say. Returns the region and where `caret` went.
 */
function expand(
  tr: Transaction,
  from: number,
  to: number,
  escaped: readonly number[],
  caret: number | null,
  inside: "left" | "right" | null,
  source: (unit: Unit) => boolean = () => false,
  empty: ReadonlyArray<{ from: number; to: number; kind: InlineKind }> = [],
): { region: Region; caret: number | null } | null {
  const units = unitsOf(tr.doc, from, to, escaped);
  if (units.length !== to - from) return null;
  const isSource = units.map(source);
  const attempts: Array<{ em: "*" | "_"; escape: boolean }> = [
    { em: "*", escape: false },
    { em: "*", escape: true },
    { em: "_", escape: true },
  ];
  for (const attempt of attempts) {
    const written = writeInline(
      units.map((u, i) => ({ ch: u.ch, marks: isSource[i] ? [] : unitMarks(u), literal: u.literal || isSource[i] })),
      attempt,
    );
    const text = written.text;
    const literalAt = new Set<number>();
    units.forEach((u, i) => {
      if (u.literal) literalAt.add(written.at[i]);
    });
    // spans kept with no text, where they went in the text
    const kept = empty
      .filter((e) => e.from >= from && e.to <= to)
      .map((e) => ({ from: written.at[e.from - from], to: written.at[e.to - 1 - from] + 1, kind: e.kind }))
      .filter((e, k) => e.to - e.from === empty[k].to - empty[k].from);
    const parse = parseInline({ text, literal: (i) => literalAt.has(i), empty: kept });
    const read = marksOfParse(parse, text.length);
    const isUnit = new Uint8Array(text.length);
    for (const at of written.at) isUnit[at] = 1;
    const creating = isSource.some(Boolean);
    // a code span just closed around formatted text takes its Markdown as text ("`a **b** c`", as in Typora)
    const inCode = (p: number) => creating && read[p].some((s) => s.kind === "code");
    let ok = true;
    for (let p = 0; p < text.length && ok; p++) if (!isUnit[p] && !parse.syntax[p] && !inCode(p)) ok = false;
    for (let i = 0; i < units.length && ok; i++) {
      if (isSource[i]) continue;
      const at = written.at[i];
      if (inCode(at)) continue;
      if (parse.syntax[at]) ok = false;
      // formatted text keeps its formatting (a span just closed around it adds its own)
      else ok = written.marks[i].every((m) => read[at].some((s) => sameMarks([toMark(s)], [m]))) && (creating || sameMarks(read[at].map(toMark), written.marks[i]));
    }
    if (!ok) continue;
    // the markers go in, right to left (positions before them stay)
    const schema = tr.doc.type.schema;
    const syntax = syntaxType(schema).create();
    const gaps: Array<{ pos: number; text: string }> = [];
    let prev = -1;
    for (let i = 0; i < units.length; i++) {
      if (written.at[i] > prev + 1) gaps.push({ pos: units[i].pos, text: text.slice(prev + 1, written.at[i]) });
      prev = written.at[i];
    }
    if (prev + 1 < text.length) gaps.push({ pos: to, text: text.slice(prev + 1) });
    // the caret: after the closers and before the openers at its place, unless it belongs inside the run there
    let caretAt: number | null = null;
    if (caret !== null) {
      const i = Math.max(0, Math.min(units.length, caret - from));
      const escapedChar = i < units.length && written.at[i] > written.split[i] && text[written.at[i] - 1] === "\\" ? 1 : 0;
      const offset = inside === "right" ? written.gapStart[i] : inside === "left" && i < units.length ? written.at[i] - escapedChar : written.split[i];
      caretAt = from + offset;
    }
    for (let k = gaps.length - 1; k >= 0; k--) tr.insert(gaps[k].pos, schema.text(gaps[k].text, [syntax]));
    const region = regionOf(from, parse, text.length);
    // the marks as read
    const atUnit = new Map<number, Unit>();
    units.forEach((u, i) => atUnit.set(written.at[i], u));
    setMarks(
      tr,
      from,
      Array.from({ length: text.length }, (_, p) => wantedMarks(schema, atUnit.get(p) ?? null, read[p].map(toMark), parse.syntax[p])),
    );
    return { region, caret: caretAt };
  }
  return null;
}

// ---------------------------------------------------------------- collapse --

/** The region's markers out of the text (the formatting stays). The region is in the transaction's current positions. */
function collapse(tr: Transaction, region: Region) {
  const type = syntaxType(tr.doc.type.schema);
  const ranges: Array<[number, number]> = region.syntax.map(([a, b]) => [a, b]);
  const from = Math.max(0, region.from);
  const to = Math.min(tr.doc.content.size, region.to);
  if (to > from) {
    tr.doc.nodesBetween(from, to, (node, pos) => {
      if (node.isText && type.isInSet(node.marks)) ranges.push([Math.max(from, pos), Math.min(to, pos + node.nodeSize)]);
      return true;
    });
  }
  deleteRanges(tr, ranges.map(([a, b]) => [Math.max(0, a), Math.min(tr.doc.content.size, b)] as [number, number]));
}

/** Deletes the ranges (overlaps merged), last first. */
function deleteRanges(tr: Transaction, ranges: Array<[number, number]>) {
  const sorted = ranges.filter(([a, b]) => b > a).sort((x, y) => x[0] - y[0]);
  const merged: Array<[number, number]> = [];
  for (const [a, b] of sorted) {
    const last = merged[merged.length - 1];
    if (last && a <= last[1]) last[1] = Math.max(last[1], b);
    else merged.push([a, b]);
  }
  for (let k = merged.length - 1; k >= 0; k--) tr.delete(merged[k][0], merged[k][1]);
}

function mapRegion(region: Region, mapping: Mappable): Region | null {
  const from = mapping.map(region.from, -1);
  const to = mapping.map(region.to, 1);
  if (to <= from) return null;
  const spans: DocSpan[] = [];
  for (const s of region.spans) {
    const span = {
      kind: s.kind,
      from: mapping.map(s.from, 1),
      to: mapping.map(s.to, -1),
      contentFrom: mapping.map(s.contentFrom, -1),
      contentTo: mapping.map(s.contentTo, 1),
      bare: s.bare,
    };
    if (span.to > span.from && span.from >= from && span.to <= to) spans.push(span);
  }
  const syntax: Array<[number, number]> = [];
  for (const [a, b] of region.syntax) {
    const x = mapping.map(a, 1);
    const y = mapping.map(b, -1);
    if (y > x) syntax.push([x, y]);
  }
  return { from, to, spans, syntax };
}

// ------------------------------------------------------------ the caret --

/** The runs of formatted text the caret at `pos` touches (inside or at their ends): their hull, or null. */
function touchedRuns(doc: Node, pos: number): [number, number] | null {
  const $pos = doc.resolve(pos);
  const parent = $pos.parent;
  if (!parent.isTextblock || parent.type.spec.code) return null;
  const start = $pos.start();
  const end = $pos.end();
  let from = pos;
  let to = pos;
  const runOf = (at: number, mark: Mark) => {
    let a = at;
    let b = at + 1;
    for (;;) {
      if (a <= start) break;
      const node = doc.nodeAt(a - 1);
      if (!node || !node.isInline || !mark.isInSet(node.marks)) break;
      a -= node.isText ? 1 : node.nodeSize;
    }
    for (;;) {
      if (b >= end) break;
      const node = doc.nodeAt(b);
      if (!node || !node.isInline || !mark.isInSet(node.marks)) break;
      b += node.isText ? 1 : node.nodeSize;
    }
    from = Math.min(from, a);
    to = Math.max(to, b);
  };
  for (const at of [pos - 1, pos]) {
    if (at < start || at >= end) continue;
    const node = doc.nodeAt(at);
    if (!node?.isInline) continue;
    for (const mark of node.marks) if (inlineMarkOf(mark)) runOf(at, mark);
  }
  return to > from ? [from, to] : null;
}

/** The spans of the region the caret touches: their hull (with what is inside it), or null. */
function touchedSpans(region: Region, pos: number): [number, number] | null {
  let from = Infinity;
  let to = -Infinity;
  for (const s of region.spans) {
    if (s.kind === "escape" || pos < s.from || pos > s.to) continue;
    from = Math.min(from, s.from);
    to = Math.max(to, s.to);
  }
  return to > from ? [from, to] : null;
}

// ---------------------------------------------------------------- reading --

/**
 * The caret's line read again, as Typora reads the whole line: the region's text (markers and all) and plain text are
 * Markdown, formatted text outside the region is one piece each. The spans that take in the region, or whose markers
 * were just typed (`typed`), are shown — the one the caret is at is the new region, the others go back to their
 * formatting; an address just ended by a space is linked. The region's text no span takes in is plain text again.
 */
function readLine(
  tr: Transaction,
  region: Region | null,
  typed: [number, number] | null,
  escaped: readonly number[],
  caret: number | null,
  inside: "left" | "right" | null,
): Region | null {
  const anchor = region ? region.from : typed ? typed[0] : null;
  if (anchor === null || anchor > tr.doc.content.size) return null;
  const $a = tr.doc.resolve(anchor);
  const block = $a.parent;
  const usable = block.isTextblock && !block.type.spec.code && block.type.allowsMarkType(syntaxType(tr.doc.type.schema));
  const lineStart = $a.start();
  const lineEnd = $a.end();
  if (region && (!usable || region.to > lineEnd || region.from < lineStart)) {
    collapse(tr, region);
    return null;
  }
  if (!usable) return null;
  if (typed && (typed[0] < lineStart || typed[1] > lineEnd)) typed = null;
  const lo = Math.min(region?.from ?? Infinity, typed?.[0] ?? Infinity);
  const hi = Math.max(region?.to ?? -Infinity, typed?.[1] ?? -Infinity);
  const from = Math.max(lineStart, lo - 2000);
  const to = Math.min(lineEnd, hi + 2000);
  const units = unitsOf(tr.doc, from, to, escaped);
  if (units.length !== to - from) {
    if (region) collapse(tr, region);
    return null;
  }
  const inRegion = (pos: number) => !!region && pos >= region.from && pos < region.to;
  const text = units.map((u) => u.ch).join("");
  // spans of the region whose text is gone, their markers still there (in document positions)
  const kept = region ? keptEmpty(region, units.slice(region.from - from, region.to - from)).map((e) => ({ ...e, from: e.from + region.from, to: e.to + region.from })) : [];
  // formatted text outside the region is one piece: its characters are never delimiters
  const parse = parseInline({
    text,
    literal: (i) => units[i].literal || (!inRegion(units[i].pos) && (units[i].syntax || unitMarks(units[i]).length > 0)),
    empty: kept.map((e) => ({ ...e, from: e.from - from, to: e.to - from })),
  });
  const rFrom = region ? region.from - from : -1;
  const rTo = region ? region.to - from : -1;
  const tFrom = typed ? typed[0] - from : -1;
  const tTo = typed ? typed[1] - from : -1;
  const accepted = parse.spans.filter((s) => {
    if (s.kind === "escape") return false;
    if (region && s.from < rTo && s.to > rFrom) return true;
    if (!typed) return false;
    // an address: linked when a space is typed after it (its trailing punctuation is the sentence's)
    if (s.bare) return s.to <= tFrom && /^[.,;:!?)\]>'"，。；：！？、）]*$/u.test(text.slice(s.to, tFrom)) && /^\s/u.test(text.slice(tFrom, tTo));
    for (let i = Math.max(s.from, tFrom); i < Math.min(s.to, tTo); i++) if (parse.syntax[i]) return true;
    return false;
  });
  const outer = accepted.filter((s) => !accepted.some((o) => o !== s && o.from <= s.from && o.to >= s.to && (o.from < s.from || o.to > s.to)));
  const schema = tr.doc.type.schema;
  // the region's text outside every span shown: plain text (its markers are characters now)
  if (region) {
    const plain = units
      .map((u, i) => ({ u, i }))
      .filter(({ u, i }) => inRegion(u.pos) && !outer.some((s) => i >= s.from && i < s.to));
    for (const { u } of plain) setMarks(tr, u.pos, [wantedMarks(schema, u, [], 0)]);
  }
  if (!outer.length) return null;
  const base = tr.steps.length;
  const map = (pos: number, assoc = 1) => tr.mapping.slice(base).map(pos, assoc);
  let shown: Region | null = null;
  let shownSteps = 0;
  // right to left: positions to the left stay
  outer.sort((x, y) => y.from - x.from);
  for (const span of outer) {
    const sFrom = map(from + span.from);
    const sTo = map(from + span.to, -1);
    if (span.bare && !(region && span.from < rTo && span.to > rFrom)) {
      const link = documentMark(schema, { kind: "link", href: span.href ?? "", title: null });
      if (link) tr.addMark(sFrom, sTo, link);
      continue;
    }
    const at = caret === null ? null : map(caret);
    const touches = at !== null && at >= sFrom && at <= sTo && !shown;
    // the region's text and plain text are Markdown as they stand; formatted text outside the region gets markers
    const rNow = region ? [map(region.from, -1), map(region.to, 1)] : null;
    const isSource = (u: Unit) => (rNow !== null && u.pos >= rNow[0] && u.pos < rNow[1]) || unitMarks(u).length === 0;
    const keptNow = kept.map((e) => ({ ...e, from: map(e.from), to: map(e.to, -1) }));
    const done = expand(tr, sFrom, sTo, escaped.map((p) => map(p)), touches ? at : null, inside, isSource, keptNow);
    if (!done) continue;
    if (touches) {
      shown = done.region;
      shownSteps = tr.steps.length;
      if (done.caret !== null) tr.setSelection(TextSelection.create(tr.doc, done.caret));
    } else {
      collapse(tr, done.region);
    }
  }
  // what was done to the left of the region moved it
  return shown ? mapRegion(shown, tr.mapping.slice(shownSteps)) : null;
}

/** Characters whose typing can close or open a span (or end an address). */
const TRIGGER = /[*_~=`[\]()<>\s.,;:!?'"\\/，。；：！？、）]/u;

// ------------------------------------------------------------- the plugin --

/** Ranges the transaction changed, in its final positions: text and nodes replaced, and marks changed. */
function changedRanges(tr: Transaction): Array<[number, number]> {
  const out: Array<[number, number]> = [];
  tr.mapping.maps.forEach((map, i) => {
    const rest = tr.mapping.slice(i + 1);
    map.forEach((_oldStart, _oldEnd, newStart, newEnd) => out.push([rest.map(newStart, -1), rest.map(newEnd, 1)]));
  });
  tr.steps.forEach((step, i) => {
    if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
      const rest = tr.mapping.slice(i + 1);
      out.push([rest.map(step.from, 1), rest.map(step.to, -1)]);
    }
  });
  return out;
}

/** Inline content the transaction put in (typed, pasted, restored), in its final positions. */
function insertedText(tr: Transaction): Array<[number, number]> {
  const out: Array<[number, number]> = [];
  tr.steps.forEach((step, i) => {
    if (!(step instanceof ReplaceStep) || !step.slice.size) return;
    let inline = false;
    step.slice.content.descendants((node) => {
      if (node.isInline) inline = true;
      return !inline;
    });
    if (!inline) return;
    step.getMap().forEach((_oldStart, _oldEnd, newStart, newEnd) => {
      if (newEnd <= newStart) return;
      const rest = tr.mapping.slice(i + 1);
      out.push([rest.map(newStart, -1), rest.map(newEnd, 1)]);
    });
  });
  return out;
}

const overlaps = (ranges: Array<[number, number]>, from: number, to: number) => ranges.some(([a, b]) => a <= to && b >= from);

function own(tr: Transaction): LiveMeta | undefined {
  return tr.getMeta(markdownLiveKey) as LiveMeta | undefined;
}

const isResult = (tr: Transaction) => {
  const meta = own(tr);
  return !!meta && "state" in meta;
};

/** The plugin's state, the region given (what was pending is done). */
function withRegion(state: LiveState, region: Region | null, escaped = state.escaped): LiveState {
  return { region, escaped, typed: null, edited: false, restyled: false, restored: [], inside: null };
}

function mapEscaped(escaped: readonly number[], mapping: Mappable): number[] {
  return escaped
    .map((pos) => ({ from: mapping.map(pos, 1), to: mapping.map(pos + 1, -1) }))
    .filter((range) => range.to - range.from === 1)
    .map((range) => range.from);
}

function sameRegion(a: Region | null, b: Region | null): boolean {
  if (!a || !b) return a === b;
  return a.from === b.from && a.to === b.to && a.syntax.length === b.syntax.length && a.syntax.every((r, i) => r[0] === b.syntax[i][0] && r[1] === b.syntax[i][1]);
}

export interface MarkdownLiveOptions {
  /** Markdown input with its inline rules on */
  enabled: Mode;
}

export function markdownLivePlugin(options: MarkdownLiveOptions): Plugin<LiveState> {
  let composing = false;
  let pressed = false;
  let settleView: EditorView | null = null;
  const settle = () => {
    const view = settleView;
    if (!view || view.isDestroyed || view.composing || composing || pressed) return;
    view.dispatch(view.state.tr.setMeta(markdownLiveKey, { settle: true } satisfies LiveMeta));
  };
  const release = () => {
    if (!pressed) return;
    pressed = false;
    setTimeout(settle);
  };
  return new Plugin<LiveState>({
    key: markdownLiveKey,
    state: {
      init: () => EMPTY_STATE,
      apply(tr, value) {
        const meta = own(tr);
        if (meta && "state" in meta) return meta.state;
        let escaped = tr.docChanged ? mapEscaped(value.escaped, tr.mapping) : value.escaped;
        if (meta && "escaped" in meta) escaped = [...escaped, meta.escaped].slice(-64);
        if (meta && "reset" in meta) {
          const reset = withRegion(value, value.region ? mapRegion(value.region, tr.mapping) : null, escaped);
          return { ...reset, region: null, inside: meta.inside ?? null };
        }
        if (!tr.docChanged) return escaped === value.escaped ? value : { ...value, escaped };
        const history = isHistoryTransaction(tr);
        const region = value.region ? mapRegion(value.region, tr.mapping) : null;
        let { edited, restyled } = value;
        if (region && overlaps(changedRanges(tr), region.from, region.to)) {
          if (history || tr.steps.some((step) => step instanceof AddMarkStep || step instanceof RemoveMarkStep)) restyled = true;
          else edited = true;
        }
        const mapRange = (r: [number, number]): [number, number] => [tr.mapping.map(r[0], -1), tr.mapping.map(r[1], 1)];
        let typed = value.typed ? mapRange(value.typed) : null;
        let restored = value.restored.map(mapRange);
        const uiEvent = tr.getMeta("uiEvent");
        if (history) restored = [...restored, ...insertedText(tr)];
        else if (uiEvent !== "paste" && uiEvent !== "drop") {
          for (const [a, b] of insertedText(tr)) typed = typed ? [Math.min(typed[0], a), Math.max(typed[1], b)] : [a, b];
        }
        return { region, escaped, typed, edited, restyled, restored, inside: value.inside };
      },
    },
    appendTransaction(trs, _oldState, state) {
      if (trs.every(isResult)) return null;
      const value = markdownLiveKey.getState(state)!;
      const schema = state.schema;
      if (!schema.marks.md_syntax) return null;
      const tr = state.tr;
      const result = (region: Region | null) =>
        tr.setMeta(markdownLiveKey, { state: withRegion(value, region, mapEscaped(value.escaped, tr.mapping)) } satisfies LiveMeta).setMeta("addToHistory", false);
      if (!options.enabled()) {
        if (!value.region && !value.typed && !value.restored.length && !value.edited && !value.restyled) return null;
        if (value.region) collapse(tr, value.region);
        return result(null);
      }
      // an input method's text, a press of the mouse: read when they are done
      if (composing || pressed || trs.some((t) => t.getMeta("composition") !== undefined)) return null;
      const escapedNow = () => mapEscaped(value.escaped, tr.mapping);
      const sel0 = state.selection;
      const caret0 = sel0 instanceof TextSelection && sel0.empty ? sel0.head : null;
      // the caret goes inside the run at its edge: text was typed into the run, or the line was split before it
      let inside: "left" | "right" | null = value.inside;
      if (!inside && caret0 !== null && value.typed && value.typed[1] === caret0) {
        const before = state.doc.resolve(caret0).nodeBefore;
        if (before && before.marks.some((m) => inlineMarkOf(m))) inside = "right";
      }
      let region = value.region;
      if (region && value.restyled) {
        // formatting changed from outside, or an undo: the markers go, the marks stay
        collapse(tr, region);
        region = null;
      }
      // markers an undo brought back: plain text again, read like typed text
      let typed = value.typed ? ([tr.mapping.map(value.typed[0], -1), tr.mapping.map(value.typed[1], 1)] as [number, number]) : null;
      let force = false;
      if (value.restored.length) {
        const type = syntaxType(schema);
        for (const [a0, b0] of value.restored) {
          const from = Math.max(0, tr.mapping.map(a0, -1));
          const to = Math.min(tr.doc.content.size, tr.mapping.map(b0, 1));
          if (to <= from || !tr.doc.rangeHasMark(from, to, type) || (region && from >= region.from && to <= region.to)) continue;
          tr.removeMark(from, to, type);
          typed = typed ? [Math.min(typed[0], from), Math.max(typed[1], to)] : [from, to];
          force = true;
        }
      }
      // typed in a place far from the region (another line): the region goes back to its formatting first
      if (region && typed && !tr.doc.resolve(region.from).sameParent(tr.doc.resolve(Math.min(typed[0], tr.doc.content.size)))) {
        const start = tr.steps.length;
        collapse(tr, region);
        region = null;
        const m = tr.mapping.slice(start);
        typed = [m.map(typed[0], -1), m.map(typed[1], 1)];
      }
      const triggered = !!typed && typed[1] > typed[0] && (force || TRIGGER.test(tr.doc.textBetween(typed[0], Math.min(typed[1], tr.doc.content.size), "\n", OBJECT)));
      if ((region && value.edited) || triggered) {
        const sel = tr.selection;
        region = readLine(tr, region, triggered ? typed : null, escapedNow(), sel instanceof TextSelection && sel.empty ? sel.head : null, inside);
      }
      region = place(tr, region, escapedNow, inside);
      if (!tr.docChanged && sameRegion(value.region, region) && !value.typed && !value.edited && !value.restyled && !value.restored.length) return null;
      return result(region);
    },
    view(view) {
      settleView = view;
      const up = () => release();
      window.addEventListener("mouseup", up, true);
      window.addEventListener("pointerup", up, true);
      window.addEventListener("pointercancel", up, true);
      window.addEventListener("blur", up);
      return {
        destroy() {
          window.removeEventListener("mouseup", up, true);
          window.removeEventListener("pointerup", up, true);
          window.removeEventListener("pointercancel", up, true);
          window.removeEventListener("blur", up);
          if (settleView === view) settleView = null;
        },
      };
    },
    props: {
      handleDOMEvents: {
        compositionstart() {
          composing = true;
          return false;
        },
        compositionend() {
          composing = false;
          setTimeout(settle);
          return false;
        },
        mousedown(_view, event) {
          if ((event as MouseEvent).button === 0) pressed = true;
          return false;
        },
        keydown() {
          // a press that never saw its release (the pointer left the window)
          if (pressed) pressed = false;
          return false;
        },
        beforeinput(view, event) {
          const type = (event as InputEvent).inputType;
          if (type === "historyUndo" || type === "historyRedo") {
            const tr = collapseTransaction(view.state);
            if (tr) view.dispatch(tr);
          }
          return false;
        },
      },
    },
  });
}

/**
 * The region for the selection: the spans the caret touches are shown (those it no longer touches go back to their
 * formatting), runs it reaches outside the region are shown with them; a range keeps the region it touches.
 */
function place(tr: Transaction, region: Region | null, escaped: () => number[], inside: "left" | "right" | null): Region | null {
  const sel = tr.selection;
  if (!(sel instanceof TextSelection)) {
    if (region) collapse(tr, region);
    return null;
  }
  if (!sel.empty) {
    if (region && sel.from <= region.to && sel.to >= region.from) return region;
    if (region) collapse(tr, region);
    return null;
  }
  const caret = sel.head;
  const $c = tr.doc.resolve(caret);
  const usable = $c.parent.isTextblock && !$c.parent.type.spec.code && $c.parent.type.allowsMarkType(syntaxType(tr.doc.type.schema));
  if (region && usable && caret >= region.from && caret <= region.to && tr.doc.resolve(region.from).parent === $c.parent) {
    const hull = touchedSpans(region, caret);
    const outside = caret === region.from || caret === region.to ? touchedRuns(tr.doc, caret) : null;
    const reachesOut = !!outside && (outside[0] < region.from || outside[1] > region.to);
    if (hull && !reachesOut) {
      if (hull[0] === region.from && hull[1] === region.to) return region;
      // the spans the caret left go back to their formatting
      const keep = (r: [number, number]) => r[0] >= hull[0] && r[1] <= hull[1];
      const kept: Region = { from: hull[0], to: hull[1], spans: region.spans.filter((s) => s.from >= hull[0] && s.to <= hull[1]), syntax: region.syntax.filter(keep) };
      const start = tr.steps.length;
      deleteRanges(tr, region.syntax.filter((r) => !keep(r)));
      return mapRegion(kept, tr.mapping.slice(start));
    }
  }
  if (region) collapse(tr, region);
  if (!usable) return null;
  const at = tr.selection.head;
  const runs = touchedRuns(tr.doc, at);
  if (!runs) return null;
  const shown = expand(tr, runs[0], runs[1], escaped(), at, inside);
  if (!shown) return null;
  if (shown.caret !== null) tr.setSelection(TextSelection.create(tr.doc, shown.caret));
  return shown.region;
}

// -------------------------------------------------------------- commands --

/** The transaction that hides the shown Markdown (null: nothing shown). */
export function collapseTransaction(state: EditorState): Transaction | null {
  const value = markdownLiveKey.getState(state);
  if (!value?.region) return null;
  const tr = state.tr;
  collapse(tr, value.region);
  return tr.setMeta(markdownLiveKey, { state: withRegion(value, null, mapEscaped(value.escaped, tr.mapping)) } satisfies LiveMeta).setMeta("addToHistory", false);
}

/** `command` on the document as it is stored (the shown Markdown hidden first). */
export function withMarkdownHidden(command: Command): Command {
  return (state, dispatch, view) => {
    const tr = collapseTransaction(state);
    if (!tr || !dispatch) return command(state, dispatch, view);
    if (view) {
      view.dispatch(tr);
      return command(view.state, view.dispatch, view);
    }
    const next = state.apply(tr);
    dispatch(tr);
    return command(next, dispatch);
  };
}

export const markdownUndo: Command = withMarkdownHidden(undo);
export const markdownRedo: Command = withMarkdownHidden(redo);

/** Records `pos` as a character typed escaped (never Markdown). */
export function markEscaped(tr: Transaction, pos: number): Transaction {
  return tr.setMeta(markdownLiveKey, { escaped: pos } satisfies LiveMeta);
}

export function isEscaped(state: EditorState, from: number, to: number): boolean {
  return (markdownLiveKey.getState(state)?.escaped ?? []).some((pos) => pos >= from && pos < to);
}

/** The text shown as Markdown, if any. */
export function shownMarkdown(state: EditorState): { from: number; to: number } | null {
  const region = markdownLiveKey.getState(state)?.region;
  return region ? { from: region.from, to: region.to } : null;
}

/** The innermost span of `kind` in the region whose text holds `pos` (between its markers). */
function spanAround(state: EditorState, kind: InlineKind, pos: number): DocSpan | null {
  const region = markdownLiveKey.getState(state)?.region;
  if (!region) return null;
  let best: DocSpan | null = null;
  for (const s of region.spans) {
    if (s.kind !== kind || s.bare || pos < s.contentFrom || pos > s.contentTo) continue;
    if (!best || s.from > best.from) best = s;
  }
  return best;
}

/** The run of `type` around `pos` in its textblock. */
function runOfMark(doc: Node, pos: number, type: MarkType): [number, number] | null {
  const $pos = doc.resolve(pos);
  const start = $pos.start();
  const end = $pos.end();
  const has = (at: number) => {
    const node = at >= start && at < end ? doc.nodeAt(at) : null;
    return !!node && !!type.isInSet(node.marks);
  };
  if (!has(pos - 1) && !has(pos)) return null;
  let a = has(pos - 1) ? pos - 1 : pos;
  let b = has(pos) ? pos + 1 : pos;
  while (has(a - 1)) a--;
  while (has(b)) b++;
  return [a, b];
}

const WORD = /[\p{L}\p{N}_]/u;

type SegmenterLike = new (locale: string, options: { granularity: string }) => { segment(text: string): Iterable<{ index: number; segment: string; isWordLike?: boolean }> };

/** The word the caret at `pos` touches (a word character before or after it), as the system cuts words (Chinese
 *  too); null between words. */
function wordAt(state: EditorState, pos: number): [number, number] | null {
  const $pos = state.doc.resolve(pos);
  const parent = $pos.parent;
  const start = $pos.start();
  const text = parent.textBetween(0, parent.content.size, undefined, OBJECT);
  const offset = $pos.parentOffset;
  const at = WORD.test(text[offset - 1] ?? "") ? offset - 1 : WORD.test(text[offset] ?? "") ? offset : -1;
  if (at < 0) return null;
  const Segmenter = (Intl as unknown as { Segmenter?: SegmenterLike }).Segmenter;
  if (Segmenter) {
    for (const part of new Segmenter("zh", { granularity: "word" }).segment(text)) {
      if (part.index <= at && part.index + part.segment.length > at) {
        return part.isWordLike === false ? null : [start + part.index, start + part.index + part.segment.length];
      }
    }
    return null;
  }
  let a = at;
  let b = at + 1;
  while (a > 0 && WORD.test(text[a - 1])) a--;
  while (b < text.length && WORD.test(text[b])) b++;
  return [start + a, start + b];
}

/** A transaction that hides the shown Markdown, then makes `change` (the caret is placed afresh after it; `inside`:
 *  inside the run at its edge). */
function afterHiding(state: EditorState, change: (tr: Transaction) => void, inside?: "left" | "right"): Transaction {
  const value = markdownLiveKey.getState(state);
  const tr = state.tr;
  if (value?.region) collapse(tr, value.region);
  change(tr);
  return tr.setMeta(markdownLiveKey, { reset: true, inside } satisfies LiveMeta).scrollIntoView();
}

/**
 * Typora's Ctrl+B / Ctrl+I / Ctrl+U / Ctrl+Shift+` / Alt+Shift+5 with no selection: in a span of that kind, the span's
 * markers go (its text stays); in a word, the word takes the formatting; elsewhere an empty pair ("****") goes in with
 * the caret between it, and what is typed there takes the formatting. A selection is `selectionCommand`'s (it takes
 * or loses the formatting as a whole).
 */
export function typoraFormat(kind: "strong" | "em" | "underline" | "code" | "strike", selectionCommand: Command): Command {
  return (state, dispatch, view) => {
    const sel = state.selection;
    if (!(sel instanceof TextSelection)) return selectionCommand(state, dispatch, view);
    const schema = state.schema;
    if (!sel.empty) {
      // text selected in one line (no formulas in it): it takes the formatting, or loses it when all of it has it; the
      // caret goes to its end, inside the span (Typora)
      const type = schema.marks[kind];
      const { $from, $to, from, to } = sel;
      let plain = $from.sameParent($to) && $from.parent.isTextblock && !$from.parent.type.spec.code && !!type && $from.parent.type.allowsMarkType(type);
      let all = true;
      state.doc.nodesBetween(from, to, (node) => {
        if (node.isInline && !node.isText) plain = false;
        if (node.isText && !node.marks.some((m) => m.type === syntaxType(schema)) && !type?.isInSet(node.marks)) all = false;
        return true;
      });
      if (!plain) return selectionCommand(state, dispatch, view);
      if (dispatch) {
        dispatch(
          afterHiding(
            state,
            (tr) => {
              const a = tr.mapping.map(from, 1);
              const b = tr.mapping.map(to, -1);
              if (all) tr.removeMark(a, b, type);
              else tr.addMark(a, b, type.create());
              tr.setSelection(TextSelection.create(tr.doc, b));
            },
            all ? undefined : "right",
          ),
        );
      }
      return true;
    }
    const $c = sel.$head;
    const type = schema.marks[kind];
    if (!type || !$c.parent.isTextblock || $c.parent.type.spec.code || !$c.parent.type.allowsMarkType(type)) return false;
    const value = markdownLiveKey.getState(state) ?? EMPTY_STATE;
    const span = spanAround(state, kind, $c.pos);
    if (span && value.region) {
      if (dispatch) {
        // the markers go: the span is text again (a step of the history)
        const markers = value.region.syntax.filter(([a, b]) => a >= span.from && b <= span.to && (b <= span.contentFrom || a >= span.contentTo));
        const tr = state.tr;
        deleteRanges(tr, markers);
        dispatch(tr.scrollIntoView());
      }
      return true;
    }
    // a run of it shown without markers (they could not be written): the run loses it
    const run = !value.region ? runOfMark(state.doc, $c.pos, type) : null;
    if (run && type.isInSet($c.marks())) {
      if (dispatch) dispatch(afterHiding(state, (tr) => tr.removeMark(run[0], run[1], type)));
      return true;
    }
    const word = wordAt(state, $c.pos);
    if (word) {
      // the caret stays where it was in the word, inside the span at its ends
      const inside = $c.pos === word[1] ? "right" : $c.pos === word[0] ? "left" : undefined;
      if (dispatch) dispatch(afterHiding(state, (tr) => tr.addMark(tr.mapping.map(word[0]), tr.mapping.map(word[1]), type.create()), inside));
      return true;
    }
    if (!dispatch) return true;
    const markers = EMPTY_MARKERS[kind]!;
    const half = markers.length / 2;
    const tr = state.tr;
    if (value.region) collapse(tr, value.region);
    const at = tr.mapping.map($c.pos);
    tr.insert(at, schema.text(markers, [syntaxType(schema).create(), ...$c.marks().filter((m) => !isMarkdownMark(m))]));
    tr.setSelection(TextSelection.create(tr.doc, at + half));
    const region: Region = {
      from: at,
      to: at + markers.length,
      spans: [{ kind, from: at, to: at + markers.length, contentFrom: at + half, contentTo: at + half }],
      syntax: [[at, at + markers.length]],
    };
    dispatch(tr.setMeta(markdownLiveKey, { state: withRegion(value, region, mapEscaped(value.escaped, tr.mapping)) } satisfies LiveMeta).setMeta("addToHistory", false).scrollIntoView());
    return true;
  };
}

/**
 * Typora's Ctrl+K: the selection becomes a link's text, "[text]()", the caret where its address goes; with no
 * selection "[]()" with the caret in the brackets. In a link, `inLink` (the app's link dialog).
 */
export function typoraLink(inLink: Command): Command {
  return (state, dispatch, view) => {
    const { $from, $to, empty, from, to } = state.selection;
    if (!$from.parent.isTextblock || $from.parent.type.spec.code || !$from.sameParent($to)) return false;
    const linkType = state.schema.marks.link;
    if (!linkType) return false;
    const value = markdownLiveKey.getState(state);
    const inShown = !!spanAround(state, "link", from);
    const inStored = !value?.region && (!!runOfMark(state.doc, from, linkType) || state.doc.rangeHasMark(from, to, linkType));
    if (inShown || inStored) return inLink(state, dispatch, view);
    if (!dispatch) return true;
    const tr = state.tr;
    const text = (t: string) => state.schema.text(t);
    if (empty) {
      tr.insert(from, text("[]()"));
      tr.setSelection(TextSelection.create(tr.doc, from + 1));
    } else {
      tr.insert(to, text("]()"));
      tr.insert(from, text("["));
      tr.setSelection(TextSelection.create(tr.doc, to + 3));
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Ctrl+\ (Typora's Clear Format): the selection's text formatting goes (links stay); with no selection, the
 *  formatting of the span the caret is in. */
export const clearFormat: Command = (state, dispatch) => {
  const sel = state.selection;
  const schema = state.schema;
  const types = Object.values(schema.marks).filter((t) => t.name !== "link" && t.name !== "md_syntax");
  let range: [number, number] | null = [sel.from, sel.to];
  if (sel.empty) {
    const value = markdownLiveKey.getState(state);
    range = value?.region ? touchedSpans(value.region, sel.head) : touchedRuns(state.doc, sel.head);
    if (!range) return false;
  }
  if (!dispatch) return true;
  const [from, to] = range;
  dispatch(
    afterHiding(state, (tr) => {
      const a = tr.mapping.map(from, 1);
      const b = tr.mapping.map(to, -1);
      for (const type of types) tr.removeMark(a, b, type);
    }),
  );
  return true;
};

// ------------------------------------------------------- what leaves the editor --

/** The document without the Markdown being shown (what is saved, exported and copied). */
export function withoutMarkdownSyntax(doc: Node): Node {
  const type = doc.type.schema.marks.md_syntax;
  if (!type || !doc.rangeHasMark(0, doc.content.size, type)) return doc;
  return doc.copy(stripFragment(doc.content, type));
}

export function sliceWithoutMarkdownSyntax(slice: Slice, schema: Schema): Slice {
  const type = schema.marks.md_syntax;
  if (!type) return slice;
  let found = false;
  slice.content.descendants((node) => {
    if (node.isText && type.isInSet(node.marks)) found = true;
    return !found;
  });
  return found ? new Slice(stripFragment(slice.content, type), slice.openStart, slice.openEnd) : slice;
}

function stripFragment(fragment: Fragment, type: MarkType): Fragment {
  const out: Node[] = [];
  fragment.forEach((node) => {
    if (node.isText) {
      if (!type.isInSet(node.marks)) out.push(node);
      return;
    }
    out.push(node.isLeaf ? node : node.copy(stripFragment(node.content, type)));
  });
  return Fragment.from(out);
}
