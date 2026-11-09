// A LaTeX document edited on its pages (personal edition), as the block editor edits its page: a click puts the caret
// between the two glyphs clicked between, typing goes into the source there and the paragraph is typeset again at
// once (the live server, the document's own TeX) and drawn over its old lines; "$" opens the formula editor where
// the caret is, a click on a formula opens it; "\" offers the LaTeX commands. The source (CodeMirror, hidden) keeps
// the text and its history; TeX makes the whole document again a moment after the typing stops.
//
// The caret's places come from texCaret.ts: the paragraph's source matched to the glyphs TeX set for it (the live
// server's map of them; the page's own SVG for a paragraph the server cannot typeset alone).
import type { ChangeSet } from "@codemirror/state";
import type { FormulaEntry } from "../../../../shared/visualtex-document/src/index.ts";
import type { LatexSourceHandle } from "./LatexSource";
import { shiftAt } from "./LatexPages";
import { lineNumberAt, pageChars, paragraphArea, visualUnit, type DocMacros, type LiveOverlay, type ParagraphChange } from "./livePara";
import {
  atomAt,
  deletionRanges,
  emptyGroupsAround,
  escapeTyped,
  layoutOf,
  lineEdges,
  lineStop,
  liveItems,
  offsetAt,
  movedLines,
  pageItems,
  selectionRects,
  stepStop,
  stopOf,
  tokenize,
  type CaretLayout,
  type LayoutLine,
} from "./texCaret";
import type { SyncTex } from "./synctex";
import type { LiveResult } from "./texService";

/** What the live server made of a unit's text: where it is drawn and what it drew, or why it could not. */
export type TypesetOutcome =
  /** (x0, y0): the paragraph's box in TeX's coordinates; (tx, ty): how far the page draws it from there */
  | { ok: true; result: LiveResult; page: number; x0: number; y0: number; tx: number; ty: number; overlay: LiveOverlay | null }
  | { ok: false; why: string; unfinished?: boolean; retry?: boolean };

export interface TypesetOptions {
  /** what the paragraph is sent as (a formula being edited: its room) */
  transform?: (paragraph: string) => string;
}

/** What the editor needs of the workspace. */
export interface VisualHost {
  source(): LatexSourceHandle | null;
  /** a file's text now (undefined: not read yet) */
  textOf(path: string): string | undefined;
  /** read a file and show it in the source editor (the editing goes through it) */
  openFile(path: string): Promise<boolean>;
  /** the build the pages show: its number, the texts it was made from */
  built(): { build: number; texts: Map<string, string> } | null;
  /** a file's text as the shown build had it */
  builtText(path: string): string | undefined;
  /** an offset of the build's text of `path` in the text now, and back */
  toCurrent(path: string, pos: number, assoc: -1 | 1): number;
  toBuilt(path: string, pos: number, assoc: -1 | 1): number;
  sync(): Promise<{ tex: SyncTex; files: Array<string | null> } | null>;
  /** a page's SVG when it shows the current build */
  pageSvg(page: number): string | null;
  pageCount(): number;
  pageHeight(page: number): number;
  macros(): DocMacros;
  commands(): Map<string, string>;
  typeset(path: string, text: string, change: ParagraphChange, options: TypesetOptions): Promise<TypesetOutcome>;
  /** the overlays changed (the pages redraw) */
  overlaysChanged(): void;
  /** the page where a source offset is typeset brought into view (its page loads) */
  revealSource(path: string, offset: number): void;
  /** a click where the page cannot be edited: said briefly */
  cannotEdit(): void;
  /** a run of the whole document: `live` when the pages are kept current by the live server meanwhile */
  scheduleCompile(live: boolean): void;
  isMac: boolean;
}

/** A stretch of a file edited on the pages: its paragraph (or structure line) now and as the shown build had it. */
export interface VUnit {
  path: string;
  start: number;
  end: number;
  build: number;
  bStart: number;
  bEnd: number;
  /** the caret's places: the unit as last typeset (its text `shown`) */
  layout: CaretLayout | null;
  shown: string | null;
  /** the live server typesets it (null: not asked yet) */
  live: boolean | null;
  overlay: LiveOverlay | null;
  busy: boolean;
  again: boolean;
  /** narrowed to its text part (its displays could not be typeset with it) */
  narrowed?: boolean;
}

export interface Caret {
  path: string;
  offset: number;
  /** the line of the layout it shows on (an offset at a line break is the end of one and the start of the next) */
  line?: number;
  /** the selection's other end (= offset: none) */
  anchor: number;
  /** the x up/down keep */
  goalX?: number;
}

/** A formula open in the formula editor on the page. */
export interface OpenFormula {
  unit: VUnit;
  /** its source (delimiters included) now; null: a new one, not in the source until it has something */
  range: [number, number] | null;
  /** where a new one goes */
  at: number;
  display: boolean;
  open: string;
  close: string;
  latex: string;
  entry: FormulaEntry;
  /** the editor's box (bp): width, height above the baseline, depth below */
  size: { w: number; h: number; d: number } | null;
  /** where it is drawn: the room TeX left for it (page bp, baseline) */
  place: { page: number; x: number; y: number } | null;
  /** its font size (bp) */
  fontSize: number;
}

/** The "\" menu: what is typed after the backslash. */
export interface CommandMenuState {
  query: string;
  index: number;
  /** where the backslash was typed */
  at: number;
}

export interface CaretView {
  page: number;
  x: number;
  top: number;
  bottom: number;
}

const SP_PER_BP = 65781.76;
const TYPESET_RETRY_MS = 350;

let measurer: CanvasRenderingContext2D | null = null;
/** The width (bp) of text drawn as pending text is (a serif face at `size` bp). */
export function textWidth(text: string, size: number): number {
  measurer ??= document.createElement("canvas").getContext("2d");
  if (!measurer) return text.length * size * 0.5;
  measurer.font = `${size}px ${PENDING_FONT}`;
  return measurer.measureText(text).width;
}
export const PENDING_FONT = '"Songti SC", "STSong", "SimSun", "Noto Serif CJK SC", "Times New Roman", serif';

export class VisualEditor {
  units: VUnit[] = [];
  active: VUnit | null = null;
  caret: Caret | null = null;
  composing = "";
  formula: OpenFormula | null = null;
  menu: CommandMenuState | null = null;
  /** the unit the pointer is selecting in */
  private dragging: VUnit | null = null;
  /** a click on a unit whose glyphs were not known yet: the caret goes there when they are */
  private pendingClick: { unit: VUnit; page: number; x: number; y: number } | null = null;
  private listeners = new Set<() => void>();
  version = 0;
  private disposed = false;

  constructor(private host: VisualHost) {}

  subscribe = (fn: () => void) => {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  };
  getVersion = () => this.version;
  private notify() {
    this.version++;
    for (const fn of this.listeners) fn();
  }
  /** The workspace is gone (answers that come after are dropped); revive: it is there again (React mounts twice). */
  dispose() {
    this.disposed = true;
  }
  revive() {
    this.disposed = false;
  }

  // ---- what is drawn ----

  overlays(): LiveOverlay[] {
    return this.units.map((u) => u.overlay).filter((o): o is LiveOverlay => !!o);
  }

  /** How far the overlays above a unit moved it on its page. */
  private shiftOf(u: VUnit): number {
    const l = u.layout;
    if (!l || !l.lines.length) return 0;
    return shiftAt(
      this.overlays().filter((o) => o !== u.overlay),
      l.page,
      l.lines[0].top,
    );
  }

  /** The caret's offset in the text a stale layout was made from (the edit since then not typeset yet). */
  private inLayout(u: VUnit, offset: number): number {
    const shown = u.shown;
    const text = this.host.textOf(u.path);
    if (shown === null || text === undefined) return offset;
    const now = text.slice(u.start, u.end);
    if (now === shown) return offset;
    // the common head of both: an offset after it stays at its end until the new layout comes
    let p = 0;
    while (p < shown.length && p < now.length && shown[p] === now[p]) p++;
    const rel = offset - u.start;
    if (rel <= p) return offset;
    let s = 0;
    while (s < shown.length - p && s < now.length - p && shown[shown.length - 1 - s] === now[now.length - 1 - s]) s++;
    if (rel >= now.length - s) return u.start + rel - (now.length - shown.length);
    return u.start + p;
  }

  caretView(): CaretView | null {
    const c = this.caret;
    const u = this.active;
    if (!c || !u?.layout || this.formula) return null;
    const stop = stopOf(u.layout, this.inLayout(u, c.offset), c.line);
    if (!stop) return null;
    const line = u.layout.lines[stop.line];
    const shift = this.shiftOf(u);
    // after text typed into a unit TeX shows only with the next build: after that text as it is drawn for now
    const pending = this.pendingOf(u);
    let x = stop.x;
    if (pending && c.offset >= pending.at && c.offset <= pending.at + pending.text.length && pending.stop.line === stop.line) {
      x = pending.stop.x + textWidth(pending.text.slice(0, c.offset - pending.at), pending.size);
    }
    return { page: u.layout.page, x, top: line.top + shift, bottom: line.bottom + shift };
  }

  /** A unit the live server does not typeset, edited: what was typed (shown at its place until the next build) and
   *  what was deleted (its glyphs covered). */
  private pendingOf(u: VUnit): { at: number; text: string; stop: { x: number; line: number }; size: number; covered: Array<{ x: number; y: number; w: number; h: number }> } | null {
    if (u.live !== false || !u.layout || u.shown === null) return null;
    const text = this.host.textOf(u.path);
    if (text === undefined) return null;
    const now = text.slice(u.start, u.end);
    const shown = u.shown;
    if (now === shown) return null;
    let p = 0;
    while (p < shown.length && p < now.length && shown[p] === now[p]) p++;
    let s = 0;
    while (s < shown.length - p && s < now.length - p && shown[shown.length - 1 - s] === now[now.length - 1 - s]) s++;
    const inserted = now.slice(p, now.length - s);
    const stop = stopOf(u.layout, u.start + p);
    if (!stop) return null;
    const covered = selectionRects(u.layout, u.start + p, u.start + shown.length - s);
    return { at: u.start + p, text: inserted, stop, size: this.sizeAt(u, u.start + p), covered };
  }

  /** The pending text and covers of the units on page `page` (drawn over the page until the next build). */
  pendingView(page: number): Array<{ x: number; top: number; bottom: number; text: string; size: number; covered: Array<{ x: number; y: number; w: number; h: number }> }> {
    const out = [];
    for (const u of this.units) {
      const l = u.layout;
      if (!l || l.page !== page) continue;
      const pending = this.pendingOf(u);
      if (!pending) continue;
      const line = l.lines[pending.stop.line];
      const shift = this.shiftOf(u);
      out.push({
        x: pending.stop.x,
        top: line.top + shift,
        bottom: line.bottom + shift,
        text: pending.text,
        size: pending.size,
        covered: pending.covered.map((r) => ({ ...r, y: r.y + shift })),
      });
    }
    return out;
  }

  selectionView(): Array<{ page: number; x: number; y: number; w: number; h: number }> {
    const c = this.caret;
    const u = this.active;
    if (!c || !u?.layout || c.anchor === c.offset) return [];
    const shift = this.shiftOf(u);
    return selectionRects(u.layout, this.inLayout(u, c.anchor), this.inLayout(u, c.offset)).map((r) => ({ page: u.layout!.page, ...r, y: r.y + shift }));
  }

  /** The font size at the caret (bp): its line's glyphs'. */
  private sizeAt(u: VUnit, offset: number): number {
    const l = u.layout;
    if (!l) return 10;
    const stop = stopOf(l, this.inLayout(u, offset));
    const on = l.items.filter((it) => !stop || it.line === stop.line).filter((it) => it.math !== true);
    const sizes = (on.length ? on : l.items).map((it) => it.size).sort((a, b) => a - b);
    return sizes[Math.floor(sizes.length / 2)] ?? 10;
  }

  // ---- units ----

  /** The source moved under the units (an edit of `path`). */
  changed(path: string, changes: ChangeSet) {
    for (const u of this.units) {
      if (u.path !== path) continue;
      u.start = changes.mapPos(u.start, -1);
      u.end = changes.mapPos(u.end, 1);
    }
    const f = this.formula;
    if (f && f.unit.path === path) {
      if (f.range) f.range = [changes.mapPos(f.range[0], -1), changes.mapPos(f.range[1], 1)];
      f.at = changes.mapPos(f.at, -1);
    }
    const c = this.caret;
    if (c && c.path === path && !this.editing) {
      c.offset = changes.mapPos(c.offset, 1);
      c.anchor = changes.mapPos(c.anchor, 1);
    }
  }

  /** A new build is shown: the units start again from it (the active one typeset anew). */
  newBuild() {
    const built = this.host.built();
    if (!built) return;
    const keep = this.active;
    this.units = keep ? [keep] : [];
    for (const u of this.units) {
      // the caret's paragraph as the source has it now (an Enter made two)
      const text = this.host.textOf(u.path);
      const c = this.caret;
      if (text !== undefined && c && c.path === u.path && !this.formula) [u.start, u.end] = visualUnit(text, Math.min(c.offset, text.length), this.host.macros());
      u.build = built.build;
      u.bStart = this.host.toBuilt(u.path, u.start, -1);
      u.bEnd = this.host.toBuilt(u.path, u.end, 1);
      u.overlay = null;
      u.live = null;
      void this.typeset(u);
    }
    this.host.overlaysChanged();
    this.notify();
  }

  /** The unit of the source around offset `pos` of `path` (its text now). */
  private unitAt(path: string, pos: number): VUnit | null {
    const text = this.host.textOf(path);
    const built = this.host.built();
    if (text === undefined || !built) return null;
    const existing = this.units.find((u) => u.path === path && pos >= u.start && pos <= u.end);
    if (existing) return existing;
    const [start, end] = visualUnit(text, pos, this.host.macros());
    const u: VUnit = {
      path,
      start,
      end,
      build: built.build,
      bStart: this.host.toBuilt(path, start, -1),
      bEnd: this.host.toBuilt(path, end, 1),
      layout: null,
      shown: null,
      live: null,
      overlay: null,
      busy: false,
      again: false,
    };
    // units edited before stay (their overlays) until the next build; ones only visited go
    this.units = this.units.filter((x) => x.overlay || x === this.active);
    this.units.push(u);
    return u;
  }

  /** The unit's text and its old text, as a change the live server is asked for. */
  private changeOf(u: VUnit): { text: string; change: ParagraphChange } | null {
    const text = this.host.textOf(u.path);
    const old = this.host.builtText(u.path);
    if (text === undefined || old === undefined) return null;
    return {
      text,
      change: {
        oldText: old.slice(u.bStart, u.bEnd),
        oldFirst: lineNumberAt(old, u.bStart),
        oldLast: lineNumberAt(old, u.bEnd),
        start: u.start,
        newText: text.slice(u.start, u.end),
      },
    };
  }

  /** The unit typeset (its overlay, its caret places): by the live server, else from the page. */
  async typeset(u: VUnit): Promise<void> {
    if (u.busy) {
      u.again = true;
      return;
    }
    u.busy = true;
    try {
      do {
        u.again = false;
        const at = this.changeOf(u);
        if (!at) return;
        const { text, change } = at;
        const f = this.formula?.unit === u ? this.formula : null;
        const out = await this.host.typeset(u.path, text, change, { transform: f ? (p) => this.withHole(u, f, p) : undefined });
        if (this.disposed) return;
        if (out.ok) {
          // (the caret's places where the page draws the paragraph)
          const { items, lines } = liveItems(out.result.map ?? [], out.result.lines, out.x0 + out.tx, out.y0 + out.ty);
          // (what was typeset: the text then; the source has the formula itself where its room was)
          u.layout = layoutOf(out.page, text, change.start, change.start + change.newText.length, items, lines, this.host.commands());
          u.shown = change.newText;
          u.live = true;
          u.overlay = out.overlay;
          if (f) {
            const hole = out.result.marks?.find((m) => m[0] === "hole");
            if (hole) f.place = { page: out.page, x: out.x0 + out.tx + hole[1], y: out.y0 + out.ty + hole[2] };
          }
        } else if (out.unfinished) {
          // a command half typed: what was shown stays
        } else if (!out.retry && !u.narrowed && this.narrow(u)) {
          // a paragraph with displays that cannot be typeset as one: the part the caret is in, as before
          u.again = true;
        } else {
          if (out.retry && !u.again) {
            window.setTimeout(() => void this.typeset(u), TYPESET_RETRY_MS);
          }
          u.live = out.retry ? null : false;
          if (change.oldText === change.newText || !u.layout) await this.fromPage(u, change);
        }
        const click = this.pendingClick;
        if (click && click.unit === u && u.layout && u.layout.page === click.page) {
          this.pendingClick = null;
          const yy = click.y - this.shiftOf(u);
          const offset = offsetAt(u.layout, click.x, yy);
          if (offset !== null && this.active === u) {
            this.caret = { path: u.path, offset, line: stopOf(u.layout, offset)?.line, anchor: offset };
            this.flushTyped();
          }
        }
        this.host.overlaysChanged();
        this.notify();
      } while (u.again);
    } finally {
      u.busy = false;
    }
  }

  /** A unit made of text and formulas shown apart, narrowed to the text part the caret is in (displays split it
   *  off); false when that is the same unit. */
  private narrow(u: VUnit): boolean {
    const text = this.host.textOf(u.path);
    const c = this.caret;
    if (text === undefined) return false;
    const at = c && c.path === u.path && c.offset >= u.start && c.offset <= u.end ? c.offset : u.start;
    const [start, end] = visualUnit(text, at, this.host.macros(), true);
    u.narrowed = true;
    if (start === u.start && end === u.end) return false;
    u.start = start;
    u.end = end;
    u.bStart = this.host.toBuilt(u.path, start, -1);
    u.bEnd = this.host.toBuilt(u.path, end, 1);
    u.layout = null;
    u.shown = null;
    return true;
  }

  /** The caret's places from the page's own glyphs (a unit the live server does not typeset alone). */
  private async fromPage(u: VUnit, change: ParagraphChange): Promise<void> {
    const s = await this.host.sync();
    if (!s) return;
    const files = s.files.map((f, i) => (f === u.path ? i : -1)).filter((i) => i >= 0);
    const area = paragraphArea(s.tex, files, change.oldFirst, change.oldLast);
    let page: number;
    let lines: LayoutLine[];
    let chars: ReturnType<typeof pageChars>;
    if (area) {
      const svg = this.host.pageSvg(area.page);
      if (!svg) return;
      page = area.page;
      chars = pageChars(svg);
      lines = area.lines.slice(area.core[0], area.core[1] + 1).map((b) => ({
        y: b.v,
        top: b.v - b.height,
        bottom: b.v + b.depth,
        left: b.h,
        right: b.h + b.width,
      }));
    } else {
      // no line boxes of its own (a frame's, a list's in a box): the baselines its source lines' records are on, and
      // the page's glyphs there
      const loose = this.looseLines(s.tex, files, change.oldFirst, change.oldLast);
      if (!loose) return;
      page = loose.page;
      chars = loose.chars;
      lines = loose.lines;
    }
    const placed = pageItems(chars, lines);
    const items = placed.items;
    lines = movedLines(lines, placed.tx, placed.ty);
    const text = this.host.textOf(u.path) ?? "";
    // (the page shows the build's text: the layout is of it, its offsets moved to the text now)
    const old = this.host.builtText(u.path) ?? "";
    const layout = layoutOf(page, old, u.bStart, u.bEnd, items, lines, this.host.commands());
    const delta = u.start - u.bStart;
    if (delta) {
      for (const unit of layout.units) {
        unit.from += delta;
        unit.to += delta;
        if (unit.inner) unit.inner = [unit.inner[0] + delta, unit.inner[1] + delta];
      }
      for (const stop of layout.stops) stop.offset += delta;
      layout.from += delta;
      layout.to += delta;
    }
    u.layout = layout;
    u.shown = change.oldText;
    void text;
  }

  /** The lines a unit's source lines [first, last] made on a page, from SyncTeX's records of them alone (no line
   *  box needed): their baselines, each line as far as the page's glyphs on it go near the records. */
  private looseLines(tex: SyncTex, files: number[], first: number, last: number): { page: number; lines: LayoutLine[]; chars: ReturnType<typeof pageChars> } | null {
    let best: { page: number; recs: SyncTex["pages"] extends Map<number, infer P> ? (P extends { records: infer R } ? R : never) : never } | null = null;
    for (const [n, p] of tex.pages) {
      const recs = p.records.filter((r) => files.includes(r.file) && r.line >= first && r.line <= last && "kg$x".includes(r.kind));
      if (recs.length && (!best || recs.length > best.recs.length)) best = { page: n, recs };
    }
    if (!best) return null;
    const svg = this.host.pageSvg(best.page);
    if (!svg) return null;
    const chars = pageChars(svg);
    const baselines: Array<{ v: number; left: number; right: number }> = [];
    for (const r of best.recs) {
      const b = baselines.find((x) => Math.abs(x.v - r.v) < 0.6);
      if (b) {
        b.left = Math.min(b.left, r.h);
        b.right = Math.max(b.right, r.h + Math.max(0, r.width));
      } else baselines.push({ v: r.v, left: r.h, right: r.h + Math.max(0, r.width) });
    }
    const lines: LayoutLine[] = [];
    for (const b of baselines.sort((p, q) => p.v - q.v)) {
      const on = chars.filter((c) => Math.abs(c.y - b.v) < 0.6 && c.x >= b.left - 40 && c.x <= b.right + 40);
      if (!on.length) continue;
      const size = Math.max(...on.map((c) => c.size));
      const left = Math.min(...on.map((c) => c.x));
      const right = Math.max(...on.map((c) => c.x + c.size * 0.5));
      lines.push({ y: b.v, top: b.v - size * 0.85, bottom: b.v + size * 0.3, left, right });
    }
    return lines.length ? { page: best.page, lines, chars } : null;
  }

  /** The paragraph as sent while a formula is open: the formula's room in place of the formula. */
  private withHole(u: VUnit, f: OpenFormula, paragraph: string): string {
    const size = f.size ?? { w: f.fontSize * 0.6, h: f.fontSize * 0.7, d: f.fontSize * 0.2 };
    const sp = (v: number) => Math.max(0, Math.round(v * SP_PER_BP));
    const hole = `\\vtxhole{${sp(size.w)}}{${sp(size.h)}}{${sp(size.d)}}`;
    const room = f.display ? `\\[${hole}\\]` : `$${hole}$`;
    const [from, to] = f.range ?? [f.at, f.at];
    const a = from - u.start;
    const b = to - u.start;
    if (a < 0 || b > paragraph.length) return paragraph;
    return paragraph.slice(0, a) + room + paragraph.slice(b);
  }

  // ---- the pointer ----

  /** The unit and offset at a page point (a unit typeset already there, else the one SyncTeX finds). */
  private async locate(page: number, x: number, y: number): Promise<{ unit: VUnit; offset: number; line: number; atom: number | null } | null> {
    const hit = (u: VUnit) => {
      const l = u.layout;
      if (!l || l.page !== page || !l.lines.length) return null;
      const shift = this.shiftOf(u);
      const top = Math.min(...l.lines.map((line) => line.top)) + shift - 2;
      const bottom = Math.max(...l.lines.map((line) => line.bottom)) + shift + 2;
      if (y < top || y > bottom) return null;
      const yy = y - shift;
      const offset = offsetAt(l, x, yy);
      if (offset === null) return null;
      const stop = stopOf(l, offset);
      return { unit: u, offset, line: stop?.line ?? 0, atom: atomAt(l, x, yy) };
    };
    for (const u of this.units) {
      const found = hit(u);
      if (found) return found;
    }
    const s = await this.host.sync();
    if (!s) return null;
    const overlays = this.overlays();
    const yb = y - shiftAt(overlays, page, y);
    // the page draws some boxes moved from TeX's coordinates (a tcolorbox's content): SyncTeX is asked at the point
    // in TeX's coordinates, as the nearest glyph on the page is moved
    const [tx, ty] = this.moveNear(page, x, yb);
    const where = s.tex.edit(page, x - tx, yb - ty, (f) => !!s.files[f]);
    if (!where) return null;
    const path = s.files[where.file];
    if (!path) return null;
    if (!(await this.host.openFile(path))) return null;
    const old = this.host.builtText(path);
    if (old === undefined) return null;
    // the line's start in the build's text, then in the text now
    let pos = 0;
    for (let n = 1; n < where.line && pos >= 0; n++) pos = old.indexOf("\n", pos) + 1;
    if (pos < 0) return null;
    const u = this.unitAt(path, this.host.toCurrent(path, pos, 1));
    if (!u) return null;
    if (!u.layout) await this.typeset(u);
    let near = hit(u);
    const trace: unknown[] = [];
    const note = (unit: VUnit) => trace.push({ start: unit.start, end: unit.end, page: unit.layout?.page, lines: unit.layout?.lines.map((l) => [+l.top.toFixed(1), +l.bottom.toFixed(1)]), live: unit.live });
    note(u);
    // SyncTeX's line may be a neighbour's (a list in a frame, the end of a paragraph): the units around it, the one
    // whose glyphs the click is on
    for (let k = 0; !near && u.layout && k < 4; k++) {
      const other = this.neighbour(u, k % 2 ? 1 : -1, k < 2 ? 1 : 2);
      if (!other || other === u) continue;
      if (!other.layout) await this.typeset(other);
      note(other);
      near = hit(other);
    }
    if (import.meta.env.DEV) (window as unknown as { __latexLocate?: unknown }).__latexLocate = { page, x, y, yb, line: where.line, near: near && { start: near.unit.start, offset: near.offset }, trace };
    // (a click the units around do not hold: the caret to the nearest place of SyncTeX's unit — never into a formula)
    const found = near ?? (u.layout ? { unit: u, offset: offsetAt(u.layout, x, yb) ?? u.start, line: 0, atom: null } : null);
    if (!found && !u.layout) {
      // its glyphs are not known yet (the live server starting): the caret goes where the click was once they are
      this.active = u;
      this.pendingClick = { unit: u, page, x, y };
    }
    return found;
  }

  private charsCache: { svg: string; chars: ReturnType<typeof pageChars> } | null = null;

  /** How far the page draws its glyphs near a page point from TeX's coordinates (the nearest glyph's move). */
  private moveNear(page: number, x: number, y: number): [number, number] {
    const svg = this.host.pageSvg(page);
    if (!svg) return [0, 0];
    if (this.charsCache?.svg !== svg) this.charsCache = { svg, chars: pageChars(svg) };
    let best: { tx: number; ty: number } | null = null;
    let dist = 24;
    for (const c of this.charsCache.chars) {
      const d = Math.hypot(c.x + c.tx - x, (c.y + c.ty - c.size * 0.35) - y);
      if (d < dist) {
        dist = d;
        best = c;
      }
    }
    return best ? [best.tx, best.ty] : [0, 0];
  }

  /** The unit `steps` units before (dir -1) or after (dir 1) `u` in its file (blank lines and lines that show
   *  nothing skipped), null at the file's start or end. */
  private neighbour(u: VUnit, dir: -1 | 1, steps: number): VUnit | null {
    const text = this.host.textOf(u.path);
    if (text === undefined) return null;
    let from = u.start;
    let to = u.end;
    for (let k = 0; k < steps; k++) {
      let pos = dir > 0 ? to : from - 1;
      if (dir > 0) while (pos < text.length && /\s/.test(text[pos])) pos++;
      else while (pos > 0 && /\s/.test(text[pos])) pos--;
      if (pos < 0 || pos >= text.length) return null;
      [from, to] = visualUnit(text, pos, this.host.macros());
    }
    const existing = this.units.find((x) => x.path === u.path && x.start === from && x.end === to);
    if (existing) return existing;
    const built = this.host.built();
    if (!built) return null;
    const unit: VUnit = { path: u.path, start: from, end: to, build: built.build, bStart: this.host.toBuilt(u.path, from, -1), bEnd: this.host.toBuilt(u.path, to, 1), layout: null, shown: null, live: null, overlay: null, busy: false, again: false };
    this.units.push(unit);
    return unit;
  }

  /** The caret to a source offset of a file (the source's own place: a message's line, the tests). */
  async placeAt(path: string, offset: number): Promise<boolean> {
    if (!(await this.host.openFile(path))) return false;
    const u = this.unitAt(path, offset);
    if (!u) return false;
    this.active = u;
    this.caret = { path, offset, anchor: offset };
    this.keyMoved = true;
    // its page in view (loaded: the paragraph is checked against it)
    this.host.revealSource(path, offset);
    if (!u.layout) await this.typeset(u);
    this.notify();
    return true;
  }

  /** the caret was moved by a key (the view follows it) */
  keyMoved = false;

  private pointerToken = 0;

  /** The pointer pressed at a page point (`client`: the same point on the screen, where a formula opened there
   *  puts its caret). */
  async pointerDown(page: number, x: number, y: number, shift: boolean, client?: { x: number; y: number }) {
    const token = ++this.pointerToken;
    this.keyMoved = false;
    if (this.formula) this.closeFormula("click");
    this.menu = null;
    // while the click is being placed (a paragraph's glyphs asked for): what is typed waits for the caret there
    this.locating = true;
    let found: Awaited<ReturnType<VisualEditor["locate"]>>;
    try {
      found = await this.locate(page, x, y);
    } finally {
      if (token === this.pointerToken) this.locating = false;
    }
    if (token !== this.pointerToken) return;
    if (!found) {
      if (!this.pendingClick) this.host.cannotEdit();
      this.notify();
      return;
    }
    this.pendingClick = null;
    const { unit, offset, line, atom } = found;
    const math = atom !== null ? unit.layout?.units[atom] : null;
    if (math && math.kind === "math" && !shift) {
      this.active = unit;
      this.openFormula(unit, atom!, client ? { kind: "point", x: client.x, y: client.y } : { kind: "end" });
      return;
    }
    const keep = shift && this.caret && this.caret.path === unit.path && this.active === unit;
    this.active = unit;
    this.caret = { path: unit.path, offset, line, anchor: keep ? this.caret!.anchor : offset };
    this.dragging = unit;
    this.notify();
    this.flushTyped();
  }

  /** a click is being placed: typing waits */
  locating = false;
  private typedWhileLocating: string[] = [];

  /** What was typed while a click was being placed, typed at its caret now. */
  private flushTyped() {
    const typed = this.typedWhileLocating;
    this.typedWhileLocating = [];
    for (const t of typed) this.insertText(t);
  }

  /** A double click: the word at the caret (letters and digits, or a run of CJK characters, between spaces and
   *  punctuation); a triple click: the whole unit. */
  selectWord(count: number) {
    const c = this.caret;
    const u = this.active;
    if (!c || !u?.layout) return;
    if (count >= 3) {
      this.selectAll();
      return;
    }
    const { text, units } = this.unitsNow(u);
    const kind = (unit: { kind: string; text?: string }) => {
      if (unit.kind !== "char") return unit.kind === "math" ? "math" : "stop";
      const ch = unit.text ?? "";
      if (/[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}]/u.test(ch)) return "cjk";
      if (/[\p{L}\p{N}_']/u.test(ch)) return "word";
      return "stop";
    };
    // the visible unit at the caret: the one starting at it, else the one ending at it
    let k = units.findIndex((x) => x.kind !== "mark" && x.from >= c.offset);
    if (k < 0 || kind(units[k]) === "stop") {
      const before = [...units.keys()].reverse().find((i) => units[i].kind !== "mark" && units[i].to <= c.offset);
      if (before !== undefined) k = before;
    }
    if (k < 0) return;
    const want = kind(units[k]);
    if (want === "stop") return;
    let a = k;
    let b = k;
    const same = (i: number) => units[i].kind === "mark" || kind(units[i]) === want;
    while (a > 0 && same(a - 1) && want !== "math") a--;
    while (b < units.length - 1 && same(b + 1) && want !== "math") b++;
    while (units[a].kind === "mark" && a < b) a++;
    while (units[b].kind === "mark" && b > a) b--;
    c.anchor = units[a].from;
    c.offset = units[b].to;
    c.line = undefined;
    void text;
    this.notify();
  }

  pointerMove(page: number, x: number, y: number) {
    const u = this.dragging;
    const c = this.caret;
    if (!u?.layout || !c || u.layout.page !== page) return;
    const yy = y - this.shiftOf(u);
    const offset = offsetAt(u.layout, x, yy);
    if (offset === null || offset === c.offset) return;
    c.offset = offset;
    c.line = stopOf(u.layout, offset)?.line;
    this.notify();
  }

  pointerUp() {
    this.dragging = null;
  }

  // ---- editing ----

  /** an edit of ours is going through the source (the caret is set by it, not moved by the change) */
  private editing = false;

  private edit(changes: Array<{ from: number; to: number; insert: string }>, caret: number, event: Parameters<LatexSourceHandle["edit"]>[3]): boolean {
    this.keyMoved = true;
    const u = this.active;
    const c = this.caret;
    const source = this.host.source();
    if (!u || !c || !source) return false;
    this.editing = true;
    try {
      if (!source.edit(u.path, changes, caret, event)) return false;
    } finally {
      this.editing = false;
    }
    c.offset = caret;
    c.anchor = caret;
    c.goalX = undefined;
    c.line = undefined;
    this.host.scheduleCompile(u.live === true);
    void this.typeset(u);
    this.notify();
    return true;
  }

  private selection(): [number, number] | null {
    const c = this.caret;
    if (!c || c.anchor === c.offset) return null;
    return c.anchor < c.offset ? [c.anchor, c.offset] : [c.offset, c.anchor];
  }

  /** The unit's source as the caret knows it now (the text now, not the one its layout was made from). */
  private unitsNow(u: VUnit) {
    const text = this.host.textOf(u.path) ?? "";
    return { text, units: tokenize(text, u.start, Math.min(u.end, text.length), this.host.commands()) };
  }

  /** The changes that delete the selection (the markup around it kept) and where the caret goes. */
  private deletion(): { changes: Array<{ from: number; to: number; insert: string }>; at: number } | null {
    const sel = this.selection();
    const u = this.active;
    if (!sel || !u) return null;
    const ranges = deletionRanges(this.unitsNow(u).units, sel[0], sel[1]);
    if (!ranges.length) return { changes: [], at: sel[0] };
    return { changes: ranges.map(([from, to]) => ({ from, to, insert: "" })), at: ranges[0][0] };
  }

  /** The selection's deletions with `insert` put where the caret goes (one change there, not an insertion beside a
   *  deletion at the same place), and the caret after the insertion. */
  private replaceSelection(insert: string): { changes: Array<{ from: number; to: number; insert: string }>; caret: number } | null {
    const c = this.caret;
    if (!c) return null;
    const del = this.deletion();
    if (!del || !del.changes.length) return { changes: [{ from: c.offset, to: c.offset, insert }], caret: c.offset + insert.length };
    const changes = del.changes.map((d) => ({ ...d }));
    const first = changes.find((d) => d.from === del.at);
    if (first) first.insert = insert;
    else changes.push({ from: del.at, to: del.at, insert });
    changes.sort((a, b) => a.from - b.from);
    const removedBefore = changes.filter((d) => d.from < del.at).reduce((n, d) => n + (d.to - d.from) - d.insert.length, 0);
    return { changes, caret: del.at - removedBefore + insert.length };
  }

  /** Text typed (or pasted) at the caret: in the source as LaTeX writes it. */
  insertText(typed: string, raw = false) {
    if (this.locating) {
      this.typedWhileLocating.push(typed);
      return;
    }
    if (!this.caret || !this.active) return;
    if (!raw && typed === "$") {
      this.newFormula(false);
      return;
    }
    if (!raw && typed === "\\") {
      this.openMenu();
      return;
    }
    const insert = raw ? typed : escapeTyped(typed);
    const r = this.replaceSelection(insert);
    if (r) this.edit(r.changes, r.caret, raw ? "input" : "input.type");
  }

  /** Backspace (dir -1) or Delete (dir 1). */
  deleteChar(dir: -1 | 1) {
    const c = this.caret;
    const u = this.active;
    if (!c || !u?.layout) return;
    const del = this.deletion();
    if (del) {
      if (del.changes.length) this.edit(del.changes, del.at, "delete.selection");
      return;
    }
    const { text, units } = this.unitsNow(u);
    const at = c.offset;
    // the unit before (after) the caret, past markup
    let k = -1;
    if (dir < 0) {
      for (let i = units.length - 1; i >= 0; i--) {
        if (units[i].to <= at && units[i].kind !== "mark") {
          k = i;
          break;
        }
      }
    } else {
      for (let i = 0; i < units.length; i++) {
        if (units[i].from >= at && units[i].kind !== "mark") {
          k = i;
          break;
        }
      }
    }
    if (k < 0) {
      this.joinUnits(dir);
      return;
    }
    const unit = units[k];
    if (unit.kind === "math" || unit.kind === "atom") {
      // a formula or a command: selected first, deleted by the next key
      c.anchor = dir < 0 ? unit.to : unit.from;
      c.offset = dir < 0 ? unit.from : unit.to;
      c.line = undefined;
      this.notify();
      return;
    }
    let [from, to] = [unit.from, unit.to];
    // a text command left empty goes too
    const after = text.slice(0, from) + text.slice(to);
    const empty = emptyGroupsAround(after, from);
    if (empty.length) {
      const [e0, e1] = empty[0];
      // (its range in the text before the deletion)
      const a0 = e0 >= from ? e0 + (to - from) : e0;
      const a1 = e1 > from ? e1 + (to - from) : e1;
      from = Math.min(from, a0);
      to = Math.max(to, a1);
    }
    this.edit([{ from, to, insert: "" }], from, dir < 0 ? "delete.backward" : "delete.forward");
  }

  /** Backspace at a paragraph's start: it joins the one before (Delete at its end: the one after). */
  private joinUnits(dir: -1 | 1) {
    const u = this.active;
    const c = this.caret;
    if (!u || !c) return;
    const text = this.host.textOf(u.path) ?? "";
    if (dir < 0) {
      let p = u.start;
      while (p > 0 && /\s/.test(text[p - 1])) p--;
      if (p === u.start || p === 0) return;
      // only blank lines between them: one line end joins them
      if (!/\n\s*\n/.test(text.slice(p, u.start))) return;
      this.edit([{ from: p, to: u.start, insert: "\n" }], p + 1, "delete.backward");
      // the joined paragraph: the unit grows over the one before
      const [start, end] = visualUnit(this.host.textOf(u.path) ?? "", p, this.host.macros());
      this.retarget(u, start, end);
    } else {
      let p = u.end;
      while (p < text.length && /\s/.test(text[p])) p++;
      if (p === u.end || p >= text.length) return;
      if (!/\n\s*\n/.test(text.slice(u.end, p))) return;
      this.edit([{ from: u.end, to: p, insert: "\n" }], c.offset, "delete.forward");
      const [start, end] = visualUnit(this.host.textOf(u.path) ?? "", u.start, this.host.macros());
      this.retarget(u, start, end);
    }
  }

  /** The unit is now [start, end) of its file (paragraphs joined or split): its old text is the build's there. */
  private retarget(u: VUnit, start: number, end: number) {
    u.start = start;
    u.end = end;
    u.bStart = this.host.toBuilt(u.path, start, -1);
    u.bEnd = this.host.toBuilt(u.path, end, 1);
    void this.typeset(u);
  }

  /** Enter: a new paragraph (a new item in a list). Shift+Enter: a line break. */
  newParagraph(lineBreak: boolean) {
    const u = this.active;
    const c = this.caret;
    if (!u || !c) return;
    if (lineBreak) {
      this.insertText("\\\\\n", true);
      return;
    }
    const text = this.host.textOf(u.path) ?? "";
    const item = /^\s*\\item(?![A-Za-z@])/.test(text.slice(u.start, u.end));
    const indent = /^[ \t]*/.exec(text.slice(text.lastIndexOf("\n", u.start - 1) + 1))?.[0] ?? "";
    const insert = item ? `\n${indent}\\item ` : "\n\n";
    const r = this.replaceSelection(insert);
    // the old unit is typeset as one, both paragraphs (what follows moves down with them); the next build makes
    // them two units again
    if (r) this.edit(r.changes, r.caret, "input");
  }

  // ---- the caret moving ----

  move(dir: -1 | 1, extend: boolean) {
    this.keyMoved = true;
    const c = this.caret;
    const u = this.active;
    if (!c || !u?.layout) return;
    const sel = this.selection();
    if (sel && !extend) {
      c.offset = c.anchor = dir < 0 ? sel[0] : sel[1];
      c.line = undefined;
      this.notify();
      return;
    }
    const stop = stepStop(u.layout, this.inLayout(u, c.offset), c.line, dir);
    if (!stop) {
      void this.crossUnit(dir, extend, null);
      return;
    }
    c.offset = stop.offset;
    c.line = stop.line;
    c.goalX = undefined;
    if (!extend) c.anchor = c.offset;
    this.notify();
  }

  moveLine(dir: -1 | 1, extend: boolean) {
    this.keyMoved = true;
    const c = this.caret;
    const u = this.active;
    if (!c || !u?.layout) return;
    const here = stopOf(u.layout, this.inLayout(u, c.offset), c.line);
    if (!here) return;
    const x = c.goalX ?? here.x;
    const target = lineStop(u.layout, here.line + dir, x);
    if (!target) {
      void this.crossUnit(dir, extend, x);
      return;
    }
    c.offset = target.offset;
    c.line = target.line;
    c.goalX = x;
    if (!extend) c.anchor = c.offset;
    this.notify();
  }

  lineEdge(dir: -1 | 1, extend: boolean) {
    this.keyMoved = true;
    const c = this.caret;
    const u = this.active;
    if (!c || !u?.layout) return;
    const here = stopOf(u.layout, this.inLayout(u, c.offset), c.line);
    if (!here) return;
    const edges = lineEdges(u.layout, here.line);
    if (!edges) return;
    const stop = dir < 0 ? edges[0] : edges[1];
    c.offset = stop.offset;
    c.line = stop.line;
    if (!extend) c.anchor = c.offset;
    this.notify();
  }

  selectAll() {
    const c = this.caret;
    const u = this.active;
    if (!c || !u?.layout?.stops.length) return;
    const s = u.layout.stops;
    c.anchor = s[0].offset;
    c.offset = s[s.length - 1].offset;
    this.notify();
  }

  /** Past the unit's first/last place: the unit before/after on the page (up/down at `x`; left/right its end). */
  private async crossUnit(dir: -1 | 1, extend: boolean, x: number | null) {
    const u = this.active;
    const c = this.caret;
    if (!u?.layout || !c || extend) return;
    const lines = u.layout.lines;
    const shift = this.shiftOf(u);
    const edge = dir > 0 ? Math.max(...lines.map((l) => l.bottom)) : Math.min(...lines.map((l) => l.top));
    const lineHeight = Math.max(4, (lines[0].bottom - lines[0].top) * 0.9);
    let page = u.layout.page;
    let px = x ?? (dir > 0 ? lines[0].left + 1 : lines[lines.length - 1].right - 1);
    // the first text below (above) the unit: step down (up) until something else is there
    for (let k = 1; k <= 40; k++) {
      let py = edge + shift + dir * (k * lineHeight * 0.5);
      if (py > this.host.pageHeight(page) || py < 0) {
        const next = page + dir;
        if (next < 1 || next > this.host.pageCount()) return;
        page = next;
        py = dir > 0 ? 40 : this.host.pageHeight(page) - 40;
      }
      const found = await this.locate(page, px, py);
      if (!found || found.unit === u) continue;
      const l = found.unit.layout;
      if (!l?.stops.length) continue;
      let stop = null;
      if (x !== null) {
        const line = dir > 0 ? 0 : l.lines.length - 1;
        stop = lineStop(l, line, x);
      }
      stop ??= dir > 0 ? l.stops[0] : l.stops[l.stops.length - 1];
      this.active = found.unit;
      this.caret = { path: found.unit.path, offset: stop.offset, line: stop.line, anchor: stop.offset, goalX: x ?? undefined };
      this.notify();
      return;
    }
    void px;
  }

  // ---- history ----

  history(dir: "undo" | "redo") {
    const u = this.active;
    const source = this.host.source();
    if (!u || !source) return;
    this.editing = true;
    let head: number | null;
    try {
      head = source.history(u.path, dir);
    } finally {
      this.editing = false;
    }
    if (head === null) return;
    // the unit the change was in
    const text = this.host.textOf(u.path) ?? "";
    const [start, end] = visualUnit(text, Math.min(head, text.length), this.host.macros());
    const unit = this.units.find((x) => x.path === u.path && head! >= x.start && head! <= x.end) ?? this.unitAt(u.path, head);
    if (!unit) return;
    if (unit.start !== start || unit.end !== end) {
      unit.start = start;
      unit.end = end;
      unit.bStart = this.host.toBuilt(unit.path, start, -1);
      unit.bEnd = this.host.toBuilt(unit.path, end, 1);
    }
    this.active = unit;
    this.caret = { path: unit.path, offset: head, anchor: head };
    this.host.scheduleCompile(unit.live === true);
    void this.typeset(unit);
    this.notify();
  }

  // ---- clipboard ----

  /** The selection as LaTeX (its markup kept when it is all inside) and as plain text. */
  copyText(): { latex: string; plain: string } | null {
    const sel = this.selection();
    const u = this.active;
    if (!sel || !u?.layout) return null;
    const text = this.host.textOf(u.path) ?? "";
    const latex = text.slice(sel[0], sel[1]);
    const plain = u.layout.units
      .filter((x) => x.from >= sel[0] && x.to <= sel[1])
      .map((x) => (x.kind === "char" ? text.slice(x.from, x.to).replace(/^\\(?=[%&#$_{}])/, "") : x.kind === "space" || x.kind === "par" || x.kind === "break" ? " " : x.kind === "math" ? text.slice(x.from, x.to) : ""))
      .join("");
    return { latex: balancedOr(latex, plain), plain };
  }

  cut(): { latex: string; plain: string } | null {
    const out = this.copyText();
    if (out) this.deleteChar(-1);
    return out;
  }

  paste(text: string) {
    // LaTeX as it is (a formula, a command pasted works as written)
    this.insertText(text.replace(/\r\n?/g, "\n"), true);
  }

  // ---- text commands ----

  /** Mod+B, Mod+I: the selection in \textbf{…} (\textit), or out of it when it is all one; no selection: an empty
   *  one at the caret, the caret in it. */
  wrap(command: string) {
    const c = this.caret;
    const u = this.active;
    if (!c || !u) return;
    const text = this.host.textOf(u.path) ?? "";
    const sel = this.selection();
    if (!sel) {
      const open = `\\${command}{`;
      this.edit([{ from: c.offset, to: c.offset, insert: open + "}" }], c.offset + open.length, "input");
      return;
    }
    const [a, b] = sel;
    const inside = text.slice(a, b);
    const head = `\\${command}{`;
    // already wrapped exactly: unwrap
    if (text.slice(a - head.length, a) === head && text[b] === "}") {
      this.edit(
        [
          { from: a - head.length, to: a, insert: "" },
          { from: b, to: b + 1, insert: "" },
        ],
        b - head.length,
        "input",
      );
      return;
    }
    this.edit([{ from: a, to: b, insert: head + inside + "}" }], b + head.length + 1, "input");
  }

  /** What an input method is composing at the caret (drawn there until it is committed). */
  setComposing(text: string) {
    if (this.composing === text) return;
    this.composing = text;
    this.notify();
  }

  // ---- the "\" menu ----

  openMenu() {
    const c = this.caret;
    if (!c) return;
    const del = this.deletion();
    if (del?.changes.length) this.edit(del.changes, del.at, "delete.selection");
    this.menu = { query: "", index: 0, at: this.caret!.offset };
    this.notify();
  }

  menuType(text: string) {
    const m = this.menu;
    if (!m) return;
    m.query += text;
    m.index = 0;
    this.notify();
  }

  menuBack() {
    const m = this.menu;
    if (!m) return;
    if (!m.query) this.menu = null;
    else m.query = m.query.slice(0, -1);
    this.notify();
  }

  closeMenu(insertTyped: boolean) {
    const m = this.menu;
    this.menu = null;
    if (m && insertTyped) this.insertText("\\" + m.query + (/^[A-Za-z@]+$/.test(m.query) ? " " : ""), true);
    else this.notify();
  }

  /** A command chosen: its LaTeX at the caret (an argument's braces with the caret in them; a formula opened). */
  runCommand(id: string, environment?: string) {
    this.menu = null;
    const c = this.caret;
    const u = this.active;
    if (!c || !u) return;
    const text = this.host.textOf(u.path) ?? "";
    const atLineStart = /(^|\n)[ \t]*$/.test(text.slice(0, c.offset));
    const block = (body: string, caretIn: number) => {
      // a block of its own: on new lines before and after
      const before = atLineStart ? "" : "\n\n";
      const insert = before + body + "\n\n";
      this.edit([{ from: c.offset, to: c.offset, insert }], c.offset + before.length + caretIn, "input");
      this.retargetAround();
    };
    switch (id) {
      case "inline-math":
        this.newFormula(false);
        return;
      case "display-math":
        this.newFormula(true);
        return;
      case "section":
      case "subsection":
      case "subsubsection":
      case "section*":
      case "subsection*":
      case "subsubsection*":
      case "title":
      case "author":
      case "date": {
        const name = id === "title" || id === "author" || id === "date" ? id : id;
        block(`\\${name}{}`, name.length + 2);
        return;
      }
      case "itemize":
      case "enumerate":
        block(`\\begin{${id}}\n  \\item \n\\end{${id}}`, `\\begin{${id}}\n  \\item `.length);
        return;
      case "item":
        this.newParagraph(false);
        return;
      case "textbf":
      case "textit":
      case "emph":
      case "underline":
      case "texttt":
      case "sout":
        this.wrap(id);
        return;
      case "newline":
        this.insertText("\\\\\n", true);
        return;
      case "footnote":
      case "cite":
      case "href": {
        const open = id === "href" ? "\\href{" : `\\${id}{`;
        const close = id === "href" ? "}{}" : "}";
        this.edit([{ from: c.offset, to: c.offset, insert: open + close }], c.offset + open.length, "input");
        return;
      }
      default: {
        if (environment) {
          const body = `\\begin{${environment}}\n\n\\end{${environment}}`;
          block(body, `\\begin{${environment}}\n`.length);
          return;
        }
        const name = id.replace(/^size-/, "");
        this.insertText(`\\${name} `, true);
      }
    }
  }

  /** After a block was put in: the unit the caret is in now (a new one). */
  private retargetAround() {
    const c = this.caret;
    const u = this.active;
    if (!c || !u) return;
    const text = this.host.textOf(u.path) ?? "";
    const [start, end] = visualUnit(text, c.offset, this.host.macros());
    const unit = this.unitAt(u.path, c.offset);
    if (unit && (unit.start !== start || unit.end !== end)) {
      unit.start = start;
      unit.end = end;
    }
    if (unit) {
      this.active = unit;
      void this.typeset(unit);
    }
  }

  // ---- formulas ----

  /** A new formula at the caret ("$", the menu): open in the formula editor, in the source once it has something. */
  newFormula(display: boolean) {
    const c = this.caret;
    const u = this.active;
    if (!c || !u) return;
    const del = this.deletion();
    if (del?.changes.length) this.edit(del.changes, del.at, "delete.selection");
    const at = this.caret!.offset;
    this.formula = {
      unit: u,
      range: null,
      at,
      display,
      open: display ? "\\[" : "$",
      close: display ? "\\]" : "$",
      latex: "",
      entry: { kind: "end" },
      size: null,
      place: null,
      fontSize: this.sizeAt(u, at),
    };
    this.placeFromCaret();
    void this.typeset(u);
    this.notify();
  }

  /** A formula of the unit (its units' index `k`) opened in the formula editor. */
  openFormula(u: VUnit, k: number, entry: FormulaEntry) {
    const l = u.layout;
    const unit = l?.units[k];
    if (!l || !unit || unit.kind !== "math" || !unit.inner) return;
    const text = this.host.textOf(u.path) ?? "";
    const [i0, i1] = unit.inner;
    const [a, b] = l.spans[k];
    const first = l.items[a];
    this.formula = {
      unit: u,
      range: [unit.from, unit.to],
      at: unit.from,
      display: !!unit.display,
      open: text.slice(unit.from, i0),
      close: text.slice(i1, unit.to),
      latex: text.slice(i0, i1),
      entry,
      size: null,
      place: first ? { page: l.page, x: Math.min(...l.items.slice(a, b).map((it) => it.x)), y: l.lines[first.line].y + this.shiftOf(u) } : null,
      fontSize: first?.size ?? this.sizeAt(u, unit.from),
    };
    this.caret = { path: u.path, offset: unit.to, anchor: unit.to };
    void this.typeset(u);
    this.notify();
  }

  private placeFromCaret() {
    const f = this.formula;
    const u = this.active;
    const c = this.caret;
    if (!f || !u?.layout || !c) return;
    const stop = stopOf(u.layout, this.inLayout(u, f.at), c.line);
    if (!stop) return;
    f.place = { page: u.layout.page, x: stop.x, y: u.layout.lines[stop.line].y + this.shiftOf(u) };
  }

  /** The formula editor's box changed (bp): TeX makes room for it. */
  formulaSized(w: number, h: number, d: number) {
    const f = this.formula;
    if (!f) return;
    const old = f.size;
    if (old && Math.abs(old.w - w) < 0.2 && Math.abs(old.h - h) < 0.2 && Math.abs(old.d - d) < 0.2) return;
    f.size = { w, h, d };
    void this.typeset(f.unit);
  }

  /** The formula's LaTeX changed in the editor: into the source. */
  formulaEdited(latex: string) {
    const f = this.formula;
    const u = f?.unit;
    if (!f || !u) return;
    f.latex = latex;
    const source = this.host.source();
    if (!source) return;
    const body = f.open + latex + f.close;
    this.editing = true;
    try {
      if (f.range) {
        if (!latex.trim()) {
          // emptied: out of the source until it has something again
          source.edit(u.path, [{ from: f.range[0], to: f.range[1], insert: "" }], f.range[0], "input");
          f.at = f.range[0];
          f.range = null;
        } else {
          source.edit(u.path, [{ from: f.range[0], to: f.range[1], insert: body }], f.range[0] + body.length, "input");
          f.range = [f.range[0], f.range[0] + body.length];
        }
      } else if (latex.trim()) {
        // (a formula right after a letter: a space keeps them apart in the source only when TeX needs one)
        source.edit(u.path, [{ from: f.at, to: f.at, insert: body }], f.at + body.length, "input");
        f.range = [f.at, f.at + body.length];
      }
    } finally {
      this.editing = false;
    }
    this.host.scheduleCompile(u.live === true);
    void this.typeset(u);
  }

  /** The formula editor closed: the formula typeset as TeX sets it, the caret before or after it. */
  closeFormula(how: "before" | "after" | "click" | "remove" | "escape") {
    const f = this.formula;
    if (!f) return;
    this.formula = null;
    const u = f.unit;
    if (how === "remove" && f.range) {
      const source = this.host.source();
      this.editing = true;
      try {
        source?.edit(u.path, [{ from: f.range[0], to: f.range[1], insert: "" }], f.range[0], "input");
      } finally {
        this.editing = false;
      }
      f.at = f.range[0];
      f.range = null;
    }
    const offset = how === "before" ? (f.range?.[0] ?? f.at) : (f.range?.[1] ?? f.at);
    this.active = u;
    if (how !== "click") this.caret = { path: u.path, offset, anchor: offset };
    this.host.scheduleCompile(u.live === true);
    void this.typeset(u);
    this.notify();
  }
}

/** LaTeX that is a whole (its braces balanced), else the plain text. */
function balancedOr(latex: string, plain: string): string {
  let depth = 0;
  for (let i = 0; i < latex.length; i++) {
    const c = latex[i];
    if (c === "\\") i++;
    else if (c === "{") depth++;
    else if (c === "}" && --depth < 0) return plain;
  }
  return depth === 0 ? latex : plain;
}
