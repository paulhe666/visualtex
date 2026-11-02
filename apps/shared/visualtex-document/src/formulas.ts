// Formula islands (docs/document-editor/DESIGN.md §3). A formula node is
// drawn as the engine's static markup; at most one formula is open at a time,
// and the app's live formula editor is mounted into it. The plugin state names
// the open formula (by id); the node views follow it through decorations.
import { Fragment, Slice, type Node } from "prosemirror-model";
import { dropPoint } from "prosemirror-transform";
import {
  NodeSelection,
  Plugin,
  PluginKey,
  Selection,
  TextSelection,
  type EditorState,
  type Transaction,
} from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView, type NodeView, type ViewMutationRecord } from "prosemirror-view";
import { isFormulaNodeName } from "./schema.ts";
import { DropMarker, dropPlace } from "./dropPlace.ts";

/** A selection inside a formula, in kernel offsets (the row editor's MathSelectionSnapshot). */
export interface FormulaSelection {
  ranges: Array<[number, number]>;
  direction: "forward" | "backward" | "none";
}

/** Where the caret goes when a formula opens. */
export type FormulaEntry =
  | { kind: "start" }
  | { kind: "end" }
  | { kind: "point"; x: number; y: number }
  | { kind: "selection"; selection: FormulaSelection };

/** How the caret leaves an open formula. */
export type FormulaExit = "before" | "after" | "up" | "down" | "new-paragraph";

/** What the live editor reports to the document. */
export interface FormulaSessionCallbacks {
  /** The formula's LaTeX changed (one history step; undo restores `before`). */
  edit(afterLatex: string, before: FormulaSelection | null, after: FormulaSelection | null): void;
  /** Leave the formula (`x`: the caret's page x, for up/down). */
  exit(how: FormulaExit, x?: number): void;
  /** Backspace in an empty formula: remove it. */
  remove(): void;
  /**
   * "$" typed in the formula. In an empty inline formula it is the second "$" of "$$": the formula becomes a display
   * formula. Otherwise it is the closing dollar: the formula is left after it (after a display formula, a "$" typed
   * next is the second one of the closing "$$" and does nothing).
   */
  dollar(empty: boolean): void;
  undo(): boolean;
  redo(): boolean;
  /** Keys that were waiting for the formula but come after it was left: the document takes them, in order. */
  handBack(keys: FormulaKey[]): void;
  /**
   * The formula's selection is dragged and the pointer is outside the formula, at a page point: the document takes
   * it there (and marks the place); false: not the document's (no place there for it).
   */
  dragOutAt?(x: number, y: number, start: number, end: number): boolean;
  /**
   * Let go there: the selection, stops [start, end] with LaTeX `latex`, leaves the formula (`copy`: stays) for the
   * document — all of it the formula itself, else as a new formula at that place.
   */
  dragOutDrop?(start: number, end: number, latex: string, copy: boolean): void;
  /** The drag is over. */
  dragOutEnd?(): void;
}

/** A key pressed while a formula was opening (it reached the document before the formula's editor had focus). */
export interface FormulaKey {
  key: string;
  code: string;
  /** the character it types, if any */
  text?: string;
  shift: boolean;
  alt: boolean;
  ctrl: boolean;
  meta: boolean;
}

/** The app's live editor mounted in an open formula. */
export interface FormulaSession {
  /** A key that reached the document while the formula was opening: the editor replays it, in order. */
  typeKey(key: FormulaKey): void;
  /** The document changed the formula (undo/redo); `selection` is where its caret was at that state. */
  update(latex: string, selection: FormulaSelection | null): void;
  focus(): void;
  /** A drag that began in the editor left the formula: the editor stops following the pointer. */
  stopPointerTracking?(): void;
  /** Report an edit the editor has made but not reported yet. */
  flush?(): void;
  destroy(): void;
  /** The part of the formula a selection through the text takes, shown selected (null: none yet). */
  showPart?(range: [number, number] | null): void;
  /** The caret stop a page point means in the formula (null until the formula is drawn). */
  stopAt?(x: number, y: number): number | null;
  /** A stop at the formula's top level: inside a structure, the stop before it ("start") or after it ("end"). */
  topLevelStop?(offset: number, side: "start" | "end"): number;
  /** The formula's last stop (null until its editor is there). */
  lastOffset?(): number | null;
  /** The formula's own selection from stop `anchor` to stop `focus` (a drag inside it). */
  selectStops?(anchor: number, focus: number): void;
  /** Where the formula's own selection is anchored (the caret's stop when it has none). */
  anchorStop?(): number | null;
  /** The LaTeX of a range of the formula. */
  latexOf?(range: [number, number]): string;
  /** The formula's LaTeX without a range (null: unchanged). */
  withoutRange?(range: [number, number]): string | null;
  /** The formula's LaTeX with a range set bold or not (null: unchanged). */
  styled?(range: [number, number], kind: "bold", enable: boolean): string | null;
  /** Whether every glyph of a range is bold (null: it holds none). */
  rangeBold?(range: [number, number]): boolean | null;
  /** The open formula's own selection made bold, or not bold when all of it is (Ctrl+B in it); false: no selection. */
  toggleBold?(): boolean;
}

/** What the app provides: static markup and the live editor. */
export interface FormulaHost {
  renderStatic(latex: string, display: boolean): string;
  open(
    container: HTMLElement,
    options: {
      latex: string;
      display: boolean;
      entry: FormulaEntry;
      callbacks: FormulaSessionCallbacks;
      /** Shown only (a part selected by a selection through the text): neither focused nor placed. */
      passive?: boolean;
    },
  ): FormulaSession;
}

/** A formula's part taken by a selection through the text (a mixed selection): kernel stops at its top level. */
export interface FormulaPart {
  id: string;
  /** null until the formula is drawn and the stop under the pointer is known */
  range: [number, number] | null;
}

/** Transaction meta: the formulas' parts that go with the text selection it sets. */
export const FORMULA_PARTS = "visualtexFormulaParts";

interface FormulaPluginState {
  openId: string | null;
  entry: FormulaEntry;
  /** bumps on every open, so reopening the same formula remounts its editor */
  generation: number;
  /** formulas partly selected by a selection through the text (cleared by any other change) */
  parts: FormulaPart[];
}

type FormulaMeta = { open: string; entry: FormulaEntry; keep?: boolean; resume?: boolean } | { close: true };

export const formulaKey = new PluginKey<FormulaPluginState>("visualtex-document-formulas");

/** A window event: the formula fonts changed, closed formulas are drawn again. */
export const FORMULAS_REFRESH_EVENT = "visualtex-document-formulas-refresh";

export function openFormulaId(state: EditorState): string | null {
  return formulaKey.getState(state)?.openId ?? null;
}

/** The formulas partly selected with the text selection (a mixed selection), if any. */
export function formulaParts(state: EditorState): FormulaPart[] {
  return formulaKey.getState(state)?.parts ?? [];
}

/** Find a formula (or any id-carrying node) by id. */
export function findNodeById(doc: Node, id: string): { node: Node; pos: number } | null {
  let found: { node: Node; pos: number } | null = null;
  doc.descendants((node, pos) => {
    if (found) return false;
    if (node.attrs.id === id) {
      found = { node, pos };
      return false;
    }
    return true;
  });
  return found;
}

/** Open the formula at `pos`. */
export function openFormula(tr: Transaction, pos: number, entry: FormulaEntry): Transaction {
  const node = tr.doc.nodeAt(pos);
  if (!node || !isFormulaNodeName(node.type.name) || !node.attrs.id) return tr;
  tr.setSelection(NodeSelection.create(tr.doc, pos));
  return tr.setMeta(formulaKey, { open: node.attrs.id, entry } satisfies FormulaMeta);
}

export function closeFormula(tr: Transaction): Transaction {
  return tr.setMeta(formulaKey, { close: true } satisfies FormulaMeta);
}

// Where each formula's caret was at each of its states: undo/redo restore it.
const selectionMemory = new Map<string, FormulaSelection>();
const memoryKey = (id: string, latex: string) => id + "\u0000" + latex;

function remember(id: string, latex: string, selection: FormulaSelection | null) {
  if (!selection) return;
  if (selectionMemory.size > 4000) selectionMemory.delete(selectionMemory.keys().next().value!);
  selectionMemory.set(memoryKey(id, latex), selection);
}

// The open formula's view per editor view: keys pressed in the document while it opens go to it.
const openViews = new WeakMap<EditorView, FormulaView>();
// Keys handed back by a formula that was left, waiting for a formula that is opening (its editor not yet there).
const waitingKeys = new WeakMap<EditorView, FormulaKey[]>();

function keyEvent(key: FormulaKey): KeyboardEvent {
  return new KeyboardEvent("keydown", {
    key: key.key,
    code: key.code,
    shiftKey: key.shift,
    altKey: key.alt,
    ctrlKey: key.ctrl,
    metaKey: key.meta,
    bubbles: true,
    cancelable: true,
  });
}

/** Keys the document takes in order: into the open formula, or as the document's own keys and typing. */
function deliverKeys(view: EditorView, keys: FormulaKey[]) {
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (formulaKey.getState(view.state)?.openId) {
      const open = openViews.get(view);
      if (open?.hasSession()) {
        open.typeKey(key);
        continue;
      }
      // A formula is opening: the rest wait for its editor.
      waitingKeys.set(view, [...(waitingKeys.get(view) ?? []), ...keys.slice(i)]);
      return;
    }
    if (view.someProp("handleKeyDown", (f) => f(view, keyEvent(key)))) continue;
    if (!key.text) continue;
    const { from, to } = view.state.selection;
    const text = key.text;
    const deflt = () => view.state.tr.insertText(text, from, to);
    if (view.someProp("handleTextInput", (f) => f(view, from, to, text, deflt))) continue;
    view.dispatch(deflt().scrollIntoView());
  }
}

/** Text typed into the formula that was just opened (it goes in once its editor is there). */
export function typeIntoOpenFormula(view: EditorView, text: string) {
  deliverKeys(
    view,
    [...text].map((ch) => ({ key: ch, code: "", text: ch, shift: false, alt: false, ctrl: false, meta: false })),
  );
}

export interface FormulaCommands {
  undo(state: EditorState, dispatch?: (tr: Transaction) => void): boolean;
  redo(state: EditorState, dispatch?: (tr: Transaction) => void): boolean;
  /** An empty display formula at the caret, open ("$$"). */
  displayFormula?(state: EditorState, dispatch?: (tr: Transaction) => void): boolean;
}

/** Meta of the transaction that left a display formula by its closing "$": the next "$" typed is let go. */
export const CLOSED_BY_DOLLAR = "visualtexClosedByDollar";

// The formula views of each editor view by formula id: a drag asks them where the pointer is.
const formulaViews = new WeakMap<EditorView, Map<string, FormulaView>>();

function formulaViewsOf(view: EditorView): Map<string, FormulaView> {
  let map = formulaViews.get(view);
  if (!map) {
    map = new Map();
    formulaViews.set(view, map);
  }
  return map;
}

/** The formula view of a formula (its live editor's questions go through it). */
export function formulaViewOf(view: EditorView, id: string): FormulaView | null {
  return formulaViewsOf(view).get(id) ?? null;
}

/** One end of a selection made with the pointer: a text position, or a caret stop in a formula. */
type SelectEnd = { kind: "text"; pos: number } | { kind: "formula"; id: string; stop: number | null };

/**
 * Selecting with the pointer, the formula rows' rules (docs/editing-core/DESIGN.md §6.9; DESIGN.md §3.5): a press in
 * the text or in a formula anchors there; while the pointer stays in the formula it was pressed in, the formula
 * selects (its own, structural selection); beyond it the selection runs through the text, and an end inside a formula
 * — the one it began in, or the one the pointer is over, inline or display — takes that formula's part up to it, at
 * the formula's top level (an end inside a structure takes the whole structure), shown selected in the formula.
 * Only moves with the left button held count (AxMath); past the top or bottom of the document it scrolls.
 */
export class SelectGesture {
  private last: { x: number; y: number } | null = null;
  private timer = 0;
  private refresh = 0;
  /** the pointer left the formula the drag began in (its editor no longer follows it) */
  private left = false;
  private anchorStopValue: number | null;
  private readonly scroller: HTMLElement | null;

  constructor(
    private readonly view: EditorView,
    private readonly press: { x: number; y: number },
    private readonly anchor: SelectEnd,
    /** the formula's editor follows the pointer itself until it leaves the formula (a press in an open formula) */
    private readonly editorFollows = false,
  ) {
    this.anchorStopValue = anchor.kind === "formula" ? anchor.stop : null;
    this.scroller = scrollParent(view.dom);
    window.addEventListener("pointermove", this.onMove, true);
    window.addEventListener("pointerup", this.onUp, true);
    window.addEventListener("pointercancel", this.finish, true);
  }

  private finish = () => {
    window.removeEventListener("pointermove", this.onMove, true);
    window.removeEventListener("pointerup", this.onUp, true);
    window.removeEventListener("pointercancel", this.finish, true);
    window.clearInterval(this.timer);
    cancelAnimationFrame(this.refresh);
    this.timer = 0;
  };

  private onMove = (event: PointerEvent) => {
    if (!(event.buttons & 1)) {
      // the release went past us (outside the window): the gesture is over
      this.finish();
      return;
    }
    this.last = { x: event.clientX, y: event.clientY };
    if (this.update()) {
      event.preventDefault();
      event.stopImmediatePropagation();
    }
    if (!this.timer && this.scroller) {
      // Past the top or bottom of the visible document: it scrolls, the selection follows.
      this.timer = window.setInterval(() => {
        if (!this.last || !this.scroller) return;
        const frame = this.scroller.getBoundingClientRect();
        const step = this.last.y < frame.top + 4 ? -16 : this.last.y > frame.bottom - 4 ? 16 : 0;
        if (!step) return;
        this.scroller.scrollTop += step;
        this.update();
      }, 30);
    }
  };

  private onUp = (event: PointerEvent) => {
    if (event.buttons & 1) return;
    this.last = { x: event.clientX, y: event.clientY };
    if (this.update()) {
      event.preventDefault();
      event.stopImmediatePropagation();
    }
    this.finish();
  };

  private posOf(id: string): number | null {
    return formulaViewOf(this.view, id)?.position() ?? null;
  }

  /** Where the pointer is: over a formula (its stop), else the text position nearest it. */
  private endAt(x: number, y: number): SelectEnd | null {
    const view = this.view;
    const element = document.elementFromPoint(x, y)?.closest?.(".vtx-math-inline, .vtx-math-display");
    if (element && view.dom.contains(element)) {
      for (const formula of formulaViewsOf(view).values()) {
        if (formula.dom === element && formula.position() !== null) return { kind: "formula", id: formula.id, stop: formula.stopAt(x, y) };
      }
    }
    const editor = view.dom.getBoundingClientRect();
    const frame = this.scroller?.getBoundingClientRect() ?? editor;
    const left = Math.min(Math.max(x, editor.left + 1), editor.right - 1);
    const top = Math.min(Math.max(y, Math.max(editor.top, frame.top) + 1), Math.min(editor.bottom, frame.bottom) - 1);
    const hit = view.posAtCoords({ left, top });
    if (!hit) return null;
    const inside = hit.inside >= 0 ? view.state.doc.nodeAt(hit.inside) : null;
    if (inside && isFormulaNodeName(inside.type.name) && inside.attrs.id) {
      const formula = formulaViewOf(view, inside.attrs.id);
      return { kind: "formula", id: inside.attrs.id, stop: formula ? formula.stopAt(x, y) : null };
    }
    return { kind: "text", pos: hit.pos };
  }

  /**
   * The stop the drag began at in its formula: where its editor's selection was anchored when the editor followed
   * the pointer, else the press point, asked again each time (the editor started with the press and is drawn a
   * frame or two later).
   */
  private anchorStop(): number | null {
    if (this.anchor.kind !== "formula") return null;
    if (this.anchorStopValue !== null) return this.anchorStopValue;
    return formulaViewOf(this.view, this.anchor.id)?.stopAt(this.press.x, this.press.y) ?? null;
  }

  /** The selection for the pointer's place; true when the gesture (not the formula's editor) made it. */
  private update(): boolean {
    if (!this.last) return false;
    if (this.editorFollows && this.anchor.kind === "formula" && formulaViewOf(this.view, this.anchor.id)?.dom.querySelector(".visualtex-selection-drag")) {
      // the press was on the formula's selection: the editor drags it (out of the formula too); no selection here
      this.finish();
      return false;
    }
    const end = this.endAt(this.last.x, this.last.y);
    if (!end) return false;
    const anchor = this.anchor;
    if (anchor.kind === "formula" && end.kind === "formula" && end.id === anchor.id) {
      // in the formula the drag began in: its own selection (the editor's, until the pointer has left it once)
      if (this.editorFollows && !this.left) return false;
      this.selectInAnchor(end.stop);
      return true;
    }
    if (anchor.kind === "formula" && !this.left) {
      this.left = true;
      const formula = formulaViewOf(this.view, anchor.id);
      if (this.editorFollows) {
        // the editor followed the pointer so far: the drag began at its selection's anchor
        this.anchorStopValue = formula?.anchorStop() ?? this.anchorStop();
        formula?.stopPointerTracking();
      }
    }
    this.selectAcross(end);
    return true;
  }

  private selectInAnchor(stop: number | null) {
    const view = this.view;
    const anchor = this.anchor as { kind: "formula"; id: string };
    const pos = this.posOf(anchor.id);
    if (pos === null) return;
    if (openFormulaId(view.state) !== anchor.id) {
      const tr = view.state.tr.setSelection(NodeSelection.create(view.state.doc, pos));
      view.dispatch(tr.setMeta(formulaKey, { open: anchor.id, entry: { kind: "end" }, resume: true } satisfies FormulaMeta).setMeta(FORMULA_PARTS, []));
    }
    const from = this.anchorStop();
    const formula = formulaViewOf(view, anchor.id);
    if (from === null || stop === null || !formula) {
      this.refreshSoon();
      return;
    }
    formula.selectStops(from, stop);
  }

  /** A selection through the text from the anchor to `end`, each end in a formula taking that formula's part. */
  private selectAcross(end: SelectEnd) {
    const view = this.view;
    const doc = view.state.doc;
    const anchor = this.anchor;
    // positions in document order: a formula sits between the positions before and after it
    const key = (e: SelectEnd) => (e.kind === "text" ? e.pos : (this.posOf(e.id) ?? 0) + 0.5);
    const forward = key(anchor) <= key(end);
    const first = forward ? anchor : end;
    const second = forward ? end : anchor;
    const parts: FormulaPart[] = [];
    let pending = false;
    const stopOf = (e: SelectEnd & { kind: "formula" }) => (e === anchor ? this.anchorStop() : e.stop);
    let from: number;
    let to: number;
    let firstBlock = false;
    let secondBlock = false;
    if (first.kind === "formula") {
      const pos = this.posOf(first.id);
      const node = pos === null ? null : doc.nodeAt(pos);
      const formula = formulaViewOf(view, first.id);
      if (pos === null || !node || !formula) return;
      from = pos + node.nodeSize;
      firstBlock = node.isBlock;
      const stop = stopOf(first);
      const last = formula.lastOffset();
      const range: [number, number] | null = stop === null || last === null ? null : [formula.topLevelStop(stop, "start"), last];
      if (!range) pending = true;
      parts.push({ id: first.id, range });
    } else from = first.pos;
    if (second.kind === "formula") {
      const pos = this.posOf(second.id);
      const node = pos === null ? null : doc.nodeAt(pos);
      const formula = formulaViewOf(view, second.id);
      to = pos ?? from;
      secondBlock = !!node?.isBlock;
      const stop = stopOf(second);
      const range: [number, number] | null = !formula || stop === null ? null : [0, formula.topLevelStop(stop, "end")];
      if (!range) pending = true;
      parts.push({ id: second.id, range });
    } else to = second.pos;
    // beside a display formula the text selection ends at the nearest text, if there is any between the ends
    if (firstBlock) {
      const near = Selection.findFrom(doc.resolve(from), 1, true);
      from = near && near.from <= to ? near.from : to;
    }
    if (secondBlock) {
      const near = Selection.findFrom(doc.resolve(to), -1, true);
      to = near && near.to >= from ? near.to : from;
    }
    if (from > to) from = to;
    const selection = forward
      ? TextSelection.between(doc.resolve(from), doc.resolve(to))
      : TextSelection.between(doc.resolve(to), doc.resolve(from));
    const before = formulaParts(view.state);
    const same =
      selection.eq(view.state.selection) &&
      before.length === parts.length &&
      before.every((p, i) => p.id === parts[i].id && JSON.stringify(p.range) === JSON.stringify(parts[i].range));
    if (!same) {
      const tr = view.state.tr.setSelection(selection).setMeta(FORMULA_PARTS, parts);
      if (openFormulaId(view.state)) closeFormula(tr);
      view.dispatch(tr);
      view.focus();
    }
    // a formula whose editor is still starting: its part once it is drawn
    if (pending) this.refreshSoon();
  }

  private refreshSoon() {
    cancelAnimationFrame(this.refresh);
    this.refresh = requestAnimationFrame(() => {
      this.refresh = requestAnimationFrame(() => {
        this.refresh = 0;
        if (this.last) this.update();
      });
    });
  }
}

/** The element that scrolls the document (for a drag past its edge). */
function scrollParent(element: HTMLElement): HTMLElement | null {
  for (let node = element.parentElement; node; node = node.parentElement) {
    const overflow = getComputedStyle(node).overflowY;
    if ((overflow === "auto" || overflow === "scroll") && node.scrollHeight > node.clientHeight) return node;
  }
  return null;
}

/**
 * An open formula's selection dragged out of the formula (DESIGN.md §3.5): beside or between the document's text it
 * lands as a formula — the formula itself when the selection is all of it, else a new formula of its part (inline
 * in a line, a display formula between blocks); the formula keeps the rest (a formula left empty goes).
 */
class FormulaDragOut {
  private readonly marker = new DropMarker();
  private target: number | null = null;

  constructor(
    private readonly view: EditorView,
    private readonly id: string,
  ) {}

  /** Whether the document takes the selection (stops [start, end]) at a page point: outside the formula, in it. */
  at(x: number, y: number, start: number, end: number): boolean {
    const view = this.view;
    const at = findNodeById(view.state.doc, this.id);
    const dom = formulaViewOf(view, this.id)?.dom;
    const editor = view.dom.getBoundingClientRect();
    const inRect = (r: DOMRect) => x >= r.left && x <= r.right && y >= r.top && y <= r.bottom;
    if (!at || !dom || inRect(dom.getBoundingClientRect()) || !inRect(editor)) {
      this.target = null;
      this.marker.remove();
      return false;
    }
    const whole = start === 0 && end === formulaViewOf(view, this.id)?.lastOffset();
    const moved = whole ? at.node : view.state.schema.nodes.math_inline.create({ latex: "x" });
    const slice = new Slice(Fragment.from(moved), 0, 0);
    this.target = dropPlace(view, x, y, slice, at.pos, at.pos + at.node.nodeSize);
    this.marker.show(view, this.target);
    return true;
  }

  drop(start: number, end: number, latex: string, copy: boolean) {
    const view = this.view;
    const target = this.target;
    this.end();
    const at = findNodeById(view.state.doc, this.id);
    const formula = formulaViewOf(view, this.id);
    if (target === null || !at || !formula || !latex.trim()) return;
    const whole = start === 0 && end === formula.lastOffset();
    const tr = closeFormula(view.state.tr);
    let node: Node;
    if (whole) {
      node = copy ? at.node.type.create({ ...at.node.attrs, id: null }) : at.node;
      if (!copy) tr.delete(at.pos, at.pos + at.node.nodeSize);
    } else {
      const rest = copy ? null : formula.withoutRange([start, end]);
      if (rest !== null && !rest.trim()) tr.delete(at.pos, at.pos + at.node.nodeSize);
      else if (rest !== null) tr.setNodeAttribute(at.pos, "latex", rest);
      node = view.state.schema.nodes.math_inline.create({ latex });
    }
    const pos = tr.mapping.map(target);
    const $pos = tr.doc.resolve(pos);
    // between blocks a formula part is a display formula; in a line a whole display formula still goes between blocks
    if (!$pos.parent.inlineContent && node.isInline) node = view.state.schema.nodes.math_display.create({ latex });
    const placeAt = node.isBlock && $pos.parent.inlineContent ? (dropPoint(tr.doc, pos, new Slice(Fragment.from(node), 0, 0)) ?? pos) : pos;
    tr.insert(placeAt, node);
    tr.setSelection(NodeSelection.create(tr.doc, placeAt));
    view.dispatch(tr.setMeta("uiEvent", "drop").scrollIntoView());
    view.focus();
  }

  end() {
    this.marker.remove();
  }
}

export class FormulaView implements NodeView {
  dom: HTMLElement;
  private staticBox: HTMLElement;
  private numberBox: HTMLElement | null = null;
  private editorBox: HTMLElement | null = null;
  private session: FormulaSession | null = null;
  private generation = -1;
  private renderedLatex: string | null = null;
  private pendingOpen: { entry: FormulaEntry; generation: number } | null = null;
  private destroyed = false;

  constructor(
    private node: Node,
    private view: EditorView,
    private getPos: () => number | undefined,
    private host: FormulaHost,
    private commands: FormulaCommands,
    decorations: readonly Decoration[],
  ) {
    const display = node.type.name === "math_display";
    this.dom = document.createElement(display ? "div" : "span");
    this.dom.className = display ? "vtx-math-display" : "vtx-math-inline";
    this.dom.contentEditable = "false";
    this.staticBox = document.createElement(display ? "div" : "span");
    this.staticBox.className = "vtx-formula-static";
    this.dom.append(this.staticBox);
    if (display) {
      // Mathcha's tag: "(n)" when numbered; while the formula is edited a grey "(...)" offers a number.
      this.numberBox = document.createElement("span");
      this.numberBox.className = "vtx-formula-number";
      this.numberBox.addEventListener("mousedown", this.onTagMouseDown);
      this.dom.append(this.numberBox);
    }
    if (!display) {
      // The rendered formula holds no text of its own (its glyphs sit in absolutely placed boxes), and Chrome then
      // loses an IME composition started right after or before it: the first letter stays, the rest is dropped and
      // the selection jumps onto the formula. A zero-width character at each end gives the caret real text to sit by.
      this.dom.prepend(document.createTextNode("\u200B"));
      this.dom.append(document.createTextNode("\u200B"));
    }
    // A click opens the formula with the caret at the clicked place.
    this.dom.addEventListener("mousedown", this.onMouseDown);
    window.addEventListener(FORMULAS_REFRESH_EVENT, this.onRefresh);
    this.registeredId = node.attrs.id ?? null;
    if (this.registeredId) formulaViewsOf(view).set(this.registeredId, this);
    this.renderStatic();
    this.applyDecorations(decorations);
  }

  /** the id this view is registered under (formulaViewOf) */
  private registeredId: string | null = null;
  /** the editor only shows a part of the formula selected by a selection through the text */
  private passive = false;
  private dragOut: FormulaDragOut | null = null;

  get id(): string {
    return this.node.attrs.id as string;
  }

  /** The formula's position in the document (null: not in it any more). */
  position(): number | null {
    return this.destroyed ? null : this.getPos() ?? null;
  }

  stopAt(x: number, y: number): number | null {
    return this.session?.stopAt?.(x, y) ?? null;
  }

  topLevelStop(offset: number, side: "start" | "end"): number {
    return this.session?.topLevelStop?.(offset, side) ?? offset;
  }

  lastOffset(): number | null {
    return this.session?.lastOffset?.() ?? null;
  }

  anchorStop(): number | null {
    return this.session?.anchorStop?.() ?? null;
  }

  selectStops(anchor: number, focus: number) {
    this.session?.selectStops?.(anchor, focus);
  }

  stopPointerTracking() {
    this.session?.stopPointerTracking?.();
  }

  latexOf(range: [number, number]): string | null {
    return this.session?.latexOf?.(range) ?? null;
  }

  withoutRange(range: [number, number]): string | null {
    return this.session?.withoutRange?.(range) ?? null;
  }

  styled(range: [number, number], enable: boolean): string | null {
    return this.session?.styled?.(range, "bold", enable) ?? null;
  }

  rangeBold(range: [number, number]): boolean | null {
    return this.session?.rangeBold?.(range) ?? null;
  }

  toggleBold(): boolean {
    return !this.passive && !!this.session?.toggleBold?.();
  }

  private get display() {
    return this.node.type.name === "math_display";
  }

  // The tag of the formula being edited numbers it or removes the number; on a closed formula a click opens it.
  private onTagMouseDown = (event: MouseEvent) => {
    if (event.button !== 0 || !this.session) return;
    event.preventDefault();
    event.stopPropagation();
    const pos = this.getPos();
    if (pos === undefined) return;
    const tr = this.view.state.tr.setNodeAttribute(pos, "numbered", !this.node.attrs.numbered);
    tr.setMeta(formulaKey, { open: this.node.attrs.id, entry: { kind: "end" }, keep: true });
    this.view.dispatch(tr);
    this.session?.focus();
  };

  private onMouseDown = (event: MouseEvent) => {
    if ((this.session && !this.passive) || event.button !== 0) return;
    if (event.shiftKey) return; // extend the text selection over the formula
    const pos = this.getPos();
    if (pos === undefined) return;
    // Part of the selection: pressing on it drags the selection (the browser's drag, as on its text); a click
    // without a drag opens the formula.
    const sel = this.view.state.selection;
    if (!sel.empty && sel.from <= pos && pos + this.node.nodeSize <= sel.to) {
      let dragged = false;
      const onDragStart = () => {
        dragged = true;
      };
      this.view.dom.addEventListener("dragstart", onDragStart, { once: true });
      window.addEventListener(
        "mouseup",
        (up) => {
          this.view.dom.removeEventListener("dragstart", onDragStart);
          if (dragged || this.destroyed) return;
          const at = this.getPos();
          if (at === undefined) return;
          this.view.dispatch(openFormula(this.view.state.tr, at, { kind: "point", x: up.clientX, y: up.clientY }));
          this.view.focus();
        },
        { once: true, capture: true },
      );
      return;
    }
    event.preventDefault();
    // The press opens the formula at that place and a drag from it selects: in the formula from the first press
    // (its editor is only starting, so the drag is followed here), beyond it through the text.
    this.view.dispatch(openFormula(this.view.state.tr, pos, { kind: "point", x: event.clientX, y: event.clientY }));
    this.view.focus();
    new SelectGesture(this.view, { x: event.clientX, y: event.clientY }, { kind: "formula", id: this.id, stop: null });
  };

  // A drag that begins in the open formula's editor: the editor selects in it; beyond it the selection runs on.
  private onEditorPointerDown = (event: PointerEvent) => {
    if (event.button !== 0 || event.shiftKey || this.passive) return;
    new SelectGesture(this.view, { x: event.clientX, y: event.clientY }, { kind: "formula", id: this.id, stop: null }, true);
  };

  private onRefresh = () => {
    this.renderedLatex = null;
    if (!this.session) this.renderStatic();
  };

  private renderStatic() {
    const latex = this.node.attrs.latex as string;
    if (this.renderedLatex === latex) return;
    this.renderedLatex = latex;
    const markup = latex.trim() ? this.host.renderStatic(latex, this.display) : "";
    this.staticBox.innerHTML = markup;
    this.dom.classList.toggle("is-empty", !latex.trim() || !markup);
    if (!markup && latex.trim()) this.staticBox.textContent = latex;
  }

  private applyDecorations(decorations: readonly Decoration[]) {
    let open: { entry: FormulaEntry; generation: number } | null = null;
    let number: string | null = null;
    let part: { range: [number, number] | null } | null = null;
    for (const d of decorations) {
      const spec = d.spec as {
        formulaOpen?: { entry: FormulaEntry; generation: number };
        equationNumber?: string;
        formulaPart?: { range: [number, number] | null };
      };
      if (spec.formulaOpen) open = spec.formulaOpen;
      if (spec.equationNumber !== undefined) number = spec.equationNumber;
      if (spec.formulaPart) part = spec.formulaPart;
    }
    this.equationNumber = number;
    this.renderTag(!!open);
    this.dom.classList.toggle("is-numbered", number !== null);
    if (open && this.session && this.passive && open.generation === this.generation) {
      // back into it during a drag: the editor that showed its part edits it again
      this.activate();
      return;
    }
    if (!open && part) {
      // a selection through the text takes a part of it: the editor shows that part (started for it if need be)
      this.pendingOpen = null;
      if (!this.session) this.startSession({ kind: "end" }, -1, true);
      else this.passivate();
      this.session?.showPart?.(part.range);
      return;
    }
    if (open && (!this.session || open.generation !== this.generation)) {
      // A formula created by this very transaction is not in the page yet: its editor starts once it is.
      if (this.dom.isConnected) this.startSession(open.entry, open.generation);
      else {
        const pending = open;
        this.pendingOpen = pending;
        queueMicrotask(() => {
          if (this.pendingOpen !== pending) return;
          this.pendingOpen = null;
          if (this.dom.isConnected && !this.destroyed) this.startSession(pending.entry, pending.generation);
        });
      }
    } else if (!open) {
      this.pendingOpen = null;
      if (this.session) this.endSession();
    }
  }

  private equationNumber: string | null = null;

  private renderTag(open: boolean) {
    if (!this.numberBox) return;
    const number = this.equationNumber;
    this.numberBox.textContent = number !== null ? number : open ? "(...)" : "";
    this.numberBox.classList.toggle("is-placeholder", number === null);
    this.numberBox.title = number !== null ? "Remove the number" : "Number this formula";
  }

  private startSession(entry: FormulaEntry, generation: number, passive = false) {
    this.endSession();
    this.generation = generation;
    this.passive = passive;
    this.editorBox = document.createElement(this.display ? "div" : "span");
    this.editorBox.className = "vtx-formula-editor";
    this.editorBox.addEventListener("pointerdown", this.onEditorPointerDown, true);
    // showing a part, the editor takes no press (a press on it is the document's: it drags or opens the formula)
    this.editorBox.inert = passive;
    this.dom.insertBefore(this.editorBox, this.staticBox);
    this.dom.classList.add(passive ? "is-partial" : "is-open");
    this.renderTag(!passive);
    if (!passive) openViews.set(this.view, this);
    const id = this.node.attrs.id as string;
    const view = this.view;
    const getPos = this.getPos;
    const commands = this.commands;
    const current = () => {
      const pos = getPos();
      const node = pos === undefined ? null : view.state.doc.nodeAt(pos);
      return pos === undefined || !node ? null : { pos, node };
    };
    this.session = this.host.open(this.editorBox, {
      latex: this.node.attrs.latex,
      display: this.display,
      entry,
      passive,
      callbacks: {
        edit: (afterLatex, before, after) => {
          const at = current();
          if (!at || at.node.attrs.latex === afterLatex) return;
          remember(id, at.node.attrs.latex, before);
          remember(id, afterLatex, after);
          // (A late edit after the formula closed only stores the LaTeX; it does not reopen it.)
          view.dispatch(view.state.tr.setNodeAttribute(at.pos, "latex", afterLatex));
        },
        exit: (how, x) => {
          const at = current();
          if (!at) return;
          view.dispatch(leaveFormula(view, at.pos, at.node, how, x));
          view.focus();
        },
        remove: () => {
          const at = current();
          if (!at) return;
          const tr = view.state.tr.delete(at.pos, at.pos + at.node.nodeSize);
          closeFormula(tr);
          view.dispatch(tr.scrollIntoView());
          view.focus();
        },
        dollar: (empty) => {
          const at = current();
          if (!at) return;
          if (!empty) {
            const tr = leaveFormula(view, at.pos, at.node, "after");
            if (at.node.isBlock) tr.setMeta(CLOSED_BY_DOLLAR, true);
            view.dispatch(tr);
            view.focus();
            return;
          }
          if (at.node.isBlock || !commands.displayFormula) return;
          // "$$": the inline formula gives way to a display formula, opened in its place.
          const tr = closeFormula(view.state.tr.delete(at.pos, at.pos + at.node.nodeSize));
          view.dispatch(tr.setSelection(TextSelection.create(tr.doc, at.pos)));
          commands.displayFormula(view.state, view.dispatch);
        },
        undo: () => commands.undo(view.state, view.dispatch),
        redo: () => commands.redo(view.state, view.dispatch),
        handBack: (keys) => deliverKeys(view, keys),
        dragOutAt: (x, y, start, end) => {
          this.dragOut ??= new FormulaDragOut(view, id);
          return this.dragOut.at(x, y, start, end);
        },
        dragOutDrop: (start, end, latex, copy) => this.dragOut?.drop(start, end, latex, copy),
        // (the kernel ends the drag before it drops: the place is kept for the drop)
        dragOutEnd: () => this.dragOut?.end(),
      },
    });
    // Keys handed back by a formula left just before this one opened.
    const waiting = waitingKeys.get(view);
    if (waiting?.length) {
      waitingKeys.delete(view);
      for (const key of waiting) this.session.typeKey(key);
    }
  }

  hasSession(): boolean {
    return !!this.session;
  }

  /** The editor stops editing and only shows the part a selection through the text takes. */
  private passivate() {
    if (!this.session || this.passive) return;
    this.passive = true;
    if (openViews.get(this.view) === this) openViews.delete(this.view);
    this.dom.classList.remove("is-open");
    this.dom.classList.add("is-partial");
    if (this.editorBox) this.editorBox.inert = true;
    this.renderTag(false);
  }

  /** The editor that showed a part edits the formula again (a drag came back into it). */
  private activate() {
    if (!this.session) return;
    this.passive = false;
    openViews.set(this.view, this);
    this.dom.classList.remove("is-partial");
    this.dom.classList.add("is-open");
    if (this.editorBox) this.editorBox.inert = false;
    this.renderTag(true);
    this.session.showPart?.(null);
    this.session.focus();
  }

  typeKey(key: FormulaKey) {
    this.session?.typeKey(key);
  }

  /** A key that reached the document while this formula is open but its editor has no focus yet. */
  forwardKey(event: KeyboardEvent): boolean {
    if (!this.session || !this.editorBox) return false;
    const active = document.activeElement;
    if (active && this.editorBox.contains(active)) return false;
    if (["Shift", "Control", "Alt", "Meta", "CapsLock"].includes(event.key) || event.isComposing) return false;
    const printable = event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey;
    this.session.typeKey({
      key: event.key,
      code: event.code,
      text: printable ? event.key : undefined,
      shift: event.shiftKey,
      alt: event.altKey,
      ctrl: event.ctrlKey,
      meta: event.metaKey,
    });
    return true;
  }

  private endSession() {
    if (!this.session) return;
    const session = this.session;
    this.session = null;
    this.passive = false;
    if (openViews.get(this.view) === this) openViews.delete(this.view);
    session.destroy();
    this.editorBox?.remove();
    this.editorBox = null;
    this.dom.classList.remove("is-open", "is-partial");
    this.renderTag(false);
    this.renderStatic();
  }

  update(node: Node, decorations: readonly Decoration[]): boolean {
    if (node.type !== this.node.type) return false;
    const latexChanged = node.attrs.latex !== this.node.attrs.latex;
    this.node = node;
    if (this.registeredId !== (node.attrs.id ?? null)) {
      const views = formulaViewsOf(this.view);
      if (this.registeredId && views.get(this.registeredId) === this) views.delete(this.registeredId);
      this.registeredId = node.attrs.id ?? null;
      if (this.registeredId) views.set(this.registeredId, this);
    }
    if (this.session && latexChanged) {
      // The document changed the open formula (undo/redo): the editor follows with the caret it had then.
      this.session.update(node.attrs.latex, selectionMemory.get(memoryKey(node.attrs.id, node.attrs.latex)) ?? null);
    }
    if (!this.session) this.renderStatic();
    this.applyDecorations(decorations);
    return true;
  }

  selectNode() {
    this.dom.classList.add("is-selected");
  }

  deselectNode() {
    this.dom.classList.remove("is-selected");
  }

  // Events inside the live editor belong to it; a click on a closed formula (or one showing a part) is ours.
  stopEvent(event: Event) {
    if (this.session && !this.passive) return true;
    return event.type === "mousedown";
  }

  ignoreMutation(_mutation: ViewMutationRecord) {
    return true;
  }

  destroy() {
    this.destroyed = true;
    const views = formulaViewsOf(this.view);
    if (this.registeredId && views.get(this.registeredId) === this) views.delete(this.registeredId);
    this.dom.removeEventListener("mousedown", this.onMouseDown);
    this.numberBox?.removeEventListener("mousedown", this.onTagMouseDown);
    window.removeEventListener(FORMULAS_REFRESH_EVENT, this.onRefresh);
    this.endSession();
  }
}

/** The transaction that closes the formula at `pos` and puts the text caret where `how` says. */
function leaveFormula(view: EditorView, pos: number, node: Node, how: FormulaExit, x?: number): Transaction {
  const tr = closeFormula(view.state.tr);
  const after = pos + node.nodeSize;
  const doc = tr.doc;
  const near = (p: number, bias: -1 | 1) => Selection.near(doc.resolve(p), bias);
  // Below a display formula there is always somewhere to go (Mathcha): the block that follows it — its first line,
  // or the display formula after it, opened at its start — and only when nothing follows is a line made.
  const below = (): Transaction => {
    const next = tr.doc.resolve(after).nodeAfter;
    if (next && isFormulaNodeName(next.type.name)) return openFormula(tr, after, { kind: "start" }).scrollIntoView();
    if (next?.isTextblock) return tr.setSelection(TextSelection.create(tr.doc, after + 1)).scrollIntoView();
    const inside = next ? Selection.findFrom(tr.doc.resolve(after), 1, true) : null;
    if (inside && inside.from <= after + next!.nodeSize) return tr.setSelection(inside).scrollIntoView();
    tr.insert(after, view.state.schema.nodes.paragraph.create());
    return tr.setSelection(TextSelection.create(tr.doc, after + 1)).scrollIntoView();
  };
  if (how === "before") return tr.setSelection(near(pos, -1)).scrollIntoView();
  if (how === "after") return node.isBlock ? below() : tr.setSelection(near(after, 1)).scrollIntoView();
  if (how === "new-paragraph") {
    const paragraph = view.state.schema.nodes.paragraph.create();
    tr.insert(after, paragraph);
    return tr.setSelection(TextSelection.create(tr.doc, after + 1)).scrollIntoView();
  }
  // Up/Down: the text line above/below the formula, at the caret's x.
  const dom = view.nodeDOM(pos) as HTMLElement | null;
  const rect = dom?.getBoundingClientRect();
  if (rect) {
    const left = x ?? (how === "up" ? rect.left : rect.right);
    // The nearest point beyond the formula's own line (an inline formula's line can be taller than the formula).
    for (const step of [4, 8, 14, 22, 34]) {
      const top = how === "up" ? rect.top - step : rect.bottom + step;
      const hit = view.posAtCoords({ left, top });
      if (!hit || (hit.pos >= pos && hit.pos <= after)) continue;
      const coords = view.coordsAtPos(hit.pos);
      const beyond = how === "up" ? coords.bottom <= rect.top + 2 : coords.top >= rect.bottom - 2;
      if (beyond || node.isBlock) return tr.setSelection(near(hit.pos, how === "up" ? -1 : 1)).scrollIntoView();
    }
  }
  if (how === "up") {
    // A display formula first in the document: a line is made above it to go to.
    if (node.isBlock && !Selection.findFrom(doc.resolve(pos), -1, true)) {
      tr.insert(pos, view.state.schema.nodes.paragraph.create());
      return tr.setSelection(TextSelection.create(tr.doc, pos + 1)).scrollIntoView();
    }
    return tr.setSelection(near(pos, -1)).scrollIntoView();
  }
  return node.isBlock ? below() : tr.setSelection(near(after, 1)).scrollIntoView();
}

/** The formula plugin: open/close state, node views, the open formula's decoration. */
export function formulaPlugin(host: FormulaHost, commands: FormulaCommands): Plugin<FormulaPluginState> {
  return new Plugin<FormulaPluginState>({
    key: formulaKey,
    state: {
      init: () => ({ openId: null, entry: { kind: "end" }, generation: 0, parts: [] }),
      apply(tr, value) {
        // The parts go with the selection that set them, while neither the document nor the selection changes.
        const setParts = tr.getMeta(FORMULA_PARTS) as FormulaPart[] | undefined;
        const parts = setParts ?? (tr.docChanged || tr.selectionSet ? [] : value.parts);
        if (parts !== value.parts) value = { ...value, parts };
        const meta = tr.getMeta(formulaKey) as FormulaMeta | undefined;
        if (meta && "close" in meta) return { ...value, openId: null };
        if (meta && "open" in meta) {
          if (meta.keep && value.openId === meta.open) return value;
          // back into a formula still on the page (a drag returning to it): its editor stays
          if (meta.resume) return { ...value, openId: meta.open, entry: meta.entry };
          return { ...value, openId: meta.open, entry: meta.entry, generation: value.generation + 1 };
        }
        if (!value.openId) return value;
        // The open formula is gone (deleted, undone away): nothing is open.
        if (tr.docChanged && !findNodeById(tr.doc, value.openId)) return { ...value, openId: null };
        // The text caret moved elsewhere (a click, a key in the text): the formula closes.
        if (tr.selectionSet && !tr.getMeta("history$")) {
          const sel = tr.selection;
          const stillOn =
            sel instanceof NodeSelection && sel.node.attrs.id === value.openId;
          if (!stillOn) return { ...value, openId: null };
        }
        return value;
      },
    },
    // An inline formula left empty goes away, however it was left (Backspace in it does the same). Outside the
    // history: undo does not bring an empty formula back.
    appendTransaction(_transactions, oldState, newState) {
      const closed = formulaKey.getState(oldState)?.openId;
      if (!closed || formulaKey.getState(newState)?.openId === closed) return null;
      const found = findNodeById(newState.doc, closed);
      if (!found || found.node.type.name !== "math_inline" || String(found.node.attrs.latex).trim()) return null;
      return newState.tr.delete(found.pos, found.pos + found.node.nodeSize).setMeta("addToHistory", false);
    },
    props: {
      // With formulas partly selected the text caret is not shown (the selection is a range, drawn in them).
      attributes(state): Record<string, string> {
        return formulaParts(state).length ? { class: "vtx-has-formula-parts" } : {};
      },
      handleDOMEvents: {
        // A press in the text anchors a selection the pointer then makes (SelectGesture): through the text, into
        // formulas, part of them. Double and triple clicks, Shift, Ctrl, tables and other atoms keep the editor's own.
        mousedown(view, event) {
          if (event.button !== 0 || event.shiftKey || event.ctrlKey || event.metaKey || event.altKey || event.detail > 1) return false;
          const target = event.target as Element | null;
          if (!target?.closest || target.closest(".vtx-math-inline, .vtx-math-display, td, th, .vtx-code-header, a.vtx-link")) return false;
          const hit = view.posAtCoords({ left: event.clientX, top: event.clientY });
          if (!hit) return false;
          const inside = hit.inside >= 0 ? view.state.doc.nodeAt(hit.inside) : null;
          if (inside?.isAtom) return false;
          event.preventDefault();
          const $pos = view.state.doc.resolve(hit.pos);
          const caret = $pos.parent.inlineContent ? TextSelection.create(view.state.doc, hit.pos) : Selection.near($pos);
          view.focus();
          const tr = view.state.tr.setSelection(caret);
          if (openFormulaId(view.state)) closeFormula(tr);
          view.dispatch(tr.scrollIntoView());
          new SelectGesture(view, { x: event.clientX, y: event.clientY }, { kind: "text", pos: view.state.selection.head });
          return true;
        },
      },
      // Keys pressed right after a formula opened, before its editor took the focus, belong to the formula.
      handleKeyDown(view, event) {
        if (!formulaKey.getState(view.state)?.openId) return false;
        if (!openViews.get(view)?.forwardKey(event)) return false;
        event.preventDefault();
        return true;
      },
      decorations(state) {
        const decorations: Decoration[] = [];
        const value = formulaKey.getState(state);
        const found = value?.openId ? findNodeById(state.doc, value.openId) : null;
        if (value && found) {
          decorations.push(
            Decoration.node(found.pos, found.pos + found.node.nodeSize, {}, {
              formulaOpen: { entry: value.entry, generation: value.generation },
            }),
          );
        }
        // The formulas a selection through the text takes a part of: their editors show it.
        for (const part of value?.parts ?? []) {
          const at = findNodeById(state.doc, part.id);
          if (at) decorations.push(Decoration.node(at.pos, at.pos + at.node.nodeSize, { class: "is-partly-selected" }, { formulaPart: { range: part.range } }));
        }
        // A formula inside a text selection is highlighted as a whole (as a selection in the formula rows), not
        // glyph box by glyph box.
        const sel = state.selection;
        if (!sel.empty && !(sel instanceof NodeSelection)) {
          state.doc.nodesBetween(sel.from, sel.to, (node, pos) => {
            if (isFormulaNodeName(node.type.name) && pos >= sel.from && pos + node.nodeSize <= sel.to) {
              decorations.push(Decoration.node(pos, pos + node.nodeSize, { class: "is-in-selection" }));
            }
            return !node.isAtom;
          });
        }
        return decorations.length ? DecorationSet.create(state.doc, decorations) : DecorationSet.empty;
      },
      nodeViews: {
        math_inline: (node, view, getPos, decorations) =>
          new FormulaView(node, view, getPos, host, commands, decorations),
        math_display: (node, view, getPos, decorations) =>
          new FormulaView(node, view, getPos, host, commands, decorations),
      },
    },
  });
}
