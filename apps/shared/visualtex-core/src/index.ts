// VisualTeX editing core (docs/editing-core/DESIGN.md): the one owner of where
// the caret is shown and where things are on the page. The layout engine
// computes the geometry of every caret stop, empty slot and atom while it lays
// the formula out; the kernel and the app ask this module instead of
// measuring the drawn page.
import type { CaretState, EditorGeometry, GeometryRect } from "../../tex-layout/src/index.ts";
import { enterSlot, firstEmptySlotStop, goToStop, moveHorizontal, moveToSlotEdge, space, tab, vertical, type EditingHost, type SpaceKernel, type TabKernel } from "./navigation.ts";
import { hitTest, itemAt, type FormulaPoint } from "./pointer.ts";
import { extend, extendTo, select, selectGroup, selectItem, selectSlot } from "./selection.ts";
import { caretMode, inTextRun } from "./text.ts";
import { ampersandAction, collapseToEnd, enterAction, escapeAction, type AmpersandAction, type EnterAction, type EscapeAction, type KeyModifiers } from "./keys.ts";
import { captureCaret, restoreCaret, type CaretSnapshot } from "./restore.ts";
import { allowedStops, isAllowedStop, type CoreModel } from "./stops.ts";
import { slotOrder, type CoreAtom, type FractionOrder, type Slot, type SlotKey } from "./structures.ts";

export type { EditingHost, SpaceKernel, TabKernel } from "./navigation.ts";
export type { AmpersandAction, EnterAction, EscapeAction, KeyModifiers } from "./keys.ts";
export type { CaretSnapshot } from "./restore.ts";
export type { CoreAtom, SlotKey } from "./structures.ts";

const TEN_PT = 10 * 65536;

/** A rectangle in page (client) pixels. */
export interface ClientRect {
  left: number;
  right: number;
  top: number;
  bottom: number;
}

interface Frame {
  left: number;
  baselineTop: number;
  /** page pixels per sp */
  px: number;
  /** horizontal space (px) the editor's row alignment adds before x */
  align: (x: number) => number;
  /** the alignment markers (sp) and the space (px) added at each */
  markers: number[];
  spaces: number[];
}

// The minimal shape of the kernel mathfield the core reads.
interface KernelMathfield {
  model: CoreModel;
  field?: HTMLElement | null;
  host?: HTMLElement | null;
  visualTexGeometry?: EditorGeometry;
}

export interface EditingCore {
  /** Where the caret is shown: a stop, an empty slot, or nowhere (a range is selected). */
  caretState(mathfield: KernelMathfield): CaretState;
  /** The geometry of the field's last render. */
  geometry(mathfield: KernelMathfield): EditorGeometry | null;
  /** A formula-unit rectangle (sp, y down from the baseline) in page pixels. */
  toClient(mathfield: KernelMathfield, rect: GeometryRect): ClientRect | null;
  /** The page rectangle of an atom (its own material and its descendants'). */
  atomBounds(mathfield: KernelMathfield, atom: unknown): ClientRect | null;
  /** The page rectangle of the caret bar's boundary (zero width), or null. */
  caretRect(mathfield: KernelMathfield): ClientRect | null;
  /** The caret point popovers and menus anchor to: right/bottom of the caret. */
  caretPoint(mathfield: KernelMathfield): { x: number; y: number; height: number } | null;
  /** Where text input lands (the IME candidate window opens beside it): the caret, or a selection's moving end. */
  imeAnchor(mathfield: KernelMathfield): ClientRect | null;

  // --- pointer and selection (DESIGN §6.9) ---
  /** The caret stop a page point means (null: the field has no geometry yet). */
  hitTest(mathfield: KernelMathfield, clientX: number, clientY: number): number | null;
  /** Select from one stop to another, widened to a run of siblings. */
  select(model: CoreModel, anchor: number, focus: number): boolean;
  /** Shift+Left/Right. */
  extend(model: CoreModel, direction: -1 | 1): boolean;
  /** Extend the selection from its anchor to a stop (Shift+click, drags, Shift+Home/End, word steps). */
  extendTo(model: CoreModel, focus: number): boolean;
  /** Double click: select the word or structure at a page point. */
  selectItemAt(mathfield: KernelMathfield, clientX: number, clientY: number): boolean;
  /** Triple click: select the slot at a page point. */
  selectSlotAt(mathfield: KernelMathfield, clientX: number, clientY: number): boolean;
  /** The selectGroup command: the item at the caret. */
  selectGroup(model: CoreModel): boolean;

  // --- structure and movement (DESIGN §4, §6) ---
  /** The keyboard order of an atom's slots (the kernel derives its atom order from it). */
  slotOrder(atom: CoreAtom, fractionOrder: FractionOrder): SlotKey[];
  /** Whether the caret may stand at a kernel offset. */
  isAllowedStop(model: CoreModel, offset: number): boolean;
  /** Every allowed stop, in keyboard order. */
  allowedStops(model: CoreModel): number[];
  moveHorizontal(model: CoreModel, direction: -1 | 1): boolean;
  /** Home/End (thenFormula: a second press goes to the formula's start/end) and Ctrl+Arrow. */
  moveToSlotEdge(model: CoreModel, direction: -1 | 1, thenFormula: boolean): boolean;
  tab(model: CoreModel, direction: -1 | 1, fractionOrder: FractionOrder, kernel: TabKernel): boolean;
  vertical(model: CoreModel, direction: -1 | 1, fractionOrder: FractionOrder, extend?: boolean): boolean;
  /** Space in math: a word space beside text, leave the structure, or insert a visible space at the top level. */
  space(model: CoreModel, kernel: SpaceKernel): boolean;

  // --- text runs (DESIGN §6.6) ---
  /** Whether the caret at a stop stands inside a text run (between two of its characters). */
  inTextRun(model: CoreModel, offset: number): boolean;
  /** The input mode at a stop: text inside a text run, else that of the material there. */
  caretMode(model: CoreModel, offset: number): string;

  // --- Enter and Escape (DESIGN §6.11) ---
  /** What Enter does (rawLatex: the LaTeX box holds typed text). */
  enterAction(model: CoreModel, keys: KeyModifiers, rawLatex: boolean): EnterAction;
  /** What Escape does (rawLatex: the LaTeX box holds typed text). */
  escapeAction(model: CoreModel, keys: KeyModifiers, rawLatex: boolean): EscapeAction;
  /** What & does (alignRows: multi-line formulas are written as align). */
  ampersandAction(model: CoreModel, keys: KeyModifiers, rawLatex: boolean, alignRows: boolean): AmpersandAction;
  /** Collapse the selection to its end. */
  collapseToEnd(model: CoreModel): boolean;

  // --- restore after a rewrite (DESIGN §6.10) ---
  /** The caret and the formula around it, before the formula is replaced. */
  captureCaret(model: CoreModel): CaretSnapshot;
  /** After the replacement: the caret where its atom went (after the replacement when that atom was rewritten). */
  restoreCaret(model: CoreModel, snapshot: CaretSnapshot): boolean;

  // --- empty slots (DESIGN §6.7) ---
  /** Put the caret at an offset; a placeholder's offset means its empty slot. */
  goToStop(model: CoreModel, offset: number): boolean;
  /** The first empty slot within freshly inserted atoms (keyboard order), if any. */
  firstEmptySlotStop(model: CoreModel, atoms: readonly CoreAtom[]): number | undefined;
  /** Enter a slot (a matrix cell for &): at its first empty slot, else at its end. */
  enterSlot(model: CoreModel, slot: Slot, direction: -1 | 1): boolean;
  /** The app's handlers for a field (Up/Down leaving the formula, key commands, ...). */
  setHost(field: Element, host: EditingHost | null): void;
  /** The field's key entry: offer a keydown to the app's key commands; true when one handled it. */
  offerKey(field: Element | undefined, event: KeyboardEvent): boolean;
}

function frameOf(mathfield: KernelMathfield, geometry: EditorGeometry): Frame | null {
  const root = mathfield.field?.querySelector<HTMLElement>(".ML__vtx");
  if (!root) return null;
  const box = root.getBoundingClientRect();
  // geometry.height/depth are the drawn box's (it covers markers and carets)
  const total = geometry.height + geometry.depth;
  const px = box.height > 0 && total > 0 ? box.height / total : parseFloat(getComputedStyle(root).fontSize) / TEN_PT;
  const height = geometry.height;
  const markers = geometry.alignMarkers;
  let spaces: number[] = [];
  if (markers.length > 0) {
    const style = getComputedStyle(mathfield.host ?? root);
    const read = (name: string) => parseFloat(style.getPropertyValue(name)) || 0;
    spaces = markers.map(
      (_, j) => read(`--visualtex-align-marker-${j}-left`) + read(`--visualtex-align-marker-${j}-right`),
    );
  }
  const align = (x: number) => {
    let s = 0;
    for (let j = 0; j < markers.length && markers[j] <= x; j++) s += spaces[j];
    return s;
  };
  return { left: box.left, baselineTop: box.top + height * px, px, align, markers, spaces };
}

/** A page point in formula units (a point inside an alignment space is at its marker). */
function fromClient(frame: Frame, clientX: number, clientY: number): FormulaPoint {
  const y = (clientY - frame.baselineTop) / frame.px;
  let shift = 0;
  for (let j = 0; j < frame.markers.length; j++) {
    const at = frame.left + frame.markers[j] * frame.px + shift;
    if (clientX < at) break;
    if (clientX < at + frame.spaces[j]) return { x: frame.markers[j], y };
    shift += frame.spaces[j];
  }
  return { x: (clientX - frame.left - shift) / frame.px, y };
}

export function createEditingCore(): EditingCore {
  const hosts = new WeakMap<Element, EditingHost>();
  const pointOf = (mathfield: KernelMathfield, clientX: number, clientY: number) => {
    const geometry = core.geometry(mathfield);
    const frame = geometry && frameOf(mathfield, geometry);
    return geometry && frame ? { geometry, point: fromClient(frame, clientX, clientY) } : null;
  };
  const hostOf = (model: CoreModel) => {
    const field = model.mathfield.host;
    return field ? hosts.get(field) : undefined;
  };
  const core: EditingCore = {
    caretState(mathfield) {
      const model = mathfield.model;
      if (model.selectionIsCollapsed) return { kind: "stop", offset: model.position };
      if (model.selectionIsPlaceholder) {
        const p = model.at(Math.max(model.position, model.anchor));
        if (p && (p.type === "placeholder" || p.type === "prompt")) return { kind: "empty", placeholder: p as never };
      }
      return { kind: "none" };
    },

    geometry(mathfield) {
      return mathfield.visualTexGeometry ?? null;
    },

    toClient(mathfield, rect) {
      const geometry = core.geometry(mathfield);
      if (!geometry) return null;
      const frame = frameOf(mathfield, geometry);
      if (!frame) return null;
      return {
        left: frame.left + rect.x0 * frame.px + frame.align(rect.x0),
        right: frame.left + rect.x1 * frame.px + frame.align(rect.x0),
        top: frame.baselineTop + rect.y0 * frame.px,
        bottom: frame.baselineTop + rect.y1 * frame.px,
      };
    },

    atomBounds(mathfield, atom) {
      const region = core.geometry(mathfield)?.regions.get(atom);
      return region ? core.toClient(mathfield, region) : null;
    },

    caretRect(mathfield) {
      const geometry = core.geometry(mathfield);
      if (!geometry) return null;
      const state = core.caretState(mathfield);
      if (state.kind === "stop") {
        const c = geometry.stops[state.offset];
        return c ? core.toClient(mathfield, { x0: c.x, x1: c.x, y0: c.top, y1: c.bottom }) : null;
      }
      if (state.kind === "empty") {
        const m = geometry.markers.get(state.placeholder);
        return m ? core.toClient(mathfield, { x0: m.x0, x1: m.x0, y0: m.y0, y1: m.y1 }) : null;
      }
      return null;
    },

    caretPoint(mathfield) {
      const r = core.caretRect(mathfield);
      return r ? { x: r.right, y: r.bottom, height: r.bottom - r.top } : null;
    },

    imeAnchor(mathfield) {
      const caret = core.caretRect(mathfield);
      if (caret) return caret;
      const c = core.geometry(mathfield)?.stops[mathfield.model.position];
      return c ? core.toClient(mathfield, { x0: c.x, x1: c.x, y0: c.top, y1: c.bottom }) : null;
    },

    hitTest(mathfield, clientX, clientY) {
      const at = pointOf(mathfield, clientX, clientY);
      return at ? hitTest(mathfield.model, at.geometry, at.point) : null;
    },
    select,
    extend,
    extendTo,
    selectItemAt(mathfield, clientX, clientY) {
      const at = pointOf(mathfield, clientX, clientY);
      const item = at ? itemAt(mathfield.model, at.geometry, at.point) : null;
      return item ? selectItem(mathfield.model, item) : false;
    },
    selectSlotAt(mathfield, clientX, clientY) {
      const stop = core.hitTest(mathfield, clientX, clientY);
      return stop === null ? false : selectSlot(mathfield.model, stop);
    },
    selectGroup,

    slotOrder,
    isAllowedStop,
    allowedStops,
    moveHorizontal,
    moveToSlotEdge,
    tab(model, direction, fractionOrder, kernel) {
      return tab(model, direction, { fractionOrder, host: hostOf(model) }, kernel);
    },
    vertical(model, direction, fractionOrder, extend = false) {
      return vertical(model, direction, { fractionOrder, host: hostOf(model) }, extend);
    },
    space,
    inTextRun,
    caretMode,
    enterAction,
    escapeAction,
    ampersandAction,
    collapseToEnd,
    captureCaret,
    restoreCaret,
    goToStop,
    firstEmptySlotStop,
    enterSlot,
    setHost(field, host) {
      if (host) hosts.set(field, host);
      else hosts.delete(field);
    },
    offerKey(field, event) {
      const commands = (field && hosts.get(field)?.keyCommands) ?? [];
      for (const command of commands) if (command(event)) return true;
      return false;
    },
  };
  return core;
}
