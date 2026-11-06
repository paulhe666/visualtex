// Dragging a selection (docs/editing-core/DESIGN.md §6.12), AxMath's drag
// (research: claude-work/axmath-mouse-reverse, drag-drop.md). A press on the
// selection's highlight starts it: the selection stays highlighted and the
// caret where it would land follows the pointer (the click hit test); near the
// edges of the view it scrolls; letting go puts it there and removes it from
// where it was, or copies it when Ctrl is held at the release. A move into the
// selection itself, or into a structure it selects, is refused. The dropped
// part is selected. Esc or the right button cancel. The drag starts with the
// first movement or after 200 ms (MFC's drag delay); let go before, or back on
// the selection, it is a click there (AxMath keeps the selection; VisualTeX
// places the caret, as Word does).
import type { _Mathfield } from './mathfield-private';
import { deleteBackward } from '../editor-model/delete';
import { visualTexSelectionFont, visualTexToggleSelectionStyle } from '../editor-model/array';
import type { InsertOptions } from '../public/core-types';
import { ModeEditor } from './mode-editor';
import { horizontalScrollingAncestor, requestUpdate } from './render';
import { visualTexEditingCore } from './visualtex-hooks';

/** A place a dragged selection would land: a formula field and a caret stop in it. */
export interface DragTarget {
  mathfield: _Mathfield;
  stop: number;
}

/** What one drag needs from its source: its targets, its edits and its scrolling. */
export interface DragDriver {
  /** Where a page point lands; null: nowhere. */
  resolve(clientX: number, clientY: number): DragTarget | null;
  /** A move there would put the selection into itself (refused). */
  inside(target: DragTarget): boolean;
  /** The containers that scroll when the pointer nears their edges. */
  scrollers(): HTMLElement[];
  /** Let go on a target: move the selection there, or copy it. */
  commit(target: DragTarget, copy: boolean): void;
  /** Let go without a drag, or on the selection itself: a click at that point. */
  click(clientX: number, clientY: number): void;
  /** The drag is over (dropped, refused or cancelled). */
  end?(): void;
  /** Where the pointer leaves every target the last target stays (AxMath's caret stays where it was). */
  keepLastTarget?: boolean;
  /** The editor's area: outside it there is no target, and letting go drops nothing (AxMath). */
  area?(): PageArea | null;
}

/** A page rectangle in CSS pixels. */
interface PageArea {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/** MFC's drag delay: a press held this long starts the drag without moving. */
const DRAG_DELAY_MS = 200;
/** AxMath's edge band (device pixels) and its scroll counter: a run of 5 ticks in the band starts scrolling. */
const SCROLL_BAND_DEVICE_PX = 50;
const SCROLL_COUNT_LIMIT = 20;
const SCROLL_COUNT_START = 5;
const SCROLL_STEP_PX = 20;
const SCROLL_TICK_MS = 50;

interface PixelRect {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/**
 * The selection's highlight in device pixels, as AxMath keeps it for this
 * test: each highlight box, merged with the one before when both have the same
 * top and bottom (across a gap between them too).
 */
function selectionPixelRects(mathfield: _Mathfield): PixelRect[] {
  const scale = globalThis.devicePixelRatio || 1;
  // the painter inserts each box before the previous one
  const boxes = Array.from(mathfield.field?.querySelectorAll<HTMLElement>('.ML__selection') ?? []).reverse();
  const rects: PixelRect[] = [];
  for (const box of boxes) {
    const r = box.getBoundingClientRect();
    if (r.width <= 0 || r.height <= 0) continue;
    const rect = {
      left: Math.trunc(r.left * scale),
      top: Math.trunc(r.top * scale),
      right: Math.trunc(r.right * scale),
      bottom: Math.trunc(r.bottom * scale),
    };
    const last = rects[rects.length - 1];
    if (last && last.top === rect.top && last.bottom === rect.bottom) {
      last.left = Math.min(last.left, rect.left);
      last.right = Math.max(last.right, rect.right);
    } else rects.push(rect);
  }
  return rects;
}

/** Whether a page point is on the field's selection highlight (left and top edges in, right and bottom out). */
export function pointOnSelection(mathfield: _Mathfield, clientX: number, clientY: number): boolean {
  const model = mathfield.model;
  if (model.selectionIsCollapsed || model.selectionIsPlaceholder) return false;
  const scale = globalThis.devicePixelRatio || 1;
  const x = Math.trunc(clientX * scale);
  const y = Math.trunc(clientY * scale);
  return selectionPixelRects(mathfield).some((r) => x >= r.left && x < r.right && y >= r.top && y < r.bottom);
}

/** The drop caret: a bar at the stop where the selection would land. */
class DropCaret {
  private bar: HTMLElement | null = null;

  show(target: DragTarget | null): void {
    const rect = target && target.stop >= 0 ? visualTexEditingCore.stopRect(target.mathfield, target.stop) : null;
    if (!target || !rect) {
      this.clear();
      return;
    }
    if (!this.bar) {
      this.bar = document.createElement('div');
      this.bar.className = 'visualtex-drop-caret';
      this.bar.style.cssText =
        'position:fixed;z-index:2147483000;width:2px;pointer-events:none;border-radius:1px;' +
        'background:var(--formula-caret, #1f638e);box-shadow:0 0 0 1px rgba(255,255,255,0.8)';
      document.body.append(this.bar);
    }
    this.bar.style.left = `${rect.left - 1}px`;
    this.bar.style.top = `${rect.top}px`;
    this.bar.style.height = `${Math.max(8, rect.bottom - rect.top)}px`;
  }

  clear(): void {
    this.bar?.remove();
    this.bar = null;
  }
}

/** AxMath's edge scrolling: per container and axis a counter in [-20, 20]; it scrolls from 5 on. */
export class EdgeScroller {
  private counts = new Map<HTMLElement, { x: number; y: number }>();

  tick(scrollers: HTMLElement[], clientX: number, clientY: number): boolean {
    const band = SCROLL_BAND_DEVICE_PX / (globalThis.devicePixelRatio || 1);
    let scrolled = false;
    for (const scroller of scrollers) {
      const r = scroller.getBoundingClientRect();
      const count = this.counts.get(scroller) ?? { x: 0, y: 0 };
      const inside = clientX >= r.left && clientX < r.right && clientY >= r.top && clientY < r.bottom;
      const step = (value: number, low: number, high: number, current: number) => {
        if (!inside) return 0;
        const direction = value < low + band ? -1 : value >= high - band ? 1 : 0;
        if (!direction) return 0;
        return Math.max(-SCROLL_COUNT_LIMIT, Math.min(SCROLL_COUNT_LIMIT, current + direction));
      };
      count.x = scroller.scrollWidth > scroller.clientWidth ? step(clientX, r.left, r.right, count.x) : 0;
      count.y = scroller.scrollHeight > scroller.clientHeight ? step(clientY, r.top, r.bottom, count.y) : 0;
      this.counts.set(scroller, count);
      const dx = Math.abs(count.x) >= SCROLL_COUNT_START ? Math.sign(count.x) * SCROLL_STEP_PX : 0;
      const dy = Math.abs(count.y) >= SCROLL_COUNT_START ? Math.sign(count.y) * SCROLL_STEP_PX : 0;
      if (!dx && !dy) continue;
      const left = scroller.scrollLeft;
      const top = scroller.scrollTop;
      scroller.scrollBy({ left: dx, top: dy });
      if (scroller.scrollLeft !== left || scroller.scrollTop !== top) scrolled = true;
    }
    return scrolled;
  }
}

/** Runs one drag of a selection from the press that started it. */
export function runDrag(driver: DragDriver, press: PointerEvent): void {
  const pointerId = press.pointerId;
  const startX = press.clientX;
  const startY = press.clientY;
  const scale = globalThis.devicePixelRatio || 1;
  const caret = new DropCaret();
  const scroller = new EdgeScroller();
  let started = false;
  let finished = false;
  let target: DragTarget | null = null;
  let last = { x: startX, y: startY, copy: press.ctrlKey };

  const place = () => {
    const area = driver.area?.();
    const outside = Boolean(area) &&
      !(last.x >= area!.left && last.x < area!.right && last.y >= area!.top && last.y < area!.bottom);
    const found = outside ? null : driver.resolve(last.x, last.y);
    if (found || outside || !driver.keepLastTarget) target = found;
    caret.show(target);
    document.documentElement.style.setProperty('cursor', last.copy ? 'copy' : 'move', 'important');
  };
  const begin = () => {
    if (started || finished) return;
    started = true;
    place();
  };
  const delay = setTimeout(begin, DRAG_DELAY_MS);
  const ticker = setInterval(() => {
    if (started && scroller.tick(driver.scrollers(), last.x, last.y)) place();
  }, SCROLL_TICK_MS);

  const finish = () => {
    if (finished) return;
    finished = true;
    clearTimeout(delay);
    clearInterval(ticker);
    window.removeEventListener('pointermove', onMove, true);
    window.removeEventListener('pointerup', onUp, true);
    window.removeEventListener('pointercancel', onCancel, true);
    window.removeEventListener('pointerdown', onOtherButton, true);
    window.removeEventListener('mousedown', onOtherButton, true);
    window.removeEventListener('keydown', onKey, true);
    caret.clear();
    document.documentElement.style.removeProperty('cursor');
    driver.end?.();
  };
  const cancel = () => {
    finish();
  };
  const onMove = (e: PointerEvent) => {
    if (e.pointerId !== pointerId || finished) return;
    if (e.buttons & 2) {
      cancelByButton();
      return;
    }
    // the left button is no longer held: its release went past us (outside
    // the window); nothing is dropped
    if (!(e.buttons & 1)) {
      cancel();
      return;
    }
    last = { x: e.clientX, y: e.clientY, copy: e.ctrlKey };
    // AxMath's drag starts once the pointer leaves the pixel it was pressed in
    if (!started && Math.trunc(e.clientX * scale) === Math.trunc(startX * scale) &&
        Math.trunc(e.clientY * scale) === Math.trunc(startY * scale)) return;
    e.preventDefault();
    e.stopImmediatePropagation();
    if (!started) {
      begin();
      return;
    }
    scroller.tick(driver.scrollers(), last.x, last.y);
    place();
  };
  const onUp = (e: PointerEvent) => {
    if (e.pointerId !== pointerId || finished) return;
    const wasStarted = started;
    const copy = e.ctrlKey;
    last = { x: e.clientX, y: e.clientY, copy };
    if (wasStarted) place();
    const landing = target;
    finish();
    e.preventDefault();
    e.stopImmediatePropagation();
    if (!wasStarted) {
      driver.click(e.clientX, e.clientY);
      return;
    }
    if (!landing) return;
    if (!copy && driver.inside(landing)) {
      // back on the selection: a click there
      driver.click(e.clientX, e.clientY);
      return;
    }
    driver.commit(landing, copy);
  };
  const onCancel = (e: PointerEvent) => {
    if (e.pointerId === pointerId) cancel();
  };
  const onKey = (e: KeyboardEvent) => {
    if (e.key !== 'Escape') return;
    e.preventDefault();
    e.stopImmediatePropagation();
    cancel();
  };
  // the right button cancels (as a Windows drag does) and opens no context menu
  const cancelByButton = () => {
    cancel();
    window.addEventListener('contextmenu', (ev) => ev.preventDefault(), { once: true, capture: true });
  };
  const onOtherButton = (e: MouseEvent) => {
    if (e.button === 0) return;
    e.preventDefault();
    e.stopImmediatePropagation();
    cancelByButton();
  };
  window.addEventListener('pointermove', onMove, true);
  window.addEventListener('pointerup', onUp, true);
  window.addEventListener('pointercancel', onCancel, true);
  window.addEventListener('pointerdown', onOtherButton, true);
  window.addEventListener('mousedown', onOtherButton, true);
  window.addEventListener('keydown', onKey, true);
}

/** The app's side of a drag out of a field (EditingHost.selectionDrag): its other formulas. */
interface FieldDragHost {
  target?(clientX: number, clientY: number): { field: Element; stop: number } | null | undefined;
  area?(): PageArea | null;
  scrollers?(): HTMLElement[];
  drop?(source: Element, start: number, end: number, latex: string, target: { field: Element; stop: number }, copy: boolean): void;
  outside?(clientX: number, clientY: number, source: Element, start: number, end: number): boolean;
  dropOutside?(source: Element, start: number, end: number, latex: string, copy: boolean): void;
  end?(): void;
}

/** The stop of a target that is the app's own place outside every formula (FieldDragHost.outside). */
const OUTSIDE = -1;

const kernelOf = (field: Element) => (field as unknown as { _mathfield?: _Mathfield })._mathfield;

const DROP_INSERT: InsertOptions = {
  format: 'latex',
  mode: 'math',
  insertionMode: 'replaceSelection',
  selectionMode: 'item',
  silenceNotifications: true,
};

/** LaTeX put in at a stop (an empty slot's: in place of its placeholder); the inserted part is selected. */
function insertAt(mathfield: _Mathfield, stop: number, latex: string): [number, number] {
  const model = mathfield.model;
  visualTexEditingCore.goToStop(model, stop);
  ModeEditor.insert(model, latex, DROP_INSERT);
  const [a, b] = model.selection.ranges[0];
  return [Math.min(a, b), Math.max(a, b)];
}

/** A range taken out (a slot left empty keeps its placeholder, a script structure goes whole). */
function removeRange(mathfield: _Mathfield, start: number, end: number): void {
  mathfield.model.setSelection(start, end);
  deleteBackward(mathfield.model);
}

/** Runs edits on a field's model without telling anyone (one notification follows, or the model is put back). */
function silently<T>(mathfield: _Mathfield, edit: () => T): T {
  const model = mathfield.model;
  const silenced = model.silenceNotifications;
  model.silenceNotifications = true;
  try {
    return edit();
  } finally {
    model.silenceNotifications = silenced;
  }
}

/** The selection [start, end] (its LaTeX `latex`) to a stop of the same formula, moved or copied: one edit. */
function dropWithin(mathfield: _Mathfield, latex: string, start: number, end: number, target: number, copy: boolean): void {
  const model = mathfield.model;
  mathfield.stopCoalescingUndo();
  silently(mathfield, () => {
    if (copy) insertAt(mathfield, target, latex);
    else if (target > end) {
      // inserted first (after the selection, its offsets stay), then the
      // selection goes: the copy moves left by what that removed (a slot left
      // empty keeps a placeholder)
      const inserted = insertAt(mathfield, target, latex);
      const before = model.lastOffset;
      removeRange(mathfield, start, end);
      const shift = before - model.lastOffset;
      model.setSelection(inserted[0] - shift, inserted[1] - shift);
    } else {
      removeRange(mathfield, start, end);
      insertAt(mathfield, target, latex);
    }
  });
  // one edit: one undo step and one input event
  mathfield.snapshot('drop');
  model.contentDidChange({ inputType: 'insertFromDrop' });
  model.selectionDidChange();
  requestUpdate(mathfield);
}

/** The editing core's selection from stop `anchor` to stop `focus` (whole structures where they part), drawn. */
export function selectStops(mathfield: _Mathfield, anchor: number, focus: number): void {
  visualTexEditingCore.select(mathfield.model, anchor, focus);
  requestUpdate(mathfield);
}

/** The caret to a stop (an empty slot's: its placeholder selected), drawn. */
export function goToStopDrawn(mathfield: _Mathfield, stop: number): boolean {
  const moved = visualTexEditingCore.goToStop(mathfield.model, stop);
  requestUpdate(mathfield);
  return moved;
}

/** What a drop makes of a formula (the app's rows): see visualTexDropResult. */
export interface DropEdit {
  /** The formula to start from instead of the field's own. */
  base?: string;
  /** A range taken out. */
  remove?: [number, number];
  /** LaTeX put in at a stop (after the removal, if both). */
  insert?: { at: number; latex: string };
}

/** A formula after a drop's edits: its LaTeX, the range put in (if any) and its last offset. */
export interface DropResult {
  latex: string;
  inserted: [number, number] | null;
  lastOffset: number;
}

/**
 * A formula's LaTeX after a drop's edits, and the range of what was put in,
 * worked out on its model, which is then put back as it was: the app makes
 * the edit across its rows as one step of its own.
 */
export function dropResult(mathfield: _Mathfield, edit: DropEdit): DropResult {
  const model = mathfield.model;
  const state = model.getState();
  return silently(mathfield, () => {
    try {
      if (edit.base !== undefined)
        ModeEditor.insert(model, edit.base, { ...DROP_INSERT, insertionMode: 'replaceAll', selectionMode: 'after' });
      if (edit.remove) removeRange(mathfield, Math.min(...edit.remove), Math.max(...edit.remove));
      const inserted = edit.insert ? insertAt(mathfield, edit.insert.at, edit.insert.latex) : null;
      return { latex: model.getValue(), inserted, lastOffset: model.lastOffset };
    } finally {
      model.setState(state, { silenceNotifications: true });
      // the restored atoms are new ones: lay them out again
      requestUpdate(mathfield);
    }
  });
}

/**
 * A press on a field's selection: drag it. Within the formula the move is one
 * edit; the app (EditingHost.selectionDrag) may offer its other formulas.
 */
export function startFieldSelectionDrag(mathfield: _Mathfield, press: PointerEvent): void {
  const core = visualTexEditingCore;
  const model = mathfield.model;
  const range = model.selection.ranges[0];
  const start = Math.min(range[0], range[1]);
  const end = Math.max(range[0], range[1]);
  const latex = model.getValue([start, end], 'latex');
  const host = core.selectionDragHost?.(mathfield.host) as FieldDragHost | undefined;
  const element = mathfield.host;
  // the app's own pointer handling (a selection across rows) stands aside while this drag lasts
  element?.classList.add('visualtex-selection-drag');
  runDrag(
    {
      keepLastTarget: Boolean(host?.target),
      area: host?.area ? () => host.area!() : undefined,
      resolve(x, y) {
        // outside the formula, a place of the app's own (a document's text): it marks the place itself
        if (element && host?.outside?.(x, y, element, start, end)) return { mathfield, stop: OUTSIDE };
        const offered = host?.target?.(x, y);
        if (offered !== undefined) {
          const kernel = offered && kernelOf(offered.field);
          return kernel && offered ? { mathfield: kernel, stop: offered.stop } : null;
        }
        // a formula on its own: its own row, beside it its start or end
        const row = (element ?? mathfield.field)?.getBoundingClientRect();
        if (!row || y < row.top || y >= row.bottom) return null;
        const stop = core.hitTest(mathfield, x, y);
        return stop === null ? null : { mathfield, stop };
      },
      // its own edges too: a move there changes nothing
      inside: (target) => target.mathfield === mathfield && target.stop !== OUTSIDE && target.stop >= start && target.stop <= end,
      scrollers: () => {
        const own = element ? horizontalScrollingAncestor(element) : null;
        return [...(own ? [own] : []), ...(host?.scrollers?.() ?? [])];
      },
      commit(target, copy) {
        if (target.stop === OUTSIDE) {
          if (element) host?.dropOutside?.(element, start, end, latex, copy);
          return;
        }
        if (target.mathfield === mathfield) {
          dropWithin(mathfield, latex, start, end, target.stop, copy);
          if (!mathfield.hasFocus()) mathfield.onFocus();
          return;
        }
        if (element && target.mathfield.host) host?.drop?.(element, start, end, latex, { field: target.mathfield.host, stop: target.stop }, copy);
      },
      click(x, y) {
        const stop = core.hitTest(mathfield, x, y);
        if (stop === null) return;
        core.goToStop(model, stop);
        requestUpdate(mathfield);
        if (!mathfield.hasFocus()) mathfield.onFocus();
      },
      end() {
        element?.classList.remove('visualtex-selection-drag');
        host?.end?.();
      },
    },
    press,
  );
}

/** Whether every glyph of a range is bold (italic); null when it holds no glyph. */
export function rangeFont(mathfield: _Mathfield, range: [number, number]): { bold: boolean; italic: boolean } | null {
  return visualTexSelectionFont(mathfield.model, { ranges: [[Math.min(...range), Math.max(...range)]], direction: 'forward' });
}

/**
 * A formula's LaTeX with a range set bold (italic) or not, worked out on its
 * model, which is then put back (the app makes the edit with its other ones).
 */
export function styledResult(mathfield: _Mathfield, range: [number, number], kind: 'bold' | 'italic', enable: boolean): string {
  const model = mathfield.model;
  const state = model.getState();
  const selection = { ranges: [[Math.min(...range), Math.max(...range)]] as [number, number][], direction: 'forward' as const };
  return silently(mathfield, () => {
    try {
      const font = visualTexSelectionFont(model, selection);
      if (font && font[kind] !== enable) visualTexToggleSelectionStyle(model, kind, selection);
      return model.getValue();
    } finally {
      model.setState(state, { silenceNotifications: true });
      requestUpdate(mathfield);
    }
  });
}

/** A target of a drag the app drives: a field (the element) and a stop in it. */
export interface AppDragTarget {
  field: Element;
  stop: number;
}

/** A drag of the app's own selection (one across its rows): as DragDriver, with its targets. */
export interface AppDragDriver {
  resolve(clientX: number, clientY: number): AppDragTarget | null;
  area?(): PageArea | null;
  inside(target: AppDragTarget): boolean;
  scrollers(): HTMLElement[];
  commit(target: AppDragTarget, copy: boolean): void;
  click(clientX: number, clientY: number): void;
  end?(): void;
}

/** Runs a drag of a selection the app holds (DESIGN §6.12): the same gesture, caret and scrolling as a field's. */
export function runAppDrag(driver: AppDragDriver, press: PointerEvent): void {
  const toKernel = (target: AppDragTarget | null): DragTarget | null => {
    const mathfield = target ? kernelOf(target.field) : undefined;
    return mathfield && target ? { mathfield, stop: target.stop } : null;
  };
  const toApp = (target: DragTarget): AppDragTarget | null =>
    target.mathfield.host ? { field: target.mathfield.host, stop: target.stop } : null;
  runDrag(
    {
      // AxMath's caret stays where it was while the pointer is below every row
      keepLastTarget: true,
      area: driver.area ? () => driver.area!() : undefined,
      resolve: (x, y) => toKernel(driver.resolve(x, y)),
      inside: (target) => {
        const app = toApp(target);
        return !app || driver.inside(app);
      },
      scrollers: () => driver.scrollers(),
      commit(target, copy) {
        const app = toApp(target);
        if (app) driver.commit(app, copy);
      },
      click: (x, y) => driver.click(x, y),
      end: () => driver.end?.(),
    },
    press,
  );
}
