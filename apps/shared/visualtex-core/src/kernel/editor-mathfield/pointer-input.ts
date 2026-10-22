import type { _Mathfield } from './mathfield-private';
import { horizontalScrollingAncestor, requestUpdate } from './render';
import { acceptCommandSuggestion } from './autocomplete';
import { visualTexEditingCore } from './visualtex-hooks';
import type { Offset } from '../public/core-types';

let gLastTap: { x: number; y: number; time: number } | null = null;
let gTapCount = 0;

export class PointerTracker {
  private static controller: AbortController | undefined;
  private static pointerId: number | undefined;
  private static element: HTMLElement | undefined;
  private static onCancel: ((ev: Event) => void) | undefined;

  /** The element whose pointer is being tracked, if any. */
  static get trackedElement(): HTMLElement | undefined {
    return PointerTracker.controller ? PointerTracker.element : undefined;
  }

  static start(
    element: HTMLElement,
    evt: Event,
    onMove: (this: HTMLElement, ev: PointerEvent | MouseEvent) => any,
    onCancel: (this: HTMLElement, ev: Event) => any
  ): void {
    PointerTracker.element = element;
    PointerTracker.onCancel = (ev: Event) => (onCancel as (ev: Event) => void).call(element, ev);

    // Have to create a new controller each time, as they can only be used once
    PointerTracker.controller?.abort();
    PointerTracker.controller = new AbortController();

    const options = { signal: PointerTracker.controller.signal };

    if ('PointerEvent' in window) {
      element.addEventListener('pointermove', onMove, options);
      // a quick drag can end without a final pointermove: the release is its last sample
      element.addEventListener('pointerup', (upEvent) => {
        (onMove as (ev: PointerEvent) => void)(upEvent);
        (onCancel as (ev: Event) => void)(upEvent);
      }, options);
      element.addEventListener('pointercancel', onCancel, options);
      if (isPointerEvent(evt)) {
        PointerTracker.pointerId = evt.pointerId;
        element.setPointerCapture(evt.pointerId);
      }
    } else {
      // @ts-ignore
      window.addEventListener('mousemove', onMove, options);
      // @ts-ignore
      window.addEventListener('blur', onCancel, options);
      // @ts-ignore
      window.addEventListener('mouseup', (upEvent: MouseEvent) => {
        (onMove as (ev: MouseEvent) => void)(upEvent);
        (onCancel as (ev: Event) => void)(upEvent);
      }, options);
    }
  }

  /** Ends the tracking as a cancelled gesture (the host took the drag over). */
  static cancel(): void {
    const onCancel = PointerTracker.onCancel;
    if (!PointerTracker.controller || !onCancel) return;
    onCancel(new Event('pointercancel'));
    PointerTracker.stop();
  }

  static stop(): void {
    PointerTracker.controller?.abort();
    PointerTracker.controller = undefined;
    PointerTracker.onCancel = undefined;

    if (typeof PointerTracker.pointerId === 'number') {
      PointerTracker.element!.releasePointerCapture(PointerTracker.pointerId);
      PointerTracker.pointerId = undefined;
    }
  }
}

function isPointerEvent(evt: Event | null): evt is PointerEvent {
  return (
    evt !== null &&
    globalThis.PointerEvent !== undefined &&
    evt instanceof PointerEvent
  );
}

function pointInRect(x: number, y: number, rect: DOMRect): boolean {
  return x >= rect.left && x <= rect.right && y >= rect.top && y <= rect.bottom;
}

export function onPointerDown(mathfield: _Mathfield, evt: PointerEvent): void {
  // one event can reach two listeners (the field and its host): handle it once
  if (mathfield.visualTexLastPointerEvent === evt) return;
  mathfield.visualTexLastPointerEvent = evt;

  // If a mouse button other than the main one was pressed, return.
  if (evt.buttons > 1) return;

  // VisualTeX editing core (pointer.ts, selection.ts): where a point is and
  // what a click, a drag or a double/triple click selects
  const core = visualTexEditingCore;
  const that = mathfield;
  let trackingPointer = false;
  let dirty: 'none' | 'selection' | 'all' = 'none';

  // Scroll right or left when the pointer is past the visible edge
  let scrollLeft = false;
  let scrollRight = false;
  // Note: evt['detail'] contains the number of clicks
  // for this event. However, it's not reliable on touch devices
  // (e.g. iOS), so we track it ourselves.
  const anchorX = evt.clientX;
  const anchorY = evt.clientY;
  const anchorTime = Date.now();
  const field = that.field!;
  // a drag past the visible edge scrolls the formula's row (the field itself
  // grows to the formula's width and never scrolls)
  const scroller = horizontalScrollingAncestor(mathfield.host!) ?? field;
  let lastPointer: { x: number; y: number } | null = null;
  const scrollInterval = setInterval(() => {
    // while it scrolls, the selection follows the formula to the visible edge
    // (the render's caret follow-up stays out of it: steady speed)
    mathfield.visualTexDragScrolling = (scrollLeft || scrollRight) && lastPointer !== null;
    if (!mathfield.visualTexDragScrolling) return;
    scroller.scrollBy({ top: 0, left: scrollLeft ? -16 : 16 });
    const visible = scroller.getBoundingClientRect();
    const focus = core.hitTest(that, scrollLeft ? visible.left + 1 : visible.right - 1, lastPointer!.y);
    if (focus !== null) {
      core.extendTo(that.model, focus);
      requestUpdate(mathfield);
    }
  }, 32);

  function endPointerTracking(): void {
    PointerTracker.stop();
    trackingPointer = false;
    clearInterval(scrollInterval);
    mathfield.visualTexDragScrolling = false;
    mathfield.element!.classList.remove('tracking');
    if (evt) evt.preventDefault();
  }

  function onPointerMove(evt: PointerEvent | MouseEvent): void {
    // If we've somehow lost focus, end tracking
    if (!that.hasFocus()) {
      endPointerTracking();
      return;
    }

    const x = evt.clientX;
    const y = evt.clientY;
    // Ignore events that are within small spatial and temporal bounds
    // of the pointer down
    const hysteresis =
      isPointerEvent(evt) && evt.pointerType === 'touch' ? 20 : 5;
    if (
      Date.now() < anchorTime + 500 &&
      Math.abs(anchorX - x) < hysteresis &&
      Math.abs(anchorY - y) < hysteresis
    ) {
      evt.preventDefault();
      evt.stopPropagation();
      return;
    }

    const visible = scroller.getBoundingClientRect();
    scrollRight = x > visible.right;
    scrollLeft = x < visible.left;
    lastPointer = { x, y };
    const focus = core.hitTest(that, x, y);
    if (focus !== null) {
      core.extendTo(that.model, focus);
      requestUpdate(mathfield);
    }
  }

  // Calculate the tap count
  if (
    gLastTap &&
    Math.abs(gLastTap.x - anchorX) < 5 &&
    Math.abs(gLastTap.y - anchorY) < 5 &&
    Date.now() < gLastTap.time + 500
  ) {
    gTapCount += 1;
    gLastTap.time = anchorTime;
  } else {
    gLastTap = {
      x: anchorX,
      y: anchorY,
      time: anchorTime,
    };
    gTapCount = 1;
  }

  const inside = [field, mathfield.container, mathfield.element].some(
    (element) => element && pointInRect(anchorX, anchorY, element.getBoundingClientRect())
  );
  if (inside) {
    // Clear out any inline shortcut buffer
    mathfield.flushInlineShortcutBuffer();

    const stop = core.hitTest(mathfield, anchorX, anchorY);
    if (stop !== null) {
      // Reset any user-specified style
      if (stop !== mathfield.model.anchor) {
        const preservedVariantStyle = mathfield.defaultStyle.variantStyle;
        mathfield.defaultStyle = preservedVariantStyle !== undefined ? { variantStyle: preservedVariantStyle } : {};
        mathfield.styleBias = 'left';
      }
      mathfield.element!.classList.add('tracking');
      if (evt.shiftKey) {
        // Extend the selection if the shift-key is down
        const wasCollapsed = mathfield.model.selectionIsCollapsed;
        core.extendTo(mathfield.model, stop);
        if (acceptCommandSuggestion(mathfield.model) || wasCollapsed) dirty = 'all';
        else dirty = 'selection';
      } else {
        core.goToStop(mathfield.model, stop);
        if (acceptCommandSuggestion(mathfield.model)) dirty = 'all';
        else dirty = 'selection';
      }
      // `evt.detail` contains the number of consecutive clicks
      // for double-click, triple-click, etc...
      // (note that evt.detail is not set when using pointerEvent)
      if (evt.detail === 3 || gTapCount > 2) {
        endPointerTracking();
        if (evt.detail === 3 || gTapCount === 3) {
          // This is a triple-click: select the slot
          core.selectSlotAt(mathfield, anchorX, anchorY);
          dirty = 'all';
        }
      } else if (!trackingPointer) {
        trackingPointer = true;
        PointerTracker.start(field, evt, onPointerMove, endPointerTracking);
        if (evt.detail === 2 || gTapCount === 2) {
          // This is a double-click: select the item under the pointer
          core.selectItemAt(mathfield, anchorX, anchorY);
          dirty = 'all';
        }
      }
    }
    if (!mathfield.hasFocus()) {
      dirty = 'none'; // focus() will refresh
      mathfield.onFocus();
      mathfield.model.announce('line');
    }
  } else gLastTap = null;

  mathfield.stopCoalescingUndo();

  // Prevent the browser from handling. In particular when this is a
  // touch event, prevent the synthetic mouseDown event from being generated
  if (dirty !== 'none') {
    if (mathfield.model.selectionIsCollapsed) dirty = 'all';
    requestUpdate(mathfield);
  }

  // Prevent the browser from handling. In particular when this is a
  // touch event, prevent the synthetic mouseDown event from being generated
  evt.preventDefault();
}

/**
 * The offset of the caret stop at a point in viewport coordinates
 * (VisualTeX editing core, pointer.ts: the one hit test of the field).
 */
export function offsetFromPoint(
  mathfield: _Mathfield,
  x: number,
  y: number,
  _options?: { bias?: -1 | 0 | 1 }
): Offset {
  return visualTexEditingCore.hitTest(mathfield, x, y) ?? 0;
}
