import {
  Rect,
  getSelectionBounds,
  isValidMathfield,
  getAtomBounds,
  adjustForScrolling,
} from './utils';
import type { _Mathfield } from './mathfield-private';

import { updateSuggestionPopoverPosition } from '../editor/suggestion-popover';
import { gFontsState } from '../core/fonts';
import { Atom } from '../core/atom-class';
import { convertLatexToMarkup } from '../public/mathlive';
import { ModeEditor } from './mode-editor';
import { visualTexEditingCore, visualTexLayoutEngine } from './visualtex-hooks';

export function requestUpdate(
  mathfield: _Mathfield | undefined | null,
  options?: { interactive: boolean }
): void {
  if (!mathfield || mathfield.dirty || !mathfield.field) return;
  mathfield.resizeObserver.unobserve(mathfield.field);
  mathfield.dirty = true;
  requestAnimationFrame(() => {
    if (isValidMathfield(mathfield) && mathfield.dirty) {
      render(mathfield, options);
      mathfield.resizeObserver.observe(mathfield.field);
      mathfield.resizeObserverStarted = true;
    }
  });
}

export function contentMarkup(
  mathfield: _Mathfield,
  renderOptions?: { forHighlighting?: boolean; interactive?: boolean }
): string {
  try {
    //
    // 1. Update selection state and blinking cursor (caret)
    //
    const { model } = mathfield;
    model.root.caret = undefined;
    model.root.isSelected = false;
    model.root.containsCaret = true;
    for (const atom of model.atoms) {
      atom.caret = undefined;
      atom.isSelected = false;
      atom.containsCaret = false;
    }
    if (model.selectionIsCollapsed) {
      const atom = model.at(model.position);
      atom.caret = mathfield.model.mode;
      let ancestor = atom.parent;
      while (ancestor) {
        ancestor.containsCaret = true;
        ancestor = ancestor.parent;
      }
    } else {
      const atoms = model.getAtoms(model.selection, {
        includeChildren: true,
        includeFirstAtoms: true,
      });
      for (const atom of atoms) atom.isSelected = true;
    }

    //
    // 2. The layout engine lays out and draws the formula
    //
    return visualTexLayoutEngine.renderMathfield(mathfield, renderOptions);
  } catch (e) {
    console.error(e);
    return '<span class="ML__latex" translate="no" aria-hidden="true">🚫</span>';
  }
}

/**
 * Layout the mathfield and generate the DOM.
 *
 * This is usually done automatically, but if the font-size, or other geometric
 * attributes are modified, outside of MathLive, this function may need to be
 * called explicitly.
 *
 */
export function render(
  mathfield: _Mathfield,
  renderOptions?: { forHighlighting?: boolean; interactive?: boolean }
): void {
  if (!isValidMathfield(mathfield)) return;
  renderOptions ??= {};

  // NVA tries (and fails) to read MathML, so skip it for now
  // mathfield.accessibleMathML.innerHTML = mathfield.options.createHTML(
  //   '<math xmlns="http://www.w3.org/1998/Math/MathML">' +
  //     toMathML(model.root, mathfield.options) +
  //     '</math>'
  // );

  //
  // 2. Render the content
  //

  const field = mathfield.field;
  if (!field) return;

  const hasFocus = mathfield.isSelectionEditable && mathfield.hasFocus();
  const isFocused = field.classList.contains('ML__focused');

  if (isFocused && !hasFocus) field.classList.remove('ML__focused');
  else if (!isFocused && hasFocus) field.classList.add('ML__focused');

  let content = contentMarkup(mathfield, renderOptions);

  //
  // 3. Render the content placeholder, if applicable
  //
  // If the mathfield is empty, display a placeholder
  if (mathfield.model.atoms.length <= 1) {
    const placeholder = mathfield.options.contentPlaceholder;
    if (placeholder) {
      content += `<span part=placeholder class="ML__content-placeholder">${convertLatexToMarkup(
        placeholder
      )}</span>`;
    }
  }

  field.innerHTML = globalThis.MathfieldElement.createHTML(content);
  mathfield.container?.classList.toggle('has-visualtex-raw-latex-command', mathfield.model.mode === 'latex');

  //
  // 4. Render the selection/caret
  //
  renderSelection(mathfield, renderOptions.interactive);

  mathfield.dirty = false;
  if (hasFocus) scheduleCaretFollowUp(mathfield);
}

// VisualTeX kernel: a long formula line grows the math-field to its full
// width and scrolls inside an outer container (the editor's formula row).
// MathLive only scrolled its own field, which never overflows, so the caret
// could leave the visible part of the row. Every render of a focused field
// and every scrollIntoView keeps the caret (or the moving end of a range
// selection) inside the nearest horizontally scrolling ancestor. The check
// runs as a microtask: the editor's mutation observer restyles the new glyphs
// (formula letter font) first, and only then is the line width final.
const CARET_SCROLL_MARGIN = 24;

export function horizontalScrollingAncestor(element: Element): HTMLElement | null {
  let node: Node | null = element;
  while (node) {
    const parent: Element | null = (node as Element).parentElement;
    node = parent ?? (node.parentNode instanceof ShadowRoot ? node.parentNode.host : null);
    if (!node || node === document.body || node === document.documentElement)
      return null;
    if ((node as HTMLElement).scrollWidth > (node as HTMLElement).clientWidth + 1) {
      const overflowX = getComputedStyle(node as Element).overflowX;
      if (overflowX === 'auto' || overflowX === 'scroll') return node as HTMLElement;
    }
  }
  return null;
}

function caretHorizontalBounds(mathfield: _Mathfield): { left: number; right: number } | null {
  if (!visualTexEditingCore) return null;
  const caret = visualTexEditingCore.caretRect(mathfield);
  if (caret) return { left: caret.left, right: caret.right };
  // A range: the end that is moving (the atom at `position`).
  const bounds = visualTexEditingCore.atomBounds(mathfield, mathfield.model.at(mathfield.model.position));
  return bounds ? { left: bounds.right, right: bounds.right } : null;
}

// After a focused render: keep the caret in view, then stand the keyboard sink
// on it, so the IME candidate window opens beside the caret.
export function scheduleCaretFollowUp(mathfield: _Mathfield): void {
  if (mathfield.caretScrollPending) return;
  mathfield.caretScrollPending = true;
  queueMicrotask(() => {
    mathfield.caretScrollPending = false;
    if (!isValidMathfield(mathfield)) return;
    keepCaretInScrollingAncestor(mathfield);
    visualTexAnchorKeyboardSink(mathfield);
  });
}

// VisualTeX editing core (imeAnchor): the keyboard sink stands on the caret.
// Also run when a key or a composition arrives: the editor may have moved the
// formula (row resizing) since the last render.
export function visualTexAnchorKeyboardSink(mathfield: _Mathfield): void {
  const anchor = visualTexEditingCore?.imeAnchor(mathfield);
  if (anchor) mathfield.keyboardDelegate.moveTo(anchor.left, anchor.top, anchor.bottom - anchor.top);
}

export function keepCaretInScrollingAncestor(mathfield: _Mathfield): void {
  if (!mathfield.host || !mathfield.field || mathfield.visualTexDragScrolling) return;
  const scroller = horizontalScrollingAncestor(mathfield.host);
  if (!scroller) return;
  const caret = caretHorizontalBounds(mathfield);
  if (!caret) return;
  const view = scroller.getBoundingClientRect();
  const margin = Math.min(CARET_SCROLL_MARGIN, view.width / 4);
  if (caret.right > view.right - margin)
    scroller.scrollLeft += caret.right - (view.right - margin);
  else if (caret.left < view.left + margin)
    scroller.scrollLeft -= view.left + margin - caret.left;
}

export function renderSelection(
  mathfield: _Mathfield,
  interactive?: boolean
): void {
  const field = mathfield.field;

  // In some rare cases, we can get called (via a timeout) when the field
  // is either no longer ready, or not yet ready. Bail.
  if (!field) return;

  // Remove existing selection
  for (const element of field.querySelectorAll(
    '.ML__selection, .ML__contains-highlight'
  ))
    element.remove();

  if (
    !(interactive ?? false) &&
    gFontsState !== 'error' &&
    gFontsState !== 'ready'
  ) {
    // If the fonts are not loaded, or if they are still loading, schedule
    // a re-render of the selection to a bit later. If after waiting a bit
    // the fonts are still not ready,
    // Once the fonts are loaded, the layout may shift due to the glyph metrics
    // being different after font-substitution, which may affect rendering of
    // the selection
    setTimeout(() => {
      if (gFontsState === 'ready') renderSelection(mathfield);
      else setTimeout(() => renderSelection(mathfield), 128);
    }, 32);
    return;
  }

  const model = mathfield.model;

  // Cache the scale factor
  // In some cases we don't need it, so we want to avoid computing it
  // since it can trigger a reflow
  let _scaleFactor: number | undefined;
  const scaleFactor = () => {
    if (_scaleFactor !== undefined) return _scaleFactor;
    // Logic to accommodate mathfield hosted in an isotropically scale-transformed element.
    // Without this, the selection indicator will not be in the right place.

    // 1. Inquire how big the mathfield thinks it is
    const offsetWidth = field.offsetWidth;

    // 2. Get the actual screen width of the box
    const actualWidth = field.getBoundingClientRect().width;
    // 3. Divide the two to get the scale factor
    _scaleFactor = Math.floor(actualWidth) / offsetWidth;
    if (isNaN(_scaleFactor)) _scaleFactor = 1;

    return _scaleFactor;
  };

  if (model.selectionIsCollapsed) {
    //
    // 1.1. Display the popover relative to the location of the caret
    //
    updateSuggestionPopoverPosition(mathfield, { deferred: true });

    //
    // 1.2. Display the 'contains' highlight
    //
    let atom = model.at(model.position);
    while (
      atom &&
      atom.type !== 'prompt' &&
      !(atom.containsCaret && atom.displayContainsHighlight)
    )
      atom = atom.parent!;

    if (atom?.containsCaret && atom.displayContainsHighlight) {
      const s = scaleFactor();
      const bounds = adjustForScrolling(
        mathfield,
        getAtomBounds(mathfield, atom),
        s
      );

      if (bounds) {
        bounds.left /= s;
        bounds.right /= s;
        bounds.top /= s;
        bounds.bottom /= s;

        const element = document.createElement('div');
        element.classList.add('ML__contains-highlight');
        element.style.position = 'absolute';
        element.style.left = `${bounds.left + 1}px`;
        element.style.top = `${Math.ceil(bounds.top)}px`;
        element.style.width = `${Math.ceil(bounds.right - bounds.left)}px`;
        element.style.height = `${Math.ceil(bounds.bottom - bounds.top)}px`;
        field.insertBefore(element, field.childNodes[0]);
      }
    }

    return;
  }

  // A selected placeholder is the caret in an empty slot (drawn by the
  // layout engine), not a range to highlight.
  if (visualTexEditingCore && visualTexEditingCore.caretState(mathfield).kind === 'empty') return;

  //
  // 2. Display the non-collapsed selection
  //
  const s = scaleFactor();
  const rects = unionRects(
    getSelectionBounds(mathfield, { excludeAtomsWithBackground: true })
  );

  for (const x of rects) {
    x.left /= s;
    x.right /= s;
    x.top /= s;
    x.bottom /= s;

    const selectionElement = document.createElement('div');
    selectionElement.classList.add('ML__selection');
    selectionElement.style.position = 'absolute';
    selectionElement.style.left = `${x.left}px`;
    selectionElement.style.top = `${x.top}px`;
    selectionElement.style.width = `${Math.ceil(x.right - x.left)}px`;
    selectionElement.style.height = `${Math.max(1, Math.ceil(x.bottom - x.top - 1))}px`;
    field.insertBefore(selectionElement, field.childNodes[0]);
  }
}

/**
 * Return the rects that are not entirely contained by other rects.
 */
function unionRects(rects: Rect[]): Rect[] {
  let result: Rect[] = [];

  // Remove duplicate rects
  for (const rect of rects) {
    let found = false;
    for (const rect2 of result) {
      if (
        rect.left === rect2.left &&
        rect.right === rect2.right &&
        rect.top === rect2.top &&
        rect.bottom === rect2.bottom
      ) {
        found = true;
        break;
      }
    }
    if (!found) result.push(rect);
  }
  rects = result;
  result = [];
  for (const rect of rects) {
    let count = 0;
    for (const rect2 of rects) {
      if (
        rect.left >= rect2.left &&
        rect.right <= rect2.right &&
        rect.top >= rect2.top &&
        rect.bottom <= rect2.bottom
      ) {
        count += 1;
        if (count > 1) break;
      }
    }
    if (count === 1) result.push(rect);
  }
  return result;
}

/**
 * Re parse the content and rerender.
 *
 * Used when context changes, for example the definition
 * of macros or the `isFunction` global option.
 *
 * @param mathfield
 */
export function reparse(mathfield: _Mathfield | null): void {
  if (!mathfield) return;
  const model = mathfield.model;
  const selection = model.selection;
  const content = Atom.serialize([model.root], {
    expandMacro: false,
    defaultMode: mathfield.options.defaultMode,
  });
  ModeEditor.insert(model, content, {
    insertionMode: 'replaceAll',
    selectionMode: 'after',
    format: 'latex',
    silenceNotifications: true,
    mode: 'math',
  });
  const wasSilent = model.silenceNotifications;
  model.silenceNotifications = true;
  model.selection = selection;
  model.silenceNotifications = wasSilent;
  requestUpdate(mathfield);
}

export function reparseAllMathfields(): void {
  for (const mathfield of document.querySelectorAll('.ML__mathfield'))
    if ('_mathfield' in mathfield) reparse(mathfield._mathfield as _Mathfield);
}
