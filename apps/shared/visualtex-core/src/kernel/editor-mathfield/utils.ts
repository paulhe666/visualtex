import { Atom } from '../core/atom-class';
import type { ElementInfo, Offset, Range } from '../public/core-types';
import { _Mathfield } from './mathfield-private';
import { visualTexEditingCore } from './visualtex-hooks';

export type Rect = {
  top: number;
  bottom: number;
  left: number;
  right: number;
};

/**
 * Checks if the argument is a valid Mathfield.
 * After a Mathfield has been destroyed (for example by calling `dispose()`
 * the Mathfield is no longer valid. However, there may be some pending
 * operations invoked via requestAnimationFrame() for example, that would
 * need to ensure the mathfield is still valid by the time they're executed.
 */
export function isValidMathfield(mf: _Mathfield): boolean {
  return mf.element?.mathfield === mf;
}

// VisualTeX editing core: the caret point and atom bounds come from the
// layout geometry, not from measuring the drawn formula.

/**
 * Return the (x,y) client coordinates of the caret in viewport coordinates
 */
export function getCaretPoint(
  mathfield: _Mathfield
): { x: number; y: number; height: number } | null {
  return visualTexEditingCore ? visualTexEditingCore.caretPoint(mathfield) : null;
}

function branchId(atom: Atom): string {
  if (!atom.parent) return 'root';
  let result = atom.parent.id ?? '';
  result +=
    typeof atom.parentBranch === 'string'
      ? '-' + atom.parentBranch
      : `-${atom.parentBranch![0]}/${atom.parentBranch![1]}`;
  return result;
}

export function adjustForScrolling(
  mathfield: _Mathfield,
  rect: Rect | null,
  scaleFactor: number
): Rect | null {
  if (!rect) return null;
  const fieldRect = mathfield.field!.getBoundingClientRect();
  const w = rect.right - rect.left;
  const h = rect.bottom - rect.top;
  const left = Math.ceil(
    rect.left - fieldRect.left + mathfield.field.scrollLeft * scaleFactor
  );

  const top = Math.ceil(rect.top - fieldRect.top);
  return { left, right: left + w, top, bottom: top + h };
}

export function getAtomBounds(mathfield: _Mathfield, atom: Atom): Rect | null {
  if (!atom || !visualTexEditingCore) return null;
  return visualTexEditingCore.atomBounds(mathfield, atom);
}

/*
 * Return an array of bounds for the specified range, at most
 * one rect per branch.
 */
function getRangeBounds(
  mathfield: _Mathfield,
  range: Range,
  options?: { excludeAtomsWithBackground?: boolean }
): Rect[] {
  // The key of the map is a 'branchId', i.e. "atom id + branch"
  const rects = new Map<string, Rect>();

  // Logic to accommodate mathfield hosted in an isotropically
  // scale-transformed element.
  // Without this, the selection indicator will not be in the right place.

  // 1. Inquire how big the mathfield thinks it is
  const field = mathfield.field;
  const offsetWidth = field.offsetWidth;

  // 2. Get the actual screen width of the box
  const actualWidth = Math.floor(field.getBoundingClientRect().width);

  // 3. Divide the two to get the scale factor
  let scaleFactor = actualWidth / offsetWidth;
  scaleFactor = isNaN(scaleFactor) ? 1 : scaleFactor;

  const selectedAtoms = mathfield.model.getAtoms(range, { includeChildren: true });
  const selectedSet = new Set(selectedAtoms);
  for (const atom of selectedAtoms) {
    // Sentinels have the surrounding font's ascent, not visible ink. Including
    // them creates tall phantom highlights in scripts and matrix cells.
    if (atom.type === 'first') continue;
    if (options?.excludeAtomsWithBackground && atom.style.backgroundColor)
      continue;

    const bounds = adjustForScrolling(
      mathfield,
      getAtomBounds(mathfield, atom),
      scaleFactor
    );
    if (bounds) {
      // Draw a fully selected structure once, including all protruding scripts.
      // Partial branches/cells keep separate rectangles and cannot tint siblings.
      let owner = atom;
      while (owner.parent && selectedSet.has(owner.parent)) owner = owner.parent;
      const id = branchId(owner);
      if (rects.has(id)) {
        const r = rects.get(id)!;
        rects.set(id, {
          left: Math.min(r.left, bounds.left),
          right: Math.max(r.right, bounds.right),
          top: Math.min(r.top, bounds.top),
          bottom: Math.max(r.bottom, bounds.bottom),
        });
      } else rects.set(id, bounds);
    }
  }

  return [...rects.values()];
}

export function getSelectionBounds(
  mathfield: _Mathfield,
  options?: { excludeAtomsWithBackground?: boolean }
): Rect[] {
  return mathfield.model.selection.ranges.reduce(
    (acc: Rect[], x) => acc.concat(...getRangeBounds(mathfield, x, options)),
    []
  );
}

export function getElementInfo(
  mf: _Mathfield | undefined | null,
  offset: Offset
): ElementInfo | undefined {
  if (!mf) return undefined;

  const atom = mf.model.at(offset);
  if (!atom) return undefined;

  const result: ElementInfo = {};

  const bounds = getAtomBounds(mf, atom);
  if (bounds) {
    result.bounds = new DOMRect(
      bounds.left,
      bounds.top,
      bounds.right - bounds.left,
      bounds.bottom - bounds.top
    );
  }

  result.depth = atom.treeDepth - 2;

  result.style = atom.style;

  // Look for some 'htmlData' in the atom or its ancestors
  let a: undefined | Atom = atom;
  while (a) {
    if (a.command === '\\htmlData' && a.args && typeof a.args[0] === 'string') {
      const entries = a.args[0].split(',');
      for (const entry of entries) {
        const matched = entry.match(/([^=]+)=(.+$)/);
        if (matched) {
          const key = matched[1].trim().replace(/ /g, '-');
          if (key) {
            if (!result.data) result.data = {};
            result.data[key] = matched[2];
          }
        } else {
          const key = entry.trim().replace(/ /g, '-');
          if (key) {
            if (!result.data) result.data = {};
            result.data[key] = undefined;
          }
        }
      }
    }

    if (a.command === '\\htmlId' || a.command === '\\cssId') {
      if (!result.id && a.args && typeof a.args[0] === 'string')
        result.id = a.args[0];
    }
    a = a.parent;
  }

  if (atom.mode === 'math' || atom.mode === 'text')
    result.latex = Atom.serialize([atom], { defaultMode: 'math' });

  return result;
}

export function getHref(mf: _Mathfield, offset: Offset): string {
  let a: Atom | undefined = mf.model.at(offset);
  while (a) {
    if (a.command === '\\href') {
      const url = a.args[0];
      if (typeof url === 'string') return url;
    }

    a = a.parent;
  }
  return '';
}
