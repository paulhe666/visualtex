import type { _Model } from './model-private';
import { LatexAtom } from '../atoms/latex';
import { TextAtom } from '../atoms/text';
import { LETTER_AND_DIGITS } from '../latex-commands/definitions-utils';
import type { Offset } from '../public/core-types';
import { getCommandSuggestionRange } from '../editor-mathfield/mode-editor-latex';
import { deleteRange } from './delete';
import { visualTexEditingCore } from '../editor-mathfield/visualtex-hooks';
import { visualTexLeapBoundedOperatorForward } from './commands-move';

/*
 * Calculates the offset of the "next word".
 * This is inspired by the behavior of text editors on macOS, namely:
    blue   yellow
      ^-
         ^-------
 * That is:

 * (1) If starts with an alphanumerical character, find the first alphanumerical
 * character which is followed by a non-alphanumerical character
 *
 * The behavior regarding non-alphanumeric characters is less consistent.
 * Here's the behavior we use:
 *
 *   +=-()_:”     blue
 * ^---------
 *   +=-()_:”     blue
 *      ^---------
 *   +=-()_:”blue
 *      ^--------
 *
 * (2) If starts in whitespace, skip whitespace, then find first non-whitespace*
 *    followed by whitespace
 * (*) Pages actually uses the character class of the first non-whitespace
 * encountered.
 *
 * (3) If starts in a non-whitespace, non alphanumerical character, find the first
 *      whitespace
 *
 */
export function wordBoundaryOffset(
  model: _Model,
  offset: Offset,
  direction: 'forward' | 'backward'
): number {
  if (model.at(offset).mode !== 'text') return offset;
  const dir = direction === 'backward' ? -1 : +1;
  let result;
  if (LETTER_AND_DIGITS.test(model.at(offset).value)) {
    // (1) We start with an alphanumerical character
    let i = offset;
    let match;
    do {
      match =
        model.at(i).mode === 'text' &&
        LETTER_AND_DIGITS.test(model.at(i).value);
      i += dir;
    } while (model.at(i) && match);

    result = model.at(i) ? i - 2 * dir : i - dir;
  } else if (/\s/.test(model.at(offset).value)) {
    // (2) We start with whitespace

    // Skip whitespace
    let i = offset;
    while (model.at(i)?.mode === 'text' && /\s/.test(model.at(i).value))
      i += dir;

    if (!model.at(i)) {
      // We've reached the end
      result = i - dir;
    } else {
      let match = true;
      do {
        match = model.at(i).mode === 'text' && !/\s/.test(model.at(i).value);
        i += dir;
      } while (model.at(i) && match);

      result = model.at(i) ? i - 2 * dir : i - dir;
    }
  } else {
    // (3)
    let i = offset;
    // Skip non-whitespace
    while (model.at(i)?.mode === 'text' && !/\s/.test(model.at(i).value))
      i += dir;

    result = model.at(i) ? i : i - dir;
    let match = true;
    while (model.at(i) && match) {
      match = model.at(i).mode === 'text' && /\s/.test(model.at(i).value);
      if (match) result = i;
      i += dir;
    }

    result = model.at(i) ? i - 2 * dir : i - dir;
  }

  return result - (dir > 0 ? 0 : 1);
}

/**
 * Keyboard navigation with alt/option:
 * Move the insertion point to the next/previous point of interest.
 * A point of interest is an atom of a different type (mbin, mord, etc...)
 * than the current focus.
 * If `extend` is true, the selection will be extended. Otherwise, it is
 * collapsed, then moved.
 * If `delete` is true, the skipped range is removed.
 * @todo array
 */
export function skip(
  model: _Model,
  direction: 'forward' | 'backward',
  options?: { extend?: boolean; delete?: boolean }
): boolean {
  const previousPosition = model.position;

  if (!(options?.extend ?? false)) model.collapseSelection(direction);

  let atom = model.at(model.position);
  if (direction === 'forward') {
    if (atom.type === 'subsup') {
      atom = atom.rightSibling;
      if (!atom) atom = model.at(model.position + 1);
    } else atom = model.at(model.position + 1);
  }

  if (!atom) {
    model.announce('plonk');
    return false;
  }

  let offset = model.offsetOf(atom);

  if (atom instanceof TextAtom) {
    //
    // We're in a text zone, skip word by word
    //
    offset = wordBoundaryOffset(model, offset, direction);
  } else if (atom instanceof LatexAtom) {
    //
    // We're in a LaTeX mode zone, skip suggestion
    //
    if (atom.isSuggestion) {
      // Since suggestions are always at the end, this must be forward
      console.assert(direction === 'forward');
      while (atom && atom instanceof LatexAtom) {
        atom.isSuggestion = false;
        offset = model.offsetOf(atom);
        atom = atom.rightSibling;
      }
    } else if (direction === 'forward') {
      atom = atom.rightSibling;
      if (!atom || !(atom instanceof LatexAtom)) {
        // At the end of the command
        model.announce('plonk');
        return false;
      }

      while (
        atom &&
        atom instanceof LatexAtom &&
        /[a-zA-Z\*]/.test(atom.value)
      ) {
        offset = model.offsetOf(atom);
        atom = atom.rightSibling;
      }
    } else {
      atom = atom.leftSibling;
      if (!atom || !(atom instanceof LatexAtom)) {
        // At the start of the command
        model.announce('plonk');
        return false;
      }

      while (
        atom &&
        atom instanceof LatexAtom &&
        /[a-zA-Z\*]/.test(atom.value)
      ) {
        offset = model.offsetOf(atom);
        atom = atom.leftSibling;
      }
    }
  } else if (direction === 'forward' && atom.type === 'mopen') {
    //
    // Right before a 'mopen', skip to the corresponding balanced fence
    //
    let level = 0;
    do {
      if (atom.type === 'mopen') level += 1;
      else if (atom.type === 'mclose') level -= 1;

      atom = atom.rightSibling;
    } while (!atom.isLastSibling && level !== 0);

    offset = model.offsetOf(atom.leftSibling);
  } else if (direction === 'backward' && atom.type === 'mclose') {
    //
    // Right after a 'mclose', skip to the corresponding balanced fence
    //
    let level = 0;
    do {
      if (atom.type === 'mopen') level += 1;
      else if (atom.type === 'mclose') level -= 1;

      atom = atom.leftSibling;
    } while (!atom.isFirstSibling && level !== 0);

    offset = model.offsetOf(atom);
  } else if (direction === 'backward') {
    //
    // We're in a regular math zone (not before/after a fence)
    //
    if (atom.type === 'first') {
      while (offset > 0 && atom.type === 'first') {
        offset -= 1;
        atom = model.at(offset);
      }
    } else {
      const type = atom.type;
      if (atom.type === 'subsup') {
        // If we're after a 'subsup', skip to its left sibling
        // (the base of the super/subscript)
        offset = model.offsetOf(model.at(offset).leftSibling);
      }

      offset -= 1;
      let nextType = model.at(offset)?.type;
      // If (nextType === 'subsup') {
      //     offset = model.offsetOf(model.at(offset).leftSibling);
      // }
      while (offset >= 0 && nextType === type) {
        if (model.at(offset)?.type === 'subsup')
          offset = model.offsetOf(model.at(offset).leftSibling);
        else offset -= 1;

        nextType = model.at(offset).type;
      }
    }
  } else {
    const { type: type } = atom;
    // If (atom.type === 'subsup') {
    //     offset = model.offsetOf(model.at(offset).rightSibling);
    // }
    let nextType = model.at(offset)?.type;
    const { lastOffset } = model;
    while (
      offset <= lastOffset &&
      (nextType === type || nextType === 'subsup')
    ) {
      while (model.at(offset).rightSibling?.type === 'subsup')
        offset = model.offsetOf(model.at(offset).rightSibling);

      offset += 1;
      nextType = model.at(offset)?.type;
    }

    offset -= 1;
  }

  if (options?.extend ?? false) {
    // VisualTeX editing core (selection.ts): one selection rule
    if (!visualTexEditingCore.extendTo(model, offset)) return false;
  } else {
    if (offset === model.position) {
      model.announce('plonk');
      return false;
    }

    if (options?.delete ?? false) {
      if (direction === 'forward')
        deleteRange(model, [previousPosition, offset], 'deleteWordForward');
      else {
        deleteRange(model, [previousPosition, offset], 'deleteWordBackward');
        model.position = offset;
      }
    } else {
      model.position = offset;
      model.announce('move', previousPosition);
    }
  }

  model.mathfield.stopCoalescingUndo();
  return true;
}

/**
 * Handle keyboard navigation (arrow keys)
 */
export function move(
  model: _Model,
  direction: 'forward' | 'backward',
  options?: { extend: boolean }
): boolean {
  options = options ?? { extend: false };
  // VisualTeX: from a big operator's last limit, a step forward leaves it
  if (direction === 'forward' && !options?.extend && visualTexLeapBoundedOperatorForward(model, false)) return true;

  model.mathfield.styleBias = direction === 'backward' ? 'right' : 'left';

  if (direction !== 'forward') {
    const [from, to] = getCommandSuggestionRange(model);
    if (from !== undefined && to !== undefined) model.deleteAtoms([from, to]);
  }

  if (model.selectionIsPlaceholder) {
    model.collapseSelection(direction);
    const result = move(model, direction);
    model.mathfield.stopCoalescingUndo();
    return result;
  }

  let pos = model.position;
  const previousPosition = pos;
  if (model.collapseSelection(direction)) {
    pos = model.position;
    if (!isValidPosition(model, pos))
      pos = nextValidPosition(model, pos, direction);
  } else pos = nextValidPosition(model, pos, direction);

  if (pos < 0 || pos > model.lastOffset) {
    // We're going out of bounds
    let success = true; // True => perform default handling
    if (!model.silenceNotifications) {
      success =
        model.mathfield.host?.dispatchEvent(
          new CustomEvent('move-out', {
            detail: { direction },
            cancelable: true,
            bubbles: true,
            composed: true,
          })
        ) ?? true;
    }
    if (success) model.announce('plonk');
    return success;
  }

  model.setPositionHandlingPlaceholder(pos);
  model.mathfield.stopCoalescingUndo();
  model.announce('move', previousPosition);

  return true;
}

function nextValidPosition(
  model: _Model,
  pos: number,
  direction: 'forward' | 'backward'
): number {
  pos = pos + (direction === 'forward' ? +1 : -1);

  if (pos < 0 || pos > model.lastOffset) return pos;

  if (!isValidPosition(model, pos))
    return nextValidPosition(model, pos, direction);

  return pos;
}

// VisualTeX editing core (stops.ts): the one rule for where the caret may stand.
function isValidPosition(model: _Model, pos: number): boolean {
  return visualTexEditingCore ? visualTexEditingCore.isAllowedStop(model, pos) : true;
}



