import { isBrowser } from '../ui/utils/capabilities';

import { Atom } from '../core/atom';
import { SubsupAtom } from '../atoms/subsup';
import { register } from '../editor/commands';

import { move, skip } from './commands';
import type { _Model } from './model-private';
import type { Branch, BranchName } from '../core/types';
import { isArray } from '../common/types';
import { visualTexEditingCore } from '../editor-mathfield/visualtex-hooks';
import { _MathEnvironment } from '../core/math-environment';
import { ModeEditor } from '../editor-mathfield/mode-editor';
import { ArrayAtom } from '../atoms/array';
import { isCellBranch } from '../core/atom-class';
import { addRowAfter, addColumnAfter } from './array';
import { range } from './selection-utils';

export function moveAfterParent(model: _Model): boolean {
  const previousPosition = model.position;
  const parent = model.at(previousPosition).parent;
  // Do nothing if at the root.
  if (!parent?.parent) {
    model.announce('plonk');
    return false;
  }

  model.position = model.offsetOf(parent);
  model.mathfield.stopCoalescingUndo();
  model.announce('move', previousPosition);
  return true;
}

function superscriptDepth(model: _Model): number {
  let result = 0;
  let atom: Atom | undefined = model.at(model.position);
  let wasSuperscript = false;
  while (atom) {
    if (!atom.hasEmptyBranch('superscript')) {
      result += 1;
      wasSuperscript = true;
    } else if (!atom.hasEmptyBranch('subscript')) wasSuperscript = false;

    atom = atom.parent;
  }

  return wasSuperscript ? result : 0;
}

function subscriptDepth(model: _Model): number {
  let result = 0;
  let atom: Atom | undefined = model.at(model.position);
  let wasSubscript = false;
  while (atom) {
    if (!atom.hasEmptyBranch('subscript')) {
      result += 1;
      wasSubscript = true;
    } else if (!atom.hasEmptyBranch('superscript')) wasSubscript = false;

    atom = atom.parent;
  }

  return wasSubscript ? result : 0;
}

/**
 * Switch the cursor to the superscript and select it. If there is no subscript
 * yet, create one.
 */
function moveToSuperscript(model: _Model): boolean {
  model.collapseSelection();
  if (superscriptDepth(model) >= (model.mathfield.options.scriptDepth as [number, number])[1]) {
    model.announce('plonk');
    return false;
  }

  let target = model.at(model.position);

  if (target.subsupPlacement === undefined) {
    // This atom can't have a superscript/subscript:
    // add an adjacent `subsup` atom instead.
    let subsup = target.rightSibling;
    if (subsup?.type !== 'subsup') {
      subsup = new SubsupAtom({ style: target.style });
      target.parent!.addChildAfter(subsup, target);
    }

    target = subsup!;
  }

  // Ensure there is a superscript branch
  target.createBranch('superscript');
  model.setSelection(
    model.getSiblingsRange(model.offsetOf(target.superscript![0]))
  );

  return true;
}

/**
 * Switch the cursor to the subscript and select it. If there is no subscript
 * yet, create one.
 */
function moveToSubscript(model: _Model): boolean {
  model.collapseSelection();
  if (subscriptDepth(model) >= (model.mathfield.options.scriptDepth as [number, number])[0]) {
    model.announce('plonk');
    return false;
  }

  let target = model.at(model.position);

  if (target.subsupPlacement === undefined) {
    // This atom can't have a superscript/subscript:
    // add an adjacent `subsup` atom instead.
    let subsup = target.rightSibling;
    if (subsup?.type !== 'subsup') {
      subsup = new SubsupAtom({ style: target.style });
      target.parent!.addChildAfter(subsup, target);
    }

    target = subsup!;
  }

  // Ensure there is a subscript branch
  target.createBranch('subscript');
  model.setSelection(
    model.getSiblingsRange(model.offsetOf(target.subscript![0]))
  );
  return true;
}

/**
 * Return an array of tabbable elements, approximately in the order a browser
 * would (the browsers are inconsistent), which is first by accounting
 * for non-null tabIndex, then null tabIndex, then document order of focusable
 * elements.
 */
function getTabbableElements(): HTMLElement[] {
  function tabbable(element: HTMLElement) {
    const regularTabbables: HTMLElement[] = [];
    const orderedTabbables: {
      documentOrder: number;
      tabIndex: number;
      node: HTMLElement;
    }[] = [];

    const candidates = [
      ...element.querySelectorAll<HTMLElement>(`input, select, textarea, a[href], button,
        [tabindex], audio[controls], video[controls],
        [contenteditable]:not([contenteditable="false"]), details>summary`),
    ].filter(isNodeMatchingSelectorTabbable);
    candidates.forEach((candidate, i) => {
      const candidateTabindex = getTabindex(candidate);
      if (candidateTabindex === 0) regularTabbables.push(candidate);
      else {
        orderedTabbables.push({
          documentOrder: i,
          tabIndex: candidateTabindex,
          node: candidate,
        });
      }
    });

    return orderedTabbables
      .sort((a, b) =>
        a.tabIndex === b.tabIndex
          ? a.documentOrder - b.documentOrder
          : a.tabIndex - b.tabIndex
      )
      .map((a) => a.node)
      .concat(regularTabbables);
  }

  function isNodeMatchingSelectorTabbable(element: HTMLElement): boolean {
    if (
      !isNodeMatchingSelectorFocusable(element) ||
      isNonTabbableRadio(element) ||
      getTabindex(element) < 0
    )
      return false;

    return true;
  }

  function isNodeMatchingSelectorFocusable(node: HTMLElement): boolean {
    if (
      (node as HTMLInputElement).disabled ||
      ((node as HTMLInputElement).type === 'hidden' &&
        node.tagName.toUpperCase() === 'INPUT') ||
      isHidden(node)
    )
      return false;

    return true;
  }

  function getTabindex(node: HTMLElement): number {
    const tabindexAttr = Number.parseInt(
      node.getAttribute('tabindex') ?? 'NaN',
      10
    );

    if (!Number.isNaN(tabindexAttr)) return tabindexAttr;

    // Browsers do not return `tabIndex` correctly for contentEditable nodes;
    // so if they don't have a tabindex attribute specifically set, assume it's 0.
    if (node.contentEditable === 'true') return 0;

    // In Chrome, <audio controls/> and <video controls/> elements get a default
    //  `tabIndex` of -1 when the 'tabindex' attribute isn't specified in the DOM,
    //  yet they are still part of the regular tab order; in FF, they get a default
    //  `tabIndex` of 0; since Chrome still puts those elements in the regular tab
    //  order, consider their tab index to be 0
    if (
      (node.nodeName === 'AUDIO' || node.nodeName === 'VIDEO') &&
      node.getAttribute('tabindex') === null
    )
      return 0;

    return node.tabIndex;
  }

  function isNonTabbableRadio(node: HTMLElement): boolean {
    return (
      node.tagName.toUpperCase() === 'INPUT' &&
      (node as HTMLInputElement).type === 'radio' &&
      !isTabbableRadio(node as HTMLInputElement)
    );
  }

  function getCheckedRadio(
    nodes: NodeListOf<Element>,
    form: HTMLFormElement | null
  ): HTMLInputElement | null {
    for (const node of nodes as NodeListOf<HTMLInputElement>)
      if (node.checked && node.form === form) return node;

    return null;
  }

  function isTabbableRadio(node: HTMLInputElement): boolean {
    if (!node.name) return true;

    const radioScope = node.form ?? node.ownerDocument;
    const radioSet = radioScope.querySelectorAll(
      'input[type="radio"][name="' + node.name + '"]'
    );
    const checked = getCheckedRadio(radioSet, node.form);
    return !checked || checked === node;
  }

  function isHidden(element: HTMLElement) {
    if (
      !isBrowser() ||
      element === document.activeElement ||
      element.contains(document.activeElement)
    )
      return false;

    if (getComputedStyle(element).visibility === 'hidden') return true;

    // Note that browsers generally don't consider the bounding rect
    // as a criteria to determine if an item is focusable, but we want
    // to exclude the invisible textareas used to capture keyoard input.
    const bounds = element.getBoundingClientRect();
    if (bounds.width === 0 || bounds.height === 0) return true;

    while (element) {
      if (getComputedStyle(element).display === 'none') return true;
      element = element.parentElement!;
    }

    return false;
  }

  if (!isBrowser()) return [];
  return tabbable(document.body);
}

// Select all the children of an atom, or a branch
function select(
  model: _Model,
  target: Atom | readonly Atom[],
  direction: 'backward' | 'forward' = 'forward'
): boolean {
  const previousPosition = model.position;
  if (isArray(target)) {
    // The target is a branch. Select all the atoms in the branch
    const first = model.offsetOf(target[0]);
    const last = model.offsetOf(target[target.length - 1]);
    if (direction === 'forward') model.setSelection(first, last);
    else model.setSelection(last, first);
    model.announce('move', previousPosition);
    model.mathfield.stopCoalescingUndo();
    return true;
  }
  if (direction === 'forward')
    return select(model, [target.leftSibling, target]);
  return select(model, [target, target.leftSibling]);
}

export function leapTo(model: _Model, target: Atom | number): boolean {
  const previousPosition = model.position;

  if (typeof target === 'number') target = model.at(target);
  // Set the selection to the next leap target
  if (target.type === 'prompt') {
    model.setSelection(
      model.offsetOf(target.firstChild),
      model.offsetOf(target.lastChild)
    );
  } else if (target.type === 'placeholder') {
    // VisualTeX editing core: the placeholder's empty slot
    visualTexEditingCore.goToStop(model, model.offsetOf(target));
  } else {
    model.position = model.offsetOf(target);
  }
  model.announce('move', previousPosition);
  model.mathfield.stopCoalescingUndo();
  return true;
}

export function visualTexIsBoundedOperatorParent(parent: Atom | undefined): boolean {
  if (!parent) return false;
  if (parent.type === "mop" || parent.type === "operator" || parent.type === "extensible-symbol") return true;
  if (parent.type !== "subsup") return false;
  const base = parent.leftSibling;
  return Boolean(base && (base.type === "error" || base.type === "mop" || base.type === "operator" || base.type === "extensible-symbol"));
}

export function visualTexBoundedOperatorBranchPlaceholders(
  _model: _Model,
  parent: Atom,
  branch: Branch
): Atom[] {
  if (!visualTexIsBoundedOperatorParent(parent)) return [];
  const branchAtoms = parent.branch(branch);
  if (!branchAtoms) return [];
  return branchAtoms.flatMap((atom) => [atom, ...atom.children]).filter((atom) => atom.type === "placeholder" || atom.type === "prompt");
}

export function visualTexLeapBoundedOperatorForward(
  model: _Model,
  allowInternalPlaceholderJump = true
): boolean {
  const selected = model.at(Math.max(model.position, model.anchor));
  if (!selected) return false;
  const parent = selected.parent;
  const branch = selected.parentBranch;
  if (!parent || (branch !== "subscript" && branch !== "superscript")) return false;
  if (!visualTexIsBoundedOperatorParent(parent)) return false;
  const lowerBranch = parent.branch("subscript");
  const upperBranch = parent.branch("superscript");
  if (!lowerBranch || !upperBranch) return false;
  const lower = visualTexBoundedOperatorBranchPlaceholders(model, parent, "subscript");
  const upper = visualTexBoundedOperatorBranchPlaceholders(model, parent, "superscript");
  const currentOffset = model.offsetOf(selected);
  const currentList = branch === "subscript" ? lower : upper;
  const nextInBranch = currentList.find((atom) => model.offsetOf(atom) > currentOffset);
  if (allowInternalPlaceholderJump && nextInBranch) return leapTo(model, nextInBranch);
  if (!allowInternalPlaceholderJump && !selected.isLastSibling) return false;
  if (branch === "subscript") return upper[0] ? leapTo(model, upper[0]) : false;
  const operatorOffset = model.offsetOf(parent);
  const target = leapTarget(model, operatorOffset + 1, "forward");
  if (target && model.offsetOf(target) > operatorOffset) return leapTo(model, target);
  return leapTo(model, parent);
}

/**
 * Move to the next/previous leap target: placeholder, editable prompt or
 * empty child list.
 * @return `false` if no placeholder found and did not move
 */
function leap(model: _Model, dir: 'forward' | 'backward'): boolean {
  // VisualTeX: from a big operator's lower limit to its upper one, then out
  if (dir === 'forward' && visualTexLeapBoundedOperatorForward(model)) return true;
  const dist = dir === 'forward' ? 1 : -1;

  // If we're already at a placeholder, move by one more (the placeholder
  // is right after the insertion point)
  if (model.at(model.anchor).type === 'placeholder') move(model, dir);

  let origin: number;

  // If we're in a prompt, start looking after/before the prompt
  const parentPrompt = model.at(model.anchor).parentPrompt;
  if (parentPrompt) {
    if (dir === 'forward') origin = model.offsetOf(parentPrompt) + 1;
    else origin = model.offsetOf(parentPrompt.leftSibling);
  } else origin = Math.max(model.position + dist, 0);

  const target = leapTarget(model, origin, dir);

  // If no leap target was found, call handler or move to the next focusable
  // element in the document
  if (
    !target ||
    (dir === 'forward' && model.offsetOf(target) < origin) ||
    (dir === 'backward' && model.offsetOf(target) > origin)
  ) {
    const success =
      model.mathfield.host?.dispatchEvent(
        new CustomEvent('move-out', {
          detail: { direction: dir },
          cancelable: true,
          bubbles: true,
          composed: true,
        })
      ) ?? true;

    if (!success) {
      model.announce('plonk');
      return false;
    }

    const tabbable = getTabbableElements();

    // If there are no other elements to focus, plonk.
    if (!document.activeElement || tabbable.length <= 1) {
      model.announce('plonk');
      return false;
    }

    //
    // Focus on next/previous tabbable element
    //
    let index = tabbable.indexOf(document.activeElement as HTMLElement) + dist;
    if (index < 0) index = tabbable.length - 1;
    if (index >= tabbable.length) index = 0;

    tabbable[index].focus();

    model.mathfield.stopCoalescingUndo();
    return true;
  }

  // Set the selection to the next leap target
  leapTo(model, target);

  return true;
}

// Candidate leap targets are atoms of type 'placeholder' or
// 'prompt' or empty children list (except for the root:
// if the root is empty, it is not a valid leap target)

export function leapTarget(
  model: _Model,
  origin = 0,
  dir: 'forward' | 'backward' = 'forward'
): Atom | undefined {
  return model.findAtom(
    (atom) =>
      atom.type === 'placeholder' ||
      atom.type === 'prompt' ||
      (!model.mathfield.readOnly &&
        atom.treeDepth > 2 &&
        atom.isFirstSibling &&
        atom.isLastSibling),
    origin,
    dir
  );
}

/**
 * If cursor is currently in:
 * - superscript: move to subscript, creating it if necessary
 * - subscript: move to superscript, creating it if necessary
 * - numerator: move to denominator
 * - denominator: move to numerator
 * - otherwise: move to superscript
 */
register(
  {
    moveToOpposite: (model: _Model): boolean => {
      const OPPOSITE_RELATIONS: Partial<Record<BranchName, BranchName>> = {
        superscript: 'subscript',
        subscript: 'superscript',
        above: 'below',
        below: 'above',
      };
      const cursor = model.at(model.position);
      const { parent } = cursor;
      if (!parent) {
        model.announce('plonk');
        return false;
      }

      const relation = cursor.parentBranch;
      let oppositeRelation: BranchName | undefined;
      if (typeof relation === 'string')
        oppositeRelation = OPPOSITE_RELATIONS[relation];

      if (!oppositeRelation) {
        const result = cursor.subsupPlacement
          ? moveToSubscript(model)
          : moveToSuperscript(model);

        model.mathfield.stopCoalescingUndo();
        return result;
      }

      if (!parent.branch(oppositeRelation)) {
        // Don't have children of the opposite relation yet
        // Add them
        parent.createBranch(oppositeRelation);
      }

      const result = model.setSelection(
        model.getBranchRange(model.offsetOf(parent), oppositeRelation)
      );
      model.mathfield.stopCoalescingUndo();
      return result;
    },
    moveBeforeParent: (model: _Model): boolean => {
      const { parent } = model.at(model.position);
      if (!parent) {
        model.announce('plonk');
        return false;
      }

      model.position = model.offsetOf(parent);
      model.mathfield.stopCoalescingUndo();
      return true;
    },
    moveAfterParent: (model: _Model): boolean => moveAfterParent(model),

    // VisualTeX editing core (navigation.ts) owns caret movement.
    moveToNextChar: (model: _Model): boolean => visualTexEditingCore.moveHorizontal(model, 1),
    moveToPreviousChar: (model: _Model): boolean => visualTexEditingCore.moveHorizontal(model, -1),
    moveUp: (model: _Model): boolean => visualTexEditingCore.vertical(model, -1, _MathEnvironment.fractionNavigationOrder),
    moveDown: (model: _Model): boolean => visualTexEditingCore.vertical(model, 1, _MathEnvironment.fractionNavigationOrder),
    visualTexSpace: (model: _Model): boolean => visualTexEditingCore.space(model, {
      insertVisibleSpace: () => {
        ModeEditor.insert(model, '\\ ', { format: 'latex', mode: 'math', selectionMode: 'after' });
        return true;
      },
      insertTextSpace: () => {
        ModeEditor.insert(model, ' ', { mode: 'text' });
        return true;
      },
    }),
    visualTexCollapseToEnd: (model: _Model): boolean => visualTexEditingCore.collapseToEnd(model),
    visualTexRestoreCaret: (model: _Model, snapshot): boolean => visualTexEditingCore.restoreCaret(model, snapshot),
    visualTexHome: (model: _Model): boolean => visualTexEditingCore.moveToSlotEdge(model, -1, true),
    visualTexEnd: (model: _Model): boolean => visualTexEditingCore.moveToSlotEdge(model, 1, true),
    moveToNextWord: (model: _Model): boolean => skip(model, 'forward'),
    moveToPreviousWord: (model: _Model): boolean => skip(model, 'backward'),
    moveToGroupStart: (model: _Model): boolean => visualTexEditingCore.moveToSlotEdge(model, -1, false),
    moveToGroupEnd: (model: _Model): boolean => visualTexEditingCore.moveToSlotEdge(model, 1, false),
    moveToNextGroup: (model: _Model): boolean => {
      if (visualTexLeapBoundedOperatorForward(model)) return true;
      //
      // 1/ If at the end of the matfield, leap to next tabbable element
      //
      if (
        model.position === model.lastOffset &&
        model.anchor === model.lastOffset
      )
        return leap(model, 'forward');

      //
      // 2/ If in text zone, move to first non-text atom
      //
      const atom = model.at(model.position);
      const mode = atom.mode;
      if (mode === 'text') {
        if (model.selectionIsCollapsed) {
          // 2.1/ Select entire zone
          let first = atom;
          while (first?.mode === 'text') first = first.leftSibling;
          let last = atom;
          while (last.rightSibling?.mode === 'text') last = last.rightSibling;
          if (first && last) return select(model, [first, last]);
        }
        if (atom.rightSibling.mode === 'text') {
          let next = atom;
          while (next?.mode === 'text') next = next.rightSibling;
          // Leap to after text zone
          if (next) {
            leapTo(model, next.leftSibling ?? next);
            model.mathfield.switchMode('math');
            return true;
          }
          // If text zone is last zone, move to end of mathfield
          return leapTo(model, model.lastOffset);
        }
      }

      //
      // 3/ Find a placeholder, a prompt or empty group.
      // They have priority over other options below
      //

      // If we're in a prompt, start looking after the prompt
      const parentPrompt = model.at(model.anchor).parentPrompt;
      const origin = parentPrompt
        ? model.offsetOf(parentPrompt) + 1
        : Math.max(model.position + 1, 0);

      const target = leapTarget(model, origin, 'forward');
      if (target && model.offsetOf(target) < origin)
        return leap(model, 'forward');

      if (target) return leapTo(model, target);

      //
      // ?/ @todo In LaTeX mode, do something?
      //

      //
      // 4/ In math zone, move to sibling branch, or out of parent
      //
      //
      // 4.1/ Find the next eligible sibling group
      //
      const sibling = findSibling(
        model,
        atom,
        (x) => x.type === 'leftright' || x.type === 'text',
        'forward'
      );

      if (sibling) {
        // If found a text atom, select the entire zone
        if (sibling.mode === 'text') {
          let last = sibling;
          while (last?.mode === 'text') last = last.rightSibling;
          return select(model, [
            sibling.leftSibling ?? sibling,
            last.leftSibling ?? last,
          ]);
        }
        return select(model, sibling);
      }

      const parent = atom.parent;
      if (parent) {
        // 4.2/ Select parent if leftright or surd without index
        if (parent.type === 'leftright' || parent.type === 'surd')
          return select(model, parent);

        // 4.3/ If in an atom with a supsub, try subsup
        if (atom.parentBranch === 'subscript' && parent.superscript)
          return select(model, parent.superscript);

        // 4.4/ If an above branch, try below
        if (atom.parentBranch === 'above' && parent.below)
          return select(model, parent.below);

        // 4.5/ If in a branch (with no "sister" branch), move after parent
        if (
          atom.parentBranch === 'superscript' ||
          atom.parentBranch === 'subscript'
        )
          return leapTo(model, parent);
        if (atom.parentBranch === 'above' || atom.parentBranch === 'below')
          return select(model, parent);
      }

      // 4.6/ No eligible group found, move to end of mathfield
      return leapTo(model, model.lastOffset);
    },
    moveToPreviousGroup: (model: _Model): boolean => {
      //
      // 1/ If at the start of the matfield, leap to previous tabbable element
      //
      if (model.position === 0 && model.anchor === 0)
        return leap(model, 'backward');

      //
      // 2/ If in text zone, move to first text atom
      //

      let atom = model.at(model.position);
      const mode = atom.mode;
      if (mode === 'text') {
        if (model.selectionIsCollapsed) {
          // 2.1/ Select entire zone
          let first = atom;
          while (first?.mode === 'text') first = first.leftSibling;
          let last = atom;
          while (last.rightSibling?.mode === 'text') last = last.rightSibling;
          if (first && last) return select(model, [first, last]);
        }
        while (atom?.mode === 'text') atom = atom.leftSibling;
        if (atom) return leapTo(model, atom);
        // If text zone is first zone, move to start of mathfield
        return leapTo(model, 0);
      }

      //
      // 3/ Find a placeholder, a prompt or empty group.
      //
      // If we're in a prompt, start looking before the prompt
      const parentPrompt = model.at(model.anchor).parentPrompt;
      const origin = parentPrompt
        ? model.offsetOf(parentPrompt.leftSibling)
        : Math.max(model.position - 1, 0);

      const target = leapTarget(model, origin, 'backward');
      if (target && model.offsetOf(target) > origin)
        return leap(model, 'backward');

      if (target) return leapTo(model, target);

      //
      // 4/ In math zone, move to sibling branch, or out of parent
      //
      if (mode === 'math') {
        //
        // 4.1/ Find the previous eligible sibling group
        //
        const sibling = findSibling(
          model,
          atom,
          (x) => x.type === 'leftright' || x.type === 'text',
          'backward'
        );

        if (sibling) {
          // If found a text atom, select the entire zone
          if (sibling.mode === 'text') {
            let first = sibling;
            while (first?.mode === 'text') first = first.leftSibling;
            return select(model, [sibling, first]);
          }
          return select(model, sibling);
        }

        const parent = atom.parent;
        if (parent) {
          // 4.2/ Select parent if leftright or surd without index
          if (parent.type === 'leftright' || parent.type === 'surd')
            return select(model, parent);

          // 4.3/ If in an atom with a supsub, try subsup
          if (atom.parentBranch === 'superscript' && parent.subscript)
            return select(model, parent.subscript);

          // 4.4/ If in a below branch, try above
          if (atom.parentBranch === 'below' && parent.above)
            return select(model, parent.above);

          // 4.5/ If in a branch (with no "sister" branch), move after parent
          if (
            atom.parentBranch === 'superscript' ||
            atom.parentBranch === 'subscript'
          )
            return leapTo(model, parent);
          if (atom.parentBranch === 'above' || atom.parentBranch === 'below')
            return select(model, parent);
        }

        // 4.6/ No eligible group found, move to start of mathfield
        return leapTo(model, 0);
      }
      //
      // 5/ @todo In LaTeX mode, do something?
      //
      return false;
    },
    moveToMathfieldStart: (model: _Model): boolean => {
      if (model.selectionIsCollapsed && model.position === 0) {
        model.announce('plonk');
        return false;
      }

      model.position = 0;
      model.mathfield.stopCoalescingUndo();
      return true;
    },
    moveToMathfieldEnd: (model: _Model): boolean => {
      if (model.selectionIsCollapsed && model.position === model.lastOffset) {
        model.announce('plonk');
        return false;
      }

      model.position = model.lastOffset;
      model.mathfield.stopCoalescingUndo();
      return true;
    },
    moveToSuperscript,
    moveToSubscript,
  },
  { target: 'model', changeSelection: true }
);

register(
  {
    moveToNextPlaceholder: (model) => leap(model, 'forward'),
    moveToPreviousPlaceholder: (model) => leap(model, 'backward'),
  },
  { target: 'model', changeSelection: true }
);

// VisualTeX kernel: Tab is the editing core's (navigation.ts). In the last cell
// of a matrix-like environment it asks the kernel for a new row.
export function visualTexTabAddRow(model: _Model): boolean {
  let child = model.at(model.position);
  while (child && child.parent && !(child.parent instanceof ArrayAtom && isCellBranch(child.parentBranch)))
    child = child.parent;
  const array = child == null ? void 0 : child.parent;
  if (!(array instanceof ArrayAtom)) return false;
  const cell = array.getCell(
    (child!.parentBranch as [number, number])[0],
    (child!.parentBranch as [number, number])[1]
  );
  if (cell != null && cell.length) model.position = model.offsetOf(cell[cell.length - 1]);
  model.mathfield.flushInlineShortcutBuffer();
  model.mathfield.snapshot("addRowAfter");
  return addRowAfter(model);
}

// VisualTeX kernel: & in a matrix, cases or array is LaTeX's column
// separator. It moves to the cell on the right, and at the end of the row it
// adds a column (Alt+Tab, MathLive's key for that, is Windows' window
// switcher). Returns false outside a cell, where & stays a character.
export function visualTexNextColumn(model: _Model): boolean {
  const atom = model.at(range(model.selection)[1]);
  const array = atom == null ? void 0 : atom.parent;
  if (!(array instanceof ArrayAtom) || !isCellBranch(atom.parentBranch) || array.environmentName === "lines")
    return false;
  const [row, col] = atom.parentBranch;
  const right = col + 1 < array.colCount ? array.getCell(row, col + 1) : void 0;
  // VisualTeX editing core: enter the cell at its first empty slot, else its end
  if (right) return visualTexEditingCore.enterSlot(model, { owner: array, key: [row, col + 1] }, 1);
  if (array.maxColumns <= array.colCount) {
    model.announce("plonk");
    return false;
  }
  model.position = model.offsetOf(atom.lastSibling);
  model.mathfield.flushInlineShortcutBuffer();
  model.mathfield.snapshot("addColumnAfter");
  return addColumnAfter(model);
}

register(
  {
    visualTexTabForward: (model) => visualTexEditingCore.tab(model, 1, _MathEnvironment.fractionNavigationOrder, { addRowAfter: () => visualTexTabAddRow(model) }),
    visualTexTabBackward: (model) => visualTexEditingCore.tab(model, -1, _MathEnvironment.fractionNavigationOrder, { addRowAfter: () => false }),
    visualTexNextColumn,
  },
  { target: 'model', changeSelection: true }
);

function findSibling(
  _model: _Model,
  atom: Atom,
  pred: (x: Atom) => boolean,
  dir: 'forward' | 'backward'
): Atom | null {
  if (dir === 'forward') {
    let result = atom.rightSibling;
    while (result && !pred(result)) result = result.rightSibling;
    return result;
  }
  let result = atom.leftSibling;
  while (result && !pred(result)) result = result.leftSibling;
  return result;
}
