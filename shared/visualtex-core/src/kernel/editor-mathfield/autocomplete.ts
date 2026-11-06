import { LatexAtom } from '../atoms/latex';
import {
  suggest,
  visualTexCompletionInfo,
} from '../latex-commands/definitions-utils';
import { visualTexLateFunctions } from './visualtex-hooks';

import type { _Model } from '../editor-model/model-private';

import {
  hideSuggestionPopover,
  showSuggestionPopover,
} from '../editor/suggestion-popover';

import type { _Mathfield } from './mathfield-private';
import {
  getLatexGroupBody,
  getCommandSuggestionRange,
  getLatexGroup,
} from './mode-editor-latex';
import { ModeEditor } from './mode-editor';
import { ParseMode } from '../public/core-types';
import { computeInsertStyle } from './styling';

export function removeSuggestion(mathfield: _Mathfield): void {
  const group = getLatexGroupBody(mathfield.model).filter(
    (x) => x.isSuggestion
  );
  if (group.length === 0) return;
  mathfield.model.position = mathfield.model.offsetOf(group[0].leftSibling);
  for (const atom of group) atom.parent!.removeChild(atom);
}

export function updateAutocomplete(
  mathfield: _Mathfield,
  options?: { atIndex?: number }
): void {
  const { model } = mathfield;
  // Remove any error indicator and any suggestions
  removeSuggestion(mathfield);
  for (const atom of getLatexGroupBody(model)) atom.isError = false;

  if (
    !model.selectionIsCollapsed ||
    mathfield.options.popoverPolicy === 'off'
  ) {
    hideSuggestionPopover(mathfield);
    return;
  }

  // The current command is the sequence of atoms around the insertion point
  // that ends on the left with a '\' and on the right with a non-command
  // character.
  const commandAtoms: LatexAtom[] = [];
  let atom = model.at(model.position);

  while (atom && atom instanceof LatexAtom && /^[a-zA-Z\*]$/.test(atom.value))
    atom = atom.leftSibling;

  if (atom && atom instanceof LatexAtom && atom.value === '\\') {
    // We've found the start of a command.
    // Go forward and collect the potential atoms of the command
    commandAtoms.push(atom);
    atom = atom.rightSibling;
    while (
      atom &&
      atom instanceof LatexAtom &&
      /^[a-zA-Z\*]$/.test(atom.value)
    ) {
      commandAtoms.push(atom);
      atom = atom.rightSibling;
    }
  }

  const command = commandAtoms.map((x) => x.value).join('');
  const suggestions = suggest(mathfield, command);

  if (suggestions.length === 0) {
    mathfield.visualTexCompletion = null;
    // This looks like a command name, but not a known one
    if (/^\\[a-zA-Z\*]+$/.test(command))
      for (const atom of commandAtoms) atom.isError = true;

    hideSuggestionPopover(mathfield);
    return;
  }

  const index = options?.atIndex ?? 0;
  mathfield.suggestionIndex =
    index < 0 ? suggestions.length - 1 : index % suggestions.length;

  // VisualTeX: the typed command stays as typed; the suggestion list keeps
  // the candidates (semantic completion), accepted as a whole
  mathfield.visualTexCompletion = { command, suggestions, group: getLatexGroup(model) };

  showSuggestionPopover(mathfield, suggestions);
}

export function acceptCommandSuggestion(model: _Model): boolean {
  const [from, to] = getCommandSuggestionRange(model, {
    before: model.position,
  });
  if (from === undefined || to === undefined) return false;
  let result = false;
  model.getAtoms([from, to]).forEach((x) => {
    if ((x as LatexAtom).isSuggestion) {
      (x as LatexAtom).isSuggestion = false;
      result = true;
    }
  });
  return result;
}

/**
 * When in LaTeX mode, insert the LaTeX being edited and leave LaTeX mode
 *
 */
export function complete(
  mathfield: _Mathfield,
  completion:
    | 'reject'
    | 'accept'
    | 'accept-suggestion'
    | 'accept-all' = 'accept',
  options?: { mode?: ParseMode; selectItem?: boolean }
): boolean {
  // A source wrapper's command group completes through the wrapper
  const visualTexOwnedSourceGroup = mathfield.visualTexSourceWrapper?.group;
  const visualTexCurrentLatexGroup = getLatexGroup(mathfield.model);
  if (visualTexOwnedSourceGroup && visualTexCurrentLatexGroup === visualTexOwnedSourceGroup) {
    if (completion === 'reject') return visualTexLateFunctions.visualTexCancelSourceWrapper!(mathfield);
    return false;
  }
  const completionState = mathfield.visualTexCompletion;
  const selectedCompletion = completionState?.suggestions[mathfield.suggestionIndex];
  hideSuggestionPopover(mathfield);
  const latexGroup = getLatexGroup(mathfield.model);
  if (!latexGroup) return false;

  const useSemanticCompletion = (completion === 'accept-suggestion' || completion === 'accept-all') && completionState?.group === latexGroup && selectedCompletion;
  if (!useSemanticCompletion && (completion === 'accept-suggestion' || completion === 'accept-all')) {
    const suggestions = getLatexGroupBody(mathfield.model).filter(
      (x) => x.isSuggestion
    );
    if (suggestions.length !== 0) {
      for (const suggestion of suggestions) suggestion.isSuggestion = false;

      mathfield.model.position = mathfield.model.offsetOf(
        suggestions[suggestions.length - 1]
      );
    }
    if (completion === 'accept-suggestion') return suggestions.length !== 0;
  }

  const body = getLatexGroupBody(mathfield.model).filter(
    (x) => !x.isSuggestion
  );

  const latex = useSemanticCompletion ? selectedCompletion : body.map((x) => x.value).join('');
  mathfield.visualTexCompletion = null;

  const newPos = latexGroup.leftSibling;
  latexGroup.parent!.removeChild(latexGroup);
  mathfield.model.position = mathfield.model.offsetOf(newPos);
  mathfield.switchMode(options?.mode ?? 'math');

  if (completion === 'reject') {
    mathfield.visualTexPendingStructuralCommandSlot = null;
    return true;
  }

  const style = { ...computeInsertStyle(mathfield) };
  // If we're inserting a non-alphanumeric character, reset the variant
  if (!/^[a-zA-Z0-9]$/.test(latex) && mathfield.styleBias !== 'none') {
    style.variant = 'normal';
    style.variantStyle = undefined;
  }

  const visualTexAcceptedInfo = useSemanticCompletion
    ? visualTexCompletionInfo(mathfield, selectedCompletion)
    : null;
  const visualTexSourceWrapperCommand =
    visualTexAcceptedInfo?.command &&
    visualTexLateFunctions.visualTexSourceWrapperKind!(visualTexAcceptedInfo.command)
      ? visualTexLateFunctions.visualTexCanonicalVariantInputCommand!(visualTexAcceptedInfo.command)
      : null;
  if (visualTexSourceWrapperCommand) {
    const visualTexSourceSlot =
      mathfield.visualTexPendingStructuralCommandSlot || null;
    const visualTexStartedSourceWrapper = visualTexLateFunctions.visualTexStartSourceWrapper!(
      mathfield.model,
      visualTexSourceWrapperCommand + '{\\placeholder{}}',
      visualTexSourceSlot,
    );
    mathfield.snapshot('insert-source-wrapper');
    if (visualTexStartedSourceWrapper) {
      const info = visualTexAcceptedInfo!;
      // MathLive contentDidChange() dispatches its public input event on a
      // zero-delay timer. Queue usage/acceptance after that timer so a
      // controlled host cannot rerender with the stale pre-completion value
      // before the model's input transaction has been observed.
      setTimeout(() => {
        if (!mathfield.host) return;
        mathfield.host.dispatchEvent(new CustomEvent('visualtex-command-accepted', {
          bubbles: true,
          composed: true,
          detail: {
            id: info.id || info.command.slice(1),
            query: completionState!.command,
            command: info.command,
            latex: selectedCompletion,
            variantInput: true,
          },
        }));
      }, 0);
      mathfield.model.announce('replacement');
      return true;
    }
  }

  ModeEditor.insert(mathfield.model, latex, {
    selectionMode: (options?.selectItem ?? false) ? 'item' : 'placeholder',
    format: 'latex',
    mode: 'math',
    style,
  });

  mathfield.snapshot();
  if (useSemanticCompletion) {
    const info = visualTexAcceptedInfo!;
    setTimeout(() => {
      if (!mathfield.host) return;
      mathfield.host.dispatchEvent(new CustomEvent('visualtex-command-accepted', {
        bubbles: true,
        composed: true,
        detail: {
          id: info.id || info.command.slice(1),
          query: completionState.command,
          command: info.command,
          latex: selectedCompletion,
          variantInput: false,
        },
      }));
    }, 0);
  }
  mathfield.model.announce('replacement');
  mathfield.switchMode('math');
  return true;
}
