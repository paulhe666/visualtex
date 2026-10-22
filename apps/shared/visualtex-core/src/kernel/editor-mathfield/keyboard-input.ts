import type { Selector } from '../public/commands';
import { splitGraphemes } from '../core/grapheme-splitter';
import { Atom } from '../core/atom';

import { keyboardEventToChar, keyboardEventToString } from '../editor/keyboard';
import { getInlineShortcut } from '../editor/shortcuts';
import { getCommandForKeybinding } from '../editor/keybindings';
import {
  getActiveKeyboardLayout,
  validateKeyboardLayout,
} from '../editor/keyboard-layout';

import { moveAfterParent } from '../editor-model/commands-move';
import { range } from '../editor-model/selection-utils';

import { complete, removeSuggestion, updateAutocomplete } from './autocomplete';
import { getLatexGroupBody } from './mode-editor-latex';
import { getDefinition } from '../latex-commands/definitions-utils';
import { requestUpdate } from './render';
import type { _Mathfield } from './mathfield-private';
import { removeIsolatedSpace, smartMode } from './smartmode';
import { ModeEditor } from './mode-editor';
import type { ParseMode, Style } from '../public/core-types';
import type { _Model } from '../editor-model/model-private';
import { LeftRightAtom } from '../atoms/leftright';
import { RIGHT_DELIM, LEFT_DELIM } from '../core/delimiters';
import { mightProducePrintableCharacter } from '../ui/events/utils';
import { computeInsertStyle } from './styling';
import {
  visualTexEditingCore,
  visualTexLateFunctions,
} from './visualtex-hooks';
import { LatexGroupAtom, LatexAtom } from '../atoms/latex';
import {
  visualTexStructuralSlotStillOccupied,
  visualTexRestoreStructuralPlaceholder,
  visualTexShouldRetainStructuralPlaceholder,
} from '../editor-model/delete';
import {
  LATEX_COMMANDS,
  MATH_SYMBOLS,
  visualTexCompletionInfo,
} from '../latex-commands/definitions-utils';
import { parseLatex } from '../core/parser';
import { hideSuggestionPopover } from '../editor/suggestion-popover';
import {
  leapTo,
  leapTarget,
  visualTexIsBoundedOperatorParent,
  visualTexBoundedOperatorBranchPlaceholders,
} from '../editor-model/commands-move';
import {
  applyInputJudgment as visualTexApplyInputJudgment,
  visualTexSourceWrapperKindsData,
  visualTexSourceWrapperCanonicalTargetsData,
} from 'virtual:visualtex-kernel-host';
import { register } from '../editor/commands';
import {
  visualTexCandidateIndexForDigit,
  visualTexCandidatePanelOpen,
} from '../latex-commands/definitions-utils';
import {
  visualTexFreshStructuralSlot,
  type VisualTexStructuralSlot,
} from '../editor-model/delete';
import { getLatexGroup } from './mode-editor-latex';
import { visualTexLiveStructuralSlot } from '../editor-model/delete';

// src/editor-mathfield/keyboard-input.ts
export function visualTexCaretIsInsideTextRun(model: _Model): boolean {
  // VisualTeX editing core (text.ts): between two characters of one run
  return model.selectionIsCollapsed && visualTexEditingCore.inTextRun(model, model.position);
}
export function visualTexNormalizeTextBoundaryMode(mathfield: _Mathfield, forceMath = false): boolean {
  const { model } = mathfield;
  if (model.mode !== "text" || !model.selectionIsCollapsed) return false;
  // MathLive stores a caret at the offset of the atom to its left. A caret is
  // genuinely inside prose only when both sides belong to the same text run.
  // Every other text-adjacent position is a math boundary in VisualTeX.
  if (!forceMath && visualTexCaretIsInsideTextRun(model)) return false;
  model.mode = "math";
  return true;
}

/**
 * Handler in response to a keystroke event (or to a virtual keyboard keycap
 * with a `key` property).
 *
 * Return `false` if the event has been handled as a shortcut or command and
 * needs no further processing.
 *
 * Return `true` if the event should be handled as a regular textual input.
 *
 *
 * ## Theory of Operation
 *
 * When the user types on the keyboard, printable keys (i.e. not arrows, shift,
 * escape, etc...) are captured in a `inlineShortcutBuffer`.
 *
 * The buffer is used to determine if the user intended to type an
 * inline shortcut (e.g. "pi" for `\pi`) or a multichar symbol.
 *
 * Characters are added to this buffer while the user type printable characters
 * consecutively. If the user change selection (with the mouse, or by
 * navigating with the keyboard), if an unambiguous match for the buffer is
 * found, the buffer is cleared.
 *
 * Associated with this buffer are `states`
 *
 */
export function onKeystroke(
  mathfield: _Mathfield,
  evt: KeyboardEvent
): boolean {
  const { model } = mathfield;
  visualTexNormalizeTextBoundaryMode(mathfield);
  let keystroke = keyboardEventToString(evt);
  if (keystroke === "[Tab]" && model.mode === "latex")
    keystroke = mathfield.visualTexSourceWrapper && !visualTexCandidatePanelOpen(mathfield) ? "[Enter]" : "[Space]";
  if (evt.isTrusted) {
    validateKeyboardLayout(evt);
    const activeLayout = getActiveKeyboardLayout();
    if (mathfield.keyboardLayout !== activeLayout.id) {
      mathfield.keyboardLayout = activeLayout.id;
      mathfield._keybindings = void 0;
    }
  }
  clearTimeout(mathfield.inlineShortcutBufferFlushTimer);
  mathfield.inlineShortcutBufferFlushTimer = 0;
  if (evt.isTrusted && evt.defaultPrevented) {
    mathfield.flushInlineShortcutBuffer();
    return false;
  }
  // VisualTeX kernel: a key that finishes a typed word runs the input
  // judgment first (part-021); the Space that finished a converted word and
  // the Backspace that reverted one are used up here.
  if (visualTexJudgeKeystroke(mathfield, evt, keystroke)) {
    mathfield.flushInlineShortcutBuffer();
    if (evt.preventDefault) {
      evt.preventDefault();
      evt.stopPropagation();
    }
    return false;
  }
  let shortcut: string | undefined;
  let shortcutLength = 0;
  let selector: Selector | '' | [Selector, ...any[]] = '';
  let stateIndex = 0;
  const buffer = mathfield.inlineShortcutBuffer;
  let placeholderReplaced = false;
  mathfield.visualTexPendingPlaceholderSelection = null;
  if (mathfield.isSelectionEditable && model.selectionIsPlaceholder && mightProducePrintableCharacter(evt)) {
    const selected = model.at(Math.max(model.position, model.anchor));
    // A direct glyph replacement is not a continuation of a previous raw command.
    mathfield.visualTexPendingStructuralCommandSlot = null;
    mathfield.visualTexPendingPlaceholderSelection = {
      position: Math.min(model.position, model.anchor),
      style: { ...selected.style },
      slot: visualTexFreshStructuralSlot(selected),
    };
    // Keep the placeholder alive until the actual insert transaction.
    // If beforeinput is cancelled, the model must remain byte-for-byte unchanged.
    mathfield.flushInlineShortcutBuffer();
    placeholderReplaced = true;
  }
  if (mathfield.isSelectionEditable) {
    if (model.mode === "math" && (!model.selectionIsPlaceholder || placeholderReplaced)) {
      if (keystroke === "[Backspace]") {
        if (mathfield.undoManager.lastOp === "insert-shortcut")
          selector = "undo";
        else buffer.pop();
      } else if (!mightProducePrintableCharacter(evt)) {
        mathfield.flushInlineShortcutBuffer();
      } else {
        const c = keyboardEventToChar(evt);
        const keystrokes = [
          ...buffer[buffer.length - 1]?.keystrokes ?? [],
          c
        ];
        buffer.push({
          state: model.getState(),
          keystrokes,
          leftSiblings: getLeftSiblings(mathfield)
        });
        shortcutLength = 0;
        let candidate = "";
        while (!shortcut && shortcutLength < keystrokes.length) {
          stateIndex = buffer.length - (keystrokes.length - shortcutLength);
          candidate = keystrokes.slice(shortcutLength).join("");
          shortcut = getInlineShortcut(
            buffer[stateIndex].leftSiblings,
            candidate,
            mathfield.options.inlineShortcuts
          );
          if (!shortcut && /^[a-zA-Z][a-zA-Z0-9]+?([_\^][a-zA-Z0-9\*\+\-]+?)?$/.test(candidate))
            shortcut = mathfield.options.onInlineShortcut(mathfield, candidate);
          shortcutLength += 1;
        }
        mathfield.flushInlineShortcutBuffer({ defer: true });
      }
    }
    if (mathfield.options.smartMode) {
      if (shortcut) {
        mathfield.switchMode("math");
      } else if (smartMode(mathfield, keystroke, evt)) {
        mathfield.switchMode(
          ({ math: "text", text: "math" } as Record<string, ParseMode>)[model.mode]
        );
        selector = "";
      }
    }
  }
  if (!shortcut) {
    if (!selector) {
      selector = getCommandForKeybinding(
        mathfield.keybindings,
        model.mode,
        evt
      );
    }
    // VisualTeX persistent variant input (source wrapper): Enter commits it,
    // Backspace in an empty one cancels it. Every other Enter belongs to the
    // editor (VisualTeX editing core, keys.ts).
    if (mathfield.visualTexSourceWrapper &&
        (keystroke === "[Enter]" || keystroke === "[Return]")) {
      const visualTexCommittedSourceWrapper = visualTexCommitSourceWrapper(mathfield);
      if (visualTexCommittedSourceWrapper) {
        mathfield.snapshot("insert-source-wrapper");
        mathfield.model.announce("replacement");
      }
      if (evt.preventDefault) evt.preventDefault();
      if (evt.stopPropagation) evt.stopPropagation();
      return false;
    }
    if (mathfield.visualTexSourceWrapper && keystroke === "[Backspace]" &&
        visualTexSourceWrapperContent(mathfield.visualTexSourceWrapper) === "") {
      visualTexCancelSourceWrapper(mathfield);
      mathfield.snapshot("delete-source-wrapper");
      if (evt.preventDefault) evt.preventDefault();
      if (evt.stopPropagation) evt.stopPropagation();
      return false;
    }
    const visualTexPickedCandidate = visualTexCandidateIndexForDigit(mathfield, evt);
    if (visualTexPickedCandidate === -2) {
      if (evt.preventDefault) { evt.preventDefault(); evt.stopPropagation(); }
      return false;
    }
    if (visualTexPickedCandidate >= 0) {
      mathfield.suggestionIndex = visualTexPickedCandidate;
      keystroke = "[Space]";
    }
    if (keystroke === "[Space]" && model.mode === "latex" && mathfield.visualTexSourceWrapper) {
      const visualTexState = mathfield.visualTexSourceWrapper;
      const visualTexGroup = getLatexGroup(model);
      if (visualTexGroup === visualTexState.group) {
        const visualTexAcceptedSuggestion = visualTexAcceptSourceWrapperCompletion(mathfield);
        visualTexSyncSourceWrapper(mathfield);
        const visualTexContent = visualTexSourceWrapperContent(visualTexState);
        const visualTexCanCommit = visualTexSourceWrapperAutoExitEnabled(mathfield, visualTexState) &&
          (Array.from(visualTexContent || "").length === 1 || visualTexSourceWrapperSimpleCommand(mathfield, visualTexContent || ""));
        if (visualTexCanCommit) {
          visualTexCommitSourceWrapper(mathfield);
          mathfield.snapshot("insert-source-wrapper");
          mathfield.model.announce("replacement");
        } else if (visualTexAcceptedSuggestion) {
          mathfield.snapshot("insert-latex-suggestion");
          requestUpdate(mathfield);
        }
        mathfield.dirty = true;
        mathfield.scrollIntoView();
        if (evt.preventDefault) { evt.preventDefault(); evt.stopPropagation(); }
        return false;
      }
    }
    if (keystroke === "[Space]" && model.mode === "latex") {
      if (complete(mathfield, "accept-all")) {
        mathfield.dirty = true;
        mathfield.scrollIntoView();
        if (evt.preventDefault) {
          evt.preventDefault();
          evt.stopPropagation();
        }
        return false;
      }
    }
    if ((!selector || keystroke === "[Space]") && model.mode === "math") {
      if (keystroke === "[Space]") {
        mathfield.styleBias = "none";
        mathfield.flushInlineShortcutBuffer();
        if (mathfield.options.mathModeSpace) {
          ModeEditor.insert(model, mathfield.options.mathModeSpace, {
            format: "latex",
            mode: "math"
          });
          mathfield.snapshot("insert-space");
          selector = "";
          mathfield.dirty = true;
          mathfield.scrollIntoView();
          if (evt.preventDefault) {
            evt.preventDefault();
            evt.stopPropagation();
          }
          return false;
        }
      }
      if ((model.at(model.position)?.isDigit()) && globalThis.MathfieldElement.decimalSeparator === "," && keyboardEventToChar(evt) === ",")
        selector = "insertDecimalSeparator";
    }
  }
  if (!shortcut && !selector) {
    if (!model.mathfield.smartFence) {
      const { parent: parent2 } = model.at(model.position);
      if (parent2 instanceof LeftRightAtom && parent2.rightDelim === "?" && model.at(model.position).isLastSibling && /^[)}\]|]$/.test(keystroke)) {
        mathfield.snapshot();
        parent2.isDirty = true;
        parent2.rightDelim = keystroke;
        model.position += 1;
        model.selectionDidChange();
        model.contentDidChange({
          data: keyboardEventToChar(evt),
          inputType: "insertText"
        });
        mathfield.snapshot("insert-fence");
        mathfield.dirty = true;
        mathfield.scrollIntoView();
        if (evt.preventDefault) evt.preventDefault();
        return false;
      }
      if (!model.selectionIsCollapsed) {
        const fence = keyboardEventToChar(evt);
        if (fence === "(" || fence === "{" || fence === "[") {
          const lDelim = { "(": "(", "{": "\\lbrace", "[": "\\lbrack" }[fence];
          const rDelim = { "(": ")", "{": "\\rbrace", "[": "\\rbrack" }[fence];
          const [start, end] = range(model.selection);
          mathfield.snapshot();
          model.position = end;
          ModeEditor.insert(model, rDelim, { format: "latex" });
          model.position = start;
          ModeEditor.insert(model, lDelim, { format: "latex" });
          model.setSelection(start + 1, end + 1);
          model.contentDidChange({
            data: fence,
            inputType: "insertText"
          });
          mathfield.snapshot("insert-fence");
          mathfield.dirty = true;
          mathfield.scrollIntoView();
          if (evt.preventDefault) evt.preventDefault();
          return false;
        }
      }
    } else if (insertSmartFence(
      model,
      keyboardEventToChar(evt),
      computeInsertStyle(mathfield)
    )) {
      mathfield.dirty = true;
      mathfield.scrollIntoView();
      if (evt.preventDefault) evt.preventDefault();
      return false;
    }
    return true;
  }
  const child = model.at(Math.max(model.position, model.anchor));
  const { parent } = child;
  if (selector === "moveAfterParent" && parent instanceof LeftRightAtom && child.isLastSibling && mathfield.options.smartFence && parent.rightDelim === "?") {
    mathfield.snapshot();
    parent.rightDelim = parent.matchingRightDelim();
    parent.isDirty = true;
    model.contentDidChange({ inputType: "insertText" });
    mathfield.snapshot("insert-fence");
  }
  mathfield.keyboardDelegate.cancelComposition();
  if (selector) mathfield.executeCommand(selector);
  else if (shortcut) {
    const style = computeInsertStyle(mathfield);
    model.setState(buffer[stateIndex].state);
    let keystrokes = buffer[buffer.length - 1].keystrokes;
    keystrokes = keystrokes.slice(shortcutLength - 1);
    for (const c of keystrokes) {
      ModeEditor.insert(model, c, {
        silenceNotifications: true,
        style
      });
    }
    mathfield.snapshot(`insert-shortcut`);
    model.setState(buffer[stateIndex].state);
    model.deferNotifications(
      {
        content: true,
        selection: true,
        data: shortcut,
        type: "insertText"
      },
      () => {
        ModeEditor.insert(model, shortcut, { format: "latex", style });
        removeIsolatedSpace(mathfield.model);
        if (shortcut.endsWith(" ")) {
          mathfield.switchMode("text");
          ModeEditor.insert(model, " ", { style, mode: "text" });
        }
        mathfield.snapshot();
        if (!model.selectionIsCollapsed) mathfield.flushInlineShortcutBuffer();
        return true;
      }
    );
    mathfield.dirty = true;
    model.announce("replacement");
  }
  mathfield.scrollIntoView();
  if (evt.preventDefault) evt.preventDefault();
  return false;
}

/**
 * Detect if the content before the cursor matches a scientific notation pattern.
 * Returns the match information if found, null otherwise.
 *
 * Pattern: digits + (e|E) + optional(+|-) + digits
 * Examples: 3.14e2, 5E-3, 1.23e+10
 */
function detectScientificNotation(model: _Model): {
  startOffset: number;
  endOffset: number;
  significand: string;
  exponent: string;
} | null {
  const { position } = model;

  // Get atoms at and to the left of the cursor
  // We need to check from position (not position-1) because when a non-inserting
  // character like space is pressed, the cursor doesn't advance
  let offset = position;
  const atoms: Atom[] = [];

  // Collect atoms going backwards from cursor position
  // Include digits, decimal points, 'e'/'E', and +/- signs
  while (offset > 0) {
    const atom = model.at(offset);
    const value = atom.value;

    // Only collect atoms that could be part of scientific notation
    if (
      atom.type === 'mord' ||
      (atom.type === 'mbin' &&
        (value === '+' || value === '-' || value === '\u2212'))
    ) {
      atoms.unshift(atom);
      offset--;
    } else break;
  }

  if (atoms.length === 0) return null;

  // Build the string from atoms
  const text = atoms.map((a) => a.value).join('');

  // Match scientific notation pattern: significand e|E [+|-] exponent
  // Respect the localized decimal separator
  const separator = globalThis.MathfieldElement?.decimalSeparator ?? '.';
  const separatorRegex = separator === '.' ? '\\.' : ',';
  // Use non-capturing group (?:...) to avoid shifting match groups
  const pattern = new RegExp(
    `^(\\d+(?:${separatorRegex}\\d*)?)[eE]([+\\-\u2212]?)(\\d+)$`
  );
  const match = text.match(pattern);

  if (!match) return null;

  const significand = match[1];
  const sign = match[2];
  const exponentDigits = match[3];
  const exponent = sign + exponentDigits;

  return {
    startOffset: offset,
    endOffset: position,
    significand,
    exponent,
  };
}

/**
 * Apply the scientific notation template to format the detected pattern.
 * Template uses #1 for significand and #2 for exponent.
 */
function applyScientificNotationTemplate(
  significand: string,
  exponent: string
): string | null {
  const template = globalThis.MathfieldElement?.scientificNotationTemplate;

  // Validate template
  if (
    !template ||
    template === '' ||
    !template.includes('#1') ||
    !template.includes('#2')
  )
    return null;

  // Replace placeholders
  let result = template.replace('#1', significand);
  result = result.replace('#2', exponent);

  return result;
}

/**
 * Check if scientific notation should be formatted and apply the template if applicable.
 * This is called when a non-digit character is typed or after a timeout.
 */
function formatScientificNotationIfApplicable(mathfield: _Mathfield): boolean {
  const { model } = mathfield;

  // Only format in math mode
  if (model.mode !== 'math') return false;

  // Detect scientific notation pattern
  const match = detectScientificNotation(model);
  if (!match) return false;

  // Apply template
  const formatted = applyScientificNotationTemplate(
    match.significand,
    match.exponent
  );
  if (!formatted) return false;

  // Replace the matched range with the formatted template
  model.deferNotifications(
    { content: true, selection: true, type: 'insertText' },
    () => {
      // Select the scientific notation atoms
      model.selection = {
        ranges: [[match.startOffset, match.endOffset]],
        direction: 'forward',
      };

      // Insert the formatted template
      ModeEditor.insert(model, formatted, {
        insertionMode: 'replaceSelection',
        selectionMode: 'after',
      });
    }
  );

  mathfield.snapshot('format-scientific-notation');
  mathfield.dirty = true;
  mathfield.scrollIntoView();

  return true;
}

/**
 * This handler is invoked when text has been input with an input method.
 * As a result, `text` can be a sequence of characters to be inserted.
 * @param {object} options
 * @param {boolean} options.focus - If true, the mathfield will be focused
 * @param {boolean} options.feedback - If true, provide audio and haptic feedback
 * @param {boolean} options.simulateKeystroke - If true, generate some synthetic
 * keystrokes (useful to trigger inline shortcuts, for example)
 * @private
 */
export function onInput(
  mathfield: _Mathfield,
  text: string,
  options?: {
    focus?: boolean;
    mode?: ParseMode;
    simulateKeystroke?: boolean;
  }
): void {
  const { model } = mathfield;
  visualTexNormalizeTextBoundaryMode(mathfield);
  if (!mathfield.isSelectionEditable) {
    model.announce("plonk");
    return;
  }
  options ??= {};
  if (options.focus) mathfield.focus();
  if (typeof options.mode === "string") {
    mathfield.switchMode(options.mode);
    mathfield.snapshot();
  }
  const graphemes = splitGraphemes(text);
  if (options.simulateKeystroke) {
    let handled = true;
    for (const c of graphemes) {
      if (onKeystroke(mathfield, new KeyboardEvent("keypress", { key: c })))
        handled = false;
    }
    if (handled) return;
  }
  if (model.mode === "latex") {
    model.deferNotifications(
      { content: true, selection: true, data: text, type: "insertText" },
      () => {
        removeSuggestion(mathfield);
        for (const c of graphemes)
          ModeEditor.insert(model, c, { insertionMode: "replaceSelection" });
        mathfield.snapshot("insert-latex");
        updateAutocomplete(mathfield);
      }
    );
    if (text === "}") {
      const latexBody = getLatexGroupBody(model);
      const latex = latexBody.map((x) => x.value).join("");
      const commandMatch = latex.match(/^\\([a-zA-Z]+)/);
      if (commandMatch) {
        const commandName = "\\" + commandMatch[1];
        const def = getDefinition(commandName, "math");
        if ((def?.definitionType) === "function") {
          const mandatoryArgCount = def.params.filter(
            (p) => !p.isOptional
          ).length;
          if (mandatoryArgCount > 0) {
            let depth = 0;
            let completedBraces = 0;
            let inCommandName = true;
            for (let i = 0; i < latex.length; i++) {
              const char = latex[i];
              if (inCommandName) {
                if (char === "\\" || /[a-zA-Z]/.test(char)) continue;
                inCommandName = false;
              }
              if (char === "{") depth++;
              else if (char === "}") {
                depth--;
                if (depth === 0) completedBraces++;
              }
            }
            if (depth === 0 && completedBraces === mandatoryArgCount) {
              if (complete(mathfield, "accept-all")) {
                mathfield.dirty = true;
                mathfield.scrollIntoView();
                return;
              }
            }
          }
        }
      }
    }
  } else if (model.mode === "text") {
    const style = { ...getSelectionStyle(model), ...mathfield.defaultStyle };
    for (const c of graphemes)
      ModeEditor.insert(model, c, { style, insertionMode: "replaceSelection" });
    mathfield.snapshot("insert-text");
  } else if (model.mode === "math")
    for (const c of graphemes) insertMathModeChar(mathfield, c);
  mathfield.dirty = true;
  mathfield.scrollIntoView();
}

// VisualTeX kernel: the input judgment. One routine and one dictionary
// (src/editor/inputJudgment.ts) decide what a typed word turns into (quick
// escapes: alpha -> \alpha, sintheta -> \sin\theta) and what is set upright
// (units, function names, constants, labels, the differential d). A word is
// judged once finished: by a key that is not a letter, a caret move, or
// leaving the field. The host sets visualTexJudgmentSettings = { upright,
// escapes } from the editor's two switches.
/** The editor element: the host gives the judgment's two switches */
type VisualTexJudgingHost = HTMLElement & {
  visualTexJudgmentSettings?: { upright?: boolean; escapes?: boolean };
};
export function visualTexJudgmentSettings(mathfield: _Mathfield): { upright: boolean; escapes: boolean } {
  const settings = (mathfield.host as VisualTexJudgingHost | undefined)
    ?.visualTexJudgmentSettings;
  return {
    upright: (settings?.upright) !== false,
    escapes: (settings?.escapes) !== false
  };
}
// A Latin letter typed in the default math style. Letters typed after the
// first carry variant "normal", which is still the default (italic);
// \mathrm letters add an upright variantStyle.
export function visualTexIsPlainLetter(atom: Atom | undefined): atom is Atom {
  return Boolean(atom) && atom!.type === "mord" && atom!.mode === "math" && /^[A-Za-z]$/.test(atom!.value || "") && (!atom!.style.variant || atom!.style.variant === "normal") && !atom!.style.variantStyle;
}
// The plain (italic) Latin letters right before the caret.
export function visualTexWordBeforeCaret(model: _Model): string {
  let text = "";
  let atom = model.at(model.position);
  while (visualTexIsPlainLetter(atom)) {
    text = atom.value + text;
    atom = atom.leftSibling;
  }
  return text;
}
// The Latin letters right before the caret in any style (km stays km when
// it only turns upright).
export function visualTexLettersBeforeCaret(model: _Model): string {
  let text = "";
  let atom = model.at(model.position);
  while (atom && atom.type === "mord" && /^[A-Za-z]$/.test(atom.value || "")) {
    text = atom.value + text;
    atom = atom.leftSibling;
  }
  return text;
}
export function visualTexSetJudgedValue(mathfield: _Mathfield, latex: string): void {
  const model = mathfield.model;
  // VisualTeX editing core (restore.ts): the caret follows its atom, or goes
  // after the conversion that replaced it
  const caret = visualTexEditingCore.captureCaret(model);
  mathfield.setValue(latex, {
    mode: "math",
    format: "latex",
    insertionMode: "replaceAll",
    selectionMode: "after"
  });
  visualTexEditingCore.restoreCaret(model, caret);
}
/**
 * Applies the input judgment to the whole formula. `finishing`: the word at
 * the caret is finished too. `typed: false` judges like stored LaTeX: upright
 * typography only, no quick escapes. Returns whether the formula changed.
 */
export function visualTexJudgeInput(
  mathfield: _Mathfield,
  options?: { typed?: boolean; finishing?: boolean }
): boolean {
  const model = mathfield.model;
  if (model.mode !== "math" || !model.selectionIsCollapsed || !mathfield.isSelectionEditable) return false;
  const settings = visualTexJudgmentSettings(mathfield);
  const typed = (options?.typed) !== false;
  const judgment = {
    upright: settings.upright,
    escapes: typed && settings.escapes,
    typed,
    deferLetterRuns: !(options?.finishing) && visualTexWordBeforeCaret(model) !== ""
  };
  const before = model.getValue();
  const after = visualTexApplyInputJudgment(before, judgment);
  if (after === before) return false;
  const revert = visualTexApplyInputJudgment(before, { ...judgment, revert: true });
  visualTexSetJudgedValue(mathfield, after);
  mathfield.visualTexLastJudgment = {
    after: model.getValue(),
    revert,
    position: model.position
  };
  return true;
}
// Backspace right after a conversion restores the typed word, marked
// explicit (\mathit) so the judgment leaves it alone from then on.
export function visualTexRevertLastJudgment(mathfield: _Mathfield): boolean {
  const last = mathfield.visualTexLastJudgment;
  mathfield.visualTexLastJudgment = void 0;
  const model = mathfield.model;
  if (!last || !model.selectionIsCollapsed || model.getValue() !== last.after || model.position !== last.position)
    return false;
  visualTexSetJudgedValue(mathfield, last.revert);
  return true;
}
/**
 * Called for every keystroke before MathLive handles it. Returns true when
 * the keystroke was used up: the Space that finished a converted word, or
 * the Backspace that reverted a conversion.
 */
export function visualTexJudgeKeystroke(
  mathfield: _Mathfield,
  evt: KeyboardEvent,
  keystroke: string
): boolean {
  const model = mathfield.model;
  if (model.mode !== "math" || evt.isComposing) return false;
  if (keystroke === "[Backspace]") return visualTexRevertLastJudgment(mathfield);
  mathfield.visualTexLastJudgment = void 0;
  if (evt.ctrlKey || evt.metaKey || evt.altKey) return false;
  const key = evt.key ?? "";
  if (/^[A-Za-z]$/.test(key) || /^(?:Shift|Control|Alt|Meta|CapsLock|Delete|Process|Dead|Unidentified)$/.test(key))
    return false;
  const wordBefore = visualTexWordBeforeCaret(model);
  if (!visualTexJudgeInput(mathfield, { finishing: true })) return false;
  // The word at the caret became something else (not just upright, as a
  // unit stays the same letters): step into a template's first slot, and
  // the Space that finished it is not inserted.
  if (!wordBefore || visualTexLettersBeforeCaret(model).endsWith(wordBefore)) return false;
  const left = model.at(model.position);
  const slot = left ? visualTexEditingCore.firstEmptySlotStop(model, [left]) : void 0;
  if (slot !== void 0) visualTexEditingCore.goToStop(model, slot);
  return keystroke === "[Space]";
}

register(
  {
    visualTexJudgeInput: (mathfield: _Mathfield, options) => visualTexJudgeInput(mathfield, options),
  },
  { target: 'mathfield' }
);

function getLeftSiblings(mf: _Mathfield): Atom[] {
  const model = mf.model;

  const result: Atom[] = [];
  let atom = model.at(Math.min(model.position, model.anchor));
  while (atom.type !== 'first') {
    result.push(atom);
    atom = atom.leftSibling!;
  }

  return result;
}

/** VisualTeX: a command typed in the LaTeX box around its argument (\mathbf{...}), not yet committed */
export type VisualTexSourceWrapperState = {
  command: string;
  kind: string;
  group: LatexGroupAtom;
  structuralSlot: VisualTexStructuralSlot | null;
};

export const VISUALTEX_SOURCE_WRAPPER_KINDS = visualTexSourceWrapperKindsData;
export const VISUALTEX_SOURCE_WRAPPER_CANONICAL_TARGETS = visualTexSourceWrapperCanonicalTargetsData;
export function visualTexCanonicalVariantInputCommand(command: string): string {
  return VISUALTEX_SOURCE_WRAPPER_CANONICAL_TARGETS[command] || command;
}
export function visualTexSourceWrapperKind(command: string): string | null {
  const canonical = visualTexCanonicalVariantInputCommand(command);
  return VISUALTEX_SOURCE_WRAPPER_KINDS[command] ||
    VISUALTEX_SOURCE_WRAPPER_KINDS[canonical] || null;
}
export function visualTexIsAccentSourceCommand(command: string): boolean {
  return visualTexSourceWrapperKind(command) === "accent";
}
export function visualTexSourceWrapperTemplate(input: unknown): { command: string; kind: string } | null {
  if (typeof input !== "string") return null;
  const match = input.match(/^(\\[a-zA-Z]+)\{\\placeholder\{\}\}$/);
  if (!match) return null;
  const command = visualTexCanonicalVariantInputCommand(match[1]);
  const kind = visualTexSourceWrapperKind(command);
  return kind ? { command, kind } : null;
}
export function visualTexSourceWrapperRaw(state: VisualTexSourceWrapperState | null | undefined): string | null {
  const group = state?.group;
  if (!group?.parent) return null;
  return (group.body || []).filter((atom) => atom.type === "latex" && !(atom as LatexAtom).isSuggestion).map((atom) => atom.value).join("");
}
export function visualTexSourceWrapperContent(state: VisualTexSourceWrapperState | null | undefined): string | null {
  const raw = visualTexSourceWrapperRaw(state);
  const prefix = state?.command + "{";
  if (raw == null || !raw.startsWith(prefix) || !raw.endsWith("}")) return null;
  return raw.slice(prefix.length, -1);
}
export function visualTexSyncSourceWrapper(mathfield: _Mathfield): void {
  let state = mathfield.visualTexSourceWrapper;
  if (state) {
    const group = state.group;
    const offset = group?.parent ? mathfield.model.offsetOf(group) : -1;
    if (!group?.parent || offset < 0 || mathfield.model.at(offset) !== group) {
      mathfield.visualTexSourceWrapper = null;
      state = null;
    }
  }
  const host = mathfield.host;
  if (!host) return;
  const content = state ? visualTexSourceWrapperContent(state) : null;
  if (state && content != null) {
    host.dataset.visualtexSourceWrapperCommand = state.command;
    host.dataset.visualtexSourceWrapperKind = state.kind;
    host.dataset.visualtexSourceWrapperContent = content;
  } else {
    delete host.dataset.visualtexSourceWrapperCommand;
    delete host.dataset.visualtexSourceWrapperKind;
    delete host.dataset.visualtexSourceWrapperContent;
  }
}
export function visualTexSourceWrapperAutoExitEnabled(
  mathfield: _Mathfield,
  state: VisualTexSourceWrapperState | null | undefined
): boolean {
  if (state?.kind === "accent") return (mathfield.host == null ? void 0 : mathfield.host.dataset.visualtexAutoExitAccent) === "true";
  return (mathfield.host == null ? void 0 : mathfield.host.dataset.visualtexAutoExitWrapperCommand) !== "false";
}
export function visualTexSourceWrapperCaretIsInside(
  mathfield: _Mathfield,
  state: VisualTexSourceWrapperState | null | undefined
): boolean {
  const current = mathfield.model.at(mathfield.model.position);
  return Boolean(state?.group?.parent && (current === state.group || current?.parent === state.group));
}
export function visualTexStartSourceWrapper(
  model: _Model,
  input: string,
  structuralSlot: VisualTexStructuralSlot | null = null
): boolean {
  const mathfield = model.mathfield;
  const info = visualTexSourceWrapperTemplate(input);
  if (!info) return false;
  const shell = info.command + "{}";
  const group = new LatexGroupAtom(shell);
  const cursor = model.at(model.position);
  if (!cursor?.parent) return false;
  cursor.parent.addChildAfter(group, cursor);
  const latexAtoms = (group.body || []).filter((atom) => atom.type === "latex");
  const close = latexAtoms[latexAtoms.length - 1];
  const beforeClose = close?.leftSibling;
  if (!close || !beforeClose) {
    group.parent?.removeChild(group);
    return false;
  }
  const state = { command: info.command, kind: info.kind, group, structuralSlot };
  group.visualTexSourceWrapper = state;
  if (structuralSlot) group.visualTexStructuralSlot = structuralSlot;
  mathfield.visualTexSourceWrapper = state;
  // Keep the raw-shell selection private until the matching input event
  // has updated host state. Publishing selection first can make a host
  // render stale content and replace this live LatexGroup.
  const visualTexSavedSilence = model.silenceNotifications;
  model.silenceNotifications = true;
  model.mode = "latex";
  model.position = model.offsetOf(beforeClose);
  model.silenceNotifications = visualTexSavedSilence;
  if (!visualTexSavedSilence) mathfield.visualTexDeferredSelectionAfterInput = true;
  visualTexSyncSourceWrapper(mathfield);
  model.contentDidChange({ data: shell, inputType: "insertText" });
  requestUpdate(mathfield);
  return true;
}
export function visualTexClearSourceWrapper(mathfield: _Mathfield): void {
  if (!mathfield.visualTexSourceWrapper) return;
  mathfield.visualTexSourceWrapper = null;
  visualTexSyncSourceWrapper(mathfield);
}
export function visualTexCancelSourceWrapper(mathfield: _Mathfield): boolean {
  const state = mathfield.visualTexSourceWrapper;
  const group = state?.group;
  if (!state || !group?.parent) { visualTexClearSourceWrapper(mathfield); return false; }
  const model = mathfield.model;
  const parent = group.parent;
  const branch = group.parentBranch;
  const left = group.leftSibling;
  parent.removeChild(group);
  model.mode = "math";
  model.position = Math.max(0, model.offsetOf(left));
  visualTexClearSourceWrapper(mathfield);
  if (state.structuralSlot?.parent &&
      !visualTexStructuralSlotStillOccupied(state.structuralSlot.parent, state.structuralSlot.branch, state.structuralSlot)) {
    visualTexRestoreStructuralPlaceholder(model, state.structuralSlot.parent, state.structuralSlot.branch, state.structuralSlot.after, true, state.structuralSlot);
  } else if (visualTexShouldRetainStructuralPlaceholder(parent, branch) && parent.hasEmptyBranch(branch)) {
    visualTexRestoreStructuralPlaceholder(model, parent, branch, left);
  }
  requestUpdate(mathfield);
  return true;
}
export function visualTexSourceWrapperSimpleCommand(mathfield: _Mathfield, content: string): boolean {
  if (!/^\\[a-zA-Z]+\*?$/.test(content)) return false;
  const native = LATEX_COMMANDS[content];
  if (MATH_SYMBOLS[content]) return true;
  if (native && !(native.params || []).some((param) => !(param as any).optional)) return true;
  const macro = mathfield.options.macros?.[content.slice(1)];
  return typeof macro === "string" && !/#\d/.test(macro);
}
export function visualTexCommitSourceWrapper(
  mathfield: _Mathfield,
  { requireComplete = false, notify = true } = {}
): boolean {
  const state = mathfield.visualTexSourceWrapper;
  const group = state?.group;
  if (!state || !group?.parent) { visualTexClearSourceWrapper(mathfield); return false; }
  const content = visualTexSourceWrapperContent(state);
  if (!content) return false;
  if (requireComplete && Array.from(content).length !== 1 && !visualTexSourceWrapperSimpleCommand(mathfield, content)) return false;
  const model = mathfield.model;
  let command = state.command;
  if (command === "\\mathcal" && /^[a-z]$/.test(content)) command = "\\mathscr";
  const latex = command + "{" + content + "}";
  const parsed = parseLatex(latex, { context: mathfield.context });
  if (!parsed?.length) return false;
  const parent = group.parent;
  const left = group.leftSibling;
  parent.removeChild(group);
  const last = parent.addChildrenAfter(parsed, left);
  if (state.structuralSlot) {
    for (const atom of parsed) atom.visualTexStructuralSlot = state.structuralSlot;
  }
  const visualTexSavedSilence = model.silenceNotifications;
  model.silenceNotifications = true;
  model.mode = "math";
  model.position = model.offsetOf(last);
  model.silenceNotifications = visualTexSavedSilence;
  visualTexClearSourceWrapper(mathfield);
  if (notify && !visualTexSavedSilence)
    mathfield.visualTexDeferredSelectionAfterInput = true;
  mathfield.stopCoalescingUndo();
  requestUpdate(mathfield);
  if (notify) model.contentDidChange({ data: latex, inputType: "insertText" });
  return true;
}
export function visualTexMaybeAutoCommitSourceWrapper(mathfield: _Mathfield): boolean {
  const state = mathfield.visualTexSourceWrapper;
  if (!state || mathfield.visualTexCompositionCommit || !visualTexSourceWrapperCaretIsInside(mathfield, state)) return false;
  if (!visualTexSourceWrapperAutoExitEnabled(mathfield, state)) { visualTexSyncSourceWrapper(mathfield); return false; }
  const content = visualTexSourceWrapperContent(state);
  visualTexSyncSourceWrapper(mathfield);
  if (content == null || Array.from(content).length !== 1 || content === "\\") return false;
  return visualTexCommitSourceWrapper(mathfield, { requireComplete: true, notify: false });
}
export function visualTexAcceptSourceWrapperCompletion(mathfield: _Mathfield): boolean {
  const state = mathfield.visualTexSourceWrapper;
  const completion = mathfield.visualTexCompletion;
  if (!state?.group?.parent || completion?.group !== state.group) return false;
  const selected = completion!.suggestions?.[mathfield.suggestionIndex];
  if (!selected) return false;
  const info = visualTexCompletionInfo(mathfield, selected);
  const token = (info?.command || selected).match(/^\\[a-zA-Z]+\*?/)?.[0];
  if (!token) return false;
  const model = mathfield.model;
  let atom = model.at(model.position);
  while (atom instanceof LatexAtom && /^[a-zA-Z\*]$/.test(atom.value)) atom = atom.leftSibling;
  if (!(atom instanceof LatexAtom) || atom.value !== "\\" || atom.parent !== state.group) return false;
  const first = atom;
  const commandAtoms = [first];
  atom = first.rightSibling;
  while (atom instanceof LatexAtom && /^[a-zA-Z\*]$/.test(atom.value)) { commandAtoms.push(atom); atom = atom.rightSibling; }
  const left = first.leftSibling;
  if (!left) return false;
  for (const item of commandAtoms) state.group.removeChild(item);
  const visualTexTemplate = selected.startsWith(token) ? selected : token;
  const visualTexSource = visualTexTemplate.replace(/\\placeholder\{\}/g, "");
  const replacement = [...visualTexSource].map((value) => new LatexAtom(value));
  const last = state.group.addChildrenAfter(replacement, left);
  const visualTexEmptyArgument = visualTexSource.indexOf("{}");
  const visualTexCaretAtom = visualTexEmptyArgument >= 0
    ? replacement[visualTexEmptyArgument]
    : last;
  model.position = model.offsetOf(visualTexCaretAtom);
  mathfield.visualTexCompletion = null;
  hideSuggestionPopover(mathfield);
  visualTexSyncSourceWrapper(mathfield);
  model.contentDidChange({ data: visualTexSource, inputType: "insertText" });
  requestUpdate(mathfield);
  return true;
}
export function visualTexAutoExitAccentAfterDirectInput(mathfield: _Mathfield, atom: Atom | undefined): boolean {
  if (mathfield.visualTexCompositionCommit) return false;
  if ((mathfield.host == null ? void 0 : mathfield.host.dataset.visualtexAutoExitAccent) !== "true") return false;
  if (atom?.type === "accent" && visualTexIsAccentSourceCommand(atom.command)) {
    const previousPosition = mathfield.model.position;
    const target = mathfield.model.offsetOf(atom);
    if (target >= 0 && target !== previousPosition) {
      mathfield.model.position = target;
      mathfield.stopCoalescingUndo();
      mathfield.model.announce("move", previousPosition);
      return true;
    }
  }
  let current = atom;
  while (current?.parent) {
    if (current.parent.type === "accent") {
      if (!visualTexIsAccentSourceCommand(current.parent.command)) return false;
      if (current !== atom) return false;
      const previousPosition = mathfield.model.position;
      const target = mathfield.model.offsetOf(current.parent);
      if (target < 0 || target === previousPosition) return false;
      mathfield.model.position = target;
      mathfield.stopCoalescingUndo();
      mathfield.model.announce("move", previousPosition);
      return true;
    }
    current = current.parent;
  }
  return false;
}
export function visualTexAutoExitAfterComposition(
  mathfield: _Mathfield,
  slot: VisualTexStructuralSlot | null | undefined
): boolean {
  if (!slot?.parent) return false;
  const model = mathfield.model;
  const parent = slot.parent;
  const branch = slot.branch;
  if (parent.type === "accent" && branch === "body" &&
      (mathfield.host == null ? void 0 : mathfield.host.dataset.visualtexAutoExitAccent) === "true" &&
      visualTexIsAccentSourceCommand(parent.command)) {
    const previous = model.position;
    const target = model.offsetOf(parent);
    if (target >= 0 && target !== previous) {
      model.position = target;
      mathfield.stopCoalescingUndo();
      model.announce("move", previous);
      return true;
    }
  }
  if (branch !== "superscript" && branch !== "subscript") return false;
  if (visualTexIsBoundedOperatorParent(parent)) {
    if (branch === "subscript") {
      const upper = visualTexBoundedOperatorBranchPlaceholders(model, parent, "superscript");
      return upper[0] ? leapTo(model, upper[0]) : false;
    }
    const operatorOffset = model.offsetOf(parent);
    const target = leapTarget(model, operatorOffset + 1, "forward");
    if (target && model.offsetOf(target) > operatorOffset) return leapTo(model, target);
    return leapTo(model, parent);
  }
  const enabled = branch === "superscript"
    ? (mathfield.host == null ? void 0 : mathfield.host.dataset.visualtexAutoExitSuperscript) === "true"
    : (mathfield.host == null ? void 0 : mathfield.host.dataset.visualtexAutoExitSubscript) === "true";
  if (!enabled || !parent.parent) return false;
  const previous = model.position;
  const target = model.offsetOf(parent);
  if (target < 0 || target === previous) return false;
  model.position = target;
  mathfield.stopCoalescingUndo();
  model.announce("move", previous);
  return true;
}
export function visualTexScriptAutoExitKey(mathfield: _Mathfield, atom: Atom | undefined): string | null {
  const branch = atom?.parentBranch;
  if (branch !== "superscript" && branch !== "subscript") return null;
  const parent = atom?.parent;
  if (!parent) return null;
  if (visualTexIsBoundedOperatorParent(parent)) return null;
  const enabled = branch === "superscript"
    ? (mathfield.host == null ? void 0 : mathfield.host.dataset.visualtexAutoExitSuperscript) === "true"
    : (mathfield.host == null ? void 0 : mathfield.host.dataset.visualtexAutoExitSubscript) === "true";
  if (!enabled) return null;
  const branchAtoms = parent.branch(branch);
  const branchId = (branchAtoms == null ? void 0 : branchAtoms[0] == null ? void 0 : branchAtoms[0].id) || parent.id;
  return branchId ? `${branch}:${branchId}` : null;
}
export function visualTexAutoExitScriptAfterDirectInput(mathfield: _Mathfield, atom: Atom | undefined): boolean {
  if (mathfield.visualTexCompositionCommit) return false;
  const key = visualTexScriptAutoExitKey(mathfield, atom);
  if (!key) return false;
  let consumed = mathfield.visualTexConsumedScriptAutoExitKeys;
  if (!consumed) {
    consumed = /* @__PURE__ */ new Set();
    mathfield.visualTexConsumedScriptAutoExitKeys = consumed;
  }
  if (consumed.has(key)) return false;
  consumed.add(key);
  const parent = atom?.parent;
  if (!parent || !parent.parent) return false;
  const model = mathfield.model;
  const previousPosition = model.position;
  // Ctrl+J builds base_{#?}^{#?}: fill the still-empty sibling script
  // first, like the bounded-operator limits above, instead of leaving it.
  const siblingBranch = atom.parentBranch === "subscript" ? "superscript" : "subscript";
  const siblingAtoms = (parent.branch(siblingBranch) || []).filter((item) => item.type !== "first");
  if (siblingAtoms.length === 1 && siblingAtoms[0].type === "placeholder") {
    const placeholderOffset = model.offsetOf(siblingAtoms[0]);
    model.setSelection(placeholderOffset - 1, placeholderOffset);
    model.mathfield.stopCoalescingUndo();
    model.announce("move", previousPosition);
    return true;
  }
  const targetPosition = model.offsetOf(parent);
  if (targetPosition < 0 || targetPosition === previousPosition) return false;
  model.position = targetPosition;
  model.mathfield.stopCoalescingUndo();
  model.announce("move", previousPosition);
  return true;
}

function insertMathModeChar(mathfield: _Mathfield, c: string): void {
  const model = mathfield.model;
  if (!/\d/.test(c)) formatScientificNotationIfApplicable(mathfield);
  // VisualTeX editing core: Space leaves the structure (a macro as a whole) or,
  // at the top level, inserts a visible space (navigation.ts).
  const selector = {
    "^": "moveToSuperscript",
    "_": "moveToSubscript",
    " ": mathfield.options.mathModeSpace ? ["insert", mathfield.options.mathModeSpace] : "visualTexSpace"
  }[c];
  if (selector === "visualTexSpace") {
    const child = model.at(Math.max(model.position, model.anchor));
    const { parent } = child;
    if (parent instanceof LeftRightAtom && child.isLastSibling && mathfield.options.smartFence && parent.rightDelim === "?") {
      mathfield.snapshot();
      parent.rightDelim = parent.matchingRightDelim();
      parent.isDirty = true;
      model.contentDidChange({ inputType: "insertText" });
      mathfield.snapshot("insert-fence");
    }
  }
  if (selector) {
    mathfield.executeCommand(selector as any);
    return;
  }
  const style = { ...computeInsertStyle(mathfield) };
  if (!/[a-zA-Z0-9]/.test(c) && mathfield.styleBias !== "none") {
    style.variant = "normal";
    style.variantStyle = void 0;
  }
  const atom = model.at(model.position);
  const wasPlaceholderSelected = model.selectionIsPlaceholder;
  const visualTexInheritedStructuralSlot = visualTexLiveStructuralSlot(model.at(model.position));
  const pendingPlaceholder = mathfield.visualTexPendingPlaceholderSelection;
  mathfield.visualTexPendingPlaceholderSelection = null;
  const replacedPlaceholder = wasPlaceholderSelected || pendingPlaceholder?.position === model.position;
  // A font command such as \mathbb is represented by style on its
  // placeholder atom. Replacing that atom must carry the style to the
  // typed glyph; otherwise the whole variant silently disappears.
  if (replacedPlaceholder && /[a-zA-Z0-9]/.test(c)) {
    const selectedPlaceholder = wasPlaceholderSelected ? model.at(Math.max(model.position, model.anchor)) : null;
    Object.assign(style, selectedPlaceholder?.style || pendingPlaceholder?.style || {});
  }
  if (/\d/.test(c) && mathfield.options.smartSuperscript && atom.parentBranch === "superscript" && atom.parent!.type !== "mop" && atom.parent!.type !== "operator" && atom.parent!.type !== "extensible-symbol" && atom.hasNoSiblings) {
    if (!ModeEditor.insert(model, c, {
      style,
      insertionMode: "replaceSelection",
      selectionMode: wasPlaceholderSelected ? "after" : "placeholder"
    })) {
      mathfield.undoManager.pop();
      return;
    }
    mathfield.snapshot("insert-mord");
    moveAfterParent(model);
    return;
  }
  let input = c;
  if (input === "{") input = "\\lbrace";
  else if (input === "}") input = "\\rbrace";
  else if (input === "&") input = "\\&";
  else if (input === "#") input = "\\#";
  else if (input === "$") input = "\\$";
  else if (input === "%") input = "\\%";
  else if (input === "~") input = "\\~";
  else if (input === "\\") input = "\\backslash";
  if (!ModeEditor.insert(model, input, {
    style,
    insertionMode: "replaceSelection",
    selectionMode: wasPlaceholderSelected ? "after" : "placeholder"
  }))
    return;
  const visualTexInsertedAtom = model.at(model.position);
  if (visualTexInsertedAtom && (pendingPlaceholder?.slot || visualTexInheritedStructuralSlot)) {
    visualTexInsertedAtom.visualTexStructuralSlot = pendingPlaceholder?.slot || visualTexInheritedStructuralSlot;
  }
  mathfield.snapshot(`insert-${visualTexInsertedAtom.type}`);
  if (!visualTexAutoExitAccentAfterDirectInput(mathfield, visualTexInsertedAtom))
    visualTexAutoExitScriptAfterDirectInput(mathfield, visualTexInsertedAtom);
  if (/\d/.test(c)) {
    const timeoutValue = mathfield.options.inlineShortcutTimeout > 0 ? mathfield.options.inlineShortcutTimeout : 1e3;
    clearTimeout(mathfield.scientificNotationTimer);
    mathfield.scientificNotationTimer = setTimeout(() => {
      formatScientificNotationIfApplicable(mathfield);
    }, timeoutValue);
  }
}

export function getSelectionStyle(model: _Model): Readonly<Style> {
  // When the selection is collapsed, we inherit the style from the
  // preceding atom
  if (model.selectionIsCollapsed) return model.at(model.position)?.style ?? {};

  // Otherwise pick the style of the first (leftmost) atom **in** the
  // selection. This is a behavior consistent with text editors such as
  // TextEdit
  const first = range(model.selection)[0];
  return model.at(first + 1)?.style ?? {};
}

/**
 * Insert a smart fence '(', '{', '[', etc...
 * If not handled (because `key` was not a fence), return false.
 */
function insertSmartFence(model: _Model, key: string, style?: Style): boolean {
  if (!key) return false;
  if (model.mode !== "math") return false;
  const atom = model.at(model.position);
  const { parent } = atom;
  // VisualTeX: "|" is not a fence key; it types one bar
  const fence = {
    "(": "(",
    ")": ")",
    "{": "\\lbrace",
    "}": "\\rbrace",
    "[": "\\lbrack",
    "]": "\\rbrack"
  }[key];
  if (!fence) return false;
  const lDelim = LEFT_DELIM[fence];
  const rDelim = RIGHT_DELIM[fence];
  if (!model.selectionIsCollapsed) {
    model.mathfield.snapshot();
    const [start, end] = range(model.selection);
    let body = model.extractAtoms([start, end]);
    body = body.filter((a) => a.type !== "first");
    const atom2 = parent!.addChildrenAfter(
      [
        new LeftRightAtom("left...right", body, {
          leftDelim: fence,
          rightDelim: rDelim
        })
      ],
      model.at(start)
    );
    model.setSelection(
      model.offsetOf(atom2.firstChild),
      model.offsetOf(atom2.lastChild)
    );
    model.mathfield.snapshot("insert-fence");
    model.contentDidChange({ data: fence, inputType: "insertText" });
    return true;
  }
  if (rDelim) {
    if (parent instanceof LeftRightAtom && parent.firstChild === atom && // At first child
    (parent.leftDelim === "?" || parent.leftDelim === ".")) {
      parent.leftDelim = fence;
      parent.isDirty = true;
      model.mathfield.snapshot();
      model.contentDidChange({ data: fence, inputType: "insertText" });
      model.mathfield.snapshot("insert-fence");
      return true;
    }
    if (!(parent instanceof LeftRightAtom)) {
      let sibling = atom;
      while (sibling) {
        if (sibling.type === "mclose" && sibling.value === rDelim) break;
        sibling = sibling.rightSibling;
      }
      if (sibling) {
        model.mathfield.snapshot();
        let body = model.extractAtoms([
          model.offsetOf(atom),
          model.offsetOf(sibling)
        ]);
        body.pop();
        body = body.filter((a) => a.type !== "first");
        const newLeftRight = new LeftRightAtom("left...right", body, {
          leftDelim: fence,
          rightDelim: rDelim
        });
        parent!.addChildrenAfter([newLeftRight], atom);
        model.position = model.offsetOf(newLeftRight.firstChild);
        model.contentDidChange({ data: fence, inputType: "insertText" });
        model.mathfield.snapshot("insert-fence");
        return true;
      }
    }
    const lastSibling = model.offsetOf(atom.lastSibling);
    let i;
    for (i = model.position; i <= lastSibling; i++) {
      const atom2 = model.at(i);
      if (atom2 instanceof LeftRightAtom && (atom2.leftDelim === "?" || atom2.leftDelim === ".") && isValidOpen(fence, atom2.rightDelim))
        break;
    }
    const match = model.at(i);
    if (i <= lastSibling && match instanceof LeftRightAtom) {
      match.leftDelim = fence;
      model.mathfield.snapshot();
      let extractedAtoms = model.extractAtoms([model.position, i - 1]);
      extractedAtoms = extractedAtoms.filter((value) => value.type !== "first");
      match.addChildren(extractedAtoms, match.parentBranch!);
      model.position += 1;
      model.contentDidChange({ data: fence, inputType: "insertText" });
      model.mathfield.snapshot("insert-fence");
      return true;
    }
    if (parent instanceof LeftRightAtom && (parent.leftDelim === "?" || parent.leftDelim === ".") && isValidOpen(fence, parent.rightDelim)) {
      parent.isDirty = true;
      parent.leftDelim = fence;
      model.mathfield.snapshot();
      const extractedAtoms = model.extractAtoms([
        model.offsetOf(atom.firstSibling),
        model.position
      ]);
      for (const extractedAtom of extractedAtoms)
        parent.parent!.addChildBefore(extractedAtom, parent);
      model.contentDidChange({ data: fence, inputType: "insertText" });
      model.mathfield.snapshot("insert-fence");
      return true;
    }
    if (!(parent instanceof LeftRightAtom && parent.leftDelim === "|")) {
      if (fence === "(") {
        let i2 = model.position - 1;
        let hasDecimalPoint = false;
        while (i2 >= 0) {
          const atom2 = model.at(i2);
          if (atom2.type === "first") break;
          if (atom2.type === "mord" && atom2.value && /^[\d]$/.test(atom2.value)) {
            i2 -= 1;
            continue;
          }
          if (atom2.type === "group" && (atom2.body?.length) === 2 && atom2.body[0].type === "first" && atom2.body[1].value === ",") {
            hasDecimalPoint = true;
            break;
          }
          if (atom2.type === "mord" && (atom2.value === "," || atom2.value === ".")) {
            hasDecimalPoint = true;
            break;
          }
          break;
        }
        if (hasDecimalPoint) return false;
      }
      model.mathfield.snapshot();
      ModeEditor.insert(model, `\\left${fence}\\right?`, {
        format: "latex",
        style
      });
      // the new fence takes the rest of the slot and the caret goes to its
      // start (VisualTeX: the insertion leaves the caret in its empty body)
      const newFence = atom.rightSibling;
      // (setting the body, even to nothing, also drops the verbatim \right?)
      if (newFence instanceof LeftRightAtom) {
        newFence.body = newFence.lastSibling !== newFence ? model.extractAtoms([model.offsetOf(newFence), model.offsetOf(newFence.lastSibling)]) : [];
        model.position = model.offsetOf(newFence.firstChild);
      }
      model.mathfield.snapshot("insert-fence");
      return true;
    }
  }
  if (lDelim) {
    if (fence === ")") {
      let i2 = model.position - 1;
      let hasDigits = false;
      while (i2 >= 0) {
        const atom2 = model.at(i2);
        if (atom2.type === "first") break;
        if (atom2.type === "mord" && atom2.value && /^[\d]$/.test(atom2.value)) {
          hasDigits = true;
          i2 -= 1;
          continue;
        }
        break;
      }
      if (hasDigits && model.at(i2).type === "mopen" && model.at(i2).value === "(")
        return false;
    }
    let sibling = atom;
    while (sibling) {
      if (sibling.type === "mopen" && sibling.value === lDelim) {
        model.mathfield.snapshot();
        const insertAfter = sibling.leftSibling;
        let body = model.extractAtoms([
          model.offsetOf(sibling.leftSibling),
          model.offsetOf(atom)
        ]);
        [, ...body] = body;
        const result = new LeftRightAtom("left...right", body, {
          leftDelim: lDelim,
          rightDelim: fence
        });
        parent!.addChildrenAfter([result], insertAfter);
        model.position = model.offsetOf(result);
        model.contentDidChange({ data: fence, inputType: "insertText" });
        model.mathfield.snapshot("insert-fence");
        return true;
      }
      sibling = sibling.leftSibling;
    }
    if (parent instanceof LeftRightAtom && atom.isLastSibling && isValidClose(parent.leftDelim, fence)) {
      model.mathfield.snapshot();
      parent.isDirty = true;
      parent.rightDelim = fence;
      model.position += 1;
      model.contentDidChange({ data: fence, inputType: "insertText" });
      model.mathfield.snapshot("insert-fence");
      return true;
    }
    const firstSibling = model.offsetOf(atom.firstSibling);
    let i;
    for (i = model.position; i >= firstSibling; i--) {
      const atom2 = model.at(i);
      if (atom2 instanceof LeftRightAtom && (atom2.rightDelim === "?" || atom2.rightDelim === ".") && isValidClose(atom2.leftDelim, fence))
        break;
    }
    const match = model.at(i);
    if (i >= firstSibling && match instanceof LeftRightAtom) {
      model.mathfield.snapshot();
      match.rightDelim = fence;
      match.addChildren(
        model.extractAtoms([i, model.position]),
        match.parentBranch!
      );
      model.contentDidChange({ data: fence, inputType: "insertText" });
      model.mathfield.snapshot("insert-fence");
      return true;
    }
    if (parent instanceof LeftRightAtom && (parent.rightDelim === "?" || parent.rightDelim === ".") && isValidClose(parent.leftDelim, fence)) {
      model.mathfield.snapshot();
      parent.isDirty = true;
      parent.rightDelim = fence;
      parent.parent!.addChildren(
        model.extractAtoms([model.position, model.offsetOf(atom.lastSibling)]),
        parent.parentBranch!
      );
      model.position = model.offsetOf(parent);
      model.contentDidChange({ data: fence, inputType: "insertText" });
      model.mathfield.snapshot("insert-fence");
      return true;
    }
    const grandparent = parent!.parent;
    if (grandparent instanceof LeftRightAtom && (grandparent.rightDelim === "?" || grandparent.rightDelim === ".") && model.at(model.position).isLastSibling) {
      model.position = model.offsetOf(grandparent);
      return insertSmartFence(model, fence, style);
    }
    return false;
  }
  return false;
}

function isValidClose(open: string | undefined, close: string): boolean {
  if (!open) return true;

  if (
    ['(', '\\lparen', '{', '\\{', '\\lbrace', '[', '\\lbrack'].includes(open)
  ) {
    return [')', '\\rparen', '}', '\\}', '\\rbrace', ']', '\\rbrack'].includes(
      close
    );
  }
  return RIGHT_DELIM[open] === close;
}

function isValidOpen(open: string, close: string | undefined): boolean {
  if (!close) return true;

  if (
    [')', '\\rparen', '}', '\\}', '\\rbrace', ']', '\\rbrack'].includes(close)
  ) {
    return ['(', '\\lparen', '{', '\\{', '\\lbrace', '[', '\\lbrack'].includes(
      open
    );
  }
  return LEFT_DELIM[close] === open;
}

visualTexLateFunctions.visualTexMaybeAutoCommitSourceWrapper = visualTexMaybeAutoCommitSourceWrapper;
visualTexLateFunctions.visualTexCancelSourceWrapper = visualTexCancelSourceWrapper;
visualTexLateFunctions.visualTexSourceWrapperKind = visualTexSourceWrapperKind;
visualTexLateFunctions.visualTexCanonicalVariantInputCommand = visualTexCanonicalVariantInputCommand;
visualTexLateFunctions.visualTexStartSourceWrapper = visualTexStartSourceWrapper;
visualTexLateFunctions.visualTexSyncSourceWrapper = visualTexSyncSourceWrapper;
