import type { MathfieldElement, Selection, Selector } from "mathlive";
import type { CommandUsage } from "../types/command";

export interface MathLivePersistentTypingStyle {
  bold: boolean;
  // null follows MathLive's normal math alphabet; false forces upright.
  italic: boolean | null;
  color: string | null;
  backgroundColor: string | null;
}

export function dismissMathLiveSuggestions(field: MathfieldElement) {
  field.executeCommand("visualTexDismissSuggestions" as Selector);
}

export function configureMathLiveCompletion(field: MathfieldElement, preferences: {
  usage: Record<string, CommandUsage>;
  personalize: boolean;
  count: number;
}) {
  (field as MathfieldElement & { visualTexCompletionPreferences?: typeof preferences })
    .visualTexCompletionPreferences = preferences;
}

// These commands are implemented by the owned MathLive kernel.
// Keep the extension cast here; callers cannot pass arbitrary command names.
export function toggleMathLiveSelectionStyle(
  field: MathfieldElement,
  kind: "bold" | "italic",
  selection: Selection,
) {
  return field.executeCommand("visualTexToggleSelectionStyle" as Selector, kind, selection);
}

// A selection that covers only the ^/_ attachment of c^2 paints a glyph-less
// container; move that paint onto the script content (c^{\textcolor{..}{2}}).
export function paintMathLiveDetachedScripts(
  field: MathfieldElement,
  selection: Selection,
) {
  return field.executeCommand("visualTexPaintDetachedScripts" as Selector, selection);
}

export function moveMathLiveToEmptySiblingScript(field: MathfieldElement) {
  return field.executeCommand("visualTexMoveToEmptySiblingScript" as Selector);
}

export function insertMathLiveAlignmentPoint(field: MathfieldElement) {
  return field.executeCommand("visualTexInsertAlignmentPoint" as Selector);
}

/**
 * After the formula was replaced: the caret where its atom went, or after the
 * rewrite that replaced it (the editing core's restore; take the snapshot with
 * visualTexEditingCore.captureCaret before replacing).
 */
export function restoreMathLiveCaret(field: MathfieldElement, snapshot: unknown) {
  return field.executeCommand(["visualTexRestoreCaret", snapshot] as unknown as Selector);
}

/** Escape with a selection: collapse it to its end (the editing core). */
export function collapseMathLiveSelectionToEnd(field: MathfieldElement) {
  return field.executeCommand("visualTexCollapseToEnd" as Selector);
}

/** & in a matrix cell: the cell on the right, or a new column at row end. */
export function moveMathLiveToNextColumn(field: MathfieldElement) {
  return field.executeCommand("visualTexNextColumn" as Selector);
}

export function insertMathLiveRowBreak(
  field: MathfieldElement,
  environment: "aligned" | "gathered",
) {
  return field.executeCommand("visualTexInsertRowBreak" as Selector, environment);
}

export function setMathLivePersistentTypingStyle(
  field: MathfieldElement,
  style: MathLivePersistentTypingStyle,
) {
  const next = {
    bold: style.bold,
    italic: style.italic,
    ...(style.color ? { color: style.color } : {}),
    ...(style.backgroundColor ? { backgroundColor: style.backgroundColor } : {}),
  };
  const target = field as unknown as {
    visualTexPersistentTypingStyle?: typeof next;
    _mathfield?: {
      visualTexPersistentTypingStyle?: typeof next;
    };
  };
  // model.mathfield is MathLive's private controller, not the custom element.
  // Mirror the state on both so the insertion kernel sees it while keeping the
  // public element inspectable in browser regressions.
  target.visualTexPersistentTypingStyle = next;
  if (target._mathfield) {
    target._mathfield.visualTexPersistentTypingStyle = next;
  }
}
