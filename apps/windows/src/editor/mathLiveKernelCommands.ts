import type { MathfieldElement, Selection, Selector } from "mathlive";
import type { CommandUsage } from "../types/command";

export interface MathLivePersistentTypingStyle {
  bold: boolean;
  italic: boolean;
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

// These commands are registered by vite.mathliveEditingKernel.ts inside MathLive.
// Keep the extension cast here; callers cannot pass arbitrary command names.
export function toggleMathLiveSelectionStyle(
  field: MathfieldElement,
  kind: "bold" | "italic",
  selection: Selection,
) {
  return field.executeCommand("visualTexToggleSelectionStyle" as Selector, kind, selection);
}

export function insertMathLiveAlignmentPoint(field: MathfieldElement) {
  return field.executeCommand("visualTexInsertAlignmentPoint" as Selector);
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
