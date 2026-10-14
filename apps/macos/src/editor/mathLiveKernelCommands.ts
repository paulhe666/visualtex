import type { MathfieldElement, Selection, Selector } from "mathlive";
import type { CommandUsage } from "../types/command";

import type { MathLivePersistentTypingStyle } from "../../vendor/mathlive/types";
export type { MathLivePersistentTypingStyle } from "../../vendor/mathlive/types";

export function dismissMathLiveSuggestions(field: MathfieldElement) {
  field.executeCommand("visualTexDismissSuggestions" as Selector);
}

export function configureMathLiveCompletion(field: MathfieldElement, preferences: {
  usage: Record<string, CommandUsage>;
  personalize: boolean;
  count: number;
}) {
  field.visualTexCompletionPreferences = preferences;
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
  field.visualTexPersistentTypingStyle = { ...style };
}
