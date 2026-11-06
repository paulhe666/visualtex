import { create } from "zustand";
import { createJSONStorage, persist } from "zustand/middleware";
import type { FormulaHotkeyChord } from "../shortcuts/formulaHotkeys";
import {
  CONFIRM_ESCAPE_KEY_STORAGE_KEY,
  confirmEscapeChordProblem,
  DEFAULT_CONFIRM_ESCAPE_CHORD,
} from "../shortcuts/confirmEscapeKey";
import { safeStorage } from "../runtime/safeStorage";

interface ConfirmEscapeKeyState {
  /** null: no key confirms a quick escape. */
  chord: FormulaHotkeyChord | null;
  setChord: (chord: FormulaHotkeyChord | null) => void;
  reset: () => void;
}

export function normalizeConfirmEscapeChord(
  value: unknown,
): FormulaHotkeyChord | null {
  if (value === null) return null;
  if (!value || typeof value !== "object") return DEFAULT_CONFIRM_ESCAPE_CHORD;
  const chord = value as Partial<FormulaHotkeyChord>;
  if (
    typeof chord.code !== "string" ||
    !chord.code ||
    typeof chord.key !== "string" ||
    typeof chord.ctrlKey !== "boolean" ||
    typeof chord.altKey !== "boolean" ||
    typeof chord.shiftKey !== "boolean" ||
    typeof chord.metaKey !== "boolean"
  ) {
    return DEFAULT_CONFIRM_ESCAPE_CHORD;
  }
  const normalized = chord as FormulaHotkeyChord;
  return confirmEscapeChordProblem(normalized)
    ? DEFAULT_CONFIRM_ESCAPE_CHORD
    : normalized;
}

export const useConfirmEscapeKeyStore = create<ConfirmEscapeKeyState>()(
  persist(
    (set) => ({
      chord: DEFAULT_CONFIRM_ESCAPE_CHORD,
      setChord: (chord) => set({ chord }),
      reset: () => set({ chord: DEFAULT_CONFIRM_ESCAPE_CHORD }),
    }),
    {
      name: CONFIRM_ESCAPE_KEY_STORAGE_KEY,
      version: 1,
      storage: createJSONStorage(() => safeStorage),
      partialize: (state) => ({ chord: state.chord }),
      merge: (persistedState, currentState) => {
        const stored = persistedState as
          | Partial<ConfirmEscapeKeyState>
          | undefined;
        return {
          ...currentState,
          chord:
            stored && "chord" in stored
              ? normalizeConfirmEscapeChord(stored.chord)
              : DEFAULT_CONFIRM_ESCAPE_CHORD,
        };
      },
    },
  ),
);
