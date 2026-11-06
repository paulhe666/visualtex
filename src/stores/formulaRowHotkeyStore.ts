import { create } from "zustand";
import { createJSONStorage, persist } from "zustand/middleware";
import type { FormulaHotkeyChord } from "../shortcuts/formulaHotkeys";
import {
  createDefaultFormulaRowHotkeyBindings,
  defaultFormulaRowHotkeyChord,
  FORMULA_ROW_HOTKEY_ACTIONS,
  FORMULA_ROW_HOTKEY_STORAGE_KEY,
  formulaRowHotkeyChordId,
  formulaRowHotkeyChordProblem,
  isFormulaRowHotkeyActionId,
  type FormulaRowHotkeyActionId,
  type FormulaRowHotkeyBinding,
} from "../shortcuts/formulaRowHotkeys";
import { safeStorage } from "../runtime/safeStorage";

interface FormulaRowHotkeyState {
  bindings: FormulaRowHotkeyBinding[];
  /** Assigns chord to actionId and takes it away from any other row action. */
  setBinding: (
    actionId: FormulaRowHotkeyActionId,
    chord: FormulaHotkeyChord,
  ) => void;
  /** Turns the action's shortcut off. */
  removeBinding: (actionId: FormulaRowHotkeyActionId) => void;
  resetBinding: (actionId: FormulaRowHotkeyActionId) => void;
}

function normalizeChord(value: unknown): FormulaHotkeyChord | null {
  if (!value || typeof value !== "object") return null;
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
    return null;
  }
  const normalized = chord as FormulaHotkeyChord;
  return formulaRowHotkeyChordProblem(normalized) ? null : normalized;
}

/** One binding per action in definition order; an action missing from the
 * stored state gets its default chord unless another action already uses it. */
export function normalizeFormulaRowHotkeyBindings(
  value: unknown,
): FormulaRowHotkeyBinding[] {
  const stored = new Map<FormulaRowHotkeyActionId, FormulaRowHotkeyBinding>();
  const usedChords = new Set<string>();
  if (Array.isArray(value)) {
    for (const item of value) {
      if (!item || typeof item !== "object") continue;
      const candidate = item as Partial<FormulaRowHotkeyBinding>;
      if (!isFormulaRowHotkeyActionId(candidate.actionId)) continue;
      if (stored.has(candidate.actionId)) continue;
      const chord =
        candidate.chord === null ? null : normalizeChord(candidate.chord);
      if (candidate.chord !== null && !chord) continue;
      if (chord) {
        const chordId = formulaRowHotkeyChordId(chord);
        if (usedChords.has(chordId)) continue;
        usedChords.add(chordId);
      }
      stored.set(candidate.actionId, {
        actionId: candidate.actionId,
        chord,
        updatedAt:
          typeof candidate.updatedAt === "number" &&
          Number.isFinite(candidate.updatedAt)
            ? candidate.updatedAt
            : Date.now(),
      });
    }
  }
  return FORMULA_ROW_HOTKEY_ACTIONS.map((action) => {
    const binding = stored.get(action.id);
    if (binding) return binding;
    const chord = defaultFormulaRowHotkeyChord(action.id);
    const chordId = formulaRowHotkeyChordId(chord);
    const free = !usedChords.has(chordId);
    if (free) usedChords.add(chordId);
    return { actionId: action.id, chord: free ? chord : null, updatedAt: 0 };
  });
}

export const useFormulaRowHotkeyStore = create<FormulaRowHotkeyState>()(
  persist(
    (set) => ({
      bindings: createDefaultFormulaRowHotkeyBindings(),
      setBinding: (actionId, chord) =>
        set((state) => {
          const chordId = formulaRowHotkeyChordId(chord);
          const now = Date.now();
          return {
            bindings: state.bindings.map((binding) => {
              if (binding.actionId === actionId) {
                return { actionId, chord, updatedAt: now };
              }
              if (
                binding.chord &&
                formulaRowHotkeyChordId(binding.chord) === chordId
              ) {
                return { ...binding, chord: null, updatedAt: now };
              }
              return binding;
            }),
          };
        }),
      removeBinding: (actionId) =>
        set((state) => ({
          bindings: state.bindings.map((binding) =>
            binding.actionId === actionId
              ? { ...binding, chord: null, updatedAt: Date.now() }
              : binding,
          ),
        })),
      resetBinding: (actionId) =>
        set((state) => {
          const chord = defaultFormulaRowHotkeyChord(actionId);
          const chordId = formulaRowHotkeyChordId(chord);
          const now = Date.now();
          return {
            bindings: state.bindings.map((binding) => {
              if (binding.actionId === actionId) {
                return { actionId, chord, updatedAt: now };
              }
              if (
                binding.chord &&
                formulaRowHotkeyChordId(binding.chord) === chordId
              ) {
                return { ...binding, chord: null, updatedAt: now };
              }
              return binding;
            }),
          };
        }),
    }),
    {
      name: FORMULA_ROW_HOTKEY_STORAGE_KEY,
      version: 1,
      storage: createJSONStorage(() => safeStorage),
      partialize: (state) => ({ bindings: state.bindings }),
      merge: (persistedState, currentState) => ({
        ...currentState,
        bindings: normalizeFormulaRowHotkeyBindings(
          (persistedState as Partial<FormulaRowHotkeyState> | undefined)
            ?.bindings,
        ),
      }),
    },
  ),
);
