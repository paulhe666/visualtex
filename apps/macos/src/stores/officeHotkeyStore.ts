import { create } from "zustand";
import { createJSONStorage, persist } from "zustand/middleware";
import {
  formulaHotkeyChordId,
  formulaHotkeyHasModifier,
  protectedFormulaHotkeyAction,
  type FormulaHotkeyChord,
} from "../shortcuts/formulaHotkeys";
import {
  isOfficeHotkeyActionId,
  OFFICE_HOTKEY_STORAGE_KEY,
  officeHotkeySupportsCode,
  type OfficeHotkeyActionId,
  type OfficeHotkeyBinding,
} from "../shortcuts/officeHotkeys";
import { safeStorage } from "../runtime/safeStorage";

interface OfficeHotkeyState {
  bindings: OfficeHotkeyBinding[];
  setBinding: (actionId: OfficeHotkeyActionId, chord: FormulaHotkeyChord) => void;
  removeBinding: (actionId: OfficeHotkeyActionId) => void;
}

function normalizeChord(value: unknown): FormulaHotkeyChord | null {
  if (!value || typeof value !== "object") return null;
  const chord = value as Partial<FormulaHotkeyChord>;
  if (
    typeof chord.code !== "string" ||
    !officeHotkeySupportsCode(chord.code) ||
    typeof chord.key !== "string" ||
    typeof chord.ctrlKey !== "boolean" ||
    typeof chord.altKey !== "boolean" ||
    typeof chord.shiftKey !== "boolean" ||
    typeof chord.metaKey !== "boolean"
  ) {
    return null;
  }
  return chord as FormulaHotkeyChord;
}

export function normalizeOfficeHotkeyBindings(value: unknown) {
  if (!Array.isArray(value)) return [];
  const usedActions = new Set<OfficeHotkeyActionId>();
  const usedChords = new Set<string>();
  const normalized: OfficeHotkeyBinding[] = [];
  for (const item of value) {
    if (!item || typeof item !== "object") continue;
    const candidate = item as Partial<OfficeHotkeyBinding>;
    const chord = normalizeChord(candidate.chord);
    if (
      !isOfficeHotkeyActionId(candidate.actionId) ||
      !chord ||
      !formulaHotkeyHasModifier(chord) ||
      protectedFormulaHotkeyAction(chord)
    ) {
      continue;
    }
    const chordId = formulaHotkeyChordId(chord);
    if (usedActions.has(candidate.actionId) || usedChords.has(chordId)) continue;
    usedActions.add(candidate.actionId);
    usedChords.add(chordId);
    normalized.push({
      actionId: candidate.actionId,
      chord,
      updatedAt:
        typeof candidate.updatedAt === "number" && Number.isFinite(candidate.updatedAt)
          ? candidate.updatedAt
          : Date.now(),
    });
  }
  return normalized;
}

export const useOfficeHotkeyStore = create<OfficeHotkeyState>()(
  persist(
    (set) => ({
      bindings: [],
      setBinding: (actionId, chord) =>
        set((state) => {
          const chordId = formulaHotkeyChordId(chord);
          return {
            bindings: [
              { actionId, chord, updatedAt: Date.now() },
              ...state.bindings.filter(
                (binding) =>
                  binding.actionId !== actionId &&
                  formulaHotkeyChordId(binding.chord) !== chordId,
              ),
            ],
          };
        }),
      removeBinding: (actionId) =>
        set((state) => ({
          bindings: state.bindings.filter(
            (binding) => binding.actionId !== actionId,
          ),
        })),
    }),
    {
      name: OFFICE_HOTKEY_STORAGE_KEY,
      version: 1,
      storage: createJSONStorage(() => safeStorage),
      partialize: (state) => ({ bindings: state.bindings }),
      merge: (persistedState, currentState) => ({
        ...currentState,
        bindings: normalizeOfficeHotkeyBindings(
          (persistedState as Partial<OfficeHotkeyState>).bindings,
        ),
      }),
    },
  ),
);
