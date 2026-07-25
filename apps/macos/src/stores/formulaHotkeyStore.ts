import { create } from "zustand";
import { createJSONStorage, persist } from "zustand/middleware";
import {
  formulaHotkeyChordId,
  formulaHotkeyHasModifier,
  protectedFormulaHotkeyAction,
  type FormulaHotkeyBinding,
  type FormulaHotkeyChord,
  type FormulaHotkeyTarget,
  type FormulaHotkeyTargetKind,
} from "../shortcuts/formulaHotkeys";
import type { LatexCommand } from "../types/command";

interface FormulaHotkeyState {
  bindings: FormulaHotkeyBinding[];
  setBinding: (target: FormulaHotkeyTarget, chord: FormulaHotkeyChord) => void;
  removeBinding: (bindingId: string) => void;
  removeBindingsForTarget: (targetId: string) => void;
  clearBindings: () => void;
}

const targetKinds = new Set<FormulaHotkeyTargetKind>([
  "command",
  "common-tile",
  "custom-tile",
  "matrix",
]);

function normalizeCommand(value: unknown): LatexCommand | null {
  if (!value || typeof value !== "object") return null;
  const command = value as Partial<LatexCommand>;
  if (
    typeof command.id !== "string" ||
    typeof command.command !== "string" ||
    typeof command.insertTemplate !== "string" ||
    typeof command.previewLatex !== "string" ||
    typeof command.labelZh !== "string" ||
    typeof command.labelEn !== "string" ||
    !Array.isArray(command.aliases) ||
    !Array.isArray(command.keywords) ||
    typeof command.category !== "string" ||
    typeof command.defaultPriority !== "number" ||
    typeof command.supportedInMathMode !== "boolean"
  ) {
    return null;
  }
  return command as LatexCommand;
}

function normalizeTarget(value: unknown): FormulaHotkeyTarget | null {
  if (!value || typeof value !== "object") return null;
  const target = value as Partial<FormulaHotkeyTarget>;
  const command = normalizeCommand(target.command);
  if (
    typeof target.id !== "string" ||
    !target.id.trim() ||
    typeof target.kind !== "string" ||
    !targetKinds.has(target.kind as FormulaHotkeyTargetKind) ||
    !command
  ) {
    return null;
  }
  return {
    id: target.id,
    kind: target.kind as FormulaHotkeyTargetKind,
    command,
  };
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
  return chord as FormulaHotkeyChord;
}

function normalizeBindings(value: unknown) {
  if (!Array.isArray(value)) return [];
  const usedTargets = new Set<string>();
  const usedChords = new Set<string>();
  const normalized: FormulaHotkeyBinding[] = [];

  for (const item of value) {
    if (!item || typeof item !== "object") continue;
    const candidate = item as Partial<FormulaHotkeyBinding>;
    const target = normalizeTarget(candidate.target);
    const chord = normalizeChord(candidate.chord);
    if (
      !target ||
      !chord ||
      !formulaHotkeyHasModifier(chord) ||
      protectedFormulaHotkeyAction(chord)
    ) {
      continue;
    }
    const chordId = formulaHotkeyChordId(chord);
    if (usedTargets.has(target.id) || usedChords.has(chordId)) continue;
    usedTargets.add(target.id);
    usedChords.add(chordId);
    normalized.push({
      id:
        typeof candidate.id === "string" && candidate.id
          ? candidate.id
          : crypto.randomUUID(),
      target,
      chord,
      updatedAt:
        typeof candidate.updatedAt === "number" &&
        Number.isFinite(candidate.updatedAt)
          ? candidate.updatedAt
          : Date.now(),
    });
  }
  return normalized;
}

export const useFormulaHotkeyStore = create<FormulaHotkeyState>()(
  persist(
    (set) => ({
      bindings: [],
      setBinding: (target, chord) =>
        set((state) => {
          const chordId = formulaHotkeyChordId(chord);
          const existing = state.bindings.find(
            (binding) => binding.target.id === target.id,
          );
          const next: FormulaHotkeyBinding = {
            id: existing?.id ?? crypto.randomUUID(),
            target,
            chord,
            updatedAt: Date.now(),
          };
          return {
            bindings: [
              next,
              ...state.bindings.filter(
                (binding) =>
                  binding.target.id !== target.id &&
                  formulaHotkeyChordId(binding.chord) !== chordId,
              ),
            ],
          };
        }),
      removeBinding: (bindingId) =>
        set((state) => ({
          bindings: state.bindings.filter(
            (binding) => binding.id !== bindingId,
          ),
        })),
      removeBindingsForTarget: (targetId) =>
        set((state) => ({
          bindings: state.bindings.filter(
            (binding) => binding.target.id !== targetId,
          ),
        })),
      clearBindings: () => set({ bindings: [] }),
    }),
    {
      name: "visualtex-formula-hotkeys-v1",
      version: 1,
      storage: createJSONStorage(() => localStorage),
      partialize: (state) => ({ bindings: state.bindings }),
      merge: (persistedState, currentState) => {
        const persisted = persistedState as Partial<FormulaHotkeyState>;
        return {
          ...currentState,
          bindings: normalizeBindings(persisted.bindings),
        };
      },
    },
  ),
);
