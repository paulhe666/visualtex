import { create } from "zustand";
import { createJSONStorage, persist } from "zustand/middleware";
import { invoke, isTauri } from "@tauri-apps/api/core";
import type { FormulaHotkeyChord } from "../shortcuts/formulaHotkeys";
import { formulaHotkeyChordId } from "../shortcuts/formulaHotkeys";
import {
  ANYWHERE_HOTKEY_STORAGE_KEY,
  anywhereHotkeyDefinition,
  defaultAnywhereHotkeyBindings,
  normalizeAnywhereHotkeyBindings,
  type AnywhereHotkeyAction,
  type AnywhereHotkeyBindings,
} from "../shortcuts/anywhereHotkeys";
import { safeStorage } from "../runtime/safeStorage";

interface AnywhereHotkeyState {
  bindings: AnywhereHotkeyBindings;
  /** Gives `action` the chord (null: none); another action that had it loses it. */
  setBinding: (action: AnywhereHotkeyAction, chord: FormulaHotkeyChord | null) => void;
  resetBinding: (action: AnywhereHotkeyAction) => void;
}

export const useAnywhereHotkeyStore = create<AnywhereHotkeyState>()(
  persist(
    (set) => ({
      bindings: defaultAnywhereHotkeyBindings(),
      setBinding: (action, chord) =>
        set((state) => {
          const key = chord ? formulaHotkeyChordId(chord) : null;
          const bindings = { ...state.bindings };
          for (const other of Object.keys(bindings) as AnywhereHotkeyAction[]) {
            const bound = bindings[other];
            if (other !== action && key && bound && formulaHotkeyChordId(bound) === key) bindings[other] = null;
          }
          bindings[action] = chord;
          return { bindings };
        }),
      resetBinding: (action) =>
        set((state) => {
          const chord = anywhereHotkeyDefinition(action).defaultChord;
          const key = chord ? formulaHotkeyChordId(chord) : null;
          const bindings = { ...state.bindings };
          for (const other of Object.keys(bindings) as AnywhereHotkeyAction[]) {
            const bound = bindings[other];
            if (other !== action && key && bound && formulaHotkeyChordId(bound) === key) bindings[other] = null;
          }
          bindings[action] = chord;
          return { bindings };
        }),
    }),
    {
      name: ANYWHERE_HOTKEY_STORAGE_KEY,
      storage: createJSONStorage(() => safeStorage),
      partialize: (state) => ({ bindings: state.bindings }),
      merge: (persisted, current) => ({
        ...current,
        bindings: normalizeAnywhereHotkeyBindings((persisted as { bindings?: unknown } | undefined)?.bindings),
      }),
    },
  ),
);

/** Registers a system-wide shortcut with macOS (it stays registered across
 *  restarts), or removes the display one (null); throws macOS's refusal. */
export async function configureAnywhereGlobalHotkey(
  action: AnywhereHotkeyAction,
  chord: FormulaHotkeyChord | null,
): Promise<string> {
  if (!isTauri()) return "";
  return invoke<string>("anywhere_configure_hotkey", {
    action,
    chord: chord
      ? {
          code: chord.code,
          ctrlKey: chord.ctrlKey,
          altKey: chord.altKey,
          shiftKey: chord.shiftKey,
          metaKey: chord.metaKey,
        }
      : null,
  });
}
