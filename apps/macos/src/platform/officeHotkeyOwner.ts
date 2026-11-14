import { useCallback } from "react";
import { isTauri } from "@tauri-apps/api/core";
import {
  formulaHotkeyChordId,
  type FormulaHotkeyChord,
} from "../../../shared/visualtex-app/src/shortcuts/formulaHotkeys";
import { useOfficeHotkeyStore } from "../stores/officeHotkeyStore";
import { officeHotkeyAction, officeHotkeyActionLabel } from "../shortcuts/officeHotkeys";
import { configureOfficeHotkeys } from "../runtime/officeHotkeys";

/** An Office action that already uses a chord: its label, and giving the chord up. */
export interface OfficeHotkeyOwner {
  label: string;
  release(): Promise<void>;
}

/**
 * macOS registers Office actions (insert or edit a formula in Word or PowerPoint) as system-wide hotkeys, so
 * another shortcut must not take their chords: the returned function says which action owns a chord.
 */
export function useOfficeHotkeyOwner() {
  const bindings = useOfficeHotkeyStore((state) => state.bindings);
  const removeBinding = useOfficeHotkeyStore((state) => state.removeBinding);
  return useCallback(
    (chord: FormulaHotkeyChord, language: "cn" | "en"): OfficeHotkeyOwner | null => {
      const id = formulaHotkeyChordId(chord);
      const binding = bindings.find((item) => formulaHotkeyChordId(item.chord) === id);
      if (!binding) return null;
      return {
        label: officeHotkeyActionLabel(officeHotkeyAction(binding.actionId), language),
        release: async () => {
          if (isTauri()) {
            await configureOfficeHotkeys(bindings.filter((item) => item.actionId !== binding.actionId));
          }
          removeBinding(binding.actionId);
        },
      };
    },
    [bindings, removeBinding],
  );
}
