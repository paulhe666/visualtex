import { useCallback } from "react";
import type { FormulaHotkeyChord } from "../../../shared/visualtex-app/src/shortcuts/formulaHotkeys";

/** An Office action that already uses a chord: its label, and giving the chord up. */
export interface OfficeHotkeyOwner {
  label: string;
  release(): Promise<void>;
}

/** No Office actions hold chords in a browser. */
export function useOfficeHotkeyOwner() {
  return useCallback((_chord: FormulaHotkeyChord, _language: "cn" | "en"): OfficeHotkeyOwner | null => null, []);
}
