import type { FormulaHotkeyChord } from "../../../shared/visualtex-app/src/shortcuts/formulaHotkeys";

/** An Office action that already uses a chord: its label, and giving the chord up. */
export interface OfficeHotkeyOwner {
  label: string;
  release(): Promise<void>;
}

const noOwner = (_chord: FormulaHotkeyChord, _language: "cn" | "en"): OfficeHotkeyOwner | null => null;

/**
 * Windows' Office shortcuts act inside Word, PowerPoint and WPS (the add-ins read them from the editor preferences),
 * not system-wide, so no chord of VisualTeX's own shortcuts belongs to an Office action.
 */
export function useOfficeHotkeyOwner() {
  return noOwner;
}
