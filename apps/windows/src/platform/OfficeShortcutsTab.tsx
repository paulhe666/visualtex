import { OfficeFormulaShortcutSettingsSection } from "../components/OfficeFormulaShortcutSettings";

/**
 * Settings → Keyboard shortcuts → Office on Windows: the Word, PowerPoint and WPS add-ins' own formula shortcuts
 * (with the output format for Word and WPS Writer), which the add-ins read from the editor preferences.
 */
export function OfficeShortcutsTab() {
  return <OfficeFormulaShortcutSettingsSection />;
}
