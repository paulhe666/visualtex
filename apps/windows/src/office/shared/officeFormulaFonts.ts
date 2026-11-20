import {
  normalizeFormulaChineseFont,
  normalizeFormulaFaces,
  normalizeFormulaLetterFont,
  persistFormulaChineseFontPreference,
  persistFormulaFaces,
  persistFormulaLetterFontPreference,
} from "../../../../shared/visualtex-app/src/editor/formulaFontPreferences";
import {
  normalizeFormulaSpacingPreferences,
  persistFormulaSpacing,
} from "../../../../shared/visualtex-app/src/editor/formulaSpacingPreferences";
import type { OfficePreferences } from "./sessionClient";

/**
 * Keeps the main app's formula fonts - the math font, the Chinese font, the spacing configurations and the fonts for
 * letters, Greek letters and digits - in this Office page's own storage (the companion's origin), where the layout
 * engine reads them, for a page that draws formulas without the editor's store (bulk import's previews). The Office
 * editor takes them through its store (applyOfficeEditorPreferences).
 */
export function storeOfficeFormulaFonts(preferences: OfficePreferences) {
  const settings = preferences.editorPreferences?.settings;
  if (!settings) return;
  if (settings.formulaLetterFont !== undefined) {
    persistFormulaLetterFontPreference(normalizeFormulaLetterFont(settings.formulaLetterFont));
  }
  if (settings.formulaChineseFont !== undefined) {
    persistFormulaChineseFontPreference(normalizeFormulaChineseFont(settings.formulaChineseFont));
  }
  if (settings.formulaSpacing !== undefined) {
    persistFormulaSpacing(normalizeFormulaSpacingPreferences(settings.formulaSpacing));
  }
  if (settings.formulaFaces !== undefined) persistFormulaFaces(normalizeFormulaFaces(settings.formulaFaces));
}
